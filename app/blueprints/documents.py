import json
import logging
import os
import re
import tempfile
import threading
import hashlib
from datetime import datetime
from pathlib import Path

import werkzeug.utils
from flask import Blueprint, jsonify, request, send_file

from database.controller import (
    delete_historial_acta,
    get_historial_actas,
    get_max_numero_acta_for_year,
    get_next_numero_acta,
    numero_acta_exists,
    get_or_create_personal,
    save_historial_acta,
)

try:
    from app.utils.word_manager import (
        extract_variables_from_template,
        generate_acta,
        get_preview_unavailable_reason,
        render_docx_preview_html,
    )
    from app.utils.sse_hub import publish_event
except ModuleNotFoundError:
    from utils.word_manager import (
        extract_variables_from_template,
        generate_acta,
        get_preview_unavailable_reason,
        render_docx_preview_html,
    )
    from utils.sse_hub import publish_event


documents_bp = Blueprint("documents", __name__)
logger = logging.getLogger(__name__)

NUMERO_ACTA_PATTERN = re.compile(r"^0\d{2,}-\d{4}$")

BASE_DIR = Path(__file__).resolve().parents[2]
UPLOAD_FOLDER = os.path.join(BASE_DIR, "plantillas")
ACTAS_OUTPUT_ROOT = os.path.join(BASE_DIR, "salidas", "inventario", "actas")
PREVIEW_OUTPUT_ROOT = os.environ.get(
    "INVENTARIO_PREVIEW_ROOT",
    os.path.join(tempfile.gettempdir(), "inventario_preview"),
)
PREVIEW_CACHE = {}
PREVIEW_CACHE_LOCK = threading.Lock()
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(ACTAS_OUTPUT_ROOT, exist_ok=True)
os.makedirs(PREVIEW_OUTPUT_ROOT, exist_ok=True)


def is_output_path_allowed(path):
    if not path:
        return False
    real_path = os.path.abspath(path)
    actas_root = os.path.abspath(ACTAS_OUTPUT_ROOT)
    preview_root = os.path.abspath(PREVIEW_OUTPUT_ROOT)
    return real_path.startswith(actas_root) or real_path.startswith(preview_root)


def _client_preview_key(req):
    forwarded = (req.headers.get("X-Forwarded-For") or "").split(",")[0].strip()
    base = forwarded or req.remote_addr or "local"
    return re.sub(r"[^A-Za-z0-9_.-]", "_", base)


def _build_preview_signature(tipo, context_data, tabla_data, tabla_columnas, prefer_pdf_preview=False):
    payload = {
        "tipo": tipo,
        "context_data": context_data,
        "tabla_data": tabla_data,
        "tabla_columnas": tabla_columnas,
        "prefer_pdf_preview": bool(prefer_pdf_preview),
    }
    raw = json.dumps(payload, sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def cleanup_preview_cache(max_age_hours=24, max_files_per_folder=8):
    preview_root = PREVIEW_OUTPUT_ROOT
    if not os.path.isdir(preview_root):
        return {"deleted_files": 0, "deleted_dirs": 0, "root": preview_root}

    now_ts = datetime.now().timestamp()
    max_age_seconds = max(1, int(max_age_hours)) * 3600
    deleted_files = 0
    deleted_dirs = 0

    for current_root, _, files in os.walk(preview_root):
        file_entries = []
        for filename in files:
            file_path = os.path.join(current_root, filename)
            try:
                age_seconds = now_ts - os.path.getmtime(file_path)
                file_entries.append((file_path, os.path.getmtime(file_path)))
                if age_seconds > max_age_seconds:
                    os.remove(file_path)
                    deleted_files += 1
            except FileNotFoundError:
                continue
            except Exception as exc:
                logger.warning("No se pudo limpiar preview %s: %s", file_path, exc)

        # Limita cantidad de archivos por carpeta de preview para evitar crecimiento infinito.
        if max_files_per_folder and len(file_entries) > max_files_per_folder:
            file_entries.sort(key=lambda t: t[1], reverse=True)
            for stale_path, _ in file_entries[max_files_per_folder:]:
                try:
                    if os.path.exists(stale_path):
                        os.remove(stale_path)
                        deleted_files += 1
                except Exception:
                    continue

    # Segunda pasada para eliminar directorios vacios, sin borrar la raiz _preview.
    for current_root, _, _ in os.walk(preview_root, topdown=False):
        if current_root == preview_root:
            continue
        try:
            if not os.listdir(current_root):
                os.rmdir(current_root)
                deleted_dirs += 1
        except Exception:
            continue

    return {"deleted_files": deleted_files, "deleted_dirs": deleted_dirs, "root": preview_root}


@documents_bp.route("/api/plantillas/upload", methods=["POST"])
def api_upload_plantilla():
    if "documento" not in request.files:
        return jsonify({"success": False, "error": "No file part"}), 400

    file = request.files["documento"]
    tipo = request.form.get("tipo", "general")

    if file.filename == "":
        return jsonify({"success": False, "error": "No selected file"}), 400

    filename = file.filename or ""
    if file and filename.endswith(".docx"):
        filename = werkzeug.utils.secure_filename(f"{tipo}.docx")
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)

        variables = extract_variables_from_template(file_path)
        publish_event("templates_changed", {"tipo": tipo})
        return jsonify(
            {
                "success": True,
                "message": "Plantilla guardada y analizada existosamente",
                "variables": variables,
            }
        )
    return jsonify({"success": False, "error": "Invalid file format, must be .docx"}), 400


@documents_bp.route("/api/plantillas/estado", methods=["GET"])
def api_estado_plantillas():
    tipo = request.args.get("tipo")
    if not tipo:
        return jsonify({"success": False, "error": "Tipo no proporcionado"}), 400

    file_path = os.path.join(UPLOAD_FOLDER, f"{tipo}.docx")
    if os.path.exists(file_path):
        variables = extract_variables_from_template(file_path)
        return jsonify({"success": True, "existe": True, "variables": variables})
    return jsonify({"success": True, "existe": False, "variables": []})


@documents_bp.route("/api/plantillas/descargar", methods=["GET"])
def api_descargar_plantilla():
    tipo = request.args.get("tipo")
    if not tipo:
        return jsonify({"success": False, "error": "Tipo no proporcionado"}), 400

    safe_tipo = werkzeug.utils.secure_filename(str(tipo))
    file_path = os.path.join(UPLOAD_FOLDER, f"{safe_tipo}.docx")
    if not os.path.exists(file_path):
        return jsonify({"success": False, "error": "No existe plantilla para este tipo."}), 404

    return send_file(
        file_path,
        as_attachment=True,
        download_name=f"{safe_tipo}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@documents_bp.route("/api/informes/generar", methods=["POST"])
def api_generar_informe():
    payload = request.json
    tipo = payload.get("tipo", "acta_entrega")
    context_data_raw = dict(payload.get("datos_formulario", {}) or {})
    context_data = dict(context_data_raw or {})
    vista_previa = payload.get("vista_previa", False)

    # Compatibilidad entre plantillas que usan usuario_final y otras que usan recibido_por.
    if context_data.get("recibido_por") and not context_data.get("usuario_final"):
        context_data["usuario_final"] = context_data.get("recibido_por")
    if context_data.get("usuario_final") and not context_data.get("recibido_por"):
        context_data["recibido_por"] = context_data.get("usuario_final")

    current_year = datetime.now().year
    numero_acta = str(context_data.get("numero_acta") or "").strip()
    if not numero_acta:
        numero_acta = get_next_numero_acta(current_year)

    if not NUMERO_ACTA_PATTERN.match(numero_acta):
        return jsonify(
            {
                "success": False,
                "error": "El numero de acta debe tener formato 0NNN-AAAA, por ejemplo 012-2026.",
            }
        ), 400

    try:
        numero_seq_str, numero_year_str = numero_acta.split("-", 1)
        numero_seq = int(numero_seq_str)
        numero_year = int(numero_year_str)
    except Exception:
        return jsonify({"success": False, "error": "Numero de acta invalido."}), 400

    if numero_year != current_year:
        return jsonify(
            {
                "success": False,
                "error": f"El numero de acta debe corresponder al año actual ({current_year}).",
            }
        ), 400

    max_for_year = get_max_numero_acta_for_year(current_year)
    if not vista_previa and numero_seq < max_for_year:
        return jsonify(
            {
                "success": False,
                "error": (
                    f"El numero de acta {numero_acta} es menor al ultimo registrado. "
                    f"Debe usar uno mayor o igual al consecutivo actual {max_for_year:03d}-{current_year}."
                ),
            }
        ), 409

    if not vista_previa and numero_acta_exists(numero_acta):
        return jsonify(
            {
                "success": False,
                "error": f"Ya existe un acta registrada con el numero {numero_acta}.",
            }
        ), 409

    context_data["numero_acta"] = numero_acta
    context_data_raw["numero_acta"] = numero_acta

    tabla_data = payload.get("datos_tabla", [])
    tabla_columnas = payload.get("datos_columnas", [])

    nombres_campos_personal = ["entregado_por", "recibido_por", "usuario_final", "administradora"]
    for campo in nombres_campos_personal:
        if campo in context_data and context_data[campo]:
            get_or_create_personal(context_data[campo])

    template_path = os.path.join(UPLOAD_FOLDER, f"{tipo}.docx")
    if not os.path.exists(template_path):
        return jsonify({"success": False, "error": "No existe plantilla cargada para este tipo."}), 404

    try:
        output_dir = None
        doc_name = f"acta_{tipo}"
        generate_pdf = False
        tipo_slug = str(tipo).replace("_", "-").strip().lower() or "entrega"

        if not vista_previa:
            folder_name = f"acta-{tipo_slug}"
            output_dir = os.path.join(ACTAS_OUTPUT_ROOT, folder_name)
            fecha_hoy = datetime.now().strftime("%d-%m-%Y")
            doc_name = f"acta-{tipo_slug}-{fecha_hoy}"
        else:
            client_key = _client_preview_key(request)
            output_dir = os.path.join(PREVIEW_OUTPUT_ROOT, f"acta-{tipo_slug}", client_key)
            cleanup_preview_cache(max_age_hours=6, max_files_per_folder=8)

            preview_signature = _build_preview_signature(
                tipo,
                context_data,
                tabla_data,
                tabla_columnas,
                prefer_pdf_preview=False,
            )
            cache_key = f"{tipo_slug}:{client_key}"
            with PREVIEW_CACHE_LOCK:
                cached = PREVIEW_CACHE.get(cache_key)
            if cached and cached.get("signature") == preview_signature:
                cached_docx = cached.get("docx_path")
                cached_pdf = cached.get("pdf_path")
                if cached_pdf and os.path.exists(cached_pdf):
                    return jsonify(
                        {
                            "success": True,
                            "docx_path": cached_docx,
                            "pdf_path": cached_pdf,
                            "html_preview": None,
                            "preview_warning": None,
                            "cached": True,
                            "message": "Vista previa recuperada desde cache.",
                        }
                    )

        docx_path, pdf_path = generate_acta(
            template_path=template_path,
            context_data=context_data,
            table_data=tabla_data,
            table_columns=tabla_columnas,
            output_dir=output_dir,
            generate_pdf=generate_pdf,
            doc_name=doc_name,
            use_date_subfolder=False,
            include_time_suffix=True,
        )

        if not docx_path:
            return jsonify(
                {
                    "success": False,
                    "error": (
                        "No se pudo escribir el archivo DOCX de vista previa. "
                        "Cierra documentos abiertos de actas o espera unos segundos e intenta de nuevo."
                    ),
                }
            ), 500

        html_preview = render_docx_preview_html(docx_path) if vista_previa and not pdf_path else None
        preview_warning = None
        if vista_previa and not pdf_path and not html_preview:
            preview_warning = get_preview_unavailable_reason() or (
                "No se pudo renderizar la vista previa por ahora. "
                "Word puede estar ocupado o la conversion tardo demasiado; intenta nuevamente."
            )

        if not vista_previa:
            datos_completos = {
                "formulario": context_data_raw,
                "tabla": tabla_data,
                "columnas": tabla_columnas,
            }
            save_historial_acta(tipo, json.dumps(datos_completos), docx_path, pdf_path, numero_acta=numero_acta)
            publish_event("actas_changed", {"tipo": tipo})
        else:
            with PREVIEW_CACHE_LOCK:
                PREVIEW_CACHE[f"{tipo_slug}:{_client_preview_key(request)}"] = {
                    "signature": _build_preview_signature(
                        tipo,
                        context_data,
                        tabla_data,
                        tabla_columnas,
                        prefer_pdf_preview=False,
                    ),
                    "docx_path": docx_path,
                    "pdf_path": pdf_path,
                }

        return jsonify(
            {
                "success": True,
                "docx_path": docx_path,
                "pdf_path": pdf_path,
                "numero_acta": numero_acta,
                "html_preview": html_preview,
                "preview_warning": preview_warning,
                "message": "Archivo generado exitosamente en " + (docx_path if docx_path else "ruta desconocida"),
            }
        )
    except Exception as error:
        logger.exception("Error no controlado generando informe tipo=%s", tipo)
        return jsonify({"success": False, "error": str(error)}), 500


@documents_bp.route("/api/informes/preview/cleanup", methods=["POST"])
def api_cleanup_preview_cache():
    payload = request.json or {}
    max_age_hours = payload.get("max_age_hours", 24)
    try:
        max_age_hours = int(max_age_hours)
    except Exception:
        max_age_hours = 24

    result = cleanup_preview_cache(max_age_hours=max_age_hours)
    return jsonify({"success": True, "data": result})


@documents_bp.route("/api/historial", methods=["GET"])
def api_get_historial_all():
    tipo = request.args.get("tipo_acta")
    historial = get_historial_actas(tipo)
    return jsonify({"success": True, "data": historial})


@documents_bp.route("/api/historial/numero-acta/siguiente", methods=["GET"])
def api_numero_acta_siguiente():
    current_year = datetime.now().year
    next_num = get_next_numero_acta(current_year)
    return jsonify({"success": True, "numero_acta": next_num, "year": current_year})


@documents_bp.route("/api/historial/numero-acta/validar", methods=["GET"])
def api_numero_acta_validar():
    numero_acta = str(request.args.get("numero_acta") or "").strip()
    current_year = datetime.now().year

    if not NUMERO_ACTA_PATTERN.match(numero_acta):
        return jsonify({"success": True, "valid": False, "reason": "format"})

    seq_str, year_str = numero_acta.split("-", 1)
    seq = int(seq_str)
    year = int(year_str)
    if year != current_year:
        return jsonify({"success": True, "valid": False, "reason": "year"})

    max_for_year = get_max_numero_acta_for_year(current_year)
    exists = numero_acta_exists(numero_acta)
    lower_than_max = seq < max_for_year

    return jsonify(
        {
            "success": True,
            "valid": (not exists and not lower_than_max),
            "exists": exists,
            "lower_than_max": lower_than_max,
            "max_numero_acta": f"{max_for_year:03d}-{current_year}" if max_for_year > 0 else None,
            "next_numero_acta": get_next_numero_acta(current_year),
        }
    )


@documents_bp.route("/api/historial/<int:acta_id>", methods=["DELETE"])
def api_delete_historial(acta_id):
    delete_historial_acta(acta_id)
    publish_event("actas_changed", {"acta_id": acta_id, "action": "delete"})
    return jsonify({"success": True})
