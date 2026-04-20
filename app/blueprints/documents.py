import json
import logging
import os
import re
import sqlite3
import shutil
import tempfile
import threading
import hashlib
import zipfile
import uuid
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

import werkzeug.utils
from flask import Blueprint, current_app, jsonify, request, send_file

from database.historial_repository import (
    count_historial_by_template_snapshot,
    delete_historial_acta,
    get_historial_actas,
    get_max_numero_acta_for_year,
    get_next_numero_acta,
    get_next_numero_informe_area,
    reserve_numero_acta,
    reserve_numeros_informe_area,
    get_or_create_personal,
    save_historial_acta,
    update_historial_acta,
)
from database.db import get_db

try:
    from app.utils.word_manager import (
        extract_variables_from_template,
        generate_acta,
        get_preview_unavailable_reason,
        render_docx_preview_html,
    )
    from app.utils.sse_hub import publish_event
    from app.utils.filesystem import cleanup_empty_parent_dirs
except ModuleNotFoundError:
    from utils.word_manager import (
        extract_variables_from_template,
        generate_acta,
        get_preview_unavailable_reason,
        render_docx_preview_html,
    )
    from utils.sse_hub import publish_event
    from utils.filesystem import cleanup_empty_parent_dirs


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
TEMPLATE_HISTORY_ROOT = os.path.join(UPLOAD_FOLDER, "_historial")
AREA_REPORTS_OUTPUT_ROOT = os.path.join(ACTAS_OUTPUT_ROOT, "informes-area")
PREVIEW_CACHE = {}
PREVIEW_CACHE_LOCK = threading.Lock()
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(ACTAS_OUTPUT_ROOT, exist_ok=True)
os.makedirs(PREVIEW_OUTPUT_ROOT, exist_ok=True)
os.makedirs(TEMPLATE_HISTORY_ROOT, exist_ok=True)
os.makedirs(AREA_REPORTS_OUTPUT_ROOT, exist_ok=True)

AREA_REPORT_TASKS = {}
AREA_REPORT_TASKS_LOCK = threading.Lock()
AREA_REPORT_EXECUTOR = ThreadPoolExecutor(max_workers=2, thread_name_prefix="area-report-worker")


def _get_area_report_task(job_id):
    with AREA_REPORT_TASKS_LOCK:
        return dict(AREA_REPORT_TASKS.get(job_id, {}) or {})


def _safe_filename_part(text):
    value = str(text or "").strip().lower()
    value = re.sub(r"[^a-z0-9_-]+", "-", value)
    value = re.sub(r"-+", "-", value).strip("-")
    return value or "sin-numero"


def _format_fecha_es_larga(value):
    raw = str(value or "").strip()
    if not raw:
        return ""
    raw = raw[:10] if re.match(r"^\d{4}-\d{2}-\d{2}", raw) else raw
    try:
        dt = datetime.strptime(raw, "%Y-%m-%d")
    except Exception:
        return raw

    meses = [
        "enero",
        "febrero",
        "marzo",
        "abril",
        "mayo",
        "junio",
        "julio",
        "agosto",
        "septiembre",
        "octubre",
        "noviembre",
        "diciembre",
    ]
    return f"{dt.day} de {meses[dt.month - 1]} del {dt.year}"


def _build_acta_procedencia_text(tipo_norm, numero_acta):
    tipo = str(tipo_norm or "").strip().lower()
    numero = str(numero_acta or "").strip()
    if not numero:
        return ""
    if tipo == "entrega":
        return f"Acta de entrega {numero}"
    if tipo == "recepcion":
        return f"Acta de recepción {numero}"
    if tipo in {"baja", "bajas"}:
        return f"Acta de baja {numero}"
    return f"Acta {numero}"


_ACTA_INVENTORY_STATE_FIELDS = [
    "area_id",
    "ubicacion",
    "estado",
    "justificacion",
    "procedencia",
]


def _snapshot_inventory_item_state(db, item_id):
    row = db.execute(
        """
        SELECT id, area_id, ubicacion, estado, justificacion, procedencia
        FROM inventario_items
        WHERE id = ?
        """,
        (int(item_id),),
    ).fetchone()
    return dict(row) if row else None


def _states_are_different(before_state, after_state):
    before = dict(before_state or {})
    after = dict(after_state or {})
    for field in _ACTA_INVENTORY_STATE_FIELDS:
        if before.get(field) != after.get(field):
            return True
    return False


def _persist_acta_inventory_mutations(db, acta_id, tipo_acta, pending_mutations):
    safe_acta_id = int(acta_id)
    safe_tipo = str(tipo_acta or "").strip().lower() or "general"
    for entry in pending_mutations or []:
        kind = str(entry.get("kind") or "").strip().lower()
        if kind not in {"update", "insert"}:
            continue
        item_id = _parse_positive_int(entry.get("item_id"))
        if kind == "insert" and not item_id:
            continue
        before_json = json.dumps(entry.get("before"), ensure_ascii=False) if entry.get("before") is not None else None
        after_json = json.dumps(entry.get("after"), ensure_ascii=False) if entry.get("after") is not None else None
        db.execute(
            """
            INSERT INTO acta_inventory_mutaciones (
                acta_id,
                tipo_acta,
                item_id,
                mutation_kind,
                before_data_json,
                after_data_json
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (safe_acta_id, safe_tipo, item_id, kind, before_json, after_json),
        )


def _revert_acta_inventory_mutations(db, acta_id):
    safe_acta_id = int(acta_id)
    rows = db.execute(
        """
        SELECT id, item_id, mutation_kind, before_data_json, after_data_json
        FROM acta_inventory_mutaciones
        WHERE acta_id = ?
        ORDER BY id DESC
        """,
        (safe_acta_id,),
    ).fetchall()

    for row in rows:
        kind = str(row["mutation_kind"] or "").strip().lower()
        item_id = _parse_positive_int(row["item_id"])

        if kind == "insert":
            if item_id:
                # Si es una inserción (como en recepción), revertir significa eliminar el ítem
                db.execute("DELETE FROM inventario_items WHERE id = ?", (item_id,))
            continue

        if kind == "transfer_out":
            # REVERSIÓN DE TRASPASO: Mover de inventario_traspasos de vuelta a inventario_items
            traspaso_row = db.execute(
                "SELECT * FROM inventario_traspasos WHERE id = ?", (item_id,)
            ).fetchone()
            if traspaso_row:
                data = dict(traspaso_row)
                # Recuperar datos originales si existen
                original_data = {}
                if data.get("datos_originales_json"):
                    try:
                        original_data = json.loads(data["datos_originales_json"])
                    except: pass
                
                db.execute(
                    """
                    INSERT INTO inventario_items (
                        id, item_numero, cod_inventario, cod_esbye, cuenta, cantidad, descripcion,
                        ubicacion, marca, modelo, serie, estado, condicion, usuario_final,
                        fecha_adquisicion, valor, observacion, justificacion, procedencia, area_id
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        data["id"], data["item_numero"], data["cod_inventario"], data["cod_esbye"],
                        data["cuenta"], data["cantidad"], data["descripcion"], data["ubicacion"],
                        data["marca"], data["modelo"], data["serie"], data["estado"],
                        data["condicion"], data["usuario_final"], data["fecha_adquisicion"],
                        data["valor"], data["observacion"], data["justificacion"],
                        data["procedencia"], data["area_id"]
                    )
                )
                db.execute("DELETE FROM inventario_traspasos WHERE id = ?", (item_id,))
            continue

        if kind != "update" or not item_id:
            continue

        before_raw = str(row["before_data_json"] or "").strip()
        if not before_raw:
            continue
        try:
            before_data = json.loads(before_raw)
        except Exception:
            continue

        db.execute(
            """
            UPDATE inventario_items
            SET
                area_id = ?,
                ubicacion = ?,
                estado = ?,
                justificacion = ?,
                procedencia = ?,
                actualizado_en = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (
                before_data.get("area_id"),
                before_data.get("ubicacion"),
                before_data.get("estado"),
                before_data.get("justificacion"),
                before_data.get("procedencia"),
                item_id,
            ),
        )


def _clear_acta_inventory_mutations(db, acta_id):
    db.execute("DELETE FROM acta_inventory_mutaciones WHERE acta_id = ?", (int(acta_id),))


def _is_fecha_key(key):
    normalized = str(key or "").strip().lower()
    if not normalized:
        return False
    return normalized == "fecha" or normalized.startswith("fecha_") or normalized.endswith("_fecha")


def _normalize_fecha_fields(record):
    data = dict(record or {})
    for key in list(data.keys()):
        if _is_fecha_key(key):
            data[key] = _format_fecha_es_larga(data.get(key))
    return data


def _normalize_table_dates(table_data):
    if not isinstance(table_data, list):
        return table_data
    normalized = []
    for row in table_data:
        if isinstance(row, dict):
            normalized.append(_normalize_fecha_fields(row))
        else:
            normalized.append(row)
    return normalized


def _normalize_informe_context(tipo, context_data):
    ctx = _normalize_fecha_fields(context_data)
    tipo_norm = str(tipo or "").strip().lower()

    # Compatibilidad de nombres históricos en plantillas.
    if not ctx.get("memorandum") and ctx.get("memorando"):
        ctx["memorandum"] = ctx.get("memorando")

    if tipo_norm == "recepcion":
        if not ctx.get("fecha_elaboracion") and ctx.get("fecha_emision"):
            ctx["fecha_elaboracion"] = ctx.get("fecha_emision")
        if not ctx.get("fecha_emision") and ctx.get("fecha_elaboracion"):
            ctx["fecha_emision"] = ctx.get("fecha_elaboracion")

    # Alias con acentos por compatibilidad si la plantilla los usa.
    if ctx.get("fecha_elaboracion") and not ctx.get("fecha_elaboración"):
        ctx["fecha_elaboración"] = ctx.get("fecha_elaboracion")
    if ctx.get("accion_personal") and not ctx.get("acción_personal"):
        ctx["acción_personal"] = ctx.get("accion_personal")

    return ctx


def _parse_positive_int(value):
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return None
    return parsed if parsed > 0 else None


def _build_allowed_recepcion_location_tokens():
    db = get_db()
    rows = db.execute(
        """
        SELECT
            b.nombre AS bloque_nombre,
            p.nombre AS piso_nombre,
            a.nombre AS area_nombre
        FROM areas a
        JOIN pisos p ON p.id = a.piso_id
        JOIN bloques b ON b.id = p.bloque_id
        ORDER BY b.nombre, p.nombre, a.nombre
        """
    ).fetchall()

    allowed = set()
    for row in rows:
        bloque = str(row["bloque_nombre"] or "").strip()
        piso = str(row["piso_nombre"] or "").strip()
        area = str(row["area_nombre"] or "").strip()
        if bloque:
            allowed.add(f"Bloque {bloque}")
        if piso and bloque:
            allowed.add(f"Piso {piso} (Bloque {bloque})")
        if area and piso and bloque:
            allowed.add(f"Area {area} (Piso {piso}, Bloque {bloque})")
    return allowed


def _resolve_area_from_form(context_data):
    area_id = _parse_positive_int(context_data.get("ubicacion_area_id"))
    if not area_id:
        return None
    row = get_db().execute(
        "SELECT id FROM areas WHERE id = ? LIMIT 1",
        (area_id,),
    ).fetchone()
    if not row:
        return None
    return int(row["id"])


def _get_target_areas_for_report(scope, area_id=None, piso_id=None, bloque_id=None):
    db = get_db()
    scope_normalized = str(scope or "area").strip().lower()
    params = []
    where = ""

    if scope_normalized == "area":
        if not area_id:
            raise ValueError("Debe seleccionar un area.")
        where = "WHERE a.id = ?"
        params = [int(area_id)]
    elif scope_normalized == "piso":
        if not piso_id:
            raise ValueError("Debe seleccionar un piso.")
        where = "WHERE p.id = ?"
        params = [int(piso_id)]
    elif scope_normalized == "bloque":
        if not bloque_id:
            raise ValueError("Debe seleccionar un bloque.")
        where = "WHERE b.id = ?"
        params = [int(bloque_id)]
    else:
        raise ValueError("Scope invalido. Use area, piso o bloque.")

    rows = db.execute(
        f"""
        SELECT
            a.id AS area_id,
            a.nombre AS area_nombre,
            p.id AS piso_id,
            p.nombre AS piso_nombre,
            b.id AS bloque_id,
            b.nombre AS bloque_nombre
        FROM areas a
        JOIN pisos p ON p.id = a.piso_id
        JOIN bloques b ON b.id = p.bloque_id
        {where}
        ORDER BY b.nombre ASC, p.nombre ASC, a.nombre ASC
        """,
        tuple(params),
    ).fetchall()
    return [dict(row) for row in rows]


def _build_area_summary_rows(area_id):
    db = get_db()
    rows = db.execute(
        """
        SELECT
            COALESCE(NULLIF(TRIM(descripcion), ''), 'SIN DESCRIPCION') AS descripcion,
            SUM(COALESCE(cantidad, 1)) AS cantidad
        FROM inventario_items
        WHERE area_id = ?
        GROUP BY COALESCE(NULLIF(TRIM(descripcion), ''), 'SIN DESCRIPCION')
        ORDER BY descripcion ASC
        """,
        (int(area_id),),
    ).fetchall()

    table_data = [
        {"descripcion": str(row["descripcion"]), "cantidad": int(row["cantidad"] or 0)}
        for row in rows
    ]
    total_bienes = sum(int(item["cantidad"] or 0) for item in table_data)
    table_data.append({"descripcion": "TOTAL DE BIENES", "cantidad": total_bienes})
    return table_data, total_bienes


def _set_area_report_task_state(job_id, **updates):
    with AREA_REPORT_TASKS_LOCK:
        current = AREA_REPORT_TASKS.get(job_id, {})
        current.update(updates)
        AREA_REPORT_TASKS[job_id] = current


def _run_area_report_job(app, job_id, scope, target_areas, numeros_acta, lot_dir, template_path, current_year):
    generated_docx = []

    try:
        with app.app_context():
            total = len(target_areas)
            for idx, area in enumerate(target_areas):
                paused_sent = False
                while True:
                    task_state = _get_area_report_task(job_id)
                    if not task_state:
                        raise RuntimeError("Estado del job no disponible.")

                    if task_state.get("cancel_requested"):
                        _set_area_report_task_state(
                            job_id,
                            status="cancelled",
                            message="Generación cancelada por el usuario.",
                        )
                        publish_event(
                            "areas_reports_cancelled",
                            {
                                "job_id": job_id,
                                "generated": len(generated_docx),
                                "total": total,
                                "message": "Generación cancelada por el usuario.",
                            },
                        )
                        for path in generated_docx:
                            _safe_remove_file(path, is_output_path_allowed)
                        _cleanup_area_reports_dir(lot_dir)
                        return

                    if task_state.get("pause_requested"):
                        _set_area_report_task_state(
                            job_id,
                            status="paused",
                            message="Generación en pausa.",
                        )
                        if not paused_sent:
                            paused_sent = True
                            publish_event(
                                "areas_reports_paused",
                                {
                                    "job_id": job_id,
                                    "generated": idx,
                                    "total": total,
                                    "message": "Generación en pausa.",
                                },
                            )
                        time.sleep(0.4)
                        continue

                    if paused_sent:
                        publish_event(
                            "areas_reports_resumed",
                            {
                                "job_id": job_id,
                                "generated": idx,
                                "total": total,
                                "message": "Generación reanudada.",
                            },
                        )
                    break

                numero_acta = numeros_acta[idx]
                area_nombre = str(area.get("area_nombre") or "AREA SIN NOMBRE")
                table_data, total_bienes = _build_area_summary_rows(area.get("area_id"))

                context_data = {
                    "numero_acta": numero_acta,
                    "nombre_area": area_nombre,
                    "piso": area.get("piso_nombre"),
                    "bloque": area.get("bloque_nombre"),
                    "total_bienes": total_bienes,
                    "titulo_acta": f"{area_nombre} ACTA No.{numero_acta}",
                }
                context_data = _normalize_informe_context("aula", context_data)
                table_columns = [
                    {"id": "descripcion", "label": "DESCRIPCION"},
                    {"id": "cantidad", "label": "CANTIDAD"},
                ]
                doc_name = (
                    f"acta-aula-{_safe_filename_part(area_nombre)}-"
                    f"{_safe_filename_part(numero_acta)}"
                )

                docx_path, _pdf_path = generate_acta(
                    template_path=template_path,
                    context_data=context_data,
                    table_data=table_data,
                    table_columns=table_columns,
                    output_dir=lot_dir,
                    doc_name=doc_name,
                    generate_pdf=False,
                    use_date_subfolder=False,
                    include_time_suffix=False,
                )

                if not docx_path:
                    raise RuntimeError(f"No se pudo generar DOCX para el area {area_nombre}.")

                generated_docx.append(docx_path)

                progress = int(((idx + 1) / max(total, 1)) * 100)
                _set_area_report_task_state(
                    job_id,
                    status="running",
                    progress=progress,
                    generated=idx + 1,
                    total=total,
                    message=f"Generado {idx + 1}/{total}: {area_nombre}",
                )
                publish_event(
                    "areas_reports_progress",
                    {
                        "job_id": job_id,
                        "progress": progress,
                        "generated": idx + 1,
                        "total": total,
                        "message": f"Generado {idx + 1}/{total}: {area_nombre}",
                    },
                )

            download_kind = "docx" if len(generated_docx) == 1 else "zip"

            zip_filename = "informes-area.zip"
            first_area = target_areas[0] if target_areas else {}
            bloque_nombre = _safe_filename_part(first_area.get("bloque_nombre") or "")
            piso_nombre = _safe_filename_part(first_area.get("piso_nombre") or "")
            scope_norm = str(scope or "").strip().lower()
            if scope_norm == "bloque" and bloque_nombre and bloque_nombre != "sin-numero":
                zip_filename = f"informes-area-{bloque_nombre}.zip"
            elif scope_norm == "piso":
                parts = [p for p in [bloque_nombre, piso_nombre] if p and p != "sin-numero"]
                if parts:
                    zip_filename = f"informes-area-{'-'.join(parts)}.zip"

            download_path = (
                generated_docx[0]
                if len(generated_docx) == 1
                else os.path.join(lot_dir, zip_filename)
            )
            zip_path = download_path if download_kind == "zip" else None

            if download_kind == "zip":
                with zipfile.ZipFile(download_path, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
                    for docx_path in generated_docx:
                        zipf.write(docx_path, arcname=os.path.basename(docx_path))

                # En modo ZIP limpiamos DOCX temporales.
                for docx_path in generated_docx:
                    _safe_remove_file(docx_path, is_output_path_allowed)

            next_numero = (
                f"{(int(numeros_acta[-1].split('-')[0]) + 1):03d}-{current_year}"
                if numeros_acta
                else get_next_numero_informe_area(current_year)
            )
            _set_area_report_task_state(
                job_id,
                status="completed",
                progress=100,
                zip_path=zip_path,
                download_path=download_path,
                download_kind=download_kind,
                total_generated=len(generated_docx),
                next_numero_acta=next_numero,
            )
            publish_event(
                "areas_reports_ready",
                {
                    "job_id": job_id,
                    "scope": scope,
                    "total_generated": len(generated_docx),
                    "zip_path": zip_path,
                    "download_path": download_path,
                    "download_kind": download_kind,
                    "start_numero_acta": numeros_acta[0] if numeros_acta else None,
                    "end_numero_acta": numeros_acta[-1] if numeros_acta else None,
                    "next_numero_acta": next_numero,
                },
            )
            publish_event(
                "areas_reports_changed",
                {
                    "scope": scope,
                    "generated": len(generated_docx),
                    "next_numero_acta": next_numero,
                },
            )
    except Exception as exc:
        logger.exception("Error en job asíncrono de informes por área job_id=%s", job_id)
        for path in generated_docx:
            _safe_remove_file(path, is_output_path_allowed)
        _cleanup_area_reports_dir(lot_dir)
        _set_area_report_task_state(job_id, status="error", error=str(exc))
        publish_event(
            "areas_reports_error",
            {
                "job_id": job_id,
                "error": str(exc),
            },
        )


def is_output_path_allowed(path):
    if not path:
        return False
    real_path = os.path.abspath(path)
    actas_root = os.path.abspath(ACTAS_OUTPUT_ROOT)
    preview_root = os.path.abspath(PREVIEW_OUTPUT_ROOT)
    return real_path.startswith(actas_root) or real_path.startswith(preview_root)


def is_template_history_path_allowed(path):
    if not path:
        return False
    real_path = os.path.abspath(path)
    history_root = os.path.abspath(TEMPLATE_HISTORY_ROOT)
    return real_path.startswith(history_root)


def _safe_remove_file(path, allowed_checker):
    if not path:
        return False
    if not allowed_checker(path):
        return False
    try:
        if os.path.exists(path):
            os.remove(path)
            return True
    except Exception as exc:
        logger.warning("No se pudo eliminar archivo %s: %s", path, exc)
    return False


def _cleanup_area_reports_dir(path):
    # Borra carpetas vacías bajo informes-area sin eliminar la raíz compartida.
    if not path:
        return
    cleanup_empty_parent_dirs(os.path.join(path, "_cleanup_.tmp"), AREA_REPORTS_OUTPUT_ROOT)


def _snapshot_template_for_historial(tipo, template_path):
    tipo_safe = werkzeug.utils.secure_filename(str(tipo or "general")) or "general"
    with open(template_path, "rb") as f:
        template_bytes = f.read()

    template_hash = hashlib.sha1(template_bytes).hexdigest()
    target_dir = os.path.join(TEMPLATE_HISTORY_ROOT, tipo_safe)
    os.makedirs(target_dir, exist_ok=True)
    snapshot_path = os.path.join(target_dir, f"{template_hash}.docx")

    if not os.path.exists(snapshot_path):
        shutil.copy2(template_path, snapshot_path)

    variables = extract_variables_from_template(snapshot_path)
    return template_hash, snapshot_path, variables


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
    context_data = _normalize_informe_context(tipo, context_data_raw)
    header_preview = str(request.headers.get("X-Preview-Request") or "").strip().lower() in {"1", "true", "yes"}
    vista_previa = bool(payload.get("vista_previa", False)) or header_preview

    # Compatibilidad entre plantillas que usan usuario_final y otras que usan recibido_por.
    if context_data.get("recibido_por") and not context_data.get("usuario_final"):
        context_data["usuario_final"] = context_data.get("recibido_por")
    if context_data.get("usuario_final") and not context_data.get("recibido_por"):
        context_data["recibido_por"] = context_data.get("usuario_final")

    current_year = datetime.now().year
    tipo_norm = str(tipo or "").strip().lower() or "entrega"
    editing_acta_id = _parse_positive_int(payload.get("editing_acta_id"))
    editing_row = None
    if editing_acta_id:
        editing_row = get_db().execute(
            """
            SELECT id, tipo_acta, numero_acta
            FROM historial_actas
            WHERE id = ?
            """,
            (editing_acta_id,),
        ).fetchone()
        if not editing_row:
            editing_acta_id = None
        elif str(editing_row["tipo_acta"] or "").strip().lower() != tipo_norm:
            return jsonify(
                {
                    "success": False,
                    "error": "La acta en edición no corresponde al tipo actual.",
                }
            ), 400

    requested_numero_acta = str(context_data.get("numero_acta") or "").strip()
    numero_acta = requested_numero_acta or get_next_numero_acta(current_year, tipo_acta=tipo_norm)

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

    context_data["numero_acta"] = numero_acta
    context_data_raw["numero_acta"] = numero_acta

    tabla_data = payload.get("datos_tabla", [])
    tabla_columnas = payload.get("datos_columnas", [])
    force_same_acta = bool(payload.get("force_same_acta"))

    if not vista_previa and not force_same_acta:
        db = get_db()
        duplicate_query = [
            "SELECT id, numero_acta, datos_json",
            "FROM historial_actas",
            "WHERE LOWER(COALESCE(tipo_acta, '')) = LOWER(?)",
        ]
        duplicate_params = [tipo_norm]
        if editing_acta_id:
            duplicate_query.append("AND id != ?")
            duplicate_params.append(editing_acta_id)
        duplicate_query.extend(["ORDER BY id DESC"])
        history_rows = db.execute("\n".join(duplicate_query), tuple(duplicate_params)).fetchall()

        current_form = dict(context_data_raw or {})
        current_form.pop("numero_acta", None)
        current_signature = {
            "formulario": current_form,
            "tabla": tabla_data,
            "columnas": tabla_columnas,
        }
        current_signature_json = json.dumps(current_signature, ensure_ascii=False, sort_keys=True)

        similar_numero_acta = None
        for row in history_rows:
            if not str(row["datos_json"] or "").strip():
                continue
            try:
                previous_payload = json.loads(row["datos_json"])
            except Exception:
                continue

            previous_form = dict((previous_payload or {}).get("formulario") or {})
            previous_form.pop("numero_acta", None)
            previous_signature = {
                "formulario": previous_form,
                "tabla": (previous_payload or {}).get("tabla") or [],
                "columnas": (previous_payload or {}).get("columnas") or [],
            }
            previous_signature_json = json.dumps(previous_signature, ensure_ascii=False, sort_keys=True)
            if previous_signature_json == current_signature_json:
                similar_numero_acta = str(row["numero_acta"] or "").strip() or None
                break

        if similar_numero_acta:
            return jsonify(
                {
                    "success": False,
                    "duplicate_previous": True,
                    "previous_numero_acta": similar_numero_acta,
                    "error": "Esta acta es igual a una acta guardada previamente para este tipo. ¿Desea guardarla de todas formas?",
                }
            ), 409

    entrega_selected_ids = []
    bajas_selected_rows = []
    pending_inventory_mutations = []

    target_area_id = None
    if not vista_previa and tipo_norm in {"entrega", "recepcion"}:
        target_area_id = _resolve_area_from_form(context_data)
        if not target_area_id:
            return jsonify({
                "success": False,
                "error": "Debe seleccionar una ubicación válida (área) desde los selectores.",
            }), 400
    target_area_id_int = int(target_area_id) if target_area_id is not None else None

    if not vista_previa and tipo_norm == "entrega":
        required_entrega = [
            "fecha_corte",
            "fecha_emision",
            "accion_personal",
            "entregado_por",
            "rol_entrega",
            "recibido_por",
            "rol_recibe",
            "area_trabajo",
        ]
        missing_entrega = [k for k in required_entrega if not str(context_data.get(k) or "").strip()]
        if missing_entrega:
            labels = ", ".join(missing_entrega)
            return jsonify({
                "success": False,
                "error": f"Faltan campos obligatorios en entrega: {labels}.",
            }), 400

        if not isinstance(tabla_data, list) or not tabla_data or not isinstance(tabla_columnas, list) or not tabla_columnas:
            return jsonify({
                "success": False,
                "error": "La variable {{tabla_dinamica}} es obligatoria en Entrega: debe extraer al menos un bien.",
            }), 400

        entrega_selected_ids = sorted(
            {
                item_id
                for item_id in (
                    _parse_positive_int((row or {}).get("id")) for row in (tabla_data or [])
                )
                if item_id
            }
        )
        if not entrega_selected_ids:
            return jsonify({
                "success": False,
                "error": "No se identificaron bienes/inmuebles válidos para mover. Extraiga desde Inventario y seleccione registros válidos.",
            }), 400

    if not vista_previa and tipo_norm == "recepcion":
        required_recepcion = [
            "entregado_por",
            "rol_entrega",
            "recibido_por",
            "rol_recibe",
            "fecha_corte",
            "fecha_elaboracion",
            "accion_personal",
            "memorandum",
            "fecha_memorandum",
            "entregado_por_segunda_delegada",
            "rol_entrega_segunda_delegada",
            "area_trabajo",
        ]
        missing = [k for k in required_recepcion if not str(context_data.get(k) or "").strip()]
        if missing:
            labels = ", ".join(missing)
            return jsonify({
                "success": False,
                "error": f"Faltan campos obligatorios en recepción: {labels}.",
            }), 400

        if not isinstance(tabla_data, list) or not tabla_data or not isinstance(tabla_columnas, list) or not tabla_columnas:
            return jsonify({
                "success": False,
                "error": "Debe registrar al menos un bien para generar el acta de recepción.",
            }), 400

        area_trabajo = str(context_data.get("area_trabajo") or "").strip()
        if not area_trabajo:
            return jsonify({
                "success": False,
                "error": "En Acta de Recepción, el campo Área de trabajo es obligatorio.",
            }), 400

    if not vista_previa and tipo_norm in {"baja", "bajas"}:
        required_bajas = [
            "numero_acta",
            "nombre_delegado",
            "recibido_por",
            "entregado_por",
            "rol_entrega",
            "fecha_emision",
        ]
        missing = [k for k in required_bajas if not str(context_data.get(k) or "").strip()]
        if missing:
            labels = ", ".join(missing)
            return jsonify({
                "success": False,
                "error": f"Faltan campos obligatorios en bajas: {labels}.",
            }), 400

        if not isinstance(tabla_data, list) or not tabla_data or not isinstance(tabla_columnas, list) or not tabla_columnas:
            return jsonify({
                "success": False,
                "error": "La variable {{tabla_dinamica}} es obligatoria en Bajas: seleccione al menos un bien.",
            }), 400

        seen_ids = set()
        for row in (tabla_data or []):
            row_data = dict(row or {})
            item_id = _parse_positive_int(row_data.get("id"))
            if not item_id or item_id in seen_ids:
                continue
            seen_ids.add(item_id)
            bajas_selected_rows.append(
                {
                    "id": item_id,
                    "estado": str(row_data.get("estado") or "").strip() or "MALO",
                    "justificacion": str(row_data.get("justificacion") or "").strip(),
                }
            )

        if not bajas_selected_rows:
            return jsonify({
                "success": False,
                "error": "No se identificaron bienes válidos para dar de baja.",
            }), 400

    if not vista_previa and tipo_norm == "traspaso":
        required_traspaso = [
            "fecha_corte",
            "fecha_emision",
            "entregado_por",
            "rol_entrega",
            "recibido_por",
            "rol_recibe",
            "facultad_entrega",
            "facultad_recibe",
            "descripcion_de_bienes"
        ]
        missing = [k for k in required_traspaso if not str(context_data.get(k) or "").strip()]
        if missing:
            return jsonify({"success": False, "error": f"Faltan campos obligatorios en traspaso: {', '.join(missing)}."}), 400

        if not isinstance(tabla_data, list) or not tabla_data:
            return jsonify({"success": False, "error": "Debe extraer al menos un bien para el traspaso."}), 400

    nombres_campos_personal = [
        "entregado_por",
        "recibido_por",
        "usuario_final",
        "administradora",
        "entregado_por_segunda_delegada",
    ]
    for campo in nombres_campos_personal:
        if campo in context_data and context_data[campo]:
            get_or_create_personal(context_data[campo])

    if isinstance(tabla_data, list):
        for row in tabla_data:
            if not isinstance(row, dict):
                continue
            usuario_final = str(row.get("usuario_final") or "").strip()
            if usuario_final:
                get_or_create_personal(usuario_final)

    requested_snapshot_path = str(payload.get("template_snapshot_path") or "").strip()
    template_path = os.path.join(UPLOAD_FOLDER, f"{tipo}.docx")
    if (
        not vista_previa
        and requested_snapshot_path
        and is_template_history_path_allowed(requested_snapshot_path)
        and os.path.exists(requested_snapshot_path)
    ):
        template_path = requested_snapshot_path

    if not os.path.exists(template_path):
        return jsonify({"success": False, "error": "No existe plantilla cargada para este tipo."}), 404

    try:
        output_dir = None
        doc_name = f"acta_{tipo}"
        generate_pdf = False
        tipo_slug = str(tipo).replace("_", "-").strip().lower() or "entrega"

        if not vista_previa:
            try:
                preserve_edit_numero = bool(
                    editing_row
                    and requested_numero_acta
                    and str(editing_row["numero_acta"] or "").strip() == requested_numero_acta
                )
                if preserve_edit_numero:
                    numero_acta = requested_numero_acta
                else:
                    numero_acta = reserve_numero_acta(
                        current_year,
                        preferred_numero_acta=numero_acta if requested_numero_acta else None,
                        tipo_acta=tipo_norm,
                    )
                context_data["numero_acta"] = numero_acta
                context_data_raw["numero_acta"] = numero_acta
            except ValueError:
                next_numero = get_next_numero_acta(current_year, tipo_acta=tipo_norm)
                return jsonify(
                    {
                        "success": False,
                        "error": (
                            f"El número de acta {numero_acta} ya no está disponible. "
                            f"Use {next_numero} o deje el campo vacío para asignación automática."
                        ),
                        "code": "numero_acta_unavailable",
                        "next_numero_acta": next_numero,
                    }
                ), 409

        if not vista_previa:
            folder_name = f"acta-{tipo_slug}"
            output_dir = os.path.join(ACTAS_OUTPUT_ROOT, folder_name)
            doc_name = f"acta-{tipo_slug}-{_safe_filename_part(numero_acta)}"
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
            include_time_suffix=bool(vista_previa),
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
            db = get_db()
            try:
                if editing_acta_id:
                    _revert_acta_inventory_mutations(db, editing_acta_id)
                    _clear_acta_inventory_mutations(db, editing_acta_id)

                if tipo_norm == "entrega":
                    before_states = {}
                    for item_id in entrega_selected_ids:
                        snapshot = _snapshot_inventory_item_state(db, item_id)
                        if snapshot:
                            before_states[int(item_id)] = snapshot

                    placeholders = ",".join(["?"] * len(entrega_selected_ids))
                    db.execute(
                        f"""
                        UPDATE inventario_items
                        SET area_id = ?, ubicacion = ?, procedencia = ?, actualizado_en = CURRENT_TIMESTAMP
                        WHERE id IN ({placeholders})
                        """,
                        (
                            target_area_id_int,
                            str(context_data.get("area_trabajo") or "").strip(),
                            _build_acta_procedencia_text(tipo_norm, numero_acta),
                            *entrega_selected_ids,
                        ),
                    )

                    for item_id in entrega_selected_ids:
                        item_id_int = int(item_id)
                        before_state = before_states.get(item_id_int)
                        after_state = _snapshot_inventory_item_state(db, item_id_int)
                        if not before_state or not after_state:
                            continue
                        if not _states_are_different(before_state, after_state):
                            continue
                        pending_inventory_mutations.append(
                            {
                                "kind": "update",
                                "item_id": item_id_int,
                                "before": before_state,
                                "after": after_state,
                            }
                        )

                if tipo_norm == "recepcion":
                    # Usar el contexto de la base de datos para manejar la transacción automáticamente
                    with db:
                        next_item_numero_row = db.execute(
                            "SELECT COALESCE(MAX(item_numero), 0) + 1 AS next_item FROM inventario_items"
                        ).fetchone()
                        next_item_numero = int(next_item_numero_row["next_item"]) if next_item_numero_row else 1

                        for idx, row in enumerate(tabla_data or []):
                            item = row or {}
                            cantidad_raw = item.get("cantidad")
                            valor_raw = item.get("valor")
                            cantidad_text = str(cantidad_raw or "").strip()
                            valor_text = str(valor_raw or "").strip()
                            try:
                                cantidad = int(cantidad_text) if cantidad_text else 1
                            except (TypeError, ValueError):
                                cantidad = 1
                            try:
                                valor = float(valor_text) if valor_text else None
                            except (TypeError, ValueError):
                                valor = None

                            db.execute(
                            """
                            INSERT INTO inventario_items (
                                item_numero,
                                cod_inventario,
                                cod_esbye,
                                cuenta,
                                cantidad,
                                descripcion,
                                ubicacion,
                                marca,
                                modelo,
                                serie,
                                estado,
                                usuario_final,
                                fecha_adquisicion,
                                valor,
                                observacion,
                                procedencia,
                                area_id
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                next_item_numero + idx,
                                str(item.get("cod_inventario") or "").strip() or None,
                                str(item.get("cod_esbye") or "").strip() or None,
                                str(item.get("cuenta") or "").strip() or None,
                                cantidad,
                                str(item.get("descripcion") or "").strip() or None,
                                str(context_data.get("area_trabajo") or "").strip() or None,
                                str(item.get("marca") or "").strip() or None,
                                str(item.get("modelo") or "").strip() or None,
                                str(item.get("serie") or "").strip() or None,
                                str(item.get("estado") or "").strip() or None,
                                str(item.get("usuario_final") or "").strip() or None,
                                str(item.get("fecha_adquisicion") or "").strip() or None,
                                valor,
                                str(item.get("observacion") or "").strip() or None,
                                _build_acta_procedencia_text(tipo_norm, numero_acta) or None,
                                target_area_id_int,
                            ),
                        )

                        inserted_id = _parse_positive_int(db.execute("SELECT last_insert_rowid() AS id").fetchone()["id"])
                        if inserted_id:
                            inserted_state = _snapshot_inventory_item_state(db, inserted_id)
                            pending_inventory_mutations.append(
                                {
                                    "kind": "insert",
                                    "item_id": inserted_id,
                                    "before": None,
                                    "after": inserted_state,
                                }
                            )

                if tipo_norm == "traspaso":
                    facultad_recibe = str(context_data.get("facultad_recibe") or "OTRA FACULTAD").strip()
                    for row in tabla_data or []:
                        item_id = _parse_positive_int(row.get("id"))
                        if not item_id:
                            continue
                        
                        # Obtener datos completos antes de mover
                        full_item = db.execute("SELECT * FROM inventario_items WHERE id = ?", (item_id,)).fetchone()
                        if not full_item:
                            continue
                        
                        data = dict(full_item)
                        fecha_local = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        # Mover a tabla de traspasos
                        db.execute(
                            """
                            INSERT INTO inventario_traspasos (
                                id, item_numero, cod_inventario, cod_esbye, cuenta, cantidad, descripcion,
                                ubicacion, marca, modelo, serie, estado, condicion, usuario_final,
                                fecha_adquisicion, valor, observacion, justificacion, procedencia, area_id,
                                facultad_destino, acta_traspaso_id, fecha_traspaso, datos_originales_json
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                data["id"], data["item_numero"], data["cod_inventario"], data["cod_esbye"],
                                data["cuenta"], data["cantidad"], data["descripcion"], data["ubicacion"],
                                data["marca"], data["modelo"], data["serie"], data["estado"],
                                data["condicion"], data["usuario_final"], data["fecha_adquisicion"],
                                data["valor"], data["observacion"], data["justificacion"],
                                data["procedencia"], data["area_id"],
                                facultad_recibe, None, fecha_local,
                                json.dumps(data)
                            )
                        )
                        # Borrar del inventario general
                        db.execute("DELETE FROM inventario_items WHERE id = ?", (item_id,))
                        
                        pending_inventory_mutations.append({
                            "kind": "transfer_out",
                            "item_id": item_id,
                            "before": data,
                            "after": {"status": "transferred", "destination": facultad_recibe}
                        })

                if tipo_norm in {"baja", "bajas"}:
                    procedencia_text = _build_acta_procedencia_text("bajas", numero_acta)
                    for row in bajas_selected_rows:
                        item_id_int = int(row.get("id"))
                        before_state = _snapshot_inventory_item_state(db, item_id_int)
                        db.execute(
                            """
                            UPDATE inventario_items
                            SET
                                estado = ?,
                                justificacion = ?,
                                procedencia = ?,
                                actualizado_en = CURRENT_TIMESTAMP
                            WHERE id = ?
                            """,
                            (
                                str(row.get("estado") or "").strip() or "MALO",
                                str(row.get("justificacion") or "").strip() or None,
                                procedencia_text or None,
                                item_id_int,
                            ),
                        )
                        after_state = _snapshot_inventory_item_state(db, item_id_int)
                        if before_state and after_state and _states_are_different(before_state, after_state):
                            pending_inventory_mutations.append(
                                {
                                    "kind": "update",
                                    "item_id": item_id_int,
                                    "before": before_state,
                                    "after": after_state,
                                }
                            )
            except sqlite3.DatabaseError:
                logger.exception("Error de base de datos al actualizar inventario en generación de acta")
                return jsonify(
                    {
                        "success": False,
                        "error": "La base de datos presenta corrupción (database disk image is malformed). Debe ejecutar reparación o restaurar respaldo antes de guardar cambios de inventario.",
                    }
                ), 500

            plantilla_hash, plantilla_snapshot_path, plantilla_variables = _snapshot_template_for_historial(
                tipo,
                template_path,
            )
            datos_completos = {
                "formulario": context_data_raw,
                "tabla": tabla_data,
                "columnas": tabla_columnas,
                "plantilla": {
                    "hash": plantilla_hash,
                    "snapshot_path": plantilla_snapshot_path,
                    "variables": plantilla_variables,
                },
            }
            target_acta_id = None
            if editing_acta_id and editing_row:
                update_historial_acta(
                    editing_acta_id,
                    tipo,
                    json.dumps(datos_completos),
                    docx_path,
                    pdf_path,
                    numero_acta=numero_acta,
                    plantilla_hash=plantilla_hash,
                    plantilla_snapshot_path=plantilla_snapshot_path,
                )
                target_acta_id = int(editing_acta_id)
            else:
                target_acta_id = save_historial_acta(
                    tipo,
                    json.dumps(datos_completos),
                    docx_path,
                    pdf_path,
                    numero_acta=numero_acta,
                    plantilla_hash=plantilla_hash,
                    plantilla_snapshot_path=plantilla_snapshot_path,
                )

            if target_acta_id and pending_inventory_mutations:
                _persist_acta_inventory_mutations(db, target_acta_id, tipo_norm, pending_inventory_mutations)
                
                if tipo_norm == "traspaso":
                    # Vincular los bienes en la tabla de traspasos con esta acta
                    db.execute(
                        "UPDATE inventario_traspasos SET acta_traspaso_id = ? WHERE acta_traspaso_id IS NULL",
                        (target_acta_id,)
                    )
                
                db.commit()

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
    except sqlite3.IntegrityError:
        try:
            get_db().rollback()
        except Exception:
            pass
        logger.exception("Conflicto de integridad al guardar acta tipo=%s numero=%s", tipo, numero_acta)
        next_numero = get_next_numero_acta(current_year, tipo_acta=tipo_norm)
        return jsonify(
            {
                "success": False,
                "error": (
                    f"El número de acta {numero_acta} ya fue usado por otra operación. "
                    f"Se sugiere usar {next_numero}."
                ),
                "code": "numero_acta_conflict",
                "next_numero_acta": next_numero,
            }
        ), 409
    except Exception as error:
        try:
            get_db().rollback()
        except Exception:
            pass
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
    tipo_acta = str(request.args.get("tipo_acta") or "entrega").strip().lower()
    next_num = get_next_numero_acta(current_year, tipo_acta=tipo_acta)
    return jsonify({"success": True, "numero_acta": next_num, "year": current_year, "tipo_acta": tipo_acta})


@documents_bp.route("/api/historial/numero-acta/validar", methods=["GET"])
def api_numero_acta_validar():
    numero_acta = str(request.args.get("numero_acta") or "").strip()
    tipo_acta = str(request.args.get("tipo_acta") or "entrega").strip().lower()
    editing_acta_id = _parse_positive_int(request.args.get("editing_acta_id"))
    current_year = datetime.now().year

    if not NUMERO_ACTA_PATTERN.match(numero_acta):
        return jsonify({"success": True, "valid": False, "reason": "format"})

    seq_str, year_str = numero_acta.split("-", 1)
    seq = int(seq_str)
    year = int(year_str)
    if year != current_year:
        return jsonify({"success": True, "valid": False, "reason": "year"})

    max_for_year = get_max_numero_acta_for_year(current_year, tipo_acta=tipo_acta)
    db = get_db()
    exists_query = [
        "SELECT id",
        "FROM historial_actas",
        "WHERE LOWER(COALESCE(tipo_acta, '')) = LOWER(?)",
        "AND numero_acta = ?",
    ]
    exists_params = [tipo_acta, numero_acta]
    if editing_acta_id:
        exists_query.append("AND id != ?")
        exists_params.append(editing_acta_id)
    exists_query.extend(["LIMIT 1"])
    exists = bool(db.execute("\n".join(exists_query), tuple(exists_params)).fetchone())

    return jsonify(
        {
            "success": True,
            "valid": (not exists),
            "exists": exists,
            "lower_than_max": seq < max_for_year,
            "max_numero_acta": f"{max_for_year:03d}-{current_year}" if max_for_year > 0 else None,
            "next_numero_acta": get_next_numero_acta(current_year, tipo_acta=tipo_acta),
            "tipo_acta": tipo_acta,
        }
    )


@documents_bp.route("/api/informes/areas/numero-acta/siguiente", methods=["GET"])
def api_numero_informe_area_siguiente():
    current_year = datetime.now().year
    next_num = get_next_numero_informe_area(current_year)
    return jsonify({"success": True, "numero_acta": next_num, "year": current_year})


@documents_bp.route("/api/informes/areas/generar-lote", methods=["POST"])
def api_generar_informes_areas_lote():
    payload = request.json or {}
    scope = str(payload.get("scope") or "area").strip().lower()
    area_id = payload.get("area_id")
    piso_id = payload.get("piso_id")
    bloque_id = payload.get("bloque_id")

    template_path = os.path.join(UPLOAD_FOLDER, "aula.docx")
    if not os.path.exists(template_path):
        return jsonify({"success": False, "error": "No existe plantilla cargada para el tipo aula."}), 404

    try:
        target_areas = _get_target_areas_for_report(
            scope=scope,
            area_id=area_id,
            piso_id=piso_id,
            bloque_id=bloque_id,
        )
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)}), 400

    if not target_areas:
        return jsonify({"success": False, "error": "No se encontraron areas para el criterio seleccionado."}), 404

    current_year = datetime.now().year
    numeros_acta = reserve_numeros_informe_area(current_year, len(target_areas))

    # Flujo rápido para una sola área: sin cola, sin SSE de job, descarga directa DOCX.
    if len(target_areas) == 1:
        area = target_areas[0]
        numero_acta = numeros_acta[0]
        area_nombre = str(area.get("area_nombre") or "AREA SIN NOMBRE")
        table_data, total_bienes = _build_area_summary_rows(area.get("area_id"))

        context_data = {
            "numero_acta": numero_acta,
            "nombre_area": area_nombre,
            "piso": area.get("piso_nombre"),
            "bloque": area.get("bloque_nombre"),
            "total_bienes": total_bienes,
            "titulo_acta": f"{area_nombre} ACTA No.{numero_acta}",
        }
        context_data = _normalize_informe_context("aula", context_data)
        table_columns = [
            {"id": "descripcion", "label": "DESCRIPCION"},
            {"id": "cantidad", "label": "CANTIDAD"},
        ]

        single_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(AREA_REPORTS_OUTPUT_ROOT, f"single-{single_stamp}")
        os.makedirs(output_dir, exist_ok=True)

        doc_name = (
            f"acta-aula-{_safe_filename_part(area_nombre)}-"
            f"{_safe_filename_part(numero_acta)}"
        )
        docx_path, _pdf_path = generate_acta(
            template_path=template_path,
            context_data=context_data,
            table_data=table_data,
            table_columns=table_columns,
            output_dir=output_dir,
            doc_name=doc_name,
            generate_pdf=False,
            use_date_subfolder=False,
            include_time_suffix=False,
        )

        if not docx_path:
            _cleanup_area_reports_dir(output_dir)
            return jsonify(
                {
                    "success": False,
                    "error": f"No se pudo generar DOCX para el area {area_nombre}.",
                }
            ), 500

        next_numero = get_next_numero_informe_area(current_year)
        publish_event(
            "areas_reports_changed",
            {
                "scope": scope,
                "generated": 1,
                "next_numero_acta": next_numero,
            },
        )

        return jsonify(
            {
                "success": True,
                "scope": scope,
                "queued": False,
                "immediate": True,
                "total_targets": 1,
                "total_generated": 1,
                "download_path": docx_path,
                "download_kind": "docx",
                "start_numero_acta": numero_acta,
                "end_numero_acta": numero_acta,
                "next_numero_acta": next_numero,
            }
        )

    batch_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    lot_dir = os.path.join(AREA_REPORTS_OUTPUT_ROOT, f"lote-{batch_stamp}")
    os.makedirs(lot_dir, exist_ok=True)

    job_id = str(uuid.uuid4())
    _set_area_report_task_state(
        job_id,
        status="queued",
        progress=0,
        scope=scope,
        total=len(target_areas),
        generated=0,
        pause_requested=False,
        cancel_requested=False,
        created_at=datetime.now().isoformat(),
    )

    app_obj = getattr(current_app, "_get_current_object", lambda: current_app)()

    AREA_REPORT_EXECUTOR.submit(
        _run_area_report_job,
        app_obj,
        job_id,
        scope,
        target_areas,
        numeros_acta,
        lot_dir,
        template_path,
        current_year,
    )

    publish_event(
        "areas_reports_progress",
        {
            "job_id": job_id,
            "progress": 0,
            "generated": 0,
            "total": len(target_areas),
            "message": "Tarea en cola...",
        },
    )
    publish_event(
        "areas_reports_changed",
        {
            "scope": scope,
            "generated": 0,
            "next_numero_acta": get_next_numero_informe_area(current_year),
        },
    )

    return jsonify(
        {
            "success": True,
            "job_id": job_id,
            "scope": scope,
            "queued": True,
            "total_targets": len(target_areas),
            "start_numero_acta": numeros_acta[0],
            "end_numero_acta": numeros_acta[-1],
            "next_numero_acta": get_next_numero_informe_area(current_year),
        }
    ), 202


@documents_bp.route("/api/informes/areas/jobs/<job_id>", methods=["GET"])
def api_get_area_report_job(job_id):
    with AREA_REPORT_TASKS_LOCK:
        payload = dict(AREA_REPORT_TASKS.get(str(job_id), {}) or {})
    if not payload:
        return jsonify({"success": False, "error": "Job no encontrado."}), 404
    return jsonify({"success": True, "data": payload})


@documents_bp.route("/api/informes/areas/jobs", methods=["GET"])
def api_list_area_report_jobs():
    active_only = str(request.args.get("active") or "0").strip().lower() in {"1", "true", "yes"}
    active_statuses = {"queued", "running", "paused"}

    with AREA_REPORT_TASKS_LOCK:
        rows = []
        for jid, payload in AREA_REPORT_TASKS.items():
            item = dict(payload or {})
            item["job_id"] = str(jid)
            status = str(item.get("status") or "")
            if active_only and status not in active_statuses:
                continue
            rows.append(item)

    rows.sort(key=lambda x: str(x.get("created_at") or ""), reverse=True)
    return jsonify({"success": True, "jobs": rows})


@documents_bp.route("/api/informes/areas/jobs/<job_id>/control", methods=["POST"])
def api_control_area_report_job(job_id):
    payload = request.json or {}
    action = str(payload.get("action") or "").strip().lower()
    if action not in {"pause", "resume", "cancel"}:
        return jsonify({"success": False, "error": "Acción inválida. Use pause, resume o cancel."}), 400

    job_id = str(job_id)
    task = _get_area_report_task(job_id)
    if not task:
        return jsonify({"success": False, "error": "Job no encontrado."}), 404

    status = str(task.get("status") or "")
    terminal_statuses = {"completed", "error", "cancelled"}
    if status in terminal_statuses:
        return jsonify({"success": False, "error": f"El job ya está en estado final: {status}."}), 409

    if action == "pause":
        _set_area_report_task_state(job_id, pause_requested=True, message="Pausa solicitada...")
        publish_event("areas_reports_paused", {"job_id": job_id, "message": "Pausa solicitada..."})
        return jsonify({"success": True, "job_id": job_id, "status": "pausing"})

    if action == "resume":
        _set_area_report_task_state(job_id, pause_requested=False, status="running", message="Reanudando...")
        publish_event("areas_reports_resumed", {"job_id": job_id, "message": "Reanudando..."})
        return jsonify({"success": True, "job_id": job_id, "status": "running"})

    _set_area_report_task_state(job_id, cancel_requested=True, message="Cancelación solicitada...")
    publish_event("areas_reports_cancelled", {"job_id": job_id, "message": "Cancelación solicitada..."})
    return jsonify({"success": True, "job_id": job_id, "status": "cancelling"})


@documents_bp.route("/api/informes/traspaso/exportar", methods=["GET"])
def api_exportar_bienes_traspaso():
    db = get_db()
    rows = db.execute(
        """
        SELECT 
            t.cod_inventario, t.cod_esbye, t.descripcion, t.ubicacion AS ubicacion_original,
            t.marca, t.modelo, t.serie, t.estado, t.usuario_final,
            t.facultad_destino, t.fecha_traspaso, h.numero_acta
        FROM inventario_traspasos t
        LEFT JOIN historial_actas h ON t.acta_traspaso_id = h.id
        ORDER BY t.fecha_traspaso DESC
        """
    ).fetchall()

    if not rows:
        return jsonify({"success": False, "error": "No hay bienes registrados fuera de la facultad."}), 404

    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Bienes Fuera"
    
    # Encabezados
    headers = [
        "CÓDIGO INV.", "CÓD. ESBYE", "DESCRIPCIÓN", "UBICACIÓN ORIG.",
        "MARCA", "MODELO", "SERIE", "ESTADO", "USUARIO FINAL",
        "FACULTAD DESTINO", "FECHA TRASPASO", "NRO. ACTA"
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Datos
    for row in rows:
        ws.append([
            row["cod_inventario"], row["cod_esbye"], row["descripcion"], row["ubicacion_original"],
            row["marca"], row["modelo"], row["serie"], row["estado"], row["usuario_final"],
            row["facultad_destino"], row["fecha_traspaso"], row["numero_acta"]
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"bienes_fuera_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@documents_bp.route("/api/historial/<int:acta_id>", methods=["DELETE"])
def api_delete_historial(acta_id):
    deleted = delete_historial_acta(acta_id)

    if deleted:
        _safe_remove_file(deleted.get("docx_path"), is_output_path_allowed)
        _safe_remove_file(deleted.get("pdf_path"), is_output_path_allowed)

        snapshot_path = deleted.get("plantilla_snapshot_path")
        if snapshot_path:
            remaining_refs = count_historial_by_template_snapshot(snapshot_path)
            if remaining_refs <= 0 and _safe_remove_file(snapshot_path, is_template_history_path_allowed):
                cleanup_empty_parent_dirs(snapshot_path, TEMPLATE_HISTORY_ROOT)

    publish_event("actas_changed", {"acta_id": acta_id, "action": "delete"})
    return jsonify({"success": True})
