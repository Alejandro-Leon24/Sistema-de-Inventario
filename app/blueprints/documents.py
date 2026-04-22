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
    numero_acta_exists,
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


import sys
from flask import Blueprint, jsonify, request, current_app

documents_bp = Blueprint("documents", __name__)
logger = logging.getLogger(__name__)

NUMERO_ACTA_PATTERN = re.compile(r"^[A-Z0-9-]{1,20}-\d{4}$")

# Manejo de rutas para entorno empaquetado (PyInstaller)
if getattr(sys, 'frozen', False):
    # Carpeta donde está el .exe
    BASE_DIR = Path(sys.executable).parent
    # Carpeta interna del bundle
    BUNDLE_DIR = Path(sys._MEIPASS)
else:
    BASE_DIR = Path(__file__).resolve().parents[2]
    BUNDLE_DIR = BASE_DIR

# Las plantillas se pueden leer del bundle, pero permitimos que existan fuera también
UPLOAD_FOLDER = os.path.join(BASE_DIR, "plantillas")
if not os.path.exists(UPLOAD_FOLDER) and getattr(sys, 'frozen', False):
    UPLOAD_FOLDER = os.path.join(BUNDLE_DIR, "plantillas")

ACTAS_OUTPUT_ROOT = os.path.join(BASE_DIR, "salidas", "inventario", "actas")
PREVIEW_OUTPUT_ROOT = os.environ.get(
    "INVENTARIO_PREVIEW_ROOT",
    os.path.join(tempfile.gettempdir(), "inventario_preview"),
)
TEMPLATE_HISTORY_ROOT = os.path.join(UPLOAD_FOLDER, "_historial")
AREA_REPORTS_OUTPUT_ROOT = os.path.join(ACTAS_OUTPUT_ROOT, "informes-area")

# Asegurar que las carpetas existan en el BASE_DIR real
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
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
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


_ACTA_INVENTORY_STATE_FIELDS = ["area_id", "ubicacion", "estado", "justificacion", "procedencia"]


def _snapshot_inventory_item_state(db, item_id):
    row = db.execute(
        "SELECT id, area_id, ubicacion, estado, justificacion, procedencia FROM inventario_items WHERE id = ?",
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
        if kind not in {"update", "insert", "transfer_out"}:
            continue
        item_id = _parse_positive_int(entry.get("item_id"))
        if (kind == "insert" or kind == "transfer_out") and not item_id:
            continue
        before_json = json.dumps(entry.get("before"), ensure_ascii=False) if entry.get("before") is not None else None
        after_json = json.dumps(entry.get("after"), ensure_ascii=False) if entry.get("after") is not None else None
        db.execute(
            """
            INSERT INTO acta_inventory_mutaciones (
                acta_id, tipo_acta, item_id, mutation_kind, before_data_json, after_data_json
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
                db.execute("DELETE FROM inventario_items WHERE id = ?", (item_id,))
            continue

        if kind == "transfer_out":
            traspaso_row = db.execute("SELECT * FROM inventario_traspasos WHERE id = ?", (item_id,)).fetchone()
            if traspaso_row:
                data = dict(traspaso_row)
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
        if not before_raw: continue
        try:
            before_data = json.loads(before_raw)
        except: continue

        db.execute(
            """
            UPDATE inventario_items
            SET area_id = ?, ubicacion = ?, estado = ?, justificacion = ?, procedencia = ?, actualizado_en = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (
                before_data.get("area_id"), before_data.get("ubicacion"), before_data.get("estado"),
                before_data.get("justificacion"), before_data.get("procedencia"), item_id,
            ),
        )


def _clear_acta_inventory_mutations(db, acta_id):
    db.execute("DELETE FROM acta_inventory_mutaciones WHERE acta_id = ?", (int(acta_id),))


def _normalize_informe_context(tipo, context_data):
    ctx = dict(context_data or {})
    for key in list(ctx.keys()):
        normalized_key = str(key).lower()
        if normalized_key == "fecha" or normalized_key.startswith("fecha_") or normalized_key.endswith("_fecha"):
            ctx[key] = _format_fecha_es_larga(ctx.get(key))

    tipo_norm = str(tipo or "").strip().lower()
    if not ctx.get("memorandum") and ctx.get("memorando"):
        ctx["memorandum"] = ctx.get("memorando")
    if tipo_norm == "recepcion":
        if not ctx.get("fecha_elaboracion") and ctx.get("fecha_emision"):
            ctx["fecha_elaboracion"] = ctx.get("fecha_emision")
    return ctx


def _parse_positive_int(value):
    try:
        parsed = int(value)
        return parsed if parsed > 0 else None
    except: return None


def _resolve_area_from_form(context_data):
    area_id = _parse_positive_int(context_data.get("ubicacion_area_id"))
    if not area_id: return None
    row = get_db().execute("SELECT id FROM areas WHERE id = ? LIMIT 1", (area_id,)).fetchone()
    return int(row["id"]) if row else None


def _get_target_areas_for_report(scope, area_id=None, piso_id=None, bloque_id=None):
    db = get_db()
    scope_norm = str(scope or "area").strip().lower()
    params = []
    where = ""
    if scope_norm == "area":
        if not area_id: raise ValueError("Debe seleccionar un area.")
        where, params = "WHERE a.id = ?", [int(area_id)]
    elif scope_norm == "piso":
        if not piso_id: raise ValueError("Debe seleccionar un piso.")
        where, params = "WHERE p.id = ?", [int(piso_id)]
    elif scope_norm == "bloque":
        if not bloque_id: raise ValueError("Debe seleccionar un bloque.")
        where, params = "WHERE b.id = ?", [int(bloque_id)]
    else: raise ValueError("Scope invalido.")

    rows = db.execute(
        f"""
        SELECT a.id AS area_id, a.nombre AS area_nombre, p.id AS piso_id, p.nombre AS piso_nombre,
               b.id AS bloque_id, b.nombre AS bloque_nombre
        FROM areas a JOIN pisos p ON p.id = a.piso_id JOIN bloques b ON b.id = p.bloque_id
        {where} ORDER BY b.nombre, p.nombre, a.nombre
        """,
        tuple(params),
    ).fetchall()
    return [dict(row) for row in rows]


def _build_area_summary_rows(area_id):
    db = get_db()
    rows = db.execute(
        """
        SELECT COALESCE(NULLIF(TRIM(descripcion), ''), 'SIN DESCRIPCION') AS descripcion, SUM(COALESCE(cantidad, 1)) AS cantidad
        FROM inventario_items WHERE area_id = ? GROUP BY 1 ORDER BY 1
        """,
        (int(area_id),),
    ).fetchall()
    table_data = [{"descripcion": str(r["descripcion"]), "cantidad": int(r["cantidad"] or 0)} for r in rows]
    total = sum(int(item["cantidad"] or 0) for item in table_data)
    table_data.append({"descripcion": "TOTAL DE BIENES", "cantidad": total})
    return table_data, total


def _set_area_report_task_state(job_id, **updates):
    with AREA_REPORT_TASKS_LOCK:
        AREA_REPORT_TASKS.setdefault(job_id, {}).update(updates)


def _run_area_report_job(app, job_id, scope, target_areas, numeros_acta, lot_dir, template_path, current_year):
    generated = []
    try:
        with app.app_context():
            total = len(target_areas)
            for idx, area in enumerate(target_areas):
                # Check status
                state = _get_area_report_task(job_id)
                if state.get("cancel_requested"):
                    _set_area_report_task_state(job_id, status="cancelled")
                    for p in generated: _safe_remove_file(p, is_output_path_allowed)
                    return
                while _get_area_report_task(job_id).get("pause_requested"): time.sleep(0.5)

                num = numeros_acta[idx]
                rows, total_bienes = _build_area_summary_rows(area["area_id"])
                ctx = _normalize_informe_context("aula", {
                    "numero_acta": num, "nombre_area": area["area_nombre"], "piso": area["piso_nombre"],
                    "bloque": area["bloque_nombre"], "total_bienes": total_bienes,
                    "titulo_acta": f"{area['area_nombre']} ACTA No.{num}"
                })
                path = generate_acta(
                    template_path=template_path, context_data=ctx, table_data=rows,
                    table_columns=[{"id": "descripcion", "label": "DESCRIPCION"}, {"id": "cantidad", "label": "CANTIDAD"}],
                    output_dir=lot_dir, doc_name=f"acta-aula-{_safe_filename_part(area['area_nombre'])}-{num}",
                    use_date_subfolder=False, include_time_suffix=False
                )
                if path: generated.append(path)
                prog = int(((idx+1)/total)*100)
                _set_area_report_task_state(job_id, status="running", progress=prog, generated=idx+1, total=total)
                publish_event("areas_reports_progress", {"job_id": job_id, "progress": prog, "generated": idx+1, "total": total})

            download_path = generated[0] if len(generated) == 1 else os.path.join(lot_dir, "informes-area.zip")
            if len(generated) > 1:
                with zipfile.ZipFile(download_path, "w", zipfile.ZIP_DEFLATED) as z:
                    for p in generated: z.write(p, os.path.basename(p))
                for p in generated: _safe_remove_file(p, is_output_path_allowed)

            _set_area_report_task_state(job_id, status="completed", progress=100, download_path=download_path, download_kind="docx" if len(generated)==1 else "zip")
            publish_event("areas_reports_ready", {"job_id": job_id, "download_path": download_path})
    except Exception as e:
        logger.exception("Error en job informes-area")
        _set_area_report_task_state(job_id, status="error", error=str(e))
        publish_event("areas_reports_error", {"job_id": job_id, "error": str(e)})


def is_output_path_allowed(path):
    if not path: return False
    p = os.path.abspath(path)
    return p.startswith(os.path.abspath(ACTAS_OUTPUT_ROOT)) or p.startswith(os.path.abspath(PREVIEW_OUTPUT_ROOT))


def is_template_history_path_allowed(path):
    return path and os.path.abspath(path).startswith(os.path.abspath(TEMPLATE_HISTORY_ROOT))


def _safe_remove_file(path, allowed_checker):
    try:
        if path and allowed_checker(path) and os.path.exists(path):
            os.remove(path)
            return True
    except: pass
    return False


def _snapshot_template_for_historial(tipo, template_path):
    tipo_safe = werkzeug.utils.secure_filename(str(tipo or "general")) or "general"
    with open(template_path, "rb") as f: data = f.read()
    h = hashlib.sha1(data).hexdigest()
    d = os.path.join(TEMPLATE_HISTORY_ROOT, tipo_safe)
    os.makedirs(d, exist_ok=True)
    snap = os.path.join(d, f"{h}.docx")
    if not os.path.exists(snap): shutil.copy2(template_path, snap)
    return h, snap, extract_variables_from_template(snap)


def _client_preview_key(req):
    ip = (req.headers.get("X-Forwarded-For") or "").split(",")[0].strip() or req.remote_addr or "local"
    return re.sub(r"[^A-Za-z0-9_.-]", "_", ip)


@documents_bp.route("/api/plantillas/upload", methods=["POST"])
def api_upload_plantilla():
    if "documento" not in request.files: return jsonify({"success": False, "error": "No file"}), 400
    f = request.files["documento"]
    t = request.form.get("tipo", "general")
    if not f.filename.endswith(".docx"): return jsonify({"success": False, "error": "Invalid file"}), 400
    p = os.path.join(UPLOAD_FOLDER, werkzeug.utils.secure_filename(f"{t}.docx"))
    f.save(p)
    publish_event("templates_changed", {"tipo": t})
    return jsonify({"success": True, "variables": extract_variables_from_template(p)})


@documents_bp.route("/api/plantillas/estado", methods=["GET"])
def api_estado_plantillas():
    t = request.args.get("tipo")
    p = os.path.join(UPLOAD_FOLDER, f"{t}.docx")
    exists = os.path.exists(p)
    return jsonify({"success": True, "existe": exists, "variables": extract_variables_from_template(p) if exists else []})


@documents_bp.route("/api/informes/generar", methods=["POST"])
def api_generar_informe():
    payload = request.json
    tipo = payload.get("tipo", "entrega")
    ctx_raw = dict(payload.get("datos_formulario", {}) or {})
    ctx = _normalize_informe_context(tipo, ctx_raw)
    vista_previa = bool(payload.get("vista_previa", False)) or str(request.headers.get("X-Preview-Request")).lower() in {"1", "true"}
    current_year = datetime.now().year
    tipo_norm = str(tipo).strip().lower()
    editing_id = _parse_positive_int(payload.get("editing_acta_id"))
    
    logger.info(f"Generando informe tipo={tipo_norm}, editing_id={editing_id}, numero_acta={ctx.get('numero_acta')}")
    
    req_num = str(ctx.get("numero_acta") or "").strip()
    numero_acta = req_num or get_next_numero_acta(current_year, tipo_acta=tipo_norm)

    if not NUMERO_ACTA_PATTERN.match(numero_acta):
        return jsonify({"success": False, "error": "Formato de acta inválido."}), 400

    tabla_data = payload.get("datos_tabla", [])
    tabla_cols = payload.get("datos_columnas", [])
    tpl_p = os.path.join(UPLOAD_FOLDER, f"{tipo}.docx")
    if not os.path.exists(tpl_p): return jsonify({"success": False, "error": "No hay plantilla."}), 404

    try:
        if not vista_previa:
            numero_acta = reserve_numero_acta(current_year, preferred_numero_acta=numero_acta if req_num else None, tipo_acta=tipo_norm, editing_acta_id=editing_id)
            ctx["numero_acta"] = numero_acta
            out_d = os.path.join(ACTAS_OUTPUT_ROOT, f"acta-{tipo}")
            doc_n = f"acta-{tipo}-{_safe_filename_part(numero_acta)}"
        else:
            out_d = os.path.join(PREVIEW_OUTPUT_ROOT, f"acta-{tipo}", _client_preview_key(request))
            doc_n = f"preview-{tipo}"

        docx_p = generate_acta(tpl_p, ctx, tabla_data, tabla_cols, out_d, doc_n, False, bool(vista_previa))
        if not docx_p: return jsonify({"success": False, "error": "Error generando DOCX."}), 500
        html = render_docx_preview_html(docx_p) if vista_previa else None

        if not vista_previa:
            db = get_db()
            if editing_id:
                _revert_acta_inventory_mutations(db, editing_id)
                _clear_acta_inventory_mutations(db, editing_id)
            
            target_area = _resolve_area_from_form(ctx)
            pending = []
            procedencia_text = _build_acta_procedencia_text(tipo_norm, numero_acta)

            # 1. Guardar primero el acta en el historial (esto valida duplicados y hace commit del acta)
            t_h, t_s, _ = _snapshot_template_for_historial(tipo, tpl_p)
            full_data_json = json.dumps({"formulario": ctx_raw, "tabla": tabla_data, "columnas": tabla_cols})
            if editing_id:
                update_historial_acta(editing_id, tipo, full_data_json, docx_p, None, numero_acta, t_h, t_s)
                acta_id = editing_id
            else:
                acta_id = save_historial_acta(tipo, full_data_json, docx_p, None, numero_acta, t_h, t_s)

            # 2. Solo si el acta se guardó bien, realizamos las mutaciones en el inventario
            if tipo_norm == "entrega" and target_area:
                for r in tabla_data:
                    iid = _parse_positive_int(r.get("id"))
                    if iid:
                        bef = _snapshot_inventory_item_state(db, iid)
                        db.execute("UPDATE inventario_items SET area_id = ?, ubicacion = ?, procedencia = ?, actualizado_en = CURRENT_TIMESTAMP WHERE id = ?",
                                   (target_area, ctx.get("area_trabajo"), procedencia_text, iid))
                        pending.append({"kind": "update", "item_id": iid, "before": bef, "after": _snapshot_inventory_item_state(db, iid)})
            
            elif tipo_norm == "recepcion" and target_area:
                row_next = db.execute("SELECT COALESCE(MAX(item_numero), 0) + 1 AS next FROM inventario_items").fetchone()
                next_item_num = int(row_next["next"]) if row_next else 1
                for idx, row in enumerate(tabla_data or []):
                    db.execute(
                        """
                        INSERT INTO inventario_items (
                            item_numero, cod_inventario, cod_esbye, cuenta, cantidad, descripcion,
                            ubicacion, marca, modelo, serie, estado, condicion, usuario_final,
                            fecha_adquisicion, valor, observacion, justificacion, procedencia, area_id,
                            descripcion_esbye, marca_esbye, modelo_esbye, serie_esbye,
                            valor_esbye, ubicacion_esbye, observacion_esbye, fecha_adquisicion_esbye
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            next_item_num + idx,
                            str(row.get("cod_inventario") or "").strip() or None,
                            str(row.get("cod_esbye") or "").strip() or None,
                            str(row.get("cuenta") or "").strip() or None,
                            int(row.get("cantidad") or 1),
                            str(row.get("descripcion") or "").strip() or None,
                            str(ctx.get("area_trabajo") or "").strip() or None,
                            str(row.get("marca") or "").strip() or None,
                            str(row.get("modelo") or "").strip() or None,
                            str(row.get("serie") or "").strip() or None,
                            str(row.get("estado") or "").strip() or None,
                            str(row.get("condicion") or "").strip() or None,
                            str(row.get("usuario_final") or "").strip() or None,
                            str(row.get("fecha_adquisicion") or "").strip() or None,
                            row.get("valor"),
                            str(row.get("observacion") or "").strip() or None,
                            str(row.get("justificacion") or "").strip() or None,
                            procedencia_text, target_area,
                            str(row.get("descripcion_esbye") or "").strip() or None,
                            str(row.get("marca_esbye") or "").strip() or None,
                            str(row.get("modelo_esbye") or "").strip() or None,
                            str(row.get("serie_esbye") or "").strip() or None,
                            row.get("valor_esbye"),
                            str(row.get("ubicacion_esbye") or "").strip() or None,
                            str(row.get("observacion_esbye") or "").strip() or None,
                            str(row.get("fecha_adquisicion_esbye") or "").strip() or None
                        ),
                    )
                    ins_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]
                    pending.append({"kind": "insert", "item_id": ins_id, "before": None, "after": _snapshot_inventory_item_state(db, ins_id)})

            elif tipo_norm == "traspaso":
                facultad_recibe = str(ctx.get("facultad_recibe") or "OTRA FACULTAD").strip()
                for row in tabla_data:
                    iid = _parse_positive_int(row.get("id"))
                    full_row = db.execute("SELECT * FROM inventario_items WHERE id = ?", (iid,)).fetchone() if iid else None
                    if full_row:
                        data = dict(full_row)
                        db.execute(
                            """
                            INSERT INTO inventario_traspasos (
                                id, item_numero, cod_inventario, cod_esbye, cuenta, cantidad, descripcion,
                                ubicacion, marca, modelo, serie, estado, condicion, usuario_final,
                                fecha_adquisicion, valor, observacion, justificacion, procedencia, area_id,
                                facultad_destino, fecha_traspaso, datos_originales_json
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, ?)
                            """,
                            (
                                data["id"], data["item_numero"], data["cod_inventario"], data["cod_esbye"],
                                data["cuenta"], data["cantidad"], data["descripcion"], data["ubicacion"],
                                data["marca"], data["modelo"], data["serie"], data["estado"],
                                data["condicion"], data["usuario_final"], data["fecha_adquisicion"],
                                data["valor"], data["observacion"], data["justificacion"],
                                data["procedencia"], data["area_id"], facultad_recibe, json.dumps(data)
                            )
                        )
                        db.execute("DELETE FROM inventario_items WHERE id = ?", (iid,))
                        pending.append({"kind": "transfer_out", "item_id": iid, "before": data, "after": {"status": "transferred", "to": facultad_recibe}})

            elif tipo_norm in {"baja", "bajas"}:
                for row in tabla_data:
                    iid = _parse_positive_int(row.get("id"))
                    if iid:
                        bef = _snapshot_inventory_item_state(db, iid)
                        db.execute("UPDATE inventario_items SET estado = ?, justificacion = ?, procedencia = ?, actualizado_en = CURRENT_TIMESTAMP WHERE id = ?",
                                   (str(row.get("estado") or "MALO"), str(row.get("justificacion") or ""), procedencia_text, iid))
                        pending.append({"kind": "update", "item_id": iid, "before": bef, "after": _snapshot_inventory_item_state(db, iid)})

            # 3. Persistir log de mutaciones y commit final
            if pending: _persist_acta_inventory_mutations(db, acta_id, tipo_norm, pending)
            if tipo_norm == "traspaso": db.execute("UPDATE inventario_traspasos SET acta_traspaso_id = ? WHERE acta_traspaso_id IS NULL", (acta_id,))
            db.commit()
            publish_event("actas_changed", {"tipo": tipo})


        return jsonify({"success": True, "docx_path": docx_p, "numero_acta": numero_acta, "html_preview": html})
    except Exception as e:
        logger.exception("Error generando informe")
        return jsonify({"success": False, "error": str(e)}), 500


@documents_bp.route("/api/historial", methods=["GET"])
def api_get_historial_all():
    return jsonify({"success": True, "data": get_historial_actas(request.args.get("tipo_acta"))})


@documents_bp.route("/api/historial/numero-acta/siguiente", methods=["GET"])
def api_numero_acta_siguiente():
    t = str(request.args.get("tipo_acta") or "entrega").strip().lower()
    return jsonify({"success": True, "numero_acta": get_next_numero_acta(datetime.now().year, tipo_acta=t)})


@documents_bp.route("/api/historial/numero-acta/validar", methods=["GET"])
def api_numero_acta_validar():
    num = str(request.args.get("numero_acta") or "").strip()
    t = str(request.args.get("tipo_acta") or "entrega").strip().lower()
    eid = _parse_positive_int(request.args.get("editing_acta_id"))
    year = datetime.now().year
    if not NUMERO_ACTA_PATTERN.match(num): return jsonify({"success": True, "valid": False, "reason": "format"})
    seq_s, year_s = num.split("-", 1)
    if int(year_s) != year: return jsonify({"success": True, "valid": False, "reason": "year"})
    db = get_db()
    if eid:
        exists = bool(db.execute(
            "SELECT 1 FROM historial_actas WHERE numero_acta=? AND LOWER(tipo_acta)=LOWER(?) AND id!=? LIMIT 1", 
            (num, t, eid)
        ).fetchone())
    else:
        exists = numero_acta_exists(num, tipo_acta=t)
    max_v = get_max_numero_acta_for_year(year, tipo_acta=t)
    return jsonify({"success": True, "valid": not exists, "exists": exists, "lower_than_max": int(seq_s) < max_v, 
                    "max_numero_acta": f"{max_v:03d}-{year}" if max_v>0 else None,
                    "next_numero_acta": get_next_numero_acta(year, tipo_acta=t)})


@documents_bp.route("/api/informes/areas/numero-acta/siguiente", methods=["GET"])
def api_numero_informe_area_siguiente():
    return jsonify({"success": True, "numero_acta": get_next_numero_informe_area(datetime.now().year)})


@documents_bp.route("/api/informes/areas/generar-lote", methods=["POST"])
def api_generar_lote_aula():
    payload = request.json or {}
    scope, aid, pid, bid = payload.get("scope", "area"), payload.get("area_id"), payload.get("piso_id"), payload.get("bloque_id")
    tpl = os.path.join(UPLOAD_FOLDER, "aula.docx")
    if not os.path.exists(tpl): return jsonify({"success": False, "error": "No hay plantilla aula."}), 404
    try:
        areas = _get_target_areas_for_report(scope, aid, pid, bid)
    except Exception as e: return jsonify({"success": False, "error": str(e)}), 400
    if not areas: return jsonify({"success": False, "error": "No hay areas."}), 404
    y = datetime.now().year
    nums = reserve_numeros_informe_area(y, len(areas))
    lot_dir = os.path.join(AREA_REPORTS_OUTPUT_ROOT, f"lote-{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    os.makedirs(lot_dir, exist_ok=True)
    jid = str(uuid.uuid4())
    _set_area_report_task_state(jid, status="queued", progress=0, total=len(areas), generated=0)
    AREA_REPORT_EXECUTOR.submit(_run_area_report_job, current_app._get_current_object(), jid, scope, areas, nums, lot_dir, tpl, y)
    return jsonify({"success": True, "job_id": jid}), 202


@documents_bp.route("/api/informes/areas/jobs/<job_id>", methods=["GET"])
def api_get_job(job_id):
    t = _get_area_report_task(job_id)
    return jsonify({"success": True, "data": t}) if t else (jsonify({"success": False}), 404)


@documents_bp.route("/api/informes/areas/jobs", methods=["GET"])
def api_list_jobs():
    with AREA_REPORT_TASKS_LOCK:
        return jsonify({"success": True, "jobs": list(AREA_REPORT_TASKS.values())})


@documents_bp.route("/api/informes/areas/jobs/<job_id>/control", methods=["POST"])
def api_control_job(job_id):
    action = request.json.get("action")
    if action == "cancel": _set_area_report_task_state(job_id, cancel_requested=True)
    elif action == "pause": _set_area_report_task_state(job_id, pause_requested=True)
    elif action == "resume": _set_area_report_task_state(job_id, pause_requested=False, status="running")
    return jsonify({"success": True})


@documents_bp.route("/api/informes/traspaso/exportar", methods=["GET"])
def api_traspaso_exportar():
    rows = get_db().execute("SELECT t.*, h.numero_acta FROM inventario_traspasos t LEFT JOIN historial_actas h ON t.acta_traspaso_id=h.id").fetchall()
    if not rows: return jsonify({"success": False, "error": "No hay datos"}), 404
    import io
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(["CÓDIGO INV.", "CÓD. ESBYE", "DESCRIPCIÓN", "DESTINO", "FECHA TRASPASO", "NRO. ACTA"])
    for r in rows: ws.append([r["cod_inventario"], r["cod_esbye"], r["descripcion"], r["facultad_destino"], r["fecha_traspaso"], r["numero_acta"]])
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, as_attachment=True, download_name="traspasos.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@documents_bp.route("/api/historial/<int:acta_id>", methods=["GET"])
def api_get_historial_by_id(acta_id):
    from database.historial_repository import get_historial_acta_by_id
    acta = get_historial_acta_by_id(acta_id)
    if not acta:
        return jsonify({"success": False, "error": "Acta no encontrada"}), 404
    return jsonify({"success": True, "data": acta})


@documents_bp.route("/api/historial/<int:acta_id>", methods=["DELETE"])
def api_delete_historial(acta_id):
    deleted = delete_historial_acta(acta_id)
    if deleted: _safe_remove_file(deleted.get("docx_path"), is_output_path_allowed)
    publish_event("actas_changed", {"acta_id": acta_id, "action": "delete"})
    return jsonify({"success": True})
