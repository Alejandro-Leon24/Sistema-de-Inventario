import logging
import os
import re
import sqlite3
import tempfile
import time
import uuid

from flask import Blueprint, jsonify, request

from database.controller import (
    ALLOWED_INVENTORY_FIELDS,
    bulk_insert_inventory_dicts,
    bulk_insert_inventory_rows,
    create_inventory_item,
    delete_inventory_item,
    find_inventory_code_duplicates,
    get_column_mappings,
    get_inventory_item,
    get_inventory_search_diagnostics,
    list_inventory_items_paginated,
    replace_column_mappings,
    set_user_preference,
    get_user_preferences,
    update_inventory_item,
)


inventory_bp = Blueprint("inventory", __name__)
DEFAULT_USER_KEY = "portable_user"
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Excel import helpers
# ---------------------------------------------------------------------------
_EXCEL_IMPORT_DIR = os.path.join(tempfile.gettempdir(), "inventario_excel_import")
os.makedirs(_EXCEL_IMPORT_DIR, exist_ok=True)

_MAX_EXCEL_PREVIEW_ROWS = 20
_MAX_EXCEL_IMPORT_ROWS = 10_000
_MAX_EXCEL_FILE_BYTES = 10 * 1024 * 1024  # 10 MB

# Strict UUID v4 pattern – prevents any path traversal attempt
_UUID_RE = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$",
    re.IGNORECASE,
)


def _clean_old_excel_imports(max_age_seconds: int = 3600) -> None:
    """Remove temp excel files older than *max_age_seconds*."""
    now = time.time()
    try:
        for fname in os.listdir(_EXCEL_IMPORT_DIR):
            if not fname.endswith(".xlsx"):
                continue
            fpath = os.path.join(_EXCEL_IMPORT_DIR, fname)
            try:
                if os.path.isfile(fpath) and (now - os.path.getmtime(fpath)) > max_age_seconds:
                    os.remove(fpath)
            except Exception:
                pass
    except Exception:
        pass


def _safe_import_session_path(session_id: str) -> str | None:
    """Return the absolute temp path for *session_id*, or None if invalid.

    Only accepts well-formed UUID strings so no path separator characters or
    directory traversal sequences can reach ``os.path.join``.
    """
    if not _UUID_RE.fullmatch(session_id):
        return None
    safe_dir = os.path.abspath(_EXCEL_IMPORT_DIR)
    path = os.path.join(safe_dir, f"{session_id}.xlsx")
    # Defense-in-depth: confirm the resolved path stays inside safe_dir
    if not os.path.abspath(path).startswith(safe_dir + os.sep):
        return None
    return path


@inventory_bp.get("/api/inventario")
def api_list_inventario():
    filters = {
        "bloque_id": request.args.get("bloque_id", type=int),
        "piso_id": request.args.get("piso_id", type=int),
        "area_id": request.args.get("area_id", type=int),
        "search": request.args.get("search", type=str),
    }
    sort_direction = request.args.get("order", default="asc", type=str)
    page = request.args.get("page", default=1, type=int)
    per_page = request.args.get("per_page", default=50, type=int)
    per_page = max(1, min(per_page, 500))

    result = list_inventory_items_paginated(
        filters=filters,
        sort_direction=sort_direction,
        page=page,
        per_page=per_page,
    )
    return jsonify(
        {
            "data": result["items"],
            "pagination": {
                "page": result["page"],
                "per_page": result["per_page"],
                "total": result["total"],
                "total_pages": result["total_pages"],
            },
        }
    )


@inventory_bp.get("/api/inventario/<int:item_id>")
def api_get_inventario(item_id):
    item = get_inventory_item(item_id)
    if not item:
        return jsonify({"error": "Elemento no encontrado."}), 404
    return jsonify({"data": item})


@inventory_bp.post("/api/inventario")
def api_create_inventario():
    payload = request.get_json(silent=True) or {}
    force_duplicate = bool(payload.get("force_duplicate"))
    duplicate_items = find_inventory_code_duplicates(
        cod_inventario=payload.get("cod_inventario"),
        cod_esbye=payload.get("cod_esbye"),
        limit=50,
    )
    if duplicate_items and not force_duplicate:
        return jsonify(
            {
                "error": "Código repetido detectado. Confirma si deseas agregarlo de todas formas.",
                "duplicates": duplicate_items,
            }
        ), 409
    try:
        item_id = create_inventory_item(payload)
    except sqlite3.IntegrityError as error:
        duplicate_items = find_inventory_code_duplicates(
            cod_inventario=payload.get("cod_inventario"),
            cod_esbye=payload.get("cod_esbye"),
            limit=50,
        )
        return jsonify(
            {
                "error": f"No se pudo guardar por un valor duplicado: {error}",
                "duplicates": duplicate_items,
            }
        ), 409
    item = get_inventory_item(item_id)
    return jsonify({"data": item}), 201


@inventory_bp.patch("/api/inventario/<int:item_id>")
def api_update_inventario(item_id):
    payload = request.get_json(silent=True) or {}
    force_duplicate = bool(payload.get("force_duplicate"))
    if "cod_inventario" in payload or "cod_esbye" in payload:
        duplicate_items = find_inventory_code_duplicates(
            cod_inventario=payload.get("cod_inventario") if "cod_inventario" in payload else None,
            cod_esbye=payload.get("cod_esbye") if "cod_esbye" in payload else None,
            limit=50,
            exclude_item_id=item_id,
        )
        if duplicate_items and not force_duplicate:
            return jsonify(
                {
                    "error": "Código repetido detectado. Confirma si deseas guardar el cambio de todas formas.",
                    "duplicates": duplicate_items,
                }
            ), 409
    try:
        ok = update_inventory_item(item_id, payload)
    except sqlite3.IntegrityError as error:
        duplicate_items = find_inventory_code_duplicates(
            cod_inventario=payload.get("cod_inventario") if "cod_inventario" in payload else None,
            cod_esbye=payload.get("cod_esbye") if "cod_esbye" in payload else None,
            limit=50,
            exclude_item_id=item_id,
        )
        return jsonify(
            {
                "error": f"No se pudo actualizar por un valor duplicado: {error}",
                "duplicates": duplicate_items,
            }
        ), 409
    if not ok:
        return jsonify({"error": "Elemento no encontrado."}), 404
    item = get_inventory_item(item_id)
    return jsonify({"data": item})


@inventory_bp.delete("/api/inventario/<int:item_id>")
def api_delete_inventario(item_id):
    ok = delete_inventory_item(item_id)
    if not ok:
        return jsonify({"error": "Elemento no encontrado."}), 404
    return jsonify({"success": True})


@inventory_bp.post("/api/inventario/pegar")
def api_paste_inventario():
    payload = request.get_json(silent=True) or {}
    rows = payload.get("rows") or []
    if not isinstance(rows, list) or not rows:
        return jsonify({"error": "Debe enviar una lista de filas para pegar."}), 400
    try:
        inserted_ids = bulk_insert_inventory_rows(rows, payload.get("area_id"))
    except sqlite3.IntegrityError as error:
        return jsonify({"error": f"Error al pegar datos por duplicado: {error}"}), 409
    return jsonify({"inserted_ids": inserted_ids}), 201


@inventory_bp.get("/api/inventario/search-diagnostics")
def api_inventory_search_diagnostics():
    search_text = request.args.get("search", default="", type=str)
    data = get_inventory_search_diagnostics(search_text)
    return jsonify({"data": data})


@inventory_bp.get("/api/inventario/duplicados")
def api_inventory_duplicates():
    cod_inventario = request.args.get("cod_inventario", default="", type=str)
    cod_esbye = request.args.get("cod_esbye", default="", type=str)
    duplicates = find_inventory_code_duplicates(
        cod_inventario=cod_inventario,
        cod_esbye=cod_esbye,
        limit=50,
    )
    return jsonify({"success": True, "duplicates": duplicates})


@inventory_bp.get("/api/preferencias")
def api_get_preferencias():
    return jsonify({"data": get_user_preferences(DEFAULT_USER_KEY)})


@inventory_bp.patch("/api/preferencias")
def api_set_preferencias():
    payload = request.get_json(silent=True) or {}
    pref_key = (payload.get("pref_key") or "").strip()
    if not pref_key:
        return jsonify({"error": "pref_key es obligatorio."}), 400
    set_user_preference(DEFAULT_USER_KEY, pref_key, payload.get("pref_value"))
    return jsonify({"success": True})


@inventory_bp.get("/api/column-mappings")
def api_get_column_mappings():
    return jsonify({"data": get_column_mappings()})


@inventory_bp.patch("/api/column-mappings")
def api_put_column_mappings():
    payload = request.get_json(silent=True) or {}
    mappings = payload.get("mappings") or []
    if not isinstance(mappings, list):
        return jsonify({"error": "mappings debe ser una lista."}), 400
    replace_column_mappings(mappings)
    return jsonify({"success": True})


# ---------------------------------------------------------------------------
# Excel bulk import  –  step 1: upload and preview
# ---------------------------------------------------------------------------

@inventory_bp.post("/api/inventario/previsualizar-excel")
def api_previsualizar_excel():
    """Upload an .xlsx file, store it temporarily and return headers + preview rows."""
    _clean_old_excel_imports()

    if "file" not in request.files:
        return jsonify({"error": "No se recibió ningún archivo."}), 400

    file = request.files["file"]
    filename = (file.filename or "").strip().lower()
    if not filename.endswith(".xlsx"):
        return jsonify({"error": "Solo se aceptan archivos .xlsx"}), 400

    content = file.read()
    if len(content) > _MAX_EXCEL_FILE_BYTES:
        return jsonify({"error": "El archivo supera el límite de 10 MB."}), 400
    if not content:
        return jsonify({"error": "El archivo está vacío."}), 400

    session_id = str(uuid.uuid4())
    temp_path = os.path.join(_EXCEL_IMPORT_DIR, f"{session_id}.xlsx")
    with open(temp_path, "wb") as f:
        f.write(content)

    try:
        from openpyxl import load_workbook  # openpyxl is an existing dependency

        wb = load_workbook(temp_path, read_only=True, data_only=True)
        ws = wb.active
        headers: list[str] = []
        preview_rows: list[list[str]] = []
        total_rows = 0

        for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
            if row_idx == 0:
                headers = [str(cell if cell is not None else "").strip() for cell in row]
            else:
                if any(cell is not None for cell in row):
                    total_rows += 1
                    if total_rows <= _MAX_EXCEL_PREVIEW_ROWS:
                        preview_rows.append(
                            [str(cell) if cell is not None else "" for cell in row]
                        )
        wb.close()
    except Exception:
        logger.exception("Error reading uploaded Excel file")
        try:
            os.remove(temp_path)
        except Exception:
            pass
        return jsonify({"error": "No se pudo leer el archivo Excel. Verifica que sea un .xlsx válido y no esté dañado."}), 400

    if not headers:
        try:
            os.remove(temp_path)
        except Exception:
            pass
        return jsonify({"error": "El archivo no tiene encabezados en la primera fila."}), 400

    return jsonify(
        {
            "session_id": session_id,
            "headers": headers,
            "preview_rows": preview_rows,
            "total_rows": total_rows,
        }
    )


# ---------------------------------------------------------------------------
# Excel bulk import  –  step 2: apply mapping and insert
# ---------------------------------------------------------------------------

@inventory_bp.post("/api/inventario/confirmar-importacion")
def api_confirmar_importacion():
    """Apply a column mapping to a previously uploaded file and bulk-insert the rows."""
    payload = request.get_json(silent=True) or {}
    session_id = str(payload.get("session_id") or "").strip()
    # mapping: {str(col_index): canonical_field_name_or_empty_string}
    mapping: dict[str, str] = payload.get("mapping") or {}
    area_id_raw = payload.get("area_id")

    if not session_id:
        return jsonify({"error": "Falta session_id."}), 400

    temp_path = _safe_import_session_path(session_id)
    if temp_path is None:
        return jsonify({"error": "session_id inválido."}), 400
    if not os.path.exists(temp_path):
        return jsonify(
            {"error": "La sesión de importación ha expirado. Por favor sube el archivo nuevamente."}
        ), 410

    try:
        area_id = int(area_id_raw) if area_id_raw not in (None, "", "null") else None
    except (TypeError, ValueError):
        area_id = None

    try:
        from openpyxl import load_workbook

        wb = load_workbook(temp_path, read_only=True, data_only=True)
        ws = wb.active
        rows_as_dicts: list[dict] = []
        first_row = True

        for row in ws.iter_rows(values_only=True):
            if first_row:
                first_row = False
                continue
            if len(rows_as_dicts) >= _MAX_EXCEL_IMPORT_ROWS:
                break
            if not any(cell is not None for cell in row):
                continue  # skip blank rows
            row_dict: dict = {}
            for col_idx, cell in enumerate(row):
                canonical = str(mapping.get(str(col_idx)) or "").strip()
                if canonical and canonical in ALLOWED_INVENTORY_FIELDS and canonical != "item_numero":
                    row_dict[canonical] = str(cell).strip() if cell is not None else None
            if row_dict:
                rows_as_dicts.append(row_dict)

        wb.close()
    except Exception:
        logger.exception("Error reading Excel file during import session_id=%s", session_id)
        return jsonify({"error": "Error al leer el archivo para importar. Por favor sube el archivo nuevamente."}), 500
    finally:
        try:
            os.remove(temp_path)
        except Exception:
            pass

    if not rows_as_dicts:
        return jsonify({"error": "No se encontraron filas con datos para importar después de aplicar el mapeo."}), 400

    try:
        result = bulk_insert_inventory_dicts(rows_as_dicts, area_id=area_id)
    except sqlite3.IntegrityError:
        logger.exception("Integrity error during Excel bulk import")
        return jsonify({"error": "Error de integridad al importar. Puede haber datos con formato inválido."}), 409

    return jsonify(
        {
            "success": True,
            "inserted": result["inserted"],
            "skipped": result["skipped"],
        }
    ), 201
