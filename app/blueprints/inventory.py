import sqlite3

from flask import Blueprint, jsonify, request

from database.controller import (
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
