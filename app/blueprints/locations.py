import sqlite3
import uuid

from flask import Blueprint, jsonify, request

from database.location_repository import (
    create_area,
    create_block,
    create_floor,
    delete_area,
    delete_block,
    delete_floor,
    get_location_dependency_summary,
    get_structure,
    update_area,
    update_block,
    update_floor,
)

try:
    from app.utils.constants import AREA_DETAILS_KEYS
except ModuleNotFoundError:
    from utils.constants import AREA_DETAILS_KEYS


locations_bp = Blueprint("locations", __name__)


@locations_bp.get("/api/estructura")
def api_estructura():
    raw_include_details = str(request.args.get("include_details", "0") or "0").strip().lower()
    include_details = raw_include_details in {"1", "true", "si", "yes", "on"}
    return jsonify({"data": get_structure(include_area_details=include_details)})


@locations_bp.post("/api/bloques")
def api_create_bloque():
    payload = request.get_json(silent=True) or {}
    nombre = (payload.get("nombre") or "").strip()
    if not nombre:
        return jsonify({"error": "El nombre del bloque es obligatorio."}), 400
    try:
        block_id = create_block(nombre, payload.get("descripcion"))
    except sqlite3.IntegrityError:
        return jsonify({"error": "Ya existe un bloque con ese nombre."}), 409
    return jsonify({"id": block_id}), 201


@locations_bp.patch("/api/bloques/<int:block_id>")
def api_update_bloque(block_id):
    payload = request.get_json(silent=True) or {}
    nombre = payload.get("nombre")
    descripcion = payload.get("descripcion")
    if nombre is not None and not str(nombre).strip():
        return jsonify({"error": "El nombre del bloque no puede estar vacío."}), 400
    try:
        ok = update_block(block_id, nombre=nombre, descripcion=descripcion)
    except sqlite3.IntegrityError:
        return jsonify({"error": "Ya existe un bloque con ese nombre."}), 409
    if not ok:
        return jsonify({"error": "Bloque no encontrado."}), 404
    return jsonify({"success": True})


@locations_bp.delete("/api/bloques/<int:block_id>")
def api_delete_bloque(block_id):
    ok = delete_block(block_id)
    if not ok:
        return jsonify({"error": "Bloque no encontrado."}), 404
    return jsonify({"success": True})


@locations_bp.post("/api/pisos")
def api_create_piso():
    payload = request.get_json(silent=True) or {}
    nombre = (payload.get("nombre") or "").strip()
    bloque_id = payload.get("bloque_id")
    if not bloque_id:
        return jsonify({"error": "El bloque es obligatorio."}), 400
    if not nombre:
        return jsonify({"error": "El nombre del piso es obligatorio."}), 400
    try:
        floor_id = create_floor(bloque_id, nombre, payload.get("descripcion"))
    except sqlite3.IntegrityError:
        return jsonify({"error": "Ya existe un piso con ese nombre en el bloque seleccionado."}), 409
    return jsonify({"id": floor_id}), 201


@locations_bp.patch("/api/pisos/<int:floor_id>")
def api_update_piso(floor_id):
    payload = request.get_json(silent=True) or {}
    nombre = payload.get("nombre")
    descripcion = payload.get("descripcion")
    if nombre is not None and not str(nombre).strip():
        return jsonify({"error": "El nombre del piso no puede estar vacío."}), 400
    try:
        ok = update_floor(floor_id, nombre=nombre, descripcion=descripcion)
    except sqlite3.IntegrityError:
        return jsonify({"error": "Ya existe un piso con ese nombre en el bloque seleccionado."}), 409
    if not ok:
        return jsonify({"error": "Piso no encontrado."}), 404
    return jsonify({"success": True})


@locations_bp.delete("/api/pisos/<int:floor_id>")
def api_delete_piso(floor_id):
    ok = delete_floor(floor_id)
    if not ok:
        return jsonify({"error": "Piso no encontrado."}), 404
    return jsonify({"success": True})


@locations_bp.post("/api/areas")
def api_create_area():
    payload = request.get_json(silent=True) or {}
    nombre = (payload.get("nombre") or "").strip()
    identificacion_ambiente = (payload.get("identificacion_ambiente") or "").strip()
    piso_id = payload.get("piso_id")
    if not piso_id:
        return jsonify({"error": "El piso es obligatorio."}), 400
    if not nombre:
        if identificacion_ambiente:
            nombre = identificacion_ambiente
        else:
            nombre = f"AREA SIN NOMBRE {uuid.uuid4().hex[:8].upper()}"

    details = {key: payload.get(key) for key in AREA_DETAILS_KEYS if key in payload}

    try:
        area_id = create_area(piso_id, nombre, payload.get("descripcion"), details=details)
    except sqlite3.IntegrityError:
        return jsonify({"error": "Ya existe un área con ese nombre en el piso seleccionado."}), 409
    return jsonify({"id": area_id}), 201


@locations_bp.patch("/api/areas/<int:area_id>")
def api_update_area(area_id):
    payload = request.get_json(silent=True) or {}
    nombre = payload.get("nombre")
    descripcion = payload.get("descripcion")
    if nombre is not None:
        nombre = str(nombre).strip() or None
    details = {key: payload.get(key) for key in AREA_DETAILS_KEYS if key in payload}

    try:
        ok = update_area(area_id, nombre=nombre, descripcion=descripcion, details=details)
    except sqlite3.IntegrityError:
        return jsonify({"error": "Ya existe un área con ese nombre en el piso seleccionado."}), 409
    if not ok:
        return jsonify({"error": "Área no encontrada."}), 404
    return jsonify({"success": True})


@locations_bp.delete("/api/areas/<int:area_id>")
def api_delete_area(area_id):
    ok = delete_area(area_id)
    if not ok:
        return jsonify({"error": "Área no encontrada."}), 404
    return jsonify({"success": True})


@locations_bp.get("/api/ubicaciones/impacto")
def api_ubicaciones_impacto():
    entity = request.args.get("entity", type=str)
    entity_id = request.args.get("id", type=int)
    if not entity or not entity_id:
        return jsonify({"error": "Parámetros entity e id son obligatorios."}), 400
    if entity not in ["bloque", "piso", "area"]:
        return jsonify({"error": "entity debe ser bloque, piso o area."}), 400
    try:
        summary = get_location_dependency_summary(entity, entity_id)
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    return jsonify({"data": summary})
