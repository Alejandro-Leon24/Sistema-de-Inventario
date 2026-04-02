import logging
import sqlite3

from flask import Blueprint, jsonify, request

from database.params_controller import (
    create_param,
    delete_param,
    get_param,
    get_universidad,
    set_universidad,
    update_param,
)

try:
    from app.utils.constants import VALID_PARAM_TYPES
except ModuleNotFoundError:
    from utils.constants import VALID_PARAM_TYPES


parameters_bp = Blueprint("parameters", __name__)
logger = logging.getLogger(__name__)


@parameters_bp.get("/api/parametros/<tipo>")
def api_get_parametros(tipo):
    if tipo not in VALID_PARAM_TYPES:
        return jsonify({"error": "Tipo de parámetro no válido."}), 400
    return jsonify({"data": get_param(tipo)})


@parameters_bp.post("/api/parametros/<tipo>")
def api_create_parametro(tipo):
    if tipo not in VALID_PARAM_TYPES:
        return jsonify({"error": "Tipo de parámetro no válido."}), 400
    payload = request.get_json(silent=True) or {}
    nombre = (payload.get("nombre") or "").strip()
    if not nombre:
        return jsonify({"error": "El nombre es obligatorio."}), 400
    try:
        param_id = create_param(tipo, nombre, payload.get("descripcion"))
    except sqlite3.IntegrityError:
        return jsonify({"error": "Ya existe un parámetro con ese nombre."}), 409
    return jsonify({"id": param_id}), 201


@parameters_bp.patch("/api/parametros/<tipo>/<int:param_id>")
def api_update_parametro(tipo, param_id):
    if tipo not in VALID_PARAM_TYPES:
        return jsonify({"error": "Tipo de parámetro no válido."}), 400
    payload = request.get_json(silent=True) or {}
    nombre = (payload.get("nombre") or "").strip()
    if not nombre:
        return jsonify({"error": "El nombre es obligatorio."}), 400
    try:
        ok = update_param(tipo, param_id, nombre, payload.get("descripcion"))
    except sqlite3.IntegrityError:
        return jsonify({"error": "Ya existe un parámetro con ese nombre."}), 409
    if not ok:
        return jsonify({"error": "Parámetro no encontrado."}), 404
    return jsonify({"success": True})


@parameters_bp.delete("/api/parametros/<tipo>/<int:param_id>")
def api_delete_parametro(tipo, param_id):
    if tipo not in VALID_PARAM_TYPES:
        return jsonify({"error": "Tipo de parámetro no válido."}), 400
    try:
        delete_param(tipo, param_id)
    except ValueError as error:
        return jsonify({"error": str(error)}), 409
    except Exception:
        logger.exception("Error no controlado eliminando parametro tipo=%s id=%s", tipo, param_id)
        return jsonify({"error": "No se pudo eliminar el parámetro."}), 500
    return jsonify({"success": True})


@parameters_bp.get("/api/universidad")
def api_get_universidad():
    return jsonify({"data": get_universidad()})


@parameters_bp.patch("/api/universidad")
def api_put_universidad():
    payload = request.get_json(silent=True) or {}
    for clave, valor in payload.items():
        set_universidad(clave, str(valor or ""))
    return jsonify({"success": True})
