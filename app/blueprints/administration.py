import logging
import sqlite3

from flask import Blueprint, jsonify, request

from database.controller import get_personal
from database.params_controller import (
    create_administrador,
    delete_administrador,
    get_administrador_dependency_summary,
    get_administradores,
    update_administrador,
)


administration_bp = Blueprint("administration", __name__)
logger = logging.getLogger(__name__)


@administration_bp.get("/api/administradores")
def api_get_administradores():
    return jsonify({"data": get_administradores()})


@administration_bp.post("/api/administradores")
def api_create_administrador_route():
    payload = request.get_json(silent=True) or {}
    try:
        admin_id = create_administrador(payload)
    except sqlite3.IntegrityError as error:
        return jsonify({"error": f"Error de integridad: {error}"}), 409
    return jsonify({"id": admin_id}), 201


@administration_bp.patch("/api/administradores/<int:admin_id>")
def api_update_administrador_route(admin_id):
    payload = request.get_json(silent=True) or {}
    try:
        update_administrador(admin_id, payload)
    except sqlite3.IntegrityError as error:
        return jsonify({"error": f"Error de integridad: {error}"}), 409
    return jsonify({"success": True})


@administration_bp.delete("/api/administradores/<int:admin_id>")
def api_delete_administrador_route(admin_id):
    try:
        delete_administrador(admin_id)
    except Exception as error:
        logger.exception("Error no controlado eliminando administrador id=%s", admin_id)
        return jsonify({"error": f"Error al eliminar: {error}"}), 500
    return jsonify({"success": True})


@administration_bp.get("/api/administradores/<int:admin_id>/impacto")
def api_administradores_impacto(admin_id):
    try:
        summary = get_administrador_dependency_summary(admin_id)
        return jsonify({"data": summary})
    except Exception as error:
        logger.exception("Error no controlado consultando impacto de administrador id=%s", admin_id)
        return jsonify({"error": str(error)}), 500


@administration_bp.get("/api/personal")
def api_get_personal():
    try:
        data = get_personal()
        return jsonify({"success": True, "data": data})
    except Exception as error:
        logger.exception("Error no controlado consultando personal")
        return jsonify({"success": False, "error": str(error)}), 500
