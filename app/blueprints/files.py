import os
from datetime import datetime

from flask import Blueprint, jsonify, request, send_file, send_from_directory

from database.controller import get_all_areas_for_export, iter_inventory_items

try:
    from app.utils.constants import AREA_EXPORT_COLUMNS, INVENTORY_EXPORT_COLUMNS
    from app.utils.excel_export import generar_excel
except ModuleNotFoundError:
    from utils.constants import AREA_EXPORT_COLUMNS, INVENTORY_EXPORT_COLUMNS
    from utils.excel_export import generar_excel

from .documents import ACTAS_OUTPUT_ROOT, is_output_path_allowed


files_bp = Blueprint("files", __name__)


@files_bp.get("/api/ubicaciones/export")
def api_export_areas():
    items = get_all_areas_for_export()

    try:
        output = generar_excel(items, AREA_EXPORT_COLUMNS, "Áreas")
    except ImportError:
        return jsonify({"error": "Para exportar a Excel se requiere openpyxl instalado."}), 500

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"ubicaciones_{timestamp}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@files_bp.get("/api/inventario/export")
def api_export_inventario():
    filters = {
        "bloque_id": request.args.get("bloque_id", type=int),
        "piso_id": request.args.get("piso_id", type=int),
        "area_id": request.args.get("area_id", type=int),
        "search": request.args.get("search", type=str),
    }
    sort_direction = request.args.get("order", default="asc", type=str)
    items = iter_inventory_items(filters=filters, sort_direction=sort_direction, batch_size=2000)

    try:
        output = generar_excel(items, INVENTORY_EXPORT_COLUMNS, "Inventario")
    except ImportError:
        return jsonify({"error": "Para exportar a Excel se requiere openpyxl instalado."}), 500

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"inventario_{timestamp}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@files_bp.route("/api/descargar", methods=["GET"])
def api_descargar():
    path = request.args.get("path")
    if path and os.path.exists(path):
        if not is_output_path_allowed(path):
            return "Ruta no permitida", 403
        return send_file(path, as_attachment=True)
    return "No encontrado", 404


@files_bp.route("/api/ver", methods=["GET"])
def api_ver_archivo():
    path = request.args.get("path")
    if path and os.path.exists(path):
        if not is_output_path_allowed(path):
            return "Ruta no permitida", 403
        return send_file(path, as_attachment=False)
    return "No encontrado", 404


@files_bp.route("/files/<path:filename>")
def serve_temp_files(filename):
    return send_from_directory(ACTAS_OUTPUT_ROOT, filename)
