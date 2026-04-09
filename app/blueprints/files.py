import os
from datetime import datetime

from flask import Blueprint, jsonify, request, send_file, send_from_directory

from database.controller import get_all_areas_for_export, iter_inventory_items

try:
    from app.utils.constants import AREA_EXPORT_COLUMNS, INVENTORY_EXPORT_COLUMNS
    from app.utils.excel_export import generar_excel
    from app.utils.filesystem import cleanup_empty_parent_dirs
except ModuleNotFoundError:
    from utils.constants import AREA_EXPORT_COLUMNS, INVENTORY_EXPORT_COLUMNS
    from utils.excel_export import generar_excel
    from utils.filesystem import cleanup_empty_parent_dirs

from .documents import ACTAS_OUTPUT_ROOT, AREA_REPORTS_OUTPUT_ROOT, PREVIEW_OUTPUT_ROOT, is_output_path_allowed


files_bp = Blueprint("files", __name__)


def _is_temporary_download_path(path):
    if not path:
        return False
    real_path = os.path.abspath(path)
    reports_root = os.path.abspath(AREA_REPORTS_OUTPUT_ROOT)
    actas_root = os.path.abspath(ACTAS_OUTPUT_ROOT)
    preview_root = os.path.abspath(PREVIEW_OUTPUT_ROOT)
    return (
        real_path.startswith(reports_root)
        or real_path.startswith(actas_root)
        or real_path.startswith(preview_root)
    )


def _cleanup_parent_dirs_for_temp_path(path):
    if not path:
        return
    real_path = os.path.abspath(path)
    reports_root = os.path.abspath(AREA_REPORTS_OUTPUT_ROOT)
    actas_root = os.path.abspath(ACTAS_OUTPUT_ROOT)
    preview_root = os.path.abspath(PREVIEW_OUTPUT_ROOT)

    stop_at = None
    if real_path.startswith(reports_root):
        stop_at = reports_root
    elif real_path.startswith(actas_root):
        stop_at = actas_root
    elif real_path.startswith(preview_root):
        stop_at = preview_root

    if stop_at:
        cleanup_empty_parent_dirs(real_path, stop_at)


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
        response = send_file(path, as_attachment=True)

        if _is_temporary_download_path(path):
            real_path = os.path.abspath(path)

            @response.call_on_close
            def _cleanup_temp_download():
                try:
                    if os.path.exists(real_path):
                        os.remove(real_path)
                except Exception:
                    pass
                _cleanup_parent_dirs_for_temp_path(real_path)

        return response
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
