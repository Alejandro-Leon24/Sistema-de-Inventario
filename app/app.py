import sqlite3
import uuid
import shutil
from io import BytesIO

from docx import Document

from flask import Flask, jsonify, render_template, request, send_file
from datetime import datetime
from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from database.controller import (
    bulk_insert_inventory_rows,
    create_area,
    create_block,
    create_floor,
    delete_block,
    delete_area,
    delete_floor,
    create_inventory_item,
    delete_inventory_item,
    find_inventory_code_duplicates,
    get_column_mappings,
    get_inventory_item,
    get_location_dependency_summary,
    get_personas,
    get_structure,
    get_user_preferences,
    init_schema,
    list_inventory_items,
    list_inventory_items_paginated,
    replace_column_mappings,
    set_user_preference,
    update_area,
    update_block,
    update_floor,
    update_inventory_item,
)
from database.params_controller import (
    get_param,
    create_param,
    delete_param,
    update_param,
    get_universidad,
    set_universidad,
    get_administradores,
    create_administrador,
    update_administrador,
    delete_administrador,
    get_administrador_dependency_summary,
)
from database.db import init_app


def _resolve_database_path(base_dir: Path) -> Path:
    target_db = base_dir / "inventario.db"
    legacy_db = base_dir / "prueba.db"

    if target_db.exists():
        return target_db

    if legacy_db.exists():
        try:
            legacy_db.replace(target_db)
        except OSError:
            shutil.copy2(legacy_db, target_db)

    return target_db

app = Flask(
    __name__
)
app.config["DATABASE"] = _resolve_database_path(BASE_DIR)
init_app(app)
DEFAULT_USER_KEY = "portable_user"
with app.app_context():
    init_schema(BASE_DIR)

@app.before_request
def before_request():
    # Aquí puedes agregar lógica que se ejecute antes de cada solicitud
    pass

@app.after_request
def after_request(response):
    # Aquí puedes agregar lógica que se ejecute después de cada solicitud
    return response


@app.route('/')
def index():
    from database.controller import get_dashboard_stats
    personas = get_personas()
    stats = get_dashboard_stats()
    data = {
        'year': datetime.now().year,
        'personas': personas,
        **stats
    }
    return render_template('index.html', data=data)


@app.route('/inventario-form')
def inventario_form():
    data = {
        'year': datetime.now().year,
    }
    return render_template('inventario-form.html', data=data)

@app.route('/inventario-list')
def inventario_list():
    data = {
        'year': datetime.now().year,
    }
    return render_template('inventario-list.html', data=data)


@app.route('/ajustes')
def ajustes():
    data = {
        'year': datetime.now().year,
    }
    return render_template('ajustes.html', data=data)

@app.route('/informe')
def informe():
    data = {
        'year': datetime.now().year,
    }
    return render_template('informe.html', data=data)


@app.get('/api/estructura')
def api_estructura():
    return jsonify({"data": get_structure()})


@app.post('/api/bloques')
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


@app.patch('/api/bloques/<int:block_id>')
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


@app.delete('/api/bloques/<int:block_id>')
def api_delete_bloque(block_id):
    ok = delete_block(block_id)
    if not ok:
        return jsonify({"error": "Bloque no encontrado."}), 404
    return jsonify({"success": True})


@app.post('/api/pisos')
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


@app.post('/api/areas')
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
            # Evita conflictos de unicidad por piso sin alterar el esquema.
            nombre = f"AREA SIN NOMBRE {uuid.uuid4().hex[:8].upper()}"

    details_keys = [
        "identificacion_ambiente", "metros_cuadrados", "alto", "senaletica", "cod_senaletica",
        "infraestructura_fisica", "estado_piso", "material_techo", "puerta", "material_puerta",
        "responsable_admin_id", "estado_paredes", "estado_techo", "estado_puerta", "cerradura",
        "nivel_seguridad", "sitio_profesor_mesa", "sitio_profesor_silla", "pc_aula", "proyector",
        "pantalla_interactiva", "pupitres_cantidad", "pupitres_funcionan", "pupitres_no_funcionan",
        "pizarra", "pizarra_estado", "ventanas_cantidad", "ventanas_funcionan", "ventanas_no_funcionan",
        "aa_cantidad", "aa_funcionan", "aa_no_funcionan", "ventiladores_cantidad",
        "ventiladores_funcionan", "ventiladores_no_funcionan", "wifi",
        "red_lan", "red_lan_funcionan", "red_lan_no_funcionan", "red_inalambrica_cantidad",
        "iluminacion_funcionan", "iluminacion_no_funcionan", "luminarias_cantidad", "puntos_electricos",
        "puntos_electricos_funcionan", "puntos_electricos_no_funcionan", "puntos_electricos_cantidad",
        "capacidad_aulica", "capacidad_distanciamiento", "ambiente_apto_retorno", "observaciones_detalle",
    ]
    details = {key: payload.get(key) for key in details_keys if key in payload}

    try:
        area_id = create_area(piso_id, nombre, payload.get("descripcion"), details=details)
    except sqlite3.IntegrityError:
        return jsonify({"error": "Ya existe un área con ese nombre en el piso seleccionado."}), 409
    return jsonify({"id": area_id}), 201


@app.patch('/api/pisos/<int:floor_id>')
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


@app.delete('/api/pisos/<int:floor_id>')
def api_delete_piso(floor_id):
    ok = delete_floor(floor_id)
    if not ok:
        return jsonify({"error": "Piso no encontrado."}), 404
    return jsonify({"success": True})


@app.patch('/api/areas/<int:area_id>')
def api_update_area(area_id):
    payload = request.get_json(silent=True) or {}
    nombre = payload.get("nombre")
    descripcion = payload.get("descripcion")
    if nombre is not None:
        nombre = str(nombre).strip() or None
    details_keys = [
        "identificacion_ambiente", "metros_cuadrados", "alto", "senaletica", "cod_senaletica",
        "infraestructura_fisica", "estado_piso", "material_techo", "puerta", "material_puerta",
        "responsable_admin_id", "estado_paredes", "estado_techo", "estado_puerta", "cerradura",
        "nivel_seguridad", "sitio_profesor_mesa", "sitio_profesor_silla", "pc_aula", "proyector",
        "pantalla_interactiva", "pupitres_cantidad", "pupitres_funcionan", "pupitres_no_funcionan",
        "pizarra", "pizarra_estado", "ventanas_cantidad", "ventanas_funcionan", "ventanas_no_funcionan",
        "aa_cantidad", "aa_funcionan", "aa_no_funcionan", "ventiladores_cantidad",
        "ventiladores_funcionan", "ventiladores_no_funcionan", "wifi",
        "red_lan", "red_lan_funcionan", "red_lan_no_funcionan", "red_inalambrica_cantidad",
        "iluminacion_funcionan", "iluminacion_no_funcionan", "luminarias_cantidad", "puntos_electricos",
        "puntos_electricos_funcionan", "puntos_electricos_no_funcionan", "puntos_electricos_cantidad",
        "capacidad_aulica", "capacidad_distanciamiento", "ambiente_apto_retorno", "observaciones_detalle",
    ]
    details = {key: payload.get(key) for key in details_keys if key in payload}

    try:
        ok = update_area(area_id, nombre=nombre, descripcion=descripcion, details=details)
    except sqlite3.IntegrityError:
        return jsonify({"error": "Ya existe un área con ese nombre en el piso seleccionado."}), 409
    if not ok:
        return jsonify({"error": "Área no encontrada."}), 404
    return jsonify({"success": True})


@app.delete('/api/areas/<int:area_id>')
def api_delete_area(area_id):
    ok = delete_area(area_id)
    if not ok:
        return jsonify({"error": "Área no encontrada."}), 404
    return jsonify({"success": True})


@app.get('/api/ubicaciones/impacto')
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


@app.get('/api/inventario')
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
    return jsonify({
        "data": result["items"],
        "pagination": {
            "page": result["page"],
            "per_page": result["per_page"],
            "total": result["total"],
            "total_pages": result["total_pages"],
        },
    })


@app.get('/api/ubicaciones/export')
def api_export_areas():
    from database.controller import get_all_areas_for_export
    from utils.constants import AREA_EXPORT_COLUMNS
    from utils.excel_export import generar_excel
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


@app.get('/api/inventario/export')
def api_export_inventario():
    from utils.constants import INVENTORY_EXPORT_COLUMNS
    from utils.excel_export import generar_excel
    filters = {
        "bloque_id": request.args.get("bloque_id", type=int),
        "piso_id": request.args.get("piso_id", type=int),
        "area_id": request.args.get("area_id", type=int),
        "search": request.args.get("search", type=str),
    }
    sort_direction = request.args.get("order", default="asc", type=str)
    items = list_inventory_items(filters=filters, sort_direction=sort_direction)

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


@app.get('/api/inventario/<int:item_id>')
def api_get_inventario(item_id):
    item = get_inventory_item(item_id)
    if not item:
        return jsonify({"error": "Elemento no encontrado."}), 404
    return jsonify({"data": item})


@app.post('/api/inventario')
def api_create_inventario():
    payload = request.get_json(silent=True) or {}
    force_duplicate = bool(payload.get("force_duplicate"))
    duplicate_items = find_inventory_code_duplicates(
        cod_inventario=payload.get("cod_inventario"),
        cod_esbye=payload.get("cod_esbye"),
        limit=50,
    )
    if duplicate_items and not force_duplicate:
        return jsonify({
            "error": "Código repetido detectado. Confirma si deseas agregarlo de todas formas.",
            "duplicates": duplicate_items,
        }), 409
    try:
        item_id = create_inventory_item(payload)
    except sqlite3.IntegrityError as error:
        duplicate_items = find_inventory_code_duplicates(
            cod_inventario=payload.get("cod_inventario"),
            cod_esbye=payload.get("cod_esbye"),
            limit=50,
        )
        return jsonify({
            "error": f"No se pudo guardar por un valor duplicado: {error}",
            "duplicates": duplicate_items,
        }), 409
    item = get_inventory_item(item_id)
    return jsonify({"data": item}), 201


@app.patch('/api/inventario/<int:item_id>')
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
            return jsonify({
                "error": "Código repetido detectado. Confirma si deseas guardar el cambio de todas formas.",
                "duplicates": duplicate_items,
            }), 409
    try:
        ok = update_inventory_item(item_id, payload)
    except sqlite3.IntegrityError as error:
        duplicate_items = find_inventory_code_duplicates(
            cod_inventario=payload.get("cod_inventario") if "cod_inventario" in payload else None,
            cod_esbye=payload.get("cod_esbye") if "cod_esbye" in payload else None,
            limit=50,
            exclude_item_id=item_id,
        )
        return jsonify({
            "error": f"No se pudo actualizar por un valor duplicado: {error}",
            "duplicates": duplicate_items,
        }), 409
    if not ok:
        return jsonify({"error": "Elemento no encontrado."}), 404
    item = get_inventory_item(item_id)
    return jsonify({"data": item})


@app.delete('/api/inventario/<int:item_id>')
def api_delete_inventario(item_id):
    ok = delete_inventory_item(item_id)
    if not ok:
        return jsonify({"error": "Elemento no encontrado."}), 404
    return jsonify({"success": True})


@app.post('/api/inventario/pegar')
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


@app.get('/api/preferencias')
def api_get_preferencias():
    return jsonify({"data": get_user_preferences(DEFAULT_USER_KEY)})


@app.patch('/api/preferencias')
def api_set_preferencias():
    payload = request.get_json(silent=True) or {}
    pref_key = (payload.get("pref_key") or "").strip()
    if not pref_key:
        return jsonify({"error": "pref_key es obligatorio."}), 400
    set_user_preference(DEFAULT_USER_KEY, pref_key, payload.get("pref_value"))
    return jsonify({"success": True})


@app.get('/api/column-mappings')
def api_get_column_mappings():
    return jsonify({"data": get_column_mappings()})


@app.patch('/api/column-mappings')
def api_put_column_mappings():
    payload = request.get_json(silent=True) or {}
    mappings = payload.get("mappings") or []
    if not isinstance(mappings, list):
        return jsonify({"error": "mappings debe ser una lista."}), 400
    replace_column_mappings(mappings)
    return jsonify({"success": True})


@app.get('/api/parametros/<tipo>')
def api_get_parametros(tipo):
    if tipo not in ["estados", "condiciones", "cuentas", "si_no", "estado_puerta", "cerraduras", "estado_piso", "material_techo", "material_puerta", "estado_pizarra"]:
        return jsonify({"error": "Tipo de parámetro no válido."}), 400
    return jsonify({"data": get_param(tipo)})


@app.post('/api/parametros/<tipo>')
def api_create_parametro(tipo):
    if tipo not in ["estados", "condiciones", "cuentas", "si_no", "estado_puerta", "cerraduras", "estado_piso", "material_techo", "material_puerta", "estado_pizarra"]:
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


@app.patch('/api/parametros/<tipo>/<int:param_id>')
def api_update_parametro(tipo, param_id):
    if tipo not in ["estados", "condiciones", "cuentas", "si_no", "estado_puerta", "cerraduras", "estado_piso", "material_techo", "material_puerta", "estado_pizarra"]:
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


@app.delete('/api/parametros/<tipo>/<int:param_id>')
def api_delete_parametro(tipo, param_id):
    if tipo not in ["estados", "condiciones", "cuentas", "si_no", "estado_puerta", "cerraduras", "estado_piso", "material_techo", "material_puerta", "estado_pizarra"]:
        return jsonify({"error": "Tipo de parámetro no válido."}), 400
    try:
        delete_param(tipo, param_id)
    except ValueError as error:
        return jsonify({"error": str(error)}), 409
    except Exception:
        return jsonify({"error": "No se pudo eliminar el parámetro."}), 500
    return jsonify({"success": True})


@app.get('/api/universidad')
def api_get_universidad():
    return jsonify({"data": get_universidad()})


@app.patch('/api/universidad')
def api_put_universidad():
    payload = request.get_json(silent=True) or {}
    for clave, valor in payload.items():
        set_universidad(clave, str(valor or ""))
    return jsonify({"success": True})


@app.get('/api/administradores')
def api_get_administradores():
    return jsonify({"data": get_administradores()})


@app.post('/api/administradores')
def api_create_administrador_route():
    payload = request.get_json(silent=True) or {}
    try:
        admin_id = create_administrador(payload)
    except sqlite3.IntegrityError as error:
        return jsonify({"error": f"Error de integridad: {error}"}), 409
    return jsonify({"id": admin_id}), 201


@app.patch('/api/administradores/<int:admin_id>')
def api_update_administrador_route(admin_id):
    payload = request.get_json(silent=True) or {}
    try:
        update_administrador(admin_id, payload)
    except sqlite3.IntegrityError as error:
        return jsonify({"error": f"Error de integridad: {error}"}), 409
    return jsonify({"success": True})


@app.delete('/api/administradores/<int:admin_id>')
def api_delete_administrador_route(admin_id):
    try:
        delete_administrador(admin_id)
    except Exception as error:
        return jsonify({"error": f"Error al eliminar: {error}"}), 500
    return jsonify({"success": True})


@app.get('/api/administradores/<int:admin_id>/impacto')
def api_administradores_impacto(admin_id):
    try:
        summary = get_administrador_dependency_summary(admin_id)
        return jsonify({"data": summary})
    except Exception as error:
        return jsonify({"error": str(error)}), 500


def pagina_no_encontrada(error):
    return render_template('404.html'), 404


app.register_error_handler(404, pagina_no_encontrada)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True, threaded=True)