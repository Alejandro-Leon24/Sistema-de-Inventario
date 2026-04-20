"""Servicio de dominio de inventario.

Mantiene contratos estables para que repositorios y blueprints no dependan
 directamente del controlador legacy.
"""

from database import controller as legacy_controller

ALLOWED_INVENTORY_FIELDS = legacy_controller.ALLOWED_INVENTORY_FIELDS


def list_inventory_items(filters=None, sort_direction="asc", limit=5000):
    return legacy_controller.list_inventory_items(filters=filters, sort_direction=sort_direction, limit=limit)


def list_inventory_items_paginated(filters=None, sort_direction="asc", page=1, per_page=50):
    return legacy_controller.list_inventory_items_paginated(
        filters=filters,
        sort_direction=sort_direction,
        page=page,
        per_page=per_page,
    )


def get_inventory_item(item_id):
    return legacy_controller.get_inventory_item(item_id)


def create_inventory_item(payload, commit=True):
    return legacy_controller.create_inventory_item(payload, commit=commit)


def update_inventory_item(item_id, payload):
    return legacy_controller.update_inventory_item(item_id, payload)


def delete_inventory_item(item_id):
    return legacy_controller.delete_inventory_item(item_id)


def clear_inventory_items(reset_sequence=True):
    return legacy_controller.clear_inventory_items(reset_sequence=reset_sequence)


def find_inventory_code_duplicates(cod_inventario=None, cod_esbye=None, limit=50, exclude_item_id=None):
    return legacy_controller.find_inventory_code_duplicates(
        cod_inventario=cod_inventario,
        cod_esbye=cod_esbye,
        limit=limit,
        exclude_item_id=exclude_item_id,
    )


def bulk_insert_inventory_rows(rows, area_id=None, procedencia_default=None):
    return legacy_controller.bulk_insert_inventory_rows(
        rows,
        area_id=area_id,
        procedencia_default=procedencia_default,
    )


def bulk_insert_inventory_dicts(rows_as_dicts, area_id=None, procedencia_default=None):
    return legacy_controller.bulk_insert_inventory_dicts(
        rows_as_dicts,
        area_id=area_id,
        procedencia_default=procedencia_default,
    )


def iter_inventory_items(filters=None, sort_direction="asc", batch_size=2000):
    return legacy_controller.iter_inventory_items(filters=filters, sort_direction=sort_direction, batch_size=batch_size)


def get_user_preferences(user_key):
    return legacy_controller.get_user_preferences(user_key)


def set_user_preference(user_key, pref_key, pref_value):
    return legacy_controller.set_user_preference(user_key, pref_key, pref_value)


def get_column_mappings():
    return legacy_controller.get_column_mappings()


def replace_column_mappings(mappings):
    return legacy_controller.replace_column_mappings(mappings)


def get_inventory_search_diagnostics(search_text=None):
    return legacy_controller.get_inventory_search_diagnostics(search_text=search_text)
