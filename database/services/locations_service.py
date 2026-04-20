"""Servicio de dominio para estructura de ubicaciones."""

from database.db import get_db

AREA_DETAIL_COLUMNS = [
    "identificacion_ambiente",
    "metros_cuadrados",
    "alto",
    "senaletica",
    "cod_senaletica",
    "infraestructura_fisica",
    "estado_piso",
    "material_techo",
    "puerta",
    "material_puerta",
    "responsable_admin_id",
    "estado_paredes",
    "estado_techo",
    "estado_puerta",
    "cerradura",
    "nivel_seguridad",
    "sitio_profesor_mesa",
    "sitio_profesor_silla",
    "pc_aula",
    "proyector",
    "pantalla_interactiva",
    "pupitres_cantidad",
    "pupitres_funcionan",
    "pupitres_no_funcionan",
    "pizarra",
    "pizarra_estado",
    "ventanas_cantidad",
    "ventanas_funcionan",
    "ventanas_no_funcionan",
    "aa_cantidad",
    "aa_funcionan",
    "aa_no_funcionan",
    "ventiladores_cantidad",
    "ventiladores_funcionan",
    "ventiladores_no_funcionan",
    "wifi",
    "red_lan",
    "red_lan_funcionan",
    "red_lan_no_funcionan",
    "red_inalambrica_cantidad",
    "iluminacion_funcionan",
    "iluminacion_no_funcionan",
    "luminarias_cantidad",
    "puntos_electricos",
    "puntos_electricos_funcionan",
    "puntos_electricos_no_funcionan",
    "puntos_electricos_cantidad",
    "capacidad_aulica",
    "capacidad_distanciamiento",
    "ambiente_apto_retorno",
    "observaciones_detalle",
]


def get_structure(include_area_details=False):
    db = get_db()
    blocks = db.execute(
        "SELECT id, nombre, descripcion, orden FROM bloques ORDER BY orden, id"
    ).fetchall()
    floors = db.execute(
        "SELECT id, bloque_id, nombre, descripcion, orden FROM pisos ORDER BY orden, id"
    ).fetchall()
    area_select = ["id", "piso_id", "nombre", "descripcion", "orden"]
    if include_area_details:
        area_select.extend(AREA_DETAIL_COLUMNS)
    areas = db.execute(
        f"""
        SELECT
            {', '.join(area_select)}
        FROM areas
        ORDER BY orden, id
        """
    ).fetchall()

    floors_by_block = {}
    for floor in floors:
        floors_by_block.setdefault(floor["bloque_id"], []).append(
            {
                "id": floor["id"],
                "nombre": floor["nombre"],
                "descripcion": floor["descripcion"],
                "orden": floor["orden"],
                "areas": [],
            }
        )

    floor_ref = {}
    for floor_list in floors_by_block.values():
        for floor in floor_list:
            floor_ref[floor["id"]] = floor

    for area in areas:
        floor = floor_ref.get(area["piso_id"])
        if floor:
            area_payload = {
                "id": area["id"],
                "nombre": area["nombre"],
                "descripcion": area["descripcion"],
                "orden": area["orden"],
            }
            if include_area_details:
                for column in AREA_DETAIL_COLUMNS:
                    area_payload[column] = area[column]
            floor["areas"].append(area_payload)

    return [
        {
            "id": block["id"],
            "nombre": block["nombre"],
            "descripcion": block["descripcion"],
            "orden": block["orden"],
            "pisos": floors_by_block.get(block["id"], []),
        }
        for block in blocks
    ]


def _next_order(table_name, parent_field=None, parent_id=None):
    db = get_db()
    if parent_field:
        row = db.execute(
            f"SELECT COALESCE(MAX(orden), 0) + 1 AS next_order FROM {table_name} WHERE {parent_field} = ?",
            (parent_id,),
        ).fetchone()
    else:
        row = db.execute(
            f"SELECT COALESCE(MAX(orden), 0) + 1 AS next_order FROM {table_name}"
        ).fetchone()
    return row["next_order"]


def _sync_inventory_location_display(db, area_filter_sql="", area_filter_params=()):
    where_clause = f"WHERE {area_filter_sql}" if area_filter_sql else ""
    db.execute(
        f"""
        UPDATE inventario_items
        SET
            ubicacion = (
                SELECT b.nombre || ' / ' || p.nombre || ' / ' || a.nombre
                FROM areas a
                JOIN pisos p ON p.id = a.piso_id
                JOIN bloques b ON b.id = p.bloque_id
                WHERE a.id = inventario_items.area_id
            ),
            actualizado_en = CURRENT_TIMESTAMP
        {where_clause}
        """,
        tuple(area_filter_params or ()),
    )


def create_block(nombre, descripcion=None):
    db = get_db()
    orden = _next_order("bloques")
    cursor = db.execute(
        "INSERT INTO bloques (nombre, descripcion, orden) VALUES (?, ?, ?)",
        (nombre.strip(), (descripcion or "").strip() or None, orden),
    )
    db.commit()
    return cursor.lastrowid


def update_block(block_id, nombre=None, descripcion=None):
    updates = []
    params = []
    if nombre is not None:
        updates.append("nombre = ?")
        params.append((nombre or "").strip())
    if descripcion is not None:
        updates.append("descripcion = ?")
        params.append((descripcion or "").strip() or None)
    if not updates:
        return False

    params.append(block_id)
    db = get_db()
    cursor = db.execute(
        f"UPDATE bloques SET {', '.join(updates)} WHERE id = ?",
        tuple(params),
    )
    if cursor.rowcount > 0:
        _sync_inventory_location_display(
            db,
            "area_id IN (SELECT a.id FROM areas a JOIN pisos p ON p.id = a.piso_id WHERE p.bloque_id = ?)",
            (block_id,),
        )
    db.commit()
    return cursor.rowcount > 0


def create_floor(bloque_id, nombre, descripcion=None):
    db = get_db()
    orden = _next_order("pisos", "bloque_id", bloque_id)
    cursor = db.execute(
        "INSERT INTO pisos (bloque_id, nombre, descripcion, orden) VALUES (?, ?, ?, ?)",
        (bloque_id, nombre.strip(), (descripcion or "").strip() or None, orden),
    )
    db.commit()
    return cursor.lastrowid


def create_area(piso_id, nombre, descripcion=None, details=None):
    db = get_db()
    orden = _next_order("areas", "piso_id", piso_id)
    details = details or {}
    columns = ["piso_id", "nombre", "descripcion", "orden", *AREA_DETAIL_COLUMNS]
    placeholders = ", ".join(["?"] * len(columns))
    values = [
        piso_id,
        nombre.strip(),
        (descripcion or "").strip() or None,
        orden,
        *[details.get(column) for column in AREA_DETAIL_COLUMNS],
    ]
    cursor = db.execute(
        f"INSERT INTO areas ({', '.join(columns)}) VALUES ({placeholders})",
        tuple(values),
    )
    db.commit()
    return cursor.lastrowid


def update_floor(floor_id, nombre=None, descripcion=None):
    updates = []
    params = []
    if nombre is not None:
        updates.append("nombre = ?")
        params.append((nombre or "").strip())
    if descripcion is not None:
        updates.append("descripcion = ?")
        params.append((descripcion or "").strip() or None)
    if not updates:
        return False

    params.append(floor_id)
    db = get_db()
    cursor = db.execute(
        f"UPDATE pisos SET {', '.join(updates)} WHERE id = ?",
        tuple(params),
    )
    if cursor.rowcount > 0:
        _sync_inventory_location_display(
            db,
            "area_id IN (SELECT id FROM areas WHERE piso_id = ?)",
            (floor_id,),
        )
    db.commit()
    return cursor.rowcount > 0


def delete_floor(floor_id):
    db = get_db()
    cursor = db.execute("DELETE FROM pisos WHERE id = ?", (floor_id,))
    db.commit()
    return cursor.rowcount > 0


def update_area(area_id, nombre=None, descripcion=None, details=None):
    updates = []
    params = []
    details = details or {}
    if nombre is not None:
        updates.append("nombre = ?")
        params.append((nombre or "").strip())
    if descripcion is not None:
        updates.append("descripcion = ?")
        params.append((descripcion or "").strip() or None)

    for column_name in AREA_DETAIL_COLUMNS:
        if column_name in details:
            updates.append(f"{column_name} = ?")
            params.append(details.get(column_name))

    if not updates:
        return False

    params.append(area_id)
    db = get_db()
    cursor = db.execute(
        f"UPDATE areas SET {', '.join(updates)} WHERE id = ?",
        tuple(params),
    )
    if cursor.rowcount > 0:
        _sync_inventory_location_display(db, "area_id = ?", (area_id,))
    db.commit()
    return cursor.rowcount > 0


def delete_area(area_id):
    db = get_db()
    cursor = db.execute("DELETE FROM areas WHERE id = ?", (area_id,))
    db.commit()
    return cursor.rowcount > 0


def delete_block(block_id):
    db = get_db()
    db.execute(
        """
        DELETE FROM inventario_items
        WHERE area_id IN (
            SELECT a.id
            FROM areas a
            JOIN pisos p ON p.id = a.piso_id
            WHERE p.bloque_id = ?
        )
        """,
        (block_id,),
    )
    cursor = db.execute("DELETE FROM bloques WHERE id = ?", (block_id,))
    db.commit()
    return cursor.rowcount > 0


def get_location_dependency_summary(entity_type, entity_id):
    db = get_db()
    entity = str(entity_type or "").strip().lower()

    if entity == "bloque":
        row = db.execute(
            """
            SELECT
                (SELECT COUNT(1) FROM pisos WHERE bloque_id = ?) AS pisos,
                (
                    SELECT COUNT(1)
                    FROM areas a
                    JOIN pisos p ON p.id = a.piso_id
                    WHERE p.bloque_id = ?
                ) AS areas,
                (
                    SELECT COUNT(1)
                    FROM inventario_items i
                    JOIN areas a ON a.id = i.area_id
                    JOIN pisos p ON p.id = a.piso_id
                    WHERE p.bloque_id = ?
                ) AS items
            """,
            (entity_id, entity_id, entity_id),
        ).fetchone()
        return {
            "pisos": row["pisos"] if row else 0,
            "areas": row["areas"] if row else 0,
            "items": row["items"] if row else 0,
        }

    if entity == "piso":
        row = db.execute(
            """
            SELECT
                (SELECT COUNT(1) FROM areas WHERE piso_id = ?) AS areas,
                (
                    SELECT COUNT(1)
                    FROM inventario_items i
                    JOIN areas a ON a.id = i.area_id
                    WHERE a.piso_id = ?
                ) AS items
            """,
            (entity_id, entity_id),
        ).fetchone()
        return {
            "areas": row["areas"] if row else 0,
            "items": row["items"] if row else 0,
        }

    if entity == "area":
        row = db.execute(
            "SELECT COUNT(1) AS items FROM inventario_items WHERE area_id = ?",
            (entity_id,),
        ).fetchone()
        return {"items": row["items"] if row else 0}

    raise ValueError("Entidad no soportada para resumen de dependencias")
