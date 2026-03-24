import sqlite3
import json
from pathlib import Path

from database.db import get_db, execute_schema

ALLOWED_ORDER_FIELDS = {"name", "permisos", "edad"}
ALLOWED_INVENTORY_FIELDS = {
    "item_numero",
    "cod_inventario",
    "cod_esbye",
    "cuenta",
    "cantidad",
    "descripcion",
    "ubicacion",
    "marca",
    "modelo",
    "serie",
    "estado",
    "condicion",
    "usuario_final",
    "fecha_adquisicion",
    "valor",
    "observacion",
    "descripcion_esbye",
    "marca_esbye",
    "modelo_esbye",
    "serie_esbye",
    "valor_esbye",
    "ubicacion_esbye",
    "observacion_esbye",
    "fecha_adquisicion_esbye",
    "area_id",
}
CANONICAL_COLUMN_ORDER = [
    "cod_inventario",
    "cod_esbye",
    "cuenta",
    "cantidad",
    "descripcion",
    "ubicacion",
    "marca",
    "modelo",
    "serie",
    "estado",
    "condicion",
    "usuario_final",
    "fecha_adquisicion",
    "valor",
    "observacion",
    "descripcion_esbye",
    "marca_esbye",
    "modelo_esbye",
    "serie_esbye",
    "fecha_adquisicion_esbye",
    "valor_esbye",
    "ubicacion_esbye",
    "observacion_esbye",
]


def _row_to_persona(row):
    return {
        "name": row["name"],
        "permisos": bool(row["permisos"]),
        "edad": row["edad"],
    }


def _row_to_inventory_item(row):
    return {
        "id": row["id"],
        "item_numero": row["item_numero"],
        "cod_inventario": row["cod_inventario"],
        "cod_esbye": row["cod_esbye"],
        "cuenta": row["cuenta"],
        "cantidad": row["cantidad"],
        "descripcion": row["descripcion"],
        "ubicacion": row["ubicacion"],
        "marca": row["marca"],
        "modelo": row["modelo"],
        "serie": row["serie"],
        "estado": row["estado"],
        "condicion": row["condicion"],
        "usuario_final": row["usuario_final"],
        "fecha_adquisicion": row["fecha_adquisicion"],
        "valor": row["valor"],
        "observacion": row["observacion"],
        "descripcion_esbye": row["descripcion_esbye"],
        "marca_esbye": row["marca_esbye"],
        "modelo_esbye": row["modelo_esbye"],
        "serie_esbye": row["serie_esbye"],
        "valor_esbye": row["valor_esbye"],
        "ubicacion_esbye": row["ubicacion_esbye"],
        "observacion_esbye": row["observacion_esbye"],
        "fecha_adquisicion_esbye": row["fecha_adquisicion_esbye"],
        "area_id": row["area_id"],
        "area_nombre": row["area_nombre"],
        "piso_id": row["piso_id"],
        "piso_nombre": row["piso_nombre"],
        "bloque_id": row["bloque_id"],
        "bloque_nombre": row["bloque_nombre"],
        "actualizado_en": row["actualizado_en"],
    }


def init_schema(base_dir: Path):
    schema_path = base_dir / "database" / "schema.sql"
    execute_schema(schema_path)
    _ensure_university_unique_constraint()
    _ensure_area_extended_columns()
    _ensure_inventory_extended_columns()
    _ensure_inventory_codes_allow_duplicates()
    _seed_default_param_values()


def _ensure_university_unique_constraint():
    db = get_db()
    db.execute(
        """
        DELETE FROM parametros_universidad
        WHERE id NOT IN (
            SELECT MAX(id)
            FROM parametros_universidad
            GROUP BY nombre
        )
        """
    )
    db.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_parametros_universidad_nombre ON parametros_universidad(nombre)"
    )
    db.commit()


def _ensure_area_extended_columns():
    db = get_db()
    existing_columns = {
        row["name"]
        for row in db.execute("PRAGMA table_info(areas)").fetchall()
    }
    required_columns = {
        "responsable_admin_id": "INTEGER",
        "estado_paredes": "TEXT",
        "estado_techo": "TEXT",
        "estado_puerta": "TEXT",
        "cerradura": "TEXT",
        "nivel_seguridad": "TEXT",
        "sitio_profesor_mesa": "TEXT",
        "sitio_profesor_silla": "TEXT",
        "pc_aula": "TEXT",
        "proyector": "TEXT",
        "pantalla_interactiva": "TEXT",
        "aa_cantidad": "INTEGER",
        "aa_funcionan": "INTEGER",
        "ventiladores_cantidad": "INTEGER",
        "wifi": "TEXT",
        "red_lan": "TEXT",
        "red_lan_funcionan": "INTEGER",
        "red_lan_no_funcionan": "INTEGER",
        "red_inalambrica_cantidad": "INTEGER",
        "iluminacion_funcionan": "INTEGER",
        "iluminacion_no_funcionan": "INTEGER",
        "luminarias_cantidad": "INTEGER",
        "puntos_electricos": "TEXT",
        "puntos_electricos_funcionan": "INTEGER",
        "puntos_electricos_no_funcionan": "INTEGER",
        "puntos_electricos_cantidad": "INTEGER",
        "capacidad_aulica": "INTEGER",
        "observaciones_detalle": "TEXT",
    }

    for column_name, column_type in required_columns.items():
        if column_name not in existing_columns:
            db.execute(f"ALTER TABLE areas ADD COLUMN {column_name} {column_type}")
    db.commit()


def _ensure_inventory_extended_columns():
    db = get_db()
    existing_columns = {
        row["name"]
        for row in db.execute("PRAGMA table_info(inventario_items)").fetchall()
    }
    required_columns = {
        "cuenta": "TEXT",
        "ubicacion": "TEXT",
        "condicion": "TEXT",
        "descripcion_esbye": "TEXT",
        "marca_esbye": "TEXT",
        "modelo_esbye": "TEXT",
        "serie_esbye": "TEXT",
        "valor_esbye": "REAL",
        "ubicacion_esbye": "TEXT",
        "observacion_esbye": "TEXT",
        "fecha_adquisicion_esbye": "TEXT",
    }

    for column_name, column_type in required_columns.items():
        if column_name not in existing_columns:
            db.execute(f"ALTER TABLE inventario_items ADD COLUMN {column_name} {column_type}")
    db.commit()


def _ensure_inventory_codes_allow_duplicates():
    db = get_db()
    table_row = db.execute(
        "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'inventario_items'"
    ).fetchone()
    table_sql = (table_row["sql"] or "") if table_row else ""
    normalized_sql = " ".join(table_sql.upper().split())

    has_unique_inventory = "UNIQUE (COD_INVENTARIO)" in normalized_sql or "UNIQUE(COD_INVENTARIO)" in normalized_sql
    has_unique_esbye = "UNIQUE (COD_ESBYE)" in normalized_sql or "UNIQUE(COD_ESBYE)" in normalized_sql
    if not (has_unique_inventory or has_unique_esbye):
        return

    db.execute("PRAGMA foreign_keys = OFF")
    try:
        db.execute("BEGIN")
        db.execute(
            """
            CREATE TABLE inventario_items_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_numero INTEGER NOT NULL,
                cod_inventario TEXT,
                cod_esbye TEXT,
                cuenta TEXT,
                cantidad INTEGER NOT NULL DEFAULT 1,
                descripcion TEXT,
                ubicacion TEXT,
                marca TEXT,
                modelo TEXT,
                serie TEXT,
                estado TEXT,
                condicion TEXT,
                usuario_final TEXT,
                fecha_adquisicion TEXT,
                valor REAL,
                observacion TEXT,
                descripcion_esbye TEXT,
                marca_esbye TEXT,
                modelo_esbye TEXT,
                serie_esbye TEXT,
                valor_esbye REAL,
                ubicacion_esbye TEXT,
                observacion_esbye TEXT,
                fecha_adquisicion_esbye TEXT,
                area_id INTEGER,
                creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                actualizado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (area_id) REFERENCES areas (id) ON DELETE SET NULL
            )
            """
        )
        db.execute(
            """
            INSERT INTO inventario_items_new (
                id, item_numero, cod_inventario, cod_esbye, cuenta, cantidad, descripcion,
                ubicacion, marca, modelo, serie, estado, condicion, usuario_final,
                fecha_adquisicion, valor, observacion, descripcion_esbye, marca_esbye,
                modelo_esbye, serie_esbye, valor_esbye, ubicacion_esbye, observacion_esbye,
                fecha_adquisicion_esbye, area_id, creado_en, actualizado_en
            )
            SELECT
                id, item_numero, cod_inventario, cod_esbye, cuenta, cantidad, descripcion,
                ubicacion, marca, modelo, serie, estado, condicion, usuario_final,
                fecha_adquisicion, valor, observacion, descripcion_esbye, marca_esbye,
                modelo_esbye, serie_esbye, valor_esbye, ubicacion_esbye, observacion_esbye,
                fecha_adquisicion_esbye, area_id, creado_en, actualizado_en
            FROM inventario_items
            """
        )
        db.execute("DROP TABLE inventario_items")
        db.execute("ALTER TABLE inventario_items_new RENAME TO inventario_items")
        db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_cod_inventario ON inventario_items(cod_inventario)")
        db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_cod_esbye ON inventario_items(cod_esbye)")
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.execute("PRAGMA foreign_keys = ON")


def _seed_default_param_values():
    db = get_db()
    db.executemany(
        "INSERT OR IGNORE INTO param_si_no (nombre, descripcion, orden) VALUES (?, ?, ?)",
        [
            ("SI", "Respuesta afirmativa", 1),
            ("NO", "Respuesta negativa", 2),
        ],
    )
    db.executemany(
        "INSERT OR IGNORE INTO param_estado_puerta (nombre, descripcion, orden) VALUES (?, ?, ?)",
        [
            ("BUENO", "Puerta en buen estado", 1),
            ("REGULAR", "Puerta en estado regular", 2),
            ("DAÑADO", "Puerta dañada", 3),
        ],
    )
    db.executemany(
        "INSERT OR IGNORE INTO param_cerraduras (nombre, descripcion, orden) VALUES (?, ?, ?)",
        [
            ("CHAPA", "Tipo chapa", 1),
            ("CANDADO", "Tipo candado", 2),
            ("MANIJA", "Tipo manija", 3),
            ("OTRO", "Otro tipo", 4),
        ],
    )
    db.commit()


def create_bd():
    get_db()


def create_table():
    db = get_db()
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS Persona (
            name TEXT,
            permisos INTEGER,
            edad INTEGER
        )
        """
    )
    db.commit()


def insert_data(name, permisos, edad):
    db = get_db()
    db.execute(
        "INSERT INTO Persona (name, permisos, edad) VALUES (?, ?, ?)",
        (name, int(bool(permisos)), edad),
    )
    db.commit()


def get_personas():
    db = get_db()
    try:
        rows = db.execute(
            "SELECT name, permisos, edad FROM Persona"
        ).fetchall()
    except sqlite3.OperationalError:
        return []
    return [_row_to_persona(row) for row in rows]


def read_rows():
    rows = get_personas()
    print(rows)
    return rows


def insert_rows(persona_list):
    db = get_db()
    rows_to_insert = [
        (name, int(bool(permisos)), edad)
        for (name, permisos, edad) in persona_list
    ]
    db.executemany(
        "INSERT INTO Persona (name, permisos, edad) VALUES (?, ?, ?)",
        rows_to_insert,
    )
    db.commit()


def read_order(field):
    if field not in ALLOWED_ORDER_FIELDS:
        raise ValueError(
            f"Campo de orden no permitido: {field}. "
            f"Use uno de: {', '.join(sorted(ALLOWED_ORDER_FIELDS))}"
        )

    db = get_db()
    rows = db.execute(
        f"SELECT name, permisos, edad FROM Persona ORDER BY {field}"
    ).fetchall()
    personas = [_row_to_persona(row) for row in rows]
    print(personas)
    return personas


def search(name):
    db = get_db()
    rows = db.execute(
        "SELECT name, permisos, edad FROM Persona WHERE name LIKE ?",
        (f"{name}%",),
    ).fetchall()
    personas = [_row_to_persona(row) for row in rows]
    print(personas)
    return personas


# Compatibilidad con nombres antiguos
def createBD():
    create_bd()


def createTable():
    create_table()


def insertData(name, permisos, edad):
    insert_data(name, permisos, edad)


def redRows():
    return read_rows()


def insertRows(persona_list):
    insert_rows(persona_list)


def readOrder(field):
    return read_order(field)


def get_structure():
    db = get_db()
    blocks = db.execute(
        "SELECT id, nombre, descripcion, orden FROM bloques ORDER BY orden, id"
    ).fetchall()
    floors = db.execute(
        "SELECT id, bloque_id, nombre, descripcion, orden FROM pisos ORDER BY orden, id"
    ).fetchall()
    areas = db.execute(
        """
        SELECT
            id,
            piso_id,
            nombre,
            descripcion,
            orden,
            responsable_admin_id,
            estado_paredes,
            estado_techo,
            estado_puerta,
            cerradura,
            nivel_seguridad,
            sitio_profesor_mesa,
            sitio_profesor_silla,
            pc_aula,
            proyector,
            pantalla_interactiva,
            aa_cantidad,
            aa_funcionan,
            ventiladores_cantidad,
            wifi,
            red_lan,
            red_lan_funcionan,
            red_lan_no_funcionan,
            red_inalambrica_cantidad,
            iluminacion_funcionan,
            iluminacion_no_funcionan,
            luminarias_cantidad,
            puntos_electricos,
            puntos_electricos_funcionan,
            puntos_electricos_no_funcionan,
            puntos_electricos_cantidad,
            capacidad_aulica,
            observaciones_detalle
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
            floor["areas"].append(
                {
                    "id": area["id"],
                    "nombre": area["nombre"],
                    "descripcion": area["descripcion"],
                    "orden": area["orden"],
                    "responsable_admin_id": area["responsable_admin_id"],
                    "estado_paredes": area["estado_paredes"],
                    "estado_techo": area["estado_techo"],
                    "estado_puerta": area["estado_puerta"],
                    "cerradura": area["cerradura"],
                    "nivel_seguridad": area["nivel_seguridad"],
                    "sitio_profesor_mesa": area["sitio_profesor_mesa"],
                    "sitio_profesor_silla": area["sitio_profesor_silla"],
                    "pc_aula": area["pc_aula"],
                    "proyector": area["proyector"],
                    "pantalla_interactiva": area["pantalla_interactiva"],
                    "aa_cantidad": area["aa_cantidad"],
                    "aa_funcionan": area["aa_funcionan"],
                    "ventiladores_cantidad": area["ventiladores_cantidad"],
                    "wifi": area["wifi"],
                    "red_lan": area["red_lan"],
                    "red_lan_funcionan": area["red_lan_funcionan"],
                    "red_lan_no_funcionan": area["red_lan_no_funcionan"],
                    "red_inalambrica_cantidad": area["red_inalambrica_cantidad"],
                    "iluminacion_funcionan": area["iluminacion_funcionan"],
                    "iluminacion_no_funcionan": area["iluminacion_no_funcionan"],
                    "luminarias_cantidad": area["luminarias_cantidad"],
                    "puntos_electricos": area["puntos_electricos"],
                    "puntos_electricos_funcionan": area["puntos_electricos_funcionan"],
                    "puntos_electricos_no_funcionan": area["puntos_electricos_no_funcionan"],
                    "puntos_electricos_cantidad": area["puntos_electricos_cantidad"],
                    "capacidad_aulica": area["capacidad_aulica"],
                    "observaciones_detalle": area["observaciones_detalle"],
                }
            )

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
    cursor = db.execute(
        """
        INSERT INTO areas (
            piso_id, nombre, descripcion, orden,
            responsable_admin_id, estado_paredes, estado_techo, estado_puerta, cerradura,
            nivel_seguridad, sitio_profesor_mesa, sitio_profesor_silla, pc_aula, proyector,
            pantalla_interactiva, aa_cantidad, aa_funcionan, ventiladores_cantidad, wifi,
            red_lan, red_lan_funcionan, red_lan_no_funcionan, red_inalambrica_cantidad,
            iluminacion_funcionan, iluminacion_no_funcionan, luminarias_cantidad,
            puntos_electricos, puntos_electricos_funcionan, puntos_electricos_no_funcionan,
            puntos_electricos_cantidad, capacidad_aulica, observaciones_detalle
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            piso_id,
            nombre.strip(),
            (descripcion or "").strip() or None,
            orden,
            details.get("responsable_admin_id"),
            details.get("estado_paredes"),
            details.get("estado_techo"),
            details.get("estado_puerta"),
            details.get("cerradura"),
            details.get("nivel_seguridad"),
            details.get("sitio_profesor_mesa"),
            details.get("sitio_profesor_silla"),
            details.get("pc_aula"),
            details.get("proyector"),
            details.get("pantalla_interactiva"),
            details.get("aa_cantidad"),
            details.get("aa_funcionan"),
            details.get("ventiladores_cantidad"),
            details.get("wifi"),
            details.get("red_lan"),
            details.get("red_lan_funcionan"),
            details.get("red_lan_no_funcionan"),
            details.get("red_inalambrica_cantidad"),
            details.get("iluminacion_funcionan"),
            details.get("iluminacion_no_funcionan"),
            details.get("luminarias_cantidad"),
            details.get("puntos_electricos"),
            details.get("puntos_electricos_funcionan"),
            details.get("puntos_electricos_no_funcionan"),
            details.get("puntos_electricos_cantidad"),
            details.get("capacidad_aulica"),
            details.get("observaciones_detalle"),
        ),
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
    db.commit()
    return cursor.rowcount > 0


def delete_floor(floor_id):
    db = get_db()
    db.execute(
        """
        DELETE FROM inventario_items
        WHERE area_id IN (
            SELECT id FROM areas WHERE piso_id = ?
        )
        """,
        (floor_id,),
    )
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

    field_map = {
        "responsable_admin_id": "responsable_admin_id",
        "estado_paredes": "estado_paredes",
        "estado_techo": "estado_techo",
        "estado_puerta": "estado_puerta",
        "cerradura": "cerradura",
        "nivel_seguridad": "nivel_seguridad",
        "sitio_profesor_mesa": "sitio_profesor_mesa",
        "sitio_profesor_silla": "sitio_profesor_silla",
        "pc_aula": "pc_aula",
        "proyector": "proyector",
        "pantalla_interactiva": "pantalla_interactiva",
        "aa_cantidad": "aa_cantidad",
        "aa_funcionan": "aa_funcionan",
        "ventiladores_cantidad": "ventiladores_cantidad",
        "wifi": "wifi",
        "red_lan": "red_lan",
        "red_lan_funcionan": "red_lan_funcionan",
        "red_lan_no_funcionan": "red_lan_no_funcionan",
        "red_inalambrica_cantidad": "red_inalambrica_cantidad",
        "iluminacion_funcionan": "iluminacion_funcionan",
        "iluminacion_no_funcionan": "iluminacion_no_funcionan",
        "luminarias_cantidad": "luminarias_cantidad",
        "puntos_electricos": "puntos_electricos",
        "puntos_electricos_funcionan": "puntos_electricos_funcionan",
        "puntos_electricos_no_funcionan": "puntos_electricos_no_funcionan",
        "puntos_electricos_cantidad": "puntos_electricos_cantidad",
        "capacidad_aulica": "capacidad_aulica",
        "observaciones_detalle": "observaciones_detalle",
    }
    for input_key, column_name in field_map.items():
        if input_key in details:
            updates.append(f"{column_name} = ?")
            params.append(details.get(input_key))

    if not updates:
        return False

    params.append(area_id)
    db = get_db()
    cursor = db.execute(
        f"UPDATE areas SET {', '.join(updates)} WHERE id = ?",
        tuple(params),
    )
    db.commit()
    return cursor.rowcount > 0


def delete_area(area_id):
    db = get_db()
    db.execute("DELETE FROM inventario_items WHERE area_id = ?", (area_id,))
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
        return {
            "items": row["items"] if row else 0,
        }

    raise ValueError(f"Tipo de ubicación no soportado: {entity_type}")


def _next_item_numero():
    db = get_db()
    row = db.execute(
        "SELECT COALESCE(MAX(item_numero), 0) + 1 AS next_item FROM inventario_items"
    ).fetchone()
    return row["next_item"]


def _audit_change(item_id, action, field=None, old_value=None, new_value=None):
    db = get_db()
    db.execute(
        """
        INSERT INTO inventario_auditoria (item_id, accion, campo, valor_anterior, valor_nuevo)
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            item_id,
            action,
            field,
            None if old_value is None else str(old_value),
            None if new_value is None else str(new_value),
        ),
    )


def _build_inventory_where_clause(filters=None):
    filters = filters or {}
    where_clauses = []
    params = []

    if filters.get("bloque_id"):
        where_clauses.append("p.bloque_id = ?")
        params.append(filters["bloque_id"])
    if filters.get("piso_id"):
        where_clauses.append("a.piso_id = ?")
        params.append(filters["piso_id"])
    if filters.get("area_id"):
        where_clauses.append("i.area_id = ?")
        params.append(filters["area_id"])
    if filters.get("search"):
        token = f"%{filters['search'].strip()}%"
        where_clauses.append(
            "(i.descripcion LIKE ? OR i.cod_inventario LIKE ? OR i.cod_esbye LIKE ? OR i.cuenta LIKE ? OR i.ubicacion LIKE ? OR i.marca LIKE ? OR i.modelo LIKE ? OR i.serie LIKE ? OR i.usuario_final LIKE ? OR i.observacion LIKE ? OR i.descripcion_esbye LIKE ? OR i.marca_esbye LIKE ? OR i.modelo_esbye LIKE ? OR i.serie_esbye LIKE ? OR i.ubicacion_esbye LIKE ? OR i.observacion_esbye LIKE ?)"
        )
        params.extend([token, token, token, token, token, token, token, token, token, token, token, token, token, token, token, token])

    where_sql = ""
    if where_clauses:
        where_sql = "WHERE " + " AND ".join(where_clauses)

    return where_sql, params


def list_inventory_items(filters=None, sort_direction="asc"):
    where_sql, params = _build_inventory_where_clause(filters)

    direction = "DESC" if str(sort_direction).lower() == "desc" else "ASC"
    db = get_db()
    rows = db.execute(
        f"""
        SELECT
            i.id,
            i.item_numero,
            i.cod_inventario,
            i.cod_esbye,
            i.cuenta,
            i.cantidad,
            i.descripcion,
            i.ubicacion,
            i.marca,
            i.modelo,
            i.serie,
            i.estado,
            i.condicion,
            i.usuario_final,
            i.fecha_adquisicion,
            i.valor,
            i.observacion,
            i.descripcion_esbye,
            i.marca_esbye,
            i.modelo_esbye,
            i.serie_esbye,
            i.valor_esbye,
            i.ubicacion_esbye,
            i.observacion_esbye,
            i.fecha_adquisicion_esbye,
            i.area_id,
            i.actualizado_en,
            a.nombre AS area_nombre,
            a.piso_id AS piso_id,
            p.nombre AS piso_nombre,
            p.bloque_id AS bloque_id,
            b.nombre AS bloque_nombre
        FROM inventario_items i
        LEFT JOIN areas a ON a.id = i.area_id
        LEFT JOIN pisos p ON p.id = a.piso_id
        LEFT JOIN bloques b ON b.id = p.bloque_id
        {where_sql}
        ORDER BY i.item_numero {direction}, i.id {direction}
        """,
        params,
    ).fetchall()
    return [_row_to_inventory_item(row) for row in rows]


def list_inventory_items_paginated(filters=None, sort_direction="asc", page=1, per_page=50):
    where_sql, params = _build_inventory_where_clause(filters)
    direction = "DESC" if str(sort_direction).lower() == "desc" else "ASC"

    safe_per_page = max(int(per_page or 50), 1)

    db = get_db()
    total_row = db.execute(
        f"""
        SELECT COUNT(1) AS total
        FROM inventario_items i
        LEFT JOIN areas a ON a.id = i.area_id
        LEFT JOIN pisos p ON p.id = a.piso_id
        LEFT JOIN bloques b ON b.id = p.bloque_id
        {where_sql}
        """,
        params,
    ).fetchone()
    total = total_row["total"] if total_row else 0
    total_pages = (total + safe_per_page - 1) // safe_per_page if total else 0
    safe_page = min(max(int(page or 1), 1), max(total_pages, 1))
    offset = (safe_page - 1) * safe_per_page

    rows = db.execute(
        f"""
        SELECT
            i.id,
            i.item_numero,
            i.cod_inventario,
            i.cod_esbye,
            i.cuenta,
            i.cantidad,
            i.descripcion,
            i.ubicacion,
            i.marca,
            i.modelo,
            i.serie,
            i.estado,
            i.condicion,
            i.usuario_final,
            i.fecha_adquisicion,
            i.valor,
            i.observacion,
            i.descripcion_esbye,
            i.marca_esbye,
            i.modelo_esbye,
            i.serie_esbye,
            i.valor_esbye,
            i.ubicacion_esbye,
            i.observacion_esbye,
            i.fecha_adquisicion_esbye,
            i.area_id,
            i.actualizado_en,
            a.nombre AS area_nombre,
            a.piso_id AS piso_id,
            p.nombre AS piso_nombre,
            p.bloque_id AS bloque_id,
            b.nombre AS bloque_nombre
        FROM inventario_items i
        LEFT JOIN areas a ON a.id = i.area_id
        LEFT JOIN pisos p ON p.id = a.piso_id
        LEFT JOIN bloques b ON b.id = p.bloque_id
        {where_sql}
        ORDER BY i.item_numero {direction}, i.id {direction}
        LIMIT ? OFFSET ?
        """,
        [*params, safe_per_page, offset],
    ).fetchall()

    return {
        "items": [_row_to_inventory_item(row) for row in rows],
        "total": total,
        "page": safe_page,
        "per_page": safe_per_page,
        "total_pages": total_pages,
    }


def get_inventory_item(item_id):
    db = get_db()
    row = db.execute(
        """
        SELECT
            i.id,
            i.item_numero,
            i.cod_inventario,
            i.cod_esbye,
            i.cuenta,
            i.cantidad,
            i.descripcion,
            i.ubicacion,
            i.marca,
            i.modelo,
            i.serie,
            i.estado,
            i.condicion,
            i.usuario_final,
            i.fecha_adquisicion,
            i.valor,
            i.observacion,
            i.descripcion_esbye,
            i.marca_esbye,
            i.modelo_esbye,
            i.serie_esbye,
            i.valor_esbye,
            i.ubicacion_esbye,
            i.observacion_esbye,
            i.fecha_adquisicion_esbye,
            i.area_id,
            i.actualizado_en,
            a.nombre AS area_nombre,
            a.piso_id AS piso_id,
            p.nombre AS piso_nombre,
            p.bloque_id AS bloque_id,
            b.nombre AS bloque_nombre
        FROM inventario_items i
        LEFT JOIN areas a ON a.id = i.area_id
        LEFT JOIN pisos p ON p.id = a.piso_id
        LEFT JOIN bloques b ON b.id = p.bloque_id
        WHERE i.id = ?
        """,
        (item_id,),
    ).fetchone()
    if not row:
        return None
    return _row_to_inventory_item(row)


def create_inventory_item(payload):
    db = get_db()
    fields = {k: payload.get(k) for k in ALLOWED_INVENTORY_FIELDS if k in payload}
    fields["item_numero"] = fields.get("item_numero") or _next_item_numero()
    fields["cantidad"] = int(fields.get("cantidad") or 1)
    fields["valor"] = float(fields.get("valor")) if fields.get("valor") not in (None, "") else None
    fields["valor_esbye"] = float(fields.get("valor_esbye")) if fields.get("valor_esbye") not in (None, "") else None

    columns = ", ".join(fields.keys())
    placeholders = ", ".join(["?"] * len(fields))
    values = list(fields.values())

    cursor = db.execute(
        f"INSERT INTO inventario_items ({columns}) VALUES ({placeholders})",
        values,
    )
    item_id = cursor.lastrowid
    _audit_change(item_id, "create")
    db.commit()
    return item_id


def update_inventory_item(item_id, payload):
    db = get_db()
    current = db.execute(
        "SELECT * FROM inventario_items WHERE id = ?",
        (item_id,),
    ).fetchone()
    if not current:
        return False

    updates = []
    values = []
    for field, value in payload.items():
        if field not in ALLOWED_INVENTORY_FIELDS:
            continue
        if field == "cantidad":
            value = int(value or 1)
        if field == "valor":
            value = float(value) if value not in (None, "") else None
        if field == "valor_esbye":
            value = float(value) if value not in (None, "") else None
        updates.append(f"{field} = ?")
        values.append(value)
        _audit_change(item_id, "update", field, current[field], value)

    if not updates:
        return True

    updates.append("actualizado_en = CURRENT_TIMESTAMP")
    values.append(item_id)
    db.execute(
        f"UPDATE inventario_items SET {', '.join(updates)} WHERE id = ?",
        values,
    )
    db.commit()
    return True


def delete_inventory_item(item_id):
    db = get_db()
    existing = db.execute("SELECT id FROM inventario_items WHERE id = ?", (item_id,)).fetchone()
    if not existing:
        return False
    _audit_change(item_id, "delete")
    db.execute("DELETE FROM inventario_items WHERE id = ?", (item_id,))
    db.commit()
    return True


def find_inventory_code_duplicates(cod_inventario=None, cod_esbye=None, limit=50, exclude_item_id=None):
    inventory_code = str(cod_inventario or "").strip()
    esbye_code = str(cod_esbye or "").strip()
    if not inventory_code and not esbye_code:
        return []

    where_parts = []
    params = []
    if inventory_code:
        where_parts.append("UPPER(TRIM(COALESCE(i.cod_inventario, ''))) = UPPER(TRIM(?))")
        params.append(inventory_code)
    if esbye_code:
        where_parts.append("UPPER(TRIM(COALESCE(i.cod_esbye, ''))) = UPPER(TRIM(?))")
        params.append(esbye_code)

    where_sql = "(" + " OR ".join(where_parts) + ")"
    if exclude_item_id is not None:
        where_sql += " AND i.id <> ?"
        params.append(int(exclude_item_id))

    db = get_db()
    rows = db.execute(
        f"""
        SELECT
            i.id,
            i.item_numero,
            i.cod_inventario,
            i.cod_esbye,
            i.descripcion,
            i.modelo,
            i.ubicacion,
            i.fecha_adquisicion,
            i.usuario_final,
            i.actualizado_en
        FROM inventario_items i
        WHERE {where_sql}
        ORDER BY i.item_numero DESC, i.id DESC
        LIMIT ?
        """,
        [*params, max(int(limit or 50), 1)],
    ).fetchall()

    duplicates = []
    for row in rows:
        duplicate = {
            "id": row["id"],
            "item_numero": row["item_numero"],
            "cod_inventario": row["cod_inventario"],
            "cod_esbye": row["cod_esbye"],
            "descripcion": row["descripcion"],
            "modelo": row["modelo"],
            "ubicacion": row["ubicacion"],
            "fecha_adquisicion": row["fecha_adquisicion"],
            "usuario_final": row["usuario_final"],
            "actualizado_en": row["actualizado_en"],
            "matches": [],
        }
        if inventory_code and str(row["cod_inventario"] or "").strip().lower() == inventory_code.lower():
            duplicate["matches"].append("cod_inventario")
        if esbye_code and str(row["cod_esbye"] or "").strip().lower() == esbye_code.lower():
            duplicate["matches"].append("cod_esbye")
        duplicates.append(duplicate)

    return duplicates


def bulk_insert_inventory_rows(rows, area_id=None):
    db = get_db()
    inserted_ids = []
    for row in rows:
        payload = {}
        for index, raw_value in enumerate(row):
            if index >= len(CANONICAL_COLUMN_ORDER):
                break
            payload[CANONICAL_COLUMN_ORDER[index]] = raw_value
        if area_id and not payload.get("area_id"):
            payload["area_id"] = area_id
        item_id = create_inventory_item(payload)
        inserted_ids.append(item_id)
    db.commit()
    return inserted_ids


def get_user_preferences(user_key):
    db = get_db()
    rows = db.execute(
        "SELECT pref_key, pref_value FROM user_preferences WHERE user_key = ?",
        (user_key,),
    ).fetchall()
    preferences = {}
    for row in rows:
        try:
            preferences[row["pref_key"]] = json.loads(row["pref_value"]) if row["pref_value"] else None
        except json.JSONDecodeError:
            preferences[row["pref_key"]] = row["pref_value"]
    return preferences


def set_user_preference(user_key, pref_key, pref_value):
    db = get_db()
    serialized_value = json.dumps(pref_value)
    db.execute(
        """
        INSERT INTO user_preferences (user_key, pref_key, pref_value, actualizado_en)
        VALUES (?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(user_key, pref_key)
        DO UPDATE SET pref_value = excluded.pref_value, actualizado_en = CURRENT_TIMESTAMP
        """,
        (user_key, pref_key, serialized_value),
    )
    db.commit()


def get_column_mappings():
    db = get_db()
    rows = db.execute(
        "SELECT columna_origen, campo_canonico, orden FROM column_mappings ORDER BY orden, id"
    ).fetchall()
    return [
        {
            "columna_origen": row["columna_origen"],
            "campo_canonico": row["campo_canonico"],
            "orden": row["orden"],
        }
        for row in rows
    ]


def replace_column_mappings(mappings):
    db = get_db()
    db.execute("DELETE FROM column_mappings")
    for idx, mapping in enumerate(mappings, start=1):
        db.execute(
            "INSERT INTO column_mappings (columna_origen, campo_canonico, orden) VALUES (?, ?, ?)",
            (
                mapping.get("columna_origen", "").strip(),
                mapping.get("campo_canonico", "").strip(),
                mapping.get("orden", idx),
            ),
        )
    db.commit()