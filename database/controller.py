import sqlite3
import json
import re
from difflib import SequenceMatcher
from datetime import datetime, date, timedelta
from pathlib import Path

from database.db import get_db, execute_schema

AREA_EXPORT_COLUMN_ORDER = [
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

_PLACEHOLDER_NO_CODE = "S/C"
_CODE_FIELDS = {"cod_inventario", "cod_esbye"}


def _normalize_inventory_code_value(value):
    text = str(value or "").strip()
    if not text:
        return _PLACEHOLDER_NO_CODE

    compact = re.sub(r"[^a-z0-9]", "", text.lower())
    if compact in {"sc", "sincodigo", "sincod"}:
        return _PLACEHOLDER_NO_CODE
    return text


def _normalize_inventory_code_fields(payload):
    if not isinstance(payload, dict):
        return payload
    for code_field in _CODE_FIELDS:
        if code_field in payload:
            payload[code_field] = _normalize_inventory_code_value(payload.get(code_field))
    return payload


def _is_placeholder_no_code(value):
    return _normalize_inventory_code_value(value).upper() == _PLACEHOLDER_NO_CODE

PROJECT_ROOT = Path(__file__).resolve().parents[1]

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


def _ensure_schema_migrations_table():
    db = get_db()
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS schema_migrations (
            name TEXT PRIMARY KEY,
            applied_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    # Compatibilidad con versiones anteriores que usaban migration_name/applied_en.
    cols = {
        row["name"]
        for row in db.execute("PRAGMA table_info(schema_migrations)").fetchall()
    }
    if "name" not in cols and "migration_name" in cols:
        db.execute("ALTER TABLE schema_migrations ADD COLUMN name TEXT")
        db.execute("UPDATE schema_migrations SET name = migration_name WHERE name IS NULL")
        db.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_schema_migrations_name ON schema_migrations(name)")

    if "applied_at" not in cols and "applied_en" in cols:
        db.execute("ALTER TABLE schema_migrations ADD COLUMN applied_at TEXT")
        db.execute("UPDATE schema_migrations SET applied_at = applied_en WHERE applied_at IS NULL")

    db.commit()


def _run_startup_migration_once(name, migration_fn):
    db = get_db()
    cols = {
        row["name"]
        for row in db.execute("PRAGMA table_info(schema_migrations)").fetchall()
    }
    key_col = "name" if "name" in cols else "migration_name"
    existing = db.execute(
        f"SELECT 1 FROM schema_migrations WHERE {key_col} = ? LIMIT 1",
        (name,),
    ).fetchone()
    if existing:
        return False

    migration_fn()
    if "name" in cols:
        db.execute("INSERT INTO schema_migrations (name) VALUES (?)", (name,))
    else:
        db.execute("INSERT INTO schema_migrations (migration_name) VALUES (?)", (name,))
    db.commit()
    return True


def _to_storage_relative_path(path_value):
    raw = str(path_value or "").strip()
    if not raw:
        return None

    path_obj = Path(raw)
    if not path_obj.is_absolute():
        return raw.replace("\\", "/")

    try:
        rel = path_obj.resolve().relative_to(PROJECT_ROOT.resolve())
        return str(rel).replace("\\", "/")
    except Exception:
        return raw


def _to_storage_absolute_path(path_value):
    raw = str(path_value or "").strip()
    if not raw:
        return None

    path_obj = Path(raw)
    if path_obj.is_absolute():
        return str(path_obj)

    return str((PROJECT_ROOT / raw).resolve())


def _normalize_historial_row_paths(row_dict):
    normalized = dict(row_dict or {})
    for field in ("docx_path", "pdf_path", "plantilla_snapshot_path"):
        normalized[field] = _to_storage_absolute_path(normalized.get(field))
    return normalized


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
    _ensure_schema_migrations_table()

    _run_startup_migration_once("20260409_university_unique_constraint", _ensure_university_unique_constraint)
    _run_startup_migration_once("20260409_area_extended_columns", _ensure_area_extended_columns)
    _run_startup_migration_once("20260409_inventory_extended_columns", _ensure_inventory_extended_columns)
    _run_startup_migration_once("20260409_inventory_codes_allow_duplicates", _ensure_inventory_codes_allow_duplicates)
    _run_startup_migration_once("20260409_inventory_search_indexes", _ensure_inventory_search_indexes)
    _run_startup_migration_once("20260409_inventory_fts", _ensure_inventory_fts)
    _run_startup_migration_once("20260409_historial_actas_numero_column", _ensure_historial_actas_numero_column)
    _run_startup_migration_once("20260409_historial_actas_template_columns", _ensure_historial_actas_template_columns)
    _run_startup_migration_once("20260409_informes_area_sequence_table", _ensure_informes_area_sequence_table)
    _run_startup_migration_once("20260409_actas_sequence_table", _ensure_actas_sequence_table)
    _run_startup_migration_once("20260414_historial_actas_numero_unique_by_type", _ensure_historial_actas_numero_unique_by_type)
    _run_startup_migration_once("20260414_actas_sequence_by_type_table", _ensure_actas_sequence_by_type_table)
    _run_startup_migration_once("20260409_seed_default_param_values", _seed_default_param_values)


def _ensure_historial_actas_numero_column():
    db = get_db()
    existing_columns = {row["name"] for row in db.execute("PRAGMA table_info(historial_actas)").fetchall()}

    if "numero_acta" not in existing_columns:
        db.execute("ALTER TABLE historial_actas ADD COLUMN numero_acta TEXT")

    db.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_historial_actas_numero_acta ON historial_actas(numero_acta)")
    db.commit()


def _ensure_historial_actas_numero_unique_by_type():
    db = get_db()
    # El esquema nuevo permite mismo numero_acta en tipos diferentes.
    db.execute("DROP INDEX IF EXISTS idx_historial_actas_numero_acta")
    db.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS idx_historial_actas_tipo_numero_acta
        ON historial_actas(tipo_acta, numero_acta)
        WHERE numero_acta IS NOT NULL AND TRIM(numero_acta) != ''
        """
    )
    db.commit()


def _ensure_historial_actas_template_columns():
    db = get_db()
    existing_columns = {row["name"] for row in db.execute("PRAGMA table_info(historial_actas)").fetchall()}

    if "plantilla_hash" not in existing_columns:
        db.execute("ALTER TABLE historial_actas ADD COLUMN plantilla_hash TEXT")
    if "plantilla_snapshot_path" not in existing_columns:
        db.execute("ALTER TABLE historial_actas ADD COLUMN plantilla_snapshot_path TEXT")

    db.execute(
        "CREATE INDEX IF NOT EXISTS idx_historial_actas_plantilla_snapshot_path ON historial_actas(plantilla_snapshot_path)"
    )
    db.commit()


def _ensure_informes_area_sequence_table():
    db = get_db()
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS secuencia_informes_area (
            anio INTEGER PRIMARY KEY,
            ultimo_numero INTEGER NOT NULL DEFAULT 0,
            actualizado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    db.commit()


def _ensure_actas_sequence_table():
    db = get_db()
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS secuencia_actas (
            anio INTEGER PRIMARY KEY,
            ultimo_numero INTEGER NOT NULL DEFAULT 0,
            actualizado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    db.commit()


def _ensure_actas_sequence_by_type_table():
    db = get_db()
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS secuencia_actas_tipo (
            tipo_acta TEXT NOT NULL,
            anio INTEGER NOT NULL,
            ultimo_numero INTEGER NOT NULL DEFAULT 0,
            actualizado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (tipo_acta, anio)
        )
        """
    )
    db.commit()


def _ensure_actas_sequence_by_type_table_runtime():
    """Asegura en runtime la tabla de secuencia por tipo para DBs antiguas o inconsistentes."""
    db = get_db()
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS secuencia_actas_tipo (
            tipo_acta TEXT NOT NULL,
            anio INTEGER NOT NULL,
            ultimo_numero INTEGER NOT NULL DEFAULT 0,
            actualizado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (tipo_acta, anio)
        )
        """
    )
    db.commit()


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
        "identificacion_ambiente": "TEXT",
        "metros_cuadrados": "TEXT",
        "alto": "REAL",
        "senaletica": "TEXT",
        "cod_senaletica": "TEXT",
        "infraestructura_fisica": "TEXT",
        "estado_piso": "TEXT",
        "material_techo": "TEXT",
        "puerta": "TEXT",
        "material_puerta": "TEXT",
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
        "pupitres_cantidad": "INTEGER",
        "pupitres_funcionan": "INTEGER",
        "pupitres_no_funcionan": "INTEGER",
        "pizarra": "TEXT",
        "pizarra_estado": "TEXT",
        "ventanas_cantidad": "INTEGER",
        "ventanas_funcionan": "INTEGER",
        "ventanas_no_funcionan": "INTEGER",
        "aa_cantidad": "INTEGER",
        "aa_funcionan": "INTEGER",
        "aa_no_funcionan": "INTEGER",
        "ventiladores_cantidad": "INTEGER",
        "ventiladores_funcionan": "INTEGER",
        "ventiladores_no_funcionan": "INTEGER",
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
        "capacidad_distanciamiento": "INTEGER",
        "ambiente_apto_retorno": "TEXT",
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
    db.executemany(
        "INSERT OR IGNORE INTO param_estado_piso (nombre, descripcion, orden) VALUES (?, ?, ?)",
        [
            ("BUENO", "Piso en buen estado", 1),
            ("REGULAR", "Piso en estado regular", 2),
            ("DAÑADO", "Piso dañado", 3),
        ],
    )
    db.executemany(
        "INSERT OR IGNORE INTO param_material_techo (nombre, descripcion, orden) VALUES (?, ?, ?)",
        [
            ("HORMIGON", "Techo de hormigon", 1),
            ("GYPSUM", "Techo de gypsum", 2),
            ("METAL", "Techo metalico", 3),
            ("OTRO", "Otro material", 4),
        ],
    )
    db.executemany(
        "INSERT OR IGNORE INTO param_material_puerta (nombre, descripcion, orden) VALUES (?, ?, ?)",
        [
            ("ALUMINIO", "Puerta de aluminio", 1),
            ("ALUMINIO CON VIDRIO", "Puerta de aluminio con vidrio", 2),
            ("MADERA", "Puerta de madera", 3),
            ("HIERRO", "Puerta de hierro", 4),
            ("OTRO", "Otro material", 5),
        ],
    )
    db.executemany(
        "INSERT OR IGNORE INTO param_estado_pizarra (nombre, descripcion, orden) VALUES (?, ?, ?)",
        [
            ("BUENO", "Pizarra en buen estado", 1),
            ("REGULAR", "Pizarra en estado regular", 2),
            ("DAÑADO", "Pizarra dañada", 3),
        ],
    )
    db.commit()


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
    detail_columns = [
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
    columns = ["piso_id", "nombre", "descripcion", "orden", *detail_columns]
    placeholders = ", ".join(["?"] * len(columns))
    values = [
        piso_id,
        nombre.strip(),
        (descripcion or "").strip() or None,
        orden,
        *[details.get(column) for column in detail_columns],
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

    field_map = {
        "identificacion_ambiente": "identificacion_ambiente",
        "metros_cuadrados": "metros_cuadrados",
        "alto": "alto",
        "senaletica": "senaletica",
        "cod_senaletica": "cod_senaletica",
        "infraestructura_fisica": "infraestructura_fisica",
        "estado_piso": "estado_piso",
        "material_techo": "material_techo",
        "puerta": "puerta",
        "material_puerta": "material_puerta",
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
        "pupitres_cantidad": "pupitres_cantidad",
        "pupitres_funcionan": "pupitres_funcionan",
        "pupitres_no_funcionan": "pupitres_no_funcionan",
        "pizarra": "pizarra",
        "pizarra_estado": "pizarra_estado",
        "ventanas_cantidad": "ventanas_cantidad",
        "ventanas_funcionan": "ventanas_funcionan",
        "ventanas_no_funcionan": "ventanas_no_funcionan",
        "aa_cantidad": "aa_cantidad",
        "aa_funcionan": "aa_funcionan",
        "aa_no_funcionan": "aa_no_funcionan",
        "ventiladores_cantidad": "ventiladores_cantidad",
        "ventiladores_funcionan": "ventiladores_funcionan",
        "ventiladores_no_funcionan": "ventiladores_no_funcionan",
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
        "capacidad_distanciamiento": "capacidad_distanciamiento",
        "ambiente_apto_retorno": "ambiente_apto_retorno",
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


def _next_item_numero_in_tx(db):
    row = db.execute(
        "SELECT COALESCE(MAX(item_numero), 0) + 1 AS next_item FROM inventario_items"
    ).fetchone()
    return int(row["next_item"] if row and row["next_item"] is not None else 1)


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


def _ensure_inventory_search_indexes():
    db = get_db()
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_descripcion ON inventario_items(descripcion)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_cod_inventario ON inventario_items(cod_inventario)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_cod_esbye ON inventario_items(cod_esbye)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_area_id ON inventario_items(area_id)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_item_numero_id ON inventario_items(item_numero, id)")
    db.commit()


def _ensure_inventory_fts():
    db = get_db()
    try:
        db.execute(
            """
            CREATE VIRTUAL TABLE IF NOT EXISTS inventario_items_fts
            USING fts5(
                descripcion,
                cod_inventario,
                cod_esbye,
                content='inventario_items',
                content_rowid='id'
            )
            """
        )
        db.execute(
            """
            CREATE TRIGGER IF NOT EXISTS inventario_items_ai
            AFTER INSERT ON inventario_items
            BEGIN
                INSERT INTO inventario_items_fts(rowid, descripcion, cod_inventario, cod_esbye)
                VALUES (new.id, new.descripcion, new.cod_inventario, new.cod_esbye);
            END
            """
        )
        db.execute(
            """
            CREATE TRIGGER IF NOT EXISTS inventario_items_ad
            AFTER DELETE ON inventario_items
            BEGIN
                INSERT INTO inventario_items_fts(inventario_items_fts, rowid, descripcion, cod_inventario, cod_esbye)
                VALUES('delete', old.id, old.descripcion, old.cod_inventario, old.cod_esbye);
            END
            """
        )
        db.execute(
            """
            CREATE TRIGGER IF NOT EXISTS inventario_items_au
            AFTER UPDATE ON inventario_items
            BEGIN
                INSERT INTO inventario_items_fts(inventario_items_fts, rowid, descripcion, cod_inventario, cod_esbye)
                VALUES('delete', old.id, old.descripcion, old.cod_inventario, old.cod_esbye);
                INSERT INTO inventario_items_fts(rowid, descripcion, cod_inventario, cod_esbye)
                VALUES (new.id, new.descripcion, new.cod_inventario, new.cod_esbye);
            END
            """
        )
        fts_count_row = db.execute("SELECT COUNT(1) AS total FROM inventario_items_fts").fetchone()
        items_count_row = db.execute("SELECT COUNT(1) AS total FROM inventario_items").fetchone()
        if (fts_count_row["total"] or 0) != (items_count_row["total"] or 0):
            db.execute("INSERT INTO inventario_items_fts(inventario_items_fts) VALUES ('rebuild')")
        db.commit()
    except sqlite3.OperationalError:
        # Some SQLite builds might not include FTS5; LIKE fallback remains active.
        db.rollback()


def _has_inventory_fts():
    db = get_db()
    row = db.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'inventario_items_fts'"
    ).fetchone()
    return row is not None


def _build_fts_query(raw_search):
    tokens = re.findall(r"[\w\-]+", str(raw_search or ""), flags=re.UNICODE)
    if not tokens:
        return ""
    return " AND ".join([f'{token}*' for token in tokens])


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
        raw_search = filters["search"].strip()
        fts_query = _build_fts_query(raw_search)
        if fts_query and _has_inventory_fts():
            where_clauses.append(
                "i.id IN (SELECT rowid FROM inventario_items_fts WHERE inventario_items_fts MATCH ?)"
            )
            params.append(fts_query)
        else:
            token = f"%{raw_search}%"
            where_clauses.append(
                "(i.descripcion LIKE ? OR i.cod_inventario LIKE ? OR i.cod_esbye LIKE ? OR i.cuenta LIKE ? OR i.ubicacion LIKE ? OR i.marca LIKE ? OR i.modelo LIKE ? OR i.serie LIKE ? OR i.usuario_final LIKE ? OR i.observacion LIKE ? OR i.descripcion_esbye LIKE ? OR i.marca_esbye LIKE ? OR i.modelo_esbye LIKE ? OR i.serie_esbye LIKE ? OR i.ubicacion_esbye LIKE ? OR i.observacion_esbye LIKE ?)"
            )
            params.extend([token, token, token, token, token, token, token, token, token, token, token, token, token, token, token, token])

    where_sql = ""
    if where_clauses:
        where_sql = "WHERE " + " AND ".join(where_clauses)

    return where_sql, params


def get_inventory_search_diagnostics(search_text=None):
    db = get_db()
    filters = {"search": search_text} if search_text else {}
    where_sql, params = _build_inventory_where_clause(filters)
    query = f"""
        SELECT i.id
        FROM inventario_items i
        LEFT JOIN areas a ON a.id = i.area_id
        LEFT JOIN pisos p ON p.id = a.piso_id
        LEFT JOIN bloques b ON b.id = p.bloque_id
        {where_sql}
        ORDER BY i.item_numero DESC, i.id DESC
        LIMIT 50
    """
    plan_rows = db.execute(f"EXPLAIN QUERY PLAN {query}", params).fetchall()
    journal_mode_row = db.execute("PRAGMA journal_mode").fetchone()
    fts_count_row = db.execute(
        "SELECT COUNT(1) AS total FROM sqlite_master WHERE type='table' AND name='inventario_items_fts'"
    ).fetchone()

    return {
        "journal_mode": journal_mode_row[0] if journal_mode_row else None,
        "fts_available": bool((fts_count_row["total"] if fts_count_row else 0) > 0),
        "using_fts": "inventario_items_fts" in where_sql,
        "search_text": search_text,
        "where_sql": where_sql,
        "query_plan": [
            {
                "id": row[0],
                "parent": row[1],
                "notused": row[2],
                "detail": row[3],
            }
            for row in plan_rows
        ],
    }


def list_inventory_items(filters=None, sort_direction="asc", limit=5000):
    where_sql, params = _build_inventory_where_clause(filters)

    direction = "DESC" if str(sort_direction).lower() == "desc" else "ASC"
    safe_limit = max(int(limit or 5000), 1)
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
        LIMIT ?
        """,
        [*params, safe_limit],
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


def create_inventory_item(payload, commit=True):
    db = get_db()
    fields = {k: payload.get(k) for k in ALLOWED_INVENTORY_FIELDS if k in payload}
    _normalize_inventory_code_fields(fields)
    fields["cantidad"] = int(fields.get("cantidad") or 1)
    fields["valor"] = float(fields.get("valor")) if fields.get("valor") not in (None, "") else None
    fields["valor_esbye"] = float(fields.get("valor_esbye")) if fields.get("valor_esbye") not in (None, "") else None

    if fields.get("item_numero") in (None, ""):
        if commit:
            db.execute("BEGIN IMMEDIATE")
        fields["item_numero"] = _next_item_numero_in_tx(db)

    columns = ", ".join(fields.keys())
    placeholders = ", ".join(["?"] * len(fields))
    values = list(fields.values())

    cursor = db.execute(
        f"INSERT INTO inventario_items ({columns}) VALUES ({placeholders})",
        values,
    )
    item_id = cursor.lastrowid
    _audit_change(item_id, "create")
    if commit:
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
    normalized_payload = dict(payload or {})
    _normalize_inventory_code_fields(normalized_payload)
    for field, value in normalized_payload.items():
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


def clear_inventory_items(reset_sequence=True):
    db = get_db()
    try:
        db.execute("BEGIN IMMEDIATE")
        deleted = db.execute("SELECT COUNT(1) AS total FROM inventario_items").fetchone()["total"]
        db.execute("DELETE FROM inventario_items")
        if reset_sequence:
            db.execute("DELETE FROM sqlite_sequence WHERE name = ?", ("inventario_items",))
        db.commit()
        return {"deleted": int(deleted or 0)}
    except Exception:
        db.rollback()
        raise


def find_inventory_code_duplicates(cod_inventario=None, cod_esbye=None, limit=50, exclude_item_id=None):
    inventory_code = _normalize_inventory_code_value(cod_inventario)
    esbye_code = _normalize_inventory_code_value(cod_esbye)
    if _is_placeholder_no_code(inventory_code):
        inventory_code = ""
    if _is_placeholder_no_code(esbye_code):
        esbye_code = ""
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
    normalized_values = []
    insert_columns = ["item_numero", *CANONICAL_COLUMN_ORDER, "area_id"]
    try:
        db.execute("BEGIN IMMEDIATE")
        start_item_numero = _next_item_numero_in_tx(db)
        for row_index, row in enumerate(rows):
            payload = {
                "item_numero": start_item_numero + row_index,
                "area_id": area_id,
            }
            for index, raw_value in enumerate(row):
                if index >= len(CANONICAL_COLUMN_ORDER):
                    break
                payload[CANONICAL_COLUMN_ORDER[index]] = raw_value
            _normalize_inventory_code_fields(payload)
            if area_id and not payload.get("area_id"):
                payload["area_id"] = area_id

            payload["cantidad"] = int(payload.get("cantidad") or 1)
            payload["valor"] = float(payload.get("valor")) if payload.get("valor") not in (None, "") else None
            payload["valor_esbye"] = float(payload.get("valor_esbye")) if payload.get("valor_esbye") not in (None, "") else None
            normalized_values.append(tuple(payload.get(column) for column in insert_columns))

        if normalized_values:
            placeholders = ", ".join(["?"] * len(insert_columns))
            db.executemany(
                f"INSERT INTO inventario_items ({', '.join(insert_columns)}) VALUES ({placeholders})",
                normalized_values,
            )
        db.commit()
    except Exception:
        db.rollback()
        raise

    if not normalized_values:
        return []

    last_insert_rowid = db.execute("SELECT last_insert_rowid() AS last_id").fetchone()["last_id"]
    first_insert_rowid = last_insert_rowid - len(normalized_values) + 1
    return list(range(first_insert_rowid, last_insert_rowid + 1))


def bulk_insert_inventory_dicts(rows_as_dicts, area_id=None):
    db = get_db()
    insert_columns = ["item_numero", *CANONICAL_COLUMN_ORDER, "area_id"]
    normalized_values = []
    skipped = 0

    try:
        db.execute("BEGIN IMMEDIATE")
        start_item_numero = _next_item_numero_in_tx(db)

        def _normalize_import_date(raw_value):
            if raw_value in (None, ""):
                return None

            if isinstance(raw_value, datetime):
                return raw_value.date().isoformat()
            if isinstance(raw_value, date):
                return raw_value.isoformat()

            text = str(raw_value).strip()
            if not text:
                return None

            # Maneja serial de Excel como texto numerico.
            if re.fullmatch(r"\d+(?:\.\d+)?", text):
                serial = float(text)
                if serial > 59:
                    # Excel epoch compatible.
                    epoch = datetime(1899, 12, 30)
                    parsed = epoch + timedelta(days=int(serial))
                    return parsed.date().isoformat()

            date_patterns = [
                "%d/%m/%Y",
                "%d/%m/%y",
                "%d-%m-%Y",
                "%d-%m-%y",
                "%Y-%m-%d",
                "%d-%b-%y",
                "%d-%b-%Y",
                "%d-%B-%Y",
                "%d-%B-%y",
            ]

            # Fuerza locale neutral para meses en ingles (Nov, Dec, etc.).
            text_normalized = text.replace(".", "-").replace("/", "/")

            for pattern in date_patterns:
                try:
                    parsed = datetime.strptime(text_normalized, pattern)
                    return parsed.date().isoformat()
                except ValueError:
                    continue

            return text

        for row_idx, raw_row in enumerate(rows_as_dicts or []):
            if not isinstance(raw_row, dict):
                skipped += 1
                continue

            payload = {
                "item_numero": start_item_numero + len(normalized_values),
                "area_id": area_id,
            }

            has_data = False
            for field in CANONICAL_COLUMN_ORDER:
                if field not in raw_row:
                    continue
                value = raw_row.get(field)
                if isinstance(value, str):
                    value = value.strip()
                if value in ("", None):
                    value = None
                else:
                    has_data = True
                payload[field] = value

            _normalize_inventory_code_fields(payload)

            if "area_id" in raw_row and raw_row.get("area_id") not in (None, ""):
                try:
                    payload["area_id"] = int(raw_row.get("area_id"))
                except (TypeError, ValueError):
                    payload["area_id"] = area_id

            if area_id and not payload.get("area_id"):
                payload["area_id"] = area_id

            if not has_data:
                skipped += 1
                continue

            try:
                payload["cantidad"] = int(payload.get("cantidad") or 1)
            except (TypeError, ValueError):
                payload["cantidad"] = 1

            for money_field in ("valor", "valor_esbye"):
                raw_money = payload.get(money_field)
                if raw_money in (None, ""):
                    payload[money_field] = None
                    continue
                text_money = str(raw_money).strip().replace(" ", "")
                if "," in text_money and "." in text_money:
                    text_money = text_money.replace(".", "").replace(",", ".")
                elif "," in text_money:
                    text_money = text_money.replace(",", ".")
                try:
                    payload[money_field] = float(text_money)
                except (TypeError, ValueError):
                    payload[money_field] = None

            for date_field in ("fecha_adquisicion", "fecha_adquisicion_esbye"):
                payload[date_field] = _normalize_import_date(payload.get(date_field))

            normalized_values.append(tuple(payload.get(column) for column in insert_columns))

        if normalized_values:
            placeholders = ", ".join(["?"] * len(insert_columns))
            db.executemany(
                f"INSERT INTO inventario_items ({', '.join(insert_columns)}) VALUES ({placeholders})",
                normalized_values,
            )

        db.commit()
    except Exception:
        db.rollback()
        raise

    return {
        "inserted": len(normalized_values),
        "skipped": skipped,
    }


def iter_inventory_items(filters=None, sort_direction="asc", batch_size=2000):
    where_sql, params = _build_inventory_where_clause(filters)
    direction = "DESC" if str(sort_direction).lower() == "desc" else "ASC"
    per_page = max(int(batch_size or 2000), 1)
    offset = 0
    db = get_db()

    while True:
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
            [*params, per_page, offset],
        ).fetchall()
        if not rows:
            break
        for row in rows:
            yield _row_to_inventory_item(row)
        offset += per_page


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


def get_all_areas_for_export():
    db = get_db()
    existing_columns = {
        row["name"] for row in db.execute("PRAGMA table_info(areas)").fetchall()
    }
    ordered_area_columns = [
        column for column in AREA_EXPORT_COLUMN_ORDER if column in existing_columns
    ]
    dynamic_select = ",\n            ".join(
        [f"a.{column} AS {column}" for column in ordered_area_columns]
    )

    extra_columns_sql = f",\n            {dynamic_select}" if dynamic_select else ""

    query = f"""
        SELECT 
            b.nombre || ' / ' || p.nombre AS ubicacion,
            a.nombre AS ambiente_aprendizaje
            {extra_columns_sql}
        FROM areas a
        JOIN pisos p ON p.id = a.piso_id
        JOIN bloques b ON b.id = p.bloque_id
        ORDER BY b.orden, p.orden, a.orden;
    """

    rows = db.execute(query).fetchall()
    return [dict(row) for row in rows]

def get_dashboard_stats():
    db = get_db()
    total_bienes = db.execute('SELECT COUNT(1) as total FROM inventario_items').fetchone()['total']
    total_actas = db.execute('SELECT COUNT(1) as total FROM historial_actas').fetchone()['total']
    total_actas_recibidas = db.execute(
        """
        SELECT COUNT(1) as total
        FROM historial_actas
        WHERE LOWER(COALESCE(tipo_acta, '')) LIKE '%recib%'
        """
    ).fetchone()['total']
    return {
        'cant_bienes': total_bienes or 0,
        'cant_actas': total_actas or 0,
        'cant_actas_recibidas': total_actas_recibidas or 0,
    }

# --- PERSONAL / ADMINISTRADORES ---
def get_personal():
    db = get_db()
    rows = db.execute("SELECT id, nombre, cargo FROM administradores WHERE activo = 1 ORDER BY nombre ASC").fetchall()
    return [{"id": row["id"], "nombre": row["nombre"], "cargo": row["cargo"]} for row in rows]


def _normalize_personal_name(value):
    text = str(value or "").strip().lower()
    if not text:
        return ""
    text = text.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u").replace("ü", "u")
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = " ".join(text.split())
    tokens = text.split()
    prefixes = {
        "ing", "ingeniero", "ingeniera", "dr", "dra", "doctor", "doctora",
        "lic", "licenciado", "licenciada", "abg", "abogada", "abogado",
        "arq", "arquitecto", "arquitecta", "tec", "tecnico", "tecnica",
        "sr", "sra", "srta", "msc", "mg", "mgs", "mgtr", "mtr", "mts", "mtro", "mt",
        "prof", "profa", "tlgo", "tlga", "ts", "phd",
    }
    while tokens and tokens[0] in prefixes:
        tokens.pop(0)
    return " ".join(tokens)


def _personal_name_variants(value):
    normalized = _normalize_personal_name(value)
    if not normalized:
        return []
    variants = [normalized]
    tokens = normalized.split()
    # Ayuda con prefijos/abreviaturas no contempladas (ej. "tlgo", "prof", etc.).
    if len(tokens) >= 2:
        variants.append(" ".join(tokens[1:]))
    return [v for v in dict.fromkeys(variants) if v]


def _resolve_existing_personal_name(nombre, min_similarity=0.86):
    db = get_db()
    raw_name = str(nombre or "").strip()
    if not raw_name:
        return None

    # Coincidencia exacta (insensible a mayúsculas) para mantener el caso ya registrado.
    exact_row = db.execute(
        "SELECT nombre FROM administradores WHERE UPPER(nombre) = UPPER(?) LIMIT 1",
        (raw_name,),
    ).fetchone()
    if exact_row:
        return exact_row["nombre"]

    candidates = db.execute(
        "SELECT nombre FROM administradores"
    ).fetchall()
    if not candidates:
        return None

    target_variants = _personal_name_variants(raw_name)
    if not target_variants:
        return None

    # Coincidencia exacta normalizada (ignora tildes/puntuación/títulos al inicio).
    for row in candidates:
        candidate_name = str(row["nombre"] or "").strip()
        candidate_variants = _personal_name_variants(candidate_name)
        if any(tv == cv for tv in target_variants for cv in candidate_variants):
            return candidate_name

    # Coincidencia difusa para casos con abreviaciones o variantes menores.
    best_name = None
    best_score = 0.0
    target_tokens = set(" ".join(target_variants).split())
    for row in candidates:
        candidate_name = str(row["nombre"] or "").strip()
        candidate_variants = _personal_name_variants(candidate_name)
        if not candidate_variants:
            continue
        seq_score = max(
            SequenceMatcher(None, tv, cv).ratio()
            for tv in target_variants
            for cv in candidate_variants
        )
        token_overlap = 0.0
        candidate_tokens = set(" ".join(candidate_variants).split())
        if target_tokens and candidate_tokens:
            token_overlap = len(target_tokens.intersection(candidate_tokens)) / max(len(target_tokens), len(candidate_tokens))
        score = (seq_score * 0.7) + (token_overlap * 0.3)
        if score > best_score:
            best_score = score
            best_name = candidate_name

    if best_name and best_score >= float(min_similarity):
        return best_name
    return None


def resolve_or_create_personal_name(nombre, cargo=None, create_if_missing=True):
    raw_name = str(nombre or "").strip()
    if not raw_name:
        return ""

    matched_name = _resolve_existing_personal_name(raw_name)
    if matched_name:
        # Si existe pero estaba inactivo, se reactiva para mantener catálogo consistente.
        db = get_db()
        db.execute(
            "UPDATE administradores SET activo = 1, actualizado_en = CURRENT_TIMESTAMP WHERE UPPER(nombre) = UPPER(?)",
            (matched_name,),
        )
        db.commit()
        return matched_name

    if not create_if_missing:
        return raw_name

    db = get_db()
    cursor = db.execute(
        "INSERT INTO administradores (nombre, cargo, activo) VALUES (?, ?, 1)",
        (raw_name, cargo.strip() if cargo else None),
    )
    db.commit()
    if cursor.lastrowid:
        created = db.execute("SELECT nombre FROM administradores WHERE id = ?", (cursor.lastrowid,)).fetchone()
        if created and created["nombre"]:
            return created["nombre"]
    return raw_name

def get_or_create_personal(nombre, cargo=None):
    canonical_name = resolve_or_create_personal_name(nombre, cargo=cargo, create_if_missing=True)
    if not canonical_name:
        return None
    db = get_db()
    row = db.execute("SELECT id FROM administradores WHERE UPPER(nombre) = UPPER(?)", (canonical_name,)).fetchone()
    if row:
        return row["id"]
    return None

# --- HISTORIAL DE ACTAS ---
def save_historial_acta(
    tipo_acta,
    datos_json,
    docx_path,
    pdf_path,
    numero_acta=None,
    plantilla_hash=None,
    plantilla_snapshot_path=None,
):
    db = get_db()
    stored_docx_path = _to_storage_relative_path(docx_path)
    stored_pdf_path = _to_storage_relative_path(pdf_path)
    stored_snapshot_path = _to_storage_relative_path(plantilla_snapshot_path)
    cursor = db.execute(
        """
        INSERT INTO historial_actas (
            tipo_acta,
            numero_acta,
            datos_json,
            docx_path,
            pdf_path,
            plantilla_hash,
            plantilla_snapshot_path
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            tipo_acta,
            numero_acta,
            datos_json,
            stored_docx_path,
            stored_pdf_path,
            plantilla_hash,
            stored_snapshot_path,
        ),
    )
    db.commit()
    return cursor.lastrowid


def _split_numero_acta(numero_acta):
    text = str(numero_acta or "").strip()
    if "-" not in text:
        return None, None
    left, right = text.split("-", 1)
    if not left.isdigit() or not right.isdigit():
        return None, None
    return int(left), int(right)


def _normalize_tipo_acta(tipo_acta):
    text = str(tipo_acta or "").strip().lower()
    return text or "general"


def get_max_numero_acta_for_year(year, tipo_acta=None):
    db = get_db()
    target_year = int(year)
    where_tipo = ""
    params = (target_year,)
    if tipo_acta is not None:
        where_tipo = " AND LOWER(COALESCE(tipo_acta, '')) = LOWER(?)"
        params = (target_year, _normalize_tipo_acta(tipo_acta))

    row = db.execute(
        f"""
        SELECT COALESCE(
            MAX(CAST(substr(numero_acta, 1, instr(numero_acta, '-') - 1) AS INTEGER)),
            0
        ) AS max_value
        FROM historial_actas
        WHERE numero_acta IS NOT NULL
          AND TRIM(numero_acta) != ''
          AND instr(numero_acta, '-') > 1
          AND CAST(substr(numero_acta, instr(numero_acta, '-') + 1) AS INTEGER) = ?
          AND substr(numero_acta, 1, instr(numero_acta, '-') - 1) GLOB '[0-9]*'
          {where_tipo}
        """,
        params,
    ).fetchone()
    return int(row["max_value"] if row else 0)


def get_next_numero_acta(year, tipo_acta=None):
    _ensure_actas_sequence_by_type_table_runtime()
    db = get_db()
    target_year = int(year)
    tipo_key = _normalize_tipo_acta(tipo_acta)
    max_historial = get_max_numero_acta_for_year(target_year, tipo_acta=tipo_key)
    seq_row = db.execute(
        "SELECT ultimo_numero FROM secuencia_actas_tipo WHERE tipo_acta = ? AND anio = ?",
        (tipo_key, target_year),
    ).fetchone()
    max_sequence = int(seq_row["ultimo_numero"] or 0) if seq_row else 0
    next_value = max(max_historial, max_sequence) + 1
    return f"{next_value:03d}-{target_year}"


def reserve_numero_acta(year, preferred_numero_acta=None, tipo_acta=None):
    _ensure_actas_sequence_by_type_table_runtime()
    db = get_db()
    target_year = int(year)
    tipo_key = _normalize_tipo_acta(tipo_acta)

    db.execute("BEGIN IMMEDIATE")
    try:
        row = db.execute(
            "SELECT ultimo_numero FROM secuencia_actas_tipo WHERE tipo_acta = ? AND anio = ?",
            (tipo_key, target_year),
        ).fetchone()

        if row:
            ultimo = int(row["ultimo_numero"] or 0)
            ultimo_historial = get_max_numero_acta_for_year(target_year, tipo_acta=tipo_key)
            if ultimo_historial > ultimo:
                ultimo = ultimo_historial
                db.execute(
                    "UPDATE secuencia_actas_tipo SET ultimo_numero = ?, actualizado_en = CURRENT_TIMESTAMP WHERE tipo_acta = ? AND anio = ?",
                    (ultimo, tipo_key, target_year),
                )
        else:
            ultimo = get_max_numero_acta_for_year(target_year, tipo_acta=tipo_key)
            db.execute(
                "INSERT INTO secuencia_actas_tipo (tipo_acta, anio, ultimo_numero) VALUES (?, ?, ?)",
                (tipo_key, target_year, ultimo),
            )

        if preferred_numero_acta:
            seq, year_in_num = _split_numero_acta(preferred_numero_acta)
            if seq is None or year_in_num != target_year:
                raise ValueError("Numero de acta invalido para reservar.")
            existing = db.execute(
                "SELECT 1 FROM historial_actas WHERE numero_acta = ? AND LOWER(COALESCE(tipo_acta, '')) = LOWER(?) LIMIT 1",
                (f"{seq:03d}-{target_year}", tipo_key),
            ).fetchone()
            if existing:
                raise ValueError("El numero de acta ya existe.")
            reserved = seq
        else:
            reserved = ultimo + 1

        nuevo_ultimo = max(ultimo, reserved)
        db.execute(
            "UPDATE secuencia_actas_tipo SET ultimo_numero = ?, actualizado_en = CURRENT_TIMESTAMP WHERE tipo_acta = ? AND anio = ?",
            (nuevo_ultimo, tipo_key, target_year),
        )
        db.commit()
        return f"{reserved:03d}-{target_year}"
    except Exception:
        db.rollback()
        raise


def numero_acta_exists(numero_acta, tipo_acta=None):
    db = get_db()
    numero = str(numero_acta or "").strip()
    if tipo_acta is None:
        row = db.execute(
            "SELECT 1 FROM historial_actas WHERE numero_acta = ? LIMIT 1",
            (numero,),
        ).fetchone()
    else:
        row = db.execute(
            "SELECT 1 FROM historial_actas WHERE numero_acta = ? AND LOWER(COALESCE(tipo_acta, '')) = LOWER(?) LIMIT 1",
            (numero, _normalize_tipo_acta(tipo_acta)),
        ).fetchone()
    return bool(row)

def get_historial_actas(tipo_acta=None):
    db = get_db()
    order_sql = """
        ORDER BY
            CASE
                WHEN numero_acta IS NOT NULL AND instr(numero_acta, '-') > 0 THEN CAST(substr(numero_acta, instr(numero_acta, '-') + 1) AS INTEGER)
                ELSE 0
            END DESC,
            CASE
                WHEN numero_acta IS NOT NULL AND instr(numero_acta, '-') > 0 THEN CAST(substr(numero_acta, 1, instr(numero_acta, '-') - 1) AS INTEGER)
                ELSE 0
            END DESC,
            id DESC
    """
    if tipo_acta:
        rows = db.execute(
            f"SELECT * FROM historial_actas WHERE tipo_acta = ? {order_sql}",
            (tipo_acta,),
        ).fetchall()
    else:
        rows = db.execute(f"SELECT * FROM historial_actas {order_sql}").fetchall()
    return [_normalize_historial_row_paths(dict(row)) for row in rows]


def count_historial_by_template_snapshot(plantilla_snapshot_path):
    db = get_db()
    normalized_snapshot = _to_storage_relative_path(plantilla_snapshot_path)
    row = db.execute(
        "SELECT COUNT(1) AS total FROM historial_actas WHERE plantilla_snapshot_path = ?",
        (str(normalized_snapshot or "").strip(),),
    ).fetchone()
    return int(row["total"] or 0)


def get_next_numero_informe_area(year):
    db = get_db()
    target_year = int(year)
    row = db.execute(
        "SELECT ultimo_numero FROM secuencia_informes_area WHERE anio = ?",
        (target_year,),
    ).fetchone()
    next_num = (int(row["ultimo_numero"]) + 1) if row else 1
    return f"{next_num:03d}-{target_year}"


def reserve_numeros_informe_area(year, count):
    safe_count = max(int(count or 0), 0)
    if safe_count <= 0:
        return []

    db = get_db()
    target_year = int(year)

    db.execute("BEGIN IMMEDIATE")
    try:
        row = db.execute(
            "SELECT ultimo_numero FROM secuencia_informes_area WHERE anio = ?",
            (target_year,),
        ).fetchone()
        last_num = int(row["ultimo_numero"]) if row else 0
        next_last = last_num + safe_count

        if row:
            db.execute(
                "UPDATE secuencia_informes_area SET ultimo_numero = ?, actualizado_en = CURRENT_TIMESTAMP WHERE anio = ?",
                (next_last, target_year),
            )
        else:
            db.execute(
                "INSERT INTO secuencia_informes_area (anio, ultimo_numero) VALUES (?, ?)",
                (target_year, next_last),
            )

        db.commit()
    except Exception:
        db.rollback()
        raise

    return [f"{n:03d}-{target_year}" for n in range(last_num + 1, next_last + 1)]


def delete_historial_acta(acta_id):
    db = get_db()
    row = db.execute("SELECT * FROM historial_actas WHERE id = ?", (acta_id,)).fetchone()
    deleted = _normalize_historial_row_paths(dict(row)) if row else None
    db.execute("DELETE FROM historial_actas WHERE id = ?", (acta_id,))
    db.commit()
    return deleted
