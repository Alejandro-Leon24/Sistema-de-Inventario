import sqlite3
import json
import re
import unicodedata
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
    "justificacion",
    "procedencia",
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
    "justificacion",
    "procedencia",
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


def _build_default_import_procedencia_text(base_date=None):
    dt = base_date or datetime.now()
    date_text = dt.strftime("%d/%m/%Y %H:%M:%S")
    return f"Exportación Masiva de Excel - {date_text} / Bienes propios de la facultad"

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


def _normalize_inventory_date_for_output(value):
    text = str(value or "").strip()
    if not text:
        return ""

    iso_match = re.match(r"^(\d{4}-\d{2}-\d{2})", text)
    if iso_match:
        return iso_match.group(1)

    dmy_match = re.match(r"^(\d{1,2}[/-]\d{1,2}[/-]\d{4})", text)
    if dmy_match:
        raw = dmy_match.group(1)
        parts = re.split(r"[/-]", raw)
        if len(parts) == 3:
            day, month, year = parts
            return f"{year}-{month.zfill(2)}-{day.zfill(2)}"

    return text


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
        "fecha_adquisicion": _normalize_inventory_date_for_output(row["fecha_adquisicion"]),
        "valor": row["valor"],
        "observacion": row["observacion"],
        "justificacion": row["justificacion"],
        "procedencia": row["procedencia"],
        "descripcion_esbye": row["descripcion_esbye"],
        "marca_esbye": row["marca_esbye"],
        "modelo_esbye": row["modelo_esbye"],
        "serie_esbye": row["serie_esbye"],
        "valor_esbye": row["valor_esbye"],
        "ubicacion_esbye": row["ubicacion_esbye"],
        "observacion_esbye": row["observacion_esbye"],
        "fecha_adquisicion_esbye": _normalize_inventory_date_for_output(row["fecha_adquisicion_esbye"]),
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
    _run_startup_migration_once("20260419_inventory_baja_columns_and_default_procedencia", _ensure_inventory_baja_columns_and_default_procedencia)
    _run_startup_migration_once("20260409_inventory_codes_allow_duplicates", _ensure_inventory_codes_allow_duplicates)
    _run_startup_migration_once("20260409_inventory_search_indexes", _ensure_inventory_search_indexes)
    _run_startup_migration_once("20260409_inventory_fts", _ensure_inventory_fts)
    _run_startup_migration_once("20260416_inventory_search_indexes_extended", _ensure_inventory_search_indexes)
    _run_startup_migration_once("20260416_inventory_fts_extended", _ensure_inventory_fts)
    _run_startup_migration_once("20260416_inventory_estado_canonical", _ensure_inventory_estado_canonical_runtime)
    _run_startup_migration_once("20260409_historial_actas_numero_column", _ensure_historial_actas_numero_column)
    _run_startup_migration_once("20260409_historial_actas_template_columns", _ensure_historial_actas_template_columns)
    _run_startup_migration_once("20260419_acta_inventory_mutaciones_table", _ensure_acta_inventory_mutaciones_table)
    _run_startup_migration_once("20260409_informes_area_sequence_table", _ensure_informes_area_sequence_table)
    _run_startup_migration_once("20260409_actas_sequence_table", _ensure_actas_sequence_table)
    _run_startup_migration_once("20260414_historial_actas_numero_unique_by_type", _ensure_historial_actas_numero_unique_by_type)
    _run_startup_migration_once("20260414_actas_sequence_by_type_table", _ensure_actas_sequence_by_type_table)
    _run_startup_migration_once("20260409_seed_default_param_values", _seed_default_param_values)

    # Refuerzo runtime para evitar esquemas de busqueda desactualizados en equipos que ya migraron antes.
    _ensure_inventory_search_indexes()
    _ensure_inventory_fts()


def _ensure_inventory_estado_canonical_runtime():
    db = get_db()
    estado_options = _get_estado_catalog_options()
    if not estado_options:
        return

    normalized_to_canonical = {}
    for name in estado_options:
        key = _normalize_catalog_text(name)
        if key and key not in normalized_to_canonical:
            normalized_to_canonical[key] = name

    if not normalized_to_canonical:
        return

    rows = db.execute(
        "SELECT id, estado FROM inventario_items WHERE TRIM(COALESCE(estado, '')) <> ''"
    ).fetchall()
    updates = []
    for row in rows:
        current = str(row["estado"] or "").strip()
        if not current:
            continue
        canonical = normalized_to_canonical.get(_normalize_catalog_text(current))
        if canonical and canonical != current:
            updates.append((canonical, row["id"]))

    if updates:
        db.executemany("UPDATE inventario_items SET estado = ? WHERE id = ?", updates)
        db.commit()


def _ensure_historial_actas_numero_column():
    db = get_db()
    existing_columns = {row["name"] for row in db.execute("PRAGMA table_info(historial_actas)").fetchall()}

    if "numero_acta" not in existing_columns:
        db.execute("ALTER TABLE historial_actas ADD COLUMN numero_acta TEXT")

    db.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_historial_actas_numero_acta ON historial_actas(numero_acta)")
    db.commit()


def _ensure_historial_actas_numero_unique_by_type():
    _ensure_historial_actas_numero_unique_by_type_runtime()
    get_db().commit()


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


def _ensure_acta_inventory_mutaciones_table():
    db = get_db()
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS acta_inventory_mutaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            acta_id INTEGER NOT NULL,
            tipo_acta TEXT NOT NULL,
            item_id INTEGER,
            mutation_kind TEXT NOT NULL,
            before_data_json TEXT,
            after_data_json TEXT,
            creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (acta_id) REFERENCES historial_actas(id) ON DELETE CASCADE
        )
        """
    )
    db.execute(
        "CREATE INDEX IF NOT EXISTS idx_acta_inventory_mutaciones_acta_id ON acta_inventory_mutaciones(acta_id)"
    )
    db.execute(
        "CREATE INDEX IF NOT EXISTS idx_acta_inventory_mutaciones_item_id ON acta_inventory_mutaciones(item_id)"
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


def _ensure_historial_actas_table_without_global_unique_runtime():
    """Reconstruye historial_actas si quedó con UNIQUE global en numero_acta."""
    db = get_db()
    has_global_unique_numero = False
    for idx in db.execute("PRAGMA index_list(historial_actas)").fetchall():
        if not int(idx["unique"] or 0):
            continue
        idx_name = str(idx["name"] or "").strip()
        if not idx_name:
            continue
        cols = [
            str(col["name"] or "").strip().lower()
            for col in db.execute(f"PRAGMA index_info({idx_name})").fetchall()
        ]
        if cols == ["numero_acta"]:
            has_global_unique_numero = True
            break

    if not has_global_unique_numero:
        table_sql_row = db.execute(
            "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'historial_actas'"
        ).fetchone()
        table_sql = str((table_sql_row["sql"] if table_sql_row else "") or "")
        normalized_sql = " ".join(table_sql.upper().replace('"', "").split())
        has_global_unique_numero = (
            "NUMERO_ACTA TEXT UNIQUE" in normalized_sql
            or "UNIQUE(NUMERO_ACTA)" in normalized_sql
            or "UNIQUE (NUMERO_ACTA)" in normalized_sql
        )

    if not has_global_unique_numero:
        return

    existing_columns = {row["name"] for row in db.execute("PRAGMA table_info(historial_actas)").fetchall()}
    if "plantilla_hash" not in existing_columns:
        db.execute("ALTER TABLE historial_actas ADD COLUMN plantilla_hash TEXT")
        existing_columns.add("plantilla_hash")
    if "plantilla_snapshot_path" not in existing_columns:
        db.execute("ALTER TABLE historial_actas ADD COLUMN plantilla_snapshot_path TEXT")
        existing_columns.add("plantilla_snapshot_path")

    db.execute("BEGIN IMMEDIATE")
    try:
        db.execute(
            """
            CREATE TABLE IF NOT EXISTS historial_actas_tmp (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo_acta TEXT NOT NULL,
                numero_acta TEXT,
                fecha TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                datos_json TEXT,
                docx_path TEXT,
                pdf_path TEXT,
                plantilla_hash TEXT,
                plantilla_snapshot_path TEXT
            )
            """
        )

        db.execute(
            """
            INSERT INTO historial_actas_tmp (
                id,
                tipo_acta,
                numero_acta,
                fecha,
                datos_json,
                docx_path,
                pdf_path,
                plantilla_hash,
                plantilla_snapshot_path
            )
            SELECT
                id,
                tipo_acta,
                numero_acta,
                fecha,
                datos_json,
                docx_path,
                pdf_path,
                plantilla_hash,
                plantilla_snapshot_path
            FROM historial_actas
            """
        )

        db.execute("DROP TABLE historial_actas")
        db.execute("ALTER TABLE historial_actas_tmp RENAME TO historial_actas")
        db.execute("CREATE INDEX IF NOT EXISTS idx_historial_actas_plantilla_snapshot_path ON historial_actas(plantilla_snapshot_path)")
        db.commit()
    except Exception:
        db.rollback()
        raise


def _ensure_historial_actas_numero_unique_by_type_runtime():
    """Asegura en runtime que la unicidad de numero_acta sea por tipo_acta."""
    db = get_db()
    _ensure_historial_actas_table_without_global_unique_runtime()
    db.execute("DROP INDEX IF EXISTS idx_historial_actas_numero_acta")
    db.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS idx_historial_actas_tipo_numero_acta
        ON historial_actas(tipo_acta, numero_acta)
        WHERE numero_acta IS NOT NULL AND TRIM(numero_acta) != ''
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
        "justificacion": "TEXT",
        "procedencia": "TEXT",
    }

    for column_name, column_type in required_columns.items():
        if column_name not in existing_columns:
            db.execute(f"ALTER TABLE inventario_items ADD COLUMN {column_name} {column_type}")
    db.commit()


def _ensure_inventory_baja_columns_and_default_procedencia():
    db = get_db()
    existing_columns = {
        row["name"]
        for row in db.execute("PRAGMA table_info(inventario_items)").fetchall()
    }

    if "justificacion" not in existing_columns:
        db.execute("ALTER TABLE inventario_items ADD COLUMN justificacion TEXT")
    if "procedencia" not in existing_columns:
        db.execute("ALTER TABLE inventario_items ADD COLUMN procedencia TEXT")

    default_procedencia = _build_default_import_procedencia_text()
    db.execute(
        "UPDATE inventario_items SET procedencia = ?",
        (default_procedencia,),
    )
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
                justificacion TEXT,
                procedencia TEXT,
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
                fecha_adquisicion, valor, observacion, justificacion, procedencia, descripcion_esbye, marca_esbye,
                modelo_esbye, serie_esbye, valor_esbye, ubicacion_esbye, observacion_esbye,
                fecha_adquisicion_esbye, area_id, creado_en, actualizado_en
            )
            SELECT
                id, item_numero, cod_inventario, cod_esbye, cuenta, cantidad, descripcion,
                ubicacion, marca, modelo, serie, estado, condicion, usuario_final,
                fecha_adquisicion, valor, observacion, justificacion, procedencia, descripcion_esbye, marca_esbye,
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
    from database.services.locations_service import get_structure as _get_structure

    return _get_structure(include_area_details=include_area_details)


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
    from database.services.locations_service import create_block as _create_block

    return _create_block(nombre, descripcion=descripcion)


def update_block(block_id, nombre=None, descripcion=None):
    from database.services.locations_service import update_block as _update_block

    return _update_block(block_id, nombre=nombre, descripcion=descripcion)


def create_floor(bloque_id, nombre, descripcion=None):
    from database.services.locations_service import create_floor as _create_floor

    return _create_floor(bloque_id, nombre, descripcion=descripcion)


def create_area(piso_id, nombre, descripcion=None, details=None):
    from database.services.locations_service import create_area as _create_area

    return _create_area(piso_id, nombre, descripcion=descripcion, details=details)


def update_floor(floor_id, nombre=None, descripcion=None):
    from database.services.locations_service import update_floor as _update_floor

    return _update_floor(floor_id, nombre=nombre, descripcion=descripcion)


def delete_floor(floor_id):
    from database.services.locations_service import delete_floor as _delete_floor

    return _delete_floor(floor_id)


def update_area(area_id, nombre=None, descripcion=None, details=None):
    from database.services.locations_service import update_area as _update_area

    return _update_area(area_id, nombre=nombre, descripcion=descripcion, details=details)


def delete_area(area_id):
    from database.services.locations_service import delete_area as _delete_area

    return _delete_area(area_id)


def delete_block(block_id):
    from database.services.locations_service import delete_block as _delete_block

    return _delete_block(block_id)


def get_location_dependency_summary(entity_type, entity_id):
    from database.services.locations_service import get_location_dependency_summary as _get_location_dependency_summary

    return _get_location_dependency_summary(entity_type, entity_id)

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
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_cuenta ON inventario_items(cuenta)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_estado ON inventario_items(estado)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_marca ON inventario_items(marca)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_modelo ON inventario_items(modelo)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_serie ON inventario_items(serie)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_usuario_final ON inventario_items(usuario_final)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_observacion ON inventario_items(observacion)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_justificacion ON inventario_items(justificacion)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_procedencia ON inventario_items(procedencia)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_desc_esbye ON inventario_items(descripcion_esbye)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_marca_esbye ON inventario_items(marca_esbye)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_modelo_esbye ON inventario_items(modelo_esbye)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_serie_esbye ON inventario_items(serie_esbye)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_ubicacion_esbye ON inventario_items(ubicacion_esbye)")
    db.execute("CREATE INDEX IF NOT EXISTS idx_inventario_observacion_esbye ON inventario_items(observacion_esbye)")
    db.commit()


def _ensure_inventory_fts():
    db = get_db()
    try:
        db.execute("DROP TRIGGER IF EXISTS inventario_items_ai")
        db.execute("DROP TRIGGER IF EXISTS inventario_items_ad")
        db.execute("DROP TRIGGER IF EXISTS inventario_items_au")
        db.execute("DROP TABLE IF EXISTS inventario_items_fts")
        db.execute(
            """
            CREATE VIRTUAL TABLE IF NOT EXISTS inventario_items_fts
            USING fts5(
                descripcion,
                cod_inventario,
                cod_esbye,
                cuenta,
                estado,
                marca,
                modelo,
                serie,
                usuario_final,
                ubicacion,
                observacion,
                justificacion,
                procedencia,
                descripcion_esbye,
                marca_esbye,
                modelo_esbye,
                serie_esbye,
                ubicacion_esbye,
                observacion_esbye,
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
                INSERT INTO inventario_items_fts(
                    rowid, descripcion, cod_inventario, cod_esbye, cuenta, estado,
                    marca, modelo, serie, usuario_final, ubicacion, observacion,
                    justificacion, procedencia,
                    descripcion_esbye, marca_esbye, modelo_esbye, serie_esbye,
                    ubicacion_esbye, observacion_esbye
                )
                VALUES (
                    new.id, new.descripcion, new.cod_inventario, new.cod_esbye, new.cuenta, new.estado,
                    new.marca, new.modelo, new.serie, new.usuario_final, new.ubicacion, new.observacion,
                    new.justificacion, new.procedencia,
                    new.descripcion_esbye, new.marca_esbye, new.modelo_esbye, new.serie_esbye,
                    new.ubicacion_esbye, new.observacion_esbye
                );
            END
            """
        )
        db.execute(
            """
            CREATE TRIGGER IF NOT EXISTS inventario_items_ad
            AFTER DELETE ON inventario_items
            BEGIN
                INSERT INTO inventario_items_fts(
                    inventario_items_fts, rowid, descripcion, cod_inventario, cod_esbye, cuenta, estado,
                    marca, modelo, serie, usuario_final, ubicacion, observacion,
                    justificacion, procedencia,
                    descripcion_esbye, marca_esbye, modelo_esbye, serie_esbye,
                    ubicacion_esbye, observacion_esbye
                )
                VALUES (
                    'delete', old.id, old.descripcion, old.cod_inventario, old.cod_esbye, old.cuenta, old.estado,
                    old.marca, old.modelo, old.serie, old.usuario_final, old.ubicacion, old.observacion,
                    old.justificacion, old.procedencia,
                    old.descripcion_esbye, old.marca_esbye, old.modelo_esbye, old.serie_esbye,
                    old.ubicacion_esbye, old.observacion_esbye
                );
            END
            """
        )
        db.execute(
            """
            CREATE TRIGGER IF NOT EXISTS inventario_items_au
            AFTER UPDATE ON inventario_items
            BEGIN
                INSERT INTO inventario_items_fts(
                    inventario_items_fts, rowid, descripcion, cod_inventario, cod_esbye, cuenta, estado,
                    marca, modelo, serie, usuario_final, ubicacion, observacion,
                    justificacion, procedencia,
                    descripcion_esbye, marca_esbye, modelo_esbye, serie_esbye,
                    ubicacion_esbye, observacion_esbye
                )
                VALUES (
                    'delete', old.id, old.descripcion, old.cod_inventario, old.cod_esbye, old.cuenta, old.estado,
                    old.marca, old.modelo, old.serie, old.usuario_final, old.ubicacion, old.observacion,
                    old.justificacion, old.procedencia,
                    old.descripcion_esbye, old.marca_esbye, old.modelo_esbye, old.serie_esbye,
                    old.ubicacion_esbye, old.observacion_esbye
                );
                INSERT INTO inventario_items_fts(
                    rowid, descripcion, cod_inventario, cod_esbye, cuenta, estado,
                    marca, modelo, serie, usuario_final, ubicacion, observacion,
                    justificacion, procedencia,
                    descripcion_esbye, marca_esbye, modelo_esbye, serie_esbye,
                    ubicacion_esbye, observacion_esbye
                )
                VALUES (
                    new.id, new.descripcion, new.cod_inventario, new.cod_esbye, new.cuenta, new.estado,
                    new.marca, new.modelo, new.serie, new.usuario_final, new.ubicacion, new.observacion,
                    new.justificacion, new.procedencia,
                    new.descripcion_esbye, new.marca_esbye, new.modelo_esbye, new.serie_esbye,
                    new.ubicacion_esbye, new.observacion_esbye
                );
            END
            """
        )
        db.execute("INSERT INTO inventario_items_fts(inventario_items_fts) VALUES ('rebuild')")
        db.commit()
    except sqlite3.OperationalError:
        # Some SQLite builds might not include FTS5; LIKE fallback remains active.
        db.rollback()


def _has_inventory_fts(required_columns=None):
    db = get_db()
    row = db.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'inventario_items_fts'"
    ).fetchone()
    if row is None:
        return False
    if not required_columns:
        return True

    fts_columns = {
        str(col["name"] or "").strip().lower()
        for col in db.execute("PRAGMA table_info(inventario_items_fts)").fetchall()
    }
    required = {str(name).strip().lower() for name in (required_columns or set())}
    return required.issubset(fts_columns)


def _build_fts_query(raw_search):
    # Evita que caracteres como '-' generen expresiones invalidas en FTS5
    # (ej. FJCP-000155 podria interpretarse como operador/columna).
    tokens = re.findall(r"[A-Za-z0-9_]+", str(raw_search or ""), flags=re.UNICODE)
    if not tokens:
        return ""
    safe_tokens = [str(token).replace('"', '""') for token in tokens if str(token).strip()]
    return " AND ".join([f'"{token}"*' for token in safe_tokens])


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
        required_fts_columns = {
            "descripcion",
            "cod_inventario",
            "cod_esbye",
            "cuenta",
            "estado",
            "marca",
            "modelo",
            "serie",
            "usuario_final",
            "ubicacion",
            "observacion",
            "justificacion",
            "procedencia",
            "descripcion_esbye",
            "marca_esbye",
            "modelo_esbye",
            "serie_esbye",
            "ubicacion_esbye",
            "observacion_esbye",
        }
        if fts_query and not _has_inventory_fts(required_columns=required_fts_columns):
            _ensure_inventory_fts()

        if fts_query and _has_inventory_fts(required_columns=required_fts_columns):
            where_clauses.append(
                "i.id IN (SELECT rowid FROM inventario_items_fts WHERE inventario_items_fts MATCH ?)"
            )
            params.append(fts_query)
        else:
            token = f"%{raw_search}%"
            where_clauses.append(
                "(i.descripcion LIKE ? OR i.cod_inventario LIKE ? OR i.cod_esbye LIKE ? OR i.cuenta LIKE ? OR i.estado LIKE ? OR i.condicion LIKE ? OR i.ubicacion LIKE ? OR i.marca LIKE ? OR i.modelo LIKE ? OR i.serie LIKE ? OR i.usuario_final LIKE ? OR i.observacion LIKE ? OR i.justificacion LIKE ? OR i.procedencia LIKE ? OR i.descripcion_esbye LIKE ? OR i.marca_esbye LIKE ? OR i.modelo_esbye LIKE ? OR i.serie_esbye LIKE ? OR i.ubicacion_esbye LIKE ? OR i.observacion_esbye LIKE ? OR i.fecha_adquisicion LIKE ? OR i.fecha_adquisicion_esbye LIKE ?)"
            )
            params.extend([
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
                token,
            ])

    where_sql = ""
    if where_clauses:
        where_sql = "WHERE " + " AND ".join(where_clauses)

    return where_sql, params


def _normalize_area_lookup_token(value):
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", text)


def _extract_compact_aula_code_for_lookup(text):
    normalized = _normalize_area_lookup_token(text)
    if not normalized:
        return ""
    match = re.search(r"(\d+[a-z]\s*-?\s*\d+)", normalized)
    if not match:
        return ""
    return re.sub(r"\s+", "", match.group(1))


def _extract_floor_hint_for_lookup(text):
    normalized = _normalize_area_lookup_token(text)
    if "planta baja" in normalized:
        return "planta baja"
    if "primer" in normalized and "piso" in normalized:
        return "primer piso"
    if "segundo" in normalized and "piso" in normalized:
        return "segundo piso"
    if "tercer" in normalized and "piso" in normalized:
        return "tercer piso"
    return ""


def _build_area_lookup_maps(db):
    rows = db.execute(
        """
        SELECT a.id, a.nombre AS area_nombre, p.nombre AS piso_nombre, b.nombre AS bloque_nombre
        FROM areas a
        JOIN pisos p ON p.id = a.piso_id
        JOIN bloques b ON b.id = p.bloque_id
        """
    ).fetchall()

    by_full_name = {}
    by_area_name = {}
    candidates = []
    for row in rows:
        area_id = int(row["id"])
        block_name = str(row["bloque_nombre"] or "").strip()
        floor_name = str(row["piso_nombre"] or "").strip()
        area_name = str(row["area_nombre"] or "").strip()
        full_name = " / ".join([part for part in [block_name, floor_name, area_name] if part])
        full_key = _normalize_area_lookup_token(full_name)
        if full_key:
            by_full_name[full_key] = {
                "id": area_id,
                "display": full_name,
            }

        area_key = _normalize_area_lookup_token(area_name)
        if area_key:
            by_area_name.setdefault(area_key, []).append(
                {
                    "id": area_id,
                    "display": full_name,
                }
            )

        candidates.append(
            {
                "id": area_id,
                "display": full_name,
                "area_name": _normalize_area_lookup_token(area_name),
                "floor_name": _normalize_area_lookup_token(floor_name),
                "block_name": _normalize_area_lookup_token(block_name),
                "full_name": _normalize_area_lookup_token(f"{block_name} {floor_name} {area_name}"),
            }
        )

    return by_full_name, by_area_name, candidates


def _resolve_area_from_location_text(raw_location, by_full_name, by_area_name, candidates):
    normalized_input = _normalize_area_lookup_token(raw_location)
    if not normalized_input:
        return None

    direct = by_full_name.get(normalized_input)
    if direct:
        return direct

    compact = re.sub(r"\s*/\s*", " / ", normalized_input)
    direct = by_full_name.get(compact)
    if direct:
        return direct

    by_area = by_area_name.get(normalized_input)
    if by_area and len(by_area) == 1:
        return by_area[0]

    # Fallback: heurística equivalente al flujo de importación Excel.
    target_tokens = [token for token in re.split(r"\s+|/|-", normalized_input) if token]
    target_aula_code = _extract_compact_aula_code_for_lookup(normalized_input)
    floor_hint = _extract_floor_hint_for_lookup(normalized_input)
    explicit_block_letter = ""
    block_token_match = re.search(r"bloque\s+([a-z])", normalized_input)
    if block_token_match:
        explicit_block_letter = block_token_match.group(1)

    best_match = None
    best_score = 0

    for candidate in candidates:
        area_name = candidate.get("area_name") or ""
        floor_name = candidate.get("floor_name") or ""
        block_name = candidate.get("block_name") or ""
        full_name = candidate.get("full_name") or ""

        score = 0
        if normalized_input == area_name or normalized_input == full_name:
            score += 30
        if normalized_input in full_name:
            score += 10

        area_code = _extract_compact_aula_code_for_lookup(area_name)
        if target_aula_code and area_code and target_aula_code == area_code:
            score += 35

        area_tokens = [token for token in re.split(r"\s+|/|-", full_name) if token]
        token_hits = 0
        for token in target_tokens:
            if len(token) < 3:
                continue
            if any(token in candidate_token or candidate_token in token for candidate_token in area_tokens):
                token_hits += 1
        score += min(token_hits * 2, 16)

        if floor_hint and floor_hint in floor_name:
            score += 8

        block_letter_match = re.search(r"\d+([a-z])", target_aula_code or "")
        if block_letter_match:
            block_letter = block_letter_match.group(1)
            if f"bloque {block_letter}" in block_name:
                score += 10
            else:
                score -= 14

        if explicit_block_letter:
            if f"bloque {explicit_block_letter}" in block_name:
                score += 12
            else:
                score -= 16

        if score > best_score:
            best_score = score
            best_match = candidate

    if best_match and best_score >= 10:
        return {
            "id": best_match["id"],
            "display": best_match["display"],
        }

    return None


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
            i.justificacion,
            i.procedencia,
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
            i.justificacion,
            i.procedencia,
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
            i.justificacion,
            i.procedencia,
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


def _normalize_catalog_text(value):
    text = str(value or "").strip().lower()
    if not text:
        return ""
    text = (
        text.replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("ü", "u")
    )
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    return " ".join(text.split())


def _resolve_catalog_name(raw_value, options):
    value = str(raw_value or "").strip()
    if not value:
        return ""

    normalized_value = _normalize_catalog_text(value)
    clean_options = [str(opt or "").strip() for opt in (options or []) if str(opt or "").strip()]
    if not normalized_value or not clean_options:
        return value

    exact = next((opt for opt in clean_options if _normalize_catalog_text(opt) == normalized_value), None)
    if exact:
        return exact

    contains = next(
        (
            opt
            for opt in clean_options
            if _normalize_catalog_text(opt) in normalized_value or normalized_value in _normalize_catalog_text(opt)
        ),
        None,
    )
    if contains:
        return contains

    target_tokens = [tok for tok in normalized_value.split() if tok]
    best_option = ""
    best_score = 0.0
    for opt in clean_options:
        opt_norm = _normalize_catalog_text(opt)
        if not opt_norm:
            continue
        opt_tokens = [tok for tok in opt_norm.split() if tok]
        overlap = sum(
            1
            for tok in target_tokens
            if any(tok in opt_tok or opt_tok in tok for opt_tok in opt_tokens)
        )
        score = (overlap / len(target_tokens)) if target_tokens else 0.0
        if score > best_score:
            best_score = score
            best_option = opt

    return best_option if best_score >= 0.6 else value


def _get_estado_catalog_options():
    db = get_db()
    rows = db.execute(
        "SELECT nombre FROM param_estados WHERE TRIM(COALESCE(nombre, '')) <> '' ORDER BY nombre ASC"
    ).fetchall()
    return [str(row["nombre"]).strip() for row in rows if str(row["nombre"] or "").strip()]


def _canonicalize_estado_value(value, estado_options=None):
    raw = str(value or "").strip()
    if not raw:
        return raw
    options = estado_options if estado_options is not None else _get_estado_catalog_options()
    return _resolve_catalog_name(raw, options)


def create_inventory_item(payload, commit=True):
    db = get_db()
    fields = {k: payload.get(k) for k in ALLOWED_INVENTORY_FIELDS if k in payload}
    _normalize_inventory_code_fields(fields)
    if fields.get("estado") not in (None, ""):
        fields["estado"] = _canonicalize_estado_value(fields.get("estado"))
    if fields.get("usuario_final") not in (None, ""):
        fields["usuario_final"] = resolve_or_create_personal_name(
            fields.get("usuario_final"),
            create_if_missing=True,
        )
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
        if field == "estado" and value not in (None, ""):
            value = _canonicalize_estado_value(value)
        if field == "usuario_final" and value not in (None, ""):
            value = resolve_or_create_personal_name(value, create_if_missing=True)
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


def bulk_insert_inventory_rows(rows, area_id=None, procedencia_default=None):
    db = get_db()
    normalized_values = []
    insert_columns = ["item_numero", *CANONICAL_COLUMN_ORDER, "area_id"]
    missing_area_names = set()
    estado_options = _get_estado_catalog_options()
    try:
        db.execute("BEGIN IMMEDIATE")
        start_item_numero = _next_item_numero_in_tx(db)
        by_full_name, by_area_name, area_candidates = _build_area_lookup_maps(db)
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

            if not payload.get("area_id") and payload.get("ubicacion") not in (None, ""):
                resolved_area = _resolve_area_from_location_text(
                    payload.get("ubicacion"),
                    by_full_name,
                    by_area_name,
                    area_candidates,
                )
                if resolved_area:
                    payload["area_id"] = resolved_area["id"]
                    payload["ubicacion"] = resolved_area["display"]
                else:
                    missing_area_names.add(str(payload.get("ubicacion") or "").strip())

            if payload.get("usuario_final") not in (None, ""):
                payload["usuario_final"] = resolve_or_create_personal_name(
                    payload.get("usuario_final"),
                    create_if_missing=True,
                )
            if payload.get("estado") not in (None, ""):
                payload["estado"] = _canonicalize_estado_value(payload.get("estado"), estado_options)
            if procedencia_default and not str(payload.get("procedencia") or "").strip():
                payload["procedencia"] = str(procedencia_default).strip()

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

            normalized_values.append(tuple(payload.get(column) for column in insert_columns))

        missing_area_names.discard("")
        if missing_area_names:
            raise ValueError(json.dumps(sorted(missing_area_names), ensure_ascii=False))

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
        return {"inserted_ids": [], "missing_area_names": []}

    last_insert_rowid = db.execute("SELECT last_insert_rowid() AS last_id").fetchone()["last_id"]
    first_insert_rowid = last_insert_rowid - len(normalized_values) + 1
    return {
        "inserted_ids": list(range(first_insert_rowid, last_insert_rowid + 1)),
        "missing_area_names": [],
    }


def bulk_insert_inventory_dicts(rows_as_dicts, area_id=None, procedencia_default=None):
    db = get_db()
    insert_columns = ["item_numero", *CANONICAL_COLUMN_ORDER, "area_id"]
    normalized_values = []
    skipped = 0
    estado_options = _get_estado_catalog_options()

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

            if payload.get("estado") not in (None, ""):
                payload["estado"] = _canonicalize_estado_value(payload.get("estado"), estado_options)
            if procedencia_default and not str(payload.get("procedencia") or "").strip():
                payload["procedencia"] = str(procedencia_default).strip()

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
                i.justificacion,
                i.procedencia,
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
    total_bloques = db.execute('SELECT COUNT(1) as total FROM bloques').fetchone()['total']
    total_areas = db.execute('SELECT COUNT(1) as total FROM areas').fetchone()['total']
    return {
        'cant_bienes': total_bienes or 0,
        'cant_bloques': total_bloques or 0,
        'cant_areas': total_areas or 0,
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
        "mrt", "prof", "profa", "tlgo", "tlga", "ts", "phd",
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
    from database.services.documents_service import get_or_create_personal as _get_or_create_personal

    return _get_or_create_personal(nombre, cargo=cargo)

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
    from database.services.documents_service import save_historial_acta as _save_historial_acta

    return _save_historial_acta(
        tipo_acta=tipo_acta,
        datos_json=datos_json,
        docx_path=docx_path,
        pdf_path=pdf_path,
        numero_acta=numero_acta,
        plantilla_hash=plantilla_hash,
        plantilla_snapshot_path=plantilla_snapshot_path,
    )


def update_historial_acta(
    acta_id,
    tipo_acta,
    datos_json,
    docx_path,
    pdf_path,
    numero_acta=None,
    plantilla_hash=None,
    plantilla_snapshot_path=None,
):
    from database.services.documents_service import update_historial_acta as _update_historial_acta

    return _update_historial_acta(
        acta_id=acta_id,
        tipo_acta=tipo_acta,
        datos_json=datos_json,
        docx_path=docx_path,
        pdf_path=pdf_path,
        numero_acta=numero_acta,
        plantilla_hash=plantilla_hash,
        plantilla_snapshot_path=plantilla_snapshot_path,
    )


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
    from database.services.documents_service import get_max_numero_acta_for_year as _get_max_numero_acta_for_year

    return _get_max_numero_acta_for_year(year=year, tipo_acta=tipo_acta)


def get_next_numero_acta(year, tipo_acta=None):
    from database.services.documents_service import get_next_numero_acta as _get_next_numero_acta

    return _get_next_numero_acta(year=year, tipo_acta=tipo_acta)


def reserve_numero_acta(year, preferred_numero_acta=None, tipo_acta=None):
    from database.services.documents_service import reserve_numero_acta as _reserve_numero_acta

    return _reserve_numero_acta(
        year=year,
        preferred_numero_acta=preferred_numero_acta,
        tipo_acta=tipo_acta,
    )


def numero_acta_exists(numero_acta, tipo_acta=None):
    from database.services.documents_service import numero_acta_exists as _numero_acta_exists

    return _numero_acta_exists(numero_acta=numero_acta, tipo_acta=tipo_acta)

def get_historial_actas(tipo_acta=None):
    from database.services.documents_service import get_historial_actas as _get_historial_actas

    return _get_historial_actas(tipo_acta=tipo_acta)


def count_historial_by_template_snapshot(plantilla_snapshot_path):
    from database.services.documents_service import count_historial_by_template_snapshot as _count_historial_by_template_snapshot

    return _count_historial_by_template_snapshot(plantilla_snapshot_path)


def get_next_numero_informe_area(year):
    from database.services.documents_service import get_next_numero_informe_area as _get_next_numero_informe_area

    return _get_next_numero_informe_area(year)


def reserve_numeros_informe_area(year, count):
    from database.services.documents_service import reserve_numeros_informe_area as _reserve_numeros_informe_area

    return _reserve_numeros_informe_area(year, count)


def delete_historial_acta(acta_id):
    from database.services.documents_service import delete_historial_acta as _delete_historial_acta

    return _delete_historial_acta(acta_id)
