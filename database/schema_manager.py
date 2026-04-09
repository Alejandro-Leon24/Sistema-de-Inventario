from pathlib import Path

from database.db import execute_schema
from database import controller as legacy_controller


def init_schema(base_dir: Path):
    schema_path = base_dir / "database" / "schema.sql"
    execute_schema(schema_path)
    legacy_controller._ensure_schema_migrations_table()

    legacy_controller._run_startup_migration_once(
        "20260409_university_unique_constraint",
        legacy_controller._ensure_university_unique_constraint,
    )
    legacy_controller._run_startup_migration_once(
        "20260409_area_extended_columns",
        legacy_controller._ensure_area_extended_columns,
    )
    legacy_controller._run_startup_migration_once(
        "20260409_inventory_extended_columns",
        legacy_controller._ensure_inventory_extended_columns,
    )
    legacy_controller._run_startup_migration_once(
        "20260409_inventory_codes_allow_duplicates",
        legacy_controller._ensure_inventory_codes_allow_duplicates,
    )
    legacy_controller._run_startup_migration_once(
        "20260409_inventory_search_indexes",
        legacy_controller._ensure_inventory_search_indexes,
    )
    legacy_controller._run_startup_migration_once(
        "20260409_inventory_fts",
        legacy_controller._ensure_inventory_fts,
    )
    legacy_controller._run_startup_migration_once(
        "20260409_historial_actas_numero_column",
        legacy_controller._ensure_historial_actas_numero_column,
    )
    legacy_controller._run_startup_migration_once(
        "20260409_historial_actas_template_columns",
        legacy_controller._ensure_historial_actas_template_columns,
    )
    legacy_controller._run_startup_migration_once(
        "20260409_informes_area_sequence_table",
        legacy_controller._ensure_informes_area_sequence_table,
    )
    legacy_controller._run_startup_migration_once(
        "20260409_actas_sequence_table",
        legacy_controller._ensure_actas_sequence_table,
    )
    legacy_controller._run_startup_migration_once(
        "20260409_seed_default_param_values",
        legacy_controller._seed_default_param_values,
    )
