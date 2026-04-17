"""Servicio de dominio para historial y numeracion de actas."""

from pathlib import Path

from database.db import get_db

PROJECT_ROOT = Path(__file__).resolve().parents[2]


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


def get_or_create_personal(nombre, cargo=None):
    from database.controller import resolve_or_create_personal_name

    canonical_name = resolve_or_create_personal_name(nombre, cargo=cargo, create_if_missing=True)
    if not canonical_name:
        return None
    db = get_db()
    row = db.execute("SELECT id FROM administradores WHERE UPPER(nombre) = UPPER(?)", (canonical_name,)).fetchone()
    if row:
        return row["id"]
    return None


def save_historial_acta(
    tipo_acta,
    datos_json,
    docx_path,
    pdf_path,
    numero_acta=None,
    plantilla_hash=None,
    plantilla_snapshot_path=None,
):
    from database.controller import _ensure_historial_actas_numero_unique_by_type_runtime

    _ensure_historial_actas_numero_unique_by_type_runtime()
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
    from database.controller import _ensure_historial_actas_numero_unique_by_type_runtime

    _ensure_historial_actas_numero_unique_by_type_runtime()
    db = get_db()
    stored_docx_path = _to_storage_relative_path(docx_path)
    stored_pdf_path = _to_storage_relative_path(pdf_path)
    stored_snapshot_path = _to_storage_relative_path(plantilla_snapshot_path)
    cursor = db.execute(
        """
        UPDATE historial_actas
        SET
            tipo_acta = ?,
            numero_acta = ?,
            fecha = CURRENT_TIMESTAMP,
            datos_json = ?,
            docx_path = ?,
            pdf_path = ?,
            plantilla_hash = ?,
            plantilla_snapshot_path = ?
        WHERE id = ?
        """,
        (
            tipo_acta,
            numero_acta,
            datos_json,
            stored_docx_path,
            stored_pdf_path,
            plantilla_hash,
            stored_snapshot_path,
            int(acta_id),
        ),
    )
    db.commit()
    return int(cursor.rowcount or 0)


def get_max_numero_acta_for_year(year, tipo_acta=None):
    from database.controller import _ensure_historial_actas_numero_unique_by_type_runtime

    _ensure_historial_actas_numero_unique_by_type_runtime()
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
    from database.controller import (
        _ensure_actas_sequence_by_type_table_runtime,
        _ensure_historial_actas_numero_unique_by_type_runtime,
    )

    _ensure_actas_sequence_by_type_table_runtime()
    _ensure_historial_actas_numero_unique_by_type_runtime()
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
    from database.controller import (
        _ensure_actas_sequence_by_type_table_runtime,
        _ensure_historial_actas_numero_unique_by_type_runtime,
    )

    _ensure_actas_sequence_by_type_table_runtime()
    _ensure_historial_actas_numero_unique_by_type_runtime()
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
    from database.controller import _ensure_historial_actas_numero_unique_by_type_runtime

    _ensure_historial_actas_numero_unique_by_type_runtime()
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
