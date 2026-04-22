"""Servicio de dominio para historial y numeracion de actas."""

from pathlib import Path
import re
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
    for field in ("docx_path", "plantilla_snapshot_path"):
        if field in normalized:
            normalized[field] = _to_storage_absolute_path(normalized.get(field))
    return normalized


def _split_numero_acta(numero_acta):
    text = str(numero_acta or "").strip()
    # Soporta formatos como "001-2026", "ACTA-001-2026", "001/2026", etc.
    # Busca el último bloque numérico como año y el anterior como secuencia.
    matches = re.findall(r"(\d+)", text)
    if len(matches) < 2:
        return None, None
    try:
        seq = int(matches[-2])
        year = int(matches[-1])
        return seq, year
    except (ValueError, IndexError):
        return None, None


def _normalize_tipo_acta(tipo_acta):
    text = str(tipo_acta or "").strip().lower()
    if text == "baja":
        return "bajas"
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
    pdf_path=None,  # Mantener por compatibilidad de firma pero no usar
    numero_acta=None,
    plantilla_hash=None,
    plantilla_snapshot_path=None,
):
    from database.controller import _ensure_historial_actas_numero_unique_by_type_runtime

    _ensure_historial_actas_numero_unique_by_type_runtime()
    db = get_db()
    stored_docx_path = _to_storage_relative_path(docx_path)
    stored_snapshot_path = _to_storage_relative_path(plantilla_snapshot_path)
    cursor = db.execute(
        """
        INSERT INTO historial_actas (
            tipo_acta,
            numero_acta,
            datos_json,
            docx_path,
            plantilla_hash,
            plantilla_snapshot_path
        ) VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            tipo_acta,
            numero_acta,
            datos_json,
            stored_docx_path,
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
    pdf_path=None,  # Mantener por compatibilidad de firma pero no usar
    numero_acta=None,
    plantilla_hash=None,
    plantilla_snapshot_path=None,
):
    from database.controller import _ensure_historial_actas_numero_unique_by_type_runtime

    _ensure_historial_actas_numero_unique_by_type_runtime()
    db = get_db()
    stored_docx_path = _to_storage_relative_path(docx_path)
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
            plantilla_hash = ?,
            plantilla_snapshot_path = ?
        WHERE id = ?
        """,
        (
            tipo_acta,
            numero_acta,
            datos_json,
            stored_docx_path,
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
    params = []
    if tipo_acta is not None:
        where_tipo = " AND LOWER(COALESCE(tipo_acta, '')) = LOWER(?)"
        params.append(_normalize_tipo_acta(tipo_acta))

    # Obtenemos todos los números del año para procesarlos en Python (más robusto que SQL puro para formatos variables)
    rows = db.execute(
        f"""
        SELECT numero_acta
        FROM historial_actas
        WHERE numero_acta IS NOT NULL
          AND TRIM(numero_acta) != ''
          {where_tipo}
        """,
        params,
    ).fetchall()
    
    max_val = 0
    for row in rows:
        seq, y = _split_numero_acta(row["numero_acta"])
        if y == target_year and seq is not None:
            if seq > max_val:
                max_val = seq
                
    return max_val


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
    
    # 1. Buscar el máximo en el historial real
    max_historial = get_max_numero_acta_for_year(target_year, tipo_acta=tipo_key)
    
    # 2. Buscar en la tabla de secuencias (reserva)
    seq_row = db.execute(
        "SELECT ultimo_numero FROM secuencia_actas_tipo WHERE tipo_acta = ? AND anio = ?",
        (tipo_key, target_year),
    ).fetchone()
    max_sequence = int(seq_row["ultimo_numero"] or 0) if seq_row else 0
    
    # El siguiente es el máximo de ambos + 1
    next_value = max(max_historial, max_sequence) + 1
    
    # Formatear como 001-2026, 002-2026, etc.
    return f"{next_value:03d}-{target_year}"


def reserve_numero_acta(year, preferred_numero_acta=None, tipo_acta=None, editing_acta_id=None):
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
            
            # Verificar si el numero ya existe en otra acta del mismo tipo
            sql_exists = "SELECT id FROM historial_actas WHERE numero_acta = ? AND LOWER(tipo_acta) = LOWER(?) "
            params_exists = [preferred_numero_acta, tipo_key]
            
            # Si estamos editando, ignoramos el acta actual
            if editing_acta_id:
                try:
                    eid = int(editing_acta_id)
                    sql_exists += " AND id != ? "
                    params_exists.append(eid)
                    from database.controller import logger as db_logger
                    db_logger.info(f"Checking existing acta for number={preferred_numero_acta}, type={tipo_key}, excluding id={eid}")
                except (ValueError, TypeError):
                    pass
            
            existing = db.execute(sql_exists + " LIMIT 1", params_exists).fetchone()
            if existing:
                from database.controller import logger as db_logger
                db_logger.warning(f"Duplicate acta found for number={preferred_numero_acta}, type={tipo_key}. Existing ID: {existing['id']}")
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
        return f"{reserved:03d}-{target_year}" if not preferred_numero_acta else preferred_numero_acta
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
            "SELECT 1 FROM historial_actas WHERE numero_acta = ? AND LOWER(tipo_acta) = LOWER(?) LIMIT 1",
            (numero, _normalize_tipo_acta(tipo_acta)),
        ).fetchone()
    return bool(row)


def get_historial_actas(tipo_acta=None):
    db = get_db()
    # Ordenar por fecha es lo más seguro si el formato de número varía
    order_sql = "ORDER BY fecha DESC, id DESC"
    
    if tipo_acta:
        rows = db.execute(
            f"SELECT * FROM historial_actas WHERE tipo_acta = ? {order_sql}",
            (tipo_acta,),
        ).fetchall()
    else:
        rows = db.execute(f"SELECT * FROM historial_actas {order_sql}").fetchall()
    return [_normalize_historial_row_paths(dict(row)) for row in rows]


def get_historial_acta_by_id(acta_id):
    import json
    db = get_db()
    row = db.execute("SELECT * FROM historial_actas WHERE id = ?", (int(acta_id),)).fetchone()
    if not row:
        return None
    
    acta = _normalize_historial_row_paths(dict(row))
    
    # Validar existencia de bienes en el inventario actual
    try:
        datos = json.loads(acta["datos_json"] or "{}")
        tabla = datos.get("tabla", [])
        if isinstance(tabla, list) and tabla:
            # Obtener todos los IDs de bienes en el inventario actual
            inventory_ids = {
                r["id"] for r in db.execute("SELECT id FROM inventario_items").fetchall()
            }
            # Marcar bienes eliminados
            for item in tabla:
                item_id = item.get("id")
                if item_id and item_id not in inventory_ids:
                    item["eliminado"] = True
            
            acta["datos_json"] = json.dumps(datos, ensure_ascii=False)
    except Exception:
        # Si falla el parseo, devolvemos el acta tal cual
        pass
        
    return acta


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
