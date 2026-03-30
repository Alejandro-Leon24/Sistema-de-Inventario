"""Controlador para parámetros de universidad y configuración"""
import sqlite3
from database.db import get_db


PARAM_TABLES = {
    "estados": "param_estados",
    "condiciones": "param_condiciones",
    "cuentas": "param_cuentas",
    "si_no": "param_si_no",
    "estado_puerta": "param_estado_puerta",
    "cerraduras": "param_cerraduras",
    "estado_piso": "param_estado_piso",
    "material_techo": "param_material_techo",
    "material_puerta": "param_material_puerta",
    "estado_pizarra": "param_estado_pizarra",
}


def _get_param_table(tipo):
    table = PARAM_TABLES.get(tipo)
    if not table:
        raise ValueError("Tipo de parámetro no soportado")
    return table


def get_param(tipo):
    """Obtener lista de parámetros por tipo (estados, condiciones, cuentas)"""
    db = get_db()
    table = _get_param_table(tipo)
    try:
        rows = db.execute(
            f"SELECT id, nombre, descripcion, orden FROM {table} ORDER BY orden, id"
        ).fetchall()
        result = []
        for row in rows:
            item = {
                "id": row["id"],
                "nombre": row["nombre"],
                "descripcion": row["descripcion"],
                "can_delete": can_delete_param(tipo, row["nombre"]),
            }
            result.append(item)
        return result
    except sqlite3.OperationalError:
        return []


def _next_order(table_name):
    db = get_db()
    row = db.execute(f"SELECT COALESCE(MAX(orden), 0) + 1 AS next_order FROM {table_name}").fetchone()
    return row["next_order"]


def create_param(tipo, nombre, descripcion=None):
    """Crear nuevo parámetro"""
    db = get_db()
    table = _get_param_table(tipo)
    orden = _next_order(table)
    cursor = db.execute(
        f"INSERT INTO {table} (nombre, descripcion, orden) VALUES (?, ?, ?)",
        (nombre.strip(), (descripcion or "").strip() or None, orden),
    )
    db.commit()
    return cursor.lastrowid


def update_param(tipo, param_id, nombre, descripcion=None):
    db = get_db()
    table = _get_param_table(tipo)
    
    old_row = db.execute(f"SELECT nombre FROM {table} WHERE id = ?", (param_id,)).fetchone()
    old_nombre = old_row["nombre"] if old_row else None
    new_nombre = nombre.strip()
    
    cursor = db.execute(
        f"UPDATE {table} SET nombre = ?, descripcion = ? WHERE id = ?",
        (new_nombre, (descripcion or "").strip() or None, param_id),
    )
    
    # Propagar los cambios a la tabla de inventario donde el valor dependa directamente de sus strings
    if old_nombre and old_nombre != new_nombre:
        if tipo == "estados":
            db.execute("UPDATE inventario_items SET estado = ? WHERE estado = ?", (new_nombre, old_nombre))
        elif tipo == "condiciones":
            db.execute("UPDATE inventario_items SET condicion = ? WHERE condicion = ?", (new_nombre, old_nombre))
        elif tipo == "cuentas":
            db.execute("UPDATE inventario_items SET cuenta = ? WHERE cuenta = ?", (new_nombre, old_nombre))

    db.commit()
    return cursor.rowcount > 0


def can_delete_param(tipo, nombre):
    db = get_db()
    if tipo == "estados":
        row = db.execute(
            "SELECT COUNT(1) AS total FROM inventario_items WHERE estado = ?",
            (nombre,),
        ).fetchone()
        return (row["total"] or 0) == 0
    if tipo == "condiciones":
        row = db.execute(
            "SELECT COUNT(1) AS total FROM inventario_items WHERE condicion = ?",
            (nombre,),
        ).fetchone()
        return (row["total"] or 0) == 0
    if tipo == "cuentas":
        row = db.execute(
            "SELECT COUNT(1) AS total FROM inventario_items WHERE cuenta = ?",
            (nombre,),
        ).fetchone()
        return (row["total"] or 0) == 0
    if tipo == "estado_puerta":
        row = db.execute(
            "SELECT COUNT(1) AS total FROM areas WHERE estado_puerta = ?",
            (nombre,),
        ).fetchone()
        return (row["total"] or 0) == 0
    if tipo == "cerraduras":
        row = db.execute(
            "SELECT COUNT(1) AS total FROM areas WHERE cerradura = ?",
            (nombre,),
        ).fetchone()
        return (row["total"] or 0) == 0
    if tipo == "estado_piso":
        row = db.execute(
            "SELECT COUNT(1) AS total FROM areas WHERE estado_piso = ?",
            (nombre,),
        ).fetchone()
        return (row["total"] or 0) == 0
    if tipo == "material_techo":
        row = db.execute(
            "SELECT COUNT(1) AS total FROM areas WHERE material_techo = ?",
            (nombre,),
        ).fetchone()
        return (row["total"] or 0) == 0
    if tipo == "material_puerta":
        row = db.execute(
            "SELECT COUNT(1) AS total FROM areas WHERE material_puerta = ?",
            (nombre,),
        ).fetchone()
        return (row["total"] or 0) == 0
    if tipo == "estado_pizarra":
        row = db.execute(
            "SELECT COUNT(1) AS total FROM areas WHERE pizarra_estado = ?",
            (nombre,),
        ).fetchone()
        return (row["total"] or 0) == 0
    if tipo == "si_no":
        si_no_columns = [
            "senaletica", "puerta", "estado_paredes", "estado_techo",
            "nivel_seguridad", "sitio_profesor_mesa", "sitio_profesor_silla",
            "pc_aula", "proyector", "pantalla_interactiva", "pizarra",
            "wifi", "red_lan", "puntos_electricos", "ambiente_apto_retorno"
        ]
        conditions = " OR ".join([f"{col} = ?" for col in si_no_columns])
        params = tuple([nombre] * len(si_no_columns))
        row = db.execute(f"SELECT COUNT(1) AS total FROM areas WHERE {conditions}", params).fetchone()
        return (row["total"] or 0) == 0
    return True


def delete_param(tipo, param_id):
    """Eliminar parámetro"""
    db = get_db()
    table = _get_param_table(tipo)
    row = db.execute(f"SELECT nombre FROM {table} WHERE id = ?", (param_id,)).fetchone()
    if not row:
        return False
    if not can_delete_param(tipo, row["nombre"]):
        raise ValueError("No se puede eliminar porque tiene relaciones activas")
    db.execute(f"DELETE FROM {table} WHERE id = ?", (param_id,))
    db.commit()
    return True


def get_universidad():
    """Obtener parámetros de universidad"""
    db = get_db()
    try:
        rows = db.execute("SELECT nombre, valor FROM parametros_universidad").fetchall()
        return {row["nombre"]: row["valor"] for row in rows}
    except sqlite3.OperationalError:
        return {}


def set_universidad(nombre, valor):
    """Establecer parámetro de universidad"""
    db = get_db()
    db.execute(
        """
        INSERT INTO parametros_universidad (nombre, valor, tipo, actualizado_en)
        VALUES (?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(nombre)
        DO UPDATE SET valor = excluded.valor, actualizado_en = CURRENT_TIMESTAMP
        """,
        (nombre.strip(), valor.strip(), "text"),
    )
    db.commit()


def get_administradores():
    """Obtener lista de administradores activos"""
    db = get_db()
    try:
        rows = db.execute(
            "SELECT id, nombre, cargo, facultad, titulo_academico, email, telefono FROM administradores WHERE activo = 1 ORDER BY id"
        ).fetchall()
        return [
            {
                "id": row["id"],
                "nombre": row["nombre"],
                "cargo": row["cargo"],
                "facultad": row["facultad"],
                "titulo_academico": row["titulo_academico"],
                "email": row["email"],
                "telefono": row["telefono"],
            }
            for row in rows
        ]
    except sqlite3.OperationalError:
        return []


def create_administrador(payload):
    """Crear nuevo administrador"""
    def clean_optional(value):
        text = str(value or "").strip()
        return text or None

    db = get_db()
    cursor = db.execute(
        """
        INSERT INTO administradores (nombre, cargo, facultad, titulo_academico, email, telefono)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            str(payload.get("nombre") or "").strip(),
            clean_optional(payload.get("cargo")),
            clean_optional(payload.get("facultad")),
            clean_optional(payload.get("titulo_academico")),
            clean_optional(payload.get("email")),
            clean_optional(payload.get("telefono")),
        ),
    )
    db.commit()
    return cursor.lastrowid


def update_administrador(admin_id, payload):
    """Actualizar administrador"""
    def clean_optional(value):
        text = str(value or "").strip()
        return text or None

    db = get_db()
    
    # Obtener el nombre anterior para actualizar las dependencias en cascada
    old_admin = db.execute("SELECT nombre FROM administradores WHERE id = ?", (admin_id,)).fetchone()
    old_nombre = old_admin['nombre'] if old_admin else None
    
    new_nombre = str(payload.get("nombre") or "").strip()
    
    db.execute(
        """
        UPDATE administradores
        SET nombre = ?, cargo = ?, facultad = ?, titulo_academico = ?, email = ?, telefono = ?, actualizado_en = CURRENT_TIMESTAMP
        WHERE id = ?
        """,
        (
            new_nombre,
            clean_optional(payload.get("cargo")),
            clean_optional(payload.get("facultad")),
            clean_optional(payload.get("titulo_academico")),
            clean_optional(payload.get("email")),
            clean_optional(payload.get("telefono")),
            admin_id,
        ),
    )
    
    # Actualizar la columna usuario_final de inventario_items cuando se modifica el nombre de un personal
    if old_nombre and old_nombre != new_nombre:
        db.execute(
            """
            UPDATE inventario_items
            SET usuario_final = ?
            WHERE usuario_final = ?
            """,
            (new_nombre, old_nombre)
        )

    db.commit()


def delete_administrador(admin_id):
    """Desactivar administrador (soft delete) y quitar reasignaciones"""
    db = get_db()
    row = db.execute("SELECT nombre FROM administradores WHERE id = ?", (admin_id,)).fetchone()
    old_nombre = row["nombre"] if row else None

    # Soft delete
    db.execute("UPDATE administradores SET activo = 0 WHERE id = ?", (admin_id,))
    
    # "Cascade" desvincular el personal de los bienes
    if old_nombre:
        db.execute("UPDATE inventario_items SET usuario_final = NULL WHERE usuario_final = ?", (old_nombre,))
        
    # Desvincular de las areas relacionadas
    db.execute("UPDATE areas SET responsable_admin_id = NULL WHERE responsable_admin_id = ?", (admin_id,))

    db.commit()

def get_administrador_dependency_summary(admin_id):
    db = get_db()
    row = db.execute("SELECT nombre FROM administradores WHERE id = ?", (admin_id,)).fetchone()
    if not row:
        return {"items": 0, "areas": 0}

    nombre = row["nombre"]
    count_items = db.execute("SELECT COUNT(1) as total FROM inventario_items WHERE usuario_final = ?", (nombre,)).fetchone()
    count_areas = db.execute("SELECT COUNT(1) as total FROM areas WHERE responsable_admin_id = ?", (admin_id,)).fetchone()
    
    return {
        "items": count_items["total"] or 0,
        "areas": count_areas["total"] or 0
    }
