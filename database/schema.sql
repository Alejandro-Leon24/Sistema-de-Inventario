CREATE TABLE IF NOT EXISTS bloques (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS pisos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bloque_id INTEGER NOT NULL,
    nombre TEXT NOT NULL,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (bloque_id, nombre),
    FOREIGN KEY (bloque_id) REFERENCES bloques (id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS areas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    piso_id INTEGER NOT NULL,
    nombre TEXT NOT NULL,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (piso_id, nombre),
    FOREIGN KEY (piso_id) REFERENCES pisos (id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS inventario_items (
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
);

CREATE INDEX IF NOT EXISTS idx_inventario_cod_inventario ON inventario_items(cod_inventario);
CREATE INDEX IF NOT EXISTS idx_inventario_cod_esbye ON inventario_items(cod_esbye);

CREATE TABLE IF NOT EXISTS column_mappings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    columna_origen TEXT NOT NULL,
    campo_canonico TEXT NOT NULL,
    orden INTEGER NOT NULL DEFAULT 0,
    UNIQUE (columna_origen)
);

CREATE TABLE IF NOT EXISTS user_preferences (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_key TEXT NOT NULL,
    pref_key TEXT NOT NULL,
    pref_value TEXT,
    actualizado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (user_key, pref_key)
);

CREATE TABLE IF NOT EXISTS inventario_auditoria (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    item_id INTEGER NOT NULL,
    accion TEXT NOT NULL,
    campo TEXT,
    valor_anterior TEXT,
    valor_nuevo TEXT,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (item_id) REFERENCES inventario_items (id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS parametros_universidad (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    valor TEXT NOT NULL,
    tipo TEXT NOT NULL,
    actualizado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS param_estados (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS param_condiciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS param_cuentas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo TEXT UNIQUE,
    nombre TEXT NOT NULL,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS param_si_no (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS param_estado_puerta (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS param_cerraduras (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS param_estado_piso (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS param_material_techo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS param_material_puerta (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS param_estado_pizarra (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    descripcion TEXT,
    orden INTEGER NOT NULL DEFAULT 0,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS administradores (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    cargo TEXT,
    facultad TEXT,
    titulo_academico TEXT,
    email TEXT UNIQUE,
    telefono TEXT,
    activo INTEGER DEFAULT 1,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    actualizado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS historial_actas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tipo_acta TEXT NOT NULL,
    numero_acta TEXT UNIQUE,
    fecha TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    datos_json TEXT,
    docx_path TEXT,
    pdf_path TEXT,
    plantilla_hash TEXT,
    plantilla_snapshot_path TEXT
);

CREATE INDEX IF NOT EXISTS idx_historial_actas_plantilla_snapshot_path ON historial_actas(plantilla_snapshot_path);

CREATE TABLE IF NOT EXISTS secuencia_informes_area (
    anio INTEGER PRIMARY KEY,
    ultimo_numero INTEGER NOT NULL DEFAULT 0,
    actualizado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
