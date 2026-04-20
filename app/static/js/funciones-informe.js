const api = window.api;

const DEFAULT_ENTREGA_COLUMN_IDS = ["cod_inventario", "descripcion", "marca", "modelo", "serie", "estado"];
const DEFAULT_RECEPCION_COLUMN_IDS = ["cod_inventario", "descripcion", "marca", "modelo", "cantidad", "estado"];
const DEFAULT_BAJAS_COLUMN_IDS = ["item_numero", "cod_inventario", "cod_esbye", "descripcion", "estado", "justificacion"];
const BAJAS_REQUIRED_COLUMN_IDS = ["estado", "justificacion"];
const INFORME_PREFS_ENTREGA_COLUMNS_KEY = "informe_entrega_columnas_acta";
const INFORME_PREFS_RECEPCION_COLUMNS_KEY = "informe_recepcion_columnas_acta";
const INFORME_PREFS_BAJAS_COLUMNS_KEY = "informe_bajas_columnas_acta";
const INFORME_LEGACY_ENTREGA_COLUMNS_KEY = "informe.prefs.entrega.columns.v1";
const INFORME_LEGACY_RECEPCION_COLUMNS_KEY = "informe.prefs.recepcion.columns.v1";

let selectedColumns = [...DEFAULT_ENTREGA_COLUMN_IDS];
const availableColumns = [
    { id: "cod_inventario", label: "COD INV." },
    { id: "cod_esbye", label: "COD. ESBYE" },
    { id: "cuenta", label: "CUENTA" },
    { id: "cantidad", label: "CANT" },
    { id: "descripcion", label: "DESCRIPCION" },
    { id: "marca", label: "MARCA" },
    { id: "modelo", label: "MODELO" },
    { id: "serie", label: "SERIE" },
    { id: "estado", label: "ESTADO" },
    { id: "ubicacion", label: "UBICACION" },
    { id: "fecha_adquisicion", label: "FECHA ADQUISICION" },
    { id: "valor", label: "VALOR" },
    { id: "usuario_final", label: "USUARIO FINAL" },
    { id: "observacion", label: "OBSERVACION" },
    { id: "descripcion_esbye", label: "DESCRIPCION ESBYE" },
    { id: "marca_esbye", label: "MARCA ESBYE" },
    { id: "modelo_esbye", label: "MODELO ESBYE" },
    { id: "serie_esbye", label: "SERIE ESBYE" },
    { id: "fecha_adquisicion_esbye", label: "FECHA ESBYE" },
    { id: "valor_esbye", label: "VALOR ESBYE" },
    { id: "ubicacion_esbye", label: "UBICACION ESBYE" },
    { id: "observacion_esbye", label: "OBSERVACION ESBYE" },
];

let structureData = [];
let inventoryDataCache = [];
let selectedItemIds = new Set();
let extraccionFilteredItems = [];
let extraccionPage = 1;
let extraccionPerPage = 25;
let recepcionTablePage = 1;
let recepcionTablePerPage = 10;
window._globalSelectedTableRows = [];
window._globalSelectedColumns = [];

let downloadProgressTimer = null;
let downloadProgressValue = 0;
let isGeneratingActa = false;
let closeDownloadModalTimer = null;
let informeEventSource = null;
let activeHistorialTemplateSnapshotPath = null;
let activeEditingActaId = null;
let activeEditingActaTipo = null;
let activeAulaBatchJobId = null;
let activeAulaBatchModal = null;
let activeAulaBatchPaused = false;
let recepcionBienesTemp = [];
let recepcionEditIndex = -1;
let recepcionSelectedColumnIds = [...DEFAULT_RECEPCION_COLUMN_IDS];
let bajasBienesTemp = [];
let bajasSelectedItemIds = new Set();
let bajasDraftBienes = [];
let bajasFilteredItems = [];
let bajasPage = 1;
let bajasPerPage = 25;
let bajasStep = 1;
let bajasEstadoOptions = [];
let bajasSelectedColumnIds = [...DEFAULT_BAJAS_COLUMN_IDS];

const RECEPCION_BIENES_COLUMNS = [
    { id: "cod_inventario", label: "COD INV." },
    { id: "cod_esbye", label: "COD. ESBYE" },
    { id: "cuenta", label: "CUENTA" },
    { id: "cantidad", label: "CANT" },
    { id: "descripcion", label: "DESCRIPCION" },
    { id: "marca", label: "MARCA" },
    { id: "modelo", label: "MODELO" },
    { id: "serie", label: "SERIE" },
    { id: "estado", label: "ESTADO" },
    { id: "ubicacion", label: "UBICACION" },
    { id: "fecha_adquisicion", label: "FECHA ADQUISICION" },
    { id: "valor", label: "VALOR" },
    { id: "usuario_final", label: "USUARIO FINAL" },
    { id: "observacion", label: "OBSERVACION" },
    { id: "descripcion_esbye", label: "DESCRIPCION ESBYE" },
    { id: "marca_esbye", label: "MARCA ESBYE" },
    { id: "modelo_esbye", label: "MODELO ESBYE" },
    { id: "serie_esbye", label: "SERIE ESBYE" },
    { id: "fecha_adquisicion_esbye", label: "FECHA ESBYE" },
    { id: "valor_esbye", label: "VALOR ESBYE" },
    { id: "ubicacion_esbye", label: "UBICACION ESBYE" },
    { id: "observacion_esbye", label: "OBSERVACION ESBYE" },
];

const BAJAS_BIENES_COLUMNS = [
    { id: "item_numero", label: "ITEM" },
    { id: "cod_inventario", label: "COD INV." },
    { id: "cod_esbye", label: "COD. ESBYE" },
    { id: "cuenta", label: "CUENTA" },
    { id: "cantidad", label: "CANT" },
    { id: "descripcion", label: "DESCRIPCION" },
    { id: "marca", label: "MARCA" },
    { id: "modelo", label: "MODELO" },
    { id: "serie", label: "SERIE" },
    { id: "estado", label: "ESTADO" },
    { id: "ubicacion", label: "UBICACION" },
    { id: "fecha_adquisicion", label: "FECHA ADQUISICION" },
    { id: "valor", label: "VALOR" },
    { id: "usuario_final", label: "USUARIO FINAL" },
    { id: "observacion", label: "OBSERVACION" },
    { id: "justificacion", label: "JUSTIFICACION" },
    { id: "procedencia", label: "PROCEDENCIA" },
    { id: "descripcion_esbye", label: "DESCRIPCION ESBYE" },
    { id: "marca_esbye", label: "MARCA ESBYE" },
    { id: "modelo_esbye", label: "MODELO ESBYE" },
    { id: "serie_esbye", label: "SERIE ESBYE" },
    { id: "fecha_adquisicion_esbye", label: "FECHA ESBYE" },
    { id: "valor_esbye", label: "VALOR ESBYE" },
    { id: "ubicacion_esbye", label: "UBICACION ESBYE" },
    { id: "observacion_esbye", label: "OBSERVACION ESBYE" },
];

const BAJAS_REGISTRADOS_COLUMNS = [
    { id: "item_numero", label: "ITEM" },
    { id: "cod_inventario", label: "COD INV." },
    { id: "cod_esbye", label: "COD. ESBYE" },
    { id: "cuenta", label: "CUENTA" },
    { id: "cantidad", label: "CANT" },
    { id: "descripcion", label: "DESCRIPCION" },
    { id: "marca", label: "MARCA" },
    { id: "modelo", label: "MODELO" },
    { id: "serie", label: "SERIE" },
    { id: "estado", label: "ESTADO" },
    { id: "ubicacion", label: "UBICACION" },
    { id: "fecha_adquisicion", label: "FECHA ADQUISICION" },
    { id: "valor", label: "VALOR" },
    { id: "usuario_final", label: "USUARIO FINAL" },
    { id: "observacion", label: "OBSERVACION" },
    { id: "justificacion", label: "JUSTIFICACION" },
    { id: "procedencia", label: "PROCEDENCIA" },
    { id: "descripcion_esbye", label: "DESCRIPCION ESBYE" },
    { id: "marca_esbye", label: "MARCA ESBYE" },
    { id: "modelo_esbye", label: "MODELO ESBYE" },
    { id: "serie_esbye", label: "SERIE ESBYE" },
    { id: "fecha_adquisicion_esbye", label: "FECHA ESBYE" },
    { id: "valor_esbye", label: "VALOR ESBYE" },
    { id: "ubicacion_esbye", label: "UBICACION ESBYE" },
    { id: "observacion_esbye", label: "OBSERVACION ESBYE" },
];

function normalizeRecepcionColumnIds(rawColumns) {
    const allIds = new Set(RECEPCION_BIENES_COLUMNS.map((c) => c.id));
    const ids = (rawColumns || [])
        .map((entry) => (typeof entry === "string" ? entry : entry?.id))
        .map((id) => String(id || "").trim())
        .filter((id) => allIds.has(id));

    const unique = [];
    ids.forEach((id) => {
        if (!unique.includes(id)) unique.push(id);
    });

    return unique.length ? unique : [...DEFAULT_RECEPCION_COLUMN_IDS];
}

function normalizeEntregaColumnIds(rawColumns) {
    const allIds = new Set(availableColumns.map((c) => c.id));
    const ids = (rawColumns || [])
        .map((entry) => (typeof entry === "string" ? entry : entry?.id))
        .map((id) => String(id || "").trim())
        .filter((id) => allIds.has(id));

    const unique = [];
    ids.forEach((id) => {
        if (!unique.includes(id)) unique.push(id);
    });

    return unique.length ? unique : [...DEFAULT_ENTREGA_COLUMN_IDS];
}

function normalizeBajasColumnIds(rawColumns) {
    const allIds = new Set(BAJAS_BIENES_COLUMNS.map((c) => c.id));
    const ids = (rawColumns || [])
        .map((entry) => (typeof entry === "string" ? entry : entry?.id))
        .map((id) => String(id || "").trim())
        .filter((id) => allIds.has(id));

    const unique = [];
    ids.forEach((id) => {
        if (!unique.includes(id)) unique.push(id);
    });

    BAJAS_REQUIRED_COLUMN_IDS.forEach((id) => {
        if (!unique.includes(id)) unique.push(id);
    });

    return unique.length ? unique : [...DEFAULT_BAJAS_COLUMN_IDS];
}

function getLegacyColumnsFromLocalStorage(storageKey, normalizer) {
    try {
        const raw = localStorage.getItem(storageKey);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        const normalized = normalizer(parsed);
        return Array.isArray(normalized) && normalized.length ? normalized : null;
    } catch (_err) {
        return null;
    }
}

function clearLegacyColumnPreferenceKeys() {
    try {
        localStorage.removeItem(INFORME_LEGACY_ENTREGA_COLUMNS_KEY);
        localStorage.removeItem(INFORME_LEGACY_RECEPCION_COLUMNS_KEY);
    } catch (_err) {
        // noop
    }
}

async function loadColumnPreferences() {
    try {
        const preferences = await window.appHelpers.loadPreferences(api);
        const hasEntregaInDb = Object.prototype.hasOwnProperty.call(preferences || {}, INFORME_PREFS_ENTREGA_COLUMNS_KEY);
        const hasRecepcionInDb = Object.prototype.hasOwnProperty.call(preferences || {}, INFORME_PREFS_RECEPCION_COLUMNS_KEY);
        const hasBajasInDb = Object.prototype.hasOwnProperty.call(preferences || {}, INFORME_PREFS_BAJAS_COLUMNS_KEY);

        let entregaColumns = hasEntregaInDb
            ? normalizeEntregaColumnIds(preferences?.[INFORME_PREFS_ENTREGA_COLUMNS_KEY])
            : null;
        let recepcionColumns = hasRecepcionInDb
            ? normalizeRecepcionColumnIds(preferences?.[INFORME_PREFS_RECEPCION_COLUMNS_KEY])
            : null;
        let bajasColumns = hasBajasInDb
            ? normalizeBajasColumnIds(preferences?.[INFORME_PREFS_BAJAS_COLUMNS_KEY])
            : null;

        if (!hasEntregaInDb) {
            const legacyEntrega = getLegacyColumnsFromLocalStorage(
                INFORME_LEGACY_ENTREGA_COLUMNS_KEY,
                normalizeEntregaColumnIds
            );
            if (legacyEntrega) {
                entregaColumns = legacyEntrega;
                await api.send("/api/preferencias", "PATCH", {
                    pref_key: INFORME_PREFS_ENTREGA_COLUMNS_KEY,
                    pref_value: legacyEntrega,
                });
            }
        }

        if (!hasRecepcionInDb) {
            const legacyRecepcion = getLegacyColumnsFromLocalStorage(
                INFORME_LEGACY_RECEPCION_COLUMNS_KEY,
                normalizeRecepcionColumnIds
            );
            if (legacyRecepcion) {
                recepcionColumns = legacyRecepcion;
                await api.send("/api/preferencias", "PATCH", {
                    pref_key: INFORME_PREFS_RECEPCION_COLUMNS_KEY,
                    pref_value: legacyRecepcion,
                });
            }
        }

        selectedColumns = entregaColumns || [...DEFAULT_ENTREGA_COLUMN_IDS];
        recepcionSelectedColumnIds = recepcionColumns || [...DEFAULT_RECEPCION_COLUMN_IDS];
        bajasSelectedColumnIds = bajasColumns || [...DEFAULT_BAJAS_COLUMN_IDS];
        clearLegacyColumnPreferenceKeys();
    } catch (_err) {
        selectedColumns = [...DEFAULT_ENTREGA_COLUMN_IDS];
        recepcionSelectedColumnIds = [...DEFAULT_RECEPCION_COLUMN_IDS];
        bajasSelectedColumnIds = [...DEFAULT_BAJAS_COLUMN_IDS];
    }
}

async function saveEntregaColumnPreferences() {
    try {
        await api.send("/api/preferencias", "PATCH", {
            pref_key: INFORME_PREFS_ENTREGA_COLUMNS_KEY,
            pref_value: normalizeEntregaColumnIds(selectedColumns),
        });
    } catch (_err) {
        // noop
    }
}

async function saveRecepcionColumnPreferences() {
    try {
        await api.send("/api/preferencias", "PATCH", {
            pref_key: INFORME_PREFS_RECEPCION_COLUMNS_KEY,
            pref_value: normalizeRecepcionColumnIds(recepcionSelectedColumnIds),
        });
    } catch (_err) {
        // noop
    }
}

async function saveBajasColumnPreferences() {
    try {
        await api.send("/api/preferencias", "PATCH", {
            pref_key: INFORME_PREFS_BAJAS_COLUMNS_KEY,
            pref_value: normalizeBajasColumnIds(bajasSelectedColumnIds),
        });
    } catch (_err) {
        // noop
    }
}

function validarNumerosComasInput(input) {
    if (!input) return;
    let value = String(input.value || "").replace(/[^0-9,]/g, "");
    const firstComma = value.indexOf(",");
    if (firstComma !== -1) {
        value = value.slice(0, firstComma + 1) + value.slice(firstComma + 1).replace(/,/g, "");
    }
    input.value = value;
}

const ACTA_TEMPLATE_REQUIRED_VARS = {
    entrega: [
        "numero_acta",
        "fecha_corte",
        "fecha_emision",
        "accion_personal",
        "entregado_por",
        "rol_entrega",
        "recibido_por",
        "rol_recibe",
        "area_trabajo",
        "tabla_dinamica",
    ],
    recepcion: [
        "numero_acta",
        "entregado_por",
        "rol_entrega",
        "recibido_por",
        "rol_recibe",
        "fecha_corte",
        "fecha_elaboracion",
        "accion_personal",
        "memorandum",
        "fecha_memorandum",
        "entregado_por_segunda_delegada",
        "rol_entrega_segunda_delegada",
        "area_trabajo",
        "tabla_dinamica",
    ],
    bajas: [
        "numero_acta",
        "nombre_delegado",
        "recibido_por",
        "entregado_por",
        "rol_entrega",
        "fecha_emision",
        "tabla_dinamica",
    ],
    aula: ["tabla_dinamica"],
};

const ACTA_TEMPLATE_REQUIRED_ALIASES = {
    entrega: {
        rol_recibe: ["usuario_final", "recibe_custodio"],
        area_trabajo: ["ubicacion", "ubicacion_entrega"],
    },
};

function normalizeImportDateLikeInventario(value) {
    const raw = String(value || "").trim();
    if (!raw) return "";
    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;

    if (/^\d+(?:\.\d+)?$/.test(raw)) {
        const serial = Number(raw);
        if (Number.isFinite(serial)) {
            const excelEpoch = new Date(Date.UTC(1899, 11, 30));
            excelEpoch.setUTCDate(excelEpoch.getUTCDate() + Math.floor(serial));
            const year = excelEpoch.getUTCFullYear();
            const month = String(excelEpoch.getUTCMonth() + 1).padStart(2, "0");
            const day = String(excelEpoch.getUTCDate()).padStart(2, "0");
            return `${year}-${month}-${day}`;
        }
    }

    const patterns = [
        /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/,
        /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/,
    ];

    const p1 = raw.match(patterns[0]);
    if (p1) {
        const day = String(Number(p1[1])).padStart(2, "0");
        const month = String(Number(p1[2])).padStart(2, "0");
        const year = p1[3];
        return `${year}-${month}-${day}`;
    }

    const p2 = raw.match(patterns[1]);
    if (p2) {
        const year = p2[1];
        const month = String(Number(p2[2])).padStart(2, "0");
        const day = String(Number(p2[3])).padStart(2, "0");
        return `${year}-${month}-${day}`;
    }

    const parsed = new Date(raw);
    if (!Number.isNaN(parsed.getTime())) {
        const year = parsed.getFullYear();
        const month = String(parsed.getMonth() + 1).padStart(2, "0");
        const day = String(parsed.getDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
    }

    return "";
}

function normalizeImportMoneyLikeInventario(value) {
    const raw = String(value ?? "").trim();
    if (!raw) return "";
    let text = raw.replace(/\s+/g, "");
    if (text.includes(",") && text.includes(".")) {
        text = text.replace(/\./g, "").replace(/,/g, ".");
    } else if (text.includes(",")) {
        text = text.replace(/,/g, ".");
    }
    const num = Number(text);
    if (!Number.isFinite(num)) return "";
    return num.toFixed(2);
}

function normalizeCodeToPlaceholder(value) {
    const text = String(value || "").trim();
    const compact = text.toLowerCase().replace(/[^a-z0-9]/g, "");
    if (!compact || compact === "sc" || compact === "sincodigo" || compact === "sincod") {
        return "S/C";
    }
    return text;
}

function isNoCodeValue(value) {
    return normalizeCodeToPlaceholder(value) === "S/C";
}

const RECEPCION_MODAL_PASTE_FIELDS = [
    "cod_inventario",
    "cod_esbye",
    "cuenta",
    "cantidad",
    "descripcion",
    "marca",
    "modelo",
    "serie",
    "estado",
    "ubicacion",
    "fecha_adquisicion",
    "valor",
    "usuario_final",
    "observacion",
    "descripcion_esbye",
    "marca_esbye",
    "modelo_esbye",
    "serie_esbye",
    "fecha_adquisicion_esbye",
    "valor_esbye",
    "ubicacion_esbye",
    "observacion_esbye",
];

const RECEPCION_MODAL_FIELD_TO_INPUT_ID = {
    cod_inventario: "recepcion-bien-cod-inventario",
    cod_esbye: "recepcion-bien-cod-esbye",
    cuenta: "recepcion-bien-cuenta",
    cantidad: "recepcion-bien-cantidad",
    descripcion: "recepcion-bien-descripcion",
    marca: "recepcion-bien-marca",
    modelo: "recepcion-bien-modelo",
    serie: "recepcion-bien-serie",
    estado: "recepcion-bien-estado",
    fecha_adquisicion: "recepcion-bien-fecha-adquisicion",
    valor: "recepcion-bien-valor",
    usuario_final: "recepcion-bien-usuario-final",
    observacion: "recepcion-bien-observacion",
    descripcion_esbye: "recepcion-bien-descripcion-esbye",
    marca_esbye: "recepcion-bien-marca-esbye",
    modelo_esbye: "recepcion-bien-modelo-esbye",
    serie_esbye: "recepcion-bien-serie-esbye",
    fecha_adquisicion_esbye: "recepcion-bien-fecha-esbye",
    valor_esbye: "recepcion-bien-valor-esbye",
    ubicacion_esbye: "recepcion-bien-ubicacion-esbye",
    observacion_esbye: "recepcion-bien-observacion-esbye",
};

function parseExcelText(text) {
    const normalizedText = String(text || "")
        .replace(/\r\n/g, "\n")
        .replace(/\r/g, "\n");

    return normalizedText
        .split("\n")
        .filter((line) => line.replace(/\t/g, "").trim().length > 0)
        .map((line) => line.split("\t").map((cell) => String(cell || "").trim()));
}

function scoreRecepcionPastedMapping(mapped = {}) {
    let score = 0;
    if (String(mapped.cod_inventario || "").trim()) score += 3;
    if (String(mapped.descripcion || "").trim()) score += 4;
    if (String(mapped.ubicacion || "").trim()) score += 3;
    if (String(mapped.usuario_final || "").trim()) score += 2;
    if (String(mapped.estado || "").trim()) score += 1;
    if (String(mapped.cantidad || "").trim().match(/^\d+$/)) score += 2;
    if (String(mapped.fecha_adquisicion || "").trim()) score += 1;
    if (String(mapped.valor || "").trim()) score += 1;
    score += Object.keys(mapped).length * 0.25;
    return score;
}

function mapPastedRowBestEffortForRecepcion(rawRow) {
    const cells = Array.isArray(rawRow) ? rawRow : [];
    if (!cells.length) return {};

    const candidateOrders = [
        RECEPCION_MODAL_PASTE_FIELDS,
        ["item_numero", ...RECEPCION_MODAL_PASTE_FIELDS],
    ];

    let bestMapped = {};
    let bestScore = Number.NEGATIVE_INFINITY;
    let bestOffset = 0;

    candidateOrders.forEach((order) => {
        const maxOffset = Math.min(8, Math.max(cells.length - 1, 0));
        for (let offset = 0; offset <= maxOffset; offset += 1) {
            const mapped = {};
            for (let idx = 0; idx < order.length; idx += 1) {
                const field = order[idx];
                if (field === "item_numero") continue;
                const srcIdx = idx + offset;
                if (srcIdx >= cells.length) break;
                const value = String(cells[srcIdx] ?? "").trim();
                if (!value) continue;
                mapped[field] = value;
            }

            const score = scoreRecepcionPastedMapping(mapped);
            if (score > bestScore || (score === bestScore && offset < bestOffset)) {
                bestScore = score;
                bestMapped = mapped;
                bestOffset = offset;
            }
        }
    });

    return bestMapped;
}

function normalizeSelectSearchText(value) {
    return String(value || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .trim();
}

function normalizePersonForSelectMatch(value) {
    const normalized = normalizeSelectSearchText(value).replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
    if (!normalized) return "";
    const prefixes = new Set([
        "ing", "ingeniero", "ingeniera", "dr", "dra", "doctor", "doctora",
        "lic", "licenciado", "licenciada", "abg", "abogada", "abogado",
        "arq", "arquitecto", "arquitecta", "tec", "tecnico", "tecnica",
        "sr", "sra", "srta", "msc", "mg", "mgs", "mgtr", "mtr", "mts", "mtro", "mt",
        "prof", "profa", "tlgo", "tlga", "ts", "phd",
    ]);
    const tokens = normalized.split(" ").filter(Boolean);
    while (tokens.length && prefixes.has(tokens[0])) {
        tokens.shift();
    }
    return tokens.join(" ");
}

function resolveSelectOptionBestMatch(options, rawValue, field = "") {
    const normalizedVal = normalizeSelectSearchText(rawValue);
    if (!normalizedVal) return null;
    const list = Array.isArray(options) ? options : [];
    const personValue = field === "usuario_final" ? normalizePersonForSelectMatch(rawValue) : "";

    let matched = list.find((opt) => normalizeSelectSearchText(opt.value) === normalizedVal);
    if (matched) return matched;

    matched = list.find((opt) => normalizeSelectSearchText(opt.textContent) === normalizedVal);
    if (matched) return matched;

    if (personValue) {
        matched = list.find((opt) => normalizePersonForSelectMatch(opt.textContent) === personValue);
        if (matched) return matched;
    }

    matched = list.find((opt) => normalizeSelectSearchText(opt.textContent).includes(normalizedVal));
    if (matched) return matched;

    matched = list.find((opt) => normalizedVal.includes(normalizeSelectSearchText(opt.value)));
    if (matched) return matched;

    const valTokens = (personValue || normalizedVal).split(/\s+/).filter(Boolean);
    let best = null;
    let bestScore = 0;
    list.forEach((opt) => {
        const optNorm = field === "usuario_final"
            ? normalizePersonForSelectMatch(opt.textContent)
            : normalizeSelectSearchText(opt.textContent).replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
        if (!optNorm) return;
        const optTokens = optNorm.split(/\s+/).filter(Boolean);
        const overlap = valTokens.filter((token) => optTokens.some((item) => item.includes(token) || token.includes(item))).length;
        const score = overlap / Math.max(valTokens.length || 1, optTokens.length || 1);
        if (score > bestScore) {
            bestScore = score;
            best = opt;
        }
    });
    return bestScore >= 0.5 ? best : null;
}

function normalizeRecepcionRowLikeInventarioImport(row, forcedLocation) {
    const src = row && typeof row === "object" ? row : {};
    const toText = (v) => String(v ?? "").trim();
    const cantidadNum = parseInt(String(src.cantidad ?? "").trim(), 10);
    return {
        ...src,
        cod_inventario: normalizeCodeToPlaceholder(src.cod_inventario),
        cod_esbye: normalizeCodeToPlaceholder(src.cod_esbye),
        cuenta: toText(src.cuenta),
        descripcion: toText(src.descripcion),
        marca: toText(src.marca),
        modelo: toText(src.modelo),
        serie: toText(src.serie),
        estado: toText(src.estado),
        usuario_final: toText(src.usuario_final),
        observacion: toText(src.observacion),
        descripcion_esbye: toText(src.descripcion_esbye),
        marca_esbye: toText(src.marca_esbye),
        modelo_esbye: toText(src.modelo_esbye),
        serie_esbye: toText(src.serie_esbye),
        observacion_esbye: toText(src.observacion_esbye),
        cantidad: Number.isFinite(cantidadNum) && cantidadNum > 0 ? cantidadNum : 1,
        valor: normalizeImportMoneyLikeInventario(src.valor),
        valor_esbye: normalizeImportMoneyLikeInventario(src.valor_esbye),
        fecha_adquisicion: normalizeImportDateLikeInventario(src.fecha_adquisicion),
        fecha_adquisicion_esbye: normalizeImportDateLikeInventario(src.fecha_adquisicion_esbye || src.fecha_esbye),
        ubicacion: toText(forcedLocation),
        ubicacion_esbye: toText(src.ubicacion_esbye),
    };
}

function normalizeTemplateVarName(value) {
    let v = String(value || "").trim();
    if (!v) return "";
    v = v.replace(/^\{\{\s*/, "").replace(/\s*\}\}$/, "");
    v = v.split("|")[0].trim();
    v = v.replace(/([a-z0-9])([A-Z])/g, "$1_$2");
    v = v.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    v = v.toLowerCase().replace(/[\s-]+/g, "_");
    v = v.replace(/_+/g, "_").replace(/^_+|_+$/g, "");
    return v;
}

function templateVarSet(variables) {
    return new Set((variables || []).map((v) => normalizeTemplateVarName(v)).filter(Boolean));
}

function isActaRequiredVarPresent(tipo, requiredKey, foundSet) {
    const normalized = normalizeTemplateVarName(requiredKey);
    if (!normalized) return true;
    if (foundSet.has(normalized)) return true;

    const aliases = (ACTA_TEMPLATE_REQUIRED_ALIASES[tipo] || {})[normalized] || [];
    return aliases.some((alias) => foundSet.has(normalizeTemplateVarName(alias)));
}

function resolveActaGuardDestination() {
    const fallback = "/ajustes";
    const rawRef = String(document.referrer || "").trim();
    if (!rawRef) return fallback;
    try {
        const refUrl = new URL(rawRef, window.location.origin);
        if (refUrl.origin !== window.location.origin) return fallback;
        const current = `${window.location.pathname}${window.location.search}`;
        const refPath = `${refUrl.pathname}${refUrl.search}`;
        if (!refPath || refPath === current || refUrl.pathname === window.location.pathname) return fallback;
        return `${refUrl.pathname}${refUrl.search}${refUrl.hash}`;
    } catch (_err) {
        return fallback;
    }
}

function showActaGuardModal(issues) {
    const modalEl = document.getElementById("modalBloqueoActasPlantilla");
    const detailEl = document.getElementById("bloqueo-actas-detalle");
    const okBtn = document.getElementById("btn-bloqueo-actas-ok");
    if (!modalEl || !detailEl || !okBtn) {
        window.location.href = resolveActaGuardDestination();
        return;
    }

    detailEl.innerHTML = "";
    const list = document.createElement("ul");
    list.className = "mb-0 ps-3";
    issues.forEach((issue) => {
        const li = document.createElement("li");
        li.innerHTML = `<strong>${issue.tipoLabel}:</strong> ${issue.reason}`;
        list.appendChild(li);
    });
    detailEl.appendChild(list);

    if (okBtn.dataset.bound !== "1") {
        okBtn.dataset.bound = "1";
        okBtn.addEventListener("click", () => {
            window.location.href = resolveActaGuardDestination();
        });
    }

    const modal = new bootstrap.Modal(modalEl, { backdrop: "static", keyboard: false });
    modal.show();
}

async function checkActaTemplatesAccessGuard() {
    const tipos = ["entrega", "recepcion", "bajas", "aula"];
    const labels = {
        entrega: "Acta de Entrega",
        recepcion: "Acta de Recepción",
        bajas: "Acta de Bienes de Baja",
        aula: "Inventario por Área",
    };
    const issues = [];

    for (const tipo of tipos) {
        try {
            const res = await fetch(`/api/plantillas/estado?tipo=${encodeURIComponent(tipo)}`);
            const payload = await res.json();
            if (!res.ok || !payload.success || !payload.existe) {
                issues.push({ tipoLabel: labels[tipo] || tipo, reason: "No existe plantilla cargada." });
                continue;
            }

            const found = templateVarSet(payload.variables || []);
            const required = ACTA_TEMPLATE_REQUIRED_VARS[tipo] || [];
            const missing = required.filter((k) => !isActaRequiredVarPresent(tipo, k, found));
            if (missing.length) {
                const rendered = missing.map((k) => `{{${k}}}`).join(", ");
                issues.push({
                    tipoLabel: labels[tipo] || tipo,
                    reason: `Faltan variables obligatorias: ${rendered}.`,
                });
            }
        } catch (_err) {
            issues.push({ tipoLabel: labels[tipo] || tipo, reason: "No se pudo verificar la plantilla." });
        }
    }

    if (!issues.length) return true;
    showActaGuardModal(issues);
    return false;
}

function _updateUbicacionFromSelects(prefix) {
    const bloqueSel = document.getElementById(`${prefix}-bloque`);
    const pisoSel = document.getElementById(`${prefix}-piso`);
    const areaSel = document.getElementById(`${prefix}-area`);
    const areaTrabajo = document.getElementById(`${prefix}-area-trabajo`);
    const areaIdHidden = document.getElementById(`${prefix}-ubicacion-area-id`);
    if (!bloqueSel || !pisoSel || !areaSel || !areaTrabajo || !areaIdHidden) return;

    const parts = [];
    if (String(bloqueSel.value || "").trim()) parts.push(bloqueSel.options[bloqueSel.selectedIndex]?.textContent || "");
    if (String(pisoSel.value || "").trim()) parts.push(pisoSel.options[pisoSel.selectedIndex]?.textContent || "");
    if (String(areaSel.value || "").trim()) parts.push(areaSel.options[areaSel.selectedIndex]?.textContent || "");

    areaTrabajo.value = parts.map((p) => String(p || "").trim()).filter(Boolean).join(", ");
    areaIdHidden.value = String(areaSel.value || "").trim();
}

function shouldRefreshNumeroActaInput(input, force = false) {
    if (!input) return false;
    if (force) return true;

    const current = String(input.value || "").trim();
    if (!current) return true;

    const lastAuto = String(input.dataset.lastAutoNumeroActa || "").trim();
    // Si el usuario sigue con el valor autocompletado, se puede refrescar en caliente.
    return Boolean(lastAuto) && current === lastAuto;
}

function setNumeroActaInputValue(input, value, markAsAuto = true) {
    if (!input) return;
    input.value = String(value || "").trim();
    if (markAsAuto) {
        input.dataset.lastAutoNumeroActa = input.value;
    }
}

function initNumeroActaTracking() {
    document.querySelectorAll('input[name="numero_acta"]').forEach((input) => {
        input.addEventListener("input", () => {
            const current = String(input.value || "").trim();
            if (!current) {
                input.dataset.lastAutoNumeroActa = "";
                return;
            }
            const lastAuto = String(input.dataset.lastAutoNumeroActa || "").trim();
            if (lastAuto && current === lastAuto) return;
            // Usuario modifico manualmente: ya no se sobreescribe en refresh suave.
            input.dataset.lastAutoNumeroActa = "";
        });
    });
}

async function refreshNumeroActaFormularioActivo(force = false) {
    const activeBtn = document.querySelector(".settings-menu-btn.active");
    const tipo = (activeBtn?.id || "tab-entrega").replace("tab-", "");
    const form = document.getElementById(`form-${tipo}`);
    if (!form) return;

    if (tipo === "aula") {
        await refreshNumeroActaAula(force);
        return;
    }

    const input = form.querySelector('input[name="numero_acta"]');
    if (!input) return;
    if (!shouldRefreshNumeroActaInput(input, force)) return;

    const nextNumero = await obtenerNumeroActaSiguiente(tipo);
    if (nextNumero) setNumeroActaInputValue(input, nextNumero, true);
}

async function refreshNumeroActaPorTipo(tipoActa, force = false) {
    const tipo = normalizeTipoActa(tipoActa || "entrega");
    if (tipo === "aula") {
        await refreshNumeroActaAula(force);
        return;
    }

    const form = document.getElementById(`form-${tipo}`);
    if (!form) {
        await refreshNumeroActaFormularioActivo(force);
        return;
    }

    const input = form.querySelector('input[name="numero_acta"]');
    if (!input) return;
    if (!shouldRefreshNumeroActaInput(input, force)) return;

    const nextNumero = await obtenerNumeroActaSiguiente(tipo);
    if (nextNumero) setNumeroActaInputValue(input, nextNumero, true);
}

async function obtenerNumeroActaAulaSiguiente() {
    try {
        const response = await fetch("/api/informes/areas/numero-acta/siguiente");
        const payload = await response.json();
        if (!response.ok || !payload.success) return null;
        return payload.numero_acta || null;
    } catch (_err) {
        return null;
    }
}

async function refreshNumeroActaAula(force = false) {
    const input = document.querySelector('#form-aula input[name="numero_acta"]');
    if (!input) return;
    if (!shouldRefreshNumeroActaInput(input, force)) return;

    const nextNumero = await obtenerNumeroActaAulaSiguiente();
    if (nextNumero) {
        setNumeroActaInputValue(input, nextNumero, true);
        updateAulaBatchPreviewCard();
    }
}

async function obtenerNumeroActaSiguiente(tipoActa = "entrega") {
    try {
        const response = await fetch(`/api/historial/numero-acta/siguiente?tipo_acta=${encodeURIComponent(String(tipoActa || "entrega"))}`);
        const payload = await response.json();
        if (!response.ok || !payload.success) return null;
        return payload.numero_acta || null;
    } catch (_err) {
        return null;
    }
}

async function autocompletarNumeroActa(form) {
    if (!form) return;
    const input = form.querySelector('input[name="numero_acta"]');
    if (!input) return;
    if (String(input.value || "").trim()) return;

    const formTipo = String(form.id || "").replace(/^form-/, "") || "entrega";
    const nextNumero = await obtenerNumeroActaSiguiente(formTipo);
    if (nextNumero) {
        setNumeroActaInputValue(input, nextNumero, true);
    }
}

async function validarNumeroActa(numeroActa, tipoActa = "entrega", editingActaId = null) {
    const value = String(numeroActa || "").trim();
    if (!value) {
        return { valid: false, error: "El número de acta es obligatorio." };
    }

    try {
        const params = new URLSearchParams({
            numero_acta: value,
            tipo_acta: String(tipoActa || "entrega"),
        });
        const editingId = Number(editingActaId || 0);
        if (Number.isInteger(editingId) && editingId > 0) {
            params.set("editing_acta_id", String(editingId));
        }
        const response = await fetch(`/api/historial/numero-acta/validar?${params.toString()}`);
        const payload = await response.json();
        if (!response.ok || !payload.success) {
            return { valid: false, error: "No se pudo validar el número de acta." };
        }
        if (payload.valid) return { valid: true };

        if (payload.exists) {
            return { valid: false, error: `Ya existe un acta con el número ${value}.` };
        }
        if (payload.reason === "format") {
            return { valid: false, error: "Formato inválido. Use 0NNN-AAAA (ej: 012-2026)." };
        }
        if (payload.reason === "year") {
            return { valid: false, error: "El año del número de acta debe ser el actual." };
        }
        return { valid: false, error: "Número de acta inválido." };
    } catch (_err) {
        return { valid: false, error: "Error de red al validar número de acta." };
    }
}

function initNumeroActaOnTabs() {
    refreshNumeroActaFormularioActivo(false);
    document.querySelectorAll(".settings-menu-btn").forEach((btn) => {
        btn.addEventListener("click", () => {
            setTimeout(() => {
                refreshNumeroActaFormularioActivo(false);
            }, 20);
        });
    });
}

function resetActaDraftState(tipo) {
    const t = String(tipo || "").trim().toLowerCase();
    activeEditingActaId = null;
    activeEditingActaTipo = null;
    if (t === "recepcion") {
        recepcionBienesTemp = [];
        clearRecepcionBienForm();
        setRecepcionEditMode(-1);
        renderRecepcionBienesTable();
        updateRecepcionSummary();
        return;
    }

    if (t === "entrega") {
        window._globalSelectedTableRows = [];
        window._globalSelectedColumns = [];
        const targetDiv = document.querySelector("#sec-entrega .resultado-tabla-container");
        hideResultadoTabla(targetDiv);
        return;
    }

    if (t === "bajas" || t === "baja") {
        bajasBienesTemp = [];
        bajasSelectedItemIds = new Set();
        bajasFilteredItems = [];
        bajasPage = 1;
        bajasStep = 1;
        updateBajasSummary();
    }
}

function getAulaBatchScopePayload() {
    const scope = String(document.getElementById("aula-scope")?.value || "area").trim().toLowerCase();
    const bloqueId = Number(document.getElementById("aula-bloque")?.value || 0);
    const pisoId = Number(document.getElementById("aula-piso")?.value || 0);
    const areaId = Number(document.getElementById("aula-area")?.value || 0);

    if (scope === "bloque" && !bloqueId) {
        return { error: "Seleccione un bloque para generar informes por bloque." };
    }
    if (scope === "piso" && !pisoId) {
        return { error: "Seleccione un piso para generar informes por piso." };
    }
    if (scope === "area" && !areaId) {
        return { error: "Seleccione un área para generar el informe." };
    }

    return {
        scope,
        bloque_id: bloqueId || null,
        piso_id: pisoId || null,
        area_id: areaId || null,
    };
}

function applyAulaScopeMode() {
    const scope = String(document.getElementById("aula-scope")?.value || "area").trim().toLowerCase();
    const bloqueSel = document.getElementById("aula-bloque");
    const pisoSel = document.getElementById("aula-piso");
    const areaSel = document.getElementById("aula-area");
    if (!bloqueSel || !pisoSel || !areaSel) return;

    if (scope === "bloque") {
        pisoSel.value = "";
        areaSel.value = "";
        pisoSel.disabled = true;
        areaSel.disabled = true;
        return;
    }

    if (scope === "piso") {
        areaSel.value = "";
        pisoSel.disabled = !String(bloqueSel.value || "").trim();
        areaSel.disabled = true;
        return;
    }

    // scope area
    pisoSel.disabled = !String(bloqueSel.value || "").trim();
    areaSel.disabled = !String(pisoSel.value || "").trim();
}

function getAulaBatchTargetCount() {
    const scope = String(document.getElementById("aula-scope")?.value || "area").trim().toLowerCase();
    const bloqueId = Number(document.getElementById("aula-bloque")?.value || 0);
    const pisoId = Number(document.getElementById("aula-piso")?.value || 0);
    const areaId = Number(document.getElementById("aula-area")?.value || 0);

    if (scope === "bloque") {
        const bloque = structureData.find((b) => Number(b.id) === bloqueId);
        if (!bloque) return 0;
        return (bloque.pisos || []).reduce((acc, piso) => acc + ((piso.areas || []).length), 0);
    }

    if (scope === "piso") {
        for (const bloque of structureData) {
            const piso = (bloque.pisos || []).find((p) => Number(p.id) === pisoId);
            if (piso) return (piso.areas || []).length;
        }
        return 0;
    }

    return areaId > 0 ? 1 : 0;
}

function parseNumeroActa(value) {
    const text = String(value || "").trim();
    const parts = text.split("-");
    if (parts.length !== 2) return null;

    const seq = Number(parts[0]);
    const year = Number(parts[1]);
    if (!Number.isFinite(seq) || !Number.isFinite(year) || seq <= 0 || year <= 0) return null;
    return { seq, year };
}

function buildNumeroActaRange(startNumeroActa, count) {
    const parsed = parseNumeroActa(startNumeroActa);
    const safeCount = Math.max(Number(count || 0), 0);
    if (!parsed || safeCount <= 0) return null;

    const start = `${String(parsed.seq).padStart(3, "0")}-${parsed.year}`;
    const endSeq = parsed.seq + safeCount - 1;
    const end = `${String(endSeq).padStart(3, "0")}-${parsed.year}`;
    return { start, end };
}

function updateAulaBatchPreviewCard() {
    applyAulaScopeMode();
    const countEl = document.getElementById("aula-batch-preview-count");
    const labelEl = document.getElementById("aula-batch-preview-label");
    const scope = String(document.getElementById("aula-scope")?.value || "area").trim().toLowerCase();
    const count = getAulaBatchTargetCount();
    const currentNumeroActa = String(document.querySelector('#form-aula input[name="numero_acta"]')?.value || "").trim();
    const estimatedRange = buildNumeroActaRange(currentNumeroActa, count);

    if (countEl) {
        countEl.textContent = String(count);
        countEl.classList.remove("bg-secondary", "bg-danger", "bg-success");
        countEl.classList.add(count > 0 ? "bg-success" : "bg-danger");
    }

    if (labelEl) {
        if (count <= 0) {
            labelEl.textContent = "No hay áreas seleccionadas para generar.";
        } else {
            const scopeText = scope === "bloque" ? "bloque" : scope === "piso" ? "piso" : "área";
            if (estimatedRange) {
                labelEl.textContent = `Se generarán ${count} acta(s) DOCX por ${scopeText}. Rango estimado: ${estimatedRange.start} a ${estimatedRange.end}.`;
            } else {
                labelEl.textContent = `Se generarán ${count} acta(s) DOCX por ${scopeText}.`;
            }
        }
    }

    document.querySelectorAll(".btn-generar-lote-aula").forEach((btn) => {
        btn.disabled = count <= 0;
    });
}

async function generarLoteAula() {
    const scopePayload = getAulaBatchScopePayload();
    if (scopePayload.error) {
        notify(scopePayload.error, true);
        return;
    }

    // Si no existe plantilla de Aula, no mostramos modal de progreso.
    try {
        const plantillaRes = await fetch("/api/plantillas/estado?tipo=aula");
        const plantillaPayload = await plantillaRes.json();
        if (!plantillaRes.ok || !plantillaPayload.success || !plantillaPayload.existe) {
            notify("No existe plantilla cargada para Aula. Cárguela en Configuración.", true);
            return;
        }
    } catch (_err) {
        notify("No se pudo verificar la plantilla de Aula.", true);
        return;
    }

    activeAulaBatchModal = showDownloadProgressModal();
    clearInterval(downloadProgressTimer);
    setDownloadBatchActionsVisible(true);
    setDownloadBatchPauseButton(false);
    setDownloadProgressTitle("Generando informes por área");
    setDownloadProgressMessage("Encolando tarea...");
    setDownloadProgressValue(2);
    try {
        notify("Generando actas DOCX masivas por ubicación...");
        const response = await fetch("/api/informes/areas/generar-lote", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(scopePayload),
        });

        const result = await response.json();
        if (!response.ok || !result.success) {
            notify(result.error || "No se pudo generar el lote de informes.", true);
            hideDownloadProgressModal(activeAulaBatchModal);
            activeAulaBatchModal = null;
            return;
        }

        if (result.immediate && result.download_path) {
            finishDownloadProgress();
            setDownloadProgressTitle("Descarga lista");
            setDownloadProgressMessage(
                String(result.download_kind || "docx").toLowerCase() === "zip"
                    ? "Paquete ZIP generado correctamente."
                    : "Documento Word generado correctamente."
            );

            const a = document.createElement("a");
            a.href = `/api/descargar?path=${encodeURIComponent(result.download_path)}`;
            a.style.display = "none";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);

            if (result.next_numero_acta) {
                const aulaInput = document.querySelector('#form-aula input[name="numero_acta"]');
                setNumeroActaInputValue(aulaInput, result.next_numero_acta, true);
                updateAulaBatchPreviewCard();
            }

            notify(
                `Se generó ${result.total_generated || 1} DOCX (${result.start_numero_acta || "-"}).`
            );
            clearTimeout(closeDownloadModalTimer);
            closeDownloadModalTimer = setTimeout(() => {
                hideDownloadProgressModal(activeAulaBatchModal);
                activeAulaBatchModal = null;
            }, 350);
            return;
        }

        activeAulaBatchJobId = String(result.job_id || "").trim() || null;
        setDownloadProgressMessage("Tarea en cola...");
        setDownloadProgressValue(5);
        notify(`Lote en cola (${result.total_targets || 0} destino(s)). Se notificará al terminar.`);
    } catch (_err) {
        notify("Error de red al generar informes masivos por ubicación.", true);
        clearTimeout(closeDownloadModalTimer);
        closeDownloadModalTimer = setTimeout(() => {
            hideDownloadProgressModal(activeAulaBatchModal);
            activeAulaBatchModal = null;
        }, 300);
    }
}

function setupTabs() {
    const tabs = document.querySelectorAll(".settings-menu-btn");
    const sections = document.querySelectorAll(".tab-section");
    const formContainer = document.getElementById("form-container");
    const previewColumn = document.getElementById("preview-column");

    const applyTabLayout = (targetId) => {
        const isAula = targetId === "sec-aula";
        if (previewColumn) {
            previewColumn.classList.toggle("d-none", isAula);
        }
        if (formContainer) {
            formContainer.classList.toggle("col-lg-6", !isAula);
            formContainer.classList.toggle("col-lg-10", isAula);
            formContainer.classList.toggle("border-end", !isAula);
        }
    };

    tabs.forEach((tab) => {
        tab.addEventListener("click", (e) => {
            e.preventDefault();
            tabs.forEach((t) => t.classList.remove("active"));
            sections.forEach((s) => s.classList.add("d-none"));
            tab.classList.add("active");
            const targetId = tab.getAttribute("data-target");
            document.getElementById(targetId)?.classList.remove("d-none");
            applyTabLayout(targetId);
        });
    });

    const activeTarget = document.querySelector(".settings-menu-btn.active")?.getAttribute("data-target") || "sec-entrega";
    applyTabLayout(activeTarget);
}

function populateLocationSelects() {
    const groups = ["entrega", "recepcion", "aula", "ext"];

    groups.forEach((pre) => {
        const selBloque = document.getElementById(`${pre}-bloque`);
        const selPiso = document.getElementById(`${pre}-piso`);
        const selArea = document.getElementById(`${pre}-area`);
        if (!selBloque || !selPiso || !selArea) return;

        selBloque.innerHTML = '<option value="">Seleccionar edificio...</option>';
        structureData.forEach((bloque) => {
            const opt = document.createElement("option");
            opt.value = bloque.id;
            opt.textContent = bloque.nombre;
            selBloque.appendChild(opt);
        });

        selBloque.addEventListener("change", () => {
            const bId = Number(selBloque.value);
            const bloque = structureData.find((b) => b.id === bId);
            selPiso.innerHTML = '<option value="">Seleccionar piso...</option>';
            selArea.innerHTML = '<option value="">Seleccionar area...</option>';
            selPiso.disabled = !bloque;
            selArea.disabled = true;
            if (!bloque) return;
            (bloque.pisos || []).forEach((piso) => {
                const opt = document.createElement("option");
                opt.value = piso.id;
                opt.textContent = piso.nombre;
                selPiso.appendChild(opt);
            });
            if (pre === "aula") updateAulaBatchPreviewCard();
            if (pre === "entrega" || pre === "recepcion") _updateUbicacionFromSelects(pre);
        });

        selPiso.addEventListener("change", () => {
            const bId = Number(selBloque.value);
            const pId = Number(selPiso.value);
            const bloque = structureData.find((b) => b.id === bId);
            const piso = (bloque?.pisos || []).find((p) => p.id === pId);
            selArea.innerHTML = '<option value="">Seleccionar area...</option>';
            selArea.disabled = !piso;
            if (!piso) return;
            (piso.areas || []).forEach((area) => {
                const opt = document.createElement("option");
                opt.value = area.id;
                opt.textContent = area.nombre;
                selArea.appendChild(opt);
            });
            if (pre === "aula") updateAulaBatchPreviewCard();
            if (pre === "entrega" || pre === "recepcion") _updateUbicacionFromSelects(pre);
        });

        if (pre === "aula") {
            selArea.addEventListener("change", () => {
                updateAulaBatchPreviewCard();
            });
        }

        if (pre === "recepcion") {
            selArea.addEventListener("change", () => {
                _updateUbicacionFromSelects(pre);
            });
        }

        if (pre === "entrega") {
            selArea.addEventListener("change", () => {
                _updateUbicacionFromSelects(pre);
            });
        }

        if (pre === "entrega" || pre === "recepcion") {
            _updateUbicacionFromSelects(pre);
        }
    });
}

function isEntregaAreaTrabajoValida(_value) {
    return String(document.getElementById("entrega-ubicacion-area-id")?.value || "").trim() !== "";
}

function isRecepcionAreaTrabajoValida(_value) {
    return String(document.getElementById("recepcion-ubicacion-area-id")?.value || "").trim() !== "";
}

function normalizeTipoActa(value) {
    return String(value || "").trim().toLowerCase();
}

function syncModalBackdropState() {
    const openModals = Array.from(document.querySelectorAll(".modal.show"));
    const backdrops = Array.from(document.querySelectorAll(".modal-backdrop"));
    const hasOpen = openModals.length > 0;

    document.body.classList.toggle("modal-open", hasOpen);
    if (!hasOpen) {
        backdrops.forEach((bd) => bd.remove());
        document.body.style.removeProperty("padding-right");
        return;
    }

    if (!backdrops.length) {
        const bd = document.createElement("div");
        bd.className = "modal-backdrop fade show";
        document.body.appendChild(bd);
    } else if (backdrops.length > 1) {
        backdrops.slice(0, -1).forEach((bd) => bd.remove());
    }
}

function setupModalBackdropGuardian() {
    if (window.__informeBackdropGuardianBound) return;
    window.__informeBackdropGuardianBound = true;

    document.addEventListener("shown.bs.modal", () => {
        setTimeout(syncModalBackdropState, 0);
    });
    document.addEventListener("hidden.bs.modal", () => {
        setTimeout(syncModalBackdropState, 0);
    });
}

function setPreviewStatus(message) {
    const status = document.getElementById("preview-status");
    if (status) status.textContent = message;
}

function showDownloadProgressModal() {
    const panelEl = document.getElementById("download-progress-floating");
    if (!panelEl) return null;
    panelEl.classList.remove("d-none");
    return panelEl;
}

function setDownloadProgressValue(percent) {
    const safe = Math.max(0, Math.min(100, Number(percent || 0)));
    const bar = document.getElementById("descarga-progress-bar");
    const text = document.getElementById("descarga-progress-text");
    downloadProgressValue = safe;
    if (bar) bar.style.width = `${safe}%`;
    if (text) text.textContent = `${safe}%`;
}

function setDownloadProgressMessage(message) {
    const msg = document.getElementById("descarga-progress-message");
    if (msg) msg.textContent = String(message || "").trim() || "Preparando descarga...";
}

function setDownloadProgressTitle(title) {
    const node = document.getElementById("descarga-progress-title");
    if (node) node.textContent = String(title || "").trim() || "Generando acta";
}

function setDownloadBatchActionsVisible(visible) {
    const actions = document.getElementById("download-progress-actions");
    if (!actions) return;
    actions.classList.toggle("d-none", !visible);
}

function setDownloadBatchPauseButton(paused) {
    activeAulaBatchPaused = Boolean(paused);
    const btn = document.getElementById("btn-download-progress-pause");
    if (!btn) return;
    if (activeAulaBatchPaused) {
        btn.innerHTML = '<i class="bi bi-play-fill me-1"></i>Reanudar';
    } else {
        btn.innerHTML = '<i class="bi bi-pause-fill me-1"></i>Pausar';
    }
}

async function controlAulaBatchJob(action) {
    const jobId = String(activeAulaBatchJobId || "").trim();
    if (!jobId) {
        notify("No hay un lote activo para controlar.", true);
        return;
    }

    try {
        const response = await fetch(`/api/informes/areas/jobs/${encodeURIComponent(jobId)}/control`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ action }),
        });
        const result = await response.json();
        if (!response.ok || !result.success) {
            notify(result.error || "No se pudo aplicar control al lote.", true);
            return;
        }

        if (action === "pause") setDownloadProgressMessage("Pausa solicitada...");
        if (action === "resume") setDownloadProgressMessage("Reanudando...");
        if (action === "cancel") setDownloadProgressMessage("Cancelación solicitada...");
    } catch (_err) {
        notify("Error de red al controlar el lote.", true);
    }
}

function hideDownloadProgressModal(modal) {
    const panelEl = modal || document.getElementById("download-progress-floating");
    if (!panelEl) return;
    panelEl.classList.add("d-none");
    setDownloadBatchActionsVisible(false);
}

function startDownloadProgress() {
    const bar = document.getElementById("descarga-progress-bar");
    const text = document.getElementById("descarga-progress-text");
    downloadProgressValue = 0;

    if (bar) bar.style.width = "0%";
    if (text) text.textContent = "0%";
    setDownloadProgressMessage("Preparando descarga...");

    clearInterval(downloadProgressTimer);
    downloadProgressTimer = setInterval(() => {
        downloadProgressValue = Math.min(92, downloadProgressValue + 4);
        if (bar) bar.style.width = `${downloadProgressValue}%`;
        if (text) text.textContent = `${downloadProgressValue}%`;
    }, 360);
}

function finishDownloadProgress() {
    const bar = document.getElementById("descarga-progress-bar");
    const text = document.getElementById("descarga-progress-text");
    clearInterval(downloadProgressTimer);
    downloadProgressValue = 100;
    if (bar) bar.style.width = "100%";
    if (text) text.textContent = "100%";
    setDownloadProgressMessage("Listo para descargar.");
}

function setManualPreviewOutput(result) {
    const iframe = document.getElementById("preview-iframe");
    const placeholder = document.getElementById("preview-placeholder");
    if (!iframe || !placeholder) return;

    if (result && result.pdf_path) {
        const pdfUrl = `/api/ver?path=${encodeURIComponent(String(result.pdf_path))}&t=${Date.now()}`;
        placeholder.classList.add("d-none");
        iframe.classList.remove("d-none");
        iframe.removeAttribute("srcdoc");
        iframe.src = pdfUrl;
        setPreviewStatus("PDF listo");
        return;
    }

    if (result && result.html_preview) {
        placeholder.classList.add("d-none");
        iframe.classList.remove("d-none");
        iframe.removeAttribute("src");
        const buildPreviewDoc = typeof window.buildInformePreviewDoc === "function"
            ? window.buildInformePreviewDoc
            : (value) => String(value || "");
        iframe.srcdoc = buildPreviewDoc(result.html_preview);
        setPreviewStatus("Vista HTML lista");
        return;
    }

    iframe.classList.add("d-none");
    placeholder.classList.remove("d-none");
    setPreviewStatus("Sin vista previa");
}

function formatFechaActa(value) {
    if (!value) return "-";
    const dt = new Date(value);
    if (Number.isNaN(dt.getTime())) return String(value);
    return dt.toLocaleString("es-EC", { hour12: false });
}

function mapTipoToTabId(tipo) {
    const t = normalizeTipoActa(tipo);
    const direct = `tab-${t}`;
    if (document.getElementById(direct)) return direct;
    return "tab-entrega";
}

function hideResultadoTabla(container) {
    if (!container) return;
    container.innerHTML = "";
    container.classList.add("d-none");
}

function showResultadoTabla(container, html) {
    if (!container) return;
    container.innerHTML = String(html || "");
    container.classList.remove("d-none");
}

function getActaTablePayload(tipo) {
    const t = normalizeTipoActa(tipo);
    if (t === "recepcion") {
        const rows = Array.isArray(recepcionBienesTemp) ? recepcionBienesTemp : [];
        const selectedIds = normalizeRecepcionColumnIds(recepcionSelectedColumnIds);
        const selected = selectedIds
            .map((id) => RECEPCION_BIENES_COLUMNS.find((col) => col.id === id))
            .filter(Boolean);
        return {
            datosTabla: rows,
            datosColumnas: selected,
        };
    }
    if (t === "bajas" || t === "baja") {
        const rows = Array.isArray(bajasBienesTemp) ? bajasBienesTemp : [];
        const selected = normalizeBajasColumnIds(bajasSelectedColumnIds)
            .map((id) => BAJAS_BIENES_COLUMNS.find((col) => col.id === id))
            .filter(Boolean);
        return {
            datosTabla: rows,
            datosColumnas: selected,
        };
    }
    return {
        datosTabla: Array.isArray(window._globalSelectedTableRows) ? window._globalSelectedTableRows : [],
        datosColumnas: Array.isArray(window._globalSelectedColumns) ? window._globalSelectedColumns : [],
    };
}

window.getInformeActaTablePayload = getActaTablePayload;

function getRecepcionResultContainer() {
    return document.querySelector("#sec-recepcion .resultado-tabla-container");
}

function getBajasResultContainer() {
    return document.querySelector("#sec-bajas .resultado-tabla-container");
}

function clearInformeFormByType(tipo) {
    const form = document.getElementById(`form-${tipo}`);
    const numeroActaInput = form?.querySelector('input[name="numero_acta"]') || null;
    const numeroActaValue = String(numeroActaInput?.value || "");
    const numeroActaLastAuto = String(numeroActaInput?.dataset?.lastAutoNumeroActa || "");

    if (form) form.reset();

    if (numeroActaInput) {
        numeroActaInput.value = numeroActaValue;
        numeroActaInput.dataset.lastAutoNumeroActa = numeroActaLastAuto;
    }

    const areaTrabajo = document.getElementById(`${tipo}-area-trabajo`);
    const areaId = document.getElementById(`${tipo}-ubicacion-area-id`);
    if (areaTrabajo) areaTrabajo.value = "";
    if (areaId) areaId.value = "";

    const bloque = document.getElementById(`${tipo}-bloque`);
    if (bloque) {
        bloque.value = "";
        bloque.dispatchEvent(new Event("change"));
    }

    const targetDiv = document.querySelector(`#sec-${tipo} .resultado-tabla-container`);
    if (targetDiv) hideResultadoTabla(targetDiv);
}

function clearAulaForm() {
    const numeroActaInput = document.querySelector('#form-aula input[name="numero_acta"]');
    const numeroActaValue = String(numeroActaInput?.value || "");
    const numeroActaLastAuto = String(numeroActaInput?.dataset?.lastAutoNumeroActa || "");

    const scope = document.getElementById("aula-scope");
    const bloque = document.getElementById("aula-bloque");

    if (scope) scope.value = "area";
    if (bloque) {
        bloque.value = "";
        bloque.dispatchEvent(new Event("change"));
    }

    if (typeof applyAulaScopeMode === "function") applyAulaScopeMode();
    if (typeof updateAulaBatchPreviewCard === "function") updateAulaBatchPreviewCard();

    if (numeroActaInput) {
        numeroActaInput.value = numeroActaValue;
        numeroActaInput.dataset.lastAutoNumeroActa = numeroActaLastAuto;
    }

    const targetDiv = document.querySelector("#sec-aula .resultado-tabla-container");
    if (targetDiv) hideResultadoTabla(targetDiv);
}

function updateRecepcionSummary() {
    const targetDiv = getRecepcionResultContainer();
    const total = Array.isArray(recepcionBienesTemp) ? recepcionBienesTemp.length : 0;
    if (!targetDiv) return;

    if (total <= 0) {
        hideResultadoTabla(targetDiv);
        return;
    }

    showResultadoTabla(
        targetDiv,
        `<div class="alert alert-success d-inline-block p-2 px-4 shadow-sm mb-0"><i class="bi bi-check2-circle me-2"></i>Recepción: <strong>${total}</strong> bien(es) nuevo(s) registrados temporalmente.</div>`
    );
}

function updateBajasSummary() {
    const targetDiv = getBajasResultContainer();
    const total = Array.isArray(bajasBienesTemp) ? bajasBienesTemp.length : 0;
    if (!targetDiv) return;

    if (total <= 0) {
        hideResultadoTabla(targetDiv);
        return;
    }

    showResultadoTabla(
        targetDiv,
        `<div class="alert alert-warning d-inline-block p-2 px-4 shadow-sm mb-0"><i class="bi bi-trash me-2"></i>Bajas: <strong>${total}</strong> bien(es) preparado(s) para dar de baja.</div>`
    );
}

function getRecepcionBienFormValues() {
    const cantidadRaw = String(document.getElementById("recepcion-bien-cantidad")?.value || "1").trim();
    const cantidad = Number(cantidadRaw);
    return {
        cod_inventario: normalizeCodeToPlaceholder(document.getElementById("recepcion-bien-cod-inventario")?.value),
        cod_esbye: normalizeCodeToPlaceholder(document.getElementById("recepcion-bien-cod-esbye")?.value),
        cuenta: String(document.getElementById("recepcion-bien-cuenta")?.value || "").trim(),
        cantidad: Number.isFinite(cantidad) && cantidad > 0 ? Math.trunc(cantidad) : 0,
        descripcion: String(document.getElementById("recepcion-bien-descripcion")?.value || "").trim(),
        marca: String(document.getElementById("recepcion-bien-marca")?.value || "").trim(),
        modelo: String(document.getElementById("recepcion-bien-modelo")?.value || "").trim(),
        serie: String(document.getElementById("recepcion-bien-serie")?.value || "").trim(),
        estado: String(document.getElementById("recepcion-bien-estado")?.value || "").trim(),
        ubicacion: String(document.getElementById("recepcion-bien-ubicacion")?.value || document.getElementById("recepcion-area-trabajo")?.value || "").trim(),
        fecha_adquisicion: String(document.getElementById("recepcion-bien-fecha-adquisicion")?.value || "").trim(),
        valor: String(document.getElementById("recepcion-bien-valor")?.value || "").trim(),
        usuario_final: String(document.getElementById("recepcion-bien-usuario-final")?.value || "").trim(),
        observacion: String(document.getElementById("recepcion-bien-observacion")?.value || "").trim(),
        descripcion_esbye: String(document.getElementById("recepcion-bien-descripcion-esbye")?.value || "").trim(),
        marca_esbye: String(document.getElementById("recepcion-bien-marca-esbye")?.value || "").trim(),
        modelo_esbye: String(document.getElementById("recepcion-bien-modelo-esbye")?.value || "").trim(),
        serie_esbye: String(document.getElementById("recepcion-bien-serie-esbye")?.value || "").trim(),
        fecha_adquisicion_esbye: String(document.getElementById("recepcion-bien-fecha-esbye")?.value || "").trim(),
        valor_esbye: String(document.getElementById("recepcion-bien-valor-esbye")?.value || "").trim(),
        ubicacion_esbye: String(document.getElementById("recepcion-bien-ubicacion-esbye")?.value || "").trim(),
        observacion_esbye: String(document.getElementById("recepcion-bien-observacion-esbye")?.value || "").trim(),
    };
}

function clearRecepcionBienForm() {
    const ids = [
        "recepcion-excel-single-row",
        "recepcion-bien-cod-inventario",
        "recepcion-bien-cod-esbye",
        "recepcion-bien-cuenta",
        "recepcion-bien-descripcion",
        "recepcion-bien-marca",
        "recepcion-bien-modelo",
        "recepcion-bien-serie",
        "recepcion-bien-estado",
        "recepcion-bien-fecha-adquisicion",
        "recepcion-bien-valor",
        "recepcion-bien-usuario-final",
        "recepcion-bien-observacion",
        "recepcion-bien-descripcion-esbye",
        "recepcion-bien-marca-esbye",
        "recepcion-bien-modelo-esbye",
        "recepcion-bien-serie-esbye",
        "recepcion-bien-fecha-esbye",
        "recepcion-bien-valor-esbye",
        "recepcion-bien-ubicacion-esbye",
        "recepcion-bien-observacion-esbye",
    ];
    ids.forEach((id) => {
        const el = document.getElementById(id);
        if (el) el.value = "";
    });
    const cantidad = document.getElementById("recepcion-bien-cantidad");
    if (cantidad) cantidad.value = "1";
}

function setRecepcionEditMode(index) {
    recepcionEditIndex = Number.isInteger(index) ? index : -1;
    const btnSave = document.getElementById("btn-recepcion-bien-guardar");
    const btnCancel = document.getElementById("btn-recepcion-bien-cancelar");
    if (btnSave) {
        btnSave.innerHTML = recepcionEditIndex >= 0
            ? '<i class="bi bi-pencil-square me-1"></i>Editar bien'
            : '<i class="bi bi-plus-circle me-1"></i>Guardar bien';
    }
    if (btnCancel) {
        btnCancel.classList.toggle("d-none", recepcionEditIndex < 0);
    }
}

function loadRecepcionBienIntoForm(index) {
    const item = recepcionBienesTemp[index];
    if (!item) return;

    const normalizeSelectText = (value) => String(value || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .trim();

    const assignSelect = (id, value) => {
        const el = document.getElementById(id);
        const val = String(value || "").trim();
        if (!el) return;
        if (!val) {
            el.value = "";
            return;
        }

        const options = Array.from(el.options || []);
        const normalizedVal = normalizeSelectText(val);
        let matched = options.find((opt) => normalizeSelectText(opt.value) === normalizedVal);
        if (!matched) {
            matched = options.find((opt) => normalizeSelectText(opt.textContent) === normalizedVal);
        }
        if (!matched) {
            matched = options.find((opt) => normalizeSelectText(opt.textContent).includes(normalizedVal));
        }
        if (!matched) {
            const valTokens = normalizedVal.split(/\s+/).filter(Boolean);
            let best = null;
            let bestScore = 0;
            options.forEach((opt) => {
                const optNorm = normalizeSelectText(opt.value || opt.textContent);
                if (!optNorm) return;
                const optTokens = optNorm.split(/\s+/).filter(Boolean);
                const overlap = valTokens.filter((token) => optTokens.some((item) => item.includes(token) || token.includes(item))).length;
                const score = valTokens.length ? overlap / valTokens.length : 0;
                if (score > bestScore) {
                    bestScore = score;
                    best = opt;
                }
            });
            if (best && bestScore >= 0.6) matched = best;
        }

        if (matched) {
            el.value = String(matched.value || "");
            return;
        }

        if (!options.some((opt) => String(opt.value) === val)) {
            const opt = document.createElement("option");
            opt.value = val;
            opt.textContent = val;
            el.appendChild(opt);
        }
        el.value = val;
    };
    const assign = (id, value) => {
        const el = document.getElementById(id);
        if (el) el.value = String(value || "");
    };
    assign("recepcion-bien-cod-inventario", item.cod_inventario);
    assign("recepcion-bien-cod-esbye", item.cod_esbye);
    assignSelect("recepcion-bien-cuenta", item.cuenta);
    assign("recepcion-bien-cantidad", item.cantidad || 1);
    assign("recepcion-bien-descripcion", item.descripcion);
    assign("recepcion-bien-marca", item.marca);
    assign("recepcion-bien-modelo", item.modelo);
    assign("recepcion-bien-serie", item.serie);
    assignSelect("recepcion-bien-estado", item.estado);
    assign("recepcion-bien-fecha-adquisicion", normalizeImportDateLikeInventario(item.fecha_adquisicion));
    assign("recepcion-bien-valor", item.valor);
    assignSelect("recepcion-bien-usuario-final", item.usuario_final);
    assign("recepcion-bien-observacion", item.observacion);
    assign("recepcion-bien-descripcion-esbye", item.descripcion_esbye);
    assign("recepcion-bien-marca-esbye", item.marca_esbye);
    assign("recepcion-bien-modelo-esbye", item.modelo_esbye);
    assign("recepcion-bien-serie-esbye", item.serie_esbye);
    assign("recepcion-bien-fecha-esbye", normalizeImportDateLikeInventario(item.fecha_adquisicion_esbye || item.fecha_esbye));
    assign("recepcion-bien-valor-esbye", item.valor_esbye);
    assign("recepcion-bien-ubicacion-esbye", item.ubicacion_esbye || "");
    assign("recepcion-bien-observacion-esbye", item.observacion_esbye);
    setRecepcionEditMode(index);
}

function renderRecepcionBienesTable() {
    const tbody = document.getElementById("tbody-recepcion-bienes");
    const thead = document.getElementById("thead-recepcion-bienes");
    const countBadge = document.getElementById("recepcion-bienes-count");
    const pageInfo = document.getElementById("recepcion-bienes-page-info");
    const pagePrev = document.getElementById("recepcion-bienes-page-prev");
    const pageNext = document.getElementById("recepcion-bienes-page-next");
    const pageSize = document.getElementById("recepcion-bienes-page-size");
    if (!tbody || !thead || !countBadge) return;

    const selectedIds = normalizeRecepcionColumnIds(recepcionSelectedColumnIds);
    const escapeCell = (text) => String(text || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#39;");
    const selectedCols = selectedIds
        .map((id) => RECEPCION_BIENES_COLUMNS.find((col) => col.id === id))
        .filter(Boolean);

    thead.innerHTML = "";
    const thIndex = document.createElement("th");
    thIndex.style.width = "55px";
    thIndex.textContent = "N";
    thead.appendChild(thIndex);
    selectedCols.forEach((col) => {
        const th = document.createElement("th");
        th.textContent = col.label;
        thead.appendChild(th);
    });
    const thActions = document.createElement("th");
    thActions.style.width = "90px";
    thActions.className = "text-center";
    thActions.textContent = "Acciones";
    thead.appendChild(thActions);

    const totalItems = recepcionBienesTemp.length;
    countBadge.textContent = String(totalItems);
    if (!totalItems) {
        tbody.innerHTML = `<tr><td colspan="${selectedCols.length + 2}" class="text-center text-muted">Aún no hay bienes registrados.</td></tr>`;
        if (pageInfo) pageInfo.textContent = "0 de 0";
        if (pagePrev) pagePrev.disabled = true;
        if (pageNext) pageNext.disabled = true;
        return;
    }

    if (pageSize) {
        pageSize.value = String(recepcionTablePerPage);
    }

    const totalPages = Math.max(1, Math.ceil(totalItems / recepcionTablePerPage));
    recepcionTablePage = Math.min(Math.max(recepcionTablePage, 1), totalPages);
    const startIndex = (recepcionTablePage - 1) * recepcionTablePerPage;
    const pagedItems = recepcionBienesTemp.slice(startIndex, startIndex + recepcionTablePerPage);

    if (pageInfo) pageInfo.textContent = `Página ${recepcionTablePage} de ${totalPages}`;
    if (pagePrev) pagePrev.disabled = recepcionTablePage <= 1;
    if (pageNext) pageNext.disabled = recepcionTablePage >= totalPages;

    tbody.innerHTML = "";
    pagedItems.forEach((item, idx) => {
        const globalIndex = startIndex + idx;
        const tr = document.createElement("tr");
        tr.style.cursor = "pointer";
        tr.dataset.index = String(globalIndex);
        let rowHtml = `<td>${globalIndex + 1}</td>`;
        selectedCols.forEach((col) => {
            const raw = item?.[col.id];
            const value = raw == null || String(raw).trim() === "" ? "-" : String(raw);
            const codeClass = (col.id === "cod_inventario" || col.id === "cod_esbye") && isNoCodeValue(value)
                ? " class=\"code-sc-cell\""
                : "";
            rowHtml += `<td${codeClass} title="${escapeCell(value)}">${escapeCell(value)}</td>`;
        });
        rowHtml += `<td class="text-center"><button type="button" class="btn btn-sm btn-outline-danger btn-recepcion-bien-eliminar" data-index="${globalIndex}"><i class="bi bi-trash"></i></button></td>`;
        tr.innerHTML = rowHtml;
        tbody.appendChild(tr);
    });

    tbody.querySelectorAll("tr").forEach((tr) => {
        tr.addEventListener("click", async (event) => {
            if (event.target.closest(".btn-recepcion-bien-eliminar")) return;
            const idx = Number(tr.dataset.index);
            if (!Number.isFinite(idx)) return;
            // Asegura que los catálogos de selects estén listos antes de cargar el item.
            if (typeof window.__loadRecepcionSelectOptions === "function") {
                await window.__loadRecepcionSelectOptions();
            }
            loadRecepcionBienIntoForm(idx);
        });
    });

    tbody.querySelectorAll(".btn-recepcion-bien-eliminar").forEach((btn) => {
        btn.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            const idx = Number(btn.dataset.index);
            if (!Number.isFinite(idx)) return;
            const confirmed = window.confirm("¿Deseas eliminar este bien de la lista temporal?");
            if (!confirmed) return;
            recepcionBienesTemp.splice(idx, 1);
            if (recepcionEditIndex === idx) {
                clearRecepcionBienForm();
                setRecepcionEditMode(-1);
            } else if (recepcionEditIndex > idx) {
                setRecepcionEditMode(recepcionEditIndex - 1);
            }
            renderRecepcionBienesTable();
            updateRecepcionSummary();
        });
    });
}

async function checkRecepcionDuplicates(payload) {
    const codInv = String(payload.cod_inventario || "").trim();
    const codEsbye = String(payload.cod_esbye || "").trim();
    if (!codInv && !codEsbye) return [];

    try {
        const query = new URLSearchParams();
        if (codInv) query.set("cod_inventario", codInv);
        if (codEsbye) query.set("cod_esbye", codEsbye);
        const response = await fetch(`/api/inventario/duplicados?${query.toString()}`);
        const result = await response.json();
        if (!response.ok || !result.success) return [];
        return Array.isArray(result.duplicates) ? result.duplicates : [];
    } catch (_err) {
        return [];
    }
}

function setupRecepcionBienesModal() {
    const modalEl = document.getElementById("modalRecepcionBienes");
    const importModalEl = document.getElementById("modalRecepcionImportarExcel");
    const columnsModalEl = document.getElementById("modalRecepcionColumnas");
    const btnOpen = document.querySelector("#sec-recepcion .btn-registrar-bienes");
    const btnSave = document.getElementById("btn-recepcion-bien-guardar");
    const btnCancel = document.getElementById("btn-recepcion-bien-cancelar");
    const btnClear = document.getElementById("btn-recepcion-bien-vaciar");
    const btnConfirm = document.getElementById("btn-confirmar-recepcion-bienes");
    const btnImportExcel = document.getElementById("btn-recepcion-importar-excel");
    const inputImportExcel = document.getElementById("recepcion-import-file-input");
    const importStatus = document.getElementById("recepcion-import-status");
    const mappingSelectsRow = document.getElementById("recepcion-import-mapping-selects-row");
    const mappingHeadersRow = document.getElementById("recepcion-import-mapping-headers-row");
    const mappingPreviewBody = document.getElementById("recepcion-import-mapping-preview-body");
    const previewInfo = document.getElementById("recepcion-import-preview-info");
    const previewRunBody = document.getElementById("recepcion-import-preview-run-body");
    const chunkRange = document.getElementById("recepcion-import-run-info");
    const rowCount = null;
    const reviewSummary = document.getElementById("recepcion-import-run-result");
    const btnImportBack = document.getElementById("btn-recepcion-import-back");
    const btnImportNext = document.getElementById("btn-recepcion-import-next");
    const btnImportRun = document.getElementById("btn-recepcion-import-run");
    const btnColumnas = document.getElementById("btn-recepcion-columnas");
    const btnColumnasGuardar = document.getElementById("btn-recepcion-columnas-guardar");
    const recepcionColumnSelector = document.getElementById("recepcion-column-selector");
    const recepcionExcelSingleRow = document.getElementById("recepcion-excel-single-row");
    const recepcionPagePrev = document.getElementById("recepcion-bienes-page-prev");
    const recepcionPageNext = document.getElementById("recepcion-bienes-page-next");
    const recepcionPageSize = document.getElementById("recepcion-bienes-page-size");
    const panel1 = document.getElementById("recepcion-import-step-file");
    const panel2 = document.getElementById("recepcion-import-step-mapping");
    const panel3 = document.getElementById("recepcion-import-step-run");
    const stepBadges = [];
    const stepLabels = [];
    const ubicacionReadonly = document.getElementById("recepcion-bien-ubicacion");
    const valorInput = document.getElementById("recepcion-bien-valor");
    const valorEsbyeInput = document.getElementById("recepcion-bien-valor-esbye");
    const cuentaSelect = document.getElementById("recepcion-bien-cuenta");
    const estadoSelect = document.getElementById("recepcion-bien-estado");
    const usuarioFinalSelect = document.getElementById("recepcion-bien-usuario-final");
    if (!modalEl || !btnOpen || !btnSave || !btnConfirm) return;

    const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
    const duplicateNodes = {
        duplicateModalEl: document.getElementById("modalDuplicadosInventario"),
        duplicateSummary: document.getElementById("duplicados-resumen"),
        duplicateList: document.getElementById("duplicados-lista"),
        duplicateCancelBtn: document.getElementById("btn-duplicados-cancelar"),
        duplicateContinueBtn: document.getElementById("btn-duplicados-continuar"),
    };
    const duplicateModal = duplicateNodes.duplicateModalEl
        ? bootstrap.Modal.getOrCreateInstance(duplicateNodes.duplicateModalEl)
        : null;
    const importModal = importModalEl
        ? bootstrap.Modal.getOrCreateInstance(importModalEl)
        : null;
    const columnsModal = columnsModalEl
        ? bootstrap.Modal.getOrCreateInstance(columnsModalEl)
        : null;

    const importConflictsModalEl = document.getElementById("modalImportConflicts");
    const importConflictsSummary = document.getElementById("import-conflicts-summary");
    const importConflictsAccordion = document.getElementById("import-conflicts-accordion");
    const importConflictsCancelBtn = document.getElementById("btn-import-conflicts-cancel");
    const importConflictsContinueBtn = document.getElementById("btn-import-conflicts-continue");
    const importConflictsModal = importConflictsModalEl
        ? bootstrap.Modal.getOrCreateInstance(importConflictsModalEl)
        : null;

    const duplicateController = typeof window.createDuplicateInventoryModal === "function"
        ? window.createDuplicateInventoryModal({
            duplicateModal,
            nodes: duplicateNodes,
            escapeHtmlText: (value) => String(value || "")
                .replaceAll("&", "&amp;")
                .replaceAll("<", "&lt;")
                .replaceAll(">", "&gt;")
                .replaceAll('"', "&quot;")
                .replaceAll("'", "&#39;"),
            buildDuplicateWarningMessage: (duplicates = [], payload = {}) => {
                const inv = String(payload.cod_inventario || "").trim();
                const esbye = String(payload.cod_esbye || "").trim();
                return `Se encontraron ${duplicates.length} coincidencia(s).${inv ? ` INV: ${inv}` : ""}${esbye ? ` ESBYE: ${esbye}` : ""}`;
            },
        })
        : null;

    const openDuplicateModal = duplicateController?.openDuplicateModal
        ? duplicateController.openDuplicateModal
        : async () => true;

    [valorInput, valorEsbyeInput].forEach((input) => {
        if (!input) return;
        input.addEventListener("input", () => validarNumerosComasInput(input));
    });

    const assignRecepcionPastedValue = (field, rawValue) => {
        const inputId = RECEPCION_MODAL_FIELD_TO_INPUT_ID[field];
        if (!inputId) return;
        const input = document.getElementById(inputId);
        if (!input) return;

        const value = String(rawValue ?? "").trim();
        if (!value) return;

        if (field === "cod_inventario" || field === "cod_esbye") {
            input.value = normalizeCodeToPlaceholder(value);
            return;
        }

        if (field === "valor" || field === "valor_esbye") {
            let normalizedMoney = value.replace(/[^\d,.-]/g, "");
            const lastComma = normalizedMoney.lastIndexOf(",");
            const lastDot = normalizedMoney.lastIndexOf(".");

            if (lastComma >= 0 && lastDot >= 0) {
                if (lastComma > lastDot) {
                    normalizedMoney = normalizedMoney.replace(/\./g, "");
                } else {
                    normalizedMoney = normalizedMoney.replace(/,/g, "");
                    const parts = normalizedMoney.split(".");
                    const decimal = parts.pop();
                    normalizedMoney = `${parts.join("")}${decimal !== undefined ? `,${decimal}` : ""}`;
                }
            } else if (lastDot >= 0) {
                const parts = normalizedMoney.split(".");
                const decimal = parts.pop();
                normalizedMoney = `${parts.join("")}${decimal !== undefined ? `,${decimal}` : ""}`;
            }

            normalizedMoney = normalizedMoney.replace(/-/g, "");
            const firstComma = normalizedMoney.indexOf(",");
            if (firstComma !== -1) {
                normalizedMoney = normalizedMoney.slice(0, firstComma + 1) + normalizedMoney.slice(firstComma + 1).replace(/,/g, "");
            }
            input.value = normalizedMoney;
            return;
        }

        if (field === "cantidad") {
            let normalized = value;
            if (value.includes(",") && value.includes(".")) {
                normalized = value.replace(/\./g, "").replace(",", ".");
            } else if (value.includes(",")) {
                normalized = value.replace(",", ".");
            }
            const numeric = Number(normalized);
            if (Number.isFinite(numeric) && numeric > 0) {
                input.value = String(Math.trunc(numeric));
            }
            return;
        }

        if (field === "fecha_adquisicion" || field === "fecha_adquisicion_esbye") {
            input.value = normalizeImportDateLikeInventario(value);
            return;
        }

        if (input.tagName === "SELECT") {
            const options = Array.from(input.options || []);
            const match = resolveSelectOptionBestMatch(options, value, field);
            if (match) {
                input.value = match.value;
            } else if (field === "usuario_final") {
                const dynamicOption = document.createElement("option");
                dynamicOption.value = value;
                dynamicOption.textContent = value;
                input.appendChild(dynamicOption);
                input.value = value;
            }
            return;
        }

        if ((input.type || "").toLowerCase() === "date") {
            input.value = normalizeImportDateLikeInventario(value);
            return;
        }

        if ((input.type || "").toLowerCase() === "number") {
            let normalized = value;
            if (value.includes(",") && value.includes(".")) {
                normalized = value.replace(/\./g, "").replace(",", ".");
            } else if (value.includes(",")) {
                normalized = value.replace(",", ".");
            }
            input.value = normalized;
            return;
        }

        input.value = value;
    };

    if (recepcionExcelSingleRow) {
        recepcionExcelSingleRow.addEventListener("paste", async (event) => {
            const text = event.clipboardData.getData("text/plain");
            if (!text.trim()) return;
            event.preventDefault();

            await loadRecepcionSelectOptions();
            const rows = parseExcelText(text);
            if (!rows.length) return;

            const mappedPrimary = mapPastedRowBestEffortForRecepcion(rows[0]);
            let mapped = mappedPrimary;

            if (rows.length > 1) {
                const mergedCells = rows.flat();
                const mappedMerged = mapPastedRowBestEffortForRecepcion(mergedCells);
                if (scoreRecepcionPastedMapping(mappedMerged) > scoreRecepcionPastedMapping(mappedPrimary)) {
                    mapped = mappedMerged;
                }
            }

            RECEPCION_MODAL_PASTE_FIELDS.forEach((field) => {
                if (field === "ubicacion") return;
                if (mapped[field] === undefined) return;
                assignRecepcionPastedValue(field, mapped[field]);
            });

            syncRecepcionUbicacionReadonly();
        });
    }

    const CHUNK_SIZE = 20;
    let shouldReturnToRecepcionModal = false;
    let shouldReturnFromColumnsModal = false;
    let pendingChildModal = "";
    let isTransientImportModalHide = false;
    let recepcionSelectsRequestId = 0;
    const RECEPCION_CANONICAL_FIELDS = [
        { value: "", label: "- Ignorar columna -" },
        { value: "cod_inventario", label: "Cod Inv." },
        { value: "cod_esbye", label: "Cod. ESBYE" },
        { value: "cuenta", label: "Cuenta" },
        { value: "cantidad", label: "Cantidad" },
        { value: "descripcion", label: "Descripcion" },
        { value: "marca", label: "Marca" },
        { value: "modelo", label: "Modelo" },
        { value: "serie", label: "Serie" },
        { value: "estado", label: "Estado" },
        { value: "ubicacion", label: "Ubicacion" },
        { value: "fecha_adquisicion", label: "Fecha Adquisicion" },
        { value: "valor", label: "Valor" },
        { value: "usuario_final", label: "Usuario Final" },
        { value: "observacion", label: "Observacion" },
        { value: "descripcion_esbye", label: "Descripcion ESBYE" },
        { value: "marca_esbye", label: "Marca ESBYE" },
        { value: "modelo_esbye", label: "Modelo ESBYE" },
        { value: "serie_esbye", label: "Serie ESBYE" },
        { value: "fecha_adquisicion_esbye", label: "Fecha ESBYE" },
        { value: "valor_esbye", label: "Valor ESBYE" },
        { value: "ubicacion_esbye", label: "Ubicacion ESBYE" },
        { value: "observacion_esbye", label: "Observacion ESBYE" },
    ];

    const BASE_TO_ESBYE_FIELD = {
        descripcion: "descripcion_esbye",
        marca: "marca_esbye",
        modelo: "modelo_esbye",
        serie: "serie_esbye",
        fecha_adquisicion: "fecha_adquisicion_esbye",
        valor: "valor_esbye",
        ubicacion: "ubicacion_esbye",
        observacion: "observacion_esbye",
    };

    const importState = {
        step: 1,
        sessionId: null,
        sourceFile: null,
        headers: [],
        previewRows: [],
        totalRows: 0,
        mappings: [],
        lockedMainLocationColIdx: null,
        startIndex: 0,
        chunkSize: CHUNK_SIZE,
        hasMore: false,
        previewChunkData: null,
        isRecoveringSession: false,
    };

    const escapeHtml = (text) => String(text || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#39;");

    const getStatusBadge = (status) => {
        if (status === "exact") return '<span class="badge text-bg-danger">Registrado (igual)</span>';
        if (status === "similar") return '<span class="badge text-bg-warning">Similar</span>';
        return '<span class="badge text-bg-success">Nuevo</span>';
    };

    const buildConflictDetailsHtml = (rowData, rowPosition) => {
        const exact = Array.isArray(rowData.exact_matches) ? rowData.exact_matches : [];
        const similar = Array.isArray(rowData.similar_matches) ? rowData.similar_matches : [];
        const allMatches = exact.length ? exact : similar;
        const statusLabel = exact.length ? "Coincidencia exacta" : "Coincidencia similar";
        const fields = [
            { label: "Fila Excel", value: (rowData.row_index || 0) + 1 },
            { label: "Cod. Inventario", value: rowData.data?.cod_inventario || "-" },
            { label: "Cod. ESBYE", value: rowData.data?.cod_esbye || "-" },
            { label: "Descripcion", value: rowData.data?.descripcion || "-" },
            { label: "Marca", value: rowData.data?.marca || "-" },
            { label: "Modelo", value: rowData.data?.modelo || "-" },
            { label: "Serie", value: rowData.data?.serie || "-" },
        ];
        const fieldsHtml = fields
            .map((item) => `<div class="col-md-4 small"><strong>${escapeHtml(item.label)}:</strong> ${escapeHtml(item.value)}</div>`)
            .join("");
        const repeatedRowsHtml = allMatches.length
            ? allMatches.map((item) => `
                <div class="import-conflict-match-row">
                    <div class="small"><strong>Item #${escapeHtml(item.item_numero || "-")}</strong></div>
                    <div class="small text-muted">INV: ${escapeHtml(item.cod_inventario || "-")} | ESBYE: ${escapeHtml(item.cod_esbye || "-")}</div>
                    <div class="small text-warning-emphasis">Coincide por: ${escapeHtml((item.match_fields || []).length ? item.match_fields.join(", ") : "revision manual")}</div>
                    <div class="small">${escapeHtml(item.descripcion || "-")}</div>
                    <div class="small text-muted">Ubicacion: ${escapeHtml(item.ubicacion || "-")} | Usuario: ${escapeHtml(item.usuario_final || "-")}</div>
                </div>
            `).join("")
            : '<div class="small text-muted">Sin coincidencias listables.</div>';

        return `
            <div class="accordion-item border-warning-subtle">
                <h2 class="accordion-header" id="recepcion-conflict-head-${rowPosition}">
                    <button class="accordion-button ${rowPosition > 0 ? "collapsed" : ""}" type="button" data-bs-toggle="collapse" data-bs-target="#recepcion-conflict-body-${rowPosition}" aria-expanded="${rowPosition === 0 ? "true" : "false"}">
                        <div class="d-flex flex-wrap align-items-center gap-2 w-100 pe-3">
                            <span class="badge text-bg-warning">${escapeHtml(statusLabel)}</span>
                            <span class="small fw-semibold">Fila ${(rowData.row_index || 0) + 1}</span>
                            <span class="small text-muted">${escapeHtml(rowData.data?.descripcion || "Sin descripcion")}</span>
                        </div>
                    </button>
                </h2>
                <div id="recepcion-conflict-body-${rowPosition}" class="accordion-collapse collapse ${rowPosition === 0 ? "show" : ""}" data-bs-parent="#import-conflicts-accordion">
                    <div class="accordion-body pt-2">
                        <div class="row g-2 mb-2">${fieldsHtml}</div>
                        <div class="import-conflict-matches-wrap">${repeatedRowsHtml}</div>
                    </div>
                </div>
            </div>
        `;
    };

    const openImportConflictsModal = (validationData) => {
        const summary = validationData?.summary || {};
        const rows = Array.isArray(validationData?.analyzed_rows) ? validationData.analyzed_rows : [];
        const conflictRows = rows.filter((row) => row.status === "exact" || row.status === "similar");

        if (!importConflictsModal || !importConflictsSummary || !importConflictsAccordion || !importConflictsCancelBtn || !importConflictsContinueBtn) {
            return Promise.resolve(window.confirm(`Conflictos detectados. Iguales: ${summary.exact || 0}, similares: ${summary.similar || 0}. ¿Deseas continuar?`));
        }

        importConflictsSummary.textContent = `Se detectaron ${summary.exact || 0} fila(s) iguales y ${summary.similar || 0} similares en este bloque.`;
        importConflictsAccordion.innerHTML = conflictRows.length
            ? conflictRows.map((row, idx) => buildConflictDetailsHtml(row, idx)).join("")
            : '<div class="alert alert-secondary mb-0">No hay conflictos detallados para mostrar.</div>';

        return new Promise((resolve) => {
            const importWasVisible = Boolean(importModalEl && importModalEl.classList.contains("show"));
            let settled = false;
            const finish = (value) => {
                if (settled) return;
                settled = true;
                cleanup();
                resolve(value);
            };
            const reopenImportIfNeeded = () => {
                if (!importWasVisible) return;
                setTimeout(() => {
                    importModal?.show();
                    syncModalBackdropState();
                }, 120);
            };
            const onContinue = () => {
                finish(true);
                importConflictsModal.hide();
                reopenImportIfNeeded();
            };
            const onCancel = () => {
                finish(false);
                importConflictsModal.hide();
                reopenImportIfNeeded();
            };
            const onHidden = () => {
                finish(false);
                reopenImportIfNeeded();
            };
            const cleanup = () => {
                importConflictsContinueBtn.removeEventListener("click", onContinue);
                importConflictsCancelBtn.removeEventListener("click", onCancel);
                importConflictsModalEl.removeEventListener("hidden.bs.modal", onHidden);
            };

            importConflictsContinueBtn.addEventListener("click", onContinue);
            importConflictsCancelBtn.addEventListener("click", onCancel);
            importConflictsModalEl.addEventListener("hidden.bs.modal", onHidden);

            const showConflicts = () => {
                importConflictsModal.show();
                syncModalBackdropState();
            };

            if (importWasVisible) {
                const onImportHiddenOnce = () => {
                    importModalEl.removeEventListener("hidden.bs.modal", onImportHiddenOnce);
                    showConflicts();
                };
                isTransientImportModalHide = true;
                importModalEl.addEventListener("hidden.bs.modal", onImportHiddenOnce);
                importModal?.hide();
            } else {
                showConflicts();
            }
        });
    };

    const fillSimpleSelect = (select, items, placeholder = "Seleccione") => {
        if (!select) return;
        const current = String(select.value || "");
        select.innerHTML = `<option value="">${placeholder}</option>`;
        (items || []).forEach((item) => {
            const option = document.createElement("option");
            option.value = String(item.nombre || "");
            option.textContent = String(item.nombre || "");
            select.appendChild(option);
        });
        if (current && Array.from(select.options).some((opt) => opt.value === current)) {
            select.value = current;
        }
    };

    const syncRecepcionUbicacionReadonly = () => {
        if (!ubicacionReadonly) return;
        ubicacionReadonly.value = String(document.getElementById("recepcion-area-trabajo")?.value || "").trim();
    };

    const renderRunChunkPreview = (validation) => {
        if (!previewRunBody) return;
        const rows = Array.isArray(validation?.analyzed_rows) ? validation.analyzed_rows : [];
        if (!rows.length) {
            previewRunBody.innerHTML = '<tr><td colspan="10" class="text-center text-muted">Sin datos de bloque para previsualizar.</td></tr>';
            return;
        }
        const badge = (status) => {
            if (status === "exact") return '<span class="badge text-bg-danger">Igual</span>';
            if (status === "similar") return '<span class="badge text-bg-warning">Similar</span>';
            return '<span class="badge text-bg-success">Nuevo</span>';
        };
        previewRunBody.innerHTML = rows.map((row) => {
            const d = row.data || {};
            const invCode = normalizeCodeToPlaceholder(d.cod_inventario);
            const esbyeCode = normalizeCodeToPlaceholder(d.cod_esbye);
            return `
                <tr>
                    <td>${Number(row.row_index || 0) + 1}</td>
                    <td>${badge(row.status)}</td>
                    <td class="${isNoCodeValue(invCode) ? "code-sc-cell" : ""}">${String(invCode || "-")}</td>
                    <td class="${isNoCodeValue(esbyeCode) ? "code-sc-cell" : ""}">${String(esbyeCode || "-")}</td>
                    <td>${String(d.descripcion || "-")}</td>
                    <td>${String(d.marca || "-")}</td>
                    <td>${String(d.modelo || "-")}</td>
                    <td>${String(d.serie || "-")}</td>
                    <td>${String(d.ubicacion || "-")}</td>
                    <td>${String(d.ubicacion_esbye || "-")}</td>
                </tr>
            `;
        }).join("");
    };

    const loadRecepcionSelectOptions = async () => {
        const requestId = ++recepcionSelectsRequestId;
        try {
            const [estadosRes, cuentasRes, adminsRes] = await Promise.all([
                api.get("/api/parametros/estados"),
                api.get("/api/parametros/cuentas"),
                api.get("/api/administradores"),
            ]);

            // Evita que respuestas viejas pisen los valores del formulario actual.
            if (requestId !== recepcionSelectsRequestId) return;

            fillSimpleSelect(estadoSelect, estadosRes.data || [], "-- Seleccionar estado --");
            fillSimpleSelect(cuentaSelect, cuentasRes.data || [], "-- Seleccionar cuenta --");
            fillSimpleSelect(usuarioFinalSelect, adminsRes.data || [], "-- Seleccionar personal --");
        } catch (_err) {
            // Permite seguir usando el modal aun si no cargan catálogos.
        }
    };

    // Exponer loader para handlers globales de la tabla temporal.
    window.__loadRecepcionSelectOptions = loadRecepcionSelectOptions;

    const showImportStatus = (html, type = "info") => {
        if (!importStatus) return;
        importStatus.className = `alert alert-${type} small mt-3 py-2`;
        importStatus.innerHTML = html;
        importStatus.classList.remove("d-none");
    };

    const clearImportStatus = () => {
        if (!importStatus) return;
        importStatus.classList.add("d-none");
        importStatus.innerHTML = "";
    };

    const resetImportState = () => {
        importState.step = 1;
        importState.sessionId = null;
        importState.sourceFile = null;
        importState.headers = [];
        importState.previewRows = [];
        importState.totalRows = 0;
        importState.mappings = [];
        importState.lockedMainLocationColIdx = null;
        importState.startIndex = 0;
        importState.hasMore = false;
        importState.previewChunkData = null;
        importState.isRecoveringSession = false;
        if (inputImportExcel) inputImportExcel.value = "";
        if (mappingSelectsRow) mappingSelectsRow.innerHTML = "";
        if (mappingHeadersRow) mappingHeadersRow.innerHTML = "";
        if (mappingPreviewBody) mappingPreviewBody.innerHTML = "";
        if (previewInfo) previewInfo.textContent = "";
        if (reviewSummary) reviewSummary.textContent = "Listo para importar el primer bloque (20 filas).";
        if (previewRunBody) {
            previewRunBody.innerHTML = '<tr><td colspan="10" class="text-center text-muted">Sin datos de bloque para previsualizar.</td></tr>';
        }
        clearImportStatus();
        if (btnImportNext) {
            btnImportNext.disabled = true;
            btnImportNext.classList.remove("d-none");
        }
        if (btnImportBack) btnImportBack.disabled = true;
    };

    const setImportStep = (step) => {
        importState.step = step;
        [panel1, panel2, panel3].forEach((p, i) => {
            if (p) p.classList.toggle("d-none", i + 1 !== step);
        });
        stepBadges.forEach((badge, i) => {
            if (!badge) return;
            const n = i + 1;
            const done = n < step;
            const active = n === step;
            badge.className = `badge rounded-pill ${done ? "bg-success" : active ? "bg-primary" : "bg-secondary"}`;
            badge.innerHTML = done ? '<i class="bi bi-check"></i>' : String(n);
        });
        stepLabels.forEach((label, i) => {
            if (!label) return;
            label.className = `small ${i + 1 === step ? "fw-semibold" : "text-muted"}`;
        });
        if (btnImportBack) btnImportBack.classList.toggle("d-none", step === 1);
        if (btnImportNext) btnImportNext.classList.toggle("d-none", step !== 2);
        if (btnImportRun) btnImportRun.classList.toggle("d-none", step !== 3);
        if (btnImportBack) btnImportBack.disabled = step === 1;
    };

    const toMappingObject = () => {
        const mapping = {};
        importState.mappings.forEach((field, idx) => {
            mapping[String(idx)] = field || "";
        });
        return mapping;
    };

    const normalizeLocationMappings = (headers, mappings) => {
        const out = Array.isArray(mappings) ? [...mappings] : [];
        const seenBase = {};

        out.forEach((field, idx) => {
            if (!field || !BASE_TO_ESBYE_FIELD[field]) return;
            const headerNorm = String(Array.isArray(headers) ? headers[idx] : "").toLowerCase();
            const seenCount = seenBase[field] || 0;
            if (headerNorm.includes("esbye") || seenCount >= 1) {
                out[idx] = BASE_TO_ESBYE_FIELD[field];
            }
            seenBase[field] = seenCount + 1;
        });

        const locationIndices = [];
        (headers || []).forEach((header, idx) => {
            const headerNorm = String(header || "").toLowerCase();
            if (headerNorm.includes("ubicacion") || out[idx] === "ubicacion") {
                locationIndices.push(idx);
            }
        });

        if (locationIndices.length) {
            const [firstIdx, secondIdx] = locationIndices;
            if (Number.isInteger(firstIdx)) out[firstIdx] = "ubicacion";
            if (Number.isInteger(secondIdx)) out[secondIdx] = "ubicacion_esbye";
            importState.lockedMainLocationColIdx = Number.isInteger(firstIdx) ? firstIdx : null;
        } else {
            importState.lockedMainLocationColIdx = null;
        }
        return out;
    };

    const applyMappingHighlights = () => {
        if (!mappingSelectsRow || !mappingHeadersRow || !mappingPreviewBody) return;
        const allSelThs = mappingSelectsRow.querySelectorAll("th");
        const allHdrThs = mappingHeadersRow.querySelectorAll("th");
        const allRows = mappingPreviewBody.querySelectorAll("tr");

        importState.mappings.forEach((mapping, colIdx) => {
            const ignored = !mapping;
            if (allSelThs[colIdx]) allSelThs[colIdx].classList.toggle("table-secondary", ignored);
            if (allHdrThs[colIdx]) allHdrThs[colIdx].classList.toggle("table-secondary", ignored);
            allRows.forEach((tr) => {
                if (tr.children[colIdx]) tr.children[colIdx].classList.toggle("table-secondary", ignored);
            });
        });
    };

    const buildMappingSelect = (colIdx) => {
        const sel = document.createElement("select");
        sel.className = "form-select form-select-sm";
        sel.style.fontSize = "0.78rem";
        RECEPCION_CANONICAL_FIELDS.forEach(({ value, label }) => {
            const opt = document.createElement("option");
            opt.value = value;
            opt.textContent = label;
            if (value === importState.mappings[colIdx]) opt.selected = true;
            sel.appendChild(opt);
        });
        const isMainLocationLocked = importState.lockedMainLocationColIdx === colIdx;
        if (isMainLocationLocked) {
            sel.value = "ubicacion";
            sel.disabled = true;
            sel.title = "La ubicación principal se asigna automáticamente para recepción.";
        }
        sel.addEventListener("change", () => {
            importState.mappings[colIdx] = sel.value;
            importState.mappings = normalizeLocationMappings(importState.headers, importState.mappings);
            renderMappingStep();
            applyMappingHighlights();
        });
        return sel;
    };

    const renderMappingStep = () => {
        if (!mappingSelectsRow || !mappingHeadersRow || !mappingPreviewBody) return;
        mappingSelectsRow.innerHTML = "";
        mappingHeadersRow.innerHTML = "";
        mappingPreviewBody.innerHTML = "";

        importState.headers.forEach((header, colIdx) => {
            const thSel = document.createElement("th");
            thSel.style.minWidth = "165px";
            thSel.style.padding = "4px 6px";
            thSel.appendChild(buildMappingSelect(colIdx));
            mappingSelectsRow.appendChild(thSel);

            const thHdr = document.createElement("th");
            thHdr.className = "text-muted small fw-normal";
            thHdr.style.padding = "2px 6px";
            thHdr.textContent = header || `(col ${colIdx + 1})`;
            mappingHeadersRow.appendChild(thHdr);
        });

        importState.previewRows.forEach((row) => {
            const tr = document.createElement("tr");
            importState.headers.forEach((_, colIdx) => {
                const td = document.createElement("td");
                td.className = "small";
                td.style.maxWidth = "200px";
                td.style.overflow = "hidden";
                td.style.textOverflow = "ellipsis";
                td.style.whiteSpace = "nowrap";
                const value = String(row[colIdx] ?? "");
                td.textContent = value;
                td.title = value;
                tr.appendChild(td);
            });
            mappingPreviewBody.appendChild(tr);
        });

        if (previewInfo) {
            previewInfo.textContent = `Vista previa: ${importState.previewRows.length} de ${importState.totalRows} filas.`;
        }
        applyMappingHighlights();
    };

    const renderRecepcionColumnSelector = () => {
        if (!recepcionColumnSelector) return;
        recepcionColumnSelector.innerHTML = "";
        const selected = new Set(normalizeRecepcionColumnIds(recepcionSelectedColumnIds));

        RECEPCION_BIENES_COLUMNS.forEach((col) => {
            const row = document.createElement("label");
            row.className = "list-group-item list-group-item-action d-flex align-items-center py-2";
            row.innerHTML = `
                <input class="form-check-input me-3 recepcion-col-chk" type="checkbox" value="${col.id}" ${selected.has(col.id) ? "checked" : ""}>
                <span class="small fw-medium">${col.label}</span>
            `;
            recepcionColumnSelector.appendChild(row);
        });
    };

    const updateImportProgress = () => {
        const endIndex = Math.min(importState.startIndex + importState.chunkSize, importState.totalRows);
        if (chunkRange) {
            if (importState.totalRows > 0) {
                chunkRange.textContent = `${importState.startIndex + 1}-${endIndex} de ${importState.totalRows}`;
            } else {
                chunkRange.textContent = "0-0 de 0";
            }
        }
        if (rowCount) {
            rowCount.textContent = String(recepcionBienesTemp.length);
        }
    };

    const handleImportFile = async (file, options = {}) => {
        const reuseMappings = Boolean(options.reuseMappings);
        const silent = Boolean(options.silent);
        const keepCurrentStep = Boolean(options.keepCurrentStep);
        if (!file) return;
        if (!String(file.name || "").toLowerCase().endsWith(".xlsx")) {
            if (!silent) showImportStatus('<i class="bi bi-x-circle me-1"></i>Solo se aceptan archivos .xlsx.', "danger");
            return false;
        }

        importState.sourceFile = file;
        if (!silent) {
            showImportStatus(`<span class="spinner-border spinner-border-sm me-2" role="status"></span>Procesando <strong>${String(file.name || "archivo")}</strong>...`);
        }
        const formData = new FormData();
        formData.append("file", file);

        try {
            const pre = await fetch("/api/inventario/previsualizar-excel", { method: "POST", body: formData });
            const preData = await pre.json();
            if (!pre.ok || preData.error) {
                if (!silent) showImportStatus(`<i class="bi bi-x-circle me-1"></i>${preData.error || "No se pudo leer el Excel."}`, "danger");
                return false;
            }

            const previousMappings = Array.isArray(importState.mappings) ? [...importState.mappings] : [];
            importState.sessionId = preData.session_id;
            importState.headers = Array.isArray(preData.headers) ? preData.headers : [];
            importState.previewRows = Array.isArray(preData.preview_rows) ? preData.preview_rows : [];
            importState.totalRows = Number(preData.total_rows || 0);
            importState.startIndex = 0;
            importState.hasMore = importState.totalRows > 0;
            const suggested = Array.isArray(preData.suggested_mapping) ? preData.suggested_mapping : [];
            const rawMappings = importState.headers.map((_, idx) => String(suggested[idx] || ""));
            if (reuseMappings && previousMappings.length === importState.headers.length && previousMappings.some(Boolean)) {
                importState.mappings = normalizeLocationMappings(importState.headers, previousMappings);
            } else {
                importState.mappings = normalizeLocationMappings(importState.headers, rawMappings);
            }

            renderMappingStep();
            updateImportProgress();
            if (!silent) clearImportStatus();
            if (btnImportNext) btnImportNext.disabled = false;
            if (!keepCurrentStep) setImportStep(2);
            return true;
        } catch (_err) {
            if (!silent) showImportStatus('<i class="bi bi-x-circle me-1"></i>Error de red al subir archivo.', "danger");
            return false;
        }
    };

    const recoverRecepcionImportSession = async () => {
        if (importState.isRecoveringSession) return false;
        if (!importState.sourceFile) {
            showImportStatus('<i class="bi bi-x-circle me-1"></i>La sesión expiró y no hay archivo en memoria para reconectar. Selecciona el Excel nuevamente.', "danger");
            return false;
        }

        importState.isRecoveringSession = true;
        const keepStep = importState.step;
        const keepStartIndex = importState.startIndex;
        try {
            showImportStatus('<span class="spinner-border spinner-border-sm me-2" role="status"></span>Reconectando sesión de importación...', "warning");
            const restored = await handleImportFile(importState.sourceFile, {
                reuseMappings: true,
                silent: true,
                keepCurrentStep: true,
            });
            if (!restored) return false;

            importState.startIndex = keepStartIndex;
            setImportStep(keepStep);
            updateImportProgress();
            showImportStatus('<i class="bi bi-check-circle-fill me-1"></i>Sesión restablecida. Reintentando...', "success");
            return true;
        } finally {
            importState.isRecoveringSession = false;
        }
    };

    const sendRecepcionImportRequest = async (payload, allowRecover = true) => {
        try {
            return await api.send("/api/inventario/excel-a-filas-recepcion", "POST", payload);
        } catch (error) {
            if (allowRecover && error?.status === 410) {
                const recovered = await recoverRecepcionImportSession();
                if (recovered) {
                    return await api.send("/api/inventario/excel-a-filas-recepcion", "POST", {
                        ...payload,
                        session_id: importState.sessionId,
                    });
                }
            }
            throw error;
        }
    };

    const importNextChunk = async () => {
        if (!importState.sessionId) {
            notify("Primero selecciona y procesa un archivo Excel.", true);
            return;
        }
        syncRecepcionUbicacionReadonly();
        const forcedLocation = String(ubicacionReadonly?.value || "").trim();
        const forcedAreaId = String(document.getElementById("recepcion-ubicacion-area-id")?.value || "").trim();
        if (!forcedLocation || !forcedAreaId) {
            notify("Seleccione Bloque, Piso y Área en el acta antes de importar bienes.", true);
            return;
        }

        btnImportRun?.setAttribute("disabled", "disabled");
        showImportStatus("<span class=\"spinner-border spinner-border-sm me-2\" role=\"status\"></span>Verificando bloque...", "info");

        try {
            const payloadBase = {
                session_id: importState.sessionId,
                mapping: toMappingObject(),
                forced_location: forcedLocation,
                forced_area_id: Number(forcedAreaId),
                start_index: importState.startIndex,
                chunk_size: importState.chunkSize,
            };

            const validation = importState.previewChunkData && Number(importState.previewChunkData.start_index) === Number(importState.startIndex)
                ? importState.previewChunkData
                : await sendRecepcionImportRequest({
                    ...payloadBase,
                    validate_only: true,
                });

            const validationSummary = validation.summary || {};
            const exactCount = Number(validationSummary.exact || 0);
            const similarCount = Number(validationSummary.similar || 0);
            const normalCount = Number(validationSummary.normal || 0);
            if (reviewSummary) {
                reviewSummary.innerHTML = `<strong>Nuevos:</strong> ${normalCount} &nbsp; <strong class="text-danger">Iguales:</strong> ${exactCount} &nbsp; <strong class="text-warning-emphasis">Similares:</strong> ${similarCount}`;
            }

            let forceDuplicate = false;
            if (exactCount > 0 || similarCount > 0) {
                const proceed = await openImportConflictsModal(validation);
                if (!proceed) {
                    showImportStatus("Bloque cancelado por el usuario. Puedes ajustar el mapeo y volver a intentar.", "warning");
                    return;
                }
                forceDuplicate = true;
            }

            showImportStatus("<span class=\"spinner-border spinner-border-sm me-2\" role=\"status\"></span>Importando bloque...", "info");
            const res = await sendRecepcionImportRequest({
                ...payloadBase,
                force_duplicate: forceDuplicate,
            });
            importState.previewChunkData = null;

            const importedRows = Array.isArray(res.rows) ? res.rows : [];
            importedRows.forEach((row) => {
                recepcionBienesTemp.push({
                    ...normalizeRecepcionRowLikeInventarioImport(row, forcedLocation),
                });
            });

            importState.hasMore = Boolean(res.has_more);
            importState.startIndex = Number(res.next_start_index || 0);
            renderRecepcionBienesTable();
            updateRecepcionSummary();
            updateImportProgress();
            document.dispatchEvent(new CustomEvent("informe:tablaExtraida"));

            if (reviewSummary) {
                const summary = res.summary || {};
                reviewSummary.innerHTML = importState.hasMore
                    ? `Se agregaron ${importedRows.length} filas. <strong>Nuevos:</strong> ${Number(summary.normal || 0)} · <strong class="text-danger">Iguales:</strong> ${Number(summary.exact || 0)} · <strong class="text-warning-emphasis">Similares:</strong> ${Number(summary.similar || 0)}. Continúa con el siguiente bloque.`
                    : `Importación completada. Último bloque: ${importedRows.length} filas. <strong>Nuevos:</strong> ${Number(summary.normal || 0)} · <strong class="text-danger">Iguales:</strong> ${Number(summary.exact || 0)} · <strong class="text-warning-emphasis">Similares:</strong> ${Number(summary.similar || 0)}.`;
            }
            if (!importState.hasMore) {
                if (btnImportRun) btnImportRun.textContent = "Importación finalizada";
                btnImportRun?.setAttribute("disabled", "disabled");
                clearImportStatus();
                notify("Importación de recepción completada.");
            } else {
                if (btnImportRun) btnImportRun.innerHTML = '<i class="bi bi-box-arrow-in-down-right me-1"></i>Importar siguiente bloque (20)';
                showImportStatus(`Bloque importado: ${importedRows.length} fila(s).`, "success");
                const nextPreview = await sendRecepcionImportRequest({
                    ...payloadBase,
                    start_index: importState.startIndex,
                    validate_only: true,
                });
                importState.previewChunkData = nextPreview;
                renderRunChunkPreview(nextPreview);
            }
        } catch (error) {
            notify(error.message || "Error al importar bloque de recepción.", true);
        } finally {
            if (importState.hasMore) btnImportRun?.removeAttribute("disabled");
        }
    };

    btnOpen.addEventListener("click", (event) => {
        event.preventDefault();
        syncRecepcionUbicacionReadonly();
        loadRecepcionSelectOptions();
        recepcionTablePage = 1;
        renderRecepcionBienesTable();
        modal.show();
    });

    if (recepcionPagePrev && recepcionPagePrev.dataset.bound !== "1") {
        recepcionPagePrev.dataset.bound = "1";
        recepcionPagePrev.addEventListener("click", () => {
            if (recepcionTablePage <= 1) return;
            recepcionTablePage -= 1;
            renderRecepcionBienesTable();
        });
    }

    if (recepcionPageNext && recepcionPageNext.dataset.bound !== "1") {
        recepcionPageNext.dataset.bound = "1";
        recepcionPageNext.addEventListener("click", () => {
            recepcionTablePage += 1;
            renderRecepcionBienesTable();
        });
    }

    if (recepcionPageSize && recepcionPageSize.dataset.bound !== "1") {
        recepcionPageSize.dataset.bound = "1";
        recepcionPageSize.addEventListener("change", () => {
            const parsed = Number(recepcionPageSize.value || 10);
            recepcionTablePerPage = Number.isFinite(parsed) && parsed > 0 ? parsed : 10;
            recepcionTablePage = 1;
            renderRecepcionBienesTable();
        });
    }

    btnImportExcel?.addEventListener("click", () => {
        syncRecepcionUbicacionReadonly();
        const forcedLocation = String(ubicacionReadonly?.value || "").trim();
        const forcedAreaId = String(document.getElementById("recepcion-ubicacion-area-id")?.value || "").trim();
        if (!forcedLocation || !forcedAreaId) {
            notify("Seleccione Bloque, Piso y Área en el acta antes de importar bienes.", true);
            return;
        }
        resetImportState();
        setImportStep(1);
        if (btnImportRun) {
            btnImportRun.innerHTML = '<i class="bi bi-box-arrow-in-down-right me-1"></i>Importar bloque (20)';
            btnImportRun.removeAttribute("disabled");
        }
        shouldReturnToRecepcionModal = true;
        pendingChildModal = "import";
        modal.hide();
    });

    inputImportExcel?.addEventListener("change", async () => {
        const file = inputImportExcel.files?.[0];
        if (!file) return;
        await handleImportFile(file);
    });

    btnImportNext?.addEventListener("click", () => {
        if (!importState.sessionId) {
            notify("Primero selecciona un archivo Excel.", true);
            return;
        }
        const selectedFields = importState.mappings.filter(Boolean);
        if (!selectedFields.length) {
            notify("Asigna al menos una columna antes de continuar.", true);
            return;
        }
        btnImportNext.setAttribute("disabled", "disabled");
        sendRecepcionImportRequest({
            session_id: importState.sessionId,
            mapping: toMappingObject(),
            forced_location: String(ubicacionReadonly?.value || "").trim(),
            forced_area_id: Number(document.getElementById("recepcion-ubicacion-area-id")?.value || 0),
            start_index: importState.startIndex,
            chunk_size: importState.chunkSize,
            validate_only: true,
        }).then((validation) => {
            importState.previewChunkData = validation;
            renderRunChunkPreview(validation);
            const s = validation.summary || {};
            if (reviewSummary) {
                reviewSummary.innerHTML = `<strong>Nuevos:</strong> ${Number(s.normal || 0)} &nbsp; <strong class="text-danger">Iguales:</strong> ${Number(s.exact || 0)} &nbsp; <strong class="text-warning-emphasis">Similares:</strong> ${Number(s.similar || 0)}`;
            }
            setImportStep(3);
            updateImportProgress();
        }).catch((error) => {
            notify(error.message || "No se pudo previsualizar el bloque actual.", true);
        }).finally(() => {
            btnImportNext.removeAttribute("disabled");
        });
    });

    btnImportBack?.addEventListener("click", () => {
        if (importState.step === 3) {
            setImportStep(2);
            return;
        }
        if (importState.step === 2) {
            setImportStep(1);
        }
    });

    btnImportRun?.addEventListener("click", async () => {
        await importNextChunk();
    });

    importModalEl?.addEventListener("hidden.bs.modal", () => {
        if (isTransientImportModalHide) {
            isTransientImportModalHide = false;
            return;
        }
        resetImportState();
        if (shouldReturnToRecepcionModal) {
            shouldReturnToRecepcionModal = false;
            modal.show();
        }
    });

    btnColumnas?.addEventListener("click", () => {
        renderRecepcionColumnSelector();
        shouldReturnFromColumnsModal = true;
        pendingChildModal = "columns";
        modal.hide();
    });

    btnColumnasGuardar?.addEventListener("click", () => {
        const checked = Array.from(document.querySelectorAll(".recepcion-col-chk:checked")).map((node) => String(node.value || "").trim());
        if (!checked.length) {
            notify("Seleccione al menos una columna para mostrar en recepción.", true);
            return;
        }
        recepcionSelectedColumnIds = normalizeRecepcionColumnIds(checked);
        saveRecepcionColumnPreferences();
        renderRecepcionBienesTable();
        updateRecepcionSummary();
        columnsModal?.hide();
    });

    columnsModalEl?.addEventListener("hidden.bs.modal", () => {
        if (shouldReturnFromColumnsModal) {
            shouldReturnFromColumnsModal = false;
            modal.show();
        }
    });

    btnCancel?.addEventListener("click", () => {
        clearRecepcionBienForm();
        setRecepcionEditMode(-1);
    });

    btnClear?.addEventListener("click", () => {
        clearRecepcionBienForm();
        setRecepcionEditMode(-1);
        syncRecepcionUbicacionReadonly();
    });

    btnSave.addEventListener("click", async () => {
        const payload = getRecepcionBienFormValues();
        if (!payload.descripcion) {
            notify("La descripción del bien es obligatoria.", true);
            return;
        }
        if (!payload.cantidad || payload.cantidad <= 0) {
            notify("La cantidad debe ser un entero mayor que 0.", true);
            return;
        }

        const duplicates = await checkRecepcionDuplicates(payload);
        if (duplicates.length) {
            const confirmed = await openDuplicateModal({ duplicates, payload, mode: "create" });
            if (!confirmed) return;
        }

        if (recepcionEditIndex >= 0 && recepcionBienesTemp[recepcionEditIndex]) {
            recepcionBienesTemp[recepcionEditIndex] = payload;
            notify("Bien actualizado en la lista temporal.");
        } else {
            recepcionBienesTemp.push(payload);
            notify("Bien agregado a la lista temporal.");
        }
        clearRecepcionBienForm();
        setRecepcionEditMode(-1);
        renderRecepcionBienesTable();
        updateRecepcionSummary();
        document.dispatchEvent(new CustomEvent("informe:tablaExtraida"));
    });

    btnConfirm.addEventListener("click", () => {
        updateRecepcionSummary();
        modal.hide();
        document.dispatchEvent(new CustomEvent("informe:tablaExtraida"));
    });

    modalEl.addEventListener("hidden.bs.modal", () => {
        if (pendingChildModal === "import") {
            pendingChildModal = "";
            importModal?.show();
            setTimeout(() => inputImportExcel?.focus(), 0);
            return;
        }
        if (pendingChildModal === "columns") {
            pendingChildModal = "";
            columnsModal?.show();
            return;
        }
        updateRecepcionSummary();
    });

    document.getElementById("recepcion-area")?.addEventListener("change", syncRecepcionUbicacionReadonly);
    document.getElementById("recepcion-piso")?.addEventListener("change", syncRecepcionUbicacionReadonly);
    document.getElementById("recepcion-bloque")?.addEventListener("change", syncRecepcionUbicacionReadonly);
}

async function loadBajasEstadoOptions() {
    if (Array.isArray(bajasEstadoOptions) && bajasEstadoOptions.length) return;
    try {
        const response = await api.get("/api/parametros/estados");
        const values = Array.isArray(response?.data)
            ? response.data.map((entry) => String(entry?.nombre || "").trim()).filter(Boolean)
            : [];
        if (!values.some((value) => String(value).toLowerCase() === "malo")) {
            values.unshift("MALO");
        }
        bajasEstadoOptions = values;
    } catch (_err) {
        bajasEstadoOptions = ["MALO"];
    }
}

function setBajasStep(step) {
    const normalizedStep = Number(step) === 2 ? 2 : 1;
    bajasStep = normalizedStep;
    const step1 = document.getElementById("bajas-step-1");
    const step2 = document.getElementById("bajas-step-2");
    const badge1 = document.getElementById("bajas-step-badge-1");
    const badge2 = document.getElementById("bajas-step-badge-2");
    const label1 = document.getElementById("bajas-step-label-1");
    const label2 = document.getElementById("bajas-step-label-2");
    const btnPrev = document.getElementById("btn-bajas-atras");
    const btnNext = document.getElementById("btn-bajas-siguiente");
    const btnConfirm = document.getElementById("btn-bajas-confirmar");
    const footerInfo = document.getElementById("bajas-modal-footer-info");

    if (step1) step1.classList.toggle("d-none", normalizedStep !== 1);
    if (step2) step2.classList.toggle("d-none", normalizedStep !== 2);

    if (badge1) badge1.className = `badge ${normalizedStep === 1 ? "bg-primary" : "bg-success"}`;
    if (badge2) badge2.className = `badge ${normalizedStep === 2 ? "bg-primary" : "bg-secondary"}`;
    if (label1) label1.className = normalizedStep === 1 ? "fw-semibold" : "text-success";
    if (label2) label2.className = normalizedStep === 2 ? "fw-semibold" : "text-muted";
    if (btnPrev) btnPrev.disabled = normalizedStep === 1;
    if (btnNext) btnNext.classList.toggle("d-none", normalizedStep !== 1);
    if (btnConfirm) btnConfirm.classList.toggle("d-none", normalizedStep !== 2);
    if (footerInfo) footerInfo.textContent = `Paso ${normalizedStep} de 2`;
}

function applyBajasSelectionFilter() {
    const needle = String(document.getElementById("bajas-buscar")?.value || "").trim().toLowerCase();
    if (!needle) {
        bajasFilteredItems = Array.isArray(inventoryDataCache) ? [...inventoryDataCache] : [];
        return;
    }

    bajasFilteredItems = (Array.isArray(inventoryDataCache) ? inventoryDataCache : []).filter((item) => {
        const bag = [
            item?.item_numero,
            item?.cod_inventario,
            item?.cod_esbye,
            item?.descripcion,
            item?.marca,
            item?.modelo,
            item?.serie,
            item?.estado,
            item?.ubicacion,
        ]
            .map((value) => String(value || "").toLowerCase())
            .join(" ");
        return bag.includes(needle);
    });
}

function renderBajasSelectionTable() {
    const tbody = document.getElementById("tbody-bajas-seleccion");
    const selectedInfo = document.getElementById("bajas-seleccionados-info");
    const checkAll = document.getElementById("bajas-check-all");
    if (!tbody) return;

    const rows = Array.isArray(bajasFilteredItems) ? bajasFilteredItems : [];
    if (!rows.length) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center text-muted">No hay bienes para mostrar.</td></tr>';
        if (selectedInfo) selectedInfo.textContent = `${bajasSelectedItemIds.size} seleccionados`;
        if (checkAll) checkAll.checked = false;
        return;
    }

    tbody.innerHTML = "";
    rows.forEach((item) => {
        const itemId = Number(item?.id || 0);
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td><input type="checkbox" class="form-check-input bajas-item-check" data-id="${itemId}" ${bajasSelectedItemIds.has(itemId) ? "checked" : ""}></td>
            <td>${String(item?.item_numero || "-")}</td>
            <td>${String(item?.cod_inventario || "")}</td>
            <td>${String(item?.cod_esbye || "")}</td>
            <td>${String(item?.descripcion || "")}</td>
            <td>${String(item?.estado || "")}</td>
            <td>${String(item?.ubicacion || "")}</td>
        `;
        tbody.appendChild(tr);
    });

    tbody.querySelectorAll(".bajas-item-check").forEach((chk) => {
        chk.addEventListener("change", (event) => {
            const itemId = Number(event.target?.dataset?.id || 0);
            if (!itemId) return;
            if (event.target.checked) {
                bajasSelectedItemIds.add(itemId);
            } else {
                bajasSelectedItemIds.delete(itemId);
            }
            if (selectedInfo) selectedInfo.textContent = `${bajasSelectedItemIds.size} seleccionados`;
        });
    });

    if (selectedInfo) selectedInfo.textContent = `${bajasSelectedItemIds.size} seleccionados`;
    if (checkAll) {
        checkAll.checked = rows.every((row) => bajasSelectedItemIds.has(Number(row?.id || 0)));
    }
}

function renderBajasEditionTable() {
    const thead = document.getElementById("thead-bajas-edicion");
    const tbody = document.getElementById("tbody-bajas-edicion");
    if (!tbody || !thead) return;

    const selectedIds = normalizeBajasColumnIds(bajasSelectedColumnIds);
    const selectedCols = selectedIds
        .map((id) => BAJAS_BIENES_COLUMNS.find((col) => col.id === id))
        .filter(Boolean);

    thead.innerHTML = "";
    const trHead = document.createElement("tr");
    selectedCols.forEach((col) => {
        const th = document.createElement("th");
        th.textContent = col.label === "ESTADO" ? "Estado nuevo" : col.label;
        if (col.id === "estado") th.style.minWidth = "160px";
        if (col.id === "justificacion") th.style.minWidth = "260px";
        trHead.appendChild(th);
    });
    thead.appendChild(trHead);

    if (!Array.isArray(bajasDraftBienes) || !bajasDraftBienes.length) {
        tbody.innerHTML = `<tr><td colspan="${Math.max(1, selectedCols.length)}" class="text-center text-muted">No hay bienes seleccionados.</td></tr>`;
        return;
    }

    const estados = Array.isArray(bajasEstadoOptions) && bajasEstadoOptions.length
        ? bajasEstadoOptions
        : ["MALO"];

    tbody.innerHTML = "";
    bajasDraftBienes.forEach((row, idx) => {
        const tr = document.createElement("tr");
        const optionsHtml = estados
            .map((estado) => `<option value="${estado}" ${String(row?.estado || "").trim() === estado ? "selected" : ""}>${estado}</option>`)
            .join("");

        const cells = selectedCols.map((col) => {
            if (col.id === "estado") {
                return `
                    <td>
                        <select class="form-select form-select-sm bajas-edit-estado" data-index="${idx}">
                            ${optionsHtml}
                        </select>
                    </td>
                `;
            }
            if (col.id === "justificacion") {
                return `
                    <td>
                        <textarea class="form-control form-control-sm bajas-edit-justificacion" rows="2" data-index="${idx}" placeholder="Justificación de la baja...">${String(row?.justificacion || "")}</textarea>
                    </td>
                `;
            }
            return `<td>${String(row?.[col.id] || (col.id === "item_numero" ? "-" : ""))}</td>`;
        });
        tr.innerHTML = cells.join("");
        tbody.appendChild(tr);
    });

    tbody.querySelectorAll(".bajas-edit-estado").forEach((select) => {
        select.addEventListener("change", (event) => {
            const idx = Number(event.target?.dataset?.index || -1);
            if (idx < 0 || !bajasDraftBienes[idx]) return;
            bajasDraftBienes[idx].estado = String(event.target.value || "").trim() || "MALO";
        });
    });

    tbody.querySelectorAll(".bajas-edit-justificacion").forEach((textarea) => {
        textarea.addEventListener("input", (event) => {
            const idx = Number(event.target?.dataset?.index || -1);
            if (idx < 0 || !bajasDraftBienes[idx]) return;
            bajasDraftBienes[idx].justificacion = String(event.target.value || "").trim();
        });
    });
}

function renderBajasColumnSelector() {
    const container = document.getElementById("bajas-column-selector");
    if (!container) return;

    const selected = new Set(normalizeBajasColumnIds(bajasSelectedColumnIds));
    container.innerHTML = "";

    BAJAS_BIENES_COLUMNS.forEach((col) => {
        const isRequired = BAJAS_REQUIRED_COLUMN_IDS.includes(col.id);
        const item = document.createElement("label");
        item.className = "list-group-item d-flex align-items-center gap-2";
        item.innerHTML = `
            <input class="form-check-input me-1" type="checkbox" value="${col.id}" ${selected.has(col.id) ? "checked" : ""} ${isRequired ? "disabled" : ""}>
            <span>${col.label}</span>
            ${isRequired ? '<span class="badge text-bg-secondary ms-auto">Obligatoria</span>' : ""}
        `;
        container.appendChild(item);
    });
}

function setupBajasBienesModal() {
    const modalEl = document.getElementById("modalBajasBienes");
    if (!modalEl) return;
    const modal = bootstrap.Modal.getOrCreateInstance(modalEl);

    const btnOpen = document.querySelector("#sec-bajas .btn-bajas-seleccionar");
    const btnBuscar = document.getElementById("btn-bajas-buscar");
    const btnPrev = document.getElementById("btn-bajas-atras");
    const btnNext = document.getElementById("btn-bajas-siguiente");
    const btnConfirm = document.getElementById("btn-bajas-confirmar");
    const checkAll = document.getElementById("bajas-check-all");
    const btnColumnas = document.getElementById("btn-bajas-columnas");
    const btnColumnasGuardar = document.getElementById("btn-bajas-columnas-guardar");
    const modalColumnasEl = document.getElementById("modalBajasColumnas");
    const modalColumnas = modalColumnasEl ? bootstrap.Modal.getOrCreateInstance(modalColumnasEl) : null;

    const refreshSelection = () => {
        applyBajasSelectionFilter();
        renderBajasSelectionTable();
    };

    btnOpen?.addEventListener("click", async (event) => {
        event.preventDefault();
        await loadExtraccionData();
        await loadBajasEstadoOptions();
        bajasSelectedItemIds = new Set((Array.isArray(bajasBienesTemp) ? bajasBienesTemp : []).map((row) => Number(row?.id || 0)).filter(Boolean));
        bajasDraftBienes = [];
        setBajasStep(1);
        refreshSelection();
        modal.show();
    });

    btnBuscar?.addEventListener("click", refreshSelection);
    document.getElementById("bajas-buscar")?.addEventListener("input", refreshSelection);

    checkAll?.addEventListener("change", (event) => {
        const shouldCheck = Boolean(event.target.checked);
        (Array.isArray(bajasFilteredItems) ? bajasFilteredItems : []).forEach((item) => {
            const itemId = Number(item?.id || 0);
            if (!itemId) return;
            if (shouldCheck) {
                bajasSelectedItemIds.add(itemId);
            } else {
                bajasSelectedItemIds.delete(itemId);
            }
        });
        renderBajasSelectionTable();
    });

    btnPrev?.addEventListener("click", () => {
        setBajasStep(1);
    });

    btnNext?.addEventListener("click", () => {
        if (!bajasSelectedItemIds.size) {
            notify("Seleccione al menos un bien para continuar.", true);
            return;
        }

        const selectedRows = (Array.isArray(inventoryDataCache) ? inventoryDataCache : [])
            .filter((item) => bajasSelectedItemIds.has(Number(item?.id || 0)))
            .map((item) => ({
                ...item,
                id: Number(item?.id || 0),
                item_numero: item?.item_numero,
                cod_inventario: String(item?.cod_inventario || ""),
                cod_esbye: String(item?.cod_esbye || ""),
                descripcion: String(item?.descripcion || ""),
                estado: "MALO",
                justificacion: String(item?.justificacion || ""),
                procedencia: String(item?.procedencia || ""),
            }));

        if (!selectedRows.length) {
            notify("No se pudo preparar la lista de bienes seleccionados.", true);
            return;
        }

        bajasDraftBienes = selectedRows;
        renderBajasEditionTable();
        setBajasStep(2);
    });

    btnColumnas?.addEventListener("click", () => {
        renderBajasColumnSelector();
        modalColumnas?.show();
    });

    btnColumnasGuardar?.addEventListener("click", async () => {
        const checks = Array.from(document.querySelectorAll("#bajas-column-selector input[type='checkbox']:checked"));
        const selected = checks.map((node) => String(node.value || "").trim()).filter(Boolean);
        bajasSelectedColumnIds = normalizeBajasColumnIds(selected);
        await saveBajasColumnPreferences();
        renderBajasEditionTable();
        modalColumnas?.hide();
    });

    btnConfirm?.addEventListener("click", () => {
        if (!Array.isArray(bajasDraftBienes) || !bajasDraftBienes.length) {
            notify("No hay bienes seleccionados para el acta de baja.", true);
            return;
        }
        bajasBienesTemp = bajasDraftBienes.map((row) => ({
            ...row,
            estado: String(row?.estado || "").trim() || "MALO",
            justificacion: String(row?.justificacion || "").trim(),
        }));
        updateBajasSummary();
        modal.hide();
        document.dispatchEvent(new CustomEvent("informe:tablaExtraida"));
    });

    modalEl.addEventListener("hidden.bs.modal", () => {
        setBajasStep(1);
        bajasDraftBienes = [];
    });
}

function setupBajasRegistradosModal() {
    const btnOpen = document.querySelector("#sec-bajas .btn-bajas-registrados");
    if (!btnOpen) return;

    btnOpen.addEventListener("click", () => {
        window.location.href = "/api/inventario/bajas/export";
    });
}

function fillActaFormFromData(tipo, formularioData) {
    const tabId = mapTipoToTabId(tipo);
    document.getElementById(tabId)?.click();

    const activeTipo = tabId.replace("tab-", "");
    const form = document.getElementById(`form-${activeTipo}`);
    if (!form || !formularioData || typeof formularioData !== "object") return;

    const aliases = {
        usuario_final: ["recibido_por"],
        recibido_por: ["usuario_final"],
        entregado: ["entregado_por"],
    };

    Object.entries(formularioData).forEach(([name, value]) => {
        const normalizedValue = value == null ? "" : String(value);
        const candidates = [name].concat(aliases[name] || []);

        for (const candidate of candidates) {
            const input = form.querySelector(`[name="${CSS.escape(candidate)}"]`);
            if (input) {
                input.value = normalizedValue;
                return;
            }
        }

        // Fallback por id si el campo no tiene atributo name.
        const idFallback = form.querySelector(`#${CSS.escape(`entrega-${name.replaceAll("_", "-")}`)}`);
        if (idFallback) {
            idFallback.value = normalizedValue;
        }
    });
}

function applyHistorialPayloadToEditor(record) {
    const raw = record && record.datos_json ? String(record.datos_json) : "";
    if (!raw) {
        notify("Esta acta no contiene datos para edición.", true);
        return;
    }

    let parsed;
    try {
        parsed = JSON.parse(raw);
    } catch (_err) {
        notify("No se pudieron leer los datos de esta acta.", true);
        return;
    }

    fillActaFormFromData(record.tipo_acta, parsed.formulario || {});
    window._globalSelectedTableRows = Array.isArray(parsed.tabla) ? parsed.tabla : [];
    window._globalSelectedColumns = Array.isArray(parsed.columnas) ? parsed.columnas : [];
    if (normalizeTipoActa(record.tipo_acta) === "recepcion") {
        recepcionBienesTemp = Array.isArray(parsed.tabla) ? parsed.tabla.map((row) => ({ ...row })) : [];
        recepcionSelectedColumnIds = normalizeRecepcionColumnIds(parsed.columnas || []);
        renderRecepcionBienesTable();
        updateRecepcionSummary();
    }
    if (normalizeTipoActa(record.tipo_acta) === "bajas" || normalizeTipoActa(record.tipo_acta) === "baja") {
        bajasBienesTemp = Array.isArray(parsed.tabla) ? parsed.tabla.map((row) => ({
            ...row,
            estado: String(row?.estado || "").trim() || "MALO",
            justificacion: String(row?.justificacion || "").trim(),
        })) : [];
        bajasSelectedColumnIds = normalizeBajasColumnIds(parsed.columnas || []);
        updateBajasSummary();
    }
    activeHistorialTemplateSnapshotPath =
        (record && record.plantilla_snapshot_path) ||
        (parsed && parsed.plantilla && parsed.plantilla.snapshot_path) ||
        null;
    {
        const parsedId = Number(record && record.id);
        activeEditingActaId = Number.isInteger(parsedId) && parsedId > 0 ? parsedId : null;
        activeEditingActaTipo = normalizeTipoActa(record && record.tipo_acta);
    }

    const activeTarget = document.querySelector(".settings-menu-btn.active")?.getAttribute("data-target");
    const targetDiv = activeTarget ? document.querySelector(`#${activeTarget} .resultado-tabla-container`) : null;
    if (targetDiv) {
        const rows = window._globalSelectedTableRows.length;
        const cols = window._globalSelectedColumns.length;
        if (rows > 0 && cols > 0) {
            showResultadoTabla(
                targetDiv,
                `<div class="alert alert-info d-inline-block p-2 px-4 shadow-sm mb-0"><i class="bi bi-pencil-square me-2"></i>Acta cargada para edición: <strong>${rows}</strong> filas y <strong>${cols}</strong> columnas.</div>`
            );
        } else {
            hideResultadoTabla(targetDiv);
        }
    }

    if (typeof window.queueInformePreview === "function") {
        window.queueInformePreview(50);
    }
    document.dispatchEvent(new CustomEvent("informe:tablaExtraida"));

    notify("Acta cargada en el formulario para edición.");
}

function resolveHistorialDownloadPath(item) {
    if (item && item.pdf_path) return item.pdf_path;
    if (item && item.docx_path) return item.docx_path;
    return null;
}

async function cargarHistorialActas() {
    const tbody = document.getElementById("lista-historial-actas");
    if (!tbody) return;

    const tipo = normalizeTipoActa(document.querySelector(".settings-menu-btn.active")?.id?.replace("tab-", ""));
    tbody.innerHTML = '<tr><td colspan="3" class="text-center text-muted">Cargando historial...</td></tr>';

    try {
        const response = await fetch(`/api/historial?tipo_acta=${encodeURIComponent(tipo)}`);
        const payload = await response.json();
        const rows = payload && payload.success && Array.isArray(payload.data) ? payload.data : [];

        if (!rows.length) {
            tbody.innerHTML = '<tr><td colspan="3" class="text-center text-muted">No hay actas registradas.</td></tr>';
            return;
        }

        tbody.innerHTML = "";
        rows.forEach((item, idx) => {
            const tr = document.createElement("tr");
            const downloadPath = resolveHistorialDownloadPath(item);

            const tdN = document.createElement("td");
            tdN.className = "fw-bold";
            tdN.textContent = String(idx + 1);

            const tdNumeroActa = document.createElement("td");
            tdNumeroActa.textContent = String(item?.numero_acta || "-");

            const tdActions = document.createElement("td");
            tdActions.className = "text-end";

            const btnDownload = document.createElement("button");
            btnDownload.className = "btn btn-sm btn-outline-primary me-2";
            btnDownload.innerHTML = '<i class="bi bi-download me-1"></i>Descargar';
            btnDownload.disabled = !downloadPath;
            btnDownload.addEventListener("click", () => {
                if (!downloadPath) return;
                const a = document.createElement("a");
                a.href = `/api/descargar?path=${encodeURIComponent(downloadPath)}`;
                a.style.display = "none";
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            });

            const btnEdit = document.createElement("button");
            btnEdit.className = "btn btn-sm btn-outline-secondary";
            btnEdit.innerHTML = '<i class="bi bi-pencil-square me-1"></i>Editar acta';
            btnEdit.addEventListener("click", () => {
                applyHistorialPayloadToEditor(item);
                const modalEl = document.getElementById("modalHistorialActas");
                bootstrap.Modal.getInstance(modalEl)?.hide();
            });

            const btnDelete = document.createElement("button");
            btnDelete.className = "btn btn-sm btn-outline-danger ms-2";
            btnDelete.innerHTML = '<i class="bi bi-trash me-1"></i>Borrar';
            btnDelete.addEventListener("click", async () => {
                const numeroActa = String(item?.numero_acta || "").trim() || `#${item?.id || ""}`;
                const ok = window.confirm(`Se eliminará el acta ${numeroActa} del historial. ¿Desea continuar?`);
                if (!ok) return;

                try {
                    const res = await fetch(`/api/historial/${item.id}`, { method: "DELETE" });
                    const payloadDelete = await res.json();
                    if (!res.ok || !payloadDelete.success) {
                        notify(payloadDelete.error || "No se pudo eliminar el acta.", true);
                        return;
                    }
                    notify(`Acta ${numeroActa} eliminada del historial.`);
                    await refreshNumeroActaFormularioActivo(true);
                    await cargarHistorialActas();
                } catch (_err) {
                    notify("Error de red al eliminar acta.", true);
                }
            });

            tdActions.appendChild(btnDownload);
            tdActions.appendChild(btnEdit);
            tdActions.appendChild(btnDelete);

            tr.appendChild(tdN);
            tr.appendChild(tdNumeroActa);
            tr.appendChild(tdActions);
            tbody.appendChild(tr);
        });
    } catch (error) {
        tbody.innerHTML = '<tr><td colspan="3" class="text-center text-danger">Error cargando historial.</td></tr>';
        console.error(error);
    }
}

function setupHistorialModal() {
    const modalEl = document.getElementById("modalHistorialActas");
    if (!modalEl) return;
    const modal = bootstrap.Modal.getOrCreateInstance(modalEl);

    document.querySelectorAll(".btn-historial").forEach((btn) => {
        btn.addEventListener("click", async (e) => {
            e.preventDefault();
            modal.show();
            await cargarHistorialActas();
        });
    });
}

function initInformeSSE() {
    if (typeof EventSource === "undefined") return;
    if (informeEventSource) return;

    const connect = () => {
        informeEventSource = new EventSource("/api/events");

        informeEventSource.addEventListener("actas_changed", async () => {
            const modalEl = document.getElementById("modalHistorialActas");
            if (modalEl && modalEl.classList.contains("show")) {
                await cargarHistorialActas();
            }
            await refreshNumeroActaFormularioActivo(false);
        });

        informeEventSource.addEventListener("areas_reports_changed", async () => {
            await refreshNumeroActaAula(false);
        });

        informeEventSource.addEventListener("areas_reports_progress", (event) => {
            let payload = {};
            try {
                payload = JSON.parse(event.data || "{}");
            } catch (_err) {
                payload = {};
            }

            const jobId = String(payload.job_id || "").trim();
            if (!activeAulaBatchJobId || jobId !== activeAulaBatchJobId) return;
            setDownloadProgressValue(Number(payload.progress || 0));
            setDownloadProgressMessage(payload.message || "Procesando informe...");
        });

        informeEventSource.addEventListener("areas_reports_paused", (event) => {
            let payload = {};
            try {
                payload = JSON.parse(event.data || "{}");
            } catch (_err) {
                payload = {};
            }

            const jobId = String(payload.job_id || "").trim();
            if (!activeAulaBatchJobId || jobId !== activeAulaBatchJobId) return;
            setDownloadBatchPauseButton(true);
            setDownloadProgressMessage(payload.message || "Generación en pausa.");
        });

        informeEventSource.addEventListener("areas_reports_resumed", (event) => {
            let payload = {};
            try {
                payload = JSON.parse(event.data || "{}");
            } catch (_err) {
                payload = {};
            }

            const jobId = String(payload.job_id || "").trim();
            if (!activeAulaBatchJobId || jobId !== activeAulaBatchJobId) return;
            setDownloadBatchPauseButton(false);
            setDownloadProgressMessage(payload.message || "Generación reanudada.");
        });

        informeEventSource.addEventListener("areas_reports_ready", async (event) => {
            let payload = {};
            try {
                payload = JSON.parse(event.data || "{}");
            } catch (_err) {
                payload = {};
            }

            const jobId = String(payload.job_id || "").trim();
            if (!activeAulaBatchJobId || jobId !== activeAulaBatchJobId) return;

            finishDownloadProgress();
            setDownloadProgressTitle("Descarga lista");
            const downloadPath = String(payload.download_path || payload.zip_path || "").trim();
            const downloadKind = String(payload.download_kind || (payload.zip_path ? "zip" : "docx")).toLowerCase();
            setDownloadProgressMessage(
                downloadKind === "zip"
                    ? "Paquete ZIP generado correctamente."
                    : "Documento Word generado correctamente."
            );
            if (downloadPath) {
                const a = document.createElement("a");
                a.href = `/api/descargar?path=${encodeURIComponent(downloadPath)}`;
                a.style.display = "none";
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            }

            if (payload.next_numero_acta) {
                const aulaInput = document.querySelector('#form-aula input[name="numero_acta"]');
                setNumeroActaInputValue(aulaInput, payload.next_numero_acta, true);
                updateAulaBatchPreviewCard();
            } else {
                await refreshNumeroActaAula(true);
            }

            notify(
                `Se generaron ${payload.total_generated || 0} DOCX (${payload.start_numero_acta || "-"} a ${payload.end_numero_acta || "-"}).`
            );

            activeAulaBatchJobId = null;
            setDownloadBatchPauseButton(false);
            clearTimeout(closeDownloadModalTimer);
            closeDownloadModalTimer = setTimeout(() => {
                hideDownloadProgressModal(activeAulaBatchModal);
                activeAulaBatchModal = null;
            }, 350);
        });

        informeEventSource.addEventListener("areas_reports_cancelled", (event) => {
            let payload = {};
            try {
                payload = JSON.parse(event.data || "{}");
            } catch (_err) {
                payload = {};
            }

            const jobId = String(payload.job_id || "").trim();
            if (!activeAulaBatchJobId || jobId !== activeAulaBatchJobId) return;

            notify(payload.message || "Lote cancelado.");
            setDownloadProgressMessage(payload.message || "Lote cancelado.");
            activeAulaBatchJobId = null;
            setDownloadBatchPauseButton(false);
            clearTimeout(closeDownloadModalTimer);
            closeDownloadModalTimer = setTimeout(() => {
                hideDownloadProgressModal(activeAulaBatchModal);
                activeAulaBatchModal = null;
            }, 300);
        });

        informeEventSource.addEventListener("areas_reports_error", (event) => {
            let payload = {};
            try {
                payload = JSON.parse(event.data || "{}");
            } catch (_err) {
                payload = {};
            }

            const jobId = String(payload.job_id || "").trim();
            if (!activeAulaBatchJobId || jobId !== activeAulaBatchJobId) return;

            notify(payload.error || "Error durante la generación del lote de informes.", true);
            setDownloadProgressMessage("Ocurrió un error durante la generación.");
            activeAulaBatchJobId = null;
            setDownloadBatchPauseButton(false);
            clearTimeout(closeDownloadModalTimer);
            closeDownloadModalTimer = setTimeout(() => {
                hideDownloadProgressModal(activeAulaBatchModal);
                activeAulaBatchModal = null;
            }, 300);
        });

        informeEventSource.addEventListener("connected", () => {
            // Conexion activa.
        });

        informeEventSource.onerror = () => {
            try {
                informeEventSource.close();
            } catch (_err) {
                // noop
            }
            informeEventSource = null;
            setTimeout(connect, 2500);
        };
    };

    connect();
}

function renderColumnSelector() {
    const colSelector = document.getElementById("column-selector");
    if (!colSelector) return;

    colSelector.innerHTML = "";
    availableColumns.forEach((c) => {
        const div = document.createElement("div");
        div.className = "list-group-item list-group-item-action d-flex align-items-center py-2";
        const isChecked = selectedColumns.includes(c.id) ? "checked" : "";
        div.innerHTML = `
            <input class="form-check-input me-3 col-chk" type="checkbox" value="${c.id}" id="chk-col-${c.id}" ${isChecked}>
            <label class="form-check-label w-100 fw-medium small mb-0" for="chk-col-${c.id}" style="cursor:pointer;">${c.label}</label>
        `;
        colSelector.appendChild(div);
    });

    colSelector.querySelectorAll(".col-chk").forEach((chk) => {
        chk.addEventListener("change", (e) => {
            if (e.target.checked) {
                if (!selectedColumns.includes(e.target.value)) selectedColumns.push(e.target.value);
            } else {
                selectedColumns = selectedColumns.filter((id) => id !== e.target.value);
            }
            saveEntregaColumnPreferences();
            extraccionPage = 1;
            applyExtraccionesFilter();
        });
    });
}

function renderExtraccionTable(items) {
    const thead = document.getElementById("thead-extraccion-tr");
    const tbody = document.getElementById("tbody-extraccion");
    const contTabla = document.getElementById("contenedor-tabla-extraccion");
    const msgSinCols = document.getElementById("mensaje-sin-columnas");
    if (!thead || !tbody || !contTabla || !msgSinCols) return;

    if (!selectedColumns.length) {
        contTabla.classList.add("d-none");
        msgSinCols.classList.remove("d-none");
        return;
    }
    contTabla.classList.remove("d-none");
    msgSinCols.classList.add("d-none");

    thead.innerHTML = '<th class="text-center" style="width:50px;"><input class="form-check-input" type="checkbox" id="chk-all-items"></th>';
    selectedColumns.forEach((id) => {
        const col = availableColumns.find((c) => c.id === id);
        if (!col) return;
        const th = document.createElement("th");
        th.textContent = col.label;
        thead.appendChild(th);
    });

    tbody.innerHTML = "";
    if (!items.length) {
        tbody.innerHTML = `<tr><td colspan="${selectedColumns.length + 1}" class="text-center text-muted">No se encontraron equipos.</td></tr>`;
        return;
    }

    items.forEach((item) => {
        const tr = document.createElement("tr");
        const itemId = Number(item.id);
        const checked = selectedItemIds.has(itemId) ? "checked" : "";
        let html = `<td class="text-center"><input class="form-check-input item-extract-chk" type="checkbox" value="${itemId}" ${checked}></td>`;
        selectedColumns.forEach((id) => {
            const value = id === "cod_inventario" || id === "cod_esbye"
                ? normalizeCodeToPlaceholder(item[id])
                : (item[id] || "-");
            const cellClass = (id === "cod_inventario" || id === "cod_esbye") && isNoCodeValue(value)
                ? ' class="code-sc-cell"'
                : "";
            html += `<td${cellClass}>${value}</td>`;
        });
        tr.innerHTML = html;
        tbody.appendChild(tr);
    });

    document.getElementById("chk-all-items")?.addEventListener("change", (e) => {
        document.querySelectorAll(".item-extract-chk").forEach((c) => {
            const id = Number(c.value);
            c.checked = e.target.checked;
            if (c.checked) selectedItemIds.add(id);
            else selectedItemIds.delete(id);
        });
        updateSelectedItemsCount();
    });

    document.querySelectorAll(".item-extract-chk").forEach((chk) => {
        chk.addEventListener("change", (e) => {
            const id = Number(e.target.value);
            if (e.target.checked) selectedItemIds.add(id);
            else selectedItemIds.delete(id);
            updateSelectedItemsCount();

            const headerChk = document.getElementById("chk-all-items");
            if (headerChk) {
                const visibleChecks = Array.from(document.querySelectorAll(".item-extract-chk"));
                const checkedCount = visibleChecks.filter((node) => node.checked).length;
                headerChk.checked = visibleChecks.length > 0 && checkedCount === visibleChecks.length;
                headerChk.indeterminate = checkedCount > 0 && checkedCount < visibleChecks.length;
            }
        });
    });

    const headerChk = document.getElementById("chk-all-items");
    if (headerChk) {
        const visibleChecks = Array.from(document.querySelectorAll(".item-extract-chk"));
        const checkedCount = visibleChecks.filter((node) => node.checked).length;
        headerChk.checked = visibleChecks.length > 0 && checkedCount === visibleChecks.length;
        headerChk.indeterminate = checkedCount > 0 && checkedCount < visibleChecks.length;
    }
    updateSelectedItemsCount();
}

function updateExtraccionPagination(totalItems) {
    const info = document.getElementById("ext-page-info");
    const prev = document.getElementById("ext-page-prev");
    const next = document.getElementById("ext-page-next");
    const size = document.getElementById("ext-page-size");

    const total = Math.max(0, Number(totalItems || 0));
    const pages = Math.max(1, Math.ceil(total / extraccionPerPage));
    extraccionPage = Math.min(Math.max(extraccionPage, 1), pages);

    if (size) size.value = String(extraccionPerPage);
    if (info) info.textContent = total > 0 ? `Página ${extraccionPage} de ${pages}` : "0 de 0";
    if (prev) prev.disabled = extraccionPage <= 1 || total <= 0;
    if (next) next.disabled = extraccionPage >= pages || total <= 0;
}

function renderExtraccionCurrentPage() {
    const total = extraccionFilteredItems.length;
    const totalPages = Math.max(1, Math.ceil(total / extraccionPerPage));
    extraccionPage = Math.min(Math.max(extraccionPage, 1), totalPages);
    const start = (extraccionPage - 1) * extraccionPerPage;
    const pageItems = extraccionFilteredItems.slice(start, start + extraccionPerPage);
    renderExtraccionTable(pageItems);
    updateExtraccionPagination(total);
}

function updateSelectedItemsCount() {
    const el = document.getElementById("items-seleccionados-count");
    if (el) el.textContent = selectedItemIds.size;
}

function applyExtraccionesFilter() {
    if (!inventoryDataCache.length) return;
    const textVal = (document.getElementById("ext-buscar")?.value || "").toLowerCase();
    const areaSel = document.getElementById("ext-area");
    const areaTxt = areaSel && areaSel.value ? areaSel.options[areaSel.selectedIndex].text.toLowerCase() : "";

    extraccionFilteredItems = inventoryDataCache.filter((it) => {
        const matchText =
            !textVal ||
            String(it.cod_inventario || "").toLowerCase().includes(textVal) ||
            String(it.descripcion || "").toLowerCase().includes(textVal) ||
            String(it.marca || "").toLowerCase().includes(textVal);
        const matchArea = !areaTxt || String(it.ubicacion || "").toLowerCase().includes(areaTxt);
        return matchText && matchArea;
    });

    renderExtraccionCurrentPage();
}

async function loadExtraccionData() {
    if (inventoryDataCache.length) return;
    const tbody = document.getElementById("tbody-extraccion");
    if (tbody) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center text-primary"><div class="spinner-border spinner-border-sm" role="status"></div> Cargando inventario...</td></tr>';
    }
    try {
        const response = await api.get("/api/inventario?per_page=500");
        inventoryDataCache = response.data || [];
        extraccionFilteredItems = [...inventoryDataCache];
        renderExtraccionCurrentPage();
    } catch (error) {
        notify("Error cargando inventario para extracción: " + error.message, true);
    }
}

function setupExtraccionModal() {
    renderColumnSelector();

    const onExtraccionFilterChanged = () => {
        extraccionPage = 1;
        applyExtraccionesFilter();
    };

    document.getElementById("btn-buscar-ext")?.addEventListener("click", onExtraccionFilterChanged);
    document.getElementById("ext-buscar")?.addEventListener("input", onExtraccionFilterChanged);
    document.getElementById("ext-area")?.addEventListener("change", onExtraccionFilterChanged);
    document.getElementById("ext-piso")?.addEventListener("change", onExtraccionFilterChanged);
    document.getElementById("ext-bloque")?.addEventListener("change", onExtraccionFilterChanged);

    const extPagePrev = document.getElementById("ext-page-prev");
    const extPageNext = document.getElementById("ext-page-next");
    const extPageSize = document.getElementById("ext-page-size");

    if (extPagePrev && extPagePrev.dataset.bound !== "1") {
        extPagePrev.dataset.bound = "1";
        extPagePrev.addEventListener("click", () => {
            if (extraccionPage <= 1) return;
            extraccionPage -= 1;
            renderExtraccionCurrentPage();
        });
    }

    if (extPageNext && extPageNext.dataset.bound !== "1") {
        extPageNext.dataset.bound = "1";
        extPageNext.addEventListener("click", () => {
            extraccionPage += 1;
            renderExtraccionCurrentPage();
        });
    }

    if (extPageSize && extPageSize.dataset.bound !== "1") {
        extPageSize.dataset.bound = "1";
        extPageSize.addEventListener("change", () => {
            const parsed = Number(extPageSize.value || 25);
            extraccionPerPage = Number.isFinite(parsed) && parsed > 0 ? parsed : 25;
            extraccionPage = 1;
            renderExtraccionCurrentPage();
        });
    }

    const modalExtraerEl = document.getElementById("modalExtraerTabla");
    const modalExtraer = modalExtraerEl ? new bootstrap.Modal(modalExtraerEl) : null;

    document.querySelectorAll(".btn-extraer-tabla").forEach((btn) => {
        btn.addEventListener("click", (e) => {
            e.preventDefault();
            modalExtraer?.show();
        });
    });

    modalExtraerEl?.addEventListener("show.bs.modal", async () => {
        if (Array.isArray(window._globalSelectedTableRows) && window._globalSelectedTableRows.length) {
            selectedItemIds = new Set(window._globalSelectedTableRows.map((it) => Number(it.id)));
        }
        extraccionPage = 1;
        await loadExtraccionData();
        onExtraccionFilterChanged();
    });

    document.getElementById("btn-confirmar-extraccion")?.addEventListener("click", () => {
        if (!selectedItemIds.size) {
            notify("Por favor selecciona al menos 1 ítem del inventario para extraer.", true);
            return;
        }

        window._globalSelectedTableRows = inventoryDataCache.filter((it) => selectedItemIds.has(Number(it.id)));
        window._globalSelectedColumns = selectedColumns
            .map((id) => availableColumns.find((c) => c.id === id))
            .filter(Boolean)
            .map((c) => ({ id: c.id, label: c.label }));

        const active = document.querySelector(".settings-menu-btn.active")?.getAttribute("data-target");
        const targetDiv = active ? document.querySelector(`#${active} .resultado-tabla-container`) : null;
        if (targetDiv) {
            if (selectedItemIds.size > 0 && selectedColumns.length > 0) {
                showResultadoTabla(
                    targetDiv,
                    `<div class="alert alert-success d-inline-block p-2 px-4 shadow-sm mb-0"><i class="bi bi-check2-circle me-2"></i>Se extrajeron <strong>${selectedItemIds.size}</strong> filas y <strong>${selectedColumns.length}</strong> columnas.</div>`
                );
            } else {
                hideResultadoTabla(targetDiv);
            }
        }

        notify(`${selectedItemIds.size} ítems seleccionados y guardados temporalmente para el acta.`);

        document.dispatchEvent(
            new CustomEvent("informe:tablaExtraida", {
                detail: {
                    rows: window._globalSelectedTableRows.length,
                    columns: window._globalSelectedColumns.length,
                },
            })
        );
        modalExtraer?.hide();
    });
}

async function cargarPersonalDatalist() {
    try {
        const respuesta = await fetch("/api/personal");
        const data = await respuesta.json();
        if (!data.success) return;

        const datalist = document.getElementById("lista-personal");
        if (!datalist) return;

        datalist.innerHTML = "";
        (data.data || []).forEach((emp) => {
            const opcion = document.createElement("option");
            opcion.value = emp.nombre;
            if (emp.cargo) opcion.textContent = emp.cargo;
            datalist.appendChild(opcion);
        });

        const campos = document.querySelectorAll('input[name="entregado_por"], input[name="recibido_por"], input[name="usuario_final"], input[id*="entregado"], input[id*="recibido"], input[id*="usuario"], input[name="administradora"]');
        campos.forEach((input) => {
            input.setAttribute("list", "lista-personal");
            input.setAttribute("autocomplete", "off");
        });
    } catch (e) {
        console.error("Error cargando personal para autocompletado", e);
    }
}

document.addEventListener("DOMContentLoaded", async () => {
    await loadColumnPreferences();
    const canAccessActas = await checkActaTemplatesAccessGuard();
    if (!canAccessActas) return;

    setupModalBackdropGuardian();

    setupTabs();
    initNumeroActaTracking();
    document.querySelectorAll(".resultado-tabla-container").forEach((el) => hideResultadoTabla(el));

    try {
        const response = await api.get("/api/estructura?include_details=0");
        structureData = response.data || [];
    } catch (e) {
        notify("Error al cargar ubicaciones: " + e.message, true);
    }

    populateLocationSelects();
    setupExtraccionModal();
    setupRecepcionBienesModal();
    setupBajasBienesModal();
    setupBajasRegistradosModal();
    cargarPersonalDatalist();
    setupHistorialModal();
    initInformeSSE();
    initNumeroActaOnTabs();
    refreshNumeroActaAula(false);
    updateAulaBatchPreviewCard();

    document.getElementById("btn-vaciar-entrega")?.addEventListener("click", () => {
        clearInformeFormByType("entrega");
    });

    document.getElementById("btn-vaciar-recepcion")?.addEventListener("click", () => {
        clearInformeFormByType("recepcion");
    });

    document.getElementById("btn-vaciar-bajas")?.addEventListener("click", () => {
        clearInformeFormByType("bajas");
    });

    document.getElementById("btn-vaciar-aula")?.addEventListener("click", () => {
        clearAulaForm();
    });

    document.getElementById("aula-scope")?.addEventListener("change", () => {
        updateAulaBatchPreviewCard();
    });

    applyAulaScopeMode();

    const getTipoActivo = (originEl = null) => {
        if (originEl && typeof originEl.closest === "function") {
            const fromForm = originEl.closest("form[id^='form-']");
            if (fromForm?.id) {
                return String(fromForm.id).replace(/^form-/, "");
            }
            const fromSection = originEl.closest(".tab-section");
            if (fromSection?.id?.startsWith("sec-")) {
                return String(fromSection.id).replace(/^sec-/, "");
            }
        }

        const activeBtn = document.querySelector(".settings-menu-btn.active");
        return (activeBtn?.id || "tab-entrega").replace("tab-", "");
    };

    const triggerFileDownload = (path) => {
        if (!path) return;
        const a = document.createElement("a");
        a.href = `/api/descargar?path=${encodeURIComponent(path)}`;
        a.style.display = "none";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    };

    const generarYDescargarActa = async (originEl = null) => {
        if (isGeneratingActa) {
            notify("Ya se está generando un acta. Espere un momento...", true);
            return;
        }

        const tipo = getTipoActivo(originEl);
        const form = document.getElementById(`form-${tipo}`);
        if (!form) {
            notify("No se encontró el formulario activo para generar el acta.", true);
            return;
        }

        isGeneratingActa = true;
        document.querySelectorAll(".btn-descargar-acta").forEach((btn) => {
            btn.disabled = true;
        });

        const datosFormulario = Object.fromEntries(new FormData(form).entries());

        if (tipo === "entrega" && !isEntregaAreaTrabajoValida(datosFormulario.area_trabajo)) {
            notify("Debe seleccionar Bloque, Piso y Área para el Acta de Entrega.", true);
            isGeneratingActa = false;
            document.querySelectorAll(".btn-descargar-acta").forEach((btn) => {
                btn.disabled = false;
            });
            return;
        }

        if (tipo === "recepcion" && !isRecepcionAreaTrabajoValida(datosFormulario.area_trabajo)) {
            notify("Debe seleccionar Bloque, Piso y Área para el Acta de Recepción.", true);
            isGeneratingActa = false;
            document.querySelectorAll(".btn-descargar-acta").forEach((btn) => {
                btn.disabled = false;
            });
            return;
        }

        const editingActaIdForTipo =
            activeEditingActaId && normalizeTipoActa(activeEditingActaTipo) === normalizeTipoActa(tipo)
                ? activeEditingActaId
                : null;

        const numeroActaValidation = await validarNumeroActa(
            datosFormulario.numero_acta,
            tipo,
            editingActaIdForTipo,
        );
        if (!numeroActaValidation.valid) {
            notify(numeroActaValidation.error, true);
            isGeneratingActa = false;
            document.querySelectorAll(".btn-descargar-acta").forEach((btn) => {
                btn.disabled = false;
            });
            return;
        }

        const tablePayload = getActaTablePayload(tipo);
        const datosTabla = tablePayload.datosTabla;
        const datosColumnas = tablePayload.datosColumnas;

        if (tipo === "recepcion" && (!datosTabla.length || !datosColumnas.length)) {
            notify("Debe registrar al menos un bien nuevo en 'Registrar bienes' para generar el acta de recepción.", true);
            isGeneratingActa = false;
            document.querySelectorAll(".btn-descargar-acta").forEach((btn) => {
                btn.disabled = false;
            });
            return;
        }

        if ((tipo === "bajas" || tipo === "baja") && !datosTabla.length) {
            notify("Debe seleccionar al menos un bien en 'Seleccionar bienes' para generar el acta de baja.", true);
            isGeneratingActa = false;
            document.querySelectorAll(".btn-descargar-acta").forEach((btn) => {
                btn.disabled = false;
            });
            return;
        }

        const modal = showDownloadProgressModal();
        setDownloadBatchActionsVisible(false);
        setDownloadProgressTitle("Generando acta");
        startDownloadProgress();

        try {
            notify("Generando acta final (DOCX)...");
            const sendGenerate = async (forceSame = false) => {
                const response = await fetch("/api/informes/generar", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                        tipo,
                        datos_formulario: datosFormulario,
                        datos_tabla: datosTabla,
                        datos_columnas: datosColumnas,
                        vista_previa: false,
                        force_same_acta: forceSame,
                        editing_acta_id: editingActaIdForTipo || undefined,
                        template_snapshot_path: activeHistorialTemplateSnapshotPath || undefined,
                    }),
                });
                const result = await response.json();
                return { response, result };
            };

            let { response, result } = await sendGenerate(false);
            if (!response.ok && result?.duplicate_previous) {
                const prev = String(result.previous_numero_acta || "").trim();
                const msg = prev
                    ? `Esta acta es igual a la acta N° "${prev}". ¿Deseas guardarla de todas formas?`
                    : "Esta acta es igual a la anterior para este tipo. ¿Deseas guardarla de todas formas?";
                const goOn = window.confirm(msg);
                if (!goOn) {
                    finishDownloadProgress();
                    return;
                }
                ({ response, result } = await sendGenerate(true));
            }
            if (!response.ok || !result.success) {
                if (response.status === 409 && result?.next_numero_acta) {
                    const numeroSugerido = String(result.next_numero_acta || "").trim();
                    if (numeroSugerido) {
                        const numeroInput = form.querySelector('input[name="numero_acta"]');
                        setNumeroActaInputValue(numeroInput, numeroSugerido, true);
                    }
                    notify(result.error || "El número de acta cambió mientras se generaba. Se actualizó al siguiente disponible.", true);
                    finishDownloadProgress();
                    return;
                }
                notify(result.error || "No se pudo generar el acta para descargar.", true);
                finishDownloadProgress();
                return;
            }

            finishDownloadProgress();

            if (result.docx_path) triggerFileDownload(result.docx_path);
            setDownloadProgressTitle("Descarga iniciada");
            setDownloadProgressMessage("El archivo DOCX se está descargando.");
            notify("Acta DOCX generada y descarga iniciada.");
            activeHistorialTemplateSnapshotPath = null;
            activeEditingActaId = null;
            activeEditingActaTipo = null;
            resetActaDraftState(tipo);
            await refreshNumeroActaPorTipo(tipo, true);
        } catch (err) {
            console.error(err);
            notify("Error de red al generar y descargar el acta.", true);
            finishDownloadProgress();
        } finally {
            clearTimeout(closeDownloadModalTimer);
            closeDownloadModalTimer = setTimeout(() => {
                hideDownloadProgressModal(modal);
            }, 300);
            isGeneratingActa = false;
            document.querySelectorAll(".btn-descargar-acta").forEach((btn) => {
                btn.disabled = false;
            });
        }
    };

    document.querySelectorAll(".btn-descargar-acta").forEach((btn) => {
        btn.addEventListener("click", async (e) => {
            e.preventDefault();
            await generarYDescargarActa(e.currentTarget || e.target || null);
        });
    });

    document.querySelectorAll(".btn-generar-lote-aula").forEach((btn) => {
        btn.addEventListener("click", async (e) => {
            e.preventDefault();
            await generarLoteAula();
        });
    });

    const closeDownloadBtn = document.getElementById("btn-download-progress-close");
    if (closeDownloadBtn) {
        closeDownloadBtn.addEventListener("click", () => {
            hideDownloadProgressModal(activeAulaBatchModal || null);
        });
    }

    const pauseBtn = document.getElementById("btn-download-progress-pause");
    if (pauseBtn) {
        pauseBtn.addEventListener("click", async () => {
            await controlAulaBatchJob(activeAulaBatchPaused ? "resume" : "pause");
        });
    }

    const cancelBtn = document.getElementById("btn-download-progress-cancel");
    if (cancelBtn) {
        cancelBtn.addEventListener("click", async () => {
            await controlAulaBatchJob("cancel");
        });
    }
});
