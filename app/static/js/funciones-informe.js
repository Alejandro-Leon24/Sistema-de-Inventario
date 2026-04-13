const api = window.api;

let selectedColumns = ["cod_inventario", "descripcion", "marca", "modelo", "serie", "estado"];
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
];

let structureData = [];
let inventoryDataCache = [];
let selectedItemIds = new Set();
window._globalSelectedTableRows = [];
window._globalSelectedColumns = [];

let downloadProgressTimer = null;
let downloadProgressValue = 0;
let isGeneratingActa = false;
let closeDownloadModalTimer = null;
let informeEventSource = null;
let activeHistorialTemplateSnapshotPath = null;
let activeAulaBatchJobId = null;
let activeAulaBatchModal = null;
let activeAulaBatchPaused = false;
let recepcionBienesTemp = [];
let recepcionEditIndex = -1;
let recepcionSelectedColumnIds = ["cod_inventario", "descripcion", "marca", "modelo", "cantidad", "estado"];

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

    return unique.length ? unique : ["cod_inventario", "descripcion", "marca", "modelo", "cantidad", "estado"];
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

function normalizeRecepcionRowLikeInventarioImport(row, forcedLocation) {
    const src = row && typeof row === "object" ? row : {};
    const toText = (v) => String(v ?? "").trim();
    const cantidadNum = parseInt(String(src.cantidad ?? "").trim(), 10);
    return {
        ...src,
        cod_inventario: toText(src.cod_inventario),
        cod_esbye: toText(src.cod_esbye),
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
    const tipos = ["entrega", "recepcion", "aula"];
    const labels = { entrega: "Acta de Entrega", recepcion: "Acta de Recepción", aula: "Inventario por Área" };
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

    const nextNumero = await obtenerNumeroActaSiguiente();
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

async function obtenerNumeroActaSiguiente() {
    try {
        const response = await fetch("/api/historial/numero-acta/siguiente");
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

    const nextNumero = await obtenerNumeroActaSiguiente();
    if (nextNumero) {
        setNumeroActaInputValue(input, nextNumero, true);
    }
}

async function validarNumeroActa(numeroActa) {
    const value = String(numeroActa || "").trim();
    if (!value) {
        return { valid: false, error: "El número de acta es obligatorio." };
    }

    try {
        const response = await fetch(`/api/historial/numero-acta/validar?numero_acta=${encodeURIComponent(value)}`);
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
        iframe.srcdoc = String(result.html_preview);
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
    return {
        datosTabla: Array.isArray(window._globalSelectedTableRows) ? window._globalSelectedTableRows : [],
        datosColumnas: Array.isArray(window._globalSelectedColumns) ? window._globalSelectedColumns : [],
    };
}

window.getInformeActaTablePayload = getActaTablePayload;

function getRecepcionResultContainer() {
    return document.querySelector("#sec-recepcion .resultado-tabla-container");
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

function getRecepcionBienFormValues() {
    const cantidadRaw = String(document.getElementById("recepcion-bien-cantidad")?.value || "1").trim();
    const cantidad = Number(cantidadRaw);
    return {
        cod_inventario: String(document.getElementById("recepcion-bien-cod-inventario")?.value || "").trim(),
        cod_esbye: String(document.getElementById("recepcion-bien-cod-esbye")?.value || "").trim(),
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
    const assignSelect = (id, value) => {
        const el = document.getElementById(id);
        const val = String(value || "").trim();
        if (!el) return;
        if (!val) {
            el.value = "";
            return;
        }
        if (!Array.from(el.options || []).some((opt) => String(opt.value) === val)) {
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

    countBadge.textContent = String(recepcionBienesTemp.length);
    if (!recepcionBienesTemp.length) {
        tbody.innerHTML = `<tr><td colspan="${selectedCols.length + 2}" class="text-center text-muted">Aún no hay bienes registrados.</td></tr>`;
        return;
    }

    tbody.innerHTML = "";
    recepcionBienesTemp.forEach((item, idx) => {
        const tr = document.createElement("tr");
        tr.style.cursor = "pointer";
        tr.dataset.index = String(idx);
        let rowHtml = `<td>${idx + 1}</td>`;
        selectedCols.forEach((col) => {
            const raw = item?.[col.id];
            const value = raw == null || String(raw).trim() === "" ? "-" : String(raw);
            rowHtml += `<td title="${escapeCell(value)}">${escapeCell(value)}</td>`;
        });
        rowHtml += `<td class="text-center"><button type="button" class="btn btn-sm btn-outline-danger btn-recepcion-bien-eliminar" data-index="${idx}"><i class="bi bi-trash"></i></button></td>`;
        tr.innerHTML = rowHtml;
        tbody.appendChild(tr);
    });

    tbody.querySelectorAll("tr").forEach((tr) => {
        tr.addEventListener("click", async (event) => {
            if (event.target.closest(".btn-recepcion-bien-eliminar")) return;
            const idx = Number(tr.dataset.index);
            if (!Number.isFinite(idx)) return;
            // Asegura que los catálogos de selects estén listos antes de cargar el item.
            await loadRecepcionSelectOptions();
            loadRecepcionBienIntoForm(idx);
        });
    });

    tbody.querySelectorAll(".btn-recepcion-bien-eliminar").forEach((btn) => {
        btn.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            const idx = Number(btn.dataset.index);
            if (!Number.isFinite(idx)) return;
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
    const panel1 = document.getElementById("recepcion-import-step-file");
    const panel2 = document.getElementById("recepcion-import-step-mapping");
    const panel3 = document.getElementById("recepcion-import-step-run");
    const stepBadges = [];
    const stepLabels = [];
    const ubicacionReadonly = document.getElementById("recepcion-bien-ubicacion");
    const cuentaSelect = document.getElementById("recepcion-bien-cuenta");
    const estadoSelect = document.getElementById("recepcion-bien-estado");
    const usuarioFinalSelect = document.getElementById("recepcion-bien-usuario-final");
    const ubicacionEsbyeSelect = document.getElementById("recepcion-bien-ubicacion-esbye");
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

    const CHUNK_SIZE = 20;
    let shouldReturnToRecepcionModal = false;
    let shouldReturnFromColumnsModal = false;
    let pendingChildModal = "";
    let isTransientImportModalHide = false;
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

    const importState = {
        step: 1,
        sessionId: null,
        headers: [],
        previewRows: [],
        totalRows: 0,
        mappings: [],
        lockedMainLocationColIdx: null,
        startIndex: 0,
        chunkSize: CHUNK_SIZE,
        hasMore: false,
        previewChunkData: null,
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

    const buildRecepcionAreaOptions = () => {
        const out = [];
        (structureData || []).forEach((block) => {
            (block.pisos || []).forEach((floor) => {
                (floor.areas || []).forEach((area) => {
                    out.push(`${block.nombre} / ${floor.nombre} / ${area.nombre}`);
                });
            });
        });
        return out;
    };

    const loadRecepcionEsbyeAreaOptions = () => {
        if (!ubicacionEsbyeSelect) return;
        const current = String(ubicacionEsbyeSelect.value || "").trim();
        ubicacionEsbyeSelect.innerHTML = '<option value="">Sin ubicación ESBYE</option>';
        buildRecepcionAreaOptions().forEach((label) => {
            const option = document.createElement("option");
            option.value = label;
            option.textContent = label;
            ubicacionEsbyeSelect.appendChild(option);
        });
        if (current && Array.from(ubicacionEsbyeSelect.options).some((opt) => opt.value === current)) {
            ubicacionEsbyeSelect.value = current;
        }
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
            return `
                <tr>
                    <td>${Number(row.row_index || 0) + 1}</td>
                    <td>${badge(row.status)}</td>
                    <td>${String(d.cod_inventario || "-")}</td>
                    <td>${String(d.cod_esbye || "-")}</td>
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
        try {
            const [estadosRes, cuentasRes, adminsRes] = await Promise.all([
                api.get("/api/parametros/estados"),
                api.get("/api/parametros/cuentas"),
                api.get("/api/administradores"),
            ]);
            fillSimpleSelect(estadoSelect, estadosRes.data || [], "-- Seleccionar estado --");
            fillSimpleSelect(cuentaSelect, cuentasRes.data || [], "-- Seleccionar cuenta --");
            fillSimpleSelect(usuarioFinalSelect, adminsRes.data || [], "-- Seleccionar personal --");
        } catch (_err) {
            // Permite seguir usando el modal aun si no cargan catálogos.
        }
    };

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
        importState.headers = [];
        importState.previewRows = [];
        importState.totalRows = 0;
        importState.mappings = [];
        importState.lockedMainLocationColIdx = null;
        importState.startIndex = 0;
        importState.hasMore = false;
        importState.previewChunkData = null;
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

    const handleImportFile = async (file) => {
        if (!file) return;
        if (!String(file.name || "").toLowerCase().endsWith(".xlsx")) {
            showImportStatus('<i class="bi bi-x-circle me-1"></i>Solo se aceptan archivos .xlsx.', "danger");
            return;
        }

        showImportStatus(`<span class="spinner-border spinner-border-sm me-2" role="status"></span>Procesando <strong>${String(file.name || "archivo")}</strong>...`);
        const formData = new FormData();
        formData.append("file", file);

        try {
            const pre = await fetch("/api/inventario/previsualizar-excel", { method: "POST", body: formData });
            const preData = await pre.json();
            if (!pre.ok || preData.error) {
                showImportStatus(`<i class="bi bi-x-circle me-1"></i>${preData.error || "No se pudo leer el Excel."}`, "danger");
                return;
            }

            importState.sessionId = preData.session_id;
            importState.headers = Array.isArray(preData.headers) ? preData.headers : [];
            importState.previewRows = Array.isArray(preData.preview_rows) ? preData.preview_rows : [];
            importState.totalRows = Number(preData.total_rows || 0);
            importState.startIndex = 0;
            importState.hasMore = importState.totalRows > 0;
            const suggested = Array.isArray(preData.suggested_mapping) ? preData.suggested_mapping : [];
            const rawMappings = importState.headers.map((_, idx) => String(suggested[idx] || ""));
            importState.mappings = normalizeLocationMappings(importState.headers, rawMappings);

            renderMappingStep();
            updateImportProgress();
            clearImportStatus();
            if (btnImportNext) btnImportNext.disabled = false;
            setImportStep(2);
        } catch (_err) {
            showImportStatus('<i class="bi bi-x-circle me-1"></i>Error de red al subir archivo.', "danger");
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
                : await api.send("/api/inventario/excel-a-filas-recepcion", "POST", {
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
            const res = await api.send("/api/inventario/excel-a-filas-recepcion", "POST", {
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
                const nextPreview = await api.send("/api/inventario/excel-a-filas-recepcion", "POST", {
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
        loadRecepcionEsbyeAreaOptions();
        renderRecepcionBienesTable();
        modal.show();
    });

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
        api.send("/api/inventario/excel-a-filas-recepcion", "POST", {
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
        recepcionSelectedColumnIds = normalizeRecepcionColumnIds(parsed.columnas || []);
        renderRecepcionBienesTable();
        updateRecepcionSummary();
    }
    activeHistorialTemplateSnapshotPath =
        (record && record.plantilla_snapshot_path) ||
        (parsed && parsed.plantilla && parsed.plantilla.snapshot_path) ||
        null;

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
            html += `<td>${item[id] || "-"}</td>`;
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
        });
    });
    updateSelectedItemsCount();
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

    const filtered = inventoryDataCache.filter((it) => {
        const matchText =
            !textVal ||
            String(it.cod_inventario || "").toLowerCase().includes(textVal) ||
            String(it.descripcion || "").toLowerCase().includes(textVal) ||
            String(it.marca || "").toLowerCase().includes(textVal);
        const matchArea = !areaTxt || String(it.ubicacion || "").toLowerCase().includes(areaTxt);
        return matchText && matchArea;
    });

    renderExtraccionTable(filtered);
}

async function loadExtraccionData() {
    if (inventoryDataCache.length) return;
    const tbody = document.getElementById("tbody-extraccion");
    if (tbody) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center text-primary"><div class="spinner-border spinner-border-sm" role="status"></div> Cargando inventario...</td></tr>';
    }
    try {
        const response = await api.get("/api/inventario");
        inventoryDataCache = response.data || [];
        renderExtraccionTable(inventoryDataCache);
    } catch (error) {
        notify("Error cargando inventario para extracción: " + error.message, true);
    }
}

function setupExtraccionModal() {
    renderColumnSelector();

    document.getElementById("btn-buscar-ext")?.addEventListener("click", applyExtraccionesFilter);
    document.getElementById("ext-buscar")?.addEventListener("input", applyExtraccionesFilter);
    document.getElementById("ext-area")?.addEventListener("change", applyExtraccionesFilter);
    document.getElementById("ext-piso")?.addEventListener("change", applyExtraccionesFilter);
    document.getElementById("ext-bloque")?.addEventListener("change", applyExtraccionesFilter);

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
        await loadExtraccionData();
        applyExtraccionesFilter();
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
    cargarPersonalDatalist();
    setupHistorialModal();
    initInformeSSE();
    initNumeroActaOnTabs();
    refreshNumeroActaAula(false);
    updateAulaBatchPreviewCard();

    document.getElementById("aula-scope")?.addEventListener("change", () => {
        updateAulaBatchPreviewCard();
    });

    applyAulaScopeMode();

    const getTipoActivo = () => {
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

    const generarYDescargarActa = async () => {
        if (isGeneratingActa) {
            notify("Ya se está generando un acta. Espere un momento...", true);
            return;
        }

        const tipo = getTipoActivo();
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

        const numeroActaValidation = await validarNumeroActa(datosFormulario.numero_acta);
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

        const modal = showDownloadProgressModal();
        setDownloadBatchActionsVisible(false);
        setDownloadProgressTitle("Generando acta");
        startDownloadProgress();

        try {
            notify("Generando acta final (DOCX)...");
            const response = await fetch("/api/informes/generar", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    tipo,
                    datos_formulario: datosFormulario,
                    datos_tabla: datosTabla,
                    datos_columnas: datosColumnas,
                    vista_previa: false,
                    template_snapshot_path: activeHistorialTemplateSnapshotPath || undefined,
                }),
            });

            const result = await response.json();
            if (!response.ok || !result.success) {
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
            await refreshNumeroActaFormularioActivo(true);
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
            await generarYDescargarActa();
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
