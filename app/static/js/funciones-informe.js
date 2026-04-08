const api = window.api;

let selectedColumns = ["cod_inventario", "descripcion", "marca", "modelo", "serie", "estado"];
const availableColumns = [
    { id: "cod_inventario", label: "CODIGO INV." },
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

const RECEPCION_BIENES_COLUMNS = [
    { id: "cod_inventario", label: "CODIGO INV." },
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
        if (payload.lower_than_max) {
            return {
                valid: false,
                error: `El número ${value} es menor al último registrado (${payload.max_numero_acta}).`,
            };
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
        return {
            datosTabla: Array.isArray(recepcionBienesTemp) ? recepcionBienesTemp : [],
            datosColumnas: RECEPCION_BIENES_COLUMNS,
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
        ubicacion: String(document.getElementById("recepcion-area-trabajo")?.value || "").trim(),
        fecha_adquisicion: String(document.getElementById("recepcion-bien-fecha-adquisicion")?.value || "").trim(),
        valor: String(document.getElementById("recepcion-bien-valor")?.value || "").trim(),
        usuario_final: String(document.getElementById("recepcion-bien-usuario-final")?.value || "").trim(),
        observacion: String(document.getElementById("recepcion-bien-observacion")?.value || "").trim(),
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
    const assign = (id, value) => {
        const el = document.getElementById(id);
        if (el) el.value = String(value || "");
    };
    assign("recepcion-bien-cod-inventario", item.cod_inventario);
    assign("recepcion-bien-cod-esbye", item.cod_esbye);
    assign("recepcion-bien-cuenta", item.cuenta);
    assign("recepcion-bien-cantidad", item.cantidad || 1);
    assign("recepcion-bien-descripcion", item.descripcion);
    assign("recepcion-bien-marca", item.marca);
    assign("recepcion-bien-modelo", item.modelo);
    assign("recepcion-bien-serie", item.serie);
    assign("recepcion-bien-estado", item.estado);
    assign("recepcion-bien-fecha-adquisicion", item.fecha_adquisicion);
    assign("recepcion-bien-valor", item.valor);
    assign("recepcion-bien-usuario-final", item.usuario_final);
    assign("recepcion-bien-observacion", item.observacion);
    setRecepcionEditMode(index);
}

function renderRecepcionBienesTable() {
    const tbody = document.getElementById("tbody-recepcion-bienes");
    const countBadge = document.getElementById("recepcion-bienes-count");
    if (!tbody || !countBadge) return;

    countBadge.textContent = String(recepcionBienesTemp.length);
    if (!recepcionBienesTemp.length) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center text-muted">Aún no hay bienes registrados.</td></tr>';
        return;
    }

    tbody.innerHTML = "";
    recepcionBienesTemp.forEach((item, idx) => {
        const tr = document.createElement("tr");
        tr.style.cursor = "pointer";
        tr.dataset.index = String(idx);
        tr.innerHTML = `
            <td>${idx + 1}</td>
            <td>${String(item.cod_inventario || "-")}</td>
            <td>${String(item.descripcion || "-")}</td>
            <td>${String(item.marca || "-")}</td>
            <td>${String(item.modelo || "-")}</td>
            <td>${String(item.cantidad || 1)}</td>
            <td class="text-center"><button type="button" class="btn btn-sm btn-outline-danger btn-recepcion-bien-eliminar" data-index="${idx}"><i class="bi bi-trash"></i></button></td>
        `;
        tbody.appendChild(tr);
    });

    tbody.querySelectorAll("tr").forEach((tr) => {
        tr.addEventListener("click", (event) => {
            if (event.target.closest(".btn-recepcion-bien-eliminar")) return;
            const idx = Number(tr.dataset.index);
            if (!Number.isFinite(idx)) return;
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
    const btnOpen = document.querySelector("#sec-recepcion .btn-registrar-bienes");
    const btnSave = document.getElementById("btn-recepcion-bien-guardar");
    const btnCancel = document.getElementById("btn-recepcion-bien-cancelar");
    const btnConfirm = document.getElementById("btn-confirmar-recepcion-bienes");
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

    btnOpen.addEventListener("click", (event) => {
        event.preventDefault();
        renderRecepcionBienesTable();
        modal.show();
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
        updateRecepcionSummary();
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

    setupTabs();
    initNumeroActaTracking();
    document.querySelectorAll(".resultado-tabla-container").forEach((el) => hideResultadoTabla(el));

    try {
        const response = await api.get("/api/estructura");
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
