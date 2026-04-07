const REQUIRED_VARS_INFORME = {
    entrega: [
        "numero_acta",
        "fecha_corte",
        "fecha_emision",
        "accion_personal",
        "entregado_por",
        "recibido_por",
        "rol_recibe",
        "area_trabajo",
    ],
    recepcion: ["numero_acta"],
    movimiento: ["numero_acta"],
    bajas: ["numero_acta"],
    traspaso: ["numero_acta"],
    "fin-gestion": ["numero_acta"],
    aula: ["tabla_dinamica"],
};

let previewTimer = null;
let previewRetryTimer = null;
let previewNoRenderRetries = 0;
let previewInFlight = false;
let previewQueuedWhileBusy = false;
const PREVIEW_DELAY_MS = 1400;
const MAX_PREVIEW_NO_RENDER_RETRIES = 3;

const INTERNAL_TEMPLATE_VARS = new Set([
    "tabla_items",
    "tabla_columnas",
    "tabla_filas",
    "tabla_dinamica",
    "celda",
]);

function isInternalTemplateVar(value) {
    const normalized = String(value || "").trim().toLowerCase();
    if (!normalized) return true;
    if (INTERNAL_TEMPLATE_VARS.has(normalized)) return true;
    if (normalized.startsWith("item.")) return true;
    if (normalized.startsWith("col.")) return true;
    if (normalized.startsWith("fila.")) return true;
    return false;
}

function getActiveTipoActa() {
    const activeBtn = document.querySelector(".settings-menu-btn.active");
    if (!activeBtn) return "entrega";
    return (activeBtn.id || "tab-entrega").replace("tab-", "");
}

function getFormByTipo(tipo) {
    return document.getElementById(`form-${tipo}`);
}

function setPreviewUiState(state) {
    const status = document.getElementById("preview-status");
    const loader = document.getElementById("preview-loader");
    if (!status || !loader) return;

    if (state === "loading") {
        status.textContent = "Generando...";
        loader.classList.remove("d-none");
        loader.classList.add("d-flex");
        return;
    }

    loader.classList.add("d-none");
    loader.classList.remove("d-flex");

    if (state === "ready") {
        status.textContent = "Actualizado";
    } else if (state === "error") {
        status.textContent = "Error de plantilla";
    } else {
        status.textContent = "Esperando datos...";
    }
}

function setPreviewError(message) {
    const iframe = document.getElementById("preview-iframe");
    const placeholder = document.getElementById("preview-placeholder");
    if (!placeholder || !iframe) return;

    iframe.classList.add("d-none");
    placeholder.classList.remove("d-none");
    placeholder.innerHTML = `
        <div class="text-danger fw-semibold mb-2">Error de plantilla en vista previa</div>
        <div class="small text-muted">${String(message || "No se pudo renderizar el documento.")}</div>
    `;
}

function setPreviewInfo(message) {
    const iframe = document.getElementById("preview-iframe");
    const placeholder = document.getElementById("preview-placeholder");
    if (!placeholder || !iframe) return;

    iframe.classList.add("d-none");
    placeholder.classList.remove("d-none");
    placeholder.innerHTML = `
        <div class="text-primary fw-semibold mb-2">Vista previa en proceso</div>
        <div class="small text-muted">${String(message || "Generando documento...")}</div>
    `;
}

function updatePreviewIframe(pdfPath) {
    const iframe = document.getElementById("preview-iframe");
    const placeholder = document.getElementById("preview-placeholder");
    if (!iframe || !placeholder || !pdfPath) return;

    const pdfUrl = `/api/ver?path=${encodeURIComponent(String(pdfPath))}&t=${Date.now()}`;

    placeholder.classList.add("d-none");
    iframe.classList.remove("d-none");
    iframe.removeAttribute("srcdoc");
    iframe.src = pdfUrl;
}

function updatePreviewHtml(htmlContent) {
    const iframe = document.getElementById("preview-iframe");
    const placeholder = document.getElementById("preview-placeholder");
    if (!iframe || !placeholder || !htmlContent) return;

    placeholder.classList.add("d-none");
    iframe.classList.remove("d-none");
    iframe.removeAttribute("src");
    iframe.srcdoc = String(htmlContent);
}

async function cargarCamposDinamicosActa(tipo) {
    const container = document.getElementById(`dinamico-${tipo}-container`);
    if (!container) return;

    container.innerHTML = "";

    try {
        const res = await fetch(`/api/plantillas/estado?tipo=${encodeURIComponent(tipo)}`);
        const payload = await res.json();
        if (!payload.success || !payload.existe) return;

        const variables = (payload.variables || []).filter(
            (v) => typeof v === "string" && !isInternalTemplateVar(v)
        );
        const required = REQUIRED_VARS_INFORME[tipo] || [];
        const extras = variables.filter((v) => !required.includes(v));
        if (!extras.length) return;

        const title = document.createElement("div");
        title.className = "col-12";
        title.innerHTML = '<h6 class="fw-bold text-primary mb-0 border-bottom pb-2"><i class="bi bi-stars me-2"></i>Campos Adicionales Detectados</h6>';
        container.appendChild(title);

        extras.forEach((ext) => {
            const col = document.createElement("div");
            col.className = "col-md-4 mt-3";
            const label = ext.replaceAll("_", " ");
            const id = `${tipo}-extra-${ext}`;
            col.innerHTML = `<label class="form-label text-capitalize">${label}</label><input type="text" class="form-control" name="${ext}" id="${id}" placeholder="Opcional">`;
            container.appendChild(col);
        });
    } catch (err) {
        console.warn("No se pudieron cargar variables dinámicas", err);
    }
}

function buildPreviewPayload(tipo) {
    const form = getFormByTipo(tipo);
    if (!form) return null;

    const entries = new FormData(form).entries();
    const datosFormulario = Object.fromEntries(entries);
    const tablaSeleccionada = Array.isArray(window._globalSelectedTableRows) ? window._globalSelectedTableRows : [];
    const columnasSeleccionadas = Array.isArray(window._globalSelectedColumns) ? window._globalSelectedColumns : [];
    const required = REQUIRED_VARS_INFORME[tipo] || [];
    const hasRequiredFormValues =
        required.length === 0 ||
        required.every((key) => String(datosFormulario[key] || "").trim() !== "");
    const hasAnyFormValues = Object.values(datosFormulario).some((v) => String(v || "").trim() !== "");
    const hasTableValues = tablaSeleccionada.length > 0 && columnasSeleccionadas.length > 0;
    // Evita llamadas pesadas mientras el usuario apenas empieza a escribir.
    if (!hasRequiredFormValues && !hasTableValues && !hasAnyFormValues) return null;
    if (!hasAnyFormValues && !hasTableValues) return null;

    return {
        tipo,
        datos_formulario: datosFormulario,
        datos_tabla: tablaSeleccionada,
        datos_columnas: columnasSeleccionadas,
        vista_previa: true,
    };
}

async function triggerPreview() {
    if (previewInFlight) {
        previewQueuedWhileBusy = true;
        return;
    }

    const tipo = getActiveTipoActa();
    if (tipo === "aula") {
        setPreviewUiState("idle");
        return;
    }
    const payload = buildPreviewPayload(tipo);
    if (!payload) {
        setPreviewUiState("idle");
        return;
    }

    previewInFlight = true;
    setPreviewUiState("loading");

    try {
        const response = await fetch("/api/informes/generar", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
        });
        const json = await response.json();
        if (response.ok && json.success && json.pdf_path) {
            clearTimeout(previewRetryTimer);
            previewNoRenderRetries = 0;
            updatePreviewIframe(json.pdf_path);
            setPreviewUiState("ready");
        } else if (response.ok && json.success && json.html_preview) {
            clearTimeout(previewRetryTimer);
            previewNoRenderRetries = 0;
            updatePreviewHtml(json.html_preview);
            setPreviewUiState("ready");
        } else if (response.ok && json.success && !json.pdf_path && !json.html_preview) {
            clearTimeout(previewRetryTimer);
            if (previewNoRenderRetries < MAX_PREVIEW_NO_RENDER_RETRIES) {
                previewNoRenderRetries += 1;
                setPreviewUiState("loading");
                setPreviewInfo("Word está procesando la vista previa. Reintentando...");
                previewRetryTimer = setTimeout(() => {
                    queuePreview(0);
                }, 1200);
            } else {
                setPreviewUiState("error");
                const warning = json && json.preview_warning
                    ? json.preview_warning
                    : "No se pudo generar la vista previa en PDF en este momento.";
                const docxHint = json && json.docx_path
                    ? `<div class=\"mt-2\"><a class=\"btn btn-sm btn-outline-primary\" href=\"/api/descargar?path=${encodeURIComponent(String(json.docx_path))}\">Descargar DOCX generado</a></div>`
                    : "";
                setPreviewError(`${warning}${docxHint}`);
                previewNoRenderRetries = 0;
            }
        } else {
            setPreviewUiState("error");
            const msg = json && json.error ? json.error : "No se generó el PDF para vista previa.";
            setPreviewError(msg);
            console.warn("Error en preview:", msg);
            previewNoRenderRetries = 0;
        }
    } catch (_err) {
        setPreviewUiState("error");
        setPreviewError("Error de red al generar la vista previa.");
        previewNoRenderRetries = 0;
    } finally {
        previewInFlight = false;
        if (previewQueuedWhileBusy) {
            previewQueuedWhileBusy = false;
            queuePreview(150);
        }
    }
}

function queuePreview(delay = PREVIEW_DELAY_MS) {
    clearTimeout(previewTimer);
    previewTimer = setTimeout(triggerPreview, delay);
}

window.queueInformePreview = queuePreview;

function bindPreviewEvents() {
    document.querySelectorAll(".tab-section form").forEach((form) => {
        form.addEventListener("input", () => queuePreview(PREVIEW_DELAY_MS));
        form.addEventListener("change", () => queuePreview(300));
    });

    document.querySelectorAll(".settings-menu-btn").forEach((tab) => {
        tab.addEventListener("click", async () => {
            const tipo = (tab.id || "tab-entrega").replace("tab-", "");
            await cargarCamposDinamicosActa(tipo);
            if (tipo !== "aula") {
                queuePreview(600);
            } else {
                setPreviewUiState("idle");
            }
        });
    });
}

document.addEventListener("DOMContentLoaded", async () => {
    await cargarCamposDinamicosActa("entrega");
    bindPreviewEvents();
    queuePreview(1200);
});

document.addEventListener("informe:tablaExtraida", () => {
    queuePreview(150);
});
