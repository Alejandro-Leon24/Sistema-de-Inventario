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
        input.value = nextNumero;
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
    const refreshActiveFormNumero = async () => {
        const activeBtn = document.querySelector(".settings-menu-btn.active");
        const tipo = (activeBtn?.id || "tab-entrega").replace("tab-", "");
        const form = document.getElementById(`form-${tipo}`);
        await autocompletarNumeroActa(form);
    };

    refreshActiveFormNumero();
    document.querySelectorAll(".settings-menu-btn").forEach((btn) => {
        btn.addEventListener("click", () => {
            setTimeout(() => {
                refreshActiveFormNumero();
            }, 20);
        });
    });
}

function setupTabs() {
    const tabs = document.querySelectorAll(".settings-menu-btn");
    const sections = document.querySelectorAll(".tab-section");

    tabs.forEach((tab) => {
        tab.addEventListener("click", (e) => {
            e.preventDefault();
            tabs.forEach((t) => t.classList.remove("active"));
            sections.forEach((s) => s.classList.add("d-none"));
            tab.classList.add("active");
            const targetId = tab.getAttribute("data-target");
            document.getElementById(targetId)?.classList.remove("d-none");
        });
    });
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
        });
    });
}

function normalizeTipoActa(value) {
    return String(value || "").trim().toLowerCase();
}

function setPreviewStatus(message) {
    const status = document.getElementById("preview-status");
    if (status) status.textContent = message;
}

function showDownloadProgressModal() {
    const modalEl = document.getElementById("modalDescargaProgress");
    if (!modalEl) return null;
    const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
    modal.show();
    return modal;
}

function hideDownloadProgressModal(modal) {
    const modalEl = document.getElementById("modalDescargaProgress");
    if (modal) {
        modal.hide();
    } else if (modalEl) {
        bootstrap.Modal.getOrCreateInstance(modalEl).hide();
    }

    // Fallback defensivo por si bootstrap deja backdrop colgado.
    setTimeout(() => {
        if (!modalEl) return;
        modalEl.classList.remove("show");
        modalEl.style.display = "none";
        document.body.classList.remove("modal-open");
        document.body.style.removeProperty("padding-right");
        document.querySelectorAll(".modal-backdrop").forEach((el) => el.remove());
    }, 20);
}

function startDownloadProgress() {
    const bar = document.getElementById("descarga-progress-bar");
    const text = document.getElementById("descarga-progress-text");
    downloadProgressValue = 0;

    if (bar) bar.style.width = "0%";
    if (text) text.textContent = "0%";

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

    const activeTarget = document.querySelector(".settings-menu-btn.active")?.getAttribute("data-target");
    const targetDiv = activeTarget ? document.querySelector(`#${activeTarget} .resultado-tabla-container`) : null;
    if (targetDiv) {
        const rows = window._globalSelectedTableRows.length;
        const cols = window._globalSelectedColumns.length;
        targetDiv.innerHTML = `<div class="alert alert-info d-inline-block p-2 px-4 shadow-sm mb-0"><i class="bi bi-pencil-square me-2"></i>Acta cargada para edición: <strong>${rows}</strong> filas y <strong>${cols}</strong> columnas.</div>`;
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

            const tdFecha = document.createElement("td");
            tdFecha.textContent = formatFechaActa(item.fecha);

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

            tdActions.appendChild(btnDownload);
            tdActions.appendChild(btnEdit);

            tr.appendChild(tdN);
            tr.appendChild(tdFecha);
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
            targetDiv.innerHTML = `<div class="alert alert-success d-inline-block p-2 px-4 shadow-sm mb-0"><i class="bi bi-check2-circle me-2"></i>Se extrajeron <strong>${selectedItemIds.size}</strong> filas y <strong>${selectedColumns.length}</strong> columnas.</div>`;
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
    setupTabs();

    try {
        const response = await api.get("/api/estructura");
        structureData = response.data || [];
    } catch (e) {
        notify("Error al cargar ubicaciones: " + e.message, true);
    }

    populateLocationSelects();
    setupExtraccionModal();
    cargarPersonalDatalist();
    setupHistorialModal();
    initInformeSSE();
    initNumeroActaOnTabs();

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
        const numeroActaValidation = await validarNumeroActa(datosFormulario.numero_acta);
        if (!numeroActaValidation.valid) {
            notify(numeroActaValidation.error, true);
            isGeneratingActa = false;
            document.querySelectorAll(".btn-descargar-acta").forEach((btn) => {
                btn.disabled = false;
            });
            return;
        }

        const datosTabla = Array.isArray(window._globalSelectedTableRows) ? window._globalSelectedTableRows : [];
        const datosColumnas = Array.isArray(window._globalSelectedColumns) ? window._globalSelectedColumns : [];

        const modal = showDownloadProgressModal();
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
            notify("Acta DOCX generada y descarga iniciada.");
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
});
