const api = {
    async get(url) {
        const response = await fetch(url);
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "Error en la red");
        return payload;
    },
    async send(url, method, body) {
        const response = await fetch(url, {
            method,
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
        });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "Error en la red");
        return payload;
    },
};

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

        const campos = document.querySelectorAll('input[name="entregado_por"], input[name="recibido_por"], input[name="usuario_final"], input[id*="entregado"], input[id*="recibido"], input[name="administradora"]');
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
        const tipo = getTipoActivo();
        const form = document.getElementById(`form-${tipo}`);
        if (!form) {
            notify("No se encontró el formulario activo para generar el acta.", true);
            return;
        }

        const datosFormulario = Object.fromEntries(new FormData(form).entries());
        const datosTabla = Array.isArray(window._globalSelectedTableRows) ? window._globalSelectedTableRows : [];
        const datosColumnas = Array.isArray(window._globalSelectedColumns) ? window._globalSelectedColumns : [];

        try {
            notify("Generando acta final (DOCX y PDF)...");
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
                return;
            }

            if (result.docx_path) triggerFileDownload(result.docx_path);
            if (result.pdf_path) {
                setTimeout(() => triggerFileDownload(result.pdf_path), 450);
            } else {
                notify("El DOCX se generó correctamente, pero no se pudo generar el PDF.", true);
            }

            notify("Acta generada y descarga iniciada.");
        } catch (err) {
            console.error(err);
            notify("Error de red al generar y descargar el acta.", true);
        }
    };

    document.querySelectorAll(".btn-descargar-acta").forEach((btn) => {
        btn.addEventListener("click", async (e) => {
            e.preventDefault();
            await generarYDescargarActa();
        });
    });
});
