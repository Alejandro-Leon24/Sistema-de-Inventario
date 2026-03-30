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

document.addEventListener("DOMContentLoaded", async () => {
    // 1. Configuración de Pestañas
    const tabs = document.querySelectorAll("#informe-tabs .nav-link");
    const sections = document.querySelectorAll(".tab-section");

    tabs.forEach(tab => {
        tab.addEventListener("click", (e) => {
            e.preventDefault();
            // Desactivar todos los tabs
            tabs.forEach(t => t.classList.remove("active"));
            // Ocultar todas las secciones
            sections.forEach(s => s.classList.add("d-none"));

            // Activar tab actual y mostrar sección objetivo
            tab.classList.add("active");
            const targetId = tab.getAttribute("data-target");
            document.getElementById(targetId).classList.remove("d-none");
        });
    });

    // 2. Carga de Estructura de Ubicaciones (Bloques, Pisos, Áreas)
    let structureData = [];
    try {
        const response = await api.get("/api/estructura");
        structureData = response.data;
    } catch (e) {
        notify("Error al cargar ubicaciones: " + e.message, true);
    }

    // Identificar los grupos de selects que actúan en cascada
    const locationGroups = [
        { pre: "entrega" },
        { pre: "recepcion" },
        { pre: "aula", routeBadgeId: "aula-ruta-seleccionada" },
        { pre: "ext" } // El de modal extracción
    ];

    locationGroups.forEach(group => {
        const selBloque = document.getElementById(`${group.pre}-bloque`);
        const selPiso = document.getElementById(`${group.pre}-piso`);
        const selArea = document.getElementById(`${group.pre}-area`);

        if (!selBloque || !selPiso || !selArea) return;

        // Llenar bloques iniciales
        structureData.forEach(bloque => {
            const opt = document.createElement("option");
            opt.value = bloque.id;
            opt.textContent = bloque.nombre;
            selBloque.appendChild(opt);
        });

        selBloque.addEventListener("change", () => {
            const bId = parseInt(selBloque.value);
            selPiso.innerHTML = '<option value="">Seleccionar piso...</option>';
            selArea.innerHTML = '<option value="">Seleccionar área...</option>';
            selPiso.disabled = true;
            selArea.disabled = true;
            updateBadge(group);

            if (!bId) return;

            const bloqueEncontrado = structureData.find(b => b.id === bId);
            if (bloqueEncontrado && bloqueEncontrado.pisos && bloqueEncontrado.pisos.length > 0) {
                selPiso.disabled = false;
                bloqueEncontrado.pisos.forEach(piso => {
                    const opt = document.createElement("option");
                    opt.value = piso.id;
                    opt.textContent = piso.nombre;
                    selPiso.appendChild(opt);
                });
            }
        });

        selPiso.addEventListener("change", () => {
            const bId = parseInt(selBloque.value);
            const pId = parseInt(selPiso.value);
            selArea.innerHTML = '<option value="">Seleccionar área...</option>';
            selArea.disabled = true;
            updateBadge(group);

            if (!bId || !pId) return;

            const bloqueEncontrado = structureData.find(b => b.id === bId);
            if (!bloqueEncontrado) return;
            const pisoEncontrado = bloqueEncontrado.pisos.find(p => p.id === pId);

            if (pisoEncontrado && pisoEncontrado.areas && pisoEncontrado.areas.length > 0) {
                selArea.disabled = false;
                pisoEncontrado.areas.forEach(area => {
                    const opt = document.createElement("option");
                    opt.value = area.id;
                    opt.textContent = area.nombre;
                    selArea.appendChild(opt);
                });
            }
        });

        selArea.addEventListener("change", () => {
            updateBadge(group);
        });
    });

    function updateBadge(group) {
        if (!group.routeBadgeId) return;
        const badgeEl = document.getElementById(group.routeBadgeId);
        
        const selBloque = document.getElementById(`${group.pre}-bloque`);
        const selPiso = document.getElementById(`${group.pre}-piso`);
        const selArea = document.getElementById(`${group.pre}-area`);

        const txtB = selBloque.options[selBloque.selectedIndex]?.text;
        const valB = selBloque.value;
        const txtA = selArea.options[selArea.selectedIndex]?.text;
        const valA = selArea.value;

        if (valB && valA) {
            badgeEl.innerHTML = `${txtB} <i class="bi bi-chevron-right mx-2"></i> ${txtA}`;
        } else {
            badgeEl.innerHTML = "Seleccione el bloque y el área correspondientes...";
        }
    }

    // 3. Modal Extracción
    const btnExtraerGeneral = document.querySelectorAll(".btn-extraer-tabla");
    const modalExtraerEl = document.getElementById("modalExtraerTabla");
    let modalExtraer = null;
    if (modalExtraerEl) {
        modalExtraer = new bootstrap.Modal(modalExtraerEl);
    }
    
    // Abrir Modal
    btnExtraerGeneral.forEach(btn => {
        btn.addEventListener("click", (e) => {
            e.preventDefault();
            e.stopPropagation(); // Evitar comportamientos no deseados en formularios
            if(modalExtraer) modalExtraer.show();
        });
    });

    let inventoryDataCache = [];
    const tbodyExtraccion = document.getElementById("tbody-extraccion");
    const checkAllItems = document.getElementById("chk-all-items");
    const selItemsCount = document.getElementById("items-seleccionados-count");

    async function loadExtraccionData() {
        if (inventoryDataCache.length > 0) return; // cacheo de la sesión
        try {
            tbodyExtraccion.innerHTML = '<tr><td colspan="7" class="text-center text-primary"><div class="spinner-border spinner-border-sm" role="status"></div> Cargando inventario...</td></tr>';
            const response = await api.get("/api/inventario");
            inventoryDataCache = response.data || [];
            renderExtraccionTable(inventoryDataCache);
        } catch (error) {
            notify("Error cargando inventario para extracción: " + error.message, true);
        }
    }

    function renderExtraccionTable(items) {
        tbodyExtraccion.innerHTML = "";
        if (items.length === 0) {
            tbodyExtraccion.innerHTML = `<tr><td colspan="7" class="text-center text-muted">No se encontraron equipos bajo estos filtros.</td></tr>`;
            return;
        }

        items.forEach(item => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td class="text-center">
                    <input class="form-check-input item-extract-chk" type="checkbox" value="${item.id}">
                </td>
                <td>${item.cod_inventario || '-'}</td>
                <td>${item.descripcion || '-'}</td>
                <td>${item.marca || '-'}</td>
                <td>${item.modelo || '-'}</td>
                <td>${item.serie || '-'}</td>
                <td><span class="badge bg-secondary">${item.estado || '-'}</span></td>
            `;
            tbodyExtraccion.appendChild(tr);
        });

        // Event listener para actualizar conteo cuando se chequea individualmente
        document.querySelectorAll(".item-extract-chk").forEach(chk => {
            chk.addEventListener("change", updateSelectedItemsCount);
        });
        updateSelectedItemsCount();
    }

    // Filtrar localmente
    document.getElementById("btn-buscar-ext")?.addEventListener("click", applyExtraccionesFilter);
    document.getElementById("ext-buscar")?.addEventListener("input", applyExtraccionesFilter);
    document.getElementById("ext-area")?.addEventListener("change", applyExtraccionesFilter);
    document.getElementById("ext-piso")?.addEventListener("change", applyExtraccionesFilter);
    document.getElementById("ext-bloque")?.addEventListener("change", applyExtraccionesFilter);

    function applyExtraccionesFilter() {
        if (!inventoryDataCache.length) return;
        
        const textVal = (document.getElementById("ext-buscar")?.value || "").toLowerCase();
        
        // Obtenemos los valores IDs seleccionados (Para ubicaciones normalmente tocaría cruzar con AreaID de Inventario, 
        // pero asumiremos filtro de texto por Area y Bloque si el inventario devuelve string de ubicacion completo)
        const areaSel = document.getElementById("ext-area");
        const areaNameSelected = areaSel && areaSel.value ? areaSel.options[areaSel.selectedIndex].text.toLowerCase() : null;

        const filtered = inventoryDataCache.filter(it => {
            // Filtro textual
            const matchText = !textVal || 
                String(it.cod_inventario || "").toLowerCase().includes(textVal) ||
                String(it.descripcion || "").toLowerCase().includes(textVal) ||
                String(it.marca || "").toLowerCase().includes(textVal);
            
            // Filtro por Área si es seleccionada (asumiendo que item.ubicacion contiene el nombre del area)
            const matchLocation = !areaNameSelected || 
                String(it.ubicacion || "").toLowerCase().includes(areaNameSelected);

            return matchText && matchLocation;
        });

        renderExtraccionTable(filtered);
    }

    // Modal Events
    modalExtraerEl?.addEventListener("show.bs.modal", async () => {
        await loadExtraccionData();
    });

    // Checkbox master (Check all)
    checkAllItems?.addEventListener("change", (e) => {
        const cks = document.querySelectorAll(".item-extract-chk");
        cks.forEach(c => c.checked = e.target.checked);
        updateSelectedItemsCount();
    });

    function updateSelectedItemsCount() {
        const checked = document.querySelectorAll(".item-extract-chk:checked").length;
        if(selItemsCount) selItemsCount.textContent = checked;
    }

    document.getElementById("btn-confirmar-extraccion")?.addEventListener("click", () => {
        const checked = document.querySelectorAll(".item-extract-chk:checked").length;
        if (checked === 0) {
            notify("Por favor selecciona al menos 1 ítem del inventario para extraer.", true);
            return;
        }
        notify(`${checked} ítems seleccionados y guardados temporalmente para el acta.`);
        if(modalExtraer) modalExtraer.hide();
    });

    // 4. Botones de Acción (Vista Previa y Descargar)
    const btnVistas = document.querySelectorAll(".btn-vista-previa");
    btnVistas.forEach(btn => btn.addEventListener("click", (e) => {
        e.preventDefault();
        notify("Generando vista previa del acta en PDF...");
    }));

    const btnDescargas = document.querySelectorAll(".btn-descargar-acta");
    btnDescargas.forEach(btn => btn.addEventListener("click", (e) => {
        e.preventDefault();
        notify("Acta descargada correctamente.");
    }));
});
