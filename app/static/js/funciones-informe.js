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
    const tabs = document.querySelectorAll(".settings-menu-btn");
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
    
    // Columnas seleccionables
    const availableColumns = [
        { id: "cod_inventario", label: "CÓDIGO INV." },
        { id: "cod_esbye", label: "CÓD. ESBYE" },
        { id: "cuenta", label: "CUENTA" },
        { id: "cantidad", label: "CANT" },
        { id: "descripcion", label: "DESCRIPCIÓN" },
        { id: "marca", label: "MARCA" },
        { id: "modelo", label: "MODELO" },
        { id: "serie", label: "SERIE" },
        { id: "estado", label: "ESTADO" },
        { id: "ubicacion", label: "UBICACIÓN" },
        { id: "fecha_adquisicion", label: "FECHA DE ADQUISICIÓN" },
        { id: "valor", label: "VALOR" },
        { id: "usuario_final", label: "USUARIO FINAL" },
        { id: "observacion", label: "OBSERVACIÓN" },
        { isSubtitle: true, label: "ESBYE" },
        { id: "descripcion_esbye", label: "DESCRIPCIÓN" },
        { id: "marca_esbye", label: "MARCA" },
        { id: "modelo_esbye", label: "MODELO" },
        { id: "serie_esbye", label: "SERIE" },
        { id: "fecha_adquisicion_esbye", label: "FECHA" },
        { id: "valor_esbye", label: "VALOR" },
        { id: "ubicacion_esbye", label: "UBICACIÓN" },
        { id: "observacion_esbye", label: "OBSERVACIÓN" }
    ];
    let selectedColumns = ["cod_inventario", "descripcion", "marca", "modelo", "serie", "estado"];

    // Obtener preferencias del usuario
    try {
        const prefRes = await api.get('/api/preferencias');
        if (prefRes.data && prefRes.data.extraccion_columnas) {
            selectedColumns = prefRes.data.extraccion_columnas;
        }
    } catch (e) {
        console.warn("No se pudieron cargar las preferencias de columnas", e);
    }

    const colSelector = document.getElementById("column-selector");
    if (colSelector) {
        availableColumns.forEach(c => {
            const div = document.createElement("div");
            if (c.isSubtitle) {
                div.className = "list-group-item bg-light fw-bold text-center border-top border-bottom py-1 mt-2 text-primary small";
                div.textContent = c.label;
            } else {
                div.className = "list-group-item list-group-item-action d-flex align-items-center py-2";
                const isChecked = selectedColumns.includes(c.id) ? "checked" : "";
                div.innerHTML = `
                    <input class="form-check-input me-3 col-chk" type="checkbox" value="${c.id}" id="chk-col-${c.id}" ${isChecked}>
                    <label class="form-check-label w-100 fw-medium small mb-0 cursor-pointer" for="chk-col-${c.id}" style="cursor:pointer;">${c.label}</label>
                `;
            }
            colSelector.appendChild(div);
        });

        document.querySelectorAll(".col-chk").forEach(chk => {
            chk.addEventListener("change", async (e) => {
                if (e.target.checked) {
                    selectedColumns.push(e.target.value);
                } else {
                    selectedColumns = selectedColumns.filter(id => id !== e.target.value);
                }
                applyExtraccionesFilter();
                
                // Guardar en base de datos la nueva preferencia
                try {
                    await api.send('/api/preferencias', 'PATCH', {
                        pref_key: 'extraccion_columnas',
                        pref_value: selectedColumns
                    });
                } catch (err) {
                    console.error("Error guardando preferencias de columnas", err);
                }
            });
        });
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
    const theadExtraccion = document.getElementById("thead-extraccion-tr");
    const checkAllItems = document.getElementById("chk-all-items");
    const selItemsCount = document.getElementById("items-seleccionados-count");
    const contTabla = document.getElementById("contenedor-tabla-extraccion");
    const msgSinCols = document.getElementById("mensaje-sin-columnas");

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
        // Actualizar thead
        if (theadExtraccion) {
            theadExtraccion.innerHTML = `<th class="text-center" style="width: 50px;"><input class="form-check-input" type="checkbox" id="chk-all-items"></th>`;
            selectedColumns.forEach(scId => {
                const c = availableColumns.find(col => col.id === scId);
                if (c) {
                    const th = document.createElement("th");
                    th.textContent = c.label;
                    theadExtraccion.appendChild(th);
                }
            });
            // Re-bind master checkbox listener
            document.getElementById("chk-all-items")?.addEventListener("change", (e) => {
                const cks = document.querySelectorAll(".item-extract-chk");
                cks.forEach(c => c.checked = e.target.checked);
                updateSelectedItemsCount();
            });
        }

        tbodyExtraccion.innerHTML = "";
        
        // Validación de al menos 1 columna seleccionada
        if (selectedColumns.length === 0) {
            contTabla.classList.add("d-none");
            msgSinCols.classList.remove("d-none");
            return;
        } else {
            contTabla.classList.remove("d-none");
            msgSinCols.classList.add("d-none");
        }

        if (items.length === 0) {
            tbodyExtraccion.innerHTML = `<tr><td colspan="${selectedColumns.length + 1}" class="text-center text-muted">No se encontraron equipos bajo estos filtros.</td></tr>`;
            return;
        }

        items.forEach(item => {
            const tr = document.createElement("tr");
            let html = `
                <td class="text-center">
                    <input class="form-check-input item-extract-chk" type="checkbox" value="${item.id}">
                </td>`;
            
            selectedColumns.forEach(scId => {
                const c = availableColumns.find(col => col.id === scId);
                if (c) {
                    if (c.id === 'estado') {
                        html += `<td><span class="badge bg-secondary">${item.estado || '-'}</span></td>`;
                    } else {
                        html += `<td>${item[c.id] || '-'}</td>`;
                    }
                }
            });
            tr.innerHTML = html;
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
        applyExtraccionesFilter(); // Disparar filtro al abrir para ocultar tabla si inicia vacío
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
        
        // Determinar qué tab está activo para mostrar el resumen en su contenedor correspondiente
        const activeTabTarget = document.querySelector(".settings-menu-btn.active")?.getAttribute("data-target");
        if(activeTabTarget) {
            const targetDiv = document.querySelector(`#${activeTabTarget} .resultado-tabla-container`);
            if (targetDiv) {
                targetDiv.innerHTML = `
                    <div class="alert alert-success d-inline-block p-2 px-4 shadow-sm mb-0 animate__animated animate__fadeIn">
                        <i class="bi bi-check2-circle me-2"></i> 
                        Se extrajeron <strong>${checked}</strong> filas y <strong>${selectedColumns.length}</strong> columnas exitosamente.
                    </div>
                `;
            }
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


// ==========================================
// M�DULO DE PLANTILLAS Y PERSONAL DIN�MICO
// ==========================================

async function cargarPersonalDatalist() {
    try {
        const respuesta = await fetch('/api/personal');
        const data = await respuesta.json();
        
        if (data.success) {
            const datalist = document.getElementById('lista-personal');
            if (!datalist) return;
            
            datalist.innerHTML = ''; // Limpiar
            const empleados = data.data || [];
            
            empleados.forEach(emp => {
                const opcion = document.createElement('option');
                opcion.value = emp.nombre;
                if(emp.cargo) {
                    opcion.textContent = emp.cargo;
                }
                datalist.appendChild(opcion);
            });
            
            // Asignar el datalist autom�ticamente a todos los campos pertinentes conocidos si no lo tienen
            const camposPersonal = document.querySelectorAll('input[name="entregado_por"], input[name="recibido_por"], input[name="usuario_final"], input[id*="entregado"], input[id*="recibido"], input[name="administradora"]');
            camposPersonal.forEach(input => {
                input.setAttribute('list', 'lista-personal');
                input.setAttribute('autocomplete', 'off'); // desactivar el estandar pata usar el datalist
            });
        }
    } catch (e) {
        console.error('Error cargando personal para autocompletado', e);
    }
}

document.addEventListener('DOMContentLoaded', () => {
    // Cargar listas
    cargarPersonalDatalist();
});


// ==========================================
// M�DULO DE VISTA PREVIA EN VIVO
// ==========================================

let _typingTimer;
const _doneTypingInterval = 1000; // delay to execute server call

function initLivePreview() {
    // Listen to changes on inputs to trigger a debounce preview request
    const formElement = document.getElementById('report-form');
    if (!formElement) return;

    formElement.addEventListener('input', () => {
        clearTimeout(_typingTimer);
        _typingTimer = setTimeout(triggerPreview, _doneTypingInterval);
    });

    // Handle initial selects like document types
    formElement.addEventListener('change', () => {
        clearTimeout(_typingTimer);
        _typingTimer = setTimeout(triggerPreview, 250);
    });
}

async function triggerPreview() {
    const iframe = document.getElementById('preview-iframe');
    const placeholder = document.getElementById('preview-placeholder');
    if (!iframe || !placeholder) return;

    // Collect Data identical to final submit but marked 'vista_previa'
    const formData = new FormData(document.getElementById('report-form'));
    const data = Object.fromEntries(formData.entries());
    
    // Si no hay tipo de documento seleccionado o faltan campos obligatorios mínimos, abortar
    if (!data.tipo_documento) return;

    data.table_data = _dataRows || []; // _dataRows from the extract logic if available
    data.vista_previa = true; 

    try {
        const payload = await api.send('/api/informes/generar', 'POST', data);
        if (payload.pdf_path) {
            // Un-hide iframe and hide placeholder
            placeholder.classList.add('d-none');
            iframe.classList.remove('d-none');
            
            // Set URL
            updateIframe(payload.pdf_path);
        }
    } catch(err) {
        // Silently fail for preview, logic could just be missing an essential field that server rejects
        console.warn('Live preview stopped:', err.message);
    }
}

// Inicializar previsualización dentro de loaded
document.addEventListener("DOMContentLoaded", () => {
    initLivePreview();
});


// Extraer nombre base y servir
// Reemplazar la URL del iframe con la ruta del servidor proxy de FLASK
async function updateIframe(pdfPath) {
    if (!pdfPath) return; // Si ocurre algo malo
    
    // Obtener sólo el nombre del archivo (todo lo que está después del último slash o backslash)
    const fileName = pdfPath.split(/[\\/]/).pop();
    const pdfUrl = '/files/' + encodeURIComponent(fileName);

    const iframe = document.getElementById('preview-iframe');
    const placeholder = document.getElementById('preview-placeholder');
    
    placeholder.classList.add('d-none');
    iframe.classList.remove('d-none');
    
    // Anexar un timestamp para evadir caché del navegador
    iframe.src = pdfUrl + '?t=' + new Date().getTime();
}
