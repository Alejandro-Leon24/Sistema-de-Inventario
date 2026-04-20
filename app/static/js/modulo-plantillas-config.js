const PLANTILLAS_REQUIRED_VARS = {
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
    movimiento: ["numero_acta"],
    bajas: [
        "numero_acta",
        "nombre_delegado",
        "recibido_por",
        "entregado_por",
        "rol_entrega",
        "fecha_emision",
        "tabla_dinamica",
    ],
    traspaso: [
        "numero_acta",
        "fecha_emision",
        "fecha_corte",
        "entregado_por",
        "rol_entrega",
        "facultad_entrega",
        "recibido_por",
        "rol_recibe",
        "facultad_recibe",
        "descripcion_de_bienes",
        "tabla_dinamica",
    ],
    "fin-gestion": ["numero_acta"],
    aula: ["tabla_dinamica"],
};

let plantillasEventSource = null;
let plantillasReloadTimer = null;

function scheduleReloadPlantillas(forms, delay = 350) {
    clearTimeout(plantillasReloadTimer);
    plantillasReloadTimer = setTimeout(() => {
        forms.forEach((form) => cargarEstadoPlantilla(form));
    }, delay);
}

function initPlantillasSSE(forms) {
    if (typeof EventSource === "undefined") return;
    if (plantillasEventSource) return;

    const connect = () => {
        plantillasEventSource = new EventSource("/api/events");

        plantillasEventSource.addEventListener("templates_changed", () => {
            scheduleReloadPlantillas(forms, 150);
        });

        plantillasEventSource.addEventListener("connected", () => {
            // Conexion activa.
        });

        plantillasEventSource.onerror = () => {
            try {
                plantillasEventSource.close();
            } catch (_err) {
                // noop
            }
            plantillasEventSource = null;
            setTimeout(connect, 2500);
        };
    };

    connect();
}

// Compatibilidad para plantillas antiguas de Entrega.
const PLANTILLAS_REQUIRED_ALIASES = {
    entrega: {
        rol_recibe: ["usuario_final", "recibe_custodio"],
        area_trabajo: ["ubicacion", "ubicacion_entrega"],
    },
};

const COMMON_REQUIRED_ALIASES = {
    numero_acta: ["numeroacta", "numero_de_acta", "nro_acta", "nroacta", "n_acta"],
};

function normalizeVarName(value) {
    let v = String(value || "").trim();
    if (!v) return "";

    // Compatibilidad con expresiones de plantilla: {{ var|filtro }} o {{ var.attr }}.
    v = v.replace(/^\{\{\s*/, "").replace(/\s*\}\}$/, "");
    v = v.split("|")[0].trim();
    v = v.replace(/([a-z0-9])([A-Z])/g, "$1_$2");
    v = v.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    v = v.toLowerCase();
    v = v.replace(/[\s-]+/g, "_");
    v = v.replace(/_+/g, "_").replace(/^_+|_+$/g, "");

    // Regla directa: cualquier variante que combine 'numero' + 'acta' se trata igual.
    const compact = v.replace(/[^a-z0-9]/g, "");
    if (compact.includes("numero") && compact.includes("acta")) return "numero_acta";

    if (v === "numero_de_acta") return "numero_acta";
    return v;
}

function canonicalVarKey(value) {
    const normalized = normalizeVarName(value);
    if (!normalized) return "";

    // Si viene namespaced (ej. data.numero_acta), usamos la ultima parte util.
    const lastSegment = normalized.split(".").pop() || normalized;
    return lastSegment.replace(/[^a-z0-9]/g, "");
}

function varsBasicas(lista) {
    const seen = new Set();
    return (lista || [])
        .filter((v) => typeof v === "string")
        .map((v) => normalizeVarName(v))
        .filter((v) => {
            if (!v) return false;
            if (v.startsWith("item.")) return false;
            if (v.startsWith("col.")) return false;
            if (v.startsWith("fila.")) return false;
            if (["tabla_items", "tabla_columnas", "tabla_filas", "col.label", "celda"].includes(v)) {
                return false;
            }
            const key = canonicalVarKey(v);
            if (!key || seen.has(key)) return false;
            seen.add(key);
            return true;
        });
}

function isRequiredVarPresent(tipo, requiredVar, foundSet) {
    const normalizedRequired = normalizeVarName(requiredVar);
    const requiredKey = canonicalVarKey(normalizedRequired);

    const foundKeys = new Set(Array.from(foundSet).map((v) => canonicalVarKey(v)).filter(Boolean));
    if (foundSet.has(normalizedRequired) || foundKeys.has(requiredKey)) return true;

    const commonAliases = COMMON_REQUIRED_ALIASES[normalizedRequired] || [];
    const aliasMap = PLANTILLAS_REQUIRED_ALIASES[tipo] || {};
    const aliases = [...commonAliases, ...(aliasMap[normalizedRequired] || [])];
    return aliases.some((alias) => {
        const normalizedAlias = normalizeVarName(alias);
        const aliasKey = canonicalVarKey(normalizedAlias);
        return foundSet.has(normalizedAlias) || foundKeys.has(aliasKey);
    });
}

function pintarEstadoVariables(tipo, container, foundVariables = [], plantillaExiste = false) {
    const required = PLANTILLAS_REQUIRED_VARS[tipo] || [];
    const basicas = varsBasicas(foundVariables);
    const basicasSet = new Set(basicas);
    const allFoundSet = new Set(
        (foundVariables || [])
            .filter((v) => typeof v === "string")
            .map((v) => normalizeVarName(v))
            .filter(Boolean)
    );

    container.innerHTML = '<strong class="mb-2 d-block text-dark mt-2">Variables Obligatorias:</strong>';

    if (!required.length) {
        container.innerHTML += '<span class="text-muted small">Aún no se ha definido esquema para esta acta.</span>';
        return;
    }

    const faltantes = [];

    required.forEach((v) => {
        const span = document.createElement("span");
        span.textContent = `{{${v}}}`;
        const ok = plantillaExiste && isRequiredVarPresent(tipo, v, allFoundSet.size ? allFoundSet : basicasSet);
        span.className = ok
            ? "badge bg-success me-2 mb-2 p-2 shadow-sm"
            : "badge bg-danger me-2 mb-2 p-2 shadow-sm";
        if (!ok) faltantes.push(v);
        container.appendChild(span);
    });

    const extras = basicas.filter((v) => {
        return !required.some((req) => isRequiredVarPresent(tipo, req, new Set([v])));
    });
    if (plantillaExiste && extras.length > 0) {
        const extra = document.createElement("span");
        extra.className = "badge bg-warning text-dark me-2 mb-2 p-2 shadow-sm border border-secondary";
        extra.innerHTML = `<i class="bi bi-stars me-1"></i>${extras.length}+`;
        extra.title = `Variables nuevas detectadas: ${extras.map((v) => `{{${v}}}`).join(", ")}`;
        container.appendChild(extra);
    }

    if (plantillaExiste) {
        const msg = document.createElement("div");
        msg.className = faltantes.length
            ? "text-danger small mt-2 fw-medium"
            : "text-success small mt-2 fw-medium";
        msg.innerHTML = faltantes.length
            ? '<i class="bi bi-exclamation-triangle-fill me-1"></i>Faltan variables obligatorias en el Word. Su acta podría fallar.'
            : '<i class="bi bi-check-circle-fill me-1"></i>Todos los campos obligatorios fueron encontrados exitosamente.';
        container.appendChild(msg);
    }
}

function pintarBotonSubida(form, existePlantilla) {
    const btn = form.querySelector('button[type="submit"]');
    if (!btn) return;
    if (existePlantilla) {
        btn.innerHTML = '<i class="bi bi-arrow-repeat me-1"></i>Actualizar Plantilla Cargada';
        btn.classList.remove("btn-outline-primary");
        btn.classList.add("btn-outline-success");
    } else {
        btn.innerHTML = '<i class="bi bi-upload me-1"></i>Subir y Analizar';
        btn.classList.remove("btn-outline-success");
        btn.classList.add("btn-outline-primary");
    }
}

function getOrCreatePlantillaInfoNode(form) {
    let info = form.querySelector(".plantilla-info");
    if (info) return info;

    info = document.createElement("div");
    info.className = "plantilla-info mt-2 d-flex align-items-center justify-content-between gap-2";
    form.appendChild(info);
    return info;
}

function pintarInfoPlantilla(form, tipo, existePlantilla) {
    const info = getOrCreatePlantillaInfoNode(form);
    if (!info) return;

    if (!existePlantilla) {
        info.innerHTML = '<span class="small text-muted">No hay plantilla cargada.</span>';
        return;
    }

    const safeTipo = encodeURIComponent(String(tipo || ""));
    info.innerHTML = `
        <span class="small text-success fw-semibold"><i class="bi bi-check-circle-fill me-1"></i>Plantilla cargada</span>
        <a class="btn btn-sm btn-outline-secondary" href="/api/plantillas/descargar?tipo=${safeTipo}">
            <i class="bi bi-download me-1"></i>Descargar actual
        </a>
    `;
}

async function cargarEstadoPlantilla(form) {
    const tipo = form.getAttribute("data-tipo");
    const container = form.nextElementSibling;
    if (!tipo || !container) return;

    try {
        const res = await fetch(`/api/plantillas/estado?tipo=${encodeURIComponent(tipo)}`);
        const payload = await res.json();
        if (payload.success && payload.existe) {
            pintarEstadoVariables(tipo, container, payload.variables || [], true);
            pintarBotonSubida(form, true);
            pintarInfoPlantilla(form, tipo, true);
        } else {
            pintarEstadoVariables(tipo, container, [], false);
            pintarBotonSubida(form, false);
            pintarInfoPlantilla(form, tipo, false);
        }
    } catch (err) {
        console.error("Error cargando estado de plantilla", err);
        pintarEstadoVariables(tipo, container, [], false);
        pintarInfoPlantilla(form, tipo, false);
    }
}

async function subirPlantilla(form) {
    const tipo = form.getAttribute("data-tipo");
    const input = form.querySelector('input[type="file"]');
    const container = form.nextElementSibling;
    const btn = form.querySelector('button[type="submit"]');
    if (!tipo || !input || !container || !btn) return;

    if (!input.files.length) {
        notify("Seleccione el documento primero.", true);
        return;
    }

    const originalText = btn.innerHTML;
    btn.innerHTML = '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Analizando...';
    btn.disabled = true;

    const formData = new FormData();
    formData.append("documento", input.files[0]);
    formData.append("tipo", tipo);

    try {
        const response = await fetch("/api/plantillas/upload", {
            method: "POST",
            body: formData,
        });
        const result = await response.json();

        if (result.success) {
            notify("Plantilla analizada y subida correctamente", false);
            pintarEstadoVariables(tipo, container, result.variables || [], true);
            pintarBotonSubida(form, true);
            pintarInfoPlantilla(form, tipo, true);
        } else {
            notify(result.error || "Error al subir plantilla", true);
            await cargarEstadoPlantilla(form);
        }
    } catch (err) {
        console.error(err);
        notify("Ocurrió un error de red al subir archivo.", true);
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalText;
        input.value = "";
    }
}

document.addEventListener("DOMContentLoaded", () => {
    const forms = document.querySelectorAll(".form-plantilla");
    forms.forEach((form) => {
        cargarEstadoPlantilla(form);
        form.addEventListener("submit", async (e) => {
            e.preventDefault();
            await subirPlantilla(form);
        });
    });

    initPlantillasSSE(forms);
});
