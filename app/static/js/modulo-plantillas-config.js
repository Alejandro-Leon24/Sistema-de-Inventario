const PLANTILLAS_REQUIRED_VARS = {
    entrega: [
        "fecha_corte",
        "fecha_emision",
        "accion_personal",
        "entregado_por",
        "recibido_por",
        "rol_recibe",
        "area_trabajo",
    ],
    recepcion: [],
    movimiento: [],
    bajas: [],
    traspaso: [],
    "fin-gestion": [],
    aula: [],
};

// Compatibilidad para plantillas antiguas de Entrega.
const PLANTILLAS_REQUIRED_ALIASES = {
    entrega: {
        rol_recibe: ["usuario_final", "recibe_custodio"],
        area_trabajo: ["ubicacion", "ubicacion_entrega"],
    },
};

function normalizeVarName(value) {
    return String(value || "").trim().toLowerCase();
}

function varsBasicas(lista) {
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
            return true;
        });
}

function isRequiredVarPresent(tipo, requiredVar, foundSet) {
    const normalizedRequired = normalizeVarName(requiredVar);
    if (foundSet.has(normalizedRequired)) return true;

    const aliasMap = PLANTILLAS_REQUIRED_ALIASES[tipo] || {};
    const aliases = aliasMap[normalizedRequired] || [];
    return aliases.some((alias) => foundSet.has(normalizeVarName(alias)));
}

function pintarEstadoVariables(tipo, container, foundVariables = [], plantillaExiste = false) {
    const required = PLANTILLAS_REQUIRED_VARS[tipo] || [];
    const basicas = varsBasicas(foundVariables);
    const basicasSet = new Set(basicas);

    container.innerHTML = '<strong class="mb-2 d-block text-dark mt-2">Variables Obligatorias:</strong>';

    if (!required.length) {
        container.innerHTML += '<span class="text-muted small">Aún no se ha definido esquema para esta acta.</span>';
        return;
    }

    const faltantes = [];

    required.forEach((v) => {
        const span = document.createElement("span");
        span.textContent = `{{${v}}}`;
        const ok = plantillaExiste && isRequiredVarPresent(tipo, v, basicasSet);
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
        } else {
            pintarEstadoVariables(tipo, container, [], false);
            pintarBotonSubida(form, false);
        }
    } catch (err) {
        console.error("Error cargando estado de plantilla", err);
        pintarEstadoVariables(tipo, container, [], false);
    }
}

async function subirPlantilla(form) {
    const tipo = form.getAttribute("data-tipo");
    const input = form.querySelector('input[type="file"]');
    const container = form.nextElementSibling;
    const btn = form.querySelector('button[type="submit"]');
    if (!tipo || !input || !container || !btn) return;

    if (!input.files.length) {
        notify("Seleccione el documento primero.", "warning");
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
            notify("Plantilla analizada y subida correctamente", "success");
            pintarEstadoVariables(tipo, container, result.variables || [], true);
            pintarBotonSubida(form, true);
        } else {
            notify(result.error || "Error al subir plantilla", "danger");
            await cargarEstadoPlantilla(form);
        }
    } catch (err) {
        console.error(err);
        notify("Ocurrió un error de red al subir archivo.", "danger");
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
});
