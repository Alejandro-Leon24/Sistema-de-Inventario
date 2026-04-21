window.appHelpers = {
    escapeHtmlText(value) {
        return String(value || "")
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;")
            .replaceAll("'", "&#39;");
    },

    async loadStructure(apiClient, { sortNatural = false, includeAreaDetails = false } = {}) {
        const includeDetailsParam = includeAreaDetails ? "1" : "0";
        const response = await apiClient.get(`/api/estructura?include_details=${includeDetailsParam}`);
        const structure = Array.isArray(response.data) ? response.data : [];

        if (!sortNatural) {
            return structure;
        }

        structure.sort((a, b) => a.nombre.localeCompare(b.nombre, undefined, { numeric: true, sensitivity: "base" }));
        structure.forEach((block) => {
            if (!Array.isArray(block.pisos)) {
                return;
            }
            block.pisos.sort((a, b) => a.nombre.localeCompare(b.nombre, undefined, { numeric: true, sensitivity: "base" }));
            block.pisos.forEach((piso) => {
                if (!Array.isArray(piso.areas)) {
                    return;
                }
                piso.areas.sort((a, b) => a.nombre.localeCompare(b.nombre, undefined, { numeric: true, sensitivity: "base" }));
            });
        });

        return structure;
    },

    async loadPreferences(apiClient) {
        const response = await apiClient.get("/api/preferencias");
        return response.data || {};
    },

    parseDecimalWithComma(value) {
        const raw = String(value ?? "").trim();
        if (!raw) return null;
        
        // Si tiene ambos, asumimos formato local (1.234,56)
        if (raw.includes(",") && raw.includes(".")) {
            const normalized = raw.replace(/\./g, "").replace(",", ".");
            const number = Number(normalized);
            return Number.isFinite(number) ? number : null;
        }
        
        // Si solo tiene coma, es el separador decimal
        if (raw.includes(",")) {
            const normalized = raw.replace(",", ".");
            const number = Number(normalized);
            return Number.isFinite(number) ? number : null;
        }
        
        // Si solo tiene punto (o nada), el Number() de JS ya lo trata como decimal (formato 1234.56)
        const number = Number(raw);
        return Number.isFinite(number) ? number : null;
    }
};

// Funciones globales para oninput
window.validarEnteros = function(input) {
    let value = String(input.value || "").replace(/[^0-9]/g, "");
    if (value.length > 1) {
        value = value.replace(/^0+/, "") || "0";
    }
    input.value = value;
};

window.validarNumerosComas = function(input) {
    let value = String(input.value || "").replace(/[^0-9,]/g, "");
    const firstComma = value.indexOf(",");
    if (firstComma !== -1) {
        value = value.slice(0, firstComma + 1) + value.slice(firstComma + 1).replace(/,/g, "");
    }
    input.value = value;
};
