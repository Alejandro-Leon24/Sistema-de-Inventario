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
};
