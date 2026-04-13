window.api = {
    async get(url, options = {}) {
        const response = await fetch(url, options);
        const payload = await parseResponsePayload(response);
        if (!response.ok) {
            throw buildApiError(payload, response.status);
        }
        return payload;
    },
    async send(url, method, body) {
        const response = await fetch(url, {
            method,
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
        });
        const payload = await parseResponsePayload(response);
        if (!response.ok) {
            throw buildApiError(payload, response.status);
        }
        return payload;
    },
};

async function parseResponsePayload(response) {
    const contentType = String(response.headers.get("content-type") || "").toLowerCase();
    const rawText = await response.text();

    if (!rawText) {
        return {};
    }

    if (contentType.includes("application/json")) {
        try {
            return JSON.parse(rawText);
        } catch (_error) {
            return {
                error: "Respuesta JSON inválida del servidor.",
                raw: rawText.slice(0, 400),
            };
        }
    }

    // Fallback para respuestas HTML (404/500) u otros formatos.
    return {
        error: "El servidor devolvió una respuesta no JSON. Si agregaste endpoints nuevos, reinicia el servidor.",
        raw: rawText.slice(0, 400),
    };
}

function buildApiError(payload, status) {
    const message = payload?.error || `Error de servidor (${status})`;
    const error = new Error(message);
    error.payload = payload || {};
    error.status = status;
    return error;
}
