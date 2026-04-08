window.api = {
    async get(url, options = {}) {
        const response = await fetch(url, options);
        const payload = await response.json();
        if (!response.ok) {
            const error = new Error(payload.error || "Error de servidor");
            error.payload = payload;
            error.status = response.status;
            throw error;
        }
        return payload;
    },
    async send(url, method, body) {
        const response = await fetch(url, {
            method,
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
        });
        const payload = await response.json();
        if (!response.ok) {
            const error = new Error(payload.error || "Error de servidor");
            error.payload = payload;
            error.status = response.status;
            throw error;
        }
        return payload;
    },
};
