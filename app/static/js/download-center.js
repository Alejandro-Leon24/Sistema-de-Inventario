(function () {
    const ACTIVE_JOB_KEY = "inventario_area_report_active_job";
    const isInformePage = String(window.location.pathname || "").startsWith("/informe");
    if (isInformePage) return;

    const panel = document.getElementById("download-progress-floating");
    if (!panel || typeof fetch === "undefined") return;

    const titleEl = document.getElementById("descarga-progress-title");
    const msgEl = document.getElementById("descarga-progress-message");
    const progressBar = document.getElementById("descarga-progress-bar");
    const progressText = document.getElementById("descarga-progress-text");
    const actions = document.getElementById("download-progress-actions");
    const pauseBtn = document.getElementById("btn-download-progress-pause");
    const cancelBtn = document.getElementById("btn-download-progress-cancel");
    const closeBtn = document.getElementById("btn-download-progress-close");

    let eventSource = null;
    let activeJobId = String(localStorage.getItem(ACTIVE_JOB_KEY) || "").trim() || null;
    let paused = false;

    function setVisible(visible) {
        panel.classList.toggle("d-none", !visible);
    }

    function setTitle(text) {
        if (titleEl) titleEl.textContent = String(text || "").trim() || "Generando acta";
    }

    function setMessage(text) {
        if (msgEl) msgEl.textContent = String(text || "").trim() || "Preparando descarga...";
    }

    function setProgress(value) {
        const safe = Math.max(0, Math.min(100, Number(value || 0)));
        if (progressBar) progressBar.style.width = `${safe}%`;
        if (progressText) progressText.textContent = `${safe}%`;
    }

    function setActionsVisible(visible) {
        if (actions) actions.classList.toggle("d-none", !visible);
    }

    function setPausedButton(isPaused) {
        paused = Boolean(isPaused);
        if (!pauseBtn) return;
        if (paused) {
            pauseBtn.innerHTML = '<i class="bi bi-play-fill me-1"></i>Reanudar';
        } else {
            pauseBtn.innerHTML = '<i class="bi bi-pause-fill me-1"></i>Pausar';
        }
    }

    function setActiveJob(jobId) {
        activeJobId = String(jobId || "").trim() || null;
        if (activeJobId) {
            localStorage.setItem(ACTIVE_JOB_KEY, activeJobId);
        } else {
            localStorage.removeItem(ACTIVE_JOB_KEY);
        }
    }

    function startDownload(path) {
        const filePath = String(path || "").trim();
        if (!filePath) return;
        const a = document.createElement("a");
        a.href = `/api/descargar?path=${encodeURIComponent(filePath)}`;
        a.style.display = "none";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }

    async function controlJob(action) {
        if (!activeJobId) return;
        try {
            const response = await fetch(`/api/informes/areas/jobs/${encodeURIComponent(activeJobId)}/control`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ action }),
            });
            const payload = await response.json();
            if (!response.ok || !payload.success) return;

            if (action === "pause") setMessage("Pausa solicitada...");
            if (action === "resume") setMessage("Reanudando...");
            if (action === "cancel") setMessage("Cancelación solicitada...");
        } catch (_err) {
            // noop
        }
    }

    function handleProgress(payload) {
        const jobId = String(payload.job_id || "").trim();
        if (!jobId) return;
        if (activeJobId && jobId !== activeJobId) return;

        setActiveJob(jobId);
        setVisible(true);
        setActionsVisible(true);
        setTitle("Generando informes por área");
        setProgress(payload.progress || 0);
        setMessage(payload.message || "Procesando informe...");
    }

    function handlePaused(payload) {
        const jobId = String(payload.job_id || "").trim();
        if (!activeJobId || jobId !== activeJobId) return;
        setPausedButton(true);
        setMessage(payload.message || "Generación en pausa.");
    }

    function handleResumed(payload) {
        const jobId = String(payload.job_id || "").trim();
        if (!activeJobId || jobId !== activeJobId) return;
        setPausedButton(false);
        setMessage(payload.message || "Generación reanudada.");
    }

    function handleReady(payload) {
        const jobId = String(payload.job_id || "").trim();
        if (!activeJobId || jobId !== activeJobId) return;

        setProgress(100);
        setTitle("Descarga lista");
        const downloadPath = String(payload.download_path || payload.zip_path || "").trim();
        const downloadKind = String(payload.download_kind || (payload.zip_path ? "zip" : "docx")).toLowerCase();
        setMessage(downloadKind === "zip" ? "Paquete ZIP generado correctamente." : "Documento Word generado correctamente.");
        setActionsVisible(false);
        setPausedButton(false);
        if (downloadPath) startDownload(downloadPath);
        setActiveJob(null);
        setTimeout(() => setVisible(false), 1800);
    }

    function handleError(payload) {
        const jobId = String(payload.job_id || "").trim();
        if (!activeJobId || jobId !== activeJobId) return;

        setTitle("Error de generación");
        setMessage(payload.error || "Ocurrió un error durante la generación.");
        setActionsVisible(false);
        setPausedButton(false);
        setActiveJob(null);
    }

    function handleCancelled(payload) {
        const jobId = String(payload.job_id || "").trim();
        if (!activeJobId || jobId !== activeJobId) return;

        setTitle("Lote cancelado");
        setMessage(payload.message || "Generación cancelada por el usuario.");
        setActionsVisible(false);
        setPausedButton(false);
        setActiveJob(null);
        setTimeout(() => setVisible(false), 1200);
    }

    function connectSSE() {
        if (typeof EventSource === "undefined") return;
        if (eventSource) return;

        const es = new EventSource("/api/events");
        eventSource = es;

        es.addEventListener("areas_reports_progress", (event) => {
            let payload = {};
            try {
                payload = JSON.parse(event.data || "{}");
            } catch (_err) {
                payload = {};
            }
            handleProgress(payload);
        });

        es.addEventListener("areas_reports_paused", (event) => {
            let payload = {};
            try {
                payload = JSON.parse(event.data || "{}");
            } catch (_err) {
                payload = {};
            }
            handlePaused(payload);
        });

        es.addEventListener("areas_reports_resumed", (event) => {
            let payload = {};
            try {
                payload = JSON.parse(event.data || "{}");
            } catch (_err) {
                payload = {};
            }
            handleResumed(payload);
        });

        es.addEventListener("areas_reports_ready", (event) => {
            let payload = {};
            try {
                payload = JSON.parse(event.data || "{}");
            } catch (_err) {
                payload = {};
            }
            handleReady(payload);
        });

        es.addEventListener("areas_reports_error", (event) => {
            let payload = {};
            try {
                payload = JSON.parse(event.data || "{}");
            } catch (_err) {
                payload = {};
            }
            handleError(payload);
        });

        es.addEventListener("areas_reports_cancelled", (event) => {
            let payload = {};
            try {
                payload = JSON.parse(event.data || "{}");
            } catch (_err) {
                payload = {};
            }
            handleCancelled(payload);
        });

        es.onerror = () => {
            try {
                es.close();
            } catch (_err) {
                // noop
            }
            eventSource = null;
            setTimeout(connectSSE, 2500);
        };
    }

    async function hydrateFromBackend() {
        if (!activeJobId) {
            try {
                const response = await fetch("/api/informes/areas/jobs?active=1");
                const payload = await response.json();
                if (!response.ok || !payload.success) return;
                const first = Array.isArray(payload.jobs) ? payload.jobs[0] : null;
                if (!first || !first.job_id) return;
                setActiveJob(first.job_id);
            } catch (_err) {
                return;
            }
        }

        if (!activeJobId) return;

        try {
            const response = await fetch(`/api/informes/areas/jobs/${encodeURIComponent(activeJobId)}`);
            const payload = await response.json();
            if (!response.ok || !payload.success || !payload.data) {
                setActiveJob(null);
                return;
            }
            const data = payload.data;
            const status = String(data.status || "");
            if (!["queued", "running", "paused"].includes(status)) {
                setActiveJob(null);
                return;
            }

            setVisible(true);
            setActionsVisible(true);
            setTitle("Generando informes por área");
            setProgress(data.progress || 0);
            setMessage(data.message || (status === "paused" ? "Generación en pausa." : "Procesando informe..."));
            setPausedButton(status === "paused" || Boolean(data.pause_requested));
        } catch (_err) {
            // noop
        }
    }

    if (closeBtn) {
        closeBtn.addEventListener("click", () => {
            setVisible(false);
        });
    }
    if (pauseBtn) {
        pauseBtn.addEventListener("click", async () => {
            await controlJob(paused ? "resume" : "pause");
        });
    }
    if (cancelBtn) {
        cancelBtn.addEventListener("click", async () => {
            await controlJob("cancel");
        });
    }

    hydrateFromBackend();
    connectSSE();
})();
