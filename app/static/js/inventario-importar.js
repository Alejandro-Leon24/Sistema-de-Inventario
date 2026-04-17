/**
 * inventario-importar.js
 * Flujo de importacion Excel en 3 pasos con revision por lotes de 20 filas.
 */
(function () {
    "use strict";

    const CHUNK_SIZE = 20;
    const CANONICAL_FIELDS = [
        { value: "", label: "- Ignorar columna -" },
        { value: "cod_inventario", label: "Cod Inv." },
        { value: "cod_esbye", label: "Cod. ESBYE" },
        { value: "cuenta", label: "Cuenta" },
        { value: "cantidad", label: "Cantidad" },
        { value: "descripcion", label: "Descripcion" },
        { value: "marca", label: "Marca" },
        { value: "modelo", label: "Modelo" },
        { value: "serie", label: "Serie" },
        { value: "estado", label: "Estado" },
        { value: "ubicacion", label: "Ubicacion" },
        { value: "fecha_adquisicion", label: "Fecha Adquisicion" },
        { value: "valor", label: "Valor" },
        { value: "usuario_final", label: "Usuario Final" },
        { value: "observacion", label: "Observacion" },
        { value: "descripcion_esbye", label: "Descripcion ESBYE" },
        { value: "marca_esbye", label: "Marca ESBYE" },
        { value: "modelo_esbye", label: "Modelo ESBYE" },
        { value: "serie_esbye", label: "Serie ESBYE" },
        { value: "fecha_adquisicion_esbye", label: "Fecha ESBYE" },
        { value: "valor_esbye", label: "Valor ESBYE" },
        { value: "ubicacion_esbye", label: "Ubicacion ESBYE" },
        { value: "observacion_esbye", label: "Observacion ESBYE" },
    ];

    const REVIEW_TABLE_FIELDS = [
        "cod_inventario",
        "cod_esbye",
        "cuenta",
        "cantidad",
        "descripcion",
        "marca",
        "modelo",
        "serie",
        "estado",
        "ubicacion",
        "fecha_adquisicion",
        "valor",
        "usuario_final",
        "observacion",
        "descripcion_esbye",
        "marca_esbye",
        "modelo_esbye",
        "serie_esbye",
        "fecha_adquisicion_esbye",
        "valor_esbye",
        "observacion_esbye",
    ];

    const EDITABLE_REVIEW_FIELDS = new Set(REVIEW_TABLE_FIELDS);
    const BASE_TO_ESBYE_FIELD = {
        descripcion: "descripcion_esbye",
        marca: "marca_esbye",
        modelo: "modelo_esbye",
        serie: "serie_esbye",
        fecha_adquisicion: "fecha_adquisicion_esbye",
        valor: "valor_esbye",
        ubicacion: "ubicacion_esbye",
        observacion: "observacion_esbye",
    };

    function escapeHtml(text) {
        return String(text || "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    function normalizeText(value) {
        return String(value || "")
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLowerCase()
            .trim();
    }

    function normalizeSelectComparable(value) {
        return normalizeText(value)
            .replace(/[^a-z0-9\s]/g, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    function normalizeCodeToPlaceholder(value) {
        const text = String(value || "").trim();
        const compact = text.toLowerCase().replace(/[^a-z0-9]/g, "");
        if (!compact || compact === "sc" || compact === "sincodigo" || compact === "sincod") {
            return "S/C";
        }
        return text;
    }

    function normalizeDateInputValue(value) {
        const raw = String(value || "").trim();
        if (!raw) return "";
        if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
        const isoDateTime = raw.match(/^(\d{4}-\d{2}-\d{2})[ T]\d{2}:\d{2}:\d{2}(?:\.\d+)?$/);
        if (isoDateTime) return isoDateTime[1];
        const dmyWithTime = raw.match(/^(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})\s+\d{1,2}:\d{2}(?::\d{2})?$/);
        if (dmyWithTime) {
            return normalizeDateInputValue(dmyWithTime[1]);
        }
        if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(raw)) {
            const [day, month, year] = raw.split(/[\/\-]/).map((part) => Number(part));
            return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
        }
        const parsed = new Date(raw);
        if (!Number.isNaN(parsed.getTime())) {
            return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
        }
        return "";
    }

    function resolveSelectOptionValue(options, rawValue) {
        const value = String(rawValue || "").trim();
        if (!value) return "";
        const normalizedValue = normalizeSelectComparable(value);
        const list = Array.isArray(options) ? options : [];
        if (!list.length) return "";

        let exact = list.find((name) => normalizeSelectComparable(name) === normalizedValue);
        if (exact) return exact;

        let contains = list.find(
            (name) =>
                normalizeSelectComparable(name).includes(normalizedValue)
                || normalizedValue.includes(normalizeSelectComparable(name))
        );
        if (contains) return contains;

        const targetTokens = normalizedValue.split(/\s+/).filter(Boolean);
        let best = "";
        let bestScore = 0;
        list.forEach((name) => {
            const norm = normalizeSelectComparable(name);
            if (!norm) return;
            const tokens = norm.split(/\s+/).filter(Boolean);
            const overlap = targetTokens.filter((token) => tokens.some((item) => item.includes(token) || token.includes(item))).length;
            const score = targetTokens.length ? overlap / targetTokens.length : 0;
            if (score > bestScore) {
                bestScore = score;
                best = name;
            }
        });
        return bestScore >= 0.6 ? best : "";
    }

    function normalizeEsbyeDuplicateMappings(headers, mappings) {
        const out = Array.isArray(mappings) ? [...mappings] : [];
        const seen = {};
        (out || []).forEach((field, idx) => {
            if (!field || !BASE_TO_ESBYE_FIELD[field]) return;
            const headerNorm = normalizeText(Array.isArray(headers) ? headers[idx] : "");
            const count = seen[field] || 0;
            if (headerNorm.includes("esbye") || count >= 1) {
                out[idx] = BASE_TO_ESBYE_FIELD[field];
            }
            seen[field] = count + 1;
        });
        return out;
    }

    function extractAulaCode(value) {
        const normalized = normalizeText(value);
        const match = normalized.match(/(\d+[a-z]\s*-?\s*\d+)/);
        if (!match) return "";
        return match[1].replace(/\s+/g, "");
    }

    function extractFloorHint(value) {
        const normalized = normalizeText(value);
        if (normalized.includes("planta baja")) return "planta baja";
        if (normalized.includes("primer") && normalized.includes("piso")) return "primer piso";
        if (normalized.includes("segundo") && normalized.includes("piso")) return "segundo piso";
        if (normalized.includes("tercer") && normalized.includes("piso")) return "tercer piso";
        return "";
    }

    function extractLocationKind(value) {
        const normalized = normalizeText(value);
        if (!normalized) return "";
        if (normalized.includes("pasillo")) return "pasillo";
        if (normalized.includes("aula") || extractAulaCode(normalized)) return "aula";
        return "";
    }

    function getStatusBadge(status) {
        if (status === "exact") return '<span class="badge text-bg-danger">Registrado (igual)</span>';
        if (status === "similar") return '<span class="badge text-bg-warning">Similar</span>';
        return '<span class="badge text-bg-success">Nuevo</span>';
    }

    function summarizeMatches(row) {
        const exact = row.exact_matches || [];
        const similar = row.similar_matches || [];
        const all = exact.length ? exact : similar;
        if (!all.length) return "Sin coincidencias en inventario.";
        const lines = all.slice(0, 4).map((it) => {
            const fields = Array.isArray(it.match_fields) && it.match_fields.length ? ` [${it.match_fields.join(", ")}]` : "";
            return `#${it.item_numero || "-"}${fields} | INV ${it.cod_inventario || "-"} | ESBYE ${it.cod_esbye || "-"}`;
        });
        const extra = all.length > 4 ? `<div class="small text-muted">... y ${all.length - 4} mas</div>` : "";
        return `<div class="small text-muted" style="max-height:90px; overflow:auto;">${lines.map((x) => `<div>${escapeHtml(x)}</div>`).join("")}${extra}</div>`;
    }

    function buildConflictDetailsHtml(rowData, rowPosition) {
        const exact = Array.isArray(rowData.exact_matches) ? rowData.exact_matches : [];
        const similar = Array.isArray(rowData.similar_matches) ? rowData.similar_matches : [];
        const allMatches = exact.length ? exact : similar;
        const statusLabel = exact.length ? "Coincidencia exacta" : "Coincidencia similar";

        const keyData = [
            { label: "Fila Excel", value: (rowData.row_index || 0) + 1 },
            { label: "Cod. Inventario", value: rowData.data?.cod_inventario || "-" },
            { label: "Cod. ESBYE", value: rowData.data?.cod_esbye || "-" },
            { label: "Descripcion", value: rowData.data?.descripcion || "-" },
            { label: "Marca", value: rowData.data?.marca || "-" },
            { label: "Modelo", value: rowData.data?.modelo || "-" },
            { label: "Serie", value: rowData.data?.serie || "-" },
            { label: "Ubicacion", value: rowData.data?.ubicacion || "-" },
            { label: "Usuario Final", value: rowData.data?.usuario_final || "-" },
        ];

        const fieldsHtml = keyData
            .map((item) => `<div class="col-md-4 small"><strong>${escapeHtml(item.label)}:</strong> ${escapeHtml(String(item.value))}</div>`)
            .join("");

        const repeatedRowsHtml = allMatches.length
            ? allMatches.map((item) => `
                <div class="import-conflict-match-row">
                    <div class="small"><strong>Item #${escapeHtml(item.item_numero || "-")}</strong></div>
                    <div class="small text-muted">INV: ${escapeHtml(item.cod_inventario || "-")} | ESBYE: ${escapeHtml(item.cod_esbye || "-")}</div>
                    <div class="small text-warning-emphasis">Coincide por: ${escapeHtml((item.match_fields || []).length ? item.match_fields.join(", ") : "revision manual")}</div>
                    <div class="small">${escapeHtml(item.descripcion || "-")}</div>
                    <div class="small text-muted">Ubicacion: ${escapeHtml(item.ubicacion || "-")} | Usuario: ${escapeHtml(item.usuario_final || "-")}</div>
                </div>
            `).join("")
            : '<div class="small text-muted">Sin coincidencias listables.</div>';

        return `
            <div class="accordion-item border-warning-subtle">
                <h2 class="accordion-header" id="import-conflict-head-${rowPosition}">
                    <button class="accordion-button ${rowPosition > 0 ? "collapsed" : ""}" type="button" data-bs-toggle="collapse" data-bs-target="#import-conflict-body-${rowPosition}" aria-expanded="${rowPosition === 0 ? "true" : "false"}">
                        <div class="d-flex flex-wrap align-items-center gap-2 w-100 pe-3">
                            <span class="badge text-bg-warning">${escapeHtml(statusLabel)}</span>
                            <span class="small fw-semibold">Fila ${(rowData.row_index || 0) + 1}</span>
                            <span class="small text-muted">${escapeHtml(rowData.data?.descripcion || "Sin descripcion")}</span>
                        </div>
                    </button>
                </h2>
                <div id="import-conflict-body-${rowPosition}" class="accordion-collapse collapse ${rowPosition === 0 ? "show" : ""}" data-bs-parent="#import-conflicts-accordion">
                    <div class="accordion-body pt-2">
                        <div class="row g-2 mb-2">
                            ${fieldsHtml}
                        </div>
                        <div class="import-conflict-matches-wrap">
                            ${repeatedRowsHtml}
                        </div>
                    </div>
                </div>
            </div>
        `;
    }

    window.initInventarioImportar = function (opts) {
        const onImportSuccess = (opts || {}).onImportSuccess;
        const modal = document.getElementById("modalImportarExcel");
        if (!modal) return;

        const bsModal = new bootstrap.Modal(modal);

        const panel1 = document.getElementById("import-panel-1");
        const panel2 = document.getElementById("import-panel-2");
        const panel3 = document.getElementById("import-panel-3");

        const dropzone = document.getElementById("import-dropzone");
        const fileInput = document.getElementById("import-file-input");
        const btnSelectFile = document.getElementById("btn-import-select-file");
        const uploadStatus = document.getElementById("import-upload-status");

        const mappingSelectsRow = document.getElementById("import-mapping-selects-row");
        const mappingHeadersRow = document.getElementById("import-mapping-headers-row");
        const mappingPreviewBody = document.getElementById("import-mapping-preview-body");
        const previewInfo = document.getElementById("import-preview-info");

        const rowCountEl = document.getElementById("import-row-count");
        const chunkRangeEl = document.getElementById("import-chunk-range");
        const areaSelect = document.getElementById("import-area-select");
        const reviewSummary = document.getElementById("import-review-summary");
        const reviewBody = document.getElementById("import-review-table-body");
        const resultArea = document.getElementById("import-result-area");

        const btnBack = document.getElementById("btn-import-back");
        const btnNext = document.getElementById("btn-import-next");
        const btnValidate = document.getElementById("btn-import-validate");
        const btnDeleteRow = document.getElementById("btn-import-row-delete");
        const btnConfirm = document.getElementById("btn-import-confirm");
        const btnSkip = document.getElementById("btn-import-skip");

        const importConflictsModalEl = document.getElementById("modalImportConflicts");
        const importConflictsSummary = document.getElementById("import-conflicts-summary");
        const importConflictsAccordion = document.getElementById("import-conflicts-accordion");
        const importConflictsCancelBtn = document.getElementById("btn-import-conflicts-cancel");
        const importConflictsContinueBtn = document.getElementById("btn-import-conflicts-continue");
        const importConflictsModal = importConflictsModalEl ? new bootstrap.Modal(importConflictsModalEl) : null;

        const importSaveConfirmModalEl = document.getElementById("modalImportSaveConfirm");
        const importSaveConfirmText = document.getElementById("import-save-confirm-text");
        const importSaveConfirmCancelBtn = document.getElementById("btn-import-save-confirm-cancel");
        const importSaveConfirmContinueBtn = document.getElementById("btn-import-save-confirm-continue");
        const importSaveConfirmModal = importSaveConfirmModalEl ? new bootstrap.Modal(importSaveConfirmModalEl) : null;

        const importInvalidLocationsModalEl = document.getElementById("modalImportInvalidLocations");
        const importInvalidLocationsSummary = document.getElementById("import-invalid-locations-summary");
        const importInvalidLocationsList = document.getElementById("import-invalid-locations-list");
        const importInvalidOpenSettingsBtn = document.getElementById("btn-import-invalid-open-settings");
        const importInvalidLocationsModal = importInvalidLocationsModalEl
            ? new bootstrap.Modal(importInvalidLocationsModalEl)
            : null;

        const stepBadges = [1, 2, 3].map((n) => document.getElementById(`import-step-badge-${n}`));
        const stepLabels = [1, 2, 3].map((n) => document.getElementById(`import-step-label-${n}`));

        function ensureNestedModalStack(modalEl, modalZ = 1090) {
            if (!modalEl) return;
            modalEl.addEventListener("shown.bs.modal", () => {
                modalEl.style.zIndex = String(modalZ);
                const backdrops = Array.from(document.querySelectorAll(".modal-backdrop.show"));
                const activeBackdrop = backdrops[backdrops.length - 1];
                if (activeBackdrop) {
                    activeBackdrop.classList.add("import-nested-backdrop");
                    activeBackdrop.style.zIndex = String(modalZ - 1);
                }
            });
        }

        ensureNestedModalStack(importSaveConfirmModalEl, 1090);
        ensureNestedModalStack(importConflictsModalEl, 1090);
        ensureNestedModalStack(importInvalidLocationsModalEl, 1090);

        let state = buildInitialState();

        function buildInitialState() {
            return {
                step: 1,
                sessionId: null,
                sourceFile: null,
                headers: [],
                previewRows: [],
                totalRows: 0,
                mappings: [],
                startIndex: 0,
                chunkSize: CHUNK_SIZE,
                reviewRows: [],
                selectedReviewRowIndex: null,
                areaOptions: [],
                cuentaOptions: [],
                usuarioOptions: [],
                estadoOptions: [],
                removedRowSignaturesByBlock: {},
                invalidLocationsModalShownByBlock: {},
                shouldRefreshAreasOnFocus: false,
                isRefreshingAreas: false,
                areaRefreshPollerId: null,
                isSaving: false,
                isRecoveringSession: false,
            };
        }

        function stopPendingAreaRefreshPolling() {
            if (state.areaRefreshPollerId) {
                clearInterval(state.areaRefreshPollerId);
                state.areaRefreshPollerId = null;
            }
        }

        function startPendingAreaRefreshPolling() {
            stopPendingAreaRefreshPolling();
            state.areaRefreshPollerId = setInterval(async () => {
                if (!state.shouldRefreshAreasOnFocus || state.step !== 3) {
                    stopPendingAreaRefreshPolling();
                    return;
                }
                await refreshAreasAfterSettingsChange();
            }, 2000);
        }

        function getCurrentBlockKey() {
            return `${state.startIndex}:${state.chunkSize}`;
        }

        function buildRowSignature(data) {
            const source = data || {};
            const fields = [
                "cod_inventario",
                "cod_esbye",
                "cuenta",
                "cantidad",
                "descripcion",
                "marca",
                "modelo",
                "serie",
                "estado",
                "ubicacion",
                "fecha_adquisicion",
                "valor",
                "usuario_final",
                "observacion",
            ];
            return fields
                .map((field) => normalizeText(source[field] == null ? "" : String(source[field])))
                .join("|");
        }

        function getRemovedRowSignaturesForCurrentBlock() {
            const key = getCurrentBlockKey();
            if (!state.removedRowSignaturesByBlock[key]) {
                state.removedRowSignaturesByBlock[key] = new Set();
            }
            return state.removedRowSignaturesByBlock[key];
        }

        function applyRemovedRowsFilter(rows) {
            const removedSignatures = getRemovedRowSignaturesForCurrentBlock();
            if (!removedSignatures.size) return rows;
            return (rows || []).filter((row) => !removedSignatures.has(buildRowSignature(row?.data || {})));
        }

        function buildInvalidLocationsFromReviewRows() {
            return (state.reviewRows || [])
                .filter((row) => row && row.invalidLocation)
                .map((row) => ({
                    row_index: row.row_index,
                    ubicacion: row?.data?.ubicacion,
                    descripcion: row?.data?.descripcion,
                    cod_inventario: row?.data?.cod_inventario,
                    cod_esbye: row?.data?.cod_esbye,
                }));
        }

        function maybeOpenInvalidLocationsModalAfterRender() {
            const invalid = buildInvalidLocationsFromReviewRows();
            if (!invalid.length) {
                return;
            }
            const blockKey = getCurrentBlockKey();
            if (state.invalidLocationsModalShownByBlock[blockKey]) {
                return;
            }

            state.invalidLocationsModalShownByBlock[blockKey] = true;

            // Espera al siguiente frame para asegurar que la tabla ya quedó pintada.
            requestAnimationFrame(() => {
                requestAnimationFrame(() => {
                    openInvalidLocationsModal({ invalid_locations: invalid });
                });
            });
        }

        async function loadReviewSelectOptions() {
            try {
                const [cuentasRes, adminsRes, estadosRes] = await Promise.all([
                    fetch("/api/parametros/cuentas"),
                    fetch("/api/administradores"),
                    fetch("/api/parametros/estados"),
                ]);
                const cuentasJson = await cuentasRes.json();
                const adminsJson = await adminsRes.json();
                const estadosJson = await estadosRes.json();
                state.cuentaOptions = Array.isArray(cuentasJson?.data)
                    ? cuentasJson.data.map((it) => String(it?.nombre || "").trim()).filter(Boolean)
                    : [];
                state.usuarioOptions = Array.isArray(adminsJson?.data)
                    ? adminsJson.data.map((it) => String(it?.nombre || "").trim()).filter(Boolean)
                    : [];
                state.estadoOptions = Array.isArray(estadosJson?.data)
                    ? estadosJson.data.map((it) => String(it?.nombre || "").trim()).filter(Boolean)
                    : [];
            } catch (_err) {
                state.cuentaOptions = [];
                state.usuarioOptions = [];
                state.estadoOptions = [];
            }
        }

        function setStep(step) {
            state.step = step;
            [panel1, panel2, panel3].forEach((p, i) => p.classList.toggle("d-none", i + 1 !== step));
            stepBadges.forEach((badge, i) => {
                const n = i + 1;
                const done = n < step;
                const active = n === step;
                badge.className = `badge rounded-pill ${done ? "bg-success" : active ? "bg-primary" : "bg-secondary"}`;
                badge.innerHTML = done ? '<i class="bi bi-check"></i>' : String(n);
            });
            stepLabels.forEach((label, i) => {
                label.className = `small ${i + 1 === step ? "fw-semibold" : "text-muted"}`;
            });
            btnBack.classList.toggle("d-none", step === 1);
            btnNext.classList.toggle("d-none", step !== 2);
            btnConfirm.classList.toggle("d-none", step !== 3);
            if (btnSkip) {
                btnSkip.classList.toggle("d-none", step !== 3);
            }
        }

        function showStatus(html, type) {
            uploadStatus.className = `alert alert-${type || "info"} small mt-3 py-2`;
            uploadStatus.innerHTML = html;
            uploadStatus.classList.remove("d-none");
        }

        function clearStatus() {
            uploadStatus.classList.add("d-none");
            uploadStatus.innerHTML = "";
        }

        function toMappingObject() {
            const mapping = {};
            state.mappings.forEach((field, idx) => {
                mapping[String(idx)] = field || "";
            });
            return mapping;
        }

        async function handleFile(file, options = {}) {
            const reuseMappings = Boolean(options.reuseMappings);
            const silent = Boolean(options.silent);
            const keepCurrentStep = Boolean(options.keepCurrentStep);

            if (!file?.name || !file.name.toLowerCase().endsWith(".xlsx")) {
                if (!silent) showStatus('<i class="bi bi-x-circle me-1"></i>Solo se aceptan archivos .xlsx.', "danger");
                return false;
            }
            if (file.size > 10 * 1024 * 1024) {
                if (!silent) showStatus('<i class="bi bi-x-circle me-1"></i>El archivo supera 10 MB.', "danger");
                return false;
            }

            state.sourceFile = file;
            if (!silent) {
                showStatus(`<span class="spinner-border spinner-border-sm me-2" role="status"></span>Procesando <strong>${escapeHtml(file.name)}</strong>...`);
            }
            const formData = new FormData();
            formData.append("file", file);

            try {
                const res = await fetch("/api/inventario/previsualizar-excel", { method: "POST", body: formData });
                const data = await res.json();
                if (!res.ok || data.error) {
                    if (!silent) {
                        showStatus(`<i class="bi bi-x-circle me-1"></i>${escapeHtml(data.error || "Error al procesar archivo")}`, "danger");
                    }
                    return false;
                }

                const previousMappings = Array.isArray(state.mappings) ? [...state.mappings] : [];
                state.sessionId = data.session_id;
                state.headers = data.headers || [];
                state.previewRows = data.preview_rows || [];
                state.totalRows = Number(data.total_rows || 0);
                const suggested = Array.isArray(data.suggested_mapping) ? data.suggested_mapping : [];
                const rawMappings = state.headers.map((_, idx) => suggested[idx] || "");
                if (reuseMappings && previousMappings.length === state.headers.length && previousMappings.some(Boolean)) {
                    state.mappings = normalizeEsbyeDuplicateMappings(state.headers, previousMappings);
                } else {
                    state.mappings = normalizeEsbyeDuplicateMappings(state.headers, rawMappings);
                }

                if (!silent) {
                    clearStatus();
                }
                renderMappingStep();
                if (!keepCurrentStep) {
                    setStep(2);
                }
                return true;
            } catch (_error) {
                if (!silent) {
                    showStatus('<i class="bi bi-x-circle me-1"></i>Error de red al subir archivo.', "danger");
                }
                return false;
            }
        }

        async function recoverExpiredImportSession() {
            if (state.isRecoveringSession) return false;
            if (!state.sourceFile) {
                showStatus('<i class="bi bi-x-circle me-1"></i>La sesión expiró y no hay archivo en memoria para reconectar. Selecciona el Excel nuevamente.', "danger");
                return false;
            }

            state.isRecoveringSession = true;
            const keepStep = state.step;
            const keepStartIndex = state.startIndex;
            try {
                showStatus('<span class="spinner-border spinner-border-sm me-2" role="status"></span>Reconectando sesión de importación...', "warning");
                const restored = await handleFile(state.sourceFile, {
                    reuseMappings: true,
                    silent: true,
                    keepCurrentStep: true,
                });
                if (!restored) return false;

                state.startIndex = keepStartIndex;
                setStep(keepStep);
                showStatus('<i class="bi bi-check-circle-fill me-1"></i>Sesión restablecida. Reintentando...', "success");
                return true;
            } finally {
                state.isRecoveringSession = false;
            }
        }

        function buildColumnSelect(colIdx) {
            const sel = document.createElement("select");
            sel.className = "form-select form-select-sm";
            sel.style.fontSize = "0.78rem";
            CANONICAL_FIELDS.forEach(({ value, label }) => {
                const opt = document.createElement("option");
                opt.value = value;
                opt.textContent = label;
                if (value === state.mappings[colIdx]) opt.selected = true;
                sel.appendChild(opt);
            });
            sel.addEventListener("change", () => {
                state.mappings[colIdx] = sel.value;
                state.mappings = normalizeEsbyeDuplicateMappings(state.headers, state.mappings);
                renderMappingStep();
                applyColumnHighlights();
            });
            return sel;
        }

        function applyColumnHighlights() {
            const allSelThs = mappingSelectsRow.querySelectorAll("th");
            const allHdrThs = mappingHeadersRow.querySelectorAll("th");
            const allDataTrs = mappingPreviewBody.querySelectorAll("tr");

            state.mappings.forEach((mapping, colIdx) => {
                const ignored = !mapping;
                if (allSelThs[colIdx]) allSelThs[colIdx].classList.toggle("table-secondary", ignored);
                if (allHdrThs[colIdx]) allHdrThs[colIdx].classList.toggle("table-secondary", ignored);
                allDataTrs.forEach((tr) => {
                    if (tr.children[colIdx]) tr.children[colIdx].classList.toggle("table-secondary", ignored);
                });
            });
        }

        function renderMappingStep() {
            mappingSelectsRow.innerHTML = "";
            mappingHeadersRow.innerHTML = "";
            mappingPreviewBody.innerHTML = "";

            state.headers.forEach((header, colIdx) => {
                const thSel = document.createElement("th");
                thSel.style.minWidth = "165px";
                thSel.style.padding = "4px 6px";
                thSel.appendChild(buildColumnSelect(colIdx));
                mappingSelectsRow.appendChild(thSel);

                const thHdr = document.createElement("th");
                thHdr.textContent = header || `(col ${colIdx + 1})`;
                thHdr.className = "text-muted small fw-normal";
                thHdr.style.padding = "2px 6px";
                mappingHeadersRow.appendChild(thHdr);
            });

            state.previewRows.forEach((row) => {
                const tr = document.createElement("tr");
                state.headers.forEach((_, colIdx) => {
                    const td = document.createElement("td");
                    td.className = "small";
                    td.style.maxWidth = "200px";
                    td.style.overflow = "hidden";
                    td.style.textOverflow = "ellipsis";
                    td.style.whiteSpace = "nowrap";
                    const val = String(row[colIdx] ?? "");
                    td.textContent = val;
                    td.title = val;
                    tr.appendChild(td);
                });
                mappingPreviewBody.appendChild(tr);
            });

            previewInfo.textContent = `Vista previa: ${state.previewRows.length} de ${state.totalRows} filas.`;
            applyColumnHighlights();
        }

        function flattenAreas(structure) {
            const result = [];
            (structure || []).forEach((block) => {
                (block.pisos || []).forEach((floor) => {
                    (floor.areas || []).forEach((area) => {
                        result.push({
                            id: Number(area.id),
                            block: String(block.nombre || ""),
                            floor: String(floor.nombre || ""),
                            area: String(area.nombre || ""),
                            label: `${block.nombre} / ${floor.nombre} / ${area.nombre}`,
                        });
                    });
                });
            });
            return result;
        }

        function suggestAreaId(ubicacionText) {
            const target = normalizeText(ubicacionText);
            if (!target || !state.areaOptions.length) return null;

            const targetTokens = target.split(/\s+|\//).filter(Boolean);
            const targetAulaCode = extractAulaCode(target);
            const targetFloorHint = extractFloorHint(target);
            const targetKind = extractLocationKind(target);
            const blockLetterMatch = (targetAulaCode || "").match(/\d+([a-z])/);
            const targetBlockLetter = blockLetterMatch ? blockLetterMatch[1] : "";

            let best = null;
            let bestScore = 0;
            let bestPasillo = null;
            let bestPasilloScore = -Infinity;
            state.areaOptions.forEach((opt) => {
                const block = normalizeText(opt.block);
                const floor = normalizeText(opt.floor);
                const area = normalizeText(opt.area);
                const label = normalizeText(opt.label);
                const areaTokens = area.split(/\s+|\//).filter(Boolean);
                const areaAulaCode = extractAulaCode(area);
                let score = 0;
                if (target === area || target === `${block} / ${floor} / ${area}`) score += 10;
                if (target.includes(area)) score += 6;
                if (area.includes(target)) score += 4;
                if (block && target.includes(block)) score += 2;
                if (floor && target.includes(floor)) score += 2;
                if (label && target.includes(label)) score += 6;

                if (targetAulaCode && areaAulaCode && targetAulaCode === areaAulaCode) {
                    score += 18;
                }
                if (targetBlockLetter && block.includes(`bloque ${targetBlockLetter}`)) {
                    score += 8;
                }

                if (targetFloorHint && floor.includes(targetFloorHint)) {
                    score += 6;
                }
                if (target.includes("alto")) {
                    if (floor.includes("alto")) score += 3;
                    if (floor.includes("bajo")) score -= 3;
                }
                if (target.includes("bajo")) {
                    if (floor.includes("bajo")) score += 3;
                    if (floor.includes("alto")) score -= 3;
                }

                if (targetKind === "pasillo") {
                    if (area.includes("pasillo")) {
                        score += 14;
                    } else {
                        score -= 16;
                    }
                } else if (targetKind === "aula") {
                    if (area.includes("aula") || areaAulaCode) {
                        score += 6;
                    }
                }

                const tokenHits = targetTokens.filter((tk) => tk.length >= 3 && areaTokens.some((ak) => ak.includes(tk) || tk.includes(ak))).length;
                if (tokenHits > 0) {
                    score += tokenHits * 2;
                }

                if (targetKind === "pasillo" && area.includes("pasillo")) {
                    let pasilloScore = score;
                    if (targetFloorHint && floor.includes(targetFloorHint)) {
                        pasilloScore += 4;
                    }
                    if (pasilloScore > bestPasilloScore) {
                        bestPasilloScore = pasilloScore;
                        bestPasillo = opt;
                    }
                }

                if (score > bestScore) {
                    bestScore = score;
                    best = opt;
                }
            });

            if (targetKind === "pasillo") {
                return bestPasillo || null;
            }
            return bestScore >= 3 ? best : null;
        }

        function applyAreaSuggestionToRows() {
            state.reviewRows.forEach((row) => {
                if (row.data.area_id) {
                    if (!row.areaLabel) {
                        const existing = state.areaOptions.find((opt) => String(opt.id) === String(row.data.area_id));
                        if (existing) row.areaLabel = existing.label;
                    }
                    return;
                }
                const suggested = suggestAreaId(row.data.ubicacion);
                if (suggested) {
                    row.data.area_id = suggested.id;
                    row.areaLabel = suggested.label;
                }
            });
        }

        function applyCatalogSuggestionToRows() {
            state.reviewRows.forEach((row) => {
                if (!row || !row.data) return;

                // Estandariza fechas para que en revision se muestren solo como fecha.
                row.data.fecha_adquisicion = normalizeDateInputValue(row.data.fecha_adquisicion) || (row.data.fecha_adquisicion || "");
                row.data.fecha_adquisicion_esbye = normalizeDateInputValue(row.data.fecha_adquisicion_esbye) || (row.data.fecha_adquisicion_esbye || "");

                const cuentaRaw = String(row.data.cuenta || "").trim();
                if (cuentaRaw) {
                    const cuentaResolved = resolveSelectOptionValue(state.cuentaOptions, cuentaRaw);
                    if (cuentaResolved) row.data.cuenta = cuentaResolved;
                }

                const usuarioRaw = String(row.data.usuario_final || "").trim();
                if (usuarioRaw) {
                    const usuarioResolved = resolveSelectOptionValue(state.usuarioOptions, usuarioRaw);
                    if (usuarioResolved) row.data.usuario_final = usuarioResolved;
                }

                const estadoRaw = String(row.data.estado || "").trim();
                if (estadoRaw) {
                    const estadoResolved = resolveSelectOptionValue(state.estadoOptions, estadoRaw);
                    if (estadoResolved) row.data.estado = estadoResolved;
                }
            });
        }

        function setReviewLoadingState(isLoading, message) {
            btnValidate.disabled = Boolean(isLoading);
            btnDeleteRow.disabled = Boolean(isLoading);
            btnConfirm.disabled = Boolean(isLoading) || state.isSaving;
            if (btnSkip) btnSkip.disabled = Boolean(isLoading) || state.isSaving;

            if (isLoading) {
                btnConfirm.classList.add("d-none");
                if (btnSkip) btnSkip.classList.add("d-none");
                resultArea.className = "alert alert-info mt-3 small";
                resultArea.innerHTML = `<span class="spinner-border spinner-border-sm me-2" role="status"></span>${escapeHtml(message || "Cargando bloque de importacion...")}`;
                resultArea.classList.remove("d-none");
            } else {
                btnConfirm.classList.remove("d-none");
                if (btnSkip) btnSkip.classList.remove("d-none");
            }
        }

        function evaluateRowLocationState(row) {
            if (!row || !row.data) {
                return;
            }

            row.invalidLocation = false;
            row.invalidLocationReason = "";

            const locationText = String(row.data.ubicacion || "").trim();
            if (!locationText) {
                return;
            }

            const areaIdValue = row.data.area_id;
            if (areaIdValue != null && String(areaIdValue).trim() !== "") {
                const existsById = state.areaOptions.some((opt) => String(opt.id) === String(areaIdValue));
                if (!existsById) {
                    row.invalidLocation = true;
                    row.invalidLocationReason = "El area asociada no existe.";
                }
                return;
            }

            const suggested = suggestAreaId(locationText);
            if (!suggested) {
                row.invalidLocation = true;
                row.invalidLocationReason = "Ubicacion no encontrada en configuracion.";
            }
        }

        function refreshInvalidLocationFlags() {
            state.reviewRows.forEach((row) => evaluateRowLocationState(row));
        }

        function renderReviewSummary(summary) {
            if (!summary) {
                reviewSummary.innerHTML = "";
                return;
            }
            reviewSummary.innerHTML = [
                `<span class="me-2"><strong>Nuevos:</strong> ${summary.normal || 0}</span>`,
                `<span class="me-2 text-danger"><strong>Iguales:</strong> ${summary.exact || 0}</span>`,
                `<span class="text-warning-emphasis"><strong>Similares:</strong> ${summary.similar || 0}</span>`,
            ].join("");
        }

        function renderSummaryFromRows() {
            renderReviewSummary({
                exact: state.reviewRows.filter((r) => r.status === "exact").length,
                similar: state.reviewRows.filter((r) => r.status === "similar").length,
                normal: state.reviewRows.filter((r) => r.status === "normal").length,
            });
        }

        function renderReviewRows() {
            const isNoCodeValue = (value) => {
                const compact = String(value || "")
                    .trim()
                    .toLowerCase()
                    .replace(/[^a-z0-9]/g, "");
                return compact === "sc" || compact === "sincodigo" || compact === "sincod";
            };

            reviewBody.innerHTML = "";
            state.reviewRows.forEach((row, localIndex) => {
                const tr = document.createElement("tr");
                tr.dataset.localIndex = String(localIndex);
                tr.classList.add(`import-row-${row.status || "normal"}`);
                if (row.invalidLocation) {
                    tr.classList.add("import-row-invalid-location");
                }
                if (state.selectedReviewRowIndex === localIndex) {
                    tr.classList.add("import-row-selected");
                }

                const tdIndex = document.createElement("td");
                tdIndex.className = "small";
                tdIndex.textContent = String((row.row_index || 0) + 1);

                const tdStatus = document.createElement("td");
                tdStatus.innerHTML = `${getStatusBadge(row.status)} ${row.areaLabel ? `<div class="small text-muted mt-1">Area: ${escapeHtml(row.areaLabel)}</div>` : ""}${row.invalidLocation ? `<div class="small text-danger fw-semibold mt-1"><i class="bi bi-exclamation-triangle-fill me-1"></i>Ubicacion invalida</div>` : ""}`;

                tr.appendChild(tdIndex);
                tr.appendChild(tdStatus);

                REVIEW_TABLE_FIELDS.forEach((field) => {
                    const td = document.createElement("td");
                    td.className = "small import-review-cell";
                    td.dataset.field = field;
                    td.title = "Doble clic para editar";
                    const value = row.data[field] == null ? "" : String(row.data[field]);
                    td.textContent = value || "-";
                    if ((field === "cod_inventario" || field === "cod_esbye") && isNoCodeValue(value)) {
                        td.classList.add("code-sc-cell");
                    }
                    if (field === "ubicacion" && row.invalidLocation) {
                        td.classList.add("import-invalid-location-cell");
                        td.title = row.invalidLocationReason || "Ubicacion no encontrada. Revisa este dato.";
                        td.innerHTML = `<i class="bi bi-exclamation-circle-fill me-1 text-danger"></i>${escapeHtml(value || "-")}`;
                    }
                    tr.appendChild(td);
                });

                const tdMatches = document.createElement("td");
                tdMatches.innerHTML = summarizeMatches(row);
                tr.appendChild(tdMatches);

                reviewBody.appendChild(tr);
            });
            btnDeleteRow.classList.toggle("d-none", state.selectedReviewRowIndex == null);
        }

        function updateSelectedRowVisual() {
            reviewBody.querySelectorAll("tr[data-local-index]").forEach((tr) => {
                const idx = Number(tr.dataset.localIndex);
                tr.classList.toggle("import-row-selected", idx === state.selectedReviewRowIndex);
            });
            btnDeleteRow.classList.toggle("d-none", state.selectedReviewRowIndex == null);
        }

        async function loadAreasForSelect() {
            const previousAreaValue = String(areaSelect.value || "");
            areaSelect.innerHTML = '<option value="">Sin area especifica</option>';
            try {
                const res = await fetch(`/api/estructura?include_details=0&_ts=${Date.now()}`, {
                    cache: "no-store",
                });
                const json = await res.json();
                state.areaOptions = flattenAreas(json.data || []);
                state.areaOptions.forEach((opt) => {
                    const option = document.createElement("option");
                    option.value = String(opt.id);
                    option.textContent = opt.label;
                    areaSelect.appendChild(option);
                });
                if (previousAreaValue) {
                    const stillExists = state.areaOptions.some((opt) => String(opt.id) === previousAreaValue);
                    if (stillExists) {
                        areaSelect.value = previousAreaValue;
                    }
                }
            } catch (_err) {
                state.areaOptions = [];
            }
        }

        async function refreshAreasAfterSettingsChange() {
            if (state.step !== 3 || state.isRefreshingAreas) {
                return;
            }
            state.isRefreshingAreas = true;
            try {
                await loadAreasForSelect();
                applyAreaSuggestionToRows();
                refreshInvalidLocationFlags();
                renderReviewRows();
                renderSummaryFromRows();

                const invalid = buildInvalidLocationsFromReviewRows();
                if (!invalid.length && importInvalidLocationsModal) {
                    importInvalidLocationsModal.hide();
                }

                if (invalid.length) {
                    resultArea.className = "alert alert-warning mt-3 small";
                    resultArea.innerHTML = `<i class="bi bi-exclamation-triangle-fill me-1"></i>Ubicaciones actualizadas desde Configuración, pero aún hay <strong>${invalid.length}</strong> fila(s) inválida(s).`;
                } else {
                    resultArea.className = "alert alert-success mt-3 small";
                    resultArea.innerHTML = '<i class="bi bi-check-circle-fill me-1"></i>Ubicaciones actualizadas desde Configuración.';
                }
                resultArea.classList.remove("d-none");

                // Mantiene activo el refresco automático hasta que todas las ubicaciones queden válidas.
                state.shouldRefreshAreasOnFocus = invalid.length > 0;
                if (!state.shouldRefreshAreasOnFocus) {
                    stopPendingAreaRefreshPolling();
                }
            } finally {
                state.isRefreshingAreas = false;
            }
        }

        async function validateCurrentChunk() {
            resultArea.classList.add("d-none");
            btnValidate.disabled = true;
            btnValidate.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Verificando...';

            try {
                const res = await fetch("/api/inventario/confirmar-importacion", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                        session_id: state.sessionId,
                        mapping: toMappingObject(),
                        start_index: state.startIndex,
                        chunk_size: state.chunkSize,
                        validate_only: true,
                        area_id: areaSelect.value ? Number(areaSelect.value) : null,
                    }),
                });
                const data = await res.json();
                if (res.status === 410) {
                    const recovered = await recoverExpiredImportSession();
                    if (recovered) {
                        await validateCurrentChunk();
                        return;
                    }
                }
                if (!res.ok || !data.success) {
                    resultArea.className = "alert alert-danger mt-3 small";
                    resultArea.innerHTML = `<i class="bi bi-x-circle me-1"></i>${escapeHtml(data.error || "No se pudo validar el bloque")}`;
                    resultArea.classList.remove("d-none");
                    return false;
                }

                state.totalRows = Number(data.total_rows || state.totalRows || 0);
                const incomingRows = Array.isArray(data.rows) ? data.rows : [];
                state.reviewRows = applyRemovedRowsFilter(incomingRows);
                state.selectedReviewRowIndex = null;
                applyAreaSuggestionToRows();
                applyCatalogSuggestionToRows();
                refreshInvalidLocationFlags();
                renderSummaryFromRows();
                renderReviewRows();
                maybeOpenInvalidLocationsModalAfterRender();

                const from = state.startIndex + 1;
                const to = Math.min(state.startIndex + state.chunkSize, state.totalRows);
                chunkRangeEl.textContent = `Bloque actual: filas ${from} a ${to} de ${state.totalRows}`;
                return true;
            } catch (_error) {
                resultArea.className = "alert alert-danger mt-3 small";
                resultArea.innerHTML = '<i class="bi bi-x-circle me-1"></i>Error de red al verificar bloque.';
                resultArea.classList.remove("d-none");
                return false;
            } finally {
                btnValidate.disabled = false;
                btnValidate.innerHTML = '<i class="bi bi-search me-1"></i>Verificar bloque actual';
            }
        }

        function buildCommitPayload(forceDuplicate) {
            const rows = state.reviewRows.map((row) => row.data || {});
            return {
                session_id: state.sessionId,
                mapping: toMappingObject(),
                area_id: areaSelect.value ? Number(areaSelect.value) : null,
                start_index: state.startIndex,
                chunk_size: state.chunkSize,
                rows,
                force_duplicate: Boolean(forceDuplicate),
            };
        }

        function openSaveConfirmModal() {
            if (!importSaveConfirmModal || !importSaveConfirmCancelBtn || !importSaveConfirmContinueBtn) {
                return Promise.resolve(window.confirm("Revisa el bloque y confirma si deseas continuar con el guardado."));
            }

            const from = state.startIndex + 1;
            const to = Math.min(state.startIndex + state.chunkSize, state.totalRows || (state.startIndex + state.chunkSize));
            if (importSaveConfirmText) {
                importSaveConfirmText.textContent = `Revisa si todas las filas del bloque ${from} al ${to} se importaron correctamente y que no haya espacios o datos vacios antes de guardar.`;
            }

            return new Promise((resolve) => {
                let settled = false;

                const finish = (value) => {
                    if (settled) return;
                    settled = true;
                    cleanup();
                    resolve(value);
                };

                const onContinue = () => {
                    finish(true);
                    importSaveConfirmModal.hide();
                };
                const onCancel = () => {
                    finish(false);
                    importSaveConfirmModal.hide();
                };
                const onHidden = () => finish(false);

                const cleanup = () => {
                    importSaveConfirmContinueBtn.removeEventListener("click", onContinue);
                    importSaveConfirmCancelBtn.removeEventListener("click", onCancel);
                    importSaveConfirmModalEl.removeEventListener("hidden.bs.modal", onHidden);
                };

                importSaveConfirmContinueBtn.addEventListener("click", onContinue);
                importSaveConfirmCancelBtn.addEventListener("click", onCancel);
                importSaveConfirmModalEl.addEventListener("hidden.bs.modal", onHidden);
                importSaveConfirmModal.show();
            });
        }

        function openImportConflictsModal(conflictData) {
            if (!importConflictsModal || !importConflictsSummary || !importConflictsAccordion || !importConflictsCancelBtn || !importConflictsContinueBtn) {
                const summary = conflictData?.summary || {};
                return Promise.resolve(window.confirm(`Conflictos detectados. Iguales: ${summary.exact || 0}, similares: ${summary.similar || 0}. Deseas guardar de todas formas?`));
            }

            const summary = conflictData?.summary || {};
            const rows = Array.isArray(conflictData?.rows) ? conflictData.rows : [];

            importConflictsSummary.textContent = `Se detectaron ${summary.exact || 0} filas iguales y ${summary.similar || 0} similares en este bloque. Revisa cada caso y decide si deseas continuar.`;

            const conflictRows = rows.filter((row) => row.status === "exact" || row.status === "similar");
            importConflictsAccordion.innerHTML = conflictRows.length
                ? conflictRows.map((row, idx) => buildConflictDetailsHtml(row, idx)).join("")
                : '<div class="alert alert-secondary mb-0">No se encontraron filas en conflicto para mostrar.</div>';

            return new Promise((resolve) => {
                let settled = false;

                const finish = (value) => {
                    if (settled) return;
                    settled = true;
                    cleanup();
                    resolve(value);
                };

                const onContinue = () => {
                    finish(true);
                    importConflictsModal.hide();
                };
                const onCancel = () => {
                    finish(false);
                    importConflictsModal.hide();
                };
                const onHidden = () => finish(false);

                const cleanup = () => {
                    importConflictsContinueBtn.removeEventListener("click", onContinue);
                    importConflictsCancelBtn.removeEventListener("click", onCancel);
                    importConflictsModalEl.removeEventListener("hidden.bs.modal", onHidden);
                };

                importConflictsContinueBtn.addEventListener("click", onContinue);
                importConflictsCancelBtn.addEventListener("click", onCancel);
                importConflictsModalEl.addEventListener("hidden.bs.modal", onHidden);
                importConflictsModal.show();
            });
        }

        async function saveCurrentChunk(forceDuplicate) {
            const payload = buildCommitPayload(forceDuplicate);
            const res = await fetch("/api/inventario/confirmar-importacion", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload),
            });
            const data = await res.json();
            return { res, data };
        }

        function openInvalidLocationsModal(payload) {
            if (!importInvalidLocationsModal || !importInvalidLocationsSummary || !importInvalidLocationsList) {
                return;
            }
            const invalid = Array.isArray(payload?.invalid_locations) ? payload.invalid_locations : [];
            importInvalidLocationsSummary.textContent = `Se encontraron ${invalid.length} filas con ubicaciones no existentes. Verifica o agrega esas ubicaciones en Configuración.`;
            importInvalidLocationsList.innerHTML = invalid.length
                ? invalid.map((item) => {
                    const excelRow = Number(item?.row_index || 0) + 1;
                    const ubicacion = escapeHtml(String(item?.ubicacion || "(sin ubicacion)"));
                    const reason = String(item?.reason || "").trim();
                    return `
                        <div class="border rounded p-2 small">
                            <div><strong>Fila ${excelRow}</strong> · <span class="text-danger-emphasis">${ubicacion}</span></div>
                            ${reason ? `<div class="text-muted mt-1">${escapeHtml(reason)}</div>` : ""}
                        </div>
                    `;
                }).join("")
                : '<div class="alert alert-secondary mb-0 small">No se encontraron detalles para mostrar.</div>';
            importInvalidLocationsModal.show();
        }

        async function handleCommitCurrentChunk() {
            if (state.isSaving) return;
            if (!state.reviewRows.length) {
                await validateCurrentChunk();
                if (!state.reviewRows.length) return;
            }

            state.isSaving = true;
            btnConfirm.disabled = true;
            btnConfirm.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Guardando...';

            try {
                let { res, data } = await saveCurrentChunk(false);
                if (res.status === 410) {
                    const recovered = await recoverExpiredImportSession();
                    if (!recovered) {
                        resultArea.className = "alert alert-danger mt-3 small";
                        resultArea.innerHTML = '<i class="bi bi-x-circle me-1"></i>La sesión de importación expiró y no se pudo restablecer automáticamente.';
                        resultArea.classList.remove("d-none");
                        return;
                    }
                    ({ res, data } = await saveCurrentChunk(false));
                }
                if (res.status === 409) {
                    const summary = data.summary || {};
                    const goOn = await openImportConflictsModal(data);
                    if (!goOn) {
                        const incomingRows = Array.isArray(data.rows) ? data.rows : state.reviewRows;
                        state.reviewRows = applyRemovedRowsFilter(incomingRows);
                        applyAreaSuggestionToRows();
                        refreshInvalidLocationFlags();
                        renderReviewSummary(summary);
                        renderSummaryFromRows();
                        renderReviewRows();
                        return;
                    }
                    ({ res, data } = await saveCurrentChunk(true));
                    if (res.status === 410) {
                        const recovered = await recoverExpiredImportSession();
                        if (!recovered) {
                            resultArea.className = "alert alert-danger mt-3 small";
                            resultArea.innerHTML = '<i class="bi bi-x-circle me-1"></i>La sesión de importación expiró y no se pudo restablecer automáticamente.';
                            resultArea.classList.remove("d-none");
                            return;
                        }
                        ({ res, data } = await saveCurrentChunk(true));
                    }
                }

                if (res.status === 422 && data?.error_code === "invalid_locations") {
                    openInvalidLocationsModal(data);
                    return;
                }

                if (!res.ok || !data.success) {
                    resultArea.className = "alert alert-danger mt-3 small";
                    resultArea.innerHTML = `<i class="bi bi-x-circle me-1"></i>${escapeHtml(data.error || "No se pudo guardar el bloque")}`;
                    resultArea.classList.remove("d-none");
                    return;
                }

                const inserted = Number(data.inserted || 0);
                const skipped = Number(data.skipped || 0);
                const locationWarnings = Array.isArray(data.location_warnings) ? data.location_warnings : [];
                resultArea.className = "alert alert-success mt-3 small";
                resultArea.innerHTML = `<i class="bi bi-check-circle-fill me-1"></i>Bloque guardado: <strong>${inserted}</strong> filas registradas${skipped > 0 ? `, ${skipped} omitidas` : ""}.`;
                if (locationWarnings.length) {
                    const warningsHtml = locationWarnings.slice(0, 3).map((item) => {
                        const row = Number(item?.row_index || 0) + 1;
                        const warning = escapeHtml(String(item?.warning || ""));
                        const suggested = escapeHtml(String(item?.suggested_area || ""));
                        const location = escapeHtml(String(item?.ubicacion || ""));
                        return `<div class="mt-1"><i class="bi bi-exclamation-triangle-fill me-1 text-warning"></i><strong>Fila ${row}:</strong> ${warning}${suggested ? ` Se asignó a <span class="fw-semibold">${suggested}</span>.` : ""}${location ? ` <span class="text-muted">(${location})</span>` : ""}</div>`;
                    }).join("");
                    const extra = locationWarnings.length > 3
                        ? `<div class="mt-1 text-muted">... y ${locationWarnings.length - 3} advertencia(s) adicional(es).</div>`
                        : "";
                    resultArea.innerHTML += `<div class="mt-2 pt-2 border-top">${warningsHtml}${extra}</div>`;
                }
                resultArea.classList.remove("d-none");

                if (typeof onImportSuccess === "function") {
                    await onImportSuccess();
                }

                if (data.has_more) {
                    state.startIndex = Number(data.next_start_index || (state.startIndex + state.chunkSize));
                    await validateCurrentChunk();
                } else {
                    reviewSummary.innerHTML = '<span class="text-success fw-semibold">Importacion finalizada. No quedan mas bloques.</span>';
                    btnConfirm.classList.add("d-none");
                    btnValidate.classList.add("d-none");
                    btnDeleteRow.classList.add("d-none");
                }
            } catch (_error) {
                resultArea.className = "alert alert-danger mt-3 small";
                resultArea.innerHTML = '<i class="bi bi-x-circle me-1"></i>Error de red al guardar bloque.';
                resultArea.classList.remove("d-none");
            } finally {
                state.isSaving = false;
                btnConfirm.disabled = false;
                btnConfirm.innerHTML = '<i class="bi bi-cloud-upload me-1"></i>Guardar bloque y continuar';
            }
        }

        async function handleSkipCurrentChunk() {
            resultArea.classList.add("d-none");
            resultArea.innerHTML = "";

            if (!state.totalRows || state.totalRows <= 0) {
                await validateCurrentChunk();
                if (!state.totalRows || state.totalRows <= 0) {
                    return;
                }
            }

            const nextStart = state.startIndex + state.chunkSize;
            if (nextStart >= state.totalRows) {
                reviewSummary.innerHTML = '<span class="text-warning-emphasis fw-semibold">No hay mas bloques para revisar. El bloque actual fue omitido.</span>';
                chunkRangeEl.textContent = `Bloque actual: filas ${state.startIndex + 1} a ${Math.min(state.startIndex + state.chunkSize, state.totalRows)} de ${state.totalRows}`;
                btnConfirm.classList.add("d-none");
                btnValidate.classList.add("d-none");
                btnDeleteRow.classList.add("d-none");
                if (btnSkip) btnSkip.classList.add("d-none");
                return;
            }

            state.startIndex = nextStart;
            await validateCurrentChunk();
            resultArea.className = "alert alert-warning mt-3 small";
            resultArea.innerHTML = '<i class="bi bi-skip-forward me-1"></i>Bloque omitido sin guardar. Se cargaron las siguientes 20 filas para revision.';
            resultArea.classList.remove("d-none");
        }

        async function buildConfirmStep() {
            const mapped = state.mappings.filter((m) => m).length;
            const ignored = state.mappings.length - mapped;
            rowCountEl.innerHTML = `<strong>${state.totalRows}</strong> filas · <strong>${mapped}</strong> columnas mapeadas · <strong>${ignored}</strong> ignoradas`;

            state.startIndex = 0;
            state.reviewRows = [];
            state.selectedReviewRowIndex = null;
            state.removedRowSignaturesByBlock = {};
            state.invalidLocationsModalShownByBlock = {};
            reviewBody.innerHTML = "";
            reviewSummary.innerHTML = "";
            chunkRangeEl.textContent = "";
            resultArea.classList.add("d-none");
            resultArea.innerHTML = "";

            btnConfirm.disabled = true;
            btnConfirm.innerHTML = '<i class="bi bi-cloud-upload me-1"></i>Guardar bloque y continuar';
            btnConfirm.classList.add("d-none");
            if (btnSkip) {
                btnSkip.disabled = true;
                btnSkip.classList.add("d-none");
            }
            btnValidate.classList.remove("d-none");
            btnDeleteRow.classList.add("d-none");

            setReviewLoadingState(true, "Preparando tabla de revision y vinculando datos...");

            await loadAreasForSelect();
            await loadReviewSelectOptions();
            const ok = await validateCurrentChunk();
            setReviewLoadingState(false);

            if (!ok) {
                return;
            }

            btnConfirm.disabled = false;
            if (btnSkip) {
                btnSkip.disabled = false;
            }
        }

        btnNext.addEventListener("click", async () => {
            if (!state.mappings.some((m) => m)) {
                window.alert("Debes asignar al menos una columna a un campo del inventario.");
                return;
            }
            setStep(3);
            btnNext.disabled = true;
            try {
                await buildConfirmStep();
            } finally {
                btnNext.disabled = false;
            }
        });

        btnBack.addEventListener("click", () => {
            if (state.step === 2) setStep(1);
            else if (state.step === 3) setStep(2);
        });

        btnValidate.addEventListener("click", async () => {
            await validateCurrentChunk();
        });

        btnConfirm.addEventListener("click", async () => {
            const accepted = await openSaveConfirmModal();
            if (!accepted) {
                return;
            }
            await handleCommitCurrentChunk();
        });

        if (btnSkip) {
            btnSkip.addEventListener("click", async () => {
                if (state.isSaving) return;
                btnSkip.disabled = true;
                try {
                    await handleSkipCurrentChunk();
                } finally {
                    btnSkip.disabled = false;
                }
            });
        }

        if (importInvalidOpenSettingsBtn) {
            importInvalidOpenSettingsBtn.addEventListener("click", () => {
                state.shouldRefreshAreasOnFocus = true;
                startPendingAreaRefreshPolling();
                window.open("/ajustes", "_blank", "noopener,noreferrer");
            });
        }

        const handleVisibilityOrFocusSync = async () => {
            if (!state.shouldRefreshAreasOnFocus) {
                return;
            }
            if (document.visibilityState && document.visibilityState !== "visible") {
                return;
            }
            await refreshAreasAfterSettingsChange();
        };

        document.addEventListener("visibilitychange", handleVisibilityOrFocusSync);
        window.addEventListener("focus", handleVisibilityOrFocusSync);
        window.addEventListener("pageshow", handleVisibilityOrFocusSync);

        reviewBody.addEventListener("click", (event) => {
            const tr = event.target.closest("tr[data-local-index]");
            if (!tr) return;
            state.selectedReviewRowIndex = Number(tr.dataset.localIndex);
            updateSelectedRowVisual();
        });

        btnDeleteRow.addEventListener("click", () => {
            if (state.selectedReviewRowIndex == null) return;
            const removedRow = state.reviewRows[state.selectedReviewRowIndex];
            if (removedRow && removedRow.data) {
                getRemovedRowSignaturesForCurrentBlock().add(buildRowSignature(removedRow.data));
            }
            state.reviewRows.splice(state.selectedReviewRowIndex, 1);
            state.selectedReviewRowIndex = null;
            renderReviewRows();
            renderSummaryFromRows();
        });

        reviewBody.addEventListener("dblclick", (event) => {
            const cell = event.target.closest(".import-review-cell");
            if (!cell) return;
            const tr = cell.closest("tr[data-local-index]");
            if (!tr) return;

            const localIndex = Number(tr.dataset.localIndex);
            const row = state.reviewRows[localIndex];
            if (!row) return;

            const field = cell.dataset.field;
            if (!field || !EDITABLE_REVIEW_FIELDS.has(field)) return;

            const current = row.data[field] == null ? "" : String(row.data[field]);
            if (field === "ubicacion") {
                const select = document.createElement("select");
                select.className = "form-select form-select-sm import-editing-select";
                const custom = document.createElement("option");
                custom.value = current;
                custom.textContent = current ? `(Ubicacion actual: ${current})` : "(Sin ubicacion)";
                select.appendChild(custom);
                state.areaOptions.forEach((opt) => {
                    const option = document.createElement("option");
                    option.value = opt.label;
                    option.textContent = opt.label;
                    option.dataset.areaId = String(opt.id);
                    select.appendChild(option);
                });

                select.value = current;
                cell.replaceWith(select);
                select.focus();

                const previousValue = String(row.data[field] || "").trim();
                const previousAreaId = row.data.area_id == null ? "" : String(row.data.area_id);

                const commit = () => {
                    const selected = select.options[select.selectedIndex];
                    const value = String(select.value || "").trim();
                    let nextAreaId = row.data.area_id;
                    let nextAreaLabel = row.areaLabel;

                    if (selected && selected.dataset.areaId) {
                        nextAreaId = Number(selected.dataset.areaId);
                        nextAreaLabel = selected.textContent;
                    } else if (!value) {
                        nextAreaId = null;
                        nextAreaLabel = "";
                    } else {
                        const suggestion = suggestAreaId(value);
                        if (suggestion) {
                            nextAreaId = suggestion.id;
                            nextAreaLabel = suggestion.label;
                        }
                    }

                    const unchanged =
                        value === previousValue
                        && String(nextAreaId == null ? "" : nextAreaId) === previousAreaId;

                    if (unchanged) {
                        renderReviewRows();
                        return;
                    }

                    row.data[field] = value;
                    row.data.area_id = nextAreaId;
                    row.areaLabel = nextAreaLabel;
                    row.status = "normal";
                    evaluateRowLocationState(row);
                    renderReviewRows();
                    renderSummaryFromRows();
                };
                select.addEventListener("change", commit, { once: true });
                select.addEventListener("blur", commit, { once: true });
                return;
            }

            if (field === "cuenta" || field === "usuario_final" || field === "estado") {
                const select = document.createElement("select");
                select.className = "form-select form-select-sm import-editing-select";

                const base = document.createElement("option");
                base.value = "";
                base.textContent = field === "cuenta"
                    ? "-- Seleccionar cuenta --"
                    : field === "usuario_final"
                        ? "-- Seleccionar personal --"
                        : "-- Seleccionar estado --";
                select.appendChild(base);

                const options = field === "cuenta"
                    ? state.cuentaOptions
                    : field === "usuario_final"
                        ? state.usuarioOptions
                        : state.estadoOptions;
                options.forEach((name) => {
                    const option = document.createElement("option");
                    option.value = name;
                    option.textContent = name;
                    select.appendChild(option);
                });

                if (current && !options.some((name) => normalizeText(name) === normalizeText(current))) {
                    const currentOption = document.createElement("option");
                    currentOption.value = current;
                    currentOption.textContent = `${current} (actual)`;
                    select.appendChild(currentOption);
                }

                select.value = resolveSelectOptionValue(options, current) || current;
                cell.replaceWith(select);
                select.focus();

                const previousValue = String(row.data[field] || "").trim();

                const commit = () => {
                    const rawSelected = String(select.value || "").trim();
                    const resolvedSelected = resolveSelectOptionValue(options, rawSelected);
                    const nextValue = resolvedSelected || rawSelected;

                    if (nextValue === previousValue) {
                        renderReviewRows();
                        return;
                    }

                    row.data[field] = nextValue;
                    row.status = "normal";
                    evaluateRowLocationState(row);
                    renderReviewRows();
                    renderSummaryFromRows();
                };

                select.addEventListener("change", commit, { once: true });
                select.addEventListener("blur", commit, { once: true });
                return;
            }

            const input = document.createElement("input");
            input.className = "form-control form-control-sm import-editing-input";
            if (field === "fecha_adquisicion" || field === "fecha_adquisicion_esbye") {
                input.type = "date";
                input.value = normalizeDateInputValue(current);
            } else {
                input.value = current;
            }
            cell.replaceWith(input);
            input.focus();
            if (input.type !== "date") input.select();

            const previousValue = String(row.data[field] == null ? "" : row.data[field]).trim();
            const previousAreaId = row.data.area_id == null ? "" : String(row.data.area_id);

            const commit = () => {
                const rawValue = String(input.value || "").trim();
                const nextValue = (field === "cod_inventario" || field === "cod_esbye")
                    ? normalizeCodeToPlaceholder(rawValue)
                    : rawValue;

                let nextAreaId = row.data.area_id;
                let nextAreaLabel = row.areaLabel;

                if (field === "ubicacion") {
                    const suggestion = suggestAreaId(nextValue);
                    if (suggestion) {
                        nextAreaId = suggestion.id;
                        nextAreaLabel = suggestion.label;
                    } else {
                        nextAreaLabel = "";
                    }
                }

                const unchanged =
                    nextValue === previousValue
                    && String(nextAreaId == null ? "" : nextAreaId) === previousAreaId;

                if (unchanged) {
                    renderReviewRows();
                    return;
                }

                row.data[field] = nextValue;
                row.status = "normal";
                if (field === "ubicacion") {
                    row.data.area_id = nextAreaId;
                    row.areaLabel = nextAreaLabel;
                }
                evaluateRowLocationState(row);
                renderReviewRows();
                renderSummaryFromRows();
            };

            input.addEventListener("keydown", (e) => {
                if (e.key === "Enter") {
                    e.preventDefault();
                    commit();
                }
                if (e.key === "Escape") {
                    renderReviewRows();
                }
            });
            input.addEventListener("blur", commit, { once: true });
        });

        dropzone.addEventListener("dragover", (e) => {
            e.preventDefault();
            dropzone.classList.add("border-success", "bg-success-subtle");
        });
        dropzone.addEventListener("dragleave", () => {
            dropzone.classList.remove("border-success", "bg-success-subtle");
        });
        dropzone.addEventListener("drop", (e) => {
            e.preventDefault();
            dropzone.classList.remove("border-success", "bg-success-subtle");
            const file = e.dataTransfer.files[0];
            if (file) handleFile(file);
        });
        dropzone.addEventListener("click", () => fileInput.click());
        btnSelectFile.addEventListener("click", (e) => {
            e.stopPropagation();
            fileInput.click();
        });
        fileInput.addEventListener("change", () => {
            if (fileInput.files[0]) handleFile(fileInput.files[0]);
        });

        modal.addEventListener("hidden.bs.modal", () => {
            stopPendingAreaRefreshPolling();
            state = buildInitialState();
            clearStatus();
            fileInput.value = "";
            setStep(1);
            btnConfirm.classList.remove("d-none");
            btnValidate.classList.remove("d-none");
            btnDeleteRow.classList.add("d-none");
            if (btnSkip) {
                btnSkip.classList.remove("d-none");
            }
            if (importConflictsModal) {
                importConflictsModal.hide();
            }
            if (importSaveConfirmModal) {
                importSaveConfirmModal.hide();
            }
            if (importInvalidLocationsModal) {
                importInvalidLocationsModal.hide();
            }
        });

        const btnAbrir = document.getElementById("btn-importar-excel");
        if (btnAbrir) {
            btnAbrir.addEventListener("click", () => {
                state = buildInitialState();
                clearStatus();
                fileInput.value = "";
                setStep(1);
                bsModal.show();
            });
        }

        setStep(1);
    };
})();
