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
        const normalizedValue = normalizeText(value);
        const list = Array.isArray(options) ? options : [];
        if (!list.length) return "";

        let exact = list.find((name) => normalizeText(name) === normalizedValue);
        if (exact) return exact;

        let contains = list.find((name) => normalizeText(name).includes(normalizedValue) || normalizedValue.includes(normalizeText(name)));
        if (contains) return contains;

        const targetTokens = normalizedValue.split(/\s+/).filter(Boolean);
        let best = "";
        let bestScore = 0;
        list.forEach((name) => {
            const norm = normalizeText(name);
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

        const importConflictsModalEl = document.getElementById("modalImportConflicts");
        const importConflictsSummary = document.getElementById("import-conflicts-summary");
        const importConflictsAccordion = document.getElementById("import-conflicts-accordion");
        const importConflictsCancelBtn = document.getElementById("btn-import-conflicts-cancel");
        const importConflictsContinueBtn = document.getElementById("btn-import-conflicts-continue");
        const importConflictsModal = importConflictsModalEl ? new bootstrap.Modal(importConflictsModalEl) : null;

        const stepBadges = [1, 2, 3].map((n) => document.getElementById(`import-step-badge-${n}`));
        const stepLabels = [1, 2, 3].map((n) => document.getElementById(`import-step-label-${n}`));

        let state = buildInitialState();

        function buildInitialState() {
            return {
                step: 1,
                sessionId: null,
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
                isSaving: false,
            };
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

        async function handleFile(file) {
            if (!file.name.toLowerCase().endsWith(".xlsx")) {
                showStatus('<i class="bi bi-x-circle me-1"></i>Solo se aceptan archivos .xlsx.', "danger");
                return;
            }
            if (file.size > 10 * 1024 * 1024) {
                showStatus('<i class="bi bi-x-circle me-1"></i>El archivo supera 10 MB.', "danger");
                return;
            }

            showStatus(`<span class="spinner-border spinner-border-sm me-2" role="status"></span>Procesando <strong>${escapeHtml(file.name)}</strong>...`);
            const formData = new FormData();
            formData.append("file", file);

            try {
                const res = await fetch("/api/inventario/previsualizar-excel", { method: "POST", body: formData });
                const data = await res.json();
                if (!res.ok || data.error) {
                    showStatus(`<i class="bi bi-x-circle me-1"></i>${escapeHtml(data.error || "Error al procesar archivo")}`, "danger");
                    return;
                }

                state.sessionId = data.session_id;
                state.headers = data.headers || [];
                state.previewRows = data.preview_rows || [];
                state.totalRows = Number(data.total_rows || 0);
                const suggested = Array.isArray(data.suggested_mapping) ? data.suggested_mapping : [];
                const rawMappings = state.headers.map((_, idx) => suggested[idx] || "");
                state.mappings = normalizeEsbyeDuplicateMappings(state.headers, rawMappings);

                clearStatus();
                renderMappingStep();
                setStep(2);
            } catch (_error) {
                showStatus('<i class="bi bi-x-circle me-1"></i>Error de red al subir archivo.', "danger");
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
            const blockLetterMatch = (targetAulaCode || "").match(/\d+([a-z])/);
            const targetBlockLetter = blockLetterMatch ? blockLetterMatch[1] : "";

            let best = null;
            let bestScore = 0;
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

                const tokenHits = targetTokens.filter((tk) => tk.length >= 3 && areaTokens.some((ak) => ak.includes(tk) || tk.includes(ak))).length;
                if (tokenHits > 0) {
                    score += tokenHits * 2;
                }

                if (score > bestScore) {
                    bestScore = score;
                    best = opt;
                }
            });
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
                if (state.selectedReviewRowIndex === localIndex) {
                    tr.classList.add("import-row-selected");
                }

                const tdIndex = document.createElement("td");
                tdIndex.className = "small";
                tdIndex.textContent = String((row.row_index || 0) + 1);

                const tdStatus = document.createElement("td");
                tdStatus.innerHTML = `${getStatusBadge(row.status)} ${row.areaLabel ? `<div class="small text-muted mt-1">Area: ${escapeHtml(row.areaLabel)}</div>` : ""}`;

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
            areaSelect.innerHTML = '<option value="">Sin area especifica</option>';
            try {
                const res = await fetch("/api/estructura?include_details=0");
                const json = await res.json();
                state.areaOptions = flattenAreas(json.data || []);
                state.areaOptions.forEach((opt) => {
                    const option = document.createElement("option");
                    option.value = String(opt.id);
                    option.textContent = opt.label;
                    areaSelect.appendChild(option);
                });
            } catch (_err) {
                state.areaOptions = [];
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
                if (!res.ok || !data.success) {
                    resultArea.className = "alert alert-danger mt-3 small";
                    resultArea.innerHTML = `<i class="bi bi-x-circle me-1"></i>${escapeHtml(data.error || "No se pudo validar el bloque")}`;
                    resultArea.classList.remove("d-none");
                    return;
                }

                state.totalRows = Number(data.total_rows || state.totalRows || 0);
                state.reviewRows = Array.isArray(data.rows) ? data.rows : [];
                state.selectedReviewRowIndex = null;
                applyAreaSuggestionToRows();
                renderReviewSummary(data.summary || null);
                renderReviewRows();

                const from = state.startIndex + 1;
                const to = Math.min(state.startIndex + state.chunkSize, state.totalRows);
                chunkRangeEl.textContent = `Bloque actual: filas ${from} a ${to} de ${state.totalRows}`;
            } catch (_error) {
                resultArea.className = "alert alert-danger mt-3 small";
                resultArea.innerHTML = '<i class="bi bi-x-circle me-1"></i>Error de red al verificar bloque.';
                resultArea.classList.remove("d-none");
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
                if (res.status === 409) {
                    const summary = data.summary || {};
                    const goOn = await openImportConflictsModal(data);
                    if (!goOn) {
                        state.reviewRows = Array.isArray(data.rows) ? data.rows : state.reviewRows;
                        applyAreaSuggestionToRows();
                        renderReviewSummary(summary);
                        renderReviewRows();
                        return;
                    }
                    ({ res, data } = await saveCurrentChunk(true));
                }

                if (!res.ok || !data.success) {
                    resultArea.className = "alert alert-danger mt-3 small";
                    resultArea.innerHTML = `<i class="bi bi-x-circle me-1"></i>${escapeHtml(data.error || "No se pudo guardar el bloque")}`;
                    resultArea.classList.remove("d-none");
                    return;
                }

                const inserted = Number(data.inserted || 0);
                const skipped = Number(data.skipped || 0);
                resultArea.className = "alert alert-success mt-3 small";
                resultArea.innerHTML = `<i class="bi bi-check-circle-fill me-1"></i>Bloque guardado: <strong>${inserted}</strong> filas registradas${skipped > 0 ? `, ${skipped} omitidas` : ""}.`;
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

        async function buildConfirmStep() {
            const mapped = state.mappings.filter((m) => m).length;
            const ignored = state.mappings.length - mapped;
            rowCountEl.innerHTML = `<strong>${state.totalRows}</strong> filas · <strong>${mapped}</strong> columnas mapeadas · <strong>${ignored}</strong> ignoradas`;

            state.startIndex = 0;
            state.reviewRows = [];
            state.selectedReviewRowIndex = null;
            reviewBody.innerHTML = "";
            reviewSummary.innerHTML = "";
            chunkRangeEl.textContent = "";
            resultArea.classList.add("d-none");
            resultArea.innerHTML = "";

            btnConfirm.classList.remove("d-none");
            btnConfirm.disabled = false;
            btnConfirm.innerHTML = '<i class="bi bi-cloud-upload me-1"></i>Guardar bloque y continuar';
            btnValidate.classList.remove("d-none");
            btnDeleteRow.classList.add("d-none");

            await loadAreasForSelect();
            await loadReviewSelectOptions();
            await validateCurrentChunk();
        }

        btnNext.addEventListener("click", async () => {
            if (!state.mappings.some((m) => m)) {
                window.alert("Debes asignar al menos una columna a un campo del inventario.");
                return;
            }
            await buildConfirmStep();
            setStep(3);
        });

        btnBack.addEventListener("click", () => {
            if (state.step === 2) setStep(1);
            else if (state.step === 3) setStep(2);
        });

        btnValidate.addEventListener("click", async () => {
            await validateCurrentChunk();
        });

        btnConfirm.addEventListener("click", async () => {
            await handleCommitCurrentChunk();
        });

        reviewBody.addEventListener("click", (event) => {
            const tr = event.target.closest("tr[data-local-index]");
            if (!tr) return;
            state.selectedReviewRowIndex = Number(tr.dataset.localIndex);
            updateSelectedRowVisual();
        });

        btnDeleteRow.addEventListener("click", () => {
            if (state.selectedReviewRowIndex == null) return;
            state.reviewRows.splice(state.selectedReviewRowIndex, 1);
            state.selectedReviewRowIndex = null;
            renderReviewRows();
            renderReviewSummary({
                exact: state.reviewRows.filter((r) => r.status === "exact").length,
                similar: state.reviewRows.filter((r) => r.status === "similar").length,
                normal: state.reviewRows.filter((r) => r.status === "normal").length,
            });
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

                const commit = () => {
                    const selected = select.options[select.selectedIndex];
                    const value = String(select.value || "").trim();
                    row.data[field] = value;
                    if (selected && selected.dataset.areaId) {
                        row.data.area_id = Number(selected.dataset.areaId);
                        row.areaLabel = selected.textContent;
                    } else if (!value) {
                        row.data.area_id = null;
                        row.areaLabel = "";
                    } else {
                        const suggestion = suggestAreaId(value);
                        if (suggestion) {
                            row.data.area_id = suggestion.id;
                            row.areaLabel = suggestion.label;
                        }
                    }
                    row.status = "normal";
                    renderReviewRows();
                    renderReviewSummary({
                        exact: state.reviewRows.filter((r) => r.status === "exact").length,
                        similar: state.reviewRows.filter((r) => r.status === "similar").length,
                        normal: state.reviewRows.filter((r) => r.status === "normal").length,
                    });
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

                const commit = () => {
                    row.data[field] = String(select.value || "").trim();
                    row.status = "normal";
                    renderReviewRows();
                    renderReviewSummary({
                        exact: state.reviewRows.filter((r) => r.status === "exact").length,
                        similar: state.reviewRows.filter((r) => r.status === "similar").length,
                        normal: state.reviewRows.filter((r) => r.status === "normal").length,
                    });
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

            const commit = () => {
                const rawValue = String(input.value || "").trim();
                row.data[field] = (field === "cod_inventario" || field === "cod_esbye")
                    ? normalizeCodeToPlaceholder(rawValue)
                    : rawValue;
                row.status = "normal";
                if (field === "ubicacion") {
                    const suggestion = suggestAreaId(row.data.ubicacion);
                    if (suggestion) {
                        row.data.area_id = suggestion.id;
                        row.areaLabel = suggestion.label;
                    }
                }
                renderReviewRows();
                renderReviewSummary({
                    exact: state.reviewRows.filter((r) => r.status === "exact").length,
                    similar: state.reviewRows.filter((r) => r.status === "similar").length,
                    normal: state.reviewRows.filter((r) => r.status === "normal").length,
                });
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
            state = buildInitialState();
            clearStatus();
            fileInput.value = "";
            setStep(1);
            btnConfirm.classList.remove("d-none");
            btnValidate.classList.remove("d-none");
            btnDeleteRow.classList.add("d-none");
            if (importConflictsModal) {
                importConflictsModal.hide();
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
