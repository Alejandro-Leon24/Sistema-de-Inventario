/**
 * inventario-importar.js
 * Modal de importación masiva desde Excel (.xlsx) – 3 pasos:
 *   1. Seleccionar / arrastrar el archivo
 *   2. Mapear cada columna al campo canónico del sistema
 *   3. Confirmar e importar
 *
 * El archivo se sube al servidor (openpyxl) para el parseo; no se usa
 * ninguna librería de terceros en el navegador para leer el xlsx.
 */
(function () {
	"use strict";

	// ── Campos canónicos del inventario ──────────────────────────────────────
	const CANONICAL_FIELDS = [
		{ value: "", label: "— Ignorar columna —" },
		{ value: "cod_inventario", label: "Cód. Inventario" },
		{ value: "cod_esbye", label: "Cód. ESBYE" },
		{ value: "cuenta", label: "Cuenta" },
		{ value: "cantidad", label: "Cantidad" },
		{ value: "descripcion", label: "Descripción" },
		{ value: "ubicacion", label: "Ubicación" },
		{ value: "marca", label: "Marca" },
		{ value: "modelo", label: "Modelo" },
		{ value: "serie", label: "Serie" },
		{ value: "estado", label: "Estado" },
		{ value: "condicion", label: "Condición" },
		{ value: "usuario_final", label: "Usuario Final" },
		{ value: "fecha_adquisicion", label: "Fecha Adquisición" },
		{ value: "valor", label: "Valor ($)" },
		{ value: "observacion", label: "Observación" },
		{ value: "descripcion_esbye", label: "Descripción ESBYE" },
		{ value: "marca_esbye", label: "Marca ESBYE" },
		{ value: "modelo_esbye", label: "Modelo ESBYE" },
		{ value: "serie_esbye", label: "Serie ESBYE" },
		{ value: "fecha_adquisicion_esbye", label: "Fecha Adq. ESBYE" },
		{ value: "valor_esbye", label: "Valor ESBYE" },
		{ value: "ubicacion_esbye", label: "Ubicación ESBYE" },
		{ value: "observacion_esbye", label: "Observación ESBYE" },
	];

	// ── Reglas de auto-mapeo de encabezados Excel ────────────────────────────
	const AUTO_MAP_RULES = {
		cod_inventario: [
			"cod inventario", "cod. inventario", "codigo inventario",
			"código inventario", "cod_inventario", "codigo", "código",
			"n° inventario", "nro inventario", "numero inventario",
			"número inventario", "item", "n° bien",
		],
		cod_esbye: [
			"cod esbye", "cod. esbye", "codigo esbye", "código esbye",
			"cod_esbye", "esbye", "n° esbye", "nro esbye",
		],
		cuenta: ["cuenta", "tipo cuenta", "cuenta contable"],
		cantidad: ["cantidad", "cant", "cant.", "cantidad de bienes", "qty"],
		descripcion: [
			"descripcion", "descripción", "descripcion del bien",
			"descripción del bien", "bien", "detalle", "nombre del bien",
		],
		ubicacion: ["ubicacion", "ubicación", "ubicacion actual", "ubicacion fisica"],
		marca: ["marca", "fabricante"],
		modelo: ["modelo", "model"],
		serie: [
			"serie", "n° serie", "no. serie", "numero de serie",
			"número de serie", "serial", "n° serial",
		],
		estado: ["estado", "estado actual", "estado del bien"],
		condicion: ["condicion", "condición", "condicion del bien"],
		usuario_final: [
			"usuario final", "usuario", "custodio", "responsable",
			"asignado a", "custodio del bien",
		],
		fecha_adquisicion: [
			"fecha adquisicion", "fecha adquisición", "fecha de adquisicion",
			"fecha de adquisición", "fecha compra", "fecha",
		],
		valor: ["valor", "valor unitario", "precio", "valor ($)", "costo"],
		observacion: ["observacion", "observación", "observaciones", "notas", "nota"],
		descripcion_esbye: ["descripcion esbye", "descripción esbye"],
		marca_esbye: ["marca esbye"],
		modelo_esbye: ["modelo esbye"],
		serie_esbye: ["serie esbye"],
		fecha_adquisicion_esbye: [
			"fecha esbye", "fecha adq esbye", "fecha adquisicion esbye",
			"fecha adquisición esbye",
		],
		valor_esbye: ["valor esbye"],
		ubicacion_esbye: ["ubicacion esbye", "ubicación esbye"],
		observacion_esbye: ["observacion esbye", "observación esbye"],
	};

	/** Sugiere un campo canónico para un encabezado Excel dado. */
	function autoMapHeader(header) {
		const normalized = String(header || "").trim().toLowerCase();
		if (!normalized) return "";
		// Coincidencia exacta primero
		for (const [field, aliases] of Object.entries(AUTO_MAP_RULES)) {
			if (aliases.includes(normalized)) return field;
		}
		// Coincidencia parcial
		for (const [field, aliases] of Object.entries(AUTO_MAP_RULES)) {
			for (const alias of aliases) {
				if (normalized.includes(alias) || alias.includes(normalized)) return field;
			}
		}
		return "";
	}

	// ── Punto de entrada público ─────────────────────────────────────────────

	/**
	 * Inicializa el modal de importación de Excel.
	 * @param {object} opts
	 * @param {Function} [opts.onImportSuccess]  Se llama tras una importación exitosa.
	 */
	window.initInventarioImportar = function (opts) {
		const { onImportSuccess } = opts || {};

		const modal = document.getElementById("modalImportarExcel");
		if (!modal) return;
		const bsModal = new bootstrap.Modal(modal);

		// Panels
		const panel1 = document.getElementById("import-panel-1");
		const panel2 = document.getElementById("import-panel-2");
		const panel3 = document.getElementById("import-panel-3");

		// Step 1 elements
		const dropzone = document.getElementById("import-dropzone");
		const fileInput = document.getElementById("import-file-input");
		const btnSelectFile = document.getElementById("btn-import-select-file");
		const uploadStatus = document.getElementById("import-upload-status");

		// Step 2 elements
		const mappingSelectsRow = document.getElementById("import-mapping-selects-row");
		const mappingHeadersRow = document.getElementById("import-mapping-headers-row");
		const mappingPreviewBody = document.getElementById("import-mapping-preview-body");
		const previewInfo = document.getElementById("import-preview-info");

		// Step 3 elements
		const rowCountEl = document.getElementById("import-row-count");
		const areaSelect = document.getElementById("import-area-select");
		const resultArea = document.getElementById("import-result-area");

		// Footer buttons
		const btnBack = document.getElementById("btn-import-back");
		const btnNext = document.getElementById("btn-import-next");
		const btnConfirm = document.getElementById("btn-import-confirm");

		// Step badges / labels
		const stepBadges = [1, 2, 3].map(n => document.getElementById(`import-step-badge-${n}`));
		const stepLabels = [1, 2, 3].map(n => document.getElementById(`import-step-label-${n}`));

		// ── State ──────────────────────────────────────────────────────────────
		let state = buildInitialState();

		function buildInitialState() {
			return {
				step: 1,
				sessionId: null,
				headers: [],       // string[]
				previewRows: [],   // string[][]
				totalRows: 0,
				mappings: [],      // canonical field per column index ("" = ignore)
			};
		}

		// ── Step navigation ────────────────────────────────────────────────────
		function setStep(step) {
			state.step = step;

			// Panels
			[panel1, panel2, panel3].forEach((p, i) =>
				p.classList.toggle("d-none", i + 1 !== step)
			);

			// Step indicator badges
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

			// Footer buttons
			btnBack.classList.toggle("d-none", step === 1);
			btnNext.classList.toggle("d-none", step !== 2);
			btnConfirm.classList.toggle("d-none", step !== 3);
		}

		// ── Upload status ──────────────────────────────────────────────────────
		function showStatus(html, type = "info") {
			uploadStatus.className = `alert alert-${type} small mt-3 py-2`;
			uploadStatus.innerHTML = html;
			uploadStatus.classList.remove("d-none");
		}

		function clearStatus() {
			uploadStatus.classList.add("d-none");
			uploadStatus.innerHTML = "";
		}

		// ── File handling ──────────────────────────────────────────────────────
		dropzone.addEventListener("dragover", e => {
			e.preventDefault();
			dropzone.classList.add("border-success", "bg-success-subtle");
		});
		dropzone.addEventListener("dragleave", () => {
			dropzone.classList.remove("border-success", "bg-success-subtle");
		});
		dropzone.addEventListener("drop", e => {
			e.preventDefault();
			dropzone.classList.remove("border-success", "bg-success-subtle");
			const file = e.dataTransfer.files[0];
			if (file) handleFile(file);
		});
		dropzone.addEventListener("click", () => fileInput.click());
		btnSelectFile.addEventListener("click", e => {
			e.stopPropagation();
			fileInput.click();
		});
		fileInput.addEventListener("change", () => {
			if (fileInput.files[0]) handleFile(fileInput.files[0]);
		});

		async function handleFile(file) {
			if (!file.name.toLowerCase().endsWith(".xlsx")) {
				showStatus(
					'<i class="bi bi-x-circle me-1"></i>Solo se aceptan archivos <strong>.xlsx</strong>.',
					"danger"
				);
				return;
			}
			if (file.size > 10 * 1024 * 1024) {
				showStatus(
					'<i class="bi bi-x-circle me-1"></i>El archivo supera el límite de <strong>10 MB</strong>.',
					"danger"
				);
				return;
			}

			showStatus(
				`<span class="spinner-border spinner-border-sm me-2" role="status"></span>` +
				`Procesando <strong>${escapeHtml(file.name)}</strong>…`
			);

			const formData = new FormData();
			formData.append("file", file);

			try {
				const res = await fetch("/api/inventario/previsualizar-excel", {
					method: "POST",
					body: formData,
				});
				const data = await res.json();
				if (!res.ok || data.error) {
					showStatus(
						`<i class="bi bi-x-circle me-1"></i>${escapeHtml(data.error || "Error al procesar el archivo.")}`,
						"danger"
					);
					return;
				}
				state.sessionId = data.session_id;
				state.headers = data.headers || [];
				state.previewRows = data.preview_rows || [];
				state.totalRows = data.total_rows || 0;
				state.mappings = state.headers.map(h => autoMapHeader(h));

				clearStatus();
				renderMappingStep();
				setStep(2);
			} catch (_) {
				showStatus(
					'<i class="bi bi-x-circle me-1"></i>Error de red al subir el archivo. Inténtalo nuevamente.',
					"danger"
				);
			}
		}

		// ── Step 2: mapping table ──────────────────────────────────────────────
		function renderMappingStep() {
			mappingSelectsRow.innerHTML = "";
			mappingHeadersRow.innerHTML = "";
			mappingPreviewBody.innerHTML = "";

			state.headers.forEach((header, colIdx) => {
				// Dropdown row
				const thSel = document.createElement("th");
				thSel.style.minWidth = "165px";
				thSel.style.padding = "4px 6px";
				thSel.appendChild(buildColumnSelect(colIdx));
				mappingSelectsRow.appendChild(thSel);

				// Original Excel header label row
				const thHdr = document.createElement("th");
				thHdr.textContent = header || `(col ${colIdx + 1})`;
				thHdr.className = "text-muted small fw-normal";
				thHdr.style.padding = "2px 6px";
				mappingHeadersRow.appendChild(thHdr);
			});

			// Data preview rows
			state.previewRows.forEach(row => {
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

			const shown = state.previewRows.length;
			const total = state.totalRows;
			previewInfo.textContent =
				`Vista previa: ${shown} de ${total} fila${total !== 1 ? "s" : ""} de datos.` +
				(total > shown ? ` (Se mostrarán las primeras ${shown}.)` : "");

			applyColumnHighlights();
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
				allDataTrs.forEach(tr => {
					if (tr.children[colIdx]) tr.children[colIdx].classList.toggle("table-secondary", ignored);
				});
			});
		}

		// ── Step 3: confirm ────────────────────────────────────────────────────
		async function buildConfirmStep() {
			const mapped = state.mappings.filter(m => m).length;
			const ignored = state.mappings.length - mapped;
			rowCountEl.innerHTML =
				`<strong>${state.totalRows}</strong> fila${state.totalRows !== 1 ? "s" : ""} · ` +
				`<strong>${mapped}</strong> columna${mapped !== 1 ? "s" : ""} mapeada${mapped !== 1 ? "s" : ""} · ` +
				`<strong>${ignored}</strong> ignorada${ignored !== 1 ? "s" : ""}`;

			// Populate area selector
			areaSelect.innerHTML = '<option value="">Sin área específica</option>';
			try {
				const res = await fetch("/api/estructura");
				const json = await res.json();
				(json.data || []).forEach(bloque => {
					(bloque.pisos || []).forEach(piso => {
						(piso.areas || []).forEach(area => {
							const opt = document.createElement("option");
							opt.value = area.id;
							opt.textContent = `${bloque.nombre} / ${piso.nombre} / ${area.nombre}`;
							areaSelect.appendChild(opt);
						});
					});
				});
			} catch (_) {
				// Area selector stays with just the default option
			}

			resultArea.classList.add("d-none");
			resultArea.innerHTML = "";
			btnConfirm.disabled = false;
			btnConfirm.innerHTML = '<i class="bi bi-cloud-upload me-1"></i>Importar ahora';
		}

		// ── Navigation buttons ─────────────────────────────────────────────────
		btnNext.addEventListener("click", () => {
			if (!state.mappings.some(m => m)) {
				alert("Debes asignar al menos una columna a un campo del inventario.");
				return;
			}
			buildConfirmStep();
			setStep(3);
		});

		btnBack.addEventListener("click", () => {
			if (state.step === 2) setStep(1);
			else if (state.step === 3) setStep(2);
		});

		btnConfirm.addEventListener("click", async () => {
			btnConfirm.disabled = true;
			btnConfirm.innerHTML =
				'<span class="spinner-border spinner-border-sm me-1" role="status"></span>Importando…';

			const mapping = {};
			state.mappings.forEach((field, idx) => {
				mapping[String(idx)] = field || "";
			});
			const areaId = areaSelect.value ? parseInt(areaSelect.value, 10) : null;

			try {
				const res = await fetch("/api/inventario/confirmar-importacion", {
					method: "POST",
					headers: { "Content-Type": "application/json" },
					body: JSON.stringify({
						session_id: state.sessionId,
						mapping,
						area_id: areaId,
					}),
				});
				const data = await res.json();

				if (!res.ok || !data.success) {
					resultArea.className = "alert alert-danger mt-3 small";
					resultArea.innerHTML =
						`<i class="bi bi-x-circle me-1"></i>${escapeHtml(data.error || "Error al importar.")}`;
					resultArea.classList.remove("d-none");
					btnConfirm.disabled = false;
					btnConfirm.innerHTML = '<i class="bi bi-cloud-upload me-1"></i>Importar ahora';
				} else {
					const ins = data.inserted;
					const skip = data.skipped || 0;
					resultArea.className = "alert alert-success mt-3 small";
					resultArea.innerHTML =
						`<i class="bi bi-check-circle-fill me-1"></i>` +
						`<strong>¡Importación exitosa!</strong> ` +
						`Se importaron <strong>${ins}</strong> bien${ins !== 1 ? "es" : ""}.` +
						(skip > 0 ? `<br><span class="text-muted">${skip} fila${skip !== 1 ? "s fueron omitidas" : " fue omitida"} (sin datos válidos).</span>` : "");
					resultArea.classList.remove("d-none");
					btnConfirm.classList.add("d-none");
					if (typeof onImportSuccess === "function") onImportSuccess();
				}
			} catch (_) {
				resultArea.className = "alert alert-danger mt-3 small";
				resultArea.innerHTML =
					'<i class="bi bi-x-circle me-1"></i>Error de red. Inténtalo nuevamente.';
				resultArea.classList.remove("d-none");
				btnConfirm.disabled = false;
				btnConfirm.innerHTML = '<i class="bi bi-cloud-upload me-1"></i>Importar ahora';
			}
		});

		// ── Reset on close ─────────────────────────────────────────────────────
		modal.addEventListener("hidden.bs.modal", () => {
			state = buildInitialState();
			clearStatus();
			fileInput.value = "";
			setStep(1);
		});

		// ── Open button ────────────────────────────────────────────────────────
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

	// ── Pequeño helper de escape HTML ─────────────────────────────────────────
	function escapeHtml(text) {
		return String(text || "")
			.replace(/&/g, "&amp;")
			.replace(/</g, "&lt;")
			.replace(/>/g, "&gt;")
			.replace(/"/g, "&quot;");
	}
})();
