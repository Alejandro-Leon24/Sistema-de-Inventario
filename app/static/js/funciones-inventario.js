const INVENTORY_COLUMNS = [
	{ field: "item_numero", label: "ITEM", editable: false },
	{ field: "cod_inventario", label: "CÓD. INVENTARIO", editable: true },
	{ field: "cod_esbye", label: "CÓD. ESBYE", editable: true },
	{ field: "cuenta", label: "CUENTA", editable: true },
	{ field: "cantidad", label: "CANTIDAD", editable: true },
	{ field: "descripcion", label: "DESCRIPCIÓN", editable: true },
	{ field: "marca", label: "MARCA", editable: true },
	{ field: "modelo", label: "MODELO", editable: true },
	{ field: "serie", label: "SERIE", editable: true },
	{ field: "estado", label: "ESTADO", editable: true },
	{ field: "ubicacion", label: "UBICACIÓN", editable: true },
	{ field: "fecha_adquisicion", label: "FECHA DE ADQUISICIÓN", editable: true },
	{ field: "valor", label: "VALOR", editable: true },
	{ field: "usuario_final", label: "USUARIO FINAL", editable: true },
	{ field: "observacion", label: "OBSERVACIÓN", editable: true },
	{ field: "descripcion_esbye", label: "DESCRIPCIÓN ESBYE", editable: true },
	{ field: "marca_esbye", label: "MARCA ESBYE", editable: true },
	{ field: "modelo_esbye", label: "MODELO ESBYE", editable: true },
	{ field: "serie_esbye", label: "SERIE ESBYE", editable: true },
	{ field: "fecha_adquisicion_esbye", label: "FECHA ESBYE", editable: true },
	{ field: "valor_esbye", label: "VALOR ESBYE", editable: true },
	{ field: "ubicacion_esbye", label: "UBICACIÓN ESBYE", editable: true },
	{ field: "observacion_esbye", label: "OBSERVACIÓN ESBYE", editable: true },
];

const api = {
	async get(url) {
		const response = await fetch(url);
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

function debounce(callback, delay = 350) {
	let timer;
	return (...args) => {
		clearTimeout(timer);
		timer = setTimeout(() => callback(...args), delay);
	};
}

function parseExcelText(text) {
	return text
		.trim()
		.split("\n")
		.filter((line) => line.trim().length > 0)
		.map((line) => line.split("\t").map((cell) => cell.trim()));
}

function notify(message, isError = false) {
	window.alert(message);
	if (isError) {
		console.error(message);
	}
}

function parseDecimalWithComma(value) {
	const raw = String(value ?? "").trim();
	if (!raw) return null;
	const normalized = raw.replace(/\./g, "").replace(",", ".");
	const number = Number(normalized);
	return Number.isFinite(number) ? number : null;
}

function formatValue(field, value) {
	if (value === null || value === undefined) return "";
	if ((field === "valor" || field === "valor_esbye") && value !== "") {
		return Number(value).toLocaleString("es-EC", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
	}
	return value;
}

function escapeHtmlText(value) {
	return String(value || "")
		.replaceAll("&", "&amp;")
		.replaceAll("<", "&lt;")
		.replaceAll(">", "&gt;")
		.replaceAll('"', "&quot;")
		.replaceAll("'", "&#39;");
}

function buildDuplicateWarningMessage(duplicates = [], payload = {}) {
	const maxRows = 8;
	const rows = duplicates.slice(0, maxRows);
	const targetInv = String(payload.cod_inventario || "").trim();
	const targetEsbye = String(payload.cod_esbye || "").trim();
	const header = [
		"⚠️ Se detectaron códigos repetidos.",
		targetInv ? `Código inventario ingresado: ${targetInv}` : null,
		targetEsbye ? `Código ESBYE ingresado: ${targetEsbye}` : null,
		"",
		"Coincidencias encontradas:",
	]
		.filter(Boolean)
		.join("\n");

	const detailLines = rows.map((item, index) => {
		const matchLabel = Array.isArray(item.matches) && item.matches.length
			? `(${item.matches.map((m) => (m === "cod_inventario" ? "INV" : "ESBYE")).join(" + ")})`
			: "";
		return [
			`${index + 1}. Ítem #${item.item_numero || "-"} ${matchLabel}`,
			`   Nombre: ${item.descripcion || "-"}`,
			`   Modelo: ${item.modelo || "-"}`,
			`   Ubicación: ${item.ubicacion || "-"}`,
			`   Fecha: ${item.fecha_adquisicion || "-"}`,
			`   Usuario final: ${item.usuario_final || "-"}`,
		].join("\n");
	});

	const more = duplicates.length > maxRows ? `\n... y ${duplicates.length - maxRows} coincidencias más.` : "";
	return `${header}\n${detailLines.join("\n\n")}${more}\n\n¿Deseas agregar este nuevo ítem de todas formas?`;
}

async function initInventoryPage() {
	const body = document.getElementById("inventario-body");
	if (!body) return;

	const state = {
		structure: [],
		items: [],
		visibleItems: [],
		selectedRowId: null,
		activeBlockId: "",
		activeFloorId: "",
		activeAreaId: "",
		order: "asc",
		search: "",
		page: 1,
		perPage: 50,
		totalItems: 0,
		totalPages: 1,
		tableDensity: "compact",
		columnOrder: INVENTORY_COLUMNS.map((column) => column.field),
		columnWidths: {},
		headerDrag: {
			active: false,
			field: null,
			targetField: null,
			startX: 0,
			startY: 0,
			moved: false,
		},
	};

	const nodes = {
		filtroBloque: document.getElementById("filtro-bloque"),
		filtroPiso: document.getElementById("filtro-piso"),
		filtroArea: document.getElementById("filtro-area"),
		filtroOrder: document.getElementById("filtro-order"),
		filtroSearch: document.getElementById("filtro-search"),
		exportExcelBtn: document.getElementById("btn-exportar-excel"),
		toggleDensityBtn: document.getElementById("btn-toggle-density"),
		gridWrapper: document.getElementById("inventory-grid-wrapper"),
		table: document.getElementById("tabla-inventario"),
		tableHeadRow: document.querySelector("#tabla-inventario thead tr"),
		addInlineBtn: document.getElementById("add-inline-row"),
		detailContainer: document.getElementById("detalle-contenido"),
		contextMenu: document.getElementById("context-menu"),
		modalAreaSelect: document.getElementById("modal-area-select"),
		modalCuentaSelect: document.getElementById("modal-cuenta-select"),
		modalUsuarioFinalSelect: document.getElementById("modal-usuario-final-select"),
		modalUbicacion: document.getElementById("modal-ubicacion"),
		modalAddButton: document.getElementById("btn-guardar-item"),
		excelSingleRow: document.getElementById("excel-single-row"),
		pageInfo: document.getElementById("inventory-pagination-info"),
		pageIndicator: document.getElementById("inventory-page-indicator"),
		pagePrev: document.getElementById("inventory-page-prev"),
		pageNext: document.getElementById("inventory-page-next"),
		pageSize: document.getElementById("inventory-page-size"),
		duplicateModalEl: document.getElementById("modalDuplicadosInventario"),
		duplicateSummary: document.getElementById("duplicados-resumen"),
		duplicateList: document.getElementById("duplicados-lista"),
		duplicateCancelBtn: document.getElementById("btn-duplicados-cancelar"),
		duplicateContinueBtn: document.getElementById("btn-duplicados-continuar"),
	};

	const detailModal = new bootstrap.Modal(document.getElementById("modalDetalle"));
	const addModalElement = document.getElementById("modalAgregarItem");
	const addModal = new bootstrap.Modal(addModalElement);
	const duplicateModal = nodes.duplicateModalEl ? new bootstrap.Modal(nodes.duplicateModalEl) : null;

	function buildDuplicateRowsHtml(duplicates = []) {
		if (!duplicates.length) {
			return '<div class="alert alert-secondary mb-0">No hay coincidencias para mostrar.</div>';
		}
		return duplicates
			.map((item) => {
				const matchBadge = Array.isArray(item.matches) && item.matches.length
					? item.matches
						.map((match) => `<span class="badge text-bg-warning me-1">${match === "cod_inventario" ? "INV" : "ESBYE"}</span>`)
						.join("")
					: '<span class="badge text-bg-secondary">Coincidencia</span>';
				return `
					<div class="card border-warning-subtle shadow-sm">
						<div class="card-body py-2 px-3">
							<div class="d-flex justify-content-between align-items-start gap-2 mb-1">
								<div class="fw-semibold">Ítem #${escapeHtmlText(item.item_numero || "-")}</div>
								<div>${matchBadge}</div>
							</div>
							<div class="small"><strong>Nombre:</strong> ${escapeHtmlText(item.descripcion || "-")}</div>
							<div class="small"><strong>Modelo:</strong> ${escapeHtmlText(item.modelo || "-")}</div>
							<div class="small"><strong>Ubicación:</strong> ${escapeHtmlText(item.ubicacion || "-")}</div>
							<div class="small"><strong>Fecha:</strong> ${escapeHtmlText(item.fecha_adquisicion || "-")}</div>
							<div class="small"><strong>Usuario final:</strong> ${escapeHtmlText(item.usuario_final || "-")}</div>
						</div>
					</div>
				`;
			})
			.join("");
	}

	function openDuplicateModal({ duplicates = [], payload = {}, mode = "create" } = {}) {
		if (!duplicateModal || !nodes.duplicateSummary || !nodes.duplicateList || !nodes.duplicateContinueBtn || !nodes.duplicateCancelBtn) {
			return Promise.resolve(window.confirm(buildDuplicateWarningMessage(duplicates, payload)));
		}

		const inv = String(payload.cod_inventario || "").trim();
		const esbye = String(payload.cod_esbye || "").trim();
		const modeText = mode === "update" ? "guardar el cambio" : "registrar este nuevo ítem";
		nodes.duplicateSummary.textContent = `Se encontraron ${duplicates.length} coincidencias para ${modeText}.${inv ? ` INV: ${inv}` : ""}${esbye ? ` ESBYE: ${esbye}` : ""}`;
		nodes.duplicateList.innerHTML = buildDuplicateRowsHtml(duplicates);

		return new Promise((resolve) => {
			let settled = false;

			const cleanup = () => {
				nodes.duplicateContinueBtn.removeEventListener("click", handleContinue);
				nodes.duplicateCancelBtn.removeEventListener("click", handleCancel);
				nodes.duplicateModalEl.removeEventListener("hidden.bs.modal", handleHidden);
			};

			const finish = (value) => {
				if (settled) return;
				settled = true;
				cleanup();
				resolve(value);
			};

			const handleContinue = () => {
				finish(true);
				duplicateModal.hide();
			};
			const handleCancel = () => {
				finish(false);
				duplicateModal.hide();
			};
			const handleHidden = () => finish(false);

			nodes.duplicateContinueBtn.addEventListener("click", handleContinue);
			nodes.duplicateCancelBtn.addEventListener("click", handleCancel);
			nodes.duplicateModalEl.addEventListener("hidden.bs.modal", handleHidden);
			duplicateModal.show();
		});
	}

	function getColumn(field) {
		return INVENTORY_COLUMNS.find((column) => column.field === field);
	}

	function applyDensityMode() {
		if (!nodes.gridWrapper || !nodes.toggleDensityBtn) return;
		const isCompact = state.tableDensity === "compact";
		nodes.gridWrapper.classList.toggle("table-compact", isCompact);
		nodes.gridWrapper.classList.toggle("table-normal", !isCompact);
		nodes.toggleDensityBtn.innerHTML = isCompact
			? '<i class="bi bi-arrows-collapse me-2"></i>Modo: Compacta'
			: '<i class="bi bi-arrows-expand me-2"></i>Modo: Normal';
		nodes.toggleDensityBtn.title = isCompact
			? 'Cambiar a vista normal'
			: 'Cambiar a vista compacta';
	}

	function getOrderedColumns() {
		return state.columnOrder.map((field) => getColumn(field)).filter(Boolean);
	}

	async function saveColumnPreference() {
		await api.send("/api/preferencias", "PATCH", {
			pref_key: "inventory_column_order",
			pref_value: state.columnOrder,
		});
	}

	async function saveColumnWidthsPreference() {
		await api.send("/api/preferencias", "PATCH", {
			pref_key: "inventory_column_widths",
			pref_value: state.columnWidths,
		});
	}

	async function savePageSizePreference() {
		await api.send("/api/preferencias", "PATCH", {
			pref_key: "inventory_page_size",
			pref_value: state.perPage,
		});
	}

	async function saveDensityPreference() {
		await api.send("/api/preferencias", "PATCH", {
			pref_key: "inventory_table_density",
			pref_value: state.tableDensity,
		});
	}

	function applyColumnWidthToField(field, widthPx) {
		const px = `${widthPx}px`;
		nodes.table.querySelectorAll(`[data-field='${field}']`).forEach((cell) => {
			cell.style.width = px;
			cell.style.minWidth = px;
			cell.style.maxWidth = px;
		});
	}

	function renderTableHead() {
		const columns = getOrderedColumns();
		nodes.tableHeadRow.innerHTML = "";
		columns.forEach((column) => {
			const th = document.createElement("th");
			th.innerHTML = `<span class="head-label">${escapeHtmlText(column.label)}</span><span class="column-resize-handle" title="Ajustar ancho"></span>`;
			th.dataset.field = column.field;
			th.draggable = true;
			th.classList.add("inventory-head-cell");
			nodes.tableHeadRow.appendChild(th);
			if (state.columnWidths[column.field]) {
				applyColumnWidthToField(column.field, state.columnWidths[column.field]);
			}
		});
	}

	function renderRows() {
		const columns = getOrderedColumns();
		body.innerHTML = "";
		state.items.forEach((item) => {
			const tr = document.createElement("tr");
			tr.dataset.id = item.id;
			columns.forEach((column) => {
				const td = document.createElement("td");
				td.dataset.field = column.field;
				td.dataset.id = item.id;
				const displayValue = formatValue(column.field, item[column.field]);
				td.textContent = displayValue;
				td.title = String(displayValue || "");
				td.classList.add("inventory-cell");
				if (column.editable) td.classList.add("editable-cell");
				if (column.field === "item_numero") td.classList.add("fw-bold", "text-primary");
				if (state.columnWidths[column.field]) {
					const px = `${state.columnWidths[column.field]}px`;
					td.style.width = px;
					td.style.minWidth = px;
					td.style.maxWidth = px;
				}
				tr.appendChild(td);
			});
			body.appendChild(tr);
		});
	}

	function renderPaginationMeta() {
		if (!nodes.pageInfo || !nodes.pageIndicator || !nodes.pagePrev || !nodes.pageNext) return;
		const total = state.totalItems || 0;
		const currentPage = Math.max(state.page, 1);
		const perPage = Math.max(state.perPage, 1);
		const first = total === 0 ? 0 : (currentPage - 1) * perPage + 1;
		const last = total === 0 ? 0 : Math.min(currentPage * perPage, total);
		const totalPages = Math.max(state.totalPages || 1, 1);

		nodes.pageInfo.textContent = `Mostrando ${first} - ${last} de ${total}`;
		nodes.pageIndicator.textContent = `${currentPage} / ${totalPages}`;
		nodes.pagePrev.disabled = currentPage <= 1;
		nodes.pageNext.disabled = currentPage >= totalPages;
		if (nodes.pageSize) nodes.pageSize.value = String(perPage);
	}

	function fillSelect(select, list, placeholder = "Todos") {
		const previous = select.value;
		select.innerHTML = `<option value="">${placeholder}</option>`;
		list.forEach((item) => {
			const option = document.createElement("option");
			option.value = item.id;
			option.textContent = item.nombre;
			select.appendChild(option);
		});
		select.value = previous;
	}

	function fillSelectStrict(select, list, options = {}) {
		const { placeholder = "Seleccione", enabled = true } = options;
		const previous = select.value;
		select.innerHTML = `<option value="">${placeholder}</option>`;
		list.forEach((item) => {
			const option = document.createElement("option");
			option.value = item.id;
			option.textContent = item.nombre;
			select.appendChild(option);
		});
		select.disabled = !enabled;
		if (enabled && previous && list.some((item) => String(item.id) === String(previous))) {
			select.value = previous;
		} else {
			select.value = "";
		}
	}

	function getFloorsByBlock(blockId) {
		const block = state.structure.find((entry) => String(entry.id) === String(blockId));
		return block ? block.pisos : [];
	}

	function getAreasByFloor(blockId, floorId) {
		if (blockId) {
			const floors = getFloorsByBlock(blockId);
			const floor = floors.find((entry) => String(entry.id) === String(floorId));
			return floor ? floor.areas : [];
		}

		for (const block of state.structure) {
			const floor = block.pisos.find((entry) => String(entry.id) === String(floorId));
			if (floor) return floor.areas;
		}
		return [];
	}

	function flattenAreas() {
		const result = [];
		state.structure.forEach((block) => {
			block.pisos.forEach((floor) => {
				floor.areas.forEach((area) => {
					result.push({
						id: area.id,
						nombre: `${block.nombre} / ${floor.nombre} / ${area.nombre}`,
						bloque_nombre: block.nombre,
						piso_nombre: floor.nombre,
						area_nombre: area.nombre,
					});
				});
			});
		});
		return result;
	}

	function composeSelectedLocation() {
		const selectedBlock = state.structure.find((block) => String(block.id) === String(state.activeBlockId));
		const selectedFloor = selectedBlock
			? selectedBlock.pisos.find((floor) => String(floor.id) === String(state.activeFloorId))
			: null;
		const selectedArea = selectedFloor
			? selectedFloor.areas.find((area) => String(area.id) === String(state.activeAreaId))
			: null;

		const parts = [selectedBlock?.nombre, selectedFloor?.nombre, selectedArea?.nombre].filter(Boolean);
		return parts.join(" / ");
	}

	function syncModalLocationFromSelection() {
		if (!nodes.modalUbicacion) return;
		nodes.modalUbicacion.value = composeSelectedLocation();
	}

	function normalizeText(value) {
		return String(value || "")
			.normalize("NFD")
			.replace(/[\u0300-\u036f]/g, "")
			.toLowerCase()
			.trim();
	}

	function toInputDate(value) {
		const raw = String(value || "").trim();
		if (!raw) return "";

		if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
			return raw;
		}

		if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(raw)) {
			const [day, month, year] = raw.split(/[\/\-]/).map((part) => Number(part));
			const d = String(day).padStart(2, "0");
			const m = String(month).padStart(2, "0");
			return `${year}-${m}-${d}`;
		}

		if (/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(raw)) {
			const [year, month, day] = raw.split(/[\/\-]/).map((part) => Number(part));
			const d = String(day).padStart(2, "0");
			const m = String(month).padStart(2, "0");
			return `${year}-${m}-${d}`;
		}

		if (/^\d{5}(?:\.\d+)?$/.test(raw)) {
			const excelSerial = Number(raw);
			if (!Number.isNaN(excelSerial)) {
				const excelEpoch = new Date(Date.UTC(1899, 11, 30));
				excelEpoch.setUTCDate(excelEpoch.getUTCDate() + Math.floor(excelSerial));
				const year = excelEpoch.getUTCFullYear();
				const month = String(excelEpoch.getUTCMonth() + 1).padStart(2, "0");
				const day = String(excelEpoch.getUTCDate()).padStart(2, "0");
				return `${year}-${month}-${day}`;
			}
		}

		const parsed = new Date(raw);
		if (!Number.isNaN(parsed.getTime())) {
			const year = parsed.getFullYear();
			const month = String(parsed.getMonth() + 1).padStart(2, "0");
			const day = String(parsed.getDate()).padStart(2, "0");
			return `${year}-${month}-${day}`;
		}

		return "";
	}

	function assignPastedValue(input, rawValue) {
		if (!input) return;
		const value = String(rawValue ?? "").trim();

		if (input.tagName === "SELECT") {
			const options = Array.from(input.options || []);
			if (!value) {
				input.value = "";
				return;
			}

			const normalizedValue = normalizeText(value);
			let match = options.find((opt) => normalizeText(opt.value) === normalizedValue);
			if (!match) {
				match = options.find((opt) => normalizeText(opt.textContent) === normalizedValue);
			}
			if (!match) {
				match = options.find((opt) => normalizeText(opt.textContent).includes(normalizedValue));
			}
			if (!match) {
				match = options.find((opt) => normalizedValue.includes(normalizeText(opt.value)));
			}

			if (match) {
				input.value = match.value;
			}
			return;
		}

		if ((input.type || "").toLowerCase() === "date") {
			input.value = toInputDate(value);
			return;
		}

		if ((input.type || "").toLowerCase() === "number") {
			let normalized = value;
			if (value.includes(",") && value.includes(".")) {
				normalized = value.replace(/\./g, "").replace(",", ".");
			} else if (value.includes(",")) {
				normalized = value.replace(",", ".");
			}
			input.value = normalized;
			return;
		}

		input.value = value;
	}

	function renderAreaModalSelect() {
		const areas = flattenAreas();
		nodes.modalAreaSelect.innerHTML = '<option value="">Sin área</option>';
		areas.forEach((area) => {
			const option = document.createElement("option");
			option.value = area.id;
			option.textContent = area.nombre;
			nodes.modalAreaSelect.appendChild(option);
		});
	}

	function updateFloorAndAreaFilters() {
		if (!state.activeBlockId) {
			fillSelectStrict(nodes.filtroPiso, [], { placeholder: "Seleccione un bloque primero", enabled: false });
			fillSelectStrict(nodes.filtroArea, [], { placeholder: "Seleccione un piso primero", enabled: false });
			syncModalLocationFromSelection();
			return;
		}

		const floors = getFloorsByBlock(state.activeBlockId);
		fillSelectStrict(nodes.filtroPiso, floors, { placeholder: "Todos", enabled: true });

		if (!state.activeFloorId) {
			fillSelectStrict(nodes.filtroArea, [], { placeholder: "Seleccione un piso primero", enabled: false });
			syncModalLocationFromSelection();
			return;
		}

		const areas = getAreasByFloor(state.activeBlockId, state.activeFloorId);
		fillSelectStrict(nodes.filtroArea, areas, { placeholder: "Todas", enabled: true });
		syncModalLocationFromSelection();
	}

	async function loadStructure() {
		const response = await api.get("/api/estructura");
		state.structure = response.data;
		fillSelect(nodes.filtroBloque, state.structure, "Todos");
		updateFloorAndAreaFilters();
		renderAreaModalSelect();
	}

	async function loadPreferences() {
		const response = await api.get("/api/preferencias");
		const prefOrder = response.data.inventory_column_order;
		if (Array.isArray(prefOrder) && prefOrder.length) {
			const valid = prefOrder.filter((field) => INVENTORY_COLUMNS.some((column) => column.field === field));
			if (valid.length === INVENTORY_COLUMNS.length) {
				state.columnOrder = valid;
			}
		}

		const prefWidths = response.data.inventory_column_widths;
		if (prefWidths && typeof prefWidths === "object" && !Array.isArray(prefWidths)) {
			const widths = {};
			Object.entries(prefWidths).forEach(([field, rawWidth]) => {
				if (!INVENTORY_COLUMNS.some((column) => column.field === field)) return;
				const width = Number(rawWidth);
				if (Number.isFinite(width) && width >= 90 && width <= 1200) {
					widths[field] = Math.round(width);
				}
			});
			state.columnWidths = widths;
		}

		const prefPageSize = Number(response.data.inventory_page_size);
		if (Number.isFinite(prefPageSize) && prefPageSize >= 25 && prefPageSize <= 500) {
			state.perPage = prefPageSize;
		}

		const prefDensity = response.data.inventory_table_density;
		if (prefDensity === "compact" || prefDensity === "normal") {
			state.tableDensity = prefDensity;
		}
	}

	async function loadItems() {
		const params = new URLSearchParams();
		if (state.activeBlockId) params.set("bloque_id", state.activeBlockId);
		if (state.activeFloorId) params.set("piso_id", state.activeFloorId);
		if (state.activeAreaId) params.set("area_id", state.activeAreaId);
		if (state.search) params.set("search", state.search);
		params.set("order", state.order);
		params.set("page", String(state.page));
		params.set("per_page", String(state.perPage));
		const response = await api.get(`/api/inventario?${params.toString()}`);
		state.items = response.data || [];
		state.visibleItems = state.items;
		state.totalItems = Number(response.pagination?.total || state.items.length || 0);
		state.totalPages = Math.max(Number(response.pagination?.total_pages || 1), 1);
		state.page = Math.min(Math.max(Number(response.pagination?.page || state.page), 1), state.totalPages);
		state.perPage = Number(response.pagination?.per_page || state.perPage);
		renderRows();
		renderPaginationMeta();
	}

	async function refreshItemsTable() {
		await loadItems();
		renderRows();
	}

	async function saveCell(id, field, value, options = {}) {
		const payload = { [field]: value };
		if (options.forceDuplicate) payload.force_duplicate = true;
		await api.send(`/api/inventario/${id}`, "PATCH", payload);
	}

	async function createEmptyItem() {
		const payload = {
			cantidad: 1,
			estado: "",
			area_id: state.activeAreaId || null,
			ubicacion: composeSelectedLocation() || "",
		};
		await api.send("/api/inventario", "POST", payload);
		await loadItems();
	}

	async function removeItem(itemId) {
		await api.send(`/api/inventario/${itemId}`, "DELETE", {});
		await loadItems();
	}

	function renderDetail(item) {
		nodes.detailContainer.innerHTML = `
			<div class="row g-3">
				<div class="col-md-6"><strong>Cód Inv:</strong> ${item.cod_inventario || "-"}</div>
				<div class="col-md-6"><strong>Cód ESBYE:</strong> ${item.cod_esbye || "-"}</div>
				<div class="col-md-6"><strong>Cuenta:</strong> ${item.cuenta || "-"}</div>
				<div class="col-md-6"><strong>Cantidad:</strong> ${item.cantidad || "-"}</div>
				<div class="col-md-6"><strong>Estado:</strong> ${item.estado || "-"}</div>
				<div class="col-12"><strong>Ubicación:</strong> ${item.ubicacion || "-"}</div>
				<div class="col-md-6"><strong>Marca:</strong> ${item.marca || "-"}</div>
				<div class="col-md-6"><strong>Modelo:</strong> ${item.modelo || "-"}</div>
				<div class="col-md-6"><strong>Serie:</strong> ${item.serie || "-"}</div>
				<div class="col-md-6"><strong>Fecha:</strong> ${item.fecha_adquisicion || "-"}</div>
				<div class="col-md-6"><strong>Valor:</strong> ${formatValue("valor", item.valor) || "-"}</div>
				<div class="col-md-6"><strong>Usuario:</strong> ${item.usuario_final || "-"}</div>
				<div class="col-12"><strong>Descripción:</strong> ${item.descripcion || "-"}</div>
				<div class="col-12"><strong>Observación:</strong> ${item.observacion || "-"}</div>
				<div class="col-12"><strong>Descripción ESBYE:</strong> ${item.descripcion_esbye || "-"}</div>
				<div class="col-md-4"><strong>Marca ESBYE:</strong> ${item.marca_esbye || "-"}</div>
				<div class="col-md-4"><strong>Modelo ESBYE:</strong> ${item.modelo_esbye || "-"}</div>
				<div class="col-md-4"><strong>Serie ESBYE:</strong> ${item.serie_esbye || "-"}</div>
				<div class="col-md-4"><strong>Fecha ESBYE:</strong> ${item.fecha_adquisicion_esbye || "-"}</div>
				<div class="col-md-4"><strong>Valor ESBYE:</strong> ${formatValue("valor_esbye", item.valor_esbye) || "-"}</div>
				<div class="col-md-4"><strong>Ubicación ESBYE:</strong> ${item.ubicacion_esbye || "-"}</div>
				<div class="col-12"><strong>Observación ESBYE:</strong> ${item.observacion_esbye || "-"}</div>
			</div>
		`;
	}

	function startEdit(cell) {
		if (!cell.classList.contains("editable-cell")) return;
		if (cell.querySelector("input")) return;
		const oldValue = cell.textContent.trim();
		const input = document.createElement("input");
		input.type = "text";
		input.className = "form-control form-control-sm";
		input.value = oldValue;
		cell.innerHTML = "";
		cell.appendChild(input);
		input.focus();
		input.select();

		const commit = async () => {
			const id = cell.dataset.id;
			const field = cell.dataset.field;
			const newValue = input.value.trim();
			cell.textContent = newValue;
			try {
				await saveCell(id, field, newValue);
			} catch (error) {
				const duplicateList = error?.payload?.duplicates;
				if (error?.status === 409 && Array.isArray(duplicateList) && duplicateList.length) {
					const confirmed = await openDuplicateModal({
						duplicates: duplicateList,
						payload: {
							cod_inventario: field === "cod_inventario" ? newValue : undefined,
							cod_esbye: field === "cod_esbye" ? newValue : undefined,
						},
						mode: "update",
					});
					if (!confirmed) {
						cell.textContent = oldValue;
						return;
					}
					try {
						await saveCell(id, field, newValue, { forceDuplicate: true });
						return;
					} catch (forcedError) {
						cell.textContent = oldValue;
						notify(forcedError.message, true);
						return;
					}
				}
				cell.textContent = oldValue;
				notify(error.message, true);
			}
		};

		input.addEventListener("blur", commit);
		input.addEventListener("keydown", (event) => {
			if (event.key === "Enter") {
				event.preventDefault();
				input.blur();
			}
			if (event.key === "Escape") {
				cell.textContent = oldValue;
			}
		});
	}

	function reorderColumns(draggedField, targetField) {
		if (!draggedField || !targetField || draggedField === targetField) return false;
		const nextOrder = state.columnOrder.filter((field) => field !== draggedField);
		const targetIndex = nextOrder.indexOf(targetField);
		if (targetIndex < 0) return false;
		nextOrder.splice(targetIndex, 0, draggedField);
		state.columnOrder = nextOrder;
		return true;
	}

	async function persistColumnOrder() {
		try {
			await saveColumnPreference();
		} catch (error) {
			notify(error.message, true);
		}
	}

	function clearHeaderDropVisual() {
		nodes.tableHeadRow.querySelectorAll("th.drag-over").forEach((th) => th.classList.remove("drag-over"));
	}

	function bindHeaderInteractions() {
		let draggedField = null;
		nodes.tableHeadRow.querySelectorAll("th").forEach((th) => {
			const field = th.dataset.field;
			const resizeHandle = th.querySelector(".column-resize-handle");

			resizeHandle?.addEventListener("mousedown", (event) => {
				event.preventDefault();
				event.stopPropagation();
				const startX = event.clientX;
				const startWidth = th.getBoundingClientRect().width;

				const onMouseMove = (moveEvent) => {
					const nextWidth = Math.max(90, Math.min(1200, Math.round(startWidth + (moveEvent.clientX - startX))));
					state.columnWidths[field] = nextWidth;
					applyColumnWidthToField(field, nextWidth);
				};

				const onMouseUp = async () => {
					document.removeEventListener("mousemove", onMouseMove);
					document.removeEventListener("mouseup", onMouseUp);
					try {
						await saveColumnWidthsPreference();
					} catch (error) {
						notify(error.message, true);
					}
				};

				document.addEventListener("mousemove", onMouseMove);
				document.addEventListener("mouseup", onMouseUp);
			});

			th.addEventListener("dragstart", (event) => {
				draggedField = field;
				event.dataTransfer.effectAllowed = "move";
				event.dataTransfer.setData("text/plain", field);
			});
			th.addEventListener("dragover", (event) => {
				event.preventDefault();
				th.classList.add("drag-over");
			});
			th.addEventListener("dragleave", () => {
				th.classList.remove("drag-over");
			});
			th.addEventListener("drop", async (event) => {
				event.preventDefault();
				clearHeaderDropVisual();
				const fromField = event.dataTransfer.getData("text/plain") || draggedField;
				if (!reorderColumns(fromField, field)) return;
				renderTableHead();
				bindHeaderInteractions();
				renderRows();
				await persistColumnOrder();
			});

			th.addEventListener("mousedown", (event) => {
				if (event.button !== 0) return;
				if (event.target.closest(".column-resize-handle")) return;
				state.headerDrag.active = true;
				state.headerDrag.field = field;
				state.headerDrag.targetField = field;
				state.headerDrag.startX = event.clientX;
				state.headerDrag.startY = event.clientY;
				state.headerDrag.moved = false;
				document.body.classList.add("inventory-column-dragging");
			});
		});
	}

	document.addEventListener("mousemove", (event) => {
		if (!state.headerDrag.active) return;
		const deltaX = Math.abs(event.clientX - state.headerDrag.startX);
		const deltaY = Math.abs(event.clientY - state.headerDrag.startY);
		if (deltaX > 5 || deltaY > 5) state.headerDrag.moved = true;
		if (!state.headerDrag.moved) return;

		clearHeaderDropVisual();
		const hovered = document.elementFromPoint(event.clientX, event.clientY)?.closest("th[data-field]");
		if (!hovered) return;
		hovered.classList.add("drag-over");
		state.headerDrag.targetField = hovered.dataset.field;
	});

	document.addEventListener("mouseup", async () => {
		if (!state.headerDrag.active) return;
		document.body.classList.remove("inventory-column-dragging");
		clearHeaderDropVisual();

		const fromField = state.headerDrag.field;
		const toField = state.headerDrag.targetField;
		const moved = state.headerDrag.moved;
		state.headerDrag.active = false;

		if (!moved || !reorderColumns(fromField, toField)) return;
		renderTableHead();
		bindHeaderInteractions();
		renderRows();
		await persistColumnOrder();
	});

	async function viewItem(rowId) {
		const response = await api.get(`/api/inventario/${rowId}`);
		renderDetail(response.data);
		detailModal.show();
	}

	body.addEventListener("dblclick", async (event) => {
		const cell = event.target.closest("td");
		if (!cell) return;
		const field = cell.dataset.field;
		if (field === "item_numero") {
			try {
				await viewItem(cell.dataset.id);
			} catch (error) {
				notify(error.message, true);
			}
			return;
		}
		startEdit(cell);
	});

	body.addEventListener("contextmenu", (event) => {
		const row = event.target.closest("tr");
		if (!row) return;
		event.preventDefault();
		state.selectedRowId = row.dataset.id;
		nodes.contextMenu.style.left = `${event.pageX}px`;
		nodes.contextMenu.style.top = `${event.pageY}px`;
		nodes.contextMenu.classList.remove("d-none");
	});

	document.addEventListener("click", () => {
		nodes.contextMenu.classList.add("d-none");
	});

	nodes.contextMenu.addEventListener("click", async (event) => {
		const actionButton = event.target.closest("button[data-action]");
		if (!actionButton || !state.selectedRowId) return;
		const action = actionButton.dataset.action;
		try {
			if (action === "view") {
				await viewItem(state.selectedRowId);
			}
			if (action === "delete") {
				const confirmed = window.confirm("¿Seguro que deseas borrar este registro?");
				if (!confirmed) return;
				await removeItem(state.selectedRowId);
			}
		} catch (error) {
			notify(error.message, true);
		}
	});

	nodes.addInlineBtn.addEventListener("click", async () => {
		try {
			await createEmptyItem();
		} catch (error) {
			notify(error.message, true);
		}
	});

	nodes.filtroBloque.addEventListener("change", async () => {
		state.activeBlockId = nodes.filtroBloque.value;
		state.activeFloorId = "";
		state.activeAreaId = "";
		state.page = 1;
		updateFloorAndAreaFilters();
		await loadItems();
	});

	nodes.filtroPiso.addEventListener("change", async () => {
		state.activeFloorId = nodes.filtroPiso.value;
		state.activeAreaId = "";
		state.page = 1;
		updateFloorAndAreaFilters();
		await loadItems();
	});

	nodes.filtroArea.addEventListener("change", async () => {
		state.activeAreaId = nodes.filtroArea.value;
		state.page = 1;
		await loadItems();
	});

	nodes.filtroOrder.addEventListener("change", async () => {
		state.order = nodes.filtroOrder.value;
		state.page = 1;
		await loadItems();
	});

	nodes.filtroSearch.addEventListener(
		"input",
		debounce(async () => {
			state.search = nodes.filtroSearch.value.trim();
			state.page = 1;
			await loadItems();
		})
	);

	nodes.pagePrev?.addEventListener("click", async () => {
		if (state.page <= 1) return;
		state.page -= 1;
		await loadItems();
	});

	nodes.pageNext?.addEventListener("click", async () => {
		if (state.page >= state.totalPages) return;
		state.page += 1;
		await loadItems();
	});

	nodes.pageSize?.addEventListener("change", async () => {
		const nextSize = Number(nodes.pageSize.value);
		if (!Number.isFinite(nextSize) || nextSize < 25 || nextSize > 500) return;
		state.perPage = nextSize;
		state.page = 1;
		try {
			await savePageSizePreference();
		} catch (error) {
			notify(error.message, true);
		}
		await loadItems();
	});

	nodes.table.addEventListener("paste", async (event) => {
		const text = event.clipboardData.getData("text/plain");
		if (!text || !text.includes("\t")) return;
		event.preventDefault();
		const rows = parseExcelText(text);
		if (!rows.length) return;
		try {
			await api.send("/api/inventario/pegar", "POST", {
				rows,
				area_id: state.activeAreaId || null,
			});
			await loadItems();
			notify("Pegado desde Excel aplicado correctamente.");
		} catch (error) {
			notify(error.message, true);
		}
	});

if (nodes.excelSingleRow) {
			nodes.excelSingleRow.addEventListener("paste", async (event) => {
				const text = event.clipboardData.getData("text/plain");
				if (!text.trim()) return;
				event.preventDefault();
				await loadModalParams();
				const rows = parseExcelText(text);
				if (!rows.length) return;
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
					"descripcion_esbye",
					"marca_esbye",
					"modelo_esbye",
					"serie_esbye",
					"fecha_adquisicion_esbye",
					"valor_esbye",
					"ubicacion_esbye",
					"observacion_esbye",
				];
				const first = rows[0];
				fields.forEach((field, idx) => {
					const input = document.querySelector(`#form-agregar-item [data-field='${field}']`);
					if (input && first[idx] !== undefined) {
						assignPastedValue(input, first[idx]);
					}
				});
				if (nodes.modalUbicacion && !nodes.modalUbicacion.value.trim()) {
					syncModalLocationFromSelection();
				}
			});
		}

	function getCurrentFilterParams() {
		const params = new URLSearchParams();
		if (state.activeBlockId) params.set("bloque_id", state.activeBlockId);
		if (state.activeFloorId) params.set("piso_id", state.activeFloorId);
		if (state.activeAreaId) params.set("area_id", state.activeAreaId);
		if (state.search) params.set("search", state.search);
		params.set("order", state.order);
		return params;
	}

	nodes.exportExcelBtn?.addEventListener("click", () => {
		const params = getCurrentFilterParams();
		window.location.href = `/api/inventario/export?${params.toString()}`;
	});

	nodes.toggleDensityBtn?.addEventListener("click", () => {
		state.tableDensity = state.tableDensity === "compact" ? "normal" : "compact";
		applyDensityMode();
		saveDensityPreference().catch((error) => notify(error.message, true));
	});

	nodes.modalAddButton.addEventListener("click", async () => {
		const payload = {};
			const inputs = document.querySelectorAll("#form-agregar-item [data-field]");
			inputs.forEach((input) => {
				let value = input.value.trim();
				if (input.dataset.field === "cantidad") {
					const quantity = Number(value);
					value = Number.isInteger(quantity) && quantity > 0 ? quantity : null;
				}
				if (input.dataset.field === "valor") {
					value = parseDecimalWithComma(value);
				}
				if (input.dataset.field === "valor_esbye") {
					value = parseDecimalWithComma(value);
				}
				if (input.dataset.field === "area_id") {
					value = value === "" ? null : parseInt(value, 10);
				}
				if (value !== "" && value !== null) {
					payload[input.dataset.field] = value;
				}
			});
			if (!payload.cantidad) {
				notify("La cantidad debe ser un entero mayor que 0.", true);
				return;
			}
			if (!payload.area_id && state.activeAreaId) {
				payload.area_id = parseInt(state.activeAreaId, 10);
			}
			if (!payload.ubicacion) {
				payload.ubicacion = composeSelectedLocation() || "";
			}
			try {
				await api.send("/api/inventario", "POST", payload);
				addModal.hide();
				inputs.forEach((input) => {
					input.value = input.dataset.field === "cantidad" ? "1" : "";
				});
				if (nodes.excelSingleRow) nodes.excelSingleRow.value = "";
				await refreshItemsTable();
				notify("Ítem guardado correctamente.");
			} catch (error) {
				const duplicateList = error?.payload?.duplicates;
				if (error?.status === 409 && Array.isArray(duplicateList) && duplicateList.length) {
					const confirmed = await openDuplicateModal({
						duplicates: duplicateList,
						payload,
						mode: "create",
					});
					if (!confirmed) {
						return;
					}
					try {
						await api.send("/api/inventario", "POST", {
							...payload,
							force_duplicate: true,
						});
						addModal.hide();
						inputs.forEach((input) => {
							input.value = input.dataset.field === "cantidad" ? "1" : "";
						});
						if (nodes.excelSingleRow) nodes.excelSingleRow.value = "";
						await refreshItemsTable();
						notify("Ítem guardado correctamente (código repetido autorizado).");
					} catch (forceError) {
						notify(forceError.message, true);
					}
					return;
				}
				notify(error.message, true);
			}
	});

	async function loadModalParams() {
		try {
			const [estadosRes, cuentasRes, adminsRes] = await Promise.all([
				api.get("/api/parametros/estados"),
				api.get("/api/parametros/cuentas"),
				api.get("/api/administradores"),
			]);

			const estadoSelect = document.querySelector("#form-agregar-item [data-field='estado']");
			const cuentaSelect = document.querySelector("#form-agregar-item [data-field='cuenta']");
			const usuarioFinalSelect = document.querySelector("#form-agregar-item [data-field='usuario_final']");

			if (estadoSelect) {
				estadoSelect.innerHTML = '<option value="">-- Seleccionar estado --</option>';
				(estadosRes.data || []).forEach((estado) => {
					const opt = document.createElement("option");
					opt.value = estado.nombre;
					opt.textContent = estado.nombre;
					estadoSelect.appendChild(opt);
				});
			}

			if (cuentaSelect) {
				cuentaSelect.innerHTML = '<option value="">-- Seleccionar cuenta --</option>';
				(cuentasRes.data || []).forEach((cuenta) => {
					const opt = document.createElement("option");
					opt.value = cuenta.nombre;
					opt.textContent = cuenta.codigo ? `${cuenta.codigo} - ${cuenta.nombre}` : cuenta.nombre;
					cuentaSelect.appendChild(opt);
				});
			}

			if (usuarioFinalSelect) {
				const currentValue = usuarioFinalSelect.value;
				usuarioFinalSelect.innerHTML = '<option value="">-- Seleccionar personal --</option>';
				(adminsRes.data || []).forEach((admin) => {
					const opt = document.createElement("option");
					opt.value = admin.nombre;
					opt.textContent = admin.nombre;
					usuarioFinalSelect.appendChild(opt);
				});
				if (currentValue) usuarioFinalSelect.value = currentValue;
			}

			renderAreaModalSelect();
		} catch (error) {
			console.error("Error loading modal params:", error);
		}
	}

	addModalElement.addEventListener("shown.bs.modal", async () => {
		await loadModalParams();
		if (state.activeAreaId) {
			nodes.modalAreaSelect.value = String(state.activeAreaId);
		}
		syncModalLocationFromSelection();
	});

	nodes.modalAreaSelect?.addEventListener("change", () => {
		const selectedAreaId = nodes.modalAreaSelect.value;
		if (!selectedAreaId) {
			syncModalLocationFromSelection();
			return;
		}
		const foundArea = flattenAreas().find((entry) => String(entry.id) === String(selectedAreaId));
		if (!foundArea || !nodes.modalUbicacion) {
			syncModalLocationFromSelection();
			return;
		}
		nodes.modalUbicacion.value = `${foundArea.bloque_nombre} / ${foundArea.piso_nombre} / ${foundArea.area_nombre}`;
	});

	try {
		await loadStructure();
		await loadPreferences();
		applyDensityMode();
		renderTableHead();
		bindHeaderInteractions();
		renderPaginationMeta();
		await refreshItemsTable();
	} catch (error) {
		notify(error.message, true);
	}
}

document.addEventListener("DOMContentLoaded", async () => {
	await initInventoryPage();
});

