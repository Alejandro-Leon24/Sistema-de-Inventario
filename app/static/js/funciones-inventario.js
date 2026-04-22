const INVENTORY_COLUMNS = [
	{ field: "item_numero", label: "ITEM", editable: false },
	{ field: "cod_inventario", label: "COD INV.", editable: true },
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
	{ field: "justificacion", label: "JUSTIFICACIÓN", editable: true },
	{ field: "procedencia", label: "PROCEDENCIA", editable: true },
	{ field: "descripcion_esbye", label: "DESCRIPCIÓN ESBYE", editable: true },
	{ field: "marca_esbye", label: "MARCA ESBYE", editable: true },
	{ field: "modelo_esbye", label: "MODELO ESBYE", editable: true },
	{ field: "serie_esbye", label: "SERIE ESBYE", editable: true },
	{ field: "fecha_adquisicion_esbye", label: "FECHA ESBYE", editable: true },
	{ field: "valor_esbye", label: "VALOR ESBYE", editable: true },
	{ field: "ubicacion_esbye", label: "UBICACIÓN ESBYE", editable: true },
	{ field: "observacion_esbye", label: "OBSERVACIÓN ESBYE", editable: true },
];

const INLINE_ADD_OPTION_VALUE = "__add_new_option__";
const INLINE_SELECT_FIELDS = new Set(["estado", "cuenta", "usuario_final", "ubicacion"]);
const INLINE_INPUT_CONFIG = {
	cantidad: { type: "number", min: "1", step: "1" },
	fecha_adquisicion: { type: "date" },
	fecha_adquisicion_esbye: { type: "date" },
};

const api = window.api;

function debounce(callback, delay = 350) {
	let timer;
	return (...args) => {
		clearTimeout(timer);
		timer = setTimeout(() => callback(...args), delay);
	};
}

function parseExcelText(text) {
	const normalizedText = String(text || "")
		.replace(/\r\n/g, "\n")
		.replace(/\r/g, "\n");

	return normalizedText
		.split("\n")
		.filter((line) => line.replace(/\t/g, "").trim().length > 0)
		.map((line) => line.split("\t").map((cell) => String(cell || "").trim()));
}

	const helper = window.appHelpers;

	function formatValue(field, value) {
		if (value === null || value === undefined) return "";
		if ((field === "valor" || field === "valor_esbye") && value !== "") {
			const num = typeof value === "number" ? value : helper.parseDecimalWithComma(value);
			if (num === null) return value;
			return num.toLocaleString("es-EC", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
		}
		return value;
	}

function normalizeCodeToPlaceholder(value) {
	const text = String(value || "").trim();
	const compact = text.toLowerCase().replace(/[^a-z0-9]/g, "");
	if (!compact || compact === "sc" || compact === "sincodigo" || compact === "sincod") {
		return "S/C";
	}
	return text;
}

const MODAL_PASTE_FIELDS = [
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

function scoreModalPastedMapping(mapped = {}) {
	let score = 0;
	if (String(mapped.cod_inventario || "").trim()) score += 3;
	if (String(mapped.descripcion || "").trim()) score += 4;
	if (String(mapped.ubicacion || "").trim()) score += 3;
	if (String(mapped.usuario_final || "").trim()) score += 2;
	if (String(mapped.estado || "").trim()) score += 1;
	if (String(mapped.cantidad || "").trim().match(/^\d+$/)) score += 2;
	if (String(mapped.fecha_adquisicion || "").trim()) score += 1;
	if (String(mapped.valor || "").trim()) score += 1;
	score += Object.keys(mapped).length * 0.25;
	return score;
}

function mapPastedRowBestEffortForModal(rawRow) {
	const cells = Array.isArray(rawRow) ? rawRow : [];
	if (!cells.length) return {};

	const candidateOrders = [
		MODAL_PASTE_FIELDS,
		["item_numero", ...MODAL_PASTE_FIELDS],
	];

	let bestMapped = {};
	let bestScore = Number.NEGATIVE_INFINITY;
	let bestOffset = 0;

	candidateOrders.forEach((order) => {
		const maxOffset = Math.min(8, Math.max(cells.length - 1, 0));
		for (let offset = 0; offset <= maxOffset; offset += 1) {
			const mapped = {};
			for (let idx = 0; idx < order.length; idx += 1) {
				const field = order[idx];
				if (field === "item_numero") continue;
				const srcIdx = idx + offset;
				if (srcIdx >= cells.length) break;
				const value = String(cells[srcIdx] ?? "").trim();
				if (!value) continue;
				mapped[field] = value;
			}

			const score = scoreModalPastedMapping(mapped);
			if (score > bestScore || (score === bestScore && offset < bestOffset)) {
				bestScore = score;
				bestMapped = mapped;
				bestOffset = offset;
			}
		}
	});

	return bestMapped;
}

const escapeHtmlText = window.appHelpers.escapeHtmlText;

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
			`${index + 1}. ítem #${item.item_numero || "-"} ${matchLabel}`,
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
		order: "desc",
		search: "",
		page: 1,
		perPage: 50,
		totalItems: 0,
		totalPages: 1,
		editingItemId: null,
		tableDensity: "compact",
		columnOrder: INVENTORY_COLUMNS.map((column) => column.field),
		columnWidths: {},
		modalOptions: {
			estados: [],
			cuentas: [],
			administradores: [],
		},
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
		btnAbrirModalAgregar: document.getElementById("btn-abrir-modal-agregar"),
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
		detailContainer: document.getElementById("detalle-contenido"),
		contextMenu: document.getElementById("context-menu"),
		modalAreaSelect: document.getElementById("modal-area-select"),
		modalCuentaSelect: document.getElementById("modal-cuenta-select"),
		modalUsuarioFinalSelect: document.getElementById("modal-usuario-final-select"),
		modalUbicacion: document.getElementById("modal-ubicacion"),
		modalUbicacionEsbye: document.querySelector("#form-agregar-item [data-field='ubicacion_esbye']"),
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
		quickAddModalEl: document.getElementById("modalAgregarParametroRapido"),
		quickAddTitle: document.getElementById("modalAgregarParametroRapidoLabel"),
		quickAddNameLabel: document.getElementById("quick-add-name-label"),
		quickAddNameInput: document.getElementById("quick-add-name"),
		quickAddDescriptionWrap: document.getElementById("quick-add-description-wrapper"),
		quickAddDescriptionLabel: document.getElementById("quick-add-description-label"),
		quickAddDescriptionInput: document.getElementById("quick-add-description"),
		quickAddSaveButton: document.getElementById("quick-add-save"),
	};

	const detailModal = new bootstrap.Modal(document.getElementById("modalDetalle"));
	const addModalElement = document.getElementById("modalAgregarItem");
	const addModal = new bootstrap.Modal(addModalElement);
	const addModalTitle = addModalElement?.querySelector(".modal-title");
	const duplicateModal = nodes.duplicateModalEl ? new bootstrap.Modal(nodes.duplicateModalEl) : null;
	const quickAddModal = nodes.quickAddModalEl ? new bootstrap.Modal(nodes.quickAddModalEl) : null;
	let itemsLoadAbortController = null;
	let itemsLoadRequestId = 0;
	const quickAddState = {
		mode: null,
		resolver: null,
	};
	const duplicateModalController = typeof window.createDuplicateInventoryModal === "function"
		? window.createDuplicateInventoryModal({
			duplicateModal,
			nodes,
			escapeHtmlText,
			buildDuplicateWarningMessage,
		})
		: null;

	const openDuplicateModal = duplicateModalController?.openDuplicateModal
		? duplicateModalController.openDuplicateModal
		: ({ duplicates = [], payload = {} } = {}) => Promise.resolve(window.confirm(buildDuplicateWarningMessage(duplicates, payload)));

	function getColumn(field) {
		return INVENTORY_COLUMNS.find((column) => column.field === field);
	}

	nodes.quickAddSaveButton?.addEventListener("click", async () => {
		const mode = quickAddState.mode;
		const config = resolveQuickAddConfig(mode);
		if (!config) return;

		const name = String(nodes.quickAddNameInput?.value || "").trim();
		const description = String(nodes.quickAddDescriptionInput?.value || "").trim();
		if (!name) {
			notify("El nombre es obligatorio.", true);
			return;
		}

		try {
			const createdValue = await config.save(name, description);
			quickAddModal?.hide();
			if (quickAddState.resolver) quickAddState.resolver(createdValue);
			quickAddState.resolver = null;
		} catch (error) {
			notify(error.message, true);
		}
	});

	nodes.quickAddModalEl?.addEventListener("hidden.bs.modal", () => {
		if (quickAddState.resolver) {
			quickAddState.resolver(null);
			quickAddState.resolver = null;
		}
		quickAddState.mode = null;
	});

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

	const tableController = typeof window.createInventarioTablaController === "function"
		? window.createInventarioTablaController({
			state,
			nodes,
			body,
			getOrderedColumns,
			formatValue,
			escapeHtmlText,
			applyColumnWidthToField,
		})
		: null;

	function renderTableHead() {
		tableController?.renderTableHead();
	}

	function renderRows() {
		tableController?.renderRows();
	}

	function renderPaginationMeta() {
		tableController?.renderPaginationMeta();
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

	function getModalSelectConfig(field) {
		if (field === "estado") {
			return {
				placeholder: "-- Seleccionar estado --",
				items: state.modalOptions.estados,
				optionValue: (item) => item.nombre,
				optionLabel: (item) => item.nombre,
				quickMode: "estados",
				quickLabel: "+ Agregar Estado",
			};
		}
		if (field === "cuenta") {
			return {
				placeholder: "-- Seleccionar cuenta --",
				items: state.modalOptions.cuentas,
				optionValue: (item) => item.nombre,
				optionLabel: (item) => (item.codigo ? `${item.codigo} - ${item.nombre}` : item.nombre),
				quickMode: "cuentas",
				quickLabel: "+ Agregar Cuenta",
			};
		}
		if (field === "usuario_final") {
			return {
				placeholder: "-- Seleccionar personal --",
				items: state.modalOptions.administradores,
				optionValue: (item) => item.nombre,
				optionLabel: (item) => item.nombre,
				quickMode: "administrador",
				quickLabel: "+ Agregar Personal",
			};
		}
		if (field === "ubicacion") {
			const areas = flattenAreas();
			return {
				placeholder: "-- Seleccionar ubicación --",
				items: areas,
				optionValue: (item) => item.nombre,
				optionLabel: (item) => item.nombre,
				strictSelection: true,
				quickMode: null,
				quickLabel: "",
			};
		}
		return null;
	}

	function renderSelectWithQuickAdd(select, config, selectedValue = "") {
		if (!select || !config) return;
		const normalizeSelectComparable = (value) =>
			normalizeText(value)
				.replace(/[^a-z0-9\s]/g, " ")
				.replace(/\s+/g, " ")
				.trim();

		select.innerHTML = `<option value="">${config.placeholder}</option>`;
		(config.items || []).forEach((item) => {
			const option = document.createElement("option");
			option.value = config.optionValue(item);
			option.textContent = config.optionLabel(item);
			select.appendChild(option);
		});
		if (config.quickMode && config.quickLabel) {
			const addOption = document.createElement("option");
			addOption.value = INLINE_ADD_OPTION_VALUE;
			addOption.textContent = config.quickLabel;
			select.appendChild(addOption);
		}
		const normalizedSelected = normalizeSelectComparable(selectedValue);
		let resolved = "";
		if (normalizedSelected) {
			const options = Array.from(select.options || []).filter((opt) => opt.value !== INLINE_ADD_OPTION_VALUE && opt.value !== "");
			let match = options.find(
				(opt) =>
					normalizeSelectComparable(opt.value) === normalizedSelected
					|| normalizeSelectComparable(opt.textContent) === normalizedSelected
			);
			if (config.strictSelection) {
				select.value = match ? match.value : "";
				return;
			}
			if (!match) {
				match = options.find(
					(opt) =>
						normalizeSelectComparable(opt.textContent).includes(normalizedSelected)
						|| normalizedSelected.includes(normalizeSelectComparable(opt.value))
				);
			}
			if (!match) {
				const selectedTokens = normalizedSelected.split(/\s+/).filter(Boolean);
				let best = null;
				let bestScore = 0;
				options.forEach((opt) => {
					const optNorm = normalizeSelectComparable(opt.value || opt.textContent);
					if (!optNorm) return;
					const optTokens = optNorm.split(/\s+/).filter(Boolean);
					const overlap = selectedTokens.filter((token) => optTokens.some((item) => item.includes(token) || token.includes(item))).length;
					const score = selectedTokens.length ? overlap / selectedTokens.length : 0;
					if (score > bestScore) {
						bestScore = score;
						best = opt;
					}
				});
				if (best && bestScore >= 0.6) match = best;
			}
			if (match) resolved = match.value;
		}
		select.value = resolved || "";
	}

	function resolveAreaEntryFromLocationValue(rawValue) {
		const value = String(rawValue || "").trim();
		if (!value) return null;

		const normalized = normalizeText(value);
		const areas = flattenAreas();

		let exact = areas.find((entry) => normalizeText(entry.nombre) === normalized);
		if (exact) return exact;

		exact = areas.find((entry) => normalizeText(entry.area_nombre) === normalized);
		if (exact) return exact;

		return resolveAreaFromPastedText(value);
	}

	async function resolveAreaEntryFromLocationValueBackend(rawValue) {
		const value = String(rawValue || "").trim();
		if (!value) return null;

		try {
			const query = new URLSearchParams({ texto: value });
			const response = await fetch(`/api/inventario/resolver-ubicacion?${query.toString()}`);
			const payload = await response.json();
			const match = payload?.match;
			if (!response.ok || !payload?.success || !match || !Number.isFinite(Number(match.area_id))) {
				return null;
			}

			const areaEntry = flattenAreas().find((entry) => String(entry.id) === String(match.area_id));
			if (areaEntry) return areaEntry;

			return {
				id: Number(match.area_id),
				nombre: String(match.display || "").trim(),
				bloque_nombre: "",
				piso_nombre: "",
				area_nombre: String(match.display || "").trim(),
			};
		} catch (_error) {
			return null;
		}
	}

	function refreshAddItemSelects() {
		const estadoSelect = document.querySelector("#form-agregar-item [data-field='estado']");
		const cuentaSelect = document.querySelector("#form-agregar-item [data-field='cuenta']");
		const usuarioFinalSelect = document.querySelector("#form-agregar-item [data-field='usuario_final']");

		const estadoCurrent = estadoSelect?.value || "";
		renderSelectWithQuickAdd(estadoSelect, getModalSelectConfig("estado"), estadoCurrent);

		const cuentaCurrent = cuentaSelect?.value || "";
		renderSelectWithQuickAdd(cuentaSelect, getModalSelectConfig("cuenta"), cuentaCurrent);

		const usuarioCurrent = usuarioFinalSelect?.value || "";
		renderSelectWithQuickAdd(usuarioFinalSelect, getModalSelectConfig("usuario_final"), usuarioCurrent);
	}

	function updateLocalItemValue(id, field, value) {
		const item = state.items.find((entry) => String(entry.id) === String(id));
		if (item) item[field] = value;
	}

	function getRowFieldValue(id, field) {
		const item = state.items.find((entry) => String(entry.id) === String(id));
		return item ? item[field] : "";
	}

	function resolveQuickAddConfig(mode) {
		const modalConfigs = {
			estados: {
				title: "Agregar Estado",
				nameLabel: "Nombre",
				descriptionLabel: "Descripcion (opcional)",
				descriptionVisible: true,
				save: async (name, description) => {
					await api.send("/api/parametros/estados", "POST", { nombre: name, descripcion: description || null });
					return name;
				},
			},
			cuentas: {
				title: "Agregar Cuenta",
				nameLabel: "Nombre",
				descriptionLabel: "Descripcion (opcional)",
				descriptionVisible: true,
				save: async (name, description) => {
					await api.send("/api/parametros/cuentas", "POST", { nombre: name, descripcion: description || null });
					return name;
				},
			},
			administrador: {
				title: "Agregar Personal",
				nameLabel: "Nombre completo",
				descriptionLabel: "Cargo (opcional)",
				descriptionVisible: true,
				save: async (name, cargo) => {
					await api.send("/api/administradores", "POST", {
						nombre: name,
						cargo: cargo || null,
						facultad: null,
						titulo_academico: null,
						email: null,
						telefono: null,
					});
					return name;
				},
			},
		};
		return modalConfigs[mode] || null;
	}

	function openQuickAddModal(mode) {
		const config = resolveQuickAddConfig(mode);
		if (!config || !quickAddModal || !nodes.quickAddSaveButton) return Promise.resolve(null);

		quickAddState.mode = mode;
		nodes.quickAddTitle.textContent = config.title;
		nodes.quickAddNameLabel.textContent = config.nameLabel;
		nodes.quickAddDescriptionLabel.textContent = config.descriptionLabel;
		nodes.quickAddDescriptionWrap.classList.toggle("d-none", !config.descriptionVisible);
		nodes.quickAddNameInput.value = "";
		nodes.quickAddDescriptionInput.value = "";

		return new Promise((resolve) => {
			quickAddState.resolver = resolve;
			quickAddModal.show();
			setTimeout(() => nodes.quickAddNameInput.focus(), 120);
		});
	}

	async function ensureModalOptions() {
		if (!state.modalOptions.estados.length && !state.modalOptions.cuentas.length && !state.modalOptions.administradores.length) {
			await loadModalParams();
		}
	}

	function attachQuickAddToSelect(select, field) {
		if (!select || select.dataset.quickAddBound === "1") return;
		select.dataset.quickAddBound = "1";
		select.addEventListener("change", async () => {
			if (select.value !== INLINE_ADD_OPTION_VALUE) return;
			const selectConfig = getModalSelectConfig(field);
			if (!selectConfig) return;
			const createdValue = await openQuickAddModal(selectConfig.quickMode);
			if (!createdValue) {
				select.value = "";
				return;
			}
			await loadModalParams();
			select.value = createdValue;
		});
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

	function normalizePersonForSelectMatch(value) {
		const normalized = normalizeText(value).replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
		if (!normalized) return "";
		const prefixes = new Set([
			"ing", "ingeniero", "ingeniera", "dr", "dra", "doctor", "doctora",
			"lic", "licenciado", "licenciada", "abg", "abogada", "abogado",
			"arq", "arquitecto", "arquitecta", "tec", "tecnico", "tecnica",
			"sr", "sra", "srta", "msc", "mg", "mgs", "mgtr", "mtr", "mts", "mtro", "mt",
			"prof", "profa", "tlgo", "tlga", "ts", "phd",
		]);
		const tokens = normalized.split(" ").filter(Boolean);
		while (tokens.length && prefixes.has(tokens[0])) {
			tokens.shift();
		}
		return tokens.join(" ");
	}

	function resolveSelectOptionBestMatch(options, field, value) {
		const rawValue = String(value || "").trim();
		if (!rawValue) return null;
		const normalizedValue = normalizeText(rawValue);
		const personValue = field === "usuario_final" ? normalizePersonForSelectMatch(rawValue) : "";

		let match = options.find((opt) => normalizeText(opt.value) === normalizedValue);
		if (match) return match;

		match = options.find((opt) => normalizeText(opt.textContent) === normalizedValue);
		if (match) return match;

		if (personValue) {
			match = options.find((opt) => normalizePersonForSelectMatch(opt.textContent) === personValue);
			if (match) return match;
		}

		match = options.find((opt) => normalizeText(opt.textContent).includes(normalizedValue));
		if (match) return match;

		match = options.find((opt) => normalizedValue.includes(normalizeText(opt.value)));
		if (match) return match;

		// Match difuso por solapamiento de tokens para capturar variantes de nombre.
		const targetTokens = (personValue || normalizedValue).split(/\s+/).filter(Boolean);
		let best = null;
		let bestScore = 0;
		options.forEach((opt) => {
			const base = field === "usuario_final"
				? normalizePersonForSelectMatch(opt.textContent)
				: normalizeText(opt.textContent).replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
			if (!base) return;
			const optionTokens = base.split(/\s+/).filter(Boolean);
			if (!targetTokens.length || !optionTokens.length) return;
			const overlap = targetTokens.filter((token) => optionTokens.some((ot) => ot.includes(token) || token.includes(ot))).length;
			const score = overlap / Math.max(targetTokens.length, optionTokens.length);
			if (score > bestScore) {
				bestScore = score;
				best = opt;
			}
		});

		return bestScore >= 0.5 ? best : null;
	}

	function resolveAreaFromPastedText(rawText) {
		const input = normalizeText(rawText);
		if (!input) return null;

		const targetTokens = input.split(/[\s/\-]+/).filter(Boolean);
		const extractCompactAulaCode = (text) => {
			const normalized = normalizeText(text);
			const match = normalized.match(/(\d+[a-z]\s*-?\s*\d+)/);
			return match ? match[1].replace(/\s+/g, "") : "";
		};
		const extractFloorHint = (text) => {
			const normalized = normalizeText(text);
			if (normalized.includes("planta baja")) return "planta baja";
			if (normalized.includes("primer") && normalized.includes("piso")) return "primer piso";
			if (normalized.includes("segundo") && normalized.includes("piso")) return "segundo piso";
			if (normalized.includes("tercer") && normalized.includes("piso")) return "tercer piso";
			return "";
		};

		const targetAulaCode = extractCompactAulaCode(input);
		const floorHint = extractFloorHint(input);
		const targetKind = input.includes("pasillo") ? "pasillo" : (input.includes("aula") || targetAulaCode ? "aula" : "");
		let explicitBlockLetter = "";
		const explicitBlockMatch = input.match(/\bbloque\s+([a-z])\b/);
		if (explicitBlockMatch) {
			explicitBlockLetter = explicitBlockMatch[1];
		}

		const ranked = [];
		let bestPasillo = null;
		let bestPasilloScore = -Infinity;

		flattenAreas().forEach((entry) => {
			const areaName = normalizeText(entry.area_nombre);
			const floorName = normalizeText(entry.piso_nombre);
			const blockName = normalizeText(entry.bloque_nombre);
			const fullName = `${blockName} ${floorName} ${areaName}`.trim();

			let score = 0;
			if (input === areaName || input === fullName) score += 30;
			if (fullName.includes(input)) score += 10;

			const areaCode = extractCompactAulaCode(areaName);
			if (targetAulaCode && areaCode && targetAulaCode === areaCode) score += 35;

			const areaTokens = fullName.split(/[\s/\-]+/).filter(Boolean);
			let tokenHits = 0;
			targetTokens.forEach((token) => {
				if (token.length < 3) return;
				if (areaTokens.some((candidate) => token.includes(candidate) || candidate.includes(token))) {
					tokenHits += 1;
				}
			});
			score += Math.min(tokenHits * 2, 16);

			if (floorHint && floorName.includes(floorHint)) score += 8;

			const blockLetterMatch = (targetAulaCode || "").match(/\d+([a-z])/);
			if (blockLetterMatch && blockName.includes(`bloque ${blockLetterMatch[1]}`)) {
				score += 10;
			}

			if (explicitBlockLetter) {
				if (blockName.includes(`bloque ${explicitBlockLetter}`)) {
					score += 12;
				} else {
					score -= 16;
				}
			}

			if (targetKind === "pasillo") {
				if (areaName.includes("pasillo")) {
					score += 22;
				} else {
					score -= 20;
				}
			} else if (targetKind === "aula") {
				if (areaName.includes("aula") || areaCode) {
					score += 8;
				}
			}

			ranked.push({
				entry,
				score,
				areaCode,
				floorName,
				areaName,
			});

			if (targetKind === "pasillo" && areaName.includes("pasillo")) {
				let pasilloScore = score;
				if (floorHint && floorName.includes(floorHint)) {
					pasilloScore += 5;
				}
				if (pasilloScore > bestPasilloScore) {
					bestPasilloScore = pasilloScore;
					bestPasillo = entry;
				}
			}
		});

		if (!ranked.length) return null;

		if (targetKind === "pasillo") {
			return bestPasillo;
		}

		ranked.sort((a, b) => b.score - a.score);
		const best = ranked[0];
		const second = ranked[1];
		const margin = second ? best.score - second.score : 999;
		const hasExactAulaCode = Boolean(targetAulaCode && best.areaCode && targetAulaCode === best.areaCode);

		if (best.score < 10) return null;
		if (!hasExactAulaCode && margin <= 1) return null;

		return best.entry;
	}

	async function resolveAreaFromPastedTextDefault(rawText) {
		const text = String(rawText || "").trim();
		if (!text) return null;

		try {
			const query = new URLSearchParams({ texto: text });
			const response = await fetch(`/api/inventario/resolver-ubicacion?${query.toString()}`);
			const payload = await response.json();
			const match = payload?.match;
			if (!response.ok || !payload?.success) return null;
			if (!match || !Number.isFinite(Number(match.area_id))) {
				return null;
			}

			const areaEntry = flattenAreas().find((entry) => String(entry.id) === String(match.area_id));
			if (areaEntry) return areaEntry;

			if (String(match.display || "").trim()) {
				return {
					id: Number(match.area_id),
					nombre: String(match.display).trim(),
					bloque_nombre: "",
					piso_nombre: "",
					area_nombre: String(match.display).trim(),
				};
			}

			return null;
		} catch (_error) {
			return null;
		}
	}

	function offerOpenSettingsForMissingAreas(missingAreas) {
		const unique = Array.from(new Set((missingAreas || []).map((x) => String(x || "").trim()).filter(Boolean)));
		if (!unique.length) return;
		const preview = unique.slice(0, 5).join(", ");
		const suffix = unique.length > 5 ? "..." : "";
		const goToSettings = window.confirm(
			`No se encontraron estas ubicaciones/areas: ${preview}${suffix}. ¿Desea ir a Configuración para registrarlas ahora?`
		);
		if (goToSettings) {
			window.open("/ajustes", "_blank", "noopener,noreferrer");
		}
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
		const field = String(input.dataset?.field || "");

		if (field === "valor" || field === "valor_esbye") {
			let normalizedMoney = value.replace(/[^\d,.-]/g, "");
			const lastComma = normalizedMoney.lastIndexOf(",");
			const lastDot = normalizedMoney.lastIndexOf(".");

			if (lastComma >= 0 && lastDot >= 0) {
				if (lastComma > lastDot) {
					normalizedMoney = normalizedMoney.replace(/\./g, "");
				} else {
					normalizedMoney = normalizedMoney.replace(/,/g, "");
					const parts = normalizedMoney.split(".");
					const decimal = parts.pop();
					normalizedMoney = `${parts.join("")}${decimal !== undefined ? `,${decimal}` : ""}`;
				}
			} else if (lastDot >= 0) {
				const parts = normalizedMoney.split(".");
				const decimal = parts.pop();
				normalizedMoney = `${parts.join("")}${decimal !== undefined ? `,${decimal}` : ""}`;
			}

			normalizedMoney = normalizedMoney.replace(/-/g, "");
			const firstComma = normalizedMoney.indexOf(",");
			if (firstComma !== -1) {
				normalizedMoney = normalizedMoney.slice(0, firstComma + 1) + normalizedMoney.slice(firstComma + 1).replace(/,/g, "");
			}

			input.value = normalizedMoney;
			return;
		}

		if (input.tagName === "SELECT") {
			const options = Array.from(input.options || []);
			if (!value) {
				input.value = "";
				return;
			}

			const match = resolveSelectOptionBestMatch(options, field, value);

			if (match) {
				input.value = match.value;
			} else if (field === "usuario_final") {
				const dynamicOption = document.createElement("option");
				dynamicOption.value = value;
				dynamicOption.textContent = value;
				input.appendChild(dynamicOption);
				input.value = value;
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
		state.structure = await window.appHelpers.loadStructure(api, { sortNatural: true });
		fillSelect(nodes.filtroBloque, state.structure, "Todos");
		updateFloorAndAreaFilters();
		renderAreaModalSelect();
	}

	async function loadPreferences() {
		const preferences = await window.appHelpers.loadPreferences(api);
		const prefOrder = preferences.inventory_column_order;
		if (Array.isArray(prefOrder) && prefOrder.length) {
			const valid = prefOrder.filter((field) => INVENTORY_COLUMNS.some((column) => column.field === field));
			if (valid.length === INVENTORY_COLUMNS.length) {
				state.columnOrder = valid;
			}
		}

		const prefWidths = preferences.inventory_column_widths;
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

		const prefPageSize = Number(preferences.inventory_page_size);
		if (Number.isFinite(prefPageSize) && prefPageSize >= 25 && prefPageSize <= 500) {
			state.perPage = prefPageSize;
		}

		const prefDensity = preferences.inventory_table_density;
		if (prefDensity === "compact" || prefDensity === "normal") {
			state.tableDensity = prefDensity;
		}
	}

	async function loadItems() {
		if (itemsLoadAbortController) {
			itemsLoadAbortController.abort();
		}
		itemsLoadAbortController = new AbortController();
		const requestId = ++itemsLoadRequestId;

		const params = new URLSearchParams();
		if (state.activeBlockId) params.set("bloque_id", state.activeBlockId);
		if (state.activeFloorId) params.set("piso_id", state.activeFloorId);
		if (state.activeAreaId) params.set("area_id", state.activeAreaId);
		if (state.search) params.set("search", state.search);
		params.set("order", state.order);
		params.set("page", String(state.page));
		params.set("per_page", String(state.perPage));

		try {
			const response = await api.get(`/api/inventario?${params.toString()}`, {
				signal: itemsLoadAbortController.signal,
			});
			if (requestId !== itemsLoadRequestId) return;

			state.items = response.data || [];
			if (state.selectedRowId && !state.items.some((item) => String(item.id) === String(state.selectedRowId))) {
				state.selectedRowId = null;
			}
			state.visibleItems = state.items;
			state.totalItems = Number(response.pagination?.total || state.items.length || 0);
			state.totalPages = Math.max(Number(response.pagination?.total_pages || 1), 1);
			state.page = Math.min(Math.max(Number(response.pagination?.page || state.page), 1), state.totalPages);
			state.perPage = Number(response.pagination?.per_page || state.perPage);
			renderRows();
			renderPaginationMeta();
		} catch (error) {
			if (error?.name === "AbortError") return;
			throw error;
		} finally {
			if (requestId === itemsLoadRequestId) {
				itemsLoadAbortController = null;
			}
		}
	}

	async function refreshItemsTable() {
		await loadItems();
	}

	async function saveCell(id, field, value, options = {}) {
		const payload = { [field]: value };
		if (field === "ubicacion") {
			const areaEntry = await resolveAreaEntryFromLocationValueBackend(value);
			if (areaEntry) {
				payload.ubicacion = areaEntry.nombre;
				payload.area_id = Number(areaEntry.id);
			}
		}
		if (options.forceDuplicate) payload.force_duplicate = true;
		const response = await api.send(`/api/inventario/${id}`, "PATCH", payload);
		return response.data || payload;
	}

	async function removeItem(itemId) {
		if (!itemId) return;
		try {
			await api.send(`/api/inventario/${itemId}`, "DELETE", {});
		} catch (error) {
			if (error?.status !== 404) {
				throw error;
			}
			notify("El registro ya no existe o fue eliminado en otra acción.");
		} finally {
			state.selectedRowId = null;
			nodes.contextMenu?.classList.add("d-none");
			document.querySelectorAll("#tabla-inventario tbody tr.row-selected").forEach((el) => {
				el.classList.remove("row-selected");
			});
		}
		await loadItems();
	}

	function resetAddItemForm() {
		const inputs = document.querySelectorAll("#form-agregar-item [data-field]");
		inputs.forEach((input) => {
			if (input.dataset.field === "cantidad") {
				input.value = "1";
				return;
			}
			input.value = "";
		});
		if (nodes.excelSingleRow) nodes.excelSingleRow.value = "";
		if (state.activeAreaId && nodes.modalAreaSelect) {
			nodes.modalAreaSelect.value = String(state.activeAreaId);
		}
		syncModalLocationFromSelection();
	}

	function setAddModalMode({ editing = false, itemId = null } = {}) {
		state.editingItemId = editing ? itemId : null;
		if (addModalTitle) {
			addModalTitle.textContent = editing ? "Editar ítem" : "Agregar nuevo ítem";
		}
		if (nodes.modalAddButton) {
			nodes.modalAddButton.innerHTML = editing
				? '<i class="bi bi-check-circle me-1"></i>Guardar cambios'
				: '<i class="bi bi-check-circle me-1"></i>Guardar ítem';
		}
	}

	function collectFormPayload() {
		const payload = {};
		const inputs = document.querySelectorAll("#form-agregar-item [data-field]");
		inputs.forEach((input) => {
			let value = input.value.trim();
			if (input.dataset.field === "cod_inventario" || input.dataset.field === "cod_esbye") {
				value = normalizeCodeToPlaceholder(value);
			}
			if (input.dataset.field === "cantidad") {
				const quantity = Number(value);
				value = Number.isInteger(quantity) && quantity > 0 ? quantity : null;
			}
			if (input.dataset.field === "valor") {
				value = helper.parseDecimalWithComma(value);
			}
			if (input.dataset.field === "valor_esbye") {
				value = helper.parseDecimalWithComma(value);
			}
			if (input.dataset.field === "area_id") {
				value = value === "" ? null : parseInt(value, 10);
			}
			if (value !== "" && value !== null) {
				payload[input.dataset.field] = value;
			}
		});
		if (!payload.cantidad) {
			throw new Error("La cantidad debe ser un entero mayor que 0.");
		}
		if (!payload.area_id && state.activeAreaId) {
			payload.area_id = parseInt(state.activeAreaId, 10);
		}
		if (!payload.ubicacion) {
			payload.ubicacion = composeSelectedLocation() || "";
		}
		return payload;
	}

	async function openEditModal(itemId) {
					   // Mostrar campos justificación y procedencia solo en edición
					   document.getElementById("wrap-justificacion")?.classList.remove("d-none");
					   document.getElementById("wrap-procedencia")?.classList.remove("d-none");
		await loadModalParams();
		const response = await api.get(`/api/inventario/${itemId}`);
		const item = response.data || {};
		const inputs = document.querySelectorAll("#form-agregar-item [data-field]");
		inputs.forEach((input) => {
			const field = input.dataset.field;
			const rawValue = item[field];
			if (field === "cantidad") {
				input.value = rawValue ?? 1;
				return;
			}
			if (field === "valor" || field === "valor_esbye") {
				if (rawValue === null || rawValue === undefined || rawValue === "") {
					input.value = "";
					return;
				}
				input.value = Number(rawValue).toLocaleString("es-EC", {
					minimumFractionDigits: 2,
					maximumFractionDigits: 2,
				});
				return;
			}
			if (field === "fecha_adquisicion" || field === "fecha_adquisicion_esbye") {
				input.value = toInputDate(rawValue);
				return;
			}
			if (field === "area_id") {
				input.value = rawValue ? String(rawValue) : "";
				return;
			}
			input.value = rawValue ?? "";
		});

		if (item.area_id && nodes.modalAreaSelect) {
			nodes.modalAreaSelect.value = String(item.area_id);
		}
		if (nodes.modalUbicacion) {
			nodes.modalUbicacion.value = String(item.ubicacion || "");
		}
		if (nodes.modalUbicacionEsbye) {
			nodes.modalUbicacionEsbye.value = String(item.ubicacion_esbye || "");
		}
		setAddModalMode({ editing: true, itemId });
		addModal.show();
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
				<div class="col-12"><strong>Justificación:</strong> ${item.justificacion || "-"}</div>
				<div class="col-12"><strong>Procedencia:</strong> ${item.procedencia || "-"}</div>
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

	async function startEdit(cell) {
		if (!cell.classList.contains("editable-cell")) return;
		if (cell.querySelector("input, select, textarea")) return;

		const id = cell.dataset.id;
		const field = cell.dataset.field;
		const oldRawValue = String(getRowFieldValue(id, field) ?? "");
		const setCellDisplay = (value) => {
			const displayValue = formatValue(field, value);
			cell.textContent = displayValue;
			cell.title = String(displayValue || "");
		};

		const trySaveValue = async (newValue) => {
			try {
				const savedPayload = await saveCell(id, field, newValue);
				const committedValue = Object.prototype.hasOwnProperty.call(savedPayload, field)
					? savedPayload[field]
					: newValue;
				setCellDisplay(committedValue);
				updateLocalItemValue(id, field, committedValue);
				if (field === "ubicacion" && Object.prototype.hasOwnProperty.call(savedPayload, "area_id")) {
					updateLocalItemValue(id, "area_id", savedPayload.area_id);
				}
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
						setCellDisplay(oldRawValue);
						return;
					}
					try {
						const savedPayload = await saveCell(id, field, newValue, { forceDuplicate: true });
						const committedValue = Object.prototype.hasOwnProperty.call(savedPayload, field)
							? savedPayload[field]
							: newValue;
						setCellDisplay(committedValue);
						updateLocalItemValue(id, field, committedValue);
						if (field === "ubicacion" && Object.prototype.hasOwnProperty.call(savedPayload, "area_id")) {
							updateLocalItemValue(id, "area_id", savedPayload.area_id);
						}
						return;
					} catch (forcedError) {
						setCellDisplay(oldRawValue);
						notify(forcedError.message, true);
						return;
					}
				}
				setCellDisplay(oldRawValue);
				notify(error.message, true);
			}
		};

		if (INLINE_SELECT_FIELDS.has(field)) {
			await ensureModalOptions();
			const selectConfig = getModalSelectConfig(field);
			if (!selectConfig) return;

			const select = document.createElement("select");
			select.className = "form-select form-select-sm";
			renderSelectWithQuickAdd(select, selectConfig, oldRawValue);
			if (field === "ubicacion") {
				const rowAreaId = getRowFieldValue(id, "area_id");
				if (rowAreaId !== null && rowAreaId !== undefined && String(rowAreaId).trim() !== "") {
					const linkedArea = flattenAreas().find((entry) => String(entry.id) === String(rowAreaId));
					if (linkedArea && linkedArea.nombre) {
						select.value = String(linkedArea.nombre);
					}
				}

				if (!String(select.value || "").trim() && String(oldRawValue || "").trim()) {
					const resolvedArea = await resolveAreaEntryFromLocationValueBackend(oldRawValue);
					if (resolvedArea && resolvedArea.nombre) {
						select.value = String(resolvedArea.nombre);
					}
				}
			}
			cell.innerHTML = "";
			cell.appendChild(select);
			select.focus();

			let finished = false;
			let openingQuickModal = false;
			let changedByUser = false;

			const finish = async (value) => {
				if (finished) return;
				finished = true;
				await trySaveValue(String(value || "").trim());
			};

			select.addEventListener("change", async () => {
				if (selectConfig.quickMode && select.value === INLINE_ADD_OPTION_VALUE) {
					openingQuickModal = true;
					const createdValue = await openQuickAddModal(selectConfig.quickMode);
					openingQuickModal = false;
					if (!createdValue) {
						setCellDisplay(oldRawValue);
						finished = true;
						return;
					}
					await loadModalParams();
					await finish(createdValue);
					return;
				}
				changedByUser = true;
				await finish(select.value);
			});

			select.addEventListener("blur", async () => {
				if (finished || openingQuickModal) return;
				const hasOldValue = String(oldRawValue || "").trim().length > 0;
				const hasCurrentValue = String(select.value || "").trim().length > 0;
				if (!changedByUser && hasOldValue && !hasCurrentValue) {
					await finish(oldRawValue);
					return;
				}
				await finish(select.value);
			});

			select.addEventListener("keydown", async (event) => {
				if (event.key === "Escape") {
					event.preventDefault();
					finished = true;
					setCellDisplay(oldRawValue);
				}
				if (event.key === "Enter") {
					event.preventDefault();
					await finish(select.value);
				}
			});
			return;
		}

		const inputConfig = INLINE_INPUT_CONFIG[field] || { type: "text" };
		const input = document.createElement("input");
		input.type = inputConfig.type;
		input.className = "form-control form-control-sm";
		if (inputConfig.min) input.min = inputConfig.min;
		if (inputConfig.step) input.step = inputConfig.step;
		input.value = input.type === "date" ? toInputDate(oldRawValue) : oldRawValue;
		cell.innerHTML = "";
		cell.appendChild(input);
		input.focus();
		if (input.type !== "date") input.select();

		let done = false;
		const commit = async () => {
			if (done) return;
			done = true;
			let newValue = String(input.value || "").trim();
			if (field === "cod_inventario" || field === "cod_esbye") {
				newValue = normalizeCodeToPlaceholder(newValue);
			}
			if (field === "cantidad") {
				const quantity = Number(newValue);
				if (!Number.isInteger(quantity) || quantity <= 0) {
					setCellDisplay(oldRawValue);
					notify("La cantidad debe ser un entero mayor que 0.", true);
					return;
				}
				newValue = String(quantity);
			}
			await trySaveValue(newValue);
		};

		input.addEventListener("blur", commit);
		input.addEventListener("keydown", async (event) => {
			if (event.key === "Enter") {
				event.preventDefault();
				await commit();
			}
			if (event.key === "Escape") {
				event.preventDefault();
				done = true;
				setCellDisplay(oldRawValue);
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
		await startEdit(cell);
	});

	// Manejador para clic derecho en tabla
	const tabla = document.getElementById("tabla-inventario");
	tabla.addEventListener("contextmenu", (event) => {
		const row = event.target.closest("tr");
		if (!row || !row.dataset.id) return; // Verificar que sea una fila de datos
		event.preventDefault();
		
		// Limpiar cualquier fila seleccionada previa
		document.querySelectorAll("#tabla-inventario tbody tr.row-selected").forEach(el => {
			el.classList.remove("row-selected");
		});
		
		// Marcar fila actual
		row.classList.add("row-selected");
		state.selectedRowId = row.dataset.id;
		
		// Mostrar menú
		nodes.contextMenu.style.left = `${event.pageX}px`;
		nodes.contextMenu.style.top = `${event.pageY}px`;
		nodes.contextMenu.classList.remove("d-none");
	});

	// Manejador para clic en opciones del menú
	nodes.contextMenu.addEventListener("click", async (event) => {
		const actionButton = event.target.closest("button[data-action]");
		if (!actionButton || !state.selectedRowId) return;
		
		const action = actionButton.dataset.action;
		try {
			if (action === "view") {
				await viewItem(state.selectedRowId);
			}
			if (action === "edit") {
				await openEditModal(state.selectedRowId);
			}
			if (action === "delete") {
				const confirmed = window.confirm("¿Seguro que deseas borrar este registro?");
				if (!confirmed) return;
				await removeItem(state.selectedRowId);
			}
		} catch (error) {
			notify(error.message, true);
		}
		
		// SIEMPRE cerrar después de la acción
		nodes.contextMenu.classList.add("d-none");
		document.querySelectorAll("#tabla-inventario tbody tr.row-selected").forEach(el => {
			el.classList.remove("row-selected");
		});
		state.selectedRowId = null;
	});

	// Manejador global para cerrar menú al hacer clic fuera
	document.addEventListener("click", (event) => {
		// Si menú está oculto, ignorar
		if (nodes.contextMenu.classList.contains("d-none")) return;
		
		// Si clic es dentro del menú, ignorar
		if (event.target.closest("#context-menu")) return;
		
		// Si clic es fuera, cerrar menú y limpiar clase
		nodes.contextMenu.classList.add("d-none");
		document.querySelectorAll("#tabla-inventario tbody tr.row-selected").forEach(el => {
			el.classList.remove("row-selected");
		});
		state.selectedRowId = null;
	});

	nodes.clearInventoryBtn?.addEventListener("click", async () => {
		await clearInventoryForTesting();
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
			if (error?.payload?.code === "area_not_found") {
				offerOpenSettingsForMissingAreas(error?.payload?.missing_areas || []);
			}
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
				const first = rows[0];
				const mapped = mapPastedRowBestEffortForModal(first);
				   MODAL_PASTE_FIELDS.forEach((field) => {
					   const input = document.querySelector(`#form-agregar-item [data-field='${field}']`);
					   // Solo asignar si el input está visible
					   if (input && mapped[field] !== undefined && input.offsetParent !== null) {
						   assignPastedValue(input, mapped[field]);
					   }
				   });

				const pastedLocation = String(mapped.ubicacion ?? "").trim();
				if (pastedLocation) {
					const resolvedArea = await resolveAreaFromPastedTextDefault(pastedLocation);
					if (resolvedArea && nodes.modalAreaSelect) {
						nodes.modalAreaSelect.value = String(resolvedArea.id);
						if (nodes.modalUbicacion) {
							nodes.modalUbicacion.value = resolvedArea.nombre;
						}
					} else {
						offerOpenSettingsForMissingAreas([pastedLocation]);
					}
				}

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

	if (typeof window.bindInventarioExport === "function") {
		window.bindInventarioExport({
			button: nodes.exportExcelBtn,
			getCurrentFilterParams,
		});
	}

	if (typeof window.initInventarioImportar === "function") {
		window.initInventarioImportar({
			onImportSuccess: async () => {
				await refreshItemsTable();
				notify("Importación de Excel completada.");
			},
		});
	}

	nodes.toggleDensityBtn?.addEventListener("click", () => {
		state.tableDensity = state.tableDensity === "compact" ? "normal" : "compact";
		applyDensityMode();
		saveDensityPreference().catch((error) => notify(error.message, true));
	});

	nodes.modalAddButton.addEventListener("click", async () => {
		const isEditing = Boolean(state.editingItemId);
		let payload;
		try {
			payload = collectFormPayload();
		} catch (validationError) {
			notify(validationError.message, true);
			return;
		}
			try {
				if (isEditing) {
					await api.send(`/api/inventario/${state.editingItemId}`, "PATCH", payload);
				} else {
					await api.send("/api/inventario", "POST", payload);
				}
				addModal.hide();
				resetAddItemForm();
				setAddModalMode({ editing: false });
				await refreshItemsTable();
				notify(isEditing ? "Ítem actualizado correctamente." : "Ítem guardado correctamente.");
			} catch (error) {
				const duplicateList = error?.payload?.duplicates;
				if (error?.status === 409 && Array.isArray(duplicateList) && duplicateList.length) {
					const modeText = isEditing ? "update" : "create";
					const confirmed = await openDuplicateModal({
						duplicates: duplicateList,
						payload,
						mode: modeText,
					});
					if (!confirmed) {
						return;
					}
					try {
						if (isEditing) {
							await api.send(`/api/inventario/${state.editingItemId}`, "PATCH", {
								...payload,
								force_duplicate: true,
							});
						} else {
							await api.send("/api/inventario", "POST", {
								...payload,
								force_duplicate: true,
							});
						}
						addModal.hide();
						resetAddItemForm();
						setAddModalMode({ editing: false });
						await refreshItemsTable();
						notify(isEditing ? "Ítem actualizado correctamente (código repetido autorizado)." : "Ítem guardado correctamente (código repetido autorizado).");
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
			state.modalOptions.estados = estadosRes.data || [];
			state.modalOptions.cuentas = cuentasRes.data || [];
			state.modalOptions.administradores = adminsRes.data || [];

			refreshAddItemSelects();
			attachQuickAddToSelect(document.querySelector("#form-agregar-item [data-field='estado']"), "estado");
			attachQuickAddToSelect(document.querySelector("#form-agregar-item [data-field='cuenta']"), "cuenta");
			attachQuickAddToSelect(document.querySelector("#form-agregar-item [data-field='usuario_final']"), "usuario_final");

			renderAreaModalSelect();
		} catch (error) {
			console.error("Error loading modal params:", error);
		}
	}

	addModalElement.addEventListener("shown.bs.modal", async () => {
                if (!state.editingItemId) {
                        await loadModalParams();
                        if (state.activeAreaId) {
                                nodes.modalAreaSelect.value = String(state.activeAreaId);
                        }
		}
	});

	addModalElement.addEventListener("hidden.bs.modal", () => {
		setAddModalMode({ editing: false });
		resetAddItemForm();
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

