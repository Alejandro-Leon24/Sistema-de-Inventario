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

const INLINE_ADD_OPTION_VALUE = "__add_new_option__";
const INLINE_SELECT_FIELDS = new Set(["estado", "cuenta", "usuario_final"]);
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
	return text
		.trim()
		.split("\n")
		.filter((line) => line.trim().length > 0)
		.map((line) => line.split("\t").map((cell) => cell.trim()));
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
		order: "asc",
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
		return null;
	}

	function renderSelectWithQuickAdd(select, config, selectedValue = "") {
		if (!select || !config) return;
		select.innerHTML = `<option value="">${config.placeholder}</option>`;
		(config.items || []).forEach((item) => {
			const option = document.createElement("option");
			option.value = config.optionValue(item);
			option.textContent = config.optionLabel(item);
			select.appendChild(option);
		});
		const addOption = document.createElement("option");
		addOption.value = INLINE_ADD_OPTION_VALUE;
		addOption.textContent = config.quickLabel;
		select.appendChild(addOption);
		if (selectedValue && Array.from(select.options).some((opt) => opt.value === selectedValue)) {
			select.value = selectedValue;
		} else {
			select.value = "";
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
		if (options.forceDuplicate) payload.force_duplicate = true;
		await api.send(`/api/inventario/${id}`, "PATCH", payload);
	}

	async function removeItem(itemId) {
		await api.send(`/api/inventario/${itemId}`, "DELETE", {});
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
			setCellDisplay(newValue);
			try {
				await saveCell(id, field, newValue);
				updateLocalItemValue(id, field, newValue);
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
						await saveCell(id, field, newValue, { forceDuplicate: true });
						updateLocalItemValue(id, field, newValue);
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
			cell.innerHTML = "";
			cell.appendChild(select);
			select.focus();

			let finished = false;
			let openingQuickModal = false;

			const finish = async (value) => {
				if (finished) return;
				finished = true;
				await trySaveValue(String(value || "").trim());
			};

			select.addEventListener("change", async () => {
				if (select.value === INLINE_ADD_OPTION_VALUE) {
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
				await finish(select.value);
			});

			select.addEventListener("blur", async () => {
				if (finished || openingQuickModal) return;
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

	if (typeof window.bindInventarioExport === "function") {
		window.bindInventarioExport({
			button: nodes.exportExcelBtn,
			getCurrentFilterParams,
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

