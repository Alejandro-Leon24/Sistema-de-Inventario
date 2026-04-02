const api = window.api;
const escapeHtmlText = window.appHelpers.escapeHtmlText;

async function initSettingsPage() {
	// Navigation between sections
	const menuButtons = document.querySelectorAll(".settings-menu-btn");
	const sections = document.querySelectorAll(".settings-section");
	const diagnosticsSearchInput = document.getElementById("diag-search-input");
	const diagnosticsRunButton = document.getElementById("btn-run-search-diagnostics");
	const diagnosticsSummary = document.getElementById("diag-search-summary");
	const diagnosticsPlan = document.getElementById("diag-search-plan");

	menuButtons.forEach((btn) => {
		btn.addEventListener("click", () => {
			const target = btn.dataset.target;
			
			menuButtons.forEach((b) => b.classList.remove("active"));
			btn.classList.add("active");
			
			sections.forEach((sec) => sec.classList.add("d-none"));
			const targetSection = document.getElementById(`section-${target}`);
			if (targetSection) targetSection.classList.remove("d-none");
		});
	});

	async function runInventoryDiagnostics() {
		if (!diagnosticsRunButton || !diagnosticsSummary || !diagnosticsPlan) return;
		diagnosticsRunButton.disabled = true;
		diagnosticsSummary.textContent = "Ejecutando diagnóstico...";
		try {
			const search = (diagnosticsSearchInput?.value || "").trim();
			const endpoint = `/api/inventario/search-diagnostics?search=${encodeURIComponent(search)}`;
			const response = await api.get(endpoint);
			const data = response.data || {};
			const summaryParts = [
				`WAL: ${String(data.journal_mode || "desconocido").toUpperCase()}`,
				`FTS disponible: ${data.fts_available ? "SI" : "NO"}`,
				`Usando FTS: ${data.using_fts ? "SI" : "NO"}`,
				`Filtro: ${data.where_sql || "(sin filtros)"}`,
			];
			diagnosticsSummary.innerHTML = summaryParts.map((part) => `<div>${escapeHtmlText(part)}</div>`).join("");
			diagnosticsPlan.textContent = JSON.stringify(data.query_plan || [], null, 2);
		} catch (error) {
			diagnosticsSummary.textContent = `Error: ${error.message}`;
			diagnosticsPlan.textContent = "[]";
		} finally {
			diagnosticsRunButton.disabled = false;
		}
	}

	diagnosticsRunButton?.addEventListener("click", runInventoryDiagnostics);
	diagnosticsSearchInput?.addEventListener("keydown", async (event) => {
		if (event.key !== "Enter") return;
		event.preventDefault();
		await runInventoryDiagnostics();
	});

	// Create shared modal instance (used for both locations and parameters)
	const modalElement = document.getElementById("modalEntidad");
	if (!modalElement) return; // Exit if modal doesn't exist
	
	const modal = new bootstrap.Modal(modalElement);
	const modalTitle = document.getElementById("modalEntidadLabel");
	const inputNombre = document.getElementById("entidad-nombre");
	const inputDescripcion = document.getElementById("entidad-descripcion");
	const btnGuardar = document.getElementById("btn-guardar-entidad");

	let modalMode = null; // Track current mode: "bloque", "piso", "area", "estados", etc.
	let loadParametros = async () => {};

	// Handle locations section (bloques, pisos, areas)
	const bloquesTabs = document.getElementById("bloques-tabs");
	const sinBloquesMsg = document.getElementById("sin-bloques-mensaje");
	const bloquesContent = document.getElementById("bloques-content");
	const areasSection = document.getElementById("areas-section");
	const sinPisosMsg = document.getElementById("sin-pisos-mensaje");
	const detalleUbicacionModalElement = document.getElementById("modalDetalleUbicacion");
	const detalleUbicacionBody = document.getElementById("modal-detalle-ubicacion-body");
	const detalleUbicacionTitle = document.getElementById("modalDetalleUbicacionLabel");
	const detalleUbicacionModal = detalleUbicacionModalElement
		? new bootstrap.Modal(detalleUbicacionModalElement)
		: null;
	if (bloquesTabs) {
		const pisosContainer = document.getElementById("pisos-container");
		const areasContainer = document.getElementById("areas-container");
		const btnAgregarBloque = document.getElementById("btn-agregar-bloque");
		const btnAgregarPiso = document.getElementById("btn-agregar-piso");
		const btnImportarUbicaciones = document.getElementById("btn-importar-ubicaciones");
		const btnExportarUbicaciones = document.getElementById("btn-exportar-ubicaciones");
		const importModalElement = document.getElementById("modalImportarUbicaciones");
		const importModal = importModalElement ? new bootstrap.Modal(importModalElement) : null;
		const importTextarea = document.getElementById("import-ubicaciones-texto");
		const importResumen = document.getElementById("import-ubicaciones-resumen");
		const btnConfirmarImport = document.getElementById("btn-confirmar-import-ubicaciones");

		const state = {
			structure: [],
			activeBlockId: null,
			activeFloorId: null,
			expandedFloorId: null,
		};

		const modalAreaElement = document.getElementById("modalArea");
		const modalArea = modalAreaElement ? new bootstrap.Modal(modalAreaElement) : null;
		const btnGuardarAreaDetalle = document.getElementById("btn-guardar-area-detalle");
		const areaFields = {
			nombre: document.getElementById("area-nombre"),
			descripcion: document.getElementById("area-descripcion"),
			identificacion_ambiente: document.getElementById("area-identificacion-ambiente"),
			metros_cuadrados: document.getElementById("area-metros-cuadrados"),
			alto: document.getElementById("area-alto"),
			senaletica: document.getElementById("area-senaletica"),
			cod_senaletica: document.getElementById("area-cod-senaletica"),
			infraestructura_fisica: document.getElementById("area-infraestructura-fisica"),
			piso_nombre: document.getElementById("area-piso-nombre"),
			estado_piso: document.getElementById("area-estado-piso"),
			material_techo: document.getElementById("area-material-techo"),
			puerta: document.getElementById("area-puerta"),
			material_puerta: document.getElementById("area-material-puerta"),
			responsable_admin_id: document.getElementById("area-responsable-admin-id"),
			estado_paredes: document.getElementById("area-estado-paredes"),
			estado_techo: document.getElementById("area-estado-techo"),
			estado_puerta: document.getElementById("area-estado-puerta"),
			cerradura: document.getElementById("area-cerradura"),
			nivel_seguridad: document.getElementById("area-nivel-seguridad"),
			sitio_profesor_mesa: document.getElementById("area-sitio-profesor-mesa"),
			sitio_profesor_silla: document.getElementById("area-sitio-profesor-silla"),
			pc_aula: document.getElementById("area-pc-aula"),
			proyector: document.getElementById("area-proyector"),
			pantalla_interactiva: document.getElementById("area-pantalla-interactiva"),
			pupitres_cantidad: document.getElementById("area-pupitres-cantidad"),
			pupitres_funcionan: document.getElementById("area-pupitres-funcionan"),
			pupitres_no_funcionan: document.getElementById("area-pupitres-no-funcionan"),
			pizarra: document.getElementById("area-pizarra"),
			pizarra_estado: document.getElementById("area-pizarra-estado"),
			ventanas_cantidad: document.getElementById("area-ventanas-cantidad"),
			ventanas_funcionan: document.getElementById("area-ventanas-funcionan"),
			ventanas_no_funcionan: document.getElementById("area-ventanas-no-funcionan"),
			aa_cantidad: document.getElementById("area-aa-cantidad"),
			aa_funcionan: document.getElementById("area-aa-funcionan"),
			aa_no_funcionan: document.getElementById("area-aa-no-funcionan"),
			ventiladores_cantidad: document.getElementById("area-ventiladores-cantidad"),
			ventiladores_funcionan: document.getElementById("area-ventiladores-funcionan"),
			ventiladores_no_funcionan: document.getElementById("area-ventiladores-no-funcionan"),
			wifi: document.getElementById("area-wifi"),
			red_lan: document.getElementById("area-red-lan"),
			red_lan_funcionan: document.getElementById("area-red-lan-funcionan"),
			red_lan_no_funcionan: document.getElementById("area-red-lan-no-funcionan"),
			red_inalambrica_cantidad: document.getElementById("area-red-inalambrica-cantidad"),
			iluminacion_funcionan: document.getElementById("area-iluminacion-funcionan"),
			iluminacion_no_funcionan: document.getElementById("area-iluminacion-no-funcionan"),
			luminarias_cantidad: document.getElementById("area-luminarias-cantidad"),
			puntos_electricos: document.getElementById("area-puntos-electricos"),
			puntos_electricos_funcionan: document.getElementById("area-puntos-electricos-funcionan"),
			puntos_electricos_no_funcionan: document.getElementById("area-puntos-electricos-no-funcionan"),
			puntos_electricos_cantidad: document.getElementById("area-puntos-electricos-cantidad"),
			capacidad_aulica: document.getElementById("area-capacidad-aulica"),
			capacidad_distanciamiento: document.getElementById("area-capacidad-distanciamiento"),
			ambiente_apto_retorno: document.getElementById("area-ambiente-apto-retorno"),
			observaciones_detalle: document.getElementById("area-observaciones-detalle"),
		};
		const btnDetalleCerrar = document.getElementById("btn-detalle-cancelar");
		const btnDetalleEliminar = document.getElementById("btn-detalle-eliminar");
		const btnDetalleEditar = document.getElementById("btn-detalle-editar");
		const btnDetalleGuardar = document.getElementById("btn-detalle-guardar");
		const btnDetalleCancelarEdicion = document.getElementById("btn-detalle-cancelar-edicion");
		const modalDeleteConfirmEl = document.getElementById("modalConfirmarEliminacionUbicacion");
		const modalDeleteConfirm = modalDeleteConfirmEl ? new bootstrap.Modal(modalDeleteConfirmEl) : null;
		const modalDeleteTitle = document.getElementById("modalConfirmarEliminacionUbicacionLabel");
		const modalDeleteBody = document.getElementById("modalConfirmarEliminacionUbicacionBody");
		const btnConfirmDeleteLocation = document.getElementById("btn-confirmar-eliminacion-ubicacion");

		function fillSelectWithText(select, values, placeholder = "Seleccionar") {
			if (!select) return;
			select.innerHTML = `<option value="">${placeholder}</option>`;
			(values || []).forEach((value) => {
				const option = document.createElement("option");
				option.value = value;
				option.textContent = value;
				select.appendChild(option);
			});
		}

		async function loadAreaFormOptions() {
			if (!modalArea) return;
			try {
				const [adminsRes, siNoRes, puertaRes, cerradurasRes, estadoPisoRes, materialTechoRes, materialPuertaRes, estadoPizarraRes] = await Promise.all([
					api.get("/api/administradores"),
					api.get("/api/parametros/si_no"),
					api.get("/api/parametros/estado_puerta"),
					api.get("/api/parametros/cerraduras"),
					api.get("/api/parametros/estado_piso"),
					api.get("/api/parametros/material_techo"),
					api.get("/api/parametros/material_puerta"),
					api.get("/api/parametros/estado_pizarra"),
				]);

				const adminSelect = areaFields.responsable_admin_id;
				if (adminSelect) {
					adminSelect.innerHTML = '<option value="">Sin asignar</option>';
					(adminsRes.data || []).forEach((admin) => {
						const option = document.createElement("option");
						option.value = admin.id;
						option.textContent = admin.nombre;
						adminSelect.appendChild(option);
					});
				}

				const yesNoValues = (siNoRes.data || []).map((item) => item.nombre);
				fillSelectWithText(areaFields.estado_paredes, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.estado_techo, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.senaletica, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.puerta, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.nivel_seguridad, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.sitio_profesor_mesa, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.sitio_profesor_silla, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.pc_aula, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.proyector, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.pantalla_interactiva, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.pizarra, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.wifi, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.red_lan, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.puntos_electricos, yesNoValues, "Seleccionar");
				fillSelectWithText(areaFields.ambiente_apto_retorno, yesNoValues, "Seleccionar");

				fillSelectWithText(areaFields.estado_puerta, (puertaRes.data || []).map((item) => item.nombre), "Seleccionar");
				fillSelectWithText(areaFields.cerradura, (cerradurasRes.data || []).map((item) => item.nombre), "Seleccionar");
				fillSelectWithText(areaFields.estado_piso, (estadoPisoRes.data || []).map((item) => item.nombre), "Seleccionar");
				fillSelectWithText(areaFields.material_techo, (materialTechoRes.data || []).map((item) => item.nombre), "Seleccionar");
				fillSelectWithText(areaFields.material_puerta, (materialPuertaRes.data || []).map((item) => item.nombre), "Seleccionar");
				fillSelectWithText(areaFields.pizarra_estado, (estadoPizarraRes.data || []).map((item) => item.nombre), "Seleccionar");

				const activeFloor = getActiveFloor();
				if (areaFields.piso_nombre) {
					areaFields.piso_nombre.value = activeFloor?.nombre || "";
				}
			} catch (error) {
				console.error("Error cargando opciones de área:", error);
			}
		}

		function resetAreaForm() {
			Object.values(areaFields).forEach((field) => {
				if (!field) return;
				field.value = "";
			});
		}

		function getActiveBlock() {
			return state.structure.find((block) => block.id === state.activeBlockId) || null;
		}

		function getActiveFloor() {
			const block = getActiveBlock();
			if (!block) return null;
			return block.pisos.find((piso) => piso.id === state.activeFloorId) || null;
		}

		function escapeHtml(value) {
			return String(value || "")
				.replaceAll("&", "&amp;")
				.replaceAll("<", "&lt;")
				.replaceAll(">", "&gt;")
				.replaceAll('"', "&quot;")
				.replaceAll("'", "&#39;");
		}

		function normalizeKey(value) {
			return String(value || "")
				.normalize("NFD")
				.replace(/[\u0300-\u036f]/g, "")
				.toLowerCase()
				.replace(/\s+/g, " ")
				.trim();
		}

		function parseLocationLine(rawLine) {
			const line = String(rawLine || "").trim();
			if (!line) return null;
			if (/^ubicacion/i.test(line) || /^nombre del edificio/i.test(line)) return null;

			const match = line.match(/^BLOQUE\s*"?([^"\-,]+)"?\s*-\s*([^,]+),\s*(.+)$/i);
			if (!match) return null;

			const blockCode = String(match[1] || "").trim().toUpperCase();
			const blockName = String(match[2] || "").trim().toUpperCase();
			const floorName = String(match[3] || "").trim().replace(/\s+/g, " ").toUpperCase();
			if (!blockCode || !blockName || !floorName) return null;

			return {
				blockDisplayName: `BLOQUE "${blockCode}" - ${blockName}`,
				blockDescription: `Edificio ${blockName} (Código ${blockCode})`,
				floorName,
			};
		}

		function summarizeImportLines(lines) {
			const parsed = lines
				.map((line) => parseLocationLine(line))
				.filter(Boolean);
			const uniqueBlocks = new Set(parsed.map((item) => normalizeKey(item.blockDisplayName)));
			const uniqueFloorPairs = new Set(parsed.map((item) => `${normalizeKey(item.blockDisplayName)}||${normalizeKey(item.floorName)}`));
			return {
				validRows: parsed.length,
				uniqueBlocks: uniqueBlocks.size,
				uniqueFloors: uniqueFloorPairs.size,
			};
		}

		async function importBlockFloorLocationsFromText(rawText) {
			const lines = String(rawText || "")
				.split(/\r?\n/)
				.map((line) => line.trim())
				.filter((line) => line.length > 0);

			const parsed = lines.map((line) => parseLocationLine(line)).filter(Boolean);
			if (!parsed.length) {
				throw new Error("No se detectaron líneas válidas. Usa formato: BLOQUE \"A\" - PRINCIPAL, PLANTA BAJA");
			}

			const existingByBlockKey = new Map(
				(state.structure || []).map((block) => [normalizeKey(block.nombre), block])
			);

			let createdBlocks = 0;
			let createdFloors = 0;

			const uniqueBlocks = [];
			const blockSeen = new Set();
			parsed.forEach((entry) => {
				const key = normalizeKey(entry.blockDisplayName);
				if (blockSeen.has(key)) return;
				blockSeen.add(key);
				uniqueBlocks.push(entry);
			});

			for (const blockEntry of uniqueBlocks) {
				const key = normalizeKey(blockEntry.blockDisplayName);
				if (existingByBlockKey.has(key)) continue;
				await api.send("/api/bloques", "POST", {
					nombre: blockEntry.blockDisplayName,
					descripcion: blockEntry.blockDescription,
				});
				createdBlocks += 1;
			}

			await loadStructure();

			const refreshedByBlockKey = new Map(
				(state.structure || []).map((block) => [normalizeKey(block.nombre), block])
			);

			const floorSeen = new Set();
			for (const entry of parsed) {
				const blockKey = normalizeKey(entry.blockDisplayName);
				const floorKey = normalizeKey(entry.floorName);
				const pairKey = `${blockKey}||${floorKey}`;
				if (floorSeen.has(pairKey)) continue;
				floorSeen.add(pairKey);

				const block = refreshedByBlockKey.get(blockKey);
				if (!block) continue;
				const existsFloor = (block.pisos || []).some((piso) => normalizeKey(piso.nombre) === floorKey);
				if (existsFloor) continue;

				await api.send("/api/pisos", "POST", {
					bloque_id: block.id,
					nombre: entry.floorName,
					descripcion: null,
				});
				createdFloors += 1;
			}

			await loadStructure();
			return {
				createdBlocks,
				createdFloors,
				rowsRead: parsed.length,
			};
		}

		function showLocationDetail(title, contentHtml) {
			if (!detalleUbicacionModal) return;
			detalleUbicacionTitle.textContent = title;
			detalleUbicacionBody.innerHTML = contentHtml;
			detalleUbicacionModal.show();
		}

		function setLocationDetailContext(type, id, name) {
			locationDetailContext = { type, id, name };
		}

		async function askDeleteLocationWithImpact(context) {
			if (!context || !modalDeleteConfirm || !modalDeleteBody || !btnConfirmDeleteLocation) {
				return false;
			}

			const response = await api.get(`/api/ubicaciones/impacto?entity=${encodeURIComponent(context.type)}&id=${context.id}`);
			const impact = response.data || {};
			const pisos = Number(impact.pisos || 0);
			const areas = Number(impact.areas || 0);
			const items = Number(impact.items || 0);
			const hasRelations = pisos > 0 || areas > 0 || items > 0;

			modalDeleteTitle.textContent = hasRelations
				? `Advertencia de eliminación en cascada (${context.type})`
				: `Confirmar eliminación de ${context.type}`;

			modalDeleteBody.innerHTML = hasRelations
				? `
					<div class="alert alert-warning mb-2">
						<strong>Ya existen datos relacionados.</strong>
					</div>
					<ul class="mb-2">
						${context.type === "bloque" ? `<li>Pisos relacionados: <strong>${pisos}</strong></li>` : ""}
						${context.type !== "area" ? `<li>Áreas relacionadas: <strong>${areas}</strong></li>` : ""}
						<li>Ítems de inventario relacionados: <strong>${items}</strong></li>
					</ul>
					<p class="mb-1">Si eliminas <strong>${escapeHtml(context.name || context.type)}</strong>, se borrará en cascada toda su información dependiente (incluyendo ítems).</p>
					<p class="text-muted small mb-0">Recomendación: si no deseas perder datos, usa Editar en lugar de Eliminar.</p>
				`
				: `<p class="mb-0">¿Seguro que deseas eliminar <strong>${escapeHtml(context.name || context.type)}</strong>?</p>`;

			return new Promise((resolve) => {
				let resolved = false;
				const onConfirm = () => {
					if (resolved) return;
					resolved = true;
					cleanup();
					resolve(true);
					modalDeleteConfirm.hide();
				};
				const onHidden = () => {
					if (resolved) return;
					resolved = true;
					cleanup();
					resolve(false);
				};
				const cleanup = () => {
					btnConfirmDeleteLocation.removeEventListener("click", onConfirm);
					modalDeleteConfirmEl.removeEventListener("hidden.bs.modal", onHidden);
				};

				btnConfirmDeleteLocation.addEventListener("click", onConfirm);
				modalDeleteConfirmEl.addEventListener("hidden.bs.modal", onHidden);
				modalDeleteConfirm.show();
			});
		}

		async function deleteLocationByContext(context) {
			if (!context) return;
			const routeMap = {
				bloque: `/api/bloques/${context.id}`,
				piso: `/api/pisos/${context.id}`,
				area: `/api/areas/${context.id}`,
			};
			const route = routeMap[context.type];
			if (!route) return;
			await api.send(route, "DELETE", {});
			await loadStructure();
			notify(`${context.type.charAt(0).toUpperCase() + context.type.slice(1)} eliminado correctamente.`);
		}

		function setDetailButtonsMode(mode = "view", enableEdit = false) {
			if (!btnDetalleCerrar || !btnDetalleEditar || !btnDetalleGuardar || !btnDetalleCancelarEdicion) return;
			if (mode === "edit") {
				btnDetalleCerrar.classList.add("d-none");
				btnDetalleEditar.classList.add("d-none");
				btnDetalleGuardar.classList.remove("d-none");
				btnDetalleCancelarEdicion.classList.remove("d-none");
				return;
			}
			btnDetalleGuardar.classList.add("d-none");
			btnDetalleCancelarEdicion.classList.add("d-none");
			btnDetalleCerrar.classList.remove("d-none");
			if (enableEdit) {
				btnDetalleEditar.classList.remove("d-none");
			} else {
				btnDetalleEditar.classList.add("d-none");
			}
		}

		function toNumberOrNull(value) {
			if (value === "" || value === null || value === undefined) return null;
			const number = Number(value);
			return Number.isNaN(number) ? null : number;
		}

		function findAreaInStructure(blockId, pisoId, areaId) {
			const block = state.structure.find((entry) => String(entry.id) === String(blockId));
			if (!block) return null;
			const piso = (block.pisos || []).find((entry) => String(entry.id) === String(pisoId));
			if (!piso) return null;
			const area = (piso.areas || []).find((entry) => String(entry.id) === String(areaId));
			if (!area) return null;
			return { block, piso, area };
		}

		async function loadAreaDetailOptions() {
			const [adminsRes, siNoRes, puertaRes, cerradurasRes, estadoPisoRes, materialTechoRes, materialPuertaRes, estadoPizarraRes] = await Promise.all([
				api.get("/api/administradores"),
				api.get("/api/parametros/si_no"),
				api.get("/api/parametros/estado_puerta"),
				api.get("/api/parametros/cerraduras"),
				api.get("/api/parametros/estado_piso"),
				api.get("/api/parametros/material_techo"),
				api.get("/api/parametros/material_puerta"),
				api.get("/api/parametros/estado_pizarra"),
			]);
			return {
				admins: adminsRes.data || [],
				yesNo: (siNoRes.data || []).map((entry) => entry.nombre),
				estadoPuerta: (puertaRes.data || []).map((entry) => entry.nombre),
				cerraduras: (cerradurasRes.data || []).map((entry) => entry.nombre),
				estadoPiso: (estadoPisoRes.data || []).map((entry) => entry.nombre),
				materialTecho: (materialTechoRes.data || []).map((entry) => entry.nombre),
				materialPuerta: (materialPuertaRes.data || []).map((entry) => entry.nombre),
				estadoPizarra: (estadoPizarraRes.data || []).map((entry) => entry.nombre),
			};
		}

		function renderAreaDetailReadOnly(context, options = {}) {
			const { block, piso, area } = context;
			const getResponsableName = () => {
				if (!area.responsable_admin_id) return "Sin asignar";
				const admin = (options.admins || []).find(a => a.id === area.responsable_admin_id);
				return admin?.nombre || "Sin asignar";
			};
			const infoRows = [
				["Bloque", block.nombre],
				["Área", area.nombre],
				["Responsable del ambiente", getResponsableName()],
				["Identificación ambiente", area.identificacion_ambiente || "-"],
				["Metros cuadrados", area.metros_cuadrados || "-"],
				["Alto", area.alto ?? "-"],
				["Señalética", area.senaletica || "-"],
				["Codificación de la señalética", area.cod_senaletica || "-"],
				["Infraestructura física", area.infraestructura_fisica || "-"],
				["Piso", piso.nombre],
				["Estado piso", area.estado_piso || "-"],
				["Estado paredes", area.estado_paredes || "-"],
				["Estado techo", area.estado_techo || "-"],
				["Material techo", area.material_techo || "-"],
				["Puerta", area.puerta || "-"],
				["Material puerta", area.material_puerta || "-"],
				["Estado puerta", area.estado_puerta || "-"],
				["Cerradura", area.cerradura || "-"],
				["Nivel seguridad", area.nivel_seguridad || "-"],
				["Sitio profesor mesa", area.sitio_profesor_mesa || "-"],
				["Sitio profesor silla", area.sitio_profesor_silla || "-"],
				["PC en aula", area.pc_aula || "-"],
				["Proyector", area.proyector || "-"],
				["Pantalla interactiva", area.pantalla_interactiva || "-"],
				["Pupitres", area.pupitres_cantidad ?? "-"],
				["Pupitres funcionando", area.pupitres_funcionan ?? "-"],
				["Pupitres no funcionando", area.pupitres_no_funcionan ?? "-"],
				["Pizarra", area.pizarra || "-"],
				["Estado pizarra", area.pizarra_estado || "-"],
				["Ventanas", area.ventanas_cantidad ?? "-"],
				["Ventanas funcionando", area.ventanas_funcionan ?? "-"],
				["Ventanas no funcionando", area.ventanas_no_funcionan ?? "-"],
				["A/A cantidad", area.aa_cantidad ?? "-"],
				["A/A funcionando", area.aa_funcionan ?? "-"],
				["A/A no funcionando", area.aa_no_funcionan ?? "-"],
				["Ventiladores", area.ventiladores_cantidad ?? "-"],
				["Ventiladores funcionando", area.ventiladores_funcionan ?? "-"],
				["Ventiladores no funcionando", area.ventiladores_no_funcionan ?? "-"],
				["WIFI", area.wifi || "-"],
				["Red LAN", area.red_lan || "-"],
				["LAN funcionando", area.red_lan_funcionan ?? "-"],
				["LAN no funcionando", area.red_lan_no_funcionan ?? "-"],
				["Red inalámbrica", area.red_inalambrica_cantidad ?? "-"],
				["Iluminación funcionando", area.iluminacion_funcionan ?? "-"],
				["Iluminación no funcionando", area.iluminacion_no_funcionan ?? "-"],
				["Luminarias", area.luminarias_cantidad ?? "-"],
				["Puntos eléctricos", area.puntos_electricos || "-"],
				["Ptos. eléctricos funcionando", area.puntos_electricos_funcionan ?? "-"],
				["Ptos. eléctricos no funcionando", area.puntos_electricos_no_funcionan ?? "-"],
				["Puntos eléctricos cantidad", area.puntos_electricos_cantidad ?? "-"],
				["Capacidad áulica", area.capacidad_aulica ?? "-"],
				["Capacidad con distanciamiento", area.capacidad_distanciamiento ?? "-"],
				["Apto retorno presencial", area.ambiente_apto_retorno || "-"],
				["Observaciones", area.observaciones_detalle || "-"],
			];
			showLocationDetail(
				`Detalle del área: ${area.nombre}`,
				`
				<div class="row g-2">
					${infoRows
						.map(
							([label, value]) =>
								`<div class="col-md-6"><strong>${escapeHtml(label)}:</strong> ${escapeHtml(value)}</div>`
						)
						.join("")}
				</div>
				`
			);
			setDetailButtonsMode("view", true);
		}

		function renderAreaDetailEditForm(context, options) {
			const { piso, area } = context;
			const ADD_ADMIN_OPTION = "__add_new_admin__";
			const makeOptions = (values, selected) =>
				[`<option value="">Seleccionar</option>`]
					.concat(
						(values || []).map(
							(value) =>
								`<option value="${escapeHtml(value)}" ${String(selected || "") === String(value) ? "selected" : ""}>${escapeHtml(value)}</option>`
						)
					)
					.join("");
			const buildAdminOptions = (admins, selectedId) =>
				[`<option value="">Sin asignar</option>`]
					.concat(
						(admins || []).map(
							(admin) =>
								`<option value="${admin.id}" ${String(selectedId || "") === String(admin.id) ? "selected" : ""}>${escapeHtml(admin.nombre)}</option>`
						)
					)
					.concat([`<option value="${ADD_ADMIN_OPTION}">+ Agregar responsable...</option>`])
					.join("");
			const adminOptions = buildAdminOptions(options.admins || [], area.responsable_admin_id || "");

			detalleUbicacionTitle.textContent = `Editar área: ${area.nombre}`;
			detalleUbicacionBody.innerHTML = `
				<div class="row g-2" id="detalle-area-edit-form">
					<div class="col-md-6"><label class="form-label small">Nombre</label><input type="text" class="form-control form-control-sm" data-field="nombre" value="${escapeHtml(area.nombre || "")}"></div>
					<div class="col-md-6">
						<label class="form-label small">Responsable</label>
						<div class="input-group input-group-sm">
							<select class="form-select form-select-sm" data-field="responsable_admin_id">${adminOptions}</select>
						</div>
					</div>
					<div class="col-md-6"><label class="form-label small">Identificación ambiente</label><input type="text" class="form-control form-control-sm" data-field="identificacion_ambiente" value="${escapeHtml(area.identificacion_ambiente || "")}"></div>
					<div class="col-md-4"><label class="form-label small">Metros cuadrados</label><input type="text" class="form-control form-control-sm" data-field="metros_cuadrados" value="${escapeHtml(area.metros_cuadrados || "")}"></div>
					<div class="col-md-2"><label class="form-label small">Alto</label><input type="number" min="0" step="0.01" class="form-control form-control-sm" data-field="alto" value="${area.alto ?? ""}"></div>
					<div class="col-md-3"><label class="form-label small">Señalética</label><select class="form-select form-select-sm" data-field="senaletica">${makeOptions(options.yesNo, area.senaletica)}</select></div>
					<div class="col-md-9"><label class="form-label small">Codificación señalética</label><input type="text" class="form-control form-control-sm" data-field="cod_senaletica" value="${escapeHtml(area.cod_senaletica || "")}"></div>
					<div class="col-12"><label class="form-label small">Infraestructura física</label><textarea rows="2" class="form-control form-control-sm" data-field="infraestructura_fisica">${escapeHtml(area.infraestructura_fisica || "")}</textarea></div>
					
					<div class="col-md-3"><label class="form-label small">Nombre piso</label><input type="text" class="form-control form-control-sm" data-field="nombre_piso" value="${escapeHtml(piso.nombre || "")}" readonly></div>
					<div class="col-md-3"><label class="form-label small">Estado piso</label><select class="form-select form-select-sm" data-field="estado_piso">${makeOptions(options.estadoPiso, area.estado_piso)}</select></div>
					<div class="col-md-3"><label class="form-label small">Estado paredes</label><select class="form-select form-select-sm" data-field="estado_paredes">${makeOptions(options.yesNo, area.estado_paredes)}</select></div>
					<div class="col-md-3"><label class="form-label small">Estado techo</label><select class="form-select form-select-sm" data-field="estado_techo">${makeOptions(options.yesNo, area.estado_techo)}</select></div>
					
					<div class="col-md-4"><label class="form-label small">Material techo</label><select class="form-select form-select-sm" data-field="material_techo">${makeOptions(options.materialTecho, area.material_techo)}</select></div>
					<div class="col-md-2"><label class="form-label small">Puerta</label><select class="form-select form-select-sm" data-field="puerta">${makeOptions(options.yesNo, area.puerta)}</select></div>
					<div class="col-md-6"><label class="form-label small">Material puerta</label><select class="form-select form-select-sm" data-field="material_puerta">${makeOptions(options.materialPuerta, area.material_puerta)}</select></div>
					

					<div class="col-md-3"><label class="form-label small">Estado puerta</label><select class="form-select form-select-sm" data-field="estado_puerta">${makeOptions(options.estadoPuerta, area.estado_puerta)}</select></div>
					<div class="col-md-3"><label class="form-label small">Cerradura</label><select class="form-select form-select-sm" data-field="cerradura">${makeOptions(options.cerraduras, area.cerradura)}</select></div>
					<div class="col-md-3"><label class="form-label small">Nivel seguridad</label><select class="form-select form-select-sm" data-field="nivel_seguridad">${makeOptions(options.yesNo, area.nivel_seguridad)}</select></div>
					<div class="col-md-3"><label class="form-label small">Profesor mesa</label><select class="form-select form-select-sm" data-field="sitio_profesor_mesa">${makeOptions(options.yesNo, area.sitio_profesor_mesa)}</select></div>
					<div class="col-md-3"><label class="form-label small">Profesor silla</label><select class="form-select form-select-sm" data-field="sitio_profesor_silla">${makeOptions(options.yesNo, area.sitio_profesor_silla)}</select></div>
					<div class="col-md-2"><label class="form-label small">PC aula</label><select class="form-select form-select-sm" data-field="pc_aula">${makeOptions(options.yesNo, area.pc_aula)}</select></div>
					<div class="col-md-3"><label class="form-label small">Proyector</label><select class="form-select form-select-sm" data-field="proyector">${makeOptions(options.yesNo, area.proyector)}</select></div>
					<div class="col-md-4"><label class="form-label small">Pantalla interactiva</label><select class="form-select form-select-sm" data-field="pantalla_interactiva">${makeOptions(options.yesNo, area.pantalla_interactiva)}</select></div>
					<div class="col-md-3"><label class="form-label small">Pupitres cant.</label><input type="number" min="0" class="form-control form-control-sm" data-field="pupitres_cantidad" value="${area.pupitres_cantidad ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">Pupitres funcionando</label><input type="number" min="0" class="form-control form-control-sm" data-field="pupitres_funcionan" value="${area.pupitres_funcionan ?? ""}"></div>
					<div class="col-md-5"><label class="form-label small">Pupitres no funcionando</label><input type="number" min="0" class="form-control form-control-sm" data-field="pupitres_no_funcionan" value="${area.pupitres_no_funcionan ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">Pizarra</label><select class="form-select form-select-sm" data-field="pizarra">${makeOptions(options.yesNo, area.pizarra)}</select></div>
					<div class="col-md-4"><label class="form-label small">Estado pizarra</label><select class="form-select form-select-sm" data-field="pizarra_estado">${makeOptions(options.estadoPizarra, area.pizarra_estado)}</select></div>
					<div class="col-md-4"><label class="form-label small">Ventanas cantidad</label><input type="number" min="0" class="form-control form-control-sm" data-field="ventanas_cantidad" value="${area.ventanas_cantidad ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">Ventanas funcionando</label><input type="number" min="0" class="form-control form-control-sm" data-field="ventanas_funcionan" value="${area.ventanas_funcionan ?? ""}"></div>
					<div class="col-md-5"><label class="form-label small">Ventanas no funcionando</label><input type="number" min="0" class="form-control form-control-sm" data-field="ventanas_no_funcionan" value="${area.ventanas_no_funcionan ?? ""}"></div>
					<div class="col-md-3"><label class="form-label small">A/A cantidad</label><input type="number" min="0" class="form-control form-control-sm" data-field="aa_cantidad" value="${area.aa_cantidad ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">A/A funcionando</label><input type="number" min="0" class="form-control form-control-sm" data-field="aa_funcionan" value="${area.aa_funcionan ?? ""}"></div>
					<div class="col-md-5"><label class="form-label small">A/A no funcionando</label><input type="number" min="0" class="form-control form-control-sm" data-field="aa_no_funcionan" value="${area.aa_no_funcionan ?? ""}"></div>
					<div class="col-md-3"><label class="form-label small">Ventiladores</label><input type="number" min="0" class="form-control form-control-sm" data-field="ventiladores_cantidad" value="${area.ventiladores_cantidad ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">Ventiladores func.</label><input type="number" min="0" class="form-control form-control-sm" data-field="ventiladores_funcionan" value="${area.ventiladores_funcionan ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">Ventiladores no func.</label><input type="number" min="0" class="form-control form-control-sm" data-field="ventiladores_no_funcionan" value="${area.ventiladores_no_funcionan ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">WIFI</label><select class="form-select form-select-sm" data-field="wifi">${makeOptions(options.yesNo, area.wifi)}</select></div>
					
					<div class="col-md-4"><label class="form-label small">Red LAN</label><select class="form-select form-select-sm" data-field="red_lan">${makeOptions(options.yesNo, area.red_lan)}</select></div>
					<div class="col-md-4"><label class="form-label small">LAN funcionando</label><input type="number" min="0" class="form-control form-control-sm" data-field="red_lan_funcionan" value="${area.red_lan_funcionan ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">LAN no funcionando</label><input type="number" min="0" class="form-control form-control-sm" data-field="red_lan_no_funcionan" value="${area.red_lan_no_funcionan ?? ""}"></div>
					
					<div class="col-md-4"><label class="form-label small">Red inalámbrica</label><input type="number" min="0" class="form-control form-control-sm" data-field="red_inalambrica_cantidad" value="${area.red_inalambrica_cantidad ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">Ilum. funcionando</label><input type="number" min="0" class="form-control form-control-sm" data-field="iluminacion_funcionan" value="${area.iluminacion_funcionan ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">Ilum. no funcionando</label><input type="number" min="0" class="form-control form-control-sm" data-field="iluminacion_no_funcionan" value="${area.iluminacion_no_funcionan ?? ""}"></div>
					<div class="col-md-3"><label class="form-label small">Luminarias</label><input type="number" min="0" class="form-control form-control-sm" data-field="luminarias_cantidad" value="${area.luminarias_cantidad ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">Puntos eléctricos</label><select class="form-select form-select-sm" data-field="puntos_electricos">${makeOptions(options.yesNo, area.puntos_electricos)}</select></div>

					<div class="col-md-5"><label class="form-label small">Ptos eléctricos func.</label><input type="number" min="0" class="form-control form-control-sm" data-field="puntos_electricos_funcionan" value="${area.puntos_electricos_funcionan ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">Ptos eléctricos no func.</label><input type="number" min="0" class="form-control form-control-sm" data-field="puntos_electricos_no_funcionan" value="${area.puntos_electricos_no_funcionan ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">Ptos eléctricos total</label><input type="number" min="0" class="form-control form-control-sm" data-field="puntos_electricos_cantidad" value="${area.puntos_electricos_cantidad ?? ""}"></div>
					<div class="col-md-4"><label class="form-label small">Capacidad áulica</label><input type="number" min="0" class="form-control form-control-sm" data-field="capacidad_aulica" value="${area.capacidad_aulica ?? ""}"></div>
					<div class="col-md-6"><label class="form-label small">Capacidad distanciamiento</label><input type="number" min="0" class="form-control form-control-sm" data-field="capacidad_distanciamiento" value="${area.capacidad_distanciamiento ?? ""}"></div>
					<div class="col-md-6"><label class="form-label small">Apto retorno presencial</label><select class="form-select form-select-sm" data-field="ambiente_apto_retorno">${makeOptions(options.yesNo, area.ambiente_apto_retorno)}</select></div>

					<div class="col-12"><label class="form-label small">Observaciones</label><textarea rows="2" class="form-control form-control-sm" data-field="observaciones_detalle">${escapeHtml(area.observaciones_detalle || "")}</textarea></div>
				</div>
			`;

			const responsableSelect = detalleUbicacionBody.querySelector('[data-field="responsable_admin_id"]');
			const addResponsableBtn = detalleUbicacionBody.querySelector('[data-action="add-responsable"]');

			const handleQuickCreateResponsable = async () => {
				const nombre = window.prompt("Nombre del responsable:");
				if (nombre === null) {
					if (responsableSelect && responsableSelect.value === ADD_ADMIN_OPTION) {
						responsableSelect.value = String(area.responsable_admin_id || "");
					}
					return;
				}
				const trimmedName = nombre.trim();
				if (!trimmedName) {
					notify("El nombre es obligatorio.", true);
					if (responsableSelect && responsableSelect.value === ADD_ADMIN_OPTION) {
						responsableSelect.value = String(area.responsable_admin_id || "");
					}
					return;
				}

				try {
					const created = await api.send("/api/administradores", "POST", {
						nombre: trimmedName,
						cargo: "",
						facultad: "",
						titulo_academico: "",
						email: null,
						telefono: null,
					});
					const newAdminId = Number(created?.id);
					const adminsRes = await api.get("/api/administradores");
					const admins = adminsRes.data || [];
					if (responsableSelect) {
						responsableSelect.innerHTML = buildAdminOptions(admins, Number.isFinite(newAdminId) ? newAdminId : "");
						if (Number.isFinite(newAdminId)) {
							responsableSelect.value = String(newAdminId);
						}
					}
					await loadAdministradores();
					notify("Responsable agregado correctamente.");
				} catch (error) {
					notify(error.message, true);
					if (responsableSelect && responsableSelect.value === ADD_ADMIN_OPTION) {
						responsableSelect.value = String(area.responsable_admin_id || "");
					}
				}
			};

			addResponsableBtn?.addEventListener("click", handleQuickCreateResponsable);
			responsableSelect?.addEventListener("change", async () => {
				if (responsableSelect.value !== ADD_ADMIN_OPTION) return;
				await handleQuickCreateResponsable();
			});
			setDetailButtonsMode("edit", true);
		}

		function collectAreaEditFormValues() {
			const form = detalleUbicacionBody.querySelector("#detalle-area-edit-form");
			if (!form) return null;
			const values = {};
			form.querySelectorAll("[data-field]").forEach((input) => {
				const field = input.dataset.field;
				let value = input.value;
				if (value === "") value = null;
				values[field] = value;
			});

			const numberFields = [
				"alto",
				"pupitres_cantidad", "pupitres_funcionan", "pupitres_no_funcionan",
				"ventanas_cantidad", "ventanas_funcionan", "ventanas_no_funcionan",
				"aa_cantidad", "aa_funcionan", "aa_no_funcionan", "ventiladores_cantidad",
				"ventiladores_funcionan", "ventiladores_no_funcionan", "red_lan_funcionan",
				"red_lan_no_funcionan", "red_inalambrica_cantidad", "iluminacion_funcionan",
				"iluminacion_no_funcionan", "luminarias_cantidad", "puntos_electricos_funcionan",
				"puntos_electricos_no_funcionan", "puntos_electricos_cantidad", "capacidad_aulica",
				"capacidad_distanciamiento",
				"responsable_admin_id",
			];
			numberFields.forEach((field) => {
				if (field in values) {
					values[field] = toNumberOrNull(values[field]);
				}
			});

			if (typeof values.nombre === "string") {
				values.nombre = values.nombre.trim();
			}
			if (typeof values.descripcion === "string") {
				values.descripcion = values.descripcion.trim();
			}
			if (typeof values.observaciones_detalle === "string") {
				values.observaciones_detalle = values.observaciones_detalle.trim();
			}
			return values;
		}

		function openFloorDetail(block, piso) {
			const areas = piso.areas || [];
			const names = areas.map((entry) => entry.nombre).join(", ") || "Sin áreas registradas";
			areaDetailContext = null;
			setLocationDetailContext("piso", piso.id, piso.nombre);
			showLocationDetail(
				`Detalle del piso: ${piso.nombre}`,
				`
				<div class="row g-2">
					<div class="col-12"><strong>Bloque:</strong> ${escapeHtml(block.nombre)}</div>
					<div class="col-12"><strong>Piso:</strong> ${escapeHtml(piso.nombre)}</div>
					<div class="col-12"><strong>Descripción:</strong> ${escapeHtml(piso.descripcion || "Sin descripción")}</div>
					<div class="col-12"><strong>Total de áreas:</strong> ${areas.length}</div>
					<div class="col-12"><strong>Áreas:</strong> ${escapeHtml(names)}</div>
				</div>
				`
			);
			setDetailButtonsMode("view", true);
		}

		function openBlockDetail(block) {
			const pisos = block.pisos || [];
			const totalAreas = pisos.reduce((acc, piso) => acc + ((piso.areas || []).length), 0);
			const pisosNames = pisos.map((entry) => entry.nombre).join(", ") || "Sin pisos registrados";
			areaDetailContext = null;
			setLocationDetailContext("bloque", block.id, block.nombre);
			showLocationDetail(
				`Detalle del bloque: ${block.nombre}`,
				`
				<div class="row g-2">
					<div class="col-12"><strong>Bloque:</strong> ${escapeHtml(block.nombre)}</div>
					<div class="col-12"><strong>Descripción:</strong> ${escapeHtml(block.descripcion || "Sin descripción")}</div>
					<div class="col-12"><strong>Total de pisos:</strong> ${pisos.length}</div>
					<div class="col-12"><strong>Total de áreas:</strong> ${totalAreas}</div>
					<div class="col-12"><strong>Pisos:</strong> ${escapeHtml(pisosNames)}</div>
				</div>
				`
			);
			setDetailButtonsMode("view", true);
		}

		async function openAreaDetail(block, piso, area) {
			areaDetailContext = {
				blockId: block.id,
				pisoId: piso.id,
				areaId: area.id,
			};
			setLocationDetailContext("area", area.id, area.nombre);
			try {
				const options = await loadAreaDetailOptions();
				renderAreaDetailReadOnly({ block, piso, area }, options);
			} catch (error) {
				console.error("Error al cargar opciones:", error);
				renderAreaDetailReadOnly({ block, piso, area });
			}
		}

		if (btnDetalleEliminar) {
			btnDetalleEliminar.addEventListener("click", async () => {
				if (!locationDetailContext) return;
				try {
					const accepted = await askDeleteLocationWithImpact(locationDetailContext);
					if (!accepted) return;
					await deleteLocationByContext(locationDetailContext);
					detalleUbicacionModal.hide();
				} catch (error) {
					notify(error.message, true);
				}
			});
		}

		if (btnDetalleEditar) {
			btnDetalleEditar.addEventListener("click", async () => {
				if (!areaDetailContext) {
					if (!locationDetailContext) return;
					if (locationDetailContext.type === "bloque") {
						const block = state.structure.find((entry) => String(entry.id) === String(locationDetailContext.id));
						if (!block) return;
						detalleUbicacionModal.hide();
						openLocationModal("bloque", block);
						return;
					}
					if (locationDetailContext.type === "piso") {
						const block = state.structure.find((entry) =>
							(entry.pisos || []).some((piso) => String(piso.id) === String(locationDetailContext.id))
						);
						const piso = (block?.pisos || []).find((entry) => String(entry.id) === String(locationDetailContext.id));
						if (!piso) return;
						detalleUbicacionModal.hide();
						openLocationModal("piso", piso);
						return;
					}
					return;
				}
				const current = findAreaInStructure(areaDetailContext.blockId, areaDetailContext.pisoId, areaDetailContext.areaId);
				if (!current) {
					notify("No se encontró el área para edición.", true);
					return;
				}
				try {
					const options = await loadAreaDetailOptions();
					renderAreaDetailEditForm(current, options);
				} catch (error) {
					notify(error.message, true);
				}
			});
		}

		if (btnDetalleCancelarEdicion) {
			btnDetalleCancelarEdicion.addEventListener("click", async () => {
				if (!areaDetailContext) {
					setDetailButtonsMode("view", false);
					return;
				}
				const current = findAreaInStructure(areaDetailContext.blockId, areaDetailContext.pisoId, areaDetailContext.areaId);
				if (!current) return;
				try {
					const options = await loadAreaDetailOptions();
					renderAreaDetailReadOnly(current, options);
				} catch (error) {
					console.error("Error al cargar opciones:", error);
					renderAreaDetailReadOnly(current);
				}
			});
		}

		if (btnDetalleGuardar) {
			btnDetalleGuardar.addEventListener("click", async () => {
				if (!areaDetailContext) return;
				const current = findAreaInStructure(areaDetailContext.blockId, areaDetailContext.pisoId, areaDetailContext.areaId);
				if (!current) {
					notify("No se encontró el área para guardar cambios.", true);
					return;
				}

				const edited = collectAreaEditFormValues();
				if (!edited) return;
				if (!edited.nombre) {
					notify("El nombre del área es obligatorio.", true);
					return;
				}

				const fieldsToCompare = [
					"nombre", "descripcion", "identificacion_ambiente", "metros_cuadrados", "alto", "senaletica",
					"cod_senaletica", "infraestructura_fisica", "estado_piso", "material_techo", "puerta", "material_puerta",
					"responsable_admin_id", "estado_paredes", "estado_techo", "estado_puerta",
					"cerradura", "nivel_seguridad", "sitio_profesor_mesa", "sitio_profesor_silla", "pc_aula", "proyector",
					"pantalla_interactiva", "pupitres_cantidad", "pupitres_funcionan", "pupitres_no_funcionan",
					"pizarra", "pizarra_estado", "ventanas_cantidad", "ventanas_funcionan", "ventanas_no_funcionan",
					"aa_cantidad", "aa_funcionan", "aa_no_funcionan", "ventiladores_cantidad",
					"ventiladores_funcionan", "ventiladores_no_funcionan", "wifi", "red_lan",
					"red_lan_funcionan", "red_lan_no_funcionan", "red_inalambrica_cantidad", "iluminacion_funcionan",
					"iluminacion_no_funcionan", "luminarias_cantidad", "puntos_electricos", "puntos_electricos_funcionan",
					"puntos_electricos_no_funcionan", "puntos_electricos_cantidad", "capacidad_aulica",
					"capacidad_distanciamiento", "ambiente_apto_retorno", "observaciones_detalle",
				];

				const normalize = (value) => (value === "" || value === undefined ? null : value);
				const patchPayload = {};
				fieldsToCompare.forEach((field) => {
					const oldValue = normalize(current.area[field]);
					const newValue = normalize(edited[field]);
					if (String(oldValue ?? "") !== String(newValue ?? "")) {
						patchPayload[field] = newValue;
					}
				});

				if (!Object.keys(patchPayload).length) {
					notify("No hay cambios para guardar.");
					try {
						const options = await loadAreaDetailOptions();
						const refreshedNoChange = findAreaInStructure(areaDetailContext.blockId, areaDetailContext.pisoId, areaDetailContext.areaId);
						if (refreshedNoChange) renderAreaDetailReadOnly(refreshedNoChange, options);
					} catch (e) {
						console.error(e);
					}
					return;
				}

				try {
					await api.send(`/api/areas/${areaDetailContext.areaId}`, "PATCH", patchPayload);
					await loadStructure();
					const refreshed = findAreaInStructure(areaDetailContext.blockId, areaDetailContext.pisoId, areaDetailContext.areaId);
					if (refreshed) {
						const options = await loadAreaDetailOptions();
						renderAreaDetailReadOnly(refreshed, options);
						notify("Área actualizada correctamente.");
					} else {
						detalleUbicacionModal.hide();
					}
				} catch (error) {
					notify(error.message, true);
				}
			});
		}

		if (detalleUbicacionModalElement) {
			detalleUbicacionModalElement.addEventListener("hidden.bs.modal", () => {
				areaDetailContext = null;
				locationDetailContext = null;
				setDetailButtonsMode("view", false);
			});
		}

		function openFloorEditModal(piso) {
			openLocationModal("piso", piso);
		}

		function renderBlocks() {
			bloquesTabs.innerHTML = "";
			state.structure.forEach((block) => {
				const wrapper = document.createElement("div");
				wrapper.className = "btn-group";
				wrapper.setAttribute("role", "group");

				const btnTab = document.createElement("button");
				btnTab.type = "button";
				btnTab.className = `btn ${block.id === state.activeBlockId ? "btn-primary" : "btn-outline-secondary"}`;
				btnTab.textContent = block.nombre;
				if (block.descripcion) btnTab.title = block.descripcion;
				
				btnTab.addEventListener("click", () => {
					state.activeBlockId = block.id;
					state.activeFloorId = block.pisos[0]?.id || null;
					renderAll();
				});

				const btnView = document.createElement("button");
				btnView.type = "button";
				btnView.className = `btn ${block.id === state.activeBlockId ? "btn-primary" : "btn-outline-secondary"} px-2 d-none`;
				btnView.innerHTML = '<i class="bi bi-eye"></i>';
				btnView.title = "Ver detalles del bloque";
				btnView.addEventListener("click", (e) => {
					e.stopPropagation();
					openBlockDetail(block);
				});

				wrapper.addEventListener("mouseenter", () => btnView.classList.remove("d-none"));
				wrapper.addEventListener("mouseleave", () => btnView.classList.add("d-none"));

				wrapper.appendChild(btnView);
				wrapper.appendChild(btnTab);
				bloquesTabs.appendChild(wrapper);
			});
		}

		function renderFloors() {
			const block = getActiveBlock();
			pisosContainer.innerHTML = "";
			areasContainer.innerHTML = "";
			if (!block) {
				pisosContainer.innerHTML = '<p class="text-muted mb-0">Agrega un bloque para comenzar.</p>';
				areasSection.classList.add("d-none");
				return;
			}
			if (!block.pisos.length) {
				pisosContainer.innerHTML = '<p class="text-muted mb-0">Este bloque aún no tiene pisos.</p>';
				areasSection.classList.add("d-none");
				return;
			}
			areasSection.classList.add("d-none");
			
			block.pisos.forEach((piso) => {
				const floorCard = document.createElement("div");
				// Determinar si este piso debe estar expandido (basado en el estado)
				const isExpanded = state.expandedFloorId === piso.id;
				
				floorCard.className = "floor-compact-card";
				const areaNames = (piso.areas || []).map((entry) => entry.nombre).join(", ");
				const summaryText = areaNames
					? `Total áreas: ${piso.areas.length}. Registradas: ${areaNames}`
					: "Total áreas: 0. No hay áreas registradas en este piso.";

				const areaButtonsHtml = (piso.areas || []).length
					? piso.areas
							.map(
								(area) => `
									<button type="button" class="area-pill-btn" data-area-id="${area.id}" title="${escapeHtml(area.descripcion || area.nombre)}">
										${escapeHtml(area.nombre)}
									</button>
								`
							)
							.join("")
					: '<p class="text-muted small mb-0">Sin áreas registradas</p>';

				floorCard.innerHTML = `
					<div class="floor-compact-head">
						<div class="floor-title-wrap">
							<span class="floor-title-text ${isExpanded ? 'text-primary' : ''}" title="${escapeHtml(piso.descripcion || piso.nombre)}">${escapeHtml(piso.nombre)}</span>
						</div>
						<div class="floor-actions">
							<button type="button" class="icon-action-btn floor-edit-btn" title="Editar piso"><i class="bi bi-pencil"></i></button>
							<button type="button" class="icon-action-btn floor-delete-btn" title="Eliminar piso"><i class="bi bi-trash"></i></button>
							<button type="button" class="icon-action-btn floor-view-btn" title="Ver detalles del piso"><i class="bi bi-eye"></i></button>
						</div>
					</div>
					<div class="floor-collapse ${isExpanded ? '' : 'd-none'}">
						<div class="floor-collapse-toolbar">
							<button type="button" class="btn btn-outline-primary btn-sm add-area-floor-btn" data-floor-id="${piso.id}">
								<i class="bi bi-plus-lg me-1"></i>Agregar Área
							</button>
						</div>
						<div class="floor-area-pills">${areaButtonsHtml}</div>
						<div class="floor-summary-note">${escapeHtml(summaryText)}</div>
					</div>
				`;
				pisosContainer.appendChild(floorCard);

				const titleText = floorCard.querySelector(".floor-title-text");
				const floorHead = floorCard.querySelector(".floor-compact-head");
				const viewBtn = floorCard.querySelector(".floor-view-btn");
				const editBtn = floorCard.querySelector(".floor-edit-btn");
				const deleteBtn = floorCard.querySelector(".floor-delete-btn");
				const collapse = floorCard.querySelector(".floor-collapse");
				const addAreaBtn = floorCard.querySelector(".add-area-floor-btn");

				const toggleCollapse = () => {
					const hidden = collapse.classList.toggle("d-none");
					titleText.classList.toggle("text-primary", !hidden);
					// Actualizamos el estado para recordar qué piso está abierto
					if (!hidden) {
						state.expandedFloorId = piso.id;
					} else if (state.expandedFloorId === piso.id) {
						state.expandedFloorId = null;
					}
				};
				const isActionClick = (event) =>
					Boolean(event.target.closest(".floor-edit-btn, .floor-delete-btn, .floor-view-btn"));

				floorHead?.addEventListener("click", (event) => {
					if (isActionClick(event)) return;
					toggleCollapse();
				});
				
				viewBtn.addEventListener("click", () => openFloorDetail(block, piso));
				editBtn.addEventListener("click", () => openLocationModal("piso", piso));
				deleteBtn.addEventListener("click", async () => {
					try {
						const context = { type: "piso", id: piso.id, name: piso.nombre };
						const accepted = await askDeleteLocationWithImpact(context);
						if (!accepted) return;
						await deleteLocationByContext(context);
					} catch (error) {
						notify(error.message, true);
					}
				});

				addAreaBtn.addEventListener("click", () => {
					state.activeFloorId = piso.id;
					openLocationModal("area");
				});

				floorCard.querySelectorAll(".area-pill-btn").forEach((areaBtn) => {
					const area = (piso.areas || []).find((entry) => String(entry.id) === areaBtn.dataset.areaId);
					if (!area) return;
					areaBtn.addEventListener("dblclick", () => openAreaDetail(block, piso, area));
				});
			});
		}

		function renderAll() {
			renderBlocks();
			
			// Mostrar/ocultar sección de pisos y áreas basado en si hay bloques
			if (state.structure && state.structure.length > 0) {
				sinBloquesMsg.classList.add("d-none");
				bloquesContent.classList.remove("d-none");
				renderFloors();
			} else {
				sinBloquesMsg.classList.remove("d-none");
				bloquesContent.classList.add("d-none");
			}
		}

		function openLocationModal(mode, dataToEdit = null) {
			modalMode = mode;
			locationEditContext = null; // Limpiar contexto previo
			
			if (mode === "area") {
				if (!state.activeFloorId) {
					notify("Primero debes seleccionar un piso.", true);
					return;
				}
				resetAreaForm();
				loadAreaFormOptions();
				if (modalArea) {
					modalArea.show();
				}
				return;
			}
			
			const mapTitle = {
				bloque: "Agregar bloque",
				"bloque-edit": "Editar bloque",
				piso: "Agregar piso",
				"piso-edit": "Editar piso",
			};
			
			// Si es edición, guardar contexto y rellenar formulario
			if (dataToEdit) {
				if (mode === "bloque") {
					mode = "bloque-edit";
					modalMode = mode;
					locationEditContext = { type: "bloque", id: dataToEdit.id };
					inputNombre.value = dataToEdit.nombre || "";
					inputDescripcion.value = dataToEdit.descripcion || "";
				} else if (mode === "piso") {
					mode = "piso-edit";
					modalMode = mode;
					locationEditContext = { type: "piso", id: dataToEdit.id };
					inputNombre.value = dataToEdit.nombre || "";
					inputDescripcion.value = dataToEdit.descripcion || "";
				}
			} else {
				inputNombre.value = "";
				inputDescripcion.value = "";
			}
			
			modalTitle.textContent = mapTitle[mode] || "Nueva entidad";
			modal.show();
		}

		async function loadStructure() {
			state.structure = await window.appHelpers.loadStructure(api, { sortNatural: true });

			if (!state.activeBlockId && state.structure.length) {
				state.activeBlockId = state.structure[0].id;
				state.activeFloorId = state.structure[0].pisos[0]?.id || null;
			}
			const activeBlockExists = state.structure.some((block) => block.id === state.activeBlockId);
			if (!activeBlockExists) {
				state.activeBlockId = state.structure[0]?.id || null;
				state.activeFloorId = state.structure[0]?.pisos[0]?.id || null;
			}
			const activeBlock = getActiveBlock();
			if (activeBlock && !activeBlock.pisos.some((piso) => piso.id === state.activeFloorId)) {
				state.activeFloorId = activeBlock.pisos[0]?.id || null;
			}
			renderAll();
		}

		btnAgregarBloque.addEventListener("click", () => openLocationModal("bloque"));
		btnExportarUbicaciones?.addEventListener("click", () => {
			window.location.href = "/api/ubicaciones/export";
		});

		btnImportarUbicaciones?.addEventListener("click", () => {
			if (importResumen) importResumen.textContent = "";
			if (importTextarea) importTextarea.value = "";
			importModal?.show();
		});

		importTextarea?.addEventListener("input", () => {
			if (!importResumen) return;
			const lines = String(importTextarea.value || "").split(/\r?\n/);
			const summary = summarizeImportLines(lines);
			importResumen.textContent = `Líneas válidas: ${summary.validRows} | Bloques únicos: ${summary.uniqueBlocks} | Pisos únicos: ${summary.uniqueFloors}`;
		});

		btnConfirmarImport?.addEventListener("click", async () => {
			const source = importTextarea?.value || "";
			if (!source.trim()) {
				notify("Pega primero las líneas de ubicación.", true);
				return;
			}
			try {
				const result = await importBlockFloorLocationsFromText(source);
				importModal?.hide();
				notify(`Importación completada. Bloques creados: ${result.createdBlocks}. Pisos creados: ${result.createdFloors}. Filas leídas: ${result.rowsRead}.`);
			} catch (error) {
				notify(error.message, true);
			}
		});

		btnAgregarPiso.addEventListener("click", () => {
			if (!state.activeBlockId) {
				notify("Primero debes crear o seleccionar un bloque.", true);
				return;
			}
			openLocationModal("piso");
		});

		// Handler for saving locations or parameters
		const handleModalSave = async () => {
			const nombre = inputNombre.value.trim();
			const descripcion = inputDescripcion.value.trim();

			if (!nombre) {
				notify("El nombre es obligatorio.", true);
				return;
			}

			try {
				// Location modes - Creación o edición
				if (modalMode === "bloque") {
					await api.send("/api/bloques", "POST", { nombre, descripcion });
				} else if (modalMode === "bloque-edit" && locationEditContext) {
					await api.send(`/api/bloques/${locationEditContext.id}`, "PATCH", { nombre, descripcion });
				} else if (modalMode === "piso") {
					await api.send("/api/pisos", "POST", {
						bloque_id: state.activeBlockId,
						nombre,
						descripcion,
					});
				} else if (modalMode === "piso-edit" && locationEditContext) {
					await api.send(`/api/pisos/${locationEditContext.id}`, "PATCH", { nombre, descripcion });
				}
				// Parameter modes
				if (["estados", "condiciones", "cuentas", "si_no", "estado_puerta", "cerraduras", "estado_piso", "material_techo", "material_puerta", "estado_pizarra"].includes(modalMode)) {
					await api.send(`/api/parametros/${modalMode}`, "POST", {
						nombre,
						descripcion: descripcion || null,
					});
				}

				modal.hide();
				
				// Reload appropriate data
				if (["bloque", "bloque-edit", "piso", "piso-edit"].includes(modalMode)) {
					await loadStructure();
				} else if (["estados", "condiciones", "cuentas", "si_no", "estado_puerta", "cerraduras", "estado_piso", "material_techo", "material_puerta", "estado_pizarra"].includes(modalMode)) {
					await loadParametros();
				}
			} catch (error) {
				notify(error.message, true);
			}
		};

		if (btnGuardarAreaDetalle) {
			btnGuardarAreaDetalle.addEventListener("click", async () => {
				const nombreArea = (areaFields.nombre?.value || "").trim();

				const numericKeys = [
					"alto",
					"pupitres_cantidad", "pupitres_funcionan", "pupitres_no_funcionan",
					"ventanas_cantidad", "ventanas_funcionan", "ventanas_no_funcionan",
					"aa_cantidad", "aa_funcionan", "aa_no_funcionan", "ventiladores_cantidad",
					"ventiladores_funcionan", "ventiladores_no_funcionan", "red_lan_funcionan",
					"red_lan_no_funcionan", "red_inalambrica_cantidad", "iluminacion_funcionan",
					"iluminacion_no_funcionan", "luminarias_cantidad", "puntos_electricos_funcionan",
					"puntos_electricos_no_funcionan", "puntos_electricos_cantidad", "capacidad_aulica",
					"capacidad_distanciamiento",
				];

				const payload = {
					piso_id: state.activeFloorId,
					descripcion: areaFields.descripcion?.value?.trim() || null,
					identificacion_ambiente: areaFields.identificacion_ambiente?.value?.trim() || null,
					metros_cuadrados: areaFields.metros_cuadrados?.value?.trim() || null,
					senaletica: areaFields.senaletica?.value || null,
					cod_senaletica: areaFields.cod_senaletica?.value?.trim() || null,
					infraestructura_fisica: areaFields.infraestructura_fisica?.value?.trim() || null,
					estado_piso: areaFields.estado_piso?.value || null,
					material_techo: areaFields.material_techo?.value || null,
					puerta: areaFields.puerta?.value || null,
					material_puerta: areaFields.material_puerta?.value || null,
					responsable_admin_id: areaFields.responsable_admin_id?.value ? Number(areaFields.responsable_admin_id.value) : null,
					estado_paredes: areaFields.estado_paredes?.value || null,
					estado_techo: areaFields.estado_techo?.value || null,
					estado_puerta: areaFields.estado_puerta?.value || null,
					cerradura: areaFields.cerradura?.value || null,
					nivel_seguridad: areaFields.nivel_seguridad?.value || null,
					sitio_profesor_mesa: areaFields.sitio_profesor_mesa?.value || null,
					sitio_profesor_silla: areaFields.sitio_profesor_silla?.value || null,
					pc_aula: areaFields.pc_aula?.value || null,
					proyector: areaFields.proyector?.value || null,
					pantalla_interactiva: areaFields.pantalla_interactiva?.value || null,
					pizarra: areaFields.pizarra?.value || null,
					pizarra_estado: areaFields.pizarra_estado?.value || null,
					wifi: areaFields.wifi?.value || null,
					red_lan: areaFields.red_lan?.value || null,
					puntos_electricos: areaFields.puntos_electricos?.value || null,
					ambiente_apto_retorno: areaFields.ambiente_apto_retorno?.value || null,
					observaciones_detalle: areaFields.observaciones_detalle?.value?.trim() || null,
				};

				numericKeys.forEach((key) => {
					const value = areaFields[key]?.value;
					payload[key] = value === "" || value === undefined ? null : Number(value);
				});

				if (nombreArea) {
					payload.nombre = nombreArea;
				}

				try {
					await api.send("/api/areas", "POST", payload);
					if (modalArea) modalArea.hide();
					await loadStructure();
					notify("Área registrada correctamente.");
				} catch (error) {
					notify(error.message, true);
				}
			});
		}

		btnGuardar.removeEventListener("click", handleModalSave);
		btnGuardar.addEventListener("click", handleModalSave);

		try {
			await loadStructure();
		} catch (error) {
			notify(error.message, true);
		}
	}

	const parametrosModule = typeof window.initConfigParametrosSection === "function"
		? window.initConfigParametrosSection({
			modal,
			modalTitle,
			inputNombre,
			inputDescripcion,
			setModalMode: (mode) => {
				modalMode = mode;
			},
			registerLoadParametros: (loader) => {
				if (typeof loader === "function") {
					loadParametros = loader;
				}
			},
		})
		: null;

	const personalModule = typeof window.initConfigPersonalSection === "function"
		? window.initConfigPersonalSection()
		: null;

	try {
		if (typeof parametrosModule?.loadParametros === "function") {
			await parametrosModule.loadParametros();
		}
		if (typeof personalModule?.loadAdministradores === "function") {
			await personalModule.loadAdministradores();
		}
	} catch (error) {
		console.error("Error initializing parametros:", error);
	}
}

document.addEventListener("DOMContentLoaded", async () => {
	if (typeof window.bindSettingsNumericInputGuards === "function") {
		window.bindSettingsNumericInputGuards();
	}

	await initSettingsPage();
});
