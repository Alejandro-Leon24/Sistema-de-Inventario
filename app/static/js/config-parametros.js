(function () {
	const api = window.api;
	const escapeHtmlText = window.appHelpers.escapeHtmlText;

	function notifyMessage(message, isError = false) {
		if (typeof window.notify === "function") {
			window.notify(message, isError);
			return;
		}
		if (isError) {
			console.error(message);
			return;
		}
		console.log(message);
	}

	window.initConfigParametrosSection = function initConfigParametrosSection(options) {
		const {
			modal,
			modalTitle,
			inputNombre,
			inputDescripcion,
			setModalMode,
			registerLoadParametros,
		} = options || {};

		const btnAddEstado = document.getElementById("btn-add-estado");
		const btnAddCondicion = document.getElementById("btn-add-condicion");
		const btnAddCuenta = document.getElementById("btn-add-cuenta");
		const btnAddSiNo = document.getElementById("btn-add-si-no");
		const btnAddEstadoPuerta = document.getElementById("btn-add-estado-puerta");
		const btnAddCerradura = document.getElementById("btn-add-cerradura");
		const btnAddEstadoPiso = document.getElementById("btn-add-estado-piso");
		const btnAddMaterialTecho = document.getElementById("btn-add-material-techo");
		const btnAddMaterialPuerta = document.getElementById("btn-add-material-puerta");
		const btnAddEstadoPizarra = document.getElementById("btn-add-estado-pizarra");
		const btnSaveUniversidad = document.getElementById("btn-save-universidad");
		const btnEditUniversidad = document.getElementById("btn-edit-universidad");
		const btnCancelUniversidad = document.getElementById("btn-cancel-universidad");
		const inputUniversidad = document.getElementById("param-universidad-nombre");

		const listEstados = document.getElementById("param-estados-list");
		const listCondiciones = document.getElementById("param-condiciones-list");
		const listCuentas = document.getElementById("param-cuentas-list");
		const listSiNo = document.getElementById("param-si-no-list");
		const listEstadoPuerta = document.getElementById("param-estado-puerta-list");
		const listCerraduras = document.getElementById("param-cerraduras-list");
		const listEstadoPiso = document.getElementById("param-estado-piso-list");
		const listMaterialTecho = document.getElementById("param-material-techo-list");
		const listMaterialPuerta = document.getElementById("param-material-puerta-list");
		const listEstadoPizarra = document.getElementById("param-estado-pizarra-list");

		const universidadController = typeof window.initConfigUniversidadSection === "function"
			? window.initConfigUniversidadSection({
				btnSave: btnSaveUniversidad,
				btnEdit: btnEditUniversidad,
				btnCancel: btnCancelUniversidad,
				input: inputUniversidad,
			})
			: null;

		async function loadParametros() {
			try {
				const [estadosRes, condicionesRes, cuentasRes, siNoRes, estadoPuertaRes, cerradurasRes, estadoPisoRes, materialTechoRes, materialPuertaRes, estadoPizarraRes, universidadRes] = await Promise.all([
					api.get("/api/parametros/estados"),
					api.get("/api/parametros/condiciones"),
					api.get("/api/parametros/cuentas"),
					api.get("/api/parametros/si_no"),
					api.get("/api/parametros/estado_puerta"),
					api.get("/api/parametros/cerraduras"),
					api.get("/api/parametros/estado_piso"),
					api.get("/api/parametros/material_techo"),
					api.get("/api/parametros/material_puerta"),
					api.get("/api/parametros/estado_pizarra"),
					api.get("/api/universidad"),
				]);

				renderParamList(listEstados, estadosRes.data || [], "estados");
				renderParamList(listCondiciones, condicionesRes.data || [], "condiciones");
				renderParamList(listCuentas, cuentasRes.data || [], "cuentas");
				renderParamList(listSiNo, siNoRes.data || [], "si_no");
				renderParamList(listEstadoPuerta, estadoPuertaRes.data || [], "estado_puerta");
				renderParamList(listCerraduras, cerradurasRes.data || [], "cerraduras");
				renderParamList(listEstadoPiso, estadoPisoRes.data || [], "estado_piso");
				renderParamList(listMaterialTecho, materialTechoRes.data || [], "material_techo");
				renderParamList(listMaterialPuerta, materialPuertaRes.data || [], "material_puerta");
				renderParamList(listEstadoPizarra, estadoPizarraRes.data || [], "estado_pizarra");

				const universidadNombre = (universidadRes.data || {}).nombre_universidad || "";
				universidadController?.applyValue(universidadNombre);
			} catch (error) {
				console.error("Error loading parametros:", error);
			}
		}

		function renderParamList(container, items, paramType) {
			if (!container) return;
			container.innerHTML = "";
			if (!items.length) {
				container.innerHTML = '<p class="text-muted small mb-0">No hay elementos creados</p>';
				return;
			}
			items.forEach((item) => {
				const div = document.createElement("div");
				div.className = "list-group-item d-flex justify-content-between align-items-center";
				const canDelete = !!item.can_delete;
				div.innerHTML = `
					<div>
						<strong>${item.nombre}</strong>
						${item.descripcion ? `<div class="text-muted small">${item.descripcion}</div>` : ""}
					</div>
					<div class="d-flex gap-2">
						<button type="button" class="btn btn-sm btn-outline-primary param-edit" data-id="${item.id}" data-name="${escapeHtmlText(item.nombre)}" data-description="${escapeHtmlText(item.descripcion || "")}" data-type="${paramType}">Editar</button>
						${canDelete ? `<button type="button" class="btn btn-sm btn-outline-danger param-delete" data-id="${item.id}" data-type="${paramType}">Eliminar</button>` : ""}
					</div>
				`;
				container.appendChild(div);
			});

			container.querySelectorAll(".param-edit").forEach((btn) => {
				btn.addEventListener("click", async () => {
					const paramId = btn.dataset.id;
					const paramTypeInner = btn.dataset.type;
					const currentName = btn.dataset.name || "";
					const currentDescription = btn.dataset.description || "";
					const newName = window.prompt("Editar nombre del parámetro:", currentName);
					if (newName === null) return;
					const trimmed = newName.trim();
					if (!trimmed) {
						notifyMessage("El nombre es obligatorio.", true);
						return;
					}
					const newDescription = window.prompt("Editar descripción (opcional):", currentDescription) ?? currentDescription;
					try {
						await api.send(`/api/parametros/${paramTypeInner}/${paramId}`, "PATCH", {
							nombre: trimmed,
							descripcion: newDescription,
						});
						await loadParametros();
					} catch (error) {
						notifyMessage(error.message, true);
					}
				});
			});

			container.querySelectorAll(".param-delete").forEach((btn) => {
				btn.addEventListener("click", async () => {
					const paramId = btn.dataset.id;
					const paramType = btn.dataset.type;
					if (!window.confirm("¿Estás seguro de que deseas eliminar este parámetro?\nSi está siendo usado en alguna tabla, la eliminación será rechazada.")) {
						return;
					}

					try {
						await api.send(`/api/parametros/${paramType}/${paramId}`, "DELETE", {});
						await loadParametros();
					} catch (error) {
						notifyMessage(error.message, true);
					}
				});
			});
		}

		function openParamModal(tipo) {
			setModalMode(tipo);
			const mapTitle = {
				estados: "Agregar Estado",
				condiciones: "Agregar Condición",
				cuentas: "Agregar Cuenta",
				si_no: "Agregar opción Sí/No",
				estado_puerta: "Agregar estado de puerta",
				cerraduras: "Agregar tipo de cerradura",
				estado_piso: "Agregar estado de piso",
				material_techo: "Agregar material de techo",
				material_puerta: "Agregar material de puerta",
				estado_pizarra: "Agregar estado de pizarra",
			};
			if (modalTitle) modalTitle.textContent = mapTitle[tipo] || "Nuevo parámetro";
			if (inputNombre) inputNombre.value = "";
			if (inputDescripcion) inputDescripcion.value = "";
			modal?.show();
		}

		btnAddEstado?.addEventListener("click", () => openParamModal("estados"));
		btnAddCondicion?.addEventListener("click", () => openParamModal("condiciones"));
		btnAddCuenta?.addEventListener("click", () => openParamModal("cuentas"));
		btnAddSiNo?.addEventListener("click", () => openParamModal("si_no"));
		btnAddEstadoPuerta?.addEventListener("click", () => openParamModal("estado_puerta"));
		btnAddCerradura?.addEventListener("click", () => openParamModal("cerraduras"));
		btnAddEstadoPiso?.addEventListener("click", () => openParamModal("estado_piso"));
		btnAddMaterialTecho?.addEventListener("click", () => openParamModal("material_techo"));
		btnAddMaterialPuerta?.addEventListener("click", () => openParamModal("material_puerta"));
		btnAddEstadoPizarra?.addEventListener("click", () => openParamModal("estado_pizarra"));

		if (typeof registerLoadParametros === "function") {
			registerLoadParametros(loadParametros);
		}

		return { loadParametros };
	};
})();
