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

	window.initConfigPersonalSection = function initConfigPersonalSection() {
		const modalDeletePersonalConfirmEl = document.getElementById("modalConfirmarEliminacionPersonal");
		const modalDeletePersonalConfirm = modalDeletePersonalConfirmEl ? new bootstrap.Modal(modalDeletePersonalConfirmEl) : null;
		const modalDeletePersonalTitle = document.getElementById("modalConfirmarEliminacionPersonalLabel");
		const modalDeletePersonalBody = document.getElementById("modalConfirmarEliminacionPersonalBody");
		const btnConfirmDeletePersonal = document.getElementById("btn-confirmar-eliminacion-personal");

		const personalContainer = document.getElementById("administradores-container");
		const formPersonal = document.getElementById("form-personal");
		const inputPersonalNombre = document.getElementById("personal-nombre");
		const inputPersonalCargo = document.getElementById("personal-cargo");
		const inputPersonalFacultad = document.getElementById("personal-facultad");
		const inputPersonalTitulo = document.getElementById("personal-titulo");

		async function askDeletePersonalWithImpact(admin) {
			if (!admin || !modalDeletePersonalConfirm || !modalDeletePersonalBody || !btnConfirmDeletePersonal) {
				return false;
			}

			const response = await api.get(`/api/administradores/${admin.id}/impacto`);
			const impact = response.data || {};
			const items = Number(impact.items || 0);
			const areas = Number(impact.areas || 0);
			const hasRelations = items > 0 || areas > 0;

			modalDeletePersonalTitle.textContent = hasRelations
				? "Advertencia de eliminación en cascada (Personal)"
				: "Confirmar eliminación de personal";

			modalDeletePersonalBody.innerHTML = hasRelations
				? `
					<div class="alert alert-warning mb-2">
						<strong>Ya existen datos relacionados al personal.</strong>
					</div>
					<ul class="mb-2">
						<li>Áreas asociadas: <strong>${areas}</strong></li>
						<li>Ítems de inventario asignados: <strong>${items}</strong></li>
					</ul>
					<p class="mb-1">Si eliminas a <strong>${escapeHtmlText(admin.nombre)}</strong>, se perderá su asignación en las áreas y en el inventario (quedarán vacíos).</p>
					<p class="text-muted small mb-0">Recomendación: si no deseas afectar reportes existentes, edita sus datos en lugar de Eliminar.</p>
				`
				: `<p class="mb-0">¿Seguro que deseas eliminar a <strong>${escapeHtmlText(admin.nombre)}</strong>?</p>`;

			return new Promise((resolve) => {
				let resolved = false;
				const onConfirm = () => {
					if (resolved) return;
					resolved = true;
					cleanup();
					resolve(true);
					modalDeletePersonalConfirm.hide();
				};
				const onHidden = () => {
					if (resolved) return;
					resolved = true;
					cleanup();
					resolve(false);
				};
				const cleanup = () => {
					btnConfirmDeletePersonal.removeEventListener("click", onConfirm);
					modalDeletePersonalConfirmEl.removeEventListener("hidden.bs.modal", onHidden);
				};

				btnConfirmDeletePersonal.addEventListener("click", onConfirm);
				modalDeletePersonalConfirmEl.addEventListener("hidden.bs.modal", onHidden);
				modalDeletePersonalConfirm.show();
			});
		}

		function buildAdminPayloadFromInputs(inputs) {
			return {
				nombre: (inputs.nombre?.value || "").trim(),
				cargo: (inputs.cargo?.value || "").trim(),
				facultad: (inputs.facultad?.value || "").trim(),
				titulo_academico: (inputs.titulo?.value || "").trim(),
				email: null,
				telefono: null,
			};
		}

		function setAdminCardEditMode(card, editing) {
			card.querySelectorAll("[data-personal-field]").forEach((input) => {
				input.readOnly = !editing;
			});
			const btnEditar = card.querySelector(".btn-admin-edit");
			const btnGuardar = card.querySelector(".btn-admin-save");
			const btnCancelar = card.querySelector(".btn-admin-cancel");
			if (btnEditar) btnEditar.classList.toggle("d-none", editing);
			if (btnGuardar) btnGuardar.classList.toggle("d-none", !editing);
			if (btnCancelar) btnCancelar.classList.toggle("d-none", !editing);
		}

		function cacheOriginalAdminValues(card) {
			const cache = {};
			card.querySelectorAll("[data-personal-field]").forEach((input) => {
				cache[input.dataset.personalField] = input.value;
			});
			card.dataset.originalValues = JSON.stringify(cache);
		}

		function restoreOriginalAdminValues(card) {
			let cache = {};
			try {
				cache = JSON.parse(card.dataset.originalValues || "{}");
			} catch (_error) {
				cache = {};
			}
			card.querySelectorAll("[data-personal-field]").forEach((input) => {
				input.value = cache[input.dataset.personalField] || "";
			});
		}

		function renderAdministradores(admins) {
			if (!personalContainer) return;
			personalContainer.innerHTML = "";
			if (!admins.length) {
				personalContainer.innerHTML = '<div class="col-12"><p class="text-muted mb-0">No hay personal registrado.</p></div>';
				return;
			}

			admins.forEach((admin) => {
				const col = document.createElement("div");
				col.className = "col-md-6";
				col.innerHTML = `
					<div class="personal-card" data-admin-id="${admin.id}">
						<div class="personal-card-head d-flex justify-content-end">
							<div class="personal-actions">
								<button type="button" class="btn btn-link btn-sm p-0 btn-admin-edit" title="Editar"><i class="bi bi-pencil"></i></button>
								<button type="button" class="btn btn-link btn-sm p-0 text-success d-none btn-admin-save" title="Guardar"><i class="bi bi-check-lg"></i></button>
								<button type="button" class="btn btn-link btn-sm p-0 text-secondary d-none btn-admin-cancel" title="Cancelar"><i class="bi bi-x-lg"></i></button>
								<button type="button" class="btn btn-link btn-sm p-0 text-danger btn-admin-delete" title="Eliminar"><i class="bi bi-trash"></i></button>
							</div>
						</div>
						<div class="personal-inline-form row">
							<div class="col-12"> <strong>Nombre:</strong> <input class="form-control" data-personal-field="nombre" value="${escapeHtmlText(admin.nombre || "")}" readonly></div>
							<div class="col-12"> <strong>Cargo:</strong> <input class="form-control" data-personal-field="cargo" value="${escapeHtmlText(admin.cargo || "")}" readonly></div>
							<div class="col-12"> <strong>Facultad:</strong> <input class="form-control" data-personal-field="facultad" value="${escapeHtmlText(admin.facultad || "")}" readonly></div>
							<div class="col-12"> <strong>Título Académico:</strong> <input class="form-control" data-personal-field="titulo" value="${escapeHtmlText(admin.titulo_academico || "")}" readonly></div>
						</div>
					</div>
				`;
				personalContainer.appendChild(col);

				const card = col.querySelector(".personal-card");
				if (!card) return;
				cacheOriginalAdminValues(card);

				const btnEdit = card.querySelector(".btn-admin-edit");
				const btnSave = card.querySelector(".btn-admin-save");
				const btnCancel = card.querySelector(".btn-admin-cancel");
				const btnDelete = card.querySelector(".btn-admin-delete");

				btnEdit?.addEventListener("click", () => {
					cacheOriginalAdminValues(card);
					setAdminCardEditMode(card, true);
					card.querySelector("[data-personal-field='nombre']")?.focus();
				});

				btnCancel?.addEventListener("click", () => {
					restoreOriginalAdminValues(card);
					setAdminCardEditMode(card, false);
				});

				btnSave?.addEventListener("click", async () => {
					const payload = buildAdminPayloadFromInputs({
						nombre: card.querySelector("[data-personal-field='nombre']"),
						cargo: card.querySelector("[data-personal-field='cargo']"),
						facultad: card.querySelector("[data-personal-field='facultad']"),
						titulo: card.querySelector("[data-personal-field='titulo']"),
					});
					if (!payload.nombre) {
						notifyMessage("El nombre es obligatorio.", true);
						return;
					}
					try {
						await api.send(`/api/administradores/${admin.id}`, "PATCH", payload);
						await loadAdministradores();
					} catch (error) {
						notifyMessage(error.message, true);
					}
				});

				btnDelete?.addEventListener("click", async () => {
					try {
						const accepted = await askDeletePersonalWithImpact(admin);
						if (!accepted) return;

						await api.send(`/api/administradores/${admin.id}`, "DELETE", {});
						await loadAdministradores();
						notifyMessage("Registro eliminado.");
					} catch (error) {
						notifyMessage(error.message, true);
					}
				});
			});
		}

		async function loadAdministradores() {
			try {
				const response = await api.get("/api/administradores");
				renderAdministradores(response.data || []);
			} catch (_error) {
				if (personalContainer) {
					personalContainer.innerHTML = '<div class="col-12"><p class="text-danger mb-0">No se pudo cargar el personal.</p></div>';
				}
			}
		}

		formPersonal?.addEventListener("submit", async (event) => {
			event.preventDefault();
			const payload = buildAdminPayloadFromInputs({
				nombre: inputPersonalNombre,
				cargo: inputPersonalCargo,
				facultad: inputPersonalFacultad,
				titulo: inputPersonalTitulo,
			});
			if (!payload.nombre) {
				notifyMessage("El nombre es obligatorio.", true);
				return;
			}
			try {
				await api.send("/api/administradores", "POST", payload);
				formPersonal?.reset();
				await loadAdministradores();
				notifyMessage("Personal registrado correctamente.");
			} catch (error) {
				notifyMessage(error.message, true);
			}
		});

		return {
			loadAdministradores,
		};
	};
})();
