(function () {
	window.createDuplicateInventoryModal = function createDuplicateInventoryModal(options) {
		const {
			duplicateModal,
			nodes,
			escapeHtmlText,
			buildDuplicateWarningMessage,
		} = options || {};

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
									<div class="fw-semibold">ítem #${escapeHtmlText(item.item_numero || "-")}</div>
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

		return {
			openDuplicateModal,
		};
	};
})();
