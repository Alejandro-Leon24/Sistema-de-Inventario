(function () {
	const api = window.api;

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

	window.initConfigUniversidadSection = function initConfigUniversidadSection(options) {
		const {
			btnSave,
			btnEdit,
			btnCancel,
			input,
		} = options || {};

		const state = {
			isEditing: false,
			originalValue: "",
		};

		function setMode(editing) {
			state.isEditing = editing;
			if (!input) return;
			input.readOnly = !editing;
			if (editing) {
				input.focus();
				input.select();
			}
			btnEdit?.classList.toggle("d-none", editing);
			btnSave?.classList.toggle("d-none", !editing);
			btnCancel?.classList.toggle("d-none", !editing);
		}

		function applyValue(value) {
			const finalValue = (value || "").trim();
			if (!input) return;
			input.value = finalValue;
			state.originalValue = finalValue;
			setMode(false);
		}

		btnEdit?.addEventListener("click", () => {
			state.originalValue = input?.value?.trim() || "";
			setMode(true);
		});

		btnCancel?.addEventListener("click", () => {
			if (input) input.value = state.originalValue;
			setMode(false);
		});

		btnSave?.addEventListener("click", async () => {
			const nombre = input?.value?.trim() || "";
			if (!nombre) {
				notifyMessage("El nombre de la universidad es obligatorio.", true);
				return;
			}
			try {
				await api.send("/api/universidad", "PATCH", { nombre_universidad: nombre });
				applyValue(nombre);
				notifyMessage("Universidad guardada correctamente.");
			} catch (error) {
				notifyMessage(error.message, true);
			}
		});

		input?.addEventListener("keydown", (event) => {
			if (!state.isEditing) return;
			if (event.key === "Escape") {
				event.preventDefault();
				btnCancel?.click();
			}
		});

		return {
			applyValue,
		};
	};
})();
