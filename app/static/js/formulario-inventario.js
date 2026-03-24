    document.addEventListener("DOMContentLoaded", function () {
        const form = document.getElementById("inventario-form");
        if (!form) return;

        const usuarioSelect = document.getElementById("usuario");
        if (usuarioSelect) {
            fetch("/api/administradores")
                .then((response) => response.ok ? response.json() : Promise.reject(new Error("No se pudo cargar personal")))
                .then((payload) => {
                    const admins = payload.data || [];
                    const previous = usuarioSelect.value;
                    usuarioSelect.innerHTML = '<option value="" selected disabled>Seleccione personal</option>';
                    admins.forEach((admin) => {
                        const option = document.createElement("option");
                        option.value = admin.nombre;
                        option.textContent = admin.nombre;
                        usuarioSelect.appendChild(option);
                    });
                    if (previous) usuarioSelect.value = previous;
                })
                .catch((error) => {
                    console.error(error.message);
                });
        }

        const fields = Array.from(
            form.querySelectorAll("input, select, textarea")
        ).filter((el) => {
            if (el.disabled || el.readOnly) return false;
            if (el.tabIndex < 0) return false;
            if (el.offsetParent === null) return false;
            const type = (el.type || "").toLowerCase();
            return type !== "hidden";
        });

        const isEmptyField = (field) => {
            if (field.tagName === "SELECT") {
                const selectedOption = field.options[field.selectedIndex];
                if (!selectedOption) return true;
                if (selectedOption.disabled) return true;
                return !field.value || field.value.trim() === "";
            }

            const type = (field.type || "").toLowerCase();

            if (type === "checkbox" || type === "radio") {
                return !field.checked;
            }

            return !field.value || field.value.trim() === "";
        };

        const setTouched = (field) => {
            field.dataset.touched = "true";
        };

        const updateBootstrapState = (field) => {
            const isTouched = field.dataset.touched === "true";
            if (!isTouched) return;

            const isEmpty = isEmptyField(field);
            field.classList.toggle("is-invalid", isEmpty);
            field.classList.toggle("is-valid", !isEmpty);
        };

        fields.forEach((field) => {
            field.addEventListener("focus", function () {
                setTouched(field);
            });

            field.addEventListener("blur", function () {
                setTouched(field);
                updateBootstrapState(field);
            });

            field.addEventListener("input", function () {
                updateBootstrapState(field);
            });

            field.addEventListener("change", function () {
                updateBootstrapState(field);
            });
        });

        form.addEventListener("keydown", function (event) {
            const isEnter = event.key === "Enter";
            const target = event.target;
            const isField = target.matches("input, select");

            if (!isEnter || !isField) return;

            event.preventDefault();

            const focusableFields = Array.from(
                form.querySelectorAll("input, select, textarea, button")
            ).filter((el) => {
                if (el.disabled || el.readOnly) return false;
                if (el.tabIndex < 0) return false;
                if (el.offsetParent === null) return false;
                if (el.tagName === "BUTTON") return false;
                const type = (el.type || "").toLowerCase();
                return type !== "hidden";
            });

            const currentIndex = focusableFields.indexOf(target);
            const nextField = focusableFields[currentIndex + 1];

            if (nextField) {
                nextField.focus();
                if (nextField.select) nextField.select();
            }
        });

        form.addEventListener("submit", function (event) {
            let hasInvalidFields = false;

            fields.forEach((field) => {
                setTouched(field);
                updateBootstrapState(field);

                if (field.classList.contains("is-invalid")) {
                    hasInvalidFields = true;
                }
            });

            if (hasInvalidFields) {
                event.preventDefault();
                event.stopPropagation();
            }
        });
    });