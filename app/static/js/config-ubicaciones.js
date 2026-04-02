(function () {
	window.bindSettingsNumericInputGuards = function bindSettingsNumericInputGuards() {
		document.addEventListener("keydown", (event) => {
			if (event.target.tagName === "INPUT" && event.target.type === "number") {
				if (event.key === "-" || event.key === "e" || event.key === "E" || event.key === "+") {
					event.preventDefault();
				}
			}
		});

		document.addEventListener("input", (event) => {
			if (event.target.tagName === "INPUT" && event.target.type === "number") {
				if (event.target.value && parseFloat(event.target.value) < 0) {
					event.target.value = Math.abs(parseFloat(event.target.value));
				}
			}
		});
	};
})();
