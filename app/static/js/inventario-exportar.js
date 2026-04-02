(function () {
	window.bindInventarioExport = function bindInventarioExport(options) {
		const { button, getCurrentFilterParams } = options || {};
		if (!button || typeof getCurrentFilterParams !== "function") return;
		button.addEventListener("click", () => {
			const params = getCurrentFilterParams();
			window.location.href = `/api/inventario/export?${params.toString()}`;
		});
	};
})();
