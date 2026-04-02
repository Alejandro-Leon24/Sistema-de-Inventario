(function () {
	window.createInventarioTablaController = function createInventarioTablaController(options) {
		const {
			state,
			nodes,
			body,
			getOrderedColumns,
			formatValue,
			escapeHtmlText,
			applyColumnWidthToField,
		} = options || {};

		function renderTableHead() {
			const columns = getOrderedColumns();
			nodes.tableHeadRow.innerHTML = "";
			columns.forEach((column) => {
				const th = document.createElement("th");
				th.innerHTML = `<span class="head-label">${escapeHtmlText(column.label)}</span><span class="column-resize-handle" title="Ajustar ancho"></span>`;
				th.dataset.field = column.field;
				th.draggable = true;
				th.classList.add("inventory-head-cell");
				nodes.tableHeadRow.appendChild(th);
				if (state.columnWidths[column.field]) {
					applyColumnWidthToField(column.field, state.columnWidths[column.field]);
				}
			});
		}

		function renderRows() {
			const columns = getOrderedColumns();
			body.innerHTML = "";
			state.items.forEach((item, index) => {
				const tr = document.createElement("tr");
				tr.dataset.id = item.id;
				columns.forEach((column) => {
					const td = document.createElement("td");
					td.dataset.field = column.field;
					td.dataset.id = item.id;

					let displayValue;
					if (column.field === "item_numero") {
						displayValue = (state.page - 1) * state.perPage + index + 1;
					} else {
						displayValue = formatValue(column.field, item[column.field]);
					}

					td.textContent = displayValue;
					td.title = String(displayValue || "");
					td.classList.add("inventory-cell");
					if (column.editable) td.classList.add("editable-cell");
					if (column.field === "item_numero") td.classList.add("fw-bold", "text-primary");
					if (state.columnWidths[column.field]) {
						const px = `${state.columnWidths[column.field]}px`;
						td.style.width = px;
						td.style.minWidth = px;
						td.style.maxWidth = px;
					}
					tr.appendChild(td);
				});
				body.appendChild(tr);
			});
		}

		function renderPaginationMeta() {
			if (!nodes.pageInfo || !nodes.pageIndicator || !nodes.pagePrev || !nodes.pageNext) return;
			const total = state.totalItems || 0;
			const currentPage = Math.max(state.page, 1);
			const perPage = Math.max(state.perPage, 1);
			const first = total === 0 ? 0 : (currentPage - 1) * perPage + 1;
			const last = total === 0 ? 0 : Math.min(currentPage * perPage, total);
			const totalPages = Math.max(state.totalPages || 1, 1);

			nodes.pageInfo.textContent = `Mostrando ${first} - ${last} de ${total}`;
			nodes.pageIndicator.textContent = `${currentPage} / ${totalPages}`;
			nodes.pagePrev.disabled = currentPage <= 1;
			nodes.pageNext.disabled = currentPage >= totalPages;
			if (nodes.pageSize) nodes.pageSize.value = String(perPage);
		}

		return {
			renderTableHead,
			renderRows,
			renderPaginationMeta,
		};
	};
})();
