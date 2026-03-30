function notify(message, isError = false) {
    if (isError) {
        console.error(message);
    }

    const modalEl = document.getElementById("modalGlobalNotificacion");
    if (!modalEl) {
        window.alert(message);
        return;
    }

    const titleEl = document.getElementById("modalGlobalNotificacionLabel");
    const headerEl = document.getElementById("modalGlobalNotificacionHeader");
    const bodyEl = document.getElementById("modalGlobalNotificacionBody");
    const btnEl = document.getElementById("modalGlobalNotificacionBtn");

    bodyEl.textContent = message;

    // Resetear clases
    headerEl.classList.remove("bg-danger", "bg-success");
    btnEl.classList.remove("btn-danger", "btn-success", "btn-primary");

    if (isError) {
        titleEl.textContent = "Error";
        headerEl.classList.add("bg-danger");
        btnEl.classList.add("btn-danger");
    } else {
        titleEl.textContent = "Éxito";
        headerEl.classList.add("bg-success");
        btnEl.classList.add("btn-success");
    }

    // Forzar z-index más alto para evitar que quede detrás si hay otros modales abiertos
    let modalInstance = bootstrap.Modal.getInstance(modalEl);
    if (!modalInstance) {
        modalInstance = new bootstrap.Modal(modalEl);
    }
    modalInstance.show();

    // Arreglar el backdrop inmediatamente después de mostrar
    setTimeout(() => {
        const backdrops = document.querySelectorAll('.modal-backdrop');
        if (backdrops.length > 0) {
            const highestZIndex = 1100 + (backdrops.length * 10);
            const lastBackdrop = backdrops[backdrops.length - 1];
            lastBackdrop.style.zIndex = highestZIndex;
            modalEl.style.zIndex = highestZIndex + 1;
        }
    }, 15);
}