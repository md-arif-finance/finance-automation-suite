document.addEventListener('DOMContentLoaded', () => {
    // Console log for now, can add interactivity later
    console.log('Automated AR Control System Portfolio Loaded');

    // Modal Interaction Logic
    const getTemplateBtn = document.getElementById('get-template-btn');
    const modalOverlay = document.getElementById('verification-modal');
    const modalConfirmBtn = document.getElementById('modal-confirm-btn');

    if (getTemplateBtn && modalOverlay && modalConfirmBtn) {
        getTemplateBtn.addEventListener('click', (e) => {
            e.preventDefault();
            modalOverlay.classList.remove('d-none');
            // Small delay to allow CSS transition to work
            setTimeout(() => {
                modalOverlay.classList.add('active');
            }, 10);
        });

        // Confirm Action
        modalConfirmBtn.addEventListener('click', () => {
            // Redirect to the actual Google Sheet in a new tab
            window.open("https://docs.google.com/spreadsheets", "_blank");
            // Close the modal
            modalOverlay.classList.remove('active');
            setTimeout(() => {
                modalOverlay.classList.add('d-none');
            }, 300);
        });

        // Close on outside click (Optional, but good UX)
        modalOverlay.addEventListener('click', (e) => {
            if (e.target === modalOverlay) {
                modalOverlay.classList.remove('active');
                setTimeout(() => {
                    modalOverlay.classList.add('d-none');
                }, 300); // Match CSS transition duration
            }
        });
    }
});
