// Shared styled confirm dialog — replaces the browser-native confirm()
// popup for destructive actions (currently: deleting a JD) with a modal
// that matches the rest of the app's look. Usage:
//   appConfirm("Are you sure?", {title: "Delete Job Description", okLabel: "Delete"})
//     .then(function (confirmed) { if (confirmed) { ... } });
(function () {
    function ensureModal() {
        if (document.getElementById('appConfirmOverlay')) return;
        var overlay = document.createElement('div');
        overlay.id = 'appConfirmOverlay';
        overlay.className = 'app-confirm-overlay';
        overlay.innerHTML =
            '<div class="app-confirm-box">' +
                '<p class="app-confirm-title" id="appConfirmTitle"></p>' +
                '<p class="app-confirm-message" id="appConfirmMessage"></p>' +
                '<div class="app-confirm-actions">' +
                    '<button type="button" class="button secondary" id="appConfirmCancelBtn">Cancel</button>' +
                    '<button type="button" class="button danger" id="appConfirmOkBtn">OK</button>' +
                '</div>' +
            '</div>';
        document.body.appendChild(overlay);
    }

    window.appConfirm = function (message, opts) {
        opts = opts || {};
        ensureModal();
        var overlay = document.getElementById('appConfirmOverlay');
        document.getElementById('appConfirmTitle').textContent = opts.title || 'Please confirm';
        document.getElementById('appConfirmMessage').textContent = message || 'Are you sure?';
        var okBtn = document.getElementById('appConfirmOkBtn');
        var cancelBtn = document.getElementById('appConfirmCancelBtn');
        okBtn.textContent = opts.okLabel || 'OK';
        overlay.style.display = 'flex';

        return new Promise(function (resolve) {
            function cleanup(result) {
                overlay.style.display = 'none';
                okBtn.removeEventListener('click', onOk);
                cancelBtn.removeEventListener('click', onCancel);
                overlay.removeEventListener('click', onOverlayClick);
                document.removeEventListener('keydown', onKeydown);
                resolve(result);
            }
            function onOk() { cleanup(true); }
            function onCancel() { cleanup(false); }
            function onOverlayClick(e) { if (e.target === overlay) cleanup(false); }
            function onKeydown(e) { if (e.key === 'Escape') cleanup(false); }

            okBtn.addEventListener('click', onOk);
            cancelBtn.addEventListener('click', onCancel);
            overlay.addEventListener('click', onOverlayClick);
            document.addEventListener('keydown', onKeydown);
        });
    };

    // Auto-wire: any <form class="app-confirm-form" data-confirm-message="...">
    // gets this modal instead of a native confirm() automatically, site-wide —
    // no per-page script needed. form.submit() (unlike a real user click)
    // does not re-fire the 'submit' event, so this never loops.
    document.addEventListener('submit', function (e) {
        var form = e.target;
        if (!(form instanceof HTMLFormElement) || !form.classList.contains('app-confirm-form')) return;
        e.preventDefault();
        appConfirm(form.dataset.confirmMessage, {
            title: form.dataset.confirmTitle,
            okLabel: form.dataset.confirmOkLabel,
        }).then(function (ok) { if (ok) form.submit(); });
    });
})();
