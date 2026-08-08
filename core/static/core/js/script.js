// Main script for RENEE Warehouse
console.log("Welcome to RENEE Warehouse.");

// Toasts are owned by the single app-wide system in enhance.js (window.B2B.toast),
// which also adopts Django's server-rendered messages and routes alert()/createToast
// through the same code path. This helper stays only as a thin, backward-compatible
// shim: it forwards to B2B.toast when present, and falls back to a minimal inline
// toast if (and only if) enhance.js hasn't loaded — so there is never a second,
// competing toast style on the page.
window.createToast = function (message, type) {
    if (window.B2B && typeof window.B2B.toast === 'function') {
        var map = { success: 'success', error: 'error', warning: 'warn', warn: 'warn', info: 'info', danger: 'error' };
        return window.B2B.toast(String(message == null ? '' : message), { type: map[type] || 'info' });
    }
    // Fallback (enhance.js absent): minimal, self-dismissing toast.
    var container = document.getElementById('toast-container');
    if (!container) return;
    var toast = document.createElement('div');
    toast.className = 'toast toast-' + (type || 'info');
    var msgSpan = document.createElement('span');
    msgSpan.innerText = message;
    var closeBtn = document.createElement('button');
    closeBtn.className = 'toast-close';
    closeBtn.innerHTML = '&times;';
    closeBtn.onclick = function () { toast.remove(); };
    toast.appendChild(msgSpan);
    toast.appendChild(closeBtn);
    container.appendChild(toast);
    setTimeout(function () {
        if (toast && toast.parentElement) {
            toast.classList.add('toast-fade-out');
            setTimeout(function () { toast.remove(); }, 1000);
        }
    }, 3000);
};
