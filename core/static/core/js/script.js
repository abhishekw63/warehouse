// Main script for RENEE Warehouse
console.log("Welcome to RENEE Warehouse.");

// Helper to create toasts programmatically
window.createToast = function(message, type) {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    const msgSpan = document.createElement('span');
    msgSpan.innerText = message;

    const closeBtn = document.createElement('button');
    closeBtn.className = 'toast-close';
    closeBtn.innerHTML = '&times;';
    closeBtn.onclick = function() { toast.remove(); };

    toast.appendChild(msgSpan);
    toast.appendChild(closeBtn);
    container.appendChild(toast);

    setTimeout(() => {
        if(toast && toast.parentElement) {
            toast.classList.add('toast-fade-out');
            setTimeout(() => toast.remove(), 1000); // Wait for the transition to finish
        }
    }, 3000);
};

// Auto-dismiss server-rendered toasts after 3 seconds
document.addEventListener("DOMContentLoaded", () => {
    const toasts = document.querySelectorAll(".toast");
    toasts.forEach((toast) => {
        setTimeout(() => {
            if(toast && toast.parentElement) {
                toast.classList.add('toast-fade-out');
                setTimeout(() => toast.remove(), 1000); // Wait for the transition to finish
            }
        }, 3000);
    });
});
