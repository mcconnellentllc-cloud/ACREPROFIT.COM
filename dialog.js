/**
 * AcreProfit shared dialog helpers.
 * Drop-in replacements for native alert(), confirm(), prompt().
 *
 *   window.toast(message, type)
 *       type: 'info' (default) | 'success' | 'error' | 'warning'
 *       Top-right stack, auto-dismiss 3s, click-to-dismiss.
 *       Fire-and-forget, no return value.
 *
 *   window.confirmDialog(message, opts)
 *       opts: { title?, ok='Confirm', cancel='Cancel', danger=false }
 *       Returns Promise<boolean>. Enter=OK, Escape=Cancel.
 *
 *   window.promptDialog(message, opts)
 *       opts: { title?, default='', placeholder='', ok='OK', cancel='Cancel' }
 *       Returns Promise<string|null>. Enter=OK, Escape=Cancel (null).
 *
 * Include once per page:
 *     <script src="dialog.js"></script>
 *
 * Loads before inline scripts; no dependencies; no module system.
 */
(function () {
    'use strict';

    if (window.toast && window.confirmDialog && window.promptDialog) return;

    // One-time CSS injection. Scoped with .ap- prefix so it can't
    // collide with existing page styles.
    const STYLE_ID = 'ap-dialog-styles';
    if (!document.getElementById(STYLE_ID)) {
        const style = document.createElement('style');
        style.id = STYLE_ID;
        style.textContent = `
            .ap-toast-container {
                position: fixed;
                top: 20px;
                right: 20px;
                z-index: 10000;
                display: flex;
                flex-direction: column;
                gap: 10px;
                pointer-events: none;
                max-width: calc(100vw - 40px);
            }
            .ap-toast {
                pointer-events: auto;
                min-width: 260px;
                max-width: 420px;
                padding: 14px 18px;
                border-radius: 8px;
                color: white;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
                font-size: 0.95rem;
                line-height: 1.4;
                box-shadow: 0 4px 16px rgba(0,0,0,0.18);
                cursor: pointer;
                opacity: 0;
                transform: translateX(20px);
                transition: opacity 0.2s ease, transform 0.2s ease;
                display: flex;
                align-items: start;
                gap: 10px;
            }
            .ap-toast.ap-visible {
                opacity: 1;
                transform: translateX(0);
            }
            .ap-toast.ap-info    { background: #6b7280; }
            .ap-toast.ap-success { background: #2d5a27; }
            .ap-toast.ap-error   { background: #dc2626; }
            .ap-toast.ap-warning { background: #d97706; }
            .ap-toast-icon { font-weight: 700; flex-shrink: 0; }

            .ap-modal-overlay {
                position: fixed;
                inset: 0;
                background: rgba(17, 24, 39, 0.55);
                z-index: 10001;
                display: flex;
                align-items: center;
                justify-content: center;
                padding: 20px;
                opacity: 0;
                transition: opacity 0.15s ease;
            }
            .ap-modal-overlay.ap-visible { opacity: 1; }
            .ap-modal {
                background: white;
                border-radius: 10px;
                padding: 24px 28px;
                max-width: 460px;
                width: 100%;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
                box-shadow: 0 20px 60px rgba(0,0,0,0.3);
                transform: translateY(10px);
                transition: transform 0.15s ease;
            }
            .ap-modal-overlay.ap-visible .ap-modal { transform: translateY(0); }
            .ap-modal-title {
                font-size: 1.05rem;
                font-weight: 700;
                color: #111;
                margin: 0 0 8px 0;
            }
            .ap-modal-message {
                color: #374151;
                line-height: 1.5;
                margin: 0 0 16px 0;
                white-space: pre-wrap;
            }
            .ap-modal-input {
                width: 100%;
                padding: 10px 12px;
                border: 2px solid #e5e7eb;
                border-radius: 6px;
                font-size: 1rem;
                margin-bottom: 16px;
                box-sizing: border-box;
                font-family: inherit;
            }
            .ap-modal-input:focus {
                outline: none;
                border-color: #2d5a27;
            }
            .ap-modal-actions {
                display: flex;
                justify-content: flex-end;
                gap: 8px;
            }
            .ap-modal-btn {
                padding: 9px 18px;
                border-radius: 6px;
                border: none;
                font-size: 0.95rem;
                font-weight: 600;
                cursor: pointer;
                font-family: inherit;
                transition: opacity 0.15s ease;
            }
            .ap-modal-btn:hover { opacity: 0.85; }
            .ap-modal-btn-cancel {
                background: #f3f4f6;
                color: #374151;
            }
            .ap-modal-btn-ok {
                background: #2d5a27;
                color: white;
            }
            .ap-modal-btn-ok.ap-danger {
                background: #dc2626;
            }
        `;
        document.head.appendChild(style);
    }

    function ensureToastContainer() {
        let container = document.querySelector('.ap-toast-container');
        if (!container) {
            container = document.createElement('div');
            container.className = 'ap-toast-container';
            document.body.appendChild(container);
        }
        return container;
    }

    const ICONS = {
        info: 'ℹ',
        success: '✓',
        error: '✕',
        warning: '⚠'
    };

    window.toast = function (message, type) {
        const t = ['info', 'success', 'error', 'warning'].includes(type) ? type : 'info';
        const container = ensureToastContainer();

        const el = document.createElement('div');
        el.className = 'ap-toast ap-' + t;
        el.innerHTML =
            '<span class="ap-toast-icon">' + ICONS[t] + '</span>' +
            '<span class="ap-toast-msg"></span>';
        el.querySelector('.ap-toast-msg').textContent = String(message == null ? '' : message);

        container.appendChild(el);
        // Next-frame to trigger transition
        requestAnimationFrame(() => el.classList.add('ap-visible'));

        const dismiss = () => {
            el.classList.remove('ap-visible');
            setTimeout(() => el.remove(), 220);
        };
        el.addEventListener('click', dismiss);

        // Error toasts linger longer (4.5s). Others auto-dismiss at 3s.
        const ttl = t === 'error' ? 4500 : 3000;
        setTimeout(dismiss, ttl);
    };

    // Internal: renders a modal overlay with actions + optional input.
    // Returns the overlay element + a cleanup function.
    function buildModal({ title, message, input = null, ok = 'OK', cancel = 'Cancel', danger = false, onResolve }) {
        const overlay = document.createElement('div');
        overlay.className = 'ap-modal-overlay';

        const modal = document.createElement('div');
        modal.className = 'ap-modal';

        if (title) {
            const t = document.createElement('div');
            t.className = 'ap-modal-title';
            t.textContent = title;
            modal.appendChild(t);
        }

        const m = document.createElement('div');
        m.className = 'ap-modal-message';
        m.textContent = String(message == null ? '' : message);
        modal.appendChild(m);

        let inputEl = null;
        if (input) {
            inputEl = document.createElement('input');
            inputEl.type = 'text';
            inputEl.className = 'ap-modal-input';
            if (input.default != null) inputEl.value = String(input.default);
            if (input.placeholder) inputEl.placeholder = String(input.placeholder);
            modal.appendChild(inputEl);
        }

        const actions = document.createElement('div');
        actions.className = 'ap-modal-actions';

        const cancelBtn = document.createElement('button');
        cancelBtn.className = 'ap-modal-btn ap-modal-btn-cancel';
        cancelBtn.textContent = cancel;

        const okBtn = document.createElement('button');
        okBtn.className = 'ap-modal-btn ap-modal-btn-ok' + (danger ? ' ap-danger' : '');
        okBtn.textContent = ok;

        actions.appendChild(cancelBtn);
        actions.appendChild(okBtn);
        modal.appendChild(actions);
        overlay.appendChild(modal);
        document.body.appendChild(overlay);

        requestAnimationFrame(() => overlay.classList.add('ap-visible'));

        const cleanup = () => {
            overlay.classList.remove('ap-visible');
            document.removeEventListener('keydown', keyHandler);
            setTimeout(() => overlay.remove(), 180);
        };

        const accept = () => {
            const value = inputEl ? inputEl.value : true;
            cleanup();
            onResolve(value);
        };
        const reject = () => {
            cleanup();
            onResolve(inputEl ? null : false);
        };

        okBtn.addEventListener('click', accept);
        cancelBtn.addEventListener('click', reject);
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) reject();
        });

        const keyHandler = (e) => {
            if (e.key === 'Escape') {
                e.preventDefault();
                reject();
            } else if (e.key === 'Enter' && !(inputEl && e.target !== inputEl && e.target.tagName === 'TEXTAREA')) {
                // Enter in input or outside textarea triggers OK
                e.preventDefault();
                accept();
            }
        };
        document.addEventListener('keydown', keyHandler);

        // Focus management: input first, otherwise OK button.
        setTimeout(() => (inputEl || okBtn).focus(), 0);
    }

    window.confirmDialog = function (message, opts) {
        const o = opts || {};
        return new Promise((resolve) => {
            buildModal({
                title: o.title || null,
                message: message,
                input: null,
                ok: o.ok || 'Confirm',
                cancel: o.cancel || 'Cancel',
                danger: !!o.danger,
                onResolve: resolve
            });
        });
    };

    window.promptDialog = function (message, opts) {
        const o = opts || {};
        return new Promise((resolve) => {
            buildModal({
                title: o.title || null,
                message: message,
                input: {
                    default: o.default || '',
                    placeholder: o.placeholder || ''
                },
                ok: o.ok || 'OK',
                cancel: o.cancel || 'Cancel',
                danger: false,
                onResolve: resolve
            });
        });
    };
})();
