document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('add-bar-form');
    // Include both inputs AND selects in the keyboard-navigation list
    const inputs = Array.from(form.querySelectorAll('input, select'));
    const sections = Array.from(form.querySelectorAll('.form-section'));
    const toast = document.getElementById('toast');
    const themeSwitch = document.getElementById('themeSwitch');
    const htmlEl = document.documentElement;

    // ── Theme ──────────────────────────────────────────────
    const savedTheme = localStorage.getItem('spliqour-theme') || 'dark';
    applyTheme(savedTheme);

    themeSwitch.addEventListener('change', () => {
        toggleTheme();
    });

    function applyTheme(theme) {
        htmlEl.setAttribute('data-theme', theme);
        themeSwitch.checked = (theme === 'dark');
        localStorage.setItem('spliqour-theme', theme);
    }

    function toggleTheme() {
        const next = htmlEl.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
        applyTheme(next);
    }

    // ── Home Navigation ────────────────────────────────────
    const homeBtn = document.getElementById('homeBtn');
    if (homeBtn) {
        homeBtn.addEventListener('click', () => window.electronAPI.navigateHome());
    }

    // ── License Prefix Auto-Sync ───────────────────────────
    const shopTypeEl = document.getElementById('shopType');
    const licensePrefix = document.getElementById('licensePrefix');
    const licenseNoInput = document.getElementById('licenseNoInput');
    const licenseNoHidden = document.getElementById('licenseNo');
    const licenseTypeEl = document.getElementById('licenseType');

    function syncLicensePrefix() {
        const val = shopTypeEl ? shopTypeEl.value : 'FL-III';
        if (licensePrefix) licensePrefix.textContent = val;
        if (licenseTypeEl) licenseTypeEl.value = val;
        updateFullLicenseNo();
    }

    function updateFullLicenseNo() {
        const prefix = licensePrefix ? licensePrefix.textContent : '';
        const num = licenseNoInput ? licenseNoInput.value.trim() : '';
        if (licenseNoHidden) licenseNoHidden.value = num ? `${prefix}/${num}` : '';
    }

    if (shopTypeEl) shopTypeEl.addEventListener('change', syncLicensePrefix);
    if (licenseNoInput) licenseNoInput.addEventListener('input', updateFullLicenseNo);

    // Init on load
    syncLicensePrefix();

    // ── Financial Year Dropdown Builder ───────────────────
    (function buildFYDropdown() {
        const fySelect = document.getElementById('financialYear');
        if (!fySelect) return;

        // Determine current FY: starts 1-Apr, ends 31-Mar
        const today = new Date();
        const month = today.getMonth() + 1; // 1–12
        const year = today.getFullYear();
        const currentFYStart = month >= 4 ? year : year - 1;

        // Generate 3 past + current + 2 future FYs
        for (let y = currentFYStart - 3; y <= currentFYStart + 2; y++) {
            const label = `01/04/${y} to 31/03/${y + 1}`;
            const opt = document.createElement('option');
            opt.value = label;
            opt.textContent = `FY ${y}-${String(y + 1).slice(2)}  (${label})`;
            if (y === currentFYStart) opt.selected = true;
            fySelect.appendChild(opt);
        }
    })();

    // ── Load Active Bar (if opening from Home) ─────────────
    let activeBarId = null;
    let activeBarCreatedAt = null;

    function loadActiveBar() {
        const raw = sessionStorage.getItem('active-bar');
        if (!raw) return;

        try {
            const data = JSON.parse(raw);
            activeBarId = data.bar_id;
            activeBarCreatedAt = data.created_at;

            inputs.forEach(input => {
                if (input.name && data[input.name] !== undefined) {
                    input.value = data[input.name];
                }
            });

            // Special case logic for split license No
            if (data.licenseNoInput) {
                const licenseNoInputEl = document.getElementById('licenseNoInput');
                if (licenseNoInputEl) licenseNoInputEl.value = data.licenseNoInput;
            }

            syncLicensePrefix();

            // Clear sessionStorage so if we click "Add New Bar" later, it starts fresh
            sessionStorage.removeItem('active-bar');

            // Show a quick hint
            showToast('Loaded: ' + (data.barName || 'Saved Bar'));

        } catch (err) {
            console.error('Failed to parse active-bar', err);
        }
    }

    loadActiveBar();

    // ── Section Highlighting ───────────────────────────────


    function highlightActiveSection(inputEl) {
        sections.forEach(sec => sec.classList.remove('active-section'));
        const parentSection = inputEl.closest('.form-section');
        if (parentSection) {
            parentSection.classList.add('active-section');
            // Scroll section into view smoothly
            parentSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }
    }

    inputs.forEach(input => {
        input.addEventListener('focus', () => highlightActiveSection(input));

        // ── Auto Title-Case for text fields ───────────────────
        // Skip: uppercase-only fields, date, email, select
        if (
            input.tagName === 'SELECT' ||
            input.type === 'date' ||
            input.type === 'email' ||
            input.style.textTransform === 'uppercase'
        ) return;

        input.addEventListener('input', () => {
            const pos = input.selectionStart; // remember cursor position
            input.value = toTitleCase(input.value);
            input.setSelectionRange(pos, pos); // restore cursor
        });
    });

    // Title Case converter
    function toTitleCase(str) {
        return str.replace(/\w\S*/g, (word) =>
            word.charAt(0).toUpperCase() + word.slice(1)
        );
    }

    // ── Focus management ───────────────────────────────────
    if (inputs.length > 0) {
        inputs[0].focus();
        highlightActiveSection(inputs[0]);
    }

    // ── Global Keyboard Shortcuts ──────────────────────────
    document.addEventListener('keydown', async (e) => {
        // Save (Ctrl/Cmd + S)
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
            e.preventDefault();
            await saveBar();
            return;
        }

        // Theme Toggle (Ctrl/Cmd + T)
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 't') {
            e.preventDefault();
            toggleTheme();
            return;
        }

        // Go Home (Ctrl/Cmd + Home)
        if ((e.ctrlKey || e.metaKey) && e.key === 'Home') {
            e.preventDefault();
            window.electronAPI.navigateHome();
            return;
        }

        // Clear Form (Esc)
        if (e.key === 'Escape') {
            e.preventDefault();
            clearForm();
            return;
        }

        // Navigate (Enter / Arrow Down / Arrow Up)
        const active = document.activeElement;
        if (!inputs.includes(active)) return;
        const idx = inputs.indexOf(active);

        if (e.key === 'Enter' || e.key === 'ArrowDown') {
            e.preventDefault();
            const nextIdx = (idx + 1) % inputs.length;
            inputs[nextIdx].focus();
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            const prevIdx = (idx - 1 + inputs.length) % inputs.length;
            inputs[prevIdx].focus();
        }
    });

    // ── Save Bar ───────────────────────────────────────────
    async function saveBar() {
        const barName = document.getElementById('barName').value.trim();
        const ownerName = document.getElementById('ownerName').value.trim();
        const phone = document.getElementById('phone').value.trim();

        if (!barName) {
            showToast('⚠️ Bar / Shop Name is required!', 'error');
            inputs[0].focus();
            return;
        }
        if (!ownerName) {
            showToast('⚠️ Owner Name is required!', 'error');
            document.getElementById('ownerName').focus();
            return;
        }
        if (!phone) {
            showToast('⚠️ Mobile Number is required!', 'error');
            document.getElementById('phone').focus();
            return;
        }

        // Build data object from all inputs
        const barData = {};
        inputs.forEach(input => {
            if (input.name) barData[input.name] = input.value.trim();
        });
        barData.bar_id = activeBarId || 'B' + Date.now();
        barData.created_at = activeBarCreatedAt || new Date().toISOString();

        try {
            const result = await window.electronAPI.saveBarData(barData);
            if (result.success) {
                const fy = document.getElementById('financialYear').value || '';
                showToast(`✅ Bar saved → data/${document.getElementById('barName').value.trim().replace(/\s+/g, '_')}/${fy ? 'FY_' + fy.match(/(\d{4})/)?.[0] + '-' + fy.match(/(\d{4})/g)?.[1]?.slice(2) : ''}/`);
                clearForm();
            } else {
                showToast('❌ Save failed: ' + result.error, 'error');
            }
        } catch (err) {
            showToast('❌ Error: ' + err.message, 'error');
        }
    }

    // ── Clear Form ─────────────────────────────────────────
    function clearForm() {
        form.reset();
        activeBarId = null;
        activeBarCreatedAt = null;
        // Restore smart defaults after reset
        if (shopTypeEl) shopTypeEl.value = 'FL-III';
        const entityTypeEl = document.getElementById('entityType');
        if (entityTypeEl) entityTypeEl.value = 'Proprietorship';
        const stateEl = document.getElementById('state');
        if (stateEl) stateEl.value = 'Maharashtra';
        // Re-sync license prefix badge & clear number field
        if (licenseNoInput) licenseNoInput.value = '';
        if (licenseNoHidden) licenseNoHidden.value = '';
        syncLicensePrefix();
        sections.forEach(sec => sec.classList.remove('active-section'));
        inputs[0].focus();
        highlightActiveSection(inputs[0]);
    }

    // ── Toast Notification ─────────────────────────────────
    let toastTimer;
    function showToast(msg, type = 'success') {
        clearTimeout(toastTimer);
        toast.textContent = msg;
        toast.className = 'toast' + (type === 'error' ? ' error' : '');
        toastTimer = setTimeout(() => {
            toast.classList.add('hidden');
        }, 3500);
    }
});
