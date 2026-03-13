document.addEventListener('DOMContentLoaded', async () => {

    // ── Theme ─────────────────────────────────────────────
    const htmlEl = document.documentElement;
    const themeSwitch = document.getElementById('themeSwitch');
    const savedTheme = localStorage.getItem('spliqour-theme') || 'dark';
    applyTheme(savedTheme);

    themeSwitch.addEventListener('change', toggleTheme);

    function applyTheme(t) {
        htmlEl.setAttribute('data-theme', t);
        themeSwitch.checked = (t === 'dark');
        localStorage.setItem('spliqour-theme', t);
    }
    function toggleTheme() {
        applyTheme(htmlEl.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
    }

    // ── Elements ───────────────────────────────────────────
    const searchInput = document.getElementById('searchInput');
    const barList = document.getElementById('barList');
    const emptyState = document.getElementById('emptyState');
    const barCountEl = document.getElementById('barCount');
    const addBarBtn = document.getElementById('addBarBtn');
    const emptyAddBtn = document.getElementById('emptyAddBtn');
    const statTotalBars = document.getElementById('statTotalBars');
    const statCurrentFY = document.getElementById('statCurrentFY');
    const statCities = document.getElementById('statCities');

    // ── FY Filter ──────────────────────────────────────────
    const fyFilterBtn = document.getElementById('fyFilterBtn');
    const fyFilterLabel = document.getElementById('fyFilterLabel');
    let fyFilterActive = true;

    function currentFYString() {
        const today = new Date();
        const m = today.getMonth() + 1; // 1–12
        const y = today.getFullYear();
        const start = m >= 4 ? y : y - 1;
        const end = start + 1;
        // Match format stored in bar.financialYear e.g. "01/04/2025 to 31/03/2026"
        return `${start}/${end}`; // used for substring match
    }

    function matchesCurrentFY(bar) {
        const fy = bar.financialYear || '';
        const today = new Date();
        const m = today.getMonth() + 1;
        const y = today.getFullYear();
        const startYear = String(m >= 4 ? y : y - 1);
        const endYear   = String(m >= 4 ? y + 1 : y);
        return fy.includes(startYear) && fy.includes(endYear);
    }

    function applyFilters() {
        const q = searchInput.value.toLowerCase();
        let results = allBars;
        if (fyFilterActive) results = results.filter(matchesCurrentFY);
        if (q) results = results.filter(b =>
            (b.barName || '').toLowerCase().includes(q) ||
            (b.city || '').toLowerCase().includes(q) ||
            (b.shopType || '').toLowerCase().includes(q) ||
            (b.bar_id || '').toLowerCase().includes(q)
        );
        focusedIdx = 0;
        render(results);
    }

    fyFilterBtn.addEventListener('click', () => {
        fyFilterActive = !fyFilterActive;
        fyFilterBtn.classList.toggle('active', fyFilterActive);
        applyFilters();
    });

    // ── Load bars ──────────────────────────────────────────
    let allBars = [];
    let filtered = [];
    let focusedIdx = 0;
    let selectedIdx = -1;
    let selectedBarKey = '';
    let openingBar = false;

    try {
        const activeRaw = sessionStorage.getItem('active-bar');
        if (activeRaw) {
            const active = JSON.parse(activeRaw);
            selectedBarKey = `${(active.barName || '').toLowerCase()}|${active.financialYear || ''}`;
        }
    } catch (_) { /* ignore parse errors */ }

    async function loadBars() {
        try {
            const result = await window.electronAPI.getBarsIndex();
            if (result.success) {
                allBars = result.bars || [];
            } else {
                console.error('get-bars-index failed:', result.error);
                allBars = [];
            }
        } catch (err) {
            console.error('getBarsIndex IPC error:', err);
            allBars = [];
        }
        filtered = [...allBars];
        updateStats();
        applyFilters();
    }

    await loadBars();
    searchInput.focus();

    // ── Update dashboard stats ─────────────────────────────
    function updateStats() {
        // Total bars
        statTotalBars.textContent = allBars.length;

        // Current FY
        if (allBars.length > 0) {
            const fy = fyShortLabel(allBars[0].financialYear);
            statCurrentFY.textContent = fy;
        } else {
            statCurrentFY.textContent = '—';
        }

        // Unique cities
        const cities = new Set(allBars.map(b => (b.city || '').trim().toLowerCase()).filter(Boolean));
        statCities.textContent = cities.size;
    }

    // ── Render cards ───────────────────────────────────────
    function render(bars) {
        barList.innerHTML = '';
        filtered = bars;
        selectedIdx = -1;

        if (bars.length === 0) {
            emptyState.classList.remove('hidden');
            barCountEl.textContent = 'No bars found';
            return;
        }
        emptyState.classList.add('hidden');
        barCountEl.textContent = `${bars.length} bar${bars.length !== 1 ? 's' : ''}`;

        bars.forEach((bar, idx) => {
            const card = document.createElement('div');
            card.className = 'bar-card';
            card.tabIndex = 0;
            card.dataset.idx = idx;
            card.style.animationDelay = `${idx * 0.04}s`;

            const cardKey = `${(bar.barName || '').toLowerCase()}|${bar.financialYear || ''}`;
            if (selectedBarKey && cardKey === selectedBarKey) {
                card.classList.add('selected');
                selectedIdx = idx;
            }

            const fyShort = fyShortLabel(bar.financialYear);
            const initial = (bar.barName || 'B').charAt(0).toUpperCase();

            card.innerHTML = `
                <div class="bar-card-header">
                    <div class="bar-card-icon">${esc(initial)}</div>
                    <span class="bar-card-badge">${esc(bar.shopType || 'FL')}</span>
                </div>
                <div class="bar-card-name" title="${esc(bar.barName)}">${esc(bar.barName)}</div>
                <div class="bar-card-meta">
                    <div class="bar-card-meta-row">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>
                        ${esc(bar.city || '—')}
                    </div>
                    <div class="bar-card-meta-row">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
                        ${esc(fyShort)}
                    </div>
                </div>
                <div class="bar-card-arrow">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"/></svg>
                </div>
            `;

            card.addEventListener('click', () => openBarByIndex(idx));
            card.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') { e.preventDefault(); openBarByIndex(idx); }
            });
            card.addEventListener('focus', () => setFocus(idx));

            barList.appendChild(card);
        });

        setFocus(selectedIdx >= 0 ? selectedIdx : 0);
    }

    // ── Focus management ───────────────────────────────────
    function setFocus(idx) {
        const cards = barList.querySelectorAll('.bar-card');
        cards.forEach(c => c.classList.remove('focused'));
        if (cards[idx]) {
            cards[idx].classList.add('focused');
            cards[idx].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
            focusedIdx = idx;
        }
    }

    function setSelected(idx) {
        const cards = barList.querySelectorAll('.bar-card');
        cards.forEach(c => c.classList.remove('selected'));
        if (cards[idx]) {
            cards[idx].classList.add('selected');
            selectedIdx = idx;
            const bar = filtered[idx];
            if (bar) selectedBarKey = `${(bar.barName || '').toLowerCase()}|${bar.financialYear || ''}`;
        }
    }

    // ── Open bar ───────────────────────────────────────────
    async function openBar(barEntry) {
        if (!barEntry || openingBar) return;
        openingBar = true;
        const idx = filtered.findIndex(b => b === barEntry);
        if (idx >= 0) setSelected(idx);
        try {
            const result = await window.electronAPI.openBar(barEntry);
            if (result.success) {
                sessionStorage.setItem('active-bar', JSON.stringify(result.data));
                window.electronAPI.navigateToApp();
            } else {
                alert('Failed to open bar: ' + result.error);
            }
        } finally {
            openingBar = false;
        }
    }

    async function openBarByIndex(idx) {
        if (idx < 0 || idx >= filtered.length) return;
        await openBar(filtered[idx]);
    }

    // ── Search ─────────────────────────────────────────────
    searchInput.addEventListener('input', () => applyFilters());

    // ── Global keyboard shortcuts ──────────────────────────
    document.addEventListener('keydown', (e) => {
        // Theme
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 't') {
            e.preventDefault(); toggleTheme(); return;
        }
        // Add new bar
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'n') {
            e.preventDefault(); window.electronAPI.navigateToAddBar(); return;
        }
        // Clear search
        if (e.key === 'Escape') {
            searchInput.value = '';
            applyFilters();
            searchInput.focus();
            return;
        }

        const cards = barList.querySelectorAll('.bar-card');
        if (!cards.length) return;

        if (e.key === 'ArrowDown' || e.key === 'ArrowRight') {
            e.preventDefault();
            setFocus(Math.min(focusedIdx + 1, cards.length - 1));
            cards[focusedIdx].focus();
        } else if (e.key === 'ArrowUp' || e.key === 'ArrowLeft') {
            e.preventDefault();
            if (focusedIdx === 0) { searchInput.focus(); return; }
            setFocus(Math.max(focusedIdx - 1, 0));
            cards[focusedIdx].focus();
        } else if (e.key === 'Enter') {
            if (document.activeElement === searchInput && cards.length > 0) {
                e.preventDefault();
                openBarByIndex(focusedIdx);
            }
        }
    });

    // ── Add Bar buttons ────────────────────────────────────
    addBarBtn.addEventListener('click', () => window.electronAPI.navigateToAddBar());
    if (emptyAddBtn) {
        emptyAddBtn.addEventListener('click', () => window.electronAPI.navigateToAddBar());
    }

    // ── Data Folder strip ─────────────────────────────────
    const dataFolderPathEl = document.getElementById('dataFolderPath');
    const openDataFolderBtn = document.getElementById('openDataFolderBtn');
    const changeDataFolderBtn = document.getElementById('changeDataFolderBtn');

    async function refreshDataFolderDisplay() {
        try {
            const res = await window.electronAPI.getDataRoot();
            if (res && res.success) {
                dataFolderPathEl.textContent = res.dataRoot;
                dataFolderPathEl.title = res.dataRoot;
            }
        } catch (e) { /* ignore */ }
    }
    await refreshDataFolderDisplay();

    openDataFolderBtn.addEventListener('click', async () => {
        await window.electronAPI.openDataFolder();
    });

    changeDataFolderBtn.addEventListener('click', async () => {
        const res = await window.electronAPI.chooseDataFolder();
        if (res && res.success) {
            dataFolderPathEl.textContent = res.dataRoot;
            dataFolderPathEl.title = res.dataRoot;
            // Reload bars from new location
            await loadBars();
        }
    });

    // ── Helpers ────────────────────────────────────────────
    function esc(str) {
        return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    function fyShortLabel(fyLabel) {
        const m = (fyLabel || '').match(/(\d{4})/g);
        if (m && m.length >= 2) return `FY ${m[0]}-${m[1].slice(2)}`;
        return fyLabel || '—';
    }
});
