document.addEventListener('DOMContentLoaded', () => {
    // ── Top-bar auto-resize on window maximize / restore ──
    (function() {
        function applyTopbarHeight(maximized) {
            var h = maximized ? 30 : 38;
            document.documentElement.style.setProperty('--topbar-drag-height', h + 'px');
            document.body.classList.toggle('win-maximized', !!maximized);
            // Swap maximize / restore icon
            var maxBtn = document.getElementById('winMaximize');
            if (maxBtn) maxBtn.title = maximized ? 'Restore' : 'Maximize';
        }
        applyTopbarHeight(false);
        if (window.electronAPI && window.electronAPI.onWindowStateChanged) {
            window.electronAPI.onWindowStateChanged(function(state) {
                applyTopbarHeight(state.maximized);
            });
        }
    })();

    // ── IST Clock (second row, center) ──
    (function() {
        var clockEl = document.getElementById('tsbClock');
        if (!clockEl) return;
        function tickClock() {
            clockEl.textContent = new Date().toLocaleTimeString('en-IN', {
                timeZone: 'Asia/Kolkata',
                hour12: true,
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            });
        }
        tickClock();
        setInterval(tickClock, 1000);
    })();

    // ── Shared print helper ── use hidden iframe so Chromium shows its own print preview ──
    let _previewMode = false;
    function downloadPdf(html, filename) {
        if (window.electronAPI?.exportHtmlPdf) {
            window.electronAPI.exportHtmlPdf({ html, filename }).catch(() => {});
        }
    }

    function printWithIframe(html) {
        if (_previewMode) {
            _previewMode = false;
            showPrintPreview(html);
            return;
        }
        let fr = document.getElementById('__spliqour_print_frame__');
        if (fr) fr.remove();
        fr = document.createElement('iframe');
        fr.id = '__spliqour_print_frame__';
        fr.style.cssText = 'position:fixed;left:-9999px;top:-9999px;width:1px;height:1px;border:none;opacity:0;pointer-events:none';
        document.body.appendChild(fr);
        const doc = fr.contentDocument || fr.contentWindow.document;
        doc.open();
        doc.write(html);
        doc.close();
        setTimeout(() => { fr.contentWindow.focus(); fr.contentWindow.print(); }, 600);
    }

    function showPrintPreview(html) {
        const modal = document.getElementById('printPreviewModal');
        const frame = document.getElementById('ppFrame');
        if (!modal || !frame) { return; }
        modal.classList.remove('hidden');
        // Sync native overlay colour to the print preview toolbar
        window.electronAPI?.setTitleBarOverlay?.({ color: '#1e293b', symbolColor: '#94a3b8', height: 38 });
        const doc = frame.contentDocument || frame.contentWindow.document;
        doc.open();
        doc.write(html);
        doc.close();
    }

    function closePrintPreview() {
        document.getElementById('printPreviewModal')?.classList.add('hidden');
        // Restore overlay to match active app theme
        const t = document.documentElement.getAttribute('data-theme') || 'dark';
        window.electronAPI?.setTitleBarOverlay?.(
            t === 'dark'
                ? { color: '#13132e', symbolColor: '#9ca3af', height: 38 }
                : { color: '#312e81', symbolColor: '#c7d2fe', height: 38 }
        );
    }

    // ── Print Preview modal controls ──
    const _ppDoPrint = () => { const fr = document.getElementById('ppFrame'); if (fr) { fr.contentWindow.focus(); fr.contentWindow.print(); } };
    document.getElementById('ppCloseBtn')?.addEventListener('click', closePrintPreview);
    document.getElementById('ppCancelBtn')?.addEventListener('click', closePrintPreview);
    document.getElementById('ppPrintBtn')?.addEventListener('click', _ppDoPrint);
    document.getElementById('ppPrintBtnBottom')?.addEventListener('click', _ppDoPrint);
    document.addEventListener('keydown', e => {
        if (e.key === 'Escape') {
            const modal = document.getElementById('printPreviewModal');
            if (modal && !modal.classList.contains('hidden')) closePrintPreview();
        }
    });

    // ── Theme ──────────────────────────────────────────────
    const htmlEl = document.documentElement;
    const themeSwitch = document.getElementById('themeSwitch');
    const themeSwitch2 = document.getElementById('themeSwitch2');
    const savedTheme = localStorage.getItem('spliqour-theme') || 'dark';
    applyTheme(savedTheme);

    [themeSwitch, themeSwitch2].forEach(sw => {
        if (sw) sw.addEventListener('change', () => applyTheme(
            htmlEl.getAttribute('data-theme') === 'dark' ? 'light' : 'dark'
        ));
    });

    function applyTheme(t) {
        htmlEl.setAttribute('data-theme', t);
        if (themeSwitch) themeSwitch.checked = (t === 'dark');
        if (themeSwitch2) themeSwitch2.checked = (t === 'dark');
        localStorage.setItem('spliqour-theme', t);
        // Update native title-bar overlay to match topbar colour
        if (window.electronAPI && window.electronAPI.setTitleBarOverlay) {
            window.electronAPI.setTitleBarOverlay(
                t === 'dark'
                    ? { color: '#13132e', symbolColor: '#9ca3af', height: 38 }
                    : { color: '#312e81', symbolColor: '#c7d2fe', height: 38 }
            );
        }
    }

    // ── Home button ────────────────────────────────────────
    const homeBtn = document.getElementById('homeBtn');
    if (homeBtn) homeBtn.addEventListener('click', () => window.electronAPI.navigateHome());

    const changeClientBtn = document.getElementById('changeClientBtn');
    if (changeClientBtn) changeClientBtn.addEventListener('click', () => window.electronAPI.navigateHome());
    const subChangeClientBtn = document.getElementById('subChangeClientBtn');
    if (subChangeClientBtn) subChangeClientBtn.addEventListener('click', () => window.electronAPI.navigateHome());

    // ── Load active bar from sessionStorage ───────────────
    let activeBar = {};
    try {
        const raw = sessionStorage.getItem('active-bar');
        if (raw) {
            activeBar = JSON.parse(raw);
            sessionStorage.removeItem('active-bar');
        }
    } catch (e) { activeBar = {}; }

    // ── Beer Shopee restriction flag ───────────────────────
    const isBeerShopee = activeBar.shopType === 'FL-BR-II';
    if (isBeerShopee) {
        // Disable Spirits option in all static category selects
        document.querySelectorAll('select option[value="Spirits"]').forEach(opt => {
            opt.disabled = true;
            opt.style.cssText = 'color:#888;font-style:italic';
            opt.textContent = 'Spirits (N/A – Beer Shopee)';
        });
    }

    // ── Populate bar name pill ─────────────────────────────
    const nameEl = document.getElementById('activeBarName');
    const fyEl = document.getElementById('activeBarFY');
    const fyLabel = activeBar.financialYear || '';
    const fyShort = fyLabel.match(/(\d{4})/g)
        ? `FY ${fyLabel.match(/(\d{4})/g)[0]}-${fyLabel.match(/(\d{4})/g)[1].slice(2)}`
        : fyLabel;

    if (nameEl) nameEl.textContent = activeBar.barName || 'Unknown Bar';
    if (fyEl) fyEl.textContent = fyShort;
    const subBarName = document.getElementById('subBarName');
    const subBarFY = document.getElementById('subBarFY');
    if (subBarName) subBarName.textContent = activeBar.barName || 'Unknown Bar';
    if (subBarFY) subBarFY.textContent = fyShort;
    document.title = `SPLIQOUR PRO — ${activeBar.barName || 'App'}`;

    // ── Populate Dashboard ─────────────────────────────────
    const dashBarName = document.getElementById('dashBarName');
    if (dashBarName) dashBarName.textContent = activeBar.barName || 'Dashboard';

    // Set today's date
    const dashTodayDate = document.getElementById('dashTodayDate');
    if (dashTodayDate) {
        const now = new Date();
        const options = { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' };
        dashTodayDate.textContent = now.toLocaleDateString('en-IN', options);
    }

    // Populate bar info grid
    const infoFields = [
        ['Bar Name', activeBar.barName],
        ['License', activeBar.licenseNo],
        ['License Type', activeBar.shopType],
        ['Owner', activeBar.ownerName],
        ['Phone', activeBar.phone],
        ['Email', activeBar.email],
        ['Address', [activeBar.address, activeBar.area, activeBar.city, activeBar.state, activeBar.pinCode].filter(Boolean).join(', ')],
        ['GSTIN', activeBar.gstin],
        ['PAN', activeBar.pan],
        ['FSSAI No.', activeBar.fssaiNo],
        ['Bank', activeBar.bankName],
        ['Account No.', activeBar.accountNo],
        ['IFSC', activeBar.ifsc],
        ['Financial Year', fyLabel],
    ];

    const infoTable = document.getElementById('barInfoTable');
    if (infoTable) {
        infoFields.forEach(([key, val]) => {
            if (!val) return;
            const isAddress = key === 'Address';
            const item = document.createElement('div');
            item.className = 'dash-info-item' + (isAddress ? ' full-width' : '');
            item.innerHTML = `<span class="dash-info-key">${key}</span><span class="dash-info-val">${val}</span>`;
            infoTable.appendChild(item);
        });
    }

    // Wire up edit bar profile button on dashboard
    const editBarProfileBtnDash = document.getElementById('editBarProfileBtnDash');
    if (editBarProfileBtnDash) {
        editBarProfileBtnDash.addEventListener('click', () => {
            sessionStorage.setItem('active-bar', JSON.stringify(activeBar));
            window.electronAPI.navigateToAddBar();
        });
    }

    // ── Populate Settings ──────────────────────────────────
    const settingFY = document.getElementById('settingFY');
    const settingFolder = document.getElementById('settingFolder');
    if (settingFY) settingFY.textContent = fyShort || '—';
    // Populate folder from actual data root IPC call
    (async () => {
        if (!settingFolder) return;
        try {
            const r = await window.electronAPI.getDataRoot();
            if (r.success) settingFolder.textContent = r.dataRoot;
        } catch (_) {
            settingFolder.textContent = `data/${activeBar.barName?.replace(/\s+/g, '_') || ''}/${fyShort?.replace(' ', '_') || ''}`;
        }
    })();

    // Edit bar profile → navigate to Add Bar form with bar data pre-loaded
    const editBarProfileBtn = document.getElementById('editBarProfileBtn');
    if (editBarProfileBtn) {
        editBarProfileBtn.addEventListener('click', () => {
            sessionStorage.setItem('active-bar', JSON.stringify(activeBar));
            window.electronAPI.navigateToAddBar();
        });
    }

    // ── Settings Tab Navigation ────────────────────────────
    const settingsNav = document.getElementById('settingsNav');
    if (settingsNav) {
        const settingsNavItems = settingsNav.querySelectorAll('.settings-nav-item');
        const settingsPanes = document.querySelectorAll('.settings-pane');

        settingsNavItems.forEach(btn => {
            btn.addEventListener('click', () => {
                const tab = btn.dataset.settingsTab;
                settingsNavItems.forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                settingsPanes.forEach(p => {
                    p.classList.toggle('active', p.id === `settingsPane-${tab}`);
                });
            });
        });
    }

    // Populate backup & transfer info
    const backupBarNameEl = document.getElementById('backupBarName');
    if (backupBarNameEl) backupBarNameEl.textContent = activeBar.barName || '—';

    const transferFromFY = document.getElementById('transferFromFY');
    const transferToFY = document.getElementById('transferToFY');
    if (transferFromFY) transferFromFY.textContent = fyShort || '—';
    if (transferToFY) {
        // Calculate next FY
        const fyMatch = (activeBar.financialYear || '').match(/(\d{4})/g);
        if (fyMatch && fyMatch.length >= 2) {
            const nextStart = parseInt(fyMatch[1]);
            const nextEnd = nextStart + 1;
            transferToFY.textContent = `FY ${nextStart}-${String(nextEnd).slice(2)}`;
        } else {
            transferToFY.textContent = '—';
        }
    }

    // Backup button
    const backupBtn = document.getElementById('backupBtn');
    const backupStatus = document.getElementById('backupStatus');
    const lastBackupTimeEl = document.getElementById('lastBackupTime');
    if (backupBtn) {
        // Show last backup from localStorage
        const lastBackup = localStorage.getItem('spliqour-last-backup');
        if (lastBackup && lastBackupTimeEl) lastBackupTimeEl.textContent = lastBackup;

        backupBtn.addEventListener('click', async () => {
            backupBtn.disabled = true;
            backupBtn.textContent = 'Creating backup...';
            try {
                const result = await window.electronAPI.backupData({
                    barName: activeBar.barName,
                    financialYear: activeBar.financialYear
                });
                if (result.success) {
                    const now = new Date().toLocaleString();
                    localStorage.setItem('spliqour-last-backup', now);
                    if (lastBackupTimeEl) lastBackupTimeEl.textContent = now;
                    if (backupStatus) {
                        backupStatus.className = 'settings-status-msg success';
                        backupStatus.textContent = `Backup saved to: ${result.path}`;
                    }
                } else {
                    if (backupStatus) {
                        backupStatus.className = 'settings-status-msg error';
                        backupStatus.textContent = result.error || 'Backup failed';
                    }
                }
            } catch (err) {
                if (backupStatus) {
                    backupStatus.className = 'settings-status-msg error';
                    backupStatus.textContent = 'Backup failed: ' + err.message;
                }
            }
            backupBtn.disabled = false;
            backupBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Create Backup Now`;
        });
    }

    // ── App & Data pane ────────────────────────────────────
    // Load & display current app version
    (async () => {
        const verEl = document.getElementById('appVersionBadge');
        if (verEl) {
            try {
                const r = await window.electronAPI.getAppVersion();
                verEl.textContent = 'v' + (r.version || '—');
            } catch (_) { verEl.textContent = 'v—'; }
        }
    })();

    // Load & display current data directory
    const dataDirPathEl = document.getElementById('dataDirPath');
    (async () => {
        if (!dataDirPathEl) return;
        try {
            const r = await window.electronAPI.getDataRoot();
            if (r.success) dataDirPathEl.textContent = r.dataRoot;
        } catch (_) { dataDirPathEl.textContent = '—'; }
    })();

    // Open data folder in Explorer
    const openDataFolderBtn = document.getElementById('openDataFolderBtn');
    if (openDataFolderBtn) {
        openDataFolderBtn.addEventListener('click', async () => {
            await window.electronAPI.openDataFolder();
        });
    }

    // Change data folder
    const changeDataFolderBtn = document.getElementById('changeDataFolderBtn');
    const dataDirStatus = document.getElementById('dataDirStatus');
    const dataDirChangeWarning = document.getElementById('dataDirChangeWarning');
    if (changeDataFolderBtn) {
        changeDataFolderBtn.addEventListener('click', async () => {
            changeDataFolderBtn.disabled = true;
            changeDataFolderBtn.textContent = 'Select folder…';
            try {
                const r = await window.electronAPI.chooseDataFolder();
                if (r.success) {
                    if (dataDirPathEl) dataDirPathEl.textContent = r.dataRoot;
                    // Also update the general-settings Data Folder badge
                    const settingFolder = document.getElementById('settingFolder');
                    if (settingFolder) settingFolder.textContent = r.dataRoot;
                    if (dataDirChangeWarning) dataDirChangeWarning.style.display = 'flex';
                    if (dataDirStatus) {
                        dataDirStatus.className = 'settings-status-msg success';
                        dataDirStatus.textContent = 'Data folder updated successfully.';
                    }
                } else if (r.error && r.error !== 'Cancelled') {
                    if (dataDirStatus) {
                        dataDirStatus.className = 'settings-status-msg error';
                        dataDirStatus.textContent = r.error;
                    }
                }
            } catch (err) {
                if (dataDirStatus) {
                    dataDirStatus.className = 'settings-status-msg error';
                    dataDirStatus.textContent = 'Failed: ' + err.message;
                }
            }
            changeDataFolderBtn.disabled = false;
            changeDataFolderBtn.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg> Change Folder`;
        });
    }

    // Check for Updates
    const checkUpdateBtn = document.getElementById('checkUpdateBtn');
    const updateStatusBox = document.getElementById('updateStatusBox');
    const updateStatusText = document.getElementById('updateStatusText');
    if (checkUpdateBtn) {
        // Listen for download progress from main process
        window.electronAPI.onUpdateDownloadProgress((progress) => {
            const pct = Math.round(progress.percent || 0);
            if (updateStatusText) updateStatusText.textContent = `Downloading… ${pct}% (${Math.round((progress.bytesPerSecond || 0) / 1024)} KB/s)`;
            if (updateStatusBox) updateStatusBox.className = 'appdata-update-status appdata-update-status--checking';
        });

        // Listen for download complete from main process
        window.electronAPI.onUpdateDownloaded(() => {
            if (updateStatusBox) updateStatusBox.className = 'appdata-update-status appdata-update-status--ok';
            if (updateStatusText) updateStatusText.textContent = 'Download complete! Click Install & Restart to apply the update.';
            checkUpdateBtn.disabled = false;
            checkUpdateBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="8 17 12 21 16 17"/><line x1="12" y1="3" x2="12" y2="21"/></svg> Install & Restart`;
            checkUpdateBtn.onclick = () => window.electronAPI.installUpdate();
        });

        checkUpdateBtn.addEventListener('click', async () => {
            checkUpdateBtn.disabled = true;
            checkUpdateBtn.textContent = 'Checking…';
            if (updateStatusBox) updateStatusBox.className = 'appdata-update-status appdata-update-status--checking';
            if (updateStatusText) updateStatusText.textContent = 'Checking for updates…';
            try {
                const r = await window.electronAPI.checkForUpdates();
                if (updateStatusBox) {
                    updateStatusBox.className = 'appdata-update-status ' +
                        (r.updateAvailable ? 'appdata-update-status--error' : 'appdata-update-status--ok');
                }
                if (updateStatusText) updateStatusText.textContent = r.message || 'Check complete.';
                if (r.updateAvailable) {
                    checkUpdateBtn.disabled = false;
                    checkUpdateBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="8 17 12 21 16 17"/><line x1="12" y1="3" x2="12" y2="21"/></svg> Download & Install`;
                    checkUpdateBtn.onclick = async () => {
                        checkUpdateBtn.disabled = true;
                        checkUpdateBtn.textContent = 'Downloading…';
                        if (updateStatusBox) updateStatusBox.className = 'appdata-update-status appdata-update-status--checking';
                        if (updateStatusText) updateStatusText.textContent = 'Starting download…';
                        await window.electronAPI.downloadUpdate();
                    };
                    return;
                }
            } catch (err) {
                if (updateStatusBox) updateStatusBox.className = 'appdata-update-status appdata-update-status--error';
                if (updateStatusText) updateStatusText.textContent = 'Error: ' + err.message;
            }
            checkUpdateBtn.disabled = false;
            checkUpdateBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-3.46"/></svg> Check for Updates`;
        });
    }

    // Transfer FY button
    const transferFYBtn = document.getElementById('transferFYBtn');
    const transferStatus = document.getElementById('transferStatus');
    if (transferFYBtn) {
        transferFYBtn.addEventListener('click', async () => {
            const confirmed = confirm(
                'Are you sure you want to transfer data to the next Financial Year?\n\n' +
                'This will create a new FY folder with closing stock carried forward as opening stock.\n' +
                'Please ensure you have created a backup first.'
            );
            if (!confirmed) return;

            transferFYBtn.disabled = true;
            transferFYBtn.textContent = 'Transferring...';
            try {
                const result = await window.electronAPI.transferToNextFY({
                    barName: activeBar.barName,
                    financialYear: activeBar.financialYear
                });
                if (result.success) {
                    if (transferStatus) {
                        transferStatus.className = 'settings-status-msg success';
                        transferStatus.textContent = `Successfully created ${result.newFY}. Reload to switch.`;
                    }
                } else {
                    if (transferStatus) {
                        transferStatus.className = 'settings-status-msg error';
                        transferStatus.textContent = result.error || 'Transfer failed';
                    }
                }
            } catch (err) {
                if (transferStatus) {
                    transferStatus.className = 'settings-status-msg error';
                    transferStatus.textContent = 'Transfer failed: ' + err.message;
                }
            }
            transferFYBtn.disabled = false;
            transferFYBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 0 1 4-4h14"/><polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 0 1-4 4H3"/></svg> Transfer & Create Next F.Y.`;
        });
    }

    // Master → Bar Profile card
    const mcBarProfile = document.getElementById('mcBarProfile');
    if (mcBarProfile) {
        mcBarProfile.addEventListener('click', () => {
            sessionStorage.setItem('active-bar', JSON.stringify(activeBar));
            window.electronAPI.navigateToAddBar();
        });
    }

    // ── Top Navigation + Sliding Indicator ──────────────
    const topNav = document.getElementById('topNav');
    const navBtns = Array.from(document.querySelectorAll('#topNav .tb-tab'));
    const panels = Array.from(document.querySelectorAll('.view-panel'));
    const indicator = document.getElementById('navIndicator');
    const statusMsg = document.getElementById('appStatusMsg');
    const purchaseWrapper = document.getElementById('purchaseDropdown');
    const purchaseBtn = document.getElementById('nav-purchase');
    const purchaseMenu = document.getElementById('purchaseMenu');
    // purchaseSubNav removed
    const purchaseItems = purchaseMenu ? Array.from(purchaseMenu.querySelectorAll('.dd-item')) : [];

    // Sales dropdown elements
    const salesWrapper = document.getElementById('salesDropdown');
    const salesBtn = document.getElementById('nav-sales');
    const salesMenu = document.getElementById('salesMenu');
    const salesItems = salesMenu ? Array.from(salesMenu.querySelectorAll('.dd-item')) : [];

    // Master dropdown elements
    const masterWrapper = document.getElementById('masterDropdown');
    const masterBtn = document.getElementById('nav-master');
    const masterMenu = document.getElementById('masterMenu');
    // masterSubNav removed
    const masterItems = masterMenu ? Array.from(masterMenu.querySelectorAll('.dd-item')) : [];

    // Report dropdown elements
    const reportWrapper = document.getElementById('reportDropdown');
    const reportBtn = document.getElementById('nav-report');
    const reportMenu = document.getElementById('reportMenu');
    const reportItems = reportMenu ? Array.from(reportMenu.querySelectorAll('.dd-item')) : [];

    // Online dropdown elements
    const onlineWrapper = document.getElementById('onlineDropdown');
    const onlineBtn = document.getElementById('nav-online');
    const onlineMenu = document.getElementById('onlineMenu');
    const onlineItems = onlineMenu ? Array.from(onlineMenu.querySelectorAll('.dd-item')) : [];

    let activeIdx = Math.max(0, navBtns.findIndex(btn => btn.classList.contains('active')));

    panels.forEach(panel => panel.setAttribute('role', 'tabpanel'));

    /* ── Theme button (replaces old toggle) ── */
    const themeBtn = document.getElementById('themeBtn');
    if (themeBtn) {
        themeBtn.addEventListener('click', () => {
            applyTheme(htmlEl.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
        });
    }

    /* ── Sliding indicator helpers ── */
    function moveIndicator(btn) {
        if (!indicator || !btn) return;
        const track = btn.closest('.tb-nav-track');
        if (!track) return;
        const trackRect = track.getBoundingClientRect();
        const btnRect = btn.getBoundingClientRect();
        indicator.style.left = (btnRect.left - trackRect.left) + 'px';
        indicator.style.width = btnRect.width + 'px';
    }

    // Position indicator on first load
    requestAnimationFrame(() => moveIndicator(navBtns[activeIdx]));
    window.addEventListener('resize', () => moveIndicator(navBtns[activeIdx]));

    function updateStatus(text) {
        if (statusMsg) statusMsg.textContent = text;
    }

    function setActiveNav(idx) {
        navBtns.forEach((btn, i) => {
            const isActive = i === idx;
            btn.classList.toggle('active', isActive);
            btn.setAttribute('aria-selected', String(isActive));
            btn.setAttribute('tabindex', isActive ? '0' : '-1');
        });
        activeIdx = idx;
        moveIndicator(navBtns[idx]);
    }

    function closePurchaseMenu({ focusButton = false } = {}) {
        if (!purchaseWrapper || !purchaseBtn) return;
        purchaseWrapper.classList.remove('open');
        purchaseBtn.setAttribute('aria-expanded', 'false');
        if (focusButton) purchaseBtn.focus();
    }

    function openPurchaseMenu({ focusFirstItem = false } = {}) {
        if (!purchaseWrapper || !purchaseBtn) return;
        closeSalesMenu();  // close sibling
        closeMasterMenu(); // close sibling dropdown
        closeReportMenu(); // close sibling dropdown
        closeOnlineMenu(); // close sibling dropdown
        purchaseWrapper.classList.add('open');
        purchaseBtn.setAttribute('aria-expanded', 'true');
        if (focusFirstItem && purchaseItems.length) {
            purchaseItems[0].focus();
        }
    }

    // ── Sales dropdown helpers ──
    function closeSalesMenu({ focusButton = false } = {}) {
        if (!salesWrapper || !salesBtn) return;
        salesWrapper.classList.remove('open');
        salesBtn.setAttribute('aria-expanded', 'false');
        if (focusButton) salesBtn.focus();
    }

    function openSalesMenu({ focusFirstItem = false } = {}) {
        if (!salesWrapper || !salesBtn) return;
        closePurchaseMenu(); // close sibling
        closeMasterMenu();   // close sibling
        closeReportMenu();   // close sibling
        closeOnlineMenu();   // close sibling
        salesWrapper.classList.add('open');
        salesBtn.setAttribute('aria-expanded', 'true');
        if (focusFirstItem && salesItems.length) {
            salesItems[0].focus();
        }
    }

    function switchSalesSub(subId) {
        const salesIdx = navBtns.findIndex(b => b.dataset.view === 'sales');
        if (salesIdx >= 0) {
            switchView('sales', salesIdx, { preserveStatus: true });
        }

        document.querySelectorAll('.sales-sub-view').forEach(sv => sv.classList.remove('active'));
        const target = document.getElementById('sub-' + subId);
        if (target) target.classList.add('active');

        salesItems.forEach(item => {
            item.classList.toggle('active', item.dataset.sub === subId);
            item.setAttribute('aria-current', item.dataset.sub === subId ? 'true' : 'false');
        });

        // Reload data when navigating to specific sub-views
        if (subId === 'customer-manager' && typeof loadCustomers === 'function') {
            loadCustomers();
        }
        if (subId === 'add-daily-sale' && typeof initDailySale === 'function') {
            initDailySale();
        }
        if (subId === 'billing' && typeof initBilling === 'function') {
            initBilling();
        }
        if (subId === 'sale-summary' && typeof initSaleSummary === 'function') {
            initSaleSummary();
        }
        if (subId === 'vat-tax-report' && typeof initVatReport === 'function') {
            initVatReport();
        }

        updateStatus('Sales › ' + subId.replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase()));
    }

    // ── Master dropdown helpers ──
    function closeMasterMenu({ focusButton = false } = {}) {
        if (!masterWrapper || !masterBtn) return;
        masterWrapper.classList.remove('open');
        masterBtn.setAttribute('aria-expanded', 'false');
        if (focusButton) masterBtn.focus();
    }

    function openMasterMenu({ focusFirstItem = false } = {}) {
        if (!masterWrapper || !masterBtn) return;
        closePurchaseMenu(); // close sibling dropdown
        closeSalesMenu();    // close sibling
        closeReportMenu();   // close sibling
        closeOnlineMenu();   // close sibling
        masterWrapper.classList.add('open');
        masterBtn.setAttribute('aria-expanded', 'true');
        if (focusFirstItem && masterItems.length) {
            masterItems[0].focus();
        }
    }

    // ── Report dropdown helpers ──
    function closeReportMenu({ focusButton = false } = {}) {
        if (!reportWrapper || !reportBtn) return;
        reportWrapper.classList.remove('open');
        reportBtn.setAttribute('aria-expanded', 'false');
        if (focusButton) reportBtn.focus();
    }

    function openReportMenu({ focusFirstItem = false } = {}) {
        if (!reportWrapper || !reportBtn) return;
        closePurchaseMenu();
        closeSalesMenu();
        closeMasterMenu();
        closeOnlineMenu();
        reportWrapper.classList.add('open');
        reportBtn.setAttribute('aria-expanded', 'true');
        if (focusFirstItem && reportItems.length) {
            reportItems[0].focus();
        }
    }

    function switchReportSub(subId) {
        const reportIdx = navBtns.findIndex(b => b.dataset.view === 'report');
        if (reportIdx >= 0) {
            switchView('report', reportIdx, { preserveStatus: true });
        }

        document.querySelectorAll('.report-sub-view').forEach(sv => sv.classList.remove('active'));
        const target = document.getElementById('sub-' + subId);
        if (target) target.classList.add('active');

        reportItems.forEach(item => {
            item.classList.toggle('active', item.dataset.sub === subId);
            item.setAttribute('aria-current', item.dataset.sub === subId ? 'true' : 'false');
        });

        updateStatus('Report › ' + subId.replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase()));

        // Show date picker modal when opening brandwise
        if (subId === 'brandwise' && typeof window._showBwDateModal === 'function') {
            setTimeout(window._showBwDateModal, 80);
        }
        // Show date range modal when opening chatai register
        if (subId === 'chatai-register' && typeof window._showCrDateModal === 'function') {
            setTimeout(window._showCrDateModal, 80);
        }
        // Show date range modal when opening MML chatai
        if (subId === 'mml-chatai' && typeof window._showMcDateModal === 'function') {
            setTimeout(window._showMcDateModal, 80);
        }
    }

    // ── Online dropdown helpers ──
    function closeOnlineMenu({ focusButton = false } = {}) {
        if (!onlineWrapper || !onlineBtn) return;
        onlineWrapper.classList.remove('open');
        onlineBtn.setAttribute('aria-expanded', 'false');
        if (focusButton) onlineBtn.focus();
    }

    function openOnlineMenu({ focusFirstItem = false } = {}) {
        if (!onlineWrapper || !onlineBtn) return;
        closePurchaseMenu();
        closeSalesMenu();
        closeMasterMenu();
        closeReportMenu();
        onlineWrapper.classList.add('open');
        onlineBtn.setAttribute('aria-expanded', 'true');
        if (focusFirstItem && onlineItems.length) {
            onlineItems[0].focus();
        }
    }

    function switchOnlineSub(subId) {
        const onlineIdx = navBtns.findIndex(b => b.dataset.view === 'online');
        if (onlineIdx >= 0) {
            switchView('online', onlineIdx, { preserveStatus: true });
        }

        document.querySelectorAll('.online-sub-view').forEach(sv => sv.classList.remove('active'));
        const target = document.getElementById('sub-' + subId);
        if (target) target.classList.add('active');

        onlineItems.forEach(item => {
            item.classList.toggle('active', item.dataset.sub === subId);
            item.setAttribute('aria-current', item.dataset.sub === subId ? 'true' : 'false');
        });

        updateStatus('Online › ' + subId.replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase()));
    }

    function switchMasterSub(subId) {
        const masterIdx = navBtns.findIndex(b => b.dataset.view === 'master');
        if (masterIdx >= 0) {
            switchView('master', masterIdx, { preserveStatus: true });
        }

        document.querySelectorAll('.master-sub-view').forEach(sv => sv.classList.remove('active'));
        const target = document.getElementById('sub-' + subId);
        if (target) target.classList.add('active');



        masterItems.forEach(item => {
            item.classList.toggle('active', item.dataset.sub === subId);
            item.setAttribute('aria-current', item.dataset.sub === subId ? 'true' : 'false');
        });

        updateStatus('Master › ' + subId.replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase()));
    }

    function switchView(viewId, idx, { preserveStatus = false } = {}) {
        panels.forEach(p => p.classList.remove('active'));
        const panel = document.getElementById('view-' + viewId);
        if (panel) panel.classList.add('active');
        setActiveNav(idx);
        if (!preserveStatus) {
            updateStatus(navBtns[idx].textContent.trim());
        }
    }

    function switchPurchaseSub(subId) {
        const purchaseIdx = navBtns.findIndex(b => b.dataset.view === 'purchase');
        if (purchaseIdx >= 0) {
            switchView('purchase', purchaseIdx, { preserveStatus: true });
        }

        const panel = document.getElementById('view-purchase');
        if (panel) panel.querySelectorAll('.sub-view').forEach(sv => sv.classList.remove('active'));
        const target = document.getElementById('sub-' + subId);
        if (target) target.classList.add('active');



        purchaseItems.forEach(item => {
            item.classList.toggle('active', item.dataset.sub === subId);
            item.setAttribute('aria-current', item.dataset.sub === subId ? 'true' : 'false');
        });

        // Load data when navigating to purchase-report
        if (subId === 'purchase-report' && typeof window._loadPurchaseReport === 'function') {
            window._loadPurchaseReport();
        }

        updateStatus('Purchase › ' + subId.replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase()));
    }

    /* ── Wire nav tab clicks & keyboard ── */
    navBtns.forEach((btn, idx) => {
        btn.addEventListener('click', (e) => {
            const isPurchase = btn.dataset.view === 'purchase';
            const isSales    = btn.dataset.view === 'sales';
            const isMaster   = btn.dataset.view === 'master';
            const isReport   = btn.dataset.view === 'report';
            const isOnline   = btn.dataset.view === 'online';
            if (isPurchase) {
                e.stopPropagation();
                closeSalesMenu();
                closeMasterMenu();
                closeReportMenu();
                closeOnlineMenu();
                if (purchaseWrapper?.classList.contains('open')) {
                    closePurchaseMenu();
                } else {
                    setActiveNav(idx);
                    openPurchaseMenu();
                }
                return;
            }
            if (isSales) {
                e.stopPropagation();
                closePurchaseMenu();
                closeMasterMenu();
                closeReportMenu();
                closeOnlineMenu();
                if (salesWrapper?.classList.contains('open')) {
                    closeSalesMenu();
                } else {
                    setActiveNav(idx);
                    openSalesMenu();
                }
                return;
            }
            if (isMaster) {
                e.stopPropagation();
                closePurchaseMenu();
                closeSalesMenu();
                closeReportMenu();
                closeOnlineMenu();
                if (masterWrapper?.classList.contains('open')) {
                    closeMasterMenu();
                } else {
                    setActiveNav(idx);
                    openMasterMenu();
                }
                return;
            }
            if (isReport) {
                e.stopPropagation();
                closePurchaseMenu();
                closeSalesMenu();
                closeMasterMenu();
                closeOnlineMenu();
                if (reportWrapper?.classList.contains('open')) {
                    closeReportMenu();
                } else {
                    setActiveNav(idx);
                    openReportMenu();
                }
                return;
            }
            if (isOnline) {
                e.stopPropagation();
                closePurchaseMenu();
                closeSalesMenu();
                closeMasterMenu();
                closeReportMenu();
                if (onlineWrapper?.classList.contains('open')) {
                    closeOnlineMenu();
                } else {
                    setActiveNav(idx);
                    openOnlineMenu();
                }
                return;
            }
            closePurchaseMenu();
            closeSalesMenu();
            closeMasterMenu();
            closeReportMenu();
            closeOnlineMenu();
            switchView(btn.dataset.view, idx);
        });

        btn.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
                e.preventDefault();
                closePurchaseMenu();
                closeSalesMenu();
                closeMasterMenu();
                closeReportMenu();
                closeOnlineMenu();
                const dir = e.key === 'ArrowRight' ? 1 : -1;
                const next = (idx + dir + navBtns.length) % navBtns.length;
                const nextBtn = navBtns[next];
                nextBtn.focus();
                // For tabs with dropdowns, just highlight — don't open sub-view
                if (nextBtn.dataset.view === 'purchase' || nextBtn.dataset.view === 'sales' || nextBtn.dataset.view === 'master' || nextBtn.dataset.view === 'report' || nextBtn.dataset.view === 'online') {
                    setActiveNav(next);
                } else {
                    switchView(nextBtn.dataset.view, next);
                }
            } else if (e.key === 'Home') {
                e.preventDefault();
                closePurchaseMenu();
                closeSalesMenu();
                closeMasterMenu();
                closeReportMenu();
                closeOnlineMenu();
                navBtns[0].focus();
                switchView(navBtns[0].dataset.view, 0);
            } else if (e.key === 'End') {
                e.preventDefault();
                closePurchaseMenu();
                closeSalesMenu();
                closeMasterMenu();
                closeReportMenu();
                closeOnlineMenu();
                const last = navBtns.length - 1;
                navBtns[last].focus();
                switchView(navBtns[last].dataset.view, last);
            } else if (btn.dataset.view === 'purchase' && (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ')) {
                e.preventDefault();
                openPurchaseMenu({ focusFirstItem: true });
            } else if (btn.dataset.view === 'purchase' && e.key === 'Escape') {
                e.preventDefault();
                closePurchaseMenu({ focusButton: true });
            } else if (btn.dataset.view === 'sales' && (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ')) {
                e.preventDefault();
                openSalesMenu({ focusFirstItem: true });
            } else if (btn.dataset.view === 'sales' && e.key === 'Escape') {
                e.preventDefault();
                closeSalesMenu({ focusButton: true });
            } else if (btn.dataset.view === 'master' && (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ')) {
                e.preventDefault();
                openMasterMenu({ focusFirstItem: true });
            } else if (btn.dataset.view === 'master' && e.key === 'Escape') {
                e.preventDefault();
                closeMasterMenu({ focusButton: true });
            } else if (btn.dataset.view === 'report' && (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ')) {
                e.preventDefault();
                openReportMenu({ focusFirstItem: true });
            } else if (btn.dataset.view === 'report' && e.key === 'Escape') {
                e.preventDefault();
                closeReportMenu({ focusButton: true });
            } else if (btn.dataset.view === 'online' && (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ')) {
                e.preventDefault();
                openOnlineMenu({ focusFirstItem: true });
            } else if (btn.dataset.view === 'online' && e.key === 'Escape') {
                e.preventDefault();
                closeOnlineMenu({ focusButton: true });
            }
        });
    });

    /* ── Wire dropdown items ── */
    purchaseItems.forEach((item, itemIdx) => {
        item.addEventListener('click', (e) => {
            e.stopPropagation();
            switchPurchaseSub(item.dataset.sub);
            closePurchaseMenu({ focusButton: true });
        });

        item.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                purchaseItems[(itemIdx + 1) % purchaseItems.length].focus();
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                purchaseItems[(itemIdx - 1 + purchaseItems.length) % purchaseItems.length].focus();
            } else if (e.key === 'Home') {
                e.preventDefault();
                purchaseItems[0].focus();
            } else if (e.key === 'End') {
                e.preventDefault();
                purchaseItems[purchaseItems.length - 1].focus();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                closePurchaseMenu({ focusButton: true });
            } else if (e.key === 'Tab') {
                closePurchaseMenu();
            } else if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                item.click();
            }
        });
    });

    /* ── Wire Master dropdown items ── */
    masterItems.forEach((item, itemIdx) => {
        item.addEventListener('click', (e) => {
            e.stopPropagation();
            switchMasterSub(item.dataset.sub);
            closeMasterMenu({ focusButton: true });
        });

        item.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                masterItems[(itemIdx + 1) % masterItems.length].focus();
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                masterItems[(itemIdx - 1 + masterItems.length) % masterItems.length].focus();
            } else if (e.key === 'Home') {
                e.preventDefault();
                masterItems[0].focus();
            } else if (e.key === 'End') {
                e.preventDefault();
                masterItems[masterItems.length - 1].focus();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                closeMasterMenu({ focusButton: true });
            } else if (e.key === 'Tab') {
                closeMasterMenu();
            } else if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                item.click();
            }
        });
    });

    /* ── Wire Sales dropdown items ── */
    salesItems.forEach((item, itemIdx) => {
        item.addEventListener('click', (e) => {
            e.stopPropagation();
            switchSalesSub(item.dataset.sub);
            closeSalesMenu({ focusButton: true });
        });

        item.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                salesItems[(itemIdx + 1) % salesItems.length].focus();
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                salesItems[(itemIdx - 1 + salesItems.length) % salesItems.length].focus();
            } else if (e.key === 'Home') {
                e.preventDefault();
                salesItems[0].focus();
            } else if (e.key === 'End') {
                e.preventDefault();
                salesItems[salesItems.length - 1].focus();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                closeSalesMenu({ focusButton: true });
            } else if (e.key === 'Tab') {
                closeSalesMenu();
            } else if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                item.click();
            }
        });
    });

    /* ── Wire Report dropdown items ── */
    reportItems.forEach((item, itemIdx) => {
        item.addEventListener('click', (e) => {
            e.stopPropagation();
            switchReportSub(item.dataset.sub);
            closeReportMenu({ focusButton: true });
        });

        item.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                reportItems[(itemIdx + 1) % reportItems.length].focus();
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                reportItems[(itemIdx - 1 + reportItems.length) % reportItems.length].focus();
            } else if (e.key === 'Home') {
                e.preventDefault();
                reportItems[0].focus();
            } else if (e.key === 'End') {
                e.preventDefault();
                reportItems[reportItems.length - 1].focus();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                closeReportMenu({ focusButton: true });
            } else if (e.key === 'Tab') {
                closeReportMenu();
            } else if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                item.click();
            }
        });
    });

    /* ── Wire Online dropdown items ── */
    onlineItems.forEach((item, itemIdx) => {
        item.addEventListener('click', (e) => {
            e.stopPropagation();
            switchOnlineSub(item.dataset.sub);
            closeOnlineMenu({ focusButton: true });
        });

        item.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                onlineItems[(itemIdx + 1) % onlineItems.length].focus();
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                onlineItems[(itemIdx - 1 + onlineItems.length) % onlineItems.length].focus();
            } else if (e.key === 'Home') {
                e.preventDefault();
                onlineItems[0].focus();
            } else if (e.key === 'End') {
                e.preventDefault();
                onlineItems[onlineItems.length - 1].focus();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                closeOnlineMenu({ focusButton: true });
            } else if (e.key === 'Tab') {
                closeOnlineMenu();
            } else if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                item.click();
            }
        });
    });

    /* ── Wire "Open in Browser" buttons for Online portal cards ── */
    document.querySelectorAll('.opc-launch-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const url = btn.dataset.url;
            if (url && window.electronAPI && window.electronAPI.openExternal) {
                window.electronAPI.openExternal(url);
            } else if (url) {
                window.open(url, '_blank');
            }
        });
    });

    /* ── SCM Auto Login ── */
    (function scmAutoLogin() {
        const PWMGR_KEY = 'spliqour-pwmgr';
        const statusEl = document.getElementById('scmCredStatus');
        const autoBtn  = document.getElementById('scmAutoLoginBtn');
        if (!statusEl || !autoBtn) return;

        function findScmCred() {
            try {
                const creds = JSON.parse(localStorage.getItem(PWMGR_KEY) || '[]');
                const barName = (activeBar.barName || '').trim().toLowerCase();
                return creds.find(c =>
                    (c.portalName || '').toLowerCase().includes('scm') &&
                    (c.barName || '').trim().toLowerCase() === barName
                ) || null;
            } catch(e) { return null; }
        }

        function refreshStatus() {
            const cred = findScmCred();
            if (cred) {
                statusEl.className = 'opc-cred-status opc-cred-found';
                statusEl.innerHTML = `
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
                    <span>Auto login ready &middot; <strong>${esc(cred.username)}</strong></span>
                `;
                autoBtn.disabled = false;
            } else {
                statusEl.className = 'opc-cred-status opc-cred-missing';
                statusEl.innerHTML = `
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
                    <span>No SCM credential saved &mdash; add one in <strong>Password Manager</strong> with portal name containing &ldquo;SCM&rdquo;</span>
                `;
                autoBtn.disabled = true;
            }
        }

        refreshStatus();

        // Re-check whenever the Online nav button is clicked (returning to tab)
        document.getElementById('nav-online')?.addEventListener('click', refreshStatus, { passive: true });

        autoBtn.addEventListener('click', async () => {
            const cred = findScmCred();
            if (!cred) { refreshStatus(); return; }
            const url = autoBtn.dataset.url;
            try {
                if (window.electronAPI?.openUrlAutoLogin) {
                    await window.electronAPI.openUrlAutoLogin({ url, username: cred.username, password: cred.password });
                } else {
                    window.open(url, '_blank');
                }
            } catch(e) { console.error('Auto login error', e); }
        });
    })();

    document.addEventListener('click', (e) => {
        if (purchaseWrapper && !purchaseWrapper.contains(e.target)) {
            closePurchaseMenu();
        }
        if (salesWrapper && !salesWrapper.contains(e.target)) {
            closeSalesMenu();
        }
        if (masterWrapper && !masterWrapper.contains(e.target)) {
            closeMasterMenu();
        }
        if (reportWrapper && !reportWrapper.contains(e.target)) {
            closeReportMenu();
        }
        if (onlineWrapper && !onlineWrapper.contains(e.target)) {
            closeOnlineMenu();
        }
    });

    // ── Global Keyboard Shortcuts ──────────────────────────
    document.addEventListener('keydown', (e) => {
        // Theme
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 't') {
            e.preventDefault();
            applyTheme(htmlEl.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
            return;
        }
        // Go Home
        if ((e.ctrlKey || e.metaKey) && e.key === 'Home') {
            e.preventDefault();
            window.electronAPI.navigateHome();
            return;
        }
        // Ctrl+P — Print active brandwise or MML brandwise report
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
            const bwActive  = document.getElementById('sub-brandwise')?.classList.contains('active');
            const mmlActive = document.getElementById('sub-mml-brandwise')?.classList.contains('active');
            if (bwActive || mmlActive) {
                e.preventDefault();
                if (bwActive)  document.getElementById('bwPrintBtn')?.click();
                else           document.getElementById('mmlBwPrintBtn')?.click();
                return;
            }
        }
        // Navigate tabs: Left / Right arrows (when no input/tab is focused)
        if (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA' || document.activeElement.tagName === 'SELECT') return;
        // Skip if a nav tab already has focus — its own keydown handler covers it
        if (document.activeElement.classList.contains('tb-tab')) return;

        if (e.key === 'ArrowRight') {
            e.preventDefault();
            const next = (activeIdx + 1) % navBtns.length;
            closePurchaseMenu();
            closeSalesMenu();
            closeMasterMenu();
            closeReportMenu();
            closeOnlineMenu();
            navBtns[next].click();
            navBtns[next].focus();
        } else if (e.key === 'ArrowLeft') {
            e.preventDefault();
            const prev = (activeIdx - 1 + navBtns.length) % navBtns.length;
            closePurchaseMenu();
            closeSalesMenu();
            closeMasterMenu();
            closeReportMenu();
            closeOnlineMenu();
            navBtns[prev].click();
            navBtns[prev].focus();
        } else if (e.key >= '1' && e.key <= '8') {
            const i = parseInt(e.key) - 1;
            if (navBtns[i]) navBtns[i].click();
        }
    });

    // ═══════════════════════════════════════════════════════
    //  TP VOUCHER ENTRY — Tally Prime-style keyboard logic
    // ═══════════════════════════════════════════════════════

    // Products array shared between Product Master and TP grid autocomplete
    let allProducts = [];

    /* ── Maharashtra standard size list ── */
    const ALL_SIZES = [
        { label: '',                  value: '',                  bpc: 0,   ml: 0    },
        { label: '50 ML',            value: '50 ML',             bpc: 120, ml: 50   },
        { label: '50 ML (192)',      value: '50 ML (192)',       bpc: 192, ml: 50   },
        { label: '50 ML (24)',       value: '50 ML (24)',        bpc: 24,  ml: 50   },
        { label: '50 ML (60)',       value: '50 ML (60)',        bpc: 60,  ml: 50   },
        { label: '50 ML (120)',      value: '50 ML (120)',       bpc: 120, ml: 50   },
        { label: '60 ML',            value: '60 ML',             bpc: 96,  ml: 60   },
        { label: '60 ML (120)',      value: '60 ML (120)',       bpc: 120, ml: 60   },
        { label: '60 ML (75)',       value: '60 ML (75)',        bpc: 75,  ml: 60   },
        { label: '60 ML (Pet)',      value: '60 ML (Pet)',       bpc: 96,  ml: 60   },
        { label: '90 ML (48)',       value: '90 ML (48)',        bpc: 48,  ml: 90   },
        { label: '90 ML (96)',       value: '90 ML (96)',        bpc: 96,  ml: 90   },
        { label: '90 ML (100)',      value: '90 ML (100)',       bpc: 100, ml: 90   },
        { label: '90 ML (Pet)-96',   value: '90 ML (Pet)-96',   bpc: 96,  ml: 90   },
        { label: '90 ML (Pet)-100',  value: '90 ML (Pet)-100',  bpc: 100, ml: 90   },
        { label: '125 ML',           value: '125 ML',            bpc: 72,  ml: 125  },
        { label: '125 ML (72)',      value: '125 ML (72)',       bpc: 72,  ml: 125  },
        { label: '180 ML',           value: '180 ML',            bpc: 48,  ml: 180  },
        { label: '180 ML (Pet)',     value: '180 ML (Pet)',      bpc: 48,  ml: 180  },
        { label: '180 ML (tetra)',   value: '180 ML (tetra)',    bpc: 48,  ml: 180  },
        { label: '180 ML (24)',      value: '180 ML (24)',       bpc: 24,  ml: 180  },
        { label: '180 ML (50)',      value: '180 ML (50)',       bpc: 50,  ml: 180  },
        { label: '187 ML (48)',      value: '187 ML (48)',       bpc: 48,  ml: 187  },
        { label: '200 ML (12)',      value: '200 ML (12)',       bpc: 12,  ml: 200  },
        { label: '200 ML (24)',      value: '200 ML (24)',       bpc: 24,  ml: 200  },
        { label: '200 ML (30)',      value: '200 ML (30)',       bpc: 30,  ml: 200  },
        { label: '200 ML (48)',      value: '200 ML (48)',       bpc: 48,  ml: 200  },
        { label: '250 ML',           value: '250 ML',            bpc: 24,  ml: 250  },
        { label: '250 ML (CAN)',     value: '250 ML (CAN)',      bpc: 24,  ml: 250  },
        { label: '250 ML (24)',      value: '250 ML (24)',       bpc: 24,  ml: 250  },
        { label: '275 ML (24)',      value: '275 ML (24)',       bpc: 24,  ml: 275  },
        { label: '330 ML',           value: '330 ML',            bpc: 24,  ml: 330  },
        { label: '330 ML (CAN)',     value: '330 ML (CAN)',      bpc: 24,  ml: 330  },
        { label: '330 ML (12)',      value: '330 ML (12)',       bpc: 12,  ml: 330  },
        { label: '350 ML (12)',      value: '350 ML (12)',       bpc: 12,  ml: 350  },
        { label: '375 ML',           value: '375 ML',            bpc: 24,  ml: 375  },
        { label: '375 ML (12)',      value: '375 ML (12)',       bpc: 12,  ml: 375  },
        { label: '375 ML (Pet)',     value: '375 ML (Pet)',      bpc: 24,  ml: 375  },
        { label: '500 ML',           value: '500 ML',            bpc: 12,  ml: 500  },
        { label: '500 ML (24)',      value: '500 ML (24)',       bpc: 24,  ml: 500  },
        { label: '500 ML (CAN)',     value: '500 ML (CAN)',      bpc: 24,  ml: 500  },
        { label: '650 ML',           value: '650 ML',            bpc: 12,  ml: 650  },
        { label: '700 ML',           value: '700 ML',            bpc: 12,  ml: 700  },
        { label: '700 ML (6)',       value: '700 ML (6)',        bpc: 6,   ml: 700  },
        { label: '750 ML',           value: '750 ML',            bpc: 12,  ml: 750  },
        { label: '750 ML (Pet)',     value: '750 ML (Pet)',      bpc: 12,  ml: 750  },
        { label: '750 ML (6)',       value: '750 ML (6)',        bpc: 6,   ml: 750  },
        { label: '1000 ML',          value: '1000 ML',           bpc: 9,   ml: 1000 },
        { label: '1000 ML (12)',     value: '1000 ML (12)',      bpc: 12,  ml: 1000 },
        { label: '1000 ML (6)',      value: '1000 ML (6)',       bpc: 6,   ml: 1000 },
        { label: '1000 ML (Pet)',    value: '1000 ML (Pet)',     bpc: 9,   ml: 1000 },
        { label: '1500 ML',          value: '1500 ML',           bpc: 6,   ml: 1500 },
        { label: '1750 ML (6)',      value: '1750 ML (6)',       bpc: 6,   ml: 1750 },
        { label: '2000 ML (4)',      value: '2000 ML (4)',       bpc: 4,   ml: 2000 },
        { label: '2000 ML (6)',      value: '2000 ML (6)',       bpc: 6,   ml: 2000 },
        { label: '2000 ML Pet (6)',  value: '2000 ML Pet (6)',   bpc: 6,   ml: 2000 },
        { label: '2000 ML (Pet)-4',  value: '2000 ML (Pet)-4',  bpc: 4,   ml: 2000 },
        { label: '4500 ML',          value: '4500 ML',           bpc: 4,   ml: 4500 },
        { label: '15 Ltr',           value: '15 Ltr',            bpc: 1,   ml: 15000 },
        { label: '20 Ltr',           value: '20 Ltr',            bpc: 1,   ml: 20000 },
        { label: '30 Ltr',           value: '30 Ltr',            bpc: 1,   ml: 30000 },
        { label: '50 Ltr',           value: '50 Ltr',            bpc: 1,   ml: 50000 },
    ];
    /* Backward-compatible alias */
    const TP_SIZES = ALL_SIZES;
    /* Quick lookup: value → size object */
    const SIZE_LOOKUP = new Map(ALL_SIZES.filter(s => s.value).map(s => [s.value, s]));

    const tpGridBody  = document.getElementById('tpGridBody');
    const tpDateEl    = document.getElementById('tpDate');
    const tpRecvDate  = document.getElementById('tpReceivedDate');
    const tpSaveBtn   = document.getElementById('tpSaveBtn');
    const tpClearBtn  = document.getElementById('tpClearBtn');
    const tpNarration = document.getElementById('tpNarration');

    /* defaults */
    const today = new Date().toISOString().split('T')[0];
    if (tpDateEl) tpDateEl.value = today;
    if (tpRecvDate) tpRecvDate.value = today;

    /* ── Row template ── */
    let tpRowCount = 0;

    function createTpRow() {
        tpRowCount++;
        const tr = document.createElement('tr');
        tr.dataset.row = tpRowCount;

        // Build size options
        const sizeOpts = TP_SIZES.map(s =>
            `<option value="${s.value}">${s.label}</option>`
        ).join('');

        tr.innerHTML = `
            <td class="gc-sr">${tpRowCount}</td>
            <td class="gc-brand">
                <input type="text" class="tp-grid-input text-left" data-col="brand" autocomplete="off" placeholder="Brand or shortcode…">
                <div class="tp-prod-ac hidden" data-role="prod-ac"></div>
            </td>
            <td class="gc-code">
                <div class="tp-cell-readonly" data-col="code" style="justify-content:flex-start">—</div>
            </td>
            <td class="gc-size">
                <select class="tp-size-select" data-col="size">${sizeOpts}</select>
            </td>
            <td class="gc-cases">
                <input type="number" class="tp-grid-input text-right" data-col="cases" min="0" placeholder="0">
            </td>
            <td class="gc-bpc">
                <div class="tp-cell-readonly" data-col="bpc">—</div>
            </td>
            <td class="gc-loose">
                <input type="number" class="tp-grid-input text-right" data-col="loose" min="0" placeholder="0">
            </td>
            <td class="gc-total">
                <div class="tp-cell-readonly" data-col="totalBtl">0</div>
            </td>
            <td class="gc-rate">
                <input type="number" class="tp-grid-input text-right" data-col="rate" min="0" step="0.01" placeholder="0.00">
            </td>
            <td class="gc-amount">
                <div class="tp-cell-readonly" data-col="amount">0.00</div>
            </td>
        `;

        tpGridBody.appendChild(tr);
        wireGridRow(tr);
        wireProdAc(tr);
        return tr;
    }

    /* ── Get interactive fields inside a row ── */
    function getRowFields(tr) {
        return Array.from(tr.querySelectorAll('input.tp-grid-input, select.tp-size-select'));
    }

    /* ── Recalculate a single row ── */
    function recalcRow(tr) {
        const sizeEl  = tr.querySelector('[data-col="size"]');
        const casesEl = tr.querySelector('[data-col="cases"]');
        const looseEl = tr.querySelector('[data-col="loose"]');
        const rateEl  = tr.querySelector('[data-col="rate"]');
        const bpcEl   = tr.querySelector('[data-col="bpc"]');
        const totalEl = tr.querySelector('[data-col="totalBtl"]');
        const amtEl   = tr.querySelector('[data-col="amount"]');

        const sizeVal = sizeEl ? sizeEl.value : '';
        const sizeObj = TP_SIZES.find(s => s.value === sizeVal) || TP_SIZES[0];
        const bpc     = sizeObj.bpc;
        const cases   = parseInt(casesEl?.value) || 0;
        const loose   = parseInt(looseEl?.value) || 0;
        const rate    = parseFloat(rateEl?.value) || 0;
        const totalBtl = (cases * bpc) + loose;
        const amount   = totalBtl * rate;

        if (bpcEl)   { bpcEl.textContent = bpc > 0 ? bpc : '—'; bpcEl.classList.toggle('has-value', bpc > 0); }
        if (totalEl) { totalEl.textContent = totalBtl; totalEl.classList.toggle('has-value', totalBtl > 0); }
        if (amtEl)   { amtEl.textContent = amount > 0 ? amount.toFixed(2) : '0.00'; amtEl.classList.toggle('has-value', amount > 0); }

        recalcTotals();
    }

    /* ── Recalculate footer totals ── */
    function recalcTotals() {
        let totalCases = 0, totalLoose = 0, totalBtl = 0, totalAmt = 0;
        tpGridBody.querySelectorAll('tr').forEach(tr => {
            const sizeEl  = tr.querySelector('[data-col="size"]');
            const sizeVal = sizeEl ? sizeEl.value : '';
            const sizeObj = TP_SIZES.find(s => s.value === sizeVal) || TP_SIZES[0];
            const cases = parseInt(tr.querySelector('[data-col="cases"]')?.value) || 0;
            const loose = parseInt(tr.querySelector('[data-col="loose"]')?.value) || 0;
            const rate  = parseFloat(tr.querySelector('[data-col="rate"]')?.value) || 0;
            const btl   = (cases * sizeObj.bpc) + loose;
            totalCases += cases;
            totalLoose += loose;
            totalBtl   += btl;
            totalAmt   += btl * rate;
        });

        const elCases   = document.getElementById('tpTotalCases');
        const elLoose   = document.getElementById('tpTotalLoose');
        const elBottles = document.getElementById('tpTotalBottles');
        const elAmount  = document.getElementById('tpTotalAmount');
        if (elCases)   elCases.textContent   = totalCases;
        if (elLoose)   elLoose.textContent   = totalLoose;
        if (elBottles) elBottles.textContent = totalBtl;
        if (elAmount)  elAmount.textContent  = '₹ ' + totalAmt.toFixed(2);
    }

    /* ── Wire a grid row for keyboard + calculation ── */
    function wireGridRow(tr) {
        const fields = getRowFields(tr);

        fields.forEach((field) => {
            // Recalc on any change
            field.addEventListener('input', () => recalcRow(tr));
            field.addEventListener('change', () => recalcRow(tr));

            // Mark row active on focus
            field.addEventListener('focus', () => {
                tpGridBody.querySelectorAll('tr').forEach(r => r.classList.remove('tp-row-active'));
                tr.classList.add('tp-row-active');
            });

            // Keyboard navigation
            field.addEventListener('keydown', (e) => {
                const col     = field.dataset.col || field.getAttribute('data-col');

                // If the product AC dropdown is open on this brand field, let it handle these keys
                if (col === 'brand' && tpProdAcActive && tpProdAcActive.tr === tr) {
                    const acDrop = tpProdAcActive.dropdown;
                    if (acDrop && !acDrop.classList.contains('hidden')) {
                        if (['ArrowDown', 'ArrowUp', 'Escape'].includes(e.key)) return;
                        if (e.key === 'Enter') return; // AC handler fires first via stopPropagation
                    }
                }

                const allRows = Array.from(tpGridBody.querySelectorAll('tr'));
                const rowIdx  = allRows.indexOf(tr);
                const fIdx    = fields.indexOf(field);

                if (e.key === 'Enter') {
                    e.preventDefault();
                    // Move to next field in row, or first field of next row
                    if (fIdx < fields.length - 1) {
                        fields[fIdx + 1].focus();
                    } else {
                        // End of row → next row (create if needed)
                        let nextRow = allRows[rowIdx + 1];
                        if (!nextRow) nextRow = createTpRow();
                        const nextFields = getRowFields(nextRow);
                        if (nextFields[0]) nextFields[0].focus();
                    }
                } else if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    let nextRow = allRows[rowIdx + 1];
                    if (!nextRow) nextRow = createTpRow();
                    const target = nextRow.querySelector(`[data-col="${col}"]`);
                    if (target && (target.tagName === 'INPUT' || target.tagName === 'SELECT')) target.focus();
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    if (rowIdx > 0) {
                        const prevRow = allRows[rowIdx - 1];
                        const target = prevRow.querySelector(`[data-col="${col}"]`);
                        if (target && (target.tagName === 'INPUT' || target.tagName === 'SELECT')) target.focus();
                    }
                } else if (e.key === 'Escape') {
                    e.preventDefault();
                    field.blur();
                } else if (e.key === 'Delete' && e.ctrlKey) {
                    e.preventDefault();
                    if (allRows.length > 1) {
                        const newFocusRow = allRows[rowIdx - 1] || allRows[rowIdx + 1];
                        tr.remove();
                        renumberRows();
                        recalcTotals();
                        if (newFocusRow) {
                            const f = getRowFields(newFocusRow);
                            if (f[0]) f[0].focus();
                        }
                    }
                }
            });
        });
    }

    /* ── Renumber serial numbers ── */
    function renumberRows() {
        tpGridBody.querySelectorAll('tr').forEach((tr, i) => {
            const srCell = tr.querySelector('.gc-sr');
            if (srCell) srCell.textContent = i + 1;
            tr.dataset.row = i + 1;
        });
        tpRowCount = tpGridBody.querySelectorAll('tr').length;
    }

    /* ── Header field → grid flow (Enter key) ── */
    const tpFlowFields = Array.from(document.querySelectorAll('.tp-flow'));
    tpFlowFields.forEach((field, idx) => {
        field.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                // If AC dropdown is open with a selection, let the AC handler deal with it
                if (field.id === 'tpSupplier' && tpSupplierAC && !tpSupplierAC.classList.contains('hidden') && acFocusIdx >= 0) {
                    return; // AC handler will fire next
                }
                e.preventDefault();
                if (idx < tpFlowFields.length - 1) {
                    tpFlowFields[idx + 1].focus();
                } else {
                    // After last header field → go to first grid row
                    const firstRow = tpGridBody.querySelector('tr');
                    if (firstRow) {
                        const f = getRowFields(firstRow);
                        if (f[0]) f[0].focus();
                    }
                }
            } else if (e.key === 'Escape') {
                e.preventDefault();
                field.blur();
            }
        });
    });

    /* ── Narration → Ctrl+S flow ── */
    if (tpNarration) {
        tpNarration.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                // Focus save button
                if (tpSaveBtn) tpSaveBtn.focus();
            }
        });
    }

    /* ── Clear form ── */
    function clearTpForm() {
        tpFlowFields.forEach(f => {
            if (f.type === 'date') {
                f.value = today;
            } else {
                f.value = '';
            }
        });
        if (tpNarration) tpNarration.value = '';
        tpGridBody.innerHTML = '';
        tpRowCount = 0;
        // Re-seed with empty rows
        for (let i = 0; i < 8; i++) createTpRow();
        recalcTotals();
        // Focus first field
        const first = document.getElementById('tpNumber');
        if (first) first.focus();
    }

    if (tpClearBtn) {
        tpClearBtn.addEventListener('click', clearTpForm);
    }

    /* ── Shortcut → Product resolver (shared) ── */
    function resolveShortcutToProduct(code) {
        if (!code || !allShortcuts.length) return null;
        const upper = code.trim().toUpperCase();
        const sc = allShortcuts.find(s => s.code.toUpperCase() === upper);
        if (!sc) return null;
        // Return the linked product from allProducts (for full data)
        const prod = allProducts.find(p => p.id === sc.productId);
        return prod || {
            id: sc.productId, brandName: sc.brandName, code: sc.prodCode,
            size: sc.size, category: sc.category, mrp: sc.mrp, costPrice: 0
        };
    }

    function getShortcutMatches(query) {
        if (!query || !allShortcuts.length) return [];
        const upper = query.trim().toUpperCase();
        return allShortcuts.filter(s =>
            s.code.toUpperCase().startsWith(upper) ||
            s.code.toUpperCase() === upper
        );
    }

    /* ── Product Autocomplete in TP Grid ── */
    let tpProdAcActive = null; // { tr, dropdown, focusIdx }

    function wireProdAc(tr) {
        const brandInput = tr.querySelector('[data-col="brand"]');
        const dropdown   = tr.querySelector('[data-role="prod-ac"]');
        if (!brandInput || !dropdown) return;

        let acIdx = -1;

        function showProdAc(query) {
            const q = (query || '').toLowerCase();

            // --- Shortcut matches first ---
            const scMatches = getShortcutMatches(query);
            const scProducts = scMatches.map(sc => {
                const prod = allProducts.find(p => p.id === sc.productId);
                return prod ? { ...prod, _shortcutCode: sc.code } : {
                    id: sc.productId, brandName: sc.brandName, code: sc.prodCode,
                    size: sc.size, category: sc.category, mrp: sc.mrp, costPrice: 0,
                    _shortcutCode: sc.code
                };
            }).filter(p => !(isBeerShopee && (p.category || '').toUpperCase() === 'SPIRITS'));
            const scProdIds = new Set(scProducts.map(p => p.id));

            // --- Regular text matches (exclude already-matched shortcut products) ---
            const textMatches = q.length > 0
                ? allProducts.filter(p =>
                    !scProdIds.has(p.id) &&
                    !(isBeerShopee && (p.category || '').toUpperCase() === 'SPIRITS') && (
                    (p.brandName || '').toLowerCase().includes(q) ||
                    (p.code || '').toLowerCase().includes(q) ||
                    (p.category || '').toLowerCase().includes(q))
                  )
                : allProducts.filter(p =>
                    !scProdIds.has(p.id) &&
                    !(isBeerShopee && (p.category || '').toUpperCase() === 'SPIRITS')
                  ).slice(0, 30);

            const combined = [...scProducts, ...textMatches];

            if (combined.length === 0 && q.length > 0) {
                dropdown.innerHTML = '<div class="tp-prod-ac-empty">No matching products</div>';
            } else if (combined.length === 0) {
                dropdown.innerHTML = '<div class="tp-prod-ac-empty">No products in master</div>';
            } else {
                dropdown.innerHTML = combined.map((p, i) => `
                    <div class="tp-prod-ac-item${p._shortcutCode ? ' sc-match' : ''}" data-idx="${i}">
                        ${p._shortcutCode ? '<span class="tp-prod-ac-sc-badge">' + esc(p._shortcutCode) + '</span>' : ''}
                        <span class="tp-prod-ac-brand">${esc(p.brandName)}</span>
                        <span class="tp-prod-ac-meta">${esc(p.code || '')} · ${p.size ? p.size + 'ml' : ''} · ${esc(p.category || '')}</span>
                    </div>
                `).join('');

                dropdown.querySelectorAll('.tp-prod-ac-item').forEach(item => {
                    item.addEventListener('mousedown', (e) => {
                        e.preventDefault();
                        const idx = parseInt(item.dataset.idx);
                        selectProdAcItem(tr, combined[idx], brandInput, dropdown);
                    });
                });
            }

            dropdown.classList.remove('hidden');
            acIdx = -1;
            tpProdAcActive = { tr, dropdown, focusIdx: -1, combined };
        }

        function hideProdAc() {
            dropdown.classList.add('hidden');
            dropdown.innerHTML = '';
            acIdx = -1;
            if (tpProdAcActive && tpProdAcActive.tr === tr) tpProdAcActive = null;
        }

        function focusAcItem(newIdx) {
            const items = dropdown.querySelectorAll('.tp-prod-ac-item');
            if (items.length === 0) return;
            acIdx = Math.max(0, Math.min(newIdx, items.length - 1));
            items.forEach((it, i) => it.classList.toggle('focused', i === acIdx));
            items[acIdx]?.scrollIntoView({ block: 'nearest' });
        }

        brandInput.addEventListener('input', () => {
            showProdAc(brandInput.value);
        });

        brandInput.addEventListener('focus', () => {
            // Close any other open product AC
            if (tpProdAcActive && tpProdAcActive.tr !== tr) {
                tpProdAcActive.dropdown.classList.add('hidden');
                tpProdAcActive = null;
            }
            if (allProducts.length > 0) {
                showProdAc(brandInput.value);
            }
        });

        brandInput.addEventListener('blur', () => {
            // Delay to allow click on items
            setTimeout(() => hideProdAc(), 150);
        });

        brandInput.addEventListener('keydown', (e) => {
            /* ── Shortcut quick-select: Enter/Tab on exact shortcode match ── */
            if (e.key === 'Enter' || e.key === 'Tab') {
                const val = brandInput.value.trim();
                if (val && acIdx < 0) {
                    const scProd = resolveShortcutToProduct(val);
                    if (scProd) {
                        e.preventDefault(); e.stopPropagation();
                        selectProdAcItem(tr, scProd, brandInput, dropdown);
                        return;
                    }
                }
            }

            if (dropdown.classList.contains('hidden')) return;
            const items = dropdown.querySelectorAll('.tp-prod-ac-item');

            if (e.key === 'ArrowDown') {
                e.preventDefault();
                e.stopPropagation();
                focusAcItem(acIdx + 1);
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                e.stopPropagation();
                focusAcItem(acIdx - 1);
            } else if (e.key === 'Enter' && acIdx >= 0 && items[acIdx]) {
                e.preventDefault();
                e.stopPropagation();
                // Use the combined list from tpProdAcActive
                const combined = tpProdAcActive?.combined;
                if (combined && combined[acIdx]) {
                    selectProdAcItem(tr, combined[acIdx], brandInput, dropdown);
                }
            } else if (e.key === 'Escape') {
                e.preventDefault();
                e.stopPropagation();
                hideProdAc();
            }
        });
    }

    function selectProdAcItem(tr, product, brandInput, dropdown) {
        // Fill brand name
        brandInput.value = product.brandName || '';

        // Fill code
        const codeEl = tr.querySelector('[data-col="code"]');
        if (codeEl) {
            codeEl.textContent = product.code || '—';
            codeEl.classList.toggle('has-value', !!product.code);
        }

        // Fill size
        const sizeEl = tr.querySelector('[data-col="size"]');
        if (sizeEl && product.size) {
            sizeEl.value = product.size;
        }

        // Fill rate from MRP or costPrice
        const rateEl = tr.querySelector('[data-col="rate"]');
        if (rateEl) {
            const rate = product.costPrice || product.mrp || 0;
            if (rate > 0) rateEl.value = rate;
        }

        // Recalculate row (updates BPC, totals)
        recalcRow(tr);

        // Hide dropdown
        dropdown.classList.add('hidden');
        dropdown.innerHTML = '';
        if (tpProdAcActive && tpProdAcActive.tr === tr) tpProdAcActive = null;

        // Move focus to the cases field
        const casesEl = tr.querySelector('[data-col="cases"]');
        if (casesEl) casesEl.focus();
    }

    /* ── Save handler (placeholder — data saving) ── */
    if (tpSaveBtn) {
        tpSaveBtn.addEventListener('click', () => saveTpEntry());
    }

    async function saveTpEntry() {
        const tpNumber   = document.getElementById('tpNumber')?.value.trim();
        const tpSupplier = document.getElementById('tpSupplier')?.value.trim();

        if (!tpNumber) {
            document.getElementById('tpNumber')?.focus();
            showTpToast('TP Number is required', true);
            return;
        }
        if (!tpSupplier) {
            document.getElementById('tpSupplier')?.focus();
            showTpToast('Supplier name is required', true);
            return;
        }

        // Gather items
        const items = [];
        tpGridBody.querySelectorAll('tr').forEach(tr => {
            const brand = tr.querySelector('[data-col="brand"]')?.value.trim();
            if (!brand) return;
            const code = tr.querySelector('[data-col="code"]')?.textContent.trim() || '';
            const sizeEl = tr.querySelector('[data-col="size"]');
            const sizeObj = TP_SIZES.find(s => s.value === (sizeEl?.value || '')) || TP_SIZES[0];
            const cases = parseInt(tr.querySelector('[data-col="cases"]')?.value) || 0;
            const loose = parseInt(tr.querySelector('[data-col="loose"]')?.value) || 0;
            const rate  = parseFloat(tr.querySelector('[data-col="rate"]')?.value) || 0;
            const totalBtl = (cases * sizeObj.bpc) + loose;
            items.push({
                brand,
                code: code !== '—' ? code : '',
                size: sizeEl?.value || '',
                sizeLabel: sizeObj.label,
                cases,
                bpc: sizeObj.bpc,
                loose,
                totalBtl,
                rate,
                amount: totalBtl * rate,
            });
        });

        if (items.length === 0) {
            const firstBrand = tpGridBody.querySelector('[data-col="brand"]');
            if (firstBrand) firstBrand.focus();
            showTpToast('Add at least one item', true);
            return;
        }

        const tpData = {
            id: 'TP_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
            tpNumber,
            tpDate: tpDateEl?.value || today,
            receivedDate: tpRecvDate?.value || today,
            supplier: tpSupplier,
            vehicle: document.getElementById('tpVehicle')?.value.trim() || '',
            narration: tpNarration?.value.trim() || '',
            items,
            totalCases:   items.reduce((s, i) => s + i.cases, 0),
            totalLoose:   items.reduce((s, i) => s + i.loose, 0),
            totalBottles: items.reduce((s, i) => s + i.totalBtl, 0),
            totalAmount:  items.reduce((s, i) => s + i.amount, 0),
            createdAt: new Date().toISOString(),
        };

        // Persist to disk
        try {
            const result = await window.electronAPI.saveTp({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || '',
                tp: tpData,
            });
            if (result.success) {
                showTpToast(`TP ${tpNumber} saved — ${items.length} items, ₹${tpData.totalAmount.toFixed(2)}`);
                setTimeout(clearTpForm, 1200);
            } else {
                showTpToast('Save failed: ' + result.error, true);
            }
        } catch (err) {
            showTpToast('Save error: ' + err.message, true);
        }
    }

    function showTpToast(msg, isError = false) {
        let toast = document.querySelector('.tp-toast');
        if (!toast) {
            toast = document.createElement('div');
            toast.className = 'tp-toast hidden';
            document.body.appendChild(toast);
        }
        toast.textContent = msg;
        toast.classList.toggle('error', isError);
        if (isError) {
            toast.style.background = '#ef4444';
            toast.style.color = '#fff';
        } else {
            toast.style.background = '';
            toast.style.color = '';
        }
        toast.classList.remove('hidden');
        clearTimeout(toast._timer);
        toast._timer = setTimeout(() => toast.classList.add('hidden'), 2500);
    }

    /* ── Ctrl+S global shortcut for TP save ── */
    document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
            // Only save if TP form is visible
            const tpPanel = document.getElementById('sub-add-tp');
            if (tpPanel && tpPanel.classList.contains('active')) {
                e.preventDefault();
                saveTpEntry();
            }
        }
    });

    /* ── Seed initial empty rows ── */
    if (tpGridBody) {
        for (let i = 0; i < 8; i++) createTpRow();
    }

    /* ── Import TP from Maharashtra Excise XLS ── */
    const tpImportXlsBtn = document.getElementById('tpImportXlsBtn');

    /* ── New Products Modal (shown before TP import save) ── */
    function showNewProductsModal(newProducts) {
        return new Promise((resolve) => {
            const overlay = document.createElement('div');
            overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.55);z-index:9999;display:flex;align-items:center;justify-content:center;';

            const rows = newProducts.map(p =>
                `<tr>
                  <td style="padding:6px 10px;border-bottom:1px solid var(--border-color,#e5e7eb);font-size:12px;color:var(--text-muted,#6b7280);font-family:monospace">${p.code}</td>
                  <td style="padding:6px 10px;border-bottom:1px solid var(--border-color,#e5e7eb);font-size:12px">${p.brand}</td>
                  <td style="padding:6px 10px;border-bottom:1px solid var(--border-color,#e5e7eb);font-size:12px;color:var(--text-muted,#6b7280)">${p.size || '—'}</td>
                </tr>`
            ).join('');

            overlay.innerHTML = `
              <div style="background:var(--surface,#fff);border-radius:12px;box-shadow:0 8px 40px rgba(0,0,0,0.22);width:560px;max-width:94vw;max-height:80vh;display:flex;flex-direction:column;overflow:hidden;">
                <div style="display:flex;align-items:center;gap:12px;padding:18px 20px 14px;border-bottom:1px solid var(--border-color,#e5e7eb);">
                  <div style="width:36px;height:36px;border-radius:8px;background:#fef3c7;display:flex;align-items:center;justify-content:center;flex-shrink:0">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
                  </div>
                  <div>
                    <h3 style="margin:0;font-size:15px;font-weight:600;">New Products Found</h3>
                    <p style="margin:2px 0 0;font-size:12px;color:var(--text-muted,#6b7280);">${newProducts.length} product${newProducts.length > 1 ? 's' : ''} not in your Product Master</p>
                  </div>
                </div>
                <div style="overflow-y:auto;flex:1;padding:14px 20px;">
                  <p style="margin:0 0 10px;font-size:13px;color:var(--text-muted,#6b7280);">The following items from the Excel file are not in your product master. They will be imported with size from the Excel sheet. You can update them in Product Master later.</p>
                  <table style="width:100%;border-collapse:collapse;">
                    <thead>
                      <tr style="background:var(--surface-alt,#f9fafb);">
                        <th style="padding:6px 10px;text-align:left;font-size:11px;font-weight:600;color:var(--text-muted,#6b7280);border-bottom:2px solid var(--border-color,#e5e7eb);">CODE</th>
                        <th style="padding:6px 10px;text-align:left;font-size:11px;font-weight:600;color:var(--text-muted,#6b7280);border-bottom:2px solid var(--border-color,#e5e7eb);">BRAND NAME</th>
                        <th style="padding:6px 10px;text-align:left;font-size:11px;font-weight:600;color:var(--text-muted,#6b7280);border-bottom:2px solid var(--border-color,#e5e7eb);">SIZE</th>
                      </tr>
                    </thead>
                    <tbody>${rows}</tbody>
                  </table>
                </div>
                <div style="display:flex;gap:10px;justify-content:flex-end;padding:14px 20px;border-top:1px solid var(--border-color,#e5e7eb);">
                  <button id="npCancelBtn" style="padding:8px 18px;border-radius:6px;border:1px solid var(--border-color,#e5e7eb);background:var(--surface,#fff);font-size:13px;cursor:pointer;">Cancel Import</button>
                  <button id="npProceedBtn" style="padding:8px 18px;border-radius:6px;border:none;background:var(--accent,#2563eb);color:#fff;font-size:13px;font-weight:600;cursor:pointer;">Proceed &amp; Save</button>
                </div>
              </div>`;

            document.body.appendChild(overlay);

            overlay.querySelector('#npProceedBtn').addEventListener('click', () => {
                document.body.removeChild(overlay);
                resolve(true);
            });
            overlay.querySelector('#npCancelBtn').addEventListener('click', () => {
                document.body.removeChild(overlay);
                resolve(false);
            });
            // Click outside to cancel
            overlay.addEventListener('click', (e) => {
                if (e.target === overlay) {
                    document.body.removeChild(overlay);
                    resolve(false);
                }
            });
        });
    }

    async function importTpFromXls() {
        showTpToast('Opening file picker…');

        try {
            const result = await window.electronAPI.importTpXls({
                barName: activeBar.barName || '',
                financialYear: activeBar.financialYear || '',
            });
            if (!result.success) {
                if (result.error !== 'Cancelled') showTpToast('Import failed: ' + result.error, true);
                return;
            }

            const tps = result.tps || [];
            if (tps.length === 0) {
                showTpToast('No TPs found in the file', true);
                return;
            }

            // If new products found, show warning modal and ask user to proceed
            const newProducts = result.newProducts || [];
            if (newProducts.length > 0) {
                const proceed = await showNewProductsModal(newProducts);
                if (!proceed) return;
            }

            // Auto-save every TP as a separate entry
            let saved = 0;
            for (let i = 0; i < tps.length; i++) {
                const tp = tps[i];
                const tpData = {
                    id: 'TP_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6) + '_' + i,
                    tpNumber: tp.tpNumber,
                    tpDate: tp.tpDate,
                    receivedDate: tp.receivedDate,
                    supplier: tp.supplier,
                    vehicle: tp.vehicle,
                    narration: 'Imported from ' + result.fileName,
                    items: tp.items,
                    totalCases: tp.totalCases,
                    totalLoose: tp.totalLoose,
                    totalBottles: tp.totalBottles,
                    totalAmount: tp.totalAmount,
                    createdAt: new Date().toISOString(),
                };
                try {
                    const saveResult = await window.electronAPI.saveTp({
                        barName: activeBar.barName,
                        financialYear: activeBar.financialYear || '',
                        tp: tpData,
                    });
                    if (saveResult.success) saved++;
                } catch (err) {
                    console.error('Save TP error:', err);
                }
            }
            showTpToast(`Imported & saved ${saved}/${tps.length} TP${tps.length > 1 ? 's' : ''} from "${result.fileName}"${result.mrpUpdated > 0 ? ` • ${result.mrpUpdated} MRP updated` : ''}`);
            await loadTpSummary();        } catch (err) {
            showTpToast('Import error: ' + err.message, true);
        }
    }

    /* ── Populate the Add TP form from imported TP data ── */
    function populateTpFormFromImport(tp) {
        // Clear existing form
        clearTpForm();

        // Fill header fields
        const tpNumEl   = document.getElementById('tpNumber');
        const tpDateEl2 = document.getElementById('tpDate');
        const tpRecvEl  = document.getElementById('tpReceivedDate');
        const tpSupEl   = document.getElementById('tpSupplier');
        const tpVehEl   = document.getElementById('tpVehicle');

        if (tpNumEl)  tpNumEl.value  = tp.tpNumber || '';
        if (tpDateEl2) tpDateEl2.value = tp.tpDate || '';
        if (tpRecvEl) tpRecvEl.value = tp.receivedDate || '';
        if (tpSupEl)  tpSupEl.value  = tp.supplier || '';
        if (tpVehEl)  tpVehEl.value  = tp.vehicle || '';

        // Clear existing grid rows
        if (tpGridBody) {
            tpGridBody.querySelectorAll('tr').forEach(tr => tr.remove());
            tpRowCount = 0;
        }

        // Create rows for each item
        for (const item of tp.items) {
            const tr = createTpRow();
            if (!tr) continue;

            // Fill brand
            const brandEl = tr.querySelector('[data-col="brand"]');
            if (brandEl) brandEl.value = item.brand || '';

            // Fill code
            const codeEl = tr.querySelector('[data-col="code"]');
            if (codeEl) {
                codeEl.textContent = item.code || '—';
                codeEl.classList.toggle('has-value', !!item.code);
            }

            // Fill size — match by ml + bpc from ALL_SIZES
            const sizeEl = tr.querySelector('[data-col="size"]');
            if (sizeEl && item.size) {
                const mlNum = parseInt(item.size) || 0;
                const bpcNum = item.bpc || 0;
                // Try exact ml+bpc match first
                let matched = ALL_SIZES.find(s => s.ml === mlNum && s.bpc === bpcNum);
                // Fallback: first entry with matching ml
                if (!matched) matched = ALL_SIZES.find(s => s.ml === mlNum);
                if (matched) sizeEl.value = matched.value;
            }

            // Fill cases
            const casesEl = tr.querySelector('[data-col="cases"]');
            if (casesEl && item.cases) casesEl.value = item.cases;

            // Fill loose
            const looseEl = tr.querySelector('[data-col="loose"]');
            if (looseEl && item.loose) looseEl.value = item.loose;

            // Fill rate
            const rateEl = tr.querySelector('[data-col="rate"]');
            if (rateEl && item.rate) rateEl.value = item.rate;

            // Recalculate row
            recalcRow(tr);
        }

        // Add a couple empty rows at the end
        createTpRow();
        createTpRow();

        // Recalculate totals
        recalcTotals();

        // Focus TP number field
        if (tpNumEl) tpNumEl.focus();
    }

    if (tpImportXlsBtn) {
        tpImportXlsBtn.addEventListener('click', importTpFromXls);
    }

    // ═══════════════════════════════════════════════════════
    //  TP SUPPLIER AUTOCOMPLETE
    // ═══════════════════════════════════════════════════════

    const tpSupplierInput = document.getElementById('tpSupplier');
    const tpSupplierAC    = document.getElementById('tpSupplierAC');
    let acSuppliers       = [];   // populated once suppliers load
    let acFiltered        = [];
    let acFocusIdx        = -1;

    /* Called after supplier list loads / changes */
    function refreshAcSuppliers(suppliers) {
        acSuppliers = suppliers || [];
    }

    function showAcDropdown() {
        if (!tpSupplierAC) return;
        const q = (tpSupplierInput?.value || '').toLowerCase();
        acFiltered = q
            ? acSuppliers.filter(s =>
                (s.name || '').toLowerCase().includes(q) ||
                (s.city || '').toLowerCase().includes(q))
            : [...acSuppliers];
        acFocusIdx = -1;
        renderAcList();
        tpSupplierAC.classList.toggle('hidden', acFiltered.length === 0 && !q);
    }

    function hideAcDropdown() {
        if (tpSupplierAC) tpSupplierAC.classList.add('hidden');
        acFocusIdx = -1;
    }

    function renderAcList() {
        if (!tpSupplierAC) return;
        tpSupplierAC.innerHTML = '';

        if (acFiltered.length === 0) {
            tpSupplierAC.innerHTML = '<div class="tp-ac-empty">No matching suppliers</div>';
            tpSupplierAC.classList.remove('hidden');
            return;
        }

        acFiltered.forEach((sup, idx) => {
            const item = document.createElement('div');
            item.className = 'tp-ac-item';
            item.role = 'option';
            item.dataset.idx = idx;
            item.innerHTML = `<span class="tp-ac-name">${esc(sup.name)}</span><span class="tp-ac-city">${esc(sup.city || '')}</span>`;
            item.addEventListener('mousedown', (e) => {
                e.preventDefault(); // prevent blur before selection
                selectAcItem(sup);
            });
            item.addEventListener('mouseenter', () => setAcFocus(idx));
            tpSupplierAC.appendChild(item);
        });
        tpSupplierAC.classList.remove('hidden');
    }

    function setAcFocus(idx) {
        acFocusIdx = idx;
        tpSupplierAC?.querySelectorAll('.tp-ac-item').forEach((el, i) => {
            el.classList.toggle('focused', i === idx);
            if (i === idx) el.scrollIntoView({ block: 'nearest' });
        });
    }

    function selectAcItem(sup) {
        if (tpSupplierInput) tpSupplierInput.value = sup.name;
        hideAcDropdown();
        // Move to next flow field
        const nextFlow = document.querySelector('.tp-flow[data-tp-idx="4"]');
        if (nextFlow) nextFlow.focus();
    }

    if (tpSupplierInput) {
        tpSupplierInput.addEventListener('input', () => showAcDropdown());
        tpSupplierInput.addEventListener('focus', () => showAcDropdown());
        tpSupplierInput.addEventListener('blur', () => {
            // Delay to allow mousedown on items to fire first
            setTimeout(hideAcDropdown, 150);
        });

        tpSupplierInput.addEventListener('keydown', (e) => {
            if (!tpSupplierAC || tpSupplierAC.classList.contains('hidden')) return;

            if (e.key === 'ArrowDown') {
                e.preventDefault();
                e.stopPropagation();
                setAcFocus(Math.min(acFocusIdx + 1, acFiltered.length - 1));
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                e.stopPropagation();
                setAcFocus(Math.max(acFocusIdx - 1, 0));
            } else if (e.key === 'Enter' && acFocusIdx >= 0 && acFiltered[acFocusIdx]) {
                e.preventDefault();
                e.stopPropagation();
                selectAcItem(acFiltered[acFocusIdx]);
            } else if (e.key === 'Escape') {
                e.preventDefault();
                e.stopPropagation();
                hideAcDropdown();
            }
        });
    }

    // ═══════════════════════════════════════════════════════
    //  TP SUMMARY — List & Detail View
    // ═══════════════════════════════════════════════════════

    const tpsTableBody     = document.getElementById('tpsTableBody');
    const tpsEmptyState    = document.getElementById('tpsEmptyState');
    const tpsDateFrom      = document.getElementById('tpsDateFrom');
    const tpsDateTo        = document.getElementById('tpsDateTo');
    const tpsSupFilter     = document.getElementById('tpsSupplierFilter');
    const tpsSearchEl      = document.getElementById('tpsSearch');
    const tpsExportBtn     = document.getElementById('tpsExportBtn');
    const tpsAddNewBtn     = document.getElementById('tpsAddNewBtn');
    const tpsTotalBadge    = document.getElementById('tpsTotalBadge');
    const tpsDetailOverlay = document.getElementById('tpsDetailOverlay');
    const tpsDetCloseBtn   = document.getElementById('tpsDetCloseBtn');
    const tpsDetDeleteBtn  = document.getElementById('tpsDetDeleteBtn');

    let allTps       = [];
    let filteredTps  = [];
    let tpsViewingId = null; // currently open detail

    /* ── Load all TPs ── */
    async function loadTpSummary() {
        if (!activeBar.barName) return;
        try {
            const result = await window.electronAPI.getTps({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || ''
            });
            if (result.success) {
                allTps = (result.tps || []).sort((a, b) => {
                    // Sort by TP date descending, then createdAt descending
                    const da = a.tpDate || a.createdAt || '';
                    const db = b.tpDate || b.createdAt || '';
                    return db.localeCompare(da);
                });
                // Assign persistent Sr.No based on ascending (chronological) order
                const ascSorted = [...allTps].sort((a, b) => {
                    const da = a.tpDate || a.createdAt || '';
                    const db = b.tpDate || b.createdAt || '';
                    return da.localeCompare(db);
                });
                ascSorted.forEach((tp, i) => { tp._srNo = i + 1; });
            }
        } catch (err) {
            console.error('loadTpSummary error:', err);
        }
        populateTpsSupplierFilter();
        applyTpsFilters();
    }

    /* ── Populate supplier dropdown from TPs ── */
    function populateTpsSupplierFilter() {
        if (!tpsSupFilter) return;
        const suppliers = [...new Set(allTps.map(t => t.supplier).filter(Boolean))].sort();
        const current = tpsSupFilter.value;
        tpsSupFilter.innerHTML = '<option value="">All Suppliers</option>' +
            suppliers.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
        tpsSupFilter.value = current;
    }

    /* ── Apply filters ── */
    function applyTpsFilters() {
        const fromDate = tpsDateFrom?.value || '';
        const toDate   = tpsDateTo?.value || '';
        const supplier = tpsSupFilter?.value || '';
        const query    = (tpsSearchEl?.value || '').toLowerCase();

        filteredTps = allTps.filter(tp => {
            // Date range
            if (fromDate && (tp.tpDate || '') < fromDate) return false;
            if (toDate && (tp.tpDate || '') > toDate) return false;
            // Supplier
            if (supplier && tp.supplier !== supplier) return false;
            // Text search
            if (query) {
                const haystack = [
                    tp.tpNumber, tp.supplier, tp.vehicle,
                    ...(tp.items || []).map(i => i.brand + ' ' + i.code)
                ].join(' ').toLowerCase();
                if (!haystack.includes(query)) return false;
            }
            return true;
        });

        renderTpsSummary();
    }

    /* ── Render table + summary cards ── */
    function renderTpsSummary() {
        if (!tpsTableBody) return;
        tpsTableBody.innerHTML = '';

        // Summary totals
        const totals = { tps: filteredTps.length, cases: 0, bottles: 0, amount: 0, items: 0 };
        filteredTps.forEach(tp => {
            totals.cases   += tp.totalCases || 0;
            totals.bottles += tp.totalBottles || 0;
            totals.amount  += tp.totalAmount || 0;
            totals.items   += (tp.items || []).length;
        });

        // Cards
        const el = id => document.getElementById(id);
        if (el('tsCardTps'))     el('tsCardTps').textContent     = totals.tps;
        if (el('tsCardCases'))   el('tsCardCases').textContent   = totals.cases.toLocaleString('en-IN');
        if (el('tsCardBottles')) el('tsCardBottles').textContent = totals.bottles.toLocaleString('en-IN');
        if (el('tsCardAmount'))  el('tsCardAmount').textContent  = '₹' + totals.amount.toLocaleString('en-IN', { minimumFractionDigits: 2 });
        if (tpsTotalBadge) tpsTotalBadge.textContent = `${totals.tps} TP${totals.tps !== 1 ? 's' : ''} · ₹${totals.amount.toLocaleString('en-IN', { minimumFractionDigits: 0 })}`;

        // Footer totals
        if (el('tpsFtItems')) el('tpsFtItems').textContent = totals.items;
        if (el('tpsFtCases')) el('tpsFtCases').textContent = totals.cases.toLocaleString('en-IN');
        if (el('tpsFtBtls'))  el('tpsFtBtls').textContent  = totals.bottles.toLocaleString('en-IN');
        if (el('tpsFtAmt'))   el('tpsFtAmt').textContent   = '₹' + totals.amount.toLocaleString('en-IN', { minimumFractionDigits: 2 });

        // Empty state
        if (tpsEmptyState) tpsEmptyState.style.display = filteredTps.length === 0 ? '' : 'none';

        // Rows
        filteredTps.forEach((tp, idx) => {
            const tr = document.createElement('tr');
            tr.dataset.tpId = tp.id;
            if (tp.id === tpsViewingId) tr.classList.add('tps-row-selected');

            const tpDate = formatDateShort(tp.tpDate);
            const recvDate = formatDateShort(tp.receivedDate);
            const itemCount = (tp.items || []).length;

            tr.innerHTML = `
                <td class="tps-td-sr">${tp._srNo ?? (idx + 1)}</td>
                <td class="tps-td-num">${esc(tp.tpNumber || '—')}</td>
                <td class="tps-td-date">${tpDate}</td>
                <td class="tps-td-recv">${recvDate}</td>
                <td class="tps-td-sup">${esc(tp.supplier || '—')}</td>
                <td class="tps-td-right">${itemCount}</td>
                <td class="tps-td-right">${(tp.totalCases || 0).toLocaleString('en-IN')}</td>
                <td class="tps-td-right">${(tp.totalBottles || 0).toLocaleString('en-IN')}</td>
                <td class="tps-td-amt">₹${(tp.totalAmount || 0).toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                <td style="text-align:center">
                    <button class="tps-action-btn tps-view-btn" title="View Details">
                        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
                    </button>
                    <button class="tps-action-btn tps-action-del tps-del-btn" title="Delete">
                        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                    </button>
                </td>
            `;

            // Click row to view
            tr.addEventListener('click', (e) => {
                if (e.target.closest('.tps-del-btn')) return;
                openTpsDetail(tp);
            });

            // View button
            tr.querySelector('.tps-view-btn')?.addEventListener('click', (e) => {
                e.stopPropagation();
                openTpsDetail(tp);
            });

            // Delete button
            tr.querySelector('.tps-del-btn')?.addEventListener('click', (e) => {
                e.stopPropagation();
                deleteTpFromSummary(tp);
            });

            tpsTableBody.appendChild(tr);
        });
    }

    /* ── Format date helper ── */
    function formatDateShort(dateStr) {
        if (!dateStr) return '—';
        try {
            const d = new Date(dateStr);
            return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
        } catch {
            return dateStr;
        }
    }

    function getBillingBarAddress() {
        return [activeBar.address, activeBar.area, activeBar.city, activeBar.state, activeBar.pinCode]
            .filter(Boolean)
            .join(', ') || 'Address: —';
    }

    function getBillingBarLicense() {
        const licNo = activeBar.licenseNo || activeBar.licNo || '';
        return licNo ? `Lic No: ${licNo}` : 'Lic No: —';
    }

    /* ── Build size <option> list for dialog grid ── */
    function tpsDetSizeOptions(selectedVal) {
        return TP_SIZES.map(s =>
            `<option value="${esc(s.value)}"${s.value === selectedVal ? ' selected' : ''}>${esc(s.label)}</option>`
        ).join('');
    }

    /* ── Recalculate a single dialog-grid row ── */
    function recalcTpsDetRow(tr) {
        const sizeVal = tr.querySelector('.tps-det-size')?.value || '';
        const sizeObj = TP_SIZES.find(s => s.value === sizeVal) || TP_SIZES[0];
        const cases = parseInt(tr.querySelector('.tps-det-cases')?.value) || 0;
        const loose = parseInt(tr.querySelector('.tps-det-loose')?.value) || 0;
        const rate  = parseFloat(tr.querySelector('.tps-det-rate')?.value) || 0;
        const totalBtl = (cases * (sizeObj.bpc || 0)) + loose;
        const amount = totalBtl * rate;
        const btlEl = tr.querySelector('.tps-det-totalbtl');
        const amtEl = tr.querySelector('.tps-det-amount');
        if (btlEl) btlEl.textContent = totalBtl;
        if (amtEl) amtEl.textContent = '₹' + amount.toLocaleString('en-IN', { minimumFractionDigits: 2 });
        recalcTpsDetTotals();
    }

    /* ── Recalculate dialog footer totals ── */
    function recalcTpsDetTotals() {
        const gridBody = document.getElementById('tpsDetGridBody');
        if (!gridBody) return;
        let totalCases = 0, totalLoose = 0, totalBtls = 0, totalAmt = 0;
        gridBody.querySelectorAll('tr').forEach(tr => {
            const sizeVal = tr.querySelector('.tps-det-size')?.value || '';
            const sizeObj = TP_SIZES.find(s => s.value === sizeVal) || TP_SIZES[0];
            const cases = parseInt(tr.querySelector('.tps-det-cases')?.value) || 0;
            const loose = parseInt(tr.querySelector('.tps-det-loose')?.value) || 0;
            const rate  = parseFloat(tr.querySelector('.tps-det-rate')?.value) || 0;
            const btl = (cases * (sizeObj.bpc || 0)) + loose;
            totalCases += cases;
            totalLoose += loose;
            totalBtls  += btl;
            totalAmt   += btl * rate;
        });
        const el = id => document.getElementById(id);
        if (el('tpsDetTotalCases')) { el('tpsDetTotalCases').textContent = totalCases; el('tpsDetTotalCases').style.textAlign = 'right'; }
        if (el('tpsDetTotalLoose')) { el('tpsDetTotalLoose').textContent = totalLoose; el('tpsDetTotalLoose').style.textAlign = 'right'; }
        if (el('tpsDetTotalBtls'))  { el('tpsDetTotalBtls').textContent = totalBtls;  el('tpsDetTotalBtls').style.textAlign = 'right'; el('tpsDetTotalBtls').style.fontWeight = '700'; }
        if (el('tpsDetTotalAmt'))   { el('tpsDetTotalAmt').textContent = '₹' + totalAmt.toLocaleString('en-IN', { minimumFractionDigits: 2 }); el('tpsDetTotalAmt').style.textAlign = 'right'; }
    }

    /* ── Create one editable item row for the dialog ── */
    function createTpsDetItemRow(item, idx) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="color:var(--text-muted);text-align:center;width:30px">${idx + 1}</td>
            <td style="position:relative">
                <input class="tps-cell-input tps-det-brand" type="text" value="${esc(item.brand || '')}" placeholder="Brand or shortcode…" autocomplete="off">
                <div class="tp-prod-ac hidden tps-det-ac-dropdown" data-role="det-prod-ac"></div>
            </td>
            <td><input class="tps-cell-input tps-det-code" type="text" value="${esc(item.code || '')}" placeholder="Code" style="width:140px"></td>
            <td><select class="tps-cell-select tps-det-size">${tpsDetSizeOptions(item.size || item.sizeLabel || '')}</select></td>
            <td><input class="tps-cell-input tps-det-cases" type="number" value="${item.cases || 0}" min="0" style="width:58px"></td>
            <td><input class="tps-cell-input tps-det-loose" type="number" value="${item.loose || 0}" min="0" style="width:58px"></td>
            <td class="tps-cell-computed tps-det-totalbtl">${item.totalBtl || 0}</td>
            <td><input class="tps-cell-input tps-det-rate" type="number" value="${(item.rate || 0).toFixed(2)}" min="0" step="0.01" style="width:80px"></td>
            <td class="tps-cell-computed tps-cell-amt tps-det-amount">₹${(item.amount || 0).toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
            <td style="text-align:center;width:32px">
                <button class="tps-row-remove" title="Remove row">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
                </button>
            </td>
        `;
        // Wire recalc on input
        tr.querySelectorAll('.tps-det-cases, .tps-det-loose, .tps-det-rate').forEach(inp => {
            inp.addEventListener('input', () => recalcTpsDetRow(tr));
        });
        tr.querySelector('.tps-det-size')?.addEventListener('change', () => recalcTpsDetRow(tr));
        // Wire remove
        tr.querySelector('.tps-row-remove')?.addEventListener('click', () => {
            tr.remove();
            renumberTpsDetRows();
            recalcTpsDetTotals();
        });

        // Wire keyboard navigation on each input/select in this row
        const editables = tr.querySelectorAll('input, select');
        editables.forEach(el => {
            el.addEventListener('keydown', (e) => tpsDetGridKeyHandler(e, tr, el));
        });

        // Wire product autocomplete on brand field
        wireDetProdAc(tr);

        return tr;
    }

    /* ── Product Autocomplete for Edit Dialog ── */
    let tpsDetAcActive = null; // { tr, dropdown, focusIdx, combined }

    function wireDetProdAc(tr) {
        const brandInput = tr.querySelector('.tps-det-brand');
        const dropdown   = tr.querySelector('[data-role="det-prod-ac"]');
        if (!brandInput || !dropdown) return;

        let acIdx = -1;

        function showDetAc(query) {
            const q = (query || '').toLowerCase();

            // shortcut matches first
            const scMatches = getShortcutMatches(query);
            const scProducts = scMatches.map(sc => {
                const prod = allProducts.find(p => p.id === sc.productId);
                return prod ? { ...prod, _shortcutCode: sc.code } : {
                    id: sc.productId, brandName: sc.brandName, code: sc.prodCode,
                    size: sc.size, category: sc.category, mrp: sc.mrp, costPrice: 0,
                    _shortcutCode: sc.code
                };
            });
            const scProdIds = new Set(scProducts.map(p => p.id));

            // text matches
            const textMatches = q.length > 0
                ? allProducts.filter(p =>
                    !scProdIds.has(p.id) && (
                    (p.brandName || '').toLowerCase().includes(q) ||
                    (p.code || '').toLowerCase().includes(q) ||
                    (p.category || '').toLowerCase().includes(q))
                  )
                : allProducts.filter(p => !scProdIds.has(p.id)).slice(0, 30);

            const combined = [...scProducts, ...textMatches];

            if (combined.length === 0 && q.length > 0) {
                dropdown.innerHTML = '<div class="tp-prod-ac-empty">No matching products</div>';
            } else if (combined.length === 0) {
                dropdown.innerHTML = '<div class="tp-prod-ac-empty">No products in master</div>';
            } else {
                dropdown.innerHTML = combined.map((p, i) => `
                    <div class="tp-prod-ac-item${p._shortcutCode ? ' sc-match' : ''}" data-idx="${i}">
                        ${p._shortcutCode ? '<span class="tp-prod-ac-sc-badge">' + esc(p._shortcutCode) + '</span>' : ''}
                        <span class="tp-prod-ac-brand">${esc(p.brandName)}</span>
                        <span class="tp-prod-ac-meta">${esc(p.code || '')} · ${p.size ? p.size + 'ml' : ''} · ${esc(p.category || '')}</span>
                    </div>
                `).join('');

                dropdown.querySelectorAll('.tp-prod-ac-item').forEach(item => {
                    item.addEventListener('mousedown', (e) => {
                        e.preventDefault();
                        const idx = parseInt(item.dataset.idx);
                        selectDetProdAcItem(tr, combined[idx], brandInput, dropdown);
                    });
                });
            }

            dropdown.classList.remove('hidden');
            acIdx = -1;
            tpsDetAcActive = { tr, dropdown, focusIdx: -1, combined };
        }

        function hideDetAc() {
            dropdown.classList.add('hidden');
            dropdown.innerHTML = '';
            acIdx = -1;
            if (tpsDetAcActive && tpsDetAcActive.tr === tr) tpsDetAcActive = null;
        }

        function focusDetAcItem(newIdx) {
            const items = dropdown.querySelectorAll('.tp-prod-ac-item');
            if (items.length === 0) return;
            acIdx = Math.max(0, Math.min(newIdx, items.length - 1));
            items.forEach((it, i) => it.classList.toggle('focused', i === acIdx));
            items[acIdx]?.scrollIntoView({ block: 'nearest' });
            if (tpsDetAcActive) tpsDetAcActive.focusIdx = acIdx;
        }

        brandInput.addEventListener('input', () => {
            showDetAc(brandInput.value);
        });

        brandInput.addEventListener('focus', () => {
            // Close any other open AC in the dialog
            if (tpsDetAcActive && tpsDetAcActive.tr !== tr) {
                tpsDetAcActive.dropdown.classList.add('hidden');
                tpsDetAcActive = null;
            }
            if (allProducts.length > 0) {
                showDetAc(brandInput.value);
            }
        });

        brandInput.addEventListener('blur', () => {
            setTimeout(() => hideDetAc(), 150);
        });

        brandInput.addEventListener('keydown', (e) => {
            // Shortcut quick-select: Enter/Tab on exact shortcode match
            if (e.key === 'Enter' || e.key === 'Tab') {
                const val = brandInput.value.trim();
                if (val && acIdx < 0) {
                    const scProd = resolveShortcutToProduct(val);
                    if (scProd) {
                        e.preventDefault(); e.stopPropagation();
                        selectDetProdAcItem(tr, scProd, brandInput, dropdown);
                        return;
                    }
                }
            }

            if (dropdown.classList.contains('hidden')) return;
            const items = dropdown.querySelectorAll('.tp-prod-ac-item');

            if (e.key === 'ArrowDown') {
                e.preventDefault(); e.stopPropagation();
                focusDetAcItem(acIdx + 1);
            } else if (e.key === 'ArrowUp') {
                e.preventDefault(); e.stopPropagation();
                focusDetAcItem(acIdx - 1);
            } else if (e.key === 'Enter' && acIdx >= 0 && items[acIdx]) {
                e.preventDefault(); e.stopPropagation();
                const combined = tpsDetAcActive?.combined;
                if (combined && combined[acIdx]) {
                    selectDetProdAcItem(tr, combined[acIdx], brandInput, dropdown);
                }
            } else if (e.key === 'Escape') {
                e.preventDefault(); e.stopPropagation();
                hideDetAc();
            }
        });
    }

    function selectDetProdAcItem(tr, product, brandInput, dropdown) {
        // Fill brand name
        brandInput.value = product.brandName || '';

        // Fill code
        const codeEl = tr.querySelector('.tps-det-code');
        if (codeEl) codeEl.value = product.code || '';

        // Fill size
        const sizeEl = tr.querySelector('.tps-det-size');
        if (sizeEl && product.size) sizeEl.value = product.size;

        // Fill rate from costPrice or mrp
        const rateEl = tr.querySelector('.tps-det-rate');
        if (rateEl) {
            const rate = product.costPrice || product.mrp || 0;
            if (rate > 0) rateEl.value = rate;
        }

        // Recalculate row
        recalcTpsDetRow(tr);

        // Hide dropdown
        dropdown.classList.add('hidden');
        dropdown.innerHTML = '';
        if (tpsDetAcActive && tpsDetAcActive.tr === tr) tpsDetAcActive = null;

        // Move focus to cases field
        const casesEl = tr.querySelector('.tps-det-cases');
        if (casesEl) casesEl.focus();
    }

    /* ── Get all focusable inputs/selects in a row ── */
    function getTpsDetRowEditables(tr) {
        return [...tr.querySelectorAll('input, select')];
    }

    /* ── Keyboard handler for grid cells ── */
    function tpsDetGridKeyHandler(e, tr, el) {
        // If autocomplete dropdown is open on this brand input, let the AC handler take over
        if (el.classList.contains('tps-det-brand') && tpsDetAcActive && tpsDetAcActive.tr === tr) {
            const dropdown = tr.querySelector('[data-role="det-prod-ac"]');
            if (dropdown && !dropdown.classList.contains('hidden')) {
                if (['ArrowDown', 'ArrowUp', 'Enter', 'Escape'].includes(e.key)) return;
            }
        }

        const gridBody = document.getElementById('tpsDetGridBody');
        if (!gridBody) return;
        const rows = [...gridBody.querySelectorAll('tr')];
        const rowIdx = rows.indexOf(tr);
        const cellsInRow = getTpsDetRowEditables(tr);
        const cellIdx = cellsInRow.indexOf(el);

        if (e.key === 'Enter' && !e.ctrlKey && !e.shiftKey) {
            e.preventDefault();
            // Move to next cell; if last cell in row, move to first cell of next row
            if (cellIdx < cellsInRow.length - 1) {
                cellsInRow[cellIdx + 1].focus();
                cellsInRow[cellIdx + 1].select?.();
            } else if (rowIdx < rows.length - 1) {
                const nextCells = getTpsDetRowEditables(rows[rowIdx + 1]);
                if (nextCells.length) { nextCells[0].focus(); nextCells[0].select?.(); }
            } else {
                // Last cell of last row → add new row and focus it
                addTpsDetEmptyRow();
                const newRows = [...gridBody.querySelectorAll('tr')];
                const lastRow = newRows[newRows.length - 1];
                const firstCell = getTpsDetRowEditables(lastRow)[0];
                if (firstCell) { firstCell.focus(); firstCell.select?.(); }
            }
            return;
        }

        if (e.key === 'ArrowDown' && !e.shiftKey) {
            if (el.tagName === 'SELECT') return; // let select handle its own arrow
            e.preventDefault();
            if (rowIdx < rows.length - 1) {
                const targetCells = getTpsDetRowEditables(rows[rowIdx + 1]);
                const target = targetCells[Math.min(cellIdx, targetCells.length - 1)];
                if (target) { target.focus(); target.select?.(); }
            }
            return;
        }

        if (e.key === 'ArrowUp' && !e.shiftKey) {
            if (el.tagName === 'SELECT') return;
            e.preventDefault();
            if (rowIdx > 0) {
                const targetCells = getTpsDetRowEditables(rows[rowIdx - 1]);
                const target = targetCells[Math.min(cellIdx, targetCells.length - 1)];
                if (target) { target.focus(); target.select?.(); }
            }
            return;
        }
    }

    /* ── Add empty row to dialog grid ── */
    function addTpsDetEmptyRow() {
        const gridBody = document.getElementById('tpsDetGridBody');
        if (!gridBody) return;
        const idx = gridBody.querySelectorAll('tr').length;
        const emptyItem = { brand: '', code: '', size: '', cases: 0, loose: 0, totalBtl: 0, rate: 0, amount: 0 };
        gridBody.appendChild(createTpsDetItemRow(emptyItem, idx));
        recalcTpsDetTotals();
    }

    /* ── Renumber row #s ── */
    function renumberTpsDetRows() {
        const gridBody = document.getElementById('tpsDetGridBody');
        if (!gridBody) return;
        gridBody.querySelectorAll('tr').forEach((tr, i) => {
            const numTd = tr.querySelector('td:first-child');
            if (numTd) numTd.textContent = i + 1;
        });
    }

    /* ── Open edit dialog ── */
    function openTpsDetail(tp) {
        tpsViewingId = tp.id;
        if (!tpsDetailOverlay) return;
        tpsDetailOverlay.classList.remove('hidden');

        // Highlight row in background table
        tpsTableBody?.querySelectorAll('tr').forEach(r => {
            r.classList.toggle('tps-row-selected', r.dataset.tpId === tp.id);
        });

        // Title
        const titleEl = document.getElementById('tpsDetTitle');
        if (titleEl) titleEl.textContent = 'Edit TP ' + (tp.tpNumber || '');

        // Populate editable meta fields
        const setVal = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ''; };
        setVal('tpsDetTpNumber', tp.tpNumber);
        setVal('tpsDetSupplier', tp.supplier);
        setVal('tpsDetTpDate', tp.tpDate);
        setVal('tpsDetRecvDate', tp.receivedDate);
        setVal('tpsDetVehicle', tp.vehicle);
        setVal('tpsDetNarration', tp.narration);

        // Build editable items grid
        const gridBody = document.getElementById('tpsDetGridBody');
        if (gridBody) {
            gridBody.innerHTML = '';
            const items = tp.items || [];
            items.forEach((item, i) => {
                gridBody.appendChild(createTpsDetItemRow(item, i));
            });
            recalcTpsDetTotals();
        }
    }

    /* ── Close dialog ── */
    function closeTpsDetail() {
        tpsViewingId = null;
        if (tpsDetailOverlay) tpsDetailOverlay.classList.add('hidden');
        tpsTableBody?.querySelectorAll('tr').forEach(r => r.classList.remove('tps-row-selected'));
    }

    /* ── Save edited TP ── */
    async function saveTpsDetailEdit() {
        const tpNumber  = document.getElementById('tpsDetTpNumber')?.value.trim();
        const supplier  = document.getElementById('tpsDetSupplier')?.value.trim();

        if (!tpNumber) {
            document.getElementById('tpsDetTpNumber')?.focus();
            showTpToast('TP Number is required', true);
            return;
        }
        if (!supplier) {
            document.getElementById('tpsDetSupplier')?.focus();
            showTpToast('Supplier name is required', true);
            return;
        }

        // Gather items from dialog grid
        const gridBody = document.getElementById('tpsDetGridBody');
        const items = [];
        gridBody?.querySelectorAll('tr').forEach(tr => {
            const brand = tr.querySelector('.tps-det-brand')?.value.trim();
            if (!brand) return;
            const code    = tr.querySelector('.tps-det-code')?.value.trim() || '';
            const sizeVal = tr.querySelector('.tps-det-size')?.value || '';
            const sizeObj = TP_SIZES.find(s => s.value === sizeVal) || TP_SIZES[0];
            const cases   = parseInt(tr.querySelector('.tps-det-cases')?.value) || 0;
            const loose   = parseInt(tr.querySelector('.tps-det-loose')?.value) || 0;
            const rate    = parseFloat(tr.querySelector('.tps-det-rate')?.value) || 0;
            const totalBtl = (cases * sizeObj.bpc) + loose;
            items.push({
                brand,
                code,
                size: sizeVal,
                sizeLabel: sizeObj.label,
                cases,
                bpc: sizeObj.bpc,
                loose,
                totalBtl,
                rate,
                amount: totalBtl * rate,
            });
        });

        if (items.length === 0) {
            showTpToast('Add at least one item', true);
            return;
        }

        // Find original TP to preserve id + createdAt
        const origTp = allTps.find(t => t.id === tpsViewingId);
        if (!origTp) {
            showTpToast('TP not found — it may have been deleted', true);
            closeTpsDetail();
            return;
        }

        const updatedTp = {
            ...origTp,
            tpNumber,
            tpDate:       document.getElementById('tpsDetTpDate')?.value || origTp.tpDate,
            receivedDate: document.getElementById('tpsDetRecvDate')?.value || origTp.receivedDate,
            supplier,
            vehicle:      document.getElementById('tpsDetVehicle')?.value.trim() || '',
            narration:    document.getElementById('tpsDetNarration')?.value.trim() || '',
            items,
            totalCases:   items.reduce((s, i) => s + i.cases, 0),
            totalLoose:   items.reduce((s, i) => s + i.loose, 0),
            totalBottles: items.reduce((s, i) => s + i.totalBtl, 0),
            totalAmount:  items.reduce((s, i) => s + i.amount, 0),
        };

        try {
            const result = await window.electronAPI.saveTp({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || '',
                tp: updatedTp,
            });
            if (result.success) {
                showTpToast(`TP "${tpNumber}" updated — ${items.length} items, ₹${updatedTp.totalAmount.toFixed(2)}`);
                closeTpsDetail();
                await loadTpSummary();
            } else {
                showTpToast('Save failed: ' + result.error, true);
            }
        } catch (err) {
            showTpToast('Save error: ' + err.message, true);
        }
    }

    /* ── Delete TP ── */
    async function deleteTpFromSummary(tp) {
        if (!confirm(`Delete TP "${tp.tpNumber}"?\n\nThis cannot be undone.`)) return;
        try {
            const result = await window.electronAPI.deleteTp({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || '',
                tpId: tp.id,
            });
            if (result.success) {
                showTpToast(`TP "${tp.tpNumber}" deleted`);
                closeTpsDetail();
                await loadTpSummary();
            } else {
                showTpToast('Delete failed: ' + result.error, true);
            }
        } catch (err) {
            showTpToast('Delete error: ' + err.message, true);
        }
    }

    /* ── Export TPs to Excel ── */
    async function exportTpSummary() {
        if (filteredTps.length === 0) {
            showTpToast('No TPs to export', true);
            return;
        }
        try {
            // Summary sheet
            const summaryRows = filteredTps.map((tp, i) => ({
                '#': tp._srNo ?? (i + 1),
                'TP Number': tp.tpNumber || '',
                'TP Date': tp.tpDate || '',
                'Received Date': tp.receivedDate || '',
                'Supplier': tp.supplier || '',
                'Items': (tp.items || []).length,
                'Cases': tp.totalCases || 0,
                'Loose': tp.totalLoose || 0,
                'Bottles': tp.totalBottles || 0,
                'Amount': tp.totalAmount || 0,
                'Vehicle': tp.vehicle || '',
            }));

            // Detail sheet with all items
            const detailRows = [];
            filteredTps.forEach(tp => {
                (tp.items || []).forEach((item, i) => {
                    detailRows.push({
                        'TP Number': tp.tpNumber || '',
                        'TP Date': tp.tpDate || '',
                        'Supplier': tp.supplier || '',
                        '#': i + 1,
                        'Brand': item.brand || '',
                        'Code': item.code || '',
                        'Size': item.sizeLabel || item.size || '',
                        'Cases': item.cases || 0,
                        'BPC': item.bpc || 0,
                        'Loose': item.loose || 0,
                        'Total Btls': item.totalBtl || 0,
                        'Rate': item.rate || 0,
                        'Amount': item.amount || 0,
                    });
                });
            });

            const result = await window.electronAPI.exportTpSummary({
                summaryRows,
                detailRows,
                barName: activeBar.barName || '',
            });

            if (result.success) {
                showTpToast(`Exported ${filteredTps.length} TPs to Excel`);
            } else if (result.error !== 'Cancelled') {
                showTpToast('Export failed: ' + result.error, true);
            }
        } catch (err) {
            showTpToast('Export error: ' + err.message, true);
        }
    }

    /* ── Wire filter controls ── */
    if (tpsDateFrom)  tpsDateFrom.addEventListener('change', applyTpsFilters);
    if (tpsDateTo)    tpsDateTo.addEventListener('change', applyTpsFilters);
    if (tpsSupFilter)  tpsSupFilter.addEventListener('change', applyTpsFilters);
    if (tpsSearchEl) {
        tpsSearchEl.addEventListener('input', applyTpsFilters);
    }

    /* ── Wire buttons ── */
    if (tpsExportBtn) tpsExportBtn.addEventListener('click', exportTpSummary);
    if (tpsAddNewBtn) tpsAddNewBtn.addEventListener('click', () => {
        switchPurchaseSub('add-tp');
    });
    if (tpsDetCloseBtn) tpsDetCloseBtn.addEventListener('click', closeTpsDetail);
    const tpsDetCancelBtn = document.getElementById('tpsDetCancelBtn');
    if (tpsDetCancelBtn) tpsDetCancelBtn.addEventListener('click', closeTpsDetail);
    const tpsDetSaveBtn = document.getElementById('tpsDetSaveBtn');
    if (tpsDetSaveBtn) tpsDetSaveBtn.addEventListener('click', saveTpsDetailEdit);
    if (tpsDetDeleteBtn) tpsDetDeleteBtn.addEventListener('click', () => {
        const tp = allTps.find(t => t.id === tpsViewingId);
        if (tp) deleteTpFromSummary(tp);
    });

    /* ── Close detail on overlay click ── */
    if (tpsDetailOverlay) {
        tpsDetailOverlay.addEventListener('click', (e) => {
            if (e.target === tpsDetailOverlay) closeTpsDetail();
        });
    }

    /* ── Global keyboard shortcuts for TP edit dialog ── */
    document.addEventListener('keydown', (e) => {
        if (!tpsDetailOverlay || tpsDetailOverlay.classList.contains('hidden')) return;

        // Escape → close dialog
        if (e.key === 'Escape') {
            e.preventDefault();
            closeTpsDetail();
            return;
        }

        // Ctrl+S → save changes
        if (e.key === 's' && (e.ctrlKey || e.metaKey) && !e.shiftKey) {
            e.preventDefault();
            saveTpsDetailEdit();
            return;
        }

        // Ctrl+N → add new empty row
        if (e.key === 'n' && (e.ctrlKey || e.metaKey) && !e.shiftKey) {
            e.preventDefault();
            addTpsDetEmptyRow();
            const gridBody = document.getElementById('tpsDetGridBody');
            if (gridBody) {
                const allRows = [...gridBody.querySelectorAll('tr')];
                const lastRow = allRows[allRows.length - 1];
                if (lastRow) {
                    const first = getTpsDetRowEditables(lastRow)[0];
                    if (first) { first.focus(); first.select?.(); }
                }
            }
            return;
        }

        // Ctrl+Delete → remove current row
        if (e.key === 'Delete' && (e.ctrlKey || e.metaKey)) {
            e.preventDefault();
            const gridBody = document.getElementById('tpsDetGridBody');
            if (!gridBody) return;
            const focused = document.activeElement;
            const tr = focused?.closest?.('#tpsDetGridBody tr');
            if (!tr) return;
            const rows = [...gridBody.querySelectorAll('tr')];
            const rowIdx = rows.indexOf(tr);
            if (rows.length <= 1) return; // don't delete last row
            tr.remove();
            renumberTpsDetRows();
            recalcTpsDetTotals();
            // Focus same position in nearest row
            const remaining = [...gridBody.querySelectorAll('tr')];
            const nextRow = remaining[Math.min(rowIdx, remaining.length - 1)];
            if (nextRow) {
                const cells = getTpsDetRowEditables(nextRow);
                if (cells.length) { cells[0].focus(); cells[0].select?.(); }
            }
            return;
        }
    });

    /* ── Load TP summary on navigation ── */
    // Observer pattern: reload when tp-summary becomes visible
    const tpsSummaryPanel = document.getElementById('sub-tp-summary');
    if (tpsSummaryPanel) {
        const tpsObserver = new MutationObserver(() => {
            if (tpsSummaryPanel.classList.contains('active')) {
                loadTpSummary();
            }
        });
        tpsObserver.observe(tpsSummaryPanel, { attributes: true, attributeFilter: ['class'] });
    }

    // Also load immediately (in case already visible)
    loadTpSummary();

    // ═══════════════════════════════════════════════════════
    //  PURCHASE REPORT MODULE
    // ═══════════════════════════════════════════════════════
    (function initPurchaseReport() {
        // Elements
        const prDateFrom      = document.getElementById('prDateFrom');
        const prDateTo        = document.getElementById('prDateTo');
        const prSupFilter     = document.getElementById('prSupplierFilter');
        const prGroupBy       = document.getElementById('prGroupBy');
        const prSearch        = document.getElementById('prSearch');
        const prResetBtn      = document.getElementById('prResetBtn');
        const prExportBtn     = document.getElementById('prExportBtn');
        const prTableBody     = document.getElementById('prTableBody');
        const prEmptyState    = document.getElementById('prEmptyState');
        // Cards
        const prBadge         = document.getElementById('prBadge');
        const prCardTps       = document.getElementById('prCardTps');
        const prCardItems     = document.getElementById('prCardItems');
        const prCardCases     = document.getElementById('prCardCases');
        const prCardBottles   = document.getElementById('prCardBottles');
        const prCardAmount    = document.getElementById('prCardAmount');
        // Footer totals
        const prFtCases       = document.getElementById('prFtCases');
        const prFtLoose       = document.getElementById('prFtLoose');
        const prFtBtls        = document.getElementById('prFtBtls');
        const prFtAmt         = document.getElementById('prFtAmt');

        // Sort state
        let prSortCol = 'date', prSortDir = 'desc';

        // All flat item rows built from TPs
        let prAllRows = [];

        function fmtDate(d) {
            if (!d) return '—';
            const dt = new Date(d);
            if (isNaN(dt)) return d;
            return dt.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
        }
        function fmtAmt(v) {
            return '₹' + Number(v || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }
        function escPr(s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

        async function loadPurchaseReport() {
            if (!activeBar.barName) return;
            try {
                const res = await window.electronAPI.getTps({
                    barName: activeBar.barName,
                    financialYear: activeBar.financialYear || '',
                });
                const tps = res.success ? (res.tps || []) : [];

                // Build flat rows: one per item per TP
                prAllRows = [];
                const uniqueTpIds = new Set();
                for (const tp of tps) {
                    uniqueTpIds.add(tp.id);
                    for (const item of (tp.items || [])) {
                        prAllRows.push({
                            tpId       : tp.id,
                            tpNumber   : tp.tpNumber || '',
                            tpDate     : tp.tpDate || tp.receivedDate || '',
                            receivedDate: tp.receivedDate || tp.tpDate || '',
                            supplier   : tp.supplier || '',
                            brand      : item.brand || item.brandName || '',
                            code       : item.code || '',
                            size       : item.size || item.sizeLabel || '',
                            cases      : item.cases || 0,
                            loose      : item.loose || 0,
                            totalBtl   : item.totalBtl || 0,
                            rate       : item.rate || 0,
                            amount     : item.amount || 0,
                        });
                    }
                }

                // Populate supplier dropdown
                const suppliers = [...new Set(prAllRows.map(r => r.supplier).filter(Boolean))].sort();
                const curSup = prSupFilter ? prSupFilter.value : '';
                if (prSupFilter) {
                    prSupFilter.innerHTML = '<option value="">All Suppliers</option>' +
                        suppliers.map(s => `<option value="${escPr(s)}"${s === curSup ? ' selected' : ''}>${escPr(s)}</option>`).join('');
                }

                // Set default date range to FY if empty
                if (prDateFrom && !prDateFrom.value && prDateTo && !prDateTo.value) {
                    const fyStart = getFyStartDate();
                    const today = new Date().toISOString().slice(0, 10);
                    if (fyStart) prDateFrom.value = fyStart;
                    prDateTo.value = today;
                }

                applyPrFilters();
            } catch (err) {
                console.error('loadPurchaseReport error:', err);
            }
        }

        function getFyStartDate() {
            const fy = activeBar.financialYear || '';
            const m = fy.match(/(\d{4})-(\d{2}|\d{4})/);
            if (!m) return '';
            return m[1] + '-04-01'; // FY starts April 1
        }

        function applyPrFilters() {
            const from    = prDateFrom ? prDateFrom.value : '';
            const to      = prDateTo   ? prDateTo.value   : '';
            const sup     = prSupFilter ? prSupFilter.value : '';
            const search  = prSearch ? prSearch.value.trim().toLowerCase() : '';
            const groupBy = prGroupBy ? prGroupBy.value : 'none';

            let rows = prAllRows.filter(r => {
                const dt = r.tpDate;
                if (from && dt && dt < from) return false;
                if (to   && dt && dt > to)   return false;
                if (sup  && r.supplier !== sup) return false;
                if (search) {
                    const haystack = [r.tpNumber, r.supplier, r.brand, r.code, r.size].join(' ').toLowerCase();
                    if (!haystack.includes(search)) return false;
                }
                return true;
            });

            // Sort
            rows.sort((a, b) => {
                let va = '', vb = '';
                if (prSortCol === 'date')     { va = a.tpDate;    vb = b.tpDate; }
                if (prSortCol === 'supplier') { va = a.supplier;  vb = b.supplier; }
                if (prSortCol === 'brand')    { va = a.brand;     vb = b.brand; }
                const cmp = va < vb ? -1 : va > vb ? 1 : 0;
                return prSortDir === 'asc' ? cmp : -cmp;
            });

            renderPrTable(rows, groupBy);
        }

        function renderPrTable(rows, groupBy) {
            if (!prTableBody) return;
            prTableBody.innerHTML = '';

            if (rows.length === 0) {
                if (prEmptyState) prEmptyState.classList.remove('hidden');
                updatePrCards([], 0);
                return;
            }
            if (prEmptyState) prEmptyState.classList.add('hidden');

            let totalCases = 0, totalLoose = 0, totalBtls = 0, totalAmt = 0;
            const uniqueTps = new Set();
            rows.forEach(r => { uniqueTps.add(r.tpId); totalCases += r.cases; totalLoose += r.loose; totalBtls += r.totalBtl; totalAmt += r.amount; });

            let srNo = 0;

            if (groupBy === 'none') {
                rows.forEach(r => {
                    srNo++;
                    prTableBody.insertAdjacentHTML('beforeend', buildPrRow(r, srNo));
                });
            } else {
                // Get group key function
                const keyOf = r => groupBy === 'tp' ? r.tpNumber : groupBy === 'supplier' ? r.supplier : r.brand;
                const groups = {};
                const groupOrder = [];
                rows.forEach(r => {
                    const k = keyOf(r) || '—';
                    if (!groups[k]) { groups[k] = []; groupOrder.push(k); }
                    groups[k].push(r);
                });

                for (const key of groupOrder) {
                    const gRows = groups[key];
                    const gCases = gRows.reduce((s,r) => s + r.cases, 0);
                    const gLoose = gRows.reduce((s,r) => s + r.loose, 0);
                    const gBtls  = gRows.reduce((s,r) => s + r.totalBtl, 0);
                    const gAmt   = gRows.reduce((s,r) => s + r.amount, 0);

                    // Group header
                    prTableBody.insertAdjacentHTML('beforeend',
                        `<tr class="pr-group-row">
                            <td colspan="13">${escPr(key)} &nbsp;<span style="font-weight:400;color:var(--text-muted);font-size:11px">${gRows.length} items</span></td>
                        </tr>`);

                    gRows.forEach(r => {
                        srNo++;
                        prTableBody.insertAdjacentHTML('beforeend', buildPrRow(r, srNo));
                    });

                    // Subtotal
                    prTableBody.insertAdjacentHTML('beforeend',
                        `<tr class="pr-subtotal-row">
                            <td colspan="8" style="text-align:right;padding-right:12px">Subtotal</td>
                            <td class="text-right">${gCases}</td>
                            <td class="text-right">${gLoose}</td>
                            <td class="text-right">${gBtls}</td>
                            <td></td>
                            <td class="text-right">${fmtAmt(gAmt)}</td>
                        </tr>`);
                }
            }

            // Footer totals
            if (prFtCases)  prFtCases.textContent  = totalCases;
            if (prFtLoose)  prFtLoose.textContent  = totalLoose;
            if (prFtBtls)   prFtBtls.textContent   = totalBtls;
            if (prFtAmt)    prFtAmt.textContent    = fmtAmt(totalAmt);

            updatePrCards(rows, uniqueTps.size);
        }

        function buildPrRow(r, sr) {
            return `<tr>
                <td class="pr-th-sr" style="text-align:center;color:var(--text-muted);font-size:11px">${sr}</td>
                <td class="pr-td-muted">${fmtDate(r.tpDate)}</td>
                <td class="pr-td-muted">${fmtDate(r.receivedDate)}</td>
                <td class="pr-td-muted" style="font-weight:500;color:var(--text-primary)">${escPr(r.tpNumber)}</td>
                <td style="max-width:150px;overflow:hidden;text-overflow:ellipsis">${escPr(r.supplier)}</td>
                <td class="pr-td-brand">${escPr(r.brand)}</td>
                <td class="pr-td-code">${escPr(r.code)}</td>
                <td class="pr-td-muted">${escPr(r.size)}</td>
                <td class="text-right">${r.cases}</td>
                <td class="text-right pr-td-muted">${r.loose || 0}</td>
                <td class="text-right" style="font-weight:500">${r.totalBtl}</td>
                <td class="text-right pr-td-muted">${r.rate > 0 ? r.rate.toFixed(2) : '—'}</td>
                <td class="pr-td-amt">${fmtAmt(r.amount)}</td>
            </tr>`;
        }

        function updatePrCards(rows, tpCount) {
            const items   = rows.length;
            const cases   = rows.reduce((s, r) => s + r.cases, 0);
            const bottles = rows.reduce((s, r) => s + r.totalBtl, 0);
            const amount  = rows.reduce((s, r) => s + r.amount, 0);

            if (prBadge)       prBadge.textContent       = `${items} items · ${fmtAmt(amount)}`;
            if (prCardTps)     prCardTps.textContent     = tpCount;
            if (prCardItems)   prCardItems.textContent   = items;
            if (prCardCases)   prCardCases.textContent   = cases;
            if (prCardBottles) prCardBottles.textContent = bottles;
            if (prCardAmount)  prCardAmount.textContent  = fmtAmt(amount);
        }

        // ── Sort click ──
        const prTable = document.getElementById('prTable');
        if (prTable) {
            prTable.querySelectorAll('thead th[data-sort]').forEach(th => {
                th.addEventListener('click', () => {
                    const col = th.dataset.sort;
                    if (prSortCol === col) prSortDir = prSortDir === 'asc' ? 'desc' : 'asc';
                    else { prSortCol = col; prSortDir = 'asc'; }
                    prTable.querySelectorAll('thead th').forEach(t => t.classList.remove('pr-sort-active'));
                    th.classList.add('pr-sort-active');
                    const icon = th.querySelector('.pr-sort-icon');
                    if (icon) icon.textContent = prSortDir === 'asc' ? '↑' : '↓';
                    applyPrFilters();
                });
            });
        }

        // ── Filter event listeners ──
        [prDateFrom, prDateTo, prSupFilter, prGroupBy].forEach(el => {
            if (el) el.addEventListener('change', applyPrFilters);
        });
        if (prSearch) {
            prSearch.addEventListener('input', applyPrFilters);
            prSearch.addEventListener('keydown', e => { if (e.key === 'Escape') { prSearch.value = ''; applyPrFilters(); } });
        }

        // ── Reset ──
        if (prResetBtn) {
            prResetBtn.addEventListener('click', () => {
                if (prDateFrom) prDateFrom.value = getFyStartDate();
                if (prDateTo)   prDateTo.value   = new Date().toISOString().slice(0, 10);
                if (prSupFilter) prSupFilter.value = '';
                if (prGroupBy)   prGroupBy.value  = 'none';
                if (prSearch)    prSearch.value   = '';
                applyPrFilters();
            });
        }

        // ── Export ──
        if (prExportBtn) {
            prExportBtn.addEventListener('click', async () => {
                const from    = prDateFrom ? prDateFrom.value : '';
                const to      = prDateTo   ? prDateTo.value   : '';
                const sup     = prSupFilter ? prSupFilter.value : '';
                const search  = prSearch ? prSearch.value.trim().toLowerCase() : '';

                const rows = prAllRows.filter(r => {
                    if (from && r.tpDate && r.tpDate < from) return false;
                    if (to   && r.tpDate && r.tpDate > to)   return false;
                    if (sup  && r.supplier !== sup) return false;
                    if (search) {
                        const h = [r.tpNumber, r.supplier, r.brand, r.code, r.size].join(' ').toLowerCase();
                        if (!h.includes(search)) return false;
                    }
                    return true;
                });

                if (rows.length === 0) { showTpToast('No data to export', true); return; }

                const exportRows = rows.map(r => ({
                    'TP Date'      : r.tpDate,
                    'Received Date': r.receivedDate,
                    'TP No.'       : r.tpNumber,
                    'Supplier'     : r.supplier,
                    'Brand'        : r.brand,
                    'Code'         : r.code,
                    'Size'         : r.size,
                    'Cases'        : r.cases,
                    'Loose Btl'    : r.loose,
                    'Total Btls'   : r.totalBtl,
                    'Rate (₹)'     : r.rate,
                    'Amount (₹)'   : r.amount,
                }));

                try {
                    prExportBtn.disabled = true;
                    const res = await window.electronAPI.exportPurchaseReport({
                        rows: exportRows,
                        barName: activeBar.barName || 'Bar',
                    });
                    if (res.success) showTpToast(`Exported ${rows.length} rows`);
                    else if (res.error !== 'Cancelled') showTpToast('Export failed: ' + res.error, true);
                } catch (e) {
                    showTpToast('Export error: ' + e.message, true);
                } finally {
                    prExportBtn.disabled = false;
                }
            });
        }

        // ── Observer: load when sub-view becomes visible ──
        const prPanel = document.getElementById('sub-purchase-report');
        if (prPanel) {
            const obs = new MutationObserver(() => {
                if (prPanel.classList.contains('active') && prAllRows.length === 0) {
                    loadPurchaseReport();
                }
            });
            obs.observe(prPanel, { attributes: true, attributeFilter: ['class'] });
        }

        // Expose load function for switchPurchaseSub
        window._loadPurchaseReport = loadPurchaseReport;
    })();
    // ═══════════════════════════════════════════════════════

    const invGridBody    = document.getElementById('invGridBody');
    const invSearchInput = document.getElementById('invSearchInput');
    const invCatFilter   = document.getElementById('invCategoryFilter');
    const invSizeFilter  = document.getElementById('invSizeFilter');
    const invStockFilter = document.getElementById('invStockFilter');
    const invRefreshBtn  = document.getElementById('invRefreshBtn');
    const invExportBtn   = document.getElementById('invExportBtn');
    const invEmptyState  = document.getElementById('invEmptyState');
    const invGridEl      = document.getElementById('invGrid');
    const invDatePicker  = document.getElementById('invDatePicker');
    const invDateToday   = document.getElementById('invDateToday');
    const invDateYesterday = document.getElementById('invDateYesterday');
    const invDateHint    = document.getElementById('invDateHint');

    let inventoryRows = []; // computed inventory data

    /* ── Date picker helpers ── */
    function getTodayStr() {
        const d = new Date();
        return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }
    function getYesterdayStr() {
        const d = new Date();
        d.setDate(d.getDate() - 1);
        return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }
    function formatDateLabel(dateStr) {
        if (!dateStr) return 'all time';
        const today = getTodayStr();
        const yesterday = getYesterdayStr();
        if (dateStr === today) return 'today';
        if (dateStr === yesterday) return 'yesterday';
        // Format as DD MMM YYYY
        const d = new Date(dateStr + 'T00:00:00');
        return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
    }
    function updateDateHint() {
        if (invDateHint) {
            const val = invDatePicker ? invDatePicker.value : '';
            invDateHint.textContent = val ? `Showing stock as of ${formatDateLabel(val)}` : 'Showing stock (all time)';
        }
    }

    // Initialize date picker to today
    if (invDatePicker) {
        invDatePicker.value = getTodayStr();
        invDatePicker.addEventListener('change', () => {
            updateDateHint();
            computeInventory();
        });
    }
    if (invDateToday) {
        invDateToday.addEventListener('click', () => {
            if (invDatePicker) invDatePicker.value = getTodayStr();
            updateDateHint();
            computeInventory();
        });
    }
    if (invDateYesterday) {
        invDateYesterday.addEventListener('click', () => {
            if (invDatePicker) invDatePicker.value = getYesterdayStr();
            updateDateHint();
            computeInventory();
        });
    }
    updateDateHint();

    /* ── Populate Inventory size filter from ALL_SIZES ── */
    if (invSizeFilter) {
        invSizeFilter.innerHTML = '<option value="">All Sizes</option>' +
            ALL_SIZES.filter(s => s.value).map(s =>
                `<option value="${s.value}">${s.label}</option>`
            ).join('');
    }

    /* ── Size helpers (use ALL_SIZES master) ── */
    function getSizeBpc(sizeVal) {
        const entry = SIZE_LOOKUP.get(sizeVal);
        return entry ? entry.bpc : 0;
    }
    function getSizeLabel(sizeVal) {
        return sizeVal || '';  // label IS the value now
    }
    const LOW_STOCK_THRESHOLD = 12; // bottles

    /* ──────────────────────────────────────────────
       CORE: Build inventory from Opening Stock + TPs
       Future: subtract Sales when Sales page exists
       ────────────────────────────────────────────── */
    async function computeInventory() {
        if (!activeBar.barName) return;
        const params = { barName: activeBar.barName, financialYear: activeBar.financialYear || '' };
        const asOfDate = (invDatePicker && invDatePicker.value) ? invDatePicker.value : ''; // YYYY-MM-DD or ''

        // Fetch all three data sources in parallel
        const [osResult, tpResult, salesData] = await Promise.all([
            window.electronAPI.getOpeningStock(params),
            window.electronAPI.getTps(params),
            getSalesData(params),   // stub — returns [] until Sales page is built
        ]);

        // Map: key = "brandName|size" → aggregated row
        const stockMap = new Map();

        function getKey(brand, size) {
            return (brand || '').trim().toUpperCase() + '|' + (size || '');
        }

        function ensureRow(key, brand, code, category, size, rate) {
            if (!stockMap.has(key)) {
                stockMap.set(key, {
                    brandName: brand || '',
                    code: code || '',
                    category: category || '',
                    size: size || '',
                    sizeLabel: getSizeLabel(size),
                    bpc: getSizeBpc(size) || 0,
                    openingBtl: 0,
                    purchasedBtl: 0,      // cumulative purchases up to asOfDate
                    soldBtl: 0,           // cumulative sales up to asOfDate
                    dayPurchasedBtl: 0,   // purchases ONLY on asOfDate
                    daySoldBtl: 0,        // sales ONLY on asOfDate
                    closingBtl: 0,
                    closingCases: 0,
                    closingLoose: 0,
                    rate: rate || 0,
                    value: 0,
                });
            }
            return stockMap.get(key);
        }

        /* 1 ── Opening Stock entries (FY start) ── */
        if (osResult.success && osResult.data && osResult.data.entries) {
            osResult.data.entries.forEach(e => {
                const key = getKey(e.brandName, e.size);
                const row = ensureRow(key, e.brandName, e.code, e.category, e.size, e.rate);
                row.openingBtl += (e.totalBtl || 0);
                if (e.code && !row.code) row.code = e.code;
                if (e.category && !row.category) row.category = e.category;
                if (e.rate && e.rate > row.rate) row.rate = e.rate;
            });
        }

        /* 2 ── TP Purchases ── */
        if (tpResult.success && tpResult.tps) {
            tpResult.tps.forEach(tp => {
                if (asOfDate && tp.tpDate && tp.tpDate > asOfDate) return; // skip future TPs
                (tp.items || []).forEach(item => {
                    const key = getKey(item.brand, item.size);
                    const row = ensureRow(key, item.brand, item.code, '', item.size, item.rate);
                    const qty = item.totalBtl || 0;
                    if (asOfDate) {
                        // Date selected: TPs before the date boost opening; TPs on the date are day purchases
                        if (!tp.tpDate || tp.tpDate < asOfDate) {
                            row.openingBtl += qty;
                        } else {
                            row.dayPurchasedBtl += qty;
                        }
                    } else {
                        // Full FY view: all TPs go into cumulative purchasedBtl
                        row.purchasedBtl += qty;
                    }
                    if (item.code && !row.code) row.code = item.code;
                    if (item.rate && item.rate > row.rate) row.rate = item.rate;
                });
            });
        }

        /* 3 ── Sales deductions ── */
        const filteredSales = asOfDate
            ? (salesData || []).filter(s => !s.billDate || s.billDate <= asOfDate)
            : (salesData || []);

        if (filteredSales && filteredSales.length > 0) {
            const codeToKey = new Map();
            stockMap.forEach((row, key) => {
                if (row.code) codeToKey.set(row.code.toUpperCase(), key);
            });

            filteredSales.forEach(sale => {
                let row = null;
                if (sale.code) {
                    const existingKey = codeToKey.get(sale.code.toUpperCase());
                    if (existingKey) row = stockMap.get(existingKey);
                }
                if (!row) {
                    const key = getKey(sale.brandName, sale.size);
                    row = ensureRow(key, sale.brandName, sale.code || '', sale.category || '', sale.size, 0);
                }
                const qty = sale.totalBtl || 0;
                if (asOfDate) {
                    // Date selected: sales before the date reduce opening; sales on the date are day sales
                    if (!sale.billDate || sale.billDate < asOfDate) {
                        row.openingBtl -= qty;
                    } else {
                        row.daySoldBtl += qty;
                    }
                } else {
                    // Full FY view: all sales go into cumulative soldBtl
                    row.soldBtl += qty;
                }
            });
        }

        /* 4 ── Compute closing stock for each product ── */
        stockMap.forEach(row => {
            if (asOfDate) {
                // Date mode: opening already includes pre-date TPs/sales; add only day purchase/sale
                row.closingBtl = row.openingBtl + row.dayPurchasedBtl - row.daySoldBtl;
            } else {
                // Full FY mode: use cumulative purchase/sale; alias day fields for uniform table display
                row.dayPurchasedBtl = row.purchasedBtl;
                row.daySoldBtl = row.soldBtl;
                row.closingBtl = row.openingBtl + row.purchasedBtl - row.soldBtl;
            }
            if (row.closingBtl < 0) row.closingBtl = 0; // safety
            const bpc = row.bpc || 1;
            row.closingCases = Math.floor(row.closingBtl / bpc);
            row.closingLoose = row.closingBtl % bpc;
            row.value = row.closingBtl * row.rate;
        });

        // Enrich category (and code) from allProducts with multi-level matching
        if (allProducts.length > 0) {
            // Helper: normalize brand for fuzzy compare (strip dots, extra spaces, lowercase)
            const normBrand = (s) => (s || '').trim().replace(/\./g, '').replace(/\s+/g, ' ').trim().toUpperCase();

            stockMap.forEach(row => {
                if (row.category && row.code) return; // already have both

                let match = null;

                // Level 1: Match by product code (most reliable)
                if (row.code) {
                    match = allProducts.find(p => (p.code || '').toUpperCase() === row.code.toUpperCase());
                }

                // Level 2: Exact brand name + size
                if (!match) {
                    match = allProducts.find(p =>
                        (p.brandName || '').toUpperCase() === row.brandName.toUpperCase() &&
                        (p.size || '') === row.size
                    );
                }

                // Level 3: Normalized brand name + size (strip dots/extra spaces)
                if (!match) {
                    const nBrand = normBrand(row.brandName);
                    match = allProducts.find(p =>
                        normBrand(p.brandName) === nBrand &&
                        (p.size || '') === row.size
                    );
                }

                // Level 4: Normalized brand name only (ignore size — different sizing conventions)
                if (!match) {
                    const nBrand = normBrand(row.brandName);
                    match = allProducts.find(p => normBrand(p.brandName) === nBrand);
                }

                // Level 5: Brand name contains or is contained (partial match)
                if (!match) {
                    const nBrand = normBrand(row.brandName);
                    if (nBrand.length > 5) {
                        match = allProducts.find(p => {
                            const np = normBrand(p.brandName);
                            return np.includes(nBrand) || nBrand.includes(np);
                        });
                    }
                }

                if (match) {
                    if (!row.category) row.category = match.category || '';
                    if (!row.code && match.code) row.code = match.code;
                }
            });
        }

        inventoryRows = Array.from(stockMap.values());
        // Sort by brand name, then size descending
        inventoryRows.sort((a, b) => {
            const nameComp = a.brandName.localeCompare(b.brandName, 'en', { sensitivity: 'base' });
            if (nameComp !== 0) return nameComp;
            return (parseInt(b.size) || 0) - (parseInt(a.size) || 0);
        });

        renderInventory();
    }

    /* ──────────────────────────────────
       Get sales data: flatten daily_sales
       bills into [{brandName, size, totalBtl}]
       ────────────────────────────────── */
    async function getSalesData(params) {
        try {
            // Load ALL daily sales (no date filter — we want full FY)
            const result = await window.electronAPI.getDailySales({
                barName: params.barName,
                financialYear: params.financialYear
            });
            if (!result.success || !result.sales) return [];

            // Flatten bill items — preserve billDate for date-based filtering
            const items = [];
            for (const bill of result.sales) {
                const billDate = bill.billDate || '';
                for (const item of (bill.items || [])) {
                    items.push({
                        brandName: item.brandName || '',
                        code: item.code || '',
                        size: item.size || '',
                        category: item.category || '',
                        totalBtl: item.qty || 0,
                        billDate: billDate,
                    });
                }
            }
            return items;
        } catch (err) {
            console.error('getSalesData error:', err);
            return [];
        }
    }

    /* ── Render filtered inventory rows ── */
    function renderInventory() {
        if (!invGridBody) return;
        const filtered = getFilteredInventory();

        invGridBody.innerHTML = '';

        if (filtered.length === 0) {
            if (invEmptyState) invEmptyState.classList.remove('hidden');
            if (invGridEl) invGridEl.style.display = 'none';
        } else {
            if (invEmptyState) invEmptyState.classList.add('hidden');
            if (invGridEl) invGridEl.style.display = '';
        }

        let totalOpening = 0, totalPurchased = 0, totalSold = 0, totalClosing = 0;
        let totalCases = 0, totalLoose = 0, totalValue = 0;
        let lowStockCount = 0;

        filtered.forEach((row, idx) => {
            const tr = document.createElement('tr');

            // Status determination
            let status, statusClass, statusLabel;
            if (row.closingBtl <= 0) {
                status = 'out-of-stock';
                statusClass = 'inv-status-out-of-stock';
                statusLabel = 'Out';
                lowStockCount++;
            } else if (row.closingBtl <= LOW_STOCK_THRESHOLD) {
                status = 'low-stock';
                statusClass = 'inv-status-low-stock';
                statusLabel = 'Low';
                lowStockCount++;
            } else {
                status = 'in-stock';
                statusClass = 'inv-status-in-stock';
                statusLabel = 'OK';
            }

            tr.innerHTML = `
                <td class="gc-sr">${idx + 1}</td>
                <td class="gc-brand"><span class="inv-brand-link" data-rowidx="${idx}" title="View stock history">${esc(row.brandName)}</span></td>
                <td class="gc-code">${esc(row.code || '—')}</td>
                <td class="gc-cat">${esc(row.category || '—')}</td>
                <td class="gc-size">${esc(row.sizeLabel || '—')}</td>
                <td class="gc-num">${row.openingBtl}</td>
                <td class="gc-num">${row.dayPurchasedBtl || ''}</td>
                <td class="gc-num inv-col-sold">${row.daySoldBtl || ''}</td>
                <td class="gc-num inv-col-closing">${row.closingBtl}</td>
                <td class="gc-num">${row.closingCases}</td>
                <td class="gc-num">${row.closingLoose}</td>
                <td class="gc-rate">${row.rate > 0 ? row.rate.toFixed(2) : '—'}</td>
                <td class="gc-amount">${row.value > 0 ? '₹ ' + row.value.toFixed(2) : '—'}</td>
                <td class="gc-status"><span class="inv-status-badge ${statusClass}">${statusLabel}</span></td>
            `;
            invGridBody.appendChild(tr);

            totalOpening   += row.openingBtl;
            totalPurchased += row.dayPurchasedBtl;
            totalSold      += row.daySoldBtl;
            totalClosing   += row.closingBtl;
            totalCases     += row.closingCases;
            totalLoose     += row.closingLoose;
            totalValue     += row.value;
        });

        // Update footer totals
        setText('invFootOpening', totalOpening);
        setText('invFootPurchased', totalPurchased);
        setText('invFootSold', totalSold);
        setText('invFootClosing', totalClosing);
        setText('invFootCases', totalCases);
        setText('invFootLoose', totalLoose);
        setText('invFootValue', '₹ ' + totalValue.toFixed(2));

        // Update summary cards
        setText('invSumProducts', filtered.length);
        setText('invSumCases', totalCases);
        setText('invSumBottles', totalClosing);
        setText('invSumValue', '₹ ' + totalValue.toFixed(2));
        setText('invSumLowStock', lowStockCount);

        // Toggle warning color on low-stock card
        const lowCard = document.getElementById('invLowStockCard');
        if (lowCard) lowCard.classList.toggle('inv-sum-warning', lowStockCount > 0);
    }

    function setText(id, text) {
        const el = document.getElementById(id);
        if (el) el.textContent = text;
    }

    /* ── Stock History Modal (inventory brand click) ── */
    if (invGridBody) {
        invGridBody.addEventListener('click', e => {
            const link = e.target.closest('.inv-brand-link');
            if (!link) return;
            const idx = parseInt(link.dataset.rowidx, 10);
            const filtered = getFilteredInventory();
            const row = filtered[idx];
            if (row) showStockHistoryModal(row);
        });
    }

    async function showStockHistoryModal(row) {
        document.getElementById('sthModal')?.remove();

        // ── Spinner ──
        const loadingModal = document.createElement('div');
        loadingModal.id = 'sthModal';
        loadingModal.className = 'sth-backdrop';
        loadingModal.innerHTML = `<div class="sth-dialog" style="align-items:center;justify-content:center;min-height:220px">
            <div style="display:flex;flex-direction:column;align-items:center;gap:14px;padding:40px">
                <div class="sth-loader"></div>
                <div style="font-size:13px;color:var(--text-muted)">Loading stock history…</div>
            </div>
        </div>`;
        document.body.appendChild(loadingModal);
        loadingModal.addEventListener('click', e => { if (e.target === loadingModal) loadingModal.remove(); });

        // ── Matching helpers (mirror computeInventory logic) ──
        const matchCode  = (row.code  || '').toUpperCase();
        const matchBrand = (row.brandName || '').trim().toUpperCase();
        const matchSize  = row.size || '';
        const normB = s => (s || '').trim().replace(/\./g, '').replace(/\s+/g, ' ').trim().toUpperCase();
        const normMatchBrand = normB(row.brandName);

        function productMatch(code, brand, size) {
            if (matchCode && code && code.toUpperCase() === matchCode) return true;
            const b = (brand || '').trim().toUpperCase();
            if (b === matchBrand && (size || '') === matchSize) return true;
            if (normB(brand) === normMatchBrand && (size || '') === matchSize) return true;
            return false;
        }

        // ── Fetch all three sources in parallel ──
        const params = { barName: activeBar.barName, financialYear: activeBar.financialYear || '' };
        let osDate = '', osEntries = [], tpList = [], salesBills = [];

        try {
            const [osRes, tpRes, srRes] = await Promise.all([
                window.electronAPI.getOpeningStock(params),
                window.electronAPI.getTps(params),
                window.electronAPI.getDailySales(params),
            ]);

            // 1. Opening Stock  (osRes.data = { asOfDate, entries[] })
            if (osRes.success && osRes.data) {
                osDate = osRes.data.asOfDate || '';
                (osRes.data.entries || []).forEach(e => {
                    if (productMatch(e.code, e.brandName, e.size)) {
                        osEntries.push({
                            date    : osDate,
                            ref     : 'Opening Stock',
                            btlIn   : e.totalBtl || 0,
                            btlOut  : 0,
                            detail  : (e.cases || 0) + ' cases + ' + (e.loose || 0) + ' loose @ ₹' + (e.rate || 0),
                            isOpening: true,
                        });
                    }
                });
            }

            // 2. TP Purchases  (items use field "brand", date = tpDate)
            if (tpRes.success && tpRes.tps) {
                tpList = tpRes.tps;
            }

            // 3. Daily Sales  (items use "brandName" + "qty")
            if (srRes.success && srRes.sales) {
                salesBills = srRes.sales;
            }
        } catch (err) {
            console.error('showStockHistoryModal fetch error:', err);
        }

        // ── Build ledger events ──
        const events = [];

        // Opening entries
        events.push(...osEntries);

        // TP purchases — one row per matching TP item
        for (const tp of tpList) {
            for (const item of (tp.items || [])) {
                if (productMatch(item.code, item.brand || item.brandName, item.size)) {
                    const parts = [];
                    if ((item.cases || 0) > 0) parts.push(item.cases + ' cases');
                    if ((item.loose || 0) > 0) parts.push(item.loose + ' loose');
                    if (tp.supplier) parts.push(tp.supplier);
                    events.push({
                        date  : tp.tpDate || tp.receivedDate || '',
                        ref   : 'Purchase  TP: ' + (tp.tpNumber || tp.id || ''),
                        btlIn : item.totalBtl || 0,
                        btlOut: 0,
                        detail: parts.join(' · '),
                    });
                }
            }
        }

        // Sales — group by billDate (sum qty per day to keep ledger clean)
        const salesByDate = new Map();
        for (const bill of salesBills) {
            const dt = bill.billDate || '';
            for (const item of (bill.items || [])) {
                if (productMatch(item.code, item.brandName, item.size)) {
                    if (!salesByDate.has(dt)) salesByDate.set(dt, { qty: 0, bills: [] });
                    const entry = salesByDate.get(dt);
                    entry.qty += item.qty || 0;
                    const bNo = bill.billNo || bill.id || '';
                    if (bNo && !entry.bills.includes(bNo)) entry.bills.push(bNo);
                }
            }
        }
        salesByDate.forEach((val, dt) => {
            events.push({
                date  : dt,
                ref   : 'Sales',
                btlIn : 0,
                btlOut: val.qty,
                detail: val.bills.length ? 'Bills: ' + val.bills.join(', ') : '',
            });
        });

        // ── Sort: opening first, then chronological ──
        events.sort((a, b) => {
            if (a.isOpening && !b.isOpening) return -1;
            if (!a.isOpening && b.isOpening)  return  1;
            if (!a.date && !b.date) return 0;
            if (!a.date) return -1;
            if (!b.date) return  1;
            return a.date < b.date ? -1 : a.date > b.date ? 1 : 0;
        });

        // ── Running balance ──
        let bal = 0;
        const ledger = events.map(ev => {
            bal += ev.btlIn - ev.btlOut;
            return Object.assign({}, ev, { balance: Math.max(0, bal) });
        });

        // ── Strip totals (computed from live data) ──
        const openBtl  = osEntries.reduce((s, e) => s + e.btlIn, 0);
        const totalIn  = events.reduce((s, e) => s + e.btlIn, 0);
        const totalOut = events.reduce((s, e) => s + e.btlOut, 0);
        const purBtl   = totalIn - openBtl;
        const finalBal = ledger.length ? ledger[ledger.length - 1].balance : 0;

        // ── Helpers ──
        function fd(d) {
            if (!d) return '—';
            const dt = new Date(d);
            if (isNaN(dt)) return d;
            return dt.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
        }
        function escM(s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

        // ── Table rows ──
        const tableRows = ledger.map((ev, i) => {
            const inBadge  = ev.btlIn  > 0 ? '<span class="sth-badge sth-badge-in">+' + ev.btlIn + ' btl</span>' : '<span class="sth-cell-dash">—</span>';
            const outBadge = ev.btlOut > 0 ? '<span class="sth-badge sth-badge-out">−' + ev.btlOut + ' btl</span>' : '<span class="sth-cell-dash">—</span>';
            const balClass = ev.balance <= 0 ? 'sth-bal-zero' : ev.balance <= 12 ? 'sth-bal-low' : '';
            const rowCls   = ev.isOpening ? 'sth-row-opening' : ev.btlIn > 0 ? 'sth-row-in' : 'sth-row-out';
            const sub      = ev.detail ? '<div class="sth-ref-sub">' + escM(ev.detail) + '</div>' : '';
            return '<tr class="' + rowCls + '">' +
                '<td class="sth-td-sr">' + (i + 1) + '</td>' +
                '<td class="sth-td-date">' + fd(ev.date) + '</td>' +
                '<td class="sth-td-ref">' + escM(ev.ref) + sub + '</td>' +
                '<td class="sth-td-in">' + inBadge + '</td>' +
                '<td class="sth-td-out">' + outBadge + '</td>' +
                '<td class="sth-td-bal ' + balClass + '">' + ev.balance + '</td>' +
                '</tr>';
        }).join('');

        // ── Replace spinner with real modal ──
        document.getElementById('sthModal')?.remove();

        const modal = document.createElement('div');
        modal.id = 'sthModal';
        modal.className = 'sth-backdrop';
        const emptyMsg = '<div class="sth-empty">No stock movements found for this product.</div>';
        const tableHtml = ledger.length === 0 ? emptyMsg :
            '<table class="sth-table"><thead><tr>' +
            '<th class="sth-th-sr">#</th>' +
            '<th class="sth-th-date">Date</th>' +
            '<th class="sth-th-ref">Event / Reference</th>' +
            '<th class="sth-th-in">Stock In</th>' +
            '<th class="sth-th-out">Stock Out</th>' +
            '<th class="sth-th-bal">Balance</th>' +
            '</tr></thead><tbody>' + tableRows + '</tbody></table>';

        modal.innerHTML =
            '<div class="sth-dialog" role="dialog" aria-modal="true">' +
              '<div class="sth-header">' +
                '<div class="sth-header-left">' +
                  '<div class="sth-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg></div>' +
                  '<div>' +
                    '<div class="sth-title">' + escM(row.brandName) + '</div>' +
                    '<div class="sth-sub">' + escM(row.sizeLabel || row.size || '') + ' &nbsp;·&nbsp; Full Year Stock Ledger &nbsp;·&nbsp; FY ' + escM(activeBar.financialYear || '') + '</div>' +
                  '</div>' +
                '</div>' +
                '<button class="sth-close" id="sthClose" title="Close"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>' +
              '</div>' +
              '<div class="sth-summary-strip">' +
                '<div class="sth-strip-item"><span class="sth-strip-label">Opening</span><span class="sth-strip-val">' + openBtl + ' btl</span></div>' +
                '<div class="sth-strip-item"><span class="sth-strip-label">Purchased</span><span class="sth-strip-val sth-val-in">+' + purBtl + ' btl</span></div>' +
                '<div class="sth-strip-item"><span class="sth-strip-label">Sold</span><span class="sth-strip-val sth-val-out">−' + totalOut + ' btl</span></div>' +
                '<div class="sth-strip-item"><span class="sth-strip-label">Closing</span><span class="sth-strip-val">' + finalBal + ' btl</span></div>' +
              '</div>' +
              '<div class="sth-body">' + tableHtml + '</div>' +
              '<div class="sth-footer">' +
                '<span class="sth-footer-note">' + ledger.length + ' movement' + (ledger.length !== 1 ? 's' : '') + ' &nbsp;·&nbsp; Closing: <strong>' + finalBal + ' btl (' + row.closingCases + ' cases + ' + row.closingLoose + ' loose)</strong></span>' +
                '<button class="sth-close-btn" id="sthCloseBtn">Close</button>' +
              '</div>' +
            '</div>';

        document.body.appendChild(modal);
        const close = () => modal.remove();
        document.getElementById('sthClose')?.addEventListener('click', close);
        document.getElementById('sthCloseBtn')?.addEventListener('click', close);
        modal.addEventListener('click', e => { if (e.target === modal) close(); });
        document.addEventListener('keydown', function escKey(e) {
            if (e.key === 'Escape') { close(); document.removeEventListener('keydown', escKey); }
        });
    }
    /* ── Filter logic ── */
    function getFilteredInventory() {
        let rows = inventoryRows;
        // Beer Shopee (FL-BR-II) cannot carry Spirits
        if (isBeerShopee) {
            rows = rows.filter(r => (r.category || '').toUpperCase() !== 'SPIRITS');
        }
        const searchQ = (invSearchInput?.value || '').trim().toLowerCase();
        const catQ    = invCatFilter?.value || '';
        const sizeQ   = invSizeFilter?.value || '';
        const stockQ  = invStockFilter?.value || '';

        if (searchQ) {
            rows = rows.filter(r =>
                (r.brandName || '').toLowerCase().includes(searchQ) ||
                (r.code || '').toLowerCase().includes(searchQ) ||
                (r.category || '').toLowerCase().includes(searchQ)
            );
        }
        if (catQ) {
            rows = rows.filter(r => (r.category || '').toUpperCase() === catQ.toUpperCase());
        }
        if (sizeQ) {
            rows = rows.filter(r => r.size === sizeQ);
        }
        if (stockQ === 'in-stock') {
            rows = rows.filter(r => r.closingBtl > LOW_STOCK_THRESHOLD);
        } else if (stockQ === 'low-stock') {
            rows = rows.filter(r => r.closingBtl > 0 && r.closingBtl <= LOW_STOCK_THRESHOLD);
        } else if (stockQ === 'out-of-stock') {
            rows = rows.filter(r => r.closingBtl <= 0);
        }

        return rows;
    }

    /* ── Wire filter events ── */
    [invSearchInput, invCatFilter, invSizeFilter, invStockFilter].forEach(el => {
        if (!el) return;
        const evt = el.tagName === 'SELECT' ? 'change' : 'input';
        el.addEventListener(evt, () => renderInventory());
    });

    /* ── Clear filters on Escape ── */
    function clearInvFilters() {
        if (invSearchInput) invSearchInput.value = '';
        if (invCatFilter) invCatFilter.value = '';
        if (invSizeFilter) invSizeFilter.value = '';
        if (invStockFilter) invStockFilter.value = '';
        if (invDatePicker) invDatePicker.value = getTodayStr();
        updateDateHint();
        computeInventory();
    }

    /* ── Export as CSV ── */
    function exportInventoryCSV() {
        if (inventoryRows.length === 0) {
            showTpToast('No inventory data to export', true);
            return;
        }
        const filtered = getFilteredInventory();
        const headers = ['#', 'Brand Name', 'Code', 'Category', 'Size', 'Opening (Btl)', 'Purchased (Btl)', 'Sold (Btl)', 'Closing (Btl)', 'Closing Cases', 'Loose', 'Rate', 'Value'];
        const csvRows = [headers.join(',')];
        filtered.forEach((row, i) => {
            csvRows.push([
                i + 1,
                '"' + (row.brandName || '').replace(/"/g, '""') + '"',
                '"' + (row.code || '') + '"',
                '"' + (row.category || '') + '"',
                row.sizeLabel || '',
                row.openingBtl,
                row.purchasedBtl,
                row.soldBtl,
                row.closingBtl,
                row.closingCases,
                row.closingLoose,
                row.rate.toFixed(2),
                row.value.toFixed(2),
            ].join(','));
        });

        const blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `inventory_${activeBar.barName || 'export'}_asof_${(invDatePicker && invDatePicker.value) || new Date().toISOString().slice(0, 10)}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showTpToast(`Exported ${filtered.length} items to CSV`);
    }

    /* ── Wire buttons ── */
    if (invRefreshBtn) invRefreshBtn.addEventListener('click', () => computeInventory());
    if (invExportBtn) invExportBtn.addEventListener('click', () => exportInventoryCSV());

    /* ── Keyboard shortcuts for Inventory panel ── */
    document.addEventListener('keydown', (e) => {
        const invPanel = document.getElementById('view-inventory');
        if (!invPanel || !invPanel.classList.contains('active')) return;

        if (e.key === 'F5') {
            e.preventDefault();
            computeInventory();
            return;
        }
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
            e.preventDefault();
            if (invSearchInput) invSearchInput.focus();
            return;
        }
        if (e.key === 'Escape') {
            e.preventDefault();
            clearInvFilters();
            return;
        }
    });

    /* ── Auto-refresh when Inventory tab is activated ── */
    const invNavBtn = document.getElementById('nav-inventory');
    if (invNavBtn) {
        invNavBtn.addEventListener('click', () => {
            // Small delay to let switchView complete and panel become active
            setTimeout(() => computeInventory(), 50);
        });
    }
    // Also observe if panel becomes active by other means (keyboard nav, etc.)
    const invPanel = document.getElementById('view-inventory');
    if (invPanel) {
        const invObserver = new MutationObserver((mutations) => {
            mutations.forEach(m => {
                if (m.attributeName === 'class' && invPanel.classList.contains('active')) {
                    computeInventory();
                }
            });
        });
        invObserver.observe(invPanel, { attributes: true, attributeFilter: ['class'] });
    }

    // ═══════════════════════════════════════════════════════
    //  BRANDWISE REPORT — Maharashtra FL3 Stock Register
    //  Pivoted: One row per brand, sizes as sub-columns
    // ═══════════════════════════════════════════════════════
    {
        const bwSingleDate     = document.getElementById('bwSingleDate');
        const bwDateFrom       = document.getElementById('bwDateFrom');
        const bwDateTo         = document.getElementById('bwDateTo');
        const bwGoBtn          = document.getElementById('bwGoBtn');
        const bwCategoryFilter = document.getElementById('bwCategoryFilter');
        const bwSearch         = document.getElementById('bwSearch');
        const bwExportBtn      = document.getElementById('bwExportBtn');
        const bwPrintBtn       = document.getElementById('bwPrintBtn');
        const bwFyBadge        = document.getElementById('bwFyBadge');
        const bwTableHead      = document.getElementById('bwTableHead');
        const bwTableBody      = document.getElementById('bwTableBody');
        const bwTableFoot      = document.getElementById('bwTableFoot');
        const bwMmlToggle    = document.getElementById('bwMmlToggle');
        const bwEmpty          = document.getElementById('bwEmpty');

        // Summary card elements
        const bwTotalBrands  = document.getElementById('bwTotalBrands');
        const bwTotalOpening = document.getElementById('bwTotalOpening');
        const bwTotalReceipt = document.getElementById('bwTotalReceipt');
        const bwTotalSale    = document.getElementById('bwTotalSale');
        const bwTotalClosing = document.getElementById('bwTotalClosing');

        let bwRows = [];      // pivoted brand rows
        let bwFiltered = [];  // after filters
        let bwSizeCols = [];  // ordered size labels detected from data e.g. ["90 ML","180 ML","375 ML","750 ML"]

        // FY defaults
        const bwFyStartYear = (() => {
            const now = new Date();
            return now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
        })();
        const bwFyStartDate = bwFyStartYear + '-04-01';
        if (bwDateFrom) bwDateFrom.value = bwFyStartDate;
        if (bwDateTo)   bwDateTo.value = new Date().toISOString().slice(0, 10);
        if (bwSingleDate) bwSingleDate.value = new Date().toISOString().slice(0, 10);
        if (bwFyBadge)  bwFyBadge.textContent = 'FY ' + bwFyStartYear + '-' + String(bwFyStartYear + 1).slice(2);

        /* ── Mode switching: 'single' | 'range' ── */
        const bwSingleGroup = document.getElementById('bwSingleGroup');
        const bwRangeGroup  = document.getElementById('bwRangeGroup');
        let bwIsRangeMode = false;

        function setBwMode(mode) {
            bwIsRangeMode = (mode === 'range');

            if (bwSingleGroup) {
                bwSingleGroup.classList.toggle('bw-single-dimmed', bwIsRangeMode);
                bwSingleGroup.classList.toggle('bw-single-group',  !bwIsRangeMode);
            }
            if (bwRangeGroup) {
                bwRangeGroup.classList.toggle('bw-range-dimmed',      !bwIsRangeMode);
                bwRangeGroup.classList.toggle('bw-range-active-group',  bwIsRangeMode);
            }

            // Enable / disable inputs
            if (bwSingleDate) bwSingleDate.disabled = bwIsRangeMode;
            if (bwDateFrom)   bwDateFrom.disabled   = !bwIsRangeMode;
            if (bwDateTo)     bwDateTo.disabled     = !bwIsRangeMode;

            // When switching to range, pre-fill From/To if empty
            if (bwIsRangeMode) {
                if (!bwDateFrom.value) bwDateFrom.value = bwFyStartDate;
                if (!bwDateTo.value)   bwDateTo.value   = new Date().toISOString().slice(0, 10);
                // Focus From field
                requestAnimationFrame(() => { if (bwDateFrom) bwDateFrom.focus(); });
            }
        }

        // Default: single date mode
        setBwMode('single');

        // Click on range group → switch to range mode
        if (bwRangeGroup) {
            bwRangeGroup.addEventListener('click', e => {
                if (!bwIsRangeMode) {
                    setBwMode('range');
                    e.stopPropagation();
                }
            });
        }

        // Click on single group → switch to single date mode
        if (bwSingleGroup) {
            bwSingleGroup.addEventListener('click', e => {
                if (bwIsRangeMode) {
                    setBwMode('single');
                    requestAnimationFrame(() => { if (bwSingleDate) bwSingleDate.focus(); });
                    e.stopPropagation();
                }
            });
        }

        /* ── Brandwise Date Picker Modal ── */
        const bwDateModal      = document.getElementById('bwDateModal');
        const bwModalDate      = document.getElementById('bwModalDate');
        const bwModalGoBtn     = document.getElementById('bwModalGoBtn');
        const bwModalCancelBtn = document.getElementById('bwModalCancelBtn');

        function showBwDateModal() {
            if (!bwDateModal || !bwModalDate) return;
            // Always open in single-date mode; pre-fill with today
            bwModalDate.value = new Date().toISOString().slice(0, 10);
            bwDateModal.classList.remove('hidden');
            requestAnimationFrame(() => bwModalDate.focus());
        }

        function hideBwDateModal() {
            if (bwDateModal) bwDateModal.classList.add('hidden');
        }

        function confirmBwDateModal() {
            if (!bwModalDate || !bwModalDate.value) return;
            // Force single-date mode and set the selected date
            setBwMode('single');
            if (bwSingleDate) bwSingleDate.value = bwModalDate.value;
            hideBwDateModal();
            generateBrandwise();
        }

        if (bwModalGoBtn)     bwModalGoBtn.addEventListener('click', confirmBwDateModal);
        if (bwModalCancelBtn) bwModalCancelBtn.addEventListener('click', hideBwDateModal);

        if (bwDateModal) {
            bwDateModal.addEventListener('keydown', e => {
                if (e.key === 'Enter')  { e.preventDefault(); confirmBwDateModal(); }
                if (e.key === 'Escape') { e.preventDefault(); hideBwDateModal(); }
            });
            bwDateModal.addEventListener('click', e => {
                if (e.target === bwDateModal) hideBwDateModal();
            });
        }

        /* ── Expose so switchReportSub can call it ── */
        window._showBwDateModal = showBwDateModal;

        /* ── Helper: extract ML number from size string for sorting ── */
        function sizeML(s) {
            const m = (s || '').match(/^(\d+)\s*ML/i);
            if (m) return parseInt(m[1]);
            const l = (s || '').match(/^(\d+)\s*Ltr/i);
            if (l) return parseInt(l[1]) * 1000;
            return 99999;
        }

        /* ── Normalize size to a group label (strip BPC/Pet variants) ── */
        function sizeGroup(s) {
            // "90 ML (96)" → "90 ML", "180 ML (Pet)" → "180 ML", "2000 ML (6)" → "2000 ML"
            const m = (s || '').match(/^(\d+\s*ML)/i);
            if (m) return m[1].replace(/\s+/g, ' ').toUpperCase().replace('ML', 'ML');
            const l = (s || '').match(/^(\d+\s*Ltr)/i);
            if (l) return l[1];
            return s || '';
        }

        const BW_CATEGORY_ORDER = ['SPIRITS', 'WINE', 'FERMENTED BEER', 'MILD BEER', 'MML'];

        function normalizeBwCategory(cat) {
            const raw = String(cat || '').trim().toUpperCase();
            if (!raw) return 'UNCATEGORIZED';
            if (raw === 'WINES') return 'WINE';
            if (raw === 'SPIRIT') return 'SPIRITS';
            return raw;
        }

        function bwCategoryDisplay(cat) {
            const n = normalizeBwCategory(cat);
            if (n === 'SPIRITS') return 'Spirits';
            if (n === 'WINE') return 'Wine';
            if (n === 'FERMENTED BEER') return 'Fermented Beer';
            if (n === 'MILD BEER') return 'Mild Beer';
            if (n === 'MML') return 'MML';
            return cat || 'Uncategorized';
        }

        function bwCategoryRank(cat) {
            const idx = BW_CATEGORY_ORDER.indexOf(normalizeBwCategory(cat));
            return idx >= 0 ? idx : 99;
        }

        function bwShortBrand(name, maxLen = 20) {
            const s = String(name || '').trim();
            if (s.length <= maxLen) return s;
            return s.slice(0, maxLen) + '...';
        }

        /* ── Main generate ── */
        async function generateBrandwise() {
            if (!activeBar.barName) return;
            const params = { barName: activeBar.barName, financialYear: activeBar.financialYear || '' };
            // Determine dates based on active mode
            let fromDate, toDate, effectiveDateStr;
            if (bwIsRangeMode) {
                fromDate         = (bwDateFrom && bwDateFrom.value) ? bwDateFrom.value : bwFyStartDate;
                toDate           = (bwDateTo   && bwDateTo.value)   ? bwDateTo.value   : new Date().toISOString().slice(0, 10);
                effectiveDateStr = toDate;
            } else {
                // Single date mode: show that date's opening/purchase/sale/closing
                // Opening = FY stock adjusted by all transactions BEFORE this date
                // Purchase/Sale = only transactions ON this date
                const singleVal  = (bwSingleDate && bwSingleDate.value) ? bwSingleDate.value : new Date().toISOString().slice(0, 10);
                fromDate         = singleVal;   // transactions before this → opening; on this date → purchase/sale
                toDate           = singleVal;
                effectiveDateStr = singleVal;
            }

            // Update AS ON DATE label
            const bwAsOnDate = document.getElementById('bwAsOnDate');
            if (bwAsOnDate) {
                const [y, m, d] = effectiveDateStr.split('-');
                bwAsOnDate.textContent = d + '/' + m + '/' + y;
            }

            const [osResult, tpResult, salesItems] = await Promise.all([
                window.electronAPI.getOpeningStock(params),
                window.electronAPI.getTps(params),
                getSalesData(params),
            ]);

            // stockMap: key = "BRAND|SIZE" → { openingBtl, purchaseBtl, saleBtl, closingBtl }
            // brandMap: brandKey → { brandName, category, subCategory, tpNumbers: Set, sizes: { sizeGroup → {op,pur,sale,cl} } }
            const brandMap = new Map();
            const allSizeGroups = new Set();

            const normBrand = (s) => (s || '').trim().replace(/\./g, '').replace(/\s+/g, ' ').trim().toUpperCase();

            function ensureBrand(brandName, category, subCat) {
                const key = normBrand(brandName);
                if (!brandMap.has(key)) {
                    brandMap.set(key, {
                        brandName: (brandName || '').trim(),
                        category: category || '',
                        subCategory: subCat || '',
                        tpNumbers: new Set(),
                        sizes: {},       // sizeGroup → { opening: 0, purchase: 0, sale: 0, closing: 0 }
                    });
                }
                const row = brandMap.get(key);
                if (category && !row.category) row.category = category;
                if (subCat && !row.subCategory) row.subCategory = subCat;
                return row;
            }

            function ensureSize(brandRow, sz) {
                const sg = sizeGroup(sz);
                allSizeGroups.add(sg);
                if (!brandRow.sizes[sg]) {
                    brandRow.sizes[sg] = { opening: 0, purchase: 0, sale: 0, closing: 0 };
                }
                return brandRow.sizes[sg];
            }

            // 1. Opening Stock (FY start — always loaded)
            if (osResult.success && osResult.data && osResult.data.entries) {
                osResult.data.entries.forEach(e => {
                    const row = ensureBrand(e.brandName, e.category, '');
                    const sz = ensureSize(row, e.size);
                    sz.opening += (e.totalBtl || 0);
                });
            }

            // 2. TP Purchases — split into "before range" (adds to opening) and "in range"
            if (tpResult.success && tpResult.tps) {
                tpResult.tps.forEach(tp => {
                    const d = tp.tpDate || '';
                    (tp.items || []).forEach(item => {
                        const row = ensureBrand(item.brand, '', '');
                        const sz = ensureSize(row, item.size);

                        if (fromDate && d && d < fromDate) {
                            // TP before the selected From date → rolls into adjusted opening
                            sz.opening += (item.totalBtl || 0);
                        } else if (toDate && d && d > toDate) {
                            // TP after the selected To date → ignore entirely
                        } else {
                            // Within range
                            sz.purchase += (item.totalBtl || 0);
                            if (tp.tpNumber) row.tpNumbers.add(tp.tpNumber);
                        }
                    });
                });
            }

            // 3. Sales — split into "before range" (subtracts from opening) and "in range"
            (salesItems || []).forEach(sale => {
                const d = sale.billDate || '';
                const row = ensureBrand(sale.brandName, sale.category, '');
                const sz = ensureSize(row, sale.size);

                if (fromDate && d && d < fromDate) {
                    // Sale before From date → reduces adjusted opening
                    sz.opening -= (sale.totalBtl || 0);
                } else if (toDate && d && d > toDate) {
                    // Sale after To date → ignore entirely
                } else {
                    // Within range
                    sz.sale += (sale.totalBtl || 0);
                }
            });

            // 4. Enrich categories from allProducts
            brandMap.forEach(row => {
                if (!row.category || !row.subCategory) {
                    const nb = normBrand(row.brandName);
                    const match = allProducts.find(p => normBrand(p.brandName) === nb);
                    if (match) {
                        if (!row.category) row.category = match.category || '';
                        if (!row.subCategory) row.subCategory = match.subCategory || '';
                    }
                }
                // Compute closing for each size
                Object.values(row.sizes).forEach(sz => {
                    if (sz.opening < 0) sz.opening = 0; // clamp adjusted opening
                    sz.closing = sz.opening + sz.purchase - sz.sale;
                    if (sz.closing < 0) sz.closing = 0;
                });
            });

            // Remove brands where all sizes have zero across all columns
            brandMap.forEach((row, key) => {
                const hasAny = Object.values(row.sizes).some(sz =>
                    sz.opening || sz.purchase || sz.sale || sz.closing
                );
                if (!hasAny) brandMap.delete(key);
            });

            // Build ordered size columns (sorted by ML)
            bwSizeCols = [...allSizeGroups].sort((a, b) => sizeML(b) - sizeML(a));

            // Build flat row array for rendering
            bwRows = Array.from(brandMap.values());
            bwRows.sort((a, b) => {
                const catOrder = bwCategoryRank(a.category) - bwCategoryRank(b.category);
                if (catOrder !== 0) return catOrder;
                const cat = bwCategoryDisplay(a.category).localeCompare(bwCategoryDisplay(b.category), 'en', { sensitivity: 'base' });
                if (cat !== 0) return cat;
                return a.brandName.localeCompare(b.brandName, 'en', { sensitivity: 'base' });
            });

            // Populate category filter
            if (bwCategoryFilter) {
                const cats = [...new Set(bwRows.map(r => r.category).filter(Boolean))]
                    .filter(c => !(isBeerShopee && normalizeBwCategory(c) === 'SPIRITS'))
                    .sort((a, b) => {
                        const ord = bwCategoryRank(a) - bwCategoryRank(b);
                        if (ord !== 0) return ord;
                        return bwCategoryDisplay(a).localeCompare(bwCategoryDisplay(b), 'en', { sensitivity: 'base' });
                    });
                const current = bwCategoryFilter.value;
                bwCategoryFilter.innerHTML = '<option value="">All Categories</option>' +
                    cats.map(c => `<option value="${c}" ${c === current ? 'selected' : ''}>${bwCategoryDisplay(c)}</option>`).join('');
            }

            applyBwFilters();
        }

        function applyBwFilters() {
            const catVal = bwCategoryFilter ? bwCategoryFilter.value : '';
            const searchVal = bwSearch ? bwSearch.value.trim().toUpperCase() : '';
            const showMml = bwMmlToggle ? bwMmlToggle.checked : true;

            // When MML is hidden, find the raw Spirits category string used in this dataset
            // so MML brands can be remapped into it (fall under Spirits section)
            const spiritsCatRaw = !showMml
                ? (bwRows.find(r => normalizeBwCategory(r.category) === 'SPIRITS')?.category || 'SPIRITS')
                : null;

            bwFiltered = bwRows
                .map(r => {
                    if (!showMml && normalizeBwCategory(r.category) === 'MML') {
                        // Remap MML brands into Spirits so they appear in that category group
                        return Object.assign({}, r, { category: spiritsCatRaw });
                    }
                    return r;
                })
                .filter(r => {
                    // Beer Shopee cannot have Spirits
                    if (isBeerShopee && normalizeBwCategory(r.category) === 'SPIRITS') return false;
                    // Use normalized comparison so remapped categories still respect the dropdown
                    if (catVal && normalizeBwCategory(r.category) !== normalizeBwCategory(catVal)) return false;
                    if (searchVal && !r.brandName.toUpperCase().includes(searchVal)) return false;
                    return true;
                });

            // Re-sort after remap (MML brands now under Spirits need to be ordered with them)
            if (!showMml) {
                bwFiltered.sort((a, b) => {
                    const catOrder = bwCategoryRank(a.category) - bwCategoryRank(b.category);
                    if (catOrder !== 0) return catOrder;
                    return a.brandName.localeCompare(b.brandName, 'en', { sensitivity: 'base' });
                });
            }

            renderBwTable();
        }

        /* ── Build the multi-level header ── */
        function buildBwHeader() {
            if (!bwTableHead) return;
            bwTableHead.innerHTML = '';
            const tr = document.createElement('tr');
            tr.className = 'bw-header-sizes';
            let headerHtml = `
                <th class="bw-th-sr bw-th-frozen">Sr.</th>
                <th class="bw-th-brand bw-th-frozen">Brand Name</th>
                <th class="bw-th-tp bw-th-frozen">TP No.</th>
            `;
            tr.innerHTML = headerHtml;
            bwTableHead.appendChild(tr);
        }

        /* ── Render table body ── */
        function renderBwTable() {
            if (!bwTableBody) return;

            buildBwHeader();
            bwTableBody.innerHTML = '';
            if (bwTableFoot) bwTableFoot.innerHTML = '';

            if (bwFiltered.length === 0 || bwSizeCols.length === 0) {
                if (bwEmpty) bwEmpty.classList.remove('hidden');
                document.getElementById('bwTable')?.classList.add('hidden');
                updateBwSummary([]);
                return;
            }
            if (bwEmpty) bwEmpty.classList.add('hidden');
            document.getElementById('bwTable')?.classList.remove('hidden');

            let catSr = 0;
            let lastCat = '';
            const fmtQtyCell = (n) => {
                const num = Number(n) || 0;
                return num === 0 ? '-' : num.toLocaleString('en-IN');
            };

            const categorySizesMap = new Map();
            bwFiltered.forEach(row => {
                const catKey = bwCategoryDisplay(row.category);
                if (!categorySizesMap.has(catKey)) categorySizesMap.set(catKey, new Set());
                const sizeSet = categorySizesMap.get(catKey);
                bwSizeCols.forEach(sz => {
                    const s = row.sizes[sz];
                    if (s && (s.opening || s.purchase || s.sale || s.closing)) {
                        sizeSet.add(sz);
                    }
                });
            });

            function appendCategorySubHeaders(sizeCols) {
                const nSizes = sizeCols.length;
                const groupBoundaryClass = (si) => `${si === 0 ? ' bw-group-start' : ''}${si === nSizes - 1 ? ' bw-group-end' : ''}`;

                const groupTr = document.createElement('tr');
                groupTr.className = 'bw-cat-subhead-group';
                groupTr.innerHTML = `
                    <td colspan="3" class="bw-cat-subhead-stub">Particulars</td>
                    <td colspan="${nSizes}" class="bw-cat-subhead-metric">OPENING</td>
                    <td colspan="${nSizes}" class="bw-cat-subhead-metric">PURCHASE</td>
                    <td colspan="${nSizes}" class="bw-cat-subhead-metric">SALES</td>
                    <td colspan="${nSizes}" class="bw-cat-subhead-metric">CLOSING</td>
                `;
                bwTableBody.appendChild(groupTr);

                const sizesTr = document.createElement('tr');
                sizesTr.className = 'bw-cat-subhead-sizes';
                let cells = `<td colspan="3" class="bw-cat-subhead-stub">Sizes</td>`;
                ['opening', 'purchase', 'sales', 'closing'].forEach(() => {
                    sizeCols.forEach(sz => {
                        const shortLabel = sz.replace(/\s*ML/i, '').trim();
                        const si = sizeCols.indexOf(sz);
                        cells += `<td class="bw-cat-subhead-size${groupBoundaryClass(si)}">${esc(shortLabel)}</td>`;
                    });
                });
                sizesTr.innerHTML = cells;
                bwTableBody.appendChild(sizesTr);
            }

            let currentCatSizes = [];
            let catOpen = [];
            let catPur = [];
            let catSale = [];
            let catCl = [];

            function pushCategoryTotalRow(catLabel) {
                if (!catLabel) return;
                const nSizes = currentCatSizes.length;
                const groupBoundaryClass = (si) => `${si === 0 ? ' bw-group-start' : ''}${si === nSizes - 1 ? ' bw-group-end' : ''}`;
                const totalTr = document.createElement('tr');
                totalTr.className = 'bw-cat-total-row';
                let totalCells = `<td colspan="3" class="bw-cat-total-label">${esc(catLabel)} Total</td>`;
                catOpen.forEach((v, si) => { totalCells += `<td class="bw-cat-total-num${groupBoundaryClass(si)}">${fmtQtyCell(v)}</td>`; });
                catPur.forEach((v, si)  => { totalCells += `<td class="bw-cat-total-num${groupBoundaryClass(si)}">${fmtQtyCell(v)}</td>`; });
                catSale.forEach((v, si) => { totalCells += `<td class="bw-cat-total-num${groupBoundaryClass(si)}">${fmtQtyCell(v)}</td>`; });
                catCl.forEach((v, si)   => { totalCells += `<td class="bw-cat-total-num${groupBoundaryClass(si)}">${fmtQtyCell(v)}</td>`; });
                totalTr.innerHTML = totalCells;
                bwTableBody.appendChild(totalTr);
            }

            bwFiltered.forEach(row => {
                // Category header
                const catKey = bwCategoryDisplay(row.category);
                if (catKey !== lastCat) {
                    if (lastCat) pushCategoryTotalRow(lastCat);
                    const sizeSet = categorySizesMap.get(catKey) || new Set();
                    currentCatSizes = bwSizeCols.filter(sz => sizeSet.has(sz));
                    catOpen = new Array(currentCatSizes.length).fill(0);
                    catPur = new Array(currentCatSizes.length).fill(0);
                    catSale = new Array(currentCatSizes.length).fill(0);
                    catCl = new Array(currentCatSizes.length).fill(0);
                    catSr = 0;
                    lastCat = catKey;
                    const catTr = document.createElement('tr');
                    catTr.className = 'bw-cat-header';
                    const totalCols = 3 + currentCatSizes.length * 4;
                    catTr.innerHTML = `<td colspan="${totalCols}"><span class="bw-cat-title">${catKey}</span></td>`;
                    bwTableBody.appendChild(catTr);
                    appendCategorySubHeaders(currentCatSizes);
                }

                catSr++;
                const tr = document.createElement('tr');
                const tpShort = [...row.tpNumbers].map(t => { const i = t.lastIndexOf('/'); return i >= 0 ? t.slice(i + 1) : t; }).join(', ') || '—';
                const tpFull = [...row.tpNumbers].join(', ') || '—';
                const brandShort = bwShortBrand(row.brandName, 20);

                let cells = `
                    <td class="bw-td-sr">${catSr}</td>
                    <td class="bw-td-brand" title="${esc(row.brandName)}">${esc(brandShort)}</td>
                    <td class="bw-td-tp" title="${esc(tpFull)}">${esc(tpShort)}</td>
                `;
                const nSizes = currentCatSizes.length;
                const groupBoundaryClass = (si) => `${si === 0 ? ' bw-group-start' : ''}${si === nSizes - 1 ? ' bw-group-end' : ''}`;

                // Opening sizes
                currentCatSizes.forEach((sz, si) => {
                    const v = (row.sizes[sz] ? row.sizes[sz].opening : 0);
                    catOpen[si] += v;
                    cells += `<td class="bw-td-num ${v ? '' : 'bw-td-zero'}${groupBoundaryClass(si)}">${fmtQtyCell(v)}</td>`;
                });
                // Purchase sizes
                currentCatSizes.forEach((sz, si) => {
                    const v = (row.sizes[sz] ? row.sizes[sz].purchase : 0);
                    catPur[si] += v;
                    cells += `<td class="bw-td-num ${v ? '' : 'bw-td-zero'}${groupBoundaryClass(si)}">${fmtQtyCell(v)}</td>`;
                });
                // Sales sizes
                currentCatSizes.forEach((sz, si) => {
                    const v = (row.sizes[sz] ? row.sizes[sz].sale : 0);
                    catSale[si] += v;
                    cells += `<td class="bw-td-num ${v ? '' : 'bw-td-zero'}${groupBoundaryClass(si)}">${fmtQtyCell(v)}</td>`;
                });
                // Closing sizes
                currentCatSizes.forEach((sz, si) => {
                    const v = (row.sizes[sz] ? row.sizes[sz].closing : 0);
                    catCl[si] += v;
                    const cls = v < 0 ? 'bw-td-neg' : (v ? '' : 'bw-td-zero');
                    cells += `<td class="bw-td-num bw-td-closing ${cls}${groupBoundaryClass(si)}">${fmtQtyCell(v)}</td>`;
                });

                tr.innerHTML = cells;
                bwTableBody.appendChild(tr);
            });

            pushCategoryTotalRow(lastCat);

            updateBwSummary(bwFiltered);
        }

        function updateBwSummary(rows) {
            let totalOpening = 0, totalPurchase = 0, totalSale = 0, totalClosing = 0;
            rows.forEach(r => {
                Object.values(r.sizes).forEach(sz => {
                    totalOpening  += sz.opening;
                    totalPurchase += sz.purchase;
                    totalSale     += sz.sale;
                    totalClosing  += sz.closing;
                });
            });
            const fmt = (n) => n.toLocaleString('en-IN');
            if (bwTotalBrands)  bwTotalBrands.textContent = rows.length;
            if (bwTotalOpening) bwTotalOpening.textContent = fmt(totalOpening);
            if (bwTotalReceipt) bwTotalReceipt.textContent = fmt(totalPurchase);
            if (bwTotalSale)    bwTotalSale.textContent = fmt(totalSale);
            if (bwTotalClosing) bwTotalClosing.textContent = fmt(totalClosing);
        }

        // Helper: html-escape
        function esc(str) {
            return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        }

        // Export CSV
        function exportBwCSV() {
            if (bwFiltered.length === 0) return;
            const fromStr = bwIsRangeMode ? (bwDateFrom ? bwDateFrom.value : bwFyStartDate) : bwFyStartDate;
            const toStr   = bwIsRangeMode ? (bwDateTo   ? bwDateTo.value   : '') : (bwSingleDate ? bwSingleDate.value : '');

            // Build header rows
            let hdr1 = ['Sr', 'Brand Name', 'TP No'];
            let hdr2 = ['', '', ''];
            ['Opening', 'Purchase', 'Sales', 'Closing'].forEach(g => {
                bwSizeCols.forEach((sz, i) => {
                    hdr1.push(i === 0 ? g : '');
                    hdr2.push(sz);
                });
            });

            const csvRows = [hdr1.join(','), hdr2.join(',')];
            bwFiltered.forEach((row, idx) => {
                const line = [
                    idx + 1,
                    '"' + (row.brandName || '').replace(/"/g, '""') + '"',
                    '"' + ([...row.tpNumbers].map(t => { const i = t.lastIndexOf('/'); return i >= 0 ? t.slice(i + 1) : t; }).join('; ') || '').replace(/"/g, '""') + '"',
                ];
                ['opening', 'purchase', 'sale', 'closing'].forEach(fld => {
                    bwSizeCols.forEach(sz => {
                        line.push(row.sizes[sz] ? row.sizes[sz][fld] : 0);
                    });
                });
                csvRows.push(line.join(','));
            });

            const blob = new Blob(['\uFEFF' + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `Brandwise_${activeBar.barName || 'Bar'}_${fromStr}_to_${toStr}.csv`;
            a.click();
            URL.revokeObjectURL(a.href);
        }

        // Print
        function printBwReport() {
            if (!bwFiltered.length || !bwSizeCols.length) return;
            const fromStr = bwIsRangeMode ? (bwDateFrom ? bwDateFrom.value : '') : bwFyStartDate;
            const toStr   = bwIsRangeMode ? (bwDateTo   ? bwDateTo.value   : '') : (bwSingleDate ? bwSingleDate.value : '');
            const barName = activeBar.barName || '';
            const licNo = activeBar.licenseNo || activeBar.licNo || '—';
            const fy = bwFyBadge ? bwFyBadge.textContent : '';
            const address = [activeBar.address, activeBar.area, activeBar.city, activeBar.state, activeBar.pinCode].filter(Boolean).join(', ');
            const fmtDate = (s) => { if (!s) return '—'; const [y,m,d] = s.split('-'); return d+'/'+m+'/'+y; };
            const effectiveDateRaw = toStr || new Date().toISOString().slice(0, 10);
            const asOnDate = fmtDate(effectiveDateRaw);
            const periodFrom = fmtDate(fromStr) || 'FY Start';
            const periodTo   = fmtDate(toStr)   || asOnDate;
            const fmtQtyCell = (n) => {
                const num = Number(n) || 0;
                return num === 0 ? '-' : num.toLocaleString('en-IN');
            };
            let thRow = '<th class="bw-print-sr">Sr.</th><th class="bw-print-brand">Brand Name</th><th class="bw-print-tp">TP No.</th>';

            let tbody = '';
            let lastCat = '';
            let lastCatNorm = '';
            let catSr = 0;
            let currentCatSizes = [];
            let catOpen = [];
            let catPur = [];
            let catSale = [];
            let catCl = [];

            const printCategorySizesMap = new Map();
            bwFiltered.forEach(row => {
                const catKey = bwCategoryDisplay(row.category);
                if (!printCategorySizesMap.has(catKey)) printCategorySizesMap.set(catKey, new Set());
                const sizeSet = printCategorySizesMap.get(catKey);
                bwSizeCols.forEach(sz => {
                    const s = row.sizes[sz];
                    if (s && (s.opening || s.purchase || s.sale || s.closing)) {
                        sizeSet.add(sz);
                    }
                });
            });

            // Dynamic column widths based on max value digits per size
            const sizeColMaxDigits = {};
            bwFiltered.forEach(row => {
                bwSizeCols.forEach(sz => {
                    const s = row.sizes[sz];
                    if (!s) return;
                    ['opening', 'purchase', 'sale', 'closing'].forEach(fld => {
                        const v = Number(s[fld]) || 0;
                        if (v > 0) {
                            const d = String(Math.round(v)).length;
                            if (!sizeColMaxDigits[sz] || d > sizeColMaxDigits[sz]) sizeColMaxDigits[sz] = d;
                        }
                    });
                });
            });
            const szW = (sz) => { const d = sizeColMaxDigits[sz] || 1; return d <= 1 ? 14 : d === 2 ? 18 : d === 3 ? 23 : d === 4 ? 28 : d === 5 ? 34 : 40; };

            // Dynamic TP No. column width based on longest formatted TP string
            let maxTpLen = 4;
            bwFiltered.forEach(row => {
                const tpStr = [...row.tpNumbers].map(t => { const j = t.lastIndexOf('/'); return j >= 0 ? t.slice(j + 1) : t; }).join(', ');
                if (tpStr.length > maxTpLen) maxTpLen = tpStr.length;
            });
            const tpW = Math.min(Math.max(Math.ceil(maxTpLen * 5.2), 32), 130);

            function printCatSubHeaders(sizeCols) {
                const nSizes = sizeCols.length;
                const groupBoundaryStyle = (si) => `${si === 0 ? 'border-left:2px solid #000;' : ''}${si === nSizes - 1 ? 'border-right:2px solid #000;' : ''}`;
                const grpHdrStyle = `border-left:2px solid #000;border-right:2px solid #000`;
                let lines = `
                    <tr class="cat-subhead-group">
                        <td class="cat-subhead-stub" style="width:20px">Sr.No</td>
                        <td class="cat-subhead-stub" style="width:88px">Item Name</td>
                        <td class="cat-subhead-stub" style="width:${tpW}px;border-right:2px solid #000">TP No.</td>
                        <td colspan="${nSizes}" style="${grpHdrStyle}">OPENING</td>
                        <td colspan="${nSizes}" style="${grpHdrStyle}">RECEIVED</td>
                        <td colspan="${nSizes}" style="${grpHdrStyle}">SALES</td>
                        <td colspan="${nSizes}" style="${grpHdrStyle}">CLOSING</td>
                    </tr>
                    <tr class="cat-subhead-sizes">
                        <td class="cat-subhead-stub" style="width:20px"></td>
                        <td class="cat-subhead-stub" style="width:88px"></td>
                        <td class="cat-subhead-stub" style="width:${tpW}px;border-right:2px solid #000"></td>
                `;
                ['opening', 'purchase', 'sales', 'closing'].forEach(() => {
                    sizeCols.forEach((sz, si) => {
                        const shortLabel = sz.replace(/\s*ML/i, '').trim();
                        lines += `<td class="bw-print-qty" style="width:${szW(sz)}px;${groupBoundaryStyle(si)}">${esc(shortLabel)}</td>`;
                    });
                });
                lines += '</tr>';
                return lines;
            }

            function printCatTotalRow(catLabel) {
                if (!catLabel) return '';
                const nSizes = currentCatSizes.length;
                const groupBoundaryStyle = (si) => `${si === 0 ? 'border-left:2px solid #000;' : ''}${si === nSizes - 1 ? 'border-right:2px solid #000;' : ''}`;
                let line = `<tr class="cat-total"><td colspan="3" style="text-align:left">${catLabel} Total</td>`;
                catOpen.forEach((v, si) => { const sz = currentCatSizes[si]; line += `<td class="bw-print-qty" style="width:${szW(sz)}px;${groupBoundaryStyle(si)}">${fmtQtyCell(v)}</td>`; });
                catPur.forEach((v, si)  => { const sz = currentCatSizes[si]; line += `<td class="bw-print-qty" style="width:${szW(sz)}px;${groupBoundaryStyle(si)}">${fmtQtyCell(v)}</td>`; });
                catSale.forEach((v, si) => { const sz = currentCatSizes[si]; line += `<td class="bw-print-qty" style="width:${szW(sz)}px;${groupBoundaryStyle(si)}">${fmtQtyCell(v)}</td>`; });
                catCl.forEach((v, si)   => { const sz = currentCatSizes[si]; line += `<td class="bw-print-qty" style="width:${szW(sz)}px;${groupBoundaryStyle(si)}">${fmtQtyCell(v)}</td>`; });
                line += '</tr>';
                return line;
            }

            bwFiltered.forEach((row) => {
                const catKey = bwCategoryDisplay(row.category);
                const currentCatNorm = normalizeBwCategory(row.category);
                if (catKey !== lastCat) {
                    if (lastCat) tbody += printCatTotalRow(lastCat);
                    const sizeSet = printCategorySizesMap.get(catKey) || new Set();
                    currentCatSizes = bwSizeCols.filter(sz => sizeSet.has(sz));
                    catOpen = new Array(currentCatSizes.length).fill(0);
                    catPur = new Array(currentCatSizes.length).fill(0);
                    catSale = new Array(currentCatSizes.length).fill(0);
                    catCl = new Array(currentCatSizes.length).fill(0);
                    catSr = 0;
                    const newPageClass = (lastCatNorm === 'SPIRITS' && currentCatNorm !== 'SPIRITS') ? ' new-page' : '';
                    lastCat = catKey;
                    lastCatNorm = currentCatNorm;
                    const totalCols = 3 + currentCatSizes.length * 4;
                    if (tbody !== '' && !newPageClass) tbody += `<tr class="cat-spacer"><td colspan="${totalCols}"></td></tr>`;
                    tbody += `<tr class="cat-head${newPageClass}"><td colspan="${totalCols}"><div class="cat-head-main">${catKey}</div></td></tr>`;
                    tbody += printCatSubHeaders(currentCatSizes);
                }
                catSr++;
                const printBrandShort = bwShortBrand(row.brandName, 20);
                const tpFormatted = [...row.tpNumbers].map(t => { const j = t.lastIndexOf('/'); return j >= 0 ? t.slice(j + 1) : t; }).join(', ') || '—';
                tbody += `<tr><td class="bw-print-sr">${catSr}</td><td class="bw-print-brand" title="${esc(row.brandName)}">${esc(printBrandShort)}</td><td class="bw-print-tp" style="width:${tpW}px">${esc(tpFormatted)}</td>`;
                const nSizes = currentCatSizes.length;
                const groupBoundaryStyle = (si) => `${si === 0 ? 'border-left:2px solid #000;' : ''}${si === nSizes - 1 ? 'border-right:2px solid #000;' : ''}`;
                ['opening', 'purchase', 'sale', 'closing'].forEach(fld => {
                    currentCatSizes.forEach((sz, si) => {
                        const v = row.sizes[sz] ? row.sizes[sz][fld] : 0;
                        if (fld === 'opening') catOpen[si] += v;
                        if (fld === 'purchase') catPur[si] += v;
                        if (fld === 'sale') catSale[si] += v;
                        if (fld === 'closing') catCl[si] += v;
                        tbody += `<td class="bw-print-qty" style="width:${szW(sz)}px;${groupBoundaryStyle(si)}">${fmtQtyCell(v)}</td>`;
                    });
                });
                tbody += '</tr>';
            });

            tbody += printCatTotalRow(lastCat);

            printWithIframe(`<!DOCTYPE html><html><head><title>Brandwise Report — ${barName}</title>
                <style>
                @page{size:legal landscape;margin:6mm 9mm 10mm}
                @page :first{margin-top:6mm}
                html,body{width:100%;margin:0;padding:0}
                body{font-family:'Segoe UI',Arial,sans-serif;padding:0 1mm;font-size:8pt;color:#000}
                .print-head{text-align:center;border-top:2px solid #000;border-bottom:1px solid #000;padding:5px 0;margin-bottom:4px}
                .ph-bar-name{font-size:14pt;font-weight:800;text-transform:uppercase;letter-spacing:.5px;color:#000;line-height:1.3}
                .ph-meta-line{font-size:8pt;color:#000;margin-top:1px;line-height:1.5}
                table{border-collapse:collapse;width:auto;table-layout:fixed}
                th,td{border:1px dashed #000;padding:1.5px 2px;font-size:7pt;line-height:1.2}
                th{background:#fff;font-weight:700;border:1.5px solid #000;white-space:nowrap;color:#000}
                td{white-space:nowrap;overflow:hidden;color:#000}
                .bw-print-sr{width:20px;text-align:center}
                .bw-print-brand{width:88px;text-align:left;overflow:hidden;text-overflow:ellipsis}
                .bw-print-tp{text-align:center;font-size:6.5pt;white-space:normal;word-break:break-word;line-height:1.2}
                .bw-print-qty{text-align:center;font-size:6.8pt;padding:1px 2px;white-space:nowrap}
                .cat-head td{background:#fff;font-weight:700;color:#000;border:2px solid #000}
                .cat-head-main{font-size:8.5pt;font-weight:700}
                .cat-head.new-page td{page-break-before:always;break-before:page}
                .cat-subhead-group td{background:#fff;font-size:6.8pt;font-weight:700;text-align:center;white-space:nowrap;line-height:1.1;padding:1px 2px;color:#000;border:1.5px solid #000}
                .cat-subhead-stub{font-size:6.8pt;text-align:left;font-weight:700;white-space:nowrap;color:#000}
                .cat-subhead-sizes td{background:#fff;font-size:6.5pt;text-align:center;white-space:nowrap;border:1.5px solid #000;font-weight:700;color:#000}
                .cat-total td{background:#fff;font-weight:700;border-top:2px solid #000;border-bottom:2px solid #000;border-left:1px dashed #000;border-right:1px dashed #000;color:#000}
                .cat-spacer td{border:none;height:5px;background:transparent;padding:0}
                tr{page-break-inside:avoid}
                thead{display:table-header-group}
                .repeat-hdr td{border-top:1.5px solid #000;border-bottom:1px solid #000;border-left:none;border-right:none;font-size:7pt;font-weight:700;padding:2px 0;white-space:nowrap;color:#000;background:#fff}
                .repeat-hdr .rh-bar{display:inline-block;margin-right:16px}
                .repeat-hdr .rh-lbl{font-weight:700;margin-right:2px;color:#000}
                .repeat-hdr .rh-val{font-weight:400;color:#000}
                .repeat-hdr .rh-pg{float:right;font-weight:600}
                </style>
            </head><body>
                <div style="width:max-content;min-width:100%">
                <div class="print-head">
                    <div class="ph-bar-name">${barName || '—'}</div>
                    <div class="ph-meta-line">License No: ${licNo}</div>
                    <div class="ph-meta-line">Address: ${address || '—'}</div>
                    <div class="ph-meta-line">Date: ${asOnDate}</div>
                </div>
                <table>
                <thead><tr class="repeat-hdr"><td colspan="999"
                ><span class="rh-bar"><span class="rh-lbl">${barName || '—'}</span></span><span class="rh-bar"><span class="rh-lbl">Lic:</span> <span class="rh-val">${licNo}</span></span><span class="rh-bar"><span class="rh-lbl">Date:</span> <span class="rh-val">${asOnDate}</span></span><span class="rh-bar"><span class="rh-lbl">FY:</span> <span class="rh-val">${fy || '—'}</span></span></td></tr></thead>
                <tbody>${tbody}</tbody></table>
                </div>
            </body></html>`);
        }

        // Event wiring
        if (bwGoBtn)          bwGoBtn.addEventListener('click', generateBrandwise);
        if (bwSingleDate)     bwSingleDate.addEventListener('change', () => {
            if (bwSingleDate.value) {
                setBwMode('single'); // ensure we're in single mode
                generateBrandwise();
            }
        });
        if (bwDateFrom)       bwDateFrom.addEventListener('change', () => {
            if (!bwIsRangeMode) setBwMode('range');
        });
        if (bwDateTo)         bwDateTo.addEventListener('change', () => {
            if (!bwIsRangeMode) setBwMode('range');
        });
        if (bwCategoryFilter) bwCategoryFilter.addEventListener('change', applyBwFilters);
        if (bwMmlToggle)      bwMmlToggle.addEventListener('change', applyBwFilters);
        if (bwSearch)         bwSearch.addEventListener('input', applyBwFilters);
        if (bwExportBtn)      bwExportBtn.addEventListener('click', exportBwCSV);
        if (bwPrintBtn)       bwPrintBtn.addEventListener('click', printBwReport);
        document.getElementById('bwPreviewBtn')?.addEventListener('click', () => { _previewMode = true; printBwReport(); });

        // Auto-generate when sub-view becomes visible
        const bwPanel = document.getElementById('sub-brandwise');
        if (bwPanel) {
            const bwObserver = new MutationObserver((mutations) => {
                mutations.forEach(m => {
                    if (m.attributeName === 'class' && bwPanel.classList.contains('active') && bwRows.length === 0) {
                        generateBrandwise();
                    }
                });
            });
            bwObserver.observe(bwPanel, { attributes: true, attributeFilter: ['class'] });
        }
    }

    // ═══════════════════════════════════════════════════════
    //  MONTHLY STATEMENT — Form F.L.R.-4
    //  Maharashtra Excise Monthly Return (Pivoted Format)
    //  Columns = Sizes grouped by Category
    //  Rows    = Opening → Received → Total → Cumul Rcpt → Sold → Cumul Sale → Breakage → Closing
    // ═══════════════════════════════════════════════════════
    {
        const msMonthSelect = document.getElementById('msMonthSelect');
        const msYearSelect  = document.getElementById('msYearSelect');
        const msGoBtn       = document.getElementById('msGoBtn');
        const msExportBtn   = document.getElementById('msExportBtn');
        const msPdfBtn      = document.getElementById('msPdfBtn');
        const msBwPdfBtn    = document.getElementById('msBwPdfBtn');
        const msFormHeader  = document.getElementById('msFormHeader');
        const msTableHead   = document.getElementById('msTableHead');
        const msTableBody   = document.getElementById('msTableBody');
        const msTable       = document.getElementById('msTable');
        const msEmpty       = document.getElementById('msEmpty');
        const msBulkSection = document.getElementById('msBulkSection');
        const msBulkMonth   = document.getElementById('msBulkMonth');
        const msBulkBody    = document.getElementById('msBulkBody');
        const msBulkFoot    = document.getElementById('msBulkFoot');

        let msGenerated = false;
        let msLastData  = null; // store for export/print

        // ── FY defaults ──
        const msFyStartYear = (() => {
            const now = new Date();
            return now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
        })();

        // Populate year dropdown
        if (msYearSelect) {
            for (let y = msFyStartYear + 1; y >= msFyStartYear - 2; y--) {
                const opt = document.createElement('option');
                opt.value = y;
                opt.textContent = y;
                msYearSelect.appendChild(opt);
            }
        }

        // Auto-select current month
        {
            const now = new Date();
            if (msMonthSelect) msMonthSelect.value = now.getMonth();
            if (msYearSelect)  msYearSelect.value  = now.getFullYear();
        }

        /* ── COLUMN DEFINITIONS for Form F.L.R.-4 ── */
        // Each category group has standard sizes (in ml). We detect sizes from actual data.
        const CATEGORY_GROUPS = [
            { key: 'Spirits', label: 'SPIRITS', categories: ['Spirits', 'MML'] },
            { key: 'Wines',   label: 'WINES',   categories: ['Wines'] },
            { key: 'BeerStrong', label: 'BEER STRONG', categories: ['Fermented Beer'] },
            { key: 'BeerMild',   label: 'BEER MILD',   categories: ['Mild Beer'] },
        ];

        /* ── ROW DEFINITIONS for Form F.L.R.-4 ── */
        const ROW_DEFS = [
            { key: 'opening',      label: 'Opening Bal. of the begin. of the mon.',     short: 'Opening Bal.' },
            { key: 'received',     label: 'Received during the month',                  short: 'Received' },
            { key: 'total',        label: 'Total',                                      short: 'Total' },
            { key: 'cumulReceipt', label: 'Rece from 1st Apr to end of the month',      short: 'Cumul. Receipt' },
            { key: 'sold',         label: 'Sold Per. Hol during month',                 short: 'Sold' },
            { key: 'cumulSale',    label: 'Sold Per. Hold. 1st Apr. to end of',         short: 'Cumul. Sale' },
            { key: 'breakage',     label: 'Breakage and wastage dur. the curr month',   short: 'Breakage' },
            { key: 'closing',      label: 'Clos Bal. of the end of the month',          short: 'Closing Bal.' },
        ];

        const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];

        /* ── Helpers ── */
        function extractML(sizeStr) {
            const m = (sizeStr || '').match(/^(\d+)\s*ML/i);
            if (m) return parseInt(m[1]);
            const l = (sizeStr || '').match(/^(\d+)\s*Ltr/i);
            if (l) return parseInt(l[1]) * 1000;
            return 0;
        }

        function sizeGroup(sizeStr) {
            const ml = extractML(sizeStr);
            return ml > 0 ? ml + ' ML' : (sizeStr || '').trim();
        }

        function rangeDate(year, month, day) {
            return `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        }

        function getMonthRange(year, monthIdx) {
            const start = new Date(year, monthIdx, 1);
            const end   = new Date(year, monthIdx + 1, 0);
            const fmt = d => d.toISOString().slice(0, 10);
            return { start: fmt(start), end: fmt(end) };
        }

        function getFyStart(year, monthIdx) {
            // FY start is April of whichever FY this month belongs to
            const fyStartYear = monthIdx >= 3 ? year : year - 1;
            return `${fyStartYear}-04-01`;
        }

        function v(n) { return n || 0; }
        function fmtN(n) { return (n || 0).toLocaleString('en-IN'); }
        function fmtBulk(n) { return (n || 0).toFixed(3); }
        function colKey(gk, ml) { return gk === '_total' ? '_total' : gk + '_' + ml; }

        /* ── Categorize a product into our form groups ── */
        function getCatGroup(category) {
            for (const g of CATEGORY_GROUPS) {
                if (g.categories.includes(category)) return g.key;
            }
            return 'Spirits'; // default fallback
        }

        /* ── MAIN GENERATE ── */
        async function generateMonthlyStatement() {
          try {
            if (!activeBar.barName) { console.warn('MS: no activeBar.barName'); return; }
            const params = { barName: activeBar.barName, financialYear: activeBar.financialYear || '' };
            console.log('MS: generating with params', params);

            const selMonth = parseInt(msMonthSelect ? msMonthSelect.value : new Date().getMonth());
            const selYear  = parseInt(msYearSelect  ? msYearSelect.value  : new Date().getFullYear());
            const monthName = MONTH_NAMES[selMonth];
            const fyStartDate = getFyStart(selYear, selMonth);

            // ── Update form header ──
            if (msFormHeader) {
                msFormHeader.innerHTML = `
                    <div class="ms-fh-row">
                        <span class="ms-fh-form">FORM F.L.R.- 4 &nbsp; [See Rule 15 (II)]</span>
                    </div>
                    <div class="ms-fh-row ms-fh-subtitle">
                        Monthly return of Transactions of Foreign Liquor effected by Vendors, Hotels, Clubs Licensee.
                    </div>
                    <div class="ms-fh-row">
                        <span>Name of the Licensee : <strong>${activeBar.barName || '—'}</strong></span>
                    </div>
                    <div class="ms-fh-row">
                        <span>Address : <strong>${[activeBar.address, activeBar.area, activeBar.city].filter(Boolean).join(', ') || '—'}</strong></span>
                        <span class="ms-fh-right">Licence No: <strong>${activeBar.licenseNo || '—'}</strong></span>
                    </div>
                    <div class="ms-fh-row">
                        <span></span>
                        <span class="ms-fh-right">Month of : <strong>${monthName}, &nbsp; ${selYear}</strong></span>
                    </div>
                `;
            }

            // ── Fetch all data ──
            const [osResult, tpResult, salesItems, productsResult] = await Promise.all([
                window.electronAPI.getOpeningStock(params),
                window.electronAPI.getTps(params),
                getSalesData(params),
                window.electronAPI.getProducts(params),
            ]);

            console.log('MS: osResult', osResult?.success, 'tpResult', tpResult?.success, 'salesItems', salesItems?.length, 'products', productsResult?.success, productsResult?.products?.length);

            // ── Build product lookup: code → { category, size, sizeML } ──
            const productMap = new Map();
            if (productsResult.success && productsResult.products) {
                for (const p of productsResult.products) {
                    productMap.set(p.code, { category: p.category || '', size: p.size || '', ml: extractML(p.size) });
                }
            }
            console.log('MS: productMap size', productMap.size);

            // ── Determine all sizes per category group from products ──
            const groupSizes = {}; // groupKey → Set of ml values
            for (const g of CATEGORY_GROUPS) groupSizes[g.key] = new Set();
            for (const [, p] of productMap) {
                const gk = getCatGroup(p.category);
                const ml = p.ml;
                if (ml > 0) groupSizes[gk].add(ml);
            }

            // Sort sizes descending within each group (matching form: 2000, 1000, 750, 375, 180, 90, 60)
            const sortedGroupSizes = {};
            for (const g of CATEGORY_GROUPS) {
                sortedGroupSizes[g.key] = [...groupSizes[g.key]].sort((a, b) => b - a);
            }

            // ── Build flat column list: [{groupKey, groupLabel, ml, label}] ──
            const columns = [];
            for (const g of CATEGORY_GROUPS) {
                for (const ml of sortedGroupSizes[g.key]) {
                    columns.push({ groupKey: g.key, groupLabel: g.label, ml, label: ml + ' ml' });
                }
            }
            // Add BOTTLE/BULK(LTR) total column
            columns.push({ groupKey: '_total', groupLabel: 'BOTTLE/BULK(LTR)', ml: 0, label: '' });

            // ── Helper: classify item into [groupKey, ml] ──
            function classifyItem(code, category, size) {
                const prod = productMap.get(code);
                const cat  = category || (prod ? prod.category : '');
                const sz   = size || (prod ? prod.size : '');
                const gk   = getCatGroup(cat);
                const ml   = extractML(sz);
                return { gk, ml };
            }

            // ── Current month range ──
            const curRange = getMonthRange(selYear, selMonth);

            // ── DATA STRUCTURES: For each row, store values per column ──
            function emptyRow() {
                const r = {};
                for (const c of columns) r[colKey(c.groupKey, c.ml)] = 0;
                return r;
            }

            const data = {};
            for (const rd of ROW_DEFS) data[rd.key] = emptyRow();

            // ── 1. OPENING STOCK (beginning of month) ──
            // Opening = FY opening + all receipts before this month - all sales before this month
            // Step 1a: FY opening stock
            const fyOpening = emptyRow();
            const osData = osResult.success ? osResult.data : null;
            if (osData && osData.entries) {
                for (const entry of osData.entries) {
                    const { gk, ml } = classifyItem(entry.code, entry.category, entry.size || entry.sizeLabel);
                    const ck = colKey(gk, ml);
                    if (ck in fyOpening) fyOpening[ck] += (entry.totalBtl || 0);
                    fyOpening['_total'] += (entry.totalBtl || 0);
                }
            }

            // Step 1b: all receipts from FY start to day before this month
            const prevRange = { start: fyStartDate, end: (() => {
                const d = new Date(selYear, selMonth, 0); // last day of prev month
                return d.toISOString().slice(0, 10);
            })() };

            const prevReceipts = emptyRow();
            const curReceipts  = emptyRow();
            if (tpResult.success && tpResult.tps) {
                for (const tp of tpResult.tps) {
                    const dt = tp.tpDate || tp.receivedDate || '';
                    for (const item of (tp.items || [])) {
                        const { gk, ml } = classifyItem(item.code, item.category, item.size || item.sizeLabel);
                        const ck = colKey(gk, ml);
                        const btl = item.totalBtl || 0;
                        if (dt >= curRange.start && dt <= curRange.end) {
                            if (ck in curReceipts) curReceipts[ck] += btl;
                            curReceipts['_total'] += btl;
                        } else if (dt >= prevRange.start && dt <= prevRange.end) {
                            if (ck in prevReceipts) prevReceipts[ck] += btl;
                            prevReceipts['_total'] += btl;
                        }
                    }
                }
            }

            // Step 1c: all sales from FY start to day before this month
            const prevSales = emptyRow();
            const curSales  = emptyRow();
            for (const item of salesItems) {
                const { gk, ml } = classifyItem(item.code, item.category, item.size);
                const ck = colKey(gk, ml);
                const btl = item.totalBtl || 0;
                const dt = item.billDate || '';
                if (dt >= curRange.start && dt <= curRange.end) {
                    if (ck in curSales) curSales[ck] += btl;
                    curSales['_total'] += btl;
                } else if (dt >= prevRange.start && dt <= prevRange.end) {
                    if (ck in prevSales) prevSales[ck] += btl;
                    prevSales['_total'] += btl;
                }
            }

            // ── Compute each row ──
            const allKeys = Object.keys(data['opening']);

            // Cumulative receipts from FY start to end of prev month
            const cumulReceiptUpToPrev = {};
            const cumulSaleUpToPrev = {};
            for (const k of allKeys) {
                cumulReceiptUpToPrev[k] = v(prevReceipts[k]);
                cumulSaleUpToPrev[k]    = v(prevSales[k]);
            }

            for (const k of allKeys) {
                // Opening = FY Opening + prev receipts - prev sales
                data['opening'][k] = v(fyOpening[k]) + v(prevReceipts[k]) - v(prevSales[k]);

                // Received during month
                data['received'][k] = v(curReceipts[k]);

                // Total = Opening + Received
                data['total'][k] = data['opening'][k] + data['received'][k];

                // Cumul receipt from 1st Apr to end of this month
                data['cumulReceipt'][k] = v(prevReceipts[k]) + v(curReceipts[k]);

                // Sold during month
                data['sold'][k] = v(curSales[k]);

                // Cumul sale from 1st Apr to end of this month
                data['cumulSale'][k] = v(prevSales[k]) + v(curSales[k]);

                // Breakage (we don't track this yet — always 0)
                data['breakage'][k] = 0;

                // Closing = Total - Sold - Breakage
                data['closing'][k] = data['total'][k] - data['sold'][k] - data['breakage'][k];
            }

            // ── Prune size columns that are zero across ALL rows ──
            for (const g of CATEGORY_GROUPS) {
                sortedGroupSizes[g.key] = sortedGroupSizes[g.key].filter(ml => {
                    const ck = colKey(g.key, ml);
                    return ROW_DEFS.some(rd => (data[rd.key][ck] || 0) !== 0);
                });
            }
            // Rebuild columns to match pruned sizes
            columns.length = 0;
            for (const g of CATEGORY_GROUPS) {
                for (const ml of sortedGroupSizes[g.key]) {
                    columns.push({ groupKey: g.key, groupLabel: g.label, ml, label: ml + ' ml' });
                }
            }
            columns.push({ groupKey: '_total', groupLabel: 'BOTTLE/BULK(LTR)', ml: 0, label: '' });

            // ── Store for export/print ──
            msLastData = { data, columns, sortedGroupSizes, selMonth, selYear, monthName };

            console.log('MS: columns count', columns.length, 'data keys', Object.keys(data));
            msGenerated = true;
            renderMsTable(data, columns, sortedGroupSizes);
            renderBulkLtrReport(data, columns, sortedGroupSizes, monthName, selYear);
          } catch (err) {
            console.error('MS generateMonthlyStatement ERROR:', err);
          }
        }

        /* ── Render the FLR-4 table ── */
        function renderMsTable(data, columns, sortedGroupSizes) {
            if (!msTableHead || !msTableBody) return;

            if (columns.length <= 1) {
                if (msEmpty) msEmpty.classList.remove('hidden');
                if (msTable) msTable.style.display = 'none';
                if (msBulkSection) msBulkSection.style.display = 'none';
                return;
            }
            if (msEmpty) msEmpty.classList.add('hidden');
            if (msTable) msTable.style.display = '';
            if (msBulkSection) msBulkSection.style.display = '';

            // ── HEADER ROW 1: Category groups ──
            let h1 = '<th class="ms-th ms-th-frozen ms-th-particulars" rowspan="3">PARTICULARS</th>';
            for (const g of CATEGORY_GROUPS) {
                const sizes = sortedGroupSizes[g.key] || [];
                if (sizes.length === 0) continue;
                const cls = g.key === 'Spirits' ? 'ms-th-spirits' :
                            g.key === 'Wines'   ? 'ms-th-wines' :
                            g.key === 'BeerStrong' ? 'ms-th-beer-strong' :
                            'ms-th-beer-mild';
                h1 += `<th class="ms-th ms-th-group ${cls}" colspan="${sizes.length}">--${g.label}--</th>`;
            }
            h1 += '<th class="ms-th ms-th-group ms-th-total-hdr" rowspan="3">BOTTLE/<br>BULK(LTR)</th>';

            // ── HEADER ROW 2: Size values (2000, 1000, 750...) ──
            let h2 = '';
            for (const g of CATEGORY_GROUPS) {
                const sizes = sortedGroupSizes[g.key] || [];
                for (const ml of sizes) {
                    h2 += `<th class="ms-th ms-th-size">${ml}</th>`;
                }
            }

            // ── HEADER ROW 3: "ml" labels ──
            let h3 = '';
            for (const g of CATEGORY_GROUPS) {
                const sizes = sortedGroupSizes[g.key] || [];
                for (let i = 0; i < sizes.length; i++) {
                    h3 += `<th class="ms-th ms-th-ml">ml</th>`;
                }
            }

            msTableHead.innerHTML = `
                <tr class="ms-header-cat">${h1}</tr>
                <tr class="ms-header-size">${h2}</tr>
                <tr class="ms-header-ml">${h3}</tr>
            `;

            // ── BODY ROWS ──
            msTableBody.innerHTML = '';
            for (const rd of ROW_DEFS) {
                const tr = document.createElement('tr');
                const isTotal = rd.key === 'total';
                const isClosing = rd.key === 'closing';
                const isCumul = rd.key === 'cumulReceipt' || rd.key === 'cumulSale';
                tr.className = 'ms-data-row' +
                    (isTotal ? ' ms-row-total' : '') +
                    (isClosing ? ' ms-row-closing' : '') +
                    (isCumul ? ' ms-row-cumul' : '');

                let cells = `<td class="ms-td ms-td-frozen ms-td-label">${rd.label}</td>`;
                for (const g of CATEGORY_GROUPS) {
                    const sizes = sortedGroupSizes[g.key] || [];
                    for (const ml of sizes) {
                        const ck = colKey(g.key, ml);
                        const val = data[rd.key][ck] || 0;
                        const cls = val === 0 ? 'ms-td ms-td-num ms-td-zero' :
                                    val < 0  ? 'ms-td ms-td-num ms-td-neg' :
                                    'ms-td ms-td-num';
                        cells += `<td class="${cls}">${val || ''}</td>`;
                    }
                }
                // Total column
                const totalVal = data[rd.key]['_total'] || 0;
                const totalCls = totalVal === 0 ? 'ms-td ms-td-num ms-td-zero' :
                                 totalVal < 0  ? 'ms-td ms-td-num ms-td-neg' :
                                 'ms-td ms-td-num ms-td-total-val';
                cells += `<td class="${totalCls}">${totalVal || ''}</td>`;

                tr.innerHTML = cells;
                msTableBody.appendChild(tr);

                // Add separator after "Total" and "Cumul Sale" rows
                if (rd.key === 'total' || rd.key === 'cumulSale') {
                    const sep = document.createElement('tr');
                    sep.className = 'ms-sep-row';
                    const totalCols = 1 + columns.length;
                    sep.innerHTML = `<td colspan="${totalCols}" class="ms-td-sep"></td>`;
                    msTableBody.appendChild(sep);
                }
            }
        }

        /* ── Bulk LTR Report ── */
        function renderBulkLtrReport(data, columns, sortedGroupSizes, monthName, year) {
            if (!msBulkBody || !msBulkFoot) return;
            if (msBulkMonth) msBulkMonth.textContent = `${monthName}, ${year}`;

            msBulkBody.innerHTML = '';

            // For each category group, compute bulk liters
            // Bulk LTR = sum of (bottles × ml / 1000) for each category
            const bulkRows = [];
            let gOpenTotal = 0, gRecTotal = 0, gSaleTotal = 0, gCloseTotal = 0;

            for (const g of CATEGORY_GROUPS) {
                const sizes = sortedGroupSizes[g.key] || [];
                if (sizes.length === 0) continue;

                let bOpen = 0, bRec = 0, bSale = 0, bClose = 0;
                for (const ml of sizes) {
                    const ck = colKey(g.key, ml);
                    bOpen  += (data['opening'][ck] || 0)  * ml / 1000;
                    bRec   += (data['received'][ck] || 0) * ml / 1000;
                    bSale  += (data['sold'][ck] || 0)     * ml / 1000;
                    bClose += (data['closing'][ck] || 0)  * ml / 1000;
                }

                bulkRows.push({ label: g.label, open: bOpen, rec: bRec, sale: bSale, close: bClose });
                gOpenTotal += bOpen; gRecTotal += bRec; gSaleTotal += bSale; gCloseTotal += bClose;
            }

            for (const row of bulkRows) {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="ms-btd ms-btd-label">${row.label}</td>
                    <td class="ms-btd ms-btd-num">${fmtBulk(row.open)}</td>
                    <td class="ms-btd ms-btd-num">${fmtBulk(row.rec)}</td>
                    <td class="ms-btd ms-btd-num">${fmtBulk(row.sale)}</td>
                    <td class="ms-btd ms-btd-num">${fmtBulk(row.close)}</td>
                `;
                msBulkBody.appendChild(tr);
            }

            msBulkFoot.innerHTML = `
                <tr class="ms-bulk-total">
                    <td class="ms-btd ms-btd-label"><strong>TOTAL :</strong></td>
                    <td class="ms-btd ms-btd-num"><strong>${fmtBulk(gOpenTotal)}</strong></td>
                    <td class="ms-btd ms-btd-num"><strong>${fmtBulk(gRecTotal)}</strong></td>
                    <td class="ms-btd ms-btd-num"><strong>${fmtBulk(gSaleTotal)}</strong></td>
                    <td class="ms-btd ms-btd-num"><strong>${fmtBulk(gCloseTotal)}</strong></td>
                </tr>
            `;
        }

        /* ── Export CSV ── */
        function exportMsCsv() {
            if (!msLastData) return;
            const { data, columns, sortedGroupSizes, monthName, selYear } = msLastData;
            const csvRows = [];

            // Header
            let hdr = 'Particulars';
            for (const g of CATEGORY_GROUPS) {
                for (const ml of (sortedGroupSizes[g.key] || [])) {
                    hdr += `,${g.label} ${ml}ml`;
                }
            }
            hdr += ',TOTAL';
            csvRows.push(hdr);

            // Data rows
            for (const rd of ROW_DEFS) {
                let line = `"${rd.label}"`;
                for (const g of CATEGORY_GROUPS) {
                    for (const ml of (sortedGroupSizes[g.key] || [])) {
                        line += `,${data[rd.key][colKey(g.key, ml)] || 0}`;
                    }
                }
                line += `,${data[rd.key]['_total'] || 0}`;
                csvRows.push(line);
            }

            const blob = new Blob(['\uFEFF' + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `FLR4_${activeBar.barName || 'Bar'}_${monthName}_${selYear}.csv`;
            a.click();
            URL.revokeObjectURL(a.href);
        }

        /* ── Print ── */
        function printMsStatement(bw = false) {
            if (!msLastData) return;
            const { data, columns, sortedGroupSizes, monthName, selYear } = msLastData;
            const barName = activeBar.barName || '';
            const address = [activeBar.address, activeBar.area, activeBar.city].filter(Boolean).join(', ') || '—';
            const licNo   = activeBar.licenseNo || '—';

            // ── Header row 1: category groups ──
            let ph1 = '<th rowspan="3" class="th-part">PARTICULARS</th>';
            const catColors = bw
                ? { Spirits:'#1a1a1a', Wines:'#1a1a1a', BeerStrong:'#2a2a2a', BeerMild:'#2a2a2a' }
                : { Spirits:'#1e3a8a', Wines:'#166534', BeerStrong:'#9a3412', BeerMild:'#854d0e' };
            for (const g of CATEGORY_GROUPS) {
                const sizes = sortedGroupSizes[g.key] || [];
                if (sizes.length === 0) continue;
                ph1 += `<th colspan="${sizes.length}" class="th-cat" style="background:${catColors[g.key]||'#1e3a8a'}">${g.label.toUpperCase()}</th>`;
            }
            ph1 += '<th rowspan="3" class="th-total-hdr">BOTTLE /<br>BULK<br>(LTR)</th>';

            // ── Header rows 2 & 3: sizes + ml ──
            let ph2 = '', ph3 = '';
            const sizeColors = bw
                ? { Spirits:'#333', Wines:'#333', BeerStrong:'#444', BeerMild:'#444' }
                : { Spirits:'#1e40af', Wines:'#15803d', BeerStrong:'#c2410c', BeerMild:'#a16207' };
            for (const g of CATEGORY_GROUPS) {
                for (const ml of (sortedGroupSizes[g.key] || [])) {
                    ph2 += `<th class="th-size" style="background:${sizeColors[g.key]||'#1e40af'}">${ml}</th>`;
                    ph3 += `<th class="th-ml" style="background:${sizeColors[g.key]||'#1e40af'}">ml</th>`;
                }
            }

            // ── Body rows ──
            let tbody = '';
            const accentRows = bw
                ? { opening:'#f5f5f5', received:'#f0f0f0', total:'#e5e5e5', cumulReceipt:'#ebebeb',
                    sold:'#f9f9f9', cumulSale:'#f0f0f0', breakage:'#f5f5f5', closing:'#e8e8e8' }
                : { opening:'#eff6ff', received:'#f0fdf4', total:'#dbeafe', cumulReceipt:'#e0f2fe',
                    sold:'#fff7ed', cumulSale:'#fef3c7', breakage:'#fee2e2', closing:'#ede9fe' };
            const boldRows = new Set(['total','closing']);
            for (const rd of ROW_DEFS) {
                const bg    = accentRows[rd.key] || '#ffffff';
                const bold  = boldRows.has(rd.key) ? 'font-weight:700;' : '';
                const fsize = boldRows.has(rd.key) ? '8pt' : '7.5pt';
                tbody += `<tr style="background:${bg};${bold}">`;
                tbody += `<td class="td-label" style="font-size:${fsize}">${rd.label}</td>`;
                for (const g of CATEGORY_GROUPS) {
                    for (const ml of (sortedGroupSizes[g.key] || [])) {
                        const val = data[rd.key][colKey(g.key, ml)] || 0;
                        tbody += `<td class="td-num" style="font-size:${fsize}">${val || ''}</td>`;
                    }
                }
                const tot = data[rd.key]['_total'] || 0;
                tbody += `<td class="td-tot" style="font-size:${fsize}">${tot || ''}</td></tr>`;
                if (rd.key === 'total' || rd.key === 'cumulSale') {
                    const totalCols = 1 + columns.length;
                    tbody += `<tr><td colspan="${totalCols}" style="height:3px;background:#cbd5e1;padding:0"></td></tr>`;
                }
            }

            // ── Bulk LTR computation ──
            const fmtB = n => (n || 0).toFixed(3);
            let bulkTbody = '';
            let gOT = 0, gRT = 0, gST = 0, gCT = 0;
            for (const g of CATEGORY_GROUPS) {
                const bSizes = sortedGroupSizes[g.key] || [];
                if (bSizes.length === 0) continue;
                let bO = 0, bR = 0, bS = 0, bC = 0;
                for (const ml of bSizes) {
                    const ck = colKey(g.key, ml);
                    bO += (data['opening'][ck]  || 0) * ml / 1000;
                    bR += (data['received'][ck] || 0) * ml / 1000;
                    bS += (data['sold'][ck]     || 0) * ml / 1000;
                    bC += (data['closing'][ck]  || 0) * ml / 1000;
                }
                bulkTbody += `<tr><td class="blk-label">${g.label}</td><td class="blk-num">${fmtB(bO)}</td><td class="blk-num">${fmtB(bR)}</td><td class="blk-num">${fmtB(bS)}</td><td class="blk-num">${fmtB(bC)}</td></tr>`;
                gOT += bO; gRT += bR; gST += bS; gCT += bC;
            }
            bulkTbody += `<tr class="blk-total"><td class="blk-label"><strong>TOTAL :</strong></td><td class="blk-num"><strong>${fmtB(gOT)}</strong></td><td class="blk-num"><strong>${fmtB(gRT)}</strong></td><td class="blk-num"><strong>${fmtB(gST)}</strong></td><td class="blk-num"><strong>${fmtB(gCT)}</strong></td></tr>`;

            const __msPdfHtml = `<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>F.L.R.-4 — ${barName} — ${monthName} ${selYear}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  html,body,*{-webkit-print-color-adjust:exact;print-color-adjust:exact;color-adjust:exact}
  body{font-family:'Segoe UI',Arial,Helvetica,sans-serif;font-size:8.5pt;color:#111;background:#fff;
       padding:8mm 8mm 6mm}
  @page{size:A4 landscape;margin:8mm 6mm}
  @media print{body{padding:0}}

  /* ── Top accent bar ── */
  .top-bar{height:5px;background:linear-gradient(90deg,#1e3a8a 0%,#1e40af 40%,#166534 40%,#166534 55%,#9a3412 55%,#9a3412 75%,#854d0e 75%,#854d0e 100%);margin-bottom:8px;border-radius:2px}

  /* ── Letterhead ── */
  .letterhead{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px}
  .lh-left{flex:1}
  .lh-center{flex:2;text-align:center}
  .lh-right{flex:1;text-align:right}
  .form-badge{display:inline-block;background:#1e3a8a;color:#fff;font-size:8pt;font-weight:700;
              padding:3px 10px;border-radius:3px;letter-spacing:.05em;margin-bottom:3px}
  .form-title{font-size:13pt;font-weight:800;color:#1e3a8a;letter-spacing:.03em;text-transform:uppercase}
  .form-sub{font-size:7pt;color:#475569;margin-top:2px;font-style:italic}
  .govt-label{font-size:7pt;color:#64748b;text-transform:uppercase;letter-spacing:.08em}
  .bar-name-big{font-size:11pt;font-weight:700;color:#0f172a}

  /* ── Info box ── */
  .info-box{border:1.5px solid #1e3a8a;border-radius:4px;padding:5px 8px;margin-bottom:8px;
            display:grid;grid-template-columns:1fr 1fr;gap:2px 16px;background:#f8fafc}
  .info-row{display:flex;gap:4px;font-size:7.5pt;color:#1e293b}
  .info-label{color:#64748b;white-space:nowrap}
  .info-val{font-weight:600;color:#0f172a}
  .info-right{text-align:right}

  /* ── Divider ── */
  .divider{height:1.5px;background:#1e3a8a;margin-bottom:6px;opacity:.25}

  /* ── Table ── */
  table{border-collapse:collapse;width:100%;font-size:7.5pt}
  .th-part{background:#0f172a;color:#fff;text-align:left;padding:4px 6px;min-width:180px;
           font-size:7pt;font-weight:700;letter-spacing:.04em;border:1px solid #1e3a8a;white-space:nowrap}
  .th-cat{color:#fff;text-align:center;padding:3px 4px;font-size:7pt;font-weight:700;
          letter-spacing:.05em;border:1px solid rgba(255,255,255,.25)}
  .th-size{color:#fff;text-align:center;padding:3px 4px;font-size:8pt;font-weight:700;
           border:1px solid rgba(255,255,255,.2)}
  .th-ml{color:rgba(255,255,255,.75);text-align:center;padding:2px 4px;font-size:6pt;
         font-weight:600;letter-spacing:.06em;border:1px solid rgba(255,255,255,.15)}
  .th-total-hdr{background:#312e81;color:#e0e7ff;text-align:center;padding:3px 4px;
                font-size:6.5pt;font-weight:700;border:1px solid #4338ca;white-space:nowrap}
  .td-label{text-align:left;white-space:nowrap;padding:3px 6px;color:#0f172a;
            border:1px solid #cbd5e1;border-left:3px solid #94a3b8;font-weight:500}
  .td-num{text-align:center;padding:3px 3px;border:1px solid #cbd5e1;color:#1e293b;min-width:28px}
  .td-tot{text-align:center;padding:3px 5px;border:1px solid #6366f1;background:#eef2ff;
          color:#312e81;font-weight:700;font-size:8pt}

  /* ── Separators & specials ── */
  tr.sep td{height:2px!important;background:#cbd5e1!important;padding:0!important;border:none!important}

  /* ── Footer ── */
  .footer{margin-top:12px;display:flex;justify-content:space-between;font-size:7.5pt;color:#475569}
  .sig-line{border-top:1px solid #64748b;padding-top:3px;min-width:140px;text-align:center;font-size:7pt}
  .print-note{font-size:6.5pt;color:#94a3b8;text-align:center;margin-top:6px}

  /* ── Bulk LTR section ── */
  .bulk-section{margin-top:14px;page-break-inside:avoid}
  .bulk-title{font-size:9pt;font-weight:800;color:#1e3a8a;margin-bottom:5px;text-transform:uppercase;letter-spacing:.04em;border-bottom:2px solid #1e3a8a;padding-bottom:3px}
  .bulk-table{border-collapse:collapse;font-size:8pt;min-width:400px}
  .bulk-table th,.bulk-table td{border:1px solid #94a3b8;padding:4px 12px}
  .bulk-table thead th{background:#1e3a8a;color:#fff;font-weight:700;text-align:center}
  .bulk-table thead th.blk-th-label{text-align:left}
  .blk-label{text-align:left;font-weight:500;white-space:nowrap}
  .blk-num{text-align:center;font-variant-numeric:tabular-nums}
  .blk-total td{background:#eff6ff;font-weight:700;border-top:2px solid #1e3a8a}
  ${bw ? 'html,body{-webkit-filter:grayscale(1)!important;filter:grayscale(1)!important}' : ''}
</style>
</head><body>
  <div class="top-bar"></div>

  <div class="letterhead">
    <div class="lh-left">
      <div class="govt-label">Government of Maharashtra</div>
      <div class="govt-label">Excise Department</div>
    </div>
    <div class="lh-center">
      <div><span class="form-badge">FORM F.L.R.- 4 &nbsp; [See Rule 15 (II)]</span></div>
      <div class="form-title">Monthly Statement of Foreign Liquor</div>
      <div class="form-sub">Monthly return of Transactions effected by Vendors, Hotels &amp; Clubs Licensee</div>
    </div>
    <div class="lh-right">
      <div class="govt-label">Month of</div>
      <div class="bar-name-big">${monthName}, ${selYear}</div>
    </div>
  </div>

  <div class="info-box">
    <div class="info-row"><span class="info-label">Licensee Name :</span><span class="info-val">&nbsp;${barName}</span></div>
    <div class="info-row info-right"><span class="info-label">Licence No :</span><span class="info-val">&nbsp;${licNo}</span></div>
    <div class="info-row"><span class="info-label">Address :</span><span class="info-val">&nbsp;${address}</span></div>
    <div class="info-row info-right"><span class="info-label">Financial Year :</span><span class="info-val">&nbsp;${activeBar.financialYear || '—'}</span></div>
  </div>

  <table>
    <thead>
      <tr>${ph1}</tr>
      <tr>${ph2}</tr>
      <tr>${ph3}</tr>
    </thead>
    <tbody>${tbody}</tbody>
  </table>

  <div class="bulk-section">
    <div class="bulk-title">BULK LTR REPORT &mdash; Month of : ${monthName}, ${selYear}</div>
    <table class="bulk-table">
      <thead><tr><th class="blk-th-label">PARTICULARS</th><th>OPENING</th><th>RECEIVED</th><th>SALE</th><th>CLOSING STOCK</th></tr></thead>
      <tbody>${bulkTbody}</tbody>
    </table>
  </div>

  <div class="footer">
    <div class="sig-line">Date &amp; Seal of Licensee</div>
    <div class="print-note">Generated by SPLIQOUR PRO &nbsp;|&nbsp; Printed on ${new Date().toLocaleDateString('en-IN')}</div>
    <div class="sig-line">Signature of Licensee</div>
  </div>

</body></html>`;
            if (_previewMode) { _previewMode = false; showPrintPreview(__msPdfHtml); }
            else { downloadPdf(__msPdfHtml, `Monthly_Statement_FL4_${(barName||'').replace(/[^a-z0-9]/gi,'_')}_${monthName}_${selYear}${bw?'_BW':''}.pdf`); }
        }

        // ── Event wiring ──
        if (msGoBtn)     msGoBtn.addEventListener('click', generateMonthlyStatement);
        if (msExportBtn) msExportBtn.addEventListener('click', exportMsCsv);
        if (msPdfBtn)    msPdfBtn.addEventListener('click', () => printMsStatement(false));
        if (msBwPdfBtn)  msBwPdfBtn.addEventListener('click', () => printMsStatement(true));
        document.getElementById('msPreviewBtn')?.addEventListener('click', () => { _previewMode = true; printMsStatement(false); });

        // Auto-generate when sub-view becomes visible
        const msPanel = document.getElementById('sub-monthly-statement');
        if (msPanel) {
            const msObserver = new MutationObserver((mutations) => {
                mutations.forEach(m => {
                    if (m.attributeName === 'class' && msPanel.classList.contains('active') && !msGenerated) {
                        generateMonthlyStatement();
                    }
                });
            });
            msObserver.observe(msPanel, { attributes: true, attributeFilter: ['class'] });
        }
    }

    // ═══════════════════════════════════════════════════════
    //  11 MONTHS STATEMENT — Purchase/Sales bottle qty by category
    //  Default selected months: April to February
    // ═══════════════════════════════════════════════════════
    {
        const m11FySelect   = document.getElementById('m11FySelect');
        const m11MonthGrid  = document.getElementById('m11MonthGrid');
        const m11GoBtn      = document.getElementById('m11GoBtn');
        const m11PrintBtn   = document.getElementById('m11PrintBtn');
        const m11PurchaseWrap = document.getElementById('m11PurchaseWrap');
        const m11SalesWrap    = document.getElementById('m11SalesWrap');
        const m11PurchaseHead = document.getElementById('m11PurchaseHead');
        const m11PurchaseBody = document.getElementById('m11PurchaseBody');
        const m11SalesHead    = document.getElementById('m11SalesHead');
        const m11SalesBody    = document.getElementById('m11SalesBody');
        const m11Empty      = document.getElementById('m11Empty');

        let m11Generated = false;

        const M11_MONTHS = [
            { idx: 3, label: 'Apr' }, { idx: 4, label: 'May' }, { idx: 5, label: 'Jun' },
            { idx: 6, label: 'Jul' }, { idx: 7, label: 'Aug' }, { idx: 8, label: 'Sep' },
            { idx: 9, label: 'Oct' }, { idx: 10, label: 'Nov' }, { idx: 11, label: 'Dec' },
            { idx: 0, label: 'Jan' }, { idx: 1, label: 'Feb' }, { idx: 2, label: 'Mar' },
        ];

        const M11_CATEGORY_ORDER = ['SPIRITS', 'WINE', 'FERMENTED BEER', 'MILD BEER', 'MML'];

        function m11NormalizeCategory(cat) {
            const raw = String(cat || '').trim().toUpperCase();
            if (!raw) return 'UNCATEGORIZED';
            if (raw === 'WINES') return 'WINE';
            if (raw === 'SPIRIT') return 'SPIRITS';
            return raw;
        }

        function m11CategoryDisplay(cat) {
            const c = m11NormalizeCategory(cat);
            if (c === 'SPIRITS') return 'Spirits';
            if (c === 'WINE') return 'Wine';
            if (c === 'FERMENTED BEER') return 'Fermented Beer';
            if (c === 'MILD BEER') return 'Mild Beer';
            if (c === 'MML') return 'MML';
            return 'Uncategorized';
        }

        function m11NormBrand(s) {
            return (s || '').trim().replace(/\./g, '').replace(/\s+/g, ' ').trim().toUpperCase();
        }

        function m11NormSize(s) {
            return String(s || '').trim().replace(/\s+/g, ' ').toUpperCase();
        }

        function m11SizeSortValue(sizeLabel) {
            const s = String(sizeLabel || '').toUpperCase();
            const ml = s.match(/(\d+)\s*ML/);
            if (ml) return Number(ml[1]);
            const ltr = s.match(/(\d+)\s*LTR/);
            if (ltr) return Number(ltr[1]) * 1000;
            return -1;
        }

        function m11DefaultFyStart() {
            const now = new Date();
            if (activeBar.financialYear) {
                const m = String(activeBar.financialYear).match(/(\d{4})/);
                if (m) return Number(m[1]);
            }
            return now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
        }

        function m11SelectedMonths() {
            const order = M11_MONTHS.map(m => m.idx);
            const picked = [];
            for (const mi of order) {
                const cb = document.getElementById(`m11Mon_${mi}`);
                if (cb && cb.checked) picked.push(mi);
            }
            return picked;
        }

        function m11IsInFyMonth(dateStr, fyStartYear, monthIdx) {
            const d = new Date(dateStr || '');
            if (!Number.isFinite(d.getTime())) return false;
            const targetYear = monthIdx >= 3 ? fyStartYear : fyStartYear + 1;
            return d.getFullYear() === targetYear && d.getMonth() === monthIdx;
        }

        function m11Fmt(n) {
            const num = Number(n) || 0;
            return num === 0 ? '-' : num.toLocaleString('en-IN');
        }

        function m11Esc(str) {
            return String(str || '')
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;');
        }

        function initM11Controls() {
            if (m11FySelect) {
                const base = m11DefaultFyStart();
                m11FySelect.innerHTML = '';
                for (let y = base + 1; y >= base - 2; y--) {
                    const fyStart = y - 1;
                    const opt = document.createElement('option');
                    opt.value = String(fyStart);
                    opt.textContent = `FY ${fyStart}-${String(y).slice(2)}`;
                    m11FySelect.appendChild(opt);
                }
                m11FySelect.value = String(base);
            }

            if (m11MonthGrid) {
                m11MonthGrid.innerHTML = M11_MONTHS.map((m) => {
                    const checked = m.idx === 2 ? '' : 'checked'; // March off by default
                    return `<label class="m11-month-chip"><input type="checkbox" id="m11Mon_${m.idx}" ${checked}>${m.label}</label>`;
                }).join('');
            }
        }

        function buildM11Header(headEl, sizeLabels) {
            if (!headEl) return;
            const labels = sizeLabels.length ? sizeLabels : ['N/A'];
            headEl.innerHTML = `
                <tr class="ms-header-cat">
                    <th class="ms-th ms-th-frozen m11-th-category">Category</th>
                    ${labels.map(l => `<th class="ms-th ms-th-size">${m11Esc(l)}</th>`).join('')}
                </tr>
            `;
        }

        async function generate11MonthsStatement() {
            if (!activeBar.barName) return;
            const selectedMonths = m11SelectedMonths();
            if (!selectedMonths.length) {
                if (m11PurchaseWrap) m11PurchaseWrap.classList.add('hidden');
                if (m11SalesWrap) m11SalesWrap.classList.add('hidden');
                if (m11Empty) m11Empty.classList.remove('hidden');
                if (m11PurchaseHead) m11PurchaseHead.innerHTML = '';
                if (m11SalesHead) m11SalesHead.innerHTML = '';
                if (m11PurchaseBody) m11PurchaseBody.innerHTML = '';
                if (m11SalesBody) m11SalesBody.innerHTML = '';
                return;
            }

            const fyStartYear = Number(m11FySelect?.value || m11DefaultFyStart());
            const params = { barName: activeBar.barName, financialYear: activeBar.financialYear || '' };

            const [tpResult, salesItems, productsResult] = await Promise.all([
                window.electronAPI.getTps(params),
                getSalesData(params),
                window.electronAPI.getProducts(params),
            ]);

            const byCode = new Map();
            const byBrandSize = new Map();
            if (productsResult.success && productsResult.products) {
                for (const p of productsResult.products) {
                    const cat = m11NormalizeCategory(p.category || '');
                    if (!M11_CATEGORY_ORDER.includes(cat)) continue;
                    const sizeNorm = m11NormSize(p.size || '');
                    const codeKey = String(p.code || '').trim().toUpperCase();
                    if (codeKey) byCode.set(codeKey, { cat, size: sizeNorm });
                    const bsKey = `${m11NormBrand(p.brandName)}|${m11NormSize(p.size)}`;
                    byBrandSize.set(bsKey, { cat, size: sizeNorm });
                }
            }

            const resolveCategory = ({ code, category, brandName, brand, size }) => {
                const direct = m11NormalizeCategory(category || '');
                if (M11_CATEGORY_ORDER.includes(direct)) return direct;
                const codeKey = String(code || '').trim().toUpperCase();
                if (codeKey && byCode.has(codeKey)) return byCode.get(codeKey).cat;
                const bsKey = `${m11NormBrand(brandName || brand)}|${m11NormSize(size)}`;
                if (byBrandSize.has(bsKey)) return byBrandSize.get(bsKey).cat;
                return 'UNCATEGORIZED';
            };

            const resolveSize = ({ code, size, brandName, brand }) => {
                const directSize = m11NormSize(size || '');
                if (directSize) return directSize;
                const codeKey = String(code || '').trim().toUpperCase();
                if (codeKey && byCode.has(codeKey)) return byCode.get(codeKey).size || 'N/A';
                const bsKey = `${m11NormBrand(brandName || brand)}|${m11NormSize(size)}`;
                if (byBrandSize.has(bsKey)) return byBrandSize.get(bsKey).size || 'N/A';
                return 'N/A';
            };

            const purchase = {};
            const sales = {};
            M11_CATEGORY_ORDER.forEach(cat => {
                purchase[cat] = {};
                sales[cat] = {};
            });
            const sizeSet = new Set();

            const ensureSizeBucket = (cat, sizeLabel) => {
                sizeSet.add(sizeLabel);
                if (!purchase[cat][sizeLabel]) {
                    purchase[cat][sizeLabel] = 0;
                }
                if (!sales[cat][sizeLabel]) {
                    sales[cat][sizeLabel] = 0;
                }
            };

            const inSelectedMonths = (dateStr) => selectedMonths.some(mi => m11IsInFyMonth(dateStr, fyStartYear, mi));

            if (tpResult.success && tpResult.tps) {
                for (const tp of tpResult.tps) {
                    const tpDate = tp.tpDate || '';
                    for (const item of (tp.items || [])) {
                        const cat = resolveCategory(item);
                        if (!M11_CATEGORY_ORDER.includes(cat)) continue;
                        const sizeLabel = resolveSize(item);
                        ensureSizeBucket(cat, sizeLabel);
                        if (inSelectedMonths(tpDate)) {
                            purchase[cat][sizeLabel] += Number(item.totalBtl || 0);
                        }
                    }
                }
            }

            for (const s of (salesItems || [])) {
                const cat = resolveCategory(s);
                if (!M11_CATEGORY_ORDER.includes(cat)) continue;
                const sizeLabel = resolveSize(s);
                ensureSizeBucket(cat, sizeLabel);
                if (inSelectedMonths(s.billDate)) {
                    sales[cat][sizeLabel] += Number(s.totalBtl || 0);
                }
            }

            const sizeLabels = [...sizeSet].sort((a, b) => {
                const delta = m11SizeSortValue(b) - m11SizeSortValue(a);
                if (delta !== 0) return delta;
                return String(a).localeCompare(String(b), 'en', { sensitivity: 'base' });
            });

            buildM11Header(m11PurchaseHead, sizeLabels);
            buildM11Header(m11SalesHead, sizeLabels);
            if (m11PurchaseBody) m11PurchaseBody.innerHTML = '';
            if (m11SalesBody) m11SalesBody.innerHTML = '';

            const grandPurchase = Object.fromEntries(sizeLabels.map(s => [s, 0]));
            const grandSales = Object.fromEntries(sizeLabels.map(s => [s, 0]));

            M11_CATEGORY_ORDER.forEach((cat) => {
                const purchaseTr = document.createElement('tr');
                let purchaseHtml = `<td class="ms-td ms-td-frozen m11-td-category">${m11CategoryDisplay(cat)}</td>`;
                sizeLabels.forEach(sizeLabel => {
                    const v = purchase[cat][sizeLabel] || 0;
                    grandPurchase[sizeLabel] += v;
                    purchaseHtml += `<td class="ms-td">${m11Fmt(v)}</td>`;
                });
                purchaseTr.innerHTML = purchaseHtml;
                m11PurchaseBody?.appendChild(purchaseTr);

                const salesTr = document.createElement('tr');
                let salesHtml = `<td class="ms-td ms-td-frozen m11-td-category">${m11CategoryDisplay(cat)}</td>`;
                sizeLabels.forEach(sizeLabel => {
                    const v = sales[cat][sizeLabel] || 0;
                    grandSales[sizeLabel] += v;
                    salesHtml += `<td class="ms-td">${m11Fmt(v)}</td>`;
                });
                salesTr.innerHTML = salesHtml;
                m11SalesBody?.appendChild(salesTr);
            });

            const totalPurchaseTr = document.createElement('tr');
            totalPurchaseTr.className = 'ms-row-total';
            let totalPurchaseHtml = '<td class="ms-td ms-td-frozen m11-td-category"><strong>Total</strong></td>';
            sizeLabels.forEach(sizeLabel => { totalPurchaseHtml += `<td class="ms-td"><strong>${m11Fmt(grandPurchase[sizeLabel])}</strong></td>`; });
            totalPurchaseTr.innerHTML = totalPurchaseHtml;
            m11PurchaseBody?.appendChild(totalPurchaseTr);

            const totalSalesTr = document.createElement('tr');
            totalSalesTr.className = 'ms-row-total';
            let totalSalesHtml = '<td class="ms-td ms-td-frozen m11-td-category"><strong>Total</strong></td>';
            sizeLabels.forEach(sizeLabel => { totalSalesHtml += `<td class="ms-td"><strong>${m11Fmt(grandSales[sizeLabel])}</strong></td>`; });
            totalSalesTr.innerHTML = totalSalesHtml;
            m11SalesBody?.appendChild(totalSalesTr);

            if (m11Empty) m11Empty.classList.add('hidden');
            if (m11PurchaseWrap) m11PurchaseWrap.classList.remove('hidden');
            if (m11SalesWrap) m11SalesWrap.classList.remove('hidden');
            m11Generated = true;
        }

        async function print11MonthsStatement() {
            if (!activeBar.barName) return;

            if (!m11PurchaseBody?.children.length && !m11SalesBody?.children.length) {
                await generate11MonthsStatement();
            }
            if (!m11PurchaseBody?.children.length && !m11SalesBody?.children.length) return;

            const fyLabel = m11FySelect?.selectedOptions?.[0]?.textContent || activeBar.financialYear || 'FY';
            const selectedMonthLabels = m11SelectedMonths()
                .map(mi => M11_MONTHS.find(m => m.idx === mi)?.label)
                .filter(Boolean)
                .join(', ');

            const barName = activeBar.barName || '—';
            const licNo = activeBar.licenseNo || activeBar.licNo || '—';
            const barAddress = [activeBar.address, activeBar.area, activeBar.city].filter(Boolean).join(', ') || '—';
            const generatedOn = new Date().toLocaleString('en-IN');

            const purchaseHead = m11PurchaseHead ? m11PurchaseHead.innerHTML : '';
            const purchaseBody = m11PurchaseBody ? m11PurchaseBody.innerHTML : '';
            const salesHead = m11SalesHead ? m11SalesHead.innerHTML : '';
            const salesBody = m11SalesBody ? m11SalesBody.innerHTML : '';

            printWithIframe(`<!DOCTYPE html><html><head><title>11 Months Statement — ${m11Esc(barName)}</title>
                <style>
                    body{font-family:'Segoe UI',Arial,sans-serif;padding:10px;color:#111;font-size:10pt}
                    .head{border:1px solid #94a3b8;border-radius:6px;padding:8px 10px;background:#f8fafc;margin-bottom:10px}
                    .head h2{margin:0 0 6px;font-size:14pt}
                    .meta{display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap;font-size:9pt}
                    .meta .lbl{font-weight:700}
                    h3{margin:10px 0 6px;font-size:11pt}
                    table{border-collapse:collapse;width:100%;margin-bottom:12px}
                    th,td{border:1px solid #8b8b8b;padding:3px 5px;font-size:8.5pt;text-align:center;white-space:nowrap}
                    th:first-child, td:first-child{text-align:left}
                    th{background:#eef2ff;font-weight:700}
                    tr.ms-row-total td{font-weight:700;background:#f8fafc}
                    @page{size:legal landscape;margin:8mm}
                </style>
            </head><body>
                <div class="head">
                    <h2>11 Months Statement (Bottle Qty)</h2>
                    <div class="meta"><div><span class="lbl">Bar Name:</span> ${m11Esc(barName)}</div><div><span class="lbl">Lic No:</span> ${m11Esc(licNo)}</div><div><span class="lbl">${m11Esc(fyLabel)}</span></div></div>
                    <div class="meta"><div><span class="lbl">Address:</span> ${m11Esc(barAddress)}</div><div><span class="lbl">Months:</span> ${m11Esc(selectedMonthLabels || '—')}</div><div><span class="lbl">Generated:</span> ${m11Esc(generatedOn)}</div></div>
                </div>

                <h3>Purchase</h3>
                <table><thead>${purchaseHead}</thead><tbody>${purchaseBody}</tbody></table>

                <h3>Sales</h3>
                <table><thead>${salesHead}</thead><tbody>${salesBody}</tbody></table>

            </body></html>`);
        }

        initM11Controls();
        if (m11GoBtn) m11GoBtn.addEventListener('click', generate11MonthsStatement);
        if (m11PrintBtn) m11PrintBtn.addEventListener('click', print11MonthsStatement);
        document.getElementById('m11PreviewBtn')?.addEventListener('click', () => { _previewMode = true; print11MonthsStatement(); });

        const m11Panel = document.getElementById('sub-eleven-months-statement');
        if (m11Panel) {
            const m11Observer = new MutationObserver((mutations) => {
                mutations.forEach(m => {
                    if (m.attributeName === 'class' && m11Panel.classList.contains('active') && !m11Generated) {
                        generate11MonthsStatement();
                    }
                });
            });
            m11Observer.observe(m11Panel, { attributes: true, attributeFilter: ['class'] });
        }
    }

    // ═══════════════════════════════════════════════════════
    //  MML MONTHLY STATEMENT — MML-only Form F.L.R.-4
    //  Same pivot layout but filtered to MML category only
    // ═══════════════════════════════════════════════════════
    {
        const mmlMonthSelect = document.getElementById('mmlMsMonthSelect');
        const mmlYearSelect  = document.getElementById('mmlMsYearSelect');
        const mmlGoBtn       = document.getElementById('mmlMsGoBtn');
        const mmlExportBtn   = document.getElementById('mmlMsExportBtn');
        const mmlMsPdfBtn    = document.getElementById('mmlMsPdfBtn');
        const mmlMsBwPdfBtn  = document.getElementById('mmlMsBwPdfBtn');
        const mmlFormHeader  = document.getElementById('mmlMsFormHeader');
        const mmlTableHead   = document.getElementById('mmlMsTableHead');
        const mmlTableBody   = document.getElementById('mmlMsTableBody');
        const mmlTable       = document.getElementById('mmlMsTable');
        const mmlEmpty       = document.getElementById('mmlMsEmpty');
        const mmlBulkSection = document.getElementById('mmlMsBulkSection');
        const mmlBulkMonth   = document.getElementById('mmlMsBulkMonth');
        const mmlBulkBody    = document.getElementById('mmlMsBulkBody');
        const mmlBulkFoot    = document.getElementById('mmlMsBulkFoot');

        let mmlGenerated = false;
        let mmlLastData  = null;

        // FY defaults
        const mmlFyStartYear = (() => {
            const now = new Date();
            return now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
        })();

        if (mmlYearSelect) {
            for (let y = mmlFyStartYear + 1; y >= mmlFyStartYear - 2; y--) {
                const opt = document.createElement('option');
                opt.value = y;
                opt.textContent = y;
                mmlYearSelect.appendChild(opt);
            }
        }
        {
            const now = new Date();
            if (mmlMonthSelect) mmlMonthSelect.value = now.getMonth();
            if (mmlYearSelect)  mmlYearSelect.value  = now.getFullYear();
        }

        const MML_ROW_DEFS = [
            { key: 'opening',      label: 'Opening Bal. of the begin. of the mon.' },
            { key: 'received',     label: 'Received during the month' },
            { key: 'total',        label: 'Total' },
            { key: 'cumulReceipt', label: 'Rece from 1st Apr to end of the month' },
            { key: 'sold',         label: 'Sold Per. Hol during month' },
            { key: 'cumulSale',    label: 'Sold Per. Hold. 1st Apr. to end of' },
            { key: 'breakage',     label: 'Breakage and wastage dur. the curr month' },
            { key: 'closing',      label: 'Clos Bal. of the end of the month' },
        ];

        const MML_MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];

        function mmlExtractML(sizeStr) {
            const m = (sizeStr || '').match(/^(\d+)\s*ML/i);
            if (m) return parseInt(m[1]);
            const l = (sizeStr || '').match(/^(\d+)\s*Ltr/i);
            if (l) return parseInt(l[1]) * 1000;
            return 0;
        }

        function mmlGetMonthRange(year, monthIdx) {
            const start = new Date(year, monthIdx, 1);
            const end   = new Date(year, monthIdx + 1, 0);
            const fmt = d => d.toISOString().slice(0, 10);
            return { start: fmt(start), end: fmt(end) };
        }

        function mmlGetFyStart(year, monthIdx) {
            const fyStartYear = monthIdx >= 3 ? year : year - 1;
            return `${fyStartYear}-04-01`;
        }

        function mmlColKey(ml) { return ml === 0 ? '_total' : 'MML_' + ml; }
        function mmlFmtBulk(n) { return (n || 0).toFixed(3); }

        async function generateMmlMonthlyStatement() {
          try {
            if (!activeBar.barName) { console.warn('MML-MS: no activeBar'); return; }
            const params = { barName: activeBar.barName, financialYear: activeBar.financialYear || '' };

            const selMonth = parseInt(mmlMonthSelect ? mmlMonthSelect.value : new Date().getMonth());
            const selYear  = parseInt(mmlYearSelect  ? mmlYearSelect.value  : new Date().getFullYear());
            const monthName = MML_MONTH_NAMES[selMonth];
            const fyStartDate = mmlGetFyStart(selYear, selMonth);

            // Form header
            if (mmlFormHeader) {
                mmlFormHeader.innerHTML = `
                    <div class="ms-fh-row">
                        <span class="ms-fh-form">MML MONTHLY STATEMENT &nbsp; [Form F.L.R.- 4 Style]</span>
                    </div>
                    <div class="ms-fh-row ms-fh-subtitle">
                        Monthly return of Transactions of MML (Maharashtra Made Liquor) effected by Licensee.
                    </div>
                    <div class="ms-fh-row">
                        <span>Name of the Licensee : <strong>${activeBar.barName || '—'}</strong></span>
                    </div>
                    <div class="ms-fh-row">
                        <span>Address : <strong>${[activeBar.address, activeBar.area, activeBar.city].filter(Boolean).join(', ') || '—'}</strong></span>
                        <span class="ms-fh-right">Licence No: <strong>${activeBar.licenseNo || '—'}</strong></span>
                    </div>
                    <div class="ms-fh-row">
                        <span></span>
                        <span class="ms-fh-right">Month of : <strong>${monthName}, &nbsp; ${selYear}</strong></span>
                    </div>
                `;
            }

            // Fetch data
            const [osResult, tpResult, salesItems, productsResult] = await Promise.all([
                window.electronAPI.getOpeningStock(params),
                window.electronAPI.getTps(params),
                getSalesData(params),
                window.electronAPI.getProducts(params),
            ]);

            // Build product map — MML only
            const productMap = new Map();
            if (productsResult.success && productsResult.products) {
                for (const p of productsResult.products) {
                    if (p.category === 'MML') {
                        productMap.set(p.code, {
                            category: p.category,
                            size: p.size || '',
                            ml: mmlExtractML(p.size)
                        });
                    }
                }
            }

            // Collect all unique sizes from MML products (flat — no sub-category)
            const allSizes = new Set();
            for (const [, p] of productMap) {
                if (p.ml > 0) allSizes.add(p.ml);
            }
            const sortedSizes = [...allSizes].sort((a, b) => b - a); // descending

            // Build flat column list: one column per size + total
            const columns = [];
            for (const ml of sortedSizes) {
                columns.push({ ml, label: ml + ' ml' });
            }
            columns.push({ ml: 0, label: 'TOTAL' }); // total column

            function isMmlItem(code, category) {
                if (productMap.has(code)) return true;
                return (category || '').toUpperCase() === 'MML';
            }

            function getMl(code, size) {
                const prod = productMap.get(code);
                const sz = size || (prod ? prod.size : '');
                return mmlExtractML(sz);
            }

            const curRange = mmlGetMonthRange(selYear, selMonth);

            function emptyRow() {
                const r = {};
                for (const c of columns) r[mmlColKey(c.ml)] = 0;
                return r;
            }

            const data = {};
            for (const rd of MML_ROW_DEFS) data[rd.key] = emptyRow();

            // 1. FY Opening stock (MML only)
            const fyOpening = emptyRow();
            const osData = osResult.success ? osResult.data : null;
            if (osData && osData.entries) {
                for (const entry of osData.entries) {
                    if (!isMmlItem(entry.code, entry.category)) continue;
                    const ml = getMl(entry.code, entry.size || entry.sizeLabel);
                    const ck = mmlColKey(ml);
                    if (ck in fyOpening) fyOpening[ck] += (entry.totalBtl || 0);
                    fyOpening['_total'] += (entry.totalBtl || 0);
                }
            }

            // 2. TPs (MML items only)
            const prevRange = { start: fyStartDate, end: (() => {
                const d = new Date(selYear, selMonth, 0);
                return d.toISOString().slice(0, 10);
            })() };

            const prevReceipts = emptyRow();
            const curReceipts  = emptyRow();
            if (tpResult.success && tpResult.tps) {
                for (const tp of tpResult.tps) {
                    const dt = tp.tpDate || tp.receivedDate || '';
                    for (const item of (tp.items || [])) {
                        if (!isMmlItem(item.code, item.category)) continue;
                        const ml = getMl(item.code, item.size || item.sizeLabel);
                        const ck = mmlColKey(ml);
                        const btl = item.totalBtl || 0;
                        if (dt >= curRange.start && dt <= curRange.end) {
                            if (ck in curReceipts) curReceipts[ck] += btl;
                            curReceipts['_total'] += btl;
                        } else if (dt >= prevRange.start && dt <= prevRange.end) {
                            if (ck in prevReceipts) prevReceipts[ck] += btl;
                            prevReceipts['_total'] += btl;
                        }
                    }
                }
            }

            // 3. Sales (MML items only)
            const prevSales = emptyRow();
            const curSales  = emptyRow();
            for (const item of salesItems) {
                if (!isMmlItem(item.code, item.category)) continue;
                const ml = getMl(item.code, item.size);
                const ck = mmlColKey(ml);
                const btl = item.totalBtl || 0;
                const dt = item.billDate || '';
                if (dt >= curRange.start && dt <= curRange.end) {
                    if (ck in curSales) curSales[ck] += btl;
                    curSales['_total'] += btl;
                } else if (dt >= prevRange.start && dt <= prevRange.end) {
                    if (ck in prevSales) prevSales[ck] += btl;
                    prevSales['_total'] += btl;
                }
            }

            // Compute rows
            const allKeys = Object.keys(data['opening']);
            for (const k of allKeys) {
                data['opening'][k] = (fyOpening[k] || 0) + (prevReceipts[k] || 0) - (prevSales[k] || 0);
                data['received'][k] = curReceipts[k] || 0;
                data['total'][k] = data['opening'][k] + data['received'][k];
                data['cumulReceipt'][k] = (prevReceipts[k] || 0) + (curReceipts[k] || 0);
                data['sold'][k] = curSales[k] || 0;
                data['cumulSale'][k] = (prevSales[k] || 0) + (curSales[k] || 0);
                data['breakage'][k] = 0;
                data['closing'][k] = data['total'][k] - data['sold'][k] - data['breakage'][k];
            }

            // ── Prune size columns that are zero across ALL rows ──
            const filteredSizes = sortedSizes.filter(ml => {
                const ck = mmlColKey(ml);
                return MML_ROW_DEFS.some(rd => (data[rd.key][ck] || 0) !== 0);
            });
            sortedSizes.splice(0, sortedSizes.length, ...filteredSizes);
            // Rebuild columns to match pruned sizes
            columns.length = 0;
            for (const ml of sortedSizes) {
                columns.push({ ml, label: ml + ' ml' });
            }
            columns.push({ ml: 0, label: 'TOTAL' });

            mmlLastData = { data, columns, sortedSizes, selMonth, selYear, monthName };
            mmlGenerated = true;
            renderMmlTable(data, columns, sortedSizes);
            renderMmlBulkReport(data, sortedSizes, monthName, selYear);
          } catch (err) {
            console.error('MML-MS generateMmlMonthlyStatement ERROR:', err);
          }
        }

        /* Render MML table — flat sizes, no sub-category grouping */
        function renderMmlTable(data, columns, sortedSizes) {
            if (!mmlTableHead || !mmlTableBody) return;

            if (sortedSizes.length === 0) {
                if (mmlEmpty) mmlEmpty.classList.remove('hidden');
                if (mmlTable) mmlTable.style.display = 'none';
                if (mmlBulkSection) mmlBulkSection.style.display = 'none';
                return;
            }
            if (mmlEmpty) mmlEmpty.classList.add('hidden');
            if (mmlTable) mmlTable.style.display = '';
            if (mmlBulkSection) mmlBulkSection.style.display = '';

            // Header row 1: single "MML" spanning all size columns
            const sizeCount = sortedSizes.length;
            let h1 = `<th class="ms-th ms-th-frozen ms-th-particulars" rowspan="3">PARTICULARS</th>`;
            h1 += `<th class="ms-th ms-th-group ms-th-spirits" colspan="${sizeCount}">-- MML (MAHARASHTRA MADE LIQUOR) --</th>`;
            h1 += `<th class="ms-th ms-th-group ms-th-total-hdr" rowspan="3">BOTTLE/<br>BULK(LTR)</th>`;

            // Header row 2: size values
            let h2 = '';
            for (const ml of sortedSizes) {
                h2 += `<th class="ms-th ms-th-size">${ml}</th>`;
            }

            // Header row 3: "ml"
            let h3 = '';
            for (let i = 0; i < sizeCount; i++) {
                h3 += '<th class="ms-th ms-th-ml">ml</th>';
            }

            mmlTableHead.innerHTML = `
                <tr class="ms-header-cat">${h1}</tr>
                <tr class="ms-header-size">${h2}</tr>
                <tr class="ms-header-ml">${h3}</tr>
            `;

            // Body rows
            mmlTableBody.innerHTML = '';
            for (const rd of MML_ROW_DEFS) {
                const tr = document.createElement('tr');
                const isTotal = rd.key === 'total';
                const isClosing = rd.key === 'closing';
                const isCumul = rd.key === 'cumulReceipt' || rd.key === 'cumulSale';
                tr.className = 'ms-data-row' +
                    (isTotal ? ' ms-row-total' : '') +
                    (isClosing ? ' ms-row-closing' : '') +
                    (isCumul ? ' ms-row-cumul' : '');

                let cells = `<td class="ms-td ms-td-frozen ms-td-label">${rd.label}</td>`;
                for (const ml of sortedSizes) {
                    const ck = mmlColKey(ml);
                    const val = data[rd.key][ck] || 0;
                    const cls = val === 0 ? 'ms-td ms-td-num ms-td-zero' :
                                val < 0  ? 'ms-td ms-td-num ms-td-neg' :
                                'ms-td ms-td-num';
                    cells += `<td class="${cls}">${val || ''}</td>`;
                }
                const totalVal = data[rd.key]['_total'] || 0;
                const totalCls = totalVal === 0 ? 'ms-td ms-td-num ms-td-zero' :
                                 totalVal < 0  ? 'ms-td ms-td-num ms-td-neg' :
                                 'ms-td ms-td-num ms-td-total-val';
                cells += `<td class="${totalCls}">${totalVal || ''}</td>`;
                tr.innerHTML = cells;
                mmlTableBody.appendChild(tr);

                if (rd.key === 'total' || rd.key === 'cumulSale') {
                    const sep = document.createElement('tr');
                    sep.className = 'ms-sep-row';
                    sep.innerHTML = `<td colspan="${2 + sizeCount}" class="ms-td-sep"></td>`;
                    mmlTableBody.appendChild(sep);
                }
            }
        }

        /* Bulk LTR Report for MML — single row for MML total */
        function renderMmlBulkReport(data, sortedSizes, monthName, year) {
            if (!mmlBulkBody || !mmlBulkFoot) return;
            if (mmlBulkMonth) mmlBulkMonth.textContent = `${monthName}, ${year}`;
            mmlBulkBody.innerHTML = '';

            let bO = 0, bR = 0, bS = 0, bC = 0;
            for (const ml of sortedSizes) {
                const ck = mmlColKey(ml);
                bO += (data['opening'][ck] || 0)  * ml / 1000;
                bR += (data['received'][ck] || 0) * ml / 1000;
                bS += (data['sold'][ck] || 0)     * ml / 1000;
                bC += (data['closing'][ck] || 0)  * ml / 1000;
            }

            mmlBulkBody.innerHTML = `
                <tr>
                    <td class="ms-btd ms-btd-label">MML</td>
                    <td class="ms-btd ms-btd-num">${mmlFmtBulk(bO)}</td>
                    <td class="ms-btd ms-btd-num">${mmlFmtBulk(bR)}</td>
                    <td class="ms-btd ms-btd-num">${mmlFmtBulk(bS)}</td>
                    <td class="ms-btd ms-btd-num">${mmlFmtBulk(bC)}</td>
                </tr>
            `;

            mmlBulkFoot.innerHTML = `
                <tr class="ms-bulk-total">
                    <td class="ms-btd ms-btd-label"><strong>TOTAL :</strong></td>
                    <td class="ms-btd ms-btd-num"><strong>${mmlFmtBulk(bO)}</strong></td>
                    <td class="ms-btd ms-btd-num"><strong>${mmlFmtBulk(bR)}</strong></td>
                    <td class="ms-btd ms-btd-num"><strong>${mmlFmtBulk(bS)}</strong></td>
                    <td class="ms-btd ms-btd-num"><strong>${mmlFmtBulk(bC)}</strong></td>
                </tr>
            `;
        }

        /* Export CSV for MML */
        function exportMmlCsv() {
            if (!mmlLastData) return;
            const { data, sortedSizes, monthName, selYear } = mmlLastData;
            const csvRows = [];
            let hdr = 'Particulars';
            for (const ml of sortedSizes) { hdr += `,${ml}ml`; }
            hdr += ',TOTAL';
            csvRows.push(hdr);
            for (const rd of MML_ROW_DEFS) {
                let line = `"${rd.label}"`;
                for (const ml of sortedSizes) { line += `,${data[rd.key][mmlColKey(ml)] || 0}`; }
                line += `,${data[rd.key]['_total'] || 0}`;
                csvRows.push(line);
            }
            const blob = new Blob(['\uFEFF' + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `MML_FLR4_${activeBar.barName || 'Bar'}_${monthName}_${selYear}.csv`;
            a.click();
            URL.revokeObjectURL(a.href);
        }

        /* Print MML statement */
        function printMmlStatement(bw = false) {
            if (!mmlLastData) return;
            const { data, sortedSizes, monthName, selYear } = mmlLastData;
            const barName  = activeBar.barName || '';
            const address  = [activeBar.address, activeBar.area, activeBar.city].filter(Boolean).join(', ') || '—';
            const licNo    = activeBar.licenseNo || '—';
            const sizeCnt  = sortedSizes.length;

            // ── Header rows ──
            let ph1 = '<th rowspan="3" class="th-part">PARTICULARS</th>';
            ph1 += `<th colspan="${sizeCnt}" class="th-cat">MML (MAHARASHTRA MADE LIQUOR)</th>`;
            ph1 += '<th rowspan="3" class="th-total-hdr">BOTTLE /<br>BULK<br>(LTR)</th>';

            let ph2 = '', ph3 = '';
            for (const ml of sortedSizes) {
                ph2 += `<th class="th-size">${ml}</th>`;
                ph3 += `<th class="th-ml">ml</th>`;
            }

            // ── Body rows ──
            let tbody = '';
            const accentRows = bw
                ? { opening:'#f5f5f5', received:'#f0f0f0', total:'#e5e5e5', cumulReceipt:'#ebebeb',
                    sold:'#f9f9f9', cumulSale:'#f0f0f0', breakage:'#f5f5f5', closing:'#e8e8e8' }
                : { opening:'#f0f9ff', received:'#f0fdf4', total:'#dbeafe', cumulReceipt:'#e0f2fe',
                    sold:'#fff7ed', cumulSale:'#fef3c7', breakage:'#fee2e2', closing:'#ede9fe' };
            const boldRows = new Set(['total','closing']);
            for (const rd of MML_ROW_DEFS) {
                const bg   = accentRows[rd.key] || '#ffffff';
                const bold = boldRows.has(rd.key) ? 'font-weight:700;' : '';
                const fsize = boldRows.has(rd.key) ? '8pt' : '7.5pt';
                tbody += `<tr style="background:${bg};${bold}">`;
                tbody += `<td class="td-label" style="font-size:${fsize}">${rd.label}</td>`;
                for (const ml of sortedSizes) {
                    const val = data[rd.key][mmlColKey(ml)] || 0;
                    tbody += `<td class="td-num" style="font-size:${fsize}">${val || ''}</td>`;
                }
                const tot = data[rd.key]['_total'] || 0;
                tbody += `<td class="td-tot" style="font-size:${fsize}">${tot || ''}</td></tr>`;
                if (rd.key === 'total' || rd.key === 'cumulSale') {
                    const totalCols = 1 + sizeCnt + 1;
                    tbody += `<tr class="sep"><td colspan="${totalCols}"></td></tr>`;
                }
            }

            // ── MML Bulk LTR computation ──
            const fmtB = n => (n || 0).toFixed(3);
            let mmlBulkO = 0, mmlBulkR = 0, mmlBulkS = 0, mmlBulkC = 0;
            for (const ml of sortedSizes) {
                const ck = mmlColKey(ml);
                mmlBulkO += (data['opening'][ck]  || 0) * ml / 1000;
                mmlBulkR += (data['received'][ck] || 0) * ml / 1000;
                mmlBulkS += (data['sold'][ck]     || 0) * ml / 1000;
                mmlBulkC += (data['closing'][ck]  || 0) * ml / 1000;
            }

            const __mmlPdfHtml = `<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>MML F.L.R.-4 — ${barName} — ${monthName} ${selYear}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  html,body,*{-webkit-print-color-adjust:exact;print-color-adjust:exact;color-adjust:exact}
  body{font-family:'Segoe UI',Arial,Helvetica,sans-serif;font-size:8.5pt;color:#111;background:#fff;
       padding:8mm 8mm 6mm}
  @page{size:A4 landscape;margin:8mm 6mm}
  @media print{body{padding:0}}

  /* ── Top accent bar ── */
  .top-bar{height:5px;background:linear-gradient(90deg,#065f46 0%,#059669 50%,#047857 100%);margin-bottom:8px;border-radius:2px}

  /* ── Letterhead ── */
  .letterhead{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px}
  .lh-left{flex:1}
  .lh-center{flex:2;text-align:center}
  .lh-right{flex:1;text-align:right}
  .form-badge{display:inline-block;background:#065f46;color:#fff;font-size:8pt;font-weight:700;
              padding:3px 10px;border-radius:3px;letter-spacing:.05em;margin-bottom:3px}
  .form-title{font-size:13pt;font-weight:800;color:#065f46;letter-spacing:.03em;text-transform:uppercase}
  .form-sub{font-size:7pt;color:#475569;margin-top:2px;font-style:italic}
  .govt-label{font-size:7pt;color:#64748b;text-transform:uppercase;letter-spacing:.08em}
  .bar-name-big{font-size:11pt;font-weight:700;color:#0f172a}

  /* ── Info box ── */
  .info-box{border:1.5px solid #065f46;border-radius:4px;padding:5px 8px;margin-bottom:8px;
            display:grid;grid-template-columns:1fr 1fr;gap:2px 16px;background:#f0fdf4}
  .info-row{display:flex;gap:4px;font-size:7.5pt;color:#1e293b}
  .info-label{color:#64748b;white-space:nowrap}
  .info-val{font-weight:600;color:#0f172a}
  .info-right{text-align:right}

  /* ── Table ── */
  table{border-collapse:collapse;width:100%;font-size:7.5pt}
  .th-part{background:#0f172a;color:#fff;text-align:left;padding:4px 6px;min-width:180px;
           font-size:7pt;font-weight:700;letter-spacing:.04em;border:1px solid #064e3b;white-space:nowrap}
  .th-cat{background:#065f46;color:#fff;text-align:center;padding:3px 4px;font-size:7pt;
          font-weight:700;letter-spacing:.05em;border:1px solid rgba(255,255,255,.25)}
  .th-size{background:#059669;color:#fff;text-align:center;padding:3px 4px;font-size:8pt;
           font-weight:700;border:1px solid rgba(255,255,255,.2)}
  .th-ml{background:#047857;color:rgba(255,255,255,.8);text-align:center;padding:2px 4px;
         font-size:6pt;font-weight:600;letter-spacing:.06em;border:1px solid rgba(255,255,255,.15)}
  .th-total-hdr{background:#134e4a;color:#99f6e4;text-align:center;padding:3px 4px;
                font-size:6.5pt;font-weight:700;border:1px solid #0f766e;white-space:nowrap}
  .td-label{text-align:left;white-space:nowrap;padding:3px 6px;color:#0f172a;
            border:1px solid #cbd5e1;border-left:3px solid #6ee7b7;font-weight:500}
  .td-num{text-align:center;padding:3px 3px;border:1px solid #d1fae5;color:#1e293b;min-width:28px}
  .td-tot{text-align:center;padding:3px 5px;border:1px solid #059669;background:#d1fae5;
          color:#065f46;font-weight:700;font-size:8pt}
  tr.sep td{height:2px!important;background:#a7f3d0!important;padding:0!important;border:none!important}

  /* ── Footer ── */
  .footer{margin-top:12px;display:flex;justify-content:space-between;font-size:7.5pt;color:#475569}
  .sig-line{border-top:1px solid #64748b;padding-top:3px;min-width:140px;text-align:center;font-size:7pt}
  .print-note{font-size:6.5pt;color:#94a3b8;text-align:center;margin-top:6px}

  /* ── Bulk LTR section ── */
  .bulk-section{margin-top:14px;page-break-inside:avoid}
  .bulk-title{font-size:9pt;font-weight:800;color:#065f46;margin-bottom:5px;text-transform:uppercase;letter-spacing:.04em;border-bottom:2px solid #065f46;padding-bottom:3px}
  .bulk-table{border-collapse:collapse;font-size:8pt;min-width:400px}
  .bulk-table th,.bulk-table td{border:1px solid #6ee7b7;padding:4px 12px}
  .bulk-table thead th{background:#065f46;color:#fff;font-weight:700;text-align:center}
  .bulk-table thead th.blk-th-label{text-align:left}
  .blk-label{text-align:left;font-weight:500;white-space:nowrap}
  .blk-num{text-align:center;font-variant-numeric:tabular-nums}
  .blk-total td{background:#d1fae5;font-weight:700;border-top:2px solid #065f46}
  ${bw ? 'html,body{-webkit-filter:grayscale(1)!important;filter:grayscale(1)!important}' : ''}
</style>
</head><body>
  <div class="top-bar"></div>

  <div class="letterhead">
    <div class="lh-left">
      <div class="govt-label">Government of Maharashtra</div>
      <div class="govt-label">Excise Department</div>
    </div>
    <div class="lh-center">
      <div><span class="form-badge">MML MONTHLY STATEMENT &nbsp; [Form F.L.R.- 4 Style]</span></div>
      <div class="form-title">Monthly Statement — Maharashtra Made Liquor</div>
      <div class="form-sub">Monthly return of Transactions of MML effected by Licensee</div>
    </div>
    <div class="lh-right">
      <div class="govt-label">Month of</div>
      <div class="bar-name-big">${monthName}, ${selYear}</div>
    </div>
  </div>

  <div class="info-box">
    <div class="info-row"><span class="info-label">Licensee Name :</span><span class="info-val">&nbsp;${barName}</span></div>
    <div class="info-row info-right"><span class="info-label">Licence No :</span><span class="info-val">&nbsp;${licNo}</span></div>
    <div class="info-row"><span class="info-label">Address :</span><span class="info-val">&nbsp;${address}</span></div>
    <div class="info-row info-right"><span class="info-label">Financial Year :</span><span class="info-val">&nbsp;${activeBar.financialYear || '—'}</span></div>
  </div>

  <table>
    <thead>
      <tr>${ph1}</tr>
      <tr>${ph2}</tr>
      <tr>${ph3}</tr>
    </thead>
    <tbody>${tbody}</tbody>
  </table>

  <div class="bulk-section">
    <div class="bulk-title">MML BULK LTR REPORT &mdash; Month of : ${monthName}, ${selYear}</div>
    <table class="bulk-table">
      <thead><tr><th class="blk-th-label">PARTICULARS</th><th>OPENING</th><th>RECEIVED</th><th>SALE</th><th>CLOSING STOCK</th></tr></thead>
      <tbody>
        <tr><td class="blk-label">MML</td><td class="blk-num">${fmtB(mmlBulkO)}</td><td class="blk-num">${fmtB(mmlBulkR)}</td><td class="blk-num">${fmtB(mmlBulkS)}</td><td class="blk-num">${fmtB(mmlBulkC)}</td></tr>
        <tr class="blk-total"><td class="blk-label"><strong>TOTAL :</strong></td><td class="blk-num"><strong>${fmtB(mmlBulkO)}</strong></td><td class="blk-num"><strong>${fmtB(mmlBulkR)}</strong></td><td class="blk-num"><strong>${fmtB(mmlBulkS)}</strong></td><td class="blk-num"><strong>${fmtB(mmlBulkC)}</strong></td></tr>
      </tbody>
    </table>
  </div>

  <div class="footer">
    <div class="sig-line">Date &amp; Seal of Licensee</div>
    <div class="print-note">Generated by SPLIQOUR PRO &nbsp;|&nbsp; Printed on ${new Date().toLocaleDateString('en-IN')}</div>
    <div class="sig-line">Signature of Licensee</div>
  </div>

</body></html>`;
            if (_previewMode) { _previewMode = false; showPrintPreview(__mmlPdfHtml); }
            else { downloadPdf(__mmlPdfHtml, `Monthly_Statement_MML_${(barName||'').replace(/[^a-z0-9]/gi,'_')}_${monthName}_${selYear}${bw?'_BW':''}.pdf`); }
        }

        // Event wiring
        if (mmlGoBtn)     mmlGoBtn.addEventListener('click', generateMmlMonthlyStatement);
        if (mmlExportBtn) mmlExportBtn.addEventListener('click', exportMmlCsv);
        if (mmlMsPdfBtn)    mmlMsPdfBtn.addEventListener('click', () => printMmlStatement(false));
        if (mmlMsBwPdfBtn)  mmlMsBwPdfBtn.addEventListener('click', () => printMmlStatement(true));
        document.getElementById('mmlMsPreviewBtn')?.addEventListener('click', () => { _previewMode = true; printMmlStatement(false); });

        // Auto-generate on panel activation
        const mmlPanel = document.getElementById('sub-mml-monthly-statement');
        if (mmlPanel) {
            const mmlObserver = new MutationObserver((mutations) => {
                mutations.forEach(m => {
                    if (m.attributeName === 'class' && mmlPanel.classList.contains('active') && !mmlGenerated) {
                        generateMmlMonthlyStatement();
                    }
                });
            });
            mmlObserver.observe(mmlPanel, { attributes: true, attributeFilter: ['class'] });
        }
    }

    // ═══════════════════════════════════════════════════════
    //  MML BRANDWISE REPORT — Maharashtra Made Liquor
    //  Same layout as main Brandwise, filtered to MML only
    // ═══════════════════════════════════════════════════════
    {
        const mmlBwDateFrom   = document.getElementById('mmlBwDateFrom');
        const mmlBwDateTo     = document.getElementById('mmlBwDateTo');
        const mmlBwGoBtn      = document.getElementById('mmlBwGoBtn');
        const mmlBwSearch     = document.getElementById('mmlBwSearch');
        const mmlBwExportBtn  = document.getElementById('mmlBwExportBtn');
        const mmlBwPrintBtn   = document.getElementById('mmlBwPrintBtn');
        const mmlBwFyBadge    = document.getElementById('mmlBwFyBadge');
        const mmlBwTableHead  = document.getElementById('mmlBwTableHead');
        const mmlBwTableBody  = document.getElementById('mmlBwTableBody');
        const mmlBwTableFoot  = document.getElementById('mmlBwTableFoot');
        const mmlBwEmpty      = document.getElementById('mmlBwEmpty');

        const mmlBwTotalBrands  = document.getElementById('mmlBwTotalBrands');
        const mmlBwTotalOpening = document.getElementById('mmlBwTotalOpening');
        const mmlBwTotalReceipt = document.getElementById('mmlBwTotalReceipt');
        const mmlBwTotalSale    = document.getElementById('mmlBwTotalSale');
        const mmlBwTotalClosing = document.getElementById('mmlBwTotalClosing');

        let mmlBwRows = [];
        let mmlBwFiltered = [];
        let mmlBwSizeCols = [];
        let mmlBwGenerated = false;

        // FY defaults
        const mmlBwFyStartYear = (() => {
            const now = new Date();
            return now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
        })();
        const mmlBwFyStartDate = mmlBwFyStartYear + '-04-01';
        if (mmlBwDateFrom) mmlBwDateFrom.value = mmlBwFyStartDate;
        if (mmlBwDateTo)   mmlBwDateTo.value = new Date().toISOString().slice(0, 10);
        if (mmlBwFyBadge)  mmlBwFyBadge.textContent = 'FY ' + mmlBwFyStartYear + '-' + String(mmlBwFyStartYear + 1).slice(2);

        /* ── Helpers ── */
        function mmlBwSizeML(s) {
            const m = (s || '').match(/^(\d+)\s*ML/i);
            if (m) return parseInt(m[1]);
            const l = (s || '').match(/^(\d+)\s*Ltr/i);
            if (l) return parseInt(l[1]) * 1000;
            return 99999;
        }

        function mmlBwSizeGroup(s) {
            const m = (s || '').match(/^(\d+\s*ML)/i);
            if (m) return m[1].replace(/\s+/g, ' ').toUpperCase().replace('ML', 'ML');
            const l = (s || '').match(/^(\d+\s*Ltr)/i);
            if (l) return l[1];
            return s || '';
        }

        function mmlBwEsc(str) {
            return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        }

        /* ── Check if a product is MML ── */
        function isMmlProduct(brandName, category) {
            if (category && category.toUpperCase() === 'MML') return true;
            const nb = (brandName || '').trim().replace(/\./g, '').replace(/\s+/g, ' ').trim().toUpperCase();
            const match = allProducts.find(p =>
                (p.brandName || '').trim().replace(/\./g, '').replace(/\s+/g, ' ').trim().toUpperCase() === nb
            );
            return match && match.category && match.category.toUpperCase() === 'MML';
        }

        /* ── Main generate ── */
        async function generateMmlBrandwise() {
            if (!activeBar.barName) return;
            const params = { barName: activeBar.barName, financialYear: activeBar.financialYear || '' };
            const fromDate = mmlBwDateFrom ? mmlBwDateFrom.value : mmlBwFyStartDate;
            const toDate   = mmlBwDateTo   ? mmlBwDateTo.value   : '';

            // Update AS ON DATE label
            const mmlBwAsOnDate = document.getElementById('mmlBwAsOnDate');
            if (mmlBwAsOnDate) {
                const effectiveDate = toDate || new Date().toISOString().slice(0, 10);
                const [y, m, d] = effectiveDate.split('-');
                mmlBwAsOnDate.textContent = d + '/' + m + '/' + y;
            }

            try {
                const [osResult, tpResult, salesItems] = await Promise.all([
                    window.electronAPI.getOpeningStock(params),
                    window.electronAPI.getTps(params),
                    getSalesData(params),
                ]);

                const brandMap = new Map();
                const allSizeGroups = new Set();
                const normBrand = (s) => (s || '').trim().replace(/\./g, '').replace(/\s+/g, ' ').trim().toUpperCase();

                function ensureBrand(brandName, category, subCat) {
                    const key = normBrand(brandName);
                    if (!brandMap.has(key)) {
                        brandMap.set(key, {
                            brandName: (brandName || '').trim(),
                            category: category || '',
                            subCategory: subCat || '',
                            tpNumbers: new Set(),
                            sizes: {},
                        });
                    }
                    const row = brandMap.get(key);
                    if (category && !row.category) row.category = category;
                    if (subCat && !row.subCategory) row.subCategory = subCat;
                    return row;
                }

                function ensureSize(brandRow, sz) {
                    const sg = mmlBwSizeGroup(sz);
                    allSizeGroups.add(sg);
                    if (!brandRow.sizes[sg]) {
                        brandRow.sizes[sg] = { opening: 0, purchase: 0, sale: 0, closing: 0 };
                    }
                    return brandRow.sizes[sg];
                }

                // 1. Opening Stock — MML only
                if (osResult.success && osResult.data && osResult.data.entries) {
                    osResult.data.entries.forEach(e => {
                        if (!isMmlProduct(e.brandName, e.category)) return;
                        const row = ensureBrand(e.brandName, e.category, e.subCategory || '');
                        const sz = ensureSize(row, e.size);
                        sz.opening += (e.totalBtl || 0);
                    });
                }

                // 2. TP Purchases — MML only
                if (tpResult.success && tpResult.tps) {
                    tpResult.tps.forEach(tp => {
                        const d = tp.tpDate || '';
                        (tp.items || []).forEach(item => {
                            if (!isMmlProduct(item.brand || item.brandName, item.category)) return;
                            const row = ensureBrand(item.brand || item.brandName, item.category || '', '');
                            const sz = ensureSize(row, item.size);
                            if (tp.tpNumber) row.tpNumbers.add(tp.tpNumber);

                            if (fromDate && d && d < fromDate) {
                                sz.opening += (item.totalBtl || 0);
                            } else if (toDate && d && d > toDate) {
                                // ignore
                            } else {
                                sz.purchase += (item.totalBtl || 0);
                            }
                        });
                    });
                }

                // 3. Sales — MML only
                (salesItems || []).forEach(sale => {
                    if (!isMmlProduct(sale.brandName, sale.category)) return;
                    const d = sale.billDate || '';
                    const row = ensureBrand(sale.brandName, sale.category, '');
                    const sz = ensureSize(row, sale.size);

                    if (fromDate && d && d < fromDate) {
                        sz.opening -= (sale.totalBtl || 0);
                    } else if (toDate && d && d > toDate) {
                        // ignore
                    } else {
                        sz.sale += (sale.totalBtl || 0);
                    }
                });

                // 4. Enrich categories
                brandMap.forEach(row => {
                    if (!row.category || !row.subCategory) {
                        const nb = normBrand(row.brandName);
                        const match = allProducts.find(p => normBrand(p.brandName) === nb);
                        if (match) {
                            if (!row.category) row.category = match.category || '';
                            if (!row.subCategory) row.subCategory = match.subCategory || '';
                        }
                    }
                    Object.values(row.sizes).forEach(sz => {
                        if (sz.opening < 0) sz.opening = 0;
                        sz.closing = sz.opening + sz.purchase - sz.sale;
                        if (sz.closing < 0) sz.closing = 0;
                    });
                });

                // Remove zero-only brands
                brandMap.forEach((row, key) => {
                    const hasAny = Object.values(row.sizes).some(sz =>
                        sz.opening || sz.purchase || sz.sale || sz.closing
                    );
                    if (!hasAny) brandMap.delete(key);
                });

                // Also filter out any non-MML that snuck through
                brandMap.forEach((row, key) => {
                    if (row.category && row.category.toUpperCase() !== 'MML') brandMap.delete(key);
                });

                mmlBwSizeCols = [...allSizeGroups].sort((a, b) => mmlBwSizeML(b) - mmlBwSizeML(a));
                mmlBwRows = Array.from(brandMap.values());
                mmlBwRows.sort((a, b) => {
                    return a.brandName.localeCompare(b.brandName, 'en', { sensitivity: 'base' });
                });

                mmlBwGenerated = true;
                applyMmlBwFilters();

            } catch (err) {
                console.error('MML Brandwise generate error:', err);
            }
        }

        function applyMmlBwFilters() {
            const searchVal = mmlBwSearch ? mmlBwSearch.value.trim().toUpperCase() : '';
            mmlBwFiltered = mmlBwRows.filter(r => {
                if (searchVal && !r.brandName.toUpperCase().includes(searchVal)) return false;
                return true;
            });
            renderMmlBwTable();
        }

        /* ── Build multi-level header ── */
        function buildMmlBwHeader() {
            if (!mmlBwTableHead) return;
            mmlBwTableHead.innerHTML = '';
            const nSizes = mmlBwSizeCols.length || 1;

            const tr1 = document.createElement('tr');
            tr1.className = 'bw-header-group';
            tr1.innerHTML = `
                <th rowspan="2" class="bw-th-sr bw-th-frozen">Sr.</th>
                <th rowspan="2" class="bw-th-brand bw-th-frozen">Brand Name</th>
                <th rowspan="2" class="bw-th-tp bw-th-frozen">TP No.</th>
                <th colspan="${nSizes}" class="bw-th-group bw-th-opening">OPENING</th>
                <th colspan="${nSizes}" class="bw-th-group bw-th-purchase">PURCHASE</th>
                <th colspan="${nSizes}" class="bw-th-group bw-th-sales">SALES</th>
                <th colspan="${nSizes}" class="bw-th-group bw-th-closing-hdr">CLOSING</th>
            `;
            mmlBwTableHead.appendChild(tr1);

            const tr2 = document.createElement('tr');
            tr2.className = 'bw-header-sizes';
            const groups = ['opening', 'purchase', 'sales', 'closing'];
            groups.forEach(g => {
                mmlBwSizeCols.forEach(sz => {
                    const shortLabel = sz.replace(/\s*ML/i, '').trim();
                    tr2.innerHTML += `<th class="bw-th-sz bw-th-sz-${g}">${shortLabel}</th>`;
                });
            });
            mmlBwTableHead.appendChild(tr2);
        }

        /* ── Render table body ── */
        function renderMmlBwTable() {
            if (!mmlBwTableBody) return;

            buildMmlBwHeader();
            mmlBwTableBody.innerHTML = '';
            if (mmlBwTableFoot) mmlBwTableFoot.innerHTML = '';

            if (mmlBwFiltered.length === 0 || mmlBwSizeCols.length === 0) {
                if (mmlBwEmpty) mmlBwEmpty.classList.remove('hidden');
                document.getElementById('mmlBwTable')?.classList.add('hidden');
                updateMmlBwSummary([]);
                return;
            }
            if (mmlBwEmpty) mmlBwEmpty.classList.add('hidden');
            document.getElementById('mmlBwTable')?.classList.remove('hidden');

            let sr = 0;
            const nSizes = mmlBwSizeCols.length;
            const totalCols = 3 + nSizes * 4;
            const fmt = (n) => n ? n.toLocaleString('en-IN') : '';

            const footOpen = new Array(nSizes).fill(0);
            const footPur  = new Array(nSizes).fill(0);
            const footSale = new Array(nSizes).fill(0);
            const footCl   = new Array(nSizes).fill(0);

            mmlBwFiltered.forEach(row => {
                sr++;
                const tr = document.createElement('tr');
                const tpShort = [...row.tpNumbers].map(t => { const i = t.lastIndexOf('/'); return i >= 0 ? t.slice(i + 1) : t; }).join(', ') || '—';
                const tpFull = [...row.tpNumbers].join(', ') || '—';

                let cells = `
                    <td class="bw-td-sr">${sr}</td>
                    <td class="bw-td-brand" title="${mmlBwEsc(row.brandName)}">${mmlBwEsc(row.brandName)}</td>
                    <td class="bw-td-tp" title="${mmlBwEsc(tpFull)}">${mmlBwEsc(tpShort)}</td>
                `;

                // Opening
                mmlBwSizeCols.forEach((sz, si) => {
                    const v = (row.sizes[sz] ? row.sizes[sz].opening : 0);
                    footOpen[si] += v;
                    cells += `<td class="bw-td-num ${v ? '' : 'bw-td-zero'}">${fmt(v)}</td>`;
                });
                // Purchase
                mmlBwSizeCols.forEach((sz, si) => {
                    const v = (row.sizes[sz] ? row.sizes[sz].purchase : 0);
                    footPur[si] += v;
                    cells += `<td class="bw-td-num ${v ? '' : 'bw-td-zero'}">${fmt(v)}</td>`;
                });
                // Sales
                mmlBwSizeCols.forEach((sz, si) => {
                    const v = (row.sizes[sz] ? row.sizes[sz].sale : 0);
                    footSale[si] += v;
                    cells += `<td class="bw-td-num ${v ? '' : 'bw-td-zero'}">${fmt(v)}</td>`;
                });
                // Closing
                mmlBwSizeCols.forEach((sz, si) => {
                    const v = (row.sizes[sz] ? row.sizes[sz].closing : 0);
                    footCl[si] += v;
                    const cls = v < 0 ? 'bw-td-neg' : (v ? '' : 'bw-td-zero');
                    cells += `<td class="bw-td-num bw-td-closing ${cls}">${fmt(v)}</td>`;
                });

                tr.innerHTML = cells;
                mmlBwTableBody.appendChild(tr);
            });

            // Footer
            if (mmlBwTableFoot) {
                const ftr = document.createElement('tr');
                ftr.className = 'bw-foot-row';
                let fcells = `<td colspan="3" class="bw-foot-label">GRAND TOTAL</td>`;
                footOpen.forEach(v => { fcells += `<td class="bw-foot-num">${fmt(v)}</td>`; });
                footPur.forEach(v  => { fcells += `<td class="bw-foot-num">${fmt(v)}</td>`; });
                footSale.forEach(v => { fcells += `<td class="bw-foot-num">${fmt(v)}</td>`; });
                footCl.forEach(v   => { fcells += `<td class="bw-foot-num">${fmt(v)}</td>`; });
                ftr.innerHTML = fcells;
                mmlBwTableFoot.appendChild(ftr);
            }

            updateMmlBwSummary(mmlBwFiltered);
        }

        function updateMmlBwSummary(rows) {
            let totalOpening = 0, totalPurchase = 0, totalSale = 0, totalClosing = 0;
            rows.forEach(r => {
                Object.values(r.sizes).forEach(sz => {
                    totalOpening  += sz.opening;
                    totalPurchase += sz.purchase;
                    totalSale     += sz.sale;
                    totalClosing  += sz.closing;
                });
            });
            const fmt = (n) => n.toLocaleString('en-IN');
            if (mmlBwTotalBrands)  mmlBwTotalBrands.textContent = rows.length;
            if (mmlBwTotalOpening) mmlBwTotalOpening.textContent = fmt(totalOpening);
            if (mmlBwTotalReceipt) mmlBwTotalReceipt.textContent = fmt(totalPurchase);
            if (mmlBwTotalSale)    mmlBwTotalSale.textContent = fmt(totalSale);
            if (mmlBwTotalClosing) mmlBwTotalClosing.textContent = fmt(totalClosing);
        }

        /* ── Export CSV ── */
        function exportMmlBwCSV() {
            if (mmlBwFiltered.length === 0) return;
            const fromStr = mmlBwDateFrom ? mmlBwDateFrom.value : '';
            const toStr = mmlBwDateTo ? mmlBwDateTo.value : '';

            let hdr1 = ['Sr', 'Brand Name', 'TP No'];
            let hdr2 = ['', '', ''];
            ['Opening', 'Purchase', 'Sales', 'Closing'].forEach(g => {
                mmlBwSizeCols.forEach((sz, i) => {
                    hdr1.push(i === 0 ? g : '');
                    hdr2.push(sz);
                });
            });

            const csvRows = [hdr1.join(','), hdr2.join(',')];
            mmlBwFiltered.forEach((row, idx) => {
                const line = [
                    idx + 1,
                    '"' + (row.brandName || '').replace(/"/g, '""') + '"',
                    '"' + ([...row.tpNumbers].map(t => { const j = t.lastIndexOf('/'); return j >= 0 ? t.slice(j + 1) : t; }).join('; ') || '').replace(/"/g, '""') + '"',
                ];
                ['opening', 'purchase', 'sale', 'closing'].forEach(fld => {
                    mmlBwSizeCols.forEach(sz => {
                        line.push(row.sizes[sz] ? row.sizes[sz][fld] : 0);
                    });
                });
                csvRows.push(line.join(','));
            });

            const blob = new Blob(['\uFEFF' + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `MML_Brandwise_${activeBar.barName || 'Bar'}_${fromStr}_to_${toStr}.csv`;
            a.click();
            URL.revokeObjectURL(a.href);
        }

        /* ── Print ── */
        function printMmlBwReport() {
            if (!mmlBwFiltered.length || !mmlBwSizeCols.length) return;
            const fromStr = mmlBwDateFrom ? mmlBwDateFrom.value : '';
            const toStr = mmlBwDateTo ? mmlBwDateTo.value : '';
            const barName = activeBar.barName || '';
            const fy = mmlBwFyBadge ? mmlBwFyBadge.textContent : '';
            const effectiveDateRaw = toStr || new Date().toISOString().slice(0, 10);
            const [edy, edm, edd] = effectiveDateRaw.split('-');
            const asOnDate = edd + '/' + edm + '/' + edy;
            const nSizes = mmlBwSizeCols.length;

            let th1 = '<th rowspan="2" style="min-width:30px">Sr.</th><th rowspan="2" style="text-align:left;min-width:180px">Brand Name</th><th rowspan="2" style="min-width:90px">TP No.</th>';
            ['OPENING', 'PURCHASE', 'SALES', 'CLOSING'].forEach(g => {
                const bgMap = { OPENING: '#dbeafe', PURCHASE: '#dcfce7', SALES: '#fee2e2', CLOSING: '#fef3c7' };
                th1 += `<th colspan="${nSizes}" style="background:${bgMap[g]};text-align:center">${g}</th>`;
            });

            let th2 = '';
            ['opening', 'purchase', 'sales', 'closing'].forEach(() => {
                mmlBwSizeCols.forEach(sz => {
                    th2 += `<th style="text-align:center;font-size:7pt;padding:2px 4px">${sz.replace(/\s*ML/i, '')}</th>`;
                });
            });

            let tbody = '';
            const totalCols = 3 + nSizes * 4;
            mmlBwFiltered.forEach((row, i) => {
                tbody += `<tr><td style="text-align:center">${i + 1}</td><td>${row.brandName}</td><td style="font-size:7pt">${[...row.tpNumbers].map(t => { const j = t.lastIndexOf('/'); return j >= 0 ? t.slice(j + 1) : t; }).join(', ') || '—'}</td>`;
                ['opening', 'purchase', 'sale', 'closing'].forEach(fld => {
                    mmlBwSizeCols.forEach(sz => {
                        const v = row.sizes[sz] ? row.sizes[sz][fld] : 0;
                        tbody += `<td style="text-align:center">${v || ''}</td>`;
                    });
                });
                tbody += '</tr>';
            });

            printWithIframe(`<!DOCTYPE html><html><head><title>MML Brandwise — ${barName}</title>
                <style>body{font-family:'Inter',Arial,sans-serif;padding:12px;font-size:8pt} h2{margin:0 0 2px;font-size:14pt} .form-flr-label{font-size:8.5pt;font-weight:600;color:#1e40af;margin-bottom:6px;text-transform:uppercase;letter-spacing:.03em} .meta{color:#666;font-size:9pt;margin-bottom:8px}
                table{border-collapse:collapse;width:100%} th,td{border:1px solid #999;padding:3px 5px;font-size:8pt} th{background:#f0f0f0}
                @page{size:landscape;margin:8mm}</style>
            </head><body>
                <h2>MML Brandwise Stock Register</h2>
                <div class="form-flr-label">FORM FLR :- 1A/2A/3A (FL) &nbsp;&nbsp;|&nbsp;&nbsp; DAILY BRANDWISE REGISTER &nbsp;&nbsp;|&nbsp;&nbsp; AS ON DATE: ${asOnDate}</div>
                <div class="meta"><strong>${barName}</strong> &nbsp;|&nbsp; ${fy} &nbsp;|&nbsp; Period: ${fromStr || 'FY Start'} to ${toStr || 'Today'}</div>
                <table><thead><tr>${th1}</tr><tr>${th2}</tr></thead><tbody>${tbody}</tbody></table>
            </body></html>`);
        }

        // Event wiring
        if (mmlBwGoBtn)     mmlBwGoBtn.addEventListener('click', generateMmlBrandwise);
        if (mmlBwSearch)    mmlBwSearch.addEventListener('input', applyMmlBwFilters);
        if (mmlBwExportBtn) mmlBwExportBtn.addEventListener('click', exportMmlBwCSV);
        if (mmlBwPrintBtn)  mmlBwPrintBtn.addEventListener('click', printMmlBwReport);
        document.getElementById('mmlBwPreviewBtn')?.addEventListener('click', () => { _previewMode = true; printMmlBwReport(); });

        // Auto-generate when panel becomes visible
        const mmlBwPanel = document.getElementById('sub-mml-brandwise');
        if (mmlBwPanel) {
            const mmlBwObserver = new MutationObserver((mutations) => {
                mutations.forEach(m => {
                    if (m.attributeName === 'class' && mmlBwPanel.classList.contains('active') && !mmlBwGenerated) {
                        generateMmlBrandwise();
                    }
                });
            });
            mmlBwObserver.observe(mmlBwPanel, { attributes: true, attributeFilter: ['class'] });
        }
    }

    // ═══════════════════════════════════════════════════════
    //  CHATAI REGISTER — Date-wise stock register
    // ═══════════════════════════════════════════════════════
    {
        const crDateFrom   = document.getElementById('crDateFrom');
        const crDateTo     = document.getElementById('crDateTo');
        const crGoBtn      = document.getElementById('crGoBtn');
        const crExportBtn  = document.getElementById('crExportBtn');
        const crPrintBtn   = document.getElementById('crPrintBtn');
        const crMmlToggle  = document.getElementById('crMmlToggle');
        const crTableHead  = document.getElementById('crTableHead');
        const crTableBody  = document.getElementById('crTableBody');
        const crTableFoot  = document.getElementById('crTableFoot');
        const crEmpty      = document.getElementById('crEmpty');
        const crFromLabel  = document.getElementById('crFromLabel');
        const crToLabel    = document.getElementById('crToLabel');
        const crTotalDays     = document.getElementById('crTotalDays');
        const crTotalOpening  = document.getElementById('crTotalOpening');
        const crTotalPurchase = document.getElementById('crTotalPurchase');
        const crTotalSale     = document.getElementById('crTotalSale');
        const crTotalClosing  = document.getElementById('crTotalClosing');

        // FY helpers
        const crFyStartYear = (() => { const n = new Date(); return n.getMonth() >= 3 ? n.getFullYear() : n.getFullYear() - 1; })();
        const crFyStart = crFyStartYear + '-04-01';

        // Default: start of current month → today
        const crToday = new Date().toISOString().slice(0, 10);
        const crMonthStart = (() => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-01'; })();
        if (crDateFrom) crDateFrom.value = crMonthStart;
        if (crDateTo)   crDateTo.value   = crToday;

        /* ── Date range modal ── */
        const crDateModal      = document.getElementById('crDateModal');
        const crModalFrom      = document.getElementById('crModalFrom');
        const crModalTo        = document.getElementById('crModalTo');
        const crModalGoBtn     = document.getElementById('crModalGoBtn');
        const crModalCancelBtn = document.getElementById('crModalCancelBtn');

        function showCrDateModal() {
            if (!crDateModal) return;
            if (crModalFrom) crModalFrom.value = crMonthStart;
            if (crModalTo)   crModalTo.value   = crToday;
            crDateModal.classList.remove('hidden');
            requestAnimationFrame(() => { if (crModalFrom) crModalFrom.focus(); });
        }
        function hideCrDateModal() { if (crDateModal) crDateModal.classList.add('hidden'); }
        function confirmCrDateModal() {
            if (!crModalFrom || !crModalTo || !crModalFrom.value || !crModalTo.value) return;
            if (crDateFrom) crDateFrom.value = crModalFrom.value;
            if (crDateTo)   crDateTo.value   = crModalTo.value;
            hideCrDateModal();
            generateChataiRegister();
        }
        if (crModalGoBtn)     crModalGoBtn.addEventListener('click', confirmCrDateModal);
        if (crModalCancelBtn) crModalCancelBtn.addEventListener('click', hideCrDateModal);
        if (crDateModal) {
            crDateModal.addEventListener('keydown', e => {
                if (e.key === 'Enter')  { e.preventDefault(); confirmCrDateModal(); }
                if (e.key === 'Escape') { e.preventDefault(); hideCrDateModal(); }
            });
            crDateModal.addEventListener('click', e => { if (e.target === crDateModal) hideCrDateModal(); });
        }
        if (crModalFrom) {
            crModalFrom.addEventListener('keydown', e => {
                if (e.key === 'Enter' && crModalFrom.value && (!crModalTo || !crModalTo.value)) {
                    e.preventDefault(); if (crModalTo) crModalTo.focus();
                }
            });
        }
        window._showCrDateModal = showCrDateModal;

        /* ── Category definitions ── */
        const CR_CATS = [
            { key: 'Spirits',        label: 'SPIRITS',         color: '#dbeafe', darkColor: 'rgba(96,165,250,.18)',  textColor: '#1e40af', darkText: '#93c5fd',
              match: raw => /^(spirits?|imfl)$/i.test(raw) },
            { key: 'Wines',          label: 'WINES',            color: '#f3e8ff', darkColor: 'rgba(192,132,252,.14)', textColor: '#6b21a8', darkText: '#d8b4fe',
              match: raw => /^wines?$/i.test(raw) },
            { key: 'FermentedBeer',  label: 'FERMENTED BEER',  color: '#dcfce7', darkColor: 'rgba(52,211,153,.14)',  textColor: '#166534', darkText: '#6ee7b7',
              match: raw => /fermented/i.test(raw) },
            { key: 'MildBeer',       label: 'MILD BEER',        color: '#fef9c3', darkColor: 'rgba(250,204,21,.14)',  textColor: '#713f12', darkText: '#fde68a',
              match: raw => /mild/i.test(raw) },
            { key: 'MML',            label: 'MML',              color: '#fee2e2', darkColor: 'rgba(248,113,113,.14)', textColor: '#991b1b', darkText: '#fca5a5',
              match: raw => /^mml$/i.test(raw) },
        ];

        function crGetCatKey(category) {
            const raw = (category || '').trim();
            for (const c of CR_CATS) { if (c.match(raw)) return c.key; }
            return 'Spirits'; // default fallback
        }

        /* ── Size helpers ── */
        function crSizeML(s) {
            const m = (s || '').match(/^(\d+)\s*ML/i); if (m) return parseInt(m[1]);
            const l = (s || '').match(/^(\d+)\s*Ltr/i); if (l) return parseInt(l[1]) * 1000;
            return 99999;
        }
        function crSizeGroup(s) {
            const m = (s || '').match(/^(\d+\s*ML)/i);
            if (m) return m[1].replace(/\s+/g, ' ').toUpperCase();
            const l = (s || '').match(/^(\d+\s*Ltr)/i); if (l) return l[1];
            return s || '';
        }

        /* ── Date helpers ── */
        function crDateRange(from, to) {
            const dates = [];
            let cur = new Date(from); const end = new Date(to);
            while (cur <= end) { dates.push(cur.toISOString().slice(0, 10)); cur.setDate(cur.getDate() + 1); }
            return dates;
        }
        function fmtDate(iso) {
            if (!iso) return '—';
            const [y, m, d] = iso.split('-');
            return d + '/' + m + '/' + y;
        }

        // Stored report data for export/print
        let crLastRows = [];
        let crLastCatCols = [];   // [{catKey, catLabel, sizes:[sz,...], color, ...}]

        /* ── Build active category-size columns ── */
        function crBuildCols(catSizeMap, showMml) {
            const cols = [];
            for (const cat of CR_CATS) {
                if (cat.key === 'MML' && !showMml) continue;
                const sizes = [...(catSizeMap[cat.key] || [])].sort((a, b) => crSizeML(b) - crSizeML(a));
                if (sizes.length === 0) continue;
                cols.push({ ...cat, sizes });
            }
            return cols;
        }

        /* ── colId: unique key for (catKey, size) ── */
        const crColId = (catKey, sz) => catKey + '|' + sz;

        /* ── Main generate ── */
        async function generateChataiRegister() {
            if (!activeBar.barName) return;
            const fromDate = crDateFrom ? crDateFrom.value : crMonthStart;
            const toDate   = crDateTo   ? crDateTo.value   : crToday;
            if (!fromDate || !toDate || fromDate > toDate) return;

            if (crFromLabel) crFromLabel.textContent = fmtDate(fromDate);
            if (crToLabel)   crToLabel.textContent   = fmtDate(toDate);

            const params = { barName: activeBar.barName, financialYear: activeBar.financialYear || '' };

            const [osResult, tpResult, salesItems, productsResult] = await Promise.all([
                window.electronAPI.getOpeningStock(params),
                window.electronAPI.getTps(params),
                getSalesData(params),
                window.electronAPI.getProducts(params),
            ]);

            // ── Build product map: code/brandName → category ──
            const prodCatMap = new Map(); // (code or brandName normalized) → catKey
            if (productsResult.success && productsResult.products) {
                for (const p of productsResult.products) {
                    const ck = crGetCatKey(p.category);
                    if (p.code) prodCatMap.set(p.code, ck);
                    if (p.brandName) prodCatMap.set((p.brandName || '').trim().toUpperCase(), ck);
                }
            }

            function resolveCatKey(item) {
                if (item.category) return crGetCatKey(item.category);
                if (item.code && prodCatMap.has(item.code)) return prodCatMap.get(item.code);
                const bn = (item.brand || item.brandName || '').trim().toUpperCase();
                if (bn && prodCatMap.has(bn)) return prodCatMap.get(bn);
                return 'Spirits';
            }

            // ── Collect which sizes exist per category ──
            const catSizeSet = {}; // catKey → Set of size strings
            for (const cat of CR_CATS) catSizeSet[cat.key] = new Set();

            const addCatSize = (catKey, sz) => {
                if (!catSizeSet[catKey]) catSizeSet[catKey] = new Set();
                catSizeSet[catKey].add(sz);
            };

            // ── Index TPs ──
            // tpByDate[date] = [ {tpNumber, items:[{catKey, sz, btl}]} ]
            const tpByDate = {};
            const tpBefore = {}; // colId → btl before fromDate

            if (tpResult.success && tpResult.tps) {
                for (const tp of tpResult.tps) {
                    const d = tp.tpDate || '';
                    for (const item of (tp.items || [])) {
                        const sz  = crSizeGroup(item.size);
                        const ck  = resolveCatKey({ category: item.category, code: item.code, brand: item.brand });
                        const btl = item.totalBtl || 0;
                        addCatSize(ck, sz);
                        const cid = crColId(ck, sz);
                        if (d < fromDate) {
                            tpBefore[cid] = (tpBefore[cid] || 0) + btl;
                        } else if (d <= toDate) {
                            if (!tpByDate[d]) tpByDate[d] = [];
                            let tpEntry = tpByDate[d].find(t => t.tpNumber === tp.tpNumber);
                            if (!tpEntry) { tpEntry = { tpNumber: tp.tpNumber, cols: {} }; tpByDate[d].push(tpEntry); }
                            tpEntry.cols[cid] = (tpEntry.cols[cid] || 0) + btl;
                        }
                    }
                }
            }

            // ── Index Sales ──
            const saleByDate = {}; // date → { colId → btl }
            const saleBefore = {}; // colId → btl before fromDate

            for (const sale of (salesItems || [])) {
                const d   = sale.billDate || '';
                const sz  = crSizeGroup(sale.size);
                const ck  = resolveCatKey(sale);
                const btl = sale.totalBtl || 0;
                addCatSize(ck, sz);
                const cid = crColId(ck, sz);
                if (d < fromDate) {
                    saleBefore[cid] = (saleBefore[cid] || 0) + btl;
                } else if (d <= toDate) {
                    if (!saleByDate[d]) saleByDate[d] = {};
                    saleByDate[d][cid] = (saleByDate[d][cid] || 0) + btl;
                }
            }

            // ── FY opening stock ──
            const fyOpen = {}; // colId → btl
            if (osResult.success && osResult.data && osResult.data.entries) {
                for (const e of osResult.data.entries) {
                    const sz  = crSizeGroup(e.size);
                    const ck  = resolveCatKey(e);
                    addCatSize(ck, sz);
                    const cid = crColId(ck, sz);
                    fyOpen[cid] = (fyOpen[cid] || 0) + (e.totalBtl || 0);
                }
            }

            // ── Compute opening balances at fromDate ──
            const allColIds = new Set([...Object.keys(fyOpen), ...Object.keys(tpBefore), ...Object.keys(saleBefore)]);
            for (const date of Object.keys(tpByDate)) for (const tp of tpByDate[date]) for (const cid of Object.keys(tp.cols)) allColIds.add(cid);
            for (const date of Object.keys(saleByDate)) for (const cid of Object.keys(saleByDate[date])) allColIds.add(cid);

            const openAtStart = {};
            for (const cid of allColIds) {
                openAtStart[cid] = Math.max(0, (fyOpen[cid] || 0) + (tpBefore[cid] || 0) - (saleBefore[cid] || 0));
            }

            // ── Build columns with MML toggle ──
            const showMml = crMmlToggle ? crMmlToggle.checked : false;
            const catCols = crBuildCols(catSizeSet, showMml);
            crLastCatCols = catCols;

            // Flat list of colIds in display order
            const orderedColIds = [];
            for (const cat of catCols) for (const sz of cat.sizes) orderedColIds.push(crColId(cat.key, sz));

            // ── Build row data per date ──
            const dates = crDateRange(fromDate, toDate);
            const rows = [];
            let prevClosing = { ...openAtStart };

            for (const date of dates) {
                const tpEntries = tpByDate[date] || [];
                const saleSizes = saleByDate[date] || {};

                const purCols = {};
                for (const tpEntry of tpEntries) for (const [cid, btl] of Object.entries(tpEntry.cols)) purCols[cid] = (purCols[cid] || 0) + btl;

                const tpNos = [...new Set(tpEntries.map(t => {
                    const i = (t.tpNumber || '').lastIndexOf('/');
                    return i >= 0 ? t.tpNumber.slice(i + 1) : (t.tpNumber || '');
                }))].filter(Boolean).join(', ');

                const opening  = {};
                const purchase = {};
                const sale     = {};
                const closing  = {};

                for (const cid of orderedColIds) {
                    opening[cid]  = prevClosing[cid] || 0;
                    purchase[cid] = purCols[cid] || 0;
                    sale[cid]     = saleSizes[cid] || 0;
                    closing[cid]  = Math.max(0, opening[cid] + purchase[cid] - sale[cid]);
                }

                const hasAny = orderedColIds.some(cid => opening[cid] || purchase[cid] || sale[cid]);
                if (hasAny || rows.length > 0) rows.push({ date, tpNos, opening, purchase, sale, closing });
                prevClosing = {};
                for (const cid of allColIds) prevClosing[cid] = (closing[cid] !== undefined ? closing[cid] : (prevClosing[cid] || 0));
            }

            crLastRows = rows;
            renderChataiTable(rows, catCols, orderedColIds);
        }

        /* ── Render table ── */
        function renderChataiTable(rows, catCols, orderedColIds) {
            // Summary cards
            const sumOpen  = orderedColIds.reduce((a, c) => a + (rows[0] ? rows[0].opening[c] || 0 : 0), 0);
            const sumPur   = rows.reduce((a, r) => a + orderedColIds.reduce((b, c) => b + (r.purchase[c] || 0), 0), 0);
            const sumSale  = rows.reduce((a, r) => a + orderedColIds.reduce((b, c) => b + (r.sale[c] || 0), 0), 0);
            const sumClose = orderedColIds.reduce((a, c) => a + (rows.length ? rows[rows.length - 1].closing[c] || 0 : 0), 0);
            if (crTotalDays)     crTotalDays.textContent     = rows.length;
            if (crTotalOpening)  crTotalOpening.textContent  = sumOpen.toLocaleString('en-IN');
            if (crTotalPurchase) crTotalPurchase.textContent = sumPur.toLocaleString('en-IN');
            if (crTotalSale)     crTotalSale.textContent     = sumSale.toLocaleString('en-IN');
            if (crTotalClosing)  crTotalClosing.textContent  = sumClose.toLocaleString('en-IN');

            if (rows.length === 0) {
                if (crEmpty)     crEmpty.classList.remove('hidden');
                if (crTableHead) crTableHead.innerHTML = '';
                if (crTableBody) crTableBody.innerHTML = '';
                if (crTableFoot) crTableFoot.innerHTML = '';
                return;
            }
            if (crEmpty) crEmpty.classList.add('hidden');

            const totalCatCols = orderedColIds.length; // per section
            const isDark = document.documentElement.getAttribute('data-theme') === 'dark';

            // ── HEADER ROW 1: Date | TP No. | PURCHASE | SALES | CLOSING (OPENING removed — shown as first body row) ──
            const sectionBg   = { PURCHASE: isDark ? 'rgba(52,211,153,.18)' : '#dcfce7', SALES: isDark ? 'rgba(248,113,113,.18)' : '#fee2e2', CLOSING: isDark ? 'rgba(251,191,36,.18)' : '#fef3c7' };
            const sectionText = { PURCHASE: isDark ? '#6ee7b7' : '#166534', SALES: isDark ? '#fca5a5' : '#991b1b', CLOSING: isDark ? '#fde68a' : '#92400e' };

            let hdr1 = `<th rowspan="3" class="cr-th-date">Date</th><th rowspan="3" class="cr-th-tp">TP No.</th>`;
            for (const label of ['PURCHASE', 'SALES', 'CLOSING']) {
                hdr1 += `<th colspan="${totalCatCols}" style="background:${sectionBg[label]};color:${sectionText[label]};text-align:center;font-size:.8rem;font-weight:800;letter-spacing:.04em">${label}</th>`;
            }

            // ── HEADER ROW 2: category labels (3 sections) ──
            let hdr2 = '';
            for (let s = 0; s < 3; s++) {
                for (const cat of catCols) {
                    const bg   = isDark ? cat.darkColor : cat.color;
                    const txt  = isDark ? cat.darkText  : cat.textColor;
                    hdr2 += `<th colspan="${cat.sizes.length}" style="background:${bg};color:${txt};text-align:center;font-size:.72rem;font-weight:700;white-space:nowrap">${cat.label}</th>`;
                }
            }

            // ── HEADER ROW 3: individual sizes (3 sections) ──
            let hdr3 = '';
            for (let s = 0; s < 3; s++) {
                for (const cat of catCols) {
                    for (const sz of cat.sizes) {
                        hdr3 += `<th style="font-size:.65rem;padding:2px 4px;min-width:36px">${sz.replace(/\s*ML$/i, '').replace(/\s*Ltr$/i, 'L')}</th>`;
                    }
                }
            }

            if (crTableHead) crTableHead.innerHTML = `<tr>${hdr1}</tr><tr>${hdr2}</tr><tr>${hdr3}</tr>`;

            // ── BODY ──
            const fmtQ = v => (v ? v.toLocaleString('en-IN') : '');
            const openBg  = isDark ? 'rgba(96,165,250,.15)' : '#eff6ff';
            const openTxt = isDark ? '#93c5fd' : '#1e40af';
            // Special opening stock row (first row)
            let bodyHtml = `<tr style="background:${openBg};font-weight:600">`;
            bodyHtml += `<td class="cr-td-date" style="color:${openTxt}">Opening Stock</td>`;
            bodyHtml += `<td class="cr-td-tp" style="color:${openTxt}">${fmtDate(rows[0].date)}</td>`;
            for (const cid of orderedColIds) bodyHtml += `<td style="color:${openTxt}">${fmtQ(rows[0].opening[cid])}</td>`;
            bodyHtml += `<td colspan="${totalCatCols * 2}"></td>`;
            bodyHtml += '</tr>';
            // Date rows — PURCHASE | SALES | CLOSING only
            for (const row of rows) {
                bodyHtml += `<tr><td class="cr-td-date">${fmtDate(row.date)}</td><td class="cr-td-tp">${row.tpNos || '—'}</td>`;
                for (const fld of ['purchase', 'sale', 'closing']) {
                    for (const cid of orderedColIds) bodyHtml += `<td>${fmtQ(row[fld][cid])}</td>`;
                }
                bodyHtml += '</tr>';
            }
            if (crTableBody) crTableBody.innerHTML = bodyHtml;

            // ── FOOTER TOTALS ──
            let footHtml = `<tr class="cr-total-row"><td class="cr-td-date" colspan="2"><strong>TOTAL</strong></td>`;
            for (const fld of ['purchase', 'sale', 'closing']) {
                for (const cid of orderedColIds) {
                    const val = fld === 'closing'
                        ? (rows.length ? rows[rows.length - 1].closing[cid] || 0 : 0)
                        : rows.reduce((a, r) => a + (r[fld][cid] || 0), 0);
                    footHtml += `<td><strong>${fmtQ(val)}</strong></td>`;
                }
            }
            footHtml += '</tr>';
            if (crTableFoot) crTableFoot.innerHTML = footHtml;
        }

        /* ── MML toggle re-render ── */
        if (crMmlToggle) {
            crMmlToggle.addEventListener('change', () => {
                if (crLastRows.length > 0) generateChataiRegister();
            });
        }

        /* ── Export CSV ── */
        function exportCrCSV() {
            if (!crLastRows.length || !crLastCatCols.length) return;
            const orderedColIds = [];
            for (const cat of crLastCatCols) for (const sz of cat.sizes) orderedColIds.push(crColId(cat.key, sz));
            const headers = ['Date', 'TP No.'];
            for (const fld of ['Open','Pur','Sale','Close']) for (const cat of crLastCatCols) for (const sz of cat.sizes) headers.push(`${fld}-${cat.label}-${sz}`);
            const csvRows = [headers.join(',')];
            for (const row of crLastRows) {
                const line = [fmtDate(row.date), '"' + row.tpNos + '"'];
                for (const fld of ['opening','purchase','sale','closing']) for (const cid of orderedColIds) line.push(row[fld][cid] || 0);
                csvRows.push(line.join(','));
            }
            const blob = new Blob(['\uFEFF' + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `ChataiRegister_${activeBar.barName || 'Bar'}_${crDateFrom?.value}_to_${crDateTo?.value}.csv`;
            a.click(); URL.revokeObjectURL(a.href);
        }

        /* ── Print ── */
        function printCrReport() {
            if (!crLastRows.length || !crLastCatCols.length) return;
            const orderedColIds = [];
            for (const cat of crLastCatCols) for (const sz of cat.sizes) orderedColIds.push(crColId(cat.key, sz));
            const barName = activeBar.barName || '';
            const fromStr = crDateFrom ? crDateFrom.value : '';
            const toStr   = crDateTo   ? crDateTo.value   : '';
            const totalCatCols = orderedColIds.length;

            const secBg  = { PURCHASE:'#dcfce7', SALES:'#fee2e2', CLOSING:'#fef3c7' };
            let th1 = `<th rowspan="3">Date</th><th rowspan="3">TP No.</th>`;
            for (const label of ['PURCHASE','SALES','CLOSING']) th1 += `<th colspan="${totalCatCols}" style="background:${secBg[label]};text-align:center">${label}</th>`;
            let th2 = '';
            for (let s = 0; s < 3; s++) for (const cat of crLastCatCols) th2 += `<th colspan="${cat.sizes.length}" style="background:${cat.color};text-align:center;font-size:7pt">${cat.label}</th>`;
            let th3 = '';
            for (let s = 0; s < 3; s++) for (const cat of crLastCatCols) for (const sz of cat.sizes) th3 += `<th>${sz.replace(/\s*ML$/i,'')}</th>`;

            const fmtQ = v => (v ? v.toLocaleString('en-IN') : '');
            // Opening stock row
            let tbody = `<tr style="background:#eff6ff;font-weight:600">`;
            tbody += `<td style="color:#1e40af">Opening Stock</td><td style="color:#1e40af;font-size:6pt">${fmtDate(crLastRows[0].date)}</td>`;
            for (const cid of orderedColIds) tbody += `<td style="color:#1e40af">${fmtQ(crLastRows[0].opening[cid])}</td>`;
            tbody += `<td colspan="${totalCatCols * 2}"></td></tr>`;
            // Date rows — PURCHASE | SALES | CLOSING only
            for (const row of crLastRows) {
                tbody += `<tr><td>${fmtDate(row.date)}</td><td style="font-size:6pt">${row.tpNos || '—'}</td>`;
                for (const fld of ['purchase','sale','closing']) for (const cid of orderedColIds) tbody += `<td>${fmtQ(row[fld][cid])}</td>`;
                tbody += '</tr>';
            }
            tbody += `<tr style="font-weight:700"><td colspan="2">TOTAL</td>`;
            for (const fld of ['purchase','sale','closing']) for (const cid of orderedColIds) {
                const val = fld === 'closing' ? (crLastRows.length ? crLastRows[crLastRows.length - 1].closing[cid] || 0 : 0) : crLastRows.reduce((a, r) => a + (r[fld][cid] || 0), 0);
                tbody += `<td>${fmtQ(val)}</td>`;
            }
            tbody += '</tr>';

            printWithIframe(`<!DOCTYPE html><html><head><title>Chatai Register — ${barName}</title>
                <style>body{font-family:'Inter',Arial,sans-serif;padding:10px;font-size:7pt}
                h2{margin:0 0 4px;font-size:12pt} .meta{color:#555;font-size:7.5pt;margin-bottom:6px}
                table{border-collapse:collapse;width:100%} th,td{border:1px solid #bbb;padding:2px 3px;font-size:6.5pt;text-align:center}
                th{background:#f0f0f0} td:first-child,td:nth-child(2),th:first-child,th:nth-child(2){text-align:left}
                @page{size:landscape;margin:6mm}</style>
            </head><body>
                <h2>\ud83d\udccb Chatai Register \u2014 ${barName}</h2>
                <div class="meta">Period: <strong>${fmtDate(fromStr)}</strong> to <strong>${fmtDate(toStr)}</strong> &nbsp;|\u00a0 Generated: ${new Date().toLocaleString('en-IN')}</div>
                <table><thead><tr>${th1}</tr><tr>${th2}</tr><tr>${th3}</tr></thead><tbody>${tbody}</tbody></table>
            </body></html>`);
        }

        /* ── Event wiring ── */
        if (crGoBtn)     crGoBtn.addEventListener('click', generateChataiRegister);
        if (crExportBtn) crExportBtn.addEventListener('click', exportCrCSV);
        if (crPrintBtn)  crPrintBtn.addEventListener('click', printCrReport);
        if (crDateFrom)  crDateFrom.addEventListener('change', generateChataiRegister);
        if (crDateTo)    crDateTo.addEventListener('change',   generateChataiRegister);
        [crDateFrom, crDateTo].forEach(el => {
            el?.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); generateChataiRegister(); } });
        });
    }

    // ═══════════════════════════════════════════════════════
    //  MML CHATAI — MML-only stock register
    // ═══════════════════════════════════════════════════════
    {
        const mcDateFrom   = document.getElementById('mcDateFrom');
        const mcDateTo     = document.getElementById('mcDateTo');
        const mcGoBtn      = document.getElementById('mcGoBtn');
        const mcExportBtn  = document.getElementById('mcExportBtn');
        const mcPrintBtn   = document.getElementById('mcPrintBtn');
        const mcTableHead  = document.getElementById('mcTableHead');
        const mcTableBody  = document.getElementById('mcTableBody');
        const mcTableFoot  = document.getElementById('mcTableFoot');
        const mcEmpty      = document.getElementById('mcEmpty');
        const mcFromLabel  = document.getElementById('mcFromLabel');
        const mcToLabel    = document.getElementById('mcToLabel');
        const mcTotalDays     = document.getElementById('mcTotalDays');
        const mcTotalOpening  = document.getElementById('mcTotalOpening');
        const mcTotalPurchase = document.getElementById('mcTotalPurchase');
        const mcTotalSale     = document.getElementById('mcTotalSale');
        const mcTotalClosing  = document.getElementById('mcTotalClosing');

        const mcToday      = new Date().toISOString().slice(0, 10);
        const mcMonthStart = (() => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-01'; })();
        if (mcDateFrom) mcDateFrom.value = mcMonthStart;
        if (mcDateTo)   mcDateTo.value   = mcToday;

        /* ── Date modal ── */
        const mcDateModal      = document.getElementById('mcDateModal');
        const mcModalFrom      = document.getElementById('mcModalFrom');
        const mcModalTo        = document.getElementById('mcModalTo');
        const mcModalGoBtn     = document.getElementById('mcModalGoBtn');
        const mcModalCancelBtn = document.getElementById('mcModalCancelBtn');

        function showMcDateModal() {
            if (!mcDateModal) return;
            if (mcModalFrom) mcModalFrom.value = mcMonthStart;
            if (mcModalTo)   mcModalTo.value   = mcToday;
            mcDateModal.classList.remove('hidden');
            requestAnimationFrame(() => { if (mcModalFrom) mcModalFrom.focus(); });
        }
        function hideMcDateModal() { if (mcDateModal) mcDateModal.classList.add('hidden'); }
        function confirmMcDateModal() {
            if (!mcModalFrom || !mcModalTo || !mcModalFrom.value || !mcModalTo.value) return;
            if (mcDateFrom) mcDateFrom.value = mcModalFrom.value;
            if (mcDateTo)   mcDateTo.value   = mcModalTo.value;
            hideMcDateModal();
            generateMmlChatai();
        }
        if (mcModalGoBtn)     mcModalGoBtn.addEventListener('click', confirmMcDateModal);
        if (mcModalCancelBtn) mcModalCancelBtn.addEventListener('click', hideMcDateModal);
        if (mcDateModal) {
            mcDateModal.addEventListener('keydown', e => {
                if (e.key === 'Enter')  { e.preventDefault(); confirmMcDateModal(); }
                if (e.key === 'Escape') { e.preventDefault(); hideMcDateModal(); }
            });
            mcDateModal.addEventListener('click', e => { if (e.target === mcDateModal) hideMcDateModal(); });
        }
        if (mcModalFrom) {
            mcModalFrom.addEventListener('keydown', e => {
                if (e.key === 'Enter' && mcModalFrom.value && (!mcModalTo || !mcModalTo.value)) {
                    e.preventDefault(); if (mcModalTo) mcModalTo.focus();
                }
            });
        }
        window._showMcDateModal = showMcDateModal;

        /* ── local helpers ── */
        const mcSizeML  = s => { const m = (s||'').match(/^(\d+)\s*ML/i); if (m) return +m[1]; const l = (s||'').match(/^(\d+)\s*Ltr/i); if (l) return +l[1]*1000; return 99999; };
        const mcSizeGrp = s => { const m = (s||'').match(/^(\d+\s*ML)/i); if (m) return m[1].replace(/\s+/g,' ').toUpperCase(); const l=(s||'').match(/^(\d+\s*Ltr)/i); if(l) return l[1]; return s||''; };
        const mcFmt     = iso => { if (!iso) return '\u2014'; const [y,m,d]=iso.split('-'); return d+'/'+m+'/'+y; };
        const mcColId   = sz => 'MML|' + sz;
        const isMml     = cat => /^mml$/i.test((cat||'').trim());

        let mcLastRows    = [];
        let mcLastSizes   = [];

        /* ── Generate ── */
        async function generateMmlChatai() {
            if (!activeBar.barName) return;
            const fromDate = mcDateFrom ? mcDateFrom.value : mcMonthStart;
            const toDate   = mcDateTo   ? mcDateTo.value   : mcToday;
            if (!fromDate || !toDate || fromDate > toDate) return;

            if (mcFromLabel) mcFromLabel.textContent = mcFmt(fromDate);
            if (mcToLabel)   mcToLabel.textContent   = mcFmt(toDate);

            const params = { barName: activeBar.barName, financialYear: activeBar.financialYear || '' };

            const [osResult, tpResult, salesItems, productsResult] = await Promise.all([
                window.electronAPI.getOpeningStock(params),
                window.electronAPI.getTps(params),
                getSalesData(params),
                window.electronAPI.getProducts(params),
            ]);

            // Product map for TP category lookup
            const prodCatMap = new Map();
            if (productsResult.success && productsResult.products) {
                for (const p of productsResult.products) {
                    if (isMml(p.category)) {
                        if (p.code) prodCatMap.set(p.code, true);
                        if (p.brandName) prodCatMap.set((p.brandName||'').trim().toUpperCase(), true);
                    }
                }
            }
            const isMmlItem = item => {
                if (isMml(item.category)) return true;
                if (item.code && prodCatMap.has(item.code)) return true;
                const bn = (item.brand || item.brandName || '').trim().toUpperCase();
                return bn && prodCatMap.has(bn);
            };

            const sizeSet = new Set();
            const tpByDate = {};
            const tpBefore = {};

            if (tpResult.success && tpResult.tps) {
                for (const tp of tpResult.tps) {
                    const d = tp.tpDate || '';
                    for (const item of (tp.items || [])) {
                        if (!isMmlItem(item)) continue;
                        const sz = mcSizeGrp(item.size); sizeSet.add(sz);
                        const btl = item.totalBtl || 0;
                        if (d < fromDate) {
                            tpBefore[sz] = (tpBefore[sz] || 0) + btl;
                        } else if (d <= toDate) {
                            if (!tpByDate[d]) tpByDate[d] = [];
                            let e = tpByDate[d].find(t => t.tpNumber === tp.tpNumber);
                            if (!e) { e = { tpNumber: tp.tpNumber, cols: {} }; tpByDate[d].push(e); }
                            e.cols[sz] = (e.cols[sz] || 0) + btl;
                        }
                    }
                }
            }

            const saleByDate = {};
            const saleBefore = {};
            for (const sale of (salesItems || [])) {
                if (!isMmlItem(sale)) continue;
                const d = sale.billDate || '', sz = mcSizeGrp(sale.size), btl = sale.totalBtl || 0;
                sizeSet.add(sz);
                if (d < fromDate) saleBefore[sz] = (saleBefore[sz]||0)+btl;
                else if (d <= toDate) { if (!saleByDate[d]) saleByDate[d]={}; saleByDate[d][sz]=(saleByDate[d][sz]||0)+btl; }
            }

            const fyOpen = {};
            if (osResult.success && osResult.data && osResult.data.entries) {
                for (const e of osResult.data.entries) {
                    if (!isMmlItem(e)) continue;
                    const sz = mcSizeGrp(e.size); sizeSet.add(sz);
                    fyOpen[sz] = (fyOpen[sz]||0) + (e.totalBtl||0);
                }
            }

            const sizes = [...sizeSet].filter(Boolean).sort((a,b) => mcSizeML(b) - mcSizeML(a));
            mcLastSizes = sizes;

            const openAtStart = {};
            for (const sz of sizes) openAtStart[sz] = Math.max(0, (fyOpen[sz]||0)+(tpBefore[sz]||0)-(saleBefore[sz]||0));

            // Build date rows
            const dates = [];
            let cur = new Date(fromDate), end = new Date(toDate);
            while (cur <= end) { dates.push(cur.toISOString().slice(0,10)); cur.setDate(cur.getDate()+1); }

            const rows = []; let prev = { ...openAtStart };
            for (const date of dates) {
                const tpEntries = tpByDate[date] || [], saleSz = saleByDate[date] || {};
                const purSz = {};
                for (const tp of tpEntries) for (const [sz,btl] of Object.entries(tp.cols)) purSz[sz]=(purSz[sz]||0)+btl;
                const tpNos = [...new Set(tpEntries.map(t => { const i=(t.tpNumber||'').lastIndexOf('/'); return i>=0?t.tpNumber.slice(i+1):(t.tpNumber||''); }))].filter(Boolean).join(', ');
                const opening={}, purchase={}, sale={}, closing={};
                for (const sz of sizes) {
                    opening[sz]=prev[sz]||0; purchase[sz]=purSz[sz]||0; sale[sz]=saleSz[sz]||0;
                    closing[sz]=Math.max(0, opening[sz]+purchase[sz]-sale[sz]);
                }
                const hasAny = sizes.some(sz => opening[sz]||purchase[sz]||sale[sz]);
                if (hasAny || rows.length>0) rows.push({date,tpNos,opening,purchase,sale,closing});
                prev={}; for (const sz of sizes) prev[sz]=closing[sz]||0;
            }

            mcLastRows = rows;

            // Summary cards
            const sumOpen  = sizes.reduce((a,sz)=>a+(rows[0]?rows[0].opening[sz]||0:0),0);
            const sumPur   = rows.reduce((a,r)=>a+sizes.reduce((b,sz)=>b+(r.purchase[sz]||0),0),0);
            const sumSale  = rows.reduce((a,r)=>a+sizes.reduce((b,sz)=>b+(r.sale[sz]||0),0),0);
            const sumClose = sizes.reduce((a,sz)=>a+(rows.length?rows[rows.length-1].closing[sz]||0:0),0);
            if (mcTotalDays)     mcTotalDays.textContent     = rows.length;
            if (mcTotalOpening)  mcTotalOpening.textContent  = sumOpen.toLocaleString('en-IN');
            if (mcTotalPurchase) mcTotalPurchase.textContent = sumPur.toLocaleString('en-IN');
            if (mcTotalSale)     mcTotalSale.textContent     = sumSale.toLocaleString('en-IN');
            if (mcTotalClosing)  mcTotalClosing.textContent  = sumClose.toLocaleString('en-IN');

            if (rows.length === 0) {
                if (mcEmpty) mcEmpty.classList.remove('hidden');
                [mcTableHead,mcTableBody,mcTableFoot].forEach(el => { if(el) el.innerHTML=''; });
                return;
            }
            if (mcEmpty) mcEmpty.classList.add('hidden');

            const n = sizes.length;
            const secBg = { OPENING:'#fee2e2', PURCHASE:'#dcfce7', SALES:'#fee2e2', CLOSING:'#fef3c7' };
            const secTxt= { OPENING:'#991b1b', PURCHASE:'#166534', SALES:'#991b1b', CLOSING:'#92400e' };
            const isDark = document.documentElement.getAttribute('data-theme')==='dark';
            const mmlBg  = isDark ? 'rgba(248,113,113,.14)' : '#fee2e2';
            const mmlTxt = isDark ? '#fca5a5' : '#991b1b';

            // 3-row header — PURCHASE | SALES | CLOSING (OPENING shown as first body row)
            let h1=`<th rowspan="3" class="cr-th-date">Date</th><th rowspan="3" class="cr-th-tp">TP No.</th>`;
            for (const [lbl,sec] of [['PURCHASE','PURCHASE'],['SALES','SALES'],['CLOSING','CLOSING']]) {
                h1+=`<th colspan="${n}" style="background:${isDark?'rgba(248,113,113,.22)':secBg[sec]};color:${isDark?'#fca5a5':secTxt[sec]};text-align:center;font-size:.8rem;font-weight:800;letter-spacing:.04em">${lbl}</th>`;
            }
            let h2='';
            for (let s=0;s<3;s++) h2+=`<th colspan="${n}" style="background:${mmlBg};color:${mmlTxt};text-align:center;font-size:.72rem;font-weight:700">MML</th>`;
            let h3='';
            for (let s=0;s<3;s++) for (const sz of sizes) h3+=`<th style="font-size:.65rem;padding:2px 4px;min-width:36px">${sz.replace(/\s*ML$/i,'').replace(/\s*Ltr$/i,'L')}</th>`;
            if (mcTableHead) mcTableHead.innerHTML=`<tr>${h1}</tr><tr>${h2}</tr><tr>${h3}</tr>`;

            const fq = v => v ? v.toLocaleString('en-IN') : '';
            const openBgMc  = isDark ? 'rgba(248,113,113,.12)' : '#fff1f2';
            const openTxtMc = isDark ? '#fca5a5' : '#991b1b';
            // Opening stock row
            let body=`<tr style="background:${openBgMc};font-weight:600">`;
            body+=`<td class="cr-td-date" style="color:${openTxtMc}">Opening Stock</td><td class="cr-td-tp" style="color:${openTxtMc}">${mcFmt(rows[0].date)}</td>`;
            for (const sz of sizes) body+=`<td style="color:${openTxtMc}">${fq(rows[0].opening[sz])}</td>`;
            body+=`<td colspan="${n*2}"></td></tr>`;
            // Date rows — PURCHASE | SALES | CLOSING only
            for (const r of rows) {
                body+=`<tr><td class="cr-td-date">${mcFmt(r.date)}</td><td class="cr-td-tp">${r.tpNos||'\u2014'}</td>`;
                for (const fld of ['purchase','sale','closing']) for (const sz of sizes) body+=`<td>${fq(r[fld][sz])}</td>`;
                body+='</tr>';
            }
            if (mcTableBody) mcTableBody.innerHTML=body;

            let foot=`<tr class="cr-total-row"><td class="cr-td-date" colspan="2"><strong>TOTAL</strong></td>`;
            for (const fld of ['purchase','sale','closing']) for (const sz of sizes) {
                const v = fld==='closing'?(rows.length?rows[rows.length-1].closing[sz]||0:0):rows.reduce((a,r)=>a+(r[fld][sz]||0),0);
                foot+=`<td><strong>${fq(v)}</strong></td>`;
            }
            foot+='</tr>';
            if (mcTableFoot) mcTableFoot.innerHTML=foot;
        }

        /* ── Export CSV ── */
        function exportMcCSV() {
            if (!mcLastRows.length) return;
            const sz = mcLastSizes;
            const hdr = ['Date','TP No.',...sz.map(s=>'Open-MML-'+s),...sz.map(s=>'Pur-MML-'+s),...sz.map(s=>'Sale-MML-'+s),...sz.map(s=>'Close-MML-'+s)];
            const lines = [hdr.join(',')];
            for (const r of mcLastRows) {
                const line=[mcFmt(r.date),'"'+r.tpNos+'"'];
                for (const fld of ['opening','purchase','sale','closing']) for (const s of sz) line.push(r[fld][s]||0);
                lines.push(line.join(','));
            }
            const blob = new Blob(['\uFEFF'+lines.join('\n')],{type:'text/csv;charset=utf-8;'});
            const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
            a.download=`MMLChatai_${activeBar.barName||'Bar'}_${mcDateFrom?.value}_to_${mcDateTo?.value}.csv`;
            a.click(); URL.revokeObjectURL(a.href);
        }

        /* ── Print ── */
        function printMcReport() {
            if (!mcLastRows.length) return;
            const sz=mcLastSizes, n=sz.length;
            const barName=activeBar.barName||'', fStr=mcDateFrom?mcDateFrom.value:'', tStr=mcDateTo?mcDateTo.value:'';
            let th1=`<th rowspan="3">Date</th><th rowspan="3">TP No.</th>`;
            for (const lbl of ['PURCHASE','SALES','CLOSING']) th1+=`<th colspan="${n}" style="background:${lbl==='PURCHASE'?'#dcfce7':lbl==='SALES'?'#fee2e2':'#fef3c7'};text-align:center">${lbl}</th>`;
            let th2=''; for (let s=0;s<3;s++) th2+=`<th colspan="${n}" style="background:#fee2e2;text-align:center;font-size:7pt">MML</th>`;
            let th3=''; for (let s=0;s<3;s++) for (const s2 of sz) th3+=`<th>${s2.replace(/\s*ML$/i,'')}</th>`;
            const fq=v=>v?v.toLocaleString('en-IN'):'';
            // Opening stock row
            let tbody=`<tr style="background:#fff1f2;font-weight:600">`;
            tbody+=`<td style="color:#991b1b">Opening Stock</td><td style="color:#991b1b;font-size:6pt">${mcFmt(mcLastRows[0].date)}</td>`;
            for (const s of sz) tbody+=`<td style="color:#991b1b">${fq(mcLastRows[0].opening[s])}</td>`;
            tbody+=`<td colspan="${n*2}"></td></tr>`;
            // Date rows — PURCHASE | SALES | CLOSING only
            for (const r of mcLastRows) {
                tbody+=`<tr><td>${mcFmt(r.date)}</td><td style="font-size:6pt">${r.tpNos||'\u2014'}</td>`;
                for (const fld of ['purchase','sale','closing']) for (const s of sz) tbody+=`<td>${fq(r[fld][s])}</td>`;
                tbody+='</tr>';
            }
            tbody+=`<tr style="font-weight:700"><td colspan="2">TOTAL</td>`;
            for (const fld of ['purchase','sale','closing']) for (const s of sz) {
                const v=fld==='closing'?(mcLastRows.length?mcLastRows[mcLastRows.length-1].closing[s]||0:0):mcLastRows.reduce((a,r)=>a+(r[fld][s]||0),0);
                tbody+=`<td>${fq(v)}</td>`;
            }
            tbody+='</tr>';
            printWithIframe(`<!DOCTYPE html><html><head><title>MML Chatai — ${barName}</title>
                <style>body{font-family:'Inter',Arial,sans-serif;padding:10px;font-size:7pt}
                h2{margin:0 0 4px;font-size:12pt} .meta{color:#555;font-size:7.5pt;margin-bottom:6px}
                table{border-collapse:collapse;width:100%} th,td{border:1px solid #bbb;padding:2px 3px;font-size:6.5pt;text-align:center}
                th{background:#f0f0f0} td:first-child,td:nth-child(2),th:first-child,th:nth-child(2){text-align:left}
                @page{size:landscape;margin:6mm}</style>
            </head><body>
                <h2>&#127991; MML Chatai — ${barName}</h2>
                <div class="meta">Period: <strong>${mcFmt(fStr)}</strong> to <strong>${mcFmt(tStr)}</strong> &nbsp;|&nbsp; Generated: ${new Date().toLocaleString('en-IN')}</div>
                <table><thead><tr>${th1}</tr><tr>${th2}</tr><tr>${th3}</tr></thead><tbody>${tbody}</tbody></table>
            </body></html>`);
        }

        /* ── Event wiring ── */
        if (mcGoBtn)     mcGoBtn.addEventListener('click', generateMmlChatai);
        if (mcExportBtn) mcExportBtn.addEventListener('click', exportMcCSV);
        if (mcPrintBtn)  mcPrintBtn.addEventListener('click', printMcReport);
        if (mcDateFrom)  mcDateFrom.addEventListener('change', generateMmlChatai);
        if (mcDateTo)    mcDateTo.addEventListener('change',   generateMmlChatai);
        [mcDateFrom, mcDateTo].forEach(el => {
            el?.addEventListener('keydown', e => { if (e.key==='Enter') { e.preventDefault(); generateMmlChatai(); } });
        });
    }

    // ═══════════════════════════════════════════════════════
    //  OPENING STOCK — FY Beginning Inventory
    // ═══════════════════════════════════════════════════════

    const osGridBody  = document.getElementById('osGridBody');
    const osDateEl    = document.getElementById('osDate');
    const osSaveBtn   = document.getElementById('osSaveBtn');
    const osClearBtn  = document.getElementById('osClearBtn');

    // Default as-of date = April 1 of current FY start
    const fyStart = (() => {
        const now = new Date();
        const year = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
        return `${year}-04-01`;
    })();
    if (osDateEl) osDateEl.value = fyStart;

    let osRowCount = 0;
    let osLoadedId = null; // track if editing existing record

    /* ── Create OS row (reuses TP grid patterns) ── */
    function createOsRow(data) {
        osRowCount++;
        const tr = document.createElement('tr');
        tr.dataset.row = osRowCount;

        const sizeOpts = TP_SIZES.map(s =>
            `<option value="${s.value}" ${data && data.size === s.value ? 'selected' : ''}>${s.label}</option>`
        ).join('');

        tr.innerHTML = `
            <td class="gc-sr">${osRowCount}</td>
            <td class="gc-brand">
                <input type="text" class="tp-grid-input text-left" data-col="brand" autocomplete="off" placeholder="Brand or shortcode…" value="${data ? esc(data.brandName || '') : ''}">
                <div class="tp-prod-ac hidden" data-role="prod-ac"></div>
            </td>
            <td class="gc-code">
                <div class="tp-cell-readonly" data-col="code" style="justify-content:flex-start">${data && data.code ? esc(data.code) : '—'}</div>
            </td>
            <td class="gc-size">
                <select class="tp-size-select" data-col="size">${sizeOpts}</select>
            </td>
            <td class="gc-cases">
                <input type="number" class="tp-grid-input text-right" data-col="cases" min="0" placeholder="0" value="${data ? (data.cases || '') : ''}">
            </td>
            <td class="gc-bpc">
                <div class="tp-cell-readonly" data-col="bpc">—</div>
            </td>
            <td class="gc-loose">
                <input type="number" class="tp-grid-input text-right" data-col="loose" min="0" placeholder="0" value="${data ? (data.loose || '') : ''}">
            </td>
            <td class="gc-total">
                <div class="tp-cell-readonly" data-col="totalBtl">0</div>
            </td>
            <td class="gc-rate">
                <input type="number" class="tp-grid-input text-right" data-col="rate" min="0" step="0.01" placeholder="0.00" value="${data ? (data.rate || '') : ''}">
            </td>
            <td class="gc-amount">
                <div class="tp-cell-readonly" data-col="amount">0.00</div>
            </td>
        `;

        osGridBody.appendChild(tr);
        wireOsGridRow(tr);
        wireOsProdAc(tr);

        // If we have data, store productId
        if (data && data.productId) tr.dataset.productId = data.productId;

        // Recalc to show loaded values
        recalcOsRow(tr);
        return tr;
    }

    /* ── Get interactive fields for a row ── */
    function getOsRowFields(tr) {
        return Array.from(tr.querySelectorAll('input.tp-grid-input, select.tp-size-select'));
    }

    /* ── Recalculate one OS row ── */
    function recalcOsRow(tr) {
        const sizeEl  = tr.querySelector('[data-col="size"]');
        const casesEl = tr.querySelector('[data-col="cases"]');
        const looseEl = tr.querySelector('[data-col="loose"]');
        const rateEl  = tr.querySelector('[data-col="rate"]');
        const bpcEl   = tr.querySelector('[data-col="bpc"]');
        const totalEl = tr.querySelector('[data-col="totalBtl"]');
        const amtEl   = tr.querySelector('[data-col="amount"]');

        const sizeVal = sizeEl ? sizeEl.value : '';
        const sizeObj = TP_SIZES.find(s => s.value === sizeVal) || TP_SIZES[0];
        const bpc     = sizeObj.bpc;
        const cases   = parseInt(casesEl?.value) || 0;
        const loose   = parseInt(looseEl?.value) || 0;
        const rate    = parseFloat(rateEl?.value) || 0;
        const totalBtl = (cases * bpc) + loose;
        const amount   = totalBtl * rate;

        if (bpcEl)   { bpcEl.textContent = bpc > 0 ? bpc : '—'; bpcEl.classList.toggle('has-value', bpc > 0); }
        if (totalEl) { totalEl.textContent = totalBtl; totalEl.classList.toggle('has-value', totalBtl > 0); }
        if (amtEl)   { amtEl.textContent = amount > 0 ? amount.toFixed(2) : '0.00'; amtEl.classList.toggle('has-value', amount > 0); }

        recalcOsTotals();
    }

    /* ── Recalculate OS footer totals + summary cards ── */
    function recalcOsTotals() {
        let totalCases = 0, totalLoose = 0, totalBtl = 0, totalVal = 0, productCount = 0;
        osGridBody.querySelectorAll('tr').forEach(tr => {
            const brand = tr.querySelector('[data-col="brand"]')?.value.trim();
            const sizeEl = tr.querySelector('[data-col="size"]');
            const sizeVal = sizeEl ? sizeEl.value : '';
            const sizeObj = TP_SIZES.find(s => s.value === sizeVal) || TP_SIZES[0];
            const cases = parseInt(tr.querySelector('[data-col="cases"]')?.value) || 0;
            const loose = parseInt(tr.querySelector('[data-col="loose"]')?.value) || 0;
            const rate  = parseFloat(tr.querySelector('[data-col="rate"]')?.value) || 0;
            const btl   = (cases * sizeObj.bpc) + loose;
            totalCases += cases;
            totalLoose += loose;
            totalBtl   += btl;
            totalVal   += btl * rate;
            if (brand) productCount++;
        });

        const elCases   = document.getElementById('osTotalCases');
        const elLoose   = document.getElementById('osTotalLoose');
        const elBottles = document.getElementById('osTotalBottles');
        const elValue   = document.getElementById('osTotalValue');
        if (elCases)   elCases.textContent   = totalCases;
        if (elLoose)   elLoose.textContent   = totalLoose;
        if (elBottles) elBottles.textContent = totalBtl;
        if (elValue)   elValue.textContent   = '₹ ' + totalVal.toFixed(2);

        // Summary cards
        const sumProd = document.getElementById('osSumProducts');
        const sumCase = document.getElementById('osSumCases');
        const sumBtl  = document.getElementById('osSumBottles');
        const sumVal  = document.getElementById('osSumValue');
        if (sumProd)  sumProd.textContent = productCount;
        if (sumCase)  sumCase.textContent = totalCases;
        if (sumBtl)   sumBtl.textContent  = totalBtl;
        if (sumVal)   sumVal.textContent   = '₹ ' + totalVal.toFixed(2);
    }

    /* ── Wire keyboard navigation for OS grid ── */
    function wireOsGridRow(tr) {
        const fields = getOsRowFields(tr);

        fields.forEach(field => {
            field.addEventListener('input', () => recalcOsRow(tr));
            field.addEventListener('change', () => recalcOsRow(tr));

            field.addEventListener('focus', () => {
                osGridBody.querySelectorAll('tr').forEach(r => r.classList.remove('tp-row-active'));
                tr.classList.add('tp-row-active');
            });

            field.addEventListener('keydown', (e) => {
                const col = field.dataset.col || field.getAttribute('data-col');

                // Defer to product AC when open
                if (col === 'brand' && osActiveProdAc && osActiveProdAc.tr === tr) {
                    const acDrop = osActiveProdAc.dropdown;
                    if (acDrop && !acDrop.classList.contains('hidden')) {
                        if (['ArrowDown', 'ArrowUp', 'Escape'].includes(e.key)) return;
                        if (e.key === 'Enter') return;
                    }
                }

                const allRows = Array.from(osGridBody.querySelectorAll('tr'));
                const rowIdx  = allRows.indexOf(tr);
                const fIdx    = fields.indexOf(field);

                if (e.key === 'Enter') {
                    e.preventDefault();
                    if (fIdx < fields.length - 1) {
                        fields[fIdx + 1].focus();
                    } else {
                        let nextRow = allRows[rowIdx + 1];
                        if (!nextRow) nextRow = createOsRow();
                        const nf = getOsRowFields(nextRow);
                        if (nf[0]) nf[0].focus();
                    }
                } else if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    let nextRow = allRows[rowIdx + 1];
                    if (!nextRow) nextRow = createOsRow();
                    const target = nextRow.querySelector(`[data-col="${col}"]`);
                    if (target && (target.tagName === 'INPUT' || target.tagName === 'SELECT')) target.focus();
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    if (rowIdx > 0) {
                        const prevRow = allRows[rowIdx - 1];
                        const target = prevRow.querySelector(`[data-col="${col}"]`);
                        if (target && (target.tagName === 'INPUT' || target.tagName === 'SELECT')) target.focus();
                    }
                } else if (e.key === 'Escape') {
                    e.preventDefault();
                    field.blur();
                } else if (e.key === 'Delete' && e.ctrlKey) {
                    e.preventDefault();
                    if (allRows.length > 1) {
                        const newFocusRow = allRows[rowIdx - 1] || allRows[rowIdx + 1];
                        tr.remove();
                        renumberOsRows();
                        recalcOsTotals();
                        if (newFocusRow) {
                            const f = getOsRowFields(newFocusRow);
                            if (f[0]) f[0].focus();
                        }
                    }
                }
            });
        });
    }

    function renumberOsRows() {
        osGridBody.querySelectorAll('tr').forEach((tr, i) => {
            const srCell = tr.querySelector('.gc-sr');
            if (srCell) srCell.textContent = i + 1;
            tr.dataset.row = i + 1;
        });
        osRowCount = osGridBody.querySelectorAll('tr').length;
    }

    /* ── Product Autocomplete for OS grid (mirrors TP pattern) ── */
    let osActiveProdAc = null;

    function wireOsProdAc(tr) {
        const brandInput = tr.querySelector('[data-col="brand"]');
        const dropdown   = tr.querySelector('[data-role="prod-ac"]');
        if (!brandInput || !dropdown) return;

        let acIdx = -1;

        function showAc(query) {
            const q = (query || '').toLowerCase();

            // --- Shortcut matches first ---
            const scMatches = getShortcutMatches(query);
            const scProducts = scMatches.map(sc => {
                const prod = allProducts.find(p => p.id === sc.productId);
                return prod ? { ...prod, _shortcutCode: sc.code } : {
                    id: sc.productId, brandName: sc.brandName, code: sc.prodCode,
                    size: sc.size, category: sc.category, mrp: sc.mrp, costPrice: 0,
                    _shortcutCode: sc.code
                };
            });
            const scProdIds = new Set(scProducts.map(p => p.id));

            const textMatches = q.length > 0
                ? allProducts.filter(p =>
                    !scProdIds.has(p.id) && (
                    (p.brandName || '').toLowerCase().includes(q) ||
                    (p.code || '').toLowerCase().includes(q) ||
                    (p.category || '').toLowerCase().includes(q))
                  )
                : allProducts.filter(p => !scProdIds.has(p.id)).slice(0, 30);

            const combined = [...scProducts, ...textMatches];

            if (combined.length === 0 && q.length > 0) {
                dropdown.innerHTML = '<div class="tp-prod-ac-empty">No matching products</div>';
            } else if (combined.length === 0) {
                dropdown.innerHTML = '<div class="tp-prod-ac-empty">No products in master</div>';
            } else {
                dropdown.innerHTML = combined.map((p, i) => `
                    <div class="tp-prod-ac-item${p._shortcutCode ? ' sc-match' : ''}" data-idx="${i}">
                        ${p._shortcutCode ? '<span class="tp-prod-ac-sc-badge">' + esc(p._shortcutCode) + '</span>' : ''}
                        <span class="tp-prod-ac-brand">${esc(p.brandName)}</span>
                        <span class="tp-prod-ac-meta">${esc(p.code || '')} · ${p.size ? p.size + 'ml' : ''} · ${esc(p.category || '')}</span>
                    </div>
                `).join('');

                dropdown.querySelectorAll('.tp-prod-ac-item').forEach(item => {
                    item.addEventListener('mousedown', (e) => {
                        e.preventDefault();
                        const idx = parseInt(item.dataset.idx);
                        selectOsProdItem(tr, combined[idx], brandInput, dropdown);
                    });
                });
            }

            dropdown.classList.remove('hidden');
            acIdx = -1;
            osActiveProdAc = { tr, dropdown, focusIdx: -1, combined };
        }

        function hideAc() {
            dropdown.classList.add('hidden');
            dropdown.innerHTML = '';
            acIdx = -1;
            if (osActiveProdAc && osActiveProdAc.tr === tr) osActiveProdAc = null;
        }

        function focusItem(newIdx) {
            const items = dropdown.querySelectorAll('.tp-prod-ac-item');
            if (items.length === 0) return;
            acIdx = Math.max(0, Math.min(newIdx, items.length - 1));
            items.forEach((it, i) => it.classList.toggle('focused', i === acIdx));
            items[acIdx]?.scrollIntoView({ block: 'nearest' });
        }

        brandInput.addEventListener('input', () => showAc(brandInput.value));

        brandInput.addEventListener('focus', () => {
            if (osActiveProdAc && osActiveProdAc.tr !== tr) {
                osActiveProdAc.dropdown.classList.add('hidden');
                osActiveProdAc = null;
            }
            if (allProducts.length > 0) showAc(brandInput.value);
        });

        brandInput.addEventListener('blur', () => {
            setTimeout(() => hideAc(), 150);
        });

        brandInput.addEventListener('keydown', (e) => {
            /* ── Shortcut quick-select: Enter/Tab on exact shortcode match ── */
            if (e.key === 'Enter' || e.key === 'Tab') {
                const val = brandInput.value.trim();
                if (val && acIdx < 0) {
                    const scProd = resolveShortcutToProduct(val);
                    if (scProd) {
                        e.preventDefault(); e.stopPropagation();
                        selectOsProdItem(tr, scProd, brandInput, dropdown);
                        return;
                    }
                }
            }

            if (dropdown.classList.contains('hidden')) return;
            const items = dropdown.querySelectorAll('.tp-prod-ac-item');

            if (e.key === 'ArrowDown') {
                e.preventDefault(); e.stopPropagation();
                focusItem(acIdx + 1);
            } else if (e.key === 'ArrowUp') {
                e.preventDefault(); e.stopPropagation();
                focusItem(acIdx - 1);
            } else if (e.key === 'Enter' && acIdx >= 0 && items[acIdx]) {
                e.preventDefault(); e.stopPropagation();
                const combined = osActiveProdAc?.combined;
                if (combined && combined[acIdx]) selectOsProdItem(tr, combined[acIdx], brandInput, dropdown);
            } else if (e.key === 'Escape') {
                e.preventDefault(); e.stopPropagation();
                hideAc();
            }
        });
    }

    function selectOsProdItem(tr, product, brandInput, dropdown) {
        brandInput.value = product.brandName || '';
        tr.dataset.productId = product.id || '';

        const codeEl = tr.querySelector('[data-col="code"]');
        if (codeEl) { codeEl.textContent = product.code || '—'; codeEl.classList.toggle('has-value', !!product.code); }

        const sizeEl = tr.querySelector('[data-col="size"]');
        if (sizeEl && product.size) sizeEl.value = product.size;

        const rateEl = tr.querySelector('[data-col="rate"]');
        if (rateEl) {
            const rate = product.costPrice || product.mrp || 0;
            if (rate > 0) rateEl.value = rate;
        }

        recalcOsRow(tr);
        dropdown.classList.add('hidden');
        dropdown.innerHTML = '';
        if (osActiveProdAc && osActiveProdAc.tr === tr) osActiveProdAc = null;

        const casesEl = tr.querySelector('[data-col="cases"]');
        if (casesEl) casesEl.focus();
    }

    /* ── Load existing opening stock ── */
    async function loadOpeningStock() {
        if (!activeBar.barName) return;
        try {
            const result = await window.electronAPI.getOpeningStock({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || ''
            });
            if (result.success && result.data) {
                osLoadedId = result.data.id || null;
                if (osDateEl && result.data.asOfDate) osDateEl.value = result.data.asOfDate;
                osGridBody.innerHTML = '';
                osRowCount = 0;
                const entries = result.data.entries || [];
                if (entries.length > 0) {
                    entries.forEach(entry => createOsRow(entry));
                }
                // Pad to at least 8 rows
                const remaining = Math.max(0, 8 - entries.length);
                for (let i = 0; i < remaining; i++) createOsRow();
                recalcOsTotals();
            } else {
                // No existing data, seed empty rows
                clearOsForm(true);
            }
        } catch (err) {
            console.error('loadOpeningStock error:', err);
            clearOsForm(true);
        }
    }

    /* ── Save opening stock ── */
    async function saveOpeningStock() {
        const asOfDate = osDateEl?.value;
        if (!asOfDate) {
            osDateEl?.focus();
            showTpToast('As-of date is required', true);
            return;
        }

        const entries = [];
        osGridBody.querySelectorAll('tr').forEach(tr => {
            const brand = tr.querySelector('[data-col="brand"]')?.value.trim();
            if (!brand) return;
            const code  = tr.querySelector('[data-col="code"]')?.textContent.trim() || '';
            const sizeEl = tr.querySelector('[data-col="size"]');
            const sizeObj = TP_SIZES.find(s => s.value === (sizeEl?.value || '')) || TP_SIZES[0];
            const cases = parseInt(tr.querySelector('[data-col="cases"]')?.value) || 0;
            const loose = parseInt(tr.querySelector('[data-col="loose"]')?.value) || 0;
            const rate  = parseFloat(tr.querySelector('[data-col="rate"]')?.value) || 0;
            const totalBtl = (cases * sizeObj.bpc) + loose;

            entries.push({
                productId: tr.dataset.productId || '',
                brandName: brand,
                code: code !== '—' ? code : '',
                category: allProducts.find(p => p.id === tr.dataset.productId)?.category || '',
                size: sizeEl?.value || '',
                sizeLabel: sizeObj.label,
                cases,
                bpc: sizeObj.bpc,
                loose,
                totalBtl,
                rate,
                value: totalBtl * rate,
            });
        });

        if (entries.length === 0) {
            const firstBrand = osGridBody.querySelector('[data-col="brand"]');
            if (firstBrand) firstBrand.focus();
            showTpToast('Add at least one product with stock', true);
            return;
        }

        const osData = {
            id: osLoadedId || ('OS_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6)),
            asOfDate,
            entries,
            totalCases:   entries.reduce((s, e) => s + e.cases, 0),
            totalLoose:   entries.reduce((s, e) => s + e.loose, 0),
            totalBottles: entries.reduce((s, e) => s + e.totalBtl, 0),
            totalValue:   entries.reduce((s, e) => s + e.value, 0),
            productCount: entries.length,
            updatedAt: new Date().toISOString(),
        };

        try {
            const result = await window.electronAPI.saveOpeningStock({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || '',
                openingStock: osData,
            });
            if (result.success) {
                osLoadedId = osData.id;
                showTpToast(`Opening stock saved — ${entries.length} products, ₹${osData.totalValue.toFixed(2)}`);
            } else {
                showTpToast('Save failed: ' + result.error, true);
            }
        } catch (err) {
            showTpToast('Save error: ' + err.message, true);
        }
    }

    /* ── Clear OS form ── */
    function clearOsForm(skipConfirm) {
        if (!skipConfirm) {
            const proceed = confirm('Clear all opening stock entries?');
            if (!proceed) return;
        }
        osGridBody.innerHTML = '';
        osRowCount = 0;
        osLoadedId = null;
        if (osDateEl) osDateEl.value = fyStart;
        for (let i = 0; i < 8; i++) createOsRow();
        recalcOsTotals();
    }

    /* ── Wire buttons ── */
    if (osSaveBtn)  osSaveBtn.addEventListener('click', saveOpeningStock);
    if (osClearBtn) osClearBtn.addEventListener('click', () => clearOsForm(false));

    /* ── Global shortcuts for OS panel ── */
    document.addEventListener('keydown', (e) => {
        const osPanel = document.getElementById('sub-opening-stock');
        if (!osPanel || !osPanel.classList.contains('active')) return;

        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
            e.preventDefault();
            saveOpeningStock();
            return;
        }
    });

    /* ── Init: seed empty rows (loadOpeningStock called later after loadProducts) ── */
    for (let i = 0; i < 8; i++) createOsRow();

    // ═══════════════════════════════════════════════════════
    //  PRODUCT MASTER — List + Form (Tally-style)
    // ═══════════════════════════════════════════════════════

    const prodListBody    = document.getElementById('prodListBody');
    const prodEmptyState  = document.getElementById('prodEmptyState');
    const prodCountEl     = document.getElementById('prodCount');
    const prodSearchEl    = document.getElementById('prodSearch');
    const prodFormPanel   = document.getElementById('prodFormPanel');
    const prodFormEl      = document.getElementById('prodForm');
    const prodPlaceholder = document.getElementById('prodPlaceholder');
    const prodNewBtn      = document.getElementById('prodNewBtn');
    const prodSaveBtn     = document.getElementById('prodSaveBtn');
    const prodCancelBtn   = document.getElementById('prodCancelBtn');
    const prodDeleteBtn   = document.getElementById('prodDeleteBtn');
    const prodFlowFields  = Array.from(document.querySelectorAll('.prod-flow'));
    const prodSizeSelect  = document.getElementById('prodSize');
    const prodBpcInput    = document.getElementById('prodBpc');

    // allProducts is declared above (before TP section) for shared access
    let filteredProds  = [];
    let prodFocusedIdx = -1;
    let prodEditingId  = null;

    /* ── Populate Product Master size dropdown from ALL_SIZES ── */
    if (prodSizeSelect) {
        prodSizeSelect.innerHTML = '<option value="">— Select —</option>' +
            ALL_SIZES.filter(s => s.value).map(s =>
                `<option value="${s.value}">${s.label}</option>`
            ).join('');
    }

    /* ── Auto-fill BPC when size changes ── */
    if (prodSizeSelect) {
        prodSizeSelect.addEventListener('change', () => {
            const entry = SIZE_LOOKUP.get(prodSizeSelect.value);
            if (prodBpcInput) prodBpcInput.value = entry ? entry.bpc : '';
        });
    }

    /* ── Load products from disk ── */
    /* ── Normalize legacy size values (e.g. "9096" → "90 ML (96)") ── */
    function normalizeLegacySize(sizeStr) {
        if (!sizeStr) return '';
        // Already in correct format (contains "ML" or "Ltr")
        if (/ml|ltr/i.test(sizeStr)) return sizeStr;
        // If it's in ALL_SIZES as-is, keep it
        if (SIZE_LOOKUP.has(sizeStr)) return sizeStr;
        // Pure numeric — could be simple (e.g. "750") or concatenated (e.g. "9096")
        const digits = sizeStr.replace(/[^0-9]/g, '');
        if (!digits) return sizeStr;
        // Try known ML values from longest to shortest to split concatenated strings
        const knownMLs = [50000, 30000, 20000, 15000, 4500, 2000, 1750, 1500, 1000, 750, 700, 650, 500, 375, 350, 330, 275, 250, 200, 187, 180, 125, 90, 60, 50];
        for (const ml of knownMLs) {
            const mlStr = String(ml);
            if (digits.startsWith(mlStr)) {
                const remainder = digits.slice(mlStr.length);
                if (remainder === '') {
                    // Simple numeric like "750" → find plain ALL_SIZES entry
                    const match = ALL_SIZES.find(s => s.ml === ml && s.value && !s.value.includes('('));
                    return match ? match.value : (ml >= 10000 ? (ml / 1000) + ' Ltr' : ml + ' ML');
                } else {
                    // Concatenated like "9096" → "90 ML (96)"
                    const bpcNum = parseInt(remainder);
                    if (bpcNum > 0) {
                        // Try exact match in ALL_SIZES
                        const match = ALL_SIZES.find(s => s.ml === ml && s.bpc === bpcNum);
                        if (match) return match.value;
                        // Construct format
                        return ml + ' ML (' + bpcNum + ')';
                    }
                }
            }
        }
        return sizeStr; // unchanged if no pattern matches
    }

    async function loadProducts() {
        if (!activeBar.barName) return;
        let needsSave = false;
        try {
            const result = await window.electronAPI.getProducts({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || ''
            });
            if (result.success) {
                allProducts = result.products || [];
                // Auto-migrate legacy size values
                allProducts.forEach(p => {
                    if (p.size && !/ml|ltr/i.test(p.size)) {
                        const fixed = normalizeLegacySize(p.size);
                        if (fixed !== p.size) {
                            p.size = fixed;
                            // Also fix BPC from ALL_SIZES
                            const entry = SIZE_LOOKUP.get(fixed);
                            if (entry) p.bpc = entry.bpc;
                            needsSave = true;
                        }
                    }
                });
                // Persist the migration
                if (needsSave) {
                    try {
                        for (const p of allProducts) {
                            await window.electronAPI.saveProduct({
                                barName: activeBar.barName,
                                financialYear: activeBar.financialYear || '',
                                product: p,
                            });
                        }
                    } catch (_) { /* silent migration save */ }
                }
            }
        } catch (err) {
            console.error('loadProducts error:', err);
        }
        filteredProds = [...allProducts];
        renderProductList();
    }

    /* ── Render product list ── */
    function renderProductList() {
        prodListBody.querySelectorAll('.prod-row').forEach(r => r.remove());

        if (filteredProds.length === 0) {
            if (prodEmptyState) prodEmptyState.style.display = '';
            if (prodCountEl) prodCountEl.textContent = `${allProducts.length} product${allProducts.length !== 1 ? 's' : ''}`;
            return;
        }
        if (prodEmptyState) prodEmptyState.style.display = 'none';
        if (prodCountEl) prodCountEl.textContent = `${filteredProds.length} product${filteredProds.length !== 1 ? 's' : ''}`;

        filteredProds.forEach((prod, idx) => {
            const row = document.createElement('div');
            row.className = 'prod-row';
            row.tabIndex = 0;
            row.dataset.idx = idx;
            row.dataset.id = prod.id;
            row.innerHTML = `
                <span class="prod-row-name">${esc(prod.brandName)}</span>
                <span class="prod-row-cat">${esc(prod.category || '—')}</span>
                <span class="prod-row-size">${prod.size ? esc(prod.size) : '—'}</span>
            `;

            row.addEventListener('click', () => {
                setProdFocus(idx);
                openProductForEdit(prod);
            });
            row.addEventListener('dblclick', () => openProductForEdit(prod));
            row.addEventListener('focus', () => setProdFocus(idx));

            row.addEventListener('keydown', (e) => {
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    const next = Math.min(idx + 1, filteredProds.length - 1);
                    focusProdRow(next);
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    if (idx === 0) { if (prodSearchEl) prodSearchEl.focus(); return; }
                    focusProdRow(idx - 1);
                } else if (e.key === 'Enter') {
                    e.preventDefault();
                    openProductForEdit(filteredProds[idx]);
                } else if (e.key === 'Delete' && e.ctrlKey) {
                    e.preventDefault();
                    deleteProductConfirm(filteredProds[idx]);
                }
            });

            if (prodEmptyState) prodListBody.insertBefore(row, prodEmptyState);
            else prodListBody.appendChild(row);
        });
    }

    function setProdFocus(idx) {
        prodFocusedIdx = idx;
        prodListBody.querySelectorAll('.prod-row').forEach((r, i) => {
            r.classList.toggle('focused', i === idx);
        });
    }

    function focusProdRow(idx) {
        const rows = prodListBody.querySelectorAll('.prod-row');
        if (rows[idx]) {
            setProdFocus(idx);
            rows[idx].focus();
            rows[idx].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        }
    }

    /* ── Open form in 'new' mode ── */
    function openNewProduct() {
        prodEditingId = null;
        clearProdForm();
        showProdForm(true);
        if (prodDeleteBtn) prodDeleteBtn.classList.add('hidden');
        prodFlowFields[0]?.focus();
    }

    /* ── Open form in 'edit' mode ── */
    function openProductForEdit(prod) {
        prodEditingId = prod.id;
        populateProdForm(prod);
        showProdForm(true);
        if (prodDeleteBtn) prodDeleteBtn.classList.remove('hidden');

        prodListBody.querySelectorAll('.prod-row').forEach(r => {
            r.classList.toggle('active', r.dataset.id === prod.id);
        });

        prodFlowFields[0]?.focus();
    }

    /* ── Show / hide form ── */
    function showProdForm(show) {
        if (prodFormEl) prodFormEl.classList.toggle('hidden', !show);
        if (prodPlaceholder) prodPlaceholder.style.display = show ? 'none' : '';
    }

    /* ── Populate form from product object ── */
    function populateProdForm(prod) {
        // Resolve size value — normalize legacy formats (e.g. "9096" → "90 ML (96)")
        let sizeVal = normalizeLegacySize(prod.size || '');

        const map = {
            prodBrandName:   prod.brandName || '',
            prodCode:        prod.code || '',
            prodCategory:    prod.category || '',
            prodSubCategory: prod.subCategory || '',
            prodSize:        sizeVal,
            prodBpc:         prod.bpc || '',
            prodMrp:         prod.mrp || '',
            prodCostPrice:   prod.costPrice || '',
            prodHsn:         prod.hsn || '',
            prodSupplier:    prod.supplier || '',
            prodBarcode:     prod.barcode || '',
            prodRemarks:     prod.remarks || '',
        };
        Object.entries(map).forEach(([id, val]) => {
            const el = document.getElementById(id);
            if (el) el.value = val;
        });
    }

    /* ── Clear form ── */
    function clearProdForm() {
        prodFlowFields.forEach(f => {
            if (f.tagName === 'SELECT') f.selectedIndex = 0;
            else f.value = '';
        });
    }

    /* ── Gather form data ── */
    function gatherProdFormData() {
        return {
            id:          prodEditingId || ('PROD_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6)),
            brandName:   document.getElementById('prodBrandName')?.value.trim() || '',
            code:        document.getElementById('prodCode')?.value.trim().toUpperCase() || '',
            category:    document.getElementById('prodCategory')?.value || '',
            subCategory: document.getElementById('prodSubCategory')?.value || '',
            size:        document.getElementById('prodSize')?.value || '',
            bpc:         parseInt(document.getElementById('prodBpc')?.value) || 0,
            mrp:         parseFloat(document.getElementById('prodMrp')?.value) || 0,
            costPrice:   parseFloat(document.getElementById('prodCostPrice')?.value) || 0,
            hsn:         document.getElementById('prodHsn')?.value.trim() || '',
            supplier:    document.getElementById('prodSupplier')?.value.trim() || '',
            barcode:     document.getElementById('prodBarcode')?.value.trim() || '',
            remarks:     document.getElementById('prodRemarks')?.value.trim() || '',
        };
    }

    /* ── Save product ── */
    async function saveProduct() {
        const data = gatherProdFormData();

        if (!data.brandName) {
            document.getElementById('prodBrandName')?.focus();
            showTpToast('Brand name is required', true);
            return;
        }
        if (!data.category) {
            document.getElementById('prodCategory')?.focus();
            showTpToast('Category is required', true);
            return;
        }
        if (!data.size) {
            document.getElementById('prodSize')?.focus();
            showTpToast('Size is required', true);
            return;
        }

        try {
            const result = await window.electronAPI.saveProduct({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || '',
                product: data,
            });
            if (result.success) {
                showTpToast(`Product "${data.brandName}" saved`);
                await loadProducts();
                cancelProdForm();
            } else {
                showTpToast('Save failed: ' + result.error, true);
            }
        } catch (err) {
            showTpToast('Save error: ' + err.message, true);
        }
    }

    /* ── Delete product ── */
    async function deleteProductConfirm(prod) {
        const target = prod || (prodEditingId && allProducts.find(p => p.id === prodEditingId));
        if (!target) return;
        const proceed = confirm(`Delete product "${target.brandName}"?`);
        if (!proceed) return;

        try {
            const result = await window.electronAPI.deleteProduct({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || '',
                productId: target.id,
            });
            if (result.success) {
                showTpToast(`Product "${target.brandName}" deleted`);
                await loadProducts();
                cancelProdForm();
            } else {
                showTpToast('Delete failed: ' + result.error, true);
            }
        } catch (err) {
            showTpToast('Delete error: ' + err.message, true);
        }
    }

    /* ── Cancel form ── */
    function cancelProdForm() {
        prodEditingId = null;
        clearProdForm();
        showProdForm(false);
        prodListBody.querySelectorAll('.prod-row').forEach(r => r.classList.remove('active'));
        const rows = prodListBody.querySelectorAll('.prod-row');
        if (rows.length > 0) focusProdRow(Math.min(prodFocusedIdx, rows.length - 1));
        else if (prodSearchEl) prodSearchEl.focus();
    }

    /* ── Wire buttons ── */
    if (prodNewBtn)    prodNewBtn.addEventListener('click', openNewProduct);
    if (prodSaveBtn)   prodSaveBtn.addEventListener('click', saveProduct);
    if (prodCancelBtn) prodCancelBtn.addEventListener('click', cancelProdForm);
    if (prodDeleteBtn) prodDeleteBtn.addEventListener('click', () => deleteProductConfirm());

    /* ── Sequential field navigation (Enter key through form) ── */
    prodFlowFields.forEach((field, idx) => {
        field.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                if (idx < prodFlowFields.length - 1) {
                    prodFlowFields[idx + 1].focus();
                } else {
                    saveProduct();
                }
            } else if (e.key === 'Escape') {
                e.preventDefault();
                cancelProdForm();
            }
        });
    });

    /* ── Search ── */
    if (prodSearchEl) {
        prodSearchEl.addEventListener('input', () => {
            const q = prodSearchEl.value.toLowerCase();
            filteredProds = allProducts.filter(p =>
                (p.brandName || '').toLowerCase().includes(q) ||
                (p.code || '').toLowerCase().includes(q) ||
                (p.category || '').toLowerCase().includes(q) ||
                (p.subCategory || '').toLowerCase().includes(q) ||
                (p.size || '').includes(q) ||
                (p.supplier || '').toLowerCase().includes(q)
            );
            renderProductList();
        });

        prodSearchEl.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                focusProdRow(0);
            } else if (e.key === 'Enter' && filteredProds.length > 0) {
                e.preventDefault();
                const target = filteredProds[prodFocusedIdx >= 0 ? prodFocusedIdx : 0];
                if (target) openProductForEdit(target);
            }
        });
    }

    /* ── Global shortcuts for product panel ── */
    document.addEventListener('keydown', (e) => {
        const prodPanel = document.getElementById('sub-product-master');
        if (!prodPanel || !prodPanel.classList.contains('active')) return;

        if (e.key === 'F2') {
            e.preventDefault();
            openNewProduct();
            return;
        }

        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
            if (prodFormEl && !prodFormEl.classList.contains('hidden')) {
                e.preventDefault();
                saveProduct();
                return;
            }
        }

        if (e.ctrlKey && e.key === 'Delete') {
            if (prodEditingId) {
                e.preventDefault();
                deleteProductConfirm();
                return;
            }
        }

        if (e.key === 'Escape') {
            if (prodFormEl && !prodFormEl.classList.contains('hidden')) {
                e.preventDefault();
                cancelProdForm();
            } else if (prodSearchEl && prodSearchEl.value) {
                e.preventDefault();
                prodSearchEl.value = '';
                prodSearchEl.dispatchEvent(new Event('input'));
                prodSearchEl.focus();
            }
        }
    });

    /* ── Bulk Import: Download Template ── */
    const prodTemplateBtn = document.getElementById('prodTemplateBtn');
    const prodImportBtn   = document.getElementById('prodImportBtn');

    if (prodTemplateBtn) {
        prodTemplateBtn.addEventListener('click', async () => {
            try {
                prodTemplateBtn.disabled = true;
                prodTemplateBtn.textContent = 'Saving…';
                const result = await window.electronAPI.downloadProductTemplate();
                if (result.success) {
                    showTpToast('Template saved — fill it and use Import Excel');
                } else if (result.error !== 'Cancelled') {
                    showTpToast('Template error: ' + result.error, true);
                }
            } catch (err) {
                showTpToast('Error: ' + err.message, true);
            } finally {
                prodTemplateBtn.disabled = false;
                prodTemplateBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Template`;
            }
        });
    }

    /* ── Bulk Import: Import Excel ── */
    if (prodImportBtn) {
        prodImportBtn.addEventListener('click', async () => {
            if (!activeBar.barName) {
                showTpToast('Open a bar first', true);
                return;
            }
            try {
                prodImportBtn.disabled = true;
                prodImportBtn.textContent = 'Importing…';
                const result = await window.electronAPI.importProductsExcel({
                    barName: activeBar.barName,
                    financialYear: activeBar.financialYear || '',
                });
                if (result.success) {
                    let msg = `Imported ${result.imported} of ${result.total} products`;
                    if (result.skipped > 0) {
                        msg += ` (${result.skipped} skipped)`;
                        // Log skipped details
                        console.table(result.skippedDetails);
                    }
                    showTpToast(msg);
                    await loadProducts();
                } else if (result.error !== 'Cancelled') {
                    showTpToast('Import failed: ' + result.error, true);
                }
            } catch (err) {
                showTpToast('Import error: ' + err.message, true);
            } finally {
                prodImportBtn.disabled = false;
                prodImportBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg> Import Excel`;
            }
        });
    }

    /* ── Load products on init, then load opening stock, then MRP ── */
    loadProducts().then(() => { loadOpeningStock(); loadMrpMaster(); });

    // ═══════════════════════════════════════════════════════
    //  MRP MASTER — Price Management with History
    // ═══════════════════════════════════════════════════════

    const mrpListBody     = document.getElementById('mrpListBody');
    const mrpEmptyState   = document.getElementById('mrpEmptyState');
    const mrpCountEl      = document.getElementById('mrpCount');
    const mrpUpdatedInfo  = document.getElementById('mrpUpdatedInfo');
    const mrpSearchEl     = document.getElementById('mrpSearch');
    const mrpFilterCatEl  = document.getElementById('mrpFilterCat');
    const mrpFilterSizeEl = document.getElementById('mrpFilterSize');
    const mrpDetailEl     = document.getElementById('mrpDetail');
    const mrpPlaceholder  = document.getElementById('mrpPlaceholder');
    const mrpSyncBtn      = document.getElementById('mrpSyncBtn');
    const mrpUpdateBtn    = document.getElementById('mrpUpdateBtn');
    const mrpCancelBtn    = document.getElementById('mrpCancelBtn');

    let allMrpEntries  = [];
    let filteredMrp    = [];
    let mrpFocusedIdx  = -1;
    let mrpSelectedId  = null; // currently selected MRP entry ID

    /* ── Populate MRP size filter from ALL_SIZES ── */
    if (mrpFilterSizeEl) {
        mrpFilterSizeEl.innerHTML = '<option value="">All Sizes</option>' +
            ALL_SIZES.filter(s => s.value).map(s =>
                `<option value="${s.value}">${s.label}</option>`
            ).join('');
    }

    /* ── Load MRP entries from disk ── */
    async function loadMrpMaster() {
        if (!activeBar.barName) return;
        try {
            const result = await window.electronAPI.getMrpMaster({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || ''
            });
            if (result.success) {
                allMrpEntries = result.entries || [];
            }
        } catch (err) {
            console.error('loadMrpMaster error:', err);
        }
        applyMrpFilters();
    }

    /* ── Sync products into MRP master (auto-create entries for new products) ── */
    async function syncProductsToMrp() {
        if (!activeBar.barName) return;
        const today = new Date().toISOString().slice(0, 10);
        let added = 0;

        for (const prod of allProducts) {
            // Check if MRP entry already exists for this product
            const exists = allMrpEntries.find(m => m.productId === prod.id);
            if (exists) {
                // Update product info (brand, code, category, size may have changed)
                exists.brandName = prod.brandName;
                exists.code = prod.code;
                exists.category = prod.category;
                exists.size = prod.size;
                continue;
            }

            // Create new MRP entry
            const mrpEntry = {
                id: 'MRP_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6) + '_' + added,
                productId: prod.id,
                brandName: prod.brandName,
                code: prod.code || '',
                category: prod.category || '',
                size: prod.size || '',
                currentMrp: prod.mrp || 0,
                currentCost: prod.costPrice || 0,
                mrpHistory: [],
            };

            // If product already has an MRP, add it as the initial history entry
            if (prod.mrp > 0) {
                mrpEntry.mrpHistory.push({
                    mrp: prod.mrp,
                    costPrice: prod.costPrice || 0,
                    effectiveFrom: today,
                    effectiveTo: null,
                    changedAt: new Date().toISOString(),
                    notes: 'Initial MRP from Product Master',
                    type: 'initial',
                });
            }

            allMrpEntries.push(mrpEntry);
            added++;
        }

        // Remove MRP entries for products that no longer exist
        const productIds = new Set(allProducts.map(p => p.id));
        allMrpEntries = allMrpEntries.filter(m => productIds.has(m.productId));

        // Save all to disk
        for (const entry of allMrpEntries) {
            try {
                await window.electronAPI.saveMrpEntry({
                    barName: activeBar.barName,
                    financialYear: activeBar.financialYear || '',
                    mrpEntry: entry,
                });
            } catch (_) { /* silent save errors */ }
        }

        applyMrpFilters();
        showTpToast(`MRP Master synced — ${added} new product${added !== 1 ? 's' : ''} added, ${allMrpEntries.length} total`);
    }

    /* ── Apply filters and render ── */
    function applyMrpFilters() {
        const catFilter  = mrpFilterCatEl?.value || '';
        const sizeFilter = mrpFilterSizeEl?.value || '';
        const searchQ    = (mrpSearchEl?.value || '').toLowerCase();

        filteredMrp = allMrpEntries.filter(m => {
            if (catFilter && m.category !== catFilter) return false;
            if (sizeFilter && m.size !== sizeFilter) return false;
            if (searchQ) {
                const hay = ((m.brandName || '') + ' ' + (m.code || '') + ' ' + (m.category || '')).toLowerCase();
                if (!hay.includes(searchQ)) return false;
            }
            return true;
        });

        // Sort by brand name
        filteredMrp.sort((a, b) => (a.brandName || '').localeCompare(b.brandName || ''));
        renderMrpList();
    }

    /* ── Render MRP list ── */
    function renderMrpList() {
        mrpListBody.querySelectorAll('.mrp-row').forEach(r => r.remove());

        if (filteredMrp.length === 0) {
            if (mrpEmptyState) mrpEmptyState.style.display = '';
            if (mrpCountEl) mrpCountEl.textContent = `${allMrpEntries.length} product${allMrpEntries.length !== 1 ? 's' : ''}`;
            return;
        }
        if (mrpEmptyState) mrpEmptyState.style.display = 'none';
        if (mrpCountEl) mrpCountEl.textContent = `${filteredMrp.length} of ${allMrpEntries.length} product${allMrpEntries.length !== 1 ? 's' : ''}`;

        filteredMrp.forEach((entry, idx) => {
            const row = document.createElement('div');
            row.className = 'mrp-row';
            row.tabIndex = 0;
            row.dataset.idx = idx;
            row.dataset.id = entry.id;

            const mrp = entry.currentMrp || 0;
            const cost = entry.currentCost || 0;
            const margin = mrp > 0 && cost > 0 ? (((mrp - cost) / mrp) * 100) : 0;
            const marginClass = margin > 0 ? 'positive' : margin < 0 ? 'negative' : 'zero';

            const hasChanges = (entry.mrpHistory || []).length > 1;
            const badge = hasChanges ? '<span class="mrp-row-badge changed">REVISED</span>' : '';

            row.innerHTML = `
                <span class="mrp-row-name">${esc(entry.brandName)}${badge}</span>
                <span class="mrp-row-size">${esc(entry.size || '—')}</span>
                <span class="mrp-row-mrp">${mrp > 0 ? '₹' + mrp.toFixed(2) : '—'}</span>
                <span class="mrp-row-cost">${cost > 0 ? '₹' + cost.toFixed(2) : '—'}</span>
                <span class="mrp-row-margin ${marginClass}">${margin !== 0 ? margin.toFixed(1) + '%' : '—'}</span>
            `;

            row.addEventListener('click', () => {
                setMrpFocus(idx);
                selectMrpEntry(entry);
            });
            row.addEventListener('focus', () => setMrpFocus(idx));

            row.addEventListener('keydown', (e) => {
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    focusMrpRow(Math.min(idx + 1, filteredMrp.length - 1));
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    if (idx === 0) { if (mrpSearchEl) mrpSearchEl.focus(); return; }
                    focusMrpRow(idx - 1);
                } else if (e.key === 'Enter') {
                    e.preventDefault();
                    selectMrpEntry(filteredMrp[idx]);
                }
            });

            if (mrpEmptyState) mrpListBody.insertBefore(row, mrpEmptyState);
            else mrpListBody.appendChild(row);
        });
    }

    function setMrpFocus(idx) {
        mrpFocusedIdx = idx;
        mrpListBody.querySelectorAll('.mrp-row').forEach((r, i) => {
            r.classList.toggle('focused', i === idx);
        });
    }

    function focusMrpRow(idx) {
        const rows = mrpListBody.querySelectorAll('.mrp-row');
        if (rows[idx]) {
            setMrpFocus(idx);
            rows[idx].focus();
            rows[idx].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        }
    }

    /* ── Select MRP entry to show detail ── */
    function selectMrpEntry(entry) {
        mrpSelectedId = entry.id;

        // Highlight active row
        mrpListBody.querySelectorAll('.mrp-row').forEach(r => {
            r.classList.toggle('active', r.dataset.id === entry.id);
        });

        // Show detail panel
        if (mrpPlaceholder) mrpPlaceholder.style.display = 'none';
        if (mrpDetailEl) mrpDetailEl.classList.remove('hidden');

        // Populate product card
        const cardBrand = document.getElementById('mrpCardBrand');
        const cardCode  = document.getElementById('mrpCardCode');
        const cardSize  = document.getElementById('mrpCardSize');
        const cardCat   = document.getElementById('mrpCardCat');
        if (cardBrand) cardBrand.textContent = entry.brandName || '—';
        if (cardCode)  cardCode.textContent  = entry.code || '—';
        if (cardSize)  cardSize.textContent  = entry.size || '—';
        if (cardCat)   cardCat.textContent   = entry.category || '—';

        // Populate current MRP
        const curVal   = document.getElementById('mrpCurrentValue');
        const curSince = document.getElementById('mrpCurrentSince');
        if (curVal) curVal.textContent = entry.currentMrp > 0 ? '₹ ' + entry.currentMrp.toFixed(2) : '₹ 0.00';

        // Find effective date of current MRP
        const history = entry.mrpHistory || [];
        const currentEntry = history.find(h => h.effectiveTo === null);
        if (curSince) {
            if (currentEntry) {
                curSince.textContent = 'Effective from ' + formatDateShort(currentEntry.effectiveFrom);
            } else if (entry.currentMrp > 0) {
                curSince.textContent = 'Set in Product Master';
            } else {
                curSince.textContent = 'No MRP set';
            }
        }

        // Pre-fill form
        const newMrpEl  = document.getElementById('mrpNewValue');
        const effDateEl = document.getElementById('mrpEffectiveDate');
        const newCostEl = document.getElementById('mrpNewCost');
        const notesEl   = document.getElementById('mrpNotes');
        if (newMrpEl)  newMrpEl.value  = '';
        if (effDateEl) effDateEl.value = new Date().toISOString().slice(0, 10);
        if (newCostEl) newCostEl.value = '';
        if (notesEl)   notesEl.value   = '';

        // Render history
        renderMrpHistory(entry);
    }

    /* ── Format date helper ── */
    function formatDateShort(dateStr) {
        if (!dateStr) return '—';
        const d = new Date(dateStr);
        const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        return d.getDate() + ' ' + months[d.getMonth()] + ' ' + d.getFullYear();
    }

    /* ── Render MRP history timeline ── */
    function renderMrpHistory(entry) {
        const container = document.getElementById('mrpHistory');
        if (!container) return;

        const history = [...(entry.mrpHistory || [])];

        if (history.length === 0) {
            container.innerHTML = '<div class="mrp-history-empty">No price changes recorded yet</div>';
            return;
        }

        // Sort by date descending (newest first)
        history.sort((a, b) => new Date(b.effectiveFrom) - new Date(a.effectiveFrom));

        let html = '<div class="mrp-timeline">';

        history.forEach((h, idx) => {
            const isFirst = idx === 0;
            const isInitial = h.type === 'initial';
            const tag = isFirst ? '<span class="mrp-tl-tag current">CURRENT</span>' :
                                  '<span class="mrp-tl-tag old">PREVIOUS</span>';

            let priceHtml = '';
            if (h.oldMrp !== undefined && h.oldMrp !== null) {
                // Rate change entry
                const diff = h.mrp - h.oldMrp;
                const pct = h.oldMrp > 0 ? ((diff / h.oldMrp) * 100).toFixed(1) : '0';
                const changeClass = diff > 0 ? 'up' : 'down';
                const changeSign = diff > 0 ? '+' : '';
                priceHtml = `
                    <span class="mrp-tl-price-val">₹${h.oldMrp.toFixed(2)}</span>
                    <span class="mrp-tl-arrow">→</span>
                    <span class="mrp-tl-price-val">₹${h.mrp.toFixed(2)}</span>
                    <span class="mrp-tl-change ${changeClass}">${changeSign}${diff.toFixed(2)} (${changeSign}${pct}%)</span>
                `;
            } else {
                // Initial entry
                priceHtml = `<span class="mrp-tl-price-val">₹${h.mrp.toFixed(2)}</span>`;
            }

            const notesHtml = h.notes ? `<div class="mrp-tl-notes">${esc(h.notes)}</div>` : '';
            const costHtml = h.costPrice ? `<div class="mrp-tl-cost">Cost: ₹${h.costPrice.toFixed(2)}</div>` : '';

            html += `
                <div class="mrp-tl-item ${isInitial ? 'initial' : ''}">
                    <div class="mrp-tl-header">
                        <span class="mrp-tl-date">${formatDateShort(h.effectiveFrom)}</span>
                        ${tag}
                    </div>
                    <div class="mrp-tl-prices">${priceHtml}</div>
                    ${costHtml}
                    ${notesHtml}
                </div>
            `;
        });

        html += '</div>';
        container.innerHTML = html;
    }

    /* ── Update MRP for selected product ── */
    async function updateMrp() {
        if (!mrpSelectedId) return;
        const entry = allMrpEntries.find(m => m.id === mrpSelectedId);
        if (!entry) return;

        const newMrpEl  = document.getElementById('mrpNewValue');
        const effDateEl = document.getElementById('mrpEffectiveDate');
        const newCostEl = document.getElementById('mrpNewCost');
        const notesEl   = document.getElementById('mrpNotes');

        const newMrp     = parseFloat(newMrpEl?.value) || 0;
        const effDate    = effDateEl?.value || '';
        const newCost    = parseFloat(newCostEl?.value) || 0;
        const notes      = notesEl?.value.trim() || '';

        if (newMrp <= 0) {
            newMrpEl?.focus();
            showTpToast('New MRP must be greater than 0', true);
            return;
        }
        if (!effDate) {
            effDateEl?.focus();
            showTpToast('Effective date is required', true);
            return;
        }

        // Check if MRP actually changed
        if (newMrp === entry.currentMrp && newCost === entry.currentCost) {
            showTpToast('MRP is the same — no change needed', true);
            return;
        }

        // Close the previous history entry's effectiveTo
        if (entry.mrpHistory && entry.mrpHistory.length > 0) {
            const lastOpen = entry.mrpHistory.find(h => h.effectiveTo === null);
            if (lastOpen) {
                // Set effectiveTo to the day before the new effective date
                const prevDate = new Date(effDate);
                prevDate.setDate(prevDate.getDate() - 1);
                lastOpen.effectiveTo = prevDate.toISOString().slice(0, 10);
            }
        }

        // Create new history entry
        const historyEntry = {
            mrp: newMrp,
            oldMrp: entry.currentMrp || 0,
            costPrice: newCost || entry.currentCost || 0,
            oldCost: entry.currentCost || 0,
            effectiveFrom: effDate,
            effectiveTo: null,
            changedAt: new Date().toISOString(),
            notes: notes || (newMrp > entry.currentMrp ? 'Price increased' : 'Price decreased'),
            type: entry.mrpHistory.length === 0 ? 'initial' : 'revision',
        };

        if (!entry.mrpHistory) entry.mrpHistory = [];
        entry.mrpHistory.push(historyEntry);

        // Update current values
        const oldMrp = entry.currentMrp;
        entry.currentMrp = newMrp;
        if (newCost > 0) entry.currentCost = newCost;

        // Also update the product in Product Master to keep in sync
        const product = allProducts.find(p => p.id === entry.productId);
        if (product) {
            product.mrp = newMrp;
            if (newCost > 0) product.costPrice = newCost;
            try {
                await window.electronAPI.saveProduct({
                    barName: activeBar.barName,
                    financialYear: activeBar.financialYear || '',
                    product: product,
                });
            } catch (_) { /* silent */ }
        }

        // Save MRP entry
        try {
            const result = await window.electronAPI.saveMrpEntry({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || '',
                mrpEntry: entry,
            });
            if (result.success) {
                const diff = newMrp - oldMrp;
                const sign = diff > 0 ? '+' : '';
                showTpToast(`MRP updated: ₹${oldMrp.toFixed(2)} → ₹${newMrp.toFixed(2)} (${sign}${diff.toFixed(2)})`);
                applyMrpFilters();
                // Re-select to refresh detail
                selectMrpEntry(entry);
            } else {
                showTpToast('Save failed: ' + result.error, true);
            }
        } catch (err) {
            showTpToast('Save error: ' + err.message, true);
        }
    }

    /* ── Cancel / deselect ── */
    function cancelMrpSelection() {
        mrpSelectedId = null;
        if (mrpDetailEl) mrpDetailEl.classList.add('hidden');
        if (mrpPlaceholder) mrpPlaceholder.style.display = '';
        mrpListBody.querySelectorAll('.mrp-row').forEach(r => r.classList.remove('active'));
    }

    /* ── Wire buttons ── */
    if (mrpSyncBtn)   mrpSyncBtn.addEventListener('click', syncProductsToMrp);
    if (mrpUpdateBtn) mrpUpdateBtn.addEventListener('click', updateMrp);
    if (mrpCancelBtn) mrpCancelBtn.addEventListener('click', cancelMrpSelection);

    /* ── Wire filters ── */
    if (mrpFilterCatEl)  mrpFilterCatEl.addEventListener('change', applyMrpFilters);
    if (mrpFilterSizeEl) mrpFilterSizeEl.addEventListener('change', applyMrpFilters);
    if (mrpSearchEl) {
        mrpSearchEl.addEventListener('input', applyMrpFilters);
        mrpSearchEl.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                focusMrpRow(0);
            } else if (e.key === 'Enter' && filteredMrp.length > 0) {
                e.preventDefault();
                const target = filteredMrp[mrpFocusedIdx >= 0 ? mrpFocusedIdx : 0];
                if (target) selectMrpEntry(target);
            }
        });
    }

    /* ── MRP form field Enter navigation ── */
    const mrpFlowFields = Array.from(document.querySelectorAll('.mrp-flow'));
    mrpFlowFields.forEach((field, idx) => {
        field.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                if (idx < mrpFlowFields.length - 1) {
                    mrpFlowFields[idx + 1].focus();
                } else {
                    updateMrp();
                }
            } else if (e.key === 'Escape') {
                e.preventDefault();
                cancelMrpSelection();
            }
        });
    });

    /* ── Global shortcuts for MRP panel ── */
    document.addEventListener('keydown', (e) => {
        const mrpPanel = document.getElementById('sub-mrp-master');
        if (!mrpPanel || !mrpPanel.classList.contains('active')) return;

        if (e.key === 'Escape') {
            if (mrpSelectedId) {
                e.preventDefault();
                cancelMrpSelection();
            } else if (mrpSearchEl && mrpSearchEl.value) {
                e.preventDefault();
                mrpSearchEl.value = '';
                applyMrpFilters();
                mrpSearchEl.focus();
            }
        }
    });

    // ═══════════════════════════════════════════════════════
    //  SHORTCUT MASTER — Quick Product Lookup Codes
    // ═══════════════════════════════════════════════════════

    const scListBody     = document.getElementById('scListBody');
    const scEmptyState   = document.getElementById('scEmptyState');
    const scCountEl      = document.getElementById('scCount');
    const scSearchEl     = document.getElementById('scSearch');
    const scFormEl       = document.getElementById('scForm');
    const scPlaceholder  = document.getElementById('scPlaceholder');
    const scNewBtn       = document.getElementById('scNewBtn');
    const scSaveBtn      = document.getElementById('scSaveBtn');
    const scCancelBtn    = document.getElementById('scCancelBtn');
    const scDeleteBtn    = document.getElementById('scDeleteBtn');
    const scAutoGenBtn   = document.getElementById('scAutoGenBtn');
    const scTestInput    = document.getElementById('scTestInput');
    const scTestResult   = document.getElementById('scTestResult');
    const scFlowFields   = Array.from(document.querySelectorAll('.sc-flow'));

    let allShortcuts   = [];
    let filteredSc     = [];
    let scFocusedIdx   = -1;
    let scEditingId    = null;
    let scLinkedProdId = null;

    /* ── Load shortcuts from disk ── */
    async function loadShortcuts() {
        if (!activeBar.barName) return;
        try {
            const result = await window.electronAPI.getShortcuts({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || ''
            });
            if (result.success) {
                allShortcuts = result.shortcuts || [];
            }
        } catch (err) {
            console.error('loadShortcuts error:', err);
        }
        filteredSc = [...allShortcuts];
        renderScList();
    }

    /* ── Render shortcut list ── */
    function renderScList() {
        scListBody.querySelectorAll('.sc-row').forEach(r => r.remove());

        if (filteredSc.length === 0) {
            if (scEmptyState) scEmptyState.style.display = '';
            if (scCountEl) scCountEl.textContent = `${allShortcuts.length} shortcut${allShortcuts.length !== 1 ? 's' : ''}`;
            return;
        }
        if (scEmptyState) scEmptyState.style.display = 'none';
        if (scCountEl) scCountEl.textContent = `${filteredSc.length} shortcut${filteredSc.length !== 1 ? 's' : ''}`;

        filteredSc.forEach((sc, idx) => {
            const row = document.createElement('div');
            row.className = 'sc-row';
            row.tabIndex = 0;
            row.dataset.idx = idx;
            row.dataset.id = sc.id;
            row.innerHTML = `
                <span class="sc-row-code">${esc(sc.code)}</span>
                <span class="sc-row-brand">${esc(sc.brandName)}</span>
                <span class="sc-row-size">${esc(sc.size || '—')}</span>
                <span class="sc-row-cat">${esc(sc.category || '—')}</span>
            `;

            row.addEventListener('click', () => {
                setScFocus(idx);
                openScForEdit(sc);
            });
            row.addEventListener('dblclick', () => openScForEdit(sc));
            row.addEventListener('focus', () => setScFocus(idx));

            row.addEventListener('keydown', (e) => {
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    focusScRow(Math.min(idx + 1, filteredSc.length - 1));
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    if (idx === 0) { if (scSearchEl) scSearchEl.focus(); return; }
                    focusScRow(idx - 1);
                } else if (e.key === 'Enter') {
                    e.preventDefault();
                    openScForEdit(filteredSc[idx]);
                } else if (e.key === 'Delete' && e.ctrlKey) {
                    e.preventDefault();
                    deleteScConfirm(filteredSc[idx]);
                }
            });

            if (scEmptyState) scListBody.insertBefore(row, scEmptyState);
            else scListBody.appendChild(row);
        });
    }

    function setScFocus(idx) {
        scFocusedIdx = idx;
        scListBody.querySelectorAll('.sc-row').forEach((r, i) => {
            r.classList.toggle('focused', i === idx);
        });
    }

    function focusScRow(idx) {
        const rows = scListBody.querySelectorAll('.sc-row');
        if (rows[idx]) {
            setScFocus(idx);
            rows[idx].focus();
            rows[idx].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        }
    }

    /* ── Show / hide form ── */
    function showScForm(show) {
        if (scFormEl) scFormEl.classList.toggle('hidden', !show);
        if (scPlaceholder) scPlaceholder.style.display = show ? 'none' : '';
    }

    /* ── Open new shortcut ── */
    function openNewSc() {
        scEditingId = null;
        scLinkedProdId = null;
        clearScForm();
        showScForm(true);
        if (scDeleteBtn) scDeleteBtn.classList.add('hidden');
        const codeEl = document.getElementById('scCode');
        if (codeEl) codeEl.focus();
    }

    /* ── Open for edit ── */
    function openScForEdit(sc) {
        scEditingId = sc.id;
        scLinkedProdId = sc.productId || null;
        populateScForm(sc);
        showScForm(true);
        if (scDeleteBtn) scDeleteBtn.classList.remove('hidden');

        scListBody.querySelectorAll('.sc-row').forEach(r => {
            r.classList.toggle('active', r.dataset.id === sc.id);
        });

        const codeEl = document.getElementById('scCode');
        if (codeEl) codeEl.focus();
    }

    /* ── Populate form ── */
    function populateScForm(sc) {
        const codeEl    = document.getElementById('scCode');
        const brandEl   = document.getElementById('scBrand');
        const pCodeEl   = document.getElementById('scProdCode');
        const pSizeEl   = document.getElementById('scProdSize');
        const pCatEl    = document.getElementById('scProdCat');
        const pMrpEl    = document.getElementById('scProdMrp');

        if (codeEl)   codeEl.value   = sc.code || '';
        if (brandEl)  brandEl.value  = sc.brandName || '';
        if (pCodeEl)  pCodeEl.value  = sc.prodCode || '';
        if (pSizeEl)  pSizeEl.value  = sc.size || '';
        if (pCatEl)   pCatEl.value   = sc.category || '';
        if (pMrpEl)   pMrpEl.value   = sc.mrp ? '₹ ' + sc.mrp.toFixed(2) : '';
    }

    /* ── Clear form ── */
    function clearScForm() {
        ['scCode','scBrand','scProdCode','scProdSize','scProdCat','scProdMrp'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = '';
        });
        scLinkedProdId = null;
    }

    /* ── Gather form data ── */
    function gatherScFormData() {
        const code = (document.getElementById('scCode')?.value || '').trim().toUpperCase();
        const brandName = (document.getElementById('scBrand')?.value || '').trim();
        const prod = allProducts.find(p => p.id === scLinkedProdId);
        return {
            id: scEditingId || ('SC_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6)),
            code,
            productId: scLinkedProdId || '',
            brandName: brandName,
            prodCode: prod ? prod.code : (document.getElementById('scProdCode')?.value || '').trim(),
            size: prod ? prod.size : (document.getElementById('scProdSize')?.value || '').trim(),
            category: prod ? prod.category : (document.getElementById('scProdCat')?.value || '').trim(),
            mrp: prod ? (prod.mrp || 0) : 0,
        };
    }

    /* ── Save shortcut ── */
    async function saveSc() {
        const data = gatherScFormData();

        if (!data.code) {
            document.getElementById('scCode')?.focus();
            showTpToast('Short code is required', true);
            return;
        }
        if (!data.brandName) {
            document.getElementById('scBrand')?.focus();
            showTpToast('Linked product is required', true);
            return;
        }

        // Check for duplicate code (ignore self when editing)
        const dupe = allShortcuts.find(s =>
            s.code.toUpperCase() === data.code && s.id !== data.id
        );
        if (dupe) {
            document.getElementById('scCode')?.focus();
            showTpToast(`Code "${data.code}" already exists for ${dupe.brandName}`, true);
            return;
        }

        try {
            const result = await window.electronAPI.saveShortcut({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || '',
                shortcut: data,
            });
            if (result.success) {
                showTpToast(`Shortcut "${data.code}" → ${data.brandName} saved`);
                await loadShortcuts();
                cancelScForm();
            } else {
                showTpToast('Save failed: ' + result.error, true);
            }
        } catch (err) {
            showTpToast('Save error: ' + err.message, true);
        }
    }

    /* ── Delete shortcut ── */
    async function deleteScConfirm(sc) {
        const target = sc || (scEditingId && allShortcuts.find(s => s.id === scEditingId));
        if (!target) return;
        const proceed = confirm(`Delete shortcut "${target.code}"?`);
        if (!proceed) return;

        try {
            const result = await window.electronAPI.deleteShortcut({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || '',
                shortcutId: target.id,
            });
            if (result.success) {
                showTpToast(`Shortcut "${target.code}" deleted`);
                await loadShortcuts();
                cancelScForm();
            } else {
                showTpToast('Delete failed: ' + result.error, true);
            }
        } catch (err) {
            showTpToast('Delete error: ' + err.message, true);
        }
    }

    /* ── Cancel form ── */
    function cancelScForm() {
        scEditingId = null;
        scLinkedProdId = null;
        clearScForm();
        showScForm(false);
        scListBody.querySelectorAll('.sc-row').forEach(r => r.classList.remove('active'));
        const rows = scListBody.querySelectorAll('.sc-row');
        if (rows.length > 0) focusScRow(Math.min(scFocusedIdx, rows.length - 1));
        else if (scSearchEl) scSearchEl.focus();
    }

    /* ── Auto-generate shortcodes for all products ── */
    async function autoGenerateShortcuts() {
        if (!allProducts.length) {
            showTpToast('No products in master — add products first', true);
            return;
        }

        const proceed = confirm(
            `Auto-generate shortcuts for ${allProducts.length} products?\n\n` +
            `This will create codes from brand initials + size digit.\n` +
            `Existing shortcuts will NOT be overwritten.`
        );
        if (!proceed) return;

        const existingCodes = new Set(allShortcuts.map(s => s.code.toUpperCase()));
        let added = 0;

        for (const prod of allProducts) {
            // Skip if product already has a shortcut
            const hasShortcut = allShortcuts.some(s => s.productId === prod.id);
            if (hasShortcut) continue;

            // Generate code from brand initials + size
            const code = generateShortCode(prod, existingCodes);
            if (!code) continue;

            existingCodes.add(code);
            const scEntry = {
                id: 'SC_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6) + '_' + added,
                code: code,
                productId: prod.id,
                brandName: prod.brandName || '',
                prodCode: prod.code || '',
                size: prod.size || '',
                category: prod.category || '',
                mrp: prod.mrp || 0,
            };
            allShortcuts.push(scEntry);
            added++;
        }

        // Save all at once
        try {
            await window.electronAPI.saveAllShortcuts({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || '',
                shortcuts: allShortcuts,
            });
            showTpToast(`Generated ${added} shortcut${added !== 1 ? 's' : ''} — ${allShortcuts.length} total`);
            filteredSc = [...allShortcuts];
            renderScList();
        } catch (err) {
            showTpToast('Save error: ' + err.message, true);
        }
    }

    /* ── Generate a unique short code from brand name + size ── */
    function generateShortCode(prod, existingCodes) {
        const brand = (prod.brandName || '').trim();
        if (!brand) return null;

        // Get size digit (first significant digit of ML)
        const sizeMatch = (prod.size || '').match(/(\d+)/);
        const sizePart = sizeMatch ? sizeMatch[1] : '';
        // Use first 1-2 chars of size (e.g. "750" → "7", "1000" → "10")
        const sizeCode = sizePart.length >= 4 ? sizePart.slice(0, 2) : sizePart.slice(0, 1);

        // Strategy 1: First letters of each word + size digit
        const words = brand.split(/\s+/).filter(w => w.length > 0);
        let initials = words.map(w => w[0].toUpperCase()).join('');
        if (initials.length > 4) initials = initials.slice(0, 4);

        let candidate = initials + sizeCode;
        if (candidate.length >= 2 && !existingCodes.has(candidate.toUpperCase())) {
            return candidate.toUpperCase();
        }

        // Strategy 2: First 2-3 chars of brand + size digit
        const prefix = brand.replace(/\s+/g, '').slice(0, 3).toUpperCase();
        candidate = prefix + sizeCode;
        if (!existingCodes.has(candidate.toUpperCase())) {
            return candidate.toUpperCase();
        }

        // Strategy 3: Add numeric suffix
        for (let i = 1; i <= 99; i++) {
            candidate = initials + sizeCode + i;
            if (!existingCodes.has(candidate.toUpperCase())) {
                return candidate.toUpperCase();
            }
        }

        return null;
    }

    /* ── Product autocomplete in shortcut form ── */
    const scBrandInput = document.getElementById('scBrand');
    const scProdAcDrop = document.getElementById('scProdAc');
    let scAcIdx = -1;

    function scShowAc(query) {
        if (!scProdAcDrop) return;
        const q = (query || '').toLowerCase();
        const matches = q.length > 0
            ? allProducts.filter(p =>
                (p.brandName || '').toLowerCase().includes(q) ||
                (p.code || '').toLowerCase().includes(q)
              )
            : allProducts.slice(0, 30);

        if (matches.length === 0) {
            scProdAcDrop.innerHTML = '<div class="tp-prod-ac-empty" style="padding:8px 12px;font-size:11px;color:var(--text-muted)">No matching products</div>';
        } else {
            scProdAcDrop.innerHTML = matches.map((p, i) => `
                <div class="tp-prod-ac-item" data-idx="${i}">
                    <span class="tp-prod-ac-brand">${esc(p.brandName)}</span>
                    <span class="tp-prod-ac-meta">${esc(p.code || '')} · ${esc(p.size || '')} · ${esc(p.category || '')}</span>
                </div>
            `).join('');

            scProdAcDrop.querySelectorAll('.tp-prod-ac-item').forEach(item => {
                item.addEventListener('mousedown', (e) => {
                    e.preventDefault();
                    const idx = parseInt(item.dataset.idx);
                    scSelectProduct(matches[idx]);
                });
            });
        }

        scProdAcDrop.classList.remove('hidden');
        scAcIdx = -1;
    }

    function scHideAc() {
        if (scProdAcDrop) {
            scProdAcDrop.classList.add('hidden');
            scProdAcDrop.innerHTML = '';
        }
        scAcIdx = -1;
    }

    function scFocusAcItem(newIdx) {
        if (!scProdAcDrop) return;
        const items = scProdAcDrop.querySelectorAll('.tp-prod-ac-item');
        if (items.length === 0) return;
        scAcIdx = Math.max(0, Math.min(newIdx, items.length - 1));
        items.forEach((it, i) => it.classList.toggle('focused', i === scAcIdx));
        items[scAcIdx]?.scrollIntoView({ block: 'nearest' });
    }

    function scSelectProduct(prod) {
        scLinkedProdId = prod.id;
        if (scBrandInput) scBrandInput.value = prod.brandName || '';
        const pCodeEl = document.getElementById('scProdCode');
        const pSizeEl = document.getElementById('scProdSize');
        const pCatEl  = document.getElementById('scProdCat');
        const pMrpEl  = document.getElementById('scProdMrp');
        if (pCodeEl) pCodeEl.value = prod.code || '';
        if (pSizeEl) pSizeEl.value = prod.size || '';
        if (pCatEl)  pCatEl.value  = prod.category || '';
        if (pMrpEl)  pMrpEl.value  = prod.mrp ? '₹ ' + prod.mrp.toFixed(2) : '';
        scHideAc();

        // Auto-suggest code if empty
        const codeEl = document.getElementById('scCode');
        if (codeEl && !codeEl.value.trim()) {
            const existingCodes = new Set(allShortcuts.map(s => s.code.toUpperCase()));
            const suggested = generateShortCode(prod, existingCodes);
            if (suggested) codeEl.value = suggested;
        }
    }

    if (scBrandInput) {
        // Make container position:relative for AC dropdown
        const container = scBrandInput.parentElement;
        if (container) container.style.position = 'relative';

        scBrandInput.addEventListener('input', () => scShowAc(scBrandInput.value));

        scBrandInput.addEventListener('focus', () => {
            if (allProducts.length > 0) scShowAc(scBrandInput.value);
        });

        scBrandInput.addEventListener('blur', () => {
            setTimeout(() => scHideAc(), 150);
        });

        scBrandInput.addEventListener('keydown', (e) => {
            if (!scProdAcDrop || scProdAcDrop.classList.contains('hidden')) return;
            const items = scProdAcDrop.querySelectorAll('.tp-prod-ac-item');

            if (e.key === 'ArrowDown') {
                e.preventDefault(); e.stopPropagation();
                scFocusAcItem(scAcIdx + 1);
            } else if (e.key === 'ArrowUp') {
                e.preventDefault(); e.stopPropagation();
                scFocusAcItem(scAcIdx - 1);
            } else if (e.key === 'Enter' && scAcIdx >= 0 && items[scAcIdx]) {
                e.preventDefault(); e.stopPropagation();
                const q = scBrandInput.value.toLowerCase();
                const matches = q.length > 0
                    ? allProducts.filter(p =>
                        (p.brandName || '').toLowerCase().includes(q) ||
                        (p.code || '').toLowerCase().includes(q)
                      )
                    : allProducts.slice(0, 30);
                if (matches[scAcIdx]) scSelectProduct(matches[scAcIdx]);
            } else if (e.key === 'Escape') {
                e.preventDefault(); e.stopPropagation();
                scHideAc();
            }
        });
    }

    /* ── Test shortcode lookup ── */
    if (scTestInput && scTestResult) {
        scTestInput.addEventListener('input', () => {
            const code = scTestInput.value.trim().toUpperCase();
            if (!code) {
                scTestResult.className = 'sc-test-result';
                scTestResult.innerHTML = '<span class="sc-test-empty">Type a code above to test</span>';
                return;
            }

            const match = allShortcuts.find(s => s.code.toUpperCase() === code);
            if (match) {
                scTestResult.className = 'sc-test-result found';
                scTestResult.innerHTML = `
                    <span class="sc-test-brand">${esc(match.brandName)}</span>
                    <span class="sc-test-meta">${esc(match.prodCode || '')} · ${esc(match.size || '')}</span>
                    <span class="sc-test-mrp">${match.mrp ? '₹' + match.mrp.toFixed(2) : ''}</span>
                `;
            } else {
                // Try partial match
                const partial = allShortcuts.filter(s =>
                    s.code.toUpperCase().startsWith(code)
                );
                if (partial.length > 0) {
                    scTestResult.className = 'sc-test-result found';
                    const first = partial[0];
                    scTestResult.innerHTML = `
                        <span class="sc-test-brand">${esc(first.brandName)}</span>
                        <span class="sc-test-meta">${esc(first.prodCode || '')} · ${esc(first.size || '')}</span>
                        <span class="sc-test-mrp">${partial.length > 1 ? '+' + (partial.length - 1) + ' more' : (first.mrp ? '₹' + first.mrp.toFixed(2) : '')}</span>
                    `;
                } else {
                    scTestResult.className = 'sc-test-result not-found';
                    scTestResult.innerHTML = '<span class="sc-test-nope">No match found</span>';
                }
            }
        });
    }

    /* ── Wire buttons ── */
    if (scNewBtn)     scNewBtn.addEventListener('click', openNewSc);
    if (scSaveBtn)    scSaveBtn.addEventListener('click', saveSc);
    if (scCancelBtn)  scCancelBtn.addEventListener('click', cancelScForm);
    if (scDeleteBtn)  scDeleteBtn.addEventListener('click', () => deleteScConfirm());
    if (scAutoGenBtn) scAutoGenBtn.addEventListener('click', autoGenerateShortcuts);

    /* ── Sequential field navigation ── */
    scFlowFields.forEach((field, idx) => {
        field.addEventListener('keydown', (e) => {
            // Don't interfere with AC dropdown in brand field
            if (field.id === 'scBrand' && scProdAcDrop && !scProdAcDrop.classList.contains('hidden')) {
                if (['ArrowDown','ArrowUp'].includes(e.key)) return;
                if (e.key === 'Enter' && scAcIdx >= 0) return;
            }
            if (e.key === 'Enter') {
                e.preventDefault();
                if (idx < scFlowFields.length - 1) {
                    scFlowFields[idx + 1].focus();
                } else {
                    saveSc();
                }
            } else if (e.key === 'Escape') {
                e.preventDefault();
                cancelScForm();
            }
        });
    });

    /* ── Search shortcuts ── */
    if (scSearchEl) {
        scSearchEl.addEventListener('input', () => {
            const q = scSearchEl.value.toLowerCase();
            filteredSc = allShortcuts.filter(s =>
                (s.code || '').toLowerCase().includes(q) ||
                (s.brandName || '').toLowerCase().includes(q) ||
                (s.prodCode || '').toLowerCase().includes(q) ||
                (s.category || '').toLowerCase().includes(q)
            );
            renderScList();
        });

        scSearchEl.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                focusScRow(0);
            } else if (e.key === 'Enter' && filteredSc.length > 0) {
                e.preventDefault();
                const target = filteredSc[scFocusedIdx >= 0 ? scFocusedIdx : 0];
                if (target) openScForEdit(target);
            }
        });
    }

    /* ── Global shortcuts for shortcut panel ── */
    document.addEventListener('keydown', (e) => {
        const scPanel = document.getElementById('sub-shortcut-master');
        if (!scPanel || !scPanel.classList.contains('active')) return;

        if (e.key === 'F2') {
            e.preventDefault();
            openNewSc();
            return;
        }

        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
            if (scFormEl && !scFormEl.classList.contains('hidden')) {
                e.preventDefault();
                saveSc();
                return;
            }
        }

        if (e.ctrlKey && e.key === 'Delete') {
            if (scEditingId) {
                e.preventDefault();
                deleteScConfirm();
                return;
            }
        }

        if (e.key === 'Escape') {
            if (scFormEl && !scFormEl.classList.contains('hidden')) {
                e.preventDefault();
                cancelScForm();
            } else if (scSearchEl && scSearchEl.value) {
                e.preventDefault();
                scSearchEl.value = '';
                filteredSc = [...allShortcuts];
                renderScList();
                scSearchEl.focus();
            }
        }
    });

    /* ── Load shortcuts on init ── */
    loadShortcuts();

    // ═══════════════════════════════════════════════════════
    //  SUPPLIER MASTER — List + Form (Tally-style)
    // ═══════════════════════════════════════════════════════

    const supListBody   = document.getElementById('supListBody');
    const supEmptyState = document.getElementById('supEmptyState');
    const supCountEl    = document.getElementById('supCount');
    const supSearchEl   = document.getElementById('supSearch');
    const supFormPanel  = document.getElementById('supFormPanel');
    const supFormEl     = document.getElementById('supForm');
    const supPlaceholder = document.getElementById('supPlaceholder');
    const supNewBtn     = document.getElementById('supNewBtn');
    const supSaveBtn    = document.getElementById('supSaveBtn');
    const supCancelBtn  = document.getElementById('supCancelBtn');
    const supDeleteBtn  = document.getElementById('supDeleteBtn');
    const supFlowFields = Array.from(document.querySelectorAll('.sup-flow'));

    let allSuppliers  = [];
    let filteredSups  = [];
    let supFocusedIdx = -1;
    let supEditingId  = null; // null = new mode, string = edit mode

    /* ── Load suppliers from disk ── */
    async function loadSuppliers() {
        if (!activeBar.barName) return;
        try {
            const result = await window.electronAPI.getSuppliers({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || ''
            });
            if (result.success) {
                allSuppliers = result.suppliers || [];
            }
        } catch (err) {
            console.error('loadSuppliers error:', err);
        }
        filteredSups = [...allSuppliers];
        renderSupplierList();
        refreshAcSuppliers(allSuppliers);
    }

    /* ── Render supplier list ── */
    function renderSupplierList() {
        // Clear existing rows (keep empty state element)
        supListBody.querySelectorAll('.sup-row').forEach(r => r.remove());

        if (filteredSups.length === 0) {
            if (supEmptyState) supEmptyState.style.display = '';
            if (supCountEl) supCountEl.textContent = `${allSuppliers.length} supplier${allSuppliers.length !== 1 ? 's' : ''}`;
            return;
        }
        if (supEmptyState) supEmptyState.style.display = 'none';
        if (supCountEl) supCountEl.textContent = `${filteredSups.length} supplier${filteredSups.length !== 1 ? 's' : ''}`;

        filteredSups.forEach((sup, idx) => {
            const row = document.createElement('div');
            row.className = 'sup-row';
            row.tabIndex = 0;
            row.dataset.idx = idx;
            row.dataset.id = sup.id;
            row.innerHTML = `
                <span class="sup-row-name">${esc(sup.name)}</span>
                <span class="sup-row-city">${esc(sup.city || '—')}</span>
                <span class="sup-row-phone">${esc(sup.phone || '—')}</span>
            `;

            row.addEventListener('click', () => {
                setSupFocus(idx);
                openSupplierForEdit(sup);
            });
            row.addEventListener('dblclick', () => openSupplierForEdit(sup));
            row.addEventListener('focus', () => setSupFocus(idx));

            row.addEventListener('keydown', (e) => {
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    const next = Math.min(idx + 1, filteredSups.length - 1);
                    focusSupRow(next);
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    if (idx === 0) { if (supSearchEl) supSearchEl.focus(); return; }
                    focusSupRow(idx - 1);
                } else if (e.key === 'Enter') {
                    e.preventDefault();
                    openSupplierForEdit(filteredSups[idx]);
                } else if (e.key === 'Delete' && e.ctrlKey) {
                    e.preventDefault();
                    deleteSupplierConfirm(filteredSups[idx]);
                }
            });

            // Insert before empty state
            if (supEmptyState) supListBody.insertBefore(row, supEmptyState);
            else supListBody.appendChild(row);
        });
    }

    function setSupFocus(idx) {
        supFocusedIdx = idx;
        supListBody.querySelectorAll('.sup-row').forEach((r, i) => {
            r.classList.toggle('focused', i === idx);
        });
    }

    function focusSupRow(idx) {
        const rows = supListBody.querySelectorAll('.sup-row');
        if (rows[idx]) {
            setSupFocus(idx);
            rows[idx].focus();
            rows[idx].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        }
    }

    /* ── Open form in 'new' mode ── */
    function openNewSupplier() {
        supEditingId = null;
        clearSupForm();
        showSupForm(true);
        if (supDeleteBtn) supDeleteBtn.classList.add('hidden');
        supFlowFields[0]?.focus();
    }

    /* ── Open form in 'edit' mode ── */
    function openSupplierForEdit(sup) {
        supEditingId = sup.id;
        populateSupForm(sup);
        showSupForm(true);
        if (supDeleteBtn) supDeleteBtn.classList.remove('hidden');

        // Highlight active row
        supListBody.querySelectorAll('.sup-row').forEach(r => {
            r.classList.toggle('active', r.dataset.id === sup.id);
        });

        supFlowFields[0]?.focus();
    }

    /* ── Show / hide form ── */
    function showSupForm(show) {
        if (supFormEl) supFormEl.classList.toggle('hidden', !show);
        if (supPlaceholder) supPlaceholder.style.display = show ? 'none' : '';
    }

    /* ── Populate form fields from supplier object ── */
    function populateSupForm(sup) {
        const map = {
            supName:     sup.name || '',
            supContact:  sup.contactPerson || '',
            supPhone:    sup.phone || '',
            supAltPhone: sup.altPhone || '',
            supEmail:    sup.email || '',
            supGstin:    sup.gstin || '',
            supPan:      sup.pan || '',
            supLicense:  sup.licenseNo || '',
            supAddress:  sup.address || '',
            supCity:     sup.city || '',
            supDistrict: sup.district || '',
            supState:    sup.state || 'Maharashtra',
            supPin:      sup.pinCode || '',
            supBank:     sup.bankName || '',
            supAccount:  sup.accountNo || '',
            supIfsc:     sup.ifsc || '',
            supRemarks:  sup.remarks || '',
        };
        Object.entries(map).forEach(([id, val]) => {
            const el = document.getElementById(id);
            if (el) el.value = val;
        });
    }

    /* ── Clear form ── */
    function clearSupForm() {
        supFlowFields.forEach(f => f.value = '');
        const stateEl = document.getElementById('supState');
        if (stateEl) stateEl.value = 'Maharashtra';
    }

    /* ── Gather form data ── */
    function gatherSupFormData() {
        return {
            id: supEditingId || ('SUP_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6)),
            name:          document.getElementById('supName')?.value.trim() || '',
            contactPerson: document.getElementById('supContact')?.value.trim() || '',
            phone:         document.getElementById('supPhone')?.value.trim() || '',
            altPhone:      document.getElementById('supAltPhone')?.value.trim() || '',
            email:         document.getElementById('supEmail')?.value.trim() || '',
            gstin:         document.getElementById('supGstin')?.value.trim().toUpperCase() || '',
            pan:           document.getElementById('supPan')?.value.trim().toUpperCase() || '',
            licenseNo:     document.getElementById('supLicense')?.value.trim() || '',
            address:       document.getElementById('supAddress')?.value.trim() || '',
            city:          document.getElementById('supCity')?.value.trim() || '',
            district:      document.getElementById('supDistrict')?.value.trim() || '',
            state:         document.getElementById('supState')?.value.trim() || 'Maharashtra',
            pinCode:       document.getElementById('supPin')?.value.trim() || '',
            bankName:      document.getElementById('supBank')?.value.trim() || '',
            accountNo:     document.getElementById('supAccount')?.value.trim() || '',
            ifsc:          document.getElementById('supIfsc')?.value.trim().toUpperCase() || '',
            remarks:       document.getElementById('supRemarks')?.value.trim() || '',
        };
    }

    /* ── Save supplier ── */
    async function saveSupplier() {
        const data = gatherSupFormData();

        if (!data.name) {
            document.getElementById('supName')?.focus();
            showTpToast('Supplier name is required', true);
            return;
        }
        if (!data.phone) {
            document.getElementById('supPhone')?.focus();
            showTpToast('Phone number is required', true);
            return;
        }

        try {
            const result = await window.electronAPI.saveSupplier({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || '',
                supplier: data,
            });
            if (result.success) {
                showTpToast(`Supplier "${data.name}" saved`);
                await loadSuppliers();
                cancelSupForm();
            } else {
                showTpToast('Save failed: ' + result.error, true);
            }
        } catch (err) {
            showTpToast('Save error: ' + err.message, true);
        }
    }

    /* ── Delete supplier ── */
    async function deleteSupplierConfirm(sup) {
        const target = sup || (supEditingId && allSuppliers.find(s => s.id === supEditingId));
        if (!target) return;
        // Simple confirm (no native dialog needed for now)
        const proceed = confirm(`Delete supplier "${target.name}"?`);
        if (!proceed) return;

        try {
            const result = await window.electronAPI.deleteSupplier({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear || '',
                supplierId: target.id,
            });
            if (result.success) {
                showTpToast(`Supplier "${target.name}" deleted`);
                await loadSuppliers();
                cancelSupForm();
            } else {
                showTpToast('Delete failed: ' + result.error, true);
            }
        } catch (err) {
            showTpToast('Delete error: ' + err.message, true);
        }
    }

    /* ── Cancel form ── */
    function cancelSupForm() {
        supEditingId = null;
        clearSupForm();
        showSupForm(false);
        supListBody.querySelectorAll('.sup-row').forEach(r => r.classList.remove('active'));
        // Refocus list
        const rows = supListBody.querySelectorAll('.sup-row');
        if (rows.length > 0) focusSupRow(Math.min(supFocusedIdx, rows.length - 1));
        else if (supSearchEl) supSearchEl.focus();
    }

    /* ── Wire buttons ── */
    if (supNewBtn) supNewBtn.addEventListener('click', openNewSupplier);
    if (supSaveBtn) supSaveBtn.addEventListener('click', saveSupplier);
    if (supCancelBtn) supCancelBtn.addEventListener('click', cancelSupForm);
    if (supDeleteBtn) supDeleteBtn.addEventListener('click', () => deleteSupplierConfirm());

    /* ── Sequential field navigation (Enter key through form) ── */
    supFlowFields.forEach((field, idx) => {
        field.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                if (idx < supFlowFields.length - 1) {
                    supFlowFields[idx + 1].focus();
                } else {
                    // Last field → trigger save
                    saveSupplier();
                }
            } else if (e.key === 'Escape') {
                e.preventDefault();
                cancelSupForm();
            }
        });
    });

    /* ── Search ── */
    if (supSearchEl) {
        supSearchEl.addEventListener('input', () => {
            const q = supSearchEl.value.toLowerCase();
            filteredSups = allSuppliers.filter(s =>
                (s.name || '').toLowerCase().includes(q) ||
                (s.city || '').toLowerCase().includes(q) ||
                (s.phone || '').includes(q) ||
                (s.gstin || '').toLowerCase().includes(q)
            );
            renderSupplierList();
        });

        supSearchEl.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                focusSupRow(0);
            } else if (e.key === 'Enter' && filteredSups.length > 0) {
                e.preventDefault();
                const target = filteredSups[supFocusedIdx >= 0 ? supFocusedIdx : 0];
                if (target) openSupplierForEdit(target);
            }
        });
    }

    /* ── Global shortcuts for supplier panel ── */
    document.addEventListener('keydown', (e) => {
        const supPanel = document.getElementById('sub-add-supplier');
        if (!supPanel || !supPanel.classList.contains('active')) return;

        // F2 = New
        if (e.key === 'F2') {
            e.preventDefault();
            openNewSupplier();
            return;
        }

        // Ctrl+S = Save (when supplier form is open)
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
            if (supFormEl && !supFormEl.classList.contains('hidden')) {
                e.preventDefault();
                saveSupplier();
                return;
            }
        }

        // Ctrl+Delete = Delete selected
        if (e.ctrlKey && e.key === 'Delete') {
            if (supEditingId) {
                e.preventDefault();
                deleteSupplierConfirm();
                return;
            }
        }

        // Esc = Cancel form or clear search
        if (e.key === 'Escape') {
            if (supFormEl && !supFormEl.classList.contains('hidden')) {
                e.preventDefault();
                cancelSupForm();
            } else if (supSearchEl && supSearchEl.value) {
                e.preventDefault();
                supSearchEl.value = '';
                supSearchEl.dispatchEvent(new Event('input'));
                supSearchEl.focus();
            }
        }
    });

    // ═══════════════════════════════════════════════════════
    //  DAILY SALE ENTRY
    // ═══════════════════════════════════════════════════════

    const VAT_RATE = 10;
    const MAX_BTL_PER_CUST = 6;

    // Data state
    let dsProducts = [];
    let dsShortcuts = [];
    let dsMrpMap = {};
    let dsStockMap = {};
    let dsCustomerList = [];
    let dsCurrentCustIdx = 0;
    let dsBillNumber = 0;
    let dsTodaySales = [];
    let dsInitialized = false;

    // Entry state
    let dsBillItems = [];
    let dsAcFocusIdx = -1;
    let dsSelectedProduct = null;

    // DOM refs
    const dsDateEl        = document.getElementById('dsDate');
    const dsItemInput     = document.getElementById('dsItemInput');
    const dsAcDropdown    = document.getElementById('dsAcDropdown');
    const dsItemBrand     = document.getElementById('dsItemBrand');
    const dsItemSize      = document.getElementById('dsItemSize');
    const dsItemMrp       = document.getElementById('dsItemMrp');
    const dsItemStock     = document.getElementById('dsItemStock');
    const dsItemQty       = document.getElementById('dsItemQty');
    const dsAddItemBtn    = document.getElementById('dsAddItemBtn');
    const dsTableBody     = document.getElementById('dsTableBody');
    const dsEmptyState    = document.getElementById('dsEmptyState');
    const dsTotalQtyEl    = document.getElementById('dsTotalQty');
    const dsSubTotalEl    = document.getElementById('dsSubTotal');
    const dsTotalVatEl    = document.getElementById('dsTotalVat');
    const dsGrandTotalEl  = document.getElementById('dsGrandTotal');
    const dsBillsNeededEl = document.getElementById('dsBillsNeeded');
    const dsSaveBillBtn   = document.getElementById('dsSaveBillBtn');
    const dsClearBillBtn  = document.getElementById('dsClearBillBtn');

    /* ══════ DATE CONFIRM DIALOG ══════ */
    const dsDateModal       = document.getElementById('dsDateModal');
    const dsDateConfirmEl   = document.getElementById('dsDateConfirm');
    const dsDateConfirmOk   = document.getElementById('dsDateConfirmOk');
    const dsConfirmDateLbl  = document.getElementById('dsConfirmDateLabel');
    const dsDateTodayBtn    = document.getElementById('dsDateTodayBtn');
    const dsDateYesterday   = document.getElementById('dsDateYesterday');

    function todayStr() { return new Date().toISOString().split('T')[0]; }
    function yesterdayStr() { const d = new Date(); d.setDate(d.getDate() - 1); return d.toISOString().split('T')[0]; }

    function fmtDateDisplay(dateStr) {
        const d = new Date(dateStr + 'T00:00:00');
        return d.toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
    }

    function updateConfirmLabel() {
        if (!dsDateConfirmEl || !dsConfirmDateLbl) return;
        const v = dsDateConfirmEl.value || todayStr();
        dsConfirmDateLbl.textContent = fmtDateDisplay(v);
    }

    function showDsDateModal() {
        if (!dsDateModal) return;
        // Pre-fill with today
        if (dsDateConfirmEl) {
            dsDateConfirmEl.value = todayStr();
            dsDateConfirmEl.addEventListener('change', updateConfirmLabel);
        }
        updateConfirmLabel();
        dsDateModal.classList.remove('hidden');
        setTimeout(() => dsDateConfirmEl?.focus(), 60);
    }

    function hideDsDateModal() {
        if (dsDateModal) dsDateModal.classList.add('hidden');
    }

    function confirmDsDate() {
        const val = dsDateConfirmEl?.value || todayStr();
        if (dsDateEl) dsDateEl.value = val;
        hideDsDateModal();
        loadDailySaleData().then(() => {
            setTimeout(() => dsItemInput?.focus(), 80);
        }).catch(() => {
            setTimeout(() => dsItemInput?.focus(), 80);
        });
    }

    if (dsDateTodayBtn) {
        dsDateTodayBtn.addEventListener('click', () => {
            if (dsDateConfirmEl) dsDateConfirmEl.value = todayStr();
            updateConfirmLabel();
        });
    }
    if (dsDateYesterday) {
        dsDateYesterday.addEventListener('click', () => {
            if (dsDateConfirmEl) dsDateConfirmEl.value = yesterdayStr();
            updateConfirmLabel();
        });
    }
    if (dsDateConfirmOk) {
        dsDateConfirmOk.addEventListener('click', confirmDsDate);
    }

    // Enter key confirms, Esc ignored (must pick a date)
    dsDateModal?.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); confirmDsDate(); }
    });

    /* ══════ DATE SHORTCUT BUTTONS (topbar) ══════ */
    const dsBtnToday     = document.getElementById('dsBtnToday');
    const dsBtnYesterday = document.getElementById('dsBtnYesterday');

    if (dsBtnToday) {
        dsBtnToday.addEventListener('click', () => {
            if (dsDateEl) dsDateEl.value = todayStr();
            loadDailySaleData();
        });
    }
    if (dsBtnYesterday) {
        dsBtnYesterday.addEventListener('click', () => {
            if (dsDateEl) dsDateEl.value = yesterdayStr();
            loadDailySaleData();
        });
    }

    /* ══════ INIT ══════ */
    async function initDailySale() {
        // Always show date confirm dialog on open — pre-filled with today
        showDsDateModal();
    }

    async function loadDailySaleData() {
        const barName = activeBar.barName;
        const financialYear = activeBar.financialYear;

        try {
            const [prodRes, scRes, mrpRes, stockRes, custRes, salesRes, counterRes] = await Promise.all([
                window.electronAPI.getProducts({ barName, financialYear }),
                window.electronAPI.getShortcuts({ barName, financialYear }),
                window.electronAPI.getMrpMaster({ barName, financialYear }),
                window.electronAPI.getCurrentStock({ barName, financialYear, asOfDate: dsDateEl?.value }),
                window.electronAPI.getCustomers({ barName, financialYear }),
                window.electronAPI.getDailySales({ barName, financialYear, date: dsDateEl?.value }),
                window.electronAPI.getSaleCounter({ barName, financialYear })
            ]);

            if (prodRes.success) dsProducts = prodRes.products || [];
            if (scRes.success) dsShortcuts = scRes.shortcuts || [];

            dsMrpMap = {};
            if (mrpRes.success) {
                for (const m of (mrpRes.entries || [])) {
                    dsMrpMap[m.productId] = m.currentMrp || 0;
                }
            }

            dsStockMap = {};
            if (stockRes.success) {
                for (const s of (stockRes.stock || [])) {
                    dsStockMap[s.code] = s.totalBtl || 0;
                }
            }

            if (custRes.success) dsCustomerList = custRes.customers || [];
            if (salesRes.success) dsTodaySales = salesRes.sales || [];

            // Deduct already-saved bills for the selected date so stock reflects
            // what was available at the start of that day minus what has already been sold
            for (const bill of dsTodaySales) {
                for (const item of (bill.items || [])) {
                    if (dsStockMap[item.code] !== undefined) {
                        dsStockMap[item.code] -= (item.qty || 0);
                    }
                }
            }

            const counter = counterRes.success ? counterRes.counter : null;
            const today = dsDateEl?.value;
            if (counter && counter.rotation && counter.rotation.date === today) {
                dsCurrentCustIdx = counter.rotation.customerIndex || 0;
            } else {
                dsCurrentCustIdx = 0;
            }
            dsBillNumber = 0;

            dsInitialized = true;
            dsBillItems = [];
            dsSelectedProduct = null;
            clearDsEntryFields();
            renderDsBillItems();
            updateDsTotals();
            renderSavedBills();
            updateDsDaySummary();

        } catch (err) {
            console.error('initDailySale error:', err);
        }
    }

    /* ══════ AUTOCOMPLETE ══════ */
    function dsSearchItems(query) {
        const q = query.toLowerCase().trim();
        if (!q) return [];
        const results = [];
        const seen = new Set();

        for (const sc of dsShortcuts) {
            if (sc.code && sc.code.toLowerCase().startsWith(q) && !seen.has(sc.productId)) {
                seen.add(sc.productId);
                const prod = dsProducts.find(p => p.id === sc.productId);
                const mrp = dsMrpMap[sc.productId] || sc.mrp || (prod ? prod.mrp : 0);
                const stock = dsStockMap[sc.prodCode || (prod ? prod.code : '')] || 0;
                if (stock <= 0) continue; // only show products with stock
                const cat = sc.category || (prod ? prod.category : '');
                if (isBeerShopee && (cat || '').toUpperCase() === 'SPIRITS') continue; // Beer Shopee: no spirits
                results.push({ productId: sc.productId, shortcode: sc.code, brandName: sc.brandName || (prod ? prod.brandName : ''), code: sc.prodCode || (prod ? prod.code : ''), size: sc.size || (prod ? prod.size : ''), category: cat, mrp, stock });
            }
            if (results.length >= 15) break;
        }
        if (results.length < 15) {
            for (const p of dsProducts) {
                if (seen.has(p.id)) continue;
                if (isBeerShopee && (p.category || '').toUpperCase() === 'SPIRITS') continue; // Beer Shopee: no spirits
                if ((p.brandName && p.brandName.toLowerCase().includes(q)) || (p.code && p.code.toLowerCase().includes(q))) {
                    seen.add(p.id);
                    const mrp = dsMrpMap[p.id] || p.mrp || 0;
                    const stock = dsStockMap[p.code] || 0;
                    if (stock <= 0) continue; // only show products with stock
                    const sc = dsShortcuts.find(s => s.productId === p.id);
                    results.push({ productId: p.id, shortcode: sc ? sc.code : '', brandName: p.brandName, code: p.code, size: p.size, category: p.category, mrp, stock });
                }
                if (results.length >= 15) break;
            }
        }
        return results;
    }

    function renderDsAcDropdown(results) {
        if (!dsAcDropdown) return;
        if (results.length === 0) { dsAcDropdown.classList.add('hidden'); dsAcFocusIdx = -1; return; }
        dsAcDropdown.innerHTML = results.map((r, i) => {
            const stockClass = r.stock <= 0 ? 'out-of-stock' : '';
            const stockText = r.stock > 0 ? `${r.stock} btl` : 'Out of stock';
            return `<div class="ds-ac-item${i === dsAcFocusIdx ? ' focused' : ''}" data-idx="${i}">
                <span class="ds-ac-item-code">${escDS(r.shortcode)}</span>
                <span class="ds-ac-item-brand">${escDS(r.brandName)}</span>
                <span class="ds-ac-item-size">${escDS(r.size)}</span>
                <span class="ds-ac-item-mrp">₹${r.mrp}</span>
                <span class="ds-ac-item-stock ${stockClass}">${stockText}</span>
            </div>`;
        }).join('');
        dsAcDropdown.classList.remove('hidden');
        dsAcDropdown.querySelectorAll('.ds-ac-item').forEach(el => {
            el.addEventListener('mousedown', (e) => { e.preventDefault(); dsSelectAcItem(results[parseInt(el.dataset.idx)]); });
        });
    }

    function dsSelectAcItem(item) {
        dsSelectedProduct = item;
        if (dsItemInput) dsItemInput.value = item.shortcode || item.brandName;
        if (dsItemBrand) dsItemBrand.value = item.brandName;
        if (dsItemSize) dsItemSize.value = item.size;
        if (dsItemMrp) dsItemMrp.value = item.mrp;
        if (dsItemStock) {
            dsItemStock.value = item.stock > 0 ? `${item.stock} btl` : 'NIL';
            dsItemStock.classList.toggle('ds-stock-low', item.stock <= 0);
        }
        if (dsAcDropdown) dsAcDropdown.classList.add('hidden');
        dsAcFocusIdx = -1;
        if (dsItemQty) { dsItemQty.value = 1; dsItemQty.focus(); dsItemQty.select(); }
    }

    function clearDsEntryFields() {
        dsSelectedProduct = null;
        if (dsItemInput) dsItemInput.value = '';
        if (dsItemBrand) dsItemBrand.value = '';
        if (dsItemSize) dsItemSize.value = '';
        if (dsItemMrp) dsItemMrp.value = '';
        if (dsItemStock) { dsItemStock.value = ''; dsItemStock.classList.remove('ds-stock-low'); }
        if (dsItemQty) dsItemQty.value = 1;
        if (dsAcDropdown) dsAcDropdown.classList.add('hidden');
        dsAcFocusIdx = -1;
    }

    /* ══════ ADD ITEM (no rotation here — just collect) ══════ */
    function dsAddItem() {
        if (!dsSelectedProduct) { if (dsItemInput) dsItemInput.focus(); return; }
        const qty = parseInt(dsItemQty?.value) || 1;
        if (qty <= 0) return;

        // Stock check
        const availStock = dsStockMap[dsSelectedProduct.code] || 0;
        const alreadyInBill = dsBillItems.filter(i => i.code === dsSelectedProduct.code).reduce((s, i) => s + i.qty, 0);
        if (availStock - alreadyInBill < qty) {
            alert(`Insufficient stock! Available: ${availStock - alreadyInBill} bottles for ${dsSelectedProduct.brandName}`);
            return;
        }

        const mrp = dsSelectedProduct.mrp || 0;
        const amount = qty * mrp;
        const vatAmt = Math.round(amount * VAT_RATE / 100 * 100) / 100;

        // Merge if same product already in list
        const existing = dsBillItems.find(i => i.code === dsSelectedProduct.code);
        if (existing) {
            existing.qty += qty;
            existing.amount = existing.qty * existing.mrp;
            existing.vatAmt = Math.round(existing.amount * VAT_RATE / 100 * 100) / 100;
            existing.lineTotal = existing.amount + existing.vatAmt;
        } else {
            dsBillItems.push({
                productId: dsSelectedProduct.productId,
                brandName: dsSelectedProduct.brandName,
                code: dsSelectedProduct.code,
                size: dsSelectedProduct.size,
                category: dsSelectedProduct.category || '',
                qty, mrp, amount,
                vatPct: VAT_RATE,
                vatAmt,
                lineTotal: amount + vatAmt
            });
        }

        renderDsBillItems();
        updateDsTotals();
        clearDsEntryFields();
        if (dsItemInput) dsItemInput.focus();
    }

    /* ══════ RENDER ITEMS TABLE ══════ */
    function renderDsBillItems() {
        if (!dsTableBody) return;
        dsTableBody.innerHTML = '';
        if (dsBillItems.length === 0) { if (dsEmptyState) dsEmptyState.style.display = ''; return; }
        if (dsEmptyState) dsEmptyState.style.display = 'none';

        dsBillItems.forEach((item, idx) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${idx + 1}</td>
                <td class="ds-td-brand">${escDS(item.brandName)}</td>
                <td class="ds-td-code">${escDS(item.code)}</td>
                <td>${escDS(item.size)}</td>
                <td class="ds-td-right">${item.qty}</td>
                <td class="ds-td-right">₹${item.mrp.toFixed(2)}</td>
                <td class="ds-td-right">₹${item.amount.toFixed(2)}</td>
                <td class="ds-td-right ds-td-vat">₹${item.vatAmt.toFixed(2)}</td>
                <td class="ds-td-right ds-td-total">₹${item.lineTotal.toFixed(2)}</td>
                <td><button class="ds-row-del" data-idx="${idx}" title="Remove">✕</button></td>
            `;
            dsTableBody.appendChild(tr);
        });

        dsTableBody.querySelectorAll('.ds-row-del').forEach(btn => {
            btn.addEventListener('click', () => {
                dsBillItems.splice(parseInt(btn.dataset.idx), 1);
                renderDsBillItems();
                updateDsTotals();
            });
        });
    }

    function updateDsTotals() {
        const totalQty = dsBillItems.reduce((s, i) => s + i.qty, 0);
        const subTotal = dsBillItems.reduce((s, i) => s + i.amount, 0);
        const totalVat = dsBillItems.reduce((s, i) => s + i.vatAmt, 0);
        const grandTotal = subTotal + totalVat;
        const billsNeeded = dsCustomerList.length > 0 ? Math.ceil(totalQty / MAX_BTL_PER_CUST) : (totalQty > 0 ? 1 : 0);

        if (dsTotalQtyEl) dsTotalQtyEl.textContent = totalQty;
        if (dsSubTotalEl) dsSubTotalEl.textContent = `₹ ${subTotal.toFixed(2)}`;
        if (dsTotalVatEl) dsTotalVatEl.textContent = `₹ ${totalVat.toFixed(2)}`;
        if (dsGrandTotalEl) dsGrandTotalEl.textContent = `₹ ${grandTotal.toFixed(2)}`;
        if (dsBillsNeededEl) dsBillsNeededEl.textContent = billsNeeded;
    }

    /* ══════ SAVE & GENERATE BILLS (auto-split by 3-btl customer rotation) ══════ */
    async function dsSaveAndGenerate() {
        if (dsBillItems.length === 0) { alert('Add at least one item first.'); return; }

        // Flatten all items into individual bottles: [{...item, qty:1}, ...]
        const allBottles = [];
        for (const item of dsBillItems) {
            for (let i = 0; i < item.qty; i++) {
                allBottles.push({ ...item, qty: 1, amount: item.mrp, vatAmt: Math.round(item.mrp * VAT_RATE / 100 * 100) / 100, lineTotal: item.mrp + Math.round(item.mrp * VAT_RATE / 100 * 100) / 100 });
            }
        }

        // Split bottles into chunks of MAX_BTL_PER_CUST
        const chunks = [];
        for (let i = 0; i < allBottles.length; i += MAX_BTL_PER_CUST) {
            chunks.push(allBottles.slice(i, i + MAX_BTL_PER_CUST));
        }

        const fy = activeBar.financialYear || '';
        const fmatch = fy.match(/(\d{4})/);
        const fyYear = fmatch ? fmatch[1] : new Date().getFullYear();
        const billDate = dsDateEl?.value || new Date().toISOString().split('T')[0];
        const barName = activeBar.barName;
        const financialYear = activeBar.financialYear;

        // Bill numbering source-of-truth: existing saved sales
        // If none found => first bill number should be 0000
        try {
            const allSalesRes = await window.electronAPI.getDailySales({ barName, financialYear });
            const allSales = allSalesRes.success ? (allSalesRes.sales || []) : [];
            let maxBillNumber = 0;

            for (const savedSale of allSales) {
                const billNo = String(savedSale.billNo || '');
                const match = billNo.match(/(\d+)$/);
                if (!match) continue;
                const num = Number(match[1]);
                if (Number.isFinite(num)) {
                    maxBillNumber = Math.max(maxBillNumber, num);
                }
            }

            dsBillNumber = maxBillNumber;
        } catch (err) {
            console.error('Failed to derive bill number from saved sales:', err);
            dsBillNumber = Math.max(0, Number(dsBillNumber) || 0);
        }

        const generatedBills = [];
        let saveError = null;

        for (const chunk of chunks) {
            dsBillNumber++;
            const billNo = `DS/${fyYear}/${String(dsBillNumber).padStart(4, '0')}`;

            // Assign customer (rotate after each bill)
            const cust = dsCustomerList.length > 0
                ? dsCustomerList[dsCurrentCustIdx % dsCustomerList.length]
                : { id: 'walk-in', name: 'Walk-in', licNo: '' };

            // Merge duplicate items within the chunk
            const merged = [];
            for (const btl of chunk) {
                const ex = merged.find(m => m.code === btl.code);
                if (ex) {
                    ex.qty += 1;
                    ex.amount = ex.qty * ex.mrp;
                    ex.vatAmt = Math.round(ex.amount * VAT_RATE / 100 * 100) / 100;
                    ex.lineTotal = ex.amount + ex.vatAmt;
                } else {
                    merged.push({ ...btl });
                }
            }

            const totalQty = merged.reduce((s, i) => s + i.qty, 0);
            const subTotal = merged.reduce((s, i) => s + i.amount, 0);
            const totalVat = merged.reduce((s, i) => s + i.vatAmt, 0);
            const grandTotal = subTotal + totalVat;

            const sale = {
                id: 'DS_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
                billNo, billDate,
                customerId: cust.id, customerName: cust.name, customerLicNo: cust.licNo,
                items: merged,
                totalQty, subTotal, vatPct: VAT_RATE, totalVat, grandTotal,
                paymentMode: 'cash', paymentStatus: 'paid'
            };

            try {
                const res = await window.electronAPI.saveDailySale({ barName, financialYear, sale });
                if (res.success) {
                    generatedBills.push(sale);
                    dsTodaySales.push(sale);
                    // Deduct stock locally
                    for (const item of merged) {
                        if (dsStockMap[item.code] !== undefined) dsStockMap[item.code] -= item.qty;
                    }
                } else {
                    saveError = res.error;
                    dsBillNumber--;
                    break;
                }
            } catch (err) {
                saveError = err.message;
                dsBillNumber--;
                break;
            }

            // Rotate to next customer
            dsCurrentCustIdx++;
            if (dsCustomerList.length > 0 && dsCurrentCustIdx >= dsCustomerList.length) dsCurrentCustIdx = 0;
        }

        // Save counter state
        try {
            await window.electronAPI.saveSaleCounter({
                barName, financialYear,
                counter: {
                    lastBillNumber: dsBillNumber,
                    prefix: 'DS', fyYear,
                    rotation: { date: billDate, customerIndex: dsCurrentCustIdx, maxPerCustomer: MAX_BTL_PER_CUST }
                }
            });
        } catch (e) { console.error('Save counter error:', e); }

        if (saveError) {
            alert('Error saving bill: ' + saveError + `\n${generatedBills.length} bills were saved before the error.`);
        }

        // Clear entry area
        dsBillItems = [];
        dsSelectedProduct = null;
        clearDsEntryFields();
        renderDsBillItems();
        updateDsTotals();
        renderSavedBills();
        updateDsDaySummary();

        // Scroll to bills section
        if (dsBillsList) dsBillsList.scrollIntoView({ behavior: 'smooth', block: 'start' });
        if (dsItemInput) dsItemInput.focus();
    }

    /* ══════ STUBS — bills display moved to Billing page ══════ */
    function renderSavedBills() { /* no-op, bills shown in Billing sub-view */ }
    function updateDsDaySummary() { /* no-op */ }

    /* ══════ CLEAR ══════ */
    function dsClearBill() {
        if (dsBillItems.length > 0 && !confirm('Clear all items? Unsaved entries will be lost.')) return;
        dsBillItems = [];
        clearDsEntryFields();
        renderDsBillItems();
        updateDsTotals();
        if (dsItemInput) dsItemInput.focus();
    }

    /* ══════ EVENT WIRING ══════ */
    if (dsItemInput) {
        let dsSearchDebounce;
        dsItemInput.addEventListener('input', () => {
            // User is typing new text — clear stale selection
            dsSelectedProduct = null;
            clearTimeout(dsSearchDebounce);
            dsSearchDebounce = setTimeout(() => {
                const q = dsItemInput.value.trim();
                if (q.length >= 1) { dsAcFocusIdx = -1; renderDsAcDropdown(dsSearchItems(q)); } else { if (dsAcDropdown) dsAcDropdown.classList.add('hidden'); dsAcFocusIdx = -1; }
            }, 120);
        });

        dsItemInput.addEventListener('keydown', (e) => {
            const items = dsAcDropdown ? dsAcDropdown.querySelectorAll('.ds-ac-item') : [];
            const dropdownOpen = dsAcDropdown && !dsAcDropdown.classList.contains('hidden') && items.length > 0;

            if (e.key === 'ArrowDown') {
                e.preventDefault();
                if (items.length > 0) { dsAcFocusIdx = Math.min(dsAcFocusIdx + 1, items.length - 1); items.forEach((el, i) => el.classList.toggle('focused', i === dsAcFocusIdx)); items[dsAcFocusIdx]?.scrollIntoView({ block: 'nearest' }); }
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                if (dsAcFocusIdx > 0) { dsAcFocusIdx--; items.forEach((el, i) => el.classList.toggle('focused', i === dsAcFocusIdx)); items[dsAcFocusIdx]?.scrollIntoView({ block: 'nearest' }); }
            } else if (e.key === 'Enter') {
                e.preventDefault();
                // Cancel pending debounce so dropdown won't re-fire
                clearTimeout(dsSearchDebounce);

                const results = dsSearchItems(dsItemInput.value.trim());

                if (dsAcFocusIdx >= 0 && results[dsAcFocusIdx]) {
                    // Arrow-navigated to a specific item — select it
                    dsSelectAcItem(results[dsAcFocusIdx]);
                } else if (dropdownOpen && results.length > 0) {
                    // Dropdown open but no arrow navigation — auto-select first result
                    dsSelectAcItem(results[0]);
                } else if (dsSelectedProduct) {
                    // Product already selected (from a previous pick) — add it
                    dsAddItem();
                } else {
                    // No dropdown, no selection — try exact shortcode match
                    const q = dsItemInput.value.trim().toLowerCase();
                    if (q) {
                        const exact = dsShortcuts.find(s => s.code && s.code.toLowerCase() === q);
                        if (exact) {
                            const prod = dsProducts.find(p => p.id === exact.productId);
                            const mrp = dsMrpMap[exact.productId] || exact.mrp || (prod ? prod.mrp : 0);
                            const stock = dsStockMap[exact.prodCode || (prod ? prod.code : '')] || 0;
                            dsSelectAcItem({ productId: exact.productId, shortcode: exact.code, brandName: exact.brandName || (prod ? prod.brandName : ''), code: exact.prodCode || (prod ? prod.code : ''), size: exact.size || (prod ? prod.size : ''), category: exact.category || (prod ? prod.category : ''), mrp, stock });
                        } else if (results.length > 0) {
                            // Partial match — select best result
                            dsSelectAcItem(results[0]);
                        }
                    }
                }
            } else if (e.key === 'Escape') {
                if (dsAcDropdown) dsAcDropdown.classList.add('hidden'); dsAcFocusIdx = -1;
            }
        });

        dsItemInput.addEventListener('blur', () => { setTimeout(() => { if (dsAcDropdown) dsAcDropdown.classList.add('hidden'); dsAcFocusIdx = -1; }, 200); });
    }

    if (dsItemQty) { dsItemQty.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); dsAddItem(); } }); }
    if (dsAddItemBtn) dsAddItemBtn.addEventListener('click', dsAddItem);
    if (dsSaveBillBtn) dsSaveBillBtn.addEventListener('click', () => dsSaveAndGenerate());
    if (dsClearBillBtn) dsClearBillBtn.addEventListener('click', dsClearBill);

    if (dsDateEl) { dsDateEl.addEventListener('change', () => loadDailySaleData()); }

    document.getElementById('sub-add-daily-sale')?.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); dsSaveAndGenerate(); }
    });

    function escDS(str) {
        return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    initDailySale();

    // ═══════════════════════════════════════════════════════
    //  BILLING — Legal page preview (12 bills/page)
    // ═══════════════════════════════════════════════════════

    const BL_BILLS_PER_PAGE = 16;

    // State
    let blAllSales = [];
    let blPagedSales = [];
    let blCurrentPage = 0;

    // DOM refs
    const blDateFilterEl   = document.getElementById('blDateFilter');
    const blMonthFilterEl  = document.getElementById('blMonthFilter');
    const blLoadBtn        = document.getElementById('blLoadBtn');
    const blTodayBtn       = document.getElementById('blTodayBtn');
    const blModeDateBtn    = document.getElementById('blModeDate');
    const blModeMonthBtn   = document.getElementById('blModeMonth');
    const blModeAllBtn     = document.getElementById('blModeAll');
    const blClearBtn       = document.getElementById('blClearBtn');
    const blPagePrevBtn    = document.getElementById('blPagePrev');
    const blPageNextBtn    = document.getElementById('blPageNext');
    const blPrintBtn       = document.getElementById('blPrintBtn');
    const blPageIndicator  = document.getElementById('blPageIndicator');
    const blSheetGrid      = document.getElementById('blSheetGrid');
    const blLegalSheet     = document.getElementById('blLegalSheet');
    const blEmpty          = document.getElementById('blEmpty');
    const blStatBillsEl    = document.getElementById('blStatBills');
    const blStatPagesEl    = document.getElementById('blStatPages');
    const blStatBottlesEl  = document.getElementById('blStatBottles');
    const blStatAmountEl   = document.getElementById('blStatAmount');
    const blVatToggle      = document.getElementById('blVatToggle');

    // ════════════════════════════════════════════════════════════════
    // ██  VAT TAX REPORT
    // ════════════════════════════════════════════════════════════════

    const vatDateFilter  = document.getElementById('vatDateFilter');
    const vatMonthFilter = document.getElementById('vatMonthFilter');
    const vatLoadBtn     = document.getElementById('vatLoadBtn');
    const vatClearBtn    = document.getElementById('vatClearBtn');
    const vatTableBody   = document.getElementById('vatTableBody');
    const vatTableEl     = document.getElementById('vatTable');
    const vatEmpty       = document.getElementById('vatEmpty');

    // Stat IDs
    const vatStatBills   = document.getElementById('vatStatBills');
    const vatStatSaleAmt = document.getElementById('vatStatSaleAmt');
    const vatStatVat     = document.getElementById('vatStatVat');
    const vatStatTotal   = document.getElementById('vatStatTotal');

    // Footer IDs
    const vatFootSpirits   = document.getElementById('vatFootSpirits');
    const vatFootWine      = document.getElementById('vatFootWine');
    const vatFootMildBeer  = document.getElementById('vatFootMildBeer');
    const vatFootFermented = document.getElementById('vatFootFermented');
    const vatFootMml       = document.getElementById('vatFootMml');
    const vatFootSaleAmt   = document.getElementById('vatFootSaleAmt');
    const vatFootVat       = document.getElementById('vatFootVat');
    const vatFootTotal     = document.getElementById('vatFootTotal');

    function initVatReport() {
        const today = new Date().toISOString().split('T')[0];
        const month = today.substring(0, 7); // YYYY-MM
        if (vatDateFilter && !vatDateFilter.value) vatDateFilter.value = '';
        if (vatMonthFilter && !vatMonthFilter.value) vatMonthFilter.value = month;
        loadVatReport();
    }

    // Mutual exclusion: picking a date clears month, and vice-versa
    if (vatDateFilter) {
        vatDateFilter.addEventListener('change', () => {
            if (vatDateFilter.value && vatMonthFilter) vatMonthFilter.value = '';
        });
    }
    if (vatMonthFilter) {
        vatMonthFilter.addEventListener('change', () => {
            if (vatMonthFilter.value && vatDateFilter) vatDateFilter.value = '';
        });
    }

    if (vatLoadBtn) vatLoadBtn.addEventListener('click', () => loadVatReport());
    if (vatClearBtn) {
        vatClearBtn.addEventListener('click', () => {
            if (vatDateFilter) vatDateFilter.value = '';
            if (vatMonthFilter) vatMonthFilter.value = '';
            clearVatReport();
        });
    }

    function clearVatReport() {
        if (vatTableBody) vatTableBody.innerHTML = '';
        if (vatTableEl) vatTableEl.style.display = 'none';
        if (vatEmpty) vatEmpty.style.display = 'flex';
        [vatStatBills, vatStatSaleAmt, vatStatVat, vatStatTotal].forEach(el => {
            if (el) el.textContent = el.id.includes('Bills') ? '0' : '₹ 0';
        });
        [vatFootSpirits, vatFootWine, vatFootMildBeer, vatFootFermented, vatFootMml].forEach(el => {
            if (el) el.textContent = '0';
        });
        [vatFootSaleAmt, vatFootVat, vatFootTotal].forEach(el => {
            if (el) el.textContent = '₹ 0';
        });
    }

    function fmtCurrency(n) {
        return '₹ ' + Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    function fmtDateShort(dateStr) {
        const d = new Date(dateStr + 'T00:00:00');
        return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
    }

    async function loadVatReport() {
        const barName = activeBar.barName;
        const financialYear = activeBar.financialYear || '';
        if (!barName) return;

        try {
            // Load all sales for the FY
            const res = await window.electronAPI.getDailySales({ barName, financialYear });
            if (!res.success) return;

            let sales = res.sales || [];

            // Filter by date or month
            const dateVal = vatDateFilter ? vatDateFilter.value : '';
            const monthVal = vatMonthFilter ? vatMonthFilter.value : '';

            if (dateVal) {
                sales = sales.filter(s => s.billDate === dateVal);
            } else if (monthVal) {
                sales = sales.filter(s => s.billDate && s.billDate.startsWith(monthVal));
            }

            if (sales.length === 0) {
                clearVatReport();
                return;
            }

            // Group by date
            const byDate = {};
            sales.forEach(sale => {
                const d = sale.billDate || 'Unknown';
                if (!byDate[d]) byDate[d] = [];
                byDate[d].push(sale);
            });

            // Sort dates
            const sortedDates = Object.keys(byDate).sort();

            // Build rows
            let html = '';
            let grandSpirits = 0, grandWine = 0, grandMildBeer = 0, grandFermented = 0, grandMml = 0;
            let grandSaleAmt = 0, grandVat = 0, grandTotal = 0;
            let totalBills = 0;

            sortedDates.forEach(date => {
                const daySales = byDate[date];
                totalBills += daySales.length;

                // Get bill range
                const billNos = daySales.map(s => s.billNo || '').filter(Boolean);
                const billFrom = billNos.length > 0 ? billNos[0] : '—';
                const billTo = billNos.length > 0 ? billNos[billNos.length - 1] : '—';

                // Sum categories
                let spirits = 0, wine = 0, mildBeer = 0, fermented = 0, mml = 0;
                let daySaleAmt = 0;

                daySales.forEach(sale => {
                    (sale.items || []).forEach(item => {
                        const cat = (item.category || '').toLowerCase().trim();
                        const amt = Number(item.amount) || 0;

                        if (cat === 'spirits' || cat === 'spirit') {
                            spirits += amt;
                        } else if (cat === 'wine') {
                            wine += amt;
                        } else if (cat === 'mild beer' || cat === 'beer') {
                            mildBeer += amt;
                        } else if (cat === 'fermented beer' || cat === 'fermented') {
                            fermented += amt;
                        } else if (cat === 'mml') {
                            mml += amt;
                        } else {
                            // Default to spirits for unknown categories
                            spirits += amt;
                        }
                    });
                });

                daySaleAmt = spirits + wine + mildBeer + fermented + mml;
                const dayVat = daySaleAmt * 0.10;
                const dayTotal = daySaleAmt + dayVat;

                grandSpirits += spirits;
                grandWine += wine;
                grandMildBeer += mildBeer;
                grandFermented += fermented;
                grandMml += mml;
                grandSaleAmt += daySaleAmt;
                grandVat += dayVat;
                grandTotal += dayTotal;

                html += `<tr>
                    <td class="vat-td-date">${fmtDateShort(date)}</td>
                    <td class="vat-td-bill">${billFrom}</td>
                    <td class="vat-td-bill">${billTo}</td>
                    <td class="vat-td-cat">${spirits ? fmtCurrency(spirits) : '—'}</td>
                    <td class="vat-td-cat">${wine ? fmtCurrency(wine) : '—'}</td>
                    <td class="vat-td-cat">${mildBeer ? fmtCurrency(mildBeer) : '—'}</td>
                    <td class="vat-td-cat">${fermented ? fmtCurrency(fermented) : '—'}</td>
                    <td class="vat-td-cat">${mml ? fmtCurrency(mml) : '—'}</td>
                    <td class="vat-td-amt">${fmtCurrency(daySaleAmt)}</td>
                    <td class="vat-td-vat">${fmtCurrency(dayVat)}</td>
                    <td class="vat-td-total">${fmtCurrency(dayTotal)}</td>
                </tr>`;
            });

            if (vatTableBody) vatTableBody.innerHTML = html;

            // Show table, hide empty
            if (vatTableEl) vatTableEl.style.display = '';
            if (vatEmpty) vatEmpty.style.display = 'none';

            // Update stats
            if (vatStatBills) vatStatBills.textContent = totalBills;
            if (vatStatSaleAmt) vatStatSaleAmt.textContent = fmtCurrency(grandSaleAmt);
            if (vatStatVat) vatStatVat.textContent = fmtCurrency(grandVat);
            if (vatStatTotal) vatStatTotal.textContent = fmtCurrency(grandTotal);

            // Update footer
            if (vatFootSpirits) vatFootSpirits.textContent = fmtCurrency(grandSpirits);
            if (vatFootWine) vatFootWine.textContent = fmtCurrency(grandWine);
            if (vatFootMildBeer) vatFootMildBeer.textContent = fmtCurrency(grandMildBeer);
            if (vatFootFermented) vatFootFermented.textContent = fmtCurrency(grandFermented);
            if (vatFootMml) vatFootMml.textContent = fmtCurrency(grandMml);
            if (vatFootSaleAmt) vatFootSaleAmt.textContent = fmtCurrency(grandSaleAmt);
            if (vatFootVat) vatFootVat.textContent = fmtCurrency(grandVat);
            if (vatFootTotal) vatFootTotal.textContent = fmtCurrency(grandTotal);

        } catch (err) {
            console.error('VAT report error:', err);
        }
    }

    // ════════════════════════════════════════════════════════════════
    // ██  BILLING
    // ════════════════════════════════════════════════════════════════

    function escBL(str) {
        return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    function formatBillDate(dateStr) {
        if (!dateStr) return '—';
        try {
            return new Date(dateStr + 'T00:00:00').toLocaleDateString('en-IN', {
                day: '2-digit',
                month: 'short',
                year: 'numeric'
            });
        } catch (err) {
            return dateStr;
        }
    }

    function paginateBills() {
        blPagedSales = [];
        for (let i = 0; i < blAllSales.length; i += BL_BILLS_PER_PAGE) {
            blPagedSales.push(blAllSales.slice(i, i + BL_BILLS_PER_PAGE));
        }
    }

    function clearBillingPreview() {
        blAllSales = [];
        blPagedSales = [];
        blCurrentPage = 0;

        if (blSheetGrid) blSheetGrid.innerHTML = '';
        if (blLegalSheet) blLegalSheet.classList.add('hidden');
        if (blEmpty) blEmpty.style.display = 'flex';
        if (blPageIndicator) blPageIndicator.textContent = 'Page 0 / 0';
        if (blPagePrevBtn) blPagePrevBtn.disabled = true;
        if (blPageNextBtn) blPageNextBtn.disabled = true;
        if (blPrintBtn) blPrintBtn.disabled = true;
        if (blStatBillsEl) blStatBillsEl.textContent = '0';
        if (blStatPagesEl) blStatPagesEl.textContent = '0';
        if (blStatBottlesEl) blStatBottlesEl.textContent = '0';
        if (blStatAmountEl) blStatAmountEl.textContent = '₹ 0';
    }

    function buildBillPreviewHtml(bill, pageIndex, slotIndex) {
        const barName = activeBar.barName || 'Bar';
        const barAddress = getBillingBarAddress();
        const barLicense = getBillingBarLicense();
        const items = (bill.items || []).slice(0, 4).map((item) => {
            return `<div class="bl-item-row">
                <span class="bl-item-brand">${escBL(item.brandName || 'Item')}</span>
                <span class="bl-item-meta">${escBL(item.size || '')} × ${Number(item.qty || 0)}</span>
            </div>`;
        }).join('');

        const extraItemCount = Math.max(0, (bill.items || []).length - 4);
        const pageBillNo = (pageIndex * BL_BILLS_PER_PAGE) + slotIndex + 1;

        return `<article class="bl-mini-bill">
            <div class="bl-bar-head">
                <div class="bl-bar-name">${escBL(barName)}</div>
                <div class="bl-bar-address">${escBL(barAddress)}</div>
                <div class="bl-bar-license">${escBL(barLicense)}</div>
            </div>
            <header class="bl-mini-head">
                <div class="bl-mini-title">Bill #${escBL(bill.billNo || String(pageBillNo))}</div>
                <div class="bl-mini-date">${escBL(formatBillDate(bill.billDate))}</div>
            </header>
            <div class="bl-mini-customer">${escBL(bill.customerName || 'Walk-in Customer')}</div>
            <div class="bl-mini-license">${escBL(bill.customerLicNo || 'Lic: —')}</div>
            <div class="bl-mini-items">
                ${items || '<div class="bl-item-row"><span class="bl-item-brand">No items</span></div>'}
                ${extraItemCount > 0 ? `<div class="bl-item-more">+ ${extraItemCount} more item(s)</div>` : ''}
            </div>
            <footer class="bl-mini-foot">${blVatToggle?.checked !== false
                ? `<span>Btl: <strong>${Number(bill.totalQty || 0)}</strong></span><span>Sub: <strong>${fmtCurrency(Number(bill.subTotal || 0))}</strong></span><span>VAT(${VAT_RATE}%): <strong>${fmtCurrency(Number(bill.totalVat || 0))}</strong></span><span>Total: <strong>${fmtCurrency(Number(bill.grandTotal || 0))}</strong></span>`
                : `<span>Btl: <strong>${Number(bill.totalQty || 0)}</strong></span><span>Total: <strong>${fmtCurrency(Number(bill.subTotal || 0))}</strong></span>`
            }</footer>
        </article>`;
    }

    function renderBillingPreview() {
        if (!blSheetGrid) return;

        if (blPagedSales.length === 0) {
            clearBillingPreview();
            return;
        }

        if (blCurrentPage < 0) blCurrentPage = 0;
        if (blCurrentPage >= blPagedSales.length) blCurrentPage = blPagedSales.length - 1;

        const pageBills = blPagedSales[blCurrentPage] || [];
        blSheetGrid.innerHTML = pageBills.map((bill, idx) => buildBillPreviewHtml(bill, blCurrentPage, idx)).join('');

        if (blLegalSheet) blLegalSheet.classList.remove('hidden');
        if (blEmpty) blEmpty.style.display = 'none';

        if (blPageIndicator) {
            blPageIndicator.textContent = `Page ${blCurrentPage + 1} / ${blPagedSales.length}`;
        }
        if (blPagePrevBtn) blPagePrevBtn.disabled = blCurrentPage <= 0;
        if (blPageNextBtn) blPageNextBtn.disabled = blCurrentPage >= (blPagedSales.length - 1);
        if (blPrintBtn) blPrintBtn.disabled = pageBills.length === 0;
    }

    function updateBillingStats() {
        const totalBills = blAllSales.length;
        const totalPages = blPagedSales.length;
        const totalBottles = blAllSales.reduce((sum, bill) => sum + Number(bill.totalQty || 0), 0);
        const totalAmount = blAllSales.reduce((sum, bill) => sum + Number(bill.grandTotal || 0), 0);

        if (blStatBillsEl) blStatBillsEl.textContent = String(totalBills);
        if (blStatPagesEl) blStatPagesEl.textContent = String(totalPages);
        if (blStatBottlesEl) blStatBottlesEl.textContent = String(totalBottles);
        if (blStatAmountEl) blStatAmountEl.textContent = fmtCurrency(totalAmount);
    }

    function getBillingFilteredSales(allSales) {
        const dateVal = blDateFilterEl ? blDateFilterEl.value : '';
        const monthVal = blMonthFilterEl ? blMonthFilterEl.value : '';
        const modeAll = blModeAllBtn?.classList.contains('active');

        let filtered = allSales || [];
        if (!modeAll) {
            if (dateVal) {
                filtered = filtered.filter(sale => sale.billDate === dateVal);
            } else if (monthVal) {
                filtered = filtered.filter(sale => sale.billDate && sale.billDate.startsWith(monthVal));
            }
        }

        return filtered.sort((a, b) => {
            const dateA = a.billDate || '';
            const dateB = b.billDate || '';
            if (dateA === dateB) return String(a.billNo || '').localeCompare(String(b.billNo || ''));
            return dateA.localeCompare(dateB);
        });
    }

    function printCurrentBillingPage() {
        if (!blPagedSales.length) return;
        const pageBills = blPagedSales[blCurrentPage] || [];
        if (pageBills.length === 0) return;
        const showVat = blVatToggle?.checked !== false;

        printWithIframe(`<!doctype html>
<html>
<head>
    <meta charset="utf-8" />
    <title>Billing Page ${blCurrentPage + 1}</title>
    <style>
        @page { size: legal landscape; margin: 6mm; }
        * { box-sizing: border-box; }
        body { margin: 0; font-family: Arial, sans-serif; color: #111; }
        .sheet-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 2mm; }
        .bill { border: 1px solid #000; border-radius: 3px; padding: 1.6mm; min-height: 35mm; display: flex; flex-direction: column; }
        .bar-head { border-bottom: 1px solid #000; margin-bottom: 1mm; padding-bottom: 1mm; }
        .bar-name { font-size: 6.6pt; font-weight: 800; line-height: 1.1; }
        .bar-address, .bar-license { font-size: 5.8pt; line-height: 1.1; }
        .head { display: flex; justify-content: space-between; gap: 1.5mm; border-bottom: 1px dashed #555; padding-bottom: 1mm; margin-bottom: 1mm; }
        .title { font-size: 6.6pt; font-weight: 700; }
        .date { font-size: 5.8pt; }
        .customer { font-size: 6.1pt; font-weight: 600; }
        .license { font-size: 5.8pt; margin-bottom: .8mm; }
        .items { flex: 1; font-size: 5.8pt; line-height: 1.12; }
        .item { display: flex; justify-content: space-between; gap: .8mm; margin-bottom: .4mm; }
        .foot { border-top: 1px dashed #555; padding-top: .8mm; display: flex; justify-content: space-between; font-size: 5.9pt; font-weight: 700; }
    </style>
</head>
<body>
    <div class="sheet-grid">
        ${pageBills.map((bill, idx) => {
            const barName = activeBar.barName || 'Bar';
            const barAddress = getBillingBarAddress();
            const barLicense = getBillingBarLicense();
            const items = (bill.items || []).slice(0, 4).map(item => `<div class="item"><span>${escBL(item.brandName || 'Item')}</span><span>${escBL(item.size || '')} × ${Number(item.qty || 0)}</span></div>`).join('');
            const extra = Math.max(0, (bill.items || []).length - 4);
            const pageBillNo = (blCurrentPage * BL_BILLS_PER_PAGE) + idx + 1;
            return `<article class="bill">
                <div class="bar-head">
                    <div class="bar-name">${escBL(barName)}</div>
                    <div class="bar-address">${escBL(barAddress)}</div>
                    <div class="bar-license">${escBL(barLicense)}</div>
                </div>
                <div class="head">
                    <div class="title">Bill #${escBL(bill.billNo || String(pageBillNo))}</div>
                    <div class="date">${escBL(formatBillDate(bill.billDate))}</div>
                </div>
                <div class="customer">${escBL(bill.customerName || 'Walk-in Customer')}</div>
                <div class="license">${escBL(bill.customerLicNo || 'Lic: —')}</div>
                <div class="items">${items}${extra > 0 ? `<div class="item"><span>+ ${extra} more item(s)</span></div>` : ''}</div>
                <div class="foot">${showVat ? `<span>Btl: ${Number(bill.totalQty || 0)}</span><span>Sub: ${fmtCurrency(Number(bill.subTotal || 0))}</span><span>VAT(${VAT_RATE}%): ${fmtCurrency(Number(bill.totalVat || 0))}</span><span>Total: ${fmtCurrency(Number(bill.grandTotal || 0))}</span>` : `<span>Btl: ${Number(bill.totalQty || 0)}</span><span>Total: ${fmtCurrency(Number(bill.subTotal || 0))}</span>`}</div>
            </article>`;
        }).join('')}
    </div>
</body>
</html>`);
    }

    /* ── Init / Load ── */
    async function initBilling() {
        const today = new Date().toISOString().split('T')[0];
        // Default: By Date mode, today
        if (blDateFilterEl && !blDateFilterEl.value) blDateFilterEl.value = today;
        _blSetMode('date');
        await loadBillingData();
    }

    async function loadBillingData() {
        const barName = activeBar.barName;
        const financialYear = activeBar.financialYear || '';
        if (!barName) {
            clearBillingPreview();
            return;
        }

        try {
            const res = await window.electronAPI.getDailySales({ barName, financialYear });
            const allSales = res.success ? (res.sales || []) : [];
            blAllSales = getBillingFilteredSales(allSales);
        } catch (err) {
            console.error('Billing load error:', err);
            blAllSales = [];
        }

        blCurrentPage = 0;
        paginateBills();
        renderBillingPreview();
        updateBillingStats();
    }

    /* ── Event Wiring ── */
    // Mode switcher helper
    function _blSetMode(mode) {
        const filterInputs = document.getElementById('blFilterInputs');
        if (!filterInputs) return;
        // Update active mode button
        [blModeDateBtn, blModeMonthBtn, blModeAllBtn].forEach(b => b?.classList.remove('active'));
        if (mode === 'date') {
            blModeDateBtn?.classList.add('active');
            if (blDateFilterEl) { blDateFilterEl.classList.remove('hidden'); filterInputs.style.display = ''; }
            if (blMonthFilterEl) blMonthFilterEl.classList.add('hidden');
            if (blTodayBtn) blTodayBtn.style.display = '';
        } else if (mode === 'month') {
            blModeMonthBtn?.classList.add('active');
            if (blMonthFilterEl) { blMonthFilterEl.classList.remove('hidden'); filterInputs.style.display = ''; }
            if (blDateFilterEl) blDateFilterEl.classList.add('hidden');
            if (blTodayBtn) blTodayBtn.style.display = 'none';
        } else { // all
            blModeAllBtn?.classList.add('active');
            filterInputs.style.display = 'none';
            if (blTodayBtn) blTodayBtn.style.display = 'none';
        }
    }

    if (blModeDateBtn) {
        blModeDateBtn.addEventListener('click', () => {
            const today = new Date().toISOString().split('T')[0];
            if (!blDateFilterEl?.value) blDateFilterEl.value = today;
            _blSetMode('date');
            loadBillingData();
        });
    }
    if (blModeMonthBtn) {
        blModeMonthBtn.addEventListener('click', () => {
            const today = new Date().toISOString().split('T')[0];
            if (!blMonthFilterEl?.value) blMonthFilterEl.value = today.substring(0, 7);
            _blSetMode('month');
            loadBillingData();
        });
    }
    if (blModeAllBtn) {
        blModeAllBtn.addEventListener('click', () => {
            _blSetMode('all');
            loadBillingData();
        });
    }
    if (blDateFilterEl) {
        blDateFilterEl.addEventListener('change', () => loadBillingData());
    }
    if (blMonthFilterEl) {
        blMonthFilterEl.addEventListener('change', () => loadBillingData());
    }
    if (blTodayBtn) {
        blTodayBtn.addEventListener('click', () => {
            const today = new Date().toISOString().split('T')[0];
            if (blDateFilterEl) blDateFilterEl.value = today;
            _blSetMode('date');
            loadBillingData();
        });
    }
    if (blLoadBtn) blLoadBtn.addEventListener('click', () => loadBillingData());
    if (blPagePrevBtn) {
        blPagePrevBtn.addEventListener('click', () => {
            if (blCurrentPage > 0) {
                blCurrentPage -= 1;
                renderBillingPreview();
            }
        });
    }
    if (blPageNextBtn) {
        blPageNextBtn.addEventListener('click', () => {
            if (blCurrentPage < blPagedSales.length - 1) {
                blCurrentPage += 1;
                renderBillingPreview();
            }
        });
    }
    if (blPrintBtn) blPrintBtn.addEventListener('click', printCurrentBillingPage);
    document.getElementById('blPreviewBtn')?.addEventListener('click', () => { _previewMode = true; printCurrentBillingPage(); });
    if (blVatToggle) {
        blVatToggle.addEventListener('change', () => renderBillingPreview());
    }

    // ═══════════════════════════════════════════════════════
    //  SALE SUMMARY — Date-wise grouped, Edit/Delete/Export
    // ═══════════════════════════════════════════════════════

    let ssAllSales = [];
    let ssEditingSale = null;
    let ssEditItems = [];
    const SS_VAT_RATE = 10;

    const ssDateFrom     = document.getElementById('ssDateFrom');
    const ssDateTo       = document.getElementById('ssDateTo');
    const ssLoadBtn      = document.getElementById('ssLoadBtn');
    const ssSalesList    = document.getElementById('ssSalesList');
    const ssStatDays     = document.getElementById('ssStatDays');
    const ssStatBills    = document.getElementById('ssStatBills');
    const ssStatBottles  = document.getElementById('ssStatBottles');
    const ssStatAmount   = document.getElementById('ssStatAmount');

    // Edit modal elements
    const ssEditOverlay   = document.getElementById('ssEditOverlay');
    const ssModalTitle    = document.getElementById('ssModalTitle');
    const ssModalClose    = document.getElementById('ssModalClose');
    const ssEditBillNo    = document.getElementById('ssEditBillNo');
    const ssEditCust      = document.getElementById('ssEditCust');
    const ssEditDateEl    = document.getElementById('ssEditDate');
    const ssEditTableBody = document.getElementById('ssEditTableBody');
    const ssEditTotalQty  = document.getElementById('ssEditTotalQty');
    const ssEditSubTotal  = document.getElementById('ssEditSubTotal');
    const ssEditVat       = document.getElementById('ssEditVat');
    const ssEditGrand     = document.getElementById('ssEditGrand');
    const ssEditSaveBtn   = document.getElementById('ssEditSaveBtn');
    const ssEditCancelBtn = document.getElementById('ssEditCancelBtn');

    function escSS(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }

    function initSaleSummary() {
        const today = new Date().toISOString().split('T')[0];
        // Default: whole financial year (April 1 → today)
        const now = new Date();
        const fyYear = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
        const fyStart = fyYear + '-04-01';
        if (ssDateFrom && !ssDateFrom.value) ssDateFrom.value = fyStart;
        if (ssDateTo && !ssDateTo.value) ssDateTo.value = today;
        loadSaleSummary();
    }

    async function loadSaleSummary() {
        const barName = activeBar.barName;
        const financialYear = activeBar.financialYear;

        try {
            const res = await window.electronAPI.getDailySales({ barName, financialYear });
            ssAllSales = res.success ? (res.sales || []) : [];
        } catch (err) {
            console.error('Sale summary load error:', err);
            ssAllSales = [];
        }

        renderSaleSummary();
    }

    function renderSaleSummary() {
        if (!ssSalesList) return;
        ssSalesList.innerHTML = '';

        const fromDate = ssDateFrom?.value || '';
        const toDate = ssDateTo?.value || '';

        // Filter by date range
        let filtered = ssAllSales;
        if (fromDate) filtered = filtered.filter(s => (s.billDate || '') >= fromDate);
        if (toDate) filtered = filtered.filter(s => (s.billDate || '') <= toDate);

        if (filtered.length === 0) {
            ssSalesList.innerHTML = `<div class="ss-empty">
                <div class="ss-empty-icon">📊</div>
                <p>No sales found for the selected date range.</p>
            </div>`;
            updateSsStats([], 0);
            return;
        }

        // Group by date
        const grouped = {};
        for (const sale of filtered) {
            const dt = sale.billDate || 'Unknown';
            if (!grouped[dt]) grouped[dt] = [];
            grouped[dt].push(sale);
        }

        // Sort dates descending (most recent first)
        const sortedDates = Object.keys(grouped).sort((a, b) => b.localeCompare(a));

        for (const dt of sortedDates) {
            const bills = grouped[dt];
            const totalQty = bills.reduce((s, b) => s + (b.totalQty || 0), 0);
            const totalAmt = bills.reduce((s, b) => s + (b.grandTotal || 0), 0);

            // Format date nicely
            let dateDisplay = dt;
            let dayName = '';
            try {
                const d = new Date(dt + 'T00:00:00');
                dateDisplay = d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
                dayName = d.toLocaleDateString('en-IN', { weekday: 'long' });
            } catch (e) { /* keep raw */ }

            const group = document.createElement('div');
            group.className = 'ss-date-group collapsed';
            group.innerHTML = `
                <div class="ss-date-header">
                    <div class="ss-date-header-left">
                        <svg class="ss-date-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
                        <span class="ss-date-text">${escSS(dateDisplay)}</span>
                        <span class="ss-date-day">${escSS(dayName)}</span>
                    </div>
                    <div class="ss-date-header-right">
                        <div class="ss-date-stats">
                            <span>Bills: <strong>${bills.length}</strong></span>
                            <span>Btl: <strong>${totalQty}</strong></span>
                            <span>₹ <strong>${totalAmt.toFixed(0)}</strong></span>
                        </div>
                        <button class="ss-btn-export" data-date="${escSS(dt)}" title="Export this day's sales as XLSX for SCM portal">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                            Export XLSX
                        </button>
                    </div>
                </div>
                <div class="ss-date-body">
                    <table class="ss-bill-table">
                        <thead>
                            <tr>
                                <th class="ss-td-sno">#</th>
                                <th class="ss-td-billno">Bill No</th>
                                <th class="ss-td-customer">Customer</th>
                                <th class="ss-td-items">Items</th>
                                <th class="ss-td-qty">Btl</th>
                                <th class="ss-td-amount">Amount</th>
                                <th class="ss-td-actions">Actions</th>
                            </tr>
                        </thead>
                        <tbody></tbody>
                        <tfoot>
                            <tr>
                                <td colspan="4" style="text-align:right">Total</td>
                                <td class="ss-td-qty">${totalQty}</td>
                                <td class="ss-td-amount">₹ ${totalAmt.toFixed(2)}</td>
                                <td></td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            `;

            // Populate bill rows
            const tbody = group.querySelector('tbody');
            bills.forEach((bill, idx) => {
                const itemSummary = (bill.items || []).map(i => `${i.brandName} ${i.size} x${i.qty}`).join(', ');
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="ss-td-sno">${idx + 1}</td>
                    <td class="ss-td-billno">${escSS(bill.billNo || '')}</td>
                    <td class="ss-td-customer" title="${escSS(bill.customerName || '')} — ${escSS(bill.customerLicNo || '')}">${escSS(bill.customerName || '—')}</td>
                    <td class="ss-td-items" title="${escSS(itemSummary)}">${escSS(itemSummary || '—')}</td>
                    <td class="ss-td-qty">${bill.totalQty || 0}</td>
                    <td class="ss-td-amount">₹ ${(bill.grandTotal || 0).toFixed(2)}</td>
                    <td class="ss-td-actions">
                        <button class="ss-act-btn ss-act-edit" title="Edit bill" data-id="${escSS(bill.id)}">✏️</button>
                        <button class="ss-act-btn ss-act-del" title="Delete bill" data-id="${escSS(bill.id)}">🗑️</button>
                    </td>
                `;
                tbody.appendChild(tr);
            });

            // Wire collapse toggle
            const header = group.querySelector('.ss-date-header');
            header.addEventListener('click', (e) => {
                // Don't toggle if clicking export button
                if (e.target.closest('.ss-btn-export')) return;
                group.classList.toggle('collapsed');
            });

            // Wire export
            const exportBtn = group.querySelector('.ss-btn-export');
            if (exportBtn) {
                exportBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    exportDaySalesXlsx(dt, bills);
                });
            }

            // Wire edit / delete buttons
            group.querySelectorAll('.ss-act-edit').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const bill = filtered.find(b => b.id === btn.dataset.id);
                    if (bill) openSsEditModal(bill);
                });
            });
            group.querySelectorAll('.ss-act-del').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const bill = filtered.find(b => b.id === btn.dataset.id);
                    if (bill) deleteSsBill(bill);
                });
            });

            ssSalesList.appendChild(group);
        }

        updateSsStats(sortedDates, filtered.length);
    }

    function updateSsStats(dates, billCount) {
        const fromDate = ssDateFrom?.value || '';
        const toDate = ssDateTo?.value || '';
        let filtered = ssAllSales;
        if (fromDate) filtered = filtered.filter(s => (s.billDate || '') >= fromDate);
        if (toDate) filtered = filtered.filter(s => (s.billDate || '') <= toDate);

        const totalBottles = filtered.reduce((s, b) => s + (b.totalQty || 0), 0);
        const totalAmount = filtered.reduce((s, b) => s + (b.grandTotal || 0), 0);

        if (ssStatDays) ssStatDays.textContent = dates.length;
        if (ssStatBills) ssStatBills.textContent = billCount;
        if (ssStatBottles) ssStatBottles.textContent = totalBottles;
        if (ssStatAmount) ssStatAmount.textContent = `₹ ${totalAmount.toFixed(0)}`;
    }

    /* ── Delete Sale ── */
    async function deleteSsBill(bill) {
        if (!confirm(`Delete bill ${bill.billNo}?\nCustomer: ${bill.customerName}\nDate: ${bill.billDate}\n\nThis cannot be undone.`)) return;

        const barName = activeBar.barName;
        const financialYear = activeBar.financialYear;

        try {
            const res = await window.electronAPI.deleteDailySale({ barName, financialYear, saleId: bill.id });
            if (res.success) {
                ssAllSales = ssAllSales.filter(b => b.id !== bill.id);
                renderSaleSummary();
            } else {
                alert('Delete failed: ' + (res.error || 'Unknown error'));
            }
        } catch (err) {
            alert('Delete error: ' + err.message);
        }
    }

    /* ── Edit Modal ── */
    function openSsEditModal(bill) {
        ssEditingSale = JSON.parse(JSON.stringify(bill));
        ssEditItems = ssEditingSale.items || [];

        if (ssModalTitle) ssModalTitle.textContent = 'Edit Bill';
        if (ssEditBillNo) ssEditBillNo.textContent = bill.billNo;
        if (ssEditCust) ssEditCust.textContent = `${bill.customerName || '—'} — ${bill.customerLicNo || ''}`;
        if (ssEditDateEl) ssEditDateEl.textContent = bill.billDate || '';

        renderSsEditTable();
        if (ssEditOverlay) ssEditOverlay.classList.remove('hidden');
    }

    function closeSsEditModal() {
        if (ssEditOverlay) ssEditOverlay.classList.add('hidden');
        ssEditingSale = null;
        ssEditItems = [];
    }

    function renderSsEditTable() {
        if (!ssEditTableBody) return;
        ssEditTableBody.innerHTML = '';

        ssEditItems.forEach((item, idx) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${idx + 1}</td>
                <td>${escSS(item.brandName)}</td>
                <td>${escSS(item.size)}</td>
                <td><input type="number" class="ss-eq" data-idx="${idx}" value="${item.qty}" min="1"></td>
                <td>₹${(item.mrp || 0).toFixed(2)}</td>
                <td class="ss-ea" data-idx="${idx}">₹${(item.amount || 0).toFixed(2)}</td>
                <td><button class="ss-edit-del" data-idx="${idx}" title="Remove item">✕</button></td>
            `;
            ssEditTableBody.appendChild(tr);
        });

        // Wire qty change
        ssEditTableBody.querySelectorAll('.ss-eq').forEach(input => {
            input.addEventListener('input', () => {
                const i = parseInt(input.dataset.idx);
                const newQty = Math.max(1, parseInt(input.value) || 1);
                ssEditItems[i].qty = newQty;
                ssEditItems[i].amount = newQty * ssEditItems[i].mrp;
                ssEditItems[i].vatAmt = Math.round(ssEditItems[i].amount * SS_VAT_RATE / 100 * 100) / 100;
                ssEditItems[i].lineTotal = ssEditItems[i].amount + ssEditItems[i].vatAmt;
                const amtEl = ssEditTableBody.querySelector(`.ss-ea[data-idx="${i}"]`);
                if (amtEl) amtEl.textContent = `₹${ssEditItems[i].amount.toFixed(2)}`;
                updateSsEditTotals();
            });
        });

        // Wire delete
        ssEditTableBody.querySelectorAll('.ss-edit-del').forEach(btn => {
            btn.addEventListener('click', () => {
                ssEditItems.splice(parseInt(btn.dataset.idx), 1);
                renderSsEditTable();
            });
        });

        updateSsEditTotals();
    }

    function updateSsEditTotals() {
        const totalQty = ssEditItems.reduce((s, i) => s + i.qty, 0);
        const subTotal = ssEditItems.reduce((s, i) => s + (i.amount || 0), 0);
        const totalVat = ssEditItems.reduce((s, i) => s + (i.vatAmt || 0), 0);
        const grandTotal = subTotal + totalVat;

        if (ssEditTotalQty) ssEditTotalQty.textContent = totalQty;
        if (ssEditSubTotal) ssEditSubTotal.textContent = `₹${subTotal.toFixed(2)}`;
        if (ssEditVat) ssEditVat.textContent = `₹${totalVat.toFixed(2)}`;
        if (ssEditGrand) ssEditGrand.textContent = `₹${grandTotal.toFixed(2)}`;
    }

    /* ── Save Edit ── */
    async function saveSsEditedBill() {
        if (!ssEditingSale) return;
        if (ssEditItems.length === 0) {
            if (confirm('All items removed. Delete this bill instead?')) {
                await deleteSsBill(ssEditingSale);
                closeSsEditModal();
                return;
            }
            return;
        }

        ssEditingSale.items = ssEditItems;
        ssEditingSale.totalQty = ssEditItems.reduce((s, i) => s + i.qty, 0);
        ssEditingSale.subTotal = ssEditItems.reduce((s, i) => s + (i.amount || 0), 0);
        ssEditingSale.totalVat = ssEditItems.reduce((s, i) => s + (i.vatAmt || 0), 0);
        ssEditingSale.grandTotal = ssEditingSale.subTotal + ssEditingSale.totalVat;

        const barName = activeBar.barName;
        const financialYear = activeBar.financialYear;

        try {
            const res = await window.electronAPI.saveDailySale({ barName, financialYear, sale: ssEditingSale });
            if (res.success) {
                const idx = ssAllSales.findIndex(b => b.id === ssEditingSale.id);
                if (idx >= 0) ssAllSales[idx] = ssEditingSale;
                renderSaleSummary();
                closeSsEditModal();
            } else {
                alert('Save failed: ' + (res.error || 'Unknown error'));
            }
        } catch (err) {
            alert('Save error: ' + err.message);
        }
    }

    /* ── Export day's sales as XLSX for SCM portal ── */
    async function exportDaySalesXlsx(date, bills) {
        // Build rows in exact SCM portal format:
        // Sale Date | Local Item Code | Brand Name | Size | Quantity(Case) | Quantity(Loose Bottle)
        // One row per unique brand+size (code), aggregated across all bills for that day

        if (!bills || bills.length === 0) {
            alert('No bills to export for this date.');
            return;
        }

        // Format date as MM/DD/YYYY for SCM portal
        let saleDate = date || '';
        if (saleDate && saleDate.includes('-')) {
            const parts = saleDate.split('-'); // YYYY-MM-DD
            saleDate = `${parts[1]}/${parts[2]}/${parts[0]}`; // MM/DD/YYYY
        }

        // Aggregate qty per brand+size (by code)
        const brandAgg = {};
        for (const bill of bills) {
            for (const item of (bill.items || [])) {
                const key = item.code || item.productId || '';
                if (!brandAgg[key]) {
                    brandAgg[key] = {
                        code: item.code || '',
                        brandName: item.brandName || '',
                        size: item.size || '',
                        totalQty: 0,
                    };
                }
                brandAgg[key].totalQty += (item.qty || 0);
            }
        }

        // Sort by brand name then size
        const sorted = Object.values(brandAgg).sort((a, b) => {
            const cmp = a.brandName.localeCompare(b.brandName);
            return cmp !== 0 ? cmp : a.size.localeCompare(b.size);
        });

        // Build SCM-format rows
        const rows = sorted.map(item => ({
            'Sale Date': saleDate,
            'Local Item Code': item.code,
            'Brand Name': item.brandName,
            'Size': item.size,
            'Quantity(Case)': '',
            'Quantity(Loose Bottle)': item.totalQty,
        }));

        try {
            const res = await window.electronAPI.exportSaleSummary({
                rows: rows,
                barName: activeBar.barName,
                date: date,
            });
            if (res.success) {
                updateStatus(`✅ Exported to ${res.filePath}`);
            } else if (res.error !== 'Cancelled') {
                alert('Export failed: ' + res.error);
            }
        } catch (err) {
            alert('Export error: ' + err.message);
        }
    }

    /* ── Event Wiring ── */
    if (ssLoadBtn) ssLoadBtn.addEventListener('click', () => renderSaleSummary());
    if (ssDateFrom) ssDateFrom.addEventListener('change', () => renderSaleSummary());
    if (ssDateTo) ssDateTo.addEventListener('change', () => renderSaleSummary());
    if (ssModalClose) ssModalClose.addEventListener('click', closeSsEditModal);
    if (ssEditCancelBtn) ssEditCancelBtn.addEventListener('click', closeSsEditModal);
    if (ssEditSaveBtn) ssEditSaveBtn.addEventListener('click', saveSsEditedBill);
    if (ssEditOverlay) {
        ssEditOverlay.addEventListener('click', (e) => {
            if (e.target === ssEditOverlay) closeSsEditModal();
        });
    }

    // ═══════════════════════════════════════════════════════
    //  CUSTOMER MANAGER
    // ═══════════════════════════════════════════════════════

    let allCustomers   = [];
    let filteredCusts  = [];
    let custEditingId  = null;
    let custFocusedIdx = -1;

    const custListBody   = document.getElementById('custListBody');
    const custEmptyState = document.getElementById('custEmptyState');
    const custFormEl     = document.getElementById('custForm');
    const custPlaceholder= document.getElementById('custPlaceholder');
    const custSearchEl   = document.getElementById('custSearch');
    const custCountEl    = document.getElementById('custCount');
    const custNameEl     = document.getElementById('custName');
    const custLicNoEl    = document.getElementById('custLicNo');
    const custSaveBtn    = document.getElementById('custSaveBtn');
    const custCancelBtn  = document.getElementById('custCancelBtn');
    const custDeleteBtn  = document.getElementById('custDeleteBtn');
    const custNewBtn     = document.getElementById('custNewBtn');

    async function loadCustomers() {
        try {
            const result = await window.electronAPI.getCustomers({
                barName: activeBar.barName, financialYear: activeBar.financialYear
            });
            if (result.success) {
                allCustomers = result.customers || [];
            }
        } catch (err) {
            console.error('loadCustomers error:', err);
        }
        filteredCusts = [...allCustomers];
        renderCustList();
    }

    function renderCustList() {
        if (!custListBody) return;
        // Remove existing rows (not empty state)
        custListBody.querySelectorAll('.cust-row').forEach(r => r.remove());

        if (custCountEl) custCountEl.textContent = `${allCustomers.length} customer${allCustomers.length !== 1 ? 's' : ''}`;

        if (filteredCusts.length === 0) {
            if (custEmptyState) custEmptyState.style.display = '';
            return;
        }
        if (custEmptyState) custEmptyState.style.display = 'none';

        filteredCusts.forEach((cust, idx) => {
            const row = document.createElement('div');
            row.className = 'cust-row';
            row.tabIndex = 0;
            row.innerHTML = `
                <span class="cust-row-name">${esc(cust.name)}</span>
                <span class="cust-row-lic">${esc(cust.licNo || '—')}</span>
            `;

            row.addEventListener('click', () => {
                setCustFocus(idx);
                openCustomerForEdit(cust);
            });

            row.addEventListener('keydown', (e) => {
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    const next = Math.min(idx + 1, filteredCusts.length - 1);
                    focusCustRow(next);
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    if (idx === 0) { if (custSearchEl) custSearchEl.focus(); return; }
                    focusCustRow(idx - 1);
                } else if (e.key === 'Enter') {
                    e.preventDefault();
                    openCustomerForEdit(filteredCusts[idx]);
                } else if (e.key === 'Delete' && e.ctrlKey) {
                    e.preventDefault();
                    deleteCustomerConfirm(filteredCusts[idx]);
                }
            });

            if (custEmptyState) custListBody.insertBefore(row, custEmptyState);
            else custListBody.appendChild(row);
        });
    }

    function setCustFocus(idx) {
        custFocusedIdx = idx;
        custListBody.querySelectorAll('.cust-row').forEach((r, i) => {
            r.classList.toggle('focused', i === idx);
        });
    }

    function focusCustRow(idx) {
        const rows = custListBody.querySelectorAll('.cust-row');
        if (rows[idx]) {
            setCustFocus(idx);
            rows[idx].focus();
            rows[idx].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        }
    }

    /* ── Open form in 'new' mode ── */
    function openNewCustomer() {
        custEditingId = null;
        clearCustForm();
        if (custPlaceholder) custPlaceholder.style.display = 'none';
        if (custFormEl) custFormEl.classList.remove('hidden');
        if (custDeleteBtn) custDeleteBtn.classList.add('hidden');
        if (custNameEl) custNameEl.focus();
    }

    /* ── Open for edit ── */
    function openCustomerForEdit(cust) {
        custEditingId = cust.id;
        if (custPlaceholder) custPlaceholder.style.display = 'none';
        if (custFormEl) custFormEl.classList.remove('hidden');
        if (custDeleteBtn) custDeleteBtn.classList.remove('hidden');

        if (custNameEl) custNameEl.value = cust.name || '';
        if (custLicNoEl) custLicNoEl.value = cust.licNo || '';

        // Highlight active row
        custListBody.querySelectorAll('.cust-row').forEach((r, i) => {
            r.classList.toggle('active', filteredCusts[i]?.id === cust.id);
        });

        if (custNameEl) custNameEl.focus();
    }

    /* ── Clear form ── */
    function clearCustForm() {
        if (custNameEl) custNameEl.value = '';
        if (custLicNoEl) custLicNoEl.value = '';
    }

    /* ── Cancel ── */
    function cancelCustForm() {
        custEditingId = null;
        clearCustForm();
        if (custFormEl) custFormEl.classList.add('hidden');
        if (custPlaceholder) custPlaceholder.style.display = '';
        if (custDeleteBtn) custDeleteBtn.classList.add('hidden');
        custListBody.querySelectorAll('.cust-row').forEach(r => r.classList.remove('active'));
    }

    /* ── Save ── */
    async function saveCustomer() {
        const name  = (custNameEl?.value || '').trim();
        const licNo = (custLicNoEl?.value || '').trim();

        if (!name) {
            custNameEl?.focus();
            return;
        }
        if (!licNo) {
            custLicNoEl?.focus();
            return;
        }

        const customer = {
            id: custEditingId || ('cust_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7)),
            name,
            licNo
        };

        try {
            const result = await window.electronAPI.saveCustomer({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear,
                customer
            });
            if (result.success) {
                await loadCustomers();
                cancelCustForm();
            } else {
                console.error('Save customer failed:', result.error);
            }
        } catch (err) {
            console.error('Save customer error:', err);
        }
    }

    /* ── Delete ── */
    async function deleteCustomerConfirm(cust) {
        const target = cust || (custEditingId && allCustomers.find(c => c.id === custEditingId));
        if (!target) return;
        if (!confirm(`Delete customer "${target.name}"?`)) return;

        try {
            const result = await window.electronAPI.deleteCustomer({
                barName: activeBar.barName,
                financialYear: activeBar.financialYear,
                customerId: target.id
            });
            if (result.success) {
                await loadCustomers();
                cancelCustForm();
            }
        } catch (err) {
            console.error('Delete customer error:', err);
        }
    }

    /* ── Wire buttons ── */
    if (custNewBtn) custNewBtn.addEventListener('click', openNewCustomer);
    if (custSaveBtn) custSaveBtn.addEventListener('click', saveCustomer);
    if (custCancelBtn) custCancelBtn.addEventListener('click', cancelCustForm);
    if (custDeleteBtn) custDeleteBtn.addEventListener('click', () => deleteCustomerConfirm());

    /* ── Search filter ── */
    if (custSearchEl) {
        custSearchEl.addEventListener('input', () => {
            const q = custSearchEl.value.toLowerCase().trim();
            filteredCusts = allCustomers.filter(c =>
                (c.name || '').toLowerCase().includes(q) ||
                (c.licNo || '').toLowerCase().includes(q)
            );
            renderCustList();
        });

        custSearchEl.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                focusCustRow(0);
            } else if (e.key === 'F2') {
                e.preventDefault();
                openNewCustomer();
            }
        });
    }

    /* ── Keyboard flow: Enter moves to next field ── */
    const custFlowFields = document.querySelectorAll('.cust-flow');
    custFlowFields.forEach((field, idx) => {
        field.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                if (idx < custFlowFields.length - 1) {
                    custFlowFields[idx + 1].focus();
                } else {
                    saveCustomer();
                }
            }
        });
    });

    /* ── Customer section keyboard shortcuts ── */
    document.getElementById('sub-customer-manager')?.addEventListener('keydown', (e) => {
        // F2 = New
        if (e.key === 'F2') {
            e.preventDefault();
            openNewCustomer();
            return;
        }

        // Ctrl+S = Save
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
            if (custFormEl && !custFormEl.classList.contains('hidden')) {
                e.preventDefault();
                saveCustomer();
                return;
            }
        }

        // Ctrl+Delete = Delete
        if (e.ctrlKey && e.key === 'Delete') {
            if (custEditingId) {
                e.preventDefault();
                deleteCustomerConfirm();
                return;
            }
        }

        // Esc = Cancel form or clear search
        if (e.key === 'Escape') {
            if (custFormEl && !custFormEl.classList.contains('hidden')) {
                e.preventDefault();
                cancelCustForm();
            } else if (custSearchEl && custSearchEl.value) {
                e.preventDefault();
                custSearchEl.value = '';
                custSearchEl.dispatchEvent(new Event('input'));
                custSearchEl.focus();
            }
        }
    });

    /* ── Load customers on init ── */
    loadCustomers();

    /* ── Helper: html-escape ── */
    function esc(str) {
        return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    /* ── Load suppliers on init ── */
    loadSuppliers();

    // ═══════════════════════════════════════════════════════
    //  PASSWORD MANAGER
    // ═══════════════════════════════════════════════════════
    {
        const PWMGR_KEY = 'spliqour-pwmgr';
        const pwmgrList    = document.getElementById('pwmgrList');
        const pwmgrSearch  = document.getElementById('pwmgrSearch');
        const pwmgrAddBtn  = document.getElementById('pwmgrAddBtn');
        const pwmgrEmptyCta  = document.getElementById('pwmgrEmptyCta');
        const pwmgrFormTitle  = document.getElementById('pwmgrFormTitle');
        const pwmgrFormBody   = document.getElementById('pwmgrFormBody');
        const pwmgrEmptyForm  = document.getElementById('pwmgrEmptyForm');
        const pwmgrDetailAvatar = document.getElementById('pwmgrDetailAvatar');
        const pwmgrDetailSub    = document.getElementById('pwmgrDetailSub');
        const pwmgrPortal   = document.getElementById('pwmgrPortal');
        const pwmgrBar      = document.getElementById('pwmgrBar');
        const pwmgrUsername = document.getElementById('pwmgrUsername');
        const pwmgrPassword = document.getElementById('pwmgrPassword');
        const pwmgrNotes    = document.getElementById('pwmgrNotes');
        const pwmgrSaveBtn  = document.getElementById('pwmgrSaveBtn');
        const pwmgrClearBtn = document.getElementById('pwmgrClearBtn');
        const pwmgrDeleteBtn = document.getElementById('pwmgrDeleteBtn');
        const pwmgrTogglePass = document.getElementById('pwmgrTogglePass');
        const pwmgrCopyUser  = document.getElementById('pwmgrCopyUser');
        const pwmgrCopyPass  = document.getElementById('pwmgrCopyPass');

        if (!pwmgrList) { /* panel not in DOM yet */ }
        else {

        let pwmgrCreds = [];
        let pwmgrEditId = null; // null = new, string = editing existing

        /* ── Persist ── */
        function pwmgrLoad() {
            try { pwmgrCreds = JSON.parse(localStorage.getItem(PWMGR_KEY) || '[]'); } catch(e) { pwmgrCreds = []; }
        }
        function pwmgrSave() {
            localStorage.setItem(PWMGR_KEY, JSON.stringify(pwmgrCreds));
        }

        /* ── Render list ── */
        function pwmgrRenderList() {
            const q = (pwmgrSearch ? pwmgrSearch.value : '').trim().toLowerCase();
            // Only show credentials saved for the active bar (or ones with no bar set)
            const activeBarName = (activeBar.barName || '').trim().toLowerCase();
            const barFiltered = pwmgrCreds.filter(c =>
                (c.barName || '').trim().toLowerCase() === activeBarName
            );
            const filtered = q
                ? barFiltered.filter(c =>
                    (c.portalName || '').toLowerCase().includes(q) ||
                    (c.username   || '').toLowerCase().includes(q) ||
                    (c.notes      || '').toLowerCase().includes(q))
                : barFiltered;

            if (!filtered.length) {
                pwmgrList.innerHTML = `<div class="pwmgr-list-empty">${q ? 'No matching credentials.' : 'No credentials saved yet.<br>Click <strong>New</strong> to add one.'}</div>`;
                return;
            }
            pwmgrList.innerHTML = filtered.map(c => {
                const letter = (c.portalName || '?')[0].toUpperCase();
                return `
                <div class="pwmgr-card${c.id === pwmgrEditId ? ' active' : ''}" data-id="${c.id}">
                    <div class="pwmgr-card-avatar">${letter}</div>
                    <div class="pwmgr-card-info">
                        <div class="pwmgr-card-portal">${esc(c.portalName)}</div>
                        ${c.barName ? `<div class="pwmgr-card-bar">${esc(c.barName)}</div>` : ''}
                        <div class="pwmgr-card-user">${esc(c.username)}</div>
                    </div>
                    <svg class="pwmgr-card-arrow" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 18l6-6-6-6"/></svg>
                </div>
                `;
            }).join('');
            pwmgrList.querySelectorAll('.pwmgr-card').forEach(card => {
                card.addEventListener('click', () => pwmgrOpenEdit(card.dataset.id));
            });
        }

        /* ── Show empty / form states ── */
        function pwmgrShowForm(show) {
            if (pwmgrFormBody)  pwmgrFormBody.style.display  = show ? '' : 'none';
            if (pwmgrEmptyForm) pwmgrEmptyForm.classList.toggle('hidden', show);
        }

        /* ── Clear / reset form ── */
        function pwmgrClearForm() {
            pwmgrEditId = null;
            if (pwmgrPortal)   pwmgrPortal.value   = '';
            if (pwmgrBar)      pwmgrBar.value       = activeBar.barName || '';
            if (pwmgrUsername) pwmgrUsername.value  = '';
            if (pwmgrPassword) pwmgrPassword.value  = '';
            if (pwmgrNotes)    pwmgrNotes.value     = '';
            if (pwmgrFormTitle) pwmgrFormTitle.textContent = 'New Credential';
            if (pwmgrDetailAvatar) pwmgrDetailAvatar.textContent = '+';
            if (pwmgrDetailSub) pwmgrDetailSub.textContent = activeBar.barName || 'Active Bar';
            if (pwmgrDetailSub) pwmgrDetailSub.textContent = 'Fill in the details below';
            if (pwmgrDeleteBtn) pwmgrDeleteBtn.classList.add('hidden');
            // reset password visibility
            if (pwmgrPassword) pwmgrPassword.type = 'password';
            if (pwmgrTogglePass) {
                pwmgrTogglePass.querySelector('.eye-show')?.classList.remove('hidden');
                pwmgrTogglePass.querySelector('.eye-hide')?.classList.add('hidden');
            }
            pwmgrShowForm(true);
            pwmgrRenderList();
            if (pwmgrPortal) setTimeout(() => pwmgrPortal.focus(), 60);
        }

        /* ── Open existing for edit ── */
        function pwmgrOpenEdit(id) {
            const cred = pwmgrCreds.find(c => c.id === id);
            if (!cred) return;
            pwmgrEditId = id;
            if (pwmgrPortal)   pwmgrPortal.value   = cred.portalName || '';
            if (pwmgrBar)      pwmgrBar.value       = activeBar.barName || '';
            if (pwmgrUsername) pwmgrUsername.value  = cred.username   || '';
            if (pwmgrPassword) pwmgrPassword.value  = cred.password   || '';
            if (pwmgrNotes)    pwmgrNotes.value     = cred.notes      || '';
            if (pwmgrFormTitle) pwmgrFormTitle.textContent = cred.portalName || 'Edit Credential';
            if (pwmgrDetailAvatar) pwmgrDetailAvatar.textContent = (cred.portalName || '?')[0].toUpperCase();
            if (pwmgrDetailSub) pwmgrDetailSub.textContent = activeBar.barName || 'Active Bar';
            if (pwmgrDeleteBtn) pwmgrDeleteBtn.classList.remove('hidden');
            // reset password visibility
            if (pwmgrPassword) pwmgrPassword.type = 'password';
            if (pwmgrTogglePass) {
                pwmgrTogglePass.querySelector('.eye-show')?.classList.remove('hidden');
                pwmgrTogglePass.querySelector('.eye-hide')?.classList.add('hidden');
            }
            pwmgrShowForm(true);
            pwmgrRenderList();
            if (pwmgrUsername) setTimeout(() => pwmgrUsername.focus(), 60);
        }

        /* ── Save ── */
        function pwmgrDoSave() {
            const portal = (pwmgrPortal?.value || '').trim();
            const user   = (pwmgrUsername?.value || '').trim();
            const pass   = (pwmgrPassword?.value || '');
            if (!portal) { pwmgrPortal?.focus(); return; }
            if (!user)   { pwmgrUsername?.focus(); return; }
            if (!pass)   { pwmgrPassword?.focus(); return; }

            if (pwmgrEditId) {
                const idx = pwmgrCreds.findIndex(c => c.id === pwmgrEditId);
                if (idx >= 0) {
                    pwmgrCreds[idx] = { ...pwmgrCreds[idx], portalName: portal, barName: pwmgrBar?.value.trim() || '', username: user, password: pass, notes: pwmgrNotes?.value.trim() || '', updatedAt: new Date().toISOString() };
                }
            } else {
                pwmgrCreds.push({ id: 'pw_' + Date.now(), portalName: portal, barName: pwmgrBar?.value.trim() || '', username: user, password: pass, notes: pwmgrNotes?.value.trim() || '', createdAt: new Date().toISOString() });
                pwmgrEditId = pwmgrCreds[pwmgrCreds.length - 1].id;
            }
            pwmgrSave();
            const savedPortal = (pwmgrPortal?.value || '').trim();
            const savedBar = (pwmgrBar?.value || '').trim();
            if (pwmgrFormTitle) pwmgrFormTitle.textContent = savedPortal || 'Edit Credential';
            if (pwmgrDetailAvatar) pwmgrDetailAvatar.textContent = (savedPortal || '?')[0].toUpperCase();
            if (pwmgrDetailSub) pwmgrDetailSub.textContent = savedBar || 'All Bars';
            if (pwmgrDeleteBtn) pwmgrDeleteBtn.classList.remove('hidden');
            pwmgrRenderList();
        }

        /* ── Delete ── */
        function pwmgrDoDelete() {
            if (!pwmgrEditId) return;
            if (!confirm('Delete this credential? This cannot be undone.')) return;
            pwmgrCreds = pwmgrCreds.filter(c => c.id !== pwmgrEditId);
            pwmgrSave();
            pwmgrEditId = null;
            pwmgrShowForm(false);
            pwmgrRenderList();
        }

        /* ── Copy helper ── */
        function pwmgrCopy(btn, text) {
            if (!text) return;
            navigator.clipboard.writeText(text).then(() => {
                btn.classList.add('copied');
                setTimeout(() => btn.classList.remove('copied'), 1400);
            }).catch(() => {});
        }

        /* ── Wire events ── */
        pwmgrLoad();
        pwmgrShowForm(false);
        pwmgrRenderList();

        pwmgrAddBtn?.addEventListener('click', pwmgrClearForm);
        pwmgrEmptyCta?.addEventListener('click', pwmgrClearForm);
        pwmgrSearch?.addEventListener('input', pwmgrRenderList);
        pwmgrSaveBtn?.addEventListener('click', pwmgrDoSave);
        pwmgrClearBtn?.addEventListener('click', () => {
            pwmgrEditId = null;
            pwmgrShowForm(false);
            pwmgrRenderList();
        });
        pwmgrDeleteBtn?.addEventListener('click', pwmgrDoDelete);

        pwmgrTogglePass?.addEventListener('click', () => {
            const shown = pwmgrPassword.type === 'text';
            pwmgrPassword.type = shown ? 'password' : 'text';
            pwmgrTogglePass.querySelector('.eye-show')?.classList.toggle('hidden', !shown);
            pwmgrTogglePass.querySelector('.eye-hide')?.classList.toggle('hidden', shown);
        });

        pwmgrCopyUser?.addEventListener('click', () => pwmgrCopy(pwmgrCopyUser, pwmgrUsername?.value));
        pwmgrCopyPass?.addEventListener('click', () => pwmgrCopy(pwmgrCopyPass, pwmgrPassword?.value));

        // Ctrl+S to save when inside the password manager panel
        document.getElementById('sub-password-manager')?.addEventListener('keydown', e => {
            if (e.ctrlKey && e.key === 's') { e.preventDefault(); pwmgrDoSave(); }
            if (e.key === 'Escape') { pwmgrEditId = null; pwmgrShowForm(false); pwmgrRenderList(); }
            if (e.key === 'Delete' && e.ctrlKey && pwmgrEditId) { e.preventDefault(); pwmgrDoDelete(); }
            if (e.key === 'F2') { e.preventDefault(); pwmgrClearForm(); }
        });

        } // end else-block
    }
    // ═══════════════════════════════════════════════════════

    // ═══════════════════════════════════════════════════════
    //  ITEM WISE SALE REPORT
    // ═══════════════════════════════════════════════════════
    {
        const iwsDateFrom  = document.getElementById('iwsDateFrom');
        const iwsDateTo    = document.getElementById('iwsDateTo');
        const iwsGoBtn     = document.getElementById('iwsGoBtn');
        const iwsExportBtn  = document.getElementById('iwsExportBtn');
        const iwsPrintBtn   = document.getElementById('iwsPrintBtn');
        const iwsCatFilter  = document.getElementById('iwsCatFilter');
        const iwsTableHead = document.getElementById('iwsTableHead');
        const iwsTableBody = document.getElementById('iwsTableBody');
        const iwsTableFoot = document.getElementById('iwsTableFoot');
        const iwsEmpty     = document.getElementById('iwsEmpty');
        const iwsBarInfo   = document.getElementById('iwsBarInfo');
        const iwsFromLabel    = document.getElementById('iwsFromLabel');
        const iwsToLabel      = document.getElementById('iwsToLabel');
        const iwsTotalItems   = document.getElementById('iwsTotalItems');
        const iwsTotalBottles = document.getElementById('iwsTotalBottles');
        const iwsTotalML      = document.getElementById('iwsTotalML');
        const iwsTotalAmount  = document.getElementById('iwsTotalAmount');

        // Default dates: full financial year
        const iwsFyStartYear = (() => { const n = new Date(); return n.getMonth() >= 3 ? n.getFullYear() : n.getFullYear() - 1; })();
        const iwsFyStart = iwsFyStartYear + '-04-01';
        const iwsFyEnd   = (iwsFyStartYear + 1) + '-03-31';
        if (iwsDateFrom) iwsDateFrom.value = iwsFyStart;
        if (iwsDateTo)   iwsDateTo.value   = iwsFyEnd;

        function iwsGetML(sizeStr) {
            const m = (sizeStr || '').match(/^(\d+)\s*ML/i);
            if (m) return parseInt(m[1]);
            const l = (sizeStr || '').match(/^(\d+(?:\.\d+)?)\s*Ltr/i);
            if (l) return Math.round(parseFloat(l[1]) * 1000);
            return 0;
        }
        function iwsFmtDate(iso) {
            if (!iso) return '—';
            const [y, m, d] = iso.split('-');
            return d + '/' + m + '/' + y;
        }
        function iwsFmtNum(v) { return (v || 0).toLocaleString('en-IN'); }
        function iwsFmtAmt(v) { return (v || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

        let iwsLastRows = [];
        let iwsLastFrom  = '';
        let iwsLastTo    = '';

        function getIwsFilteredRows() {
            const cat = iwsCatFilter ? iwsCatFilter.value : '';
            return cat ? iwsLastRows.filter(r => r.category === cat) : iwsLastRows;
        }

        async function generateIws() {
            if (!activeBar.barName) return;
            const fromDate = iwsDateFrom ? iwsDateFrom.value : iwsFyStart;
            const toDate   = iwsDateTo   ? iwsDateTo.value   : iwsFyEnd;
            if (!fromDate || !toDate || fromDate > toDate) return;

            if (iwsFromLabel) iwsFromLabel.textContent = iwsFmtDate(fromDate);
            if (iwsToLabel)   iwsToLabel.textContent   = iwsFmtDate(toDate);

            const params = { barName: activeBar.barName, financialYear: activeBar.financialYear || '' };
            const result = await window.electronAPI.getDailySales(params);
            if (!result.success || !result.sales) {
                iwsLastRows = [];
                renderIwsTable([], fromDate, toDate);
                return;
            }

            // Aggregate by (brandName + size)
            const aggMap = new Map();
            for (const bill of result.sales) {
                const billDate = bill.billDate || '';
                if (billDate < fromDate || billDate > toDate) continue;
                for (const item of (bill.items || [])) {
                    const key = (item.brandName || '') + '|||' + (item.size || '');
                    if (!aggMap.has(key)) {
                        aggMap.set(key, {
                            brandName: item.brandName || '',
                            size: item.size || '',
                            category: item.category || '',
                            bottles: 0,
                            amount: 0,
                            mlPerBottle: iwsGetML(item.size),
                        });
                    }
                    const r = aggMap.get(key);
                    r.bottles += item.qty || 0;
                    r.amount  += item.amount || 0;
                }
            }

            // Sort: category → brandName → size
            const rows = [...aggMap.values()].sort((a, b) => {
                const cc = (a.category || '').localeCompare(b.category || '');
                if (cc !== 0) return cc;
                const bc = (a.brandName || '').localeCompare(b.brandName || '');
                if (bc !== 0) return bc;
                return (a.size || '').localeCompare(b.size || '');
            });
            rows.forEach((r, i) => { r.srNo = i + 1; r.mlQty = r.bottles * r.mlPerBottle; });

            iwsLastRows = rows;
            iwsLastFrom  = fromDate;
            iwsLastTo    = toDate;

            // Populate category filter
            if (iwsCatFilter) {
                const prevSel = iwsCatFilter.value;
                const cats = [...new Set(rows.map(r => r.category).filter(Boolean))].sort();
                iwsCatFilter.innerHTML = '<option value="">All Categories</option>' +
                    cats.map(c => `<option value="${c}"${c === prevSel ? ' selected' : ''}>${c}</option>`).join('');
            }

            renderIwsTable(getIwsFilteredRows(), fromDate, toDate);
        }

        function renderIwsTable(rows, fromDate, toDate) {
            // Bar info header
            if (iwsBarInfo) {
                const licNo   = activeBar.licenseNo || activeBar.licNo || '—';
                const address = [activeBar.address, activeBar.area, activeBar.city, activeBar.state, activeBar.pinCode].filter(Boolean).join(', ') || '—';
                iwsBarInfo.innerHTML = `
                    <div class="iws-bar-header">
                        <div class="iws-bar-name">${activeBar.barName || '—'}</div>
                        <div class="iws-bar-meta">
                            <span><strong>Lic No:</strong> ${licNo}</span>
                            <span class="iws-dot">·</span>
                            <span><strong>Address:</strong> ${address}</span>
                        </div>
                        ${fromDate ? `<div class="iws-bar-period"><strong>Period:</strong> ${iwsFmtDate(fromDate)} to ${iwsFmtDate(toDate)}</div>` : ''}
                    </div>`;
            }

            // Summary cards
            const totalBottles = rows.reduce((a, r) => a + r.bottles, 0);
            const totalML      = rows.reduce((a, r) => a + r.mlQty, 0);
            const totalAmount  = rows.reduce((a, r) => a + r.amount, 0);
            if (iwsTotalItems)   iwsTotalItems.textContent   = rows.length.toLocaleString('en-IN');
            if (iwsTotalBottles) iwsTotalBottles.textContent = iwsFmtNum(totalBottles);
            if (iwsTotalML)      iwsTotalML.textContent      = (totalML / 1000).toFixed(3) + ' Ltr';
            if (iwsTotalAmount)  iwsTotalAmount.textContent  = '₹' + iwsFmtAmt(totalAmount);

            if (rows.length === 0) {
                if (iwsEmpty)     iwsEmpty.classList.remove('hidden');
                if (iwsTableHead) iwsTableHead.innerHTML = '';
                if (iwsTableBody) iwsTableBody.innerHTML = '';
                if (iwsTableFoot) iwsTableFoot.innerHTML = '';
                return;
            }
            if (iwsEmpty) iwsEmpty.classList.add('hidden');

            // Header
            if (iwsTableHead) {
                iwsTableHead.innerHTML = `<tr>
                    <th class="iws-th iws-th-sr">Sr.No</th>
                    <th class="iws-th iws-th-item">Item Name</th>
                    <th class="iws-th iws-th-size">Size</th>
                    <th class="iws-th iws-th-btl">Bottles</th>
                    <th class="iws-th iws-th-ml">ML Qty (ml)</th>
                    <th class="iws-th iws-th-ltr">Bulk Ltr</th>
                    <th class="iws-th iws-th-amt">Amount (₹)</th>
                </tr>`;
            }

            // Body
            let bodyHtml = '';
            let prevCat = null;
            for (const r of rows) {
                if (r.category !== prevCat) {
                    bodyHtml += `<tr class="iws-cat-row"><td colspan="7">${r.category || 'Uncategorized'}</td></tr>`;
                    prevCat = r.category;
                }
                bodyHtml += `<tr>
                    <td class="iws-td iws-td-sr">${r.srNo}</td>
                    <td class="iws-td iws-td-item">${r.brandName}</td>
                    <td class="iws-td iws-td-size">${r.size}</td>
                    <td class="iws-td iws-td-btl">${iwsFmtNum(r.bottles)}</td>
                    <td class="iws-td iws-td-ml">${iwsFmtNum(r.mlQty)}</td>
                    <td class="iws-td iws-td-ltr">${(r.mlQty / 1000).toFixed(3)}</td>
                    <td class="iws-td iws-td-amt">${iwsFmtAmt(r.amount)}</td>
                </tr>`;
            }
            if (iwsTableBody) iwsTableBody.innerHTML = bodyHtml;

            // Footer
            if (iwsTableFoot) {
                iwsTableFoot.innerHTML = `<tr class="iws-total-row">
                    <td colspan="3"><strong>TOTAL</strong></td>
                    <td><strong>${iwsFmtNum(totalBottles)}</strong></td>
                    <td><strong>${iwsFmtNum(totalML)}</strong></td>
                    <td><strong>${(totalML / 1000).toFixed(3)}</strong></td>
                    <td><strong>${iwsFmtAmt(totalAmount)}</strong></td>
                </tr>`;
            }
        }

        /* ── Export CSV ── */
        function exportIwsCSV() {
            if (!iwsLastRows.length) return;
            const filtered = getIwsFilteredRows();
            const headers = ['Sr.No', 'Item Name', 'Size', 'Bottles', 'ML Qty (ml)', 'Bulk Ltr', 'Amount (INR)'];
            const csvRows = [headers.join(',')];
            for (const r of filtered) {
                csvRows.push([r.srNo, `"${r.brandName}"`, `"${r.size}"`, r.bottles, r.mlQty, (r.mlQty / 1000).toFixed(3), r.amount.toFixed(2)].join(','));
            }
            const blob = new Blob(['\uFEFF' + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `ItemWiseSale_${activeBar.barName || 'Bar'}_${iwsDateFrom?.value}_to_${iwsDateTo?.value}.csv`;
            a.click(); URL.revokeObjectURL(a.href);
        }

        /* ── Print ── */
        function printIwsReport() {
            if (!iwsLastRows.length) return;
            const filtered  = getIwsFilteredRows();
            const barName   = activeBar.barName || '';
            const licNo     = activeBar.licenseNo || activeBar.licNo || '—';
            const address   = [activeBar.address, activeBar.area, activeBar.city, activeBar.state, activeBar.pinCode].filter(Boolean).join(', ') || '—';
            const fromStr   = iwsDateFrom ? iwsDateFrom.value : '';
            const toStr     = iwsDateTo   ? iwsDateTo.value   : '';
            const catLabel  = (iwsCatFilter && iwsCatFilter.value) ? iwsCatFilter.value : 'All Categories';
            const totalBottles = filtered.reduce((a, r) => a + r.bottles, 0);
            const totalML      = filtered.reduce((a, r) => a + r.mlQty, 0);
            const totalAmount  = filtered.reduce((a, r) => a + r.amount, 0);

            let prevCat = null;
            let tbody = '';
            for (const r of filtered) {
                if (r.category !== prevCat) {
                    tbody += `<tr><td colspan="7" style="background:#e8f5e9;font-weight:700;color:#1b5e20;text-align:left;padding:4px 6px">${r.category || 'Uncategorized'}</td></tr>`;
                    prevCat = r.category;
                }
                tbody += `<tr>
                    <td>${r.srNo}</td>
                    <td style="text-align:left">${r.brandName}</td>
                    <td>${r.size}</td>
                    <td>${r.bottles.toLocaleString('en-IN')}</td>
                    <td>${r.mlQty.toLocaleString('en-IN')}</td>
                    <td>${(r.mlQty / 1000).toFixed(3)}</td>
                    <td style="text-align:right">${r.amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                </tr>`;
            }
            tbody += `<tr style="font-weight:700;background:#f0f4ff">
                <td colspan="3" style="text-align:left">TOTAL</td>
                <td>${totalBottles.toLocaleString('en-IN')}</td>
                <td>${totalML.toLocaleString('en-IN')}</td>
                <td>${(totalML / 1000).toFixed(3)}</td>
                <td style="text-align:right">${totalAmount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
            </tr>`;

            printWithIframe(`<!DOCTYPE html><html><head><title>Item Wise Sale — ${barName}</title>
                <style>
                    body{font-family:'Inter',Arial,sans-serif;padding:12px;font-size:8pt}
                    .bar-header{border-bottom:2px solid #1b5e20;padding-bottom:6px;margin-bottom:6px}
                    .bar-name{font-size:13pt;font-weight:800;color:#1b5e20}
                    .bar-meta{font-size:7.5pt;color:#444;margin-top:2px}
                    .report-title{font-size:11pt;font-weight:700;text-align:center;margin:4px 0;color:#333}
                    .period{text-align:center;font-size:7pt;color:#666;margin-bottom:6px}
                    table{border-collapse:collapse;width:100%}
                    th,td{border:1px solid #ccc;padding:3px 5px;font-size:7pt;text-align:center}
                    th{background:#e8f5e9;color:#1b5e20;font-weight:700}
                    td:nth-child(2){text-align:left}
                    @page{size:A4 portrait;margin:8mm}
                </style>
            </head><body>
                <div class="bar-header">
                    <div class="bar-name">${barName}</div>
                    <div class="bar-meta">Lic No: <strong>${licNo}</strong> &nbsp;|&nbsp; ${address}</div>
                </div>
                <div class="report-title">Item Wise Sale Report</div>
                <div class="period">Period: <strong>${iwsFmtDate(fromStr)}</strong> to <strong>${iwsFmtDate(toStr)}</strong> &nbsp;|&nbsp; Category: <strong>${catLabel}</strong> &nbsp;|&nbsp; Generated: ${new Date().toLocaleString('en-IN')}</div>
                <table>
                    <thead><tr><th>Sr.No</th><th>Item Name</th><th>Size</th><th>Bottles</th><th>ML Qty (ml)</th><th>Bulk Ltr</th><th>Amount (&#x20B9;)</th></tr></thead>
                    <tbody>${tbody}</tbody>
                </table>
            </body></html>`);
        }

        /* ── Event wiring ── */
        if (iwsGoBtn)     iwsGoBtn.addEventListener('click', generateIws);
        if (iwsExportBtn) iwsExportBtn.addEventListener('click', exportIwsCSV);
        if (iwsPrintBtn)  iwsPrintBtn.addEventListener('click', printIwsReport);
        if (iwsCatFilter) iwsCatFilter.addEventListener('change', () => {
            if (iwsLastRows.length) renderIwsTable(getIwsFilteredRows(), iwsLastFrom, iwsLastTo);
        });
        [iwsDateFrom, iwsDateTo].forEach(el => {
            el?.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); generateIws(); } });
        });

        // Auto-generate when the panel first becomes active
        const iwsPanel = document.getElementById('sub-item-wise-sale');
        if (iwsPanel) {
            const iwsObs = new MutationObserver(() => {
                if (iwsPanel.classList.contains('active') && !iwsLastRows.length && activeBar.barName) {
                    generateIws();
                }
            });
            iwsObs.observe(iwsPanel, { attributes: true, attributeFilter: ['class'] });
        }
    }
    // ═══════════════════════════════════════════════════════

});
