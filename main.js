const { app, BrowserWindow, ipcMain, shell, dialog } = require('electron');
const { autoUpdater } = require('electron-updater');
const path = require('path');
const fs = require('fs');

// ── Auto-Updater Setup ─────────────────────────────────────
autoUpdater.autoDownload = false;
autoUpdater.autoInstallOnAppQuit = true;
autoUpdater.on('download-progress', (progress) => {
  if (mainWindow) mainWindow.webContents.send('update-download-progress', progress);
});
autoUpdater.on('update-downloaded', () => {
  if (mainWindow) mainWindow.webContents.send('update-downloaded');
});

let mainWindow;

// ── Config & Data-Root ─────────────────────────────────────
function getConfigPath() {
  return path.join(app.getPath('userData'), 'spliqour-config.json');
}
let _cfgCache = null;
function loadConfig() {
  if (_cfgCache) return _cfgCache;
  try {
    const cp = getConfigPath();
    if (fs.existsSync(cp)) {
      _cfgCache = JSON.parse(fs.readFileSync(cp, 'utf8'));
      return _cfgCache;
    }
  } catch (e) {}
  return {};
}
function saveConfig(cfg) {
  _cfgCache = cfg;
  const cp = getConfigPath();
  fs.mkdirSync(path.dirname(cp), { recursive: true });
  fs.writeFileSync(cp, JSON.stringify(cfg, null, 2));
}
function getDataRoot() {
  const cfg = loadConfig();
  return cfg.dataRoot || path.join(app.getPath('userData'), 'SpliqourProData');
}
// Strips legacy 'data/' prefix stored in older bars_index entries
function normalizeFolderPath(fp) {
  return (fp || '').replace(/^data[\/\\]/, '');
}
// Path to the tesseract lang data (works in both dev & packaged)
function getTesseractLangPath() {
  return app.isPackaged ? process.resourcesPath : __dirname;
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    },
    titleBarStyle: 'hidden',
    titleBarOverlay: {
      color: '#13132e',
      symbolColor: '#9ca3af',
      height: 38
    }
  });

  // Start at the bar-selection home screen
  if (!app.isPackaged) mainWindow.webContents.openDevTools();
  mainWindow.loadFile('home.html');

  // ── Notify renderer of window state changes (topbar resize) ──
  function sendWindowState() {
    if (!mainWindow || mainWindow.isDestroyed()) return;
    mainWindow.webContents.send('window-state-changed', {
      maximized: mainWindow.isMaximized(),
    });
  }
  mainWindow.on('maximize',   sendWindowState);
  mainWindow.on('unmaximize', sendWindowState);
  mainWindow.on('restore',    sendWindowState);
  mainWindow.on('resize',     sendWindowState);
}

// ── Dynamic title bar overlay (theme-aware window buttons) ──
ipcMain.handle('set-title-bar-overlay', (event, opts) => {
  if (mainWindow && !mainWindow.isDestroyed()) {
    try { mainWindow.setTitleBarOverlay(opts); } catch(e) {}
  }
});

// ── Custom window controls ──
ipcMain.on('window-minimize', () => {
  if (mainWindow && !mainWindow.isDestroyed()) mainWindow.minimize();
});
ipcMain.on('window-maximize', () => {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.isMaximized() ? mainWindow.unmaximize() : mainWindow.maximize();
  }
});
ipcMain.on('window-close', () => {
  if (mainWindow && !mainWindow.isDestroyed()) mainWindow.close();
});

app.whenReady().then(async () => {
  // ── First-run: ensure data root is configured ───────────
  const cfg = loadConfig();
  if (!cfg.dataRoot) {
    if (!app.isPackaged) {
      // Development: use the project's own data/ folder
      cfg.dataRoot = path.join(__dirname, 'data');
      fs.mkdirSync(cfg.dataRoot, { recursive: true });
      saveConfig(cfg);
    } else {
      // Production: create a small visible setup window to parent dialogs
      // (dialogs without a parent window may not appear/focus on Windows)
      const setupWin = new BrowserWindow({
        width: 480,
        height: 240,
        center: true,
        resizable: false,
        alwaysOnTop: true,
        show: true,
        frame: true,
        title: 'SpliqourPro — First-Run Setup',
        webPreferences: { nodeIntegration: false, contextIsolation: true }
      });
      // Load a blank page so the window has something to render
      setupWin.loadURL('data:text/html,<body style="background:#1e1e24;color:#fff;font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0"><h2 style="opacity:.7">Setting up SpliqourPro…</h2></body>');

      await dialog.showMessageBox(setupWin, {
        type: 'info',
        title: 'SpliqourPro — First-Run Setup',
        message: 'Welcome to SpliqourPro!',
        detail: 'Please choose a folder where SpliqourPro will store all your bar data (sales, inventory, reports, etc.).\n\nYou can change this later from Settings.',
        buttons: ['Choose Folder'],
        defaultId: 0
      });

      // Default data folder = user's Documents (survives app updates/reinstalls)
      const defaultDataPath = path.join(app.getPath('documents'), 'SpliqourPro Data');

      const pick = await dialog.showOpenDialog(setupWin, {
        title: 'Choose Data Folder for SpliqourPro',
        properties: ['openDirectory', 'createDirectory'],
        defaultPath: defaultDataPath
      });

      if (!pick.canceled && pick.filePaths.length > 0) {
        cfg.dataRoot = pick.filePaths[0];
      } else {
        cfg.dataRoot = defaultDataPath;
      }
      fs.mkdirSync(cfg.dataRoot, { recursive: true });
      saveConfig(cfg);

      if (!setupWin.isDestroyed()) setupWin.close();
    }
  } else {
    // Ensure folder still exists
    fs.mkdirSync(cfg.dataRoot, { recursive: true });

    // ── Migrate data out of install dir (if user had old default path) ──
    // Old default was next to the .exe (C:\Program Files\SpliqourPro\SpliqourPro Data)
    // which gets wiped on update. Move it to Documents if still in install dir.
    if (app.isPackaged) {
      const installDir = path.normalize(path.dirname(app.getPath('exe')));
      const currentRoot = path.normalize(cfg.dataRoot);
      if (currentRoot.startsWith(installDir)) {
        const newRoot = path.join(app.getPath('documents'), 'SpliqourPro Data');
        if (!fs.existsSync(newRoot) && fs.existsSync(currentRoot)) {
          try {
            fs.cpSync(currentRoot, newRoot, { recursive: true });
            cfg.dataRoot = newRoot;
            saveConfig(cfg);
          } catch (e) {
            // If copy fails, just update the path and let user know on next boot
          }
        } else if (fs.existsSync(newRoot)) {
          cfg.dataRoot = newRoot;
          saveConfig(cfg);
        }
      }
    }
  }

  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ── Helpers ────────────────────────────────────────────────
function sanitizeFolderName(name) {
  return name.trim()
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, '')
    .replace(/\s+/g, '_')
    .substring(0, 60);
}

function fyToFolder(fyLabel) {
  const match = fyLabel.match(/(\d{4})/g);
  if (match && match.length >= 2) {
    return `FY_${match[0]}-${match[1].slice(2)}`;
  }
  return fyLabel.replace(/\//g, '-').replace(/\s+/g, '_');
}

// ── IPC: Get bars index ────────────────────────────────────
ipcMain.handle('get-bars-index', async () => {
  try {
    const indexPath = path.join(getDataRoot(), 'bars_index.json');
    if (!fs.existsSync(indexPath)) return { success: true, bars: [] };
    const raw = fs.readFileSync(indexPath, 'utf8');
    const bars = raw ? JSON.parse(raw) : [];
    return { success: true, bars };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ── IPC: Open a specific bar ───────────────────────────────
ipcMain.handle('open-bar', async (event, barEntry) => {
  try {
    const masterPath = path.join(getDataRoot(), normalizeFolderPath(barEntry.folderPath), 'bar_master.json');
    if (!fs.existsSync(masterPath)) {
      return { success: false, error: 'bar_master.json not found' };
    }
    const data = JSON.parse(fs.readFileSync(masterPath, 'utf8'));
    return { success: true, data };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ── IPC: Navigate back to Home ─────────────────────────────
ipcMain.on('navigate-home', () => {
  if (mainWindow) mainWindow.loadFile('home.html');
});

// ── IPC: Backup Data ───────────────────────────────────────
ipcMain.handle('backup-data', async (event, params) => {
  try {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: 'Select Backup Destination Folder',
      properties: ['openDirectory', 'createDirectory']
    });
    if (result.canceled || !result.filePaths.length) {
      return { success: false, error: 'Backup cancelled' };
    }
    const destBase = result.filePaths[0];
    const barFolder = sanitizeFolderName(params.barName || 'Unknown_Bar');
    const fyFolder = params.financialYear ? fyToFolder(params.financialYear) : 'FY_Unknown';
    const srcDir = path.join(getDataRoot(), barFolder, fyFolder);
    if (!fs.existsSync(srcDir)) {
      return { success: false, error: 'Source data folder not found' };
    }
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const backupFolderName = `SPLIQOUR_Backup_${barFolder}_${timestamp}`;
    const destDir = path.join(destBase, backupFolderName);
    fs.mkdirSync(destDir, { recursive: true });
    // Copy all files recursively
    function copyDirSync(src, dest) {
      fs.mkdirSync(dest, { recursive: true });
      const entries = fs.readdirSync(src, { withFileTypes: true });
      for (const entry of entries) {
        const srcPath = path.join(src, entry.name);
        const destPath = path.join(dest, entry.name);
        if (entry.isDirectory()) {
          copyDirSync(srcPath, destPath);
        } else {
          fs.copyFileSync(srcPath, destPath);
        }
      }
    }
    copyDirSync(srcDir, destDir);
    return { success: true, path: destDir };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ── IPC: Transfer to Next Financial Year ───────────────────
ipcMain.handle('transfer-to-next-fy', async (event, params) => {
  try {
    const barFolder = sanitizeFolderName(params.barName || 'Unknown_Bar');
    const fyMatch = (params.financialYear || '').match(/(\d{4})/g);
    if (!fyMatch || fyMatch.length < 2) {
      return { success: false, error: 'Invalid financial year format' };
    }
    const nextStart = parseInt(fyMatch[1]);
    const nextEnd = nextStart + 1;
    const newFYLabel = `${nextStart}-${nextEnd}`;
    const newFYFolder = `FY_${nextStart}-${String(nextEnd).slice(2)}`;
    const currentFYFolder = fyToFolder(params.financialYear);
    const srcDir = path.join(getDataRoot(), barFolder, currentFYFolder);
    const destDir = path.join(getDataRoot(), barFolder, newFYFolder);
    if (fs.existsSync(destDir)) {
      return { success: false, error: `Next FY folder already exists: ${newFYFolder}` };
    }
    fs.mkdirSync(destDir, { recursive: true });
    // Copy products.json as-is
    const productsFile = path.join(srcDir, 'products.json');
    if (fs.existsSync(productsFile)) {
      fs.copyFileSync(productsFile, path.join(destDir, 'products.json'));
    }
    // Copy bar_master.json and update FY
    const masterFile = path.join(srcDir, 'bar_master.json');
    if (fs.existsSync(masterFile)) {
      const master = JSON.parse(fs.readFileSync(masterFile, 'utf8'));
      master.financialYear = newFYLabel;
      fs.writeFileSync(path.join(destDir, 'bar_master.json'), JSON.stringify(master, null, 2));
    }
    // Copy MRP master
    const mrpFile = path.join(srcDir, 'mrp_master.json');
    if (fs.existsSync(mrpFile)) {
      fs.copyFileSync(mrpFile, path.join(destDir, 'mrp_master.json'));
    }
    // Copy shortcuts
    const shortcutsFile = path.join(srcDir, 'shortcuts.json');
    if (fs.existsSync(shortcutsFile)) {
      fs.copyFileSync(shortcutsFile, path.join(destDir, 'shortcuts.json'));
    }
    // Copy suppliers
    const suppliersFile = path.join(srcDir, 'suppliers.json');
    if (fs.existsSync(suppliersFile)) {
      fs.copyFileSync(suppliersFile, path.join(destDir, 'suppliers.json'));
    }
    // Copy customers
    const customersFile = path.join(srcDir, 'customers.json');
    if (fs.existsSync(customersFile)) {
      fs.copyFileSync(customersFile, path.join(destDir, 'customers.json'));
    }
    // Create opening stock from current stock (empty for now — can be enhanced)
    fs.writeFileSync(path.join(destDir, 'opening_stock.json'), '[]');
    // Initialize empty data files
    fs.writeFileSync(path.join(destDir, 'daily_sales.json'), '[]');
    fs.writeFileSync(path.join(destDir, 'tp_entries.json'), '[]');
    fs.writeFileSync(path.join(destDir, 'sale_counter.json'), '{"lastNumber": 0}');
    // Add to bars index
    const indexPath = path.join(getDataRoot(), 'bars_index.json');
    let index = [];
    if (fs.existsSync(indexPath)) {
      const raw = fs.readFileSync(indexPath, 'utf8');
      if (raw) index = JSON.parse(raw);
    }
    const existingBar = index.find(b => b.barName === params.barName);
    index.push({
      bar_id: (existingBar?.bar_id || '') + '_' + newFYFolder,
      barName: params.barName,
      financialYear: newFYLabel,
      shopType: existingBar?.shopType || '',
      city: existingBar?.city || '',
      folderPath: path.join(barFolder, newFYFolder),
      created_at: new Date().toISOString()
    });
    fs.writeFileSync(indexPath, JSON.stringify(index, null, 2));
    return { success: true, newFY: `FY ${nextStart}-${String(nextEnd).slice(2)}` };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ── IPC: Navigate to Main App (after bar selection) ───────
ipcMain.on('navigate-to-app', () => {
  if (mainWindow) mainWindow.loadFile('app.html');
});

// ── IPC: Navigate to Add Bar form ─────────────────────────
ipcMain.on('navigate-to-add-bar', () => {
  if (mainWindow) mainWindow.loadFile('index.html');
});

// ── IPC: Save Bar Data ─────────────────────────────────────
ipcMain.handle('save-bar-data', async (event, barData) => {
  try {
    const barName = barData.barName || 'Unknown_Bar';
    const fyLabel = barData.financialYear || '';

    const barFolder = sanitizeFolderName(barName);
    const fyFolder = fyLabel ? fyToFolder(fyLabel) : 'FY_Unknown';

    const dirPath = path.join(getDataRoot(), barFolder, fyFolder);
    fs.mkdirSync(dirPath, { recursive: true });

    const masterPath = path.join(dirPath, 'bar_master.json');
    fs.writeFileSync(masterPath, JSON.stringify(barData, null, 2));

    const indexPath = path.join(getDataRoot(), 'bars_index.json');
    let index = [];
    if (fs.existsSync(indexPath)) {
      const raw = fs.readFileSync(indexPath, 'utf8');
      if (raw) index = JSON.parse(raw);
    }

    const existingIdx = index.findIndex(b => b.bar_id === barData.bar_id);
    const indexEntry = {
      bar_id: barData.bar_id,
      barName,
      financialYear: fyLabel,
      shopType: barData.shopType || '',
      city: barData.city || '',
      folderPath: path.join(barFolder, fyFolder),
      created_at: barData.created_at
    };

    if (existingIdx >= 0) index[existingIdx] = indexEntry;
    else index.push(indexEntry);

    fs.writeFileSync(indexPath, JSON.stringify(index, null, 2));
    return { success: true, folderPath: dirPath };
  } catch (error) {
    console.error('Error saving bar data:', error);
    return { success: false, error: error.message };
  }
});

// ── IPC: Supplier CRUD ─────────────────────────────────────
function getBarDir(barName, financialYear) {
  const barFolder = sanitizeFolderName(barName);
  const fyFolder = financialYear ? fyToFolder(financialYear) : 'FY_Unknown';
  return path.join(getDataRoot(), barFolder, fyFolder);
}

ipcMain.handle('get-suppliers', async (event, { barName, financialYear }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'suppliers.json');
    if (!fs.existsSync(filePath)) return { success: true, suppliers: [] };
    const raw = fs.readFileSync(filePath, 'utf8');
    return { success: true, suppliers: raw ? JSON.parse(raw) : [] };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('save-supplier', async (event, { barName, financialYear, supplier }) => {
  try {
    const dir = getBarDir(barName, financialYear);
    fs.mkdirSync(dir, { recursive: true });
    const filePath = path.join(dir, 'suppliers.json');
    let suppliers = [];
    if (fs.existsSync(filePath)) {
      const raw = fs.readFileSync(filePath, 'utf8');
      if (raw) suppliers = JSON.parse(raw);
    }
    const idx = suppliers.findIndex(s => s.id === supplier.id);
    if (idx >= 0) {
      suppliers[idx] = { ...suppliers[idx], ...supplier, updatedAt: new Date().toISOString() };
    } else {
      supplier.createdAt = new Date().toISOString();
      suppliers.push(supplier);
    }
    fs.writeFileSync(filePath, JSON.stringify(suppliers, null, 2));
    return { success: true, supplier };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('delete-supplier', async (event, { barName, financialYear, supplierId }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'suppliers.json');
    if (!fs.existsSync(filePath)) return { success: false, error: 'No suppliers file' };
    let suppliers = JSON.parse(fs.readFileSync(filePath, 'utf8') || '[]');
    suppliers = suppliers.filter(s => s.id !== supplierId);
    fs.writeFileSync(filePath, JSON.stringify(suppliers, null, 2));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ── IPC: Customer CRUD ─────────────────────────────────────
ipcMain.handle('get-customers', async (event, { barName, financialYear }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'customers.json');
    if (!fs.existsSync(filePath)) return { success: true, customers: [] };
    const raw = fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, '');
    return { success: true, customers: raw ? JSON.parse(raw) : [] };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('save-customer', async (event, { barName, financialYear, customer }) => {
  try {
    const dir = getBarDir(barName, financialYear);
    fs.mkdirSync(dir, { recursive: true });
    const filePath = path.join(dir, 'customers.json');
    let customers = [];
    if (fs.existsSync(filePath)) {
      const raw = fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, '');
      if (raw) customers = JSON.parse(raw);
    }
    const idx = customers.findIndex(c => c.id === customer.id);
    if (idx >= 0) {
      customers[idx] = { ...customers[idx], ...customer, updatedAt: new Date().toISOString() };
    } else {
      customer.createdAt = new Date().toISOString();
      customers.push(customer);
    }
    fs.writeFileSync(filePath, JSON.stringify(customers, null, 2));
    return { success: true, customer };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('delete-customer', async (event, { barName, financialYear, customerId }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'customers.json');
    if (!fs.existsSync(filePath)) return { success: false, error: 'No customers file' };
    let customers = JSON.parse(fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, '') || '[]');
    customers = customers.filter(c => c.id !== customerId);
    fs.writeFileSync(filePath, JSON.stringify(customers, null, 2));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ── IPC: Daily Sales ───────────────────────────────────────
ipcMain.handle('get-daily-sales', async (event, { barName, financialYear, date }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'daily_sales.json');
    if (!fs.existsSync(filePath)) return { success: true, sales: [] };
    const raw = fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, '');
    let sales = raw ? JSON.parse(raw) : [];
    if (date) sales = sales.filter(s => s.billDate === date);
    return { success: true, sales };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('save-daily-sale', async (event, { barName, financialYear, sale }) => {
  try {
    const dir = getBarDir(barName, financialYear);
    fs.mkdirSync(dir, { recursive: true });
    const filePath = path.join(dir, 'daily_sales.json');
    let sales = [];
    if (fs.existsSync(filePath)) {
      const raw = fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, '');
      if (raw) sales = JSON.parse(raw);
    }
    const idx = sales.findIndex(s => s.id === sale.id);
    if (idx >= 0) {
      sales[idx] = { ...sales[idx], ...sale, updatedAt: new Date().toISOString() };
    } else {
      sale.createdAt = new Date().toISOString();
      sales.push(sale);
    }
    fs.writeFileSync(filePath, JSON.stringify(sales, null, 2));
    return { success: true, sale };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('delete-daily-sale', async (event, { barName, financialYear, saleId }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'daily_sales.json');
    if (!fs.existsSync(filePath)) return { success: false, error: 'No sales file' };
    let sales = JSON.parse(fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, '') || '[]');
    sales = sales.filter(s => s.id !== saleId);
    fs.writeFileSync(filePath, JSON.stringify(sales, null, 2));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('get-sale-counter', async (event, { barName, financialYear }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'sale_counter.json');
    if (!fs.existsSync(filePath)) return { success: true, counter: null };
    const raw = fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, '');
    return { success: true, counter: raw ? JSON.parse(raw) : null };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('save-sale-counter', async (event, { barName, financialYear, counter }) => {
  try {
    const dir = getBarDir(barName, financialYear);
    fs.mkdirSync(dir, { recursive: true });
    const filePath = path.join(dir, 'sale_counter.json');
    fs.writeFileSync(filePath, JSON.stringify(counter, null, 2));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('get-current-stock', async (event, { barName, financialYear, asOfDate }) => {
  try {
    const dir = getBarDir(barName, financialYear);

    // 1. Opening stock
    const osPath = path.join(dir, 'opening_stock.json');
    let openingEntries = [];
    if (fs.existsSync(osPath)) {
      const osRaw = fs.readFileSync(osPath, 'utf8').replace(/^\uFEFF/, '');
      if (osRaw) {
        const osData = JSON.parse(osRaw);
        openingEntries = osData.entries || [];
      }
    }

    // 2. TP receipts
    const tpPath = path.join(dir, 'tp_entries.json');
    let tpEntries = [];
    if (fs.existsSync(tpPath)) {
      const tpRaw = fs.readFileSync(tpPath, 'utf8').replace(/^\uFEFF/, '');
      if (tpRaw) tpEntries = JSON.parse(tpRaw);
    }

    // 3. Daily sales
    const dsPath = path.join(dir, 'daily_sales.json');
    let dailySales = [];
    if (fs.existsSync(dsPath)) {
      const dsRaw = fs.readFileSync(dsPath, 'utf8').replace(/^\uFEFF/, '');
      if (dsRaw) dailySales = JSON.parse(dsRaw);
    }

    // Build stock map: productId => { brandName, code, size, bpc, totalBtl }
    const stockMap = {};

    // Add opening stock
    for (const e of openingEntries) {
      if (!stockMap[e.productId]) {
        stockMap[e.productId] = { productId: e.productId, brandName: e.brandName, code: e.code, size: e.size, bpc: e.bpc || 1, totalBtl: 0 };
      }
      stockMap[e.productId].totalBtl += (e.totalBtl || 0);
    }

    // Add TP receipts — if asOfDate supplied, only include TPs received on/before that date
    for (const tp of tpEntries) {
      if (asOfDate && tp.tpDate && tp.tpDate > asOfDate) continue;
      for (const item of (tp.items || [])) {
        const key = Object.keys(stockMap).find(k => stockMap[k].code === item.code) || ('code_' + item.code);
        if (!stockMap[key]) {
          stockMap[key] = { productId: key, brandName: item.brand || item.brandName, code: item.code, size: item.size, bpc: item.bpc || 1, totalBtl: 0 };
        }
        stockMap[key].totalBtl += (item.totalBtl || 0);
      }
    }

    // Subtract daily sales — if asOfDate supplied, only subtract sales BEFORE that date
    // (sales ON the date are handled live in the renderer so the user sees real-time stock)
    for (const sale of dailySales) {
      if (asOfDate && sale.billDate >= asOfDate) continue;
      for (const item of (sale.items || [])) {
        const key = Object.keys(stockMap).find(k => stockMap[k].code === item.code) || item.productId;
        if (stockMap[key]) {
          stockMap[key].totalBtl -= (item.qty || 0);
        }
      }
    }

    const stock = Object.values(stockMap);
    return { success: true, stock };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ── IPC: Open External URL ─────────────────────────────────
ipcMain.handle('open-external-url', async (event, url) => {
  try {
    if (url && (url.startsWith('https://') || url.startsWith('http://'))) {
      await shell.openExternal(url);
      return { success: true };
    }
    return { success: false, error: 'Invalid URL' };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ── IPC: Open URL with Auto-Login + Tesseract CAPTCHA solver ─
ipcMain.handle('open-url-autologin', async (event, { url, username, password }) => {
  try {
    const { createWorker } = require('tesseract.js');

    const win = new BrowserWindow({
      width: 1280,
      height: 800,
      title: 'SCM Retailer Portal — Auto Login',
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        webSecurity: false,   // allow cross-origin canvas reads for captcha OCR
      }
    });
    win.loadURL(url);

    // ── Fill inputs in the portal page ──
    const fillInputs = (u, p, captchaAnswer) => `(function(u,p,ca){
      function fill(el,val){
        if(!el) return;
        el.value=val;
        ['input','change','keyup'].forEach(function(t){el.dispatchEvent(new Event(t,{bubbles:true}));});
      }
      var userSel=['#txtUserName','#txtUserId','#UserName','input[name="UserName"]','input[name="txtUserName"]','input[id*="UserName"]'];
      var passSel=['#txtPassword','#Password','input[name="Password"]','input[name="txtPassword"]','input[type="password"]'];
      var capSel=['input[name*="aptcha"]','input[id*="aptcha"]','input[placeholder*="TYPE"]','input[placeholder*="type"]','input[id*="Captcha"]','input[name*="Captcha"]'];
      var uf=null,pf=null,cf=null;
      for(var s of userSel){try{uf=document.querySelector(s);if(uf)break;}catch(e){}}
      for(var s of passSel){try{pf=document.querySelector(s);if(pf)break;}catch(e){}}
      for(var s of capSel){try{cf=document.querySelector(s);if(cf)break;}catch(e){}}
      if(!cf){
        var all=Array.from(document.querySelectorAll('input[type="text"],input:not([type])'));
        var rem=all.filter(function(i){return i!==uf&&i!==pf&&i.offsetParent!==null&&!i.value;});
        cf=rem[rem.length-1]||null;
      }
      fill(uf,u); fill(pf,p);
      if(ca!==null&&cf) fill(cf,String(ca));
      return {user:!!uf,pass:!!pf,cap:!!cf};
    })(${JSON.stringify(u)},${JSON.stringify(p)},${JSON.stringify(captchaAnswer)})`;

    // ── Read captcha via canvas (uses browser session — no separate HTTP fetch) ──
    // Upscales 4x and applies grayscale+threshold for much better OCR accuracy.
    const getCaptchaDataUrl = `(function(){
      return new Promise(function(resolve){
        var imgSels=['img[id*="aptcha"]','img[id*="Captcha"]','img[name*="aptcha"]',
          'img[src*="captcha"]','img[src*="Captcha"]','img[src*="CaptchaImage"]',
          'img[src*="GetCaptcha"]','img[src*="generatecaptcha"]'];
        var el=null;
        for(var s of imgSels){el=document.querySelector(s);if(el)break;}
        if(!el){
          var capInput=null;
          var cSel=['input[name*="aptcha"]','input[id*="aptcha"]','input[placeholder*="TYPE"]','input[placeholder*="type"]'];
          for(var s of cSel){capInput=document.querySelector(s);if(capInput)break;}
          if(capInput){
            var par=capInput.closest('tr,td,div,form')||capInput.parentElement;
            if(par)el=par.querySelector('img');
          }
        }
        if(!el){return resolve(null);}
        function draw(){
          try{
            var scale=4;
            var W=(el.naturalWidth||el.width||150)*scale;
            var H=(el.naturalHeight||el.height||50)*scale;
            var c=document.createElement('canvas'); c.width=W; c.height=H;
            var ctx=c.getContext('2d');
            ctx.imageSmoothingEnabled=false;
            ctx.drawImage(el,0,0,W,H);
            var id=ctx.getImageData(0,0,W,H); var d=id.data;
            var sum=0; for(var i=0;i<d.length;i+=4) sum+=(0.299*d[i]+0.587*d[i+1]+0.114*d[i+2]);
            var avgLum=sum/(d.length/4);
            var darkBg=avgLum<128;
            for(var i=0;i<d.length;i+=4){
              var gray=0.299*d[i]+0.587*d[i+1]+0.114*d[i+2];
              if(darkBg) gray=255-gray;
              var bin=gray<160?0:255;
              d[i]=d[i+1]=d[i+2]=bin; d[i+3]=255;
            }
            ctx.putImageData(id,0,0);
            resolve(c.toDataURL('image/png'));
          }catch(e){resolve(null);}
        }
        if(el.complete && el.naturalWidth>0){draw();}
        else{el.onload=draw; setTimeout(function(){if(el.naturalWidth>0)draw();else resolve(null);},3000);}
      });
    })()`;

    // ── Parse math from OCR text ──
    function parseMath(rawText) {
      let text = rawText
        .replace(/\r?\n/g, ' ')
        .replace(/[oO]/g, '0')
        .replace(/[lI|!]/g, '1')
        .replace(/[B]/g, '8')
        .replace(/[gq]/g, '9')
        .replace(/[zZ]/g, '2')
        .replace(/[×xX✕]/g, '*')
        .replace(/[—–]/g, '-');

      const patterns = [
        /(\d{1,3})\s*([+\-*])\s*(\d{1,3})/,
        /(\d{1,3})([+\-*])(\d{1,3})/,
      ];
      for (const re of patterns) {
        const m = text.match(re);
        if (!m) continue;
        const a = parseInt(m[1], 10), b = parseInt(m[3], 10);
        if (isNaN(a) || isNaN(b)) continue;
        const op = m[2];
        const result = op === '+' ? a + b : op === '-' ? a - b : a * b;
        console.log(`[AutoLogin] parseMath: "${m[0]}" → ${a} ${op} ${b} = ${result}`);
        return result;
      }
      return null;
    }

    win.webContents.on('did-finish-load', async () => {
      try {
        // Wait for ASP.NET to fully render
        await new Promise(r => setTimeout(r, 2500));

        // Fill user + pass first (captcha = null)
        const fillResult = await win.webContents.executeJavaScript(fillInputs(username, password, null)).catch(() => ({}));
        console.log('[AutoLogin] Fill result:', fillResult);

        // Read captcha image via browser canvas (same session, correct image)
        const dataUrl = await win.webContents.executeJavaScript(getCaptchaDataUrl, true).catch(() => null);
        let imgBuffer;
        if (dataUrl) {
          const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');
          imgBuffer = Buffer.from(base64, 'base64');
          console.log('[AutoLogin] Canvas captcha PNG size:', imgBuffer.length, 'bytes');
        } else {
          // Fallback: capturePage on the captcha element's bounding box
          console.log('[AutoLogin] Canvas failed — falling back to capturePage');
          const rect = await win.webContents.executeJavaScript(`(function(){
            var imgSels=['img[id*="aptcha"]','img[id*="Captcha"]','img[name*="aptcha"]','img[src*="captcha"]','img[src*="Captcha"]'];
            var el=null; for(var s of imgSels){el=document.querySelector(s);if(el)break;}
            if(!el) return null;
            var r=el.getBoundingClientRect();
            return {x:Math.round(r.left),y:Math.round(r.top),width:Math.round(r.width)||150,height:Math.round(r.height)||50};
          })()`).catch(() => null);
          if (!rect || rect.width < 10) {
            console.log('[AutoLogin] capturePage rect invalid — skipping OCR');
            return;
          }
          const ni = await win.webContents.capturePage(rect);
          imgBuffer = ni.toPNG();
          console.log('[AutoLogin] capturePage PNG size:', imgBuffer.length, 'bytes');
        }
        if (!imgBuffer || imgBuffer.length < 100) {
          console.log('[AutoLogin] No valid captcha image — skipping OCR');
          return;
        }

        // Run Tesseract OCR — PSM 7 single-line, OEM 1 LSTM
        let worker;
        try {
          worker = await createWorker('eng', 1, { logger: () => {}, errorHandler: () => {}, langPath: getTesseractLangPath() });
          await worker.setParameters({
            tessedit_char_whitelist: '0123456789+- *=',
            tessedit_pageseg_mode:   '7',
          });
          const { data: { text } } = await worker.recognize(imgBuffer);
          console.log('[AutoLogin] OCR raw text:', JSON.stringify(text));
          const answer = parseMath(text);
          console.log('[AutoLogin] Computed answer:', answer);
          if (answer !== null) {
            await win.webContents.executeJavaScript(fillInputs(username, password, answer)).catch(() => {});
          } else {
            console.log('[AutoLogin] Could not parse math from OCR output');
          }
        } finally {
          if (worker) await worker.terminate().catch(() => {});
        }

      } catch (err) {
        console.error('[AutoLogin] Error during captcha solve:', err && (err.message || String(err)), err);
      }
    });

    return { success: true };
  } catch (err) {
    return { success: false, error: err && (err.message || String(err)) };
  }
});

// ── IPC: Product CRUD ──────────────────────────────────────
ipcMain.handle('get-products', async (event, { barName, financialYear }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'products.json');
    if (!fs.existsSync(filePath)) return { success: true, products: [] };
    const raw = fs.readFileSync(filePath, 'utf8');
    return { success: true, products: raw ? JSON.parse(raw) : [] };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('save-product', async (event, { barName, financialYear, product }) => {
  try {
    const dir = getBarDir(barName, financialYear);
    fs.mkdirSync(dir, { recursive: true });
    const filePath = path.join(dir, 'products.json');
    let products = [];
    if (fs.existsSync(filePath)) {
      const raw = fs.readFileSync(filePath, 'utf8');
      if (raw) products = JSON.parse(raw);
    }
    const idx = products.findIndex(p => p.id === product.id);
    if (idx >= 0) {
      products[idx] = { ...products[idx], ...product, updatedAt: new Date().toISOString() };
    } else {
      product.createdAt = new Date().toISOString();
      products.push(product);
    }
    fs.writeFileSync(filePath, JSON.stringify(products, null, 2));
    return { success: true, product };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('delete-product', async (event, { barName, financialYear, productId }) => {
  try {
    const dir = getBarDir(barName, financialYear);
    const filePath = path.join(dir, 'products.json');
    if (!fs.existsSync(filePath)) return { success: false, error: 'No products file' };
    let products = JSON.parse(fs.readFileSync(filePath, 'utf8') || '[]');

    // Find the product being deleted (need brandName + size for cleaning linked data)
    const deletedProduct = products.find(p => p.id === productId);
    products = products.filter(p => p.id !== productId);
    fs.writeFileSync(filePath, JSON.stringify(products, null, 2));

    // Clean up linked data if product was found
    if (deletedProduct) {
      const brandUpper = (deletedProduct.brandName || '').toUpperCase();
      const sizeStr = String(deletedProduct.size || '');

      // 1) Remove matching entries from Opening Stock
      const osFile = path.join(dir, 'opening_stock.json');
      if (fs.existsSync(osFile)) {
        try {
          const osData = JSON.parse(fs.readFileSync(osFile, 'utf8'));
          if (osData && osData.entries) {
            const before = osData.entries.length;
            osData.entries = osData.entries.filter(e =>
              !((e.brandName || '').toUpperCase() === brandUpper && String(e.size || '') === sizeStr)
            );
            if (osData.entries.length !== before) {
              // Recalculate totals
              osData.totalCases = osData.entries.reduce((s, e) => s + (e.cases || 0), 0);
              osData.totalLoose = osData.entries.reduce((s, e) => s + (e.loose || 0), 0);
              osData.totalBottles = osData.entries.reduce((s, e) => s + (e.totalBtl || 0), 0);
              osData.totalValue = osData.entries.reduce((s, e) => s + (e.value || 0), 0);
              osData.productCount = osData.entries.length;
              osData.updatedAt = new Date().toISOString();
              fs.writeFileSync(osFile, JSON.stringify(osData, null, 2));
            }
          }
        } catch (_) { /* ignore OS cleanup errors */ }
      }

      // 2) Remove matching items from TP entries
      const tpFile = path.join(dir, 'tp_entries.json');
      if (fs.existsSync(tpFile)) {
        try {
          let tps = JSON.parse(fs.readFileSync(tpFile, 'utf8') || '[]');
          let tpChanged = false;
          tps = tps.map(tp => {
            if (!tp.items) return tp;
            const before = tp.items.length;
            tp.items = tp.items.filter(item =>
              !((item.brand || '').toUpperCase() === brandUpper && String(item.size || '') === sizeStr)
            );
            if (tp.items.length !== before) {
              tpChanged = true;
              // Recalculate TP totals
              tp.totalCases = tp.items.reduce((s, i) => s + (i.cases || 0), 0);
              tp.totalLoose = tp.items.reduce((s, i) => s + (i.loose || 0), 0);
              tp.totalBottles = tp.items.reduce((s, i) => s + (i.totalBtl || 0), 0);
              tp.totalAmount = tp.items.reduce((s, i) => s + (i.amount || 0), 0);
            }
            return tp;
          });
          // Remove TPs that have no items left
          tps = tps.filter(tp => !tp.items || tp.items.length > 0);
          if (tpChanged) {
            fs.writeFileSync(tpFile, JSON.stringify(tps, null, 2));
          }
        } catch (_) { /* ignore TP cleanup errors */ }
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ── IPC: TP (Transport Permit) CRUD ───────────────────────
ipcMain.handle('get-tps', async (event, { barName, financialYear }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'tp_entries.json');
    if (!fs.existsSync(filePath)) return { success: true, tps: [] };
    const raw = fs.readFileSync(filePath, 'utf8');
    return { success: true, tps: raw ? JSON.parse(raw) : [] };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('save-tp', async (event, { barName, financialYear, tp }) => {
  try {
    const dir = getBarDir(barName, financialYear);
    fs.mkdirSync(dir, { recursive: true });
    const filePath = path.join(dir, 'tp_entries.json');
    let tps = [];
    if (fs.existsSync(filePath)) {
      const raw = fs.readFileSync(filePath, 'utf8');
      if (raw) tps = JSON.parse(raw);
    }
    const idx = tps.findIndex(t => t.id === tp.id);
    if (idx >= 0) {
      tps[idx] = { ...tps[idx], ...tp, updatedAt: new Date().toISOString() };
    } else {
      tp.createdAt = new Date().toISOString();
      tps.push(tp);
    }
    fs.writeFileSync(filePath, JSON.stringify(tps, null, 2));
    return { success: true, tp };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('delete-tp', async (event, { barName, financialYear, tpId }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'tp_entries.json');
    if (!fs.existsSync(filePath)) return { success: false, error: 'No TP file' };
    let tps = JSON.parse(fs.readFileSync(filePath, 'utf8') || '[]');
    tps = tps.filter(t => t.id !== tpId);
    fs.writeFileSync(filePath, JSON.stringify(tps, null, 2));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('export-tp-summary', async (event, { summaryRows, detailRows, barName }) => {
  try {
    const { dialog } = require('electron');
    const safeName = (barName || 'Bar').replace(/[^a-zA-Z0-9_\- ]/g, '');
    const today = new Date().toISOString().slice(0, 10);
    const defaultName = `TP_Summary_${safeName}_${today}.xlsx`;

    const result = await dialog.showSaveDialog({
      title: 'Export TP Summary',
      defaultPath: defaultName,
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    });
    if (result.canceled || !result.filePath) return { success: false, error: 'Cancelled' };

    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.json_to_sheet(summaryRows);
    const ws2 = XLSX.utils.json_to_sheet(detailRows);

    ws1['!cols'] = [
      { wch: 4 }, { wch: 22 }, { wch: 12 }, { wch: 12 },
      { wch: 26 }, { wch: 6 }, { wch: 8 }, { wch: 8 },
      { wch: 10 }, { wch: 14 }, { wch: 14 }
    ];
    ws2['!cols'] = [
      { wch: 22 }, { wch: 12 }, { wch: 26 }, { wch: 4 },
      { wch: 30 }, { wch: 12 }, { wch: 8 }, { wch: 8 },
      { wch: 6 }, { wch: 8 }, { wch: 10 }, { wch: 10 }, { wch: 14 }
    ];

    XLSX.utils.book_append_sheet(wb, ws1, 'TP Summary');
    XLSX.utils.book_append_sheet(wb, ws2, 'TP Items Detail');
    XLSX.writeFile(wb, result.filePath);
    return { success: true, filePath: result.filePath };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ─── Export Sale Summary XLSX (SCM Portal format) ─────────
ipcMain.handle('export-sale-summary', async (event, params) => {
  try {
    const { dialog } = require('electron');
    const rows = params.rows || [];
    const barName = params.barName || 'Bar';
    const date = params.date || '';
    const safeName = barName.replace(/[^a-zA-Z0-9_\- ]/g, '');
    const safeDate = (date || new Date().toISOString().slice(0, 10)).replace(/[^0-9\-]/g, '');
    const defaultName = `Sale_${safeName}_${safeDate}.xlsx`;

    if (rows.length === 0) return { success: false, error: 'No data to export' };

    const result = await dialog.showSaveDialog({
      title: 'Export Sale Summary for SCM Portal',
      defaultPath: defaultName,
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    });
    if (result.canceled || !result.filePath) return { success: false, error: 'Cancelled' };

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [
      { wch: 12 }, // Sale Date
      { wch: 18 }, // Local Item Code
      { wch: 50 }, // Brand Name
      { wch: 18 }, // Size
      { wch: 16 }, // Quantity(Case)
      { wch: 22 }, // Quantity(Loose Bottle)
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    XLSX.writeFile(wb, result.filePath);
    return { success: true, filePath: result.filePath };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ─── Export Purchase Report XLSX ─────────────────────────
ipcMain.handle('export-purchase-report', async (event, params) => {
  try {
    const { dialog } = require('electron');
    const rows = params.rows || [];
    const barName = params.barName || 'Bar';
    const safeName = barName.replace(/[^a-zA-Z0-9_\- ]/g, '');
    const today = new Date().toISOString().slice(0, 10);
    const defaultName = `Purchase_Report_${safeName}_${today}.xlsx`;

    if (rows.length === 0) return { success: false, error: 'No data to export' };

    const result = await dialog.showSaveDialog({
      title: 'Export Purchase Report',
      defaultPath: defaultName,
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    });
    if (result.canceled || !result.filePath) return { success: false, error: 'Cancelled' };

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [
      { wch: 12 }, // TP Date
      { wch: 12 }, // Received Date
      { wch: 16 }, // TP No.
      { wch: 28 }, // Supplier
      { wch: 40 }, // Brand
      { wch: 14 }, // Code
      { wch: 14 }, // Size
      { wch:  8 }, // Cases
      { wch:  8 }, // Loose Btl
      { wch: 10 }, // Total Btls
      { wch: 10 }, // Rate
      { wch: 14 }, // Amount
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Purchase Report');
    XLSX.writeFile(wb, result.filePath);
    return { success: true, filePath: result.filePath };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ─── Import TP from SCD XLS ──────────────────────────────
ipcMain.handle('import-tp-xls', async (event, params = {}) => {
  try {
    const { barName, financialYear } = params;

    // Build product code → full product lookup from product master
    const productLookup = {};   // code.toUpperCase() → { size, mrp, id }
    let prodFilePath = null;
    let allProducts = [];
    if (barName) {
      try {
        prodFilePath = path.join(getBarDir(barName, financialYear || ''), 'products.json');
        if (fs.existsSync(prodFilePath)) {
          allProducts = JSON.parse(fs.readFileSync(prodFilePath, 'utf8') || '[]');
          for (const p of allProducts) {
            if (p.code) productLookup[p.code.trim().toUpperCase()] = { size: p.size || '', mrp: p.mrp || 0, id: p.id };
          }
        }
      } catch (_) { /* ignore product lookup errors */ }
    }

    const { dialog } = require('electron');
    const result = await dialog.showOpenDialog({
      title: 'Import TP from Maharashtra Excise XLS',
      filters: [
        { name: 'Excel / HTML Table', extensions: ['xls', 'xlsx', 'html', 'htm'] }
      ],
      properties: ['openFile'],
    });
    if (result.canceled || !result.filePaths.length) return { success: false, error: 'Cancelled' };

    const filePath = result.filePaths[0];
    const raw = fs.readFileSync(filePath, 'utf8');

    // --- Parse: split on </tr> and extract <td> contents ---
    const chunks = raw.split(/<\/tr>/i);
    const allRows = [];
    for (const ch of chunks) {
      const cells = [];
      const reg = /<td[^>]*>(.*?)<\/td>/gi;
      let m;
      while ((m = reg.exec(ch)) !== null) cells.push(m[1].trim());
      if (cells.length >= 20) allRows.push(cells);
    }

    if (allRows.length < 2) {
      return { success: false, error: 'Could not find enough data rows. Is this a Maharashtra Excise TP export?' };
    }

    // First row with 27 cols is the header; find column mapping
    const header = allRows[0];
    const colMap = {};
    header.forEach((h, i) => {
      const lc = h.toLowerCase().trim();
      if (lc.includes('received date')) colMap.receivedDate = i;
      else if (lc.includes('auto tp')) colMap.autoTpNo = i;
      else if (lc.includes('manual tp')) colMap.manualTpNo = i;
      else if (lc === 'tp date') colMap.tpDate = i;
      else if (lc.includes('party name')) colMap.partyName = i;
      else if (lc.includes('scm item code') || lc === 'item code') colMap.itemCode = i;
      else if (lc === 'item name') colMap.itemName = i;
      else if (lc === 'size') colMap.size = i;
      else if (lc.includes('qty (cases)') || lc === 'cases') colMap.cases = i;
      else if (lc.includes('qty (bottles)') || lc === 'bottles') colMap.bottles = i;
      else if (lc === 'mrp') colMap.mrp = i;
      else if (lc.includes('total bot')) colMap.totalBtl = i;
      else if (lc.includes('vehicle')) colMap.vehicle = i;
      else if (lc.includes('scm party code')) colMap.partyCode = i;
    });

    // Data rows (skip header)
    const dataRows = allRows.slice(1);
    if (dataRows.length === 0) {
      return { success: false, error: 'No data rows found in the file.' };
    }

    // Group rows by Manual TP No (prefer manual over auto)
    const tpGroups = {};
    for (const row of dataRows) {
      const tpKey = row[colMap.manualTpNo] || row[colMap.autoTpNo] || 'UNKNOWN';
      if (!tpGroups[tpKey]) tpGroups[tpKey] = [];
      tpGroups[tpKey].push(row);
    }

    // Parse size string from Excel: "90 ML(100)" / "2000 ML Pet (6)" / "180 ML" etc.
    // Returns the proper ALL_SIZES value string.
    const SIZE_VALUE_LOOKUP = [
      { ml: 50,   bpc: 120, value: '50 ML' },
      { ml: 50,   bpc: 192, value: '50 ML (192)' },
      { ml: 50,   bpc: 24,  value: '50 ML (24)' },
      { ml: 50,   bpc: 60,  value: '50 ML (60)' },
      { ml: 60,   bpc: 96,  value: '60 ML' },
      { ml: 60,   bpc: 120, value: '60 ML (120)' },
      { ml: 60,   bpc: 75,  value: '60 ML (75)' },
      { ml: 90,   bpc: 48,  value: '90 ML (48)' },
      { ml: 90,   bpc: 96,  value: '90 ML (96)' },
      { ml: 90,   bpc: 100, value: '90 ML (100)' },
      { ml: 125,  bpc: 72,  value: '125 ML' },
      { ml: 180,  bpc: 48,  value: '180 ML' },
      { ml: 180,  bpc: 24,  value: '180 ML (24)' },
      { ml: 180,  bpc: 50,  value: '180 ML (50)' },
      { ml: 187,  bpc: 48,  value: '187 ML (48)' },
      { ml: 200,  bpc: 12,  value: '200 ML (12)' },
      { ml: 200,  bpc: 24,  value: '200 ML (24)' },
      { ml: 200,  bpc: 30,  value: '200 ML (30)' },
      { ml: 200,  bpc: 48,  value: '200 ML (48)' },
      { ml: 250,  bpc: 24,  value: '250 ML' },
      { ml: 275,  bpc: 24,  value: '275 ML (24)' },
      { ml: 330,  bpc: 24,  value: '330 ML' },
      { ml: 330,  bpc: 12,  value: '330 ML (12)' },
      { ml: 350,  bpc: 12,  value: '350 ML (12)' },
      { ml: 375,  bpc: 24,  value: '375 ML' },
      { ml: 375,  bpc: 12,  value: '375 ML (12)' },
      { ml: 500,  bpc: 12,  value: '500 ML' },
      { ml: 500,  bpc: 24,  value: '500 ML (24)' },
      { ml: 650,  bpc: 12,  value: '650 ML' },
      { ml: 700,  bpc: 12,  value: '700 ML' },
      { ml: 700,  bpc: 6,   value: '700 ML (6)' },
      { ml: 750,  bpc: 12,  value: '750 ML' },
      { ml: 750,  bpc: 6,   value: '750 ML (6)' },
      { ml: 1000, bpc: 9,   value: '1000 ML' },
      { ml: 1000, bpc: 12,  value: '1000 ML (12)' },
      { ml: 1000, bpc: 6,   value: '1000 ML (6)' },
      { ml: 1500, bpc: 6,   value: '1500 ML' },
      { ml: 1750, bpc: 6,   value: '1750 ML (6)' },
      { ml: 2000, bpc: 6,   value: '2000 ML (6)' },
      { ml: 2000, bpc: 4,   value: '2000 ML (4)' },
      { ml: 4500, bpc: 4,   value: '4500 ML' },
    ];
    function sizeToValue(ml, bpc) {
      const mlNum = parseInt(ml) || 0;
      const matched = SIZE_VALUE_LOOKUP.find(e => e.ml === mlNum && e.bpc === bpc);
      if (matched) return matched.value;
      if (mlNum > 0) return mlNum + ' ML' + (bpc ? ' (' + bpc + ')' : '');
      return '';
    }

    function parseSize(str) {
      if (!str) return { ml: '', bpc: 0, value: '' };
      // Strip HTML tags if any
      const clean = str.replace(/<[^>]*>/g, '').trim();
      // Extract ML number (also handles "Ltr" sizes)
      const mlMatch = clean.match(/(\d+)\s*ML/i);
      const ml = mlMatch ? mlMatch[1] : '';
      // BPC: explicit (number) in parentheses takes priority (handles "Pet (6)", "(100)" etc.)
      // Find the LAST (digits) occurrence to handle "Pet (6)"
      const bpcMatches = [...clean.matchAll(/\((\d+)\)/g)];
      let bpc = 0;
      if (bpcMatches.length > 0) {
        bpc = parseInt(bpcMatches[bpcMatches.length - 1][1]);
      } else {
        // Standard Maharashtra BPC
        const stdBpc = { '50': 120, '60': 96, '90': 96, '125': 72, '180': 48, '375': 24, '500': 12, '650': 12, '750': 12, '1000': 9, '2000': 6 };
        bpc = stdBpc[ml] || 0;
      }
      return { ml, bpc, value: sizeToValue(ml, bpc) };
    }

    // Parse date: "14-Feb-2026" → "2026-02-14"
    function parseDate(str) {
      if (!str) return '';
      try {
        const d = new Date(str);
        if (isNaN(d.getTime())) return str;
        return d.toISOString().slice(0, 10);
      } catch { return str; }
    }

    // Build TP objects
    const newProductsMap = {};
    const mrpUpdatesMap = {}; // code.toUpperCase() → new MRP
    const tps = [];
    for (const [tpKey, rows] of Object.entries(tpGroups)) {
      const first = rows[0];
      const items = rows.map(row => {
        const itemCode = (row[colMap.itemCode] || '').trim();
        const excelMrp = parseFloat(row[colMap.mrp] || '0') || 0;
        // 1. Try to get size from product master by item code
        const prodEntry = itemCode ? (productLookup[itemCode.toUpperCase()] || null) : null;
        const prodSize  = prodEntry ? prodEntry.size : '';
        // Normalise product size through SIZE_VALUE_LOOKUP so it exactly matches ALL_SIZES values
        function normaliseProdSize(sz) {
          if (!sz) return '';
          // If it already exists as a SIZE_VALUE_LOOKUP value, return as-is
          if (SIZE_VALUE_LOOKUP.find(e => e.value === sz)) return sz;
          // Otherwise re-parse it to get the canonical value
          const m = sz.match(/(\d+)\s*ML/i);
          if (!m) return sz;
          const ml = m[1];
          const bpcM = [...sz.matchAll(/\((\d+)\)/g)];
          let bpc = bpcM.length > 0 ? parseInt(bpcM[bpcM.length - 1][1]) : 0;
          if (!bpc) {
            const stdBpc = { '50': 120, '60': 96, '90': 96, '125': 72, '180': 48, '375': 24, '500': 12, '650': 12, '750': 12, '1000': 9, '2000': 6 };
            bpc = stdBpc[ml] || 0;
          }
          return sizeToValue(ml, bpc) || sz;
        }
        const normProdSize = normaliseProdSize(prodSize);
        // 2. Fallback: parse size from Excel column
        const sizeInfo = parseSize(row[colMap.size] || '');
        const resolvedSize = normProdSize || sizeInfo.value;
        // Get bpc from SIZE_VALUE_LOOKUP for the resolved size
        const sizeEntry = SIZE_VALUE_LOOKUP.find(e => e.value === resolvedSize);
        const resolvedBpc = sizeEntry ? sizeEntry.bpc : sizeInfo.bpc;
        // Track new products (not in product master)
        if (itemCode && !prodEntry) {
          const brand = row[colMap.itemName] || '';
          if (!newProductsMap[itemCode.toUpperCase()]) {
            newProductsMap[itemCode.toUpperCase()] = { code: itemCode, brand, size: sizeInfo.value };
          }
        }
        // Track MRP updates needed (product exists but MRP differs)
        if (itemCode && prodEntry && excelMrp > 0 && excelMrp !== prodEntry.mrp) {
          mrpUpdatesMap[itemCode.toUpperCase()] = excelMrp;
        }
        const cases = parseFloat(row[colMap.cases] || '0') || 0;
        const loose = parseFloat(row[colMap.bottles] || '0') || 0;
        const totalBtl = parseFloat(row[colMap.totalBtl] || '0') || (cases * resolvedBpc + loose);
        const rate = excelMrp;
        return {
          brand: row[colMap.itemName] || '',
          code: itemCode,
          size: resolvedSize,
          sizeLabel: resolvedSize,
          cases: Math.round(cases),
          bpc: resolvedBpc,
          loose: Math.round(loose),
          totalBtl: Math.round(totalBtl),
          rate,
          amount: Math.round(totalBtl) * rate,
        };
      });

      tps.push({
        tpNumber: first[colMap.manualTpNo] || first[colMap.autoTpNo] || '',
        manualTpNo: first[colMap.manualTpNo] || '',
        tpDate: parseDate(first[colMap.tpDate]),
        receivedDate: parseDate(first[colMap.receivedDate]),
        supplier: first[colMap.partyName] || '',
        supplierCode: first[colMap.partyCode] || '',
        vehicle: (first[colMap.vehicle] || '').replace(/\./g, ''),
        items,
        totalCases: items.reduce((s, i) => s + i.cases, 0),
        totalLoose: items.reduce((s, i) => s + i.loose, 0),
        totalBottles: items.reduce((s, i) => s + i.totalBtl, 0),
        totalAmount: items.reduce((s, i) => s + i.amount, 0),
      });
    }

    // Apply MRP updates to products.json
    let mrpUpdatedCount = 0;
    if (prodFilePath && allProducts.length > 0 && Object.keys(mrpUpdatesMap).length > 0) {
      try {
        let changed = false;
        for (const p of allProducts) {
          const key = (p.code || '').trim().toUpperCase();
          if (mrpUpdatesMap[key] !== undefined) {
            p.mrp = mrpUpdatesMap[key];
            p.updatedAt = new Date().toISOString();
            changed = true;
            mrpUpdatedCount++;
          }
        }
        if (changed) fs.writeFileSync(prodFilePath, JSON.stringify(allProducts, null, 2));
      } catch (_) { /* ignore MRP update errors */ }
    }

    return { success: true, tps, newProducts: Object.values(newProductsMap), mrpUpdated: mrpUpdatedCount, fileName: path.basename(filePath) };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ─── Product Bulk Import ─────────────────────────────────
const XLSX = require('xlsx');

ipcMain.handle('download-product-template', async (event) => {
  try {
    const { dialog } = require('electron');
    const templateHeaders = [
      'Brand Name *', 'Code', 'Category *', 'Sub-Category',
      'Size (ml) *', 'BPC', 'MRP', 'Cost Price',
      'HSN', 'Supplier', 'Barcode', 'Remarks'
    ];
    const sampleRows = [
      ['Royal Stag', 'RS750', 'Spirits', 'Whisky', '750 ML', '12', '650', '520', '22041000', 'ABC Distributors', '', 'Sample row'],
      ['Kingfisher Premium', 'KFP650', 'Fermented Beer', 'Beer', '650 ML', '12', '180', '150', '22030000', '', '', ''],
      ['Mansion House', 'MH180', 'MML', 'Brandy', '180 ML', '48', '120', '90', '', '', '', ''],
    ];
    const wsData = [templateHeaders, ...sampleRows];
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'Product Template');
    // Set column widths
    ws['!cols'] = [
      { wch: 22 }, { wch: 12 }, { wch: 16 }, { wch: 14 },
      { wch: 10 }, { wch: 6 },  { wch: 10 }, { wch: 10 },
      { wch: 12 }, { wch: 20 }, { wch: 14 }, { wch: 20 },
    ];
    const result = await dialog.showSaveDialog({
      title: 'Save Product Import Template',
      defaultPath: 'Product_Import_Template.xlsx',
      filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
    });
    if (result.canceled || !result.filePath) return { success: false, error: 'Cancelled' };
    XLSX.writeFile(wb, result.filePath);
    return { success: true, filePath: result.filePath };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('import-products-excel', async (event, { barName, financialYear }) => {
  try {
    const { dialog } = require('electron');
    const result = await dialog.showOpenDialog({
      title: 'Select Product Excel File',
      filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls', 'csv'] }],
      properties: ['openFile'],
    });
    if (result.canceled || !result.filePaths.length) return { success: false, error: 'Cancelled' };

    const filePath = result.filePaths[0];
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

    if (rows.length === 0) return { success: false, error: 'No data rows found in the file' };

    // Map Excel headers to product fields
    const headerMap = {
      'brand name *': 'brandName', 'brand name': 'brandName', 'brandname': 'brandName', 'brand': 'brandName',
      'code': 'code', 'product code': 'code',
      'category *': 'category', 'category': 'category',
      'sub-category': 'subCategory', 'subcategory': 'subCategory', 'sub category': 'subCategory',
      'size (ml) *': 'size', 'size (ml)': 'size', 'size': 'size', 'sizeml': 'size',
      'bpc': 'bpc', 'btl/cs': 'bpc', 'bottles per case': 'bpc',
      'mrp': 'mrp', 'mrp *': 'mrp',
      'cost price': 'costPrice', 'costprice': 'costPrice', 'cost': 'costPrice',
      'hsn': 'hsn', 'hsn code': 'hsn',
      'supplier': 'supplier',
      'barcode': 'barcode',
      'remarks': 'remarks', 'notes': 'remarks',
    };

    const SIZE_BPC_MAP = { '1000': 9, '750': 12, '650': 12, '500': 12, '375': 24, '330': 24, '180': 48, '90': 96 };

    // ── Auto-detect sub-category from brand name keywords ──
    const SUBCATEGORY_KEYWORDS = [
      // Whisky brands
      { keywords: ['black dog', 'blenders pride', 'royal stag', 'imperial blue', 'mcdowell', 'signature', 'antiquity',
        'teachers', 'vat 69', 'black & white', 'jack daniel', 'jim beam', 'johnnie walker', 'ballantine',
        '100 pipers', 'bagpiper', 'peter scot', 'officer', 'old monk', 'hayward', 'oaksmith',
        'royal challenge', 'sixer', 'rockford', 'mansion house', 'director', 'gold riband',
        'chivas', 'dewars', 'famous grouse', 'grants', 'glenfiddich', 'glenlivet', 'monkey shoulder',
        'jameson', 'bushmills', 'wild turkey', 'makers mark', 'woodford', 'buffalo trace',
        'amrut', 'paul john', 'rampur', 'indri', 'gianchand', 'godawan',
        'morpheus', 'oak & cane', 'sterling reserve', 'double black', 'red label', 'black label',
        'gold label', 'green label', 'blue label', 'platinum label',
        'windsor', 'something special', 'grand master', 'after dark', 'all seasons', '8pm',
        'whisky', 'whiskey', 'scotch', 'bourbon', 'single malt', 'blended'], sub: 'Whisky' },
      // Rum brands
      { keywords: ['old monk', 'bacardi', 'captain morgan', 'malibu', 'old port', 'hercules',
        'havana club', 'ron', 'mount gay', 'appleton', 'sailor jerry', 'kraken',
        'magician', 'contessa', 'jolly roger', 'xxx rum', 'mcginns', 'bermuda',
        'campa rum', 'romanov rum', 'white mischief rum', 'cabo',
        'rum', 'dark rum', 'white rum', 'gold rum', 'spiced rum'], sub: 'Rum' },
      // Vodka brands
      { keywords: ['smirnoff', 'absolut', 'grey goose', 'belvedere', 'ciroc', 'ketel one',
        'stolichnaya', 'finlandia', 'skyy', 'titos', 'romanov', 'magic moments',
        'white mischief', 'fuel', 'vladmir', 'zippy', 'happy hours',
        'ping', 'nucleus', 'morpheus xo', 'smoke', 'eriksson',
        'vodka'], sub: 'Vodka' },
      // Gin brands
      { keywords: ['bombay sapphire', 'tanqueray', 'hendricks', 'beefeater', 'gordon',
        'blue riband', 'greater than', 'hapusa', 'jaisalmer', 'stranger & sons',
        'roku', 'monkey 47', 'the botanist', 'jin jiji', 'tickle',
        'gin'], sub: 'Gin' },
      // Brandy brands
      { keywords: ['mansion house brandy', 'honey bee', 'mcginns brandy', 'old admiral',
        'doctor brandy', 'dreher', 'courrier napoleon', 'morpheus',
        'three barrels', 'remy martin', 'hennessy', 'courvoisier', 'martell',
        'brandy', 'cognac', 'vsop', 'xo brandy'], sub: 'Brandy' },
      // Beer brands
      { keywords: ['kingfisher', 'tuborg', 'carlsberg', 'heineken', 'budweiser', 'corona',
        'stella artois', 'hoegaarden', 'foster', 'haywards', 'knockout', 'kalyani',
        'godfather', 'bira', 'simba', 'white owl', 'medusa', 'goa kings',
        'royal challenge beer', 'miller', 'amstel', 'peroni', 'leffe', 'grimbergen',
        'london pilsner', 'zingaro', 'bull', 'thunderbolt', 'nine hills', 'six fields',
        'beer', 'lager', 'ale', 'stout', 'pilsner', 'wheat beer', 'ipa', 'craft beer'], sub: 'Beer' },
      // Wine brands
      { keywords: ['sula', 'fratelli', 'grover zampa', 'york', 'four seasons', 'reveilo',
        'charosa', 'soma', 'vallonne', 'myra', 'good earth', 'big banyan',
        'nandi hills', 'chandon', 'moet', 'dom perignon', 'veuve clicquot',
        'jacob', 'yellow tail', 'barefoot', 'robert mondavi', 'casillero',
        'wine', 'champagne', 'prosecco', 'cabernet', 'merlot', 'shiraz',
        'sauvignon', 'chardonnay', 'pinot', 'rose', 'sangria', 'port wine'], sub: 'Wine' },
      // Tequila brands
      { keywords: ['jose cuervo', 'patron', 'don julio', 'herradura', 'sauza', 'olmeca',
        'casamigos', 'clase azul', 'el jimador', 'milagro',
        'tequila', 'mezcal'], sub: 'Tequila' },
      // Liqueur brands
      { keywords: ['baileys', 'kahlua', 'cointreau', 'grand marnier', 'amaretto',
        'jagermeister', 'campari', 'aperol', 'sambuca', 'absinthe', 'drambuie',
        'chambord', 'chartreuse', 'disaronno', 'limoncello', 'triple sec',
        'liqueur', 'cream liqueur', 'schnapps'], sub: 'Liqueur' },
      // Desi Daru / Country Liquor
      { keywords: ['santra', 'gavran', 'desi daru', 'country liquor', 'tadi', 'mahua',
        'feni', 'cashew feni', 'coconut feni', 'arrack', 'toddy'], sub: 'Desi Daru' },
    ];

    function guessSubCategory(brandName) {
      if (!brandName) return '';
      const lower = brandName.toLowerCase();
      for (const entry of SUBCATEGORY_KEYWORDS) {
        for (const kw of entry.keywords) {
          if (lower.includes(kw)) return entry.sub;
        }
      }
      return '';
    }

    // Load existing products
    const prodFilePath = path.join(getBarDir(barName, financialYear), 'products.json');
    let existingProducts = [];
    if (fs.existsSync(prodFilePath)) {
      const raw = fs.readFileSync(prodFilePath, 'utf8');
      if (raw) existingProducts = JSON.parse(raw);
    }

    const imported = [];
    const skipped = [];
    const errors = [];

    rows.forEach((row, rowIdx) => {
      // Normalize headers
      const mapped = {};
      Object.keys(row).forEach(key => {
        const normalKey = key.trim().toLowerCase();
        const field = headerMap[normalKey];
        if (field) mapped[field] = String(row[key]).trim();
      });

      // Validate required fields
      if (!mapped.brandName) {
        skipped.push({ row: rowIdx + 2, reason: 'Missing brand name' });
        return;
      }
      if (!mapped.category) {
        skipped.push({ row: rowIdx + 2, reason: 'Missing category' });
        return;
      }
      if (!mapped.size) {
        skipped.push({ row: rowIdx + 2, reason: 'Missing size' });
        return;
      }

      // ── Normalize size to match ALL_SIZES format ──
      const rawSize = String(mapped.size).trim();
      // Extract numeric ML portion
      const mlMatch = rawSize.match(/^(\d+)/);
      const mlNum = mlMatch ? mlMatch[1] : '';
      // Try to find exact match with ALL_SIZES values (case-insensitive)
      const ALL_SIZE_VALUES = [
        '50 ML','50 ML (192)','50 ML (24)','50 ML (60)','50 ML (120)',
        '60 ML','60 ML (120)','60 ML (75)','60 ML (Pet)',
        '90 ML (48)','90 ML (96)','90 ML (100)','90 ML (Pet)-96','90 ML (Pet)-100',
        '125 ML','125 ML (72)',
        '180 ML','180 ML (Pet)','180 ML (tetra)','180 ML (24)','180 ML (50)',
        '187 ML (48)',
        '200 ML (12)','200 ML (24)','200 ML (30)','200 ML (48)',
        '250 ML','250 ML (CAN)','250 ML (24)',
        '275 ML (24)',
        '330 ML','330 ML (CAN)','330 ML (12)',
        '350 ML (12)',
        '375 ML','375 ML (12)','375 ML (Pet)',
        '500 ML','500 ML (24)','500 ML (CAN)',
        '650 ML',
        '700 ML','700 ML (6)',
        '750 ML','750 ML (Pet)','750 ML (6)',
        '1000 ML','1000 ML (12)','1000 ML (6)','1000 ML (Pet)',
        '1500 ML','1750 ML (6)',
        '2000 ML (4)','2000 ML (6)','2000 ML Pet (6)','2000 ML (Pet)-4',
        '4500 ML',
        '15 Ltr','20 Ltr','30 Ltr','50 Ltr',
      ];
      let sizeVal = '';
      // 1) Try exact match
      const exactMatch = ALL_SIZE_VALUES.find(v => v.toLowerCase() === rawSize.toLowerCase());
      if (exactMatch) {
        sizeVal = exactMatch;
      } else {
        // 2) Try matching with BPC hint: e.g. "750" with bpc=12 → "750 ML"
        const bpcHint = parseInt(mapped.bpc) || 0;
        if (mlNum) {
          // Find candidates with this ML
          const candidates = ALL_SIZE_VALUES.filter(v => {
            const m = v.match(/^(\d+)/);
            return m && m[1] === mlNum;
          });
          if (candidates.length === 1) {
            sizeVal = candidates[0];
          } else if (candidates.length > 1 && bpcHint > 0) {
            // Match by BPC in brackets
            const bpcMatch = candidates.find(v => {
              const b = v.match(/\((\d+)\)/);
              return b && parseInt(b[1]) === bpcHint;
            });
            sizeVal = bpcMatch || candidates[0]; // fallback to first candidate
          } else if (candidates.length > 1) {
            // Pick the plain one (no brackets) or first
            const plain = candidates.find(v => !v.includes('('));
            sizeVal = plain || candidates[0];
          } else {
            sizeVal = rawSize; // keep as-is if no match
          }
        } else {
          sizeVal = rawSize;
        }
      }
      const sizeNumeric = mlNum || rawSize.replace(/[^0-9]/g, '');
      const bpc = parseInt(mapped.bpc) || SIZE_BPC_MAP[sizeNumeric] || 0;

      // Check for duplicate (same brand + size combo)
      const isDupe = existingProducts.some(p =>
        p.brandName && p.brandName.toUpperCase() === mapped.brandName.toUpperCase() &&
        String(p.size).toLowerCase() === sizeVal.toLowerCase()
      );
      if (isDupe) {
        skipped.push({ row: rowIdx + 2, reason: `Duplicate: ${mapped.brandName} ${sizeVal} already exists` });
        return;
      }

      const product = {
        id: 'PROD_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6) + '_' + rowIdx,
        brandName: mapped.brandName,
        code: (mapped.code || '').toUpperCase(),
        category: mapped.category,
        subCategory: mapped.subCategory || guessSubCategory(mapped.brandName),
        size: sizeVal,
        bpc: bpc,
        mrp: parseFloat(mapped.mrp) || 0,
        costPrice: parseFloat(mapped.costPrice) || 0,
        hsn: mapped.hsn || '',
        supplier: mapped.supplier || '',
        barcode: mapped.barcode || '',
        remarks: mapped.remarks || '',
        createdAt: new Date().toISOString(),
        importedFrom: path.basename(filePath),
      };

      existingProducts.push(product);
      imported.push(product);
    });

    // Save updated products
    const dir = getBarDir(barName, financialYear);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, 'products.json'), JSON.stringify(existingProducts, null, 2));

    return {
      success: true,
      imported: imported.length,
      skipped: skipped.length,
      skippedDetails: skipped,
      total: rows.length,
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ─── Opening Stock ───────────────────────────────────────
ipcMain.handle('get-opening-stock', async (event, { barName, financialYear }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'opening_stock.json');
    if (!fs.existsSync(filePath)) return { success: true, data: null };
    const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    return { success: true, data };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('save-opening-stock', async (event, { barName, financialYear, openingStock }) => {
  try {
    const dirPath = getBarDir(barName, financialYear);
    if (!fs.existsSync(dirPath)) fs.mkdirSync(dirPath, { recursive: true });
    const filePath = path.join(dirPath, 'opening_stock.json');
    fs.writeFileSync(filePath, JSON.stringify(openingStock, null, 2));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ─── MRP Master ──────────────────────────────────────────
ipcMain.handle('get-mrp-master', async (event, { barName, financialYear }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'mrp_master.json');
    if (!fs.existsSync(filePath)) return { success: true, entries: [] };
    const raw = fs.readFileSync(filePath, 'utf8');
    return { success: true, entries: raw ? JSON.parse(raw) : [] };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('save-mrp-entry', async (event, { barName, financialYear, mrpEntry }) => {
  try {
    const dir = getBarDir(barName, financialYear);
    fs.mkdirSync(dir, { recursive: true });
    const filePath = path.join(dir, 'mrp_master.json');
    let entries = [];
    if (fs.existsSync(filePath)) {
      const raw = fs.readFileSync(filePath, 'utf8');
      if (raw) entries = JSON.parse(raw);
    }
    const idx = entries.findIndex(e => e.id === mrpEntry.id);
    if (idx >= 0) {
      entries[idx] = { ...entries[idx], ...mrpEntry, updatedAt: new Date().toISOString() };
    } else {
      mrpEntry.createdAt = new Date().toISOString();
      entries.push(mrpEntry);
    }
    fs.writeFileSync(filePath, JSON.stringify(entries, null, 2));
    return { success: true, mrpEntry };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('delete-mrp-entry', async (event, { barName, financialYear, mrpEntryId }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'mrp_master.json');
    if (!fs.existsSync(filePath)) return { success: false, error: 'No MRP file' };
    let entries = JSON.parse(fs.readFileSync(filePath, 'utf8') || '[]');
    entries = entries.filter(e => e.id !== mrpEntryId);
    fs.writeFileSync(filePath, JSON.stringify(entries, null, 2));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ─── Shortcut Master ─────────────────────────────────────
ipcMain.handle('get-shortcuts', async (event, { barName, financialYear }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'shortcuts.json');
    if (!fs.existsSync(filePath)) return { success: true, shortcuts: [] };
    const raw = fs.readFileSync(filePath, 'utf8');
    return { success: true, shortcuts: raw ? JSON.parse(raw) : [] };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('save-shortcut', async (event, { barName, financialYear, shortcut }) => {
  try {
    const dir = getBarDir(barName, financialYear);
    fs.mkdirSync(dir, { recursive: true });
    const filePath = path.join(dir, 'shortcuts.json');
    let shortcuts = [];
    if (fs.existsSync(filePath)) {
      const raw = fs.readFileSync(filePath, 'utf8');
      if (raw) shortcuts = JSON.parse(raw);
    }
    const idx = shortcuts.findIndex(s => s.id === shortcut.id);
    if (idx >= 0) {
      shortcuts[idx] = { ...shortcuts[idx], ...shortcut, updatedAt: new Date().toISOString() };
    } else {
      shortcut.createdAt = new Date().toISOString();
      shortcuts.push(shortcut);
    }
    fs.writeFileSync(filePath, JSON.stringify(shortcuts, null, 2));
    return { success: true, shortcut };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('delete-shortcut', async (event, { barName, financialYear, shortcutId }) => {
  try {
    const filePath = path.join(getBarDir(barName, financialYear), 'shortcuts.json');
    if (!fs.existsSync(filePath)) return { success: false, error: 'No shortcuts file' };
    let shortcuts = JSON.parse(fs.readFileSync(filePath, 'utf8') || '[]');
    shortcuts = shortcuts.filter(s => s.id !== shortcutId);
    fs.writeFileSync(filePath, JSON.stringify(shortcuts, null, 2));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('save-all-shortcuts', async (event, { barName, financialYear, shortcuts }) => {
  try {
    const dir = getBarDir(barName, financialYear);
    fs.mkdirSync(dir, { recursive: true });
    const filePath = path.join(dir, 'shortcuts.json');
    fs.writeFileSync(filePath, JSON.stringify(shortcuts, null, 2));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ─── Data Root Management IPC ─────────────────────────────
ipcMain.handle('get-data-root', async () => {
  return { success: true, dataRoot: getDataRoot() };
});

ipcMain.handle('choose-data-folder', async () => {
  try {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: 'Choose Data Folder for SpliqourPro',
      properties: ['openDirectory', 'createDirectory'],
      defaultPath: getDataRoot()
    });
    if (result.canceled || !result.filePaths.length) {
      return { success: false, error: 'Cancelled' };
    }
    const newRoot = result.filePaths[0];
    const cfg = loadConfig();
    cfg.dataRoot = newRoot;
    _cfgCache = null;          // clear cache so next call re-reads
    saveConfig(cfg);
    fs.mkdirSync(newRoot, { recursive: true });
    return { success: true, dataRoot: newRoot };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('get-app-version', async () => {
  return { version: app.getVersion() };
});

ipcMain.handle('open-data-folder', async () => {
  try {
    await shell.openPath(getDataRoot());
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ── IPC: Check for Updates (lightweight GitHub releases check) ──
ipcMain.handle('check-for-updates', async () => {
  const current = app.getVersion();
  try {
    const result = await autoUpdater.checkForUpdates();
    if (!result || !result.updateInfo) {
      return { success: true, currentVersion: current, updateAvailable: false, message: `You are running the latest version (v${current}).` };
    }
    const latest = result.updateInfo.version;
    const updateAvailable = latest !== current;
    return {
      success: true,
      currentVersion: current,
      latestVersion: latest,
      updateAvailable,
      message: updateAvailable
        ? `Update available: v${latest} is ready to download.`
        : `You are running the latest version (v${current}).`
    };
  } catch (err) {
    return { success: true, currentVersion: current, updateAvailable: false, message: 'Could not check for updates: ' + err.message };
  }
});

ipcMain.handle('download-update', async () => {
  try {
    autoUpdater.downloadUpdate();
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
});

ipcMain.handle('install-update', () => {
  autoUpdater.quitAndInstall();
});
