const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    // Window state
    onWindowStateChanged: (cb) => ipcRenderer.on('window-state-changed', (_e, state) => cb(state)),
    setTitleBarOverlay: (opts) => ipcRenderer.invoke('set-title-bar-overlay', opts),
    // Custom window controls
    windowMinimize: () => ipcRenderer.send('window-minimize'),
    windowMaximize: () => ipcRenderer.send('window-maximize'),
    windowClose:    () => ipcRenderer.send('window-close'),
    // Bar data
    saveBarData: (data) => ipcRenderer.invoke('save-bar-data', data),
    getBarsIndex: () => ipcRenderer.invoke('get-bars-index'),
    openBar: (entry) => ipcRenderer.invoke('open-bar', entry),
    // Navigation
    navigateToAddBar: () => ipcRenderer.send('navigate-to-add-bar'),
    navigateToApp: () => ipcRenderer.send('navigate-to-app'),
    navigateHome: () => ipcRenderer.send('navigate-home'),
    // Suppliers
    getSuppliers: (params) => ipcRenderer.invoke('get-suppliers', params),
    saveSupplier: (params) => ipcRenderer.invoke('save-supplier', params),
    deleteSupplier: (params) => ipcRenderer.invoke('delete-supplier', params),
    // Customers
    getCustomers: (params) => ipcRenderer.invoke('get-customers', params),
    saveCustomer: (params) => ipcRenderer.invoke('save-customer', params),
    deleteCustomer: (params) => ipcRenderer.invoke('delete-customer', params),
    // Products
    getProducts: (params) => ipcRenderer.invoke('get-products', params),
    saveProduct: (params) => ipcRenderer.invoke('save-product', params),
    deleteProduct: (params) => ipcRenderer.invoke('delete-product', params),
    downloadProductTemplate: () => ipcRenderer.invoke('download-product-template'),
    importProductsExcel: (params) => ipcRenderer.invoke('import-products-excel', params),
    // Transport Permits
    getTps: (params) => ipcRenderer.invoke('get-tps', params),
    saveTp: (params) => ipcRenderer.invoke('save-tp', params),
    deleteTp: (params) => ipcRenderer.invoke('delete-tp', params),
    exportTpSummary: (params) => ipcRenderer.invoke('export-tp-summary', params),
    importTpXls: (params) => ipcRenderer.invoke('import-tp-xls', params),
    exportPurchaseReport: (params) => ipcRenderer.invoke('export-purchase-report', params),
    // Opening Stock
    getOpeningStock: (params) => ipcRenderer.invoke('get-opening-stock', params),
    saveOpeningStock: (params) => ipcRenderer.invoke('save-opening-stock', params),
    // MRP Master
    getMrpMaster: (params) => ipcRenderer.invoke('get-mrp-master', params),
    saveMrpEntry: (params) => ipcRenderer.invoke('save-mrp-entry', params),
    deleteMrpEntry: (params) => ipcRenderer.invoke('delete-mrp-entry', params),
    // Shortcut Master
    getShortcuts: (params) => ipcRenderer.invoke('get-shortcuts', params),
    saveShortcut: (params) => ipcRenderer.invoke('save-shortcut', params),
    deleteShortcut: (params) => ipcRenderer.invoke('delete-shortcut', params),
    saveAllShortcuts: (params) => ipcRenderer.invoke('save-all-shortcuts', params),
    // Daily Sales
    getDailySales: (params) => ipcRenderer.invoke('get-daily-sales', params),
    saveDailySale: (params) => ipcRenderer.invoke('save-daily-sale', params),
    deleteDailySale: (params) => ipcRenderer.invoke('delete-daily-sale', params),
    getSaleCounter: (params) => ipcRenderer.invoke('get-sale-counter', params),
    saveSaleCounter: (params) => ipcRenderer.invoke('save-sale-counter', params),
    getCurrentStock: (params) => ipcRenderer.invoke('get-current-stock', params),
    exportSaleSummary: (params) => ipcRenderer.invoke('export-sale-summary', params),
    // Open external URL
    openExternal: (url) => ipcRenderer.invoke('open-external-url', url),
    // Auto-login in new window
    openUrlAutoLogin: (params) => ipcRenderer.invoke('open-url-autologin', params),
    // Backup & Transfer
    backupData: (params) => ipcRenderer.invoke('backup-data', params),
    transferToNextFY: (params) => ipcRenderer.invoke('transfer-to-next-fy', params),
    // Data Root / Settings
    getDataRoot: () => ipcRenderer.invoke('get-data-root'),
    chooseDataFolder: () => ipcRenderer.invoke('choose-data-folder'),
    openDataFolder: () => ipcRenderer.invoke('open-data-folder'),
    getAppVersion: () => ipcRenderer.invoke('get-app-version'),
    checkForUpdates: () => ipcRenderer.invoke('check-for-updates'),
});
