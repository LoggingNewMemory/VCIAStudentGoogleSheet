const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    // --- Main Actions ---
    loginWithGoogle: () => ipcRenderer.send('login-with-google'),
    submitAuthCode: (code) => ipcRenderer.send('submit-auth-code', code),
    logout: () => ipcRenderer.send('logout'),
    saveSpreadsheetId: (spreadsheetId) => ipcRenderer.send('save-spreadsheet-id', spreadsheetId),
    testSpreadsheetAccess: (fileId) => ipcRenderer.send('test-spreadsheet-access', fileId),
    processSpreadsheet: (fileId) => ipcRenderer.send('process-spreadsheet', fileId),
    revertLastChange: () => ipcRenderer.send('revert-last-change'),

    // --- Event Listeners from Main to Renderer ---
    // Auth
    receiveRestoreSession: (callback) => ipcRenderer.on('restore-session', (event, data) => callback(data)),
    receiveGoogleAuthSuccess: (callback) => ipcRenderer.on('google-auth-success', (event, message) => callback(message)),
    receiveGoogleAuthError: (callback) => ipcRenderer.on('google-auth-error', (event, message) => callback(message)),
    receiveLogoutComplete: (callback) => ipcRenderer.on('logout-complete', (event) => callback()),
    receiveAuthExpired: (callback) => ipcRenderer.on('auth-expired', (event) => callback()),
    
    // Spreadsheet Test
    receiveTestAccessSuccess: (callback) => ipcRenderer.on('test-access-success', (event, data) => callback(data)),
    receiveTestAccessError: (callback) => ipcRenderer.on('test-access-error', (event, message) => callback(message)),
    receiveTestAccessExcelDetected: (callback) => ipcRenderer.on('test-access-excel-detected', (event) => callback()),

    // Spreadsheet Process
    receiveProcessingComplete: (callback) => ipcRenderer.on('processing-complete', (event, message) => callback(message)),
    receiveProcessingError: (callback) => ipcRenderer.on('processing-error', (event, message) => callback(message)),

    // Revert
    receiveRevertComplete: (callback) => ipcRenderer.on('revert-complete', (event, message) => callback(message)),
    receiveRevertError: (callback) => ipcRenderer.on('revert-error', (event, message) => callback(message)),
});