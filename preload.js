const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    // Authentication
    loginWithGoogle: () => ipcRenderer.send('login-with-google'),
    submitAuthCode: (code) => ipcRenderer.send('submit-auth-code', code),
    logout: () => ipcRenderer.send('logout'),
    
    // Spreadsheet operations
    testSpreadsheetAccess: (fileId) => ipcRenderer.send('test-spreadsheet-access', fileId),
    processSpreadsheet: (fileId) => ipcRenderer.send('process-spreadsheet', fileId),
    saveSpreadsheetId: (spreadsheetId) => ipcRenderer.send('save-spreadsheet-id', spreadsheetId),
    
    // Event listeners for responses from main process
    receiveGoogleAuthSuccess: (callback) => ipcRenderer.on('google-auth-success', (event, message) => callback(message)),
    receiveGoogleAuthError: (callback) => ipcRenderer.on('google-auth-error', (event, message) => callback(message)),
    
    receiveTestAccessSuccess: (callback) => ipcRenderer.on('test-access-success', (event, data) => callback(data)),
    receiveTestAccessError: (callback) => ipcRenderer.on('test-access-error', (event, message) => callback(message)),
    receiveTestAccessExcelDetected: (callback) => ipcRenderer.on('test-access-excel-detected', (event, data) => callback(data)),
    
    receiveProcessingComplete: (callback) => ipcRenderer.on('processing-complete', (event, message) => callback(message)),
    receiveProcessingError: (callback) => ipcRenderer.on('processing-error', (event, message) => callback(message)),
    
    // Session management event listeners
    receiveRestoreSession: (callback) => ipcRenderer.on('restore-session', (event, data) => callback(data)),
    receiveLogoutComplete: (callback) => ipcRenderer.on('logout-complete', (event) => callback()),
    receiveAuthExpired: (callback) => ipcRenderer.on('auth-expired', (event) => callback()),
});