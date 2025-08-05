const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { google } = require('googleapis');

// Global references
let mainWindow;
let authWindow;
let lastSuccessfulMoves = []; // For the revert feature

// File paths for storing persistent data
const CONFIG_FILE = path.join(app.getPath('userData'), 'config.json');
const TOKENS_FILE = path.join(app.getPath('userData'), 'tokens.json');

// --- Google Auth Setup ---
let credentials;
try {
    const content = fs.readFileSync('client_secret.json');
    credentials = JSON.parse(content).installed;
} catch (err) {
    console.error('Error loading client_secret.json:', err);
    app.quit();
}
const { client_secret, client_id } = credentials;
const REDIRECT_URI = 'urn:ietf:wg:oauth:2.0:oob';
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, REDIRECT_URI);

// --- Config & Token Management (Helper Functions) ---
function loadConfig() {
    try {
        if (fs.existsSync(CONFIG_FILE)) {
            return JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8'));
        }
    } catch (error) { console.error('Error loading config:', error); }
    return { spreadsheetId: '' };
}

function saveConfig(config) {
    try {
        fs.writeFileSync(CONFIG_FILE, JSON.stringify(config, null, 2));
    } catch (error) { console.error('Error saving config:', error); }
}

function loadTokens() {
    try {
        if (fs.existsSync(TOKENS_FILE)) {
            const tokens = JSON.parse(fs.readFileSync(TOKENS_FILE, 'utf8'));
            if (tokens.expiry_date && tokens.expiry_date <= Date.now()) {
                if (tokens.refresh_token) return 'refresh_needed';
                return false;
            }
            oAuth2Client.setCredentials(tokens);
            return true;
        }
    } catch (error) { console.error('Error loading tokens:', error); }
    return false;
}

function saveTokens(tokens) {
    try {
        fs.writeFileSync(TOKENS_FILE, JSON.stringify(tokens, null, 2));
    } catch (error) { console.error('Error saving tokens:', error); }
}

function clearTokens() {
    try {
        if (fs.existsSync(TOKENS_FILE)) fs.unlinkSync(TOKENS_FILE);
        lastSuccessfulMoves = []; // Clear undo history on logout
    } catch (error) { console.error('Error clearing tokens:', error); }
}

async function refreshAccessToken() {
    try {
        const { credentials } = await oAuth2Client.refreshAccessToken();
        oAuth2Client.setCredentials(credentials);
        saveTokens(credentials);
        return true;
    } catch (error) {
        console.error('Error refreshing token:', error);
        clearTokens();
        return false;
    }
}

// --- Electron Window Management ---
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 850,
    height: 700,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      enableRemoteModule: false,
      nodeIntegration: false,
    },
  });
  mainWindow.loadFile('index.html');
  mainWindow.on('closed', () => { mainWindow = null; });

  mainWindow.webContents.once('did-finish-load', async () => {
    const config = loadConfig();
    const tokenStatus = loadTokens();
    let authenticated = false;
    if (tokenStatus === true) {
      authenticated = true;
    } else if (tokenStatus === 'refresh_needed') {
      authenticated = await refreshAccessToken();
    }
    mainWindow.webContents.send('restore-session', { authenticated, spreadsheetId: config.spreadsheetId || '' });
  });
}

app.on('ready', createWindow);
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
app.on('activate', () => { if (mainWindow === null) createWindow(); });

// --- Google API Helper ---
function getErrorMessage(error) {
    if (error.code === 404) return 'File not found. Please check the Spreadsheet ID.';
    if (error.code === 403) return 'Access denied. Ensure the file is shared with your Google account and you have editor permissions.';
    if (error.code === 401) {
        clearTokens();
        mainWindow.webContents.send('auth-expired');
        return 'Authentication expired. Please log in again.';
    }
    return error.message || 'An unknown error occurred.';
}

// --- IPC Handlers ---
ipcMain.on('login-with-google', () => {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        prompt: 'consent', // Important to get a refresh token every time
        scope: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly'],
    });
    authWindow = new BrowserWindow({ width: 600, height: 800 });
    authWindow.loadURL(authUrl);
    authWindow.on('closed', () => { authWindow = null });
});

ipcMain.on('submit-auth-code', async (event, code) => {
    if (authWindow) authWindow.close();
    try {
        const { tokens } = await oAuth2Client.getToken(code);
        oAuth2Client.setCredentials(tokens);
        saveTokens(tokens);
        mainWindow.webContents.send('google-auth-success', 'Successfully authenticated with Google!');
    } catch (error) {
        console.error('Error retrieving access token', error);
        mainWindow.webContents.send('google-auth-error', 'Error authenticating. Please try logging in again.');
    }
});

ipcMain.on('logout', () => {
    clearTokens();
    oAuth2Client.setCredentials({});
    mainWindow.webContents.send('logout-complete');
});

ipcMain.on('save-spreadsheet-id', (event, spreadsheetId) => {
    const config = loadConfig();
    config.spreadsheetId = spreadsheetId;
    saveConfig(config);
    lastSuccessfulMoves = []; // Clear undo history if sheet changes
});

ipcMain.on('test-spreadsheet-access', async (event, spreadsheetId) => {
    lastSuccessfulMoves = []; // New test clears old undo history
    try {
        const drive = google.drive({ version: 'v3', auth: oAuth2Client });
        const fileResponse = await drive.files.get({ fileId: spreadsheetId, fields: 'name,mimeType' });
        const file = fileResponse.data;

        if (file.mimeType !== 'application/vnd.google-apps.spreadsheet') {
            event.sender.send('test-access-excel-detected');
            return;
        }

        const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
        const response = await sheets.spreadsheets.get({ spreadsheetId });
        const sheetTitles = response.data.sheets.map(s => ({ title: s.properties.title }));
        
        const potentialMoves = await analyzeStudentMoves(spreadsheetId);
        
        event.sender.send('test-access-success', {
            title: file.name,
            sheets: sheetTitles,
            potentialMoves: potentialMoves
        });

    } catch (error) {
        console.error('Error testing access:', error);
        event.sender.send('test-access-error', getErrorMessage(error));
    }
});

ipcMain.on('process-spreadsheet', async (event, spreadsheetId) => {
    try {
        await processGoogleSheet(spreadsheetId, event);
    } catch (error) {
        console.error('Error processing file:', error);
        event.sender.send('processing-error', getErrorMessage(error));
    }
});

ipcMain.on('revert-last-change', async (event, spreadsheetId) => {
    if (!lastSuccessfulMoves || lastSuccessfulMoves.length === 0) {
        event.sender.send('revert-error', 'No recent operation to revert.');
        return;
    }

    try {
        const config = loadConfig();
        const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
        let revertedCount = 0;

        // The "undo" is to move students from their new sheet back to their old one
        const undoMoves = lastSuccessfulMoves.map(move => ({
            ...move,
            // Swap current and new sheets for the revert operation
            currentSheet: move.newSheet,
            newSheet: move.oldSheet,
            studentData: move.originalRowData // Use original data for re-insertion
        }));

        // Delete students from the sheet they were just moved TO
        const deleteRequests = [];
        const allSheetsInfo = (await sheets.spreadsheets.get({ spreadsheetId: config.spreadsheetId })).data.sheets;

        for (const move of lastSuccessfulMoves) {
            const sheetData = await sheets.spreadsheets.values.get({ spreadsheetId: config.spreadsheetId, range: move.newSheet });
            const rows = sheetData.data.values || [];
            
            // Find the student in the new sheet to delete them
            const rowIndexToDelete = rows.findIndex(row => row.includes(move.studentName)); // Simple name check
            
            if (rowIndexToDelete !== -1) {
                const sheetInfo = allSheetsInfo.find(s => s.properties.title === move.newSheet);
                if (sheetInfo) {
                    deleteRequests.push({
                        deleteDimension: {
                            range: {
                                sheetId: sheetInfo.properties.sheetId,
                                dimension: 'ROWS',
                                startIndex: rowIndexToDelete,
                                endIndex: rowIndexToDelete + 1,
                            },
                        },
                    });
                }
            }
        }
        
        // Append students back to their original sheets
        for (const move of lastSuccessfulMoves) {
             await sheets.spreadsheets.values.append({
                spreadsheetId: config.spreadsheetId,
                range: move.currentSheet, // The original sheet
                valueInputOption: 'USER_ENTERED',
                resource: {
                    values: [move.studentData], // The original student data
                },
            });
            revertedCount++;
        }
        
        if (deleteRequests.length > 0) {
            await sheets.spreadsheets.batchUpdate({ spreadsheetId: config.spreadsheetId, resource: { requests: deleteRequests } });
        }

        event.sender.send('revert-complete', `Revert successful! Restored ${revertedCount} students.`);
        lastSuccessfulMoves = []; // Clear history after reverting

    } catch(error) {
        console.error('Error reverting changes:', error);
        event.sender.send('revert-error', getErrorMessage(error));
    }
});


// --- Core Logic Functions ---

async function analyzeStudentMoves(spreadsheetId) {
    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
    const sheetMetadata = await sheets.spreadsheets.get({ spreadsheetId });
    const allSheets = sheetMetadata.data.sheets;

    const blacklist = ['ALL STUDENTS', 'NEW STUDENTS', 'Sheet4'];
    const availableSheets = allSheets.filter(s => !blacklist.includes(s.properties.title));

    const sheetProgression = [
        { namePattern: /young ws|young w/i, range: { min: 6, max: 8 }, displayName: 'Young WS' },
        { namePattern: /^ws(?!.*young)/i, range: { min: 9, max: 11 }, displayName: 'WS' },
        { namePattern: /young victor/i, range: { min: 12, max: 14 }, displayName: 'Young Victor' },
        { namePattern: /^victor(?!.*young)/i, range: { min: 15, max: 17 }, displayName: 'Victor' }
    ];

    let potentialMoves = [];

    for (let i = 0; i < sheetProgression.length; i++) {
        const currentLevel = sheetProgression[i];
        const currentSheetInfo = availableSheets.find(s => currentLevel.namePattern.test(s.properties.title));
        if (!currentSheetInfo) continue;
        
        const sheetName = currentSheetInfo.properties.title;
        const response = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!A:Z` });
        const rows = response.data.values;
        if (!rows || rows.length === 0) continue;

        let headerRowIndex = -1, dobIndex = -1, nameIndex = -1;
        for (let j = 0; j < rows.length; j++) {
            if (rows[j] && rows[j].includes('DOB')) {
                headerRowIndex = j;
                dobIndex = rows[j].indexOf('DOB');
                nameIndex = rows[j].findIndex(cell => cell && /name/i.test(cell) && !/user/i.test(cell));
                if (nameIndex === -1) nameIndex = rows[j].findIndex(cell => /name/i.test(cell)); // Fallback
                if (nameIndex === -1) nameIndex = 0; // Default to first col
                break;
            }
        }
        if (headerRowIndex === -1) continue;

        for (let j = headerRowIndex + 1; j < rows.length; j++) {
            const row = rows[j];
            if (!row || !row[dobIndex]) continue;

            try {
                const dob = new Date(row[dobIndex]);
                if (isNaN(dob.getTime())) continue;

                const today = new Date();
                let age = today.getFullYear() - dob.getFullYear();
                const m = today.getMonth() - dob.getMonth();
                if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) age--;

                if (age > currentLevel.range.max) {
                    let nextLevelIndex = i + 1;
                    while (nextLevelIndex < sheetProgression.length && age > sheetProgression[nextLevelIndex].range.max) {
                        nextLevelIndex++;
                    }
                    if (nextLevelIndex < sheetProgression.length) {
                        const nextLevel = sheetProgression[nextLevelIndex];
                        const nextSheetInfo = availableSheets.find(s => nextLevel.namePattern.test(s.properties.title));
                        if (nextSheetInfo && nextSheetInfo.properties.title !== sheetName) {
                            potentialMoves.push({
                                studentName: row[nameIndex] || `Student in row ${j + 1}`,
                                age: age,
                                currentSheet: sheetName,
                                currentSheetId: currentSheetInfo.properties.sheetId,
                                newSheet: nextSheetInfo.properties.title,
                                newSheetId: nextSheetInfo.properties.sheetId,
                                rowIndex: j + 1,
                                studentData: row,
                            });
                        }
                    }
                }
            } catch (e) {
                console.log(`Could not parse date for row ${j+1} in ${sheetName}: ${row[dobIndex]}`);
            }
        }
    }
    return potentialMoves;
}

async function processGoogleSheet(spreadsheetId, event) {
    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
    const moves = await analyzeStudentMoves(spreadsheetId);

    if (moves.length === 0) {
        event.sender.send('processing-complete', 'No students needed to be moved.');
        return;
    }

    // Store moves for potential revert action
    lastSuccessfulMoves = moves; 
    let movedCount = 0;

    // --- 1. Append Students to New Sheets ---
    const sourceSheetHeaders = {}; // Cache headers to avoid re-fetching
    const targetSheetHeaders = {};

    for (const move of moves) {
        // Get source headers if not cached
        if (!sourceSheetHeaders[move.currentSheet]) {
            const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: move.currentSheet });
            const headerRow = (res.data.values || []).find(r => r.includes('DOB'));
            sourceSheetHeaders[move.currentSheet] = headerRow || [];
        }
        // Get target headers if not cached
        if (!targetSheetHeaders[move.newSheet]) {
            const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: move.newSheet });
            const headerRow = (res.data.values || []).find(r => r.includes('DOB'));
            targetSheetHeaders[move.newSheet] = headerRow || [];
        }

        const sourceHeaders = sourceSheetHeaders[move.currentSheet];
        const targetHeaders = targetSheetHeaders[move.newSheet];
        
        // Map data based on headers
        const mappedData = targetHeaders.map(targetHeader => {
            const sourceIndex = sourceHeaders.indexOf(targetHeader);
            return sourceIndex !== -1 ? move.studentData[sourceIndex] : ''; // Default to empty string if column doesn't exist in source
        });

        // Use append for safety
        await sheets.spreadsheets.values.append({
            spreadsheetId,
            range: move.newSheet,
            valueInputOption: 'USER_ENTERED',
            resource: { values: [mappedData] },
        });
    }

    // --- 2. Delete Students from Old Sheets ---
    const deletionsBySheet = {};
    for (const move of moves) {
        if (!deletionsBySheet[move.currentSheetId]) {
            deletionsBySheet[move.currentSheetId] = [];
        }
        deletionsBySheet[move.currentSheetId].push(move.rowIndex);
    }

    const deleteRequests = [];
    for (const sheetId in deletionsBySheet) {
        // Sort row indices in descending order to avoid shifting issues
        const sortedRows = deletionsBySheet[sheetId].sort((a, b) => b - a);
        for (const rowIndex of sortedRows) {
            deleteRequests.push({
                deleteDimension: {
                    range: {
                        sheetId: parseInt(sheetId, 10),
                        dimension: 'ROWS',
                        startIndex: rowIndex - 1, // API is 0-indexed
                        endIndex: rowIndex,
                    },
                },
            });
        }
    }

    if (deleteRequests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            resource: { requests: deleteRequests },
        });
        movedCount = moves.length;
    }
    
    event.sender.send('processing-complete', `Successfully moved ${movedCount} students.`);
}