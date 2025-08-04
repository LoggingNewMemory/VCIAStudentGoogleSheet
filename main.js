const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { google } = require('googleapis');

// Keep a global reference of the window object
let mainWindow;
let authWindow;

// File paths for storing persistent data
const CONFIG_FILE = path.join(app.getPath('userData'), 'config.json');
const TOKENS_FILE = path.join(app.getPath('userData'), 'tokens.json');

// Load client secrets from a local file
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

const oAuth2Client = new google.auth.OAuth2(
  client_id,
  client_secret,
  REDIRECT_URI
);

// Load saved configuration
function loadConfig() {
    try {
        if (fs.existsSync(CONFIG_FILE)) {
            const configData = fs.readFileSync(CONFIG_FILE, 'utf8');
            return JSON.parse(configData);
        }
    } catch (error) {
        console.error('Error loading config:', error);
    }
    return { spreadsheetId: '' };
}

// Save configuration
function saveConfig(config) {
    try {
        fs.writeFileSync(CONFIG_FILE, JSON.stringify(config, null, 2));
    } catch (error) {
        console.error('Error saving config:', error);
    }
}

// Load saved tokens
function loadTokens() {
    try {
        if (fs.existsSync(TOKENS_FILE)) {
            const tokensData = fs.readFileSync(TOKENS_FILE, 'utf8');
            const tokens = JSON.parse(tokensData);
            
            if (tokens.expiry_date && tokens.expiry_date > Date.now()) {
                oAuth2Client.setCredentials(tokens);
                return true;
            } else if (tokens.refresh_token) {
                oAuth2Client.setCredentials(tokens);
                return 'refresh_needed';
            }
        }
    } catch (error) {
        console.error('Error loading tokens:', error);
    }
    return false;
}

// Save tokens
function saveTokens(tokens) {
    try {
        fs.writeFileSync(TOKENS_FILE, JSON.stringify(tokens, null, 2));
    } catch (error) {
        console.error('Error saving tokens:', error);
    }
}

// Clear saved tokens
function clearTokens() {
    try {
        if (fs.existsSync(TOKENS_FILE)) {
            fs.unlinkSync(TOKENS_FILE);
        }
    } catch (error) {
        console.error('Error clearing tokens:', error);
    }
}

// Refresh access token using refresh token
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

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 700, // Increased height for more content
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      enableRemoteModule: false,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile('index.html');

  mainWindow.on('closed', function () {
    mainWindow = null;
  });

  // Check for saved login and configuration when window is ready
  mainWindow.webContents.once('did-finish-load', async () => {
    const config = loadConfig();
    const tokenStatus = loadTokens();
    
    if (tokenStatus === true) {
      mainWindow.webContents.send('restore-session', {
        authenticated: true,
        spreadsheetId: config.spreadsheetId || ''
      });
    } else if (tokenStatus === 'refresh_needed') {
      const refreshed = await refreshAccessToken();
      mainWindow.webContents.send('restore-session', {
        authenticated: refreshed,
        spreadsheetId: config.spreadsheetId || ''
      });
    } else {
      mainWindow.webContents.send('restore-session', {
        authenticated: false,
        spreadsheetId: config.spreadsheetId || ''
      });
    }
  });
}

app.on('ready', createWindow);

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', function () {
  if (mainWindow === null) {
    createWindow();
  }
});

function startGoogleAuth() {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive.readonly'
        ],
    });

    authWindow = new BrowserWindow({
        width: 600,
        height: 800,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true
        }
    });

    authWindow.loadURL(authUrl);
    authWindow.show();

    authWindow.on('closed', () => {
        authWindow = null;
    });
}

async function processAuthCode(code) {
    try {
        const { tokens } = await oAuth2Client.getToken(code);
        oAuth2Client.setCredentials(tokens);
        saveTokens(tokens);
        mainWindow.webContents.send('google-auth-success', 'Successfully authenticated with Google!');
    } catch (error) {
        console.error('Error retrieving access token', error);
        mainWindow.webContents.send('google-auth-error', 'Error authenticating. Please check the authorization code and try again.');
    }
}

ipcMain.on('logout', () => {
    clearTokens();
    oAuth2Client.setCredentials({});
    mainWindow.webContents.send('logout-complete');
});

ipcMain.on('save-spreadsheet-id', (event, spreadsheetId) => {
    const config = loadConfig();
    config.spreadsheetId = spreadsheetId;
    saveConfig(config);
});

ipcMain.on('login-with-google', startGoogleAuth);

ipcMain.on('submit-auth-code', async (event, code) => {
    if (authWindow) authWindow.close();
    await processAuthCode(code);
});

// Function to test access and analyze potential moves
ipcMain.on('test-spreadsheet-access', async (event, fileId) => {
    try {
        const drive = google.drive({ version: 'v3', auth: oAuth2Client });
        
        const fileResponse = await drive.files.get({
            fileId: fileId,
            fields: 'id,name,mimeType',
        });
        
        const file = fileResponse.data;
        
        if (file.mimeType !== 'application/vnd.google-apps.spreadsheet') {
            throw new Error(`Unsupported file type: ${file.mimeType}. Please use a Google Sheets document.`);
        }
        
        // It's a Google Sheet, proceed with analysis
        const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
        const response = await sheets.spreadsheets.get({
            spreadsheetId: fileId,
        });
        
        const potentialMoves = await analyzeStudentMoves(fileId);
        
        const fileInfo = {
            title: file.name,
            sheets: response.data.sheets.map(sheet => ({
                title: sheet.properties.title,
                id: sheet.properties.sheetId
            })),
            potentialMoves: potentialMoves,
        };
        
        event.sender.send('test-access-success', fileInfo);
        
    } catch (error) {
        console.error('Error testing file access:', error);
        let errorMessage = 'Unknown error occurred';
        
        if (error.code === 404) {
            errorMessage = 'File not found. Please check the file ID.';
        } else if (error.code === 403) {
            errorMessage = 'Access denied. Please make sure the file is shared with your Google account.';
        } else if (error.code === 401) {
            errorMessage = 'Authentication expired. Please log in again.';
            clearTokens();
            mainWindow.webContents.send('auth-expired');
        } else if (error.message) {
            errorMessage = error.message;
        }
        
        event.sender.send('test-access-error', errorMessage);
    }
});

// Main processing function
ipcMain.on('process-spreadsheet', async (event, spreadsheetId) => {
    try {
        const drive = google.drive({ version: 'v3', auth: oAuth2Client });
        
        const fileResponse = await drive.files.get({
            fileId: spreadsheetId,
            fields: 'mimeType',
        });

        if (fileResponse.data.mimeType !== 'application/vnd.google-apps.spreadsheet') {
            throw new Error('Unsupported file type. Only Google Sheets are supported for processing.');
        }

        await processGoogleSheet(spreadsheetId, event);

    } catch (error) {
        console.error('Error processing file:', error);
        let errorMessage = 'An error occurred while processing the file.';
        if (error.code === 403) {
            errorMessage = 'Access denied. Ensure you have edit permissions for the sheet.';
        } else if (error.code === 401) {
            errorMessage = 'Authentication expired. Please log in again.';
            clearTokens();
            mainWindow.webContents.send('auth-expired');
        } else if (error.message) {
            errorMessage = error.message;
        }
        event.sender.send('processing-error', errorMessage);
    }
});


async function analyzeStudentMoves(spreadsheetId) {
    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });

    const sheetMetadata = await sheets.spreadsheets.get({ spreadsheetId });
    const blacklist = ['ALL STUDENTS', 'NEW STUDENTS', 'Sheet4'];
    const availableSheets = sheetMetadata.data.sheets.filter(s => !blacklist.includes(s.properties.title));

    const sheetProgression = [
        { namePattern: /young ws|young w/i, range: { min: 6, max: 8 }, displayName: 'Young WS' },
        { namePattern: /^ws(?!.*young)/i, range: { min: 9, max: 11 }, displayName: 'WS' },
        { namePattern: /young victor/i, range: { min: 12, max: 14 }, displayName: 'Young Victor' },
        { namePattern: /^victor(?!.*young)/i, range: { min: 15, max: 17 }, displayName: 'Victor' }
    ];

    let potentialMoves = [];

    for (let i = 0; i < sheetProgression.length; i++) {
        const currentLevel = sheetProgression[i];
        const currentSheet = availableSheets.find(s => currentLevel.namePattern.test(s.properties.title));
        if (!currentSheet) continue;

        const sheetName = currentSheet.properties.title;
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `${sheetName}!A:Z`,
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) continue;

        let headerRowIndex = rows.findIndex(row => row && row.includes('DOB'));
        if (headerRowIndex === -1) continue;

        const dobIndex = rows[headerRowIndex].indexOf('DOB');
        let nameIndex = rows[headerRowIndex].findIndex(cell => cell && cell.toLowerCase().includes('name') && !cell.toLowerCase().includes('username'));
        if (nameIndex === -1) nameIndex = 0;

        for (let j = headerRowIndex + 1; j < rows.length; j++) {
            const row = rows[j];
            if (!row || !row[dobIndex]) continue;

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
                    const nextSheet = availableSheets.find(s => nextLevel.namePattern.test(s.properties.title));
                    
                    if (nextSheet && nextSheet.properties.title !== sheetName) {
                        potentialMoves.push({
                            studentName: row[nameIndex] || `Student in row ${j + 1}`,
                            currentSheet: sheetName,
                            newSheet: nextSheet.properties.title,
                            age: age,
                            rowIndex: j + 1,
                            studentData: row,
                            sourceSheetId: currentSheet.properties.sheetId,
                        });
                    }
                }
            }
        }
    }
    return potentialMoves;
}

async function processGoogleSheet(spreadsheetId, event) {
    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
    const moves = await analyzeStudentMoves(spreadsheetId);
    let movedStudentsSummary = [];
    let errors = [];

    // Process moves in reverse order of row index to avoid shifting issues
    moves.sort((a, b) => b.rowIndex - a.rowIndex);

    for (const move of moves) {
        try {
            // Append data to the new sheet
            await sheets.spreadsheets.values.append({
                spreadsheetId,
                range: move.newSheet, // Append to the table in this sheet
                valueInputOption: 'USER_ENTERED',
                insertDataOption: 'INSERT_ROWS',
                resource: {
                    values: [move.studentData],
                },
            });

            // Delete the original row
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                resource: {
                    requests: [{
                        deleteDimension: {
                            range: {
                                sheetId: move.sourceSheetId,
                                dimension: 'ROWS',
                                startIndex: move.rowIndex - 1,
                                endIndex: move.rowIndex,
                            },
                        },
                    }],
                },
            });

            const summaryLine = `✓ Moved <strong>${move.studentName}</strong> from ${move.currentSheet} to ${move.newSheet}`;
            movedStudentsSummary.push(summaryLine);
            console.log(`✓ Moved "${move.studentName}" from "${move.currentSheet}" to "${move.newSheet}"`);
        } catch (moveError) {
            console.error(`✗ Error moving student "${move.studentName}":`, moveError.message);
            errors.push(move.studentName);
        }
    }
    
    let completionMessage = `Processing complete! Moved ${movedStudentsSummary.length} students.`;
    if (movedStudentsSummary.length > 0) {
        completionMessage += `\n\n<strong>Summary of Changes:</strong>\n${movedStudentsSummary.join('\n')}`;
    }

    if (errors.length > 0) {
        completionMessage += `\n\nFailed to move: ${errors.join(', ')}.`;
    }
    
    event.sender.send('processing-complete', completionMessage);
}