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
            
            // Check if tokens are expired
            if (tokens.expiry_date && tokens.expiry_date > Date.now()) {
                oAuth2Client.setCredentials(tokens);
                return true;
            } else if (tokens.refresh_token) {
                // Try to refresh the token
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
    height: 600,
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
      // Valid tokens found
      mainWindow.webContents.send('restore-session', {
        authenticated: true,
        spreadsheetId: config.spreadsheetId || ''
      });
    } else if (tokenStatus === 'refresh_needed') {
      // Try to refresh token
      const refreshed = await refreshAccessToken();
      if (refreshed) {
        mainWindow.webContents.send('restore-session', {
          authenticated: true,
          spreadsheetId: config.spreadsheetId || ''
        });
      } else {
        mainWindow.webContents.send('restore-session', {
          authenticated: false,
          spreadsheetId: config.spreadsheetId || ''
        });
      }
    } else {
      // No valid tokens
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

    async function handleAuthRedirect(url) {
        const title = authWindow.getTitle();
        if (title.startsWith('Success code=')) {
            const code = title.split('=')[1];
            authWindow.close();
            await processAuthCode(code);
        }
    }
    
    authWindow.webContents.on('did-navigate', (event, url) => {
        handleAuthRedirect(url);
    });

    authWindow.on('closed', () => {
        authWindow = null;
    });
}

async function processAuthCode(code) {
    try {
        const { tokens } = await oAuth2Client.getToken(code);
        oAuth2Client.setCredentials(tokens);
        
        // Save tokens for future use
        saveTokens(tokens);
        
        mainWindow.webContents.send('google-auth-success', 'Successfully authenticated with Google!');
    } catch (error) {
        console.error('Error retrieving access token', error);
        mainWindow.webContents.send('google-auth-error', 'Error authenticating with Google. Please check the authorization code and try again.');
    }
}

// Handle logout
ipcMain.on('logout', () => {
    clearTokens();
    oAuth2Client.setCredentials({});
    mainWindow.webContents.send('logout-complete');
});

// Handle saving spreadsheet ID
ipcMain.on('save-spreadsheet-id', (event, spreadsheetId) => {
    const config = loadConfig();
    config.spreadsheetId = spreadsheetId;
    saveConfig(config);
});

ipcMain.on('login-with-google', () => {
    startGoogleAuth();
});

ipcMain.on('submit-auth-code', async (event, code) => {
    if (authWindow) {
        authWindow.close();
        authWindow = null;
    }
    await processAuthCode(code);
});

// Enhanced function to detect file type and handle only Google Sheets
ipcMain.on('test-spreadsheet-access', async (event, fileId) => {
    try {
        const drive = google.drive({ version: 'v3', auth: oAuth2Client });
        
        console.log(`Testing access to file: ${fileId}`);
        
        // Get file metadata to determine type
        const fileResponse = await drive.files.get({
            fileId: fileId,
            fields: 'id,name,mimeType,size'
        });
        
        const file = fileResponse.data;
        console.log('File info:', {
            name: file.name,
            mimeType: file.mimeType,
            size: file.size
        });
        
        let fileInfo = {
            title: file.name,
            mimeType: file.mimeType,
            sheets: []
        };
        
        if (file.mimeType === 'application/vnd.google-apps.spreadsheet') {
            // It's a Google Sheets document
            const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
            const response = await sheets.spreadsheets.get({
                spreadsheetId: fileId,
            });
            
            fileInfo.sheets = response.data.sheets.map(sheet => ({
                title: sheet.properties.title,
                id: sheet.properties.sheetId
            }));
            fileInfo.type = 'google-sheets';
            
            // Analyze potential moves
            const moves = await analyzeStudentMoves(fileId);
            fileInfo.potentialMoves = moves;
            
            event.sender.send('test-access-success', fileInfo);
            
        } else if (file.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                   file.mimeType === 'application/vnd.ms-excel') {
            // It's an Excel file - send Excel detected event
            fileInfo.type = 'excel';
            event.sender.send('test-access-excel-detected', fileInfo);
            
        } else {
            throw new Error(`Unsupported file type: ${file.mimeType}. Please use a Google Sheets document.`);
        }
        
    } catch (error) {
        console.error('Error testing file access:', error);
        let errorMessage = 'Unknown error occurred';
        
        if (error.code === 404) {
            errorMessage = 'File not found. Please check the file ID.';
        } else if (error.code === 403) {
            errorMessage = 'Access denied. Please make sure the file is shared with your Google account.';
        } else if (error.code === 400) {
            errorMessage = 'Invalid request. The file ID might be incorrect.';
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

// Function to analyze potential student moves without actually moving them
async function analyzeStudentMoves(spreadsheetId) {
    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });

    // Get spreadsheet metadata
    const sheetMetadata = await sheets.spreadsheets.get({
        spreadsheetId,
    });

    console.log('Google Sheets title:', sheetMetadata.data.properties.title);

    // Define a blacklist of sheet names to ignore
    const blacklist = ['ALL STUDENTS', 'NEW STUDENTS', 'Sheet4'];
    const availableSheets = sheetMetadata.data.sheets.filter(s => !blacklist.includes(s.properties.title));

    console.log('Available sheets after blacklist:', availableSheets.map(s => s.properties.title));

    // Define sheet progression (from youngest to oldest)
    const sheetProgression = [
        { namePattern: /young ws|young w/i, range: { min: 6, max: 8 }, displayName: 'Young WS' },
        { namePattern: /^ws(?!.*young)/i, range: { min: 9, max: 11 }, displayName: 'WS' },
        { namePattern: /young victor/i, range: { min: 12, max: 14 }, displayName: 'Young Victor' },
        { namePattern: /^victor(?!.*young)/i, range: { min: 15, max: 17 }, displayName: 'Victor' }
    ];

    let potentialMoves = [];

    // Process each sheet in the progression
    for (let i = 0; i < sheetProgression.length; i++) {
        const currentLevel = sheetProgression[i];
        
        const currentSheet = availableSheets.find(s => 
            currentLevel.namePattern.test(s.properties.title)
        );
        
        if (!currentSheet) {
            console.log(`No sheet found matching pattern: ${currentLevel.namePattern}`);
            continue;
        }
        
        const sheetName = currentSheet.properties.title;
        const sheetId = currentSheet.properties.sheetId;
        
        console.log(`Analyzing sheet: ${sheetName}`);
        
        // Get all data from the sheet
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `${sheetName}!A:Z`,
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) {
            console.log(`No data found in sheet: ${sheetName}`);
            continue;
        }

        // Find header row and required column indices
        let headerRowIndex = -1;
        let dobIndex = -1;
        let nameIndex = -1;
        
        for (let j = 0; j < rows.length; j++) {
            const row = rows[j];
            if (row && row.includes('DOB')) {
                headerRowIndex = j;
                dobIndex = row.indexOf('DOB');
                // Look for name column (could be 'Name', 'Student Name', 'Full Name', etc.)
                nameIndex = row.findIndex(cell => 
                    cell && cell.toLowerCase().includes('name') && !cell.toLowerCase().includes('username')
                );
                if (nameIndex === -1) {
                    nameIndex = 0; // Default to first column if no name column found
                }
                break;
            }
        }
        
        if (dobIndex === -1) {
            console.log(`No DOB column found in sheet: ${sheetName}`);
            continue;
        }

        console.log(`Found DOB column at index ${dobIndex}, Name column at index ${nameIndex} in row ${headerRowIndex + 1}`);

        // Process each student row
        for (let j = headerRowIndex + 1; j < rows.length; j++) {
            const row = rows[j];
            if (!row || !row[dobIndex]) continue;

            // Parse DOB
            let dob;
            if (typeof row[dobIndex] === 'string' && row[dobIndex].includes('-')) {
                const parts = row[dobIndex].split('-');
                if (parts.length === 3) {
                    dob = new Date(`${parts[0]}-${parts[1]}-${parts[2]}`);
                }
            } else if (row[dobIndex] instanceof Date) {
                dob = new Date(row[dobIndex]);
            } else {
                dob = new Date(row[dobIndex]);
            }
            
            if (!dob || isNaN(dob.getTime())) {
                console.log(`Invalid date format for row ${j + 1}: ${row[dobIndex]}`);
                continue;
            }

            // Calculate age
            const today = new Date();
            let age = today.getFullYear() - dob.getFullYear();
            const m = today.getMonth() - dob.getMonth();
            if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) {
                age--;
            }

            console.log(`Student in row ${j + 1}: age ${age}, current level: ${currentLevel.range.min}-${currentLevel.range.max}`);

            // Check if student has outgrown current level
            if (age > currentLevel.range.max) {
                let nextLevelIndex = i + 1;
                while (nextLevelIndex < sheetProgression.length && 
                       age > sheetProgression[nextLevelIndex].range.max) {
                    nextLevelIndex++;
                }
                
                if (nextLevelIndex < sheetProgression.length) {
                    const nextLevel = sheetProgression[nextLevelIndex];
                    const nextSheet = availableSheets.find(s => 
                        nextLevel.namePattern.test(s.properties.title)
                    );
                    
                    if (nextSheet && nextSheet.properties.title !== sheetName) {
                        const studentName = row[nameIndex] || `Student in row ${j + 1}`;
                        potentialMoves.push({
                            studentName: studentName,
                            currentSheet: sheetName,
                            newSheet: nextSheet.properties.title,
                            currentLevel: currentLevel.displayName,
                            newLevel: nextLevel.displayName,
                            age: age,
                            rowIndex: j + 1,
                            studentData: row
                        });
                        console.log(`Student "${studentName}" (age ${age}) should move from "${sheetName}" to "${nextSheet.properties.title}"`);
                    }
                }
            }
        }
    }

    return potentialMoves;
}

// Process spreadsheet with confirmation
ipcMain.on('process-spreadsheet', async (event, customFileId) => {
    try {
        const drive = google.drive({ version: 'v3', auth: oAuth2Client });
        const fileId = customFileId || ' ';
        
        console.log(`Processing file: ${fileId}`);

        // Get file metadata
        const fileResponse = await drive.files.get({
            fileId: fileId,
            fields: 'id,name,mimeType'
        });
        
        const file = fileResponse.data;
        
        if (file.mimeType === 'application/vnd.google-apps.spreadsheet') {
            // Process as Google Sheets
            await processGoogleSheet(fileId, event);
        } else {
            throw new Error(`Unsupported file type: ${file.mimeType}. Only Google Sheets are supported.`);
        }
        
    } catch (error) {
        console.error('Error processing file:', error);
        
        let errorMessage = 'An error occurred while processing the file.';
        if (error.code === 404) {
            errorMessage = 'File not found. Please check the file ID.';
        } else if (error.code === 403) {
            errorMessage = 'Access denied. Please make sure the file is shared with your Google account and you have edit permissions.';
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

async function processGoogleSheet(spreadsheetId, event) {
    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });

    // Get the moves that were previously analyzed
    const moves = await analyzeStudentMoves(spreadsheetId);
    
    let totalMoved = 0;

    // Process each move
    for (const move of moves) {
        try {
            // First, get the target sheet structure to understand column layout
            const targetSheetResponse = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: `${move.newSheet}!A:Z`,
            });

            const targetRows = targetSheetResponse.data.values || [];
            
            // Find header row in target sheet
            let targetHeaderRowIndex = -1;
            let targetHeaders = [];
            
            for (let i = 0; i < Math.min(targetRows.length, 5); i++) {
                const row = targetRows[i];
                if (row && row.includes('DOB')) {
                    targetHeaderRowIndex = i;
                    targetHeaders = row;
                    break;
                }
            }
            
            if (targetHeaderRowIndex === -1) {
                console.log(`No header row found in target sheet: ${move.newSheet}`);
                continue;
            }

            // Get source sheet structure
            const sourceSheetResponse = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: `${move.currentSheet}!A:Z`,
            });

            const sourceRows = sourceSheetResponse.data.values || [];
            let sourceHeaderRowIndex = -1;
            let sourceHeaders = [];
            
            for (let i = 0; i < Math.min(sourceRows.length, 5); i++) {
                const row = sourceRows[i];
                if (row && row.includes('DOB')) {
                    sourceHeaderRowIndex = i;
                    sourceHeaders = row;
                    break;
                }
            }

            // Create a mapping between source and target columns
            let mappedData = new Array(targetHeaders.length);
            
            for (let i = 0; i < targetHeaders.length; i++) {
                const targetHeader = targetHeaders[i];
                const sourceIndex = sourceHeaders.indexOf(targetHeader);
                
                if (sourceIndex !== -1 && sourceIndex < move.studentData.length) {
                    mappedData[i] = move.studentData[sourceIndex];
                } else {
                    mappedData[i] = ''; // Empty cell if no matching column
                }
            }

            // Find the last row with data in each column of target sheet
            let insertRowIndex = targetHeaderRowIndex + 1;
            
            // Find the actual last row by checking each column
            for (let col = 0; col < targetHeaders.length; col++) {
                for (let row = targetRows.length - 1; row > targetHeaderRowIndex; row--) {
                    if (targetRows[row] && targetRows[row][col] && targetRows[row][col].trim() !== '') {
                        insertRowIndex = Math.max(insertRowIndex, row + 2); // +2 because we want the next empty row
                        break;
                    }
                }
            }
            
            console.log(`Inserting student "${move.studentName}" at row ${insertRowIndex} in sheet "${move.newSheet}"`);

            // Insert the mapped data at the correct position
            await sheets.spreadsheets.values.update({
                spreadsheetId,
                range: `${move.newSheet}!A${insertRowIndex}:${String.fromCharCode(65 + mappedData.length - 1)}${insertRowIndex}`,
                valueInputOption: 'USER_ENTERED',
                resource: {
                    values: [mappedData],
                },
            });

            // Get the current sheet metadata to find the sheet ID for deletion
            const sheetMetadata = await sheets.spreadsheets.get({
                spreadsheetId,
            });
            
            const currentSheetObj = sheetMetadata.data.sheets.find(s => 
                s.properties.title === move.currentSheet
            );
            
            if (currentSheetObj) {
                // Delete from old sheet
                await sheets.spreadsheets.batchUpdate({
                    spreadsheetId,
                    resource: {
                        requests: [{
                            deleteDimension: {
                                range: {
                                    sheetId: currentSheetObj.properties.sheetId,
                                    dimension: 'ROWS',
                                    startIndex: move.rowIndex - 1,
                                    endIndex: move.rowIndex
                                }
                            }
                        }]
                    }
                });
            }
            
            totalMoved++;
            console.log(`✓ Moved "${move.studentName}" (age ${move.age}) from "${move.currentSheet}" to "${move.newSheet}"`);
        } catch (moveError) {
            console.error(`✗ Error moving student "${move.studentName}":`, moveError.message);
        }
    }
    
    event.sender.send('processing-complete', `Google Sheets processing complete! Moved ${totalMoved} students.`);
}