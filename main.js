const { app, BrowserWindow, ipcMain, screen } = require('electron');
const path = require('path');
const fs = require('fs');
const { google } = require('googleapis');
const http = require('http');
const url = require('url');

// Global reference
let mainWindow;
let authWindow;
let authServer;

// File paths for storing persistent data
const CONFIG_FILE = path.join(app.getPath('userData'), 'config.json');
const TOKENS_FILE = path.join(app.getPath('userData'), 'tokens.json');

// --- Google Auth Setup ---
let credentials;
try {
    // It's recommended to bundle client_secret.json or load it from a secure location
    const secretPath = path.join(__dirname, 'client_secret.json');
    const content = fs.readFileSync(secretPath);
    credentials = JSON.parse(content).installed;
} catch (err) {
    console.error('Error loading client_secret.json:', err);
    // Use dialog box in production to inform user
    app.quit();
}
const { client_secret, client_id } = credentials;
const REDIRECT_URI = 'http://localhost:8080'; // Use specific port
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, REDIRECT_URI);

// --- Config & Token Management (Helper Functions) ---
function loadConfig() {
    try {
        if (fs.existsSync(CONFIG_FILE)) {
            return JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8'));
        }
    } catch (error) { console.error('Error loading config:', error); }
    return { spreadsheetId: '', userName: '', userEmail: '' };
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
                if (tokens.refresh_token) {
                    oAuth2Client.setCredentials(tokens); // Set old tokens to use refresh_token
                    return 'refresh_needed';
                }
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

// --- OAuth Server Helper ---
function createAuthServer() {
    return new Promise((resolve, reject) => {
        const server = http.createServer((req, res) => {
            const query = url.parse(req.url, true).query;
            
            if (query.code) {
                // Success - got the authorization code
                res.writeHead(200, {'Content-Type': 'text/html'});
                res.end(`
                    <html>
                        <body>
                            <h1>Authorization Successful!</h1>
                            <p>You can close this window and return to the application.</p>
                            <script>window.close();</script>
                        </body>
                    </html>
                `);
                
                server.close();
                resolve(query.code);
            } else if (query.error) {
                // Error in authorization
                res.writeHead(400, {'Content-Type': 'text/html'});
                res.end(`
                    <html>
                        <body>
                            <h1>Authorization Failed</h1>
                            <p>Error: ${query.error}</p>
                            <p>You can close this window and try again.</p>
                        </body>
                    </html>
                `);
                
                server.close();
                reject(new Error(query.error));
            } else {
                // Unexpected request
                res.writeHead(404, {'Content-Type': 'text/plain'});
                res.end('Not found');
            }
        });
        
        server.listen(8080, 'localhost', () => {
            console.log('Auth server listening on http://localhost:8080');
        });
        
        server.on('error', (err) => {
            reject(err);
        });
        
        // Store server reference for cleanup
        authServer = server;
    });
}

// --- Electron Window Management ---
function createWindow() {
  const { width, height } = screen.getPrimaryDisplay().workAreaSize; // Get system resolution

  mainWindow = new BrowserWindow({
    width: width,
    height: height,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      enableRemoteModule: false,
      nodeIntegration: false,
    },
  });
  mainWindow.loadFile('index.html');
  mainWindow.on('closed', () => { mainWindow = null; });

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    require('electron').shell.openExternal(url);
    return { action: 'deny' };
  });

  mainWindow.webContents.once('did-finish-load', async () => {
    const config = loadConfig();
    const tokenStatus = loadTokens();
    let authenticated = false;

    if (tokenStatus === true) {
      authenticated = true;
    } else if (tokenStatus === 'refresh_needed') {
      authenticated = await refreshAccessToken();
    }
    
    if (!authenticated) clearTokens();

    mainWindow.webContents.send('restore-session', { 
        authenticated, 
        spreadsheetId: config.spreadsheetId || '',
        userName: config.userName || ''
    });
  });
}

app.on('ready', createWindow);
app.on('window-all-closed', () => { 
    // Clean up auth server if it exists
    if (authServer) {
        authServer.close();
    }
    if (process.platform !== 'darwin') app.quit(); 
});
app.on('activate', () => { if (mainWindow === null) createWindow(); });

// --- Google API Helper ---
function getErrorMessage(error) {
    const code = error.code || (error.response && error.response.status);
    if (code === 404) return 'File not found. Please check the Spreadsheet ID.';
    if (code === 403) return 'Permission denied. Ensure your Google account has "Editor" access to this Sheet.';
    if (code === 401) {
        clearTokens();
        if(mainWindow) mainWindow.webContents.send('auth-expired');
        return 'Authentication expired. Please log in again.';
    }
    return error.message || 'An unknown error occurred with the Google API.';
}

// --- Helper function to find essential columns ---
function findEssentialColumns(row) {
    const dobIndex = row.findIndex(cell => cell && /^DOB$/i.test(cell.toString().trim()));
    
    // Look for name column - prioritize exact matches, then partial matches
    let nameIndex = row.findIndex(cell => cell && /^name$/i.test(cell.toString().trim()));
    if (nameIndex === -1) {
        nameIndex = row.findIndex(cell => cell && /name/i.test(cell.toString()) && !/user/i.test(cell.toString()));
    }
    // If still no name column found, assume first column
    if (nameIndex === -1) nameIndex = 0;
    
    return { dobIndex, nameIndex };
}

// --- Helper function to extract year from DOB ---
function getYearFromDOB(dobString) {
    try {
        // Handle various date formats
        const dobStr = dobString.toString().trim();
        
        // Try parsing as a full date first
        const date = new Date(dobStr);
        if (!isNaN(date.getTime())) {
            return date.getFullYear();
        }
        
        // If that fails, try to extract year from common formats
        // Format like "5-May-2019", "May-5-2019", "2019-05-05", etc.
        const yearMatch = dobStr.match(/\b(19|20)\d{2}\b/);
        if (yearMatch) {
            return parseInt(yearMatch[0], 10);
        }
        
        // If still no match, return null
        return null;
    } catch (e) {
        console.log(`Could not extract year from DOB: ${dobString}`);
        return null;
    }
}

// --- Helper function to calculate age based on birth year ---
function calculateAgeFromYear(birthYear) {
    const currentYear = new Date().getFullYear();
    return currentYear - birthYear;
}

// --- Helper function to renumber the "No." column in all sheets ---
async function autoCorrectNumbering(spreadsheetId) {
    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
    const sheetMetadata = await sheets.spreadsheets.get({ spreadsheetId });
    const allSheets = sheetMetadata.data.sheets;
    
    // Only process sheets that are not in the blacklist
    const blacklist = ['ALL STUDENTS', 'NEW STUDENTS', 'Sheet4'];
    const sheetsToCorrect = allSheets.filter(s => !blacklist.includes(s.properties.title));
    
    const updateRequests = [];
    
    for (const sheetInfo of sheetsToCorrect) {
        const sheetName = sheetInfo.properties.title;
        const sheetId = sheetInfo.properties.sheetId;
        
        try {
            // Get all data from the sheet
            const response = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: `${sheetName}!A:Z`
            });
            
            const rows = response.data.values || [];
            if (rows.length === 0) continue;
            
            // Find the header row (contains DOB)
            let headerRowIndex = -1;
            for (let i = 0; i < rows.length; i++) {
                if (rows[i] && rows[i].some(cell => cell && /DOB/i.test(cell.toString()))) {
                    headerRowIndex = i;
                    break;
                }
            }
            
            if (headerRowIndex === -1) {
                console.log(`Skipping sheet ${sheetName}: No DOB column found`);
                continue;
            }
            
            // Check if the first column should be numbered (not a name or DOB column)
            const headerRow = rows[headerRowIndex];
            const essentials = findEssentialColumns(headerRow);
            
            // Only renumber if the first column is not the name or DOB column
            const shouldRenumber = (0 !== essentials.nameIndex && 0 !== essentials.dobIndex);
            
            if (!shouldRenumber) {
                console.log(`Skipping sheet ${sheetName}: First column appears to be name or DOB`);
                continue;
            }
            
            // Find all data rows (after header) that have content
            const dataRows = [];
            for (let i = headerRowIndex + 1; i < rows.length; i++) {
                if (rows[i] && rows[i].some(cell => cell && cell.toString().trim() !== '')) {
                    dataRows.push({
                        rowIndex: i,
                        data: rows[i]
                    });
                }
            }
            
            if (dataRows.length === 0) continue;
            
            // Create update requests to renumber the first column
            for (let i = 0; i < dataRows.length; i++) {
                const newNumber = (i + 1).toString();
                const rowIndex = dataRows[i].rowIndex;
                
                updateRequests.push({
                    updateCells: {
                        range: {
                            sheetId: sheetId,
                            startRowIndex: rowIndex,
                            endRowIndex: rowIndex + 1,
                            startColumnIndex: 0,
                            endColumnIndex: 1
                        },
                        rows: [{
                            values: [{
                                userEnteredValue: {
                                    stringValue: newNumber
                                }
                            }]
                        }],
                        fields: 'userEnteredValue'
                    }
                });
            }
            
            console.log(`Prepared ${dataRows.length} numbering updates for sheet: ${sheetName}`);
            
        } catch (error) {
            console.error(`Error processing sheet ${sheetName} for numbering correction:`, error);
        }
    }
    
    // Execute all updates in batches to avoid API limits
    if (updateRequests.length > 0) {
        const batchSize = 100; // Google Sheets API limit
        for (let i = 0; i < updateRequests.length; i += batchSize) {
            const batch = updateRequests.slice(i, i + batchSize);
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                resource: { requests: batch }
            });
        }
        
        console.log(`Applied ${updateRequests.length} numbering corrections across all sheets`);
        return updateRequests.length;
    }
    
    return 0;
}

// --- IPC Handlers ---
ipcMain.on('login-with-google', async () => {
    try {
        const authUrl = oAuth2Client.generateAuthUrl({
            access_type: 'offline',
            prompt: 'consent',
            scope: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/userinfo.email', 'https://www.googleapis.com/auth/userinfo.profile'],
        });

        // Clean up any existing auth server
        if (authServer) {
            authServer.close();
            authServer = null;
        }

        // Create and start the auth server
        const authCodePromise = createAuthServer();
        
        // Open the auth URL in a browser window
        authWindow = new BrowserWindow({ 
            width: 600, 
            height: 800, 
            webPreferences: { 
                nodeIntegration: false, 
                contextIsolation: true 
            } 
        });
        
        authWindow.loadURL(authUrl);
        
        // Handle auth window being closed manually
        authWindow.on('closed', () => { 
            console.log('Auth window was closed');
            authWindow = null;
            
            // Clean up auth server
            if (authServer) {
                authServer.close();
                authServer = null;
            }
            
            // Re-enable login button and hide instructions
            if (mainWindow && !mainWindow.isDestroyed()) {
                mainWindow.webContents.send('google-auth-cancelled');
            }
        });

        try {
            const code = await authCodePromise;
            if (authWindow) {
                authWindow.close();
                authWindow = null;
            }
            const { tokens } = await oAuth2Client.getToken(code);
            oAuth2Client.setCredentials(tokens);
            saveTokens(tokens);

            // Fetch user info after authentication
            const oauth2 = google.oauth2({ version: 'v2', auth: oAuth2Client });
            const userInfo = await oauth2.userinfo.get();
            const userEmail = userInfo.data.email;
            const userName = userInfo.data.name;

            // Save userName and userEmail in config
            const config = loadConfig();
            config.userEmail = userEmail;
            config.userName = userName;
            saveConfig(config);

            mainWindow.webContents.send('google-auth-success', {
                message: 'Successfully authenticated with Google!',
                userName: userName
            });

        } catch (authError) {
            console.error('Error during authorization:', authError);
            
            // Clean up auth window if still open
            if (authWindow) {
                authWindow.close();
                authWindow = null;
            }
            
            mainWindow.webContents.send('google-auth-error', 'Error during authorization. Please try again.');
        }
        
    } catch (error) {
        console.error('Error starting auth process:', error);
        mainWindow.webContents.send('google-auth-error', 'Error starting authorization process.');
    }
});

// Remove the old submit-auth-code handler as it's no longer needed
// ipcMain.on('submit-auth-code', async (event, code) => { ... });

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

ipcMain.on('test-spreadsheet-access', async (event, spreadsheetId) => {
    try {
        // 1. Check file metadata and type first
        const drive = google.drive({ version: 'v3', auth: oAuth2Client });
        const fileResponse = await drive.files.get({ fileId: spreadsheetId, fields: 'name,mimeType' });
        const file = fileResponse.data;

        if (file.mimeType !== 'application/vnd.google-apps.spreadsheet') {
            event.sender.send('test-access-excel-detected');
            return;
        }

        // 2. Get sheet titles
        const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
        const response = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties.title' });
        const sheetTitles = response.data.sheets.map(s => ({ title: s.properties.title }));
        
        // 3. Analyze for potential moves
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

// --- Core Logic Functions ---
async function analyzeStudentMoves(spreadsheetId) {
    const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
    const sheetMetadata = await sheets.spreadsheets.get({ spreadsheetId });
    const allSheets = sheetMetadata.data.sheets;

    const blacklist = ['ALL STUDENTS', 'NEW STUDENTS', 'Sheet4'];
    const availableSheets = allSheets.filter(s => !blacklist.includes(s.properties.title));

    // Updated age ranges based on year-only calculation
    const sheetProgression = [
        { namePattern: /young ws|young w\/i/i, range: { min: 6, max: 8 } },
        { namePattern: /^ws(?!.*young)/i, range: { min: 9, max: 11 } },
        { namePattern: /young victor/i, range: { min: 12, max: 14 } },
        { namePattern: /^victor(?!.*young)/i, range: { min: 15, max: 17 } }
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
            if (rows[j] && rows[j].some(cell => cell && /DOB/i.test(cell.toString()))) {
                headerRowIndex = j;
                const essentialColumns = findEssentialColumns(rows[j]);
                dobIndex = essentialColumns.dobIndex;
                nameIndex = essentialColumns.nameIndex;
                break;
            }
        }
        
        // Skip if no DOB column found
        if (headerRowIndex === -1 || dobIndex === -1) {
            console.log(`Skipping sheet ${sheetName}: No DOB column found`);
            continue;
        }

        for (let j = headerRowIndex + 1; j < rows.length; j++) {
            const row = rows[j];
            if (!row || !row[dobIndex]) continue;

            // Check if this row has at least name and DOB
            const hasName = nameIndex !== -1 && row[nameIndex] && row[nameIndex].toString().trim() !== '';
            const hasDOB = row[dobIndex] && row[dobIndex].toString().trim() !== '';
            
            if (!hasName && !hasDOB) continue; // Skip if missing both essential fields

            // Extract year from DOB and calculate age
            const birthYear = getYearFromDOB(row[dobIndex]);
            if (birthYear === null) {
                console.log(`Could not extract year from DOB for row ${j + 1} in ${sheetName}: ${row[dobIndex]}`);
                continue;
            }

            const age = calculateAgeFromYear(birthYear);

            // Check if student should move to a higher level
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
                            studentName: hasName ? row[nameIndex] : `Student in row ${j + 1}`,
                            age: age,
                            birthYear: birthYear,
                            currentSheet: sheetName,
                            currentSheetId: currentSheetInfo.properties.sheetId,
                            newSheet: nextSheetInfo.properties.title,
                            newSheetId: nextSheetInfo.properties.sheetId,
                            rowIndex: j, // 0-indexed row
                            studentData: row,
                            hasOnlyEssentials: !hasName || row.filter(cell => cell && cell.toString().trim() !== '').length <= 2
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

    if (moves.length === 0) {
        event.sender.send('processing-complete', 'No students needed to be moved.');
        return;
    }

    // Step 1: Delete students from old sheets FIRST (before insertion affects row numbers)
    const deletionsBySheet = {};
    for (const move of moves) {
        if (!deletionsBySheet[move.currentSheetId]) {
            deletionsBySheet[move.currentSheetId] = [];
        }
        deletionsBySheet[move.currentSheetId].push(move.rowIndex);
    }

    const deleteRequests = [];
    for (const sheetId in deletionsBySheet) {
        // Sort row indices descending to avoid shifting issues during deletion
        const sortedRows = deletionsBySheet[sheetId].sort((a, b) => b - a);
        for (const rowIndex of sortedRows) {
            deleteRequests.push({
                deleteDimension: {
                    range: {
                        sheetId: parseInt(sheetId, 10),
                        dimension: 'ROWS',
                        startIndex: rowIndex, 
                        endIndex: rowIndex + 1,
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
    }

    // Step 2: Insert students into new sheets and apply borders
    const sourceSheetHeaders = {};
    const targetSheetHeaders = {};
    const targetSheetData = {};
    const insertedRows = [];
    
    // Track next number and next insertion row for each target sheet
    const sheetNextNumbers = {};
    const sheetNextInsertionRows = {};

    for (const move of moves) {
        // Get source sheet headers if not cached
        if (!sourceSheetHeaders[move.currentSheet]) {
            const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: move.currentSheet });
            const headerRow = (res.data.values || []).find(r => r && r.some(cell => cell && /DOB/i.test(cell.toString())));
            sourceSheetHeaders[move.currentSheet] = headerRow || [];
        }
        
        // Get target sheet data if not cached (refresh after deletions)
        if (!targetSheetData[move.newSheet]) {
            const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: move.newSheet });
            const allRows = res.data.values || [];
            targetSheetData[move.newSheet] = allRows;
            
            // Find and cache headers
            const headerRow = allRows.find(r => r && r.some(cell => cell && /DOB/i.test(cell.toString())));
            targetSheetHeaders[move.newSheet] = headerRow || [];
        }

        const sourceHeaders = sourceSheetHeaders[move.currentSheet];
        const targetHeaders = targetSheetHeaders[move.newSheet];
        
        // Find essential column indices for both source and target
        const sourceEssentials = findEssentialColumns(sourceHeaders);
        const targetEssentials = findEssentialColumns(targetHeaders);
        
        // Initialize tracking for this target sheet (only once per sheet)
        if (!sheetNextNumbers[move.newSheet]) {
            const targetRows = targetSheetData[move.newSheet];
            const headerRowIndex = targetRows.findIndex(r => r && r.some(cell => cell && /DOB/i.test(cell.toString())));
            
            let maxNumber = 0;
            let lastRowWithData = 0;
            
            // Find the highest number in the first column and the last row with data
            for (let i = targetRows.length - 1; i >= 0; i--) {
                if (targetRows[i] && targetRows[i].some(cell => cell && cell.toString().trim() !== '')) {
                    if (lastRowWithData === 0) lastRowWithData = i + 1; // First non-empty row from bottom
                    
                    // Check if this row is after the header and has a number in the first column
                    if (i > headerRowIndex && targetRows[i][0]) {
                        const cellValue = targetRows[i][0].toString().trim();
                        const number = parseInt(cellValue, 10);
                        
                        // Only consider valid numbers (not text, dates, etc.)
                        if (!isNaN(number) && number > 0 && cellValue === number.toString()) {
                            maxNumber = Math.max(maxNumber, number);
                        }
                    }
                }
            }
            
            // If no data found, start after potential header row
            if (lastRowWithData === 0) {
                lastRowWithData = headerRowIndex >= 0 ? headerRowIndex + 1 : 1;
            }
            
            // Set next number to be one higher than the maximum found
            sheetNextNumbers[move.newSheet] = maxNumber + 1;
            sheetNextInsertionRows[move.newSheet] = lastRowWithData;
        }

        // Use the current next number for this sheet and increment it
        const currentNumber = sheetNextNumbers[move.newSheet];
        const insertionRow = sheetNextInsertionRows[move.newSheet];
        
        sheetNextNumbers[move.newSheet]++;
        sheetNextInsertionRows[move.newSheet]++;

        // Map data from source to target format
        const mappedData = targetHeaders.map((targetHeader, targetIndex) => {
            // For the first column (assumed to be the numbering column), use sequential number
            if (targetIndex === 0 && targetIndex !== targetEssentials.nameIndex && targetIndex !== targetEssentials.dobIndex) {
                return currentNumber.toString();
            }
            
            // Map essential fields directly
            if (targetIndex === targetEssentials.nameIndex && sourceEssentials.nameIndex !== -1) {
                return move.studentData[sourceEssentials.nameIndex] || '';
            }
            
            if (targetIndex === targetEssentials.dobIndex && sourceEssentials.dobIndex !== -1) {
                return move.studentData[sourceEssentials.dobIndex] || '';
            }
            
            // For other columns, try to map by header name if available
            if (targetHeader && sourceHeaders.length > 0) {
                const sourceIndex = sourceHeaders.indexOf(targetHeader);
                if (sourceIndex !== -1 && move.studentData[sourceIndex]) {
                    return move.studentData[sourceIndex];
                }
            }
            
            // Return empty string for unmapped columns
            return '';
        });

        // Use append with a specific range starting from the tracked insertion row
        const appendRange = `${move.newSheet}!A${insertionRow + 1}:Z${insertionRow + 1}`;
        
        await sheets.spreadsheets.values.append({
            spreadsheetId,
            range: appendRange,
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: [mappedData] },
        });

        // Track inserted row for border formatting
        insertedRows.push({
            sheetId: move.newSheetId,
            rowIndex: insertionRow, // 0-indexed
            columnCount: targetHeaders.length
        });

        // Update cached data to include the new row
        targetSheetData[move.newSheet][insertionRow] = mappedData;
    }

    // Step 3: Apply borders to newly inserted rows
    if (insertedRows.length > 0) {
        const borderRequests = [];
        
        for (const row of insertedRows) {
            borderRequests.push({
                updateBorders: {
                    range: {
                        sheetId: row.sheetId,
                        startRowIndex: row.rowIndex,
                        endRowIndex: row.rowIndex + 1,
                        startColumnIndex: 0,
                        endColumnIndex: row.columnCount
                    },
                    top: {
                        style: 'SOLID',
                        width: 1,
                        color: { red: 0, green: 0, blue: 0 }
                    },
                    bottom: {
                        style: 'SOLID',
                        width: 1,
                        color: { red: 0, green: 0, blue: 0 }
                    },
                    left: {
                        style: 'SOLID',
                        width: 1,
                        color: { red: 0, green: 0, blue: 0 }
                    },
                    right: {
                        style: 'SOLID',
                        width: 1,
                        color: { red: 0, green: 0, blue: 0 }
                    },
                    innerHorizontal: {
                        style: 'SOLID',
                        width: 1,
                        color: { red: 0, green: 0, blue: 0 }
                    },
                    innerVertical: {
                        style: 'SOLID',
                        width: 1,
                        color: { red: 0, green: 0, blue: 0 }
                    }
                }
            });
        }

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            resource: { requests: borderRequests },
        });
    }
    
    // Step 4: Auto-correct numbering in all sheets
    console.log('Starting auto-correction of numbering...');
    const correctionCount = await autoCorrectNumbering(spreadsheetId);
    
    const essentialOnlyCount = moves.filter(m => m.hasOnlyEssentials).length;
    const regularCount = moves.length - essentialOnlyCount;
    
    let message = `Successfully moved ${moves.length} student(s) with borders applied.`;
    if (essentialOnlyCount > 0) {
        message += ` (${essentialOnlyCount} with minimal data, ${regularCount} with complete data)`;
    }
    
    if (correctionCount > 0) {
        message += ` Auto-corrected ${correctionCount} numbering entries.`;
    }
    
    event.sender.send('processing-complete', message);
}