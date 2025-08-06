const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { google } = require('googleapis');

// ===========================
// CONSTANTS AND CONFIGURATION
// ===========================

const WINDOW_CONFIG = {
  width: 850,
  height: 750,
  webPreferences: {
    preload: path.join(__dirname, 'preload.js'),
    contextIsolation: true,
    enableRemoteModule: false,
    nodeIntegration: false,
  },
};

const AUTH_WINDOW_CONFIG = {
  width: 600,
  height: 800,
  webPreferences: {
    nodeIntegration: false,
    contextIsolation: true,
  },
};

const GOOGLE_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.readonly',
];

const REDIRECT_URI = 'urn:ietf:wg:oauth:2.0:oob';

const SHEET_BLACKLIST = ['ALL STUDENTS', 'NEW STUDENTS', 'Sheet4'];

const SHEET_PROGRESSION = [
  { namePattern: /young ws|young w\/i/i, range: { min: 6, max: 8 } },
  { namePattern: /^ws(?!.*young)/i, range: { min: 9, max: 11 } },
  { namePattern: /young victor/i, range: { min: 12, max: 14 } },
  { namePattern: /^victor(?!.*young)/i, range: { min: 15, max: 17 } },
];

// ===========================
// GLOBAL STATE
// ===========================

class AppState {
  constructor() {
    this.mainWindow = null;
    this.authWindow = null;
    this.oAuth2Client = null;
    this.configFile = path.join(app.getPath('userData'), 'config.json');
    this.tokensFile = path.join(app.getPath('userData'), 'tokens.json');
  }

  initializeGoogleAuth() {
    try {
      const secretPath = path.join(__dirname, 'client_secret.json');
      const content = fs.readFileSync(secretPath);
      const credentials = JSON.parse(content).installed;
      const { client_secret, client_id } = credentials;
      
      this.oAuth2Client = new google.auth.OAuth2(client_id, client_secret, REDIRECT_URI);
      return true;
    } catch (err) {
      console.error('Error loading client_secret.json:', err);
      return false;
    }
  }
}

const appState = new AppState();

// ===========================
// UTILITY FUNCTIONS
// ===========================

const Utils = {
  safeJsonRead(filePath, defaultValue = {}) {
    try {
      if (fs.existsSync(filePath)) {
        return JSON.parse(fs.readFileSync(filePath, 'utf8'));
      }
    } catch (error) {
      console.error(`Error reading ${filePath}:`, error);
    }
    return defaultValue;
  },

  safeJsonWrite(filePath, data) {
    try {
      fs.writeFileSync(filePath, JSON.stringify(data, null, 2));
      return true;
    } catch (error) {
      console.error(`Error writing ${filePath}:`, error);
      return false;
    }
  },

  safeFileDelete(filePath) {
    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
      return true;
    } catch (error) {
      console.error(`Error deleting ${filePath}:`, error);
      return false;
    }
  },

  calculateAge(dateOfBirth) {
    const dob = new Date(dateOfBirth);
    if (isNaN(dob.getTime())) {
      throw new Error('Invalid date format');
    }

    const today = new Date();
    let age = today.getFullYear() - dob.getFullYear();
    const monthDiff = today.getMonth() - dob.getMonth();
    
    if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < dob.getDate())) {
      age--;
    }
    
    return age;
  },

  getGoogleApiErrorMessage(error) {
    const code = error.code || (error.response && error.response.status);
    
    switch (code) {
      case 404:
        return 'File not found. Please check the Spreadsheet ID.';
      case 403:
        return 'Permission denied. Ensure your Google account has "Editor" access to this Sheet.';
      case 401:
        TokenManager.clearTokens();
        if (appState.mainWindow) {
          appState.mainWindow.webContents.send('auth-expired');
        }
        return 'Authentication expired. Please log in again.';
      default:
        return error.message || 'An unknown error occurred with the Google API.';
    }
  },
};

// ===========================
// CONFIGURATION MANAGEMENT
// ===========================

const ConfigManager = {
  load() {
    return Utils.safeJsonRead(appState.configFile, { spreadsheetId: '' });
  },

  save(config) {
    return Utils.safeJsonWrite(appState.configFile, config);
  },

  updateSpreadsheetId(spreadsheetId) {
    const config = this.load();
    config.spreadsheetId = spreadsheetId;
    return this.save(config);
  },
};

// ===========================
// TOKEN MANAGEMENT
// ===========================

const TokenManager = {
  load() {
    try {
      const tokens = Utils.safeJsonRead(appState.tokensFile, null);
      if (!tokens) return false;

      // Check token expiry
      if (tokens.expiry_date && tokens.expiry_date <= Date.now()) {
        return tokens.refresh_token ? 'refresh_needed' : false;
      }

      appState.oAuth2Client.setCredentials(tokens);
      return true;
    } catch (error) {
      console.error('Error loading tokens:', error);
      return false;
    }
  },

  save(tokens) {
    return Utils.safeJsonWrite(appState.tokensFile, tokens);
  },

  clearTokens() {
    Utils.safeFileDelete(appState.tokensFile);
    appState.oAuth2Client.setCredentials({});
  },

  async refreshAccessToken() {
    try {
      const { credentials } = await appState.oAuth2Client.refreshAccessToken();
      appState.oAuth2Client.setCredentials(credentials);
      this.save(credentials);
      return true;
    } catch (error) {
      console.error('Error refreshing token:', error);
      this.clearTokens();
      return false;
    }
  },
};

// ===========================
// WINDOW MANAGEMENT
// ===========================

const WindowManager = {
  createMainWindow() {
    appState.mainWindow = new BrowserWindow(WINDOW_CONFIG);
    appState.mainWindow.loadFile('index.html');
    
    appState.mainWindow.on('closed', () => {
      appState.mainWindow = null;
    });

    appState.mainWindow.webContents.setWindowOpenHandler(({ url }) => {
      require('electron').shell.openExternal(url);
      return { action: 'deny' };
    });

    appState.mainWindow.webContents.once('did-finish-load', async () => {
      await this.restoreSession();
    });
  },

  createAuthWindow(authUrl) {
    appState.authWindow = new BrowserWindow(AUTH_WINDOW_CONFIG);
    appState.authWindow.loadURL(authUrl);
    appState.authWindow.on('closed', () => {
      appState.authWindow = null;
    });
  },

  closeAuthWindow() {
    if (appState.authWindow) {
      appState.authWindow.close();
      appState.authWindow = null;
    }
  },

  async restoreSession() {
    const config = ConfigManager.load();
    const tokenStatus = TokenManager.load();
    let authenticated = false;

    if (tokenStatus === true) {
      authenticated = true;
    } else if (tokenStatus === 'refresh_needed') {
      authenticated = await TokenManager.refreshAccessToken();
    }

    if (!authenticated) {
      TokenManager.clearTokens();
    }

    appState.mainWindow.webContents.send('restore-session', {
      authenticated,
      spreadsheetId: config.spreadsheetId || '',
    });
  },
};

// ===========================
// GOOGLE SHEETS OPERATIONS
// ===========================

const SheetsOperations = {
  async getFileMetadata(spreadsheetId) {
    const drive = google.drive({ version: 'v3', auth: appState.oAuth2Client });
    const response = await drive.files.get({
      fileId: spreadsheetId,
      fields: 'name,mimeType',
    });
    return response.data;
  },

  async getSheetTitles(spreadsheetId) {
    const sheets = google.sheets({ version: 'v4', auth: appState.oAuth2Client });
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: 'sheets.properties.title',
    });
    return response.data.sheets.map(s => ({ title: s.properties.title }));
  },

  async getSheetData(spreadsheetId, sheetName) {
    const sheets = google.sheets({ version: 'v4', auth: appState.oAuth2Client });
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:Z`,
    });
    return response.data.values || [];
  },

  async getAllSheetsData(spreadsheetId) {
    const sheets = google.sheets({ version: 'v4', auth: appState.oAuth2Client });
    const metadata = await sheets.spreadsheets.get({ spreadsheetId });
    return metadata.data.sheets;
  },

  async appendStudentData(spreadsheetId, sheetName, data) {
    const sheets = google.sheets({ version: 'v4', auth: appState.oAuth2Client });
    return await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: sheetName,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [data] },
    });
  },

  async deleteRows(spreadsheetId, deleteRequests) {
    if (deleteRequests.length === 0) return;
    
    const sheets = google.sheets({ version: 'v4', auth: appState.oAuth2Client });
    return await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: { requests: deleteRequests },
    });
  },
};

// ===========================
// STUDENT MOVEMENT LOGIC
// ===========================

const StudentMover = {
  async analyzeStudentMoves(spreadsheetId) {
    const allSheets = await SheetsOperations.getAllSheetsData(spreadsheetId);
    const availableSheets = allSheets.filter(
      s => !SHEET_BLACKLIST.includes(s.properties.title)
    );

    const potentialMoves = [];

    for (let i = 0; i < SHEET_PROGRESSION.length; i++) {
      const currentLevel = SHEET_PROGRESSION[i];
      const currentSheetInfo = availableSheets.find(s =>
        currentLevel.namePattern.test(s.properties.title)
      );
      
      if (!currentSheetInfo) continue;

      const moves = await this.analyzeSingleSheet(
        spreadsheetId,
        currentSheetInfo,
        currentLevel,
        availableSheets,
        i
      );
      
      potentialMoves.push(...moves);
    }

    return potentialMoves;
  },

  async analyzeSingleSheet(spreadsheetId, sheetInfo, currentLevel, availableSheets, levelIndex) {
    const sheetName = sheetInfo.properties.title;
    const rows = await SheetsOperations.getSheetData(spreadsheetId, sheetName);
    
    if (!rows || rows.length === 0) return [];

    const headerInfo = this.findHeaderRow(rows);
    if (!headerInfo) return [];

    const moves = [];
    
    for (let rowIndex = headerInfo.index + 1; rowIndex < rows.length; rowIndex++) {
      const row = rows[rowIndex];
      if (!row || !row[headerInfo.dobIndex]) continue;

      try {
        const age = Utils.calculateAge(row[headerInfo.dobIndex]);
        
        if (age > currentLevel.range.max) {
          const targetSheet = this.findTargetSheet(age, levelIndex, availableSheets);
          
          if (targetSheet && targetSheet.properties.title !== sheetName) {
            moves.push({
              studentName: row[headerInfo.nameIndex] || `Student in row ${rowIndex + 1}`,
              age: age,
              currentSheet: sheetName,
              currentSheetId: sheetInfo.properties.sheetId,
              newSheet: targetSheet.properties.title,
              rowIndex: rowIndex,
              studentData: row,
            });
          }
        }
      } catch (error) {
        console.log(`Could not parse date for row ${rowIndex + 1} in ${sheetName}: ${row[headerInfo.dobIndex]}`);
      }
    }

    return moves;
  },

  findHeaderRow(rows) {
    for (let i = 0; i < rows.length; i++) {
      if (rows[i] && rows[i].includes('DOB')) {
        const dobIndex = rows[i].indexOf('DOB');
        const nameIndex = rows[i].findIndex(
          cell => cell && /name/i.test(cell) && !/user/i.test(cell)
        );
        
        return {
          index: i,
          dobIndex: dobIndex,
          nameIndex: nameIndex === -1 ? 0 : nameIndex,
        };
      }
    }
    return null;
  },

  findTargetSheet(age, currentLevelIndex, availableSheets) {
    let nextLevelIndex = currentLevelIndex + 1;
    
    // Find the appropriate level for the student's age
    while (nextLevelIndex < SHEET_PROGRESSION.length && age > SHEET_PROGRESSION[nextLevelIndex].range.max) {
      nextLevelIndex++;
    }
    
    if (nextLevelIndex >= SHEET_PROGRESSION.length) return null;
    
    const nextLevel = SHEET_PROGRESSION[nextLevelIndex];
    return availableSheets.find(s => nextLevel.namePattern.test(s.properties.title));
  },

  async processStudentMoves(spreadsheetId, moves) {
    if (moves.length === 0) return 'No students needed to be moved.';

    const headers = await this.collectHeaders(spreadsheetId, moves);
    
    // 1. Move students to new sheets
    await this.moveStudentsToNewSheets(spreadsheetId, moves, headers);
    
    // 2. Delete students from old sheets
    await this.deleteStudentsFromOldSheets(spreadsheetId, moves);
    
    return `Successfully moved ${moves.length} student(s).`;
  },

  async collectHeaders(spreadsheetId, moves) {
    const sourceHeaders = {};
    const targetHeaders = {};
    
    for (const move of moves) {
      if (!sourceHeaders[move.currentSheet]) {
        const rows = await SheetsOperations.getSheetData(spreadsheetId, move.currentSheet);
        const headerRow = rows.find(r => r.includes('DOB'));
        sourceHeaders[move.currentSheet] = headerRow || [];
      }
      
      if (!targetHeaders[move.newSheet]) {
        const rows = await SheetsOperations.getSheetData(spreadsheetId, move.newSheet);
        const headerRow = rows.find(r => r.includes('DOB'));
        targetHeaders[move.newSheet] = headerRow || [];
      }
    }
    
    return { sourceHeaders, targetHeaders };
  },

  async moveStudentsToNewSheets(spreadsheetId, moves, headers) {
    for (const move of moves) {
      const sourceHeaderRow = headers.sourceHeaders[move.currentSheet];
      const targetHeaderRow = headers.targetHeaders[move.newSheet];
      
      const mappedData = targetHeaderRow.map(targetHeader => {
        const sourceIndex = sourceHeaderRow.indexOf(targetHeader);
        return sourceIndex !== -1 ? (move.studentData[sourceIndex] || '') : '';
      });
      
      await SheetsOperations.appendStudentData(spreadsheetId, move.newSheet, mappedData);
    }
  },

  async deleteStudentsFromOldSheets(spreadsheetId, moves) {
    const deletionsBySheet = this.groupDeletionsBySheet(moves);
    const deleteRequests = this.buildDeleteRequests(deletionsBySheet);
    
    await SheetsOperations.deleteRows(spreadsheetId, deleteRequests);
  },

  groupDeletionsBySheet(moves) {
    const deletions = {};
    
    for (const move of moves) {
      if (!deletions[move.currentSheetId]) {
        deletions[move.currentSheetId] = [];
      }
      deletions[move.currentSheetId].push(move.rowIndex);
    }
    
    return deletions;
  },

  buildDeleteRequests(deletionsBySheet) {
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
    
    return deleteRequests;
  },
};

// ===========================
// IPC HANDLERS
// ===========================

const IPCHandlers = {
  registerAll() {
    ipcMain.on('login-with-google', this.handleGoogleLogin);
    ipcMain.on('submit-auth-code', this.handleAuthCode);
    ipcMain.on('logout', this.handleLogout);
    ipcMain.on('save-spreadsheet-id', this.handleSaveSpreadsheetId);
    ipcMain.on('test-spreadsheet-access', this.handleTestSpreadsheetAccess);
    ipcMain.on('process-spreadsheet', this.handleProcessSpreadsheet);
  },

  handleGoogleLogin() {
    const authUrl = appState.oAuth2Client.generateAuthUrl({
      access_type: 'offline',
      prompt: 'consent',
      scope: GOOGLE_SCOPES,
    });
    
    WindowManager.createAuthWindow(authUrl);
  },

  async handleAuthCode(event, code) {
    WindowManager.closeAuthWindow();
    
    try {
      const { tokens } = await appState.oAuth2Client.getToken(code);
      appState.oAuth2Client.setCredentials(tokens);
      TokenManager.save(tokens);
      
      appState.mainWindow.webContents.send('google-auth-success', 'Successfully authenticated with Google!');
    } catch (error) {
      console.error('Error retrieving access token', error);
      appState.mainWindow.webContents.send('google-auth-error', 'Error authenticating. The code may be invalid or expired.');
    }
  },

  handleLogout() {
    TokenManager.clearTokens();
    appState.mainWindow.webContents.send('logout-complete');
  },

  handleSaveSpreadsheetId(event, spreadsheetId) {
    ConfigManager.updateSpreadsheetId(spreadsheetId);
  },

  async handleTestSpreadsheetAccess(event, spreadsheetId) {
    try {
      // Check file metadata and type
      const file = await SheetsOperations.getFileMetadata(spreadsheetId);
      
      if (file.mimeType !== 'application/vnd.google-apps.spreadsheet') {
        event.sender.send('test-access-excel-detected');
        return;
      }

      // Get sheet information and analyze moves
      const [sheetTitles, potentialMoves] = await Promise.all([
        SheetsOperations.getSheetTitles(spreadsheetId),
        StudentMover.analyzeStudentMoves(spreadsheetId),
      ]);

      event.sender.send('test-access-success', {
        title: file.name,
        sheets: sheetTitles,
        potentialMoves: potentialMoves,
      });
    } catch (error) {
      console.error('Error testing access:', error);
      event.sender.send('test-access-error', Utils.getGoogleApiErrorMessage(error));
    }
  },

  async handleProcessSpreadsheet(event, spreadsheetId) {
    try {
      const moves = await StudentMover.analyzeStudentMoves(spreadsheetId);
      const result = await StudentMover.processStudentMoves(spreadsheetId, moves);
      event.sender.send('processing-complete', result);
    } catch (error) {
      console.error('Error processing file:', error);
      event.sender.send('processing-error', Utils.getGoogleApiErrorMessage(error));
    }
  },
};

// ===========================
// APPLICATION INITIALIZATION
// ===========================

function initializeApp() {
  // Initialize Google Auth
  if (!appState.initializeGoogleAuth()) {
    console.error('Failed to initialize Google Auth. Exiting...');
    app.quit();
    return;
  }

  // Register IPC handlers
  IPCHandlers.registerAll();

  // Handle app events
  app.on('ready', WindowManager.createMainWindow);
  
  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
      app.quit();
    }
  });
  
  app.on('activate', () => {
    if (appState.mainWindow === null) {
      WindowManager.createMainWindow();
    }
  });
}

// Start the application
initializeApp();