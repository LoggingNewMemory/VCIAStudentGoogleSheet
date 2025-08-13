document.addEventListener('DOMContentLoaded', () => {
    // --- Element References ---
    const loginBtn = document.getElementById('login-btn');
    const logoutBtn = document.getElementById('logout-btn');
    const testBtn = document.getElementById('test-btn');
    const processBtn = document.getElementById('process-btn');
    const statusMessage = document.getElementById('status-message');
    const loginSection = document.getElementById('login-section');
    const authenticatedSection = document.getElementById('authenticated-section');
    const spreadsheetIdInput = document.getElementById('spreadsheet-id-input');
    const spreadsheetInfo = document.getElementById('spreadsheet-info');
    const warningInfo = document.getElementById('warning-info');
    const errorInfo = document.getElementById('error-info');
    const postProcessingInfo = document.getElementById('post-processing-info');
    const movesPreview = document.getElementById('moves-preview');
    const movesList = document.getElementById('moves-list');
    const authStatus = document.getElementById('auth-status');
    const username = document.getElementById('username');
    const logsContainer = document.getElementById('logs-container');
    
    let pendingMoves = [];

    // --- Utility Functions ---
    function hideAllInfoBoxes() {
        spreadsheetInfo.style.display = 'none';
        warningInfo.style.display = 'none';
        errorInfo.style.display = 'none';
        movesPreview.style.display = 'none';
        postProcessingInfo.style.display = 'none';
    }

    function updateStatusMessage(message, type = 'info') {
        statusMessage.textContent = message;
        statusMessage.className = `status-message ${type}`;
    }

    function logMessage(message, type = 'info') {
        const timestamp = new Date().toLocaleTimeString();
        const logEntry = `[${timestamp}] ${type.toUpperCase()}: ${message}`;
        
        if (logsContainer.textContent === 'Initializing...') {
            logsContainer.textContent = '';
        }
        
        logsContainer.textContent += logEntry + '\n';
        logsContainer.scrollTop = logsContainer.scrollHeight;
    }

    function extractSpreadsheetId(input) {
        const trimmedInput = input.trim();
        
        // If it looks like a URL, extract the ID
        if (trimmedInput.includes('docs.google.com/spreadsheets')) {
            const match = trimmedInput.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
            return match ? match[1] : null;
        }
        
        // If it's already just an ID (no slashes or dots), return as is
        if (trimmedInput && !trimmedInput.includes('/') && !trimmedInput.includes('.')) {
            return trimmedInput;
        }
        
        return null;
    }

    function updateAuthStatus(authenticated, name = '') {
        if (authenticated) {
            authStatus.textContent = '✓ Authenticated';
            authStatus.className = 'status-indicator authenticated';
            loginSection.style.display = 'none';
            authenticatedSection.style.display = 'block';
            username.textContent = name || 'User';
            logMessage('Successfully authenticated with Google', 'success');
        } else {
            authStatus.textContent = 'Not Logged In';
            authStatus.className = 'status-indicator unauthenticated';
            loginSection.style.display = 'block';
            authenticatedSection.style.display = 'none';
            username.textContent = 'User';
            hideAllInfoBoxes();
            updateStatusMessage('Please sign in to continue');
            logMessage('User logged out or session expired', 'info');
        }
    }

    function setButtonState(button, loading = false, text = null) {
        if (loading) {
            button.disabled = true;
            button.originalText = button.textContent;
            button.textContent = text || 'Loading...';
        } else {
            button.disabled = false;
            if (button.originalText) {
                button.textContent = button.originalText;
                delete button.originalText;
            }
        }
    }
    
    // --- Event Listeners ---
    spreadsheetIdInput.addEventListener('input', () => {
        const input = spreadsheetIdInput.value.trim();
        const extractedId = extractSpreadsheetId(input);
        
        if (extractedId) {
            // Update the input field to show just the ID for clarity
            spreadsheetIdInput.value = extractedId;
            window.electronAPI.saveSpreadsheetId(extractedId);
            logMessage(`Spreadsheet ID saved: ${extractedId}`, 'info');
        } else if (input === '') {
            // Clear the saved ID if input is empty
            window.electronAPI.saveSpreadsheetId('');
        }
    });

    spreadsheetIdInput.addEventListener('paste', (e) => {
        // Small delay to let the paste complete, then process
        setTimeout(() => {
            const input = spreadsheetIdInput.value.trim();
            const extractedId = extractSpreadsheetId(input);
            
            if (extractedId && extractedId !== input) {
                spreadsheetIdInput.value = extractedId;
                window.electronAPI.saveSpreadsheetId(extractedId);
                logMessage(`Spreadsheet ID extracted from URL: ${extractedId}`, 'info');
            }
        }, 10);
    });

    loginBtn.addEventListener('click', () => {
        updateStatusMessage('Opening Google sign-in...');
        setButtonState(loginBtn, true, 'Signing In...');
        logMessage('Initiating Google authentication', 'info');
        window.electronAPI.loginWithGoogle();
    });

    logoutBtn.addEventListener('click', () => {
        updateStatusMessage('Signing out...');
        setButtonState(logoutBtn, true, 'Signing Out...');
        logMessage('Logging out user', 'info');
        window.electronAPI.logout();
    });

    testBtn.addEventListener('click', () => {
        const input = spreadsheetIdInput.value.trim();
        const spreadsheetId = extractSpreadsheetId(input) || input;
        
        if (!spreadsheetId) {
            updateStatusMessage('Please enter a valid spreadsheet URL or ID');
            errorInfo.innerHTML = '<h4>Invalid Input</h4><div>Please provide a valid Google Sheets URL or ID.</div>';
            errorInfo.style.display = 'block';
            return;
        }
        
        // Update the input field to show the extracted ID
        if (spreadsheetId !== input) {
            spreadsheetIdInput.value = spreadsheetId;
        }
        
        updateStatusMessage('Testing spreadsheet access...');
        hideAllInfoBoxes();
        processBtn.style.display = 'none';
        setButtonState(testBtn, true, 'Testing...');
        logMessage(`Testing access to spreadsheet: ${spreadsheetId}`, 'info');
        window.electronAPI.testSpreadsheetAccess(spreadsheetId);
    });

    processBtn.addEventListener('click', () => {
        const input = spreadsheetIdInput.value.trim();
        const spreadsheetId = extractSpreadsheetId(input) || input;
        
        if (pendingMoves.length === 0) {
            updateStatusMessage('No moves to process');
            return;
        }
        
        updateStatusMessage('Processing spreadsheet... This may take a moment');
        setButtonState(processBtn, true, 'Processing...');
        setButtonState(testBtn, true);
        logMessage(`Processing ${pendingMoves.length} student moves`, 'info');
        window.electronAPI.processSpreadsheet(spreadsheetId);
    });
    
    // --- Electron API Event Handlers ---

    window.electronAPI.receiveGoogleAuthCancelled(() => {
        updateStatusMessage('Sign-in was cancelled. You can try again.');
        setButtonState(loginBtn, false);
        logMessage('Google authentication cancelled by user', 'warning');
    });

    window.electronAPI.receiveRestoreSession((data) => {
        updateAuthStatus(data.authenticated, data.userName);
        if (data.spreadsheetId) {
            spreadsheetIdInput.value = data.spreadsheetId;
            logMessage(`Session restored with spreadsheet: ${data.spreadsheetId}`, 'success');
        }
    });

    window.electronAPI.receiveGoogleAuthSuccess((data) => {
        updateStatusMessage('Successfully authenticated with Google!');
        updateAuthStatus(true, data.userName);
        setButtonState(loginBtn, false);
        logMessage(`Authentication successful for user: ${data.userName}`, 'success');
    });

    window.electronAPI.receiveGoogleAuthError((message) => {
        updateStatusMessage(`Authentication error: ${message}`);
        setButtonState(loginBtn, false);
        logMessage(`Authentication failed: ${message}`, 'error');
    });

    window.electronAPI.receiveLogoutComplete(() => {
        updateStatusMessage('Successfully signed out');
        updateAuthStatus(false);
        setButtonState(logoutBtn, false);
    });

    window.electronAPI.receiveAuthExpired(() => {
        updateStatusMessage('Your session has expired. Please sign in again.');
        updateAuthStatus(false);
        logMessage('Session expired, user needs to re-authenticate', 'warning');
    });

    window.electronAPI.receiveTestAccessSuccess((data) => {
        updateStatusMessage('Spreadsheet access verified successfully!');
        setButtonState(testBtn, false);
        hideAllInfoBoxes();
        
        spreadsheetInfo.innerHTML = `
            <h4>✓ Spreadsheet Connected</h4>
            <div><strong>Title:</strong> ${data.title}</div>
            <div><strong>Available Sheets:</strong> ${data.sheets.map(sheet => sheet.title).join(', ')}</div>
        `;
        spreadsheetInfo.style.display = 'block';

        if (data.potentialMoves && data.potentialMoves.length > 0) {
            pendingMoves = data.potentialMoves;
            let movesHtml = data.potentialMoves.map(move => `
                <div class="move-item">
                    <div class="student-name">${move.studentName}</div>
                    <div class="move-details">Age ${move.age} • From: ${move.currentSheet} → To: ${move.newSheet}</div>
                </div>
            `).join('');
            
            movesList.innerHTML = movesHtml;
            movesPreview.style.display = 'block';
            processBtn.style.display = 'inline-block';
            setButtonState(processBtn, false);
            updateStatusMessage(`Found ${data.potentialMoves.length} student(s) ready to move`);
            logMessage(`Found ${data.potentialMoves.length} students needing to be moved`, 'info');
        } else {
            pendingMoves = [];
            processBtn.style.display = 'none';
            updateStatusMessage('All students are in their correct age groups!');
            logMessage('No student moves required - all students in correct groups', 'success');
        }
    });
    
    window.electronAPI.receiveTestAccessExcelDetected(() => {
        updateStatusMessage('Excel format detected - please convert to Google Sheets');
        setButtonState(testBtn, false);
        hideAllInfoBoxes();
        warningInfo.innerHTML = `
            <h4>⚠️ Excel Format Detected</h4>
            <div>This appears to be an Excel file. Please convert it to Google Sheets format:</div>
            <ol style="margin-left: 20px; margin-top: 8px;">
                <li>Open the file in Google Sheets</li>
                <li>Go to File → Save as Google Sheets</li>
                <li>Use the new Google Sheets URL</li>
            </ol>
        `;
        warningInfo.style.display = 'block';
        logMessage('Excel format detected, conversion required', 'warning');
    });

    window.electronAPI.receiveTestAccessError((message) => {
        updateStatusMessage('Failed to access spreadsheet');
        setButtonState(testBtn, false);
        hideAllInfoBoxes();
        errorInfo.innerHTML = `<h4>Access Error</h4><div>${message}</div>`;
        errorInfo.style.display = 'block';
        logMessage(`Spreadsheet access error: ${message}`, 'error');
    });

    window.electronAPI.receiveProcessingComplete((message) => {
        updateStatusMessage('Processing completed successfully!');
        processBtn.style.display = 'none';
        setButtonState(testBtn, false);
        hideAllInfoBoxes();
        
        const input = spreadsheetIdInput.value.trim();
        const spreadsheetId = extractSpreadsheetId(input) || input;
        const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/`;

        postProcessingInfo.innerHTML = `
            <h4>✅ Processing Complete</h4>
            <div style="margin-bottom: 16px;">${message}</div>
            <div><strong>Need to undo?</strong> Use Google Sheets' Version History:</div>
            <ol style="margin-left: 20px; margin-top: 8px;">
                <li><a href="${spreadsheetUrl}" target="_blank" style="color: #3b82f6;">Open your Google Sheet</a></li>
                <li>Go to <strong>File → Version history → See version history</strong></li>
                <li>Select the version from before the changes</li>
                <li>Click <strong>"Restore this version"</strong></li>
            </ol>
        `;
        postProcessingInfo.style.display = 'block';
        logMessage(`Processing completed: ${message}`, 'success');
    });

    window.electronAPI.receiveProcessingError((message) => {
        updateStatusMessage('Processing failed');
        processBtn.style.display = 'inline-block';
        setButtonState(processBtn, false);
        setButtonState(testBtn, false);
        hideAllInfoBoxes();
        errorInfo.innerHTML = `<h4>Processing Error</h4><div>${message}</div>`;
        errorInfo.style.display = 'block';
        logMessage(`Processing failed: ${message}`, 'error');
    });

    // Initialize logs
    logMessage('Application initialized', 'info');
});