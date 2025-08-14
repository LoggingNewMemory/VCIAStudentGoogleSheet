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
    
    let pendingMoves = [];
    let authInProgress = false;

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
        } else {
            authStatus.textContent = 'Not Logged In';
            authStatus.className = 'status-indicator unauthenticated';
            loginSection.style.display = 'block';
            authenticatedSection.style.display = 'none';
            username.textContent = 'User';
            hideAllInfoBoxes();
            updateStatusMessage('Please sign in to continue');
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

    function showAuthInstructions() {
        updateStatusMessage('Opening your default browser for secure sign-in...');
        hideAllInfoBoxes();
        
        warningInfo.innerHTML = `
            <h4>🌐 Browser Authentication</h4>
            <div style="margin-bottom: 12px;">Your default browser should now be opening for secure Google sign-in.</div>
            <div style="margin-bottom: 12px;"><strong>If the browser didn't open:</strong></div>
            <ol style="margin-left: 20px; margin-bottom: 12px;">
                <li>Check if a new browser tab/window opened</li>
                <li>Look for any browser permission prompts</li>
                <li>Make sure your default browser is set correctly</li>
            </ol>
            <div style="margin-bottom: 12px;"><strong>After signing in:</strong></div>
            <ul style="margin-left: 20px;">
                <li>Complete the Google authorization process</li>
                <li>You'll see a success message in the browser</li>
                <li>Return to this application</li>
            </ul>
        `;
        warningInfo.style.display = 'block';
    }

    function hideAuthInstructions() {
        if (warningInfo.innerHTML.includes('Browser Authentication')) {
            warningInfo.style.display = 'none';
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
            }
        }, 10);
    });

    loginBtn.addEventListener('click', () => {
        if (authInProgress) return; // Prevent multiple clicks
        
        authInProgress = true;
        setButtonState(loginBtn, true, 'Opening Browser...');
        showAuthInstructions();
        window.electronAPI.loginWithGoogle();
    });

    logoutBtn.addEventListener('click', () => {
        updateStatusMessage('Signing out...');
        setButtonState(logoutBtn, true, 'Signing Out...');
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
        window.electronAPI.processSpreadsheet(spreadsheetId);
    });
    
    // --- Electron API Event Handlers ---

    window.electronAPI.receiveGoogleAuthCancelled(() => {
        authInProgress = false;
        updateStatusMessage('Sign-in was cancelled. You can try again.');
        setButtonState(loginBtn, false);
        hideAuthInstructions();
    });

    window.electronAPI.receiveRestoreSession((data) => {
        updateAuthStatus(data.authenticated, data.userName);
        if (data.spreadsheetId) {
            spreadsheetIdInput.value = data.spreadsheetId;
        }
    });

    window.electronAPI.receiveGoogleAuthSuccess((data) => {
        authInProgress = false;
        updateStatusMessage('Successfully authenticated with Google!');
        updateAuthStatus(true, data.userName);
        setButtonState(loginBtn, false);
        hideAuthInstructions();
    });

    window.electronAPI.receiveGoogleAuthError((message) => {
        authInProgress = false;
        updateStatusMessage(`Authentication failed: ${message}`);
        setButtonState(loginBtn, false);
        hideAuthInstructions();
        
        // Show specific error information
        errorInfo.innerHTML = `
            <h4>Authentication Error</h4>
            <div style="margin-bottom: 12px;">${message}</div>
            <div><strong>Common solutions:</strong></div>
            <ul style="margin-left: 20px;">
                <li>Make sure port 8080 is not blocked by firewall</li>
                <li>Close any other applications using port 8080</li>
                <li>Check your default browser settings</li>
                <li>Try again after a few moments</li>
            </ul>
        `;
        errorInfo.style.display = 'block';
    });

    window.electronAPI.receiveLogoutComplete(() => {
        updateStatusMessage('Successfully signed out');
        updateAuthStatus(false);
        setButtonState(logoutBtn, false);
    });

    window.electronAPI.receiveAuthExpired(() => {
        updateStatusMessage('Your session has expired. Please sign in again.');
        updateAuthStatus(false);
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
        } else {
            pendingMoves = [];
            processBtn.style.display = 'none';
            updateStatusMessage('All students are in their correct age groups!');
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
    });

    window.electronAPI.receiveTestAccessError((message) => {
        updateStatusMessage('Failed to access spreadsheet');
        setButtonState(testBtn, false);
        hideAllInfoBoxes();
        errorInfo.innerHTML = `<h4>Access Error</h4><div>${message}</div>`;
        errorInfo.style.display = 'block';
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
    });

    window.electronAPI.receiveProcessingError((message) => {
        updateStatusMessage('Processing failed');
        processBtn.style.display = 'inline-block';
        setButtonState(processBtn, false);
        setButtonState(testBtn, false);
        hideAllInfoBoxes();
        errorInfo.innerHTML = `<h4>Processing Error</h4><div>${message}</div>`;
        errorInfo.style.display = 'block';
    });
});