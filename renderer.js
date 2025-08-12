document.addEventListener('DOMContentLoaded', () => {
    // --- Element Constants ---
    const loginBtn = document.getElementById('login-btn');
    const logoutBtn = document.getElementById('logout-btn');
    const testBtn = document.getElementById('test-btn');
    const processBtn = document.getElementById('process-btn');
    const statusDiv = document.getElementById('status');
    const loginSection = document.getElementById('login-section');
    const authenticatedSection = document.getElementById('authenticated-section');
    const spreadsheetIdInput = document.getElementById('spreadsheet-id-input');
    const spreadsheetInfo = document.getElementById('spreadsheet-info');
    const warningInfo = document.getElementById('warning-info');
    const errorInfo = document.getElementById('error-info');
    const postProcessingInfo = document.getElementById('post-processing-info');
    const movesPreview = document.getElementById('moves-preview');
    const movesList = document.getElementById('moves-list');
    const userStatus = document.getElementById('user-status');
    const loginInstructions = document.getElementById('login-instructions');
    const username = document.getElementById('username');
    
    let pendingMoves = [];

    // --- Utility Functions ---
    function hideAllInfoBoxes() {
        spreadsheetInfo.style.display = 'none';
        warningInfo.style.display = 'none';
        errorInfo.style.display = 'none';
        movesPreview.style.display = 'none';
        postProcessingInfo.style.display = 'none';
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

    function updateUI(authenticated, spreadsheetId = '', userEmail = '') {
        if (authenticated) {
            loginSection.style.display = 'none';
            authenticatedSection.style.display = 'block';
            userStatus.textContent = '✓ Logged in';
            statusDiv.textContent = 'Ready to work with spreadsheets!';
            
            // Update username display
            if (userEmail) {
                username.textContent = userEmail;
            }
            
            if (spreadsheetId) {
                spreadsheetIdInput.value = spreadsheetId;
            }
        } else {
            loginSection.style.display = 'block';
            authenticatedSection.style.display = 'none';
            loginInstructions.style.display = 'none';
            userStatus.textContent = '';
            hideAllInfoBoxes();
            statusDiv.textContent = 'Please log in to continue.';
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
        statusDiv.textContent = 'Opening Google login in your browser...';
        loginBtn.disabled = true;
        loginInstructions.style.display = 'block';
        window.electronAPI.loginWithGoogle();
    });

    logoutBtn.addEventListener('click', () => {
        statusDiv.textContent = 'Logging out...';
        window.electronAPI.logout();
    });

    testBtn.addEventListener('click', () => {
        const input = spreadsheetIdInput.value.trim();
        const spreadsheetId = extractSpreadsheetId(input) || input;
        
        if (!spreadsheetId) {
            statusDiv.textContent = 'Please enter a valid spreadsheet URL or ID.';
            return;
        }
        
        // Update the input field to show the extracted ID
        if (spreadsheetId !== input) {
            spreadsheetIdInput.value = spreadsheetId;
        }
        
        statusDiv.textContent = 'Testing spreadsheet access...';
        hideAllInfoBoxes();
        processBtn.style.display = 'none';
        testBtn.disabled = true;
        window.electronAPI.testSpreadsheetAccess(spreadsheetId);
    });

    processBtn.addEventListener('click', () => {
        const input = spreadsheetIdInput.value.trim();
        const spreadsheetId = extractSpreadsheetId(input) || input;
        
        if (pendingMoves.length === 0) {
            statusDiv.textContent = 'No moves to process.';
            return;
        }
        statusDiv.textContent = 'Processing spreadsheet... This may take a moment.';
        processBtn.disabled = true;
        testBtn.disabled = true;
        window.electronAPI.processSpreadsheet(spreadsheetId);
    });
    
    // --- Electron API Listeners ---
    window.electronAPI.receiveRestoreSession((data) => {
        updateUI(data.authenticated, data.spreadsheetId, data.userEmail);
    });

    window.electronAPI.receiveGoogleAuthSuccess((message) => {
        statusDiv.textContent = message;
        updateUI(true, spreadsheetIdInput.value);
        loginBtn.disabled = false;
    });

    window.electronAPI.receiveGoogleAuthError((message) => {
        statusDiv.textContent = `Error: ${message}`;
        loginBtn.disabled = false;
        loginInstructions.style.display = 'none';
    });

    window.electronAPI.receiveLogoutComplete(() => {
        statusDiv.textContent = 'Successfully logged out.';
        updateUI(false);
    });

    window.electronAPI.receiveAuthExpired(() => {
        statusDiv.textContent = 'Your session has expired. Please log in again.';
        updateUI(false);
    });

    window.electronAPI.receiveTestAccessSuccess((data) => {
        statusDiv.textContent = 'Spreadsheet access test successful!';
        testBtn.disabled = false;
        hideAllInfoBoxes();
        
        spreadsheetInfo.innerHTML = `
            <strong>Title:</strong> ${data.title}<br>
            <strong>Available Sheets:</strong>
            <ul>${data.sheets.map(sheet => `<li>${sheet.title}</li>`).join('')}</ul>
        `;
        spreadsheetInfo.style.display = 'block';

        if (data.potentialMoves && data.potentialMoves.length > 0) {
            pendingMoves = data.potentialMoves;
            let movesHtml = data.potentialMoves.map(move => `
                <div class="move-item">
                    <strong>${move.studentName}</strong> (Age ${move.age})<br>
                    From: <em>${move.currentSheet}</em> → To: <em>${move.newSheet}</em>
                </div>
            `).join('');
            
            movesList.innerHTML = movesHtml;
            movesPreview.style.display = 'block';
            processBtn.style.display = 'inline-block';
            processBtn.disabled = false; 
            statusDiv.innerHTML = `Found <strong>${data.potentialMoves.length}</strong> student(s) to move. Ready to process.`;
        } else {
            pendingMoves = [];
            processBtn.style.display = 'none';
            statusDiv.textContent = 'All students are in their correct age groups!';
        }
    });
    
    window.electronAPI.receiveTestAccessExcelDetected(() => {
        statusDiv.textContent = 'Excel format detected.';
        testBtn.disabled = false;
        hideAllInfoBoxes();
        warningInfo.innerHTML = '<h4>⚠️ Excel Format Detected</h4><div>Please convert the file to Google Sheets format (File > Save as Google Sheets) and use the new Sheet ID.</div>';
        warningInfo.style.display = 'block';
    });

    window.electronAPI.receiveTestAccessError((message) => {
        statusDiv.textContent = 'Spreadsheet access test failed.';
        testBtn.disabled = false;
        hideAllInfoBoxes();
        errorInfo.innerHTML = `<h4>Error:</h4><div>${message}</div>`;
        errorInfo.style.display = 'block';
    });

    window.electronAPI.receiveProcessingComplete((message) => {
        statusDiv.innerHTML = `<strong>Processing completed successfully!</strong>`;
        processBtn.style.display = 'none';
        testBtn.disabled = false;
        hideAllInfoBoxes();
        
        const input = spreadsheetIdInput.value.trim();
        const spreadsheetId = extractSpreadsheetId(input) || input;
        const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/`;

        postProcessingInfo.innerHTML = `
            <h4>Undo Instructions</h4>
            <p>${message}</p>
            <p><strong>To revert this change, use Google Sheets' built-in "Version History" feature.</strong></p>
            <ol style="text-align: left; padding-left: 25px;">
                <li><a href="${spreadsheetUrl}" target="_blank">Click here to open your Google Sheet.</a></li>
                <li>In the Google Sheet, go to <strong>File > Version history > See version history</strong>.</li>
                <li>Select the version from just before the changes were made.</li>
                <li>Click the green "Restore this version" button at the top.</li>
            </ol>
        `;
        postProcessingInfo.style.display = 'block';
    });

    window.electronAPI.receiveProcessingError((message) => {
        statusDiv.textContent = 'Processing failed.';
        processBtn.style.display = 'inline-block';
        processBtn.disabled = false;
        testBtn.disabled = false;
        hideAllInfoBoxes();
        errorInfo.innerHTML = `<h4>Error:</h4><div>${message}</div>`;
        errorInfo.style.display = 'block';
    });
});