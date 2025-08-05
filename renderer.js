document.addEventListener('DOMContentLoaded', () => {
    // --- Element Constants ---
    const loginBtn = document.getElementById('login-btn');
    const logoutBtn = document.getElementById('logout-btn');
    const testBtn = document.getElementById('test-btn');
    const processBtn = document.getElementById('process-btn');
    const statusDiv = document.getElementById('status');
    const codeInputSection = document.getElementById('code-input-section');
    const spreadsheetSection = document.getElementById('spreadsheet-section');
    const authCodeInput = document.getElementById('auth-code-input');
    const spreadsheetIdInput = document.getElementById('spreadsheet-id-input');
    const submitCodeBtn = document.getElementById('submit-code-btn');
    const spreadsheetInfo = document.getElementById('spreadsheet-info');
    const warningInfo = document.getElementById('warning-info');
    const errorInfo = document.getElementById('error-info');
    const postProcessingInfo = document.getElementById('post-processing-info');
    const movesPreview = document.getElementById('moves-preview');
    const movesList = document.getElementById('moves-list');
    const userStatus = document.getElementById('user-status');
    
    let pendingMoves = [];

    // --- Utility Functions ---
    function hideAllInfoBoxes() {
        spreadsheetInfo.style.display = 'none';
        warningInfo.style.display = 'none';
        errorInfo.style.display = 'none';
        movesPreview.style.display = 'none';
        postProcessingInfo.style.display = 'none';
    }

    function updateUI(authenticated, spreadsheetId = '') {
        if (authenticated) {
            loginBtn.style.display = 'none';
            logoutBtn.style.display = 'inline-block';
            codeInputSection.style.display = 'none';
            spreadsheetSection.style.display = 'block';
            userStatus.textContent = '✓ Logged in';
            statusDiv.textContent = 'Ready to work with spreadsheets!';
            if (spreadsheetId) {
                spreadsheetIdInput.value = spreadsheetId;
            }
        } else {
            loginBtn.style.display = 'inline-block';
            logoutBtn.style.display = 'none';
            spreadsheetSection.style.display = 'none';
            userStatus.textContent = '';
            authCodeInput.value = '';
            hideAllInfoBoxes();
            statusDiv.textContent = 'Please log in to continue.';
        }
    }
    
    // --- Event Listeners ---
    spreadsheetIdInput.addEventListener('input', () => {
        window.electronAPI.saveSpreadsheetId(spreadsheetIdInput.value.trim());
    });

    loginBtn.addEventListener('click', () => {
        statusDiv.textContent = 'Opening Google login...';
        codeInputSection.style.display = 'block';
        loginBtn.style.display = 'none';
        window.electronAPI.loginWithGoogle();
    });

    logoutBtn.addEventListener('click', () => {
        statusDiv.textContent = 'Logging out...';
        window.electronAPI.logout();
    });

    submitCodeBtn.addEventListener('click', () => {
        const authCode = authCodeInput.value.trim();
        if (!authCode) {
            statusDiv.textContent = 'Please enter the authorization code.';
            return;
        }
        statusDiv.textContent = 'Processing authorization code...';
        submitCodeBtn.disabled = true;
        window.electronAPI.submitAuthCode(authCode);
    });

    authCodeInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') submitCodeBtn.click();
    });

    testBtn.addEventListener('click', () => {
        const spreadsheetId = spreadsheetIdInput.value.trim();
        if (!spreadsheetId) {
            statusDiv.textContent = 'Please enter a spreadsheet ID.';
            return;
        }
        statusDiv.textContent = 'Testing spreadsheet access...';
        hideAllInfoBoxes();
        processBtn.style.display = 'none';
        testBtn.disabled = true;
        window.electronAPI.testSpreadsheetAccess(spreadsheetId);
    });

    processBtn.addEventListener('click', () => {
        const spreadsheetId = spreadsheetIdInput.value.trim();
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
        updateUI(data.authenticated, data.spreadsheetId);
    });

    window.electronAPI.receiveGoogleAuthSuccess((message) => {
        statusDiv.textContent = message;
        updateUI(true, spreadsheetIdInput.value);
        submitCodeBtn.disabled = false;
    });

    window.electronAPI.receiveGoogleAuthError((message) => {
        statusDiv.textContent = `Error: ${message}`;
        submitCodeBtn.disabled = false;
        loginBtn.style.display = 'block';
        codeInputSection.style.display = 'none';
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
        
        const spreadsheetId = spreadsheetIdInput.value.trim();
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