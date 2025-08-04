document.addEventListener('DOMContentLoaded', () => {
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
    const spreadsheetDetails = document.getElementById('spreadsheet-details');
    const errorInfo = document.getElementById('error-info');
    const errorDetails = document.getElementById('error-details');
    const authSection = document.getElementById('auth-section');
    const userStatus = document.getElementById('user-status');
    const proposedMovesDiv = document.getElementById('proposed-moves');

    let isAuthenticated = false;

    // Auto-save spreadsheet ID when it changes
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

    // Handle Enter key in the input field
    authCodeInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            submitCodeBtn.click();
        }
    });

    testBtn.addEventListener('click', () => {
        const spreadsheetId = spreadsheetIdInput.value.trim();
        if (!spreadsheetId) {
            statusDiv.textContent = 'Please enter a spreadsheet ID.';
            return;
        }
        statusDiv.textContent = 'Testing spreadsheet access...';
        spreadsheetInfo.style.display = 'none';
        errorInfo.style.display = 'none';
        proposedMovesDiv.innerHTML = ''; // Clear previous moves
        testBtn.disabled = true;
        processBtn.disabled = true; // Disable until test is successful
        window.electronAPI.testSpreadsheetAccess(spreadsheetId);
    });

    processBtn.addEventListener('click', () => {
        const spreadsheetId = spreadsheetIdInput.value.trim();
        if (!spreadsheetId) {
            statusDiv.textContent = 'Please enter a spreadsheet ID.';
            return;
        }
        statusDiv.textContent = 'Processing spreadsheet... This may take a moment.';
        processBtn.disabled = true;
        testBtn.disabled = true;
        window.electronAPI.processSpreadsheet(spreadsheetId);
    });

    // Function to update UI based on authentication status
    function updateUI(authenticated, spreadsheetId = '') {
        isAuthenticated = authenticated;
        
        if (authenticated) {
            statusDiv.textContent = 'Ready to work with spreadsheets!';
            loginBtn.style.display = 'none';
            logoutBtn.style.display = 'inline-block';
            codeInputSection.style.display = 'none';
            spreadsheetSection.style.display = 'block';
            userStatus.textContent = '✓ Logged in';
            processBtn.disabled = true; // Should be disabled initially
            
            if (spreadsheetId) {
                spreadsheetIdInput.value = spreadsheetId;
            }
        } else {
            statusDiv.textContent = 'Please log in to continue.';
            loginBtn.style.display = 'inline-block';
            logoutBtn.style.display = 'none';
            codeInputSection.style.display = 'none';
            spreadsheetSection.style.display = 'none';
            spreadsheetInfo.style.display = 'none';
            errorInfo.style.display = 'none';
            userStatus.textContent = '';
            
            authCodeInput.value = '';
        }
    }

    // Handle session restoration
    window.electronAPI.receiveRestoreSession((data) => {
        console.log('Restoring session:', data);
        updateUI(data.authenticated, data.spreadsheetId);
        
        if (data.authenticated) {
            statusDiv.textContent = 'Welcome back! Your previous session has been restored.';
            setTimeout(() => {
                statusDiv.textContent = 'Ready to work with spreadsheets!';
            }, 3000);
        }
    });

    // Handle successful authentication
    window.electronAPI.receiveGoogleAuthSuccess((message) => {
        statusDiv.textContent = message;
        updateUI(true, spreadsheetIdInput.value);
    });

    // Handle authentication errors
    window.electronAPI.receiveGoogleAuthError((message) => {
        statusDiv.textContent = `Error: ${message}`;
        submitCodeBtn.disabled = false;
        loginBtn.style.display = 'block';
        codeInputSection.style.display = 'none';
    });

    // Handle logout completion
    window.electronAPI.receiveLogoutComplete(() => {
        statusDiv.textContent = 'Successfully logged out.';
        updateUI(false);
    });

    // Handle expired authentication
    window.electronAPI.receiveAuthExpired(() => {
        statusDiv.textContent = 'Your session has expired. Please log in again.';
        updateUI(false);
    });

    // Handle test results
    window.electronAPI.receiveTestAccessSuccess((data) => {
        statusDiv.textContent = 'Spreadsheet access test successful!';
        testBtn.disabled = false;
        errorInfo.style.display = 'none';
        
        spreadsheetDetails.innerHTML = `
            <strong>Title:</strong> ${data.title}<br>
            <strong>Available Sheets:</strong>
            <ul>
                ${data.sheets.map(sheet => `<li>${sheet.title} (ID: ${sheet.id})</li>`).join('')}
            </ul>
        `;
        
        // Display proposed moves
        proposedMovesDiv.innerHTML = ''; // Clear previous results
        if (data.potentialMoves && data.potentialMoves.length > 0) {
            let movesHtml = '<h4>Proposed Changes:</h4><ul>';
            data.potentialMoves.forEach(move => {
                movesHtml += `<li>${move.studentName} (Age ${move.age}) will be moved from <strong>${move.currentSheet}</strong> to <strong>${move.newSheet}</strong></li>`;
            });
            movesHtml += '</ul>';
            proposedMovesDiv.innerHTML = movesHtml;
            processBtn.disabled = false; // Enable processing
        } else {
            proposedMovesDiv.innerHTML = '<h4>No students need to be moved at this time.</h4>';
            processBtn.disabled = true; // Keep disabled
        }

        spreadsheetInfo.style.display = 'block';
    });

    window.electronAPI.receiveTestAccessError((message) => {
        statusDiv.textContent = 'Spreadsheet access test failed.';
        testBtn.disabled = false;
        spreadsheetInfo.style.display = 'none';
        
        if (message.includes('Unsupported file type')) {
            errorDetails.textContent = 'Excel Format Detected. Please convert to a Google Sheet (File > Save as Google Sheets) and use the new Sheet ID.';
        } else {
            errorDetails.textContent = message;
        }
        errorInfo.style.display = 'block';
    });

    window.electronAPI.receiveProcessingComplete((message) => {
    // Show results in a formatted way, allowing for HTML content
    const lines = message.split('\n');
    statusDiv.innerHTML = lines.map(line => line.trim()).filter(line => line).join('<br>');
    
    processBtn.disabled = true; // Disable after processing
    testBtn.disabled = false;
    spreadsheetInfo.style.display = 'none'; // Hide info to encourage a new test
    });

});