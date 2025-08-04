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
    const sessionIndicator = document.getElementById('session-indicator');
    const userStatus = document.getElementById('user-status');

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
        testBtn.disabled = true;
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
            sessionIndicator.textContent = '● Online';
            sessionIndicator.className = 'session-indicator';
            sessionIndicator.style.display = 'block';
            
            // Restore spreadsheet ID if provided
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
            sessionIndicator.textContent = '● Offline';
            sessionIndicator.className = 'session-indicator offline';
            sessionIndicator.style.display = 'block';
            
            // Clear sensitive data from UI
            authCodeInput.value = '';
        }
    }

    // Handle session restoration
    window.electronAPI.receiveRestoreSession((data) => {
        console.log('Restoring session:', data);
        updateUI(data.authenticated, data.spreadsheetId);
        
        if (data.authenticated) {
            statusDiv.textContent = 'Welcome back! Your previous session has been restored.';
            // Show a brief welcome message, then return to normal status
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
            <strong>Type:</strong> ${data.type === 'google-sheets' ? 'Google Sheets' : 'Excel File'}<br>
            <strong>Available Sheets:</strong><br>
            <ul>
                ${data.sheets.map(sheet => `<li>${sheet.title} (ID: ${sheet.id})</li>`).join('')}
            </ul>
        `;
        spreadsheetInfo.style.display = 'block';
    });

    window.electronAPI.receiveTestAccessError((message) => {
        statusDiv.textContent = 'Spreadsheet access test failed.';
        testBtn.disabled = false;
        spreadsheetInfo.style.display = 'none';
        
        errorDetails.textContent = message;
        errorInfo.style.display = 'block';
    });

    window.electronAPI.receiveProcessingComplete((message) => {
        statusDiv.textContent = 'Processing completed successfully!';
        processBtn.disabled = false;
        
        // Show results in a formatted way
        const lines = message.split('\n');
        if (lines.length > 1) {
            // Multi-line message (like Excel analysis)
            statusDiv.innerHTML = lines.map(line => line.trim()).filter(line => line).join('<br>');
        } else {
            statusDiv.textContent = message;
        }
    });

});