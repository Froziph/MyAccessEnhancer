document.addEventListener('DOMContentLoaded', async function() {
    const apiKeyInput = document.getElementById('openai-key');
    const saveBtn = document.getElementById('save-btn');
    const clearBtn = document.getElementById('clear-btn');
    const toggleVisibilityBtn = document.getElementById('toggle-visibility');
    const statusIndicator = document.getElementById('status-indicator');
    const statusText = document.getElementById('status-text');
    const message = document.getElementById('message');

    // Load existing API key
    await loadSettings();

    // Save settings
    saveBtn.addEventListener('click', async function() {
        const apiKey = apiKeyInput.value.trim();
        
        if (!apiKey) {
            showMessage('Please enter an API key', 'error');
            return;
        }

        if (!isValidOpenAIKey(apiKey)) {
            showMessage('Invalid API key format. Key should start with "sk-"', 'error');
            return;
        }

        try {
            await chrome.storage.sync.set({ 'openai_api_key': apiKey });
            await updateStatus();
            showMessage('Settings saved successfully!', 'success');
            
            // Notify content scripts about the key update
            notifyContentScripts();
        } catch (error) {
            showMessage('Failed to save settings: ' + error.message, 'error');
        }
    });

    // Clear API key
    clearBtn.addEventListener('click', async function() {
        if (confirm('Are you sure you want to clear your API key?')) {
            try {
                await chrome.storage.sync.remove('openai_api_key');
                apiKeyInput.value = '';
                await updateStatus();
                showMessage('API key cleared', 'info');
                
                // Notify content scripts about the key removal
                notifyContentScripts();
            } catch (error) {
                showMessage('Failed to clear settings: ' + error.message, 'error');
            }
        }
    });

    // Toggle password visibility
    toggleVisibilityBtn.addEventListener('click', function() {
        if (apiKeyInput.type === 'password') {
            apiKeyInput.type = 'text';
            toggleVisibilityBtn.textContent = '🙈';
        } else {
            apiKeyInput.type = 'password';
            toggleVisibilityBtn.textContent = '👁️';
        }
    });

    // Real-time validation
    apiKeyInput.addEventListener('input', function() {
        const apiKey = this.value.trim();
        if (apiKey && !isValidOpenAIKey(apiKey)) {
            this.style.borderColor = '#d83b01';
        } else {
            this.style.borderColor = '';
        }
    });

    // Load settings from storage
    async function loadSettings() {
        try {
            const result = await chrome.storage.sync.get('openai_api_key');
            if (result.openai_api_key) {
                apiKeyInput.value = result.openai_api_key;
            }
            await updateStatus();
        } catch (error) {
            showMessage('Failed to load settings: ' + error.message, 'error');
        }
    }

    // Update status indicator
    async function updateStatus() {
        try {
            const result = await chrome.storage.sync.get('openai_api_key');
            const hasKey = result.openai_api_key && result.openai_api_key.trim().length > 0;
            
            if (hasKey) {
                statusIndicator.className = 'status-indicator status-active';
                statusText.textContent = 'API key configured ✓';
            } else {
                statusIndicator.className = 'status-indicator status-inactive';
                statusText.textContent = 'API key not configured';
            }
        } catch (error) {
            statusIndicator.className = 'status-indicator status-error';
            statusText.textContent = 'Error checking status';
        }
    }

    // Show message to user
    function showMessage(text, type = 'info') {
        message.textContent = text;
        message.className = `message message-${type}`;
        message.classList.remove('hidden');
        
        setTimeout(() => {
            message.classList.add('hidden');
        }, 3000);
    }

    // Validate OpenAI API key format
    function isValidOpenAIKey(key) {
        // OpenAI keys typically start with "sk-" and are 51+ characters
        return key.startsWith('sk-') && key.length >= 20;
    }

    // Notify content scripts about API key changes
    async function notifyContentScripts() {
        try {
            const tabs = await chrome.tabs.query({
                url: ["https://myaccess.microsoft.com/*", "https://myapplications.microsoft.com/*"]
            });
            
            for (const tab of tabs) {
                try {
                    await chrome.tabs.sendMessage(tab.id, {
                        type: 'API_KEY_UPDATED'
                    });
                } catch (e) {
                    // Tab might not have content script loaded, ignore
                }
            }
        } catch (error) {
            // Ignore errors when notifying content scripts
        }
    }
});