// MyAccess Package approvers Tab Enhanced - Chrome Extension Version

// ========================================
// AI PROMPTS CONFIGURATION
// ========================================

const AI_PROMPTS = {
    // Default prompt for business justification enhancement
    ENHANCE_BUSINESS_JUSTIFICATION: {
        systemPrompt: `You are a SimCorp access request specialist. Transform business justifications into the proper SimCorp format using the required structure.

REQUIRED FORMAT:
As a (role and team), I need access so that (justification). I have read the policies and understand the implications of having such access.

GOOD EXAMPLES FROM DOCUMENTATION:
- As an SRE team operator, I need access so that I can monitor system health (Ticket #12345). I have read the policies and understand the implications of having such access.
- As an SRE team reader for the DEP system, I need access so that I can review logs and reports. I have read the policies and understand the implications of having such access.
- As a DevOps Engineer of Straw Hats, I need access so that I can resolve an issue related to KV password change (Alert #2376, Ticket #54321). I have read the policies and understand the implications of having such access.
- As a Product Owner of Cockpit, I need access so I will be the new approver for DEP packages, as approved by my manager (manager's initials). I have read the policies and understand the implications of having such access.

IMPORTANT: Do not include quotation marks around the output.

RULES:
- NEVER add information not in the original text
- Use only details provided in the original request
- If role/team is unclear from context, use generic terms like "team member" or "engineer"
- Preserve all ticket numbers, system names, and specific details exactly
- Always end with the required policy acknowledgment`,

        userPromptTemplate: `Transform this into proper SimCorp business justification format. Use only the information provided - do not add details not mentioned:

Original: "{originalText}"

SimCorp format:`,

        maxTokens: 300,
        model: "gpt-4o-mini"
    },

    // Alternative prompt for minimal SimCorp format
    ENHANCE_CONCISE: {
        systemPrompt: `You are a SimCorp specialist creating minimal but compliant justifications.

FORMAT: As a (role), I need access so that (essential purpose). I have read the policies and understand the implications of having such access.

RULES:
- Use the shortest possible language while meeting requirements
- NEVER add information not provided
- Keep to absolute essentials only
- Preserve ticket numbers and system names exactly
- Do not include quotation marks around the output`,

        userPromptTemplate: `Create the most concise SimCorp format justification using only the provided information:

"{originalText}"

Minimal SimCorp format:`,

        maxTokens: 500,
        model: "gpt-4o-mini"
    },

    // Alternative prompt for detailed SimCorp format
    ENHANCE_DETAILED: {
        systemPrompt: `You are a SimCorp specialist creating detailed but appropriate justifications.

FORMAT: As a (role and team), I need access so that (detailed justification with context). I have read the policies and understand the implications of having such access.

APPROACH:
- Include role and team if mentioned
- Provide clear context for the access need
- Reference tickets, alerts, or user stories if provided
- Answer: Why this user? Why this access? What purpose?
- NEVER add information not in original text
- Keep to one well-structured sentence plus policy statement
- Do not include quotation marks around the output`,

        userPromptTemplate: `Create a detailed SimCorp format justification with appropriate context using only the provided information:

"{originalText}"

Detailed SimCorp format:`,

        maxTokens: 1500,
        model: "gpt-4o-mini"
    }
};

// ========================================
// AI ENHANCEMENT FUNCTIONALITY
// ========================================

// Check if extension context is valid
function isExtensionContextValid() {
    try {
        return !!(chrome && chrome.runtime && chrome.runtime.id);
    } catch (error) {
        return false;
    }
}

// Get OpenAI API key from storage
async function getOpenAIKey() {
    try {
        // Check if extension context is valid
        if (!isExtensionContextValid()) {
            return null;
        }
        const result = await chrome.storage.sync.get('openai_api_key');
        return result.openai_api_key || null;
    } catch (error) {
        // Extension context invalidated or other error
        return null;
    }
}

// Check if API key is configured
async function hasOpenAIKey() {
    const key = await getOpenAIKey();
    return key && key.trim().length > 0;
}

// Enhance business justification using OpenAI API
async function enhanceBusinessJustification(text, apiKey) {
    if (!text || !text.trim()) {
        throw new Error('No text provided');
    }

    if (!apiKey) {
        throw new Error('OpenAI API key not configured');
    }

    const prompt = AI_PROMPTS.ENHANCE_BUSINESS_JUSTIFICATION;
    
    try {
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: prompt.model,
                messages: [
                    {
                        role: 'system',
                        content: prompt.systemPrompt
                    },
                    {
                        role: 'user',
                        content: prompt.userPromptTemplate.replace('{originalText}', text)
                    }
                ],
                max_tokens: prompt.maxTokens,
                temperature: 0.3
            })
        });

        if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            throw new Error(`OpenAI API error: ${response.status} - ${errorData.error?.message || 'Unknown error'}`);
        }

        const data = await response.json();
        
        if (!data.choices || !data.choices[0] || !data.choices[0].message) {
            throw new Error('Invalid response from OpenAI API');
        }

        return data.choices[0].message.content.trim();

    } catch (error) {
        if (error.message.includes('fetch')) {
            throw new Error('Network error: Please check your internet connection');
        }
        throw error;
    }
}

// Handle AI enhancement process
async function handleAIEnhancement(textarea) {
    const apiKey = await getOpenAIKey();
    if (!apiKey) {
        showAIMessage('OpenAI API key not configured. Please set it in the extension settings.', 'error');
        return;
    }

    const originalText = textarea.value.trim();
    if (!originalText) {
        showAIMessage('Please enter some text first.', 'error');
        return;
    }

    // Find or create AI button to show loading state
    const aiButton = document.querySelector('.ai-enhance-btn');
    if (aiButton) {
        aiButton.disabled = true;
        aiButton.classList.add('loading');
    }

    try {
        const enhancedText = await enhanceBusinessJustification(originalText, apiKey);
        textarea.value = enhancedText;
        
        // Trigger input event to ensure form validation updates
        textarea.dispatchEvent(new Event('input', { bubbles: true }));
        
        // Show checkmark for success
        if (aiButton) {
            aiButton.classList.remove('loading');
            aiButton.classList.add('success');
            aiButton.innerHTML = '✓';
            
            // Reset to sparkles after 2 seconds
            setTimeout(() => {
                aiButton.classList.remove('success');
                aiButton.innerHTML = '✨';
            }, 2000);
        }
        
    } catch (error) {
        showAIMessage(`Enhancement failed: ${error.message}`, 'error');
    } finally {
        if (aiButton) {
            aiButton.disabled = false;
            aiButton.classList.remove('loading');
        }
    }
}

// Show AI operation message
function showAIMessage(message, type = 'info') {
    // Remove existing message
    const existingMessage = document.querySelector('.ai-message');
    if (existingMessage) {
        existingMessage.remove();
    }

    // Create new message
    const messageDiv = document.createElement('div');
    messageDiv.className = `ai-message ai-message-${type}`;
    messageDiv.textContent = message;

    // Find the textarea container to position the message
    const textareaContainer = document.querySelector('.ms-TextField.is-required.ms-TextField--multiline');
    if (textareaContainer) {
        textareaContainer.insertAdjacentElement('afterend', messageDiv);
        
        // Auto-remove after 4 seconds
        setTimeout(() => {
            if (messageDiv.parentNode) {
                messageDiv.remove();
            }
        }, 4000);
    }
}

// Add AI enhancement button to Additional Questions modal
async function addAIButtonToBusinessJustification() {
    try {
        // Check if extension context is valid
        if (!isExtensionContextValid()) {
            return;
        }
        
        // Check if API key is configured
        const hasKey = await hasOpenAIKey();
        if (!hasKey) {
            return; // Don't show button if no API key
        }
    } catch (error) {
        // Extension context issues, silently return
        return;
    }

    // Look for the business justification textarea (dynamic ID)
    const textarea = document.querySelector('.ms-TextField.is-required.ms-TextField--multiline textarea');
    if (!textarea) {
        return;
    }

    // Check if button already exists
    if (document.querySelector('.ai-enhance-btn')) {
        return;
    }

    // Find the submit button container
    const submitContainer = document.querySelector('._1W7zJMqqfOIlOV9PXVVYxz');
    if (!submitContainer) {
        return;
    }

    // Create AI enhancement button
    const aiButton = document.createElement('button');
    aiButton.type = 'button';
    aiButton.className = 'ai-enhance-btn';
    aiButton.innerHTML = '✨';
    aiButton.title = 'Enhance business justification using AI';
    aiButton.setAttribute('data-is-focusable', 'true');

    // Add click handler
    aiButton.addEventListener('click', () => handleAIEnhancement(textarea));

    // Insert button next to submit button
    const submitButton = submitContainer.querySelector('.ms-Button--primary');
    if (submitButton) {
        submitButton.insertAdjacentElement('beforebegin', aiButton);
    } else {
        submitContainer.appendChild(aiButton);
    }
}

// Listen for API key updates from popup
if (chrome?.runtime?.onMessage) {
    chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
        try {
            if (message.type === 'API_KEY_UPDATED') {
                // Re-check for Additional Questions modal and update AI button visibility
                setTimeout(() => {
                    const modal = document.querySelector('.ms-Modal-scrollableContent');
                    if (modal && modal.textContent.includes('Additional questions')) {
                        // Remove existing button if any
                        const existingButton = document.querySelector('.ai-enhance-btn');
                        if (existingButton) {
                            existingButton.remove();
                        }
                        // Re-add button based on new key status
                        addAIButtonToBusinessJustification();
                    }
                }, 100);
            }
        } catch (error) {
            // Extension context invalidated, ignore
        }
    });
}

// ========================================
// AUTHENTICATION
// ========================================

function getAuthToken() {
    for (let i = 0; i < window.sessionStorage.length; i++) {
        const key = window.sessionStorage.key(i);
        if (!key) continue;
        
        if (key.includes('react-auth') && key.includes('session')) {
            try {
                const value = window.sessionStorage.getItem(key);
                if (value) {
                    let parsed;
                    try {
                        parsed = JSON.parse(value);
                    } catch (e) {
                        continue;
                    }
                    
                    if (parsed && parsed.accessTokens) {
                        for (const subKey in parsed.accessTokens) {
                            if (subKey.includes('elm.iga.azure.com')) {
                                const accessObj = parsed.accessTokens[subKey];
                                if (accessObj && accessObj.token) {
                                    return accessObj.token;
                                }
                            }
                        }
                    }
                    
                    if (parsed && parsed.access_token) {
                        return parsed.access_token;
                    }
                }
            } catch (e) {
                continue;
            }
        }
        
        if (key.includes('msal') && (key.includes('accesstoken') || key.includes('idtoken'))) {
            try {
                const value = window.sessionStorage.getItem(key);
                if (value) {
                    const parsed = JSON.parse(value);
                    if (parsed && parsed.secret) {
                        return parsed.secret;
                    }
                }
            } catch (e) {
                continue;
            }
        }
    }
    
    for (let key in localStorage) {
        if (key.includes('msal') && (key.includes('accesstoken') || key.includes('idtoken'))) {
            try {
                const data = JSON.parse(localStorage.getItem(key));
                if (data && data.secret) {
                    return data.secret;
                }
            } catch (e) {}
        }
    }
    
    return null;
}

function getGraphToken() {
    // Look for react-auth session tokens with graph.microsoft.com
    for (let i = 0; i < window.sessionStorage.length; i++) {
        const key = window.sessionStorage.key(i);
        if (!key) continue;
        
        if (key.includes('react-auth') && key.includes('session')) {
            try {
                const value = window.sessionStorage.getItem(key);
                if (value) {
                    let parsed;
                    try {
                        parsed = JSON.parse(value);
                    } catch (e) {
                        continue;
                    }
                    
                    if (parsed && parsed.accessTokens) {
                        for (const subKey in parsed.accessTokens) {
                            if (subKey.includes('graph.microsoft.com')) {
                                const accessObj = parsed.accessTokens[subKey];
                                if (accessObj && accessObj.token) {
                                    return accessObj.token;
                                }
                            }
                        }
                    }
                }
            } catch (e) {
                continue;
            }
        }
    }
    
    for (let key in localStorage) {
        if (key.includes('msal') && key.includes('accesstoken')) {
            try {
                const data = JSON.parse(localStorage.getItem(key));
                if (data && data.secret) {
                    // Try to validate this is a Graph token by checking the audience
                    try {
                        const payload = JSON.parse(atob(data.secret.split('.')[1]));
                        if (payload.aud && (payload.aud.includes('graph.microsoft.com') || payload.aud === '00000003-0000-0000-c000-000000000000')) {
                            return data.secret;
                        }
                    } catch (e) {
                        // If we can't decode, skip this token
                    }
                }
            } catch (e) {}
        }
    }
    
    return null;
}

async function lookupPackageGUID(displayName, context = 'general') {
    const token = getAuthToken();
    if (!token) {
        throw new Error('Authentication token not found');
    }

    try {
        const searchUrl = `https://elm.iga.azure.com/api/v1/accessPackages`;
        const urlParams = new URLSearchParams({
            '$filter': `contains(displayName,'${displayName}')`
        });

        const response = await fetch(`${searchUrl}?${urlParams}`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });

        if (!response.ok) {
            if (response.status === 403) {
                throw new Error('Access denied to elm.iga.azure.com API. Token may lack required permissions.');
            }
            throw new Error(`API request failed with status ${response.status}`);
        }

        const data = await response.json();

        if (!data.value || data.value.length === 0) {
            let message;
            switch(context) {
                case 'approvers':
                    if (displayName.toLowerCase().includes('stakeholder')) {
                        message = `Cannot display approvers for access package "${displayName}" - Missing permissions to access package information. Once you obtain the Stakeholder package, you will be able to see approver information for this and other packages in the same group.`;
                    } else {
                        message = `Cannot display approvers for access package "${displayName}" - Missing permissions to access package information. You may need elevated permissions or the appropriate Stakeholder package to view this data.`;
                    }
                    break;
                case 'requestors':
                    if (displayName.toLowerCase().includes('stakeholder')) {
                        message = `Cannot display requestors for access package "${displayName}" - Missing permissions to access package information. Once you obtain the Stakeholder package, you will be able to see requestor information for this and other packages in the same group.`;
                    } else {
                        message = `Cannot display requestors for access package "${displayName}" - Missing permissions to access package information. You may need elevated permissions or the appropriate Stakeholder package to view this data.`;
                    }
                    break;
                default:
                    message = `Cannot find access package "${displayName}" - This may be due to missing permissions or the package name not matching exactly.`;
                    break;
            }
            return {
                found: false,
                message: message
            };
        }

        const exactMatch = data.value.find(pkg => pkg.displayName === displayName);
        if (exactMatch) {
            return {
                found: true,
                package: exactMatch,
                matches: data.value.length,
                message: `Exact match found (${data.value.length} total matches)`
            };
        }

        return {
            found: true,
            package: data.value[0],
            matches: data.value.length,
            message: `Partial match found (${data.value.length} total matches)`
        };

    } catch (error) {
        throw error;
    }
}

async function getPackageDetails(packageGuid) {
    const token = getAuthToken();
    if (!token) {
        throw new Error('Authentication token not found');
    }

    try {
        const searchUrl = `https://elm.iga.azure.com/api/v1/accessPackages`;
        const urlParams = new URLSearchParams({
            '$filter': `id eq '${packageGuid}'`
        });

        const response = await fetch(`${searchUrl}?${urlParams}`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });

        if (!response.ok) {
            throw new Error(`Failed to fetch package details: ${response.status}`);
        }

        const data = await response.json();
        return data.value && data.value.length > 0 ? data.value[0] : null;
    } catch (error) {
        throw error;
    }
}

async function getAssignmentPolicies(packageGuid) {
    const token = getAuthToken();
    if (!token) {
        throw new Error('Authentication token not found');
    }

    try {
        const searchUrl = `https://elm.iga.azure.com/api/v1/accessPackageAssignmentPolicies`;
        const urlParams = new URLSearchParams({
            '$filter': `accessPackage/id eq '${packageGuid}'`
        });

        const response = await fetch(`${searchUrl}?${urlParams}`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });

        if (!response.ok) {
            throw new Error(`Failed to fetch assignment policies: ${response.status}`);
        }

        const data = await response.json();
        return data.value || [];
    } catch (error) {
        throw error;
    }
}

async function getGroupMembers(groupId) {
    const graphToken = getGraphToken();
    const fallbackToken = getAuthToken();
    
    if (!graphToken && !fallbackToken) {
        return { error: 'No authentication token found', members: [] };
    }

    const tokensToTry = [
        { token: graphToken, type: 'Graph-specific' },
        { token: fallbackToken, type: 'Fallback' }
    ].filter(t => t.token);

    for (const tokenInfo of tokensToTry) {
        try {
            const graphUrl = `https://graph.microsoft.com/v1.0/groups/${groupId}/members`;
            const urlParams = new URLSearchParams({
                '$select': 'id,displayName,givenName,surname,mail,jobTitle,department,companyName,userPrincipalName'
            });

            const response = await fetch(`${graphUrl}?${urlParams}`, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${tokenInfo.token}`,
                    'Content-Type': 'application/json'
                }
            });

            if (response.ok) {
                const data = await response.json();
                return { members: data.value || [], error: null };
            }

            if (response.status === 403 || response.status === 401) {
                continue; // Try next token
            }
            
        } catch (error) {
            continue; // Try next token
        }
    }

    return { 
        error: 'Authentication failed - no valid token found for Microsoft Graph API. Graph API access may require additional permissions.', 
        members: [] 
    };
}

function createCollapsibleSection(sectionId, contentId, title, subtitle = 'Show/hide details') {
    return `
        <div class="ms-List-cell" data-list-index="0" style="margin-bottom: 8px;">
            <div class="collapsible-wrapper">
                <div class="collapsible-section" data-section="${sectionId}" 
                     role="button" data-is-focusable="true" tabindex="0">
                    <div class="_2Wk_euHnEs1IEUewKXsi4D">
                        <div class="collapsible-section-title">${title}</div>
                        <div class="collapsible-section-subtitle">${subtitle}</div>
                    </div>
                    <div aria-live="assertive" style="position: absolute; width: 1px; height: 1px; margin: -1px; padding: 0px; border: 0px; overflow: hidden; clip: rect(0px, 0px, 0px, 0px);">Collapsed</div>
                </div>
                <div class="collapsible-content-area" id="${contentId}" data-expanded="false">
                    <div class="content-loading">
                        Click to load details...
                    </div>
                </div>
            </div>
        </div>
    `;
}

const collapsibleHandlers = new Map();

function createCollapsibleClickHandler(sectionSelector, contentId, loadingMessage, contentLoader) {
    let contentLoaded = false;
    
    // Store handler configuration
    collapsibleHandlers.set(sectionSelector, {
        contentId,
        loadingMessage,
        contentLoader,
        contentLoaded: false
    });
    
    // Function to setup click handlers with better reliability
    const setupHandlers = () => {
        const sectionHeader = document.querySelector(`.collapsible-section[data-section="${sectionSelector}"]`);
        
        if (!sectionHeader) {
            return false;
        }
        
        // Mark as ready for delegation
        sectionHeader.setAttribute('data-handler-ready', 'true');
        
        return true;
    };
    
    // Try to setup handlers with multiple attempts
    const attemptSetup = (retryCount = 0) => {
        if (setupHandlers()) {
            return;
        }
        
        if (retryCount < 5) {
            setTimeout(() => attemptSetup(retryCount + 1), 100 * (retryCount + 1));
        }
    };
    
    attemptSetup();
}

document.addEventListener('click', async function(e) {
    const target = e.target.closest('.collapsible-section[data-section]');
    if (!target || target.getAttribute('data-handler-ready') !== 'true') {
        return;
    }
    
    e.preventDefault();
    e.stopPropagation();
    
    const sectionSelector = target.getAttribute('data-section');
    const handlerConfig = collapsibleHandlers.get(sectionSelector);
    
    if (!handlerConfig) {
        return;
    }
    
    const content = document.getElementById(handlerConfig.contentId);
    const ariaLive = target.querySelector('[aria-live="assertive"]');
    
    if (!content) {
        return;
    }
    
    const expandedAttr = content.getAttribute('data-expanded');
    const isExpanded = expandedAttr === 'true';
    
    if (isExpanded) {
        // Collapse
        content.style.display = 'none';
        content.setAttribute('data-expanded', 'false');
        if (ariaLive) ariaLive.textContent = 'Collapsed';
    } else {
        // Expand
        content.style.display = 'block';
        content.setAttribute('data-expanded', 'true');
        if (ariaLive) ariaLive.textContent = 'Expanded';
        
        // Load content only when expanding for the first time
        if (!handlerConfig.contentLoaded) {
            handlerConfig.contentLoaded = true;
            
            // Show loading message
            content.innerHTML = `
                <div class="loading-container">
                    <div class="loading-title">${handlerConfig.loadingMessage}</div>
                    <div class="loading-progress">
                        <div class="loading-bar"></div>
                    </div>
                </div>
            `;
            
            try {
                // Create content using the provided loader function
                const loadedContent = await handlerConfig.contentLoader();
                
                // Update content with loaded data
                content.innerHTML = '';
                content.appendChild(loadedContent);
                
            } catch (error) {
                content.innerHTML = `
                    <div class="error-message">
                        <strong>Error loading content:</strong><br>
                        ${error.message}
                    </div>
                `;
            }
        }
    }
});

document.addEventListener('keydown', function(e) {
    if (e.key !== 'Enter' && e.key !== ' ') {
        return;
    }
    
    const target = e.target.closest('.collapsible-section[data-section]');
    if (!target || target.getAttribute('data-handler-ready') !== 'true') {
        return;
    }
    
    e.preventDefault();
    target.click();
});


function generateInitials(givenName, surname, displayName) {
    if (givenName && surname) {
        return (givenName.charAt(0) + surname.charAt(0)).toUpperCase();
    }
    if (displayName) {
        const parts = displayName.split(' ').filter(part => part.length > 0 && !part.includes('('));
        if (parts.length >= 2) {
            return (parts[0].charAt(0) + parts[1].charAt(0)).toUpperCase();
        }
        return displayName.charAt(0).toUpperCase();
    }
    return '??';
}

function createMemberCard(member) {
    if (member['@odata.type'] !== '#microsoft.graph.user') {
        return '';
    }
    
    const initials = generateInitials(member.givenName, member.surname, member.displayName);
    const jobTitle = member.jobTitle || 'No title';
    const department = member.department || 'No department';
    const email = member.mail || member.userPrincipalName;
    
    return `
        <div class="member-card" title="${email}">
            <div class="member-avatar">${initials}</div>
            <div class="member-info">
                <div class="member-name">${member.displayName}</div>
                <div class="member-email">${jobTitle}</div>
                <div class="member-email" style="font-size: 10px; color: #8a8886;">${department}</div>
            </div>
        </div>
    `;
}

function createMemberGrid(members) {
    if (!members || members.length === 0) {
        return '';
    }
    
    const memberCards = members.map(member => createMemberCard(member)).join('');
    return `
        <div style="margin-top: 8px;">
            <div class="metadata-label" style="margin-bottom: 6px;">Group Members (${members.length})</div>
            <div class="members-grid">${memberCards}</div>
        </div>
    `;
}

function extractRiskLevel(displayName) {
    const riskMatch = displayName.match(/\[(\w+)\]/i);
    if (riskMatch) {
        return riskMatch[1].toLowerCase();
    }
    return 'unknown';
}

function formatDate(dateString) {
    if (!dateString) return 'Not available';
    try {
        const date = new Date(dateString);
        return date.toLocaleString('en-US', {
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch (e) {
        return dateString;
    }
}

let currentPackageId = null;
let currentPackageName = null;

function getCurrentPackageName() {
    // Try to extract from the modal title
    const modalTitle = document.querySelector('.ms-Modal h3, .ms-Dialog-title, [data-automation-id="DetailsPanelHeaderText"]');
    if (modalTitle && modalTitle.textContent) {
        return modalTitle.textContent.trim();
    }

    // Fallback to stored package name
    if (currentPackageName) {
        return currentPackageName;
    }
    
    return null;
}

function interceptPackageClicks() {
    document.addEventListener('click', function(e) {
        // Check if click is on a Request button or package row
        const requestButton = e.target.closest('.ms-Link');
        const row = e.target.closest('.ms-DetailsRow');
        
        if (requestButton && requestButton.textContent === 'Request' && row) {
            // Extract package name from the row
            const packageNameElement = row.querySelector('.ms-pii, [data-automation-id="DetailsRowCell"]');
            if (packageNameElement) {
                currentPackageName = packageNameElement.textContent.trim();
                currentPackageId = currentPackageName.toLowerCase().replace(/[^a-z0-9]+/g, '-');
            }
        }
    }, true);
}

function copyToClipboard(text) {
    navigator.clipboard.writeText(text).catch(() => {
        // Fallback for older browsers
        const textArea = document.createElement('textarea');
        textArea.value = text;
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
    });
}

async function createSimpleRequestorsContent() {
    const container = document.createElement('div');
    container.className = 'simple-requestors-content';

    const packageName = getCurrentPackageName();
    if (!packageName) {
        container.innerHTML = '<div class="error-message">Could not determine package name from modal</div>';
        return container;
    }

    const token = getAuthToken();
    if (!token) {
        container.innerHTML = `
            <div class="error-message">
                <strong>Authentication token not found</strong><br>
                Please try refreshing the page or signing out and back in.
            </div>`;
        return container;
    }

    try {
        // Look up the package GUID and metadata
        const lookupResult = await lookupPackageGUID(packageName, 'requestors');
        
        if (lookupResult.found) {
            // Get full package details and assignment policies
            const packageDetails = await getPackageDetails(lookupResult.package.id);
            const pkg = packageDetails || lookupResult.package;
            const assignmentPolicies = await getAssignmentPolicies(pkg.id);
            
            if (assignmentPolicies && assignmentPolicies.length > 0) {
                const allRequestorGroups = new Set();
                
                // Collect all requestor groups from all policies
                for (const policy of assignmentPolicies) {
                    if (policy.specificAllowedTargets && policy.specificAllowedTargets.length > 0) {
                        policy.specificAllowedTargets.forEach(target => {
                            const displayName = target.displayName || target.description || 'Unknown Group';
                            if (displayName !== 'Unknown Group') {
                                allRequestorGroups.add(displayName);
                            }
                        });
                    }
                    
                    if (policy.requestorSettings?.allowedRequestors && policy.requestorSettings.allowedRequestors.length > 0) {
                        policy.requestorSettings.allowedRequestors.forEach(requestor => {
                            const displayName = requestor.displayName || requestor.description || 'Unknown Group';
                            if (displayName !== 'Unknown Group') {
                                allRequestorGroups.add(displayName);
                            }
                        });
                    }
                }
                
                // Display each unique requestor group
                const uniqueRequestorGroups = Array.from(allRequestorGroups);
                let requestorsHtml = '';
                
                if (uniqueRequestorGroups.length > 0) {
                    for (const groupName of uniqueRequestorGroups) {
                        requestorsHtml += `
                            <div class="simple-requestor-group">
                                <div class="requestor-group-header">${groupName}</div>
                            </div>
                        `;
                    }
                } else {
                    requestorsHtml = `
                        <div class="open-access-message">
                            <strong>ℹ️ Open access</strong><br>
                            No specific requestor groups configured - this package may be open to all users.
                        </div>
                    `;
                }
                
                container.innerHTML = requestorsHtml;
            } else {
                container.innerHTML = `
                    <div class="no-policies-message">
                        ⚠️ No assignment policies found for this package.
                    </div>
                `;
            }
        } else {
            container.innerHTML = `<div class="warning-message">⚠️ ${lookupResult.message}</div>`;
        }

    } catch (error) {
        container.innerHTML = `
            <div class="error-message">
                <strong>Error looking up package information:</strong><br>
                ${error.message}
            </div>`;
    }

    return container;
}

async function createSimpleApproversContent() {
    const container = document.createElement('div');
    container.className = 'simple-approvers-content';

    const packageName = getCurrentPackageName();
    if (!packageName) {
        container.innerHTML = '<div class="error-message">Could not determine package name from modal</div>';
        return container;
    }

    const token = getAuthToken();
    if (!token) {
        container.innerHTML = `
            <div class="error-message">
                <strong>Authentication token not found</strong><br>
                Please try refreshing the page or signing out and back in.
            </div>`;
        return container;
    }

    try {
        // Look up the package GUID and metadata
        const lookupResult = await lookupPackageGUID(packageName, 'approvers');
        
        if (lookupResult.found) {
            // Get full package details and assignment policies
            const packageDetails = await getPackageDetails(lookupResult.package.id);
            const pkg = packageDetails || lookupResult.package;
            const assignmentPolicies = await getAssignmentPolicies(pkg.id);
            
            if (assignmentPolicies && assignmentPolicies.length > 0) {
                // Collect all approver groups from all policies and stages
                const allApproverGroups = new Set();
                
                for (const policy of assignmentPolicies) {
                    const approvalSettings = policy.approvalSettings || policy.requestApprovalSettings;
                    let stages = [];
                    
                    if (approvalSettings?.stages) {
                        stages = approvalSettings.stages;
                    } else if (approvalSettings?.approvalStages) {
                        stages = approvalSettings.approvalStages;
                    }
                    
                    if (stages && stages.length > 0) {
                        for (const stage of stages) {
                            if (stage.primaryApprovers) {
                                stage.primaryApprovers.forEach(approver => {
                                    const displayName = approver.displayName || approver.description || 'Unknown Group';
                                    const groupId = approver.groupId || approver.id || approver.objectId;
                                    if (groupId) {
                                        allApproverGroups.add(JSON.stringify({ displayName, groupId }));
                                    }
                                });
                            }
                            if (stage.escalationApprovers) {
                                stage.escalationApprovers.forEach(approver => {
                                    const displayName = approver.displayName || approver.description || 'Unknown Group';
                                    const groupId = approver.groupId || approver.id || approver.objectId;
                                    if (groupId) {
                                        allApproverGroups.add(JSON.stringify({ displayName, groupId }));
                                    }
                                });
                            }
                        }
                    }
                }
                
                // Display each unique approver group with all members shown by default
                const uniqueGroups = Array.from(allApproverGroups).map(g => JSON.parse(g));
                let approversHtml = '';
                
                if (uniqueGroups.length > 0) {
                    for (let i = 0; i < uniqueGroups.length; i++) {
                        const group = uniqueGroups[i];
                        
                        try {
                            const memberData = await getGroupMembers(group.groupId);
                            
                            approversHtml += `
                                <div class="simple-approver-group">
                                    <div class="approver-group-header">${group.displayName}</div>
                            `;
                            
                            // Show all members by default
                            if (memberData && !memberData.error && memberData.members.length > 0) {
                                approversHtml += createMemberGrid(memberData.members);
                            } else if (memberData && memberData.error) {
                                if (memberData.error.includes('Authentication failed')) {
                                    approversHtml += `
                                        <div class="graph-api-warning">
                                            <div class="graph-api-warning-title">⚠️ Graph API Access Required</div>
                                            <div class="graph-api-warning-text">
                                                To see group members, Microsoft Graph API access is needed.<br>
                                                <strong>Workaround:</strong> Run this command in a separate terminal:<br>
                                                <code class="graph-api-command">az rest --method GET --url "https://graph.microsoft.com/v1.0/groups/${group.groupId}/members"</code>
                                            </div>
                                        </div>
                                    `;
                                } else {
                                    approversHtml += `
                                        <div class="member-load-error">
                                            ⚠️ Could not load group members: ${memberData.error}
                                        </div>
                                    `;
                                }
                            } else if (memberData && memberData.members.length === 0) {
                                approversHtml += `
                                    <div class="no-members-message">
                                        No members found in this group
                                    </div>
                                `;
                            } else {
                                approversHtml += `
                                    <div class="loading-members-message">
                                        Loading group members...
                                    </div>
                                `;
                            }
                            
                            approversHtml += '</div>';
                            
                        } catch (error) {
                            approversHtml += `
                                <div class="simple-approver-group">
                                    <div class="approver-group-header">${group.displayName}</div>
                                    <div style="margin-top: 8px; font-size: 12px; color: #a80000; font-style: italic;">
                                        ⚠️ Could not load group members: ${error.message}
                                    </div>
                                </div>
                            `;
                        }
                    }
                } else {
                    approversHtml = `
                        <div style="padding: 12px; background: #fff4ce; border: 1px solid #ffb900; border-radius: 6px;">
                            ⚠️ No approver groups configured for this package.
                        </div>
                    `;
                }
                
                container.innerHTML = approversHtml;
            } else {
                container.innerHTML = `
                    <div class="no-policies-message">
                        ⚠️ No assignment policies found for this package.
                    </div>
                `;
            }
        } else {
            container.innerHTML = `<div class="warning-message">⚠️ ${lookupResult.message}</div>`;
        }

    } catch (error) {
        container.innerHTML = `
            <div class="error-message">
                <strong>Error looking up package information:</strong><br>
                ${error.message}
            </div>`;
    }

    return container;
}


function addApproversToRequestDetailsTab() {
    // Look for the Request details tab panel content
    const tabPanel = document.querySelector('[role="tabpanel"]');
    if (!tabPanel) return;
    
    // Check if this is the Request details tab by looking for "Share a link" content
    const shareSection = tabPanel.querySelector('button[id="share-button-id"]');
    if (!shareSection) return;
    
    // Check if we already added approvers
    if (tabPanel.querySelector('[data-request-details-approvers-section]')) return;
    
    // Extract package name from current package name
    const packageName = getCurrentPackageName();
    if (!packageName) {
        return;
    }
    
    // Store the package name for createApproversContent
    currentPackageName = packageName;
    
    // Find the container with the share button section
    const shareContainer = shareSection.closest('div');
    if (!shareContainer) {
        return;
    }
    
    // Create collapsible approvers section
    const approversSection = document.createElement('div');
    approversSection.setAttribute('data-request-details-approvers-section', 'true');
    approversSection.style.cssText = `
        margin-top: 16px;
        border-top: 1px solid #edebe9;
        padding-top: 12px;
    `;
    
    // Use reusable collapsible section for approvers
    approversSection.innerHTML = createCollapsibleSection('request-details-approvers', 'request-details-approvers-content', 'Who can approve?');
    
    // Insert before the share container
    shareContainer.insertAdjacentElement('beforebegin', approversSection);
    
    // Add requestor section
    const requestorsSection = document.createElement('div');
    requestorsSection.setAttribute('data-request-details-requestors-section', 'true');
    requestorsSection.style.cssText = `
        margin-top: 8px;
    `;
    
    // Use reusable collapsible section for requestors
    requestorsSection.innerHTML = createCollapsibleSection('request-details-requestors', 'request-details-requestors-content', 'Who can request?');
    
    // Insert after the approvers section
    approversSection.insertAdjacentElement('afterend', requestorsSection);
    
    // Setup click handlers with delays to ensure DOM is ready
    setTimeout(() => {
        createCollapsibleClickHandler('request-details-approvers', 'request-details-approvers-content', 'Loading who can approve...', createSimpleApproversContent);
        createCollapsibleClickHandler('request-details-requestors', 'request-details-requestors-content', 'Loading who can request...', createSimpleRequestorsContent);
    }, 50);
}

function addApproversToAccessRequestPanel() {
    // Look for the access request panel
    const accessRequestPanel = document.querySelector('.ms-Panel-content');
    if (!accessRequestPanel) return;
    
    // Check if this is an access request panel by looking for "Access request" header
    const headerElement = document.querySelector('.ms-Panel-headerText');
    if (!headerElement || !headerElement.textContent.includes('Access request')) return;
    
    // Check if we already added approvers section
    if (accessRequestPanel.querySelector('[data-access-request-approvers-section]')) return;
    
    
    // Extract package name from "You submitted a request for" text
    let packageName = null;
    
    // Look for the activity text that shows the request
    const activityTexts = accessRequestPanel.querySelectorAll('.ms-ActivityItem-activityText');
    for (const activityText of activityTexts) {
        const text = activityText.textContent.trim();
        if (text.includes('You submitted a request for')) {
            // Extract package name after "You submitted a request for"
            const match = text.match(/You submitted a request for (.+)/);
            if (match && match[1]) {
                packageName = match[1].trim();
                break;
            }
        }
    }
    
    if (!packageName) {
        packageName = 'Unknown Package';
    }
    
    // Store the package name for createSimpleApproversContent
    currentPackageName = packageName;
    
    // Create collapsible approvers section
    const approversSection = document.createElement('div');
    approversSection.setAttribute('data-access-request-approvers-section', 'true');
    approversSection.style.cssText = `
        margin-top: 16px;
        border-top: 1px solid #edebe9;
        padding-top: 12px;
    `;
    
    // Use reusable collapsible section for approvers
    approversSection.innerHTML = createCollapsibleSection('access-request-approvers', 'access-request-approvers-content', 'Who can approve?');
    
    // Append section to panel
    accessRequestPanel.appendChild(approversSection);
    
    // Add requestor section
    const requestorsSection = document.createElement('div');
    requestorsSection.setAttribute('data-access-request-requestors-section', 'true');
    requestorsSection.style.cssText = `
        margin-top: 8px;
    `;
    
    // Use reusable collapsible section for requestors
    requestorsSection.innerHTML = createCollapsibleSection('access-request-requestors', 'access-request-requestors-content', 'Who can request?');
    
    // Insert after the approvers section
    approversSection.insertAdjacentElement('afterend', requestorsSection);
    
    // Setup click handlers with delays to ensure DOM is ready
    setTimeout(() => {
        createCollapsibleClickHandler('access-request-approvers', 'access-request-approvers-content', 'Loading who can approve...', createSimpleApproversContent);
        createCollapsibleClickHandler('access-request-requestors', 'access-request-requestors-content', 'Loading who can request...', createSimpleRequestorsContent);
    }, 50);
}


function watchForModal() {
    let lastCheckTime = 0;
    let lastPanelContent = '';
    
    const checkForModal = () => {
        const now = Date.now();
        if (now - lastCheckTime < 50) return; // Reduced throttle for better responsiveness
        lastCheckTime = now;
        
        // Check for traditional modal with tabs
        const modal = document.querySelector('.ms-Modal.is-open') || 
                     document.querySelector('[role="dialog"]');
        
        if (modal) {
            const hasCorrectTabs = modal.textContent.includes('Request details') && 
                                 modal.textContent.includes('Resources');
            
            if (hasCorrectTabs) {
                // Always attempt to add approvers to Request details tab when modal is detected
                setTimeout(addApproversToRequestDetailsTab, 100);
                setTimeout(addApproversToRequestDetailsTab, 300);
                setTimeout(addApproversToRequestDetailsTab, 600);
                
                // Also add click handlers to tabs if not already added
                const tabList = modal.querySelector('[role="tablist"]');
                if (tabList && !tabList.hasAttribute('data-approvers-listeners-added')) {
                    tabList.setAttribute('data-approvers-listeners-added', 'true');
                    
                    const tabs = tabList.querySelectorAll('[role="tab"]');
                    tabs.forEach(tab => {
                        tab.addEventListener('click', function() {
                            // If this is the Request details tab, add approvers after a delay
                            if (this.textContent.includes('Request details')) {
                                setTimeout(addApproversToRequestDetailsTab, 100);
                                setTimeout(addApproversToRequestDetailsTab, 300);
                                setTimeout(addApproversToRequestDetailsTab, 600);
                            }
                        });
                    });
                }
            }

            // Check for Additional Questions modal
            const additionalQuestionsTitle = modal.querySelector('h3[title="Additional questions"]');
            if (additionalQuestionsTitle) {
                // Attempt to add AI button to business justification field
                setTimeout(addAIButtonToBusinessJustification, 100);
                setTimeout(addAIButtonToBusinessJustification, 300);
                setTimeout(addAIButtonToBusinessJustification, 600);
            } else {
                // Also check for modal content that might contain "Additional questions"
                const hasAdditionalQuestions = modal.textContent.includes('Additional questions');
                if (hasAdditionalQuestions) {
                    setTimeout(addAIButtonToBusinessJustification, 100);
                    setTimeout(addAIButtonToBusinessJustification, 300);
                    setTimeout(addAIButtonToBusinessJustification, 600);
                }
            }
        }
        
        // Check for access request panel
        const headerElement = document.querySelector('.ms-Panel-headerText');
        if (headerElement && headerElement.textContent.includes('Access request')) {
            const accessRequestPanel = document.querySelector('.ms-Panel-content');
            if (accessRequestPanel) {
                const currentContent = accessRequestPanel.textContent;
                
                // Check if we have an access request panel and it's not already processed
                if (!accessRequestPanel.querySelector('[data-approvers-section]')) {
                    
                    // Try to add button immediately if we have access request content
                    if (currentContent.includes('You submitted a request for') || 
                        currentContent.includes('Request history') ||
                        currentContent.length > 100) {
                        
                        addApproversToAccessRequestPanel();
                    }
                }
                
                // Also check for content changes
                if (currentContent !== lastPanelContent && currentContent.length > 100) {
                    lastPanelContent = currentContent;
                    
                    // Multiple attempts with different delays to catch dynamic loading
                    setTimeout(addApproversToAccessRequestPanel, 50);
                    setTimeout(addApproversToAccessRequestPanel, 200);
                    setTimeout(addApproversToAccessRequestPanel, 500);
                    setTimeout(addApproversToAccessRequestPanel, 1000);
                }
            }
        }
    };
    
    const observer = new MutationObserver(checkForModal);
    observer.observe(document.body, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ['class', 'style']
    });
    
    // More frequent checking for better responsiveness
    setInterval(checkForModal, 100);
}

function init() {
    interceptPackageClicks();
    watchForModal();
    
    // Check for existing access request panels on page load
    setTimeout(() => {
        addApproversToAccessRequestPanel();
    }, 500);
    setTimeout(() => {
        addApproversToAccessRequestPanel();
    }, 1500);
    setTimeout(() => {
        addApproversToAccessRequestPanel();
    }, 3000);
    
    // Add global function for manual testing
    window.addAIButton = addAIButtonToBusinessJustification;
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}