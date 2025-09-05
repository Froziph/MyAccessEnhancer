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

        maxTokens: 300
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

        maxTokens: 500
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

        maxTokens: 1500
    },

    // Access Packages Documentation Assistant
    ACCESS_PACKAGES_ASSISTANT: {
        systemPrompt: `You are an expert assistant specializing in Azure Active Directory (AAD) access management and DevOps Center of Excellence (CoE) access packages. You have comprehensive knowledge about the access package system described in the provided documentation.

## Access Package Categories
- **Process Governance**: Approval rights for other access packages (no direct resource access)
- **System Access Packages**: Cross-landing zone access (Engineers, DEV/TEST Owner, Operators)
- **Landing Zone Access Packages**: Environment-specific access (PRD/STA Reader, Contributor, Owner, etc.)
- **Azure DevOps Packages**: DevOps-specific permissions and deployment approvals

## Risk-Based Policies
- **Critical Risk**: 1 hour expiration, Manager approval required
- **High Risk**: 1 hour expiration, SME approval required  
- **Medium Risk**: 8 hours expiration, SME approval required
- **Low Risk**: 180 days expiration, auto-approval
- **Process Governance**: Various durations with different approval levels

## Specific Package Details

### Process Governance
- **[SystemName] Managers**: Approve Process Governance packages (3-6 months duration)
- **[SystemName] Subject Matter Expert**: Approve high/medium risk packages (3 months, team requestor)
- **[SystemName] Stakeholder**: View/request System and PRD/STA packages (6 months)

### System Access Packages
- **[SystemName] Engineers**: DEV/TEST groups, Contributor rights, Azure DevOps access (Medium Risk, 8hr, SME approval)
- **[SystemName] DEV/TEST owner**: Owner on DEV/TST subscriptions (Medium Risk, 8hr, SME approval)
- **[SystemName] Operators**: Production monitoring, Support Request Contributor (6 months, Manager approval)

### Landing Zone Packages
- **[SystemName] [LandingZoneName] [PRD/STA] Reader**: Reader role (Low Risk, 180 days, auto-approval)
- **[SystemName] [LandingZoneName] [PRD/STA] Contributor**: Contributor role (High Risk, 1hr, SME approval)
- **[SystemName] [LandingZoneName] [PRD/STA] Owner**: Owner role (Critical Risk, 1hr, Manager approval)
- **ConfReader/ConfWriter**: Confidential data access (High Risk, 1hr, SME approval)
- **ConfPersonalReader/ConfPersonalWriter**: Personal data access (High Risk, 1hr, SME approval)

### Azure DevOps Packages
- **[SystemName] ADO Production Deployment Approver**: Approve prod deployments (ADO Medium Risk, 6 months, SME approval)
- **[SystemName] ADO Project Manager**: Repo policies, permissions (ADO Critical Risk, 24hr, SME approval)
- **[SystemName] ADO Endpoint Admin**: System endpoints (ADO High Risk, 5 days, SME approval)

## Response Guidelines

### Structure Your Responses
1. **Start with a clear, direct answer**
2. **Use proper headings and sections**
3. **Include specific details with proper formatting**
4. **End with actionable next steps**

### HTML Formatting Rules
- Use h3 for main section headings
- Use h4 for subsections  
- Use p for paragraphs with proper spacing
- Use ul and li for lists
- Use div with class="access-package-info" for main content blocks
- Use div with class="warning-box" for important warnings

### Required Classes
- span with class="risk-badge [level]" for risk levels (with appropriate emoji)
- span with class="package-name" for package names
- span with class="duration" for time durations

### Response Structure Template
Always follow this structure:
- h3 with direct answer
- div with class="access-package-info" containing:
  - h4 for Package Details
  - p and ul with specific information
  - h4 for Approval Process
  - p with risk badge and duration
  - h4 for Next Steps
  - p with actionable instructions

Replace [SystemName] with actual system name if provided by user. Always be helpful and actionable.`,

        userPromptTemplate: `{userQuestion}`,

        maxTokens: 800,
        temperature: 0.7
    }
};

// ========================================
// AI CHAT ASSISTANT FUNCTIONALITY
// ========================================

// Chat assistant state
let chatMessages = [];
let chatModal = null;
let conversationHistory = []; // Store last 6 messages (3 exchanges) for context

// Create and show AI chat modal dialog with modern styling
function showChatModal() {
    if (chatModal && document.body.contains(chatModal)) {
        return;
    }

    // Create the main modal overlay
    chatModal = document.createElement('div');
    chatModal.className = 'ms-Modal is-open ai-chat-modal';
    chatModal.setAttribute('role', 'dialog');
    chatModal.setAttribute('aria-modal', 'true');
    chatModal.style.zIndex = '1000';

    // Create the modal HTML structure with modern styling
    chatModal.innerHTML = `
        <div class="ms-Overlay ms-Overlay--dark ai-modal-overlay"></div>
        <div class="ms-Modal-scrollableContent ai-modal-content">
            <div class="ms-Modal-main ai-modal-main">
                <div class="ai-modal-header">
                    <div class="ai-modal-header-content">
                        <div class="ai-modal-icon">
                            <div class="ai-modal-icon-inner">🤖</div>
                        </div>
                        <div class="ai-modal-title">
                            <h2>Access Packages Assistant</h2>
                            <p>Your AI guide to access management</p>
                        </div>
                    </div>
                    <button class="ai-modal-close" type="button" aria-label="Close dialog">
                        <span class="ai-close-icon">×</span>
                    </button>
                </div>
                
                <div class="ai-modal-body">
                    <div class="ai-chat-messages" id="ai-chat-messages">
                        <!-- Welcome message -->
                        <div class="ai-welcome-card">
                            <div class="ai-welcome-icon">🤖</div>
                            <div class="ai-welcome-content">
                                <h3>Welcome! I'm here to help you navigate access packages.</h3>
                                <div class="ai-capabilities">
                                    <div class="ai-capability">
                                        <span class="ai-capability-icon">●</span>
                                        <span>Access package types & risk levels</span>
                                    </div>
                                    <div class="ai-capability">
                                        <span class="ai-capability-icon">◆</span>
                                        <span>Approval workflows & permissions</span>
                                    </div>
                                    <div class="ai-capability">
                                        <span class="ai-capability-icon">◖</span>
                                        <span>Expiration policies & renewals</span>
                                    </div>
                                    <div class="ai-capability">
                                        <span class="ai-capability-icon">◻</span>
                                        <span>Data classification patterns</span>
                                    </div>
                                </div>
                                <p class="ai-welcome-prompt">What would you like to know about access packages?</p>
                            </div>
                        </div>
                    </div>
                </div>
                
                <div class="ai-modal-footer">
                    <div class="ai-input-container">
                        <div class="ai-input-wrapper">
                            <textarea class="ai-input-field" id="ai-chat-input" 
                                     placeholder="Ask me anything about access packages, risk levels, or permissions..."
                                     rows="1"></textarea>
                            <button class="ai-send-button" id="ai-chat-send" type="button" title="Send message">
                                <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
                                    <path d="M2.01 21L23 12 2.01 3 2 10l15 2-15 2z"/>
                                </svg>
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;

    // Add to document body
    document.body.appendChild(chatModal);

    // Add event listeners
    const closeButton = chatModal.querySelector('.ai-modal-close');
    const sendButton = chatModal.querySelector('.ai-send-button');
    const input = chatModal.querySelector('.ai-input-field');
    const overlay = chatModal.querySelector('.ai-modal-overlay');

    if (closeButton) {
        closeButton.addEventListener('click', closeChatModal);
    }

    if (sendButton) {
        sendButton.addEventListener('click', sendChatMessage);
    }

    if (input) {
        input.addEventListener('keypress', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                sendChatMessage();
            }
        });
        
        // Auto-resize textarea
        input.addEventListener('input', () => {
            input.style.height = 'auto';
            input.style.height = Math.min(input.scrollHeight, 120) + 'px';
        });
    }
    
    // Close on overlay click
    if (overlay) {
        overlay.addEventListener('click', closeChatModal);
    }

    // Focus on input after panel is shown
    setTimeout(() => {
        if (input) {
            input.focus();
        }
    }, 100);

}

// Close chat modal
function closeChatModal() {
    if (chatModal && document.body.contains(chatModal)) {
        chatModal.remove();
        chatModal = null;
        // Clear conversation history when modal is closed to save memory
        conversationHistory = [];
    }
}

// Send message to AI assistant
async function sendChatMessage() {
    const input = document.querySelector('.ai-input-field');
    const messagesContainer = document.getElementById('ai-chat-messages');
    const sendButton = document.querySelector('.ai-send-button');

    const userMessage = input.value.trim();
    if (!userMessage) return;

    const apiKey = await getOpenAIKey();
    if (!apiKey) {
        showChatError('OpenAI API key not configured. Please set it in the extension settings.');
        return;
    }

    // Add user message
    addChatMessage(userMessage, 'user');
    input.value = '';

    // Check cache first for instant response
    const cacheKey = userMessage.toLowerCase().trim();
    if (responseCache.has(cacheKey)) {
        addChatMessage(responseCache.get(cacheKey), 'bot');
        return;
    }

    // Show typing indicator for new queries
    const typingId = addTypingIndicator();
    sendButton.disabled = true;

    try {
        // Use streaming response for better UX
        const response = await queryChatAssistantWithStreaming(userMessage, apiKey, typingId);
        removeTypingIndicator(typingId);
        if (response) {
            addChatMessage(response, 'bot');
        }
    } catch (error) {
        removeTypingIndicator(typingId);
        showChatError(`Error: ${error.message}`);
    } finally {
        sendButton.disabled = false;
    }
}

// Well-formatted response cache based on actual documentation
const responseCache = new Map([
    ['what are the risk levels', 
        '<h3>Access Package Risk Categories</h3>' +
        '<div class="access-package-info">' +
        '<ul>' +
        '<li><span class="risk-badge critical">🔴 Critical Risk</span> - <span class="duration">1 hour</span> expiration, <strong>Manager approval</strong></li>' +
        '<li><span class="risk-badge high">🟠 High Risk</span> - <span class="duration">1 hour</span> expiration, <strong>SME approval</strong></li>' +
        '<li><span class="risk-badge medium">🟡 Medium Risk</span> - <span class="duration">8 hours</span> expiration, <strong>SME approval</strong></li>' +
        '<li><span class="risk-badge low">🟢 Low Risk</span> - <span class="duration">180 days</span> expiration, <strong>auto-approval</strong></li>' +
        '<li><span class="risk-badge governance">🔵 Process Governance</span> - Various durations, different approval levels</li>' +
        '</ul>' +
        '</div>'],

    ['risk levels', 
        '<h3>Quick Risk Reference</h3>' +
        '<p><span class="risk-badge critical">Critical</span> 1hr (Manager) | <span class="risk-badge high">High</span> 1hr (SME) | <span class="risk-badge medium">Medium</span> 8hr (SME) | <span class="risk-badge low">Low</span> 180d (auto)</p>'],

    ['engineers package', 
        '<h3>Engineers Access Package</h3>' +
        '<div class="access-package-info">' +
        '<h4>Package Details</h4>' +
        '<p><span class="package-name">[SystemName] Engineers</span> provides development access:</p>' +
        '<ul>' +
        '<li>All DEV/TEST AAD Groups (except Owner)</li>' +
        '<li>Contributor + Key Vault Admin on DEV/TST subscriptions</li>' +
        '<li>Reader access on staging/production subscriptions</li>' +
        '<li>Azure DevOps Contributor + Build admin rights</li>' +
        '</ul>' +
        '<h4>Approval Process</h4>' +
        '<p><span class="risk-badge medium">🟡 Medium Risk</span> - <span class="duration">8 hours</span></p>' +
        '<p>Approved by: <strong>SME (Subject Matter Expert)</strong></p>' +
        '</div>'],

    ['production access', 
        '<h3>Production Access Options</h3>' +
        '<div class="access-package-info">' +
        '<h4>Available Packages</h4>' +
        '<ul>' +
        '<li><span class="package-name">[SystemName] [LandingZoneName] PRD Reader</span><br>' +
        '    <span class="risk-badge low">🟢 Low Risk</span> - <span class="duration">180 days</span>, auto-approved</li>' +
        '<li><span class="package-name">[SystemName] [LandingZoneName] PRD Contributor</span><br>' +
        '    <span class="risk-badge high">🟠 High Risk</span> - <span class="duration">1 hour</span>, SME approval</li>' +
        '<li><span class="package-name">[SystemName] [LandingZoneName] PRD Owner</span><br>' +
        '    <span class="risk-badge critical">🔴 Critical Risk</span> - <span class="duration">1 hour</span>, Manager approval</li>' +
        '</ul>' +
        '<h4>Next Steps</h4>' +
        '<p>Choose based on your access needs. Reader for monitoring, Contributor for changes, Owner for full control.</p>' +
        '</div>'],

    ['what is sme', 
        '<h3>Subject Matter Expert (SME)</h3>' +
        '<div class="access-package-info">' +
        '<p><strong>SME</strong> = Subject Matter Expert for your system.</p>' +
        '<h4>What They Do</h4>' +
        '<ul>' +
        '<li>Approve <span class="risk-badge high">🟠 High Risk</span> access requests</li>' +
        '<li>Approve <span class="risk-badge medium">🟡 Medium Risk</span> access requests</li>' +
        '<li>Provide technical expertise for access decisions</li>' +
        '</ul>' +
        '<h4>How to Contact</h4>' +
        '<p>Find members of your <span class="package-name">[SystemName] Subject Matter Expert</span> group in the access packages portal.</p>' +
        '</div>']
]);

// Query the chat assistant with conversation context and caching
async function queryChatAssistant(userQuestion, apiKey) {
    const prompt = AI_PROMPTS.ACCESS_PACKAGES_ASSISTANT;
    const selectedModel = await getOpenAIModel();
    
    // Check cache for common/exact queries
    const cacheKey = userQuestion.toLowerCase().trim();
    if (responseCache.has(cacheKey)) {
        return responseCache.get(cacheKey);
    }
    
    // Add page context to help AI understand what user is looking at
    const pageContext = detectPageContext();
    
    // Build messages with conversation history
    const messages = [
        {
            role: 'system',
            content: prompt.systemPrompt + (pageContext ? `\n\nCURRENT CONTEXT: ${pageContext}` : '')
        }
    ];
    
    // Add recent conversation history for context (last 6 messages = 3 exchanges)
    conversationHistory.slice(-6).forEach(msg => {
        messages.push(msg);
    });
    
    // Add current user question
    messages.push({
        role: 'user',
        content: userQuestion
    });
    
    try {
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: selectedModel,
                messages: messages,
                max_tokens: prompt.maxTokens,
                temperature: prompt.temperature || 0.7,
                stream: false // Will implement streaming in next phase
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

        const aiResponse = data.choices[0].message.content.trim();
        
        // Update conversation history
        conversationHistory.push(
            { role: 'user', content: userQuestion },
            { role: 'assistant', content: aiResponse }
        );
        
        // Keep only last 6 messages (3 exchanges)
        if (conversationHistory.length > 6) {
            conversationHistory = conversationHistory.slice(-6);
        }
        
        // Cache responses for common/short queries
        if (userQuestion.length < 100) {
            responseCache.set(cacheKey, aiResponse);
            // Clear old cache entries if too many
            if (responseCache.size > 50) {
                const firstKey = responseCache.keys().next().value;
                responseCache.delete(firstKey);
            }
        }

        return aiResponse;

    } catch (error) {
        if (error.message.includes('fetch')) {
            throw new Error('Network error: Please check your internet connection');
        }
        throw error;
    }
}

// Detect current page context for better responses
function detectPageContext() {
    let context = '';
    
    // Check for "Additional questions" modal (business justification)
    if (document.querySelector('h3[title="Additional questions"]')) {
        context = 'User is in the "Additional questions" modal filling out business justification for an access request. ';
    }
    
    // Check for package selection/listing page
    else if (document.querySelector('[data-testid*="package"], [aria-label*="package"]')) {
        context = 'User is on the access packages listing/selection page. ';
    }
    
    // Check for approval workflow page
    else if (document.querySelector('[aria-label*="approval"], [aria-label*="request"]')) {
        context = 'User is in an approval or request workflow. ';
    }
    
    // Check for MyAccess main page
    else if (window.location.hostname.includes('myaccess')) {
        context = 'User is on the MyAccess portal. ';
    }
    
    return context;
}

// Enhanced query with simulated streaming for better UX
async function queryChatAssistantWithStreaming(userQuestion, apiKey, typingId) {
    const prompt = AI_PROMPTS.ACCESS_PACKAGES_ASSISTANT;
    const selectedModel = await getOpenAIModel();
    
    // Add page context to help AI understand what user is looking at
    const pageContext = detectPageContext();
    
    // Build messages with conversation history
    const messages = [
        {
            role: 'system',
            content: prompt.systemPrompt + (pageContext ? `\n\nCURRENT CONTEXT: ${pageContext}` : '')
        }
    ];
    
    // Add recent conversation history for context
    conversationHistory.slice(-6).forEach(msg => {
        messages.push(msg);
    });
    
    // Add current user question
    messages.push({
        role: 'user',
        content: userQuestion
    });
    
    try {
        // Create a placeholder message for streaming effect
        const messagesContainer = document.getElementById('ai-chat-messages');
        const responseDiv = document.createElement('div');
        responseDiv.className = 'ai-message ai-message-bot ai-streaming-response';
        responseDiv.innerHTML = `
            <div class="ai-message-avatar ai-avatar-bot">
                <span>🤖</span>
            </div>
            <div class="ai-message-content">
                <div class="ai-message-text">
                    <span class="streaming-cursor">▋</span>
                </div>
            </div>
        `;
        
        // Remove typing indicator and add streaming response
        removeTypingIndicator(typingId);
        messagesContainer.appendChild(responseDiv);
        scrollChatToBottom();
        
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: selectedModel,
                messages: messages,
                max_tokens: prompt.maxTokens,
                temperature: prompt.temperature || 0.7
            })
        });

        if (!response.ok) {
            responseDiv.remove();
            const errorData = await response.json().catch(() => ({}));
            throw new Error(`OpenAI API error: ${response.status} - ${errorData.error?.message || 'Unknown error'}`);
        }

        const data = await response.json();
        
        if (!data.choices || !data.choices[0] || !data.choices[0].message) {
            responseDiv.remove();
            throw new Error('Invalid response from OpenAI API');
        }

        const aiResponse = data.choices[0].message.content.trim();
        const processedResponse = processAIMessage(aiResponse);
        
        // Simulate typing effect by revealing text gradually
        const textElement = responseDiv.querySelector('.ai-message-text');
        textElement.innerHTML = '';
        
        // Split response into words for smoother streaming effect
        const words = processedResponse.split(' ');
        let currentText = '';
        
        for (let i = 0; i < words.length; i++) {
            currentText += (i === 0 ? '' : ' ') + words[i];
            textElement.innerHTML = currentText + '<span class="streaming-cursor">▋</span>';
            scrollChatToBottom();
            
            // Small delay between words for streaming effect
            await new Promise(resolve => setTimeout(resolve, 50));
        }
        
        // Final update without cursor
        textElement.innerHTML = processedResponse;
        responseDiv.className = 'ai-message ai-message-bot'; // Remove streaming class
        scrollChatToBottom();
        
        // Update conversation history
        conversationHistory.push(
            { role: 'user', content: userQuestion },
            { role: 'assistant', content: aiResponse }
        );
        
        // Keep only last 6 messages (3 exchanges)
        if (conversationHistory.length > 6) {
            conversationHistory = conversationHistory.slice(-6);
        }
        
        // Cache responses for common/short queries
        if (userQuestion.length < 100) {
            responseCache.set(userQuestion.toLowerCase().trim(), aiResponse);
            // Clear old cache entries if too many
            if (responseCache.size > 50) {
                const firstKey = responseCache.keys().next().value;
                responseCache.delete(firstKey);
            }
        }

        return null; // Response already displayed via streaming

    } catch (error) {
        // Remove any streaming UI on error
        const streamingResponse = document.querySelector('.ai-streaming-response');
        if (streamingResponse) {
            streamingResponse.remove();
        }
        
        if (error.message.includes('fetch')) {
            throw new Error('Network error: Please check your internet connection');
        }
        throw error;
    }
}

// Escape HTML for user messages to prevent XSS
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Clean and process AI message content
function processAIMessage(message) {
    // Remove code block markers completely
    let cleanedMessage = message
        .replace(/```html\s*/gi, '')
        .replace(/```\s*/g, '')
        .replace(/^html\s*/i, '')
        .trim();
    
    // If the message starts with HTML tags, ensure it's properly formatted
    if (cleanedMessage.startsWith('<')) {
        // Already HTML content, return as is
        return cleanedMessage;
    }
    
    // Convert newlines to <br> for non-HTML text
    cleanedMessage = cleanedMessage.replace(/\n/g, '<br>');
    
    return cleanedMessage;
}

// Add message to chat using modern card-based styling
function addChatMessage(message, sender) {
    const messagesContainer = document.getElementById('ai-chat-messages');
    const messageDiv = document.createElement('div');
    
    if (sender === 'user') {
        messageDiv.className = 'ai-message ai-message-user';
        messageDiv.innerHTML = `
            <div class="ai-message-content">
                <div class="ai-message-text">
                    ${escapeHtml(message).replace(/\n/g, '<br>')}
                </div>
            </div>
            <div class="ai-message-avatar ai-avatar-user">
                <span>👤</span>
            </div>
        `;
    } else {
        messageDiv.className = 'ai-message ai-message-bot';
        const processedMessage = processAIMessage(message);
        messageDiv.innerHTML = `
            <div class="ai-message-avatar ai-avatar-bot">
                <span>🤖</span>
            </div>
            <div class="ai-message-content">
                <div class="ai-message-text">
                    ${processedMessage}
                </div>
            </div>
        `;
    }
    
    messagesContainer.appendChild(messageDiv);
    
    // Force scroll to bottom with multiple methods for reliability
    setTimeout(() => {
        scrollChatToBottom();
    }, 50);
    
    // Store message
    chatMessages.push({ sender, message, timestamp: Date.now() });
}

// Add typing indicator using modern styling
function addTypingIndicator() {
    const messagesContainer = document.getElementById('ai-chat-messages');
    const typingDiv = document.createElement('div');
    const typingId = 'typing-' + Date.now();
    typingDiv.id = typingId;
    typingDiv.className = 'ai-message ai-message-bot ai-typing-indicator';
    typingDiv.innerHTML = `
        <div class="ai-message-avatar ai-avatar-bot ai-avatar-typing">
            <span>🤖</span>
        </div>
        <div class="ai-message-content">
            <div class="ai-typing-bubble">
                <div class="ai-typing-dots">
                    <span></span><span></span><span></span>
                </div>
            </div>
        </div>
    `;
    messagesContainer.appendChild(typingDiv);
    
    // Scroll to bottom
    setTimeout(() => {
        scrollChatToBottom();
    }, 50);
    
    return typingId;
}

// Remove typing indicator
function removeTypingIndicator(typingId) {
    const indicator = document.getElementById(typingId);
    if (indicator) {
        indicator.remove();
    }
}

// Show error in chat
function showChatError(errorMessage) {
    addChatMessage(`❌ ${errorMessage}`, 'bot');
}

// Scroll chat to bottom with multiple fallback methods
function scrollChatToBottom() {
    const messagesContainer = document.getElementById('ai-chat-messages');
    if (!messagesContainer) return;
    
    // Force immediate scroll without smooth behavior for reliability
    const scrollToBottom = () => {
        // Method 1: Scroll the modal body container
        const modalBody = messagesContainer.closest('.ai-modal-body');
        if (modalBody) {
            modalBody.scrollTop = modalBody.scrollHeight;
        }
        
        // Method 2: Scroll the messages container itself
        messagesContainer.scrollTop = messagesContainer.scrollHeight;
        
        // Method 3: Use scrollIntoView on the last message
        const lastMessage = messagesContainer.lastElementChild;
        if (lastMessage) {
            lastMessage.scrollIntoView({ behavior: 'instant', block: 'end' });
        }
    };
    
    // Execute immediately
    scrollToBottom();
    
    // Execute again after a short delay to handle dynamic content
    setTimeout(scrollToBottom, 10);
    setTimeout(scrollToBottom, 50);
    setTimeout(scrollToBottom, 100);
}

// Add separate AI chat button next to help button
async function addAIChatButton() {
    try {
        if (!isExtensionContextValid()) {
            return;
        }
        
        const hasKey = await hasOpenAIKey();
        if (!hasKey) {
            return;
        }
    } catch (error) {
        return;
    }

    // Remove any existing buttons first to prevent duplicates
    const existingButtons = document.querySelectorAll('.ai-chat-button');
    existingButtons.forEach(button => button.remove());

    // Find the help button to insert after it
    const helpButton = document.querySelector('button[title="Help"]');
    if (!helpButton) {
        return;
    }

    // Create AI chat button with same styling as other header buttons
    const aiChatButton = document.createElement('button');
    aiChatButton.type = 'button';
    aiChatButton.className = 'ms-Button ms-Button--commandBar root-239 ai-chat-button';
    aiChatButton.setAttribute('data-is-focusable', 'true');
    aiChatButton.title = 'Access Packages AI Assistant - Ask questions about access packages, risk levels, and approvals';
    
    // Create button structure matching other header buttons
    aiChatButton.innerHTML = `
        <span class="ms-Button-flexContainer flexContainer-230" data-automationid="splitbuttonprimary">
            <i data-icon-name="Robot" aria-hidden="true" class="ms-Button-icon icon-242 ai-chat-icon">🤖</i>
        </span>
    `;

    // Add click handler
    aiChatButton.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        showChatModal();
    });

    // Insert after help button
    helpButton.insertAdjacentElement('afterend', aiChatButton);
}

// Update chat button visibility based on API key availability
async function updateChatButtonVisibility() {
    try {
        // Check if extension context is valid
        if (!isExtensionContextValid()) {
            return;
        }
        
        const hasKey = await hasOpenAIKey();
        const existingButtons = document.querySelectorAll('.ai-chat-button');
        
        if (hasKey) {
            // If we have a key and no button, or multiple buttons exist, refresh the button
            if (existingButtons.length === 0 || existingButtons.length > 1) {
                // Remove all existing buttons to prevent duplicates
                existingButtons.forEach(button => button.remove());
                // Add a single new button
                addAIChatButton();
            }
        } else {
            // Remove all buttons if we have no key
            existingButtons.forEach(button => button.remove());
        }
    } catch (error) {
        // Extension context issues, silently return
        return;
    }
}

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

// Get selected OpenAI model from storage
async function getOpenAIModel() {
    try {
        // Check if extension context is valid
        if (!isExtensionContextValid()) {
            return 'gpt-4o-mini'; // Default fallback
        }
        const result = await chrome.storage.sync.get('openai_model');
        return result.openai_model || 'gpt-4o-mini'; // Default to gpt-4o-mini if not set
    } catch (error) {
        // Extension context invalidated or other error, return default
        return 'gpt-4o-mini';
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
    const selectedModel = await getOpenAIModel();
    
    try {
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: selectedModel,
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
                aiButton.innerHTML = '✦';
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

    // Check if button already exists
    if (document.querySelector('.ai-enhance-btn')) {
        return;
    }

    // Look for Additional questions modal and business justification textarea
    const additionalQuestionsModal = document.querySelector('h3[title="Additional questions"]');
    if (!additionalQuestionsModal) {
        return;
    }

    // Look for the business justification textarea using multiple selectors
    const textarea = document.querySelector('textarea[aria-labelledby*="TextFieldLabel"]') ||
                    document.querySelector('textarea[id*="TextField"]') ||
                    document.querySelector('.ms-TextField--multiline textarea');
    if (!textarea) {
        return;
    }

    // Find the submit button container - look for the submit button and get its parent
    const submitButton = document.querySelector('button[aria-label="Submit request"]');
    if (!submitButton) {
        return;
    }
    
    const submitContainer = submitButton.parentElement;

    // Create AI enhancement button
    const aiButton = document.createElement('button');
    aiButton.type = 'button';
    aiButton.className = 'ai-enhance-btn';
    aiButton.innerHTML = '✦';
    aiButton.title = 'Enhance business justification using AI';
    aiButton.setAttribute('data-is-focusable', 'true');

    // Add click handler
    aiButton.addEventListener('click', () => handleAIEnhancement(textarea));

    // Insert button next to submit button
    const primarySubmitButton = submitContainer.querySelector('.ms-Button--primary');
    if (primarySubmitButton) {
        primarySubmitButton.insertAdjacentElement('beforebegin', aiButton);
    } else {
        submitContainer.appendChild(aiButton);
    }
}

// Update business justification button visibility based on API key availability
async function updateBusinessJustificationButtonVisibility() {
    try {
        // Check if extension context is valid
        if (!isExtensionContextValid()) {
            return;
        }
        
        const hasKey = await hasOpenAIKey();
        const existingButton = document.querySelector('.ai-enhance-btn');
        
        // Look for Additional questions modal more reliably
        const additionalQuestionsModal = document.querySelector('h3[title="Additional questions"]');
        // Try multiple selectors for the textarea
        const businessJustificationTextarea = document.querySelector('textarea[aria-labelledby*="TextFieldLabel"]') ||
                                            document.querySelector('textarea[id*="TextField"]') ||
                                            document.querySelector('.ms-TextField--multiline textarea');
        
        // Only show button if we have API key and the Additional Questions modal is open
        if (hasKey && additionalQuestionsModal && businessJustificationTextarea && !existingButton) {
            // Add button if we have a key, correct modal, but no button
            addAIButtonToBusinessJustification();
        } else if (!hasKey && existingButton) {
            // Remove button if we have no key but button exists
            existingButton.remove();
        }
    } catch (error) {
        // Extension context issues, silently return
        return;
    }
}

// Start periodic checks for button visibility
function startPeriodicVisibilityCheck() {
    // Check chat button visibility every 5 seconds
    setInterval(() => {
        updateChatButtonVisibility();
    }, 5000);
    
    // Check business justification button visibility more frequently when modal might be open
    setInterval(() => {
        const modal = document.querySelector('.ms-Modal.is-open') || 
                     document.querySelector('[role="dialog"]');
        if (modal && modal.textContent.includes('Additional questions')) {
            updateBusinessJustificationButtonVisibility();
        }
    }, 2000);
}

// Listen for API key updates from popup
if (chrome?.runtime?.onMessage) {
    chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
        try {
            if (message.type === 'API_KEY_UPDATED') {
                // Update chat button visibility immediately
                updateChatButtonVisibility();
                
                // Update business justification button visibility
                updateBusinessJustificationButtonVisibility();
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
                // Update AI button visibility based on API key
                setTimeout(updateBusinessJustificationButtonVisibility, 100);
                setTimeout(updateBusinessJustificationButtonVisibility, 300);
                setTimeout(updateBusinessJustificationButtonVisibility, 600);
            } else {
                // Also check for modal content that might contain "Additional questions"
                const hasAdditionalQuestions = modal.textContent.includes('Additional questions');
                if (hasAdditionalQuestions) {
                    setTimeout(updateBusinessJustificationButtonVisibility, 100);
                    setTimeout(updateBusinessJustificationButtonVisibility, 300);
                    setTimeout(updateBusinessJustificationButtonVisibility, 600);
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
    
    // Start periodic button visibility checks
    startPeriodicVisibilityCheck();
    
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
    
    // Update AI chat button visibility based on API key
    setTimeout(() => {
        updateChatButtonVisibility();
    }, 1000);
    setTimeout(() => {
        updateChatButtonVisibility();
    }, 2000);
    setTimeout(() => {
        updateChatButtonVisibility();
    }, 5000);
    setTimeout(() => {
        updateChatButtonVisibility();
    }, 10000);
    
    // Add global function for manual testing
    window.addAIButton = addAIButtonToBusinessJustification;
    window.showAIChat = showChatModal;
    window.addAIChatButton = addAIChatButton;
    window.updateChatButtonVisibility = updateChatButtonVisibility;
    window.updateBusinessJustificationButtonVisibility = updateBusinessJustificationButtonVisibility;
    window.checkAPIKey = async () => {
        const hasKey = await hasOpenAIKey();
        const key = await getOpenAIKey();
        const additionalQuestionsModal = document.querySelector('h3[title="Additional questions"]');
        const businessJustificationTextarea = document.querySelector('textarea[aria-labelledby*="TextFieldLabel"]') ||
                                            document.querySelector('textarea[id*="TextField"]') ||
                                            document.querySelector('.ms-TextField--multiline textarea');
        const submitButton = document.querySelector('button[aria-label="Submit request"]');
        
        console.log('Has API key:', hasKey);
        console.log('API key exists:', !!key);
        console.log('Chat button exists:', !!document.querySelector('.ai-chat-button'));
        console.log('Enhancement button exists:', !!document.querySelector('.ai-enhance-btn'));
        console.log('Additional questions modal exists:', !!additionalQuestionsModal);
        console.log('Business justification textarea exists:', !!businessJustificationTextarea);
        console.log('Submit button exists:', !!submitButton);
        
        return { 
            hasKey, 
            keyExists: !!key, 
            chatButton: !!document.querySelector('.ai-chat-button'), 
            enhanceButton: !!document.querySelector('.ai-enhance-btn'),
            modalExists: !!additionalQuestionsModal,
            textareaExists: !!businessJustificationTextarea,
            submitButtonExists: !!submitButton
        };
    };
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}