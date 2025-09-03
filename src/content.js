// MyAccess Package approvers Tab Enhanced - Chrome Extension Version
console.log('[MyAccess Enhanced] Chrome Extension loaded');

// ========================================
// UI INITIALIZATION
// ========================================


// ========================================
// AUTHENTICATION
// ========================================

// Get auth token for elm.iga.azure.com API calls
function getAuthToken() {
    // Look for react-auth session tokens
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
                    
                    // Look for elm.iga.azure.com tokens first
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
                    
                    // Check for direct access_token property
                    if (parsed && parsed.access_token) {
                        return parsed.access_token;
                    }
                }
            } catch (e) {
                continue;
            }
        }
        
        // Check for MSAL tokens (fallback)
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
    
    // Last resort: check localStorage for MSAL tokens
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

// Get Microsoft Graph API token
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
                    
                    // Look specifically for graph.microsoft.com tokens
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
    
    // Fallback: try MSAL tokens that might work with Graph
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

// ========================================
// API FUNCTIONS
// ========================================

// Lookup package GUID and metadata by display name
async function lookupPackageGUID(displayName) {
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
            return {
                found: false,
                message: `Cannot display approvers for access package "${displayName}"`
            };
        }

        // Look for exact match first
        const exactMatch = data.value.find(pkg => pkg.displayName === displayName);
        if (exactMatch) {
            return {
                found: true,
                package: exactMatch,
                matches: data.value.length,
                message: `Exact match found (${data.value.length} total matches)`
            };
        }

        // Return first match if no exact match
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

// Get detailed package metadata by GUID
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

// Get assignment policies for a package GUID
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

// Get members of a group from Microsoft Graph API
async function getGroupMembers(groupId) {
    const graphToken = getGraphToken();
    const fallbackToken = getAuthToken();
    
    if (!graphToken && !fallbackToken) {
        return { error: 'No authentication token found', members: [] };
    }

    // Try Graph-specific token first, then fallback token
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

// ========================================
// UTILITY FUNCTIONS
// ========================================

// Create reusable collapsible section HTML
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

// Global event delegation handler for collapsible sections
const collapsibleHandlers = new Map();

// Create reusable click handler for collapsible sections
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
            console.log(`[MyAccess Enhanced] Section header not found for: ${sectionSelector}`);
            return false;
        }
        
        // Mark as ready for delegation
        sectionHeader.setAttribute('data-handler-ready', 'true');
        
        console.log(`[MyAccess Enhanced] Handler ready for: ${sectionSelector}`);
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

// Global event delegation for all collapsible sections
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
        console.log(`[MyAccess Enhanced] No handler config found for: ${sectionSelector}`);
        return;
    }
    
    console.log(`[MyAccess Enhanced] Click handler triggered for: ${sectionSelector}`);
    
    const content = document.getElementById(handlerConfig.contentId);
    const ariaLive = target.querySelector('[aria-live="assertive"]');
    
    if (!content) {
        console.log(`[MyAccess Enhanced] Content area not found: ${handlerConfig.contentId}`);
        return;
    }
    
    const expandedAttr = content.getAttribute('data-expanded');
    const isExpanded = expandedAttr === 'true';
    console.log(`[MyAccess Enhanced] Current state - expanded attr: ${expandedAttr}, isExpanded: ${isExpanded}`);
    
    if (isExpanded) {
        // Collapse
        content.style.display = 'none';
        content.setAttribute('data-expanded', 'false');
        if (ariaLive) ariaLive.textContent = 'Collapsed';
        console.log(`[MyAccess Enhanced] Collapsed: ${sectionSelector}`);
    } else {
        // Expand
        content.style.display = 'block';
        content.setAttribute('data-expanded', 'true');
        if (ariaLive) ariaLive.textContent = 'Expanded';
        console.log(`[MyAccess Enhanced] Expanded: ${sectionSelector}`);
        
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
                console.error(`[MyAccess Enhanced] Error loading content for ${sectionSelector}:`, error);
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

// Global keyboard support for collapsible sections
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


// Generate initials from a name
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

// Create a member card for a user
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

// Create member grid HTML from members array
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

// Extract risk level from display name
function extractRiskLevel(displayName) {
    const riskMatch = displayName.match(/\[(\w+)\]/i);
    if (riskMatch) {
        return riskMatch[1].toLowerCase();
    }
    return 'unknown';
}

// Format date for display
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

// ========================================
// PACKAGE TRACKING
// ========================================

// Store the current package ID when a row is clicked
let currentPackageId = null;
let currentPackageName = null;

// Extract package name from current modal
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

// Intercept clicks on package rows to capture package info
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

// Copy text to clipboard
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

// ========================================
// UI CONTENT CREATION
// ========================================

// Create simple requestors content for Request details tab (non-collapsible)
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
        const lookupResult = await lookupPackageGUID(packageName);
        
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

// Create simple approvers content for Request details tab (non-collapsible)
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
        const lookupResult = await lookupPackageGUID(packageName);
        
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


// ========================================
// TAB MANAGEMENT
// ========================================

// Add collapsible approvers content to Request details tab
function addApproversToRequestDetailsTab() {
    // Look for the Request details tab panel content
    const tabPanel = document.querySelector('[role="tabpanel"]');
    if (!tabPanel) return;
    
    // Check if this is the Request details tab by looking for "Share a link" content
    const shareSection = tabPanel.querySelector('button[id="share-button-id"]');
    if (!shareSection) return;
    
    // Check if we already added approvers
    if (tabPanel.querySelector('[data-request-details-approvers-section]')) return;
    
    console.log('[MyAccess Enhanced] Request details tab found, attempting to add collapsible approvers...');
    
    // Extract package name from current package name
    const packageName = getCurrentPackageName();
    if (!packageName) {
        console.log('[MyAccess Enhanced] Could not extract package name from request details tab');
        return;
    }
    
    console.log('[MyAccess Enhanced] Extracted package name for request details:', packageName);
    
    // Store the package name for createApproversContent
    currentPackageName = packageName;
    
    // Find the container with the share button section
    const shareContainer = shareSection.closest('div');
    if (!shareContainer) {
        console.log('[MyAccess Enhanced] Could not find share container');
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

// [REMOVED] addApproversToResourcesTab function - replaced with collapsible implementation

// Add collapsible approver information to access request panel
function addApproversToAccessRequestPanel() {
    // Look for the access request panel
    const accessRequestPanel = document.querySelector('.ms-Panel-content');
    if (!accessRequestPanel) return;
    
    // Check if this is an access request panel by looking for "Access request" header
    const headerElement = document.querySelector('.ms-Panel-headerText');
    if (!headerElement || !headerElement.textContent.includes('Access request')) return;
    
    // Check if we already added approvers section
    if (accessRequestPanel.querySelector('[data-access-request-approvers-section]')) return;
    
    console.log('[MyAccess Enhanced] Access request panel found, attempting to add collapsible approvers...');
    
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
        console.log('[MyAccess Enhanced] Could not extract package name from access request, using fallback');
        packageName = 'Unknown Package';
    }
    
    console.log('[MyAccess Enhanced] Extracted package name:', packageName);
    
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


// Monitor for modal opening
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
                        
                        console.log('[MyAccess Enhanced] Found access request panel, attempting to add button...');
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

// ========================================
// INITIALIZATION
// ========================================

// Initialize the script
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
    
    // Debug helpers
    window.debugRequestDetailsTab = () => {
        console.log('=== Debug Request Details Tab ===');
        const tabPanel = document.querySelector('[role="tabpanel"]');
        console.log('Tab panel found:', !!tabPanel);
        
        const shareSection = tabPanel ? tabPanel.querySelector('button[id="share-button-id"]') : null;
        console.log('Share button found:', !!shareSection);
        
        const shareContainer = shareSection ? shareSection.closest('div') : null;
        console.log('Share container found:', !!shareContainer);
        
        const packageName = getCurrentPackageName();
        console.log('Package name:', packageName);
        
        addApproversToRequestDetailsTab();
    };
    window.debugAccessRequestPanel = () => {
        console.log('=== Debug Access Request Panel ===');
        const headerElement = document.querySelector('.ms-Panel-headerText');
        console.log('Header found:', !!headerElement);
        console.log('Header text:', headerElement ? headerElement.textContent : 'N/A');
        
        const panel = document.querySelector('.ms-Panel-content');
        console.log('Panel found:', !!panel);
        if (panel) {
            console.log('Panel content includes "You submitted a request for":', panel.textContent.includes('You submitted a request for'));
            console.log('Panel text content:', panel.textContent.substring(0, 300));
            
            const activityTexts = panel.querySelectorAll('.ms-ActivityItem-activityText');
            console.log('Activity texts found:', activityTexts.length);
            activityTexts.forEach((text, i) => {
                console.log(`Activity ${i}:`, text.textContent);
            });
        }
        addApproversToAccessRequestPanel();
    };
}

// Start when page is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}