// MyAccess Package Approvers Tab Enhanced - Chrome Extension Version
console.log('[content.js] MyAccess Enhanced Chrome Extension loaded');

// ========================================
// UI INITIALIZATION
// ========================================

// Add test banner to verify extension is working
function addTestBanner() {
    if (!document.querySelector('.tampermonkey-test-banner')) {
        const banner = document.createElement('div');
        banner.className = 'tampermonkey-test-banner';
        banner.textContent = '🚀 Enhanced MyAccess Script - GUID Lookup Enabled 🚀';
        document.body.appendChild(banner);
        
        // Remove after 5 seconds
        setTimeout(() => {
            banner.style.transition = 'opacity 1s';
            banner.style.opacity = '0';
            setTimeout(() => banner.remove(), 1000);
        }, 5000);
    }
}

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
                message: `No packages found matching "${displayName}"`
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

// Create approvers tab content with GUID lookup
async function createApproversContent() {
    const container = document.createElement('div');
    container.className = 'approvers-tab-content';
    
    // Create progress bar container
    const progressContainer = document.createElement('div');
    progressContainer.innerHTML = `
        <div style="margin-bottom: 20px;">
            <div style="font-size: 14px; color: #323130; margin-bottom: 8px;">Loading package information...</div>
            <div style="background: #f3f2f1; border-radius: 8px; height: 8px; overflow: hidden;">
                <div id="progress-bar" style="background: linear-gradient(90deg, #0078d4, #40e0d0); height: 100%; width: 0%; transition: width 0.3s ease; border-radius: 8px;"></div>
            </div>
            <div id="progress-text" style="font-size: 12px; color: #605e5c; margin-top: 4px;">Initializing...</div>
        </div>
    `;
    container.appendChild(progressContainer);
    
    const updateProgress = (percentage, text) => {
        const progressBar = document.getElementById('progress-bar');
        const progressText = document.getElementById('progress-text');
        if (progressBar) progressBar.style.width = percentage + '%';
        if (progressText) progressText.textContent = text;
    };

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
                Please try:<br>
                1. Refresh the page<br>
                2. Sign out and sign back in<br>
                3. Run <code>window.debugTokens()</code> in console for details
            </div>`;
        return container;
    }

    try {
        updateProgress(10, 'Looking up package GUID...');
        // Look up the package GUID and metadata
        const lookupResult = await lookupPackageGUID(packageName);
        
        updateProgress(30, 'Found package, fetching details...');
        container.removeChild(progressContainer);

        if (lookupResult.found) {
            // Restore progress bar for remaining calls
            container.appendChild(progressContainer);
            updateProgress(40, 'Fetching package metadata...');
            
            // Get full package details and assignment policies
            const packageDetails = await getPackageDetails(lookupResult.package.id);
            const pkg = packageDetails || lookupResult.package;
            
            updateProgress(60, 'Loading assignment policies...');
            const assignmentPolicies = await getAssignmentPolicies(pkg.id);
            
            updateProgress(70, 'Processing approval workflows...');
            
            // Collect all group IDs from approval stages
            const groupIds = new Set();
            for (const policy of assignmentPolicies) {
                const approvalSettings = policy.approvalSettings || policy.requestApprovalSettings;
                let stages = [];
                
                if (approvalSettings?.stages) {
                    stages = approvalSettings.stages;
                } else if (approvalSettings?.approvalStages) {
                    stages = approvalSettings.approvalStages;
                }
                
                for (const stage of stages) {
                    if (stage.primaryApprovers) {
                        stage.primaryApprovers.forEach(approver => {
                            const groupId = approver.groupId || approver.id || approver.objectId;
                            if (groupId) groupIds.add(groupId);
                        });
                    }
                    if (stage.escalationApprovers) {
                        stage.escalationApprovers.forEach(approver => {
                            const groupId = approver.groupId || approver.id || approver.objectId;
                            if (groupId) groupIds.add(groupId);
                        });
                    }
                }
            }
            
            updateProgress(80, `Loading group members (${groupIds.size} groups)...`);
            
            // Fetch all group members
            const groupMembersCache = new Map();
            const groupArray = Array.from(groupIds);
            
            for (let i = 0; i < groupArray.length; i++) {
                const groupId = groupArray[i];
                const progress = 80 + (i + 1) * (15 / groupArray.length);
                updateProgress(progress, `Loading members for group ${i + 1}/${groupArray.length}...`);
                
                try {
                    const memberData = await getGroupMembers(groupId);
                    groupMembersCache.set(groupId, memberData);
                } catch (error) {
                    console.warn(`Failed to load members for group ${groupId}:`, error);
                    groupMembersCache.set(groupId, { error: error.message, members: [] });
                }
            }
            
            updateProgress(95, 'Finalizing display...');
            
            // Remove progress bar after all API calls
            container.removeChild(progressContainer);


            // Add simplified approvers section
            const approversSection = document.createElement('div');
            approversSection.style.marginTop = '20px';
            
            if (assignmentPolicies && assignmentPolicies.length > 0) {
                let approversHtml = '<h4 style="margin: 0 0 16px 0; color: #323130;">👥 Approvers</h4>';
                
                const allApproverGroups = new Set();
                
                // Collect all approver groups from all policies and stages
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
                
                // Display each unique approver group
                const uniqueGroups = Array.from(allApproverGroups).map(g => JSON.parse(g));
                
                for (const group of uniqueGroups) {
                    const memberData = groupMembersCache.get(group.groupId);
                    
                    approversHtml += `
                        <div style="margin-bottom: 24px;">
                            <div style="font-size: 15px; color: #0078d4; font-weight: 500; margin-bottom: 12px; padding: 8px 12px; background: #f8f9fa; border-left: 3px solid #0078d4; border-radius: 4px;">${group.displayName}</div>
                    `;
                    
                    if (memberData && !memberData.error && memberData.members.length > 0) {
                        approversHtml += createMemberGrid(memberData.members);
                    } else if (memberData && memberData.error) {
                        if (memberData.error.includes('Authentication failed')) {
                            approversHtml += `
                                <div style="padding: 12px; background: #fff4ce; border: 1px solid #ffb900; border-radius: 6px; margin-top: 8px;">
                                    <div style="font-size: 12px; color: #8a6914; margin-bottom: 4px;"><strong>⚠️ Graph API Access Required</strong></div>
                                    <div style="font-size: 11px; color: #8a6914; line-height: 1.4;">
                                        To see group members, Microsoft Graph API access is needed.<br>
                                        <strong>Workaround:</strong> Run this command in a separate terminal:<br>
                                        <code style="background: #f3f2f1; padding: 2px 4px; border-radius: 2px; font-family: monospace;">az rest --method GET --url "https://graph.microsoft.com/v1.0/groups/${group.groupId}/members"</code>
                                    </div>
                                </div>
                            `;
                        } else {
                            approversHtml += `
                                <div style="margin-top: 8px; font-size: 12px; color: #a80000; font-style: italic;">
                                    ⚠️ Could not load group members: ${memberData.error}
                                </div>
                            `;
                        }
                    } else if (memberData && memberData.members.length === 0) {
                        approversHtml += `
                            <div style="margin-top: 8px; font-size: 12px; color: #605e5c; font-style: italic;">
                                No members found in this group
                            </div>
                        `;
                    } else {
                        approversHtml += `
                            <div style="margin-top: 8px; font-size: 12px; color: #8a6914; font-style: italic;">
                                Loading group members...
                            </div>
                        `;
                    }
                    
                    approversHtml += '</div>';
                }
                
                if (uniqueGroups.length === 0) {
                    approversHtml += `
                        <div style="font-style: italic; color: #605e5c; margin-top: 16px;">
                            No approver groups configured for this package.
                        </div>
                    `;
                }
                
                approversSection.innerHTML = approversHtml;
            } else {
                approversSection.innerHTML = `
                    <h4 style="margin: 0 0 16px 0; color: #323130;">👥 Approvers</h4>
                    <div style="padding: 12px; background: #fff4ce; border: 1px solid #ffb900; border-radius: 6px;">
                        ⚠️ No assignment policies found for this package.
                    </div>
                `;
            }
            
            container.appendChild(approversSection);
            
            // Add requestor groups section
            const requestorSection = document.createElement('div');
            requestorSection.style.marginTop = '20px';
            
            if (assignmentPolicies && assignmentPolicies.length > 0) {
                let requestorHtml = '<h4 style="margin: 0 0 16px 0; color: #323130;">👤 Who Can Request</h4>';
                
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
                
                if (uniqueRequestorGroups.length > 0) {
                    for (const groupName of uniqueRequestorGroups) {
                        requestorHtml += `
                            <div style="margin-bottom: 12px;">
                                <div style="font-size: 15px; color: #107c10; font-weight: 500; margin-bottom: 4px; padding: 8px 12px; background: #f8f9fa; border-left: 3px solid #107c10; border-radius: 4px;">${groupName}</div>
                            </div>
                        `;
                    }
                } else {
                    requestorHtml += `
                        <div style="font-style: italic; color: #605e5c;">
                            No specific requestor groups configured - may be open to all users.
                        </div>
                    `;
                }
                
                requestorSection.innerHTML = requestorHtml;
            } else {
                requestorSection.innerHTML = `
                    <h4 style="margin: 0 0 16px 0; color: #323130;">👤 Who Can Request</h4>
                    <div style="font-style: italic; color: #605e5c;">
                        No assignment policies found.
                    </div>
                `;
            }
            
            container.appendChild(requestorSection);

        } else {
            // Show not found message but no separate box
            container.innerHTML += `<div class="warning-message">⚠️ ${lookupResult.message}</div>`;
        }

    } catch (error) {
        container.innerHTML = `
            <div class="error-message">
                <strong>Error looking up package GUID:</strong><br>
                ${error.message}<br><br>
                <strong>Package Name:</strong> ${packageName}<br>
                <strong>Token Available:</strong> ${token ? 'Yes' : 'No'}
            </div>`;
    }

    return container;
}

// ========================================
// TAB MANAGEMENT
// ========================================

// Add Approvers tab to the modal
function addApproversTab() {
    // Look specifically for the modal's tab list
    let tabList = null;
    let modalTabLists = document.querySelectorAll('[role="tablist"]');
    
    for (let list of modalTabLists) {
        // Check if this tab list contains the Request details and Resources tabs
        const tabs = list.querySelectorAll('[role="tab"]');
        let hasRequestDetails = false;
        let hasResources = false;
        
        for (let tab of tabs) {
            if (tab.textContent.includes('Request details')) hasRequestDetails = true;
            if (tab.textContent.includes('Resources')) hasResources = true;
        }
        
        // This is the modal's tab list
        if (hasRequestDetails && hasResources) {
            tabList = list;
            break;
        }
    }
    
    if (!tabList) return;
    
    // Check if we already added the tab
    if (tabList.querySelector('[data-approvers-tab]')) return;
    
    // Find existing tabs in the modal
    const existingTabs = Array.from(tabList.querySelectorAll('[role="tab"]'));
    const resourcesTab = existingTabs.find(tab => tab.textContent.includes('Resources'));
    
    if (!resourcesTab) return;
    
    // Clone the Resources tab
    const approversTab = resourcesTab.cloneNode(true);
    
    // Update text content - handle nested elements
    const textElement = approversTab.querySelector('.ms-Pivot-text') || 
                      approversTab.querySelector('span');
    if (textElement) {
        textElement.textContent = ' Approvers';
    } else {
        approversTab.textContent = 'Approvers';
    }
    
    approversTab.setAttribute('data-approvers-tab', 'true');
    
    // Reset selection state
    if (approversTab.hasAttribute('aria-selected')) {
        approversTab.setAttribute('aria-selected', 'false');
    }
    
    // Remove any active/selected classes
    const classesToRemove = ['is-selected', 'selected', 'active', 'is-active'];
    classesToRemove.forEach(cls => {
        approversTab.classList.remove(cls);
        const nested = approversTab.querySelectorAll('*');
        nested.forEach(el => el.classList.remove(cls));
    });
    
    // Add to tab list
    tabList.appendChild(approversTab);
    
    // Add click handler
    approversTab.addEventListener('click', async function(e) {
        e.preventDefault();
        e.stopPropagation();
        
        // Only clear selection from other tabs, don't forcefully style them
        // Let native tabs manage their own styling
        const allTabs = tabList.querySelectorAll('[role="tab"]');
        allTabs.forEach(tab => {
            if (tab !== approversTab) {
                tab.setAttribute('aria-selected', 'false');
                
                // Only remove selection classes from other tabs, let them keep their other styling
                const classesToRemove = [];
                tab.classList.forEach(className => {
                    if (className.startsWith('linkIsSelected') || className === 'is-selected' || className === 'ms-Pivot-link--selected') {
                        classesToRemove.push(className);
                    }
                });
                classesToRemove.forEach(className => tab.classList.remove(className));
                
                // Remove underline from other tabs
                const textElement = tab.querySelector('.ms-Pivot-text') || tab.querySelector('span') || tab;
                if (textElement) {
                    textElement.style.borderBottom = '';
                    textElement.style.borderBottomLeftRadius = '';
                    textElement.style.borderBottomRightRadius = '';
                }
            }
        });
        
        // Find the tab panel first
        let tabPanel = document.querySelector('[role="tabpanel"]');
        
        if (!tabPanel) {
            // Try to find content area another way
            const possiblePanels = document.querySelectorAll('div[class*="content"], div[class*="panel"], div[class*="body"]');
            for (let panel of possiblePanels) {
                if (panel.textContent.includes('Share a link') ||
                    panel.querySelector('button[id="share-button-id"]')) {
                    tabPanel = panel.parentElement;
                    console.log('Found panel by content search');
                    break;
                }
            }
        }

        // Select the Approvers tab with proper styling AND load content (use same delay as existing tabs)
        setTimeout(async () => {
            approversTab.setAttribute('aria-selected', 'true');
            approversTab.classList.remove('link-321');
            approversTab.classList.add('is-selected', 'linkIsSelected-315', 'ms-Pivot-link--selected');
            
            // Add underline styling to match native tabs
            const textElement = approversTab.querySelector('.ms-Pivot-text') || approversTab.querySelector('span') || approversTab;
            if (textElement) {
                textElement.style.borderBottom = '2px solid #0078d4';
                textElement.style.borderBottomLeftRadius = '0';
                textElement.style.borderBottomRightRadius = '0';
            }
            
            // Now load content (after aria-selected is set)
            if (tabPanel) {
                tabPanel.setAttribute('aria-labelledby', approversTab.id || 'approvers-tab');
                tabPanel.innerHTML = '';
                
                // Load content and check if still selected after async operation
                const content = await createApproversContent();
                
                // Double-check we're still on the approvers tab after async operation
                if (approversTab.getAttribute('aria-selected') === 'true') {
                    tabPanel.appendChild(content);
                }
                // If not, the content is discarded (user switched tabs)
            } else {
                // Load content and check if still selected after async operation
                const content = await createApproversContent();
                
                // Double-check we're still on the approvers tab after async operation
                if (approversTab.getAttribute('aria-selected') === 'true') {
                    content.style.marginTop = '20px';
                    tabList.parentElement.appendChild(content);
                }
                // If not, the content is discarded (user switched tabs)
            }
        }, 10);
    });
    
    // Add lightweight handlers to existing tabs that only manage approvers tab cleanup
    existingTabs.forEach(tab => {
        // Store original click handler before adding our own
        const originalClickHandlers = [];
        
        // Add our handler with capture=true to run before native handlers
        tab.addEventListener('click', function(e) {
            // Only clear the approvers tab selection, let native tabs handle themselves
            if (approversTab && approversTab.getAttribute('aria-selected') === 'true') {
                approversTab.setAttribute('aria-selected', 'false');
                
                // Remove approvers tab selection styling
                const approversClassesToRemove = [];
                approversTab.classList.forEach(className => {
                    if (className.startsWith('linkIsSelected') || className === 'is-selected' || className === 'ms-Pivot-link--selected') {
                        approversClassesToRemove.push(className);
                    }
                });
                approversClassesToRemove.forEach(className => approversTab.classList.remove(className));
                
                // Ensure approvers tab has unselected classes
                if (!approversTab.classList.contains('ms-Button')) approversTab.classList.add('ms-Button');
                if (!approversTab.classList.contains('ms-Button--action')) approversTab.classList.add('ms-Button--action');
                if (!approversTab.classList.contains('ms-Button--command')) approversTab.classList.add('ms-Button--command');
                if (!approversTab.classList.contains('ms-Pivot-link')) approversTab.classList.add('ms-Pivot-link');
                if (!approversTab.classList.contains('link-321')) approversTab.classList.add('link-321');
                
                // Remove underline from approvers tab
                const approversTextElement = approversTab.querySelector('.ms-Pivot-text') || approversTab.querySelector('span') || approversTab;
                if (approversTextElement) {
                    approversTextElement.style.borderBottom = '';
                    approversTextElement.style.borderBottomLeftRadius = '';
                    approversTextElement.style.borderBottomRightRadius = '';
                }
            }
            
            // Let the native click handler proceed normally for the clicked tab
            // DON'T prevent default or stop propagation - let Microsoft handle the tab switch
        }, true); // Use capture=true to ensure we run before native handlers
    });
    
    // Try to trigger a click on request details to ensure UI is ready
    const requestTab = existingTabs.find(tab => tab.textContent.includes('Request details'));
    if (requestTab) {
        setTimeout(() => {
            requestTab.click();
        }, 200);
    }
}

// Add approver information to access request panel
function addApproversToAccessRequestPanel() {
    // Look for the access request panel
    const accessRequestPanel = document.querySelector('.ms-Panel-content');
    if (!accessRequestPanel) return;
    
    // Check if this is an access request panel by looking for "Access request" header
    const headerElement = document.querySelector('.ms-Panel-headerText');
    if (!headerElement || !headerElement.textContent.includes('Access request')) return;
    
    // Check if we already added approvers link/section
    if (accessRequestPanel.querySelector('[data-approvers-section]') || 
        accessRequestPanel.querySelector('.view-approvers-link')) return;
    
    console.log('[MyAccess Enhanced] Access request panel found, attempting to add approvers...');
    
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
    
    // Store the package name for createApproversContent
    currentPackageName = packageName;
    
    // Create View Approvers link (styled like existing Details links)
    const viewApproversLink = document.createElement('button');
    viewApproversLink.textContent = '👥 View Approvers';
    viewApproversLink.className = 'ms-Link view-approvers-link';
    viewApproversLink.type = 'button';
    viewApproversLink.style.cssText = `
        background: none;
        border: none;
        color: #0078d4;
        cursor: pointer;
        font-size: 14px;
        text-decoration: underline;
        padding: 0;
        margin: 4px 0;
        font-family: inherit;
        display: inline;
    `;
    
    // Add hover effect to match existing Details links
    viewApproversLink.addEventListener('mouseenter', () => {
        viewApproversLink.style.color = '#106ebe';
    });
    viewApproversLink.addEventListener('mouseleave', () => {
        viewApproversLink.style.color = '#0078d4';
    });
    
    // Create approvers section - initially hidden
    const approversSection = document.createElement('div');
    approversSection.setAttribute('data-approvers-section', 'true');
    approversSection.style.cssText = `
        margin-top: 12px;
        border-top: 1px solid #edebe9;
        padding-top: 20px;
        display: none;
    `;
    
    // Add loading message initially
    approversSection.innerHTML = `
        <div style="font-size: 16px; font-weight: 600; color: #323130; margin-bottom: 12px;">👥 Package Approvers</div>
        <div style="font-size: 14px; color: #605e5c;">Loading approver information...</div>
    `;
    
    // Add link click handler to toggle approvers section
    let isLoaded = false;
    viewApproversLink.addEventListener('click', async (e) => {
        e.preventDefault(); // Prevent any default link behavior
        e.stopPropagation(); // Stop event bubbling
        
        if (approversSection.style.display === 'none') {
            // Show the section
            approversSection.style.display = 'block';
            viewApproversLink.textContent = '👥 Hide Approvers';
            
            // Load content if not already loaded
            if (!isLoaded) {
                try {
                    const approverContent = await createApproversContent();
                    approversSection.innerHTML = `
                        <div style="font-size: 16px; font-weight: 600; color: #323130; margin-bottom: 12px;">👥 Package Approvers</div>
                    `;
                    approversSection.appendChild(approverContent);
                    isLoaded = true;
                } catch (error) {
                    approversSection.innerHTML = `
                        <div style="font-size: 16px; font-weight: 600; color: #323130; margin-bottom: 12px;">👥 Package Approvers</div>
                        <div class="error-message">
                            <strong>Error loading approver information:</strong><br>
                            ${error.message}
                        </div>
                    `;
                    isLoaded = true;
                }
            }
        } else {
            // Hide the section
            approversSection.style.display = 'none';
            viewApproversLink.textContent = '👥 View Approvers';
        }
    });
    
    // Append link and section to panel (simple positioning)
    accessRequestPanel.appendChild(viewApproversLink);
    accessRequestPanel.appendChild(approversSection);
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
                setTimeout(addApproversTab, 50);
                setTimeout(addApproversTab, 200);
                setTimeout(addApproversTab, 500);
            }
        }
        
        // Check for access request panel
        const headerElement = document.querySelector('.ms-Panel-headerText');
        if (headerElement && headerElement.textContent.includes('Access request')) {
            const accessRequestPanel = document.querySelector('.ms-Panel-content');
            if (accessRequestPanel) {
                const currentContent = accessRequestPanel.textContent;
                
                // Check if we have an access request panel and it's not already processed
                if (!accessRequestPanel.querySelector('.view-approvers-link') && 
                    !accessRequestPanel.querySelector('[data-approvers-section]')) {
                    
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
    addTestBanner();
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
    window.debugApproversTab = () => addApproversTab();
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