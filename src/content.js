// MyAccess Package approvers Tab Enhanced - Chrome Extension Version
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

// Create reusable approvers section HTML
function createApproversSection(sectionId, contentId) {
    return `
        <div class="ms-List-cell" data-list-index="0" style="margin-bottom: 8px;">
            <div class="approvers-wrapper" style="background: #f8f9fa; border: 1px solid #edebe9; border-radius: 4px; overflow: hidden;">
                <div class="collapsible-section" data-section="${sectionId}" 
                     role="button" data-is-focusable="true" tabindex="0"
                     style="cursor: pointer; padding: 12px; background: transparent;">
                    <div class="_2Wk_euHnEs1IEUewKXsi4D">
                        <div class="_2Mud-yt14z94cvqhazDnLC" style="font-size: 16px; color: #323130; font-weight: 600;">Who can approve?</div>
                        <div class="hMwWTN0DSGG3aIKsHkP3f" style="font-size: 11px; color: #605e5c;">Show/hide details</div>
                    </div>
                    <div aria-live="assertive" style="position: absolute; width: 1px; height: 1px; margin: -1px; padding: 0px; border: 0px; overflow: hidden; clip: rect(0px, 0px, 0px, 0px);">Collapsed</div>
                </div>
                <div class="collapsible-content" id="${contentId}" style="display: none; padding: 0 12px 12px 12px; background: transparent;">
                    <div class="approver-content-loading" style="padding: 20px; text-align: center; color: #605e5c;">
                        Click to load approver details...
                    </div>
                </div>
            </div>
        </div>
    `;
}

// Create reusable approvers click handler
function createApproversClickHandler(sectionSelector, contentId) {
    const sectionHeader = document.querySelector(`.collapsible-section[data-section="${sectionSelector}"]`);
    let approversLoaded = false;
    
    if (!sectionHeader) return;
    
    sectionHeader.addEventListener('click', async function(e) {
        e.preventDefault();
        e.stopPropagation();
        
        const content = document.getElementById(contentId);
        const ariaLive = this.querySelector('[aria-live="assertive"]');
        
        if (content) {
            const isExpanded = content.style.display !== 'none';
            
            if (isExpanded) {
                // Collapse
                content.style.display = 'none';
                ariaLive.textContent = 'Collapsed';
            } else {
                // Expand
                content.style.display = 'block';
                ariaLive.textContent = 'Expanded';
                
                // Load approvers data only when expanding for the first time
                if (!approversLoaded) {
                    approversLoaded = true;
                    
                    // Show loading message
                    content.innerHTML = `
                        <div style="padding: 20px; text-align: center;">
                            <div style="font-size: 14px; color: #323130; margin-bottom: 8px;">Loading who can approve...</div>
                            <div style="background: #f3f2f1; border-radius: 8px; height: 4px; width: 200px; margin: 0 auto; overflow: hidden;">
                                <div style="background: linear-gradient(90deg, #0078d4, #40e0d0); height: 100%; width: 100%; animation: pulse 1.5s ease-in-out infinite;"></div>
                            </div>
                        </div>
                    `;
                    
                    try {
                        // Create a simple approvers content
                        const approverContent = await createSimpleApproversContent();
                        
                        // Update content with loaded data
                        content.innerHTML = '';
                        content.appendChild(approverContent);
                        
                    } catch (error) {
                        content.innerHTML = `
                            <div style="padding: 12px; background: #fde7e9; border: 1px solid #a80000; border-radius: 6px;">
                                <strong>Error loading approvers:</strong><br>
                                ${error.message}
                            </div>
                        `;
                    }
                }
            }
        }
    });
    
    // Add keyboard support
    sectionHeader.addEventListener('keydown', function(e) {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            this.click();
        }
    });
    
    // Prevent clicks on the content area from bubbling up
    const contentArea = document.getElementById(contentId);
    if (contentArea) {
        contentArea.addEventListener('click', function(e) {
            e.stopPropagation();
        });
    }
}

// Create consistent loading message
function createLoadingMessage(title = 'Loading who can approve...', subtitle = 'Please wait while we fetch approver information') {
    return `
        <div style="padding: 40px; text-align: center;">
            <div style="font-size: 16px; color: #323130; margin-bottom: 12px;">${title}</div>
            <div style="font-size: 14px; color: #605e5c;">${subtitle}</div>
            <div style="margin-top: 16px;">
                <div style="background: #f3f2f1; border-radius: 8px; height: 4px; width: 200px; margin: 0 auto; overflow: hidden;">
                    <div style="background: linear-gradient(90deg, #0078d4, #40e0d0); height: 100%; width: 100%; animation: pulse 1.5s ease-in-out infinite;"></div>
                </div>
            </div>
        </div>
    `;
}

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
                                <div class="simple-approver-group" style="margin-bottom: 24px;">
                                    <div style="font-size: 15px; color: #0078d4; font-weight: 500; margin-bottom: 12px; padding: 8px 12px; background: #f8f9fa; border-left: 3px solid #0078d4; border-radius: 4px;">${group.displayName}</div>
                            `;
                            
                            // Show all members by default
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
                            
                        } catch (error) {
                            approversHtml += `
                                <div class="simple-approver-group" style="margin-bottom: 24px;">
                                    <div style="font-size: 15px; color: #0078d4; font-weight: 500; margin-bottom: 12px; padding: 8px 12px; background: #f8f9fa; border-left: 3px solid #0078d4; border-radius: 4px;">${group.displayName}</div>
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
                    <div style="padding: 12px; background: #fff4ce; border: 1px solid #ffb900; border-radius: 6px;">
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

// Create simple approvers content for Resources tab and Access panels (non-collapsible)
async function createApproversContent() {
    const container = document.createElement('div');
    container.className = 'approvers-tab-content';

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
                
                // Create simple approvers section
                let approversHtml = '<h4 style="margin: 0 0 16px 0; color: #323130;">👥 Who can approve?</h4>';
                
                // Display each unique approver group with all members shown by default
                const uniqueGroups = Array.from(allApproverGroups).map(g => JSON.parse(g));
                
                if (uniqueGroups.length > 0) {
                    for (let i = 0; i < uniqueGroups.length; i++) {
                        const group = uniqueGroups[i];
                        
                        try {
                            const memberData = await getGroupMembers(group.groupId);
                            
                            approversHtml += `
                                <div class="simple-approver-group" style="margin-bottom: 24px;">
                                    <div style="font-size: 15px; color: #0078d4; font-weight: 500; margin-bottom: 12px; padding: 8px 12px; background: #f8f9fa; border-left: 3px solid #0078d4; border-radius: 4px;">${group.displayName}</div>
                            `;
                            
                            // Show all members by default
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
                            
                        } catch (error) {
                            approversHtml += `
                                <div class="simple-approver-group" style="margin-bottom: 24px;">
                                    <div style="font-size: 15px; color: #0078d4; font-weight: 500; margin-bottom: 12px; padding: 8px 12px; background: #f8f9fa; border-left: 3px solid #0078d4; border-radius: 4px;">${group.displayName}</div>
                                    <div style="margin-top: 8px; font-size: 12px; color: #a80000; font-style: italic;">
                                        ⚠️ Could not load group members: ${error.message}
                                    </div>
                                </div>
                            `;
                        }
                    }
                } else {
                    approversHtml += `
                        <div style="padding: 12px; background: #fff4ce; border: 1px solid #ffb900; border-radius: 6px;">
                            ⚠️ No approver groups configured for this package.
                        </div>
                    `;
                }
                
                // Add requestor groups section
                approversHtml += '<h4 style="margin: 20px 0 16px 0; color: #323130;">👤 Who Can Request</h4>';
                
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
                        approversHtml += `
                            <div style="margin-bottom: 12px;">
                                <div style="font-size: 15px; color: #107c10; font-weight: 500; margin-bottom: 4px; padding: 8px 12px; background: #f8f9fa; border-left: 3px solid #107c10; border-radius: 4px;">${groupName}</div>
                            </div>
                        `;
                    }
                } else {
                    approversHtml += `
                        <div style="padding: 12px; background: #e6f7ff; border: 1px solid #0078d4; border-radius: 6px;">
                            <strong>ℹ️ Open Access</strong><br>
                            No specific requestor groups configured - this package may be open to all users.
                        </div>
                    `;
                }
                
                container.innerHTML = approversHtml;
            } else {
                container.innerHTML = `
                    <h4 style="margin: 0 0 16px 0; color: #323130;">👥 Who can approve?</h4>
                    <div style="padding: 12px; background: #fff4ce; border: 1px solid #ffb900; border-radius: 6px;">
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
    
    // Use reusable approvers section
    approversSection.innerHTML = createApproversSection('request-details-approvers', 'request-details-approvers-content');
    
    // Insert before the share container
    shareContainer.insertAdjacentElement('beforebegin', approversSection);
    
    // Use reusable click handler
    createApproversClickHandler('request-details-approvers', 'request-details-approvers-content');
}

// Add approvers content to Resources tab (keeping for backwards compatibility)
function addApproversToResourcesTab() {
    // Look for the Resources tab panel content
    const tabPanel = document.querySelector('[role="tabpanel"]');
    if (!tabPanel) return;
    
    // Check if this is the Resources tab by looking for "View by:" dropdown
    const viewByLabel = tabPanel.querySelector('label[id*="Dropdown"][id$="-label"]');
    if (!viewByLabel || !viewByLabel.textContent.includes('View by:')) return;
    
    // Check if we already added approvers
    if (tabPanel.querySelector('[data-resources-approvers-section]')) return;
    
    console.log('[MyAccess Enhanced] Resources tab found, attempting to add approvers...');
    
    // Extract package name from current package name
    const packageName = getCurrentPackageName();
    if (!packageName) {
        console.log('[MyAccess Enhanced] Could not extract package name from resources tab');
        return;
    }
    
    console.log('[MyAccess Enhanced] Extracted package name for resources:', packageName);
    
    // Store the package name for createApproversContent
    currentPackageName = packageName;
    
    // Find the container after the dropdown and resources list
    const resourcesContainer = tabPanel.querySelector('.ms-FocusZone');
    if (!resourcesContainer) {
        console.log('[MyAccess Enhanced] Could not find resources container');
        return;
    }
    
    // Create approvers section with minimal spacing
    const approversSection = document.createElement('div');
    approversSection.setAttribute('data-resources-approvers-section', 'true');
    approversSection.style.cssText = `
        margin-top: 0px;
        border-top: 1px solid #edebe9;
        padding-top: 4px;
    `;
    
    // Add header
    const approversHeader = document.createElement('div');
    approversHeader.style.cssText = `
        font-size: 16px;
        font-weight: 600;
        color: #323130;
        margin-bottom: 8px;
        margin-top: 8px;
    `;
    approversHeader.textContent = '👥 Who can approve?';
    
    // Add loading message
    const loadingDiv = document.createElement('div');
    loadingDiv.innerHTML = createLoadingMessage('Loading who can approve...', 'Fetching approver information for this access package');
    
    approversSection.appendChild(approversHeader);
    approversSection.appendChild(loadingDiv);
    
    // Insert directly after the resources container within the same parent
    resourcesContainer.insertAdjacentElement('afterend', approversSection);
    
    // Load approver content immediately
    (async () => {
        try {
            const approverContent = await createApproversContent();
            
            // Replace loading with actual content
            approversSection.removeChild(loadingDiv);
            approversSection.appendChild(approverContent);
        } catch (error) {
            // Replace loading with error message
            const errorDiv = document.createElement('div');
            errorDiv.innerHTML = `
                <div style="padding: 20px; text-align: center;">
                    <div style="font-size: 16px; color: #a80000; margin-bottom: 12px;">⚠️ Error loading approvers</div>
                    <div style="font-size: 14px; color: #605e5c;">${error.message}</div>
                </div>
            `;
            
            approversSection.removeChild(loadingDiv);
            approversSection.appendChild(errorDiv);
        }
    })();
}

// Removed approvers tab functionality - keeping only Request details tab integration

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
    
    // Use reusable approvers section
    approversSection.innerHTML = createApproversSection('access-request-approvers', 'access-request-approvers-content');
    
    // Append section to panel
    accessRequestPanel.appendChild(approversSection);
    
    // Use reusable click handler
    createApproversClickHandler('access-request-approvers', 'access-request-approvers-content');
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