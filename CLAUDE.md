# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Chrome extension called "MyAccess Enhanced - GUID Lookup & Approvers" that enhances the Microsoft MyAccess portal experience. Converted from a streamlined Tampermonkey userscript, it provides comprehensive package information, GUID lookup capabilities, and detailed approver workflows with group member listings.

## Architecture

The extension follows Chrome Extension Manifest V3 architecture with a pure content-script approach:

### Core Files
- `manifest.json` - Extension configuration with permissions for Microsoft and Azure domains
- `src/content.js` - Complete functionality including GUID lookup, tab injection, and API calls
- `src/styles.css` - Comprehensive CSS styles for the enhanced UI components

### Main Features

#### 1. **GUID Lookup System**
- Uses `elm.iga.azure.com` API for comprehensive package searches
- Extracts detailed package metadata including creation dates, risk levels, and descriptions
- Provides exact and partial matching with match count reporting

#### 2. **Enhanced Tab Interface**
- Automatically injects "Approvers" tab into MyAccess package detail modals
- Displays comprehensive package information with risk level badges
- Shows GUID with one-click copy functionality
- Progress bars for long-running API operations

#### 3. **Simplified Approval Display**
- Shows unique approver groups from all policies
- Attempts to resolve group membership via Microsoft Graph API
- Displays member information with job titles and departments
- Shows requestor groups section

#### 4. **Token Management**
- Sophisticated token extraction from sessionStorage and localStorage
- Supports both `elm.iga.azure.com` and `graph.microsoft.com` tokens
- Falls back through multiple token storage patterns
- Includes MSAL token support

### API Integration

#### ELM IGA API (elm.iga.azure.com)
- `GET /api/v1/accessPackages` - Package search and metadata retrieval
- `GET /api/v1/accessPackageAssignmentPolicies` - Assignment policy retrieval
- Uses bearer token authentication from sessionStorage

#### Microsoft Graph API (graph.microsoft.com)
- `GET /v1.0/groups/{id}/members` - Group member enumeration
- Dual-token approach: Graph-specific tokens preferred, ELM tokens as fallback
- Graceful degradation when Graph API access is unavailable

### Authentication Flow
1. **Primary**: Extract `elm.iga.azure.com` tokens from `react-auth` session keys
2. **Graph API**: Look for `graph.microsoft.com` specific tokens
3. **MSAL Fallback**: Check MSAL tokens in sessionStorage and localStorage
4. **Validation**: Attempt to decode JWT tokens to verify audience

## Development

### Loading the Extension
1. Open Chrome and navigate to `chrome://extensions/`
2. Enable "Developer mode"
3. Click "Load unpacked" and select this directory
4. Navigate to `https://myaccess.microsoft.com/` to test

### Testing the Features
1. Navigate to a package detail page and look for "Approvers" tab
2. **Debug Commands**: Debug functions have been removed from production code

### Key Dependencies
- Chrome Extension APIs: storage, identity, scripting, activeTab, cookies
- Host permissions: myaccess.microsoft.com, myapplications.microsoft.com, graph.microsoft.com, elm.iga.azure.com
- CSS injection via content_scripts and web_accessible_resources
- DOM manipulation for tab injection and modal enhancement

## File Structure
```
/
├── manifest.json          # Extension configuration (no popup)
├── CLAUDE.md              # This documentation file
└── src/
    ├── content.js         # Complete functionality (streamlined version)
    └── styles.css         # Comprehensive UI styling
```

## Important Implementation Details

### Streamlined Architecture
- **Pure Content Script**: No popup or background scripts needed
- **Direct Injection**: Automatically enhances MyAccess pages on load
- **Simplified Display**: Shows unique approver and requestor groups clearly
- **Clean Organization**: Well-structured code with clear section comments

### Token Extraction Patterns
- Looks for sessionStorage keys containing `react-auth` and `session`
- Prioritizes `elm.iga.azure.com` tokens for package operations
- Uses `graph.microsoft.com` tokens for group member queries
- Includes MSAL token parsing with audience validation

### UI Components
- Custom styled progress bars for API operations
- Risk level badges extracted from package names (e.g., `[HIGH]`, `[LOW]`)
- Member cards with avatars generated from initials
- Clean, simplified approver group display

### Error Handling
- Graceful degradation when APIs are unavailable
- Clear user messaging for authentication issues
- Helpful workarounds provided (e.g., Azure CLI commands)
- Console debugging information for troubleshooting

### Performance Optimizations
- Batched API calls for group member retrieval
- Progress indicators for long-running operations
- Efficient DOM manipulation with minimal reflows
- CSS-only animations and transitions

## Key Differences from Complex Version
- **No Popup**: Pure content script extension with automatic injection
- **Simplified Display**: Focuses on unique groups rather than detailed policy breakdown
- **Streamlined Code**: Better organized with clear section headers
- **Cleaner UI**: More focused on essential information
- **Better Performance**: Reduced complexity and overhead

## Known Limitations
- Graph API access may require additional permissions for some tenants
- Package lookup requires exact or partial name matching
- Group member enumeration depends on Graph API token availability
- Some legacy MSAL token formats may not be supported

## Conversion Notes
This extension was converted from a streamlined Tampermonkey userscript (v3.0). The conversion maintained the simplified architecture while adapting to Chrome Extension security constraints. The focus is on clean, essential functionality without unnecessary complexity.