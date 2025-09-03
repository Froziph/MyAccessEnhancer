# Chrome Web Store Submission Guide

## Files Created for Submission

✅ **manifest.json** - Updated for store requirements
✅ **store-listing.md** - Complete store listing content
✅ **privacy-policy.md** - Privacy policy content
✅ **create-icons.html** - Icon generator tool
✅ **icons/** - Directory for icons (you need to generate them)

## Pre-Submission Steps

### 1. Generate Icons
1. Open `create-icons.html` in your browser
2. Click "Generate Icons" 
3. Right-click each icon and "Save image as..."
4. Save in the `icons/` folder with these exact names:
   - `icon-16.png`
   - `icon-48.png` 
   - `icon-128.png`

### 2. Publish Privacy Policy
You need a publicly accessible privacy policy URL:

**Option A: GitHub Pages**
1. Go to your GitHub repo settings
2. Enable GitHub Pages
3. Your privacy policy will be at: `https://froziph.github.io/MyAccessEnhancer/privacy-policy.html`

**Option B: Copy privacy-policy.md content to your website**

### 3. Update manifest.json
If you used Option B above, update the homepage_url in manifest.json to your privacy policy URL.

### 4. Test Extension Locally
1. Load unpacked extension in Chrome
2. Test on MyAccess to ensure everything works
3. Check for any console errors

## Chrome Web Store Submission Steps

### 1. Create Extension Package
1. Create a ZIP file containing:
   ```
   manifest.json
   src/content.js
   src/styles.css
   icons/icon-16.png
   icons/icon-48.png
   icons/icon-128.png
   ```
2. **DO NOT** include: 
   - `.git/` folder
   - `node_modules/`
   - `*.md` files
   - `create-icons.html`

### 2. Upload to Chrome Web Store
1. Go to [Chrome Web Store Developer Dashboard](https://chrome.google.com/webstore/devconsole)
2. Click "Add new item"
3. Upload your ZIP file
4. Fill out the store listing:

#### Store Listing Details:
- **Name**: MyAccess Enhanced - GUID Lookup & Approvers
- **Summary**: Enhance Microsoft MyAccess with package GUID lookup, detailed approver information, and group member visibility.
- **Description**: Copy from `store-listing.md` (the detailed description section)
- **Category**: Productivity
- **Language**: English
- **Website**: https://github.com/Froziph/MyAccessEnhancer
- **Support URL**: https://github.com/Froziph/MyAccessEnhancer/issues
- **Privacy Policy**: Your published privacy policy URL

#### Required Images:
You'll need to create these screenshots/images:

**Screenshots (1280x800 or 640x400):**
- Screenshot 1: MyAccess page showing the Approvers tab
- Screenshot 2: Approvers tab open showing group information
- Screenshot 3: Group members display

**Small Tile (440x280):**
- Create a promotional tile with your extension logo and name

**Marquee (1400x560) - Optional but recommended:**
- Large promotional banner for featured placement

### 3. Review Settings
- **Visibility**: Public (or unlisted if you prefer)
- **Pricing**: Free
- **Regions**: Worldwide (or specific regions if needed)

### 4. Review & Submit
1. Review all information carefully
2. Click "Submit for Review"
3. Pay the $5 developer fee (one-time)

## Expected Review Process

- **Initial Review**: 1-3 business days
- **Possible Issues**: 
  - Privacy policy concerns (common for enterprise extensions)
  - Permission justification requests
  - Icon/screenshot quality

## Post-Approval Steps

1. **Test the published extension**
2. **Monitor reviews and feedback**
3. **Update as needed** (updates typically review faster)

## Tips for Approval

✅ **Strong Points:**
- Clear, legitimate business purpose
- No data collection
- Uses existing Microsoft authentication
- Professional presentation
- Good documentation

⚠️ **Potential Review Issues:**
- Host permissions for Microsoft domains (be ready to justify)
- Enterprise-focused extensions sometimes get extra scrutiny
- Make sure privacy policy is accessible

## Support During Review

If Google requests clarification:
- Emphasize the legitimate business use case
- Highlight that it enhances existing Microsoft tools
- Reference that it uses existing authentication
- Point to clear privacy policy showing no data collection

Good luck with your submission! The extension looks professional and serves a clear business purpose, which should help with approval.