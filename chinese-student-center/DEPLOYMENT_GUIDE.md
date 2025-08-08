# 🚀 Chinese Student Center - Deployment & Sharing Guide

## 🔐 Permission Issues Fix

If users see the error: "You do not have permission to call SpreadsheetApp.getActiveSpreadsheet", follow these steps:

## ✅ Step 1: Set Up Proper Sharing

### For Spreadsheet Owner:

1. **Share the Spreadsheet**
   - Open your Google Spreadsheet
   - Click "Share" button (top right)
   - Add user emails with "Editor" permissions (not "Viewer")
   - Make sure "Notify people" is checked

2. **Enable Apps Script Sharing**
   - Go to Extensions → Apps Script
   - Click "Deploy" → "New Deployment"
   - Choose type: "Web app"
   - Execute as: "Me (your email)"
   - Who has access: "Anyone with Google account" or specific users
   - Click "Deploy"

## ✅ Step 2: Apps Script Configuration

### Required OAuth Scopes:
The script automatically requests these permissions:
- `https://www.googleapis.com/auth/spreadsheets.currentonly`
- `https://www.googleapis.com/auth/script.external_request`

### Authorization Process:
1. First-time users will see an authorization screen
2. Click "Review permissions"
3. Select your Google account
4. Click "Allow" to grant necessary permissions

## ✅ Step 3: Deployment Options

### Option A: Spreadsheet-Bound Script (Recommended)
- Script runs within the spreadsheet
- Users need "Editor" access to the spreadsheet
- Most secure option

### Option B: Web App Deployment
1. In Apps Script: Deploy → New Deployment
2. Type: Web app
3. Execute as: "Me" (script owner's account)
4. Who has access: "Anyone with Google account"
5. Copy the web app URL and share with users

## ✅ Step 4: User Instructions

### For New Users:
1. **Get Access**: Request "Editor" access to the spreadsheet
2. **First Time Setup**:
   - Open the spreadsheet
   - Go to Extensions → Apps Script
   - Click any function to trigger authorization
   - Follow the authorization prompts
3. **Use the System**: Access features through the "🏛️ X.Lab Center" menu

### For Recurring Users:
- Simply use the menu: Extensions → Apps Script → Run function
- Or use the custom menu: "🏛️ X.Lab Center"

## 🔧 Troubleshooting

### Common Issues:

1. **"Permission denied"**
   - Solution: Spreadsheet owner must share with "Editor" access
   - Check: User is signed into correct Google account

2. **"Script function not found"**
   - Solution: Make sure the latest script version is deployed
   - Try: Refresh the spreadsheet and try again

3. **"Authorization required"**
   - Solution: User needs to complete OAuth authorization
   - Go to: Extensions → Apps Script → Authorize

4. **"Spreadsheet not found"**
   - Solution: Ensure user has direct access to the spreadsheet
   - Check: Spreadsheet sharing settings

## 🔒 Security Best Practices

1. **Limited Scope**: Script only accesses the current spreadsheet
2. **User Permissions**: Only editors can modify data
3. **Audit Trail**: All actions are logged with user email
4. **Safe Execution**: Built-in error handling prevents crashes

## 📞 Support

If users continue experiencing issues:
1. Check Google Account permissions
2. Verify spreadsheet sharing settings
3. Re-deploy the Apps Script if needed
4. Contact system administrator

## 🎯 Quick Setup Checklist

- [ ] Spreadsheet shared with "Editor" access
- [ ] Apps Script deployed as web app (if needed)
- [ ] Users completed OAuth authorization
- [ ] Test all major functions work
- [ ] Documentation shared with users

---

💡 **Pro Tip**: The system now includes built-in permission checking and user-friendly error messages to guide users through any access issues!
