# Azure AD App Registration Setup Guide

This guide walks you through setting up an Azure Active Directory application for the Outlook Auto-Unsubscriber add-in.

## Prerequisites

- Azure AD tenant (comes with Microsoft 365)
- Global Administrator or Application Administrator role
- Access to [Azure Portal](https://portal.azure.com)

## 📌 Important: Personal vs Organizational Accounts

Your Azure AD app registration will work for **BOTH** types of Microsoft accounts:

### ✅ Office 365 / Microsoft 365 (Work/School Accounts)
- Full Microsoft Graph API access
- Enterprise features available
- Can be deployed organization-wide by IT admin
- Supports all scanning features

### ✅ Personal Microsoft Accounts (outlook.com, hotmail.com, live.com)
- Also fully supported!
- Users can install the add-in individually
- Same scanning functionality works
- Cannot be deployed organization-wide (no IT admin)

### How to Support Both:

When registering your app in Step 1, choose:
```
Supported account types:
→ "Accounts in any organizational directory and personal Microsoft accounts (Any Azure AD directory - Multitenant + Personal Microsoft accounts)"
```

**Note:** The code in this project already uses `authority: "common"` which supports both account types automatically.

### Limitations for Personal Accounts:
- Some advanced Microsoft Graph features may be limited
- Users must install add-in themselves (no centralized deployment)
- No organization-wide management

### For Enterprise Only:
If you only want to support your organization:
```
Supported account types:
→ "Accounts in this organizational directory only (Single tenant)"
```

---

## Step 1: Register the Application

1. **Navigate to Azure Portal**
   - Go to https://portal.azure.com
   - Sign in with your admin account

2. **Open Azure Active Directory**
   - Click on "Azure Active Directory" in the left menu
   - Or search for "Azure Active Directory" in the top search bar

3. **Create App Registration**
   - Click on "App registrations" in the left menu
   - Click "+ New registration" button

4. **Fill in Registration Details**
   ```
   Name: Outlook Auto-Unsubscriber

   Supported account types: (See "Personal vs Organizational Accounts" section above)

   RECOMMENDED (supports both personal and organizational accounts):
   → "Accounts in any organizational directory and personal Microsoft accounts"

   Other options:
   - For enterprise only: "Accounts in this organizational directory only"
   - For multi-tenant organizations only: "Accounts in any organizational directory"

   Redirect URI:
   Platform: Single-page application (SPA)
   URI: https://YOUR-GITHUB-USERNAME.github.io/Outlook-auto-unsubscriber/outlook-addin/src/taskpane/taskpane.html
   ```

   **Note:** Replace `YOUR-GITHUB-USERNAME` with your actual GitHub username!

5. **Click "Register"**

## Step 2: Configure API Permissions

1. **Add Microsoft Graph Permissions**
   - In your app registration, click "API permissions" in the left menu
   - Click "+ Add a permission"
   - Select "Microsoft Graph"
   - Select "Delegated permissions"

2. **Select Required Permissions**
   - Search for and select these permissions:
     - `User.Read` - Read user profile
     - `Mail.Read` - Read user mail
     - `Mail.ReadBasic` - Read basic mail properties

3. **Grant Admin Consent** (Optional but recommended for enterprise)
   - Click "Grant admin consent for [Your Organization]"
   - Click "Yes" to confirm
   - This prevents users from having to consent individually

## Step 3: Configure Authentication

1. **Navigate to Authentication**
   - Click "Authentication" in the left menu

2. **Add Platform (if not already added)**
   - Click "+ Add a platform"
   - Select "Single-page application"
   - Add redirect URI: `https://YOUR-GITHUB-USERNAME.github.io/Outlook-auto-unsubscriber/src/taskpane/taskpane.html`

3. **Configure Token Settings**
   - Under "Implicit grant and hybrid flows":
     - ✅ Check "Access tokens"
     - ✅ Check "ID tokens"

   - Under "Allow public client flows":
     - Select "No"

4. **Click "Save"**

## Step 4: Get Your Application (Client) ID

1. **Copy the Client ID**
   - Go to "Overview" in the left menu
   - Copy the "Application (client) ID" value
   - It looks like: `12345678-1234-1234-1234-123456789abc`

2. **Update Your Code**
   - Open `src/auth/auth.js`
   - Replace `YOUR_CLIENT_ID_HERE` with your actual Client ID:
   ```javascript
   const msalConfig = {
       auth: {
           clientId: "12345678-1234-1234-1234-123456789abc", // Your Client ID
           authority: "https://login.microsoftonline.com/common",
           redirectUri: window.location.origin + "/outlook-addin/src/taskpane/taskpane.html"
       },
       // ...
   };
   ```

## Step 5: Configure Manifest

1. **Update manifest.json**
   - Open `manifest.json`
   - Replace all instances of `YOUR-GITHUB-USERNAME` with your actual GitHub username
   - Update the GUID `id` field with a new unique GUID

   Generate a new GUID at: https://www.guidgenerator.com/

## Step 6: Test Authentication

1. **Deploy to GitHub Pages** (see DEPLOYMENT.md)

2. **Load the add-in in Outlook**
   - Follow the testing instructions in DEPLOYMENT.md

3. **Test Sign In**
   - Click "Sign In with Microsoft 365"
   - You should be redirected to Microsoft login
   - Sign in with your Microsoft 365 account
   - Grant permissions when prompted
   - You should be redirected back to the add-in

## Common Issues

### Issue: "AADSTS50011: Reply URL does not match"

**Solution:**
- Ensure the redirect URI in Azure AD matches exactly with your GitHub Pages URL
- Include the full path: `https://username.github.io/Outlook-auto-unsubscriber/src/taskpane/taskpane.html`
- Check for typos and case sensitivity

### Issue: "Consent required"

**Solution:**
- Admin needs to grant consent in Azure AD
- Or users need to consent individually on first sign-in

### Issue: "AADSTS65001: User or administrator has not consented"

**Solution:**
- Go to API permissions in Azure AD
- Click "Grant admin consent for [Your Organization]"

### Issue: Token acquisition fails silently

**Solution:**
- Check browser console for errors
- Verify API permissions are correctly configured
- Ensure tokens are not expired (clear browser cache)

## Security Best Practices

1. **Use specific permissions** - Only request the minimum permissions needed
2. **Enable Conditional Access** - Require MFA for sensitive operations
3. **Monitor app usage** - Use Azure AD sign-in logs to monitor authentication
4. **Rotate secrets** - If using client secrets, rotate them regularly (not needed for SPA)
5. **Restrict tenant access** - For enterprise, use "Single tenant" mode

## For AppSource Submission

If you plan to submit to AppSource:

1. Use "Multitenant" account type
2. Add privacy policy URL
3. Add terms of service URL
4. Complete app verification process
5. Follow Microsoft's validation guidelines

## Additional Resources

- [Azure AD App Registration Documentation](https://docs.microsoft.com/azure/active-directory/develop/quickstart-register-app)
- [Microsoft Graph Permissions Reference](https://docs.microsoft.com/graph/permissions-reference)
- [MSAL.js Documentation](https://docs.microsoft.com/azure/active-directory/develop/msal-js-initializing-client-applications)

## Support

If you encounter issues:
1. Check Azure AD sign-in logs for error details
2. Review browser console for JavaScript errors
3. Verify all URLs and GUIDs are correct
4. Test with a different account
