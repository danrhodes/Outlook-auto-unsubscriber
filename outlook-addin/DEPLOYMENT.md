# Deployment Guide - Outlook Auto-Unsubscriber Add-in

This guide covers multiple deployment options for the Outlook Auto-Unsubscriber add-in.

## Deployment Options

1. **GitHub Pages (Free)** - Recommended for testing and small deployments
2. **Azure Static Web Apps** - For production with custom domain
3. **AppSource** - For public distribution to all Outlook users
4. **Organization Deployment** - For enterprise-wide rollout

---

## Option 1: GitHub Pages Deployment (Recommended)

### Step 1: Prepare Your Repository

1. **Fork or Clone this repository**
   ```bash
   git clone https://github.com/YOUR-USERNAME/Outlook-auto-unsubscriber.git
   cd Outlook-auto-unsubscriber
   ```

2. **Update Configuration Files**

   Edit `manifest.json` - Replace placeholders:
   ```json
   "id": "YOUR-NEW-GUID-HERE",
   "iconUrl": "https://YOUR-GITHUB-USERNAME.github.io/Outlook-auto-unsubscriber/assets/icon-64.png",
   ```

   Edit `src/auth/auth.js` - Add your Azure AD Client ID:
   ```javascript
   clientId: "YOUR_CLIENT_ID_HERE",
   ```

### Step 2: Enable GitHub Pages

1. **Push changes to GitHub**
   ```bash
   git add .
   git commit -m "Configure add-in for GitHub Pages"
   git push origin main
   ```

2. **Enable GitHub Pages**
   - Go to repository Settings
   - Scroll to "Pages" section
   - Source: Deploy from a branch
   - Branch: `main` / `root` (or wherever your files are)
   - Click "Save"

3. **Wait for deployment** (usually 1-2 minutes)
   - Your site will be available at: `https://YOUR-USERNAME.github.io/Outlook-auto-unsubscriber/`

### Step 3: Update Azure AD Redirect URIs

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to your App Registration
3. Click "Authentication"
4. Update redirect URI to match your GitHub Pages URL:
   ```
   https://YOUR-USERNAME.github.io/Outlook-auto-unsubscriber/outlook-addin/src/taskpane/taskpane.html
   ```

### Step 4: Test Your Add-in

See "Testing Your Add-in" section below.

---

## Option 2: Azure Static Web Apps

Azure Static Web Apps provides:
- Free tier available
- Custom domains
- Automatic SSL
- CI/CD from GitHub
- Better performance

### Setup Steps:

1. **Create Azure Static Web App**
   ```bash
   # Install Azure CLI
   curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash

   # Login
   az login

   # Create resource group
   az group create --name outlook-addin-rg --location eastus

   # Create static web app
   az staticwebapp create \
       --name outlook-auto-unsubscriber \
       --resource-group outlook-addin-rg \
       --source https://github.com/YOUR-USERNAME/Outlook-auto-unsubscriber \
       --location eastus \
       --branch main \
       --app-location "/outlook-addin" \
       --login-with-github
   ```

2. **Get your Azure Static Web App URL**
   ```bash
   az staticwebapp show \
       --name outlook-auto-unsubscriber \
       --resource-group outlook-addin-rg \
       --query "defaultHostname" -o tsv
   ```

3. **Update manifest.json and Azure AD** with new URL

---

## Option 3: AppSource Distribution

To list your add-in in the Microsoft AppSource (Office Add-ins Store):

### Requirements:

- ✅ Azure AD app registration with multi-tenant + personal accounts support
- ✅ Privacy policy URL
- ✅ Terms of service URL
- ✅ Support URL
- ✅ App validation requirements met
- ✅ Secure hosting (HTTPS)

### Steps:

1. **Prepare for Submission**
   - Complete all metadata in manifest
   - Test thoroughly on all platforms
   - Create promotional materials (screenshots, video)

2. **Submit to Partner Center**
   - Go to [Partner Center](https://partner.microsoft.com/)
   - Create a new Office Add-in offer
   - Upload your manifest
   - Complete all required information
   - Submit for validation

3. **Wait for Approval** (typically 1-2 weeks)

4. **Once Approved**
   - Users can install directly from Outlook
   - No sideloading needed
   - Automatic updates

### Benefits:
- ✅ Reach millions of users
- ✅ Trusted installation source
- ✅ Automatic updates
- ✅ Microsoft validation badge

---

## Option 4: Enterprise Organization-Wide Deployment

For IT admins to deploy to all users in an organization.

### Prerequisites:
- Microsoft 365 Admin Center access
- Global Administrator or Exchange Administrator role

### Centralized Deployment Steps:

1. **Sign in to Microsoft 365 Admin Center**
   - Go to https://admin.microsoft.com
   - Navigate to Settings > Integrated apps

2. **Upload Custom App**
   - Click "Upload custom apps"
   - Click "Upload manifest file (.xml or .json)"
   - Select your `manifest.json` file
   - Click "Upload"

3. **Configure Deployment**
   - Choose deployment type:
     - **Everyone**: All users automatically get the add-in
     - **Specific users/groups**: Deploy to selected groups
     - **Just me**: Test deployment

   - Set deployment options:
     - ✅ "Pin to ribbon" (recommended)
     - ✅ "Make add-in available immediately"

4. **Deploy**
   - Click "Deploy"
   - Users will see the add-in in Outlook within 24 hours

### Benefits:
- ✅ No user installation required
- ✅ Central management
- ✅ Easy updates
- ✅ Usage analytics
- ✅ Can enforce usage policies

---

## Testing Your Add-in

### Method 1: Sideloading (Development/Testing)

#### For Outlook on the Web:
1. Go to https://outlook.office.com
2. Click Settings (gear icon) → View all Outlook settings
3. Go to Mail → Customize actions → Add-ins
4. Click "Add from file"
5. Browse and select your `manifest.json`
6. Click "Install"

#### For Outlook Desktop (Windows):
1. Open Outlook
2. Click "Get Add-ins" from the ribbon
3. Click "My add-ins" on the left
4. Under "Custom add-ins", click "Add a custom add-in"
5. Select "Add from file"
6. Browse and select your `manifest.json`

#### For New Outlook (Windows):
1. Open New Outlook
2. Click Apps icon in the sidebar
3. Click "Get Add-ins"
4. Follow same steps as Outlook Desktop

### Method 2: URL Sideloading

1. In Outlook on the Web:
2. Go to Settings → View all Outlook settings
3. Mail → Customize actions → Add-ins
4. Click "Add from URL"
5. Enter: `https://YOUR-USERNAME.github.io/Outlook-auto-unsubscriber/outlook-addin/manifest.json`

---

## Verification Checklist

Before deploying to users:

- [ ] Azure AD app is configured correctly
- [ ] All URLs in manifest.json are updated
- [ ] Client ID in auth.js is set
- [ ] Manifest GUID is unique
- [ ] Icons are accessible (create placeholder icons if needed)
- [ ] Redirect URIs match in Azure AD
- [ ] Tested sign-in flow
- [ ] Tested scan all emails
- [ ] Tested scan selected email
- [ ] Tested on Outlook Web
- [ ] Tested on Outlook Desktop
- [ ] Tested on New Outlook

---

## Updating Your Add-in

### For GitHub Pages:
1. Make changes to your code
2. Commit and push to GitHub
3. GitHub Pages auto-deploys
4. Users get updates on next load (may need cache clear)

### For AppSource:
1. Update version in manifest.json
2. Submit update to Partner Center
3. Wait for approval
4. Users auto-updated after approval

### For Enterprise Deployment:
1. Update manifest version
2. Re-upload in Admin Center
3. Update propagates to users within 24 hours

---

## Troubleshooting

### Add-in doesn't appear
- Wait 5-10 minutes after sideloading
- Clear Outlook cache
- Try different Outlook client (web vs desktop)
- Check manifest.json for errors

### Authentication fails
- Verify Azure AD redirect URIs
- Check browser console for errors
- Ensure pop-ups are not blocked
- Verify API permissions are granted

### Can't load task pane
- Check all URLs are HTTPS
- Verify GitHub Pages is deployed
- Check browser console for CORS errors
- Ensure Office.js is loading

### Icons not showing
- Verify icon URLs are accessible
- Icons must be HTTPS
- Supported formats: PNG, JPG, SVG
- Recommended sizes: 16x16, 32x32, 64x64, 128x128

---

## Performance Optimization

1. **Minify JavaScript**
   ```bash
   npm install -g terser
   terser src/taskpane/taskpane.js -o src/taskpane/taskpane.min.js
   ```

2. **Optimize Images**
   - Use compressed PNGs
   - Consider SVG for icons
   - Maximum 100KB per image

3. **Enable Caching**
   - Add cache headers on hosting
   - Use service workers for offline support

---

## Security Considerations

1. **HTTPS Only** - Never deploy over HTTP
2. **Content Security Policy** - Add CSP headers
3. **Input Validation** - Always validate user input
4. **Token Storage** - MSAL handles secure token storage
5. **Regular Updates** - Keep dependencies updated

---

## Support

- GitHub Issues: Report bugs and feature requests
- Documentation: Check AZURE_AD_SETUP.md for auth issues
- Microsoft Docs: [Outlook Add-ins Documentation](https://docs.microsoft.com/office/dev/add-ins/outlook/)

---

## Next Steps

1. ✅ Complete Azure AD setup (see AZURE_AD_SETUP.md)
2. ✅ Deploy to GitHub Pages
3. ✅ Test thoroughly
4. ✅ Deploy to organization or submit to AppSource
5. ✅ Monitor usage and gather feedback
