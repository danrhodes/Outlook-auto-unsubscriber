# 🚀 Quick Start Guide

Get your Outlook Auto-Unsubscriber add-in running in 15 minutes!

## ⚡ Prerequisites

- [ ] GitHub account
- [ ] Microsoft account (personal or Office 365)
- [ ] Access to Azure Portal (free)

## 📝 Step-by-Step Setup

### Step 1: Get Azure Client ID (5 minutes)

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **+ New registration**

   Fill in:
   ```
   Name: Outlook Auto-Unsubscriber

   Supported account types:
   → "Accounts in any organizational directory and personal Microsoft accounts"

   Redirect URI:
   Platform: Single-page application (SPA)
   URI: https://YOUR-GITHUB-USERNAME.github.io/Outlook-auto-unsubscriber/outlook-addin/src/taskpane/taskpane.html
   ```

4. Click **Register**
5. Copy the **Application (client) ID** (looks like `12345678-1234-...`)

6. Add API permissions:
   - Click **API permissions** → **+ Add a permission**
   - Select **Microsoft Graph** → **Delegated permissions**
   - Add: `User.Read`, `Mail.Read`, `Mail.ReadBasic`
   - Click **Add permissions**

---

### Step 2: Configure Your Code (2 minutes)

1. **Edit `src/auth/auth.js`**

   Replace line 8:
   ```javascript
   clientId: "YOUR_CLIENT_ID_HERE",  // ← Paste your Client ID here
   ```

2. **Edit `manifest.json`**

   Replace:
   - All instances of `YOUR-GITHUB-USERNAME` with your actual GitHub username
   - Line 5: `"id"` with a new GUID from https://www.guidgenerator.com/

---

### Step 3: Deploy to GitHub Pages (3 minutes)

1. **Commit and push:**
   ```bash
   cd outlook-addin
   git add .
   git commit -m "Configure Outlook add-in for deployment"
   git push origin main
   ```

2. **Enable GitHub Pages:**
   - Go to your repo on GitHub.com
   - Click **Settings** → **Pages**
   - Source: **main branch**
   - Click **Save**
   - Wait 2 minutes for deployment

3. **Verify it's live:**
   - Visit: `https://YOUR-USERNAME.github.io/Outlook-auto-unsubscriber/outlook-addin/manifest.json`
   - Should see JSON content (not 404)

---

### Step 4: Update Azure Redirect URI (1 minute)

1. Back in Azure Portal
2. Your app registration → **Authentication**
3. Verify redirect URI matches:
   ```
   https://YOUR-GITHUB-USERNAME.github.io/Outlook-auto-unsubscriber/outlook-addin/src/taskpane/taskpane.html
   ```
4. Save if you made changes

---

### Step 5: Install in Outlook (4 minutes)

#### Option A: Outlook on the Web
1. Go to https://outlook.office.com
2. Click ⚙️ **Settings** → **View all Outlook settings**
3. Go to **Mail** → **Customize actions** → **Add-ins**
4. Click **+ Add from file**
5. Browse to your `manifest.json` file
6. Click **Install**

#### Option B: Desktop Outlook
1. Open Outlook Desktop
2. Click **Get Add-ins** (ribbon)
3. **My add-ins** (left menu)
4. **Add a custom add-in** → **Add from file**
5. Select your `manifest.json`

---

### Step 6: Test It! (1 minute)

1. Open any email in Outlook
2. Look for **"Scan for Unsubscribe Links"** button in the ribbon
3. Click it to open the task pane
4. Click **"Sign In with Microsoft 365"**
5. Sign in and grant permissions
6. Click **"Scan All Emails"**
7. 🎉 See your results!

---

## ✅ You're Done!

Your add-in is now:
- ✅ Deployed on GitHub Pages
- ✅ Authenticated with Azure AD
- ✅ Installed in Outlook
- ✅ Ready to use!

---

## 🐛 Troubleshooting

### "Reply URL mismatch" error
→ Make sure redirect URI in Azure AD **exactly** matches your GitHub Pages URL

### Can't load task pane
→ Check browser console (F12) for errors
→ Verify GitHub Pages is deployed and accessible

### Authentication fails
→ Verify Client ID in `auth.js` is correct
→ Check API permissions are added in Azure AD

### Add-in doesn't appear
→ Wait 5-10 minutes after sideloading
→ Try refreshing Outlook

---

## 📚 Next Steps

- [ ] Read [DEPLOYMENT.md](DEPLOYMENT.md) for production deployment
- [ ] Learn about enterprise deployment options
- [ ] Customize the UI to match your brand
- [ ] Create proper icons (see [assets/README.md](assets/README.md))
- [ ] Consider submitting to AppSource

---

## 🆘 Need Help?

- **Detailed setup:** [AZURE_AD_SETUP.md](AZURE_AD_SETUP.md)
- **Deployment options:** [DEPLOYMENT.md](DEPLOYMENT.md)
- **GitHub Issues:** Report problems
- **Microsoft Docs:** [Outlook Add-ins](https://docs.microsoft.com/office/dev/add-ins/outlook/)

---

**Happy Unsubscribing! 📧✨**
