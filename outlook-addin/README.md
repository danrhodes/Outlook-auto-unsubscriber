# Outlook Auto-Unsubscriber Add-in

A modern Outlook add-in for New Outlook that helps you find and manage email unsubscribe links directly within your Outlook client.

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Platform](https://img.shields.io/badge/platform-Outlook-orange)
![License](https://img.shields.io/badge/license-MIT-green)

## ✨ Features

- **🔍 Smart Scanning** - Automatically detects unsubscribe links using multiple methods
- **📊 Priority Detection** - Identifies high-volume senders (10+ emails)
- **📈 Statistics Dashboard** - See email counts, domains, and priorities at a glance
- **🎯 Dual Scan Modes** - Scan all emails OR just selected email
- **🔐 Secure Authentication** - Microsoft 365 OAuth integration
- **💅 Modern UI** - Fluent UI design aligned with New Outlook
- **⚡ Fast & Efficient** - Processes hundreds of emails in seconds
- **📱 Cross-Platform** - Works on Outlook Web, Desktop, and Mobile

## 🎯 Works With

- ✅ **New Outlook** (Windows, Mac)
- ✅ **Outlook on the Web** (Office 365)
- ✅ **Outlook Desktop** (Microsoft 365)
- ✅ **Outlook Mobile** (iOS, Android)
- ✅ **Personal Microsoft Accounts** (outlook.com, hotmail.com, live.com)
- ✅ **Work/School Accounts** (Office 365, Microsoft 365)

## 📋 Prerequisites

- Microsoft 365 account OR personal Microsoft account (Outlook.com)
- Azure AD app registration (free)
- GitHub account (for hosting)

## 🚀 Quick Start

### 1. Clone the Repository

```bash
git clone https://github.com/YOUR-USERNAME/Outlook-auto-unsubscriber.git
cd Outlook-auto-unsubscriber/outlook-addin
```

### 2. Set Up Azure AD App

Follow the detailed guide: **[AZURE_AD_SETUP.md](AZURE_AD_SETUP.md)**

Quick steps:
1. Register app in Azure Portal
2. Add API permissions (Mail.Read, Mail.ReadBasic, User.Read)
3. Configure authentication
4. Copy Client ID
5. Update `src/auth/auth.js` with your Client ID

### 3. Configure for Your Environment

Edit `manifest.json`:
```json
{
  "id": "YOUR-UNIQUE-GUID",
  "iconUrl": "https://YOUR-GITHUB-USERNAME.github.io/...",
  // ... update all URLs
}
```

Edit `src/auth/auth.js`:
```javascript
const msalConfig = {
    auth: {
        clientId: "YOUR-CLIENT-ID-HERE",
        // ...
    }
};
```

### 4. Deploy to GitHub Pages

```bash
git add .
git commit -m "Configure add-in"
git push origin main
```

Enable GitHub Pages:
- Go to Settings → Pages
- Source: main branch
- Save

### 5. Test Your Add-in

#### Sideload in Outlook:
1. Open Outlook on the Web
2. Settings → View all Outlook settings
3. Mail → Customize actions → Add-ins
4. "Add from file" → select `manifest.json`

See **[DEPLOYMENT.md](DEPLOYMENT.md)** for detailed deployment options.

## 📖 Usage

### Scanning All Emails

1. Open the add-in from Outlook ribbon
2. Sign in with your Microsoft 365 account
3. Select options:
   - Days to scan back (1-365)
   - Folder to scan (Inbox, Junk, etc.)
4. Click "Scan All Emails"
5. View results in the interactive table

### Scanning Selected Email

1. Select an email in Outlook
2. Open the add-in
3. Click "Scan Selected Email"
4. View unsubscribe link if found

### Using the Report

- **Search** - Filter by domain or subject
- **Sort** - By priority, domain, or count
- **Filter** - Show only high/medium/low priority
- **Click to Unsubscribe** - Opens link in new tab
- **Track Clicks** - Automatically marks clicked links

## 🏗️ Architecture

```
outlook-addin/
├── manifest.json           # Add-in manifest
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html  # Main UI
│   │   ├── taskpane.css   # Fluent UI styles
│   │   └── taskpane.js    # Main controller
│   ├── auth/
│   │   └── auth.js        # MSAL authentication
│   └── services/
│       ├── graph.js       # Microsoft Graph API
│       └── scanner.js     # Email scanning logic
├── assets/
│   └── icon.svg          # App icon (convert to PNG)
├── AZURE_AD_SETUP.md     # Azure AD setup guide
├── DEPLOYMENT.md         # Deployment guide
└── README.md            # This file
```

## 🔧 Technologies

- **Office.js** - Outlook Add-in API
- **MSAL.js 2.x** - Microsoft Authentication Library
- **Microsoft Graph API** - Email access
- **Fluent UI** - Microsoft design system
- **Vanilla JavaScript** - No frameworks, lightweight
- **HTML5 + CSS3** - Modern web standards

## 📦 Deployment Options

### 1. GitHub Pages (Free)
- Perfect for testing and small teams
- See [DEPLOYMENT.md](DEPLOYMENT.md)

### 2. Azure Static Web Apps
- Production-ready with custom domain
- Automatic SSL and CI/CD
- See [DEPLOYMENT.md](DEPLOYMENT.md)

### 3. Microsoft AppSource
- Public distribution to all Outlook users
- Requires validation
- See [DEPLOYMENT.md](DEPLOYMENT.md)

### 4. Enterprise Centralized Deployment
- IT admin deploys to entire organization
- No user installation needed
- See [DEPLOYMENT.md](DEPLOYMENT.md)

## 🔒 Security & Privacy

- ✅ **Read-only access** - Cannot modify or delete emails
- ✅ **OAuth 2.0** - Secure Microsoft authentication
- ✅ **No data storage** - All processing happens client-side
- ✅ **No third-party services** - Direct Graph API calls only
- ✅ **Open source** - Fully auditable code
- ✅ **Token security** - MSAL handles secure token storage
- ✅ **HTTPS only** - All connections encrypted

## 🎨 Icons

The add-in needs icons in these sizes:
- 16x16, 32x32, 64x64, 80x80, 128x128 pixels

An SVG template is provided in `assets/icon.svg`.

Convert to PNG using ImageMagick:
```bash
convert icon.svg -resize 16x16 icon-16.png
convert icon.svg -resize 32x32 icon-32.png
convert icon.svg -resize 64x64 icon-64.png
convert icon.svg -resize 80x80 icon-80.png
convert icon.svg -resize 128x128 icon-128.png
```

Or use online tools:
- https://cloudconvert.com/svg-to-png
- https://svgtopng.com/

## 🐛 Troubleshooting

### Add-in won't load
- Verify all URLs in manifest.json are correct and accessible
- Check Azure AD redirect URIs match
- Clear browser cache

### Authentication fails
- Verify Client ID in auth.js
- Check API permissions in Azure AD
- Ensure pop-ups are not blocked

### No results found
- Try increasing days to scan
- Select "All Mail" folder
- Verify email has unsubscribe links

See [DEPLOYMENT.md](DEPLOYMENT.md) for more troubleshooting.

## 📊 API Permissions

The add-in requests these Microsoft Graph permissions:

| Permission | Type | Reason |
|------------|------|--------|
| User.Read | Delegated | Read user profile |
| Mail.Read | Delegated | Read mailbox emails |
| Mail.ReadBasic | Delegated | Read basic email properties |

## 🆚 Comparison: PowerShell vs Add-in

| Feature | PowerShell Script | Outlook Add-in |
|---------|------------------|----------------|
| **Platform** | Windows only | All platforms |
| **Integration** | External tool | Inside Outlook |
| **Installation** | Manual execution | One-click install |
| **Updates** | Manual download | Automatic |
| **User Experience** | Command line | Modern web UI |
| **Authentication** | Popup login | Seamless OAuth |
| **Deployment** | Per-user script | Enterprise-wide |
| **Mobile Support** | ❌ | ✅ |

## 🤝 Contributing

Contributions are welcome!

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📝 License

MIT License - See LICENSE file for details

## 🙏 Credits

Created by Dan Rhodes

Built with:
- Microsoft Office Add-ins platform
- Microsoft Graph API
- MSAL.js
- Fluent UI

Original PowerShell version: [../Find_Unsubscribe_Links-Enhanced.ps1](../Find_Unsubscribe_Links-Enhanced.ps1)

## 📞 Support

- **Documentation**: See [AZURE_AD_SETUP.md](AZURE_AD_SETUP.md) and [DEPLOYMENT.md](DEPLOYMENT.md)
- **Issues**: [GitHub Issues](https://github.com/YOUR-USERNAME/Outlook-auto-unsubscriber/issues)
- **Microsoft Docs**: [Outlook Add-ins](https://docs.microsoft.com/office/dev/add-ins/outlook/)

## 🗺️ Roadmap

- [ ] One-click unsubscribe support (RFC 8058)
- [ ] Batch unsubscribe operations
- [ ] Export reports to CSV
- [ ] Email category filtering
- [ ] Dark mode support
- [ ] Localization (multiple languages)
- [ ] Advanced analytics dashboard

---

**Made with ❤️ to help you reclaim your inbox!**
