# 📧 Outlook Auto-Unsubscriber

A powerful tool to help you find and unsubscribe from unwanted email newsletters and mailing lists in your Microsoft 365/Outlook account.

## 🌟 Features

- **Automatic Detection** - Scans your emails and finds unsubscribe links
- **Priority System** - Identifies high-volume senders (🔴 High, 🟡 Medium, 🟢 Low)
- **Email Count Tracking** - Shows how many emails you received from each sender
- **Smart Filtering** - Search and filter by domain, priority, or subject
- **Click Tracking** - Remembers which links you've already clicked
- **One-Click Unsubscribe** - Automatically unsubscribes from supported emails
- **Beautiful Report** - Generates an interactive HTML report with all unsubscribe links
- **Statistics Dashboard** - See total emails, domains, and storage insights

## 📋 Prerequisites

Before you start, make sure you have:

1. **Windows 10 or 11** (with PowerShell pre-installed)
2. **Microsoft 365 Account** (Outlook, Office 365, or Exchange Online)
3. **Internet Connection** (to connect to Microsoft Graph API)

That's it! No programming knowledge required.

## 🚀 Quick Start Guide

### Step 1: Download the Script

1. Download all files from this repository
2. Save them to a folder on your computer (e.g., `C:\UnsubscribeTool\`)

### Step 2: Open PowerShell

**Option A - Right-click Method:**
1. Navigate to the folder containing the scripts
2. Hold **Shift** and **right-click** in the empty space
3. Select **"Open PowerShell window here"** or **"Open in Terminal"**

**Option B - Start Menu Method:**
1. Press **Windows Key**
2. Type `PowerShell`
3. Right-click **Windows PowerShell**
4. Select **"Run as Administrator"** (recommended)
5. Navigate to your folder: `cd C:\UnsubscribeTool\`

### Step 3: Enable Script Execution (First Time Only)

If this is your first time running PowerShell scripts, you need to allow them:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

When prompted, type **Y** and press **Enter**.

> **What does this do?** This allows PowerShell to run scripts you've written or downloaded. It's a one-time setup.

### Step 4: Run the Script

Type this command and press Enter:

```powershell
.\Find_Unsubscribe_Links-Enhanced.ps1
```

### Step 5: Follow the Prompts

The script will ask you several questions:

1. **Email Address**: Enter your Microsoft 365 email
   ```
   Enter your email address: john.smith@company.com
   ```

2. **Sign In**: A browser window will open - sign in with your Microsoft account

3. **Grant Permissions**: Click **Accept** to allow the app to read your emails

4. **Number of Days**: How far back to scan (e.g., `30` for last 30 days)
   ```
   Enter the number of days to check back on emails: 30
   ```

5. **Folder Selection**: Which folder to scan
   ```
   1. Inbox (default)
   2. Sent Items
   3. Junk Email
   4. All Mail (entire mailbox)
   Select folder option (1-4, default: 1): 1
   ```

6. **One-Click Unsubscribe** (Enhanced version only): Enable automatic unsubscription?
   ```
   Enable automatic one-click unsubscribe? (y/n): n
   ```
   > ⚠️ **Recommendation**: Type `n` (No) for your first run to review links manually

### Step 6: Wait for Processing

The script will now:
- Connect to your Microsoft 365 account
- Download email metadata (not full emails)
- Search for unsubscribe links
- Generate a beautiful HTML report

This may take 1-5 minutes depending on how many emails you have.

### Step 7: View Your Report

When complete:
1. The script will ask: **"Would you like to open the report now? (y/n)"**
2. Type `y` and press **Enter**
3. Your default browser will open with an interactive report!

The report is saved as: `unsubscribe-links-enhanced.html`

## 📊 Understanding Your Report

### Statistics Dashboard
At the top of the report, you'll see:
- **Total Emails**: How many emails were scanned
- **Unique Domains**: Number of different senders
- **With Unsubscribe**: Domains that have unsubscribe links
- **Priority Levels**: Breakdown by volume (High/Medium/Low)

### Interactive Features

#### 🔍 Search Box
Type a domain name or keyword to filter results instantly.

#### 📑 Sorting Options
- **Sort by Priority** - Shows high-volume senders first
- **Sort by Domain** - Alphabetical order
- **Sort by Email Count** - Most emails first

#### 🎯 Filters
- **Show All** - Display everything
- **High Priority Only** - Senders with 10+ emails
- **Medium Priority Only** - Senders with 3-9 emails
- **Low Priority Only** - Senders with 1-2 emails
- **Unclicked Only** - Links you haven't visited yet

#### ✅ Click Tracking
When you click an unsubscribe link, the row turns gray automatically. This helps you track which ones you've already processed.

#### 🔄 Reset Tracking Button
Clear all click history to start fresh.

## 🎨 Priority Badges Explained

- 🔴 **HIGH** - 10 or more emails (definitely consider unsubscribing!)
- 🟡 **MEDIUM** - 3 to 9 emails (probably worth unsubscribing)
- 🟢 **LOW** - 1 to 2 emails (might want to keep)

## 🛡️ Security & Privacy

### Is this safe?

**Yes!** Here's why:
- ✅ Uses official Microsoft Graph API
- ✅ Read-only access (can't delete or modify emails)
- ✅ OAuth authentication (doesn't store your password)
- ✅ All processing happens locally on your computer
- ✅ No data is sent to third parties
- ✅ Open-source code (you can review it)

### What permissions does it request?

- **Mail.Read** - Read your emails to find unsubscribe links
- **Mail.ReadBasic** - Access basic email metadata

### Can I revoke access?

Yes! At any time:
1. Go to [Microsoft Account Permissions](https://account.microsoft.com/permissions)
2. Find the Graph PowerShell app
3. Click **Revoke**

## ⚙️ Advanced Options

### Compile to EXE (No PowerShell Required)

Want to share with non-technical friends? Compile it to an executable:

```powershell
.\Compile-To-EXE.ps1
```

This creates `UnsubscribeLinksScanner.exe` that anyone can double-click to run!

> ⚠️ **Note**: Antivirus software may flag the EXE as suspicious. This is a false positive. You may need to add an exception.

### Scheduled Runs

Want to run this automatically every month?

1. Open **Task Scheduler** (Windows Key → type "Task Scheduler")
2. Create a new task
3. Set trigger: Monthly
4. Set action: Run the PowerShell script
5. Done! You'll get fresh reports automatically

## ❓ Troubleshooting

### "Script cannot be loaded because running scripts is disabled"

**Solution**: Run this command in PowerShell:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### "Microsoft Graph PowerShell SDK is not installed"

**Solution**: The script will auto-install it. Just wait and press `Y` when prompted.

### "Could not find a property named 'size'"

**Solution**: This is already fixed in the latest version. Make sure you have the most recent script.

### "No emails found in the specified date range"

**Possible causes**:
- The date range is too short (try 60 or 90 days)
- Selected folder is empty
- Email filtering issue

**Solution**: Try selecting "All Mail" and increase the days.

### Browser doesn't open for sign-in

**Solution**:
1. Look in PowerShell for a URL
2. Copy and paste it into your browser manually
3. Complete sign-in

### "Request failed" or "Authentication error"

**Solution**:
1. Run: `Disconnect-MgGraph`
2. Re-run the script
3. Sign in again

## 🔧 Configuration Files

### Exclude List (Optional)

Don't want to see certain domains? Create `exclude-domains.txt`:

```
company.com
important-service.com
newsletter-i-like.com
```

### Save Your Settings

The script remembers your folder choice and preferences for next time!

## 📝 Tips & Best Practices

1. **Start with 30 days** - Don't scan too far back initially
2. **Review before clicking** - Check the preview URL before unsubscribing
3. **Use filters** - Focus on high-priority senders first
4. **Run monthly** - Keep your inbox clean with regular scans
5. **Check Junk folder too** - Many newsletters end up there
6. **One-click unsubscribe** - Only enable after you trust the tool

## 🆘 Need Help?

### Common Questions

**Q: Will this delete my emails?**
A: No! It only reads emails and finds unsubscribe links. Nothing is deleted.

**Q: Do I need to keep PowerShell open?**
A: Only while the script is running. Once the HTML report is generated, you can close PowerShell.

**Q: Can I run this on Mac or Linux?**
A: Currently Windows only, but it could be adapted for PowerShell Core.

**Q: How long does it take?**
A: Usually 1-5 minutes, depending on email volume.

**Q: Will senders know I'm checking for unsubscribe links?**
A: No. The script only reads metadata, it doesn't "open" or mark emails.

### Still Stuck?

1. Check the [Issues](https://github.com/danrhodes/Outlook-auto-unsubscriber/issues) page
2. Create a new issue with:
   - Your Windows version
   - PowerShell version (`$PSVersionTable.PSVersion`)
   - Error message (copy the red text)
   - Steps you took

## 📜 Version History

### v2.0 - Enhanced Edition (Current)
- ✨ Email count per domain
- ✨ Priority detection system
- ✨ Interactive filtering and sorting
- ✨ Click tracking with localStorage
- ✨ Email subject previews
- ✨ One-click unsubscribe (RFC 8058)
- ✨ Statistics dashboard
- ✨ Beautiful modern UI
- ✅ Microsoft Graph API integration
- ✅ Multi-method unsubscribe detection
- ✅ HTML report generation
- ✅ Link validation

## 📄 License

MIT License - Free to use, modify, and distribute!

## 🙏 Credits

Created by Dan Rhodes

Built with:
- Microsoft Graph PowerShell SDK
- PowerShell 5.1+
- HTML5 & CSS3
- Vanilla JavaScript

## 🌟 Contributing

Found a bug? Have an idea? Pull requests are welcome!

---

**Made with ❤️ to help you reclaim your inbox!**

---

## Quick Reference Card

```
┌─────────────────────────────────────────────┐
│  QUICK START COMMANDS                       │
├─────────────────────────────────────────────┤
│  1. Open PowerShell in script folder        │
│  2. Set-ExecutionPolicy RemoteSigned        │
│  3. .\Find_Unsubscribe_Links-Enhanced.ps1   │
│  4. Follow prompts                          │
│  5. Open HTML report                        │
└─────────────────────────────────────────────┘

┌─────────────────────────────────────────────┐
│  PRIORITY LEVELS                            │
├─────────────────────────────────────────────┤
│  🔴 HIGH    = 10+ emails                    │
│  🟡 MEDIUM  = 3-9 emails                    │
│  🟢 LOW     = 1-2 emails                    │
└─────────────────────────────────────────────┘
```

**Remember**: You're in control. The script just helps you find links - you decide which ones to click!
