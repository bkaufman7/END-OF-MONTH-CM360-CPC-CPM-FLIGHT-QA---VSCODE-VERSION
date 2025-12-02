# VS Code Setup Guide - CM360 End of Month Audit

**Complete setup instructions for development environment**

---

## 📋 Prerequisites

### Required Software
- **Node.js** (v14 or higher) - [Download](https://nodejs.org/)
- **Git** - [Download](https://git-scm.com/)
- **VS Code** - [Download](https://code.visualstudio.com/)
- **Google Account** with Apps Script access

---

## 🔧 Step-by-Step Setup

### 1. Install Node.js and npm
```powershell
# Verify installation
node --version
npm --version
```

Expected output:
```
v18.x.x (or higher)
9.x.x (or higher)
```

---

### 2. Install Google Apps Script CLI (clasp)

```powershell
# Install clasp globally
npm install -g @google/clasp

# Verify installation
clasp --version
```

Expected output:
```
2.4.x (or higher)
```

---

### 3. Authenticate clasp with Google

```powershell
# Login to Google Account
clasp login
```

This will:
1. Open a browser window
2. Ask you to sign in to Google
3. Request permission to manage your Apps Script projects
4. Save credentials locally

**Important**: Use the same Google account that owns your spreadsheet!

---

### 4. Clone the GitHub Repository

```powershell
# Navigate to your projects folder
cd C:\Users\bkaufman

# Clone the repository
git clone https://github.com/bkaufman7/END-OF-MONTH-CM360-CPC-CPM-FLIGHT-QA---VSCODE-VERSION.git

# Navigate into the project
cd "CM360 END OF MONTH AUDIT"
```

---

### 5. Link to Your Apps Script Project

#### Option A: If you already have a script project bound to your spreadsheet

```powershell
# Get your Script ID
# 1. Open your Google Spreadsheet
# 2. Extensions → Apps Script
# 3. Click "Project Settings" (gear icon)
# 4. Copy the "Script ID"

# Create .clasp.json file
echo '{"scriptId":"YOUR_SCRIPT_ID_HERE","rootDir":"C:\\Users\\bkaufman\\CM360 END OF MONTH AUDIT"}' | Out-File -FilePath .clasp.json -Encoding utf8
```

Replace `YOUR_SCRIPT_ID_HERE` with your actual Script ID.

#### Option B: Create a new standalone script project

```powershell
# Create new script project
clasp create --title "CM360 End of Month Audit" --type standalone

# This creates .clasp.json automatically
```

---

### 6. Configure .clasp.json

Your `.clasp.json` should look like this:

```json
{
  "scriptId": "YOUR_SCRIPT_ID",
  "rootDir": "C:\\Users\\bkaufman\\CM360 END OF MONTH AUDIT"
}
```

**Note**: Make sure to use double backslashes (`\\`) in Windows paths!

---

### 7. Push Code to Apps Script

```powershell
# Push all .gs files and appsscript.json
clasp push

# Verify files were pushed
clasp status
```

Expected output:
```
Pushed 3 files.
└─ appsscript.json
└─ AuditSystems.gs
└─ Code.gs
```

---

### 8. Open the Project in Apps Script Editor

```powershell
# Open in browser
clasp open
```

This opens the Apps Script web editor where you can:
- Run functions manually
- View logs
- Set up triggers
- Check execution history

---

## 🎨 VS Code Extensions (Recommended)

### Install These Extensions

1. **GitHub Copilot** (optional but helpful)
   - AI-powered code completion
   - Install: `Ctrl+Shift+X` → Search "GitHub Copilot"

2. **Apps Script** by Eiji Kitamura (optional)
   - Syntax highlighting for .gs files
   - Install: Search "Apps Script" in extensions

3. **Git Graph** (optional)
   - Visualize git history
   - Install: Search "Git Graph"

4. **GitLens** (optional)
   - Enhanced git capabilities
   - Install: Search "GitLens"

---

## 📁 Project File Structure

After setup, you should have:

```
CM360 END OF MONTH AUDIT/
├── .clasp.json                    # Clasp configuration (DO NOT commit)
├── .claspignore                   # Files to ignore when pushing
├── .git/                          # Git repository
├── .gitignore                     # Git ignore rules
├── appsscript.json                # Apps Script manifest
├── Code.gs                        # Main QA script (7,226 lines)
├── AuditSystems.gs                # Audit systems (567 lines)
├── README.md                      # User documentation
├── DEPLOYMENT.md                  # Setup instructions
├── VSCODE-SETUP-GUIDE.md          # This file
├── DEVELOPMENT-PROGRESS.md        # Git history & milestones
├── COMPREHENSIVE-AUDIT-GUIDE.md   # Audit system guide
└── RAW-DATA-ARCHIVE-GUIDE.md      # Archive system guide
```

---

## 🔄 Daily Development Workflow

### Making Changes

```powershell
# 1. Edit code in VS Code
# 2. Save changes

# 3. Push to Apps Script
clasp push

# 4. Test in spreadsheet
# Open your spreadsheet and use the menu

# 5. Commit to git
git add .
git commit -m "Description of changes"
git push origin main
```

### Pulling Latest from Apps Script

```powershell
# If you made changes in the web editor
clasp pull

# This downloads the latest from Apps Script
```

---

## 🛠️ Common Commands Reference

### clasp Commands

```powershell
# Push local code to Apps Script
clasp push

# Pull code from Apps Script to local
clasp pull

# Open project in browser
clasp open

# View recent deployments
clasp deployments

# View project info
clasp status

# Run a function
clasp run functionName

# View logs
clasp logs

# Create a new version
clasp version "Version description"

# Deploy as web app
clasp deploy
```

### Git Commands

```powershell
# Check status
git status

# View changes
git diff

# View commit history
git log --oneline

# View all commits
git log --oneline --all

# Add all changes
git add .

# Commit with message
git commit -m "Your message"

# Push to GitHub
git push origin main

# Pull from GitHub
git pull origin main
```

---

## 🔐 Security Best Practices

### Files to NEVER Commit

Add these to `.gitignore`:

```
.clasp.json
.clasprc.json
node_modules/
.env
```

Your `.gitignore` should already have:
```gitignore
# clasp local auth
.clasp.json
.clasprc.json

# Node modules
node_modules/

# Environment variables
.env
```

---

## 🚨 Troubleshooting

### Issue: "clasp push" fails with authentication error

**Solution**:
```powershell
# Re-authenticate
clasp logout
clasp login
```

---

### Issue: "clasp push" says "No files to push"

**Solution**:
```powershell
# Check .claspignore - make sure it's not ignoring .gs files
# The file should look like:
**/**
!appsscript.json
!**/*.gs
```

---

### Issue: Changes in VS Code don't appear in spreadsheet

**Solution**:
```powershell
# 1. Make sure you pushed
clasp push

# 2. Refresh your spreadsheet (F5)
# 3. Check Apps Script editor
clasp open
```

---

### Issue: "Script ID not found"

**Solution**:
```powershell
# Check your .clasp.json has the correct Script ID
# Get Script ID from:
# Spreadsheet → Extensions → Apps Script → Project Settings → Script ID
```

---

### Issue: Permission denied errors

**Solution**:
```powershell
# Make sure you're using the same Google account that owns:
# - The spreadsheet
# - The Apps Script project

# Re-login with correct account
clasp logout
clasp login
```

---

### Issue: Code appears in Apps Script but menu doesn't show

**Solution**:
1. Close and reopen the spreadsheet
2. Wait a few seconds for `onOpen()` to run
3. Refresh the page (F5)
4. Check Apps Script execution logs for errors:
   - Extensions → Apps Script → Executions

---

## 📊 Binding Script to Spreadsheet

### If you need to bind a standalone script to a spreadsheet:

```powershell
# 1. Open your spreadsheet
# 2. Extensions → Apps Script
# 3. Copy the existing code (backup!)
# 4. Delete all code in the editor
# 5. In your terminal:
clasp push

# 6. Refresh spreadsheet
# 7. Menu should appear
```

---

## ⚙️ Apps Script Configuration

Your `appsscript.json` should have:

```json
{
  "timeZone": "America/New_York",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Gmail",
        "version": "v1",
        "serviceId": "gmail"
      },
      {
        "userSymbol": "Drive",
        "version": "v3",
        "serviceId": "drive"
      }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

---

## 🎯 First-Time Setup Checklist

- [ ] Node.js installed and verified
- [ ] clasp installed globally (`npm install -g @google/clasp`)
- [ ] Authenticated with Google (`clasp login`)
- [ ] Repository cloned from GitHub
- [ ] `.clasp.json` created with correct Script ID
- [ ] Code pushed successfully (`clasp push`)
- [ ] Spreadsheet opened and menu visible
- [ ] VS Code extensions installed (optional)
- [ ] Gmail API enabled in Apps Script project
- [ ] Drive API enabled in Apps Script project
- [ ] Test run: "Run It All" executes successfully

---

## 📚 Additional Resources

### Official Documentation
- [clasp Documentation](https://github.com/google/clasp)
- [Apps Script Documentation](https://developers.google.com/apps-script)
- [Apps Script Reference](https://developers.google.com/apps-script/reference)

### Project Documentation
- `README.md` - System overview
- `DEPLOYMENT.md` - Deployment instructions
- `DEVELOPMENT-PROGRESS.md` - Git history & milestones
- `COMPREHENSIVE-AUDIT-GUIDE.md` - Audit system details
- `RAW-DATA-ARCHIVE-GUIDE.md` - Archive system details

---

## 🆘 Getting Help

### Check Execution Logs
```powershell
# View recent logs
clasp logs

# Or in Apps Script web editor:
# Extensions → Apps Script → Executions
```

### View Script Properties
In Apps Script web editor:
- File → Project properties → Script properties
- File → Project properties → Document properties

### Enable API Services
1. Open Apps Script web editor (`clasp open`)
2. Click "Services" (+ icon on left sidebar)
3. Add Gmail API (v1)
4. Add Drive API (v3)

---

**Last Updated**: December 2, 2025  
**Author**: bkaufman7  
**Project**: CM360 End of Month Audit
