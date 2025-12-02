# Complete Development Environment Setup Guide

**From zero to fully functional Google Apps Script development with VS Code, GitHub Copilot, and clasp**

*This guide assumes you're on Windows and new to coding. We'll walk through every step.*

---

## 🎯 What You'll Have When Done

- ✅ Professional code editor (VS Code)
- ✅ AI coding assistant (GitHub Copilot)
- ✅ Version control (Git + GitHub)
- ✅ Google Apps Script development tools (clasp)
- ✅ Complete workflow: Write code → Test in Google Sheets → Deploy

---

## 📋 Prerequisites

### Accounts You'll Need (Free)

1. **GitHub Account** - For storing code and accessing Copilot
   - Go to [github.com](https://github.com)
   - Click "Sign up"
   - Follow the prompts
   - ⚠️ **IMPORTANT**: Remember your username and password!

2. **Google Account** - For Google Sheets and Apps Script
   - Use your existing Google/Gmail account
   - Or create one at [google.com](https://google.com)

### Software You'll Install (Free)

- **VS Code** - Your code editor
- **Node.js** - Required for clasp (the Google Apps Script tool)
- **Git** - Version control system
- **clasp** - Google Apps Script command-line tool

---

## 🔧 Step-by-Step Setup

### STEP 1: Create Your GitHub Account (If You Don't Have One)

1. Go to [github.com](https://github.com)
2. Click **"Sign up"** in the top-right corner
3. Enter your email address
4. Create a password (use something secure!)
5. Choose a username (this will be public)
6. Complete the verification
7. Choose the free plan

**✅ Checkpoint**: You should be logged into GitHub and see your dashboard

---

### STEP 2: Install Git

**What is Git?** Version control - it tracks changes to your code so you can go back to earlier versions if needed.

1. Download Git: [git-scm.com/download/win](https://git-scm.com/download/win)
2. Run the installer
3. **IMPORTANT**: Use these settings during installation:
   - ✅ Check "Git from the command line and also from 3rd-party software"
   - ✅ Check "Use Windows' default console window"
   - ✅ Accept all other defaults (just click "Next")

4. After installation, open **PowerShell**:
   - Press `Windows Key + X`
   - Click "Windows PowerShell" or "Terminal"

5. Verify Git is installed:
```powershell
git --version
```

Expected output:
```
git version 2.x.x
```

**✅ Checkpoint**: Git version displays without errors

---

### STEP 3: Configure Git with Your Info

Tell Git who you are (this will show up in your commits):

```powershell
git config --global user.name "Your Name"
git config --global user.email "your.email@example.com"
```

**Replace with your actual name and email** (use the email from your GitHub account)

Example:
```powershell
git config --global user.name "John Smith"
git config --global user.email "john.smith@gmail.com"
```

Verify:
```powershell
git config --global user.name
git config --global user.email
```

**✅ Checkpoint**: Your name and email display correctly

---

### STEP 4: Install Node.js

**What is Node.js?** A JavaScript runtime required to run clasp (Google Apps Script command-line tool).

1. Download Node.js: [nodejs.org](https://nodejs.org/)
   - Choose the **LTS** version (Long Term Support)
   - Download the Windows Installer (.msi)

2. Run the installer
   - ✅ Accept the license
   - ✅ Use default installation location
   - ✅ Click "Next" through all options
   - ✅ Click "Install"

3. After installation, **close and reopen PowerShell** (this is important!)

4. Verify installation:
```powershell
node --version
npm --version
```

Expected output:
```
v20.x.x (or similar)
10.x.x (or similar)
```

**✅ Checkpoint**: Both node and npm versions display

---

### STEP 5: Install VS Code

**What is VS Code?** Your code editor - where you'll write and edit code.

1. Download VS Code: [code.visualstudio.com](https://code.visualstudio.com/)
2. Run the installer
3. **IMPORTANT**: Check these boxes during installation:
   - ✅ "Add 'Open with Code' action to Windows Explorer file context menu"
   - ✅ "Add 'Open with Code' action to Windows Explorer directory context menu"
   - ✅ "Add to PATH"

4. Finish installation and launch VS Code

**✅ Checkpoint**: VS Code opens successfully

---

### STEP 6: Install GitHub Copilot (AI Assistant)

**What is GitHub Copilot?** Your AI pair programmer - it suggests code as you type!

#### 6A. Sign Up for GitHub Copilot

**Free Option** (Limited):
1. Go to [github.com/features/copilot](https://github.com/features/copilot)
2. Sign in to GitHub
3. Click "Start my free trial" or "Try Copilot Free"
4. Follow the prompts

**Pro Option** ($10/month - recommended for serious development):
1. Go to [github.com/settings/copilot](https://github.com/settings/copilot)
2. Choose "Copilot Pro" plan
3. Add payment method
4. Enable Copilot

#### 6B. Install Copilot Extensions in VS Code

1. Open VS Code
2. Click the **Extensions** icon on the left sidebar (or press `Ctrl+Shift+X`)
3. Search for "**GitHub Copilot**"
4. Click **Install** on these two extensions:
   - **GitHub Copilot** (by GitHub)
   - **GitHub Copilot Chat** (by GitHub)

5. After installation, you'll see a "Sign in to GitHub" button
6. Click it and sign in with your GitHub account
7. Authorize VS Code to access your GitHub account

**✅ Checkpoint**: You see the Copilot icon (sparkle/star) in VS Code's status bar

---

### STEP 7: Install Additional VS Code Extensions

These extensions are currently installed and recommended:

#### Essential Extensions:

1. **GitHub Copilot** ✅ (already installed)
   - AI code suggestions

2. **GitHub Copilot Chat** ✅ (already installed)
   - Chat with AI about your code

3. **PowerShell** (search and install)
   - Better PowerShell terminal support
   - Publisher: Microsoft

4. **GitHub Actions** (search and install)
   - If you plan to use GitHub automation
   - Publisher: GitHub

#### Install These in VS Code:
- Press `Ctrl+Shift+X` to open Extensions
- Search for each extension name
- Click "Install"

#### Optional Extensions (Install as needed):

- **Python** - If you work with Python scripts
- **GitLens** - Advanced Git features and history
- **Git Graph** - Visualize your Git commit history
- **Prettier** - Code formatting
- **ESLint** - JavaScript linting
- **Live Share** - Collaborate with others in real-time
- **Markdown All in One** - Better Markdown editing
- **Better Comments** - Color-coded comments
- **Bracket Pair Colorizer** - Color-matched brackets

**✅ Checkpoint**: At minimum, GitHub Copilot and Copilot Chat are installed

---

### STEP 8: Install clasp (Google Apps Script CLI)

**What is clasp?** Command-line tool to push/pull code between VS Code and Google Apps Script.

1. In PowerShell, run:
```powershell
npm install -g @google/clasp
```

This will take a minute. You'll see installation progress.

2. Verify installation:
```powershell
clasp --version
```

Expected output:
```
2.4.x
```

**✅ Checkpoint**: clasp version displays

---

### STEP 9: Authenticate clasp with Google

**This connects clasp to your Google Account:**

```powershell
clasp login
```

**What happens:**
1. A browser window opens
2. You'll see "Sign in to Google"
3. Choose your Google account
4. Click "Allow" to give clasp permission
5. You'll see "Success! Logged in as your.email@gmail.com"

**⚠️ IMPORTANT**: Use the same Google account that owns your spreadsheet!

**✅ Checkpoint**: Terminal shows "Logged in as [your email]"

---

## 🚀 Starting a New Google Apps Script Project

Now that all tools are installed, here's how to start ANY new project:

---

### STEP 10: Choose Your Starting Point

#### Option A: Start from Scratch (New Project)

1. Create a project folder:
```powershell
# Navigate to where you want your projects
cd C:\Users\YourUsername\Documents

# Create a folder for your project
mkdir MyGoogleSheetsProject
cd MyGoogleSheetsProject
```

2. Initialize Git:
```powershell
git init
```

3. Create a new Apps Script project:
```powershell
clasp create --title "My Project Name" --type standalone
```

**What this does:**
- Creates a new Apps Script project in your Google Drive
- Creates `.clasp.json` file (links your folder to the script)
- Creates `appsscript.json` file (configuration)

4. Create your first script file:
   - In VS Code: File → New File
   - Save as `Code.gs`
   - Start coding!

**✅ Checkpoint**: You have `.clasp.json` and can run `clasp push`

---

#### Option B: Clone an Existing Project from GitHub

**If someone shared a project with you:**

1. Navigate to your projects folder:
```powershell
cd C:\Users\YourUsername\Documents
```

2. Clone the repository:
```powershell
git clone https://github.com/USERNAME/REPOSITORY-NAME.git
cd REPOSITORY-NAME
```

**Replace with actual GitHub URL** (example):
```powershell
git clone https://github.com/bkaufman7/END-OF-MONTH-CM360-CPC-CPM-FLIGHT-QA---VSCODE-VERSION.git
cd "CM360 END OF MONTH AUDIT"
```

3. Link to YOUR Apps Script project (see Step 11)

**✅ Checkpoint**: Repository files are on your computer

---

### STEP 11: Link to Your Google Apps Script Project

#### If Linking to an EXISTING Apps Script Project:

**Get your Script ID:**
1. Open your Google Spreadsheet
2. Click **Extensions → Apps Script**
3. In the Apps Script editor, click the **Settings** (gear) icon
4. Copy the **Script ID**

**Create `.clasp.json` file:**

In PowerShell (make sure you're in your project folder):
```powershell
# Replace YOUR_SCRIPT_ID with the ID you copied
# Replace the path with your actual project path

$claspConfig = @"
{
  "scriptId": "YOUR_SCRIPT_ID_HERE"
}
"@

$claspConfig | Out-File -FilePath .clasp.json -Encoding utf8
```

**Example** (with real values):
```powershell
$claspConfig = @"
{
  "scriptId": "1a2b3c4d5e6f7g8h9i0j"
}
"@

$claspConfig | Out-File -FilePath .clasp.json -Encoding utf8
```

**✅ Checkpoint**: `.clasp.json` file exists in your project folder

---

### STEP 12: Open Project in VS Code

```powershell
# Open current folder in VS Code
code .
```

VS Code will open with your project files visible on the left.

**✅ Checkpoint**: You see your project files in VS Code's Explorer pane

---

### STEP 13: Push Code to Apps Script

**In VS Code terminal** (or PowerShell in your project folder):

```powershell
clasp push
```

**What happens:**
- VS Code sends your `.gs` files to Google Apps Script
- You'll see a list of files pushed

Expected output:
```
Pushed 2 files.
└─ appsscript.json
└─ Code.gs
```

**✅ Checkpoint**: Files pushed without errors

---

### STEP 14: Verify in Apps Script Web Editor

```powershell
# Open your project in browser
clasp open
```

**You should see:**
- Your code files in the Apps Script editor
- Same code as in VS Code

**✅ Checkpoint**: Code appears in Apps Script editor

---

### STEP 15: Bind Script to a Google Spreadsheet (If Needed)

**If you want your script to work with a specific Google Sheet:**

1. Open your Google Spreadsheet
2. Click **Extensions → Apps Script**
3. Delete any existing code
4. Go back to PowerShell and run:
```powershell
clasp push
```
5. Refresh your spreadsheet
6. The custom menu should appear (if your code has `onOpen()` function)

**✅ Checkpoint**: Custom menu appears in your spreadsheet

---

### STEP 16: Set Up GitHub Repository (Version Control)

**Why?** Save your code online and track changes.

#### Create a New Repository on GitHub:

1. Go to [github.com](https://github.com)
2. Click the **"+"** icon (top-right) → **"New repository"**
3. Name your repository (e.g., "my-google-sheets-automation")
4. Choose **Private** or **Public**
5. **Do NOT** check "Initialize with README" (you already have files)
6. Click **"Create repository"**

#### Connect Your Local Project to GitHub:

GitHub will show you commands. Copy them, or use these:

```powershell
# Add GitHub as the remote repository (replace with YOUR repository URL)
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git

# Add all files to git
git add .

# Make your first commit
git commit -m "Initial commit"

# Push to GitHub
git push -u origin main
```

**Example** (with real values):
```powershell
git remote add origin https://github.com/johnsmith/my-sheets-project.git
git add .
git commit -m "Initial commit"
git push -u origin main
```

**You'll be asked for credentials:**
- Username: Your GitHub username
- Password: Use a **Personal Access Token** (not your password!)

**How to create a Personal Access Token:**
1. GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)
2. Click "Generate new token"
3. Check "repo" scope
4. Copy the token and save it somewhere safe
5. Use this token as your password

**✅ Checkpoint**: Your code is visible on GitHub.com

---

## 📁 Typical Project File Structure

After complete setup:

```
YourProjectFolder/
├── .clasp.json                    # Links to your Apps Script project (DON'T commit to Git)
├── .claspignore                   # Files to ignore when pushing to Apps Script
├── .git/                          # Git repository data
├── .gitignore                     # Files to ignore in Git
├── appsscript.json                # Apps Script configuration
├── Code.gs                        # Your main script file
├── OtherFile.gs                   # Additional script files (optional)
└── README.md                      # Project documentation
```

**Important files explained:**

- **`.clasp.json`**: Links your folder to your Google Apps Script project. **Never commit this to Git!** (Has your Script ID)
- **`.gitignore`**: Tells Git which files to ignore (like `.clasp.json`)
- **`appsscript.json`**: Configuration for your Apps Script project (timezone, API services, etc.)
- **`Code.gs`**: Your actual script code (can have multiple `.gs` files)

---

## 🔄 Daily Development Workflow

**Once everything is set up, here's your daily workflow:**

### Making Changes to Your Code

1. **Open your project in VS Code**
   ```powershell
   cd C:\Users\YourUsername\YourProject
   code .
   ```

2. **Edit your code**
   - Make changes to your `.gs` files
   - GitHub Copilot will suggest code as you type!
   - Press `Tab` to accept suggestions
   - Press `Ctrl+Enter` to see more suggestions

3. **Save your changes**
   - Press `Ctrl+S` to save

4. **Push to Apps Script** (so changes work in Google Sheets)
   ```powershell
   clasp push
   ```

5. **Test in your spreadsheet**
   - Open your Google Spreadsheet
   - Use your custom menu
   - Test the functionality

6. **If it works, commit to Git** (save version)
   ```powershell
   git add .
   git commit -m "Describe what you changed"
   git push
   ```

### Example Workflow:

```powershell
# 1. Open project
code .

# 2. Make changes in VS Code, save

# 3. Push to Google Apps Script
clasp push

# 4. Test in Google Sheets

# 5. If it works, save to Git
git add .
git commit -m "Fixed date calculation bug"
git push
```

### If You Made Changes in Apps Script Web Editor

**Pull changes back to VS Code:**

```powershell
clasp pull
```

This downloads the latest code from Apps Script to your local files.

---

## 🤖 Using GitHub Copilot

### How Copilot Helps You

As you type code, Copilot suggests completions:

1. **Inline suggestions** (gray text)
   - Press `Tab` to accept
   - Press `Esc` to reject
   - Keep typing to ignore

2. **Chat with Copilot**
   - Press `Ctrl+I` for inline chat
   - Or click the Copilot Chat icon in the sidebar
   - Ask questions like:
     - "How do I loop through rows in Google Sheets?"
     - "Write a function to send an email"
     - "Explain this code"

3. **Generate entire functions**
   - Type a comment describing what you want:
   ```javascript
   // Function to find duplicate values in column A
   ```
   - Press `Enter`
   - Copilot suggests the entire function!

### Copilot Tips:

- **Be specific in comments**: "Create a function that emails a summary to admin@company.com every Monday at 9am"
- **Use descriptive variable names**: Copilot understands context better
- **Ask questions**: "Why is this function slow?" or "How can I optimize this?"
- **Learn from suggestions**: Read what Copilot suggests to learn new techniques

---

## 🛠️ Essential Commands Reference

**Keep this handy - you'll use these all the time!**

### clasp Commands (Google Apps Script)

```powershell
# Push your code TO Apps Script (local → cloud)
clasp push

# Pull code FROM Apps Script (cloud → local)
clasp pull

# Open your project in the browser
clasp open

# View your Apps Script logs
clasp logs

# Check which project you're connected to
clasp status

# Login to Google (if logged out)
clasp login

# Logout
clasp logout
```

### Git Commands (Version Control)

```powershell
# See what files changed
git status

# Add all changes to be committed
git add .

# Save changes with a message
git commit -m "Describe what you changed"

# Send changes to GitHub
git push

# Get latest changes from GitHub
git pull

# View your commit history
git log --oneline

# See what changed in files
git diff
```

### VS Code Shortcuts

```
Ctrl+S          Save file
Ctrl+Shift+P    Command palette (search for any VS Code command)
Ctrl+`          Open/close terminal
Ctrl+B          Toggle sidebar
Ctrl+Shift+X    Extensions
Ctrl+Shift+E    Explorer (file tree)
Ctrl+I          Copilot inline chat
Ctrl+/          Comment/uncomment line
Ctrl+F          Find in current file
Ctrl+Shift+F    Find in all files
```

---

## 🔐 Security & Best Practices

### Critical: Files to NEVER Commit to GitHub

**Your `.gitignore` file should include:**

```gitignore
# clasp authentication (has your Script ID!)
.clasp.json
.clasprc.json

# Node modules
node_modules/

# Environment variables (passwords, API keys)
.env

# OS files
.DS_Store
Thumbs.db
```

**Why?**
- `.clasp.json` contains your Script ID (could expose your project)
- `.env` might contain passwords or API keys
- Committing these = security risk!

### Create a `.gitignore` File:

In your project folder:

```powershell
$gitignore = @"
# clasp authentication
.clasp.json
.clasprc.json

# Node modules
node_modules/

# Environment variables
.env

# OS files
.DS_Store
Thumbs.db
"@

$gitignore | Out-File -FilePath .gitignore -Encoding utf8
```

**✅ Verify**: `.gitignore` file exists before your first `git commit`

---

## 🚨 Troubleshooting Common Issues

### 😕 "I can't push to Apps Script" - Authentication Error

**Problem**: `clasp push` says "Not authorized" or "Login required"

**Solution**:
```powershell
clasp logout
clasp login
```
Then try `clasp push` again.

---

### 😕 "No files to push" Error

**Problem**: `clasp push` says "No files to push"

**Fix 1 - Check if `.claspignore` exists:**

If you have a `.claspignore` file, it should look like:
```
**/**
!appsscript.json
!**/*.gs
```

**Fix 2 - Make sure you have `.gs` files:**

Your project needs at least one `.gs` file (e.g., `Code.gs`)

---

### 😕 "Changes don't appear in my spreadsheet"

**Problem**: You pushed code but the spreadsheet doesn't reflect changes

**Checklist**:
1. ✅ Did you run `clasp push`?
2. ✅ Did it say "Pushed X files"?
3. ✅ Refresh your spreadsheet (press `F5`)
4. ✅ Wait 10 seconds (sometimes takes a moment)
5. ✅ Check Apps Script editor: `clasp open`

If menu doesn't appear:
- Make sure your code has an `onOpen()` function
- Close and reopen the spreadsheet
- Check Executions log in Apps Script for errors

---

### 😕 "Git push failed" - Authentication Error

**Problem**: `git push` asks for password but it's rejected

**Solution**: You need a **Personal Access Token**, not your password!

**Create a token:**
1. Go to GitHub.com
2. Settings → Developer settings → Personal access tokens → Tokens (classic)
3. Click "Generate new token (classic)"
4. Check "repo" scope
5. Click "Generate token"
6. **Copy the token** (you won't see it again!)
7. Use this token as your password when `git push` asks

**Save credentials** (so you don't have to enter every time):
```powershell
git config --global credential.helper wincred
```

---

### 😕 "Permission denied" or "Forbidden"

**Problem**: Can't access spreadsheet or Apps Script

**Cause**: You're using the wrong Google account

**Solution**:
1. Make sure you're logged into the correct Google account in your browser
2. Logout and login to clasp with the RIGHT account:
   ```powershell
   clasp logout
   clasp login
   ```
3. Choose the account that owns the spreadsheet

---

### 😕 "Script ID not found"

**Problem**: `.clasp.json` has wrong Script ID

**Fix**:
1. Open your spreadsheet
2. Extensions → Apps Script
3. Click Settings (gear icon)
4. Copy the Script ID
5. Update `.clasp.json`:
   ```json
   {
     "scriptId": "PASTE_CORRECT_ID_HERE"
   }
   ```

---

### 😕 "Copilot isn't suggesting anything"

**Problem**: Copilot installed but not working

**Checklist**:
1. ✅ Are you signed into GitHub in VS Code?
   - Click Copilot icon in status bar
2. ✅ Do you have an active Copilot subscription?
   - Check [github.com/settings/copilot](https://github.com/settings/copilot)
3. ✅ Is Copilot enabled for this file type?
   - `.gs` files should work
4. ✅ Try restarting VS Code

---

### 😕 "Command not found: clasp" or "Command not found: git"

**Problem**: Installed but terminal doesn't recognize commands

**Solution**:
1. **Close and reopen PowerShell** (very important!)
2. If still not working, restart your computer
3. Verify installation:
   ```powershell
   node --version
   npm --version
   git --version
   clasp --version
   ```

---

### 😕 "How do I enable Gmail/Drive APIs?"

**Problem**: Your script needs to access Gmail or Drive

**Solution**:
1. Open Apps Script editor: `clasp open`
2. Click "Services" (+ icon on left sidebar)
3. Find "Gmail API" or "Drive API"
4. Click "Add"
5. Choose version (latest)
6. Click "Add"

**Or update `appsscript.json`** in VS Code:
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

Then run `clasp push`

---

## 🎯 Complete Setup Checklist

Use this to verify everything is working:

### Software Installation
- [ ] Git installed (`git --version` works)
- [ ] Git configured with your name and email
- [ ] Node.js installed (`node --version` works)
- [ ] npm installed (`npm --version` works)
- [ ] VS Code installed and opens
- [ ] clasp installed (`clasp --version` works)

### Accounts & Authentication
- [ ] GitHub account created
- [ ] Personal Access Token created (for Git)
- [ ] Logged into clasp (`clasp login` completed)
- [ ] Google account ready

### VS Code Setup
- [ ] GitHub Copilot extension installed
- [ ] GitHub Copilot Chat extension installed
- [ ] Signed into GitHub in VS Code
- [ ] Copilot icon shows in status bar
- [ ] PowerShell extension installed (optional)

### Project Setup
- [ ] Project folder created
- [ ] `.clasp.json` exists with Script ID
- [ ] `.gitignore` exists
- [ ] Code pushed successfully (`clasp push`)
- [ ] Can open project in browser (`clasp open`)

### Spreadsheet Integration (If Applicable)
- [ ] Script bound to spreadsheet
- [ ] Menu appears when spreadsheet opens
- [ ] Can run functions from menu

### Version Control
- [ ] GitHub repository created
- [ ] Local project connected to GitHub (`git remote -v` shows repo)
- [ ] Initial commit made
- [ ] Pushed to GitHub successfully

---

## 💡 Pro Tips

### Speed Up Your Workflow

1. **Use keyboard shortcuts**
   - `Ctrl+S` to save
   - `Ctrl+`` to open terminal
   - `Ctrl+I` for Copilot chat

2. **Terminal in VS Code**
   - No need to switch to PowerShell window
   - Terminal → New Terminal (or `Ctrl+'`)
   - Run `clasp push` right in VS Code!

3. **Multi-file editing**
   - Open multiple `.gs` files as tabs
   - Split view: Drag tab to side

4. **Copilot shortcuts**
   - Type `//` for comment suggestions
   - Start typing a function name and let Copilot complete it
   - Ask Copilot to explain errors

### Learning Resources

- **Apps Script Tutorials**: [developers.google.com/apps-script/guides](https://developers.google.com/apps-script/guides)
- **GitHub Copilot Guide**: [docs.github.com/en/copilot](https://docs.github.com/en/copilot)
- **VS Code Tips**: [code.visualstudio.com/docs/getstarted/tips-and-tricks](https://code.visualstudio.com/docs/getstarted/tips-and-tricks)

---

## 🚀 Advanced: Google Cloud Platform (Optional)

**When you might need GCP:**
- Using APIs beyond Gmail/Drive (BigQuery, Cloud Storage, etc.)
- Need higher quota limits
- Building production apps with many users
- OAuth authentication for users

**How to set up:**
1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project
3. Enable APIs you need
4. Link to your Apps Script project

**For most projects, the basic Apps Script setup is sufficient!**

---

## 📚 Where to Go from Here

### You're Ready To:
✅ Write code in VS Code with AI assistance  
✅ Test in Google Sheets  
✅ Save versions with Git  
✅ Share on GitHub  
✅ Collaborate with others

### Next Steps:
1. **Start coding!** Open VS Code and create `Code.gs`
2. **Use Copilot**: Type comments describing what you want, let it suggest code
3. **Test frequently**: `clasp push` after every change and test in your spreadsheet
4. **Commit often**: Save your progress with Git
5. **Ask Copilot for help**: It can explain errors, suggest improvements, and teach you!

### Common First Scripts:
- Custom menu with functions
- Email automation
- Data import/export
- Report generation
- Data validation

**Ask Copilot Chat**: "Show me a simple Google Sheets Apps Script example with a custom menu"

---

## 🆘 Still Stuck?

### Check These:
1. **All commands work in PowerShell?**
   - `git --version`
   - `node --version`
   - `clasp --version`

2. **Logged into everything?**
   - clasp: `clasp login`
   - VS Code: Click account icon (bottom-left)
   - GitHub: Check [github.com](https://github.com)

3. **Files in the right place?**
   - `.clasp.json` in your project folder
   - `.gs` files in your project folder
   - PowerShell is IN your project folder (`cd` to it!)

4. **Still not working?**
   - Ask GitHub Copilot! (Press `Ctrl+I` and describe the issue)
   - Check error messages carefully
   - Google the error message
   - Restart VS Code and PowerShell

---

**Last Updated**: December 2, 2025  
**Perfect for**: Ad Ops teams, Marketing Ops, Anyone automating Google Sheets  
**Difficulty**: Beginner-friendly (step-by-step)  
**Time to Complete**: 30-45 minutes
