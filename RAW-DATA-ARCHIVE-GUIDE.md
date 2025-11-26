# CM360 Raw Data Archive System - Complete Guide

## Overview
The Raw Data Archive System saves ALL daily "BKCM360 Global QA Check" emails from April-November 2025 to Google Drive for ROI analysis. This is Phase 1 of a two-phase system.

## Two-Phase Approach

### Phase 1: Save Everything (Weekend Job)
**What it does:**
- Processes emails month-by-month (April through November 2025)
- Saves ALL CSV/XLSX/ZIP attachments from every email
- No network filtering during save (prevents file loss)
- Auto-resumes every 10 minutes via trigger
- Organizes by date first: `Raw Data/2025/04-April/2025-04-15/`

**Expected Output:**
- ~2,400 emails processed (8 months × 30 days)
- ~9,000+ files saved (all networks, all days)
- Completion time: ~20 hours (runs over weekend)

### Phase 2: Categorize by Network (After Archive Complete)
**What it does:**
- Scans all saved files and extracts network ID from filename
- Moves files to network folders: `Raw Data/Networks/898158 - A Place for Mom/2025-04-15/`
- Preserves date organization within network folders
- Provides categorization statistics (files per network)

**Expected Output:**
- Files organized by 37 networks
- Uncategorized files remain in date folders (for manual review)
- Statistics report via email

---

## Step-by-Step Usage Guide

### Weekend Setup (Friday Evening)

1. **Start Archive**
   - Open spreadsheet
   - Go to menu: `CM360 QA Tools > 📦 Raw Data Archive > 📦 Archive All Raw Data`
   - Confirm dialog (shows ~2,400 emails expected)
   - System starts processing immediately

2. **Create Auto-Resume Trigger**
   - Go to menu: `CM360 QA Tools > 📦 Raw Data Archive > ⏰ Create Auto-Resume Trigger`
   - Confirms trigger created (runs every 10 minutes)
   - System will now auto-resume until complete

3. **Go Home for Weekend**
   - System runs automatically Friday night → Monday morning
   - Emails sent after each month completes (8 total)
   - Final completion email when done

### Monday Morning (Check Results)

1. **View Progress**
   - Menu: `📦 Raw Data Archive > 📊 View Raw Data Progress`
   - Shows: Status, emails processed, files saved, current month

2. **Check Completion Email**
   - Subject: "✅ CM360 Raw Data Archive Complete"
   - Contains: Total emails, total files, duration, Drive link

3. **Delete Trigger** (if complete)
   - Menu: `📦 Raw Data Archive > 🛑 Delete Auto-Resume Trigger`
   - Stops auto-resume from running

### After Archive Complete

1. **Categorize Files by Network**
   - Menu: `📦 Raw Data Archive > 📂 Categorize Files by Network`
   - Confirm dialog
   - System scans all files and organizes by network ID
   - Completion email with stats

---

## Folder Structure

### During Archive (Phase 1)
```
Raw Data/
├── 2025/
│   ├── 04-April/
│   │   ├── 2025-04-01/
│   │   │   ├── file1.csv
│   │   │   ├── file2.xlsx
│   │   ├── 2025-04-02/
│   │   ├── ... (all 30 days)
│   ├── 05-May/
│   ├── ... (through 11-November)
```

### After Categorization (Phase 2)
```
Raw Data/
├── 2025/
│   └── ... (original date folders remain)
├── Networks/
│   ├── 898158 - A Place for Mom/
│   │   ├── 2025-04-01/
│   │   ├── 2025-04-02/
│   │   └── ... (all dates with files for this network)
│   ├── 1283860 - ADT/
│   │   └── ... (dates)
│   └── ... (37 networks total)
```

---

## Email Notifications

### Progress Emails (8 total)
- **Frequency:** After each month completes
- **Subject:** "📦 CM360 Raw Data: April 2025 Complete" (example)
- **Contains:** Emails processed, files saved, total progress, next month

### Archive Completion Email
- **Subject:** "✅ CM360 Raw Data Archive Complete - Summary Report"
- **Contains:** 
  - Overall statistics (emails, files, averages)
  - Performance metrics (duration, start/end times)
  - File location and structure
  - Next steps for categorization

### Error Email (if issue occurs)
- **Subject:** "⚠️ CM360 Raw Data Archive Error"
- **Contains:** Current month, error message, progress stats
- **Action:** Use "Resume Raw Data Archive" to continue

### Categorization Completion Email
- **Subject:** "✅ CM360 Raw Data Categorization Complete - Summary Report"
- **Contains:**
  - Overall statistics (total files, categorized %, uncategorized %)
  - Performance metrics (duration, processing rate)
  - Top 10 networks by file count
  - File locations and next steps for ROI analysis

---

## Menu Functions Reference

### Archive All Raw Data (Apr-Nov 2025)
- **Purpose:** Start archiving all raw data files
- **When:** Run once on Friday evening
- **Duration:** Processes first month, then relies on trigger

### View Raw Data Progress
- **Purpose:** Check current status
- **When:** Anytime during weekend/Monday
- **Shows:** Current month, emails processed, files saved

### Resume Raw Data Archive
- **Purpose:** Manually resume if stopped
- **When:** After error or if trigger deleted
- **Action:** Continues from last saved state

### Create Auto-Resume Trigger
- **Purpose:** Set up automatic resumption
- **When:** After starting archive
- **Frequency:** Every 10 minutes
- **Note:** Auto-deletes when archive completes

### Delete Auto-Resume Trigger
- **Purpose:** Stop automatic resumption
- **When:** After archive completes or to pause
- **Effect:** Stops scheduled executions

### Categorize Files by Network
- **Purpose:** Organize saved files by network
- **When:** After archive 100% complete
- **Duration:** ~10-20 minutes for all files

---

## Technical Details

### Performance
- **Batch Size:** 20 emails per execution (conservative)
- **Execution Time:** ~5 minutes per month
- **Total Runtime:** ~40 minutes actual execution (8 months × 5 min)
- **Wall Clock Time:** ~20 hours (with 10-min intervals)

### State Management
- Archive state saved in Script Properties
- Tracks: current month, emails processed, files saved
- Resumes from last successful month
- Never re-processes completed months

### File Handling
- **ZIP Files:** Auto-extracts CSV/XLSX files
- **Duplicates:** Skips if filename already exists
- **Naming:** Preserves original filenames from email
- **Formats:** CSV, XLSX (ignores other file types)

### Network Detection
- Scans filename for 3-7 digit numbers
- Compares against Networks tab (Column A)
- Supports patterns like: `898158_report.csv`, `DCM_898158.zip`
- Files without network ID marked "uncategorized"

---

## Troubleshooting

### Archive Not Progressing
1. Check trigger exists: Menu > Delete Trigger (see if any found)
2. View progress to see current state
3. Manually resume if needed

### Missing Files in Archive
- Archive saves ALL attachments (no filtering)
- Check email count matches expected
- Review Drive folder structure

### Categorization Missing Files
- Files without network ID in filename stay uncategorized
- Check "filesUncategorized" count in completion email
- Manually review uncategorized files if needed

### Trigger Not Auto-Deleting
- Manually delete after completion confirmed
- Check Script Properties for state (should show "completed")

---

## ROI Analysis (Future)

Once archive is complete, you'll have:
- **All raw data files** from April-November 2025
- **Organized by network** and date
- **Ready for analysis:**
  - Violations detected per network/month
  - Violations rectified (changes made)
  - Cost savings ($) per network
  - Trend analysis (improving vs. declining)

**Next Steps After Archive:**
1. Build ROI dashboard (analyze historical data)
2. Compare violations detected vs. billed
3. Calculate total $ saved over 8 months
4. Present to leadership with data-backed results

---

## Important Notes

⚠️ **Do NOT run archive twice** - Check progress first to avoid duplicates

✅ **Set up trigger** - Archive won't auto-resume without it

📧 **Check emails** - Progress updates keep you informed

🕒 **Weekend job** - Plan to run Friday night through Monday

📂 **Categorize after** - Phase 2 only after Phase 1 complete

💾 **Drive space** - Archive will use ~1-2 GB (estimate)

---

## Quick Reference

| Task | Menu Path | When |
|------|-----------|------|
| Start Archive | Raw Data Archive > Archive All | Friday Evening |
| Set Trigger | Raw Data Archive > Create Trigger | After Start |
| Check Status | Raw Data Archive > View Progress | Anytime |
| Stop Archive | Raw Data Archive > Delete Trigger | To Pause |
| Resume Archive | Raw Data Archive > Resume | After Error |
| Organize Files | Raw Data Archive > Categorize | After Complete |

---

## Questions?

**Is it saving all days or just last 2 days?**
✅ NEW VERSION SAVES ALL DAYS (no network filtering during save)

**How long will it take?**
⏱️ ~20 hours wall clock time (runs over weekend with trigger)

**Can I pause it?**
⚠️ Delete trigger to pause, Resume to continue (picks up where left off)

**What if something goes wrong?**
📧 Error email sent with details, use Resume to continue

**How do I know it's done?**
✅ Completion email + Status shows "completed"
