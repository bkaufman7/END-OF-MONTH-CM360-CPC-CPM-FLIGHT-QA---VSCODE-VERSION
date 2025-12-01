# Comprehensive Archive Audit Guide

## Overview
The Comprehensive Audit validates that **ALL** email attachments from Gmail are properly saved in Google Drive. Unlike the Quick Audit (which only checks if dates have at least one file), this audit verifies every single file.

## How It Works

### Three-Phase Execution
Due to the 6-minute execution limit, the audit runs in phases:

1. **Phase 1: Gmail Scan** (15-20 minutes total)
   - Searches `subject:"BKCM360 Global QA Check"`
   - Processes 100 email threads at a time
   - Extracts all CSV/XLSX files (including inside ZIPs)
   - Builds map: `"date|networkId|filename"` → `filename`
   - Saves progress every batch
   - Auto-resumes until complete (~8,880 emails)

2. **Phase 2: Drive Scan** (3-5 minutes)
   - Scans `Raw Data/2025/MM-Month/YYYY-MM-DD/` folders
   - Counts all CSV/XLSX files
   - Builds same format map as Gmail
   - Usually completes in single run

3. **Phase 3: Compare & Report** (< 1 minute)
   - Compares Gmail vs Drive
   - Identifies missing files (in Gmail, not Drive)
   - Identifies extra files (in Drive, not Gmail)
   - Sends detailed email report
   - Cleans up state

### State Management
The audit uses DocumentProperties to save progress:
```javascript
{
  phase: 'gmail_scan' | 'drive_scan' | 'compare',
  gmailStartIndex: 4500,  // Threads scanned so far
  expectedFilesJson: '{"2025-05-01|485401|file.csv":"file.csv",...}',
  actualFilesJson: '{"2025-05-01|485401|file.csv":"file.csv",...}',
  startTime: '2025-12-01T10:00:00.000Z'
}
```

## How to Run

### Option 1: Manual Execution
1. **Menu → ARCHIVE TOOLS → 🔬 Comprehensive Audit (Gmail vs Drive)**
2. Confirm the audit start
3. **Menu → ARCHIVE TOOLS → ⏰ Create Auto-Resume Trigger**
   - This creates a 10-minute trigger for automatic continuation
4. Wait for completion email (typically 20-30 minutes total)

### Option 2: Manual Resume (if interrupted)
- **Menu → ARCHIVE TOOLS → 🔄 Resume Comprehensive Audit**
  - Continues from last saved position
  - Use this if trigger fails or you want to speed up completion

### Check Progress
- **Menu → ARCHIVE TOOLS → 📊 View Audit Progress**
  - Shows current phase
  - Displays counts (threads scanned, files found)
  - Shows start time

### Reset Audit
- **Menu → ARCHIVE TOOLS → 🔄 Reset Comprehensive Audit**
  - Clears all state
  - Use if audit gets stuck or you want to start fresh

## Expected Results

### Success Email
```
Subject: ✅ Comprehensive Archive Audit Complete

Gmail Expected Files: 15,240
Drive Actual Files: 15,240
Missing Files: 0
Extra Files: 0

✅ Archive is complete and matches Gmail perfectly!
```

### Issues Found Email
```
Subject: ⚠️ Comprehensive Archive Audit Complete (Issues Found)

Gmail Expected Files: 15,240
Drive Actual Files: 15,180
Missing Files: 60
Extra Files: 0

=== MISSING FILES (In Gmail, Not in Drive) ===
Date: 2025-05-15 | Network: 485401 | File: BKCM360_485401_20250515.csv
Date: 2025-05-15 | Network: 485402 | File: BKCM360_485402_20250515.csv
...

These files need to be re-archived from Gmail emails.
```

## Performance Notes

### Timing Estimates
- **8,880 emails**: ~15-20 minutes (Gmail scan)
- **15K files**: ~3-5 minutes (Drive scan)
- **Comparison**: < 1 minute
- **Total**: 20-30 minutes with auto-resume

### Batch Processing
- Gmail: 100 threads per batch
- Time budget: 5 minutes per run (1-minute buffer)
- Progress logged every 500 threads
- Safety limit: 20,000 threads max

### Auto-Resume Trigger
- Runs every 10 minutes
- Continues any in-progress audit
- Also handles raw data archive
- Automatically deleted when audit completes

## Troubleshooting

### Audit Stuck?
1. Check progress: **View Audit Progress**
2. If not advancing after 30 min:
   - **Reset Comprehensive Audit**
   - Check Apps Script logs for errors
   - Verify Gmail API quota (daily limit: 10,000 reads)

### Error Email Received?
- Progress is saved automatically
- Simply run **Resume Comprehensive Audit**
- Or let auto-resume trigger continue (if enabled)

### Want to Start Over?
1. **Reset Comprehensive Audit**
2. Start fresh with **Comprehensive Audit (Gmail vs Drive)**

### Quota Exceeded?
- Gmail API limit: 10,000 reads/day
- Audit uses ~9,000 reads for full scan
- If hit: Wait 24 hours, then resume
- Progress is saved, won't lose work

## Key Differences: Quick vs Comprehensive

| Feature | Quick Audit | Comprehensive Audit |
|---------|-------------|---------------------|
| Speed | ~30 seconds | ~20-30 minutes |
| Validation | Date has ≥1 file | ALL files verified |
| Gmail Scan | No | Yes (full scan) |
| Drive Scan | Folder check | Full file count |
| Missing Files | Date level | File level |
| Chunking | No | Yes (3 phases) |
| Auto-Resume | No | Yes (trigger-based) |

## When to Use Each

### Use Quick Audit When:
- Quick sanity check needed
- Verifying gap-fill completion
- Daily monitoring
- Just need to know if dates are covered

### Use Comprehensive Audit When:
- Preparing for ROI analysis
- Ensuring 100% archive completeness
- Investigating missing data reports
- Annual data verification
- Before major migration/reorganization

## Technical Details

### File Detection Logic
```javascript
// Direct CSV/XLSX attachments
if (filename.endsWith('.csv') || filename.endsWith('.xlsx')) {
  networkId = extractNetworkId(filename);
  key = `${date}|${networkId}|${filename}`;
}

// Files inside ZIP attachments
if (filename.endsWith('.zip')) {
  unzipped = Utilities.unzip(attachment);
  for (file in unzipped) {
    if (file.endsWith('.csv') || file.endsWith('.xlsx')) {
      key = `${date}|${networkId}|${file.getName()}`;
    }
  }
}
```

### Network ID Extraction
- Uses Networks sheet mapping
- Matches filename patterns: `BKCM360_485401_20250515.csv`
- Falls back to pattern matching if not in Networks sheet
- Unknown networks logged but not counted

### Drive Structure
```
Raw Data/
  2025/
    05-May/
      2025-05-01/
        BKCM360_485401_20250501.csv
        BKCM360_485402_20250501.csv
        ...
      2025-05-02/
        ...
    06-June/
      ...
```

## Integration with Existing Tools

### Works With
- ✅ Auto-Resume Trigger (shared with raw data archive)
- ✅ Networks sheet (for network ID mapping)
- ✅ Raw Data Drive folder structure
- ✅ Gmail "BKCM360 Global QA Check" emails

### Independent Of
- ❌ Daily QA processing (different workflow)
- ❌ Violation tracking (different purpose)
- ❌ Monthly summaries (different schedule)

## Best Practices

1. **Create Auto-Resume Trigger First**
   - Ensures completion even if you close browser
   - Handles execution timeouts automatically

2. **Run During Off-Hours**
   - Lower API quota usage
   - Won't interfere with daily processing

3. **Monitor First Run**
   - Check progress every 10-15 minutes
   - Verify state advancing correctly
   - Ensure no errors in Apps Script logs

4. **Save Results Email**
   - Document your archive completeness
   - Reference for future audits
   - Track improvement over time

5. **Address Issues Promptly**
   - Missing files = data gaps
   - Use gap-fill archive for missing dates
   - Extra files = review for duplicates

## Future Enhancements

Potential improvements:
- Email-by-email verification (match Drive files to source emails)
- Network-by-network breakdowns in report
- Historical audit comparison
- Automated gap-fill trigger when missing files found
- Drive folder cleanup suggestions (extra files)
