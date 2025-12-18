# Smart Gap Fill Automation - Complete Flow Analysis

## 🎯 What You Click

**Menu Path:**
```
CM360 QA Tools → 🎯 Violations Gap Fill → 🤖 Start Smart Automation (15min process + 10min refresh)
```

**Function Called:** `startSmartGapFillAutomation()`

---

## 🔍 DEEP DIVE: Complete Execution Flow

### **PHASE 1: Initialization (30-60 seconds)**

#### Step 1.1: Check for Existing Automation
```javascript
startSmartGapFillAutomation()
├── Check ScriptProperties for existing triggers
├── If found: Alert user "Already Running" → Ask to restart Y/N
└── If No: Return and exit
```

**Properties Checked:**
- `SMART_PROCESS_TRIGGER_KEY` = 'smart_gap_fill_process_trigger'
- `SMART_REFRESH_TRIGGER_KEY` = 'smart_gap_fill_refresh_trigger'

#### Step 1.2: Scan Drive for Missing Reports
```javascript
setupAndRefreshViolationsAudit()  [Line 11399 in Code.gs]
├── setupViolationsAudit()  [AuditSystems.gs:395]
│   ├── Create "Violations Audit" sheet if doesn't exist
│   ├── Set headers: Date | Status | Report Date | Folder
│   └── Format: Frozen rows, bold headers, green background
│
└── refreshViolationsAudit()  [AuditSystems.gs:436]
    ├── Get Drive root: '1F53lLe3z5cup338IRY4nhTZQdUmJ9_wk'
    ├── Navigate: Root → Violations Reports folder
    ├── Generate date range: April 15, 2025 → TODAY (only 15th-31st)
    ├── For each date:
    │   ├── Look for folder: YYYY-MM (e.g., "2025-04")
    │   ├── Look for file: Violations_YYYY-MM-DD.xlsx
    │   ├── If found: Mark "✅ FOUND" + report date + folder path
    │   └── If missing: Mark "❌ MISSING"
    └── Write results to "Violations Audit" sheet
```

**Critical Constants:**
```javascript
VIOLATIONS_ROOT_FOLDER_ID = '1F53lLe3z5cup338IRY4nhTZQdUmJ9_wk'
START_DATE = new Date('2025-04-15')
END_DATE = new Date()  // Today
```

**Date Range Logic:**
```javascript
// Only dates 15-31 of each month
for each month from April 2025 to current month:
  for day = 15 to last day of month:
    add YYYY-MM-DD to list
```

**Expected Output:**
- Violations Audit sheet populated with ~135 dates (Apr 15 - Dec 18, 2025)
- Each row shows: Date | ✅/❌ Status | Report Date | Folder Path

#### Step 1.3: Get Missing Dates
```javascript
getMissingDatesFromAudit_()  [Line 6913 in Code.gs]
├── Read "Violations Audit" sheet
├── Filter rows where Status contains "MISSING"
├── Skip dates before 2025-04-14 (no data exists)
└── Return array of date strings
```

**Example Return:**
```javascript
[
  "Wed Apr 23 2025 00:00:00 GMT-0400",
  "Thu Apr 24 2025 00:00:00 GMT-0400",
  "Fri Apr 25 2025 00:00:00 GMT-0400",
  ... // 60 total dates
]
```

**🚨 CRITICAL CHECK #1:**
- If `missingDates.length === 0`:
  - Shows toast: "All violations reports are present"
  - **STOPS** - No automation created
- Else: Continue to next step

#### Step 1.4: Create Progress Tracking Sheet
```javascript
setupGapFillProgressSheet()  [Line 6880 in Code.gs]
├── Create "Gap Fill Progress" sheet if doesn't exist
├── Headers: Date | Status | Last Updated | Attempts | Error Message | Drive File
├── Format: Bold headers, green background, frozen rows
└── Auto-sizing columns

initializeGapFillProgress_(missingDates)  [Line 6954 in Code.gs]
├── Clear existing data in sheet
├── For each missing date:
│   └── Add row: [date, "🔄 Queued", timestamp, 0, "", ""]
└── Return count of rows added
```

**Example Progress Sheet:**
| Date | Status | Last Updated | Attempts | Error Message | Drive File |
|------|--------|--------------|----------|---------------|------------|
| Wed Apr 23 2025 | 🔄 Queued | 12/18/2025, 6:18:02 PM | 0 | | |
| Thu Apr 24 2025 | 🔄 Queued | 12/18/2025, 6:18:02 PM | 0 | | |

#### Step 1.5: Initialize State (ScriptProperties)
```javascript
state = {
  queue: [array of 60 date strings],
  currentDate: null,
  currentStep: null,
  startTime: "2025-12-18T18:18:02.000Z",
  processed: 0,
  successful: 0,
  failed: 0,
  totalToProcess: 60,
  completedFiles: []
}

saveGapFillState_(state)  [Stores in ScriptProperties]
```

**Storage Key:** `'gap_fill_state'`

#### Step 1.6: Create Automation Triggers
```javascript
// Process Trigger (Every 15 minutes)
processTrig = ScriptApp.newTrigger('smartProcessNextDate')
  .timeBased()
  .everyMinutes(15)
  .create()

// Store trigger ID
props.setProperty(SMART_PROCESS_TRIGGER_KEY, processTrig.getUniqueId())

// Refresh Trigger (Every 10 minutes)
refreshTrig = ScriptApp.newTrigger('smartRefreshAudit')
  .timeBased()
  .everyMinutes(10)
  .create()

// Store trigger ID
props.setProperty(SMART_REFRESH_TRIGGER_KEY, refreshTrig.getUniqueId())
```

**🚨 CRITICAL CHECK #2:**
- Google Apps Script only allows: 1, 5, 10, 15, or 30 minute intervals
- Using 15 minutes for processing (was 20, caused error)
- Using 10 minutes for refresh

#### Step 1.7: Process First Date Immediately
```javascript
smartProcessNextDate()  // Don't wait 15 min, start now!
```

---

### **PHASE 2: Per-Date Processing (Every 15 minutes)**

**Trigger Function:** `smartProcessNextDate()` [Line 11552]

#### Step 2.1: Load State
```javascript
state = getGapFillState_()  // Load from ScriptProperties

if (!state || state.queue.length === 0):
  ├── Log: "Gap fill complete"
  ├── stopSmartGapFillAutomation()
  ├── sendCompletionNotification_()
  └── EXIT
```

#### Step 2.2: Get Next Date
```javascript
dateStr = state.queue[0]  // e.g., "Wed Apr 23 2025 00:00:00 GMT-0400"
state.currentDate = dateStr

updateGapFillProgress_(dateStr, '🔄 Running Time Machine...', '', '')
```

**Gap Fill Progress Updated:**
| Date | Status | Last Updated | Attempts | Error Message | Drive File |
|------|--------|--------------|----------|---------------|------------|
| Wed Apr 23 2025 | 🔄 Running Time Machine... | 12/18/2025, 6:18:05 PM | 1 | | |

#### Step 2.3: Run Time Machine
```javascript
runTimeMachineForDate_(dateStr)  [Line 7283]
```

**🔍 TIME MACHINE DEEP DIVE:**

##### Step 2.3.1: Validate Date
```javascript
if (dateStr < '2025-04-14'):
  return {success: false, error: 'No data before 4.14.25'}
```

##### Step 2.3.2: Clear Sheets
```javascript
clearRawData()  [Line 208]
├── Get "Raw Data" sheet
├── Auto-create if missing (with headers)
└── Clear all data rows (keep headers)

clearViolations()  [Line 183]
├── Get "Violations" sheet
├── Auto-create if missing (with headers)
└── Clear all data rows (keep headers)
```

**Raw Data Headers:**
```
Network ID | Advertiser | Placement ID | Placement | Campaign |
Placement Start Date | Placement End Date | Campaign Start Date |
Campaign End Date | Ad | Impressions | Clicks | Report Date
```

**Violations Headers:**
```
Network | Advertiser | Campaign | Placement | Placement ID |
Start Date | End Date | Cost Structure | Issue Type | Owner |
Severity | $ at Risk | Low Priority | Report Date | Notes
```

##### Step 2.3.3: Download Raw Data
```javascript
downloadRawDataForDate_(dateStr)  [Line 6442]
```

**🔍 DATA DOWNLOAD FLOW:**

**Option A: Try Drive First (FAST)**
```javascript
downloadRawDataFromDrive_(dateStr)  [Line 6470]
├── Parse date: "Wed Apr 23 2025" → year=2025, month=4, day=23
├── Build path: "2025/04-April/2025-04-23/"
├── Navigate Drive:
│   ├── Root: '1qA77_YET8RLiES7X7NoUT5jzTHDJ3k61'
│   ├── Year folder: "2025"
│   ├── Month folder: "04-April"
│   └── Date folder: "2025-04-23"
├── Get all .csv files in folder
├── For each CSV:
│   ├── Extract Network ID: filename.split('_')[0]
│   ├── Example: "1068_BKCM360_..." → networkId = "1068"
│   ├── Read file content
│   ├── Parse CSV with processCSV()
│   └── Write to Raw Data sheet starting row 2
└── Return {success: true, filesProcessed: 34, source: 'Drive'}
```

**If Drive Folder Not Found:**
```
Log: "⚠️ Drive folder not found. Falling back to Gmail..."
```

**Option B: Gmail Fallback (SLOWER)**
```javascript
importDCMReportsForDate_(dateStr)  [Line 6468 - NEW FUNCTION]
├── Format date for Gmail: "2025/04/23"
├── Calculate next day: "2025/04/24"
├── Search Gmail:
│   └── Query: "label:CM360 QA after:2025/04/23 before:2025/04/24"
├── Found 34 threads
├── For each thread:
│   └── For each message:
│       └── For each attachment:
│           ├── If CSV:
│           │   ├── Extract Network ID from filename
│           │   ├── Parse with processCSV()
│           │   └── Append to extractedData[]
│           └── If ZIP:
│               ├── Unzip
│               ├── Find CSVs inside
│               ├── Extract Network IDs
│               └── Append to extractedData[]
├── Write all data to Raw Data sheet (bulk write)
└── Return {success: true, filesProcessed: 34, rowsImported: 4523, source: 'Gmail (CM360 QA label)'}
```

**🚨 CRITICAL CHECK #3: Gmail Label**
- **MUST** have emails labeled: "CM360 QA"
- If no label found: Returns error "No emails found"
- User needs to verify Gmail labels are applied

**CSV Processing:**
```javascript
processCSV(content, networkId)  [Line 227]
├── Split by newlines
├── Find header row (starts with "Advertiser")
├── Parse CSV data
├── Remove header row
├── For each data row:
│   └── Prepend networkId + append reportDate
└── Return array of rows
```

**Example Row Output:**
```javascript
[
  "1068",  // Network ID
  "Acme Corp",  // Advertiser
  "12345",  // Placement ID
  "Banner 728x90",  // Placement
  "Summer Campaign",  // Campaign
  "2025-06-01",  // Placement Start
  "2025-08-31",  // Placement End
  "2025-06-01",  // Campaign Start
  "2025-08-31",  // Campaign End
  "Display Ad",  // Ad
  "15000",  // Impressions
  "245",  // Clicks
  "2025-04-23"  // Report Date (today)
]
```

##### Step 2.3.4: Run QA Analysis
```javascript
runQAOnly()  [Line 1058]
```

**🔍 QA ANALYSIS DEEP DIVE:**

**Document Lock (Prevent Overlapping Runs):**
```javascript
dlock = LockService.getDocumentLock()
if (!dlock.tryLock(30000)):  // 30 second timeout
  scheduleNextQAChunk_(2)  // Retry in 2 min
  return
```

**Load Configuration:**
```javascript
ignoreSet = loadIgnoreAdvertisers()  // From "Advertisers to ignore" sheet
ownerMap = loadOwnerMapFromNetworks_()  // From "Networks" sheet
vMap = loadViolationChangeMap_()  // From cache
```

**🚨 CRITICAL CHECK #4: Networks Sheet**
- **REQUIRED** for owner assignment
- Format: Network ID | Advertiser | Owner
- If missing: Violations will have empty Owner column

**Chunked Processing:**
```javascript
QA_CHUNK_ROWS = 3500  // Process 3500 rows at a time
QA_TIME_BUDGET_MS = 4.2 * 60 * 1000  // 4.2 minutes max

totalRows = rawData.length  // e.g., 10,500 rows
chunk1: rows 0-3499
chunk2: rows 3500-6999
chunk3: rows 7000-10499
```

**For Each Row:**
```javascript
1. Extract data: advertiser, campaign, placement, dates, impressions, clicks, costStructure

2. Check if advertiser in ignoreSet → Skip if true

3. Validate Cost Structure:
   if (costStructure === "CPC" && clicks === 0):
     violation = "CPC placement has 0 clicks"
     
   if (costStructure === "CPM" && impressions === 0):
     violation = "CPM placement has 0 impressions"

4. Check Stale Placements:
   if (endDate < today - 7 days && (impressions > 0 || clicks > 0)):
     violation = "Stale: Spend after end date"

5. Calculate Costs:
   cpcCost = clicks * 0.008  // $8 per 1000 clicks
   cpmCost = (impressions / 1000) * 0.034  // $34 per 1000 impressions

6. Assign Owner:
   owner = ownerMap[networkId] || advertiserName

7. Check Low Priority:
   if (same violation yesterday with same stats):
     lowPriority = "Yes"

8. Build violation row:
   [network, advertiser, campaign, placement, placementId,
    startDate, endDate, costStructure, issueType, owner,
    severity, $atRisk, lowPriority, reportDate, notes]
```

**Write to Violations Sheet:**
```javascript
violationsSheet.getRange(outputRow, 1, violationRows.length, 15).setValues(violationRows)
```

**Session Tracking:**
```javascript
// If time running low or rows remaining:
saveQAState_({
  session: sessionId,
  rowStart: 3500,
  rowsProcessed: 3500,
  totalRows: 10500
})
scheduleNextQAChunk_(1)  // Continue in 1 minute
```

##### Step 2.3.5: Save Violations Report to Drive
```javascript
saveViolationsReportToDrive_(dateStr, violationCount)  [Line 6719]
├── Check if Violations sheet has data (>1 row)
├── Get root folder: '1F53lLe3z5cup338IRY4nhTZQdUmJ9_wk'
├── Navigate/Create: Root → Violations Reports folder
├── Get/Create month folder: "2025-04" (YYYY-MM format)
├── Create XLSX from Violations sheet:
│   └── Filename: "Violations_2025-04-23.xlsx"
├── Delete existing file if present
├── Upload new file to Drive
└── Return {success: true, filename, folderPath, fileUrl}
```

**Drive Structure:**
```
Root (1F53lLe3z5cup338IRY4nhTZQdUmJ9_wk)
└── Violations Reports/
    ├── 2025-04/
    │   ├── Violations_2025-04-23.xlsx  ← NEW FILE
    │   ├── Violations_2025-04-24.xlsx
    │   └── ...
    ├── 2025-05/
    └── ...
```

##### Step 2.3.6: Send Email Summary (Conditional)
```javascript
sendEmailSummary()  [Line 1296]
├── Check: Is today >= 15th of month?
│   └── If No: Log "before 15th" and SKIP
├── Check: Is QA still running? (session active)
│   └── If Yes: Log "QA in progress" and SKIP
├── Get recipients from "EMAIL LIST" sheet
├── Build HTML email with violations grouped by owner
├── Apply filters: Low Priority = No, max 30 per owner
└── Send to all recipients
```

**🚨 CRITICAL CHECK #5: Email Timing**
- Only sends on/after 15th of month
- Gap fill runs any time, but email may not send

##### Step 2.3.7: Return Result
```javascript
return {
  success: true,
  filename: "Violations_2025-04-23.xlsx",
  fileUrl: "https://drive.google.com/file/d/..."
  violationCount: 45,
  folderPath: "Violations Reports/2025-04"
}
```

#### Step 2.4: Update Progress & State
```javascript
if (tmResult.success):
  updateGapFillProgress_(dateStr, '✅ Complete', '', tmResult.filename)
  state.completedFiles.push({
    date: dateStr,
    filename: tmResult.filename,
    violationCount: 45,
    folderPath: "Violations Reports/2025-04",
    fileUrl: tmResult.fileUrl
  })
  state.successful++
else:
  updateGapFillProgress_(dateStr, '❌ Failed', tmResult.error, '')
  state.failed++

state.processed++
state.queue.shift()  // Remove first date from queue
saveGapFillState_(state)  // Save to ScriptProperties
```

**Gap Fill Progress Updated:**
| Date | Status | Last Updated | Attempts | Error Message | Drive File |
|------|--------|--------------|----------|---------------|------------|
| Wed Apr 23 2025 | ✅ Complete | 12/18/2025, 6:23:15 PM | 1 | | Violations_2025-04-23.xlsx |

#### Step 2.5: Wait for Next Trigger
```
Next execution: 15 minutes later (6:33 PM)
Processes: Thu Apr 24 2025
```

---

### **PHASE 3: Audit Refresh (Every 10 minutes)**

**Trigger Function:** `smartRefreshAudit()` [Line 11610]

```javascript
setupAndRefreshViolationsAudit()
├── Scan Drive for newly saved reports
├── Update "Violations Audit" sheet
├── Change ❌ MISSING → ✅ FOUND for completed dates
└── Update report date and folder path
```

**Purpose:**
- Keep Violations Audit current
- Show progress visually
- Verify files actually saved to Drive

**Timeline:**
```
6:18 PM - Start automation (60 missing)
6:23 PM - Complete 1st date, save to Drive
6:28 PM - Refresh audit, show 59 missing
6:33 PM - Complete 2nd date
6:38 PM - Refresh audit, show 58 missing
...continues every 10-15 minutes
```

---

### **PHASE 4: Completion (After Last Date)**

**When:** `state.queue.length === 0`

#### Step 4.1: Stop Automation
```javascript
stopSmartGapFillAutomation()  [Line 11490]
├── Delete process trigger (15 min)
├── Delete refresh trigger (10 min)
├── Remove trigger IDs from ScriptProperties
└── Log final stats
```

#### Step 4.2: Send Completion Email
```javascript
sendCompletionNotification_()  [Line 11628]
├── Load state from ScriptProperties
├── Build HTML email:
│   ├── Total Processed: 60
│   ├── Successful: 59
│   ├── Failed: 1
│   ├── Duration: "15 hours 23 minutes"
│   └── Table of all created files with clickable links
├── Send to default email
└── Toast notification: "Automation Complete"
```

---

## ⚠️ CRITICAL CHECKS SUMMARY

### **1. Gmail Labels**
**Issue:** No emails found
**Check:**
```
1. Open Gmail
2. Search: label:"CM360 QA"
3. Verify emails exist with this label
4. If not, apply label to CM360 emails
```

### **2. Google Sheets Required**
**Auto-Created:**
- ✅ Raw Data
- ✅ Violations
- ✅ Gap Fill Progress
- ✅ Violations Audit

**User Must Create:**
- ❌ Networks (Network ID | Advertiser | Owner)
- ❌ EMAIL LIST (Column A: email addresses)
- ❌ Advertisers to ignore (optional)

### **3. Drive Folder Access**
**Root Folder ID:** `1qA77_YET8RLiES7X7NoUT5jzTHDJ3k61`
**Check:**
```
1. Open: https://drive.google.com/drive/folders/1qA77_YET8RLiES7X7NoUT5jzTHDJ3k61
2. Verify you have edit access
3. Check folder structure: YYYY/MM-MonthName/YYYY-MM-DD/
```

### **4. Execution Time Per Date**
**Expected:** 3-6 minutes per date
**Breakdown:**
- Import data: 30-90 seconds
- Run QA: 120-250 seconds (chunked if needed)
- Save to Drive: 10-20 seconds
- Email: 10-30 seconds (if post-15th)

**With 60 dates:**
- Total time: ~15-18 hours
- Triggers: 60 × 15 min = 900 minutes = 15 hours minimum

### **5. Trigger Limits**
**Google Apps Script Quotas:**
- Max 20 triggers per script
- This uses 2 triggers (safe)
- Max 6 minute execution time per trigger
- Time Machine designed to finish in 4-5 minutes

### **6. State Management**
**If Script Stops:**
- State saved in ScriptProperties
- Gap Fill Progress sheet shows last status
- Triggers remain active
- Will resume on next trigger (15 min)

**Manual Restart:**
```
CM360 QA Tools → Violations Gap Fill → Stop Smart Automation
CM360 QA Tools → Violations Gap Fill → Start Smart Automation
```

---

## 🐛 Common Issues & Solutions

### **Issue 1: "No emails found for date"**
**Cause:** Gmail search returned 0 results
**Solutions:**
1. Check Gmail label exists: "CM360 QA"
2. Verify date has emails (before 4/14/25 has no data)
3. Check Drive folder exists: `YYYY/MM-MonthName/YYYY-MM-DD/`

### **Issue 2: "Raw Data sheet not found"**
**Cause:** Sheet doesn't exist
**Solution:** Auto-created now, but verify sheet tabs visible

### **Issue 3: "Exceeded maximum execution time"**
**Cause:** Processing >6 minutes
**Solution:** Chunked QA automatically schedules continuation

### **Issue 4: Empty Owner column in Violations**
**Cause:** Networks sheet missing or incorrectly formatted
**Solution:**
```
1. Create "Networks" sheet
2. Columns: Network ID | Advertiser | Owner
3. Add all network mappings
```

### **Issue 5: No completion email**
**Cause:** EMAIL LIST sheet missing
**Solution:**
```
1. Create "EMAIL LIST" sheet
2. Add email addresses in column A (starting row 2)
```

---

## 📊 Performance Metrics

### **Expected Timeline for 60 Dates**

| Time | Event | Dates Remaining |
|------|-------|-----------------|
| 6:18 PM | Start automation | 60 |
| 6:23 PM | Complete date 1 | 59 |
| 6:38 PM | Complete date 2 | 58 |
| 6:53 PM | Complete date 3 | 57 |
| ... | ... | ... |
| 9:18 AM (next day) | Complete date 60 | 0 |

**Total Duration:** ~15-18 hours

### **Resource Usage**
- Script execution time: ~4-5 min per date
- Total execution time: ~300 minutes (5 hours)
- Idle time: ~15 hours (waiting between triggers)
- Drive storage: ~60 × 50KB = 3MB (XLSX files)
- Gmail API calls: ~120 searches (2 per date)

---

## ✅ Pre-Flight Checklist

**Before Starting Automation:**
- [ ] Violations Audit shows ❌ MISSING dates
- [ ] Networks sheet populated with owner mappings
- [ ] EMAIL LIST sheet has recipient emails
- [ ] Gmail label "CM360 QA" applied to emails
- [ ] Drive folder accessible: `1qA77_YET8RLiES7X7NoUT5jzTHDJ3k61`
- [ ] No other automation triggers running
- [ ] Ready to leave running for 15-18 hours

---

**Last Updated:** December 18, 2025  
**Status:** Production Ready ✅
