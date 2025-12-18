# CM360 End of Month Audit System - `runItAll()` Reference Documentation

## 📋 Overview

The **CM360 End of Month Audit System** is a Google Apps Script automation that processes daily Campaign Manager 360 (CM360) reports to identify advertising placement violations. The core orchestration function `runItAll()` coordinates the entire daily workflow.

---

## 🎯 Primary Purpose

**Automated Daily QA Processing**
- Import raw CM360 placement data from Gmail
- Analyze placements for violations (cost structure mismatches, stale placements, etc.)
- Generate violation reports with owner assignments
- Send alerts and email summaries to stakeholders

---

## 🏗️ System Architecture

### Core Components

```
runItAll() [Master Orchestrator]
├── trimAllSheetsToData_()          → Clean up sheet formatting
├── importDCMReports()              → Download & import CSVs from Gmail
├── runQAOnly()                     → Analyze data & generate violations
├── sendPerformanceSpikeAlertIfPre15() → Pre-15th performance alerts
└── sendEmailSummary()              → Post-15th violation summary emails
```

### Key Google Sheets

| Sheet Name | Purpose | Auto-Created |
|------------|---------|--------------|
| **Raw Data** | Imported CM360 placement data | ✅ Yes |
| **Violations** | QA analysis results | ✅ Yes |
| **Networks** | Network ID → Owner mapping | ❌ Required |
| **EMAIL LIST** | Email recipients | ❌ Optional |
| **Advertisers to ignore** | Filter list | ❌ Optional |
| **Gap Fill Progress** | Automation tracking | ✅ Yes |
| **Violations Audit** | Historical reports dashboard | ✅ Yes |

---

## 🔄 `runItAll()` Execution Flow

### **Phase 1: Data Preparation** (30-60 seconds)

**Step 1: `trimAllSheetsToData_()`**
- Removes blank rows/columns from all sheets
- Ensures consistent data structure
- Prevents formatting drift over time

**Step 2: `importDCMReports()`**
- Searches Gmail for emails with label: `"CM360 QA"`
- Filters emails to today's date only: `after:YYYY/MM/DD`
- Processes CSV/ZIP attachments:
  - Extracts Network ID from filename (digits before first `_`)
  - Parses CSV data (Advertiser, Campaign, Placement, etc.)
  - Appends report date column
- Writes to **Raw Data** sheet with headers:
  ```
  Network ID | Advertiser | Placement ID | Placement | Campaign | 
  Placement Start Date | Placement End Date | Campaign Start Date | 
  Campaign End Date | Ad | Impressions | Clicks | Report Date
  ```

### **Phase 2: Time Budget Check** (1 second)

**Critical Decision Point:**
```javascript
totalMs = Date.now() - runStart;
quotaMs = 6 * 60 * 1000;  // 6 minutes
timeLeft = quotaMs - totalMs;

if (timeLeft < 2 * 60 * 1000) {  // Less than 2 minutes left?
  // Schedule QA for later trigger, exit now
  scheduleNextQAChunk_(1);
  return;
}
```

**Why this matters:**
- Google Apps Script has a 6-minute execution limit
- QA analysis can take 3-5 minutes for large datasets
- If <2 min remaining, schedule QA as a separate trigger to avoid timeout

### **Phase 3: QA Analysis** (2-5 minutes)

**Step 3: `runQAOnly()` - Chunked Processing**

**Input:** Raw Data sheet (3,500+ rows typical)
**Output:** Violations sheet

**QA Checks Performed:**

1. **Cost Structure Validation**
   - Placement marked as "CPC" but has 0 clicks → Flag as "CPC placement has 0 clicks"
   - Placement marked as "CPM" but has 0 impressions → Flag as "CPM placement has 0 impressions"

2. **Stale Placement Detection**
   - End date >7 days ago but still showing impressions/clicks
   - Flags as "Stale: Spend after end date"

3. **Flight Completion Analysis**
   - Calculates % of flight completed based on days elapsed
   - Calculates % of month remaining
   - Identifies performance spikes or lulls

4. **Cost Calculation**
   - CPC Rate: $0.008 per click ($8 per 1,000 clicks)
   - CPM Rate: $0.034 per 1,000 impressions
   - Calculates "$ at Risk" for each violation

5. **Owner Assignment**
   - Maps Network ID → Owner via **Networks** sheet
   - Falls back to Advertiser name if Network mapping unavailable

6. **Low-Priority Classification**
   - Checks violation change cache
   - If same placement violated yesterday with same stats → Mark "Low Priority"

**Chunked Execution:**
```
Total Rows: 10,500
Chunk Size: 3,500 rows
Processing: Chunk 1 → Schedule Chunk 2 → Schedule Chunk 3
Max Time: 4.2 minutes per chunk
```

**Violations Sheet Columns:**
```
Network | Advertiser | Campaign | Placement | Placement ID | 
Start Date | End Date | Cost Structure | Issue Type | Owner | 
Severity | $ at Risk | Low Priority | Report Date | Notes
```

### **Phase 4: Alerting & Reporting** (30-90 seconds)

**Step 4: `sendPerformanceSpikeAlertIfPre15()`**
- **Only runs before 15th of month**
- Monitors for unusual activity spikes
- Sends immediate alert emails to stakeholders

**Step 5: `sendEmailSummary()`**
- **Only runs on/after 15th of month**
- Generates violation summary grouped by owner
- Includes:
  - Total violations per owner
  - $ at Risk breakdown
  - Clickable Drive links to detailed reports
  - Filters out "Low Priority" violations by default
- Max email size: 90KB HTML (auto-truncates if needed)
- Sends to all emails in **EMAIL LIST** sheet

---

## ⏰ Trigger Variants

### **`runItAll()` - Full Daily Run**
**When:** Manual execution or custom trigger
**Duration:** ~5-7 minutes
**Includes:** Import → QA → Performance Alert → Email Summary

### **`runItAllMorning()` - Morning Automation**
**When:** Time-driven trigger (e.g., 8:00 AM daily)
**Duration:** ~5 minutes
**Includes:** Import → QA → Performance Alert
**Excludes:** Email Summary (runs separately)

### **`runDailyEmailSummary()` - Email Only**
**When:** Time-driven trigger (e.g., 12:00 PM daily)
**Duration:** ~1-2 minutes
**Includes:** Email Summary only
**Why separate:** Prevents email timeout when QA takes full 6 minutes

---

## 📊 Data Flow Diagram

```
┌─────────────────────────────────────────────────────────┐
│  Gmail (CM360 QA Label)                                 │
│  Subject: "BKCM360 Global QA Check"                     │
│  Attachments: 1068_*.csv, 2524_*.zip, etc.             │
└───────────────────┬─────────────────────────────────────┘
                    │
                    ↓ importDCMReports()
┌─────────────────────────────────────────────────────────┐
│  Raw Data Sheet (Google Sheets)                         │
│  Network ID | Advertiser | Campaign | Placement |...   │
│  1068       | Acme Corp  | Summer   | Banner 728x90    │
│  2524       | XYZ Inc    | Fall     | Video Pre-roll   │
└───────────────────┬─────────────────────────────────────┘
                    │
                    ↓ runQAOnly()
┌─────────────────────────────────────────────────────────┐
│  QA Analysis Engine                                      │
│  ├─ Cost structure validation                           │
│  ├─ Stale placement detection                           │
│  ├─ Owner mapping (Networks sheet)                      │
│  ├─ $ at Risk calculation                               │
│  └─ Low-priority flagging (change cache)                │
└───────────────────┬─────────────────────────────────────┘
                    │
                    ↓
┌─────────────────────────────────────────────────────────┐
│  Violations Sheet (Google Sheets)                       │
│  Issue Type              | Owner  | $ at Risk | Priority│
│  CPC placement has 0 clicks | John   | $245.00  | Normal │
│  Stale: Spend after end   | Sarah  | $892.00  | Low    │
└───────────────────┬─────────────────────────────────────┘
                    │
                    ↓ sendEmailSummary()
┌─────────────────────────────────────────────────────────┐
│  Email Recipients (EMAIL LIST sheet)                    │
│  john@example.com                                        │
│  sarah@example.com                                       │
│  team@example.com                                        │
└─────────────────────────────────────────────────────────┘
```

---

## 🔧 Configuration Constants

```javascript
// Execution Time Management
const APPROX_QUOTA_MINUTES = 6;        // Google Apps Script limit
const QA_TIME_BUDGET_MS = 4.2 * 60 * 1000;  // 4.2 min for QA
const QA_CHUNK_ROWS = 3500;            // Process 3,500 rows per chunk

// Cost Calculation Rates
const CPC_RATE = 0.008;                // $8 per 1,000 clicks
const CPM_RATE = 0.034;                // $34 per 1,000 impressions

// Email Size Limits
const MAX_HTML_CHARS = 90000;          // 90KB max email size
const MAX_ROWS_PER_OWNER = 30;         // Max violations per owner in email
const MAX_TOTAL_OWNER_ROWS = 1000;     // Max total violations in email

// Stale Threshold
const DEFAULT_STALE_DAYS = 7;          // Configurable via Networks!H1
```

---

## 🚨 Error Handling & Resilience

### **Automatic Sheet Creation**
If critical sheets don't exist, they're auto-created:
```javascript
if (!rawSheet) {
  rawSheet = ss.insertSheet("Raw Data");
  rawSheet.getRange("A1:H1").setValues([headers]);
}
```

### **Document Lock Protection**
Prevents concurrent QA runs:
```javascript
const dlock = LockService.getDocumentLock();
if (!dlock.tryLock(30000)) {
  scheduleNextQAChunk_(2);  // Retry in 2 minutes
  return;
}
```

### **Time Budget Monitoring**
Every step logs execution time:
```javascript
logStep_('importDCMReports', function(){ importDCMReports(); }, runStart, APPROX_QUOTA_MINUTES);
// Logs: "✅ importDCMReports (1.2s) [0.3% quota]"
```

### **Chunked QA Processing**
If QA times out mid-execution:
```javascript
// Save progress to DocumentProperties
saveQAState_({
  session: sessionId,
  rowStart: 3501,
  rowsProcessed: 3500,
  totalRows: 10500
});

// Schedule next chunk
scheduleNextQAChunk_(1);  // Resume in 1 minute
```

---

## 📈 Performance Metrics

### **Typical Execution Times**
| Step | Duration | % of Budget |
|------|----------|-------------|
| trimAllSheetsToData_ | 10-30s | 3-8% |
| importDCMReports | 30-90s | 8-25% |
| runQAOnly (per chunk) | 120-250s | 33-69% |
| sendPerformanceSpikeAlert | 10-20s | 3-6% |
| sendEmailSummary | 30-90s | 8-25% |

### **Data Volume Handled**
- Raw Data: 5,000-15,000 rows daily
- Violations: 50-500 violations typical
- Networks: 10-50 network mappings
- Email Recipients: 5-20 users

---

## 🎛️ Automation Trigger Setup

### **Recommended Daily Schedule**

**Morning Run (Data Import + QA):**
```
Trigger: Time-driven, Day timer, 8:00 AM - 9:00 AM
Function: runItAllMorning
Frequency: Every day
Failure notifications: Notify me immediately
```

**Afternoon Email (Violation Summary):**
```
Trigger: Time-driven, Day timer, 12:00 PM - 1:00 PM
Function: runDailyEmailSummary
Frequency: Every day
Failure notifications: Notify me immediately
```

---

## 🔍 Troubleshooting Guide

### **"Exceeded maximum execution time"**
**Cause:** QA processing >6 minutes
**Solution:** Automatic - chunked execution resumes via trigger

### **"Raw Data sheet not found"**
**Cause:** Sheet deleted manually
**Solution:** Auto-created on next run with proper headers

### **"No violations found"**
**Possible causes:**
1. Gmail label "CM360 QA" not applied to emails
2. No emails for today's date
3. All placements are compliant (rare!)

### **"Owner (Ops) column empty"**
**Cause:** Networks sheet missing or incomplete
**Solution:** Populate Networks sheet with:
```
Network ID | Advertiser | Owner
1068       | Acme Corp  | John Smith
2524       | XYZ Inc    | Sarah Jones
```

---

## 📝 Key Differences: `runItAll()` vs Automation Variants

| Feature | runItAll() | runItAllMorning() | runDailyEmailSummary() |
|---------|-----------|-------------------|------------------------|
| Data Import | ✅ Yes | ✅ Yes | ❌ No |
| QA Analysis | ✅ Yes | ✅ Yes | ❌ No |
| Performance Alert | ✅ Yes | ✅ Yes | ❌ No |
| Email Summary | ✅ Yes | ❌ No | ✅ Yes |
| Use Case | Manual/On-demand | Daily automation | Separate email trigger |

---

## 🔗 Related Systems

### **Smart Gap Fill Automation**
- Function: `startSmartGapFillAutomation()`
- Purpose: Backfill historical missing violation reports (April-December 2025)
- Trigger: 15-minute intervals
- Details: See [Gap Fill Progress] sheet

### **Raw Data Archive**
- Function: `archiveAllRawData()`
- Purpose: Save all historical raw CSV files to Google Drive
- Folder: `Raw Data/YYYY/MM-Month/YYYY-MM-DD/`
- Status: One-time backfill completed

### **Violations Audit Dashboard**
- Function: `setupAndRefreshViolationsAudit()`
- Purpose: Scan Drive for existing violation reports
- Shows: ✅ FOUND / ❌ MISSING for each date (15th-31st of month)

---

## 📚 Data Sources

### **Gmail Search Pattern**
```
Label: "CM360 QA"
Date: Today (after:YYYY/MM/DD)
Subject: "BKCM360 Global QA Check"
Attachments: CSV/ZIP files
```

### **CSV Filename Pattern**
```
{NetworkID}_BKCM360_Global_QA_Check_{YYYYMMDD}_{HHMMSS}_{ReportID}.csv

Example:
1068_BKCM360_Global_QA_Check_20250423_010611_5077781354.csv
     ^^^^                         ^^^^^^^^
     Network ID                   Date: 2025-04-23
```

---

## 🎯 Success Criteria

**Daily Run Considered Successful If:**
1. ✅ Raw Data sheet populated (>100 rows)
2. ✅ QA analysis completed (no hanging sessions)
3. ✅ Violations sheet generated (0+ violations is valid)
4. ✅ Email sent (if post-15th) or alert sent (if pre-15th)
5. ✅ Execution completed <6 minutes (or chunked properly)

---

## 📊 Monthly Workflow

### **Days 1-14: Performance Monitoring**
- `runItAll()` runs daily
- `sendPerformanceSpikeAlertIfPre15()` active
- Monitors for unusual activity spikes
- No violation emails sent

### **Days 15-31: Violation Reporting**
- `runItAll()` runs daily
- `sendEmailSummary()` active
- Daily violation emails to stakeholders
- Includes $ at Risk calculations
- Filters low-priority repeat violations

---

## 🛠️ Maintenance Notes

### **Monthly Tasks**
- Review Networks sheet for new/removed advertisers
- Update EMAIL LIST for recipient changes
- Check Gap Fill Progress for any failed historical dates

### **Quarterly Review**
- Analyze violation trends via Violations Audit
- Adjust cost rates (CPC_RATE, CPM_RATE) if needed
- Review stale threshold days (Networks!H1)

### **Annual Cleanup**
- Archive old violation reports to Drive
- Clean up Gmail "CM360 QA" label (keep last 90 days)
- Verify Drive folder structure integrity

---

## 🔒 Security & Permissions

**Required Google OAuth Scopes:**
- `https://www.googleapis.com/auth/gmail.readonly` - Read Gmail
- `https://www.googleapis.com/auth/spreadsheets` - Read/Write Sheets
- `https://www.googleapis.com/auth/drive` - Save reports to Drive
- `https://www.googleapis.com/auth/script.send_mail` - Send emails

**Document-Level Access:**
- Script must be bound to the Google Spreadsheet
- Users need "Editor" access to run functions manually

---

## 📞 Support & Documentation

**Primary Files:**
- `Code.gs` - Main script (12,000+ lines)
- `AuditSystems.gs` - Audit dashboard functions
- `RUNALL-REFERENCE.md` - This document

**Key Functions Reference:**
- `runItAll()` - Line 1758
- `importDCMReports()` - Line 238
- `runQAOnly()` - Line 1058
- `sendEmailSummary()` - Line 1296

---

**Last Updated:** December 18, 2025  
**Version:** 2.0 (Smart Automation + Drive Integration)  
**Status:** Production Ready ✅
