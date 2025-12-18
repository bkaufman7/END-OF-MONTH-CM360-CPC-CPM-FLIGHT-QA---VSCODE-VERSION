// =====================
// CM360 QA Tools Script
// =====================
// Adds custom menu, imports CM360 reports via Gmail, runs QA checks,
// filters out ignored advertisers, and emails a summary of violations.

// ---------------------
// onOpen: Menu Setup
// ---------------------
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("CM360 QA Tools")
    // === DAILY OPERATIONS ===
    .addItem("▶️ Run It All", "runItAll")
    .addItem("📥 Pull Data", "importDCMReports")
    .addItem("🔍 Run QA Only", "runQAOnly")
    .addItem("📧 Send Email Only", "sendEmailSummary")
    .addSeparator()
    
    // === ARCHIVE AUDITS ===
    .addSubMenu(ui.createMenu("📋 Archive Audits")
      .addItem("📊 Raw Data Audit (Check Drive)", "setupAndRefreshRawDataAudit")
      .addItem("📦 Violations Audit (Gmail + Drive)", "setupAndRefreshViolationsAudit"))
    .addSeparator()
    
    // === RAW DATA GAP FILL ===
    .addSubMenu(ui.createMenu("🔄 Raw Data Gap Fill")
      .addItem("▶️ Start Raw Data Gap Fill", "startRawDataGapFill")
      .addItem("📊 View Status", "viewRawDataGapFillStatus")
      .addSeparator()
      .addItem("⏰ Create Auto-Resume Trigger (10 min)", "createRawGapFillAutoResumeTrigger")
      .addItem("🛑 Stop & Delete Trigger", "stopRawDataGapFillAndDeleteTrigger")
      .addSeparator()
      .addItem("🤖 Start Smart Automation (15min process + 10min refresh)", "startSmartRawDataAutomation")
      .addItem("🛑 Stop Smart Automation", "stopSmartRawDataAutomation")
      .addSeparator()
      .addItem("🔄 Reset (Start Over)", "resetRawDataGapFill"))
    .addSeparator()
    
    // === RAW DATA GAP FILL (TEST MODE - 2 PHASE) ===
    .addSubMenu(ui.createMenu("🧪 Raw Data Gap Fill (TEST)")
      .addItem("▶️ Phase 1: Download All Attachments", "startTestPhase1Download")
      .addItem("📦 Phase 2: Extract All ZIPs", "startTestPhase2Extraction")
      .addSeparator()
      .addItem("📊 View Download Progress (Phase 1)", "viewTestPhase1Status")
      .addItem("📊 View Extraction Progress (Phase 2)", "viewTestPhase2Status")
      .addItem("📈 View Today's Progress", "viewTodayProgress")
      .addSeparator()
      .addItem("⏰ Create Phase 1 Auto-Resume", "createTestPhase1Trigger")
      .addItem("⏰ Create Phase 2 Auto-Resume", "createTestPhase2Trigger")
      .addSeparator()
      .addItem("🤖 Start Complete TEST Automation", "startCompleteTestAutomation")
      .addItem("🛑 Stop All TEST Automation", "stopCompleteTestAutomation")
      .addSeparator()
      .addItem("🌅 Setup Daily Morning Automation (7-8 AM)", "setupDailyMorningAutomation")
      .addItem("📧 Create Daily Email Trigger (7:30 PM)", "createDailyEmailTrigger")
      .addItem("📅 Setup Weekly Auto-Download (Sat 11:30 PM)", "setupWeeklyAutoDownload")
      .addSeparator()
      .addItem("🛑 Stop All TEST Triggers", "stopAllTestTriggers")
      .addSeparator()
      .addItem("🔍 Audit Test Folder", "auditTestFolder")
      .addItem("🔧 Fix Incomplete Dates (Auto)", "fixIncompleteDatesAuto")
      .addItem("🧹 Cleanup & Verify", "cleanupAndVerifyTest")
      .addItem("🔄 Reset TEST Mode", "resetTestMode"))
    .addSeparator()
    
    // === VIOLATIONS GAP FILL ===
    .addSubMenu(ui.createMenu("🎯 Violations Gap Fill")
      .addItem("💾 Setup Progress Sheet", "setupGapFillProgressSheet")
      .addItem("▶️ Start Auto Gap Fill", "startAutoGapFill")
      .addItem("📊 View Status", "viewGapFillStatus")
      .addSeparator()
      .addItem("⏰ Create Auto-Resume Trigger (10 min)", "createGapFillAutoResumeTrigger")
      .addItem("🛑 Stop & Delete Trigger", "stopGapFillAndDeleteTrigger")
      .addSeparator()
      .addItem("🤖 Start Smart Automation (15min process + 10min refresh)", "startSmartGapFillAutomation")
      .addItem("🛑 Stop Smart Automation", "stopSmartGapFillAutomation")
      .addSeparator()
      .addItem("🔄 Reset (Start Over)", "resetGapFill"))
    .addSeparator()
    
    // === TIME MACHINE ===
    .addSubMenu(ui.createMenu("⏰ Time Machine")
      .addItem("⚙️ Setup Time Machine", "setupTimeMachineSheet")
      .addItem("▶️ Run QA for Selected Date", "runTimeMachineQA"))
    .addSeparator()
    
    // === REPORTS & DASHBOARDS ===
    .addSubMenu(ui.createMenu("📈 Reports & Dashboards")
      .addItem("📊 Generate V2 Dashboard", "generateViolationsV2Dashboard")
      .addItem("💾 Export V2 to Drive", "exportV2ToDrive")
      .addItem("📋 Monthly Summary Report", "generateMonthlySummaryReport")
      .addItem("📊 Month-over-Month Analysis", "runMonthOverMonthAnalysis")
      .addItem("💰 Calculate Financial Impact", "displayFinancialImpact"))
    .addSeparator()
    
    // === HISTORICAL ARCHIVE ===
    .addSubMenu(ui.createMenu("📁 Historical Archive")
      .addItem("📦 Archive All (April-Nov 2025)", "archiveAllHistoricalReports")
      .addItem("📅 Archive Single Month", "archiveSingleMonth")
      .addItem("📊 View Archive Progress", "viewArchiveProgress")
      .addItem("▶️ Resume Archive", "resumeArchive"))
    .addSeparator()
    .addSubMenu(ui.createMenu("📂 Raw Data Archive")
      .addItem("📦 Archive All Raw Data (Apr-Nov 2025)", "archiveAllRawData")
      .addItem("📊 View Raw Data Progress", "viewRawDataProgress")
      .addItem("📧 Email Detailed Progress Report", "emailDetailedProgressReport")
      .addItem("▶️ Resume Raw Data Archive", "resumeRawDataArchive")
      .addSeparator()
      .addItem("⏰ Create Auto-Resume Trigger", "createRawDataAutoResumeTrigger")
      .addItem("🛑 Delete Auto-Resume Trigger", "deleteRawDataAutoResumeTrigger")
      .addSeparator()
      .addItem("📧 Create Daily Progress Report (7:30 PM)", "createDailyProgressReportTrigger")
      .addItem("🛑 Delete Daily Progress Report", "deleteDailyProgressReportTrigger")
      .addSeparator()
      .addItem("📂 Categorize Files by Network", "categorizeRawDataByNetwork")
      .addItem("🔬 Audit Archive Completeness (Quick)", "auditRawDataArchive")
      .addItem("🔬 Comprehensive Audit (Gmail vs Drive)", "auditRawDataArchiveComprehensive")
      .addSeparator()
      .addItem("▶️ Resume Comprehensive Audit", "processComprehensiveAuditBatch_")
      .addItem("📊 View Audit Progress", "viewComprehensiveAuditProgress")
      .addItem("🔄 Reset Comprehensive Audit", "resetComprehensiveAudit"))
    .addSeparator()
    
    // === UTILITIES ===
    .addSubMenu(ui.createMenu("⚙️ Settings & Utilities")
      .addItem("🔓 Authorize Email (one-time)", "authorizeMail_")
      .addItem("🕒 Create Daily Email Trigger (9am)", "createDailyEmailTrigger")
      .addSeparator()
      .addItem("🧹 Clear Violations", "clearViolations"))
    .addToUi();
  
  // Setup Time Machine sheet if it exists
  setupTimeMachineIfExists_();
}

/**
 * Show menu loading error details
 */
function showMenuError_() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '`u{26A0}`u{FE0F} Menu Loading Error',
    'The CM360 QA Tools menu failed to load.\n\n' +
    'Check the Apps Script execution log (View > Executions) for details.\n\n' +
    'Common causes:\n' +
    '`u{2022} Missing or renamed functions\n' +
    '`u{2022} Syntax errors in the script\n' +
    '`u{2022} Authorization issues',
    ui.ButtonSet.OK
  );
}
// ---------------------
function authorizeMail_() {
  // Running this from the editor or from the menu will force the OAuth prompt
  MailApp.sendEmail({
    to: 'platformsolutionsadopshorizon@gmail.com',
    subject: 'Apps Script auth test',
    htmlBody: 'If you received this, MailApp is authorized.'
  });
}

// ---------------------
// Create an installable time trigger for the email-only run
// ---------------------
function createDailyEmailTrigger() {
  // Runs runDailyEmailSummary daily at 9am local time with full auth
  ScriptApp.newTrigger('runDailyEmailSummary')
    .timeBased()
    .atHour(9)       // change if you prefer another hour
    .everyDays(1)
    .create();
}




// ---------------------
// clearViolations
// ---------------------
function clearViolations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Violations");
  
  // Auto-create Violations sheet if it doesn't exist
  if (!sheet) {
    Logger.log('⚠️ Violations sheet not found - creating it now...');
    sheet = ss.insertSheet("Violations");
    sheet.getRange("A1:O1").setValues([[
      "Network", "Advertiser", "Campaign", "Placement", "Placement ID",
      "Start Date", "End Date", "Cost Structure", "Issue Type", 
      "Owner", "Severity", "$ at Risk", "Low Priority", "Report Date", "Notes"
    ]]).setFontWeight("bold");
    Logger.log('✅ Violations sheet created');
    return; // Nothing to clear on new sheet
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

// ---------------------
// clearRawData
// ---------------------
function clearRawData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Data");
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

// ---------------------
// extractNetworkId
// ---------------------
function extractNetworkId(fileName) {
  const match = fileName.match(/^([^_]+)_/);
  return match ? String(match[1]) : "Unknown";
}

// ---------------------
// processCSV
// ---------------------
function processCSV(fileContent, networkId) {
  const lines = fileContent.split("\n").map(line => line.trim()).filter(Boolean);
  const startIndex = lines.findIndex(line => line.startsWith("Advertiser"));
  if (startIndex === -1) return [];
  const csvData = Utilities.parseCsv(lines.slice(startIndex).join("\n"));
  csvData.shift(); // remove header row in the attachment
  const reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  return csvData.map(function(row){ return [networkId].concat(row).concat([reportDate]); });
}

function importDCMReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Raw Data") || ss.insertSheet("Raw Data");
  const outputSheet = ss.getSheetByName("Violations") || ss.insertSheet("Violations");
  const label = "CM360 QA";
  const formattedToday = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");

  const dataHeaders = [
    "Network ID","Advertiser","Placement ID","Placement","Campaign",
    "Placement Start Date","Placement End Date","Campaign Start Date","Campaign End Date",
    "Ad","Impressions","Clicks","Report Date"
  ];
  // APPENDED "Owner (Ops)" to be column Y (25th)
  const outputHeaders = [
    "Network ID","Report Date","Advertiser","Campaign","Campaign Start Date","Campaign End Date",
    "Ad","Placement ID","Placement","Placement Start Date","Placement End Date",
    "Impressions","Clicks","CTR (%)","Days Until Placement End","Flight Completion %",
    "Days Left in the Month","CPC Risk","$CPC","$CPM","Issue Type","Details",
    "Last Imp Change","Last Click Change","Owner (Ops)"
  ];

  dataSheet.clearContents().getRange(1,1,1,dataHeaders.length).setValues([dataHeaders]);
  outputSheet.clearContents().getRange(1,1,1,outputHeaders.length).setValues([outputHeaders]);

  const threads = GmailApp.search('label:' + label + ' after:' + formattedToday);
  let extractedData = [];

  threads.forEach(function(thread){
    thread.getMessages().forEach(function(message){
      message.getAttachments().forEach(function(att){
        const netId = extractNetworkId(att.getName());
        if (att.getContentType() === "text/csv" || att.getName().endsWith(".csv")) {
          extractedData = extractedData.concat(processCSV(att.getDataAsString(), netId));
        } else if (att.getContentType() === "application/zip") {
          Utilities.unzip(att.copyBlob()).forEach(function(file){
            if (file.getContentType() === "text/csv" || file.getName().endsWith(".csv")) {
              extractedData = extractedData.concat(processCSV(file.getDataAsString(), extractNetworkId(file.getName())));
            }
          });
        }
      });
    });
  });

  if (extractedData.length) {
    dataSheet.getRange(2, 1, extractedData.length, dataHeaders.length).setValues(extractedData);
  }
}

// ====== Chunked QA execution control ======
const QA_CHUNK_ROWS = 3500;
const QA_TIME_BUDGET_MS = 4.2 * 60 * 1000;
const QA_STATE_KEY = 'qa_progress_v2';      // DocumentProperties key

// ====== Cost calculation rates ======
const CPC_RATE = 0.008;  // Cost per click ($8 per 1000 clicks)
const CPM_RATE = 0.034;  // Cost per 1000 impressions ($0.034 per 1000 impressions)

// --- Auto-resume trigger control for QA chunks ---
const QA_TRIGGER_KEY = 'qa_chunk_trigger_id';   // ScriptProperties key for one-shot trigger
const QA_LOCK_KEY = 'qa_chunk_lock';            // logical name only

function getScriptProps_() { return PropertiesService.getScriptProperties(); }

function scheduleNextQAChunk_(minutesFromNow) {
  minutesFromNow = Math.max(1, Math.min(10, Math.floor(minutesFromNow || 1))); // 1..10 min
  const props = getScriptProps_();

  // If a trigger is already scheduled, do nothing (unless it no longer exists)
  const existingId = props.getProperty(QA_TRIGGER_KEY);
  if (existingId) {
    const stillThere = ScriptApp.getProjectTriggers().some(function(t){ return t.getUniqueId() === existingId; });
    if (stillThere) return;
    props.deleteProperty(QA_TRIGGER_KEY);
  }

  const trig = ScriptApp
    .newTrigger('runQAOnly')      // re-enter same function
    .timeBased()
    .after(minutesFromNow * 60 * 1000)
    .create();

  props.setProperty(QA_TRIGGER_KEY, trig.getUniqueId());
}

function cancelQAChunkTrigger_() {
  const props = getScriptProps_();
  const id = props.getProperty(QA_TRIGGER_KEY);
  if (!id) return;
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getUniqueId() === id) ScriptApp.deleteTrigger(t);
  });
  props.deleteProperty(QA_TRIGGER_KEY);
}

function getQAState_() {
  const raw = PropertiesService.getDocumentProperties().getProperty(QA_STATE_KEY);
  return raw ? JSON.parse(raw) : null;
}
function saveQAState_(obj) {
  PropertiesService.getDocumentProperties().setProperty(QA_STATE_KEY, JSON.stringify(obj));
}
function clearQAState_() {
  PropertiesService.getDocumentProperties().deleteProperty(QA_STATE_KEY);
}

// ---------------------
// getHeaderMap
// ---------------------
function getHeaderMap(headers) {
  const map = {};
  headers.forEach(function(h,i){ map[String(h).trim()] = i; });
  return map;
}

// ===== Helpers for change detection cache (PERFORMANCE alert snapshots use a sheet) =====
function getPerfAlertCacheSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = "_Perf Alert Cache";

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.hideSheet();
    }

    const needed = ["date","key","impressions","clicks"];
    const current = sh.getRange(1, 1, 1, 4).getValues()[0] || [];
    const ok = current.length === 4 && current
      .map(function(v){ return String(v).toLowerCase(); })
      .every(function(v, i){ return v === needed[i]; });

    if (!ok) {
      sh.getRange(1, 1, 1, 4).setValues([needed]);
    }
    return sh;
  } finally {
    lock.releaseLock();
  }
}

// Returns a map of latest snapshot by key: { key: { date: 'yyyy-MM-dd', imp: number, clk: number } }
function loadLatestCacheMap_() {
  const sh = getPerfAlertCacheSheet_();
  const vals = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < vals.length; i++) {
    const d   = vals[i][0];
    const key = String(vals[i][1] || "");
    const imp = Number(vals[i][2] || 0);
    const clk = Number(vals[i][3] || 0);
    if (!key) continue;
    const ds = (d && d.getFullYear) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d || "");
    if (!map[key] || ds > map[key].date) {
      map[key] = { date: ds, imp: imp, clk: clk };
    }
  }
  return map;
}

// Appends today's snapshots for all evaluated rows
function appendTodaySnapshots_(rowsForSnapshot) {
  if (!rowsForSnapshot.length) return;
  const sh = getPerfAlertCacheSheet_();
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const out = rowsForSnapshot.map(function(r){ return [todayStr, r.key, r.imp, r.clk]; });
  sh.getRange(sh.getLastRow()+1, 1, out.length, 4).setValues(out);
}

// Compact PERF ALERT cache to last N days
function compactPerfAlertCache_(keepDays) {
  keepDays = keepDays || 35;
  const sh = getPerfAlertCacheSheet_();
  const cutoff = new Date(Date.now() - keepDays*86400000);
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return;

  const keep = [vals[0]];
  for (let i = 1; i < vals.length; i++) {
    const d = vals[i][0] instanceof Date ? vals[i][0] : new Date(vals[i][0]);
    if (d >= cutoff) keep.push(vals[i]);
  }
  sh.clearContents();
  sh.getRange(1,1,keep.length,4).setValues(keep);
}

// ---------------------
// Ignore Advertisers sheet
// ---------------------
function loadIgnoreAdvertisers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Advertisers to ignore");
  if (!sheet) return new Set();
  const rows = sheet.getDataRange().getValues();
  const ignoreMap = {};

  for (let i = 1; i < rows.length; i++) {
    const name = rows[i][0] && rows[i][0].toString().trim().toLowerCase();
    if (name) ignoreMap[name] = { row: i + 1, set: new Set() };
  }

  const raw = ss.getSheetByName("Raw Data");
  if (raw) {
    const data = raw.getDataRange().getValues();
    const m = getHeaderMap(data[0]);
    data.slice(1).forEach(function(r){
      const adv = r[m["Advertiser"]] && r[m["Advertiser"]].toString().trim().toLowerCase();
      const net = r[m["Network ID"]];
      if (adv && ignoreMap[adv]) ignoreMap[adv].set.add(net);
    });
    // Batch write all counts at once
    const updates = [];
    Object.values(ignoreMap).forEach(function(o){
      updates.push([o.row, o.set.size]);
    });
    if (updates.length > 0) {
      updates.sort((a, b) => a[0] - b[0]);
      const updateData = updates.map(u => [u[1]]);
      const startRow = updates[0][0];
      sheet.getRange(startRow, 2, updates.length, 1).setValues(updateData);
    }
  }

  return new Set(Object.keys(ignoreMap));
}

// ---------------------
// sendPerformanceSpikeAlertIfPre15
// ---------------------
function sendPerformanceSpikeAlertIfPre15() {
  const today = new Date();
  const dayOfMonth = today.getDate();
  if (dayOfMonth >= 15) return; // Only before 15th

  // Ensures the cache sheet exists before proceeding
  getPerfAlertCacheSheet_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Violations");
  const recipientsSheet = ss.getSheetByName("EMAIL LIST");
  if (!sheet || !recipientsSheet) return;

  // Recipient list
  const emails = recipientsSheet.getRange("A2:A").getValues()
    .flat()
    .map(function(e){ return String(e || "").trim(); })
    .filter(Boolean);
  const uniqueEmails = Array.from(new Set(emails));
  if (uniqueEmails.length === 0) return;

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return;

  const headers = values[0];
  const hMap = {};
  headers.forEach(function(h, i){ hMap[h] = i; });

  const req = [
    "Network ID", "Report Date", "Advertiser", "Campaign",
    "Placement ID", "Placement", "Impressions", "Clicks", "Issue Type", "Details"
  ];
  if (req.some(function(k){ return hMap[k] === undefined; })) return;

  const MATCH_TEXT = "� PERFORMANCE: CTR �� 90% & CPM �� $10";
  const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const latestMap = loadLatestCacheMap_();

  const candidateRows = [];
  const snapshots = [];

  values.slice(1).forEach(function(r){
    const issueStr = String(r[hMap["Issue Type"]] || "");
    if (!issueStr.includes(MATCH_TEXT)) return;

    const rd = new Date(r[hMap["Report Date"]]);
    if (isNaN(rd) || rd < startOfMonth || rd > today) return;

    const netId = String(r[hMap["Network ID"]] || "");
    const adv   = String(r[hMap["Advertiser"]] || "");
    const camp  = String(r[hMap["Campaign"]] || "");
    const pid   = String(r[hMap["Placement ID"]] || "");
    const plc   = String(r[hMap["Placement"]] || "");
    const imp   = Number(r[hMap["Impressions"]] || 0);
    const clk   = Number(r[hMap["Clicks"]] || 0);
    const det   = String(r[hMap["Details"]] || "");

    const key = pid ? ('pid:' + pid) : ('k:' + netId + '|' + camp + '|' + plc);
    snapshots.push({ key: key, imp: imp, clk: clk });

    const prev = latestMap[key];
    const isNew = !prev;
    const changed = isNew || prev.imp !== imp || prev.clk !== clk;

    if (changed) {
      const trimmedCampaign  = camp.length > 20 ? camp.substring(0, 20) + "��" : camp;
      const trimmedPlacement = plc.length > 20 ? plc.substring(0, 20) + "��" : plc;

      candidateRows.push({
        netId: netId, adv: adv,
        camp: trimmedCampaign,
        pid: pid,
        plc: trimmedPlacement,
        imp: imp, clk: clk, det: det
      });
    }
  });

  appendTodaySnapshots_(snapshots);
  if (!candidateRows.length) { compactPerfAlertCache_(35); return; }

  const htmlRows = candidateRows.map(function(o){
    return (
      '<tr>' +
      '<td>' + o.netId + '</td>' +
      '<td>' + o.adv + '</td>' +
      '<td>' + o.camp + '</td>' +
      '<td>' + o.pid + '</td>' +
      '<td>' + o.plc + '</td>' +
      '<td>' + o.imp + '</td>' +
      '<td>' + o.clk + '</td>' +
      '<td>' + o.det + '</td>' +
      '</tr>'
    );
  }).join("");

  const table = ''
    + '<p><b>ALERT:</b> ' + MATCH_TEXT + '</p>'
    + '<p>This report lists placements that continue to meet the performance-alert criteria. Items drop off once metrics are corrected or fall below the thresholds, but will continue to be listed within the CM360 CPC/CPM FLIGHT QA reports.</p>'
    + '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-size:11px;">'
    + '<tr style="background:#f2f2f2;font-weight:bold;">'
    + '<th>Network ID</th><th>Advertiser</th><th>Campaign</th><th>Placement ID</th>'
    + '<th>Placement</th><th>Impressions</th><th>Clicks</th><th>Details</th>'
    + '</tr>'
    + htmlRows
    + '</table>'
    + '<br/>'
    + '<p><i>Brought to you by Platform Solutions Automation. (Made by: BK)</i></p>';

  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "M/d/yy");
  const subject = 'ALERT �� PERFORMANCE (pre-monthly-summary) �� ' + todayStr + ' �� ' + candidateRows.length + ' changed/new row(s)';

  uniqueEmails.forEach(function(addr){
    try {
      MailApp.sendEmail({ to: addr, subject: subject, htmlBody: table });
      Utilities.sleep(500);
    } catch (err) {
      Logger.log('�� Failed to email ' + addr + ': ' + err);
    }
  });

  compactPerfAlertCache_(35);
}




// ===== Violation last-change cache (sidecar workbook, retry & batched) =====
function withBackoff_(fn, label, maxTries) {
  label = label || "op";
  maxTries = maxTries || 5;
  let wait = 250;
  for (let i = 1; i <= maxTries; i++) {
    try { return fn(); } catch (e) {
      if (i === maxTries) throw e;
      Utilities.sleep(wait);
      wait = Math.min(wait * 2, 4000);
    }
  }
}

function getVChangeBook_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty('vChangeBookId');
    if (id) return withBackoff_(function(){ return SpreadsheetApp.openById(id); }, "open sidecar");
    const book = withBackoff_(function(){ return SpreadsheetApp.create("_CM360_QA_VChangeCache_" + Date.now()); }, "create sidecar");
    props.setProperty('vChangeBookId', book.getId());
    return book;
  } finally {
    lock.releaseLock();
  }
}

function getVChangeSheet_() {
  const book = getVChangeBook_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    let sh = withBackoff_(function(){ return book.getSheetByName("_Violation Change Cache"); }, "get sheet");
    if (!sh) {
      sh = withBackoff_(function(){ return book.insertSheet("_Violation Change Cache"); }, "insert sheet");
      withBackoff_(function(){ sh.hideSheet(); }, "hide sheet");
    }
    const header = ["key","pe","lastReport","lastImp","lastClk","lastImpChange","lastClkChange"];
    const cur = withBackoff_(function(){ return (sh.getRange(1,1,1,header.length).getValues()[0] || []); }, "read header");
    const ok = header.every(function(h,i){ return String(cur[i]||"").toLowerCase()===h.toLowerCase(); });
    if (!ok) withBackoff_(function(){ sh.getRange(1,1,1,header.length).setValues([header]); }, "write header");
    return sh;
  } finally {
    lock.releaseLock();
  }
}

function migrateViolationPropsToSheetOnce_() {
  const propsDoc = PropertiesService.getDocumentProperties();
  const raw = propsDoc.getProperty('violationChangeMap');
  if (!raw) return;
  let obj; try { obj = JSON.parse(raw); } catch(e) { obj = {}; }
  saveViolationChangeMap_(obj);
  propsDoc.deleteProperty('violationChangeMap');
}

function loadViolationChangeMap_() {
  migrateViolationPropsToSheetOnce_();
  const sh = getVChangeSheet_();
  const lastRow = withBackoff_(function(){ return sh.getLastRow(); }, "getLastRow");
  if (lastRow <= 1) return {};
  const vals = withBackoff_(function(){ return sh.getRange(2,1,lastRow-1,7).getValues(); }, "read cache rows");
  const map = {};
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    const key = String(r[0] || "").trim();
    if (!key) continue;
    map[key] = {
      key:            key,
      pe:            r[1] ? String(r[1]) : null,
      lastReport:    r[2] ? String(r[2]) : null,
      lastImp:       Number(r[3] || 0),
      lastClk:       Number(r[4] || 0),
      lastImpChange: r[5] ? String(r[5]) : null,
      lastClkChange: r[6] ? String(r[6]) : null
    };
  }
  return map;
}

function saveViolationChangeMap_(mapObj) {
  const sh = getVChangeSheet_();
  const keys = Object.keys(mapObj).sort();
  const rows = new Array(keys.length);
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    const r = mapObj[k] || {};
    rows[i] = [
      k,
      r.pe || null,
      r.lastReport || null,
      Number(r.lastImp || 0),
      Number(r.lastClk || 0),
      r.lastImpChange || null,
      r.lastClkChange || null
    ];
  }

  const COLS = 7;
  const last = withBackoff_(function(){ return sh.getLastRow(); }, "getLastRow before clear");
  if (last > 1) withBackoff_(function(){ sh.getRange(2,1,last-1,COLS).clearContent(); }, "clear body");

  if (!rows.length) {
    PropertiesService.getDocumentProperties().deleteProperty('violationChangeMap');
    return;
  }

  const BATCH = 10000;
  for (let start = 0; start < rows.length; start += BATCH) {
    const chunk = rows.slice(start, start + BATCH);
    withBackoff_(function(){ sh.getRange(2 + start, 1, chunk.length, COLS).setValues(chunk); }, "write batch");
    Utilities.sleep(50);
  }

  PropertiesService.getDocumentProperties().deleteProperty('violationChangeMap');
}

function cleanupViolationCache_(mapObj, today) {
  for (const k in mapObj) {
    if (!mapObj.hasOwnProperty(k)) continue;
    const r = mapObj[k];
    const pe  = r.pe ? new Date(r.pe) : null;
    const lic = r.lastImpChange ? new Date(r.lastImpChange) : null;
    const lcc = r.lastClkChange ? new Date(r.lastClkChange) : null;
    if (pe && today > pe) {
      const impOk = !lic || lic <= pe;
      const clkOk = !lcc || lcc <= pe;
      if (impOk && clkOk) delete mapObj[k];
    }
  }
  const ninetyDaysAgo = new Date(Date.now() - 90 * 86400000);
  for (const k2 in mapObj) {
    if (!mapObj.hasOwnProperty(k2)) continue;
    const r2 = mapObj[k2];
    const lr = r2.lastReport ? new Date(r2.lastReport) : null;
    if (lr && lr < ninetyDaysAgo) delete mapObj[k2];
  }
  const remaining = Object.keys(mapObj).map(function(k3){
    const v = mapObj[k3];
    return [k3, v.lastReport ? new Date(v.lastReport).getTime() : 0];
  }).sort(function(a,b){ return b[1]-a[1]; });

  const MAX = 150000;
  if (remaining.length > MAX) {
    for (let i = MAX; i < remaining.length; i++) delete mapObj[remaining[i][0]];
  }
}

function upsertViolationChange_(mapObj, key, rd, imp, clk, pe) {
  const rdISO = rd ? Utilities.formatDate(rd, Session.getScriptTimeZone(), "yyyy-MM-dd") : null;
  const peISO = pe ? Utilities.formatDate(pe, Session.getScriptTimeZone(), "yyyy-MM-dd") : null;

  let rec = mapObj[key];
  if (!rec) {
    rec = mapObj[key] = {
      key: key,
      pe: peISO,
      lastReport: rdISO,
      lastImp: Number(imp || 0),
      lastClk: Number(clk || 0),
      lastImpChange: rdISO,
      lastClkChange: rdISO
    };
  } else {
    if (peISO && peISO !== rec.pe) rec.pe = peISO;
    if (!rec.lastReport || (rdISO && rdISO > rec.lastReport)) rec.lastReport = rdISO;
    if (typeof imp === "number" && imp !== Number(rec.lastImp || 0)) {
      rec.lastImp = Number(imp);
      rec.lastImpChange = rdISO;
    }
    if (typeof clk === "number" && clk !== Number(rec.lastClk || 0)) {
      rec.lastClk = Number(clk);
      rec.lastClkChange = rdISO;
    }
  }
  return {
    lastImpChange: rec.lastImpChange ? new Date(rec.lastImpChange) : null,
    lastClkChange: rec.lastClkChange ? new Date(rec.lastClkChange) : null
  };
}

// ---------------------
// Owner/Rep mapping helpers + lookup from "Networks" (prefer OPS in P��S)
// ---------------------
function normalizeAdv_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/\(.*?\)/g, '')
    .replace(/\[.*?\]/g, '')
    .replace(/\b(inc|llc|ltd|corp|corporation|group)\b/g, '')
    .replace(/[^a-z0-9+]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function resolveRep_(ownerMap, netId, adv) {
  const rawKey  = netId + "|||" + String(adv || "").toLowerCase().trim();
  const normKey = netId + "|||" + normalizeAdv_(adv || "");
  const rr = ownerMap.byKey[rawKey];
  const nr = ownerMap.byKey[normKey];
  return (rr && rr.rep) || (nr && nr.rep) || "Unassigned";
}

function loadOwnerMapFromNetworks_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Networks");
  const byKey = {};

  if (!sh || sh.getLastRow() < 2) return { byKey: byKey };

  const vals = sh.getDataRange().getValues();
  const hdr  = vals[0].map(function(h){ return String(h || "").trim().toLowerCase(); });

  const idIdx = (function() {
    const cands = ["network id","network_id","networkid","cm360 network id"];
    for (let i = 0; i < cands.length; i++) { const c = cands[i]; const idx = hdr.indexOf(c); if (idx !== -1) return idx; }
    return -1;
  })();
  const advIdx = (function() {
    const cands = ["advertiser","advertiser name","advertiser_name","cm360 advertiser","cm360 advertiser name"];
    for (let i = 0; i < cands.length; i++) { const c = cands[i]; const idx = hdr.indexOf(c); if (idx !== -1) return idx; }
    return -1;
  })();

  function findOpsInRange_(hdrArr, start, end) {
    for (let i = start; i <= end && i < hdrArr.length; i++) {
      const name = hdrArr[i];
      if (/ops/.test(name)) return i;
    }
    return -1;
  }
  let repIdx = findOpsInRange_(hdr, 15, 18);

  if (repIdx === -1) {
    const repCands = [
      "account rep ops","rep ops","ops owner","ops member","ops",
      "owner (ops)","operations owner","account owner","owner","rep","sales rep","account lead"
    ];
    for (let i = 0; i < repCands.length; i++) {
      const c = repCands[i];
      const j = hdr.indexOf(c);
      if (j !== -1) { repIdx = j; break; }
    }
  }

  if (idIdx === -1 || advIdx === -1 || repIdx === -1) return { byKey: byKey };

  for (let r = 1; r < vals.length; r++) {
    const netId = String(vals[r][idIdx] || "").trim();
    const adv   = String(vals[r][advIdx] || "").trim();
    const theRep = String(vals[r][repIdx] || "").trim();
    if (!netId || !adv) continue;

    const rawKey  = netId + "|||" + adv.toLowerCase();
    const normKey = netId + "|||" + normalizeAdv_(adv);
    const payload = { rep: theRep || "Unassigned" };

    byKey[rawKey]  = payload;
    byKey[normKey] = payload;
  }

  return { byKey: byKey };
}

// Export a single Sheet as XLSX blob (robust via export endpoint)
function createXLSXFromSheet(sheet) {
  if (!sheet) throw new Error("createXLSXFromSheet: sheet is required");

  const tmp = SpreadsheetApp.create("TMP_EXPORT_" + Date.now());
  const tmpId = tmp.getId();
  const tmpSs = SpreadsheetApp.openById(tmpId);

  const copied = sheet.copyTo(tmpSs).setName(sheet.getName());
  tmpSs.getSheets().forEach(function(s){
    if (s.getSheetId() !== copied.getSheetId()) tmpSs.deleteSheet(s);
  });
  tmpSs.setActiveSheet(copied);
  tmpSs.moveActiveSheet(0);

  const url = 'https://docs.google.com/spreadsheets/d/' + tmpId + '/export?format=xlsx';
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });

  DriveApp.getFileById(tmpId).setTrashed(true);
  return response.getBlob();
}

function getStaleThresholdDays_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const networksSheet = ss.getSheetByName("Networks");
  if (!networksSheet) return 7;

  const raw = String(networksSheet.getRange("H1").getDisplayValue() || "").trim();
  const m = raw.match(/-?\d+(\.\d+)?/);
  let v = m ? Number(m[0]) : NaN;

  if (!isFinite(v) || v <= 0) v = 7;
  v = Math.floor(v);
  Logger.log("Stale threshold days used (from Networks!H1): " + v + " (raw='" + raw + "')");
  return v;
}


/*******************************************************
 * Low-Priority Scoring �� Lightweight (NO sheets/logging)
 *******************************************************/

// Keep these defaults (same signal quality, no sheet I/O)
const X_CH = "[x��]";
const DEFAULT_LP_PATTERNS = [
  ['Impression Pixel/Beacon', `\\b0\\s*${X_CH}\\s*0\\b|\\bzero\\s*by\\s*zero\\b`, 40, 'Zero-size creative', 'Y'],
  ['Impression Pixel/Beacon', `\\b1\\s*${X_CH}\\s*1\\b|\\b1\\s*by\\s*1\\b|\\b1x1(?:cc)?\\b`, 30, '1x1 variants', 'Y'],
  ['Impression Pixel/Beacon', `\\bpixel(?:\\s*only)?\\b|\\bbeacon\\b|\\bclear\\s*pixel\\b|\\btransparent\\s*pixel\\b|\\bspacer\\b|\\bshim\\b`, 20, 'Pixel-ish words', 'Y'],

  ['Click Tracker', `\\bclick\\s*tr(?:ac)?k(?:er)?\\b`, 28, 'click tracker', 'Y'],
  ['Click Tracker', `\\bclick[_-]?(?:trk|tr)\\b|\\bclk[_-]?trk\\b|\\bclktrk\\b|\\bctrk\\b`, 26, 'click/clk tracker shorthands', 'Y'],
  ['Click Tracker', `(^|[^A-Za-z0-9])ct(?:_?trk)\\b`, 22, 'bounded CT_TRK', 'Y'],
  ['Click Tracker', `tracking\\s*1\\s*${X_CH}\\s*1|track(?:ing)?\\s*1x1`, 20, 'tracking 1x1', 'Y'],
  ['Click Tracker', `dfa\\s*zero\\s*placement|zero\\s*placement`, 18, 'legacy DFA zero placement', 'Y'],

  ['VAST/CTV Tracking Tag', `\\bvid(?:eo)?[\\s_\\-]*tag\\b`, 25, 'VID_TAG / video tag', 'Y'],
  ['VAST/CTV Tracking Tag', `\\bvid[\\s_\\-]*:(?:06|15|30)s?\\b`, 22, 'VID:06/15/30 shorthand', 'Y'],
  ['VAST/CTV Tracking Tag', `\\bvast[\\s_\\-]*(?:tag|pixel|tracker)\\b`, 30, 'VAST tag/pixel/tracker', 'Y'],
  ['VAST/CTV Tracking Tag', `\\bdv[_\\-]?tag\\b|\\bgcm[_\\-]?(?:non[_\\-]?)?tag\\b|\\bgcm[_\\-]?dv[_\\-]?tag\\b`, 30, 'DV_TAG/GCM tags', 'Y'],
  ['VAST/CTV Tracking Tag', `\\bvpaid\\b|\\bomsdk\\b|\\bavoc\\b`, 18, 'VPAID/OMSDK/AVOC', 'Y'],

  ['Viewability/Verification', `\\bom(id)?\\b|\\bmoat\\b|\\bias\\b|\\bintegral\\s*ad\\s*science\\b|\\bdoubleverify\\b|\\bcomscore\\b|\\bpixalate\\b|\\bverification\\b|\\bviewability\\b`, 18, 'Verification vendors/terms', 'Y'],

  ['Placeholder/Tag-Only/Test', `\\b[_-]?tag\\b|\\bnon[_-]?tag\\b|\\bplaceholder\\b|\\bdefault\\s*tag\\b|\\bqa\\b|\\btest\\b|\\bsample\\b`, 15, 'Non-serving / test-ish', 'Y'],

  ['Impression-Only Keywords', `\\bimp(?:ression)?[\\s_\\-]*only\\b|\\bimpr[\\s_\\-]*only\\b|\\bview[\\s_\\-]*through\\b`, 20, 'Impr-only phrasing', 'Y'],

  ['Social/3P Pixel', `\\b(meta|facebook|tiktok|snap|pinterest|youtube)[\\s_\\-]*(pixel|tag)\\b`, 15, 'Social pixel/tag', 'Y'],
  ['Social/3P Pixel', `\\bfbq\\b|\\bttq\\b|\\bsnaptr\\b|\\bpintrk\\b|\\btwq\\b|\\bgads\\b`, 15, 'SDK shorthands', 'Y'],

  ['Descriptor Only', `\\b(?:added\\s*value|sponsorship)\\b`, 5, 'Descriptor-only if CPM-only', 'Y'],
  ['Signal', `\\bN\\/A\\b`, 10, 'N/A token in piped name', 'Y']
];

// Negatives used only to *reduce* likelihood when both metrics are present
const DEFAULT_NEG_PATTERNS = [
  ['DisplaySize', `\\b(120\\s*${X_CH}\\s*600|160\\s*${X_CH}\\s*600|300\\s*${X_CH}\\s*50|300\\s*${X_CH}\\s*100|300\\s*${X_CH}\\s*250|300\\s*${X_CH}\\s*600|320\\s*${X_CH}\\s*50|320\\s*${X_CH}\\s*100|336\\s*${X_CH}\\s*280|468\\s*${X_CH}\\s*60|728\\s*${X_CH}\\s*90|970\\s*${X_CH}\\s*90|970\\s*${X_CH}\\s*250|980\\s*${X_CH}\\s*120|980\\s*${X_CH}\\s*240|640\\s*${X_CH}\\s*360|1280\\s*${X_CH}\\s*720|1920\\s*${X_CH}\\s*1080)\\b`, 35, 'Standard creative sizes', 'Y'],
  ['AssetExt', `\\b(?:jpg|jpeg|png|gif|mp4|mov|webm)\\b`, 10, 'Creative file type mentioned', 'Y'],
  ['RealCreativeKeywords', `\\b(?:interstitial|masthead|takeover|homepage|roadblock)\\b`, 15, 'Likely real creatives', 'Y']
];

// Probability tuning (same math, no logging)
const LP_THRESHOLDS = { VERY_LIKELY: 85, LIKELY: 70, POSSIBLE: 55 };
const LP_BASE_SCORE = 40;

let _lpCompiled = null;
let _negCompiled = null;

function compileLPPatternsIfNeeded_() {
  if (_lpCompiled && _negCompiled) return;

  _lpCompiled = DEFAULT_LP_PATTERNS.map(function(r){
    let re = null; try { re = new RegExp(String(r[1]), 'i'); } catch (e) { /* noop */ }
    return {
      category: String(r[0]),
      re: re,
      weight: Number(r[2] || 0),
      label: String(r[0]) + ':' + String(r[1]),
      enabled: String(r[4] || 'Y').toUpperCase().startsWith('Y') && !!re
    };
  });

  _negCompiled = DEFAULT_NEG_PATTERNS.map(function(r){
    let re = null; try { re = new RegExp(String(r[1]), 'i'); } catch (e) { /* noop */ }
    return {
      category: r[0],
      re: re,
      weight: Number(r[2] || 0),
      label: String(r[0]) + ':' + String(r[1]),
      enabled: !!re
    };
  });
}

function normalizeName_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[��]/g, 'x')
    .replace(/\|/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
function clamp_(n, a, b) { return Math.max(a, Math.min(b, n)); }

/**
 * Lightweight classifier:
 * - NO sheet reads/writes
 * - Returns descriptor string or '' (no tag)
 * - gating: 'CPM-only' | 'CPC-only' | 'Mixed'
 */
function scoreAndLabelLowPriority_(placementName, clicks, impr, rowIdOrIndex, gating) {
  gating = gating || ((impr > 0 && clicks === 0) ? 'CPM-only' :
                      (impr === 0 && clicks > 0) ? 'CPC-only' : 'Mixed');

  compileLPPatternsIfNeeded_();

  if (gating === 'Mixed') {
    // Don��t LP-tag rows where both metrics present (or pathological both+clicks>impr)
    return '';
  }

  const s = normalizeName_(placementName);
  let pos = 0, neg = 0;
  const catScores = Object.create(null);

  for (var i=0; i<_lpCompiled.length; i++) {
    var p = _lpCompiled[i];
    if (!p.enabled || !p.re) continue;
    if (p.re.test(s)) {
      pos += p.weight;
      catScores[p.category] = (catScores[p.category] || 0) + p.weight;
    }
  }

  // If Mixed, we��d subtract negatives; for single-metric add a tiny boost when size present
  if (gating !== 'Mixed') {
    var sizeRgx = _negCompiled[0].re;
    if (sizeRgx && sizeRgx.test(s)) {
      pos += 15; // helps 1x1 & obvious ��pixel-ish�� names
      catScores['Impression Pixel/Beacon'] = (catScores['Impression Pixel/Beacon'] || 0) + 15;
    }
  } else {
    for (var j=0; j<_negCompiled.length; j++) {
      var n = _negCompiled[j];
      if (n.enabled && n.re && n.re.test(s)) neg += n.weight;
    }
  }

  var has0x0  = /\b0\s*x\s*0\b|\bzero\s*by\s*zero\b/.test(s);
  var hasTag  = /\bvid(?:eo)?[\s_\-]*tag\b/.test(s) || /\b(?:gcm|dv)[\s_\-]*(?:non[\s_\-]*)?tag\b|\bdv[_\-]?tag\b/.test(s);
  var hasDur  = /\bvid[\s_\-]*:(?:06|15|30)s?\b/.test(s);
  if (has0x0 && (hasTag || hasDur)) {
    pos += 20;
    catScores['VAST/CTV Tracking Tag'] = (catScores['VAST/CTV Tracking Tag'] || 0) + 20;
  }

  if (gating === 'CPC-only' && (catScores['Click Tracker'] || 0) > 0) {
    pos += 10;
  }
  if (gating === 'CPM-only' && (catScores['Impression Pixel/Beacon'] || 0) > 0) {
    pos += 10;
  }

  var probability = clamp_(LP_BASE_SCORE + pos - neg, 0, 100);
  var band = (probability >= LP_THRESHOLDS.VERY_LIKELY) ? 'Very likely'
          : (probability >= LP_THRESHOLDS.LIKELY)      ? 'Likely'
          : (probability >= LP_THRESHOLDS.POSSIBLE)    ? 'Possible'
          : 'Unlikely';

  if (band === 'Unlikely') return '';

  var topCat = '';
  var maxCatScore = -1;
  for (var cat in catScores) {
    if (catScores[cat] > maxCatScore) { maxCatScore = catScores[cat]; topCat = cat; }
  }
  if (!topCat) topCat = 'Impression Pixel/Beacon';

  // Descriptor only; no writes/logging
  return 'Low Priority �� ' + topCat + ' (' + band + ')';
}




// ---------------------
// runQAOnly (auto-resume, chunked, lock-guarded)
// ---------------------
function runQAOnly() {
  // Prevent overlapping runs
  const dlock = LockService.getDocumentLock();
  if (!dlock.tryLock(30000)) { scheduleNextQAChunk_(2); return; }

  // Clear any stale scheduled id right as we start a chunk
  cancelQAChunkTrigger_();

  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    let raw = ss.getSheetByName("Raw Data");
    let out = ss.getSheetByName("Violations");
    
    // Auto-create sheets if they don't exist
    if (!raw) {
      Logger.log('⚠️ Raw Data sheet not found - creating it now...');
      raw = ss.insertSheet("Raw Data");
      raw.getRange("A1:H1").setValues([[
        "Network ID", "Advertiser", "Campaign", "Placement", 
        "Start Date", "End Date", "Cost Structure", "Report Date"
      ]]).setFontWeight("bold");
      Logger.log('✅ Raw Data sheet created');
      return; // No data to process yet
    }
    
    if (!out) {
      Logger.log('⚠️ Violations sheet not found - creating it now...');
      out = ss.insertSheet("Violations");
      out.getRange("A1:O1").setValues([[
        "Network", "Advertiser", "Campaign", "Placement", "Placement ID",
        "Start Date", "End Date", "Cost Structure", "Issue Type", 
        "Owner", "Severity", "$ at Risk", "Low Priority", "Report Date", "Notes"
      ]]).setFontWeight("bold");
      Logger.log('✅ Violations sheet created');
    }

    const data = raw.getDataRange().getValues();
    if (!data || data.length <= 1) return;

    const headers = data[0];
    const m = getHeaderMap(headers);

    const ignoreSet = loadIgnoreAdvertisers();
    const ownerMap  = loadOwnerMapFromNetworks_();
    const vMap      = loadViolationChangeMap_();



    compileLPPatternsIfNeeded_();

    let state = getQAState_();
    const totalRows = data.length - 1; // excluding header
    const freshStart = !state || state.totalRows !== totalRows;

    if (freshStart) {
      clearViolations();
      state = { session: String(Date.now()), next: 2, totalRows: totalRows };
      saveQAState_(state);
      cancelQAChunkTrigger_();
    }

    const startTime = Date.now();
    const today = new Date();
    const firstOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);

    // ���� Tweak these constants in your file (outside this function) ����
    // const QA_CHUNK_ROWS = 3500;
    // const QA_TIME_BUDGET_MS = 4.2 * 60 * 1000;

    let processed = 0;
    const resultsChunk = [];

    for (let r = state.next; r < data.length; r++) {
      const row = data[r];
      const adv  = row[m["Advertiser"]] && String(row[m["Advertiser"]]).trim();
      const camp = row[m["Campaign"]]   || "";

      // Progress logging every 500 rows
      if (processed > 0 && processed % 500 === 0) {
        Logger.log('[runQAOnly] Progress: Processed ' + processed + ' rows in current chunk');
      }

      const advLower = adv ? adv.toLowerCase() : "";
      if (advLower && (ignoreSet.has(advLower) || advLower.includes("bidmanager"))) { state.next = r + 1; continue; }
      if (camp && String(camp).includes("DART Search"))                               { state.next = r + 1; continue; }
      if (adv === "Grand Total:")                                                     { state.next = r + 1; continue; }

      const imp = Number(row[m["Impressions"]] || 0);
      const clk = Number(row[m["Clicks"]] || 0);
      if (imp === 0 && clk === 0) { state.next = r + 1; continue; }

      const ctr = imp > 0 ? (clk / imp) * 100 : 0;

      // Your CPC/CPM formulas
      const cpc = clk * 0.008;
      const cpm = (imp / 1000) * 0.034;

      const ps  = new Date(row[m["Placement Start Date"]]);
      const pe  = new Date(row[m["Placement End Date"]]);
      const rd  = new Date(row[m["Report Date"]]);

      const daysRem  = Math.ceil((pe - rd) / 86400000);
      const eom      = new Date(rd.getFullYear(), rd.getMonth() + 1, 0);
      const daysLeft = Math.ceil((eom - rd) / 86400000);

      const flen = (pe - ps) / 86400000;
      const din  = (rd - ps) / 86400000;
      const pctComplete = pe.getTime() === ps.getTime()
        ? (rd > pe ? 100 : 0)
        : Math.min(100, Math.max(0, (din / flen) * 100));

      const issueTypes = [];
      const details    = [];
      let risk = "";

      // ?? BILLING
      if (pe < firstOfMonth && clk > imp) {
        issueTypes.push("?? BILLING: Expired CPC Risk");
        details.push("Ended " + pe.toDateString() + " with clicks (" + clk + ") > impressions (" + imp + ")");
        risk = "?? Expired Risk";
      } else if (pe < rd && clk > imp) {
        issueTypes.push("� BILLING: Recently Expired CPC Risk");
        details.push("Ended " + pe.toDateString() + " and still has clicks > impressions");
        risk = "��️ Expired This Month";
      } else if (rd <= pe && clk > imp && cpc > 10) {
        issueTypes.push("?? BILLING: Active CPC Billing Risk");
        details.push("Active: clicks (" + clk + ") > impressions (" + imp + "), $CPC = $" + cpc.toFixed(2));
        risk = "��️ Active CPC Risk";
      }

      // � DELIVERY
      if (pe < firstOfMonth && rd >= firstOfMonth && (imp > 0 || clk > 0)) {
        issueTypes.push("� DELIVERY: Post-Flight Activity");
        details.push("Ended " + pe.toDateString() + " but has " + imp + " impressions and " + clk + " clicks");
      }

      // � PERFORMANCE
      if (ctr >= 90 && cpm >= 10) {
        issueTypes.push("� PERFORMANCE: CTR �� 90% & CPM �� $10");
        details.push("CTR = " + ctr.toFixed(2) + "%, $CPM = $" + cpm.toFixed(2));
      }

      // � COST
      let isCPMOnly = false;
      let isCPCOnly = false;
      if (cpc > 0 && cpm === 0 && cpc > 10) {
        issueTypes.push("� COST: CPC Only > $10");
        details.push("No CPM spend, $CPC = $" + cpc.toFixed(2));
        if (imp === 0 && clk > 0) isCPCOnly = true;
      }
      if (cpm > 0 && cpc === 0 && cpm > 10) {
        issueTypes.push("� COST: CPM Only > $10");
        details.push("No CPC spend, $CPM = $" + cpm.toFixed(2));
        if (imp > 0 && clk === 0) isCPMOnly = true;
      }
      if (cpc > 0 && cpm > 0 && clk > imp && cpc > 10) {
        issueTypes.push("� COST: CPC+CPM Clicks > Impr & CPC > $10");
        details.push("Clicks > impressions with both CPC and CPM charges (CPC = $" + cpc.toFixed(2) + ")");
      }

      // --- Low-priority tagging via scorer (gating-aware) �� no sheet writes ---
      const bothMetricsPresent = imp > 0 && clk > 0;
      const clicksExceedImprWithBoth = bothMetricsPresent && (clk > imp);
      const gating = (imp > 0 && clk === 0) ? 'CPM-only' :
                     (imp === 0 && clk > 0) ? 'CPC-only' : 'Mixed';

      if (!bothMetricsPresent && !clicksExceedImprWithBoth) {
        const placement = row[m["Placement"]];
        const rowIdOrIndex = String(row[m["Placement ID"]] || (r + 1));
        const lpDescriptor = scoreAndLabelLowPriority_(placement, clk, imp, rowIdOrIndex, gating);
        if (lpDescriptor) {
          issueTypes.push("� COST: (Low Priority) " + lpDescriptor.replace(/^Low Priority ��\s*/, ""));
        }
      }
      // --- end Low-priority tagging ---

      if (!issueTypes.length) { state.next = r + 1; continue; }

      const pid = String(row[m["Placement ID"]] || "");
      const key = pid ? ("pid:" + pid) : ("k:" + row[m["Network ID"]] + "|" + camp + "|" + row[m["Placement"]]);
      const changes = upsertViolationChange_(vMap, key, rd, imp, clk, pe);

      function daysSince_(lastChangeDate, reportDate) {
        if (!(lastChangeDate instanceof Date) || isNaN(lastChangeDate) || !(reportDate instanceof Date) || isNaN(reportDate)) return "";
        const ms = reportDate.getTime() - lastChangeDate.getTime();
        if (ms < 0) return "";
        return Math.floor(ms / 86400000);
      }
      const lastImpDays = changes.lastImpChange ? daysSince_(changes.lastImpChange, rd) : "";
      const lastClkDays = changes.lastClkChange ? daysSince_(changes.lastClkChange, rd) : "";

      const ownerOps = resolveRep_(ownerMap, String(row[m["Network ID"]] || ""), adv) || "Unassigned";

      resultsChunk.push([
        row[m["Network ID"]], row[m["Report Date"]], row[m["Advertiser"]], row[m["Campaign"]],
        row[m["Campaign Start Date"]], row[m["Campaign End Date"]], row[m["Ad"]], row[m["Placement ID"]],
        row[m["Placement"]], row[m["Placement Start Date"]], row[m["Placement End Date"]],
        imp, clk, ctr.toFixed(2) + "%", daysRem, pctComplete.toFixed(1) + "%", daysLeft,
        risk, "$" + cpc.toFixed(2), "$" + cpm.toFixed(2), issueTypes.join(", "), details.join(" | "),
        lastImpDays, lastClkDays, ownerOps
      ]);

      processed++;
      state.next = r + 1;

      // Respect chunk size & time budget
      if (processed >= QA_CHUNK_ROWS) break;
      if ((Date.now() - startTime) >= QA_TIME_BUDGET_MS) break;
    }

    // Persist violation-change snapshot
    cleanupViolationCache_(vMap, today);
    saveViolationChangeMap_(vMap);

    // Write this chunk's rows
    if (resultsChunk.length) {
      const width = resultsChunk[0].length;
      const startWriteRow = out.getLastRow() + 1;
      out.getRange(startWriteRow, 1, resultsChunk.length, width).setValues(resultsChunk);
    }

    // Decide: finished or schedule next chunk
    if (state.next >= (data.length)) {
      clearQAState_();
      cancelQAChunkTrigger_();
      Logger.log("� runQAOnly complete. Processed all " + totalRows + " data rows.");
    } else {
      saveQAState_(state);
      Logger.log("�� runQAOnly partial: processed " + processed + " rows this run. Next row index: "
        + state.next + " / " + (data.length - 1));
      scheduleNextQAChunk_(2); // resume soon
    }
  } finally {
    dlock.releaseLock();
  }
}




// === Helpers for "Immediate Attention" selection ===
function _parseMoney_(s) { // "$12.34" -> 12.34
  var n = String(s || "").replace(/[^\d.-]/g, "");
  var v = parseFloat(n);
  return isFinite(v) ? v : 0;
}
function _parsePct_(s) { // "95.00%" -> 95
  var n = String(s || "").replace(/[^\d.-]/g, "");
  var v = parseFloat(n);
  return isFinite(v) ? v : 0;
}






// ---------------------
// sendEmailSummary (size-safe) �� UPDATED with extra buckets
// ---------------------
function sendEmailSummary() {
  // Skip if QA is still running in chunks
  const _qaState = getQAState_();
  if (_qaState && _qaState.session) {
    Logger.log("sendEmailSummary skipped: QA still in progress (chunked).");
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();

  // Only send on/after the 15th
  if (today.getDate() < 15) {
    Logger.log("Email summary skipped: before the 15th of the month.");
    return;
  }

  // --- Email size & filtering controls ---
  const INCLUDE_APPENDIX = false;
  const INCLUDE_ZERO_NETS = false;
  const MAX_ROWS_PER_OWNER = 30;
  const MAX_TOTAL_OWNER_ROWS = 1000;
  const MAX_HTML_CHARS = 90000;

  // Sheets
  const sheet           = ss.getSheetByName("Violations");
  const rawSheet        = ss.getSheetByName("Raw Data");
  const networksSheet   = ss.getSheetByName("Networks");
  const recipientsSheet = ss.getSheetByName("EMAIL LIST");
  if (!sheet || !rawSheet || !recipientsSheet) return;

  // Recipients
  const emails = recipientsSheet.getRange("A2:A").getValues()
    .flat().map(function(e){ return String(e || "").trim(); }).filter(Boolean);
  const uniqueEmails = Array.from(new Set(emails));
  if (uniqueEmails.length === 0) return;

  // Data
  const violations = sheet.getDataRange().getValues();
  const rawData    = rawSheet.getDataRange().getValues();
  if (violations.length <= 1) return;

  const hMap = getHeaderMap(violations[0]);
  const rMap = getHeaderMap(rawData[0]);

  // --- Network ID -> Network Name ---
  function buildNetworkNameMap_() {
    if (!networksSheet) return {};
    const vals = networksSheet.getDataRange().getValues();
    const map = {};
    for (let r = 1; r < vals.length; r++) {
      const idRaw = vals[r][0];
      const name  = String(vals[r][1] == null ? "" : vals[r][1]).replace(/\u00A0/g, " ").trim();
      if (!idRaw) continue;
      let id = "";
      if (typeof idRaw === "number") id = String(Math.trunc(idRaw));
      else {
        let s = String(idRaw).replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
        s = s.replace(/,/g, "");
        const digits = s.replace(/\D+/g, "");
        id = digits || s;
      }
      if (id) map[id] = name;
    }
    return map;
  }
  const networkNameMap = buildNetworkNameMap_();

  // --- Counts per network ---
  const placementCounts = {};
  rawData.slice(1).forEach(function(r){
    const id = String(r[rMap["Network ID"]] || "");
    if (id) placementCounts[id] = (placementCounts[id] || 0) + 1;
  });

  // --- Violation counts per network (by group) ---
  const violationCounts = {};
  violations.slice(1).forEach(function(r){
    const id    = String(r[hMap["Network ID"]] || "");
    const types = String(r[hMap["Issue Type"]] || "").split(", ");
    if (!violationCounts[id]) {
      violationCounts[id] = { "� BILLING": 0, "� DELIVERY": 0, "� PERFORMANCE": 0, "� COST": 0 };
    }
    types.forEach(function(t){
      if (t.startsWith("�")) violationCounts[id]["� BILLING"]++;
      if (t.startsWith("�")) violationCounts[id]["� DELIVERY"]++;
      if (t.startsWith("�")) violationCounts[id]["� PERFORMANCE"]++;
      if (t.startsWith("�")) violationCounts[id]["� COST"]++;
    });
  });

  // --- Network summary table ---
  let networkSummary =
      '<p><b>Network-Level QA Summary</b></p>'
    + '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse; font-size: 11px;">'
    + '<tr style="background-color: #f2f2f2; font-weight: bold;">'
    + '<th>Network ID</th><th>Network Name</th><th>Placements Checked</th>'
    + '<th>� BILLING</th><th>� DELIVERY</th><th>� PERFORMANCE</th><th>� COST</th>'
    + '</tr>';

  Object.entries(networkNameMap)
    .filter(function(pair){
      const id = pair[0];
      if (INCLUDE_ZERO_NETS) return true;
      const vc = violationCounts[id] || { "� BILLING":0,"� DELIVERY":0,"� PERFORMANCE":0,"� COST":0 };
      const total = vc["� BILLING"] + vc["� DELIVERY"] + vc["� PERFORMANCE"] + vc["� COST"];
      return total > 0;
    })
    .sort(function(a, b){ return a[1].localeCompare(b[1]); })
    .forEach(function(entry){
      const id = entry[0], name = entry[1];
      const pc = placementCounts[id] || 0;
      const vc = violationCounts[id] || { "� BILLING":0,"� DELIVERY":0,"� PERFORMANCE":0,"� COST":0 };
      networkSummary += '<tr>'
        + '<td>' + id + '</td><td>' + name + '</td><td>' + pc + '</td>'
        + '<td>' + vc["� BILLING"] + '</td><td>' + vc["� DELIVERY"] + '</td><td>' + vc["� PERFORMANCE"] + '</td><td>' + vc["� COST"] + '</td>'
        + '</tr>';
    });
  networkSummary += '</table><br/>';

  // --- Grouped issue summary (unchanged) ---
  const groupedCounts = { "� BILLING": {}, "� DELIVERY": {}, "� PERFORMANCE": {}, "� COST": {} };
  violations.slice(1).forEach(function(r){
    const types = String(r[hMap["Issue Type"]] || "").split(", ");
    types.forEach(function(t){
      const match = t.match(/^(�|�|�|�)\s(\w+):\s(.+)/);
      if (match) {
        const emoji = match[1], group = match[2], subtype = match[3];
        const key = emoji + " " + group;
        groupedCounts[key] = groupedCounts[key] || {};
        groupedCounts[key][subtype] = (groupedCounts[key][subtype] || 0) + 1;
      }
    });
  });
  let summaryHtml = "";
  Object.entries(groupedCounts).forEach(function(entry){
    const groupLabel = entry[0], subtypes = entry[1];
    summaryHtml += "<b>" + groupLabel + "</b><ul>";
    Object.entries(subtypes).forEach(function(st){
      const subtype = st[0], count = st[1];
      if (count > 0) summaryHtml += "<li>" + subtype + ": " + count + "</li>";
    });
    summaryHtml += "</ul>";
  });

  // --- Immediate Attention �� Key Issues (by Owner) �� UPDATED bucket logic
  function buildImmediateAttentionByOwner_() {
    const ownerMap = loadOwnerMapFromNetworks_();
  const perOwner = {};

  // Column indexes
  const idx = {
    netId: hMap["Network ID"],
    adv:   hMap["Advertiser"],
    camp:  hMap["Campaign"],
    pid:   hMap["Placement ID"],
    plc:   hMap["Placement"],
    impr:  hMap["Impressions"],
    clk:   hMap["Clicks"],
    ctr:   hMap["CTR (%)"],
    cpc$:  hMap["$CPC"],
    cpm$:  hMap["$CPM"],
    issues:hMap["Issue Type"],
    rd:    hMap["Report Date"],
    pe:    hMap["Placement End Date"]
  };

  // bucket order (lower = higher priority in sort)
  const BUCKETS = {
    PERF: 1,               // � Performance
    COST_BIMBAL: 2,        // � CPC+CPM clicks>impr & $CPC>10
    BILLING: 3,            // � (Active/Recently Expired/Expired) + tightened rules
    DELIV_STRICT: 4,       // � Post-flight + clicks>impr + $CPC>10
    DELIV_CPM_ONLY: 5,     // � Post-flight + CPM-only >$10
    DELIV_GENERAL: 6       // � Post-flight (any activity) but only if $CPC>10 || $CPM>10
  };

  const today = new Date();
  const firstOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);

  function qualifies_(row) {
    const issues = String(row[idx.issues] || "");
    // exclude Low Priority rows entirely
    if (/\(Low Priority\)/i.test(issues)) return null;

    const imp = Number(row[idx.impr] || 0);
    const clk = Number(row[idx.clk] || 0);
    const both = imp > 0 && clk > 0;
    const clicksGtImpr = both && (clk > imp);

    const cpc = _parseMoney_(row[idx.cpc$]);
    const cpm = _parseMoney_(row[idx.cpm$]);
    const ctrPct = _parsePct_(row[idx.ctr]);

    const rd = new Date(row[idx.rd]);
    const pe = new Date(row[idx.pe]);
    const isPostFlight = pe < firstOfMonth && rd >= firstOfMonth;

    // === Your inclusion rules ===

    // � PERFORMANCE: CTR �� 90% & CPM �� $10
    const isPerformance = /�\s*PERFORMANCE: CTR �� 90% & CPM �� \$?10/.test(issues) ||
                          (ctrPct >= 90 && cpm >= 10);

    // � CPC+CPM Clicks > Impr & CPC > $10  (both metrics, clicks>impr & CPC>10)
    const isCostBothMetricsClicksGtImpr = /�\s*COST: CPC\+CPM Clicks > Impr.*CPC > \$?10/i.test(issues) ||
                                          (both && clicksGtImpr && cpc > 10);

    // � BILLING (tightened to both metrics, clicks>impr & $CPC>10)
    const isBillingActive   = /�\s*BILLING: Active CPC Billing Risk/i.test(issues)   && both && clicksGtImpr && cpc > 10;
    const isBillingRecent   = /�\s*BILLING: Recently Expired CPC Risk/i.test(issues) && both && clicksGtImpr && cpc > 10;
    const isBillingExpired  = /�\s*BILLING: Expired CPC Risk/i.test(issues)          && both && clicksGtImpr && cpc > 10;

    // � DELIVERY (Post-Flight) inclusions you selected
    // 1) Strict: post-flight + both metrics + clicks>impr + $CPC>10
    const isDelivStrict = /�\s*DELIVERY: Post-Flight Activity/i.test(issues) && isPostFlight && both && clicksGtImpr && cpc > 10;
    // 2) CPM-only > $10 (post-flight)
    const isDelivCpmOnly = /�\s*DELIVERY: Post-Flight Activity/i.test(issues) && isPostFlight && (imp > 0 && clk === 0) && cpm > 10;
    // 3) General: post-flight, include only if $CPC>10 OR $CPM>10
    const isDelivGeneral = /�\s*DELIVERY: Post-Flight Activity/i.test(issues) && isPostFlight && (cpc > 10 || cpm > 10);

    // �� Explicit excludes
    const isCpcOnly = /�\s*COST:\s*CPC\s*Only\s*>\s*\$?10/i.test(issues) || (imp === 0 && clk > 0 && cpc > 10);
    const isCpmOnly = /�\s*COST:\s*CPM\s*Only\s*>\s*\$?10/i.test(issues) || (imp > 0 && clk === 0 && cpm > 10);
    if (isCpcOnly || isCpmOnly) return null;

    // decide bucket (highest priority match wins)
    if (isPerformance)                    return { bucket: BUCKETS.PERF };
    if (isCostBothMetricsClicksGtImpr)    return { bucket: BUCKETS.COST_BIMBAL };
    if (isBillingActive || isBillingRecent || isBillingExpired)
                                           return { bucket: BUCKETS.BILLING };
    if (isDelivStrict)                    return { bucket: BUCKETS.DELIV_STRICT };
    if (isDelivCpmOnly)                   return { bucket: BUCKETS.DELIV_CPM_ONLY };
    if (isDelivGeneral)                   return { bucket: BUCKETS.DELIV_GENERAL };

    return null; // not included
  }

  // collect rows per owner
  for (let i = 1; i < violations.length; i++) {
    const row = violations[i];
    const q = qualifies_(row);
    if (!q) continue;

    const netId = String(row[idx.netId] || "").trim();
    const adv   = String(row[idx.adv]   || "").trim();
    const rep   = resolveRep_(ownerMap, netId, adv);

    if (!perOwner[rep]) perOwner[rep] = [];
    perOwner[rep].push({
      bucket: q.bucket,
      adv: adv,
      camp: String(row[idx.camp] || ""),
      pid:  String(row[idx.pid]  || ""),
      plc:  String(row[idx.plc]  || ""),
      imp:  Number(row[idx.impr] || 0),
      clk:  Number(row[idx.clk]  || 0),
      issue:String(row[idx.issues] || "")
    });
  }

  const owners = Object.keys(perOwner).sort((a,b)=> a.toLowerCase().localeCompare(b.toLowerCase()));
  if (!owners.length) return "";

  let html = "<p><b>Immediate Attention �� Key Issues (by Owner)</b></p>";
  let totalRows = 0;

  for (const rep of owners) {
    if (totalRows >= MAX_TOTAL_OWNER_ROWS) break;
    const arr = perOwner[rep];

    // sort: bucket �� advertiser A��Z �� clicks desc �� impressions desc �� placement id
    arr.sort(function(a, b){
      if (a.bucket !== b.bucket) return a.bucket - b.bucket;
      const aAdv = String(a.adv||"").toLowerCase(), bAdv = String(b.adv||"").toLowerCase();
      if (aAdv !== bAdv) return aAdv.localeCompare(bAdv);
      if (b.clk !== a.clk) return b.clk - a.clk;
      if (b.imp !== a.imp) return b.imp - a.imp;
      return a.pid.localeCompare(b.pid);
    });

    const take = Math.min(arr.length, MAX_ROWS_PER_OWNER, MAX_TOTAL_OWNER_ROWS - totalRows);
    if (take <= 0) break;
    totalRows += take;

    html += "<p><b>" + rep + "</b> (Showing " + take + " of " + arr.length + ")</p>";
    html += '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse; font-size: 11px;">'
         +  '<tr style="background-color:#f9f9f9;font-weight:bold;">'
         +  '<th>Advertiser</th><th>Campaign</th><th>Placement ID</th><th>Placement</th><th>Impr</th><th>Clicks</th><th>Issue(s)</th>'
         +  '</tr>';

    for (let i = 0; i < take; i++) {
      const o = arr[i];
      const campShort = o.camp.length > 40 ? o.camp.substring(0, 40) + "��" : o.camp;
      const plcShort  = o.plc.length  > 30 ? o.plc.substring(0, 30)  + "��" : o.plc;
      html += "<tr>"
           +  "<td>" + o.adv + "</td>"
           +  "<td>" + campShort + "</td>"
           +  "<td>" + o.pid + "</td>"
           +  "<td>" + plcShort + "</td>"
           +  "<td>" + o.imp + "</td>"
           +  "<td>" + o.clk + "</td>"
           +  "<td>" + o.issue + "</td>"
           +  "</tr>";
    }
    html += "</table><br/>";
  }

  return html;
}

const immediateAttentionHtml = buildImmediateAttentionByOwner_(); // still inside sendEmailSummary()


  // --- Stale metrics (unchanged) ---
  const thresholdDays = getStaleThresholdDays_();
  let staleImp = 0, staleClk = 0;
  const impIdx = hMap["Last Imp Change"], clkIdx = hMap["Last Click Change"];
  if (impIdx !== undefined || clkIdx !== undefined) {
    for (let i = 1; i < violations.length; i++) {
      const r = violations[i];
      const impDays = impIdx !== undefined ? Number(r[impIdx]) : NaN;
      const clkDays = clkIdx !== undefined ? Number(r[clkIdx]) : NaN;
      if (isFinite(impDays) && impDays >= thresholdDays) staleImp++;
      if (isFinite(clkDays) && clkDays >= thresholdDays) staleClk++;
    }
  }
  const staleHtml =
      "<b>Stale Metrics (this month)</b><ul>"
    + "<li>Placements with no new impressions since last change (�� " + thresholdDays + " days): " + staleImp + "</li>"
    + "<li>Placements with no new clicks since last change (�� " + thresholdDays + " days): " + staleClk + "</li>"
    + "</ul>";

  // Appendix (optional)
  const violationsAppendixHtml =
      '<p><b>What the Violations tab tracks</b></p>'
    + '<ul>'
    + '<li><b>� BILLING</b><ul>'
    + '<li><b>Expired CPC Risk</b> �� Ended before this month and clicks &gt; impressions.</li>'
    + '<li><b>Recently Expired CPC Risk</b> �� Ended earlier this month and still clicks &gt; impressions.</li>'
    + '<li><b>Active CPC Billing Risk</b> �� Active (report date �� end date), clicks &gt; impressions, and $CPC &gt; $10.</li>'
    + '</ul></li>'
    + '<li><b>� DELIVERY</b><ul>'
    + '<li><b>Post-Flight Activity</b> �� Ended before this month but shows impressions or clicks this month.</li>'
    + '</ul></li>'
    + '<li><b>� PERFORMANCE</b><ul>'
    + '<li><b>CTR �� 90% &amp; CPM �� $10</b> �� Extreme CTR with meaningful CPM spend.</li>'
    + '</ul></li>'
    + '<li><b>� COST</b><ul>'
    + '<li><b>CPC Only &gt; $10</b> �� No CPM spend and $CPC &gt; $10.</li>'
    + '<li><b>CPM Only &gt; $10</b> �� No CPC spend and $CPM &gt; $10.</li>'
    + '<li><b>CPC+CPM Clicks &gt; Impr &amp; CPC &gt; $10</b> �� Both CPC &amp; CPM, clicks &gt; impressions, and $CPC &gt; $10.</li>'
    + '<li><i>(Low Priority tags exist in attachment but are excluded from this section)</i></li>'
    + '</ul></li>'
    + '</ul>';

  // Attachment
  const todayformatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "M.d.yy");
  const fileName = "CM360_QA_Violations_" + todayformatted + ".xlsx";
  const xlsxBlob = createXLSXFromSheet(sheet).setName(fileName);

  // Assemble body
  const subject = "!!!TESTING VS CODE VERSION!!!!!CM360 CPC/CPM FLIGHT QA �� " + todayformatted;
  let htmlBody =
      networkSummary
    + '<p>The below is a table of the following Billing, Delivery, Performance and Cost issues:</p>'
    + summaryHtml
    + (immediateAttentionHtml ? ('<br/>' + immediateAttentionHtml) : '')
    + '<br/>' + staleHtml
    + (INCLUDE_APPENDIX ? ('<br/>' + violationsAppendixHtml) : '')
    + '<p><i>Brought to you by the Platform Solutions Automation. (Made by: BK)</i></p>';

  // Safety trim if needed
  if (htmlBody.length > MAX_HTML_CHARS) {
    htmlBody = htmlBody.slice(0, MAX_HTML_CHARS - 1200)
             + '<p><i>(trimmed for size �� full detail in the attached XLSX)</i></p>';
  }

  // Send
  uniqueEmails.forEach(function(addr){
    try {
      MailApp.sendEmail({ to: addr, subject: subject, htmlBody: htmlBody, attachments: [xlsxBlob] });
      Utilities.sleep(500);
    } catch (err) {
      Logger.log("Failed to email " + addr + ": " + err);
    }
  });
}



function fmtMs_(ms) {
  if (ms < 0) ms = 0;
  var s = Math.floor(ms / 1000);
  var m = Math.floor(s / 60);
  var r = s % 60;
  return (m + 'm ' + r + 's');
}

function logStep_(label, fn, runStartMs, quotaMinutes) {
  var stepStart = Date.now();
  Logger.log('�� ' + label + ' �� START @ ' + new Date(stepStart).toISOString());
  try {
    var out = fn();
    SpreadsheetApp.flush();
    var stepMs = Date.now() - stepStart;
    var totalMs = Date.now() - runStartMs;
    var quotaMs = (quotaMinutes || 6) * 60 * 1000;
    var leftMs = quotaMs - totalMs;

    Logger.log('� ' + label + ' �� DONE in ' + fmtMs_(stepMs)
      + ' (since run start: ' + fmtMs_(totalMs)
      + ', est. time left: ' + fmtMs_(leftMs) + ')');

    if (leftMs <= 60000) {
      Logger.log('�� WARNING: ~' + Math.max(0, Math.floor(leftMs/1000)) + 's left in Apps Script quota window.');
    }
    return out;
  } catch (e) {
    Logger.log('�� ' + label + ' �� ERROR: ' + (e && e.stack ? e.stack : e));
    throw e;
  }
}

// ---------------------
// runItAll (with execution logging per step) �� MANUAL USE
// ---------------------
function runItAll() {
  var APPROX_QUOTA_MINUTES = 6; // leave at 6 unless your domain truly has more
  var runStart = Date.now();
  Logger.log('� runItAll �� START @ ' + new Date(runStart).toISOString()
             + ' (approx quota: ' + APPROX_QUOTA_MINUTES + ' min)');

  try {
    // 1) Prep & ingest
    logStep_('trimAllSheetsToData_', function(){ trimAllSheetsToData_(); }, runStart, APPROX_QUOTA_MINUTES);
    logStep_('importDCMReports',     function(){ importDCMReports();      }, runStart, APPROX_QUOTA_MINUTES);

    // 2) If low on time, schedule QA and exit (handoff)
    var totalMs  = Date.now() - runStart;
    var quotaMs  = APPROX_QUOTA_MINUTES * 60 * 1000;
    var timeLeft = Math.max(0, quotaMs - totalMs);

    if (timeLeft < 2 * 60 * 1000) {
      Logger.log('�� Not enough time left for QA (' + Math.floor(timeLeft/1000) + 's). Scheduling QA handoff.');
      clearQAState_();           // ensure a fresh QA session
      cancelQAChunkTrigger_();   // clear any stale chunk trigger
      scheduleNextQAChunk_(1);   // kick off the first QA chunk shortly
      return;                    // exit cleanly to avoid hitting the 6-min wall
    }

    // 3) Otherwise, run at most one QA chunk now
    logStep_('runQAOnly (single chunk)', function(){ runQAOnly(); }, runStart, APPROX_QUOTA_MINUTES);

    // 4) Alerts & summary (summary already guards on QA completion & date)
    logStep_('sendPerformanceSpikeAlertIfPre15', function(){ sendPerformanceSpikeAlertIfPre15(); }, runStart, APPROX_QUOTA_MINUTES);
    logStep_('sendEmailSummary',                 function(){ sendEmailSummary();                 }, runStart, APPROX_QUOTA_MINUTES);
  } finally {
    var totalMs = Date.now() - runStart;
    Logger.log('🏁 runItAll �� FINISHED in ' + fmtMs_(totalMs));
  }
}

// ---------------------
// runItAllMorning (no email, for time-driven trigger)
// ---------------------
function runItAllMorning() {
  var APPROX_QUOTA_MINUTES = 6; // same budget, but we stop before email
  var runStart = Date.now();
  Logger.log('� runItAllMorning �� START @ ' + new Date(runStart).toISOString()
             + ' (approx quota: ' + APPROX_QUOTA_MINUTES + ' min)');

  try {
    // 1) Prep & ingest
    logStep_('trimAllSheetsToData_', function(){ trimAllSheetsToData_(); }, runStart, APPROX_QUOTA_MINUTES);
    logStep_('importDCMReports',     function(){ importDCMReports();      }, runStart, APPROX_QUOTA_MINUTES);

    // 2) If low on time, schedule QA and exit (handoff)
    var totalMs  = Date.now() - runStart;
    var quotaMs  = APPROX_QUOTA_MINUTES * 60 * 1000;
    var timeLeft = Math.max(0, quotaMs - totalMs);

    if (timeLeft < 2 * 60 * 1000) {
      Logger.log('�� Not enough time left for QA (' + Math.floor(timeLeft/1000) + 's). Scheduling QA handoff.');
      clearQAState_();           // ensure a fresh QA session
      cancelQAChunkTrigger_();   // clear any stale chunk trigger
      scheduleNextQAChunk_(1);   // kick off the first QA chunk shortly
      return;                    // exit cleanly to avoid hitting the 6-min wall
    }

    // 3) Run at most one QA chunk now
    logStep_('runQAOnly (single chunk)', function(){ runQAOnly(); }, runStart, APPROX_QUOTA_MINUTES);

    // 4) Performance spike alert (fast; safe to keep here)
    logStep_('sendPerformanceSpikeAlertIfPre15', function(){ sendPerformanceSpikeAlertIfPre15(); }, runStart, APPROX_QUOTA_MINUTES);

    // �� NO sendEmailSummary here �� that gets its own trigger/window
  } finally {
    var totalMs = Date.now() - runStart;
    Logger.log('🏁 runItAllMorning �� FINISHED in ' + fmtMs_(totalMs));
  }
}

// ---------------------
// runDailyEmailSummary (email only, for separate trigger)
// ---------------------
function runDailyEmailSummary() {
  var APPROX_QUOTA_MINUTES = 6;
  var runStart = Date.now();
  Logger.log('� runDailyEmailSummary �� START @ ' + new Date(runStart).toISOString()
             + ' (approx quota: ' + APPROX_QUOTA_MINUTES + ' min)');

  try {
    // sendEmailSummary already:
    //  - skips if QA still has an active session
    //  - skips before the 15th of the month
    logStep_('sendEmailSummary', function(){ sendEmailSummary(); }, runStart, APPROX_QUOTA_MINUTES);
  } finally {
    var totalMs = Date.now() - runStart;
    Logger.log('🏁 runDailyEmailSummary �� FINISHED in ' + fmtMs_(totalMs));
  }
}



// ---------------------
// arrayToCsv (utility)
// ---------------------
function arrayToCsv(data) {
  return data.map(function(row){ return row.map(function(cell){ return '"' + cell + '"'; }).join(","); }).join("\n");
}

// ---------------------
// Trim all sheets' grids (reclaim cells)
// ---------------------
function trimAllSheetsToData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sh){
    const lastRow = Math.max(1, sh.getLastRow());
    const lastCol = Math.max(1, sh.getLastColumn());

    const maxRows = sh.getMaxRows();
    const targetRows = Math.max(2, lastRow);
    if (maxRows > targetRows) {
      sh.deleteRows(targetRows + 1, maxRows - targetRows);
    }

    const maxCols = sh.getMaxColumns();
    const targetCols = Math.max(1, lastCol);
    if (maxCols > targetCols) {
      sh.deleteColumns(targetCols + 1, maxCols - targetCols);
    }
  });
}


// =====================================================================================================================
// ========================================== V2 DASHBOARD SYSTEM (BETA) ==============================================
// =====================================================================================================================
// 
// Purpose: Enhanced violations dashboard with priority scoring, financial impact tracking, 
//          Google Drive archiving, and month-over-month analysis
//
// Features:
// - Priority-based scoring (������ / ���� / ��)
// - Status badges (🔴 URGENT | � REVIEW | � MONITOR)
// - Financial impact calculation ($ At Risk)
// - Google Drive monthly archiving
// - Month-over-month trend analysis
// - Conditional formatting with color coding
// - Severity scoring (1-5 scale)
// - Violation resolution tracking
// =====================================================================================================================

// ---------------------
// V2 CONSTANTS & CONFIG
// ---------------------
const V2_SHEET_NAME = "Violations V2";
const V2_DRIVE_FOLDER_ID = "1u28i_kcx9D-LQoSiOj08sKfEAZyc7uWN"; // Your Google Drive folder
const V2_ADMIN_EMAIL = "platformsolutionsadopshorizon@gmail.com";

// Color scheme for conditional formatting
const V2_COLORS = {
  URGENT_BG: "#cc0000",      // Dark red
  URGENT_TEXT: "#ffffff",    // White
  REVIEW_BG: "#ffd966",      // Yellow
  REVIEW_TEXT: "#000000",    // Black
  MONITOR_BG: "#93c47d",     // Light green
  MONITOR_TEXT: "#000000",   // Black
  PRIORITY_HIGH: "#f4cccc",  // Light red
  PRIORITY_MED: "#fff2cc",   // Light yellow
  PRIORITY_LOW: "#ffffff",   // White
  STALE_SEVERE: "#ea9999",   // Salmon red
  STALE_HIGH: "#f9cb9c",     // Orange
  STALE_MED: "#ffe599",      // Light yellow
  STALE_LOW: "#d9ead3"       // Light green
};

// V2 Headers (21 columns - includes billing breakdown)
const V2_HEADERS = [
  "Priority", "Status", "Owner (Ops)", "Network ID", "Network Name", "Advertiser",
  "Placement ID", "Placement Name", "Flight Dates", "Issue Category", "Issue Severity",
  "Specific Issue", "Impressions", "Clicks", "CTR %", "CPC Cost", "CPM Cost",
  "Days Stale", "Total Cost", "Overcharge", "Action Required"
];

// ---------------------
// V2 DASHBOARD GENERATION
// ---------------------
function generateViolationsV2Dashboard() {
  const startTime = Date.now();
  Logger.log("[V2] Starting dashboard generation...");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const violationsSheet = ss.getSheetByName("Violations");
  
  if (!violationsSheet) {
    SpreadsheetApp.getUi().alert("Error: Violations sheet not found. Run QA first.");
    return;
  }
  
  // Get or create V2 sheet
  let v2Sheet = ss.getSheetByName(V2_SHEET_NAME);
  if (v2Sheet) {
    v2Sheet.clear();
  } else {
    v2Sheet = ss.insertSheet(V2_SHEET_NAME);
  }
  
  // Load source data from Violations tab
  const violationsData = violationsSheet.getDataRange().getValues();
  if (violationsData.length < 2) {
    SpreadsheetApp.getUi().alert("No violations found to process.");
    return;
  }
  
  const vHeaders = violationsData[0];
  const vMap = getHeaderMap(vHeaders);
  
  // Load Networks sheet for Network Name lookup
  const networkNameMap = buildNetworkNameMap_();
  
  // Process each violation row and transform to V2 format
  const v2Data = [V2_HEADERS];
  
  for (let i = 1; i < violationsData.length; i++) {
    const row = violationsData[i];
    const v2Row = transformToV2Row_(row, vMap, networkNameMap);
    if (v2Row) v2Data.push(v2Row);
  }
  
  // Write data to V2 sheet
  if (v2Data.length > 1) {
    v2Sheet.getRange(1, 1, v2Data.length, V2_HEADERS.length).setValues(v2Data);
    
    // Apply formatting
    formatV2Sheet_(v2Sheet);
    applyV2ConditionalFormatting_(v2Sheet);
    
    // Freeze header and priority columns
    v2Sheet.setFrozenRows(1);
    v2Sheet.setFrozenColumns(3);
    
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    Logger.log(`[V2] � Dashboard generated with ${v2Data.length - 1} rows in ${elapsed}s`);
    
    SpreadsheetApp.getUi().alert(`� V2 Dashboard generated!\n\n${v2Data.length - 1} violations processed\nTime: ${elapsed}s`);
  } else {
    SpreadsheetApp.getUi().alert("No violations to display in V2 dashboard.");
  }
}

// ---------------------
// TRANSFORM ROW TO V2 FORMAT
// ---------------------
function transformToV2Row_(row, vMap, networkNameMap) {
  // Extract data from original Violations row
  const networkId = String(row[vMap["Network ID"]] || "");
  const networkName = networkNameMap[networkId] || networkId;
  const advertiser = String(row[vMap["Advertiser"]] || "");
  const placementId = String(row[vMap["Placement ID"]] || "");
  const placementName = String(row[vMap["Placement"]] || "");
  const placementStart = row[vMap["Placement Start Date"]];
  const placementEnd = row[vMap["Placement End Date"]];
  const reportDate = row[vMap["Report Date"]];
  const impressions = row[vMap["Impressions"]] || 0;
  const clicks = row[vMap["Clicks"]] || 0;
  const ctrStr = String(row[vMap["CTR (%)"]] || "0%");
  const ctr = parseFloat(ctrStr.replace("%", "")) || 0;
  const cpcStr = String(row[vMap["$CPC"]] || "$0");
  const cpc = parseFloat(cpcStr.replace("$", "")) || 0;
  const cpmStr = String(row[vMap["$CPM"]] || "$0");
  const cpm = parseFloat(cpmStr.replace("$", "")) || 0;
  const issueType = String(row[vMap["Issue Type"]] || "");
  const details = String(row[vMap["Details"]] || "");
  const lastImpChange = row[vMap["Last Imp Change"]];
  const lastClkChange = row[vMap["Last Click Change"]];
  const ownerOps = String(row[vMap["Owner (Ops)"]] || "Unassigned");
  
  // Calculate derived fields
  const flightDates = formatFlightDates_(placementStart, placementEnd, reportDate);
  const issueCategory = extractIssueCategory_(issueType);
  
  // Check if this is a click tracker/pixel (de-escalate if 0 impressions)
  const isTracker = isClickTrackerOrPixel_(placementName);
  
  const issueSeverity = calculateSeverityScore_(issueType, impressions, clicks, cpc, cpm, placementEnd, reportDate, isTracker);
  const priority = calculatePriority_(issueSeverity, issueCategory);
  const status = calculateStatus_(priority, issueSeverity, issueCategory, cpc, placementEnd, reportDate);
  const specificIssue = formatSpecificIssue_(issueType, details, impressions, clicks, cpc, cpm);
  const daysStale = calculateDaysStale_(lastImpChange, lastClkChange, reportDate);
  
  // Calculate billing costs using correct CPC/CPM rates
  const cpcCost = clicks * CPC_RATE;
  const cpmCost = (impressions / 1000) * CPM_RATE;
  
  // Total cost and overcharge calculation
  let totalCost = 0;
  let overcharge = 0;
  
  if (clicks > impressions) {
    // Billing error: Billed at CPM for impressions + CPC for excess clicks
    const excessClicks = clicks - impressions;
    overcharge = excessClicks * CPC_RATE;
    totalCost = cpmCost + overcharge;
  } else {
    // Normal: Billed at CPM only
    totalCost = cpmCost;
    overcharge = 0;
  }
  
  return [
    priority,           // Priority (������ / ���� / ��)
    status,             // Status (🔴/�/�)
    ownerOps,           // Owner (Ops)
    networkId,          // Network ID
    networkName,        // Network Name
    advertiser,         // Advertiser
    placementId,        // Placement ID
    placementName,      // Placement Name
    flightDates,        // Flight Dates (combined)
    issueCategory,      // Issue Category
    issueSeverity,      // Issue Severity (1-5)
    specificIssue,      // Specific Issue
    impressions,        // Impressions
    clicks,             // Clicks
    ctr + "%",          // CTR %
    "$" + cpcCost.toFixed(2), // CPC Cost (total)
    "$" + cpmCost.toFixed(2), // CPM Cost (total)
    daysStale,          // Days Stale
    "$" + totalCost.toFixed(2),   // Total Cost (actual bill)
    "$" + overcharge.toFixed(2),  // Overcharge (extra due to error)
    ""                  // Action Required (blank for manual entry)
  ];
}

// ---------------------
// HELPER: Build Network Name Map
// ---------------------
function buildNetworkNameMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const networksSheet = ss.getSheetByName("Networks");
  const map = {};
  
  if (!networksSheet || networksSheet.getLastRow() < 2) return map;
  
  const data = networksSheet.getDataRange().getValues();
  const hdr = data[0].map(h => String(h || "").trim().toLowerCase());
  
  // Find Network ID and Network Name columns
  const idCandidates = ["network id", "network_id", "networkid", "cm360 network id"];
  const nameCandidates = ["network name", "network_name", "networkname", "cm360 network name", "name"];
  
  let idIdx = -1, nameIdx = -1;
  
  for (let i = 0; i < hdr.length; i++) {
    const h = hdr[i];
    if (idIdx === -1 && idCandidates.some(c => h.includes(c))) idIdx = i;
    if (nameIdx === -1 && nameCandidates.some(c => h.includes(c))) nameIdx = i;
  }
  
  if (idIdx === -1 || nameIdx === -1) return map;
  
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][idIdx] || "").trim();
    const name = String(data[i][nameIdx] || "").trim();
    if (id && name) map[id] = name;
  }
  
  Logger.log(`[V2] Loaded ${Object.keys(map).length} network names`);
  return map;
}

// ---------------------
// HELPER: Calculate Google Billing (Dual Methodology)
// ---------------------
function calculateGoogleBilling_(imp, clk, cpc, cpm) {
  let expectedCost = 0;
  let actualCost = 0;
  let overcharge = 0;
  
  // Google's billing methodology:
  // 1. CPC only (no CPM): Bill all clicks at CPC
  // 2. CPM only (no CPC): Bill all impressions at CPM
  // 3. Both present + Impressions > Clicks: Bill impressions at CPM (normal)
  // 4. Both present + Clicks > Impressions: Bill impressions at CPM + excess clicks at CPC (RISK!)
  
  const hasCPC = cpc > 0;
  const hasCPM = cpm > 0;
  
  if (!hasCPC && !hasCPM) {
    // No pricing - no cost
    return { expectedCost: 0, actualCost: 0, overcharge: 0 };
  }
  
  if (hasCPC && !hasCPM) {
    // CPC only billing
    actualCost = clk * cpc;
    expectedCost = actualCost; // This is normal
    overcharge = 0;
  } else if (hasCPM && !hasCPC) {
    // CPM only billing
    actualCost = (imp / 1000) * cpm;
    expectedCost = actualCost; // This is normal
    overcharge = 0;
  } else if (hasCPC && hasCPM) {
    // Both metrics present - check for dual billing scenario
    if (imp >= clk) {
      // Normal: Impressions >= Clicks, billed at CPM
      actualCost = (imp / 1000) * cpm;
      expectedCost = actualCost;
      overcharge = 0;
    } else {
      // BILLING RISK: Clicks > Impressions
      // Expected: Should only pay CPM for impressions
      expectedCost = (imp / 1000) * cpm;
      
      // Actual: Google bills CPM for impressions + CPC for excess clicks
      const cpmCost = (imp / 1000) * cpm;
      const excessClicks = clk - imp;
      const cpcCost = excessClicks * cpc;
      actualCost = cpmCost + cpcCost;
      
      // Overcharge = the extra CPC charge
      overcharge = cpcCost;
    }
  }
  
  return {
    expectedCost: expectedCost,
    actualCost: actualCost,
    overcharge: overcharge
  };
}

// ---------------------
// HELPER: Format Flight Dates
// ---------------------
function formatFlightDates_(startDate, endDate, reportDate) {
  const start = startDate instanceof Date ? startDate : new Date(startDate);
  const end = endDate instanceof Date ? endDate : new Date(endDate);
  const report = reportDate instanceof Date ? reportDate : new Date(reportDate);
  
  if (isNaN(end)) return "Unknown";
  
  const isExpired = end < report;
  const startStr = isNaN(start) ? "?" : Utilities.formatDate(start, Session.getScriptTimeZone(), "M/d");
  const endStr = Utilities.formatDate(end, Session.getScriptTimeZone(), "M/d");
  
  if (isExpired) {
    return `ENDED ${endStr}`;
  } else {
    return `${startStr} - ${endStr}`;
  }
}

// ---------------------
// HELPER: Detect Click Tracker/Impression Pixel
// ---------------------
function isClickTrackerOrPixel_(placementName) {
  if (!placementName) return false;
  
  const name = normalizeName_(placementName);
  
  // Compile patterns if needed
  compileLPPatternsIfNeeded_();
  
  // Check for click tracker or impression pixel patterns
  for (let i = 0; i < _lpCompiled.length; i++) {
    const p = _lpCompiled[i];
    if (!p.enabled || !p.re) continue;
    
    const cat = p.category;
    if ((cat === 'Click Tracker' || cat === 'Impression Pixel/Beacon') && p.re.test(name)) {
      return true;
    }
  }
  
  return false;
}

// ---------------------
// HELPER: Extract Issue Category
// ---------------------
function extractIssueCategory_(issueType) {
  const types = issueType.toUpperCase();
  if (types.includes("BILLING")) return "BILLING";
  if (types.includes("DELIVERY")) return "DELIVERY";
  if (types.includes("PERFORMANCE")) return "PERFORMANCE";
  if (types.includes("COST")) return "COST";
  return "OTHER";
}

// ---------------------
// HELPER: Calculate Severity Score (1-5)
// ---------------------
function calculateSeverityScore_(issueType, imp, clk, cpc, cpm, placementEnd, reportDate, isTracker) {
  const types = issueType.toUpperCase();
  const end = placementEnd instanceof Date ? placementEnd : new Date(placementEnd);
  const report = reportDate instanceof Date ? reportDate : new Date(reportDate);
  const isExpired = !isNaN(end) && end < report;
  
  // De-escalate click trackers/pixels with 0 impressions
  if (isTracker && imp === 0) {
    // Click trackers with 0 impressions are expected behavior
    // Downgrade to severity 1 (INFO) regardless of issue type
    return 1;
  }
  
  // 5 = CRITICAL: Billing risk with both metrics + clicks > impressions
  if (types.includes("BILLING") && clk > imp && cpc > 0 && cpm > 0) return 5;
  
  // 5 = CRITICAL: Expired CPC risk with high cost
  if (types.includes("EXPIRED CPC RISK") && cpc > 20) return 5;
  
  // 4 = HIGH: Active billing risk
  if (types.includes("ACTIVE CPC") && clk > imp) return 4;
  
  // 4 = HIGH: Extreme performance (CTR �� 90% + CPM �� $10)
  if (types.includes("PERFORMANCE") && cpm >= 10) return 4;
  
  // 3 = MEDIUM: Recently expired with activity
  if (types.includes("RECENTLY EXPIRED") || (types.includes("DELIVERY") && isExpired)) return 3;
  
  // 3 = MEDIUM: High cost issues
  if (types.includes("COST") && (cpc > 10 || cpm > 10)) return 3;
  
  // 2 = LOW: Cost-only issues
  if (types.includes("COST")) return 2;
  
  // 1 = INFO: Everything else
  return 1;
}

// ---------------------
// HELPER: Calculate Priority
// ---------------------
function calculatePriority_(severity, category) {
  if (severity >= 4) return "������";
  if (severity === 3) return "����";
  return "��";
}

// ---------------------
// HELPER: Calculate Status
// ---------------------
function calculateStatus_(priority, severity, category, cpc, placementEnd, reportDate) {
  const end = placementEnd instanceof Date ? placementEnd : new Date(placementEnd);
  const report = reportDate instanceof Date ? reportDate : new Date(reportDate);
  const isExpired = !isNaN(end) && end < report;
  
  // 🔴 URGENT: High priority + severe conditions
  if (priority === "������" && (category === "BILLING" || cpc > 20 || (isExpired && severity >= 4))) {
    return "🔴 URGENT";
  }
  
  // � REVIEW: Medium priority or specific categories
  if (priority === "����" || category === "PERFORMANCE" || category === "DELIVERY") {
    return "� REVIEW";
  }
  
  // � MONITOR: Everything else
  return "� MONITOR";
}

// ---------------------
// HELPER: Format Specific Issue
// ---------------------
function formatSpecificIssue_(issueType, details, imp, clk, cpc, cpm) {
  const types = issueType.split(", ");
  const primary = types[0] || issueType;
  
  // Extract key details and format concisely
  if (primary.includes("BILLING")) {
    if (clk > imp) {
      return `CPC Billing Risk: Clicks (${clk}) > Impr (${imp}), $CPC=$${cpc.toFixed(2)}`;
    }
  }
  
  if (primary.includes("PERFORMANCE")) {
    const ctrMatch = details.match(/CTR = ([\d.]+)%/);
    const ctr = ctrMatch ? ctrMatch[1] : "N/A";
    return `Extreme Performance: CTR=${ctr}%, $CPM=$${cpm.toFixed(2)}`;
  }
  
  if (primary.includes("DELIVERY")) {
    const dateMatch = details.match(/Ended ([\d/]+)/);
    const endDate = dateMatch ? dateMatch[1] : "Unknown";
    return `Post-Flight Activity: Ended ${endDate}, still serving`;
  }
  
  if (primary.includes("COST")) {
    if (cpc > 10 && cpm === 0) {
      return `High CPC Only: $CPC=$${cpc.toFixed(2)} (no CPM)`;
    }
    if (cpm > 10 && cpc === 0) {
      return `High CPM Only: $CPM=$${cpm.toFixed(2)} (no CPC)`;
    }
    if (cpc > 10 && cpm > 10) {
      return `Both Metrics High: $CPC=$${cpc.toFixed(2)}, $CPM=$${cpm.toFixed(2)}`;
    }
  }
  
  // Fallback: use first detail snippet
  const detailParts = details.split(" | ");
  return detailParts[0] || primary.replace(/�|�|�|�/g, "").trim();
}

// ---------------------
// HELPER: Calculate Days Stale
// ---------------------
function calculateDaysStale_(lastImpChange, lastClkChange, reportDate) {
  const report = reportDate instanceof Date ? reportDate : new Date(reportDate);
  
  const impDays = (lastImpChange && !isNaN(lastImpChange)) ? lastImpChange : 999;
  const clkDays = (lastClkChange && !isNaN(lastClkChange)) ? lastClkChange : 999;
  
  const stale = Math.max(impDays, clkDays);
  
  if (stale === 999) return "";
  return stale;
}

// ---------------------
// HELPER: Calculate Financial Impact ($ At Risk)
// ---------------------
function calculateFinancialImpact_(issueType, imp, clk, cpc, cpm, placementEnd, reportDate, billingCalc) {
  const types = issueType.toUpperCase();
  let atRisk = 0;
  
  // BILLING RISK: Use the overcharge from Google's dual billing calculation
  if (types.includes("BILLING") && billingCalc && billingCalc.overcharge > 0) {
    atRisk += billingCalc.overcharge;
  }
  
  // PERFORMANCE WASTE: Potential bot traffic (CTR �� 90% + high CPM)
  if (types.includes("PERFORMANCE") && cpm >= 10) {
    // Estimate 50% of impressions as potential waste
    atRisk += (imp * 0.5) * (cpm / 1000);
  }
  
  // POST-FLIGHT OVERSPEND: Spending after placement end date
  const end = placementEnd instanceof Date ? placementEnd : new Date(placementEnd);
  const report = reportDate instanceof Date ? reportDate : new Date(reportDate);
  if (types.includes("DELIVERY") && !isNaN(end) && end < report) {
    // Estimate all current spend is post-flight waste
    // Use actual cost from billing calculation if available
    if (billingCalc && billingCalc.actualCost > 0) {
      atRisk += billingCalc.actualCost;
    } else if (cpm > 0) {
      atRisk += (imp * cpm) / 1000;
    } else if (cpc > 0) {
      atRisk += clk * cpc;
    }
  }
  
  // HIGH COST: Flag but don't add to at-risk (already planned spend)
  // Just monitoring, not financial risk
  
  return atRisk;
}

// ---------------------
// FORMAT V2 SHEET
// ---------------------
function formatV2Sheet_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, V2_HEADERS.length);
  
  // Header formatting
  headerRange
    .setBackground("#4a86e8")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  
  // Set column widths
  sheet.setColumnWidth(1, 80);   // Priority
  sheet.setColumnWidth(2, 100);  // Status
  sheet.setColumnWidth(3, 120);  // Owner
  sheet.setColumnWidth(4, 90);   // Network ID
  sheet.setColumnWidth(5, 150);  // Network Name
  sheet.setColumnWidth(6, 150);  // Advertiser
  sheet.setColumnWidth(7, 100);  // Placement ID
  sheet.setColumnWidth(8, 250);  // Placement Name
  sheet.setColumnWidth(9, 120);  // Flight Dates
  sheet.setColumnWidth(10, 100); // Issue Category
  sheet.setColumnWidth(11, 90);  // Issue Severity
  sheet.setColumnWidth(12, 350); // Specific Issue
  sheet.setColumnWidth(13, 90);  // Impressions
  sheet.setColumnWidth(14, 80);  // Clicks
  sheet.setColumnWidth(15, 80);  // CTR %
  sheet.setColumnWidth(16, 100); // CPC Cost
  sheet.setColumnWidth(17, 100); // CPM Cost
  sheet.setColumnWidth(18, 90);  // Days Stale
  sheet.setColumnWidth(19, 110); // Total Cost
  sheet.setColumnWidth(20, 110); // Overcharge
  sheet.setColumnWidth(21, 150); // Action Required
  
  // Auto-resize row heights
  sheet.setRowHeights(2, sheet.getMaxRows() - 1, 21);
  
  Logger.log("[V2] Sheet formatting applied");
}

// ---------------------
// APPLY CONDITIONAL FORMATTING
// ---------------------
function applyV2ConditionalFormatting_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  // Priority column (A) - Background colors
  const priorityRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("������")
      .setBackground(V2_COLORS.PRIORITY_HIGH)
      .setRanges([sheet.getRange(2, 1, lastRow - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("����")
      .setBackground(V2_COLORS.PRIORITY_MED)
      .setRanges([sheet.getRange(2, 1, lastRow - 1, 1)])
      .build()
  ];
  
  // Status column (B) - Background + text colors
  const statusRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("🔴 URGENT")
      .setBackground(V2_COLORS.URGENT_BG)
      .setFontColor(V2_COLORS.URGENT_TEXT)
      .setRanges([sheet.getRange(2, 2, lastRow - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("� REVIEW")
      .setBackground(V2_COLORS.REVIEW_BG)
      .setFontColor(V2_COLORS.REVIEW_TEXT)
      .setRanges([sheet.getRange(2, 2, lastRow - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("� MONITOR")
      .setBackground(V2_COLORS.MONITOR_BG)
      .setFontColor(V2_COLORS.MONITOR_TEXT)
      .setRanges([sheet.getRange(2, 2, lastRow - 1, 1)])
      .build()
  ];
  
  // Days Stale column (R/18) - Color gradient
  const staleRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(15)
      .setBackground(V2_COLORS.STALE_SEVERE)
      .setRanges([sheet.getRange(2, 18, lastRow - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(8, 14)
      .setBackground(V2_COLORS.STALE_HIGH)
      .setRanges([sheet.getRange(2, 18, lastRow - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(4, 7)
      .setBackground(V2_COLORS.STALE_MED)
      .setRanges([sheet.getRange(2, 18, lastRow - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0, 3)
      .setBackground(V2_COLORS.STALE_LOW)
      .setRanges([sheet.getRange(2, 18, lastRow - 1, 1)])
      .build()
  ];
  
  // Flight Dates column (I/9) - Highlight expired
  const flightRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("ENDED")
      .setFontColor("#cc0000")
      .setBold(true)
      .setRanges([sheet.getRange(2, 9, lastRow - 1, 1)])
      .build()
  ];
  
  // Overcharge column (T/20) - Highlight any overcharges > $0
  const overchargeRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=VALUE(SUBSTITUTE(T2,"$",""))>0')
      .setBackground("#f4cccc")  // Light red
      .setBold(true)
      .setRanges([sheet.getRange(2, 20, lastRow - 1, 1)])
      .build()
  ];
  
  // Combine all rules
  const allRules = [].concat(priorityRules, statusRules, staleRules, flightRules, overchargeRules);
  sheet.setConditionalFormatRules(allRules);
  
  Logger.log(`[V2] Applied ${allRules.length} conditional formatting rules`);
}

// ---------------------
// EXPORT V2 TO GOOGLE DRIVE
// ---------------------
function exportV2ToDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const v2Sheet = ss.getSheetByName(V2_SHEET_NAME);
  
  if (!v2Sheet || v2Sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("Error: V2 Dashboard is empty. Generate it first.");
    return;
  }
  
  try {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    const monthName = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM");
    
    // Create folder structure: YYYY-MM-MonthName
    const parentFolder = DriveApp.getFolderById(V2_DRIVE_FOLDER_ID);
    const monthFolderName = `${year}-${month}-${monthName}`;
    
    let monthFolder;
    const existingFolders = parentFolder.getFoldersByName(monthFolderName);
    if (existingFolders.hasNext()) {
      monthFolder = existingFolders.next();
    } else {
      monthFolder = parentFolder.createFolder(monthFolderName);
      Logger.log(`[V2] Created folder: ${monthFolderName}`);
    }
    
    // Export as XLSX
    const fileName = `Violations_V2_${year}-${month}-${day}.xlsx`;
    const xlsxBlob = createXLSXFromSheet(v2Sheet);
    xlsxBlob.setName(fileName);
    
    // Delete old file with same name if exists
    const existingFiles = monthFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }
    
    // Create new file
    const file = monthFolder.createFile(xlsxBlob);
    const fileUrl = file.getUrl();
    
    Logger.log(`[V2] � Exported to Drive: ${fileUrl}`);
    SpreadsheetApp.getUi().alert(`� V2 Dashboard exported to Google Drive!\n\nFile: ${fileName}\nFolder: ${monthFolderName}\n\nURL: ${fileUrl}`);
    
    return fileUrl;
    
  } catch (error) {
    Logger.log("[V2] �� Export failed: " + error);
    SpreadsheetApp.getUi().alert("�� Export failed:\n\n" + error);
    return null;
  }
}

// ---------------------
// MONTHLY SUMMARY REPORT
// ---------------------
function generateMonthlySummaryReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const v2Sheet = ss.getSheetByName(V2_SHEET_NAME);
  
  if (!v2Sheet || v2Sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("Error: V2 Dashboard is empty. Generate it first.");
    return;
  }
  
  const data = v2Sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  // Build header map
  const hMap = {};
  headers.forEach((h, i) => { hMap[h] = i; });
  
  // Calculate statistics
  let totalViolations = rows.length;
  let urgentCount = 0;
  let reviewCount = 0;
  let monitorCount = 0;
  let totalAtRisk = 0;
  
  const categoryBreakdown = {};
  const ownerBreakdown = {};
  const severityBreakdown = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  
  rows.forEach(row => {
    const status = String(row[hMap["Status"]] || "");
    const category = String(row[hMap["Issue Category"]] || "OTHER");
    const owner = String(row[hMap["Owner (Ops)"]] || "Unassigned");
    const severity = row[hMap["Issue Severity"]] || 1;
    const overchargeStr = String(row[hMap["Overcharge"]] || "$0");
    const overcharge = parseFloat(overchargeStr.replace("$", "")) || 0;
    
    if (status.includes("🔴")) urgentCount++;
    else if (status.includes("�")) reviewCount++;
    else if (status.includes("�")) monitorCount++;
    
    totalAtRisk += overcharge; // Use overcharge instead of old "$ At Risk"
    
    categoryBreakdown[category] = (categoryBreakdown[category] || 0) + 1;
    ownerBreakdown[owner] = (ownerBreakdown[owner] || 0) + 1;
    severityBreakdown[severity] = (severityBreakdown[severity] || 0) + 1;
  });
  
  // Create summary sheet
  let summarySheet = ss.getSheetByName("Monthly Summary");
  if (summarySheet) {
    summarySheet.clear();
  } else {
    summarySheet = ss.insertSheet("Monthly Summary");
  }
  
  const today = new Date();
  const monthName = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM yyyy");
  
  // Build summary data
  const summaryData = [
    ["CM360 QA Monthly Summary Report"],
    ["Generated:", Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM dd, yyyy HH:mm")],
    ["Month:", monthName],
    [""],
    ["📊 OVERVIEW"],
    ["Total Violations:", totalViolations],
    ["🔴 Urgent:", urgentCount],
    ["� Review:", reviewCount],
    ["� Monitor:", monitorCount],
    ["💰 Total $ At Risk:", "$" + totalAtRisk.toFixed(2)],
    [""],
    ["📂 BY CATEGORY"],
  ];
  
  Object.keys(categoryBreakdown).sort().forEach(cat => {
    summaryData.push([cat, categoryBreakdown[cat]]);
  });
  
  summaryData.push([""]);
  summaryData.push(["👥 BY OWNER"]);
  
  Object.keys(ownerBreakdown).sort().forEach(owner => {
    summaryData.push([owner, ownerBreakdown[owner]]);
  });
  
  summaryData.push([""]);
  summaryData.push(["�� BY SEVERITY"]);
  for (let i = 5; i >= 1; i--) {
    const stars = i >= 4 ? "������" : i === 3 ? "����" : "��";
    summaryData.push([`${i} - ${stars}`, severityBreakdown[i]]);
  }
  
  // Write to sheet
  summarySheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);
  
  // Format
  summarySheet.getRange(1, 1, 1, 2).merge().setBackground("#4a86e8").setFontColor("#ffffff").setFontWeight("bold").setFontSize(14);
  summarySheet.setColumnWidth(1, 200);
  summarySheet.setColumnWidth(2, 150);
  
  Logger.log("[V2] � Monthly summary generated");
  SpreadsheetApp.getUi().alert(`� Monthly Summary Report Generated!\n\nTotal Violations: ${totalViolations}\n🔴 Urgent: ${urgentCount}\n💰 At Risk: $${totalAtRisk.toFixed(2)}`);
}

// ---------------------
// MONTH-OVER-MONTH ANALYSIS
// ---------------------
function runMonthOverMonthAnalysis() {
  SpreadsheetApp.getUi().alert("📈 Month-over-Month Analysis\n\nThis feature tracks trends by comparing archived monthly reports.\n\nComing soon: Automatically compare violation counts, $ at risk, and resolution rates across months.");
  
  // TODO: Implement full MoM analysis
  // - Load previous month's archived V2 file from Drive
  // - Compare violation counts by category
  // - Track $ at risk trends
  // - Calculate resolution rate (violations disappeared)
  // - Generate trend charts
}

// ---------------------
// DISPLAY FINANCIAL IMPACT
// ---------------------
function displayFinancialImpact() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const v2Sheet = ss.getSheetByName(V2_SHEET_NAME);
  
  if (!v2Sheet || v2Sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("Error: V2 Dashboard is empty. Generate it first.");
    return;
  }
  
  const data = v2Sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  const hMap = {};
  headers.forEach((h, i) => { hMap[h] = i; });
  
  let totalAtRisk = 0;
  let totalOvercharge = 0;
  let billingRisk = 0;
  let performanceWaste = 0;
  let postFlightSpend = 0;
  
  rows.forEach(row => {
    const category = String(row[hMap["Issue Category"]] || "");
    const overchargeStr = String(row[hMap["Overcharge"]] || "$0");
    const overcharge = parseFloat(overchargeStr.replace("$", "")) || 0;
    
    totalOvercharge += overcharge;
    
    if (category === "BILLING") billingRisk += overcharge;
    if (category === "PERFORMANCE") performanceWaste += overcharge;
    if (category === "DELIVERY") postFlightSpend += overcharge;
  });
  
  const message = `💰 FINANCIAL IMPACT ANALYSIS\n\n` +
    `Total Overcharge (Billing Risk): $${totalOvercharge.toFixed(2)}\n\n` +
    `Breakdown:\n` +
    `  �� Billing Overcharge: $${billingRisk.toFixed(2)}\n` +
    `  �� Performance Waste: $${performanceWaste.toFixed(2)}\n` +
    `  �� Post-Flight Spend: $${postFlightSpend.toFixed(2)}\n\n` +
    `Note: Billing Overcharge shows the dual billing impact where\n` +
    `Google charges CPM for impressions + CPC for excess clicks.\n\n` +
    `This represents potential savings from catching and resolving these violations.`;
  
  SpreadsheetApp.getUi().alert(message);
  Logger.log(`[V2] Financial Impact - Total Overcharge: $${totalOvercharge.toFixed(2)}`);
}


// =====================================================================================================================
// ============================================ END V2 DASHBOARD SYSTEM ================================================
// =====================================================================================================================


// =====================================================================================================================
// ======================================== HISTORICAL ARCHIVE SYSTEM ==================================================
// =====================================================================================================================

// ---------------------
// CONSTANTS
// ---------------------
const ARCHIVE_FOLDER_ID = '1u28i_kcx9D-LQoSiOj08sKfEAZyc7uWN'; // Same as V2 exports
const GMAIL_SEARCH_SUBJECT = 'CM360 CPC/CPM FLIGHT QA';
const BATCH_SIZE = 25; // Process 25 emails per execution (safe for 6-min limit)

// ---------------------
// MAIN: Archive All Historical Reports (April-November 2025)
// ---------------------
function archiveAllHistoricalReports() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Archive Historical Reports',
    'This will process all CM360 QA reports from April-November 2025.\n\n' +
    'Expected: ~128 emails (8 months � 16 days)\n' +
    'Processing: 25 emails per run\n' +
    'You will receive email updates after each batch.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Archive cancelled.');
    return;
  }
  
  // Initialize archive state
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ARCHIVE_STATE', JSON.stringify({
    status: 'running',
    currentMonth: 4, // Start with April
    currentYear: 2025,
    emailsProcessed: 0,
    attachmentsSaved: 0,
    startTime: new Date().toISOString()
  }));
  
  // Start processing
  processNextBatch_();
}

// ---------------------
// Archive Single Month (User selects)
// ---------------------
function archiveSingleMonth() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Archive Single Month',
    'Enter month to archive (1-12):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const month = parseInt(response.getResponseText());
  if (isNaN(month) || month < 1 || month > 12) {
    ui.alert('Invalid month. Please enter 1-12.');
    return;
  }
  
  const yearResponse = ui.prompt(
    'Archive Single Month',
    'Enter year (2025):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (yearResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const year = parseInt(yearResponse.getResponseText());
  if (isNaN(year)) {
    ui.alert('Invalid year.');
    return;
  }
  
  // Process single month
  const stats = processSingleMonthArchive_(year, month);
  
  ui.alert(
    'Archive Complete',
    `Processed ${stats.emailsProcessed} emails\n` +
    `Saved ${stats.attachmentsSaved} attachments\n\n` +
    `Folder: Historical Violation Reports/${year}/${String(month).padStart(2, '0')}-${getMonthName_(month)}`,
    ui.ButtonSet.OK
  );
}

// ---------------------
// View Archive Progress
// ---------------------
function viewArchiveProgress() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('ARCHIVE_STATE');
  
  const ui = SpreadsheetApp.getUi();
  
  if (!stateJson) {
    ui.alert('Archive Progress', 'No archive in progress.', ui.ButtonSet.OK);
    return;
  }
  
  const state = JSON.parse(stateJson);
  const monthName = getMonthName_(state.currentMonth);
  
  ui.alert(
    'Archive Progress',
    `Status: ${state.status}\n` +
    `Current: ${monthName} ${state.currentYear}\n` +
    `Emails processed: ${state.emailsProcessed}\n` +
    `Attachments saved: ${state.attachmentsSaved}\n` +
    `Started: ${new Date(state.startTime).toLocaleString()}`,
    ui.ButtonSet.OK
  );
}

// ---------------------
// Resume Archive (if interrupted)
// ---------------------
function resumeArchive() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('ARCHIVE_STATE');
  
  const ui = SpreadsheetApp.getUi();
  
  if (!stateJson) {
    ui.alert('Resume Archive', 'No archive in progress to resume.', ui.ButtonSet.OK);
    return;
  }
  
  const state = JSON.parse(stateJson);
  
  if (state.status === 'completed') {
    ui.alert('Resume Archive', 'Archive already completed.', ui.ButtonSet.OK);
    return;
  }
  
  const response = ui.alert(
    'Resume Archive',
    `Resume from ${getMonthName_(state.currentMonth)} ${state.currentYear}?\n\n` +
    `Progress: ${state.emailsProcessed} emails, ${state.attachmentsSaved} attachments`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  processNextBatch_();
}

// ---------------------
// INTERNAL: Process Next Batch (recursive until complete)
// ---------------------
function processNextBatch_() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('ARCHIVE_STATE');
  
  if (!stateJson) {
    Logger.log('No archive state found');
    return;
  }
  
  const state = JSON.parse(stateJson);
  
  // Check if we're done (November = month 11)
  if (state.currentMonth > 11) {
    state.status = 'completed';
    props.setProperty('ARCHIVE_STATE', JSON.stringify(state));
    
    // Send completion email
    MailApp.sendEmail({
      to: 'platformsolutionsadopshorizon@gmail.com',
      subject: '� CM360 Historical Archive Complete',
      htmlBody: `<h3>Archive Complete</h3>
        <p><strong>Total emails processed:</strong> ${state.emailsProcessed}</p>
        <p><strong>Total attachments saved:</strong> ${state.attachmentsSaved}</p>
        <p><strong>Duration:</strong> ${new Date(state.startTime).toLocaleString()} - ${new Date().toLocaleString()}</p>
        <p><strong>Location:</strong> <a href="https://drive.google.com/drive/folders/${ARCHIVE_FOLDER_ID}">Historical Violation Reports</a></p>`
    });
    
    return;
  }
  
  // Process current month
  try {
    const monthStats = processSingleMonthArchive_(state.currentYear, state.currentMonth);
    
    state.emailsProcessed += monthStats.emailsProcessed;
    state.attachmentsSaved += monthStats.attachmentsSaved;
    state.currentMonth++;
    
    props.setProperty('ARCHIVE_STATE', JSON.stringify(state));
    
    // Send progress email
    const monthName = getMonthName_(state.currentMonth - 1);
    MailApp.sendEmail({
      to: 'platformsolutionsadopshorizon@gmail.com',
      subject: `📁 CM360 Archive: ${monthName} ${state.currentYear} Complete`,
      htmlBody: `<h3>${monthName} ${state.currentYear} Archived</h3>
        <p><strong>Emails:</strong> ${monthStats.emailsProcessed}</p>
        <p><strong>Attachments:</strong> ${monthStats.attachmentsSaved}</p>
        <p><strong>Total progress:</strong> ${state.emailsProcessed} emails, ${state.attachmentsSaved} attachments</p>
        <p><strong>Next:</strong> ${getMonthName_(state.currentMonth)} ${state.currentYear}</p>`
    });
    
    // Continue to next month
    processNextBatch_();
    
  } catch (error) {
    Logger.log('Error processing batch: ' + error);
    
    // Send error email
    MailApp.sendEmail({
      to: 'platformsolutionsadopshorizon@gmail.com',
      subject: '��️ CM360 Archive Error',
      htmlBody: `<h3>Archive Error</h3>
        <p><strong>Month:</strong> ${getMonthName_(state.currentMonth)} ${state.currentYear}</p>
        <p><strong>Error:</strong> ${error}</p>
        <p><strong>Progress:</strong> ${state.emailsProcessed} emails, ${state.attachmentsSaved} attachments</p>
        <p>Use "Resume Archive" to continue.</p>`
    });
  }
}

// ---------------------
// INTERNAL: Process Single Month Archive
// ---------------------
function processSingleMonthArchive_(year, month) {
  const monthStr = String(month).padStart(2, '0');
  const monthName = getMonthName_(month);
  
  // Search Gmail for this month's reports
  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 0, 23, 59, 59); // Last day of month
  
  const query = `subject:"${GMAIL_SEARCH_SUBJECT}" after:${startDate.getTime()/1000} before:${endDate.getTime()/1000}`;
  
  const threads = GmailApp.search(query, 0, BATCH_SIZE);
  
  let emailsProcessed = 0;
  let attachmentsSaved = 0;
  
  // Get or create Drive folder
  const monthFolder = getOrCreateArchiveFolder_(year, month);
  
  for (const thread of threads) {
    const messages = thread.getMessages();
    
    for (const message of messages) {
      const subject = message.getSubject();
      const date = extractDateFromSubject_(subject);
      
      if (!date) {
        Logger.log(`Could not extract date from subject: ${subject}`);
        continue;
      }
      
      // Process attachments
      const attachments = message.getAttachments();
      
      for (const attachment of attachments) {
        const filename = attachment.getName();
        const lowerFilename = filename.toLowerCase();
        
        // Handle ZIP files
        if (lowerFilename.endsWith('.zip')) {
          const zipBlob = attachment.copyBlob();
          const unzipped = Utilities.unzip(zipBlob);
          
          for (const file of unzipped) {
            const unzippedName = file.getName().toLowerCase();
            if (unzippedName.endsWith('.csv') || unzippedName.endsWith('.xlsx')) {
              saveAttachmentToDrive_(file, monthFolder, date);
              attachmentsSaved++;
            }
          }
        }
        // Handle CSV and XLSX files
        else if (lowerFilename.endsWith('.csv') || lowerFilename.endsWith('.xlsx')) {
          saveAttachmentToDrive_(attachment, monthFolder, date);
          attachmentsSaved++;
        }
      }
      
      emailsProcessed++;
    }
  }
  
  return {
    emailsProcessed: emailsProcessed,
    attachmentsSaved: attachmentsSaved
  };
}

// ---------------------
// INTERNAL: Get or Create Archive Folder Structure
// ---------------------
function getOrCreateArchiveFolder_(year, month) {
  const monthStr = String(month).padStart(2, '0');
  const monthName = getMonthName_(month);
  
  const rootFolder = DriveApp.getFolderById(ARCHIVE_FOLDER_ID);
  
  // Get or create "Historical Violation Reports" folder
  let histFolder;
  const histFolders = rootFolder.getFoldersByName('Historical Violation Reports');
  if (histFolders.hasNext()) {
    histFolder = histFolders.next();
  } else {
    histFolder = rootFolder.createFolder('Historical Violation Reports');
  }
  
  // Get or create year folder
  let yearFolder;
  const yearFolders = histFolder.getFoldersByName(String(year));
  if (yearFolders.hasNext()) {
    yearFolder = yearFolders.next();
  } else {
    yearFolder = histFolder.createFolder(String(year));
  }
  
  // Get or create month folder
  let monthFolder;
  const monthFolderName = `${monthStr}-${monthName}`;
  const monthFolders = yearFolder.getFoldersByName(monthFolderName);
  if (monthFolders.hasNext()) {
    monthFolder = monthFolders.next();
  } else {
    monthFolder = yearFolder.createFolder(monthFolderName);
  }
  
  return monthFolder;
}

// ---------------------
// INTERNAL: Save Attachment to Drive
// ---------------------
function saveAttachmentToDrive_(attachment, folder, date) {
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const originalName = attachment.getName();
  const extension = originalName.substring(originalName.lastIndexOf('.'));
  const filename = `CM360_Report_${dateStr}${extension}`;
  
  // Check if file already exists
  const existingFiles = folder.getFilesByName(filename);
  if (existingFiles.hasNext()) {
    Logger.log(`File already exists: ${filename}`);
    return;
  }
  
  // Create file
  folder.createFile(attachment.copyBlob().setName(filename));
  Logger.log(`Saved: ${filename}`);
}

// ---------------------
// INTERNAL: Extract Date from Email Subject
// ---------------------
function extractDateFromSubject_(subject) {
  // Remove any prefixes (RE:, Fwd:, Automatic reply:, etc.)
  const cleanSubject = subject.replace(/^(RE:|FWD:|Automatic reply:)\s*/i, '').trim();
  
  // Try format: "MM.DD.YY" (e.g., "11.25.25")
  let match = cleanSubject.match(/(\d{1,2})\.(\d{1,2})\.(\d{2})/);
  
  if (match) {
    const month = parseInt(match[1]) - 1; // JS months are 0-indexed
    const day = parseInt(match[2]);
    const year = 2000 + parseInt(match[3]); // Assuming 20xx
    return new Date(year, month, day);
  }
  
  // Try format: "M/D/YY" (e.g., "4/30/25")
  match = cleanSubject.match(/(\d{1,2})\/(\d{1,2})\/(\d{2})/);
  
  if (match) {
    const month = parseInt(match[1]) - 1; // JS months are 0-indexed
    const day = parseInt(match[2]);
    const year = 2000 + parseInt(match[3]); // Assuming 20xx
    return new Date(year, month, day);
  }
  
  return null;
}

// ---------------------
// INTERNAL: Get Month Name
// ---------------------
function getMonthName_(month) {
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                  'July', 'August', 'September', 'October', 'November', 'December'];
  return months[month - 1];
}

// =====================================================================================================================
// ======================================= END HISTORICAL ARCHIVE SYSTEM ===============================================
// =====================================================================================================================


// =====================================================================================================================
// =========================================== RAW DATA ARCHIVE SYSTEM =================================================
// =====================================================================================================================

// ---------------------
// CONSTANTS
// ---------------------
const RAW_DATA_FOLDER_ID = '1u28i_kcx9D-LQoSiOj08sKfEAZyc7uWN'; // Same root as other archives
const RAW_DATA_SEARCH_SUBJECT = 'CM360 CPC/CPM FLIGHT QA';
const RAW_BATCH_SIZE = 100; // Process 100 emails per execution (increased from 20)
const RAW_SEARCH_MAX = 500; // Gmail's max threads per search

// ---------------------
// MAIN: Archive All Raw Data (Complete Inbox)
// ---------------------
function archiveAllRawData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Archive Raw Data',
    'This will save ALL raw data files from your CM360 inbox.\n\n' +
    'Strategy: Retrieve ALL emails with subject "BKCM360 Global QA Check"\n' +
    'All CSV/ZIP attachments will be extracted and saved.\n' +
    'Files organized by date automatically.\n\n' +
    'Expected: 100 emails per batch (auto-resumes every 10 min)\n' +
    'Process runs in background until complete.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Archive cancelled.');
    return;
  }
  
  // Initialize archive state
  const props = PropertiesService.getScriptProperties();
  props.setProperty('RAW_ARCHIVE_STATE', JSON.stringify({
    status: 'running',
    startIndex: 0, // Gmail search pagination index
    emailsProcessed: 0,
    filesExtracted: 0,
    filesSaved: 0,
    startTime: new Date().toISOString()
  }));
  
  ui.alert(
    'Archive Started',
    'Raw data archive started.\n\n' +
    'Run "Create Auto-Resume Trigger" to enable automatic processing every 10 minutes.\n\n' +
    'Or manually run "Resume Raw Data Archive" to continue.',
    ui.ButtonSet.OK
  );
  
  // Start first batch
  processNextRawDataBatch_();
}

// ---------------------
// View Raw Data Archive Progress
// ---------------------
function viewRawDataProgress() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('RAW_ARCHIVE_STATE');
  
  const ui = SpreadsheetApp.getUi();
  
  if (!stateJson) {
    ui.alert('Raw Data Archive Progress', 'No archive in progress.', ui.ButtonSet.OK);
    return;
  }
  
  const state = JSON.parse(stateJson);
  
  ui.alert(
    'Raw Data Archive Progress',
    `Status: ${state.status}\n` +
    `Current search index: ${state.startIndex}\n` +
    `Emails processed: ${state.emailsProcessed}\n` +
    `Files saved: ${state.filesSaved}\n` +
    `Started: ${new Date(state.startTime).toLocaleString()}`,
    ui.ButtonSet.OK
  );
}

// ---------------------
// Generate Detailed Progress Report (Email)
// ---------------------
function emailDetailedProgressReport() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('RAW_ARCHIVE_STATE');
  
  if (!stateJson) {
    SpreadsheetApp.getUi().alert('No archive in progress or completed.');
    return;
  }
  
  const state = JSON.parse(stateJson);
  const startTime = new Date(state.startTime);
  const now = new Date();
  const elapsed = now - startTime;
  const hours = Math.floor(elapsed / (1000 * 60 * 60));
  const minutes = Math.floor((elapsed % (1000 * 60 * 60)) / (1000 * 60));
  
  // Check Drive folder stats
  let driveStats = { monthFolders: [], totalFiles: 0, sampleCounts: [] };
  try {
    driveStats = analyzeDriveProgress_();
  } catch (e) {
    Logger.log('Error analyzing Drive: ' + e);
  }
  
  // Get recent execution history
  const executionSummary = getRecentExecutionSummary_();
  
  // Calculate projections
  const avgFilesPerEmail = state.emailsProcessed > 0 ? (state.filesSaved / state.emailsProcessed).toFixed(1) : 0;
  const processingRate = hours > 0 ? Math.round(state.emailsProcessed / hours) : 0;
  const estimatedTotal = state.emailsProcessed > 0 ? Math.round((driveStats.totalFiles / state.emailsProcessed) * state.emailsProcessed) : 8880;
  const percentComplete = estimatedTotal > 0 ? ((state.filesSaved / estimatedTotal) * 100).toFixed(1) : 0;
  
  const htmlReport = `
    <h2 style="color: #0066cc;">📊 CM360 Raw Data Archive - Progress Report</h2>
    
    <h3>📈 Current Status</h3>
    <table style="border-collapse: collapse; width: 100%;">
      <tr style="background-color: #f0f0f0;">
        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Status</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${state.status}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Emails Processed</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${state.emailsProcessed}</td>
      </tr>
      <tr style="background-color: #f0f0f0;">
        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Files Saved</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${state.filesSaved}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Avg Files per Email</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${avgFilesPerEmail}</td>
      </tr>
      <tr style="background-color: #f0f0f0;">
        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Current Search Index</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${state.startIndex}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Estimated Progress</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${percentComplete}%</td>
      </tr>
    </table>
    
    <h3>��️ Timing</h3>
    <table style="border-collapse: collapse; width: 100%;">
      <tr style="background-color: #f0f0f0;">
        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Started</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${startTime.toLocaleString()}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Elapsed Time</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${hours}h ${minutes}m</td>
      </tr>
      <tr style="background-color: #f0f0f0;">
        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Processing Rate</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${processingRate} emails/hour</td>
      </tr>
    </table>
    
    <h3>📁 Google Drive Analysis</h3>
    <table style="border-collapse: collapse; width: 100%;">
      <tr style="background-color: #f0f0f0;">
        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Month Folders Created</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${driveStats.monthFolders.join(', ') || 'None yet'}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Files in Drive</strong></td>
        <td style="padding: 8px; border: 1px solid #ddd;">${driveStats.totalFiles}</td>
      </tr>
    </table>
    
    ${driveStats.sampleCounts.length > 0 ? `
    <h3>🔍 Sample File Counts</h3>
    <table style="border-collapse: collapse; width: 100%;">
      <tr style="background-color: #f0f0f0;">
        <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Date Folder</th>
        <th style="padding: 8px; border: 1px solid #ddd; text-align: right;">Files</th>
      </tr>
      ${driveStats.sampleCounts.map((item, i) => `
        <tr${i % 2 === 0 ? ' style="background-color: #f9f9f9;"' : ''}>
          <td style="padding: 8px; border: 1px solid #ddd;">${item.folder}</td>
          <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${item.count}</td>
        </tr>
      `).join('')}
    </table>
    ` : ''}
    
    <h3>🔄 Recent Execution History</h3>
    <p>${executionSummary}</p>
    
    <hr style="border: 1px solid #ddd; margin: 20px 0;">
    <p style="color: #666; font-size: 12px;">Report generated: ${now.toLocaleString()}</p>
  `;
  
  MailApp.sendEmail({
    to: 'platformsolutionsadopshorizon@gmail.com',
    subject: `📊 CM360 Archive Progress - ${percentComplete}% Complete`,
    htmlBody: htmlReport
  });
  
  SpreadsheetApp.getUi().alert(
    'Progress Report Sent',
    `Detailed progress report sent to your email.\n\n` +
    `Status: ${state.status}\n` +
    `Files saved: ${state.filesSaved}\n` +
    `Estimated progress: ${percentComplete}%`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ---------------------
// INTERNAL: Analyze Drive Folder Progress
// ---------------------
function analyzeDriveProgress_() {
  const rootFolder = DriveApp.getFolderById(RAW_DATA_FOLDER_ID);
  const rawDataFolders = rootFolder.getFoldersByName('Raw Data');
  
  if (!rawDataFolders.hasNext()) {
    return { monthFolders: [], totalFiles: 0, sampleCounts: [] };
  }
  
  const rawDataFolder = rawDataFolders.next();
  const yearFolders = rawDataFolder.getFoldersByName('2025');
  
  if (!yearFolders.hasNext()) {
    return { monthFolders: [], totalFiles: 0, sampleCounts: [] };
  }
  
  const yearFolder = yearFolders.next();
  const monthFolders = [];
  const sampleCounts = [];
  let totalFiles = 0;
  
  const monthIterator = yearFolder.getFolders();
  while (monthIterator.hasNext()) {
    const monthFolder = monthIterator.next();
    const monthName = monthFolder.getName();
    monthFolders.push(monthName);
    
    // Count files in first 3 date folders of each month as sample
    const dateFolders = monthFolder.getFolders();
    let sampleCount = 0;
    let dateFoldersChecked = 0;
    
    while (dateFolders.hasNext() && dateFoldersChecked < 3) {
      const dateFolder = dateFolders.next();
      const files = dateFolder.getFiles();
      let count = 0;
      while (files.hasNext()) {
        files.next();
        count++;
        totalFiles++;
      }
      
      if (dateFoldersChecked === 0) {
        sampleCounts.push({ folder: `${monthName}/${dateFolder.getName()}`, count: count });
      }
      
      dateFoldersChecked++;
    }
  }
  
  return {
    monthFolders: monthFolders,
    totalFiles: totalFiles,
    sampleCounts: sampleCounts
  };
}

// ---------------------
// INTERNAL: Get Recent Execution Summary
// ---------------------
function getRecentExecutionSummary_() {
  // Note: This is a simple text summary since we can't programmatically access execution logs
  return 'Check Apps Script executions at: https://script.google.com/home/executions for detailed run history.';
}

// ---------------------
// Resume Raw Data Archive (if interrupted)
// ---------------------
function resumeRawDataArchive() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('RAW_ARCHIVE_STATE');
  
  const ui = SpreadsheetApp.getUi();
  
  if (!stateJson) {
    ui.alert('Resume Archive', 'No archive in progress to resume.', ui.ButtonSet.OK);
    return;
  }
  
  const state = JSON.parse(stateJson);
  
  if (state.status === 'completed') {
    ui.alert('Resume Archive', 'Archive already completed.', ui.ButtonSet.OK);
    return;
  }
  
  processNextRawDataBatch_();
}

// ---------------------
// Create Auto-Resume Trigger (Every 10 Minutes)
// ---------------------
function createRawDataAutoResumeTrigger() {
  // Delete any existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'autoResumeRawDataArchive') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Create new trigger: every 10 minutes
  ScriptApp.newTrigger('autoResumeRawDataArchive')
    .timeBased()
    .everyMinutes(10)
    .create();
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Auto-Resume Trigger Created',
    'Raw data archive will auto-resume every 10 minutes until complete.\n\n' +
    'Use "Delete Auto-Resume Trigger" to stop.',
    ui.ButtonSet.OK
  );
}

// ---------------------
// Delete Auto-Resume Trigger
// ---------------------
function deleteRawDataAutoResumeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'autoResumeRawDataArchive') {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    }
  }
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Auto-Resume Trigger Deleted',
    `Removed ${deleted} trigger(s). Archive will no longer auto-resume.`,
    ui.ButtonSet.OK
  );
}

// ---------------------
// Create Daily Evening Progress Report Trigger (7:30 PM)
// ---------------------
function createDailyProgressReportTrigger() {
  // Delete any existing triggers for this function first
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'sendDailyProgressReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Create new trigger: Daily at 7:30 PM
  ScriptApp.newTrigger('sendDailyProgressReport')
    .timeBased()
    .atHour(19) // 7 PM in 24-hour format
    .everyDays(1)
    .create();
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Daily Progress Report Trigger Created',
    'You will receive a detailed progress email every evening at 7:30 PM.\n\n' +
    'This will continue daily until you delete the trigger.\n\n' +
    'Use "Delete Daily Progress Report Trigger" to stop.',
    ui.ButtonSet.OK
  );
}

// ---------------------
// Delete Daily Evening Progress Report Trigger
// ---------------------
function deleteDailyProgressReportTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'sendDailyProgressReport') {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    }
  }
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Daily Progress Report Trigger Deleted',
    `Removed ${deleted} trigger(s). No more daily progress emails.`,
    ui.ButtonSet.OK
  );
}

// ---------------------
// Send Daily Progress Report (Called by Trigger)
// ---------------------
function sendDailyProgressReport() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('RAW_ARCHIVE_STATE');
  
  // If no archive in progress or completed, send notification
  if (!stateJson) {
    MailApp.sendEmail({
      to: 'platformsolutionsadopshorizon@gmail.com',
      subject: '📊 CM360 Archive - No Active Archive',
      htmlBody: `
        <h3>Daily Progress Report</h3>
        <p>No archive is currently in progress or no state found.</p>
        <p><strong>Time:</strong> ${new Date().toLocaleString()}</p>
        <p>If you expected an archive to be running, check the Scripts and Triggers.</p>
      `
    });
    return;
  }
  
  const state = JSON.parse(stateJson);
  
  // If completed, send final summary and delete this trigger
  if (state.status === 'completed') {
    MailApp.sendEmail({
      to: 'platformsolutionsadopshorizon@gmail.com',
      subject: '� CM360 Archive COMPLETE - Daily Report Trigger Stopping',
      htmlBody: `
        <h2 style="color: #00cc00;">� Archive Complete!</h2>
        <p>The raw data archive has finished successfully.</p>
        <p><strong>Total emails processed:</strong> ${state.emailsProcessed}</p>
        <p><strong>Total files saved:</strong> ${state.filesSaved}</p>
        <p><strong>Completed:</strong> ${state.endTime ? new Date(state.endTime).toLocaleString() : 'Recently'}</p>
        <hr>
        <p>Daily progress report trigger will now stop automatically.</p>
        <p><strong>Next step:</strong> Run "Categorize Files by Network" to organize by network folders.</p>
      `
    });
    
    // Auto-delete this daily trigger since archive is done
    deleteDailyProgressReportTrigger();
    return;
  }
  
  // Otherwise, send the detailed progress report
  emailDetailedProgressReport();
}

// ---------------------
// Auto-Resume Function (Called by Trigger)
// ---------------------
function autoResumeRawDataArchive() {
  const props = PropertiesService.getScriptProperties();
  const docProps = PropertiesService.getDocumentProperties();
  const archiveStateJson = props.getProperty('RAW_ARCHIVE_STATE');
  const auditStateJson = docProps.getProperty('comprehensive_audit_state');
  
  // Check for comprehensive audit first
  if (auditStateJson) {
    const auditState = JSON.parse(auditStateJson);
    Logger.log('Auto-resuming comprehensive audit...');
    processComprehensiveAuditBatch_();
    return;
  }
  
  // Then check for archive
  if (!archiveStateJson) {
    return; // No archive or audit in progress
  }
  
  const state = JSON.parse(archiveStateJson);
  
  if (state.status === 'completed') {
    // Archive complete, delete trigger
    deleteRawDataAutoResumeTrigger();
    return;
  }
  
  // Route to appropriate processor based on mode
  if (state.mode === 'gap-fill') {
    Logger.log('Auto-resuming gap-fill archive...');
    processGapFillBatch_();
  } else {
    Logger.log('Auto-resuming full archive...');
    processNextRawDataBatch_();
  }
}

// ---------------------
// INTERNAL: Process Next Batch (Search All Emails)
// ---------------------
function processNextRawDataBatch_() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('RAW_ARCHIVE_STATE');
  
  if (!stateJson) {
    Logger.log('No raw data archive state found');
    return;
  }
  
  const state = JSON.parse(stateJson);
  
  try {
    // Search for ALL emails with the subject (no date filters)
    const query = `subject:"${RAW_DATA_SEARCH_SUBJECT}"`;
    
    Logger.log(`Searching Gmail from index ${state.startIndex} with batch size ${RAW_BATCH_SIZE}`);
    
    const threads = GmailApp.search(query, state.startIndex, RAW_BATCH_SIZE);
    
    Logger.log(`Found ${threads.length} threads`);
    
    // If no threads found, check for new emails before declaring complete
    if (threads.length === 0) {
      Logger.log('No more threads at current index. Checking for new emails...');
      
      // Check if any new emails arrived at the top of inbox (index 0)
      const newEmailCheck = checkForNewEmails_(state);
      
      if (newEmailCheck.hasNewEmails) {
        Logger.log(`Found ${newEmailCheck.newEmailCount} new emails. Restarting from index 0...`);
        
        // Reset to start but preserve stats
        state.startIndex = 0;
        state.lastCheckTime = new Date().toISOString();
        props.setProperty('RAW_ARCHIVE_STATE', JSON.stringify(state));
        
        // Continue processing
        return;
      }
      
      // No new emails, we're truly done
      state.status = 'completed';
      state.endTime = new Date().toISOString();
      props.setProperty('RAW_ARCHIVE_STATE', JSON.stringify(state));
      
      // Send completion email
      sendRawDataCompletionEmail_(state);
      
      // Auto-delete trigger
      deleteRawDataAutoResumeTrigger();
      
      Logger.log('� Raw data archive complete!');
      return;
    }
    
    // Process this batch of threads
    let batchEmailsProcessed = 0;
    let batchFilesSaved = 0;
    let lastProcessedEmailIndex = state.startIndex; // Track actual progress
    
    for (let i = 0; i < threads.length; i++) {
      try {
        const thread = threads[i];
        const messages = thread.getMessages();
        
        for (const message of messages) {
          const emailDate = message.getDate();
          const year = emailDate.getFullYear();
          const month = emailDate.getMonth() + 1; // JavaScript months are 0-indexed
          const dateStr = Utilities.formatDate(emailDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          
          // Get or create folder structure: Raw Data/2025/05-May/2025-05-15/
          const monthFolder = getOrCreateRawDataMonthFolder_(year, month);
          const dateFolder = getOrCreateDateFolder_(monthFolder, dateStr);
          
          const attachments = message.getAttachments();
          
          for (const attachment of attachments) {
            const filename = attachment.getName();
            const lowerFilename = filename.toLowerCase();
            
            // Handle ZIP files
            if (lowerFilename.endsWith('.zip')) {
              try {
                const zipBlob = attachment.copyBlob();
                const unzipped = Utilities.unzip(zipBlob);
                
                for (const file of unzipped) {
                  const unzippedName = file.getName().toLowerCase();
                  if (unzippedName.endsWith('.csv') || unzippedName.endsWith('.xlsx')) {
                    if (saveRawFileToDrive_(file, dateFolder, file.getName())) {
                      batchFilesSaved++;
                    }
                  }
                }
              } catch (error) {
                Logger.log(`Error unzipping ${filename}: ${error}`);
              }
            }
            // Handle CSV and XLSX files directly
            else if (lowerFilename.endsWith('.csv') || lowerFilename.endsWith('.xlsx')) {
              if (saveRawFileToDrive_(attachment, dateFolder, filename)) {
                batchFilesSaved++;
              }
            }
          }
          
          batchEmailsProcessed++;
        }
        
        // Update last processed index after each thread completes
        lastProcessedEmailIndex = state.startIndex + i + 1;
        
        // Save state every 10 threads (balance between safety and performance)
        // This protects against data loss while minimizing Script Properties writes
        if ((i + 1) % 10 === 0 || i === threads.length - 1) {
          state.startIndex = lastProcessedEmailIndex;
          state.emailsProcessed += batchEmailsProcessed;
          state.filesSaved = (state.filesSaved || 0) + batchFilesSaved;
          props.setProperty('RAW_ARCHIVE_STATE', JSON.stringify(state));
          
          Logger.log(`State saved at index ${state.startIndex} (${batchEmailsProcessed} emails, ${batchFilesSaved} files in this checkpoint)`);
          
          // Reset batch counters since we just saved
          batchEmailsProcessed = 0;
          batchFilesSaved = 0;
        }
        
      } catch (threadError) {
        Logger.log(`Error processing thread at index ${state.startIndex + i}: ${threadError}`);
        // Skip this thread and continue with next one
        // State is already saved from previous successful thread
        continue;
      }
    }
    
    Logger.log(`Batch complete. Next index: ${state.startIndex}`);
    
    // Send progress email every 500 emails
    if (state.emailsProcessed % 500 === 0) {
      MailApp.sendEmail({
        to: 'platformsolutionsadopshorizon@gmail.com',
        subject: `📊 CM360 Raw Data Archive Progress - ${state.emailsProcessed} emails`,
        htmlBody: `<h3>Raw Data Archive Progress</h3>
        <p><strong>Emails processed:</strong> ${state.emailsProcessed}</p>
        <p><strong>Files saved:</strong> ${state.filesSaved}</p>
        <p><strong>Current search index:</strong> ${state.startIndex}</p>
        <p><strong>Started:</strong> ${new Date(state.startTime).toLocaleString()}</p>`
      });
    }
    
  } catch (error) {
    Logger.log('Error processing batch: ' + error);
    
    // Save current state before erroring out (preserve progress)
    // Note: state updates happen in the main loop, so we need to check if we have partial progress
    if (state) {
      props.setProperty('RAW_ARCHIVE_STATE', JSON.stringify(state));
      Logger.log(`State saved on error. Current index: ${state.startIndex}, Files saved: ${state.filesSaved}`);
    }
    
    MailApp.sendEmail({
      to: 'platformsolutionsadopshorizon@gmail.com',
      subject: '��️ CM360 Raw Data Archive Error',
      htmlBody: `<h3>Raw Data Archive Error</h3>
        <p><strong>Error:</strong> ${error}</p>
        <p><strong>Progress:</strong> ${state ? state.emailsProcessed : 'unknown'} emails, ${state ? state.filesSaved : 'unknown'} files saved</p>
        <p><strong>Search index:</strong> ${state ? state.startIndex : 'unknown'}</p>
        <p><strong>Stack:</strong> ${error.stack || 'No stack trace'}</p>
        <p>Use "Resume Raw Data Archive" to continue.</p>`
    });
  }
}

// ---------------------
// INTERNAL: Send Completion Email
// ---------------------
function sendRawDataCompletionEmail_(state) {
  const startTime = new Date(state.startTime);
  const endTime = new Date(state.endTime);
  const durationMs = endTime - startTime;
  const hours = Math.floor(durationMs / (1000 * 60 * 60));
  const minutes = Math.floor((durationMs % (1000 * 60 * 60)) / (1000 * 60));
  
  const avgFilesPerEmail = (state.emailsProcessed > 0 ? (state.filesSaved / state.emailsProcessed).toFixed(1) : 0);
  
  MailApp.sendEmail({
    to: 'platformsolutionsadopshorizon@gmail.com',
    subject: '� CM360 Raw Data Archive Complete - Full Inbox Archived',
    htmlBody: `
      <h2 style="color: #0066cc;">� CM360 Raw Data Archive Complete</h2>
      
      <h3>📊 Overall Statistics</h3>
      <table style="border-collapse: collapse; width: 100%;">
        <tr style="background-color: #f0f0f0;">
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Emails Processed</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${state.emailsProcessed}</td>
        </tr>
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Files Saved</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${state.filesSaved}</td>
        </tr>
        <tr style="background-color: #f0f0f0;">
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Average Files per Email</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${avgFilesPerEmail}</td>
        </tr>
      </table>
      
      <h3>��️ Performance</h3>
      <table style="border-collapse: collapse; width: 100%;">
        <tr style="background-color: #f0f0f0;">
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Start Time</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${startTime.toLocaleString()}</td>
        </tr>
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>End Time</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${endTime.toLocaleString()}</td>
        </tr>
        <tr style="background-color: #f0f0f0;">
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Duration</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${hours}h ${minutes}m</td>
        </tr>
      </table>
      
      <h3>📁 File Location</h3>
      <p><a href="https://drive.google.com/drive/folders/${RAW_DATA_FOLDER_ID}" style="color: #0066cc; font-weight: bold;">View Raw Data Archive in Google Drive</a></p>
      <p><strong>Folder Structure:</strong> Raw Data/[Year]/[Month]/[Date]/files</p>
      
      <h3>📋 Next Steps</h3>
      <ol>
        <li><strong>Review the data:</strong> Check Drive folder to verify all files saved correctly</li>
        <li><strong>Audit completeness:</strong> Run "Audit Archive Completeness" to check for gaps</li>
        <li><strong>Categorize by network:</strong> Run "Categorize Raw Data by Network" from the menu</li>
        <li><strong>Build ROI dashboard:</strong> Use categorized data to analyze violations and cost savings</li>
      </ol>
      
      <hr style="border: 1px solid #ddd; margin: 20px 0;">
      <p style="color: #34a853; font-size: 12px;">� Archive checked for new emails that arrived during processing - all caught up!</p>
      <p style="color: #666; font-size: 12px;">Auto-resume trigger has been automatically deleted. Archive state saved in Script Properties.</p>
    `
  });
}

// ---------------------
// INTERNAL: Check for New Emails
// ---------------------
function checkForNewEmails_(state) {
  try {
    // Search for emails from index 0 (top of inbox)
    const query = `subject:"${RAW_DATA_SEARCH_SUBJECT}"`;
    const recentThreads = GmailApp.search(query, 0, 10); // Check first 10 emails
    
    if (recentThreads.length === 0) {
      return { hasNewEmails: false, newEmailCount: 0 };
    }
    
    // Use last check time if available, otherwise use start time (first check only)
    const lastCheck = state.lastCheckTime ? new Date(state.lastCheckTime) : new Date(state.startTime);
    
    let newEmailCount = 0;
    for (const thread of recentThreads) {
      const messages = thread.getMessages();
      for (const message of messages) {
        const messageDate = message.getDate();
        
        // If email is newer than last time we checked, it's new
        if (messageDate > lastCheck) {
          newEmailCount++;
        }
      }
    }
    
    Logger.log(`New email check: Found ${newEmailCount} emails newer than last check (${lastCheck.toISOString()})`);
    
    return {
      hasNewEmails: newEmailCount > 0,
      newEmailCount: newEmailCount
    };
    
  } catch (error) {
    Logger.log('Error checking for new emails: ' + error);
    return { hasNewEmails: false, newEmailCount: 0 };
  }
}

// ---------------------
// INTERNAL: Get or Create Month Folder
// ---------------------
function getOrCreateRawDataMonthFolder_(year, month) {
  const monthStr = String(month).padStart(2, '0');
  const monthName = getMonthName_(month);
  
  const rootFolder = DriveApp.getFolderById(RAW_DATA_FOLDER_ID);
  
  // Get or create "Raw Data" folder
  let rawDataFolder;
  const rawDataFolders = rootFolder.getFoldersByName('Raw Data');
  if (rawDataFolders.hasNext()) {
    rawDataFolder = rawDataFolders.next();
  } else {
    rawDataFolder = rootFolder.createFolder('Raw Data');
  }
  
  // Get or create year folder
  let yearFolder;
  const yearFolders = rawDataFolder.getFoldersByName(String(year));
  if (yearFolders.hasNext()) {
    yearFolder = yearFolders.next();
  } else {
    yearFolder = rawDataFolder.createFolder(String(year));
  }
  
  // Get or create month folder: "04-April"
  let monthFolder;
  const monthFolderName = `${monthStr}-${monthName}`;
  const monthFolders = yearFolder.getFoldersByName(monthFolderName);
  if (monthFolders.hasNext()) {
    monthFolder = monthFolders.next();
  } else {
    monthFolder = yearFolder.createFolder(monthFolderName);
  }
  
  return monthFolder;
}

// ---------------------
// INTERNAL: Get or Create Date Folder
// ---------------------
function getOrCreateDateFolder_(monthFolder, dateStr) {
  // dateStr format: "2025-04-15"
  let dateFolder;
  const dateFolders = monthFolder.getFoldersByName(dateStr);
  if (dateFolders.hasNext()) {
    dateFolder = dateFolders.next();
  } else {
    dateFolder = monthFolder.createFolder(dateStr);
  }
  
  return dateFolder;
}

// ---------------------
// INTERNAL: Save Raw File to Drive
// ---------------------
function saveRawFileToDrive_(attachment, folder, filename) {
  // Check if file already exists
  const existingFiles = folder.getFilesByName(filename);
  if (existingFiles.hasNext()) {
    Logger.log(`File already exists: ${filename}`);
    return false; // File not saved (already exists)
  }
  
  // Create file
  folder.createFile(attachment.copyBlob().setName(filename));
  Logger.log(`Saved: ${filename}`);
  return true; // File saved successfully
}

// ---------------------
// CATEGORIZE: Organize Files by Network (Run After Archive Complete)
// ---------------------
function categorizeRawDataByNetwork() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Categorize Raw Data by Network',
    'This will organize all saved raw data files into network folders.\n\n' +
    'Files will be analyzed and moved to:\n' +
    'Raw Data/Networks/[NetworkID - NetworkName]/[Date]/\n\n' +
    'This may take 10-20 minutes depending on file count.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  ui.alert('Categorization Started', 'Processing all files... You will receive an email when complete with detailed statistics.', ui.ButtonSet.OK);
  
  const startTime = new Date();
  
  try {
    const networkMap = loadNetworkMapping_();
    
    if (Object.keys(networkMap).length === 0) {
      ui.alert('Error', 'No networks found in Networks tab. Please check columns A and B.', ui.ButtonSet.OK);
      return;
    }
    
    const stats = categorizeAllFiles_(networkMap);
    const endTime = new Date();
    const durationMin = Math.round((endTime - startTime) / (1000 * 60));
    
    // Send detailed completion email
    MailApp.sendEmail({
      to: 'platformsolutionsadopshorizon@gmail.com',
      subject: '� CM360 Raw Data Categorization Complete - Summary Report',
      htmlBody: `
        <h2 style="color: #0066cc;">� File Categorization Complete</h2>
        
        <h3>📊 Overall Statistics</h3>
        <table style="border-collapse: collapse; width: 100%;">
          <tr style="background-color: #f0f0f0;">
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Files Processed</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd;">${stats.totalFiles}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Files Categorized</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd;">${stats.filesCategorized} (${((stats.filesCategorized / stats.totalFiles) * 100).toFixed(1)}%)</td>
          </tr>
          <tr style="background-color: #f0f0f0;">
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Files Uncategorized</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd;">${stats.filesUncategorized} (${((stats.filesUncategorized / stats.totalFiles) * 100).toFixed(1)}%)</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Networks Found</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd;">${stats.networksFound} of ${Object.keys(networkMap).length} total</td>
          </tr>
          <tr style="background-color: #f0f0f0;">
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Date Folders Processed</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd;">${stats.dateFoldersProcessed}</td>
          </tr>
        </table>
        
        <h3>��️ Performance</h3>
        <table style="border-collapse: collapse; width: 100%;">
          <tr style="background-color: #f0f0f0;">
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Start Time</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd;">${startTime.toLocaleString()}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>End Time</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd;">${endTime.toLocaleString()}</td>
          </tr>
          <tr style="background-color: #f0f0f0;">
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Duration</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd;">${durationMin} minutes</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Processing Rate</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd;">${Math.round(stats.totalFiles / durationMin)} files/minute</td>
          </tr>
        </table>
        
        <h3>🌐 Top Networks by File Count</h3>
        <table style="border-collapse: collapse; width: 100%;">
          <tr style="background-color: #f0f0f0;">
            <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Network</th>
            <th style="padding: 8px; border: 1px solid #ddd; text-align: right;">Files</th>
          </tr>
          ${stats.networkBreakdown.slice(0, 10).map((net, i) => `
            <tr${i % 2 === 0 ? ' style="background-color: #f9f9f9;"' : ''}>
              <td style="padding: 8px; border: 1px solid #ddd;">${net.name}</td>
              <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${net.count}</td>
            </tr>
          `).join('')}
        </table>
        ${stats.networkBreakdown.length > 10 ? `<p style="color: #666; font-size: 12px;">...and ${stats.networkBreakdown.length - 10} more networks</p>` : ''}
        
        <h3>📁 File Locations</h3>
        <p><strong>Categorized Files:</strong> <a href="https://drive.google.com/drive/folders/${RAW_DATA_FOLDER_ID}" style="color: #0066cc;">Raw Data/Networks/</a></p>
        <p><strong>Uncategorized Files:</strong> Remain in Raw Data/2025/[Month]/[Date]/ folders</p>
        
        <h3>📋 Next Steps</h3>
        <ol>
          <li><strong>Review uncategorized files:</strong> ${stats.filesUncategorized > 0 ? 'Check files without network IDs in filename' : 'None to review! �'}</li>
          <li><strong>Verify network folders:</strong> Spot-check a few networks to confirm proper organization</li>
          <li><strong>Build ROI analysis:</strong> Ready to analyze violations and cost savings per network</li>
        </ol>
        
        <hr style="border: 1px solid #ddd; margin: 20px 0;">
        <p style="color: #666; font-size: 12px;">Categorization process completed successfully. Original date-organized folders preserved.</p>
      `
    });
    
    ui.alert(
      'Categorization Complete',
      `� ${stats.filesCategorized} files organized into ${stats.networksFound} network folders\n` +
      `��️ ${stats.filesUncategorized} files remain uncategorized\n\n` +
      `Duration: ${durationMin} minutes\n\n` +
      'Check your email for detailed statistics.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('Categorization error: ' + error);
    
    MailApp.sendEmail({
      to: 'platformsolutionsadopshorizon@gmail.com',
      subject: '��️ CM360 Raw Data Categorization Error',
      htmlBody: `
        <h3 style="color: #cc0000;">Categorization Error</h3>
        <p><strong>Error:</strong> ${error}</p>
        <p><strong>Time:</strong> ${new Date().toLocaleString()}</p>
        <p><strong>Action:</strong> Try running "Categorize Raw Data by Network" again or check the Networks tab data.</p>
      `
    });
    
    ui.alert('Error', 'Categorization failed: ' + error + '\n\nCheck your email for details.', ui.ButtonSet.OK);
  }
}

// ---------------------
// INTERNAL: Categorize All Files
// ---------------------
function categorizeAllFiles_(networkMap) {
  const rootFolder = DriveApp.getFolderById(RAW_DATA_FOLDER_ID);
  const rawDataFolder = rootFolder.getFoldersByName('Raw Data').next();
  
  // Create Networks folder
  let networksFolder;
  const networksFolders = rawDataFolder.getFoldersByName('Networks');
  if (networksFolders.hasNext()) {
    networksFolder = networksFolders.next();
  } else {
    networksFolder = rawDataFolder.createFolder('Networks');
  }
  
  let filesCategorized = 0;
  let filesUncategorized = 0;
  let dateFoldersProcessed = 0;
  const networksFound = new Set();
  const networkFileCounts = {}; // Track files per network
  
  // Iterate through year folders (2025)
  const yearFolders = rawDataFolder.getFolders();
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    if (yearFolder.getName() === 'Networks') continue; // Skip Networks folder
    
    Logger.log(`Processing year folder: ${yearFolder.getName()}`);
    
    // Iterate through month folders (04-April, 05-May, etc.)
    const monthFolders = yearFolder.getFolders();
    while (monthFolders.hasNext()) {
      const monthFolder = monthFolders.next();
      Logger.log(`  Processing month folder: ${monthFolder.getName()}`);
      
      // Iterate through date folders (2025-04-15, etc.)
      const dateFolders = monthFolder.getFolders();
      while (dateFolders.hasNext()) {
        const dateFolder = dateFolders.next();
        const dateStr = dateFolder.getName(); // e.g., "2025-04-15"
        dateFoldersProcessed++;
        
        Logger.log(`    Processing date folder: ${dateStr}`);
        
        // Iterate through files in this date folder
        const files = dateFolder.getFiles();
        while (files.hasNext()) {
          const file = files.next();
          const filename = file.getName();
          
          // Extract network ID from filename
          const networkId = extractNetworkIdFromFilename_(filename, networkMap);
          
          if (networkId && networkMap[networkId]) {
            // Create network folder structure
            const networkName = networkMap[networkId];
            const networkFolder = getOrCreateNetworkFolder_(networksFolder, networkId, networkName);
            const networkDateFolder = getOrCreateDateFolder_(networkFolder, dateStr);
            
            // Rename file to include friendly network name
            const newFilename = renameFileWithNetworkName_(filename, networkId, networkName);
            
            // Move and rename file
            const movedFile = file.moveTo(networkDateFolder);
            if (newFilename !== filename) {
              movedFile.setName(newFilename);
              Logger.log(`Renamed and moved: ${filename} �� ${newFilename}`);
            }
            
            filesCategorized++;
            networksFound.add(networkId);
            
            // Track count per network
            if (!networkFileCounts[networkId]) {
              networkFileCounts[networkId] = { name: networkName, count: 0 };
            }
            networkFileCounts[networkId].count++;
            
            Logger.log(`Categorized: ${newFilename} �� ${networkId} - ${networkName}/${dateStr}`);
          } else {
            filesUncategorized++;
            Logger.log(`Uncategorized: ${filename}`);
          }
        }
      }
    }
  }
  
  // Sort networks by file count (descending)
  const networkBreakdown = Object.keys(networkFileCounts)
    .map(id => ({
      id: id,
      name: `${id} - ${networkFileCounts[id].name}`,
      count: networkFileCounts[id].count
    }))
    .sort((a, b) => b.count - a.count);
  
  return {
    filesCategorized: filesCategorized,
    filesUncategorized: filesUncategorized,
    totalFiles: filesCategorized + filesUncategorized,
    networksFound: networksFound.size,
    dateFoldersProcessed: dateFoldersProcessed,
    networkBreakdown: networkBreakdown
  };
}

// ---------------------
// INTERNAL: Get or Create Network Folder
// ---------------------
function getOrCreateNetworkFolder_(networksFolder, networkId, networkName) {
  const networkFolderName = `${networkId} - ${networkName}`;
  
  let networkFolder;
  const folders = networksFolder.getFoldersByName(networkFolderName);
  if (folders.hasNext()) {
    networkFolder = folders.next();
  } else {
    networkFolder = networksFolder.createFolder(networkFolderName);
  }
  
  return networkFolder;
}

// ---------------------
// INTERNAL: Load Network Mapping from Networks Tab
// ---------------------
function loadNetworkMapping_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const networksSheet = ss.getSheetByName('Networks');
  
  if (!networksSheet) {
    Logger.log('Networks tab not found');
    return {};
  }
  
  const data = networksSheet.getDataRange().getValues();
  const networkMap = {};
  
  // Skip header row, read columns A (Network ID) and B (Network Name)
  for (let i = 1; i < data.length; i++) {
    const networkId = String(data[i][0]).trim();
    const networkName = String(data[i][1]).trim();
    
    if (networkId && networkName) {
      networkMap[networkId] = networkName;
    }
  }
  
  Logger.log(`Loaded ${Object.keys(networkMap).length} networks from Networks tab`);
  return networkMap;
}

// ---------------------
// INTERNAL: Extract Network ID from Filename
// ---------------------
function extractNetworkIdFromFilename_(filename, networkMap) {
  // Filename patterns might include network ID
  // Examples: "898158_report.csv", "DCM_898158.zip", etc.
  // Try to extract any number sequence that matches known network IDs
  
  const matches = filename.match(/\d{3,7}/g); // Look for 3-7 digit numbers
  
  if (!matches) {
    return null;
  }
  
  // Check each number found to see if it's a valid network ID
  for (const match of matches) {
    if (networkMap[match]) {
      return match;
    }
  }
  
  return null;
}

// ---------------------
// INTERNAL: Rename File with Friendly Network Name
// ---------------------
function renameFileWithNetworkName_(filename, networkId, networkName) {
  // Clean network name for filename (remove special characters, limit length)
  const cleanNetworkName = networkName
    .replace(/[^a-zA-Z0-9\s-]/g, '') // Remove special chars except spaces and hyphens
    .replace(/\s+/g, '_') // Replace spaces with underscores
    .substring(0, 50); // Limit to 50 chars
  
  // Get file extension
  const lastDot = filename.lastIndexOf('.');
  const extension = lastDot > 0 ? filename.substring(lastDot) : '';
  const nameWithoutExt = lastDot > 0 ? filename.substring(0, lastDot) : filename;
  
  // Check if network name is already in the filename
  if (nameWithoutExt.includes(cleanNetworkName)) {
    return filename; // Already has friendly name
  }
  
  // Build new filename: NetworkID_NetworkName_OriginalFilename.ext
  // Example: 898158_Advertiser_Inc_BKCM360_Global_QA_Check_20250515.csv
  const newFilename = `${networkId}_${cleanNetworkName}_${nameWithoutExt}${extension}`;
  
  return newFilename;
}

/**
 * COMPREHENSIVE AUDIT: Validates ALL attachments from Gmail are in Drive
 * Scans Gmail for expected files and compares to actual Drive contents
 * Supports chunked execution with auto-resume
 */
function auditRawDataArchiveComprehensive() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const stateKey = 'comprehensive_audit_state';
  
  // Check if audit is already in progress
  const existingState = props.getProperty(stateKey);
  if (existingState) {
    const response = ui.alert(
      '��️ Audit In Progress',
      'An audit is already running.\n\n' +
      'Continue from where it left off?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Resume existing audit
    processComprehensiveAuditBatch_();
    return;
  }
  
  const response = ui.alert(
    '🔍 Comprehensive Archive Audit',
    'This will:\n\n' +
    '1. Scan ALL emails with subject "BKCM360 Global QA Check"\n' +
    '2. Count attachments per date/network in Gmail\n' +
    '3. Count actual files in Drive per date/network\n' +
    '4. Report any missing files\n\n' +
    'Estimated time: Multiple 6-min runs (auto-resumes)\n' +
    'Create auto-resume trigger recommended!\n\n' +
    'Proceed?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // Initialize state
  const state = {
    phase: 'gmail_scan',
    gmailStartIndex: 0,
    expectedCountsJson: '{}',
    actualCountsJson: '{}',
    startTime: new Date().toISOString()
  };
  
  props.setProperty(stateKey, JSON.stringify(state));
  
  ui.alert(
    '� Audit Started',
    'Phase 1: Scanning Gmail\n\n' +
    'Create an auto-resume trigger to continue automatically every 10 minutes.\n\n' +
    'Menu �� ARCHIVE TOOLS �� Create Auto-Resume Trigger',
    ui.ButtonSet.OK
  );
  
  // Start first batch
  processComprehensiveAuditBatch_();
}

/**
 * Process one batch of comprehensive audit (called by trigger or manually)
 */
function processComprehensiveAuditBatch_() {
  const props = PropertiesService.getDocumentProperties();
  const stateKey = 'comprehensive_audit_state';
  const stateJson = props.getProperty(stateKey);
  
  if (!stateJson) {
    Logger.log('No audit in progress');
    return;
  }
  
  const state = JSON.parse(stateJson);
  const startTime = new Date();
  const MAX_EXECUTION_MS = 5 * 60 * 1000; // 5 minutes
  
  try {
    if (state.phase === 'gmail_scan') {
      Logger.log(`Gmail scan: Starting at index ${state.gmailStartIndex}`);
      
      // Continue Gmail scan
      const expectedCounts = new Map(Object.entries(JSON.parse(state.expectedCountsJson)).map(([k, v]) => [k, Number(v)]));
      const result = scanGmailBatch_(state.gmailStartIndex, expectedCounts, startTime, MAX_EXECUTION_MS);
      
      state.expectedCountsJson = JSON.stringify(Object.fromEntries(result.expectedCounts));
      state.gmailStartIndex = result.nextIndex;
      
      if (result.complete) {
        const totalFiles = Array.from(result.expectedCounts.values()).reduce((sum, count) => sum + count, 0);
        Logger.log(`Gmail scan complete: ${totalFiles} files found`);
        state.phase = 'drive_scan';
      }
      
      props.setProperty(stateKey, JSON.stringify(state));
      Logger.log(`Progress saved: ${state.phase}, Gmail index: ${state.gmailStartIndex}`);
      
    } else if (state.phase === 'drive_scan') {
      Logger.log('Drive scan: Scanning folders...');
      
      // Scan Drive (usually completes in one run)
      const actualCounts = scanDriveForActualCounts_();
      state.actualCountsJson = JSON.stringify(Object.fromEntries(actualCounts));
      state.phase = 'compare';
      
      props.setProperty(stateKey, JSON.stringify(state));
      const totalFiles = Array.from(actualCounts.values()).reduce((sum, count) => sum + count, 0);
      Logger.log(`Drive scan complete: ${totalFiles} files found`);
      
      // Continue to comparison immediately if time allows
      if (new Date() - startTime < MAX_EXECUTION_MS) {
        processComprehensiveAuditBatch_();
      }
      
    } else if (state.phase === 'compare') {
      Logger.log('Comparison phase: Analyzing differences...');
      
      const expectedCounts = new Map(Object.entries(JSON.parse(state.expectedCountsJson)).map(([k, v]) => [k, Number(v)]));
      const actualCounts = new Map(Object.entries(JSON.parse(state.actualCountsJson)).map(([k, v]) => [k, Number(v)]));
      
      const missingDateNetworks = [];
      const extraDateNetworks = [];
      const countMismatches = [];
      
      // Check for missing date/networks (in Gmail but not Drive)
      expectedCounts.forEach((expectedCount, key) => {
        const actualCount = actualCounts.get(key) || 0;
        const parts = key.split('|');
        
        if (actualCount === 0) {
          missingDateNetworks.push({
            date: parts[0],
            networkId: parts[1],
            expectedCount: expectedCount
          });
        } else if (actualCount !== expectedCount) {
          countMismatches.push({
            date: parts[0],
            networkId: parts[1],
            expectedCount: expectedCount,
            actualCount: actualCount,
            difference: actualCount - expectedCount
          });
        }
      });
      
      // Check for extra date/networks (in Drive but not Gmail)
      actualCounts.forEach((actualCount, key) => {
        if (!expectedCounts.has(key)) {
          const parts = key.split('|');
          extraDateNetworks.push({
            date: parts[0],
            networkId: parts[1],
            actualCount: actualCount
          });
        }
      });
      
      const totalExpected = Array.from(expectedCounts.values()).reduce((sum, count) => sum + count, 0);
      const totalActual = Array.from(actualCounts.values()).reduce((sum, count) => sum + count, 0);
      
      // Send report
      sendComprehensiveAuditReportCounts_(totalExpected, totalActual, missingDateNetworks, extraDateNetworks, countMismatches);
      
      // Clean up state
      props.deleteProperty(stateKey);
      
      Logger.log('� Comprehensive audit complete and email sent');
    }
    
  } catch (error) {
    Logger.log('Error in audit batch: ' + error);
    
    // Send error email
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: '�� Comprehensive Audit Error',
      body: `Error in ${state.phase} phase: ${error.toString()}\n\nProgress saved. Run again to resume.`
    });
  }
}

/**
 * Scan Gmail in batches with time limit
 * Returns: { expectedCounts: Map, nextIndex: number, complete: boolean }
 * expectedCounts format: "date|networkId" => count
 */
function scanGmailBatch_(startIndex, expectedCounts, startTime, maxExecutionMs) {
  const query = 'subject:"CM360 CPC/CPM FLIGHT QA"';
  const batchSize = 100;
  let currentIndex = startIndex;
  
  while (true) {
    // Check time limit
    const elapsed = new Date() - startTime;
    if (elapsed > maxExecutionMs) {
      Logger.log(`Time limit reached at index ${currentIndex}`);
      return { expectedCounts, nextIndex: currentIndex, complete: false };
    }
    
    const threads = GmailApp.search(query, currentIndex, batchSize);
    if (threads.length === 0) {
      Logger.log(`Gmail scan complete at index ${currentIndex}`);
      return { expectedCounts, nextIndex: currentIndex, complete: true };
    }
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      
      for (const message of messages) {
        const messageDate = message.getDate();
        const dateStr = Utilities.formatDate(messageDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const attachments = message.getAttachments();
        
        for (const attachment of attachments) {
          const filename = attachment.getName();
          const lowerFilename = filename.toLowerCase();
          
          // Only count CSV/XLSX files (or files inside ZIPs)
          if (lowerFilename.endsWith('.zip')) {
            const networkId = extractNetworkIdFromFilename_(filename, getNetworkMap_());
            if (networkId) {
              try {
                const zipBlob = attachment.copyBlob();
                const unzipped = Utilities.unzip(zipBlob);
                
                for (const file of unzipped) {
                  const unzippedName = file.getName().toLowerCase();
                  if (unzippedName.endsWith('.csv') || unzippedName.endsWith('.xlsx')) {
                    const key = `${dateStr}|${networkId}`;
                    expectedCounts.set(key, (expectedCounts.get(key) || 0) + 1);
                  }
                }
              } catch (e) {
                Logger.log(`Error unzipping ${filename}: ${e}`);
              }
            }
          } else if (lowerFilename.endsWith('.csv') || lowerFilename.endsWith('.xlsx')) {
            const networkId = extractNetworkIdFromFilename_(filename, getNetworkMap_());
            if (networkId) {
              const key = `${dateStr}|${networkId}`;
              expectedCounts.set(key, (expectedCounts.get(key) || 0) + 1);
            }
          }
        }
      }
    }
    
    currentIndex += batchSize;
    
    // Progress log every 500 emails
    if (currentIndex % 500 === 0) {
      const totalFiles = Array.from(expectedCounts.values()).reduce((sum, count) => sum + count, 0);
      Logger.log(`Scanned ${currentIndex} threads, found ${totalFiles} expected files so far...`);
    }
    
    // Safety limit
    if (currentIndex > 20000) {
      Logger.log('Hit safety limit of 20,000 threads');
      return { expectedCounts, nextIndex: currentIndex, complete: true };
    }
  }
}

/**
 * Reset comprehensive audit state
 */
function resetComprehensiveAudit() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty('comprehensive_audit_state');
  
  SpreadsheetApp.getUi().alert(
    '� Audit Reset',
    'Comprehensive audit state has been cleared.\n\nYou can start a new audit from the menu.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  Logger.log('Comprehensive audit state reset');
}

/**
 * View comprehensive audit progress
 */
function viewComprehensiveAuditProgress() {
  const props = PropertiesService.getDocumentProperties();
  const stateJson = props.getProperty('comprehensive_audit_state');
  
  if (!stateJson) {
    SpreadsheetApp.getUi().alert(
      '��️ No Audit In Progress',
      'There is no comprehensive audit currently running.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const state = JSON.parse(stateJson);
  const expectedCount = Object.keys(JSON.parse(state.expectedCountsJson || '{}')).length;
  const actualCount = Object.keys(JSON.parse(state.actualCountsJson || '{}')).length;
  
  let message = `Phase: ${state.phase.toUpperCase()}\n\n`;
  
  if (state.phase === 'gmail_scan') {
    message += `Gmail threads scanned: ${state.gmailStartIndex}\n`;
    message += `Date/Network combinations found: ${expectedCount}\n\n`;
    message += 'Status: Scanning emails for attachments...';
  } else if (state.phase === 'drive_scan') {
    message += `Expected date/networks (from Gmail): ${expectedCount}\n`;
    message += `Actual date/networks (from Drive): ${actualCount}\n\n`;
    message += 'Status: Scanning Drive folders...';
  } else if (state.phase === 'compare') {
    message += `Expected date/networks: ${expectedCount}\n`;
    message += `Actual date/networks: ${actualCount}\n\n`;
    message += 'Status: Comparing and generating report...';
  }
  
  message += `\n\nStarted: ${new Date(state.startTime).toLocaleString()}`;
  
  SpreadsheetApp.getUi().alert(
    '📊 Comprehensive Audit Progress',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Scans Gmail for all "BKCM360 Global QA Check" emails and builds expected file map
 * LEGACY VERSION - Not used in chunked audit, kept for reference
 * Returns Map: "date|networkId|filename" => filename
 */
function scanGmailForExpectedFiles_() {
  const expectedFiles = new Map();
  const query = 'subject:"CM360 CPC/CPM FLIGHT QA"';
  let startIndex = 0;
  const batchSize = 100;
  
  while (true) {
    const threads = GmailApp.search(query, startIndex, batchSize);
    if (threads.length === 0) break;
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      
      for (const message of messages) {
        const messageDate = message.getDate();
        const dateStr = Utilities.formatDate(messageDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const attachments = message.getAttachments();
        
        for (const attachment of attachments) {
          const filename = attachment.getName();
          const lowerFilename = filename.toLowerCase();
          
          // Only count CSV/XLSX files (or files inside ZIPs)
          if (lowerFilename.endsWith('.zip')) {
            // Extract network ID from ZIP filename
            const networkId = extractNetworkIdFromFilename_(filename, getNetworkMap_());
            if (networkId) {
              try {
                const zipBlob = attachment.copyBlob();
                const unzipped = Utilities.unzip(zipBlob);
                
                for (const file of unzipped) {
                  const unzippedName = file.getName().toLowerCase();
                  if (unzippedName.endsWith('.csv') || unzippedName.endsWith('.xlsx')) {
                    const key = `${dateStr}|${networkId}|${file.getName()}`;
                    expectedFiles.set(key, file.getName());
                  }
                }
              } catch (e) {
                Logger.log(`Error unzipping ${filename}: ${e}`);
              }
            }
          } else if (lowerFilename.endsWith('.csv') || lowerFilename.endsWith('.xlsx')) {
            // Direct CSV/XLSX file
            const networkId = extractNetworkIdFromFilename_(filename, getNetworkMap_());
            if (networkId) {
              const key = `${dateStr}|${networkId}|${filename}`;
              expectedFiles.set(key, filename);
            }
          }
        }
      }
    }
    
    startIndex += batchSize;
    
    // Progress log every 500 emails
    if (startIndex % 500 === 0) {
      Logger.log(`Scanned ${startIndex} threads, found ${expectedFiles.size} expected files so far...`);
    }
    
    // Safety check: avoid infinite loop
    if (startIndex > 20000) {
      Logger.log('Hit safety limit of 20,000 threads');
      break;
    }
  }
  
  return expectedFiles;
}

/**
 * Scans Drive for all actual files in date folders
 * Returns Map: "date|networkId|filename" => filename
 */
/**
 * Scans Drive folders and builds actual file count map
 * Returns Map: "date|networkId" => count
 */
function scanDriveForActualCounts_() {
  const actualCounts = new Map();
  const rootFolderId = '1F53lLe3z5cup338IRY4nhTZQdUmJ9_wk'; // Raw Data folder ID
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  
  // Navigate: Raw Data/2025/05-May/2025-05-15/file.csv
  const yearFolders = rootFolder.getFolders();
  
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    const yearName = yearFolder.getName();
    
    // Skip Networks folder (categorized data)
    if (yearName === 'Networks') continue;
    
    // Process year folders
    if (/^\d{4}$/.test(yearName)) {
      const monthFolders = yearFolder.getFolders();
      
      while (monthFolders.hasNext()) {
        const monthFolder = monthFolders.next();
        const dateFolders = monthFolder.getFolders();
        
        while (dateFolders.hasNext()) {
          const dateFolder = dateFolders.next();
          const dateStr = dateFolder.getName(); // e.g., "2025-05-15"
          
          const files = dateFolder.getFiles();
          while (files.hasNext()) {
            const file = files.next();
            const filename = file.getName();
            
            // Extract network ID from filename
            const networkId = extractNetworkIdFromFilename_(filename, getNetworkMap_());
            if (networkId) {
              const key = `${dateStr}|${networkId}`;
              actualCounts.set(key, (actualCounts.get(key) || 0) + 1);
            }
          }
        }
      }
    }
  }
  
  return actualCounts;
}

/**
 * LEGACY: Scans Drive for actual files (kept for reference)
 * Returns Map: "date|networkId|filename" => filename
 */
function scanDriveForActualFiles_() {
  const actualFiles = new Map();
  const rootFolderId = '1F53lLe3z5cup338IRY4nhTZQdUmJ9_wk'; // Raw Data folder ID
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  
  // Navigate: Raw Data/2025/05-May/2025-05-15/file.csv
  const yearFolders = rootFolder.getFolders();
  
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    const yearName = yearFolder.getName();
    
    // Skip Networks folder (categorized data)
    if (yearName === 'Networks') continue;
    
    // Process year folders
    if (/^\d{4}$/.test(yearName)) {
      const monthFolders = yearFolder.getFolders();
      
      while (monthFolders.hasNext()) {
        const monthFolder = monthFolders.next();
        const dateFolders = monthFolder.getFolders();
        
        while (dateFolders.hasNext()) {
          const dateFolder = dateFolders.next();
          const dateStr = dateFolder.getName(); // e.g., "2025-05-15"
          
          const files = dateFolder.getFiles();
          while (files.hasNext()) {
            const file = files.next();
            const filename = file.getName();
            
            // Extract network ID from filename
            const networkId = extractNetworkIdFromFilename_(filename, getNetworkMap_());
            if (networkId) {
              const key = `${dateStr}|${networkId}|${filename}`;
              actualFiles.set(key, filename);
            }
          }
        }
      }
    }
  }
  
  return actualFiles;
}

/**
 * Sends comprehensive audit report with count-based comparison
 */
function sendComprehensiveAuditReportCounts_(totalExpected, totalActual, missingDateNetworks, extraDateNetworks, countMismatches) {
  const networkMap = getNetworkMap_();
  
  const hasIssues = missingDateNetworks.length > 0 || extraDateNetworks.length > 0 || countMismatches.length > 0;
  
  let htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 900px; margin: 0 auto;">
      <h2 style="color: ${hasIssues ? '#d93025' : '#1e8e3e'};">🔍 Comprehensive Archive Audit Report</h2>
      
      <h3>📊 Summary</h3>
      <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
        <tr style="background-color: #f0f0f0;">
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Expected Files (Gmail)</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${totalExpected.toLocaleString()}</td>
        </tr>
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Actual Files (Drive)</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${totalActual.toLocaleString()}</td>
        </tr>
        <tr style="background-color: ${totalExpected === totalActual ? '#d4edda' : '#fff3cd'};">
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Difference</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${totalActual - totalExpected}</td>
        </tr>
      </table>
      
      <h3>🔍 Issues Found</h3>
      <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
        <tr style="background-color: ${missingDateNetworks.length > 0 ? '#fff3cd' : '#d4edda'};">
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Missing Date/Networks</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${missingDateNetworks.length}</td>
        </tr>
        <tr style="background-color: ${extraDateNetworks.length > 0 ? '#f8d7da' : '#d4edda'};">
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Extra Date/Networks</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${extraDateNetworks.length}</td>
        </tr>
        <tr style="background-color: ${countMismatches.length > 0 ? '#fff3cd' : '#d4edda'};">
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Count Mismatches</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${countMismatches.length}</td>
        </tr>
      </table>
  `;
  
  // Missing date/networks section
  if (missingDateNetworks.length > 0) {
    htmlBody += `
      <h3>�� Missing Date/Networks (In Gmail, Not in Drive)</h3>
      <p>These dates have emails but no files saved in Drive:</p>
      <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 12px;">
        <thead>
          <tr style="background-color: #f0f0f0;">
            <th style="padding: 6px; border: 1px solid #ddd; text-align: left;">Date</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: left;">Network ID</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: left;">Network Name</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: right;">Expected Files</th>
          </tr>
        </thead>
        <tbody>
    `;
    
    missingDateNetworks.sort((a, b) => a.date.localeCompare(b.date));
    for (const item of missingDateNetworks) {
      const networkName = networkMap.get(item.networkId) || 'Unknown';
      htmlBody += `
        <tr>
          <td style="padding: 6px; border: 1px solid #ddd;">${item.date}</td>
          <td style="padding: 6px; border: 1px solid #ddd;">${item.networkId}</td>
          <td style="padding: 6px; border: 1px solid #ddd;">${networkName}</td>
          <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${item.expectedCount}</td>
        </tr>
      `;
    }
    
    htmlBody += `
        </tbody>
      </table>
    `;
  }
  
  // Extra date/networks section
  if (extraDateNetworks.length > 0) {
    htmlBody += `
      <h3>�� Extra Date/Networks (In Drive, Not in Gmail)</h3>
      <p>These dates have files in Drive but no corresponding emails:</p>
      <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 12px;">
        <thead>
          <tr style="background-color: #f0f0f0;">
            <th style="padding: 6px; border: 1px solid #ddd; text-align: left;">Date</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: left;">Network ID</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: left;">Network Name</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: right;">Actual Files</th>
          </tr>
        </thead>
        <tbody>
    `;
    
    extraDateNetworks.sort((a, b) => a.date.localeCompare(b.date));
    for (const item of extraDateNetworks) {
      const networkName = networkMap.get(item.networkId) || 'Unknown';
      htmlBody += `
        <tr>
          <td style="padding: 6px; border: 1px solid #ddd;">${item.date}</td>
          <td style="padding: 6px; border: 1px solid #ddd;">${item.networkId}</td>
          <td style="padding: 6px; border: 1px solid #ddd;">${networkName}</td>
          <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${item.actualCount}</td>
        </tr>
      `;
    }
    
    htmlBody += `
        </tbody>
      </table>
    `;
  }
  
  // Count mismatches section
  if (countMismatches.length > 0) {
    htmlBody += `
      <h3>��️ File Count Mismatches</h3>
      <p>These date/networks exist in both Gmail and Drive but have different file counts:</p>
      <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 12px;">
        <thead>
          <tr style="background-color: #f0f0f0;">
            <th style="padding: 6px; border: 1px solid #ddd; text-align: left;">Date</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: left;">Network ID</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: left;">Network Name</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: right;">Expected</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: right;">Actual</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: right;">Difference</th>
          </tr>
        </thead>
        <tbody>
    `;
    
    countMismatches.sort((a, b) => a.date.localeCompare(b.date) || a.networkId.localeCompare(b.networkId));
    for (const item of countMismatches) {
      const networkName = networkMap.get(item.networkId) || 'Unknown';
      const diffColor = item.difference > 0 ? '#d4edda' : '#f8d7da';
      htmlBody += `
        <tr>
          <td style="padding: 6px; border: 1px solid #ddd;">${item.date}</td>
          <td style="padding: 6px; border: 1px solid #ddd;">${item.networkId}</td>
          <td style="padding: 6px; border: 1px solid #ddd;">${networkName}</td>
          <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${item.expectedCount}</td>
          <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${item.actualCount}</td>
          <td style="padding: 6px; border: 1px solid #ddd; text-align: right; background-color: ${diffColor};">${item.difference > 0 ? '+' : ''}${item.difference}</td>
        </tr>
      `;
    }
    
    htmlBody += `
        </tbody>
      </table>
    `;
  }
  
  // Final status
  if (!hasIssues) {
    htmlBody += `
      <div style="background-color: #d4edda; border: 1px solid #c3e6cb; border-radius: 4px; padding: 15px; margin-top: 20px;">
        <h3 style="color: #155724; margin: 0;">� Archive is Complete!</h3>
        <p style="margin: 10px 0 0 0;">All Gmail attachments are properly saved in Drive with matching counts.</p>
      </div>
    `;
  } else {
    htmlBody += `
      <div style="background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 4px; padding: 15px; margin-top: 20px;">
        <h3 style="color: #856404; margin: 0;">��️ Action Required</h3>
        <p style="margin: 10px 0 0 0;">Please review the issues above and use the gap-fill archive tool to correct missing data.</p>
      </div>
    `;
  }
  
  htmlBody += `
    </div>
  `;
  
  const subject = hasIssues 
    ? '��️ Comprehensive Archive Audit Complete (Issues Found)'
    : '� Comprehensive Archive Audit Complete';
  
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: subject,
    htmlBody: htmlBody
  });
  
  Logger.log('Comprehensive audit report sent');
}

/**
 * LEGACY: Sends comprehensive audit report with file-level details
 */
function sendComprehensiveAuditReport_(expectedCount, actualCount, missingFiles, extraFiles) {
  const networkMap = getNetworkMap_();
  
  // Group missing files by date
  const missingByDate = {};
  for (const item of missingFiles) {
    if (!missingByDate[item.date]) {
      missingByDate[item.date] = [];
    }
    missingByDate[item.date].push(item);
  }
  
  // Group extra files by date
  const extraByDate = {};
  for (const item of extraFiles) {
    if (!extraByDate[item.date]) {
      extraByDate[item.date] = [];
    }
    extraByDate[item.date].push(item);
  }
  
  let htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 900px; margin: 0 auto;">
      <h2 style="color: #1a73e8;">🔍 Comprehensive Archive Audit Report</h2>
      
      <h3>📊 Summary</h3>
      <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
        <tr style="background-color: #f0f0f0;">
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Expected Files (Gmail)</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${expectedCount}</td>
        </tr>
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Actual Files (Drive)</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${actualCount}</td>
        </tr>
        <tr style="background-color: ${missingFiles.length > 0 ? '#fff3cd' : '#d4edda'};">
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Missing Files</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${missingFiles.length}</td>
        </tr>
        <tr style="background-color: ${extraFiles.length > 0 ? '#f8d7da' : '#d4edda'};">
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>Extra Files (not in Gmail)</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${extraFiles.length}</td>
        </tr>
      </table>
  `;
  
  // Missing files section
  if (missingFiles.length > 0) {
    htmlBody += `
      <h3>��️ Missing Files (In Gmail, Not in Drive)</h3>
      <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
        <tr style="background-color: #f0f0f0;">
          <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Date</th>
          <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Network</th>
          <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Filename</th>
        </tr>
    `;
    
    // Show first 500 missing files (prevent email size issues)
    const displayMissing = missingFiles.slice(0, 500);
    for (const item of displayMissing) {
      const networkName = networkMap[item.networkId] || 'Unknown';
      htmlBody += `
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;">${item.date}</td>
          <td style="padding: 8px; border: 1px solid #ddd;">${item.networkId} - ${networkName}</td>
          <td style="padding: 8px; border: 1px solid #ddd; font-family: monospace; font-size: 11px;">${item.filename}</td>
        </tr>
      `;
    }
    
    if (missingFiles.length > 500) {
      htmlBody += `
        <tr>
          <td colspan="3" style="padding: 8px; border: 1px solid #ddd; background-color: #fff3cd;">
            ... and ${missingFiles.length - 500} more missing files (showing first 500)
          </td>
        </tr>
      `;
    }
    
    htmlBody += '</table>';
  } else {
    htmlBody += '<p style="color: green;">� No missing files! All Gmail attachments are in Drive.</p>';
  }
  
  // Extra files section
  if (extraFiles.length > 0) {
    htmlBody += `
      <h3>��️ Extra Files (In Drive, Not in Gmail)</h3>
      <p style="color: #856404; background-color: #fff3cd; padding: 10px; border-radius: 4px;">
        These files exist in Drive but were not found in Gmail. This could be due to:
        <ul>
          <li>Emails deleted after archiving</li>
          <li>Manual file uploads</li>
          <li>Files from other sources</li>
        </ul>
      </p>
      <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
        <tr style="background-color: #f0f0f0;">
          <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Date</th>
          <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Network</th>
          <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Filename</th>
        </tr>
    `;
    
    // Show first 100 extra files
    const displayExtra = extraFiles.slice(0, 100);
    for (const item of displayExtra) {
      const networkName = networkMap[item.networkId] || 'Unknown';
      htmlBody += `
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;">${item.date}</td>
          <td style="padding: 8px; border: 1px solid #ddd;">${item.networkId} - ${networkName}</td>
          <td style="padding: 8px; border: 1px solid #ddd; font-family: monospace; font-size: 11px;">${item.filename}</td>
        </tr>
      `;
    }
    
    if (extraFiles.length > 100) {
      htmlBody += `
        <tr>
          <td colspan="3" style="padding: 8px; border: 1px solid #ddd; background-color: #fff3cd;">
            ... and ${extraFiles.length - 100} more extra files (showing first 100)
          </td>
        </tr>
      `;
    }
    
    htmlBody += '</table>';
  }
  
  htmlBody += `
      <h3>📋 Next Steps</h3>
      <ul>
        <li>Review missing files and run gap-fill archive to retrieve them</li>
        <li>Extra files can generally be ignored unless you suspect data corruption</li>
      </ul>
      
      <hr style="border: 1px solid #ddd; margin: 20px 0;">
      <p style="color: #666; font-size: 12px;">Report generated: ${new Date().toLocaleString()}</p>
    </div>
  `;
  
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: `[CM360 Archive] Comprehensive Audit Report - ${missingFiles.length} Missing, ${extraFiles.length} Extra`,
    htmlBody: htmlBody
  });
}

/**
 * AUDIT FUNCTION: Validates archive completeness by checking for missing files
 * Compares expected files (Networks � Date Range) vs actual files in Drive
 * Generates detailed report of gaps for manual retrieval
 */
function auditRawDataArchive() {
  try {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      '🔍 Audit Raw Data Archive',
      'This will scan your Drive to identify missing files.\n\n' +
      'Expected files = All Networks � All Dates in archive period.\n' +
      'This may take several minutes.\n\nProceed?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      ui.alert('Audit cancelled.');
      return;
    }
    
    ui.alert('Starting audit... This will take a few minutes. Check your email for results.');
    
    // Get all networks
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var networksSheet = ss.getSheetByName('Networks');
    if (!networksSheet) {
      throw new Error('Networks sheet not found');
    }
    
    var networkData = networksSheet.getDataRange().getValues();
    var networkMap = {};
    
    for (var i = 1; i < networkData.length; i++) {
      var networkId = String(networkData[i][0]).trim();
      var networkName = String(networkData[i][1]).trim();
      if (networkId && networkName) {
        networkMap[networkId] = networkName;
      }
    }
    
    Logger.log('Found ' + Object.keys(networkMap).length + ' networks');
    
    // Scan Drive for existing files
    var rootFolderId = PropertiesService.getScriptProperties().getProperty('RAW_DATA_FOLDER_ID');
    if (!rootFolderId) {
      throw new Error('Raw Data folder ID not found in Script Properties');
    }
    
    var rootFolder = DriveApp.getFolderById(rootFolderId);
    var existingFiles = scanAllFilesInDrive_(rootFolder);
    
    Logger.log('Found ' + existingFiles.size + ' unique date|network combinations in Drive');
    
    // Build expected file list from existing dates
    var expectedFiles = buildExpectedFileList_(rootFolder, networkMap);
    
    Logger.log('Expected ' + expectedFiles.length + ' files based on date folders');
    
    // Compare and identify gaps
    var missingFiles = [];
    var foundFiles = [];
    
    for (var i = 0; i < expectedFiles.length; i++) {
      var expected = expectedFiles[i];
      var key = expected.date + '|' + expected.networkId;
      
      if (existingFiles.has(key)) {
        foundFiles.push(expected);
      } else {
        missingFiles.push(expected);
      }
    }
    
    // Send detailed audit report
    sendAuditReport_(networkMap, expectedFiles.length, foundFiles.length, missingFiles);
    
    ui.alert(
      '� Audit Complete',
      'Found: ' + foundFiles.length + ' files\n' +
      'Missing: ' + missingFiles.length + ' files\n\n' +
      'Detailed report sent to your email.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('Error in auditRawDataArchive: ' + error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Scans all files in Drive and returns a Set of date|networkId keys
 */
function scanAllFilesInDrive_(rootFolder) {
  var fileKeys = new Set();
  
  // Check both date-based structure and network-based structure
  var yearFolders = rootFolder.getFolders();
  
  while (yearFolders.hasNext()) {
    var yearFolder = yearFolders.next();
    var yearName = yearFolder.getName();
    
    // Skip Networks folder (that's categorized data)
    if (yearName === 'Networks') {
      Logger.log('Skipping Networks folder during audit');
      continue;
    }
    
    // Process year folders (2024, 2025, etc.)
    if (/^\d{4}$/.test(yearName)) {
      var monthFolders = yearFolder.getFolders();
      
      while (monthFolders.hasNext()) {
        var monthFolder = monthFolders.next();
        var dateFolders = monthFolder.getFolders();
        
        while (dateFolders.hasNext()) {
          var dateFolder = dateFolders.next();
          var dateName = dateFolder.getName(); // e.g., "2025-05-15"
          
          var files = dateFolder.getFiles();
          while (files.hasNext()) {
            var file = files.next();
            var fileName = file.getName();
            
            // Extract network ID from filename
            var networkMap = getNetworkMap_();
            var networkId = extractNetworkIdFromFilename_(fileName, networkMap);
            if (networkId) {
              var key = dateName + '|' + networkId;
              fileKeys.add(key);
            }
          }
        }
      }
    }
  }
  
  return fileKeys;
}

/**
 * Helper to get network map for audit functions
 */
function getNetworkMap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var networksSheet = ss.getSheetByName('Networks');
  if (!networksSheet) {
    return {};
  }
  
  var networkData = networksSheet.getDataRange().getValues();
  var networkMap = {};
  
  for (var i = 1; i < networkData.length; i++) {
    var networkId = String(networkData[i][0]).trim();
    var networkName = String(networkData[i][1]).trim();
    if (networkId && networkName) {
      networkMap[networkId] = networkName;
    }
  }
  
  return networkMap;
}

/**
 * Builds expected file list based on actual date folders in Drive
 */
function buildExpectedFileList_(rootFolder, networkMap) {
  var expectedFiles = [];
  var dates = new Set();
  
  // Scan for all date folders
  var yearFolders = rootFolder.getFolders();
  
  while (yearFolders.hasNext()) {
    var yearFolder = yearFolders.next();
    var yearName = yearFolder.getName();
    
    if (yearName === 'Networks' || !/^\d{4}$/.test(yearName)) {
      continue;
    }
    
    var monthFolders = yearFolder.getFolders();
    
    while (monthFolders.hasNext()) {
      var monthFolder = monthFolders.next();
      var dateFolders = monthFolder.getFolders();
      
      while (dateFolders.hasNext()) {
        var dateFolder = dateFolders.next();
        dates.add(dateFolder.getName()); // e.g., "2025-05-15"
      }
    }
  }
  
  Logger.log('Found ' + dates.size + ' unique dates in Drive');
  
  // For each date � network, expect a file
  var dateArray = Array.from(dates);
  var networkIds = Object.keys(networkMap);
  
  for (var i = 0; i < dateArray.length; i++) {
    for (var j = 0; j < networkIds.length; j++) {
      expectedFiles.push({
        date: dateArray[i],
        networkId: networkIds[j],
        networkName: networkMap[networkIds[j]]
      });
    }
  }
  
  return expectedFiles;
}

/**
 * Sends detailed audit report via email
 */
function sendAuditReport_(networkMap, expectedCount, foundCount, missingFiles) {
  try {
    var missingByNetwork = {};
    var missingByDate = {};
    
    // Group missing files by network and date
    for (var i = 0; i < missingFiles.length; i++) {
      var missing = missingFiles[i];
      
      // By network
      if (!missingByNetwork[missing.networkId]) {
        missingByNetwork[missing.networkId] = {
          name: missing.networkName,
          dates: []
        };
      }
      missingByNetwork[missing.networkId].dates.push(missing.date);
      
      // By date
      if (!missingByDate[missing.date]) {
        missingByDate[missing.date] = [];
      }
      missingByDate[missing.date].push(missing.networkId + ' - ' + missing.networkName);
    }
    
    // Sort networks by most missing files
    var networkBreakdown = Object.keys(missingByNetwork).map(function(netId) {
      return {
        id: netId,
        name: missingByNetwork[netId].name,
        missingCount: missingByNetwork[netId].dates.length,
        dates: missingByNetwork[netId].dates.sort()
      };
    }).sort(function(a, b) { return b.missingCount - a.missingCount; });
    
    // Sort dates
    var dateBreakdown = Object.keys(missingByDate).sort().map(function(date) {
      return {
        date: date,
        networks: missingByDate[date].sort()
      };
    });
    
    var htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 800px;">
        <h2 style="color: #1a73e8;">🔍 Raw Data Archive Audit Report</h2>
        <p><strong>Generated:</strong> ${new Date().toLocaleString()}</p>
        
        <h3>📊 Summary</h3>
        <table style="border-collapse: collapse; width: 100%;">
          <tr style="background-color: #f0f0f0;">
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Expected Files</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd;">${expectedCount}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Files Found</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd; color: green;">${foundCount} (${((foundCount / expectedCount) * 100).toFixed(1)}%)</td>
          </tr>
          <tr style="background-color: #f0f0f0;">
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Files Missing</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd; color: ${missingFiles.length > 0 ? 'red' : 'green'};">${missingFiles.length} (${((missingFiles.length / expectedCount) * 100).toFixed(1)}%)</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Networks</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd;">${Object.keys(networkMap).length}</td>
          </tr>
          <tr style="background-color: #f0f0f0;">
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Networks with Gaps</strong></td>
            <td style="padding: 8px; border: 1px solid #ddd;">${Object.keys(missingByNetwork).length}</td>
          </tr>
        </table>
        
        ${missingFiles.length > 0 ? `
        <h3 style="color: #d93025;">��️ Missing Files by Network</h3>
        <p style="color: #666; font-size: 12px;">Networks sorted by most missing files (showing top 20):</p>
        <table style="border-collapse: collapse; width: 100%; font-size: 12px;">
          <tr style="background-color: #f0f0f0;">
            <th style="padding: 6px; border: 1px solid #ddd; text-align: left;">Network</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: right;">Missing</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: left;">Sample Dates</th>
          </tr>
          ${networkBreakdown.slice(0, 20).map((net, i) => `
            <tr${i % 2 === 0 ? ' style="background-color: #f9f9f9;"' : ''}>
              <td style="padding: 6px; border: 1px solid #ddd;">${net.id} - ${net.name}</td>
              <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${net.missingCount}</td>
              <td style="padding: 6px; border: 1px solid #ddd; font-size: 11px;">${net.dates.slice(0, 5).join(', ')}${net.dates.length > 5 ? '...' : ''}</td>
            </tr>
          `).join('')}
        </table>
        ${networkBreakdown.length > 20 ? `<p style="color: #666; font-size: 12px;">...and ${networkBreakdown.length - 20} more networks with missing files</p>` : ''}
        
        <h3 style="color: #d93025;">📅 Missing Files by Date</h3>
        <p style="color: #666; font-size: 12px;">Dates with missing files (showing first 10):</p>
        <table style="border-collapse: collapse; width: 100%; font-size: 12px;">
          <tr style="background-color: #f0f0f0;">
            <th style="padding: 6px; border: 1px solid #ddd; text-align: left;">Date</th>
            <th style="padding: 6px; border: 1px solid #ddd; text-align: right;">Networks Missing</th>
          </tr>
          ${dateBreakdown.slice(0, 10).map((d, i) => `
            <tr${i % 2 === 0 ? ' style="background-color: #f9f9f9;"' : ''}>
              <td style="padding: 6px; border: 1px solid #ddd;">${d.date}</td>
              <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${d.networks.length}</td>
            </tr>
          `).join('')}
        </table>
        ${dateBreakdown.length > 10 ? `<p style="color: #666; font-size: 12px;">...and ${dateBreakdown.length - 10} more dates</p>` : ''}
        
        <h3>🔧 Next Steps to Fill Gaps</h3>
        <ol>
          <li><strong>Search Gmail for missing dates:</strong> <code>subject:"BKCM360 Global QA Check" after:YYYY-MM-DD before:YYYY-MM-DD</code></li>
          <li><strong>Download missing CSV/ZIP files</strong> from those emails manually</li>
          <li><strong>Upload to Drive:</strong> Place in correct folders (Raw Data/YYYY/Month/YYYY-MM-DD/)</li>
          <li><strong>Re-run categorization:</strong> Organize new files by network</li>
          <li><strong>Re-audit:</strong> Run this audit again to verify gaps are filled</li>
        </ol>
        ` : `
        <h3 style="color: #34a853;">� Archive Complete!</h3>
        <p>All expected files are present in your Drive. No gaps detected.</p>
        <p style="color: #666; font-size: 12px;">Note: This assumes 1 file per network per date. Some networks may legitimately have no data for certain dates.</p>
        `}
        
        <hr style="border: 1px solid #ddd; margin: 20px 0;">
        <p style="color: #666; font-size: 12px;">This audit scans all date folders and expects 1 file per network per date. Files saved in network folders (categorized) are also counted.</p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: '[CM360 Archive] Audit Report - ' + (missingFiles.length > 0 ? missingFiles.length + ' Files Missing' : 'Complete'),
      htmlBody: htmlBody
    });
    
  } catch (error) {
    Logger.log('Error sending audit report: ' + error);
  }
}

// =====================================================================================================================
// ======================================= GAP-FILLING ARCHIVE FUNCTIONS ===============================================
// =====================================================================================================================

// ---------------------
// HELPER: Check if file exists in folder
// ---------------------
function fileExistsInFolder_(folder, filename) {
  const files = folder.getFilesByName(filename);
  return files.hasNext();
}

/**
 * Archive only missing dates (gap-filling)
 * Identifies which dates are missing from Drive and archives only those
 */
function archiveRawDataGapFill() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '🔍 Gap-Filling Archive',
    'This will archive ONLY the missing dates from May 1 to today.\n\n' +
    'Based on your existing data, you have ~90 missing dates.\n\n' +
    'Estimated:\n' +
    '�� ~1,620 emails to process\n' +
    '�� ~2-3 hours duration\n\n' +
    'The system will skip dates that already have data.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    return;
  }
  
  // Generate list of ALL dates from May 1, 2025 to today
  const allDates = generateDateRange_(new Date(2025, 4, 1), new Date()); // May is month 4 (0-indexed)
  
  ui.alert(
    '� Ready to Start',
    `Will check ${allDates.length} dates from May 1 to today.\n\n` +
    'Dates that already have files will be skipped automatically.\n\n' +
    'Click OK to start.',
    ui.ButtonSet.OK
  );
  
  // Initialize state
  const props = PropertiesService.getScriptProperties();
  const state = {
    status: 'running',
    mode: 'gap-fill',
    allDates: allDates.map(d => formatDateForFolder_(d)),
    currentDateIndex: 0,
    startTime: new Date().toISOString(),
    emailsProcessed: 0,
    filesSaved: 0,
    datesCompleted: 0,
    datesSkipped: 0
  };
  
  props.setProperty('RAW_ARCHIVE_STATE', JSON.stringify(state));
  
  // Start processing
  processGapFillBatch_();
  
  ui.alert(
    '� Gap-Fill Archive Started',
    `Processing ${allDates.length} dates (will skip existing).\n\n` +
    'Create an Auto-Resume trigger to continue automatically every 10 minutes.',
    ui.ButtonSet.OK
  );
}

/**
 * Process one batch of gap-fill archive
 * Processes emails for a few missing dates at a time
 */
function processGapFillBatch_() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('RAW_ARCHIVE_STATE');
  
  if (!stateJson) {
    Logger.log('No gap-fill state found');
    return;
  }
  
  const state = JSON.parse(stateJson);
  
  if (state.mode !== 'gap-fill') {
    Logger.log('Not in gap-fill mode');
    return;
  }
  
  const startTime = new Date();
  const MAX_EXECUTION_MS = 5 * 60 * 1000; // 5 minutes
  
  Logger.log(`Gap-fill batch starting at date index ${state.currentDateIndex}/${state.allDates.length}`);
  
  // Process dates until we run out of time
  while (state.currentDateIndex < state.allDates.length) {
    const elapsed = new Date() - startTime;
    if (elapsed > MAX_EXECUTION_MS) {
      Logger.log('Time limit reached, saving progress...');
      break;
    }
    
    const dateStr = state.allDates[state.currentDateIndex];
    const date = parseFolderDateString_(dateStr);
    
    // Check if this date already has data (skip if so)
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    
    if (dateAlreadyHasFiles_(year, month, date)) {
      Logger.log(`Skipping ${dateStr} - already has files`);
      state.datesSkipped++;
      state.currentDateIndex++;
      props.setProperty('RAW_ARCHIVE_STATE', JSON.stringify(state));
      continue;
    }
    
    Logger.log(`Processing missing date: ${dateStr}`);
    
    // Search for emails on this specific date
    const result = archiveSingleDate_(date);
    
    state.emailsProcessed += result.emailsProcessed;
    state.filesSaved += result.filesSaved;
    state.datesCompleted++;
    state.currentDateIndex++;
    
    // Save progress after each date
    props.setProperty('RAW_ARCHIVE_STATE', JSON.stringify(state));
    
    Logger.log(`Completed ${dateStr}: ${result.emailsProcessed} emails, ${result.filesSaved} files`);
  }
  
  // Check if complete
  if (state.currentDateIndex >= state.allDates.length) {
    state.status = 'completed';
    state.completedTime = new Date().toISOString();
    props.setProperty('RAW_ARCHIVE_STATE', JSON.stringify(state));
    
    // Send completion email
    sendGapFillCompletionEmail_(state);
    
    Logger.log('� Gap-fill archive COMPLETED!');
  } else {
    const progress = ((state.currentDateIndex/state.allDates.length)*100).toFixed(1);
    Logger.log(`Progress: ${state.currentDateIndex}/${state.allDates.length} dates checked (${progress}%) - ${state.datesCompleted} archived, ${state.datesSkipped} skipped`);
  }
}

/**
 * Archive emails for a single specific date
 */
function archiveSingleDate_(date) {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  
  // Format date for Gmail search: "2025/05/12"
  const searchDate = `${year}/${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}`;
  
  // Search for emails on this specific date
  const query = `subject:"CM360 CPC/CPM FLIGHT QA" after:${searchDate} before:${getNextDay_(searchDate)}`;
  
  Logger.log(`Searching: ${query}`);
  
  const threads = GmailApp.search(query, 0, 10); // Max 10 threads per day (safety)
  
  let emailsProcessed = 0;
  let filesSaved = 0;
  
  for (const thread of threads) {
    const messages = thread.getMessages();
    
    for (const message of messages) {
      const subject = message.getSubject();
      const messageDate = message.getDate();
      
      // Double-check the date matches (Gmail search can be fuzzy)
      if (messageDate.getFullYear() !== year || 
          messageDate.getMonth() + 1 !== month || 
          messageDate.getDate() !== day) {
        continue;
      }
      
      // Get folder for this date
      const monthFolder = getOrCreateRawDataMonthFolder_(year, month);
      const dateStr = formatDateForFolder_(date);
      const dateFolder = getOrCreateDateFolder_(monthFolder, dateStr);
      
      // Process attachments
      const attachments = message.getAttachments();
      
      for (const attachment of attachments) {
        const filename = attachment.getName();
        const lowerFilename = filename.toLowerCase();
        
        // Skip if already exists (duplicate protection)
        if (fileExistsInFolder_(dateFolder, filename)) {
          Logger.log(`Skipping duplicate: ${filename}`);
          continue;
        }
        
        // Handle ZIP files
        if (lowerFilename.endsWith('.zip')) {
          const zipBlob = attachment.copyBlob();
          const unzipped = Utilities.unzip(zipBlob);
          
          for (const file of unzipped) {
            const unzippedName = file.getName().toLowerCase();
            if (unzippedName.endsWith('.csv') || unzippedName.endsWith('.xlsx')) {
              if (!fileExistsInFolder_(dateFolder, file.getName())) {
                dateFolder.createFile(file);
                filesSaved++;
                Logger.log(`Saved from ZIP: ${file.getName()}`);
              }
            }
          }
        }
        // Handle CSV and XLSX files
        else if (lowerFilename.endsWith('.csv') || lowerFilename.endsWith('.xlsx')) {
          dateFolder.createFile(attachment.copyBlob());
          filesSaved++;
          Logger.log(`Saved: ${filename}`);
        }
      }
      
      emailsProcessed++;
    }
  }
  
  return {
    emailsProcessed: emailsProcessed,
    filesSaved: filesSaved
  };
}

/**
 * Check if a specific date already has files in Drive
 */
function dateAlreadyHasFiles_(year, month, date) {
  try {
    const monthFolder = getOrCreateRawDataMonthFolder_(year, month);
    const dateStr = formatDateForFolder_(date);
    
    // Check if date folder exists and has files
    const dateFolders = monthFolder.getFoldersByName(dateStr);
    if (!dateFolders.hasNext()) {
      return false; // Folder doesn't exist
    }
    
    const dateFolder = dateFolders.next();
    const files = dateFolder.getFiles();
    
    // Return true if folder has at least one file
    return files.hasNext();
    
  } catch (error) {
    Logger.log(`Error checking if ${formatDateForFolder_(date)} has files: ${error}`);
    return false; // Assume doesn't exist on error
  }
}

/**
 * Scan existing Drive data to find which dates already exist
 * Returns a Set of date strings like "2025-05-12"
 */
function scanExistingRawDataDates_() {
  const existingDates = new Set();
  
  try {
    const folderId = '1F53lLe3z5cup338IRY4nhTZQdUmJ9_wk';
    const rawDataFolder = DriveApp.getFolderById(folderId);
    
    // Loop through year folders
    const yearFolders = rawDataFolder.getFolders();
    while (yearFolders.hasNext()) {
      const yearFolder = yearFolders.next();
      
      // Loop through month folders
      const monthFolders = yearFolder.getFolders();
      while (monthFolders.hasNext()) {
        const monthFolder = monthFolders.next();
        
        // Loop through day folders
        const dayFolders = monthFolder.getFolders();
        while (dayFolders.hasNext()) {
          const dayFolder = dayFolders.next();
          const dayName = dayFolder.getName(); // e.g., "2025-05-12"
          existingDates.add(dayName);
        }
      }
    }
    
    Logger.log(`Found ${existingDates.size} existing dates in Drive`);
    
  } catch (error) {
    Logger.log('Error scanning existing dates: ' + error);
  }
  
  return existingDates;
}

/**
 * Generate array of dates between start and end (inclusive)
 */
function generateDateRange_(startDate, endDate) {
  const dates = [];
  const current = new Date(startDate);
  
  while (current <= endDate) {
    dates.push(new Date(current));
    current.setDate(current.getDate() + 1);
  }
  
  return dates;
}

/**
 * Format date for folder name: "2025-05-12"
 */
function formatDateForFolder_(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Format date for display: "May 12, 2025"
 */
function formatDateForDisplay_(date) {
  return date.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
}

/**
 * Parse folder date string back to Date object
 */
function parseFolderDateString_(dateStr) {
  const parts = dateStr.split('-');
  return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
}

/**
 * Get next day in Gmail search format
 */
function getNextDay_(gmailDateStr) {
  const parts = gmailDateStr.split('/');
  const date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  date.setDate(date.getDate() + 1);
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  return `${year}/${month}/${day}`;
}

/**
 * Send completion email for gap-fill archive
 */
function sendGapFillCompletionEmail_(state) {
  const startTime = new Date(state.startTime);
  const endTime = new Date(state.completedTime);
  const duration = endTime - startTime;
  const hours = Math.floor(duration / (1000 * 60 * 60));
  const minutes = Math.floor((duration % (1000 * 60 * 60)) / (1000 * 60));
  
  MailApp.sendEmail({
    to: 'platformsolutionsadopshorizon@gmail.com',
    subject: '� CM360 Gap-Fill Archive COMPLETED',
    htmlBody: `
      <h2>🎉 Gap-Fill Archive Complete!</h2>
      
      <h3>📊 Summary</h3>
      <ul>
        <li><strong>Missing Dates Filled:</strong> ${state.datesCompleted}</li>
        <li><strong>Emails Processed:</strong> ${state.emailsProcessed}</li>
        <li><strong>Files Saved:</strong> ${state.filesSaved}</li>
        <li><strong>Duration:</strong> ${hours}h ${minutes}m</li>
      </ul>
      
      <h3>� Next Steps</h3>
      <ol>
        <li>Delete Auto-Resume trigger (if created)</li>
        <li>Run audit to verify all dates present</li>
        <li>Categorize by network</li>
        <li>Begin ROI analysis</li>
      </ol>
      
      <p><strong>Archive Start:</strong> ${startTime.toLocaleString()}</p>
      <p><strong>Archive End:</strong> ${endTime.toLocaleString()}</p>
    `
  });
}

/**
 * DIAGNOSTIC: Check what's actually in the Raw Data Drive folder by FOLDER ID
 * Checks nested structure: Raw Data/Year/Month/Day/files
 */
function checkDriveRawDataFolder() {
  try {
    // Use the specific folder ID from the user's Drive link
    const folderId = '1F53lLe3z5cup338IRY4nhTZQdUmJ9_wk';
    const rawDataFolder = DriveApp.getFolderById(folderId);
    
    let totalFiles = 0;
    let totalDays = 0;
    let folderStructure = {};
    
    Logger.log('=== RAW DATA FOLDER ANALYSIS ===');
    Logger.log('Folder ID: ' + folderId);
    Logger.log('Folder Name: ' + rawDataFolder.getName());
    
    // Check year folders
    const yearFolders = rawDataFolder.getFolders();
    while (yearFolders.hasNext()) {
      const yearFolder = yearFolders.next();
      const yearName = yearFolder.getName();
      folderStructure[yearName] = {};
      
      Logger.log('\n📅 Year: ' + yearName);
      
      // Check month folders
      const monthFolders = yearFolder.getFolders();
      while (monthFolders.hasNext()) {
        const monthFolder = monthFolders.next();
        const monthName = monthFolder.getName();
        
        let monthFileCount = 0;
        let monthDayCount = 0;
        let dayFolders = [];
        
        // Check for DAY subfolders within each month
        const dayFoldersIterator = monthFolder.getFolders();
        while (dayFoldersIterator.hasNext()) {
          const dayFolder = dayFoldersIterator.next();
          const dayName = dayFolder.getName();
          dayFolders.push(dayName);
          
          // Count files in this day folder
          const files = dayFolder.getFiles();
          let dayFileCount = 0;
          while (files.hasNext()) {
            files.next();
            dayFileCount++;
            monthFileCount++;
            totalFiles++;
          }
          
          if (dayFileCount > 0) {
            monthDayCount++;
            totalDays++;
          }
        }
        
        // Sort day folders
        dayFolders.sort();
        
        folderStructure[yearName][monthName] = {
          totalFiles: monthFileCount,
          totalDays: monthDayCount,
          days: dayFolders
        };
        
        if (monthFileCount > 0) {
          Logger.log('  📁 ' + monthName + ': ' + monthFileCount + ' files across ' + monthDayCount + ' days');
          Logger.log('     Days: ' + dayFolders.join(', '));
        }
      }
    }
    
    // Summary
    Logger.log('\n=== SUMMARY ===');
    Logger.log('� Total Files: ' + totalFiles);
    Logger.log('� Total Days with Data: ' + totalDays);
    
    // Detailed breakdown
    Logger.log('\n=== DETAILED BREAKDOWN ===');
    for (const year in folderStructure) {
      for (const month in folderStructure[year]) {
        const data = folderStructure[year][month];
        if (data.totalFiles > 0) {
          Logger.log(month + ': ' + data.totalFiles + ' files, ' + data.totalDays + ' days');
          Logger.log('  Days present: ' + data.days.join(', '));
        }
      }
    }
    
    // Recommendation
    Logger.log('\n=== RECOMMENDATION ===');
    if (totalFiles > 0) {
      Logger.log('� You have ' + totalFiles + ' files from ' + totalDays + ' days already archived!');
      Logger.log('📋 NEXT STEP: Run audit to identify missing dates');
      Logger.log('�� Then archive ONLY the missing dates (much faster than re-doing everything)');
    } else {
      Logger.log('📁 Folders exist but empty - need to run full archive');
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Found ' + totalFiles + ' files across ' + totalDays + ' days', 
      '📊 Drive Analysis Complete', 
      15
    );
    
    return {
      totalFiles: totalFiles,
      totalDays: totalDays,
      structure: folderStructure
    };
    
  } catch (error) {
    Logger.log('Error checking Drive folder: ' + error);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + error.message, '�� Check Failed', 10);
  }
}

// =====================================================================================================================
// ======================================= END RAW DATA ARCHIVE SYSTEM ================================================
// =====================================================================================================================

// =====================================================================================================================
// ========================================== TIME MACHINE SYSTEM ======================================================
// =====================================================================================================================

/**
 * Setup Time Machine sheet with date picker and run button
 */
function setupTimeMachineSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Time Machine");
  
  if (!sheet) {
    sheet = ss.insertSheet("Time Machine");
  }
  
  // Clear existing content
  sheet.clear();
  
  // Set up the interface
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 400);
  
  // Title
  sheet.getRange("A1:C1").merge();
  sheet.getRange("A1").setValue("�� TIME MACHINE - Run QA for Past Dates")
    .setFontSize(16)
    .setFontWeight("bold")
    .setBackground("#4285f4")
    .setFontColor("#ffffff")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(1, 40);
  
  // Instructions
  sheet.getRange("A2:C2").merge();
  sheet.getRange("A2").setValue("Select a date below and click 'Run QA' to process that day's data from Gmail")
    .setFontSize(11)
    .setBackground("#e8f0fe")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(2, 30);
  
  // Date selector row
  sheet.getRange("A4").setValue("Select Date:")
    .setFontWeight("bold")
    .setVerticalAlignment("middle");
  
  // Set default date to yesterday
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  sheet.getRange("B4").setValue(yesterday)
    .setNumberFormat("yyyy-mm-dd")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  
  // Add data validation for date
  const dateValidation = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText("Select a date to run QA for")
    .build();
  sheet.getRange("B4").setDataValidation(dateValidation);
  
  // Status row
  sheet.getRange("A6").setValue("Status:")
    .setFontWeight("bold")
    .setVerticalAlignment("middle");
  sheet.getRange("B6:C6").merge();
  sheet.getRange("B6").setValue("Ready - Select a date and run QA from the menu")
    .setFontColor("#666666")
    .setVerticalAlignment("middle");
  
  // Last run info
  sheet.getRange("A8").setValue("Last Run:")
    .setFontWeight("bold")
    .setVerticalAlignment("middle");
  sheet.getRange("B8:C8").merge();
  sheet.getRange("B8").setValue("Never")
    .setFontColor("#666666")
    .setVerticalAlignment("middle");
  
  // Results summary
  sheet.getRange("A10:C10").merge();
  sheet.getRange("A10").setValue("Last Run Results")
    .setFontSize(12)
    .setFontWeight("bold")
    .setBackground("#f0f0f0")
    .setVerticalAlignment("middle");
  
  sheet.getRange("A11").setValue("Files Processed:")
    .setFontWeight("bold")
    .setVerticalAlignment("middle");
  sheet.getRange("B11").setValue("��")
    .setVerticalAlignment("middle");
  
  sheet.getRange("A12").setValue("Placements Checked:")
    .setFontWeight("bold")
    .setVerticalAlignment("middle");
  sheet.getRange("B12").setValue("��")
    .setVerticalAlignment("middle");
  
  sheet.getRange("A13").setValue("Violations Found:")
    .setFontWeight("bold")
    .setVerticalAlignment("middle");
  sheet.getRange("B13").setValue("��")
    .setVerticalAlignment("middle");
  
  sheet.getRange("A14").setValue("Report Saved:")
    .setFontWeight("bold")
    .setVerticalAlignment("middle");
  sheet.getRange("B14:C14").merge();
  sheet.getRange("B14").setValue("��")
    .setVerticalAlignment("middle");
  
  // Instructions section
  sheet.getRange("A16:C16").merge();
  sheet.getRange("A16").setValue("How to Use")
    .setFontSize(12)
    .setFontWeight("bold")
    .setBackground("#f0f0f0")
    .setVerticalAlignment("middle");
  
  const instructions = [
    ["1.", "Click on cell B4 and select a date from the date picker"],
    ["2.", "Go to Menu �� Time Machine �� Run QA for Selected Date"],
    ["3.", "Wait for processing to complete (may take a few minutes)"],
    ["4.", "Check results above and review Violations sheet"],
    ["5.", "Report will be automatically saved to Drive"]
  ];
  
  for (let i = 0; i < instructions.length; i++) {
    sheet.getRange(17 + i, 1).setValue(instructions[i][0]).setFontWeight("bold");
    sheet.getRange(17 + i, 2, 1, 2).merge();
    sheet.getRange(17 + i, 2).setValue(instructions[i][1]);
  }
  
  // Freeze header
  sheet.setFrozenRows(2);
  
  SpreadsheetApp.getUi().alert(
    '� Time Machine Ready',
    'Time Machine sheet has been set up!\n\n' +
    '1. Click on cell B4 to select a date\n' +
    '2. Use Menu �� Time Machine �� Run QA for Selected Date\n\n' +
    'The system will download that day\'s data from Gmail and run full QA analysis.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Setup Time Machine sheet automatically if it exists (called on onOpen)
 */
function setupTimeMachineIfExists_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Time Machine");
  
  if (sheet) {
    // Add button using drawing (if not already exists)
    // Note: Buttons need to be manually added via Insert > Drawing
    // This just ensures the sheet formatting is correct
  }
}

/**
 * Run QA for the date selected in Time Machine sheet
 */
function runTimeMachineQA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tmSheet = ss.getSheetByName("Time Machine");
  
  if (!tmSheet) {
    SpreadsheetApp.getUi().alert(
      '�� Time Machine Not Found',
      'Please run "Setup Time Machine Sheet" first from the menu.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // Get selected date
  const dateCell = tmSheet.getRange("B4");
  const dateValue = dateCell.getValue();
  
  if (!dateValue || !(dateValue instanceof Date)) {
    SpreadsheetApp.getUi().alert(
      '�� No Date Selected',
      'Please select a date in cell B4 first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const dateStr = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  // Update status
  tmSheet.getRange("B6").setValue("🔄 Processing " + dateStr + "...")
    .setFontColor("#ff6d00")
    .setFontWeight("bold");
  
  SpreadsheetApp.flush();
  
  // Confirm
  const confirm = SpreadsheetApp.getUi().alert(
    '🔄 Run QA for ' + dateStr,
    'This will:\n\n' +
    '1. Clear current Raw Data and Violations sheets\n' +
    '2. Download raw CSV files from Gmail for ' + dateStr + '\n' +
    '3. Run full QA analysis\n' +
    '4. Generate violations report\n' +
    '5. Save to Drive\n\n' +
    'This may take several minutes. Proceed?',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (confirm !== SpreadsheetApp.getUi().Button.YES) {
    tmSheet.getRange("B6").setValue("Ready - Select a date and run QA from the menu")
      .setFontColor("#666666")
      .setFontWeight("normal");
    return;
  }
  
  try {
    // Run the time machine process
    const result = runTimeMachineQA_(dateStr);
    
    if (result.success) {
      // Update status
      tmSheet.getRange("B6").setValue("� Complete - " + dateStr)
        .setFontColor("#0f9d58")
        .setFontWeight("bold");
      
      // Update last run
      tmSheet.getRange("B8").setValue(new Date().toLocaleString())
        .setFontColor("#000000");
      
      // Update results
      tmSheet.getRange("B11").setValue(result.filesProcessed);
      tmSheet.getRange("B12").setValue(result.placementsChecked);
      tmSheet.getRange("B13").setValue(result.violationCount);
      tmSheet.getRange("B14").setValue(result.fileUrl)
        .setFontColor("#1155cc")
        .setFontStyle("italic");
      
      SpreadsheetApp.getUi().alert(
        '� QA Complete for ' + dateStr,
        'Processing complete!\n\n' +
        'Files processed: ' + result.filesProcessed + '\n' +
        'Placements checked: ' + result.placementsChecked + '\n' +
        'Violations found: ' + result.violationCount + '\n\n' +
        'Report saved to: ' + result.folderPath + '\n' +
        'File: ' + result.filename + '\n\n' +
        'Check the Violations sheet for details.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      tmSheet.getRange("B6").setValue("�� Error - " + result.error)
        .setFontColor("#d93025")
        .setFontWeight("bold");
      
      SpreadsheetApp.getUi().alert('�� Error', result.error, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    tmSheet.getRange("B6").setValue("�� Error - " + error.toString())
      .setFontColor("#d93025")
      .setFontWeight("bold");
    
    SpreadsheetApp.getUi().alert('�� Error', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Internal function to run Time Machine QA for specific date
 */
function runTimeMachineQA_(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Raw Data");
  const violationsSheet = ss.getSheetByName("Violations");
  
  if (!rawSheet || !violationsSheet) {
    return { success: false, error: "Required sheets not found (Raw Data, Violations)" };
  }
  
  // Step 1: Clear existing data
  Logger.log(`Time Machine: Clearing sheets for ${dateStr}...`);
  clearRawData();
  clearViolations();
  
  // Step 2: Download raw CSVs from Gmail
  Logger.log(`Time Machine: Downloading from Gmail...`);
  const downloadResult = downloadRawDataForDate_(dateStr);
  
  if (!downloadResult.success) {
    return { success: false, error: downloadResult.error };
  }
  
  Logger.log(`Time Machine: Downloaded ${downloadResult.filesProcessed} files`);
  
  // Step 3: Run QA
  Logger.log(`Time Machine: Running QA analysis...`);
  try {
    processCSVData();
  } catch (error) {
    return { success: false, error: `QA processing failed: ${error.toString()}` };
  }
  
  const violationCount = violationsSheet.getLastRow() - 1;
  Logger.log(`Time Machine: Found ${violationCount} violations`);
  
  // Step 4: Save to Drive
  Logger.log(`Time Machine: Saving to Drive...`);
  const saveResult = saveViolationsReportToDrive_(dateStr, violationCount);
  
  if (!saveResult.success) {
    return { success: false, error: saveResult.error };
  }
  
  return {
    success: true,
    filename: saveResult.filename,
    folderPath: saveResult.folderPath,
    fileUrl: saveResult.fileUrl,
    filesProcessed: downloadResult.filesProcessed,
    placementsChecked: rawSheet.getLastRow() - 1,
    violationCount: violationCount
  };
}

/**
 * Download raw CSV files from Gmail for a specific date
 */
/**
 * GAP FILL: Import DCM Reports for a specific date (follows runItAll pattern)
 * This is the Time Machine version - filters by missing date instead of today
 * DOES NOT modify the original importDCMReports() function
 */
function downloadRawDataForDate_(dateStr) {
  // Try Drive first (much faster!)
  Logger.log(`🔍 Attempting to load raw data from Drive for ${dateStr}...`);
  const driveResult = downloadRawDataFromDrive_(dateStr);
  
  if (driveResult.success) {
    Logger.log(`✅ Drive: Loaded ${driveResult.filesProcessed} CSV files from Drive`);
    return driveResult;
  }
  
  // Fallback to Gmail if Drive folder doesn't exist
  Logger.log(`⚠️ Drive folder not found. Falling back to Gmail...`);
  return importDCMReportsForDate_(dateStr);
}

/**
 * GAP FILL: Gmail import for specific date (pattern from importDCMReports)
 * Phase 1: Data Preparation - Import raw CM360 data from Gmail
 * Searches label "CM360 QA" filtered to specific missing date
 */
function importDCMReportsForDate_(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName("Raw Data");
  
  // Auto-create Raw Data sheet if it doesn't exist
  if (!dataSheet) {
    Logger.log('⚠️ Raw Data sheet not found - creating it now...');
    dataSheet = ss.insertSheet("Raw Data");
    const dataHeaders = [
      "Network ID","Advertiser","Placement ID","Placement","Campaign",
      "Placement Start Date","Placement End Date","Campaign Start Date","Campaign End Date",
      "Ad","Impressions","Clicks","Report Date"
    ];
    dataSheet.getRange(1,1,1,dataHeaders.length).setValues([dataHeaders]).setFontWeight("bold");
    Logger.log('✅ Raw Data sheet created');
  }
  
  // Format date for Gmail search: YYYY/MM/DD
  const date = new Date(dateStr);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const formattedDate = `${year}/${month}/${day}`;
  
  // Get next day for before: parameter
  const nextDay = new Date(date);
  nextDay.setDate(nextDay.getDate() + 1);
  const nextYear = nextDay.getFullYear();
  const nextMonth = String(nextDay.getMonth() + 1).padStart(2, '0');
  const nextDayNum = String(nextDay.getDate()).padStart(2, '0');
  const formattedNextDate = `${nextYear}/${nextMonth}/${nextDayNum}`;
  
  // Search Gmail with CM360 QA label (same as importDCMReports)
  const label = "CM360 QA";
  const searchQuery = `label:${label} after:${formattedDate} before:${formattedNextDate}`;
  
  Logger.log(`📧 Searching Gmail: ${searchQuery}`);
  const threads = GmailApp.search(searchQuery);
  
  if (threads.length === 0) {
    return {
      success: false,
      error: `No emails found for ${formattedDate} with label "${label}". Check Gmail labels.`
    };
  }
  
  Logger.log(`Found ${threads.length} email thread(s) with label "${label}"`);
  
  let extractedData = [];
  let filesProcessed = 0;
  
  // Process each thread (same as importDCMReports)
  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(message) {
      message.getAttachments().forEach(function(att) {
        const filename = att.getName();
        const netId = extractNetworkId(filename);
        
        try {
          if (att.getContentType() === "text/csv" || filename.endsWith(".csv")) {
            // Process CSV directly
            const csvData = processCSV(att.getDataAsString(), netId);
            extractedData = extractedData.concat(csvData);
            filesProcessed++;
            Logger.log(`  ✅ ${filename} (${csvData.length} rows, Network ${netId})`);
            
          } else if (att.getContentType() === "application/zip" || filename.endsWith(".zip")) {
            // Unzip and process CSV files inside
            Utilities.unzip(att.copyBlob()).forEach(function(file) {
              const unzippedName = file.getName();
              if (file.getContentType() === "text/csv" || unzippedName.endsWith(".csv")) {
                const unzippedNetId = extractNetworkId(unzippedName);
                const csvData = processCSV(file.getDataAsString(), unzippedNetId);
                extractedData = extractedData.concat(csvData);
                filesProcessed++;
                Logger.log(`  ✅ (ZIP) ${unzippedName} (${csvData.length} rows, Network ${unzippedNetId})`);
              }
            });
          }
        } catch (error) {
          Logger.log(`  ❌ Error processing ${filename}: ${error}`);
        }
      });
    });
  });
  
  if (extractedData.length === 0) {
    return {
      success: false,
      error: `Found ${threads.length} email(s) but no CSV/ZIP data could be processed.`
    };
  }
  
  // Write to Raw Data sheet (starting at row 2, after headers)
  const dataHeaders = [
    "Network ID","Advertiser","Placement ID","Placement","Campaign",
    "Placement Start Date","Placement End Date","Campaign Start Date","Campaign End Date",
    "Ad","Impressions","Clicks","Report Date"
  ];
  
  if (extractedData.length > 0) {
    dataSheet.getRange(2, 1, extractedData.length, dataHeaders.length).setValues(extractedData);
    Logger.log(`✅ Imported ${extractedData.length} total rows from ${filesProcessed} files`);
  }
  
  return {
    success: true,
    filesProcessed: filesProcessed,
    rowsImported: extractedData.length,
    source: 'Gmail (CM360 QA label)'
  };
}

/**
 * Download raw data from Google Drive (FAST!)
 */
function downloadRawDataFromDrive_(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let rawSheet = ss.getSheetByName("Raw Data");
  
  // Auto-create Raw Data sheet if it doesn't exist
  if (!rawSheet) {
    Logger.log('⚠️ Raw Data sheet not found - creating it now...');
    rawSheet = ss.insertSheet("Raw Data");
    rawSheet.getRange("A1:H1").setValues([[
      "Network ID", "Advertiser", "Campaign", "Placement", 
      "Start Date", "End Date", "Cost Structure", "Report Date"
    ]]).setFontWeight("bold");
    Logger.log('✅ Raw Data sheet created');
  }
  try {
    // Build folder path: Root/2025/04-April/2025-04-15/
    const RAW_DATA_ROOT_ID = '1qA77_YET8RLiES7X7NoUT5jzTHDJ3k61';
    const date = new Date(dateStr);
    const year = date.getFullYear();
    const month = date.getMonth() + 1; // 1-12
    const day = date.getDate();
    
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 
                        'July', 'August', 'September', 'October', 'November', 'December'];
    const monthFolder = `${String(month).padStart(2, '0')}-${monthNames[month - 1]}`;
    const dateFolder = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    
    Logger.log(`📁 Looking for: ${year}/${monthFolder}/${dateFolder}/`);
    
    // Navigate folder structure
    const rootFolder = DriveApp.getFolderById(RAW_DATA_ROOT_ID);
    
    // Find year folder
    const yearFolders = rootFolder.getFoldersByName(String(year));
    if (!yearFolders.hasNext()) {
      return { success: false, error: `Year folder ${year} not found in Drive` };
    }
    const yearFolderObj = yearFolders.next();
    
    // Find month folder
    const monthFolders = yearFolderObj.getFoldersByName(monthFolder);
    if (!monthFolders.hasNext()) {
      return { success: false, error: `Month folder ${monthFolder} not found in Drive` };
    }
    const monthFolderObj = monthFolders.next();
    
    // Find date folder
    const dateFolders = monthFolderObj.getFoldersByName(dateFolder);
    if (!dateFolders.hasNext()) {
      return { success: false, error: `Date folder ${dateFolder} not found in Drive` };
    }
    const dateFolderObj = dateFolders.next();
    
    // Get all CSV files
    const csvFiles = dateFolderObj.getFilesByName('');
    const allFiles = [];
    while (csvFiles.hasNext()) {
      allFiles.push(csvFiles.next());
    }
    
    // Filter CSV files only
    const csvOnly = allFiles.filter(f => f.getName().toLowerCase().endsWith('.csv'));
    
    if (csvOnly.length === 0) {
      return { success: false, error: `No CSV files found in ${dateFolder}` };
    }
    
    Logger.log(`📂 Found ${csvOnly.length} CSV files in Drive`);
    
    // Process each CSV
    let filesProcessed = 0;
    let currentRow = 2;
    
    for (const file of csvOnly) {
      try {
        const filename = file.getName();
        const content = file.getBlob().getDataAsString();
        
        // Extract NetworkID: first digits before first underscore
        // Example: "1068_BKCM360_Global_QA_Check_20250801_005620_5225494517.csv" → "1068"
        const networkId = filename.split('_')[0];
        
        const rows = processCSV(content, networkId);
        
        if (rows.length > 0) {
          rawSheet.getRange(currentRow, 1, rows.length, rows[0].length).setValues(rows);
          currentRow += rows.length;
          filesProcessed++;
          Logger.log(`  ✅ ${filename} → ${rows.length} rows (Network ${networkId})`);
        }
      } catch (error) {
        Logger.log(`  ❌ Error processing ${file.getName()}: ${error}`);
      }
    }
    
    return {
      success: true,
      filesProcessed: filesProcessed,
      source: 'Drive'
    };
    
  } catch (error) {
    return { success: false, error: `Drive access error: ${error}` };
  }
}

/**
 * Download raw data from Gmail (FALLBACK - slower)
 */

/**
 * Save violations report to Drive
 */
function saveViolationsReportToDrive_(dateStr, violationCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const violationsSheet = ss.getSheetByName("Violations");
  
  if (!violationsSheet || violationsSheet.getLastRow() < 2) {
    return {
      success: false,
      error: "No violations data to save (sheet is empty or has only headers)"
    };
  }
  
  const rootFolderId = '1F53lLe3z5cup338IRY4nhTZQdUmJ9_wk';
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  
  let violationsReportsFolder;
  const vFolders = rootFolder.getFoldersByName('Violations Reports');
  if (vFolders.hasNext()) {
    violationsReportsFolder = vFolders.next();
  } else {
    violationsReportsFolder = rootFolder.createFolder('Violations Reports');
  }
  
  const yearMonth = dateStr.substring(0, 7);
  let monthFolder;
  const mFolders = violationsReportsFolder.getFoldersByName(yearMonth);
  if (mFolders.hasNext()) {
    monthFolder = mFolders.next();
  } else {
    monthFolder = violationsReportsFolder.createFolder(yearMonth);
  }
  
  const filename = `Violations_${dateStr}.xlsx`;
  const xlsxBlob = createXLSXFromSheet(violationsSheet);
  xlsxBlob.setName(filename);
  
  const existingFiles = monthFolder.getFilesByName(filename);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }
  
  const file = monthFolder.createFile(xlsxBlob);
  const fileUrl = file.getUrl();
  
  Logger.log(`� Saved violations report: ${filename} (${violationCount} violations)`);
  
  return {
    success: true,
    filename: filename,
    folderPath: `Violations Reports/${yearMonth}`,
    fileUrl: fileUrl
  };
}

/**
 * Helper: Get next date
 */
function getNextDate_(dateStr) {
  const date = new Date(dateStr);
  date.setDate(date.getDate() + 1);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// =====================================================================================================================
// ======================================= END TIME MACHINE SYSTEM ====================================================
// =====================================================================================================================
// =====================================================================================================================
// ======================================= AUDIT SYSTEMS (see AuditSystems.gs) ========================================
// =====================================================================================================================
// All audit dashboard functionality has been moved to AuditSystems.gs for better organization
// - Raw Data Audit (Drive scanning)
// - Violations Audit (Historical reports)
// Functions: setupAndRefreshRawDataAudit(), setupAndRefreshViolationsAudit(), resetRawDataAudit()
// =====================================================================================================================

// =====================================================================================================================
// ========================================== DRIVE FOLDER CRAWLER ====================================================
// =====================================================================================================================

/**
 * Crawl a Drive folder and log its structure
 */
function crawlDriveFolder() {
  const folderId = '1uOXQ-zgCZ5-d9E2ewR-XO11c1sperj5S';
  const folder = DriveApp.getFolderById(folderId);
  
  Logger.log('=== DRIVE FOLDER STRUCTURE ===');
  Logger.log('Root: ' + folder.getName());
  Logger.log('');
  
  crawlFolder_(folder, 0);
  
  SpreadsheetApp.getUi().alert(
    'Folder Crawl Complete',
    'Check the execution log (View > Logs) to see the folder structure.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function crawlFolder_(folder, depth) {
  const indent = '  '.repeat(depth);
  
  // List subfolders
  const subfolders = folder.getFolders();
  const folderList = [];
  while (subfolders.hasNext()) {
    folderList.push(subfolders.next());
  }
  
  // List files
  const files = folder.getFiles();
  const fileList = [];
  while (files.hasNext()) {
    fileList.push(files.next());
  }
  
  Logger.log(indent + '📁 ' + folder.getName() + ' (' + folderList.length + ' folders, ' + fileList.length + ' files)');
  
  // Show sample files (first 3)
  for (let i = 0; i < Math.min(3, fileList.length); i++) {
    Logger.log(indent + '  📄 ' + fileList[i].getName());
  }
  if (fileList.length > 3) {
    Logger.log(indent + '  ... and ' + (fileList.length - 3) + ' more files');
  }
  
  // Recurse into subfolders (max depth 3)
  if (depth < 3) {
    for (const subfolder of folderList) {
      crawlFolder_(subfolder, depth + 1);
    }
  } else if (folderList.length > 0) {
    Logger.log(indent + '  [' + folderList.length + ' subfolders not shown - max depth reached]');
  }
}

// =====================================================================================================================
// ======================================== END DRIVE FOLDER CRAWLER ==================================================
// =====================================================================================================================
// NOTE: Audit system functions (setupAndRefreshRawDataAudit, setupAndRefreshViolationsAudit, etc.)
// and related constants (GAP_FILL_STATE_KEY, VIOLATIONS_ROOT_FOLDER_ID, etc.)
// are now located in AuditSystems.gs file for better code organization and to avoid duplicate definitions.
// =====================================================================================================================

/**
 * Setup Gap Fill Progress Sheet
 */
function setupGapFillProgressSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Gap Fill Progress");
  
  if (!sheet) {
    sheet = ss.insertSheet("Gap Fill Progress");
  }
  
  sheet.clear();
  
  // Column widths
  sheet.setColumnWidth(1, 120);  // Date
  sheet.setColumnWidth(2, 150);  // Status
  sheet.setColumnWidth(3, 180);  // Last Updated
  sheet.setColumnWidth(4, 80);   // Attempts
  sheet.setColumnWidth(5, 300);  // Error Message
  sheet.setColumnWidth(6, 150);  // Drive File
  
  // Headers
  const headers = [["Date", "Status", "Last Updated", "Attempts", "Error Message", "Drive File"]];
  sheet.getRange(1, 1, 1, 6).setValues(headers)
    .setFontWeight("bold")
    .setBackground("#34a853")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  sheet.setFrozenRows(1);
  
  Logger.log('Gap Fill Progress sheet created successfully');
  // No notification - called internally during automation
}

/**
 * Get missing dates from Violations Audit sheet
 */
function getMissingDatesFromAudit_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName("Violations Audit");
  
  if (!auditSheet || auditSheet.getLastRow() < 2) {
    Logger.log('No Violations Audit sheet found or sheet is empty');
    return [];
  }
  
  const data = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 2).getValues();
  const missingDates = [];
  const startDate = new Date('2025-04-14');
  
  Logger.log(`Scanning ${data.length} rows in Violations Audit`);
  
  for (const row of data) {
    const dateStr = String(row[0]);
    const status = String(row[1]).trim();
    
    // Check for MISSING status (handle different formats)
    if (status.includes('MISSING') && dateStr) {
      const checkDate = new Date(dateStr);
      // Skip dates before 4.14.25 - no data exists
      if (checkDate >= startDate) {
        missingDates.push(dateStr);
      }
    }
  }
  
  Logger.log(`Found ${missingDates.length} missing dates`);
  return missingDates;
}

/**
 * Initialize Gap Fill Progress sheet with missing dates
 */
function initializeGapFillProgress_(missingDates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Gap Fill Progress");
  
  if (!sheet) {
    setupGapFillProgressSheet();
    sheet = ss.getSheetByName("Gap Fill Progress");
  }
  
  // Clear existing data
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clear();
  }
  
  // Add all missing dates as "Queued"
  const rows = missingDates.map(date => [
    date,
    '�� Queued',
    new Date().toLocaleString(),
    0,
    '',
    ''
  ]);
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  
  return rows.length;
}

/**
 * Update progress for a specific date
 */
function updateGapFillProgress_(dateStr, status, errorMsg, driveFile) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Gap Fill Progress");
  
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === dateStr) {
      const attempts = Number(data[i][3] || 0) + 1;
      sheet.getRange(i + 1, 2).setValue(status);
      sheet.getRange(i + 1, 3).setValue(new Date().toLocaleString());
      sheet.getRange(i + 1, 4).setValue(attempts);
      sheet.getRange(i + 1, 5).setValue(errorMsg || '');
      sheet.getRange(i + 1, 6).setValue(driveFile || '');
      
      // Color code status
      const statusCell = sheet.getRange(i + 1, 2);
      if (status.includes('�')) {
        statusCell.setBackground('#d4edda').setFontColor('#155724');
      } else if (status.includes('��')) {
        statusCell.setBackground('#f8d7da').setFontColor('#721c24');
      } else if (status.includes('🔄')) {
        statusCell.setBackground('#cfe2ff').setFontColor('#084298');
      } else if (status.includes('��')) {
        statusCell.setBackground('#fff3cd').setFontColor('#856404');
      }
      
      break;
    }
  }
}

/**
 * Get gap fill state from DocumentProperties
 */
function getGapFillState_() {
  try {
    const props = PropertiesService.getDocumentProperties();
    const stateJson = props.getProperty(GAP_FILL_STATE_KEY);
    return stateJson ? JSON.parse(stateJson) : null;
  } catch (e) {
    Logger.log('Error loading gap fill state: ' + e);
    return null;
  }
}

/**
 * Save gap fill state to DocumentProperties
 */
function saveGapFillState_(state) {
  try {
    const props = PropertiesService.getDocumentProperties();
    props.setProperty(GAP_FILL_STATE_KEY, JSON.stringify(state));
  } catch (e) {
    Logger.log('Error saving gap fill state: ' + e);
  }
}

/**
 * Clear gap fill state
 */
function clearGapFillState_() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty(GAP_FILL_STATE_KEY);
}

/**
 * Format date to MM.DD.YY format for email attachment search
 */
function formatDateForEmail_(dateStr) {
  // dateStr is "2025-04-23"
  const parts = dateStr.split('-');
  const month = parts[1];
  const day = parts[2];
  const year = parts[0].slice(2); // "25"
  return `${month}.${day}.${year}`;
}

/**
 * Format date to month folder name "MM-MonthName"
 */
function getMonthFolderName_(dateStr) {
  const parts = dateStr.split('-');
  const monthNum = parseInt(parts[1]);
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                      'July', 'August', 'September', 'October', 'November', 'December'];
  return `${parts[1]}-${monthNames[monthNum - 1]}`;
}

/**
 * Search Gmail for violations email attachment
 * Returns {found: boolean, attachment: Blob, filename: string}
 */
function searchGmailForViolationsAttachment_(dateStr) {
  const emailDateFormat = formatDateForEmail_(dateStr); // "04.23.25"
  
  // Search for violations email on that date
  const query = `subject:"CM360 CPC/CPM FLIGHT QA" after:${dateStr} before:${getNextDate_(dateStr)} has:attachment`;
  const threads = GmailApp.search(query, 0, 5);
  
  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const msg of messages) {
      const attachments = msg.getAttachments();
      for (const attachment of attachments) {
        const filename = attachment.getName();
        
        // Check if filename contains the date in either format
        if (filename.includes(emailDateFormat) || filename.includes(dateStr)) {
          return {
            found: true,
            attachment: attachment,
            filename: filename
          };
        }
      }
    }
  }
  
  return { found: false, attachment: null, filename: null };
}

/**
 * Save violations attachment to Drive with uniform naming
 */
function saveViolationsAttachmentToDrive_(dateStr, attachment, originalFilename) {
  const rootFolder = DriveApp.getFolderById(VIOLATIONS_ROOT_FOLDER_ID);
  
  // Get year folder (2025)
  const yearFolders = rootFolder.getFoldersByName('2025');
  let yearFolder;
  if (yearFolders.hasNext()) {
    yearFolder = yearFolders.next();
  } else {
    yearFolder = rootFolder.createFolder('2025');
  }
  
  // Get month folder
  const monthFolderName = getMonthFolderName_(dateStr);
  const monthFolders = yearFolder.getFoldersByName(monthFolderName);
  let monthFolder;
  if (monthFolders.hasNext()) {
    monthFolder = monthFolders.next();
  } else {
    monthFolder = yearFolder.createFolder(monthFolderName);
  }
  
  // Uniform filename
  const uniformFilename = `CM360_Report_${dateStr}.xlsx`;
  
  // Check if file already exists and delete it
  const existingFiles = monthFolder.getFilesByName(uniformFilename);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }
  
  // Save attachment
  const file = monthFolder.createFile(attachment);
  file.setName(uniformFilename);
  
  return {
    success: true,
    filename: uniformFilename,
    url: file.getUrl(),
    folderPath: `Historical Violation Reports/2025/${monthFolderName}/`
  };
}

/**
 * Start Auto Gap Fill process
 */
function startAutoGapFill() {
  const ui = SpreadsheetApp.getUi();
  
  // First, run Violations Audit to get latest missing dates
  ui.alert('🔄 Running Violations Audit', 'Scanning Drive for missing violations reports...', ui.ButtonSet.OK);
  setupAndRefreshViolationsAudit();
  
  // Get missing dates
  const missingDates = getMissingDatesFromAudit_();
  
  if (missingDates.length === 0) {
    ui.alert(
      '� No Gaps Found',
      'All violations reports are present in Drive!\n\nNo gap-fill needed.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Initialize progress sheet
  const count = initializeGapFillProgress_(missingDates);
  
  // Initialize state
  const state = {
    queue: missingDates,
    currentDate: null,
    currentStep: null,
    startTime: new Date().toISOString(),
    processed: 0,
    successful: 0,
    failed: 0
  };
  saveGapFillState_(state);
  
  ui.alert(
    '� Gap Fill Started',
    `Found ${count} missing violations reports.\n\n` +
    `Auto gap-fill will process them automatically.\n\n` +
    `Create an auto-resume trigger (10 min) from the menu to enable continuous processing.`,
    ui.ButtonSet.OK
  );
  
  // Start first chunk
  processGapFillChunk_();
}

/**
 * Process one chunk of gap fill (respects time budget)
 */
function processGapFillChunk_() {
  const startTime = Date.now();
  const state = getGapFillState_();
  
  if (!state || !state.queue || state.queue.length === 0) {
    Logger.log('� Gap fill complete or no state found');
    return;
  }
  
  // Process dates from queue
  while (state.queue.length > 0 && (Date.now() - startTime) < GAP_FILL_TIME_BUDGET_MS) {
    const dateStr = state.queue[0];
    state.currentDate = dateStr;
    
    Logger.log(`🔄 Processing date: ${dateStr}`);
    updateGapFillProgress_(dateStr, '🔄 Checking Email...', '', '');
    
    try {
      // Step 1: Check Gmail for existing violations email
      const emailResult = searchGmailForViolationsAttachment_(dateStr);
      
      if (emailResult.found) {
        Logger.log(`� Found email attachment for ${dateStr}`);
        updateGapFillProgress_(dateStr, '🔄 Saving to Drive...', '', '');
        
        // Save to Drive
        const saveResult = saveViolationsAttachmentToDrive_(dateStr, emailResult.attachment, emailResult.filename);
        
        updateGapFillProgress_(dateStr, '� Complete (from email)', '', saveResult.filename);
        state.successful++;
        state.processed++;
        state.queue.shift(); // Remove from queue
        saveGapFillState_(state);
        continue;
      }
      
      // Step 2: Email not found, need to run Time Machine
      Logger.log(`��️ No email found for ${dateStr}, running Time Machine`);
      updateGapFillProgress_(dateStr, '🔄 Running Time Machine...', '', '');
      
      const tmResult = runTimeMachineForDate_(dateStr);
      
      if (tmResult.success) {
        updateGapFillProgress_(dateStr, '� Complete (regenerated)', '', tmResult.filename);
        state.successful++;
      } else {
        updateGapFillProgress_(dateStr, '�� Failed', tmResult.error, '');
        state.failed++;
      }
      
      state.processed++;
      state.queue.shift();
      saveGapFillState_(state);
      
    } catch (e) {
      Logger.log(`�� Error processing ${dateStr}: ${e}`);
      updateGapFillProgress_(dateStr, '�� Failed', String(e), '');
      state.failed++;
      state.processed++;
      state.queue.shift();
      saveGapFillState_(state);
    }
  }
  
  // Save final state
  saveGapFillState_(state);
  
  if (state.queue.length === 0) {
    Logger.log(`� Gap fill complete! Processed: ${state.processed}, Successful: ${state.successful}, Failed: ${state.failed}`);
    clearGapFillState_();
  } else {
    Logger.log(`��️ Gap fill paused. Remaining: ${state.queue.length}/${state.processed + state.queue.length}`);
  }
}

/**
 * Run Time Machine for a specific date (internal version)
 * Returns {success, filename, error}
 */
function runTimeMachineForDate_(dateStr) {
  try {
    // Check if date is before data collection started (4.14.25)
    const startDate = new Date('2025-04-14');
    const checkDate = new Date(dateStr);
    
    if (checkDate < startDate) {
      return { 
        success: false, 
        error: 'No data available before 4.14.25 (data collection start date)' 
      };
    }
    
    // Clear sheets
    clearRawData();
    clearViolations();
    
    // Step 1: Download raw data
    const downloadResult = downloadRawDataForDate_(dateStr);
    if (!downloadResult.success) {
      return { success: false, error: downloadResult.error };
    }
    
    // Step 2: Run QA
    runQAOnly();
    
    // Step 3: Save violations report
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const violationsSheet = ss.getSheetByName("Violations");
    const violationCount = violationsSheet ? Math.max(0, violationsSheet.getLastRow() - 1) : 0;
    
    const saveResult = saveViolationsReportToDrive_(dateStr, violationCount);
    
    if (!saveResult.success) {
      return { success: false, error: saveResult.error };
    }
    
    // Step 4: Send email summary
    sendEmailSummary();
    
    return {
      success: true,
      filename: saveResult.filename,
      fileUrl: saveResult.fileUrl
    };
    
  } catch (e) {
    return { success: false, error: String(e) };
  }
}

/**
 * View Gap Fill Status
 */
function viewGapFillStatus() {
  const state = getGapFillState_();
  const ui = SpreadsheetApp.getUi();
  
  if (!state) {
    ui.alert(
      '📊 Gap Fill Status',
      'No gap fill process is currently running.\n\n' +
      'Run "Start Auto Gap Fill" to begin.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  const remaining = state.queue ? state.queue.length : 0;
  const total = state.processed + remaining;
  
  ui.alert(
    '📊 Gap Fill Status',
    `Started: ${new Date(state.startTime).toLocaleString()}\n\n` +
    `Total Dates: ${total}\n` +
    `Processed: ${state.processed}\n` +
    `Successful: ${state.successful}\n` +
    `Failed: ${state.failed}\n` +
    `Remaining: ${remaining}\n\n` +
    `Current: ${state.currentDate || 'None'}\n` +
    `Step: ${state.currentStep || 'N/A'}\n\n` +
    `Check "Gap Fill Progress" sheet for details.`,
    ui.ButtonSet.OK
  );
}

/**
 * Reset Gap Fill
 */
function resetGapFill() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '��️ Reset Gap Fill',
    'This will clear all progress and start over.\n\nAre you sure?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    clearGapFillState_();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Gap Fill Progress");
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    
    ui.alert('� Reset Complete', 'Gap fill has been reset.', ui.ButtonSet.OK);
  }
}

/**
 * Create Auto-Resume Trigger for Gap Fill
 */
function createGapFillAutoResumeTrigger() {
  // Delete existing trigger
  deleteGapFillAutoResumeTrigger_();
  
  // Create new 10-minute recurring trigger
  const trigger = ScriptApp.newTrigger('processGapFillChunk_')
    .timeBased()
    .everyMinutes(10)
    .create();
  
  // Save trigger ID
  const props = PropertiesService.getScriptProperties();
  props.setProperty(GAP_FILL_TRIGGER_KEY, trigger.getUniqueId());
  
  SpreadsheetApp.getUi().alert(
    '� Auto-Resume Trigger Created',
    'Gap fill will automatically resume every 10 minutes.\n\n' +
    'The trigger will process missing violations reports continuously until all gaps are filled.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Delete Gap Fill Auto-Resume Trigger
 */
function deleteGapFillAutoResumeTrigger_() {
  const props = PropertiesService.getScriptProperties();
  const triggerId = props.getProperty(GAP_FILL_TRIGGER_KEY);
  
  if (triggerId) {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
        break;
      }
    }
    props.deleteProperty(GAP_FILL_TRIGGER_KEY);
  }
}

/**
 * Stop Gap Fill and Delete Trigger
 */
function stopGapFillAndDeleteTrigger() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '🛑 Stop Gap Fill',
    'This will stop the auto gap-fill process and delete the trigger.\n\n' +
    'Progress will be saved and you can resume later.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    deleteGapFillAutoResumeTrigger_();
    ui.alert(
      '� Stopped',
      'Gap fill process stopped and trigger deleted.\n\n' +
      'Progress has been saved. Run "Start Auto Gap Fill" to resume.',
      ui.ButtonSet.OK
    );
  }
}

// =====================================================================================================================
// ======================================= END AUTO GAP FILL SYSTEM (VIOLATIONS) ======================================
// =====================================================================================================================


// =====================================================================================================================
// ================================= RAW DATA GAP FILL SYSTEM ==========================================================
// =====================================================================================================================

// NOTE: Constants for Raw Data Gap Fill are now in AuditSystems.gs to avoid duplicate declarations

/**
 * Analyze network lifecycle from Audit Dashboard
 * Tracks BOTH found and missing networks to understand when each network is expected
 * Returns map of { networkId: { firstExpected: 'YYYY-MM-DD', lastExpected: 'YYYY-MM-DD', allDates: Set } }
 */
function analyzeNetworkLifecycle_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Audit Dashboard");
  
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  const networkLifecycle = {};
  
  // Get all networks from Networks sheet
  const networksSheet = ss.getSheetByName("Networks");
  if (!networksSheet) return {};
  
  const networkData = networksSheet.getDataRange().getValues();
  const allPossibleNetworks = [];
  for (let j = 1; j < networkData.length; j++) {
    const netId = String(networkData[j][0] || '').trim();
    if (netId) allPossibleNetworks.push(netId);
  }
  
  // Find where data starts (dates are stored as Date objects)
  let startRow = 0;
  for (let i = 0; i < data.length; i++) {
    const cellValue = data[i][0];
    if (cellValue instanceof Date && !isNaN(cellValue)) {
      startRow = i;
      Logger.log(`Found data starting at row ${i}`);
      break;
    }
  }
  
  if (startRow === 0) {
    Logger.log(`Could not find date data. Total rows: ${data.length}`);
    return {};
  }
  
  // Scan all dates to find when each network is EXPECTED (found OR missing means expected)
  for (let i = startRow; i < data.length; i++) {
    const dateObj = data[i][0];
    if (!(dateObj instanceof Date) || isNaN(dateObj)) continue;
    
    const dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const status = String(data[i][1] || '').trim();
    const filesInDrive = Number(data[i][2]) || 0;
    const missingNetworks = String(data[i][4] || '').trim();
    
    // Skip dates with no activity at all (0 files, no status)
    if (filesInDrive === 0 && status !== '�� MISSING' && status !== '��️ PARTIAL') continue;
    
    // Get missing networks list
    const missingList = missingNetworks && missingNetworks !== '��' 
      ? missingNetworks.split(',').map(n => n.trim()).filter(n => n && n !== '��')
      : [];
    
    // Get present networks (all possible minus missing)
    const presentNetworks = allPossibleNetworks.filter(n => !missingList.includes(n));
    
    // ANY network that appears (found OR missing) is EXPECTED on this date
    const expectedNetworks = [...new Set([...presentNetworks, ...missingList])];
    
    // Update lifecycle for all expected networks
    for (const netId of expectedNetworks) {
      if (!networkLifecycle[netId]) {
        networkLifecycle[netId] = {
          firstExpected: dateStr,
          lastExpected: dateStr,
          allDates: new Set([dateStr]),
          foundCount: 0,
          missingCount: 0
        };
      } else {
        if (dateStr < networkLifecycle[netId].firstExpected) {
          networkLifecycle[netId].firstExpected = dateStr;
        }
        if (dateStr > networkLifecycle[netId].lastExpected) {
          networkLifecycle[netId].lastExpected = dateStr;
        }
        networkLifecycle[netId].allDates.add(dateStr);
      }
      
      // Track if found or missing
      if (missingList.includes(netId)) {
        networkLifecycle[netId].missingCount++;
      } else {
        networkLifecycle[netId].foundCount++;
      }
    }
  }
  
  Logger.log(`Analyzed lifecycle for ${Object.keys(networkLifecycle).length} networks`);
  
  // Log some examples
  for (const netId of Object.keys(networkLifecycle).slice(0, 3)) {
    const info = networkLifecycle[netId];
    Logger.log(`Network ${netId}: ${info.firstExpected} to ${info.lastExpected}, Found: ${info.foundCount}, Missing: ${info.missingCount}`);
  }
  
  return networkLifecycle;
}

/**
 * Check if a network gap should be filled based on lifecycle
 * Returns { shouldFill: boolean, reason: string }
 */
function shouldFillNetworkGap_(networkId, dateStr, lifecycle) {
  const networkInfo = lifecycle[networkId];
  
  if (!networkInfo) {
    // Network never seen anywhere - might be newly added, fill it
    return { shouldFill: true, reason: 'Unknown network - attempting fill' };
  }
  
  const dateObj = new Date(dateStr);
  const firstExpectedObj = new Date(networkInfo.firstExpected);
  const lastExpectedObj = new Date(networkInfo.lastExpected);
  
  // Check if date is before network started
  if (dateObj < firstExpectedObj) {
    return { shouldFill: false, reason: `Before first expected (${networkInfo.firstExpected})` };
  }
  
  // Check if date is after network ended (7-day threshold)
  const daysSinceLastExpected = Math.floor((dateObj - lastExpectedObj) / 86400000);
  if (daysSinceLastExpected > 7) {
    return { shouldFill: false, reason: `${daysSinceLastExpected} days after last expected (${networkInfo.lastExpected})` };
  }
  
  // Date is within expected period - fill it
  return { shouldFill: true, reason: 'Within active period' };
}

/**
 * Get missing items from Audit Dashboard with smart lifecycle filtering
 * Returns array of { date, networks: [] }
 */
function getMissingRawDataFromAudit_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Audit Dashboard");
  
  if (!sheet) {
    throw new Error("Audit Dashboard sheet not found. Run the audit first.");
  }
  
  // Analyze network lifecycle first
  Logger.log("Analyzing network lifecycle...");
  const lifecycle = analyzeNetworkLifecycle_();
  
  const data = sheet.getDataRange().getValues();
  const missing = [];
  const skipped = { beforeStart: 0, afterEnd: 0, total: 0 };
  const byNetwork = {}; // Track gaps per network
  
  // Find where data starts (skip header row)
  let startRow = 1; // Start at row 1 (0-based), which is row 2 in sheet
  
  // Verify we have actual date data
  if (data.length < 2) {
    Logger.log(`Not enough data rows. Total rows: ${data.length}`);
    return [];
  }
  
  // Check if row 1 (index 1) has a date
  const firstDataRow = data[1];
  if (!(firstDataRow[0] instanceof Date) || isNaN(firstDataRow[0])) {
    Logger.log(`Row 2 (index 1) doesn't contain a valid date. Value: ${firstDataRow[0]}`);
    // Try to find first date row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] instanceof Date && !isNaN(data[i][0])) {
        startRow = i;
        Logger.log(`Found data starting at row ${i + 1} (0-based index ${i})`);
        break;
      }
    }
  } else {
    Logger.log(`Data starts at row 2 (0-based index 1)`);
  }

  // Process data rows
  for (let i = startRow; i < data.length; i++) {
    const dateObj = data[i][0];
    if (!(dateObj instanceof Date) || isNaN(dateObj)) continue;
    
    const dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const status = String(data[i][1] || '').trim();
    const missingNetworks = String(data[i][4] || '').trim();
    
    Logger.log(`Checking ${dateStr}: Status="${status}", Missing="${missingNetworks}"`);
    
    // Only process MISSING or PARTIAL statuses (handle emoji variations)
    if (status.includes('MISSING') || status.includes('PARTIAL')) {
      let networksList = [];
      
      if (status.includes('MISSING')) {
        // Get all networks for this date
        const networksSheet = ss.getSheetByName("Networks");
        if (networksSheet) {
          const networkData = networksSheet.getDataRange().getValues();
          for (let j = 1; j < networkData.length; j++) {
            const netId = String(networkData[j][0] || '').trim();
            if (netId) networksList.push(netId);
          }
        }
      } else if (status.includes('PARTIAL') && missingNetworks) {
        // Parse missing networks from column E (index 4)
        // Handle comma-separated or space-separated values
        networksList = missingNetworks
          .split(/[,\s]+/)
          .map(n => n.trim())
          .filter(n => n && n !== '✓' && !n.includes('✓'));
      }
      
      // Filter networks based on lifecycle
      const validNetworks = [];
      for (const netId of networksList) {
        const check = shouldFillNetworkGap_(netId, dateStr, lifecycle);
        
        if (check.shouldFill) {
          validNetworks.push(netId);
          byNetwork[netId] = (byNetwork[netId] || 0) + 1;
        } else {
          skipped.total++;
          if (check.reason.includes('Before first expected')) {
            skipped.beforeStart++;
          } else if (check.reason.includes('after last expected')) {
            skipped.afterEnd++;
          }
        }
      }
      
      if (validNetworks.length > 0) {
        missing.push({
          date: dateStr,
          networks: validNetworks,
          status: status
        });
      }
    }
  }
  
  // Log summary
  Logger.log(`\n=== Gap Fill Analysis ===`);
  Logger.log(`Found ${missing.length} dates with valid gaps to fill`);
  Logger.log(`Skipped (lifecycle filtering): ${skipped.total}`);
  Logger.log(`  - Before network started: ${skipped.beforeStart}`);
  Logger.log(`  - After network ended (7+ days): ${skipped.afterEnd}`);
  
  // Show top networks with gaps
  const topNetworks = Object.entries(byNetwork)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  
  Logger.log(`\nTop networks with valid gaps:`);
  for (const [netId, count] of topNetworks) {
    const info = lifecycle[netId];
    if (info) {
      Logger.log(`  ${netId}: ${count} gaps (active ${info.firstExpected} to ${info.lastExpected}, Found: ${info.foundCount}, Missing: ${info.missingCount})`);
    }
  }
  
  return missing;
}

/**
 * Update Audit Dashboard Notes column (Column G)
 * Now accumulates messages per date instead of overwriting
 */
function updateRawDataAuditNotes_(dateStr, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Audit Dashboard");
  
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 0; i < data.length; i++) {
    const cellValue = data[i][0];
    
    // Handle Date objects
    let cellDateStr = '';
    if (cellValue instanceof Date && !isNaN(cellValue)) {
      cellDateStr = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      cellDateStr = String(cellValue || '').trim();
    }
    
    if (cellDateStr === dateStr) {
      // Column G = index 6 (0-based), row is i+1 (1-based)
      const currentNote = String(sheet.getRange(i + 1, 7).getValue() || '').trim();
      
      // Don't append duplicates or "Processing..." messages
      if (message.includes('Processing network')) {
        // Temporary status - overwrite
        sheet.getRange(i + 1, 7).setValue(message);
      } else if (!currentNote || currentNote.includes('Processing network')) {
        // First real result or replacing temp status - overwrite
        sheet.getRange(i + 1, 7).setValue(message);
      } else {
        // Append to existing (multiple networks per date)
        sheet.getRange(i + 1, 7).setValue(currentNote + ' | ' + message);
      }
      break; // Don't flush here - let batch updates happen
    }
  }
}

/**
 * Get/Save/Clear state for Raw Data Gap Fill
 */
function getRawGapFillState_() {
  const props = PropertiesService.getDocumentProperties();
  const stateJson = props.getProperty(RAW_GAP_FILL_STATE_KEY);
  return stateJson ? JSON.parse(stateJson) : null;
}

function saveRawGapFillState_(state) {
  try {
    const props = PropertiesService.getDocumentProperties();
    const stateJson = JSON.stringify(state);
    
    // Check size before saving (9KB limit = ~9000 chars)
    if (stateJson.length > 8500) {
      Logger.log(`⚠️ WARNING: State size (${stateJson.length} chars) approaching 9KB limit!`);
      // Consider chunking or cleanup if this happens
    }
    
    props.setProperty(RAW_GAP_FILL_STATE_KEY, stateJson);
  } catch (e) {
    Logger.log(`❌ ERROR saving raw gap fill state: ${e.message}`);
    Logger.log(`State size: ${JSON.stringify(state).length} characters`);
    throw new Error(`Failed to save gap fill state: ${e.message}`);
  }
}

function clearRawGapFillState_() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty(RAW_GAP_FILL_STATE_KEY);
}

/**
 * Download raw data CSVs for a specific date and network
 * Returns { success: boolean, filesFound: number, errorMsg: string }
 */
/**
 * Download raw data for a specific date and network from Gmail
 * Returns: { success, filesFound, errorMsg, details }
 */
function downloadRawDataForDateNetwork_(dateStr, networkId) {
  try {
    // Convert date to YYYYMMDD format for filename matching
    const filenameDateStr = dateStr.replace(/-/g, ''); // 2025-05-11 -> 20250511
    
    // Search Gmail for raw data emails on specific date
    const targetDate = new Date(dateStr);
    
    // Gmail's date search: use after:(day before) before:(day after) to catch emails that arrive just after midnight
    const dayBefore = new Date(targetDate);
    dayBefore.setDate(dayBefore.getDate() - 1);
    const dayAfter = new Date(targetDate);
    dayAfter.setDate(dayAfter.getDate() + 1);
    
    const afterStr = Utilities.formatDate(dayBefore, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    const beforeStr = Utilities.formatDate(dayAfter, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    
    const fullQuery = `subject:"CM360 CPC/CPM FLIGHT QA" has:attachment after:${afterStr} before:${beforeStr}`;
    
    Logger.log(`Searching for ${dateStr} (${filenameDateStr}): ${fullQuery}`);
    
    const threads = GmailApp.search(fullQuery, 0, 50); // Increased to 50 to catch all networks
    
    if (threads.length === 0) {
      return { 
        success: false, 
        filesFound: 0, 
        errorMsg: 'No emails found on target date',
        details: '? No raw data emails found for this date'
      };
    }
    
    let csvsSaved = 0;
    let zipFilesExtracted = 0;
    const savedFiles = [];
    
    // Search through all matching emails
    for (const thread of threads) {
      const messages = thread.getMessages();
      
      for (const message of messages) {
        const attachments = message.getAttachments();
        
        for (const attachment of attachments) {
          const filename = attachment.getName();
          const lowerFilename = filename.toLowerCase();
          
          // Pattern: {networkId}_BKCM360_Global_QA_Check_{YYYYMMDD}_{time}_{reportId}.{csv|zip}
          if (filename.startsWith(`${networkId}_`) && filename.includes(`_${filenameDateStr}_`)) {
            Logger.log(`  ? MATCH: ${filename}`);
            
            // Handle ZIP files - extract CSVs
            if (lowerFilename.endsWith('.zip')) {
              try {
                const zipBlob = attachment.copyBlob();
                const unzipped = Utilities.unzip(zipBlob);
                
                for (const file of unzipped) {
                  const unzippedName = file.getName();
                  if (unzippedName.toLowerCase().endsWith('.csv')) {
                    const saved = saveRawDataFileToDrive_(dateStr, networkId, file, unzippedName);
                    if (saved) {
                      csvsSaved++;
                      savedFiles.push(unzippedName);
                      Logger.log(`    ?? Extracted CSV from ZIP: ${unzippedName}`);
                    }
                  }
                }
                zipFilesExtracted++;
              } catch (unzipError) {
                Logger.log(`    ? Failed to extract ZIP ${filename}: ${unzipError.message}`);
                return {
                  success: false,
                  filesFound: 0,
                  errorMsg: `Failed to extract ZIP: ${unzipError.message}`,
                  details: `? ZIP extraction failed: ${filename}`
                };
              }
            } 
            // Handle CSV files directly
            else if (lowerFilename.endsWith('.csv')) {
              const saved = saveRawDataFileToDrive_(dateStr, networkId, attachment, filename);
              if (saved) {
                csvsSaved++;
                savedFiles.push(filename);
              }
            }
          }
        }
      }
    }
    
    if (csvsSaved === 0) {
      return { 
        success: false, 
        filesFound: 0, 
        errorMsg: `No files found for network ${networkId} on ${filenameDateStr}`,
        details: `? Network ${networkId} not found (may not exist on this date)`
      };
    }
    
    // Success - build details message
    let details = `? Downloaded ${csvsSaved} CSV file${csvsSaved > 1 ? 's' : ''}`;
    if (zipFilesExtracted > 0) {
      details += ` (?? Extracted from ${zipFilesExtracted} ZIP${zipFilesExtracted > 1 ? 's' : ''})`;
    }
    details += ` for network ${networkId}`;
    
    return { 
      success: true, 
      filesFound: csvsSaved, 
      errorMsg: '',
      details: details
    };
    
  } catch (error) {
    Logger.log(`Error downloading raw data for ${dateStr} / ${networkId}: ${error.message}`);
    return { 
      success: false, 
      filesFound: 0, 
      errorMsg: error.message,
      details: `? Error: ${error.message}`
    };
  }
}

/**
 * Save raw data file to Drive with proper folder structure
 * Path: 2025/MM-Month/YYYY-MM-DD/filename
 */
function saveRawDataFileToDrive_(dateStr, networkId, attachment, originalFilename) {
  try {
    const rootFolder = DriveApp.getFolderById(RAW_DATA_ROOT_FOLDER_ID);
    
    // Get or create 2025 folder
    let yearFolder;
    const yearFolders = rootFolder.getFoldersByName('2025');
    if (yearFolders.hasNext()) {
      yearFolder = yearFolders.next();
    } else {
      yearFolder = rootFolder.createFolder('2025');
    }
    
    // Parse date for folder structure
    const dateParts = dateStr.split('-');
    const monthNum = dateParts[1];
    const monthNames = ['', 'January', 'February', 'March', 'April', 'May', 'June', 
                        'July', 'August', 'September', 'October', 'November', 'December'];
    const monthName = monthNames[parseInt(monthNum)];
    const monthFolderName = `${monthNum}-${monthName}`;
    
    // Get or create month folder
    let monthFolder;
    const monthFolders = yearFolder.getFoldersByName(monthFolderName);
    if (monthFolders.hasNext()) {
      monthFolder = monthFolders.next();
    } else {
      monthFolder = yearFolder.createFolder(monthFolderName);
    }
    
    // Get or create date folder
    let dateFolder;
    const dateFolders = monthFolder.getFoldersByName(dateStr);
    if (dateFolders.hasNext()) {
      dateFolder = dateFolders.next();
    } else {
      dateFolder = monthFolder.createFolder(dateStr);
    }
    
    // Check if file already exists
    const existingFiles = dateFolder.getFilesByName(originalFilename);
    if (existingFiles.hasNext()) {
      Logger.log(`File already exists: ${originalFilename}`);
      return true; // Consider it saved
    }
    
    // Save the file
    const blob = attachment.copyBlob();
    dateFolder.createFile(blob.setName(originalFilename));
    
    Logger.log(`� Saved: ${dateStr} / ${networkId} / ${originalFilename}`);
    return true;
    
  } catch (e) {
    Logger.log(`Error saving file to Drive: ${e}`);
    return false;
  }
}

/**
 * Start Raw Data Gap Fill
 */
function startRawDataGapFill() {
  const ui = SpreadsheetApp.getUi();
  
  // Check if already running
  const existingState = getRawGapFillState_();
  if (existingState && existingState.status === 'running') {
    const response = ui.alert(
      '��️ Gap Fill In Progress',
      'Raw data gap fill is already running.\n\nDo you want to continue from where it left off?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      processRawDataGapFillChunk_();
      return;
    } else {
      return;
    }
  }
  
  // Get missing items from audit (with smart lifecycle filtering)
  ui.alert(
    '🔍 Analyzing Data',
    'Analyzing network lifecycles and finding valid gaps...\n\nThis may take a moment.',
    ui.ButtonSet.OK
  );
  
  const missing = getMissingRawDataFromAudit_();
  
  if (missing.length === 0) {
    ui.alert('� No Gaps Found', 'All raw data is complete or all gaps are outside network active periods!', ui.ButtonSet.OK);
    return;
  }
  
  // Build queue: expand each date's networks into separate items
  const queue = [];
  for (const item of missing) {
    for (const networkId of item.networks) {
      queue.push({
        date: item.date,
        network: networkId,
        status: 'pending'
      });
    }
  }
  
  // Initialize state
  const state = {
    status: 'running',
    queue: queue,
    currentIndex: 0,
    startTime: new Date().toISOString(),
    processed: 0,
    successful: 0,
    failed: 0
  };
  
  saveRawGapFillState_(state);
  
  // Ask if user wants auto-resume trigger
  const triggerResponse = ui.alert(
    '� Raw Data Gap Fill Ready',
    `Found ${queue.length} valid date/network combinations to process.\n\n` +
    `This will take multiple runs due to Gmail quota and time limits.\n\n` +
    `Do you want to create an AUTO-RESUME TRIGGER?\n` +
    `(Recommended - will continue automatically every 10 minutes until complete)`,
    ui.ButtonSet.YES_NO
  );
  
  if (triggerResponse === ui.Button.YES) {
    createRawGapFillAutoResumeTrigger();
  }
  
  ui.alert(
    '��️ Starting Now',
    `Processing will begin now and update the Audit Dashboard Notes column.\n\n` +
    `You can check progress with "View Status" or watch the Notes column.`,
    ui.ButtonSet.OK
  );
  
  // Start processing
  processRawDataGapFillChunk_();
}


/**
 * Process a chunk of raw data gap fill (called by trigger or manually)
 */
function processRawDataGapFillChunk_() {
  const startTime = Date.now();
  const state = getRawGapFillState_();
  
  if (!state) {
    Logger.log('No raw data gap fill state found');
    return;
  }
  
  if (state.status !== 'running') {
    Logger.log('Raw data gap fill not running');
    return;
  }
  
  // Track Gmail quota usage
  const docProps = PropertiesService.getDocumentProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const quotaKey = `RAW_GAP_FILL_QUOTA_${today}`;
  const dailyEmailCount = parseInt(docProps.getProperty(quotaKey) || '0', 10);
  
  const MAX_EMAILS_PER_CHUNK = 30; // Conservative limit per run
  const MAX_EMAILS_PER_DAY = 100; // Daily Gmail quota limit
  
  Logger.log(`📧 Starting chunk - Daily quota used: ${dailyEmailCount}/${MAX_EMAILS_PER_DAY}`);
  
  if (dailyEmailCount >= MAX_EMAILS_PER_DAY) {
    Logger.log('⚠️ Daily Gmail quota exhausted - pausing until tomorrow');
    
    // Update audit note for current date
    if (state.currentIndex < state.queue.length) {
      const currentDate = state.queue[state.currentIndex].date;
      updateRawDataAuditNotes_(currentDate, '⏸️ Paused: Daily Gmail quota reached. Will resume tomorrow.');
    }
    
    // Don't show alert if trigger is active
    const triggerId = PropertiesService.getDocumentProperties().getProperty(RAW_GAP_FILL_TRIGGER_KEY);
    if (!triggerId) {
      SpreadsheetApp.getUi().alert(
        '⏸️ Gap Fill Paused',
        `Daily Gmail quota limit reached (${MAX_EMAILS_PER_DAY} emails).\n\n` +
        `Processed today: ${state.processed} items\n` +
        `Remaining: ${state.queue.length - state.currentIndex} items\n\n` +
        `Will automatically resume tomorrow when quota resets.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return;
  }
  
  const queue = state.queue;
  let processedThisRun = 0;
  let emailsThisRun = 0;
  let currentDateResults = {}; // Track results per date for batching
  
  while (state.currentIndex < queue.length) {
    // Re-check quota key in case midnight passed mid-run
    const currentToday = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (currentToday !== today) {
      Logger.log(`🕐 Midnight passed! Quota reset. Old: ${today}, New: ${currentToday}`);
      // Update quota counter with current run, then exit to pick up fresh quota next run
      docProps.setProperty(quotaKey, String(dailyEmailCount + emailsThisRun));
      saveRawGapFillState_(state);
      Logger.log(`Saved state and exiting to leverage fresh daily quota`);
      return;
    }
    
    // Check time budget
    if ((Date.now() - startTime) >= RAW_GAP_FILL_TIME_BUDGET_MS) {
      Logger.log(`⏱️ Time budget reached. Processed ${state.processed}/${queue.length}`);
      saveRawGapFillState_(state);
      
      // Update quota counter
      docProps.setProperty(quotaKey, String(dailyEmailCount + emailsThisRun));
      Logger.log(`📧 Chunk complete - Emails this run: ${emailsThisRun}, Daily total: ${dailyEmailCount + emailsThisRun}`);
      
      // Don't show alert if trigger is active (silent resume)
      const props = PropertiesService.getDocumentProperties();
      const triggerId = props.getProperty(RAW_GAP_FILL_TRIGGER_KEY);
      
      if (!triggerId) {
        // Manual run - show alert
        SpreadsheetApp.getUi().alert(
          '⏱️ Gap Fill Paused',
          `Time limit reached. Progress saved.\n\n` +
          `Processed: ${state.processed}/${queue.length}\n` +
          `Successful: ${state.successful}\n` +
          `Failed: ${state.failed}\n` +
          `Emails used: ${emailsThisRun}\n` +
          `Daily quota: ${dailyEmailCount + emailsThisRun}/${MAX_EMAILS_PER_DAY}\n\n` +
          `Run again to continue, or create auto-resume trigger.`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } else {
        Logger.log(`Auto-resume trigger active - will continue in 10 minutes`);
      }
      return;
    }
    
    // Check Gmail quota limits
    if (emailsThisRun >= MAX_EMAILS_PER_CHUNK || 
        (dailyEmailCount + emailsThisRun) >= MAX_EMAILS_PER_DAY) {
      Logger.log(`📧 Gmail quota limit reached for this chunk (${emailsThisRun} emails processed)`);
      saveRawGapFillState_(state);
      docProps.setProperty(quotaKey, String(dailyEmailCount + emailsThisRun));
      
      const triggerId = PropertiesService.getDocumentProperties().getProperty(RAW_GAP_FILL_TRIGGER_KEY);
      if (!triggerId) {
        SpreadsheetApp.getUi().alert(
          '📧 Gmail Quota Limit',
          `Email quota limit reached for this run.\n\n` +
          `Processed: ${state.processed}/${queue.length}\n` +
          `Emails this run: ${emailsThisRun}\n` +
          `Daily total: ${dailyEmailCount + emailsThisRun}/${MAX_EMAILS_PER_DAY}\n\n` +
          `Will resume automatically if trigger is active.`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
      return;
    }
    
    const item = queue[state.currentIndex];
    const dateStr = item.date;
    const networkId = item.network;
    
    Logger.log(`Processing [${state.currentIndex + 1}/${queue.length}]: ${dateStr} / ${networkId}`);
    
    // Only show processing message if this is the first network for this date
    if (!currentDateResults[dateStr]) {
      currentDateResults[dateStr] = [];
      updateRawDataAuditNotes_(dateStr, `🔍 Processing network ${networkId}...`);
    }
    
    // Try to download from Gmail
    const result = downloadRawDataForDateNetwork_(dateStr, networkId);
    emailsThisRun++; // Count Gmail search
    
    // Accumulate result for this date
    currentDateResults[dateStr].push({
      networkId,
      success: result.success,
      filesFound: result.filesFound,
      details: result.details
    });
    
    if (result.success) {
      item.status = 'success';
      state.successful++;
      Logger.log(`✅ Success: ${dateStr} / ${networkId} - ${result.filesFound} files`);
    } else {
      item.status = 'failed';
      state.failed++;
      Logger.log(`❌ Failed: ${dateStr} / ${networkId} - ${result.errorMsg}`);
    }
    
    state.processed++;
    state.currentIndex++;
    processedThisRun++;
    
    // Check if we've finished all networks for this date
    const nextItem = queue[state.currentIndex];
    const dateChanged = !nextItem || nextItem.date !== dateStr;
    
    if (dateChanged && currentDateResults[dateStr]) {
      // Build combined message for this date
      const results = currentDateResults[dateStr];
      const successful = results.filter(r => r.success);
      const failed = results.filter(r => !r.success);
      
      let summary = [];
      if (successful.length > 0) {
        const totalFiles = successful.reduce((sum, r) => sum + r.filesFound, 0);
        summary.push(`✅ ${successful.length} network${successful.length > 1 ? 's' : ''} (${totalFiles} file${totalFiles > 1 ? 's' : ''})`);
      }
      if (failed.length > 0) {
        summary.push(`❌ ${failed.length} failed: ${failed.map(r => r.networkId).join(', ')}`);
      }
      
      updateRawDataAuditNotes_(dateStr, summary.join(' | '));
      delete currentDateResults[dateStr]; // Clean up
    }
    
    // Save state every 5 items
    if (state.processed % 5 === 0) {
      saveRawGapFillState_(state);
    }
    
    // Throttle to avoid rate limiting (100ms between searches)
    Utilities.sleep(100);
  }
  
  // All done
  state.status = 'completed';
  state.endTime = new Date().toISOString();
  saveRawGapFillState_(state);
  
  // Update final quota counter
  docProps.setProperty(quotaKey, String(dailyEmailCount + emailsThisRun));
  
  const totalTime = (Date.now() - new Date(state.startTime).getTime()) / 1000;
  const thisRunTime = (Date.now() - startTime) / 1000;
  
  // Auto-delete trigger
  deleteRawGapFillAutoResumeTrigger_();
  
  Logger.log(`✅ Raw Data Gap Fill Complete! Total: ${state.processed} items, Time: ${totalTime.toFixed(1)}s`);
  
  SpreadsheetApp.getUi().alert(
    '✅ Raw Data Gap Fill Complete',
    `Finished processing ${state.processed} items.\n\n` +
    `Successful: ${state.successful}\n` +
    `Failed: ${state.failed}\n` +
    `Emails processed: ${emailsThisRun}\n` +
    `Daily quota used: ${dailyEmailCount + emailsThisRun}/${MAX_EMAILS_PER_DAY}\n\n` +
    `Total time: ${(totalTime / 60).toFixed(1)} minutes\n\n` +
    `Check the Audit Dashboard Notes column for details.\n` +
    `Re-run the Raw Data Audit to update statuses.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * View Raw Data Gap Fill status
 */
function viewRawDataGapFillStatus() {
  const state = getRawGapFillState_();
  const ui = SpreadsheetApp.getUi();
  
  if (!state) {
    ui.alert('📊 Gap Fill Status', 'No gap fill in progress.', ui.ButtonSet.OK);
    return;
  }
  
  const progress = state.queue.length > 0 ? ((state.processed / state.queue.length) * 100).toFixed(1) : 0;
  
  ui.alert(
    '📊 Raw Data Gap Fill Status',
    `Status: ${state.status}\n\n` +
    `Progress: ${state.processed}/${state.queue.length} (${progress}%)\n` +
    `Successful: ${state.successful}\n` +
    `Failed: ${state.failed}\n\n` +
    `Started: ${new Date(state.startTime).toLocaleString()}`,
    ui.ButtonSet.OK
  );
}

/**
 * Reset Raw Data Gap Fill
 */
function resetRawDataGapFill() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '��️ Reset Raw Data Gap Fill',
    'This will clear all progress and start over.\n\nAre you sure?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    clearRawGapFillState_();
    ui.alert('� Reset Complete', 'Raw data gap fill has been reset.', ui.ButtonSet.OK);
  }
}

/**
 * Create Auto-Resume Trigger for Raw Data Gap Fill
 */
function createRawGapFillAutoResumeTrigger() {
  deleteRawGapFillAutoResumeTrigger_();
  
  const trigger = ScriptApp.newTrigger('processRawDataGapFillChunk_')
    .timeBased()
    .everyMinutes(10)
    .create();
  
  const props = PropertiesService.getDocumentProperties(); // Changed from ScriptProperties
  props.setProperty(RAW_GAP_FILL_TRIGGER_KEY, trigger.getUniqueId());
  
  SpreadsheetApp.getUi().alert(
    '✅ Auto-Resume Trigger Created',
    'Raw data gap fill will automatically resume every 10 minutes.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Delete Raw Data Gap Fill Auto-Resume Trigger
 */
function deleteRawGapFillAutoResumeTrigger_() {
  const props = PropertiesService.getDocumentProperties(); // Changed from ScriptProperties
  const triggerId = props.getProperty(RAW_GAP_FILL_TRIGGER_KEY);
  
  if (triggerId) {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
        break;
      }
    }
    props.deleteProperty(RAW_GAP_FILL_TRIGGER_KEY);
  }
}

/**
 * Stop Raw Data Gap Fill and Delete Trigger
 */
function stopRawDataGapFillAndDeleteTrigger() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '🛑 Stop Raw Data Gap Fill',
    'This will stop the gap fill process and delete the auto-resume trigger.\n\nProgress will be saved.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    deleteRawGapFillAutoResumeTrigger_();
    
    const state = getRawGapFillState_();
    if (state && state.status === 'running') {
      state.status = 'paused';
      saveRawGapFillState_(state);
    }
    
    ui.alert(
      '� Stopped',
      'Raw data gap fill stopped and trigger deleted.\n\n' +
      'Run "Start Raw Data Gap Fill" to resume.',
      ui.ButtonSet.OK
    );
  }
}

// =====================================================================================================================
// ================================ RAW DATA GAP FILL (TEST MODE - 2 PHASE SYSTEM) ==================================
// =====================================================================================================================

/**
 * Setup Test Audit Dashboard sheet
 */
function setupTestAuditDashboard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Audit Dashboard (TEST)");
  
  if (!sheet) {
    sheet = ss.insertSheet("Audit Dashboard (TEST)");
    
    // Set up columns
    sheet.setColumnWidth(1, 120); // Date
    sheet.setColumnWidth(2, 100); // Phase 1 Status
    sheet.setColumnWidth(3, 150); // Files Downloaded
    sheet.setColumnWidth(4, 100); // Phase 2 Status
    sheet.setColumnWidth(5, 150); // ZIPs Extracted
    sheet.setColumnWidth(6, 300); // Notes
    
    // Headers
    const headers = [
      ["Date", "Phase 1 Status", "Files Downloaded", "Phase 2 Status", "ZIPs Extracted", "Notes"]
    ];
    
    sheet.getRange(1, 1, 1, 6).setValues(headers)
      .setFontWeight("bold")
      .setBackground("#4285f4")
      .setFontColor("#ffffff");
  }
  
  return sheet;
}

/**
 * Phase 1: Download all attachments for date range
 */
function startTestPhase1Download() {
  const ui = SpreadsheetApp.getUi();
  
  // Setup dashboard
  setupTestAuditDashboard_();
  
  // Check for existing state
  const existingState = getTestPhase1State_();
  if (existingState && existingState.status === 'running') {
    const response = ui.alert(
      '⚠️ Phase 1 In Progress',
      'Phase 1 download is already running.\n\nContinue from where it left off?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      processTestPhase1Chunk_();
      return;
    } else {
      return;
    }
  }
  
  // Show calendar picker
  const dateRange = showTestDateRangePicker_();
  if (!dateRange) return;
  
  const startDate = dateRange.startDate;
  const endDate = dateRange.endDate;
  
  // Build date queue
  const queue = [];
  const start = new Date(startDate);
  const end = new Date(endDate);
  
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    queue.push({
      date: dateStr,
      status: 'pending',
      filesDownloaded: 0,
      csvCount: 0,
      zipCount: 0
    });
  }
  
  // Initialize state
  const state = {
    status: 'running',
    queue: queue,
    currentIndex: 0,
    startTime: new Date().toISOString(),
    startDate: startDate,
    endDate: endDate,
    processed: 0,
    totalFiles: 0,
    totalCSVs: 0,
    totalZIPs: 0
  };
  
  saveTestPhase1State_(state);
  
  ui.alert(
    '▶️ Phase 1 Starting',
    `Will download all attachments from ${startDate} to ${endDate}\n\n` +
    `Total dates: ${queue.length}\n\n` +
    `Create auto-resume trigger for automatic processing?`,
    ui.ButtonSet.OK
  );
  
  // Start processing
  processTestPhase1Chunk_();
}

/**
 * Process Phase 1 chunk (download attachments)
 */
function processTestPhase1Chunk_() {
  const startTime = Date.now();
  const state = getTestPhase1State_();
  
  if (!state || state.status !== 'running') {
    Logger.log('Phase 1 not running');
    return;
  }
  
  const TIME_BUDGET_MS = 5.5 * 60 * 1000;
  const queue = state.queue;
  
  // Initialize daily stats tracking
  const dailyStats = getTodayStats_();
  
  while (state.currentIndex < queue.length) {
    // Check time budget
    if ((Date.now() - startTime) >= TIME_BUDGET_MS) {
      Logger.log(`⏱️ Time budget reached`);
      saveTestPhase1State_(state);
      saveTodayStats_(dailyStats);
      return;
    }
    
    const item = queue[state.currentIndex];
    const dateStr = item.date;
    
    Logger.log(`Processing [${state.currentIndex + 1}/${queue.length}]: ${dateStr}`);
    updateTestPhase1Note_(dateStr, `🔍 Downloading...`);
    
    // Try to download - catch quota errors
    let result;
    try {
      result = downloadAllAttachmentsForDate_(dateStr);
      
      // Only count Gmail search if not skipped
      if (!result.skipped) {
        dailyStats.gmailSearches++;
      }
      
      // Clear any previous pause message
      if (item.status === 'paused') {
        item.status = 'pending';
      }
      
    } catch (e) {
      const errorMsg = String(e.message || e);
      
      // Check for Gmail quota error
      if (errorMsg.includes('Service invoked too many times') || 
          errorMsg.includes('quota') || 
          errorMsg.includes('rate limit')) {
        Logger.log(`⚠️ Gmail quota exceeded: ${errorMsg}`);
        updateTestPhase1Note_(dateStr, '⏸️ Paused: Gmail quota exceeded');
        item.status = 'paused';
        saveTestPhase1State_(state);
        saveTodayStats_(dailyStats);
        return; // Exit and wait for next trigger
      } else {
        // Other error - log and skip this date
        Logger.log(`❌ Error processing ${dateStr}: ${errorMsg}`);
        updateTestPhase1Note_(dateStr, `❌ Error: ${errorMsg}`);
        item.status = 'error';
        state.currentIndex++;
        saveTestPhase1State_(state);
        continue;
      }
    }
    
    // Success - update stats
    item.status = 'completed';
    item.filesDownloaded = result.totalFiles;
    item.csvCount = result.csvsSaved;
    item.zipCount = result.zipsSaved;
    
    state.totalFiles += result.totalFiles;
    state.totalCSVs += result.csvsSaved;
    state.totalZIPs += result.zipsSaved;
    state.processed++;
    
    // Track daily progress (only if not skipped)
    if (!result.skipped) {
      dailyStats.csvsSaved += result.csvsSaved;
      dailyStats.zipsSaved += result.zipsSaved;
      dailyStats.datesCompleted.push(dateStr);
    }
    
    state.currentIndex++;
    
    // Update note based on whether skipped or downloaded
    if (result.skipped) {
      updateTestPhase1Note_(dateStr, `⏭️ Skipped (already has ${result.totalFiles} files: ${result.csvsSaved} CSVs, ${result.zipsSaved} ZIPs)`);
    } else {
      updateTestPhase1Note_(dateStr, `✅ Downloaded ${result.totalFiles} files (${result.csvsSaved} CSVs, ${result.zipsSaved} ZIPs)`);
    }
    
    // Save state after each date
    saveTestPhase1State_(state);
    saveTodayStats_(dailyStats);
    
    Utilities.sleep(result.skipped ? 50 : 100); // Faster for skipped dates
  }
  
  // All done - Phase 1 complete!
  state.status = 'completed';
  state.endTime = new Date().toISOString();
  saveTestPhase1State_(state);
  saveTodayStats_(dailyStats);
  
  Logger.log(`✅ Phase 1 Complete! Files: ${state.totalFiles}, CSVs: ${state.totalCSVs}, ZIPs: ${state.totalZIPs}`);
  
  // Send immediate completion email
  sendPhase1CompletionEmail_(state);
  
  // Stop auto-trigger
  const props = PropertiesService.getDocumentProperties();
  const triggerId = props.getProperty(RAW_TEST_PHASE1_TRIGGER_KEY);
  if (triggerId) {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
        Logger.log('✅ Auto-trigger stopped');
        break;
      }
    }
    props.deleteProperty(RAW_TEST_PHASE1_TRIGGER_KEY);
  }
}

/**
 * Download all attachments for a specific date
 */
function downloadAllAttachmentsForDate_(dateStr) {
  // Check if this date already has files downloaded (skip if so)
  const existingFiles = checkExistingFilesForDate_(dateStr);
  if (existingFiles.hasFiles) {
    Logger.log(`  ⏭️ Skipping ${dateStr} - already has ${existingFiles.csvCount} CSVs and ${existingFiles.zipCount} ZIPs`);
    return { 
      totalFiles: existingFiles.csvCount + existingFiles.zipCount, 
      csvsSaved: existingFiles.csvCount, 
      zipsSaved: existingFiles.zipCount,
      skipped: true
    };
  }
  
  // Parse date string: "2025-06-15" → year, month, day
  const dateParts = dateStr.split('-');
  const year = parseInt(dateParts[0], 10);
  const month = parseInt(dateParts[1], 10);
  const day = parseInt(dateParts[2], 10);
  
  // Build Gmail search dates - use YYYY/MM/DD format (matches Google's date format exactly)
  // Gmail's after: includes the specified date and all days after
  // Gmail's before: excludes the specified date (stops before it)
  // To get ONLY the target date, we need: after:(target) before:(target+1)
  
  const afterDate = `${year}/${month}/${day}`;
  
  // Calculate next day using string manipulation instead of Date object
  let nextYear = year;
  let nextMonth = month;
  let nextDay = day + 1;
  
  // Handle month/year rollover
  const daysInMonth = new Date(year, month, 0).getDate(); // Get days in current month
  if (nextDay > daysInMonth) {
    nextDay = 1;
    nextMonth++;
    if (nextMonth > 12) {
      nextMonth = 1;
      nextYear++;
    }
  }
  
  const beforeDate = `${nextYear}/${nextMonth}/${nextDay}`;
  
  // Use exact production search format (lowercase label, hyphen, after/before)
  const query = `label:cm360-qa subject:"CM360 CPC/CPM FLIGHT QA" after:${afterDate} before:${beforeDate}`;
  Logger.log(`Gmail search: ${query}`);
  
  let threads;
  try {
    threads = GmailApp.search(query, 0, 50);
    Logger.log(`  Found ${threads.length} threads for ${dateStr}`);
  } catch (e) {
    Logger.log(`  Gmail search error: ${e.message}`);
    throw e;
  }
  
  let csvsSaved = 0;
  let zipsSaved = 0;
  
  if (threads.length === 0) {
    Logger.log(`  No emails found for ${dateStr}`);
    return { totalFiles: 0, csvsSaved: 0, zipsSaved: 0, skipped: false };
  }
  
  // Get/create date folder
  const dateFolder = getOrCreateTestDateFolder_(dateStr);
  
  for (const thread of threads) {
    const messages = thread.getMessages();
    
    for (const message of messages) {
      const attachments = message.getAttachments();
      
      for (const attachment of attachments) {
        const filename = attachment.getName();
        const lowerFilename = filename.toLowerCase();
        
        if (lowerFilename.endsWith('.csv') || lowerFilename.endsWith('.zip')) {
          // Check if file already exists
          const existingFiles = dateFolder.getFilesByName(filename);
          if (existingFiles.hasNext()) {
            Logger.log(`  Skipping ${filename} (already exists)`);
            continue;
          }
          
          // Save file as-is (no processing)
          dateFolder.createFile(attachment.copyBlob().setName(filename));
          
          if (lowerFilename.endsWith('.csv')) {
            csvsSaved++;
            Logger.log(`  ✅ Saved CSV: ${filename}`);
          } else if (lowerFilename.endsWith('.zip')) {
            zipsSaved++;
            Logger.log(`  📦 Saved ZIP: ${filename}`);
          }
        }
      }
    }
  }
  
  return { 
    totalFiles: csvsSaved + zipsSaved, 
    csvsSaved: csvsSaved, 
    zipsSaved: zipsSaved 
  };
}

/**
 * Check if a date already has files downloaded in Drive
 */
function checkExistingFilesForDate_(dateStr) {
  try {
    const rootFolder = DriveApp.getFolderById(RAW_DATA_TEST_FOLDER_ID);
    const dateParts = dateStr.split('-');
    const year = dateParts[0];
    const month = dateParts[1];
    
    const monthNames = ['','January','February','March','April','May','June','July','August','September','October','November','December'];
    const monthName = `${month}-${monthNames[parseInt(month, 10)]}`;
    
    // Try to find year folder
    const yearFolders = rootFolder.getFoldersByName(year);
    if (!yearFolders.hasNext()) {
      return { hasFiles: false, csvCount: 0, zipCount: 0 };
    }
    const yearFolder = yearFolders.next();
    
    // Try to find month folder
    const monthFolders = yearFolder.getFoldersByName(monthName);
    if (!monthFolders.hasNext()) {
      return { hasFiles: false, csvCount: 0, zipCount: 0 };
    }
    const monthFolder = monthFolders.next();
    
    // Try to find date folder
    const dateFolders = monthFolder.getFoldersByName(dateStr);
    if (!dateFolders.hasNext()) {
      return { hasFiles: false, csvCount: 0, zipCount: 0 };
    }
    const dateFolder = dateFolders.next();
    
    // Count CSV and ZIP files
    let csvCount = 0;
    let zipCount = 0;
    
    const csvFiles = dateFolder.getFilesByType(MimeType.CSV);
    while (csvFiles.hasNext()) {
      csvFiles.next();
      csvCount++;
    }
    
    const zipFiles = dateFolder.getFilesByType(MimeType.ZIP);
    while (zipFiles.hasNext()) {
      zipFiles.next();
      zipCount++;
    }
    
    const totalFiles = csvCount + zipCount;
    
    return {
      hasFiles: totalFiles > 0,
      csvCount: csvCount,
      zipCount: zipCount
    };
    
  } catch (e) {
    Logger.log(`Error checking existing files for ${dateStr}: ${e.message}`);
    return { hasFiles: false, csvCount: 0, zipCount: 0 };
  }
}

/**
 * Get or create date folder in TEST directory
 */
function getOrCreateTestDateFolder_(dateStr) {
  const rootFolder = DriveApp.getFolderById(RAW_DATA_TEST_FOLDER_ID);
  const dateParts = dateStr.split('-');
  const year = dateParts[0];
  const month = dateParts[1];
  
  const monthNames = ['','January','February','March','April','May','June','July','August','September','October','November','December'];
  const monthName = `${month}-${monthNames[parseInt(month, 10)]}`;
  
  // Get/create year folder
  let yearFolder;
  const yearFolders = rootFolder.getFoldersByName(year);
  if (yearFolders.hasNext()) {
    yearFolder = yearFolders.next();
  } else {
    yearFolder = rootFolder.createFolder(year);
  }
  
  // Get/create month folder
  let monthFolder;
  const monthFolders = yearFolder.getFoldersByName(monthName);
  if (monthFolders.hasNext()) {
    monthFolder = monthFolders.next();
  } else {
    monthFolder = yearFolder.createFolder(monthName);
  }
  
  // Get/create date folder
  let dateFolder;
  const dateFolders = monthFolder.getFoldersByName(dateStr);
  if (dateFolders.hasNext()) {
    dateFolder = dateFolders.next();
  } else {
    dateFolder = monthFolder.createFolder(dateStr);
  }
  
  return dateFolder;
}

/**
 * Show calendar date range picker using HTML dialog
 */
function showTestDateRangePicker_() {
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            background: #f5f5f5;
          }
          .container {
            background: white;
            padding: 25px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            max-width: 500px;
            margin: 0 auto;
          }
          h2 {
            color: #1a73e8;
            margin-top: 0;
            font-size: 20px;
          }
          .date-group {
            margin: 20px 0;
          }
          label {
            display: block;
            font-weight: bold;
            margin-bottom: 8px;
            color: #333;
          }
          input[type="date"] {
            width: 100%;
            padding: 10px;
            font-size: 16px;
            border: 2px solid #dadce0;
            border-radius: 4px;
            box-sizing: border-box;
          }
          input[type="date"]:focus {
            outline: none;
            border-color: #1a73e8;
          }
          .button-group {
            margin-top: 25px;
            text-align: right;
          }
          button {
            padding: 10px 24px;
            font-size: 14px;
            font-weight: 500;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            margin-left: 10px;
          }
          .cancel-btn {
            background: #f1f3f4;
            color: #5f6368;
          }
          .cancel-btn:hover {
            background: #e8eaed;
          }
          .submit-btn {
            background: #1a73e8;
            color: white;
          }
          .submit-btn:hover {
            background: #1557b0;
          }
          .info {
            background: #e8f0fe;
            padding: 12px;
            border-radius: 4px;
            margin-top: 15px;
            font-size: 13px;
            color: #1967d2;
          }
          .error {
            background: #fce8e6;
            color: #c5221f;
            padding: 10px;
            border-radius: 4px;
            margin-top: 10px;
            display: none;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>📅 Select Date Range</h2>
          
          <div class="date-group">
            <label for="startDate">Start Date:</label>
            <input type="date" id="startDate" required>
          </div>
          
          <div class="date-group">
            <label for="endDate">End Date:</label>
            <input type="date" id="endDate" required>
          </div>
          
          <div class="info">
            💡 Select the date range for downloading raw data attachments from Gmail.
          </div>
          
          <div class="error" id="error"></div>
          
          <div class="button-group">
            <button class="cancel-btn" onclick="google.script.host.close()">Cancel</button>
            <button class="submit-btn" onclick="submitDates()">Continue</button>
          </div>
        </div>
        
        <script>
          // Set default dates (last 7 days)
          window.onload = function() {
            const today = new Date();
            const weekAgo = new Date();
            weekAgo.setDate(today.getDate() - 7);
            
            document.getElementById('endDate').valueAsDate = today;
            document.getElementById('startDate').valueAsDate = weekAgo;
          };
          
          function submitDates() {
            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;
            const errorDiv = document.getElementById('error');
            
            if (!startDate || !endDate) {
              errorDiv.textContent = 'Please select both start and end dates.';
              errorDiv.style.display = 'block';
              return;
            }
            
            if (new Date(startDate) > new Date(endDate)) {
              errorDiv.textContent = 'Start date must be before or equal to end date.';
              errorDiv.style.display = 'block';
              return;
            }
            
            // Return dates to Apps Script
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .setTestDateRange(startDate, endDate);
          }
        </script>
      </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(550)
    .setHeight(420);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Date Range');
  
  // Wait for user input (handled by setTestDateRange callback)
  return null; // Will be set by callback
}

/**
 * Callback to receive date range from calendar picker
 */
function setTestDateRange(startDate, endDate) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('TEST_DATE_RANGE_START', startDate);
  props.setProperty('TEST_DATE_RANGE_END', endDate);
  
  // Trigger the actual download process
  continueTestPhase1WithDates_();
}

/**
 * Continue Phase 1 download with selected dates
 */
function continueTestPhase1WithDates_() {
  const props = PropertiesService.getDocumentProperties();
  const startDate = props.getProperty('TEST_DATE_RANGE_START');
  const endDate = props.getProperty('TEST_DATE_RANGE_END');
  
  if (!startDate || !endDate) {
    Logger.log('No dates selected');
    return;
  }
  
  // Clear temp properties
  props.deleteProperty('TEST_DATE_RANGE_START');
  props.deleteProperty('TEST_DATE_RANGE_END');
  
  // Build date queue
  const queue = [];
  
  // Parse dates manually to avoid timezone issues
  // Input format: "2025-12-05"
  const startParts = startDate.split('-');
  const endParts = endDate.split('-');
  
  const start = new Date(parseInt(startParts[0]), parseInt(startParts[1]) - 1, parseInt(startParts[2]), 12, 0, 0); // Noon local time
  const end = new Date(parseInt(endParts[0]), parseInt(endParts[1]) - 1, parseInt(endParts[2]), 12, 0, 0);
  
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    queue.push({
      date: dateStr,
      status: 'pending',
      filesDownloaded: 0,
      csvCount: 0,
      zipCount: 0
    });
  }
  
  // Initialize state
  const state = {
    status: 'running',
    queue: queue,
    currentIndex: 0,
    startTime: new Date().toISOString(),
    startDate: startDate,
    endDate: endDate,
    processed: 0,
    totalFiles: 0,
    totalCSVs: 0,
    totalZIPs: 0
  };
  
  saveTestPhase1State_(state);
  
  SpreadsheetApp.getUi().alert(
    '▶️ Phase 1 Starting',
    `Will download all attachments from ${startDate} to ${endDate}\n\n` +
    `Total dates: ${queue.length}\n\n` +
    `Processing will begin now. You can create an auto-resume trigger if needed.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  // Start processing
  processTestPhase1Chunk_();
}

/**
 * Phase 2: Extract all ZIPs in TEST folder
 */
function startTestPhase2Extraction() {
  const ui = SpreadsheetApp.getUi();
  
  // Ask user for scope
  const scopeResponse = ui.alert(
    '📦 Select Extraction Scope',
    'Choose scope:\n\n' +
    'YES = Extract all ZIPs (full folder)\n' +
    'NO = Select specific month(s)\n' +
    'CANCEL = Exit',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (scopeResponse === ui.Button.CANCEL) return;
  
  let monthFilter = null; // null = all months
  
  if (scopeResponse === ui.Button.NO) {
    // Ask which month(s) to extract
    const monthInput = ui.prompt(
      '📅 Select Month(s)',
      'Enter month(s) to extract ZIPs from:\n\n' +
      'Examples:\n' +
      '  "4" = April only\n' +
      '  "6-8" = June through August\n' +
      '  "4,6,9" = April, June, September\n\n' +
      'Enter month number(s):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (monthInput.getSelectedButton() !== ui.Button.OK) return;
    
    const input = monthInput.getResponseText().trim();
    monthFilter = parseMonthFilter_(input);
    
    if (!monthFilter) {
      ui.alert('❌ Invalid Input', 'Could not parse month selection. Please try again.', ui.ButtonSet.OK);
      return;
    }
  }
  
  const scopeMsg = monthFilter 
    ? `Month(s): ${monthFilter.join(', ')}`
    : 'All months';
  
  const confirmMsg = 
    `📦 Phase 2: ZIP Extraction\n\n` +
    `Scope: ${scopeMsg}\n\n` +
    `This will:\n` +
    `1. Scan for ZIP files\n` +
    `2. Extract CSVs from each ZIP\n` +
    `3. Verify extraction succeeded\n` +
    `4. Delete ZIP only if verified\n\n` +
    `Continue?`;
  
  const confirm = ui.alert('📦 Confirm Extraction', confirmMsg, ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) return;
  
  // Count ZIPs (quick scan without storing full list)
  const zipCount = countZipsInTestFolder_(monthFilter);
  
  if (zipCount === 0) {
    ui.alert('✅ No ZIPs Found', `No ZIP files found in selected scope.\n\nScope: ${scopeMsg}`, ui.ButtonSet.OK);
    return;
  }
  
  // Initialize lightweight state (no ZIP list stored)
  const state = {
    status: 'running',
    monthFilter: monthFilter,
    startTime: new Date().toISOString(),
    processed: 0,
    successful: 0,
    failed: 0,
    csvsExtracted: 0,
    estimatedTotal: zipCount
  };
  
  saveTestPhase2State_(state);
  
  ui.alert(
    '▶️ Starting ZIP Extraction',
    `Found ~${zipCount} ZIP files to extract.\n\n` +
    `Scope: ${scopeMsg}\n\n` +
    `This will run in chunks with auto-resume.\n` +
    `Processing on-the-fly (no memory limits).`,
    ui.ButtonSet.OK
  );
  
  processTestPhase2Chunk_();
}

/**
 * Process Phase 2 chunk (extract ZIPs) - ON-THE-FLY SCANNING
 */
function processTestPhase2Chunk_() {
  const startTime = Date.now();
  const state = getTestPhase2State_();
  
  if (!state || state.status !== 'running') {
    Logger.log('Phase 2 not running');
    return;
  }
  
  const TIME_BUDGET_MS = 5.5 * 60 * 1000;
  const BATCH_SIZE = 50; // Process 50 ZIPs per chunk
  
  // Scan for ZIPs on-the-fly (don't store in state)
  const rootFolder = DriveApp.getFolderById(RAW_DATA_TEST_FOLDER_ID);
  let processedThisRun = 0;
  let foundAnyZips = false;
  
  // Scan year folders
  const yearFolders = rootFolder.getFolders();
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    
    // Scan month folders
    const monthFolders = yearFolder.getFolders();
    while (monthFolders.hasNext()) {
      const monthFolder = monthFolders.next();
      
      // Apply month filter if specified
      if (state.monthFilter) {
        const monthName = monthFolder.getName();
        const monthNum = parseInt(monthName.split('-')[0], 10);
        if (!state.monthFilter.includes(monthNum)) {
          continue;
        }
      }
      
      // Scan date folders
      const dateFolders = monthFolder.getFolders();
      while (dateFolders.hasNext()) {
        const dateFolder = dateFolders.next();
        const dateStr = dateFolder.getName();
        
        // Find ZIP files in this date folder
        const files = dateFolder.getFilesByType(MimeType.ZIP);
        while (files.hasNext()) {
          foundAnyZips = true;
          
          // Check time budget
          if ((Date.now() - startTime) >= TIME_BUDGET_MS) {
            Logger.log(`⏱️ Time budget reached - processed ${processedThisRun} ZIPs this run`);
            saveTestPhase2State_(state);
            return;
          }
          
          // Check batch limit
          if (processedThisRun >= BATCH_SIZE) {
            Logger.log(`📦 Batch limit reached (${BATCH_SIZE}) - saving and resuming`);
            saveTestPhase2State_(state);
            return;
          }
          
          const zipFile = files.next();
          const zipInfo = {
            id: zipFile.getId(),
            name: zipFile.getName(),
            folderId: dateFolder.getId(),
            dateStr: dateStr
          };
          
          Logger.log(`Extracting [${state.processed + 1}]: ${zipInfo.name}`);
          
          try {
            const zipBlob = zipFile.getBlob();
            
            // Extract files
            const unzipped = Utilities.unzip(zipBlob);
            let csvsThisZip = 0;
            const extractedFiles = [];
            
            for (const file of unzipped) {
              if (file.getName().toLowerCase().endsWith('.csv')) {
                const createdFile = dateFolder.createFile(file);
                extractedFiles.push(createdFile.getId());
                csvsThisZip++;
              }
            }
            
            // VERIFICATION: Ensure all CSVs were created
            let allVerified = true;
            for (const fileId of extractedFiles) {
              try {
                DriveApp.getFileById(fileId);
              } catch (e) {
                allVerified = false;
                Logger.log(`  ⚠️ Verification failed for file ID: ${fileId}`);
                break;
              }
            }
            
            if (allVerified && csvsThisZip > 0) {
              // Only delete ZIP if all CSVs verified
              zipFile.setTrashed(true);
              
              state.successful++;
              state.csvsExtracted += csvsThisZip;
              Logger.log(`  ✅ Extracted ${csvsThisZip} CSVs, verified, deleted ZIP`);
              
              updateTestPhase2Note_(zipInfo.dateStr, `📦 Extracted ${csvsThisZip} CSVs from ${zipInfo.name}`);
            } else {
              state.failed++;
              Logger.log(`  ⚠️ Verification failed or no CSVs - ZIP kept for safety`);
              updateTestPhase2Note_(zipInfo.dateStr, `⚠️ Extraction issue: ${zipInfo.name} - ZIP preserved`);
            }
            
          } catch (e) {
            Logger.log(`  ❌ Failed: ${e.message}`);
            state.failed++;
            updateTestPhase2Note_(zipInfo.dateStr, `❌ Extraction failed: ${zipInfo.name} - ZIP preserved`);
          }
          
          state.processed++;
          processedThisRun++;
          
          // Save state every 10 ZIPs
          if (state.processed % 10 === 0) {
            saveTestPhase2State_(state);
          }
        }
      }
    }
  }
  
  // If we scanned everything and found no ZIPs, we're done
  if (!foundAnyZips) {
    state.status = 'completed';
    state.endTime = new Date().toISOString();
    saveTestPhase2State_(state);
    
    Logger.log(`✅ Phase 2 Complete! Processed: ${state.processed}, Successful: ${state.successful}, CSVs: ${state.csvsExtracted}`);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `ZIP extraction finished! ZIPs: ${state.processed} | Successful: ${state.successful} | Failed: ${state.failed} | CSVs: ${state.csvsExtracted}`,
      '✅ Phase 2 Complete',
      10
    );
  }
}

/**
 * Count ZIP files in TEST folder (lightweight, doesn't store list)
 */
function countZipsInTestFolder_(monthFilter = null) {
  const rootFolder = DriveApp.getFolderById(RAW_DATA_TEST_FOLDER_ID);
  let count = 0;
  
  const yearFolders = rootFolder.getFolders();
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    
    const monthFolders = yearFolder.getFolders();
    while (monthFolders.hasNext()) {
      const monthFolder = monthFolders.next();
      
      if (monthFilter) {
        const monthName = monthFolder.getName();
        const monthNum = parseInt(monthName.split('-')[0], 10);
        if (!monthFilter.includes(monthNum)) {
          continue;
        }
      }
      
      const dateFolders = monthFolder.getFolders();
      while (dateFolders.hasNext()) {
        const dateFolder = dateFolders.next();
        const files = dateFolder.getFilesByType(MimeType.ZIP);
        while (files.hasNext()) {
          files.next();
          count++;
        }
      }
    }
  }
  
  return count;
}

// State management functions
function getTestPhase1State_() {
  const props = PropertiesService.getDocumentProperties();
  const json = props.getProperty(RAW_TEST_PHASE1_STATE_KEY);
  return json ? JSON.parse(json) : null;
}

function saveTestPhase1State_(state) {
  try {
    const props = PropertiesService.getDocumentProperties();
    props.setProperty(RAW_TEST_PHASE1_STATE_KEY, JSON.stringify(state));
  } catch (e) {
    Logger.log(`❌ Error saving Phase 1 state: ${e.message}`);
  }
}

function getTestPhase2State_() {
  const props = PropertiesService.getDocumentProperties();
  const json = props.getProperty(RAW_TEST_PHASE2_STATE_KEY);
  return json ? JSON.parse(json) : null;
}

function saveTestPhase2State_(state) {
  try {
    const props = PropertiesService.getDocumentProperties();
    props.setProperty(RAW_TEST_PHASE2_STATE_KEY, JSON.stringify(state));
  } catch (e) {
    Logger.log(`❌ Error saving Phase 2 state: ${e.message}`);
  }
}

function updateTestPhase1Note_(dateStr, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Audit Dashboard (TEST)");
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === dateStr) {
      sheet.getRange(i + 1, 6).setValue(message);
      return;
    }
  }
  
  // Add new row if date not found
  sheet.appendRow([dateStr, '', '', '', '', message]);
}

function updateTestPhase2Note_(dateStr, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Audit Dashboard (TEST)");
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === dateStr) {
      const current = String(sheet.getRange(i + 1, 6).getValue() || '');
      sheet.getRange(i + 1, 6).setValue(current + ' | ' + message);
      return;
    }
  }
}

// View status functions
function viewTestPhase1Status() {
  const state = getTestPhase1State_();
  const ui = SpreadsheetApp.getUi();
  
  if (!state) {
    ui.alert('📊 Phase 1 Status', 'No Phase 1 download in progress.', ui.ButtonSet.OK);
    return;
  }
  
  const progress = state.queue.length > 0 ? ((state.processed / state.queue.length) * 100).toFixed(1) : 0;
  
  ui.alert(
    '📊 Phase 1 Download Status',
    `Status: ${state.status}\n\n` +
    `Date Range: ${state.startDate} to ${state.endDate}\n` +
    `Progress: ${state.processed}/${state.queue.length} (${progress}%)\n` +
    `Files Downloaded: ${state.totalFiles}\n` +
    `  ├─ CSVs: ${state.totalCSVs}\n` +
    `  └─ ZIPs: ${state.totalZIPs}\n\n` +
    `Started: ${new Date(state.startTime).toLocaleString()}`,
    ui.ButtonSet.OK
  );
}

function viewTestPhase2Status() {
  const state = getTestPhase2State_();
  const ui = SpreadsheetApp.getUi();
  
  if (!state) {
    ui.alert('📊 Phase 2 Status', 'No Phase 2 extraction in progress.', ui.ButtonSet.OK);
    return;
  }
  
  const estimatedTotal = state.estimatedTotal || state.processed || 'Unknown';
  const progress = (estimatedTotal !== 'Unknown' && estimatedTotal > 0) 
    ? ((state.processed / estimatedTotal) * 100).toFixed(1) 
    : '0.0';
  
  const scopeMsg = state.monthFilter 
    ? `Month(s): ${state.monthFilter.join(', ')}`
    : 'All months';
  
  ui.alert(
    '📊 Phase 2 Extraction Status',
    `Status: ${state.status}\n` +
    `Scope: ${scopeMsg}\n\n` +
    `Progress: ${state.processed}/${estimatedTotal} (~${progress}%)\n` +
    `Successful: ${state.successful}\n` +
    `Failed: ${state.failed}\n` +
    `CSVs Extracted: ${state.csvsExtracted}\n\n` +
    `Started: ${new Date(state.startTime).toLocaleString()}`,
    ui.ButtonSet.OK
  );
}

// Trigger functions
function createTestPhase1Trigger() {
  // Delete existing trigger if any
  const props = PropertiesService.getDocumentProperties();
  const existingId = props.getProperty(RAW_TEST_PHASE1_TRIGGER_KEY);
  if (existingId) {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === existingId) {
        ScriptApp.deleteTrigger(trigger);
        break;
      }
    }
  }
  
  // Create new trigger
  const trigger = ScriptApp.newTrigger('processTestPhase1Chunk_')
    .timeBased()
    .everyMinutes(10)
    .create();
  
  props.setProperty(RAW_TEST_PHASE1_TRIGGER_KEY, trigger.getUniqueId());
  
  SpreadsheetApp.getUi().alert(
    '✅ Phase 1 Trigger Created',
    'Phase 1 will auto-resume every 10 minutes.\n\n' +
    'The system will pause if Gmail quota is exceeded and auto-retry when quota is available.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function createTestPhase2Trigger() {
  const trigger = ScriptApp.newTrigger('processTestPhase2Chunk_')
    .timeBased()
    .everyMinutes(10)
    .create();
  
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(RAW_TEST_PHASE2_TRIGGER_KEY, trigger.getUniqueId());
  
  SpreadsheetApp.getUi().alert(
    '✅ Phase 2 Trigger Created',
    'Phase 2 will auto-resume every 10 minutes.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function stopAllTestTriggers() {
  const props = PropertiesService.getDocumentProperties();
  const scriptProps = PropertiesService.getScriptProperties();
  
  const phase1Id = props.getProperty(RAW_TEST_PHASE1_TRIGGER_KEY);
  const phase2Id = props.getProperty(RAW_TEST_PHASE2_TRIGGER_KEY);
  const dailyId = props.getProperty(RAW_TEST_DAILY_TRIGGER_KEY);
  
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  for (const trigger of triggers) {
    const id = trigger.getUniqueId();
    const funcName = trigger.getHandlerFunction();
    
    // Delete TEST-related triggers
    if (id === phase1Id || id === phase2Id || id === dailyId || 
        funcName === 'runWeeklyAutoDownload' || 
        funcName === 'fixIncompleteDatesAuto' ||
        funcName === 'runDailyMorningPhase1' ||
        funcName === 'runDailyMorningPhase2') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  }
  
  props.deleteProperty(RAW_TEST_PHASE1_TRIGGER_KEY);
  props.deleteProperty(RAW_TEST_PHASE2_TRIGGER_KEY);
  props.deleteProperty(RAW_TEST_DAILY_TRIGGER_KEY);
  
  // Also delete daily morning triggers
  deleteDailyMorningTriggers_();
  
  SpreadsheetApp.getUi().alert(
    '🛑 Triggers Stopped',
    `Deleted ${deletedCount} TEST mode trigger(s).\n\nThis includes:\n` +
    `• Phase 1 auto-resume\n` +
    `• Phase 2 auto-resume\n` +
    `• Daily email (7:30 PM)\n` +
    `• Daily morning automation (7-8 AM)\n` +
    `• Weekly auto-download (Sat 11:30 PM)\n` +
    `• Fix incomplete auto-resume`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Audit and cleanup functions
/**
 * Parse month filter input (e.g., "4", "6-8", "4,6,9")
 */
function parseMonthFilter_(input) {
  if (!input) return null;
  
  const months = [];
  const parts = input.split(',');
  
  for (const part of parts) {
    const trimmed = part.trim();
    
    if (trimmed.includes('-')) {
      // Range: "6-8"
      const range = trimmed.split('-');
      const start = parseInt(range[0], 10);
      const end = parseInt(range[1], 10);
      
      if (isNaN(start) || isNaN(end) || start < 1 || end > 12 || start > end) {
        return null;
      }
      
      for (let m = start; m <= end; m++) {
        if (!months.includes(m)) months.push(m);
      }
    } else {
      // Single month: "4"
      const month = parseInt(trimmed, 10);
      if (isNaN(month) || month < 1 || month > 12) {
        return null;
      }
      if (!months.includes(month)) months.push(month);
    }
  }
  
  return months.sort((a, b) => a - b);
}

/**
 * Helper: Apply status color coding to a cell
 */
function applyStatusColor_(cell, status) {
  if (status.includes('✅')) {
    cell.setBackground("#d4edda").setFontColor("#155724").setFontWeight("bold");
  } else if (status.includes('❌')) {
    cell.setBackground("#f8d7da").setFontColor("#721c24").setFontWeight("bold");
  } else if (status.includes('⚠️')) {
    cell.setBackground("#fff3cd").setFontColor("#856404").setFontWeight("bold");
  } else if (status.includes('⚪')) {
    cell.setBackground("#e9ecef").setFontColor("#6c757d");
  }
}

/**
 * Comprehensive audit: Compare Gmail email counts vs Drive file counts
 * Writes detailed results to "OVERALL CROSS AUDIT GTE" sheet
 */
function auditTestFolder() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ask user to select scope
  const scopeResponse = ui.alert(
    '🔍 Select Audit Scope',
    'Choose audit scope:\n\n' +
    'YES = Full audit (all dates, ~5-10 min)\n' +
    'NO = Select specific month(s)\n' +
    'CANCEL = Exit',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (scopeResponse === ui.Button.CANCEL) return;
  
  let monthFilter = null; // null = all months
  
  if (scopeResponse === ui.Button.NO) {
    // Ask which month(s) to audit
    const monthInput = ui.prompt(
      '📅 Select Month(s)',
      'Enter month(s) to audit:\n\n' +
      'Examples:\n' +
      '  "4" = April only\n' +
      '  "6-8" = June through August\n' +
      '  "4,6,9" = April, June, September\n\n' +
      'Enter month number(s):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (monthInput.getSelectedButton() !== ui.Button.OK) return;
    
    const input = monthInput.getResponseText().trim();
    monthFilter = parseMonthFilter_(input);
    
    if (!monthFilter) {
      ui.alert('❌ Invalid Input', 'Could not parse month selection. Please try again.', ui.ButtonSet.OK);
      return;
    }
  }
  
  const scopeMsg = monthFilter 
    ? `Month(s): ${monthFilter.join(', ')}`
    : 'All dates';
  
  const response = ui.alert(
    '🔍 Start Gmail vs Drive Audit',
    `Audit Scope: ${scopeMsg}\n\n` +
    'This will:\n' +
    '1. Scan Drive folders\n' +
    '2. Check Gmail for each date\n' +
    '3. Compare: 1 email = 1 file\n' +
    '4. Write to "OVERALL CROSS AUDIT GTE"\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  ui.alert('⏳ Step 1/3: Scanning Drive', 'Reading date folders and counting files...', ui.ButtonSet.OK);
  
  // Step 1: Scan Drive for all dates and file counts
  const rootFolder = DriveApp.getFolderById(RAW_DATA_TEST_FOLDER_ID);
  const driveData = [];
  
  const yearFolders = rootFolder.getFolders();
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    const monthFolders = yearFolder.getFolders();
    
    while (monthFolders.hasNext()) {
      const monthFolder = monthFolders.next();
      const dateFolders = monthFolder.getFolders();
      
      while (dateFolders.hasNext()) {
        const dateFolder = dateFolders.next();
        const dateStr = dateFolder.getName();
        
        // Apply month filter if specified
        if (monthFilter) {
          const dateMonth = parseInt(dateStr.split('-')[1], 10);
          if (!monthFilter.includes(dateMonth)) {
            continue; // Skip this date
          }
        }
        
        let csvCount = 0;
        let zipCount = 0;
        
        const csvs = dateFolder.getFilesByType(MimeType.CSV);
        while (csvs.hasNext()) {
          csvs.next();
          csvCount++;
        }
        
        const zips = dateFolder.getFilesByType(MimeType.ZIP);
        while (zips.hasNext()) {
          zips.next();
          zipCount++;
        }
        
        driveData.push({
          date: dateStr,
          csvCount: csvCount,
          zipCount: zipCount,
          totalFiles: csvCount + zipCount
        });
      }
    }
  }
  
  driveData.sort((a, b) => a.date.localeCompare(b.date));
  
  if (driveData.length === 0) {
    ui.alert('⚠️ No Dates Found', 'No dates found matching the selected month(s).', ui.ButtonSet.OK);
    return;
  }
  
  const estimatedMinutes = Math.ceil(driveData.length / 60); // ~1 date per second
  ui.alert('⏳ Step 2/3: Checking Gmail', `Found ${driveData.length} dates in Drive.\n\nNow checking Gmail for each date...\n\nEstimated time: ${estimatedMinutes} minute(s)`, ui.ButtonSet.OK);
  
  // Step 2: Check Gmail for email counts for each date
  const auditResults = [];
  let mismatches = 0;
  let missing = 0;
  let complete = 0;
  let extraFiles = 0;
  
  for (let i = 0; i < driveData.length; i++) {
    const item = driveData[i];
    
    // Parse date
    const dateParts = item.date.split('-');
    const year = parseInt(dateParts[0], 10);
    const month = parseInt(dateParts[1], 10);
    const day = parseInt(dateParts[2], 10);
    
    // Build Gmail query with timezone-safe date calculation
    const afterDate = `${year}/${month}/${day}`;
    let nextYear = year;
    let nextMonth = month;
    let nextDay = day + 1;
    
    const daysInMonth = new Date(year, month, 0).getDate();
    if (nextDay > daysInMonth) {
      nextDay = 1;
      nextMonth++;
      if (nextMonth > 12) {
        nextMonth = 1;
        nextYear++;
      }
    }
    
    const beforeDate = `${nextYear}/${nextMonth}/${nextDay}`;
    const query = `label:cm360-qa subject:"CM360 CPC/CPM FLIGHT QA" after:${afterDate} before:${beforeDate}`;
    
    let emailCount = 0;
    try {
      const threads = GmailApp.search(query, 0, 50);
      emailCount = threads.length;
    } catch (e) {
      emailCount = -1; // Error flag
    }
    
    // Compare: 1 Gmail email should = 1 Drive file
    let status = '✅ Match';
    let notes = '';
    let matchDiff = 0;
    
    if (emailCount === -1) {
      status = '⚠️ Error';
      notes = 'Gmail search failed - quota or API error';
    } else if (item.totalFiles === 0 && emailCount === 0) {
      status = '⚪ No Data';
      notes = 'No emails or files (expected for pre-production dates)';
    } else if (item.totalFiles === 0 && emailCount > 0) {
      status = '❌ Missing All';
      notes = `${emailCount} emails in Gmail but 0 files in Drive - NEEDS DOWNLOAD`;
      matchDiff = emailCount;
      missing++;
    } else if (emailCount > item.totalFiles) {
      status = '⚠️ Incomplete';
      matchDiff = emailCount - item.totalFiles;
      notes = `${emailCount} emails but only ${item.totalFiles} files - MISSING ${matchDiff} files`;
      mismatches++;
    } else if (emailCount < item.totalFiles) {
      status = '⚠️ Extra Files';
      matchDiff = item.totalFiles - emailCount;
      notes = `${emailCount} emails but ${item.totalFiles} files - ${matchDiff} EXTRA files (possible duplicates)`;
      extraFiles++;
    } else {
      status = '✅ Match';
      notes = `Perfect match: ${emailCount} emails = ${item.totalFiles} files`;
      complete++;
    }
    
    auditResults.push([
      item.date,
      emailCount >= 0 ? emailCount : 'ERROR',
      item.totalFiles,
      item.csvCount,
      item.zipCount,
      matchDiff !== 0 ? matchDiff : '-',
      status,
      notes
    ]);
    
    // Progress logging every 25 dates
    if ((i + 1) % 25 === 0) {
      Logger.log(`Audited ${i + 1}/${driveData.length} dates...`);
    }
  }
  
  ui.alert('⏳ Step 3/3: Writing Results', `Audit complete!\n\nWriting ${auditResults.length} rows to spreadsheet...`, ui.ButtonSet.OK);
  
  // Step 3: Write results to audit sheet (append mode - keeps existing data)
  let sheet = ss.getSheetByName("OVERALL CROSS AUDIT GTE");
  if (!sheet) {
    sheet = ss.insertSheet("OVERALL CROSS AUDIT GTE");
  }
  
  // Check if sheet already has data
  const lastRow = sheet.getLastRow();
  const hasExistingData = lastRow > 1; // More than just header
  
  if (!hasExistingData) {
    // First-time setup: Create headers
    sheet.clear();
    
    // Set up columns
    sheet.setColumnWidth(1, 120);  // Date
    sheet.setColumnWidth(2, 110);  // Gmail Emails
    sheet.setColumnWidth(3, 110);  // Drive Files
    sheet.setColumnWidth(4, 80);   // CSVs
    sheet.setColumnWidth(5, 80);   // ZIPs
    sheet.setColumnWidth(6, 100);  // Difference
    sheet.setColumnWidth(7, 130);  // Status
    sheet.setColumnWidth(8, 450);  // Notes
    
    // Headers
    const headers = [
      ["Date", "Gmail Emails", "Drive Files", "CSVs", "ZIPs", "Difference", "Status", "Notes"]
    ];
    
    sheet.getRange(1, 1, 1, 8).setValues(headers)
      .setFontWeight("bold")
      .setBackground("#4285f4")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    
    // Freeze header row
    sheet.setFrozenRows(1);
  }
  
  // Update or append audit results
  if (auditResults.length > 0) {
    if (hasExistingData) {
      // Read existing data (skip header)
      const existingData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      const existingDates = new Map();
      
      // Build map of existing dates to row numbers
      for (let i = 0; i < existingData.length; i++) {
        const dateStr = String(existingData[i][0]);
        if (dateStr) {
          existingDates.set(dateStr, i + 2); // +2 for header row and 1-based index
        }
      }
      
      // Update existing rows and collect new rows
      const newRows = [];
      for (const result of auditResults) {
        const dateStr = result[0];
        const rowNum = existingDates.get(dateStr);
        
        if (rowNum) {
          // Update existing row
          sheet.getRange(rowNum, 1, 1, 8).setValues([result]);
          
          // Apply color coding
          const statusCell = sheet.getRange(rowNum, 7);
          const status = result[6];
          applyStatusColor_(statusCell, status);
        } else {
          // New date - will append later
          newRows.push(result);
        }
      }
      
      // Append new rows
      if (newRows.length > 0) {
        const nextRow = sheet.getLastRow() + 1;
        sheet.getRange(nextRow, 1, newRows.length, 8).setValues(newRows);
        
        // Color-code new rows
        for (let i = 0; i < newRows.length; i++) {
          const statusCell = sheet.getRange(nextRow + i, 7);
          const status = newRows[i][6];
          applyStatusColor_(statusCell, status);
        }
      }
      
      Logger.log(`Updated ${auditResults.length - newRows.length} existing rows, added ${newRows.length} new rows`);
      
    } else {
      // No existing data - write all results
      sheet.getRange(2, 1, auditResults.length, 8).setValues(auditResults);
      
      // Color-code status column
      for (let i = 0; i < auditResults.length; i++) {
        const statusCell = sheet.getRange(i + 2, 7);
        const status = auditResults[i][6];
        applyStatusColor_(statusCell, status);
      }
    }
    
    // Center-align numeric columns for all data
    const dataRows = sheet.getLastRow() - 1;
    if (dataRows > 0) {
      sheet.getRange(2, 2, dataRows, 5).setHorizontalAlignment("center");
    }
  }
  
  // Summary
  const totalIssues = missing + mismatches + extraFiles;
  const summary = 
    `✅ Cross-Audit Complete!\n\n` +
    `📊 SUMMARY:\n` +
    `Total Dates: ${driveData.length}\n` +
    `✅ Perfect Match: ${complete}\n` +
    `⚠️ Incomplete: ${mismatches}\n` +
    `❌ Missing All Files: ${missing}\n` +
    `⚠️ Extra Files: ${extraFiles}\n\n` +
    `${totalIssues > 0 ? `⚠️ ${totalIssues} dates need attention` : `✅ All dates verified!`}\n\n` +
    `Results written to "OVERALL CROSS AUDIT GTE" sheet.`;
  
  ui.alert('🔍 Audit Results', summary, ui.ButtonSet.OK);
}

/**
 * Fix Incomplete Dates: Archive incomplete folders and auto-restart download
 * Uses chunked execution to avoid timeout (processes 10 dates at a time)
 */
function fixIncompleteDatesAuto() {
  const props = PropertiesService.getScriptProperties();
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if resuming from previous run
  const stateJson = props.getProperty('FIX_INCOMPLETE_STATE');
  
  if (!stateJson) {
    // === INITIAL RUN: Identify incomplete dates ===
    
    // Step 1: Check if audit sheet exists
    const auditSheet = ss.getSheetByName("OVERALL CROSS AUDIT GTE");
    if (!auditSheet) {
      ui.alert(
        '❌ Audit Not Found',
        'Please run "Audit Test Folder" first to identify incomplete dates.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Step 2: Read audit data and find incomplete dates
    const data = auditSheet.getDataRange().getValues();
    if (data.length <= 1) {
      ui.alert('❌ No Data', 'Audit sheet is empty. Run audit first.', ui.ButtonSet.OK);
      return;
    }
    
    const incompleteDates = [];
    
    for (let i = 1; i < data.length; i++) {
      const dateCell = data[i][0];
      const status = String(data[i][6] || '');
      
      // Convert date to YYYY-MM-DD string format
      let dateStr = '';
      if (dateCell instanceof Date) {
        const year = dateCell.getFullYear();
        const month = String(dateCell.getMonth() + 1).padStart(2, '0');
        const day = String(dateCell.getDate()).padStart(2, '0');
        dateStr = `${year}-${month}-${day}`;
      } else {
        dateStr = String(dateCell || '').trim();
      }
      
      if (!dateStr || dateStr === '') continue;
      
      // Find rows with "⚠️ Incomplete" or "⚠️ Extra Files"
      if (status.includes('⚠️')) {
        incompleteDates.push({
          date: dateStr,
          status: status,
          gmailEmails: data[i][1],
          driveFiles: data[i][2],
          difference: data[i][5]
        });
      }
    }
    
    if (incompleteDates.length === 0) {
      ui.alert(
        '✅ All Complete',
        'No incomplete dates found! All dates match Gmail.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Step 3: Show summary and ask confirmation
    const summary = incompleteDates.slice(0, 10).map(d => 
      `  ${d.date}: ${d.status} (${d.difference} files)`
    ).join('\n');
    
    const confirmMsg = 
      `🔧 Fix Incomplete Dates\n\n` +
      `Found ${incompleteDates.length} incomplete dates:\n\n` +
      `${summary}` +
      (incompleteDates.length > 10 ? `\n  ...and ${incompleteDates.length - 10} more` : '') +
      `\n\n` +
      `This will:\n` +
      `1. Archive incomplete folders (10 per run)\n` +
      `2. Re-download fresh data from Gmail\n` +
      `3. Auto-resume every 10 min if needed\n\n` +
      `Total runs needed: ${Math.ceil(incompleteDates.length / 10)}\n` +
      `Estimated time: ${Math.ceil(incompleteDates.length / 10) * 10} minutes\n\n` +
      `Continue?`;
    
    const response = ui.alert('🔧 Fix Incomplete Dates', confirmMsg, ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) return;
    
    // Initialize state
    const state = {
      allDates: incompleteDates.map(d => d.date).sort(),
      archivedDates: [],
      currentIndex: 0,
      totalCount: incompleteDates.length,
      startTime: new Date().toISOString()
    };
    
    props.setProperty('FIX_INCOMPLETE_STATE', JSON.stringify(state));
    
    // Create auto-resume trigger
    createFixIncompleteResumeTrigger_();
    
    Logger.log(`🔧 Fix started: ${state.totalCount} dates to archive`);
  }
  
  // === CHUNKED EXECUTION: Process batch ===
  
  const state = JSON.parse(props.getProperty('FIX_INCOMPLETE_STATE'));
  const CHUNK_SIZE = 10; // Archive 10 dates per execution
  const rootFolder = DriveApp.getFolderById(RAW_DATA_TEST_FOLDER_ID);
  const archiveFolder = DriveApp.getFolderById(RAW_DATA_TEST_ARCHIVE_FOLDER_ID);
  
  const endIndex = Math.min(state.currentIndex + CHUNK_SIZE, state.allDates.length);
  const chunk = state.allDates.slice(state.currentIndex, endIndex);
  
  Logger.log(`📦 Processing batch: ${state.currentIndex + 1}-${endIndex} of ${state.totalCount}`);
  
  // Archive this chunk
  for (const dateStr of chunk) {
    try {
      const dateFolder = findDateFolder_(rootFolder, dateStr);
      
      if (dateFolder) {
        // Create archive subfolder with date name (or use existing)
        let archiveDateFolder;
        const existingArchiveFolders = archiveFolder.getFoldersByName(dateStr);
        
        if (existingArchiveFolders.hasNext()) {
          archiveDateFolder = existingArchiveFolders.next();
          Logger.log(`📁 Using existing archive folder: ${dateStr}`);
        } else {
          archiveDateFolder = archiveFolder.createFolder(dateStr);
        }
        
        // Move all files to archive
        const files = dateFolder.getFiles();
        let fileCount = 0;
        while (files.hasNext()) {
          const file = files.next();
          file.moveTo(archiveDateFolder);
          fileCount++;
        }
        
        // Delete empty date folder from TEST
        dateFolder.setTrashed(true);
        
        state.archivedDates.push(dateStr);
        Logger.log(`✅ Archived: ${dateStr} (${fileCount} files) - ${state.archivedDates.length}/${state.totalCount}`);
      } else {
        // Folder not found - likely already archived in previous run
        state.archivedDates.push(dateStr);
        Logger.log(`⏭️ Already archived: ${dateStr} (${state.archivedDates.length}/${state.totalCount})`);
      }
    } catch (e) {
      Logger.log(`❌ Error archiving ${dateStr}: ${e.message}`);
      // Still count it to avoid infinite loop
      state.archivedDates.push(dateStr);
    }
  }
  
  // Update state
  state.currentIndex = endIndex;
  
  if (state.currentIndex >= state.allDates.length) {
    // === ALL ARCHIVING COMPLETE ===
    
    Logger.log(`✅ All archiving complete: ${state.archivedDates.length} dates`);
    
    // Start re-download Phase 1
    const datesToRedownload = state.allDates.sort();
    const startDate = datesToRedownload[0];
    const endDate = datesToRedownload[datesToRedownload.length - 1];
    
    continueTestPhase1WithDates_(startDate, endDate, datesToRedownload);
    createTestPhase1Trigger();
    
    // Clean up state and trigger
    props.deleteProperty('FIX_INCOMPLETE_STATE');
    deleteFixIncompleteResumeTrigger_();
    
    // Final summary
    const finalMsg =
      `✅ Archive Complete!\n\n` +
      `Archived: ${state.archivedDates.length} folders\n` +
      `Re-downloading: ${datesToRedownload.length} dates\n` +
      `Date Range: ${startDate} to ${endDate}\n\n` +
      `Phase 1 auto-resume trigger created.\n\n` +
      `Monitor progress:\n` +
      `  • View Download Progress (Phase 1)\n` +
      `  • Check logs for details`;
    
    SpreadsheetApp.getUi().alert('🔧 Fix Process Started', finalMsg, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } else {
    // === MORE BATCHES REMAINING ===
    
    props.setProperty('FIX_INCOMPLETE_STATE', JSON.stringify(state));
    
    Logger.log(`⏳ Progress: ${state.archivedDates.length}/${state.totalCount} archived`);
    Logger.log(`⏰ Next batch in 10 minutes...`);
  }
}

/**
 * Create auto-resume trigger for fix incomplete process
 */
function createFixIncompleteResumeTrigger_() {
  deleteFixIncompleteResumeTrigger_(); // Remove existing
  
  ScriptApp.newTrigger('fixIncompleteDatesAuto')
    .timeBased()
    .after(10 * 60 * 1000) // 10 minutes
    .create();
  
  Logger.log('⏰ Created fix incomplete resume trigger (10 min)');
}

/**
 * Delete fix incomplete resume trigger
 */
function deleteFixIncompleteResumeTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'fixIncompleteDatesAuto') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('🗑️ Deleted fix incomplete resume trigger');
    }
  }
}

/**
 * Helper: Find date folder in Drive hierarchy
 */
function findDateFolder_(rootFolder, dateStr) {
  const dateParts = dateStr.split('-');
  const year = dateParts[0];
  const month = dateParts[1];
  const monthNames = ['', 'January', 'February', 'March', 'April', 'May', 'June', 
                      'July', 'August', 'September', 'October', 'November', 'December'];
  const monthName = monthNames[parseInt(month, 10)];
  const monthFolderName = `${month}-${monthName}`;
  
  // Navigate: Year → Month → Date
  const yearFolders = rootFolder.getFolders();
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    if (yearFolder.getName() === year) {
      
      const monthFolders = yearFolder.getFolders();
      while (monthFolders.hasNext()) {
        const monthFolder = monthFolders.next();
        if (monthFolder.getName() === monthFolderName) {
          
          const dateFolders = monthFolder.getFolders();
          while (dateFolders.hasNext()) {
            const dateFolder = dateFolders.next();
            if (dateFolder.getName() === dateStr) {
              return dateFolder;
            }
          }
        }
      }
    }
  }
  
  return null;
}

function cleanupAndVerifyTest() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🧹 Cleanup & Verify',
    'This will:\n' +
    '1. Extract any remaining ZIPs\n' +
    '2. Delete empty folders\n' +
    '3. Generate final report\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    // Count remaining ZIPs (lightweight)
    const zipCount = countZipsInTestFolder_();
    
    if (zipCount > 0) {
      ui.alert('📦 Found ZIPs', `Found ${zipCount} remaining ZIPs. Running extraction...`, ui.ButtonSet.OK);
      startTestPhase2Extraction();
    } else {
      ui.alert('✅ Complete', 'No ZIPs found. All files are extracted!', ui.ButtonSet.OK);
    }
  }
}

function resetTestMode() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '⚠️ Reset TEST Mode',
    'This will clear all TEST progress and state.\n\n' +
    'Files in Drive will NOT be deleted.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const props = PropertiesService.getDocumentProperties();
    props.deleteProperty(RAW_TEST_PHASE1_STATE_KEY);
    props.deleteProperty(RAW_TEST_PHASE2_STATE_KEY);
    props.deleteProperty(RAW_TEST_DAILY_STATS_KEY);
    
    stopAllTestTriggers();
    
    ui.alert('✅ Reset Complete', 'TEST mode has been reset.', ui.ButtonSet.OK);
  }
}

// =====================================================================================================================
// ====================================== DAILY MORNING AUTOMATION SYSTEM =============================================
// =====================================================================================================================

const RAW_TEST_DAILY_MORNING_TRIGGER_KEY = 'raw_test_daily_morning_trigger_id';

/**
 * Setup Daily Morning Automation (runs 7-8 AM every day)
 * Phase 1 at 7:00 AM → Phase 2 at 8:00 AM
 */
function setupDailyMorningAutomation() {
  const ui = SpreadsheetApp.getUi();
  
  // Confirm setup
  const confirmMsg = 
    '🌅 Daily Morning Automation Setup\n\n' +
    'Schedule:\n' +
    '  • 7:00 AM - Phase 1 (Download new emails)\n' +
    '  • 8:00 AM - Phase 2 (Extract ZIPs)\n\n' +
    'This will run EVERY DAY:\n' +
    '1. Download all new CM360 email attachments (Phase 1)\n' +
    '2. Extract CSVs from ZIPs (Phase 2) 1 hour later\n' +
    '3. Auto-resume if needed (chunked processing)\n\n' +
    'Perfect for daily data collection!\n\n' +
    'Create triggers?';
  
  const confirm = ui.alert('🌅 Confirm Daily Automation', confirmMsg, ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) return;
  
  // Delete existing daily morning triggers if any
  deleteDailyMorningTriggers_();
  
  // Create Phase 1 trigger (7:00 AM daily)
  const phase1Trigger = ScriptApp.newTrigger('runDailyMorningPhase1')
    .timeBased()
    .atHour(7)
    .nearMinute(0)
    .everyDays(1)
    .create();
  
  // Create Phase 2 trigger (8:00 AM daily)
  const phase2Trigger = ScriptApp.newTrigger('runDailyMorningPhase2')
    .timeBased()
    .atHour(8)
    .nearMinute(0)
    .everyDays(1)
    .create();
  
  // Store trigger IDs
  const props = PropertiesService.getScriptProperties();
  props.setProperty(RAW_TEST_DAILY_MORNING_TRIGGER_KEY, JSON.stringify({
    phase1: phase1Trigger.getUniqueId(),
    phase2: phase2Trigger.getUniqueId()
  }));
  
  Logger.log('✅ Created daily morning automation triggers (7 AM Phase 1, 8 AM Phase 2)');
  
  ui.alert(
    '✅ Daily Automation Enabled',
    'Daily morning automation created!\n\n' +
    'Phase 1 (Download): Every day at 7:00 AM\n' +
    'Phase 2 (Extract): Every day at 8:00 AM\n\n' +
    'Both phases will auto-resume if needed.\n\n' +
    'To stop: Menu → 🧪 Raw Data Gap Fill (TEST) → 🛑 Stop All TEST Triggers',
    ui.ButtonSet.OK
  );
}

/**
 * Daily Morning Phase 1 (runs at 7:00 AM)
 */
function runDailyMorningPhase1() {
  Logger.log('🌅 Daily Morning Phase 1 - Download Started');
  
  try {
    // Download yesterday's data (emails arrive overnight)
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const dateStr = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    Logger.log(`📥 Downloading data for ${dateStr}`);
    
    // Use the existing download function
    const result = downloadAllAttachmentsForDate_(dateStr);
    
    Logger.log(`✅ Phase 1 Complete: ${result.totalFiles} files (${result.csvs} CSVs, ${result.zips} ZIPs)`);
    
    // If download had issues, create auto-resume trigger
    if (result.errors && result.errors.length > 0) {
      createTestPhase1Trigger();
      Logger.log('⏰ Created Phase 1 auto-resume trigger (errors detected)');
    }
    
  } catch (e) {
    Logger.log(`❌ Daily Phase 1 Error: ${e.message}`);
    
    // Send error email
    MailApp.sendEmail({
      to: RAW_TEST_EMAIL_TARGET,
      subject: '❌ Daily Phase 1 Failed',
      body: `Daily morning Phase 1 (download) failed at ${new Date().toLocaleString()}:\n\n${e.message}\n\n${e.stack || ''}`
    });
  }
}

/**
 * Daily Morning Phase 2 (runs at 8:00 AM)
 */
function runDailyMorningPhase2() {
  Logger.log('🌅 Daily Morning Phase 2 - Extract Started');
  
  try {
    // Check if there are any ZIPs to extract
    const zipCount = countZipsInTestFolder_();
    
    if (zipCount === 0) {
      Logger.log('ℹ️ No ZIPs found - skipping Phase 2');
      return;
    }
    
    Logger.log(`📦 Found ${zipCount} ZIPs to extract`);
    
    // Initialize Phase 2 state for all months (or yesterday's month specifically)
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const monthNum = yesterday.getMonth() + 1;
    
    const state = {
      status: 'running',
      monthFilter: null, // Process all ZIPs
      startTime: new Date().toISOString(),
      processed: 0,
      successful: 0,
      failed: 0,
      csvsExtracted: 0,
      estimatedTotal: zipCount
    };
    
    saveTestPhase2State_(state);
    
    // Run first Phase 2 batch
    processTestPhase2Chunk_();
    
    // Create auto-resume trigger for Phase 2
    createTestPhase2Trigger();
    Logger.log('⏰ Created Phase 2 auto-resume trigger');
    
  } catch (e) {
    Logger.log(`❌ Daily Phase 2 Error: ${e.message}`);
    
    // Send error email
    MailApp.sendEmail({
      to: RAW_TEST_EMAIL_TARGET,
      subject: '❌ Daily Phase 2 Failed',
      body: `Daily morning Phase 2 (extract) failed at ${new Date().toLocaleString()}:\n\n${e.message}\n\n${e.stack || ''}`
    });
  }
}

/**
 * Delete daily morning triggers
 */
function deleteDailyMorningTriggers_() {
  const props = PropertiesService.getScriptProperties();
  const triggerDataJson = props.getProperty(RAW_TEST_DAILY_MORNING_TRIGGER_KEY);
  
  if (!triggerDataJson) return;
  
  try {
    const triggerData = JSON.parse(triggerDataJson);
    const triggers = ScriptApp.getProjectTriggers();
    
    for (const trigger of triggers) {
      const id = trigger.getUniqueId();
      if (id === triggerData.phase1 || id === triggerData.phase2) {
        ScriptApp.deleteTrigger(trigger);
        Logger.log(`🗑️ Deleted daily morning trigger: ${trigger.getHandlerFunction()}`);
      }
    }
    
    props.deleteProperty(RAW_TEST_DAILY_MORNING_TRIGGER_KEY);
  } catch (e) {
    Logger.log(`⚠️ Error deleting daily morning triggers: ${e.message}`);
  }
}

// =====================================================================================================================
// ======================================= WEEKLY AUTO-DOWNLOAD SYSTEM =================================================
// =====================================================================================================================

/**
 * Setup Weekly Automatic Download (Saturday 11:30 PM)
 */
function setupWeeklyAutoDownload() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  // Ask user for email notification preference
  const emailPref = ui.alert(
    '📧 Email Notifications',
    'When should you receive email notifications?\n\n' +
    'YES = Always (weekly summary)\n' +
    'NO = Only when errors occur\n' +
    'CANCEL = Exit setup',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (emailPref === ui.Button.CANCEL) return;
  
  const emailMode = emailPref === ui.Button.YES ? 'always' : 'errors-only';
  props.setProperty('WEEKLY_AUTO_EMAIL_MODE', emailMode);
  
  // Confirm setup
  const confirmMsg = 
    '📅 Weekly Auto-Download Setup\n\n' +
    'Schedule: Every Saturday at 11:30 PM\n\n' +
    'Actions:\n' +
    '1. Fix incomplete dates (auto-archive + re-download)\n' +
    '2. Download past 7 days (skip existing files)\n' +
    '3. Auto-trigger Phase 2 (ZIP extraction)\n' +
    '4. Run audit to verify completeness\n' +
    `5. Email ${emailMode === 'always' ? 'weekly summary' : 'only if errors'}\n\n` +
    'This will run automatically every week.\n\n' +
    'Create trigger?';
  
  const confirm = ui.alert('📅 Confirm Setup', confirmMsg, ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) return;
  
  // Delete existing weekly trigger if any
  deleteWeeklyAutoDownloadTrigger_();
  
  // Create new weekly trigger (Saturday 11:30 PM)
  ScriptApp.newTrigger('runWeeklyAutoDownload')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(23)
    .nearMinute(30)
    .create();
  
  Logger.log('✅ Created weekly auto-download trigger (Saturday 11:30 PM)');
  
  ui.alert(
    '✅ Setup Complete',
    `Weekly auto-download trigger created!\n\n` +
    `Next run: This Saturday at 11:30 PM\n\n` +
    `Email notifications: ${emailMode}\n\n` +
    `To stop: CM360 QA Tools → 🧪 Raw Data Gap Fill (TEST) → 🛑 Stop All TEST Triggers`,
    ui.ButtonSet.OK
  );
}

/**
 * Main weekly automation function (runs Saturday 11:30 PM)
 */
function runWeeklyAutoDownload() {
  const startTime = new Date();
  const props = PropertiesService.getScriptProperties();
  const emailMode = props.getProperty('WEEKLY_AUTO_EMAIL_MODE') || 'always';
  
  const log = [];
  let errors = [];
  
  try {
    log.push('🤖 WEEKLY AUTO-DOWNLOAD STARTED');
    log.push(`Time: ${startTime.toLocaleString()}`);
    log.push('');
    
    // STEP 1: Fix incomplete dates from previous weeks
    log.push('STEP 1: Checking for incomplete dates...');
    Logger.log('📋 Step 1: Checking for incomplete dates');
    
    const incompleteCount = checkAndFixIncompleteDates_();
    
    if (incompleteCount > 0) {
      log.push(`  ✅ Fixed ${incompleteCount} incomplete dates (archived + re-download started)`);
      Logger.log(`✅ Fixed ${incompleteCount} incomplete dates`);
    } else {
      log.push(`  ✅ No incomplete dates found`);
      Logger.log('✅ No incomplete dates');
    }
    log.push('');
    
    // STEP 2: Download past 7 days (smart: skip existing)
    log.push('STEP 2: Downloading past 7 days...');
    Logger.log('📥 Step 2: Downloading past 7 days');
    
    const downloadResult = downloadPast7DaysSmart_();
    
    log.push(`  📊 Scanned: ${downloadResult.scanned} dates`);
    log.push(`  ⏭️ Skipped (already have files): ${downloadResult.skipped}`);
    log.push(`  📥 Downloaded: ${downloadResult.downloaded} dates`);
    log.push(`  📄 Total files: ${downloadResult.totalFiles} (${downloadResult.csvs} CSVs, ${downloadResult.zips} ZIPs)`);
    
    if (downloadResult.errors.length > 0) {
      log.push(`  ⚠️ Errors: ${downloadResult.errors.length}`);
      errors = errors.concat(downloadResult.errors);
    }
    log.push('');
    
    // STEP 3: Auto-create Phase 2 trigger (ZIP extraction)
    log.push('STEP 3: Setting up ZIP extraction...');
    Logger.log('📦 Step 3: Creating Phase 2 trigger');
    
    if (downloadResult.zips > 0) {
      createTestPhase2Trigger();
      log.push(`  ✅ Phase 2 trigger created (will extract ${downloadResult.zips} ZIPs)`);
      Logger.log(`✅ Phase 2 trigger created`);
    } else {
      log.push(`  ℹ️ No ZIPs to extract`);
      Logger.log('ℹ️ No ZIPs to extract');
    }
    log.push('');
    
    // STEP 4: Run audit on the week's dates
    log.push('STEP 4: Running audit...');
    Logger.log('🔍 Step 4: Running audit');
    
    const auditResult = auditPast7Days_();
    
    log.push(`  📊 Audited: ${auditResult.total} dates`);
    log.push(`  ✅ Perfect match: ${auditResult.complete}`);
    log.push(`  ⚠️ Incomplete: ${auditResult.incomplete}`);
    log.push(`  ❌ Missing all: ${auditResult.missing}`);
    log.push(`  ⚠️ Extra files: ${auditResult.extraFiles}`);
    
    if (auditResult.incomplete > 0 || auditResult.missing > 0) {
      errors.push(`Audit found ${auditResult.incomplete + auditResult.missing} dates with issues`);
    }
    log.push('');
    
    // Summary
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);
    
    log.push('✅ WEEKLY AUTO-DOWNLOAD COMPLETE');
    log.push(`Duration: ${duration} seconds`);
    log.push(`Errors: ${errors.length}`);
    
    Logger.log('✅ Weekly auto-download complete');
    
    // Send email if needed
    if (emailMode === 'always' || (emailMode === 'errors-only' && errors.length > 0)) {
      sendWeeklyEmail_(log.join('\n'), errors);
    }
    
  } catch (e) {
    log.push('');
    log.push(`❌ FATAL ERROR: ${e.message}`);
    log.push(`Stack: ${e.stack}`);
    
    Logger.log(`❌ Fatal error: ${e.message}`);
    
    // Always send email on fatal error
    sendWeeklyEmail_(log.join('\n'), [`FATAL: ${e.message}`]);
  }
}

/**
 * Check audit sheet and fix incomplete dates automatically
 */
function checkAndFixIncompleteDates_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName("OVERALL CROSS AUDIT GTE");
  
  if (!auditSheet) return 0; // No audit data yet
  
  const data = auditSheet.getDataRange().getValues();
  if (data.length <= 1) return 0;
  
  const incompleteDates = [];
  
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][6] || '');
    if (status.includes('⚠️')) {
      const dateCell = data[i][0];
      let dateStr = '';
      
      if (dateCell instanceof Date) {
        const year = dateCell.getFullYear();
        const month = String(dateCell.getMonth() + 1).padStart(2, '0');
        const day = String(dateCell.getDate()).padStart(2, '0');
        dateStr = `${year}-${month}-${day}`;
      } else {
        dateStr = String(dateCell || '').trim();
      }
      
      if (dateStr) incompleteDates.push(dateStr);
    }
  }
  
  if (incompleteDates.length === 0) return 0;
  
  Logger.log(`Found ${incompleteDates.length} incomplete dates - starting fix process`);
  
  // Trigger the fix process (it will handle archiving and re-downloading)
  const props = PropertiesService.getScriptProperties();
  const state = {
    allDates: incompleteDates,
    archivedDates: [],
    currentIndex: 0,
    totalCount: incompleteDates.length,
    startTime: new Date().toISOString()
  };
  
  props.setProperty('FIX_INCOMPLETE_STATE', JSON.stringify(state));
  createFixIncompleteResumeTrigger_();
  
  // Run first batch immediately
  fixIncompleteDatesAuto();
  
  return incompleteDates.length;
}

/**
 * Download past 7 days (smart: skip dates that already have files)
 */
function downloadPast7DaysSmart_() {
  const result = {
    scanned: 0,
    skipped: 0,
    downloaded: 0,
    totalFiles: 0,
    csvs: 0,
    zips: 0,
    errors: []
  };
  
  // Calculate past 7 days
  const today = new Date();
  const dates = [];
  
  for (let i = 6; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    dates.push(dateStr);
  }
  
  result.scanned = dates.length;
  
  // Check each date and download if needed
  for (const dateStr of dates) {
    try {
      const existing = checkExistingFilesForDate_(dateStr);
      
      if (existing.hasFiles) {
        result.skipped++;
        Logger.log(`⏭️ Skipping ${dateStr} - already has ${existing.csvCount + existing.zipCount} files`);
        continue;
      }
      
      // Download this date
      Logger.log(`📥 Downloading ${dateStr}...`);
      const downloadResult = downloadAllAttachmentsForDate_(dateStr);
      
      if (!downloadResult.skipped) {
        result.downloaded++;
        result.totalFiles += downloadResult.totalFiles;
        result.csvs += downloadResult.csvsSaved;
        result.zips += downloadResult.zipsSaved;
        
        Logger.log(`✅ Downloaded ${dateStr}: ${downloadResult.totalFiles} files`);
      }
      
    } catch (e) {
      result.errors.push(`${dateStr}: ${e.message}`);
      Logger.log(`❌ Error downloading ${dateStr}: ${e.message}`);
    }
  }
  
  return result;
}

/**
 * Audit past 7 days and return summary
 */
function auditPast7Days_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rootFolder = DriveApp.getFolderById(RAW_DATA_TEST_FOLDER_ID);
  
  const result = {
    total: 0,
    complete: 0,
    incomplete: 0,
    missing: 0,
    extraFiles: 0
  };
  
  // Calculate past 7 days
  const today = new Date();
  const dates = [];
  
  for (let i = 6; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    dates.push(dateStr);
  }
  
  const auditResults = [];
  
  for (const dateStr of dates) {
    result.total++;
    
    // Get Drive file count
    const existing = checkExistingFilesForDate_(dateStr);
    const driveFiles = existing.csvCount + existing.zipCount;
    
    // Get Gmail email count
    const dateParts = dateStr.split('-');
    const year = parseInt(dateParts[0], 10);
    const month = parseInt(dateParts[1], 10);
    const day = parseInt(dateParts[2], 10);
    
    const afterDate = `${year}/${month}/${day}`;
    let nextYear = year;
    let nextMonth = month;
    let nextDay = day + 1;
    
    const daysInMonth = new Date(year, month, 0).getDate();
    if (nextDay > daysInMonth) {
      nextDay = 1;
      nextMonth++;
      if (nextMonth > 12) {
        nextMonth = 1;
        nextYear++;
      }
    }
    
    const beforeDate = `${nextYear}/${nextMonth}/${nextDay}`;
    const query = `label:cm360-qa subject:"CM360 CPC/CPM FLIGHT QA" after:${afterDate} before:${beforeDate}`;
    
    let emailCount = 0;
    try {
      const threads = GmailApp.search(query, 0, 50);
      emailCount = threads.length;
    } catch (e) {
      emailCount = -1;
    }
    
    // Categorize
    let status = '✅ Match';
    let notes = '';
    let matchDiff = 0;
    
    if (emailCount === -1) {
      status = '⚠️ Error';
      notes = 'Gmail search failed';
    } else if (driveFiles === 0 && emailCount === 0) {
      status = '⚪ No Data';
      notes = 'No emails or files';
    } else if (driveFiles === 0 && emailCount > 0) {
      status = '❌ Missing All';
      notes = `${emailCount} emails but 0 files`;
      matchDiff = emailCount;
      result.missing++;
    } else if (emailCount > driveFiles) {
      status = '⚠️ Incomplete';
      matchDiff = emailCount - driveFiles;
      notes = `Missing ${matchDiff} files`;
      result.incomplete++;
    } else if (emailCount < driveFiles) {
      status = '⚠️ Extra Files';
      matchDiff = driveFiles - emailCount;
      notes = `${matchDiff} extra files`;
      result.extraFiles++;
    } else {
      status = '✅ Match';
      notes = 'Perfect match';
      result.complete++;
    }
    
    auditResults.push([
      dateStr,
      emailCount >= 0 ? emailCount : 'ERROR',
      driveFiles,
      existing.csvCount,
      existing.zipCount,
      matchDiff !== 0 ? matchDiff : '-',
      status,
      notes
    ]);
  }
  
  // Write to audit sheet (append mode)
  let sheet = ss.getSheetByName("OVERALL CROSS AUDIT GTE");
  if (!sheet) {
    sheet = ss.insertSheet("OVERALL CROSS AUDIT GTE");
    
    // Setup headers
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 110);
    sheet.setColumnWidth(3, 110);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 130);
    sheet.setColumnWidth(8, 450);
    
    const headers = [["Date", "Gmail Emails", "Drive Files", "CSVs", "ZIPs", "Difference", "Status", "Notes"]];
    sheet.getRange(1, 1, 1, 8).setValues(headers)
      .setFontWeight("bold")
      .setBackground("#4285f4")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    
    sheet.setFrozenRows(1);
  }
  
  // Update or append results
  const lastRow = sheet.getLastRow();
  const hasExistingData = lastRow > 1;
  
  if (hasExistingData) {
    const existingData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const existingDates = new Map();
    
    for (let i = 0; i < existingData.length; i++) {
      const dateStr = String(existingData[i][0]);
      if (dateStr) existingDates.set(dateStr, i + 2);
    }
    
    const newRows = [];
    for (const result of auditResults) {
      const dateStr = result[0];
      const rowNum = existingDates.get(dateStr);
      
      if (rowNum) {
        sheet.getRange(rowNum, 1, 1, 8).setValues([result]);
        applyStatusColor_(sheet.getRange(rowNum, 7), result[6]);
      } else {
        newRows.push(result);
      }
    }
    
    if (newRows.length > 0) {
      const nextRow = sheet.getLastRow() + 1;
      sheet.getRange(nextRow, 1, newRows.length, 8).setValues(newRows);
      
      for (let i = 0; i < newRows.length; i++) {
        applyStatusColor_(sheet.getRange(nextRow + i, 7), newRows[i][6]);
      }
    }
  } else {
    sheet.getRange(2, 1, auditResults.length, 8).setValues(auditResults);
    
    for (let i = 0; i < auditResults.length; i++) {
      applyStatusColor_(sheet.getRange(i + 2, 7), auditResults[i][6]);
    }
  }
  
  // Center-align numeric columns
  const dataRows = sheet.getLastRow() - 1;
  if (dataRows > 0) {
    sheet.getRange(2, 2, dataRows, 5).setHorizontalAlignment("center");
  }
  
  return result;
}

/**
 * Send weekly summary email
 */
function sendWeeklyEmail_(logText, errors) {
  const userEmail = Session.getActiveUser().getEmail();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  const subject = errors.length > 0 
    ? `⚠️ CM360 Weekly Auto-Download - ERRORS - ${today}`
    : `✅ CM360 Weekly Auto-Download - Success - ${today}`;
  
  const body = 
    `CM360 Weekly Auto-Download Report\n` +
    `Date: ${today}\n\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n` +
    `${logText}\n\n` +
    `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  
  if (errors.length > 0) {
    const errorList = errors.map((e, i) => `${i + 1}. ${e}`).join('\n');
    const fullBody = body + `\n⚠️ ERRORS:\n${errorList}\n\nPlease check the logs for details.`;
    GmailApp.sendEmail(userEmail, subject, fullBody);
  } else {
    GmailApp.sendEmail(userEmail, subject, body);
  }
  
  Logger.log(`📧 Email sent to ${userEmail}`);
}

/**
 * Delete weekly auto-download trigger
 */
function deleteWeeklyAutoDownloadTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'runWeeklyAutoDownload') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('🗑️ Deleted existing weekly auto-download trigger');
    }
  }
}

// =====================================================================================================================
// ===================================== DAILY STATS & EMAIL FUNCTIONS ==============================================
// =====================================================================================================================

/**
 * Get today's statistics (resets daily)
 */
function getTodayStats_() {
  const props = PropertiesService.getDocumentProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const json = props.getProperty(RAW_TEST_DAILY_STATS_KEY);
  
  if (json) {
    const stats = JSON.parse(json);
    // Check if stats are from today
    if (stats.date === today) {
      return stats;
    }
  }
  
  // Create new stats for today
  return {
    date: today,
    gmailSearches: 0,
    csvsSaved: 0,
    zipsSaved: 0,
    datesCompleted: [],
    lastEmailSent: null
  };
}

/**
 * Save today's statistics
 */
function saveTodayStats_(stats) {
  try {
    const props = PropertiesService.getDocumentProperties();
    props.setProperty(RAW_TEST_DAILY_STATS_KEY, JSON.stringify(stats));
  } catch (e) {
    Logger.log(`❌ Error saving daily stats: ${e.message}`);
  }
}

/**
 * View today's progress
 */
function viewTodayProgress() {
  const stats = getTodayStats_();
  const state = getTestPhase1State_();
  
  let progressMsg = '';
  if (state && state.queue) {
    const totalDates = state.queue.length;
    const completed = state.processed || 0;
    const remaining = totalDates - completed;
    const percentComplete = totalDates > 0 ? ((completed / totalDates) * 100).toFixed(1) : 0;
    
    progressMsg = `\nOverall Progress: ${completed}/${totalDates} dates (${percentComplete}%)\nRemaining: ${remaining} dates`;
    
    // Estimate time remaining
    if (stats.datesCompleted.length > 0 && remaining > 0) {
      const daysElapsed = state.startTime ? Math.max(1, Math.ceil((Date.now() - new Date(state.startTime).getTime()) / (24 * 60 * 60 * 1000))) : 1;
      const ratePerDay = completed / daysElapsed;
      if (ratePerDay > 0) {
        const daysRemaining = Math.ceil(remaining / ratePerDay);
        progressMsg += `\nETA: ~${daysRemaining} days`;
      }
    }
  }
  
  const datesMsg = stats.datesCompleted.length > 0 
    ? `\n\nDates completed today:\n${stats.datesCompleted.join('\n')}`
    : '\n\nNo dates completed yet today.';
  
  SpreadsheetApp.getUi().alert(
    '📊 Today\'s Progress',
    `Date: ${stats.date}\n\n` +
    `Gmail Searches: ${stats.gmailSearches}\n` +
    `CSVs Saved: ${stats.csvsSaved}\n` +
    `ZIPs Saved: ${stats.zipsSaved}\n` +
    `Dates Completed: ${stats.datesCompleted.length}` +
    progressMsg +
    datesMsg,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Create daily email trigger (7-8 PM)
 */
function createDailyEmailTrigger() {
  const props = PropertiesService.getDocumentProperties();
  const existingId = props.getProperty(RAW_TEST_DAILY_TRIGGER_KEY);
  
  // Delete existing trigger if any
  if (existingId) {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === existingId) {
        ScriptApp.deleteTrigger(trigger);
        break;
      }
    }
  }
  
  // Create new trigger at 7:30 PM daily
  const trigger = ScriptApp.newTrigger('sendDailyProgressEmail_')
    .timeBased()
    .atHour(19) // 7 PM
    .nearMinute(30)
    .everyDays(1)
    .create();
  
  props.setProperty(RAW_TEST_DAILY_TRIGGER_KEY, trigger.getUniqueId());
  
  SpreadsheetApp.getUi().alert(
    '✅ Daily Email Trigger Created',
    'Progress summary will be sent to:\n' + RAW_TEST_EMAIL_TARGET + '\n\n' +
    'Every day at approximately 7:30 PM.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Send daily progress email (triggered at 7-8 PM)
 */
function sendDailyProgressEmail_() {
  const stats = getTodayStats_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  // Check if already sent today
  if (stats.lastEmailSent === today) {
    Logger.log('Daily email already sent today');
    return;
  }
  
  const state = getTestPhase1State_();
  
  // Build email content
  let subject = `[CM360 TEST] Daily Progress Report - ${today}`;
  
  let body = `<html><body style="font-family: Arial, sans-serif;">`;
  body += `<h2 style="color: #1a73e8;">📊 Phase 1 Daily Progress Report</h2>`;
  body += `<p><strong>Date:</strong> ${today}</p>`;
  body += `<hr>`;
  
  // Today's work
  body += `<h3>Today's Activity</h3>`;
  body += `<ul>`;
  body += `<li><strong>Gmail Searches:</strong> ${stats.gmailSearches}</li>`;
  body += `<li><strong>CSVs Saved:</strong> ${stats.csvsSaved}</li>`;
  body += `<li><strong>ZIPs Saved:</strong> ${stats.zipsSaved}</li>`;
  body += `<li><strong>Dates Completed:</strong> ${stats.datesCompleted.length}</li>`;
  body += `</ul>`;
  
  // Dates completed today
  if (stats.datesCompleted.length > 0) {
    body += `<h4>Dates Completed Today:</h4>`;
    body += `<ul>`;
    for (const date of stats.datesCompleted) {
      body += `<li>${date}</li>`;
    }
    body += `</ul>`;
  } else {
    body += `<p><em>No dates completed today.</em></p>`;
  }
  
  // Overall progress
  if (state && state.queue) {
    const totalDates = state.queue.length;
    const completed = state.processed || 0;
    const remaining = totalDates - completed;
    const percentComplete = totalDates > 0 ? ((completed / totalDates) * 100).toFixed(1) : 0;
    
    body += `<hr>`;
    body += `<h3>Overall Progress</h3>`;
    body += `<ul>`;
    body += `<li><strong>Total Dates:</strong> ${totalDates}</li>`;
    body += `<li><strong>Completed:</strong> ${completed} (${percentComplete}%)</li>`;
    body += `<li><strong>Remaining:</strong> ${remaining}</li>`;
    body += `<li><strong>Total CSVs:</strong> ${state.totalCSVs || 0}</li>`;
    body += `<li><strong>Total ZIPs:</strong> ${state.totalZIPs || 0}</li>`;
    body += `<li><strong>Total Files:</strong> ${state.totalFiles || 0}</li>`;
    body += `</ul>`;
    
    // ETA calculation
    if (stats.datesCompleted.length > 0 && remaining > 0) {
      const daysElapsed = state.startTime ? Math.max(1, Math.ceil((Date.now() - new Date(state.startTime).getTime()) / (24 * 60 * 60 * 1000))) : 1;
      const ratePerDay = completed / daysElapsed;
      if (ratePerDay > 0) {
        const daysRemaining = Math.ceil(remaining / ratePerDay);
        body += `<p><strong>Estimated Time Remaining:</strong> ~${daysRemaining} days</p>`;
      }
    }
    
    // Date range
    body += `<p><strong>Date Range:</strong> ${state.startDate} to ${state.endDate}</p>`;
    body += `<p><strong>Status:</strong> ${state.status}</p>`;
  }
  
  body += `<hr>`;
  body += `<p style="color: #666; font-size: 12px;">This is an automated report from CM360 TEST Mode Phase 1.</p>`;
  body += `</body></html>`;
  
  // Send email
  try {
    MailApp.sendEmail({
      to: RAW_TEST_EMAIL_TARGET,
      subject: subject,
      htmlBody: body
    });
    
    Logger.log(`✅ Daily email sent to ${RAW_TEST_EMAIL_TARGET}`);
    
    // Mark as sent
    stats.lastEmailSent = today;
    saveTodayStats_(stats);
    
  } catch (e) {
    Logger.log(`❌ Error sending daily email: ${e.message}`);
  }
}

/**
 * Send Phase 1 completion email (immediate)
 */
function sendPhase1CompletionEmail_(state) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  let subject = `[CM360 TEST] ✅ Phase 1 COMPLETE!`;
  
  let body = `<html><body style="font-family: Arial, sans-serif;">`;
  body += `<h2 style="color: #34a853;">✅ Phase 1 Download Complete!</h2>`;
  body += `<p><strong>Completion Time:</strong> ${today}</p>`;
  body += `<hr>`;
  
  body += `<h3>Final Statistics</h3>`;
  body += `<ul>`;
  body += `<li><strong>Dates Processed:</strong> ${state.processed}</li>`;
  body += `<li><strong>Total Files:</strong> ${state.totalFiles}</li>`;
  body += `<li><strong>CSVs:</strong> ${state.totalCSVs}</li>`;
  body += `<li><strong>ZIPs:</strong> ${state.totalZIPs}</li>`;
  body += `<li><strong>Date Range:</strong> ${state.startDate} to ${state.endDate}</li>`;
  body += `</ul>`;
  
  body += `<hr>`;
  body += `<h3>Next Steps</h3>`;
  body += `<ol>`;
  body += `<li>Run <strong>Phase 2: Extract All ZIPs</strong> to unpack the ${state.totalZIPs} ZIP files</li>`;
  body += `<li>Run <strong>Audit TEST Folder</strong> to verify all files</li>`;
  body += `<li>Compare with production folder for validation</li>`;
  body += `</ol>`;
  
  body += `<hr>`;
  body += `<p style="color: #666; font-size: 12px;">The auto-trigger has been stopped. Phase 1 is complete.</p>`;
  body += `</body></html>`;
  
  try {
    MailApp.sendEmail({
      to: RAW_TEST_EMAIL_TARGET,
      subject: subject,
      htmlBody: body
    });
    
    Logger.log(`✅ Completion email sent to ${RAW_TEST_EMAIL_TARGET}`);
  } catch (e) {
    Logger.log(`❌ Error sending completion email: ${e.message}`);
  }
}

// =====================================================================================================================
// ======================================= END RAW DATA GAP FILL (TEST MODE) ========================================
// =====================================================================================================================

// =====================================================================================================================
// ======================================= END RAW DATA GAP FILL SYSTEM ===============================================
// =====================================================================================================================



// ======================================= END HISTORICAL ARCHIVE SYSTEM ===============================================
// =====================================================================================================================

// =====================================================================================================================
// ======================================= SMART VIOLATIONS GAP FILL AUTOMATION ========================================
// =====================================================================================================================

const SMART_PROCESS_TRIGGER_KEY = 'smart_gap_fill_process_trigger';
const SMART_REFRESH_TRIGGER_KEY = 'smart_gap_fill_refresh_trigger';

/**
 * Start Smart Gap Fill Automation
 * - Processes one date every 15 minutes
 * - Refreshes violations audit every 10 minutes
 */
function startSmartGapFillAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  Logger.log('=== Starting Smart Gap Fill Automation ===');
  
  // Check if already running
  const processTrigger = props.getProperty(SMART_PROCESS_TRIGGER_KEY);
  const refreshTrigger = props.getProperty(SMART_REFRESH_TRIGGER_KEY);
  
  if (processTrigger || refreshTrigger) {
    Logger.log('Automation already running, asking user to restart');
    const response = ui.alert(
      '⚠️ Automation Already Running',
      'Smart automation triggers are already active.\n\n' +
      'Do you want to stop and restart them?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      Logger.log('User chose to restart automation');
      stopSmartGapFillAutomation();
    } else {
      Logger.log('User cancelled restart');
      return;
    }
  }
  
  // First, run violations audit to get current status
  Logger.log('Starting Violations Audit scan');
  ss.toast('Scanning Drive for missing violations reports...', '🔄 Initializing', 10);
  
  setupAndRefreshViolationsAudit();
  Logger.log('Violations Audit complete');
  
  // Get missing dates
  const missingDates = getMissingDatesFromAudit_();
  Logger.log(`Found ${missingDates.length} missing dates`);
  
  if (missingDates.length === 0) {
    Logger.log('No gaps found - all reports present');
    ss.toast('All violations reports are present in Drive!', '✅ No Gaps Found', 5);
    return;
  }
  
  // Initialize progress sheet
  Logger.log(`Creating Gap Fill Progress sheet for ${missingDates.length} missing dates`);
  ss.toast('Setting up progress tracking...', '📊 Creating Progress Sheet', 5);
  setupGapFillProgressSheet();
  
  // Verify sheet was created
  const progressSheet = ss.getSheetByName("Gap Fill Progress");
  if (!progressSheet) {
    Logger.log('ERROR: Failed to create Gap Fill Progress sheet');
    ss.toast('Failed to create Gap Fill Progress sheet. Please try again.', '❌ Error', 10);
    return;
  }
  
  const count = initializeGapFillProgress_(missingDates);
  Logger.log(`Initialized ${count} dates in Gap Fill Progress sheet`);
  
  // Initialize state
  const state = {
    queue: missingDates,
    currentDate: null,
    currentStep: null,
    startTime: new Date().toISOString(),
    processed: 0,
    successful: 0,
    failed: 0,
    totalToProcess: missingDates.length,
    completedFiles: [] // Track completed files with details
  };
  saveGapFillState_(state);
  Logger.log(`Gap fill state saved. Queue size: ${state.queue.length}`);
  
  // Create triggers
  try {
    Logger.log('Creating automation triggers...');
    ss.toast('Creating automation triggers...', '⚙️ Setting Up', 3);
    
    // Process trigger: Every 15 minutes (Google only allows 1, 5, 10, 15, 30)
    const processTrig = ScriptApp.newTrigger('smartProcessNextDate')
      .timeBased()
      .everyMinutes(15)
      .create();
    props.setProperty(SMART_PROCESS_TRIGGER_KEY, processTrig.getUniqueId());
    Logger.log(`Created process trigger: ${processTrig.getUniqueId()}`);
    
    // Refresh trigger: Every 10 minutes
    const refreshTrig = ScriptApp.newTrigger('smartRefreshAudit')
      .timeBased()
      .everyMinutes(10)
      .create();
    props.setProperty(SMART_REFRESH_TRIGGER_KEY, refreshTrig.getUniqueId());
    Logger.log(`Created refresh trigger: ${refreshTrig.getUniqueId()}`);
    
    Logger.log('Starting first date processing immediately...');
    ss.toast('Processing first date...', '▶️ Starting', 3);
    
    // Process first date immediately
    smartProcessNextDate();
    
    const estimatedTime = formatEstimatedTime_(count * 15);
    Logger.log(`=== Automation Started Successfully ===`);
    Logger.log(`Missing dates: ${count}`);
    Logger.log(`Estimated completion: ${estimatedTime}`);
    Logger.log(`Process interval: 15 minutes`);
    Logger.log(`Refresh interval: 10 minutes`);
    
    ss.toast(
      `Processing ${count} dates. Est. completion: ${estimatedTime}. Check "Gap Fill Progress" sheet.`,
      '🤖 Automation Started',
      10
    );
    
  } catch (e) {
    Logger.log(`ERROR creating triggers: ${e.toString()}`);
    ss.toast(`Failed to create triggers: ${e.toString()}`, '❌ Error', 15);
  }
}

/**
 * Stop Smart Gap Fill Automation
 */
function stopSmartGapFillAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  
  Logger.log('=== Stopping Smart Gap Fill Automation ===');
  
  let stopped = 0;
  
  // Delete process trigger
  const processId = props.getProperty(SMART_PROCESS_TRIGGER_KEY);
  if (processId) {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getUniqueId() === processId) {
        ScriptApp.deleteTrigger(trigger);
        stopped++;
        Logger.log(`Deleted process trigger: ${processId}`);
      }
    });
    props.deleteProperty(SMART_PROCESS_TRIGGER_KEY);
  }
  
  // Delete refresh trigger
  const refreshId = props.getProperty(SMART_REFRESH_TRIGGER_KEY);
  if (refreshId) {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getUniqueId() === refreshId) {
        ScriptApp.deleteTrigger(trigger);
        stopped++;
        Logger.log(`Deleted refresh trigger: ${refreshId}`);
      }
    });
    props.deleteProperty(SMART_REFRESH_TRIGGER_KEY);
  }
  
  if (stopped > 0) {
    const state = getGapFillState_();
    const remaining = state ? state.queue.length : 0;
    
    Logger.log(`Automation stopped. Processed: ${state ? state.processed : 0}, Successful: ${state ? state.successful : 0}, Failed: ${state ? state.failed : 0}, Remaining: ${remaining}`);
    
    ss.toast(
      `Processed: ${state ? state.processed : 0} | Successful: ${state ? state.successful : 0} | Failed: ${state ? state.failed : 0} | Remaining: ${remaining}`,
      '🛑 Automation Stopped',
      10
    );
  } else {
    Logger.log('No active automation triggers found');
    ss.toast('No smart automation triggers were found.', 'ℹ️ No Active Automation', 5);
  }
}

/**
 * Smart process next date (triggered every 20 minutes)
 */
function smartProcessNextDate() {
  const startTime = Date.now();
  const state = getGapFillState_();
  
  if (!state || !state.queue || state.queue.length === 0) {
    Logger.log('✅ Gap fill complete - all dates processed');
    stopSmartGapFillAutomation();
    sendCompletionNotification_();
    return;
  }
  
  // Process ONE date
  const dateStr = state.queue[0];
  state.currentDate = dateStr;
  
  Logger.log(`🔄 Processing date: ${dateStr}`);
  updateGapFillProgress_(dateStr, '🔄 Running Time Machine...', '', '');
  
  try {
    // Run Time Machine to generate violations report
    // This will: Import raw data from Gmail → Run QA → Save violations to Drive
    const tmResult = runTimeMachineForDate_(dateStr);
    
    if (tmResult.success) {
      updateGapFillProgress_(dateStr, '✅ Complete', '', tmResult.filename);
      
      // Track completed file details
      if (!state.completedFiles) state.completedFiles = [];
      state.completedFiles.push({
        date: dateStr,
        filename: tmResult.filename,
        violationCount: tmResult.violationCount,
        folderPath: tmResult.folderPath,
        fileUrl: tmResult.fileUrl
      });
      
      state.successful++;
    } else {
      updateGapFillProgress_(dateStr, '❌ Failed', tmResult.error, '');
      state.failed++;
    }
    
    state.processed++;
    state.queue.shift();
    saveGapFillState_(state);
    
    const elapsed = (Date.now() - startTime) / 1000;
    Logger.log(`✅ Completed ${dateStr} in ${elapsed.toFixed(1)}s. Remaining: ${state.queue.length}`);
    
  } catch (e) {
    Logger.log(`❌ Error processing ${dateStr}: ${e.toString()}`);
    updateGapFillProgress_(dateStr, '❌ Error', e.toString(), '');
    state.failed++;
    state.processed++;
    state.queue.shift();
    saveGapFillState_(state);
  }
}

/**
 * Smart refresh audit (triggered every 10 minutes)
 */
function smartRefreshAudit() {
  Logger.log('🔄 Smart refresh: Updating Violations Audit...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Violations Audit");
    
    if (sheet) {
      // Only refresh if the audit sheet exists
      setupAndRefreshViolationsAudit();
      Logger.log('✅ Violations Audit refreshed');
    }
  } catch (e) {
    Logger.log(`⚠️ Could not refresh audit: ${e.toString()}`);
  }
}

/**
 * Send completion notification
 */
function sendCompletionNotification_() {
  const state = getGapFillState_();
  if (!state) return;
  
  const subject = '✅ Violations Gap Fill Complete';
  
  let body = `<html><body style="font-family: Arial, sans-serif;">`;
  body += `<h2>✅ Gap Fill Automation Complete</h2>`;
  body += `<p><strong>Total Processed:</strong> ${state.processed}</p>`;
  body += `<p><strong>Successful:</strong> ${state.successful}</p>`;
  body += `<p><strong>Failed:</strong> ${state.failed}</p>`;
  body += `<p><strong>Duration:</strong> ${calculateDuration_(state.startTime)}</p>`;
  body += `<hr>`;
  
  // List successfully created files
  if (state.completedFiles && state.completedFiles.length > 0) {
    body += `<h3>📄 Successfully Created Files (${state.completedFiles.length}):</h3>`;
    body += `<table style="border-collapse: collapse; width: 100%;">`;
    body += `<tr style="background-color: #f0f0f0; font-weight: bold;">`;
    body += `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Date</th>`;
    body += `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">File</th>`;
    body += `<th style="border: 1px solid #ddd; padding: 8px; text-align: center;">Placements</th>`;
    body += `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Folder</th>`;
    body += `</tr>`;
    
    for (const file of state.completedFiles) {
      body += `<tr>`;
      body += `<td style="border: 1px solid #ddd; padding: 8px;">${file.date}</td>`;
      body += `<td style="border: 1px solid #ddd; padding: 8px;"><a href="${file.fileUrl}">${file.filename}</a></td>`;
      body += `<td style="border: 1px solid #ddd; padding: 8px; text-align: center; font-weight: bold;">${file.violationCount}</td>`;
      body += `<td style="border: 1px solid #ddd; padding: 8px; font-size: 11px;">${file.folderPath}</td>`;
      body += `</tr>`;
    }
    
    body += `</table>`;
    body += `<p style="color: #666; font-size: 12px; margin-top: 10px;">`;
    body += `<strong>Total Violations Processed:</strong> ${state.completedFiles.reduce((sum, f) => sum + f.violationCount, 0)} placements`;
    body += `</p>`;
  }
  
  // List failed dates
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressSheet = ss.getSheetByName("Gap Fill Progress");
  if (progressSheet && state.failed > 0) {
    body += `<hr>`;
    body += `<h3>❌ Failed Dates (${state.failed}):</h3>`;
    body += `<ul>`;
    
    const data = progressSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][1]);
      if (status.includes('❌')) {
        const date = data[i][0];
        const error = data[i][4] || 'Unknown error';
        body += `<li><strong>${date}:</strong> ${error}</li>`;
      }
    }
    
    body += `</ul>`;
  }
  
  body += `<hr>`;
  body += `<p>Check the "Gap Fill Progress" and "Violations Audit" sheets for full details.</p>`;
  body += `<p><a href="https://drive.google.com/drive/folders/1lJm0K1LLo9ez29AcKCc4qtIbBC2uK3a9">📁 View All Reports in Drive</a></p>`;
  body += `</body></html>`;
  
  try {
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: subject,
      htmlBody: body
    });
    Logger.log(`✅ Completion email sent with ${state.completedFiles ? state.completedFiles.length : 0} files listed`);
  } catch (e) {
    Logger.log(`Could not send completion email: ${e.toString()}`);
  }
}

/**
 * Format estimated time
 */
function formatEstimatedTime_(minutes) {
  if (minutes < 60) {
    return `${Math.round(minutes)} minutes`;
  }
  const hours = Math.floor(minutes / 60);
  const mins = Math.round(minutes % 60);
  return `${hours}h ${mins}m`;
}

/**
 * Calculate duration
 */
function calculateDuration_(startISO) {
  const start = new Date(startISO);
  const end = new Date();
  const diffMs = end - start;
  const diffMins = Math.round(diffMs / 60000);
  return formatEstimatedTime_(diffMins);
}

// =====================================================================================================================
// ======================================= END SMART VIOLATIONS GAP FILL AUTOMATION ===================================
// =====================================================================================================================

// =====================================================================================================================
// ======================================= SMART RAW DATA GAP FILL AUTOMATION ==========================================
// =====================================================================================================================

const SMART_RAW_PROCESS_TRIGGER_KEY = 'smart_raw_gap_fill_process_trigger';
const SMART_RAW_REFRESH_TRIGGER_KEY = 'smart_raw_gap_fill_refresh_trigger';

/**
 * Start Smart Raw Data Gap Fill Automation
 * - Processes raw data downloads every 15 minutes
 * - Refreshes raw data audit every 10 minutes
 */
function startSmartRawDataAutomation() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  // Check if already running
  const processTrigger = props.getProperty(SMART_RAW_PROCESS_TRIGGER_KEY);
  const refreshTrigger = props.getProperty(SMART_RAW_REFRESH_TRIGGER_KEY);
  
  if (processTrigger || refreshTrigger) {
    const response = ui.alert(
      '⚠️ Automation Already Running',
      'Smart raw data automation triggers are already active.\n\n' +
      'Do you want to stop and restart them?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      stopSmartRawDataAutomation();
    } else {
      return;
    }
  }
  
  // First, run raw data audit to get current status
  ui.alert(
    '🔄 Initializing',
    'Scanning Drive for missing raw data files...\n\nThis will take a moment.',
    ui.ButtonSet.OK
  );
  
  setupAndRefreshRawDataAudit();
  
  // Get missing dates from Audit Dashboard
  const missingDates = getMissingRawDataDatesFromAudit_();
  
  if (missingDates.length === 0) {
    ui.alert(
      '✅ No Gaps Found',
      'All raw data files are present in Drive!\n\nNo automation needed.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Initialize state
  const state = {
    missingDates: missingDates,
    currentIndex: 0,
    startTime: new Date().toISOString(),
    processed: 0,
    successful: 0,
    failed: 0,
    totalToProcess: missingDates.length
  };
  props.setProperty('smart_raw_state', JSON.stringify(state));
  
  // Create triggers
  try {
    // Process trigger: Every 15 minutes (faster since it's just Gmail downloads)
    const processTrig = ScriptApp.newTrigger('smartProcessRawDataBatch')
      .timeBased()
      .everyMinutes(15)
      .create();
    props.setProperty(SMART_RAW_PROCESS_TRIGGER_KEY, processTrig.getUniqueId());
    
    // Refresh trigger: Every 10 minutes
    const refreshTrig = ScriptApp.newTrigger('smartRefreshRawDataAudit')
      .timeBased()
      .everyMinutes(10)
      .create();
    props.setProperty(SMART_RAW_REFRESH_TRIGGER_KEY, refreshTrig.getUniqueId());
    
    // Process first batch immediately
    smartProcessRawDataBatch();
    
    ui.alert(
      '🤖 Smart Raw Data Automation Started',
      `Found ${missingDates.length} missing date(s).\n\n` +
      `⚙️ AUTOMATION ACTIVE:\n` +
      `• Processing Gmail downloads every 15 minutes\n` +
      `• Refreshing audit every 10 minutes\n\n` +
      `📊 Progress tracked in execution logs\n` +
      `📋 Results updated in "Audit Dashboard" sheet\n\n` +
      `Estimated completion: ${formatEstimatedTime_(missingDates.length * 15)}\n\n` +
      `Use "Stop Smart Automation" to cancel at any time.`,
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    ui.alert('❌ Error', `Failed to create triggers: ${e.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Stop Smart Raw Data Gap Fill Automation
 */
function stopSmartRawDataAutomation() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  let stopped = 0;
  
  // Delete process trigger
  const processId = props.getProperty(SMART_RAW_PROCESS_TRIGGER_KEY);
  if (processId) {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getUniqueId() === processId) {
        ScriptApp.deleteTrigger(trigger);
        stopped++;
      }
    });
    props.deleteProperty(SMART_RAW_PROCESS_TRIGGER_KEY);
  }
  
  // Delete refresh trigger
  const refreshId = props.getProperty(SMART_RAW_REFRESH_TRIGGER_KEY);
  if (refreshId) {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getUniqueId() === refreshId) {
        ScriptApp.deleteTrigger(trigger);
        stopped++;
      }
    });
    props.deleteProperty(SMART_RAW_REFRESH_TRIGGER_KEY);
  }
  
  if (stopped > 0) {
    const stateJson = props.getProperty('smart_raw_state');
    const state = stateJson ? JSON.parse(stateJson) : null;
    
    ui.alert(
      '🛑 Raw Data Automation Stopped',
      `Smart raw data automation has been stopped.\n\n` +
      `Processed: ${state ? state.processed : 0}\n` +
      `Successful: ${state ? state.successful : 0}\n` +
      `Failed: ${state ? state.failed : 0}\n\n` +
      `You can restart automation at any time.`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      'ℹ️ No Active Automation',
      'No smart raw data automation triggers were found.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Smart process raw data batch (triggered every 15 minutes)
 */
function smartProcessRawDataBatch() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('smart_raw_state');
  
  if (!stateJson) {
    Logger.log('✅ Raw data gap fill complete - no state found');
    stopSmartRawDataAutomation();
    return;
  }
  
  const state = JSON.parse(stateJson);
  
  if (state.currentIndex >= state.missingDates.length) {
    Logger.log('✅ All raw data dates processed');
    stopSmartRawDataAutomation();
    sendRawDataCompletionNotification_(state);
    return;
  }
  
  const dateStr = state.missingDates[state.currentIndex];
  Logger.log(`🔄 Processing raw data for date: ${dateStr}`);
  
  try {
    // Download raw data from Gmail for this date
    const result = downloadRawDataForDate_(dateStr);
    
    if (result.success) {
      Logger.log(`✅ Downloaded ${result.filesProcessed} files for ${dateStr}`);
      state.successful++;
    } else {
      Logger.log(`❌ Failed to download for ${dateStr}: ${result.error}`);
      state.failed++;
    }
    
    state.processed++;
    state.currentIndex++;
    props.setProperty('smart_raw_state', JSON.stringify(state));
    
    Logger.log(`Progress: ${state.processed}/${state.totalToProcess} dates`);
    
  } catch (e) {
    Logger.log(`❌ Error processing ${dateStr}: ${e.toString()}`);
    state.failed++;
    state.processed++;
    state.currentIndex++;
    props.setProperty('smart_raw_state', JSON.stringify(state));
  }
}

/**
 * Smart refresh raw data audit (triggered every 10 minutes)
 */
function smartRefreshRawDataAudit() {
  Logger.log('🔄 Smart refresh: Updating Raw Data Audit...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Audit Dashboard");
    
    if (sheet) {
      refreshAuditDashboardChunked();
      Logger.log('✅ Raw Data Audit refreshed');
    }
  } catch (e) {
    Logger.log(`⚠️ Could not refresh audit: ${e.toString()}`);
  }
}

/**
 * Get missing raw data dates from Audit Dashboard
 */
function getMissingRawDataDatesFromAudit_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Audit Dashboard");
  
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const missingDates = [];
  
  for (const row of data) {
    const dateStr = row[0];
    const status = row[1];
    
    if (status === '❌ MISSING' && dateStr) {
      missingDates.push(String(dateStr));
    }
  }
  
  return missingDates;
}

/**
 * Send raw data completion notification
 */
function sendRawDataCompletionNotification_(state) {
  const subject = '✅ Raw Data Gap Fill Complete';
  const body = `
    <h2>Raw Data Gap Fill Automation Complete</h2>
    <p><strong>Total Processed:</strong> ${state.processed}</p>
    <p><strong>Successful:</strong> ${state.successful}</p>
    <p><strong>Failed:</strong> ${state.failed}</p>
    <p><strong>Duration:</strong> ${calculateDuration_(state.startTime)}</p>
    <hr>
    <p>Check the "Audit Dashboard" sheet for details.</p>
  `;
  
  try {
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: subject,
      htmlBody: body
    });
  } catch (e) {
    Logger.log(`Could not send completion email: ${e.toString()}`);
  }
}

// =====================================================================================================================
// ======================================= END SMART RAW DATA GAP FILL AUTOMATION ======================================
// =====================================================================================================================

// =====================================================================================================================
// ======================================= SMART TEST MODE AUTOMATION ==================================================
// =====================================================================================================================

const SMART_TEST_COMPLETE_TRIGGER_KEY = 'smart_test_complete_trigger';

/**
 * Start Complete TEST Automation - Runs both Phase 1 and Phase 2 automatically
 */
function startCompleteTestAutomation() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  // Check if already running
  const existingTrigger = props.getProperty(SMART_TEST_COMPLETE_TRIGGER_KEY);
  
  if (existingTrigger) {
    const response = ui.alert(
      '⚠️ Automation Already Running',
      'Complete TEST automation is already active.\n\n' +
      'Do you want to stop and restart it?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      stopCompleteTestAutomation();
    } else {
      return;
    }
  }
  
  const response = ui.alert(
    '🤖 Start Complete TEST Automation',
    'This will automatically:\n\n' +
    '1. Run Phase 1 (Download all attachments)\n' +
    '2. Wait for Phase 1 to complete\n' +
    '3. Run Phase 2 (Extract all ZIPs)\n' +
    '4. Send completion email\n\n' +
    'The system will handle everything automatically.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // Initialize complete test state
  const state = {
    phase: 1,
    startTime: new Date().toISOString(),
    phase1Complete: false,
    phase2Complete: false
  };
  props.setProperty('smart_test_complete_state', JSON.stringify(state));
  
  // Create monitoring trigger (checks every 15 minutes)
  const trigger = ScriptApp.newTrigger('monitorCompleteTestAutomation')
    .timeBased()
    .everyMinutes(15)
    .create();
  props.setProperty(SMART_TEST_COMPLETE_TRIGGER_KEY, trigger.getUniqueId());
  
  // Start Phase 1
  startTestPhase1Download();
  createTestPhase1Trigger();
  
  ui.alert(
    '🤖 Complete TEST Automation Started',
    'Phase 1 (Download) has been started.\n\n' +
    'The system will automatically:\n' +
    '• Monitor Phase 1 progress\n' +
    '• Start Phase 2 when Phase 1 completes\n' +
    '• Send completion email when done\n\n' +
    'You can check progress in the TEST sheets or logs.',
    ui.ButtonSet.OK
  );
}

/**
 * Stop Complete TEST Automation
 */
function stopCompleteTestAutomation() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  // Delete monitoring trigger
  const triggerId = props.getProperty(SMART_TEST_COMPLETE_TRIGGER_KEY);
  if (triggerId) {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    props.deleteProperty(SMART_TEST_COMPLETE_TRIGGER_KEY);
  }
  
  // Also stop phase-specific triggers
  stopAllTestTriggers();
  
  ui.alert(
    '🛑 Complete TEST Automation Stopped',
    'All TEST automation triggers have been stopped.\n\n' +
    'You can restart automation at any time.',
    ui.ButtonSet.OK
  );
}

/**
 * Monitor complete test automation progress
 */
function monitorCompleteTestAutomation() {
  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('smart_test_complete_state');
  
  if (!stateJson) {
    Logger.log('No complete test state found');
    stopCompleteTestAutomation();
    return;
  }
  
  const state = JSON.parse(stateJson);
  
  // Check Phase 1 status
  if (state.phase === 1 && !state.phase1Complete) {
    const phase1StateJson = props.getProperty(RAW_TEST_PHASE1_STATE_KEY);
    if (phase1StateJson) {
      const phase1State = JSON.parse(phase1StateJson);
      
      if (phase1State.completed) {
        Logger.log('✅ Phase 1 complete, starting Phase 2');
        state.phase = 2;
        state.phase1Complete = true;
        props.setProperty('smart_test_complete_state', JSON.stringify(state));
        
        // Stop Phase 1 trigger
        stopTestPhase1Trigger_();
        
        // Start Phase 2
        Utilities.sleep(2000);
        startTestPhase2Extraction();
        createTestPhase2Trigger();
      }
    }
  }
  
  // Check Phase 2 status
  if (state.phase === 2 && !state.phase2Complete) {
    const phase2StateJson = props.getProperty(RAW_TEST_PHASE2_STATE_KEY);
    if (phase2StateJson) {
      const phase2State = JSON.parse(phase2StateJson);
      
      if (phase2State.completed) {
        Logger.log('✅ Phase 2 complete, TEST automation finished');
        state.phase2Complete = true;
        props.setProperty('smart_test_complete_state', JSON.stringify(state));
        
        // Stop Phase 2 trigger
        stopTestPhase2Trigger_();
        
        // Send completion notification
        sendCompleteTestNotification_(state);
        
        // Stop monitoring
        stopCompleteTestAutomation();
      }
    }
  }
}

/**
 * Send complete test automation notification
 */
function sendCompleteTestNotification_(state) {
  const subject = '✅ Complete TEST Automation Finished';
  const body = `
    <h2>Complete TEST Mode Automation Finished</h2>
    <p>Both Phase 1 (Download) and Phase 2 (Extraction) are complete!</p>
    <hr>
    <p><strong>Phase 1:</strong> ✅ Complete</p>
    <p><strong>Phase 2:</strong> ✅ Complete</p>
    <p><strong>Total Duration:</strong> ${calculateDuration_(state.startTime)}</p>
    <hr>
    <h3>Next Steps:</h3>
    <ol>
      <li>Run "Audit TEST Folder" to verify all files</li>
      <li>Run "Cleanup & Verify" to finalize</li>
    </ol>
  `;
  
  try {
    MailApp.sendEmail({
      to: RAW_TEST_EMAIL_TARGET,
      subject: subject,
      htmlBody: body
    });
  } catch (e) {
    Logger.log(`Could not send completion email: ${e.toString()}`);
  }
}

/**
 * Stop Phase 1 trigger
 */
function stopTestPhase1Trigger_() {
  const props = PropertiesService.getScriptProperties();
  const triggerId = props.getProperty(RAW_TEST_PHASE1_TRIGGER_KEY);
  if (triggerId) {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    props.deleteProperty(RAW_TEST_PHASE1_TRIGGER_KEY);
  }
}

/**
 * Stop Phase 2 trigger
 */
function stopTestPhase2Trigger_() {
  const props = PropertiesService.getScriptProperties();
  const triggerId = props.getProperty(RAW_TEST_PHASE2_TRIGGER_KEY);
  if (triggerId) {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    props.deleteProperty(RAW_TEST_PHASE2_TRIGGER_KEY);
  }
}

// =====================================================================================================================
// ======================================= END SMART TEST MODE AUTOMATION ==============================================
// =====================================================================================================================


