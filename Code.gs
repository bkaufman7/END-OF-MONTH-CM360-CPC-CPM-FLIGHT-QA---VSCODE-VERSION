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
    .addItem("Run It All", "runItAll")
    .addItem("Pull Data", "importDCMReports")
    .addItem("Run QA Only", "runQAOnly")
    .addItem("Send Email Only", "sendEmailSummary")
    .addSeparator()
    .addItem("Authorize Email (one-time)", "authorizeMail_")         // <-- add this
    .addItem("Create Daily Email Trigger (9am)", "createDailyEmailTrigger") // <-- and this
    .addSeparator()
    .addItem("Clear Violations", "clearViolations")
    .addToUi();
}



// ---------------------
// one-time MailApp authorization helper
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
// Create cascading triggers for better timeout management
// ---------------------
function createCascadingTriggers() {
  // Clear any existing cascade triggers first
  clearAllCascadeTriggers_();
  
  // Trigger 1: Data Ingestion at 1:15 AM (after files arrive 12:42-12:56 AM)
  ScriptApp.newTrigger('runDataIngestion')
    .timeBased()
    .atHour(1)
    .nearMinute(15)
    .everyDays(1)
    .create();
    
  Logger.log('✅ Created cascading triggers: Data Ingestion (1:15 AM) → QA Processing → Email Reporting');
}

// Legacy function - kept for backward compatibility
function createDailyEmailTrigger() {
  createCascadingTriggers();
}



// ---------------------
// clearViolations
// ---------------------
function clearViolations() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Violations");
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

// ====== CASCADING TRIGGER SYSTEM ======

const CASCADE_STATE_KEY = 'cascade_progress_v1';
const CASCADE_TRIGGER_PREFIX = 'cascade_';

// ====== Chunked QA execution control ======
const QA_CHUNK_ROWS = 3500;
const QA_TIME_BUDGET_MS = 4.2 * 60 * 1000;
const QA_STATE_KEY = 'qa_progress_v2';

const QA_TRIGGER_KEY = 'qa_chunk_trigger_id';
const QA_LOCK_KEY = 'qa_chunk_lock';

function getScriptProps_() { return PropertiesService.getScriptProperties(); }

// --- CASCADE TRIGGER MANAGEMENT ---
function clearAllCascadeTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    const funcName = trigger.getHandlerFunction();
    if (funcName.startsWith('run') && (funcName.includes('Ingestion') || funcName.includes('QAProcessing') || funcName.includes('EmailReporting'))) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('🗑️ Removed trigger: ' + funcName);
    }
  });
}

function scheduleCascadeTrigger_(functionName, delayMinutes) {
  // Clear any existing trigger for this function
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger
  const newTrigger = ScriptApp.newTrigger(functionName)
    .timeBased()
    .after(delayMinutes * 60 * 1000)
    .create();
    
  Logger.log('⏰ Scheduled ' + functionName + ' in ' + delayMinutes + ' minutes');
  return newTrigger.getUniqueId();
}

function setCascadeState_(step, status, data) {
  const state = {
    currentStep: step,
    status: status,
    timestamp: new Date().toISOString(),
    data: data || {}
  };
  PropertiesService.getDocumentProperties().setProperty(CASCADE_STATE_KEY, JSON.stringify(state));
}

function getCascadeState_() {
  const raw = PropertiesService.getDocumentProperties().getProperty(CASCADE_STATE_KEY);
  return raw ? JSON.parse(raw) : null;
}

// --- CASCADE STEP 1: DATA INGESTION ---
function runDataIngestion() {
  try {
    setCascadeState_('ingestion', 'started');
    Logger.log('🚀 CASCADE STEP 1: Data Ingestion - START');
    
    // Trim sheets and import data
    trimAllSheetsToData_();
    importDCMReports();
    
    setCascadeState_('ingestion', 'completed');
    Logger.log('✅ CASCADE STEP 1: Data Ingestion - COMPLETED');
    
    // Schedule next step
    scheduleCascadeTrigger_('runQAProcessing', 15); // 15 minutes later
    
  } catch (error) {
    setCascadeState_('ingestion', 'failed', { error: error.toString() });
    Logger.log('❌ CASCADE STEP 1: Data Ingestion - FAILED: ' + error.toString());
    
    // Still proceed to QA step in case of partial success
    scheduleCascadeTrigger_('runQAProcessing', 20); // 20 minutes later with extra buffer
  }
}

// --- CASCADE STEP 2: QA PROCESSING ---
function runQAProcessing() {
  try {
    setCascadeState_('qa', 'started');
    Logger.log('🚀 CASCADE STEP 2: QA Processing - START');
    
    // Run QA (this will handle chunking automatically)
    runQAOnly();
    
    // Send performance alerts if pre-15th
    sendPerformanceSpikeAlertIfPre15();
    
    // Check if QA is truly complete
    const qaState = getQAState_();
    if (qaState && qaState.session) {
      // QA is still running in chunks, let it complete first
      Logger.log('⏳ CASCADE STEP 2: QA still chunking, will wait for completion');
      setCascadeState_('qa', 'chunking');
      // Don't schedule email yet - QA chunking will handle final step
      return;
    }
    
    setCascadeState_('qa', 'completed');
    Logger.log('✅ CASCADE STEP 2: QA Processing - COMPLETED');
    
    // Schedule final step
    scheduleCascadeTrigger_('runEmailReporting', 15); // 15 minutes later
    
  } catch (error) {
    setCascadeState_('qa', 'failed', { error: error.toString() });
    Logger.log('❌ CASCADE STEP 2: QA Processing - FAILED: ' + error.toString());
    
    // Still proceed to email step for any available data
    scheduleCascadeTrigger_('runEmailReporting', 20);
  }
}

// --- CASCADE STEP 3: EMAIL REPORTING ---
function runEmailReporting() {
  try {
    setCascadeState_('email', 'started');
    Logger.log('🚀 CASCADE STEP 3: Email Reporting - START');
    
    // Send email summary (has built-in >15th filter)
    sendEmailSummary();
    
    setCascadeState_('email', 'completed');
    Logger.log('✅ CASCADE STEP 3: Email Reporting - COMPLETED');
    Logger.log('🏁 CASCADE COMPLETE: All steps finished');
    
    // Clear cascade state
    PropertiesService.getDocumentProperties().deleteProperty(CASCADE_STATE_KEY);
    
  } catch (error) {
    setCascadeState_('email', 'failed', { error: error.toString() });
    Logger.log('❌ CASCADE STEP 3: Email Reporting - FAILED: ' + error.toString());
  }
}

// Enhanced QA chunking with cascade awareness
function scheduleNextQAChunk_(minutesFromNow) {
  minutesFromNow = Math.max(1, Math.min(10, Math.floor(minutesFromNow || 1)));
  const props = getScriptProps_();

  const existingId = props.getProperty(QA_TRIGGER_KEY);
  if (existingId) {
    const stillThere = ScriptApp.getProjectTriggers().some(function(t){ return t.getUniqueId() === existingId; });
    if (stillThere) return;
    props.deleteProperty(QA_TRIGGER_KEY);
  }

  const trig = ScriptApp
    .newTrigger('runQAChunkAndCheckComplete')  // Modified to check cascade state
    .timeBased()
    .after(minutesFromNow * 60 * 1000)
    .create();

  props.setProperty(QA_TRIGGER_KEY, trig.getUniqueId());
}

// Enhanced QA chunk runner that integrates with cascade
function runQAChunkAndCheckComplete() {
  runQAOnly(); // Run the QA chunk
  
  // Check if QA is complete
  const qaState = getQAState_();
  if (!qaState || !qaState.session) {
    // QA is complete, check if we're in a cascade
    const cascadeState = getCascadeState_();
    if (cascadeState && cascadeState.currentStep === 'qa' && cascadeState.status === 'chunking') {
      Logger.log('🔄 QA chunking complete, resuming cascade');
      setCascadeState_('qa', 'completed');
      scheduleCascadeTrigger_('runEmailReporting', 5); // Quick transition to email
    }
  }
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
    Object.values(ignoreMap).forEach(function(o){
      sheet.getRange(o.row, 2).setValue(o.set.size);
    });
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

  const MATCH_TEXT = "🟨 PERFORMANCE: CTR ≥ 90% & CPM ≥ $10";
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
      const trimmedCampaign  = camp.length > 20 ? camp.substring(0, 20) + "…" : camp;
      const trimmedPlacement = plc.length > 20 ? plc.substring(0, 20) + "…" : plc;

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
  const subject = 'ALERT – PERFORMANCE (pre-monthly-summary) – ' + todayStr + ' – ' + candidateRows.length + ' changed/new row(s)';

  uniqueEmails.forEach(function(addr){
    try {
      MailApp.sendEmail({ to: addr, subject: subject, htmlBody: table });
      Utilities.sleep(300);
    } catch (err) {
      Logger.log('❌ Failed to email ' + addr + ': ' + err);
    }
  });

  compactPerfAlertCache_(35);
}

// --- CONTINUATION OF FULL IMPLEMENTATION ---
// Note: This is a partial implementation - copy your complete original code here
// including all the remaining functions:
// - runQAOnly (chunked QA processing)
// - sendEmailSummary (email reporting system)  
// - Low-priority classification patterns
// - Owner resolution functions
// - Violation tracking functions
// - trimAllSheetsToData_
// - All helper functions

// Placeholder for remaining implementation
function runQAOnly() {
  Logger.log('⚠️ PLACEHOLDER: runQAOnly - Replace with your complete implementation');
  // TODO: Add your complete runQAOnly implementation here
}

function sendEmailSummary() {
  Logger.log('⚠️ PLACEHOLDER: sendEmailSummary - Replace with your complete implementation');
  // TODO: Add your complete sendEmailSummary implementation here
}

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