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
    .addItem("Authorize Email (one-time)", "authorizeMail_")
    .addItem("Create Daily Email Trigger (9am)", "createDailyEmailTrigger")
    .addSeparator()
    .addItem("Clear Violations", "clearViolations")
    .addSeparator()
    .addSubMenu(ui.createMenu("📊 V2 Dashboard (BETA)")
      .addItem("🎯 Generate V2 Dashboard", "generateViolationsV2Dashboard")
      .addItem("💾 Export V2 to Drive", "exportV2ToDrive")
      .addItem("📊 Monthly Summary Report", "generateMonthlySummaryReport")
      .addItem("📈 Month-over-Month Analysis", "runMonthOverMonthAnalysis")
      .addItem("💰 Calculate Financial Impact", "displayFinancialImpact"))
    .addSeparator()
    .addSubMenu(ui.createMenu("📁 Historical Archive")
      .addItem("📁 Archive All (April-Nov 2025)", "archiveAllHistoricalReports")
      .addItem("📅 Archive Single Month", "archiveSingleMonth")
      .addItem("📊 View Archive Progress", "viewArchiveProgress")
      .addItem("🔄 Resume Archive", "resumeArchive"))
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
      Utilities.sleep(500);
    } catch (err) {
      Logger.log('❌ Failed to email ' + addr + ': ' + err);
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
// Owner/Rep mapping helpers + lookup from "Networks" (prefer OPS in P–S)
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
 * Low-Priority Scoring — Lightweight (NO sheets/logging)
 *******************************************************/

// Keep these defaults (same signal quality, no sheet I/O)
const X_CH = "[x×✕]";
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
    .replace(/[×✕]/g, 'x')
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
    // Don’t LP-tag rows where both metrics present (or pathological both+clicks>impr)
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

  // If Mixed, we’d subtract negatives; for single-metric add a tiny boost when size present
  if (gating !== 'Mixed') {
    var sizeRgx = _negCompiled[0].re;
    if (sizeRgx && sizeRgx.test(s)) {
      pos += 15; // helps 1x1 & obvious “pixel-ish” names
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
  return 'Low Priority — ' + topCat + ' (' + band + ')';
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
    const raw = ss.getSheetByName("Raw Data");
    const out = ss.getSheetByName("Violations");
    if (!raw || !out) return;

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

    // —— Tweak these constants in your file (outside this function) ——
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

      // 🟥 BILLING
      if (pe < firstOfMonth && clk > imp) {
        issueTypes.push("🟥 BILLING: Expired CPC Risk");
        details.push("Ended " + pe.toDateString() + " with clicks (" + clk + ") > impressions (" + imp + ")");
        risk = "🚨 Expired Risk";
      } else if (pe < rd && clk > imp) {
        issueTypes.push("🟥 BILLING: Recently Expired CPC Risk");
        details.push("Ended " + pe.toDateString() + " and still has clicks > impressions");
        risk = "⚠️ Expired This Month";
      } else if (rd <= pe && clk > imp && cpc > 10) {
        issueTypes.push("🟥 BILLING: Active CPC Billing Risk");
        details.push("Active: clicks (" + clk + ") > impressions (" + imp + "), $CPC = $" + cpc.toFixed(2));
        risk = "⚠️ Active CPC Risk";
      }

      // 🟦 DELIVERY
      if (pe < firstOfMonth && rd >= firstOfMonth && (imp > 0 || clk > 0)) {
        issueTypes.push("🟦 DELIVERY: Post-Flight Activity");
        details.push("Ended " + pe.toDateString() + " but has " + imp + " impressions and " + clk + " clicks");
      }

      // 🟨 PERFORMANCE
      if (ctr >= 90 && cpm >= 10) {
        issueTypes.push("🟨 PERFORMANCE: CTR ≥ 90% & CPM ≥ $10");
        details.push("CTR = " + ctr.toFixed(2) + "%, $CPM = $" + cpm.toFixed(2));
      }

      // 🟩 COST
      let isCPMOnly = false;
      let isCPCOnly = false;
      if (cpc > 0 && cpm === 0 && cpc > 10) {
        issueTypes.push("🟩 COST: CPC Only > $10");
        details.push("No CPM spend, $CPC = $" + cpc.toFixed(2));
        if (imp === 0 && clk > 0) isCPCOnly = true;
      }
      if (cpm > 0 && cpc === 0 && cpm > 10) {
        issueTypes.push("🟩 COST: CPM Only > $10");
        details.push("No CPC spend, $CPM = $" + cpm.toFixed(2));
        if (imp > 0 && clk === 0) isCPMOnly = true;
      }
      if (cpc > 0 && cpm > 0 && clk > imp && cpc > 10) {
        issueTypes.push("🟩 COST: CPC+CPM Clicks > Impr & CPC > $10");
        details.push("Clicks > impressions with both CPC and CPM charges (CPC = $" + cpc.toFixed(2) + ")");
      }

      // --- Low-priority tagging via scorer (gating-aware) — no sheet writes ---
      const bothMetricsPresent = imp > 0 && clk > 0;
      const clicksExceedImprWithBoth = bothMetricsPresent && (clk > imp);
      const gating = (imp > 0 && clk === 0) ? 'CPM-only' :
                     (imp === 0 && clk > 0) ? 'CPC-only' : 'Mixed';

      if (!bothMetricsPresent && !clicksExceedImprWithBoth) {
        const placement = row[m["Placement"]];
        const rowIdOrIndex = String(row[m["Placement ID"]] || (r + 1));
        const lpDescriptor = scoreAndLabelLowPriority_(placement, clk, imp, rowIdOrIndex, gating);
        if (lpDescriptor) {
          issueTypes.push("🟩 COST: (Low Priority) " + lpDescriptor.replace(/^Low Priority —\s*/, ""));
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
      Logger.log("✅ runQAOnly complete. Processed all " + totalRows + " data rows.");
    } else {
      saveQAState_(state);
      Logger.log("⏳ runQAOnly partial: processed " + processed + " rows this run. Next row index: "
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
// sendEmailSummary (size-safe) — UPDATED with extra buckets
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
      violationCounts[id] = { "🟥 BILLING": 0, "🟦 DELIVERY": 0, "🟨 PERFORMANCE": 0, "🟩 COST": 0 };
    }
    types.forEach(function(t){
      if (t.startsWith("🟥")) violationCounts[id]["🟥 BILLING"]++;
      if (t.startsWith("🟦")) violationCounts[id]["🟦 DELIVERY"]++;
      if (t.startsWith("🟨")) violationCounts[id]["🟨 PERFORMANCE"]++;
      if (t.startsWith("🟩")) violationCounts[id]["🟩 COST"]++;
    });
  });

  // --- Network summary table ---
  let networkSummary =
      '<p><b>Network-Level QA Summary</b></p>'
    + '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse; font-size: 11px;">'
    + '<tr style="background-color: #f2f2f2; font-weight: bold;">'
    + '<th>Network ID</th><th>Network Name</th><th>Placements Checked</th>'
    + '<th>🟥 BILLING</th><th>🟦 DELIVERY</th><th>🟨 PERFORMANCE</th><th>🟩 COST</th>'
    + '</tr>';

  Object.entries(networkNameMap)
    .filter(function(pair){
      const id = pair[0];
      if (INCLUDE_ZERO_NETS) return true;
      const vc = violationCounts[id] || { "🟥 BILLING":0,"🟦 DELIVERY":0,"🟨 PERFORMANCE":0,"🟩 COST":0 };
      const total = vc["🟥 BILLING"] + vc["🟦 DELIVERY"] + vc["🟨 PERFORMANCE"] + vc["🟩 COST"];
      return total > 0;
    })
    .sort(function(a, b){ return a[1].localeCompare(b[1]); })
    .forEach(function(entry){
      const id = entry[0], name = entry[1];
      const pc = placementCounts[id] || 0;
      const vc = violationCounts[id] || { "🟥 BILLING":0,"🟦 DELIVERY":0,"🟨 PERFORMANCE":0,"🟩 COST":0 };
      networkSummary += '<tr>'
        + '<td>' + id + '</td><td>' + name + '</td><td>' + pc + '</td>'
        + '<td>' + vc["🟥 BILLING"] + '</td><td>' + vc["🟦 DELIVERY"] + '</td><td>' + vc["🟨 PERFORMANCE"] + '</td><td>' + vc["🟩 COST"] + '</td>'
        + '</tr>';
    });
  networkSummary += '</table><br/>';

  // --- Grouped issue summary (unchanged) ---
  const groupedCounts = { "🟥 BILLING": {}, "🟦 DELIVERY": {}, "🟨 PERFORMANCE": {}, "🟩 COST": {} };
  violations.slice(1).forEach(function(r){
    const types = String(r[hMap["Issue Type"]] || "").split(", ");
    types.forEach(function(t){
      const match = t.match(/^(🟥|🟦|🟨|🟩)\s(\w+):\s(.+)/);
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

  // --- Immediate Attention — Key Issues (by Owner) — UPDATED bucket logic
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
    PERF: 1,               // 🟨 Performance
    COST_BIMBAL: 2,        // 🟩 CPC+CPM clicks>impr & $CPC>10
    BILLING: 3,            // 🟥 (Active/Recently Expired/Expired) + tightened rules
    DELIV_STRICT: 4,       // 🟦 Post-flight + clicks>impr + $CPC>10
    DELIV_CPM_ONLY: 5,     // 🟦 Post-flight + CPM-only >$10
    DELIV_GENERAL: 6       // 🟦 Post-flight (any activity) but only if $CPC>10 || $CPM>10
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

    // 🟨 PERFORMANCE: CTR ≥ 90% & CPM ≥ $10
    const isPerformance = /🟨\s*PERFORMANCE: CTR ≥ 90% & CPM ≥ \$?10/.test(issues) ||
                          (ctrPct >= 90 && cpm >= 10);

    // 🟩 CPC+CPM Clicks > Impr & CPC > $10  (both metrics, clicks>impr & CPC>10)
    const isCostBothMetricsClicksGtImpr = /🟩\s*COST: CPC\+CPM Clicks > Impr.*CPC > \$?10/i.test(issues) ||
                                          (both && clicksGtImpr && cpc > 10);

    // 🟥 BILLING (tightened to both metrics, clicks>impr & $CPC>10)
    const isBillingActive   = /🟥\s*BILLING: Active CPC Billing Risk/i.test(issues)   && both && clicksGtImpr && cpc > 10;
    const isBillingRecent   = /🟥\s*BILLING: Recently Expired CPC Risk/i.test(issues) && both && clicksGtImpr && cpc > 10;
    const isBillingExpired  = /🟥\s*BILLING: Expired CPC Risk/i.test(issues)          && both && clicksGtImpr && cpc > 10;

    // 🟦 DELIVERY (Post-Flight) inclusions you selected
    // 1) Strict: post-flight + both metrics + clicks>impr + $CPC>10
    const isDelivStrict = /🟦\s*DELIVERY: Post-Flight Activity/i.test(issues) && isPostFlight && both && clicksGtImpr && cpc > 10;
    // 2) CPM-only > $10 (post-flight)
    const isDelivCpmOnly = /🟦\s*DELIVERY: Post-Flight Activity/i.test(issues) && isPostFlight && (imp > 0 && clk === 0) && cpm > 10;
    // 3) General: post-flight, include only if $CPC>10 OR $CPM>10
    const isDelivGeneral = /🟦\s*DELIVERY: Post-Flight Activity/i.test(issues) && isPostFlight && (cpc > 10 || cpm > 10);

    // ❌ Explicit excludes
    const isCpcOnly = /🟩\s*COST:\s*CPC\s*Only\s*>\s*\$?10/i.test(issues) || (imp === 0 && clk > 0 && cpc > 10);
    const isCpmOnly = /🟩\s*COST:\s*CPM\s*Only\s*>\s*\$?10/i.test(issues) || (imp > 0 && clk === 0 && cpm > 10);
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

  let html = "<p><b>Immediate Attention — Key Issues (by Owner)</b></p>";
  let totalRows = 0;

  for (const rep of owners) {
    if (totalRows >= MAX_TOTAL_OWNER_ROWS) break;
    const arr = perOwner[rep];

    // sort: bucket → advertiser A–Z → clicks desc → impressions desc → placement id
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
      const campShort = o.camp.length > 40 ? o.camp.substring(0, 40) + "…" : o.camp;
      const plcShort  = o.plc.length  > 30 ? o.plc.substring(0, 30)  + "…" : o.plc;
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
    + "<li>Placements with no new impressions since last change (≥ " + thresholdDays + " days): " + staleImp + "</li>"
    + "<li>Placements with no new clicks since last change (≥ " + thresholdDays + " days): " + staleClk + "</li>"
    + "</ul>";

  // Appendix (optional)
  const violationsAppendixHtml =
      '<p><b>What the Violations tab tracks</b></p>'
    + '<ul>'
    + '<li><b>🟥 BILLING</b><ul>'
    + '<li><b>Expired CPC Risk</b> — Ended before this month and clicks &gt; impressions.</li>'
    + '<li><b>Recently Expired CPC Risk</b> — Ended earlier this month and still clicks &gt; impressions.</li>'
    + '<li><b>Active CPC Billing Risk</b> — Active (report date ≤ end date), clicks &gt; impressions, and $CPC &gt; $10.</li>'
    + '</ul></li>'
    + '<li><b>🟦 DELIVERY</b><ul>'
    + '<li><b>Post-Flight Activity</b> — Ended before this month but shows impressions or clicks this month.</li>'
    + '</ul></li>'
    + '<li><b>🟨 PERFORMANCE</b><ul>'
    + '<li><b>CTR ≥ 90% &amp; CPM ≥ $10</b> — Extreme CTR with meaningful CPM spend.</li>'
    + '</ul></li>'
    + '<li><b>🟩 COST</b><ul>'
    + '<li><b>CPC Only &gt; $10</b> — No CPM spend and $CPC &gt; $10.</li>'
    + '<li><b>CPM Only &gt; $10</b> — No CPC spend and $CPM &gt; $10.</li>'
    + '<li><b>CPC+CPM Clicks &gt; Impr &amp; CPC &gt; $10</b> — Both CPC &amp; CPM, clicks &gt; impressions, and $CPC &gt; $10.</li>'
    + '<li><i>(Low Priority tags exist in attachment but are excluded from this section)</i></li>'
    + '</ul></li>'
    + '</ul>';

  // Attachment
  const todayformatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "M.d.yy");
  const fileName = "CM360_QA_Violations_" + todayformatted + ".xlsx";
  const xlsxBlob = createXLSXFromSheet(sheet).setName(fileName);

  // Assemble body
  const subject = "!!!TESTING VS CODE VERSION!!!!!CM360 CPC/CPM FLIGHT QA – " + todayformatted;
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
             + '<p><i>(trimmed for size — full detail in the attached XLSX)</i></p>';
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
  Logger.log('▶ ' + label + ' — START @ ' + new Date(stepStart).toISOString());
  try {
    var out = fn();
    SpreadsheetApp.flush();
    var stepMs = Date.now() - stepStart;
    var totalMs = Date.now() - runStartMs;
    var quotaMs = (quotaMinutes || 6) * 60 * 1000;
    var leftMs = quotaMs - totalMs;

    Logger.log('✅ ' + label + ' — DONE in ' + fmtMs_(stepMs)
      + ' (since run start: ' + fmtMs_(totalMs)
      + ', est. time left: ' + fmtMs_(leftMs) + ')');

    if (leftMs <= 60000) {
      Logger.log('⏳ WARNING: ~' + Math.max(0, Math.floor(leftMs/1000)) + 's left in Apps Script quota window.');
    }
    return out;
  } catch (e) {
    Logger.log('❌ ' + label + ' — ERROR: ' + (e && e.stack ? e.stack : e));
    throw e;
  }
}

// ---------------------
// runItAll (with execution logging per step) — MANUAL USE
// ---------------------
function runItAll() {
  var APPROX_QUOTA_MINUTES = 6; // leave at 6 unless your domain truly has more
  var runStart = Date.now();
  Logger.log('🚀 runItAll — START @ ' + new Date(runStart).toISOString()
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
      Logger.log('⏭ Not enough time left for QA (' + Math.floor(timeLeft/1000) + 's). Scheduling QA handoff.');
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
    Logger.log('🏁 runItAll — FINISHED in ' + fmtMs_(totalMs));
  }
}

// ---------------------
// runItAllMorning (no email, for time-driven trigger)
// ---------------------
function runItAllMorning() {
  var APPROX_QUOTA_MINUTES = 6; // same budget, but we stop before email
  var runStart = Date.now();
  Logger.log('🚀 runItAllMorning — START @ ' + new Date(runStart).toISOString()
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
      Logger.log('⏭ Not enough time left for QA (' + Math.floor(timeLeft/1000) + 's). Scheduling QA handoff.');
      clearQAState_();           // ensure a fresh QA session
      cancelQAChunkTrigger_();   // clear any stale chunk trigger
      scheduleNextQAChunk_(1);   // kick off the first QA chunk shortly
      return;                    // exit cleanly to avoid hitting the 6-min wall
    }

    // 3) Run at most one QA chunk now
    logStep_('runQAOnly (single chunk)', function(){ runQAOnly(); }, runStart, APPROX_QUOTA_MINUTES);

    // 4) Performance spike alert (fast; safe to keep here)
    logStep_('sendPerformanceSpikeAlertIfPre15', function(){ sendPerformanceSpikeAlertIfPre15(); }, runStart, APPROX_QUOTA_MINUTES);

    // ❌ NO sendEmailSummary here — that gets its own trigger/window
  } finally {
    var totalMs = Date.now() - runStart;
    Logger.log('🏁 runItAllMorning — FINISHED in ' + fmtMs_(totalMs));
  }
}

// ---------------------
// runDailyEmailSummary (email only, for separate trigger)
// ---------------------
function runDailyEmailSummary() {
  var APPROX_QUOTA_MINUTES = 6;
  var runStart = Date.now();
  Logger.log('🚀 runDailyEmailSummary — START @ ' + new Date(runStart).toISOString()
             + ' (approx quota: ' + APPROX_QUOTA_MINUTES + ' min)');

  try {
    // sendEmailSummary already:
    //  - skips if QA still has an active session
    //  - skips before the 15th of the month
    logStep_('sendEmailSummary', function(){ sendEmailSummary(); }, runStart, APPROX_QUOTA_MINUTES);
  } finally {
    var totalMs = Date.now() - runStart;
    Logger.log('🏁 runDailyEmailSummary — FINISHED in ' + fmtMs_(totalMs));
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
// - Priority-based scoring (⭐⭐⭐ / ⭐⭐ / ⭐)
// - Status badges (🔴 URGENT | 🟡 REVIEW | 🟢 MONITOR)
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
    Logger.log(`[V2] ✅ Dashboard generated with ${v2Data.length - 1} rows in ${elapsed}s`);
    
    SpreadsheetApp.getUi().alert(`✅ V2 Dashboard generated!\n\n${v2Data.length - 1} violations processed\nTime: ${elapsed}s`);
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
    priority,           // Priority (⭐⭐⭐ / ⭐⭐ / ⭐)
    status,             // Status (🔴/🟡/🟢)
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
  
  // 4 = HIGH: Extreme performance (CTR ≥ 90% + CPM ≥ $10)
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
  if (severity >= 4) return "⭐⭐⭐";
  if (severity === 3) return "⭐⭐";
  return "⭐";
}

// ---------------------
// HELPER: Calculate Status
// ---------------------
function calculateStatus_(priority, severity, category, cpc, placementEnd, reportDate) {
  const end = placementEnd instanceof Date ? placementEnd : new Date(placementEnd);
  const report = reportDate instanceof Date ? reportDate : new Date(reportDate);
  const isExpired = !isNaN(end) && end < report;
  
  // 🔴 URGENT: High priority + severe conditions
  if (priority === "⭐⭐⭐" && (category === "BILLING" || cpc > 20 || (isExpired && severity >= 4))) {
    return "🔴 URGENT";
  }
  
  // 🟡 REVIEW: Medium priority or specific categories
  if (priority === "⭐⭐" || category === "PERFORMANCE" || category === "DELIVERY") {
    return "🟡 REVIEW";
  }
  
  // 🟢 MONITOR: Everything else
  return "🟢 MONITOR";
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
  return detailParts[0] || primary.replace(/🟥|🟨|🟩|🟦/g, "").trim();
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
  
  // PERFORMANCE WASTE: Potential bot traffic (CTR ≥ 90% + high CPM)
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
      .whenTextEqualTo("⭐⭐⭐")
      .setBackground(V2_COLORS.PRIORITY_HIGH)
      .setRanges([sheet.getRange(2, 1, lastRow - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("⭐⭐")
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
      .whenTextContains("🟡 REVIEW")
      .setBackground(V2_COLORS.REVIEW_BG)
      .setFontColor(V2_COLORS.REVIEW_TEXT)
      .setRanges([sheet.getRange(2, 2, lastRow - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("🟢 MONITOR")
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
    
    Logger.log(`[V2] ✅ Exported to Drive: ${fileUrl}`);
    SpreadsheetApp.getUi().alert(`✅ V2 Dashboard exported to Google Drive!\n\nFile: ${fileName}\nFolder: ${monthFolderName}\n\nURL: ${fileUrl}`);
    
    return fileUrl;
    
  } catch (error) {
    Logger.log("[V2] ❌ Export failed: " + error);
    SpreadsheetApp.getUi().alert("❌ Export failed:\n\n" + error);
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
    else if (status.includes("🟡")) reviewCount++;
    else if (status.includes("🟢")) monitorCount++;
    
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
    ["🟡 Review:", reviewCount],
    ["🟢 Monitor:", monitorCount],
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
  summaryData.push(["⚡ BY SEVERITY"]);
  for (let i = 5; i >= 1; i--) {
    const stars = i >= 4 ? "⭐⭐⭐" : i === 3 ? "⭐⭐" : "⭐";
    summaryData.push([`${i} - ${stars}`, severityBreakdown[i]]);
  }
  
  // Write to sheet
  summarySheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);
  
  // Format
  summarySheet.getRange(1, 1, 1, 2).merge().setBackground("#4a86e8").setFontColor("#ffffff").setFontWeight("bold").setFontSize(14);
  summarySheet.setColumnWidth(1, 200);
  summarySheet.setColumnWidth(2, 150);
  
  Logger.log("[V2] ✅ Monthly summary generated");
  SpreadsheetApp.getUi().alert(`✅ Monthly Summary Report Generated!\n\nTotal Violations: ${totalViolations}\n🔴 Urgent: ${urgentCount}\n💰 At Risk: $${totalAtRisk.toFixed(2)}`);
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
    `  • Billing Overcharge: $${billingRisk.toFixed(2)}\n` +
    `  • Performance Waste: $${performanceWaste.toFixed(2)}\n` +
    `  • Post-Flight Spend: $${postFlightSpend.toFixed(2)}\n\n` +
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
    'Expected: ~128 emails (8 months × 16 days)\n' +
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
      subject: '✅ CM360 Historical Archive Complete',
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
      subject: '⚠️ CM360 Archive Error',
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


