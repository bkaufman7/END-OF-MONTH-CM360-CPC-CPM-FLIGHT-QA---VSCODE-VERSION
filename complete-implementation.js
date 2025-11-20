// ===== CM360 End of Month Audit - Complete Implementation =====
// This file contains the full implementation of your CM360 audit system
// Copy this content to replace the Code.gs file in your Google Apps Script project

// ====== Chunked QA execution control ======
const QA_CHUNK_ROWS = 3500;
const QA_TIME_BUDGET_MS = 4.2 * 60 * 1000;
const QA_STATE_KEY = 'qa_progress_v2';      // DocumentProperties key

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