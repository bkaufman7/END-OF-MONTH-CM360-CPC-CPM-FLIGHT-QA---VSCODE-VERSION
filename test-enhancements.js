// =====================================================================================================================
// ========= ENHANCED TEST MODE FUNCTIONS - Gmail Quota Management & Daily Email Reports ==========================
// =====================================================================================================================

/**
 * Process Phase 1 chunk (download attachments) - ENHANCED VERSION
 */
function processTestPhase1Chunk_ENHANCED() {
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
      dailyStats.gmailSearches++;
      
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
    
    // Track daily progress
    dailyStats.csvsSaved += result.csvsSaved;
    dailyStats.zipsSaved += result.zipsSaved;
    dailyStats.datesCompleted.push(dateStr);
    
    state.currentIndex++;
    
    updateTestPhase1Note_(dateStr, `✅ Downloaded ${result.totalFiles} files (${result.csvsSaved} CSVs, ${result.zipsSaved} ZIPs)`);
    
    // Save state after each date
    saveTestPhase1State_(state);
    saveTodayStats_(dailyStats);
    
    Utilities.sleep(100);
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

/**
 * Enhanced trigger creation - removes old trigger first
 */
function createTestPhase1Trigger_ENHANCED() {
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

/**
 * Enhanced stop triggers - includes daily email trigger
 */
function stopAllTestTriggers_ENHANCED() {
  const props = PropertiesService.getDocumentProperties();
  
  const phase1Id = props.getProperty(RAW_TEST_PHASE1_TRIGGER_KEY);
  const phase2Id = props.getProperty(RAW_TEST_PHASE2_TRIGGER_KEY);
  const dailyId = props.getProperty(RAW_TEST_DAILY_TRIGGER_KEY);
  
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    const id = trigger.getUniqueId();
    if (id === phase1Id || id === phase2Id || id === dailyId) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  props.deleteProperty(RAW_TEST_PHASE1_TRIGGER_KEY);
  props.deleteProperty(RAW_TEST_PHASE2_TRIGGER_KEY);
  props.deleteProperty(RAW_TEST_DAILY_TRIGGER_KEY);
  
  SpreadsheetApp.getUi().alert(
    '🛑 Triggers Stopped',
    'All TEST mode triggers have been deleted.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Enhanced reset - includes daily stats
 */
function resetTestMode_ENHANCED() {
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
    
    stopAllTestTriggers_ENHANCED();
    
    ui.alert('✅ Reset Complete', 'TEST mode has been reset.', ui.ButtonSet.OK);
  }
}
