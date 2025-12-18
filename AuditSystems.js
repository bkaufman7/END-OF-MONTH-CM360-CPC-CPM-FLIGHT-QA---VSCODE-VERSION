// =====================================================================================================================
// ========================================== AUDIT SYSTEMS ========================================================
// =====================================================================================================================
// This file contains all archive auditing and gap-fill systems
// - Raw Data Audit (Drive scanning)
// - Violations Audit (Historical reports)
// - Raw Data Gap Fill (Gmail → Drive)
// - Violations Gap Fill (Time Machine automation)

// Constants for Raw Data Audit
const RAW_DATA_AUDIT_STATE_KEY = 'raw_data_audit_state';
const RAW_DATA_AUDIT_TIME_BUDGET_MS = 5.5 * 60 * 1000;

// Constants for Violations Gap Fill
const GAP_FILL_STATE_KEY = 'gap_fill_state';
const GAP_FILL_TRIGGER_KEY = 'gap_fill_trigger_id';
const GAP_FILL_TIME_BUDGET_MS = 5.5 * 60 * 1000;
const VIOLATIONS_ROOT_FOLDER_ID = '1lJm0K1LLo9ez29AcKCc4qtIbBC2uK3a9';

// Constants for Raw Data Gap Fill
const RAW_GAP_FILL_STATE_KEY = 'raw_gap_fill_state';
const RAW_GAP_FILL_TRIGGER_KEY = 'raw_gap_fill_trigger_id';
const RAW_GAP_FILL_TIME_BUDGET_MS = 5.5 * 60 * 1000;
const RAW_DATA_ROOT_FOLDER_ID = '1F53lLe3z5cup338IRY4nhTZQdUmJ9_wk';

// Constants for Raw Data Gap Fill (TEST MODE - 2-Phase System)
const RAW_DATA_TEST_FOLDER_ID = '1qA77_YET8RLiES7X7NoUT5jzTHDJ3k61';
const RAW_DATA_TEST_ARCHIVE_FOLDER_ID = '1WkI8lpIVLW7xtga1MfKZXgCtAFbSD_e6';
const RAW_TEST_PHASE1_STATE_KEY = 'raw_test_phase1_state';
const RAW_TEST_PHASE2_STATE_KEY = 'raw_test_phase2_state';
const RAW_TEST_PHASE1_TRIGGER_KEY = 'raw_test_phase1_trigger_id';
const RAW_TEST_PHASE2_TRIGGER_KEY = 'raw_test_phase2_trigger_id';
const RAW_TEST_DAILY_TRIGGER_KEY = 'raw_test_daily_email_trigger_id';
const RAW_TEST_DAILY_STATS_KEY = 'raw_test_daily_stats';
const RAW_TEST_EMAIL_TARGET = 'platformsolutionsadopshorizon@gmail.com';

// =====================================================================================================================
// ========================================= RAW DATA AUDIT SYSTEM ==================================================
// =====================================================================================================================

/**
 * Combined function to setup and refresh Raw Data audit in one click
 */
function setupAndRefreshRawDataAudit() {
  setupAuditDashboard();
  refreshAuditDashboardChunked();
}

/**
 * Get Raw Data audit state
 */
function getRawDataAuditState_() {
  try {
    const props = PropertiesService.getDocumentProperties();
    const stateJson = props.getProperty(RAW_DATA_AUDIT_STATE_KEY);
    return stateJson ? JSON.parse(stateJson) : null;
  } catch (e) {
    Logger.log('Error loading raw data audit state: ' + e);
    return null;
  }
}

/**
 * Save Raw Data audit state
 */
function saveRawDataAuditState_(state) {
  try {
    const props = PropertiesService.getDocumentProperties();
    props.setProperty(RAW_DATA_AUDIT_STATE_KEY, JSON.stringify(state));
  } catch (e) {
    Logger.log('Error saving raw data audit state: ' + e);
  }
}

/**
 * Clear Raw Data audit state
 */
function clearRawDataAuditState_() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty(RAW_DATA_AUDIT_STATE_KEY);
}

/**
 * Setup Audit Dashboard sheet with date tracking
 */
function setupAuditDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Audit Dashboard");
  
  if (!sheet) {
    sheet = ss.insertSheet("Audit Dashboard");
  }
  
  sheet.clear();
  
  // Set up columns
  sheet.setColumnWidth(1, 120); // Date
  sheet.setColumnWidth(2, 100); // Status
  sheet.setColumnWidth(3, 150); // Files in Drive
  sheet.setColumnWidth(4, 150); // Networks Found
  sheet.setColumnWidth(5, 300); // Missing Networks
  sheet.setColumnWidth(6, 200); // Action
  sheet.setColumnWidth(7, 300); // Notes
  
  // Headers
  const headers = [
    ["Date", "Status", "Files in Drive", "Networks Found", "Missing Networks", "Action", "Notes"]
  ];
  
  sheet.getRange(1, 1, 1, 7).setValues(headers)
    .setFontWeight("bold")
    .setBackground("#4285f4")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  
  sheet.setFrozenRows(1);
  
  SpreadsheetApp.getUi().alert(
    '✅ Audit Dashboard Ready',
    'Audit Dashboard sheet created!\n\n' +
    'Click "Refresh Audit" from the menu to scan your Drive and populate the dashboard.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Refresh audit dashboard with chunking support
 */
function refreshAuditDashboardChunked() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Audit Dashboard");
  
  if (!sheet) {
    setupAuditDashboard();
    sheet = ss.getSheetByName("Audit Dashboard");
  }
  
  const startTime = Date.now();
  let state = getRawDataAuditState_();
  
  // Initialize state if first run
  if (!state) {
    state = {
      dateData: {},
      monthsProcessed: [],
      allNetworks: [],
      startTime: new Date().toISOString()
    };
    
    // Get all networks
    const networksSheet = ss.getSheetByName("Networks");
    if (networksSheet) {
      const networkData = networksSheet.getDataRange().getValues();
      for (let i = 1; i < networkData.length; i++) {
        const networkId = String(networkData[i][0] || '').trim();
        if (networkId) {
          state.allNetworks.push(networkId);
        }
      }
    }
  }
  
  const allNetworks = new Set(state.allNetworks);
  
  // Scan Drive with chunking
  const rootFolderId = RAW_DATA_ROOT_FOLDER_ID;
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  
  const yearFolders = rootFolder.getFoldersByName('2025');
  if (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    const monthFolders = yearFolder.getFolders();
    const monthFoldersList = [];
    
    while (monthFolders.hasNext()) {
      monthFoldersList.push(monthFolders.next());
    }
    
    // Process month folders
    for (const monthFolder of monthFoldersList) {
      const monthName = monthFolder.getName();
      
      // Skip if already processed
      if (state.monthsProcessed.includes(monthName)) {
        Logger.log(`Skipping already processed month: ${monthName}`);
        continue;
      }
      
      // Check time budget BEFORE processing month
      if ((Date.now() - startTime) >= RAW_DATA_AUDIT_TIME_BUDGET_MS) {
        Logger.log(`⏸️ Time budget reached. Saving progress. Processed ${state.monthsProcessed.length}/${monthFoldersList.length} months.`);
        saveRawDataAuditState_(state);
        
        SpreadsheetApp.getUi().alert(
          '⏸️ Audit Paused',
          `Time limit reached. Progress saved.\n\n` +
          `Processed: ${state.monthsProcessed.length}/${monthFoldersList.length} months\n\n` +
          `Run again to continue, or create an auto-resume trigger.`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return;
      }
      
      Logger.log(`Processing month folder: ${monthName}`);
      
      const dateFolders = monthFolder.getFolders();
      while (dateFolders.hasNext()) {
        const dateFolder = dateFolders.next();
        const dateStr = dateFolder.getName();
        
        if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) continue;
        
        if (!state.dateData[dateStr]) {
          state.dateData[dateStr] = { files: 0, networks: [] };
        }
        
        const files = dateFolder.getFiles();
        while (files.hasNext()) {
          const file = files.next();
          const filename = file.getName();
          
          const networkId = extractNetworkIdFromFilename_(filename, getNetworkMap_());
          if (networkId) {
            state.dateData[dateStr].files++;
            if (!state.dateData[dateStr].networks.includes(networkId)) {
              state.dateData[dateStr].networks.push(networkId);
            }
          }
        }
      }
      
      // Mark month as processed and save state immediately
      state.monthsProcessed.push(monthName);
      saveRawDataAuditState_(state);
      Logger.log(`✅ Completed and saved month: ${monthName} (${state.monthsProcessed.length}/${monthFoldersList.length})`);
    }
  }
  
  // All months processed - generate final report
  generateRawDataAuditReport_(sheet, state.dateData, allNetworks);
  clearRawDataAuditState_();
  
  const elapsed = (Date.now() - startTime) / 1000;
  SpreadsheetApp.getUi().alert(
    '✅ Raw Data Audit Complete',
    `Finished scanning all months in ${elapsed.toFixed(1)}s.\n\n` +
    `Check the Audit Dashboard sheet for results.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Generate the audit report from collected data
 */
function generateRawDataAuditReport_(sheet, dateData, allNetworks) {
  // Generate date range (May 1 - Nov 30, 2025)
  const startDate = new Date('2025-05-01');
  const endDate = new Date('2025-11-30');
  const allDates = [];
  
  const current = new Date(startDate);
  while (current <= endDate) {
    allDates.push(Utilities.formatDate(current, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    current.setDate(current.getDate() + 1);
  }
  
  // Build rows
  const rows = [];
  let missingCount = 0;
  let partialCount = 0;
  let completeCount = 0;
  
  for (const dateStr of allDates) {
    const data = dateData[dateStr];
    
    if (!data || data.files === 0) {
      rows.push([
        dateStr,
        '❌ MISSING',
        0,
        0,
        'All networks',
        'Use Time Machine',
        '' // Notes column
      ]);
      missingCount++;
    } else {
      const foundNetworks = data.networks.length;
      const missingNetworks = [];
      
      allNetworks.forEach(netId => {
        if (!data.networks.includes(netId)) {
          missingNetworks.push(netId);
        }
      });
      
      if (missingNetworks.length === 0) {
        rows.push([
          dateStr,
          '✅ COMPLETE',
          data.files,
          foundNetworks,
          '—',
          '—',
          '' // Notes column
        ]);
        completeCount++;
      } else {
        rows.push([
          dateStr,
          '⚠️ PARTIAL',
          data.files,
          foundNetworks,
          missingNetworks.join(', '),
          'Use Gap-Fill',
          '' // Notes column
        ]);
        partialCount++;
      }
    }
  }
  
  // Clear existing data (keep headers)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clear();
  }
  
  // Write data
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
    
    // Format status column
    for (let i = 0; i < rows.length; i++) {
      const statusCell = sheet.getRange(i + 2, 2);
      const status = rows[i][1];
      
      if (status === '✅ COMPLETE') {
        statusCell.setBackground('#d4edda').setFontColor('#155724');
      } else if (status === '⚠️ PARTIAL') {
        statusCell.setBackground('#fff3cd').setFontColor('#856404');
      } else if (status === '❌ MISSING') {
        statusCell.setBackground('#f8d7da').setFontColor('#721c24');
      }
    }
  }
  
  // Add summary at top
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, 7).merge();
  sheet.getRange(1, 1).setValue(
    `📊 Archive Audit Summary: ${completeCount} Complete | ${partialCount} Partial | ${missingCount} Missing | Total: ${allDates.length} days`
  )
    .setFontSize(12)
    .setFontWeight("bold")
    .setBackground("#e8f0fe")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  
  sheet.setRowHeight(1, 35);
}

/**
 * Reset Raw Data Audit state and start fresh
 */
function resetRawDataAudit() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Reset Raw Data Audit',
    'This will clear the audit progress and start scanning from the beginning.\n\nAre you sure?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    clearRawDataAuditState_();
    ui.alert('✅ Reset Complete', 'Raw Data Audit state has been cleared. Run the audit again to start fresh.', ui.ButtonSet.OK);
  }
}

// =====================================================================================================================
// ==================================== VIOLATIONS AUDIT DASHBOARD ===================================================
// =====================================================================================================================

/**
 * Combined function to setup and refresh Violations audit in one click
 */
function setupAndRefreshViolationsAudit() {
  setupViolationsAudit();
  refreshViolationsAudit();
}

/**
 * Setup Violations Audit Dashboard sheet
 */
function setupViolationsAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Violations Audit");
  
  if (!sheet) {
    sheet = ss.insertSheet("Violations Audit");
  }
  
  sheet.clear();
  
  // Set up columns
  sheet.setColumnWidth(1, 120); // Date
  sheet.setColumnWidth(2, 100); // Status
  sheet.setColumnWidth(3, 250); // Drive File
  sheet.setColumnWidth(4, 200); // Drive URL
  
  // Headers
  const headers = [
    ["Date", "Status", "Drive File", "Drive URL"]
  ];
  
  sheet.getRange(1, 1, 1, 4).setValues(headers)
    .setFontWeight("bold")
    .setBackground("#f4b400")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  
  sheet.setFrozenRows(1);
  
  Logger.log('Violations Audit sheet created successfully');
  ss.toast('Violations Audit sheet created!', '✅ Ready', 3);
}

/**
 * Refresh violations audit by scanning Drive ONLY
 */
function refreshViolationsAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Violations Audit");
  
  if (!sheet) {
    Logger.log('ERROR: Violations Audit sheet not found');
    ss.toast('Please run "Violations Audit" first.', '❌ Dashboard Not Found', 5);
    return;
  }
  
  Logger.log('Starting Violations Audit refresh - scanning Drive');
  ss.toast('Scanning Drive for Violations Reports...', '🔄 Scanning', 5);
  
  // Generate date range (April 15, 2025 - Today) - Only 15th onwards of each month
  const allDates = [];
  
  const startDate = new Date('2025-04-15');
  const today = new Date();
  today.setHours(23, 59, 59, 999); // Include full day
  
  // Generate all dates from 15th onwards for each month
  let currentDate = new Date(startDate);
  
  while (currentDate <= today) {
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth() + 1; // JS months are 0-indexed
    const day = currentDate.getDate();
    
    // Only include dates from 15th onwards
    if (day >= 15) {
      const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      allDates.push(dateStr);
    }
    
    // Move to next day
    currentDate.setDate(currentDate.getDate() + 1);
  }
  
  Logger.log(`Generated ${allDates.length} dates from ${startDate.toISOString().split('T')[0]} to ${today.toISOString().split('T')[0]}`);
  
  // Scan Drive for Violations Reports
  const driveData = {};
  const violationsRoot = DriveApp.getFolderById(VIOLATIONS_ROOT_FOLDER_ID);
  
  const monthFolders = violationsRoot.getFolders();
  
  while (monthFolders.hasNext()) {
    const monthFolder = monthFolders.next();
    const files = monthFolder.getFiles();
    
    while (files.hasNext()) {
      const file = files.next();
      const filename = file.getName();
      
      const dateMatch = filename.match(/(\d{4}-\d{2}-\d{2})/);
      if (dateMatch) {
        const dateStr = dateMatch[1];
        driveData[dateStr] = {
          filename: filename,
          url: file.getUrl()
        };
      }
    }
  }
  
  // Build rows
  const rows = [];
  let missingCount = 0;
  let foundCount = 0;
  
  for (const dateStr of allDates) {
    const drive = driveData[dateStr];
    
    if (drive) {
      rows.push([
        dateStr,
        '✅ FOUND',
        drive.filename,
        drive.url
      ]);
      foundCount++;
    } else {
      rows.push([
        dateStr,
        '❌ MISSING',
        '—',
        '—'
      ]);
      missingCount++;
    }
  }
  
  // Clear existing data (keep headers)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clear();
  }
  
  // Write data
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    
    // Format status column
    for (let i = 0; i < rows.length; i++) {
      const statusCell = sheet.getRange(i + 2, 2);
      const status = rows[i][1];
      
      if (status === '✅ FOUND') {
        statusCell.setBackground('#d4edda').setFontColor('#155724');
      } else if (status === '❌ MISSING') {
        statusCell.setBackground('#f8d7da').setFontColor('#721c24');
      }
    }
    
    // Make Drive URLs clickable
    for (let i = 0; i < rows.length; i++) {
      const url = rows[i][3];
      if (url && url !== '—') {
        const urlCell = sheet.getRange(i + 2, 4);
        urlCell.setFormula(`=HYPERLINK("${url}", "Open File")`);
      }
    }
  }
  
  // Add summary at top
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, 4).merge();
  sheet.getRange(1, 1).setValue(
    `📊 Violations Report Audit (Drive Only): ${foundCount} Found | ${missingCount} Missing | Total: ${allDates.length} days (15th-31st only)`
  )
    .setFontSize(12)
    .setFontWeight("bold")
    .setBackground("#fef7e0")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  
  sheet.setRowHeight(1, 35);
  
  Logger.log(`Violations Audit complete: ${foundCount} found, ${missingCount} missing out of ${allDates.length} dates`);
  ss.toast(`Scanned ${allDates.length} dates | ✅ Found: ${foundCount} | ❌ Missing: ${missingCount}`, '✅ Audit Complete', 8);
}
