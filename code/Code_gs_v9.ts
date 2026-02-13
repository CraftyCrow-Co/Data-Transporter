/**
 * ========================================
 * DATA MIGRATOR - OPTIMIZED CODE.GS
 * ========================================
 * Removed unused functions, consolidated logic,
 * improved performance with caching
 */

// ==================== CACHE & CONSTANTS ====================
const CONFIG_FILE_NAME = 'configs.json';
const CONFIG_FOLDER_NAME = 'Data Transporter_User_Configs';
const MAX_SPREADSHEETS = 50;

// ==================== INITIALIZATION & MENU ====================
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Open Sidebar', 'showSidebar')
    .addItem('Refresh Dynamic Menu', 'updateCustomMenus')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Data Transporter')
    .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

function updateCustomMenus() {
  const ui = SpreadsheetApp.getUi();
  let configs = [];
  
  try {
    configs = getSavedConfigs();
  } catch (e) {
    ui.alert("Please authorize the script first by clicking 'Open Sidebar'.");
    return;
  }

  let menu = ui.createAddonMenu().addItem('Open Sidebar', 'showSidebar');
  
  if (configs.length > 0) {
    menu.addSeparator();
    configs.forEach((cfg, i) => {
      if (cfg.hasMenu) menu.addItem('▶ Run: ' + cfg.name, 'runSavedConfig' + i);
    });
  }
  menu.addToUi();
}

// ==================== DRIVE & FILE MANAGEMENT ====================
function getUserSpreadsheets() {
  try {
    // Request authorization explicitly
    const files = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
    const list = [];
    let count = 0;
    
    while (files.hasNext() && count < MAX_SPREADSHEETS) {
      const file = files.next();
      list.push({ name: file.getName(), id: file.getId() });
      count++;
    }
    
    return list.sort((a, b) => a.name.localeCompare(b.name));
  } catch (e) {
  throw new Error("Cannot access Google Drive. Please authorize the add-on and try again.");
    return [];
  }
}

function requestDriveAuthorization() {
  // This function explicitly requests Drive authorization
  // by attempting to access Drive - this will trigger the auth dialog
  try {
    // Try to get files - this requires Drive authorization
    const files = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    // If we can iterate (even just once), authorization worked
    let hasAccess = false;
    if (files.hasNext()) {
      files.next(); // Just try to get one file
      hasAccess = true;
    }
    
    return { 
      success: true, 
      message: hasAccess ? 'Authorization successful' : 'No files found but access granted'
    };
  } catch (e) {
    Logger.log('Authorization error: ' + e.message);
    return { 
      success: false, 
      error: e.message,
      message: 'Failed to access Drive. The authorization may have been declined or there was a network error.'
    };
  }
}

function getUserConfigFile_() {
  const folders = DriveApp.getFoldersByName(CONFIG_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG_FOLDER_NAME);

  const files = folder.getFilesByName(CONFIG_FILE_NAME);
  if (files.hasNext()) return files.next();

  return folder.createFile(
    CONFIG_FILE_NAME,
    JSON.stringify({ configs: [] }, null, 2),
    MimeType.PLAIN_TEXT
  );
}

function readConfigData_() {
  const file = getUserConfigFile_();
  return JSON.parse(file.getBlob().getDataAsString() || '{"configs":[]}');
}

function writeConfigData_(data) {
  const file = getUserConfigFile_();
  file.setContent(JSON.stringify(data, null, 2));
}

// ==================== SPREADSHEET UTILITIES ====================
function getActiveSpreadsheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getName();
}

function getActiveSheetContext() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  return {
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheetName: sheet.getName()
  };
}

function getSheetsInSpreadsheet(id) {
  if (!id || id === "CREATE_NEW") return ["Sheet1"];
  
  try {
    const ss = SpreadsheetApp.openById(id);
    return ss.getSheets().map(s => s.getName());
  } catch (e) {
    return ["Error: Access Denied"];
  }
}

function getSheetColumns(ssId, sheetName, headerRow) {
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName(sheetName);
  const rowNum = parseInt(headerRow) || 1;
  
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  
  const headers = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  return headers.map(h => h.toString().trim()).filter(h => h !== "");
}

function getInitialData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    return {
      activeName: ss.getName(),
      activeId: ss.getId(),
      sheets: sheets.map(s => s.getName()),
      files: getUserSpreadsheets(), // Load spreadsheets list
      currentSpreadsheet: {
        id: ss.getId(),
        name: ss.getName(),
        url: ss.getUrl()
      }
    };
  } catch (e) {
    throw new Error("Initialization Error: " + e.message);
  }
}

// ==================== EXECUTION ROUTER ====================
function runExecution(config) {
  if (!config) {
    return { status: 'error', message: 'No configuration provided' };
  }
  
  // Update lastRun timestamp if this is a saved config
  if (config.id) {
    try {
      const data = readConfigData_();
      const idx = data.configs.findIndex(c => c.id === config.id);
      if (idx !== -1) {
        data.configs[idx].lastRun = new Date().toISOString();
        writeConfigData_(data);
      }
    } catch (e) {
      // Continue even if tracking fails
      Logger.log('Failed to update lastRun: ' + e.message);
    }
  }
  
  // Check if this is an Archive config
  if (config.mode === 'archive' || config.archive) {
    const result = runArchive(config);
    
    // Store the file URL in the config if successful
    if (result.status === 'success' && result.fileUrl && config.id) {
      try {
        const data = readConfigData_();
        const idx = data.configs.findIndex(c => c.id === config.id);
        if (idx !== -1) {
          data.configs[idx].lastArchivedFileUrl = result.fileUrl;
          writeConfigData_(data);
        }
      } catch (e) {
        Logger.log('Failed to store file URL: ' + e.message);
      }
    }
    
    return result;
  }
  
  // Otherwise it's a Transfer config
  return runMigration(config);
}

// ==================== CORE MIGRATION ENGINE ====================
function checkSheetProtection_(sheet) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (protections.length > 0) {
    const protection = protections[0];
    if (!protection.canEdit()) {
      throw new Error('Destination sheet "' + sheet.getName() + '" is protected and cannot be edited. Please unprotect the sheet or select a different destination.');
    }
  }
}

function runMigration(config) {
  try {
    config.dataMode = (config.dataMode || "").toLowerCase();
    const execType = config.exec || 'std';

    // 1. DESTINATION SPREADSHEET
    let destSS;
    if (config.destSS === "new") {
      destSS = SpreadsheetApp.create(
        "Migrated Data - " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
      );
    } else if (config.destSS === "current") {
      destSS = SpreadsheetApp.getActiveSpreadsheet();
    } else {
      destSS = SpreadsheetApp.openById(config.destSS);
    }

    // 2. DESTINATION SHEET
    let destSheet = destSS.getSheetByName(config.destSheet);
    if (!destSheet) {
      destSheet = destSS.insertSheet(config.destSheet || "Migrated_Data");
    }
    
    // Check if destination sheet is protected before proceeding
    checkSheetProtection_(destSheet);

    // 3. TABLE STYLE (FULL SHEET SYNC)
    if (config.includeTableFormats) {
      const src = config.sources[0];
      const sourceSS = SpreadsheetApp.openById(src.srcSS.replace(/.*\/d\/(.*)\/.*/, "$1"));
      const sourceSheet = sourceSS.getSheetByName(src.srcSheet) || sourceSS.getSheets()[0];
      const sheetName = sourceSheet.getName();

      const tempSheet = sourceSheet.copyTo(destSS);
      const tempName = `__TEMP_SYNC__${Date.now()}`;
      tempSheet.setName(tempName);

      const existing = destSS.getSheetByName(sheetName);
      if (existing) destSS.deleteSheet(existing);

      tempSheet.setName(sheetName);
      return { status: "success", message: "Table Style synced successfully" };
    }

    // 4. START CELL & MODE
    const startCell = destSheet.getRange(config.destStart || "A1");
    const startCol = startCell.getColumn();
    let writeRow;
    let headersWritten = false;
    let headerWriteRow = null;
    let totalWritten = 0;

    if (config.dataMode === "replace") {
      destSheet.clear();
      writeRow = startCell.getRow();
    } else {
      const lastRow = destSheet.getLastRow();
      writeRow = lastRow === 0 ? startCell.getRow() : lastRow + 1;
      headersWritten = lastRow > 0;
    }

    // 5. PROCESS SOURCES
    for (const src of config.sources) {
      const sourceSS = SpreadsheetApp.openById(src.srcSS.replace(/.*\/d\/(.*)\/.*/, "$1"));
      const sourceSheet = sourceSS.getSheetByName(src.srcSheet);
      
      // Validate that the sheet exists
      if (!sourceSheet) {
        throw new Error(`Sheet "${src.srcSheet}" not found in source spreadsheet "${sourceSS.getName()}". Please check the sheet name.`);
      }
      
      const headerRow = parseInt(src.headerRow) || 1;
      const range = src.srcRange ? sourceSheet.getRange(src.srcRange) : sourceSheet.getDataRange();

      const rawValues = range.getValues();
      const rawFormulas = range.getFormulas();
      const values = config.includeFormulas
        ? rawFormulas.map((r, i) => r.map((f, j) => f || rawValues[i][j]))
        : rawValues;

      const rawHeaders = rawValues[headerRow - 1];
      let data = values.slice(headerRow);

      // COLUMN SELECTION
      let activeHeaders = [...rawHeaders];
      let colIndices = rawHeaders.map((_, i) => i);

      if (src.includedColumns?.length) {
        colIndices = src.includedColumns.map(h => rawHeaders.indexOf(h)).filter(i => i > -1);
        activeHeaders = colIndices.map(i => rawHeaders[i]);
        data = data.map(r => colIndices.map(i => r[i]));
      }

      // TIMESTAMP
      if (config.addTimestamp) {
        const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        if (!headersWritten) activeHeaders.unshift("Import Timestamp");
        data = data.map(r => [ts, ...r]);
      }

      // WRITE HEADERS
      if (config.includeHeaders && !headersWritten) {
        headerWriteRow = writeRow;
        destSheet.getRange(writeRow, startCol, 1, activeHeaders.length).setValues([activeHeaders]);
        writeRow++;
        headersWritten = true;
      }

      // WRITE DATA BASED ON EXECUTION TYPE
      if (data.length) {
        if (execType === 'row') {
          // ROW BY ROW EXECUTION
          for (let i = 0; i < data.length; i++) {
            destSheet.getRange(writeRow, startCol, 1, data[i].length).setValues([data[i]]);
            writeRow++;
            totalWritten++;
            SpreadsheetApp.flush(); // Force update after each row
          }
        } else if (execType === 'batch') {
          // BATCH EXECUTION
          const batchSize = parseInt(config.batchSize) || 100;
          for (let i = 0; i < data.length; i += batchSize) {
            const batch = data.slice(i, i + batchSize);
            destSheet.getRange(writeRow, startCol, batch.length, batch[0].length).setValues(batch);
            writeRow += batch.length;
            totalWritten += batch.length;
            SpreadsheetApp.flush(); // Force update after each batch
          }
        } else {
          // STANDARD EXECUTION (ALL AT ONCE)
          destSheet.getRange(writeRow, startCol, data.length, data[0].length).setValues(data);
          writeRow += data.length;
          totalWritten += data.length;
        }
      }

      // 6. FORMATTING
      if (config.includeFormatting || config.includeHeaderFormats || config.includeCondFormats) {
        const tempSheet = sourceSheet.copyTo(destSS);
        tempSheet.setName(`__FORMAT_TEMP__${Date.now()}`);

        const numCols = tempSheet.getLastColumn();
        const numRows = tempSheet.getLastRow();

        if (config.includeFormatting) {
          tempSheet.getRange(headerRow + 1, 1, numRows - headerRow, numCols)
            .copyTo(destSheet.getRange(headerRow + 1, startCol, numRows - headerRow, numCols), { formatOnly: true });
        }

        if (config.includeHeaderFormats && headerWriteRow !== null) {
          tempSheet.getRange(headerRow, 1, 1, numCols)
            .copyTo(destSheet.getRange(headerWriteRow, startCol, 1, numCols), { formatOnly: true });
        }

        if (config.includeCondFormats) {
          destSheet.setConditionalFormatRules(tempSheet.getConditionalFormatRules());
        }

        destSS.deleteSheet(tempSheet);
      }
    }

    const execLabel = execType === 'row' ? ' (Row-by-Row)' : execType === 'batch' ? ' (Batch)' : '';
    return { status: "success", message: `Migration complete${execLabel}. Rows written: ${totalWritten}` };

  } catch (e) {
    return { status: "error", message: e.stack || e.toString() };
  }
}

// ==================== CONFIGURATION MANAGEMENT ====================
function saveConfiguration(config, existingId = null) {
  const data = readConfigData_();
  const configs = data.configs || [];

  if (!config.id) config.id = "cfg_" + Utilities.getUuid();
  
  config.updatedAt = new Date().toISOString();
  config.timestamp = config.updatedAt;

  if (existingId) {
    const idx = configs.findIndex(c => c.id === existingId);
    if (idx !== -1) {
      const oldConfig = configs[idx];
      configs[idx] = {
        ...oldConfig,
        ...config,
        name: oldConfig.name,
        triggerActive: oldConfig.triggerActive || false,
        hasMenu: oldConfig.hasMenu || false,
        updatedAt: new Date().toISOString()
      };
    } else {
      configs.push(config);
    }
  } else {
    config.triggerActive = false;
    configs.push(config);
  }

  data.configs = configs;
  writeConfigData_(data);
  
  return { 
    status: 'success', 
    message: existingId ? 'Configuration updated successfully' : 'Configuration saved successfully' 
  };
}

function getSavedConfigs(filterToCurrent = false) {
  const data = readConfigData_();
  let configs = data.configs || [];
  
  if (filterToCurrent) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeId = ss.getId();
    const activeSheetName = ss.getActiveSheet().getName();
    
    configs = configs.filter(cfg => {
      const isDest = (cfg.destSS === activeId || cfg.destSS === 'current') && (cfg.destSheet === activeSheetName);
      const isSource = cfg.sources && cfg.sources.some(src => 
        (src.srcSS === activeId || src.srcSS === 'current') && src.srcSheet === activeSheetName
      );
      return isDest || isSource;
    });
  }

  return configs.sort((a, b) => {
    const ta = new Date(a.updatedAt || a.timestamp || 0).getTime();
    const tb = new Date(b.updatedAt || b.timestamp || 0).getTime();
    return tb - ta;
  });
}

function getConfigById(id) {
  const data = readConfigData_();
  return data.configs.find(c => c.id === id);
}

function deleteConfigById(id) {
  const data = readConfigData_();
  const index = data.configs.findIndex(c => c.id === id);
  
  if (index !== -1) {
    deleteTriggersForConfigId_(id);
    data.configs.splice(index, 1);
    writeConfigData_(data);
    return true;
  }
  return false;
}

function toggleMenuFlagById(id) {
  const data = readConfigData_();
  const index = data.configs.findIndex(c => c.id === id);
  if (index === -1) return false;

  data.configs[index].hasMenu = !data.configs[index].hasMenu;
  writeConfigData_(data);
  updateCustomMenus();
  return true;
}

// ==================== SCHEDULING ====================
function createScheduleComplex(id, scheduleData) {
  const data = readConfigData_();
  const cfg = data.configs.find(c => c.id === id);
  if (!cfg) throw new Error("Configuration not found for ID: " + id);

  deleteTriggersForConfigId_(cfg.id);

  const handler = createTriggerHandler_(cfg.id);
  let tb = ScriptApp.newTrigger(handler).timeBased();

  const hour = parseInt(scheduleData.hour) || 9;
  const every = parseInt(scheduleData.every) || 1;

  if (scheduleData.unit === 'minute') {
    tb.everyMinutes(every).create();
  } else if (scheduleData.unit === 'hour') {
    tb.everyHours(every).create();
  } else if (scheduleData.unit === 'day') {
    tb.everyDays(every).atHour(hour).create();
  } else if (scheduleData.unit === 'week') {
    scheduleData.days.forEach(d =>
      ScriptApp.newTrigger(handler).timeBased().onWeekDay(getDayEnum(parseInt(d))).atHour(hour).create()
    );
  } else if (scheduleData.unit === 'month') {
    tb.onMonthDay(parseInt(scheduleData.dayNum) || 1).atHour(hour).create();
  }

  cfg.schedule = scheduleData;
  cfg.alertEmail = scheduleData.alertEmail;
  cfg.triggerActive = true;

  writeConfigData_(data);
  return "Trigger created successfully";
}

function createTriggerHandler_(configId) {
  const shortId = configId.replace(/-/g, '');
  const fnName = "runConfig_" + shortId;
  PropertiesService.getScriptProperties().setProperty(fnName, configId);
  return "triggerRouter";
}

function triggerRouter(e) {
  const allTriggers = ScriptApp.getProjectTriggers();
  const currentTrigger = allTriggers.find(t => t.getUniqueId() === e.triggerUid);
  const handlerName = currentTrigger.getHandlerFunction();
  const configId = PropertiesService.getScriptProperties().getProperty(handlerName);
  
  if (configId) executeByConfigId(configId);
}

function executeByConfigId(configId) {
  const data = readConfigData_();
  const cfg = data.configs.find(c => c.id === configId);
  if (!cfg) return;

  cfg.lastRun = new Date().toLocaleString();
  cfg.triggerActive = true;
  writeConfigData_(data);
  runMigration(cfg);
}

function deleteTriggersForConfigId_(configId) {
  const shortId = configId.replace(/-/g, '');
  const targetHandler = "runConfig_" + shortId;
  const triggers = ScriptApp.getProjectTriggers();
  
  triggers.forEach(t => {
    const handler = t.getHandlerFunction();
    if (handler === targetHandler || handler === "triggerRouter") {
      ScriptApp.deleteTrigger(t);
    }
  });

  PropertiesService.getScriptProperties().deleteProperty(targetHandler);
}

function stopScheduleById(id) {
  const data = readConfigData_();
  const cfg = data.configs.find(c => c.id === id);
  if (!cfg) return { status: "error", message: "Config not found" };

  deleteTriggersForConfigId_(cfg.id);
  cfg.triggerActive = false;
  writeConfigData_(data);

  return { status: "success", message: "Schedule stopped." };
}

function getDayEnum(index) {
  const days = [
    ScriptApp.WeekDay.SUNDAY, ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY,
    ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY,
    ScriptApp.WeekDay.SATURDAY
  ];
  return days[index];
}

// ==================== EXPORT CONFIGURATION ====================
function exportConfigAsCodeById(id) {
  const data = readConfigData_();
  const cfg = data.configs.find(c => c.id === id);
  if (!cfg) return "Error: Configuration not found.";

  const cfgName = cfg.name || "Unnamed Config";
  
  // Check if this is an Archive config
  if (cfg.mode === 'archive') {
    return generateArchiveCode_(cfg, cfgName);
  } else {
    return generateTransferCode_(cfg, cfgName);
  }
}

function generateArchiveCode_(cfg, cfgName) {
  const arc = cfg.archive;
  const configJson = JSON.stringify(arc, null, 2);
  
  return `/**
 * Standalone Archive Script for: ${cfgName}
 * Copy this into the Script Editor (Extensions > Apps Script) of ANY Google Sheet.
 * This will create an archive of the specified spreadsheet.
 */
function runStandaloneArchive() {
  const archiveConfig = ${configJson};
  
  try {
    // Get source file
    const sourceFile = DriveApp.getFileById(archiveConfig.sourceId);
    const sourceName = archiveConfig.sourceName || sourceFile.getName();
    
    // Get destination folder (if specified)
    let targetFolder = null;
    if (archiveConfig.destinationMode === 'select' && archiveConfig.folderId) {
      targetFolder = DriveApp.getFolderById(archiveConfig.folderId);
    }
    
    // Generate archive name with timestamp
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
    
    let finalName = '';
    if (archiveConfig.fileName && archiveConfig.fileName.trim()) {
      finalName = archiveConfig.fileName.trim();
    } else if (archiveConfig.sourceType === 'sheet' && archiveConfig.sheets && archiveConfig.sheets.length > 0) {
      const sheetNames = archiveConfig.sheets.join(', ');
      finalName = sourceName + ' (' + sheetNames + ')_Archived_' + timestamp;
    } else {
      finalName = sourceName + '_Archived_' + timestamp;
    }
    
    // Create copy
    const copiedFile = targetFolder
      ? sourceFile.makeCopy(finalName, targetFolder)
      : sourceFile.makeCopy(finalName);
    
    // If specific sheets selected, delete others
    if (archiveConfig.sourceType === 'sheet' && archiveConfig.sheets && archiveConfig.sheets.length > 0) {
      const ss = SpreadsheetApp.openById(copiedFile.getId());
      const sheets = ss.getSheets();
      sheets.forEach(function(sh) {
        if (archiveConfig.sheets.indexOf(sh.getName()) === -1) {
          ss.deleteSheet(sh);
        }
      });
    }
    
    // Cleanup blank rows/columns if enabled
    if (archiveConfig.cleanup) {
      try {
        cleanupBlankSpace(copiedFile.getId());
      } catch (e) {
        // Continue even if cleanup fails
      }
    }
    
    SpreadsheetApp.getUi().alert('Archive created: ' + finalName + '\\n\\nFile URL: ' + copiedFile.getUrl());
  } catch (e) {
    SpreadsheetApp.getUi().alert('Archive Error: ' + e.toString());
  }
}

function cleanupBlankSpace(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  ss.getSheets().forEach(function(sh) {
    try {
      const dataRange = sh.getDataRange();
      const values = dataRange.getValues();
      const maxRows = sh.getMaxRows();
      const maxCols = sh.getMaxColumns();
      
      // Find last row with data
      let lastRow = 0;
      for (let i = values.length - 1; i >= 0; i--) {
        let hasData = false;
        for (let j = 0; j < values[i].length; j++) {
          if (values[i][j] !== '' && values[i][j] !== null) {
            hasData = true;
            break;
          }
        }
        if (hasData) {
          lastRow = i + 1;
          break;
        }
      }
      
      // Find last column with data
      let lastCol = 0;
      for (let j = values[0].length - 1; j >= 0; j--) {
        let hasData = false;
        for (let i = 0; i < values.length; i++) {
          if (values[i][j] !== '' && values[i][j] !== null) {
            hasData = true;
            break;
          }
        }
        if (hasData) {
          lastCol = j + 1;
          break;
        }
      }
      
      // Delete empty rows
      if (lastRow > 0 && lastRow < maxRows) {
        sh.deleteRows(lastRow + 1, maxRows - lastRow);
      }
      
      // Delete empty columns
      if (lastCol > 0 && lastCol < maxCols) {
        sh.deleteColumns(lastCol + 1, maxCols - lastCol);
      }
    } catch (e) {
      // Skip if frozen or protected
    }
  });
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('▶ Archive Script')
    .addItem('Create Archive: ${cfgName}', 'runStandaloneArchive')
    .addToUi();
}`;
}

function generateTransferCode_(cfg, cfgName) {
  const configJson = JSON.stringify(cfg, null, 2);
  
  return `/**
 * Standalone Migration Script for: ${cfgName}
 * Copy this into the Script Editor (Extensions > Apps Script) of ANY Google Sheet.
 */
function runStandaloneMigration() {
  const cfg = ${configJson};
  const ui = SpreadsheetApp.getUi();
  
  try {
    const destSS = SpreadsheetApp.getActiveSpreadsheet();
    let destSheet = destSS.getSheetByName(cfg.destSheet) || destSS.insertSheet(cfg.destSheet);
    
    cfg.sources.forEach(src => {
      const srcSS = SpreadsheetApp.openById(src.srcSS);
      const srcSheet = srcSS.getSheetByName(src.srcSheet);
      const range = src.srcRange ? srcSheet.getRange(src.srcRange) : srcSheet.getDataRange();
      const values = range.getValues();
      
      const targetRange = destSheet.getRange(cfg.destStart || "A1");
      destSheet.getRange(
        targetRange.getRow(),
        targetRange.getColumn(),
        values.length,
        values[0].length
      ).setValues(values);
    });
    
    ui.alert("Success: Data migrated using standalone script.");
  } catch (e) {
    ui.alert("Migration Error: " + e.toString());
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('▶ Dynamic Migrator')
    .addItem('Execute: ${cfgName}', 'runStandaloneMigration')
    .addToUi();
}`;
}

// ==================== MENU EXECUTION HANDLERS ====================
function runSavedConfig0() { runSavedConfigByIndex(0); }
function runSavedConfig1() { runSavedConfigByIndex(1); }
function runSavedConfig2() { runSavedConfigByIndex(2); }
function runSavedConfig3() { runSavedConfigByIndex(3); }
function runSavedConfig4() { runSavedConfigByIndex(4); }
function runSavedConfig5() { runSavedConfigByIndex(5); }
function runSavedConfig6() { runSavedConfigByIndex(6); }
function runSavedConfig7() { runSavedConfigByIndex(7); }
function runSavedConfig8() { runSavedConfigByIndex(8); }
function runSavedConfig9() { runSavedConfigByIndex(9); }

function runSavedConfigByIndex(index) {
  const configs = getSavedConfigs();
  if (!configs[index]) throw new Error('No config at index: ' + index);
  return runExecution(configs[index]);
}

function runSavedConfigById(id) {
  const configs = getSavedConfigs();
  const cfg = configs.find(c => c.id === id);
  if (!cfg) throw new Error('Config not found for ID: ' + id);
  return runExecution(cfg);
}

// ==================== UTILITY FUNCTIONS ====================
function loadSavedConfigs() {
  return getSavedConfigs(false);
}

// ==================== ARCHIVE ENGINE ====================
function runArchive(config) {
  try {
    const arc = config.archive;
    
    if (!arc || !arc.sourceId) {
      throw new Error("Archive source spreadsheet is missing");
    }

    // 1. RESOLVE SOURCE FILE
    const sourceFileId = extractFileId_(arc.sourceId);
    const sourceFile = DriveApp.getFileById(sourceFileId);
    const sourceName = arc.sourceName || sourceFile.getName();

    // 2. RESOLVE DESTINATION FOLDER
    let targetFolder = null;
    
    Logger.log('Archive destination mode: ' + arc.destinationMode);
    Logger.log('Archive folder ID: ' + arc.folderId);
    
    if (arc.destinationMode === 'select' && arc.folderId && arc.folderId.trim() !== '') {
      try {
        targetFolder = DriveApp.getFolderById(arc.folderId);
        Logger.log('Target folder resolved: ' + targetFolder.getName());
      } catch (e) {
        throw new Error('Cannot access the selected folder. It may have been deleted or you lost access. Please select again.');
      }
    }
    
  if (arc.destinationMode === 'paste' && arc.folderId && arc.folderId.trim() !== '') {
  try {
    const folderId = extractFileId_(arc.folderId);
    if (!folderId || folderId.length < 20) {  // Google IDs are ~33 chars
      throw new Error('Invalid folder URL or ID');
    }
    targetFolder = DriveApp.getFolderById(folderId);
    Logger.log('Target folder from paste: ' + targetFolder.getName());
  } catch (e) {
    throw new Error('Invalid folder URL/ID pasted. Please check the URL and try again.');
  }
}

    // 3. RESOLVE FILE NAME WITH HUMAN-READABLE DATE
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    let finalName = '';
    
    // If custom name provided, sanitize and append suffix to it
    if (arc.fileName && arc.fileName.trim()) {
      const sanitizedName = sanitizeFileName_(arc.fileName.trim());
      finalName = sanitizedName + '_Archived_' + timestamp;
    } else {
      // Auto-generate based on source type
      if (arc.sourceType === 'sheet' && arc.sheets && arc.sheets.length > 0) {
        // Specific sheets: SourceName (Sheet1, Sheet2)_Archived_dd-mm-yyyy
        const sheetNames = arc.sheets.join(', ');
        finalName = `${sourceName} (${sheetNames})_Archived_${timestamp}`;
      } else {
        // Entire file: SourceName_Archived_dd-mm-yyyy
        finalName = `${sourceName}_Archived_${timestamp}`;
      }
    }

    // 4. CREATE COPY
    const copiedFile = targetFolder
      ? sourceFile.makeCopy(finalName, targetFolder)
      : sourceFile.makeCopy(finalName);

    // 5. SHEET-LEVEL FILTERING
    if (arc.sourceType === 'sheet' && arc.sheets?.length) {
      retainOnlySelectedSheets_(copiedFile.getId(), arc.sheets);
    }

    // 6. APPLY FILTERS (if any)
    if (arc.filters && arc.filters.length > 0) {
      applyArchiveFilters_(copiedFile.getId(), arc.filters);
    }

    // 7. OPTIONAL CLEANUP
    let cleanupMessage = '';
    if (arc.cleanup === true) {
      const cleanupResult = cleanupBlankSpace_(copiedFile.getId());
      if (cleanupResult.errors && cleanupResult.errors.length > 0) {
        cleanupMessage = ' Note: ' + cleanupResult.message;
      }
    }

    return {
      status: "success",
      message: "Archive created successfully: " + finalName + cleanupMessage,
      fileUrl: copiedFile.getUrl()
    };

  } catch (e) {
    return {
      status: "error",
      message: e.message || String(e)
    };
  }
}

// ==================== ARCHIVE HELPERS ====================
function sanitizeFileName_(name) {
  // Remove prohibited Google Drive filename characters: / \ ? * : < > | "
  return name.replace(/[\/\\?*:<>|"]/g, '-');
}

function extractFileId_(input) {
  if (!input) return null;
  const match = input.match(/[-\w]{25,}/);
  return match ? match[0] : input;
}

function retainOnlySelectedSheets_(spreadsheetId, allowedSheetNames) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheets = ss.getSheets();
  
  sheets.forEach(sh => {
    if (!allowedSheetNames.includes(sh.getName())) {
      ss.deleteSheet(sh);
    }
  });
}

function applyArchiveFilters_(spreadsheetId, filters) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  
  filters.forEach(filter => {
    const sheetName = filter.sheet;
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) return;
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length <= 1) return; // No data rows
    
    // Get header row
    const headers = values[0];
    const columnIndex = headers.indexOf(filter.column);
    
    if (columnIndex === -1) return; // Column not found
    
    // Filter rows (keep header)
    const filteredRows = [values[0]]; // Keep header
    
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][columnIndex];
      
      if (matchesFilter_(cellValue, filter)) {
        filteredRows.push(values[i]);
      }
    }
    
    // Clear sheet and write filtered data
    sheet.clear();
    if (filteredRows.length > 0) {
      sheet.getRange(1, 1, filteredRows.length, filteredRows[0].length).setValues(filteredRows);
    }
  });
}

function matchesFilter_(cellValue, filter) {
  const cellStr = String(cellValue ?? "").trim().toLowerCase();
  const target = String(filter.value ?? "").toLowerCase();
  
  // Date filters
  if (filter.type === 'date') {
    return checkDateFilter_(cellValue, filter.operator, filter.value);
  }
  
  // Text/Number filters
  switch (filter.operator) {
    case "is_empty": return cellStr === "";
    case "is_not_empty": return cellStr !== "";
    case "equals": return cellStr === target;
    case "contains": return cellStr.includes(target);
    case "not_equal": return cellStr !== target;
    case "starts_with": return cellStr.startsWith(target);
    case "ends_with": return cellStr.endsWith(target);
    case "greater_than": return parseFloat(cellValue) > parseFloat(filter.value);
    case "less_than": return parseFloat(cellValue) < parseFloat(filter.value);
    default: return true;
  }
}

function checkDateFilter_(cellVal, operator, nValue) {
  if (cellVal === "" || cellVal === null || cellVal === undefined) return false;
  
  const cellDate = new Date(cellVal);
  if (isNaN(cellDate.getTime())) return false;
  cellDate.setHours(0,0,0,0);
  
  const today = new Date();
  today.setHours(0,0,0,0);
  
  const oneDay = 24 * 60 * 60 * 1000;
  const diffTime = cellDate.getTime() - today.getTime();
  const diffDays = Math.round(diffTime / oneDay);

  switch (operator) {
    case "today": return diffDays === 0;
    case "yesterday": return diffDays === -1;
    case "tomorrow": return diffDays === 1;
    case "last_n_days": 
      const nLast = parseInt(nValue) || 0;
      return diffDays <= 0 && diffDays >= -nLast;
    case "next_n_days":
      const nNext = parseInt(nValue) || 0;
      return diffDays >= 0 && diffDays <= nNext;
    default: return false;
  }
}

function cleanupBlankSpace_(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const errors = [];
  let successCount = 0;
  
  ss.getSheets().forEach(sh => {
    try {
      const sheetName = sh.getName();
      
      // Check if sheet is protected
      const protections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      if (protections.length > 0) {
        errors.push(`Sheet "${sheetName}" is protected - skipped`);
        return;
      }
      
      const maxRows = sh.getMaxRows();
      const maxCols = sh.getMaxColumns();
      
      if (maxRows === 0 || maxCols === 0) return;
      
      const dataRange = sh.getDataRange();
      const values = dataRange.getValues();
      
      if (values.length === 0) return;
      
      // Find last row with data
      let lastRowWithData = 0;
      for (let i = values.length - 1; i >= 0; i--) {
        if (values[i].some(cell => cell !== "" && cell !== null)) {
          lastRowWithData = i + 1;
          break;
        }
      }
      
      // Find last column with data
      let lastColWithData = 0;
      for (let j = values[0].length - 1; j >= 0; j--) {
        if (values.some(row => row[j] !== "" && row[j] !== null)) {
          lastColWithData = j + 1;
          break;
        }
      }
      
      // Check for frozen rows
      const frozenRows = sh.getFrozenRows();
      const frozenCols = sh.getFrozenColumns();
      
      // Delete empty rows at the end (if not frozen)
      if (lastRowWithData > 0 && lastRowWithData < maxRows) {
        const rowsToDelete = maxRows - lastRowWithData;
        if (rowsToDelete > 0) {
          // Check if we're trying to delete frozen rows
          if (lastRowWithData < frozenRows) {
            errors.push(`Sheet "${sheetName}" has frozen rows - cannot delete all blank rows`);
          } else {
            try {
              sh.deleteRows(lastRowWithData + 1, rowsToDelete);
              successCount++;
            } catch (e) {
              errors.push(`Sheet "${sheetName}" rows: ${e.message}`);
            }
          }
        }
      }
      
      // Delete empty columns at the end (if not frozen)
      if (lastColWithData > 0 && lastColWithData < maxCols) {
        const colsToDelete = maxCols - lastColWithData;
        if (colsToDelete > 0) {
          // Check if we're trying to delete frozen columns
          if (lastColWithData < frozenCols) {
            errors.push(`Sheet "${sheetName}" has frozen columns - cannot delete all blank columns`);
          } else {
            try {
              sh.deleteColumns(lastColWithData + 1, colsToDelete);
              successCount++;
            } catch (e) {
              errors.push(`Sheet "${sheetName}" columns: ${e.message}`);
            }
          }
        }
      }
    } catch (e) {
      errors.push(`Sheet "${sh.getName()}": ${e.message}`);
    }
  });
  
  return {
    success: successCount > 0,
    errors: errors,
    message: errors.length > 0 
      ? `Cleanup completed with warnings: ${errors.join('; ')}` 
      : 'Cleanup completed successfully'
  };
}

function showFolderPicker() {
  const html = HtmlService
    .createHtmlOutputFromFile('FolderPicker')
    .setWidth(400)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Drive Folder');
}

function getDriveFolders() {
  const folders = [];
  const it = DriveApp.searchFolders('trashed = false');
  
  while (it.hasNext()) {
    const f = it.next();
    folders.push({ id: f.getId(), name: f.getName() });
  }
  return folders;
}

function setSelectedArchiveFolder(id, name) {
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty('ARCHIVE_FOLDER_ID', id);
  userProps.setProperty('ARCHIVE_FOLDER_NAME', name);
  userProps.setProperty('ARCHIVE_FOLDER_TIMESTAMP', new Date().getTime().toString());
  Logger.log('Folder stored: ' + name + ' (ID: ' + id + ')');
  return { id, name };
}

function getSelectedArchiveFolder() {
  const userProps = PropertiesService.getUserProperties();
  const id = userProps.getProperty('ARCHIVE_FOLDER_ID') || '';
  const name = userProps.getProperty('ARCHIVE_FOLDER_NAME') || '';
  const timestamp = userProps.getProperty('ARCHIVE_FOLDER_TIMESTAMP') || '';
  
  Logger.log('Folder retrieved: ' + name + ' (ID: ' + id + ', Timestamp: ' + timestamp + ')');
  
  return {
    id: id,
    name: name,
    timestamp: timestamp
  };
}

function clearSelectedArchiveFolder() {
  const userProps = PropertiesService.getUserProperties();
  userProps.deleteProperty('ARCHIVE_FOLDER_ID');
  userProps.deleteProperty('ARCHIVE_FOLDER_NAME');
  userProps.deleteProperty('ARCHIVE_FOLDER_TIMESTAMP');
}

function getArchiveFolderUrl(folderId, destinationMode) {
  try {
    if (!folderId || destinationMode === 'root') {
      return null; // My Drive root - no specific folder
    }
    
    const folder = DriveApp.getFolderById(folderId);
    return folder.getUrl();
  } catch (e) {
    Logger.log('Error getting folder URL: ' + e.message);
    return null;
  }
}

function getLastArchivedFileUrl(configId) {
  try {
    const data = readConfigData_();
    const config = data.configs.find(c => c.id === configId);
    
    if (!config || !config.lastArchivedFileUrl) {
      return null;
    }
    
    // Verify the file still exists
    try {
      // Extract file ID from URL if needed
      let fileId = config.lastArchivedFileUrl;
      if (fileId.includes('/d/')) {
        fileId = fileId.split('/d/')[1].split('/')[0];
      }
      
      const file = DriveApp.getFileById(fileId);
      return file.getUrl();
    } catch (e) {
      // File no longer exists, return null
      Logger.log('Archived file no longer exists: ' + e.message);
      return null;
    }
  } catch (e) {
    Logger.log('Error getting last archived file URL: ' + e.message);
    return null;
  }
}

function getArchiveSheetColumns(ssId, sheetName) {
  try {
    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) return [];
    
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return [];
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    return headers.map(h => h.toString().trim()).filter(h => h !== "");
  } catch (e) {
    return [];
  }
}
