
/////////////////////////////////////////////////////////////////////Autorize/////////////////////////////////////////////////////////////////////

function doAutorizeScript() {
  const sheet_co = fetchSheetByName('Config');
  if (!sheet_co) {
    LogDebug("Sheet 'Config' not found.", 'MIN');
    return;
  }
  LogDebug(`Autorizing Script`, 'MIN');

  const cell = sheet_co.getRange("L2");
  cell.setBackground("#006600");                             // Dark Green (#006600)
  cell.setFontColor("#FFFFFF");                              // White Font (#FFFFFF)

  LogDebug("L2 cell color updated to dark green with white font.", 'MIN');
}

function checkAutorizeScript() {
  const sheet_co = fetchSheetByName('Config');

  const cell = sheet_co.getRange("L2");
  const bgColor = cell.getBackground();                     // Get background color
  const fontColor = cell.getFontColor();                    // Get font color

  const expectedBgColor = "#006600";
  const expectedFontColor = "#ffffff";                      // Note: Google Sheets may return lowercase

  const isMatch = (bgColor.toLowerCase() === expectedBgColor && fontColor.toLowerCase() === expectedFontColor);

  LogDebug(`L2 Background: ${bgColor}, Font: ${fontColor}`, 'MIN');
  LogDebug(`Match: ${isMatch ? "✅ Colors are correct" : "❌ Colors are incorrect"}`, 'MIN');

  return isMatch;
}

/////////////////////////////////////////////////////////////////////Triggers/////////////////////////////////////////////////////////////////////

// Central trigger definitions (single source of truth)
const triggerMap = {
  STOCK: [
    { fn: 'doSaveAllBasics',    cfgKey: TG1 },    // Basic Trigger Event
    { fn: 'doSaveAllFinancials', cfgKey: TG2 },    // Financial Trigger Event
    { fn: 'doSaveAllExtras',     cfgKey: TG3 },    // Extras Trigger Event
    { fn: 'doSettings',          cfgKey: TG4 },    // Settings Trigger Event
    { fn: 'doSaveAll',           cfgKey: TG5 },    // SaveAll Trigger Event
  ],
  BDR:   [{ fn: 'doSaveAllBasics', hour: 20 }],
  ETF:   [{ fn: 'doSaveAllBasics', hour: 20 }],
  ADR:   [{ fn: 'doSaveSWING',     hour: 20 }],
};

// Compare two sets for equality
function setsEqual(a, b) {
  if (a.size !== b.size) return false;
  for (let x of a) if (!b.has(x)) return false;
  return true;
}

// Check and reconcile project triggers
function doCheckTriggers() {
  const Class       = getConfigValue(IST, 'Config');
  const desiredList = triggerMap[Class] || [];

  // Only consider time-based triggers
  const existing    = ScriptApp.getProjectTriggers()
                        .filter(t => t.getEventType() === ScriptApp.EventType.CLOCK);
  const haveFns     = existing.map(t => t.getHandlerFunction());
  const wantFns     = desiredList.map(d => d.fn);

  if (!setsEqual(new Set(haveFns), new Set(wantFns))) {
    LogDebug(`Triggers mismatch (have:${haveFns.length}, want:${wantFns.length}). Rebuilding…`, 'MIN');
    doDeleteTriggers();
    doCreateTriggers();
    writeTriggersToSheet();
  } else {
    LogDebug(`Triggers up-to-date (${haveFns.length}).`, 'MIN');
  }
}

// Create triggers based on mapping
function doCreateTriggers() {
  const Class     = getConfigValue(IST, 'Config');
  const desired   = triggerMap[Class] || [];
  if (desired.length === 0) return;

  LogDebug(`📝 Creating ${desired.length} triggers for class '${Class}'`, 'MIN');

  desired.forEach(d => {
    // If cfgKey is provided, pull the configured hour
    const Hour = d.cfgKey
      ? getConfigValue(d.cfgKey, 'Config')
      : d.hour;

    ScriptApp.newTrigger(d.fn)
      .timeBased()
      .atHour(Hour)
      .everyDays(1)
      .create();

    LogDebug(` → ${d.fn} @ ${Hour}`, 'MIN');
  });
}

function doDeleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  if (triggers.length === 0) {
    LogDebug(`No triggers found to delete.`, 'MIN');
  }

  for (const trigger of triggers) {
    LogDebug(`⚠️ Deleting trigger: ${trigger.getHandlerFunction()} (ID: ${trigger.getUniqueId()})`, 'MIN');
    ScriptApp.deleteTrigger(trigger);
  }

  LogDebug(`All triggers deleted.`, 'MIN');
}


function getSheetTriggers() {
  const sheet_Triggers = ScriptApp.getProjectTriggers();

  return sheet_Triggers.length;
};

function getSheetTriggersHandle() {
  const triggers = ScriptApp.getProjectTriggers();
  const handlerFunctions = [];

  for (let i = 0; i < triggers.length; i++) {
    const funcName = triggers[i].getHandlerFunction();
    handlerFunctions.push(funcName);
  }
  return handlerFunctions;
}

function writeTriggersToSheet() {
  const sheet = fetchSheetByName('Config');
  if (!sheet) return;

  const triggers = getSheetTriggersHandle();
  const count    = getSheetTriggers();

  // Write count to L21
  sheet.getRange(21, 12).setValue(count);

  const startRow   = 24; // L24
  const startCol   = 12; // Column L
  const lastRow    = sheet.getLastRow();

  // Clear old handler names
  if (lastRow >= startRow) {
    sheet.getRange(startRow, startCol, lastRow - startRow + 1, 1).clearContent();
  }

  // Write handler names or a placeholder
  if (triggers.length > 0) {
    sheet.getRange(startRow, startCol, triggers.length, 1)
         .setValues(triggers.map(t => [t]));
  } else {
    sheet.getRange(startRow, startCol).setValue('No active triggers');
  }

  LogDebug(`🖋️ Wrote ${triggers.length} triggers and count ${count} to Config`, 'MIN');
}

/////////////////////////////////////////////////////////////////////IMPORT FUNCTIONS/////////////////////////////////////////////////////////////////////
