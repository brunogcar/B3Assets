
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
  if (!sheet_co) {
    LogDebug("Sheet 'Config' not found.", 'MIN');
    return false;
  }

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

function doCheckTriggers() {
  const Class = getConfigValue(IST, 'Config');                                      // IST = Is Stock?

  var Triggers = ScriptApp.getProjectTriggers().length;

  LogDebug(`Number of existing triggers: ${Triggers}`, 'MIN');

  if (Class == 'STOCK')
  {
    if (Triggers == 0)
    {
      LogDebug("No triggers found. Creating new triggers...", 'MIN');
      doCreateTriggers();
    }
    else if (Triggers > 0 && Triggers < 5)
    {
      LogDebug("Found 1-4 triggers. Deleting and creating new triggers...", 'MIN');
      doDeleteTriggers();
      doCreateTriggers();
    }
    else if (Triggers > 5)
    {
      LogDebug("Found more than 5 triggers. Deleting and creating new triggers...", 'MIN');
      doDeleteTriggers();
      doCreateTriggers();
    }
  }
  else
  {
    if (Triggers == 0)
    {
      LogDebug("No triggers found. Creating new triggers...", 'MIN');
      doCreateTriggers();
    }
    else if (Triggers > 1)
    {
      LogDebug("Found more than 1 triggers. Deleting and creating new triggers...", 'MIN');
      doDeleteTriggers();
      doCreateTriggers();
    }
  }
};

function doCreateTriggers() {
  const Class = getConfigValue(IST, 'Config');                                      // IST = Is Stock?

  // Check existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  let shouldCreateTrigger = true;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getEventType() === ScriptApp.EventType.CLOCK) {
      shouldCreateTrigger = false;
      break;
    }
  }

  if (!shouldCreateTrigger) return;

  if (Class === 'STOCK') {
    LogDebug("Creating new triggers...", 'MIN');

    const Hour_1 = getConfigValue(TG1, 'Config');                                    // Basic Trigger Event
    const Hour_2 = getConfigValue(TG2, 'Config');                                    // Financial Trigger Event
    const Hour_3 = getConfigValue(TG3, 'Config');                                    // Extras Trigger Event
    const Hour_4 = getConfigValue(TG4, 'Config');                                    // Settings Trigger Event
    const Hour_5 = getConfigValue(TG5, 'Config');                                    // SaveAll Trigger Event

    ScriptApp.newTrigger("doSaveAllBasics")
      .timeBased().atHour(Hour_1).everyDays(1).create();

    ScriptApp.newTrigger("doSaveAllFinancials")
      .timeBased().atHour(Hour_2).everyDays(1).create();

    ScriptApp.newTrigger("doSaveAllExtras")
      .timeBased().atHour(Hour_3).everyDays(1).create();

    ScriptApp.newTrigger("doSettings")
      .timeBased().atHour(Hour_4).everyDays(1).create();

    ScriptApp.newTrigger("doSaveAll")
      .timeBased().atHour(Hour_5).everyDays(1).create();

  } else if (Class === 'BDR' || Class === 'ETF') {
    ScriptApp.newTrigger("doSaveAllBasics")
      .timeBased().atHour(20).everyDays(1).create();

  } else if (Class === 'ADR') {
    ScriptApp.newTrigger("doSaveSWING")
      .timeBased().atHour(20).everyDays(1).create();
  }
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
  const sheet = fetchSheetByName("Config");
  if (!sheet) return;

  const triggers = getSheetTriggersHandle();
  const startRow = 24; // L24
  const startColumn = 12; // Column "L" = 12th column

  // Clear old values from L24 downward
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, startColumn, lastRow - startRow + 1, 1).clearContent();
  }

  // Write triggers if available
  if (triggers.length > 0) {
    sheet.getRange(startRow, startColumn, triggers.length, 1).setValues(triggers.map(t => [t]));
  } else {
    sheet.getRange(startRow, startColumn).setValue("No active triggers");
  }
  LogDebug(`Wrote ${triggers.length} triggers to Config`, 'MIN');
}

function doDeleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  if (triggers.length === 0) {
    LogDebug(`No triggers found to delete.`, 'MIN');
  }

  for (const trigger of triggers) {
    LogDebug(`Deleting trigger: ${trigger.getHandlerFunction()} (ID: ${trigger.getUniqueId()})`, 'MIN');
    ScriptApp.deleteTrigger(trigger);
  }

  LogDebug(`All triggers deleted.`, 'MIN');
}

/////////////////////////////////////////////////////////////////////IMPORT FUNCTIONS/////////////////////////////////////////////////////////////////////
