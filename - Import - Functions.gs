
/////////////////////////////////////////////////////////////////////Autorize/////////////////////////////////////////////////////////////////////

function doAutorizeScript() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');  
  if (!sheet_co) {
    Logger.log("Sheet 'Config' not found.");
    return;
  }
    Logger.log(`Autorizing Script`);

  const cell = sheet_co.getRange("L2");
  cell.setBackground("#006600"); // Dark Green (#006600)
  cell.setFontColor("#FFFFFF"); // White Font (#FFFFFF)

  Logger.log("L2 cell color updated to dark green with white font.");
}

function checkAutorizeScript() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');  
  if (!sheet_co) {
    Logger.log("Sheet 'Config' not found.");
    return false;
  }

  const cell = sheet_co.getRange("L2");
  const bgColor = cell.getBackground();  // Get background color
  const fontColor = cell.getFontColor(); // Get font color

  const expectedBgColor = "#006600";
  const expectedFontColor = "#ffffff"; // Note: Google Sheets may return lowercase

  const isMatch = (bgColor.toLowerCase() === expectedBgColor && fontColor.toLowerCase() === expectedFontColor);

  Logger.log(`L2 Background: ${bgColor}, Font: ${fontColor}`);
  Logger.log(`Match: ${isMatch ? "✅ Colors are correct" : "❌ Colors are incorrect"}`);

  return isMatch;
}
/////////////////////////////////////////////////////////////////////Triggers/////////////////////////////////////////////////////////////////////

function doCheckTriggers() 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');
  var Class = sheet_co.getRange(IST).getDisplayValue();                             // IST = Is Stock? 

  var Triggers = ScriptApp.getProjectTriggers().length;

  Logger.log(`Number of existing triggers: ${Triggers}`);

  if (Class == 'STOCK') 
  {
    if (Triggers == 0) 
    {
      Logger.log("No triggers found. Creating new triggers...");
      doCreateTriggers();
    } 
    else if (Triggers > 0 && Triggers < 5) 
    {
      Logger.log("Found 1-4 triggers. Deleting and creating new triggers...");
      doDeleteTriggers();
      doCreateTriggers();
    } 
    else if (Triggers > 5) 
    {
      Logger.log("Found more than 5 triggers. Deleting and creating new triggers...");
      doDeleteTriggers();
      doCreateTriggers();
    }
  }
  else 
  {
    if (Triggers == 0) 
    {
      Logger.log("No triggers found. Creating new triggers...");
      doCreateTriggers();
    } 
    else if (Triggers > 1) 
    {
      Logger.log("Found more than 1 triggers. Deleting and creating new triggers...");
      doDeleteTriggers();
      doCreateTriggers();
    }
  }
};

function doCreateTriggers()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');
  var Class = sheet_co.getRange(IST).getDisplayValue();                             // IST = Is Stock? 

  var triggers = ScriptApp.getProjectTriggers();
  var shouldCreateTrigger = true;
  triggers.forEach(function (trigger)
  {
    if(trigger.getEventType() === ScriptApp.EventType.CLOCK ) 
    {
      shouldCreateTrigger = false; 
    }
  });
  
  if(shouldCreateTrigger)
  {
    if (Class == 'STOCK') 
    {
      Logger.log("Creating new triggers...");

      var Hour_1 = sheet_co.getRange(TG1).getValue();                              // TG1 = Sheet Trigger Event
      var Hour_2 = sheet_co.getRange(TG2).getValue();                              // TG2 = Data  Trigger Event
      var Hour_3 = sheet_co.getRange(TG3).getValue();                              // TG3 = Extra Trigger Event
      var Hour_4 = sheet_co.getRange(TG4).getValue();                              // TG4 = Settings Trigger Event
      var Hour_5 = sheet_co.getRange(TG5).getValue();                              // TG5 = SaveAll Trigger Event

      ScriptApp.newTrigger("doSaveAllSheets")
       .timeBased()
       .atHour(Hour_1)
       .everyDays(1)
       .create();

      ScriptApp.newTrigger("doSaveAllDatas")
       .timeBased()
       .atHour(Hour_2)
       .everyDays(1)
       .create();

      ScriptApp.newTrigger("doSaveAllExtras")
       .timeBased()
       .atHour(Hour_3)
       .everyDays(1)
       .create();

      ScriptApp.newTrigger("doSettings")
       .timeBased()
       .atHour(Hour_4)
       .everyDays(1)
       .create();

        ScriptApp.newTrigger("doSaveAll")
       .timeBased()
       .atHour(Hour_5)
       .everyDays(1)
       .create();
    }
    else if (Class == 'BDR' || Class == 'ETF') 
    {
      ScriptApp.newTrigger("doSaveAllSheets")                                   // Change to not stock settings, to not export, etc
       .timeBased()
       .atHour(20)
       .everyDays(1)
       .create();
    }
    else if (Class == 'ADR') 
    {
      ScriptApp.newTrigger("doSaveSWING")                                   // Change to not stock settings, to not export, etc
       .timeBased()
       .atHour(20)
       .everyDays(1)
       .create();
    }
  }
}

function getSheetTriggers()
{
  const sheet_Triggers = ScriptApp.getProjectTriggers();

  return sheet_Triggers.length;
};

function getSheetTriggersHandle() 
{
  const triggers = ScriptApp.getProjectTriggers();
  const handlerFunctions = [];

  for (let i = 0; i < triggers.length; i++) {
    const funcName = triggers[i].getHandlerFunction();
    handlerFunctions.push(funcName);
  }
  return handlerFunctions;
}

function writeTriggersToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = fetchSheetByName("Config");

  if (!sheet) {
    Logger.log("Sheet 'Config' not found.");
    return;
  }

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

  Logger.log(`Wrote ${triggers.length} triggers to Config`);
}

function doDeleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  if (triggers.length === 0) {
    Logger.log("No triggers found to delete.");
  }

  for (const trigger of triggers) {
    Logger.log(`Deleting trigger: ${trigger.getHandlerFunction()} (ID: ${trigger.getUniqueId()})`);
    ScriptApp.deleteTrigger(trigger);
  }

  Logger.log("All triggers deleted.");
}


/////////////////////////////////////////////////////////////////////IMPORT FUNCTIONS/////////////////////////////////////////////////////////////////////