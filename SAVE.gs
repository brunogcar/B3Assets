/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doSaveAll() {
  Logger.log(SNAME(2));

  processSheets([BLC, DRE, FLC, DVA], doCheckDATA, doSaveData);

  doSaveShares();
  doProventos();

  processSheets([OPCOES, BTC, TERMO], doCheckDATA, doSaveSheet);
  
  SpreadsheetApp.flush();

  processSheets([FUND, SWING_4, SWING_12, SWING_52], doCheckDATA, doSaveSheet);

  processSheets([
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3,
    RIGHT_1, RIGHT_2,
    RECEIPT_9, RECEIPT_10,
    WARRANT_11, WARRANT_12, WARRANT_13,
    BLOCK
  ], doCheckDATA, doSaveSheet);

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}

function processSheets(sheetNames, checkCallback, saveCallback) {
  var sheetsToSave = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Gather sheets that are available and pass the check.
  sheetNames.forEach(function(SheetName) {
    var sheet = ss.getSheetByName(SheetName);
    if (sheet) {
      var availableData = checkCallback(SheetName);
      if (availableData === "TRUE") {
        sheetsToSave.push(SheetName);
      }
    } else {
      Logger.log(`ERROR SAVE: ${SheetName} - Does not exist`);
    }
  });
  
  var totalSheets = sheetsToSave.length;
  if (totalSheets > 0) {
    SpreadsheetApp.flush();
    let count = 0;
    sheetsToSave.forEach(function(SheetName) {
      count++;
      const progress = Math.round((count / totalSheets) * 100);
      Logger.log(`[${count}/${totalSheets}] (${progress}%) saving ${SheetName}...`);
      try {
        saveCallback(SheetName);
        Logger.log(`[${count}/${totalSheets}] (${progress}%) ${SheetName} saved successfully`);
      } catch (error) {
        Logger.log(`[${count}/${totalSheets}] (${progress}%) Error saving ${SheetName}: ${error}`);
      }
    });
  } else {
    Logger.log(`No valid data found. Skipping save operation.`);
  }
}

/////////////////////////////////////////////////////////////////////Individual/////////////////////////////////////////////////////////////////////

//-------------------------------------------------------------------SHEETS-------------------------------------------------------------------//
function doSaveAllSheets() {
  Logger.log(SNAME(2));

  doSaveShares();
  doProventos();

  // Process first group (OPCOES, BTC, TERMO)
  processSheets([OPCOES, BTC, TERMO], doCheckDATA, doSaveSheet);
  
  SpreadsheetApp.flush();

  // Process second group (FUND, SWING_4, SWING_12, SWING_52)
  processSheets([FUND, SWING_4, SWING_12, SWING_52], doCheckDATA, doSaveSheet);
  
  doExportExtras();
  doExportDatas();
  
  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}
//-------------------------------------------------------------------EXTRAS-------------------------------------------------------------------//
function doSaveAllExtras() {
  Logger.log(SNAME(2));

  processSheets([
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3,
    RIGHT_1, RIGHT_2,
    RECEIPT_9, RECEIPT_10,
    WARRANT_11, WARRANT_12, WARRANT_13,
    BLOCK
  ], doCheckDATA, doSaveSheet);

  doSaveShares();
  doProventos();

  doExportSheets();
  doExportDatas();

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}
//-------------------------------------------------------------------DATAS-------------------------------------------------------------------//
function doSaveAllDatas() {
  Logger.log(SNAME(2));

  processSheets([BLC, DRE, FLC, DVA], doCheckDATA, doSaveData);

  doSaveShares();
  doProventos();

  doExportSheets();
  doExportExtras();

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

  // could add addicional checks with && "and" || "or"
  // can get SheetName from ss.getName() as well

/////////////////////////////////////////////////////////////////////SHEETS/////////////////////////////////////////////////////////////////////

function doSaveSheets() {
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND];

  SheetNames.forEach(SheetName => {
    try { doSaveSheet(SheetName); }
    catch (error) { Logger.error(`Error saving sheet ${SheetName}: ${error}`); }
  });
}

/////////////////////////////////////////////////////////////////////EXTRAS/////////////////////////////////////////////////////////////////////

function doSaveExtras() {
  const SheetNames = [FUTURE, FUTURE_1, FUTURE_2, FUTURE_3, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13, BLOCK];

  SheetNames.forEach(SheetName => {
    try { doSaveSheet(SheetName); }
    catch (error) { Logger.error(`Error saving sheet ${SheetName}: ${error}`); }
  });
}

/////////////////////////////////////////////////////////////////////DATAS/////////////////////////////////////////////////////////////////////

function doSaveDatas()
{
  const SheetNames = [BLC, DRE, FLC, DVA];                             //Balanço, Resultado, Fluxo and Valor are saved after parent SheetNames

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_u = ss.getSheetByName(`UPDATE`);                         // UPDATE sheet

    var ACTV = sheet_u.getRange(`B3`).getValue();
    var SOMA = sheet_u.getRange(`K8`).getValue();

  if (!ACTV || (ACTV && ((SOMA >= 450 && SOMA <= 460) || (SOMA == 0 || SOMA > 125000))))
  {
    SheetNames.forEach(SheetName => {
      try { doSaveData(SheetName); }
      catch (error) { Logger.error(`Error saving sheet ${SheetName}: ${error}`); }
    });
  }
}

/////////////////////////////////////////////////////////////////////SAVE/////////////////////////////////////////////////////////////////////