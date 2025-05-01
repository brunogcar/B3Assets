/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doSaveAll() {
  Logger.log(SNAME(2));

  SpreadsheetApp.flush();

  processSave([BLC, DRE, FLC, DVA], doCheckDATA, doSaveData);

  doSaveShares();
  doProventos();

  processSave([OPCOES, BTC, TERMO], doCheckDATA, doSaveSheet);
  
  processSave([FUND, SWING_4, SWING_12, SWING_52], doCheckDATA, doSaveSheet);

  processSave([
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

/////////////////////////////////////////////////////////////////////Individual/////////////////////////////////////////////////////////////////////

//-------------------------------------------------------------------SHEETS-------------------------------------------------------------------//
function doSaveAllSheets() {
  Logger.log(SNAME(2));

  SpreadsheetApp.flush();

  processSave([OPCOES, BTC, TERMO], doCheckDATA, doSaveSheet);
  
  processSave([FUND, SWING_4, SWING_12, SWING_52], doCheckDATA, doSaveSheet);
  
  doSaveShares();
  doExportProventos();

  doExportExtras();
  doExportDatas();

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}
//-------------------------------------------------------------------EXTRAS-------------------------------------------------------------------//
function doSaveAllExtras() {
  Logger.log(SNAME(2));

  processSave([
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3,
    RIGHT_1, RIGHT_2,
    RECEIPT_9, RECEIPT_10,
    WARRANT_11, WARRANT_12, WARRANT_13,
    BLOCK
  ], doCheckDATA, doSaveSheet);

  SpreadsheetApp.flush();

  doSaveShares();
  doExportProventos();

  doExportSheets();
  doExportDatas();

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}
//-------------------------------------------------------------------DATAS-------------------------------------------------------------------//
function doSaveAllDatas() {
  Logger.log(SNAME(2));

  SpreadsheetApp.flush();

  processSave([BLC, DRE, FLC, DVA], doCheckDATA, doSaveData);

  doSaveShares();
  doExportProventos();

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

  const sheet_up = fetchSheetByName(`UPDATE`);                         // UPDATE sheet

    var ACTV = sheet_up.getRange(`B3`).getValue();
    var SOMA = sheet_up.getRange(`K8`).getValue();

  if (!ACTV || (ACTV && ((SOMA >= 450 && SOMA <= 460) || (SOMA == 0 || SOMA > 125000))))
  {
    SheetNames.forEach(SheetName => {
      try { doSaveData(SheetName); }
      catch (error) { Logger.error(`Error saving sheet ${SheetName}: ${error}`); }
    });
  }
}

/////////////////////////////////////////////////////////////////////SAVE/////////////////////////////////////////////////////////////////////