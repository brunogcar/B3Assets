/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doSaveAll() {
  Logger.log(SNAME(2));

  SpreadsheetApp.flush();

  doSaveGroup([BLC, DRE, FLC, DVA], doCheckDATA, doSaveFinancial);

  doSaveShares();
  doProventos();

  doSaveGroup([OPCOES, BTC, TERMO], doCheckDATA, doSaveBasic);
  
  doSaveGroup([FUND, SWING_4, SWING_12, SWING_52], doCheckDATA, doSaveBasic);

  doSaveGroup([
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3,
    RIGHT_1, RIGHT_2,
    RECEIPT_9, RECEIPT_10,
    WARRANT_11, WARRANT_12, WARRANT_13,
    BLOCK
  ], doCheckDATA, doSaveBasic);

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}

/////////////////////////////////////////////////////////////////////Individual/////////////////////////////////////////////////////////////////////

//-------------------------------------------------------------------BASICS-------------------------------------------------------------------//
function doSaveAllBasics() {
  Logger.log(SNAME(2));

  SpreadsheetApp.flush();

  doSaveGroup([OPCOES, BTC, TERMO], doCheckDATA, doSaveBasic);
  
  doSaveGroup([FUND, SWING_4, SWING_12, SWING_52], doCheckDATA, doSaveBasic);
  
  doSaveShares();
  doExportProventos();

  doExportExtras();
  doExportFinancials();

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}
//-------------------------------------------------------------------EXTRAS-------------------------------------------------------------------//
function doSaveAllExtras() {
  Logger.log(SNAME(2));

  doSaveGroup([
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3,
    RIGHT_1, RIGHT_2,
    RECEIPT_9, RECEIPT_10,
    WARRANT_11, WARRANT_12, WARRANT_13,
    BLOCK
  ], doCheckDATA, doSaveBasic);

  SpreadsheetApp.flush();

  doSaveShares();
  doExportProventos();

  doExportBasics();
  doExportFinancials();

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}
//-------------------------------------------------------------------DATAS-------------------------------------------------------------------//
function doSaveAllFinancials() {
  Logger.log(SNAME(2));

  SpreadsheetApp.flush();

  doSaveGroup([BLC, DRE, FLC, DVA], doCheckDATA, doSaveFinancial);

  doSaveShares();
  doExportProventos();

  doExportBasics();
  doExportExtras();

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

  // could add addicional checks with && "and" || "or"
  // can get SheetName from ss.getName() as well

/////////////////////////////////////////////////////////////////////BASICS/////////////////////////////////////////////////////////////////////

function doSaveBasics() {
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

function doSaveFinancials()
{
  const SheetNames = [BLC, DRE, FLC, DVA];                             //Balanço, Resultado, Fluxo and Valor are saved after parent SheetNames

  const sheet_up = fetchSheetByName(`UPDATE`);                         // UPDATE sheet

    var ACTV = sheet_up.getRange(`B3`).getValue();
    var SOMA = sheet_up.getRange(`K8`).getValue();

  if (!ACTV || (ACTV && ((SOMA >= 450 && SOMA <= 460) || (SOMA == 0 || SOMA > 125000))))
  {
    SheetNames.forEach(SheetName => {
      try { doSaveFinancial(SheetName); }
      catch (error) { Logger.error(`Error saving sheet ${SheetName}: ${error}`); }
    });
  }
}

/////////////////////////////////////////////////////////////////////SAVE/////////////////////////////////////////////////////////////////////