/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doSaveAll() {
  LogDebug(SNAME(2), 'MIN');

  SpreadsheetApp.flush();

  doSaveGroup(SheetsFinancial, doCheckDATA, doSaveFinancial);

  doSaveShares();
  doProventos();

  doSaveGroup([...SheetsBasic,...SheetsExtra], doCheckDATA, doSaveBasic);

  doIsFormula();
  doDisableSheets();
  doCheckTriggers();
}

/////////////////////////////////////////////////////////////////////Individual/////////////////////////////////////////////////////////////////////

//-------------------------------------------------------------------BASICS-------------------------------------------------------------------//
function doSaveAllBasics() {
  LogDebug(SNAME(2), 'MIN');

  SpreadsheetApp.flush();

  doSaveGroup(SheetsBasic, doCheckDATA, doSaveBasic);

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
  LogDebug(SNAME(2), 'MIN');

  doSaveGroup(SheetsExtra, doCheckDATA, doSaveBasic);

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
  LogDebug(SNAME(2), 'MIN');

  SpreadsheetApp.flush();

  doSaveGroup(SheetsFinancial, doCheckDATA, doSaveFinancial);

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
  const SheetNames = SheetsBasic;

  for (let i = 0; i < SheetNames.length; i++) {
    const SheetName = SheetNames[i];
    try { doSaveBasic(SheetName); }
    catch (error) { Logger.log(`Error saving: ${SheetName}: ${error}`); }
  }
}


/////////////////////////////////////////////////////////////////////EXTRAS/////////////////////////////////////////////////////////////////////

function doSaveExtras() {
  const SheetNames = SheetsExtra;

  for (let i = 0; i < SheetNames.length; i++) {
    const SheetName = SheetNames[i];
    try { doSaveBasic(SheetName); }
    catch (error) { Logger.log(`Error saving: ${SheetName}: ${error}`); }
  }
}

/////////////////////////////////////////////////////////////////////DATAS/////////////////////////////////////////////////////////////////////

function doSaveFinancials() {
  const SheetNames = SheetsFinancial;                             //Balanço, Resultado, Fluxo and Valor are saved after parent SheetNames

  const sheet_up = getSheet(`UPDATE`);
  const ACTV = sheet_up.getRange(`B3`).getValue();
  const SOMA = sheet_up.getRange(`K8`).getValue();

  if (!ACTV || (ACTV && ((SOMA >= 450 && SOMA <= 460) || (SOMA === 0 || SOMA > 125000)))) {
    for (let i = 0; i < SheetNames.length; i++) {
      const SheetName = SheetNames[i];
      try { doSaveFinancial(SheetName); }
      catch (error) { Logger.log(`Error saving: ${SheetName}: ${error}`); }
    }
  }
}

/////////////////////////////////////////////////////////////////////SAVE/////////////////////////////////////////////////////////////////////
