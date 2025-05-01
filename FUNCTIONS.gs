//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////Helper functions/////////////////////////////////////////////////////////////////////
/**
 * IMPORTANT:
 * Because this function already calls SpreadsheetApp.getActiveSpreadsheet(), 
 * you do NOT need to manually declare `var ss = SpreadsheetApp.getActiveSpreadsheet()` 
 * in any function that only uses this to access sheets.
 *
 * @param {string} SheetName - The exact name of the sheet to fetch.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null} - The Sheet object if found; otherwise, null.
 */
function fetchSheetByName(SheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
  if (!sheet) {
    Logger.log(`Sheet not found: ${SheetName}`);
    return null;
  }
  return sheet;
}

/**
 * Retrieves a configuration value based on the provided acronym and source.
 *
 * This function checks the 'Settings' sheet first (if enabled),
 * then falls back to the 'Config' sheet if needed. You can explicitly
 * force which sheet to check by passing "Settings" or "Config" as the source.
 *
 * If the value is "DEFAULT" or one of the error values, it tries to get
 * the value from the other sheet (when using default behavior).
 *
 * @param {string} Acronym - A named cell reference (e.g., "ETE", "ITR") representing a config setting.
 * @param {string} [source="Both"] - Optional source: "Settings", "Config", or "Both".
 * @returns {string|null} The retrieved value, or null if not found or invalid.
 *
 * Usage:
 *   const val1 = getConfigValue(ETE);                // Default behavior (Settings -> Config)
 *   const val2 = getConfigValue(ETE, 'Settings');    // Only from Settings
 *   const val3 = getConfigValue(ETE, 'Config');      // Only from Config
 */
function getConfigValue(Acronym, Source = 'Both') {
  const sheet_se = (Source !== "Config") ? fetchSheetByName('Settings') : null;
  const sheet_co = (Source !== "Settings") ? fetchSheetByName('Config') : null;

  if (!sheet_se || !sheet_co){
    Logger.log('Settings or Config sheet not found');
    return null;
  }

  let Value = null;

  // If Source is Settings or Both, try Settings first
  if (sheet_se) {
    try {
      Value = sheet_se.getRange(Acronym).getDisplayValue().trim();
      if (!Value || Value === "DEFAULT" || ErrorValues.includes(Value)) {
        Value = null;                                                        // fallback
      } else if (Source === "Settings") {
        return Value;                                                        // shortcut if only using Settings
      }
    } catch (e) {
      Logger.log(`Acronym ${Acronym} not found in Settings`);
    }
  }

  // If Source is Config or fallback from Settings
  if (!Value && sheet_co) {
    try {
      Value = sheet_co.getRange(Acronym).getDisplayValue().trim();
      if (!Value || ErrorValues.includes(Value)) {
        Value = null;
      }
    } catch (e) {
      Logger.log(`Acronym ${Acronym} not found in Config`);
    }
  }

  return Value;
}

/////////////////////////////////////////////////////////////////////Compare arrays/////////////////////////////////////////////////////////////////////

function arraysAreEqual(arr1, arr2) {
  if (arr1.length !== arr2.length) return false;
  for (let i = 0; i < arr1.length; i++) {
    if (arr1[i].length !== arr2[i].length) return false;
    for (let j = 0; j < arr1[i].length; j++) {
      if (arr1[i][j] !== arr2[i][j]) return false;
    }
  }
  return true;
}

/////////////////////////////////////////////////////////////////////Settings/////////////////////////////////////////////////////////////////////

function doSettings() {
  const sheet_co = fetchSheetByName('Config');
  const Class = getConfigValue(IST, 'Config');                                          // IST = Is Stock?
  const sheet_sr = fetchSheetByName('Settings');
  var Activate  = sheet_sr.getRange(ACT).getDisplayValue();                             // ACT = Activate

  if (Class == 'STOCK') 
  {
    if ( Activate == "TRUE")                                              // TRUE
    {
      var True = sheet_sr.getRange(TRU).getDisplayValue();                               // TRU = True 

      if ( True == 'SAVE')                                                // SAVE
      {
        var Save = sheet_sr.getRange(SAV).getDisplayValue();                             // SAV = SAVE

        if ( Save == 'SHEETS') { doSaveAllBasics(); }
        if ( Save == 'EXTRAS') { doSaveAllExtras(); }
        if ( Save == 'DATAS')  { doSaveAllFinancials(); }
        if ( Save == 'ALL')    { doSaveAll(); }
        if ( Save == 'INDIVIDUAL')
          { 
            var Individual = sheet_sr.getRange(IND).getDisplayValue();                   // IND = INDIVIDUAL

            if ( Individual == 'SWING')  { doSaveSWING(); }
            if ( Individual == 'OPCOES') { doSaveBasic(OPCOES); }
            if ( Individual == 'BTC')    { doSaveBasic(BTC); }
            if ( Individual == 'TERMO')  { doSaveBasic(TERMO); }
            if ( Individual == 'FUND')   { doSaveBasic(FUND); }
            if ( Individual == 'FUTURE') { doSaveBasic(FUTURE); }
          }
      }
      if ( True == 'EXPORT') {doExportAll(); }
      if ( True == 'OTHER')                                               // OTHER
      {
        var Other = sheet_sr.getRange(EXT).getDisplayValue();                             // EXT = Extra

        if ( Other == 'ZEROS')    { doCleanZeros(); }
        if ( Other == 'TRIGGERS') { doCheckTriggers(); }
        if ( Other == 'CHECK')    { doCheckDATAS(); }                                    // Check and hide or show Sheets
        if ( Other == 'PROV')     { doSaveProventos(); }
        if ( Other == 'SHARES')   { doSaveShares(); }
        if ( Other == 'RIGHTS')   { doRestoreRight(); }
      }
    }
  }
};

/////////////////////////////////////////////////////////////////////RETIRE/////////////////////////////////////////////////////////////////////

function doRetire() {
  copypasteSheets();
  doClearSheetID();
  doClearExportAll();
  doDeleteSheets();
  moveSpreadsheetToARQUIVO();

  doDeleteTriggers();
  revokeOwnAccess();
};

function copypasteSheets() {
  const SheetNames = [
    'Index', 'Info', 'Comunicados', 'Prov', 'Preço', 'Cotações', 'OPT', 'DATA',
    'Value', 'Balanco', 'Resultado', 'Fluxo', 'Valor'
  ];

  SheetNames.forEach(Name => {
    const sheet = fetchSheetByName(Name);
    if (sheet) {
      const range = sheet.getDataRange();
      range.copyTo(range, { contentsOnly: true });
    }
  });
}

function doDeleteSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SheetNames = [
    'Balanço Ativo', 
    'Balanço Passivo', 
    'Demonstração', 
    'Fluxo de Caixa', 
    'Demonstração do Valor Adicionado'
  ];

  SheetNames.forEach(Name => {
    const sheet = fetchSheetByName(Name);
    if (sheet) {
      try {
        ss.deleteSheet(sheet);
        Logger.log(`Sheet deleted: ${Name}`);
      } catch (error) {
        Logger.log(`Error deleting sheet "${Name}": ${error}`);
      }
    }
  });
}



function moveSpreadsheetToFolder(folderName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());

  const folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) {
    Logger.log(`Folder "${folderName}" not found.`);
    return;
  }

  const folder = folders.next();
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  Logger.log(`Spreadsheet moved to ${folderName}`);
}

function moveSpreadsheetToARQUIVO() {
  moveSpreadsheetToFolder("-=ARQUIVO=-");
}

/////////////////////////////////////////////////////////////////////DELETE/////////////////////////////////////////////////////////////////////

function doDelete() {
  doDeleteTriggers();
  moveSpreadsheetToBACKUP();
  revokeOwnAccess();
}

function revokeOwnAccess() {
  // Invalidate the script's authorization
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  if (authInfo) {
    ScriptApp.invalidateAuth();
    Logger.log('Script access revoked successfully.');
  } else {
    Logger.log('Script is not authorized or access has already been revoked.');
  }
}

function moveSpreadsheetToBACKUP() {
  moveSpreadsheetToFolder("-=BACKUP=-");
}

function doDeleteSpreadsheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fileId = ss.getId();
  
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    Logger.log('Spreadsheet deleted successfully.');
  } catch (error) {
    Logger.log(`Error deleting spreadsheet: ${error}`);
  }
}

/////////////////////////////////////////////////////////////////////Name/////////////////////////////////////////////////////////////////////

function SNAME(option) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  switch (option) {
    case 0: {                                                 // Active sheet name
      const activeSheet = ss.getActiveSheet();
      return activeSheet ? activeSheet.getName() : "#N/A";
    }
    case 1:                                                   // All sheet names
      return ss.getSheets().map(sheet => sheet.getName());

    case 2:                                                   // Spreadsheet name
      return ss.getName();

    case 3: {                                                 // Extract version from spreadsheet name (after hyphen)
      const Name = ss.getName();
      const match = Name.match(/-(.*)/);
      return match ? match[1].trim() : "No version found";
    }

    default:
      return "#N/A";
  }
}

/////////////////////////////////////////////////////////////////////CLEAN SHEETS/////////////////////////////////////////////////////////////////////

function doCleanZeros() {
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUTURE, FUND];

  SheetNames.forEach(SheetName => {
    const sheet = fetchSheetByName(SheetName);
    if (!sheet) return;

    const range = sheet.getRange(5, 1, sheet.getLastRow() - 4, sheet.getLastColumn());
    const Data = range.getValues();
    let Modified = false;

    for (let i = 0; i < Data.length; i++) {
      for (let j = 0; j < Data[i].length; j++) {
        if (Data[i][j] === 0) {
          Data[i][j] = "";
          Modified = true;
        }
      }
    }

    if (Modified) {
      range.setValues(Data);
      Logger.log(`Zeros cleaned in sheet: ${SheetName}`);
    }
  });
}

/////////////////////////////////////////////////////////////////////reverse/////////////////////////////////////////////////////////////////////

function reverseColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const SheetName = sheet.getName();
  const active = fetchSheetByName(SheetName);
  if (!active) return;

  const range = active.getRange(1, 4, active.getLastRow(), active.getLastColumn() - 3);
  const values = range.getValues();

  const reversedValues = values.map(row => row.reverse());
  range.setValues(reversedValues);
  Logger.log(`Columns reversed in sheet: ${SheetName}`);
}

function reverseRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const SheetName = sheet.getName();
  const active = fetchSheetByName(SheetName);
  if (!active) return;

  const range = active.getRange(5, 1, active.getLastRow() - 4, active.getLastColumn());
  const values = range.getValues();

  values.reverse();
  range.setValues(values);
  Logger.log(`Rows reversed in sheet: ${SheetName}`);
}

/////////////////////////////////////////////////////////////////////RESTORE Functions/////////////////////////////////////////////////////////////////////

function doRestoreFundExport() {
  const sheet_co = fetchSheetByName('Config');

  var Value = '=IF(OR(AND(Fund!A5="";Fund!A1=""); L18<>"STOCK"); FALSE;TRUE)';                              

    sheet_co.getRange(EFU).setValue(Value);                                 // EFU = Export to Fund 
}

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////