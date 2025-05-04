//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////Helper functions/////////////////////////////////////////////////////////////////////

/**
 * Generic batch‐runner for sheet operations (edit/export/import).
 *
 * @param {string[]} SheetNames         List of sheet‐name constants.
 * @param {function(string):void} fn    Operation to perform on each sheet.
 * @param {string} actionLabel          Verb in gerund form ("Editing", "Exporting", "Importing").
 * @param {string} resultLabel          Past‐tense for summary ("edited", "exported", "imported").
 * @param {string} groupLabel           Descriptor for logging ("basic", "extra", "financial", etc.).
 *
 * Behavior:
 * - If a sheet does not exist, logs an error and skips it.
 * - If no sheets have data (`totalSheets === 0`), logs a “skipping” message.
 * - Otherwise, for each sheet:
 *    • Logs `[i/N] (P%) action <SheetName>...`
 *    • Calls `fn(SheetName)` inside a try/catch
 *    • On success or error, logs the outcome **only** if DEBUG is `"TRUE"`.
 */
function _doGroup(SheetNames, fn, actionLabel, resultLabel, groupLabel) {
  const totalSheets = SheetNames.length;
  let count = 0;

  const DEBUG    = getConfigValue(DBG, 'Config');               // DBG = Debug Mode

  if (DEBUG == "TRUE") {
    Logger.log(`Starting ${actionLabel.toLowerCase()} of ${totalSheets} ${groupLabel} sheets...`);
  }

  for (let i = 0; i < totalSheets; i++) {
    const SheetName = SheetNames[i];
    count++;
    const progress = Math.round((count / totalSheets) * 100);

    if (DEBUG == "TRUE") {
      Logger.log(`[${count}/${totalSheets}] (${progress}%) ${actionLabel} ${SheetName}...`);
    }

    try {
      fn(SheetName);
      if (DEBUG == "TRUE") {
        Logger.log(`[${count}/${totalSheets}] (${progress}%) ${SheetName} ${resultLabel} successfully`);
      }
    } catch (error) {
      if (DEBUG == "TRUE") {
        Logger.log(`[${count}/${totalSheets}] (${progress}%) Error ${actionLabel.toLowerCase()} ${SheetName}: ${error}`);
      }
    }
  }

  if (DEBUG == "TRUE") {
    Logger.log(
      `${actionLabel} completed: ${count} of ${totalSheets} ` +
      `${groupLabel} sheets ${resultLabel} successfully`
    );
  }
}

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

  if (!sheet_se || !sheet_co){ Logger.log('Settings or Config sheet not found'); return null; }

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
  const Class    = getConfigValue(IST, 'Config');                                       // IST = Is Stock?
  const sheet = fetchSheetByName('Settings');
  if (!sheet) return;

  const Activate = getConfigValue(ACT, 'Settings');                                     // ACT = Activate

  if (Class == 'STOCK')
  {
    if ( Activate == "TRUE")                                              // TRUE
    {
      const True = getConfigValue(TRU, 'Settings');                                     // TRU = True

      if ( True == 'SAVE')                                                // SAVE
      {
        const Save = getConfigValue(SAV, 'Settings');                                    // SAV = SAVE

        if ( Save == 'SHEETS') { doSaveAllBasics(); }
        if ( Save == 'EXTRAS') { doSaveAllExtras(); }
        if ( Save == 'DATAS')  { doSaveAllFinancials(); }
        if ( Save == 'ALL')    { doSaveAll(); }
        if ( Save == 'INDIVIDUAL')
          {
            const Individual = getConfigValue(IND, 'Settings');                          // IND = INDIVIDUAL

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
        const Other = getConfigValue(EXT, 'Settings');                                    // EXT = Extra

        if ( Other == 'ZEROS')           { doCleanZeros(); }
        if ( Other == 'TRIGGERS')        { doCheckTriggers(); }
        if ( Other == 'CHECK')           { doCheckDATAS(); }                              // Check and hide or show Sheets
        if ( Other == 'PROV')            { doSaveProventos(); }
        if ( Other == 'SHARES')          { doSaveShares(); }
        if ( Other == 'RIGHTS')          { doRestoreRight(); }
        if ( Other == 'ZEROS OPTIONS')   { doDeleteZeroOptions(); }
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
    'Index', 'Info', 'Comunicados', 'Prov', 'Preço', 'Cotações',
    'OPT', 'DATA', 'Value', 'Balanco', 'Resultado', 'Fluxo', 'Valor'
  ];

  for (let i = 0; i < SheetNames.length; i++) {
    const Name = SheetNames[i];
    const sheet = fetchSheetByName(Name);
    if (!sheet) continue;

    const range = sheet.getDataRange();
    range.copyTo(range, { contentsOnly: true });
  }
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

  for (let i = 0; i < SheetNames.length; i++) {
    const Name = SheetNames[i];
    const sheet = fetchSheetByName(Name);
    if (!sheet) continue;

    try {
      ss.deleteSheet(sheet);
      Logger.log(`Sheet deleted: ${Name}`);
    } catch (error) {
      Logger.log(`Error deleting sheet "${Name}": ${error}`);
    }
  }
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

  for (let idx = 0; idx < SheetNames.length; idx++) {
    const SheetName = SheetNames[idx];
    const sheet     = fetchSheetByName(SheetName);
    if (!sheet) continue;

    const LR = sheet.getLastRow();
    const LC = sheet.getLastColumn();
    if (LR < 5) continue; // nothing to clean

    const Range = sheet.getRange(5, 1, LR - 4, LC);
    const Data = Range.getValues();
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
      Range.setValues(Data);
      Logger.log(`Zeros cleaned in sheet: ${SheetName}`);
    }
  }
}

function doDeleteZeroOptions() {
  Logger.log(`DELETE: 0 values from call put / blank values from ratios on Sheet ${OPCOES}`);

  const sheet = fetchSheetByName(OPCOES);
  if (!sheet) return;

  const DEBUG   = getConfigValue(DBG, 'Config') === "TRUE";
  const lastRow = sheet.getLastRow();

  for (let row = lastRow; row > 4; row--) {
    try {
      const C = sheet.getRange(row, 3).getValue();
      const E = sheet.getRange(row, 5).getValue();
      const H = sheet.getRange(row, 8).getDisplayValue().trim();
      const I = sheet.getRange(row, 9).getDisplayValue().trim();
      const J = sheet.getRange(row,10).getDisplayValue().trim();

      const zeroCE   = (C === 0 || E === 0);
      const allBlank = (H === "" && I === "" && J === "");

      if (zeroCE || allBlank) {
        // Decide the reason for logging
        let reason;
        if (zeroCE) {
          reason = `zero in C or E (C=${C}, E=${E})`;
        } else {
          reason = `blank H/I/J (H='${H}', I='${I}', J='${J}')`;
        }
        if (DEBUG) { Logger.log(`[doDeleteZeroOptions] Deleting row ${row} due to ${reason}`); }
        sheet.deleteRow(row);
      }
    } catch (err) {
      if (DEBUG) { Logger.log(`[doDeleteZeroOptions] Error on row ${row}: ${err}`); }
    }
  }
  if (DEBUG) { Logger.log(`[doDeleteZeroOptions] Completed scanning rows 5–${lastRow}.`); }
}

function tryCleanOpcaoExportRow(sheet_tr, TKT) {
  Logger.log(`CLEAN: rows with values from call put / blank values from ratios from EXPORTED Source SpreadSheet on Sheet ${sheet_tr}`);

  const colA = sheet_tr.getRange(2, 1, sheet_tr.getLastRow() - 1).getValues();     // only column A, skip header
  const rowIndex = colA.findIndex(row => row[0] === TKT);

  const DEBUG = getConfigValue(DBG, 'Config') === "TRUE";

  if (rowIndex > -1) {
    const rowNum = rowIndex + 2;                                                   // +2 because we started from row 2
    const colCount = sheet_tr.getLastColumn();
    sheet_tr.getRange(rowNum, 1, 1, colCount).clearContent();
    if (DEBUG) Logger.log(`EXPORT CLEAN: OPCOES - Row for ticket ${TKT} cleaned from export sheet.`);
  } else {
    if (DEBUG) Logger.log(`EXPORT CLEAN: OPCOES - Ticket ${TKT} not found on export sheet.`);
  }
}

function normalizeFund() {
  const sheet = fetchSheetByName(FUND);
  if (!sheet) return;

  const MINIMUM = getConfigValue(MIN, 'Settings');
  const MAXIMUM = getConfigValue(MAX, 'Settings');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const rowStart = 3;
  const colStart = 4;  // D
  const colEnd   = 61; // BI

  // Read the Block once
  const Block = sheet.getRange(rowStart, colStart, lastRow - rowStart + 1, colEnd - colStart + 1).getValues();

  // Clamp in-place in the 2D array
  for (let r = 0; r < Block.length; r++) {
    for (let c = 0; c < Block[0].length; c++) {
      const v = Block[r][c];
      if (typeof v === 'number') {
        if (v < MINIMUM)      Block[r][c] = MINIMUM;
        else if (v > MAXIMUM) Block[r][c] = MAXIMUM;
      }
    }
  }

  // Write back the adjusted Block
  sheet.getRange(rowStart, colStart, Block.length, Block[0].length)
       .setValues(Block);

  Logger.log(`NORMALIZE: Clamped FUND cols D–BI, rows ${rowStart}–${lastRow} to [${MINIMUM}, ${MAXIMUM}]`);
}

/////////////////////////////////////////////////////////////////////reverse/////////////////////////////////////////////////////////////////////

function reverseColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const SheetName = sheet.getName();
  const active = fetchSheetByName(SheetName);
  if (!active) return;

  const LR = sheet.getLastRow();
  const LC = sheet.getLastColumn();
  const Range = active.getRange(1, 4, LR, LC - 3);
  const Values = Range.getValues();

  const ReversedValues = Values.map(row => row.reverse());
  Range.setValues(ReversedValues);
  Logger.log(`Columns reversed in sheet: ${SheetName}`);
}

function reverseRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const SheetName = sheet.getName();
  const active = fetchSheetByName(SheetName);
  if (!active) return;

  const LR = sheet.getLastRow();
  const LC = sheet.getLastColumn();
  const Range = active.getRange(5, 1, LR - 4, LC);
  const Values = Range.getValues();

  Values.reverse();
  Range.setValues(Values);
  Logger.log(`Rows reversed in sheet: ${SheetName}`);
}

/////////////////////////////////////////////////////////////////////RESTORE Functions/////////////////////////////////////////////////////////////////////

function doRestoreFundExport() {
  const sheet_co = fetchSheetByName('Config');
  if (!sheet_co) return;

  var Value = '=IF(OR(AND(Fund!A5="";Fund!A1=""); L18<>"STOCK"); FALSE;TRUE)';

    sheet_co.getRange(EFU).setValue(Value);                                 // EFU = Export to Fund
}

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////
