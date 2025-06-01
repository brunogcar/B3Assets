//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////Helper functions/////////////////////////////////////////////////////////////////////

/**
 * Conditional debug logger based on Config‑tab cell DBG.
 * DBG cell must contain one of: "MIN", "MID", or "MAX".
 *
 * @param {string} msg     The message to log.
 * @param {"MIN"|"MID"|"MAX"} level  How verbose this message is.
 */
function LogDebug(msg, level = "MIN") {
  // Define verbosity order
  const ORDER = ["MIN", "MID", "MAX"];

  // Read the current debug setting (cell L12 on your Config sheet)
  const dbgLevel = getConfigValue(DBG, 'Config');  // DBG = "L12"

  // Only log if the message’s level is at or below the configured level
  if (ORDER.indexOf(dbgLevel) >= ORDER.indexOf(level)) {
    Logger.log(msg);
  }
}

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
 */
function _doGroup(SheetNames, fn, actionLabel, resultLabel, groupLabel) {
  const totalSheets = SheetNames.length;
  let count = 0;

  LogDebug(`Starting ${actionLabel.toLowerCase()} of ${totalSheets} ${groupLabel} sheets...`, "MAX");

  for (let i = 0; i < totalSheets; i++) {
    const SheetName = SheetNames[i];
    count++;
    const progress = Math.round((count / totalSheets) * 100);

    LogDebug(`[⏳ ${count}/${totalSheets}] (${progress}%) ${actionLabel} ${SheetName}...`, "MAX");

    try {
      fn(SheetName);
      LogDebug(`[🆗 ${count}/${totalSheets}] (${progress}%) ${SheetName} ${resultLabel} successfully`, "MAX");

    } catch (error) {
      LogDebug(`[🛑 ${count}/${totalSheets}] (${progress}%) Error ${actionLabel.toLowerCase()} ${SheetName}: ${error}`, "MAX");
    }
  }
  LogDebug(
      `💾 ` +
      `${actionLabel} completed: ${count} of ${totalSheets} ` +
      `${groupLabel} sheets ${resultLabel} successfully`
    , "MAX");
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
  if (!sheet) { LogDebug(`⚠️ Sheet not found: ${SheetName}`, "MIN"); return null; }
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
  // Only fetch the sheets you need
  const sheet_se = (Source !== 'Config')   ? fetchSheetByName('Settings') : null;
  const sheet_co = (Source !== 'Settings') ? fetchSheetByName('Config')   : null;

  // If we needed Settings but didn't get it, bail
  if (Source !== 'Config' && !sheet_se) { LogDebug('⚠️ Settings sheet not found', "MIN");
    return null;
  }
  // If we needed Config but didn't get it, bail
  if (Source !== 'Settings' && !sheet_co) { LogDebug('⚠️ Config sheet not found', "MIN");
    return null;
  }

  let Value = null;

  // Try Settings first if applicable
  if (sheet_se) {
    try {
      Value = sheet_se.getRange(Acronym).getDisplayValue().trim();
      if (!Value || Value === 'DEFAULT' || ErrorValues.includes(Value)) {
        Value = null;  // fall back to Config
      } else if (Source === 'Settings') {
        return Value;  // short‑circuit if only pulling from Settings
      }
    } catch (e) {
      LogDebug(`⚠️ const ${Acronym} not found in Settings: getConfigValue`, "MIN");
    }
  }

  // Then Config if we still need a value
  if (!Value && sheet_co) {
    try {
      Value = sheet_co.getRange(Acronym).getDisplayValue().trim();
      if (!Value || ErrorValues.includes(Value)) {
        Value = null;
      }
    } catch (e) {
      LogDebug(`⚠️ const ${Acronym} not found in Config: getConfigValue`, "MIN");
    }
  }
  return Value;
}

/**
 * Writes a single value into the Config sheet at the given A1‑notation.
 *
 * @param {string} Acronym  The A1‑notation of the cell (e.g. EXR).
 * @param {string|number} value  The value to write into that cell.
 * @returns {boolean}  True if the write succeeded, false otherwise.
 */
function setConfigValue(Acronym, value) {
  // Fetch the Config sheet
  const sheet_co = fetchSheetByName('Config');
  if (!sheet_co) {
    LogDebug(`⚠️ Config sheet not found; cannot set ${Acronym}`, "MIN");
    return false;
  }

  try {
    // Write the value
    sheet_co.getRange(Acronym).setValue(value);
    LogDebug(`🆗 Wrote value "${value}" to Config!${Acronym}`, "MID");
    return true;
  } catch (e) {
    LogDebug(`🛑 Failed to write ${Acronym} to Config: ${e.message}`, "MIN");
    return false;
  }
}

/////////////////////////////////////////////////////////////////////Check Dates/////////////////////////////////////////////////////////////////////

/**
 * Reads and validates the “New” and “Old” date values from both target (TR) and source (SR) sheets.
 *
 * @param {Sheet}      sheet_tr   The “target” sheet (ticker sheet).
 * @param {Sheet}      sheet_sr   The “source” sheet (template sheet).
 * @param {Object}     cfg        The financialMap entry for this sheet.
 * @param {string}     SheetName  The sheet’s name (for logging).
 * @param {string}     action     Either "SAVE" or "EDIT" (for clearer logging).
 *
 * @returns {{New_tr: Date, Old_tr: Date, New_sr: Date, Old_sr: Date}|null}
 *   Returns the four parsed Date objects if all are valid.
 *   If any is invalid, logs which one(s) and returns null.
 */
function extractAndValidateDates(sheet_tr, sheet_sr, cfg, SheetName, action) {
  // 1) Read TR dates
  const raw_New_tr = sheet_tr.getRange(1, cfg.col_new).getDisplayValue();
  const raw_Old_tr = sheet_tr.getRange(1, cfg.col_old).getDisplayValue();
  LogDebug(`[${cfg.sh_tr}] Raw Dates (TR): New=${raw_New_tr}, Old=${raw_Old_tr}, col_new=${cfg.col_new}, col_old=${cfg.col_old}`, 'MAX');

  const [New_tr, Old_tr] = doFinancialDateHelper([raw_New_tr, raw_Old_tr]);

  // 2) Read SR dates (conditional old-date column)
  const raw_New_sr = sheet_sr.getRange(1, cfg.col_new).getDisplayValue();
  const oldCol     = cfg.recurse ? cfg.col_old_src : cfg.col_old;
  const raw_Old_sr = sheet_sr.getRange(1, oldCol).getDisplayValue();
  LogDebug(`[${cfg.sh_sr}] Raw Dates (SR): New=${raw_New_sr}, Old=${raw_Old_sr}, col_new=${cfg.col_new}, col_old_src=${oldCol}`, 'MAX');
  const [New_sr, Old_sr] = doFinancialDateHelper([raw_New_sr, raw_Old_sr]);

  // 3) Validate each Date using isValidDate()
  const dateNames  = ['New_tr','Old_tr','New_sr','Old_sr'];
  const dateValues = [New_tr,  Old_tr,  New_sr,  Old_sr];

  const badDates = [];
  for (let i = 0; i < dateValues.length; i++) {
    if (!isValidDate(dateValues[i])) {
      badDates.push(`${dateNames[i]}='${dateValues[i]}'`);
    }
  }
  if (badDates.length) {
    // Example log: “❌ ERROR SAVE: BalanceSheet2019 – Invalid date(s): New_sr='-', Old_tr='foo'”
    LogDebug(
      `❌ ERROR ${action}: ${SheetName} - Invalid date(s): ${badDates.join(', ')}`,
      'MID'
    );
    return null;
  }

  // 4) Everything’s valid—return parsed Dates
  return { New_tr, Old_tr, New_sr, Old_sr };
}

/**
 * @param {Date|string} dateCandidate
 * @returns {boolean} true if `dateCandidate` is a valid Date or parseable string
 */
function isValidDate(dateCandidate) {
  // If it’s already a Date, check .valueOf()
  if (dateCandidate instanceof Date) {
    return !isNaN(dateCandidate.valueOf());
  }
  // If it’s a string, try to convert
  const parsed = new Date(dateCandidate);
  return !isNaN(parsed.valueOf());
}

/////////////////////////////////////////////////////////////////////Compare Columns/////////////////////////////////////////////////////////////////////

/**
 * Compares two single‐column ranges (same number of rows) and returns an array of differences.
 *
 * @param {Sheet}   sheetA    The “source” sheet (where updated values live).
 * @param {Sheet}   sheetB    The “target” sheet (where current values live).
 * @param {number}  colA      Column index (1-based) in sheetA.
 * @param {number}  colB      Column index (1-based) in sheetB.
 * @param {number}  lastRow   Number of rows to compare (starting at row 1).
 *
 * @return {Array<{row: number, value: any}>}
 *   An array of objects, one per row where sheetA ≠ sheetB:
 *   – `row`: the 1-based row index
 *   – `value`: the sheetA value at that row/column
 *
 * Example:
 *   //   If sheetA!A1:A3 = [10, 20, 30]
 *   //   and sheetB!B1:B3 = [10, 25, 30]
 *   //   getColumnDifferences(sheetA, sheetB, 1, 2, 3)
 *   //   → [ {row: 2, value: 20} ]
 */
function getColumnDifferences(sheetA, sheetB, colA, colB, lastRow) {
  // Read both columns in one go each, then flatten to 1-D arrays
  const valuesA = sheetA.getRange(1, colA, lastRow, 1).getValues().flat();
  const valuesB = sheetB.getRange(1, colB, lastRow, 1).getValues().flat();
  const diffs   = [];

  // Compare row by row
  for (let i = 0; i < lastRow; i++) {
    if (valuesA[i] !== valuesB[i]) {
      diffs.push({ row: i + 1, value: valuesA[i] });
    }
  }

  return diffs;
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
        if ( Other == 'NORM FUND')       { normalizeFund(); }
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
  LogDebug('copypasteSheets: Starting formula clear on core sheets', 'MIN');

  const SheetNames = [
    'Index', 'Info', 'Comunicados', 'Prov', 'Preço', 'Cotações',
    'OPT', 'DATA', 'Value', 'Balanco', 'Resultado', 'Fluxo', 'Valor'
  ];

  for (const Name of SheetNames) {
    LogDebug(`copypasteSheets: Processing sheet "${Name}"`, 'MID');

    const sheet = fetchSheetByName(Name);
    if (!sheet) {
      LogDebug(`copypasteSheets: Sheet not found, skipping "${Name}"`, 'MID');
      continue;
    }

    try {
      const range = sheet.getDataRange();
      range.copyTo(range, { contentsOnly: true });
      LogDebug(`copypasteSheets: Cleared formulas: "${Name}"`, 'MIN');
    } catch (e) {
      LogDebug(`copypasteSheets: Error copying: "${Name}": ${e.message}`, 'MIN');
    }
  }
}

function doDeleteSheets() {
  LogDebug('doDeleteSheets: Starting deletion of obsolete sheets', 'MIN');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SheetNames = [
    'Balanço Ativo',
    'Balanço Passivo',
    'Demonstração',
    'Fluxo de Caixa',
    'Demonstração do Valor Adicionado'
  ];

  for (const Name of SheetNames) {
    LogDebug(`doDeleteSheets: Attempting to delete "${Name}"`, 'MID');

    const sheet = fetchSheetByName(Name);
    if (!sheet) {
      LogDebug(`doDeleteSheets: Sheet not found, skipping "${Name}"`, 'MID');
      continue;
    }

    try {
      ss.deleteSheet(sheet);
      LogDebug(`doDeleteSheets: Deleted sheet "${Name}"`, 'MIN');
    } catch (error) {
      LogDebug(`doDeleteSheets: Error deleting "${Name}": ${error.message}`, 'MIN');
    }
  }
}

function moveSpreadsheetToFolder(folderName) {

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const file  = DriveApp.getFileById(ss.getId());
  LogDebug(`moveSpreadsheetToFolder: File ID = ${file.getId()}`, 'MAX');

  const folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) {
    LogDebug(`moveSpreadsheetToFolder: Folder "${folderName}" not found`, 'MID');
    return;
  }

  const folder = folders.next();
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  LogDebug(`moveSpreadsheetToFolder: Moved file to "${folderName}"`, 'MIN');
}

function moveSpreadsheetToARQUIVO() {
  LogDebug('moveSpreadsheetToARQUIVO: Starting', 'MIN');
  moveSpreadsheetToFolder('-=ARQUIVO=-');
  LogDebug('moveSpreadsheetToARQUIVO: Finished', 'MIN');
}


/////////////////////////////////////////////////////////////////////DELETE/////////////////////////////////////////////////////////////////////

function doDelete() {
  doDeleteTriggers();
  moveSpreadsheetToBACKUP();
  revokeOwnAccess();
}

/**
 * Revokes the script’s own authorization token so it will prompt for re‑authorization
 * on the next run.
 *
 * @returns {void}
 */
function revokeOwnAccess() {
  LogDebug('revokeOwnAccess: Starting', 'MIN');

  // Check current authorization info
  const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  LogDebug(`revokeOwnAccess: Status = ${authInfo.getAuthorizationStatus()}`, 'MAX');

  if (authInfo) {
    ScriptApp.invalidateAuth();
    LogDebug('revokeOwnAccess: Script access revoked successfully.', 'MIN');
  } else {
    LogDebug('revokeOwnAccess: Script is not authorized or access already revoked.', 'MIN');
  }
}

function moveSpreadsheetToBACKUP() {
  LogDebug('moveSpreadsheetToBACKUP: Starting', 'MIN');
  moveSpreadsheetToFolder('-=BACKUP=-');
  LogDebug('moveSpreadsheetToBACKUP: Finished', 'MIN');
}

function doDeleteSpreadsheet() {
  LogDebug('doDeleteSpreadsheet: Starting permanent deletion', 'MIN');

  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const fileId = ss.getId();
  LogDebug(`doDeleteSpreadsheet: File ID = ${fileId}`, 'MAX');

  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    LogDebug('doDeleteSpreadsheet: Spreadsheet trashed successfully', 'MIN');
  } catch (error) {
    LogDebug(`doDeleteSpreadsheet: Error deleting spreadsheet: ${error}`, 'MIN');
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
  LogDebug(`DELETE: 0 values from call put / blank values from ratios: ${OPCOES}`, "MIN");

  const sheet = fetchSheetByName(OPCOES);
  if (!sheet) return;

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
        LogDebug(`[doDeleteZeroOptions] Deleting row ${row} due to ${reason}`, "MIN");
        sheet.deleteRow(row);
      }
    } catch (err) {
      LogDebug(`[doDeleteZeroOptions] Error: row ${row}: ${err}`, "MIN");
    }
  }
}

function tryCleanOpcaoExportRow(sheet_tr, TKT) {
  LogDebug(`CLEAN: rows with values from call put / blank values from ratios from EXPORTED Source SpreadSheet: ${sheet_tr}`, "MIN");

  const colA = sheet_tr.getRange(2, 1, sheet_tr.getLastRow() - 1).getValues();     // only column A, skip header
  const rowIndex = colA.findIndex(row => row[0] === TKT);

  if (rowIndex > -1) {
    const rowNum = rowIndex + 2;                                                   // +2 because we started from row 2
    const colCount = sheet_tr.getLastColumn();
    sheet_tr.getRange(rowNum, 1, 1, colCount).clearContent();
    LogDebug(`EXPORT CLEAN: OPCOES - Row for ticket ${TKT} cleaned from exported sheet ${sheet_tr}.`, "MIN");
  } else {
    LogDebug(`EXPORT CLEAN: OPCOES - Ticket ${TKT} not found: ${sheet_tr}.`, "MIN");
  }
}

function normalizeFund() {
  LogDebug(`NORMALIZE: Values: ${FUND}`, "MIN");

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

  LogDebug(`NORMALIZE: Clamped FUND cols D–BI, rows ${rowStart}–${lastRow} to [${MINIMUM}, ${MAXIMUM}]`, "MIN");
}

/////////////////////////////////////////////////////////////////////reverse/////////////////////////////////////////////////////////////////////

function reverseColumns() {
  const sheet     = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const SheetName = sheet.getName();
  LogDebug(`reverseColumns: Starting: "${SheetName}"`, 'MIN');

  const active = fetchSheetByName(SheetName);
  if (!active) {
    LogDebug(`reverseColumns: "${SheetName}" not found`, 'MIN');
    return;
  }

  const LR = active.getLastRow();
  const LC = active.getLastColumn();
  const Range = active.getRange(1, 4, LR, LC - 3);  // cols D→last
  LogDebug(`reverseColumns: Range = ${Range.getA1Notation()}`, 'MAX');

  const Values = Range.getValues();
  LogDebug(`reverseColumns: Original values snapshot: ${JSON.stringify(Values)}`, 'MAX');

  const reversed = Values.map(row => row.reverse());
  Range.setValues(reversed);
  LogDebug(`reverseColumns: Columns reversed for ${Values.length} rows`, 'MIN');
}

function reverseRows() {
  const sheet     = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const SheetName = sheet.getName();
  LogDebug(`reverseRows: Starting: "${SheetName}"`, 'MIN');

  const active = fetchSheetByName(SheetName);
  if (!active) {
    LogDebug(`reverseRows: "${SheetName}" not found`, 'MIN');
    return;
  }

  const LR = active.getLastRow();
  const LC = active.getLastColumn();
  const Range = active.getRange(5, 1, LR - 4, LC);  // rows 5→last
  LogDebug(`reverseRows: Range = ${Range.getA1Notation()}`, 'MAX');

  const Values = Range.getValues();
  LogDebug(`reverseRows: Original values snapshot: ${JSON.stringify(Values)}`, 'MAX');

  const reversed = Values.reverse();
  Range.setValues(reversed);
  LogDebug(`reverseRows: Rows reversed (count = ${Values.length})`, 'MIN');
}

/////////////////////////////////////////////////////////////////////RESTORE Functions/////////////////////////////////////////////////////////////////////

function doRestoreFundExport() {
  const sheet_co = fetchSheetByName('Config');
  if (!sheet_co) return;

  var Value = '=IF(OR(AND(Fund!A5="";Fund!A1=""); L18<>"STOCK"); FALSE;TRUE)';

    sheet_co.getRange(EFU).setValue(Value);                                 // EFU = Export to Fund
}

/////////////////////////////////////////////////////////////////////Unicode emoji or symbol/////////////////////////////////////////////////////////////////////

/*                                                                  to be added to log messages
Meaning	Emoji/Symbol	Codepoint
Success / OK	✅ ✔️ 🆗	U+2705 U+2714 U+1F197
Failure / Error	❌ ✖️ 🛑	U+274C U+2716 U+1F6D1
Warning	⚠️ 🔶 🔸	U+26A0 U+1F536 U+1F538
Info / Notice	ℹ️ 🛈 📘	U+2139 U+1F6C8 U+1F4D8
Debug / Trace	🐛 🔍 🛠️	U+1F41B U+1F50D U+1F6E0
In Progress	🔄 ⏳ ⏱️	U+1F504 U+23F3 U+23F1
Data / I/O	📈 📉 💾	U+1F4C8 U+1F4C9 U+1F4BE
Locks / Sync	🔒 🔓 🔐	U+1F512 U+1F513 U+1F510
Flags / Pins	📌 🚩 🏷️	U+1F4CC U+1F6A9 U+1F3F7
Notifications	🔔 🔕 🔕	U+1F514 U+1F515
*/

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////
