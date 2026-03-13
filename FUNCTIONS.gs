//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////Helper functions/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////DEBUG/////////////////////////////////////////////////////////////////////

/**
 * Conditional debug logger based on Config‑tab cell DBG.
 * DBG cell must contain one of: 'MIN', 'MID', or 'MAX'.
 *
 * @param {string} msg     The message to log.
 * @param {'MIN'|'MID'|'MAX'} level  How verbose this message is.
 */
let _DBG_CACHE = null;

function LogDebug(msg, level = 'MIN') {
  const ORDER = ['MIN', 'MID', 'MAX'];
  // lazy fetch / cache dbgLevel; refresh if cache is null
  if (_DBG_CACHE === null) {
    const dbgVal = getConfigValue(DBG, 'Config');                                                  // DBG = "L12"

    _DBG_CACHE = (dbgVal && typeof dbgVal === 'string' && dbgVal.trim()) ? dbgVal.trim() : 'MIN';
  }
  const dbgLevel = _DBG_CACHE;
  if (!ORDER.includes(dbgLevel)) _DBG_CACHE = 'MIN';                                               // fallback if invalid

  try {
    if (ORDER.indexOf(dbgLevel) >= ORDER.indexOf(level)) {
      // Prefer Logger.log so Apps Script IDE shows messages
      Logger.log(msg);
    }
  } catch {
    // last resort
    console.log(msg);
  }
}

/////////////////////////////////////////////////////////////////////Group try/catch/////////////////////////////////////////////////////////////////////

/**
 * Executes a batch operation over a list of sheet names with progress tracking
 * and per-item error isolation.
 *
 * This helper runs a provided function sequentially for each sheet name in the
 * supplied list while providing structured logging, execution progress, and
 * safe failure handling. If an operation fails for a specific sheet, the error
 * is logged and processing continues for the remaining sheets.
 *
 * Typical usage includes batch workflows such as saving, importing, exporting,
 * editing, or cleaning groups of sheets.
 *
 * Execution model:
 * - Sequential processing (one sheet at a time)
 * - Each operation runs inside its own try/catch block
 * - Failures do not stop the batch
 *
 * Logging behavior:
 * - Logs the start of the batch operation
 * - Logs progress for each sheet: `[i/N] (P%)`
 * - Logs success or failure for each sheet
 * - Logs a final execution summary
 *
 * @param {string[]} SheetNames         List of item constants.
 * @param {function(string):void} fn    Operation to perform on each sheet.
 * @param {string} actionLabel          Verb in gerund form ("Editing", "Exporting", "Importing").
 * @param {string} resultLabel          Past tense for summary ("edited", "exported", "imported").
 * @param {string} groupLabel           Descriptor for logging ("basic", "extra", "financial", etc.).
 *
 * @returns {void}
 */
function _doGroup(SheetNames, fn, actionLabel, resultLabel, groupLabel) {

  const totalSheets = SheetNames.length;
  let processed = 0;
  let successCount = 0;

  const actionLower = actionLabel.toLowerCase();

  LogDebug(`Starting ${actionLower} of ${totalSheets} ${groupLabel} sheets...`, 'MAX');

  for (const SheetName of SheetNames) {
    processed++;
    const progress = Math.round((processed / totalSheets) * 100);

    LogDebug(`[⏳ ${processed}/${totalSheets}] (${progress}%) ${actionLabel} ${SheetName}...`, 'MAX');

    try {
      fn(SheetName);
      successCount++;

      LogDebug(
        `[🆗 ${processed}/${totalSheets}] (${progress}%) ${SheetName} ${resultLabel} successfully`, 'MAX');

    } catch (error) {
      LogDebug(`[🛑 ${processed}/${totalSheets}] (${progress}%) Error ${actionLower} ${SheetName}: ${error.message || error}`, 'MAX');
    }
  }
  LogDebug(`💾 ${actionLabel} completed: ${successCount} of ${totalSheets} ${groupLabel} sheets ${resultLabel} successfully`, 'MAX');
}

/////////////////////////////////////////////////////////////////////CACHE/////////////////////////////////////////////////////////////////////

//-------------------------------------------------------------------SPREADSHEET-------------------------------------------------------------------//

/**
 * In-memory cache for opened spreadsheets.
 * Used to avoid repeated calls to SpreadsheetApp.openById() within the same script execution.
 * @type {Object.<string, GoogleAppsScript.Spreadsheet.Spreadsheet>}
 */
const _SPREADSHEET_CACHE = {};

/**
 * Returns a cached spreadsheet object for the given ID.
 * If the spreadsheet hasn't been opened yet in this execution, it opens and caches it.
 *
 * @param {string} id  The ID of the spreadsheet to retrieve.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The spreadsheet object.
 *
 * Behavior:
 * - The cache is purely in‑memory and lasts only for the current script execution.
 * - If the same ID is requested multiple times, the cached object is returned immediately,
 *   saving time and avoiding redundant API calls.
 * - If the ID is invalid or the spreadsheet cannot be opened, the error from
 *   SpreadsheetApp.openById() is propagated.
 */
function getSpreadsheetById(id) {
  if (!id) {
    throw new Error("getSpreadsheetById: missing spreadsheet ID");
  }

  if (!_SPREADSHEET_CACHE[id]) {
    LogDebug(`📂 Opening spreadsheet: ${id}`, "MAX");
    _SPREADSHEET_CACHE[id] = SpreadsheetApp.openById(id);
  } else {
    LogDebug(`📦 Spreadsheet cache hit: ${id}`, "MAX");
  }
  return _SPREADSHEET_CACHE[id];
}

//-------------------------------------------------------------------SS-------------------------------------------------------------------//

/**
 * The active spreadsheet instance, cached for the entire script execution.
 * This avoids repeated calls to SpreadsheetApp.getActiveSpreadsheet().
 * @type {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
const _SS_CACHE = SpreadsheetApp.getActiveSpreadsheet();

/**
 * In‑memory cache for sheets retrieved from the active spreadsheet.
 * Keys are sheet names (strings); values are Sheet objects.
 * @type {Object.<string, GoogleAppsScript.Spreadsheet.Sheet>}
 */
const _SHEET_CACHE = {};

/**
 * Retrieves a sheet by name from the active spreadsheet, with caching.
 *
 * @param {string}  SheetName      The exact name of the sheet to retrieve.
 * @param {boolean} [forceRefresh] If true, bypass the cache and force a fresh lookup.
 *                                  Defaults to false.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} The sheet object if found, otherwise null.
 *
 * Behavior:
 * - If SheetName is falsy (empty, null, undefined), returns null immediately.
 * - If forceRefresh is false and the sheet is already in _SHEET_CACHE, returns the cached sheet.
 * - Otherwise, attempts to get the sheet via _SS_CACHE.getSheetByName(SheetName).
 * - If the sheet is found, it is stored in the cache (overwriting any stale entry) and returned.
 */
function getSheet(SheetName, forceRefresh = false) {

  if (!SheetName) return null;

  if (!forceRefresh && _SHEET_CACHE[SheetName]) {
    LogDebug(`📦 Sheet cache hit: ${SheetName}`, "MAX");
    return _SHEET_CACHE[SheetName];
  }

  const sh = _SS_CACHE.getSheetByName(SheetName);

  if (sh) {
    _SHEET_CACHE[SheetName] = sh;
    LogDebug(`✅ Sheet found: ${SheetName}`, "MAX");
  } else {
    LogDebug(`⚠️ Sheet not found: ${SheetName}`, "MIN");
  }

  return sh;
}

/**
 * Clear sheet cache (call when sheets are added/renamed programmatically).
 */
function clearSheetCache() {
  for (const k in _SHEET_CACHE) delete _SHEET_CACHE[k];
}

/////////////////////////////////////////////////////////////////////CONFIG/////////////////////////////////////////////////////////////////////

/**
 * Retrieves a configuration value from named ranges in the Settings and/or Config sheets.
 *
 * @param {string} Acronym         The named range to look up (e.g., "TAX_RATE", "COMPANY_NAME").
 * @param {string} [Source='Both'] Which sheet(s) to search:
 *                                  - 'Settings' : only the Settings sheet.
 *                                  - 'Config'   : only the Config sheet.
 *                                  - 'Both'     : try Settings first, fall back to Config.
 * @returns {string|null}          The trimmed value if found, otherwise null.
 *
 * Behavior:
 * - If the required sheet(s) cannot be obtained via `getSheet()` (e.g., sheet missing),
 *   a warning is logged at level "MIN" and the function returns null for that source.
 * - For each sheet accessed, `getRange(Acronym).getDisplayValue()` is attempted.
 *   - If the named range does not exist or causes an error, the error is caught and
 *     a warning is logged at level "MIN".
 *   - The returned value is trimmed (`trim()`).
 *   - If the trimmed value is empty, equals "DEFAULT", or is included in the external
 *     `ErrorValues` array, it is treated as not found (value = null).
 * - Source-specific logic:
 *   - `'Settings'`: Returns the value from Settings only (or null).
 *   - `'Config'`  : Returns the value from Config only (or null).
 *   - `'Both'`    : If Settings yields a non-null value, that value is returned immediately.
 *                   Otherwise, Config is consulted and its value (or null) is returned.
 *
 * Dependencies:
 * - `getSheet(sheetName)`         – retrieves a cached sheet object.
 * - `LogDebug(message, level)`    – logging function with "MIN"/"MAX" levels.
 * - `ErrorValues` (global array)  – list of strings that should be treated as errors/missing.
 */
function getConfigValue(Acronym, Source = 'Both') {

  const sheet_se = (Source !== 'Config')   ? getSheet('Settings') : null;
  const sheet_co = (Source !== 'Settings') ? getSheet('Config')   : null;

  if (Source !== 'Config' && !sheet_se) {
    LogDebug('⚠️ Settings sheet not found', 'MIN');
    return null;
  }

  if (Source !== 'Settings' && !sheet_co) {
    LogDebug('⚠️ Config sheet not found', 'MIN');
    return null;
  }

  let Value = null;

  if (sheet_se) {
    try {
      Value = sheet_se.getRange(Acronym).getDisplayValue().trim();

      if (!Value || Value === 'DEFAULT' || ErrorValues.includes(Value))
        Value = null;
      else if (Source === 'Settings')
        return Value;

    } catch (e) {
      LogDebug(`⚠️ const ${Acronym} not found in Settings: getConfigValue`, 'MIN');
    }
  }

  if (!Value && sheet_co) {
    try {
      Value = sheet_co.getRange(Acronym).getDisplayValue().trim();

      if (!Value || ErrorValues.includes(Value))
        Value = null;

    } catch (e) {
      LogDebug(`⚠️ const ${Acronym} not found in Config: getConfigValue`, 'MIN');
    }
  }

  return Value;
}

/////////////////////////////////////////////////////////////////////Settings/////////////////////////////////////////////////////////////////////

function doSettings() {
  const Class    = getConfigValue(IST, 'Config');                                       // IST = Is Stock?
  const sheet = getSheet('Settings');
  if (!sheet) return;

  const Activate = getConfigValue(ACT, 'Settings');                                     // ACT = Activate

  if (Class !== 'STOCK' || Activate !== 'TRUE') return;

  const True = getConfigValue(TRU, 'Settings');

  switch (True) {
    case 'SAVE': {
      const Save = getConfigValue(SAV, 'Settings');                                    // SAV = SAVE

      switch (Save) {
        case 'SHEETS':     doSaveAllBasics();    break;
        case 'EXTRAS':     doSaveAllExtras();    break;
        case 'DATAS':      doSaveAllFinancials(); break;
        case 'ALL':        doSaveAll();           break;
        case 'INDIVIDUAL': {
          const Individual = getConfigValue(IND, 'Settings');                          // IND = INDIVIDUAL

          switch (Individual) {
            case 'SWING':  doSaveSWING();          break;
            case 'OPCOES': doSaveBasic(OPCOES);    break;
            case 'BTC':    doSaveBasic(BTC);       break;
            case 'TERMO':  doSaveBasic(TERMO);     break;
            case 'FUND':   doSaveBasic(FUND);      break;
            case 'FUTURE': doSaveBasic(FUTURE);    break;
          }
          break;
        }
      }
      break;
    }

    case 'EXPORT':
      doExportAll();
      break;

    case 'OTHER': {
      const Other = getConfigValue(EXT, 'Settings');                                    // EXT = Extra

      switch (Other) {
        case 'ZEROS':        doCleanZeros();          break;
        case 'TRIGGERS':     doCheckTriggers();       break;
        case 'CHECK':        doCheckDATAS();          break;                            // Check and hide or show Sheets
        case 'PROV':         doSaveProventos();       break;
        case 'SHARES':       doSaveShares();          break;
        case 'ZEROS OPTIONS':doDeleteZeroOptions();   break;
        case 'NORM FUND':    normalizeFund();         break;
      }
      break;
    }
  }
}

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

    const sheet = getSheet(Name);
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

    const sheet = getSheet(Name);
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
  copypasteSheets();
  doDeleteSheets();

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

/**
 * Replaces numeric zero values with blank cells in all basic data sheets.
 *
 * The function iterates through the sheet names defined in `SheetsBasic`
 * and scans their data region starting from row 5 (rows 1–4 are assumed
 * to contain headers or metadata). Any cell containing the numeric value
 * `0` is replaced with an empty string (`""`).
 *
 * Processing strategy:
 * - The entire data block is read into memory with `getValues()`.
 * - Zeros are replaced in-place within the 2D array.
 * - The sheet is only updated with `setValues()` if at least one change
 *   was made, minimizing unnecessary write operations.
 *
 * Range affected:
 * - Rows: 5 → last row
 * - Columns: 1 → last column
 *
 * Logging:
 * - Logs a message for each sheet where zero values were cleaned.
 *
 * Dependencies:
 * - `SheetsBasic` list defining which sheets to process
 * - `getSheet()` helper for sheet retrieval
 * - `LogDebug()` for controlled logging
 *
 * @function doCleanZeros
 * @returns {void}
 */
function doCleanZeros() {
  const SheetNames = SheetsBasic;

  for (const SheetName of SheetNames) {
    const sheet = getSheet(SheetName);
    if (!sheet) continue;

    const LR = sheet.getLastRow();
    const LC = sheet.getLastColumn();
    if (LR < 5) continue; // nothing to clean

    const Range = sheet.getRange(5, 1, LR - 4, LC);
    const Data = Range.getValues();
    let Modified = false;

    for (const row of Data) {
      for (let j = 0; j < row.length; j++) {
        if (row[j] === 0) {
          row[j] = "";
          Modified = true;
        }
      }
    }

    if (Modified) {
      Range.setValues(Data);
      LogDebug(`Zeros cleaned in sheet: ${SheetName}`, 'MIN');
    }
  }
}


/**
 * Removes invalid or incomplete option rows from the OPCOES sheet.
 *
 * The function scans rows starting at row 5 and deletes rows where:
 *
 * 1. Column C or Column E contains the numeric value 0
 *    (typically indicating invalid call/put values).
 *
 * 2. Columns H, I, and J are all blank
 *    (usually indicating missing ratio or pricing data).
 *
 * To improve performance, the function reads the data block A:J into
 * memory once, determines which rows should be removed, and then deletes
 * them from bottom to top to avoid row index shifting.
 *
 * Range scanned:
 * - Rows: 5 → last row
 * - Columns: A → J
 *
 * Logging:
 * - Logs the reason for each row deletion.
 * - Reports when no rows require deletion.
 * - Reports the final number of rows removed.
 *
 * @function doDeleteZeroOptions
 * @returns {void}
 */
function doDeleteZeroOptions() {
  LogDebug(`DELETE: 0 values from call put / blank values from ratios: ${OPCOES}`, 'MIN');

  const sheet = getSheet(OPCOES);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const startRow = 5;

  if (lastRow < startRow) {
    LogDebug(`[doDeleteZeroOptions] No rows to scan in ${OPCOES}`, 'MIN');
    return;
  }

  // Read block once (A:J)
  const block = sheet.getRange(startRow, 1, lastRow - (startRow - 1), 10).getValues();

  const rowsToDelete = [];

  for (let r = 0; r < block.length; r++) {

    const C = block[r][2];
    const E = block[r][4];

    const H = (block[r][7] || '').toString().trim();
    const I = (block[r][8] || '').toString().trim();
    const J = (block[r][9] || '').toString().trim();

    const zeroCE   = (C === 0 || E === 0);
    const allBlank = (H === "" && I === "" && J === "");

    if (zeroCE || allBlank) {

      const rowIndex = r + startRow;

      let reason;
      if (zeroCE) {
        reason = `zero in C or E (C=${C}, E=${E})`;
      } else {
        reason = `blank H/I/J (H='${H}', I='${I}', J='${J}')`;
      }

      LogDebug(
        `[doDeleteZeroOptions] Deleting row ${rowIndex} due to ${reason}`,
        'MIN'
      );

      rowsToDelete.push(rowIndex);
    }
  }

  if (!rowsToDelete.length) {
    LogDebug(`[doDeleteZeroOptions] No rows deleted from ${OPCOES}`, 'MIN');
    return;
  }

  // Delete bottom → top to avoid shifting issues
  rowsToDelete.sort((a, b) => b - a);

  for (const row of rowsToDelete) {
    sheet.deleteRow(row);
  }

  LogDebug(
    `[doDeleteZeroOptions] Deleted ${rowsToDelete.length} rows from ${OPCOES}`,
    'MIN'
  );
}

/**
 * Searches an exported options sheet for a specific ticker and clears
 * the entire row if it exists.
 *
 * The function scans column A (starting at row 2 to skip headers) to find
 * the first row matching the provided ticker. If found, the entire row
 * content is cleared. This is typically used to remove rows generated
 * from incomplete or inconsistent option export data (e.g. mismatched
 * call/put values or missing ratios).
 *
 * Behavior:
 * - Only column A is scanned for the ticker.
 * - Row 1 is skipped (assumed header).
 * - When a match is found, all columns in that row are cleared.
 *
 * Logging:
 * - Reports successful row cleanup.
 * - Reports when the ticker cannot be found.
 *
 * @function tryCleanOpcaoExportRow
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet_tr - Target exported options sheet where the cleanup should occur.
 * @param {string} TKT                                  - Ticker symbol to search for in column A.
 * @returns {void}
 */
function tryCleanOpcaoExportRow(sheet_tr, TKT) {
  LogDebug(`CLEAN: rows with values from call put / blank values from ratios from EXPORTED Source SpreadSheet: ${sheet_tr}`, 'MIN');

  const colA = sheet_tr.getRange(2, 1, sheet_tr.getLastRow() - 1).getValues();     // only column A, skip header
  const rowIndex = colA.findIndex(row => row[0] === TKT);

  if (rowIndex > -1) {
    const rowNum = rowIndex + 2;                                                   // +2 because we started from row 2
    const colCount = sheet_tr.getLastColumn();
    sheet_tr.getRange(rowNum, 1, 1, colCount).clearContent();
    LogDebug(`EXPORT CLEAN: OPCOES - Row for ticket ${TKT} cleaned from exported sheet ${sheet_tr}.`, 'MIN');
  } else {
    LogDebug(`EXPORT CLEAN: OPCOES - Ticket ${TKT} not found: ${sheet_tr}.`, 'MIN');
  }
}

/**
 * Normalizes numeric values in the FUND sheet by clamping them to a
 * configured minimum and maximum range.
 *
 * The function reads the data block from columns D–BI (rows 5 → last row)
 * and adjusts every numeric value so it remains within the range defined
 * in the Settings sheet.
 *
 * Behavior:
 * - Values lower than MINIMUM are replaced with MINIMUM.
 * - Values greater than MAXIMUM are replaced with MAXIMUM.
 * - Non-numeric cells remain unchanged.
 *
 * Configuration:
 * - MINIMUM value is retrieved from Settings via the MIN key.
 * - MAXIMUM value is retrieved from Settings via the MAX key.
 *
 * Range affected:
 * - Rows: 5 → last row
 * - Columns: D → BI
 *
 * Processing strategy:
 * - The entire block is read once into memory.
 * - Values are adjusted directly in the 2D array.
 * - The block is written back in a single `setValues()` call to minimize
 *   Spreadsheet API operations.
 *
 * Logging:
 * - Reports the normalization range applied and the processed rows.
 *
 * @function normalizeFund
 * @returns {void}
 */
function normalizeFund() {
  LogDebug(`NORMALIZE: Values: ${FUND}`, 'MIN');

  const sheet = getSheet(FUND);
  if (!sheet) return;

  const MINIMUM = getConfigValue(MIN, 'Settings');
  const MAXIMUM = getConfigValue(MAX, 'Settings');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const rowStart = 5;
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

  LogDebug(`NORMALIZE: Clamped FUND cols D–BI, rows ${rowStart}–${lastRow} to [${MINIMUM}, ${MAXIMUM}]`, 'MIN');
}

/////////////////////////////////////////////////////////////////////fixNumericFormatting Function/////////////////////////////////////////////////////////////////////

/**
 * Scans financial statement sheets and corrects numeric formatting issues
 * caused by import or copy-paste inconsistencies.
 *
 * The function processes rows 5 → last row of several financial sheets and
 * fixes two common data problems:
 *
 * 1. Text numbers containing thousand separators (e.g. "54.334.248")
 *    → converted to real numeric values (54334248).
 *
 * 2. Numeric values incorrectly interpreted as decimals instead of thousands
 *    (e.g. 462.764 instead of 462764). When a value appears suspiciously
 *    small but has a large fractional component, it is multiplied by 1000
 *    to restore the expected magnitude.
 *
 * Corrections are applied in memory and written back to the sheet only when
 * changes are detected, minimizing write operations.
 *
 * Sheets processed:
 * - Balanço Ativo
 * - Balanço Passivo
 * - Demonstração
 * - Fluxo de Caixa
 * - Demonstração do Valor Adicionado
 *
 * Rows 1–4 are intentionally skipped to preserve headers and metadata.
 *
 * Logging:
 * - Reports per-sheet correction counts.
 * - Warns if a target sheet is missing.
 *
 * @function fixNumericFormatting
 * @returns {void}
 */
function fixNumericFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SheetNames = [
    'Balanço Ativo',
    'Balanço Passivo',
    'Demonstração',
    'Fluxo de Caixa',
    'Demonstração do Valor Adicionado'
  ];

  LogDebug('🧹 Starting cleanup: correcting numeric formatting issues', 'MIN');

  for (const name of SheetNames) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      LogDebug(`⚠️ Sheet not found: ${name}`, 'MID');
      continue;
    }

    const LR = sheet.getLastRow();
    const LC = sheet.getLastColumn();
    if (LR < 5 || LC < 1) continue; // skip if no data

    const range = sheet.getRange(5, 1, LR - 4, LC);
    const data = range.getValues();

    let fixedCount = 0;

    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const val = data[r][c];

        // 1️⃣ Fix text numbers with thousands separators (e.g. "54.334.248")
        if (typeof val === 'string' && /^\-?\d{1,3}(\.\d{3})+$/.test(val)) {
          const cleaned = parseFloat(val.replace(/\./g, ''));
          data[r][c] = cleaned;
          fixedCount++;
          continue;
        }

        // 2️⃣ Fix numeric values that look like they were treated as decimals
        if (typeof val === 'number' && Math.abs(val) > 0 && Math.abs(val) < 10000 && val % 1 !== 0) {
          // Check if the fractional part is large enough to indicate a missing thousand separator
          const fraction = Math.abs(val) % 1;
          if (fraction > 0.1 && fraction < 0.999) {  // e.g., 462.764 or -22.501
            const scaled = Math.round(val * 1000);
            if (Math.abs(scaled) > Math.abs(val) * 10) { // sanity check
              data[r][c] = scaled;
              fixedCount++;
            }
          }
        }
      }
    }

    if (fixedCount > 0) {
      range.setValues(data);
      LogDebug(`✅ ${name}: ${fixedCount} values corrected`, 'MIN');
    } else {
      LogDebug(`ℹ️ ${name}: No corrections needed`, 'MID');
    }
  }

  LogDebug('🎯 Cleanup completed for all sheets.', 'MIN');
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
