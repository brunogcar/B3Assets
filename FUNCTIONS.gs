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
    const dbgVal = runSafely(() => getConfigValue(DBG, 'Config'), 'LogDebug:getConfigValue');      // DBG = "L12"
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

/**
 * Generic batch runner with progress logging and error isolation. (save/edit/export/import/etc).
 *
 * Executes an operation for each item in a list while providing:
 * - progress logging
 * - per-item try/catch protection
 * - execution summary
 *
 * @param {string[]} SheetNames         List of itens constants.
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

  LogDebug(`Starting ${actionLabel.toLowerCase()} of ${totalSheets} ${groupLabel} sheets...`, 'MAX');

  for (let i = 0; i < totalSheets; i++) {
    const SheetName = SheetNames[i];
    count++;
    const progress = Math.round((count / totalSheets) * 100);

    LogDebug(`[⏳ ${count}/${totalSheets}] (${progress}%) ${actionLabel} ${SheetName}...`, 'MAX');

    try {
      fn(SheetName);
      LogDebug(`[🆗 ${count}/${totalSheets}] (${progress}%) ${SheetName} ${resultLabel} successfully`, 'MAX');

    } catch (error) {
      LogDebug(`[🛑 ${count}/${totalSheets}] (${progress}%) Error ${actionLabel.toLowerCase()} ${SheetName}: ${error}`, 'MAX');
    }
  }
  LogDebug(
      `💾 ` +
      `${actionLabel} completed: ${count} of ${totalSheets} ` +
      `${groupLabel} sheets ${resultLabel} successfully`
    , 'MAX');
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
  if (!_SPREADSHEET_CACHE[id]) {
    _SPREADSHEET_CACHE[id] = SpreadsheetApp.openById(id);
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

  if (!forceRefresh && _SHEET_CACHE[SheetName])
    return _SHEET_CACHE[SheetName];

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

/////////////////////////////////////////////////////////////////////VALUES/////////////////////////////////////////////////////////////////////

/**
 * Safe getValues wrapper; returns empty array if range invalid.
 * @param {Sheet} sh
 * @param {number} row
 * @param {number} col
 * @param {number} numRows
 * @param {number} numCols
 */
function getValuesSafe(sh, r, c, numRows, numCols) {
  if (!sh) return [];
  if (numRows <= 0 || numCols <= 0) return [];
  try {
    return sh.getRange(r, c, numRows, numCols).getValues();
  } catch (e) {
    LogDebug(`getValuesSafe failed: ${e.message}`, 'MIN');
    return [];
  }
}

/**
 * Safe setValues wrapper with a small guard.
 */
function setValuesSafe(sh, r, c, values) {
  if (!sh) return false;
  if (!Array.isArray(values) || values.length === 0) return false;
  try {
    sh.getRange(r, c, values.length, values[0].length).setValues(values);
    return true;
  } catch (e) {
    LogDebug(`setValuesSafe failed: ${e.message}`, 'MIN');
    return false;
  }
}

/**
 * Utility to run a function with error handling and consistent logging.
 * @param {function():any} fn
 * @param {string} ctx
 */
function runSafely(fn, ctx) {
  try {
    return fn();
  } catch (e) {
    LogDebug(`Error in ${ctx}: ${e && e.message ? e.message : e}`, 'MIN');
    return null;
  }
}

/////////////////////////////////////////////////////////////////////CONFIG/////////////////////////////////////////////////////////////////////

/**
 * In‑memory cache for configuration values retrieved from named ranges.
 * The cache key is a combination of Source and Acronym (e.g., "Both:TAX_RATE").
 * @type {Object.<string, string|null>}
 */
const _CONFIG_VALUE_CACHE = {};

/**
 * Retrieves a configuration value from named ranges in the Settings and/or Config sheets.
 * Results are cached to avoid repeated lookups within the same script execution.
 *
 * @param {string} Acronym         The named range to look up (e.g., "TAX_RATE", "COMPANY_NAME").
 * @param {string} [Source='Both'] Which sheet(s) to search:
 *                                  - 'Settings' : only the Settings sheet.
 *                                  - 'Config'   : only the Config sheet.
 *                                  - 'Both'     : try Settings first, fall back to Config.
 * @returns {string|null}          The trimmed value if found, otherwise null.
 *
 * Behavior:
 * - The cache key is `${Source}:${Acronym}`. If the key exists in `_CONFIG_VALUE_CACHE`,
 *   its value (which may be null) is returned immediately.
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
 *   - `'Both'`    : If Settings yields a non‑null value, that value is returned immediately.
 *                   Otherwise, Config is consulted and its value (or null) is returned.
 * - The final value (even null) is stored in the cache before being returned.
 *
 * Dependencies:
 * - `getSheet(sheetName)`         – retrieves a cached sheet object.
 * - `LogDebug(message, level)`    – logging function with "MIN"/"MAX" levels.
 * - `ErrorValues` (global array)  – list of strings that should be treated as errors/missing.
 */
function getConfigValue(Acronym, Source = 'Both') {

  const key = `${Source}:${Acronym}`;

  if (_CONFIG_VALUE_CACHE[key] !== undefined)
    return _CONFIG_VALUE_CACHE[key];

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
      else if (Source === 'Settings') {
        _CONFIG_VALUE_CACHE[key] = Value;
        return Value;
      }

    } catch (e) {
      LogDebug(`⚠️ const ${Acronym} not found in Settings`, 'MIN');
    }
  }

  if (!Value && sheet_co) {
    try {
      Value = sheet_co.getRange(Acronym).getDisplayValue().trim();

      if (!Value || ErrorValues.includes(Value))
        Value = null;

    } catch (e) {
      LogDebug(`⚠️ const ${Acronym} not found in Config`, 'MIN');
    }
  }

  _CONFIG_VALUE_CACHE[key] = Value;

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
  const sheet_co = getSheet('Config');
  if (!sheet_co) {
    LogDebug(`⚠️ Config sheet not found; cannot set ${Acronym}`, 'MIN');
    return false;
  }

  try {
    // Write the value
    sheet_co.getRange(Acronym).setValue(value);
    LogDebug(`🆗 Wrote value "${value}" to Config!${Acronym}`, 'MID');
    return true;
  } catch (e) {
    LogDebug(`🛑 Failed to write ${Acronym} to Config: ${e.message}`, 'MIN');
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
  LogDebug(`[${SheetName}] ⏳ ${action} DATES: SR New=${New_sr}-(${raw_New_sr}), TR New=${New_tr}-(${raw_New_tr})`, 'MAX');

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

/////////////////////////////////////////////////////////////////////reverse/////////////////////////////////////////////////////////////////////

function reverseColumns() {
  const sheet     = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const SheetName = sheet.getName();
  LogDebug(`reverseColumns: Starting: "${SheetName}"`, 'MIN');

  const active = getSheet(SheetName);
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

  const active = getSheet(SheetName);
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
  const sheet_co = getSheet('Config');
  if (!sheet_co) return;

  var Value = '=IF(OR(AND(Fund!A5="";Fund!A1=""); L18<>"STOCK"); FALSE;TRUE)';

    sheet_co.getRange(EFU).setValue(Value);                                 // EFU = Export to Fund
}

/////////////////////////////////////////////////////////////////////fixNumericFormatting Function/////////////////////////////////////////////////////////////////////

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
