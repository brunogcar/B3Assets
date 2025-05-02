/////////////////////////////////////////////////////////////////////PROCESS SAVE/////////////////////////////////////////////////////////////////////
/**
 * Processes a batch “save” operation over a list of sheet names.
 *
 * This function:
 * 1. Fetches each sheet by name and calls `checkCallback(sheetName)` to decide
 *    whether that sheet has data to save (expects `"TRUE"` for “yes”).
 * 2. Collects all sheets that pass into an array.
 * 3. Flushes the spreadsheet to ensure any pending writes are applied.
 * 4. Iterates over the filtered list, calls `saveCallback(sheetName)` on each,
 *    and logs progress and success/errors conditionally based on DEBUG mode.
 *
 * @param {string[]} SheetNames      An array of sheet‐name constants to consider.
 * @param {function(string):string} checkCallback   Returns "TRUE" or "FALSE" for whether that sheet has data to save.
 * @param {function(string):void} saveCallback      The function to perform the actual save on a sheet name.
 *
 * Behavior:
 * - If a sheet does not exist, logs an error and skips it.
 * - If no sheets have data (`totalSheets === 0`), logs a “skipping” message.
 * - Otherwise, for each sheet:
 *    • Logs `[i/N] (P%) saving <SheetName>...`  
 *    • Calls `saveCallback(SheetName)` inside a try/catch  
 *    • On success or error, logs the outcome **only** if DEBUG is `"TRUE"`.
 */

/**
 * Processes a batch “save” operation over a list of sheet names.
 *
 * @param {string[]} SheetNames      An array of sheet‐name constants to consider.
 * @param {function(string):string} checkCallback   Returns "TRUE" or "FALSE" for whether that sheet has data to save.
 * @param {function(string):void} saveCallback      The function to perform the actual save on a sheet name.
 */
function processSave(SheetNames, checkCallback, saveCallback) {
  const sheet_co = fetchSheetByName('Config');                  // Config sheet
  const DEBUG    = getConfigValue(DBG, 'Config');               // DBG = Debug Mode

  // 1) Gather sheets that exist and pass the check
  const sheetsToSave = [];
  for (let i = 0; i < SheetNames.length; i++) {
    const Name = SheetNames[i];
    const sheet = fetchSheetByName(Name);
    if (!sheet) { Logger.log(`ERROR SAVE: ${Name} - Does not exist`); continue; }
    if (checkCallback(Name) === "TRUE") sheetsToSave.push(Name);
  }

  const totalSheets = sheetsToSave.length;
  if (totalSheets === 0) { Logger.log(`No valid data found. Skipping save operation.`); return; }

  // 2) Flush pending changes
  SpreadsheetApp.flush();

  // 3) Iterate with a simple for-loop
  for (let idx = 0; idx < totalSheets; idx++) {
    const SheetName = sheetsToSave[idx];
    const progress  = Math.round(((idx + 1) / totalSheets) * 100);

    if (DEBUG == "TRUE") Logger.log(`[${idx+1}/${totalSheets}] (${progress}%) saving ${SheetName}...`);

    try {
      saveCallback(SheetName);
      if (DEBUG == "TRUE") Logger.log(`[${idx+1}/${totalSheets}] (${progress}%) ${SheetName} saved successfully`);
    } catch (error) {
      if (DEBUG == "TRUE") Logger.log(`[${idx+1}/${totalSheets}] (${progress}%) Error saving ${SheetName}: ${error}`);
    }
  }
}

/**
 * Converts an array of date strings into timestamps (milliseconds since epoch),
 * handling either "DD/MM/YYYY" or "YYYY-MM-DD" formats. Invalid or unparsable
 * entries return an empty string.
 *
 * @param {string[]} dateStrings
 *   An array of date strings to convert. Supported formats:
 *     • "DD/MM/YYYY" (e.g. "12/5/2024")
 *     • "YYYY-MM-DD" (e.g. "2024-05-12")
 *
 * @returns {(number|string)[]}
 *   An array where each element is:
 *     • A number representing the timestamp (ms since epoch) if parsing succeeded
 *     • An empty string ("") if the input was invalid or not in a supported format
 *
 * @example
 *   const inputs = ["12/5/2024", "2023-10-03", "invalid", null];
 *   const results = doFinancialDateHelper(inputs);
 *   // results might be [1715558400000, 1696281600000, "", ""]
 */
function doFinancialDateHelper(dateStrings) {
  return dateStrings.map(v => {
    // Reject nullish or non-string/non-number values
    if (v == null || typeof v.toString !== "function") {
      return "";
    }

    const str = v.toString().trim();

    // Case 1: "DD/MM/YYYY"
    if (str.includes("/")) {
      const [d, m, y] = str.split("/");
      if (d && m && y) {
        return new Date(+y, +m - 1, +d).getTime();
      }
    }
    // Case 2: "YYYY-MM-DD"
    else if (str.includes("-")) {
      const [y, m, d] = str.split("-");
      if (y && m && d) {
        return new Date(+y, +m - 1, +d).getTime();
      }
    }

    // Fallback: return empty string for invalid formats
    return "";
  });
}

/////////////////////////////////////////////////////////////////////CHECK/////////////////////////////////////////////////////////////////////

/**
 * Iterates through all data-related sheets and runs doCheckDATA on each.
 * Errors are caught per sheet to prevent one failure from stopping the loop.
 */
function doCheckDATAS() {
  const SheetNames = [
    SWING_4, SWING_12, SWING_52,
    PROV, OPCOES, BTC, TERMO, FUND,
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3,
    RIGHT_1, RIGHT_2,
    RECEIPT_9, RECEIPT_10,
    WARRANT_11, WARRANT_12, WARRANT_13,
    BLOCK
  ];

  for (let i = 0; i < SheetNames.length; i++) {
    const SheetName = SheetNames[i];
    try {
      doCheckDATA(SheetName);
    } catch (error) {
      Logger.error(`Error checking DATA for sheet ${SheetName}: ${error}`);
    }
  }
}

/////////////////////////////////////////////////////////////////////DO CHECK TEMPLATE/////////////////////////////////////////////////////////////////////

function doCheckDATA(SheetName) {
  const sheet_sr = fetchSheetByName(SheetName);    // Source sheet
  const sheet_i = fetchSheetByName('Index');       // Index sheet
  const sheet_d = fetchSheetByName('DATA');        // DATA sheet
  const sheet_p = fetchSheetByName(PROV);          // PROV sheet
  const sheet_o = fetchSheetByName('OPT');         // OPT sheet
  const sheet_b = fetchSheetByName(Balanco);       // Balanco sheet
  const sheet_r = fetchSheetByName(Resultado);     // Resultado sheet
  const sheet_f = fetchSheetByName(Fluxo);         // Fluxo sheet
  const sheet_v = fetchSheetByName(Valor);         // Valor sheet

  let Check;

  Logger.log(`CHECK Sheet: ${SheetName}`);

  switch (SheetName) {
//-------------------------------------------------------------------PROV-------------------------------------------------------------------//
    case PROV:
      Check = sheet_p.getRange("B3").getValue();
      break;

//-------------------------------------------------------------------OPCOES-------------------------------------------------------------------//
    case OPCOES:
      Check = sheet_o.getRange("B2").getValue();
      if (Check === '') {
        sheet_o.hideSheet();
        Logger.log(`HIDDEN:`, `OPT`);
      } else if (sheet_o.isSheetHidden()) {
        sheet_o.showSheet();
        Logger.log(`DISPLAYED:  ${SheetName}`);
      }
      break;
//-------------------------------------------------------------------SWING-------------------------------------------------------------------//
    case SWING_4:
    case SWING_12:
    case SWING_52:
      const sheet_co = fetchSheetByName('Config');
      const Class = getConfigValue(IST, 'Config');                                     // IST = Is Stock?
      Check = Class === 'STOCK' ? sheet_d.getRange('B16').getValue() : 'TRUE';
      break;
//-------------------------------------------------------------------BTC-------------------------------------------------------------------//
    case BTC:
      Check = sheet_d.getRange("B3").getValue();
      break;
//-------------------------------------------------------------------TERMO-------------------------------------------------------------------//
    case TERMO:
      Check = sheet_d.getRange("B24").getValue();
      break;
//-------------------------------------------------------------------FUND-------------------------------------------------------------------//
    case FUND:
      Check = sheet_i.getRange("D2").getValue();
      break;
//-------------------------------------------------------------------FUTURE-------------------------------------------------------------------//
    case FUTURE:
      const futureChecks = ["B32", "B33", "B34"];
      for (let i = 0; i < futureChecks.length; i++) {
        Check = sheet_d.getRange(futureChecks[i]).getValue();
        if (!ErrorValues.includes(Check)) break;
      }
      break;

    case FUTURE_1:
      Check = sheet_d.getRange("B32").getValue();
      break;

    case FUTURE_2:
      Check = sheet_d.getRange("B33").getValue();
      break;

    case FUTURE_3:
      Check = sheet_d.getRange("B34").getValue();
      break;
//-------------------------------------------------------------------RIGHT-------------------------------------------------------------------//
    case RIGHT_1:
      Check = sheet_d.getRange("C38").getValue();
      break;

    case RIGHT_2:
      Check = sheet_d.getRange("C39").getValue();
      break;
//-------------------------------------------------------------------RECEIPT-------------------------------------------------------------------//
    case RECEIPT_9:
      Check = sheet_d.getRange("C44").getValue();
      break;

    case RECEIPT_10:
      Check = sheet_d.getRange("C45").getValue();
      break;
//-------------------------------------------------------------------WARRANT-------------------------------------------------------------------//
    case WARRANT_11:
      Check = sheet_d.getRange("C50").getValue();
      break;

    case WARRANT_12:
      Check = sheet_d.getRange("C51").getValue();
      break;

    case WARRANT_13:
      Check = sheet_d.getRange("C52").getValue();
      break;
//-------------------------------------------------------------------BLOCK-------------------------------------------------------------------//
    case BLOCK:
      const blockChecks = ["C56", "C57", "C58"];
      for (let i = 0; i < blockChecks.length; i++) {
        Check = sheet_d.getRange(blockChecks[i]).getValue();
        if (!ErrorValues.includes(Check)) break;
      }
      break;
//-------------------------------------------------------------------BLC / DRE / FLC / DVA-------------------------------------------------------------------//
    case BLC:
      Check = sheet_b.getRange("B1").getValue();
      break;

    case DRE:
      Check = sheet_r.getRange("C1").getValue();
      break;

    case FLC:
      Check = sheet_f.getRange("C1").getValue();
      break;

    case DVA:
      Check = sheet_v.getRange("C1").getValue();
      break;
//-------------------------------------------------------------------DEFAULT-------------------------------------------------------------------//
    default:
      Check = 'FALSE';
      Logger.log(`Sheet Name ${SheetName} not recognized.`);
      break;
  }

  return processCheckDATA(sheet_sr, SheetName, Check);
}

/////////////////////////////////////////////////////////////////////DO CHECK Process/////////////////////////////////////////////////////////////////////

function processCheckDATA(sheet_sr, SheetName, Check) {
  const fixedSheets = [BLC, DRE, FLC, DVA];

  if (ErrorValues.includes(Check)) {
    if (fixedSheets.includes(SheetName)) {
      Logger.log(`DATA Check: FALSE for ${SheetName}`);
      return "FALSE";
    }
    if (!sheet_sr.isSheetHidden()) {
      sheet_sr.hideSheet();
      Logger.log(`Sheet ${SheetName} HIDDEN`);
    }
    Logger.log(`DATA Check: FALSE for ${SheetName}`);
    return "FALSE";
  }

  if (sheet_sr.isSheetHidden()) {
    sheet_sr.showSheet();
    Logger.log(`Sheet ${SheetName} DISPLAYED`);
  }

  Logger.log(`DATA Check: TRUE for ${SheetName}`);
  return "TRUE";
}

/////////////////////////////////////////////////////////////////////TRIM TEMPLATE/////////////////////////////////////////////////////////////////////

function doTrim() {
  const SheetNames = [SWING_4, SWING_12, SWING_52];

  for (let i = 0; i < SheetNames.length; i++) {
    const SheetName = SheetNames[i];
    try {
      doTrimSheet(SheetName);
    } catch (error) {
      Logger.error(`Error trimming sheet ${SheetName}: ${error}`);
    }
  }
}

function doTrimSheet(SheetName) {
  const sheet_sr = fetchSheetByName(SheetName);
  Logger.log(`TRIM: ${SheetName}`);
  if (!sheet_sr) { Logger.error(`Sheet ${SheetName} not found.`); return; }

  const LR = sheet_sr.getLastRow();
  const LC = sheet_sr.getLastColumn();

  switch (SheetName) {
    case SWING_4:
      if (LR > 126) {
        sheet_sr.getRange(127, 1, LR - 126, LC).clearContent();
        Logger.log(`SUCCESS TRIM. Cleared rows 127→${LR} in ${SheetName}.`);
      }
      break;

    case SWING_12:
      if (LR > 366) {
        sheet_sr.getRange(367, 1, LR - 366, LC).clearContent();
        Logger.log(`SUCCESS TRIM. Cleared rows 367→${LR} in ${SheetName}.`);
      }
      break;

    case SWING_52:
      Logger.log(`NOTHING TO TRIM. ${SheetName} stays at ${LR} rows.`);
      break;

    default:
      Logger.log(`No trim logic defined for ${SheetName}.`);
  }
}

/////////////////////////////////////////////////////////////////////Hide and Show Sheets/////////////////////////////////////////////////////////////////////
/**
 * Hides or deletes specific sheets based on the asset class defined in the Config sheet.
 *
 * This function needs direct access to the `SpreadsheetApp.getActiveSpreadsheet()` object 
 * because it performs actions that require the spreadsheet context itself — such as:
 * - Fetching *all* sheets using `getSheets()`
 * - Deleting sheets using `ss.deleteSheet(sheet)`
 *
 * These operations go beyond simply fetching a sheet by name (which `fetchSheetByName()` handles),
 * so we must declare `const ss = SpreadsheetApp.getActiveSpreadsheet();` here directly.
 *
 * @returns {void}
 */

/**
 * Hides or deletes sheets based on the asset class in Config!IST.
 *
 * - STOCK: hides specific sheets listed in `stockHidden`.  
 * - ADR: deletes all sheets *not* in the `adrKeep` set.  
 * - BDR/ETF: deletes all sheets *not* in the `bdrKeep` set.  
 * 
 * Always keeps Config and Settings visible via hideConfig() at end.
 */
function doDisableSheets() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();          // cant remove 
  const sheet_co = fetchSheetByName('Config');
  if (!sheet_co) return;

  const Class = getConfigValue(IST, 'Config');                     // IST = asset class

  const sheets = ss.getSheets();

  switch (Class) {
    case 'STOCK': {
      const stockHidden = [
        'DATA','Prov_','FIBO','Cotações','UPDATE','Balanço',
        'Balanço Ativo','Balanço Passivo','Resultado','Demonstração',
        'Fluxo','Fluxo de Caixa','Valor','Demonstração do Valor Adicionado'
      ];
      for (let i = 0; i < sheets.length; i++) {
        const sh = sheets[i];
        const name = sh.getName();
        if (!sh.isSheetHidden() && stockHidden.indexOf(name) !== -1) {
          sh.hideSheet();
          Logger.log(`Sheet hidden: ${name}`);
        }
      }
      break;
    }
    case 'ADR': {
      const adrKeep = new Set([
        'Config','Settings','Index','Preço','FIBO',
        SWING_4, SWING_12, SWING_52,'Cotações'
      ]);
      // reverse order to safely delete
      for (let i = sheets.length - 1; i >= 0; i--) {
        const sh = sheets[i];
        const name = sh.getName();
        if (!adrKeep.has(name)) {
          ss.deleteSheet(sh);
          Logger.log(`Sheet deleted: ${name}`);
        }
      }
      break;
    }
    case 'BDR':
    case 'ETF': {
      const bdrKeep = new Set([
        'Config','Settings','Index','Prov','Prov_','Preço','FIBO',
        SWING_4, SWING_12, SWING_52,'Cotações','DATA','OPT','Opções','BTC','Termo'
      ]);
      for (let i = sheets.length - 1; i >= 0; i--) {
        const sh = sheets[i];
        const name = sh.getName();
        if (!bdrKeep.has(name)) {
          ss.deleteSheet(sh);
          Logger.log(`Sheet deleted: ${name}`);
        }
      }
      break;
    }
    default:
      Logger.log(`Class "${Class}" not recognized. No sheets modified.`);
  }

  // Always run hideConfig() to re‐hide Config/Settings if needed
  hideConfig();
}

/////////////////////////////////////////////////////////////////////HIDE CONFIG/////////////////////////////////////////////////////////////////////

function hideConfig() {
  const sheet_sr = fetchSheetByName(`Settings`);                        // Source sheet
  const sheet_co = fetchSheetByName(`Config`);                          // Config sheet

  var Hide_Config = sheet_co.getRange(HCR).getDisplayValue();            // HCR = Hide Config Range

  if (Hide_Config == "TRUE") {
    if (sheet_sr && !sheet_sr.isSheetHidden()) {
      sheet_sr.hideSheet();
      Logger.log(`HIDDEN: ${sheet_sr.getName()}`);
    }
    if (sheet_co && !sheet_co.isSheetHidden()) {
      sheet_co.hideSheet();
      Logger.log(`HIDDEN: ${sheet_co.getName()}`);
    }
  }
};

/////////////////////////////////////////////////////////////////////SAVE FUNCTIONS/////////////////////////////////////////////////////////////////////