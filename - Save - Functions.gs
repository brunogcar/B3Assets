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
 * @param {string[]} SheetNames
 *   An array of sheet‐name constants (e.g. [`SWING_4`, `SWING_12`, …]) to consider.
 * @param {function(string):string} checkCallback
 *   A function that returns `"TRUE"` or `"FALSE"` when given a sheet name,
 *   indicating whether there is new data to save.
 * @param {function(string):void} saveCallback
 *   The function to perform the actual save on a sheet name.
 *
 * Behavior:
 * - If a sheet does not exist, logs an error and skips it.
 * - If no sheets have data (`totalSheets === 0`), logs a “skipping” message.
 * - Otherwise, for each sheet:
 *    • Logs `[i/N] (P%) saving <SheetName>...`  
 *    • Calls `saveCallback(SheetName)` inside a try/catch  
 *    • On success or error, logs the outcome **only** if DEBUG is `"TRUE"`.
 */
function processSave(SheetNames, checkCallback, saveCallback) {
  var sheetsToSave = [];
  const sheet_co = fetchSheetByName('Config');                                  // Config sheet
  const DEBUG = sheet_co.getRange(DBG).getDisplayValue();                       // DBG = Debug Mode

  // Gather sheets that are available and pass the check.
  SheetNames.forEach(function(SheetName) {
    var sheet = fetchSheetByName(SheetName);
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
      if (DEBUG = TRUE) Logger.log(`[${count}/${totalSheets}] (${progress}%) saving ${SheetName}...`);
      try {
        saveCallback(SheetName);
        if (DEBUG = TRUE) Logger.log(`[${count}/${totalSheets}] (${progress}%) ${SheetName} saved successfully`);
      } catch (error) {
        if (DEBUG = TRUE) Logger.log(`[${count}/${totalSheets}] (${progress}%) Error saving ${SheetName}: ${error}`);
      }
    });
  } else {
    Logger.log(`No valid data found. Skipping save operation.`);
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

function doCheckDATAS() 
{
  const SheetNames = [
    SWING_4, SWING_12, SWING_52,
    PROV, OPCOES, BTC, TERMO, FUND,
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3,
    RIGHT_1, RIGHT_2,
    RECEIPT_9, RECEIPT_10,
    WARRANT_11, WARRANT_12, WARRANT_13,
    BLOCK
  ];

  SheetNames.forEach(SheetName => 
  {
    try 
    {
      doCheckDATA(SheetName);
    } 
    catch (error) 
    {
      // Handle the error here, you can log it or take appropriate actions.
      Logger.error(`Error checking DATA for sheet ${SheetName}: ${error}`);
    }
  });
}

/////////////////////////////////////////////////////////////////////DO CHECK TEMPLATE/////////////////////////////////////////////////////////////////////

function doCheckDATA(SheetName) {
  const sheet_sr = fetchSheetByName(SheetName);     // Source sheet
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
      var Class = sheet_co.getRange(IST).getDisplayValue();                            // IST = Is Stock? 
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
  const SheetNames = [
    SWING_4, SWING_12, SWING_52
  ];

  SheetNames.forEach(SheetName => 
  {
    try { doTrimSheet(SheetName); } 
    catch (error) { Logger.error(`Error saving sheet ${SheetName}: ${error}`); }
  });
}

function doTrimSheet(SheetName) {
  const sheet_sr = fetchSheetByName(SheetName);

  Logger.log(`TRIM: ${SheetName}`);

  if (!sheet_sr) { 
    Logger.error(`Sheet ${SheetName} not found.`); 
    return; 
  }

  var LR = sheet_sr.getLastRow();
  var LC = sheet_sr.getLastColumn();

  switch (SheetName) {
    case SWING_4:
      if (LR > 126) {
        sheet_sr.getRange(127, 1, LR - 126, LC).clearContent();
        Logger.log(`SUCCESS TRIM. Sheet: ${SheetName}.`);
        Logger.log(`Cleared data below row 126 in ${SheetName}.`);
      }
      break;

    case SWING_12:
      if (LR > 366) {
        sheet_sr.getRange(367, 1, LR - 366, LC).clearContent();
        Logger.log(`SUCCESS TRIM. Sheet: ${SheetName}.`);
        Logger.log(`Cleared data below row 366 in ${SheetName}.`);
      }
      break;

    case SWING_52:
      Logger.log(`NOTHING TO TRIM. Sheet: ${SheetName}.`);
      break;

    default:
      Logger.log(`No specific logic defined  to Trim for ${SheetName}.`);
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
function doDisableSheets(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');
  const sheets = ss.getSheets();

  var Class = sheet_co.getRange(IST).getDisplayValue();                                                                 // IST = Is Stock?
  let SheetNames = [];

  switch (Class) {
    case 'STOCK':
      SheetNames = ['DATA', 'Prov_', 'FIBO', 'Cotações', 'UPDATE', 'Balanço', 'Balanço Ativo', 'Balanço Passivo', 'Resultado', 'Demonstração', 'Fluxo', 'Fluxo de Caixa', 'Valor', 'Demonstração do Valor Adicionado'];

      sheets.forEach(sheet => {
        if (!sheet.isSheetHidden() && SheetNames.includes(sheet.getName())) {
          sheet.hideSheet();
          Logger.log(`Sheet: ${sheet.getName()} HIDDEN`);
        }
      });
      break;

    case 'ADR':
      SheetNames = new Set(['Config', 'Settings', 'Index', 'Preço', 'FIBO', SWING_4, SWING_12, SWING_52, 'Cotações']);

      for (let i = sheets.length - 1; i >= 0; i--) {                                                                    // Reverse iteration to avoid index shifting
        const sheet = sheets[i];
        if (!SheetNames.has(sheet.getName())) {                                                                         // Delete all but SheetNames
          Logger.log(`Deleting sheet: ${sheet.getName()}`);
          ss.deleteSheet(sheet);
        }
      }
      break;

    case 'BDR':
    case 'ETF':
      SheetNames = new Set(['Config', 'Settings', 'Index', 'Prov', 'Prov_', 'Preço', 'FIBO', SWING_4, SWING_12, SWING_52, 'Cotações', 'DATA', 'OPT', 'Opções', 'BTC', 'Termo']);

      for (let i = sheets.length - 1; i >= 0; i--) {                                                                    // Reverse iteration to avoid index shifting
        const sheet = sheets[i];
        if (!SheetNames.has(sheet.getName())) {                                                                         // Delete all but SheetNames
          Logger.log(`Deleting sheet: ${sheet.getName()}`);
          ss.deleteSheet(sheet);
        }
      }
      break;
      
    default:
      Logger.log(`Class ${Class} not recognized. No sheets modified.`);
  }
  hideConfig();
}

/////////////////////////////////////////////////////////////////////HIDE CONFIG/////////////////////////////////////////////////////////////////////

function hideConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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