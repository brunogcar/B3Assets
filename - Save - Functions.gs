/////////////////////////////////////////////////////////////////////SAVE FUNCTIONS/////////////////////////////////////////////////////////////////////
/**
 * Processes a batch “save” operation over a list of sheet names.
 *
 * This function:
 * 1. Fetches each sheet by name and calls `checkCallback(sheetName)` to decide
 *    whether that sheet has data to save (expects `"TRUE"` for “yes”).
 * 2. Collects all sheets that pass into an array.
 * 3. Flushes the spreadsheet to ensure any pending writes are applied.
 * 4. Iterates over the filtered list, calls `saveFunction(sheetName)` on each,
 *    and logs progress and success/errors conditionally based on DEBUG mode.
 *
 * @param {string[]} SheetNames      An array of sheet‐name constants to consider.
 * @param {function(string):string}  checkCallback   Returns "TRUE" or "FALSE" for whether that sheet has data to save.
 * @param {function(string):void}    saveFunction      The function to perform the actual save on a sheet name.
 */
function doSaveGroup(SheetNames, checkCallback, saveFunction) {
  const SheetNamesToSave = [];
  for (let i = 0; i < SheetNames.length; i++) {
    const Name = SheetNames[i];
    const sheet = fetchSheetByName(Name);
    if (!sheet) return;

    if (checkCallback(Name) === "TRUE") {
      SheetNamesToSave.push(Name);
    }
  }

  const totalSheets = SheetNamesToSave.length;
  if (totalSheets === 0) {
    LogDebug(`No valid data found. Skipping save operation.`, 'MIN');
    return;
  }

  SpreadsheetApp.flush();

  _doGroup(SheetNamesToSave, saveFunction, "Saving", "saved", "");
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
      LogDebug(`Error checking DATA for sheet ${SheetName}: ${error}`, 'MIN');
    }
  }
}

/////////////////////////////////////////////////////////////////////DO CHECK TEMPLATE/////////////////////////////////////////////////////////////////////

function doCheckDATA(SheetName) {
  const sheet_sr = fetchSheetByName(SheetName);    // Source sheet
  const sheet_i  = fetchSheetByName('Index');      // Index sheet
  const sheet_d  = fetchSheetByName('DATA');       // DATA sheet
  const sheet_p  = fetchSheetByName(PROV);         // PROV sheet
  const sheet_o  = fetchSheetByName('OPT');        // OPT sheet
  const sheet_b  = fetchSheetByName(Balanco);      // Balanco sheet
  const sheet_r  = fetchSheetByName(Resultado);    // Resultado sheet
  const sheet_f  = fetchSheetByName(Fluxo);        // Fluxo sheet
  const sheet_v  = fetchSheetByName(Valor);        // Valor sheet

  LogDebug(`CHECK Sheet: ${SheetName}`, 'MIN');

  const cfg = {
    [PROV]:       { sheetVar: sheet_p, cell:  "B3",        toggleHide: false, classSheet: false, cells: null },
    [OPCOES]:     { sheetVar: sheet_o, cell:  "B2",        toggleHide: true,  classSheet: false, cells: null },
    [SWING_4]:    { sheetVar: sheet_d, cell:  "B16",       toggleHide: false, classSheet: true,  cells: null },
    [SWING_12]:   { sheetVar: sheet_d, cell:  "B16",       toggleHide: false, classSheet: true,  cells: null },
    [SWING_52]:   { sheetVar: sheet_d, cell:  "B16",       toggleHide: false, classSheet: true,  cells: null },
    [BTC]:        { sheetVar: sheet_d, cell:  "B3",        toggleHide: false, classSheet: false, cells: null },
    [TERMO]:      { sheetVar: sheet_d, cell:  "B24",       toggleHide: false, classSheet: false, cells: null },
    [FUND]:       { sheetVar: sheet_i, cell:  "D2",        toggleHide: false, classSheet: false, cells: null },
    [FUTURE]:     { sheetVar: sheet_d, cell:  null,        toggleHide: false, classSheet: false, cells: ["B32","B33","B34"] },
    [FUTURE_1]:   { sheetVar: sheet_d, cell:  "B32",       toggleHide: false, classSheet: false, cells: null },
    [FUTURE_2]:   { sheetVar: sheet_d, cell:  "B33",       toggleHide: false, classSheet: false, cells: null },
    [FUTURE_3]:   { sheetVar: sheet_d, cell:  "B34",       toggleHide: false, classSheet: false, cells: null },
    [RIGHT_1]:    { sheetVar: sheet_d, cell:  "C38",       toggleHide: false, classSheet: false, cells: null },
    [RIGHT_2]:    { sheetVar: sheet_d, cell:  "C39",       toggleHide: false, classSheet: false, cells: null },
    [RECEIPT_9]:  { sheetVar: sheet_d, cell:  "C44",       toggleHide: false, classSheet: false, cells: null },
    [RECEIPT_10]: { sheetVar: sheet_d, cell:  "C45",       toggleHide: false, classSheet: false, cells: null },
    [WARRANT_11]: { sheetVar: sheet_d, cell:  "C50",       toggleHide: false, classSheet: false, cells: null },
    [WARRANT_12]: { sheetVar: sheet_d, cell:  "C51",       toggleHide: false, classSheet: false, cells: null },
    [WARRANT_13]: { sheetVar: sheet_d, cell:  "C52",       toggleHide: false, classSheet: false, cells: null },
    [BLOCK]:      { sheetVar: sheet_d, cell:  null,        toggleHide: false, classSheet: false, cells: ["C56","C57","C58"] },
    [BLC]:        { sheetVar: sheet_b, cell:  "B1",        toggleHide: false, classSheet: false, cells: null },
    [DRE]:        { sheetVar: sheet_r, cell:  "C1",        toggleHide: false, classSheet: false, cells: null },
    [FLC]:        { sheetVar: sheet_f, cell:  "C1",        toggleHide: false, classSheet: false, cells: null },
    [DVA]:        { sheetVar: sheet_v, cell:  "C1",        toggleHide: false, classSheet: false, cells: null },
  }[SheetName];

  if (!cfg) {
    LogDebug(`Sheet Name "${SheetName}" not recognized.`, 'MIN');
    return processCheckDATA(sheet_sr, SheetName, 'FALSE');
  }

  let Check = 'FALSE';

  // 1) toggleHide case (OPCOES)
  if (cfg.toggleHide) {
    Check = cfg.sheetVar.getRange(cfg.cell).getValue();
    if (Check === '') {
      sheet_o.hideSheet();
      LogDebug(`HIDDEN: OPT`, 'MIN');
    } else if (sheet_o.isSheetHidden()) {
      sheet_o.showSheet();
      LogDebug(`DISPLAYED: ${SheetName}`, 'MIN');
    }

  // 2) classSheet case (SWING_x)
  } else if (cfg.classSheet) {
    const Class   = getConfigValue(IST, 'Config');
    Check = (Class === 'STOCK')
      ? sheet_d.getRange(cfg.cell).getValue()
      : 'TRUE';

  // 3) cells array case (FUTURE, BLOCK)
  } else if (cfg.cells) {
    for (let addr of cfg.cells) {
      const val = sheet_d.getRange(addr).getValue();
      if (!ErrorValues.includes(val)) {
        Check = val;
        break;
      }
    }

  // 4) simple cell case
  } else {
    Check = cfg.sheetVar.getRange(cfg.cell).getValue();
  }

  return processCheckDATA(sheet_sr, SheetName, Check);
}

/////////////////////////////////////////////////////////////////////DO CHECK Process/////////////////////////////////////////////////////////////////////

function processCheckDATA(sheet_sr, SheetName, Check) {
  const fixedSheets = [BLC, DRE, FLC, DVA];

  if (ErrorValues.includes(Check)) {
    if (fixedSheets.includes(SheetName)) {
      LogDebug(`DATA Check: FALSE for ${SheetName}`, 'MIN');
      return "FALSE";
    }
    if (!sheet_sr.isSheetHidden()) {
      sheet_sr.hideSheet();
      LogDebug(`Sheet ${SheetName} HIDDEN`, 'MIN');
    }
    LogDebug(`DATA Check: FALSE for ${SheetName}`, 'MIN');
    return "FALSE";
  }

  if (sheet_sr.isSheetHidden()) {
    sheet_sr.showSheet();
    LogDebug(`Sheet ${SheetName} DISPLAYED`, 'MIN');
  }

  LogDebug(`DATA Check: TRUE for ${SheetName}`, 'MIN');
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
      LogDebug(`Error trimming sheet ${SheetName}: ${error}`, 'MIN');
    }
  }
}

function doTrimSheet(SheetName) {
  LogDebug(`TRIM: ${SheetName}`, 'MIN');

  const sheet_sr = fetchSheetByName(SheetName);
  if (!sheet_sr) return;

  const LR = sheet_sr.getLastRow();
  const LC = sheet_sr.getLastColumn();

  switch (SheetName) {
    case SWING_4:
      if (LR > 126) {
        sheet_sr.getRange(127, 1, LR - 126, LC).clearContent();
        LogDebug(`SUCCESS TRIM. Cleared rows 127→${LR} in ${SheetName}.`, 'MIN');
      }
      break;

    case SWING_12:
      if (LR > 366) {
        sheet_sr.getRange(367, 1, LR - 366, LC).clearContent();
        LogDebug(`SUCCESS TRIM. Cleared rows 367→${LR} in ${SheetName}.`, 'MIN');
      }
      break;

    case SWING_52:
      LogDebug(`NOTHING TO TRIM. ${SheetName} stays at ${LR} rows.`, 'MIN');
      break;

    default:
      LogDebug(`No trim logic defined for ${SheetName}.`, 'MIN');
  }
}

/////////////////////////////////////////////////////////////////////Hide and Show Sheets/////////////////////////////////////////////////////////////////////
/**
 * Hides or deletes specific sheets based on the asset class defined in the Config!IST.
 *
 * This function needs direct access to the `SpreadsheetApp.getActiveSpreadsheet()` object
 * because it performs actions that require the spreadsheet context itself — such as:
 * - Fetching *all* sheets using `getSheets()`
 * - Deleting sheets using `ss.deleteSheet(sheet)`
 *
 * These operations go beyond simply fetching a sheet by name (which `fetchSheetByName()` handles),
 * so we must declare `const ss = SpreadsheetApp.getActiveSpreadsheet();` here directly.
 *
 * - STOCK: hides specific sheets listed in `Hidden`.
 * - ADR/BDR/ETF: deletes all sheets *not* in the `Keep` set.
 *
 * @returns {void}
 */
function doDisableSheets() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();          // cant remove
  const sheet_co = fetchSheetByName('Config');
  if (!sheet_co) return;

  const Class = getConfigValue(IST, 'Config');                     // IST = asset class

  const sheets = ss.getSheets();

  switch (Class) {
    case 'STOCK': {
      var Hidden = [
        'DATA','Prov_','FIBO','Cotações','UPDATE','Balanço',
        'Balanço Ativo','Balanço Passivo','Resultado','Demonstração',
        'Fluxo','Fluxo de Caixa','Valor','Demonstração do Valor Adicionado'
      ];
      for (let i = 0; i < sheets.length; i++) {
        const sheet = sheets[i];
        const SheetName = sheet.getName();
        if (!sheet.isSheetHidden() && Hidden.indexOf(SheetName) !== -1) {
          sheet.hideSheet();
          LogDebug(`Sheet hidden: ${SheetName}`, 'MIN');
        }
      }
      break;
    }
    case 'ADR': {
      var Keep = new Set([
        'Config','Settings','Index','Preço','FIBO',
        SWING_4, SWING_12, SWING_52,'Cotações'
      ]);
      // reverse order to safely delete
      for (let i = sheets.length - 1; i >= 0; i--) {
        const sheet = sheets[i];
        const SheetName = sheet.getName();
        if (!Keep.has(SheetName)) {
          ss.deleteSheet(sheet);
          LogDebug(`Sheet deleted: ${SheetName}`, 'MIN');
        }
      }
      break;
    }
    case 'BDR':
    case 'ETF': {
      var Keep = new Set([
        'Config','Settings','Index','Prov','Prov_','Preço','FIBO',
        SWING_4, SWING_12, SWING_52,'Cotações','DATA','OPT','Opções','BTC','Termo'
      ]);
      for (let i = sheets.length - 1; i >= 0; i--) {
        const sheet = sheets[i];
        const SheetName = sheet.getName();
        if (!Keep.has(SheetName)) {
          ss.deleteSheet(sheet);
          LogDebug(`Sheet deleted: ${SheetName}`, 'MIN');
        }
      }
      break;
    }
    default:
      LogDebug(`Class "${Class}" not recognized. No sheets modified.`, 'MIN');
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
      LogDebug(`HIDDEN: ${sheet_sr.getName()}`, 'MIN');
    }
    if (sheet_co && !sheet_co.isSheetHidden()) {
      sheet_co.hideSheet();
      LogDebug(`HIDDEN: ${sheet_co.getName()}`, 'MIN');
    }
  }
};

/////////////////////////////////////////////////////////////////////SAVE FUNCTIONS/////////////////////////////////////////////////////////////////////
