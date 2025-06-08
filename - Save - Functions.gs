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
 * Convert an array of “date‐like” values (strings or Date objects)
 * into timestamps (ms since epoch), or return “” if invalid.
 *
 * @param {Array<string|Date|number>} dateArr
 * @returns {Array<number|string>} [newDateTs, oldDateTs, …]
 */
function doFinancialDateHelper(dateArr) {
  return dateArr.map(v => {
    // 1) If it’s already a Date, grab its ms
    if (v instanceof Date && !isNaN(v)) {
      return v.getTime();
    }
    // 2) If it’s a number (timestamp), leave it
    if (typeof v === 'number' && !isNaN(v)) {
      return v;
    }
    // 3) Everything else -> string
    const str = v != null ? v.toString().trim() : '';
    // DD/MM/YYYY
    if (str.includes('/')) {
      const [d,m,y] = str.split('/');
      if (d && m && y) {
        return new Date(+y, +m - 1, +d).getTime();
      }
    }
    // YYYY-MM-DD
    if (str.includes('-')) {
      const [y,m,d] = str.split('-');
      if (y && m && d) {
        return new Date(+y, +m - 1, +d).getTime();
      }
    }
    // fallback
    return '';
  });
}

/////////////////////////////////////////////////////////////////////CHECK/////////////////////////////////////////////////////////////////////

function doCheckDATAS() {
  const SheetNames = [...SheetsBasic,...SheetsExtra];

  for (let i = 0; i < SheetNames.length; i++) {
    const SheetName = SheetNames[i];
    try {
      doCheckDATA(SheetName);
    } catch (error) {
      LogDebug(`DATA Check: ❌ ERROR: ${SheetName}: ${error}`, 'MIN');
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

  LogDebug(`DATA CHECK Sheet: ${SheetName}`, 'MIN');

  const cfg = {
    [PROV]:       { sheetVar: sheet_p, cell:  "B3",        toggleHide: false, classSheet: false, forceHide:"DEFAULT", cells: null },
    [OPCOES]:     { sheetVar: sheet_o, cell:  "B2",        toggleHide: true,  classSheet: false, forceHide:HOP,       cells: null },
    [SWING_4]:    { sheetVar: sheet_d, cell:  "B3",       toggleHide: false, classSheet: true,  forceHide:HTR ,       cells: null },
    [SWING_12]:   { sheetVar: sheet_d, cell:  "B3",       toggleHide: false, classSheet: true,  forceHide:HTR ,       cells: null },
    [SWING_52]:   { sheetVar: sheet_d, cell:  "B3",       toggleHide: false, classSheet: true,  forceHide:HTR ,       cells: null },
    [BTC]:        { sheetVar: sheet_d, cell:  "B7",        toggleHide: false, classSheet: false, forceHide:HBT,       cells: null },
    [TERMO]:      { sheetVar: sheet_d, cell:  "B28",       toggleHide: false, classSheet: false, forceHide:HTE,       cells: null },
    [FUND]:       { sheetVar: sheet_i, cell:  "D2",        toggleHide: false, classSheet: false, forceHide:HFU,       cells: null },
    [FUTURE]:     { sheetVar: sheet_d, cell:  null,        toggleHide: false, classSheet: false, forceHide:HFT,       cells: ["B36","B37","B38"] },
    [FUTURE_1]:   { sheetVar: sheet_d, cell:  "B36",       toggleHide: false, classSheet: false, forceHide:true,      cells: null },
    [FUTURE_2]:   { sheetVar: sheet_d, cell:  "B37",       toggleHide: false, classSheet: false, forceHide:true,      cells: null },
    [FUTURE_3]:   { sheetVar: sheet_d, cell:  "B38",       toggleHide: false, classSheet: false, forceHide:true,      cells: null },
    [RIGHT_1]:    { sheetVar: sheet_d, cell:  "C42",       toggleHide: false, classSheet: false, forceHide:HRT,       cells: null },
    [RIGHT_2]:    { sheetVar: sheet_d, cell:  "C43",       toggleHide: false, classSheet: false, forceHide:HRT,       cells: null },
    [RECEIPT_9]:  { sheetVar: sheet_d, cell:  "C48",       toggleHide: false, classSheet: false, forceHide:HRC,       cells: null },
    [RECEIPT_10]: { sheetVar: sheet_d, cell:  "C49",       toggleHide: false, classSheet: false, forceHide:HRC,       cells: null },
    [WARRANT_11]: { sheetVar: sheet_d, cell:  "C54",       toggleHide: false, classSheet: false, forceHide:HWT,       cells: null },
    [WARRANT_12]: { sheetVar: sheet_d, cell:  "C55",       toggleHide: false, classSheet: false, forceHide:HWT,       cells: null },
    [WARRANT_13]: { sheetVar: sheet_d, cell:  "C56",       toggleHide: false, classSheet: false, forceHide:HWT,       cells: null },
    [BLOCK]:      { sheetVar: sheet_d, cell:  null,        toggleHide: false, classSheet: false, forceHide:HBK,       cells: ["C60","C61","C62"] },
    [AFTER]:      { sheetVar: sheet_d, cell:  "C66",       toggleHide: false, classSheet: false, forceHide:HAF,       cells: null },
    [BLC]:        { sheetVar: sheet_b, cell:  "B1",        toggleHide: false, classSheet: false, forceHide:HBL,       cells: null },
    [DRE]:        { sheetVar: sheet_r, cell:  "C1",        toggleHide: false, classSheet: false, forceHide:HDE,       cells: null },
    [FLC]:        { sheetVar: sheet_f, cell:  "C1",        toggleHide: false, classSheet: false, forceHide:HFL,       cells: null },
    [DVA]:        { sheetVar: sheet_v, cell:  "C1",        toggleHide: false, classSheet: false, forceHide:HDV,       cells: null },
  }[SheetName];

  if (!cfg) {
    LogDebug(`Sheet Name "${SheetName}" not recognized.`, 'MIN');
    return processCheckDATA(sheet_sr, SheetName, 'FALSE');
  }

  let Check = '';

  // 1) toggleHide case (OPCOES)
  if (cfg.toggleHide) {
    Check = cfg.sheetVar.getRange(cfg.cell).getValue();
    if (Check === '') {
      sheet_o.hideSheet();
      LogDebug(`🔒 HIDDEN: OPT`, 'MIN');
    } else if (sheet_o.isSheetHidden()) {
      sheet_o.showSheet();
      LogDebug(`🔓 DISPLAYED: ${SheetName}`, 'MIN');
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


  let hideOpt;
  if (typeof cfg.forceHide === 'boolean') {
    hideOpt = cfg.forceHide ? 'TRUE' : 'FALSE';
  } else {
    hideOpt = getConfigValue(cfg.forceHide) || 'DEFAULT';
  }
  return processCheckDATA(sheet_sr, SheetName, Check, hideOpt);
}

/////////////////////////////////////////////////////////////////////DO CHECK Process/////////////////////////////////////////////////////////////////////

/**
 * Purely evaluates pass/fail of the data check.
 * @param {*}       Check
 * @param {string}  SheetName
 * @returns {"TRUE"|"FALSE"}
 */
function evaluateCheck(Check, SheetName) {
  const fixedSheets = [BLC, DRE, FLC, DVA];
  if (ErrorValues.includes(Check)) {
    return fixedSheets.includes(SheetName) ? "TRUE" : "FALSE";
  }
  return "TRUE";
}

/**
 * Shows or hides a sheet based on the three-state setting and the check result.
 * @param {Sheet}   sheet_sr
 * @param {string}  SheetName
 * @param {"TRUE"|"FALSE"} result       the outcome from evaluateCheck()
 * @param {"TRUE"|"FALSE"|"DEFAULT"} hideSetting
 */
function applyVisibility(sheet_sr, SheetName, result, hideSetting) {
  if (hideSetting === "TRUE") {
    if (!sheet_sr.isSheetHidden()) {
      sheet_sr.hideSheet();
      LogDebug(`🔐 FORCED HIDE via setting: ${SheetName}`, 'MIN');
    }
    return;
  }
  if (hideSetting === "FALSE") {
    if (sheet_sr.isSheetHidden()) {
      sheet_sr.showSheet();
      LogDebug(`🔓 FORCED SHOW via setting: ${SheetName}`, 'MIN');
    }
    return;
  }
  // DEFAULT: fallback
  if (result === "FALSE") {
    if (!sheet_sr.isSheetHidden()) {
      sheet_sr.hideSheet();
      LogDebug(`🔒 HIDDEN: ${SheetName}`, 'MIN');
    }
  } else {
    if (sheet_sr.isSheetHidden()) {
      sheet_sr.showSheet();
      LogDebug(`🔓 DISPLAYED: ${SheetName}`, 'MIN');
    }
  }
}

function processCheckDATA(sheet_sr, SheetName, Check, hideSetting) {
  const result = evaluateCheck(Check, SheetName);
  applyVisibility(sheet_sr, SheetName, result, hideSetting);
  LogDebug(`DATA Check: ${ result === "TRUE" ? "✅ TRUE" : "❌ FALSE" }: ${SheetName}`, 'MIN');
  return result;
}

/////////////////////////////////////////////////////////////////////TRIM TEMPLATE/////////////////////////////////////////////////////////////////////

function doTrim() {
  const SheetNames = [SWING_4, SWING_12, SWING_52];

  for (let i = 0; i < SheetNames.length; i++) {
    const SheetName = SheetNames[i];
    try {
      doTrimSheet(SheetName);
    } catch (error) {
      LogDebug(`❌ ERROR trimming ${SheetName}: ${error}`, 'MIN');
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
        LogDebug(`✂️ SUCCESS TRIM: Cleared rows 127→${LR} in ${SheetName}.`, 'MIN');
      }
      break;

    case SWING_12:
      if (LR > 366) {
        sheet_sr.getRange(367, 1, LR - 366, LC).clearContent();
        LogDebug(`✂️ SUCCESS TRIM: Cleared rows 367→${LR} in ${SheetName}.`, 'MIN');
      }
      break;

    case SWING_52:
      LogDebug(`NOTHING TO TRIM: ${SheetName} stays at ${LR} rows.`, 'MIN');
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
          LogDebug(`🔒 HIDDEN: ${SheetName}`, 'MIN');
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
          LogDebug(`🗑️ DELETED: ${SheetName}`, 'MIN');
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
          LogDebug(`🗑️ DELETED: ${SheetName}`, 'MIN');
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
      LogDebug(`🔒 HIDDEN: ${sheet_sr.getName()}`, 'MIN');
    }
    if (sheet_co && !sheet_co.isSheetHidden()) {
      sheet_co.hideSheet();
      LogDebug(`🔒 HIDDEN: ${sheet_co.getName()}`, 'MIN');
    }
  }
};

/////////////////////////////////////////////////////////////////////SAVE FUNCTIONS/////////////////////////////////////////////////////////////////////
