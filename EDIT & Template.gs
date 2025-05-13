/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doEditAll()
{
  doEditBasics();
  doEditExtras();
  doEditFinancials();
  doIsFormula();
};

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

function doEditGroup(SheetNames, editFunction, label) {
  _doGroup(SheetNames, editFunction, "Editing", "edited", label);
}

//-------------------------------------------------------------------BASICS-------------------------------------------------------------------//

function doEditBasics() {
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND];
  doEditGroup(SheetNames, doEditSheet, 'basic');
}

//-------------------------------------------------------------------EXTRAS-------------------------------------------------------------------//

function doEditExtras() {
  const SheetNames = [FUTURE, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13, BLOCK];
  doEditGroup(SheetNames, doEditSheet, 'extra');
}

//-------------------------------------------------------------------FINANCIALS-------------------------------------------------------------------//

function doEditFinancials() {
  const SheetNames = [BLC, Balanco, DRE, Resultado, FLC, Fluxo, DVA, Valor];
  doEditGroup(SheetNames, doEditFinancial, 'financial');
}

/////////////////////////////////////////////////////////////////////BASIC TEMPLATE/////////////////////////////////////////////////////////////////////

function doEditBasic(SheetName) {
  LogDebug(`EDIT: ${SheetName}`, 'MIN');

  const sheet_sr = fetchSheetByName(SheetName);
  if (!sheet_sr) return;
  Utilities.sleep(2500);

  const editTable = [
    {
      names: [SWING_4, SWING_12, SWING_52],
      editKey: DTR,                                          // DTR = Edit to Swing
      cells: ['C2'],
      test: ([c2]) => {
        const cls = getConfigValue(IST, 'Config');       // IST = Is Stock?
        return c2 > 0 && ['STOCK','BDR','ETF','ADR'].includes(cls);
      },
      handler: processEditBasic
    },
    {
      names: [OPCOES],
      editKey: DOP,                                          // DOP = Edit to Option
      cells: ['C2','E2'],
      test: ([call, put]) => (call != 0 && put != 0 && call !== '' && put !== ''),
      handler: processEditBasic
    },
    {
      names: [BTC],
      editKey: DBT,                                          // DBT = Edit to BTC
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditBasic
    },
    {
      names: [TERMO],
      editKey: DTE,                                          // DTE = Edit to Termo
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditBasic
    },
    {
      names: [FUND],
      editKey: DFU,                                          // DFU = Edit to Fund
      cells: ['B2'],
      test: ([b2]) => !ErrorValues.includes(b2),
      handler: processEditBasic
    },
    {
      names: [FUTURE],
      editKey: DFT,                                          // DFT = Edit to Future
      cells: ['C2','E2','G2'],
      test: vals => vals.some(v => !ErrorValues.includes(v)),
      handler: processEditBasic
    },
    {
      names: [FUTURE_1, FUTURE_2, FUTURE_3],
      editKey: DFT,                                          // DFT = Edit to Future
      cells: ['C2'],
      test: ([c2]) => !ErrorValues.includes(c2),
      handler: processEditExtra
    },
    {
      names: [RIGHT_1, RIGHT_2],
      editKey: DRT,                                          // DRT = Edit to Right
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditExtra
    },
    {
      names: [RECEIPT_9, RECEIPT_10],
      editKey: DRC,                                          // DRC = Edit to Receipt
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditExtra
    },
    {
      names: [WARRANT_11, WARRANT_12, WARRANT_13],
      editKey: DWT,                                          // DWT = Edit to Warrant
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditExtra
    },
    {
      names: [BLOCK],
      editKey: DBK,                                          // DBK = Edit to Block
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditExtra
    }
  ];

  const cfg = editTable.find(e => e.names.includes(SheetName));
  if (!cfg) {
    LogDebug(`ERROR EDIT: ${SheetName} - Unhandled sheet type in doEditBasic`, 'MIN');
    return;
  }

  const Edit = getConfigValue(cfg.editKey);
  const vals = cfg.cells.map(a1 => sheet_sr.getRange(a1).getValue());

  if (cfg.test(vals)) {
    cfg.handler(sheet_sr, SheetName, Edit);
  } else {
    LogDebug(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditBasic`, 'MIN');
  }
}



/////////////////////////////////////////////////////////////////////FINANCIAL TEMPLATE/////////////////////////////////////////////////////////////////////

function doEditFinancial(SheetName) {
  LogDebug(`EDIT: ${SheetName}`, 'MIN');

  const cfg = Object.values(financialMap)
                    .find(c => c.sh_tr === SheetName);
  if (!cfg) {
    LogDebug(`No financialMap entry for ${SheetName}`, 'MIN');
    return;
  }

  const sheet_sr = fetchSheetByName(cfg.sh_sr);
  if (!sheet_sr) return;
  const sheet_tr = cfg.sh_tr === cfg.sh_sr
    ? sheet_sr
    : fetchSheetByName(cfg.sh_tr);
  if (!sheet_tr) return;

  const Edit = getConfigValue(cfg.editKey);
  if (Edit !== "TRUE") {
    LogDebug(`ERROR EDIT: ${SheetName} - EDIT disabled`, 'MIN');
    return;
  }

  const raw_New_tr = sheet_tr.getRange(1, cfg.col_new).getDisplayValue();
  const raw_Old_tr = sheet_tr.getRange(1, cfg.col_old).getDisplayValue();
  LogDebug(`[${cfg.sh_tr}] Raw Dates (TR): New=${raw_New_tr}, Old=${raw_Old_tr}, col_new=${cfg.col_new}, col_old=${cfg.col_old}`, 'MAX');
  const [New_tr, Old_tr] = doFinancialDateHelper([raw_New_tr, raw_Old_tr]);

  // — Read SR dates (with conditional old‐date column) —
  const raw_New_sr = sheet_sr.getRange(1, cfg.col_new).getDisplayValue();
  const oldCol     = cfg.recurse ? cfg.col_old_src : cfg.col_old;
  const raw_Old_sr = sheet_sr.getRange(1, oldCol).getDisplayValue();
  LogDebug(`[${cfg.sh_sr}] Raw Dates (SR): New=${raw_New_sr}, Old=${raw_Old_sr}, col_new=${cfg.col_new}, col_old_src=${oldCol}`, 'MAX');
  const [New_sr, Old_sr] = doFinancialDateHelper([raw_New_sr, raw_Old_sr]);

  LogDebug(`[${SheetName}] Edit dates: SR New=${New_sr}, TR New=${New_tr}`, 'MAX');

  // Row-specific conditions on source template
  if (cfg.conditions && !cfg.conditions(sheet_sr)) {
    LogDebug(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditFinancial`, 'MIN');
    return;
  }

  const validNewDate = New_sr.valueOf() !== "-" && New_sr.valueOf() !== "";

  if (validNewDate) {
  processEditFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr);
    // Recurse if needed
    if (cfg.recurse) {
      doEditFinancial(cfg.sh_sr);
    }
  }
  else {
    LogDebug(`ERROR EDIT: ${SheetName} - New_sr '${New_sr}' is invalid on doEditFinancial`, 'MIN');
    return;
  }
}

/////////////////////////////////////////////////////////////////////EDIT TEMPLATE/////////////////////////////////////////////////////////////////////
