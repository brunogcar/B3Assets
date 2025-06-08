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
  const SheetNames = SheetsBasic;
  doEditGroup(SheetNames, doEditSheet, 'basic');
}

//-------------------------------------------------------------------EXTRAS-------------------------------------------------------------------//

function doEditExtras() {
  const SheetNames = SheetsExtra;
  doEditGroup(SheetNames, doEditSheet, 'extra');
}

//-------------------------------------------------------------------FINANCIALS-------------------------------------------------------------------//

function doEditFinancials() {
  const SheetNames = SheetsFinancialFull;
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
      checks: ['C2'],
      conditions: ([c2]) => {
        const Class = getConfigValue(IST, 'Config');       // IST = Is Stock?
        return c2 > 0 && ['STOCK','BDR','ETF','ADR'].includes(Class);
      },
      handler: processEditBasic
    },
    {
      names: [OPCOES],
      editKey: DOP,                                          // DOP = Edit to Option
      checks: ['C2','E2'],
      conditions: ([call, put]) => (call != 0 && put != 0 && call !== '' && put !== ''),
      handler: processEditBasic
    },
    {
      names: [BTC],
      editKey: DBT,                                          // DBT = Edit to BTC
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditBasic
    },
    {
      names: [TERMO],
      editKey: DTE,                                          // DTE = Edit to Termo
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditBasic
    },
    {
      names: [FUND],
      editKey: DFU,                                          // DFU = Edit to Fund
      checks: ['B2'],
      conditions: ([b2]) => !ErrorValues.includes(b2),
      handler: processEditBasic
    },
    {
      names: [FUTURE],
      editKey: DFT,                                          // DFT = Edit to Future
      checks: ['C2','E2','G2'],
      conditions: vals => vals.some(v => !ErrorValues.includes(v)),
      handler: processEditBasic
    },
    {
      names: [FUTURE_1, FUTURE_2, FUTURE_3],
      editKey: DFT,                                          // DFT = Edit to Future
      checks: ['C2'],
      conditions: ([c2]) => !ErrorValues.includes(c2),
      handler: processEditExtra
    },
    {
      names: [RIGHT_1, RIGHT_2],
      editKey: DRT,                                          // DRT = Edit to Right
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditExtra
    },
    {
      names: [RECEIPT_9, RECEIPT_10],
      editKey: DRC,                                          // DRC = Edit to Receipt
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditExtra
    },
    {
      names: [WARRANT_11, WARRANT_12, WARRANT_13],
      editKey: DWT,                                          // DWT = Edit to Warrant
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditExtra
    },
    {
      names: [BLOCK],
      editKey: DBK,                                          // DBK = Edit to Block
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditExtra
    },
      {
      names: [AFTER],
      editKey: DAF,                                          // DTE = Edit to Termo
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processEditBasic
    }
  ];

  const cfg = editTable.find(e => e.names.includes(SheetName));
  if (!cfg) {
    LogDebug(`🚩 ERROR EDIT: ${SheetName} - No entry in editTable: doEditBasic`, 'MIN');
    return;
  }

  const Edit = getConfigValue(cfg.editKey);
  const vals = cfg.checks.map(a1 => sheet_sr.getRange(a1).getValue());

  if (cfg.conditions(vals)) {
    cfg.handler(sheet_sr, SheetName, Edit);
  } else {
    LogDebug(`❌ ERROR EDIT: ${SheetName} - Conditions arent met: doEditBasic`, 'MIN');
  }
}

/////////////////////////////////////////////////////////////////////FINANCIAL TEMPLATE/////////////////////////////////////////////////////////////////////

function doEditFinancial(SheetName) {
  LogDebug(`EDIT: ${SheetName}`, 'MIN');

  const cfg = Object.values(financialMap)
                    .find(c => c.sh_tr === SheetName);
  if (!cfg) {
    LogDebug(`🚩 ERROR EDIT: ${SheetName} - No entry in financialMap: doEditFinancial`, 'MIN');
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
    LogDebug(`❌ ERROR EDIT: ${SheetName} - EDIT is set to FALSE`, 'MIN');
    return;
  }

  // ─── Read & validate dates via helper ───────────────────────────
  const dates = extractAndValidateDates(sheet_tr, sheet_sr, cfg, SheetName, 'EDIT');
  if (!dates) {
    return;
  }
  const { New_tr, Old_tr, New_sr, Old_sr } = dates;

  // Row-specific conditions on source template
  if (cfg.conditions && !cfg.conditions(sheet_sr)) {
    LogDebug(`❌ ERROR EDIT: ${SheetName} - Conditions arent met: doEditFinancial`, 'MIN');
    return;
  }

  if (isValidDate(New_sr)) {
  processEditFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr);
    // Recurse if needed
    if (cfg.recurse) {
      doEditFinancial(cfg.sh_sr);
    }
  }
  else {
    LogDebug(`❌ ERROR EDIT: ${SheetName} - New_sr '${New_sr}' is invalid: doEditFinancial`, 'MIN');
    return;
  }
}

/////////////////////////////////////////////////////////////////////EDIT TEMPLATE/////////////////////////////////////////////////////////////////////
