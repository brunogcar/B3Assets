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
  doEditGroup(SheetNames, doEditBasic, 'basic');
}

//-------------------------------------------------------------------EXTRAS-------------------------------------------------------------------//

function doEditExtras() {
  const SheetNames = SheetsExtra;
  doEditGroup(SheetNames, doEditBasic, 'extra');
}

//-------------------------------------------------------------------FINANCIALS-------------------------------------------------------------------//

function doEditFinancials() {
  const SheetNames = SheetsFinancialFull;
  doEditGroup(SheetNames, doEditFinancial, 'financial');
}

/////////////////////////////////////////////////////////////////////BASIC TEMPLATE/////////////////////////////////////////////////////////////////////

const basicEditMap = [
  {
    names: [SWING_4, SWING_12, SWING_52],
    editKey: DTR,
    checks: ['C2'],
    conditions: ([c2], Class) => {
      return c2 > 0 && ['STOCK','BDR','ETF','ADR'].includes(Class);
    },
    handler: processEditBasic
  },
  {
    names: [OPCOES],
    editKey: DOP,
    checks: ['C2','E2'],
    conditions: ([call, put]) => (call != 0 && put != 0 && call !== '' && put !== ''),
    handler: processEditBasic
  },
  {
    names: [BTC],
    editKey: DBT,
    checks: ['D2'],
    conditions: ([d2]) => !ErrorValues.includes(d2),
    handler: processEditBasic
  },
  {
    names: [TERMO],
    editKey: DTE,
    checks: ['D2'],
    conditions: ([d2]) => !ErrorValues.includes(d2),
    handler: processEditBasic
  },
  {
    names: [FUND],
    editKey: DFU,
    checks: ['B2'],
    conditions: ([b2]) => !ErrorValues.includes(b2),
    handler: processEditBasic
  },
  {
    names: [FUTURE],
    editKey: DFT,
    checks: ['C2','E2','G2'],
    conditions: vals => vals.some(v => !ErrorValues.includes(v)),
    handler: processEditBasic
  },
  {
    names: [FUTURE_1, FUTURE_2, FUTURE_3],
    editKey: DFT,
    checks: ['C2'],
    conditions: ([c2]) => !ErrorValues.includes(c2),
    handler: processEditExtra
  },
  {
    names: [RIGHT_1, RIGHT_2],
    editKey: DRT,
    checks: ['D2'],
    conditions: ([d2]) => !ErrorValues.includes(d2),
    handler: processEditExtra
  },
  {
    names: [RECEIPT_9, RECEIPT_10],
    editKey: DRC,
    checks: ['D2'],
    conditions: ([d2]) => !ErrorValues.includes(d2),
    handler: processEditExtra
  },
  {
    names: [WARRANT_11, WARRANT_12, WARRANT_13],
    editKey: DWT,
    checks: ['D2'],
    conditions: ([d2]) => !ErrorValues.includes(d2),
    handler: processEditExtra
  },
  {
    names: [BLOCK],
    editKey: DBK,
    checks: ['D2'],
    conditions: ([d2]) => !ErrorValues.includes(d2),
    handler: processEditExtra
  },
  {
    names: [AFTER],
    editKey: DAF,
    checks: ['D2'],
    conditions: ([d2]) => !ErrorValues.includes(d2),
    handler: processEditBasic
  }
];

const basicEditLookup = Object.fromEntries(
  basicEditMap.flatMap(cfg =>
    cfg.names.map(name => [name, cfg])
  )
);

function doEditBasic(SheetName) {
  LogDebug(`EDIT: ${SheetName}`, 'MIN');

  const sheet_sr = getSheet(SheetName);
  if (!sheet_sr) return;
//  SpreadsheetApp.flush()                                                              //   Utilities.sleep(2500); // 2.5 secs // called from doSaveAll() for exemple instead

  const cfg = basicEditLookup[SheetName];
  if (!cfg) {
    LogDebug(`🚩 ERROR EDIT: ${SheetName} - No entry in basicEditMap: doEditBasic`, 'MIN');
    return;
  }

  const Edit = getConfigValue(cfg.editKey);
  const vals = cfg.checks.map(a1 => sheet_sr.getRange(a1).getValue());
  const Class = getConfigValue(IST, 'Config');                                          // read once here, not inside lambda

  if (cfg.conditions(vals, Class)) {
    if (SheetName === FUND) {
      const Minimum = getConfigValue(MIN, 'Settings');                                  // -500 - Default
      const Maximum = getConfigValue(MAX, 'Settings');                                  //  500 - Default
      const LC = sheet_sr.getLastColumn();
      const row = sheet_sr.getRange(2, 1, 1, LC-1).getValues()[0];

      const filtered = filterFundRow(row, Minimum, Maximum);                            // function in Save - Function

      sheet_sr.getRange(2, 1, 1, LC-1).setValues([filtered]);
    }
    cfg.handler(sheet_sr, SheetName, Edit);
  } else {
    LogDebug(`❌ ERROR EDIT: ${SheetName} - Conditions arent met: doEditBasic`, 'MIN');
  }
}

/////////////////////////////////////////////////////////////////////FINANCIAL TEMPLATE/////////////////////////////////////////////////////////////////////

function doEditFinancial(SheetName) {
  LogDebug(`EDIT: ${SheetName}`, 'MIN');

  const cfg = financialSaveMap[SheetName];
  if (!cfg) {
    LogDebug(`🚩 ERROR EDIT: ${SheetName} - No entry in financialSaveMap: doEditFinancial`, 'MIN');
    return;
  }

  const sheet_sr = getSheet(cfg.sh_sr);
  if (!sheet_sr) return;
  const sheet_tr = cfg.sh_tr === cfg.sh_sr
    ? sheet_sr
    : getSheet(cfg.sh_tr);
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

  processEditFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr);
  // Recurse if needed
  if (cfg.recurse) {
    doEditFinancial(cfg.sh_sr);
  }
}

/////////////////////////////////////////////////////////////////////EDIT TEMPLATE/////////////////////////////////////////////////////////////////////
