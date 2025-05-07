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

  const configs = {
    [BLC]: {
      editKey: DBL,
      targetSheet: BLC,
      sourceSheet: Balanco,
      targetRange: "B1:C1",
      sourceRange: "B1:C1",
      conditions: (sheet) => {
        const [B2, B27] = ["B2", "B27"].map(r => sheet.getRange(r).getDisplayValue());
        return B2 != 0 && B2 !== "" && B27 != 0 && B27 !== "";
      },
      recurse: Balanco
    },
    [Balanco]: {
      editKey: DBL,
      targetSheet: Balanco,
      sourceSheet: Balanco,
      sourceRange: "B1:C1",
      conditions: (sheet) => {
        const [C4, C27] = ["C4", "C27"].map(r => sheet.getRange(r).getDisplayValue());
        return C4 != 0 && C4 !== "" && C27 != 0 && C27 !== "";
      }
    },
    [DRE]: {
      editKey: DDE,
      targetSheet: DRE,
      sourceSheet: Resultado,
      targetRange: "B1:C1",
      sourceRange: "B1:D1",
      conditions: (sheet) => {
        const [C4, C27] = ["C4", "C27"].map(r => sheet.getRange(r).getDisplayValue());
        return C4 != 0 && C4 !== "" && C27 != 0 && C27 !== "";
      },
      recurse: Resultado
    },
    [Resultado]: {
      editKey: DDE,
      targetSheet: Resultado,
      sourceSheet: Resultado,
      sourceRange: "B1:D1",
      conditions: (sheet) => {
        const [C4, C27] = ["C4", "C27"].map(r => sheet.getRange(r).getDisplayValue());
        return C4 !== "" && C27 != 0 && C27 !== "";
      }
    },
    [FLC]: {
      editKey: DFL,
      targetSheet: FLC,
      sourceSheet: Fluxo,
      targetRange: "B1:C1",
      sourceRange: "B1:D1",
      conditions: (sheet) => sheet.getRange("C2").getDisplayValue() != 0 && sheet.getRange("C2").getDisplayValue() !== "",
      recurse: Fluxo
    },
    [Fluxo]: {
      editKey: DFL,
      targetSheet: Fluxo,
      sourceSheet: Fluxo,
      sourceRange: "B1:D1",
      conditions: (sheet) => sheet.getRange("C2").getDisplayValue() != 0 && sheet.getRange("C2").getDisplayValue() !== ""
    },
    [DVA]: {
      editKey: DDV,
      targetSheet: DVA,
      sourceSheet: Valor,
      targetRange: "B1:C1",
      sourceRange: "B1:D1",
      conditions: (sheet) => sheet.getRange("C2").getDisplayValue() != 0 && sheet.getRange("C2").getDisplayValue() !== "",
      recurse: Valor
    },
    [Valor]: {
      editKey: DDV,
      targetSheet: Valor,
      sourceSheet: Valor,
      sourceRange: "B1:D1",
      conditions: (sheet) => sheet.getRange("C2").getDisplayValue() != 0 && sheet.getRange("C2").getDisplayValue() !== ""
    }
  };

  const cfg = configs[SheetName];
  if (!cfg) return;

  const Edit = getConfigValue(cfg.editKey);

  const sheet_sr = fetchSheetByName(cfg.sourceSheet);
  if (!sheet_sr) return;

  const sheet_tr = cfg.targetSheet === cfg.sourceSheet ? sheet_sr : fetchSheetByName(cfg.targetSheet);
  if (!sheet_tr) return;

  const Values_sr = sheet_sr.getRange(cfg.sourceRange).getValues()[0];
  const [New_sr, , Old_sr] = doFinancialDateHelper(Values_sr);

  let New_tr = '', Old_tr = '';
  if (cfg.targetRange) {
    const Values_tr = sheet_tr.getRange(cfg.targetRange).getValues()[0];
    [New_tr, Old_tr] = doFinancialDateHelper(Values_tr);
  }

  const validNewDate = New_sr.valueOf() !== "-" && New_sr.valueOf() !== "";

  if (validNewDate && cfg.conditions(sheet_sr)) {
    processEditFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Edit);
    if (cfg.recurse) doEditFinancial(cfg.recurse);
  } else {
    LogDebug(`ERROR EDIT: ${SheetName} - Conditions arent met on doEditFinancial`, 'MIN');
  }
}

/////////////////////////////////////////////////////////////////////EDIT TEMPLATE/////////////////////////////////////////////////////////////////////
