/////////////////////////////////////////////////////////////////////SAVE BASICS/////////////////////////////////////////////////////////////////////

function doSaveBasic(SheetName) {
  LogDebug(`SAVE: ${SheetName}`, 'MIN');
  const sheet_sr = fetchSheetByName(SheetName);
  if (!sheet_sr) return;
  Utilities.sleep(2500); // 2.5 secs

  const saveTable = [
    {
      names: [SWING_4, SWING_12, SWING_52],
      saveKey: STR,
      editKey: DTR,
      checks: ['B2','C2'],
      conditions: ([b2, c2]) => {
        const Class = getConfigValue(IST, 'Config');
        if (Class === 'STOCK')     return b2 != 0 && c2 > 0;
        return Class.match(/BDR|ETF|ADR/) && c2 > 0;
      },
      handler: processSaveBasic
    },
    {
      names: [OPCOES],
      saveKey: SOP,
      editKey: DOP,
      checks: ['C2','E2','D2','F2','K3','N3'],
      conditions: ([call, put,call_PM, put_PM, diff_simples, diff_composto]) =>
        call && put &&
        (call_PM || put_PM) &&
        (diff_simples || diff_composto),
      handler: processSaveBasic
    },
    {
      names: [BTC],
      saveKey: SBT,
      editKey: DBT,
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveBasic
    },
    {
      names: [TERMO],
      saveKey: STE,
      editKey: DTE,
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveBasic
    },
    {
      names: [FUND],
      saveKey: SFU,
      editKey: DFU,
      checks: ['B2'],
      conditions: ([b2]) => !ErrorValues.includes(b2),
      handler: processSaveBasic
    },
    {
      names: [FUTURE],
      saveKey: SFT,
      editKey: DFT,
      checks: ['C2','E2','G2'],
      conditions: vals => vals.some(v => !ErrorValues.includes(v)),
      handler: processSaveBasic
    },
    {
      names: [FUTURE_1, FUTURE_2, FUTURE_3],
      saveKey: SFT,
      editKey: DFT,
      checks: ['C2'],
      conditions: ([c2]) => !ErrorValues.includes(c2),
      handler: processSaveExtra
    },
    {
      names: [RIGHT_1, RIGHT_2],
      saveKey: SRT,
      editKey: DRT,
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveExtra
    },
    {
      names: [RECEIPT_9, RECEIPT_10],
      saveKey: SRC,
      editKey: DRC,
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveExtra
    },
    {
      names: [WARRANT_11, WARRANT_12, WARRANT_13],
      saveKey: SWT,
      editKey: DWT,
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveExtra
    },
    {
      names: [BLOCK],
      saveKey: SBK,
      editKey: DBK,
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveExtra
    },
    {
      names: [AFTER],
      saveKey: SAF,
      editKey: DAF,
      checks: ['D2'],
      conditions: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveBasic
    }
  ];

  const cfg = saveTable.find(e => e.names.includes(SheetName));
  if (!cfg) {
    LogDebug(`🚩 ERROR SAVE: ${SheetName} - No entry in saveTable: doSaveBasic`, 'MIN');
    return;
  }

  const Save = getConfigValue(cfg.saveKey);
  const Edit = getConfigValue(cfg.editKey);
  const vals = cfg.checks.map(a1 => sheet_sr.getRange(a1).getValue());

  if (cfg.conditions(vals)) {
    cfg.handler(sheet_sr, SheetName, Save, Edit);
  } else {
    LogDebug(`🚩 ERROR SAVE: ${SheetName} - Conditions arent met: doSaveBasic`, 'MIN');
  }
}

/////////////////////////////////////////////////////////////////////FINANCIAL TEMPLATE/////////////////////////////////////////////////////////////////////

const financialMap = {
  BLC: {
    saveKey: SBL,   editKey: DBL,
    sh_sr: Balanco, sh_tr: BLC,   recurse: true,
    col_new: 2, col_old: 3, col_old_src: 3,
    col_src: 2, col_trg: 2, col_bak: 3,
    checks: ["K3","K4"],
    conditions: sheet => {
      const [B2,B27] = ["B2","B27"].map(r=>sheet.getRange(r).getDisplayValue());
      return B2!=0 && B2!=="" && B27!=0 && B27!=="";
    }
  },
  DRE: {
    saveKey: SDE,   editKey: DDE,
    sh_sr: Resultado, sh_tr: DRE, recurse: true,
    col_new: 2, col_old: 3, col_old_src: 4,
    col_src: 2, col_trg: 2, col_bak: 3,
    checks: ["K5"],
    conditions: sheet => {
      const [B4,B27] = ["B4","B27"].map(r=>sheet.getRange(r).getDisplayValue());
      return B4!=0 && B4!=="" && B27!=0 && B27!=="";
    }
  },
  FLC: {
    saveKey: SFL,   editKey: DFL,
    sh_sr: Fluxo,     sh_tr: FLC, recurse: true,
    col_new: 2, col_old: 3, col_old_src: 4,
    col_src: 2, col_trg: 2, col_bak: 3,
    checks: ["K6"],
    conditions: sheet => {
      const v = sheet.getRange("B2").getDisplayValue();
      return v!=0 && v!=="";
    }
  },
  DVA: {
    saveKey: SDV,   editKey: DDV,
    sh_sr: Valor,     sh_tr: DVA, recurse: true,
    col_new: 2, col_old: 3, col_old_src: 4,
    col_src: 2, col_trg: 2, col_bak: 3,
    checks: ["K7"],
    conditions: sheet => {
      const v = sheet.getRange("B2").getDisplayValue();
      return v!=0 && v!=="";
    }
  },
  Balanco: {
    saveKey: SBL,   editKey: DBL,
    sh_sr: Balanco, sh_tr: Balanco, recurse: false,
    col_new: 2, col_old: 3, col_old_src: 3,
    col_src: 2, col_trg: 3, col_bak: 4,
    conditions: sheet => {
      const [B2,B27] = ["B2","B27"].map(r=>sheet.getRange(r).getDisplayValue());
      return B2!=0 && B2!=="" && B27!=0 && B27!=="";
    }
  },
  Resultado: {
    saveKey: SDE,   editKey: DDE,
    sh_sr: Resultado, sh_tr: Resultado, recurse: false,
    col_new: 3, col_old: 4, col_old_src: 4,
    col_src: 3, col_trg: 4, col_bak: 5,
    conditions: sheet => {
      const [B4,B27] = ["B4","B27"].map(r=>sheet.getRange(r).getDisplayValue());
      return B4!=="" && B27!=0 && B27!=="";
    }
  },
  Fluxo: {
    saveKey: SFL,   editKey: DFL,
    sh_sr: Fluxo,     sh_tr: Fluxo, recurse: false,
    col_new: 3, col_old: 4, col_old_src: 4,
    col_src: 3, col_trg: 4, col_bak: 5,
    conditions: sheet => {
      const v = sheet.getRange("B2").getDisplayValue();
      return v!=0 && v!=="";
    }
  },
  Valor: {
    saveKey: SDV,   editKey: DDV,
    sh_sr: Valor,     sh_tr: Valor, recurse: false,
    col_new: 3, col_old: 4, col_old_src: 4,
    col_src: 3, col_trg: 4, col_bak: 5,
    conditions: sheet => {
      const v = sheet.getRange("B2").getDisplayValue();
      return v!=0 && v!=="";
    }
  }
};

function doSaveFinancial(SheetName) {
  LogDebug(`SAVE: ${SheetName}`, 'MIN');

  const sheet_up = fetchSheetByName('UPDATE');
  if (!sheet_up) return;

  const cfg = Object.values(financialMap)
                    .find(c => c.sh_tr === SheetName);
  if (!cfg) {
    LogDebug(`🚩 ERROR SAVE: ${SheetName} - No entry in financialMap: doSaveFinancial`, 'MIN');
    return;
  }

  const Save = getConfigValue(cfg.saveKey);
  if (Save !== "TRUE") {
    LogDebug(`❌ ERROR SAVE: ${SheetName} - SAVE is set to FALSE`, 'MIN');
    return;
  }

  const sheet_sr = fetchSheetByName(cfg.sh_sr);
  if (!sheet_sr) return;
  const sheet_tr = cfg.sh_tr === cfg.sh_sr
    ? sheet_sr
    : fetchSheetByName(cfg.sh_tr);
  if (!sheet_tr) return;

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

  LogDebug(`[${SheetName}] ⏳ SAVE DATES: SR New=${New_sr}-(${raw_New_sr}), TR New=${New_tr}-(${raw_New_tr})`, 'MAX');

  if (cfg.checks) {
    const checkVals = cfg.checks.map(a => sheet_up.getRange(a).getValue());
    const valid = checkVals.every(v => (v >= 90 && v <= 92) || v === 0 || v > 40000);
    if (!valid) {
      LogDebug(`❌ ERROR SAVE: ${SheetName} - Checks failed: ${JSON.stringify(checkVals)}`, 'MID');
      return;
    }
  }
  if (cfg.conditions && !cfg.conditions(sheet_sr)) {
    LogDebug(`❌ ERROR SAVE: ${SheetName} - Conditions arent met: doSaveFinancial`, 'MIN');
    return;
  }

  const validNewDate = New_sr.valueOf() !== "-" && New_sr.valueOf() !== "";

  if (validNewDate) {
  processSaveFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr);
    // Recurse if needed
    if (cfg.recurse) {
      doSaveFinancial(cfg.sh_sr);
    }
  }
  else {
    LogDebug(`❌ ERROR SAVE: ${SheetName} - New_sr '${New_sr}' is invalid: doSaveFinancial`, 'MIN');
    return;
  }
}

/////////////////////////////////////////////////////////////////////OTHER/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////PROVENTOS TEMPLATE/////////////////////////////////////////////////////////////////////

function doProventos()
{
  doCheckDATA(PROV);
  doGetProventos();
  doSaveProventos();
  doExportProventos();
}

function doSaveProventos() {
  const ProvNames = [
    { name: 'Proventos',  checkCell: "B3",  expectedValue: 'Proventos',     sourceRange: "B3:H60",   targetRange: "B3:H60"   },
    { name: 'Subscrição', checkCell: "L3",  expectedValue: 'Tipo',          sourceRange: "L3:T60",   targetRange: "L3:T60"   },
    { name: 'Ativos',     checkCell: "B64", expectedValue: 'Proventos',     sourceRange: "B64:H200", targetRange: "B64:H200" },
    { name: 'Historico',  checkCell: "L64", expectedValue: 'Tipo de Ativo', dynamicRange: true }
  ];

  for (let i = 0; i < ProvNames.length; i++) {
    const Prov_Values = ProvNames[i];
    _doGroup(
      [ Prov_Values.name ],            // sheetNames (1‑item array)
      () => doSaveProv(Prov_Values),   // callback uses full config object
      "Saving",                // actionLabel
      "saved",                 // resultLabel
      Prov_Values.name                 // groupLabel ← your per‑item name
    );
  }
}

function doSaveProv(Prov_Values) {
  const sheet_sr = fetchSheetByName('Prov_');
  if (!sheet_sr) return;

  const sheet_tr = fetchSheetByName('Prov');
  if (!sheet_tr) return;

  const checkValue = sheet_sr.getRange(Prov_Values.checkCell).getDisplayValue().trim();

  if (checkValue === Prov_Values.expectedValue) {
    let Data;

    if (Prov_Values.dynamicRange) {
      const LR = sheet_sr.getLastRow();
      const LC = sheet_sr.getLastColumn();
      const sourceRange = sheet_sr.getRange(64, 12, LR - 63, LC - 11);
      const targetRange = sheet_tr.getRange(64, 12, LR - 63, LC - 11);

      Data = sourceRange.getValues();
      targetRange.clearContent();
      targetRange.setValues(Data);
    } else {
      const sourceRange = sheet_sr.getRange(Prov_Values.sourceRange);
      const targetRange = sheet_tr.getRange(Prov_Values.targetRange);

      Data = sourceRange.getValues();
      targetRange.clearContent();
      targetRange.setValues(Data);
    }

    LogDebug(`✅ SUCCESS SAVE: ${Prov_Values.name}.`, 'MIN');
  } else {
    LogDebug(`❌ ERROR SAVE: ${Prov_Values.name}, ${Prov_Values.checkCell} != ${Prov_Values.expectedValue}`, 'MIN');
  }
}

function doGetProventos() {
  const sheet_tr = fetchSheetByName('Prov_');
  if (!sheet_tr) return;

  const TKT      = getConfigValue(TKR, 'Config');                     // TKR = Ticket Range
  const ticker   = TKT.substring(0, 4);
  const language = 'pt-br';

  const data = JSON.stringify({ issuingCompany: ticker, language });
  const base64Params = Utilities.base64Encode(data);
  const url = `https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall/GetListedSupplementCompany/${base64Params}`;

  LogDebug(`SAVE: Proventos`, 'MIN');

  LogDebug(`URL: ${url}`, 'MIN');

  let responseText;
  try {
    const response = UrlFetchApp.fetch(url);
    responseText = response.getContentText().trim();
    LogDebug(`API Response: ${responseText}`, 'MIN');
  } catch (error) {
    LogDebug(`❌ ERROR: Failed to fetch API response. ${error}`, 'MIN');
    return;
  }

  if (!responseText) {
    LogDebug(`❌ ERROR: Empty response from API.`, 'MIN');
    return;
  }

  let content;
  try {
    content = JSON.parse(responseText);
  } catch (error) {
    LogDebug(`❌ ERROR: Failed to parse JSON response. ${error}`, 'MIN');
    return;
  }

  if (!content || !content[0]) {
    LogDebug(`❌ ERROR: No data returned from API.`, 'MIN');
    return;
  }

  fillCashDividends(sheet_tr, content[0]?.cashDividends || []);
  fillStockDividends(sheet_tr, content[0]?.stockDividends || []);
  fillSubscriptions(sheet_tr, content[0]?.subscriptions || []);
}


/**
 * Generic function to fill a titled table section with header and data rows.
 * Logs start/finish at MIN, header actions at MID, raw data at MAX.
 * Only writes header and data if at least one row passes validation.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet            The target sheet.
 * @param {string}                       titleCell        A1 cell for the section title.
 * @param {string}                       clearRange       A1 range to clear before writing.
 * @param {string}                       titleText        The title text to set in titleCell.
 * @param {string}                       headerRange      A1 range for the header row.
 * @param {Array<string>}                headers          Array of header labels.
 * @param {Array<Object>}                items            Array of data objects.
 * @param {function(Object): any[]}      mapRow           Function mapping an item to a row array.
 * @param {function(Object, any[]): boolean} validateRow   Predicate to decide if a row should be written.
 * @param {number}                       startRow         1-based row where data begins.
 * @param {number}                       colStart         1-based column where data begins.
 * @param {number}                       maxRows          Maximum number of rows to process.
 * @returns {number} Number of rows written.
 */
function fillSection(
  sheet,            // target sheet
  titleCell,        // titleCell
  clearRange,       // clearRange
  titleText,        // titleText
  headerRange,      // headerRange
  headers,          // headers
  items,            // items
  mapRow,           // mapRow
  validateRow,      // validateRow
  startRow,         // startRow
  colStart,         // colStart
  maxRows           // maxRows
) {
  const colCount = headers.length;

  LogDebug(`${titleText}: ℹ️ Starting with ${items.length} items`, 'MIN');
  LogDebug(`${titleText}: Raw items: ${JSON.stringify(items)}`, 'MAX');

  // Blank slate: clear and title
  sheet.getRange(clearRange).clearContent();
  sheet.getRange(titleCell).setValue(titleText);
  LogDebug(`${titleText}: 🧽Cleared ${clearRange} and set title`, 'MID');

  // Prepare filtered rows
  const rows = items
    .slice(0, maxRows)
    .map(mapRow)
    .filter((row, idx) => {
      const ok = validateRow(items[idx], row);
      if (!ok) LogDebug(`${titleText}: ⚠️ Skipping invalid row #${idx}`, 'MAX');
      return ok;
    });

  // Only write header/data if we have rows
  if (rows.length > 0) {
    sheet.getRange(headerRange).setValues([headers]);
    LogDebug(`${titleText}: 💾 Wrote headers at ${headerRange}`, 'MID');

    LogDebug(`${titleText}: 💾 Writing ${rows.length} rows`, 'MIN');
    sheet.getRange(startRow, colStart, rows.length, colCount).setValues(rows);
  } else {
    LogDebug(`${titleText}: ⚠️ No valid rows — header not written`, 'MIN');
  }
  return rows.length;
}

/** Populates Cash Dividends section. */
function fillCashDividends(sheet_tr, dividends) {
  return fillSection(
    sheet_tr,                                             // target sheet
    'B2',                                                 // titleCell
    'B3:H60',                                             // clearRange
    'Proventos em Dinheiro',                              // titleText
    'B3:H3',                                              // headerRange
    [                                                     // headers
      'Proventos','Código ISIN','Data de Aprovação',
      'Última Data Com','Valor (R$)','Relacionado a','Data de Pagamento'
    ],
    dividends,                                            // items
    d => [                                                // mapRow
      d.label, d.isinCode, d.approvedOn,
      d.lastDatePrior, d.rate, d.relatedTo, d.paymentDate
    ],
    (item, row) => row.some(c => c !== '' && c != null),  // validateRow: any non-blank
    4,                                                    // startRow
    2,                                                    // colStart (B)
    57                                                    // maxRows
  );
}

/** Populates Stock Dividends section. */
function fillStockDividends(sheet_tr, stockDividends) {
  const start = 63;
  return fillSection(
    sheet_tr,                                             // target sheet
    `B${start}`,                                          // titleCell
    `B${start}:G${start + 1 + stockDividends.length}`,    // clearRange
    'Dividendos em Ações',                                // titleText
    `B${start+1}:G${start+1}`,                            // headerRange
    [                                                     // headers
      'Proventos','Código ISIN','Data de Aprovação',
      'Última Data Com','Fator','Ativo Emitido'
    ],
    stockDividends,                                       // items
    d => [                                                // mapRow
      d.label, d.isinCode, d.approvedOn,
      d.lastDatePrior, d.factor, d.assetIssued
    ],
    (item, row) => row.some(c => c !== '' && c != null),  // validateRow
    start + 2,                                            // startRow
    2,                                                    // colStart (B)
    stockDividends.length                                 // maxRows
  );
}

/** Populates Subscriptions section. */
function fillSubscriptions(sheet_tr, subscriptions) {
  return fillSection(
    sheet_tr,                                             // target sheet
    'L2',                                                 // titleCell
    'L3:T60',                                             // clearRange
    'Subscrições',                                        // titleText
    'L3:T3',                                              // headerRange
    [                                                     // headers
      'Tipo','Código ISIN','Data de Aprovação',
      'Última Data Com','Percentual (%)','Ativo Emitido',
      'Preço Emissão (R$)','Período de Negociação','Data de Subscrição'
    ],
    subscriptions,                                        // items
    s => [                                                // mapRow
      s.label, s.isinCode, s.approvedOn,
      s.lastDatePrior, s.percentage, s.assetIssued,
      s.priceUnit, s.tradingPeriod, s.subscriptionDate
    ],
    (item, row) => row.some(c => c !== '' && c != null),  // validateRow
    4,                                                    // startRow
    12,                                                   // colStart (L)
    57                                                    // maxRows
  );
}

/////////////////////////////////////////////////////////////////////CodeCVM/////////////////////////////////////////////////////////////////////

function doGetCodeCVM() {
  const sheet_tr = fetchSheetByName('Info');
  if (!sheet_tr) return;

  const TKT      = getConfigValue(TKR, 'Config');                     // TKR = Ticket Range
  const ticker   = TKT.substring(0, 4);
  const language = 'pt-br';

  const data = JSON.stringify({ issuingCompany: ticker, language });
  const base64Params = Utilities.base64Encode(data);


  LogDebug(`GET Code CVM`, 'MIN');

  const url = `https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall/GetListedSupplementCompany/${base64Params}`;
  LogDebug(`URL: ${url}`, 'MIN');

  let responseText;
  try {
    const response = UrlFetchApp.fetch(url);
    responseText = response.getContentText().trim();
    LogDebug(`API Response: ${responseText}`, 'MID');
  } catch (error) {
    LogDebug(`❌ ERROR: Failed to fetch API response. ${error}`, 'MIN');
    return;
  }

  if (!responseText) {
    LogDebug(`❌ ERROR: Empty response from API.`, 'MIN');
  }

  let content;
  try {
    content = JSON.parse(responseText);
  } catch (error) {
    LogDebug(`❌ ERROR: Failed to parse JSON response. ${error}`, 'MIN');
  }

  if (!content || !content[0]) {
    LogDebug(`❌ ERROR: No data returned from API.`, 'MIN');
  }

  const codeCVM = content[0]?.codeCVM || 'N/A';
  LogDebug(`Extracted codeCVM: ${codeCVM}`, 'MIN');

  sheet_tr.getRange("C3").setValue(codeCVM);
}

/////////////////////////////////////////////////////////////////////SAVE AND SHARES TEMPLATE/////////////////////////////////////////////////////////////////////

function doSaveShares() {
  const sheet_sr = fetchSheetByName('DATA');
  if (!sheet_sr) return;

  try {
    let M1 = sheet_sr.getRange("M1").getValue();
    let M2 = sheet_sr.getRange("M2").getValue();

    LogDebug(`SAVE: Shares and FF`, 'MIN');

    if (!isNaN(M1) && !isNaN(M2) && !ErrorValues.includes(M1) && !ErrorValues.includes(M2)) {
      M1 = Number(M1);
      M2 = Number(M2);

      const Data = sheet_sr.getRange("M1:M2").getValues();
      sheet_sr.getRange("L1:L2").setValues(Data);

      LogDebug(`✅ SUCCESS SAVE: Shares and FF`, 'MIN');
    } else {
      LogDebug(`❌ ERROR SAVE: Invalid values in M1 ${M1} / M2 ${M2}`, 'MIN');
    }
  } catch (error) {
    LogDebug(`❌ ERROR in doSaveShares: ${error.message}`, 'MIN');
  }
}

/////////////////////////////////////////////////////////////////////SAVE TEMPLATE/////////////////////////////////////////////////////////////////////
