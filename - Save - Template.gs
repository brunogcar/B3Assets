/////////////////////////////////////////////////////////////////////SAVE BASICS/////////////////////////////////////////////////////////////////////

function doSaveBasic(SheetName) {
  Logger.log(`SAVE: ${SheetName}`);
  const sheet_sr = fetchSheetByName(SheetName);
  if (!sheet_sr) return;
  Utilities.sleep(2500); // 2.5 secs

  const saveTable = [
    {
      names: [SWING_4, SWING_12, SWING_52],
      saveKey: STR,             // STR = Save to Swing
      editKey: DTR,             // DTR = Edit to Swing
      cells: ['B2','C2'],
      test: ([b2, c2]) => {
        const cls = getConfigValue(IST, 'Config'); // IST = Is Stock?
        if (cls === 'STOCK')     return b2 != 0 && c2 > 0;
        return cls.match(/BDR|ETF|ADR/) && c2 > 0;
      },
      handler: processSaveBasic
    },
    {
      names: [OPCOES],
      saveKey: SOP,             // SOP = Save to Option
      editKey: DOP,             // DOP = Edit to Option
      cells: ['C2','E2','C3','E3','D2','F2','D3','F3','K3','N3'],
      test: ([call, put, call_, put_, callPM, putPM, callPM_, putPM_, diff, diff2]) =>
        call && put &&
        (callPM || putPM) &&
        (diff || diff2),
      handler: processSaveBasic
    },
    {
      names: [BTC],
      saveKey: SBT,             // SBT = Save to BTC
      editKey: DBT,             // DBT = Edit to BTC
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveBasic
    },
    {
      names: [TERMO],
      saveKey: STE,             // STE = Save to Termo
      editKey: DTE,             // DTE = Edit to Termo
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveBasic
    },
    {
      names: [FUND],
      saveKey: SFU,             // SFU = Save to Fund
      editKey: DFU,             // DFU = Edit to Fund
      cells: ['B2'],
      test: ([b2]) => !ErrorValues.includes(b2),
      handler: processSaveBasic
    },
    {
      names: [FUTURE],
      saveKey: SFT,             // SFT = Save to Future
      editKey: DFT,             // DFT = Edit to Future
      cells: ['C2','E2','G2'],
      test: vals => vals.some(v => !ErrorValues.includes(v)),
      handler: processSaveBasic
    },
    {
      names: [FUTURE_1, FUTURE_2, FUTURE_3],
      saveKey: SFT,             // SFT = Save to Future
      editKey: DFT,             // DFT = Edit to Future
      cells: ['C2'],
      test: ([c2]) => !ErrorValues.includes(c2),
      handler: processSaveExtra
    },
    {
      names: [RIGHT_1, RIGHT_2],
      saveKey: SRT,             // SRT = Save to Right
      editKey: DRT,             // DRT = Edit to Right
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveExtra
    },
    {
      names: [RECEIPT_9, RECEIPT_10],
      saveKey: SRC,             // SRC = Save to Receipt
      editKey: DRC,             // DRC = Edit to Receipt
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveExtra
    },
    {
      names: [WARRANT_11, WARRANT_12, WARRANT_13],
      saveKey: SWT,             // SWT = Save to Warrant
      editKey: DWT,             // DWT = Edit to Warrant
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveExtra
    },
    {
      names: [BLOCK],
      saveKey: SBK,             // SBK = Save to Block
      editKey: DBK,             // DBK = Edit to Block
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2),
      handler: processSaveExtra
    }
  ];

  const cfg = saveTable.find(e => e.names.includes(SheetName));
  if (!cfg) { Logger.log(`ERROR SAVE: ${SheetName} - Unhandled sheet type in doSaveBasic`); return; }

  // grab the Save/Edit modes
  const Save = getConfigValue(cfg.saveKey);
  const Edit = getConfigValue(cfg.editKey);

  // read all cells in one go
  const vals = cfg.cells.map(a1 => sheet_sr.getRange(a1).getValue());

  if (cfg.test(vals)) {
    cfg.handler(sheet_sr, SheetName, Save, Edit);
  } else { Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveBasic`); }
}

/////////////////////////////////////////////////////////////////////FINANCIAL TEMPLATE/////////////////////////////////////////////////////////////////////
// sheet_sr and sheet_tr are checked  inside the blocks

function doSaveFinancial(SheetName) {
  Logger.log(`SAVE: ${SheetName}`);
  const sheet_up = fetchSheetByName('UPDATE');
  if (!sheet_up) return;

  let Save, Edit;
  let sheet_tr, sheet_sr;

  switch (SheetName) {
    // -------------------------------------------------------------------BLC -------------------------------------------------------------------//
    case BLC:
      Save = getConfigValue(SBL)                                                     // SBL = Save to BLC
      Edit = getConfigValue(DBL)                                                     // DBL = Edit to BLC

      sheet_tr = fetchSheetByName(BLC);
      if (!sheet_tr) return;

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];
      var [New_tr, Old_tr]  = doFinancialDateHelper(Values_tr);

      sheet_sr = fetchSheetByName(Balanco);
      if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange("B1:C1").getValues()[0];
      var [New_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [B2_sr, B27_sr] = ["B2", "B27"].map(r => sheet_sr.getRange(r).getDisplayValue());

      var CHECK1 = sheet_up.getRange("K3").getValue();
      var CHECK2 = sheet_up.getRange("K4").getValue();

      if (((CHECK1 >= 90 && CHECK1 <= 92) || (CHECK1 == 0 || CHECK1 > 40000)) &&
          ((CHECK2 >= 90 && CHECK2 <= 92) || (CHECK2 == 0 || CHECK1 > 40000))) {
        if ((New_sr.valueOf() != "-" && New_sr.valueOf() != "") &&
            (B2_sr != 0 && B2_sr != "") &&
            (B27_sr != 0 && B27_sr != "")) {
          processSaveFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Save, Edit);
          doSaveFinancial(Balanco);
        } else {
          Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveFinancial`);
        }
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveFinancial`);
      }
      break;
    // -------------------------------------------------------------------Balanço -------------------------------------------------------------------//
    case Balanco:
      Save = getConfigValue(SBL)                                                     // SBL = Save to BLC
      Edit = getConfigValue(DBL)                                                     // DBL = Edit to BLC

      sheet_sr = fetchSheetByName(Balanco);
      if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange("B1:C1").getValues()[0];
      var [New_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [C4_sr, C27_sr] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if ((New_sr.valueOf() != "-" && New_sr.valueOf() != "") &&
          (C4_sr != 0 && C4_sr != "") &&
          (C27_sr != 0 && C27_sr != "")) {
        processSaveFinancial(sheet_sr, sheet_sr, '', '', New_sr, Old_sr, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveFinancial`);
      }
      break;
    // -------------------------------------------------------------------DRE -------------------------------------------------------------------//
    case DRE:
      Save = getConfigValue(SDE)                                                     // SDE = Save to DRE
      Edit = getConfigValue(DDE)                                                     // DDE = Edit to DRE

      sheet_tr = fetchSheetByName(DRE);
      if (!sheet_tr) return;

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];
      var [New_tr, Old_tr]  = doFinancialDateHelper(Values_tr);

      sheet_sr = fetchSheetByName(Resultado);
      if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [New_sr, dud_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [C4_sr, C27_sr] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());
      var CHECK = sheet_up.getRange("K5").getValue();

      if (((CHECK >= 90 && CHECK <= 92) || (CHECK == 0 || CHECK > 40000))) {
        if ((New_sr.valueOf() != "-" && New_sr.valueOf() != "") &&
            (C4_sr != 0 && C4_sr != "") &&
            (C27_sr != 0 && C27_sr != "")) {
          processSaveFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Save, Edit);
          doSaveFinancial(Resultado);
        } else {
          Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveFinancial`);
        }
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveFinancial`);
      }
      break;
    // -------------------------------------------------------------------Resultado -------------------------------------------------------------------//
    case Resultado:
      Save = getConfigValue(SDE)                                                     // SDE = Save to DRE
      Edit = getConfigValue(DDE)                                                     // DDE = Edit to DRE

      sheet_sr = fetchSheetByName(Resultado);
      if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [New_sr, dud_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [C4_sr, C27_sr] = ["C4", "C27"].map(r => sheet_sr.getRange(r).getDisplayValue());
      if (sheet_sr) {
        if ((New_sr.valueOf() != "-" && New_sr.valueOf() != "") &&
            (C4_sr != "") &&
            (C27_sr != 0 && C27_sr != "")) {
          processSaveFinancial(sheet_sr, sheet_sr, '', '', New_sr, Old_sr, Save, Edit);
        } else {
          Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveFinancial`);
        }
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveFinancial`);
      }
      break;
    // -------------------------------------------------------------------FLC -------------------------------------------------------------------//
    case FLC:
      Save = getConfigValue(SFL)                                                     // SFL = Save to FLC
      Edit = getConfigValue(DFL)                                                     // DFL = Edit to FLC

      sheet_tr = fetchSheetByName(FLC);
      if (!sheet_tr) return;

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];
      var [New_tr, Old_tr]  = doFinancialDateHelper(Values_tr);

      sheet_sr = fetchSheetByName(Fluxo);
      if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [New_sr, dud_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [C2_sr] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());
      var CHECK = sheet_up.getRange("K6").getValue();

      if ((CHECK >= 90 && CHECK <= 92) || (CHECK == 0 || CHECK > 40000)) {
        if ((New_sr.valueOf() != "-" && New_sr.valueOf() != "") &&
            (C2_sr != 0 && C2_sr !== "")) {
          processSaveFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Save, Edit);
          doSaveFinancial(Fluxo);
        } else {
          Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveFinancial`);
        }
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveFinancial`);
      }
      break;
    // -------------------------------------------------------------------Fluxo -------------------------------------------------------------------//
    case Fluxo:
      Save = getConfigValue(SFL)                                                     // SFL = Save to FLC
      Edit = getConfigValue(DFL)                                                     // DFL = Edit to FLC

      sheet_sr = fetchSheetByName(Fluxo);
      if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [New_sr, dud_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [C2_sr] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if ((New_sr.valueOf() != "-" && New_sr.valueOf() != "") &&
          (C2_sr != 0 && C2_sr !== "")) {
        processSaveFinancial(sheet_sr, sheet_sr, '', '', New_sr, Old_sr, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveFinancial`);
      }
      break;
    // -------------------------------------------------------------------DVA -------------------------------------------------------------------//
    case DVA:
      Save = getConfigValue(SDV)                                                     // SDV = Save to DVA
      Edit = getConfigValue(DDV)                                                     // DDV = Edit to DVA

      sheet_tr = fetchSheetByName(DVA);
      if (!sheet_tr) return;

      var Values_tr = sheet_tr.getRange("B1:C1").getValues()[0];
      var [New_tr, Old_tr]  = doFinancialDateHelper(Values_tr);

      sheet_sr = fetchSheetByName(Valor);
      if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [New_sr, dud_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [C2_sr] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());
      var CHECK = sheet_up.getRange("K7").getValue();

      if ((CHECK >= 90 && CHECK <= 92) || (CHECK == 0 || CHECK > 40000)) {
        if ((New_sr.valueOf() != "-" && New_sr.valueOf() != "") &&
            (C2_sr != 0 && C2_sr !== "")) {
          processSaveFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Save, Edit);
          doSaveFinancial(Valor);
        } else {
          Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveFinancial`);
        }
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Does not exist on doSaveFinancial`);
      }
      break;
    // -------------------------------------------------------------------Valor -------------------------------------------------------------------//
    case Valor:
      Save = getConfigValue(SDV)                                                     // SDV = Save to DVA
      Edit = getConfigValue(DDV)                                                     // DDV = Edit to DVA

      sheet_sr = fetchSheetByName(Valor);
      if (!sheet_sr) return;

      var Values_sr = sheet_sr.getRange("B1:D1").getValues()[0];
      var [New_sr, dud_sr, Old_sr] = doFinancialDateHelper(Values_sr);

      var [C2_sr] = ["C2"].map(r => sheet_sr.getRange(r).getDisplayValue());

      if ((New_sr.valueOf() != "-" && New_sr.valueOf() != "") &&
          (C2_sr != 0 && C2_sr !== "")) {
        processSaveFinancial(sheet_sr, sheet_sr, '', '', New_sr, Old_sr, Save, Edit);
      } else {
        Logger.log(`ERROR SAVE: ${SheetName} - Conditions arent met on doSaveFinancial`);
      }
      break;
    default:
      Logger.log(`ERROR SAVE: ${SheetName} - Unhandled sheet type in doSaveFinancial`);
      break;
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
  const sheet_sr = fetchSheetByName('Prov_');                                    // Source Sheet
  if (!sheet_sr) return;

  const sheet_tr = fetchSheetByName('Prov');                                     // Target Sheet
  if (!sheet_tr) return;

  const checkValue = sheet_sr.getRange(Prov_Values.checkCell).getDisplayValue().trim();

  if (checkValue === Prov_Values.expectedValue) {
    let Data;

    if (Prov_Values.dynamicRange) {
      const lr = sheet_sr.getLastRow();
      const lc = sheet_sr.getLastColumn();
      const sourceRange = sheet_sr.getRange(64, 12, lr - 63, lc - 11);
      const targetRange = sheet_tr.getRange(64, 12, lr - 63, lc - 11);

      Data = sourceRange.getValues();
      targetRange.clearContent(); // Clear target range before writing data
      targetRange.setValues(Data);
    } else {
      const sourceRange = sheet_sr.getRange(Prov_Values.sourceRange);
      const targetRange = sheet_tr.getRange(Prov_Values.targetRange);

      Data = sourceRange.getValues();
      targetRange.clearContent(); // Clear target range before writing data
      targetRange.setValues(Data);
    }

    Logger.log(`SUCCESS SAVE: ${Prov_Values.name}.`);
  } else {
    Logger.log(`ERROR SAVE: ${Prov_Values.name}, ${Prov_Values.checkCell} != ${Prov_Values.expectedValue}`);
  }
}

function doGetProventos() {
  const sheet_tr = fetchSheetByName('Prov_');
  if (!sheet_tr) return;

  const TKT      = getConfigValue(TKR, 'Config');                                     // TKR = Ticket Range
  const ticker   = TKT.substring(0, 4);
  const language = 'pt-br';

  const data = JSON.stringify({
    issuingCompany: ticker,
    language: language
  });

  const base64Params = Utilities.base64Encode(data);

  const url = `https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall/GetListedSupplementCompany/${base64Params}`;
  Logger.log(`URL: ${url}`);

  let responseText;
  try {
    const response = UrlFetchApp.fetch(url);
    responseText = response.getContentText().trim();
    Logger.log(`API Response: ${responseText}`);
  } catch (error) {
    Logger.log(`ERROR: Failed to fetch API response. ${error}`);
    return; // Exit if the API request fails
  }

  if (!responseText) { Logger.log("ERROR: Empty response from API."); }

  let content;
  try { content = JSON.parse(responseText); }
  catch (error) { Logger.log(`ERROR: Failed to parse JSON response. ${error}`); }

  if (!content || !content[0]) { Logger.log("ERROR: No data returned from API."); }

  fillCashDividends(sheet_tr, content[0]?.cashDividends || []);
  fillStockDividends(sheet_tr, content[0]?.stockDividends || []);
  fillSubscriptions(sheet_tr, content[0]?.subscriptions || []);
}

// Fill Cash Dividends from B2 to B60
function fillCashDividends(sheet_tr, dividends) {
  const headerRange = "B3:H3";
  const startRow = 4;
  const maxRows = 57;

  sheet_tr.getRange("B2").setValue('Proventos em Dinheiro');
  sheet_tr.getRange("B3:H60").clearContent();

  const headers = ['Proventos', 'Código ISIN', 'Data de Aprovação', 'Última Data Com', 'Valor (R$)', 'Relacionado a', 'Data de Pagamento'];
  sheet_tr.getRange(headerRange).setValues([headers]);

  dividends.slice(0, maxRows).forEach((div, i) => {
    sheet_tr.getRange(startRow + i, 2, 1, 7).setValues([[
      div.label, div.isinCode, div.approvedOn, div.lastDatePrior, div.rate, div.relatedTo, div.paymentDate
    ]]);
  });
}

// Fill Stock Dividends from row 63
function fillStockDividends(sheet_tr, stockDividends) {
  const startRow = 63;
  const headerRange = `B${startRow + 1}:G${startRow + 1}`;

  sheet_tr.getRange(`B${startRow}:G${startRow + stockDividends.length + 1}`).clearContent();
  sheet_tr.getRange(`B${startRow}`).setValue("Dividendos em Ações");

  const headers = ['Proventos', 'Código ISIN', 'Data de Aprovação', 'Última Data Com', 'Fator', 'Ativo Emitido'];
  sheet_tr.getRange(headerRange).setValues([headers]);

  stockDividends.forEach((stockDiv, i) => {
    sheet_tr.getRange(startRow + 2 + i, 2, 1, 6).setValues([[
      stockDiv.label, stockDiv.isinCode, stockDiv.approvedOn, stockDiv.lastDatePrior, stockDiv.factor, stockDiv.assetIssued
    ]]);
  });
}

// Fill Subscriptions starting from column L, row 2
function fillSubscriptions(sheet_tr, subscriptions) {
  const headerRange = "L3:T3";
  const startRow = 4;

  sheet_tr.getRange("L2").setValue('Subscrições');
  sheet_tr.getRange("L3:T60").clearContent();

  const headers = ['Tipo', 'Código ISIN', 'Data de Aprovação', 'Última Data Com', 'Percentual (%)', 'Ativo Emitido', 'Preço Emissão (R$)', 'Período de Negociação', 'Data de Subscrição'];
  sheet_tr.getRange(headerRange).setValues([headers]);

  subscriptions.forEach((sub, i) => {
    sheet_tr.getRange(startRow + i, 12, 1, 9).setValues([[
      sub.label, sub.isinCode, sub.approvedOn, sub.lastDatePrior, sub.percentage, sub.assetIssued, sub.priceUnit, sub.tradingPeriod, sub.subscriptionDate
    ]]);
  });
}

/////////////////////////////////////////////////////////////////////CodeCVM/////////////////////////////////////////////////////////////////////

function doGetCodeCVM() {
  const sheet_tr = fetchSheetByName('Info');                                    // Target sheet
  if (!sheet_tr) return;

  const TKT      = getConfigValue(TKR, 'Config');                                // TKR = Ticket Range
  const ticker   = TKT.substring(0, 4);
  const language = 'pt-br';

  const data = JSON.stringify({
    issuingCompany: ticker,
    language: language
  });

  const base64Params = Utilities.base64Encode(data);

  const url = `https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall/GetListedSupplementCompany/${base64Params}`;
  Logger.log("URL:", url);

  let responseText;
  try {
    const response = UrlFetchApp.fetch(url);
    responseText = response.getContentText().trim();
    Logger.log("API Response:", responseText);
  }
  catch (error) {
    Logger.log(`ERROR: Failed to fetch API response. ${error}`);

    return; // Exit if the API request fails
  }

  if (!responseText) { Logger.log("ERROR: Empty response from API."); }

  let content;
  try { content = JSON.parse(responseText); }
  catch (error) { Logger.log(`ERROR: Failed to parse JSON response. ${error}`); }

  if (!content || !content[0]) { Logger.log(`ERROR: No data returned from API.`); }

  const codeCVM = content[0]?.codeCVM || 'N/A';                                     // Default to 'N/A' if codeCVM is missing
  Logger.log(`Extracted codeCVM: ${codeCVM}`);

  // Write to the Info sheet
  sheet_tr.getRange("C3").setValue(codeCVM);
}

/////////////////////////////////////////////////////////////////////SAVE AND SHARES TEMPLATE/////////////////////////////////////////////////////////////////////

function doSaveShares() {
  const sheet_sr = fetchSheetByName('DATA');
  if (!sheet_sr) return;

  try {
    var M1 = sheet_sr.getRange("M1").getValue();
    var M2 = sheet_sr.getRange("M2").getValue();

    Logger.log(`SAVE: Shares and FF`);

    if (!isNaN(M1) && !isNaN(M2) && !ErrorValues.includes(M1) && !ErrorValues.includes(M2)) {
      M1 = Number(M1); // Convert to number if not already
      M2 = Number(M2);

      var Data = sheet_sr.getRange("M1:M2").getValues();
      sheet_sr.getRange("L1:L2").setValues(Data);
      Logger.log(`SUCCESS SAVE: Shares and FF`);
    } else { Logger.log(`ERROR SAVE: Invalid values in M1/M2`); }
  }
  catch (error) { Logger.log(`ERROR in doSaveShares:`, error.message); }
}

/////////////////////////////////////////////////////////////////////SAVE TEMPLATE/////////////////////////////////////////////////////////////////////
