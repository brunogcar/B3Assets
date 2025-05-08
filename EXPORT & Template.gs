/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doExportAll()
{
  doExportBasics();
  doExportExtras();
  doExportFinancials();
};

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

function doExportGroup(SheetNames, exportFunction, label) {
  _doGroup(SheetNames, exportFunction, "Exporting", "exported", label);
}

//-------------------------------------------------------------------BASICS-------------------------------------------------------------------//

function doExportBasics() {
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND];
  doExportGroup(SheetNames, doExportBasic, 'basic');
}

//-------------------------------------------------------------------EXTRAS-------------------------------------------------------------------//

function doExportExtras() {
  const SheetNames = [FUTURE, FUTURE_1, FUTURE_2, FUTURE_3, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13, BLOCK];
  doExportGroup(SheetNames, doExportExtra, 'extra');
}

//-------------------------------------------------------------------FINANCIALS-------------------------------------------------------------------//

function doExportFinancials() {
  const SheetNames = [BLC, DRE, FLC, DVA];
  doExportGroup(SheetNames, doExportFinancial, 'financial');
}

/////////////////////////////////////////////////////////////////////SHEETS TEMPLATE/////////////////////////////////////////////////////////////////////

function doExportBasic(SheetName) {
  LogDebug(`EXPORT: ${SheetName}`, 'MIN');

  const Class     = getConfigValue(IST, 'Config');                                   // IST = Is Stock?
  const TKT       = getConfigValue(TKR, 'Config');                                   // TKR = Ticket Range
  const Target_Id = getConfigValue(TDR, 'Config');                                   // Target sheet ID
  if (!Target_Id) { LogDebug("ERROR EXPORT: Target ID is empty.", 'MIN'); return; }

  const Minimum = getConfigValue(MIN, 'Settings');                                  // -500 - Default
  const Maximum = getConfigValue(MAX, 'Settings');                                  //  500 - Default

  if (Class !== 'STOCK') {
    LogDebug(`ERROR EXPORT: ${SheetName} - Class != STOCK (${Class}) on doExportBasic`, 'MIN');
    return;
  }

  const sheet_sr = fetchSheetByName(SheetName);
  if (!sheet_sr) return;

  const exportTable = [
    {
      names: [SWING_4, SWING_12, SWING_52],
      exportKey: ETR,                                                                   // ETR = Export to Swing
      cells: ['C2'],
      test: ([c2]) => c2 > 0
    },
    {
      names: [OPCOES],
      exportKey: EOP,                                                                   // EOP = Export to Option
      cells: ['C2','E2'],
      test: ([call, put]) => call != 0 && put != 0 && call !== '' && put !== ''
    },
    {
      names: [BTC],
      exportKey: EBT,                                                                   // EBT = Export to BTC
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2)
    },
    {
      names: [TERMO],
      exportKey: ETE,                                                                   // ETE = Export to Termo
      cells: ['D2'],
      test: ([d2]) => !ErrorValues.includes(d2)
    },
    {
      names: [FUTURE],
      exportKey: ETF,                                                                   // ETF = Export to Future
      cells: ['C2','E2','G2'],
      test: vals => vals.some(v => !ErrorValues.includes(v))
    },
    {
      names: [FUTURE_1, FUTURE_2, FUTURE_3],
      exportKey: ETF,                                                                   // ETF = Export to Future
      cells: ['B2','C2'],
      test: ([b2, c2]) => !ErrorValues.includes(b2) && c2 > 0
    },
    {
      names: [FUND],
      exportKey: EFU,                                                                   // EFU = Export to Fund
      cells: ['B2'],
      test: ([b2]) => !ErrorValues.includes(b2)
    }
  ];

  const cfg = exportTable.find(e => e.names.includes(SheetName));
  if (!cfg) {
    LogDebug(`ERROR EXPORT: ${SheetName} - Sheet name not recognized in doExportBasic`, 'MIN');
    return;
  }

  const ss_tr = SpreadsheetApp.openById(Target_Id);                                     // Target spreadsheet
  const sheet_tr = ss_tr.getSheetByName(SheetName);                                     // Target sheet - does not use fetchSheetByName, because gets data from diferent spreadsheet
  if (!sheet_tr) return;

  const vals = cfg.cells.map(a1 => sheet_sr.getRange(a1).getValue());
  if (!cfg.test(vals)) {
    LogDebug(`EXPORT: Skipped ${SheetName} - Conditions for export not met on doExportBasic.`);

    if (SheetName === OPCOES) {
     tryCleanOpcaoExportRow(sheet_tr, TKT);
    }
    return;
  }

  const Export = getConfigValue(cfg.exportKey);
  if (Export !== 'TRUE') {
    LogDebug(`EXPORT: ${SheetName} - Export on config is set to FALSE on doExportBasic.`);
    return;
  }

  const LC = sheet_sr.getLastColumn();
  let filtered;
  if (SheetName === FUND) {
    const row = sheet_sr.getRange(2, 1, 1, LC).getValues()[0];
    filtered = row.map((v, i) => {
      // keep date cols 1–2 and beyond BJ (col 62), else filter
      if (i < 2 || i >= 62) return v;
      return (v > Minimum && v < Maximum) ? v : '';
    });
  } else {
    filtered = sheet_sr.getRange(2, 1, 1, LC).getValues()[0];
  }
  processExport(TKT, [filtered], sheet_tr, SheetName);
}

/////////////////////////////////////////////////////////////////////EXTRA TEMPLATE/////////////////////////////////////////////////////////////////////

function doExportExtra(SheetName) {
  LogDebug(`EXPORT: ${SheetName}`, 'MIN');

  const sheet_sr = fetchSheetByName(SheetName);
  if (!sheet_sr) return;

  const TKT       = getConfigValue(TKR, 'Config');                                   // TKR = Ticket Range
  const Target_Id = getConfigValue(TDR, 'Config');                                   // Target sheet ID
  if (!Target_Id) {
    LogDebug("ERROR EXPORT: Target ID is empty.", 'MIN');
    return;
  }

  const target_co = {
    [RIGHT_1]: ERT,  [RIGHT_2]: ERT,
    [RECEIPT_9]: ERC, [RECEIPT_10]: ERC,
    [WARRANT_11]: EWT, [WARRANT_12]: EWT, [WARRANT_13]: EWT,
    [BLOCK]: EBK
  };

  var Export = getConfigValue(target_co[SheetName]) || FALSE;
//-------------------------------------------------------------------Structure-------------------------------------------------------------------//
  var A = sheet_sr.getRange("A2").getValue();                                        // Data
  var B = sheet_sr.getRange("B2").getValue();                                        // Cotação
  var C = sheet_sr.getRange("C2").getValue();                                        // PM
  var D = sheet_sr.getRange("D2").getValue();                                        // Contratos
  var E = sheet_sr.getRange("E2").getValue();                                        // Mínimo
  var F = sheet_sr.getRange("F2").getValue();                                        // Máximo
  var G = sheet_sr.getRange("G2").getValue();                                        // Volume
  var H = sheet_sr.getRange("H2").getValue();                                        // Negócios
  var I = sheet_sr.getRange("I2").getValue();                                        // Ratio

  var N = sheet_sr.getRange("N2").getValue();                                        // Início
  var O = sheet_sr.getRange("O2").getValue();                                        // Fim

  var J = sheet_sr.getRange("J2").getValue();                                        // Emissão
  var K = sheet_sr.getRange("K2").getValue();                                        // Preço
  var L = sheet_sr.getRange("L2").getValue();                                        // Diff

  var Range = [B, C, D, E, F, G, H, I];

  var hasNonBlankCell = Range.some(cell => cell !== '' && cell !== null);            // Check if at least one cell is not blank

  let Data = [];
  let ShouldExport = false;

  if (hasNonBlankCell && !ErrorValues.some(error => Range.includes(error)))
  {
    Data.push([A, B, C, D, E, F, G, H, I, N, O, J, K, L]);
    ShouldExport = true;                                                             // Set ShouldExport to true if conditions are met
  }
//-------------------------------------------------------------------Foot-------------------------------------------------------------------//
  if (ErrorValues.includes(A)) {
    LogDebug(`EXPORT Skipped: ${SheetName} - Data (A) failed ErrorValues on doExportExtra.`, 'MIN');
    return;
  }

  if (ShouldExport != true) {
    LogDebug(`EXPORT: Skipped ${SheetName} - Conditions for export not met on doExportExtra.`, 'MIN');
    return;
  }

  if (Export != "TRUE") {
        LogDebug(`EXPORT: ${SheetName} - Export on config is set to FALSE on doExportExtra.`, 'MIN');
    return;
  }

  const target_sh = {
    [RIGHT_1]: 'Right', [RIGHT_2]: 'Right',
    [RECEIPT_9]: 'Receipt', [RECEIPT_10]: 'Receipt',
    [WARRANT_11]: 'Warrant', [WARRANT_12]: 'Warrant', [WARRANT_13]: 'Warrant',
    [BLOCK]: 'Block'
  };

  const ss_tr = SpreadsheetApp.openById(Target_Id);                                   // Target spreadsheet
  const sheet_tr = ss_tr.getSheetByName(target_sh[SheetName] || SheetName);           // Declare sheet_tr outside the conditional scope
  if (!sheet_tr) {
    LogDebug(`ERROR EXPORT: ${SheetName} - Does not exist on doExportFinancial from sheet_tr`, 'MIN');
    return;
  }
  processExport(TKT, Data, sheet_tr, SheetName);
}

/////////////////////////////////////////////////////////////////////FINANCIAL TEMPLATE/////////////////////////////////////////////////////////////////////

function doExportFinancial(SheetName) {
  LogDebug(`EXPORT: ${SheetName}`, 'MIN');

  const TKT       = getConfigValue(TKR, 'Config');                                   // TKR = Ticket Range
  const Target_Id = getConfigValue(TDR, 'Config');
  if (!Target_Id) {
    LogDebug("ERROR EXPORT: Target ID is empty.", 'MIN');
    return;
  }

  const sheet_sr = fetchSheetByName('Index');
  if (!sheet_sr) return;

  const ss_tr = SpreadsheetApp.openById(Target_Id);                                    // Target spreadsheet
  const sheet_tr = ss_tr.getSheetByName(SheetName);                                    // Target sheet - does not use fetchSheetByName, because gets data from diferent spreadsheet
  if (!sheet_tr) {
    LogDebug(`ERROR EXPORT: ${SheetName} - Does not exist on doExportFinancial from sheet_tr`, 'MIN');
    return;
  }

  const target_co = {
    [BLC]: EBL, [DRE]: EDR, [FLC]: EFL, [DVA]: EDV
  };

  var Export = getConfigValue(target_co[SheetName]) || FALSE;
  if (Export !== "TRUE") {
    LogDebug(`ERROR EXPORT: ${SheetName} - EXPORT on config is set to FALSE on doExportFinancial`, 'MIN');
    return;
  }

  let Data = [];

  switch (SheetName)
  {
//-------------------------------------------------------------------BLC-------------------------------------------------------------------//
    case BLC:

    Export = getConfigValue(EBL)                                                      // EBL = Export to BLC

    var A = sheet_sr.getRange("D5").getValue();                                       // Balanço Atual

    var B = sheet_sr.getRange("B43").getValue();                                      // Ativo
    var C = sheet_sr.getRange("B44").getValue();                                      // A. Circulante
    var D = sheet_sr.getRange("B45").getValue();                                      // A. Não Circulante
    var E = sheet_sr.getRange("B46").getValue();                                      // Passivo
    var F = sheet_sr.getRange("B47").getValue();                                      // Passivo Circulante
    var G = sheet_sr.getRange("B48").getValue();                                      // Passivo Não Circ
    var H = sheet_sr.getRange("B49").getValue();                                      // Patrim. Líq

    Data.push([A, B, C, D, E, F, G, H]);

    break;
//-------------------------------------------------------------------DRE-------------------------------------------------------------------//
    case DRE:

    Export = getConfigValue(EDR)                                                     // EDR = Export to DRE

    var A = sheet_sr.getRange("D5").getValue();                                      // Balanço Atual

    var B = sheet_sr.getRange("B52").getValue();                                     // Receita Líquida 12 MESES
    var C = sheet_sr.getRange("B53").getValue();                                     // Resultado Bruto 12 MESES
    var D = sheet_sr.getRange("B54").getValue();                                     // EBIT 12 MESES
    var E = sheet_sr.getRange("B55").getValue();                                     // EBITDA 12 MESES
    var F = sheet_sr.getRange("B57").getValue();                                     // Lucro Líquido 12 MESES

    var G = sheet_sr.getRange("D52").getValue();                                     // Receita Líquida 3 MESES
    var H = sheet_sr.getRange("D53").getValue();                                     // Resultado Bruto 3 MESES
    var I = sheet_sr.getRange("D54").getValue();                                     // EBIT 3 MESES
    var J = sheet_sr.getRange("D55").getValue();                                     // EBITDA 3 MESES
    var K = sheet_sr.getRange("D57").getValue();                                     // Lucro Líquido 3 MESES

    Data.push([A, B, C, D, E, F, G, H, I, J, K]);

    break;
//-------------------------------------------------------------------FLC-------------------------------------------------------------------//
    case FLC:

    Export = getConfigValue(EFL)                                                     // EFL = Export to FLC

    var A = sheet_sr.getRange("D5").getValue();                                      // Balanço Atual

    var B = sheet_sr.getRange("B69").getValue();                                     // FCO
    var C = sheet_sr.getRange("B70").getValue();                                     // FCI
    var D = sheet_sr.getRange("B71").getValue();                                     // FCF
    var E = sheet_sr.getRange("B72").getValue();                                     // FCT
    var F = sheet_sr.getRange("B73").getValue();                                     // FCL
    var G = sheet_sr.getRange("B74").getValue();                                     // Saldo Inicial
    var H = sheet_sr.getRange("B75").getValue();                                     // Saldo Final

    Data.push([A, B, C, D, E, F, G, H]);

    break;
//-------------------------------------------------------------------DVA-------------------------------------------------------------------//
    case DVA:

    Export = getConfigValue(EDV)                                                     // EDV = Export to DVA

    var A = sheet_sr.getRange("D5").getValue();                                      // Balanço Atual

    var B = sheet_sr.getRange("B77").getValue();                                     // Receitas
    var C = sheet_sr.getRange("B78").getValue();                                     // Insumos Adquiridos de Terceiros
    var D = sheet_sr.getRange("D77").getValue();                                     // Valor Adicionado Bruto
    var E = sheet_sr.getRange("B79").getValue();                                     // Depreciação, Amortização e Exaustão
    var F = sheet_sr.getRange("D78").getValue();                                     // Valor Adicionado Recebido em Transferência
    var G = sheet_sr.getRange("D79").getValue();                                     // Valor Adicionado Total a Distribuir

    Data.push([A, B, C, D, E, F, G]);

    break;

    default:
      LogDebug(`ERROR EXPORT: ${SheetName} - Invalid sheet name`, 'MIN');
      return;
  }
processExport(TKT, Data, sheet_tr, SheetName);
}

/////////////////////////////////////////////////////////////////////INFO/////////////////////////////////////////////////////////////////////

function doExportInfo() {
  const sheet_in = fetchSheetByName('Info');
  if (!sheet_in) return;

  var SheetName = sheet_in.getName();
  LogDebug(`Exporting: ${SheetName}`, 'MIN');

  const Data_Id = getConfigValue(DIR, 'Config');                     // DIR = DATA Source ID
  if (!Data_Id) {
    LogDebug("ERROR EXPORT: Target ID is empty.", 'MIN');
    return;
  }

  const Exported = getConfigValue(EXR, 'Config');                   // EXR = Exported?
  if (Exported === "TRUE") {
    LogDebug("ERROR EXPORT: already exported.", 'MIN');
    return;
  }

  var A = sheet_in.getRange("C11").getValue();                      // Ticket
  var B = sheet_in.getRange("C3").getValue();                       // Código CVM
  var C = sheet_in.getRange("C4").getValue();                       // CNPJ
  var D = sheet_in.getRange("C5").getValue();                       // Empresa
  var E = sheet_in.getRange("C6").getValue();                       // Razão Social
  var F = sheet_in.getRange("C13").getValue();                      // Tipo de Ação
  var G = sheet_in.getRange("C9").getValue();                       // Listagem
  var H = sheet_in.getRange("C18").getValue();                      // Setor
  var I = sheet_in.getRange("C19").getValue();                      // Subsetor
  var J = sheet_in.getRange("C20").getValue();                      // Segmento
  var K = sheet_in.getRange("C7").getValue();                       // Situação Registro

  // Convert 0 values to blank ("")
  var Data = [[A, B, C, D, E, F, G, H, I, J, K]].map(row => row.map(value => value === 0 ? "" : value));

  var ss_tr = SpreadsheetApp.openById(Data_Id);                     // Target spreadsheet
  var sheet_tr = ss_tr.getSheetByName('Relação');                   // Target sheet

  if (!sheet_tr) {
    LogDebug(`ERROR EXPORT: Target sheet 'Relação' not found in spreadsheet ID ${Data_Id}`, 'MIN');
    return;
  }

  var LR = sheet_tr.getLastRow();

  // Export data to the next available row
  sheet_tr.getRange(LR + 1, 1, 1, Data[0].length).setValues(Data);

  setSheetID();                                                     // Mark as exported

  LogDebug(`SUCCESS EXPORT. Sheet: ${SheetName}.`, 'MIN');
}

/////////////////////////////////////////////////////////////////////PROVENTOS/////////////////////////////////////////////////////////////////////

function doExportProventos() {
  const sheet_pv = fetchSheetByName(PROV);
  if (!sheet_pv) return;

  const sheet_ix = fetchSheetByName('Index');
  if (!sheet_ix) return;

  if (!sheet_ix || !sheet_pv) return;

  const Class     = getConfigValue(IST, 'Config');                  // IST = Is Stock?
  const Target_Id = getConfigValue(TDR, 'Config');                  // Target sheet ID
  if (!Target_Id) {
    LogDebug("ERROR EXPORT: Target ID is empty.", 'MIN');
    return;
  }

  var SheetName = sheet_pv.getName();
  LogDebug(`Export Proventos: ${SheetName}`, 'MIN');

  var ISIN  = sheet_pv.getRange("C61").getDisplayValue().trim();    // Código ISIN
  const TKT = getConfigValue(TKR, 'Config');                        // TKR = Ticket Range

  var B = sheet_pv.getRange("J2").getValue();                       // Date
  var C = sheet_ix.getRange("D2").getValue();                       // Price - Index Sheet
  var D = sheet_ix.getRange("B57").getValue();                      // Lucro - Index Sheet

  var E = sheet_pv.getRange("M67").getValue();                      // DY
  var F = sheet_pv.getRange("M68").getValue();                      // Payout
  var G = sheet_pv.getRange("P67").getValue();                      // EVP - DPA
  var H = sheet_pv.getRange("Q67").getValue();                      // EQP
  var I = sheet_pv.getRange("P68").getValue();                      // EVA
  var J = sheet_pv.getRange("Q68").getValue();                      // EQA
  var K = sheet_pv.getRange("R67").getValue();                      // GVP
  var L = sheet_pv.getRange("S67").getValue();                      // GQP
  var M = sheet_pv.getRange("R68").getValue();                      // GVA
  var N = sheet_pv.getRange("S68").getValue();                      // GQA

  var P = sheet_pv.getRange("N76").getValue();                      // TOTAL Ações
  var Q = sheet_pv.getRange("P76").getValue();                      // TOTAL Proventos

  if (ErrorValues.includes(B) || ErrorValues.includes(ISIN)) {
    LogDebug(`ERROR EXPORT PROVENTOS: ${SheetName} - Date / ISIN error or missing`, 'MIN');
    return;
  }

  let Data;

  if (ErrorValues.includes(P)) {
    Data = [[B, C, D, E, F, G, H, I, J, K, L, M, N]];
  } else {
    Data = [[B, C, D, E, F, G, H, I, J, K, L, M, N, "", P, Q]];
  }

  // Convert any 0 values to blank ("")
  Data = Data.map(row => row.map(value => value === 0 ? "" : value));

  var ss_tr = SpreadsheetApp.openById(Target_Id);
  var sheet_tr = ss_tr.getSheetByName('Poventos');

  if (!sheet_tr) {
    LogDebug(`ERROR EXPORT: Target sheet 'Poventos' not found in spreadsheet ID ${Target_Id}`, 'MIN');
    return;
  }

  var LR = sheet_tr.getLastRow();

  if (Class !== 'STOCK') {
    LogDebug(`ERROR EXPORT: ${SheetName} - Class != STOCK - ${Class} on doExportProventos`, 'MIN');
    return;
  }

  var nonExportValues = Data[0].slice(3, 7);                                              // From index 3 (E) to index 7 (not inclusive), i.e. columns E through H.
  var isBlankOrZero = nonExportValues.some(value => value === "" || value === 0);         // nonExportValues.every to select ALL

  if (isBlankOrZero) {
    var Search = sheet_tr.getRange("A2:A" + LR).createTextFinder(TKT).findNext();

    if (Search) {
      // Clear the entire row (including TKT)
      var rowToClear = Search.getRow();
      sheet_tr.getRange(rowToClear, 1, 1, Data[0].length + 1).clearContent();
      LogDebug(`CLEARED EXPORT: Entire row for ${TKT} cleared due to values being blank/zero.`, 'MIN');
    } else {
      LogDebug(`NO ACTION: No existing row found for ${TKT}, and values are blank/zero.`, 'MIN');
    }
    return; // Stop processing further for this ticker
  } else {
    processExport(TKT, Data, sheet_tr, SheetName);
  }
}

/////////////////////////////////////////////////////////////////////PROCESS EXPORT/////////////////////////////////////////////////////////////////////

function processExport(TKT, Data, sheet_tr, SheetName) {
  if (!Data || Data.length <= 0) {
    LogDebug(`EXPORT: Skipped ${SheetName} - No valid data to export.`, 'MIN');
    return;
  }

  // Get the target sheet's last row
  var LR = sheet_tr.getLastRow();

  // Look for the ticker in column A (starting from row 2)
  var Search = sheet_tr.getRange("A2:A" + LR).createTextFinder(TKT).findNext();

  if (Search) {
    // Update adjacent columns with Data
    Search.offset(0, 1, 1, Data[0].length).setValues(Data);
    LogDebug(`SUCCESS EXPORT. Data for ${TKT} updated on Sheet: ${SheetName}.`, 'MIN');
  } else {
    // Ticker not found; add a new row with the ticker in column A...
    sheet_tr.getRange(LR + 1, 1, 1, 1).setValue(TKT);
    LogDebug(`SUCCESS EXPORT. Ticker: ${TKT} added to ${SheetName}.`, 'MIN');
    // ...and then write Data to the adjacent columns.
    sheet_tr.getRange(LR + 1, 2, 1, Data[0].length).setValues(Data);
    LogDebug(`SUCCESS EXPORT. Data for ${TKT} exported on Sheet: ${SheetName}.`, 'MIN');
  }
}

/////////////////////////////////////////////////////////////////////EXPORT TEMPLATE/////////////////////////////////////////////////////////////////////
