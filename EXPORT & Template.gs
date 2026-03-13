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
  const SheetNames = SheetsBasic;
  doExportGroup(SheetNames, doExportBasic, 'basic');
}

//-------------------------------------------------------------------EXTRAS-------------------------------------------------------------------//

function doExportExtras() {
  const SheetNames = SheetsExtra;
  doExportGroup(SheetNames, doExportExtra, 'extra');
}

//-------------------------------------------------------------------FINANCIALS-------------------------------------------------------------------//

function doExportFinancials() {
  const SheetNames = SheetsFinancial;
  doExportGroup(SheetNames, doExportFinancial, 'financial');
}

/////////////////////////////////////////////////////////////////////SHEETS TEMPLATE/////////////////////////////////////////////////////////////////////

const basicExportMap = [
  {
    names: [SWING_4, SWING_12, SWING_52],
    exportKey: ETR,
    checks: ['C2'],
    conditions: ([c2]) => c2 > 0
  },
  {
    names: [OPCOES],
    exportKey: EOP,
    checks: ['C2','E2'],
    conditions: ([call, put]) => call != 0 && put != 0 && call !== '' && put !== ''
  },
  {
    names: [BTC],
    exportKey: EBT,
    checks: ['D2'],
    conditions: ([d2]) => !ErrorValues.includes(d2)
  },
  {
    names: [TERMO],
    exportKey: ETE,
    checks: ['D2'],
    conditions: ([d2]) => !ErrorValues.includes(d2)
  },
  {
    names: [FUTURE],
    exportKey: ETF,
    checks: ['C2','E2','G2'],
    conditions: vals => vals.some(v => !ErrorValues.includes(v))
  },
  {
    names: [FUTURE_1, FUTURE_2, FUTURE_3],
    exportKey: ETF,
    checks: ['B2','C2'],
    conditions: ([b2, c2]) => !ErrorValues.includes(b2) && c2 > 0
  },
  {
    names: [FUND],
    exportKey: EFU,
    checks: ['B2'],
    conditions: ([b2]) => !ErrorValues.includes(b2)
  },
  {
    names: [AFTER],
    exportKey: EAF,
    checks: ['D2'],
    conditions: ([d2]) => !ErrorValues.includes(d2)
  }
];

const basicExportLookup = Object.fromEntries(
  basicExportMap.flatMap(cfg =>
    cfg.names.map(name => [name, cfg])
  )
);

function doExportBasic(SheetName) {
  LogDebug(`EXPORT: ${SheetName}`, 'MIN');

  const Class     = getConfigValue(IST, 'Config');                                   // IST = Is Stock?
  const TKT       = getConfigValue(TKR, 'Config');                                   // TKR = Ticket Range
  const Target_Id = getConfigValue(TDR, 'Config');                                   // Target sheet ID
  if (!Target_Id) { LogDebug(`❌ ERROR EXPORT: Target ID is empty.`, 'MIN'); return; }

  if (Class !== 'STOCK') {
    LogDebug(`❌ ERROR EXPORT: ${SheetName} - Class != STOCK (${Class}): doExportBasic`, 'MIN');
    return;
  }

  const sheet_sr = getSheet(SheetName);
  if (!sheet_sr) return;

  const cfg = basicExportLookup[SheetName];
  if (!cfg) {
    LogDebug(`❌ ERROR EXPORT: ${SheetName} - No entry in basicExportMap: doExportBasic`, 'MIN');
    return;
  }

  const ss_tr    = getSpreadsheetById(Target_Id);                                       // Target spreadsheet
  const sheet_tr = ss_tr.getSheetByName(SheetName);                                     // Target sheet - does not use getSheet, because gets data from diferent spreadsheet
  if (!sheet_tr) return;

  const A2      = sheet_sr.getRange('A2').getValue();
  const A5      = sheet_sr.getRange('A5').getValue();

  if (ErrorValues.includes(A2) || A5 === '') {
    LogDebug(`❌ ERROR EXPORT: ${SheetName} - A2 or A5 invalid (A2=${A2}, A5=${A5})`, 'MIN');
    return;
  }

  const vals = cfg.checks.map(a1 => sheet_sr.getRange(a1).getValue());
  if (!cfg.conditions(vals)) {
    LogDebug(`❌ ERROR EXPORT: ${SheetName} - Conditions arent met: doExportBasic.`, 'MIN');

    if (SheetName === OPCOES && [1, 16].includes(new Date().getDate())) {
     tryCleanOpcaoExportRow(sheet_tr, TKT);
    }
    return;
  }

  const Export = getConfigValue(cfg.exportKey);
  if (Export !== 'TRUE') {
    LogDebug(`EXPORT: ${SheetName} - Export is set to FALSE: doExportBasic.`, 'MIN');

    return;
  }

  const LC = sheet_sr.getLastColumn();
  let filtered;
  if (SheetName === FUND) {
    const Minimum = getConfigValue(MIN, 'Settings');                                  // -500 - Default
    const Maximum = getConfigValue(MAX, 'Settings');                                  //  500 - Default
    const row = sheet_sr.getRange(2, 1, 1, LC-1).getValues()[0];

    filtered = filterFundRow(row, Minimum, Maximum);                                  // function in Save - Function

  } else {
    filtered = sheet_sr.getRange(2, 1, 1, LC-1).getValues()[0];
  }
  processExport(TKT, [filtered], sheet_tr, SheetName);
}

/////////////////////////////////////////////////////////////////////EXTRA TEMPLATE/////////////////////////////////////////////////////////////////////

const exportExtraConfig = {
  target_co: {
    [RIGHT_1]: ERT,  [RIGHT_2]: ERT,
    [RECEIPT_9]: ERC, [RECEIPT_10]: ERC,
    [WARRANT_11]: EWT, [WARRANT_12]: EWT, [WARRANT_13]: EWT,
    [BLOCK]: EBK
  },

  target_sh: {
    [RIGHT_1]: 'Right', [RIGHT_2]: 'Right',
    [RECEIPT_9]: 'Receipt', [RECEIPT_10]: 'Receipt',
    [WARRANT_11]: 'Warrant', [WARRANT_12]: 'Warrant', [WARRANT_13]: 'Warrant',
    [BLOCK]: 'Block'
  }
};

function doExportExtra(SheetName) {
  LogDebug(`EXPORT: ${SheetName}`, 'MIN');

  const sheet_sr = getSheet(SheetName);
  if (!sheet_sr) return;

  const Target_Id = getConfigValue(TDR, 'Config');                                   // Target sheet ID
  if (!Target_Id) {
    LogDebug(`❌ EXPORT: Target ID is empty.`, 'MIN');
    return;
  }

  const Export = getConfigValue(exportExtraConfig.target_co[SheetName], 'Config');
//-------------------------------------------------------------------Structure-------------------------------------------------------------------//
  const row = sheet_sr.getRange("A2:O2").getValues()[0];

  const [
    A,      // Data
    B,      // Cotação
    C,      // PM
    D,      // Contratos
    E,      // Mínimo
    F,      // Máximo
    G,      // Volume
    H,      // Negócios
    I,      // Ratio

    J,      // Emissão
    K,      // Preço
    L,      // Diff

    TKT,    // Ticker

    N,      // Início
    O       // Fim
  ] = row;

  if (ErrorValues.includes(A)) {
    LogDebug(`❌ ERROR EXPORT: ${SheetName} - ErrorValues in Data A ${A}: doExportExtra.`, 'MIN');
    return;
  }

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
  if (ShouldExport != true) {
    LogDebug(`❌ ERROR EXPORT: ${SheetName} - Conditions arent met: doExportExtra.`, 'MIN');
    return;
  }

  if (Export != "TRUE") {
        LogDebug(`❌ ERROR EXPORT: ${SheetName} - Export is set to FALSE: doExportExtra.`, 'MIN');
    return;
  }

  const ss_tr    = getSpreadsheetById(Target_Id);                                                      // Target spreadsheet
  const sheet_tr = ss_tr.getSheetByName(exportExtraConfig.target_sh[SheetName] || SheetName);          // Declare sheet_tr outside the conditional scope
  if (!sheet_tr) {
    LogDebug(`❌ ERROR EXPORT: ${SheetName} - Does not exist: doExportBasic`, 'MIN');
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
    LogDebug(`❌ ERROR EXPORT: Target ID is empty.`, 'MIN');
    return;
  }

  const sheet_sr = getSheet('Index');
  if (!sheet_sr) return;

  const ss_tr    = getSpreadsheetById(Target_Id);                                      // Target spreadsheet
  const sheet_tr = ss_tr.getSheetByName(SheetName);                                    // Target sheet - does not use getSheet, because gets data from diferent spreadsheet
  if (!sheet_tr) {
    LogDebug(`❌ ERROR EXPORT: ${SheetName} - Does not exist: doExportFinancial`, 'MIN');
    return;
  }

  const target_co = {
    [BLC]: EBL,                                        // EBL = Export to BLC
    [DRE]: EDR,                                        // EDR = Export to DRE
    [FLC]: EFL,                                        // EFL = Export to FLC
    [DVA]: EDV                                         // EDV = Export to DVA
  };

  const Export = getConfigValue(target_co[SheetName], 'Config');
  if (Export !== "TRUE") {
    LogDebug(`❌ ERROR EXPORT: ${SheetName} - EXPORT is set to FALSE: doExportFinancial`, 'MIN');
    return;
  }

  let Data = [];

  switch (SheetName)
  {
//-------------------------------------------------------------------BLC-------------------------------------------------------------------//
    case BLC:

      var A = sheet_sr.getRange("D5").getValue();                         // Balanço Atual

      var rows = sheet_sr.getRange("B43:B49").getValues();                 // Ativo → Patrim. Líq

      var B = rows[0][0];                                                  // Ativo
      var C = rows[1][0];                                                  // A. Circulante
      var D = rows[2][0];                                                  // A. Não Circulante
      var E = rows[3][0];                                                  // Passivo
      var F = rows[4][0];                                                  // Passivo Circulante
      var G = rows[5][0];                                                  // Passivo Não Circ
      var H = rows[6][0];                                                  // Patrim. Líq

      Data.push([A, B, C, D, E, F, G, H]);

    break;
//-------------------------------------------------------------------DRE-------------------------------------------------------------------//
    case DRE:

      var A = sheet_sr.getRange("D5").getValue();                           // Balanço Atual

      var colB = sheet_sr.getRange("B52:B57").getValues();                   // 12 MESES
      var colD = sheet_sr.getRange("D52:D57").getValues();                   // 3 MESES

      var B = colB[0][0];                                                    // Receita Líquida 12 MESES
      var C = colB[1][0];                                                    // Resultado Bruto 12 MESES
      var D = colB[2][0];                                                    // EBIT 12 MESES
      var E = colB[3][0];                                                    // EBITDA 12 MESES
// colB[4] = Depr & Amort, intentionally skipped
      var F = colB[5][0];                                                    // Lucro Líquido 12 MESES

      var G = colD[0][0];                                                    // Receita Líquida 3 MESES
      var H = colD[1][0];                                                    // Resultado Bruto 3 MESES
      var I = colD[2][0];                                                    // EBIT 3 MESES
      var J = colD[3][0];                                                    // EBITDA 3 MESES
// colD[4] = Depr & Amort, intentionally skipped
      var K = colD[5][0];                                                    // Lucro Líquido 3 MESES

      Data.push([A, B, C, D, E, F, G, H, I, J, K]);

    break;
//-------------------------------------------------------------------FLC-------------------------------------------------------------------//
    case FLC:

      var A = sheet_sr.getRange("D5").getValue();                            // Balanço Atual

      var rows = sheet_sr.getRange("B69:B75").getValues();                    // Fluxo de Caixa

      var B = rows[0][0];                                                     // FCO
      var C = rows[1][0];                                                     // FCI
      var D = rows[2][0];                                                     // FCF
      var E = rows[3][0];                                                     // FCT
      var F = rows[4][0];                                                     // FCL
      var G = rows[5][0];                                                     // Saldo Inicial
      var H = rows[6][0];                                                     // Saldo Final

      Data.push([A, B, C, D, E, F, G, H]);

    break;
//-------------------------------------------------------------------DVA-------------------------------------------------------------------//
    case DVA:

      var A = sheet_sr.getRange("D5").getValue();                                 // Balanço Atual

      var colB = sheet_sr.getRange("B77:B79").getValues();                    // Receitas → Depreciação
      var colD = sheet_sr.getRange("D77:D79").getValues();                    // Valores adicionados

      var B = colB[0][0];                                                     // Receitas
      var C = colB[1][0];                                                     // Insumos Adquiridos de Terceiros
      var D = colB[2][0];                                                     // Depreciação, Amortização e Exaustão

      var E = colD[0][0];                                                     // Valor Adicionado Bruto
      var F = colD[1][0];                                                     // Valor Adicionado Recebido em Transferência
      var G = colD[2][0];                                                     // Valor Adicionado Total a Distribuir

    Data.push([A, B, C, D, E, F, G]);

    break;

    default:
      LogDebug(`❌ ERROR EXPORT: ${SheetName} - Invalid sheet name`, 'MIN');
      return;
  }
processExport(TKT, Data, sheet_tr, SheetName);
}

/////////////////////////////////////////////////////////////////////INFO/////////////////////////////////////////////////////////////////////

function doExportInfo() {
  const sheet_in = getSheet('Info');
  if (!sheet_in) return;

  var SheetName = sheet_in.getName();
  LogDebug(`Exporting: ${SheetName}`, 'MIN');

  const Data_Id = getConfigValue(DIR, 'Config');                     // DIR = DATA Source ID
  if (!Data_Id) {
    LogDebug(`❌ ERROR EXPORT: Target ID is empty.`, 'MIN');
    return;
  }

  const Exported = getConfigValue(EXR, 'Config');                   // EXR = Exported?
  if (Exported === "TRUE") {
    LogDebug(`❌ ERROR EXPORT: already exported.`, 'MIN');
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

  var ss_tr    = getSpreadsheetById(Data_Id);                       // Target spreadsheet
  var sheet_tr = ss_tr.getSheetByName('Relação');                   // Target sheet

  if (!sheet_tr) {
    LogDebug(`❌ ERROR EXPORT: 'Relação' not found in spreadsheet ID ${Data_Id}`, 'MIN');
    return;
  }

  var LR = sheet_tr.getLastRow();

  // Export data to the next available row
  sheet_tr.getRange(LR + 1, 1, 1, Data[0].length).setValues(Data);

  setSheetID();                                                     // Mark as exported

  LogDebug(`✅ SUCCESS EXPORT: ${SheetName}.`, 'MIN');
}

/////////////////////////////////////////////////////////////////////PROVENTOS/////////////////////////////////////////////////////////////////////

function doExportProventos() {
  const sheet_pv = getSheet(PROV);
  const sheet_ix = getSheet('Index');

  if (!sheet_ix || !sheet_pv) return;

  const Class     = getConfigValue(IST, 'Config');                  // IST = Is Stock?
  const Target_Id = getConfigValue(TDR, 'Config');                  // Target sheet ID
  if (!Target_Id) {
    LogDebug(`❌ ERROR EXPORT: Target ID is empty.`, 'MIN');
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
    LogDebug(`❌ ERROR EXPORT PROVENTOS: ${SheetName} - Date / ISIN error or missing`, 'MIN');
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

  var ss_tr    = getSpreadsheetById(Target_Id);
  var sheet_tr = ss_tr.getSheetByName('Proventos');

  if (!sheet_tr) {
    LogDebug(`❌ ERROR EXPORT: 'Proventos' not found in spreadsheet ID ${Target_Id}`, 'MIN');
    return;
  }

  var LR = sheet_tr.getLastRow();

  if (Class !== 'STOCK') {
    LogDebug(`❌ ERROR EXPORT: ${SheetName} - Class != STOCK - ${Class}: doExportProventos`, 'MIN');
    return;
  }

  var nonExportValues = Data[0].slice(4, 7);                                              // From index 4 (F) to index 7 (not inclusive), i.e. columns F through H.
  var isBlankOrZero = nonExportValues.some(value => value === "" || value === 0);         // nonExportValues.every to select ALL

  if (isBlankOrZero) {
    var Search = sheet_tr.getRange("A2:A" + LR).createTextFinder(TKT).findNext();

    if (Search) {
      // Clear the entire row (including TKT)
      var rowToClear = Search.getRow();
      sheet_tr.getRange(rowToClear, 1, 1, Data[0].length + 1).clearContent();
      LogDebug(`🧽 CLEARED EXPORT: Entire row for ${TKT} cleared due to values being blank/zero.`, 'MIN');
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
    LogDebug(`❌ ERROR EXPORT: ${SheetName} - No valid data to export.`, 'MIN');
    return;
  }

  // Get the target sheet's last row
  var LR = sheet_tr.getLastRow();

  // Look for the ticker in column A (starting from row 2)
  var Search = sheet_tr.getRange("A2:A" + LR).createTextFinder(TKT).findNext();

  if (Search) {
    // Update adjacent columns with Data
    Search.offset(0, 1, 1, Data[0].length).setValues(Data);
    LogDebug(`✅ SUCCESS EXPORT. Data for ${TKT} Updated: ${SheetName}.`, 'MIN');
  } else {
    // Ticker not found; add a new row with the ticker in column A...
    sheet_tr.getRange(LR + 1, 1, 1, 1).setValue(TKT);
    LogDebug(`✅ SUCCESS EXPORT. Ticker: ${TKT} Added: ${SheetName}.`, 'MIN');
    // ...and then write Data to the adjacent columns.
    sheet_tr.getRange(LR + 1, 2, 1, Data[0].length).setValues(Data);
    LogDebug(`✅ SUCCESS EXPORT. Data for ${TKT} Exported: ${SheetName}.`, 'MIN');
  }
}

/////////////////////////////////////////////////////////////////////EXPORT TEMPLATE/////////////////////////////////////////////////////////////////////
