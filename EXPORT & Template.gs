/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doExportAll()
{
  doExportBasics();
  doExportExtras();
  doExportFinancials();
};

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

//-------------------------------------------------------------------BASICS-------------------------------------------------------------------//

function doExportBasics() {
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND];
  const totalSheets = SheetNames.length;
  let Count = 0;

  const sheet_co = fetchSheetByName('Config');                                  // Config sheet
  const DEBUG = sheet_co.getRange(DBG).getDisplayValue();                       // DBG = Debug Mode

  if (DEBUG = "TRUE") Logger.log(`Starting export of ${totalSheets} sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100);
    if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) Exporting ${SheetName}...`);

    try {
      doExportBasic(SheetName);
      if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} exported successfully`);
    } catch (error) {
      if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error exporting ${SheetName}: ${error}`);
    }
  });
  if (DEBUG = "TRUE") Logger.log(`Export completed: ${Count} of ${totalSheets} basics exported successfully`);
}

//-------------------------------------------------------------------DATA-------------------------------------------------------------------//

function doExportFinancials() {
  const SheetNames = [BLC, DRE, FLC, DVA];
  const totalSheets = SheetNames.length;
  let Count = 0;

  const sheet_co = fetchSheetByName('Config');                                  // Config sheet
  const DEBUG = sheet_co.getRange(DBG).getDisplayValue();                       // DBG = Debug Mode

  if (DEBUG = "TRUE") Logger.log(`Starting export of ${totalSheets} sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100);
    if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) Exporting ${SheetName}...`);

    try {
      doExportFinancial(SheetName);
      if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} exported successfully`);
    } catch (error) {
      if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error exporting ${SheetName}: ${error}`);
    }
  });
  if (DEBUG = "TRUE") Logger.log(`Export completed: ${Count} of ${totalSheets} Financial sheets exported successfully`);
}

//-------------------------------------------------------------------EXTRAS-------------------------------------------------------------------//

function doExportExtras() {
  const SheetNames = [FUTURE, FUTURE_1, FUTURE_2, FUTURE_3, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13, BLOCK];
  const totalSheets = SheetNames.length;
  let Count = 0;

  const sheet_co = fetchSheetByName('Config');                                  // Config sheet
  const DEBUG = sheet_co.getRange(DBG).getDisplayValue();                       // DBG = Debug Mode

  if (DEBUG = "TRUE") Logger.log(`Starting export of ${totalSheets} sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100);
    if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) Exporting ${SheetName}...`);

    try {
      doExportExtra(SheetName);
      if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} exported successfully`);
    } catch (error) {
      if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error exporting ${SheetName}: ${error}`);
    }
  });
  if (DEBUG = "TRUE") Logger.log(`Export completed: ${Count} of ${totalSheets} extra sheets exported successfully`);
}

/////////////////////////////////////////////////////////////////////SHEETS TEMPLATE/////////////////////////////////////////////////////////////////////

function doExportBasic(SheetName) {
  Logger.log(`EXPORT: ${SheetName}`);

  const sheet_co = fetchSheetByName('Config');                                       // Config sheet
  const sheet_se = fetchSheetByName('Settings');                                     // Settings sheet
  if (!sheet_co || !sheet_se) return;

  const Class = getConfigValue(IST, 'Config');                                       // IST = Is Stock?
  const TKT = getConfigValue(TKR, 'Config');                                         // TKR = Ticket Range
  const Target_Id = getConfigValue(TDR, 'Config');                                   // Target sheet ID
  if (!Target_Id) {
    Logger.log("ERROR EXPORT: Target ID is empty."); 
  return;
  }

  var Minimum = sheet_se.getRange(MIN).getValue();                                   // -1000 - Default
  var Maximum = sheet_se.getRange(MAX).getValue();                                   //  1000 - Default

  const sheet_sr = fetchSheetByName(SheetName);                                      // Source sheet
  if (!sheet_sr) {
    Logger.log(`ERROR EXPORT: Source sheet ${SheetName} - Source sheet does not exist on doExportBasic from sheet_sr`);
    return;
  }

  var [A2, A5] = sheet_sr.getRange("A2:A5").getValues().flat();
  var LC = sheet_sr.getLastColumn();

  const ss_tr = SpreadsheetApp.openById(Target_Id);                                     // Target spreadsheet
  const sheet_tr = ss_tr.getSheetByName(SheetName);                                     // Target sheet - does not use fetchSheetByName, because gets data from diferent spreadsheet
  if (!sheet_tr) {
    Logger.log(`ERROR EXPORT: Target sheet ${SheetName} - Target sheet does not exist on doExportBasic from sheet_tr`);
    return;
  }

  let ShouldExport = false;
  let Export;                                                                          // Declare Export without an initial value

  if (ErrorValues.includes(A2) && A5 == "") {
    Logger.log(`ERROR EXPORT: ${SheetName} - ErrorValues in A2 or A5_ = "" on doExportBasic`);
    return;
  }

  if (Class !== 'STOCK') {
    Logger.log(`ERROR EXPORT: ${SheetName} - Class != STOCK - ${Class} on doExportBasic`);
    return;
  }

  // Export logic specific to each sheet
  switch (SheetName) {
//-------------------------------------------------------------------Swing-------------------------------------------------------------------//
    case SWING_4:
    case SWING_12:
    case SWING_52:
      Export = getConfigValue(ETR);                                             // ETR = Export to Swing
      var C2 = sheet_sr.getRange("C2").getValue();
      if (C2 > 0) {
        ShouldExport = true;
      }
      break;
//-------------------------------------------------------------------Opções-------------------------------------------------------------------//
    case OPCOES:
      Export = getConfigValue(EOP);                                             // EOP = Export to Option
      var [Call, Put] = ["C2", "E2"].map(r => sheet_sr.getRange(r).getValue());
      if (Call != 0 && Put != 0 && Call != "" && Put != "") {
        ShouldExport = true;
      }
      break;
//-------------------------------------------------------------------BTC-------------------------------------------------------------------//
    case BTC:
      Export = getConfigValue(EBT);                                             // EBT = Export to BTC
      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        ShouldExport = true;
      }
      break;
//-------------------------------------------------------------------Termo-------------------------------------------------------------------//
    case TERMO:
      Export = getConfigValue(ETE);                                             // ETE = Export to Termo
      var D2 = sheet_sr.getRange("D2").getValue();
      if (!ErrorValues.includes(D2)) {
        ShouldExport = true;
      }
      break;
//-------------------------------------------------------------------Future-------------------------------------------------------------------//
    case FUTURE:
      Export = getConfigValue(ETF);                                             // ETF = Export to Future
      var C2 = sheet_sr.getRange("C2").getValue();
      var E2 = sheet_sr.getRange("E2").getValue();
      var G2 = sheet_sr.getRange("G2").getValue();
      // Using OR here (if at least one value is valid, then export)
      if (!ErrorValues.includes(C2) || !ErrorValues.includes(E2) || !ErrorValues.includes(G2)) {
        ShouldExport = true;
      }
      break;
    // -------------------- Future Variants -------------------- // NOT USED
    case FUTURE_1:
    case FUTURE_2:
    case FUTURE_3:
      Export = getConfigValue(ETF);                                             // ETF = Export to Future
      var C2 = sheet_sr.getRange("C2").getValue();
      var B2 = sheet_sr.getRange("B2").getValue();
      if (!ErrorValues.includes(B2) && C2 > 0) {
        ShouldExport = true;
      }
      break;
//-------------------------------------------------------------------Fund-------------------------------------------------------------------//
    case FUND:
      Export = getConfigValue(EFU);                                             // EFU = Export to Fund
      var B2 = sheet_sr.getRange("B2").getValue();
      if (!ErrorValues.includes(B2)) {
        ShouldExport = true;
      }
      break;

    default:
      Logger.log(`ERROR EXPORT: ${SheetName} - Sheet name not recognized.`);
      return;
  }
//-------------------------------------------------------------------Foot-------------------------------------------------------------------//
  if (ShouldExport != true) {
    Logger.log(`EXPORT: Skipped ${SheetName} - Conditions for export not met on doExportBasic.`);
    return;
  }

  if (Export != "TRUE") {
        Logger.log(`EXPORT: ${SheetName} - Export on config is set to FALSE on doExportBasic.`);
    return;
  }

  let FilteredData;
  if (SheetName === FUND) {
    // Retrieve data as 2D array and filter row data
    var Data = sheet_sr.getRange(2, 1, 1, LC).getValues(); 
    FilteredData = Data[0].map((Value, ColIndex) => {
      if (ColIndex + 1 < 3) {
        return Value; // Keep the date as-is (columns 1-2)
      } else if (ColIndex + 1 > 62) {
        return Value; // For columns beyond BJ, use original
      } else {
        return (Value > Minimum && Value < Maximum) ? Value : "";
      }
    });
  } else {
    FilteredData = sheet_sr.getRange(2, 1, 1, LC).getValues()[0];
  }
  processExport(TKT, [FilteredData], sheet_tr, SheetName);                           // [FilteredData] instead of Data
}

/////////////////////////////////////////////////////////////////////EXTRA TEMPLATE/////////////////////////////////////////////////////////////////////

function doExportExtra(SheetName) {
  Logger.log(`EXPORT: ${SheetName}`);

  const sheet_co = fetchSheetByName('Config');                                       // Config sheet
  const sheet_se = fetchSheetByName('Settings');                                     // Settings sheet
  if (!sheet_co || !sheet_se) return;
  const sheet_sr = fetchSheetByName(SheetName);                                      // Source sheet

  if (!sheet_sr) {
    Logger.log(`ERROR EXPORT: ${SheetName} - Does not exist on doExportExtra from sheet_sr`);
    return;
  }

  const TKT = getConfigValue(TKR, 'Config');                                         // TKR = Ticket Range
  const Target_Id = getConfigValue(TDR, 'Config');                                   // Target sheet ID
  if (!Target_Id) {
    Logger.log("ERROR EXPORT: Target ID is empty."); 
  return;
  }

  // Mapping export config settings
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
    Logger.log(`EXPORT Skipped: ${SheetName} - Data (A) failed ErrorValues on doExportExtra.`);
    return;
  }

  if (ShouldExport != true) {
    Logger.log(`EXPORT: Skipped ${SheetName} - Conditions for export not met on doExportExtra.`);
    return;
  }

  if (Export != "TRUE") {
        Logger.log(`EXPORT: ${SheetName} - Export on config is set to FALSE on doExportExtra.`);
    return;
  }

  // Determine target sheet
  const target_sh = {
    [RIGHT_1]: 'Right', [RIGHT_2]: 'Right',
    [RECEIPT_9]: 'Receipt', [RECEIPT_10]: 'Receipt',
    [WARRANT_11]: 'Warrant', [WARRANT_12]: 'Warrant', [WARRANT_13]: 'Warrant',
    [BLOCK]: 'Block'
  };

  const ss_tr = SpreadsheetApp.openById(Target_Id);                                   // Target spreadsheet
  const sheet_tr = ss_tr.getSheetByName(target_sh[SheetName] || SheetName);           // Declare sheet_tr outside the conditional scope
  if (!sheet_tr) {
    Logger.log(`ERROR EXPORT: ${SheetName} - Does not exist on doExportFinancial from sheet_tr`);
    return;
  }
  processExport(TKT, Data, sheet_tr, SheetName);
}

/////////////////////////////////////////////////////////////////////FINANCIAL TEMPLATE/////////////////////////////////////////////////////////////////////

function doExportFinancial(SheetName) {
  Logger.log(`EXPORT: ${SheetName}`);

  const sheet_co = fetchSheetByName('Config');                                       // Config sheet
  const sheet_se = fetchSheetByName('Settings');                                     // Settings sheet
  if (!sheet_co || !sheet_se) return;

  const TKT = getConfigValue(TKR, 'Config');                                         // TKR = Ticket Range
  var Target_Id = sheet_co.getRange(TDR).getValues();                                // Target sheet ID
  if (!Target_Id) {
    Logger.log("ERROR EXPORT: Target ID is empty."); 
  return;
  }

  const sheet_sr = fetchSheetByName('Index');                                         // Source sheet
  if (!sheet_sr) {
    Logger.log(`ERROR EXPORT: ${SheetName} - Does not exist on doExportFinancial from sheet_sr`);
    return;
  }

  const ss_tr = SpreadsheetApp.openById(Target_Id);                                    // Target spreadsheet
  const sheet_tr = ss_tr.getSheetByName(SheetName);                                    // Target sheet - does not use fetchSheetByName, because gets data from diferent spreadsheet
  if (!sheet_tr) {
    Logger.log(`ERROR EXPORT: ${SheetName} - Does not exist on doExportFinancial from sheet_tr`);
    return; }

  // Mapping export config settings
  const target_co = {
    [BLC]: EBL, [DRE]: EDR, [FLC]: EFL, [DVA]: EDV
  };

  var Export = getConfigValue(target_co[SheetName]) || FALSE;
  if (Export !== "TRUE") {
    Logger.log(`ERROR EXPORT: ${SheetName} - EXPORT on config is set to FALSE on doExportFinancial`);
    return;
  }

  let Data = [];

  switch (SheetName) 
  {
//-------------------------------------------------------------------BLC-------------------------------------------------------------------//
    case BLC:

    Export = getConfigValue(EBL)                                                      // EBL = Export to BLC

    var A = sheet_co.getRange("B18").getValue();                                      // Balanço Atual

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

    var A = sheet_co.getRange("B18").getValue();                                     // Balanço Atual

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

    var A = sheet_co.getRange("B18").getValue();                                     // Balanço Atual

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

    var A = sheet_co.getRange("B18").getValue();                                     // Balanço Atual
      
    var B = sheet_sr.getRange("B77").getValue();                                     // Receitas
    var C = sheet_sr.getRange("B78").getValue();                                     // Insumos Adquiridos de Terceiros
    var D = sheet_sr.getRange("D77").getValue();                                     // Valor Adicionado Bruto
    var E = sheet_sr.getRange("B79").getValue();                                     // Depreciação, Amortização e Exaustão
    var F = sheet_sr.getRange("D78").getValue();                                     // Valor Adicionado Recebido em Transferência
    var G = sheet_sr.getRange("D79").getValue();                                     // Valor Adicionado Total a Distribuir

    Data.push([A, B, C, D, E, F, G]);

    break;

    default:
      Logger.log(`ERROR EXPORT: ${SheetName} - Invalid sheet name`);
      return;
  }
//-------------------------------------------------------------------Foot-------------------------------------------------------------------//
processExport(TKT, Data, sheet_tr, SheetName);
}

/////////////////////////////////////////////////////////////////////INFO/////////////////////////////////////////////////////////////////////

function doExportInfo() {
  const sheet_co = fetchSheetByName('Config');                      // Config sheet
  const sheet_in = fetchSheetByName('Info');                        // Info sheet

  if (!sheet_co || !sheet_in) return;

  var SheetName = sheet_in.getName();
  Logger.log(`Exporting: ${SheetName}`);

  const Data_Id = getConfigValue(DIR, 'Config');                    // DIR = DATA Source ID
  if (!Data_Id) {
    Logger.log("ERROR EXPORT: Target ID is empty."); 
  return;
  }
  var Exported = sheet_co.getRange(EXR).getDisplayValue();          // EXR = Exported?
  if (Exported === "TRUE") { Logger.log("ERROR EXPORT: already exported."); return;  }

  var A = sheet_co.getRange("B3").getValue();                       // Ticket
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

  var ss_tr = SpreadsheetApp.openById(Data_Id);                    // Target spreadsheet
  var sheet_tr = ss_tr.getSheetByName('Relação');                    // Target sheet

  if (!sheet_tr) { Logger.log(`ERROR EXPORT: Target sheet 'Relação' not found in spreadsheet ID ${Data_Id}`); return; }

  var LR = sheet_tr.getLastRow();

  // Export data to the next available row
  sheet_tr.getRange(LR + 1, 1, 1, Data[0].length).setValues(Data);

  setSheetID(); // Mark as exported

  Logger.log(`SUCCESS EXPORT. Sheet: ${SheetName}.`);
}

/////////////////////////////////////////////////////////////////////PROVENTOS/////////////////////////////////////////////////////////////////////

function doExportProventos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config'); 
  const sheet_pv = fetchSheetByName(PROV);  
  const sheet_ix = fetchSheetByName('Index');

  if (!sheet_co || !sheet_ix || !sheet_pv) return;

  const Class = getConfigValue(IST, 'Config');                      // IST = Is Stock?
  const Target_Id = getConfigValue(TDR, 'Config');                  // Target sheet ID
  if (!Target_Id) {
    Logger.log("ERROR EXPORT: Target ID is empty."); 
  return;
  }

  var SheetName = sheet_pv.getName();
  Logger.log(`Export Proventos: ${SheetName}`);

  var ISIN = sheet_pv.getRange("C61").getDisplayValue().trim();     // Código ISIN
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
    Logger.log(`ERROR EXPORT PROVENTOS: ${SheetName} - Date / ISIN error or missing`);
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
    Logger.log(`ERROR EXPORT: Target sheet 'Poventos' not found in spreadsheet ID ${Target_Id}`);
    return;
  }

  var LR = sheet_tr.getLastRow();

  if (Class !== 'STOCK') {
    Logger.log(`ERROR EXPORT: ${SheetName} - Class != STOCK - ${Class} on doExportProventos`);
    return;
  }

  var nonExportValues = Data[0].slice(3, 13);                        // From index 1 (C) to index 11 (not inclusive), i.e. columns C through L.
  var isAllBlankOrZero = nonExportValues.every(value => value === "" || value === 0);

  if (isAllBlankOrZero) {
    var Search = sheet_tr.getRange("A2:A" + LR).createTextFinder(TKT).findNext();

    if (Search) {
      // Clear the entire row (including TKT)
      var rowToClear = Search.getRow();
      sheet_tr.getRange(rowToClear, 1, 1, Data[0].length + 1).clearContent();
      Logger.log(`CLEARED EXPORT: Entire row for ${TKT} cleared due to all values being blank/zero.`);
    } else {
      Logger.log(`NO ACTION: No existing row found for ${TKT}, and all values are blank/zero.`);
    }
    return; // Stop processing further for this ticker
  } else {
    processExport(TKT, Data, sheet_tr, SheetName);
  }
};

/////////////////////////////////////////////////////////////////////PROCESS EXPORT/////////////////////////////////////////////////////////////////////

function processExport(TKT, Data, sheet_tr, SheetName) {
  if (!Data || Data.length <= 0) {
    Logger.log(`EXPORT: Skipped ${SheetName} - No valid data to export.`);
    return;
  }

  // Get the target sheet's last row
  var LR = sheet_tr.getLastRow();

  // Look for the ticker in column A (starting from row 2)
  var Search = sheet_tr.getRange("A2:A" + LR).createTextFinder(TKT).findNext();

  if (Search) {
    // Update adjacent columns with Data
    Search.offset(0, 1, 1, Data[0].length).setValues(Data);
    Logger.log(`SUCCESS EXPORT. Data for ${TKT} updated on Sheet: ${SheetName}.`);
  } else {
    // Ticker not found; add a new row with the ticker in column A...
    sheet_tr.getRange(LR + 1, 1, 1, 1).setValue(TKT);
    Logger.log(`SUCCESS EXPORT. Ticker: ${TKT} added to ${SheetName}.`);
    // ...and then write Data to the adjacent columns.
    sheet_tr.getRange(LR + 1, 2, 1, Data[0].length).setValues(Data);
    Logger.log(`SUCCESS EXPORT. Data for ${TKT} exported on Sheet: ${SheetName}.`);
  }
}

/////////////////////////////////////////////////////////////////////EXPORT TEMPLATE/////////////////////////////////////////////////////////////////////