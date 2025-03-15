/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doExportAll()
{
  doExportSheets();
  doExportExtras();
  doExportDatas();
};

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////SHEETS/////////////////////////////////////////////////////////////////////

function doExportSheets() {
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND];
  const totalSheets = SheetNames.length;
  let Count = 0;

  Logger.log(`Starting export of ${totalSheets} sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100);
    Logger.log(`[${Count}/${totalSheets}] (${progress}%) Exporting ${SheetName}...`);

    try {
      doExportSheet(SheetName);
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} exported successfully`);
    } catch (error) {
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error exporting ${SheetName}: ${error}`);
    }
  });
  Logger.log(`Export completed: ${Count} of ${totalSheets} sheets exported successfully`);
}

/////////////////////////////////////////////////////////////////////DATAS/////////////////////////////////////////////////////////////////////

function doExportDatas() {
  const SheetNames = [BLC, DRE, FLC, DVA];
  const totalSheets = SheetNames.length;
  let Count = 0;

  Logger.log(`Starting export of ${totalSheets} sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100);
    Logger.log(`[${Count}/${totalSheets}] (${progress}%) Exporting ${SheetName}...`);

    try {
      doExportData(SheetName);
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} exported successfully`);
    } catch (error) {
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error exporting ${SheetName}: ${error}`);
    }
  });
  Logger.log(`Export completed: ${Count} of ${totalSheets} data sheets exported successfully`);
}

/////////////////////////////////////////////////////////////////////EXTRAS/////////////////////////////////////////////////////////////////////

function doExportExtras() {
  const SheetNames = [FUTURE, FUTURE_1, FUTURE_2, FUTURE_3, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13, BLOCK];
  const totalSheets = SheetNames.length;
  let Count = 0;

  Logger.log(`Starting export of ${totalSheets} sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100);
    Logger.log(`[${Count}/${totalSheets}] (${progress}%) Exporting ${SheetName}...`);

    try {
      doExportExtra(SheetName);
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} exported successfully`);
    } catch (error) {
      Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error exporting ${SheetName}: ${error}`);
    }
  });
  Logger.log(`Export completed: ${Count} of ${totalSheets} extra sheets exported successfully`);
}

/////////////////////////////////////////////////////////////////////SHEETS TEMPLATE/////////////////////////////////////////////////////////////////////

function doExportSheet(SheetName) 
{
  Logger.log('EXPORT:', SheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');                                        // Config sheet
    var Class = sheet_co.getRange(IST).getDisplayValue();                             // IST = Is Stock? 
    var TKT = sheet_co.getRange(TKR).getValue();                                      // TKR = Ticket Range
    var Target_Id = sheet_co.getRange(TDR).getValues();                               // Target sheet ID

  const sheet_se = fetchSheetByName('Settings');                                      // Settings sheet
    var Minimum = sheet_se.getRange(MIN).getValue();                                  // -1000 - Default
    var Maximum = sheet_se.getRange(MAX).getValue();                                  //  1000 - Default
  if (!sheet_co || !sheet_se) return;

  const sheet_sr = fetchSheetByName(SheetName);                                       // Source sheet
  if (!sheet_sr) { Logger.log('ERROR EXPORT: Source sheet', SheetName, 'Source sheet does not exist on doExportSheet from sheet_sr'); return; }
    var A2 = sheet_sr.getRange('A2').getValue();
    var A5 = sheet_sr.getRange('A5').getValue();
    var LR_S = sheet_sr.getLastRow();
    var LC_S = sheet_sr.getLastColumn();

  const trg = SpreadsheetApp.openById(Target_Id);                                     // Target spreadsheet
  const sheet_tr = trg.getSheetByName(SheetName);                                     // Target sheet - does not use fetchSheetByName, because gets data from diferent spreadsheet
  if (!sheet_tr) { Logger.log('ERROR EXPORT: Target sheet', SheetName, 'Target sheet does not exist on doExportSheet from sheet_tr'); return; }
    var LR_T = sheet_tr.getLastRow();
    var LC_T = sheet_tr.getLastColumn();

  let ShouldExport = false;
  let Export;                 // Declare Export without an initial value

  if(Class == 'STOCK')
  {
    if(!ErrorValues.includes(A2) && A5 != "")
    {
      // Export logic specific to each sheet
      switch (SheetName) 
      {

//-------------------------------------------------------------------Swing-------------------------------------------------------------------//

        case SWING_4:
        case SWING_12:
        case SWING_52:

        Export = getConfigValue(ETR)                                                  // ETR = Export to Swing

        var C2 = sheet_sr.getRange('C2').getValue();

        if (Class == 'STOCK') 
        {
          if( C2 > 0 )
          {
            ShouldExport = true;
          }
        }
          break;

//-------------------------------------------------------------------Opções-------------------------------------------------------------------//

        case OPCOES:

        Export = getConfigValue(EOP)                                                  // EOP = Export to Option

        var [Call, Put] = ['C2', 'E2'].map(r => sheet_sr.getRange(r).getValue());

        if( ( Call != 0 && Put != 0 ) &&
            ( Call != "" && Put != "" ) )
        {
          ShouldExport = true;
        }
        break;

//-------------------------------------------------------------------BTC-------------------------------------------------------------------//

        case BTC:

        Export = getConfigValue(EBT)                                                  // EBT = Export to BTC

        var D2 = sheet_sr.getRange('D2').getValue();

        if( !ErrorValues.includes(D2) )
        {
          ShouldExport = true;
        }
        break;

//-------------------------------------------------------------------Termo-------------------------------------------------------------------//

        case TERMO:

        Export = getConfigValue(ETE)                                                  // ETE = Export to Termo

        var D2 = sheet_sr.getRange('D2').getValue();

        if( !ErrorValues.includes(D2) )
        {
          ShouldExport = true;
        }
        break;

//-------------------------------------------------------------------Future-------------------------------------------------------------------//

        case FUTURE:

        Export = getConfigValue(ETF)                                                  // ETF = Export to Future

        var C2 = sheet_sr.getRange('C2').getValue();
        var E2 = sheet_sr.getRange('E2').getValue();
        var G2 = sheet_sr.getRange('G2').getValue();

        if( ( !ErrorValues.includes(C2) || !ErrorValues.includes(E2) || !ErrorValues.includes(G2) ) )
        {
          ShouldExport = true;
        }
        break;

/////////////////////////////////////////////////////////////////////NOT USED/////////////////////////////////////////////////////////////////////

        case FUTURE_1:
        case FUTURE_2:
        case FUTURE_3:

        Export = getConfigValue(ETF)                                                  // ETF = Export to Future

        var C2 = sheet_sr.getRange('C2').getValue();
        var B2 = sheet_sr.getRange('B2').getValue();

        if( ( !ErrorValues.includes(B2) ) && 
            ( C2 > 0 ))
        {
          ShouldExport = true;
        }
        break;

//-------------------------------------------------------------------Fund-------------------------------------------------------------------//

        case FUND:

        Export = getConfigValue(EFU)                                                  // EFU = Export to Fund

        var B2 = sheet_sr.getRange('B2').getValue();

        if( !ErrorValues.includes(B2) )
        {
          ShouldExport = true;
        }
        break;

        default:
          Logger.log('ERROR EXPORT:', SheetName, 'Sheet name not recognized.');
          return;
      }

//-------------------------------------------------------------------Foot-------------------------------------------------------------------//

      if (Export == 'TRUE' && ShouldExport) 
      {
        let FilteredData;

        if (SheetName === FUND) 
        {
          var Data = sheet_sr.getRange(2, 1, 1, LC_S).getValues();                    // 2D array
          FilteredData = Data[0].map((Value, ColIndex) => 
          {
            if (ColIndex + 1 < 3) 
            {
              return Value; // Keep the date as-is
            } 
            else if (ColIndex + 1 > 62) 
            {
              return Value; // Use the original value for columns > BJ
            } 
            else 
            {
              return (Value > Minimum && Value < Maximum) ? Value : "";               // Apply filtering for columns <= BJ (except column B)
            }
          });
        } 
        else 
        {
          FilteredData = sheet_sr.getRange(2, 1, 1, LC_S).getValues()[0];             // Use unfiltered data
        }

        var Search = sheet_tr.getRange('A2:A' + LR_T).createTextFinder(TKT).findNext();

        if (Search) 
        {
          Search.offset(0, 1, 1, FilteredData.length).setValues([FilteredData]);      // Ensure it's a 2D array
          Logger.log(`SUCCESS EXPORT. Data for ${TKT}. Sheet: ${SheetName}.`);
        } 
        else 
        {
          var NewRow = sheet_tr.getRange(LR_T + 1, 1, 1, 1).setValue([TKT]);
          Logger.log(`SUCCESS EXPORT. Ticker: ${TKT}. Sheet: ${SheetName}.`);

          NewRow.offset(0, 1, 1, FilteredData.length).setValues([FilteredData]);      // Ensure it's a 2D array
          Logger.log(`SUCCESS EXPORT. Data for ${TKT}. Sheet: ${SheetName}.`);
        }
      }
      else 
      {
        Logger.log('ERROR EXPORT:', SheetName, 'EXPORT on config is set to FALSE or Conditions arent met on doExportSheet');
      }
    }
    else 
    {
      Logger.log('ERROR EXPORT:', SheetName, 'ErrorValues in A2 or A5_ = "" on doExportSheet');
    }
  }
  else 
  {
    Logger.log('ERROR EXPORT:', SheetName, 'Class != STOCK', Class);
  }
}

/////////////////////////////////////////////////////////////////////DATA TEMPLATE/////////////////////////////////////////////////////////////////////

function doExportData(SheetName) 
{
  Logger.log('EXPORT:', SheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');                                        // Config sheet
    var TKT = sheet_co.getRange(TKR).getValue();                                      // TKR = Ticket Range
    var Target_Id = sheet_co.getRange(TDR).getValues();                               // Target sheet ID

  const sheet_se = fetchSheetByName('Settings');                                      // Settings sheet
  if (!sheet_co || !sheet_se) return;
  
  const sheet_sr = fetchSheetByName('Index');                                         // Source sheet
  if (!sheet_sr) { Logger.log('ERROR EXPORT:', SheetName, 'Does not exist on doExportData from sheet_sr'); return; }
   const trg = SpreadsheetApp.openById(Target_Id);                                    // Target spreadsheet

  const sheet_tr = trg.getSheetByName(SheetName);                                     // Target sheet - does not use fetchSheetByName, because gets data from diferent spreadsheet
  if (!sheet_tr) { Logger.log('ERROR EXPORT:', SheetName, 'Does not exist on doExportData from sheet_tr'); return; }
    var LR_T = sheet_tr.getLastRow();
    var LC_T = sheet_tr.getLastColumn();

  let Data = [];
  let Export;


  switch (SheetName) 
  {

//-------------------------------------------------------------------BLC-------------------------------------------------------------------//

    case BLC:

    Export = getConfigValue(EBL)                                                      // EBL = Export to BLC

    var A = sheet_co.getRange('B18').getValue();                                      // Balanço Atual

    var B = sheet_sr.getRange('B43').getValue();                                      // Ativo
    var C = sheet_sr.getRange('B44').getValue();                                      // A. Circulante
    var D = sheet_sr.getRange('B45').getValue();                                      // A. Não Circulante
    var E = sheet_sr.getRange('B46').getValue();                                      // Passivo
    var F = sheet_sr.getRange('B47').getValue();                                      // Passivo Circulante
    var G = sheet_sr.getRange('B48').getValue();                                      // Passivo Não Circ
    var H = sheet_sr.getRange('B49').getValue();                                      // Patrim. Líq

    Data.push([A, B, C, D, E, F, G, H]);

    break;

//-------------------------------------------------------------------DRE-------------------------------------------------------------------//

    case DRE:

    Export = getConfigValue(EDR)                                                     // EDR = Export to DRE

    var A = sheet_co.getRange('B18').getValue();                                     // Balanço Atual

    var B = sheet_sr.getRange('B52').getValue();                                     // Receita Líquida 12 MESES
    var C = sheet_sr.getRange('B53').getValue();                                     // Resultado Bruto 12 MESES
    var D = sheet_sr.getRange('B54').getValue();                                     // EBIT 12 MESES
    var E = sheet_sr.getRange('B55').getValue();                                     // EBITDA 12 MESES
    var F = sheet_sr.getRange('B57').getValue();                                     // Lucro Líquido 12 MESES

    var G = sheet_sr.getRange('D52').getValue();                                     // Receita Líquida 3 MESES
    var H = sheet_sr.getRange('D53').getValue();                                     // Resultado Bruto 3 MESES
    var I = sheet_sr.getRange('D54').getValue();                                     // EBIT 3 MESES
    var J = sheet_sr.getRange('D55').getValue();                                     // EBITDA 3 MESES
    var K = sheet_sr.getRange('D57').getValue();                                     // Lucro Líquido 3 MESES

    Data.push([A, B, C, D, E, F, G, H, I, J, K]);

    break;

//-------------------------------------------------------------------FLC-------------------------------------------------------------------//

    case FLC:

    Export = getConfigValue(EFL)                                                     // EFL = Export to FLC

    var A = sheet_co.getRange('B18').getValue();                                     // Balanço Atual

    var B = sheet_sr.getRange('B69').getValue();                                     // FCO
    var C = sheet_sr.getRange('B70').getValue();                                     // FCI
    var D = sheet_sr.getRange('B71').getValue();                                     // FCF
    var E = sheet_sr.getRange('B72').getValue();                                     // FCT
    var F = sheet_sr.getRange('B73').getValue();                                     // FCL
    var G = sheet_sr.getRange('B74').getValue();                                     // Saldo Inicial
    var H = sheet_sr.getRange('B75').getValue();                                     // Saldo Final

    Data.push([A, B, C, D, E, F, G, H]);

    break;

//-------------------------------------------------------------------DVA-------------------------------------------------------------------//

    case DVA:

    Export = getConfigValue(EDV)                                                     // EDV = Export to DVA

    var A = sheet_co.getRange('B18').getValue();                                     // Balanço Atual
      
    var B = sheet_sr.getRange('B77').getValue();                                     // Receitas
    var C = sheet_sr.getRange('B78').getValue();                                     // Insumos Adquiridos de Terceiros
    var D = sheet_sr.getRange('D77').getValue();                                     // Valor Adicionado Bruto
    var E = sheet_sr.getRange('B79').getValue();                                     // Depreciação, Amortização e Exaustão
    var F = sheet_sr.getRange('D78').getValue();                                     // Valor Adicionado Recebido em Transferência
    var G = sheet_sr.getRange('D79').getValue();                                     // Valor Adicionado Total a Distribuir

    Data.push([A, B, C, D, E, F, G]);

    break;

    default:
      Export = null;
    break;
  }

//-------------------------------------------------------------------Foot-------------------------------------------------------------------//

  if( Export == "TRUE" )
  {
    var Search = sheet_tr.getRange('A2:A' + LR_T).createTextFinder(TKT).findNext();

    if (Search)
    {
      Search.offset(0, 1, 1 , Data[0].length).setValues(Data);
      Logger.log(`SUCCESS EXPORT. Data for ${TKT}. Sheet: ${SheetName}.`);
    }
    else
    {
      var NewRow = sheet_tr.getRange(LR_T+1,1,1,1).setValue([TKT]);
      Logger.log(`SUCCESS EXPORT. Ticker: ${TKT}. Sheet: ${SheetName}.`);

      NewRow.offset(0, 1, 1 , Data[0].length).setValues(Data);
      Logger.log(`SUCCESS EXPORT. Data for ${TKT}. Sheet: ${SheetName}.`);
    }
  }
  else
  {
    Logger.log('ERROR EXPORT:', SheetName, 'EXPORT on config is set to FALSE');
  }
}

/////////////////////////////////////////////////////////////////////EXTRA TEMPLATE/////////////////////////////////////////////////////////////////////

function doExportExtra(SheetName)
{
  Logger.log('EXPORT:', SheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');                                       // Config sheet
    var Target_Id = sheet_co.getRange(TDR).getValues();                              // TDR = Target ID Range
  const sheet_se = fetchSheetByName('Settings');                                     // Settings sheet
  const sheet_sr = fetchSheetByName(SheetName);                                      // Source sheet
  if (!sheet_sr) { Logger.log('ERROR EXPORT:', SheetName, 'Does not exist on doExportExtra from sheet_sr'); return; }

  let ShouldExport = false; // Initialize ShouldExport as false
  let Data = [];
  let Export;                 // Declare Export without an initial value

  switch (SheetName) 
  {

//-------------------------------------------------------------------Right-------------------------------------------------------------------//

    case RIGHT_1:
    case RIGHT_2:

    Export = getConfigValue(ERT)                                                     // ERT = Export to Right

    break;

//-------------------------------------------------------------------Receipt-------------------------------------------------------------------//

    case RECEIPT_9:
    case RECEIPT_10:

    Export = getConfigValue(ERC)                                                     // ERC = Export to Receipt

    break;

//-------------------------------------------------------------------Warrant-------------------------------------------------------------------//

    case WARRANT_11:
    case WARRANT_12:
    case WARRANT_13:

    Export = getConfigValue(EWT)                                                     // EWT = Export to Warrant

    break;

//-------------------------------------------------------------------Block-------------------------------------------------------------------//

    case BLOCK:

    Export = getConfigValue(EBK)                                                     // EBK = Export to Block

    break;

    default:
      Export = null;
    break;
  }

  var M = sheet_sr.getRange('M2').getValue();                                        // Ticker

  var A = sheet_sr.getRange('A2').getValue();                                        // Data
  var B = sheet_sr.getRange('B2').getValue();                                        // Cotação
  var C = sheet_sr.getRange('C2').getValue();                                        // PM
  var D = sheet_sr.getRange('D2').getValue();                                        // Contratos
  var E = sheet_sr.getRange('E2').getValue();                                        // Mínimo
  var F = sheet_sr.getRange('F2').getValue();                                        // Máximo
  var G = sheet_sr.getRange('G2').getValue();                                        // Volume
  var H = sheet_sr.getRange('H2').getValue();                                        // Negócios
  var I = sheet_sr.getRange('I2').getValue();                                        // Ratio

  var N = sheet_sr.getRange('N2').getValue();                                        // Início
  var O = sheet_sr.getRange('O2').getValue();                                        // Fim

  var J = sheet_sr.getRange('J2').getValue();                                        // Emissão
  var K = sheet_sr.getRange('K2').getValue();                                        // Preço
  var L = sheet_sr.getRange('L2').getValue();                                        // Diff

  var Range = [B, C, D, E, F, G, H, I];

  var hasNonBlankCell = Range.some(cell => cell !== '' && cell !== null);            // Check if at least one cell is not blank

  if (hasNonBlankCell && !ErrorValues.some(error => Range.includes(error))) 
  {
    Data.push([A, B, C, D, E, F, G, H, I, N, O, J, K, L]);
    ShouldExport = true;                                                             // Set ShouldExport to true if conditions are met
  }

//-------------------------------------------------------------------Foot-------------------------------------------------------------------//

  if( !ErrorValues.includes(A) )
  {
    if (ShouldExport) 
    {
      if (Export == "TRUE") 
      {
        const trg = SpreadsheetApp.openById(Target_Id);                             // Target spreadsheet
        let sheet_tr;                                                               // Declare sheet_tr outside the conditional scope

        if (SheetName === RIGHT_1 || SheetName === RIGHT_2 ) 
        {
          sheet_tr = trg.getSheetByName('Right');
        }
        else if (SheetName === RECEIPT_9 || SheetName === RECEIPT_10 ) 
        {
          sheet_tr = trg.getSheetByName('Receipt');
        }
        else if (SheetName === WARRANT_11 || SheetName === WARRANT_12 || SheetName === WARRANT_13) 
        {
          sheet_tr = trg.getSheetByName('Warrant');
        }
        else if (SheetName === BLOCK) 
        {
          sheet_tr = trg.getSheetByName('Block');
        }
        else
        {
          sheet_tr = trg.getSheetByName(SheetName);
        }

        var LR_T = sheet_tr.getLastRow();
        var LC_T = sheet_tr.getLastColumn();

        var Search = sheet_tr.getRange("A2:A" + LR_T).createTextFinder(M).findNext();

        if (Search)
        {
          Search.offset(0, 1, 1 , Data[0].length).setValues(Data);
          Logger.log(`SUCCESS EXPORT. Data for ${M} . Sheet: ${SheetName}.`);
        }
        else
        {
          // Value not found, add a new row with the ticker (M)
          var NewRow = sheet_tr.getRange(LR_T + 1, 1, 1, 1).setValue([M]);
          Logger.log(`SUCCESS EXPORT. Ticker: ${M}. Sheet: ${SheetName}.`);

          // Now set the adjacent values (Data) in the new row
          NewRow.offset(0, 1, 1, Data[0].length).setValues(Data);
          Logger.log(`SUCCESS EXPORT. Data for ${M}. Sheet: ${SheetName}.`);
        }
      }
      else
      {
        Logger.log('EXPORT:', SheetName, 'Export on config is set to FALSE');
      }
    }
    else
    {
      Logger.log('EXPORT:', SheetName, 'ShouldExport is FALSE');
    }
  }
  else
  {
    Logger.log('EXPORT:', SheetName, 'Data (A) failed ErrorValues');
  }
}

/////////////////////////////////////////////////////////////////////INFO/////////////////////////////////////////////////////////////////////

function doExportInfo()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');                                      // Config sheet

  var Exported = sheet_co.getRange(EXR).getDisplayValue();                          // EXR = Exported?
  var Target_Id = sheet_co.getRange(DIR).getValues();                               // DIR = DATA Source ID

  const sheet_in = fetchSheetByName('Info');                                        // Info sheet
  var SheetName =  sheet_in.getName()
  Logger.log('Export:', SheetName);
  if (!sheet_co || !sheet_in) return;

  var A = sheet_co.getRange('B3').getValue();                                       // Ticket
  var B = sheet_in.getRange('C3').getValue();                                       // Codigo CVM
  var C = sheet_in.getRange('C4').getValue();                                       // CNPJ
  var D = sheet_in.getRange('C5').getValue();                                       // Empresa
  var E = sheet_in.getRange('C6').getValue();                                       // Razão Social
  var F = sheet_in.getRange('C13').getValue();                                      // Tipo de Ação
  var G = sheet_in.getRange('C9').getValue();                                       // Listagem
  var H = sheet_in.getRange('C18').getValue();                                      // Setor
  var I = sheet_in.getRange('C19').getValue();                                      // Subsetor
  var J = sheet_in.getRange('C20').getValue();                                      // Segmento
  var K = sheet_in.getRange('C7').getValue();                                       // Situação Registro

  var Data = [];
  Data.push(A,B,C,D,E,F,G,H,I,J,K);

  var ss_t = SpreadsheetApp.openById(Target_Id);                                    // Target spreadsheet
  var sheet_tr = ss_t.getSheetByName('Relação');                                    // Target sheet

  var LR = sheet_tr.getLastRow();
  var LC = sheet_tr.getLastColumn();

  var columnNumber = Data.length;

  if( Exported !== "TRUE")
  {
    var export_r = sheet_tr.getRange(LR+1,1,1,columnNumber).setValues([Data]);

    setSheetID()

    Logger.log(`SUCCESS EXPORT. Sheet: ${SheetName}.`);
  }
};

/////////////////////////////////////////////////////////////////////EXPORT TEMPLATE/////////////////////////////////////////////////////////////////////