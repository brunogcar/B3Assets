/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function doExportAll()
{
  doExportSheets();
  doExportExtras();
  doExportDatas();
};

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

function doExportSheets() 
{
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND];
  
  SheetNames.forEach(SheetName => 
  {
    try 
    {
      doExportSheet(SheetName);
    } 
    catch (error) 
    {
      // Handle the error here, you can log it or take appropriate actions.
      console.error(`Error exporting sheet ${SheetName}:`, error);
    }
  });
}

function doExportDatas() 
{
  const SheetNames = [BLC, DRE, FLC, DVA];

  SheetNames.forEach(SheetName => 
  {
    try 
    {
      doExportData(SheetName);
    } 
    catch (error) 
    {
      // Handle the error here, you can log it or take appropriate actions.
      console.error(`Error exporting sheet ${SheetName}:`, error);
    }
  });
}

function doExportExtras() 
{
  const SheetNames = [FUTURE, FUTURE_1, FUTURE_2, FUTURE_3, RIGHT_1, RIGHT_2, RECEIPT_9, RECEIPT_10, WARRANT_11, WARRANT_12, WARRANT_13, BLOCK];

  SheetNames.forEach(SheetName => 
  {
    try 
    {
      doExportExtra(SheetName);
    } 
    catch (error) 
    {
      // Handle the error here, you can log it or take appropriate actions.
      console.error(`Error exporting sheet ${SheetName}:`, error);
    }
  });
}

/////////////////////////////////////////////////////////////////////SHEETS TEMPLATE/////////////////////////////////////////////////////////////////////

function doExportSheet(SheetName) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = ss.getSheetByName('Config');
    var Class = sheet_co.getRange(IST).getDisplayValue();                             // IST = Is Stock? 
    var TKT = sheet_co.getRange(TKR).getValue();                                      // TKR = Ticket Range
    var Target_Id = sheet_co.getRange(TDR).getValues();
  const sheet_se = ss.getSheetByName('Settings');
    var Minimum = sheet_se.getRange(MIN).getValue();                                  // -1000 - Default
    var Maximum = sheet_se.getRange(MAX).getValue();                                  //  1000 - Default
  const sheet_sr = ss.getSheetByName(SheetName);                                      // Source sheet
    var A2 = sheet_sr.getRange('A2').getValue();
    var A5 = sheet_sr.getRange('A5').getValue();
    var LR_S = sheet_sr.getLastRow();
    var LC_S = sheet_sr.getLastColumn();
  const trg = SpreadsheetApp.openById(Target_Id);                                     // Target spreadsheet
  const sheet_tr = trg.getSheetByName(SheetName);                                     // Target sheet

  if( sheet_tr ) 
  {
    var LR_T = sheet_tr.getLastRow();
    var LC_T = sheet_tr.getLastColumn();

    console.log('EXPORT:', SheetName);

    let ShouldExport = false;
    let Export;                 // Declare Export without an initial value
    let Value_se = "DEFAULT";   // Initialize Value_se with "DEFAULT" as the fallback

    if( sheet_sr ) 
    {
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

            Value_se = sheet_se.getRange(ETR).getDisplayValue().trim();                 // ETR = Export to Swing

            Export = (Value_se === "DEFAULT") 
              ? sheet_co.getRange(ETR).getDisplayValue().trim()                         // Use Config value if Settings has "DEFAULT"
              : Value_se;                                                               // Use Settings value

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

            Value_se = sheet_se.getRange(EOP).getDisplayValue().trim();                 // EOP = Export to Option

            Export = (Value_se === "DEFAULT") 
              ? sheet_co.getRange(EOP).getDisplayValue().trim()
              : Value_se;    

            var [Call, Put] = ['C2', 'E2'].map(r => sheet_sr.getRange(r).getValue());

            if( ( Call != 0 && Put != 0 ) &&
                ( Call != "" && Put != "" ) )
            {
              ShouldExport = true;
            }
            break;

//-------------------------------------------------------------------BTC-------------------------------------------------------------------//

            case BTC:

            Value_se = sheet_se.getRange(EBT).getDisplayValue().trim();                 // EBT = Export to BTC

            Export = (Value_se === "DEFAULT") 
              ? sheet_co.getRange(EBT).getDisplayValue().trim()
              : Value_se;   

            var D2 = sheet_sr.getRange('D2').getValue();

            if( !ErrorValues.includes(D2) )
            {
              ShouldExport = true;
            }
            break;

//-------------------------------------------------------------------Termo-------------------------------------------------------------------//

            case TERMO:

            Value_se = sheet_se.getRange(ETE).getDisplayValue().trim();                 // ETE = Export to Termo

            Export = (Value_se === "DEFAULT") 
              ? sheet_co.getRange(ETE).getDisplayValue().trim()
              : Value_se;   

            var D2 = sheet_sr.getRange('D2').getValue();

            if( !ErrorValues.includes(D2) )
            {
              ShouldExport = true;
            }
            break;

//-------------------------------------------------------------------Future-------------------------------------------------------------------//

            case FUTURE:

            Value_se = sheet_se.getRange(ETF).getDisplayValue().trim();                 // ETF = Export to Future

            Export = (Value_se === "DEFAULT") 
              ? sheet_co.getRange(ETF).getDisplayValue().trim()
              : Value_se;   

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

            Value_se = sheet_se.getRange(ETF).getDisplayValue().trim();                 // ETF = Export to Future

            Export = (Value_se === "DEFAULT") 
              ? sheet_co.getRange(ETF).getDisplayValue().trim()
              : Value_se;   

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

            Value_se = sheet_se.getRange(EFU).getDisplayValue().trim();                 // EFU = Export to Fund 

            Export = (Value_se === "DEFAULT") ? sheet_co.getRange(EFU).getDisplayValue().trim() : Value_se;

            var B2 = sheet_sr.getRange('B2').getValue();

            if( !ErrorValues.includes(B2) )
            {
              ShouldExport = true;
            }
            break;

            default:
              console.log('ERROR EXPORT:', SheetName, 'Sheet name not recognized.');
              return;
          }

//-------------------------------------------------------------------Foot-------------------------------------------------------------------//

          if (Export == 'TRUE' && ShouldExport) 
          {
            let FilteredData;

            if (SheetName === FUND) 
            {
              var Data = sheet_sr.getRange(2, 1, 1, LC_S).getValues();                            // 2D array
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
                  return (Value > Minimum && Value < Maximum) ? Value : "";                      // Apply filtering for columns <= BJ (except column B)
                }
              });
            } 
            else 
            {
              FilteredData = sheet_sr.getRange(2, 1, 1, LC_S).getValues()[0];                    // Use unfiltered data
            }


            var Search = sheet_tr.getRange('A2:A' + LR_T).createTextFinder(TKT).findNext();

            if (Search) 
            {
              Search.offset(0, 1, 1, FilteredData.length).setValues([FilteredData]);            // Ensure it's a 2D array
              console.log(`SUCCESS EXPORT. Data for ${TKT}. Sheet: ${SheetName}.`);
            } 
            else 
            {
              var NewRow = sheet_tr.getRange(LR_T + 1, 1, 1, 1).setValue([TKT]);
              console.log(`SUCCESS EXPORT. Ticker: ${TKT}. Sheet: ${SheetName}.`);

              NewRow.offset(0, 1, 1, FilteredData.length).setValues([FilteredData]);            // Ensure it's a 2D array
              console.log(`SUCCESS EXPORT. Data for ${TKT}. Sheet: ${SheetName}.`);
            }
          }
          else 
          {
            console.log('ERROR EXPORT:', SheetName, 'EXPORT on config is set to FALSE or Conditions arent met on doExportSheet');
          }
        }
        else 
        {
          console.log('ERROR EXPORT:', SheetName, 'ErrorValues in A2 or A5_ = "" on doExportSheet');
        }
      }
      else 
      {
        console.log('ERROR EXPORT:', SheetName, 'Class != STOCK', Class);
      }
    }
    else 
    {
      console.log('ERROR EXPORT:', SheetName, 'Does not exist on doExportSheet from sheet_sr');
    }
  }
  else 
  {
    console.log('ERROR EXPORT:', SheetName, 'Does not exist on doExportSheet from sheet_tr');
  }
}

/////////////////////////////////////////////////////////////////////DATA TEMPLATE/////////////////////////////////////////////////////////////////////

function doExportData(SheetName) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = ss.getSheetByName('Config');
    var TKT = sheet_co.getRange(TKR).getValue();               // TKR = Ticket Range
    var Target_Id = sheet_co.getRange(TDR).getValues();        // TDR = Target ID Range
  const sheet_se = ss.getSheetByName('Settings');
  const sheet_sr = ss.getSheetByName('Index');               // Source sheet
  const sheet_tr = ss.getSheetByName(SheetName);             // Targer sheet of check

  console.log('EXPORT:', SheetName);

  let Data = [];
  let Export;                 // Declare Export without an initial value
  let Value_se = "DEFAULT";   // Initialize Value_se with "DEFAULT" as the fallback

  if (sheet_tr)
  {
    switch (SheetName) 
    {

//-------------------------------------------------------------------BLC-------------------------------------------------------------------//

      case BLC:

      Value_se = sheet_se.getRange(EBL).getDisplayValue().trim();    // EBL = Export to BLC 

      Export = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(EBL).getDisplayValue().trim()
        : Value_se;   

      var A = sheet_co.getRange('B18').getValue();                  // Balanço Atual

      var B = sheet_sr.getRange('B43').getValue();                  // Ativo
      var C = sheet_sr.getRange('B44').getValue();                  // A. Circulante
      var D = sheet_sr.getRange('B45').getValue();                  // A. Não Circulante
      var E = sheet_sr.getRange('B46').getValue();                  // Passivo
      var F = sheet_sr.getRange('B47').getValue();                  // Passivo Circulante
      var G = sheet_sr.getRange('B48').getValue();                  // Passivo Não Circ
      var H = sheet_sr.getRange('B49').getValue();                  // Patrim. Líq

      Data.push([A, B, C, D, E, F, G, H]);

      break;

//-------------------------------------------------------------------DRE-------------------------------------------------------------------//

      case DRE:

      Value_se = sheet_se.getRange(EDR).getDisplayValue().trim();    // EDR = Export to DRE 

      Export = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(EDR).getDisplayValue().trim()
        : Value_se;  

      var A = sheet_co.getRange('B18').getValue();                  // Balanço Atual

      var B = sheet_sr.getRange('B52').getValue();                  // Receita Líquida 12 MESES
      var C = sheet_sr.getRange('B53').getValue();                  // Resultado Bruto 12 MESES
      var D = sheet_sr.getRange('B54').getValue();                  // EBIT 12 MESES
      var E = sheet_sr.getRange('B55').getValue();                  // EBITDA 12 MESES
      var F = sheet_sr.getRange('B57').getValue();                  // Lucro Líquido 12 MESES

      var G = sheet_sr.getRange('D52').getValue();                  // Receita Líquida 3 MESES
      var H = sheet_sr.getRange('D53').getValue();                  // Resultado Bruto 3 MESES
      var I = sheet_sr.getRange('D54').getValue();                  // EBIT 3 MESES
      var J = sheet_sr.getRange('D55').getValue();                  // EBITDA 3 MESES
      var K = sheet_sr.getRange('D57').getValue();                  // Lucro Líquido 3 MESES

      Data.push([A, B, C, D, E, F, G, H, I, J, K]);

      break;

//-------------------------------------------------------------------FLC-------------------------------------------------------------------//

      case FLC:

      Value_se = sheet_se.getRange(EFL).getDisplayValue().trim();   // EFL = Export to FLC 

      Export = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(EFL).getDisplayValue().trim()
        : Value_se;  

      var A = sheet_co.getRange('B18').getValue();                  // Balanço Atual

      var B = sheet_sr.getRange('B69').getValue();                  // FCO
      var C = sheet_sr.getRange('B70').getValue();                  // FCI
      var D = sheet_sr.getRange('B71').getValue();                  // FCF
      var E = sheet_sr.getRange('B72').getValue();                  // FCT
      var F = sheet_sr.getRange('B73').getValue();                  // FCL
      var G = sheet_sr.getRange('B74').getValue();                  // Saldo Inicial
      var H = sheet_sr.getRange('B75').getValue();                  // Saldo Final

      Data.push([A, B, C, D, E, F, G, H]);

      break;

//-------------------------------------------------------------------DVA-------------------------------------------------------------------//

      case DVA:

      Value_se = sheet_se.getRange(EDV).getDisplayValue().trim();   // EDV = Export to DVA 

      Export = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(EDV).getDisplayValue().trim()
        : Value_se;  

      var A = sheet_co.getRange('B18').getValue();                  // Balanço Atual
      
      var B = sheet_sr.getRange('B77').getValue();                  // Receitas
      var C = sheet_sr.getRange('B78').getValue();                  // Insumos Adquiridos de Terceiros
      var D = sheet_sr.getRange('D77').getValue();                  // Valor Adicionado Bruto
      var E = sheet_sr.getRange('B79').getValue();                  // Depreciação, Amortização e Exaustão
      var F = sheet_sr.getRange('D78').getValue();                  // Valor Adicionado Recebido em Transferência
      var G = sheet_sr.getRange('D79').getValue();                  // Valor Adicionado Total a Distribuir

      Data.push([A, B, C, D, E, F, G]);

      break;

      default:
        Export = null;
      break;
    }

//-------------------------------------------------------------------Foot-------------------------------------------------------------------//

    if( Export == "TRUE" )
    {
      const trg = SpreadsheetApp.openById(Target_Id);        // Target spreadsheet
      const sheet_tr = trg.getSheetByName(SheetName);         // Target sheet

      var LR_T = sheet_tr.getLastRow();
      var LC_T = sheet_tr.getLastColumn();

      var Search = sheet_tr.getRange('A2:A' + LR_T).createTextFinder(TKT).findNext();

      if (Search)
      {
        Search.offset(0, 1, 1 , Data[0].length).setValues(Data);
        console.log(`SUCESS EXPORT. Data for ${TKT}. Sheet: ${SheetName}.`);
      }
      else
      {
        var NewRow = sheet_tr.getRange(LR_T+1,1,1,1).setValue([TKT]);
        console.log(`SUCESS EXPORT. Ticker: ${TKT}. Sheet: ${SheetName}.`);

        NewRow.offset(0, 1, 1 , Data[0].length).setValues(Data);
        console.log(`SUCCESS EXPORT. Data for ${TKT}. Sheet: ${SheetName}.`);
      }
    }
    else
    {
      console.log('ERROR EXPORT:', SheetName, 'EXPORT on config is set to FALSE');
    }
  }
  else
  {
    console.log('ERROR EXPORT:', SheetName, 'Does not exist on doExportData');
  }
}

/////////////////////////////////////////////////////////////////////EXTRA TEMPLATE/////////////////////////////////////////////////////////////////////

function doExportExtra(SheetName) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = ss.getSheetByName('Config');

  var Target_Id = sheet_co.getRange(TDR).getValues();                         // TDR = Target ID Range

  const sheet_se = ss.getSheetByName('Settings');
  const sheet_sr = ss.getSheetByName(SheetName);                              // Source sheet

  console.log('EXPORT:', SheetName);

  let ShouldExport = false; // Initialize ShouldExport as false
  let Data = [];
  let Export;                 // Declare Export without an initial value
  let Value_se = "DEFAULT";   // Initialize Value_se with "DEFAULT" as the fallback

  if (sheet_sr)
  {
    switch (SheetName) 
    {
      case RIGHT_1:
      case RIGHT_2:

      Value_se = sheet_se.getRange(ERT).getDisplayValue().trim();            // ERT = Export to Right 

      Export = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(ERT).getDisplayValue().trim()
        : Value_se;  

      break;

      case RECEIPT_9:
      case RECEIPT_10:

      Value_se = sheet_se.getRange(ERC).getDisplayValue().trim();            // ERC = Export to Receipt 

      Export = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(ERC).getDisplayValue().trim()
        : Value_se;  

      break;

      case WARRANT_11:
      case WARRANT_12:
      case WARRANT_13:

      Value_se = sheet_se.getRange(EWT).getDisplayValue().trim();            // EWT = Export to Warrant 

      Export = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(EWT).getDisplayValue().trim()
        : Value_se;  

      break;

      case BLOCK:

      Value_se = sheet_se.getRange(EBK).getDisplayValue().trim();            // EBK = Export to Block

      Export = (Value_se === "DEFAULT") 
        ? sheet_co.getRange(EBK).getDisplayValue().trim()
        : Value_se;  

      break;

      default:
        Export = null;
      break;
    }

    var M = sheet_sr.getRange('M2').getValue();                             // Ticker

    var A = sheet_sr.getRange('A2').getValue();                             // Data
    var B = sheet_sr.getRange('B2').getValue();                             // Cotação
    var C = sheet_sr.getRange('C2').getValue();                             // PM
    var D = sheet_sr.getRange('D2').getValue();                             // Contratos
    var E = sheet_sr.getRange('E2').getValue();                             // Mínimo
    var F = sheet_sr.getRange('F2').getValue();                             // Máximo
    var G = sheet_sr.getRange('G2').getValue();                             // Volume
    var H = sheet_sr.getRange('H2').getValue();                             // Negócios
    var I = sheet_sr.getRange('I2').getValue();                             // Ratio

    var N = sheet_sr.getRange('N2').getValue();                             // Início
    var O = sheet_sr.getRange('O2').getValue();                             // Fim

    var J = sheet_sr.getRange('J2').getValue();                             // Emissão
    var K = sheet_sr.getRange('K2').getValue();                             // Preço
    var L = sheet_sr.getRange('L2').getValue();                             // Diff

    var Range = [B, C, D, E, F, G, H, I];

    var hasNonBlankCell = Range.some(cell => cell !== '' && cell !== null); // Check if at least one cell is not blank

    if (hasNonBlankCell && !ErrorValues.some(error => Range.includes(error))) 
    {
      Data.push([A, B, C, D, E, F, G, H, I, N, O, J, K, L]);
      ShouldExport = true;                                                  // Set ShouldExport to true if conditions are met
    }

//-------------------------------------------------------------------Foot-------------------------------------------------------------------//

    if( !ErrorValues.includes(A) )
    {
      if (ShouldExport) 
      {
        if (Export == "TRUE") 
        {
          const trg = SpreadsheetApp.openById(Target_Id);   // Target spreadsheet
          let sheet_tr;                                      // Declare sheet_tr outside the conditional scope

          if (SheetName === RIGHT_1 || SheetName === RIGHT_2 ) 
          {
            sheet_tr = trg.getSheetByName('Right');          // Target sheet
          }
          else if (SheetName === RECEIPT_9 || SheetName === RECEIPT_10 ) 
          {
            sheet_tr = trg.getSheetByName('Receipt');        // Target sheet
          }
          else if (SheetName === WARRANT_11 || SheetName === WARRANT_12 || SheetName === WARRANT_13) 
          {
            sheet_tr = trg.getSheetByName('Warrant');        // Target sheet
          }
          else if (SheetName === BLOCK) 
          {
            sheet_tr = trg.getSheetByName('Block');        // Target sheet
          }
          else
          {
            sheet_tr = trg.getSheetByName(SheetName);        // Target sheet
          }

          var LR_T = sheet_tr.getLastRow();
          var LC_T = sheet_tr.getLastColumn();

          var Search = sheet_tr.getRange("A2:A" + LR_T).createTextFinder(M).findNext();

          if (Search)
          {
            Search.offset(0, 1, 1 , Data[0].length).setValues(Data);
            console.log(`SUCESS EXPORT. Data for ${M} . Sheet: ${SheetName}.`);
          }
          else
          {
            // Value not found, add a new row with the ticker (M)
            var NewRow = sheet_tr.getRange(LR_T + 1, 1, 1, 1).setValue([M]);
            console.log(`SUCCESS EXPORT. Ticker: ${M}. Sheet: ${SheetName}.`);

            // Now set the adjacent values (Data) in the new row
            NewRow.offset(0, 1, 1, Data[0].length).setValues(Data);
            console.log(`SUCCESS EXPORT. Data for ${M}. Sheet: ${SheetName}.`);
          }
        }
        else
        {
          console.log('EXPORT:', SheetName, 'Export on config is set to FALSE');
        }
      }
      else
      {
        console.log('EXPORT:', SheetName, 'ShouldExport is FALSE');
      }
    }
    else
    {
      console.log('EXPORT:', SheetName, 'Data (A) failed ErrorValues');
    }
  }
  else
  {
    console.log('ERROR EXPORT:', SheetName, 'Does not exist on doExportExtra');
  }
}

/////////////////////////////////////////////////////////////////////INFO/////////////////////////////////////////////////////////////////////

function doExportInfo()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet_co = ss.getSheetByName('Config');

  var Exported = sheet_co.getRange(EXR).getDisplayValue();    // EXR = Exported?
  var Target_Id = sheet_co.getRange(DIR).getValues();         // DIR = DATA Source ID

  var sheet_in = ss.getSheetByName('Info');                   // Target sheet
  var SheetName =  sheet_in.getName()

  var A = sheet_co.getRange('B3').getValue();                 // Ticket
  var B = sheet_in.getRange('C3').getValue();                 // Codigo CVM
  var C = sheet_in.getRange('C4').getValue();                 // CNPJ
  var D = sheet_in.getRange('C5').getValue();                 // Empresa
  var E = sheet_in.getRange('C6').getValue();                 // Razão Social
  var F = sheet_in.getRange('C13').getValue();                // Tipo de Ação
  var G = sheet_in.getRange('C9').getValue();                 // Listagem
  var H = sheet_in.getRange('C18').getValue();                // Setor
  var I = sheet_in.getRange('C19').getValue();                // Subsetor
  var J = sheet_in.getRange('C20').getValue();                // Segmento
  var K = sheet_in.getRange('C7').getValue();                 // Situação Registro

  console.log('Export:', SheetName);

  var Data = [];
  Data.push(A,B,C,D,E,F,G,H,I,J,K);

  var ss_t = SpreadsheetApp.openById(Target_Id);              // Target spreadsheet
  var sheet_tr = ss_t.getSheetByName('Relação');                // Target sheet

  var LR = sheet_tr.getLastRow();
  var LC = sheet_tr.getLastColumn();

  var columnNumber = Data.length;

  if( Exported !== "TRUE")
  {
    var export_r = sheet_tr.getRange(LR+1,1,1,columnNumber).setValues([Data]);

    setSheetID()

    console.log(`SUCESS EXPORT. Sheet: ${SheetName}.`);
  }
};

/////////////////////////////////////////////////////////////////////EXPORT/////////////////////////////////////////////////////////////////////