//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////IMPORT/////////////////////////////////////////////////////////////////////

function Import(){
  const sheet_co = fetchSheetByName('Config');                                   // Config sheet
  if (!sheet_co) {Logger.log("ERROR: 'Config' sheet not found."); return;}

  // Check if L2 has the expected colors
  if (!checkAutorizeScript()) {
    Logger.log("Import aborted: L2 does not have the correct background and font colors.");
    return;
  }

  const Source_Id = getConfigValue(SIR, 'Config');                               // SIR = Source ID
  if (!Source_Id) {Logger.log("Warning: Source ID is empty."); return;}
  const Option = sheet_co.getRange(OPR).getDisplayValue();                       // OPR = Option

  if (Option === "AUTO") 
  {
    // Check for specific sheets
    const hasSwing4 = fetchSheetByName(SWING_4) !== null;
    const hasSwing12 = fetchSheetByName(SWING_12) !== null;
    const hasSwing52 = fetchSheetByName(SWING_52) !== null;
    const hasTrade = fetchSheetByName('Trade') !== null;

    if (hasSwing4 && hasSwing12 && hasSwing52) 
    {
      import_Current();
    } 
    else if (hasSwing12 && hasSwing52) 
    {
      import_15x_to_161();
    } 
    else if (hasTrade) 
    {
      import_14x_to_161();
    } 
    else 
    {
      Logger.log(`No matching sheets found for AUTO mode.`);
    }
  } 
  else 
  {
    // Manual Option Handling
    if (Option == 1) 
    {
      import_Current();
    } 
    else if (Option == 2) 
    {
      import_15x_to_163();
    }
    else 
    {
      Logger.log(`Invalid Option: ${Option}`);
    }
  }
}

/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function import_Current(){
  console.log('Import: import_Current');

  import_config();

  doImportProventos();
  doImportShares();

  doImportBasics();
  doImportFinancials();

  doCheckTriggers();
  update_form();

// doCleanZeros();

  console.log('Import: Finished');
};

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////

function doImportProventos() 
{
  const ProvNames = ['Proventos'];

  ProvNames.forEach(ProvName =>
  {
    try
    {
      doImportProv(ProvName);
    }
    catch (error)
    {
      // Handle the error here, you can log it or take appropriate actions.
      Logger.error(`Error importing ${ProvName}:`, error);
    }
  });
}

//-------------------------------------------------------------------BASICS-------------------------------------------------------------------//

function doImportBasics(){
  const SheetNames = [
    SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUND, 
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3, 
    RIGHT_1, RIGHT_2, 
    RECEIPT_9, RECEIPT_10, 
    WARRANT_11, WARRANT_12, WARRANT_13, 
    BLOCK
  ];

  const totalSheets = SheetNames.length;
  let Count = 0;

  const sheet_co = fetchSheetByName('Config');                                  // Config sheet
  const DEBUG = sheet_co.getRange(DBG).getDisplayValue();                       // DBG = Debug Mode

  if (DEBUG = "TRUE") Logger.log(`Starting import of ${totalSheets} sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100); // Calculate percentage
    if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) Importing ${SheetName}...`);

    try {
      doImportBasic(SheetName);
      if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} imported successfully`);
    } catch (error) {
      if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error importing ${SheetName}: ${error}`);
    }
  });
  if (DEBUG = "TRUE") Logger.log(`Import completed: ${Count} of ${totalSheets} sheets imported successfully`);
}

//-------------------------------------------------------------------DATA-------------------------------------------------------------------//

function doImportFinancials(){
  const SheetNames = [BLC, Balanco, DRE, Resultado, FLC, Fluxo, DVA, Valor];
  const totalSheets = SheetNames.length;
  let Count = 0;

  const sheet_co = fetchSheetByName('Config');                                  // Config sheet
  const DEBUG = sheet_co.getRange(DBG).getDisplayValue();                       // DBG = Debug Mode

  if (DEBUG = "TRUE") Logger.log(`Starting import of ${totalSheets} data sheets...`);

  SheetNames.forEach((SheetName, index) => {
    Count++;
    const progress = Math.round((Count / totalSheets) * 100); // Calculate percentage
    if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) Importing ${SheetName}...`);

    try {
      doImportFinancial(SheetName);
      if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) ${SheetName} imported successfully`);
    } catch (error) {
      if (DEBUG = "TRUE") Logger.log(`[${Count}/${totalSheets}] (${progress}%) Error importing ${SheetName}: ${error}`);
    }
  });
  if (DEBUG = "TRUE") Logger.log(`Import completed: ${Count} of ${totalSheets} data sheets imported successfully`);
}

/////////////////////////////////////////////////////////////////////Update Form/////////////////////////////////////////////////////////////////////

function update_form(){
  const sheet_co = fetchSheetByName('Config');                                        // Config sheet
  var Update_Form = ss.getRange(UFR).getDisplayValue();                               // UFR = Update Form

  switch (Update_Form) 
  {
    case 'EDIT':
      doEditAll();
      break;
    case 'SAVE':
      doSaveAll();
      break;

    default:
      Logger.error(`Invalid update form value: ${Update_Form}`);
      break;
  }
}

/////////////////////////////////////////////////////////////////////Functions/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////Config/////////////////////////////////////////////////////////////////////

function import_config(){
  const sheet_co = fetchSheetByName('Config');
  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('Config');       // Source Sheet
  {
    var Data = sheet_sr.getRange(COR).getValues();                                    // Does not use getConfigValue because it gets data from another spreadsheet
    sheet_co.getRange(COR).setValues(Data);
  }
};

/////////////////////////////////////////////////////////////////////SHARES and FF/////////////////////////////////////////////////////////////////////

function doImportShares() 
{
  const sheet_co = fetchSheetByName('Config');
  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('DATA');         // Source Sheet
    var L1 = sheet_sr.getRange("L1").getValue();
    var L2 = sheet_sr.getRange("L2").getValue();
  const sheet_tr = fetchSheetByName('DATA');                                          // Target Sheet
    var SheetName = sheet_tr.getName()

  Logger.log(`IMPORT: Shares and FF`);

  if (!ErrorValues.includes(L1) && !ErrorValues.includes(L2)) 
  {
    var Data = sheet_sr.getRange("L1:L2").getValues();
    sheet_tr.getRange("L1:L2").setValues(Data);
  }
  else
  {
    Logger.log(`ERROR IMPORT: ${SheetName} - ErrorValues on L1 or L2 on doImportShares`);
  }
Logger.log(`SUCCESS IMPORT: Shares and FF`);
}

/////////////////////////////////////////////////////////////////////Proventos/////////////////////////////////////////////////////////////////////

function doImportProv(ProvName){
  const sheet_co = fetchSheetByName('Config');
  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('Prov');         // Source Sheet
  const sheet_tr = fetchSheetByName('Prov');  

  Logger.log(`IMPORT: ${ProvName}`);

  if (ProvName == 'Proventos') 
  {
    var Check = sheet_sr.getRange("B3").getDisplayValue();  

    if( Check == "Proventos" )  // check if error
    {
      var Data = sheet_sr.getRange(PRV).getValues();                              // PRV = Provento Range
      sheet_tr.getRange(PRV).setValues(Data);
    }
    else
    {
      Logger.log(`ERROR IMPORT: ${ProvName} - B3 != Proventos on doImportProv`);
    }
  }
}

/////////////////////////////////////////////////////////////////////BASIC/////////////////////////////////////////////////////////////////////

function doImportBasic(SheetName){
  Logger.log(`IMPORT: ${SheetName}`);
  const sheet_co = fetchSheetByName('Config');                                        // Config sheet
  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
    if (!Source_Id) { Logger.log(`ERROR: Source ID not found in Config sheet`); return;}
  const sheet_se = fetchSheetByName('Settings');                                      // Settings sheet
  if (!sheet_co || !sheet_se) return;
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName(SheetName);      // Source Sheet
  if (!sheet_sr) { Logger.log(`ERROR IMPORT: Source sheet ${SheetName} - Does not exist on doImportBasic from ${Source_Id}`); return; }
  const sheet_tr = fetchSheetByName(SheetName);                                       // Target Sheet
  if (!sheet_tr) { Logger.log(`ERROR IMPORT: Target sheet ${SheetName} - does not exist on doImportBasic.`); return; }

  let Import;

  switch (SheetName) 
  {

//-------------------------------------------------------------------Swing-------------------------------------------------------------------//
    case SWING_4:
    case SWING_12:
    case SWING_52:

    Import = getConfigValue(ITR)                                                     // ITR = Import to Swing
    break;
//-------------------------------------------------------------------Opções-------------------------------------------------------------------//
    case OPCOES:

    Import = getConfigValue(IOP)                                                     // IOP = Import to Option
    break;
//-------------------------------------------------------------------BTC-------------------------------------------------------------------//
    case BTC:

    Import = getConfigValue(IBT)                                                     // IBT = Import to BTC
    break;
//-------------------------------------------------------------------Termo-------------------------------------------------------------------//
    case TERMO:

    Import = getConfigValue(ITE)                                                     // ITE = Import to Termo
    break;
//-------------------------------------------------------------------Future-------------------------------------------------------------------//
    case FUTURE:
    case FUTURE_1:
    case FUTURE_2:
    case FUTURE_3:

    Import = getConfigValue(IFT)                                                     // IFT = Import to Future
    break;
//-------------------------------------------------------------------Fund-------------------------------------------------------------------//
    case FUND:

    Import = getConfigValue(IFU)                                                     // IFU = Import to Fund
    break;
//-------------------------------------------------------------------Right-------------------------------------------------------------------//
    case RIGHT_1:
    case RIGHT_2:

    Import = getConfigValue(IRT)                                                     // IRT = Import to Right
    break;
//-------------------------------------------------------------------Receipt-------------------------------------------------------------------//
    case RECEIPT_9:
    case RECEIPT_10:

    Import = getConfigValue(IRC)                                                     // IRC = Import to Receipt
    break;
//-------------------------------------------------------------------Warrant-------------------------------------------------------------------//
    case WARRANT_11:
    case WARRANT_12:
    case WARRANT_13:

    Import = getConfigValue(IWT)                                                     // IWT = Import to Warrant
    break;
//-------------------------------------------------------------------Block-------------------------------------------------------------------//
    case BLOCK:

    Import = getConfigValue(IBK)                                                     // IBK = Import to Block
    break;
      
    default:
      Import = null;
    break;
  }

//-------------------------------------------------------------------Foot-------------------------------------------------------------------//

  if (Import == "TRUE") 
  {
    var Check = sheet_sr.getRange("A5").getValue();

    if( Check !== "" ) 
    {
      var LR = sheet_sr.getLastRow();
      var LC = sheet_sr.getLastColumn();
      
      var Data = sheet_sr.getRange(5,1,LR-4,LC).getValues();
        sheet_tr.getRange(5,1,LR-4,LC).setValues(Data);
      var Data_1 = sheet_sr.getRange(1,1,1,LC).getValues();
        sheet_tr.getRange(1,1,1,LC).setValues(Data_1);

        Logger.log(`SUCCESS IMPORT. Sheet:  ${SheetName}.`);
    }
    else
    {
      Logger.log(`ERROR IMPORT: ${SheetName} - A5 cell is Blank on doImportBasic`);
    }
  }
  else
  {
    Logger.log(`ERROR IMPORT: ${SheetName} - IMPORT on config is set to FALSE`);
  }
}

/////////////////////////////////////////////////////////////////////FINANCIAL////////////////////////////////////////////////////////////////////

function doImportFinancial(SheetName){
  Logger.log(`IMPORT: ${SheetName}`);
  const sheet_co = fetchSheetByName('Config');                                        // Config sheet
  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID
  const sheet_se = fetchSheetByName('Settings');                                      // Settings sheet
  if (!sheet_co || !sheet_se) return;
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName(SheetName);      // Source Sheet
  if (!sheet_sr) { Logger.log(`ERROR IMPORT: ${SheetName} - Does not exist on doImportFinancial from ${Source_Id}`); return; }
    var LR = sheet_sr.getLastRow();
    var LC = sheet_sr.getLastColumn();
  const sheet_tr = fetchSheetByName(SheetName);                                       // Target Sheet
  if (!sheet_tr) { Logger.log(`WARNING: Target sheet ${SheetName} - does not exist on doImportSheet. Skipping.`); return; }

  let Import;

//-------------------------------------------------------------------BLC-------------------------------------------------------------------//

  if (SheetName === BLC) 
  {
    Import = getConfigValue(IBL)                                                     // IBL = Import to BLC / Balanco

    if ( Import == "TRUE" )
    {
      var Check = sheet_sr.getRange("B1").getValue();

      if( Check !== "" )
      {
        var Data = sheet_sr.getRange(1,2,LR,LC-1).getValues();
        sheet_tr.getRange(1,2,LR,LC-1).setValues(Data);
      }
      else
      {
        Logger.log(`ERROR IMPORT: ${SheetName} - B1 cell is Blank on doImportFinancial`);
      }
    }
    else
    {
      Logger.log(`ERROR IMPORT: ${SheetName} - IMPORT on config is set to FALSE`);
    }
  }

//-------------------------------------------------------------------Balanco-------------------------------------------------------------------//

  if (SheetName === Balanco) 
  {
    Import = getConfigValue(IBL)                                                     // IBL = Import to BLC / Balanco

    if ( Import == "TRUE" )
    {
      var Check = sheet_sr.getRange("C1").getValue();

      if( Check !== "" )
      {
        var Data = sheet_sr.getRange(1,3,LR,LC-2).getValues();
        sheet_tr.getRange(1,3,LR,LC-2).setValues(Data);
      }
      else
      {
        Logger.log(`ERROR IMPORT: ${SheetName} - C1 cell is Blank on doImportFinancial`);
      }
    }
    else
    {
      Logger.log(`ERROR IMPORT: ${SheetName} - IMPORT on config is set to FALSE`);
    }
  }

//-------------------------------------------------------------------DRE-------------------------------------------------------------------//

  if (SheetName === DRE) 
  {
    Import = getConfigValue(IDE)                                                     // IDE = Import to DRE / Resultado

    if ( Import == "TRUE" )
    {
      var Check = sheet_sr.getRange("B1").getValue();

      if( Check !== "" )
      {
        var Data = sheet_sr.getRange(1,2,LR,LC-1).getValues();
        sheet_tr.getRange(1,2,LR,LC-1).setValues(Data);
      }
      else
      {
        Logger.log(`ERROR IMPORT: ${SheetName} - B1 cell is Blank on doImportFinancial`);
      }
    }
    else
    {
      Logger.log(`ERROR IMPORT: ${SheetName} - IMPORT on config is set to FALSE`);
    }
  }

//-------------------------------------------------------------------Resultado-------------------------------------------------------------------//

  if (SheetName === Resultado) 
  {
    Import = getConfigValue(IDE)                                                     // IDE = Import to DRE / Resultado

    if ( Import == "TRUE" )
    {
      var Check = sheet_sr.getRange("D1").getValue();

      if( Check !== "" )
      {
        var Data = sheet_sr.getRange(1,4,LR,LC-3).getValues();
        sheet_tr.getRange(1,4,LR,LC-3).setValues(Data);
      }
      else
      {
        Logger.log(`ERROR IMPORT: ${SheetName} - D1 cell is Blank on doImportFinancial`);
      }
    }
    else
    {
      Logger.log(`ERROR IMPORT: ${SheetName} - IMPORT on config is set to FALSE`);
    }
  }

//-------------------------------------------------------------------FLC-------------------------------------------------------------------//

  if (SheetName === FLC) 
  {
    Import = getConfigValue(IFL)                                                     // IFL = Import to FLC / Fluxo

    if ( Import == "TRUE" )
    {
      var Check = sheet_sr.getRange("B1").getValue();

      if( Check !== "" )
      {
        var Data = sheet_sr.getRange(1,2,LR,LC-1).getValues();
        sheet_tr.getRange(1,2,LR,LC-1).setValues(Data);
      }
      else
      {
        Logger.log(`ERROR IMPORT: ${SheetName} - B1 cell is Blank on doImportFinancial`);
      }
    }
    else
    {
      Logger.log(`ERROR IMPORT: ${SheetName} - IMPORT on config is set to FALSE`);
    }
  }

//-------------------------------------------------------------------Fluxo-------------------------------------------------------------------//

  if (SheetName === Fluxo) 
  {
    Import = getConfigValue(IFL)                                                     // IFL = Import to FLC / Fluxo

    if ( Import == "TRUE" )
    {
      var Check = sheet_sr.getRange("D1").getValue();

      if( Check !== "" )
      {
        var Data = sheet_sr.getRange(1,4,LR,LC-3).getValues();
        sheet_tr.getRange(1,4,LR,LC-3).setValues(Data);
      }
      else
      {
        Logger.log(`ERROR IMPORT: ${SheetName} - C1 cell is Blank on doImportFinancial`);
      }
    }
    else
    {
      Logger.log(`ERROR IMPORT: ${SheetName} - IMPORT on config is set to FALSE`);
    }
  }

//-------------------------------------------------------------------DVA-------------------------------------------------------------------//

  if (SheetName === DVA) 
  {
    Import = getConfigValue(IDV)                                                     // IDV = Import to DVA / Valor

    if ( Import == "TRUE" )
    {
      var Check = sheet_sr.getRange("B1").getValue();

      if( Check !== "" )
      {
        var Data = sheet_sr.getRange(1,2,LR,LC-1).getValues();
        sheet_tr.getRange(1,2,LR,LC-1).setValues(Data);
      }
      else
      {
        Logger.log(`ERROR IMPORT: ${SheetName} - B1 cell is Blank on doImportFinancial`);
      }
    }
    else
    {
      Logger.log(`ERROR IMPORT: ${SheetName} - IMPORT on config is set to FALSE`);
    }
  }

//-------------------------------------------------------------------Valor-------------------------------------------------------------------//

  if (SheetName === Valor) 
  {
    Import = getConfigValue(IDV)                                                     // IDV = Import to DVA / Valor

    if ( Import == "TRUE" )
    {
      var Check = sheet_sr.getRange("D1").getValue();

      if( Check !== "" )
      {
        var Data = sheet_sr.getRange(1,4,LR,LC-3).getValues();
        sheet_tr.getRange(1,4,LR,LC-3).setValues(Data);
      }
      else
      {
        Logger.log(`ERROR IMPORT: ${SheetName} - C1 cell is Blank on doImportFinancial`);
      }
    }
    else
    {
      Logger.log(`ERROR IMPORT: ${SheetName} - IMPORT on config is set to FALSE`);
    }
  }
  Logger.log(`SUCCESS IMPORT for sheet ${SheetName}.`);
}

/////////////////////////////////////////////////////////////////////IMPORT TEMPLATE/////////////////////////////////////////////////////////////////////