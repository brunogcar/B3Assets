//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////IMPORT/////////////////////////////////////////////////////////////////////

function Import() 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();                              // Active spreadsheet
  const sheet_co = ss.getSheetByName('Config');                                  // Config sheet
  const Source_Id = sheet_co.getRange(SIR).getDisplayValue().trim();             // SIR = Source ID
  const Option = sheet_co.getRange(OPR).getDisplayValue();                       // OPR = Option
  const sheet_sr = SpreadsheetApp.openById(Source_Id);                           // Open source spreadsheet by ID


  // Helper function to check if a sheet exists in the source spreadsheet
  function sheetExists(SheetName) 
  {
    const exists = sheet_sr.getSheetByName(SheetName) !== null;
    Logger.log(`Sheet "${SheetName}" exists: ${exists}`); // Log the result
    return exists;
  }

  if (Option === "AUTO") 
  {
    // Check for specific sheets
    const hasSwing4 = sheetExists(SWING_4);
    const hasSwing12 = sheetExists(SWING_12);
    const hasSwing52 = sheetExists(SWING_52);
    const hasTrade = sheetExists('Trade');

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
      Logger.log('No matching sheets found for AUTO mode.');
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
      import_15x_to_161();
    }
    else 
    {
      Logger.log(`Invalid Option: ${Option}`);
    }
  }
}


/////////////////////////////////////////////////////////////////////MENU/////////////////////////////////////////////////////////////////////

function import_Current()
{
  console.log('Import: import_Current');

  import_config();

  doImportProventos();

  doImportShares();

  doImportData(BLC);
  doImportData(Balanco);
  doImportData(DRE);
  doImportData(Resultado);
  doImportData(FLC);
  doImportData(Fluxo);
  doImportData(DVA);
  doImportData(Valor);

  doImportSheet(SWING_4);
  doImportSheet(SWING_12);
  doImportSheet(SWING_52);
  doImportSheet(OPCOES);
  doImportSheet(BTC);
  doImportSheet(TERMO);
  doImportSheet(FUND);

  doImportSheet(FUTURE);
  doImportSheet(FUTURE_1);
  doImportSheet(FUTURE_2);
  doImportSheet(FUTURE_3);

  doImportSheet(RIGHT_1);
  doImportSheet(RIGHT_2);
  doImportSheet(RECEIPT_9);
  doImportSheet(RECEIPT_10);
  doImportSheet(WARRANT_11);
  doImportSheet(WARRANT_12);
  doImportSheet(WARRANT_13);
  doImportSheet(BLOCK);

  doCheckTriggers();
  update_form();

// doCleanZeros();
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
      console.error(`Error importing ${ProvName}:`, error);
    }
  });
}

function doImportSheets() 
{
  const SheetNames = [
    SWING_4, SWING_52, OPCOES, BTC, TERMO, FUND, 
    FUTURE, FUTURE_1, FUTURE_2, FUTURE_3, 
    RIGHT_1, RIGHT_2, 
    RECEIPT_9, RECEIPT_10, 
    WARRANT_11, WARRANT_12, WARRANT_13, 
    BLOCK
  ];

  SheetNames.forEach(SheetName =>
  {
    try
    {
      doImportSheet(SheetName);
    }
    catch (error)
    {
      // Handle the error here, you can log it or take appropriate actions.
      console.error(`Error importing sheet ${SheetName}:`, error);
    }
  });
}


function doImportDatas() 
{
  const SheetNames = [BLC, Balanco, DRE, Resultado, FLC, Fluxo, DVA, Valor];

  SheetNames.forEach(SheetName => 
  {
    try 
    {
      doImportData(SheetName);
    } 
    catch (error) 
    {
      // Handle the error here, you can log it or take appropriate actions.
      console.error(`Error importing sheet ${SheetName}:`, error);
    }
  });
}

/////////////////////////////////////////////////////////////////////Update Form/////////////////////////////////////////////////////////////////////

function update_form() 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var Update_Form = ss.getRange(UFR).getDisplayValue();                       // UFR = Update Form

  switch (Update_Form) 
  {
    case 'EDIT':
      doEditAll();
      break;
    case 'SAVE':
      doSaveAll();
      break;

    default:
      console.error('Invalid update form value:', Update_Form);
      break;
  }
}

/////////////////////////////////////////////////////////////////////Functions/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////Config/////////////////////////////////////////////////////////////////////

function import_config()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet_co = ss.getSheetByName('Config');
    var Source_Id = sheet_co.getRange(SIR).getValues();                           // SIR = Source ID
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('Config');   // Source Sheet
  {
    var Data = sheet_sr.getRange(COR).getValues();
    sheet_co.getRange(COR).setValues(Data);
  }
};

/////////////////////////////////////////////////////////////////////SHARES and FF/////////////////////////////////////////////////////////////////////

function doImportShares() 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = ss.getSheetByName('Config');
    var Source_Id = sheet_co.getRange(SIR).getValues();                          // SIR = Source ID
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('DATA');    // Source Sheet
    var L1 = sheet_sr.getRange('L1').getValue();
    var L2 = sheet_sr.getRange('L2').getValue();
  const sheet_tr = ss.getSheetByName('DATA');                                    // Target Sheet
    var SheetName =  sheet_tr.getName()

  console.log('IMPORT: Shares and FF');

  if (!ErrorValues.includes(L1) && !ErrorValues.includes(L2)) 
  {
    var Data = sheet_sr.getRange("L1:L2").getValues();
    sheet_tr.getRange("L1:L2").setValues(Data);
  }
  else
  {
    console.log('ERROR IMPORT:', SheetName, 'ErrorValues on L1 or L2 on doImportShares');
  }
console.log(`SUCESS IMPORT: Shares and FF`);
}

/////////////////////////////////////////////////////////////////////Proventos/////////////////////////////////////////////////////////////////////

function doImportProv(ProvName)
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = ss.getSheetByName('Config');
    var Source_Id = sheet_co.getRange(SIR).getValues();                           // SIR = Source ID
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('Prov');     // Source Sheet
  const sheet_tr = ss.getSheetByName('Prov');  

  console.log('IMPORT:', ProvName);

  if (ProvName == 'Proventos') 
  {
    var Check = sheet_sr.getRange('B3').getDisplayValue();  

    if( Check == "Proventos" )  // check if error
    {
      var Data = sheet_sr.getRange(PRV).getValues();                              // PRV = Provento Range
      sheet_tr.getRange(PRV).setValues(Data);
    }
    else
    {
      console.log('ERROR IMPORT:', ProvName, 'B3 != Proventos on doImportProv');
    }
  }
}

/////////////////////////////////////////////////////////////////////Sheets/////////////////////////////////////////////////////////////////////

function doImportSheet(SheetName) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = ss.getSheetByName('Config');
    var Source_Id = sheet_co.getRange(SIR).getValues();                             // SIR = Source ID
  const sheet_se = ss.getSheetByName('Settings');
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName(SheetName);    // Source Sheet
  if (!sheet_sr) 
  {
    console.log('ERROR IMPORT:', SheetName, 'Does not exist on doImportSheet from', Source_Id);
    return;
  }

  const sheet_tr = ss.getSheetByName(SheetName);                                    // Target Sheet
  if (!sheet_tr) 
  {
      console.log('WARNING: Target sheet', SheetName, 'does not exist. Skipping.');
    return; // Skip to the next sheet
  }
  console.log('IMPORT:', SheetName);

  let Import;
  let Value_se = "DEFAULT";                                                         // Initialize Value_se with "DEFAULT" as the fallback

  switch (SheetName) 
  {
    case SWING_4:
    case SWING_12:
    case SWING_52:

    Value_se = sheet_se.getRange(ITR).getDisplayValue().trim();                    // ITR = Import to Swing

    Import = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(ITR).getDisplayValue().trim()                            // Use Config value if Settings has "DEFAULT"
      : Value_se;                                                                  // Use Settings value

    break;

    case OPCOES:

    Value_se = sheet_se.getRange(IOP).getDisplayValue().trim();                    // IOP = Import to Option

    Import = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(IOP).getDisplayValue().trim()
      : Value_se; 

    break;

    case BTC:

    Value_se = sheet_se.getRange(IBT).getDisplayValue().trim();                    // IBT = Import to BTC

    Import = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(IBT).getDisplayValue().trim()
      : Value_se; 

    break;

    case TERMO:

    Value_se = sheet_se.getRange(ITE).getDisplayValue().trim();                    // ITE = Import to Termo

    Import = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(ITE).getDisplayValue().trim()
      : Value_se; 

    break;

    case FUTURE:
    case FUTURE_1:
    case FUTURE_2:
    case FUTURE_3:

    Value_se = sheet_se.getRange(IFT).getDisplayValue().trim();                    // IFT = Import to Future

    Import = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(IFT).getDisplayValue().trim()
      : Value_se; 

    break;

    case FUND:

    Value_se = sheet_se.getRange(IFU).getDisplayValue().trim();                    // IFU = Import to Fund

    Import = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(IFU).getDisplayValue().trim()
      : Value_se; 

    break;

    case RIGHT_1:
    case RIGHT_2:

    Value_se = sheet_se.getRange(IRT).getDisplayValue().trim();                    // IRT = Import to Right

    Import = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(IRT).getDisplayValue().trim()
      : Value_se; 

    break;

    case RECEIPT_9:
    case RECEIPT_10:

    Value_se = sheet_se.getRange(IRC).getDisplayValue().trim();                    // IRC = Import to Receipt

    Import = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(IRC).getDisplayValue().trim()
      : Value_se; 

    break;

    case WARRANT_11:
    case WARRANT_12:
    case WARRANT_13:

    Value_se = sheet_se.getRange(IWT).getDisplayValue().trim();                    // IWT = Import to Warrant

    Import = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(IWT).getDisplayValue().trim()
      : Value_se; 

    break;

    case BLOCK:

    Value_se = sheet_se.getRange(IBK).getDisplayValue().trim();                    //  IBK = Import to Block

    Import = (Value_se === "DEFAULT") 
      ? sheet_co.getRange(IBK).getDisplayValue().trim()
      : Value_se; 

    break;
      
    default:
      var Source_Id = null;
      Import = null;
    break;
  }

//-------------------------------------------------------------------Foot-------------------------------------------------------------------//

  if (Import == "TRUE") 
  {
    var Check = sheet_sr.getRange('A5').getValue();

    if( Check !== "" ) 
    {
      var LR = sheet_sr.getLastRow();
      var LC = sheet_sr.getLastColumn();
      
      var Data = sheet_sr.getRange(5,1,LR-4,LC).getValues();
        sheet_tr.getRange(5,1,LR-4,LC).setValues(Data);
      var Data_1 = sheet_sr.getRange(1,1,1,LC).getValues();
        sheet_tr.getRange(1,1,1,LC).setValues(Data_1);

        console.log(`SUCESS IMPORT. Sheet:  ${SheetName}.`);
    }
    else
    {
      console.log('ERROR IMPORT:', SheetName, 'A5 cell is Blank on doImportSheet');
    }
  }
  else
  {
    console.log('ERROR IMPORT:', SheetName, 'IMPORT on config is set to FALSE');
  }
}

/////////////////////////////////////////////////////////////////////DATA////////////////////////////////////////////////////////////////////

function doImportData(SheetName) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet_co = ss.getSheetByName('Config');
    var Source_Id = sheet_co.getRange(SIR).getValues();                                  // SIR = Source ID Range
  const sheet_se = ss.getSheetByName('Settings');
  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName(SheetName);         // Source Sheet

  if (!sheet_sr) 
  {
    console.log('ERROR IMPORT:', SheetName, 'Does not exist on doImportData from', Source_Id);
    return;                                                                              // Exit the function
  }

    var LR = sheet_sr.getLastRow();
    var LC = sheet_sr.getLastColumn();
  const sheet_tr = ss.getSheetByName(SheetName);                                         // Target Sheet

  console.log('IMPORT:', SheetName);

  let Import;
  let Value_se = "DEFAULT";   // Initialize Value_se with "DEFAULT" as the fallback

//-------------------------------------------------------------------BLC-------------------------------------------------------------------//

   if (SheetName === BLC) 
   {
     Value_se = sheet_se.getRange(IBL).getDisplayValue().trim();            // IBL = Import to BLC / Balanco

     if (Value_se === "DEFAULT") 
     {
       Import = sheet_co.getRange(IBL).getDisplayValue().trim();            // Use Config value if Settings has "DEFAULT"
     } 
     else 
     {
       Import = Value_se;                                                   // Use the imported value from Settings
     }

     if ( Import == "TRUE" )
     {
       var Check = sheet_sr.getRange('B1').getValue();

       if( Check !== "" )
       {
         var Data = sheet_sr.getRange(1,2,LR,LC-1).getValues();
         sheet_tr.getRange(1,2,LR,LC-1).setValues(Data);
       }
       else
       {
         console.log('ERROR IMPORT:', SheetName, 'B1 cell is Blank on doImportData');
       }
     }
     else
     {
       console.log('ERROR IMPORT:', SheetName, 'IMPORT on config is set to FALSE');
     }
   }

//-------------------------------------------------------------------Balanco-------------------------------------------------------------------//

   if (SheetName === Balanco) 
   {
     Value_se = sheet_se.getRange(IBL).getDisplayValue().trim();            // IBL = Import to BLC / Balanco

     if (Value_se === "DEFAULT") 
     {
       Import = sheet_co.getRange(IBL).getDisplayValue().trim();            // Use Config value if Settings has "DEFAULT"
     } 
     else 
     {
       Import = Value_se;                                                   // Use the imported value from Settings
     }

     if ( Import == "TRUE" )
     {
       var Check = sheet_sr.getRange('C1').getValue();

       if( Check !== "" )
       {
         var Data = sheet_sr.getRange(1,3,LR,LC-2).getValues();
         sheet_tr.getRange(1,3,LR,LC-2).setValues(Data);
       }
       else
       {
         console.log('ERROR IMPORT:', SheetName, 'C1 cell is Blank on doImportData');
       }
     }
     else
     {
       console.log('ERROR IMPORT:', SheetName, 'IMPORT on config is set to FALSE');
     }
   }

//-------------------------------------------------------------------DRE-------------------------------------------------------------------//

   if (SheetName === DRE) 
   {
     Value_se = sheet_se.getRange(IDE).getDisplayValue().trim();            // IDE = Import to DRE / Resultado

     if (Value_se === "DEFAULT") 
     {
       Import = sheet_co.getRange(IDE).getDisplayValue().trim();            // Use Config value if Settings has "DEFAULT"
     } 
     else 
     {
       Import = Value_se;                                                   // Use the imported value from Settings
     }

     if ( Import == "TRUE" )
     {
       var Check = sheet_sr.getRange('B1').getValue();

       if( Check !== "" )
       {
         var Data = sheet_sr.getRange(1,2,LR,LC-1).getValues();
         sheet_tr.getRange(1,2,LR,LC-1).setValues(Data);
       }
       else
       {
         console.log('ERROR IMPORT:', SheetName, 'B1 cell is Blank on doImportData');
       }
     }
     else
     {
       console.log('ERROR IMPORT:', SheetName, 'IMPORT on config is set to FALSE');
     }
   }

//-------------------------------------------------------------------Resultado-------------------------------------------------------------------//

   if (SheetName === Resultado) 
   {
     Value_se = sheet_se.getRange(IDE).getDisplayValue().trim();            // IDE = Import to DRE / Resultado

     if (Value_se === "DEFAULT") 
     {
       Import = sheet_co.getRange(IDE).getDisplayValue().trim();            // Use Config value if Settings has "DEFAULT"
     } 
     else 
     {
       Import = Value_se;                                                   // Use the imported value from Settings
     }

     if ( Import == "TRUE" )
     {
       var Check = sheet_sr.getRange('D1').getValue();

       if( Check !== "" )
       {
         var Data = sheet_sr.getRange(1,4,LR,LC-3).getValues();
         sheet_tr.getRange(1,4,LR,LC-3).setValues(Data);
       }
       else
       {
         console.log('ERROR IMPORT:', SheetName, 'D1 cell is Blank on doImportData');
       }
     }
     else
     {
       console.log('ERROR IMPORT:', SheetName, 'IMPORT on config is set to FALSE');
     }
   }

//-------------------------------------------------------------------FLC-------------------------------------------------------------------//

   if (SheetName === FLC) 
   {
     Value_se = sheet_se.getRange(IFL).getDisplayValue().trim();            // IFL = Import to FLC / Fluxo

     if (Value_se === "DEFAULT") 
     {
       Import = sheet_co.getRange(IFL).getDisplayValue().trim();            // Use Config value if Settings has "DEFAULT"
     } 
     else 
     {
       Import = Value_se;                                                   // Use the imported value from Settings
     }

     if ( Import == "TRUE" )
     {
       var Check = sheet_sr.getRange('B1').getValue();

       if( Check !== "" )
       {
         var Data = sheet_sr.getRange(1,2,LR,LC-1).getValues();
         sheet_tr.getRange(1,2,LR,LC-1).setValues(Data);
       }
       else
       {
         console.log('ERROR IMPORT:', SheetName, 'B1 cell is Blank on doImportData');
       }
     }
     else
     {
       console.log('ERROR IMPORT:', SheetName, 'IMPORT on config is set to FALSE');
     }
   }

//-------------------------------------------------------------------Fluxo-------------------------------------------------------------------//

   if (SheetName === Fluxo) 
   {
     Value_se = sheet_se.getRange(IFL).getDisplayValue().trim();            // IFL = Import to FLC / Fluxo

     if (Value_se === "DEFAULT") 
     {
       Import = sheet_co.getRange(IFL).getDisplayValue().trim();            // Use Config value if Settings has "DEFAULT"
     } 
     else 
     {
       Import = Value_se;                                                   // Use the imported value from Settings
     }

     if ( Import == "TRUE" )
     {
       var Check = sheet_sr.getRange('D1').getValue();

       if( Check !== "" )
       {
         var Data = sheet_sr.getRange(1,4,LR,LC-3).getValues();
         sheet_tr.getRange(1,4,LR,LC-3).setValues(Data);
       }
       else
       {
         console.log('ERROR IMPORT:', SheetName, 'C1 cell is Blank on doImportData');
       }
     }
     else
     {
       console.log('ERROR IMPORT:', SheetName, 'IMPORT on config is set to FALSE');
     }
   }

//-------------------------------------------------------------------DVA-------------------------------------------------------------------//

   if (SheetName === DVA) 
   {
     Value_se = sheet_se.getRange(IDV).getDisplayValue().trim();            // IDV = Import to DVA / Valor

     if (Value_se === "DEFAULT") 
     {
       Import = sheet_co.getRange(IDV).getDisplayValue().trim();            // Use Config value if Settings has "DEFAULT"
     } 
     else 
     {
       Import = Value_se;                                                   // Use the imported value from Settings
     }

     if ( Import == "TRUE" )
     {
       var Check = sheet_sr.getRange('B1').getValue();

       if( Check !== "" )
       {
         var Data = sheet_sr.getRange(1,2,LR,LC-1).getValues();
         sheet_tr.getRange(1,2,LR,LC-1).setValues(Data);
       }
       else
       {
         console.log('ERROR IMPORT:', SheetName, 'B1 cell is Blank on doImportData');
       }
     }
     else
     {
       console.log('ERROR IMPORT:', SheetName, 'IMPORT on config is set to FALSE');
     }
   }

//-------------------------------------------------------------------Valor-------------------------------------------------------------------//

   if (SheetName === Valor) 
   {
     Value_se = sheet_se.getRange(IDV).getDisplayValue().trim();            // IDV = Import to DVA / Valor

     if (Value_se === "DEFAULT") 
     {
       Import = sheet_co.getRange(IDV).getDisplayValue().trim();            // Use Config value if Settings has "DEFAULT"
     } 
     else 
     {
       Import = Value_se;                                                   // Use the imported value from Settings
     }

     if ( Import == "TRUE" )
     {
       var Check = sheet_sr.getRange('D1').getValue();

       if( Check !== "" )
       {
         var Data = sheet_sr.getRange(1,4,LR,LC-3).getValues();
         sheet_tr.getRange(1,4,LR,LC-3).setValues(Data);
       }
       else
       {
         console.log('ERROR IMPORT:', SheetName, 'C1 cell is Blank on doImportData');
       }
     }
     else
     {
       console.log('ERROR IMPORT:', SheetName, 'IMPORT on config is set to FALSE');
     }
   }
   console.log(`SUCESS IMPORT for sheet ${SheetName}.`);
}

/////////////////////////////////////////////////////////////////////IMPORT/////////////////////////////////////////////////////////////////////