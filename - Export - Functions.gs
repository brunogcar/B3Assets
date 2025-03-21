/////////////////////////////////////////////////////////////////////ID/////////////////////////////////////////////////////////////////////

function getSheetID()
{
  const sheet_Id = SpreadsheetApp.getActiveSpreadsheet().getId();

  return sheet_Id;
};

function setSheetID()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');                                     // Config sheet
  if (!sheet_co) return;

  var Target_Id = sheet_co.getRange(DIR).getValues();                              // DIR = DATA ID Range

  var TKT = sheet_co.getRange(TKR).getDisplayValue();                              // TKR = Ticket 
  var EXP = sheet_co.getRange(EXR).getDisplayValue();                              // EXR = Export 
  var SHI = sheet_co.getRange(ICR).getDisplayValue();                              // ICR = Sheet ID 

  Logger.log('Setting Sheet ID');

  const Sheet_Id = SpreadsheetApp.getActiveSpreadsheet().getId();

//  const [sheet_co, sheet_tr] = ["Config", "Relação"].map(s => ss.getSheetByName(s));
//  const [b3, d10] = ["B3", "D10"].map(r => sheet_co.getRange(r).getValue());

  var ss_t = SpreadsheetApp.openById(Target_Id);                                    // Target spreadsheet
  var sheet_tr = ss_t.getSheetByName('Relação');                                    // Target sheet

  var bgcolor = sheet_co.getRange(IDR).getBackground();
  var colour = '#d9ead3';

  const search = sheet_tr.getRange("A2:A" + sheet_tr.getLastRow()).createTextFinder(TKT).findNext();
  if (!search) return;
  if( bgcolor == colour)
  {
    if ( EXP == "TRUE" && SHI != "TRUE")                                           //Check conditions to export Sheet ID
    {
      search.offset(0, 11).setValue(Sheet_Id);
      search.offset(0, 12).setValue(SNAME(3));

      Logger.log(`Sheet ID Set: ${Sheet_Id}`);
    }
  }
};

function doIsIdExported()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');                                     // Config sheet
  if (!sheet_co) return;

  var IEP = sheet_co.getRange(IER).getDisplayValue();                              // EXR = Export Range

  if( IEP === "FALSE" )
  {
    setSheetID()
  }
};

function doClearSheetID()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');                                     // Config sheet
  if (!sheet_co) return;

  var Target_Id = sheet_co.getRange(DIR).getValues();                              // DIR = DATA ID Range

  var TKT = sheet_co.getRange(TKR).getDisplayValue();                              // TKR = Ticket Range

  var dst = SpreadsheetApp.openById(Target_Id);                                    // Target spreadsheet
  var sheet_tr = dst.getSheetByName('Relação');                                    // Target sheet

  const search = sheet_tr.getRange("A2:A" + sheet_tr.getLastRow()).createTextFinder(TKT).findNext();
  if (!search) return;
  {
    search.offset(0, 11, 1, 2).clearContent();

    Logger.log('Sheet ID Cleared');
  }
};

/////////////////////////////////////////////////////////////////////EXPORT CHECKS/////////////////////////////////////////////////////////////////////

function doIsFormula()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');                                     // Config sheet
  if (!sheet_co) return;

  var Formula = sheet_co.getRange(FOR).getDisplayValue();                          // FOR = Formula Range
  var Sheet_ID = sheet_co.getRange(ICR).getDisplayValue();                         // ICR = Sheet ID Check Range

  if( Formula == "TRUE" ) //Check if formula true to export info
  {
    doIsExportable()
  }
  else if ( Formula == "FALSE" && Sheet_ID != "TRUE" )                             //Check if formula true to export info
  {
    setSheetID()
  }
};

function doIsExportable()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');                                    // Config sheet
  if (!sheet_co) return;

  var EPD = sheet_co.getRange(EPR).getDisplayValue();                             // EPR = Exportable? Check Range
  var EXP = sheet_co.getRange(EXR).getDisplayValue();                             // EXR = Export Range

  if( EPD === "TRUE" )
  {
    doIsInfoExported()
  }
};

function doIsInfoExported()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');                                   // Config sheet
  if (!sheet_co) return;

  var EXP = sheet_co.getRange(EXR).getDisplayValue();                            // EXR = Export Range

  if( EXP === "TRUE" )
  {
    const sheet_in = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Info');

    const Data = sheet_in.getRange(TIR).getValues();                             // TIR = Tab Info Range
    sheet_in.getRange(TIR).setValues(Data);                                      // Copy Paste Info   

    const Data_2 = sheet_co.getRange(EXR).getValues();                           // TIR = Tab Info Range
    sheet_co.getRange(EXR).setValues(Data_2);                                    // Set Formula to TRUE // EXP === "TRUE"

    setSheetID()
  }
  else if ( EXP !== "TRUE" )
  {
    doExportInfo()
  }
};

/////////////////////////////////////////////////////////////////////CLEAR EXPORTED to EXPORTED Source/////////////////////////////////////////////////////////////////////

function clearExportALL() 
{
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUTURE, FUND, BLC, DRE, FLC, DVA];

  SheetNames.forEach(SheetName => 
  {
    clearExport(SheetName);
  });
}

function clearExport(SheetName) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');                                   // Config sheet
  if (!sheet_co) return;

  var TKT = sheet_co.getRange(TKR).getDisplayValue();                            // TKR = Ticket Range
  var Target_Id = sheet_co.getRange(TDR).getValues();                            // TDR = Target ID Range

  const trg = SpreadsheetApp.openById(Target_Id);                                // Target spreadsheet
  const sheet_tr = trg.getSheetByName(SheetName);                                // Target sheet

  let success = false; // Initialize success flag to false

  var search = sheet_tr.getRange("A2:A" + sheet_tr.getLastRow()).createTextFinder(TKT).findNext();

  Logger.log('Clear Export:', SheetName);

  if (search) 
  {
    search.offset(0, 0, 1, sheet_tr.getLastColumn()).clearContent();

    success = true; // Set the success flag to true if data was cleared
  }
  if (success) 
  {
    Logger.log(`Exported data cleared successfully. Sheet: ${SheetName}.`);
  } 
  else 
  {
    Logger.log('Clear EXPORT:', SheetName, "Didn't find Ticket:", TKT);
  }
}

/////////////////////////////////////////////////////////////////////EXPORT FUNCTIONS/////////////////////////////////////////////////////////////////////