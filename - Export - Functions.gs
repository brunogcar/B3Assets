/////////////////////////////////////////////////////////////////////ID/////////////////////////////////////////////////////////////////////

function getSheetID()
{
  const sheet_Id = SpreadsheetApp.getActiveSpreadsheet().getId();
  return sheet_Id;
};

function setSheetID(){
  const sheet_co = fetchSheetByName('Config');                                     // Config sheet
  if (!sheet_co) return;

  const Data_Id = getConfigValue(DIR, 'Config');                                   // DIR = DATA Source ID
  const TKT = getConfigValue(TKR, 'Config');                                       // TKR = Ticket Range
  const EXP = getConfigValue(EXR, 'Config');                                       // EXR = Export 
  const SHI = getConfigValue(ICR, 'Config');                                       // ICR = Sheet ID 

  Logger.log('Setting Sheet ID');

  const Sheet_Id = SpreadsheetApp.getActiveSpreadsheet().getId();

//  const [sheet_co, sheet_tr] = ["Config", "Relação"].map(s => ss.getSheetByName(s));
//  const [b3, d10] = ["B3", "D10"].map(r => sheet_co.getRange(r).getValue());

  var ss_tr = SpreadsheetApp.openById(Data_Id);                                    // Target spreadsheet
  var sheet_tr = ss_tr.getSheetByName('Relação');                                    // Target sheet
  if (!sheet_tr) {Logger.log(`Target sheet not found: ${SheetName}`); return;}

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

function doIsIdExported(){
  const sheet_co = fetchSheetByName('Config');                                     // Config sheet
  if (!sheet_co) return;

  var IEP = sheet_co.getRange(IER).getDisplayValue();                              // EXR = Export Range

  if( IEP === "FALSE" )
  {
    setSheetID()
  }
};

function doClearSheetID(){
  const sheet_co = fetchSheetByName('Config');                                     // Config sheet
  if (!sheet_co) return;

  const Data_Id = getConfigValue(DIR, 'Config');                                   // DIR = DATA Source ID
  const TKT = getConfigValue(TKR, 'Config');                                       // TKR = Ticket Range

  var ss_tr = SpreadsheetApp.openById(Data_Id);                                    // Target spreadsheet
  var sheet_tr = ss_tr.getSheetByName('Relação');                                  // Target sheet
  if (!sheet_tr) {Logger.log(`Target sheet not found: ${SheetName}`); return;}

  const search = sheet_tr.getRange("A2:A" + sheet_tr.getLastRow()).createTextFinder(TKT).findNext();
  if (!search) return;
  {
    search.offset(0, 11, 1, 2).clearContent();

    Logger.log('Sheet ID Cleared');
  }
};

/////////////////////////////////////////////////////////////////////EXPORT CHECKS/////////////////////////////////////////////////////////////////////

function doIsFormula(){
  const sheet_co = fetchSheetByName('Config');                                     // Config sheet
  if (!sheet_co) return;

  const Formula = getConfigValue(FOR, 'Config');                                   // FOR = Formula Range
  const Sheet_Id = getConfigValue(ICR, 'Config');                                  // ICR = Sheet ID Check Range

  if( Formula == "TRUE" ) //Check if formula true to export info
  {
    doIsExportable()
  }
  else if ( Formula == "FALSE" && Sheet_Id != "TRUE" )                             //Check if formula true to export info
  {
    setSheetID()
  }
};

function doIsExportable(){
  const sheet_co = fetchSheetByName('Config');                                    // Config sheet
  if (!sheet_co) return;

  const EPD = getConfigValue(EPR, 'Config');                                      // EPR = Exportable? Check Range

  if( EPD === "TRUE" )
  {
    doIsInfoExported()
  }
};

function doIsInfoExported(){
  const sheet_co = fetchSheetByName('Config');                                   // Config sheet
  if (!sheet_co) return;

  const EXP = getConfigValue(EXR, 'Config');                                     // EXR = Export 

  if( EXP === "TRUE" )
  {
    const sheet_in = fetchSheetByName('Info');

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

function doClearExportAll(){
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUTURE, FUND, BLC, DRE, FLC, DVA];

  SheetNames.forEach(SheetName => 
  {
    doClearExport(SheetName);
  });
}

function doClearExport(SheetName){
  const sheet_co = fetchSheetByName('Config');                                   // Config sheet
  if (!sheet_co) return;

  const Target_Id = getConfigValue(TDR, 'Config');                               // Target sheet ID
  const TKT = getConfigValue(TKR, 'Config');                                     // TKR = Ticket Range

  const ss_tr = SpreadsheetApp.openById(Target_Id);                              // Target spreadsheet
  const sheet_tr = ss_tr.getSheetByName(SheetName);                              // Target sheet
  if (!sheet_tr) {Logger.log(`Target sheet not found: ${SheetName}`); return;}

  let success = false;                                                           // Initialize success flag to false

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
    Logger.log(`Clear EXPORT: ${SheetName} | Didn't find Ticket: ${TKT}`);
  }
}

/////////////////////////////////////////////////////////////////////EXPORT FUNCTIONS/////////////////////////////////////////////////////////////////////