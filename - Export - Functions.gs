/////////////////////////////////////////////////////////////////////ID/////////////////////////////////////////////////////////////////////

function getSheetID() {
  const sheet_Id = SpreadsheetApp.getActiveSpreadsheet().getId();
  Logger.log(`Sheet ID not found: ${sheet_Id}`)
  return sheet_Id;
};

function setSheetID() {
  // 1) Pull all your config flags
  const Data_Id = getConfigValue(DIR, 'Config');   // DIR = DATA Source ID
  const TKT     = getConfigValue(TKR, 'Config');   // TKR = Ticket Range
  const EXP     = getConfigValue(EXR, 'Config');   // EXR = Export
  const SHI     = getConfigValue(ICR, 'Config');   // ICR = Sheet ID
  const bgcolor = getConfigValue(IDR, 'Config');   // IDR = ID Sheet
  const colour  = '#d9ead3';

  // 2) Grab your active sheet’s ID
  const Sheet_Id = SpreadsheetApp.getActiveSpreadsheet().getId();

  // 3) Open the external “Relação” sheet
  const ss_tr    = SpreadsheetApp.openById(Data_Id);
  const sheet_tr = ss_tr.getSheetByName('Relação');
  if (!sheet_tr) {
    Logger.log('Target sheet not found: Relação');
    return;
  }
  Logger.log('Found Relação sheet, searching for ticket...');

  // 4) Find your ticket in column A
  const search = sheet_tr
    .getRange("A2:A" + sheet_tr.getLastRow())
    .createTextFinder(TKT)
    .findNext();

  if (!search) {
    Logger.log(`Ticket "${TKT}" not found in Relação sheet`);
    return;
  }
  Logger.log(`Ticket "${TKT}" found at row ${search.getRow()}`);

  // 5) Check your combined condition
  if (bgcolor == colour) {
    Logger.log('bgcolor matches expected colour');
  } else {
    Logger.log(`bgcolor "${bgcolor}" does NOT match "${colour}"`);
  }

  if (EXP == "TRUE" && SHI != "TRUE") {
    Logger.log('Conditions met: EXP is TRUE and SHI is not TRUE — setting Sheet ID');
    search.offset(0, 11).setValue(Sheet_Id);
    search.offset(0, 12).setValue(SNAME(3));
    Logger.log(`Sheet ID ${Sheet_Id} written to Relação!`);
  } else {
    Logger.log(
      `Skipping write: EXP="${EXP}", SHI="${SHI}" — need EXP=="TRUE" && SHI!="TRUE"`
    );
  }
}

function doIsIdExported() {
  const IEP     = getConfigValue(IER, 'Config');                                   // IER = ID Exported?

  if( IEP === "FALSE" )
  {
    setSheetID()
  }
};

function doClearSheetID() {
  const TKT     = getConfigValue(TKR, 'Config');                                   // TKR = Ticket Range
  const Data_Id = getConfigValue(DIR, 'Config');                                   // DIR = DATA Source ID
  if (!Data_Id) { Logger.log("ERROR EXPORT: DATA ID is empty."); return; }

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

function doIsFormula() {
  const Formula  = getConfigValue(FOR, 'Config');                                  // FOR = Formula Range
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

function doIsExportable() {
  const EPD = getConfigValue(EPR, 'Config');                                      // EPR = Exportable? Check Range

  if( EPD === "TRUE" )
  {
    doIsInfoExported()
  }
};

/**
 * Checks the “Exported?” flag in the Config sheet and either
 * 1) If TRUE: copies data from the Info sheet back into Config,
 *    marks the flag TRUE again, and calls setSheetID().
 * 2) If not TRUE: invokes doExportInfo() to perform the export.
 *
 * @returns {void}
 */
function doIsInfoExported() {
  const EXP = getConfigValue(EXR, 'Config');                                      // EXR = Exported?

  if (EXP === "TRUE") {
    const sheet_in = fetchSheetByName('Info');
    if (!sheet_in) return;

    const Range = sheet_in.getRange(TIR).getValues();                              // TIR = Tab Info Range
    sheet_in.getRange(TIR).setValues(Range);                                       // Copy Paste Info

    setConfigValue(EXR, "TRUE");                                                   // Set Formula to TRUE // EXP === "TRUE"

    setSheetID();
  }
  else {
    doExportInfo();
  }
}

/////////////////////////////////////////////////////////////////////CLEAR EXPORTED to EXPORTED Source/////////////////////////////////////////////////////////////////////

function doClearExportAll() {
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUTURE, FUND, BLC, DRE, FLC, DVA];

  _doGroup(SheetNames, doClearExport, "Clearing", "cleared", "");
}

function doClearExport(SheetName) {
  const TKT       = getConfigValue(TKR, 'Config');                                   // TKR = Ticket Range
  const Target_Id = getConfigValue(TDR, 'Config');                                   // Target sheet ID
  if (!Target_Id) { Logger.log("ERROR EXPORT: Target ID is empty."); return; }

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
