//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////Helper functions/////////////////////////////////////////////////////////////////////

function fetchSheetByName(SheetName) 
{
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
  if (!sheet) {
    Logger.log(`Sheet not found: ${SheetName}`);
    return null;
  }
  return sheet;
}

function getConfigValue(Acronym) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet_se = ss.getSheetByName('Settings');
  const sheet_co = ss.getSheetByName('Config');
  
  if (!sheet_se || !sheet_co) {
    Logger.log('Settings or Config sheet not found');
    return null;
  }

  // Get value from Settings
  let Value = sheet_se.getRange(Acronym).getDisplayValue().trim();

  // Fallback to Config if value is DEFAULT or in ErrorValues
  if (Value === "DEFAULT" || ErrorValues.includes(Value)) 
  {
    Value = sheet_co.getRange(Acronym).getDisplayValue().trim();
    // Verify Config value isn't also invalid
    return ErrorValues.includes(Value) ? null : Value;
  }
  return Value;
}

/////////////////////////////////////////////////////////////////////Settings/////////////////////////////////////////////////////////////////////

function doSettings()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet_co = ss.getSheetByName('Config');
  var Class     = sheet_co.getRange(IST).getDisplayValue();                             // IST = Is Stock? 
  const sheet_sr = ss.getSheetByName('Settings');
  var Activate  = sheet_sr.getRange(ACT).getDisplayValue();                             // ACT = Activate

  if (Class == 'STOCK') 
  {
    if ( Activate == "TRUE")                                              // TRUE
    {
      var True = sheet_sr.getRange(TRU).getDisplayValue();                               // TRU = True 

      if ( True == 'SAVE')                                                // SAVE
      {
        var Save = sheet_sr.getRange(SAV).getDisplayValue();                             // SAV = SAVE

        if ( Save == 'SHEETS') { doSaveAllSheets(); }
        if ( Save == 'EXTRAS') { doSaveAllExtras(); }
        if ( Save == 'DATAS')  { doSaveAllDatas(); }
        if ( Save == 'ALL')    { doSaveAll(); }
        if ( Save == 'INDIVIDUAL')
          { 
            var Individual = sheet_sr.getRange(IND).getDisplayValue();                   // IND = INDIVIDUAL

            if ( Individual == 'SWING')  { doSaveSWING(); }
            if ( Individual == 'OPCOES') { doSaveSheet(OPCOES); }
            if ( Individual == 'BTC')    { doSaveSheet(BTC); }
            if ( Individual == 'TERMO')  { doSaveSheet(TERMO); }
            if ( Individual == 'FUND')   { doSaveSheet(FUND); }
            if ( Individual == 'FUTURE') { doSaveSheet(FUTURE); }
          }
      }
      if ( True == 'EXPORT') {doExportAll(); }
      if ( True == 'OTHER')                                               // OTHER
      {
        var Other = sheet_sr.getRange(EXT).getDisplayValue();                             // EXT = Extra

        if ( Other == 'ZEROS')    { doCleanZeros(); }
        if ( Other == 'TRIGGERS') { doCheckTriggers(); }
        if ( Other == 'CHECK')    { doCheckDATAS(); }                                    // Check and hide or show Sheets
        if ( Other == 'PROV')     { doSaveProventos(); }
        if ( Other == 'SHARES')   { doSaveShares(); }
        if ( Other == 'RIGHTS')   { doRestoreRight(); }
        
      }
    }
  }
};

/////////////////////////////////////////////////////////////////////RETIRE/////////////////////////////////////////////////////////////////////

function doRetire() 
{
  copypasteSheets();
  doClearSheetID();
  clearExportALL();
  doDeleteSheets();
  moveSpreadsheetToARQUIVO();

  doDeleteTriggers();
  revokeOwnAccess();
};

function copypasteSheets() 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const SheetNames = new Set(['Index', 'Info', 'Comunicados', 'Prov', 'Preço', 'Cotações', 'OPT', 'DATA', 'Value', Balanco, Resultado, Fluxo, Valor]);

  sheets.forEach(sheet => {
    if (SheetNames.has(sheet.getName())) 
    {
      sheet.getDataRange().copyTo(sheet.getDataRange(), { contentsOnly: true });
    }
  });
}

function doDeleteSheets() 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SheetNames = new Set(['Balanço Ativo', 'Balanço Passivo', 'Demonstração', 'Fluxo de Caixa', 'Demonstração do Valor Adicionado']);
  
  ss.getSheets().forEach(sheet => {
    if (SheetNames.has(sheet.getName())) 
    {
      ss.deleteSheet(sheet);
    }
  });
}

function moveSpreadsheetToFolder(folderName) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());

  const folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) {
    Logger.log(`Folder "${folderName}" not found.`);
    return;
  }

  const folder = folders.next();
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  Logger.log(`Spreadsheet moved to ${folderName}`);
}

function moveSpreadsheetToARQUIVO() 
{
  moveSpreadsheetToFolder("-=ARQUIVO=-");
}

/////////////////////////////////////////////////////////////////////DELETE/////////////////////////////////////////////////////////////////////

function doDelete() 
{
  doDeleteTriggers();
  moveSpreadsheetToBACKUP();
  revokeOwnAccess();
}

function revokeOwnAccess() 
{
  // Invalidate the script's authorization
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  if (authInfo) {
    ScriptApp.invalidateAuth();
    Logger.log('Script access revoked successfully.');
  } else {
    Logger.log('Script is not authorized or access has already been revoked.');
  }
}

function moveSpreadsheetToBACKUP() 
{
  moveSpreadsheetToFolder("-=BACKUP=-");
}

function doDeleteSpreadsheet() 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fileId = ss.getId();
  
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    Logger.log('Spreadsheet deleted successfully.');
  } catch (error) {
    Logger.log('Error deleting spreadsheet: ' + error);
  }
}

/////////////////////////////////////////////////////////////////////Name/////////////////////////////////////////////////////////////////////

function SNAME(option) 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var thisSheet = sheet.getName(); 

  if(option === 0)                               // ACTIVE SHEET NAME =SNAME(0)
  {
    return thisSheet;
  }
  else if(option === 1)                          // ALL SHEET NAMES =SNAME(1)
  {
    var sheetList = [];
    ss.getSheets().forEach(function(val){
       sheetList.push(val.getName())
    });
    return sheetList;
  }
  else if(option === 2)                         // SPREADSHEET NAME =SNAME(2)
  {
    return ss.getName(); 
  }
  else if(option === 3)                         // SPREADSHEET VERSION  =SNAME(3)
  {
    var SheetName = ss.getName();
    var regex = /-(.*)/;
    var matches = regex.exec(SheetName);
      if (matches) 
      {
        return matches[1].trim();
      }
      else 
      {
        return "No match found";
      }
  }
  else
  {
    return "#N/A";
  }
};

/////////////////////////////////////////////////////////////////////CLEAN SHEETS/////////////////////////////////////////////////////////////////////

function doCleanZeros() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUTURE, FUND];

  SheetNames.forEach(SheetName => {
    const sheet = ss.getSheetByName(SheetName);
    if (!sheet) {
      Logger.log(`Sheet not found: ${SheetName}`);
      return;
    }

    const range = sheet.getRange(5, 1, sheet.getLastRow() - 4, sheet.getLastColumn());
    const Data = range.getValues();
    let Modified = false;

    Data.forEach(row => {
      row.forEach((cell, i) => {
        if (cell === 0) {
          row[i] = "";
          Modified = true;
        }
      });
    });

    if (Modified) range.setValues(Data); // Only update if changes were made
  });
}

/////////////////////////////////////////////////////////////////////reverse/////////////////////////////////////////////////////////////////////

function reverseColumns() 
{
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(1, 4, sheet.getLastRow(), sheet.getLastColumn() - 2);
  const values = range.getValues();

  const reversedValues = values.map(row => row.reverse());
  range.setValues(reversedValues);
}

function reverseRows() 
{
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(5, 1, sheet.getLastRow() - 4, sheet.getLastColumn());
  const values = range.getValues();

  values.reverse();
  range.setValues(values);
}

/////////////////////////////////////////////////////////////////////RESTORE Functions/////////////////////////////////////////////////////////////////////

function doRestoreFundExport()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();                        // Source spreadsheet
  const sheet_co = ss.getSheetByName('Config');                             // Source sheet

  var Value = '=IF(OR(AND(Fund!A5="";Fund!A1=""); L18<>"STOCK"); FALSE;TRUE)';                              

    sheet_co.getRange(EFU).setValue(Value);                                 // EFU = Export to Fund 
}

/////////////////////////////////////////////////////////////////////FUNCTIONS/////////////////////////////////////////////////////////////////////