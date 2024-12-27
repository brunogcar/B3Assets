//@NotOnlyCurrentDoc
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
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var SheetNames = ['Index', 'Info', 'Comunicados', 'Prov', 'Preço', 'Cotações', 'OPT' , 'DATA', 'Value', Balanco, Resultado, Fluxo, Valor];
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if(SheetNames.indexOf(sheet.getName()) != -1){
      // Copy the values of the sheet
      var source = sheet.getDataRange();
      source.copyTo(source, {contentsOnly: true});
    }
  }
}

function doDeleteSheets() 
{
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var SheetNames = ['Balanço Ativo', 'Balanço Passivo', 'Demonstração', 'Fluxo de Caixa', 'Demonstração do Valor Adicionado'];
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if(SheetNames.indexOf(sheet.getName()) != -1)
    {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    }
  }
}

function moveSpreadsheetToARQUIVO() 
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var folder = DriveApp.getFoldersByName(`-=ARQUIVO=-`).next();
  var file = DriveApp.getFileById(spreadsheet.getId());
  
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  console.log('Spreadsheet Moved To ARQUIVO');
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
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var folder = DriveApp.getFoldersByName(`-=BACKUP=-`).next();
  var file = DriveApp.getFileById(spreadsheet.getId());
  
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  console.log('Spreadsheet Moved To BACKUP');
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

function doCleanZeros() 
{
  const SheetNames = [SWING_4, SWING_12, SWING_52, OPCOES, BTC, TERMO, FUTURE, FUND];
  
  for (var k = 0; k < SheetNames.length; k++) 
  {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetNames[k]);

    if (!sheet) 
    {
      Logger.log(`Sheet not found: ${SheetNames[k]}`);                               // Log the missing sheet for debugging
      continue;                                                                      // Skip to the next sheet if not found
    }    
    var A5 = sheet.getRange('A5').getValue();

    if (A5 !== "") 
    {
      var lastRow = sheet.getLastRow();
      var lastColumn = sheet.getLastColumn();
      var Data = sheet.getRange(5, 1, lastRow-4, lastColumn).getValues();
    
      for (var i = 0; i < Data.length; i++) 
      {
        for (var j = 0; j < Data[i].length; j++) 
        {
          if (Data[i][j] === 0) 
          {
            Data[i][j] = "";                                                         // Set the value to an empty string
          }
        }
      }
  
      // Now, set the modified values back to the sheet
      sheet.getRange(5, 1, lastRow-4, lastColumn).setValues(Data);
    }
  }
}

/////////////////////////////////////////////////////////////////////reverse/////////////////////////////////////////////////////////////////////

function reverseColumns() 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var LC = sheet.getLastColumn();
  var LR = sheet.getLastRow();
  var range = sheet.getRange(1, 4, LR, LC - 2);                       // Adjust the range to start from column C dynamically
  var values = range.getValues();
  var numRows = values.length;
  var numCols = values[0].length;
  var reversedValues = [];

  for (var i = 0; i < numRows; i++) 
  {
    reversedValues[i] = [];
    for (var j = numCols - 1; j >= 0; j--) 
    {
      reversedValues[i][numCols - 1 - j] = values[i][j];
    }
  }
  range.setValues(reversedValues);
}

function reverseRows() 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var LC = sheet.getLastColumn();
  var LR = sheet.getLastRow();
  var range = sheet.getRange(5, 1, LR - 4, LC);                       // Starting from row 5
  var values = range.getValues();
  var numRows = values.length;
  var numCols = values[0].length;

  for (var i = 0; i < Math.floor(numRows / 2); i++) 
  {
    var temp = values[i];
    values[i] = values[numRows - 1 - i];
    values[numRows - 1 - i] = temp;
  }
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