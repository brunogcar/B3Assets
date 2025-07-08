//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////MENU FUNCTIONS/////////////////////////////////////////////////////////////////////

function import_Upgrade()
{
  Logger.log('Import: import_16x_to_17x');

  import_config();

  doImportProventos();

  doImportShares_Upgrade();

//  import_Upgrade_Sheets();
  doImportBasics();
  doImportFinancials();

  deleteRowByDisplayDate_OPCOES();

  doCheckTriggers();
  update_form();

// doCleanZeros();
};

function import_Upgrade_Sheets()
{
  doImportBasic(SWING_4);
  doImportBasic(SWING_12);
  doImportBasic(SWING_52);
  doImportBasic(OPCOES);
  doImportBasic(BTC);
  doImportBasic(TERMO);
  doImportBasic(FUND);
  doImportBasic(AFTER);

  doImportBasic(FUTURE);
  doImportBasic(FUTURE_1);
  doImportBasic(FUTURE_2);
  doImportBasic(FUTURE_3);

  doImportBasic(RIGHT_1);
  doImportBasic(RIGHT_2);
  doImportBasic(RECEIPT_9);
  doImportBasic(RECEIPT_10);
  doImportBasic(WARRANT_11);
  doImportBasic(WARRANT_12);
  doImportBasic(WARRANT_13);
  doImportBasic(BLOCK);

  doImportFinancial(BLC);
  doImportFinancial(Balanco);
  doImportFinancial(DRE);
  doImportFinancial(Resultado);
  doImportFinancial(FLC);
  doImportFinancial(Fluxo);
  doImportFinancial(DVA);
  doImportFinancial(Valor);
}

/////////////////////////////////////////////////////////////////////IMPORT FUNCTIONS/////////////////////////////////////////////////////////////////////

function doImportShares_Upgrade() {
  const Source_Id = getConfigValue(SIR, 'Config');                                    // SIR = Source ID

  const sheet_sr = SpreadsheetApp.openById(Source_Id).getSheetByName('DATA');         // Source Sheet
  const L1 = sheet_sr.getRange("L1").getValue();
  const L2 = sheet_sr.getRange("L2").getValue();
  const sheet_tr = fetchSheetByName('DATA');                                          // Target Sheet
  if (!sheet_tr) return;

    var SheetName = sheet_tr.getName()

  LogDebug(`IMPORT: Shares and FF`, 'MIN');

  if (!ErrorValues.includes(L1) && !ErrorValues.includes(L2)) {
    // read the block L1:L2, then write it into L5:L6
    const Data = sheet_sr.getRange("L1:L2").getValues();
    sheet_tr.getRange("L5:L6").setValues(Data);
  }
  else
  {
    LogDebug(`❌ ERROR IMPORT: ${SheetName} - ErrorValues - L1=${L1} or L2=${L2}: doImportShares`, 'MIN');
  }
  LogDebug(`✅ SUCCESS IMPORT: Shares and FF`, 'MIN');
}

function deleteRowByDisplayDate_OPCOES() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName('Opções');
  if (!sheet) throw new Error('Sheet “Opções” not found');

  const targetDisplay = '30/05/2025';
  LogDebug(`🔍 Looking for display-value ${targetDisplay} in Opções!A5:A`, 'MIN');

  const LR = sheet.getLastRow();
  if (LR < 5) {
    LogDebug(`⚠️ No data rows (LR=${LR}); nothing to delete.`, 'MID');
    return;
  }

  // Grab all displayed values in A5:A
  const displays = sheet
    .getRange(5, 1, LR - 4, 1)
    .getDisplayValues()
    .flat();

  let deletedCount = 0;
  // Loop backwards so row numbers stay valid as we delete
  for (let i = displays.length - 1; i >= 0; i--) {
    if (displays[i] === targetDisplay) {
      const rowNum = i + 5;  // because array starts at A5
      sheet.deleteRow(rowNum);
      deletedCount++;
      LogDebug(`✅ Deleted row ${rowNum} (display="${targetDisplay}")`, 'MIN');
    }
  }
  if (deletedCount === 0) {
    LogDebug(`❌ No matches for ${targetDisplay}; nothing deleted.`, 'MIN');
  }
  LogDebug(`🗑️ Finished deletions: ${deletedCount} row(s) removed.`, 'MID');
}

/////////////////////////////////////////////////////////////////////IMPORT UPGRADE/////////////////////////////////////////////////////////////////////
