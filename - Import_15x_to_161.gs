//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////MENU FUNCTIONS/////////////////////////////////////////////////////////////////////

function import_15x_to_161()
{
  Logger.log('Import: import_15x_to_161');

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

//  doImportSheet(SWING_4);
  doImport_SWING_12_to_SWING_4();
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

/////////////////////////////////////////////////////////////////////IMPORT FUNCTIONS/////////////////////////////////////////////////////////////////////



/////////////////////////////////////////////////////////////////////IMPORT AND COPY SHEET/////////////////////////////////////////////////////////////////////

function doImport_SWING_12_to_SWING_4()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = ss.getSheetByName('Config');

  const Source_Id = sheet_co.getRange(SIR).getDisplayValue().trim();                     // Get Source ID from Config
  const sourceSpreadsheet = SpreadsheetApp.openById(Source_Id);
  const sheet_sr = sourceSpreadsheet.getSheetByName(SWING_12);                            // Source Sheet (Trade)

  Logger.log('Import: Trade to Swing Sheets');

  // Check if the source sheet exists
  if (!sheet_sr) {
    Logger.log('ERROR IMPORT: Trade sheet does not exist in the source spreadsheet.');
    return;
  }

  // Split handling for SWING_4 and (SWING_12, SWING_52)
  const targetSheets = [
    {
      name: SWING_4,
      sourceRanges: ['A5:D125', 'Y5:Y125', 'X5:X125', 'S5:V125'],                                             // Add source ranges specific to SWING_4
      targetRanges: ['A5:D125', 'E5:E125', 'F5:F125', 'V5:Y125']                                              // Corresponding target ranges
    }
  ];

  targetSheets.forEach(config => {
    let targetSheet = ss.getSheetByName(config.name);

    if (!targetSheet) {
      Logger.log(`${config.name} sheet does not exist.`);
    }

    // Loop through source and target ranges
    config.sourceRanges.forEach((sourceRange, index) => {
      const targetRange = config.targetRanges[index];

      // Fetch data from source range
      const data = sheet_sr.getRange(sourceRange).getValues();

      // Clear and set data in the target range
      targetSheet.getRange(targetRange).clearContent();
      targetSheet.getRange(targetRange).setValues(data);

      Logger.log(`SUCCESS IMPORT. Data from 'Trade' (${sourceRange}) copied to '${config.name}' (${targetRange}).`);
    });
  });
}

/////////////////////////////////////////////////////////////////////IMPORT/////////////////////////////////////////////////////////////////////