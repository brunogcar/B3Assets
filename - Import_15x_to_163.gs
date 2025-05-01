//@NotOnlyCurrentDoc
/////////////////////////////////////////////////////////////////////MENU FUNCTIONS/////////////////////////////////////////////////////////////////////

function import_15x_to_163()
{
  Logger.log('Import: import_15x_to_163');

  import_config();

  doImportProventos();

  doImportShares();

  doImportFinancial(BLC);
  doImportFinancial(Balanco);
  doImportFinancial(DRE);
  doImportFinancial(Resultado);
  doImportFinancial(FLC);
  doImportFinancial(Fluxo);
  doImportFinancial(DVA);
  doImportFinancial(Valor);

//  doImportBasic(SWING_4);
  doImport_SWING_12_to_SWING_4();
  doImportBasic(SWING_12);
  doImportBasic(SWING_52);
  doImportBasic(OPCOES);
  doImportBasic(BTC);
  doImportBasic(TERMO);
  doImportBasic(FUND);

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

  doCheckTriggers();
  update_form();

// doCleanZeros();
};

/////////////////////////////////////////////////////////////////////IMPORT FUNCTIONS/////////////////////////////////////////////////////////////////////



/////////////////////////////////////////////////////////////////////IMPORT AND COPY SHEET/////////////////////////////////////////////////////////////////////

function doImport_SWING_12_to_SWING_4()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_co = fetchSheetByName('Config');

  const Source_Id = sheet_co.getRange(SIR).getDisplayValue().trim();                     // Get Source ID from Config
  const ss_sr = SpreadsheetApp.openById(Source_Id);
  const sheet_sr = ss_sr.getSheetByName(SWING_12);                            // Source Sheet (Trade)

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

/////////////////////////////////////////////////////////////////////IMPORT UPGRADE/////////////////////////////////////////////////////////////////////