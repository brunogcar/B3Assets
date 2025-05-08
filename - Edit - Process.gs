/////////////////////////////////////////////////////////////////////PROCESS EDIT/////////////////////////////////////////////////////////////////////

function processEditGeneric(sheet_sr, SheetName, Edit, exportFn) {
  const LC = sheet_sr.getLastColumn();

  const A1 = sheet_sr.getRange("A1").getValue();
  const A2 = sheet_sr.getRange("A2").getValue();
  const A5 = sheet_sr.getRange("A5").getValue();

  if (Edit !== "TRUE") {
    LogDebug(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
    return;
  }

  if (ErrorValues.includes(A2)) {
    LogDebug(`ERROR EDIT: ${SheetName} - ErrorValues in A2 on processEdit`, 'MIN');
    return;
  }

  if (A5 === "" || A2 > A5 || A2 > A1) {
    doSaveBasic(SheetName);
    return;
  }

  if (
    A2 >= A5 || A2 >= A1 ||
    ErrorValues.includes(A1) || ErrorValues.includes(A5)
  ) {
    const condition = (SheetName === FUND);
    const columnCount = condition ? LC : LC - 4;

    const Data_Header = sheet_sr.getRange(2, 1, 1, columnCount).getValues();
    sheet_sr.getRange(5, 1, 1, columnCount).setValues(Data_Header);
    sheet_sr.getRange(1, 1, 1, columnCount).setValues(Data_Header);

    LogDebug(`SUCCESS EDIT. Sheet: ${SheetName}.`, 'MIN');
    exportFn(SheetName);
    return;
  }

  // Final fallback
  LogDebug(`ERROR EDIT: ${SheetName} - Conditions aren't met on processEdit`, 'MIN');
}


/////////////////////////////////////////////////////////////////////PROCESS BASIC AND EXTRA/////////////////////////////////////////////////////////////////////

function processEditBasic(sheet_sr, SheetName, Edit) {
  processEditGeneric(sheet_sr, SheetName, Edit, doExportBasic);
}

function processEditExtra(sheet_sr, SheetName, Edit) {
  processEditGeneric(sheet_sr, SheetName, Edit, doExportExtra);
}

/////////////////////////////////////////////////////////////////////PROCESS FINANCIAL/////////////////////////////////////////////////////////////////////

/**
 * Applies an “edit” sync to financial sheets when the source & template dates match.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet|null} sheet_tr The template sheet (or null if source-only).
 * @param {GoogleAppsScript.Spreadsheet.Sheet}      sheet_sr The source sheet.
 * @param {number|string}                           New_tr    New template date millis or blank.
 * @param {number|string}                           Old_tr    Old template date millis or blank.
 * @param {number|string}                           New_sr    New source date millis or blank.
 * @param {number|string}                           Old_sr    Old source date millis or blank.
 * @param {string}                                  Edit      “TRUE” if EDIT is enabled.
 */
function processEditFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr, Edit) {
  const SheetName = sheet_tr ? sheet_tr.getSheetName() : sheet_sr.getSheetName();
  const LR        = sheet_tr ? sheet_tr.getLastRow()  : sheet_sr.getLastRow();

  let range_sr, range_tr, mappingFunc;

  if (Edit === "TRUE") {
    switch (SheetName) {
      //-------------------------------------------------------------------BLC / DRE / FLC / DVA-------------------------------------------------------------------//
      case BLC:
      case DRE:
      case FLC:
      case DVA:
        if (New_sr.valueOf() > New_tr.valueOf()) {
          doSaveFinancial(SheetName);
          return;
        }
        if (New_sr.valueOf() === New_tr.valueOf()) {
          // For these sheets, the mapping is applied on the target values.
          range_sr = sheet_sr.getRange("B1:B" + LR);
          range_tr = sheet_tr.getRange("B1:B" + LR);
          mappingFunc = (source, target) =>
            target.map((row, index) => [row[0] !== source[index][0] ? source[index][0] : row[0]]);
        }
        break;

      //-------------------------------------------------------------------Balanco-------------------------------------------------------------------//
      case Balanco:
        if (New_sr.valueOf() > Old_sr.valueOf()) {
          doSaveFinancial(SheetName);
          return;
        }
        if (New_sr.valueOf() === Old_sr.valueOf()) {
          range_sr = sheet_sr.getRange("B1:B" + LR);
          range_tr = sheet_sr.getRange("C1:C" + LR);
          mappingFunc = (source, target) =>
            source.map((row, index) => [row[0] !== target[index][0] ? row[0] : target[index][0]]);
        }
        break;

      //-------------------------------------------------------------------Resultado / Valor / Fluxo-------------------------------------------------------------------//
      case Resultado:
      case Valor:
      case Fluxo:
        if (New_sr.valueOf() > Old_sr.valueOf()) {
          doSaveFinancial(SheetName);
          return;
        }
        if (New_sr.valueOf() === Old_sr.valueOf()) {
          range_sr = sheet_sr.getRange("C1:C" + LR);
          range_tr = sheet_sr.getRange("D1:D" + LR);
          mappingFunc = (source, target) =>
            source.map((row, index) => [row[0] !== target[index][0] ? row[0] : target[index][0]]);
        }
        break;

      default:
        LogDebug(`ERROR EDIT: ${SheetName} - Conditions aren’t met on processEditFinancial`, 'MIN');
        return;
    }
  }

  if (Edit !== "TRUE") {
    LogDebug(`ERROR EDIT: ${SheetName} - EDIT on config is set to FALSE`, 'MIN');
  }

  /////////////////////////////////////////////////////////////////////PROCESS - END/////////////////////////////////////////////////////////////////////
  // Common code block: update the values based on the mapping function
  if (range_sr && range_tr && mappingFunc) {
    const values_sr = range_sr.getValues();
    const values_tr = range_tr.getValues();
    const updatedValues = mappingFunc(values_sr, values_tr);
    range_tr.setValues(updatedValues);
    LogDebug(`SUCCESS EDIT. Sheet: ${SheetName}.`, 'MIN');
    doExportFinancial(SheetName);
  }
}

/////////////////////////////////////////////////////////////////////EDIT PROCESS/////////////////////////////////////////////////////////////////////
