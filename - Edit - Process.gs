/////////////////////////////////////////////////////////////////////PROCESS EDIT/////////////////////////////////////////////////////////////////////

function processEditGeneric(sheet_sr, SheetName, Edit, exportFn) {

  if (Edit !== "TRUE") {
    LogDebug(`❌ ERROR EDIT: ${SheetName} - EDIT is set to FALSE`, 'MIN');
    return;
  }

  const LC = sheet_sr.getLastColumn();

  // Read rows 1–5 once
  const data = sheet_sr.getRange(1,1,5,LC).getValues();

  const A1 = data[0][0];
  const A2 = data[1][0];
  const A5 = data[4][0];

  if (ErrorValues.includes(A2)) {
    LogDebug(`❌ ERROR EDIT: ${SheetName} - ErrorValues in A2 ${A2} : processEditGeneric`, 'MIN');
    return;
  }

  if (A5 === "" || A2 > A5 || A2 > A1) {
    doSaveBasic(SheetName);
    return;
  }

  if (
    A2 >= A5 || A2 >= A1 ||
    ErrorValues.includes(A1) ||
    ErrorValues.includes(A5)
  ) {

    const condition = (SheetName === FUND);
    const columnCount = condition ? LC : LC - 4;

    const Header = [data[1].slice(0, columnCount)];

    sheet_sr.getRange(5,1,1,columnCount).setValues(Header);
    sheet_sr.getRange(1,1,1,columnCount).setValues(Header);

    LogDebug(`✅ SUCCESS EDIT: ${SheetName}.`, 'MIN');
    exportFn(SheetName);
    return;
  }

  LogDebug(`❌ ERROR EDIT: ${SheetName} - Conditions arent met: processEditGeneric`, 'MIN');
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
function processEditFinancial(sheet_tr, sheet_sr, New_tr, Old_tr, New_sr, Old_sr) {
  const SheetName = sheet_tr.getSheetName();
  const cfg       = Object.values(financialSaveMap)
                            .find(c => c.sh_tr === SheetName);
  if (!cfg) {
    LogDebug(`🚩 No financialSaveMap entry: ${SheetName}`, 'MIN');
    return;
  }

  const LR = sheet_sr.getLastRow();

  let doEdit = false;

  if (New_sr.valueOf() === New_tr.valueOf()) {
    doEdit = true;
  } else {
    LogDebug(`⚠️ EDIT WARNING: Skipping as dates differ (SR:${New_sr} vs TR:${New_tr})`, 'MIN');
    return;
  }

  if (doEdit) {
    const updates = getColumnDifferences(sheet_sr, sheet_tr, cfg.col_src, cfg.col_trg, LR);

    if (updates.length === 0) {
      LogDebug(`🏷️ EDIT not detected: ${SheetName}`, 'MID');
      return;
    }

    // Apply updates one‐by‐one (to preserve blanks/unmodified cells)
    updates.forEach(u => {
      sheet_tr.getRange(u.row, cfg.col_trg).setValue(u.value);
    });
    LogDebug(`✏️ EDIT Applied ${updates.length} edits on ${SheetName} col ${cfg.col_trg}`, 'MIN');

    if (cfg.recurse) {
      doExportFinancial(SheetName);
    }
  }
}

/////////////////////////////////////////////////////////////////////EDIT PROCESS/////////////////////////////////////////////////////////////////////
