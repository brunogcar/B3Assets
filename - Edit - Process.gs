/////////////////////////////////////////////////////////////////////PROCESS EDIT/////////////////////////////////////////////////////////////////////

function processEditGeneric(sheet_sr, SheetName, Edit, exportFn) {
  const LC = sheet_sr.getLastColumn();

  const A1 = sheet_sr.getRange("A1").getValue();
  const A2 = sheet_sr.getRange("A2").getValue();
  const A5 = sheet_sr.getRange("A5").getValue();

  if (Edit !== "TRUE") {
    LogDebug(`ERROR EDIT: ${SheetName} - EDIT is set to FALSE`, 'MIN');
    return;
  }

  if (ErrorValues.includes(A2)) {
    LogDebug(`ERROR EDIT: ${SheetName} - ErrorValues in A2: processEditGeneric`, 'MIN');
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
  LogDebug(`ERROR EDIT: ${SheetName} - Conditions aren't met: processEditGeneric`, 'MIN');
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
  const cfg       = Object.values(financialMap)
                            .find(c => c.sh_tr === SheetName);
  if (!cfg) {
    LogDebug(`No financialMap entry: ${SheetName}`, 'MIN');
    return;
  }

  const LR = sheet_sr.getLastRow();

  let doEdit = false;

  if (New_sr.valueOf() === New_tr.valueOf()) {
    doEdit = true;
  } else {
    LogDebug(`Skipping edit: dates differ (SR:${New_sr} vs TR:${New_tr})`, 'MIN');
    return;
  }

  if (doEdit) {
    const edit_sr = sheet_sr.getRange(1, cfg.col_src, LR, 1);
    const edit_tr = sheet_tr.getRange(1, cfg.col_trg, LR, 1);

    const values_sr = edit_sr.getValues();
    const values_tr = edit_tr.getValues();

    const updates = [];
    values_sr.forEach((row, i) => {
      const vSr = row[0], vTr = values_tr[i][0];
      if (vSr !== vTr) {
        updates.push({ row: i+1, value: vSr });
      }
    });

    if (updates.length === 0) {
      LogDebug(`No edits detected: ${SheetName}`, 'MID');
      return;
    }

    // Apply updates one‐by‐one (to preserve blanks/unmodified cells)
    updates.forEach(u => {
      sheet_tr.getRange(u.row, cfg.col_trg).setValue(u.value);
    });
    LogDebug(`Applied ${updates.length} edits on ${SheetName} col ${cfg.col_trg}`, 'MIN');
    if (cfg.recurse) {
      doExportFinancial(SheetName);
    }
  }
}

/////////////////////////////////////////////////////////////////////EDIT PROCESS/////////////////////////////////////////////////////////////////////
