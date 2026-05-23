/**
 * Returns the full wishlist dataset to the web app on initial load.
 */
function webAppGetAllSets() {
  try {
    return getAllSets(); // Calls public function from SheetService.gs
  } catch (error) {
    return { error: error.message };
  }
}

/**
 * Updates a single editable field for a set.
 */
function webAppUpdateField(setId, colName, value) {
  try {
    const permittedCols = ['targetPrice', 'priority', 'reason'];
    if (!permittedCols.includes(colName)) {
      return { success: false, error: 'Invalid field: ' + colName };
    }

    let coercedValue = value;

    if (colName === 'targetPrice') {
      coercedValue = parseFloat(value);
      if (isNaN(coercedValue)) {
        return { success: false, error: 'Invalid price value' };
      }
    } else if (colName === 'priority') {
      coercedValue = parseInt(value, 10);
      if (isNaN(coercedValue) || coercedValue < 1 || coercedValue > 5) {
        return { success: false, error: 'Priority must be 1–5' };
      }
    } else if (colName === 'reason') {
      coercedValue = String(value).trim();
    }

    return updateCell(setId, colName, coercedValue);
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Deletes a set from the Wishlist sheet by Set ID.
 */
function webAppDeleteSet(setId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const row = findSetRow(setId); // Uses the existing helper from SheetService.gs
    
    if (row === null) {
      return { success: false, error: 'Set not found in sheet.' };
    }
    
    sheet.deleteRow(row);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}
