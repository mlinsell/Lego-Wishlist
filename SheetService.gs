/**
 * Private helper to get the Wishlist sheet.
 */
function getWishlistSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(`Sheet "${SHEET_NAME}" not found.`);
  }
  return sheet;
}

/**
 * Returns all wishlist rows as an array of plain objects.
 */
function getAllSets() {
  const sheet = getWishlistSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < DATA_START_ROW) {
    return [];
  }
  
  const numRows = lastRow - DATA_START_ROW + 1;
  const dataRange = sheet.getRange(DATA_START_ROW, 1, numRows, TOTAL_COLUMNS);
  const values = dataRange.getValues();
  
  return values.map(row => ({
    setId: row[COL.SET_ID - 1] !== '' ? String(row[COL.SET_ID - 1]) : '',
    setName: row[COL.SET_NAME - 1],
    theme: row[COL.THEME - 1],
    subtheme: row[COL.SUBTHEME - 1],
    year: row[COL.YEAR - 1],
    pieces: row[COL.PIECES - 1],
    ukRrp: row[COL.UK_RRP - 1],
    targetPrice: row[COL.TARGET_PRICE - 1],
    retirementDate: row[COL.RETIREMENT_DATE - 1],
    priority: row[COL.PRIORITY - 1],
    reason: row[COL.REASON - 1],
    ownedQty: row[COL.OWNED_QTY - 1],
    ownedStatus: row[COL.OWNED_STATUS - 1],
    collectionNotes: row[COL.COLLECTION_NOTES - 1],
    imageUrl: row[COL.IMAGE_URL - 1],
    imagePreview: row[COL.IMAGE_PREVIEW - 1]
  }));
}

/**
 * Searches column COL.SET_ID for an exact string match.
 * Returns the 1-based row number of the first match, or null if not found.
 */
function findSetRow(setId) {
  const sheet = getWishlistSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < DATA_START_ROW) return null;
  
  const numRows = lastRow - DATA_START_ROW + 1;
  const idValues = sheet.getRange(DATA_START_ROW, COL.SET_ID, numRows, 1).getValues();
  const targetId = String(setId);
  
  for (let i = 0; i < idValues.length; i++) {
    if (String(idValues[i][0]) === targetId) {
      return i + DATA_START_ROW;
    }
  }
  
  return null;
}

/**
 * Adds a new set to the wishlist sheet.
 */
function addSet(setData) {
  try {
    const sheet = getWishlistSheet();
    const row = sheet.getLastRow() + 1;
    
    // Calculate the target price
    let targetPriceVal = "";
    if (typeof setData.targetPrice === 'number') {
      targetPriceVal = setData.targetPrice;
    } else if (typeof setData.ukRrp === 'number') {
      targetPriceVal = Math.round(setData.ukRrp * TARGET_PRICE_MULTIPLIER * 100) / 100;
    }
    
    // Build values array (1-16)
    const values = new Array(TOTAL_COLUMNS).fill("");
    values[COL.SET_ID - 1] = setData.setId != null ? String(setData.setId) : "";
    values[COL.SET_NAME - 1] = setData.name != null ? setData.name : "";
    values[COL.THEME - 1] = setData.theme != null ? setData.theme : "";
    values[COL.SUBTHEME - 1] = setData.subtheme != null ? setData.subtheme : "";
    values[COL.YEAR - 1] = setData.year != null ? setData.year : "";
    values[COL.PIECES - 1] = setData.pieces != null ? setData.pieces : "";
    values[COL.UK_RRP - 1] = setData.ukRrp != null ? setData.ukRrp : "";
    values[COL.TARGET_PRICE - 1] = targetPriceVal;
    values[COL.RETIREMENT_DATE - 1] = setData.retirementDate != null ? setData.retirementDate : "";
    values[COL.PRIORITY - 1] = setData.priority != null ? setData.priority : "";
    values[COL.REASON - 1] = setData.reason != null ? setData.reason : "";
    values[COL.IMAGE_URL - 1] = setData.imageUrl != null ? setData.imageUrl : "";
    
    // Write standard values first
    sheet.getRange(row, 1, 1, TOTAL_COLUMNS).setValues([values]);
    
    // Note: IMPORTRANGE requires a one-time manual authorisation in the sheet UI. 
    // The first time these formulas are written, cells 12-14 will show a #REF! error 
    // until the user clicks "Allow access" in the sheet.
    
    sheet.getRange(row, COL.OWNED_QTY)
      .setFormula(`=COUNTIF(IMPORTRANGE("${COLLECTION_SHEET_ID}","${COLLECTION_SHEET_TAB}!A:A"),A${row})`);
      
    sheet.getRange(row, COL.OWNED_STATUS)
      .setFormula(`=IFERROR(VLOOKUP(A${row},IMPORTRANGE("${COLLECTION_SHEET_ID}","${COLLECTION_SHEET_TAB}!A:B"),2,FALSE),"")`);
      
    sheet.getRange(row, COL.COLLECTION_NOTES)
      .setFormula(`=IFERROR(VLOOKUP(A${row},IMPORTRANGE("${COLLECTION_SHEET_ID}","${COLLECTION_SHEET_TAB}!A:C"),3,FALSE),"")`);
      
    sheet.getRange(row, COL.IMAGE_PREVIEW)
      .setFormula(`=IMAGE(O${row})`);
      
    return { success: true, rowNum: row };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Updates a single permitted cell for a given set ID.
 */
function updateCell(setId, colName, value) {
  try {
    const permittedCols = {
      'targetPrice': COL.TARGET_PRICE,
      'priority': COL.PRIORITY,
      'reason': COL.REASON
    };
    
    if (!permittedCols.hasOwnProperty(colName)) {
      return { success: false, error: 'Invalid field: ' + colName };
    }
    
    const colIndex = permittedCols[colName];
    const row = findSetRow(setId);
    
    if (row === null) {
      return { success: false, error: 'Set not found: ' + setId };
    }
    
    const sheet = getWishlistSheet();
    sheet.getRange(row, colIndex).setValue(value);
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}
