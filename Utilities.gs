/**
 * Scans the Wishlist sheet for any missing data and fills it 
 * using the Rebrickable and Brickset APIs safely.
 */
function fillMissingSetData() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < DATA_START_ROW) {
    ui.alert('No sets found on the Wishlist.');
    return;
  }

  // Notify user that the script has started
  SpreadsheetApp.getActiveSpreadsheet().toast('Scanning wishlist and querying APIs. This may take a minute...', 'Scan Started', -1);

  // Read only up to Column 15 (IMAGE_URL) to avoid the IMAGE() formula crashing the service
  const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 15);
  const values = dataRange.getValues();
  
  let updatedCount = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowNum = i + DATA_START_ROW;
    const setId = String(row[COL.SET_ID - 1]);
    
    if (!setId) continue; // Skip blank rows

    let needsUpdate = false;

    // 1. Check Rebrickable fields
    const missingReb = !row[COL.SET_NAME - 1] || !row[COL.THEME - 1] || 
                       !row[COL.YEAR - 1] || !row[COL.PIECES - 1] || !row[COL.IMAGE_URL - 1];
                       
    if (missingReb) {
      const res = rebrickableGetSet(setId);
      if (res.success && res.data) {
        if (!row[COL.SET_NAME - 1] && res.data.name) sheet.getRange(rowNum, COL.SET_NAME).setValue(res.data.name);
        if (!row[COL.THEME - 1] && res.data.theme) sheet.getRange(rowNum, COL.THEME).setValue(res.data.theme);
        if (!row[COL.YEAR - 1] && res.data.year) sheet.getRange(rowNum, COL.YEAR).setValue(res.data.year);
        if (!row[COL.PIECES - 1] && res.data.pieces) sheet.getRange(rowNum, COL.PIECES).setValue(res.data.pieces);
        if (!row[COL.IMAGE_URL - 1] && res.data.imageUrl) sheet.getRange(rowNum, COL.IMAGE_URL).setValue(res.data.imageUrl);
        needsUpdate = true;
      }
    }

    // 2. Check Brickset fields
    const missingBrick = !row[COL.SUBTHEME - 1] || row[COL.UK_RRP - 1] === '' || 
                         row[COL.RETIREMENT_DATE - 1] === '';
                         
    if (missingBrick) {
      const res = bricksetGetSet(setId);
      if (res.success && res.data) {
        
        if (!row[COL.SUBTHEME - 1] && res.data.subtheme) {
          sheet.getRange(rowNum, COL.SUBTHEME).setValue(res.data.subtheme);
          needsUpdate = true;
        }
        
        if (row[COL.UK_RRP - 1] === '' && res.data.ukRrp !== null) {
          sheet.getRange(rowNum, COL.UK_RRP).setValue(res.data.ukRrp);
          
          // Auto-calculate target price ONLY if it is currently blank
          if (row[COL.TARGET_PRICE - 1] === '') {
            const target = Math.round(res.data.ukRrp * TARGET_PRICE_MULTIPLIER * 100) / 100;
            sheet.getRange(rowNum, COL.TARGET_PRICE).setValue(target);
          }
          needsUpdate = true;
        }
        
        if (row[COL.RETIREMENT_DATE - 1] === '' && res.data.retirementDate !== null) {
          sheet.getRange(rowNum, COL.RETIREMENT_DATE).setValue(res.data.retirementDate);
          needsUpdate = true;
        }
      }
    }

    if (needsUpdate) {
      updatedCount++;
      // Sleep for 500ms between updates to avoid hitting API rate limits
      Utilities.sleep(500); 
    }
  }

  if (updatedCount > 0) {
    ui.alert(`Update Complete\n\nSuccessfully filled missing data for ${updatedCount} set(s).`);
  } else {
    ui.alert('Scan Complete\n\nNo missing data was found. Your wishlist is fully populated!');
  }
}
