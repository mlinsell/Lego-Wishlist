/**
 * Entry point for the search box.
 * Handles both numeric (set number) and text (name) queries.
 */
function sidebarSearch(query) {
  try {
    const trimmed = (query || '').trim();
    if (!trimmed) {
      return { success: false, error: 'Please enter a set number or name' };
    }

    const isNumeric = /^\d+$/.test(trimmed);

    if (isNumeric) {
      const result = rebrickableGetSet(trimmed);
      if (!result.success) {
        return { success: false, error: result.error };
      }
      return {
        success: true,
        directHit: true,
        results: [result.data]
      };
    } else {
      const result = rebrickableSearchSets(trimmed);
      if (!result.success) {
        return { success: false, error: result.error };
      }
      return {
        success: true,
        directHit: false,
        results: result.data
      };
    }
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Fetches Brickset data and checks ownership/wishlist status.
 */
function sidebarGetSetDetail(setId) {
  try {
    const targetId = String(setId);
    
    // 1. Fetch Brickset data (fail gracefully)
    let ukRrp = null;
    let retirementDate = null;
    let subtheme = ""; // Default to empty
    const bricksetResult = bricksetGetSet(targetId);
    
    if (bricksetResult.success && bricksetResult.data) {
      ukRrp = bricksetResult.data.ukRrp;
      retirementDate = bricksetResult.data.retirementDate;
      subtheme = bricksetResult.data.subtheme; // Assign the fetched subtheme
    } else {
      console.log('Brickset fetch failed for ' + targetId + ': ' + (bricksetResult.error || 'Unknown error'));
    }

    // 2. Check if already on wishlist
    const rowNum = findSetRow(targetId);
    const alreadyOnWishlist = (rowNum !== null);

    // 3. Get owned quantity
    let ownedQty = 0;
    const allSets = getAllSets();
    const existingSet = allSets.find(row => String(row.setId) === targetId);
    
    if (existingSet && typeof existingSet.ownedQty === 'number' && existingSet.ownedQty > 0) {
      ownedQty = existingSet.ownedQty;
    }

    return {
      success: true,
      data: {
        setId: targetId,
        ukRrp: ukRrp,
        retirementDate: retirementDate,
        subtheme: subtheme, // Pass it to the UI
        alreadyOnWishlist: alreadyOnWishlist,
        ownedQty: ownedQty
      }
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Writes a new row to the wishlist sheet.
 */
function sidebarAddSet(setData) {
  try {
    // 1. Validate required fields
    const requiredFields = ['setId', 'name', 'theme', 'year', 'pieces', 'imageUrl', 'priority'];
    for (const field of requiredFields) {
      if (setData[field] === undefined || setData[field] === null || setData[field] === '') {
        return { success: false, error: `Missing required field: ${field}` };
      }
    }

    // 2. Validate priority
    const priority = parseInt(setData.priority, 10);
    if (isNaN(priority) || priority < 1 || priority > 5) {
      return { success: false, error: 'Priority must be between 1 and 5' };
    }

    // 3. Call SheetService.addSet (globally available)
    const result = addSet(setData);

    if (!result.success) {
      return { success: false, error: result.error };
    }

    return {
      success: true,
      message: `${setData.setId} · ${setData.name} added ✓`
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}
