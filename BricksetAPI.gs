const BRICKSET_BASE_URL = 'https://brickset.com/api/v3.asmx';

/**
 * Fetches UK RRP, retirement date, and subtheme for a set.
 */
function bricksetGetSet(setId) {
  try {
    const params = { setNumber: setId + '-1', pageSize: 1 };
    const encodedParams = encodeURIComponent(JSON.stringify(params));
    
    // Create a safe URL string for logging (hiding the API key)
    const logUrl = `${BRICKSET_BASE_URL}/getSets?apiKey=HIDDEN_KEY&userHash=&params=${encodedParams}`;
    console.log(`Fetching: ${logUrl}`);
    
    const fetchUrl = `${BRICKSET_BASE_URL}/getSets?apiKey=${BRICKSET_KEY}&userHash=&params=${encodedParams}`;
    
    const response = UrlFetchApp.fetch(fetchUrl, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    console.log(`Response code: ${responseCode}`);
    
    if (responseCode !== 200) {
      return { success: false, error: 'Brickset error: ' + responseCode };
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (data.status !== "success") {
      return { success: false, error: 'Brickset status: ' + data.status };
    }
    
    if (!data.sets || data.sets.length === 0) {
      return { 
        success: true, 
        data: { ukRrp: null, retirementDate: null, subtheme: "" } 
      };
    }
    
    const setObj = data.sets[0];
    let ukRrp = null;
    let retirementDate = null;
    let subtheme = setObj.subtheme || ""; // Extract subtheme from Brickset
    
    // Navigate nested LEGOCom structure safely
    if (setObj.LEGOCom && setObj.LEGOCom.UK) {
      const ukData = setObj.LEGOCom.UK;
      
      if (ukData.retailPrice && ukData.retailPrice !== 0) {
        ukRrp = Math.round(ukData.retailPrice * 100) / 100;
      }
      
      if (ukData.dateLastAvailable && ukData.dateLastAvailable !== "") {
        // Convert "YYYY-MM-DDTHH:MM:SS" into "YYYY-MM-DD"
        retirementDate = new Date(ukData.dateLastAvailable).toISOString().split('T')[0];
      }
    }
    
    return {
      success: true,
      data: {
        ukRrp: ukRrp,
        retirementDate: retirementDate,
        subtheme: subtheme // Return it here
      }
    };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}
