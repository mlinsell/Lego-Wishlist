const REBRICKABLE_BASE_URL = 'https://rebrickable.com/api/v3/lego';

/**
 * Returns the standard authentication headers for Rebrickable API calls.
 */
function getRebrickableHeaders() {
  return {
    'Authorization': 'key ' + REBRICKABLE_KEY
  };
}

/**
 * Fetches full details for a single set by set number.
 */
function rebrickableGetSet(setId) {
  try {
    const url = `${REBRICKABLE_BASE_URL}/sets/${setId}-1/`;
    console.log(`Fetching: ${url}`);
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: getRebrickableHeaders()
    });
    
    const responseCode = response.getResponseCode();
    console.log(`Response code: ${responseCode}`);
    
    if (responseCode === 404) {
      return { success: false, error: 'Set not found' };
    }
    if (responseCode !== 200) {
      return { success: false, error: 'Rebrickable error: ' + responseCode };
    }
    
    const data = JSON.parse(response.getContentText());
    const themeName = rebrickableGetThemeName(data.theme_id);
    
    return {
      success: true,
      data: {
        setId: data.set_num.split('-')[0],
        name: data.name,
        theme: themeName,
        subtheme: '', // Rebrickable basic endpoint doesn't return subtheme
        year: data.year,
        pieces: data.num_parts,
        imageUrl: data.set_img_url
      }
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Searches for sets by name keyword.
 */
function rebrickableSearchSets(query) {
  try {
    const encodedQuery = encodeURIComponent(query);
    const url = `${REBRICKABLE_BASE_URL}/sets/?search=${encodedQuery}&page_size=10&ordering=year`;
    console.log(`Fetching: ${url}`);
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: getRebrickableHeaders()
    });
    
    const responseCode = response.getResponseCode();
    console.log(`Response code: ${responseCode}`);
    
    if (responseCode !== 200) {
      return { success: false, error: 'Rebrickable error: ' + responseCode };
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data.results || data.results.length === 0) {
      return { success: true, data: [] };
    }
    
    const results = data.results.map(item => {
      return {
        setId: item.set_num.split('-')[0],
        name: item.name,
        theme: rebrickableGetThemeName(item.theme_id),
        year: item.year,
        pieces: item.num_parts,
        imageUrl: item.set_img_url
      };
    });
    
    return { success: true, data: results };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Resolves a numeric theme ID to a theme name string with caching.
 */
function rebrickableGetThemeName(themeId) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'rebrickable_theme_' + themeId;
    
    const cachedName = cache.get(cacheKey);
    if (cachedName) {
      return cachedName;
    }
    
    const url = `${REBRICKABLE_BASE_URL}/themes/${themeId}/`;
    console.log(`Fetching: ${url}`);
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: getRebrickableHeaders()
    });
    
    const responseCode = response.getResponseCode();
    console.log(`Response code: ${responseCode}`);
    
    if (responseCode !== 200) {
      return 'Unknown';
    }
    
    const data = JSON.parse(response.getContentText());
    const name = data.name;
    
    cache.put(cacheKey, name, THEME_CACHE_DURATION_SECONDS);
    
    return name;
  } catch (error) {
    return 'Unknown';
  }
}
