/**
 * Code.gs — FINAL with chunked cache + progress + log window
 *
 * Main sheet: Wishlist App Data
 * Columns:
 * A Set_ID
 * B Set Name
 * C Theme
 * D Year
 * E UK RRP
 * F Target Price
 * G Wishlists Containing Set
 * H Pieces
 * I Image URL
 * J Image Preview
 *
 * Additional sheets:
 * - SetsCache
 * - WishlistSources
 * - CacheMeta
 */

/* -------------------------
   Configuration / Columns
   ------------------------- */
const SHEET_NAME = 'Wishlist App Data';
const COL = {
  SET_ID: 1,
  SET_NAME: 2,
  THEME: 3,
  YEAR: 4,
  UK_RRP: 5,
  TARGET: 6,
  WISHLISTS: 7,
  PIECES: 8,
  IMAGE_URL: 9,
  IMAGE_PREVIEW: 10
};

/* -------------------------
   Progress / Log globals
   ------------------------- */
// Note: globals persist only while the Apps Script instance is warm.
// They are used to coordinate chunked work across repeated client calls.
var CACHE_PROGRESS = {
  year: null,
  page: 1,
  pageSize: 100,
  done: false,
  lastMessage: '',
  log: []
};

/* -------------------------
   Menu
   ------------------------- */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Lego Wishlist')
    .addItem('Search Sets', 'openSetSearchSidebar')
    .addItem('Sync LEGO Wishlists', 'syncAllWishlists')
    .addItem('Refresh All', 'refreshAll')
    .addItem('Refresh Selected Row', 'refreshSelectedRow')
    .addToUi();
}

/* -------------------------
   Helpers
   ------------------------- */
function normaliseSetId(input) {
  const clean = String(input || '').trim().replace(/-1$/, '');
  return clean ? (clean + '-1') : '';
}

/* -------------------------
   Refresh logic
   ------------------------- */
function refreshAll() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const last = sheet.getLastRow();
  for (let r = 2; r <= last; r++) {
    const id = sheet.getRange(r, COL.SET_ID).getValue();
    if (id) updateRow(sheet, r, id);
  }
}

function refreshSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  if (row < 2) return;
  const id = sheet.getRange(row, COL.SET_ID).getValue();
  if (id) updateRow(sheet, row, id);
}

/* -------------------------
   Update a single row
   ------------------------- */
function updateRow(sheet, row, setIdRaw) {
  const setId = normaliseSetId(setIdRaw);
  if (!setId) return;

  const data = fetchRebrickableSet(setId);
  if (!data) return;

  sheet.getRange(row, COL.SET_NAME).setValue(data.name || '');
  sheet.getRange(row, COL.THEME).setValue(fetchRebrickableTheme(data.theme_id));
  sheet.getRange(row, COL.YEAR).setValue(data.year || '');
  sheet.getRange(row, COL.PIECES).setValue(data.num_parts || '');

  if (data.set_img_url) {
    sheet.getRange(row, COL.IMAGE_URL).setValue(data.set_img_url);
    try { sheet.setColumnWidth(COL.IMAGE_PREVIEW, 150); sheet.setRowHeight(row, 150); } catch(e) {}
    sheet.getRange(row, COL.IMAGE_PREVIEW).setFormula(`=IMAGE("${data.set_img_url}",1)`);
  }

  const rrp = fetchBricksetRRP(setId);
  if (rrp !== null && rrp !== '') {
    const numeric = parseFloat(String(rrp).replace(/[^\d\.]/g, ''));
    if (!isNaN(numeric)) {
      sheet.getRange(row, COL.UK_RRP).setValue(numeric);
      sheet.getRange(row, COL.UK_RRP).setNumberFormat("£#,##0.00");
      const targetCell = sheet.getRange(row, COL.TARGET);
      if (!targetCell.getValue()) {
        const targetValue = Math.round((numeric * 0.75) * 100) / 100;
        targetCell.setValue(targetValue);
        targetCell.setNumberFormat("£#,##0.00");
      }
    } else {
      sheet.getRange(row, COL.UK_RRP).setValue(rrp);
    }
  }
}

/* -------------------------
   Rebrickable helpers
   ------------------------- */
function fetchRebrickableSet(setId) {
  const key = Config.rebrickableKey();
  const url = `https://rebrickable.com/api/v3/lego/sets/${encodeURIComponent(setId)}/?key=${encodeURIComponent(key)}`;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return null;
  return JSON.parse(resp.getContentText());
}

function fetchRebrickableTheme(themeId) {
  if (!themeId) return "";
  const key = Config.rebrickableKey();
  const url = `https://rebrickable.com/api/v3/lego/themes/${encodeURIComponent(themeId)}/?key=${encodeURIComponent(key)}`;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return "";
  return JSON.parse(resp.getContentText()).name || "";
}

/* -------------------------
   Brickset RRP
   ------------------------- */
function fetchBricksetRRP(setId) {
  const apiKey = Config.bricksetKey();
  if (!apiKey) return null;
  const params = JSON.stringify({ setNumber: setId });
  const url =
    `https://brickset.com/api/v3.asmx/getSets` +
    `?apiKey=${encodeURIComponent(apiKey)}` +
    `&userHash=` +
    `&params=${encodeURIComponent(params)}`;

  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return null;
  const json = JSON.parse(resp.getContentText());
  if (!json.sets || json.sets.length === 0) return null;
  const s = json.sets[0];
  if (s.retailPrice && s.retailPrice.UK !== undefined) return s.retailPrice.UK;
  if (s.LEGOCom && s.LEGOCom.UK && s.LEGOCom.UK.retailPrice !== undefined) return s.LEGOCom.UK.retailPrice;
  return null;
}

/* -------------------------
   Chunked Cache Builder API
   ------------------------- */

/**
 * Start a cache job for a year.
 * This clears existing rows for that year and initialises progress/log.
 * Returns initial status message.
 */
function cacheYearWithProgress(year) {
  CACHE_PROGRESS.year = String(year);
  CACHE_PROGRESS.page = 1;
  CACHE_PROGRESS.pageSize = 100;
  CACHE_PROGRESS.done = false;
  CACHE_PROGRESS.log = [];
  CACHE_PROGRESS.lastMessage = `Starting cache for ${year}`;

  // Remove existing rows for this year from SetsCache (safe idempotent)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SetsCache') || ss.insertSheet('SetsCache');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Set_Num', 'Name', 'Theme_ID', 'Theme_Name',
      'Year', 'Num_Parts', 'Image_URL', 'Brickset_RRP', 'LastUpdated'
    ]);
  } else if (sheet.getLastRow() > 1) {
    const data = sheet.getRange(2,1,sheet.getLastRow()-1,9).getValues();
    const keep = data.filter(r => String(r[4]) !== String(year));
    sheet.getRange(2,1,sheet.getLastRow()-1,9).clearContent();
    if (keep.length > 0) sheet.getRange(2,1,keep.length,9).setValues(keep);
  }

  CACHE_PROGRESS.log.push(`Cleared existing cache rows for year ${year}`);
  CACHE_PROGRESS.log.push(`Initialised cache job (pageSize=${CACHE_PROGRESS.pageSize})`);
  CACHE_PROGRESS.lastMessage = `Ready to fetch page ${CACHE_PROGRESS.page} for ${year}`;
  return { message: CACHE_PROGRESS.lastMessage, done: CACHE_PROGRESS.done, log: CACHE_PROGRESS.log.slice(-20) };
}

/**
 * Continue the cache job by fetching one page.
 * Returns { message, done, log }.
 */
function cacheYearContinue() {
  if (!CACHE_PROGRESS.year) {
    return { message: 'No cache job running', done: true, log: CACHE_PROGRESS.log || [] };
  }
  if (CACHE_PROGRESS.done) {
    return { message: 'DONE', done: true, log: CACHE_PROGRESS.log || [] };
  }

  const year = CACHE_PROGRESS.year;
  const page = CACHE_PROGRESS.page;
  const pageSize = CACHE_PROGRESS.pageSize;
  const key = Config.rebrickableKey();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SetsCache') || ss.insertSheet('SetsCache');

  CACHE_PROGRESS.log.push(`Fetching page ${page} for ${year}...`);
  CACHE_PROGRESS.lastMessage = `Fetching page ${page} for ${year}...`;

  const url =
    `https://rebrickable.com/api/v3/lego/sets/?key=${encodeURIComponent(key)}` +
    `&year=${encodeURIComponent(year)}&page=${page}&page_size=${pageSize}`;

  let resp;
  try {
    resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (e) {
    CACHE_PROGRESS.log.push(`Fetch error: ${String(e)}`);
    CACHE_PROGRESS.done = true;
    CACHE_PROGRESS.lastMessage = `Error fetching page ${page}`;
    return { message: CACHE_PROGRESS.lastMessage, done: true, log: CACHE_PROGRESS.log.slice(-200) };
  }

  if (resp.getResponseCode() !== 200) {
    CACHE_PROGRESS.log.push(`HTTP ${resp.getResponseCode()} fetching page ${page}`);
    CACHE_PROGRESS.done = true;
    CACHE_PROGRESS.lastMessage = `Error: HTTP ${resp.getResponseCode()}`;
    return { message: CACHE_PROGRESS.lastMessage, done: true, log: CACHE_PROGRESS.log.slice(-200) };
  }

  const json = JSON.parse(resp.getContentText());
  const results = json.results || [];

  if (!results.length) {
    CACHE_PROGRESS.log.push(`No results on page ${page}. Finalising...`);
    // Finalise: fill theme names and update meta
    CACHE_PROGRESS.lastMessage = 'Updating theme names...';
    CACHE_PROGRESS.log.push(CACHE_PROGRESS.lastMessage);
    fillThemeNamesInCache(); // blocking call
    updateCacheMeta(year);
    CACHE_PROGRESS.log.push(`Cache complete for ${year}`);
    CACHE_PROGRESS.done = true;
    CACHE_PROGRESS.lastMessage = `DONE: Cache complete for ${year}`;
    return { message: CACHE_PROGRESS.lastMessage, done: true, log: CACHE_PROGRESS.log.slice(-200) };
  }

  // Append rows
  const rows = results.map(s => [
    s.set_num,
    s.name,
    s.theme_id || '',
    '',
    s.year || '',
    s.num_parts || '',
    s.set_img_url || '',
    '',
    new Date()
  ]);

  sheet.getRange(sheet.getLastRow()+1, 1, rows.length, 9).setValues(rows);
  CACHE_PROGRESS.log.push(`Fetched page ${page} (${rows.length} sets) and appended to SetsCache`);
  CACHE_PROGRESS.lastMessage = `Fetched page ${page} (${rows.length} sets)`;

  // Prepare next page or finish
  CACHE_PROGRESS.page = page + 1;
  if (!json.next) {
    CACHE_PROGRESS.log.push('No more pages reported by Rebrickable; finalising...');
    // Finalise
    CACHE_PROGRESS.lastMessage = 'Updating theme names...';
    CACHE_PROGRESS.log.push(CACHE_PROGRESS.lastMessage);
    fillThemeNamesInCache();
    updateCacheMeta(year);
    CACHE_PROGRESS.log.push(`Cache complete for ${year}`);
    CACHE_PROGRESS.done = true;
    CACHE_PROGRESS.lastMessage = `DONE: Cache complete for ${year}`;
    return { message: CACHE_PROGRESS.lastMessage, done: true, log: CACHE_PROGRESS.log.slice(-200) };
  }

  // Not done yet
  return { message: CACHE_PROGRESS.lastMessage, done: false, log: CACHE_PROGRESS.log.slice(-200) };
}

/**
 * Return the current log (last N lines).
 */
function getCacheLog(limit) {
  limit = parseInt(limit,10) || 500;
  return (CACHE_PROGRESS.log || []).slice(-limit);
}

/* -------------------------
   Cache helpers
   ------------------------- */
function fillThemeNamesInCache() {
  const key = Config.rebrickableKey();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SetsCache');
  if (!sheet || sheet.getLastRow() <= 1) return;

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,9).getValues();
  const themeIds = [...new Set(data.map(r => r[2]).filter(Boolean))];
  const themeMap = {};

  themeIds.forEach(id => {
    try {
      const url = `https://rebrickable.com/api/v3/lego/themes/${encodeURIComponent(id)}/?key=${encodeURIComponent(key)}`;
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        themeMap[id] = JSON.parse(resp.getContentText()).name || '';
      }
    } catch (e) {}
  });

  const updated = data.map(r => {
    r[3] = themeMap[r[2]] || '';
    return r;
  });

  sheet.getRange(2,1,updated.length,9).setValues(updated);
}

function updateCacheMeta(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CacheMeta') || ss.insertSheet('CacheMeta');
  if (sheet.getLastRow() === 0) sheet.appendRow(['Year','LastSynced']);

  const data = sheet.getRange(2,1,Math.max(0,sheet.getLastRow()-1),2).getValues();
  let found = false;
  for (let i=0;i<data.length;i++) {
    if (String(data[i][0]) === String(year)) {
      sheet.getRange(i+2,2).setValue(new Date());
      found = true;
      break;
    }
  }
  if (!found) sheet.appendRow([year, new Date()]);
}

/* -------------------------
   Local cache search
   ------------------------- */
function searchLocalCache(query, themeFilter, yearFilter, page) {
  page = parseInt(page,10) || 1;
  const pageSize = 10;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SetsCache');
  if (!sheet || sheet.getLastRow() <= 1) {
    return { results: [], page, pageSize, count: 0, totalPages: 0 };
  }

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,9).getValues();
  const q = String(query || '').trim().toLowerCase();

  let filtered = data.filter(r => {
    const setNum = String(r[0]).toLowerCase();
    const name = String(r[1]).toLowerCase();
    const year = String(r[4]);
    const themeId = String(r[2]);

    if (themeFilter && themeId !== String(themeFilter)) return false;
    if (yearFilter && year !== String(yearFilter)) return false;
    if (!q) return true;
    return setNum.includes(q) || name.includes(q);
  });

  const totalCount = filtered.length;
  const totalPages = Math.max(1, Math.ceil(totalCount / pageSize));
  const start = (page - 1) * pageSize;
  const pageRows = filtered.slice(start, start + pageSize);

  const results = pageRows.map(r => ({
    set_num: r[0],
    name: r[1],
    year: r[4],
    theme_id: r[2],
    theme_name: r[3],
    img: r[6],
    rrp: r[7]
  }));

  return { results, page, pageSize, count: totalCount, totalPages };
}

/* -------------------------
   Insert set into main sheet
   ------------------------- */
function insertSetIntoSheet(setNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);
  const lastRow = Math.max(1, sheet.getLastRow());
  const targetRow = lastRow + 1;
  const normalised = normaliseSetId(setNum);
  sheet.getRange(targetRow, COL.SET_ID).setValue(normalised);
  updateRow(sheet, targetRow, normalised);
  return { ok: true, row: targetRow };
}

/* -------------------------
   LEGO.com wishlist sync
   ------------------------- */
function fetchLegoWishlistSetNumbers(url) {
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return [];
    const html = resp.getContentText();
    const match = html.match(/"wishlistItems":(\[.*?\])/s);
    if (!match) return [];
    const items = JSON.parse(match[1]);
    return items.map(i => i.product.productCode);
  } catch (e) {
    return [];
  }
}

function syncAllWishlists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName('WishlistSources');
  if (!src || src.getLastRow() < 2) return { ok: false, message: 'No wishlist sources' };

  const rows = src.getRange(2,1,src.getLastRow()-1,2).getValues();
  const wishlistMap = {};

  rows.forEach(([name, url]) => {
    if (!name || !url) return;
    const setNums = fetchLegoWishlistSetNumbers(url);
    setNums.forEach(num => {
      if (!wishlistMap[num]) wishlistMap[num] = [];
      wishlistMap[num].push(name);
    });
  });

  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { ok: false, message: 'Main sheet not found' };
  const last = sheet.getLastRow();
  if (last < 2) return { ok: true, message: 'No rows to update' };

  const ids = sheet.getRange(2, COL.SET_ID, last-1, 1).getValues();
  const out = [];
  for (let i=0;i<ids.length;i++) {
    const raw = String(ids[i][0] || '');
    const normalised = raw.replace(/-1$/, '');
    const lists = wishlistMap[normalised] || [];
    out.push([lists.join(', ')]);
  }
  sheet.getRange(2, COL.WISHLISTS, out.length, 1).setValues(out);
  return { ok: true, message: 'Wishlists synced' };
}

/* -------------------------
   Theme list for sidebar
   ------------------------- */
function getThemeList() {
  const key = Config.rebrickableKey();
  const url = `https://rebrickable.com/api/v3/lego/themes/?key=${encodeURIComponent(key)}&page_size=1000`;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return [];
  const json = JSON.parse(resp.getContentText());
  return (json.results || []).map(t => ({ id: t.id, name: t.name }));
}

/* -------------------------
   Sidebar opener
   ------------------------- */
function openSetSearchSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Search LEGO Sets');
  SpreadsheetApp.getUi().showSidebar(html);
}