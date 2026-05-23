const PROPS = PropertiesService.getScriptProperties();
const REBRICKABLE_KEY = PROPS.getProperty('REBRICKABLE_KEY');
const BRICKSET_KEY = PROPS.getProperty('BRICKSET_KEY');
const COLLECTION_SHEET_ID = PROPS.getProperty('COLLECTION_SHEET_ID');

const SHEET_NAME = 'Wishlist';

const COL = {
  SET_ID:           1,
  SET_NAME:         2,
  THEME:            3,
  SUBTHEME:         4,
  YEAR:             5,
  PIECES:           6,
  UK_RRP:           7,
  TARGET_PRICE:     8,
  RETIREMENT_DATE:  9,
  PRIORITY:         10,
  REASON:           11,
  OWNED_QTY:        12,
  OWNED_STATUS:     13,
  COLLECTION_NOTES: 14,
  IMAGE_URL:        15,
  IMAGE_PREVIEW:    16
};

const COLLECTION_SHEET_TAB = 'Collection';
const HEADER_ROW = 1;
const DATA_START_ROW = 2;
const TOTAL_COLUMNS = 16;
const TARGET_PRICE_MULTIPLIER = 0.75;
const THEME_CACHE_DURATION_SECONDS = 21600; // 6 hours

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Wishlist')
    .addItem('Open Sidebar', 'openSidebar')
    .addSeparator() // Adds a nice dividing line in the menu
    .addItem('Fill Missing Set Data', 'fillMissingSetData')
    .addToUi();
}

function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('LEGO Wishlist');
  SpreadsheetApp.getUi().showSidebar(html);
}

function doGet() {
  return HtmlService.createTemplateFromFile('webapp')
    .evaluate()
    .setTitle('LEGO Wishlist')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


function setCollectionSheetId() {
  PropertiesService.getScriptProperties().setProperty(
    'COLLECTION_SHEET_ID', 
    '1YACslAO0Zv6aKrPH2mIuEgY9feApG6fbAjYuhHMxxUo'
  );
  console.log('Collection Sheet ID successfully saved!');
}



