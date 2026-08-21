/**
 * Named-range backed settings. Scripts must read by name, never by A1.
 */

function getSetting_(name) {
  const ss = SpreadsheetApp.getActive();
  const range = ss.getRangeByName(name);
  if (!range) {
    throw new Error('Named range missing: ' + name + '. Run bootstrapTemplate() to repair the sheet.');
  }
  return range.getValue();
}

function setSetting_(name, value) {
  const ss = SpreadsheetApp.getActive();
  const range = ss.getRangeByName(name);
  if (!range) {
    throw new Error('Named range missing: ' + name + '. Run bootstrapTemplate() to repair the sheet.');
  }
  range.setValue(value);
}

function getPropertyUrl_() {
  return String(getSetting_(SETTINGS.PROPERTY_URL) || '').trim();
}

function getDailyCap_() {
  const n = Number(getSetting_(SETTINGS.DAILY_CAP));
  return isFinite(n) && n > 0 ? n : DEFAULTS.DAILY_CAP;
}

function getBatchSize_() {
  const n = Number(getSetting_(SETTINGS.BATCH_SIZE));
  return isFinite(n) && n > 0 ? Math.floor(n) : DEFAULTS.BATCH_SIZE;
}

function getLanguageCode_() {
  const v = String(getSetting_(SETTINGS.LANGUAGE_CODE) || '').trim();
  return v || DEFAULTS.LANGUAGE_CODE;
}

function getRunHour_() {
  const n = Number(getSetting_(SETTINGS.RUN_HOUR));
  if (!isFinite(n)) return DEFAULTS.RUN_HOUR;
  return Math.max(0, Math.min(23, Math.floor(n)));
}

function isTruthySetting_(name) {
  const v = getSetting_(name);
  return v === true || v === 'TRUE' || v === 'true' || v === 1 || v === '1';
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet missing: ' + name + '. Run bootstrapTemplate() to repair the sheet.');
  }
  return sheet;
}

function props_() {
  return PropertiesService.getDocumentProperties();
}

function getProp_(key, fallback) {
  const v = props_().getProperty(key);
  return v === null || v === undefined ? fallback : v;
}

function setProp_(key, value) {
  props_().setProperty(key, String(value));
}

function deleteProp_(key) {
  props_().deleteProperty(key);
}

function isSetupComplete_() {
  if (getProp_(PROPS.SETUP_COMPLETE) === 'true') return true;
  try {
    if (getPropertyUrl_()) {
      setProp_(PROPS.SETUP_COMPLETE, 'true');
      return true;
    }
  } catch (e) {}
  return false;
}

function sheetToObjects_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const width = headers.length;
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  return values.map(function (row, i) {
    const obj = { _row: i + 2 };
    headers.forEach(function (h, idx) {
      obj[h] = row[idx];
    });
    return obj;
  });
}

function appendRows_(sheet, rows) {
  if (!rows || !rows.length) return;
  const last = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(last + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Header-row lookup so column order can shift without breaking writes.
 */
function headerIndexMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function (h, i) {
    map[String(h)] = i;
  });
  return map;
}
