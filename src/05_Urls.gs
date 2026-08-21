/**
 * URLs sheet read/write, Next Due, Last Checked, Last Verdict.
 */

function readUrlRecords_() {
  const sheet = getSheet_(SHEETS.URLS);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, URL_HEADERS.length).getValues();
  return values.map(function (row, i) {
    const url = String(row[URL_COL.URL - 1] || '').trim();
    const lastChecked = row[URL_COL.LAST_CHECKED - 1];
    const nextDueRaw = row[URL_COL.NEXT_DUE - 1];
    return {
      row: i + 2,
      url: url,
      active: asBool_(row[URL_COL.ACTIVE - 1]),
      frequency: String(row[URL_COL.FREQUENCY - 1] || FREQUENCY.DAILY),
      section: String(row[URL_COL.SECTION - 1] || ''),
      tags: String(row[URL_COL.TAGS - 1] || ''),
      priority: asNumber_(row[URL_COL.PRIORITY - 1], 3),
      lastChecked: lastChecked instanceof Date && !isNaN(lastChecked.getTime()) ? lastChecked : '',
      nextDue: nextDueRaw instanceof Date && !isNaN(nextDueRaw.getTime()) ? nextDueRaw : '',
      lastVerdict: String(row[URL_COL.LAST_VERDICT - 1] || ''),
      orphaned: asBool_(row[URL_COL.ORPHANED - 1])
    };
  }).filter(function (u) { return !!u.url; });
}

function urlMapByUrl_() {
  const map = {};
  readUrlRecords_().forEach(function (u) {
    map[u.url] = u;
  });
  return map;
}

function applyUrlUpdates_(updates) {
  if (!updates || !updates.length) return;
  const sheet = getSheet_(SHEETS.URLS);
  const today = spreadsheetToday_();
  updates.forEach(function (u) {
    if (!u.row) return;
    if (u.lastChecked) {
      sheet.getRange(u.row, URL_COL.LAST_CHECKED).setValue(u.lastChecked);
    }
    if (u.lastVerdict !== undefined) {
      sheet.getRange(u.row, URL_COL.LAST_VERDICT).setValue(u.lastVerdict);
    }
    if (u.priority !== undefined) {
      sheet.getRange(u.row, URL_COL.PRIORITY).setValue(u.priority);
    }
    if (u.orphaned !== undefined) {
      sheet.getRange(u.row, URL_COL.ORPHANED).setValue(u.orphaned);
    }
    if (u.frequency || u.lastChecked) {
      const freq = u.frequency || sheet.getRange(u.row, URL_COL.FREQUENCY).getValue();
      const last = u.lastChecked || sheet.getRange(u.row, URL_COL.LAST_CHECKED).getValue();
      sheet.getRange(u.row, URL_COL.NEXT_DUE).setValue(computeNextDue_(freq, last, today));
    }
    if (u.nextDue !== undefined) {
      sheet.getRange(u.row, URL_COL.NEXT_DUE).setValue(u.nextDue);
    }
  });
}

/**
 * Batch-write Last Checked / Next Due / Last Verdict after a run.
 * Faster than per-row writes for hundreds of URLs.
 */
function applyUrlUpdatesBatch_(results, now) {
  if (!results || !results.length) return;
  const sheet = getSheet_(SHEETS.URLS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, 1, lastRow - 1, URL_HEADERS.length);
  const values = range.getValues();
  const byRow = {};
  results.forEach(function (r) {
    if (r.urlsRow) byRow[r.urlsRow] = r;
  });
  const today = spreadsheetToday_();
  let changed = false;
  values.forEach(function (row, i) {
    const sheetRow = i + 2;
    const r = byRow[sheetRow];
    if (!r) return;
    if (r.requeueTomorrow) {
      row[URL_COL.PRIORITY - 1] = 1;
      row[URL_COL.NEXT_DUE - 1] = addDays_(today, 1);
      changed = true;
      return;
    }
    if (r.error && !r.verdict) return;
    row[URL_COL.LAST_CHECKED - 1] = now;
    row[URL_COL.LAST_VERDICT - 1] = r.verdict || row[URL_COL.LAST_VERDICT - 1];
    row[URL_COL.NEXT_DUE - 1] = computeNextDue_(row[URL_COL.FREQUENCY - 1], now, today);
    changed = true;
  });
  if (changed) range.setValues(values);
}

function addDays_(date, days) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  d.setDate(d.getDate() + days);
  return d;
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEETS.URLS) return;
    const col = e.range.getColumn();
    if (col !== URL_COL.FREQUENCY && col !== URL_COL.LAST_CHECKED && col !== URL_COL.ACTIVE) return;
    const start = e.range.getRow();
    const num = e.range.getNumRows();
    if (start < 2) return;
    const today = spreadsheetToday_();
    for (let r = start; r < start + num; r++) {
      const freq = sheet.getRange(r, URL_COL.FREQUENCY).getValue();
      const last = sheet.getRange(r, URL_COL.LAST_CHECKED).getValue();
      sheet.getRange(r, URL_COL.NEXT_DUE).setValue(computeNextDue_(freq, last, today));
    }
  } catch (err) {
    // Simple triggers must not throw to the user.
  }
}

function addUrlRows_(records) {
  if (!records.length) return;
  compactUrlSheetIfNeeded_();
  const sheet = getSheet_(SHEETS.URLS);
  const today = spreadsheetToday_();
  const rows = records.map(function (r) {
    return [
      r.url,
      r.active !== false,
      r.frequency || FREQUENCY.DAILY,
      r.section || '',
      r.tags || '',
      r.priority || 3,
      '',
      computeNextDue_(r.frequency || FREQUENCY.DAILY, '', today),
      '',
      false
    ];
  });
  const start = lastUrlRow_(sheet) + 1;
  sheet.getRange(start, 1, rows.length, URL_HEADERS.length).setValues(rows);
  sheet.getRange(start, URL_COL.ACTIVE, rows.length, 1).insertCheckboxes();
  sheet.getRange(start, URL_COL.ORPHANED, rows.length, 1).insertCheckboxes();
}

function lastUrlRow_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 1;
  const urls = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = urls.length - 1; i >= 0; i--) {
    if (String(urls[i][0] || '').trim()) return i + 2;
  }
  return 1;
}

/**
 * Whole-column checkboxes make getLastRow() ~1000, so imports were appended
 * far below the visible header. Pack real URL rows back to row 2.
 */
function compactUrlSheetIfNeeded_() {
  const sheet = getSheet_(SHEETS.URLS);
  if (sheet.getLastRow() < 2) return;
  if (String(sheet.getRange(2, 1).getValue() || '').trim()) return;
  const records = readUrlRecords_();
  if (!records.length) return;

  const last = sheet.getLastRow();
  const width = URL_HEADERS.length;
  sheet.getRange(2, 1, last - 1, width).clearContent();
  try {
    sheet.getRange(2, URL_COL.ACTIVE, last - 1, 1).removeCheckboxes();
    sheet.getRange(2, URL_COL.ORPHANED, last - 1, 1).removeCheckboxes();
  } catch (e) {}

  const today = spreadsheetToday_();
  const rows = records.map(function (r) {
    return [
      r.url,
      r.active !== false,
      r.frequency || FREQUENCY.DAILY,
      r.section || '',
      r.tags || '',
      r.priority || 3,
      r.lastChecked || '',
      r.nextDue || computeNextDue_(r.frequency, r.lastChecked, today),
      r.lastVerdict || '',
      r.orphaned === true
    ];
  });
  sheet.getRange(2, 1, rows.length, width).setValues(rows);
  sheet.getRange(2, URL_COL.ACTIVE, rows.length, 1).insertCheckboxes();
  sheet.getRange(2, URL_COL.ORPHANED, rows.length, 1).insertCheckboxes();
}
