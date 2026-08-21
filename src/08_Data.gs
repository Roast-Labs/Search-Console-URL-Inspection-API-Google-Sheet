/**
 * Append-only Data store, change detection, retention purge.
 * Every inspection writes a row, whether or not the status changed.
 */

function appendInspectionRows_(runId, timestamp, results, urlLookup) {
  if (!results || !results.length) return;
  const lookup = urlLookup || urlMapByUrl_();
  const lastVerdict = lastVerdictMap_();
  const rows = results.map(function (r) {
    const meta = lookup[r.url] || {};
    const previous = lastVerdict[r.url];
    const statusChanged = hasStatusChanged_(previous, r);
    if (r.verdict) lastVerdict[r.url] = r.verdict;
    return [
      runId,
      timestamp,
      r.url,
      r.section || meta.section || '',
      r.tags || meta.tags || '',
      r.verdict || '',
      r.coverageState || '',
      r.robotsTxtState || '',
      r.indexingState || '',
      r.googleCanonical || '',
      r.userCanonical || '',
      r.canonicalMatch === '' ? '' : r.canonicalMatch,
      r.lastCrawlTime || '',
      r.crawledAs || '',
      r.pageFetchState || '',
      r.sitemaps || '',
      r.referringUrls || '',
      r.mobileVerdict || '',
      r.richResultsVerdict || '',
      r.httpStatus || '',
      statusChanged,
      r.error || ''
    ];
  });
  appendRows_(getSheet_(SHEETS.DATA), rows);
}

function hasStatusChanged_(previousVerdict, result) {
  if (!previousVerdict) return false;
  if (result.error && !result.verdict) return false;
  return String(previousVerdict) !== String(result.verdict || '');
}

function lastVerdictMap_() {
  const map = {};
  readUrlRecords_().forEach(function (u) {
    if (u.url && u.lastVerdict) map[u.url] = u.lastVerdict;
  });
  return map;
}

function updateUrlsFromRunId_(runId) {
  if (!runId) return;
  const data = getSheet_(SHEETS.DATA);
  const last = data.getLastRow();
  if (last < 2) return;
  const values = data.getRange(2, 1, last - 1, DATA_HEADERS.length).getValues();
  const byUrl = {};
  values.forEach(function (row) {
    if (String(row[DATA_COL.RUN_ID - 1]) !== runId) return;
    byUrl[String(row[DATA_COL.URL - 1])] = {
      verdict: row[DATA_COL.VERDICT - 1],
      error: row[DATA_COL.ERROR - 1],
      timestamp: row[DATA_COL.TIMESTAMP - 1]
    };
  });
  const sheet = getSheet_(SHEETS.URLS);
  const urlLast = sheet.getLastRow();
  if (urlLast < 2) return;
  const urlRange = sheet.getRange(2, 1, urlLast - 1, URL_HEADERS.length);
  const urlValues = urlRange.getValues();
  const today = spreadsheetToday_();
  let changed = false;
  urlValues.forEach(function (row) {
    const url = String(row[0] || '').trim();
    const hit = byUrl[url];
    if (!hit) return;
    if (hit.error && !hit.verdict) return;
    row[URL_COL.LAST_CHECKED - 1] = hit.timestamp || new Date();
    row[URL_COL.LAST_VERDICT - 1] = hit.verdict || row[URL_COL.LAST_VERDICT - 1];
    row[URL_COL.NEXT_DUE - 1] = computeNextDue_(row[URL_COL.FREQUENCY - 1], row[URL_COL.LAST_CHECKED - 1], today);
    changed = true;
  });
  if (changed) urlRange.setValues(urlValues);
}

function purgeOldDataMenu() {
  const ui = SpreadsheetApp.getUi();
  const days = asNumber_(getSetting_(SETTINGS.DATA_RETENTION_DAYS), DEFAULTS.DATA_RETENTION_DAYS);
  const resp = ui.alert(
    'Purge old data',
    'Delete Data rows older than ' + days + ' days (DATA_RETENTION_DAYS)? This cannot be undone.',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;
  const n = purgeOldData_();
  ui.alert('Deleted ' + n + ' row(s).');
}

function purgeOldData_() {
  const days = asNumber_(getSetting_(SETTINGS.DATA_RETENTION_DAYS), DEFAULTS.DATA_RETENTION_DAYS);
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  const sheet = getSheet_(SHEETS.DATA);
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const timestamps = sheet.getRange(2, DATA_COL.TIMESTAMP, last - 1, 1).getValues();
  const keep = [];
  const all = sheet.getRange(2, 1, last - 1, DATA_HEADERS.length).getValues();
  all.forEach(function (row, i) {
    const ts = timestamps[i][0];
    if (ts instanceof Date && ts.getTime() < cutoff.getTime()) return;
    keep.push(row);
  });
  const deleted = all.length - keep.length;
  clearDataRows_(sheet);
  if (keep.length) appendRows_(sheet, keep);
  return deleted;
}

function countChangedInRun_(runId) {
  if (!runId) return 0;
  const sheet = getSheet_(SHEETS.DATA);
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const values = sheet.getRange(2, DATA_COL.RUN_ID, last - 1, DATA_COL.STATUS_CHANGED).getValues();
  let n = 0;
  values.forEach(function (row) {
    if (String(row[0]) === runId && asBool_(row[DATA_COL.STATUS_CHANGED - DATA_COL.RUN_ID])) n++;
  });
  return n;
}

function changedRowsForRun_(runId) {
  const sheet = getSheet_(SHEETS.DATA);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, DATA_HEADERS.length).getValues();
  return values.filter(function (row) {
    return String(row[DATA_COL.RUN_ID - 1]) === runId && asBool_(row[DATA_COL.STATUS_CHANGED - 1]);
  }).map(function (row) {
    return {
      url: row[DATA_COL.URL - 1],
      verdict: row[DATA_COL.VERDICT - 1],
      coverage: row[DATA_COL.COVERAGE - 1]
    };
  });
}

function updateDashboardBuckets_() {
  const runId = getSetting_(SETTINGS.LATEST_RUN_ID);
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.DASHBOARD);
  if (!sheet || !runId) return;
  const data = getSheet_(SHEETS.DATA);
  const last = data.getLastRow();
  const buckets = {
    '0–1 days': 0,
    '1–7 days': 0,
    '7–30 days': 0,
    '30–90 days': 0,
    '90+ days': 0,
    'Never crawled': 0
  };
  if (last >= 2) {
    const values = data.getRange(2, 1, last - 1, DATA_COL.LAST_CRAWL).getValues();
    const now = Date.now();
    values.forEach(function (row) {
      if (String(row[0]) !== String(runId)) return;
      const crawl = row[DATA_COL.LAST_CRAWL - 1];
      if (!(crawl instanceof Date) || isNaN(crawl.getTime())) {
        buckets['Never crawled']++;
        return;
      }
      const days = (now - crawl.getTime()) / 86400000;
      if (days <= 1) buckets['0–1 days']++;
      else if (days <= 7) buckets['1–7 days']++;
      else if (days <= 30) buckets['7–30 days']++;
      else if (days <= 90) buckets['30–90 days']++;
      else buckets['90+ days']++;
    });
  }
  const labels = Object.keys(buckets);
  sheet.getRange(32, 1, labels.length, 2).setValues(labels.map(function (k) {
    return [k, buckets[k]];
  }));
}
