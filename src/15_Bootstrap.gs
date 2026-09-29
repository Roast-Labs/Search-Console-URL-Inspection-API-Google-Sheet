/**
 * Builds the template skeleton: tabs, named ranges, formulas, validations, pivots, charts.
 * Idempotent for structure. Does not wipe Data, URLs, Run Log, or _Quota history.
 */

function bootstrapTemplate() {
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();
  lock.waitLock(API_LIMITS.LOCK_WAIT_MS);
  try {
    bootstrapLocked_(ss);
    ss.toast('Sheet structure is ready. Use Index Checker → Setup.', 'Index Checker', 8);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function sheetStructureReady_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  return !!(ss.getSheetByName(SHEETS.SETTINGS) &&
    ss.getSheetByName(SHEETS.URLS) &&
    ss.getSheetByName(SHEETS.DATA));
}

/**
 * Create tabs + named ranges if this is a blank copy (script only, no Settings tab).
 * Safe to call from Setup. Does not rebuild when Settings/URLs/Data already exist.
 */
function ensureSheetStructure_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  if (sheetStructureReady_(ss)) {
    ensureNamedRanges_(ss);
    return true;
  }
  if (ensureSheetStructure_.running) return false;
  const lock = LockService.getDocumentLock();
  lock.waitLock(API_LIMITS.LOCK_WAIT_MS);
  try {
    if (sheetStructureReady_(ss)) {
      ensureNamedRanges_(ss);
      return true;
    }
    ensureSheetStructure_.running = true;
    bootstrapLocked_(ss);
    return true;
  } finally {
    ensureSheetStructure_.running = false;
    try { lock.releaseLock(); } catch (e) {}
  }
}

function bootstrapLocked_(ss) {
  const existing = ss.getSheets();
  if (existing.length === 1 && ['Sheet1', 'Sheet 1'].indexOf(existing[0].getName()) !== -1) {
    existing[0].setName(SHEETS.SETTINGS);
  }

  const settings = ensureSheet_(ss, SHEETS.SETTINGS, { tabColor: COLORS.ACCENT });
  const urls = ensureSheet_(ss, SHEETS.URLS, { tabColor: COLORS.OK });
  const data = ensureSheet_(ss, SHEETS.DATA, { tabColor: COLORS.HEADER_BG });
  const runLog = ensureSheet_(ss, SHEETS.RUN_LOG, { tabColor: '#9aa0a6' });
  const existingPivots = ss.getSheetByName(SHEETS.PIVOTS);
  if (existingPivots && ss.getSheets().length > 1) {
    ss.deleteSheet(existingPivots);
  }
  const pivots = ensureSheet_(ss, SHEETS.PIVOTS, { tabColor: '#9334e6' });
  const trends = ensureSheet_(ss, SHEETS.TRENDS, { tabColor: '#9334e6' });
  const charts = ensureSheet_(ss, SHEETS.CHARTS, { tabColor: COLORS.OK });
  deleteSheetIfExists_(ss, SHEETS.DASHBOARD);
  const glossary = ensureSheet_(ss, SHEETS.GLOSSARY, { tabColor: '#f9ab00' });
  const queue = ensureSheet_(ss, SHEETS.QUEUE, { hidden: true, tabColor: COLORS.HIDDEN_TAB });
  const quota = ensureSheet_(ss, SHEETS.QUOTA, { hidden: true, tabColor: COLORS.HIDDEN_TAB });
  const importSheet = ensureSheet_(ss, SHEETS.IMPORT, { hidden: true, tabColor: COLORS.HIDDEN_TAB });
  const deferred = ensureSheet_(ss, SHEETS.DEFERRED, { hidden: true, tabColor: COLORS.HIDDEN_TAB });

  writeSettings_(settings);
  writeHeaderSheet_(urls, URL_HEADERS, { checkboxes: [URL_COL.ACTIVE, URL_COL.ORPHANED] });
  writeCreditCalculator_(urls);
  writeUrlValidations_(urls);
  repairDataSheet_(ss);
  writeHeaderSheet_(ss.getSheetByName(SHEETS.DATA), DATA_HEADERS);
  writeHeaderSheet_(runLog, RUN_LOG_HEADERS);
  writeHeaderSheet_(queue, QUEUE_HEADERS);
  writeQuotaSheet_(quota);
  writeHeaderSheet_(importSheet, IMPORT_HEADERS);
  writeHeaderSheet_(deferred, DEFERRED_HEADERS);
  writeGlossary_(glossary);
  const dataSheet = ss.getSheetByName(SHEETS.DATA);
  createNamedRanges_(ss, settings, urls, dataSheet);
  writePivotsSheet_(ss.getSheetByName(SHEETS.PIVOTS));
  refreshTrendAndChartSheets_();

  const order = [
    SHEETS.SETTINGS, SHEETS.URLS, SHEETS.DATA, SHEETS.RUN_LOG,
    SHEETS.PIVOTS, SHEETS.TRENDS, SHEETS.CHARTS, SHEETS.GLOSSARY,
    SHEETS.QUEUE, SHEETS.QUOTA, SHEETS.IMPORT, SHEETS.DEFERRED
  ];
  order.forEach(function (name, i) {
    const sh = ss.getSheetByName(name);
    if (sh) ss.setActiveSheet(sh) && ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(settings);
  SpreadsheetApp.flush();
}

function ensureSheet_(ss, name, opt) {
  opt = opt || {};
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (opt.hidden) sheet.hideSheet();
  else if (sheet.isSheetHidden()) sheet.showSheet();
  if (opt.tabColor) sheet.setTabColor(opt.tabColor);
  return sheet;
}

function deleteSheetIfExists_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (sheet && ss.getSheets().length > 1) ss.deleteSheet(sheet);
}

function writeHeaderSheet_(sheet, headers, opt) {
  opt = opt || {};
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, headers.length);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  if (sheet.getLastRow() < 2 && opt.checkboxes) {
    // Checkboxes are applied when rows are added.
  }
}

function styleHeaderRow_(sheet, width) {
  sheet.getRange(1, 1, 1, width)
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold')
    .setWrap(true);
}

function writeSettings_(sheet) {
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:C1').merge()
    .setValue('GSC Index Status Checker')
    .setFontSize(18).setFontWeight('bold')
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG);
  sheet.getRange('A2:C2').merge()
    .setValue('Uses the URL Inspection API (searchconsole.googleapis.com/v1/urlInspection/index:inspect). Not the Indexing API.')
    .setFontColor('#5f6368');

  const rows = [
    ['PROPERTY_URL', '', 'Set by Setup. Example: sc-domain:example.com or https://www.example.com/'],
    ['DAILY_CAP', DEFAULTS.DAILY_CAP, 'Leaves headroom under the 2,000 per-property official limit.'],
    ['BATCH_SIZE', DEFAULTS.BATCH_SIZE, 'URLs per execution slice before checking the 4.5 minute cutoff.'],
    ['RUN_HOUR', DEFAULTS.RUN_HOUR, 'Local hour for the daily trigger. Uses the Apps Script project timezone.'],
    ['ALERT_EMAIL', '', 'Blank disables email.'],
    ['ALERT_ON_CHANGE_ONLY', true, 'Only email when at least one Indexing Verdict changed.'],
    ['DATA_RETENTION_DAYS', DEFAULTS.DATA_RETENTION_DAYS, 'Purge Data rows older than this.'],
    ['LANGUAGE_CODE', DEFAULTS.LANGUAGE_CODE, 'Passed to the API as languageCode (BCP-47).'],
    ['ADHOC_BORROW_ALLOWED', true, 'If FALSE, the Borrow from tonight button is hidden.'],
    ['TODAY_PT', '', 'Pacific date, maintained by the script. Used by Dashboard quota formulas.'],
    ['LATEST_RUN_ID', '', 'Written at the end of each completed run.']
  ];
  sheet.getRange(4, 1, rows.length, 3).setValues(rows);
  rows.forEach(function (r, i) {
    sheet.getRange(4 + i, 1).setFontWeight('bold');
    sheet.getRange(4 + i, 3).setFontColor('#5f6368').setWrap(true);
  });

  sheet.getRange(4 + 5, 2).insertCheckboxes(); // ALERT_ON_CHANGE_ONLY
  sheet.getRange(4 + 8, 2).insertCheckboxes(); // ADHOC_BORROW_ALLOWED

  sheet.getRange(16, 1, 1, 3).merge()
    .setValue('Script timezone must match the operator\'s local timezone (Apps Script → Project Settings). Quota is always Pacific.')
    .setBackground(COLORS.SOFT).setWrap(true);

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 520);
}

function createNamedRanges_(ss, settings, urls, data) {
  applyNamedRanges_(ss, settings, urls, data, true);
}

/**
 * File → Make a copy / /copy often keeps tabs but drops named ranges.
 * Recreate any that are missing from Settings column A + URLs/Data sheets.
 */
function ensureNamedRanges_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (!settings) return false;
  applyNamedRanges_(ss, settings, ss.getSheetByName(SHEETS.URLS), ss.getSheetByName(SHEETS.DATA), false);
  return !!ss.getRangeByName(SETTINGS.PROPERTY_URL) || !!findSettingsValueCell_(settings, SETTINGS.PROPERTY_URL);
}

function findSettingsValueCell_(settings, name) {
  const last = Math.max(settings.getLastRow(), 1);
  const keys = settings.getRange(1, 1, last, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === name) return settings.getRange(i + 1, 2);
  }
  return null;
}

function applyNamedRanges_(ss, settings, urls, data, overwrite) {
  const map = {};
  Object.keys(SETTINGS).forEach(function (key) {
    const name = SETTINGS[key];
    if (name === SETTINGS.SETUP_COMPLETE_CELL) return;
    const cell = findSettingsValueCell_(settings, name);
    if (cell) map[name] = cell;
  });
  if (urls) {
    map.URL_COL = urls.getRange('A2:A');
    map.URL_ACTIVE = urls.getRange('B2:B');
    map.URL_FREQUENCY = urls.getRange('C2:C');
    map.URL_LAST_VERDICT = urls.getRange('I2:I');
    map.URL_ORPHANED = urls.getRange('J2:J');
  }
  if (data) map.DATA_TABLE = data.getRange('A:V');

  Object.keys(map).forEach(function (name) {
    const existing = ss.getRangeByName(name);
    if (existing && !overwrite) return;
    if (existing) ss.removeNamedRange(name);
    ss.setNamedRange(name, map[name]);
  });
}

function writeCreditCalculator_(urls) {
  urls.getRange('L1').setValue('Credit calculator').setFontWeight('bold').setFontSize(12).setFontColor(COLORS.HEADER_BG);
  urls.getRange('L2:L7').setValues([
    ['Daily commitments'],
    ['Weekly amortised'],
    ['Monthly amortised'],
    ['Projected daily average'],
    ['Worst-case day'],
    ['Headroom']
  ]);
  urls.getRange('M2').setFormula('=COUNTIFS(URL_ACTIVE,TRUE,URL_FREQUENCY,"Daily")');
  urls.getRange('M3').setFormula('=COUNTIFS(URL_ACTIVE,TRUE,URL_FREQUENCY,"Weekly")/7');
  urls.getRange('M4').setFormula('=COUNTIFS(URL_ACTIVE,TRUE,URL_FREQUENCY,"Monthly")/30');
  urls.getRange('M5').setFormula('=M2+M3+M4');
  urls.getRange('M6').setFormula('=COUNTIFS(URL_ACTIVE,TRUE,URL_FREQUENCY,"Daily")+COUNTIFS(URL_ACTIVE,TRUE,URL_FREQUENCY,"Weekly")+COUNTIFS(URL_ACTIVE,TRUE,URL_FREQUENCY,"Monthly")');
  urls.getRange('M7').setFormula('=DAILY_CAP-M6');
  urls.getRange('M3:M5').setNumberFormat('0.00');
  urls.getRange('M2').setNumberFormat('0');
  urls.getRange('M6:M7').setNumberFormat('0');
  urls.getRange('L8:M10').merge()
    .setValue('Live formulas, not script output. Ad-hoc URLs are excluded. Red on worst-case means it exceeds DAILY_CAP; amber means within 10% of the cap.')
    .setFontColor('#5f6368').setWrap(true);
  urls.setColumnWidth(12, 180);
  urls.setColumnWidth(13, 80);

  const rules = urls.getConditionalFormatRules();
  const red = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=M6>DAILY_CAP')
    .setBackground('#fce8e6')
    .setFontColor('#c5221f')
    .setRanges([urls.getRange('M6')])
    .build();
  const amber = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(M6<=DAILY_CAP,M6>=DAILY_CAP*0.9)')
    .setBackground('#fef7e0')
    .setFontColor('#e37400')
    .setRanges([urls.getRange('M6')])
    .build();
  urls.setConditionalFormatRules(rules.concat([red, amber]));
}

function writeUrlValidations_(urls) {
  const freq = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Daily', 'Weekly', 'Monthly', 'Ad-hoc'], true)
    .setAllowInvalid(false)
    .build();
  const pri = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5'], true)
    .setAllowInvalid(false)
    .build();
  urls.getRange('C2:C').setDataValidation(freq);
  urls.getRange('F2:F').setDataValidation(pri);
  urls.getRange('B2:B').setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build()
  );
  urls.getRange('J2:J').setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build()
  );
  urls.setColumnWidth(1, 360);
  for (let c = 2; c <= 10; c++) urls.setColumnWidth(c, 120);
}

function writeQuotaSheet_(sheet) {
  writeHeaderSheet_(sheet, QUOTA_HEADERS);
  sheet.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",DAILY_CAP-B2:B-C2:C-D2:D))');
}

function refreshPivotsSnapshot_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEETS.PIVOTS);
  if (!sheet) sheet = ensureSheet_(ss, SHEETS.PIVOTS, { tabColor: '#9334e6' });
  writePivotsSheet_(sheet);
  refreshTrendAndChartSheets_();
}

function refreshTrendAndChartSheets_() {
  const ss = SpreadsheetApp.getActive();
  const trends = ensureSheet_(ss, SHEETS.TRENDS, { tabColor: '#9334e6' });
  const charts = ensureSheet_(ss, SHEETS.CHARTS, { tabColor: COLORS.OK });
  const meta = writeTrendsSheet_(trends);
  writeChartsSheet_(charts, trends, meta);
}

/**
 * Latest-run snapshot. The whole Pivots tab is replaced at the end of every run.
 * Tables sit in a row so the user scrolls right, not down through overlapping QUERY spills.
 */
function writePivotsSheet_(sheet) {
  sheet.clear();
  const snap = latestRunSnapshot_();
  const tz = Session.getScriptTimeZone();
  const when = snap.timestamp
    ? Utilities.formatDate(snap.timestamp, tz, 'EEEE, d MMMM yyyy HH:mm')
    : 'No completed run yet';

  sheet.getRange(1, 1, 1, 12).merge()
    .setValue('Latest run — ' + when + ' (' + tz + ')')
    .setFontSize(16).setFontWeight('bold')
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG);
  sheet.getRange(2, 1, 1, 12).merge()
    .setValue(
      (snap.runId ? 'Run ID ' + snap.runId + '  ·  ' : '') +
      snap.rows.length + ' URL' + (snap.rows.length === 1 ? '' : 's') +
      ' in this run. This tab is overwritten each time the scheduled trigger, Run Now, or Ad-hoc run finishes.'
    )
    .setFontColor('#5f6368');

  const rows = snap.rows;
  const blocks = [
    { title: 'Coverage by section', table: countPivot_(rows, 3, 6) },
    { title: 'Verdict by tag', table: countPivot_(rows, 4, 5) },
    { title: 'Verdict this run', table: countPivot_(rows, function (r) {
      const ts = r[1];
      if (ts instanceof Date) return Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
      return String(ts || '');
    }, 5) },
    { title: 'Canonical mismatches', table: countPivot_(rows, 3, 11) },
    { title: 'Changed since last check', table: changedTable_(rows) },
    { title: 'Fetch state by section', table: countPivot_(rows, 14, 3) }
  ];

  let col = 1;
  let farthest = 12;
  blocks.forEach(function (b, i) {
    if (i === 1) col = Math.max(13, col);
    const width = Math.max(1, b.table[0].length);
    farthest = Math.max(farthest, col + width);
    col = col + width + 1;
  });
  ensureSheetColumns_(sheet, farthest + 2);

  col = 1;
  blocks.forEach(function (b, i) {
    if (i === 1) col = Math.max(13, col);
    writePivotBlock_(sheet, 4, col, b.title, b.table);
    col = col + Math.max(1, b.table[0].length) + 1;
  });
}

function writeTrendsSheet_(sheet) {
  sheet.clear();
  const tz = Session.getScriptTimeZone();
  const crawls = buildCrawlSeries_(tz);
  sheet.getRange(1, 1, 1, 11).merge()
    .setValue('Pivots over time — one row per crawl')
    .setFontSize(16).setFontWeight('bold')
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG);
  sheet.getRange(2, 1, 1, 11).merge()
    .setValue('Rebuilt at the end of every run from the full Data tab. Charts tab plots this table.')
    .setFontColor('#5f6368');

  const verdictHeaders = ['Date', 'Run ID', 'URLs', 'PASS', 'PARTIAL', 'FAIL', 'NEUTRAL', 'Other', 'Changed', 'Canonical mismatches', 'Indexed %'];
  sheet.getRange(4, 1).setValue('Verdict by crawl').setFontWeight('bold').setFontColor(COLORS.HEADER_BG);
  const verdictTable = [verdictHeaders];
  crawls.forEach(function (c) { verdictTable.push(c.verdictRow); });
  sheet.getRange(5, 1, verdictTable.length, verdictHeaders.length).setValues(verdictTable);
  sheet.getRange(5, 1, 1, verdictHeaders.length).setFontWeight('bold').setBackground(COLORS.BAND);
  sheet.getRange(6, 11, Math.max(crawls.length, 1), 1).setNumberFormat('0.0%');
  for (let c = 1; c <= 11; c++) sheet.setColumnWidth(c, c === 2 ? 180 : 130);

  const coverageKeys = {};
  crawls.forEach(function (c) {
    Object.keys(c.coverage).forEach(function (k) { coverageKeys[k] = true; });
  });
  const covCols = Object.keys(coverageKeys).sort();
  const covHeaders = ['Date', 'Run ID'].concat(covCols);
  const covStart = 7 + Math.max(crawls.length, 1);
  sheet.getRange(covStart, 1).setValue('Coverage by crawl').setFontWeight('bold').setFontColor(COLORS.HEADER_BG);
  const covTable = [covHeaders];
  crawls.forEach(function (c) {
    covTable.push([c.date, c.runId].concat(covCols.map(function (k) { return c.coverage[k] || 0; })));
  });
  if (covHeaders.length) {
    ensureSheetColumns_(sheet, covHeaders.length);
    sheet.getRange(covStart + 1, 1, covTable.length, covHeaders.length).setValues(covTable);
    sheet.getRange(covStart + 1, 1, 1, covHeaders.length).setFontWeight('bold').setBackground(COLORS.BAND);
  }

  const chartSrc = [['Date', 'PASS', 'PARTIAL', 'FAIL', 'NEUTRAL', 'Indexed %', 'Changed', 'Canonical mismatches']];
  crawls.forEach(function (c) {
    chartSrc.push([c.date, c.PASS, c.PARTIAL, c.FAIL, c.NEUTRAL, c.urls ? (c.PASS + c.PARTIAL) / c.urls : 0, c.changed, c.mismatch]);
  });
  ensureSheetColumns_(sheet, 20);
  sheet.getRange(4, 13).setValue('Chart source').setFontWeight('bold').setFontColor(COLORS.HEADER_BG);
  sheet.getRange(5, 13, chartSrc.length, 8).setValues(chartSrc);
  sheet.getRange(5, 13, 1, 8).setFontWeight('bold').setBackground(COLORS.BAND);
  if (crawls.length) sheet.getRange(6, 18, crawls.length, 1).setNumberFormat('0.0%');

  return {
    headerRow: 5,
    rows: verdictTable.length,
    chartCol: 13,
    chartRows: chartSrc.length
  };
}

function buildCrawlSeries_(tz) {
  const data = SpreadsheetApp.getActive().getSheetByName(SHEETS.DATA);
  if (!data || data.getLastRow() < 2) return [];
  const values = data.getRange(2, 1, data.getLastRow() - 1, DATA_HEADERS.length).getValues();
  const byRun = {};
  values.forEach(function (row) {
    const runId = String(row[0] || '').trim();
    if (!runId || runId === 'Run ID') return;
    if (!byRun[runId]) {
      const ts = row[1] instanceof Date ? row[1] : new Date();
      byRun[runId] = {
        runId: runId,
        timestamp: ts,
        date: ts instanceof Date && !isNaN(ts.getTime()) ? Utilities.formatDate(ts, tz, 'yyyy-MM-dd HH:mm') : '',
        urls: 0,
        PASS: 0,
        PARTIAL: 0,
        FAIL: 0,
        NEUTRAL: 0,
        other: 0,
        changed: 0,
        mismatch: 0,
        coverage: {}
      };
    }
    const b = byRun[runId];
    b.urls++;
    const v = String(row[5] || '');
    if (b[v] !== undefined) b[v]++;
    else b.other++;
    if (asBool_(row[20])) b.changed++;
    if (row[11] === false || row[11] === 'FALSE' || row[11] === 'false') b.mismatch++;
    const cov = String(row[6] || '(blank)');
    b.coverage[cov] = (b.coverage[cov] || 0) + 1;
    if (row[1] instanceof Date && row[1].getTime() > b.timestamp.getTime()) {
      b.timestamp = row[1];
      b.date = Utilities.formatDate(row[1], tz, 'yyyy-MM-dd HH:mm');
    }
  });
  const list = Object.keys(byRun).map(function (k) { return byRun[k]; });
  list.sort(function (a, b) { return a.timestamp.getTime() - b.timestamp.getTime(); });
  list.forEach(function (c) {
    const indexed = c.PASS + c.PARTIAL;
    c.verdictRow = [
      c.date,
      c.runId,
      c.urls,
      c.PASS,
      c.PARTIAL,
      c.FAIL,
      c.NEUTRAL,
      c.other,
      c.changed,
      c.mismatch,
      c.urls ? indexed / c.urls : 0
    ];
  });
  return list;
}

function writeChartsSheet_(sheet, trends, meta) {
  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  sheet.clear();
  sheet.getRange(1, 1, 1, 8).merge()
    .setValue('Charts — crawl history')
    .setFontSize(16).setFontWeight('bold')
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG);
  sheet.getRange(2, 1, 1, 8).merge()
    .setValue('Source: Pivots over time. Replaced at the end of every run.')
    .setFontColor('#5f6368');

  if (!meta || (meta.chartRows || meta.rows) < 2) {
    sheet.getRange(4, 1).setValue('No crawl history to chart yet. Complete a run first.');
    return;
  }

  const n = meta.chartRows || meta.rows;
  const start = meta.headerRow;
  const col = meta.chartCol || 13;
  try {
    sheet.insertChart(
      sheet.newChart()
        .setChartType(Charts.ChartType.AREA)
        .addRange(trends.getRange(start, col, n, 5))
        .setPosition(4, 1, 0, 0)
        .setNumHeaders(1)
        .setOption('title', 'Verdicts over time (stacked)')
        .setOption('isStacked', true)
        .setOption('legend', { position: 'bottom' })
        .setOption('width', 680)
        .setOption('height', 380)
        .build()
    );
  } catch (err) {
    sheet.getRange(4, 1).setValue('Verdict chart: ' + err.message);
  }
  try {
    sheet.insertChart(
      sheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(trends.getRange(start, col, n, 1))
        .addRange(trends.getRange(start, col + 5, n, 1))
        .setPosition(4, 9, 0, 0)
        .setNumHeaders(1)
        .setOption('title', 'Indexed % over time')
        .setOption('legend', { position: 'none' })
        .setOption('width', 480)
        .setOption('height', 380)
        .build()
    );
  } catch (err) {
    sheet.getRange(4, 9).setValue('Indexed % chart: ' + err.message);
  }
  try {
    sheet.insertChart(
      sheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(trends.getRange(start, col, n, 1))
        .addRange(trends.getRange(start, col + 6, n, 1))
        .setPosition(24, 1, 0, 0)
        .setNumHeaders(1)
        .setOption('title', 'Statuses changed per crawl')
        .setOption('legend', { position: 'none' })
        .setOption('width', 680)
        .setOption('height', 360)
        .build()
    );
  } catch (err) {
    sheet.getRange(24, 1).setValue('Changed chart: ' + err.message);
  }
  try {
    sheet.insertChart(
      sheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(trends.getRange(start, col, n, 1))
        .addRange(trends.getRange(start, col + 7, n, 1))
        .setPosition(24, 9, 0, 0)
        .setNumHeaders(1)
        .setOption('title', 'Canonical mismatches per crawl')
        .setOption('legend', { position: 'none' })
        .setOption('width', 480)
        .setOption('height', 360)
        .build()
    );
  } catch (err) {
    sheet.getRange(24, 9).setValue('Mismatch chart: ' + err.message);
  }
}

function ensureSheetColumns_(sheet, needed) {
  const have = sheet.getMaxColumns();
  if (needed > have) sheet.insertColumnsAfter(have, needed - have);
}

function latestRunSnapshot_() {
  const empty = { runId: '', timestamp: '', rows: [] };
  const data = SpreadsheetApp.getActive().getSheetByName(SHEETS.DATA);
  if (!data || data.getLastRow() < 2) return empty;
  const values = data.getRange(2, 1, data.getLastRow() - 1, DATA_HEADERS.length).getValues();
  let runId = '';
  try { runId = String(getSetting_(SETTINGS.LATEST_RUN_ID) || '').trim(); } catch (e) {}
  if (!runId) {
    let latest = 0;
    values.forEach(function (row) {
      const t = row[1] instanceof Date ? row[1].getTime() : 0;
      if (String(row[0] || '') && t >= latest) {
        latest = t;
        runId = String(row[0]);
      }
    });
  }
  const rows = values.filter(function (row) { return String(row[0] || '') === runId; });
  let timestamp = '';
  rows.forEach(function (row) {
    if (row[1] instanceof Date) timestamp = row[1];
  });
  return { runId: runId, timestamp: timestamp, rows: rows };
}

function countPivot_(rows, rowKey, colKey) {
  const rowFn = typeof rowKey === 'function' ? rowKey : function (r) { return r[rowKey]; };
  const colFn = typeof colKey === 'function' ? colKey : function (r) { return r[colKey]; };
  const grid = {};
  const colSet = {};
  rows.forEach(function (r) {
    const rk = String(rowFn(r) == null || rowFn(r) === '' ? '(blank)' : rowFn(r));
    const ck = String(colFn(r) == null || colFn(r) === '' ? '(blank)' : colFn(r));
    colSet[ck] = true;
    if (!grid[rk]) grid[rk] = {};
    grid[rk][ck] = (grid[rk][ck] || 0) + 1;
  });
  const cols = Object.keys(colSet).sort();
  const table = [[' '].concat(cols)];
  Object.keys(grid).sort().forEach(function (rk) {
    table.push([rk].concat(cols.map(function (c) { return grid[rk][c] || 0; })));
  });
  return table.length > 1 ? table : [['No data for this run']];
}

function changedTable_(rows) {
  const table = [['URL', 'Indexing Verdict', 'Coverage State']];
  rows.forEach(function (r) {
    if (!asBool_(r[20])) return;
    table.push([r[2] || '', r[5] || '', r[6] || '']);
  });
  return table.length > 1 ? table : [['No statuses changed in this run']];
}

function writePivotBlock_(sheet, row, col, title, table) {
  const width = Math.max(1, table[0].length);
  ensureSheetColumns_(sheet, col + width);
  sheet.getRange(row, col, 1, width).merge()
    .setValue(title)
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_FG);
  sheet.getRange(row + 1, col, table.length, width).setValues(table);
  sheet.getRange(row + 1, col, 1, width).setFontWeight('bold').setBackground(COLORS.BAND);
  sheet.setColumnWidth(col, 200);
}

/**
 * Restore Data headers and strip a leftover pivot on the Data tab.
 * A failed createPivotTable can land on Data and turn A1 into #REF!.
 */
function repairDataSheet_(ss) {
  const data = ss.getSheetByName(SHEETS.DATA);
  if (!data) return;
  let hasPivot = false;
  try {
    hasPivot = data.getPivotTables().length > 0;
  } catch (e) {}
  const header = String(data.getRange(1, 1).getValue() || '');
  const headerBroken = header !== 'Run ID';
  if (!hasPivot && !headerBroken) return;

  const lastRow = data.getLastRow();
  const lastCol = Math.max(data.getLastColumn(), DATA_HEADERS.length);
  let body = [];
  if (lastRow >= 1) {
    const values = data.getRange(1, 1, lastRow, lastCol).getValues();
    values.forEach(function (row) {
      const a = String(row[0] || '');
      const url = String(row[2] || '');
      const looksLikeRun = /^\d{8}-/.test(a);
      const looksLikeUrl = url.indexOf('http') === 0;
      if (looksLikeRun || looksLikeUrl) body.push(row);
    });
  }

  if (hasPivot) {
    const idx = data.getIndex();
    ss.deleteSheet(data);
    const neu = ss.insertSheet(SHEETS.DATA, idx);
    writeHeaderSheet_(neu, DATA_HEADERS);
    writeDataBody_(neu, body);
    ss.setNamedRange('DATA_TABLE', neu.getRange('A:V'));
    return;
  }

  writeHeaderSheet_(data, DATA_HEADERS);
}

function writeDataBody_(sheet, body) {
  if (!body.length) return;
  const width = DATA_HEADERS.length;
  const trimmed = body.map(function (row) {
    const out = [];
    for (let i = 0; i < width; i++) out.push(row[i] !== undefined && row[i] !== null ? row[i] : '');
    return out;
  });
  sheet.getRange(2, 1, trimmed.length, width).setValues(trimmed);
}

function rebuildPivots() {
  const ss = SpreadsheetApp.getActive();
  repairDataSheet_(ss);
  const existing = ss.getSheetByName(SHEETS.PIVOTS);
  if (existing && ss.getSheets().length > 1) ss.deleteSheet(existing);
  const pivots = ensureSheet_(ss, SHEETS.PIVOTS, { tabColor: '#9334e6' });
  writePivotsSheet_(pivots);
  const data = ss.getSheetByName(SHEETS.DATA);
  if (data) ss.setNamedRange('DATA_TABLE', data.getRange('A:V'));
  deleteSheetIfExists_(ss, SHEETS.DASHBOARD);
  try {
    refreshTrendAndChartSheets_();
    placeSheetAfter_(ss, SHEETS.TRENDS, SHEETS.PIVOTS);
    placeSheetAfter_(ss, SHEETS.CHARTS, SHEETS.TRENDS);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Pivots were updated but the new tabs failed: ' + err.message);
    return;
  }
  SpreadsheetApp.getUi().alert('Created/updated Pivots, Pivots over time, and Charts. Those tabs sit next to each other and are overwritten at the end of every run.');
}

function placeSheetAfter_(ss, name, afterName) {
  const sheet = ss.getSheetByName(name);
  const after = ss.getSheetByName(afterName);
  if (!sheet || !after) return;
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(after.getIndex() + 1);
}

