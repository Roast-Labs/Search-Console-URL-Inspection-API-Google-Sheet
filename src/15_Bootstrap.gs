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
  const dashboard = ensureSheet_(ss, SHEETS.DASHBOARD, { tabColor: COLORS.OK });
  const glossary = ensureSheet_(ss, SHEETS.GLOSSARY, { tabColor: '#f9ab00' });
  const queue = ensureSheet_(ss, SHEETS.QUEUE, { hidden: true, tabColor: COLORS.HIDDEN_TAB });
  const quota = ensureSheet_(ss, SHEETS.QUOTA, { hidden: true, tabColor: COLORS.HIDDEN_TAB });
  const importSheet = ensureSheet_(ss, SHEETS.IMPORT, { hidden: true, tabColor: COLORS.HIDDEN_TAB });
  const deferred = ensureSheet_(ss, SHEETS.DEFERRED, { hidden: true, tabColor: COLORS.HIDDEN_TAB });

  writeSettings_(settings);
  writeHeaderSheet_(urls, URL_HEADERS, { checkboxes: [URL_COL.ACTIVE, URL_COL.ORPHANED] });
  writeCreditCalculator_(urls);
  writeUrlValidations_(urls);
  writeHeaderSheet_(data, DATA_HEADERS);
  writeHeaderSheet_(runLog, RUN_LOG_HEADERS);
  writeHeaderSheet_(queue, QUEUE_HEADERS);
  writeQuotaSheet_(quota);
  writeHeaderSheet_(importSheet, IMPORT_HEADERS);
  writeHeaderSheet_(deferred, DEFERRED_HEADERS);
  writeGlossary_(glossary);
  createNamedRanges_(ss, settings, urls, data);
  writePivotsSheet_(pivots);
  writeDashboard_(dashboard);
  tryCreatePivotTables_(ss, data, pivots);
  tryCreateCharts_(dashboard, quota);

  const order = [
    SHEETS.SETTINGS, SHEETS.URLS, SHEETS.DATA, SHEETS.RUN_LOG,
    SHEETS.PIVOTS, SHEETS.DASHBOARD, SHEETS.GLOSSARY,
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
  const map = {
    PROPERTY_URL: settings.getRange('B4'),
    DAILY_CAP: settings.getRange('B5'),
    BATCH_SIZE: settings.getRange('B6'),
    RUN_HOUR: settings.getRange('B7'),
    ALERT_EMAIL: settings.getRange('B8'),
    ALERT_ON_CHANGE_ONLY: settings.getRange('B9'),
    DATA_RETENTION_DAYS: settings.getRange('B10'),
    LANGUAGE_CODE: settings.getRange('B11'),
    ADHOC_BORROW_ALLOWED: settings.getRange('B12'),
    TODAY_PT: settings.getRange('B13'),
    LATEST_RUN_ID: settings.getRange('B14'),
    URL_COL: urls.getRange('A2:A'),
    URL_ACTIVE: urls.getRange('B2:B'),
    URL_FREQUENCY: urls.getRange('C2:C'),
    URL_LAST_VERDICT: urls.getRange('I2:I'),
    URL_ORPHANED: urls.getRange('J2:J'),
    DATA_TABLE: data.getRange('A:V')
  };
  Object.keys(map).forEach(function (name) {
    const existing = ss.getRangeByName(name);
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

function writePivotsSheet_(sheet) {
  sheet.clear();
  sheet.getRange('A1').setValue('Pivots over Data!A:V. Built once in the template so they expand with new rows.')
    .setFontColor('#5f6368');
  const labels = [
    [3, 'Coverage by section'],
    [22, 'Verdict by tag'],
    [41, 'Verdict trend'],
    [60, 'Canonical mismatches'],
    [79, 'Changed since last check'],
    [98, 'Fetch failures']
  ];
  labels.forEach(function (item) {
    sheet.getRange(item[0], 1).setValue(item[1]).setFontWeight('bold').setFontColor(COLORS.HEADER_BG);
  });
}

function tryCreatePivotTables_(ss, data, pivots) {
  const source = data.getRange('A1:V');
  try { pivots.getPivotTables().forEach(function (p) { /* cannot easily delete */ }); } catch (e) {}
  function make(anchorA1, rowCols, colCols, valueCol, filterCol, filterValue) {
    const table = source.createPivotTable(pivots.getRange(anchorA1));
    (rowCols || []).forEach(function (c) { table.addRowGroup(c); });
    (colCols || []).forEach(function (c) { table.addColumnGroup(c); });
    table.addPivotValue(valueCol, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
    if (filterCol) {
      table.addFilter(filterCol, SpreadsheetApp.newFilterCriteria().whenTextEqualTo(filterValue).build());
    }
    return table;
  }
  try {
    make('A4', [DATA_COL.SECTION], [DATA_COL.COVERAGE], DATA_COL.URL);
    make('A23', [DATA_COL.TAGS], [DATA_COL.VERDICT], DATA_COL.URL);
    const trend = make('A42', [DATA_COL.TIMESTAMP], [DATA_COL.VERDICT], DATA_COL.URL);
    try {
      trend.getRowGroups()[0].setDateTimeGroupingRule(SpreadsheetApp.DateTimeGroupingRuleType.DAY);
    } catch (e) {}
    make('A61', [DATA_COL.SECTION], [DATA_COL.CANONICAL_MATCH], DATA_COL.URL);
    make('A80', [DATA_COL.URL], [], DATA_COL.URL, DATA_COL.STATUS_CHANGED, 'TRUE');
    make('A99', [DATA_COL.PAGE_FETCH], [DATA_COL.SECTION], DATA_COL.URL);
  } catch (err) {
    pivots.getRange('A2').setValue('Pivot creation failed (' + err.message + '). Use the QUERY helpers on the Dashboard, or recreate pivots manually from Data!A:V.');
  }
}

function writeDashboard_(sheet) {
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:F1').merge()
    .setValue('Index status dashboard')
    .setFontSize(18).setFontWeight('bold')
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG);

  const labels = ['Active URLs', 'Indexed %', 'Changed today', 'Quota used today (PT)', 'Orphaned'];
  sheet.getRange(2, 1, 1, labels.length).setValues([labels]).setFontColor('#5f6368');
  sheet.getRange('A3').setFormula('=COUNTIF(URL_ACTIVE,TRUE)');
  sheet.getRange('B3').setFormula('=IFERROR((COUNTIF(URL_LAST_VERDICT,"PASS")+COUNTIF(URL_LAST_VERDICT,"PARTIAL"))/MAX(COUNTA(URL_LAST_VERDICT),1),0)');
  sheet.getRange('B3').setNumberFormat('0.0%');
  sheet.getRange('C3').setFormula('=COUNTIFS(Data!B:B,">="&TODAY(),Data!B:B,"<"&TODAY()+1,Data!U:U,TRUE)');
  sheet.getRange('D3').setFormula('=IFERROR(SUMIF(\'_Quota\'!A:A,TODAY_PT,\'_Quota\'!B:B)+SUMIF(\'_Quota\'!A:A,TODAY_PT,\'_Quota\'!C:C),0)');
  sheet.getRange('E3').setFormula('=COUNTIF(URL_ORPHANED,TRUE)');
  sheet.getRange('A3:E3').setFontSize(18).setFontWeight('bold').setFontColor(COLORS.HEADER_BG);

  sheet.getRange('A5').setValue('Charts sit on the QUERY helpers in columns L–Z and the crawl-recency table at A32. Scorecards are live formulas.')
    .setFontColor('#5f6368');

  sheet.getRange('L1').setValue('Verdict trend (for stacked area)');
  sheet.getRange('L2').setFormula('=IFERROR(QUERY(Data!B:F,"select toDate(B), count(C) where B is not null and F<>\'\' group by toDate(B) pivot F label toDate(B) \'Date\'",1),"No data yet")');

  sheet.getRange('L20').setValue('Coverage by section, latest run (for stacked bar)');
  sheet.getRange('L21').setFormula('=IFERROR(QUERY(Data!A:G,"select D, count(C) where A = \'"&LATEST_RUN_ID&"\' and D is not null group by D pivot G",1),"No data yet")');

  sheet.getRange('L40').setValue('Quota last 30 rows (for line chart)');
  sheet.getRange('L41').setFormula('=IFERROR(QUERY(\'_Quota\'!A:D,"select A, B+C, D, E order by A desc limit 30 label B+C \'Used\', D \'Reserved\', E \'Remaining\'",1),"No quota rows")');

  sheet.getRange('A30').setValue('Crawl recency (latest run)').setFontWeight('bold');
  sheet.getRange('A31:B31').setValues([['Bucket', 'URLs']]).setFontWeight('bold');
  sheet.getRange('A32:B37').setValues([
    ['0–1 days', 0],
    ['1–7 days', 0],
    ['7–30 days', 0],
    ['30–90 days', 0],
    ['90+ days', 0],
    ['Never crawled', 0]
  ]);

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 120);
}

function tryCreateCharts_(dashboard, quota) {
  dashboard.getCharts().forEach(function (c) {
    dashboard.removeChart(c);
  });
  function add(type, range, title, row, col, stacked) {
    try {
      let b = dashboard.newChart()
        .setChartType(type)
        .addRange(range)
        .setPosition(row, col, 0, 0)
        .setNumHeaders(1)
        .setOption('title', title)
        .setOption('legend', { position: 'bottom' });
      if (stacked) b = b.setOption('isStacked', true);
      dashboard.insertChart(b.build());
    } catch (err) {
      dashboard.getRange('A6').setValue('Chart "' + title + '" was not created: ' + err.message);
    }
  }
  add(Charts.ChartType.AREA, dashboard.getRange('L2:Q16'), 'Indexed vs not indexed over time', 8, 1, true);
  add(Charts.ChartType.BAR, dashboard.getRange('L21:Q35'), 'Coverage state by section, latest run', 8, 6, true);
  add(Charts.ChartType.COLUMN, dashboard.getRange('A31:B37'), 'Crawl recency distribution', 24, 1, false);
  add(Charts.ChartType.LINE, dashboard.getRange('L41:O71'), 'Quota used vs cap, last 30 days', 24, 6, false);
}
