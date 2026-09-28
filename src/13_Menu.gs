/**
 * Menu, setup wizard, quota viewer.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Index Checker')
    .addItem('Setup', 'showSetupWizard')
    .addItem('Run Now', 'runNow')
    .addItem('Run Ad-hoc', 'showAdhocSidebar')
    .addItem('Import from Sitemap', 'showSitemapWizard')
    .addSeparator()
    .addItem('Build Queue Only', 'buildQueueOnly')
    .addItem('View Quota', 'showQuotaView')
    .addItem('Install / reset daily trigger', 'resetTriggers')
    .addItem('Purge Old Data', 'purgeOldDataMenu')
    .addSeparator()
    .addItem('Rebuild pivots', 'rebuildPivots')
    .addItem('Rebuild sheet structure', 'bootstrapTemplate')
    .addToUi();

    try {
      refreshTodayPtCell_();
      compactUrlSheetIfNeeded_();
      if (isSetupComplete_()) refreshReservation_();
    } catch (e) {}
}

function onInstall(e) {
  onOpen(e);
}

function showSetupWizard() {
  const html = HtmlService.createTemplateFromFile('Setup')
    .evaluate()
    .setWidth(520)
    .setHeight(560)
    .setTitle('Index Checker setup');
  SpreadsheetApp.getUi().showModalDialog(html, 'Index Checker setup');
}

function setupBegin() {
  const account = detectAccountType_();
  let sites = [];
  let sitesError = '';
  try {
    sites = listSearchConsoleSites_();
  } catch (err) {
    sitesError = err.message;
  }
  return {
    account: account,
    sites: sites,
    sitesError: sitesError,
    triggerCapNote: account.isConsumer
      ? 'This looks like a consumer (gmail.com) account. Trigger total runtime is ~90 minutes/day, which will cap practical throughput well below 2,000 URLs. A Google Workspace account has ~6 hours/day.'
      : 'Workspace-class account detected. Daily trigger runtime allowance is ~6 hours.',
    currentProperty: getPropertyUrl_(),
    runHour: getRunHour_(),
    setupComplete: isSetupComplete_(),
    timezone: Session.getScriptTimeZone()
  };
}

function setupSaveProperty(siteUrl, runHour) {
  if (!siteUrl) throw new Error('Choose a Search Console property.');
  setSetting_(SETTINGS.PROPERTY_URL, siteUrl);
  if (runHour !== undefined && runHour !== null && runHour !== '') {
    setSetting_(SETTINGS.RUN_HOUR, Number(runHour));
  }
  setProp_(PROPS.SETUP_COMPLETE, 'true');
  createDailyTrigger_();
  refreshReservation_();
  return { property: siteUrl };
}

function setupFinish(opts) {
  opts = opts || {};
  createDailyTrigger_();
  setProp_(PROPS.SETUP_COMPLETE, 'true');
  refreshReservation_();
  let smoke = null;
  if (opts.smokeTest) {
    smoke = runSmokeTest();
  }
  return {
    ok: true,
    timezone: Session.getScriptTimeZone(),
    runHour: getRunHour_(),
    smoke: smoke
  };
}

function showSitemapWizard() {
  if (!isSetupComplete_() && !getPropertyUrl_()) {
    SpreadsheetApp.getUi().alert('Run Setup first so a Search Console property is set.');
    return;
  }
  const html = HtmlService.createTemplateFromFile('SitemapWizard')
    .evaluate()
    .setWidth(640)
    .setHeight(620)
    .setTitle('Import from sitemap');
  SpreadsheetApp.getUi().showModalDialog(html, 'Import from sitemap');
}

function showQuotaView() {
  refreshReservation_();
  const html = HtmlService.createTemplateFromFile('QuotaView')
    .evaluate()
    .setWidth(420)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Quota');
}

function showQuotaWarningDialog(rawUrls) {
  const plan = getQuotaWarningPlan(rawUrls);
  const t = HtmlService.createTemplateFromFile('QuotaWarning');
  t.planJson = JSON.stringify(plan);
  t.rawUrls = String(rawUrls || '');
  const html = t.evaluate().setWidth(480).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Not enough quota');
}

function showBorrowConfirm(rawUrls) {
  const t = HtmlService.createTemplateFromFile('BorrowConfirm');
  t.rawUrls = String(rawUrls || '');
  const html = t.evaluate().setWidth(420).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'Borrow from tonight?');
}

/** Public aliases — google.script.run cannot call functions whose names end in _. */
function listGscSitemapsForUi() { return listGscSitemapsForUi_(); }
function startSitemapImportFromGsc(path) { return startSitemapImportFromGsc_(path); }
function startSitemapImportFromUrl(url) { return startSitemapImportFromUrl_(url); }
function filterImport(includeRegex, excludeRegex) { return filterImport_(includeRegex, excludeRegex); }
function applySectionMap(mapping) { return applySectionMap_(mapping); }
function previewReconciliation(defaults) { return previewReconciliation_(defaults); }
function commitSitemapImport(defaults) { return commitSitemapImport_(defaults); }
function viewQuotaPayload() { return viewQuotaPayload_(); }
