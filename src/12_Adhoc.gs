/**
 * Ad-hoc runs with quota warning modal (trim / borrow / queue tomorrow).
 */

function showAdhocSidebar() {
  if (!isSetupComplete_()) {
    SpreadsheetApp.getUi().alert('Run Setup first.');
    return;
  }
  refreshReservation_();
  ensureDailyTrigger_();
  const html = HtmlService.createTemplateFromFile('AdhocSidebar')
    .evaluate()
    .setTitle('Run Ad-hoc')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getAdhocContext() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  let selected = [];
  if (sheet.getName() === SHEETS.URLS) {
    const range = ss.getActiveRange();
    if (range && range.getRow() >= 2) {
      const urls = sheet.getRange(range.getRow(), 1, range.getNumRows(), 1).getValues();
      selected = urls.map(function (r) { return String(r[0] || '').trim(); }).filter(Boolean);
    }
  }
  const snap = viewQuotaPayload_();
  const borrowAllowed = isTruthySetting_(SETTINGS.ADHOC_BORROW_ALLOWED);
  return {
    selected: selected,
    quota: snap,
    borrowAllowed: borrowAllowed
  };
}

function planAdhoc_(rawUrls) {
  const urls = normaliseAdhocUrls_(rawUrls);
  const snap = getQuotaSnapshot_();
  const requested = urls.length;
  const remaining = Math.max(0, snap.remaining);
  const used = snap.scheduledUsed + snap.adhocUsed;
  const hardMax = Math.max(0, snap.cap - used);
  return {
    urls: urls,
    requested: requested,
    remaining: remaining,
    used: used,
    reserved: snap.reserved,
    cap: snap.cap,
    hardMax: hardMax,
    canRunNow: requested > 0 && requested <= remaining,
    needsWarning: requested > remaining,
    borrowAllowed: isTruthySetting_(SETTINGS.ADHOC_BORROW_ALLOWED) && snap.reserved > 0,
    property: getPropertyUrl_()
  };
}

function normaliseAdhocUrls_(rawUrls) {
  const incoming = [];
  String(rawUrls || '').split(/\r?\n/).forEach(function (line) {
    const u = line.trim();
    if (u) incoming.push(u);
  });
  const lookup = urlMapByUrl_();
  const property = getPropertyUrl_();
  const seen = {};
  const out = [];
  incoming.forEach(function (url) {
    if (seen[url]) return;
    seen[url] = true;
    const meta = lookup[url] || {};
    out.push({
      url: url,
      row: meta.row || 0,
      section: meta.section || deriveSection_(url),
      tags: meta.tags || '',
      priority: 1,
      outside: property ? !urlBelongsToProperty_(url, property) : false
    });
  });
  return out;
}

function runAdhocNow(rawUrls) {
  const plan = planAdhoc_(rawUrls);
  if (!plan.requested) throw new Error('No URLs provided.');
  if (!plan.canRunNow) {
    return { needsWarning: true, plan: publicPlan_(plan) };
  }
  return executeAdhoc_(plan.urls, false, '');
}

function getQuotaWarningPlan(rawUrls) {
  return publicPlan_(planAdhoc_(rawUrls));
}

function publicPlan_(plan) {
  return {
    requested: plan.requested,
    remaining: plan.remaining,
    used: plan.used,
    reserved: plan.reserved,
    cap: plan.cap,
    hardMax: plan.hardMax,
    borrowAllowed: plan.borrowAllowed,
    message: 'You\'ve asked for **' + plan.requested + '** checks.\n' +
      '**' + plan.used + '** used today. **' + plan.reserved + '** reserved for tonight\'s scheduled run.\n' +
      'Only **' + plan.remaining + '** remaining.'
  };
}

function adhocTrimToFit(rawUrls) {
  const plan = planAdhoc_(rawUrls);
  const fit = plan.urls.slice(0, plan.remaining);
  const rest = plan.urls.slice(plan.remaining);
  if (rest.length) {
    deferUrlsUntilTomorrow_(rest.map(function (u) { return u.url; }), 'Ad-hoc trim to fit');
    rest.forEach(function (u) {
      if (u.row) {
        getSheet_(SHEETS.URLS).getRange(u.row, URL_COL.PRIORITY).setValue(1);
      }
    });
  }
  if (!fit.length) {
    return { ok: true, message: 'Nothing fitted in remaining quota. Remainder queued for tomorrow at Priority 1.' };
  }
  return executeAdhoc_(fit, false, 'Trimmed to fit; ' + rest.length + ' queued for tomorrow.');
}

function adhocBorrowFromTonight(rawUrls) {
  if (!isTruthySetting_(SETTINGS.ADHOC_BORROW_ALLOWED)) {
    throw new Error('Borrowing from the scheduled reservation is disabled (ADHOC_BORROW_ALLOWED).');
  }
  const plan = planAdhoc_(rawUrls);
  const take = plan.urls.slice(0, plan.hardMax);
  const rest = plan.urls.slice(plan.hardMax);
  const need = Math.max(0, take.length - plan.remaining);
  releaseReservation_(need);
  setProp_(PROPS.CURRENT_RUN_BORROWED, 'true');
  if (rest.length) {
    deferUrlsUntilTomorrow_(rest.map(function (u) { return u.url; }), 'Ad-hoc borrow overflow');
  }
  const notes = 'Borrowed ' + need + ' from tonight\'s reservation.' +
    (rest.length ? ' ' + rest.length + ' still over the cap were queued for tomorrow.' : '');
  return executeAdhoc_(take, true, notes);
}

function adhocQueueTomorrow(rawUrls) {
  const plan = planAdhoc_(rawUrls);
  deferUrlsUntilTomorrow_(plan.urls.map(function (u) { return u.url; }), 'Ad-hoc queued for tomorrow');
  plan.urls.forEach(function (u) {
    if (u.row) getSheet_(SHEETS.URLS).getRange(u.row, URL_COL.PRIORITY).setValue(1);
  });
  return { ok: true, message: 'Queued ' + plan.requested + ' URL(s) for tomorrow at Priority 1. No quota spent.' };
}

function executeAdhoc_(items, borrowed, notes) {
  const lock = acquireLock_();
  try {
    resetRunState_(RUN_TYPE.ADHOC, notes || 'Ad-hoc run', borrowed);
    buildAdhocQueue_(items, true);
    if (!items.length) {
      writeRunLog_({ status: RUN_STATUS.COMPLETE, notes: 'Empty ad-hoc run' });
      clearRunState_();
      return { ok: true, message: 'Nothing to run.' };
    }
    processQueueLocked_(Date.now());
    return { ok: true, message: 'Ad-hoc run started. If the queue is large, continuation triggers will finish it.' };
  } catch (err) {
    if (err.name === 'AbortRunError') {
      finalizeAbort_(err.message);
      return { ok: false, message: err.message };
    }
    throw err;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
