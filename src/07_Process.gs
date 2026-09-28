/**
 * Resumable batch processor: lock, fetchAll chunks, continuation triggers, retry policy.
 */

function onScheduledRun() {
  startRun_(RUN_TYPE.SCHEDULED, 'Scheduled daily run', false);
}

function runNow() {
  ensureDailyTrigger_();
  startRun_(RUN_TYPE.SCHEDULED, 'Manual run', true);
}

function startRun_(type, notes, interactive) {
  if (!isSetupComplete_() && interactive) {
    SpreadsheetApp.getUi().alert('Run Setup from the Index Checker menu before inspecting URLs.');
    return;
  }
  const lock = acquireLock_();
  try {
    resetRunState_(type, notes);
    if (type === RUN_TYPE.SCHEDULED) {
      const built = buildQueueLocked_(true);
      if (!built.queued) {
        writeRunLog_({
          status: RUN_STATUS.COMPLETE,
          notes: (notes || '') + (notes ? ' — ' : '') + 'Nothing due.'
        });
        setProp_(PROPS.SCHEDULED_COMPLETE_PT, getPacificDate());
        setReserved_(0);
        clearRunState_();
        if (interactive) showToast_('No URLs are due.', 'Index Checker');
        return;
      }
    }
    processQueueLocked_(Date.now());
  } catch (err) {
    if (err.name === 'AbortRunError') {
      finalizeAbort_(err.message);
      if (interactive) SpreadsheetApp.getUi().alert(err.message);
      return;
    }
    finalizeAbort_(err.message || String(err));
    throw err;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function processQueue() {
  const lock = acquireLock_();
  try {
    bumpContinuationCount_();
    processQueueLocked_(Date.now());
  } catch (err) {
    if (err.name === 'AbortRunError') {
      finalizeAbort_(err.message);
      return;
    }
    finalizeAbort_(err.message || String(err));
    throw err;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function processQueueLocked_(startTime) {
  const batchSize = getBatchSize_();
  const propertyUrl = getPropertyUrl_();
  const runId = getProp_(PROPS.CURRENT_RUN_ID) || newRunId_();
  const runType = getProp_(PROPS.CURRENT_RUN_TYPE) || RUN_TYPE.SCHEDULED;
  setProp_(PROPS.CURRENT_RUN_ID, runId);

  const urlLookup = urlMapByUrl_();
  const processedThisSlice = [];

  while (Date.now() - startTime < API_LIMITS.MAX_RUNTIME_MS) {
    if (getProp_(PROPS.ABORT_FLAG) === 'true') {
      throw new AbortRunError_('Run aborted.');
    }
    const cursor = asNumber_(getProp_(PROPS.QUEUE_CURSOR, '2'), 2);
    const items = readQueuePending_(cursor, batchSize);
    if (!items.length) {
      finalizeSuccess_(processedThisSlice);
      return;
    }

    const chunks = chunk_(items, API_LIMITS.FETCH_CHUNK);
    for (let c = 0; c < chunks.length; c++) {
      if (Date.now() - startTime >= API_LIMITS.MAX_RUNTIME_MS) {
        persistAndContinue_();
        return;
      }
      const chunkItems = chunks[c];
      chunkItems.forEach(function (item) {
        if (!item.section && urlLookup[item.url]) {
          item.section = urlLookup[item.url].section;
          item.tags = urlLookup[item.url].tags;
          if (!item.urlsRow) item.urlsRow = urlLookup[item.url].row;
        }
      });

      const outside = [];
      const inside = [];
      chunkItems.forEach(function (item) {
        if (propertyUrl && !urlBelongsToProperty_(item.url, propertyUrl)) {
          outside.push(Object.assign({}, item, {
            httpStatus: 400,
            error: 'URL is outside property ' + propertyUrl,
            verdict: '',
            coverageState: '',
            robotsTxtState: '',
            indexingState: '',
            googleCanonical: '',
            userCanonical: '',
            canonicalMatch: '',
            lastCrawlTime: '',
            crawledAs: '',
            pageFetchState: '',
            sitemaps: '',
            referringUrls: '',
            mobileVerdict: '',
            richResultsVerdict: ''
          }));
        } else {
          inside.push(item);
        }
      });

      const results = outside.slice();
      if (inside.length) {
        const inspected = inspectChunk_(inside);
        inspected.forEach(function (r) { results.push(r); });
      }

      const abort = results.filter(function (r) { return r.abort; })[0];
      const now = new Date();
      appendInspectionRows_(runId, now, results, urlLookup);
      const apiCalls = inside.length;
      incrementQuota_(apiCalls, runType === RUN_TYPE.ADHOC ? RUN_TYPE.ADHOC : RUN_TYPE.SCHEDULED);

      if (abort) {
        throw new AbortRunError_(abort.error);
      }

      const requeue = results.filter(function (r) { return r.requeueTomorrow; }).map(function (r) { return r.url; });
      if (requeue.length) deferUrlsUntilTomorrow_(requeue, '429 after retries');

      const succeeded = results.filter(function (r) { return !r.error && !r.requeueTomorrow; }).length;
      const failed = results.filter(function (r) { return r.error && !r.requeueTomorrow; }).length;
      addRunCounter_(PROPS.CURRENT_RUN_ATTEMPTED, results.length);
      addRunCounter_(PROPS.CURRENT_RUN_SUCCEEDED, succeeded);
      addRunCounter_(PROPS.CURRENT_RUN_FAILED, failed);

      markQueueRows_(chunkItems.map(function (i) { return i.queueRow; }), 'Done');
      setProp_(PROPS.QUEUE_CURSOR, String(chunkItems[chunkItems.length - 1].queueRow + 1));
      processedThisSlice.push.apply(processedThisSlice, results);

      Utilities.sleep(API_LIMITS.CHUNK_SLEEP_MS);
    }
  }

  persistAndContinue_();
}

function persistAndContinue_() {
  addRunCounter_(PROPS.CURRENT_RUN_CONTINUATIONS, 1);
  deleteContinuationTriggers_();
  ScriptApp.newTrigger('processQueue').timeBased().after(60 * 1000).create();
  showToast_('Paused at the 4.5 minute mark. A continuation trigger will resume in ~1 minute.', 'Index Checker');
}

function bumpContinuationCount_() {
  if (!getProp_(PROPS.CURRENT_RUN_ID)) {
    resetRunState_(RUN_TYPE.CONTINUATION, 'Continuation without parent run state');
  }
}

function resetRunState_(type, notes, borrowed) {
  setProp_(PROPS.CURRENT_RUN_ID, newRunId_());
  setProp_(PROPS.CURRENT_RUN_TYPE, type);
  setProp_(PROPS.CURRENT_RUN_START, String(Date.now()));
  setProp_(PROPS.CURRENT_RUN_ATTEMPTED, '0');
  setProp_(PROPS.CURRENT_RUN_SUCCEEDED, '0');
  setProp_(PROPS.CURRENT_RUN_FAILED, '0');
  setProp_(PROPS.CURRENT_RUN_CONTINUATIONS, '0');
  setProp_(PROPS.CURRENT_RUN_BORROWED, borrowed ? 'true' : 'false');
  setProp_(PROPS.CURRENT_RUN_NOTES, notes || '');
  deleteProp_(PROPS.ABORT_FLAG);
}

function addRunCounter_(key, n) {
  const cur = asNumber_(getProp_(key, '0'), 0);
  setProp_(key, String(cur + n));
}

function finalizeSuccess_(sliceResults) {
  const now = new Date();
  applyUrlUpdatesBatch_(sliceResults || [], now);
  updateUrlsFromRunId_(getProp_(PROPS.CURRENT_RUN_ID, ''));

  const type = getProp_(PROPS.CURRENT_RUN_TYPE) || RUN_TYPE.SCHEDULED;
  if (type === RUN_TYPE.SCHEDULED) {
    setProp_(PROPS.SCHEDULED_COMPLETE_PT, getPacificDate());
    setReserved_(0);
  }
  deleteContinuationTriggers_();
  writeRunLog_({ status: RUN_STATUS.COMPLETE });
  setSetting_(SETTINGS.LATEST_RUN_ID, getProp_(PROPS.CURRENT_RUN_ID, ''));
  try { refreshPivotsSnapshot_(); } catch (e) { console.error(e); }
  sendRunAlert_();
  clearRunState_();
  showToast_('Run complete.', 'Index Checker');
}

function finalizeAbort_(message) {
  try {
    deleteContinuationTriggers_();
    writeRunLog_({ status: RUN_STATUS.ABORTED, notes: message });
    sendAbortAlert_(message);
  } catch (err) {
    console.error(err);
  }
  clearRunState_();
}

function clearRunState_() {
  [
    PROPS.CURRENT_RUN_ID,
    PROPS.CURRENT_RUN_TYPE,
    PROPS.CURRENT_RUN_START,
    PROPS.CURRENT_RUN_ATTEMPTED,
    PROPS.CURRENT_RUN_SUCCEEDED,
    PROPS.CURRENT_RUN_FAILED,
    PROPS.CURRENT_RUN_CONTINUATIONS,
    PROPS.CURRENT_RUN_BORROWED,
    PROPS.CURRENT_RUN_NOTES,
    PROPS.ABORT_FLAG
  ].forEach(deleteProp_);
}

function deleteContinuationTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processQueue') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function resetTriggers() {
  const ui = SpreadsheetApp.getUi();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
  if (isSetupComplete_()) {
    createDailyTrigger_();
    ui.alert(
      'Daily trigger installed.\n\nFunction: onScheduledRun\nHour: ' + getRunHour_() +
      ' (' + Session.getScriptTimeZone() + ')\n\nView it in Apps Script → Triggers (clock icon). ' +
      'It does not show on the spreadsheet menu.'
    );
  } else {
    ui.alert('All triggers were removed. Run Setup and choose a property to create the daily trigger.');
  }
}

function ensureDailyTrigger_() {
  if (!isSetupComplete_()) return;
  const existing = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'onScheduledRun';
  });
  if (existing) return;
  createDailyTrigger_();
}

function createDailyTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onScheduledRun') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onScheduledRun').timeBased().atHour(getRunHour_()).everyDays(1).create();
}

function buildQueueOnly() {
  try {
    const result = buildQueue();
    SpreadsheetApp.getUi().alert(
      'Queue built.\n\nDue: ' + result.due +
      '\nQueued: ' + result.queued +
      '\nTruncated to cap: ' + result.truncated +
      '\nRemaining quota at build time: ' + result.remaining
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert(err.message || String(err));
  }
}

/**
 * 5-URL smoke test used by the setup wizard.
 */
function runSmokeTest() {
  const urls = readUrlRecords_().filter(function (u) { return u.active; }).slice(0, 5);
  if (!urls.length) {
    return { ok: false, message: 'No active URLs to test. Import a sitemap or add URLs first.' };
  }
  const results = inspectChunk_(urls.map(function (u) {
    return { url: u.url, section: u.section, tags: u.tags, urlsRow: u.row, priority: u.priority };
  }));
  const runId = 'smoke-' + newRunId_();
  appendInspectionRows_(runId, new Date(), results, urlMapByUrl_());
  incrementQuota_(results.length, RUN_TYPE.ADHOC);
  applyUrlUpdatesBatch_(results, new Date());
  setSetting_(SETTINGS.LATEST_RUN_ID, runId);
  return {
    ok: true,
    message: 'Inspected ' + results.length + ' URL(s).',
    results: results.map(function (r) {
      return {
        url: r.url,
        verdict: r.verdict || '(none)',
        coverage: r.coverageState || '',
        error: r.error || ''
      };
    })
  };
}
