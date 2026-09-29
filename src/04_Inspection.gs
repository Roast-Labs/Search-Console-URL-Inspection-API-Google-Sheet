/**
 * URL Inspection API client.
 * Uses UrlFetchApp + ScriptApp.getOAuthToken() — not the Advanced Service.
 */

function inspectChunk_(items) {
  const propertyUrl = getPropertyUrl_();
  const language = getLanguageCode_();
  if (!propertyUrl) {
    throw new AbortRunError_('No Search Console property configured. Run Setup first.');
  }

  const pending = items.slice();
  const finished = [];
  let attempt = 0;

  while (pending.length && attempt < API_LIMITS.RETRY_ATTEMPTS) {
    if (attempt > 0) {
      Utilities.sleep(Math.pow(2, attempt) * 1000);
    }
    const requests = pending.map(function (item) {
      return buildInspectRequest_(item.url, propertyUrl, language);
    });
    const responses = UrlFetchApp.fetchAll(requests);
    const retry = [];
    let abortResult = null;

    responses.forEach(function (resp, i) {
      const item = pending[i];
      const parsed = parseInspectResponse_(resp, item);
      if (parsed.abort) {
        abortResult = parsed;
        finished.push(parsed);
        return;
      }
      if (parsed.retry && attempt < API_LIMITS.RETRY_ATTEMPTS - 1) {
        retry.push(item);
        return;
      }
      if (parsed.retry && parsed.httpStatus === 429) {
        parsed.requeueTomorrow = true;
        parsed.error = parsed.error || 'Rate limited (429) after 3 attempts — requeued for tomorrow at Priority 1.';
      }
      finished.push(parsed);
    });

    if (abortResult) return finished;
    pending.length = 0;
    retry.forEach(function (r) { pending.push(r); });
    attempt++;
  }

  pending.forEach(function (item) {
    finished.push({
      url: item.url,
      section: item.section,
      tags: item.tags,
      urlsRow: item.urlsRow,
      priority: item.priority,
      httpStatus: 429,
      requeueTomorrow: true,
      error: 'Still pending after retries — requeued for tomorrow.',
      verdict: '',
      coverageState: '',
      robotsTxtState: '',
      indexingState: '',
      googleCanonical: '',
      userCanonical: '',
      lastCrawlTime: '',
      crawledAs: '',
      pageFetchState: '',
      sitemaps: '',
      referringUrls: '',
      mobileVerdict: '',
      richResultsVerdict: ''
    });
  });

  return finished;
}

function buildInspectRequest_(inspectionUrl, siteUrl, languageCode) {
  return {
    url: ENDPOINTS.INSPECT,
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify({
      inspectionUrl: inspectionUrl,
      siteUrl: siteUrl,
      languageCode: languageCode
    }),
    muteHttpExceptions: true,
    followRedirects: true
  };
}

function parseInspectResponse_(resp, item) {
  const httpStatus = resp.getResponseCode();
  const raw = resp.getContentText() || '';
  let body = {};
  try {
    body = raw ? JSON.parse(raw) : {};
  } catch (err) {
    body = {};
  }

  const base = {
    url: item.url,
    section: item.section || '',
    tags: item.tags || '',
    urlsRow: item.urlsRow,
    priority: item.priority,
    httpStatus: httpStatus,
    abort: false,
    retry: false,
    requeueTomorrow: false,
    error: '',
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
    richResultsVerdict: '',
    inspectionResultLink: ''
  };

  if (httpStatus === 403) {
    const apiMessage = extractApiError_(body, raw);
    base.abort = true;
    base.error = 'Search Console returned 403 — the authorised account almost certainly does not have access to ' +
      getPropertyUrl_() + '. Aborting the run so we do not burn the daily quota. ' + apiMessage;
    return base;
  }

  if (httpStatus === 429) {
    base.retry = true;
    base.error = extractApiError_(body, 'HTTP 429');
    return base;
  }

  if (httpStatus >= 500) {
    base.retry = true;
    base.error = extractApiError_(body, 'HTTP ' + httpStatus);
    return base;
  }

  // 404 from the inspect API is a valid inspection result, not a failure.
  if (httpStatus === 404) {
    fillInspectionFields_(base, body);
    if (!base.coverageState) base.coverageState = 'Not found (404)';
    return base;
  }

  if (httpStatus >= 400) {
    base.error = extractApiError_(body, 'HTTP ' + httpStatus);
    return base;
  }

  fillInspectionFields_(base, body);
  return base;
}

function fillInspectionFields_(base, body) {
  const result = (body && body.inspectionResult) || {};
  const idx = result.indexStatusResult || {};
  const mobile = result.mobileUsabilityResult || {};
  const rich = result.richResultsResult || {};

  base.verdict = idx.verdict || '';
  base.coverageState = idx.coverageState || '';
  base.robotsTxtState = idx.robotsTxtState || '';
  base.indexingState = idx.indexingState || '';
  base.googleCanonical = idx.googleCanonical || '';
  base.userCanonical = idx.userCanonical || '';
  base.canonicalMatch = canonicalMatch_(base.googleCanonical, base.userCanonical);
  base.lastCrawlTime = parseCrawlTime_(idx.lastCrawlTime);
  base.crawledAs = idx.crawledAs || '';
  base.pageFetchState = idx.pageFetchState || '';
  base.sitemaps = joinList_(idx.sitemap);
  base.referringUrls = joinList_(idx.referringUrls);
  base.mobileVerdict = mobile.verdict || '';
  base.richResultsVerdict = rich.verdict || '';
  base.inspectionResultLink = result.inspectionResultLink || '';
}

function parseCrawlTime_(value) {
  if (!value) return '';
  const d = new Date(value);
  return isNaN(d.getTime()) ? String(value) : d;
}

function extractApiError_(body, fallback) {
  if (body && body.error) {
    if (body.error.message) return body.error.message;
    if (typeof body.error === 'string') return body.error;
  }
  return fallback || 'Unknown API error';
}

function inspectOne_(url) {
  const results = inspectChunk_([{ url: url, section: '', tags: '', urlsRow: 0, priority: 3 }]);
  return results[0];
}

function mapSiteEntries_(entries) {
  return (entries || []).map(function (e) {
    return {
      siteUrl: e.siteUrl,
      permissionLevel: e.permissionLevel || ''
    };
  });
}

function isGscApiDisabledError_(text) {
  return /has not been used in project|ACCESS_NOT_CONFIGURED|accessNotConfigured|API has not been enabled|it is disabled|accessNotConfigured/i.test(String(text || ''));
}

function gcpLinkInstructions_() {
  return (
    'This copy uses your own Apps Script project, which starts on Google\'s Default Cloud project. ' +
    'The Search Console API has to be enabled on a Cloud project you own — not the template owner\'s.\n\n' +
    'Once per copy:\n' +
    '1. In Google Cloud Console (same Google account), create or pick a project.\n' +
    '2. Enable "Google Search Console API".\n' +
    '3. Configure the OAuth consent screen (Internal, or External + Testing with your email as a test user).\n' +
    '4. Copy the project NUMBER (digits only).\n' +
    '5. In this spreadsheet: Extensions → Apps Script → Project Settings (gear) → Google Cloud Platform (GCP) Project → Change project → paste YOUR number.\n' +
    '6. Close this dialog, reload the sheet, run Index Checker → Setup again, and authorise.'
  );
}

function listSearchConsoleSites_() {
  const resp = gscFetch_(ENDPOINTS.SITES);
  const code = resp.getResponseCode();
  const raw = resp.getContentText() || '';
  let body = {};
  try {
    body = raw ? JSON.parse(raw) : {};
  } catch (parseErr) {
    body = {};
  }
  const apiMessage = extractApiError_(body, raw);
  if (code === 403) {
    throw new Error(
      (isGscApiDisabledError_(apiMessage + ' ' + raw)
        ? 'Cannot list Search Console properties — the API is not enabled on this script\'s Cloud project.\n\n'
        : 'Cannot list Search Console properties (403). ' + (apiMessage ? apiMessage + '\n\n' : '')) +
      gcpLinkInstructions_()
    );
  }
  if (code >= 400) {
    throw new Error('sites.list failed (' + code + '): ' + apiMessage);
  }
  return mapSiteEntries_(body.siteEntry);
}

function listSubmittedSitemaps_(siteUrl) {
  const encoded = encodeURIComponent(siteUrl);
  const resp = gscFetch_(ENDPOINTS.SITEMAPS + encoded + '/sitemaps');
  const code = resp.getResponseCode();
  const raw = resp.getContentText() || '';
  let body = {};
  try {
    body = raw ? JSON.parse(raw) : {};
  } catch (err) {
    body = {};
  }
  if (code >= 400) {
    throw new Error('sitemaps.list failed (' + code + '): ' + extractApiError_(body, raw));
  }
  return (body.sitemap || []).map(function (s) {
    return {
      path: s.path,
      lastSubmitted: s.lastSubmitted,
      lastDownloaded: s.lastDownloaded,
      isPending: s.isPending,
      isSitemapsIndex: s.isSitemapsIndex,
      type: s.type,
      errors: s.errors,
      warnings: s.warnings
    };
  });
}
