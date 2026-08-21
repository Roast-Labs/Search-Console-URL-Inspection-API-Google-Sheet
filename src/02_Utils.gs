/**
 * Date, lock, account-type, and small helpers.
 */

function getPacificDate(date) {
  return Utilities.formatDate(date || new Date(), 'America/Los_Angeles', 'yyyy-MM-dd');
}

function nextPacificDate_(fromDate) {
  const iso = fromDate || getPacificDate();
  const parts = iso.split('-');
  const dt = new Date(Date.UTC(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]) + 1, 12, 0, 0));
  return Utilities.formatDate(dt, 'UTC', 'yyyy-MM-dd');
}

function spreadsheetToday_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const iso = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const parts = iso.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function newRunId_() {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const rand = Utilities.getUuid().split('-')[0];
  return stamp + '-' + rand;
}

function chunk_(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) {
    out.push(arr.slice(i, i + size));
  }
  return out;
}

function acquireLock_() {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(API_LIMITS.LOCK_WAIT_MS);
    return lock;
  } catch (err) {
    throw new AbortRunError_(
      'Another Index Checker run is in progress. Wait for it to finish, or use Reset Triggers if it is stuck.'
    );
  }
}

function AbortRunError_(message) {
  this.name = 'AbortRunError';
  this.message = message;
  this.stack = (new Error()).stack;
}
AbortRunError_.prototype = Object.create(Error.prototype);
AbortRunError_.prototype.constructor = AbortRunError_;

function detectAccountType_() {
  const email = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const consumer = /@(gmail|googlemail)\.com$/.test(email);
  const type = consumer ? 'consumer' : 'workspace';
  setProp_(PROPS.ACCOUNT_TYPE, type);
  return {
    email: email,
    type: type,
    isConsumer: consumer
  };
}

/**
 * Whether a URL falls inside the configured Search Console property.
 * Domain properties: host equals or is a subdomain of the registered domain.
 * URL-prefix properties: URL must start with the prefix (which should end in /).
 */
function urlBelongsToProperty_(url, propertyUrl) {
  const property = String(propertyUrl || '').trim();
  const target = String(url || '').trim();
  if (!property || !target) return false;
  if (property.indexOf('sc-domain:') === 0) {
    const domain = property.substring('sc-domain:'.length).toLowerCase();
    const host = hostFromUrl_(target);
    if (!host) return false;
    return host === domain || host.endsWith('.' + domain);
  }
  return target === property || target.indexOf(property) === 0;
}

function hostFromUrl_(url) {
  const m = String(url).match(/^https?:\/\/([^\/:?#]+)/i);
  return m ? m[1].toLowerCase() : '';
}

function pathFromUrl_(url) {
  const m = String(url).match(/^https?:\/\/[^\/]+(\/[^?#]*)?/i);
  if (!m) return '/';
  return m[1] || '/';
}

function firstPathSegment_(url) {
  const path = pathFromUrl_(url);
  const parts = path.split('/').filter(function (p) { return p; });
  return parts.length ? parts[0] : '(root)';
}

function joinList_(arr) {
  if (!arr || !arr.length) return '';
  return arr.filter(function (v) { return v; }).join(', ');
}

function canonicalMatch_(googleCanonical, userCanonical) {
  const g = String(googleCanonical || '').trim();
  const u = String(userCanonical || '').trim();
  if (!g && !u) return '';
  return g === u;
}

function computeNextDue_(frequency, lastChecked, today) {
  const freq = String(frequency || '');
  if (freq === FREQUENCY.ADHOC) return '';
  const base = lastChecked instanceof Date && !isNaN(lastChecked.getTime())
    ? new Date(lastChecked.getTime())
    : (today || spreadsheetToday_());
  const d = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  if (!lastChecked) return d;
  if (freq === FREQUENCY.DAILY) {
    d.setDate(d.getDate() + 1);
    return d;
  }
  if (freq === FREQUENCY.WEEKLY) {
    d.setDate(d.getDate() + 7);
    return d;
  }
  if (freq === FREQUENCY.MONTHLY) {
    d.setMonth(d.getMonth() + 1);
    return d;
  }
  return d;
}

function asBool_(v) {
  return v === true || v === 'TRUE' || v === 'true' || v === 1 || v === '1';
}

function asNumber_(v, fallback) {
  const n = Number(v);
  return isFinite(n) ? n : fallback;
}

function escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function showToast_(msg, title) {
  SpreadsheetApp.getActive().toast(msg, title || 'Index Checker', 8);
}

function oauthHeaders_() {
  return {
    Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    'Content-Type': 'application/json'
  };
}

function gscFetch_(url, opt) {
  const options = {
    method: (opt && opt.method) || 'get',
    headers: oauthHeaders_(),
    muteHttpExceptions: true,
    followRedirects: true
  };
  if (opt && opt.payload) {
    options.payload = typeof opt.payload === 'string' ? opt.payload : JSON.stringify(opt.payload);
    options.method = options.method || 'post';
  }
  return UrlFetchApp.fetch(url, options);
}

function refreshTodayPtCell_() {
  try {
    setSetting_(SETTINGS.TODAY_PT, getPacificDate());
  } catch (err) {
    // Named range may not exist before bootstrap.
  }
}
