/**
 * GSC Index Status Checker — constants.
 *
 * Quota figures verified 2026-08-21 against:
 *   https://developers.google.com/webmaster-tools/limits
 *   https://developers.google.com/apps-script/guides/services/quotas
 *   https://developers.google.com/webmaster-tools/v1/urlInspection.index/inspect
 * Google changes these without much notice — re-check before treating them as gospel.
 */

var SHEETS = {
  SETTINGS: 'Settings',
  URLS: 'URLs',
  DATA: 'Data',
  RUN_LOG: 'Run Log',
  PIVOTS: 'Pivots',
  DASHBOARD: 'Dashboard',
  GLOSSARY: 'Glossary',
  QUEUE: '_Queue',
  QUOTA: '_Quota',
  IMPORT: '_Import',
  DEFERRED: '_Deferred'
};

var SETTINGS = {
  PROPERTY_URL: 'PROPERTY_URL',
  DAILY_CAP: 'DAILY_CAP',
  BATCH_SIZE: 'BATCH_SIZE',
  RUN_HOUR: 'RUN_HOUR',
  ALERT_EMAIL: 'ALERT_EMAIL',
  ALERT_ON_CHANGE_ONLY: 'ALERT_ON_CHANGE_ONLY',
  DATA_RETENTION_DAYS: 'DATA_RETENTION_DAYS',
  LANGUAGE_CODE: 'LANGUAGE_CODE',
  ADHOC_BORROW_ALLOWED: 'ADHOC_BORROW_ALLOWED',
  TODAY_PT: 'TODAY_PT',
  LATEST_RUN_ID: 'LATEST_RUN_ID',
  SETUP_COMPLETE_CELL: 'SETUP_COMPLETE_CELL'
};

var DEFAULTS = {
  DAILY_CAP: 1900,
  BATCH_SIZE: 60,
  RUN_HOUR: 3,
  DATA_RETENTION_DAYS: 180,
  LANGUAGE_CODE: 'en-GB',
  ALERT_ON_CHANGE_ONLY: true,
  ADHOC_BORROW_ALLOWED: true
};

/** Official URL Inspection API per-site caps (Aug 2026). */
var API_LIMITS = {
  QPD_PER_SITE: 2000,
  QPM_PER_SITE: 600,
  FETCH_CHUNK: 10,
  /** Sleep between fetchAll chunks so we stay under 600 QPM. */
  CHUNK_SLEEP_MS: 1100,
  MAX_RUNTIME_MS: 4.5 * 60 * 1000,
  LOCK_WAIT_MS: 30000,
  RETRY_ATTEMPTS: 3,
  SITEMAP_MAX_DEPTH: 3
};

var ENDPOINTS = {
  INSPECT: 'https://searchconsole.googleapis.com/v1/urlInspection/index:inspect',
  SITES: 'https://www.googleapis.com/webmasters/v3/sites',
  SITEMAPS: 'https://www.googleapis.com/webmasters/v3/sites/'
};

var SCOPE = 'https://www.googleapis.com/auth/webmasters.readonly';

var PROPS = {
  SETUP_COMPLETE: 'SETUP_COMPLETE',
  QUEUE_CURSOR: 'QUEUE_CURSOR',
  CURRENT_RUN_ID: 'CURRENT_RUN_ID',
  CURRENT_RUN_TYPE: 'CURRENT_RUN_TYPE',
  CURRENT_RUN_START: 'CURRENT_RUN_START',
  CURRENT_RUN_ATTEMPTED: 'CURRENT_RUN_ATTEMPTED',
  CURRENT_RUN_SUCCEEDED: 'CURRENT_RUN_SUCCEEDED',
  CURRENT_RUN_FAILED: 'CURRENT_RUN_FAILED',
  CURRENT_RUN_CONTINUATIONS: 'CURRENT_RUN_CONTINUATIONS',
  CURRENT_RUN_BORROWED: 'CURRENT_RUN_BORROWED',
  CURRENT_RUN_NOTES: 'CURRENT_RUN_NOTES',
  SCHEDULED_COMPLETE_PT: 'SCHEDULED_COMPLETE_PT',
  ACCOUNT_TYPE: 'ACCOUNT_TYPE',
  ABORT_FLAG: 'ABORT_FLAG'
};

var RUN_TYPE = {
  SCHEDULED: 'Scheduled',
  ADHOC: 'Ad-hoc',
  CONTINUATION: 'Continuation'
};

var RUN_STATUS = {
  COMPLETE: 'Complete',
  PARTIAL: 'Partial',
  ABORTED: 'Aborted'
};

var FREQUENCY = {
  DAILY: 'Daily',
  WEEKLY: 'Weekly',
  MONTHLY: 'Monthly',
  ADHOC: 'Ad-hoc'
};

/** URLs sheet columns (1-based). */
var URL_COL = {
  URL: 1,
  ACTIVE: 2,
  FREQUENCY: 3,
  SECTION: 4,
  TAGS: 5,
  PRIORITY: 6,
  LAST_CHECKED: 7,
  NEXT_DUE: 8,
  LAST_VERDICT: 9,
  ORPHANED: 10
};

var URL_HEADERS = [
  'URL',
  'Active',
  'Frequency',
  'Section',
  'Tags',
  'Priority',
  'Last Checked',
  'Next Due',
  'Last Verdict',
  'Orphaned'
];

/** Data sheet columns (1-based). */
var DATA_COL = {
  RUN_ID: 1,
  TIMESTAMP: 2,
  URL: 3,
  SECTION: 4,
  TAGS: 5,
  VERDICT: 6,
  COVERAGE: 7,
  ROBOTS: 8,
  INDEXING_STATE: 9,
  GOOGLE_CANONICAL: 10,
  USER_CANONICAL: 11,
  CANONICAL_MATCH: 12,
  LAST_CRAWL: 13,
  CRAWLED_AS: 14,
  PAGE_FETCH: 15,
  SITEMAPS: 16,
  REFERRING: 17,
  MOBILE: 18,
  RICH_RESULTS: 19,
  HTTP_STATUS: 20,
  STATUS_CHANGED: 21,
  ERROR: 22
};

var DATA_HEADERS = [
  'Run ID',
  'Timestamp',
  'URL',
  'Section',
  'Tags',
  'Indexing Verdict',
  'Coverage State',
  'Robots.txt State',
  'Indexing State',
  'Google Canonical',
  'User Canonical',
  'Canonical Match',
  'Last Crawl Time',
  'Crawled As',
  'Page Fetch State',
  'Sitemaps',
  'Referring URLs',
  'Mobile Verdict',
  'Rich Results Verdict',
  'HTTP Status',
  'Status Changed',
  'Error'
];

var QUEUE_HEADERS = ['URL', 'Priority', 'Next Due', 'Section', 'Tags', 'Frequency', 'Status', 'Urls Row'];

var QUOTA_HEADERS = [
  'Date (PT)',
  'Scheduled Used',
  'Ad-hoc Used',
  'Reserved',
  'Remaining'
];

var RUN_LOG_HEADERS = [
  'Run ID',
  'Type',
  'Start',
  'End',
  'Duration',
  'Attempted',
  'Succeeded',
  'Failed',
  'Quota Used',
  'Quota Remaining After',
  'Continuations',
  'Borrowed From Scheduled',
  'Status',
  'Notes'
];

var DEFERRED_HEADERS = ['URL', 'Priority', 'Due Date (PT)', 'Reason'];

var IMPORT_HEADERS = ['URL', 'Lastmod', 'Section', 'Include'];

var COLORS = {
  HEADER_BG: '#174ea6',
  HEADER_FG: '#ffffff',
  ACCENT: '#1a73e8',
  HIDDEN_TAB: '#666666',
  BAND: '#e8f0fe',
  WARN: '#fbbc04',
  ERROR: '#ea4335',
  OK: '#34a853',
  SOFT: '#f8f9fa'
};
