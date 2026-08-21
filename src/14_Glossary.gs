/**
 * Glossary tab content — written by bootstrap, not at runtime.
 */

function writeGlossary_(sheet) {
  sheet.clear();
  sheet.setHiddenGridlines(true);
  const rows = glossaryBody_();
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1, 1, 3).merge()
    .setFontSize(18).setFontWeight('bold').setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG);
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 420);
  sheet.setColumnWidth(3, 360);
  sheet.setFrozenRows(1);

  const headerRows = [];
  rows.forEach(function (r, i) {
    if (r[0] && !r[1] && String(r[0]).indexOf('GSC Index') !== 0) headerRows.push(i + 1);
  });
  headerRows.forEach(function (r) {
    sheet.getRange(r, 1, 1, 3).merge().setFontWeight('bold').setFontSize(13)
      .setBackground(COLORS.BAND).setFontColor(COLORS.HEADER_BG);
  });
}

function glossaryBody_() {
  return [
    ['GSC Index Status Checker — Glossary', '', ''],
    ['', '', ''],
    ['This tool does not use the Google Indexing API',
      'The Indexing API only accepts JobPosting and BroadcastEvent submissions. Its metadata endpoint reports what you submitted, not whether Google has indexed a page. This sheet calls POST https://searchconsole.googleapis.com/v1/urlInspection/index:inspect',
      ''],
    ['', '', ''],
    ['Indexing verdict (indexStatusResult.verdict)', '', ''],
    ['Value', 'Meaning', 'Notes'],
    ['PASS', 'URL is indexed', 'Equivalent to Valid in Search Console'],
    ['PARTIAL', 'Indexed with issues (historical)', 'Official schema currently marks PARTIAL as reserved / no longer in use. Capture it if it appears; do not rely on it in client reporting.'],
    ['FAIL', 'Not indexed', 'Equivalent to Error / Invalid'],
    ['NEUTRAL', 'Excluded, but not necessarily an error — often a correctly canonicalised duplicate', 'Equivalent to Excluded'],
    ['VERDICT_UNSPECIFIED', 'Unknown — treat as needing manual inspection', ''],
    ['', '', ''],
    ['Coverage state (coverageState)', '', ''],
    ['Free-text strings mirroring the Search Console UI. Google adds and rewords these. Unrecognised strings must be stored as-is, never rejected.', '', ''],
    ['Value', 'Meaning', 'Action'],
    ['Submitted and indexed', 'In sitemap and indexed', 'None'],
    ['Indexed, not submitted in sitemap', 'Indexed but missing from sitemap', 'Add to sitemap'],
    ['Discovered – currently not indexed', 'Google knows the URL but has not crawled it', 'Often crawl budget or perceived low value'],
    ['Crawled – currently not indexed', 'Crawled but Google chose not to index', 'Usually a quality or duplication signal'],
    ['Duplicate without user-selected canonical', 'Google picked a canonical, you did not declare one', 'Declare a canonical'],
    ['Duplicate, Google chose different canonical than user', 'Your canonical was overridden', 'Investigate — often thin or near-identical content'],
    ['Duplicate, submitted URL not selected as canonical', 'Sitemap URL lost the canonical contest', 'Align sitemap with the chosen canonical'],
    ['Alternate page with proper canonical tag', 'Working as intended', 'None'],
    ['Excluded by \'noindex\' tag', 'Deliberately or accidentally blocked', 'Verify intent'],
    ['Blocked by robots.txt', 'Crawl blocked', 'Verify intent'],
    ['Indexed, though blocked by robots.txt', 'Indexed without being crawled', 'Usually undesirable — use noindex instead'],
    ['Page with redirect', 'URL redirects elsewhere', 'Update sitemap to the destination'],
    ['Not found (404)', 'Missing', 'Fix or remove'],
    ['Soft 404', 'Returns 200 but looks empty to Google', 'Fix response code or content'],
    ['Blocked due to unauthorized request (401)', 'Auth wall', 'Expected on staging'],
    ['Blocked due to access forbidden (403)', 'Forbidden', 'Check server config'],
    ['URL is unknown to Google', 'Never discovered', 'Add internal links and submit in sitemap'],
    ['', '', ''],
    ['Robots.txt state (robotsTxtState)', '', ''],
    ['ALLOWED', 'Crawl allowed by robots.txt', ''],
    ['DISALLOWED', 'Crawl blocked by robots.txt', ''],
    ['ROBOTS_TXT_STATE_UNSPECIFIED', 'Unknown — typically the page was not fetched, or robots.txt could not be reached', ''],
    ['', '', ''],
    ['Indexing state (indexingState)', '', ''],
    ['INDEXING_ALLOWED', 'Indexing allowed', ''],
    ['BLOCKED_BY_META_TAG', 'noindex in robots meta tag', ''],
    ['BLOCKED_BY_HTTP_HEADER', 'noindex in X-Robots-Tag header', ''],
    ['BLOCKED_BY_ROBOTS_TXT', 'Reserved, no longer in use', ''],
    ['INDEXING_STATE_UNSPECIFIED', 'Unknown indexing status', ''],
    ['', '', ''],
    ['Page fetch state (pageFetchState)', '', ''],
    ['SUCCESSFUL', 'Successful fetch', ''],
    ['SOFT_404', 'Soft 404', ''],
    ['BLOCKED_ROBOTS_TXT', 'Blocked by robots.txt', ''],
    ['NOT_FOUND', 'Not found (404)', ''],
    ['ACCESS_DENIED', 'Blocked due to unauthorized request (401)', ''],
    ['SERVER_ERROR', 'Server error (5xx)', ''],
    ['REDIRECT_ERROR', 'Redirection error', ''],
    ['ACCESS_FORBIDDEN', 'Blocked due to access forbidden (403)', ''],
    ['BLOCKED_4XX', 'Blocked due to other 4xx issue (not 403, 404)', ''],
    ['INTERNAL_CRAWL_ERROR', 'Internal crawl error', ''],
    ['INVALID_URL', 'Invalid URL', ''],
    ['PAGE_FETCH_STATE_UNSPECIFIED', 'Unknown fetch state', ''],
    ['', '', ''],
    ['Crawled as (crawledAs)', '', ''],
    ['MOBILE', 'Mobile user agent', 'Primary crawler under mobile-first indexing'],
    ['DESKTOP', 'Desktop user agent', ''],
    ['CRAWLING_USER_AGENT_UNSPECIFIED', 'Unknown user agent', ''],
    ['', '', ''],
    ['Canonical fields', '', ''],
    ['userCanonical', 'The canonical you declared', 'Absent if you did not declare one'],
    ['googleCanonical', 'The canonical Google actually selected', 'Absent if the page was not indexed'],
    ['Canonical Match', 'TRUE when both are present and equal', 'A mismatch is the single most useful diagnostic this tool surfaces. Google overriding your canonical means it does not agree with your view of the site\'s structure.'],
    ['', '', ''],
    ['Mobile usability caveat', '', ''],
    ['mobileUsabilityResult', 'Still present on the API schema but marked deprecated. The Search Console Mobile Usability report, Mobile-Friendly Test, and Mobile-Friendly Test API were retired on 1 December 2023.', 'This sheet captures Mobile Verdict if returned. Do not build reporting or client-facing commentary on it. Use Core Web Vitals / Lighthouse instead.'],
    ['', '', ''],
    ['API limits (verified 2026-08-21 against developers.google.com/webmaster-tools/limits)', '', ''],
    ['Inspections per property per day', '2,000', 'Per-site quota. Cannot be raised with extra API keys. DAILY_CAP defaults to 1,900 for headroom.'],
    ['Inspections per property per minute', '600', 'The processor sleeps between fetchAll chunks of 10 to stay under this.'],
    ['URLs per request', '1 (no batching)', 'fetchAll only parallelises HTTP, it does not change the API contract.'],
    ['Quota reset', 'Pacific time midnight', 'A 03:00 London run still draws on the previous Pacific day. The _Quota ledger is keyed on PT date via getPacificDate().'],
    ['Scope', 'Only URLs within a property the user has verified access to', 'webmasters.readonly. 403 aborts the entire run.'],
    ['Index vs live', 'API reports the version in Google\'s index, not a live test', 'Data can lag the Search Console UI. Do not treat a single check as definitive during a live incident.'],
    ['', '', ''],
    ['Apps Script limits (verified 2026-08-21 against developers.google.com/apps-script/guides/services/quotas)', '', ''],
    ['Limit', 'Consumer (gmail.com)', 'Google Workspace'],
    ['Script runtime per execution', '6 min', '6 min'],
    ['Trigger total runtime per day', '~90 min', '~6 hr'],
    ['UrlFetch calls per day', '~20,000', '~100,000'],
    ['Emails per day', '~100', '~1,500'],
    ['Cells per spreadsheet', '10 million', '10 million'],
    ['The trigger runtime cap is the real constraint on consumer accounts: 90 minutes of continuation triggers per day limits practical throughput well below 2,000 URLs. Setup warns if the account looks like a consumer account.', '', ''],
    ['At ~2,000 Data rows/day, the 10 million cell limit arrives in roughly nine months without DATA_RETENTION_DAYS purge (default 180).', '', '']
  ];
}
