/**
 * Email alerts. Blank ALERT_EMAIL disables sending.
 */

function sendRunAlert_() {
  const to = String(getSetting_(SETTINGS.ALERT_EMAIL) || '').trim();
  if (!to) return;
  const runId = getProp_(PROPS.CURRENT_RUN_ID, '');
  const changed = changedRowsForRun_(runId);
  if (isTruthySetting_(SETTINGS.ALERT_ON_CHANGE_ONLY) && !changed.length) return;

  const attempted = getProp_(PROPS.CURRENT_RUN_ATTEMPTED, '0');
  const succeeded = getProp_(PROPS.CURRENT_RUN_SUCCEEDED, '0');
  const failed = getProp_(PROPS.CURRENT_RUN_FAILED, '0');
  const snap = getQuotaSnapshot_();
  const property = getPropertyUrl_();

  let body = 'Index Checker run finished for ' + property + '\n\n';
  body += 'Run ID: ' + runId + '\n';
  body += 'Attempted: ' + attempted + '  Succeeded: ' + succeeded + '  Failed: ' + failed + '\n';
  body += 'Quota used today (PT ' + getPacificDate() + '): ' + (snap.scheduledUsed + snap.adhocUsed) +
    ' / ' + snap.cap + '  Remaining: ' + snap.remaining + '\n';
  body += 'Statuses changed: ' + changed.length + '\n\n';
  if (changed.length) {
    body += 'Changed URLs\n';
    changed.slice(0, 50).forEach(function (c) {
      body += '- ' + c.url + ' → ' + (c.verdict || c.coverage || '(unknown)') + '\n';
    });
    if (changed.length > 50) body += '… and ' + (changed.length - 50) + ' more\n';
  }
  body += '\nSheet: ' + SpreadsheetApp.getActive().getUrl() + '\n';
  MailApp.sendEmail(to, 'Index Checker: ' + property + ' (' + changed.length + ' changed)', body);
}

function sendAbortAlert_(message) {
  const to = String(getSetting_(SETTINGS.ALERT_EMAIL) || '').trim();
  if (!to) return;
  const property = getPropertyUrl_() || '(no property)';
  MailApp.sendEmail(
    to,
    'Index Checker ABORTED: ' + property,
    'The Index Checker run was aborted.\n\n' + message + '\n\nSheet: ' + SpreadsheetApp.getActive().getUrl()
  );
}
