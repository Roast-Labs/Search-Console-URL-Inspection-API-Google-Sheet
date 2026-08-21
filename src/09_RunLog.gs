/**
 * One row per run in Run Log.
 */

function writeRunLog_(opts) {
  opts = opts || {};
  const startMs = asNumber_(getProp_(PROPS.CURRENT_RUN_START, String(Date.now())), Date.now());
  const end = new Date();
  const start = new Date(startMs);
  const durationMin = ((end.getTime() - start.getTime()) / 60000);
  const snap = getQuotaSnapshot_();
  const attempted = asNumber_(getProp_(PROPS.CURRENT_RUN_ATTEMPTED, '0'), 0);
  const succeeded = asNumber_(getProp_(PROPS.CURRENT_RUN_SUCCEEDED, '0'), 0);
  const failed = asNumber_(getProp_(PROPS.CURRENT_RUN_FAILED, '0'), 0);
  const notes = [getProp_(PROPS.CURRENT_RUN_NOTES, ''), opts.notes || ''].filter(Boolean).join(' — ');
  const row = [
    getProp_(PROPS.CURRENT_RUN_ID, newRunId_()),
    getProp_(PROPS.CURRENT_RUN_TYPE, RUN_TYPE.SCHEDULED),
    start,
    end,
    Math.round(durationMin * 10) / 10,
    attempted,
    succeeded,
    failed,
    snap.scheduledUsed + snap.adhocUsed,
    snap.remaining,
    asNumber_(getProp_(PROPS.CURRENT_RUN_CONTINUATIONS, '0'), 0),
    getProp_(PROPS.CURRENT_RUN_BORROWED, 'false') === 'true',
    opts.status || RUN_STATUS.COMPLETE,
    notes
  ];
  appendRows_(getSheet_(SHEETS.RUN_LOG), [row]);
}
