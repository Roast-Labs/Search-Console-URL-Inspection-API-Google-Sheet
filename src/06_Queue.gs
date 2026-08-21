/**
 * Queue builder: due URLs, priority sort, truncate to remaining quota.
 */

function buildQueue() {
  const lock = acquireLock_();
  try {
    return buildQueueLocked_(false);
  } finally {
    lock.releaseLock();
  }
}

function buildQueueLocked_(skipReservationRefresh) {
  const queueSheet = getSheet_(SHEETS.QUEUE);
  clearDataRows_(queueSheet);

  const propertyUrl = getPropertyUrl_();
  const pt = getPacificDate();
  ensureQuotaRow_(pt);
  setReserved_(0, pt);

  const snap = getQuotaSnapshot_(pt);
  const remaining = Math.max(0, snap.remaining);
  const today = spreadsheetToday_();
  const due = [];

  readUrlRecords_().forEach(function (u) {
    if (!u.active) return;
    if (u.frequency === FREQUENCY.ADHOC) return;
    if (u.nextDue && u.nextDue.getTime() > today.getTime()) return;
    due.push(u);
  });

  readDeferredDue_(pt).forEach(function (d) {
    const already = due.some(function (u) { return u.url === d.url; });
    if (already) {
      due.forEach(function (u) {
        if (u.url === d.url) u.priority = Math.min(u.priority, 1);
      });
      return;
    }
    const map = urlMapByUrl_();
    const existing = map[d.url];
    due.push({
      row: existing ? existing.row : 0,
      url: d.url,
      active: true,
      frequency: existing ? existing.frequency : FREQUENCY.DAILY,
      section: existing ? existing.section : '',
      tags: existing ? existing.tags : '',
      priority: 1,
      nextDue: today,
      lastVerdict: existing ? existing.lastVerdict : ''
    });
  });

  due.sort(function (a, b) {
    const pa = asNumber_(a.priority, 3);
    const pb = asNumber_(b.priority, 3);
    if (pa !== pb) return pa - pb;
    const da = a.nextDue ? a.nextDue.getTime() : 0;
    const db = b.nextDue ? b.nextDue.getTime() : 0;
    return da - db;
  });

  const truncated = due.slice(0, remaining);
  const rows = truncated.map(function (u) {
    return [
      u.url,
      asNumber_(u.priority, 3),
      u.nextDue || today,
      u.section || '',
      u.tags || '',
      u.frequency || '',
      'Pending',
      u.row || 0
    ];
  });
  if (rows.length) appendRows_(queueSheet, rows);

  setReserved_(rows.length, pt);
  setProp_(PROPS.QUEUE_CURSOR, '2');
  return {
    due: due.length,
    queued: rows.length,
    truncated: Math.max(0, due.length - rows.length),
    remaining: remaining
  };
}

function buildAdhocQueue_(items, skipLock) {
  const run = function () {
    const queueSheet = getSheet_(SHEETS.QUEUE);
    clearDataRows_(queueSheet);
    const rows = items.map(function (u) {
      return [
        u.url,
        asNumber_(u.priority, 1),
        '',
        u.section || '',
        u.tags || '',
        FREQUENCY.ADHOC,
        'Pending',
        u.row || 0
      ];
    });
    if (rows.length) appendRows_(queueSheet, rows);
    setProp_(PROPS.QUEUE_CURSOR, '2');
    return rows.length;
  };
  if (skipLock) return run();
  const lock = acquireLock_();
  try {
    return run();
  } finally {
    lock.releaseLock();
  }
}

function readQueuePending_(cursor, limit) {
  const sheet = getSheet_(SHEETS.QUEUE);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const start = Math.max(Number(cursor) || 2, 2);
  if (start > last) return [];
  const end = Math.min(last, start + limit - 1);
  const num = end - start + 1;
  const values = sheet.getRange(start, 1, num, QUEUE_HEADERS.length).getValues();
  return values.map(function (row, i) {
    return {
      queueRow: start + i,
      url: String(row[0] || '').trim(),
      priority: asNumber_(row[1], 3),
      nextDue: row[2],
      section: String(row[3] || ''),
      tags: String(row[4] || ''),
      frequency: String(row[5] || ''),
      status: String(row[6] || ''),
      urlsRow: asNumber_(row[7], 0)
    };
  }).filter(function (q) { return q.url && q.status !== 'Done'; });
}

function markQueueRows_(queueRows, status) {
  if (!queueRows.length) return;
  const sheet = getSheet_(SHEETS.QUEUE);
  queueRows.forEach(function (r) {
    sheet.getRange(r, 7).setValue(status);
  });
}

function clearDataRows_(sheet) {
  const last = sheet.getLastRow();
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, sheet.getMaxColumns()).clearContent();
  }
}

function queueLength_() {
  const sheet = getSheet_(SHEETS.QUEUE);
  return Math.max(0, sheet.getLastRow() - 1);
}

function readDeferredDue_(ptDate) {
  const sheet = getSheet_(SHEETS.DEFERRED);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, DEFERRED_HEADERS.length).getValues();
  const due = [];
  const keep = [];
  values.forEach(function (row) {
    const url = String(row[0] || '').trim();
    if (!url) return;
    const dueDate = row[2] instanceof Date
      ? Utilities.formatDate(row[2], 'America/Los_Angeles', 'yyyy-MM-dd')
      : String(row[2] || '');
    if (dueDate && dueDate <= ptDate) {
      due.push({ url: url, priority: asNumber_(row[1], 1), dueDate: dueDate });
    } else {
      keep.push(row);
    }
  });
  clearDataRows_(sheet);
  if (keep.length) appendRows_(sheet, keep);
  return due;
}

function deferUrlsUntilTomorrow_(urls, reason) {
  if (!urls || !urls.length) return;
  const sheet = getSheet_(SHEETS.DEFERRED);
  const tomorrow = nextPacificDate_();
  const rows = urls.map(function (url) {
    return [url, 1, tomorrow, reason || '429 requeue'];
  });
  appendRows_(sheet, rows);
}
