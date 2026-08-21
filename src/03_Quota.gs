/**
 * Pacific-date quota ledger.
 *
 * Search Console URL Inspection quota resets at Pacific midnight.
 * A 03:00 London run still draws on the previous Pacific day's allowance,
 * so this ledger is keyed on PT date, never local date.
 */

function ensureQuotaRow_(ptDate) {
  const date = ptDate || getPacificDate();
  refreshTodayPtCell_();
  const sheet = getSheet_(SHEETS.QUOTA);
  const row = findQuotaRow_(sheet, date);
  if (row) return readQuotaRow_(sheet, row);

  const last = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(last + 1, 1, 1, 4).setValues([[date, 0, 0, 0]]);
  sheet.getRange(last + 1, 1).setNumberFormat('yyyy-mm-dd');
  return readQuotaRow_(sheet, last + 1);
}

function findQuotaRow_(sheet, ptDate) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const dates = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < dates.length; i++) {
    if (String(dates[i][0]).trim() === ptDate) return i + 2;
  }
  return 0;
}

function readQuotaRow_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, 5).getValues()[0];
  const cap = getDailyCap_();
  const scheduled = asNumber_(values[1], 0);
  const adhoc = asNumber_(values[2], 0);
  const reserved = asNumber_(values[3], 0);
  return {
    row: row,
    date: values[0],
    scheduledUsed: scheduled,
    adhocUsed: adhoc,
    reserved: reserved,
    remaining: cap - scheduled - adhoc - reserved,
    cap: cap
  };
}

function getQuotaSnapshot_(ptDate) {
  const date = ptDate || getPacificDate();
  const sheet = getSheet_(SHEETS.QUOTA);
  let row = findQuotaRow_(sheet, date);
  if (!row) {
    ensureQuotaRow_(date);
    row = findQuotaRow_(sheet, date);
  }
  return readQuotaRow_(sheet, row);
}

/**
 * Keep Reserved equal to the number of scheduled URLs still due today,
 * unless today's scheduled run has already finished.
 */
function refreshReservation_(ptDate) {
  const date = ptDate || getPacificDate();
  const sheet = getSheet_(SHEETS.QUOTA);
  let row = findQuotaRow_(sheet, date);
  if (!row) return;

  if (getProp_(PROPS.SCHEDULED_COMPLETE_PT) === date) {
    sheet.getRange(row, 4).setValue(0);
    return;
  }

  const due = countScheduledDue_();
  const values = sheet.getRange(row, 1, 1, 4).getValues()[0];
  const scheduled = asNumber_(values[1], 0);
  const adhoc = asNumber_(values[2], 0);
  const cap = getDailyCap_();
  const maxReserve = Math.max(0, cap - scheduled - adhoc);
  sheet.getRange(row, 4).setValue(Math.min(due, maxReserve));
}

function incrementQuota_(count, runType, ptDate) {
  if (!count) return getQuotaSnapshot_(ptDate);
  const date = ptDate || getPacificDate();
  const sheet = getSheet_(SHEETS.QUOTA);
  let row = findQuotaRow_(sheet, date);
  if (!row) {
    ensureQuotaRow_(date);
    row = findQuotaRow_(sheet, date);
  }
  const col = runType === RUN_TYPE.ADHOC ? 3 : 2;
  const cell = sheet.getRange(row, col);
  cell.setValue(asNumber_(cell.getValue(), 0) + count);

  const reservedCell = sheet.getRange(row, 4);
  const reserved = asNumber_(reservedCell.getValue(), 0);
  if (runType !== RUN_TYPE.ADHOC) {
    reservedCell.setValue(Math.max(0, reserved - count));
  }
  SpreadsheetApp.flush();
  return readQuotaRow_(sheet, row);
}

function setReserved_(amount, ptDate) {
  const date = ptDate || getPacificDate();
  const sheet = getSheet_(SHEETS.QUOTA);
  let row = findQuotaRow_(sheet, date);
  if (!row) {
    ensureQuotaRow_(date);
    row = findQuotaRow_(sheet, date);
  }
  sheet.getRange(row, 4).setValue(Math.max(0, amount));
}

function releaseReservation_(amount, ptDate) {
  const snap = getQuotaSnapshot_(ptDate);
  const next = Math.max(0, snap.reserved - amount);
  setReserved_(next, ptDate);
  return next;
}

function countScheduledDue_() {
  const urls = readUrlRecords_();
  const today = spreadsheetToday_();
  let n = 0;
  urls.forEach(function (u) {
    if (!u.active) return;
    if (u.frequency === FREQUENCY.ADHOC) return;
    if (!u.nextDue) {
      n++;
      return;
    }
    if (u.nextDue.getTime() <= today.getTime()) n++;
  });
  return n;
}

function viewQuotaPayload_() {
  ensureQuotaRow_();
  refreshReservation_();
  const snap = getQuotaSnapshot_();
  return {
    datePt: getPacificDate(),
    cap: snap.cap,
    apiLimit: API_LIMITS.QPD_PER_SITE,
    scheduledUsed: snap.scheduledUsed,
    adhocUsed: snap.adhocUsed,
    reserved: snap.reserved,
    remaining: snap.remaining,
    used: snap.scheduledUsed + snap.adhocUsed
  };
}
