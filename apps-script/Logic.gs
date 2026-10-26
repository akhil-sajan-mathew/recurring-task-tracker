/**
 * Pure date / recurrence / SLA logic. No Apps Script services are used in this file,
 * so it behaves identically in Apps Script and in the local Node test harness.
 * All Date objects are in the script time zone (appsscript.json → Asia/Kolkata).
 */

const DOW_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function pad2(n) { return (n < 10 ? '0' : '') + n; }
function ymd(d) { return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
function hm(d) { return pad2(d.getHours()) + ':' + pad2(d.getMinutes()); }
function fmtDateTime(d) { return d instanceof Date ? ymd(d) + ' ' + hm(d) : ''; }
function fmtFriendly(d) {
  return d instanceof Date ? DOW_NAMES[d.getDay()] + ' ' + d.getDate() + ' ' + MONTH_NAMES[d.getMonth()] + ', ' + hm(d) : '';
}

/** "HH:MM" string (or a Sheets time value) → {h, m}. */
function parseHM(v) {
  if (v instanceof Date) return { h: v.getHours(), m: v.getMinutes() };
  const m = String(v).trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!m || +m[1] > 23 || +m[2] > 59) throw new Error('Bad time "' + v + '" – expected HH:MM');
  return { h: +m[1], m: +m[2] };
}

/** Date, "YYYY-MM-DD" or "YYYY-MM-DD HH:MM" → Date (local time). Returns null for blanks. */
function parseDateTime(v) {
  if (v instanceof Date) return new Date(v.getTime());
  const s = String(v == null ? '' : v).trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2}))?$/);
  if (!m) throw new Error('Bad date "' + s + '" – expected YYYY-MM-DD or YYYY-MM-DD HH:MM');
  return new Date(+m[1], +m[2] - 1, +m[3], m[4] ? +m[4] : 0, m[5] ? +m[5] : 0);
}

function dayStart(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n, d.getHours(), d.getMinutes(), d.getSeconds());
}
function atTime(d, t) { return new Date(d.getFullYear(), d.getMonth(), d.getDate(), t.h, t.m); }
function addMinutes(d, n) { return new Date(d.getTime() + n * 60000); }

function isWorkingDay(d, cfg) { return cfg.workDays.has(d.getDay()) && !cfg.holidays.has(ymd(d)); }

/** The same day if it is a working day, otherwise the next working day (at 00:00). */
function nextWorkingDay(d, cfg) {
  let x = dayStart(d);
  for (let i = 0; i < 370; i++) {
    if (isWorkingDay(x, cfg)) return x;
    x = addDays(x, 1);
  }
  throw new Error('No working day within a year – check WORK_DAYS / Holidays');
}

/** Move n working days forward (n > 0) or backward (n < 0) from d. Returns a day at 00:00. */
function addWorkingDays(d, n, cfg) {
  let x = dayStart(d);
  const step = n < 0 ? -1 : 1;
  let left = Math.abs(n);
  let guard = 0;
  while (left > 0) {
    x = addDays(x, step);
    if (isWorkingDay(x, cfg)) left--;
    if (++guard > 3700) throw new Error('addWorkingDays ran away – check WORK_DAYS / Holidays');
  }
  return x;
}

/** Add working hours, pausing outside WORK_START–WORK_END and on non-working days. */
function addWorkingHours(start, hours, cfg) {
  let remaining = hours * 3600000;
  let t = new Date(start.getTime());
  for (let i = 0; i < 400; i++) {
    const open = atTime(t, cfg.workStart);
    const close = atTime(t, cfg.workEnd);
    if (!isWorkingDay(t, cfg) || t >= close) {
      t = atTime(nextWorkingDay(addDays(dayStart(t), 1), cfg), cfg.workStart);
      continue;
    }
    if (t < open) t = open;
    const available = close - t;
    if (remaining <= available) return new Date(t.getTime() + remaining);
    remaining -= available;
    t = atTime(nextWorkingDay(addDays(dayStart(t), 1), cfg), cfg.workStart);
  }
  throw new Error('addWorkingHours ran away');
}

function cycleId(taskId, nominal) { return taskId + '|' + ymd(nominal); }

/**
 * Every cycle of a task that should exist by `now` (its Open At has passed).
 * FIXED basis: nominal due dates are Start + k × CycleDays and never drift with late completion.
 * Each cycle: {id, cycleNo, nominal, openAt, dueAt}.
 */
function expectedCycles(task, now, cfg) {
  const out = [];
  const start = dayStart(task.startDate);
  const floor = Math.max(start.getTime(), addDays(dayStart(now), -cfg.backfillDays).getTime());
  const dueHM = task.dueTime || (task.freq === 'DAILY' ? cfg.dailyDue : cfg.portalDue);

  if (task.freq === 'DAILY') {
    let n = 0;
    for (let d = start; d <= now; d = addDays(d, 1)) {
      if (!isWorkingDay(d, cfg)) continue;
      const openAt = atTime(d, cfg.workStart);
      if (d.getTime() >= floor && openAt <= now) {
        out.push({ id: cycleId(task.id, d), cycleNo: n, nominal: d, openAt: openAt, dueAt: atTime(d, dueHM) });
      }
      n++;
    }
    return out;
  }

  const N = task.cycleDays;
  // Round because the day difference is an exact multiple of 24h in IST (no DST).
  let k = Math.max(0, Math.ceil(Math.round((floor - start.getTime()) / 86400000) / N));
  for (;; k++) {
    const nominal = addDays(start, k * N);
    const dueDay = nextWorkingDay(nominal, cfg);
    const openDay = task.leadDays > 0 ? addWorkingDays(dueDay, -task.leadDays, cfg) : dueDay;
    const openAt = atTime(openDay, cfg.workStart);
    if (openAt > now) break;
    out.push({ id: cycleId(task.id, nominal), cycleNo: k, nominal: nominal, openAt: openAt, dueAt: atTime(dueDay, dueHM) });
  }
  return out;
}

const ST = { OPEN: 'OPEN', OVERDUE: 'OVERDUE', DONE: 'DONE', DONE_LATE: 'DONE_LATE', WAIVED: 'WAIVED' };

function isClosed(c) { return c.Status === ST.DONE || c.Status === ST.DONE_LATE || c.Status === ST.WAIVED; }
function effectiveDue(c) { return c['Extended Due At'] instanceof Date ? c['Extended Due At'] : c['Due At']; }

/** Status from facts. Lateness is always measured against the original due time. */
function statusOf(c, now) {
  if (c.Status === ST.WAIVED) return ST.WAIVED;
  if (c['Completed At'] instanceof Date) return c['Completed At'] <= c['Due At'] ? ST.DONE : ST.DONE_LATE;
  return now > effectiveDue(c) ? ST.OVERDUE : ST.OPEN;
}

/** When the manager alert goes: due + grace, capped at end of the working day so it is always same-day. */
function l1Time(dueAt, cfg) {
  const t = addMinutes(dueAt, cfg.graceMin);
  const close = atTime(dueAt, cfg.workEnd);
  return dueAt < close && t > close ? close : t;
}

function ackDeadline(l1SentAt, cfg) { return addWorkingHours(l1SentAt, cfg.ackHrs, cfg); }

/** By when the task itself must be done before leadership hears about it. */
function resolveDeadline(c, isDaily, cfg) {
  if (c['Extended Due At'] instanceof Date) return c['Extended Due At'];
  const due = c['Due At'];
  if (isDaily) return atTime(due, cfg.dailyResolve);
  return atTime(addWorkingDays(due, cfg.portalResolveDays, cfg), { h: due.getHours(), m: due.getMinutes() });
}

/** Why this cycle must go to leadership now, or null. */
function l2Reason(c, isDaily, now, cfg) {
  if (isClosed(c) || c['L2 Sent At'] instanceof Date) return null;
  if (c['Manager Action'] === ACTIONS.ESCALATE) return 'Escalated by manager';
  if (!(c['L1 Sent At'] instanceof Date)) return null;
  if (!(c['Ack At'] instanceof Date) && now >= ackDeadline(c['L1 Sent At'], cfg)) return 'Manager did not acknowledge';
  // A late daily update is handled by the manager and shows up in the weekly pattern digest instead.
  if (isDaily && !cfg.dailyL2OnUnresolved) return null;
  if (now >= resolveDeadline(c, isDaily, cfg)) {
    return c['Ack At'] instanceof Date ? 'Acknowledged but unresolved' : 'Unresolved past deadline';
  }
  return null;
}

function normalizeText(s) { return String(s == null ? '' : s).toLowerCase().replace(/\s+/g, ' ').trim(); }

const ACTIONS = {
  REMINDED: 'Reminded employee',
  EXTEND: 'Extend deadline',
  WAIVE: 'Waive (reason required)',
  ESCALATE: 'Escalate to leadership now'
};
