/** The scheduled engine: create cycles, update statuses, send alerts/escalations, refresh views. */

/** Time-driven trigger target (every 15 min). */
function scheduledRun() {
  const flag = PropertiesService.getScriptProperties().getProperty('TESTS_RUNNING');
  // The test suite owns the workbook while it runs; skip instead of fighting over the simulated clock.
  if (flag && Date.now() - Number(flag) < 15 * 60000) return;
  runEngine();
}

function runEngine(opts) {
  opts = opts || {};
  return withLock_(function () {
    const t0 = Date.now();
    const ctx = loadContext_();
    const stats = { created: 0, changed: 0, emails: 0, notes: [] };
    try {
      stats.created = generateCycles_(ctx, stats.notes);
      stats.changed = updateStatuses_(ctx);
      saveContext_(ctx);

      // No emails on non-working days; anything pending goes out on the next working day.
      if (isWorkingDay(ctx.now, ctx.cfg)) {
        stats.emails += sendReminders_(ctx);
        stats.emails += sendManagerAlerts_(ctx);
        stats.emails += sendLeadershipEscalations_(ctx);
        stats.emails += sendWeeklyDigest_(ctx);
        saveContext_(ctx);
      }
      if (opts.refreshViews !== false) refreshViews_(ctx);

      PropertiesService.getScriptProperties().setProperty('LAST_RUN', new Date().toISOString());
      // Only runs that did something are logged, so the log stays readable at 96 runs a day.
      if (stats.created || stats.changed || stats.emails || stats.notes.length) {
        logRun_(ctx, stats, Date.now() - t0, stats.notes.join('; '));
      }
      return stats;
    } catch (e) {
      logRun_(ctx, stats, Date.now() - t0, 'ERROR: ' + (e && e.stack || e));
      throw e; // Apps Script emails the owner about failed trigger runs.
    }
  });
}

function logRun_(ctx, stats, ms, note) {
  appendRows_(SHEETS.RUN, [[new Date(), ctx.cfg.testNow ? ctx.now : '', stats.created, stats.changed, stats.emails, ms, note]]);
}

function generateCycles_(ctx, notes) {
  let created = 0;
  Object.keys(ctx.tasks).forEach(function (id) {
    const task = ctx.tasks[id];
    if (!task.active) return;
    const emp = ctx.empById[task.empId];
    if (!emp || !emp.active) { notes.push(id + ': assignee ' + task.empId + ' missing or inactive'); return; }
    if (!task.startDate) { notes.push(id + ': no Start Date'); return; }
    expectedCycles(task, ctx.now, ctx.cfg).forEach(function (x) {
      if (ctx.byId[x.id]) return; // deterministic ID → re-running never duplicates
      const c = {};
      HEADERS.Cycles.forEach(function (h) { c[h] = ''; });
      Object.assign(c, {
        CycleID: x.id, TaskID: task.id, 'Task Name': task.name, EmpID: emp.id, Employee: emp.name,
        'Cycle No': x.cycleNo, 'Nominal Date': x.nominal, 'Open At': x.openAt, 'Due At': x.dueAt,
        Status: ST.OPEN, Posted: false, 'Last Updated': ctx.now
      });
      c.Status = statusOf(c, ctx.now);
      ctx.cycles.push(c);
      ctx.byId[x.id] = c;
      created++;
    });
  });
  return created;
}

function updateStatuses_(ctx) {
  let changed = 0;
  ctx.cycles.forEach(function (c) {
    const s = statusOf(c, ctx.now);
    if (s !== c.Status) { c.Status = s; c['Last Updated'] = ctx.now; changed++; }
  });
  return changed;
}

/* ───────────────────────── Notifications ───────────────────────── */

/**
 * Logs first (status SENDING), then sends, then marks SENT. A crash between send and stamp leaves
 * a SENDING row, which still blocks a resend: one missed stamp is better than a duplicate email.
 * items: [{key, cycleId, details}]
 */
function sendDigest_(ctx, type, to, cc, subject, html, items) {
  const ref = type + '-' + ymd(ctx.now) + '-' + hm(ctx.now).replace(':', '') + '-' + (++ctx.emailSeq);
  const status = ctx.cfg.dryRun ? 'DRY_RUN' : 'SENDING';
  const rows = items.map(function (it) {
    return [it.key, type, it.cycleId || '', to + (cc ? ' cc ' + cc : ''), subject, ref, it.details || '', ctx.now, status];
  });
  const first = appendRows_(SHEETS.NOTIF, rows);
  items.forEach(function (it) { ctx.notifKeys.add(it.key); });
  if (ctx.cfg.dryRun) return true;

  SpreadsheetApp.flush();
  const statusCol = HEADERS.Notification_Log.indexOf('Status') + 1;
  const statusRange = sheet_(SHEETS.NOTIF).getRange(first, statusCol, rows.length, 1);
  try {
    const msg = { to: to, subject: subject, htmlBody: html };
    if (cc) msg.cc = cc;
    MailApp.sendEmail(msg);
  } catch (e) {
    // FAILED rows are ignored when loading keys, so the next run retries them.
    statusRange.setValues(rows.map(function () { return ['FAILED: ' + e.message]; }));
    items.forEach(function (it) { ctx.notifKeys.delete(it.key); });
    return false;
  }
  statusRange.setValues(rows.map(function () { return ['SENT']; }));
  return true;
}

function subjectPrefix_(ctx) { return ctx.cfg.testMode ? '[TEST] ' : ''; }

function managerOf_(ctx, c) {
  const e = ctx.empById[c.EmpID];
  return (e && e.managerEmail) || ctx.cfg.adminEmail;
}

function proofLabel_(ctx, c) {
  const t = ctx.tasks[c.TaskID];
  if (!t) return '';
  return t.proofType === 'DAILY_UPDATE' ? 'Daily update + Posted tick' : 'Ticket ID';
}

function groupBy_(arr, keyFn) {
  const out = {};
  arr.forEach(function (x) { const k = keyFn(x); (out[k] = out[k] || []).push(x); });
  return out;
}

function sendReminders_(ctx) {
  if (!ctx.cfg.reminderEnabled) return 0;
  const now = ctx.now;
  const due = ctx.cycles.filter(function (c) {
    if (c.Status !== ST.OPEN || ctx.notifKeys.has('REM|' + c.CycleID)) return false;
    const d = effectiveDue(c);
    return ymd(d) === ymd(now) && now < d && now >= addMinutes(atTime(d, ctx.cfg.workStart), ctx.cfg.reminderAfterStartMin);
  });
  let sent = 0;
  const groups = groupBy_(due, function (c) { return c.EmpID; });
  Object.keys(groups).forEach(function (empId) {
    const emp = ctx.empById[empId];
    if (!emp || !emp.email) return;
    const list = groups[empId];
    const html = '<p>Hi ' + esc_(emp.name) + ',</p><p>These recurring tasks are due today:</p>' +
      table_(['Task', 'Due', 'What to enter'], list.map(function (c) {
        return [c['Task Name'], fmtFriendly(effectiveDue(c)), proofLabel_(ctx, c)];
      })) + footer_(ctx, emp.tab, 'Open your task sheet');
    if (sendDigest_(ctx, 'REM', emp.email, '', subjectPrefix_(ctx) + 'Reminder: ' + list.length + ' task(s) due today',
      html, list.map(function (c) { return { key: 'REM|' + c.CycleID, cycleId: c.CycleID }; }))) sent++;
  });
  return sent;
}

/** L1: one same-day digest per manager for tasks that missed their due time. */
function sendManagerAlerts_(ctx) {
  const now = ctx.now;
  const due = ctx.cycles.filter(function (c) {
    return c.Status === ST.OVERDUE && !(c['L1 Sent At'] instanceof Date) &&
      !ctx.notifKeys.has('L1|' + c.CycleID) && now >= l1Time(c['Due At'], ctx.cfg);
  });
  let sent = 0;
  const groups = groupBy_(due, function (c) { return managerOf_(ctx, c); });
  Object.keys(groups).forEach(function (mgr) {
    if (!mgr) return;
    const list = groups[mgr];
    const ackBy = ackDeadline(now, ctx.cfg);
    const html = '<p>Hi,</p><p>The following recurring tasks have <b>missed their due time</b>:</p>' +
      table_(['Employee', 'Task', 'Due', 'Overdue by', 'Required proof'], list.map(function (c) {
        return [c.Employee, c['Task Name'], fmtFriendly(c['Due At']), overdueBy_(c, now), proofLabel_(ctx, c)];
      })) +
      '<p><b>What to do:</b> follow up with the employee, then open the Manager Dashboard and choose an <b>Action</b> ' +
      'for each item by <b>' + fmtFriendly(ackBy) + '</b>. Choosing an action records your acknowledgement. ' +
      'Items with no action by then are escalated to leadership automatically.</p>' + footer_(ctx, SHEETS.DASH, 'Open Manager Dashboard');
    const ok = sendDigest_(ctx, 'L1', mgr, '', subjectPrefix_(ctx) + 'Pending tasks: ' + list.length + ' overdue – ' + ymd(now), html,
      list.map(function (c) { return { key: 'L1|' + c.CycleID, cycleId: c.CycleID, details: c.Employee + ' / ' + c['Task Name'] }; }));
    if (ok) {
      list.forEach(function (c) { c['L1 Sent At'] = now; c['Last Updated'] = now; });
      sent++;
    }
  });
  return sent;
}

/** L2: one digest to leadership (cc managers) for new escalations, plus escalations that are now resolved. */
function sendLeadershipEscalations_(ctx) {
  const now = ctx.now;
  const escalate = [];
  ctx.cycles.forEach(function (c) {
    if (ctx.notifKeys.has('L2|' + c.CycleID)) return;
    const reason = l2Reason(c, isDailyTask_(ctx, c), now, ctx.cfg);
    if (reason) escalate.push({ c: c, reason: reason });
  });
  const resolved = ctx.cycles.filter(function (c) {
    return isClosed(c) && c['L2 Sent At'] instanceof Date && !ctx.notifKeys.has('RES|' + c.CycleID);
  });
  if (!escalate.length && !resolved.length) return 0;

  const to = ctx.cfg.leadershipEmail || ctx.cfg.adminEmail;
  if (!to) return 0;
  const cc = Array.from(new Set(escalate.map(function (x) { return managerOf_(ctx, x.c); }).filter(Boolean))).join(',');

  let html = '<p>Hi,</p>';
  if (escalate.length) {
    html += '<p>These recurring tasks are <b>escalated</b>:</p>' +
      table_(['Employee', 'Task', 'Due', 'Overdue by', 'Reason', 'Manager action', 'Manager note'], escalate.map(function (x) {
        return [x.c.Employee, x.c['Task Name'], fmtFriendly(x.c['Due At']), overdueBy_(x.c, now), x.reason,
          x.c['Manager Action'] || '—', x.c['Manager Note'] || ''];
      }));
  }
  if (resolved.length) {
    html += '<p>Previously escalated items that are now <b>closed</b>:</p>' +
      table_(['Employee', 'Task', 'Due', 'Outcome', 'Closed / proof'], resolved.map(function (c) {
        return [c.Employee, c['Task Name'], fmtFriendly(c['Due At']), c.Status,
          c.Status === ST.WAIVED ? 'Waived: ' + (c['Manager Note'] || '') : fmtFriendly(c['Completed At']) + ' · ' + c.Proof];
      }));
  }
  html += footer_(ctx, SHEETS.DASH, 'Open Manager Dashboard');

  const subject = subjectPrefix_(ctx) + (escalate.length
    ? 'Escalation: ' + escalate.length + ' unresolved task(s)' + (resolved.length ? ', ' + resolved.length + ' resolved' : '')
    : 'Resolved: ' + resolved.length + ' escalated task(s) now closed');
  const items = escalate.map(function (x) {
    return { key: 'L2|' + x.c.CycleID, cycleId: x.c.CycleID, details: x.reason };
  }).concat(resolved.map(function (c) { return { key: 'RES|' + c.CycleID, cycleId: c.CycleID, details: c.Status }; }));

  if (!sendDigest_(ctx, 'L2', to, cc, subject, html, items)) return 0;
  escalate.forEach(function (x) {
    x.c['L2 Sent At'] = now; x.c['L2 Reason'] = x.reason; x.c['Last Updated'] = now;
  });
  return 1;
}

/** Weekly leadership summary (first working-hours run of the week) – where daily-update patterns surface. */
function sendWeeklyDigest_(ctx) {
  const now = ctx.now;
  if (now < atTime(now, ctx.cfg.workStart)) return 0;
  const monday = addDays(dayStart(now), -((now.getDay() + 6) % 7));
  const key = 'WEEKLY|' + ymd(monday);
  if (ctx.notifKeys.has(key)) return 0;
  const from = addDays(monday, -7);
  const inWeek = ctx.cycles.filter(function (c) { return c['Due At'] >= from && c['Due At'] < monday; });
  if (!inWeek.length) return 0;

  const rows = [];
  const flags = [];
  ctx.employees.forEach(function (e) {
    const mine = inWeek.filter(function (c) { return c.EmpID === e.id; });
    if (!mine.length) return;
    const count = function (s) { return mine.filter(function (c) { return c.Status === s; }).length; };
    const dailyMisses = mine.filter(function (c) {
      return isDailyTask_(ctx, c) && (c.Status === ST.DONE_LATE || c.Status === ST.OVERDUE);
    }).length;
    const flag = dailyMisses >= ctx.cfg.patternMisses ? 'PATTERN: ' + dailyMisses + ' late/missed daily updates' : '';
    if (flag) flags.push(e.name + ' – ' + flag);
    rows.push([e.name, mine.length, count(ST.DONE), count(ST.DONE_LATE), count(ST.OVERDUE) + count(ST.OPEN), count(ST.WAIVED), dailyMisses, flag]);
  });
  const to = ctx.cfg.leadershipEmail || ctx.cfg.adminEmail;
  if (!to) return 0;
  const cc = Array.from(new Set(ctx.employees.map(function (e) { return e.managerEmail; }).filter(Boolean))).join(',');
  const html = '<p>Recurring-task summary for ' + ymd(from) + ' to ' + ymd(addDays(monday, -1)) + ':</p>' +
    table_(['Employee', 'Due', 'On time', 'Late', 'Still open', 'Waived', 'Daily misses', 'Flag'], rows) +
    footer_(ctx, SHEETS.DASH, 'Open Manager Dashboard');
  return sendDigest_(ctx, 'WEEKLY', to, cc, subjectPrefix_(ctx) + 'Weekly recurring-task summary – week of ' + ymd(from), html,
    [{ key: key, cycleId: '', details: flags.length ? flags.join(' | ') : 'No patterns' }]) ? 1 : 0;
}

function overdueBy_(c, now) {
  const h = (now - c['Due At']) / 3600000;
  if (h <= 0) return '—';
  return h < 24 ? (Math.round(h * 10) / 10) + ' h' : (Math.round(h / 2.4) / 10) + ' days';
}

function esc_(v) {
  return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function table_(headers, rows) {
  return '<table cellpadding="6" style="border-collapse:collapse;font:13px Arial,sans-serif" border="1">' +
    '<tr>' + headers.map(function (h) { return '<th style="background:#f1f3f4;text-align:left">' + esc_(h) + '</th>'; }).join('') + '</tr>' +
    rows.map(function (r) { return '<tr>' + r.map(function (v) { return '<td>' + esc_(v) + '</td>'; }).join('') + '</tr>'; }).join('') +
    '</table>';
}

function footer_(ctx, sheetName, label) {
  const sh = ctx.ss.getSheetByName(sheetName);
  const url = ctx.ss.getUrl() + (sh ? '#gid=' + sh.getSheetId() : '');
  return '<p><a href="' + esc_(url) + '">' + esc_(label) + '</a></p>' +
    (ctx.cfg.testNow ? '<p style="color:#888">Simulated time: ' + fmtDateTime(ctx.now) + '</p>' : '') +
    '<p style="color:#888">Sent automatically by Allo Ops recurring-task tracker.</p>';
}
