/** Generated views. Never edit these as data: they are rebuilt from the Cycles ledger on every run. */

const EMP_HEADER_ROW = 5;
const EMP_FIRST_ROW = 6;
const EMP_HEADERS = ['CycleID', 'Task', 'Due', 'Status', 'Proof required', 'Ticket ID / Update ✎', 'Posted to WhatsApp ✎', 'Share', 'Completed at'];
const EMP_COL = { ID: 1, PROOF: 6, POSTED: 7, SHARE: 8 };

const DASH_SUMMARY_ROW = 4;
const DASH_HEADERS = ['CycleID', 'Employee', 'Task', 'Due', 'Status', 'Overdue by', 'Manager alerted', 'Act by',
  'Action ✎', 'Note ✎', 'New due ✎', 'Acknowledged', 'Escalated', 'Escalation reason'];
const DASH_COL = { ACTION: 9, NOTE: 10, NEWDUE: 11 };
const DASH_SUMMARY_HEADERS = ['Employee', 'Open', 'Overdue', 'Done on time (7d)', 'Done late (7d)', 'On-time % (30d)', 'Awaiting manager action'];

const COLORS = { overdue: '#f4cccc', today: '#fff2cc', open: '#ffffff', done: '#d9ead3', history: '#efefef' };

function ensureSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function refreshViews_(ctx) {
  const emp = {};
  ctx.employees.forEach(function (e) { if (e.active && e.tab) emp[e.id] = refreshEmployeeTab_(ctx, e); });
  return { emp: emp, dash: refreshDashboard_(ctx) };
}

function statusLabel_(c, now) {
  switch (c.Status) {
    case ST.OVERDUE: return 'OVERDUE';
    case ST.DONE: return 'Done';
    case ST.DONE_LATE: return 'Done (late)';
    case ST.WAIVED: return 'Waived';
    default:
      if (c['Extended Due At'] instanceof Date) return 'Open (extended)';
      return ymd(effectiveDue(c)) === ymd(now) ? 'Due today' : 'Open';
  }
}

/** Clears the generated body (row `first` down) of a view. */
function clearBody_(sh, first, width) {
  const last = Math.max(sh.getLastRow(), first);
  sh.getRange(first, 1, last - first + 1, width)
    .clearContent().clearDataValidations().clearNote().setBackground(null);
}

/** Returns {CycleID: row} so callers can annotate cells. */
function refreshEmployeeTab_(ctx, emp) {
  const sh = ensureSheet_(emp.tab);
  const now = ctx.now;
  const mine = ctx.cycles.filter(function (c) { return c.EmpID === emp.id; });
  const pending = mine.filter(function (c) { return !isClosed(c); }).sort(function (a, b) {
    return (b.Status === ST.OVERDUE) - (a.Status === ST.OVERDUE) || effectiveDue(a) - effectiveDue(b);
  });
  const history = mine.filter(isClosed).sort(function (a, b) {
    return (b['Completed At'] || b['Due At']) - (a['Completed At'] || a['Due At']);
  }).slice(0, 5);
  const rows = pending.concat(history);
  const overdue = pending.filter(function (c) { return c.Status === ST.OVERDUE; }).length;

  // Column A holds the hidden CycleID, so visible text starts in column B.
  sh.getRange(1, 2, 3, 1).setValues([
    [EMP_TAB_PREFIX + emp.name],
    [pending.length + ' pending · ' + overdue + ' overdue · refreshed ' + fmtFriendly(now) + (ctx.cfg.testNow ? ' (simulated time)' : '')],
    ['Portal task: type the Ticket ID in the ✎ column – that completes it. Daily update: type it in the ✎ column, ' +
      'click "Share to WhatsApp", post it in the group, then tick Posted. Grey rows are your recent history.']
  ]);
  sh.getRange(1, 2).setFontWeight('bold').setFontSize(14);
  sh.getRange(EMP_HEADER_ROW, 1, 1, EMP_HEADERS.length).setValues([EMP_HEADERS]).setFontWeight('bold').setBackground('#cfe2f3');
  clearBody_(sh, EMP_FIRST_ROW, EMP_HEADERS.length);
  sh.setFrozenRows(EMP_HEADER_ROW);
  sh.hideColumns(EMP_COL.ID);

  const pos = {};
  if (!rows.length) { sh.getRange(EMP_FIRST_ROW, 2).setValue('Nothing pending right now.'); return pos; }

  const values = [];
  const bgs = [];
  rows.forEach(function (c, i) {
    const r = EMP_FIRST_ROW + i;
    pos[c.CycleID] = r;
    const task = ctx.tasks[c.TaskID] || {};
    const daily = task.proofType === 'DAILY_UPDATE';
    const label = statusLabel_(c, now);
    values.push([
      c.CycleID, c['Task Name'], effectiveDue(c), label,
      daily ? 'Update text + Posted tick' : 'Ticket ID' + (task.pattern ? ' (format ' + task.pattern + ')' : ''),
      safeCell_(c.Proof), daily ? c.Posted === true : '',
      daily ? '=HYPERLINK("https://wa.me/?text="&ENCODEURL(F' + r + '),"Share to WhatsApp")' : '',
      c['Completed At'] || ''
    ]);
    const color = isClosed(c) ? COLORS.history : c.Status === ST.OVERDUE ? COLORS.overdue : label === 'Due today' ? COLORS.today : COLORS.open;
    bgs.push(EMP_HEADERS.map(function () { return color; }));
    // insertCheckboxes resets values to FALSE, so it has to run before setValues.
    if (daily) sh.getRange(r, EMP_COL.POSTED).insertCheckboxes();
  });
  const body = sh.getRange(EMP_FIRST_ROW, 1, rows.length, EMP_HEADERS.length);
  sh.getRange(EMP_FIRST_ROW, EMP_COL.PROOF, rows.length, 1).setNumberFormat('@');
  body.setValues(values).setBackgrounds(bgs);
  sh.getRange(EMP_FIRST_ROW, 3, rows.length, 1).setNumberFormat('ddd d mmm, hh:mm');
  sh.getRange(EMP_FIRST_ROW, 9, rows.length, 1).setNumberFormat('ddd d mmm, hh:mm');
  return pos;
}

function refreshDashboard_(ctx) {
  const sh = ensureSheet_(SHEETS.DASH);
  const now = ctx.now;
  const lastRun = PropertiesService.getScriptProperties().getProperty('LAST_RUN');
  const minsAgo = lastRun ? Math.round((Date.now() - new Date(lastRun).getTime()) / 60000) : null;

  // Summary: one row per employee, so the manager never checks employees one by one.
  const summary = ctx.employees.filter(function (e) { return e.active; }).map(function (e) {
    const mine = ctx.cycles.filter(function (c) { return c.EmpID === e.id; });
    const within = function (d, days) { return d instanceof Date && d <= now && d >= addDays(now, -days); };
    const judged = mine.filter(function (c) { return within(c['Due At'], 30) && c.Status !== ST.WAIVED; });
    const onTime = judged.filter(function (c) { return c.Status === ST.DONE; }).length;
    return [
      e.name,
      mine.filter(function (c) { return c.Status === ST.OPEN; }).length,
      mine.filter(function (c) { return c.Status === ST.OVERDUE; }).length,
      mine.filter(function (c) { return c.Status === ST.DONE && within(c['Completed At'], 7); }).length,
      mine.filter(function (c) { return c.Status === ST.DONE_LATE && within(c['Completed At'], 7); }).length,
      judged.length ? Math.round(100 * onTime / judged.length) + '%' : '—',
      mine.filter(function (c) { return !isClosed(c) && c['L1 Sent At'] instanceof Date && !(c['Ack At'] instanceof Date); }).length
    ];
  });

  const actionHeaderRow = DASH_SUMMARY_ROW + summary.length + 3;
  const first = actionHeaderRow + 1;
  const items = ctx.cycles.filter(function (c) {
    return !isClosed(c) && (c.Status === ST.OVERDUE || c['L1 Sent At'] instanceof Date);
  }).sort(function (a, b) { return a['Due At'] - b['Due At']; });

  clearBody_(sh, 1, DASH_HEADERS.length);
  // Column A holds the hidden CycleID, so visible text starts in column B.
  sh.getRange(1, 2, 2, 1).setValues([
    ['Manager Dashboard – recurring tasks'],
    ['Refreshed ' + fmtFriendly(now) + (ctx.cfg.testNow ? ' (simulated time)' : '') +
      ' · last engine run ' + (minsAgo === null ? 'never' : minsAgo + ' min ago') +
      (minsAgo !== null && minsAgo > 60 ? ' ⚠ engine may be stopped – check triggers' : '')]
  ]);
  sh.getRange(1, 2).setFontWeight('bold').setFontSize(14);
  if (minsAgo === null || minsAgo > 60) sh.getRange(2, 2).setBackground(COLORS.overdue);

  sh.getRange(DASH_SUMMARY_ROW, 2, 1, DASH_SUMMARY_HEADERS.length).setValues([DASH_SUMMARY_HEADERS])
    .setFontWeight('bold').setBackground('#cfe2f3');
  if (summary.length) sh.getRange(DASH_SUMMARY_ROW + 1, 2, summary.length, DASH_SUMMARY_HEADERS.length).setValues(summary);

  sh.getRange(actionHeaderRow - 1, 2).setValue(
    'ACTION LIST – choose an Action for each item (this records your acknowledgement). ' +
    'For Waive, type the reason in Note first. For Extend, enter New due first.').setFontWeight('bold');
  sh.getRange(actionHeaderRow, 1, 1, DASH_HEADERS.length).setValues([DASH_HEADERS]).setFontWeight('bold').setBackground('#cfe2f3');
  sh.hideColumns(1);

  const pos = {};
  if (!items.length) { sh.getRange(first, 2).setValue('No overdue items.'); return pos; }

  const values = items.map(function (c, i) {
    pos[c.CycleID] = first + i;
    return [
      c.CycleID, c.Employee, c['Task Name'], effectiveDue(c), statusLabel_(c, now), overdueBy_(c, now),
      c['L1 Sent At'] || '', c['L1 Sent At'] instanceof Date ? ackDeadline(c['L1 Sent At'], ctx.cfg) : '',
      c['Manager Action'] || '', safeCell_(c['Manager Note']), c['Requested Due'] || c['Extended Due At'] || '',
      c['Ack At'] || '', c['L2 Sent At'] || '', c['L2 Reason'] || ''
    ];
  });
  sh.getRange(first, 1, values.length, DASH_HEADERS.length).setValues(values).setBackgrounds(items.map(function (c) {
    const color = c['L2 Sent At'] instanceof Date ? COLORS.overdue : c['Ack At'] instanceof Date ? COLORS.today : COLORS.open;
    return DASH_HEADERS.map(function () { return color; });
  }));
  const actionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(ACTIONS).map(function (k) { return ACTIONS[k]; }), true).setAllowInvalid(false).build();
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  sh.getRange(first, DASH_COL.ACTION, values.length, 1).setDataValidation(actionRule);
  sh.getRange(first, DASH_COL.NEWDUE, values.length, 1).setDataValidation(dateRule);
  [4, 7, 8, 11, 12, 13].forEach(function (col) { sh.getRange(first, col, values.length, 1).setNumberFormat('ddd d mmm, hh:mm'); });
  return pos;
}
