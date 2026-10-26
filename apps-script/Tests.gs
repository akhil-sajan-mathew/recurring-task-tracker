/**
 * Automated test suite. Runs the real code paths against a simulated clock (Config TEST_NOW)
 * with DRY_RUN_EMAIL on, and writes PASS/FAIL to Test_Results.
 * Demo calendar: the schedule starts Mon 2 Nov 2026. Work days Mon–Sat.
 */

const TEST_START = new Date(2026, 10, 2);
const MON_TEXT = 'Dispatched 42 orders, 3 returns pending QC, Amazon claim 118 filed';
const TUE_TEXT = 'Dispatched 38 orders, all returns cleared, 2 listings fixed on Flipkart';

function runAllTests() { return runTestGroups([0, 1, 2, 3, 4, 5, 6, 7, 8]); }

function runTestGroups(groups) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('TESTS_RUNNING', String(Date.now()));
  const raw = loadConfig_().raw;
  const saved = {};
  ['TEST_NOW', 'DRY_RUN_EMAIL', 'TEST_MODE'].forEach(function (k) { saved[k] = raw[k]; });
  const results = [];
  try {
    setConfigValue('DRY_RUN_EMAIL', 'TRUE');
    setConfigValue('TEST_MODE', 'TRUE');
    groups.forEach(function (g) { TEST_GROUPS[g].forEach(function (t) { results.push(runOneTest_(t)); }); });
  } finally {
    Object.keys(saved).forEach(function (k) { setConfigValue(k, saved[k] == null ? '' : saved[k]); });
    // Leave a clean live demo behind (schedule starting this week).
    loadDemoData_(mondayOf_(new Date()), true);
    resetData_();
    props.deleteProperty('TESTS_RUNNING');
    saveTestResults_(results);
  }
  return results;
}

/** Keeps earlier results for tests that were not re-run, so two half-runs leave the full table. */
function saveTestResults_(results) {
  const order = [];
  TEST_GROUPS.forEach(function (g) { g.forEach(function (t) { order.push(t.id); }); });
  const fresh = {};
  results.forEach(function (r) { fresh[r.id] = [r.id, r.name, r.expected, r.actual, r.result, new Date()]; });
  const rows = {};
  sheet_(SHEETS.TESTS).getDataRange().getValues().slice(1).forEach(function (r) { if (r[0] !== '') rows[String(r[0])] = r; });
  Object.keys(fresh).forEach(function (id) { rows[id] = fresh[id]; });
  clearRows_(SHEETS.TESTS);
  appendRows_(SHEETS.TESTS, order.filter(function (id) { return rows[id]; }).map(function (id) { return rows[id]; }));
}

function runOneTest_(t) {
  let r;
  try { r = t.run(); } catch (e) { r = { pass: false, actual: 'ERROR: ' + (e && e.message || e) }; }
  return { id: t.id, name: t.name, expected: t.expected, actual: r.actual,
    result: r.pass === null ? 'MANUAL' : r.pass ? 'PASS' : 'FAIL' };
}

/* ───── helpers ───── */

function tReset_() { loadDemoData_(TEST_START, true); resetData_(); }
function tAt_(s) { setConfigValue('TEST_NOW', s); }
function tRun_() { return runEngine({ refreshViews: false }); }
function tCyc_(id) { return readObjects_(SHEETS.CYCLES).filter(function (c) { return String(c.CycleID) === id; })[0]; }
function tCycles_() { return readObjects_(SHEETS.CYCLES); }
function tNotifs_(type) { return readObjects_(SHEETS.NOTIF).filter(function (n) { return !type || n.Type === type; }); }
function tNotifFor_(type, id) { return tNotifs_(type).filter(function (n) { return n.NotifKey === type + '|' + id; }); }
function tEmp_(id) { return readObjects_(SHEETS.EMP).filter(function (e) { return e.EmpID === id; })[0]; }
function f_(d) { return d instanceof Date ? fmtDateTime(d) : String(d === '' || d == null ? '(blank)' : d); }
function res_(pass, actual) { return { pass: !!pass, actual: actual }; }
function tRowOf_(sh, id, firstRow) {
  const last = sh.getLastRow();
  if (last < firstRow) return 0;
  const ids = sh.getRange(firstRow, 1, last - firstRow + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return firstRow + i;
  return 0;
}

/* ───── test groups ───── */

const TEST_GROUPS = [
  // 0 – date/SLA building blocks
  [
    { id: 'U.1', name: 'Working-hours clock pauses overnight and on Sunday',
      expected: 'Sat 17:30 + 4 working hours = Mon 12:30',
      run: function () { tReset_(); const d = ackDeadline(new Date(2026, 10, 7, 17, 30), loadConfig_()); return res_(fmtDateTime(d) === '2026-11-09 12:30', fmtDateTime(d)); } },
    { id: 'U.2', name: 'Working-day arithmetic skips Sunday',
      expected: 'Mon 2 Nov − 1 = Sat 31 Oct; Sat 7 Nov + 1 = Mon 9 Nov',
      run: function () {
        const cfg = loadConfig_();
        const a = ymd(addWorkingDays(new Date(2026, 10, 2), -1, cfg));
        const b = ymd(addWorkingDays(new Date(2026, 10, 7), 1, cfg));
        return res_(a === '2026-10-31' && b === '2026-11-09', a + ', ' + b);
      } },
    { id: 'U.3', name: '7-day schedule generated from the start date',
      expected: 'By Sat 14 Nov 10:00: cycles due 2, 9, 16 Nov (16 Nov opens 1 working day early)',
      run: function () {
        const task = { id: 'TX', freq: 'EVERY_N_DAYS', cycleDays: 7, startDate: TEST_START, leadDays: 1, dueTime: null };
        const ids = expectedCycles(task, new Date(2026, 10, 14, 10, 0), loadConfig_()).map(function (x) { return x.id; });
        return res_(ids.join() === 'TX|2026-11-02,TX|2026-11-09,TX|2026-11-16', ids.join(', '));
      } }
  ],

  // 1 – Test 1: recurring cycle
  [
    { id: '1.1', name: 'Cycle appears automatically when its lead window opens',
      expected: 'Sat 31 Oct 10:00: T02 cycle due Mon 2 Nov 17:00 exists, OPEN; no daily cycle before the start date',
      run: function () {
        tReset_(); tAt_('2026-10-31 10:00'); tRun_();
        const c = tCyc_('T02|2026-11-02');
        return res_(c && c.Status === ST.OPEN && f_(c['Due At']) === '2026-11-02 17:00' && !tCyc_('T01|2026-10-31'),
          c ? 'T02 due ' + f_(c['Due At']) + ', ' + c.Status : 'T02 cycle missing');
      } },
    { id: '1.2', name: 'Completing closes the current cycle',
      expected: 'Mon 15:00 ticket PTA-10234 → DONE, completed by Employee B at 15:00',
      run: function () {
        tAt_('2026-11-02 15:00'); tRun_();
        const r = apiCompleteTicket('T02|2026-11-02', 'PTA-10234');
        const c = tCyc_('T02|2026-11-02');
        const who = tEmp_('E02').Name;
        return res_(r.ok && c.Status === ST.DONE && f_(c['Completed At']) === '2026-11-02 15:00' && String(c['Completed By']).indexOf(who) === 0,
          c.Status + ' by ' + c['Completed By'] + ' at ' + f_(c['Completed At']));
      } },
    { id: '1.3', name: 'Next cycle is created by the system, not the employee',
      expected: 'Fri 6 Nov: no next cycle yet. Sat 7 Nov 10:00: T02 cycle 1 due Mon 9 Nov 17:00 created',
      run: function () {
        tAt_('2026-11-06 18:00'); tRun_();
        const before = !!tCyc_('T02|2026-11-09');
        tAt_('2026-11-07 10:00'); tRun_();
        const c = tCyc_('T02|2026-11-09');
        return res_(!before && c && f_(c['Due At']) === '2026-11-09 17:00' && Number(c['Cycle No']) === 1,
          'before: ' + (before ? 'exists' : 'none') + '; after: ' + (c ? 'cycle ' + c['Cycle No'] + ' due ' + f_(c['Due At']) : 'missing'));
      } },
    { id: '1.4', name: 'Re-running does not duplicate the next cycle',
      expected: 'Engine run 3 more times: still exactly one T02|2026-11-09, total cycle count unchanged',
      run: function () {
        const n0 = tCycles_().length;
        tRun_(); tRun_(); tRun_();
        const all = tCycles_();
        const dup = all.filter(function (c) { return c.CycleID === 'T02|2026-11-09'; }).length;
        return res_(all.length === n0 && dup === 1, 'cycles ' + n0 + ' → ' + all.length + ', T02|2026-11-09 rows: ' + dup);
      } },
    { id: '1.5', name: 'Late completion does not drift the schedule (FIXED basis)',
      expected: 'Cycle due 9 Nov done Wed 11 Nov → DONE_LATE; next cycle still due 16 Nov (not 18 Nov)',
      run: function () {
        tAt_('2026-11-11 12:00'); tRun_();
        apiCompleteTicket('T02|2026-11-09', 'PTA-10300');
        const late = tCyc_('T02|2026-11-09');
        tAt_('2026-11-14 10:00'); tRun_();
        const next = tCyc_('T02|2026-11-16');
        return res_(late.Status === ST.DONE_LATE && next && !tCyc_('T02|2026-11-18'),
          '9 Nov cycle ' + late.Status + '; next ' + (next ? 'due ' + f_(next['Due At']) : 'missing'));
      } },
    { id: '1.6', name: 'An unfinished cycle stays visible when the next one opens',
      expected: 'T03 (15-day) cycle due 2 Nov still OVERDUE and new cycle due 17 Nov OPEN; overdue listed first on C\'s tab',
      run: function () {
        tReset_(); tAt_('2026-11-14 10:00'); runEngine();
        const a = tCyc_('T03|2026-11-02'), b = tCyc_('T03|2026-11-17');
        const sh = sheet_(String(tEmp_('E03')['Tab Name']));
        const ra = tRowOf_(sh, 'T03|2026-11-02', EMP_FIRST_ROW), rb = tRowOf_(sh, 'T03|2026-11-17', EMP_FIRST_ROW);
        return res_(a && b && a.Status === ST.OVERDUE && b.Status === ST.OPEN && ra && rb && ra < rb,
          (a ? a.Status : 'missing') + ' / ' + (b ? b.Status : 'missing') + '; tab rows ' + ra + ', ' + rb);
      } },
    { id: '1.7', name: 'Due date on a holiday moves to the next working day',
      expected: 'Holiday Mon 2 Nov: T04 due Tue 3 Nov 17:00 (ID keeps nominal date); no daily cycle on the holiday',
      run: function () {
        tReset_(); appendRows_(SHEETS.HOL, [[new Date(2026, 10, 2), 'Test holiday']]);
        tAt_('2026-11-03 10:00'); tRun_();
        const c = tCyc_('T04|2026-11-02');
        return res_(c && f_(c['Due At']) === '2026-11-03 17:00' && !tCyc_('T01|2026-11-02') && tCyc_('T01|2026-11-03'),
          c ? 'T04 due ' + f_(c['Due At']) + '; daily on holiday: ' + (tCyc_('T01|2026-11-02') ? 'yes' : 'no') : 'missing');
      } }
  ],

  // 2 – Test 2: single-point completion
  [
    { id: '2.1', name: 'Employee completes from their own tab only',
      expected: 'Typing PTA-10234 on "My Tasks – Employee B" → ledger DONE, row moves to history as Done; no other sheet touched',
      run: function () {
        tReset_(); tAt_('2026-11-02 15:00'); runEngine();
        const sh = sheet_(String(tEmp_('E02')['Tab Name']));
        const row = tRowOf_(sh, 'T02|2026-11-02', EMP_FIRST_ROW);
        if (!row) return res_(false, 'task not on tab');
        const cell = sh.getRange(row, EMP_COL.PROOF);
        cell.setValue('PTA-10234');
        onEditInstalled({ range: cell });
        const c = tCyc_('T02|2026-11-02');
        const row2 = tRowOf_(sh, 'T02|2026-11-02', EMP_FIRST_ROW);
        const shown = row2 ? sh.getRange(row2, 4).getValue() : '';
        return res_(c.Status === ST.DONE && c.Proof === 'PTA-10234' && shown === 'Done',
          'ledger ' + c.Status + ' ' + c.Proof + '; tab shows "' + shown + '"');
      } },
    { id: '2.2', name: 'Wrong Ticket ID format is rejected',
      expected: '"abc" on T05 → rejected with format message; cycle stays OPEN; REJECTED in Audit_Log',
      run: function () {
        const r = apiCompleteTicket('T05|2026-11-03', 'abc');
        const c = tCyc_('T05|2026-11-03');
        const a = readObjects_(SHEETS.AUDIT).filter(function (x) { return x.CycleID === 'T05|2026-11-03' && x.Result === 'REJECTED'; });
        return res_(!r.ok && /format/.test(r.reason) && c.Status === ST.OPEN && a.length === 1, (r.reason || 'accepted') + '; status ' + c.Status);
      } },
    { id: '2.3', name: 'A ticket cannot close two cycles',
      expected: 'PTA-10234 (already used by T02) on T05 → rejected as duplicate',
      run: function () {
        const r = apiCompleteTicket('T05|2026-11-03', 'PTA-10234');
        return res_(!r.ok && /already used/.test(r.reason), r.reason || 'accepted');
      } },
    { id: '2.4', name: 'Correcting a ticket does not create a second completion',
      expected: 'T02 ticket changed to PTA-10235 → CORRECTION, Completed At unchanged (15:00)',
      run: function () {
        const r = apiCompleteTicket('T02|2026-11-02', 'pta-10235');
        const c = tCyc_('T02|2026-11-02');
        return res_(r.code === 'CORRECTION' && c.Proof === 'PTA-10235' && f_(c['Completed At']) === '2026-11-02 15:00',
          r.code + '; proof ' + c.Proof + '; completed ' + f_(c['Completed At']));
      } },
    { id: '2.5', name: 'A cycle cannot be completed before it opens',
      expected: 'Clock moved before Open At → rejected "not open yet"',
      run: function () {
        tReset_(); tAt_('2026-10-31 10:00'); tRun_();
        tAt_('2026-10-31 09:00');
        const r = apiCompleteTicket('T02|2026-11-02', 'PTA-10001');
        return res_(!r.ok && /not open/.test(r.reason), r.reason || 'accepted');
      } },
    { id: '2.6', name: 'Employees cannot edit each other\'s tabs',
      expected: 'Each tab\'s input range is protected to that employee (needs separate Google accounts to demo)',
      run: function () { return { pass: null, actual: 'Enforced by sheet protection – shown in the demo / Data → Protected ranges' }; } }
  ],

  // 3 – Test 3: daily task
  [
    { id: '3.1', name: 'Daily update recorded from the sheet',
      expected: 'Mon 10:15: update text + Posted tick → DONE, text stored as proof, time 10:15',
      run: function () {
        tReset_(); tAt_('2026-11-02 10:15'); tRun_();
        const r = apiSubmitDaily('T01|2026-11-02', MON_TEXT, true);
        const c = tCyc_('T01|2026-11-02');
        return res_(r.ok && c.Status === ST.DONE && c.Proof === MON_TEXT && f_(c['Completed At']) === '2026-11-02 10:15',
          c.Status + ' at ' + f_(c['Completed At']));
      } },
    { id: '3.2', name: 'Ticking Posted without a real update is rejected',
      expected: '"ok" + Posted → rejected (minimum length)',
      run: function () {
        tAt_('2026-11-03 10:15'); tRun_();
        const r = apiSubmitDaily('T01|2026-11-03', 'ok', true);
        return res_(!r.ok && /at least/.test(r.reason), r.reason || 'accepted');
      } },
    { id: '3.3', name: 'Copy-pasting yesterday\'s update is rejected',
      expected: 'Monday\'s text (different case/spacing) on Tuesday → rejected',
      run: function () {
        const r = apiSubmitDaily('T01|2026-11-03', '  ' + MON_TEXT.toUpperCase() + '  ', true);
        return res_(!r.ok && /same text/.test(r.reason), r.reason || 'accepted');
      } },
    { id: '3.4', name: 'Draft first, tick later',
      expected: 'Text without tick → DRAFT (still OPEN); tick afterwards → DONE',
      run: function () {
        const r1 = apiSubmitDaily('T01|2026-11-03', TUE_TEXT, false);
        const s1 = tCyc_('T01|2026-11-03').Status;
        const r2 = apiSubmitDaily('T01|2026-11-03', TUE_TEXT, true);
        const s2 = tCyc_('T01|2026-11-03').Status;
        return res_(r1.code === 'DRAFT' && s1 === ST.OPEN && r2.ok && s2 === ST.DONE, r1.code + '/' + s1 + ' → ' + r2.code + '/' + s2);
      } },
    { id: '3.5', name: 'No daily task on a non-working day',
      expected: 'Sunday 8 Nov: no daily cycle; Saturday 7 Nov has one',
      run: function () {
        tAt_('2026-11-08 12:00'); tRun_();
        return res_(!tCyc_('T01|2026-11-08') && !!tCyc_('T01|2026-11-07'),
          'Sun: ' + (tCyc_('T01|2026-11-08') ? 'yes' : 'no') + ', Sat: ' + (tCyc_('T01|2026-11-07') ? 'yes' : 'no'));
      } },
    { id: '3.6', name: 'Daily update through the employee tab (type, then tick)',
      expected: 'Text typed in the ✎ cell is kept as a draft; ticking Posted completes it; Share link formula present',
      run: function () {
        tAt_('2026-11-09 10:20'); runEngine();
        const sh = sheet_(String(tEmp_('E01')['Tab Name']));
        const id = 'T01|2026-11-09';
        let row = tRowOf_(sh, id, EMP_FIRST_ROW);
        if (!row) return res_(false, 'task not on tab');
        // getValue() returns the link's display text in real Sheets; the wa.me URL lives in the formula.
        const share = String(sh.getRange(row, EMP_COL.SHARE).getFormula());
        sh.getRange(row, EMP_COL.PROOF).setValue('Monday: 51 orders out, 2 RTOs, listing audit done');
        onEditInstalled({ range: sh.getRange(row, EMP_COL.PROOF) });
        const draft = tCyc_(id).Status;
        row = tRowOf_(sh, id, EMP_FIRST_ROW);
        sh.getRange(row, EMP_COL.POSTED).setValue(true);
        onEditInstalled({ range: sh.getRange(row, EMP_COL.POSTED) });
        const c = tCyc_(id);
        return res_(draft === ST.OPEN && c.Status === ST.DONE && c.Posted === true && share.indexOf('wa.me') >= 0,
          'after typing: ' + draft + '; after tick: ' + c.Status + '; share link ' + (share.indexOf('wa.me') >= 0 ? 'present' : 'missing'));
      } }
  ],

  // 4 – Test 4: same-day manager alert
  [
    { id: '4.1', name: 'No alert inside the grace period',
      expected: 'Mon 11:20, daily update missing: status OVERDUE, no manager email yet',
      run: function () {
        tReset_(); tAt_('2026-11-02 11:20'); tRun_();
        const c = tCyc_('T01|2026-11-02');
        return res_(c.Status === ST.OVERDUE && tNotifs_('L1').length === 0, c.Status + ', ' + tNotifs_('L1').length + ' alerts');
      } },
    { id: '4.2', name: 'Manager emailed the same day',
      expected: 'Mon 11:31: one email to the manager naming Employee A / Daily Operations Update',
      run: function () {
        tAt_('2026-11-02 11:31'); tRun_();
        const l1 = tNotifs_('L1');
        const c = tCyc_('T01|2026-11-02');
        return res_(l1.length === 1 && l1[0].CycleID === 'T01|2026-11-02' && c['L1 Sent At'] instanceof Date,
          l1.length + ' alert(s)' + (l1[0] ? ' to ' + l1[0].To + ': ' + l1[0].Details : ''));
      } },
    { id: '4.3', name: 'Later alert contains only the new misses',
      expected: 'Mon 17:31: a second digest with T02, T03, T04 – not T01 again',
      run: function () {
        tAt_('2026-11-02 17:31'); tRun_();
        const l1 = tNotifs_('L1');
        const refs = Array.from(new Set(l1.map(function (n) { return n['Email Ref']; })));
        const second = l1.filter(function (n) { return n['Email Ref'] === refs[1]; }).map(function (n) { return n.CycleID; });
        const t01 = l1.filter(function (n) { return n.CycleID === 'T01|2026-11-02'; }).length;
        return res_(refs.length === 2 && second.length === 3 && second.every(function (id) { return /^T0[234]\|/.test(id); }) && t01 === 1,
          refs.length + ' emails; second: ' + second.join(', '));
      } }
  ],

  // 5 – Test 5: manager action / acknowledgement
  [
    { id: '5.1', name: 'Acknowledged and then done → no escalation',
      expected: 'Manager picks "Reminded" 12:00, A completes 13:00 → DONE_LATE, Ack 12:00, no leadership email',
      run: function () {
        tReset_(); tAt_('2026-11-02 11:31'); tRun_();
        tAt_('2026-11-02 12:00'); apiManagerAction('T01|2026-11-02', ACTIONS.REMINDED);
        tAt_('2026-11-02 13:00'); apiSubmitDaily('T01|2026-11-02', MON_TEXT, true);
        tAt_('2026-11-02 15:32'); tRun_();
        tAt_('2026-11-02 19:01'); tRun_();
        const c = tCyc_('T01|2026-11-02');
        return res_(c.Status === ST.DONE_LATE && f_(c['Ack At']) === '2026-11-02 12:00' && tNotifFor_('L2', 'T01|2026-11-02').length === 0,
          c.Status + ', ack ' + f_(c['Ack At']) + ', L2: ' + tNotifFor_('L2', 'T01|2026-11-02').length);
      } },
    { id: '5.2', name: 'Waive needs a reason',
      expected: 'Waive with empty note → rejected; with note → WAIVED and acknowledged',
      run: function () {
        tAt_('2026-11-02 17:31'); tRun_();
        const r1 = apiManagerAction('T03|2026-11-02', ACTIONS.WAIVE, { note: '' });
        const r2 = apiManagerAction('T03|2026-11-02', ACTIONS.WAIVE, { note: 'Portal down all day, vendor ticket raised' });
        const c = tCyc_('T03|2026-11-02');
        return res_(!r1.ok && r2.ok && c.Status === ST.WAIVED && c['Ack At'] instanceof Date,
          (r1.reason || 'accepted') + ' → ' + c.Status);
      } },
    { id: '5.3', name: 'Extension moves the resolution deadline',
      expected: 'T02 extended to Wed 17:00: no escalation Tue 17:01; done Wed 16:00 → no escalation',
      run: function () {
        const r = apiManagerAction('T02|2026-11-02', ACTIONS.EXTEND, { newDue: '2026-11-04 17:00' });
        tAt_('2026-11-03 17:01'); tRun_();
        const l2a = tNotifFor_('L2', 'T02|2026-11-02').length;
        tAt_('2026-11-04 16:00'); apiCompleteTicket('T02|2026-11-02', 'PTA-10500');
        tAt_('2026-11-04 17:01'); tRun_();
        const c = tCyc_('T02|2026-11-02');
        return res_(r.ok && l2a === 0 && c.Status === ST.DONE_LATE && tNotifFor_('L2', 'T02|2026-11-02').length === 0,
          'extend ' + (r.ok ? 'ok' : r.reason) + '; L2 Tue: ' + l2a + '; final ' + c.Status);
      } },
    { id: '5.4', name: 'Extension beyond the limit is rejected',
      expected: 'Extending T04 to 20 Nov (more than 3 working days) → rejected',
      run: function () {
        const r = apiManagerAction('T04|2026-11-02', ACTIONS.EXTEND, { newDue: '2026-11-20 17:00' });
        return res_(!r.ok && /limited/.test(r.reason), r.reason || 'accepted');
      } },
    { id: '5.5', name: 'Dashboard summary matches the ledger',
      expected: 'Open / Overdue counts per employee on Manager_Dashboard equal the Cycles ledger',
      run: function () {
        runEngine();
        const sh = sheet_(SHEETS.DASH);
        const emps = readObjects_(SHEETS.EMP);
        const rows = sh.getRange(DASH_SUMMARY_ROW + 1, 2, emps.length, 3).getValues();
        const cyc = tCycles_();
        const bad = rows.filter(function (r) {
          const e = emps.filter(function (x) { return x.Name === r[0]; })[0];
          const mine = cyc.filter(function (c) { return c.EmpID === e.EmpID; });
          return r[1] !== mine.filter(function (c) { return c.Status === ST.OPEN; }).length ||
            r[2] !== mine.filter(function (c) { return c.Status === ST.OVERDUE; }).length;
        });
        return res_(rows.length === emps.length && bad.length === 0, rows.map(function (r) { return r[0] + ' ' + r[1] + '/' + r[2]; }).join('; '));
      } },
    { id: '5.6', name: 'Manager acts from the dashboard itself',
      expected: 'Note typed, then Action chosen on Manager_Dashboard → ledger has action, note, Ack At and Ack By',
      run: function () {
        const sh = sheet_(SHEETS.DASH);
        const id = 'T04|2026-11-02';
        const row = tRowOf_(sh, id, DASH_SUMMARY_ROW + 1);
        if (!row) return res_(false, 'item not on dashboard');
        sh.getRange(row, DASH_COL.NOTE).setValue('Called D, portal login issue');
        onEditInstalled({ range: sh.getRange(row, DASH_COL.NOTE) });
        const row2 = tRowOf_(sh, id, DASH_SUMMARY_ROW + 1);
        sh.getRange(row2, DASH_COL.ACTION).setValue(ACTIONS.REMINDED);
        onEditInstalled({ range: sh.getRange(row2, DASH_COL.ACTION) });
        const c = tCyc_(id);
        return res_(c['Manager Action'] === ACTIONS.REMINDED && c['Manager Note'] === 'Called D, portal login issue' &&
          c['Ack At'] instanceof Date && !!c['Ack By'],
          c['Manager Action'] + ' / ' + c['Manager Note'] + ' / ack ' + f_(c['Ack At']) + ' by ' + c['Ack By']);
      } }
  ],

  // 6 – Test 6: leadership escalation
  [
    { id: '6.1', name: 'Escalates when the manager does not act',
      expected: 'Alert 11:31, no action: nothing at 15:30; leadership email at 15:32 "Manager did not acknowledge"',
      run: function () {
        tReset_(); tAt_('2026-11-02 11:31'); tRun_();
        tAt_('2026-11-02 15:30'); tRun_();
        const early = tNotifFor_('L2', 'T01|2026-11-02').length;
        tAt_('2026-11-02 15:32'); tRun_();
        const l2 = tNotifFor_('L2', 'T01|2026-11-02');
        return res_(early === 0 && l2.length === 1 && l2[0].Details === 'Manager did not acknowledge',
          'at 15:30: ' + early + '; at 15:32: ' + (l2[0] ? l2[0].Details : 'none'));
      } },
    { id: '6.2', name: 'Escalates when acknowledged but still not done',
      expected: 'T02 acknowledged Tue 10:00, still open Tue 17:01 → "Acknowledged but unresolved"; T04 (no action) → "Manager did not acknowledge"',
      run: function () {
        tAt_('2026-11-02 17:31'); tRun_();
        tAt_('2026-11-03 10:00');
        apiManagerAction('T02|2026-11-02', ACTIONS.REMINDED);
        apiManagerAction('T03|2026-11-02', ACTIONS.REMINDED);
        tAt_('2026-11-03 17:01'); tRun_();
        const a = tNotifFor_('L2', 'T02|2026-11-02')[0], b = tNotifFor_('L2', 'T04|2026-11-02')[0];
        return res_(a && a.Details === 'Acknowledged but unresolved' && b && b.Details === 'Manager did not acknowledge',
          'T02: ' + (a ? a.Details : 'none') + '; T04: ' + (b ? b.Details : 'none'));
      } },
    { id: '6.3', name: 'Manager can escalate immediately',
      expected: 'Manager picks "Escalate now" on T05 → leadership email on the next run, "Escalated by manager"',
      run: function () {
        tAt_('2026-11-03 17:31'); tRun_();
        apiManagerAction('T05|2026-11-03', ACTIONS.ESCALATE);
        tAt_('2026-11-03 17:40'); tRun_();
        const l2 = tNotifFor_('L2', 'T05|2026-11-03')[0];
        return res_(l2 && l2.Details === 'Escalated by manager', l2 ? l2.Details : 'none');
      } },
    { id: '6.4', name: 'Leadership told once when an escalated item is resolved',
      expected: 'T02 completed after escalation → one "resolved" entry, not repeated on later runs',
      run: function () {
        tAt_('2026-11-04 11:00'); apiCompleteTicket('T02|2026-11-02', 'PTA-20001');
        tRun_(); tRun_();
        const n = tNotifs_().filter(function (x) { return x.NotifKey === 'RES|T02|2026-11-02'; }).length;
        return res_(n === 1, n + ' resolved notice(s)');
      } }
  ],

  // 7 – Test 7: escalation timing
  [
    { id: '7.1', name: 'Acknowledgement SLA spans the weekend correctly',
      expected: 'Alert Sat 17:30 → manager deadline Mon 12:30 (not Sunday)',
      run: function () { const d = ackDeadline(new Date(2026, 10, 7, 17, 30), loadConfig_()); return res_(fmtDateTime(d) === '2026-11-09 12:30', fmtDateTime(d)); } },
    { id: '7.2', name: 'No emails on a non-working day; they go out next working day',
      expected: 'First run Sun 8 Nov with a backlog: 0 emails. Mon 10:05: alerts sent',
      run: function () {
        tReset_(); tAt_('2026-11-08 12:00'); tRun_();
        const sun = tNotifs_().length;
        const overdue = tCycles_().filter(function (c) { return c.Status === ST.OVERDUE; }).length;
        tAt_('2026-11-09 10:05'); tRun_();
        const mon = tNotifs_('L1').length;
        return res_(sun === 0 && overdue > 0 && mon > 0, 'Sunday: ' + sun + ' emails (' + overdue + ' overdue); Monday: ' + mon + ' alert rows');
      } },
    { id: '7.3', name: 'Alert is always same-day',
      expected: 'Due 17:00 → alert 17:30; due 18:50 → alert capped at 19:00',
      run: function () {
        const cfg = loadConfig_();
        const a = hm(l1Time(new Date(2026, 10, 2, 17, 0), cfg)), b = fmtDateTime(l1Time(new Date(2026, 10, 2, 18, 50), cfg));
        return res_(a === '17:30' && b === '2026-11-02 19:00', a + ', ' + b);
      } },
    { id: '7.4', name: 'Repeated daily misses go to the weekly digest, not one email each',
      expected: 'Daily update late Mon–Wed (manager acted each time): no leadership escalation; Monday weekly digest flags Employee A',
      run: function () {
        tReset_();
        ['2026-11-02', '2026-11-03', '2026-11-04'].forEach(function (d) {
          tAt_(d + ' 11:31'); tRun_();
          apiManagerAction('T01|' + d, ACTIONS.REMINDED);
          tAt_(d + ' 12:00'); apiSubmitDaily('T01|' + d, 'Ops update for ' + d + ': orders dispatched, returns checked', true);
          tAt_(d + ' 19:01'); tRun_();
        });
        tAt_('2026-11-09 10:05'); tRun_();
        const w = tNotifs_('WEEKLY');
        const l2 = ['02', '03', '04'].filter(function (d) { return tNotifFor_('L2', 'T01|2026-11-' + d).length > 0; });
        const name = tEmp_('E01').Name;
        return res_(w.length === 1 && String(w[0].Details).indexOf(name + ' – PATTERN') >= 0 && l2.length === 0,
          'weekly: ' + (w[0] ? w[0].Details : 'none') + '; L2 for Mon–Wed: ' + l2.length);
      } }
  ],

  // 8 – Test 8: no duplicates
  [
    { id: '8.1', name: 'Running the engine repeatedly changes nothing',
      expected: 'Engine run 5× at Mon 17:31: cycle and notification counts unchanged after the first run',
      run: function () {
        tReset_(); tAt_('2026-11-02 17:31'); tRun_();
        const c0 = tCycles_().length, n0 = tNotifs_().length;
        tRun_(); tRun_(); tRun_(); tRun_();
        const c1 = tCycles_().length, n1 = tNotifs_().length;
        return res_(c0 === c1 && n0 === n1 && n0 > 0, 'cycles ' + c0 + '→' + c1 + ', notifications ' + n0 + '→' + n1);
      } },
    { id: '8.2', name: 'Same ticket entered twice = one completion',
      expected: 'Second identical entry → NO_CHANGE; one ACCEPTED audit row; Completed At unchanged',
      run: function () {
        tAt_('2026-11-02 17:40'); const r1 = apiCompleteTicket('T02|2026-11-02', 'PTA-30001');
        tAt_('2026-11-02 17:45'); const r2 = apiCompleteTicket('T02|2026-11-02', 'PTA-30001');
        const c = tCyc_('T02|2026-11-02');
        const acc = readObjects_(SHEETS.AUDIT).filter(function (a) { return a.CycleID === 'T02|2026-11-02' && a.Result === 'ACCEPTED'; }).length;
        return res_(r1.code === 'ACCEPTED' && r2.code === 'NO_CHANGE' && acc === 1 && f_(c['Completed At']) === '2026-11-02 17:40',
          r1.code + ' then ' + r2.code + '; accepted rows ' + acc + '; completed ' + f_(c['Completed At']));
      } },
    { id: '8.3', name: 'Same daily update submitted twice = one completion',
      expected: 'Second identical submission → NO_CHANGE',
      run: function () {
        tAt_('2026-11-03 10:15'); tRun_();
        const r1 = apiSubmitDaily('T01|2026-11-03', TUE_TEXT, true);
        const r2 = apiSubmitDaily('T01|2026-11-03', TUE_TEXT, true);
        return res_(r1.code === 'ACCEPTED' && r2.code === 'NO_CHANGE', r1.code + ' then ' + r2.code);
      } },
    { id: '8.4', name: 'Re-running setup creates no duplicate sheets, triggers or config rows',
      expected: 'Counts identical before and after setupWorkbook()',
      run: function () {
        const ss = SpreadsheetApp.getActive();
        const s0 = ss.getSheets().length, t0 = ScriptApp.getProjectTriggers().length, k0 = readObjects_(SHEETS.CONFIG).length;
        setupWorkbook();
        const s1 = ss.getSheets().length, t1 = ScriptApp.getProjectTriggers().length, k1 = readObjects_(SHEETS.CONFIG).length;
        return res_(s0 === s1 && t0 === t1 && k0 === k1, 'sheets ' + s0 + '→' + s1 + ', triggers ' + t0 + '→' + t1 + ', config ' + k0 + '→' + k1);
      } }
  ]
];
