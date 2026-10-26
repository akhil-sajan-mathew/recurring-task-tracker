/** Workbook setup, demo data, triggers, protections and the custom menu. Safe to re-run. */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Allo Ops')
    .addItem('1. Set up workbook', 'menuSetup')
    .addItem('2. Load demo data (this week)', 'menuLoadDemo')
    .addItem('2b. Load demo on the test calendar (Mon 2 Nov 2026)', 'menuLoadDemoTestCalendar')
    .addSeparator()
    .addItem('Run engine now', 'menuRunEngine')
    .addItem('Set test clock…', 'menuSetClock')
    .addItem('Advance test clock +1 hour', 'menuAdvanceHour')
    .addItem('Advance test clock +1 day', 'menuAdvanceDay')
    .addItem('Use real time (clear test clock)', 'menuClearClock')
    .addSeparator()
    .addItem('Run all tests', 'menuRunAllTests')
    .addItem('Run tests 1–4 only', 'menuRunTestsA')
    .addItem('Run tests 5–8 only', 'menuRunTestsB')
    .addItem('Reset cycles and logs', 'menuReset')
    .addToUi();
}

function setupWorkbook() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Asia/Kolkata');
  [SHEETS.CONFIG, SHEETS.EMP, SHEETS.HOL, SHEETS.TASKS, SHEETS.CYCLES, SHEETS.NOTIF, SHEETS.RUN, SHEETS.AUDIT, SHEETS.TESTS]
    .forEach(function (name) {
      const sh = ensureSheet_(name);
      const h = HEADERS[name];
      sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#cfe2f3');
      sh.setFrozenRows(1);
    });
  ensureSheet_(SHEETS.DASH);

  // Text format stops Sheets turning "10:00" into a 1899 time value and "TRUE" into a boolean.
  sheet_(SHEETS.CONFIG).getRange('B:B').setNumberFormat('@');
  sheet_(SHEETS.TASKS).getRange('G:G').setNumberFormat('@');
  sheet_(SHEETS.TASKS).getRange('J:J').setNumberFormat('@');
  sheet_(SHEETS.CYCLES).getRange('K:K').setNumberFormat('@');

  const existing = new Set(readObjects_(SHEETS.CONFIG).map(function (r) { return String(r.Key); }));
  const missing = CONFIG_DEFAULTS.filter(function (d) { return !existing.has(d[0]); });
  if (missing.length) appendRows_(SHEETS.CONFIG, missing);
  const me = Session.getEffectiveUser().getEmail();
  const cfgNow = loadConfig_(ss);
  if (!cfgNow.adminEmail && me) setConfigValue('ADMIN_EMAIL', me);

  const tasks = sheet_(SHEETS.TASKS);
  tasks.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['DAILY', 'EVERY_N_DAYS'], true).build());
  tasks.getRange('I2:I').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TICKET_ID', 'DAILY_UPDATE'], true).build());

  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  installTriggers_();
  applyProtections_();
}

function installTriggers_() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['scheduledRun', 'onEditInstalled'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scheduledRun').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('onEditInstalled').forSpreadsheet(ss).onEdit().create();
}

function aliasOf_(email, tag) { return email ? email.replace('@', '+' + tag + '@') : ''; }

/**
 * Demo employees (Gmail +aliases of the owner, so every email lands in one inbox) and tasks T01–T06.
 * start: the Monday the schedule starts. keepEmployees: leave an edited Employees sheet alone.
 */
function loadDemoData_(start, keepEmployees) {
  const me = Session.getEffectiveUser().getEmail();
  const emps = ['A', 'B', 'C', 'D'].map(function (x, i) {
    return { EmpID: 'E0' + (i + 1), Name: 'Employee ' + x, Email: aliasOf_(me, 'emp' + x.toLowerCase()),
      'Manager Email': aliasOf_(me, 'manager'), 'Tab Name': EMP_TAB_PREFIX + 'Employee ' + x, Active: true };
  });
  if (!keepEmployees || readObjects_(SHEETS.EMP).length === 0) writeObjects_(SHEETS.EMP, emps);
  if (!loadConfig_().leadershipEmail) setConfigValue('LEADERSHIP_EMAIL', aliasOf_(me, 'leadership'));

  const t = function (id, name, emp, freq, days, startOffset, lead, proof, pattern, active) {
    return { TaskID: id, 'Task Name': name, 'Assignee EmpID': emp, Frequency: freq, 'Cycle Days': days,
      'Start Date': addDays(start, startOffset), 'Due Time': '', 'Lead Days': lead, 'Proof Type': proof,
      'Proof Pattern': pattern, Active: active };
  };
  writeObjects_(SHEETS.TASKS, [
    t('T01', 'Daily Operations Update', 'E01', 'DAILY', 1, 0, 0, 'DAILY_UPDATE', '', true),
    t('T02', 'Portal Task A', 'E02', 'EVERY_N_DAYS', 7, 0, 1, 'TICKET_ID', '^PTA-\\d{4,8}$', true),
    t('T03', 'Portal Task B', 'E03', 'EVERY_N_DAYS', 15, 0, 2, 'TICKET_ID', '^PTB-\\d{4,8}$', true),
    t('T04', 'Portal Task C', 'E04', 'EVERY_N_DAYS', 30, 0, 3, 'TICKET_ID', '^PTC-\\d{4,8}$', true),
    // Extra cases: a second task for B (multi-task view + cross-task ticket uniqueness), and an inactive task.
    t('T05', 'Portal Task A2 (extra test case)', 'E02', 'EVERY_N_DAYS', 7, 1, 1, 'TICKET_ID', '^PTA-\\d{4,8}$', true),
    t('T06', 'Retired task (inactive test case)', 'E03', 'EVERY_N_DAYS', 7, 0, 0, 'TICKET_ID', '', false)
  ]);
  clearRows_(SHEETS.HOL);
  readObjects_(SHEETS.EMP).forEach(function (e) { ensureSheet_(String(e['Tab Name'])); });
}

function resetData_() {
  [SHEETS.CYCLES, SHEETS.NOTIF, SHEETS.AUDIT, SHEETS.RUN].forEach(clearRows_);
}

function applyProtections_() {
  const ss = SpreadsheetApp.getActive();
  const me = Session.getEffectiveUser().getEmail();
  ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).concat(ss.getProtections(SpreadsheetApp.ProtectionType.RANGE))
    .forEach(function (p) { if (String(p.getDescription()).indexOf('AlloOps') === 0) p.remove(); });

  const restrict = function (p, editors) {
    p.addEditor(me);
    p.removeEditors(p.getEditors().filter(function (u) { return u.getEmail() !== me; }));
    if (p.canDomainEdit()) p.setDomainEdit(false);
    // A +alias of the owner is the owner's own account, so adding it can fail; that is fine for the demo.
    editors.forEach(function (e) { try { if (e) p.addEditor(e); } catch (err) { /* not a separate Google account */ } });
  };

  [SHEETS.CONFIG, SHEETS.EMP, SHEETS.HOL, SHEETS.TASKS, SHEETS.CYCLES, SHEETS.NOTIF, SHEETS.RUN, SHEETS.AUDIT, SHEETS.TESTS]
    .forEach(function (name) { restrict(sheet_(name).protect().setDescription('AlloOps: admin only'), []); });

  const emps = readObjects_(SHEETS.EMP);
  emps.forEach(function (e) {
    const sh = ss.getSheetByName(String(e['Tab Name']));
    if (!sh) return;
    const input = sh.getRange(EMP_FIRST_ROW, EMP_COL.PROOF, 500, 2);
    restrict(sh.protect().setDescription('AlloOps: ' + e.Name + ' tab').setUnprotectedRanges([input]), []);
    restrict(input.protect().setDescription('AlloOps: ' + e.Name + ' input'), [String(e.Email)]);
  });

  const dash = ss.getSheetByName(SHEETS.DASH);
  if (dash) {
    const managers = Array.from(new Set(emps.map(function (e) { return String(e['Manager Email']); }).filter(Boolean)));
    const input = dash.getRange(1, DASH_COL.ACTION, 1000, 3);
    restrict(dash.protect().setDescription('AlloOps: dashboard').setUnprotectedRanges([input]), []);
    restrict(input.protect().setDescription('AlloOps: dashboard input'), managers);
  }
}

/* ───────────────────────── Menu handlers ───────────────────────── */

function mondayOf_(d) { return addDays(dayStart(d), -((d.getDay() + 6) % 7)); }

function menuSetup() {
  setupWorkbook();
  SpreadsheetApp.getActive().toast('Workbook ready. Next: Allo Ops → Load demo data.', 'Allo Ops', 6);
}

function menuLoadDemo() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Load demo data?', 'This replaces Task_Master and clears Cycles and the logs.', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  loadDemoData_(mondayOf_(new Date()), true);
  resetData_();
  applyProtections_();
  runEngine();
  SpreadsheetApp.getActive().toast('Demo data loaded and engine run.', 'Allo Ops', 6);
}

/** Same demo, pinned to the dates the tests and the walkthrough script use, with the clock at Mon 2 Nov 10:00. */
function menuLoadDemoTestCalendar() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Load test-calendar demo?', 'Replaces Task_Master, clears Cycles and the logs, and sets the test clock to 2026-11-02 10:00.',
    ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  loadDemoData_(new Date(2026, 10, 2), true);
  resetData_();
  applyProtections_();
  setConfigValue('TEST_NOW', '2026-11-02 10:00');
  menuRunEngine();
}

function menuRunEngine() {
  const s = runEngine();
  SpreadsheetApp.getActive().toast('Created ' + s.created + ' cycles, ' + s.changed + ' status changes, ' + s.emails + ' emails.', 'Engine run', 6);
}

function menuSetClock() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Set test clock', 'Simulated time as YYYY-MM-DD HH:MM (e.g. 2026-11-02 11:31)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  parseDateTime(r.getResponseText()); // throws on a bad format before anything is saved
  setConfigValue('TEST_NOW', r.getResponseText().trim());
  menuRunEngine();
}

function advanceClock_(minutes) {
  const cfg = loadConfig_();
  setConfigValue('TEST_NOW', addMinutes(getNow(cfg), minutes));
  menuRunEngine();
}
function menuAdvanceHour() { advanceClock_(60); }
function menuAdvanceDay() { advanceClock_(24 * 60); }
function menuClearClock() { setConfigValue('TEST_NOW', ''); menuRunEngine(); }

function confirmTests_() {
  const ui = SpreadsheetApp.getUi();
  return ui.alert('Run tests?', 'Tests wipe Cycles and the logs, use a simulated clock and send no real email. ' +
    'The workbook is reset to fresh demo data afterwards.', ui.ButtonSet.OK_CANCEL) === ui.Button.OK;
}
function menuRunAllTests() { if (confirmTests_()) toastResults_(runAllTests()); }
function menuRunTestsA() { if (confirmTests_()) toastResults_(runTestGroups([0, 1, 2, 3, 4])); }
function menuRunTestsB() { if (confirmTests_()) toastResults_(runTestGroups([5, 6, 7, 8])); }
function toastResults_(results) {
  const failed = results.filter(function (r) { return r.result === 'FAIL'; }).length;
  SpreadsheetApp.getActive().toast(results.length + ' tests, ' + failed + ' failed. See Test_Results.', 'Tests finished', 10);
  sheet_(SHEETS.TESTS).activate();
}

function menuReset() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Reset?', 'Clears Cycles and all logs.', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  resetData_();
  runEngine();
}
