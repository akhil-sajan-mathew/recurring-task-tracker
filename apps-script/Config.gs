/** Sheet names, headers, configuration defaults and the (simulated) clock. */

const SHEETS = {
  CONFIG: 'Config',
  EMP: 'Employees',
  HOL: 'Holidays',
  TASKS: 'Task_Master',
  CYCLES: 'Cycles',
  DASH: 'Manager_Dashboard',
  NOTIF: 'Notification_Log',
  RUN: 'Run_Log',
  AUDIT: 'Audit_Log',
  TESTS: 'Test_Results'
};

const EMP_TAB_PREFIX = 'My Tasks – ';

const HEADERS = {
  Config: ['Key', 'Value', 'Description'],
  Employees: ['EmpID', 'Name', 'Email', 'Manager Email', 'Tab Name', 'Active'],
  Holidays: ['Date', 'Name'],
  Task_Master: ['TaskID', 'Task Name', 'Assignee EmpID', 'Frequency', 'Cycle Days', 'Start Date',
    'Due Time', 'Lead Days', 'Proof Type', 'Proof Pattern', 'Active'],
  Cycles: ['CycleID', 'TaskID', 'Task Name', 'EmpID', 'Employee', 'Cycle No', 'Nominal Date', 'Open At',
    'Due At', 'Status', 'Proof', 'Posted', 'Completed By', 'Completed At', 'L1 Sent At', 'Ack At', 'Ack By',
    'Manager Action', 'Manager Note', 'Requested Due', 'Extended Due At', 'L2 Sent At', 'L2 Reason', 'Last Updated'],
  Notification_Log: ['NotifKey', 'Type', 'CycleID', 'To', 'Subject', 'Email Ref', 'Details', 'Created At', 'Status'],
  Run_Log: ['Run At', 'Simulated Now', 'Cycles Created', 'Status Changes', 'Emails', 'Duration ms', 'Notes / Error'],
  Audit_Log: ['When', 'Simulated Now', 'Who', 'CycleID', 'Field', 'Old', 'New', 'Result', 'Reason'],
  Test_Results: ['Test ID', 'Name', 'Expected', 'Actual', 'Result', 'Run At']
};

const CONFIG_DEFAULTS = [
  ['TIMEZONE', 'Asia/Kolkata', 'Must match File → Settings and the script project time zone'],
  ['WORK_DAYS', 'Mon,Tue,Wed,Thu,Fri,Sat', 'Working days; SLA clocks pause on other days'],
  ['WORK_START', '10:00', 'Start of working hours (HH:MM)'],
  ['WORK_END', '19:00', 'End of working hours (HH:MM)'],
  ['DAILY_DUE_TIME', '11:00', 'Default due time for DAILY tasks'],
  ['PORTAL_DUE_TIME', '17:00', 'Default due time for portal (ticket) tasks'],
  ['L1_GRACE_MIN', '30', 'Minutes after due time before the manager is alerted (capped at end of day)'],
  ['ACK_SLA_WORK_HRS', '4', 'Working hours the manager has to choose an action after the alert'],
  ['DAILY_RESOLVE_TIME', '19:00', 'A daily task must be done by this time on its own day'],
  ['PORTAL_RESOLVE_WORK_DAYS', '1', 'Working days after the due date by which a portal task must be done'],
  ['DAILY_L2_ON_UNRESOLVED', 'FALSE', 'FALSE: a late daily task that the manager acknowledged goes to the weekly digest, not an immediate escalation'],
  ['MAX_EXTEND_WORK_DAYS', '3', 'Furthest a manager may extend a deadline'],
  ['MIN_DAILY_UPDATE_CHARS', '20', 'Minimum length of a daily update'],
  ['REMINDER_ENABLED', 'TRUE', 'Email employees on the morning of a due date'],
  ['REMINDER_AFTER_START_MIN', '30', 'Reminder goes this many minutes after WORK_START'],
  ['PATTERN_MISSES', '3', 'Late/missed daily updates in a week that flag a pattern in the weekly digest'],
  ['BACKFILL_DAYS', '60', 'Never create cycles older than this (protects against a long-paused trigger)'],
  ['LEADERSHIP_EMAIL', '', 'Leadership escalations go here'],
  ['ADMIN_EMAIL', '', 'Fallback recipient when an employee has no manager email'],
  ['TEST_MODE', 'TRUE', 'Prefix email subjects with [TEST]'],
  ['DRY_RUN_EMAIL', 'FALSE', 'TRUE: log emails in Notification_Log without sending'],
  ['TEST_NOW', '', 'Simulated clock "YYYY-MM-DD HH:MM". Blank = real time']
];

function parseBool_(v) { return v === true || String(v).trim().toUpperCase() === 'TRUE'; }
function parseNum_(v, key) {
  const n = Number(v);
  if (String(v).trim() === '' || isNaN(n)) throw new Error('Config ' + key + ' must be a number, got "' + v + '"');
  return n;
}

function loadConfig_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  const raw = {};
  CONFIG_DEFAULTS.forEach(function (d) { raw[d[0]] = d[1]; });
  readObjects_(SHEETS.CONFIG).forEach(function (r) { raw[String(r.Key).trim()] = r.Value; });

  const dayIdx = {};
  DOW_NAMES.forEach(function (n, i) { dayIdx[n.toLowerCase()] = i; });
  const workDays = new Set(String(raw.WORK_DAYS).split(',').map(function (s) {
    const i = dayIdx[s.trim().toLowerCase().slice(0, 3)];
    if (i === undefined) throw new Error('Config WORK_DAYS has an unknown day "' + s + '"');
    return i;
  }));
  const holidays = new Set(readObjects_(SHEETS.HOL)
    .map(function (h) { return parseDateTime(h.Date); })
    .filter(Boolean)
    .map(ymd));

  return {
    raw: raw,
    tz: String(raw.TIMEZONE),
    workDays: workDays,
    holidays: holidays,
    workStart: parseHM(raw.WORK_START),
    workEnd: parseHM(raw.WORK_END),
    dailyDue: parseHM(raw.DAILY_DUE_TIME),
    portalDue: parseHM(raw.PORTAL_DUE_TIME),
    graceMin: parseNum_(raw.L1_GRACE_MIN, 'L1_GRACE_MIN'),
    ackHrs: parseNum_(raw.ACK_SLA_WORK_HRS, 'ACK_SLA_WORK_HRS'),
    dailyResolve: parseHM(raw.DAILY_RESOLVE_TIME),
    portalResolveDays: parseNum_(raw.PORTAL_RESOLVE_WORK_DAYS, 'PORTAL_RESOLVE_WORK_DAYS'),
    dailyL2OnUnresolved: parseBool_(raw.DAILY_L2_ON_UNRESOLVED),
    maxExtendDays: parseNum_(raw.MAX_EXTEND_WORK_DAYS, 'MAX_EXTEND_WORK_DAYS'),
    minDailyChars: parseNum_(raw.MIN_DAILY_UPDATE_CHARS, 'MIN_DAILY_UPDATE_CHARS'),
    reminderEnabled: parseBool_(raw.REMINDER_ENABLED),
    reminderAfterStartMin: parseNum_(raw.REMINDER_AFTER_START_MIN, 'REMINDER_AFTER_START_MIN'),
    patternMisses: parseNum_(raw.PATTERN_MISSES, 'PATTERN_MISSES'),
    backfillDays: parseNum_(raw.BACKFILL_DAYS, 'BACKFILL_DAYS'),
    leadershipEmail: String(raw.LEADERSHIP_EMAIL || '').trim(),
    adminEmail: String(raw.ADMIN_EMAIL || '').trim(),
    testMode: parseBool_(raw.TEST_MODE),
    dryRun: parseBool_(raw.DRY_RUN_EMAIL),
    testNow: parseDateTime(raw.TEST_NOW)
  };
}

/** Every piece of logic asks this for the time, so tests can replay days in minutes. */
function getNow(cfg) { return cfg.testNow ? new Date(cfg.testNow.getTime()) : new Date(); }

function setConfigValue(key, value) {
  const sh = sheet_(SHEETS.CONFIG);
  const last = sh.getLastRow();
  const keys = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  const i = keys.indexOf(key);
  const v = value instanceof Date ? fmtDateTime(value) : String(value);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(v);
  else sh.getRange(last + 1, 1, 1, 3).setValues([[key, v, '']]);
}
