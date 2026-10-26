/** Sheet I/O, the shared context object, locking and audit logging. */

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet "' + name + '" – run Allo Ops → Set up workbook');
  return sh;
}

function toObj_(headers, row) {
  const o = {};
  headers.forEach(function (h, i) { o[h] = row[i]; });
  return o;
}

/** Rows below the header as objects keyed by header name. Blank first column = blank row. */
function readObjects_(name) {
  const sh = sheet_(name);
  const h = HEADERS[name];
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, h.length).getValues()
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
    .map(function (r) { return toObj_(h, r); });
}

function writeObjects_(name, objs) {
  const sh = sheet_(name);
  const h = HEADERS[name];
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, h.length).clearContent();
  if (objs.length) {
    sh.getRange(2, 1, objs.length, h.length)
      .setValues(objs.map(function (o) { return h.map(function (k) { return safeCell_(o[k]); }); }));
  }
}

/** Appends rows and returns the first row number written. */
function appendRows_(name, rows) {
  const sh = sheet_(name);
  const start = sh.getLastRow() + 1;
  if (rows.length) {
    sh.getRange(start, 1, rows.length, rows[0].length)
      .setValues(rows.map(function (r) { return r.map(safeCell_); }));
  }
  return start;
}

function clearRows_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, HEADERS[name].length).clearContent();
}

/** User text that starts with = + - @ would otherwise be stored as a formula. */
function safeCell_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

/** Runs fn while holding the script lock, so engine runs and edits never interleave. */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another Allo Ops run is in progress – try again in a minute');
  try { return fn(); } finally { lock.releaseLock(); }
}

function loadContext_() {
  const ss = SpreadsheetApp.getActive();
  const cfg = loadConfig_(ss);

  const employees = readObjects_(SHEETS.EMP).map(function (e) {
    return {
      id: String(e.EmpID).trim(), name: String(e.Name).trim(), email: String(e.Email).trim(),
      managerEmail: String(e['Manager Email']).trim(), tab: String(e['Tab Name']).trim(), active: parseBool_(e.Active)
    };
  });
  const empById = {};
  employees.forEach(function (e) { empById[e.id] = e; });

  const tasks = {};
  readObjects_(SHEETS.TASKS).forEach(function (t) {
    tasks[String(t.TaskID).trim()] = {
      id: String(t.TaskID).trim(),
      name: String(t['Task Name']).trim(),
      empId: String(t['Assignee EmpID']).trim(),
      freq: String(t.Frequency).trim().toUpperCase(),
      cycleDays: Number(t['Cycle Days']) || 1,
      startDate: parseDateTime(t['Start Date']),
      dueTime: String(t['Due Time']).trim() ? parseHM(t['Due Time']) : null,
      leadDays: Number(t['Lead Days']) || 0,
      proofType: String(t['Proof Type']).trim().toUpperCase(),
      pattern: String(t['Proof Pattern'] || '').trim(),
      active: parseBool_(t.Active)
    };
  });

  const cycles = readObjects_(SHEETS.CYCLES);
  const byId = {};
  cycles.forEach(function (c) { c.CycleID = String(c.CycleID); byId[c.CycleID] = c; });

  const notifKeys = new Set();
  readObjects_(SHEETS.NOTIF).forEach(function (n) {
    if (String(n.Status).indexOf('FAILED') !== 0) notifKeys.add(String(n.NotifKey));
  });

  return {
    ss: ss, cfg: cfg, now: getNow(cfg), employees: employees, empById: empById, tasks: tasks,
    cycles: cycles, byId: byId, notifKeys: notifKeys, audit: [], emailSeq: 0
  };
}

function saveContext_(ctx) {
  writeObjects_(SHEETS.CYCLES, ctx.cycles);
  if (ctx.audit.length) {
    appendRows_(SHEETS.AUDIT, ctx.audit.map(function (a) {
      return [new Date(), ctx.cfg.testNow ? ctx.now : '', a.who || '', a.cycleId || '', a.field || '',
        a.old == null ? '' : a.old, a.new == null ? '' : a.new, a.result, a.reason || ''];
    }));
    ctx.audit = [];
  }
}

function audit_(ctx, cyc, field, oldV, newV, result, reason, who) {
  ctx.audit.push({ cycleId: cyc ? cyc.CycleID : '', field: field, old: oldV, new: newV, result: result, reason: reason, who: who });
}

function isDailyTask_(ctx, c) {
  const t = ctx.tasks[c.TaskID];
  return !!t && t.proofType === 'DAILY_UPDATE';
}
