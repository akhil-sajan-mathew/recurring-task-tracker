/** Completion capture (employee tabs) and manager actions (dashboard). All writes land in the Cycles ledger. */

function ok_(code) { return { ok: true, code: code || 'ACCEPTED' }; }
function reject_(ctx, cyc, field, value, reason, who) {
  audit_(ctx, cyc, field, '', value, 'REJECTED', reason, who);
  return { ok: false, code: 'REJECTED', reason: reason };
}

/** Installable onEdit trigger (installed by Set up workbook). */
function onEditInstalled(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const name = sh.getName();
  const isDash = name === SHEETS.DASH;
  if (!isDash && name.indexOf(EMP_TAB_PREFIX) !== 0) return;
  // On consumer Gmail, e.user is often blank for editors other than the owner; the protected tab is the identity.
  const who = (e.user && e.user.getEmail && e.user.getEmail()) || '';

  withLock_(function () {
    const ctx = loadContext_();
    const emp = isDash ? null : ctx.employees.filter(function (x) { return x.tab === name; })[0];
    if (!isDash && !emp) return;

    const r0 = e.range.getRow(), nr = e.range.getNumRows(), c0 = e.range.getColumn(), nc = e.range.getNumColumns();
    const width = isDash ? DASH_HEADERS.length : EMP_HEADERS.length;
    const rows = sh.getRange(r0, 1, nr, width).getValues();
    const rejections = [];
    rows.forEach(function (vals) {
      const cyc = ctx.byId[String(vals[0])];
      if (!cyc) return;
      for (let col = c0; col < c0 + nc; col++) {
        const res = isDash ? handleDashEdit_(ctx, cyc, col, vals, who) : handleEmpEdit_(ctx, emp, cyc, col, vals, who);
        if (res && !res.ok) rejections.push({ id: cyc.CycleID, col: col, reason: res.reason });
      }
    });
    saveContext_(ctx);

    // Re-render from the ledger: accepted edits show their new status, rejected edits snap back.
    const pos = isDash ? refreshViews_(ctx).dash : (function () { const p = refreshEmployeeTab_(ctx, emp); refreshDashboard_(ctx); return p; })();
    rejections.forEach(function (r) {
      if (pos[r.id]) sh.getRange(pos[r.id], r.col).setNote('Not saved: ' + r.reason);
    });
    if (rejections.length) ctx.ss.toast(rejections[0].reason, 'Not saved', 8);
  });
}

function handleEmpEdit_(ctx, emp, cyc, col, vals, who) {
  if (col !== EMP_COL.PROOF && col !== EMP_COL.POSTED) return null;
  if (cyc.EmpID !== emp.id) return reject_(ctx, cyc, 'Proof', vals[EMP_COL.PROOF - 1], 'This task belongs to another employee', who);
  const by = emp.name + (who ? ' (' + who + ')' : '');
  const proof = vals[EMP_COL.PROOF - 1];
  if (isDailyTask_(ctx, cyc)) return submitDaily_(ctx, cyc, proof, vals[EMP_COL.POSTED - 1] === true, by);
  if (col !== EMP_COL.PROOF) return null;
  return completeTicket_(ctx, cyc, proof, by);
}

function completeTicket_(ctx, cyc, raw, by) {
  const task = ctx.tasks[cyc.TaskID];
  const ticket = String(raw == null ? '' : raw).trim().toUpperCase();
  if (!task || task.proofType !== 'TICKET_ID') return reject_(ctx, cyc, 'Proof', ticket, 'This task does not take a Ticket ID', by);
  if (cyc.Status === ST.WAIVED) return reject_(ctx, cyc, 'Proof', ticket, 'This cycle was waived by the manager', by);
  if (!ticket) {
    if (isClosed(cyc)) return reject_(ctx, cyc, 'Proof', '', 'A completed task cannot be cleared – ask your manager', by);
    return null;
  }
  if (ctx.now < cyc['Open At']) return reject_(ctx, cyc, 'Proof', ticket, 'This cycle is not open yet', by);
  if (task.pattern && !new RegExp(task.pattern).test(ticket)) {
    return reject_(ctx, cyc, 'Proof', ticket, 'Ticket ID has the wrong format (expected ' + task.pattern + ')', by);
  }
  const dup = ctx.cycles.filter(function (c) {
    return c.CycleID !== cyc.CycleID && String(c.Proof).trim().toUpperCase() === ticket && !isDailyTask_(ctx, c);
  })[0];
  if (dup) return reject_(ctx, cyc, 'Proof', ticket, 'Ticket ' + ticket + ' is already used for ' + dup.CycleID, by);

  if (isClosed(cyc)) {
    if (String(cyc.Proof).toUpperCase() === ticket) { audit_(ctx, cyc, 'Proof', cyc.Proof, ticket, 'NO_CHANGE', '', by); return ok_('NO_CHANGE'); }
    audit_(ctx, cyc, 'Proof', cyc.Proof, ticket, 'CORRECTION', 'Ticket corrected after completion', by);
    cyc.Proof = ticket;
    cyc['Last Updated'] = ctx.now;
    return ok_('CORRECTION');
  }
  cyc.Proof = ticket;
  cyc['Completed By'] = by;
  cyc['Completed At'] = ctx.now;
  cyc.Status = statusOf(cyc, ctx.now);
  cyc['Last Updated'] = ctx.now;
  audit_(ctx, cyc, 'Proof', '', ticket, 'ACCEPTED', cyc.Status, by);
  return ok_();
}

function submitDaily_(ctx, cyc, raw, posted, by) {
  const text = String(raw == null ? '' : raw).trim();
  if (!isDailyTask_(ctx, cyc)) return reject_(ctx, cyc, 'Proof', text, 'This task is not a daily update', by);
  if (cyc.Status === ST.WAIVED) return reject_(ctx, cyc, 'Proof', text, 'This cycle was waived by the manager', by);

  if (isClosed(cyc)) {
    if (!posted) return reject_(ctx, cyc, 'Posted', false, 'A completed daily update cannot be unticked – ask your manager', by);
    if (normalizeText(text) === normalizeText(cyc.Proof)) return ok_('NO_CHANGE');
    if (text.length < ctx.cfg.minDailyChars) return reject_(ctx, cyc, 'Proof', text, 'Update must be at least ' + ctx.cfg.minDailyChars + ' characters', by);
    audit_(ctx, cyc, 'Proof', cyc.Proof, text, 'CORRECTION', 'Update text edited after completion', by);
    cyc.Proof = text;
    cyc['Last Updated'] = ctx.now;
    return ok_('CORRECTION');
  }
  if (!posted) {
    // Typing before ticking is normal: keep the draft in the ledger so the next refresh does not wipe it.
    if (text !== String(cyc.Proof)) { cyc.Proof = text; cyc['Last Updated'] = ctx.now; audit_(ctx, cyc, 'Proof', '', text, 'DRAFT', '', by); }
    return ok_('DRAFT');
  }
  if (ctx.now < cyc['Open At']) return reject_(ctx, cyc, 'Posted', true, 'This cycle is not open yet', by);
  if (text.length < ctx.cfg.minDailyChars) {
    return reject_(ctx, cyc, 'Posted', true, 'Write the update (at least ' + ctx.cfg.minDailyChars + ' characters) before ticking Posted', by);
  }
  const prev = ctx.cycles.filter(function (c) {
    return c.TaskID === cyc.TaskID && c.CycleID !== cyc.CycleID && c['Completed At'] instanceof Date && c['Nominal Date'] < cyc['Nominal Date'];
  }).sort(function (a, b) { return b['Nominal Date'] - a['Nominal Date']; })[0];
  if (prev && normalizeText(prev.Proof) === normalizeText(text)) {
    return reject_(ctx, cyc, 'Posted', true, 'This is the same text as the previous update (' + ymd(prev['Nominal Date']) + ')', by);
  }
  cyc.Proof = text;
  cyc.Posted = true;
  cyc['Completed By'] = by;
  cyc['Completed At'] = ctx.now;
  cyc.Status = statusOf(cyc, ctx.now);
  cyc['Last Updated'] = ctx.now;
  audit_(ctx, cyc, 'Posted', false, true, 'ACCEPTED', cyc.Status, by);
  return ok_();
}

function handleDashEdit_(ctx, cyc, col, vals, who) {
  const by = who || 'Manager';
  const note = String(vals[DASH_COL.NOTE - 1] || '').trim();
  if (col === DASH_COL.NOTE) {
    if (note !== String(cyc['Manager Note'])) {
      audit_(ctx, cyc, 'Manager Note', cyc['Manager Note'], note, 'ACCEPTED', '', by);
      cyc['Manager Note'] = note;
    }
    return ok_();
  }
  if (col === DASH_COL.NEWDUE) {
    const v = vals[DASH_COL.NEWDUE - 1];
    if (cyc['Manager Action'] === ACTIONS.EXTEND && v) return managerAction_(ctx, cyc, ACTIONS.EXTEND, { note: note, newDue: v }, by);
    cyc['Requested Due'] = v || '';
    return ok_('DRAFT');
  }
  if (col === DASH_COL.ACTION) {
    const action = String(vals[DASH_COL.ACTION - 1] || '').trim();
    if (!action) return null;
    return managerAction_(ctx, cyc, action, { note: note, newDue: vals[DASH_COL.NEWDUE - 1] || cyc['Requested Due'] }, by);
  }
  return null;
}

/** The manager's action IS the acknowledgement: it forces a decision instead of a bare "seen" tick. */
function managerAction_(ctx, cyc, action, opts, by) {
  opts = opts || {};
  const now = ctx.now;
  const note = String(opts.note || '').trim();
  const valid = Object.keys(ACTIONS).map(function (k) { return ACTIONS[k]; });
  if (valid.indexOf(action) < 0) return reject_(ctx, cyc, 'Manager Action', action, 'Unknown action', by);
  if (isClosed(cyc)) return reject_(ctx, cyc, 'Manager Action', action, 'This cycle is already closed (' + cyc.Status + ')', by);
  if (action === ACTIONS.WAIVE && !note) return reject_(ctx, cyc, 'Manager Action', action, 'Type a reason in Note first, then choose Waive', by);

  if (action === ACTIONS.EXTEND) {
    let d = parseDateTime(opts.newDue);
    if (!d) return reject_(ctx, cyc, 'Manager Action', action, 'Enter the New Due date first, then choose Extend', by);
    if (d.getHours() === 0 && d.getMinutes() === 0) d = atTime(d, { h: cyc['Due At'].getHours(), m: cyc['Due At'].getMinutes() });
    const max = atTime(addWorkingDays(now, ctx.cfg.maxExtendDays, ctx.cfg), ctx.cfg.workEnd);
    if (d <= now) return reject_(ctx, cyc, 'Extended Due At', fmtDateTime(d), 'New due date must be in the future', by);
    if (d > max) return reject_(ctx, cyc, 'Extended Due At', fmtDateTime(d), 'Extensions are limited to ' + ctx.cfg.maxExtendDays + ' working days (latest ' + fmtDateTime(max) + ')', by);
    audit_(ctx, cyc, 'Extended Due At', cyc['Extended Due At'], d, 'ACCEPTED', '', by);
    cyc['Extended Due At'] = d;
    cyc['Requested Due'] = '';
  }
  if (!(cyc['Ack At'] instanceof Date)) { cyc['Ack At'] = now; cyc['Ack By'] = by; }
  audit_(ctx, cyc, 'Manager Action', cyc['Manager Action'], action, 'ACCEPTED', note, by);
  cyc['Manager Action'] = action;
  if (note) cyc['Manager Note'] = note;
  if (action === ACTIONS.WAIVE) cyc.Status = ST.WAIVED;
  else cyc.Status = statusOf(cyc, now);
  cyc['Last Updated'] = now;
  return ok_();
}

/* Public entry points (used by the tests; same code path as the sheet edits). */

function apiCall_(cycleIdStr, fn) {
  return withLock_(function () {
    const ctx = loadContext_();
    const c = ctx.byId[cycleIdStr];
    if (!c) return { ok: false, code: 'REJECTED', reason: 'Cycle ' + cycleIdStr + ' not found' };
    const r = fn(ctx, c) || ok_('NO_CHANGE');
    saveContext_(ctx);
    return r;
  });
}
function apiCompleteTicket(id, ticket, by) {
  return apiCall_(id, function (ctx, c) { return completeTicket_(ctx, c, ticket, by || c.Employee); });
}
function apiSubmitDaily(id, text, posted, by) {
  return apiCall_(id, function (ctx, c) { return submitDaily_(ctx, c, text, posted, by || c.Employee); });
}
function apiManagerAction(id, action, opts, by) {
  return apiCall_(id, function (ctx, c) { return managerAction_(ctx, c, action, opts, by || 'Manager'); });
}
