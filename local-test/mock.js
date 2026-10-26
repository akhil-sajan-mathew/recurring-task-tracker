// Minimal in-memory stand-in for the Apps Script services this project uses.
// Loaded into the same VM context as the .gs files (so Dates and Arrays share a realm).
// shortcut: models values only (no formatting, formulas are stored as text); good enough to exercise the logic.

var __mail = [];
var __triggers = [];
var __props = {};
var __lockHeld = false;
var __sheetSeq = 0;

function __isDate(v) { return Object.prototype.toString.call(v) === '[object Date]'; }
function __store(v) {
  if (__isDate(v)) return new Date(v.getTime());
  if (typeof v === 'string' && v.charAt(0) === "'") return v.slice(1); // Sheets strips the text-forcing apostrophe
  return v === undefined || v === null ? '' : v;
}
function __read(v) { return __isDate(v) ? new Date(v.getTime()) : v; }

function MockProtection(target) { this.target = target; this.desc = ''; this.editors = ['owner@gmail.com']; }
MockProtection.prototype.setDescription = function (d) { this.desc = d; return this; };
MockProtection.prototype.getDescription = function () { return this.desc; };
MockProtection.prototype.addEditor = function (e) { if (this.editors.indexOf(e) < 0) this.editors.push(e); return this; };
MockProtection.prototype.removeEditors = function (list) {
  const drop = list.map(function (u) { return typeof u === 'string' ? u : u.getEmail(); });
  this.editors = this.editors.filter(function (e) { return drop.indexOf(e) < 0; });
  return this;
};
MockProtection.prototype.getEditors = function () { return this.editors.map(function (e) { return { getEmail: function () { return e; } }; }); };
MockProtection.prototype.canDomainEdit = function () { return false; };
MockProtection.prototype.setDomainEdit = function () { return this; };
MockProtection.prototype.setUnprotectedRanges = function () { return this; };
MockProtection.prototype.remove = function () {
  const all = __ss.protections; const i = all.indexOf(this); if (i >= 0) all.splice(i, 1);
};

function MockRange(sheet, r, c, nr, nc) {
  if (r < 1 || c < 1 || nr < 1 || nc < 1) throw new Error('Bad range ' + [r, c, nr, nc].join(','));
  this.sheet = sheet; this.r = r; this.c = c; this.nr = nr; this.nc = nc;
}
MockRange.prototype.getValues = function () {
  const out = [];
  for (let i = 0; i < this.nr; i++) {
    const row = [];
    for (let j = 0; j < this.nc; j++) row.push(__read(this.sheet.get(this.r + i, this.c + j)));
    out.push(row);
  }
  return out;
};
MockRange.prototype.setValues = function (vals) {
  if (vals.length !== this.nr || vals.some(function (r) { return r.length !== this.nc; }, this)) {
    throw new Error('setValues size mismatch: range ' + this.nr + 'x' + this.nc + ', data ' + vals.length + 'x' + (vals[0] || []).length);
  }
  for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet.set(this.r + i, this.c + j, __store(vals[i][j]));
  return this;
};
MockRange.prototype.getValue = function () { return __read(this.sheet.get(this.r, this.c)); };
// Real Sheets: getValue() of a HYPERLINK formula returns its display text, getFormula() the formula.
MockRange.prototype.getFormula = function () {
  const v = this.sheet.get(this.r, this.c);
  return typeof v === 'string' && v.charAt(0) === '=' ? v : '';
};
MockRange.prototype.setValue = function (v) {
  for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet.set(this.r + i, this.c + j, __store(v));
  return this;
};
MockRange.prototype.clearContent = function () {
  for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet.set(this.r + i, this.c + j, '');
  return this;
};
MockRange.prototype.insertCheckboxes = function () { return this.setValue(false); };
MockRange.prototype.setNote = function (n) { this.sheet.notes[this.r + ',' + this.c] = n; return this; };
MockRange.prototype.getNote = function () { return this.sheet.notes[this.r + ',' + this.c] || ''; };
MockRange.prototype.clearNote = function () {
  for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) delete this.sheet.notes[(this.r + i) + ',' + (this.c + j)];
  return this;
};
['clearDataValidations', 'setNumberFormat', 'setBackground', 'setBackgrounds', 'setFontWeight', 'setFontSize', 'setDataValidation', 'activate']
  .forEach(function (m) { MockRange.prototype[m] = function () { return this; }; });
MockRange.prototype.getRow = function () { return this.r; };
MockRange.prototype.getColumn = function () { return this.c; };
MockRange.prototype.getNumRows = function () { return this.nr; };
MockRange.prototype.getNumColumns = function () { return this.nc; };
MockRange.prototype.getSheet = function () { return this.sheet; };
MockRange.prototype.protect = function () { const p = new MockProtection(this); p.type = 'RANGE'; __ss.protections.push(p); return p; };

function MockSheet(name) { this.name = name; this.cells = {}; this.notes = {}; this.id = ++__sheetSeq; }
MockSheet.prototype.get = function (r, c) { const v = this.cells[r + ',' + c]; return v === undefined ? '' : v; };
MockSheet.prototype.set = function (r, c, v) { if (v === '') delete this.cells[r + ',' + c]; else this.cells[r + ',' + c] = v; };
MockSheet.prototype.getName = function () { return this.name; };
MockSheet.prototype.getSheetId = function () { return this.id; };
MockSheet.prototype.getLastRow = function () {
  let m = 0; Object.keys(this.cells).forEach(function (k) { m = Math.max(m, +k.split(',')[0]); }); return m;
};
MockSheet.prototype.getLastColumn = function () {
  let m = 0; Object.keys(this.cells).forEach(function (k) { m = Math.max(m, +k.split(',')[1]); }); return m;
};
MockSheet.prototype.getRange = function (r, c, nr, nc) {
  if (typeof r === 'string') {
    // Column-style A1 ("B:B", "D2:D") – only used for formatting/validation, which the mock ignores.
    const m = r.match(/^([A-Z]+)(\d*):([A-Z]+)$/);
    if (!m) throw new Error('Mock getRange A1 unsupported: ' + r);
    const col = m[1].charCodeAt(0) - 64;
    return new MockRange(this, m[2] ? +m[2] : 1, col, 1000, 1);
  }
  return new MockRange(this, r, c, nr || 1, nc || 1);
};
MockSheet.prototype.getDataRange = function () { return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); };
MockSheet.prototype.setFrozenRows = function () { return this; };
MockSheet.prototype.hideColumns = function () { return this; };
MockSheet.prototype.activate = function () { return this; };
MockSheet.prototype.protect = function () { const p = new MockProtection(this); p.type = 'SHEET'; __ss.protections.push(p); return p; };

var __ss = {
  sheets: [new MockSheet('Sheet1')],
  protections: [],
  getSheetByName: function (n) { return this.sheets.filter(function (s) { return s.name === n; })[0] || null; },
  insertSheet: function (n) {
    if (this.getSheetByName(n)) throw new Error('Sheet exists: ' + n);
    const s = new MockSheet(n); this.sheets.push(s); return s;
  },
  getSheets: function () { return this.sheets.slice(); },
  deleteSheet: function (s) { this.sheets.splice(this.sheets.indexOf(s), 1); },
  getUrl: function () { return 'https://docs.google.com/spreadsheets/d/MOCK/edit'; },
  setSpreadsheetTimeZone: function () {},
  toast: function () {},
  getProtections: function (type) { return this.protections.filter(function (p) { return p.type === type; }); }
};

var SpreadsheetApp = {
  getActive: function () { return __ss; },
  flush: function () {},
  ProtectionType: { SHEET: 'SHEET', RANGE: 'RANGE' },
  newDataValidation: function () {
    const b = {};
    ['requireValueInList', 'requireDate', 'setAllowInvalid'].forEach(function (m) { b[m] = function () { return b; }; });
    b.build = function () { return {}; };
    return b;
  },
  getUi: function () { throw new Error('No UI in local tests'); }
};

var LockService = {
  getScriptLock: function () {
    return {
      // Not re-entrant, like a second execution: a nested withLock_ shows up as a failure here.
      tryLock: function () { if (__lockHeld) return false; __lockHeld = true; return true; },
      releaseLock: function () { __lockHeld = false; }
    };
  }
};

var PropertiesService = {
  getScriptProperties: function () {
    return {
      getProperty: function (k) { return __props.hasOwnProperty(k) ? __props[k] : null; },
      setProperty: function (k, v) { __props[k] = String(v); },
      deleteProperty: function (k) { delete __props[k]; }
    };
  }
};

var MailApp = {
  sendEmail: function (msg) { __mail.push(msg); },
  getRemainingDailyQuota: function () { return 100; }
};

var Session = { getEffectiveUser: function () { return { getEmail: function () { return 'owner@gmail.com'; } }; } };

var ScriptApp = {
  getProjectTriggers: function () { return __triggers.slice(); },
  deleteTrigger: function (t) { __triggers.splice(__triggers.indexOf(t), 1); },
  newTrigger: function (fn) {
    const t = { getHandlerFunction: function () { return fn; } };
    const b = {
      timeBased: function () { return b; }, everyMinutes: function () { return b; },
      forSpreadsheet: function () { return b; }, onEdit: function () { return b; },
      create: function () { __triggers.push(t); return t; }
    };
    return b;
  }
};
