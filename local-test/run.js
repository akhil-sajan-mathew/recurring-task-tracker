// Runs the Apps Script test suite locally against the in-memory mock.
// Usage: node local-test/run.js
process.env.TZ = 'Asia/Kolkata';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const srcDir = path.join(__dirname, '..', 'apps-script');
const files = fs.readdirSync(srcDir).filter((f) => f.endsWith('.gs')).sort();
const code = [fs.readFileSync(path.join(__dirname, 'mock.js'), 'utf8')]
  .concat(files.map((f) => '// ---- ' + f + '\n' + fs.readFileSync(path.join(srcDir, f), 'utf8')))
  .join('\n;\n') + `
;
setupWorkbook();
loadDemoData_(new Date(2026, 10, 2), false);
const __results = runAllTests();
const __testMail = __mail.length;

// Real-send path (the suite itself runs with DRY_RUN_EMAIL on): send, no resend, failure then retry.
loadDemoData_(new Date(2026, 10, 2), true); resetData_();
setConfigValue('DRY_RUN_EMAIL', 'FALSE'); setConfigValue('TEST_NOW', '2026-11-02 11:31');
__mail.length = 0;
runEngine();
const __sent1 = __mail.length, __status1 = readObjects_(SHEETS.NOTIF).map(function (n) { return n.Status; }).join();
runEngine();
const __sent2 = __mail.length;
const __realSend = MailApp.sendEmail;
MailApp.sendEmail = function () { throw new Error('quota'); };
setConfigValue('TEST_NOW', '2026-11-02 17:31'); runEngine();
const __failed = readObjects_(SHEETS.NOTIF).filter(function (n) { return String(n.Status).indexOf('FAILED') === 0; }).length;
MailApp.sendEmail = __realSend;
setConfigValue('TEST_NOW', '2026-11-02 17:46'); runEngine();
const __retried = __mail.length - __sent2;
({ results: __results, mail: __mail, testMail: __testMail, lockHeld: __lockHeld, run: readObjects_(SHEETS.RUN),
   send: { sent1: __sent1, status1: __status1, sent2: __sent2, failed: __failed, retried: __retried, subject: __mail[0] && __mail[0].subject } });
`;

const ctx = vm.createContext({ console });
const out = vm.runInContext(code, ctx, { filename: 'bundle.js' });

let failed = 0;
for (const r of out.results) {
  if (r.result === 'FAIL') failed++;
  console.log(`${r.result.padEnd(6)} ${r.id.padEnd(4)} ${r.name}\n         ${r.actual}`);
}
console.log(`\n${out.results.length} tests, ${failed} failed, ${out.results.filter((r) => r.result === 'MANUAL').length} manual`);
console.log(`real emails sent during the suite: ${out.testMail}; lock left held: ${out.lockHeld}`);
const s = out.send;
// 11:31: L1 for A + morning reminders for B, C, D = 4. 17:31 (mail down): L1 digest (3 rows) + L2 for T01 (1 row)
// fail; 17:46 retries those 2 emails.
const sendOk = s.sent1 === 4 && s.status1 === 'SENT,SENT,SENT,SENT' && s.sent2 === 4 && s.failed === 4 && s.retried === 2;
console.log(`${sendOk ? 'PASS' : 'FAIL'}   send path: first run ${s.sent1} email (${s.status1}), re-run total ${s.sent2}, ` +
  `failed rows ${s.failed}, retried ${s.retried} – "${s.subject}"`);
if (!sendOk) failed++;
const errors = out.run.filter((r) => String(r['Notes / Error']).startsWith('ERROR'));
if (errors.length) console.log('engine errors:', errors.map((e) => e['Notes / Error']).join('\n'));
process.exit(failed ? 1 : 0);
