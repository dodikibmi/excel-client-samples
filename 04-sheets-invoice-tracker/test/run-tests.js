// Runs the Apps Script logic in Node, without Google:
//   1. the unit tests in src/Tests.gs (the same ones "Run self-test" runs inside Google Sheets)
//   2. the whole demo (src/DemoData.gs) through invoiceState / planReminders / summarize, compared with
//      test/expected.json, which scripts/make_data.py worked out separately in Python.
// Run:  node test/run-tests.js      Exit code 1 if anything fails.
const fs = require('fs'), path = require('path'), vm = require('vm');
const src = (f) => fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8');
const ctx = vm.createContext({ console });
// const declarations stay inside a script; re-export what the test needs
vm.runInContext(src('Code.gs') + src('Tests.gs') + src('DemoData.gs') +
  '\nthis.api = {runAllTests, invoiceState, planReminders, summarize, buildReminder, DEMO};', ctx);
const { runAllTests, invoiceState, planReminders, summarize, DEMO } = ctx.api;
const expected = JSON.parse(fs.readFileSync(path.join(__dirname, 'expected.json'), 'utf8'));
let pass = 0, fail = 0;
const check = (name, ok, detail) => { ok ? pass++ : fail++; console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${ok || !detail ? '' : '  (' + detail + ')'}`); };

for (const t of runAllTests()) check('Unit: ' + t.name, t.pass, t.detail);

// demo: day 0 = today; any fixed date works because the data is relative
const today = new Date(2026, 10, 14);
const day = (o) => { const d = new Date(today); d.setDate(d.getDate() + o); return d; };
const inv = DEMO.invoices.map((i) => ({ no: i.no, customer: i.customer, amount: i.amount, paid: i.paid, due: day(i.due) }));
const bad = inv.filter((i) => {
  const s = invoiceState(i, today), e = expected.states[i.no];
  return Math.abs(s.balance - e.balance) > 0.004 || s.status !== e.status || s.days !== e.days || s.aging !== e.aging;
});
check(`Demo: state of all ${inv.length} invoices = Python`, bad.length === 0, bad.map((b) => b.no).join(','));

const customers = {}; DEMO.customers.forEach((c) => { customers[c.name] = { email: c.email, contact: c.contact }; });
const sent = {}; DEMO.log.forEach((e) => { (sent[e.no] = sent[e.no] || []).push(e.stage); });
const plan = planReminders(inv, customers, sent, { stages: DEMO.stages, maxPerRun: DEMO.maxPerRun }, today);
const got = JSON.stringify(plan.map((p) => [p.no, p.stage, p.action]));
const want = JSON.stringify(expected.plan.map((p) => [p.no, p.stage, p.action]));
check(`Demo: reminder plan (${plan.length} lines) = Python`, got === want, got + ' vs ' + want);
check('Demo: no invoice gets a stage it already had', plan.every((p) => !(sent[p.no] || []).includes(p.stage)));

const sum = summarize(inv.map((i) => ({ no: i.no, customer: i.customer, state: invoiceState(i, today) })));
const es = expected.summary;
check('Demo: outstanding, overdue and aging = Python',
  Math.abs(sum.outstanding - es.outstanding) < 0.005 && Math.abs(sum.overdue - es.overdue) < 0.005 && sum.overdueCount === es.overdue_count &&
  Object.keys(es.aging).every((b) => Math.abs(sum.aging[b] - es.aging[b]) < 0.005), JSON.stringify([sum.outstanding, sum.overdue, sum.aging]));

console.log(`\n${pass}/${pass + fail} checks passed`);
process.exit(fail ? 1 : 0);
