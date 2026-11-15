// Runs the Apps Script logic in Node, without Google:
//   1. the unit tests in src/Tests.gs (the same ones "Run self-test" runs inside Google Sheets)
//   2. the whole demo (src/DemoData.gs) through the script logic, compared with test/expected.json, which
//      scripts/make_data.py worked out separately in Python: invoice states, pauses, reminder plan, totals,
//      6-month trends, DSO, collection rate, customer risk scores and the owner's action queue.
// Run:  node test/run-tests.js      Exit code 1 if anything fails.
const fs = require('fs'), path = require('path'), vm = require('vm');
const src = (f) => fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8');
const ctx = vm.createContext({ console });
// const declarations stay inside a script; re-export what the test needs
vm.runInContext(src('Code.gs') + src('Tests.gs') + src('DemoData.gs') +
  '\nthis.api = {runAllTests, invoiceState, planReminders, summarize, buildReminder, pauseState, trendPoints, kpiCards, riskScores, actionQueue, DEMO};', ctx);
const { runAllTests, invoiceState, planReminders, summarize, pauseState, trendPoints, kpiCards, riskScores, actionQueue, DEMO } = ctx.api;
const expected = JSON.parse(fs.readFileSync(path.join(__dirname, 'expected.json'), 'utf8'));
let pass = 0, fail = 0;
const check = (name, ok, detail) => { ok ? pass++ : fail++; console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${ok || !detail ? '' : '  (' + detail + ')'}`); };
const near = (a, b, tol) => Math.abs(a - b) < (tol || 0.005);

for (const t of runAllTests()) check('Unit: ' + t.name, t.pass, t.detail);

// demo: day 0 = today; Python works with the same day (14 Nov 2026) so the month ends match
const VDate = vm.runInContext("Date", ctx);   // dates must come from the script's own realm, or instanceof Date fails
const today = new VDate(2026, 10, 14);
const day = (o) => { const d = new VDate(today); d.setDate(d.getDate() + o); return d; };
const inv = DEMO.invoices.map((i) => ({ no: i.no, customer: i.customer, amount: i.amount, paid: i.paid, issue: day(i.issue), due: day(i.due),
  paidOn: i.paid_on === null ? null : day(i.paid_on) }));
const bad = inv.filter((i) => {
  const s = invoiceState(i, today), e = expected.states[i.no];
  return Math.abs(s.balance - e.balance) > 0.004 || s.status !== e.status || s.days !== e.days || s.aging !== e.aging;
});
check(`Demo: state of all ${inv.length} invoices = Python`, bad.length === 0, bad.map((b) => b.no).join(','));

const customers = {}; DEMO.customers.forEach((c) => { customers[c.name] = { email: c.email, contact: c.contact, status: c.status }; });
const sent = {}; DEMO.log.forEach((e) => { (sent[e.no] = sent[e.no] || []).push(e.stage); });
const replies = DEMO.replies.map((r) => ({ date: day(r.day), no: r.no, type: r.type, promised: r.promised === null ? null : day(r.promised) }));
const pauses = pauseState(replies, today, DEMO.pauseDays);
const pz = (m) => JSON.stringify(Object.keys(m).sort().map((k) => [k, m[k].paused, m[k].broken, m[k].reason]));
check(`Demo: pause rule on ${replies.length} logged replies/promises = Python`, pz(pauses) === pz(expected.pauses), pz(pauses) + ' vs ' + pz(expected.pauses));

const plan = planReminders(inv, customers, sent, { stages: DEMO.stages, maxPerRun: DEMO.maxPerRun }, today, pauses);
const got = JSON.stringify(plan.map((p) => [p.no, p.stage, p.action, p.reason]));
const want = JSON.stringify(expected.plan.map((p) => [p.no, p.stage, p.action, p.reason]));
check(`Demo: reminder plan (${plan.length} lines, ${plan.filter((p) => p.action === 'pause').length} paused) = Python`, got === want, got + ' vs ' + want);
check('Demo: no invoice gets a stage it already had', plan.every((p) => !(sent[p.no] || []).includes(p.stage)));
check('Demo: nothing is sent for a paused invoice', plan.filter((p) => pauses[p.no] && pauses[p.no].paused).every((p) => p.action === 'pause'));

const sum = summarize(inv.map((i) => ({ no: i.no, customer: i.customer, state: invoiceState(i, today) })));
const es = expected.summary;
check('Demo: outstanding, overdue and aging = Python',
  near(sum.outstanding, es.outstanding) && near(sum.overdue, es.overdue) && sum.overdueCount === es.overdue_count &&
  Object.keys(es.aging).every((b) => near(sum.aging[b], es.aging[b])), JSON.stringify([sum.outstanding, sum.overdue, sum.aging]));

const pts = trendPoints(inv, today);
const ptBad = pts.filter((p, k) => {
  const e = expected.points[k];
  return p.label !== e.label || !near(p.outstanding, e.outstanding) || !near(p.overdue, e.overdue) || !near(p.collected, e.collected) ||
    !near(p.billed, e.billed) || !near(p.dso, e.dso, 0.001) || !near(p.rate, e.rate, 0.00001);
});
check('Demo: 6-month trend (outstanding, overdue, collected, billed, DSO, collection rate) = Python', ptBad.length === 0,
  JSON.stringify(ptBad));
const k = kpiCards(inv, today), ek = expected.kpis;
check(`Demo: KPI cards (collected ${k.collectedMtd} MTD vs ${k.collectedPrevMtd}, DSO ${k.dso.toFixed(1)}, rate ${(k.rate * 100).toFixed(1)}%) = Python`,
  near(k.collectedMtd, ek.collected_mtd) && near(k.collectedPrevMtd, ek.collected_prev_mtd) && near(k.dso, ek.dso, 0.001) &&
  near(k.dsoPrev, ek.dso_prev, 0.001) && near(k.rate, ek.rate, 0.00001) && near(k.ratePrev, ek.rate_prev, 0.00001) &&
  near(k.outstandingPrev, ek.outstanding_prev) && near(k.overduePrev, ek.overdue_prev), JSON.stringify(k));
check('Demo: DSO cross-check: outstanding / sales of the last 90 days x 90',
  near(k.dso, k.outstanding / inv.filter((i) => Math.round((today - i.issue) / 864e5) < 90 && i.issue <= today).reduce((a, i) => a + i.amount, 0) * 90, 0.001));

const risk = riskScores(inv, customers, today);
const rBad = Object.keys(expected.risk).filter((c) => {
  const r = risk[c], e = expected.risk[c];
  return !r || r.score !== e.score || r.level !== e.level || r.maxDays !== e.max_days || !near(r.overdue, e.overdue) || !near(r.avgLate, e.avg_late, 0.001);
});
check(`Demo: risk score of all ${Object.keys(expected.risk).length} customers = Python`, rBad.length === 0,
  rBad.map((c) => c + ' ' + JSON.stringify(risk[c])).join('; '));

const q = actionQueue(inv, customers, plan, pauses, today);
const qg = JSON.stringify(q.map((x) => [x.priority, x.action, x.no])), qw = JSON.stringify(expected.queue.map((x) => [x.priority, x.action, x.no]));
check(`Demo: owner action queue (${q.length} actions, ${q.filter((x) => x.priority === 'Urgent').length} urgent) = Python`, qg === qw, qg + ' vs ' + qw);
check('Demo: the broken promise is an Urgent call', q.some((x) => x.action === 'Call: promise not kept' && pauses[x.no] && pauses[x.no].broken));

console.log(`\n${pass}/${pass + fail} checks passed`);
process.exit(fail ? 1 : 0);
