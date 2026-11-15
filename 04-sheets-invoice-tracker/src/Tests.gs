/**
 * Unit tests for the pure logic in Code.gs. They run in two places with the same code:
 * in Google Sheets (Invoice tracker > Run self-test, results on the Checks sheet) and in Node (test/run-tests.js).
 */
function runAllTests() {
  const out = [];
  const t = function (name, actual, expected) {
    const a = JSON.stringify(actual), e = JSON.stringify(expected);
    out.push({ name: name, pass: a === e, detail: a === e ? '' : 'got ' + a + ', expected ' + e });
  };
  const asOf = new Date(2026, 10, 14);            // 14 Nov 2026
  const d = function (y, m, day) { return new Date(y, m - 1, day); };
  const st = function (amount, paid, due) { return invoiceState({ amount: amount, paid: paid, due: due }, asOf); };

  // invoice state
  t('Paid in full', st(500, 500, d(2026, 10, 1)), { balance: 0, status: 'Paid', days: 0, aging: 'Paid' });
  t('Paid with float noise (0.1 + 0.2) is still paid', st(0.3, 0.1 + 0.2, d(2026, 10, 1)).status, 'Paid');
  t('Overpaid is flagged, not hidden', st(820, 850, d(2026, 10, 1)), { balance: -30, status: 'Overpaid', days: 0, aging: 'Paid' });
  t('Due today is not overdue yet', st(100, 0, d(2026, 11, 14)), { balance: 100, status: 'Not due', days: 0, aging: 'Not due' });
  t('One day late', st(100, 0, d(2026, 11, 13)), { balance: 100, status: 'Overdue', days: 1, aging: '1-30' });
  t('30 days late is still 1-30', st(100, 0, d(2026, 10, 15)).aging, '1-30');
  t('31 days late moves to 31-60', st(100, 0, d(2026, 10, 14)).aging, '31-60');
  t('90 days late is 61-90, 91 is 90+', [st(1, 0, d(2026, 8, 16)).aging, st(1, 0, d(2026, 8, 15)).aging], ['61-90', '90+']);
  t('Part paid, not due', st(1000, 400, d(2026, 11, 30)), { balance: 600, status: 'Part paid', days: 0, aging: 'Not due' });
  t('Part paid and late counts as overdue', st(1000, 400, d(2026, 11, 1)).status, 'Overdue');
  t('Days are calendar days across a month end', daysBetween(d(2026, 10, 31), d(2026, 11, 1)), 1);

  // reminder stage
  const S = [{ name: 'Upcoming', day: -3 }, { name: 'Friendly', day: 1 }, { name: 'Second', day: 14 }, { name: 'Final', day: 30 }];
  t('Nothing 4 days before due', reminderStage(-4, S, []), null);
  t('Heads-up 3 days before due', reminderStage(-3, S, []), 'Upcoming');
  t('Friendly the day after due', reminderStage(1, S, ['Upcoming']), 'Friendly');
  t('No second Friendly', reminderStage(10, S, ['Friendly']), null);
  t('Second at exactly 14 days', reminderStage(14, S, ['Friendly']), 'Second');
  t('Jumps straight to Final when nothing was sent', reminderStage(45, S, []), 'Final');
  t('Never a softer email after a firmer one', reminderStage(20, S, ['Final']), null);
  t('Stage order on Settings does not matter', reminderStage(15, S.slice().reverse(), []), 'Second');

  // plan
  const cust = { A: { email: 'a@example.com', contact: 'Ann' }, B: { email: '', contact: 'Bob' } };
  const settings = { stages: S, maxPerRun: 2 };
  const inv = [
    { no: 'X1', customer: 'A', amount: 100, paid: 0, due: d(2026, 9, 1) },    // 74 days: Final
    { no: 'X2', customer: 'B', amount: 100, paid: 0, due: d(2026, 10, 1) },   // 44 days, no email: skip
    { no: 'X3', customer: 'A', amount: 100, paid: 100, due: d(2026, 10, 1) }, // paid: nothing
    { no: 'X4', customer: 'A', amount: 100, paid: 0, due: d(2026, 11, 10) },  // Friendly
    { no: 'X5', customer: 'A', amount: 100, paid: 0, due: d(2026, 11, 12) },  // Friendly, but over the limit of 2
    { no: 'X6', customer: 'A', amount: 100, paid: 0, due: d(2026, 11, 13) },  // Friendly already sent
    { no: 'X7', customer: 'Z', amount: 100, paid: 0, due: d(2026, 11, 16) },  // customer not on the list: skip
  ];
  const plan = planReminders(inv, cust, { X6: ['Friendly'] }, settings, asOf);
  t('Plan: most overdue first, with the right action', plan.map(function (p) { return [p.no, p.stage, p.action]; }),
    [['X1', 'Final', 'send'], ['X2', 'Final', 'skip'], ['X4', 'Friendly', 'send'], ['X5', 'Friendly', 'hold'], ['X7', 'Upcoming', 'skip']]);
  t('Plan: skipped lines say why', plan[1].reason, 'No email for this customer');

  // email text
  t('Template fills placeholders and keeps unknown ones visible',
    fillTemplate('Hi {contact}, {invoice} {oops}', { contact: 'Ann', invoice: 'NB-1' }), 'Hi Ann, NB-1 {oops}');
  t('Money format', [money(1234567.5, '$'), money(-30, '$'), money(0.005, '')], ['$1,234,567.50', '-$30.00', '0.01']);
  const mail = buildReminder({ no: 'NB-9', customer: 'A Co', contact: 'Ann', stage: 'Second', balance: 2100, days: 14, due: d(2026, 10, 31) },
    { currency: '$', company: 'Northbeam', senderName: 'Accounts',
      templates: { Second: { subject: 'Invoice {invoice} is {days} days overdue', body: 'Hi {contact}: {amount} due {due}. {sender}, {company}' } } });
  t('Reminder email text', mail, { subject: 'Invoice NB-9 is 14 days overdue', body: 'Hi Ann: $2,100.00 due 31 Oct 2026. Accounts, Northbeam' });

  // summary
  const sum = summarize([{ customer: 'A', state: st(100, 0, d(2026, 11, 1)) }, { customer: 'B', state: st(50, 0, d(2026, 12, 1)) },
                         { customer: 'A', state: st(70, 70, d(2026, 9, 1)) }, { customer: 'B', state: st(10, 20, d(2026, 9, 1)) }]);
  t('Summary totals ignore paid and overpaid lines', [sum.outstanding, sum.overdue, sum.overdueCount, sum.aging['1-30'], sum.aging['Not due']],
    [150, 100, 1, 100, 50]);
  t('Summary top customers', sum.top, [{ customer: 'A', owes: 100 }, { customer: 'B', owes: 50 }]);

  // pause rule: replies and promises to pay
  const R = function (date, no, type, promised) { return { date: date, no: no, type: type, promised: promised || null }; };
  const pz = pauseState([R(d(2026, 11, 13), 'X4', 'Reply'), R(d(2026, 11, 1), 'X1', 'Reply'),
    R(d(2026, 11, 10), 'X5', 'Promise to pay', d(2026, 11, 20)), R(d(2026, 11, 2), 'X6', 'Promise to pay', d(2026, 11, 9)),
    R(d(2026, 11, 7), 'X7', 'Reply')], asOf, 7);
  t('Pause: a reply yesterday pauses reminders', [pz.X4.paused, pz.X4.reason], [true, 'Customer replied on 13 Nov 2026']);
  t('Pause: a reply 13 days ago no longer pauses', pz.X1.paused, false);
  t('Pause: a reply exactly 7 days ago still pauses', pz.X7.paused, true);
  t('Pause: promise for a future date pauses until then', [pz.X5.paused, pz.X5.broken], [true, false]);
  t('Pause: a promise whose date passed is broken, not paused', [pz.X6.paused, pz.X6.broken, pz.X6.reason],
    [false, true, 'Promise to pay by 9 Nov 2026 not kept']);
  t('Pause: only the latest entry counts', pauseState([R(d(2026, 11, 1), 'Y', 'Promise to pay', d(2026, 11, 5)),
    R(d(2026, 11, 12), 'Y', 'Reply')], asOf, 7).Y, { paused: true, broken: false, type: 'Reply', promised: null, reason: 'Customer replied on 12 Nov 2026' });
  const plan2 = planReminders(inv, cust, { X6: ['Friendly'] }, { stages: S, maxPerRun: 2 }, asOf, { X1: { paused: true, reason: 'held' } });
  t('Pause: a paused invoice is not emailed and does not use the run limit', plan2.map(function (p) { return [p.no, p.action]; }),
    [['X1', 'pause'], ['X2', 'skip'], ['X4', 'send'], ['X5', 'send'], ['X7', 'skip']]);
  const cust2 = { A: { email: 'a@example.com', contact: 'Ann', status: 'On hold' } };
  t('Pause: customer status On hold stops reminders', planReminders([inv[0]], cust2, {}, settings, asOf)[0].reason, 'Customer status: On hold');

  // DSO, collection rate, trends
  const I = function (no, issue, due, amount, paid, paidOn) { return { no: no, customer: 'A', issue: issue, due: due, amount: amount, paid: paid, paidOn: paidOn || null }; };
  const book = [I('D1', d(2026, 9, 1), d(2026, 10, 1), 1000, 1000, d(2026, 10, 5)), I('D2', d(2026, 10, 1), d(2026, 10, 31), 2000, 500, d(2026, 11, 2)),
                I('D3', d(2026, 11, 1), d(2026, 12, 1), 3000, 0), I('D4', d(2026, 5, 1), d(2026, 5, 31), 900, 900, d(2026, 6, 10))];
  const m = metricsAt(book, asOf, asOf);
  t('DSO = outstanding / sales of the last 90 days x 90', [m.outstanding, Math.round(m.dso * 100) / 100], [4500, 67.5]);
  t('Collection rate = paid share of what fell due in the last 90 days', Math.round(m.rate * 10000) / 10000, 0.5);
  t('Outstanding as it stood on an earlier day', metricsAt(book, d(2026, 10, 31), asOf).outstanding, 2000);
  t('Overdue as on an earlier day counts only invoices past due then', metricsAt(book, d(2026, 11, 1), asOf).overdue, 2000);
  const tp = trendPoints(book, asOf);
  t('Trend: 6 months ending this month, cash collected per month', tp.map(function (p) { return [p.label, p.collected]; }),
    [['Jun', 900], ['Jul', 0], ['Aug', 0], ['Sep', 0], ['Oct', 1000], ['Nov', 500]]);
  t('KPI: collected month to date vs the same days last month', [kpiCards(book, asOf).collectedMtd, kpiCards(book, asOf).collectedPrevMtd], [500, 1000]);
  t('KPI: a payment typed without a date counts today', collectedBetween([I('E', d(2026, 11, 1), d(2026, 11, 2), 50, 50)], d(2026, 11, 14), d(2026, 11, 14), asOf), 50);

  // risk score
  const rk = riskScores([I('R1', d(2026, 8, 1), d(2026, 8, 31), 1000, 1000, d(2026, 10, 10)),     // paid 40 days late
                         I('R2', d(2026, 9, 1), d(2026, 10, 1), 5000, 0)], { A: { email: '' } }, asOf); // 44 days overdue
  t('Risk: habit 40x0.5 + late 44x0.3 + overdue 5000/500 + no email 10', [rk.A.score, rk.A.level], [Math.round(20 + 13.2 + 10 + 10), 'Medium']);
  const rk2 = riskScores([I('R3', d(2026, 1, 1), d(2026, 1, 31), 99999, 0)], { A: { email: 'a@x' } }, asOf);
  t('Risk: each part is capped (max 40 + 30 + 20)', [rk2.A.score, rk2.A.level], [50, 'Medium']);
  t('Risk: a customer who always pays on time scores 0', riskScores([I('R4', d(2026, 10, 1), d(2026, 10, 31), 100, 100, d(2026, 10, 20))],
    { A: { email: 'a@x' } }, asOf).A, { avgLate: 0, maxDays: 0, overdue: 0, score: 0, level: 'Low' });

  // action queue
  const qInv = [I('Q1', d(2026, 7, 1), d(2026, 8, 1), 500, 0), I('Q2', d(2026, 10, 1), d(2026, 10, 10), 900, 0),
                I('Q3', d(2026, 10, 1), d(2026, 11, 10), 300, 0), I('Q4', d(2026, 10, 1), d(2026, 10, 20), 800, 0),
                I('Q5', d(2026, 10, 1), d(2026, 11, 16), 200, 0)];
  const qPlan = planReminders(qInv, { A: { email: 'a@x', contact: 'Ann' } }, {}, { stages: S, maxPerRun: 25 }, asOf,
    { Q4: { paused: false, broken: true, reason: 'not kept' }, Q3: { paused: true, type: 'Reply', reason: 'replied' } });
  const qq = actionQueue(qInv, { A: { email: 'a@x' } }, qPlan, { Q4: { paused: false, broken: true, reason: 'not kept' },
    Q3: { paused: true, type: 'Reply', reason: 'replied' } }, asOf);
  t('Queue: ranked Urgent > High > Medium > Low, bigger balance first', qq.map(function (x) { return [x.no, x.priority, x.action]; }),
    [['Q4', 'Urgent', 'Call: promise not kept'], ['Q1', 'Urgent', 'Call: over 60 days late'], ['Q2', 'High', 'Send Final reminder'],
     ['Q3', 'Medium', 'Record a promise to pay'], ['Q5', 'Low', 'Send Upcoming reminder']]);
  t('Queue: a promise due in 2 days becomes a Low check, later ones need nothing',
    [actionQueue([qInv[1]], { A: {} }, [], { Q2: { paused: true, type: 'Promise to pay', promised: d(2026, 11, 16) } }, asOf).length,
     actionQueue([qInv[1]], { A: {} }, [], { Q2: { paused: true, type: 'Promise to pay', promised: d(2026, 11, 17) } }, asOf).length], [1, 0]);
  return out;
}
