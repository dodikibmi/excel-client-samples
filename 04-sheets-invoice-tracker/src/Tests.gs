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
  return out;
}
