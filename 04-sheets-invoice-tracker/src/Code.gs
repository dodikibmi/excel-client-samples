/**
 * Invoice & payment tracker - Google Sheets + Apps Script
 *
 * What it does
 *   - Invoices sheet: type invoices and payments; balance, status, days overdue and aging are ARRAYFORMULAs.
 *   - Payment reminders by email, in stages (Upcoming, Friendly, Second, Final), never the same stage twice,
 *     every email written to the Reminder log. Email mode on Settings: Preview only / Send to me (test) / Send to customers.
 *   - Weekly summary email to the owner (outstanding, overdue, aging, who to chase).
 *   - Daily automation with a time-driven trigger; a lock stops two runs at once.
 *   - Self-test: unit tests + checks that every formula on the sheet agrees with this script, written to the Checks sheet.
 *
 * Layout: pure logic first (no Google services, unit-tested in Node by test/run-tests.js and in Sheets by Run self-test),
 * then the sheet / mail layer, then the demo builder.
 */

const SHEET = { INVOICES: 'Invoices', CUSTOMERS: 'Customers', SETTINGS: 'Settings', DASHBOARD: 'Dashboard',
                LOG: 'Reminder log', CHECKS: 'Checks' };
const MODES = ['Preview only', 'Send to me (test)', 'Send to customers'];
const AGING = ['Not due', '1-30', '31-60', '61-90', '90+'];

// ============================================================================================ pure logic

/** Whole days since 1970 for a date, ignoring the time of day and daylight saving. */
function dayNumber(d) {
  return Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 864e5);
}

/** Days from a to b (positive when b is later). */
function daysBetween(a, b) {
  return dayNumber(b) - dayNumber(a);
}

function round2(x) {
  return Math.round((x + (x >= 0 ? 1 : -1) * 1e-9) * 100) / 100;
}

/** Balance, status, days overdue and aging bucket of one invoice {amount, paid, due} as of a date. */
function invoiceState(inv, asOf) {
  const balance = round2((Number(inv.amount) || 0) - (Number(inv.paid) || 0));
  if (balance < 0) return { balance: balance, status: 'Overpaid', days: 0, aging: 'Paid' };
  if (balance === 0) return { balance: 0, status: 'Paid', days: 0, aging: 'Paid' };
  const rel = daysBetween(inv.due, asOf);
  const days = Math.max(0, rel);
  const status = rel > 0 ? 'Overdue' : (Number(inv.paid) > 0 ? 'Part paid' : 'Not due');
  return { balance: balance, status: status, days: days, aging: agingBucket(days) };
}

function agingBucket(days) {
  if (days <= 0) return 'Not due';
  if (days <= 30) return '1-30';
  if (days <= 60) return '31-60';
  if (days <= 90) return '61-90';
  return '90+';
}

/**
 * Which reminder stage is due now, or null.
 * stages: [{name, day}] where day is days after the due date (negative = before it).
 * The highest stage reached is sent; nothing is sent if that stage or a later one was already sent,
 * so a customer never gets the same reminder twice and never gets an earlier, softer one after a firmer one.
 */
function reminderStage(daysAfterDue, stages, sentStages) {
  const sorted = stages.slice().sort(function (a, b) { return a.day - b.day; });
  const reached = sorted.filter(function (s) { return daysAfterDue >= s.day; });
  if (!reached.length) return null;
  const target = reached[reached.length - 1];
  const already = sorted.some(function (s) { return s.day >= target.day && sentStages.indexOf(s.name) >= 0; });
  return already ? null : target.name;
}

/**
 * The reminder plan for a run. invoices: [{no, customer, amount, paid, due}], customers: {name: {email, contact}},
 * sent: {invoiceNo: [stage names already sent]}. Returns [{no, customer, email, stage, action, reason, ...}]
 * most overdue first; action is send, skip (no email) or hold (over the per-run limit).
 */
function planReminders(invoices, customers, sent, settings, asOf) {
  const rows = invoices.slice().sort(function (a, b) {
    return dayNumber(a.due) - dayNumber(b.due) || String(a.no).localeCompare(String(b.no));
  });
  const plan = [];
  let sends = 0;
  rows.forEach(function (inv) {
    const st = invoiceState(inv, asOf);
    if (st.balance <= 0) return;
    const stage = reminderStage(daysBetween(inv.due, asOf), settings.stages, sent[inv.no] || []);
    if (!stage) return;
    const cust = customers[inv.customer] || {};
    const item = { no: inv.no, customer: inv.customer, contact: cust.contact || '', email: cust.email || '', stage: stage,
                   balance: st.balance, days: st.days, due: inv.due, action: 'send', reason: '' };
    if (!item.email) {
      item.action = 'skip'; item.reason = 'No email for this customer';
    } else if (++sends > settings.maxPerRun) {
      item.action = 'hold'; item.reason = 'Run limit reached';
    }
    plan.push(item);
  });
  return plan;
}

/** Replaces {placeholders} in a template. Unknown placeholders are left as they are, so a typo is visible. */
function fillTemplate(text, values) {
  return String(text).replace(/\{(\w+)\}/g, function (m, key) {
    return Object.prototype.hasOwnProperty.call(values, key) ? String(values[key]) : m;
  });
}

function money(x, currency) {
  const s = Math.abs(x).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (x < 0 ? '-' : '') + (currency || '') + s;
}

function fmtDate(d) {
  const m = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return d.getDate() + ' ' + m[d.getMonth()] + ' ' + d.getFullYear();
}

/** Subject and body of one reminder, from the stage's templates on Settings. */
function buildReminder(item, settings) {
  const t = settings.templates[item.stage];
  const values = { contact: item.contact || item.customer, customer: item.customer, invoice: item.no,
                   amount: money(item.balance, settings.currency), due: fmtDate(item.due), days: item.days,
                   company: settings.company, sender: settings.senderName };
  return { subject: fillTemplate(t.subject, values), body: fillTemplate(t.body, values) };
}

/** Totals for the summary email and the self-test. states: [{no, customer, state}] */
function summarize(states) {
  const s = { outstanding: 0, overdue: 0, overdueCount: 0, aging: {}, owing: {} };
  AGING.forEach(function (b) { s.aging[b] = 0; });
  states.forEach(function (r) {
    const st = r.state;
    if (st.balance <= 0) return;
    s.outstanding += st.balance;
    s.aging[st.aging] += st.balance;
    s.owing[r.customer] = (s.owing[r.customer] || 0) + st.balance;
    if (st.status === 'Overdue') { s.overdue += st.balance; s.overdueCount++; }
  });
  s.outstanding = round2(s.outstanding); s.overdue = round2(s.overdue);
  AGING.forEach(function (b) { s.aging[b] = round2(s.aging[b]); });
  s.top = Object.keys(s.owing).map(function (k) { return { customer: k, owes: round2(s.owing[k]) }; })
    .sort(function (a, b) { return b.owes - a.owes; }).slice(0, 5);
  return s;
}

function summaryHtml(sum, overdueRows, sentThisWeek, settings, asOf) {
  const td = 'style="padding:4px 10px;border-bottom:1px solid #e5e5e5"';
  const tdr = 'style="padding:4px 10px;border-bottom:1px solid #e5e5e5;text-align:right"';
  const c = settings.currency;
  let h = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1b2d48">';
  h += '<h2 style="margin:0 0 4px">' + settings.company + ': money owed, ' + fmtDate(asOf) + '</h2>';
  h += '<p style="margin:0 0 14px;color:#666">Outstanding <b>' + money(sum.outstanding, c) + '</b> &middot; overdue <b style="color:#b00020">' +
       money(sum.overdue, c) + '</b> on ' + sum.overdueCount + ' invoices &middot; reminders sent in the last 7 days: ' + sentThisWeek + '</p>';
  h += '<table style="border-collapse:collapse"><tr><th align="left" ' + td + '>Age</th><th ' + tdr + '>Owed</th></tr>';
  AGING.forEach(function (b) { h += '<tr><td ' + td + '>' + b + (b === 'Not due' ? '' : ' days') + '</td><td ' + tdr + '>' + money(sum.aging[b], c) + '</td></tr>'; });
  h += '</table><h3 style="margin:16px 0 4px">Chase first</h3><table style="border-collapse:collapse">';
  h += '<tr><th align="left" ' + td + '>Invoice</th><th align="left" ' + td + '>Customer</th><th ' + tdr + '>Days late</th><th ' + tdr + '>Balance</th></tr>';
  overdueRows.slice(0, 10).forEach(function (r) {
    h += '<tr><td ' + td + '>' + r.no + '</td><td ' + td + '>' + r.customer + '</td><td ' + tdr + '>' + r.state.days + '</td><td ' + tdr + '>' + money(r.state.balance, c) + '</td></tr>';
  });
  return h + '</table><p style="color:#888;font-size:12px">Sent by the invoice tracker in your Google Sheet.</p></div>';
}

// ============================================================================================ sheet layer

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invoice tracker')
    .addItem('Preview reminders', 'menuPreview')
    .addItem('Send reminders now', 'menuSend')
    .addItem('Email me the summary', 'menuSummary')
    .addSeparator()
    .addItem('Turn on daily automation', 'installTriggers')
    .addItem('Turn off automation', 'removeTriggers')
    .addSeparator()
    .addItem('Run self-test', 'menuSelfTest')
    .addItem('Rebuild demo data', 'menuRebuildDemo')
    .addToUi();
}

function ss_() { return SpreadsheetApp.getActive(); }

function readSettings_() {
  const v = {};
  ss_().getSheetByName(SHEET.SETTINGS).getRange('A2:B12').getValues().forEach(function (r) { if (r[0]) v[r[0]] = r[1]; });
  const stages = [], templates = {};
  ss_().getSheetByName(SHEET.SETTINGS).getRange('A16:D19').getValues().forEach(function (r) {
    if (!r[0]) return;
    stages.push({ name: String(r[0]), day: Number(r[1]) });
    templates[r[0]] = { subject: r[2], body: r[3] };
  });
  return {
    company: v['Company'], senderName: v['Sender name'], ownerEmail: v['Owner email'] || Session.getEffectiveUser().getEmail(),
    currency: v['Currency symbol'] || '', mode: MODES.indexOf(v['Email mode']) >= 0 ? v['Email mode'] : 'Preview only',
    maxPerRun: Number(v['Max emails per run']) || 25, summaryDay: v['Summary day'] || 'Monday',
    asOf: v['As-of used'] instanceof Date ? v['As-of used'] : new Date(),  // the same date the formulas use
    stages: stages, templates: templates,
  };
}

function readInvoices_() {
  const sh = ss_().getSheetByName(SHEET.INVOICES);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 13).getValues().map(function (r, i) {
    return { row: i + 2, no: r[0], customer: r[1], issue: r[2], due: r[3], amount: r[4], paid: r[5], paidOn: r[6],
             sheet: { balance: r[8], status: r[9], days: r[10], aging: r[11] } };
  }).filter(function (r) { return r.no !== '' && r.due instanceof Date; });
}

function readCustomers_() {
  const sh = ss_().getSheetByName(SHEET.CUSTOMERS);
  const out = {};
  if (sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
    if (r[0]) out[r[0]] = { contact: r[1], email: String(r[2]).trim() };
  });
  return out;
}

/** {invoiceNo: [stages]} from the log, counting only emails that really went to the customer. */
function readSent_() {
  const sh = ss_().getSheetByName(SHEET.LOG);
  const sent = {};
  if (sh.getLastRow() < 2) return sent;
  sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues().forEach(function (r) {
    if (r[6] === 'Sent to customer') (sent[r[1]] = sent[r[1]] || []).push(r[4]);
  });
  return sent;
}

/** Plans and (depending on Email mode) sends reminders. Returns the plan with what happened to each line. */
function runReminders_(forceMode) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another run is still busy. Try again in a minute.');
  try {
    const s = readSettings_();
    const mode = forceMode || s.mode;
    const plan = planReminders(readInvoices_(), readCustomers_(), readSent_(), s, s.asOf);
    const log = ss_().getSheetByName(SHEET.LOG);
    const now = new Date();
    const rows = [];
    plan.forEach(function (p) {
      const mail = buildReminder(p, s);
      p.subject = mail.subject; p.body = mail.body;
      if (mode === 'Preview only') { p.result = p.action === 'send' ? 'Preview' : p.reason; return; }
      if (p.action !== 'send') { p.result = p.reason; rows.push([now, p.no, p.customer, p.email, p.stage, mode, p.reason]); return; }
      if (MailApp.getRemainingDailyQuota() < 1) { p.result = 'Daily email quota used up'; rows.push([now, p.no, p.customer, p.email, p.stage, mode, p.result]); return; }
      const to = mode === 'Send to customers' ? p.email : s.ownerEmail;
      MailApp.sendEmail({ to: to, subject: (mode === 'Send to customers' ? '' : '[TEST to ' + p.email + '] ') + mail.subject,
                          body: mail.body, name: s.senderName, replyTo: s.ownerEmail });
      p.result = mode === 'Send to customers' ? 'Sent to customer' : 'Sent to me (test)';
      rows.push([now, p.no, p.customer, to, p.stage, mode, p.result]);
      if (mode === 'Send to customers') {
        ss_().getSheetByName(SHEET.INVOICES).getRange(rowOf_(p.no), 13).setValue(p.stage + ' ' + fmtDate(now));
      }
    });
    if (rows.length) log.getRange(log.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    return { mode: mode, plan: plan };
  } finally {
    lock.releaseLock();
  }
}

function rowOf_(no) {
  const inv = readInvoices_().filter(function (r) { return r.no === no; })[0];
  return inv ? inv.row : null;
}

function showPlan_(title, res) {
  const esc = function (x) { return String(x).replace(/&/g, '&amp;').replace(/</g, '&lt;'); };
  let h = '<div style="font-family:Arial;font-size:13px"><p>Email mode: <b>' + esc(res.mode) + '</b> &middot; ' + res.plan.length + ' reminder(s)</p>';
  h += '<table style="border-collapse:collapse;width:100%">';
  res.plan.forEach(function (p) {
    h += '<tr style="border-top:1px solid #ddd"><td style="padding:4px;vertical-align:top"><b>' + esc(p.no) + '</b><br>' + esc(p.stage) +
         '<br><span style="color:' + (p.action === 'send' ? '#00796b' : '#b00020') + '">' + esc(p.result || p.reason || p.action) + '</span></td>' +
         '<td style="padding:4px"><b>To:</b> ' + esc(p.email || '(none)') + '<br><b>' + esc(p.subject) + '</b><pre style="white-space:pre-wrap;font-family:Arial">' +
         esc(p.body) + '</pre></td></tr>';
  });
  if (!res.plan.length) h += '<tr><td>Nothing to send today.</td></tr>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(h + '</table></div>').setWidth(760).setHeight(560), title);
}

function menuPreview() { showPlan_('Reminders that would go out today', runReminders_('Preview only')); }

function menuSend() {
  const s = readSettings_();
  const ui = SpreadsheetApp.getUi();
  if (s.mode === 'Preview only') {
    ui.alert('Email mode on Settings is "Preview only", so nothing is sent. Change it to "Send to me (test)" or "Send to customers".');
    return;
  }
  if (ui.alert('Send reminders now (' + s.mode + ')?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  showPlan_('Reminders: what happened', runReminders_());
}

function sendSummary_() {
  const s = readSettings_();
  const inv = readInvoices_();
  const states = inv.map(function (r) { return { no: r.no, customer: r.customer, state: invoiceState(r, s.asOf) }; });
  const sum = summarize(states);
  const overdue = states.filter(function (r) { return r.state.status === 'Overdue'; })
    .sort(function (a, b) { return b.state.days - a.state.days; });
  const log = ss_().getSheetByName(SHEET.LOG);
  const week = log.getLastRow() < 2 ? 0 : log.getRange(2, 1, log.getLastRow() - 1, 7).getValues().filter(function (r) {
    return r[0] instanceof Date && daysBetween(r[0], s.asOf) <= 7 && String(r[6]).indexOf('Sent') === 0;
  }).length;
  MailApp.sendEmail({ to: s.ownerEmail, subject: s.company + ': ' + money(sum.overdue, s.currency) + ' overdue on ' + sum.overdueCount + ' invoices',
                      htmlBody: summaryHtml(sum, overdue, week, s, s.asOf), name: 'Invoice tracker' });
  return s.ownerEmail;
}

function menuSummary() { SpreadsheetApp.getUi().alert('Summary sent to ' + sendSummary_()); }

/** Runs every morning: reminders (in the Email mode on Settings), and the summary on the chosen weekday. */
function dailyRun() {
  runReminders_();
  const s = readSettings_();
  const day = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][new Date().getDay()];
  if (day === s.summaryDay) sendSummary_();
}

function installTriggers() {
  removeTriggers(true);
  ScriptApp.newTrigger('dailyRun').timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getUi().alert('Daily automation is on: reminders every morning around 8:00 (' +
    Session.getScriptTimeZone() + '), summary every ' + readSettings_().summaryDay + '.');
}

function removeTriggers(quiet) {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyRun') ScriptApp.deleteTrigger(t); });
  if (quiet !== true) SpreadsheetApp.getUi().alert('Daily automation is off.');
}

// ============================================================================================ self-test

/** Unit tests (Tests.gs) + every formula row on Invoices compared with invoiceState() + dashboard totals. */
function selfTest() {
  const results = runAllTests().map(function (t) { return ['Unit test', t.name, t.pass ? 'PASS' : 'FAIL', t.detail || '']; });
  const s = readSettings_();
  const inv = readInvoices_();
  let bad = [];
  inv.forEach(function (r) {
    const st = invoiceState(r, s.asOf);
    const ok = Math.abs(st.balance - r.sheet.balance) < 0.005 && st.status === r.sheet.status && st.days === r.sheet.days && st.aging === r.sheet.aging;
    if (!ok) bad.push(r.no + ' sheet ' + [r.sheet.status, r.sheet.days, r.sheet.aging].join('/') + ' script ' + [st.status, st.days, st.aging].join('/'));
  });
  results.push(['Sheet formulas', 'Balance, status, days and aging on all ' + inv.length + ' invoices = script', bad.length ? 'FAIL' : 'PASS', bad.slice(0, 3).join('; ')]);
  const sum = summarize(inv.map(function (r) { return { no: r.no, customer: r.customer, state: invoiceState(r, s.asOf) }; }));
  const d = ss_().getRangeByName('DashOutstanding').getValue(), o = ss_().getRangeByName('DashOverdue').getValue();
  results.push(['Dashboard', 'Outstanding on Dashboard = script', Math.abs(d - sum.outstanding) < 0.005 ? 'PASS' : 'FAIL', d + ' vs ' + sum.outstanding]);
  results.push(['Dashboard', 'Overdue on Dashboard = script', Math.abs(o - sum.overdue) < 0.005 ? 'PASS' : 'FAIL', o + ' vs ' + sum.overdue]);
  const aging = ss_().getRangeByName('DashAging').getValues();
  const agingOk = aging.every(function (r) { return Math.abs(r[1] - sum.aging[r[0]]) < 0.005; });
  results.push(['Dashboard', 'Aging table on Dashboard = script', agingOk ? 'PASS' : 'FAIL', JSON.stringify(aging)]);
  const per = {};
  inv.forEach(function (r) {
    const st = invoiceState(r, s.asOf), p = per[r.customer] = per[r.customer] || [0, 0];
    if (st.balance > 0) p[0] += st.balance;
    if (st.status === 'Overdue') p[1] += st.balance;
  });
  const cs = ss_().getSheetByName(SHEET.CUSTOMERS).getDataRange().getValues().slice(1).filter(function (r) { return r[0] !== ''; });
  const custBad = cs.filter(function (r) { const p = per[r[0]] || [0, 0]; return Math.abs(r[3] - p[0]) > 0.005 || Math.abs(r[4] - p[1]) > 0.005; });
  results.push(['Customers', 'Open balance and overdue for all ' + cs.length + ' customers = script', custBad.length ? 'FAIL' : 'PASS',
    custBad.slice(0, 3).map(function (r) { return r[0] + ' ' + r[3] + '/' + r[4]; }).join('; ')]);
  const sh = ss_().getSheetByName(SHEET.CHECKS);
  sh.getRange('A4:D200').clearContent();
  sh.getRange(4, 1, results.length, 4).setValues(results);
  sh.getRange('B2').setValue(new Date());
  const failed = results.filter(function (r) { return r[2] !== 'PASS'; }).length;
  sh.getRange('C2').setValue(failed ? failed + ' FAILED' : 'ALL ' + results.length + ' PASS');
  return { total: results.length, failed: failed };
}

function menuSelfTest() {
  const r = selfTest();
  ss_().getSheetByName(SHEET.CHECKS).activate();
  SpreadsheetApp.getUi().alert(r.failed ? r.failed + ' of ' + r.total + ' checks FAILED - see the Checks sheet.' : 'All ' + r.total + ' checks pass.');
}
