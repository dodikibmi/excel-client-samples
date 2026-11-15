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
                LOG: 'Reminder log', CHECKS: 'Checks', REPLIES: 'Replies', STATEMENT: 'Statement', CALC: 'Calc' };
const STATUSES = ['Active', 'On hold', 'Collections'];   // On hold / Collections: no automatic reminders
const REPLY_TYPES = ['Reply', 'Promise to pay'];
const PRIORITY = ['Urgent', 'High', 'Medium', 'Low'];
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
 * most overdue first; action is send, skip (no email), hold (over the per-run limit) or pause (the customer replied or
 * promised to pay - see pauseState - or the customer's status is not Active). pauses is optional.
 */
function planReminders(invoices, customers, sent, settings, asOf, pauses) {
  pauses = pauses || {};
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
    const status = cust.status || 'Active', pz = pauses[inv.no];
    if (status !== 'Active') {
      item.action = 'pause'; item.reason = 'Customer status: ' + status;
    } else if (pz && pz.paused) {
      item.action = 'pause'; item.reason = pz.reason;
    } else if (!item.email) {
      item.action = 'skip'; item.reason = 'No email for this customer';
    } else if (++sends > settings.maxPerRun) {
      item.action = 'hold'; item.reason = 'Run limit reached';
    }
    plan.push(item);
  });
  return plan;
}

/**
 * Pause rule. replies: [{date, no, type: 'Reply' | 'Promise to pay', promised}] as logged on the Replies sheet.
 * Only the latest entry per invoice counts (on the same day, the lower row wins):
 *   Promise to pay, promised date today or later -> reminders paused until then.
 *   Promise to pay, promised date passed and the invoice still open -> not paused, flagged broken (call the customer).
 *   Reply within pauseDays -> paused, so nobody gets a reminder the day after they answered; older replies do nothing.
 */
function pauseState(replies, asOf, pauseDays) {
  const latest = {};
  replies.forEach(function (r, i) {
    if (!r.no || !(r.date instanceof Date)) return;
    const cur = latest[r.no];
    if (!cur || dayNumber(r.date) >= dayNumber(cur.date)) latest[r.no] = r;
  });
  const out = {};
  Object.keys(latest).forEach(function (k) {
    const r = latest[k];
    if (r.type === 'Promise to pay' && r.promised instanceof Date) {
      const ok = daysBetween(asOf, r.promised) >= 0;
      out[k] = { paused: ok, broken: !ok, type: r.type, promised: r.promised,
                 reason: ok ? 'Promised to pay by ' + fmtDate(r.promised) : 'Promise to pay by ' + fmtDate(r.promised) + ' not kept' };
    } else {
      const paused = daysBetween(r.date, asOf) <= pauseDays;
      out[k] = { paused: paused, broken: false, type: 'Reply', promised: null, reason: paused ? 'Customer replied on ' + fmtDate(r.date) : '' };
    }
  });
  return out;
}

/** Day the money came in: the paid date, or the as-of date when a payment was typed without one. */
function payDate_(inv, asOf) { return inv.paidOn instanceof Date ? inv.paidOn : asOf; }

/**
 * Receivables as on day p (a Date): outstanding and overdue as they stood that day, DSO (outstanding / sales of the
 * 90 days to p x 90) and collection rate (share of the amount falling due in the 90 days to p that was paid by p).
 */
function metricsAt(invoices, p, asOf) {
  const P = dayNumber(p);
  let out = 0, ovd = 0, sales = 0, dueAmt = 0, dueColl = 0;
  invoices.forEach(function (i) {
    const amt = Number(i.amount) || 0, paid = Number(i.paid) || 0;
    const paidBy = paid && dayNumber(payDate_(i, asOf)) <= P ? paid : 0;
    const bal = amt - paidBy, iss = dayNumber(i.issue), due = dayNumber(i.due);
    if (iss <= P && bal > 1e-7) { out += bal; if (due < P) ovd += bal; }
    if (iss > P - 90 && iss <= P) sales += amt;
    if (due > P - 90 && due <= P) { dueAmt += amt; dueColl += Math.min(paidBy, amt); }
  });
  return { outstanding: round2(out), overdue: round2(ovd), dso: sales ? out / sales * 90 : 0, rate: dueAmt ? dueColl / dueAmt : 0 };
}

/** Cash received from day a to day b (Dates, inclusive). */
function collectedBetween(invoices, a, b, asOf) {
  const A = dayNumber(a), B = dayNumber(b);
  let s = 0;
  invoices.forEach(function (i) {
    const paid = Number(i.paid) || 0, d = dayNumber(payDate_(i, asOf));
    if (paid && d >= A && d <= B) s += paid;
  });
  return round2(s);
}

/** Six points for the trends: the last day of each of the 5 previous months, then the as-of date. */
function trendPoints(invoices, asOf) {
  const mon = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const pts = [];
  for (let k = 5; k >= 0; k--) {
    const start = new Date(asOf.getFullYear(), asOf.getMonth() - k, 1);
    const end = k ? new Date(asOf.getFullYear(), asOf.getMonth() - k + 1, 0) : asOf;
    let billed = 0;
    invoices.forEach(function (i) {
      const d = dayNumber(i.issue);
      if (d >= dayNumber(start) && d <= dayNumber(end)) billed += Number(i.amount) || 0;
    });
    const m = metricsAt(invoices, end, asOf);
    pts.push({ label: mon[start.getMonth()], start: start, end: end, collected: collectedBetween(invoices, start, end, asOf),
               billed: round2(billed), outstanding: m.outstanding, overdue: m.overdue, dso: m.dso, rate: m.rate });
  }
  return pts;
}

/** The KPI cards: now against the end of last month (collected: month to date against the same days last month). */
function kpiCards(invoices, asOf) {
  const pts = trendPoints(invoices, asOf), now = pts[5], prev = pts[4];
  const prevEnd = new Date(asOf.getFullYear(), asOf.getMonth(), 0);
  const sameDay = new Date(prevEnd.getFullYear(), prevEnd.getMonth(), Math.min(asOf.getDate(), prevEnd.getDate()));
  return { outstanding: now.outstanding, outstandingPrev: prev.outstanding, overdue: now.overdue, overduePrev: prev.overdue,
           collectedMtd: now.collected,
           collectedPrevMtd: collectedBetween(invoices, new Date(prevEnd.getFullYear(), prevEnd.getMonth(), 1), sameDay, asOf),
           dso: now.dso, dsoPrev: prev.dso, rate: now.rate, ratePrev: prev.rate, points: pts };
}

/**
 * Customer risk score 0-100 = payment habit (average days late on paid invoices x 0.5, max 40)
 * + how late now (most days overdue x 0.3, max 30) + how much is overdue (1 point per 500, max 20) + no email (10).
 * 60+ High, 30+ Medium, else Low. The same formula runs on the Customers sheet.
 */
function riskScores(invoices, customers, asOf) {
  const out = {};
  Object.keys(customers).forEach(function (name) {
    const late = [];
    let maxd = 0, ovd = 0;
    invoices.forEach(function (i) {
      if (i.customer !== name) return;
      const st = invoiceState(i, asOf);
      if (st.balance <= 0 && i.paidOn instanceof Date) late.push(Math.max(0, daysBetween(i.due, i.paidOn)));
      maxd = Math.max(maxd, st.days);
      if (st.status === 'Overdue') ovd += st.balance;
    });
    let sum = 0;
    late.forEach(function (x) { sum += x; });
    const avg = late.length ? sum / late.length : 0;
    ovd = round2(ovd);
    const score = Math.round(Math.min(40, avg * 0.5) + Math.min(30, maxd * 0.3) + Math.min(20, ovd / 500) + (customers[name].email ? 0 : 10));
    out[name] = { avgLate: avg, maxDays: maxd, overdue: ovd, score: score, level: score >= 60 ? 'High' : score >= 30 ? 'Medium' : 'Low' };
  });
  return out;
}

/**
 * The owner's to-do list for today: one action per open invoice, Urgent > High > Medium > Low, then the biggest balance.
 *   Urgent: a promise to pay was not kept, or the invoice is over 60 days late (call, do not email).
 *   High:   Final / Second reminder to send, or a reminder is due but there is no email (call instead).
 *   Medium: Friendly reminder, or the customer replied: record a promise to pay.
 *   Low:    Upcoming heads-up, a held reminder, or a promised payment due within 2 days: check it came in.
 */
function actionQueue(invoices, customers, plan, pauses, asOf) {
  const stagePri = { Final: 'High', Second: 'High', Friendly: 'Medium', Upcoming: 'Low' };
  const byNo = {};
  plan.forEach(function (p) { byNo[p.no] = p; });
  const q = [];
  invoices.forEach(function (i) {
    const st = invoiceState(i, asOf);
    if (st.balance <= 0) return;
    const pz = pauses[i.no], pl = byNo[i.no], status = (customers[i.customer] || {}).status || 'Active';
    let a = null;
    if (pz && pz.broken) a = ['Urgent', 'Call: promise not kept'];
    else if (status !== 'Active') a = null;
    else if (pz && pz.paused && pz.type === 'Reply') a = ['Medium', 'Record a promise to pay'];
    else if (pz && pz.paused) a = daysBetween(asOf, pz.promised) <= 2 ? ['Low', 'Check the promised payment'] : null;
    else if (st.days > 60) a = ['Urgent', 'Call: over 60 days late'];
    else if (pl && pl.action === 'skip') a = ['High', 'Call: no email on file'];
    else if (pl && pl.action === 'send') a = [stagePri[pl.stage], 'Send ' + pl.stage + ' reminder'];
    else if (pl && pl.action === 'hold') a = ['Low', 'Send ' + pl.stage + ' reminder (tomorrow)'];
    if (a) {
      q.push({ priority: a[0], action: a[1], no: i.no, customer: i.customer, balance: st.balance, days: st.days,
               why: pz && pz.reason ? pz.reason : st.days ? st.days + ' days late' : 'due ' + fmtDate(i.due) });
    }
  });
  return q.sort(function (x, y) {
    return PRIORITY.indexOf(x.priority) - PRIORITY.indexOf(y.priority) || y.balance - x.balance || String(x.no).localeCompare(String(y.no));
  });
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

/** Dark-gold palette shared by the sheets, the sidebar and the dialogs. */
const THEME = { BG: '#0B0F13', PAGE: '#12171D', CARD: '#1A2129', CARD2: '#202832', LINE: '#2A333D', GOLD: '#C9A35B',
                GOLD_LT: '#E6CF9A', IVORY: '#ECE6DA', MUTED: '#8E98A3', JADE: '#4FA38C', COPPER: '#B5754A', SKY: '#7FA7C9',
                GREY: '#4A5560', GOOD: '#5BBF8A', BAD: '#E5735F', AMBER: '#D9A441' };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invoice tracker')
    .addItem('Open collections desk', 'showSidebar')
    .addItem('Preview reminders', 'menuPreview')
    .addItem('Send reminders now', 'menuSend')
    .addItem('Email me the summary', 'menuSummary')
    .addItem('Refresh action queue', 'refreshDesk')
    .addSeparator()
    .addItem('Turn on daily automation', 'installTriggers')
    .addItem('Turn off automation', 'removeTriggers')
    .addSeparator()
    .addItem('Run self-test', 'menuSelfTest')
    .addItem('Rebuild demo data', 'menuRebuildDemo')
    .addToUi();
}

/** Simple trigger: typing a payment, a reply or a customer status refreshes the action queue and the risk table. */
function onEdit(e) {
  try {
    const name = e && e.range ? e.range.getSheet().getName() : '';
    if ([SHEET.INVOICES, SHEET.CUSTOMERS, SHEET.REPLIES, SHEET.SETTINGS].indexOf(name) >= 0) refreshDesk_();
  } catch (err) {
    console.warn('onEdit refresh skipped: ' + err);
  }
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
    company: v['Company'], senderName: v['Sender name'], ownerEmail: v['Owner email'] || '',
    currency: v['Currency symbol'] || '', mode: MODES.indexOf(v['Email mode']) >= 0 ? v['Email mode'] : 'Preview only',
    maxPerRun: Number(v['Max emails per run']) || 25, summaryDay: v['Summary day'] || 'Monday',
    pauseDays: v['Pause after a reply (days)'] === '' || v['Pause after a reply (days)'] === undefined ? 7 : Number(v['Pause after a reply (days)']),
    asOf: v['As-of used'] instanceof Date ? v['As-of used'] : new Date(),  // the same date the formulas use
    stages: stages, templates: templates,
  };
}

/** Owner email from Settings, or the account running the script (asked only when an email is really sent). */
function ownerEmail_(s) { return s.ownerEmail || Session.getEffectiveUser().getEmail(); }

function readInvoices_() {
  const sh = ss_().getSheetByName(SHEET.INVOICES);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 13).getValues().map(function (r, i) {
    return { row: i + 2, no: r[0], customer: r[1], issue: r[2], due: r[3], amount: r[4], paid: r[5], paidOn: r[6] instanceof Date ? r[6] : null,
             sheet: { balance: r[8], status: r[9], days: r[10], aging: r[11] } };
  }).filter(function (r) { return r.no !== '' && r.due instanceof Date; });
}

function readCustomers_() {
  const sh = ss_().getSheetByName(SHEET.CUSTOMERS);
  const out = {};
  if (sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues().forEach(function (r) {
    if (r[0]) out[r[0]] = { contact: r[1], email: String(r[2]).trim(), status: r[9] || 'Active', sheet: { avg: r[5], max: r[6], score: r[7], level: r[8] } };
  });
  return out;
}

/** Replies and promises to pay typed on the Replies sheet. */
function readReplies_() {
  const sh = ss_().getSheetByName(SHEET.REPLIES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().map(function (r) {
    return { date: r[0], no: r[1], type: r[3], promised: r[4] instanceof Date ? r[4] : null };
  }).filter(function (r) { return r.no && r.date instanceof Date; });
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

/** Everything the desk needs, read once: settings, invoices, pauses, today's plan, the queue and the risk scores. */
function deskState_() {
  const s = readSettings_(), inv = readInvoices_(), cust = readCustomers_();
  const pauses = pauseState(readReplies_(), s.asOf, s.pauseDays);
  const plan = planReminders(inv, cust, readSent_(), s, s.asOf, pauses);
  return { s: s, inv: inv, cust: cust, pauses: pauses, plan: plan, queue: actionQueue(inv, cust, plan, pauses, s.asOf),
           risk: riskScores(inv, cust, s.asOf) };
}

/** Menu: Refresh action queue (functions ending in _ are private in Apps Script and cannot be menu items). */
function refreshDesk() { refreshDesk_(); }

const DESK = { QUEUE_ROW: 29, QUEUE_ROWS: 10, RISK_ROW: 42, RISK_ROWS: 6 };

/** Writes the owner action queue and the riskiest customers on the Dashboard. */
function refreshDesk_() {
  const d = deskState_();
  const sh = ss_().getSheetByName(SHEET.DASHBOARD);
  const rows = [];
  for (let i = 0; i < DESK.QUEUE_ROWS; i++) {
    const q = d.queue[i];
    rows.push(q ? [q.priority, q.no, q.customer, '', q.action, '', q.why, '', q.days, q.balance] : ['', '', '', '', '', '', '', '', '', '']);
  }
  if (!d.queue.length) rows[0][4] = 'Nothing to do today';
  sh.getRange(DESK.QUEUE_ROW, 2, DESK.QUEUE_ROWS, 10).setValues(rows);
  const urgent = d.queue.filter(function (q) { return q.priority === 'Urgent'; }).length;
  sh.getRange('H27').setValue(d.queue.length + ' actions today  ·  ' + urgent + ' urgent' +
    (d.queue.length > DESK.QUEUE_ROWS ? '  ·  top ' + DESK.QUEUE_ROWS + ' shown' : ''));
  const top = Object.keys(d.risk).map(function (k) { return [k, d.risk[k]]; })
    .sort(function (a, b) { return b[1].score - a[1].score || b[1].overdue - a[1].overdue; }).slice(0, DESK.RISK_ROWS);
  const rr = [];
  for (let i = 0; i < DESK.RISK_ROWS; i++) {
    const t = top[i];
    rr.push(t ? [t[0], '', t[1].level, t[1].score, t[1].overdue, Math.round(t[1].avgLate)] : ['', '', '', '', '', '']);
  }
  sh.getRange(DESK.RISK_ROW, 6, DESK.RISK_ROWS, 6).setValues(rr);
  sh.getRange('B57').setValue('Action queue refreshed ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy HH:mm') +
    ' (on every edit, daily run and Menu > Refresh action queue)');
  return d;
}

/** Plans and (depending on Email mode) sends reminders. Returns the plan with what happened to each line. */
function runReminders_(forceMode) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another run is still busy. Try again in a minute.');
  try {
    const d = deskState_(), s = d.s;
    const mode = forceMode || s.mode;
    const plan = d.plan;
    const log = ss_().getSheetByName(SHEET.LOG);
    const now = new Date();
    const rows = [];
    plan.forEach(function (p) {
      const mail = buildReminder(p, s);
      p.subject = mail.subject; p.body = mail.body;
      if (mode === 'Preview only') { p.result = p.action === 'send' ? 'Preview' : p.reason; return; }
      if (p.action !== 'send') { p.result = p.reason; rows.push([now, p.no, p.customer, p.email, p.stage, mode, p.reason]); return; }
      if (MailApp.getRemainingDailyQuota() < 1) { p.result = 'Daily email quota used up'; rows.push([now, p.no, p.customer, p.email, p.stage, mode, p.result]); return; }
      const to = mode === 'Send to customers' ? p.email : ownerEmail_(s);
      MailApp.sendEmail({ to: to, subject: (mode === 'Send to customers' ? '' : '[TEST to ' + p.email + '] ') + mail.subject,
                          body: mail.body, name: s.senderName, replyTo: ownerEmail_(s) });
      p.result = mode === 'Send to customers' ? 'Sent to customer' : 'Sent to me (test)';
      rows.push([now, p.no, p.customer, to, p.stage, mode, p.result]);
      if (mode === 'Send to customers') {
        ss_().getSheetByName(SHEET.INVOICES).getRange(rowOf_(p.no), 13).setValue(p.stage + ' ' + fmtDate(now));
      }
    });
    if (rows.length) log.getRange(log.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    return { mode: mode, plan: plan, asOf: s.asOf };
  } finally {
    lock.releaseLock();
  }
}

function rowOf_(no) {
  const inv = readInvoices_().filter(function (r) { return r.no === no; })[0];
  return inv ? inv.row : null;
}

function esc_(x) { return String(x === null || x === undefined ? '' : x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

/** Shared CSS for the sidebar and the dialogs: the same dark-gold look as the sheets. */
function css_() {
  const T = THEME;
  return '<style>' +
    'body{margin:0;background:' + T.PAGE + ';color:' + T.IVORY + ';font:13px/1.45 "Segoe UI",Roboto,Arial,sans-serif}' +
    '.brand{color:' + T.GOLD + ';letter-spacing:2px;font-size:9px;font-weight:700;white-space:nowrap}' +
    'h1{font:22px Georgia,serif;margin:2px 0 4px;color:' + T.IVORY + '}h2{font:15px Georgia,serif;margin:0 0 8px;color:' + T.GOLD_LT + '}' +
    '.rule{height:2px;background:' + T.GOLD + ';margin:8px 0 12px}.muted{color:' + T.MUTED + '}' +
    '.card{background:' + T.CARD + ';border:1px solid ' + T.LINE + ';border-left:3px solid ' + T.GOLD + ';border-radius:6px;padding:10px 12px;margin:0 0 10px}' +
    '.kpis{display:grid;grid-template-columns:1fr 1fr;gap:6px}.kpi{background:' + T.CARD2 + ';border-radius:5px;padding:6px 8px}' +
    '.kpi b{display:block;font:18px Georgia,serif;color:' + T.IVORY + '}.kpi span{font-size:10px;color:' + T.GOLD + ';letter-spacing:1px}' +
    'label{display:block;font-size:11px;color:' + T.MUTED + ';margin:6px 0 2px}' +
    'select,input,textarea{width:100%;box-sizing:border-box;background:' + T.BG + ';color:' + T.IVORY + ';border:1px solid ' + T.LINE + ';border-radius:4px;padding:6px;font:13px Arial}' +
    'button{background:' + T.GOLD + ';color:' + T.BG + ';border:0;border-radius:4px;padding:7px 12px;font-weight:700;cursor:pointer;margin-top:8px}' +
    'button.ghost{background:transparent;color:' + T.GOLD_LT + ';border:1px solid ' + T.GOLD + '}button:disabled{opacity:.5}' +
    '.nav{display:flex;flex-wrap:wrap;gap:4px}.nav button{margin:0;padding:4px 8px;font-size:11px;font-weight:400}' +
    '.chip{display:inline-block;border-radius:10px;padding:1px 8px;font-size:11px;font-weight:700}' +
    '.Urgent{background:#4A1F1C;color:' + T.BAD + '}.High{background:#3D2A1C;color:#E39A6B}.Medium{background:#3A3220;color:' + T.AMBER + '}' +
    '.Low{background:#1E2D3A;color:' + T.SKY + '}.send{background:#1F3A33;color:' + T.GOOD + '}.skip{background:#4A1F1C;color:' + T.BAD + '}' +
    '.pause{background:#3A3220;color:' + T.AMBER + '}.hold{background:#1E2D3A;color:' + T.SKY + '}' +
    '.q{display:flex;justify-content:space-between;gap:6px;padding:5px 0;border-bottom:1px solid ' + T.LINE + '}.msg{color:' + T.GOOD + ';min-height:16px;font-size:12px}' +
    'a{color:' + T.GOLD_LT + '}pre{white-space:pre-wrap;font:12px/1.5 Arial;color:' + T.IVORY + ';margin:6px 0 0}' +
    '</style>';
}

function showPlan_(title, res) {
  const T = THEME;
  const n = function (a) { return res.plan.filter(function (p) { return p.action === a; }).length; };
  let h = '<!DOCTYPE html><html><head><base target="_blank">' + css_() + '</head><body style="padding:14px 18px">';
  h += '<div class="brand">NORTHBEAM STUDIO</div><h1>' + esc_(title) + '</h1><div class="rule"></div>';
  h += '<p class="muted" style="margin:0 0 10px">Email mode: <b style="color:' + T.GOLD_LT + '">' + esc_(res.mode) + '</b> &middot; ' +
       res.plan.length + ' line(s): ' + n('send') + ' to send, ' + n('pause') + ' paused (reply or promise logged), ' + n('skip') +
       ' skipped, ' + n('hold') + ' held. Fictional addresses (example.com).</p>';
  res.plan.forEach(function (p, i) {
    h += '<div class="card"><div style="display:flex;justify-content:space-between;align-items:center">' +
         '<div><b style="font-size:14px">' + esc_(p.no) + '</b> &nbsp;' + esc_(p.customer) + ' &nbsp;<span class="chip Low">' + esc_(p.stage) + '</span></div>' +
         '<span class="chip ' + p.action + '">' + esc_(p.action === 'send' ? (p.result || 'send') : p.action + ': ' + (p.reason || '')) + '</span></div>' +
         '<div class="muted" style="margin-top:4px">To: ' + esc_(p.email || '(no email on file)') + ' &middot; ' + money(p.balance, '') + ' &middot; ' + p.days + ' days late</div>' +
         '<div style="margin-top:6px;color:' + T.GOLD_LT + ';font-weight:700">' + esc_(p.subject) + '</div><pre>' + esc_(p.body) + '</pre>' +
         '<div style="margin-top:6px;font-size:12px">&#128206; Statement for ' + esc_(p.customer) + ': <a href="#" onclick="st(this,' + i + ');return false">prepare PDF</a></div></div>';
  });
  if (!res.plan.length) h += '<div class="card">Nothing to send today.</div>';
  h += '<script>const C=' + JSON.stringify(res.plan.map(function (p) { return p.customer; })) + ';' +
       'function st(a,i){a.textContent="preparing...";google.script.run.withSuccessHandler(function(u){a.textContent="open PDF (Statement sheet export)";' +
       'a.href=u;a.onclick=null;}).sidebarStatement(C[i]);}</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(h).setWidth(820).setHeight(600), title);
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
  refreshDesk_();
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
  const to = ownerEmail_(s);
  MailApp.sendEmail({ to: to, subject: s.company + ': ' + money(sum.overdue, s.currency) + ' overdue on ' + sum.overdueCount + ' invoices',
                      htmlBody: summaryHtml(sum, overdue, week, s, s.asOf), name: 'Invoice tracker' });
  return to;
}

function menuSummary() { SpreadsheetApp.getUi().alert('Summary sent to ' + sendSummary_()); }

/** Runs every morning: reminders (in the Email mode on Settings), the action queue, and the summary on the chosen weekday. */
function dailyRun() {
  runReminders_();
  refreshDesk_();
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

// ============================================================================================ collections desk (sidebar)

function showSidebar() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(sidebarHtml_()).setTitle('Collections desk'));
}

function isoDate_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function parseIso_(t) { const p = String(t).split('-'); return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])); }

/** What the sidebar shows: KPIs, the top of the queue, open invoices and customers. */
function sidebarData() {
  const d = deskState_(), s = d.s;
  const k = kpiCards(d.inv, s.asOf);
  const open = d.inv.map(function (r) { return { no: r.no, customer: r.customer, st: invoiceState(r, s.asOf) }; })
    .filter(function (r) { return r.st.balance > 0; })
    .sort(function (a, b) { return b.st.days - a.st.days || b.st.balance - a.st.balance; })
    .map(function (r) { return { no: r.no, customer: r.customer, balance: r.st.balance, days: r.st.days }; });
  return { asOf: fmtDate(s.asOf), asOfIso: isoDate_(s.asOf), mode: s.mode,
           kpi: { outstanding: money(k.outstanding, ''), overdue: money(k.overdue, ''), collected: money(k.collectedMtd, ''),
                  dso: k.dso.toFixed(1), rate: (k.rate * 100).toFixed(1) + '%' },
           queue: d.queue.slice(0, 6).map(function (q) { return { priority: q.priority, action: q.action, no: q.no, customer: q.customer }; }),
           queueCount: d.queue.length, open: open, customers: Object.keys(d.cust),
           sends: d.plan.filter(function (p) { return p.action === 'send'; }).length,
           paused: d.plan.filter(function (p) { return p.action === 'pause'; }).length };
}

/** Record payment: adds the amount to Paid and sets the paid date, then refreshes the desk. */
function sidebarRecordPayment(no, amount, iso) {
  const a = round2(Number(amount));
  if (!(a > 0)) throw new Error('Type an amount above 0.');
  const row = rowOf_(no);
  if (!row) throw new Error('Invoice ' + no + ' not found.');
  const sh = ss_().getSheetByName(SHEET.INVOICES);
  const before = Number(sh.getRange(row, 6).getValue()) || 0;
  sh.getRange(row, 6, 1, 2).setValues([[round2(before + a), parseIso_(iso)]]);
  SpreadsheetApp.flush();
  refreshDesk_();
  const st = invoiceState(readInvoices_().filter(function (r) { return r.no === no; })[0], readSettings_().asOf);
  return 'Recorded ' + money(a, '') + ' on ' + no + ': now ' + st.status + (st.balance > 0 ? ', ' + money(st.balance, '') + ' still open' : '') + '.';
}

/** Log a reply or a promise to pay on the Replies sheet; reminders for that invoice pause by the rule in pauseState. */
function sidebarLogReply(no, type, promisedIso, note) {
  if (REPLY_TYPES.indexOf(type) < 0) throw new Error('Unknown type.');
  if (type === 'Promise to pay' && !promisedIso) throw new Error('Pick the promised date.');
  const s = readSettings_();
  const sh = ss_().getSheetByName(SHEET.REPLIES);
  const r = Math.max(sh.getLastRow(), 1) + 1;
  sh.getRange(r, 1, 1, 2).setValues([[s.asOf, no]]);
  sh.getRange(r, 4, 1, 3).setValues([[type, promisedIso ? parseIso_(promisedIso) : '', note || '']]);
  SpreadsheetApp.flush();
  const d = refreshDesk_();
  const pz = d.pauses[no];
  return 'Logged. ' + no + ': ' + (pz && pz.paused ? 'reminders paused (' + pz.reason + ').' : 'no pause.');
}

function sidebarPreview() { menuPreview(); return 'Preview opened.'; }

/** Demo-safe: in Preview only nothing is sent, the sidebar just says what would go out. */
function sidebarSendDue() {
  const s = readSettings_();
  if (s.mode === 'Preview only') {
    const r = runReminders_('Preview only');
    const n = function (a) { return r.plan.filter(function (p) { return p.action === a; }).length; };
    return 'Email mode is Preview only: nothing was sent. Today ' + n('send') + ' reminder(s) would go out, ' + n('pause') +
           ' are paused (reply or promise), ' + n('skip') + ' skipped. Switch the mode on Settings to send.';
  }
  const r = runReminders_();
  refreshDesk_();
  return r.plan.filter(function (p) { return /^Sent/.test(p.result || ''); }).length + ' reminder(s) sent (' + r.mode + '). See Reminder log.';
}

/** Puts a customer on the Statement sheet and returns the PDF export link of that sheet. */
function sidebarStatement(customer) {
  const sh = ss_().getSheetByName(SHEET.STATEMENT);
  sh.getRange('B6').setValue(customer);
  SpreadsheetApp.flush();
  return statementUrl_();
}

function statementUrl_() {
  const sh = ss_().getSheetByName(SHEET.STATEMENT);
  return 'https://docs.google.com/spreadsheets/d/' + ss_().getId() + '/export?format=pdf&gid=' + sh.getSheetId() +
         '&range=A1:I38&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED&size=A4' +
         '&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4';
}

function goTo(name) { ss_().getSheetByName(name).activate(); return name; }

function sidebarHtml_() {
  const nav = [SHEET.DASHBOARD, SHEET.INVOICES, SHEET.CUSTOMERS, SHEET.REPLIES, SHEET.STATEMENT, SHEET.LOG, SHEET.SETTINGS, SHEET.CHECKS];
  return '<!DOCTYPE html><html><head><base target="_blank">' + css_() + '</head><body style="padding:12px">' +
    '<div class="brand">NORTHBEAM STUDIO</div><h1>Collections desk</h1><div class="muted" id="asof"></div><div class="rule"></div>' +
    '<div class="nav">' + nav.map(function (n) { return '<button class="ghost" onclick="go(\'' + n + '\')">' + n + '</button>'; }).join('') + '</div>' +
    '<div class="kpis" style="margin:10px 0">' +
    '<div class="kpi"><span>OUTSTANDING</span><b id="k1">-</b></div><div class="kpi"><span>OVERDUE</span><b id="k2" style="color:' + THEME.BAD + '">-</b></div>' +
    '<div class="kpi"><span>COLLECTED MTD</span><b id="k3">-</b></div><div class="kpi"><span>DSO / RATE</span><b id="k4">-</b></div></div>' +
    '<div class="card"><h2>Today&#39;s actions <span class="muted" id="qn" style="font:12px Arial"></span></h2><div id="queue"></div></div>' +
    '<div class="card"><h2>Record payment</h2><label>Invoice</label><select id="pinv" onchange="fill()"></select>' +
    '<label>Amount</label><input id="pamt" type="number" step="0.01"><label>Paid on</label><input id="pdate" type="date">' +
    '<button id="b1" onclick="pay()">Record payment</button><div class="msg" id="m1"></div></div>' +
    '<div class="card"><h2>Log a reply or promise to pay</h2><label>Invoice</label><select id="rinv"></select>' +
    '<label>Type</label><select id="rtype" onchange="pd()"><option>Promise to pay</option><option>Reply</option></select>' +
    '<div id="pdw"><label>Promised date</label><input id="rdate" type="date"></div><label>Note</label><input id="rnote" placeholder="e.g. paying on their next run">' +
    '<button id="b2" onclick="reply()">Log it</button><div class="msg" id="m2"></div></div>' +
    '<div class="card"><h2>Reminders</h2><div class="muted" id="rem"></div><button class="ghost" onclick="prev()">Preview reminders</button> ' +
    '<button onclick="send()">Send due reminders</button><div class="msg" id="m3"></div></div>' +
    '<div class="card"><h2>Statement PDF</h2><label>Customer</label><select id="scust"></select>' +
    '<button class="ghost" onclick="stmt()">Prepare statement</button><div class="msg" id="m4"></div></div>' +
    '<script>' +
    'let D=null;const $=function(i){return document.getElementById(i)};' +
    'function m(i,t,bad){$(i).style.color=bad?"' + THEME.BAD + '":"' + THEME.GOOD + '";$(i).textContent=t}' +
    'function load(){google.script.run.withSuccessHandler(render).sidebarData()}' +
    'function render(d){D=d;$("asof").textContent="As of "+d.asOf+" \u00b7 Email mode: "+d.mode;$("k1").textContent=d.kpi.outstanding;' +
    '$("k2").textContent=d.kpi.overdue;$("k3").textContent=d.kpi.collected;$("k4").textContent=d.kpi.dso+"d \u00b7 "+d.kpi.rate;' +
    '$("qn").textContent=d.queueCount+" in total";$("queue").innerHTML=d.queue.map(function(q){return "<div class=q><span><span class=\\"chip "+q.priority+"\\">"+q.priority+"</span> "+q.action+"</span><span class=muted>"+q.no+"</span></div>"}).join("")||"Nothing to do today";' +
    'const o=d.open.map(function(r){return "<option value=\\""+r.no+"\\">"+r.no+" \u00b7 "+r.customer+" \u00b7 "+r.balance.toFixed(2)+(r.days?" \u00b7 "+r.days+"d late":"")+"</option>"}).join("");' +
    '$("pinv").innerHTML=o;$("rinv").innerHTML=o;$("scust").innerHTML=d.customers.map(function(c){return "<option>"+c+"</option>"}).join("");' +
    '$("pdate").value=d.asOfIso;fill();$("rem").textContent=d.sends+" due today, "+d.paused+" paused by a reply or promise."}' +
    'function fill(){const r=D.open.filter(function(x){return x.no===$("pinv").value})[0];$("pamt").value=r?r.balance.toFixed(2):""}' +
    'function pd(){$("pdw").style.display=$("rtype").value==="Reply"?"none":"block"}' +
    'function fail(i){return function(e){m(i,e.message,1)}}' +
    'function pay(){m("m1","Saving...");google.script.run.withSuccessHandler(function(t){m("m1",t);load()}).withFailureHandler(fail("m1")).sidebarRecordPayment($("pinv").value,$("pamt").value,$("pdate").value)}' +
    'function reply(){m("m2","Saving...");google.script.run.withSuccessHandler(function(t){m("m2",t);load()}).withFailureHandler(fail("m2")).sidebarLogReply($("rinv").value,$("rtype").value,$("rdate").value,$("rnote").value)}' +
    'function prev(){google.script.run.sidebarPreview()}' +
    'function send(){if(D.mode!=="Preview only"&&!confirm("Send reminders now ("+D.mode+")?"))return;m("m3","Working...");google.script.run.withSuccessHandler(function(t){m("m3",t);load()}).withFailureHandler(fail("m3")).sidebarSendDue()}' +
    'function stmt(){m("m4","Preparing...");google.script.run.withSuccessHandler(function(u){$("m4").innerHTML="Ready: <a href=\\""+u+"\\">open the PDF</a> (Statement sheet export)"}).withFailureHandler(fail("m4")).sidebarStatement($("scust").value)}' +
    'function go(n){google.script.run.goTo(n)}' +
    'load();</script></body></html>';
}

// ============================================================================================ self-test

/** Unit tests (Tests.gs) + every formula on the sheets compared with the script logic. */
function selfTest() {
  const results = runAllTests().map(function (t) { return ['Unit test', t.name, t.pass ? 'PASS' : 'FAIL', t.detail || '']; });
  const add = function (type, name, ok, detail) { results.push([type, name, ok ? 'PASS' : 'FAIL', ok ? '' : detail]); };
  const near = function (a, b, tol) { return Math.abs(Number(a) - Number(b)) < (tol || 0.005); };
  const d = refreshDesk_(), s = d.s, inv = d.inv;
  let bad = [];
  inv.forEach(function (r) {
    const st = invoiceState(r, s.asOf);
    const ok = Math.abs(st.balance - r.sheet.balance) < 0.005 && st.status === r.sheet.status && st.days === r.sheet.days && st.aging === r.sheet.aging;
    if (!ok) bad.push(r.no + ' sheet ' + [r.sheet.status, r.sheet.days, r.sheet.aging].join('/') + ' script ' + [st.status, st.days, st.aging].join('/'));
  });
  add('Sheet formulas', 'Balance, status, days and aging on all ' + inv.length + ' invoices = script', !bad.length, bad.slice(0, 3).join('; '));
  const sum = summarize(inv.map(function (r) { return { no: r.no, customer: r.customer, state: invoiceState(r, s.asOf) }; }));
  const nr = function (n) { return ss_().getRangeByName(n).getValue(); };
  add('Dashboard', 'Outstanding on Dashboard = script', near(nr('DashOutstanding'), sum.outstanding), nr('DashOutstanding') + ' vs ' + sum.outstanding);
  add('Dashboard', 'Overdue on Dashboard = script', near(nr('DashOverdue'), sum.overdue), nr('DashOverdue') + ' vs ' + sum.overdue);
  const aging = ss_().getRangeByName('DashAging').getValues();
  add('Dashboard', 'Aging buckets on Dashboard = script', aging.every(function (r) { return near(r[1], sum.aging[r[0]]); }), JSON.stringify(aging));
  const k = kpiCards(inv, s.asOf);
  add('Dashboard', 'Collected this month ' + money(k.collectedMtd, '') + ' and the same days last month = script',
    near(nr('DashCollected'), k.collectedMtd) && near(nr('CalcPrevMtd'), k.collectedPrevMtd), nr('DashCollected') + '/' + nr('CalcPrevMtd'));
  add('Dashboard', 'DSO ' + k.dso.toFixed(1) + ' days = script (outstanding / sales of the last 90 days x 90)', near(nr('DashDSO'), k.dso, 0.001), nr('DashDSO') + ' vs ' + k.dso);
  add('Dashboard', 'Collection rate ' + (k.rate * 100).toFixed(1) + '% = script', near(nr('DashRate'), k.rate, 0.00001), nr('DashRate') + ' vs ' + k.rate);
  const calc = ss_().getRangeByName('CalcTrend').getValues();   // D:J of the 6 points
  const tBad = k.points.filter(function (p, i) {
    const c = calc[i];
    return !near(c[0], p.outstanding) || !near(c[1], p.overdue) || !near(c[2], p.collected) || !near(c[3], p.billed) ||
           !near(c[5], p.dso, 0.001) || !near(c[6], p.rate, 0.00001);
  });
  add('Trends', '6-month trends (outstanding, overdue, collected, billed, DSO, rate) on Calc = script', !tBad.length,
    tBad.map(function (p) { return p.label; }).join(','));
  const per = {};
  inv.forEach(function (r) {
    const st = invoiceState(r, s.asOf), p = per[r.customer] = per[r.customer] || [0, 0];
    if (st.balance > 0) p[0] += st.balance;
    if (st.status === 'Overdue') p[1] += st.balance;
  });
  const cs = ss_().getSheetByName(SHEET.CUSTOMERS).getDataRange().getValues().slice(1).filter(function (r) { return r[0] !== ''; });
  const custBad = cs.filter(function (r) { const p = per[r[0]] || [0, 0]; return Math.abs(r[3] - p[0]) > 0.005 || Math.abs(r[4] - p[1]) > 0.005; });
  add('Customers', 'Open balance and overdue for all ' + cs.length + ' customers = script', !custBad.length,
    custBad.slice(0, 3).map(function (r) { return r[0] + ' ' + r[3] + '/' + r[4]; }).join('; '));
  const rBad = Object.keys(d.cust).filter(function (c) {
    const sh = d.cust[c].sheet, r = d.risk[c];
    return sh.score !== r.score || sh.level !== r.level || sh.max !== r.maxDays || !near(sh.avg, r.avgLate, 0.001);
  });
  add('Customers', 'Risk score and level for all ' + Object.keys(d.cust).length + ' customers (formulas) = script', !rBad.length,
    rBad.slice(0, 3).map(function (c) { return c + ' ' + JSON.stringify(d.cust[c].sheet) + ' vs ' + d.risk[c].score; }).join('; '));
  const qs = ss_().getSheetByName(SHEET.DASHBOARD).getRange(DESK.QUEUE_ROW, 2, DESK.QUEUE_ROWS, 5).getValues()
    .filter(function (r) { return r[0]; }).map(function (r) { return [r[0], r[1], r[4]].join('|'); });
  const qw = d.queue.slice(0, DESK.QUEUE_ROWS).map(function (q) { return [q.priority, q.no, q.action].join('|'); });
  add('Action queue', 'Owner action queue on Dashboard = script (' + d.queue.length + ' actions, ranked)', qs.join() === qw.join(), qs.join() + ' vs ' + qw.join());
  const pausedNos = Object.keys(d.pauses).filter(function (k2) { return d.pauses[k2].paused; });
  const leak = d.plan.filter(function (p) { return pausedNos.indexOf(p.no) >= 0 && p.action !== 'pause'; });
  add('Pause rule', pausedNos.length + ' invoice(s) paused by a logged reply/promise; none of them is in today\'s emails', !leak.length,
    leak.map(function (p) { return p.no; }).join(','));
  const sh = ss_().getSheetByName(SHEET.CHECKS);
  sh.getRange('A5:D200').clearContent();
  sh.getRange(5, 1, results.length, 4).setValues(results);
  sh.getRange('B3').setValue(new Date());
  const failed = results.filter(function (r) { return r[2] !== 'PASS'; }).length;
  sh.getRange('C3').setValue(failed ? failed + ' FAILED' : 'ALL ' + results.length + ' PASS');
  return { total: results.length, failed: failed };
}

function menuSelfTest() {
  const r = selfTest();
  ss_().getSheetByName(SHEET.CHECKS).activate();
  SpreadsheetApp.getUi().alert(r.failed ? r.failed + ' of ' + r.total + ' checks FAILED - see the Checks sheet.' : 'All ' + r.total + ' checks pass.');
}
