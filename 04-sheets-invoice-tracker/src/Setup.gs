/**
 * Builds the whole tracker in the current spreadsheet: Dashboard, Invoices, Customers, Replies, Statement, Reminder log,
 * Settings, Checks (+ a hidden Calc sheet for the trends), with formulas, drop-downs, the dark-gold look, charts,
 * navigation buttons and the fictional demo data from DemoData.gs.
 * Run "Rebuild demo data" from the menu (or buildDemo from the editor) in an empty Google Sheet.
 */

const T = THEME;
const CHIP = { Urgent: ['#4A1F1C', T.BAD], High: ['#3D2A1C', '#E39A6B'], Medium: ['#3A3220', T.AMBER], Low: ['#1E2D3A', T.SKY] };
const HEAT = { 'Not due': ['#1F3A33', T.GOOD, T.JADE], '1-30': ['#3A3220', T.AMBER, T.AMBER], '31-60': ['#3D2A1C', '#E39A6B', T.COPPER],
               '61-90': ['#4A1F1C', T.BAD, T.BAD], '90+': ['#5C1A17', '#FF8A75', '#C2452F'] };

function menuRebuildDemo() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Rebuild the demo?', 'This replaces every sheet in this file with fresh demo data.', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  buildDemo();
  ui.alert('Demo rebuilt. Try Invoice tracker > Open collections desk.');
}

function buildDemo() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetLocale('en_US');
  ss.setSpreadsheetTimeZone(Session.getScriptTimeZone());
  const temp = ss.insertSheet('tmp' + Date.now());
  // charts and cross-sheet formulas first, then one sheet at a time: deleting the whole linked set in one go
  // made Google fail with "Service Spreadsheets failed" on a rebuild
  ss.getSheets().forEach(function (s) { s.getCharts().forEach(function (c) { s.removeChart(c); }); });
  ss.getSheets().forEach(function (s) { if (s.getSheetId() !== temp.getSheetId()) s.clear(); });
  SpreadsheetApp.flush();
  ss.getSheets().forEach(function (s) { if (s.getSheetId() !== temp.getSheetId()) { ss.deleteSheet(s); SpreadsheetApp.flush(); } });
  ss.getNamedRanges().forEach(function (n) { n.remove(); });
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const day = function (offset) { const d = new Date(today); d.setDate(d.getDate() + offset); return d; };

  const order = [SHEET.DASHBOARD, SHEET.INVOICES, SHEET.CUSTOMERS, SHEET.REPLIES, SHEET.STATEMENT, SHEET.LOG, SHEET.SETTINGS, SHEET.CHECKS, SHEET.CALC];
  const sh = {};
  order.forEach(function (n) { sh[n] = ss.insertSheet(n); });
  ss.deleteSheet(temp);
  const tabColor = { Dashboard: T.GOLD, Statement: T.GOLD_LT, Checks: T.GOOD };
  order.forEach(function (n) { sh[n].setTabColor(tabColor[n] || T.GREY); });

  buildSettings_(ss, sh[SHEET.SETTINGS]);
  buildCustomers_(ss, sh[SHEET.CUSTOMERS]);
  buildInvoices_(ss, sh[SHEET.INVOICES], day);
  buildLog_(ss, sh[SHEET.LOG], sh[SHEET.INVOICES], day);
  buildReplies_(ss, sh[SHEET.REPLIES], day);
  buildCalc_(ss, sh[SHEET.CALC]);
  buildDashboard_(ss, sh[SHEET.DASHBOARD], sh[SHEET.CALC]);
  buildStatement_(ss, sh[SHEET.STATEMENT]);
  buildChecks_(sh[SHEET.CHECKS]);
  buildNav_(ss, sh);
  SpreadsheetApp.flush();
  selfTest();
  sh[SHEET.DASHBOARD].activate();
}

// ------------------------------------------------------------------------------------------- look helpers

/** Trims the sheet to rows x cols and paints it in the page colour with ivory text. */
function darkSheet_(sh, rows, cols) {
  if (sh.getMaxRows() > rows) sh.deleteRows(rows + 1, sh.getMaxRows() - rows);
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() > cols) sh.deleteColumns(cols + 1, sh.getMaxColumns() - cols);
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  sh.getRange(1, 1, rows, cols).setBackground(T.PAGE).setFontColor(T.IVORY).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sh.setHiddenGridlines(true);
}

function header_(range, color) {
  range.setFontWeight('bold').setFontColor(color || T.GOLD).setBackground(T.BG).setVerticalAlignment('middle')
    .setBorder(null, null, true, null, null, null, T.GOLD, SpreadsheetApp.BorderStyle.SOLID);
}

/** Title block used on every sheet: spaced gold brand line, serif title, thin gold rule. */
function titleBlock_(sh, title, sub, cols) {
  sh.getRange(1, 1, 2, cols).setBackground(T.BG);
  sh.getRange('B1').setValue('N O R T H B E A M   S T U D I O').setFontColor(T.GOLD).setFontSize(8).setFontWeight('bold');
  sh.getRange('B2').setValue(title).setFontFamily('Georgia').setFontSize(18).setFontColor(T.IVORY);
  if (sub) sh.getRange(2, 4).setValue(sub).setFontColor(T.MUTED).setFontSize(9);
  sh.getRange(3, 1, 1, cols).setBackground(T.GOLD);
  sh.setRowHeight(1, 20); sh.setRowHeight(2, 36); sh.setRowHeight(3, 2);
}

function chips_(range, map) {
  return Object.keys(map).map(function (k) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(k).setBackground(map[k][0]).setFontColor(map[k][1]).setBold(true).setRanges([range]).build();
  });
}

function banding_(range) {
  range.getBandings().forEach(function (b) { b.remove(); });
  range.applyRowBanding().setHeaderRowColor(T.BG).setFirstRowColor(T.CARD).setSecondRowColor(T.CARD2);
}

function chartStyle_(b) {
  return b.setOption('backgroundColor', T.CARD).setOption('legend', { position: 'top', textStyle: { color: T.IVORY } })
    .setOption('hAxis', { textStyle: { color: T.IVORY } }).setOption('vAxis', { textStyle: { color: T.MUTED }, gridlines: { color: T.LINE } });
}

// ------------------------------------------------------------------------------------------- sheets

function buildSettings_(ss, sh) {
  darkSheet_(sh, 30, 6);
  const rows = [['Setting', 'Value', 'Note'],
    ['Company', 'Northbeam Studio', 'Fictional design agency used for the demo'],
    ['Sender name', 'Northbeam Studio Accounts', 'Name shown on reminder emails'],
    ['Owner email', '', 'Summary and test emails go here. Blank = the account running the script'],
    ['Currency symbol', '$', ''],
    ['Email mode', 'Preview only', 'Preview only: nothing is sent. Send to me (test): every reminder comes to you. Send to customers: live.'],
    ['Max emails per run', 25, 'Safety limit; the rest wait for the next run'],
    ['Summary day', 'Monday', 'Weekday the summary email goes out (daily automation)'],
    ['As-of date', '', 'Blank = today. Type a date to see the tracker as on that day'],
    ['As-of used', '=IF(ISDATE(B9),B9,TODAY())', 'What every formula and the script use'],
    ['Pause after a reply (days)', DEMO.pauseDays, 'A logged reply pauses that invoice\'s reminders this long; a promise to pay pauses them until the promised date']];
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  header_(sh.getRange('A1:C1'));
  sh.getRange('A2:C11').setBackground(T.CARD).setBorder(null, null, true, null, null, true, T.LINE, SpreadsheetApp.BorderStyle.SOLID);
  [2, 3, 4, 5, 6, 7, 8, 9, 11].forEach(function (r) { sh.getRange(r, 2).setBackground(T.CARD2).setFontColor(T.GOLD_LT).setFontWeight('bold'); });
  sh.getRange('B10').setFontColor(T.SKY);
  sh.getRange('B6').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(MODES, true).build());
  sh.getRange('B8').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'], true).build());
  sh.getRange('B9:B10').setNumberFormat('d mmm yyyy').setHorizontalAlignment('left');
  sh.getRange('B7').setHorizontalAlignment('left'); sh.getRange('B11').setHorizontalAlignment('left');
  ss.setNamedRange('AsOf', sh.getRange('B10'));
  ss.setNamedRange('PauseDays', sh.getRange('B11'));

  const st = [['Stage', 'Days after due', 'Email subject', 'Email body']];
  const sig = '\n\nBest regards,\n{sender}\n{company}';
  st.push(['Upcoming', -3, 'Invoice {invoice} is due on {due}',
    'Hi {contact},\n\nA quick heads-up: invoice {invoice} for {amount} is due on {due}. If payment is already on its way, thank you!' + sig]);
  st.push(['Friendly', 1, 'Friendly reminder: invoice {invoice} is now due',
    'Hi {contact},\n\nOur records show invoice {invoice} for {amount} was due on {due}. Could you let us know when we can expect payment? If you have already paid, please ignore this note.' + sig]);
  st.push(['Second', 14, 'Second reminder: invoice {invoice} is {days} days overdue',
    'Hi {contact},\n\nInvoice {invoice} still shows {amount} outstanding, {days} days after the due date of {due}. Please arrange payment this week, or reply if there is a problem with the invoice.' + sig]);
  st.push(['Final', 30, 'Final reminder: invoice {invoice} ({amount}) is {days} days overdue',
    'Hi {contact},\n\nThis is our final reminder for invoice {invoice}: {amount}, now {days} days overdue (due {due}). Please pay within 7 days or contact us to agree a payment plan.' + sig]);
  sh.getRange('A14').setValue('Reminder stages   ·   placeholders: {contact} {customer} {invoice} {amount} {due} {days} {company} {sender}')
    .setFontFamily('Georgia').setFontSize(12).setFontColor(T.GOLD_LT);
  sh.getRange(15, 1, st.length, 4).setValues(st);
  header_(sh.getRange('A15:D15'));
  sh.getRange('A16:D19').setBackground(T.CARD2).setFontColor(T.IVORY).setWrap(true).setVerticalAlignment('top')
    .setBorder(null, null, true, null, null, true, T.LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('A16:A19').setFontColor(T.GOLD_LT).setFontWeight('bold');
  sh.getRange('B16:B19').setHorizontalAlignment('center');
  sh.getRange('A21').setValue('Gold text = you type it. Blue = calculated. The script reads these cells by their label, so keep column A as it is.')
    .setFontColor(T.MUTED).setFontStyle('italic');
  sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 330); sh.setColumnWidth(4, 520); sh.setColumnWidth(5, 20); sh.setColumnWidth(6, 120);
  sh.getRange('C2:C11').setFontColor(T.MUTED).setWrap(true);
  sh.setFrozenRows(1);
}

function buildCustomers_(ss, sh) {
  darkSheet_(sh, 120, 12);
  const rows = [['Customer', 'Contact', 'Email', 'Open balance', 'Overdue', 'Avg days late (paid)', 'Most days late now', 'Risk score', 'Risk', 'Status']];
  DEMO.customers.forEach(function (c) { rows.push([c.name, c.contact, c.email, '', '', '', '', '', '', c.status || 'Active']); });
  sh.getRange(1, 1, rows.length, 10).setValues(rows);
  // BYROW / MAP, because SUMIFS inside ARRAYFORMULA does not run once per row
  sh.getRange('D2').setFormula('=BYROW(A2:A,LAMBDA(c,IF(c="",,SUMIFS(Invoices!I2:I,Invoices!B2:B,c,Invoices!I2:I,">0"))))');
  sh.getRange('E2').setFormula('=BYROW(A2:A,LAMBDA(c,IF(c="",,SUMIFS(Invoices!I2:I,Invoices!B2:B,c,Invoices!J2:J,"Overdue"))))');
  sh.getRange('F2').setFormula('=BYROW(A2:A,LAMBDA(c,IF(c="",,IFERROR(AVERAGEIF(Invoices!B2:B,c,Invoices!N2:N),0))))');
  sh.getRange('G2').setFormula('=BYROW(A2:A,LAMBDA(c,IF(c="",,MAXIFS(Invoices!K2:K,Invoices!B2:B,c))))');
  // risk score: payment habit (max 40) + how late now (max 30) + overdue money (max 20) + no email (10); same rule as riskScores()
  sh.getRange('H2').setFormula('=MAP(A2:A,C2:C,E2:E,F2:F,G2:G,LAMBDA(c,em,ov,av,mx,IF(c="",,ROUND(MIN(40,av*0.5)+MIN(30,mx*0.3)+MIN(20,ROUND(ov,2)/500)+IF(em="",10,0)))))');
  sh.getRange('I2').setFormula('=MAP(A2:A,H2:H,LAMBDA(c,h,IF(c="",,IF(h>=60,"High",IF(h>=30,"Medium","Low")))))');
  header_(sh.getRange('A1:J1'));
  sh.getRange('D1:I1').setFontColor(T.SKY);
  banding_(sh.getRange('A1:J120'));
  sh.getRange('A2:C120').setFontColor(T.GOLD_LT);
  sh.getRange('D2:E120').setNumberFormat('#,##0.00');
  sh.getRange('F2:F120').setNumberFormat('0.0');
  sh.getRange('G2:H120').setNumberFormat('0');
  sh.getRange('F2:I120').setHorizontalAlignment('center');
  sh.getRange('J2:J120').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build()).setFontColor(T.GOLD_LT);
  sh.setFrozenRows(1);
  [200, 130, 250, 105, 100, 130, 125, 85, 80, 105, 16, 120].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  const blank = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",$C2="")')
    .setBackground('#4A1F1C').setRanges([sh.getRange('C2:C120')]).build();
  const status = chips_(sh.getRange('J2:J120'), { 'On hold': CHIP.Medium, Collections: CHIP.Urgent, Active: ['#1F3A33', T.GOOD] });
  sh.setConditionalFormatRules([blank].concat(chips_(sh.getRange('I2:I120'), { High: CHIP.Urgent, Medium: CHIP.Medium, Low: ['#1F3A33', T.GOOD] }), status));
}

function buildInvoices_(ss, sh, day) {
  darkSheet_(sh, 500, 15);
  const hdr = ['Invoice no', 'Customer', 'Issue date', 'Due date', 'Amount', 'Paid', 'Paid date', 'Notes'];
  const rows = DEMO.invoices.map(function (i) {
    return [i.no, i.customer, day(i.issue), day(i.due), i.amount, i.paid || '', i.paid_on === null ? '' : day(i.paid_on), i.notes];
  });
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  sh.getRange(2, 1, rows.length, hdr.length).setValues(rows);
  sh.getRange('I1').setFormula('={"Balance";ARRAYFORMULA(IF(A2:A="",,ROUND(E2:E-F2:F,2)))}');
  sh.getRange('J1').setFormula('={"Status";ARRAYFORMULA(IF(A2:A="",,IF(I2:I<0,"Overpaid",IF(I2:I=0,"Paid",IF(D2:D<AsOf,"Overdue",IF(F2:F>0,"Part paid","Not due"))))))}');
  sh.getRange('K1').setFormula('={"Days overdue";ARRAYFORMULA(IF(A2:A="",,IF((I2:I>0)*(D2:D<AsOf),AsOf-D2:D,0)))}');
  sh.getRange('L1').setFormula('={"Aging";ARRAYFORMULA(IF(A2:A="",,IF(I2:I<=0,"Paid",IF(K2:K=0,"Not due",IF(K2:K<=30,"1-30",IF(K2:K<=60,"31-60",IF(K2:K<=90,"61-90","90+")))))))}');
  sh.getRange('M1').setValue('Last reminder');
  sh.getRange('N1').setFormula('={"Paid late (days)";ARRAYFORMULA(IF(A2:A="",,IF((I2:I<=0)*(G2:G<>""),IF(G2:G-D2:D>0,G2:G-D2:D,0),)))}');
  const n = 499;
  header_(sh.getRange('A1:N1'));
  sh.getRange('I1:N1').setFontColor(T.SKY);
  banding_(sh.getRange('A1:N500'));
  sh.getRange(2, 1, n, 8).setFontColor(T.IVORY);
  sh.getRange(2, 9, n, 6).setFontColor('#B9C6D3');
  sh.setFrozenRows(1);
  sh.getRange(2, 3, n, 2).setNumberFormat('d mmm yyyy');
  sh.getRange(2, 7, n, 1).setNumberFormat('d mmm yyyy');
  sh.getRange(2, 5, n, 2).setNumberFormat('#,##0.00');
  sh.getRange(2, 9, n, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, 10, n, 5).setHorizontalAlignment('center');
  sh.getRange(2, 2, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(SHEET.CUSTOMERS).getRange('A2:A120'), true).setAllowInvalid(false).build());
  sh.getRange(2, 3, n, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  sh.getRange(2, 7, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  sh.getRange(2, 5, n, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build());
  const dup = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",COUNTIF($A$2:$A,$A2)>1)')
    .setBackground('#4A1F1C').setRanges([sh.getRange('A2:A500')]).build();
  const status = chips_(sh.getRange('J2:J500'), { Overdue: CHIP.Urgent, Paid: ['#1F3A33', T.GOOD], Overpaid: CHIP.Medium, 'Part paid': CHIP.Low, 'Not due': ['#22303A', '#B9C6D3'] });
  const age = chips_(sh.getRange('L2:L500'), { '1-30': HEAT['1-30'], '31-60': HEAT['31-60'], '61-90': HEAT['61-90'], '90+': HEAT['90+'] });
  sh.setConditionalFormatRules(status.concat(age, [dup]));
  [95, 185, 92, 92, 92, 92, 92, 150, 92, 88, 92, 70, 140, 100, 16].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange('I:N').protect().setDescription('Calculated columns').setWarningOnly(true);
}

function buildLog_(ss, sh, inv, day) {
  darkSheet_(sh, 300, 9);
  const hdr = ['Timestamp', 'Invoice no', 'Customer', 'Email', 'Stage', 'Mode', 'Result'];
  sh.getRange(1, 1, 1, 7).setValues([hdr]);
  const rows = DEMO.log.map(function (e) {
    const d = day(e.day); d.setHours(8, 2, 0, 0);
    return [d, e.no, e.customer, e.email, e.stage, 'Send to customers', 'Sent to customer'];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, 7).setValues(rows);
  header_(sh.getRange('A1:G1'));
  banding_(sh.getRange('A1:G300'));
  sh.getRange('A2:A').setNumberFormat('d mmm yyyy hh:mm');
  sh.setFrozenRows(1);
  [140, 100, 200, 260, 90, 140, 160, 16, 120].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setConditionalFormatRules(chips_(sh.getRange('E2:E300'), { Final: CHIP.Urgent, Second: CHIP.High, Friendly: CHIP.Medium, Upcoming: CHIP.Low }));
  sh.protect().setDescription('Written by the script').setWarningOnly(true);
  // last reminder per invoice on the Invoices sheet
  const last = {};
  DEMO.log.forEach(function (e) { last[e.no] = e.stage + ' ' + fmtDate(day(e.day)); });
  const nos = inv.getRange(2, 1, DEMO.invoices.length, 1).getValues();
  inv.getRange(2, 13, nos.length, 1).setValues(nos.map(function (r) { return [last[r[0]] || '']; }));
}

function buildReplies_(ss, sh, day) {
  darkSheet_(sh, 200, 9);
  sh.getRange('A1:F1').setValues([['Date', 'Invoice no', 'Customer', 'Type', 'Promised date', 'Note']]);
  sh.getRange('C1').setFormula('={"Customer";ARRAYFORMULA(IF(B2:B="",,IFERROR(VLOOKUP(B2:B,Invoices!A2:B,2,FALSE),"unknown invoice")))}');
  sh.getRange('G1').setFormula('={"Effect on reminders";ARRAYFORMULA(IF(B2:B="",,IF((D2:D="Promise to pay")*(E2:E<>""),' +
    'IF(E2:E>=AsOf,"Paused until "&TEXT(E2:E,"d mmm"),"Promise not kept: call"),' +
    'IF(AsOf-A2:A<=PauseDays,"Paused until "&TEXT(A2:A+PauseDays,"d mmm"),"Pause ended"))))}');
  const rows = DEMO.replies.map(function (r) { return [day(r.day), r.no, '', r.type, r.promised === null ? '' : day(r.promised), r.note]; });
  sh.getRange(2, 1, rows.length, 2).setValues(rows.map(function (r) { return r.slice(0, 2); }));
  sh.getRange(2, 4, rows.length, 3).setValues(rows.map(function (r) { return r.slice(3); }));
  header_(sh.getRange('A1:G1'));
  sh.getRange('C1').setFontColor(T.SKY); sh.getRange('G1').setFontColor(T.SKY);
  banding_(sh.getRange('A1:G200'));
  sh.getRange('A2:A200').setNumberFormat('d mmm yyyy'); sh.getRange('E2:E200').setNumberFormat('d mmm yyyy');
  sh.getRange('B2:B200').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(SHEET.INVOICES).getRange('A2:A500'), true).setAllowInvalid(false).build());
  sh.getRange('D2:D200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(REPLY_TYPES, true).build());
  sh.getRange('A2:A200').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
  sh.getRange('E2:E200').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
  sh.setFrozenRows(1);
  [100, 100, 190, 120, 115, 360, 190, 16, 120].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  const g = sh.getRange('G2:G200');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Paused').setBackground('#3A3220').setFontColor(T.AMBER).setRanges([g]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Promise not kept').setBackground('#4A1F1C').setFontColor(T.BAD).setBold(true).setRanges([g]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pause ended').setFontColor(T.MUTED).setRanges([g]).build()]);
}

/** Hidden helper: the six points of the trends (5 month ends + the as-of date), one formula per number. */
function buildCalc_(ss, sh) {
  // visible (charts in Google Sheets do not plot a hidden sheet's data), styled like the rest
  darkSheet_(sh, 20, 13);
  [70, 95, 95, 105, 105, 105, 105, 105, 80, 110, 150, 60, 16].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange('A1:L1').setValues([['Month', 'Start', 'End', 'Outstanding', 'Overdue', 'Collected', 'Billed', 'Sales 90d', 'DSO', 'Collection rate', 'Same days last month', 'Good?']]);
  const L = 'LET(n,Invoices!$A$2:$A$500,iss,Invoices!$C$2:$C$500,du,Invoices!$D$2:$D$500,amt,Invoices!$E$2:$E$500,pd,Invoices!$F$2:$F$500,' +
            'pon,IF(Invoices!$G$2:$G$500="",AsOf,Invoices!$G$2:$G$500),';
  for (let k = 5; k >= 0; k--) {
    const r = 7 - k;
    sh.getRange(r, 2).setFormula('=EOMONTH(AsOf,' + (-k - 1) + ')+1');
    sh.getRange(r, 3).setFormula(k ? '=EOMONTH(AsOf,' + (-k) + ')' : '=AsOf');
    sh.getRange(r, 1).setFormula('=TEXT(C' + r + ',"mmm")');
    const P = '$C' + r, S = '$B' + r, pb = 'paidby,pd*(pon<=' + P + '),bal,amt-paidby,';
    sh.getRange(r, 4).setFormula('=ARRAYFORMULA(' + L + pb + 'SUMPRODUCT((n<>"")*(iss<=' + P + ')*(bal>0.0000001)*bal)))');
    sh.getRange(r, 5).setFormula('=ARRAYFORMULA(' + L + pb + 'SUMPRODUCT((n<>"")*(iss<=' + P + ')*(du<' + P + ')*(bal>0.0000001)*bal)))');
    sh.getRange(r, 6).setFormula('=ARRAYFORMULA(' + L + 'SUMPRODUCT((n<>"")*(pd>0)*(pon>=' + S + ')*(pon<=' + P + ')*pd)))');
    sh.getRange(r, 7).setFormula('=ARRAYFORMULA(' + L + 'SUMPRODUCT((n<>"")*(iss>=' + S + ')*(iss<=' + P + ')*amt)))');
    sh.getRange(r, 8).setFormula('=ARRAYFORMULA(' + L + 'SUMPRODUCT((n<>"")*(iss>' + P + '-90)*(iss<=' + P + ')*amt)))');
    sh.getRange(r, 9).setFormula('=IF(H' + r + '=0,0,D' + r + '/H' + r + '*90)');
    sh.getRange(r, 10).setFormula('=ARRAYFORMULA(' + L + pb + 'da,SUMPRODUCT((n<>"")*(du>' + P + '-90)*(du<=' + P + ')*amt),' +
      'dc,SUMPRODUCT((n<>"")*(du>' + P + '-90)*(du<=' + P + ')*IF(paidby>amt,amt,paidby)),IF(da=0,0,dc/da)))');
  }
  sh.getRange('K2').setFormula('=ARRAYFORMULA(' + L + 'SUMPRODUCT((n<>"")*(pd>0)*(pon>=EOMONTH(AsOf,-2)+1)*(pon<=EDATE(AsOf,-1))*pd)))');
  // 1 = the KPI moved the right way since last month (drives the green / red note colour)
  sh.getRange('L2:L6').setFormulas([['=--(D7<=D6)'], ['=--(E7<=E6)'], ['=--(F7>=K2)'], ['=--(I7<=I6)'], ['=--(J7>=J6)']]);
  sh.getRange('A10:B13').setValues([['Where the money is', 'Amount'], ['Collected', ''], ['Overdue', ''], ['Not due yet', '']]);
  sh.getRange('B11').setFormula('=SUM(Invoices!F2:F500)');
  sh.getRange('B12').setFormula('=SUMIFS(Invoices!I2:I500,Invoices!J2:J500,"Overdue")');
  sh.getRange('B13').setFormula('=SUMIF(Invoices!I2:I500,">0")-B12');
  sh.getRange('B2:C7').setNumberFormat('d mmm yyyy');
  sh.getRange('D2:H7').setNumberFormat('#,##0');
  sh.getRange('B11:B13').setNumberFormat('#,##0');
  header_(sh.getRange('A1:L1')); header_(sh.getRange('A10:B10'));
  sh.getRange('I2:I7').setNumberFormat('0.0'); sh.getRange('J2:J7').setNumberFormat('0.0%'); sh.getRange('K2').setNumberFormat('#,##0');
  sh.getRange('A15').setValue('Helper for the Dashboard trends and charts: every number is a formula on Invoices. Checked by Run self-test.')
    .setFontColor(T.MUTED).setFontStyle('italic');
  ss.setNamedRange('CalcTrend', sh.getRange('D2:J7'));
  ss.setNamedRange('CalcPrevMtd', sh.getRange('K2'));
}

function buildDashboard_(ss, sh, calc) {
  darkSheet_(sh, 58, 12);
  sh.setColumnWidth(1, 18); sh.setColumnWidth(12, 18);
  for (let c = 2; c <= 11; c++) sh.setColumnWidth(c, 128);
  const set = function (a1, v, fmt) { const r = sh.getRange(a1); (String(v).charAt(0) === '=' ? r.setFormula(v) : r.setValue(v)); if (fmt) r.setNumberFormat(fmt); return r; };
  const hr = function (row, h) { sh.setRowHeight(row, h); };

  // header
  sh.getRange('A1:L2').setBackground(T.BG);
  set('B1', 'N O R T H B E A M   S T U D I O').setFontColor(T.GOLD).setFontSize(8).setFontWeight('bold');
  sh.getRange('B2:F2').merge();
  set('B2', 'Accounts receivable').setFontFamily('Georgia').setFontSize(22).setFontColor(T.IVORY);
  sh.getRange('G2:K2').merge();
  set('G2', '="Fictional data   ·   as of "&TEXT(AsOf,"d mmm yyyy")&"   ·   self-test: "&Checks!C3')
    .setHorizontalAlignment('right').setFontColor(T.MUTED).setFontSize(9);
  sh.getRange('A3:L3').setBackground(T.GOLD);
  hr(1, 22); hr(2, 44); hr(3, 2); hr(4, 10); hr(5, 28); hr(6, 12);

  // KPI cards (rows 7-10), 2 columns each
  const cards = [
    ['B', '◆  OUTSTANDING', '=SUMIF(Invoices!I2:I,">0")', '#,##0.00', '=SPARKLINE(Calc!D2:D7,{"charttype","line";"color","' + T.GOLD + '";"linewidth",2})',
      '=IF(Calc!L2,"▼ ","▲ ")&TEXT(ABS(Calc!D7-Calc!D6),"#,##0")&" vs "&TEXT(Calc!C6,"d mmm")', 'L2', 'DashOutstanding'],
    ['D', '▲  OVERDUE', '=SUMIFS(Invoices!I2:I,Invoices!J2:J,"Overdue")', '#,##0.00', '=SPARKLINE(Calc!E2:E7,{"charttype","line";"color","' + T.BAD + '";"linewidth",2})',
      '=IF(Calc!L3,"▼ ","▲ ")&TEXT(ABS(Calc!E7-Calc!E6),"#,##0")&" vs "&TEXT(Calc!C6,"d mmm")&"  ·  "&TEXT(IFERROR(D8/B8,0),"0%")&" of open"', 'L3', 'DashOverdue'],
    ['F', '✚  COLLECTED THIS MONTH', '=Calc!F7', '#,##0.00', '=SPARKLINE(Calc!F2:F7,{"charttype","column";"color","' + T.JADE + '";"highcolor","' + T.GOLD + '"})',
      '=IF(Calc!L4,"▲ ","▼ ")&TEXT(ABS(Calc!F7-Calc!K2),"#,##0")&" vs same days last month"', 'L4', 'DashCollected'],
    ['H', '◷  DSO (DAYS)', '=Calc!I7', '0.0', '=SPARKLINE(Calc!I2:I7,{"charttype","line";"color","' + T.SKY + '";"linewidth",2})',
      '=IF(Calc!L5,"▼ ","▲ ")&TEXT(ABS(Calc!I7-Calc!I6),"0.0")&" days vs "&TEXT(Calc!C6,"d mmm")', 'L5', 'DashDSO'],
    ['J', '◎  COLLECTION RATE', '=Calc!J7', '0.0%', '=SPARKLINE(Calc!J2:J7,{"charttype","line";"color","' + T.GOOD + '";"linewidth",2})',
      '=IF(Calc!L6,"▲ ","▼ ")&TEXT(ABS(Calc!J7-Calc!J6)*100,"0.0")&" pts vs "&TEXT(Calc!C6,"d mmm")', 'L6', 'DashRate']];
  const rules = [];
  cards.forEach(function (c) {
    const col = sh.getRange(c[0] + '7').getColumn();
    const box = sh.getRange(7, col, 4, 2);
    box.setBackground(T.CARD);
    for (let r = 7; r <= 10; r++) sh.getRange(r, col, 1, 2).merge();
    sh.getRange(7, col).setValue(c[1]).setFontSize(8).setFontWeight('bold').setFontColor(T.GOLD);
    sh.getRange(8, col).setFormula(c[2]).setNumberFormat(c[3]).setFontFamily('Georgia').setFontSize(22).setFontColor(T.IVORY);
    sh.getRange(9, col).setFormula(c[4]);
    const note = sh.getRange(10, col).setFormula(c[5]).setFontSize(9).setFontWeight('bold');
    box.setHorizontalAlignment('left');
    sh.getRange(7, col, 4, 1).setBorder(null, true, null, null, null, null, T.GOLD, SpreadsheetApp.BorderStyle.SOLID_THICK);
    sh.getRange(7, col + 1, 4, 1).setBorder(null, null, null, true, null, null, T.PAGE, SpreadsheetApp.BorderStyle.SOLID_THICK);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=INDIRECT("Calc!' + c[6] + '")=1').setFontColor(T.GOOD).setRanges([note]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=INDIRECT("Calc!' + c[6] + '")=0').setFontColor(T.BAD).setRanges([note]).build());
    ss.setNamedRange(c[7], sh.getRange(8, col));
  });
  sh.getRange('D8').setFontColor(T.BAD);
  hr(7, 26); hr(8, 44); hr(9, 34); hr(10, 24); hr(11, 14);
  sh.getRange('B7:K7').setVerticalAlignment('bottom');

  // charts (rows 12-25) on cards
  const section = function (row, c1, c2, title, rowsDown) {
    sh.getRange(row, c1, rowsDown, c2 - c1 + 1).setBackground(T.CARD);
    sh.getRange(row, c1).setValue(title).setFontFamily('Georgia').setFontSize(13).setFontColor(T.IVORY);
    sh.getRange(row, c1, 1, c2 - c1 + 1).setBorder(true, null, null, null, null, null, T.GOLD, SpreadsheetApp.BorderStyle.SOLID);
  };
  section(12, 2, 5, 'Where the money is', 14);
  section(12, 6, 11, 'Collections, last 6 months', 14);
  sh.getRange(12, 5, 14, 1).setBorder(null, null, null, true, null, null, T.PAGE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  set('H12', 'gold = cash collected, grey = billed').setFontColor(T.MUTED).setFontSize(9);
  hr(12, 30);
  for (let r = 13; r <= 25; r++) hr(r, 21);
  hr(26, 14);
  const donut = chartStyle_(sh.newChart().asPieChart().addRange(calc.getRange('A10:B13')).setNumHeaders(1))
    .setOption('pieHole', 0.58).setOption('colors', [T.JADE, T.BAD, T.SKY]).setOption('pieSliceBorderColor', T.CARD)
    .setOption('legend', { position: 'right', textStyle: { color: T.IVORY } }).setOption('pieSliceText', 'percentage')
    .setOption('pieSliceTextStyle', { color: T.BG })
    .setOption('width', 500).setOption('height', 268).setPosition(13, 2, 4, 2).build();
  sh.insertChart(donut);
  const trend = chartStyle_(sh.newChart().asColumnChart().addRange(calc.getRange('A1:A7')).addRange(calc.getRange('F1:G7')).setNumHeaders(1))
    .setOption('colors', [T.GOLD, T.GREY]).setOption('legend', { position: 'none' })
    .setOption('width', 756).setOption('height', 268).setPosition(13, 6, 4, 2).build();
  sh.insertChart(trend);

  // owner action queue (rows 27-38), written by refreshDesk_()
  section(27, 2, 11, 'Owner action queue: today', 12);
  sh.getRange('H27:K27').merge().setHorizontalAlignment('right').setFontColor(T.GOLD_LT).setFontSize(10);
  hr(27, 30);
  sh.getRange('B28:K28').setValues([['Priority', 'Invoice', 'Customer', '', 'Action', '', 'Why', '', 'Days late', 'Balance']]);
  header_(sh.getRange('B28:K28'));
  for (let r = 28; r <= 38; r++) {
    sh.getRange(r, 4, 1, 2).merge(); sh.getRange(r, 6, 1, 2).merge(); sh.getRange(r, 8, 1, 2).merge();
    if (r > 28) hr(r, 22);
  }
  sh.getRange('B29:K38').setBackground(T.CARD).setBorder(null, null, true, null, null, true, T.LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('B29:B38').setHorizontalAlignment('center');
  sh.getRange('C29:C38').setFontColor(T.GOLD_LT);
  sh.getRange('H29:I38').setFontColor(T.MUTED).setFontSize(9);
  sh.getRange('J29:J38').setHorizontalAlignment('center');
  sh.getRange('K29:K38').setNumberFormat('#,##0.00');
  rules.push.apply(rules, chips_(sh.getRange('B29:B38'), CHIP));
  hr(39, 14);

  // aging heat table (rows 40-47) and customer risk (rows 40-47)
  section(40, 2, 5, 'Aging: how late the money is', 8);
  section(40, 6, 11, 'Customer risk (payment habit, lateness, overdue)', 8);
  sh.getRange(40, 5, 8, 1).setBorder(null, null, null, true, null, null, T.PAGE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  hr(40, 30);
  sh.getRange('B41:E41').setValues([['Days late', 'Owed', 'Share', '']]);
  header_(sh.getRange('B41:E41'));
  AGING.forEach(function (b, i) {
    const r = 42 + i;
    sh.getRange('B' + r).setNumberFormat('@').setValue(b).setBackground(HEAT[b][0]).setFontColor(HEAT[b][1]).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange('C' + r).setFormula('=SUMIFS(Invoices!I2:I,Invoices!L2:L,B' + r + ',Invoices!I2:I,">0")');
    sh.getRange('D' + r).setFormula('=IFERROR(C' + r + '/$B$8,0)');
    sh.getRange('E' + r).setFormula('=SPARKLINE(C' + r + ',{"charttype","bar";"max",MAX($C$42:$C$46);"color1","' + HEAT[b][2] + '"})');
    hr(r, 24);
  });
  sh.getRange('B47').setValue('Total').setFontWeight('bold').setHorizontalAlignment('center').setFontColor(T.GOLD_LT);
  sh.getRange('C47').setFormula('=SUM(C42:C46)').setFontWeight('bold').setFontColor(T.GOLD_LT);
  sh.getRange('D47').setFormula('=SUM(D42:D46)').setFontWeight('bold').setFontColor(T.GOLD_LT);
  sh.getRange('B47:E47').setBorder(true, null, null, null, null, null, T.GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('C42:C47').setNumberFormat('#,##0.00'); sh.getRange('D42:D47').setNumberFormat('0%');
  ss.setNamedRange('DashAging', sh.getRange('B42:C46'));
  sh.getRange('F41:K41').setValues([['Customer', '', 'Risk', 'Score', 'Overdue', 'Avg days late']]);
  header_(sh.getRange('F41:K41'));
  for (let r = 41; r <= 47; r++) sh.getRange(r, 6, 1, 2).merge();
  sh.getRange('H42:K47').setHorizontalAlignment('center');
  sh.getRange('J42:J47').setNumberFormat('#,##0.00');
  sh.getRange('I42:I47').setFontFamily('Georgia').setFontSize(12);
  rules.push.apply(rules, chips_(sh.getRange('H42:H47'), { High: CHIP.Urgent, Medium: CHIP.Medium, Low: ['#1F3A33', T.GOOD] }));
  hr(47, 24); hr(48, 14);

  // data checks (rows 49-55)
  section(49, 2, 11, 'Data checks', 7);
  hr(49, 30);
  const checks = [['Every invoice number is unique', '=IF(COUNTA(Invoices!A2:A)=COUNTUNIQUE(Invoices!A2:A),"PASS","FAIL")'],
    ['Every invoice\'s customer is on Customers', '=IF(SUMPRODUCT((Invoices!B2:B<>"")*(COUNTIF(Customers!A2:A,Invoices!B2:B)=0))=0,"PASS","FAIL")'],
    ['Aging table adds up to outstanding', '=IF(ROUND(SUM(C42:C46)-B8,2)=0,"PASS","FAIL")'],
    ['Script self-test (menu: Run self-test)', '=Checks!C3'],
    ['Overpaid invoices: issue a credit note', '=COUNTIF(Invoices!J2:J,"Overpaid")'],
    ['Customers with no email: reminders are skipped', '=COUNTIFS(Customers!A2:A,"<>",Customers!C2:C,"")']];
  checks.forEach(function (c, i) {
    sh.getRange(50 + i, 2).setValue(c[0]);
    sh.getRange(50 + i, 10, 1, 2).merge().setFormula(c[1]).setFontWeight('bold').setHorizontalAlignment('center');
    hr(50 + i, 22);
  });
  const res = sh.getRange('J50:K55');
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('PASS').setBackground('#1F3A33').setFontColor(T.GOOD).setRanges([res]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('FAIL').setBackground('#4A1F1C').setFontColor(T.BAD).setRanges([res]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#3A3220').setFontColor(T.AMBER).setRanges([res]).build());
  sh.setConditionalFormatRules(rules);
  set('B57', '').setFontColor(T.MUTED).setFontSize(8);
}

function buildStatement_(ss, sh) {
  const PAPER = '#FBF8F2', INK = '#1A2129', RULE = '#E2D9C6';
  if (sh.getMaxRows() > 40) sh.deleteRows(41, sh.getMaxRows() - 40);
  if (sh.getMaxColumns() > 12) sh.deleteColumns(13, sh.getMaxColumns() - 12);
  sh.setHiddenGridlines(true);
  sh.getRange('A1:L40').setBackground(T.PAGE).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sh.getRange('A4:I38').setBackground(PAPER).setFontColor(INK);
  [24, 105, 100, 100, 110, 100, 115, 85, 24, 20, 150, 150].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange('A1:I2').setBackground(T.BG);
  sh.getRange('B1').setValue('N O R T H B E A M   S T U D I O').setFontColor(T.GOLD).setFontSize(8).setFontWeight('bold');
  sh.getRange('B2').setValue('Statement of account').setFontFamily('Georgia').setFontSize(20).setFontColor(T.IVORY);
  sh.getRange('E2:H2').merge().setValue('accounts@northbeam.example.com')
    .setFontColor(T.MUTED).setFontSize(8).setHorizontalAlignment('right');
  sh.getRange('A3:I3').setBackground(T.GOLD);
  sh.setRowHeight(1, 22); sh.setRowHeight(2, 44); sh.setRowHeight(3, 3); sh.setRowHeight(4, 14);
  sh.getRange('B5').setValue('BILL TO').setFontSize(8).setFontWeight('bold').setFontColor(T.COPPER);
  sh.getRange('B6:E6').merge().setValue('Inkwell Bookshop').setFontFamily('Georgia').setFontSize(14).setFontWeight('bold')
    .setBackground('#F3ECDD').setBorder(true, true, true, true, null, null, T.GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('B6').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(SHEET.CUSTOMERS).getRange('A2:A120'), true).build());
  sh.getRange('B7:E7').merge().setFormula('=IFERROR(VLOOKUP(B6,Customers!A2:C120,2,FALSE),"")').setFontColor('#4A5560');
  sh.getRange('B8:E8').merge().setFormula('=IFERROR(VLOOKUP(B6,Customers!A2:C120,3,FALSE),"")').setFontColor('#4A5560');
  sh.getRange('F5:H5').merge().setValue('STATEMENT DATE').setFontSize(8).setFontWeight('bold').setFontColor(T.COPPER).setHorizontalAlignment('right');
  sh.getRange('F6:H6').merge().setFormula('=AsOf').setNumberFormat('d mmmm yyyy').setHorizontalAlignment('right').setFontSize(11);
  sh.getRange('F7:H7').merge().setValue('TOTAL DUE').setFontSize(8).setFontWeight('bold').setFontColor(T.COPPER).setHorizontalAlignment('right');
  sh.getRange('F8:H8').merge().setFormula('=Settings!B5&TEXT(SUMIFS(Invoices!I2:I500,Invoices!B2:B500,B6,Invoices!I2:I500,">0"),"#,##0.00")')
    .setFontFamily('Georgia').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('right');
  sh.setRowHeight(6, 30); sh.setRowHeight(8, 32); sh.setRowHeight(9, 12);
  sh.getRange('B10:H10').setValues([['Invoice', 'Issued', 'Due', 'Amount', 'Paid', 'Balance', 'Days late']])
    .setBackground(INK).setFontColor(T.GOLD_LT).setFontWeight('bold');
  sh.getRange('C10:H10').setHorizontalAlignment('right');
  sh.getRange('B11').setFormula('=IFERROR(SORT(CHOOSECOLS(FILTER(Invoices!A2:K500,Invoices!B2:B500=B6,Invoices!I2:I500>0),1,3,4,5,6,9,11),3,TRUE),"No open invoices")');
  sh.getRange('B11:H28').setBorder(null, null, true, null, null, true, RULE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('C11:D28').setNumberFormat('d mmm yyyy');
  sh.getRange('E11:G28').setNumberFormat('#,##0.00');
  sh.getRange('G11:G28').setFontWeight('bold');
  sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor('#B23A2A').setBold(true)
    .setRanges([sh.getRange('H11:H28')]).build()]);
  sh.getRange('B30').setValue('AGING OF THE BALANCE').setFontSize(8).setFontWeight('bold').setFontColor(T.COPPER);
  sh.getRange('B31:F31').setNumberFormat('@').setValues([['Not due', '1-30', '31-60', '61-90', '90+']]).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#F3ECDD');
  ['B', 'C', 'D', 'E', 'F'].forEach(function (c) {
    sh.getRange(c + '32').setFormula('=SUMIFS(Invoices!I2:I500,Invoices!B2:B500,$B$6,Invoices!L2:L500,' + c + '31,Invoices!I2:I500,">0")');
  });
  sh.getRange('B32:F32').setNumberFormat('#,##0.00').setHorizontalAlignment('center');
  sh.getRange('B31:F32').setBorder(true, true, true, true, true, true, RULE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('B34').setValue('Please pay by bank transfer to Northbeam Studio, account 00-0000-0000 (fictional), quoting the invoice number.').setFontColor('#4A5560');
  sh.getRange('B35').setValue('Already paid? Thank you, please ignore this statement. Questions: reply to our reminder email.').setFontColor('#4A5560');
  sh.getRange('B37').setValue('Thank you for your business.').setFontFamily('Georgia').setFontStyle('italic').setFontColor(T.COPPER);
  sh.getRange('A38:I38').setBackground(T.GOLD); sh.setRowHeight(38, 3);
  sh.getRange('K5').setValue('PDF of this statement').setFontColor(T.GOLD).setFontWeight('bold').setFontSize(9);
  sh.getRange('K6:L6').merge().setFormula('=HYPERLINK("' + statementUrl_() + '","⤓  Download PDF")')
    .setBackground(T.CARD).setFontColor(T.GOLD_LT).setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, null, null, T.GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('K7:L9').merge().setValue('Pick the customer in the gold box; the link exports the paper area (A1:I38) as an A4 PDF. ' +
    'The sidebar and the reminder preview prepare it per customer.').setWrap(true).setFontColor(T.MUTED).setFontSize(8).setVerticalAlignment('top');
}

function buildChecks_(sh) {
  darkSheet_(sh, 120, 6);
  titleBlock_(sh, 'Checks', '', 6);
  sh.getRange('B2').setValue('Checks   ·   written by Invoice tracker ▸ Run self-test');
  sh.getRange('A1:A2').setBackground(T.BG);
  sh.getRange('A3:F3').setBackground(T.PAGE);
  sh.setRowHeight(3, 34);
  sh.getRange('A3').setValue('Last run').setFontColor(T.MUTED);
  sh.getRange('B3').setNumberFormat('d mmm yyyy hh:mm').setFontColor(T.MUTED).setHorizontalAlignment('left');
  sh.getRange('C3').setValue('not run yet').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sh.getRange('A4:D4').setValues([['Type', 'Check', 'Result', 'Detail']]);
  header_(sh.getRange('A4:D4'));
  banding_(sh.getRange('A4:D120'));
  sh.getRange('A1:F2').setBackground(T.BG);
  sh.getRange('A3:F3').setBorder(true, null, null, null, null, null, T.GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sh.setColumnWidth(1, 120); sh.setColumnWidth(2, 600); sh.setColumnWidth(3, 120); sh.setColumnWidth(4, 360); sh.setColumnWidth(5, 16); sh.setColumnWidth(6, 120);
  sh.getRange('C5:C120').setHorizontalAlignment('center');
  sh.getRange('D5:D120').setFontColor(T.MUTED).setFontSize(9);
  sh.getRange('A5:A120').setFontColor(T.GOLD_LT);
  const rng = sh.getRange('C3:C120');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('PASS').setBackground('#1F3A33').setFontColor(T.GOOD).setBold(true).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('FAIL').setBackground('#4A1F1C').setFontColor(T.BAD).setBold(true).setRanges([rng]).build()]);
  sh.setFrozenRows(4);
}

/**
 * Navigation buttons: on the Dashboard a bar of all sheets (the current one in gold), on every other sheet a
 * "Dashboard" button in the first free column. HYPERLINK("#gid=...") jumps to the sheet inside the same file.
 */
function buildNav_(ss, sh) {
  const names = [SHEET.DASHBOARD, SHEET.INVOICES, SHEET.CUSTOMERS, SHEET.REPLIES, SHEET.STATEMENT, SHEET.LOG, SHEET.SETTINGS, SHEET.CHECKS];
  const link = function (n, label) { return '=HYPERLINK("#gid=' + sh[n].getSheetId() + '","' + (label || n) + '")'; };
  const btn = function (r, current) {
    r.setBackground(current ? T.GOLD : T.CARD).setFontColor(current ? T.BG : T.IVORY).setFontWeight(current ? 'bold' : 'normal')
      .setHorizontalAlignment('center').setFontSize(10)
      .setBorder(true, true, true, true, null, null, current ? T.GOLD : '#5A4A2E', SpreadsheetApp.BorderStyle.SOLID);
  };
  const dash = sh[SHEET.DASHBOARD];
  names.forEach(function (n, i) { const r = dash.getRange(5, 2 + i).setFormula(link(n)); btn(r, n === SHEET.DASHBOARD); });
  dash.getRange('J5:K5').merge().setValue('Menu ▸ Invoice tracker ▸ Open collections desk').setFontColor(T.MUTED).setFontSize(8)
    .setFontStyle('italic').setHorizontalAlignment('right');
  const back = { Invoices: 'O1', Customers: 'L1', Replies: 'I1', 'Reminder log': 'I1', Settings: 'F1', Checks: 'F1', Statement: 'K2' };
  Object.keys(back).forEach(function (n) {
    const r = sh[n].getRange(back[n]).setFormula(link(SHEET.DASHBOARD, '◀  Dashboard'));
    btn(r, false);
  });
  sh[SHEET.INVOICES].setColumnWidth(15, 110);
}
