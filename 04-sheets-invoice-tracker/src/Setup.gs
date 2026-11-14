/**
 * Builds the whole tracker in the current spreadsheet: Dashboard, Invoices, Customers, Reminder log, Settings, Checks,
 * with formulas, drop-downs, colours, a chart and the fictional demo data from DemoData.gs.
 * Run "Rebuild demo data" from the menu (or buildDemo from the editor) in an empty Google Sheet.
 */

const NAVY = '#182d48', TEAL = '#008080', LIGHT = '#f7f9fc', YELLOW = '#fff2cc';

function menuRebuildDemo() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Rebuild the demo?', 'This replaces every sheet in this file with fresh demo data.', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  buildDemo();
  ui.alert('Demo rebuilt. Try Invoice tracker > Preview reminders.');
}

function buildDemo() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetLocale('en_US');
  ss.setSpreadsheetTimeZone(Session.getScriptTimeZone());
  const temp = ss.insertSheet('tmp' + Date.now());
  ss.getSheets().forEach(function (s) { if (s.getSheetId() !== temp.getSheetId()) ss.deleteSheet(s); });
  ss.getNamedRanges().forEach(function (n) { n.remove(); });
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const day = function (offset) { const d = new Date(today); d.setDate(d.getDate() + offset); return d; };

  const dash = ss.insertSheet(SHEET.DASHBOARD);
  const inv = ss.insertSheet(SHEET.INVOICES);
  const cust = ss.insertSheet(SHEET.CUSTOMERS);
  const log = ss.insertSheet(SHEET.LOG);
  const set = ss.insertSheet(SHEET.SETTINGS);
  const chk = ss.insertSheet(SHEET.CHECKS);
  ss.deleteSheet(temp);

  buildSettings_(ss, set);
  buildCustomers_(ss, cust);
  buildInvoices_(ss, inv, day);
  buildLog_(ss, log, inv, day);
  buildDashboard_(ss, dash);
  buildChecks_(chk);
  SpreadsheetApp.flush();
  selfTest();
  dash.activate();
}

function header_(range) {
  range.setFontWeight('bold').setFontColor('white').setBackground(NAVY).setVerticalAlignment('middle');
}

function buildSettings_(ss, sh) {
  const rows = [['Setting', 'Value', 'Note'],
    ['Company', 'Northbeam Studio', 'Fictional design agency used for the demo'],
    ['Sender name', 'Northbeam Studio Accounts', 'Name shown on reminder emails'],
    ['Owner email', '', 'Summary and test emails go here. Blank = the account running the script'],
    ['Currency symbol', '$', ''],
    ['Email mode', 'Preview only', 'Preview only: nothing is sent. Send to me (test): every reminder comes to you. Send to customers: live.'],
    ['Max emails per run', 25, 'Safety limit; the rest wait for the next run'],
    ['Summary day', 'Monday', 'Weekday the summary email goes out (daily automation)'],
    ['As-of date', '', 'Blank = today. Type a date to see the tracker as on that day'],
    ['As-of used', '=IF(ISDATE(B9),B9,TODAY())', 'What every formula and the script use']];
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  header_(sh.getRange('A1:C1'));
  sh.getRange('B2:B9').setBackground(YELLOW);
  sh.getRange('B6').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(MODES, true).build());
  sh.getRange('B8').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'], true).build());
  sh.getRange('B9:B10').setNumberFormat('d mmm yyyy');
  ss.setNamedRange('AsOf', sh.getRange('B10'));

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
  sh.getRange('A14').setValue('Reminder stages - placeholders: {contact} {customer} {invoice} {amount} {due} {days} {company} {sender}').setFontWeight('bold').setFontColor(NAVY);
  sh.getRange(15, 1, st.length, 4).setValues(st);
  header_(sh.getRange('A15:D15'));
  sh.getRange('A16:D19').setBackground(YELLOW).setWrap(true).setVerticalAlignment('top');
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 190); sh.setColumnWidth(3, 330); sh.setColumnWidth(4, 520);
  sh.getRange('C2:C10').setFontColor('#777777');
}

function buildCustomers_(ss, sh) {
  const rows = [['Customer', 'Contact', 'Email', 'Open balance', 'Overdue']];
  DEMO.customers.forEach(function (c) { rows.push([c.name, c.contact, c.email, '', '']); });
  sh.getRange(1, 1, rows.length, 5).setValues(rows);
  // BYROW, because SUMIFS inside ARRAYFORMULA does not run once per row
  sh.getRange('D2').setFormula('=BYROW(A2:A,LAMBDA(c,IF(c="",,SUMIFS(Invoices!I2:I,Invoices!B2:B,c,Invoices!I2:I,">0"))))');
  sh.getRange('E2').setFormula('=BYROW(A2:A,LAMBDA(c,IF(c="",,SUMIFS(Invoices!I2:I,Invoices!B2:B,c,Invoices!J2:J,"Overdue"))))');
  header_(sh.getRange('A1:E1'));
  sh.getRange('A2:C200').setBackground(YELLOW);
  sh.getRange('D2:E200').setNumberFormat('#,##0.00');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 140); sh.setColumnWidth(3, 260);
  const blank = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",$C2="")')
    .setBackground('#ffc7ce').setRanges([sh.getRange('C2:C200')]).build();
  sh.setConditionalFormatRules([blank]);
}

function buildInvoices_(ss, sh, day) {
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
  header_(sh.getRange('A1:M1'));
  sh.getRange('I1:L1').setBackground('#2f4a6d');
  sh.setFrozenRows(1);
  const n = 400;
  sh.getRange(2, 1, n, 8).setBackground(YELLOW);
  sh.getRange(2, 3, n, 2).setNumberFormat('d mmm yyyy');
  sh.getRange(2, 7, n, 1).setNumberFormat('d mmm yyyy');
  sh.getRange(2, 5, n, 2).setNumberFormat('#,##0.00');
  sh.getRange(2, 9, n, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, 2, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(SHEET.CUSTOMERS).getRange('A2:A200'), true).setAllowInvalid(false).build());
  sh.getRange(2, 3, n, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  sh.getRange(2, 7, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  sh.getRange(2, 5, n, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build());
  const st = sh.getRange('J2:J' + (n + 1));
  const rule = function (text, bg, fg) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(bg).setFontColor(fg).setRanges([st]).build();
  };
  const dup = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",COUNTIF($A$2:$A,$A2)>1)')
    .setBackground('#ffc7ce').setRanges([sh.getRange('A2:A' + (n + 1))]).build();
  sh.setConditionalFormatRules([rule('Overdue', '#ffc7ce', '#9c0006'), rule('Paid', '#c6efce', '#006100'),
    rule('Overpaid', '#fff2cc', '#9c5700'), rule('Part paid', '#ddebf7', '#1f4e79'), dup]);
  [110, 200, 95, 95, 95, 95, 95, 170, 95, 85, 95, 70, 150].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange('I:L').protect().setDescription('Calculated columns').setWarningOnly(true);
}

function buildLog_(ss, sh, inv, day) {
  const hdr = ['Timestamp', 'Invoice no', 'Customer', 'Email', 'Stage', 'Mode', 'Result'];
  sh.getRange(1, 1, 1, 7).setValues([hdr]);
  header_(sh.getRange('A1:G1'));
  const rows = DEMO.log.map(function (e) {
    const d = day(e.day); d.setHours(8, 2, 0, 0);
    return [d, e.no, e.customer, e.email, e.stage, 'Send to customers', 'Sent to customer'];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, 7).setValues(rows);
  sh.getRange('A2:A').setNumberFormat('d mmm yyyy hh:mm');
  sh.setFrozenRows(1);
  [140, 100, 200, 260, 90, 140, 160].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.protect().setDescription('Written by the script').setWarningOnly(true);
  // last reminder per invoice on the Invoices sheet
  const last = {};
  DEMO.log.forEach(function (e) { last[e.no] = e.stage + ' ' + fmtDate(day(e.day)); });
  const nos = inv.getRange(2, 1, DEMO.invoices.length, 1).getValues();
  inv.getRange(2, 13, nos.length, 1).setValues(nos.map(function (r) { return [last[r[0]] || '']; }));
}

function buildDashboard_(ss, sh) {
  sh.setHiddenGridlines(true);
  [20, 150, 170, 100, 150, 130, 145, 175, 130, 20, 150, 130].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange('A1:L2').setBackground(NAVY);
  sh.getRange('B1').setFormula('="Money owed   |   "&Settings!B2&"   |   as of "&TEXT(AsOf,"d mmm yyyy")')
    .setFontSize(18).setFontWeight('bold').setFontColor('white');
  sh.getRange('B2').setValue('Invoices and payments are typed on Invoices; everything here updates by itself. Menu: Invoice tracker.')
    .setFontColor('#c8d7e6');
  sh.setRowHeight(1, 40);

  const tiles = [['B', 'OUTSTANDING', '=SUMIF(Invoices!I2:I,">0")', '#,##0.00', '=COUNTIF(Invoices!I2:I,">0")&" open invoices"', 'DashOutstanding'],
    ['E', 'OVERDUE', '=SUMIFS(Invoices!I2:I,Invoices!J2:J,"Overdue")', '#,##0.00', '=COUNTIF(Invoices!J2:J,"Overdue")&" invoices · "&TEXT(IFERROR(E5/B5,0),"0%")&" of outstanding"', 'DashOverdue'],
    ['H', 'COLLECTED, LAST 30 DAYS', '=SUMIFS(Invoices!F2:F,Invoices!G2:G,">="&(AsOf-30),Invoices!G2:G,"<="&AsOf)', '#,##0.00', '=COUNTIFS(Invoices!G2:G,">="&(AsOf-30),Invoices!G2:G,"<="&AsOf)&" payments"', null],
    ['K', 'OVERDUE 60+ DAYS', '=SUMIFS(Invoices!I2:I,Invoices!K2:K,">60")', '#,##0.00', '=COUNTIFS(Invoices!K2:K,">60")&" invoices: call, do not email"', null]];
  tiles.forEach(function (t) {
    const c = sh.getRange(t[0] + '4').getColumn();
    sh.getRange(4, c, 1, 2).merge().setValue(t[1]).setFontSize(9).setFontWeight('bold').setFontColor('#5a6473').setBackground(LIGHT);
    sh.getRange(5, c, 1, 2).merge().setFormula(t[2]).setFontSize(22).setFontWeight('bold').setFontColor(NAVY).setNumberFormat(t[3]).setBackground(LIGHT);
    sh.getRange(6, c, 1, 2).merge().setFormula(t[4]).setFontSize(9).setFontColor('#6e6e6e').setBackground(LIGHT);
    sh.getRange(3, c, 1, 2).setBackground(TEAL);
    sh.getRange(3, c, 4, 2).setHorizontalAlignment('center');
    if (t[5]) ss.setNamedRange(t[5], sh.getRange(5, c));
  });
  sh.setRowHeight(3, 5); sh.setRowHeight(5, 44);
  sh.getRange('E5').setFontColor('#b00020');
  sh.getRange('B8').setValue('Owed by age').setFontSize(13).setFontWeight('bold').setFontColor(NAVY);
  sh.getRange('B9:E9').setValues([['Age (days late)', 'Owed', 'Share', '']]);
  header_(sh.getRange('B9:E9'));
  AGING.forEach(function (b, i) {
    const r = 10 + i;
    sh.getRange('B' + r).setNumberFormat('@').setValue(b);   // text, or Sheets turns "1-30" into 30 Jan
    sh.getRange('C' + r).setFormula('=SUMIFS(Invoices!I2:I,Invoices!L2:L,B' + r + ',Invoices!I2:I,">0")');
    sh.getRange('D' + r).setFormula('=IFERROR(C' + r + '/$B$5,0)');
    sh.getRange('E' + r).setFormula('=SPARKLINE(C' + r + ',{"charttype","bar";"max",MAX($C$10:$C$14);"color1","' + (i >= 3 ? '#b00020' : TEAL) + '"})');
  });
  sh.getRange('C10:C14').setNumberFormat('#,##0.00'); sh.getRange('D10:D14').setNumberFormat('0%');
  ss.setNamedRange('DashAging', sh.getRange('B10:C14'));

  sh.getRange('H8').setValue('Who owes the most').setFontSize(13).setFontWeight('bold').setFontColor(NAVY);
  sh.getRange('H9').setFormula('=QUERY(Invoices!A1:L,"select B, sum(I) where I > 0 group by B order by sum(I) desc limit 5 label B \'Customer\', sum(I) \'Owes\'",1)');
  header_(sh.getRange('H9:I9'));
  sh.getRange('I10:I14').setNumberFormat('#,##0.00');

  sh.getRange('B17').setValue('Chase first: overdue invoices, oldest first').setFontSize(13).setFontWeight('bold').setFontColor(NAVY);
  sh.getRange('B18').setFormula('=QUERY(Invoices!A1:M,"select A, B, D, I, K, M where J = \'Overdue\' order by K desc limit 12 ' +
    'label A \'Invoice\', B \'Customer\', D \'Due\', I \'Balance\', K \'Days late\', M \'Last reminder\'",1)');
  header_(sh.getRange('B18:G18'));
  sh.getRange('D19:D30').setNumberFormat('d mmm yyyy'); sh.getRange('E19:E30').setNumberFormat('#,##0.00');

  sh.getRange('B33').setValue('Data checks').setFontSize(13).setFontWeight('bold').setFontColor(NAVY);
  const checks = [['Every invoice number is unique', '=IF(COUNTA(Invoices!A2:A)=COUNTUNIQUE(Invoices!A2:A),"PASS","FAIL")'],
    ['Every invoice\'s customer is on Customers', '=IF(SUMPRODUCT((Invoices!B2:B<>"")*(COUNTIF(Customers!A2:A,Invoices!B2:B)=0))=0,"PASS","FAIL")'],
    ['Aging table adds up to outstanding', '=IF(ROUND(SUM(C10:C14)-B5,2)=0,"PASS","FAIL")'],
    ['Script self-test (menu: Run self-test)', '=Checks!C2'],
    ['Overpaid invoices: issue a credit note', '=COUNTIF(Invoices!J2:J,"Overpaid")'],
    ['Customers with no email: reminders are skipped', '=COUNTIFS(Customers!A2:A,"<>",Customers!C2:C,"")']];
  sh.getRange(34, 2, checks.length, 1).setValues(checks.map(function (c) { return [c[0]]; }));
  checks.forEach(function (c, i) { sh.getRange(34 + i, 6).setFormula(c[1]).setFontWeight('bold').setHorizontalAlignment('center'); });
  const res = sh.getRange('F34:F39');
  const cf = function (cond, bg, fg) { return cond.setBackground(bg).setFontColor(fg).setRanges([res]).build(); };
  sh.setConditionalFormatRules([
    cf(SpreadsheetApp.newConditionalFormatRule().whenTextContains('PASS'), '#c6efce', '#006100'),
    cf(SpreadsheetApp.newConditionalFormatRule().whenTextContains('FAIL'), '#ffc7ce', '#9c0006'),
    cf(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0), '#fff2cc', '#9c5700')]);

  const chart = sh.newChart().asColumnChart().addRange(sh.getRange('B9:C14')).setNumHeaders(1)
    .setOption('title', 'Owed by age (days late)').setOption('legend', { position: 'none' })
    .setOption('colors', [TEAL]).setOption('height', 260).setOption('width', 470)
    .setPosition(17, 8, 30, 0).build();
  sh.insertChart(chart);
}

function buildChecks_(sh) {
  sh.getRange('A1').setValue('Checks - written by Invoice tracker > Run self-test').setFontSize(14).setFontWeight('bold').setFontColor(NAVY);
  sh.getRange('A2').setValue('Last run');
  sh.getRange('B2').setNumberFormat('d mmm yyyy hh:mm');
  sh.getRange('C2').setValue('not run yet').setFontWeight('bold');
  sh.getRange('A3:D3').setValues([['Type', 'Check', 'Result', 'Detail']]);
  header_(sh.getRange('A3:D3'));
  sh.setColumnWidth(1, 120); sh.setColumnWidth(2, 520); sh.setColumnWidth(3, 80); sh.setColumnWidth(4, 420);
  const rng = sh.getRange('C4:C200');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PASS').setBackground('#c6efce').setFontColor('#006100').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FAIL').setBackground('#ffc7ce').setFontColor('#9c0006').setRanges([rng]).build()]);
}
