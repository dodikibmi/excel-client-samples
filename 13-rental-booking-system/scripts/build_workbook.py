"""Builds Event-Rental-Desk.xlsm with real Excel (COM), then replays 10 weeks of sample history through the workbook's
own macros, so every booking, fee, invoice PDF and log line in the sample was produced by the tool itself.

Sheets
  Home          sign in, message bar, 6 KPI cards, the selected booking with its actions (confirm, check out, check in
                with damaged / missing quantities, invoice + PDF, payment, cancel, lock period), lists: going out,
                due back / overdue, unpaid invoices, latest log lines; booked value by event month
  Availability  free quantity of every item for 21 days from a start date (red = none left, amber = under a quarter)
  Bookings      one row per booking (values written by the macros): status, dates, hire, delivery, late fee,
                damage, tax, total, deposit, paid, balance, invoice, PDF, who and when
  Lines         every booked item; columns L-P are formulas used by Availability and Checks
  Payments      deposits and balances received
  Catalog       yellow = typed: item, quantity owned, day rate, late fee per day, replacement cost; formulas: written
                off, out now, free today
  Customers, Users (role + PIN hash), Log (every attempt, allowed or refused), Invoice (A4 page that becomes the PDF)
  Checks        9 checks, all must say PASS     Settings     How to use     Calc (hidden)
VBA: modCore (rules), modButtons (Home buttons), frmSignIn, frmBooking.

    python scripts/build_workbook.py            (run under tools/with_lock.py)
"""
import csv, json, shutil, sys, time
from datetime import date
from pathlib import Path
import win32com.client as win32
from openpyxl.utils import get_column_letter as CL

sys.path.insert(0, str(Path(__file__).parent))
from vbom import vba_access
from watchdog import watch
import form

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Event-Rental-Desk.xlsm"
SRC = ROOT / "src"
SET = json.loads((ROOT / "data" / "settings.json").read_text())
HIST = json.loads((ROOT / "data" / "history.json").read_text())
CAT = list(csv.reader(open(ROOT / "data" / "catalog.csv", encoding="utf-8")))
CUST = list(csv.reader(open(ROOT / "data" / "customers.csv", encoding="utf-8")))
USERS = list(csv.reader(open(ROOT / "data" / "users.csv", encoding="utf-8")))
SER = lambda iso: (date.fromisoformat(iso) - date(1899, 12, 30)).days       # Excel serial (a datetime shifts a day)


def hexrgb(h):
    h = h.lstrip("#"); return int(h[:2], 16) + int(h[2:4], 16) * 256 + int(h[4:], 16) * 65536


PAGE, CARD, LINE = "#FAF7F2", "#FFFFFF", "#E7DFD3"
ACCENT, ACCENT2, S2, S3 = "#5B2A4E", "#8C5A7E", "#C8693F", "#7FA38A"
TEXT, MUTED, GOOD, BAD, AMBER, INPUT = "#2B2230", "#7A6E76", "#2F8A4E", "#B83A2E", "#D9A23B", "#FFF6D6"
SERIF, BODY = "Georgia", "Segoe UI"
BRAND = "   ".join(" ".join(w) for w in "LUMEN & LINEN EVENT RENTALS".split())
USD = "[$$-409]#,##0.00;-[$$-409]#,##0.00;\"-\""
USD0 = "[$$-409]#,##0;-[$$-409]#,##0"
DT = "d mmm yyyy"
NL, NB = 1500, 600                                  # rows kept for Lines / Bookings in formulas
MON = ",".join(f'"{m}"' for m in ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"))


def dtxt(x, year=True):
    """date -> 'd Mmm yyyy' as a formula, without TEXT() date codes (they follow the PC's language)"""
    return f'(DAY({x})&" "&CHOOSE(MONTH({x}),{MON})' + (f'&" "&YEAR({x}))' if year else ")")


def style(r, fg=None, bg=None, size=None, bold=None, name=None, italic=None, h=None, v=None, fmt=None, wrap=None):
    if fg: r.Font.Color = hexrgb(fg)
    if bg: r.Interior.Color = hexrgb(bg)
    if size: r.Font.Size = size
    if bold is not None: r.Font.Bold = bold
    if name: r.Font.Name = name
    if italic is not None: r.Font.Italic = italic
    if h: r.HorizontalAlignment = {"l": -4131, "c": -4108, "r": -4152}[h]
    if v: r.VerticalAlignment = {"t": -4160, "c": -4108, "b": -4107}[v]
    if fmt: r.NumberFormat = fmt
    if wrap is not None: r.WrapText = wrap
    return r


def border(r, color, edges=(7, 8, 9, 10), weight=2):
    for e in edges:
        b = r.Borders(e); b.LineStyle = 1; b.Weight = weight; b.Color = hexrgb(color)


def page(ws, title, sub, xl):
    ws.Cells.Font.Name = BODY; ws.Cells.Font.Size = 10; ws.Cells.Font.Color = hexrgb(TEXT)
    ws.Cells.Interior.Color = hexrgb(PAGE)
    style(ws.Range("B2"), fg=ACCENT2, size=8, bold=True).Value = BRAND
    style(ws.Range("B3"), size=20, name=SERIF, fg=ACCENT).Value = title
    style(ws.Range("B4"), fg=MUTED, size=9).Value = sub
    ws.Rows(3).RowHeight = 30
    ws.Columns("A").ColumnWidth = 2
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = False


def table(ws, cols, row=5, first_col=1):
    for i, (h, w) in enumerate(cols):
        c = ws.Cells(row, first_col + i); c.Value = h; ws.Columns(first_col + i).ColumnWidth = w
    style(ws.Range(ws.Cells(row, first_col), ws.Cells(row, first_col + len(cols) - 1)), fg="#FFFFFF", bg=ACCENT, bold=True, wrap=True, v="c")
    ws.Rows(row).RowHeight = 30


def button(ws, text, macro, left, top, w, h, primary=False, size=9, danger=False):
    b = ws.Shapes.AddShape(5, left, top, w, h)
    b.Fill.ForeColor.RGB = hexrgb(ACCENT if primary else ("#F6E3DE" if danger else CARD))
    b.Line.ForeColor.RGB = hexrgb(ACCENT if primary else (BAD if danger else LINE))
    tf = b.TextFrame2; tf.TextRange.Text = text
    tf.TextRange.Font.Size = size; tf.TextRange.Font.Bold = True; tf.TextRange.Font.Name = BODY
    tf.TextRange.Font.Fill.ForeColor.RGB = hexrgb("#FFFFFF" if primary else (BAD if danger else ACCENT))
    tf.TextRange.ParagraphFormat.Alignment = 2; tf.VerticalAnchor = 3
    tf.MarginLeft = tf.MarginRight = 2
    b.OnAction = macro
    b.Placement = 3
    return b


with vba_access():
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False; xl.ScreenUpdating = False
    for f in (OUT,):
        if f.exists(): f.unlink()
    shutil.rmtree(ROOT / SET["pdf_folder"], ignore_errors=True)
    try:
        wb = xl.Workbooks.Add()
        SHEETS = ["Home", "Availability", "Bookings", "Lines", "Payments", "Catalog", "Customers", "Users", "Log",
                  "Invoice", "Checks", "Settings", "How to use", "Calc"]
        while wb.Worksheets.Count < len(SHEETS):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, n in enumerate(SHEETS, 1):
            wb.Worksheets(i).Name = n
        W = {n: wb.Worksheets(n) for n in SHEETS}
        xl.Calculation = -4135
        name = lambda nm, ref: wb.Names.Add(nm, "=" + ref)

        # ------------------------------------------------------------------ Settings
        ws = W["Settings"]
        page(ws, "Settings", "Yellow = typed. Locked through, signed-in user and role are written by the macros.", xl)
        ws.Columns("B").ColumnWidth = 34; ws.Columns("C").ColumnWidth = 30; ws.Columns("D").ColumnWidth = 60
        rows = [("Company", "Company name", SET["company"], "@", "On the invoice", True),
                ("Address", "Address", SET["address"], "@", "", True),
                ("CoEmail", "Email", SET["email"], "@", "", True),
                ("Cur", "Currency sign", SET["currency"], "@", "", True),
                ("TaxRate", "Tax rate", SET["tax_rate"], "0.0%", "On hire, delivery, late fees and damage", True),
                ("DepositPct", "Deposit", SET["deposit_pct"], "0%", "Of the estimated total, asked when a booking is made", True),
                ("GraceDays", "Grace days before a late fee", SET["grace_days"], "0", "Late fee per item per day after that", True),
                ("MaxDays", "Longest hire (days)", SET["max_days"], "0", "", True),
                ("PdfFolder", "Invoice PDF folder", SET["pdf_folder"], "@", "Next to the workbook unless a full path", True),
                ("AsOfDate", "As-of date (blank = today)", "", DT, "Sample: frozen at 19 Mar 2027 so the screens stay the same", True),
                ("AsOf", "Today for the desk", "=IF(AsOfDate=\"\",TODAY(),AsOfDate)", DT, "", False),
                ("LockedThrough", "Period locked through", "", DT, "Written by Lock period (Manager)", False),
                ("LockedAt", "Locked at", "", "d mmm yyyy hh:mm", "", False),
                ("SignedUser", "Signed in", "", "@", "Written by Sign in", False),
                ("SignedRole", "Role", "", "@", "", False)]
        for i, (nm, lab, v, fmt, note, inp) in enumerate(rows, 6):
            ws.Cells(i, 2).Value = lab
            c = ws.Cells(i, 3); c.NumberFormat = fmt
            if isinstance(v, str) and v.startswith("="): c.Formula = v
            else: c.Value = v
            ws.Cells(i, 4).Value = note; ws.Cells(i, 4).Font.Color = hexrgb(MUTED)
            if inp: c.Interior.Color = hexrgb(INPUT); c.Locked = False
            name(nm, f"Settings!$C${i}")

        # ------------------------------------------------------------------ Catalog, Customers, Users
        ws = W["Catalog"]
        page(ws, "Catalog", "Yellow = typed. Written off, out now and free today are formulas on the bookings.", xl)
        table(ws, [("Code", 10), ("Category", 13), ("Item", 30), ("Owned", 8), ("Day rate", 10), ("Late fee / day", 10),
                   ("Replace-ment cost", 11), ("Written off", 9), ("In stock", 9), ("Out now", 9), ("Free today", 9)])
        rows = [[r[0], r[1], r[2], int(r[3]), float(r[4]), float(r[5]), float(r[6])] for r in CAT[1:]]
        last = 5 + len(rows)
        ws.Range(f"A6:G{last}").Value = rows
        style(ws.Range(f"A6:G{last}"), bg=INPUT).Locked = False
        L = lambda c: f"Lines!${c}$6:${c}${NL}"
        ws.Range(f"H6:H{last}").Formula = f"=SUMIFS({L('K')},{L('C')},A6)"
        ws.Range(f"I6:I{last}").Formula = "=D6-H6"
        ws.Range(f"J6:J{last}").Formula = f'=SUMIFS({L("E")},{L("C")},A6,{L("L")},"Out")'
        ws.Range(f"K6:K{last}").Formula = f'=D6-SUMIFS({L("K")},{L("C")},A6,{L("P")},"<="&AsOf)-SUMIFS({L("E")},{L("C")},A6,{L("O")},1,{L("M")},"<="&AsOf,{L("N")},">="&AsOf)'
        ws.Range(f"E6:G{last}").NumberFormat = USD
        ws.Range(f"H6:K{last}").Interior.Color = hexrgb(CARD)
        CAT_LAST = last
        ws = W["Customers"]
        page(ws, "Customers", "Yellow = typed.", xl)
        table(ws, [("ID", 8), ("Customer", 30), ("Contact", 16), ("Email", 34), ("Type", 12)])
        last = 5 + len(CUST) - 1
        ws.Range(f"A6:E{last}").Value = CUST[1:]
        style(ws.Range(f"A6:E{last}"), bg=INPUT).Locked = False
        ws = W["Users"]
        page(ws, "Users", "Roles: Staff (bookings, check out / in, invoices, payments) and Manager (also cancel and lock a period). Only a hash of each PIN is kept.", xl)
        table(ws, [("User", 10), ("Name", 18), ("Role", 10), ("PIN hash", 16)])
        ws.Range(f"A6:C{5 + len(USERS) - 1}").Value = [u[:3] for u in USERS[1:]]

        # ------------------------------------------------------------------ Bookings, Lines, Payments, Log
        ws = W["Bookings"]
        page(ws, "Bookings", "One row per booking. Written by the macros only (the sheet is protected); filter freely.", xl)
        table(ws, [("Booking", 13), ("Status", 10), ("Cust.", 6), ("Customer", 24), ("Event", 11), ("Out", 11), ("Back by", 11),
                   ("Returned", 11), ("Days", 5), ("Hire", 10), ("Delivery & setup", 9), ("Late fee", 9), ("Damaged / missing", 10),
                   ("Subtotal", 10), ("Tax", 9), ("Total", 10), ("Deposit due", 9), ("Paid", 10), ("Balance", 10),
                   ("Invoice", 13), ("Invoiced", 11), ("PDF", 30), ("By", 6), ("Created", 15), ("Notes", 30)])
        ws.Range("E:H").NumberFormat = DT; ws.Range("U:U").NumberFormat = DT; ws.Range("X:X").NumberFormat = "d mmm yy hh:mm"
        ws.Range("J:S").NumberFormat = USD
        for c in ("E5", "J5", "U5", "X5"):
            ws.Range(c).NumberFormat = "General"
        ws.Range("A6").Select(); xl.ActiveWindow.FreezePanes = True
        st = ws.Range(f"B6:B{NB}")
        for txt, col in (("Enquiry", "#8C8C8C"), ("Confirmed", ACCENT), ("Out", S2), ("Returned", AMBER), ("Invoiced", BAD), ("Paid", GOOD), ("Cancelled", "#B0A8AE")):
            fc = st.FormatConditions.Add(1, 3, f'="{txt}"'); fc.Font.Color = hexrgb(col); fc.Font.Bold = True
        ws = W["Lines"]
        page(ws, "Booking lines", "Columns A-K written by the macros; L-P are formulas (status, out date, held until, holds stock, returned) used by Availability and Checks.", xl)
        table(ws, [("Booking", 13), ("#", 4), ("Code", 8), ("Item", 26), ("Qty", 6), ("Day rate", 9), ("Days", 5), ("Amount", 10),
                   ("Late fee / day", 8), ("Replace-ment", 9), ("Damaged / missing", 9), ("Status", 10), ("Out", 11), ("Held until", 11),
                   ("Holds stock", 7), ("Returned", 11)])
        ws.Range("F:F").NumberFormat = USD; ws.Range("H:J").NumberFormat = USD
        ws.Range("M:N").NumberFormat = DT; ws.Range("P:P").NumberFormat = DT
        for c in ("F5", "H5", "I5", "J5", "M5", "N5", "P5"): ws.Range(c).NumberFormat = "General"
        style(ws.Range("L5:P5"), bg=ACCENT2)
        ws.Range("A6").Select(); xl.ActiveWindow.FreezePanes = True
        ws = W["Payments"]
        page(ws, "Payments", "Deposits and balances received. Written by Record payment only.", xl)
        table(ws, [("Booking", 13), ("Date", 11), ("Kind", 9), ("Amount", 11), ("By", 6), ("Recorded", 16)])
        ws.Range("B:B").NumberFormat = DT; ws.Range("D:D").NumberFormat = USD; ws.Range("F:F").NumberFormat = "d mmm yy hh:mm"
        for c in ("B5", "D5", "F5"): ws.Range(c).NumberFormat = "General"
        ws = W["Log"]
        page(ws, "Log (audit trail)", "Every attempt: who, role, what, on which booking, allowed (OK) or refused (BLOCKED) and why. Written by the macros only.", xl)
        table(ws, [("When", 16), ("User", 7), ("Role", 9), ("Action", 13), ("Booking", 13), ("Result", 9), ("Detail", 90)])
        ws.Range("A:A").NumberFormat = "d mmm yy hh:mm"; ws.Range("A5").NumberFormat = "General"
        fc = ws.Range("F6:F20000").FormatConditions.Add(1, 3, '="BLOCKED"'); fc.Font.Color = hexrgb(BAD); fc.Font.Bold = True
        ws.Range("A6").Select(); xl.ActiveWindow.FreezePanes = True

        # ------------------------------------------------------------------ Calc (hidden): lists and the full stock grid
        ws = W["Calc"]
        B = lambda c: f"Bookings!${c}$6:${c}${NB}"
        RB = f"ROW(Bookings!$A$6:$A${NB})"
        ws.Range("A1").Value = "Helpers for Home and Checks. Hidden; nothing to type."
        # open bookings for the Home picker (not Paid / Cancelled), in booking order
        ws.Range("A2").Value = "Open bookings"
        ws.Range("A3:A302").Formula = (f'=IFERROR(INDEX(Bookings!$A:$A,AGGREGATE(15,6,{RB}/(({B("A")}<>"")*({B("B")}<>"Paid")*({B("B")}<>"Cancelled")),ROW()-2)),"")')
        # lists: key = date + row/100000, sorted
        lists = {"C": ("Going out", f'{B("F")}', f'({B("B")}="Confirmed")*({B("F")}<=AsOf+7)'),
                 "H": ("Due back", f'{B("G")}', f'({B("B")}="Out")'),
                 "M": ("Unpaid invoices", f'{B("U")}', f'({B("B")}="Invoiced")')}
        for col, (title, key, cond) in lists.items():
            c0 = ws.Range(f"{col}1").Column
            ws.Cells(2, c0).Value = title
            for k in range(1, 9):
                r = 2 + k
                ws.Cells(r, c0).Formula = f"=IFERROR(AGGREGATE(15,6,({key}+{RB}/100000)/({cond}),{k}),\"\")"
                a1 = f"{CL(c0)}{r}"
                ws.Cells(r, c0 + 1).Formula = f'=IF({a1}="","",ROUND(MOD({a1},1)*100000,0))'
        # lines of the selected booking (Home)
        ws.Range("R2").Value = "Selected booking lines (row in Lines)"
        RL = f"ROW(Lines!$A$6:$A${NL})"
        ws.Range("R3:R22").Formula = f'=IFERROR(AGGREGATE(15,6,{RL}/(Lines!$A$6:$A${NL}=SelBooking),ROW()-2),"")'
        # booked value by event month (not cancelled / enquiry)
        ws.Range("T2:V2").Value = ("Month", "Booked value", "Bookings")
        for i, (m, lab) in enumerate(((1, "Jan 2027"), (2, "Feb 2027"), (3, "Mar 2027"), (4, "Apr 2027")), 3):
            ws.Cells(i, 20).NumberFormat = "@"; ws.Cells(i, 20).Value = lab
            rng = f'{B("E")},">="&DATE(2027,{m},1),{B("E")},"<"&DATE(2027,{m + 1},1),{B("B")},"<>Cancelled",{B("B")},"<>Enquiry"'
            ws.Cells(i, 21).Formula = f"=SUMIFS({B('P')},{rng})"
            ws.Cells(i, 22).Formula = f"=COUNTIFS({rng})"
        # the full stock grid: every item x every day from 4 Jan 2027 for 120 days (Checks: never below 0)
        ws.Range("X2").Value = "Stock grid: free quantity per item and day"
        ws.Range("Y3").Formula = "=DATE(2027,1,4)"
        ws.Range("Z3:EN3").Formula = "=Y3+1"
        n_items = CAT_LAST - 5
        ws.Range(f"X4:X{3 + n_items}").Formula = "=INDEX(Catalog!$A:$A,ROW()+2)"
        ws.Range(f"Y4:EN{3 + n_items}").Formula = (
            f'=INDEX(Catalog!$D:$D,ROW()+2)-SUMIFS(Lines!$K$6:$K${NL},Lines!$C$6:$C${NL},$X4,Lines!$P$6:$P${NL},"<="&Y$3)'
            f'-SUMIFS(Lines!$E$6:$E${NL},Lines!$C$6:$C${NL},$X4,Lines!$O$6:$O${NL},1,Lines!$M$6:$M${NL},"<="&Y$3,Lines!$N$6:$N${NL},">="&Y$3)')
        GRID = f"Calc!$Y$4:$EN${3 + n_items}"
        ws.Visible = 0

        # ------------------------------------------------------------------ Availability
        ws = W["Availability"]
        page(ws, "Availability", "Free quantity of every item for 21 days (owned - written off - held by confirmed or out bookings). Red = none left, amber = under a quarter.", xl)
        ws.Range("B6").Value = "Start date (blank = today)"; style(ws.Range("D6"), bg=INPUT, fmt=DT).Locked = False
        name("AvailStart", "Availability!$D$6")
        ws.Columns("B").ColumnWidth = 9; ws.Columns("C").ColumnWidth = 28; ws.Columns("D").ColumnWidth = 11
        ws.Range("B8:D8").Value = ("Code", "Item", "Owned")
        ws.Range("E8").Formula = "=IF(AvailStart=\"\",AsOf,AvailStart)"
        ws.Range("F8:Y8").Formula = "=E8+1"
        ws.Range("E7:Y7").Formula = '=CHOOSE(WEEKDAY(E8),"Sun","Mon","Tue","Wed","Thu","Fri","Sat")'
        ws.Range("E8:Y8").NumberFormat = "d mmm"
        style(ws.Range("B8:Y8"), fg="#FFFFFF", bg=ACCENT, bold=True); style(ws.Range("E7:Y7"), fg=MUTED, size=8, h="c")
        ws.Range("E8:Y8").HorizontalAlignment = -4108
        for c in range(5, 26): ws.Columns(c).ColumnWidth = 6.2
        a_last = 8 + n_items
        ws.Range(f"B9:B{a_last}").Formula = "=INDEX(Catalog!$A:$A,ROW()-3)"
        ws.Range(f"C9:C{a_last}").Formula = "=INDEX(Catalog!$C:$C,ROW()-3)"
        ws.Range(f"D9:D{a_last}").Formula = "=INDEX(Catalog!$D:$D,ROW()-3)"
        ws.Range(f"E9:Y{a_last}").Formula = (
            f'=$D9-SUMIFS(Lines!$K$6:$K${NL},Lines!$C$6:$C${NL},$B9,Lines!$P$6:$P${NL},"<="&E$8)'
            f'-SUMIFS(Lines!$E$6:$E${NL},Lines!$C$6:$C${NL},$B9,Lines!$O$6:$O${NL},1,Lines!$M$6:$M${NL},"<="&E$8,Lines!$N$6:$N${NL},">="&E$8)')
        body = ws.Range(f"E9:Y{a_last}"); body.HorizontalAlignment = -4108; body.Interior.Color = hexrgb(CARD)
        fc = body.FormatConditions.Add(1, 8, "=0"); fc.Interior.Color = hexrgb("#F2C4BE"); fc.Font.Color = hexrgb(BAD); fc.Font.Bold = True
        fc = body.FormatConditions.Add(2, Formula1="=E9<$D9/4"); fc.Interior.Color = hexrgb("#FBE7C2")
        border(ws.Range(f"B8:Y{a_last}"), LINE, (7, 8, 9, 10, 11, 12))
        ws.Range(f"B{a_last + 2}").Value = "Held = Confirmed, Out, Returned, Invoiced or Paid bookings, from the out date to the back-by date; a booking still out after its back-by date keeps its items until it comes back. Enquiries hold nothing."
        style(ws.Range(f"B{a_last + 2}"), fg=MUTED, size=9)
        ws.Range("E9").Select(); xl.ActiveWindow.FreezePanes = True

        # ------------------------------------------------------------------ Invoice (A4)
        ws = W["Invoice"]
        ws.Cells.Font.Name = BODY; ws.Cells.Font.Size = 10; ws.Cells.Font.Color = hexrgb(TEXT)
        ws.Activate(); xl.ActiveWindow.DisplayGridlines = False
        for col, w in zip("ABCDEFGH", (2, 12, 34, 8, 11, 8, 13, 2)): ws.Columns(col).ColumnWidth = w
        ws.Range("J1").Value = "Booking shown:"; style(ws.Range("K1"), bg=INPUT).Locked = False
        name("CurInv", "Invoice!$K$1")
        BR = "MATCH(CurInv,Bookings!$A:$A,0)"
        bk = lambda c: f"INDEX(Bookings!${c}:${c},{BR})"
        style(ws.Range("B2"), size=20, name=SERIF, fg=ACCENT).Formula = "=Company"
        style(ws.Range("B3"), fg=MUTED, size=9).Formula = '=Address&"  ·  "&CoEmail'
        style(ws.Range("F2:G2").Cells(1, 2), size=18, name=SERIF, h="r", fg=ACCENT).Value = "INVOICE"
        lab = [("Invoice no.", bk("T") + '&""', "@"), ("Date", bk("U"), DT), ("Booking", "CurInv&\"\"", "@")]
        for i, (t, f, fmt) in enumerate(lab, 4):
            ws.Cells(i, 6).Value = t; ws.Cells(i, 6).Font.Color = hexrgb(MUTED); ws.Cells(i, 6).HorizontalAlignment = -4152
            c = ws.Cells(i, 7); c.NumberFormat = "General" if fmt == "@" else fmt; c.Formula = "=" + f; c.HorizontalAlignment = -4152; c.Font.Bold = True
        ws.Range("B6").Value = "Bill to"; ws.Range("B6").Font.Color = hexrgb(MUTED)
        style(ws.Range("B7"), bold=True, size=11).Formula = "=" + bk("D")
        ws.Range("B8").Formula = f'=INDEX(Customers!$C:$C,MATCH({bk("C")},Customers!$A:$A,0))&"  ·  "&INDEX(Customers!$D:$D,MATCH({bk("C")},Customers!$A:$A,0))'
        ws.Range("B10").Formula = (f'="Event "&{dtxt(bk("E"))}&"   ·   out "&{dtxt(bk("F"), False)}&"   ·   back by "&{dtxt(bk("G"), False)}'
                                   f'&IF({bk("H")}="","","   ·   returned "&{dtxt(bk("H"), False)})&"   ·   "&{bk("I")}&" hire days"')
        ws.Range("B10").Font.Color = hexrgb(MUTED)
        style(ws.Range("B12:G12"), fg="#FFFFFF", bg=ACCENT, bold=True).Value = ("Code", "Item", "Qty", "Day rate", "Days", "Amount")
        ws.Range("D12:G12").HorizontalAlignment = -4152
        RLI = f"AGGREGATE(15,6,ROW(Lines!$A$6:$A${NL})/(Lines!$A$6:$A${NL}=CurInv),ROW()-12)"
        for col, src in zip("BCDEFG", "CDEFGH"):
            ws.Range(f"{col}13:{col}32").Formula = f'=IFERROR(INDEX(Lines!${src}:${src},{RLI}),"")'
        ws.Range("C13:C32").Formula = (f'=IFERROR(INDEX(Lines!$D:$D,{RLI})&IF(INDEX(Lines!$K:$K,{RLI})>0,"  ·  "&INDEX(Lines!$K:$K,{RLI})'
                                       f'&" damaged / missing","")&"","")')
        ws.Range("E13:E32").NumberFormat = USD; ws.Range("G13:G32").NumberFormat = USD
        border(ws.Range("B12:G32"), LINE, (8, 12))
        tot = [("Hire", "J"), ("Delivery & setup", "K"), ("Late return fee", "L"), ("Damaged / missing items", "M"), ("Subtotal", "N"),
               ("Tax", "O"), ("TOTAL", "P"), ("Paid so far", "R"), ("BALANCE DUE", "S")]
        for i, (t, c) in enumerate(tot, 34):
            ws.Cells(i, 5).Value = t; ws.Range(ws.Cells(i, 5), ws.Cells(i, 6)).Merge(); ws.Cells(i, 5).HorizontalAlignment = -4152
            v = ws.Cells(i, 7); v.NumberFormat = USD; v.Formula = "=" + bk(c); v.HorizontalAlignment = -4152
            if t in ("TOTAL", "BALANCE DUE"):
                style(ws.Range(ws.Cells(i, 5), ws.Cells(i, 7)), bold=True, fg=ACCENT, size=11)
        ws.Range("B34").Formula = f'=IF({bk("L")}>0,"Late fee: returned "&({bk("H")}-{bk("G")})&" days after the back-by date ("&TEXT(GraceDays,"0")&" grace day)","")'
        ws.Range("B35").Formula = f'=IF({bk("M")}>0,"Damaged or missing items are charged at replacement cost.","")'
        ws.Range("B34:B35").Font.Color = hexrgb(MUTED); ws.Range("B34:B35").Font.Size = 9
        style(ws.Range("B45"), fg=MUTED, size=8).Value = "Sample project, fictional company and customers. Payment within 14 days. Thank you!"
        ps = ws.PageSetup; ps.PrintArea = "$A$1:$H$46"; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
        ps.LeftMargin = ps.RightMargin = 28; ps.TopMargin = ps.BottomMargin = 28; ps.CenterHorizontally = True

        # ------------------------------------------------------------------ Checks
        ws = W["Checks"]
        page(ws, "Checks", "The data checks itself. All must say PASS.", xl)
        ws.Columns("B").ColumnWidth = 74; ws.Columns("C").ColumnWidth = 12; ws.Columns("D").ColumnWidth = 10
        style(ws.Range("B6:D6"), fg="#FFFFFF", bg=ACCENT, bold=True).Value = ("Check", "Found", "Result")
        NBK = f'({B("A")}<>"")'
        chk = [("No item over-booked on any day (every item, 4 Jan - 3 May 2027)", f'=COUNTIF({GRID},"<0")', "=C{r}=0"),
               ("Booking and invoice numbers are unique", f'=SUMPRODUCT({NBK}*(COUNTIF({B("A")},{B("A")})>1))+SUMPRODUCT(({B("T")}<>"")*(COUNTIF({B("T")},{B("T")})>1))', "=C{r}=0"),
               ("Every line belongs to a booking and to an item in the catalog",
                f'=SUMPRODUCT((Lines!$A$6:$A${NL}<>"")*ISNA(MATCH(Lines!$A$6:$A${NL},{B("A")},0)))+SUMPRODUCT((Lines!$C$6:$C${NL}<>"")*ISNA(MATCH(Lines!$C$6:$C${NL},Catalog!$A$6:$A$200,0)))', "=C{r}=0"),
               ("Hire on every booking = its line amounts", f'=SUMPRODUCT({NBK}*(ABS({B("J")}-SUMIF(Lines!$A$6:$A${NL},{B("A")},Lines!$H$6:$H${NL}))>0.005))', "=C{r}=0"),
               ("Total = hire + delivery + late fee + damage + tax (tax rounded to the cent)",
                f'=SUMPRODUCT({NBK}*(ABS({B("N")}-{B("J")}-{B("K")}-{B("L")}-{B("M")})+ABS({B("O")}-ROUND({B("N")}*TaxRate,2))+ABS({B("P")}-{B("N")}-{B("O")})>0.005))', "=C{r}=0"),
               ("Paid = the payments recorded; balance = total - paid",
                f'=SUMPRODUCT({NBK}*(ABS({B("R")}-SUMIF(Payments!$A$6:$A$5000,{B("A")},Payments!$D$6:$D$5000))+ABS({B("S")}-{B("P")}+{B("R")})>0.005))', "=C{r}=0"),
               ("Status matches the money: Paid has no balance, Invoiced still has one",
                f'=COUNTIFS({B("B")},"Paid",{B("S")},"<>0")+COUNTIFS({B("B")},"Invoiced",{B("S")},"<=0")', "=C{r}=0"),
               ("Damaged / missing never more than hired", f'=SUMPRODUCT(--(Lines!$K$6:$K${NL}>Lines!$E$6:$E${NL}))', "=C{r}=0"),
               ("Nothing recorded inside the locked period after it was locked (payments)",
                f'=IF(LockedThrough="",0,COUNTIFS(Payments!$B$6:$B$5000,"<="&LockedThrough,Payments!$F$6:$F$5000,">"&LockedAt))', "=C{r}=0"),
               ("Every refused attempt in the Log says why", '=COUNTIFS(Log!$F$6:$F$20000,"BLOCKED",Log!$G$6:$G$20000,"")', "=C{r}=0")]
        for k, (txt, ff, test) in enumerate(chk, 7):
            ws.Cells(k, 2).Value = txt; ws.Cells(k, 3).Formula = ff
            ws.Cells(k, 4).Formula = f'=IF({test.format(r=k)[1:]},"PASS","FAIL")'
        last = 6 + len(chk)
        d = ws.Range(f"D7:D{last}")
        fc = d.FormatConditions.Add(1, 3, '="PASS"'); fc.Font.Color = hexrgb(GOOD); fc.Interior.Color = hexrgb("#E2F0E5")
        fc = d.FormatConditions.Add(1, 3, '="FAIL"'); fc.Font.Color = hexrgb(BAD); fc.Interior.Color = hexrgb("#F6DEDA")
        d.Font.Bold = True; d.HorizontalAlignment = -4108
        border(ws.Range(f"B6:D{last}"), LINE, (7, 8, 9, 10, 12))
        ws.Range(f"B{last + 2}").Value = "Checks passed"; ws.Range(f"D{last + 2}").Formula = f'=COUNTIF(D7:D{last},"PASS")'
        ws.Range(f"B{last + 3}").Value = "Checks in total"; ws.Range(f"D{last + 3}").Value = len(chk)
        name("ChecksPassed", f"Checks!$D${last + 2}"); name("ChecksTotal", f"Checks!$D${last + 3}")

        # ------------------------------------------------------------------ Home
        ws = W["Home"]
        page(ws, "Event rental desk", "", xl)
        xl.ActiveWindow.DisplayHeadings = False
        for c in range(2, 21):
            ws.Columns(c).ColumnWidth = 2.2 if c in (5, 9, 13, 17) else 10
        ws.Range("B4").Formula = ('=IF(SignedUser="","Nobody signed in",INDEX(Users!$B:$B,MATCH(SignedUser,Users!$A:$A,0))&" ("&SignedRole&")")'
                                  f'&"   ·   today "&{dtxt("AsOf")}&IF(LockedThrough="","","   ·   locked through "&{dtxt("LockedThrough")})')
        ws.Range("B4").Font.Color = hexrgb(MUTED)
        x0 = ws.Range("K2").Left
        for i, (t, m, prim) in enumerate((("Sign in", "BtnSignIn", True), ("Sign out", "BtnSignOut", False), ("New booking", "BtnNew", True),
                                          ("Availability", "GoAvailability", False), ("Bookings", "GoBookings", False), ("Log", "GoLog", False))):
            button(ws, t, m, x0 + i * 76, 14, 72, 22, primary=prim)
        ws.Range("B6:T6").Merge()
        style(ws.Range("B6"), bold=True, bg="#F1EBE3", size=10, v="c").Value = "Sign in to start (Sign in button, top right)"
        ws.Range("B6").IndentLevel = 1; ws.Rows(6).RowHeight = 22
        name("HomeMsg", "Home!$B$6")
        W["Calc"].Range("B1").Formula = '=LEFT(HomeMsg,7)="BLOCKED"'
        fc = ws.Range("B6").FormatConditions.Add(2, Formula1="=Calc!$B$1"); fc.Font.Color = hexrgb(BAD); fc.Interior.Color = hexrgb("#F6DEDA")
        # KPI cards
        Bk = lambda c: B(c)
        cards = [("CONFIRMED, NEXT 30 DAYS", f'=COUNTIFS({Bk("B")},"Confirmed",{Bk("F")},">="&AsOf,{Bk("F")},"<="&AsOf+30)', "0",
                  f'=TEXT(SUMIFS({Bk("P")},{Bk("B")},"Confirmed",{Bk("F")},">="&AsOf,{Bk("F")},"<="&AsOf+30),"[$$-409]#,##0")&" booked"'),
                 ("ITEMS OUT NOW", f'=SUMIFS(Lines!$E$6:$E${NL},Lines!$L$6:$L${NL},"Out")', "#,##0",
                  f'="on "&COUNTIF({Bk("B")},"Out")&" bookings"'),
                 ("OVERDUE RETURNS", f'=COUNTIFS({Bk("B")},"Out",{Bk("G")},"<"&AsOf)', "0", '="late fees run per item per day"'),
                 ("INVOICED THIS MONTH", f'=SUMIFS({Bk("P")},{Bk("U")},">="&DATE(YEAR(AsOf),MONTH(AsOf),1),{Bk("U")},"<="&AsOf)', USD0,
                  f'=COUNTIFS({Bk("U")},">="&DATE(YEAR(AsOf),MONTH(AsOf),1),{Bk("U")},"<="&AsOf)&" invoices"'),
                 ("UNPAID BALANCE", f'=SUMIFS({Bk("S")},{Bk("B")},"Invoiced")', USD0, f'=COUNTIF({Bk("B")},"Invoiced")&" invoices open"')]
        for i, (lab, ff, fmt, note) in enumerate(cards):
            c0 = 2 + i * 4 if i < 4 else 18
            box = ws.Range(ws.Cells(8, c0), ws.Cells(11, c0 + (2 if i < 4 else 2)))
            box.Interior.Color = hexrgb(CARD); border(box, LINE)
            border(ws.Range(ws.Cells(8, c0), ws.Cells(8, c0 + 2)), ACCENT, (8,), -4138)
            ws.Range(ws.Cells(8, c0), ws.Cells(8, c0 + 2)).Merge()
            style(ws.Cells(8, c0), fg=MUTED, size=8, bold=True).Value = lab; ws.Cells(8, c0).IndentLevel = 1
            ws.Range(ws.Cells(9, c0), ws.Cells(10, c0 + 2)).Merge()
            v = ws.Cells(9, c0); v.NumberFormat = fmt; v.Formula = ff
            style(v, fg=ACCENT, size=20, name=SERIF, v="c", h="l"); v.IndentLevel = 1
            ws.Range(ws.Cells(11, c0), ws.Cells(11, c0 + 2)).Merge()
            style(ws.Cells(11, c0), fg=MUTED, size=8).Formula = note; ws.Cells(11, c0).IndentLevel = 1
            if lab == "OVERDUE RETURNS":
                fc = v.FormatConditions.Add(1, 5, "=0"); fc.Font.Color = hexrgb(BAD)
        for r, hgt in ((7, 8), (8, 20), (9, 18), (10, 18), (11, 16), (12, 10)):
            ws.Rows(r).RowHeight = hgt
        # selected booking
        style(ws.Range("B13"), fg=ACCENT2, size=8, bold=True).Value = "S E L E C T E D   B O O K I N G"
        ws.Range("B14").Value = "Booking"; style(ws.Range("C14:D14"), bg=INPUT, bold=True, size=11).Merge()
        ws.Range("C14:D14").Locked = False
        name("SelBooking", "Home!$C$14")
        dv = ws.Range("C14").Validation; dv.Delete(); dv.Add(3, 1, 1, "=Calc!$A$3:$A$302")
        SR = "MATCH(SelBooking,Bookings!$A:$A,0)"
        sb = lambda c: f'IFERROR(INDEX(Bookings!${c}:${c},{SR}),"")'
        sbt = lambda c: f'IFERROR(INDEX(Bookings!${c}:${c},{SR})&"","")'
        det = [("Status", sbt("B"), "@"), ("Customer", sbt("D"), "@"),
               ("Dates", f'IFERROR("event "&{dtxt(sb("E"), False)}&"  ·  out "&{dtxt(sb("F"), False)}&"  ·  back by "&{dtxt(sb("G"), False)}&IF({sb("H")}="","","  ·  returned "&{dtxt(sb("H"), False)}),"")', "@"),
               ("Total", sb("P"), USD), ("Paid", sb("R"), USD), ("Balance", sb("S"), USD), ("Invoice", sbt("T"), "@")]
        for i, (t, f, fmt) in enumerate(det, 15):
            ws.Cells(i, 2).Value = t; ws.Cells(i, 2).Font.Color = hexrgb(MUTED)
            ws.Range(ws.Cells(i, 3), ws.Cells(i, 8)).Merge()
            c = ws.Cells(i, 3); c.NumberFormat = "General" if fmt == "@" else fmt; c.Formula = "=" + f; c.HorizontalAlignment = -4131
        ws.Range("C15").Font.Bold = True
        inputs = [(22, "Date", "ActionDate", DT, "action date, blank = today"), (23, "Amount", "PayAmount", USD, "received, for Record payment"),
                  (24, "Reason", "CancelReason", "@", "for Cancel (Manager)"), (25, "Lock to", "LockDate", DT, "last day to close (Manager)")]
        for r, t, nm, fmt, note in inputs:
            ws.Cells(r, 2).Value = t; ws.Cells(r, 2).Font.Color = hexrgb(MUTED)
            ws.Range(f"C{r}:D{r}").Merge(); style(ws.Range(f"C{r}:D{r}"), bg=INPUT, fmt=fmt)
            ws.Range(f"C{r}:D{r}").Locked = False
            ws.Cells(r, 6).Value = note; ws.Cells(r, 6).Font.Color = hexrgb(MUTED); ws.Cells(r, 6).Font.Size = 8
            name(nm, f"Home!$C${r}")
        top = ws.Range("B27").Top
        for i, (t, m, prim, danger) in enumerate((("Confirm", "BtnConfirm", False, False), ("Check out", "BtnCheckOut", True, False),
                                                  ("Check in", "BtnCheckIn", True, False), ("Invoice + PDF", "BtnInvoice", True, False))):
            button(ws, t, m, ws.Range("B27").Left + i * 74, top, 70, 22, primary=prim)
        top2 = ws.Range("B29").Top
        for i, (t, m, prim, danger) in enumerate((("Record payment", "BtnPay", False, False), ("Cancel (Mgr)", "BtnCancel", False, True),
                                                  ("Lock period (Mgr)", "BtnLock", False, True))):
            button(ws, t, m, ws.Range("B27").Left + i * 98, top2, 94, 22, danger=danger)
        ws.Rows(28).RowHeight = 8
        style(ws.Range("B32"), fg=ACCENT2, size=8, bold=True).Value = "L I N E S   ·   type damaged or missing before Check in"
        style(ws.Range("B33:H33"), fg="#FFFFFF", bg=ACCENT, bold=True).Value = ("Code", "Item", "", "", "Qty", "Damaged / missing", "")
        ws.Range("C33:E33").Merge()
        for k in range(12):
            r = 34 + k
            ws.Cells(r, 2).Formula = f'=IFERROR(INDEX(Lines!$C:$C,Calc!$R${3 + k}),"")'
            ws.Range(f"C{r}:E{r}").Merge(); ws.Cells(r, 3).Formula = f'=IFERROR(INDEX(Lines!$D:$D,Calc!$R${3 + k}),"")'
            ws.Cells(r, 6).Formula = f'=IFERROR(INDEX(Lines!$E:$E,Calc!$R${3 + k}),"")'
            ws.Range(f"G{r}:H{r}").Merge(); style(ws.Range(f"G{r}:H{r}"), bg=INPUT, h="c"); ws.Range(f"G{r}:H{r}").Locked = False
        border(ws.Range("B33:H45"), LINE, (8, 12))
        name("HomeLineCode", "Home!$B$34:$B$45"); name("HomeDamaged", "Home!$G$34:$H$45")             # merged G:H, so ClearContents works
        # right side lists
        CalcR = lambda col, k: f"Calc!${col}${2 + k}"

        def lst(r0, title, keycol, cols):
            style(ws.Range(f"J{r0}"), fg=ACCENT2, size=8, bold=True).Value = title
            style(ws.Range(f"J{r0 + 1}:T{r0 + 1}"), fg="#FFFFFF", bg=ACCENT, bold=True)
            for j, (h, c0, c1, src, fmt) in enumerate(cols):
                ws.Range(f"{c0}{r0 + 1}:{c1}{r0 + 1}").Merge(); ws.Range(f"{c0}{r0 + 1}").Value = h
                for k in range(1, 9):
                    r = r0 + 1 + k
                    ws.Range(f"{c0}{r}:{c1}{r}").Merge()
                    c = ws.Range(f"{c0}{r}"); c.NumberFormat = "General" if fmt == "@" else fmt
                    c.Formula = f'=IF({CalcR(keycol, k)}="","",INDEX(Bookings!${src}:${src},{CalcR(keycol, k)}))'
            border(ws.Range(f"J{r0 + 1}:T{r0 + 9}"), LINE, (8, 12))
        lst(13, "G O I N G   O U T   ( N E X T   7   D A Y S )", "D",
            [("Booking", "J", "L", "A", "@"), ("Customer", "N", "Q", "D", "@"), ("Out", "R", "R", "F", "d mmm"), ("Total", "S", "T", "P", USD0)])
        lst(24, "D U E   B A C K   /   O V E R D U E", "I",
            [("Booking", "J", "L", "A", "@"), ("Customer", "N", "Q", "D", "@"), ("Back by", "R", "R", "G", "d mmm"), ("Total", "S", "T", "P", USD0)])
        fc = ws.Range("R26:R33").FormatConditions.Add(2, Formula1="=($R26<>\"\")*($R26<AsOf)"); fc.Font.Color = hexrgb(BAD); fc.Font.Bold = True
        lst(35, "U N P A I D   I N V O I C E S", "N",
            [("Booking", "J", "L", "A", "@"), ("Customer", "N", "Q", "D", "@"), ("Invoiced", "R", "R", "U", "d mmm"), ("Balance", "S", "T", "S", USD0)])
        # chart: booked value by event month
        a, b = ws.Range("B48"), ws.Range("T62")
        co = ws.ChartObjects().Add(a.Left, a.Top, b.Left + b.Width - a.Left, b.Top + b.Height - a.Top)
        ch = co.Chart; ch.ChartType = 51
        while ch.SeriesCollection().Count: ch.SeriesCollection(1).Delete()
        se = ch.SeriesCollection().NewSeries(); se.Name = "Booked value"; se.Values = W["Calc"].Range("U3:U6"); se.XValues = W["Calc"].Range("T3:T6")
        se.Format.Fill.ForeColor.RGB = hexrgb(ACCENT); se.HasDataLabels = True; se.DataLabels().NumberFormat = USD0
        ch.ChartArea.Format.Fill.ForeColor.RGB = hexrgb(CARD); ch.ChartArea.Format.Line.ForeColor.RGB = hexrgb(LINE)
        ch.PlotArea.Format.Fill.Visible = False
        tr = ch.ChartArea.Format.TextFrame2.TextRange.Font; tr.Fill.ForeColor.RGB = hexrgb(MUTED); tr.Size = 9; tr.Name = BODY
        ch.HasTitle = True; ch.ChartTitle.Text = "Booked value by event month (confirmed and later)"
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11; ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True
        ch.HasLegend = False; ch.ChartGroups(1).GapWidth = 70
        ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = hexrgb(LINE); ch.Axes(2).TickLabels.NumberFormat = USD0
        co.Name = "chtBooked"
        ws.Range("B64:T64").Merge()
        style(ws.Range("B64"), fg=MUTED, size=8).Formula = (
            '="Checks: "&ChecksPassed&" / "&ChecksTotal&" PASS   ·   Log: "&COUNTA(Log!$A$6:$A$20000)&" lines, "&COUNTIF(Log!$F$6:$F$20000,"BLOCKED")&" refused   ·   Sample project, fictional data"')
        fc = ws.Range("B64").FormatConditions.Add(2, Formula1="=ChecksPassed<ChecksTotal"); fc.Font.Color = hexrgb(BAD)
        ps = ws.PageSetup; ps.PrintArea = "$A$1:$U$65"; ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1

        # ------------------------------------------------------------------ How to use
        ws = W["How to use"]
        page(ws, "How to use", "Excel for Windows with macros enabled. Each person signs in with their own PIN.", xl)
        ws.Columns("B").ColumnWidth = 130
        for i, t in enumerate([
                "1. Sign in (top right of Home). Staff can book, check out, check in, invoice and record payments; a Manager can also cancel and lock a period.",
                "2. New booking: pick the customer, type the dates (yyyy-mm-dd), add items. The form shows how many are free for those dates while you type.",
                "   Save as enquiry holds no stock. Confirm booking holds it; if anything does not fit, the booking is refused and the reason is shown.",
                "3. On the day: pick the booking on Home, Check out. When it comes back: type damaged or missing quantities next to its lines, then Check in.",
                "   Late returns pay the late fee per item per day after the grace day; damaged or missing items are charged at replacement cost and leave the stock.",
                "4. Invoice + PDF numbers the invoice and saves the PDF in the Invoices folder. Record payment with the amount received; at 0 balance it shows Paid.",
                "5. Availability shows 21 days of free stock from any start date. Catalog: change quantities, prices, late fees and replacement costs (yellow).",
                "6. Month end: a Manager locks the period; after that nothing can be recorded on those dates. The Log keeps every attempt, allowed or refused.",
                "Checks must all say PASS. Sheets are protected without a password: it stops mistakes, it is not security. A PIN keeps honest people honest.",
        ], 6):
            ws.Cells(i, 2).Value = t

        for n in ("Availability", "Bookings", "Lines", "Payments", "Log", "Users", "Checks", "Settings", "Invoice"):
            button(W[n], "Home", "GoHome", W[n].Range("A1").Left + 700, 8, 60, 20)

        # ------------------------------------------------------------------ VBA
        vbp = wb.VBProject
        for nm in ("modCore", "modButtons"):
            code = (SRC / f"{nm}.bas").read_text(encoding="utf-8")
            code = "\n".join(l for l in code.splitlines() if not l.startswith("Attribute VB_Name"))
            m = vbp.VBComponents.Add(1); m.Name = nm
            m.CodeModule.AddFromString(code)
        vbp.VBComponents(wb.CodeName).CodeModule.AddFromString((SRC / "ThisWorkbook.code.vba").read_text(encoding="utf-8"))
        form.build(vbp, "frmSignIn", "Sign in", 296, 192, form.SIGNIN, (SRC / "frmSignIn.code.vba").read_text(encoding="utf-8"))
        form.build(vbp, "frmBooking", "New booking", 648, 540, form.BOOKING, (SRC / "frmBooking.code.vba").read_text(encoding="utf-8"))
        wb.SaveAs(str(OUT), FileFormat=52)
        watch(xl)
        run = lambda *a: xl.Run(*a)
        # PIN hashes, written by the workbook's own function
        for i, u in enumerate(USERS[1:], 6):
            W["Users"].Cells(i, 4).Value = run("PinHash", u[0], u[3])
        W["Users"].Range("D6:D20").NumberFormat = "@"

        # ------------------------------------------------------------------ replay the sample history through the macros
        ITEM = {r[0]: r for r in CAT[1:]}
        BY = {b["ref"]: b for b in HIST["bookings"]}
        PIN = {u[0]: u[3] for u in USERS[1:]}
        no, minute, t0 = {}, {}, time.time()
        for e in HIST["events"]:
            d = e["date"]; minute[d] = minute.get(d, 0) + 9
            run("SetClock", f"{d} {9 + minute[d] // 60:02d}:{minute[d] % 60:02d}")
            run("ReplayAs", e["user"])
            dt = SER(d); a = e["action"]; exp = e.get("expect", "")
            if a == "create":
                b = BY[e["ref"]]
                spec = "|".join(f"{it}~{q}" for it, q in b["lines"])
                res = run("NewBooking", b["customer"], SER(b["event"]), SER(b["out"]), SER(b["due"]), spec, float(b["delivery"]), "", bool(e["confirm"]), dt)
                assert res.startswith("B-"), (e, res); no[e["ref"]] = res
            elif a == "create_over":
                spec = "|".join(f"{it}~{q}" for it, q in e["lines"])
                res = run("NewBooking", e["customer"], SER(e["event"]), SER(e["out"]), SER(e["due"]), spec, 0.0, "", True, dt)
            elif a == "confirm": res = run("ConfirmBooking", no[e["ref"]], dt)
            elif a == "deposit":
                r = run("BookRow", no[e["ref"]]); res = run("RecordPayment", no[e["ref"]], float(W["Bookings"].Cells(r, 17).Value), dt)
            elif a == "cancel": res = run("CancelBooking", no[e["ref"]], dt, e["reason"])
            elif a == "checkout": res = run("CheckOut", no[e["ref"]], dt)
            elif a == "checkin": res = run("CheckIn", no[e["ref"]], dt, "|".join(f"{it}~{q}" for it, q in e["damaged"]))
            elif a == "invoice": res = run("MakeInvoice", no[e["ref"]], dt)
            elif a == "pay":
                r = run("BookRow", no[e["ref"]]); res = run("RecordPayment", no[e["ref"]], float(W["Bookings"].Cells(r, 19).Value), dt)
            elif a == "pay_backdated": res = run("RecordPayment", no[e["ref"]], float(e["amount"]), SER(e["on"]))
            elif a == "lock": res = run("LockPeriod", SER(e["through"]), dt)
            elif a == "login_bad": res = run("SignIn", e["user"], "0000")
            if exp == "BLOCKED":
                assert str(res).startswith("BLOCKED"), (e, res)
            elif a not in ("create",):
                assert res == "", (e, res)
        run("SetClock", "")
        run("ReplayAs", "")                           # nobody signed in when the file is opened (no log line)
        print(f"replayed {len(HIST['events'])} events in {time.time() - t0:.0f}s, {len(no)} bookings")
        (ROOT / "data" / "numbers.json").write_text(json.dumps(no, indent=1))

        # ------------------------------------------------------------------ finish: frozen as-of date, a selected booking
        xl.Run("SheetsOpen")
        wb.Names("AsOfDate").RefersToRange.Value = SER(SET["as_of"])
        W["Home"].Range("C14").Value = no[next(e["ref"] for e in HIST["events"] if e["action"] == "checkout"
                                              and not any(x.get("ref") == e["ref"] and x["action"] == "checkin" for x in HIST["events"]))]
        W["Home"].Range("B6").Value = "Sign in to start (Sign in button, top right)"
        inv_show = no[next(e["ref"] for e in reversed(HIST["events"]) if e["action"] == "checkin" and e["damaged"])]
        W["Invoice"].Range("K1").Value = inv_show
        for n, last_col in (("Bookings", "Y"), ("Lines", "P"), ("Payments", "F"), ("Log", "G")):
            lr = W[n].Cells(W[n].Rows.Count, 1).End(-4162).Row
            W[n].Range(f"A5:{last_col}{lr}").AutoFilter(1)
        xl.Run("SheetsLock")
        for n in ("Checks", "Availability", "Catalog", "Customers", "How to use", "Calc"):
            W[n].Protect(DrawingObjects=True, Contents=True)
        for n in reversed(SHEETS[:-1]):
            W[n].Activate(); W[n].Range("A1").Select()
        W["Home"].Activate(); W["Home"].Range("C14").Select()
        xl.ScreenUpdating = True
        xl.Calculation = -4105
        xl.CalculateFull()
        wb.Save()
        print("saved", OUT.name, "| checks", W["Checks"].Range(f"D{last + 2}").Value, "/", len(chk),
              "| log lines", W["Log"].Cells(W["Log"].Rows.Count, 1).End(-4162).Row - 5)
        for k in range(7, last + 1):
            print("  ", W["Checks"].Cells(k, 4).Value, W["Checks"].Cells(k, 3).Value, W["Checks"].Cells(k, 2).Value)
        wb.Close(False)
    finally:
        xl.Quit()
