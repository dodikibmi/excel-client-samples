"""Builds Quote-Invoice-Builder.xlsm with real Excel (COM), then replays the sample history through the workbook's
own macros, so every number, PDF and log line in the sample was produced by the tool itself.

Sheets
  Home        dark-gold cockpit: selected document + action buttons, 6 KPI cards, quoted/invoiced/collected chart,
              "needs attention" list, last 10 documents, data checks
  Document    the printable quote / invoice (A4) for any document number; this page becomes the PDF
  Register    one row per document (frozen at issue): totals, status, dates, PDF, link quote <-> invoice
  Lines       every document line
  Customers   yellow = typed: contact, email, payment terms, tax exempt
  Products    yellow = typed: category, unit, list price, minimum price, max discount
  Log         every create / edit / issue / block / status change, with time and user
  Checks      data checks, all must say PASS
  Settings    company, tax, numbering, quote validity, discount limit, PDF / email mode, as-of date
  How to use
  Calc        (hidden) chart and list helpers
VBA: modCore (numbering, totals, checks, PDF, email, status, log), modButtons (Home buttons), frmDocument (UserForm).

    python scripts/build_workbook.py            (needs pywin32; run under tools/with_lock.py)
"""
import csv, json, os, sys, time
from datetime import date, datetime
from pathlib import Path
import win32com.client as win32

sys.path.insert(0, str(Path(__file__).parent))
from vbom import vba_access
from watchdog import watch

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Quote-Invoice-Builder.xlsm"
SRC = ROOT / "src"
SET = json.loads((ROOT / "data" / "settings.json").read_text())
HIST = json.loads((ROOT / "data" / "history.json").read_text())
CUST = list(csv.reader(open(ROOT / "data" / "customers.csv", encoding="utf-8")))
PROD = list(csv.reader(open(ROOT / "data" / "products.csv", encoding="utf-8")))


def rgb(c):
    return c[0] + c[1] * 256 + c[2] * 65536


BG, PAGE, CARD, CARD2, LINE = (11, 15, 19), (18, 23, 29), (26, 33, 41), (32, 40, 50), (42, 51, 61)
GOLD, GOLD_LT, IVORY, MUTED = (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163)
JADE, COPPER, SKY = (79, 163, 140), (181, 117, 74), (127, 167, 201)
GOOD, BAD, AMBER = (91, 191, 138), (229, 115, 95), (217, 164, 65)
INPUT_DK = (52, 44, 28)
WHITE, INK, YELLOW, PAPER, SOFT = (255, 255, 255), (38, 44, 52), (255, 242, 204), (251, 249, 244), (232, 226, 212)
GREEN_BG, GREEN_FG, RED_BG, RED_FG = (198, 239, 206), (0, 97, 0), (255, 199, 206), (156, 0, 6)
GEORGIA, BODY = "Georgia", "Segoe UI"
SHEETS = ["Home", "Document", "Register", "Lines", "Customers", "Products", "Log", "Checks", "Settings", "How to use", "Calc"]
BRAND = "   ".join(" ".join(w.upper()) for w in SET["company"].split())

REG_COLS = [("Document", 13), ("Type", 8), ("Date", 11), ("Cust. ID", 8), ("Customer", 26), ("Email", 28), ("Lines", 6),
            ("Subtotal", 11), ("Disc %", 7), ("Discount", 10), ("Tax rate", 8), ("Tax", 10), ("Total", 12),
            ("Status", 9), ("Valid until / due", 13), ("Sent", 11), ("Accepted / paid", 13), ("Linked", 13),
            ("PDF file", 34), ("Issue check (drafts)", 40), ("Notes", 22),
            ("Status now", 10), ("Days late", 8), ("Lines check", 9), ("Needs attention", 40), ("Quote link", 9)]
LIN_COLS = [("Document", 13), ("Line", 5), ("Code", 8), ("Description", 34), ("Unit", 6), ("Qty", 7), ("Unit price", 10),
            ("Disc %", 7), ("Amount", 11), ("Key", 16), ("Amount check", 10), ("Min price", 9), ("Doc status", 10),
            ("Under min", 10)]
MAXR = 2000


def style(rng, fg=None, bg=None, size=None, bold=None, name=None, italic=None, h=None, v=None, fmt=None, wrap=None):
    if fg is not None: rng.Font.Color = rgb(fg)
    if bg is not None: rng.Interior.Color = rgb(bg)
    if size is not None: rng.Font.Size = size
    if bold is not None: rng.Font.Bold = bold
    if name is not None: rng.Font.Name = name
    if italic is not None: rng.Font.Italic = italic
    if h is not None: rng.HorizontalAlignment = {"l": -4131, "c": -4108, "r": -4152}[h]
    if v is not None: rng.VerticalAlignment = {"t": -4160, "c": -4108, "b": -4107}[v]
    if fmt is not None: rng.NumberFormat = fmt
    if wrap is not None: rng.WrapText = wrap
    return rng


def border(rng, color, weight=2, edges=(7, 8, 9, 10)):
    for e in edges:
        b = rng.Borders(e); b.LineStyle = 1; b.Weight = weight; b.Color = rgb(color)


def name(wb, nm, ref):
    wb.Names.Add(Name=nm, RefersTo=ref)


def title(ws, text, sub, light=True):
    ws.Cells.Font.Name = BODY; ws.Cells.Font.Size = 10
    style(ws.Range("A1"), fg=GOLD, size=8, bold=True).Value = BRAND
    style(ws.Range("A2"), fg=INK if light else IVORY, size=18, name=GEORGIA).Value = text
    style(ws.Range("A3"), fg=MUTED, size=9, italic=True).Value = sub
    ws.Rows(2).RowHeight = 28


def home_button(ws, left=None):
    left = left if left is not None else ws.Range("A1").Left + 560
    b = ws.Shapes.AddShape(5, left, 6, 78, 20)
    b.Fill.ForeColor.RGB = rgb(CARD); b.Line.ForeColor.RGB = rgb(GOLD)
    tf = b.TextFrame2.TextRange; tf.Text = "< Home"; tf.Font.Size = 9; tf.Font.Fill.ForeColor.RGB = rgb(GOLD)
    b.TextFrame2.VerticalAnchor = 3; tf.ParagraphFormat.Alignment = 2
    b.OnAction = "GoHome"; b.Name = "btnHome"


def table_sheet(ws, cols, header_row=5):
    for i, (h, w) in enumerate(cols, 1):
        ws.Columns(i).ColumnWidth = w
        ws.Cells(header_row, i).Value = h
    hr = ws.Range(ws.Cells(header_row, 1), ws.Cells(header_row, len(cols)))
    style(hr, fg=IVORY, bg=CARD, bold=True, size=9, wrap=True, v="c")
    ws.Rows(header_row).RowHeight = 30
    ws.Activate()
    ws.Range(ws.Cells(header_row + 1, 1), ws.Cells(header_row + 1, 1)).Select()
    xl.ActiveWindow.FreezePanes = False
    ws.Cells(header_row + 1, 2).Select()
    xl.ActiveWindow.FreezePanes = True
    xl.ActiveWindow.DisplayGridlines = False


def button(ws, text, macro, left, top, w, h, primary=False, size=10):
    b = ws.Shapes.AddShape(5, left, top, w, h)
    b.Fill.ForeColor.RGB = rgb(GOLD if primary else CARD2)
    b.Line.ForeColor.RGB = rgb(GOLD)
    b.Line.Weight = 0.75
    tf = b.TextFrame2.TextRange
    tf.Text = text; tf.Font.Size = size; tf.Font.Name = BODY; tf.Font.Bold = primary
    tf.Font.Fill.ForeColor.RGB = rgb(BG if primary else GOLD_LT)
    b.TextFrame2.VerticalAnchor = 3; tf.ParagraphFormat.Alignment = 2
    b.TextFrame2.MarginLeft = 2; b.TextFrame2.MarginRight = 2
    b.OnAction = macro
    b.Name = "btn" + macro
    return b


# ===================================================================================================== build
if OUT.exists():
    OUT.unlink()
with vba_access():
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False; xl.ScreenUpdating = False
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < len(SHEETS):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, n in enumerate(SHEETS, 1):
            wb.Worksheets(i).Name = n
        W = {n: wb.Worksheets(n) for n in SHEETS}

        # ------------------------------------------------------------------ Settings
        ws = W["Settings"]
        title(ws, "Settings", "Yellow cells are yours to change. Everything else reads them.")
        ws.Columns("A").ColumnWidth = 30; ws.Columns("B").ColumnWidth = 46; ws.Columns("C").ColumnWidth = 60
        rows = [("CoName", "Company name", SET["company"], None, "Printed on every quote and invoice"),
                ("CoAddress", "Address", SET["address"], None, ""),
                ("CoPhone", "Phone", SET["phone"], None, ""),
                ("CoEmail", "Email (sender)", SET["email"], None, "Used as From: in the email files"),
                ("Cur", "Currency symbol", SET["currency"], None, ""),
                ("TaxLabel", "Tax name", SET["tax_label"], None, "e.g. VAT, GST, Sales tax"),
                ("TaxRate", "Tax rate", SET["tax_rate"], "0.0%", "Customers marked tax exempt pay 0%"),
                ("QuoteValidDays", "Quote valid for (days)", SET["quote_valid_days"], "0", "Valid until = quote date + this"),
                ("MaxDocDisc", "Max document discount", SET["max_doc_discount"], "0%", "Above this, a document cannot be issued"),
                ("QuotePrefix", "Quote number prefix", SET["quote_prefix"], "@", "Numbers: prefix + year + 4 digits, e.g. Q-2026-0001"),
                ("InvPrefix", "Invoice number prefix", SET["invoice_prefix"], "@", ""),
                ("PdfFolder", "PDF folder", SET["pdf_folder"], "@", "A folder name next to this file, or a full path"),
                ("EmailMode", "Email when issuing", SET["email_mode"], "@", "Outlook draft = opens the mail in Outlook, ready to send; .eml file = an unsent email with the PDF attached, in the Outbox folder; Off = PDF only"),
                ("OutboxFolder", "Outbox folder (.eml)", SET["outbox_folder"], "@", ""),
                ("DeskUser", "Your name (log and email)", SET["user"], "@", "Written in the Log and under each email"),
                ("AsOfInput", "Dashboard as-of date", (date.fromisoformat(SET["as_of"]) - date(1899, 12, 30)).days, "d mmm yyyy", "Empty = today. Fixed here so the sample always shows the same numbers"),
                ("BankLine", "Bank details (invoices)", SET["bank"], "@", ""),
                ("TermsNote", "Payment note (invoices)", SET["terms_note"], "@", "")]
        style(ws.Range("A5:C5"), fg=IVORY, bg=CARD, bold=True).Value = ("Setting", "Value", "What it does")
        for i, (nm, label, val, fmt, note) in enumerate(rows, 6):
            ws.Cells(i, 1).Value = label
            c = ws.Cells(i, 2)
            if fmt: c.NumberFormat = fmt
            c.Value = val
            style(c, bg=YELLOW, h="l")
            style(ws.Cells(i, 3), fg=MUTED, size=9, wrap=True).Value = note
            name(wb, nm, f"=Settings!$B${i}")
        r = 6 + len(rows)
        ws.Cells(r, 1).Value = "As-of date in use"
        ws.Cells(r, 2).Formula = "=IF(AsOfInput=\"\",TODAY(),AsOfInput)"
        ws.Cells(r, 2).NumberFormat = "d mmm yyyy"; ws.Cells(r, 2).HorizontalAlignment = -4131
        name(wb, "AsOf", f"=Settings!$B${r}")
        em = 6 + [x[0] for x in rows].index("EmailMode")
        dv = ws.Cells(em, 2).Validation; dv.Delete(); dv.Add(3, 1, 1, "Outlook draft,.eml file,Off")
        border(ws.Range(f"A5:C{r}"), SOFT, 2, (7, 8, 9, 10, 12))

        # ------------------------------------------------------------------ Customers / Products
        ws = W["Customers"]
        title(ws, "Customers", "Yellow = typed. Payment terms set the invoice due date; tax exempt customers pay 0% tax.")
        table_sheet(ws, [(h, w) for h, w in zip(CUST[0], (11, 30, 18, 34, 14, 12, 11))])
        ws.Range(ws.Cells(6, 1), ws.Cells(5 + len(CUST) - 1, 7)).Value = [
            [r[0], r[1], r[2], r[3], r[4], int(r[5]), r[6]] for r in CUST[1:]]
        style(ws.Range(f"A6:G{5 + len(CUST) - 1 + 20}"), bg=YELLOW)
        dv = ws.Range("G6:G200").Validation; dv.Delete(); dv.Add(3, 1, 1, "No,Yes")
        ws.Range("F6:F200").NumberFormat = "0"

        ws = W["Products"]
        title(ws, "Products & prices", "Yellow = typed. A line priced under the minimum, or discounted above the max, blocks issuing.")
        table_sheet(ws, [(h, w) for h, w in zip(PROD[0], (9, 18, 38, 7, 11, 11, 12))])
        ws.Range(ws.Cells(6, 1), ws.Cells(5 + len(PROD) - 1, 7)).Value = [
            [r[0], r[1], r[2], r[3], float(r[4]), float(r[5]), int(r[6])] for r in PROD[1:]]
        style(ws.Range(f"A6:G{5 + len(PROD) - 1 + 20}"), bg=YELLOW)
        ws.Range("E6:F300").NumberFormat = "#,##0.00"; ws.Range("G6:G300").NumberFormat = "0"

        # ------------------------------------------------------------------ Register / Lines / Log
        ws = W["Register"]
        title(ws, "Register: every quote and invoice", "Written by the macros only. Totals are frozen when a document is issued; a price change later never alters it. Filter with the arrows.")
        table_sheet(ws, REG_COLS)
        for col in ("C", "O", "P", "Q"):
            ws.Range(f"{col}6:{col}{MAXR}").NumberFormat = "d mmm yyyy"
        for col in ("H", "J", "L", "M"):
            ws.Range(f"{col}6:{col}{MAXR}").NumberFormat = "#,##0.00"
        ws.Range(f"I6:I{MAXR}").NumberFormat = "0%"; ws.Range(f"K6:K{MAXR}").NumberFormat = "0%"
        style(ws.Range("V5:Z5"), bg=(60, 50, 30))
        st = ws.Range(f"V6:V{MAXR}")
        for txt, fg, bg in (("Overdue", RED_FG, RED_BG), ("Expired", (120, 90, 0), (255, 235, 156)), ("Paid", GREEN_FG, GREEN_BG),
                            ("Draft", (90, 90, 90), (230, 230, 230))):
            fc = st.FormatConditions.Add(1, 3, f'="{txt}"'); fc.Font.Color = rgb(fg); fc.Interior.Color = rgb(bg)
        fc = ws.Range(f"X6:X{MAXR}").FormatConditions.Add(1, 4, '="OK"'); fc.Font.Color = rgb(RED_FG); fc.Interior.Color = rgb(RED_BG)

        ws = W["Lines"]
        title(ws, "Lines: every line of every document", "Written by the macros only. Amount = qty x unit price x (1 - discount), rounded to the cent.")
        table_sheet(ws, LIN_COLS)
        ws.Range(f"F6:F{MAXR * 4}").NumberFormat = "General"
        ws.Range(f"G6:G{MAXR * 4}").NumberFormat = "#,##0.00"; ws.Range(f"I6:I{MAXR * 4}").NumberFormat = "#,##0.00"
        ws.Range(f"L6:L{MAXR * 4}").NumberFormat = "#,##0.00"; ws.Range(f"H6:H{MAXR * 4}").NumberFormat = "0%"
        style(ws.Range("J5:N5"), bg=(60, 50, 30))

        ws = W["Log"]
        title(ws, "Log: who did what, when", "Every create, edit, issue, blocked issue, status change and conversion. Written by the macros only.")
        table_sheet(ws, [("When", 17), ("User", 12), ("Action", 13), ("Document", 13), ("Detail", 110)])
        ws.Range(f"A6:A{MAXR * 6}").NumberFormat = "d mmm yyyy hh:mm"
        fc = ws.Range(f"C6:C{MAXR * 6}").FormatConditions.Add(1, 3, '="Issue blocked"'); fc.Font.Color = rgb(RED_FG); fc.Interior.Color = rgb(RED_BG)


        # ------------------------------------------------------------------ Document (printable)
        ws = W["Document"]
        ws.Cells.Font.Name = BODY; ws.Cells.Font.Size = 10; ws.Cells.Font.Color = rgb(INK)
        for col, w in zip("ABCDEFGHIJ", (2, 4, 40, 9, 7, 7, 11, 8, 13, 2)):
            ws.Columns(col).ColumnWidth = w
        ws.Columns("K").ColumnWidth = 3; ws.Columns("L").ColumnWidth = 18; ws.Columns("M").ColumnWidth = 20
        style(ws.Range("L2"), fg=MUTED, size=9).Value = "Show document:"
        style(ws.Range("M2"), bg=YELLOW, bold=True, h="c")
        name(wb, "CurDoc", "=Document!$M$2")
        ws.Range("M2").Locked = False   # the drop-down list is set by the macros (RefreshDocList)
        style(ws.Range("L3"), fg=MUTED, size=8, italic=True, wrap=True).Value = "(not printed: helper cells for the page)"
        ws.Range("L4").Value = "Register row"; ws.Range("M4").Formula = '=IFERROR(MATCH(CurDoc,Register!$A:$A,0),"")'
        style(ws.Range("L4:M40"), fg=MUTED, size=8)
        R = lambda col: f'IF($M$4="","",INDEX(Register!${col}:${col},$M$4)&"")' if col in ("E", "F", "U", "N", "A") else             f'IF($M$4="","",INDEX(Register!${col}:${col},$M$4))'
        isinv = '(INDEX(Register!$B:$B,$M$4)="Invoice")'

        style(ws.Range("B2"), fg=INK, size=22, name=GEORGIA).Formula = "=CoName"
        style(ws.Range("H2"), fg=GOLD, size=26, name=GEORGIA, h="r").Formula = f'=IF($M$4="","",IF({isinv},"INVOICE","QUOTE"))'
        ws.Range("H2:I2").Merge()
        style(ws.Range("B3"), fg=MUTED, size=9).Formula = '=CoAddress&"   ·   "&CoPhone&"   ·   "&CoEmail'
        border(ws.Range("B4:I4"), GOLD, -4138, (9,))
        ws.Rows(2).RowHeight = 34; ws.Rows(4).RowHeight = 6
        style(ws.Range("B6"), fg=GOLD, size=8, bold=True).Value = "B I L L   T O"
        style(ws.Range("B7"), size=12, bold=True).Formula = "=" + R("E")
        crow = 'MATCH(INDEX(Register!$D:$D,$M$4),Customers!$A:$A,0)'
        ws.Range("B8").Formula = f'=IFERROR(INDEX(Customers!$C:$C,{crow})&"","")'
        ws.Range("B9").Formula = "=" + R("F")
        ws.Range("B10").Formula = f'=IFERROR(INDEX(Customers!$E:$E,{crow})&"","")'
        for c in ("B8", "B9", "B10"):
            ws.Range(c).Font.Color = rgb((90, 96, 104))
        meta = [("G6", "Number", "=" + R("A"), "General"),
                ("G7", "Date", "=" + R("C"), "d mmm yyyy"),
                ("G8", f'=IF($M$4="","",IF({isinv},"Due date","Valid until"))', "=" + R("O"), "d mmm yyyy"),
                ("G9", "Status", "=" + R("N"), "General"),
                ("G10", f'=IF($M$4="","",IF(INDEX(Register!$R:$R,$M$4)="","",IF({isinv},"Quote ref.","Invoice")))',
                 f'=IF($M$4="","",IF(INDEX(Register!$R:$R,$M$4)="","",INDEX(Register!$R:$R,$M$4)))', "General")]
        for cell, lab, f, fmt in meta:
            c = ws.Range(cell)
            if lab.startswith("="): c.Formula = lab
            else: c.Value = lab
            style(c, fg=MUTED, size=9, h="r")
            rr = int(cell[1:]); ws.Range(f"H{rr}:I{rr}").Merge(); v = ws.Range(f"H{rr}")
            v.NumberFormat = fmt; v.Formula = f
            style(v, bold=True, h="r")
        fc = ws.Range("H9").FormatConditions.Add(1, 3, '="Draft"'); fc.Font.Color = rgb(BAD)
        style(ws.Range("B11"), fg=BAD, bold=True, size=9).Formula = (
            f'=IF($M$4="","Pick a document number in the yellow cell on the right",'
            f'IF(INDEX(Register!$N:$N,$M$4)="Draft","DRAFT  ·  not issued yet",""))')
        hdr = ["#", "Description", "Code", "Qty", "Unit", "Unit price", "Disc", "Amount"]
        for i, h in enumerate(hdr):
            ws.Cells(12, 2 + i).Value = h
        style(ws.Range("B12:I12"), fg=IVORY, bg=CARD, bold=True, size=9, v="c")
        for c in ("E12", "G12", "H12", "I12"):
            ws.Range(c).HorizontalAlignment = -4152
        ws.Rows(12).RowHeight = 22
        ws.Range("L12").Value = "Line row in Lines"
        for i in range(1, 16):
            r = 12 + i
            ws.Range(f"M{r}").Formula = f'=IF(CurDoc="","",IFERROR(MATCH(CurDoc&"|{i}",Lines!$J:$J,0),""))'
            L = lambda col: f'=IF($M{r}="","",INDEX(Lines!${col}:${col},$M{r}))'
            ws.Range(f"B{r}").Formula = f'=IF($M{r}="","",{i})'
            ws.Range(f"C{r}").Formula = L("D"); ws.Range(f"D{r}").Formula = L("C")
            ws.Range(f"E{r}").Formula = L("F"); ws.Range(f"F{r}").Formula = L("E")
            ws.Range(f"G{r}").Formula = L("G")
            ws.Range(f"H{r}").Formula = f'=IF($M{r}="","",IF(INDEX(Lines!$H:$H,$M{r})=0,"",INDEX(Lines!$H:$H,$M{r})))'
            ws.Range(f"I{r}").Formula = f'=IF($M{r}="","",ROUND(E{r}*G{r}*(1-N(H{r})),2))'
            if i % 2 == 0:
                ws.Range(f"B{r}:I{r}").Interior.Color = rgb(PAPER)
            ws.Rows(r).RowHeight = 18
        ws.Range("E13:E27").NumberFormat = "General"; ws.Range("G13:G27").NumberFormat = "#,##0.00"
        ws.Range("H13:H27").NumberFormat = "0%"; ws.Range("I13:I27").NumberFormat = "#,##0.00"
        ws.Range("B13:B27").HorizontalAlignment = -4108; ws.Range("F13:F27").Font.Color = rgb(MUTED)
        border(ws.Range("B27:I27"), SOFT, 2, (9,))
        tot = [(29, '="Subtotal"', "=SUM(I13:I27)"),
               (30, f'=IF($M$4="","Discount","Discount "&TEXT(INDEX(Register!$I:$I,$M$4),"0%"))',
                f'=IF($M$4="","",-ROUND(I29*INDEX(Register!$I:$I,$M$4),2))'),
               (31, f'=IF($M$4="",TaxLabel,TaxLabel&" "&TEXT(INDEX(Register!$K:$K,$M$4),"0%"))',
                f'=IF($M$4="","",ROUND((I29+I30)*INDEX(Register!$K:$K,$M$4),2))'),
               (32, '="TOTAL  "&Cur', "=IF($M$4=\"\",\"\",I29+I30+I31)")]
        for r, lab, f in tot:
            ws.Range(f"F{r}:H{r}").Merge()
            style(ws.Range(f"F{r}"), fg=MUTED, h="r").Formula = lab
            style(ws.Range(f"I{r}"), h="r", fmt="#,##0.00").Formula = f
        style(ws.Range("F32:I32"), fg=GOLD_LT, bg=CARD, bold=True, size=12, v="c")
        ws.Rows(32).RowHeight = 26
        ws.Range("L32").Value = "Register total"; ws.Range("M32").Formula = "=" + R("M")
        ws.Range("L33").Value = "Page = register?"
        ws.Range("M33").Formula = '=IF($M$4="","",IF(ABS(I32-M32)<0.005,"matches register","MISMATCH"))'
        name(wb, "DocViewCheck", "=Document!$M$33")
        style(ws.Range("B34"), fg=GOLD, size=8, bold=True).Value = "N O T E S"
        ws.Range("B35:I35").Merge(); ws.Range("B35").Formula = "=" + R("U"); ws.Range("B35").WrapText = True
        ws.Range("B37:I38").Merge()
        style(ws.Range("B37"), fg=(90, 96, 104), size=9, wrap=True, v="t").Formula = (
            f'=IF($M$4="","",IF({isinv},TermsNote&"  "&BankLine,'
            f'"This quote is valid until "&TEXT(INDEX(Register!$O:$O,$M$4),"d mmmm yyyy")&". To accept, reply to the email or call "&CoPhone&"."))')
        ws.Rows(37).RowHeight = 16; ws.Rows(38).RowHeight = 16
        border(ws.Range("B39:I39"), SOFT, 2, (9,))
        ws.Range("B40:I40").Merge()
        style(ws.Range("B40"), fg=MUTED, size=8, h="c").Formula = '="Thank you for your business   ·   "&CoName'
        ps = ws.PageSetup
        ps.PrintArea = "$A$1:$J$41"; ps.Orientation = 1; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
        ps.LeftMargin = xl.CentimetersToPoints(1.2); ps.RightMargin = xl.CentimetersToPoints(1.2)
        ps.TopMargin = xl.CentimetersToPoints(1.4); ps.BottomMargin = xl.CentimetersToPoints(1.2)
        ps.CenterHorizontally = True
        try:
            ps.PaperSize = 9  # A4
        except Exception:
            pass
        ws.Activate(); xl.ActiveWindow.DisplayGridlines = False
        home_button(ws, ws.Range("L1").Left + 160)

        # ------------------------------------------------------------------ Calc (hidden helpers)
        ws = W["Calc"]
        ws.Range("A1").Value = "Helpers for Home (chart and lists). Hidden; nothing to type here."
        ws.Range("A3:D3").Value = ("Month", "Quoted (issued)", "Invoiced (issued)", "Collected")
        for i in range(4):
            r = 4 + i
            ws.Range(f"A{r}").Formula = f"=DATE(YEAR(AsOf),MONTH(AsOf)-{3 - i},1)"
            ws.Range(f"A{r}").NumberFormat = "mmm yyyy"
            rng = f'Register!$C$6:$C${MAXR},">="&$A{r},Register!$C$6:$C${MAXR},"<"&EDATE($A{r},1)'
            ws.Range(f"B{r}").Formula = f'=SUMIFS(Register!$M$6:$M${MAXR},Register!$B$6:$B${MAXR},"Quote",Register!$N$6:$N${MAXR},"<>Draft",{rng})'
            ws.Range(f"C{r}").Formula = f'=SUMIFS(Register!$M$6:$M${MAXR},Register!$B$6:$B${MAXR},"Invoice",Register!$N$6:$N${MAXR},"<>Draft",{rng})'
            ws.Range(f"D{r}").Formula = (f'=SUMIFS(Register!$M$6:$M${MAXR},Register!$B$6:$B${MAXR},"Invoice",Register!$N$6:$N${MAXR},"Paid",'
                                         f'Register!$Q$6:$Q${MAXR},">="&$A{r},Register!$Q$6:$Q${MAXR},"<"&EDATE($A{r},1))')
        ws.Range("B4:D7").NumberFormat = "#,##0"
        ws.Range("F3").Value = "k"; ws.Range("G3").Value = "Register row needing attention"; ws.Range("H3").Value = "Recent row"
        for k in range(1, 13):
            r = 3 + k
            ws.Range(f"F{r}").Value = k
            ws.Range(f"G{r}").Formula = f'=IFERROR(AGGREGATE(15,6,ROW(Register!$A$6:$A${MAXR})/(Register!$Y$6:$Y${MAXR}<>""),F{r}),"")'
            ws.Range(f"H{r}").Formula = f'=IF(5+COUNTA(Register!$A$6:$A${MAXR})-F{r}+1>=6,5+COUNTA(Register!$A$6:$A${MAXR})-F{r}+1,"")'
        ws.Visible = 0

        # ------------------------------------------------------------------ Checks
        ws = W["Checks"]
        title(ws, "Checks", "Recomputed by formulas from the Register and Lines. Every check must say PASS.")
        ws.Columns("A").ColumnWidth = 4; ws.Columns("B").ColumnWidth = 58; ws.Columns("C").ColumnWidth = 10
        ws.Columns("D").ColumnWidth = 10; ws.Columns("E").ColumnWidth = 70
        style(ws.Range("A5:E5"), fg=IVORY, bg=CARD, bold=True).Value = ("#", "Check", "Found", "Result", "Why it matters")
        RG = f"Register!$A$6:$A${MAXR}"
        checks = [
            ("Register subtotal = sum of the document's lines", f'=COUNTIF(Register!$X$6:$X${MAXR},"Lines")',
             "A total that does not match its own lines is the classic copy-paste invoice error"),
            ("Total = subtotal - discount + tax, on every document", f'=COUNTIF(Register!$X$6:$X${MAXR},"Total")', ""),
            ("Line amount = qty x price x (1 - discount), rounded", f'=COUNTIF(Lines!$K$6:$K${MAXR * 4},"Wrong")', ""),
            ("No document number used twice", f'=SUMPRODUCT(({RG}<>"")*(COUNTIF({RG},{RG})>1))',
             "Duplicate invoice numbers are a tax-audit finding"),
            ("Every line belongs to a document in the register", f'=COUNTIF(Lines!$M$6:$M${MAXR * 4},"Missing")', ""),
            ("Every document has at least one line", f'=SUMPRODUCT(({RG}<>"")*(COUNTIF(Lines!$A$6:$A${MAXR * 4},{RG})=0))', ""),
            ("No issued line priced under the product's minimum", f'=COUNTIF(Lines!$N$6:$N${MAXR * 4},"Under min")',
             "Discounts beyond the floor need the owner's OK; the macro blocks them before issuing"),
            ("Every invoice comes from one accepted quote, and back", f'=COUNTIF(Register!$Z$6:$Z${MAXR},"Bad")',
             "No invoice without an agreed quote; no quote invoiced twice"),
            ("Every issued document has its PDF file", f'=COUNTIFS(Register!$N$6:$N${MAXR},"<>Draft",{RG},"<>",Register!$S$6:$S${MAXR},"")', ""),
            ("Every document has a 'Created' entry in the Log",
             f'=SUMPRODUCT(({RG}<>"")*(COUNTIFS(Log!$D$6:$D${MAXR * 6},{RG},Log!$C$6:$C${MAXR * 6},"Created")=0))', ""),
            ("Printed page total = register total (document shown)", '=IF(DocViewCheck="MISMATCH",1,0)', ""),
        ]
        for i, (txt, f, why) in enumerate(checks, 1):
            r = 5 + i
            ws.Cells(r, 1).Value = i; ws.Cells(r, 2).Value = txt; ws.Cells(r, 3).Formula = f
            ws.Cells(r, 4).Formula = f'=IF(C{r}=0,"PASS","FAIL")'
            style(ws.Cells(r, 5), fg=MUTED, size=9).Value = why
        last = 5 + len(checks)
        d = ws.Range(f"D6:D{last}")
        fc = d.FormatConditions.Add(1, 3, '="PASS"'); fc.Font.Color = rgb(GREEN_FG); fc.Interior.Color = rgb(GREEN_BG)
        fc = d.FormatConditions.Add(1, 3, '="FAIL"'); fc.Font.Color = rgb(RED_FG); fc.Interior.Color = rgb(RED_BG)
        d.HorizontalAlignment = -4108; d.Font.Bold = True
        border(ws.Range(f"A5:E{last}"), SOFT, 2, (7, 8, 9, 10, 12))
        ws.Cells(last + 2, 2).Value = "Checks passed"
        ws.Cells(last + 2, 4).Formula = f'=COUNTIF(D6:D{last},"PASS")'
        ws.Cells(last + 3, 2).Value = "Checks in total"
        ws.Cells(last + 3, 4).Value = len(checks)
        style(ws.Range(f"B{last + 2}:D{last + 3}"), bold=True)
        name(wb, "ChecksPassed", f"=Checks!$D${last + 2}")
        name(wb, "ChecksTotal", f"=Checks!$D${last + 3}")
        ws.Activate(); xl.ActiveWindow.DisplayGridlines = False

        # ------------------------------------------------------------------ How to use
        ws = W["How to use"]
        title(ws, "How to use", "Five buttons do the whole job. Macros must be enabled (File > Options > Trust Center, or 'Enable Content').")
        ws.Columns("A").ColumnWidth = 120
        steps = [
            "1. Settings: type your company, tax, quote validity, discount limit, numbering prefixes and how emails are made.",
            "2. Customers and Products: type or paste your own (yellow cells). Minimum price and max discount are your price floor.",
            "3. Home > + New quote: pick the customer, then category > product, quantity, price, discount, Add line.",
            "   The form shows the totals and the checks live. Save draft keeps it; Issue makes the PDF and the email.",
            "4. Issue is blocked (and the reason is shown and logged) when: no email for the customer, a price under the minimum,",
            "   a line discount above the product's max, a document discount above the limit, a number used twice, no lines.",
            "5. When the customer answers: pick the quote in 'Selected document' and click Accepted or Declined.",
            "6. Convert to invoice: makes a draft invoice with the quote's own lines and prices; Issue sends it with its due date.",
            "7. Mark paid when the money arrives. Overdue invoices and quotes about to expire show on Home under Needs attention.",
            "8. Issued documents are frozen. To change one, create a new document; the Log keeps who did what and when.",
            "9. Reset (on Home) clears all documents to start your own book; customers, products and settings stay.",
            "",
            "Email modes: 'Outlook draft' opens the mail in Outlook with the PDF attached (you press Send);",
            "'.eml file' writes an unsent email file in the Outbox folder (double-click opens it ready to send); 'Off' = PDF only.",
            "Sheets are protected without a password so nobody types over the register by accident; Review > Unprotect if you must.",
            "Works in Excel 2016 and later for Windows. The sample history (Aug-Nov 2026) was created by these same macros, with email Off.",
        ]
        for i, t in enumerate(steps, 5):
            ws.Cells(i, 1).Value = t
        ws.Activate(); xl.ActiveWindow.DisplayGridlines = False

        # ------------------------------------------------------------------ Home (dark-gold cockpit)
        ws = W["Home"]
        ws.Activate()
        ws.Cells.Font.Name = BODY; ws.Cells.Font.Size = 10
        ws.Cells.Interior.Color = rgb(BG); ws.Cells.Font.Color = rgb(IVORY)
        xl.ActiveWindow.DisplayGridlines = False; xl.ActiveWindow.DisplayHeadings = False
        ws.Columns("A").ColumnWidth = 2
        ws.Range("B:S").ColumnWidth = 9.3
        ws.Columns("T").ColumnWidth = 2
        style(ws.Range("B2"), fg=GOLD, size=8, bold=True).Value = BRAND
        style(ws.Range("B3"), fg=IVORY, size=24, name=GEORGIA).Value = "Quotes & invoices"
        ws.Rows(3).RowHeight = 36
        style(ws.Range("B4"), fg=MUTED, size=9).Formula = (
            '="As of "&TEXT(AsOf,"d mmm yyyy")&"   ·   "&COUNTIF(Register!$B:$B,"Quote")&" quotes, "&COUNTIF(Register!$B:$B,"Invoice")'
            '&" invoices   ·   every document numbered, checked before it is issued, and logged"')
        ws.Rows(5).RowHeight = 8; ws.Rows(6).RowHeight = 24; ws.Rows(7).RowHeight = 10
        nav = [("Document", "GoDocument"), ("Register", "GoRegister"), ("Lines", "GoLines"), ("Customers", "GoCustomers"),
               ("Products", "GoProducts"), ("Log", "GoLog"), ("Checks", "GoChecks"), ("Settings", "GoSettings"),
               ("How to use", "GoHelp")]
        x0 = ws.Range("B6").Left; y0 = ws.Range("B6").Top + 2
        for i, (t, m) in enumerate(nav):
            button(ws, t, m, x0 + i * 76, y0, 72, 20, size=9)
        # selected document + actions
        style(ws.Range("B8"), fg=GOLD, size=8, bold=True).Value = "S E L E C T E D   D O C U M E N T"
        ws.Range("B9:C9").Merge()
        style(ws.Range("B9"), fg=GOLD_LT, bg=INPUT_DK, size=12, bold=True, h="c", v="c")
        name(wb, "SelDoc", "=Home!$B$9")
        ws.Range("B9:C9").Locked = False
        border(ws.Range("B9:C9"), GOLD, 2)
        ws.Rows(9).RowHeight = 24
        sr = 'MATCH(SelDoc,Register!$A:$A,0)'
        ws.Range("D9:K9").Merge()
        style(ws.Range("D9"), fg=IVORY, size=10, v="c").Formula = (
            f'=IF(SelDoc="","   Pick a document number (drop-down) or select a row on Register",IFERROR("   "&INDEX(Register!$B:$B,{sr})&"  ·  "'
            f'&INDEX(Register!$E:$E,{sr})&"  ·  "&Cur&TEXT(INDEX(Register!$M:$M,{sr}),"#,##0.00")&"  ·  "&INDEX(Register!$V:$V,{sr})'
            f'&IF(INDEX(Register!$Y:$Y,{sr})="","","  ·  "&INDEX(Register!$Y:$Y,{sr})),"   not found"))')
        ws.Rows(10).RowHeight = 6; ws.Rows(11).RowHeight = 26; ws.Rows(12).RowHeight = 12
        y = ws.Range("B11").Top + 2
        button(ws, "+  New quote", "BtnNewQuote", ws.Range("B11").Left, y, 112, 24, primary=True, size=11)
        acts = [("Edit draft", "BtnEditDraft"), ("Issue: PDF + email", "BtnIssue"), ("Accepted", "BtnAccept"),
                ("Declined", "BtnDecline"), ("Convert to invoice", "BtnConvert"), ("Mark paid", "BtnPaid"),
                ("Show page", "BtnShowDoc"), ("Open PDF", "BtnOpenPdf")]
        x = ws.Range("B11").Left + 122
        for t, m in acts:
            wdt = 64 if len(t) < 10 else 104
            button(ws, t, m, x, y, wdt, 24, size=9)
            x += wdt + 6
        # KPI cards
        RGs = lambda c: f"Register!${c}$6:${c}${MAXR}"
        cards = [
            ("OPEN QUOTES", f'=SUMIFS({RGs("M")},{RGs("B")},"Quote",{RGs("V")},"Sent")', "#,##0",
             f'=COUNTIFS({RGs("B")},"Quote",{RGs("V")},"Sent")&" waiting for an answer"'),
            ("WIN RATE", f'=IFERROR(COUNTIFS({RGs("B")},"Quote",{RGs("N")},"Accepted")/(COUNTIFS({RGs("B")},"Quote",{RGs("N")},"Accepted")'
                         f'+COUNTIFS({RGs("B")},"Quote",{RGs("N")},"Declined")+COUNTIFS({RGs("B")},"Quote",{RGs("V")},"Expired")),0)', "0%",
             f'=COUNTIFS({RGs("B")},"Quote",{RGs("N")},"Accepted")&" won · "&COUNTIFS({RGs("B")},"Quote",{RGs("N")},"Declined")'
             f'&" declined · "&COUNTIFS({RGs("B")},"Quote",{RGs("V")},"Expired")&" expired"'),
            ("INVOICED THIS MONTH", f'=SUMIFS({RGs("M")},{RGs("B")},"Invoice",{RGs("N")},"<>Draft",{RGs("C")},">="&DATE(YEAR(AsOf),MONTH(AsOf),1),{RGs("C")},"<="&AsOf)', "#,##0",
             f'=COUNTIFS({RGs("B")},"Invoice",{RGs("N")},"<>Draft",{RGs("C")},">="&DATE(YEAR(AsOf),MONTH(AsOf),1),{RGs("C")},"<="&AsOf)&" invoices in "&TEXT(AsOf,"mmmm")'),
            ("UNPAID", f'=SUMIFS({RGs("M")},{RGs("B")},"Invoice",{RGs("N")},"Sent")', "#,##0",
             f'=COUNTIFS({RGs("B")},"Invoice",{RGs("N")},"Sent")&" invoices sent, not paid"'),
            ("OVERDUE", f'=SUMIFS({RGs("M")},{RGs("V")},"Overdue")', "#,##0",
             f'=COUNTIF({RGs("V")},"Overdue")&" invoices · oldest "&MAX(0,MAX({RGs("W")}))&" days late"'),
            ("NEEDS ATTENTION", f'=COUNTIF({RGs("Y")},"?*")', "0",
             f'=COUNTIF({RGs("Y")},"Blocked*")&" drafts blocked · "&COUNTIF({RGs("Y")},"Expires*")&" quotes expiring"'),
        ]
        for i, (lab, f, fmt, note) in enumerate(cards):
            c0 = 2 + i * 3
            box = ws.Range(ws.Cells(13, c0), ws.Cells(17, c0 + 2))
            box.Interior.Color = rgb(CARD)
            border(box, LINE, 2)
            ws.Range(ws.Cells(13, c0), ws.Cells(13, c0 + 2)).Merge()
            style(ws.Cells(13, c0), fg=MUTED, size=8, bold=True, h="l").Value = "  " + lab
            ws.Range(ws.Cells(14, c0), ws.Cells(15, c0 + 2)).Merge()
            v = ws.Cells(14, c0)
            v.NumberFormat = ("[$" + SET["currency"] + "]" + fmt) if fmt == "#,##0" else fmt
            v.Formula = f
            style(v, fg=GOLD_LT if i < 4 else (BAD if i == 4 else AMBER), size=22, name=GEORGIA, h="l", v="c")
            v.IndentLevel = 1
            ws.Range(ws.Cells(16, c0), ws.Cells(16, c0 + 2)).Merge()
            style(ws.Cells(16, c0), fg=MUTED, size=8, h="l").Formula = note
            ws.Cells(16, c0).IndentLevel = 1
        ws.Rows(13).RowHeight = 18; ws.Rows(14).RowHeight = 18; ws.Rows(15).RowHeight = 18; ws.Rows(16).RowHeight = 16
        ws.Rows(17).RowHeight = 6; ws.Rows(18).RowHeight = 12
        # chart (left) + needs attention (right)
        style(ws.Range("B19"), fg=GOLD, size=8, bold=True).Value = "Q U O T E D   ·   I N V O I C E D   ·   C O L L E C T E D   B Y   M O N T H"
        style(ws.Range("J19"), fg=GOLD, size=8, bold=True).Value = "N E E D S   A T T E N T I O N"
        top, left = ws.Range("B20").Top, ws.Range("B20").Left
        width, height = ws.Range("J20").Left - left - 10, ws.Range("B35").Top - top
        co = ws.ChartObjects().Add(left, top, width, height)
        ch = co.Chart
        ch.ChartType = 51
        ch.SetSourceData(W["Calc"].Range("A3:D7"))
        ch.PlotBy = 2
        ch.HasTitle = False
        ch.ChartArea.Format.Fill.ForeColor.RGB = rgb(CARD); ch.ChartArea.Format.Line.ForeColor.RGB = rgb(LINE)
        ch.PlotArea.Format.Fill.Visible = False
        ch.ChartArea.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(MUTED)
        ch.ChartArea.Format.TextFrame2.TextRange.Font.Size = 9
        ch.HasLegend = True; ch.Legend.Position = -4160
        for s, col in zip(range(1, 4), (GOLD, SKY, JADE)):
            se = ch.SeriesCollection(s)
            se.Format.Fill.ForeColor.RGB = rgb(col)
            se.HasDataLabels = True
            se.DataLabels().NumberFormat = '#,##0;;'
            se.DataLabels().Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(IVORY)
            se.DataLabels().Format.TextFrame2.TextRange.Font.Size = 8
        ch.ChartGroups(1).GapWidth = 70; ch.ChartGroups(1).Overlap = -5
        ax = ch.Axes(2); ax.HasMajorGridlines = True
        ax.MajorGridlines.Format.Line.ForeColor.RGB = rgb(LINE); ax.TickLabels.NumberFormat = "#,##0"
        ax.Format.Line.Visible = False
        ch.Axes(1).Format.Line.ForeColor.RGB = rgb(LINE)
        co.Name = "chtMonths"
        # attention list J..S rows 20..31
        hdr = [("J", "K", "Document"), ("L", "N", "Customer"), ("O", "R", "Why"), ("S", "S", "Amount")]
        for a, b, t in hdr:
            ws.Range(f"{a}20:{b}20").Merge()
            style(ws.Range(f"{a}20"), fg=MUTED, size=8, bold=True).Value = t
        border(ws.Range("J20:S20"), LINE, 2, (9,))
        for k in range(1, 13):
            r = 20 + k
            g = f"Calc!$G${3 + k}"
            for a, b, col, fmt in (("J", "K", "A", "General"), ("L", "N", "E", "General"), ("O", "R", "Y", "General"), ("S", "S", "M", "#,##0")):
                ws.Range(f"{a}{r}:{b}{r}").Merge()
                c = ws.Range(f"{a}{r}")
                c.NumberFormat = fmt
                c.Formula = f'=IF({g}="","",INDEX(Register!${col}:${col},{g}))'
                style(c, size=9, fg=GOLD_LT if col == "A" else (BAD if col == "Y" else IVORY))
            ws.Range(f"S{r}").HorizontalAlignment = -4152
            ws.Rows(r).RowHeight = 16
        fcr = ws.Range("O21:O32")
        fc = fcr.FormatConditions.Add(Type=9, String="Expires", TextOperator=2); fc.Font.Color = rgb(AMBER)
        fc = fcr.FormatConditions.Add(Type=9, String="Draft", TextOperator=2); fc.Font.Color = rgb(MUTED)
        # recent documents
        ws.Rows(34).RowHeight = 8; ws.Rows(35).RowHeight = 8
        style(ws.Range("B36"), fg=GOLD, size=8, bold=True).Value = "L A S T   1 0   D O C U M E N T S"
        cols = [("B", "C", "Document", "A", "General"), ("D", "D", "Type", "B", "General"), ("E", "F", "Date", "C", "d mmm yyyy"),
                ("G", "J", "Customer", "E", "General"), ("K", "L", "Total", "M", "#,##0.00"), ("M", "N", "Status now", "V", "General"),
                ("O", "P", "Valid / due", "O", "d mmm yyyy"), ("Q", "S", "Linked / PDF", "R", "General")]
        for a, b, t, _, _ in cols:
            ws.Range(f"{a}37:{b}37").Merge()
            style(ws.Range(f"{a}37"), fg=MUTED, size=8, bold=True, h="r" if t == "Total" else "l").Value = t
        border(ws.Range("B37:S37"), LINE, 2, (9,))
        for k in range(1, 11):
            r = 37 + k
            g = f"Calc!$H${3 + k}"
            for a, b, _, col, fmt in cols:
                ws.Range(f"{a}{r}:{b}{r}").Merge()
                c = ws.Range(f"{a}{r}")
                c.NumberFormat = fmt
                if col == "R":
                    c.Formula = f'=IF({g}="","",IF(INDEX(Register!$R:$R,{g})<>"",INDEX(Register!$R:$R,{g}),IF(INDEX(Register!$S:$S,{g})<>"","PDF saved","")))'
                else:
                    c.Formula = f'=IF({g}="","",INDEX(Register!${col}:${col},{g}))'
                style(c, size=9, fg=GOLD_LT if col == "A" else IVORY, h="l")
            ws.Range(f"K{r}").HorizontalAlignment = -4152
            ws.Rows(r).RowHeight = 16
        st = ws.Range("M38:M47")
        for txt, col in (("Overdue", BAD), ("Expired", AMBER), ("Paid", GOOD), ("Draft", MUTED)):
            fc = st.FormatConditions.Add(1, 3, f'="{txt}"'); fc.Font.Color = rgb(col)
        ws.Rows(48).RowHeight = 10
        ws.Range("B49:S49").Merge()
        style(ws.Range("B49"), fg=MUTED, size=8).Formula = (
            '="Data checks: "&ChecksPassed&" / "&ChecksTotal&" PASS   ·   Log: "&COUNTA(Log!$A$6:$A$20000)&" entries   ·   '
            'Customers: "&COUNTA(Customers!$A$6:$A$500)&"   ·   Products: "&COUNTA(Products!$A$6:$A$500)&"   ·   Sample project, fictional data"')
        fc = ws.Range("B49").FormatConditions.Add(2, Formula1="=ChecksPassed<ChecksTotal"); fc.Font.Color = rgb(BAD)
        ps = ws.PageSetup; ps.PrintArea = "$A$1:$T$50"; ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1

        for n in ("Register", "Lines", "Log", "Customers", "Products", "Settings", "Checks", "How to use"):
            home_button(W[n], W[n].Range("A1").Left + 640)
            W[n].Activate(); xl.ActiveWindow.DisplayGridlines = False

        # ------------------------------------------------------------------ VBA
        vbp = wb.VBProject
        for nm in ("modCore", "modButtons"):
            code = (SRC / f"{nm}.bas").read_text(encoding="utf-8")
            code = "\n".join(l for l in code.splitlines() if not l.startswith("Attribute VB_Name"))
            m = vbp.VBComponents.Add(1); m.Name = nm
            m.CodeModule.AddFromString(code)
        twb = vbp.VBComponents(wb.CodeName)
        twb.CodeModule.AddFromString((SRC / "ThisWorkbook.code.vba").read_text(encoding="utf-8"))
        import form
        form.build(vbp, (SRC / "frmDocument.code.vba").read_text(encoding="utf-8"))

        wb.SaveAs(str(OUT), FileFormat=52)

        # ------------------------------------------------------------------ replay the sample history through the macros
        import shutil
        for sub in (SET["pdf_folder"], SET["outbox_folder"]):
            shutil.rmtree(ROOT / sub, ignore_errors=True)
        W["Settings"].Range("B" + str(6 + [x[0] for x in rows].index("EmailMode"))).Value = "Off"
        watch(xl)
        xl.Calculation = -4135   # manual while replaying; IssueDocument recalculates before each PDF
        docs = {d["ref"]: d for d in HIST["documents"]}
        qno, ino, minute = {}, {}, {}
        run = lambda *a: xl.Run(*a)
        t0 = time.time()
        for e in HIST["events"]:
            d = date.fromisoformat(e["date"])
            minute[d] = minute.get(d, 0) + 7
            run("SetClock", f"{d.isoformat()} {9 + minute[d] // 60:02d}:{minute[d] % 60:02d}")
            dt = (d - date(1899, 12, 30)).days   # an Excel serial: a datetime would be shifted by the time zone
            doc = docs[e["ref"]]
            a = e["action"]
            if a == "create":
                spec = "|".join(f"{c}~{q}~{p}~{dd}" for c, q, p, dd in doc["lines"])
                res = run("NewDocument", "Quote", doc["customer"], dt, float(doc["doc_discount"]), doc["notes"], spec)
                assert not res.startswith("ERR"), (e, res)
                qno[e["ref"]] = res
            elif a == "issue":
                res = run("IssueDocument", qno[e["ref"]], dt); assert res == "", (e, res)
            elif a in ("accept", "decline"):
                res = run("SetStatus", qno[e["ref"]], "Accepted" if a == "accept" else "Declined", dt); assert res == "", (e, res)
            elif a == "convert":
                res = run("ConvertToInvoice", qno[e["ref"]], dt); assert res.startswith("INV"), (e, res)
                ino[e["ref"]] = res
            elif a == "issue_invoice":
                res = run("IssueDocument", ino[e["ref"]], dt); assert res == "", (e, res)
            elif a == "paid":
                res = run("SetStatus", ino[e["ref"]], "Paid", dt); assert res == "", (e, res)
        W["Settings"].Range("B" + str(6 + [x[0] for x in rows].index("EmailMode"))).Value = SET["email_mode"]
        # the two planted drafts: try to issue them once, so the Log shows the block and the reason
        lastday = HIST["events"][-1]["date"]
        for ref in ("d1", "d2"):
            run("SetClock", f"{lastday} 17:30")
            res = run("IssueDocument", qno[ref], (date.fromisoformat(lastday) - date(1899, 12, 30)).days)
            assert res.startswith("BLOCKED"), (ref, res)
        run("SetClock", "")
        print(f"replayed {len(HIST['events'])} events in {time.time() - t0:.0f}s:", len(qno), "quotes,", len(ino), "invoices")
        (ROOT / "data" / "numbers.json").write_text(json.dumps({"quotes": qno, "invoices": ino}, indent=1))

        # ------------------------------------------------------------------ finish: show a paid invoice, select the blocked draft
        paid_inv = max((ino[e["ref"]] for e in HIST["events"] if e["action"] == "paid"),
                       key=lambda n: len(docs[[r for r, v in ino.items() if v == n][0]]["lines"]))
        W["Document"].Range("M2").Value = paid_inv
        W["Home"].Range("B9").Value = qno["d2"]
        xl.Run("SheetsOpen")
        for n, last_col in (("Register", "Z"), ("Lines", "N"), ("Log", "E")):
            lr = W[n].Cells(W[n].Rows.Count, 1).End(-4162).Row
            W[n].Range(f"A5:{last_col}{lr}").AutoFilter(1)
        xl.Run("SheetsLock")
        for n in ("Checks", "How to use", "Calc"):
            W[n].Protect(DrawingObjects=True, Contents=True)
        for n in reversed(SHEETS[:-1]):
            W[n].Activate(); W[n].Range("A1").Select() if n not in ("Home",) else None
        W["Home"].Activate(); W["Home"].Range("B9").Select()
        xl.ScreenUpdating = True
        xl.Calculation = -4105
        xl.CalculateFull()
        wb.Save()
        print("saved", OUT.name, "checks", xl.Evaluate("ChecksPassed"), "/", xl.Evaluate("ChecksTotal"))
        wb.Close(False)
    finally:
        xl.Quit()
