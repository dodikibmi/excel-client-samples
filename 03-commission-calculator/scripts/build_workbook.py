"""Builds Sales-Commission-Calculator.xlsx with real Excel (COM). No macros.

Sheets
  Commission  pick a month: KPI cards, commission per sales rep (tiered bands + target bonus), held-back lines, checks
  Statement   pick a rep and a month: a printable commission statement with every order line and the working
  Orders      the only place the sales team types (yellow): drop-downs, validation; everything else is a locked formula
  Plan        prices, max discounts, reps and targets, commission bands, tax, bonus, method (the manager edits here)
  Test cases  17 hand-worked cases run through the same formulas; all must say PASS
  How to use

Run with a Python that has pywin32:  python scripts/build_workbook.py
"""
import csv, json
from datetime import date
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Sales-Commission-Calculator.xlsx"
PLAN = json.loads((ROOT / "data" / "plan.json").read_text())
ORDER_ROWS = 600  # pre-sized Orders table: a protected sheet cannot grow a table


def rgb(r, g, b):
    return r + g * 256 + b * 65536


NAVY, TEAL, ORANGE, GREY = (24, 45, 72), (0, 128, 128), (230, 126, 34), (110, 110, 110)
YELLOW = (255, 242, 204)
GREEN_BG, GREEN_FG, RED_BG, RED_FG = (198, 239, 206), (0, 97, 0), (255, 199, 206), (156, 0, 6)
SHEETS = ["Commission", "Statement", "Orders", "Plan", "Test cases", "How to use"]
serial = lambda d: (d - date(1899, 12, 30)).days


def comm(net, method):
    """Commission formula, used by Commission, Statement and Test cases so all three share one logic.
    Whole amount: all net sales at the rate of the band reached. Stepped: each band's slice at its own rate
    (written as net above each band start x the rate increase at that band = Bands[Step])."""
    return (f'IF({method}="Stepped",ROUND(SUMPRODUCT(({net}>Bands[From])*({net}-Bands[From])*Bands[Step]),2),'
            f'ROUND({net}*INDEX(Bands[Rate],MATCH({net},Bands[From],1)),2))')


# Order-line formulas: the same strings fill the Orders table and the OrderTests table on Test cases.
ORDER_CALC = [
    ("Unit price", '=IF([@Product]="","",IFERROR(INDEX(Products[Unit price],MATCH([@Product],Products[Product],0)),""))', "#,##0.00"),
    ("Gross", '=IF([@[Unit price]]="","",ROUND(N([@Qty])*[@[Unit price]],2))', "#,##0.00"),
    ("Discount", '=IF([@Gross]="","",ROUND([@Gross]*N([@[Discount %]]),2))', "#,##0.00"),
    ("Net", '=IF([@Gross]="","",[@Gross]-[@Discount])', "#,##0.00"),
    ("Tax", '=IF([@Net]="","",ROUND([@Net]*TaxRate,2))', "#,##0.00"),
    ("Total", '=IF([@Net]="","",[@Net]+[@Tax])', "#,##0.00"),
    ("Month", '=IF([@Date]="","",DATE(YEAR([@Date]),MONTH([@Date]),1))', "mmm yyyy"),
    ("Status", '=IF(AND([@Date]="",[@[Sales rep]]="",[@Product]=""),"",'
               'IF([@Date]="","Missing date",'
               'IF(COUNTIF(Reps[Sales rep],[@[Sales rep]])=0,"Unknown sales rep",'
               'IF(COUNTIF(Products[Product],[@Product])=0,"Unknown product",'
               'IF(N([@Qty])<1,"Qty must be 1 or more",'
               'IF(N([@[Discount %]])>INDEX(Products[Max discount],MATCH([@Product],Products[Product],0))+0.000001,'
               '"Discount above "&TEXT(INDEX(Products[Max discount],MATCH([@Product],Products[Product],0)),"0%")&" max","OK"))))))', "@"),
]
ORDER_HELPERS = [  # Orders only, hidden: feed the held-back list and the statement without FILTER (works in Excel 2016)
    ("Held", '=IF(AND([@Status]<>"",[@Status]<>"OK",OR([@Month]=SelMonth,[@Month]="")),1,0)'),
    ("Held line", '=IF([@Held]=1,SUM(INDEX([Held],1):[@Held]),"")'),
    ("In statement", '=IF(AND([@Status]="OK",[@[Sales rep]]=StmtRep,[@Month]=StmtMonth),1,0)'),
    ("Statement line", '=IF([@[In statement]]=1,SUM(INDEX([In statement],1):[@[In statement]]),"")'),
]
INPUTS = ["Date", "Order no", "Sales rep", "Customer", "Product", "Qty", "Discount %"]


def font(rng, size=None, bold=None, color=None):
    if size: rng.Font.Size = size
    if bold is not None: rng.Font.Bold = bold
    if color: rng.Font.Color = rgb(*color)


def heading(ws, cell, text):
    ws.Range(cell).Value = text
    font(ws.Range(cell), 13, True, NAVY)


def title_band(ws, last_col, formula, sub):
    ws.Range(f"A1:{last_col}2").Interior.Color = rgb(*NAVY)
    ws.Range("B1").Formula = formula
    font(ws.Range("B1"), 18, True, (255, 255, 255))
    ws.Range("B2").Value = sub
    font(ws.Range("B2"), 10, False, (200, 215, 230))
    ws.Rows(1).RowHeight = 34; ws.Rows(2).RowHeight = 18; ws.Rows(3).RowHeight = 8


def selector(ws, label_cell, rng, name, wb, source, fmt=None, value=None):
    lab = ws.Range(label_cell)
    font(lab, 11, True, NAVY); lab.HorizontalAlignment = -4152
    sel = ws.Range(rng)
    sel.Merge()
    sel.Interior.Color = rgb(*YELLOW)
    sel.Borders.LineStyle = 1; sel.Borders.Color = rgb(191, 143, 0)
    font(sel, 12, True, NAVY); sel.HorizontalAlignment = -4108
    sel.Locked = False
    first = sel.Cells(1, 1)
    first.Validation.Add(3, 1, 1, source)
    if fmt:
        first.NumberFormat = fmt
    if value is not None:
        first.Value2 = value
    wb.Names.Add(name, f"='{ws.Name}'!{first.Address}")


def table_header(ws, row, cols_titles):
    for c, t in cols_titles:
        h = ws.Range(f"{c}{row}"); h.Value = t
        font(h, 10, True, (255, 255, 255)); h.Interior.Color = rgb(*NAVY)
        h.HorizontalAlignment = -4131 if t in ("Sales rep", "Region", "Order no", "Customer", "Product", "Problem", "Case", "Method", "Date") else -4152


def check_cell(cell):
    font(cell, 10, True); cell.HorizontalAlignment = -4108
    for val, bg, fg in (("PASS", GREEN_BG, GREEN_FG), ("FAIL", RED_BG, RED_FG)):
        fc = cell.FormatConditions.Add(1, 3, f'="{val}"'); fc.Interior.Color = rgb(*bg); fc.Font.Color = rgb(*fg)


def main():
    if OUT.exists():
        OUT.unlink()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < len(SHEETS):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, n in enumerate(SHEETS, 1):
            wb.Worksheets(i).Name = n
            wb.Worksheets(i).Cells.Font.Name = "Calibri"
        xl.Calculation = -4135  # manual while building
        build_plan(wb)
        # selector names first, so the Orders helper formulas can refer to them
        wb.Worksheets("Commission").Range("C4").Value2 = serial(date(2026, 10, 1))
        wb.Names.Add("SelMonth", "=Commission!$C$4")
        wb.Worksheets("Statement").Range("C4").Value = PLAN["reps"][4][0]
        wb.Worksheets("Statement").Range("G4").Value2 = serial(date(2026, 10, 1))
        wb.Names.Add("StmtRep", "=Statement!$C$4")
        wb.Names.Add("StmtMonth", "=Statement!$G$4")
        build_orders(wb)
        build_tests(wb)
        build_commission(wb)
        build_statement(wb)
        build_howto(wb)
        xl.Calculation = -4105
        xl.CalculateFull()
        for name in ("Commission", "Statement", "Orders", "Test cases"):
            ws = wb.Worksheets(name)
            ws.Protect(Password="", DrawingObjects=True, Contents=True, Scenarios=True, AllowFiltering=True)
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            ws.Range("A1").Select()
        wb.Worksheets("Commission").Activate()
        xl.ActiveWindow.Zoom = 90
        wb.SaveAs(str(OUT), 51)
        print("saved", OUT.name)
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        xl.Quit()


# ------------------------------------------------------------------------------------------------- Plan
def add_table(ws, top_left, headers, rows, name, fmts):
    r0 = ws.Range(top_left).Row; c0 = ws.Range(top_left).Column
    for j, h in enumerate(headers):
        ws.Cells(r0, c0 + j).Value = h
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row):
            ws.Cells(r0 + i, c0 + j).Value = v
    rng = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + len(rows), c0 + len(headers) - 1))
    lo = ws.ListObjects.Add(1, rng, None, 1)
    lo.Name = name
    lo.TableStyle = "TableStyleMedium2"
    for col, fmt in fmts.items():
        lo.ListColumns(col).DataBodyRange.NumberFormat = fmt
    return lo


def build_plan(wb):
    p = wb.Worksheets("Plan")
    p.Range("A1").Value = "Commission plan & lists"
    font(p.Range("A1"), 16, True, NAVY)
    p.Range("A2").Value = "The sales manager edits this sheet: prices, discounts, reps, targets, bands. Every other sheet follows."
    font(p.Range("A2"), 10, False, GREY)
    settings = [("Company", PLAN["company"], "Company", "@"), ("Sales tax rate", PLAN["tax_rate"], "TaxRate", "0.00%"),
                ("Target bonus", PLAN["target_bonus"], "TargetBonus", "#,##0"),
                ("Discount cap (any product)", PLAN["discount_cap"], "DiscountCap", "0%"),
                ("Commission method", "Whole amount", "Method", "@")]
    p.Range("A4:B4").Value = ("Setting", "Value")
    font(p.Range("A4:B4"), 10, True)
    for i, (label, val, nm, fmt) in enumerate(settings, 5):
        p.Cells(i, 1).Value = label
        c = p.Cells(i, 2)
        c.NumberFormat = fmt
        c.Value = val
        c.Interior.Color = rgb(*YELLOW)
        wb.Names.Add(nm, f"=Plan!$B${i}")
    p.Range("Q3").Value = "Methods"
    p.Range("Q4").Value = "Whole amount"; p.Range("R4").Value = "All net sales are paid at the rate of the band reached"
    p.Range("Q5").Value = "Stepped"; p.Range("R5").Value = "Each band's slice of net sales is paid at that band's own rate"
    wb.Names.Add("MethodList", "=Plan!$Q$4:$Q$5")
    p.Range("B9").Validation.Add(3, 1, 1, "=MethodList")
    p.Range("C9").Value = "Whole amount or Stepped (see the note on the right)"
    p.Range("C6").Value = "Paid when a rep's net sales reach the monthly target"
    p.Range("C7").Value = "No discount above this is accepted in Orders; each product also has its own maximum"
    font(p.Range("C5:C9"), 9, False, GREY)

    add_table(p, "A12", ["Product", "Category", "Unit price", "Max discount"], PLAN["products"], "Products",
              {"Unit price": "#,##0.00", "Max discount": "0%"})
    add_table(p, "F12", ["Sales rep", "Region", "Monthly target"], PLAN["reps"], "Reps", {"Monthly target": "#,##0"})
    bands = add_table(p, "J12", ["From", "Rate"], PLAN["bands"], "Bands", {"From": "#,##0", "Rate": "0.0%"})
    col = bands.ListColumns.Add(); col.Name = "Step"
    col.DataBodyRange.Formula2 = "=IF(ROW()-ROW($K$12)=1,[@Rate],[@Rate]-INDEX([Rate],ROW()-ROW($K$12)-1))"  # $K$12 = header
    col.DataBodyRange.NumberFormat = "+0.0%"
    col = bands.ListColumns.Add(); col.Name = "Order ok"
    col.DataBodyRange.Formula2 = '=IF(ROW()-ROW($J$12)=1,--([@From]=0),--(INDEX([From],ROW()-ROW($J$12)-1)<[@From]))'
    p.Range("A11").Value = "Products"; p.Range("F11").Value = "Sales reps"; p.Range("J11").Value = "Commission bands (monthly net sales)"
    for c in ("A11", "F11", "J11"):
        font(p.Range(c), 12, True, NAVY)
    p.Range("J19").Value = "Step = rate increase at that band (used by the Stepped method). Order ok = bands go up from 0."
    font(p.Range("J19"), 9, False, GREY)
    p.Range("Q11").Value = "Months"
    font(p.Range("Q11"), 12, True, NAVY)
    font(p.Range("Q3"), 12, True, NAVY)
    for i in range(24):
        c = p.Range(f"Q{12 + i}")
        c.Value2 = serial(date(2026 + i // 12, i % 12 + 1, 1))
        c.NumberFormat = "mmm yyyy"
        c.HorizontalAlignment = -4131
    wb.Names.Add("MonthList", "=Plan!$Q$12:$Q$35")
    bands.Resize(p.Range(f"J12:M{12 + len(PLAN['bands'])}"))  # keep the table to its own rows and columns
    for col, w in zip("ABCDEFGHIJKLMNOPQR", (28, 12, 11, 12, 2, 16, 10, 14, 2, 10, 8, 8, 9, 2, 2, 2, 14, 60)):
        p.Columns(col).ColumnWidth = w
    p.Columns("C").ColumnWidth = 11
    p.Range("C5:C9").WrapText = False


# ------------------------------------------------------------------------------------------------- Orders
def build_orders(wb):
    o = wb.Worksheets("Orders")
    o.Range("A1").Value = "Orders - type or paste order lines in the yellow columns; everything else is calculated"
    font(o.Range("A1"), 14, True, NAVY)
    o.Range("A2").Value = ("Status must say OK to be paid. Lines with a problem are held back and listed on the Commission sheet "
                           "until fixed. Room for 600 lines.")
    font(o.Range("A2"), 10, False, GREY)
    with open(ROOT / "data" / "orders.csv", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    hdr = INPUTS + [n for n, _, _ in ORDER_CALC] + [n for n, _ in ORDER_HELPERS]
    o.Range(o.Cells(4, 1), o.Cells(4, len(hdr))).Value = tuple(hdr)
    rng = o.Range(o.Cells(4, 1), o.Cells(4 + ORDER_ROWS, len(hdr)))
    lo = o.ListObjects.Add(1, rng, None, 1)
    lo.Name = "Orders"
    lo.TableStyle = "TableStyleLight9"
    data = [(serial(date.fromisoformat(r["Date"])), r["Order no"], r["Sales rep"], r["Customer"], r["Product"],
             int(r["Qty"]), float(r["Discount %"])) for r in rows]
    o.Range(o.Cells(5, 1), o.Cells(4 + len(data), 7)).Value2 = data
    for name, f, fmt in ORDER_CALC:
        lc = lo.ListColumns(name)
        lc.DataBodyRange.Formula = f
        lc.DataBodyRange.NumberFormat = fmt
        lc.DataBodyRange.Interior.Color = rgb(247, 249, 252)
    for name, f in ORDER_HELPERS:
        lo.ListColumns(name).DataBodyRange.Formula = f
    style_input_columns(o, lo, 5, 4 + ORDER_ROWS)
    st = lo.ListColumns("Status").DataBodyRange
    fc = st.FormatConditions.Add(2, None, f'=(${col_letter(lo, "Status")}5<>"")*(${col_letter(lo, "Status")}5<>"OK")')
    fc.Interior.Color = rgb(*RED_BG); fc.Font.Color = rgb(*RED_FG); fc.Font.Bold = True
    widths = {"Date": 11, "Order no": 10, "Sales rep": 16, "Customer": 25, "Product": 26, "Qty": 6, "Discount %": 10,
              "Unit price": 10, "Gross": 11, "Discount": 10, "Net": 11, "Tax": 9, "Total": 11, "Month": 10, "Status": 24}
    for i, h in enumerate(hdr, 1):
        o.Columns(i).ColumnWidth = widths.get(h, 8)
    first_helper = len(INPUTS) + len(ORDER_CALC) + 1
    o.Range(o.Cells(1, first_helper), o.Cells(1, len(hdr))).EntireColumn.Hidden = True
    o.Activate(); o.Range("A5").Select(); wb.Application.ActiveWindow.FreezePanes = True


def col_letter(lo, name):
    a = lo.ListColumns(name).DataBodyRange.Cells(1, 1).Address  # like $O$5
    return a.split("$")[1]


def style_input_columns(ws, lo, r1, r2):
    """Yellow, unlocked input cells with drop-downs and validation (pasted values bypass validation: Status catches them)."""
    body = lambda n: lo.ListColumns(n).DataBodyRange
    for n in INPUTS:
        b = body(n)
        b.Locked = False
        b.Interior.Color = rgb(*YELLOW)
    rules = {
        "Date": (4, 1, "=MIN(MonthList)", "=MAX(MonthList)+30", "dd mmm yyyy", "Order date", "A date in the plan's months."),
        "Sales rep": (3, 1, "=Reps[Sales rep]", None, None, "Sales rep", "Pick from the list (reps live on Plan)."),
        "Product": (3, 1, "=Products[Product]", None, None, "Product", "Pick from the list; price comes from Plan."),
        "Qty": (1, 7, "1", None, "0", "Quantity", "Whole number, 1 or more."),
        "Discount %": (2, 1, "0", "=DiscountCap", "0%", "Discount", "0% to the cap; each product has its own max (Plan)."),
    }
    for n, (typ, op, f1, f2, fmt, title, msg) in rules.items():
        b = body(n)
        v = b.Validation
        v.Delete()
        src = f1
        if typ == 3:  # list from a table column: a defined name keeps it working in every Excel version
            nm = {"Sales rep": "RepList", "Product": "ProductList"}[n]
            ws.Parent.Names.Add(nm, f1)
            src = "=" + nm
        if f2:
            v.Add(typ, 1, op, src, f2)
        else:
            v.Add(typ, 1, op, src)
        v.InputTitle = title; v.InputMessage = msg
        v.ErrorTitle = title; v.ErrorMessage = msg
        if fmt:
            b.NumberFormat = fmt


# ------------------------------------------------------------------------------------------------- Commission
REP_ROWS = range(14, 24)  # up to 10 reps


def build_commission(wb):
    c = wb.Worksheets("Commission")
    c.Columns("A").ColumnWidth = 2
    for col, w in zip("BCDEFGHIJKL", (17, 9, 9, 13, 12, 11, 11, 12, 10, 13, 2)):
        c.Columns(col).ColumnWidth = w
    title_band(c, "L", '="Sales commission   |   "&Company&"   |   "&TEXT(SelMonth,"mmmm yyyy")',
               "Pick a month. Commission is paid on order lines with Status OK; the plan (bands, bonus, method) lives on Plan.")
    c.Range("B4").Value = "Month"
    selector(c, "B4", "C4:D4", "SelMonth", wb, "=MonthList", "mmm yyyy", serial(date(2026, 10, 1)))
    c.Range("F4").Value = '="Method: "&Method'
    c.Range("F4").Formula = '="Method: "&Method&"   |   target bonus "&TEXT(TargetBonus,"#,##0")'
    font(c.Range("F4"), 10, True, GREY)
    c.Rows(5).RowHeight = 8

    cards = [("B", "C", "NET SALES PAID ON", "=E24", "#,##0", '=TEXT(D24,"#,##0")&" order lines"', None),
             ("D", "E", "COMMISSION + BONUS", "=K24", "#,##0.00", '="bonus "&TEXT(J24,"#,##0")', None),
             ("F", "G", "PAY AS % OF SALES", "=IFERROR(K24/E24,0)", "0.0%", '="method: "&Method', None),
             ("H", "I", "REPS ON TARGET", '=COUNTIF(G14:G23,">=1")&" of "&COUNTA(Reps[Sales rep])', "@",
              '="bonus "&TEXT(TargetBonus,"#,##0")&" each"', None),
             ("J", "K", "LINES HELD BACK", "=SUM(Orders[Held])", "0",
              '=IF(SUM(Orders[Held])=0,"nothing to fix","fix them in Orders")', "$J$8")]  # CF cannot use table refs
    for a, b, label, value, fmt, note, warn in cards:
        box = c.Range(f"{a}6:{b}10")
        box.Interior.Color = rgb(247, 249, 252)
        for edge in (7, 8, 9, 10):
            box.Borders(edge).LineStyle = 1; box.Borders(edge).Color = rgb(255, 255, 255); box.Borders(edge).Weight = -4138
        c.Range(f"{a}6:{b}6").Interior.Color = rgb(*TEAL)
        for row, val, size, color in ((7, label, 9, (90, 100, 115)), (8, value, 22, NAVY), (10, note, 9, GREY)):
            rng = c.Range(f"{a}{row}:{b}{row + (1 if row == 8 else 0)}")
            rng.Merge()
            cell = c.Range(f"{a}{row}")
            if val.startswith("="):
                cell.Formula = val
            else:
                cell.Value = val
            font(cell, size, True, color)
            rng.HorizontalAlignment = -4108; rng.VerticalAlignment = -4108
        c.Range(f"{a}8").NumberFormat = fmt
        if warn:
            for r in (8, 10):
                fc = c.Range(f"{a}{r}").FormatConditions.Add(2, None, f"={warn}>0"); fc.Font.Color = rgb(192, 0, 0)
    c.Rows(6).RowHeight = 5; c.Rows(8).RowHeight = 20; c.Rows(9).RowHeight = 16; c.Rows(11).RowHeight = 10

    heading(c, "B12", "Commission by sales rep, selected month")
    cols = list("BCDEFGHIJK")
    table_header(c, 13, zip(cols, ("Sales rep", "Region", "Lines", "Net sales", "Target", "Achieved", "Band rate",
                                     "Commission", "Bonus", "Total pay")))
    for n, r in enumerate(REP_ROWS, 1):
        b = f"$B{r}"
        f = {
            "B": f'=IF({n}>ROWS(Reps[Sales rep]),"",INDEX(Reps[Sales rep],{n}))',
            "C": f'=IF({b}="","",INDEX(Reps[Region],{n}))',
            "D": f'=IF({b}="","",COUNTIFS(Orders[Sales rep],{b},Orders[Month],SelMonth,Orders[Status],"OK"))',
            "E": f'=IF({b}="","",SUMIFS(Orders[Net],Orders[Sales rep],{b},Orders[Month],SelMonth,Orders[Status],"OK"))',
            "F": f'=IF({b}="","",INDEX(Reps[Monthly target],{n}))',
            "G": f'=IF({b}="","",IFERROR(E{r}/F{r},0))',
            "H": f'=IF({b}="","",INDEX(Bands[Rate],MATCH(E{r},Bands[From],1)))',
            "I": f'=IF({b}="","",{comm(f"E{r}", "Method")})',
            "J": f'=IF({b}="","",IF(E{r}>=F{r},TargetBonus,0))',
            "K": f'=IF({b}="","",I{r}+J{r})',
        }
        for col, formula in f.items():
            c.Range(f"{col}{r}").Formula = formula
        c.Range(f"B{r}:K{r}").Borders(9).LineStyle = 1; c.Range(f"B{r}:K{r}").Borders(9).Color = rgb(220, 220, 220)
    c.Range("B24").Value = "Total"
    for col in "DEFIJK":
        c.Range(f"{col}24").Formula = f"=SUM({col}14:{col}23)"
    c.Range("G24").Formula = "=IFERROR(E24/F24,0)"
    font(c.Range("B24:K24"), None, True)
    c.Range("B24:K24").Borders(8).LineStyle = 1
    for rng, fmt in (("D14:D24", "#,##0"), ("E14:F24", "#,##0.00"), ("G14:G24", "0.0%"), ("H14:H23", "0.0%"), ("I14:K24", "#,##0.00")):
        c.Range(rng).NumberFormat = fmt
    fc = c.Range("G14:G23").FormatConditions.Add(2, None, "=N($G14)>=1"); fc.Interior.Color = rgb(*GREEN_BG); fc.Font.Color = rgb(*GREEN_FG)
    fc = c.Range("J14:J23").FormatConditions.Add(2, None, "=N($J14)>0"); fc.Font.Color = rgb(*GREEN_FG); fc.Font.Bold = True

    # chart: net sales vs target per rep, and the plan next to it
    heading(c, "B26", "Net sales vs target")
    n_reps = len(PLAN["reps"])
    co = c.ChartObjects().Add(c.Range("B27").Left, c.Range("B27").Top, c.Range("B27:F27").Width, c.Range("B27:B41").Height)
    ch = co.Chart
    ch.ChartType = 57
    s1 = ch.SeriesCollection().NewSeries(); s1.Name = "Net sales"; s1.Values = c.Range(f"E14:E{13 + n_reps}"); s1.XValues = c.Range(f"B14:B{13 + n_reps}")
    s2 = ch.SeriesCollection().NewSeries(); s2.Name = "Target"; s2.Values = c.Range(f"F14:F{13 + n_reps}"); s2.XValues = c.Range(f"B14:B{13 + n_reps}")
    s1.Format.Fill.ForeColor.RGB = rgb(*TEAL); s2.Format.Fill.ForeColor.RGB = rgb(*ORANGE)
    ch.HasTitle = False
    ch.HasLegend = True; ch.Legend.Position = -4160
    ch.Axes(1).ReversePlotOrder = True
    ch.Axes(2).TickLabels.NumberFormat = '#,##0,"k"'
    ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = rgb(230, 230, 230)
    ch.ChartArea.Format.Line.Visible = False; ch.ChartArea.Font.Size = 9

    heading(c, "H26", "Commission plan")
    table_header(c, 27, (("H", "From"), ("I", "To"), ("J", "Rate"), ("K", "Paid on")))
    for k in range(1, 7):
        r = 27 + k
        c.Range(f"H{r}").Formula = f'=IF({k}>ROWS(Bands[From]),"",INDEX(Bands[From],{k}))'
        c.Range(f"I{r}").Formula = f'=IF(H{r}="","",IFERROR(INDEX(Bands[From],{k + 1})-0.01,"and up"))'
        c.Range(f"J{r}").Formula = f'=IF(H{r}="","",INDEX(Bands[Rate],{k}))'
        c.Range(f"K{r}").Formula = f'=IF(H{r}="","",IF(Method="Stepped","this slice","all sales"))'
        c.Range(f"H{r}:I{r}").NumberFormat = "#,##0"; c.Range(f"J{r}").NumberFormat = "0.0%"
        c.Range(f"I{r}:K{r}").HorizontalAlignment = -4152
    c.Range("H35").Formula = '="Method: "&Method'
    font(c.Range("H35"), 10, True, NAVY)
    c.Range("H36").Formula = "=INDEX(Plan!$R$4:$R$5,MATCH(Method,MethodList,0))"
    c.Range("H37").Formula = '="Target bonus "&TEXT(TargetBonus,"#,##0")&" when net sales reach the monthly target."'
    c.Range("H38").Value = "Commission is worked on net sales (after discount, before tax)."
    font(c.Range("H36:H38"), 9, False, GREY)

    # held-back lines
    heading(c, "B43", "Lines held back: not paid until fixed in Orders")
    hdr = (("B", "Order no"), ("C", "Sales rep"), ("E", "Product"), ("H", "Qty"), ("I", "Problem"))
    table_header(c, 44, hdr)
    for col in "DFGJK":
        c.Range(f"{col}44").Interior.Color = rgb(*NAVY)
    for k in range(1, 11):
        r = 44 + k
        pick = lambda col: f'=IFERROR(INDEX(Orders[{col}],MATCH({k},Orders[Held line],0)),"")'
        c.Range(f"B{r}").Formula = pick("Order no")
        c.Range(f"C{r}:D{r}").Merge(); c.Range(f"C{r}").Formula = pick("Sales rep")
        c.Range(f"E{r}:G{r}").Merge(); c.Range(f"E{r}").Formula = pick("Product")
        c.Range(f"H{r}").Formula = pick("Qty")
        c.Range(f"I{r}:K{r}").Merge(); c.Range(f"I{r}").Formula = pick("Status")
        font(c.Range(f"I{r}"), None, True, RED_FG)
        c.Range(f"B{r}:K{r}").Borders(9).LineStyle = 1; c.Range(f"B{r}:K{r}").Borders(9).Color = rgb(220, 220, 220)
    c.Range("B55").Formula = '=IF(SUM(Orders[Held])>10,"... and "&(SUM(Orders[Held])-10)&" more: filter Status in Orders","")'
    font(c.Range("B55"), 9, False, GREY)

    heading(c, "B57", "Data checks")
    checks = [("Every order line in the month is either paid or held back",
               '=IF(COUNTIFS(Orders[Month],SelMonth)=D24+COUNTIFS(Orders[Month],SelMonth,Orders[Held],1),"PASS","FAIL")'),
              ("Rep table = all OK net sales in the month",
               '=IF(ROUND(E24-SUMIFS(Orders[Net],Orders[Month],SelMonth,Orders[Status],"OK"),2)=0,"PASS","FAIL")'),
              ("Commission bands start at 0 and go up", '=IF(SUM(Bands[Order ok])=ROWS(Bands[Order ok]),"PASS","FAIL")'),
              ("All test cases pass (Test cases sheet)", '=IF(TestsFailed=0,"PASS","FAIL")')]
    for i, (label, f) in enumerate(checks, 58):
        c.Range(f"B{i}").Value = label
        c.Range(f"K{i}").Formula = f
        check_cell(c.Range(f"K{i}"))
    ps = c.PageSetup
    ps.Orientation = 1; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$L$61"
    for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
        setattr(ps, m, 18)


# ------------------------------------------------------------------------------------------------- Statement
STMT_LINES = range(19, 99)  # 80 lines


def build_statement(wb):
    s = wb.Worksheets("Statement")
    s.Columns("A").ColumnWidth = 2
    for col, w in zip("BCDEFGHIJ", (11, 10, 25, 26, 6, 10, 10, 12, 2)):
        s.Columns(col).ColumnWidth = w
    title_band(s, "J", '="Commission statement   |   "&StmtRep&"   |   "&TEXT(StmtMonth,"mmmm yyyy")',
               "")
    s.Range("B2").Formula = '=Company&"  -  pick a rep and a month, then print or save as PDF for the rep."'
    s.Range("B4").Value = "Sales rep"
    selector(s, "B4", "C4:D4", "StmtRep", wb, "=RepList", None, None)
    s.Range("F4").Value = "Month"
    selector(s, "F4", "G4:H4", "StmtMonth", wb, "=MonthList", "mmm yyyy", serial(date(2026, 10, 1)))
    s.Range("F4").HorizontalAlignment = -4152

    rows = [("Region", '=IFERROR(INDEX(Reps[Region],MATCH(StmtRep,Reps[Sales rep],0)),"not on Plan")', "@"),
            ("Monthly target", '=IFERROR(INDEX(Reps[Monthly target],MATCH(StmtRep,Reps[Sales rep],0)),0)', "#,##0.00"),
            ("Net sales (paid lines)", '=SUMIFS(Orders[Net],Orders[Sales rep],StmtRep,Orders[Month],StmtMonth,Orders[Status],"OK")', "#,##0.00"),
            ("Target achieved", "=IFERROR(D8/D7,0)", "0.0%"),
            ("Band reached", "=INDEX(Bands[Rate],MATCH(D8,Bands[From],1))", "0.0%"),
            ("Commission", "=" + comm("D8", "Method"), "#,##0.00"),
            ("Target bonus", "=IF(AND(D7>0,D8>=D7),TargetBonus,0)", "#,##0.00"),
            ("Total pay", "=D11+D12", "#,##0.00")]
    for i, (label, f, fmt) in enumerate(rows, 6):
        s.Range(f"B{i}").Value = label
        s.Range(f"D{i}").Formula = f
        s.Range(f"D{i}").NumberFormat = fmt
        s.Range(f"D{i}").HorizontalAlignment = -4152
        s.Range(f"B{i}:D{i}").Borders(9).LineStyle = 1; s.Range(f"B{i}:D{i}").Borders(9).Color = rgb(220, 220, 220)
    font(s.Range("B13:D13"), 12, True, NAVY)
    s.Range("B13:D13").Interior.Color = rgb(*YELLOW)

    s.Range("E6").Formula = '="How the commission was worked out ("&Method&")"'
    font(s.Range("E6"), 11, True, NAVY)
    for k in range(1, 7):
        r = 6 + k
        frm, nxt, rate = f"INDEX(Bands[From],{k})", f"IFERROR(INDEX(Bands[From],{k + 1}),$D$8)", f"INDEX(Bands[Rate],{k})"
        part = f"MAX(0,MIN($D$8,{nxt})-{frm})"
        stepped = (f'IF({k}>ROWS(Bands[From]),"",IF({part}<=0,"",TEXT({part},"#,##0.00")&" x "&TEXT({rate},"0.0%")&" = "'
                   f'&TEXT({part}*{rate},"#,##0.00")))')
        whole = '"All net sales "&TEXT($D$8,"#,##0.00")&" x "&TEXT($D$10,"0.0%")&" = "&TEXT($D$11,"#,##0.00")' if k == 1 else '""'
        s.Range(f"E{r}").Formula = f'=IF(Method="Stepped",{stepped},{whole})'
    s.Range("E13").Formula = '=IF(Method="Stepped","Commission (rounded once) = "&TEXT(D11,"#,##0.00"),"")&IF(D12>0,"   +  target bonus "&TEXT(D12,"#,##0.00"),"")'
    font(s.Range("E7:E13"), 10, False, (60, 60, 60))
    s.Range("B15").Formula = ('="Lines held back for this rep and month: "&COUNTIFS(Orders[Sales rep],StmtRep,Orders[Month],StmtMonth,'
                              'Orders[Status],"<>OK")&" (not paid until fixed in Orders)"')
    font(s.Range("B15"), 9, True, RED_FG)

    heading(s, "B17", "Order lines paid on")
    cols = list("BCDEFGHI")
    names = ("Date", "Order no", "Customer", "Product", "Qty", "Unit price", "Discount", "Net")
    table_header(s, 18, zip(cols, names))
    for k, r in enumerate(STMT_LINES, 1):
        for col, n in zip(cols, names):
            s.Range(f"{col}{r}").Formula = f'=IFERROR(INDEX(Orders[{n}],MATCH({k},Orders[Statement line],0)),"")'
    last = STMT_LINES[-1]
    s.Range(f"B{STMT_LINES[0]}:B{last}").NumberFormat = "dd mmm yyyy"
    s.Range(f"G{STMT_LINES[0]}:I{last}").NumberFormat = "#,##0.00"
    s.Range(f"B{STMT_LINES[0]}:B{last}").HorizontalAlignment = -4131
    s.Range(f"B{last + 1}").Value = "Total of the lines listed"
    s.Range(f"I{last + 1}").Formula = f"=SUM(I{STMT_LINES[0]}:I{last})"
    s.Range(f"I{last + 1}").NumberFormat = "#,##0.00"
    font(s.Range(f"B{last + 1}:I{last + 1}"), None, True)
    s.Range(f"B{last + 1}:I{last + 1}").Borders(8).LineStyle = 1
    s.Range(f"B{last + 2}").Value = "Check: every paid line is listed and adds up to net sales"
    s.Range(f"I{last + 2}").Formula = (f'=IF(AND(ROUND(I{last + 1}-D8,2)=0,COUNT(I{STMT_LINES[0]}:I{last})='
                                       f'COUNTIFS(Orders[Sales rep],StmtRep,Orders[Month],StmtMonth,Orders[Status],"OK")),"PASS","FAIL")')
    check_cell(s.Range(f"I{last + 2}"))
    # hide empty statement lines when printing: rows stay, text is blank
    ps = s.PageSetup
    ps.Orientation = 1; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = False
    ps.PrintArea = f"$A$1:$J${last + 2}"
    ps.PrintTitleRows = "$18:$18"
    for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
        setattr(ps, m, 18)


# ------------------------------------------------------------------------------------------------- Test cases
COMM_CASES = [  # case, net sales, target, method, expected commission, expected bonus  (worked by hand, bands 0/15k/30k/45k at 2/4/6/8%)
    ("No sales", 0, 20000, "Whole amount", 0.00, 0),
    ("One cent under the first step", 14999.99, 20000, "Whole amount", 300.00, 0),       # 14,999.99 x 2% = 299.9998
    ("Exactly on the 15,000 step", 15000, 20000, "Whole amount", 600.00, 0),             # 15,000 x 4%
    ("Same 15,000, stepped", 15000, 20000, "Stepped", 300.00, 0),                        # 15,000 x 2%
    ("30,000 stepped", 30000, 35000, "Stepped", 900.00, 0),                              # 15,000 x 2% + 15,000 x 4%
    ("Top band, stepped", 52340.50, 40000, "Stepped", 2387.24, 300),                     # 300 + 600 + 900 + 7,340.50 x 8% (587.24)
    ("Top band, whole amount", 52340.50, 40000, "Whole amount", 4187.24, 300),           # 52,340.50 x 8%
    ("Target met exactly", 25000, 25000, "Whole amount", 1000.00, 300),                  # 25,000 x 4%, bonus
    ("One cent under target", 24999.99, 25000, "Whole amount", 1000.00, 0),              # 999.9996 rounds to 1,000.00, no bonus
]
ORDER_CASES = [  # case, date, rep, product, qty, discount, expected net, expected tax, expected status
    ("Discount exactly at the product max", date(2026, 10, 5), "Grace Holloway", "Gel Pens (pack 12)", 30, 0.15, 249.90, 20.62, "OK"),
    ("Discount above the product max", date(2026, 10, 14), "Diego Santos", "A4 Copy Paper (box)", 40, 0.25, 1260.00, 103.95, "Discount above 10% max"),
    ("Quantity zero", date(2026, 10, 21), "Priya Nair", "Standing Desk", 0, 0, 0.00, 0.00, "Qty must be 1 or more"),
    ("Product not on the price list", date(2026, 10, 27), "Lena Fischer", "Desk Lamp LED", 6, 0, None, None, "Unknown product"),
    ("Sales rep name mistyped", date(2026, 9, 18), "Tom Baker", "Ergonomic Chair", 3, 0, 867.00, 71.53, "Unknown sales rep"),
    ("Several units with a discount", date(2026, 10, 9), "Marcus Reid", "Ergonomic Chair", 7, 0.05, 1921.85, 158.55, "OK"),
    ("Tax of exactly half a cent rounds up", date(2026, 10, 12), "Tom Becker", "Monitor Arm", 2, 0, 238.00, 19.64, "OK"),
    ("Missing order date", None, "Priya Nair", "Laser Printer", 1, 0, 379.00, 31.27, "Missing date"),
]


def build_tests(wb):
    t = wb.Worksheets("Test cases")
    t.Range("A1").Value = "Test cases - worked out by hand, run through the same formulas as the calculator"
    font(t.Range("A1"), 14, True, NAVY)
    t.Range("A2").Value = ("Expected values are typed in; Calculator values are formulas. Every Result must say PASS. "
                           "If the plan on Plan changes (bands, tax, bonus), update the expected values with it.")
    font(t.Range("A2"), 10, False, GREY)

    t.Range("A4").Value = "Commission"
    font(t.Range("A4"), 12, True, NAVY)
    hdr = ["#", "Case", "Net sales", "Target", "Method", "Expected commission", "Expected bonus",
           "Calculator commission", "Calculator bonus"]
    t.Range("A5:I5").Value = tuple(hdr)
    t.Range("J5").Value = "Result"
    for i, (case, net, target, method, ec, eb) in enumerate(COMM_CASES, 6):
        t.Range(f"A{i}:G{i}").Value = (i - 5, case, net, target, method, ec, eb)
        t.Range(f"H{i}").Formula = "=" + comm(f"C{i}", f"E{i}")
        t.Range(f"I{i}").Formula = f"=IF(C{i}>=D{i},TargetBonus,0)"
        t.Range(f"J{i}").Formula = f'=IF(AND(ROUND(H{i}-F{i},2)=0,ROUND(I{i}-G{i},2)=0),"PASS","FAIL")'
        check_cell(t.Range(f"J{i}"))
    end1 = 5 + len(COMM_CASES)
    t.Range(f"C6:D{end1}").NumberFormat = "#,##0.00"; t.Range(f"F6:I{end1}").NumberFormat = "#,##0.00"
    style_header(t.Range("A5:J5"))
    t.Range("A5:J5").WrapText = True; t.Rows(5).RowHeight = 30; t.Range("A5:J5").VerticalAlignment = -4108

    top = end1 + 3
    t.Range(f"A{top - 1}").Value = "Order lines"
    font(t.Range(f"A{top - 1}"), 12, True, NAVY)
    # an OrderTests table with the same column names as Orders, so the very same formulas fill it
    hdr2 = ["#", "Case"] + ["Date", "Sales rep", "Product", "Qty", "Discount %"] + [n for n, _, _ in ORDER_CALC] + \
           ["Expected net", "Expected tax", "Expected status", "Result"]
    t.Range(t.Cells(top, 1), t.Cells(top, len(hdr2))).Value = tuple(hdr2)
    lo = t.ListObjects.Add(1, t.Range(t.Cells(top, 1), t.Cells(top + len(ORDER_CASES), len(hdr2))), None, 1)
    lo.Name = "OrderTests"
    lo.TableStyle = "TableStyleLight9"
    for i, (case, d, rep, prod, qty, disc, en, et, es) in enumerate(ORDER_CASES, 1):
        r = top + i
        t.Range(f"A{r}").Value = i
        t.Range(f"B{r}").Value = case
        if d:
            t.Range(f"C{r}").Value2 = serial(d)
        t.Range(f"D{r}:G{r}").Value = (rep, prod, qty, disc)
        if en is not None:
            lo.ListColumns("Expected net").DataBodyRange.Cells(i, 1).Value = en
            lo.ListColumns("Expected tax").DataBodyRange.Cells(i, 1).Value = et
        lo.ListColumns("Expected status").DataBodyRange.Cells(i, 1).Value = es
    for name, f, fmt in ORDER_CALC:
        lo.ListColumns(name).DataBodyRange.Formula = f
        lo.ListColumns(name).DataBodyRange.NumberFormat = fmt
    lo.ListColumns("Result").DataBodyRange.Formula = (
        '=IF(AND(ROUND(N([@Net])-N([@[Expected net]]),2)=0,ROUND(N([@Tax])-N([@[Expected tax]]),2)=0,'
        '[@Status]=[@[Expected status]]),"PASS","FAIL")')
    check_cell(lo.ListColumns("Result").DataBodyRange)
    lo.ListColumns("Date").DataBodyRange.NumberFormat = "dd mmm yyyy"
    lo.ListColumns("Discount %").DataBodyRange.NumberFormat = "0%"
    for n in ("Expected net", "Expected tax"):
        lo.ListColumns(n).DataBodyRange.NumberFormat = "#,##0.00"
    # Result column for both blocks is column M on the commission block; for the table it is its own column
    res_col = lo.ListColumns("Result").Range.Column
    for col, w in zip(range(1, len(hdr2) + 1), [4, 34, 12, 15, 26, 12, 12, 12, 12, 10, 10, 9, 11, 10, 22, 11, 11, 22, 9]):
        t.Columns(col).ColumnWidth = w
    t.Range(f"A{top + len(ORDER_CASES) + 2}").Formula = (
        f'="Cases passed: "&(COUNTIF($J$6:$J${end1},"PASS")+COUNTIF(OrderTests[Result],"PASS"))&" of "&'
        f'({len(COMM_CASES)}+ROWS(OrderTests[Result]))')
    font(t.Range(f"A{top + len(ORDER_CASES) + 2}"), 11, True, NAVY)
    wb.Names.Add("TestsFailed", f"='Test cases'!$A${top + len(ORDER_CASES) + 3}")
    t.Range(f"A{top + len(ORDER_CASES) + 3}").Formula = f'=COUNTIF($J$6:$J${end1},"FAIL")+COUNTIF(OrderTests[Result],"FAIL")'
    t.Range(f"A{top + len(ORDER_CASES) + 3}").Font.Color = rgb(255, 255, 255)
    return res_col


def style_header(rng):
    font(rng, 10, True, (255, 255, 255))
    rng.Interior.Color = rgb(*NAVY)


def build_howto(wb):
    ws = wb.Worksheets("How to use")
    ws.Columns("A").ColumnWidth = 115
    text = [("How to use the commission calculator", True), ("", False),
            ("Every day (sales team)", True),
            ("1. Orders: type or paste each order line in the yellow columns. Rep and product come from drop-downs.", False),
            ("2. Price, discount, tax and total are calculated and locked. Status must say OK; red = fix that line.", False),
            ("", False), ("Month end (sales manager)", True),
            ("1. Commission: pick the month. Fix every line under 'Lines held back', then the totals are final.", False),
            ("2. Statement: pick each rep, print or save as PDF, send it to the rep.", False),
            ("3. Check that the four data checks at the bottom of Commission say PASS.", False),
            ("", False), ("Changing the plan", True),
            ("Plan is not locked: change prices, max discounts, reps, targets, bands, tax rate, bonus or the method there.", False),
            ("Whole amount = all net sales at the band reached. Stepped = each slice of sales at its own band rate.", False),
            ("After a plan change, update the expected values on Test cases so they still describe the agreed plan.", False),
            ("", False), ("Protection", True),
            ("Commission, Statement, Orders and Test cases are protected without a password (Review > Unprotect Sheet).", False),
            ("Orders has room for 600 lines; for more, unprotect and drag the table's corner down.", False),
            ("Drop-downs and validation stop typing mistakes; pasted data skips validation, which is why Status checks every line.", False),
            ("", False), ("Built and tested in Excel for Windows (2016+ formulas only). No macros, no Power Query.", False)]
    for i, (s, bold) in enumerate(text, 1):
        ws.Cells(i, 1).Value = s
        ws.Cells(i, 1).Font.Bold = bold
    ws.Range("A1").Font.Size = 16


if __name__ == "__main__":
    main()
