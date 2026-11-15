"""Builds Sales-Commission-Calculator.xlsx with real Excel (COM). No macros.

Sheets
  Dashboard   pick a month: KPI cards, 12-month pay trend per rep, share of commission, leaderboard with % of target
              gauges, insights written by formulas, check summary
  Commission  the selected month in detail: pay per rep (commission, accelerator, bonus, manager override, clawback),
              net sales vs target, the plan in words, lines held back, returns clawed back, 8 data checks
  Statement   pick a rep and a month: a printable statement (logo slot, attainment gauge, receipt-style working,
              year to date, every order line credited)
  What-if     change band rates, targets, accelerator and bonus: cost impact against the current plan
  Payouts     payout status per rep and month (Pending / Approved / Paid + paid date) and the pay ledger behind
              every number (12 months x 6 reps)
  Orders      the only place the sales team types order lines (drop-downs, validation, split deals)
  Returns     goods returned / orders cancelled: the commission paid on them is clawed back
  Plan        prices, reps, managers, targets, bands, tax, bonus, override, accelerator, method
  Test cases  hand-worked cases for every rule, run through the same formulas; all must say PASS
  How to use
  Calc        (hidden) chart and dashboard helpers

Run with a Python that has pywin32:  python scripts/build_workbook.py
"""
import csv, json
from datetime import date
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Sales-Commission-Calculator.xlsx"
PLAN = json.loads((ROOT / "data" / "plan.json").read_text())
ORDER_ROWS = 3000   # pre-sized Orders table: a protected sheet cannot grow a table
RETURN_ROWS = 60
STMT_ROWS = 100     # order lines a statement can list
FIRST = date.fromisoformat(PLAN["first_month"])
MONTHS = [date(FIRST.year + (FIRST.month - 1 + i) // 12, (FIRST.month - 1 + i) % 12 + 1, 1) for i in range(PLAN["months"])]
REPS = [r[0] for r in PLAN["reps"]]
serial = lambda d: (d - date(1899, 12, 30)).days


def rgb(c):
    return c[0] + c[1] * 256 + c[2] * 65536


# ---- dark luxury palette (same as sample E9)
BG, PAGE, CARD, CARD2, LINE = (11, 15, 19), (18, 23, 29), (26, 33, 41), (32, 40, 50), (42, 51, 61)
GOLD, GOLD_LT, IVORY, MUTED = (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163)
JADE, COPPER, SKY, GREYB = (79, 163, 140), (181, 117, 74), (127, 167, 201), (74, 85, 96)
GOOD, BAD, AMBER = (91, 191, 138), (229, 115, 95), (217, 164, 65)
INPUT_DK, NAV_ON = (52, 44, 28), (59, 49, 31)
# ---- light working sheets
WHITE, INK, YELLOW, PAPER = (255, 255, 255), (38, 44, 52), (255, 242, 204), (251, 249, 244)
GREEN_BG, GREEN_FG, RED_BG, RED_FG = (198, 239, 206), (0, 97, 0), (255, 199, 206), (156, 0, 6)
REP_COLORS = [GOLD, JADE, COPPER, SKY, (178, 146, 204), MUTED]

SHEETS = ["Dashboard", "Commission", "Statement", "What-if", "Payouts", "Orders", "Returns", "Plan", "Test cases",
          "How to use", "Calc"]
NAV = SHEETS[:9]
DARK = {"Dashboard", "Commission", "What-if", "Payouts"}
GEORGIA, BODY = "Georgia", "Segoe UI"
BRAND = "   ".join(" ".join(w.upper()) for w in PLAN["company"].split())


# ================================================================================================ shared formulas
def comm(net, method, bands="Bands"):
    """Commission: Whole amount = all sales at the rate of the band reached; Stepped = each slice at its own rate
    (net above each band start x the rate increase at that band = [Step]). Used by every sheet and by Test cases."""
    return (f'IF({method}="Stepped",ROUND(SUMPRODUCT(({net}>{bands}[From])*({net}-{bands}[From])*{bands}[Step]),2),'
            f'ROUND({net}*INDEX({bands}[Rate],MATCH({net},{bands}[From],1)),2))')


def accel(net, target, at="AccelAt", rate="AccelRate"):
    """Accelerator: extra rate on the net sales above (at x target). No target = no accelerator."""
    return f"IF(N({target})>0,ROUND(MAX(0,{net}-{at}*{target})*{rate},2),0)"


def bonus(net, target, amount="TargetBonus"):
    return f"IF(AND(N({target})>0,{net}>={target}),{amount},0)"


def override(team_sales):
    return f"ROUND(OverrideRate*{team_sales},2)"


def returned_net(line_net, qty_ret, qty_sold):
    return f"ROUND({line_net}*N({qty_ret})/{qty_sold},2)"


def split_part(amount, pct):
    return f"ROUND({amount}*N({pct}),2)"


def clawback(share, comm_orig, net_orig):
    """Clawback = the returned share x the rate the rep was actually paid in the order's month (commission / net)."""
    return f"ROUND({share}*IFERROR({comm_orig}/{net_orig},0),2)"


def to_pay(status, total):
    return f'IF({status}="Paid",0,{total})'


def payout_check(status, paid_date):
    return (f'IF(AND({status}="Paid",{paid_date}=""),"Paid date missing",IF(AND({status}<>"Paid",{paid_date}<>""),'
            f'"Date but not Paid",IF(COUNTIF(StatusList,{status})=0,"Unknown status","OK")))')


ORDER_CALC = [  # same strings fill the Orders table and the OrderTests table on Test cases
    ("Unit price", '=IF([@Product]="","",IFERROR(INDEX(Products[Unit price],MATCH([@Product],Products[Product],0)),""))', "#,##0.00"),
    ("Gross", '=IF([@[Unit price]]="","",ROUND(N([@Qty])*[@[Unit price]],2))', "#,##0.00"),
    ("Discount", '=IF([@Gross]="","",ROUND([@Gross]*N([@[Discount %]]),2))', "#,##0.00"),
    ("Net", '=IF([@Gross]="","",[@Gross]-[@Discount])', "#,##0.00"),
    ("Tax", '=IF([@Net]="","",ROUND([@Net]*TaxRate,2))', "#,##0.00"),
    ("Total", '=IF([@Net]="","",[@Net]+[@Tax])', "#,##0.00"),
    ("Split credit", '=IF([@Net]="","",IF([@[Split rep]]="",0,' + split_part("[@Net]", "[@[Split %]]") + "))", "#,##0.00"),
    ("Rep credit", '=IF([@Net]="","",[@Net]-[@[Split credit]])', "#,##0.00"),
    ("Month", '=IF([@Date]="","",DATE(YEAR([@Date]),MONTH([@Date]),1))', "mmm yyyy"),
    ("Status", '=IF(AND([@Date]="",[@[Sales rep]]="",[@Product]=""),"",'
               'IF([@Date]="","Missing date",'
               'IF(COUNTIF(Reps[Sales rep],[@[Sales rep]])=0,"Unknown sales rep",'
               'IF(COUNTIF(Products[Product],[@Product])=0,"Unknown product",'
               'IF(N([@Qty])<1,"Qty must be 1 or more",'
               'IF(N([@[Discount %]])>INDEX(Products[Max discount],MATCH([@Product],Products[Product],0))+0.000001,'
               '"Discount above "&TEXT(INDEX(Products[Max discount],MATCH([@Product],Products[Product],0)),"0%")&" max",'
               'IF(AND([@[Split rep]]<>"",COUNTIF(Reps[Sales rep],[@[Split rep]])=0),"Unknown split rep",'
               'IF(AND([@[Split rep]]<>"",[@[Split rep]]=[@[Sales rep]]),"Split rep same as sales rep",'
               'IF(AND([@[Split rep]]<>"",OR(N([@[Split %]])<0.01,N([@[Split %]])>0.99)),"Split % must be 1% to 99%",'
               'IF(AND([@[Split rep]]="",N([@[Split %]])<>0),"Split % without a split rep","OK"))))))))))', "General"),
]
ORDER_HELPERS = [  # Orders only, hidden: lookups and lists without FILTER (works in Excel 2016)
    ("Key", '=[@[Order no]]&"|"&[@Product]'),
    ("Held", '=IF(AND([@Status]<>"",[@Status]<>"OK",OR([@Month]=SelMonth,[@Month]="")),1,0)'),
    ("Held line", '=IF([@Held]=1,SUM(INDEX([Held],1):[@Held]),"")'),
    ("In statement", '=IF(AND([@Status]="OK",[@Month]=StmtMonth,OR([@[Sales rep]]=StmtRep,[@[Split rep]]=StmtRep)),1,0)'),
    ("Statement line", '=IF([@[In statement]]=1,SUM(INDEX([In statement],1):[@[In statement]]),"")'),
    ("Stmt share", '=IF([@[In statement]]=1,IF([@[Sales rep]]=StmtRep,1-IF([@[Split rep]]="",0,N([@[Split %]])),N([@[Split %]])),"")'),
    ("Stmt credit", '=IF([@[In statement]]=1,IF([@[Sales rep]]=StmtRep,[@[Rep credit]],[@[Split credit]]),"")'),
]
INPUTS = ["Date", "Order no", "Sales rep", "Customer", "Product", "Qty", "Discount %", "Split rep", "Split %"]

RET_INPUTS = ["Return date", "Order no", "Product", "Qty returned", "Reason"]
LEDGER_KEY = lambda month, rep: f'YEAR({month})*100+MONTH({month})&"|"&{rep}'
RET_CALC = [  # same strings fill the Returns table and the ReturnTests table on Test cases
    ("Line", '=IF([@[Order no]]="","",IFERROR(MATCH([@[Order no]]&"|"&[@Product],Orders[Key],0),""))', "0"),
    ("Sales rep", '=IF(N([@Line])=0,"",INDEX(Orders[Sales rep],[@Line]))', "General"),
    ("Split rep", '=IF(N([@Line])=0,"",INDEX(Orders[Split rep],[@Line])&"")', "General"),
    ("Split %", '=IF(N([@Line])=0,"",N(INDEX(Orders[Split %],[@Line])))', "0%"),
    ("Order date", '=IF(N([@Line])=0,"",INDEX(Orders[Date],[@Line]))', "dd mmm yyyy"),
    ("Qty sold", '=IF(N([@Line])=0,"",INDEX(Orders[Qty],[@Line]))', "0"),
    ("Returned net", '=IF(N([@Line])=0,"",IFERROR(' + returned_net("INDEX(Orders[Net],[@Line])", "[@[Qty returned]]", "[@[Qty sold]]") + ',""))', "#,##0.00"),
    ("Booked month", '=IF([@[Return date]]="","",DATE(YEAR([@[Return date]]),MONTH([@[Return date]]),1))', "mmm yyyy"),
    ("Status", '=IF(AND([@[Return date]]="",[@[Order no]]=""),"",IF([@[Return date]]="","Missing date",'
               'IF(N([@Line])=0,"Order line not found",IF(N([@[Qty returned]])<1,"Qty must be 1 or more",'
               'IF([@[Qty returned]]>[@[Qty sold]],"Qty above qty sold",IF([@[Return date]]<[@[Order date]],'
               '"Return before order date",IF(INDEX(Orders[Status],[@Line])<>"OK","Line was held back (never paid)","OK")))))))', "General"),
    ("Split share", '=IF([@Status]<>"OK",0,IF([@[Split rep]]="",0,' + split_part("[@[Returned net]]", "[@[Split %]]") + "))", "#,##0.00"),
    ("Rep share", '=IF([@Status]<>"OK",0,[@[Returned net]]-[@[Split share]])', "#,##0.00"),
    ("Rep commission", '=IF([@Status]<>"OK",0,IFERROR(INDEX(Ledger[Commission],MATCH(' + LEDGER_KEY("[@[Order date]]", "[@[Sales rep]]") + ',Ledger[Key],0)),0))', "#,##0.00"),
    ("Rep net", '=IF([@Status]<>"OK",0,IFERROR(INDEX(Ledger[Credited net],MATCH(' + LEDGER_KEY("[@[Order date]]", "[@[Sales rep]]") + ',Ledger[Key],0)),0))', "#,##0.00"),
    ("Split commission", '=IF(OR([@Status]<>"OK",[@[Split rep]]=""),0,IFERROR(INDEX(Ledger[Commission],MATCH(' + LEDGER_KEY("[@[Order date]]", "[@[Split rep]]") + ',Ledger[Key],0)),0))', "#,##0.00"),
    ("Split net", '=IF(OR([@Status]<>"OK",[@[Split rep]]=""),0,IFERROR(INDEX(Ledger[Credited net],MATCH(' + LEDGER_KEY("[@[Order date]]", "[@[Split rep]]") + ',Ledger[Key],0)),0))', "#,##0.00"),
    ("Clawback rep", "=" + clawback("[@[Rep share]]", "[@[Rep commission]]", "[@[Rep net]]"), "#,##0.00"),
    ("Clawback split", "=" + clawback("[@[Split share]]", "[@[Split commission]]", "[@[Split net]]"), "#,##0.00"),
    ("Clawback", '=IF([@Status]="","",[@[Clawback rep]]+[@[Clawback split]])', "#,##0.00"),
]
RET_HELPERS = [
    ("In month", '=IF(AND([@Status]<>"",[@[Booked month]]=SelMonth),1,0)'),
    ("Month line", '=IF([@[In month]]=1,SUM(INDEX([In month],1):[@[In month]]),"")'),
]
LEDGER_COLS = [  # (name, formula or None for typed values, number format)
    ("Sales rep", None, "@"), ("Month", None, "mmm yyyy"),
    ("Manager", '=IFERROR(INDEX(Reps[Manager],MATCH([@[Sales rep]],Reps[Sales rep],0))&"","")', "General"),
    ("Target", "=IFERROR(INDEX(Reps[Monthly target],MATCH([@[Sales rep]],Reps[Sales rep],0)),0)", "#,##0"),
    ("Lines", '=COUNTIFS(Orders[Sales rep],[@[Sales rep]],Orders[Month],[@Month],Orders[Status],"OK")'
              '+COUNTIFS(Orders[Split rep],[@[Sales rep]],Orders[Month],[@Month],Orders[Status],"OK")', "0"),
    ("Credited net", '=ROUND(SUMIFS(Orders[Rep credit],Orders[Sales rep],[@[Sales rep]],Orders[Month],[@Month],Orders[Status],"OK")'
                     '+SUMIFS(Orders[Split credit],Orders[Split rep],[@[Sales rep]],Orders[Month],[@Month],Orders[Status],"OK"),2)', "#,##0.00"),
    ("Achieved", "=IF([@Target]>0,[@[Credited net]]/[@Target],0)", "0.0%"),
    ("Band rate", "=INDEX(Bands[Rate],MATCH([@[Credited net]],Bands[From],1))", "0.0%"),
    ("Commission", "=" + comm("[@[Credited net]]", "Method"), "#,##0.00"),
    ("Accelerator", "=" + accel("[@[Credited net]]", "[@Target]"), "#,##0.00"),
    ("Bonus", "=" + bonus("[@[Credited net]]", "[@Target]"), "#,##0.00"),
    ("Override", "=" + override("SUMIFS([Credited net],[Manager],[@[Sales rep]],[Month],[@Month])"), "#,##0.00"),
    ("Clawback", '=SUMIFS(Returns[Clawback rep],Returns[Sales rep],[@[Sales rep]],Returns[Booked month],[@Month],Returns[Status],"OK")'
                 '+SUMIFS(Returns[Clawback split],Returns[Split rep],[@[Sales rep]],Returns[Booked month],[@Month],Returns[Status],"OK")', "#,##0.00"),
    ("Total pay", "=[@Commission]+[@Accelerator]+[@Bonus]+[@Override]-[@Clawback]", "#,##0.00"),
    ("Status", None, "@"), ("Paid date", None, "dd mmm yyyy"),
    ("To pay", "=" + to_pay("[@Status]", "[@[Total pay]]"), "#,##0.00"),
    ("Check", "=" + payout_check("[@Status]", "[@[Paid date]]"), "General"),
    ("What-if target", "=ROUND([@Target]*(1+ScTargetChange),2)", "#,##0.00"),
    ("What-if commission", "=" + comm("[@[Credited net]]", "Method", "ScBands"), "#,##0.00"),
    ("What-if accelerator", "=" + accel("[@[Credited net]]", "[@[What-if target]]", "ScAccelAt", "ScAccelRate"), "#,##0.00"),
    ("What-if bonus", "=" + bonus("[@[Credited net]]", "[@[What-if target]]", "ScBonus"), "#,##0.00"),
    ("What-if pay", "=[@[What-if commission]]+[@[What-if accelerator]]+[@[What-if bonus]]+[@Override]-[@Clawback]", "#,##0.00"),
    ("Key", "=" + LEDGER_KEY("[@Month]", "[@[Sales rep]]"), "General"),
]


# ================================================================================================ styling helpers
def font(rng, size=None, bold=None, color=None, name=None, italic=None):
    if name: rng.Font.Name = name
    if size: rng.Font.Size = size
    if bold is not None: rng.Font.Bold = bold
    if italic is not None: rng.Font.Italic = italic
    if color: rng.Font.Color = rgb(color)


def put(ws, addr, val, fmt=None):
    c = ws.Range(addr)
    if isinstance(val, str) and val.startswith("="):
        c.Formula = val          # formula first: a text (@) format set before it would store the formula as text
        if fmt:
            c.NumberFormat = fmt
    else:
        if fmt:
            c.NumberFormat = fmt
        c.Value = val
    return c


def merge(ws, addr, val=None, fmt=None, h=None, v=-4108):
    rng = ws.Range(addr)
    rng.Merge()
    if val is not None:
        put(ws, addr.split(":")[0], val, fmt)
    if h: rng.HorizontalAlignment = h
    rng.VerticalAlignment = v
    return rng


def fill(rng, color):
    rng.Interior.Color = rgb(color)


def underline(rng, color, weight=2):
    b = rng.Borders(9); b.LineStyle = 1; b.Color = rgb(color); b.Weight = weight


def cf_font(rng, formula, color, bold=None):
    fc = rng.FormatConditions.Add(2, None, formula)
    fc.Font.Color = rgb(color)
    if bold is not None: fc.Font.Bold = bold
    return fc


def cf_fill(rng, formula, bg, fg=None):
    fc = rng.FormatConditions.Add(2, None, formula)
    fc.Interior.Color = rgb(bg)
    if fg: fc.Font.Color = rgb(fg)
    return fc


def pass_fail(cell, dark=False):
    font(cell, 10, True); cell.HorizontalAlignment = -4108
    for val, bg, fg in (("PASS", (24, 58, 44) if dark else GREEN_BG, GOOD if dark else GREEN_FG),
                        ("FAIL", (70, 30, 26) if dark else RED_BG, BAD if dark else RED_FG)):
        fc = cell.FormatConditions.Add(1, 3, f'="{val}"'); fc.Interior.Color = rgb(bg); fc.Font.Color = rgb(fg)


def widths(ws, spec):
    for col, w in spec.items():
        ws.Columns(col).ColumnWidth = w


def header(ws, last_col, title, sub, dark=True, tc="B"):
    """Charcoal band: spaced gold brand line, serif title, subtitle, thin gold rule, navigation strip (row 5)."""
    if dark:
        ws.Cells.Interior.Color = rgb(PAGE)
        ws.Cells.Font.Color = rgb(IVORY)
    ws.Cells.Font.Name = BODY
    ws.Cells.Font.Size = 10
    fill(ws.Range(f"A1:{last_col}5"), BG)
    fill(ws.Range(f"A4:{last_col}4"), GOLD)
    for r, h in ((1, 20), (2, 30), (3, 18), (4, 2.25), (5, 27), (6, 10)):
        ws.Rows(r).RowHeight = h
    put(ws, f"{tc}1", BRAND); font(ws.Range(f"{tc}1"), 8, True, GOLD)
    put(ws, f"{tc}2", title); font(ws.Range(f"{tc}2"), 20, False, IVORY, GEORGIA)
    put(ws, f"{tc}3", sub); font(ws.Range(f"{tc}3"), 9, False, MUTED)
    for c in (f"{tc}1", f"{tc}2", f"{tc}3"):
        ws.Range(c).VerticalAlignment = -4108


def nav(ws, current, row=5, col="B", printable=True):
    x = ws.Range(f"{col}{row}").Left
    top = ws.Range(f"{col}{row}").Top + 4
    for name in NAV:
        w = 18 + 5.4 * len(name)
        shp = ws.Shapes.AddShape(5, x, top, w, 19)
        shp.Name = "nav_" + name.replace(" ", "_").replace("-", "_")
        on = name == current
        shp.Fill.ForeColor.RGB = rgb(NAV_ON if on else CARD)
        shp.Line.ForeColor.RGB = rgb(GOLD if on else LINE)
        shp.Line.Weight = 1 if on else 0.75
        tf = shp.TextFrame2
        tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
        tf.VerticalAnchor = 3
        tr = tf.TextRange
        tr.Text = name
        tr.Font.Size = 8.5
        tr.Font.Name = BODY
        tr.Font.Bold = on
        tr.Font.Fill.ForeColor.RGB = rgb(GOLD_LT if on else IVORY)
        tr.ParagraphFormat.Alignment = 2
        shp.Placement = 3
        if not printable:
            shp.DrawingObject.PrintObject = False
        if not on:
            ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress=f"'{name}'!A1", ScreenTip=f"Go to {name}")
        x += w + 5


def icon(ws, cell, glyph, size=17):
    c = ws.Range(cell)
    shp = ws.Shapes.AddShape(9, c.Left + c.Width - size - 6, c.Top + 3, size, size)
    shp.Fill.Visible = False
    shp.Line.ForeColor.RGB = rgb(GOLD); shp.Line.Weight = 1
    tf = shp.TextFrame2
    tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
    tf.VerticalAnchor = 3; tf.WordWrap = False
    tr = tf.TextRange; tr.Text = glyph
    tr.Font.Size = 8.5; tr.Font.Bold = True; tr.Font.Name = "Segoe UI Symbol"
    tr.Font.Fill.ForeColor.RGB = rgb(GOLD); tr.ParagraphFormat.Alignment = 2
    shp.Placement = 3
    return shp


def panel(ws, addr, title=None, title_row=None):
    rng = ws.Range(addr)
    fill(rng, CARD)
    for e in (7, 8, 9, 10):
        rng.Borders(e).LineStyle = 1; rng.Borders(e).Color = rgb(LINE)
    if title:
        c = ws.Range(addr.split(":")[0]).Offset(2 if title_row is None else title_row, 2)
        c.Value = title
        font(c, 12.5, False, IVORY, GEORGIA)
        c.VerticalAlignment = -4108


def kpi(ws, first_col, last_col, r0, label, value, fmt, note, glyph, sign=None, note_color=MUTED):
    """KPI card r0..r0+4: label, icon, big serif number, note (green/red from the sign formula)."""
    fc, lc = first_col, last_col
    for r, h in ((r0, 5), (r0 + 1, 18), (r0 + 2, 34), (r0 + 3, 18), (r0 + 4, 6)):
        ws.Rows(r).RowHeight = h
    box = ws.Range(f"{fc}{r0}:{lc}{r0 + 4}")
    fill(box, CARD)
    for e in (8, 9, 10):
        box.Borders(e).LineStyle = 1; box.Borders(e).Color = rgb(LINE)
    b = box.Borders(7); b.LineStyle = 1; b.Color = rgb(GOLD); b.Weight = 4
    lab = ws.Range(f"{fc}{r0 + 1}")
    lab.Value = "  " + label
    font(lab, 7.5, True, MUTED)
    lab.VerticalAlignment = -4108
    merge(ws, f"{fc}{r0 + 2}:{lc}{r0 + 2}", value, fmt, -4108)
    font(ws.Range(f"{fc}{r0 + 2}"), 20, False, IVORY, GEORGIA)
    merge(ws, f"{fc}{r0 + 3}:{lc}{r0 + 3}", note, None, -4108)
    font(ws.Range(f"{fc}{r0 + 3}"), 8.5, True, note_color)
    if sign:
        cf_font(ws.Range(f"{fc}{r0 + 3}"), f"={sign}>0", GOOD)
        cf_font(ws.Range(f"{fc}{r0 + 3}"), f"={sign}<0", BAD)
    icon(ws, f"{lc}{r0 + 1}", glyph)


def dark_chart(ch, legend=True):
    ch.ChartArea.Format.Fill.Visible = True
    ch.ChartArea.Format.Fill.ForeColor.RGB = rgb(CARD)
    ch.ChartArea.Format.Line.Visible = False
    ch.PlotArea.Format.Fill.Visible = False
    ch.ChartArea.Font.Color = rgb(MUTED)
    ch.ChartArea.Font.Size = 9
    ch.ChartArea.Font.Name = BODY
    ch.HasTitle = False
    ch.HasLegend = legend
    if legend:
        ch.Legend.Font.Color = rgb(IVORY)


def dark_axes(ch, value_fmt=None):
    for ax in (1, 2):
        try:
            a = ch.Axes(ax)
        except Exception:
            continue
        a.Format.Line.Visible = ax == 1
        if ax == 1:
            a.Format.Line.ForeColor.RGB = rgb(LINE)
        a.TickLabels.Font.Color = rgb(MUTED)
        a.MajorTickMark = -4142
    ch.Axes(2).HasMajorGridlines = True
    ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = rgb(LINE)
    if value_fmt:
        ch.Axes(2).TickLabels.NumberFormat = value_fmt


def add_chart(ws, addr, ctype):
    rng = ws.Range(addr)
    co = ws.ChartObjects().Add(rng.Left + 4, rng.Top + 2, rng.Width - 8, rng.Height - 4)
    co.Chart.ChartType = ctype
    while co.Chart.SeriesCollection().Count:
        co.Chart.SeriesCollection(1).Delete()
    co.Placement = 2
    return co


def table_header(ws, row, cols_titles, dark=True, left=()):
    for c, t in cols_titles:
        h = ws.Range(f"{c}{row}")
        h.Value = t
        font(h, 8.5, True, GOLD_LT if dark else WHITE)
        fill(h, BG if dark else (40, 48, 58))
        h.HorizontalAlignment = -4131 if t in left else -4152
        h.VerticalAlignment = -4108


def selector(ws, addr, name, wb, source, fmt=None, value=None, dark=True):
    sel = ws.Range(addr)
    sel.Merge()
    fill(sel, INPUT_DK if dark else (250, 243, 226))
    for e in (7, 8, 9, 10):
        sel.Borders(e).LineStyle = 1; sel.Borders(e).Color = rgb(GOLD)
    font(sel, 11, True, GOLD_LT if dark else INK)
    sel.HorizontalAlignment = -4108; sel.VerticalAlignment = -4108
    sel.Locked = False
    first = sel.Cells(1, 1)
    first.Validation.Delete()
    first.Validation.Add(3, 1, 1, source)
    first.Validation.InputTitle = "Pick from the list"
    first.Validation.InputMessage = "Every figure follows this choice."
    if fmt:
        first.NumberFormat = fmt
    if value is not None:
        first.Value2 = value
    wb.Names.Add(name, f"='{ws.Name}'!{first.Address}")


def dark_table_style(wb):
    ts = wb.TableStyles.Add("HarborDark")
    whole = ts.TableStyleElements(0)    # xlWholeTable
    whole.Interior.Color = rgb(CARD); whole.Font.Color = rgb(IVORY)
    for e in (9, 12):                   # inside horizontal, bottom
        whole.Borders(e).LineStyle = 1; whole.Borders(e).Color = rgb(LINE)
    hdr = ts.TableStyleElements(1)      # xlHeaderRow
    hdr.Interior.Color = rgb(BG); hdr.Font.Color = rgb(GOLD_LT); hdr.Font.Bold = True
    stripe = ts.TableStyleElements(5)   # xlRowStripe1
    stripe.Interior.Color = rgb(CARD2)
    return ts


def make_table(ws, top_left, headers, rows, name, fmts=None, style="TableStyleLight9"):
    r0 = ws.Range(top_left).Row; c0 = ws.Range(top_left).Column
    ws.Range(ws.Cells(r0, c0), ws.Cells(r0, c0 + len(headers) - 1)).Value = tuple(headers)
    if rows:
        ws.Range(ws.Cells(r0 + 1, c0), ws.Cells(r0 + len(rows), c0 + len(headers) - 1)).Value = [tuple(r) for r in rows]
    n = max(len(rows), 1)
    lo = ws.ListObjects.Add(1, ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + n, c0 + len(headers) - 1)), None, 1)
    lo.Name = name
    lo.TableStyle = style
    if style == "HarborDark":
        undark(lo)
    for col, fmt in (fmts or {}).items():
        lo.ListColumns(col).DataBodyRange.NumberFormat = fmt
    return lo


def undark(lo):
    """Direct cell fill/font beats a table style: clear them so HarborDark shows its stripes."""
    lo.Range.Interior.Pattern = -4142
    lo.Range.Font.Color = rgb(IVORY)
    lo.HeaderRowRange.Font.Color = rgb(GOLD_LT)
    lo.HeaderRowRange.Font.Bold = True


def add_columns(lo, specs):
    for name, *_ in specs:
        lo.ListColumns.Add().Name = name


def fill_formulas(lo, specs):
    for spec in specs:
        name, f = spec[0], spec[1]
        if f is None:
            continue
        lc = lo.ListColumns(name)
        lc.DataBodyRange.Formula = f
        if len(spec) > 2:
            lc.DataBodyRange.NumberFormat = spec[2]


# ================================================================================================ main
import time as _t
_T0 = _t.time()


def step(msg):
    print(f"{_t.time() - _T0:6.1f}s  {msg}", flush=True)


def main():
    if OUT.exists():
        OUT.unlink()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < len(SHEETS):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, n in enumerate(SHEETS, 1):
            wb.Worksheets(i).Name = n
            wb.Worksheets(i).Cells.Font.Name = BODY
            wb.Worksheets(i).Cells.Font.Size = 10
        xl.Calculation = -4135  # manual while building
        dark_table_style(wb)
        step('build_plan'); build_plan(wb)
        step('build_whatif_inputs'); build_whatif_inputs(wb)
        # the three selectors first, so helper formulas can refer to them
        for sheet, cell, name, val in (("Dashboard", "F7", "SelMonth", serial(MONTHS[-1])), ("Statement", "C14", "StmtRep", PLAN["reps"][4][0]),
                                       ("Statement", "G14", "StmtMonth", serial(MONTHS[-1]))):
            wb.Worksheets(sheet).Range(cell).Value2 = val
            wb.Names.Add(name, f"='{sheet}'!${cell[0]}${cell[1:]}")
        step('build_orders'); build_orders(wb)
        step('build_returns_and_ledger'); build_returns_and_ledger(wb)
        step('build_tests'); build_tests(wb)
        step('build_commission'); build_commission(wb)
        step('build_statement'); build_statement(wb)
        step('build_whatif'); build_whatif(wb)
        step('build_payouts'); build_payouts(wb)
        step('build_calc'); build_calc(wb)
        step('build_dashboard'); build_dashboard(wb)
        step('build_howto'); build_howto(wb)
        step('calculate'); xl.Calculation = -4105
        xl.CalculateFull()
        step('protect/save')
        for sheet, name in (("Dashboard", "donut"), ("Statement", "gauge")):
            ch = wb.Worksheets(sheet).ChartObjects(name).Chart
            ch.HasTitle = True
            ch.ChartTitle.Delete()      # like deleting it by hand: marks the automatic one-series title as deleted for good
        for name in ("Dashboard", "Commission", "Statement", "What-if", "Payouts", "Orders", "Returns", "Test cases", "Calc"):
            wb.Worksheets(name).Protect(Password="", DrawingObjects=True, Contents=True, Scenarios=True, AllowFiltering=True)
        for ws in wb.Worksheets:
            if ws.Name == "Calc":
                continue
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            if ws.Name in DARK or ws.Name == "Statement":
                xl.ActiveWindow.DisplayHeadings = False
            xl.ActiveWindow.Zoom = 100 if ws.Name in DARK else 90
            ws.Range("A1").Select()
            xl.ActiveWindow.ScrollRow = 1
        wb.Worksheets("Calc").Visible = 0
        wb.Worksheets("Dashboard").Activate()
        wb.SaveAs(str(OUT), 51)
        print("saved", OUT.name)
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        xl.Quit()


# ================================================================================================ Plan
def build_plan(wb):
    p = wb.Worksheets("Plan")
    header(p, "R", "Commission plan & lists",
           "The sales manager edits this sheet (not locked): prices, reps, managers, targets, bands and the rules. "
           "Every other sheet follows.", dark=False, tc="A")
    nav(p, "Plan", col="A")
    put(p, "A7", "Settings"); font(p.Range("A7"), 12, True, INK, GEORGIA)
    settings = [("Company", PLAN["company"], "Company", "@", "Shown on every sheet and statement"),
                ("Sales tax rate", PLAN["tax_rate"], "TaxRate", "0.00%", "Added to the order total; commission is on net sales"),
                ("Target bonus", PLAN["target_bonus"], "TargetBonus", "#,##0", "Paid when a rep's credited net sales reach the monthly target"),
                ("Discount cap (any product)", PLAN["discount_cap"], "DiscountCap", "0%", "No discount above this in Orders; each product also has its own max"),
                ("Commission method", "Whole amount", "Method", "@", "Whole amount or Stepped (see Methods on the right)"),
                ("Manager override", PLAN["override_rate"], "OverrideRate", "0.0%", "Paid to a manager on the credited net sales of the reps they manage"),
                ("Accelerator from", PLAN["accel_at"], "AccelAt", "0%", "of the monthly target ..."),
                ("Accelerator extra rate", PLAN["accel_rate"], "AccelRate", "0.0%", "... paid on top, on the net sales above that point")]
    p.Range("A8:C8").Value = ("Setting", "Value", "What it does")
    font(p.Range("A8:C8"), 9, True, MUTED)
    for i, (label, val, nm, fmt, note) in enumerate(settings, 9):
        p.Cells(i, 1).Value = label
        c = p.Cells(i, 2)
        c.NumberFormat = fmt
        c.Value = val
        fill(c, YELLOW)
        p.Cells(i, 3).Value = note
        wb.Names.Add(nm, f"=Plan!$B${i}")
    font(p.Range("C9:C16"), 9, False, (110, 110, 110))
    put(p, "P7", "Methods"); font(p.Range("P7"), 12, True, INK, GEORGIA)
    p.Range("P8:Q9").Value = (("Whole amount", "All credited net sales are paid at the rate of the band reached"),
                              ("Stepped", "Each band's slice of credited net sales is paid at that band's own rate"))
    wb.Names.Add("MethodList", "=Plan!$P$8:$P$9")
    p.Range("B13").Validation.Add(3, 1, 1, "=MethodList")
    put(p, "P11", "Payout statuses"); font(p.Range("P11"), 12, True, INK, GEORGIA)
    p.Range("P12:Q14").Value = (("Pending", "Calculated, waiting for the manager's approval"),
                                ("Approved", "Approved, waiting to be paid (counts in To pay)"),
                                ("Paid", "Paid; type the paid date next to it"))
    wb.Names.Add("StatusList", "=Plan!$P$12:$P$14")

    products = make_table(p, "A19", ["Product", "Category", "Unit price", "Max discount"], PLAN["products"], "Products",
                          {"Unit price": "#,##0.00", "Max discount": "0%"}, "TableStyleMedium2")
    reps = make_table(p, "F19", ["Sales rep", "Region", "Monthly target", "Manager"], PLAN["reps"], "Reps",
                      {"Monthly target": "#,##0"}, "TableStyleMedium2")
    bands = make_table(p, "K19", ["From", "Rate"], PLAN["bands"], "Bands", {"From": "#,##0", "Rate": "0.0%"}, "TableStyleMedium2")
    band_helpers(bands, "Bands")
    for c, t in (("A18", "Products"), ("F18", "Sales reps (Manager = who earns the override on this rep)"),
                 ("K18", "Commission bands (monthly credited net sales)")):
        put(p, c, t); font(p.Range(c), 12, True, INK, GEORGIA)
    put(p, "K25", "Step = rate increase at that band (Stepped method). Order ok = bands go up from 0.")
    font(p.Range("K25"), 8, False, (110, 110, 110))
    put(p, "P18", "Months"); font(p.Range("P18"), 12, True, INK, GEORGIA)
    for i, m in enumerate(MONTHS):
        c = p.Range(f"P{19 + i}")
        c.Value2 = serial(m); c.NumberFormat = "mmm yyyy"; c.HorizontalAlignment = -4131
    wb.Names.Add("MonthList", f"=Plan!$P$19:$P${18 + len(MONTHS)}")
    # last order date the drop-downs accept (a cell, because validation formulas cannot hold commas in every locale)
    put(p, f"P{20 + len(MONTHS)}", "Last day"); font(p.Range(f"P{20 + len(MONTHS)}"), 8, True, (110, 110, 110))
    put(p, f"P{21 + len(MONTHS)}", "=DATE(YEAR(MAX(MonthList)),MONTH(MAX(MonthList))+1,0)", "dd mmm yyyy")
    wb.Names.Add("LastDay", f"=Plan!$P${21 + len(MONTHS)}")
    wb.Names.Add("RepList", "=Plan!$F$20:$F$" + str(19 + len(PLAN["reps"])))
    wb.Names.Add("ProductList", "=Plan!$A$20:$A$" + str(19 + len(PLAN["products"])))
    widths(p, {"A": 27, "B": 14, "C": 12, "D": 12, "E": 2, "F": 16, "G": 9, "H": 14, "I": 16, "J": 2, "K": 10, "L": 8,
               "M": 8, "N": 9, "O": 2, "P": 15, "Q": 52, "R": 2})
    p.Range("C9:C16").WrapText = False
    font(p.Range("P8:P14"), 10, True, INK)
    font(p.Range("Q8:Q14"), 9, False, (110, 110, 110))


def band_helpers(lo, name):
    hdr = f"ROW({name}[[#Headers],[Rate]])"
    col = lo.ListColumns.Add(); col.Name = "Step"
    col.DataBodyRange.Formula2 = f"=IF(ROW()-{hdr}=1,[@Rate],[@Rate]-INDEX([Rate],ROW()-{hdr}-1))"
    col.DataBodyRange.NumberFormat = "+0.0%;-0.0%;0.0%"
    col = lo.ListColumns.Add(); col.Name = "Order ok"
    col.DataBodyRange.Formula2 = f"=IF(ROW()-{hdr}=1,--([@From]=0),--(INDEX([From],ROW()-{hdr}-1)<[@From]))"


# ================================================================================================ Orders
def build_orders(wb):
    o = wb.Worksheets("Orders")
    header(o, "Z", "Orders",
           "Type or paste order lines in the yellow columns; everything else is calculated and locked. Status must say OK to "
           "be paid. Split rep + Split % = a deal shared with a colleague.", dark=False, tc="A")
    nav(o, "Orders", col="A")
    with open(ROOT / "data" / "orders.csv", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    hdr = INPUTS + [n for n, _, _ in ORDER_CALC] + [n for n, _ in ORDER_HELPERS]
    o.Range(o.Cells(7, 1), o.Cells(7, len(hdr))).Value = tuple(hdr)
    lo = o.ListObjects.Add(1, o.Range(o.Cells(7, 1), o.Cells(7 + ORDER_ROWS, len(hdr))), None, 1)
    lo.Name = "Orders"
    lo.TableStyle = "TableStyleLight9"
    data = [(serial(date.fromisoformat(r["Date"])), r["Order no"], r["Sales rep"], r["Customer"], r["Product"],
             int(r["Qty"]), float(r["Discount %"]), r["Split rep"] or None, float(r["Split %"]) if r["Split %"] else None)
            for r in rows]
    o.Range(o.Cells(8, 1), o.Cells(7 + len(data), len(INPUTS))).Value2 = data
    for name, f, fmt in ORDER_CALC:
        lc = lo.ListColumns(name)
        lc.DataBodyRange.Formula = f
        lc.DataBodyRange.NumberFormat = fmt
        fill(lc.DataBodyRange, (247, 249, 252))
    for name, f in ORDER_HELPERS:
        lo.ListColumns(name).DataBodyRange.Formula = f
    style_order_inputs(o, lo)
    st = lo.ListColumns("Status").DataBodyRange
    sc = col_letter(lo, "Status")
    cf_fill(st, f'=(${sc}8<>"")*(${sc}8<>"OK")', RED_BG, RED_FG).Font.Bold = True
    w = {"Date": 11, "Order no": 9, "Sales rep": 15, "Customer": 23, "Product": 24, "Qty": 5, "Discount %": 9,
         "Split rep": 14, "Split %": 7, "Unit price": 9, "Gross": 10, "Discount": 9, "Net": 10, "Tax": 8, "Total": 10,
         "Split credit": 10, "Rep credit": 10, "Month": 9, "Status": 24}
    for i, h in enumerate(hdr, 1):
        o.Columns(i).ColumnWidth = w.get(h, 8)
    first_helper = len(INPUTS) + len(ORDER_CALC) + 1
    o.Range(o.Cells(1, first_helper), o.Cells(1, len(hdr))).EntireColumn.Hidden = True
    o.Activate(); o.Range("A8").Select(); wb.Application.ActiveWindow.FreezePanes = True


def col_letter(lo, name):
    return lo.ListColumns(name).DataBodyRange.Cells(1, 1).Address.split("$")[1]


def style_order_inputs(ws, lo):
    """Yellow, unlocked input cells with drop-downs and validation (pasted values bypass validation: Status catches them)."""
    body = lambda n: lo.ListColumns(n).DataBodyRange
    for n in INPUTS:
        b = body(n); b.Locked = False; fill(b, YELLOW)
    rules = {
        "Date": (4, 1, "=MIN(MonthList)", "=LastDay", "dd mmm yyyy", "Order date", "A date in the plan's months."),
        "Sales rep": (3, 1, "=RepList", None, None, "Sales rep", "Pick from the list (reps live on Plan)."),
        "Product": (3, 1, "=ProductList", None, None, "Product", "Pick from the list; price comes from Plan."),
        "Qty": (1, 7, "1", None, "0", "Quantity", "Whole number, 1 or more."),
        "Discount %": (2, 1, "0", "=DiscountCap", "0%", "Discount", "0% to the cap; each product has its own max (Plan)."),
        "Split rep": (3, 1, "=RepList", None, None, "Split deal", "Only for a deal shared with a colleague: pick the colleague."),
        "Split %": (2, 1, "=1/100", "=99/100", "0%", "Split %", "The colleague's share of this line, 1% to 99%."),
    }
    for n, (typ, op, f1, f2, fmt, title, msg) in rules.items():
        print("  validation", n)
        v = body(n).Validation
        v.Delete()
        if f2:
            v.Add(typ, 1, op, f1, f2)
        else:
            v.Add(typ, 1, op, f1)
        v.InputTitle = title; v.InputMessage = msg; v.ErrorTitle = title; v.ErrorMessage = msg
        if fmt:
            body(n).NumberFormat = fmt


# ================================================================================================ Returns + ledger tables
def build_returns_and_ledger(wb):
    r = wb.Worksheets("Returns")
    header(r, "AB", "Returns & cancellations (clawback)",
           "Type returned goods or cancelled lines in the yellow columns. The commission paid on them is taken back in the "
           "month of the return, at the rate the rep was paid in the order's month.", dark=False, tc="A")
    nav(r, "Returns", col="A")
    with open(ROOT / "data" / "returns.csv", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    data = [(serial(date.fromisoformat(x["Return date"])), x["Order no"], x["Product"], int(x["Qty returned"]), x["Reason"]) for x in rows]
    rlo = make_table(r, "A7", RET_INPUTS, [], "Returns")
    rlo.Resize(r.Range(r.Cells(7, 1), r.Cells(7 + RETURN_ROWS, len(RET_INPUTS))))
    r.Range(r.Cells(8, 1), r.Cells(7 + len(data), 5)).Value2 = data
    add_columns(rlo, RET_CALC + RET_HELPERS)

    pay = wb.Worksheets("Payouts")
    ldata = []
    status = {}
    with open(ROOT / "data" / "payouts.csv", encoding="utf-8") as f:
        for x in csv.DictReader(f):
            status[(x["Month"], x["Sales rep"])] = (x["Status"], x["Paid date"])
    for m in MONTHS:
        for rep in REPS:
            ldata.append((m, rep))
    llo = make_table(pay, "B24", ["Sales rep", "Month"], [(rep, serial(m)) for m, rep in ldata], "Ledger", style="HarborDark")
    add_columns(llo, LEDGER_COLS[2:])
    for i, (m, rep) in enumerate(ldata, 1):
        st, pdte = status[(m.isoformat(), rep)]
        llo.ListColumns("Status").DataBodyRange.Cells(i, 1).Value = st
        if pdte:
            llo.ListColumns("Paid date").DataBodyRange.Cells(i, 1).Value2 = serial(date.fromisoformat(pdte))
    fill_formulas(llo, LEDGER_COLS)
    for n, _, fmt in LEDGER_COLS:
        llo.ListColumns(n).DataBodyRange.NumberFormat = fmt
    fill_formulas(rlo, RET_CALC + [(n, f) for n, f in RET_HELPERS])

    # Returns sheet styling
    for n in RET_INPUTS:
        b = rlo.ListColumns(n).DataBodyRange; b.Locked = False; fill(b, YELLOW)
    rlo.ListColumns("Return date").DataBodyRange.NumberFormat = "dd mmm yyyy"
    v = rlo.ListColumns("Return date").DataBodyRange.Validation
    v.Add(4, 1, 1, "=MIN(MonthList)", "=LastDay"); v.InputMessage = "Date the goods came back or the order was cancelled."
    v = rlo.ListColumns("Product").DataBodyRange.Validation
    v.Add(3, 1, 1, "=ProductList"); v.InputMessage = "Product of the order line that came back."
    v = rlo.ListColumns("Qty returned").DataBodyRange.Validation
    v.Add(1, 1, 7, "1"); v.InputMessage = "Whole number, 1 or more, not more than was sold."
    for n, *_ in RET_CALC:
        fill(rlo.ListColumns(n).DataBodyRange, (247, 249, 252))
    sc = col_letter(rlo, "Status")
    cf_fill(rlo.ListColumns("Status").DataBodyRange, f'=(${sc}8<>"")*(${sc}8<>"OK")', RED_BG, RED_FG).Font.Bold = True
    w = {"Return date": 11, "Order no": 9, "Product": 24, "Qty returned": 8, "Reason": 19, "Line": 6, "Sales rep": 14,
         "Split rep": 13, "Split %": 7, "Order date": 11, "Qty sold": 7, "Returned net": 10, "Booked month": 10,
         "Status": 26, "Split share": 9, "Rep share": 9, "Rep commission": 11, "Rep net": 11, "Split commission": 11,
         "Split net": 10, "Clawback rep": 10, "Clawback split": 10, "Clawback": 10}
    hdrs = RET_INPUTS + [n for n, *_ in RET_CALC] + [n for n, _ in RET_HELPERS]
    for i, h in enumerate(hdrs, 1):
        r.Columns(i).ColumnWidth = w.get(h, 8)
    r.Range("A7").EntireRow.WrapText = True
    r.Rows(7).RowHeight = 30
    r.Range(r.Cells(1, len(RET_INPUTS) + len(RET_CALC) + 1), r.Cells(1, len(hdrs))).EntireColumn.Hidden = True
    r.Activate(); r.Range("A8").Select(); wb.Application.ActiveWindow.FreezePanes = True


# ================================================================================================ Commission (month detail)
REP_ROWS = range(11, 21)   # up to 10 reps


def build_commission(wb):
    c = wb.Worksheets("Commission")
    widths(c, {"A": 1.5, "B": 16, "C": 8, "D": 6, "E": 11, "F": 10, "G": 9, "H": 7, "I": 10, "J": 10, "K": 8,
               "L": 9, "M": 9, "N": 11, "O": 10, "P": 1.5, "Q": 4})
    header(c, "P", "Commission by rep", "")
    put(c, "B3", '=Company&" (fictional)  ·  "&TEXT(SelMonth,"mmmm yyyy")&"  ·  "&Method&" method  ·  pick the month on the Dashboard"')
    nav(c, "Commission")
    c.Rows(7).RowHeight = 20
    put(c, "B7", '="Month: "&TEXT(SelMonth,"mmmm yyyy")'); font(c.Range("B7"), 11, True, GOLD_LT)
    put(c, "E7", ('="Bands "&Method&"  ·  accelerator +"&TEXT(AccelRate,"0.0%")&" above "&TEXT(AccelAt,"0%")&" of target  ·  '
                  'bonus "&TEXT(TargetBonus,"#,##0")&"  ·  manager override "&TEXT(OverrideRate,"0.0%")'))
    font(c.Range("E7"), 9, False, MUTED)

    panel(c, "B9:O22")
    cols = list("BCDEFGHIJKLMNO")
    titles = ("Sales rep", "Region", "Lines", "Credited net", "Target", "Achieved", "Band", "Commission", "Accelerator",
              "Bonus", "Override", "Clawback", "Total pay", "Payout")
    table_header(c, 10, zip(cols, titles), left=("Sales rep", "Region", "Payout"))
    c.Rows(10).RowHeight = 20
    lk = lambda col, r: f'IF($Q{r}=0,0,INDEX(Ledger[{col}],$Q{r}))'
    for n, r in enumerate(REP_ROWS, 1):
        b = f"$B{r}"
        f = {
            "B": f'=IF({n}>ROWS(Reps[Sales rep]),"",INDEX(Reps[Sales rep],{n}))',
            "C": f'=IF({b}="","",INDEX(Reps[Region],{n}))',
            "Q": f'=IF({b}="",0,IFERROR(MATCH({LEDGER_KEY("SelMonth", b)},Ledger[Key],0),0))',
        }
        for col, name in zip("DEFGHIJKLMN", ("Lines", "Credited net", "Target", "Achieved", "Band rate", "Commission",
                                              "Accelerator", "Bonus", "Override", "Clawback", "Total pay")):
            f[col] = f'=IF({b}="","",{lk(name, r)})'
        f["O"] = f'=IF({b}="","",IF($Q{r}=0,"",INDEX(Ledger[Status],$Q{r})))'
        for col, formula in f.items():
            c.Range(f"{col}{r}").Formula = formula
        underline(c.Range(f"B{r}:O{r}"), LINE)
    c.Range("Q11:Q20").Font.Color = rgb(PAGE)
    put(c, "B21", "Total"); font(c.Range("B21:O21"), 10, True, GOLD_LT)
    for col in "DEFIJKLMN":
        put(c, f"{col}21", f"=SUM({col}11:{col}20)")
    put(c, "G21", "=IFERROR(E21/F21,0)")
    b = c.Range("B21:O21").Borders(8); b.LineStyle = 1; b.Color = rgb(GOLD)
    for rng, fmt in (("D11:D21", "#,##0"), ("E11:F21", "#,##0.00"), ("G11:G21", "0.0%"), ("H11:H20", "0.0%"),
                     ("I11:N21", "#,##0.00;-#,##0.00;-")):
        c.Range(rng).NumberFormat = fmt
    c.Range("D11:N21").HorizontalAlignment = -4152
    cf_font(c.Range("G11:G20"), "=N($G11)>=1", GOOD, True)
    cf_font(c.Range("G11:G20"), f"=N($G11)>=AccelAt", GOLD_LT, True)
    cf_font(c.Range("J11:L20"), "=N(J11)>0", GOOD)
    cf_font(c.Range("M11:M20"), "=N($M11)>0", BAD)
    cf_font(c.Range("O11:O20"), '=$O11="Paid"', GOOD, True)
    cf_font(c.Range("O11:O20"), '=$O11="Approved"', AMBER, True)
    font(c.Range("N11:N20"), None, True, IVORY)

    # chart + plan in words
    panel(c, "B24:H41")
    c.Rows(24).RowHeight = 22
    put(c, "B24", "  Credited net sales vs target"); font(c.Range("B24"), 12, False, IVORY, GEORGIA)
    n_reps = len(REPS)
    co = add_chart(c, "B25:H41", 57)
    ch = co.Chart
    s1 = ch.SeriesCollection().NewSeries(); s1.Name = "Credited net"; s1.Values = c.Range(f"E11:E{10 + n_reps}"); s1.XValues = c.Range(f"B11:B{10 + n_reps}")
    s2 = ch.SeriesCollection().NewSeries(); s2.Name = "Target"; s2.Values = c.Range(f"F11:F{10 + n_reps}"); s2.XValues = c.Range(f"B11:B{10 + n_reps}")
    dark_chart(ch); dark_axes(ch, '#,##0,"k"')
    s1.Format.Fill.ForeColor.RGB = rgb(GOLD); s2.Format.Fill.ForeColor.RGB = rgb(GREYB)
    ch.ChartGroups(1).GapWidth = 60; ch.ChartGroups(1).Overlap = -5
    ch.Legend.Position = -4160
    ch.Axes(1).ReversePlotOrder = True
    ch.Axes(1).TickLabels.Font.Color = rgb(IVORY)

    panel(c, "J24:O41")
    put(c, "J24", "  The plan"); font(c.Range("J24"), 12, False, IVORY, GEORGIA)
    table_header(c, 25, (("J", "From"), ("K", "To"), ("L", "Rate"), ("M", "Step")))
    merge(c, "N25:O25", "Paid on", None, -4152); font(c.Range("N25"), 8.5, True, GOLD_LT); fill(c.Range("N25:O25"), BG)
    for k in range(1, 7):
        r = 25 + k
        put(c, f"J{r}", f'=IF({k}>ROWS(Bands[From]),"",INDEX(Bands[From],{k}))', "#,##0")
        put(c, f"K{r}", f'=IF(J{r}="","",IFERROR(INDEX(Bands[From],{k + 1})-0.01,"and up"))', "#,##0")
        put(c, f"L{r}", f'=IF(J{r}="","",INDEX(Bands[Rate],{k}))', "0.0%")
        put(c, f"M{r}", f'=IF(J{r}="","",INDEX(Bands[Step],{k}))', "+0.0%;-0.0%;0.0%")
        merge(c, f"N{r}:O{r}", f'=IF(J{r}="","",IF(Method="Stepped","this slice","all sales"))', None, -4152)
        c.Range(f"J{r}:O{r}").HorizontalAlignment = -4152
    rules = [('="Method: "&Method', GOLD_LT, True),
             ("=INDEX(Plan!$Q$8:$Q$9,MATCH(Method,MethodList,0))", MUTED, False),
             ('="Accelerator: +"&TEXT(AccelRate,"0.0%")&" on credited sales above "&TEXT(AccelAt,"0%")&" of target."', MUTED, False),
             ('="Target bonus "&TEXT(TargetBonus,"#,##0")&" when credited sales reach the monthly target."', MUTED, False),
             ('="Manager override: "&TEXT(OverrideRate,"0.0%")&" of the credited sales of the reps they manage."', MUTED, False),
             ('="Split deals: each rep is credited with their share of the line."', MUTED, False),
             ('="Clawback: returns take back the commission paid on them, in the month of the return."', MUTED, False),
             ('="Commission is worked on net sales (after discount, before tax)."', MUTED, False)]
    for i, (f, color, bold) in enumerate(rules, 33):
        put(c, f"J{i}", f); font(c.Range(f"J{i}"), 9 if not bold else 10, bold, color)

    # held-back lines
    panel(c, "B43:O56")
    put(c, "B43", "  Lines held back: not paid until fixed in Orders"); font(c.Range("B43"), 12, False, IVORY, GEORGIA)
    c.Rows(43).RowHeight = 22
    table_header(c, 44, (("B", "Order no"), ("C", "Sales rep"), ("F", "Product"), ("J", "Qty"), ("K", "Problem")),
                 left=("Order no", "Sales rep", "Product", "Problem"))
    for col in "DEGHILMNO":
        fill(c.Range(f"{col}44"), BG)
    for k in range(1, 11):
        r = 44 + k
        pick = lambda col: f'=IFERROR(INDEX(Orders[{col}],MATCH({k},Orders[Held line],0)),"")'
        put(c, f"B{r}", pick("Order no"))
        merge(c, f"C{r}:E{r}", pick("Sales rep"), None, -4131)
        merge(c, f"F{r}:I{r}", pick("Product"), None, -4131)
        put(c, f"J{r}", pick("Qty"))
        merge(c, f"K{r}:O{r}", pick("Status"), None, -4131)
        font(c.Range(f"K{r}"), None, True, BAD)
        underline(c.Range(f"B{r}:O{r}"), LINE)
    put(c, "B55", '=IF(SUM(Orders[Held])>10,"... and "&(SUM(Orders[Held])-10)&" more: filter Status in Orders","")')
    font(c.Range("B55"), 9, False, MUTED)

    # returns booked this month
    panel(c, "B58:O68")
    put(c, "B58", "  Returns booked this month: commission clawed back"); font(c.Range("B58"), 12, False, IVORY, GEORGIA)
    c.Rows(58).RowHeight = 22
    table_header(c, 59, (("B", "Order no"), ("C", "Sales rep"), ("F", "Product"), ("I", "Qty"), ("J", "Returned"),
                         ("K", "Status"), ("N", "Clawback")), left=("Order no", "Sales rep", "Product", "Status"))
    for col in "DEGHLMO":
        fill(c.Range(f"{col}59"), BG)
    for k in range(1, 9):
        r = 59 + k
        pick = lambda col: f'=IFERROR(INDEX(Returns[{col}],MATCH({k},Returns[Month line],0)),"")'
        put(c, f"B{r}", pick("Order no"))
        merge(c, f"C{r}:E{r}", f'=IFERROR(INDEX(Returns[Sales rep],MATCH({k},Returns[Month line],0))&IF(INDEX(Returns[Split rep],'
                               f'MATCH({k},Returns[Month line],0))="",""," + "&INDEX(Returns[Split rep],MATCH({k},Returns[Month line],0))),"")', None, -4131)
        merge(c, f"F{r}:H{r}", pick("Product"), None, -4131)
        put(c, f"I{r}", pick("Qty returned"))
        put(c, f"J{r}", pick("Returned net"), "#,##0.00")
        merge(c, f"K{r}:M{r}", pick("Status"), None, -4131)
        put(c, f"N{r}", pick("Clawback"), "-#,##0.00;#,##0.00;-")
        cf_font(c.Range(f"K{r}"), f'=($K{r}<>"")*($K{r}<>"OK")', BAD, True)
        cf_font(c.Range(f"K{r}"), f'=$K{r}="OK"', GOOD)
        underline(c.Range(f"B{r}:O{r}"), LINE)
    c.Range("I60:J67").HorizontalAlignment = -4152; c.Range("N60:N67").HorizontalAlignment = -4152

    # data checks
    panel(c, "B70:O80")
    put(c, "B70", "  Data checks"); font(c.Range("B70"), 12, False, IVORY, GEORGIA)
    c.Rows(70).RowHeight = 22
    checks = [("Every order line in the month is either paid or held back",
               '=IF(COUNTIFS(Orders[Month],SelMonth)=COUNTIFS(Orders[Month],SelMonth,Orders[Status],"OK")+COUNTIFS(Orders[Month],SelMonth,Orders[Held],1),"PASS","FAIL")'),
              ("Credited net of all reps = every OK net sale in the month (split deals add up to 100%)",
               '=IF(ROUND(E21-SUMIFS(Orders[Net],Orders[Month],SelMonth,Orders[Status],"OK"),2)=0,"PASS","FAIL")'),
              ("Commission bands start at 0 and go up", '=IF(SUM(Bands[Order ok])=ROWS(Bands[Order ok]),"PASS","FAIL")'),
              ("All hand-worked test cases pass (Test cases sheet)", '=IF(TestsFailed=0,"PASS","FAIL")'),
              ("Every payout marked Paid has a paid date (Payouts)", '=IF(COUNTIF(Ledger[Check],"OK")=ROWS(Ledger[Check]),"PASS","FAIL")'),
              ("Clawbacks in the ledger = clawbacks on Returns (every return falls in a ledger month)",
               '=IF(ROUND(SUM(Ledger[Clawback])-SUMIFS(Returns[Clawback],Returns[Status],"OK"),2)=0,"PASS","FAIL")'),
              ("Ledger: total pay = commission + accelerator + bonus + override - clawback",
               '=IF(ROUND(SUM(Ledger[Total pay])-SUM(Ledger[Commission])-SUM(Ledger[Accelerator])-SUM(Ledger[Bonus])-SUM(Ledger[Override])+SUM(Ledger[Clawback]),2)=0,"PASS","FAIL")'),
              ("What-if bands start at 0 and go up", '=IF(SUM(ScBands[Order ok])=ROWS(ScBands[Order ok]),"PASS","FAIL")')]
    for i, (label, f) in enumerate(checks, 71):
        put(c, f"B{i}", "  " + label); font(c.Range(f"B{i}"), 9.5, False, IVORY)
        merge(c, f"N{i}:O{i}", f)
        pass_fail(c.Range(f"N{i}:O{i}"), dark=True)
        underline(c.Range(f"B{i}:O{i}"), LINE)
    put(c, "Q71", '=COUNTIF($N$71:$N$78,"PASS")')
    wb.Names.Add("ChecksPassed", "=Commission!$Q$71")
    c.Columns("Q").Hidden = True
    ps = c.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$P$80"


# ================================================================================================ Statement (paper)
STMT_FIRST = 46


def build_statement(wb):
    s = wb.Worksheets("Statement")
    widths(s, {"A": 1.5, "B": 11, "C": 9, "D": 22, "E": 24, "F": 6, "G": 10, "H": 9, "I": 12, "J": 1.5, "K": 3, "L": 12})
    header(s, "J", "Rep statement", "Pick a rep and a month in the gold boxes. Print or save as PDF; "
                                    "scripts/export_statements.py exports every rep in one go.", dark=False)
    nav(s, "Statement", printable=False)
    selector(s, "C14:D14", "StmtRep", wb, "=RepList", None, None, dark=False)
    selector(s, "G14:H14", "StmtMonth", wb, "=MonthList", "mmm yyyy", None, dark=False)
    fill(s.Range("A6:J160"), WHITE)
    # the printable document starts at row 8
    fill(s.Range("B8:I11"), BG)
    for r, h in ((8, 18), (9, 30), (10, 18), (11, 8), (12, 2.25), (13, 10), (14, 22), (15, 10)):
        s.Rows(r).RowHeight = h
    put(s, "B8", "  " + BRAND); font(s.Range("B8"), 8, True, GOLD)
    put(s, "B9", "  Commission statement"); font(s.Range("B9"), 20, False, IVORY, GEORGIA)
    put(s, "B10", '="  "&StmtRep&"  ·  "&TEXT(StmtMonth,"mmmm yyyy")&"  ·  "&Method&" method"')
    font(s.Range("B10"), 9.5, False, GOLD_LT)
    fill(s.Range("B12:I12"), GOLD)
    logo = s.Shapes.AddShape(5, s.Range("H8").Left, s.Range("H8").Top + 6, s.Range("H8:I8").Width - 8, s.Range("H8:H10").Height - 10)
    logo.Name = "logo_slot"
    logo.Fill.Visible = False
    logo.Line.ForeColor.RGB = rgb(GOLD); logo.Line.DashStyle = 4; logo.Line.Weight = 1
    tr = logo.TextFrame2.TextRange; tr.Text = "YOUR LOGO"
    tr.Font.Size = 9; tr.Font.Bold = True; tr.Font.Fill.ForeColor.RGB = rgb(GOLD); tr.ParagraphFormat.Alignment = 2
    logo.TextFrame2.VerticalAnchor = 3
    logo.Placement = 2

    put(s, "B14", "Sales rep"); put(s, "F14", "Month")
    for c in ("B14", "F14"):
        font(s.Range(c), 9, True, MUTED); s.Range(c).HorizontalAlignment = -4152 if c == "F14" else -4131
    s.Range("F14:F14").HorizontalAlignment = -4152
    wb.Names.Add("StmtKey", "=Statement!$L$14")
    put(s, "L14", "=" + LEDGER_KEY("StmtMonth", "StmtRep"))
    put(s, "L15", "=IFERROR(MATCH(StmtKey,Ledger[Key],0),0)")
    lk = lambda col: f"IF(Statement!$L$15=0,0,INDEX(Ledger[{col}],Statement!$L$15))"

    # details (left) + attainment gauge (right)
    put(s, "B16", "Your month"); font(s.Range("B16"), 12, False, INK, GEORGIA)
    details = [("Region", '=IFERROR(INDEX(Reps[Region],MATCH(StmtRep,Reps[Sales rep],0)),"not on Plan")', "General"),
               ("Manager", '=IFERROR(IF(INDEX(Reps[Manager],MATCH(StmtRep,Reps[Sales rep],0))="","-",INDEX(Reps[Manager],MATCH(StmtRep,Reps[Sales rep],0))),"-")', "General"),
               ("Monthly target", "=" + lk("Target"), "#,##0.00"),
               ("Credited net sales", "=" + lk("Credited net"), "#,##0.00"),
               ("Target achieved", "=" + lk("Achieved"), "0.0%"),
               ("Band reached", "=" + lk("Band rate"), "0.0%"),
               ("Payout status", f'=IF(Statement!$L$15=0,"",INDEX(Ledger[Status],Statement!$L$15))', "General"),
               ("Paid on", f'=IF(Statement!$L$15=0,"",IF(INDEX(Ledger[Paid date],Statement!$L$15)="","not yet",INDEX(Ledger[Paid date],Statement!$L$15)))', "dd mmm yyyy")]
    for i, (label, f, fmt) in enumerate(details, 17):
        put(s, f"B{i}", label); font(s.Range(f"B{i}"), 9.5, False, (90, 96, 104))
        put(s, f"D{i}", f, fmt); s.Range(f"D{i}").HorizontalAlignment = -4152
        font(s.Range(f"D{i}"), 10, True, INK)
        underline(s.Range(f"B{i}:D{i}"), (225, 220, 210), 1)
    cf_font(s.Range("D23"), '=$D$23="Paid"', (40, 130, 85))
    cf_font(s.Range("D23"), '=$D$23="Approved"', (176, 120, 20))
    # gauge data (hidden column L)
    put(s, "L17", "=MIN(1,MAX(0,D21))"); put(s, "L18", "=1-L17")
    put(s, "L19", '=TEXT(D21,"0%")')
    put(s, "L20", '=IF(D21>=AccelAt,"accelerator reached",IF(D21>=1,"target reached",TEXT(MAX(0,D19-D20),"#,##0")&" to target"))')
    co = s.ChartObjects().Add(s.Range("F16").Left + 10, s.Range("F16").Top, s.Range("F16:I16").Width - 20, s.Range("F16:F24").Height)
    co.Name = "gauge"
    ch = co.Chart
    ch.ChartType = -4120
    ch.PlotVisibleOnly = False          # the gauge data sits in a hidden column
    while ch.SeriesCollection().Count:
        ch.SeriesCollection(1).Delete()
    g = ch.SeriesCollection().NewSeries(); g.Values = s.Range("L17:L18"); g.Name = "Attainment"
    ch.HasTitle = False; ch.HasLegend = False
    ch.ChartArea.Format.Fill.Visible = False; ch.ChartArea.Format.Line.Visible = False
    ch.PlotArea.Format.Fill.Visible = False
    ch.ChartGroups(1).DoughnutHoleSize = 72
    ch.ChartGroups(1).FirstSliceAngle = 0
    g.Points(1).Format.Fill.ForeColor.RGB = rgb(GOLD)
    g.Points(2).Format.Fill.ForeColor.RGB = rgb((232, 228, 220))
    g.Format.Line.Visible = False
    ch.SetElement(0)                    # msoElementChartTitleNone: a one-series doughnut gets an automatic title otherwise
    co.Placement = 2
    gh = s.Shapes.AddTextbox(1, co.Left + co.Width / 2 - 55, co.Top + co.Height / 2 - 24, 110, 30)
    gh.Name = "gauge_pct"
    gh.DrawingObject.Formula = "=Statement!$L$19"
    gh.TextFrame2.TextRange.Font.Size = 20; gh.TextFrame2.TextRange.Font.Name = GEORGIA
    gh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(INK); gh.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    gh.Fill.Visible = False; gh.Line.Visible = False
    gl = s.Shapes.AddTextbox(1, co.Left + co.Width / 2 - 60, co.Top + co.Height / 2 + 5, 120, 30)
    gl.Name = "gauge_note"
    gl.DrawingObject.Formula = "=Statement!$L$20"
    gl.TextFrame2.TextRange.Font.Size = 7.5
    gl.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb((110, 110, 110)); gl.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    gl.Fill.Visible = False; gl.Line.Visible = False
    for shp in (gh, gl):
        shp.Placement = 2

    # receipt-style working
    put(s, "B26", '="How your pay was worked out  ("&Method&")"'); font(s.Range("B26"), 12, False, INK, GEORGIA)
    s.Rows(26).RowHeight = 22
    fill(s.Range("B27:I39"), PAPER)
    net, rate_band = "$D$20", "$D$22"
    for k in range(1, 7):
        r = 26 + k
        frm, nxt, rate = f"INDEX(Bands[From],{k})", f"IFERROR(INDEX(Bands[From],{k + 1}),{net})", f"INDEX(Bands[Rate],{k})"
        part = f"MAX(0,MIN({net},{nxt})-{frm})"
        last = f"{k}=ROWS(Bands[From])"
        desc_st = (f'IF({k}>ROWS(Bands[From]),"",IF({part}<=0,"","Band {k}:  "&TEXT({frm},"#,##0")&IF({last}," and up",'
                   f'" to "&TEXT(INDEX(Bands[From],{k + 1}),"#,##0"))))')
        work_st = f'IF({k}>ROWS(Bands[From]),"",IF({part}<=0,"",TEXT({part},"#,##0.00")&"  ×  "&TEXT({rate},"0.0%")))'
        amt_st = f'IF({k}>ROWS(Bands[From]),"",IF({part}<=0,"",{part}*{rate}))'
        desc_wh = {1: '"All credited sales at the band reached"',
                   2: '"  band reached: "&TEXT(INDEX(Bands[From],MATCH(' + net + ',Bands[From],1)),"#,##0")&" and up"',
                   3: '"  (on Plan, Stepped pays each slice at its own rate)"'}.get(k, '""')
        work_wh = f'TEXT({net},"#,##0.00")&"  ×  "&TEXT({rate_band},"0.0%")' if k == 1 else '""'
        amt_wh = "$I$33" if k == 1 else '""'
        merge(s, f"B{r}:E{r}", f'=IF(Method="Stepped",{desc_st},{desc_wh})', None, -4131)
        merge(s, f"F{r}:H{r}", f'=IF(Method="Stepped",{work_st},{work_wh})', None, -4152)
        put(s, f"I{r}", f'=IF(Method="Stepped",{amt_st},{amt_wh})', "#,##0.00")
        font(s.Range(f"B{r}:I{r}"), 9.5, False, (70, 74, 80))
    lines = [(33, '="Commission"&IF(Method="Stepped","  (sum of the bands, rounded once)","")', '=""', "=" + lk("Commission")),
             (34, '="Accelerator:  sales above "&TEXT(AccelAt,"0%")&" of target ("&TEXT(AccelAt*D19,"#,##0")&")"',
              '=IF(D20>AccelAt*D19,TEXT(D20-AccelAt*D19,"#,##0.00")&"  ×  "&TEXT(AccelRate,"0.0%"),"not reached")', "=" + lk("Accelerator")),
             (35, '="Target bonus  ("&IF(D20>=D19,"target reached","target not reached")&")"', '=""', "=" + lk("Bonus")),
             (36, '="Manager override:  "&TEXT(OverrideRate,"0.0%")&" of your team\'s credited sales"',
              f'=IF(COUNTIF(Reps[Manager],StmtRep)=0,"not a manager","team "&TEXT(SUMIFS(Ledger[Credited net],Ledger[Manager],StmtRep,Ledger[Month],StmtMonth),"#,##0.00")&"  ×  "&TEXT(OverrideRate,"0.0%"))',
              "=" + lk("Override")),
             (37, '="Clawback:  returns booked this month"',
              '=COUNTIFS(Returns[Sales rep],StmtRep,Returns[Booked month],StmtMonth,Returns[Status],"OK")+COUNTIFS(Returns[Split rep],StmtRep,Returns[Booked month],StmtMonth,Returns[Status],"OK")&" line(s), see Returns"',
              "=-" + lk("Clawback")),
             (38, '="TOTAL PAY"', '=IF(D23="Paid","paid "&TEXT(D24,"dd mmm yyyy"),"status: "&D23)', "=" + lk("Total pay"))]
    for r, d, w, a in lines:
        merge(s, f"B{r}:E{r}", d, None, -4131)
        merge(s, f"F{r}:H{r}", w, None, -4152)
        put(s, f"I{r}", a, "#,##0.00;-#,##0.00;0.00")
        font(s.Range(f"B{r}:I{r}"), 9.5, False, INK)
    b = s.Range("B33:I33").Borders(8); b.LineStyle = 4; b.Color = rgb((170, 160, 140))   # dotted, like a till receipt
    font(s.Range("B33:I33"), 10, True, INK)
    b = s.Range("B38:I38").Borders(8); b.LineStyle = -4119; b.Color = rgb(INK)          # double rule above the total
    font(s.Range("B38:I38"), 12, True, INK); font(s.Range("I38"), 13, True, INK, GEORGIA)
    s.Rows(38).RowHeight = 22
    font(s.Range("F38"), 9, True, (110, 110, 110))
    cf_font(s.Range("F38"), '=$D$23="Paid"', (40, 130, 85))
    cf_font(s.Range("I37"), "=$I$37<0", (180, 60, 45))
    s.Range("B27:I39").Font.Name = "Consolas"
    s.Range("B38").Font.Name = BODY; s.Range("I38").Font.Name = GEORGIA
    put(s, "B39", '=IF(ROUND(I38-SUM(I33:I37),2)=0,"  Checked: lines add up to the total","  CHECK: lines do not add up")')
    font(s.Range("B39"), 8, False, (110, 110, 110))

    # year to date
    put(s, "B41", '="Year to date  (Jan to "&TEXT(StmtMonth,"mmm yyyy")&")"'); font(s.Range("B41"), 12, False, INK, GEORGIA)
    s.Rows(41).RowHeight = 22
    ytd = lambda col, extra="": (f'SUMIFS(Ledger[{col}],Ledger[Sales rep],StmtRep,Ledger[Month],">="&DATE(YEAR(StmtMonth),1,1),'
                                 f'Ledger[Month],"<="&StmtMonth{extra})')
    cards = [("B42:C42", "B43:C43", "Credited sales", "=" + ytd("Credited net"), "#,##0"),
             ("D42", "D43", "Total pay", "=" + ytd("Total pay"), "#,##0.00"),
             ("E42", "E43", "Paid so far", "=" + ytd("Total pay", ',Ledger[Status],"Paid"'), "#,##0.00"),
             ("F42:H42", "F43:H43", "Still to pay", "=" + ytd("To pay"), "#,##0.00"),
             ("I42", "I43", "Months on target", f'=COUNTIFS(Ledger[Sales rep],StmtRep,Ledger[Month],">="&DATE(YEAR(StmtMonth),1,1),'
                                                f'Ledger[Month],"<="&StmtMonth,Ledger[Achieved],">=1")&" of "&MONTH(StmtMonth)', "General")]
    for la, va, label, f, fmt in cards:
        merge(s, la, label, None, -4108); font(s.Range(la.split(":")[0]), 8, True, (120, 110, 90))
        merge(s, va, f, fmt, -4108); font(s.Range(va.split(":")[0]), 13, False, INK, GEORGIA)
        box = s.Range(la.split(":")[0] + ":" + va.split(":")[-1])
        fill(box, PAPER)
        for e in (7, 8, 9, 10):
            box.Borders(e).LineStyle = 1; box.Borders(e).Color = rgb((226, 214, 186))
    s.Rows(43).RowHeight = 24
    put(s, "B44", ('="Lines held back for this rep and month: "&COUNTIFS(Orders[Sales rep],StmtRep,Orders[Month],StmtMonth,'
                   'Orders[Status],"<>OK")&"  (not paid until fixed in Orders)"'))
    font(s.Range("B44"), 8.5, True, RED_FG)

    # order lines
    put(s, f"B{STMT_FIRST - 1}", "Order lines credited to you"); font(s.Range(f"B{STMT_FIRST - 1}"), 12, False, INK, GEORGIA)
    s.Rows(STMT_FIRST - 1).RowHeight = 22
    cols = list("BCDEFGHI")
    src = ("Date", "Order no", "Customer", "Product", "Qty", "Net", "Stmt share", "Stmt credit")
    names = ("Date", "Order no", "Customer", "Product", "Qty", "Line net", "Your share", "Credited")
    table_header(s, STMT_FIRST, zip(cols, names), dark=True, left=("Date", "Order no", "Customer", "Product"))
    first, last = STMT_FIRST + 1, STMT_FIRST + STMT_ROWS
    for k, r in enumerate(range(first, last + 1), 1):
        for col, n in zip(cols, src):
            s.Range(f"{col}{r}").Formula = f'=IFERROR(INDEX(Orders[{n}],MATCH({k},Orders[Statement line],0)),"")'
    s.Range(f"B{first}:B{last}").NumberFormat = "dd mmm yyyy"
    s.Range(f"G{first}:G{last}").NumberFormat = "#,##0.00"
    s.Range(f"H{first}:H{last}").NumberFormat = "0%"
    s.Range(f"I{first}:I{last}").NumberFormat = "#,##0.00"
    s.Range(f"B{first}:B{last}").HorizontalAlignment = -4131
    s.Range(f"B{first}:I{last}").Font.Size = 9
    s.Range(f"B{first}:I{last}").Font.Color = rgb(INK)
    stripe = s.Range(f"B{first}:I{last}").FormatConditions.Add(2, None, f'=($B{first}<>"")*ISEVEN(ROW())')
    stripe.Interior.Color = rgb(PAPER)
    put(s, f"B{last + 1}", "Total credited"); put(s, f"I{last + 1}", f"=SUM(I{first}:I{last})", "#,##0.00")
    font(s.Range(f"B{last + 1}:I{last + 1}"), 10, True, INK)
    b = s.Range(f"B{last + 1}:I{last + 1}").Borders(8); b.LineStyle = 1; b.Color = rgb(INK)
    put(s, f"B{last + 2}", "Check: every credited line is listed and adds up to the credited net sales")
    font(s.Range(f"B{last + 2}"), 8.5, False, (110, 110, 110))
    put(s, f"I{last + 2}", (f'=IF(AND(ROUND(I{last + 1}-D20,2)=0,COUNT(I{first}:I{last})=COUNTIFS(Orders[Sales rep],StmtRep,'
                            f'Orders[Month],StmtMonth,Orders[Status],"OK")+COUNTIFS(Orders[Split rep],StmtRep,Orders[Month],StmtMonth,'
                            f'Orders[Status],"OK")),"PASS","FAIL")'))
    pass_fail(s.Range(f"I{last + 2}"))
    wb.Names.Add("StmtCheck", f"=Statement!$I${last + 2}")
    wb.Names.Add("StmtLines", f"=Statement!$B${first}:$I${last}")
    s.Columns("K:L").Hidden = True
    ps = s.PageSetup
    ps.Orientation = 1; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = False
    ps.PrintArea = f"$A$7:$J${last + 2}"
    ps.PrintTitleRows = f"${STMT_FIRST}:${STMT_FIRST}"
    ps.CenterHorizontally = True
    for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
        setattr(ps, m, 24)


# ================================================================================================ What-if
def build_whatif_inputs(wb):
    w = wb.Worksheets("What-if")
    widths(w, {"A": 1.5, "B": 16, "C": 11, "D": 11, "E": 11, "F": 11, "G": 2, "H": 16, "I": 12, "J": 12, "K": 12,
               "L": 10, "M": 1.5})
    header(w, "M", "What-if: change the plan, see the cost",
           "Change the gold cells. Every rep-month of the last 12 months is recalculated with the scenario and compared "
           "with the current plan. Nothing on the other sheets changes.")
    nav(w, "What-if")
    sc = PLAN["scenario"]
    panel(w, "B14:F34")
    w.Rows(14).RowHeight = 6; w.Rows(15).RowHeight = 24
    put(w, "B15", "  Scenario plan"); font(w.Range("B15"), 12.5, False, IVORY, GEORGIA)
    lo = make_table(w, "B16", ["From", "Rate"], sc["bands"], "ScBands", {"From": "#,##0", "Rate": "0.0%"}, "HarborDark")
    band_helpers(lo, "ScBands")
    lo.ListColumns("Step").DataBodyRange.Font.Color = rgb(MUTED)
    lo.ListColumns("Order ok").DataBodyRange.NumberFormat = ";;;"
    lo.ListColumns("Order ok").Range.Cells(1, 1).Font.Color = rgb(BG)
    for n in ("From", "Rate"):
        b = lo.ListColumns(n).DataBodyRange
        b.Locked = False; fill(b, INPUT_DK); font(b, 10, True, GOLD_LT)
    lo.ListColumns("Rate").DataBodyRange.Validation.Add(2, 1, 1, "0", "=1/2")
    put(w, "F16", "Rate now"); font(w.Range("F16"), 10, True, GOLD_LT); fill(w.Range("F16"), BG)
    w.Range("F16").HorizontalAlignment = -4152
    for k in range(1, 5):
        put(w, f"F{16 + k}", f'=IFERROR(INDEX(Bands[Rate],{k}),"")', "0.0%")
        font(w.Range(f"F{16 + k}"), 10, False, MUTED)
    put(w, "F22", "Now"); font(w.Range("F22"), 8.5, True, GOLD_LT); w.Range("F22").HorizontalAlignment = -4152
    inputs = [(23, "Targets change by", sc["target_change"], "ScTargetChange", "+0%;-0%;0%", "0%", "=-1/2", "1"),
              (24, "Accelerator from", sc["accel_at"], "ScAccelAt", "0%", '=TEXT(AccelAt,"0%")', 1, 3),
              (25, "Accelerator rate", sc["accel_rate"], "ScAccelRate", "0.0%", '=TEXT(AccelRate,"0.0%")', "0", "=1/2"),
              (26, "Target bonus", sc["target_bonus"], "ScBonus", "#,##0", '=TEXT(TargetBonus,"#,##0")', 0, 100000)]
    for r, label, val, nm, fmt, now, lo_v, hi_v in inputs:
        put(w, f"B{r}", "  " + label); font(w.Range(f"B{r}"), 9.5, False, IVORY)
        merge(w, f"C{r}:D{r}", val, fmt, -4152)
        cell = w.Range(f"C{r}:D{r}")
        cell.Locked = False; fill(cell, INPUT_DK); font(cell, 10, True, GOLD_LT)
        w.Range(f"C{r}").Validation.Add(2, 1, 1, str(lo_v), str(hi_v))
        put(w, f"F{r}", now); font(w.Range(f"F{r}"), 9, False, MUTED)
        w.Range(f"F{r}").HorizontalAlignment = -4152
        wb.Names.Add(nm, f"='What-if'!$C${r}")
        underline(w.Range(f"B{r}:F{r}"), LINE)
    for i, txt in enumerate(["  Gold cells = the scenario (type here); Now = the current plan.",
                             "  Manager overrides and clawbacks stay as they are.",
                             '="  Method: "&Method&" (switch it on Plan); it applies to both."'], 29):
        put(w, f"B{i}", txt); font(w.Range(f"B{i}"), 8.5, False, MUTED)


def build_whatif(wb):
    w = wb.Worksheets("What-if")
    cards = [("B", "C", "12-MONTH COST NOW", "=SUM(Ledger[Total pay])", "#,##0",
              '="current plan, "&TEXT(SUM(Ledger[Total pay])/SUM(Ledger[Credited net]),"0.0%")&" of sales"', "$", None),
             ("D", "F", "WITH THE SCENARIO", "=SUM(Ledger[What-if pay])", "#,##0",
              '="scenario, "&TEXT(Calc!$B$62,"0.0%")&" of sales"', "◇", None),
             ("H", "I", "DIFFERENCE, 12 MONTHS", "=Calc!$B$60", "+#,##0;-#,##0;0",
              '=TEXT(Calc!$B$61,"+0.0%;-0.0%;0.0%")&" vs the current plan"', "Δ", "-Calc!$B$60"),
             ("J", "L", "SELECTED MONTH", "=Calc!$B$63", "+#,##0;-#,##0;0",
              '=TEXT(SelMonth,"mmm yyyy")&": "&TEXT(Calc!$B$64,"#,##0")&" → "&TEXT(Calc!$B$65,"#,##0")', "◷", "-Calc!$B$63")]
    for a, b, label, val, fmt, note, glyph, sign in cards:
        kpi(w, a, b, 8, label, val, fmt, note, glyph, sign)
    w.Rows(13).RowHeight = 12
    # per-rep table
    panel(w, "H14:L34")
    put(w, "H15", "  12 months per rep: now vs scenario"); font(w.Range("H15"), 12.5, False, IVORY, GEORGIA)
    table_header(w, 16, (("H", "Sales rep"), ("I", "Pay now"), ("J", "Scenario"), ("K", "Difference"), ("L", "Change")), left=("Sales rep",))
    for n, r in enumerate(range(17, 17 + len(REPS)), 1):
        put(w, f"H{r}", f'=INDEX(Reps[Sales rep],{n})')
        put(w, f"I{r}", f"=SUMIFS(Ledger[Total pay],Ledger[Sales rep],H{r})", "#,##0.00")
        put(w, f"J{r}", f"=SUMIFS(Ledger[What-if pay],Ledger[Sales rep],H{r})", "#,##0.00")
        put(w, f"K{r}", f"=J{r}-I{r}", "+#,##0.00;-#,##0.00;0.00")
        put(w, f"L{r}", f"=IFERROR(K{r}/I{r},0)", "+0.0%;-0.0%;0.0%")
        underline(w.Range(f"H{r}:L{r}"), LINE)
    tr = 17 + len(REPS)
    put(w, f"H{tr}", "  Total")
    for col, f, fmt in (("I", f"=SUM(I17:I{tr - 1})", "#,##0.00"), ("J", f"=SUM(J17:J{tr - 1})", "#,##0.00"),
                        ("K", f"=J{tr}-I{tr}", "+#,##0.00;-#,##0.00;0.00"), ("L", f"=IFERROR(K{tr}/I{tr},0)", "+0.0%;-0.0%;0.0%")):
        put(w, f"{col}{tr}", f, fmt)
    font(w.Range(f"H{tr}:L{tr}"), 10, True, GOLD_LT)
    b = w.Range(f"H{tr}:L{tr}").Borders(8); b.LineStyle = 1; b.Color = rgb(GOLD)
    cf_font(w.Range(f"K17:L{tr}"), "=K17>0", BAD)    # more cost = red
    cf_font(w.Range(f"K17:L{tr}"), "=K17<0", GOOD)
    co = add_chart(w, f"H{tr + 2}:L34", 51)
    ch = co.Chart
    s1 = ch.SeriesCollection().NewSeries(); s1.Name = "Now"; s1.Values = w.Range(f"I17:I{tr - 1}"); s1.XValues = w.Range(f"H17:H{tr - 1}")
    s2 = ch.SeriesCollection().NewSeries(); s2.Name = "Scenario"; s2.Values = w.Range(f"J17:J{tr - 1}"); s2.XValues = w.Range(f"H17:H{tr - 1}")
    dark_chart(ch); dark_axes(ch, '#,##0,"k"')
    s1.Format.Fill.ForeColor.RGB = rgb(GREYB); s2.Format.Fill.ForeColor.RGB = rgb(GOLD)
    ch.ChartGroups(1).GapWidth = 60
    ch.Legend.Position = -4160
    ch.Axes(1).TickLabels.Font.Size = 7.5
    # monthly line
    panel(w, "B36:L52")
    w.Rows(36).RowHeight = 6; w.Rows(37).RowHeight = 24
    put(w, "B37", "  Commission cost per month: now vs scenario"); font(w.Range("B37"), 12.5, False, IVORY, GEORGIA)
    co = add_chart(w, "B38:L52", 4)
    ch = co.Chart
    calc = wb.Worksheets("Calc")
    for name, col, color, wdt in (("Now", "H", GREYB, 2.5), ("Scenario", "I", GOLD, 2.75)):
        srs = ch.SeriesCollection().NewSeries(); srs.Name = name
        srs.Values = calc.Range(f"{col}3:{col}14"); srs.XValues = calc.Range("A3:A14")
        srs.Format.Line.ForeColor.RGB = rgb(color); srs.Format.Line.Weight = wdt; srs.MarkerStyle = -4142
    dark_chart(ch); dark_axes(ch, '#,##0')
    ch.Axes(1).TickLabels.NumberFormat = "mmm yy"
    ch.Legend.Position = -4160
    ps = w.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$M$53"


# ================================================================================================ Payouts
def build_payouts(wb):
    p = wb.Worksheets("Payouts")
    widths(p, {"A": 1.5, "B": 19})
    for i in range(3, 3 + 25):
        p.Columns(i).ColumnWidth = 9.6
    header(p, "Z", "Payouts", "Approve and pay each rep's month here: Status and Paid date are the gold columns of the ledger "
                              "below. To pay = everything not yet Paid.")
    nav(p, "Payouts")
    cards = [("B", "C", "TO PAY NOW (APPROVED)", '=SUMIFS(Ledger[To pay],Ledger[Status],"Approved")', "#,##0.00",
              '=COUNTIFS(Ledger[Status],"Approved")&" rep-months approved"', "✓", None),
             ("D", "F", "WAITING FOR APPROVAL", '=SUMIFS(Ledger[To pay],Ledger[Status],"Pending")', "#,##0.00",
              '=COUNTIFS(Ledger[Status],"Pending")&" rep-months pending"', "…", None),
             ("G", "I", "PAID, LAST 12 MONTHS", '=SUMIFS(Ledger[Total pay],Ledger[Status],"Paid")', "#,##0.00",
              '=COUNTIFS(Ledger[Status],"Paid")&" rep-months paid"', "$", None),
             ("J", "L", "PAYOUT CHECKS", '=COUNTIF(Ledger[Check],"<>OK")', '0" issues"',
              '=IF(COUNTIF(Ledger[Check],"<>OK")=0,"every Paid month has a date","fix the Check column")', "!", "Calc!$B$66")]
    for a, b, label, val, fmt, note, glyph, sign in cards:
        kpi(p, a, b, 7, label, val, fmt, note, glyph, sign)
    # status matrix
    panel(p, "B13:O22")
    put(p, "B13", "  Total pay and payout status, per rep and month"); font(p.Range("B13"), 12.5, False, IVORY, GEORGIA)
    p.Rows(13).RowHeight = 24
    put(p, "B14", "Sales rep"); font(p.Range("B14"), 8.5, True, GOLD_LT); fill(p.Range("B14:O14"), BG)
    for j, m in enumerate(MONTHS):
        col = p.Cells(14, 3 + j)
        col.Value2 = serial(m); col.NumberFormat = "mmm yy"; col.HorizontalAlignment = -4152
        font(col, 8.5, True, GOLD_LT)
    put(p, "O14", "To pay"); font(p.Range("O14"), 8.5, True, GOLD_LT); p.Range("O14").HorizontalAlignment = -4152
    for i, rep in enumerate(REPS):
        r = 15 + i
        put(p, f"B{r}", f"=INDEX(Reps[Sales rep],{i + 1})")
        for j in range(len(MONTHS)):
            cell = p.Cells(r, 3 + j)
            ml = p.Cells(14, 3 + j).Address.replace("$", "")
            mcol = ml.rstrip("0123456789")
            cell.Formula = f"=SUMIFS(Ledger[Total pay],Ledger[Sales rep],$B{r},Ledger[Month],{mcol}$14)"
            # status, mirrored in hidden columns AB.. for the colour rules
            sc = p.Cells(r, 28 + j)
            sc.Formula = f'=IFERROR(INDEX(Ledger[Status],MATCH({LEDGER_KEY(mcol + "$14", "$B" + str(r))},Ledger[Key],0)),"")'
        put(p, f"O{r}", f"=SUMIFS(Ledger[To pay],Ledger[Sales rep],$B{r})", "#,##0.00")
    body = p.Range(f"C15:N{14 + len(REPS)}")
    body.NumberFormat = "#,##0"
    body.HorizontalAlignment = -4152
    cf_fill(body, '=AB15="Paid"', (24, 52, 42), GOOD)
    cf_fill(body, '=AB15="Approved"', (66, 52, 22), AMBER)
    cf_fill(body, '=AB15="Pending"', CARD2, IVORY)
    for r in range(15, 15 + len(REPS)):
        underline(p.Range(f"B{r}:O{r}"), PAGE, 2)
    font(p.Range(f"O15:O{14 + len(REPS)}"), 10, True, GOLD_LT)
    lr = 15 + len(REPS) + 1
    legend = [("C", "Paid", (24, 52, 42), GOOD), ("E", "Approved", (66, 52, 22), AMBER), ("G", "Pending", CARD2, IVORY)]
    for col, label, bg, fg in legend:
        c = p.Range(f"{col}{lr}")
        c.Value = label; fill(c, bg); font(c, 8.5, True, fg); c.HorizontalAlignment = -4108
    put(p, f"I{lr}", "Amounts = total pay for that rep and month."); font(p.Range(f"I{lr}"), 8.5, False, MUTED)
    p.Range("AB:AM").EntireColumn.Hidden = True

    # the ledger table
    put(p, "B23", "  Pay ledger: every rep and month (gold = you type)"); font(p.Range("B23"), 12.5, False, IVORY, GEORGIA)
    p.Rows(23).RowHeight = 26
    lo = p.ListObjects("Ledger")
    undark(lo)
    for n in ("Status", "Paid date"):
        b = lo.ListColumns(n).DataBodyRange
        b.Locked = False; fill(b, INPUT_DK); font(b, 10, True, GOLD_LT)
    v = lo.ListColumns("Status").DataBodyRange.Validation
    v.Add(3, 1, 1, "=StatusList"); v.InputMessage = "Pending, Approved or Paid."
    v = lo.ListColumns("Paid date").DataBodyRange.Validation
    v.Add(4, 1, 5, "36526"); v.InputMessage = "The date the pay went out (only when Status = Paid)."
    st = lo.ListColumns("Status").DataBodyRange
    cf_font(st, f'={st.Cells(1, 1).Address.replace("$", "")}="Paid"', GOOD, True)
    cf_font(st, f'={st.Cells(1, 1).Address.replace("$", "")}="Approved"', AMBER, True)
    ck = lo.ListColumns("Check").DataBodyRange
    cf_font(ck, f'={ck.Cells(1, 1).Address.replace("$", "")}<>"OK"', BAD, True)
    lo.HeaderRowRange.WrapText = True
    lo.HeaderRowRange.Font.Size = 8.5
    lo.HeaderRowRange.RowHeight = 30
    p.Columns("D").ColumnWidth = 13
    lo.HeaderRowRange.VerticalAlignment = -4108
    lo.HeaderRowRange.HorizontalAlignment = -4108
    lo.ListColumns("Key").Range.EntireColumn.Hidden = True
    p.Activate(); p.Range("A25").Select(); wb.Application.ActiveWindow.FreezePanes = False


# ================================================================================================ Calc (hidden)
def build_calc(wb):
    k = wb.Worksheets("Calc")
    k.Range("A1").Value = "Helpers for the Dashboard and What-if charts (hidden sheet, all formulas)."
    # trend: month x rep total pay; H = all reps now; I = all reps scenario
    k.Range("A2").Value = "Month"
    for j, rep in enumerate(REPS):
        k.Cells(2, 2 + j).Formula = f"=INDEX(Reps[Sales rep],{j + 1})"
    k.Range("H2").Value = "Now"; k.Range("I2").Value = "Scenario"; k.Range("J2").Value = "Sales"
    for i, m in enumerate(MONTHS):
        r = 3 + i
        k.Range(f"A{r}").Value2 = serial(m); k.Range(f"A{r}").NumberFormat = "mmm yy"
        for j in range(len(REPS)):
            col = k.Cells(2, 2 + j).Address.replace("$", "").rstrip("0123456789")
            k.Cells(r, 2 + j).Formula = f"=SUMIFS(Ledger[Total pay],Ledger[Month],$A{r},Ledger[Sales rep],{col}$2)"
        k.Range(f"H{r}").Formula = f"=SUMIFS(Ledger[Total pay],Ledger[Month],$A{r})"
        k.Range(f"I{r}").Formula = f"=SUMIFS(Ledger[What-if pay],Ledger[Month],$A{r})"
        k.Range(f"J{r}").Formula = f"=SUMIFS(Ledger[Credited net],Ledger[Month],$A{r})"
    # selected month and prior month (rows 20-40)
    rows = [
        (20, "Prev month", "=EDATE(SelMonth,-1)"),
        (21, "Has prev", "=COUNTIF(Ledger[Month],B20)>0"),
        (22, "Sales", "=SUMIFS(Ledger[Credited net],Ledger[Month],SelMonth)"),
        (23, "Sales prev", "=SUMIFS(Ledger[Credited net],Ledger[Month],B20)"),
        (24, "Cost", "=SUMIFS(Ledger[Total pay],Ledger[Month],SelMonth)"),
        (25, "Cost prev", "=SUMIFS(Ledger[Total pay],Ledger[Month],B20)"),
        (26, "Cost %", "=IFERROR(B24/B22,0)"),
        (27, "Cost % prev", "=IFERROR(B25/B23,0)"),
        (28, "On target", '=COUNTIFS(Ledger[Month],SelMonth,Ledger[Achieved],">=1")'),
        (29, "Accelerator reps", '=COUNTIFS(Ledger[Month],SelMonth,Ledger[Accelerator],">0")'),
        (30, "Held lines", "=SUM(Orders[Held])"),
        (31, "Held value", "=SUMIFS(Orders[Net],Orders[Held],1)"),
        (32, "To pay", "=SUM(Ledger[To pay])"),
        (33, "Approved", '=SUMIFS(Ledger[To pay],Ledger[Status],"Approved")'),
        (34, "Pending", '=SUMIFS(Ledger[To pay],Ledger[Status],"Pending")'),
        (35, "Sales sign", "=IF(B21,SIGN(ROUND(B22-B23,2)),0)"),
        (36, "Cost sign", "=IF(B21,-SIGN(ROUND(B24-B25,2)),0)"),        # cost up = red
        (37, "Cost % sign", "=IF(B21,-SIGN(ROUND(B26-B27,4)),0)"),
        (38, "Held sign", "=-SIGN(B30)"),
        (39, "Commission", "=SUMIFS(Ledger[Commission],Ledger[Month],SelMonth)"),
        (40, "Extras", "=SUMIFS(Ledger[Accelerator],Ledger[Month],SelMonth)+SUMIFS(Ledger[Bonus],Ledger[Month],SelMonth)+SUMIFS(Ledger[Override],Ledger[Month],SelMonth)"),
        (41, "Clawback", "=SUMIFS(Ledger[Clawback],Ledger[Month],SelMonth)"),
        (42, "Returns n", '=COUNTIFS(Returns[Booked month],SelMonth,Returns[Status],"OK")'),
        (43, "Accelerator sum", "=SUMIFS(Ledger[Accelerator],Ledger[Month],SelMonth)"),
        (44, "Override sum", "=SUMIFS(Ledger[Override],Ledger[Month],SelMonth)"),
        (45, "Tests total", "=TestsTotal"),
        (46, "Tests passed", "=TestsTotal-TestsFailed"),
        (60, "What-if diff", "=SUM(Ledger[What-if pay])-SUM(Ledger[Total pay])"),
        (61, "What-if diff %", "=IFERROR(B60/SUM(Ledger[Total pay]),0)"),
        (62, "What-if % of sales", "=IFERROR(SUM(Ledger[What-if pay])/SUM(Ledger[Credited net]),0)"),
        (63, "What-if month diff", "=B65-B64"),
        (64, "Month now", "=SUMIFS(Ledger[Total pay],Ledger[Month],SelMonth)"),
        (65, "Month scenario", "=SUMIFS(Ledger[What-if pay],Ledger[Month],SelMonth)"),
        (66, "Payout check sign", '=-SIGN(COUNTIF(Ledger[Check],"<>OK"))'),
    ]
    for r, label, f in rows:
        k.Range(f"A{r}").Value = label
        k.Range(f"B{r}").Formula = f
    k.Range("B20").NumberFormat = "mmm yyyy"
    # per-rep block for the selected month (rows 50-55): rep, attainment, total pay, status, change vs prev, sort key
    k.Range("A49:G49").Value = ("Rep", "Achieved", "Total pay", "Status", "Change", "Sort key", "Rank")
    for i in range(len(REPS)):
        r = 50 + i
        key = LEDGER_KEY("SelMonth", f"$A{r}")
        pkey = LEDGER_KEY("$B$20", f"$A{r}")
        k.Range(f"A{r}").Formula = f"=INDEX(Reps[Sales rep],{i + 1})"
        k.Range(f"B{r}").Formula = f"=IFERROR(INDEX(Ledger[Achieved],MATCH({key},Ledger[Key],0)),0)"
        k.Range(f"C{r}").Formula = f"=IFERROR(INDEX(Ledger[Total pay],MATCH({key},Ledger[Key],0)),0)"
        k.Range(f"D{r}").Formula = f'=IFERROR(INDEX(Ledger[Status],MATCH({key},Ledger[Key],0)),"")'
        k.Range(f"E{r}").Formula = f"=IF($B$21,C{r}-IFERROR(INDEX(Ledger[Total pay],MATCH({pkey},Ledger[Key],0)),0),0)"
        k.Range(f"F{r}").Formula = f"=B{r}-ROW()/1000000"
        k.Range(f"G{r}").Formula = f"=RANK(F{r},$F$50:$F${49 + len(REPS)})"
    # leaderboard sorted by attainment (rows 70-75)
    k.Range("A69:E69").Value = ("Rank", "Rep", "Achieved", "Total pay", "Status")
    for i in range(len(REPS)):
        r = 70 + i
        row = f"MATCH({i + 1},$G$50:$G${49 + len(REPS)},0)"
        k.Range(f"A{r}").Value = i + 1
        k.Range(f"B{r}").Formula = f"=INDEX($A$50:$A${49 + len(REPS)},{row})"
        k.Range(f"C{r}").Formula = f"=INDEX($B$50:$B${49 + len(REPS)},{row})"
        k.Range(f"D{r}").Formula = f"=INDEX($C$50:$C${49 + len(REPS)},{row})"
        k.Range(f"E{r}").Formula = f"=INDEX($D$50:$D${49 + len(REPS)},{row})"
    # insights (rows 80-86)
    n = len(REPS)
    top = f"INDEX($A$50:$A${49 + n},MATCH(MAX($C$50:$C${49 + n}),$C$50:$C${49 + n},0))"
    mover = f"INDEX($A$50:$A${49 + n},MATCH(MAX($E$50:$E${49 + n}),$E$50:$E${49 + n},0))"
    ins = [
        ('="Top earner: "&' + top + '&", "&TEXT(MAX($C$50:$C$55),"#,##0.00")&" ("&TEXT(IFERROR(MAX($C$50:$C$55)/B24,0),"0%")&" of the month\'s pay)"'),
        ('=IF(B21,"Biggest rise vs "&TEXT(B20,"mmm")&": "&' + mover + '&" "&TEXT(MAX($E$50:$E$55),"+#,##0.00;-#,##0.00"),'
         '"First month in the ledger: no prior month to compare")'),
        ('=IF(B29=0,"Accelerator: nobody above "&TEXT(AccelAt,"0%")&" of target this month",'
         'B29&" rep(s) above "&TEXT(AccelAt,"0%")&" of target earn +"&TEXT(B43,"#,##0.00")&" accelerator")'),
        ('="Manager override: "&TEXT(B44,"#,##0.00")&" to team leads; target bonus to "&B28&" rep(s)"'),
        ('=IF(B42=0,"Clawback: no returns booked this month","Clawback: "&B42&" return(s) take back "&TEXT(B41,"#,##0.00"))'),
        ('=IF(B30=0,"Held back: nothing to fix",B30&" order line(s) held back, worth "&TEXT(B31,"#,##0.00")&", not paid until fixed")'),
        ('="To pay: "&TEXT(B33,"#,##0.00")&" approved + "&TEXT(B34,"#,##0.00")&" waiting for approval"'),
    ]
    for i, f in enumerate(ins):
        k.Range(f"A{80 + i}").Formula = f
    # donut source: total pay per rep in the selected month (only positive values make sense in a donut)
    k.Range("A90").Value = "Donut"
    for i in range(n):
        k.Range(f"A{91 + i}").Formula = f"=A{50 + i}"
        k.Range(f"B{91 + i}").Formula = f"=MAX(0,C{50 + i})"


# ================================================================================================ Dashboard
def build_dashboard(wb):
    d = wb.Worksheets("Dashboard")
    d.Columns("A").ColumnWidth = 1.5
    for i in range(2, 37):
        d.Columns(i).ColumnWidth = 3.57
    d.Columns(37).ColumnWidth = 1.5
    header(d, "AK", "Sales commission dashboard", "")
    put(d, "B3", '=Company&" (fictional)  ·  "&TEXT(SelMonth,"mmmm yyyy")&"  ·  "&Method&" bands  ·  12 months of pay in the ledger"')
    nav(d, "Dashboard")
    selector(d, "F7:I7", "SelMonth", wb, "=MonthList", "mmm yyyy", None)
    d.Rows(7).RowHeight = 24
    merge(d, "B7:E7", "MONTH", None, -4131); font(d.Range("B7"), 8.5, True, GOLD)
    put(d, "K7", ('="Every figure follows the month you pick.  Green = better than the month before, red = worse.  '
                  'Method: "&Method&"  ·  accelerator +"&TEXT(AccelRate,"0%")&" above "&TEXT(AccelAt,"0%")&" of target"'))
    font(d.Range("K7"), 9, False, MUTED); d.Range("K7").VerticalAlignment = -4108
    d.Rows(8).RowHeight = 10
    for r, h in ((9, 5), (10, 18), (11, 34), (12, 18), (13, 6), (14, 12)):
        d.Rows(r).RowHeight = h
    K = "Calc!$B$"
    cards = [
        ("B", "F", "NET SALES", f"={K}22", "#,##0",
         f'=IF({K}21,TEXT(IFERROR({K}22/{K}23-1,0),"▲ 0.0%;▼ 0.0%")&" vs "&TEXT({K}20,"mmm"),"first month")', "$", f"{K}35"),
        ("H", "L", "COMMISSION COST", f"={K}24", "#,##0",
         f'=IF({K}21,TEXT(IFERROR({K}24/{K}25-1,0),"▲ 0.0%;▼ 0.0%")&" vs "&TEXT({K}20,"mmm"),"first month")', "Σ", f"{K}36"),
        ("N", "R", "COST % OF SALES", f"={K}26", "0.0%",
         f'=IF({K}21,TEXT({K}27,"0.0%")&" in "&TEXT({K}20,"mmm"),"first month")', "%", f"{K}37"),
        ("T", "X", "REPS AT/ABOVE TARGET", f'={K}28&" of "&ROWS(Reps[Sales rep])', "@",
         f'={K}29&" above "&TEXT(AccelAt,"0%")&" (accelerator)"', "★", None),
        ("Z", "AD", "HELD-BACK LINES", f"={K}30", "0",
         f'=IF({K}30=0,"nothing to fix",TEXT({K}31,"#,##0")&" not paid")', "!", f"{K}38"),
        ("AF", "AJ", "TO PAY (ALL MONTHS)", f"={K}32", "#,##0",
         f'=TEXT({K}33,"#,##0")&" approved, ready to pay"', "⇥", None),
    ]
    for a, b, label, val, fmt, note, glyph, sign in cards:
        kpi(d, a, b, 9, label, val, fmt, note, glyph, sign, GOLD_LT if sign is None else MUTED)

    # trend + donut
    panel(d, "B15:X33"); panel(d, "Z15:AJ33")
    d.Rows(15).RowHeight = 6; d.Rows(16).RowHeight = 22
    put(d, "C16", "Total pay per rep, last 12 months"); font(d.Range("C16"), 12.5, False, IVORY, GEORGIA)
    put(d, "AA16", '="Share of pay, "&TEXT(SelMonth,"mmm yyyy")'); font(d.Range("AA16"), 12.5, False, IVORY, GEORGIA)
    calc = wb.Worksheets("Calc")
    co = add_chart(d, "C17:W32", 4); co.Name = "trend"
    ch = co.Chart
    for j, rep in enumerate(REPS):
        srs = ch.SeriesCollection().NewSeries()
        srs.Name = f"=Calc!${chr(66 + j)}$2"
        srs.Values = calc.Range(f"{chr(66 + j)}3:{chr(66 + j)}14")
        srs.XValues = calc.Range("A3:A14")
        srs.Format.Line.ForeColor.RGB = rgb(REP_COLORS[j]); srs.Format.Line.Weight = 2.25
        srs.MarkerStyle = 8; srs.MarkerSize = 4
        srs.MarkerBackgroundColor = rgb(REP_COLORS[j]); srs.MarkerForegroundColor = rgb(REP_COLORS[j])
    dark_chart(ch); dark_axes(ch, '#,##0')
    ch.Legend.Position = -4107
    ch.Axes(1).TickLabels.NumberFormat = "mmm yy"
    co = add_chart(d, "AA17:AI32", -4120); co.Name = "donut"
    ch = co.Chart
    srs = ch.SeriesCollection().NewSeries()
    srs.Values = calc.Range(f"B91:B{90 + len(REPS)}"); srs.XValues = calc.Range(f"A91:A{90 + len(REPS)}"); srs.Name = "Share of pay"
    dark_chart(ch)
    ch.Legend.Position = -4107; ch.Legend.Font.Size = 8
    ch.ChartGroups(1).DoughnutHoleSize = 58
    srs.Format.Line.ForeColor.RGB = rgb(CARD); srs.Format.Line.Weight = 1.5
    for j in range(len(REPS)):
        srs.Points(j + 1).Format.Fill.ForeColor.RGB = rgb(REP_COLORS[j])
    srs.HasDataLabels = True
    dl = srs.DataLabels()
    dl.ShowValue = False; dl.ShowPercentage = True; dl.ShowCategoryName = False; dl.NumberFormat = "0%"
    dl.Font.Color = rgb(BG); dl.Font.Bold = True; dl.Font.Size = 8
    ch.SetElement(0)                    # msoElementChartTitleNone

    # leaderboard + insights
    panel(d, "B35:R46"); panel(d, "T35:AJ46")
    d.Rows(34).RowHeight = 12; d.Rows(35).RowHeight = 6; d.Rows(36).RowHeight = 22
    put(d, "C36", "Leaderboard: % of target"); font(d.Range("C36"), 12.5, False, IVORY, GEORGIA)
    put(d, "U36", "What the numbers say"); font(d.Range("U36"), 12.5, False, IVORY, GEORGIA)
    hdr = [("C", "C", "#"), ("D", "G", "Sales rep"), ("H", "I", "Target"), ("J", "M", "Gauge (100% = target)"), ("N", "O", "Pay"), ("P", "Q", "Payout")]
    for a, b, t in hdr:
        merge(d, f"{a}37:{b}37", t, None, -4131 if t in ("#", "Sales rep", "Gauge (100% = target)", "Payout") else -4152)
        font(d.Range(f"{a}37"), 8, True, GOLD_LT)
    fill(d.Range("C37:Q37"), BG)
    for i in range(len(REPS)):
        r = 38 + i
        d.Rows(r).RowHeight = 18
        put(d, f"C{r}", f"=Calc!A{70 + i}")
        merge(d, f"D{r}:G{r}", f"=Calc!B{70 + i}", None, -4131)
        merge(d, f"H{r}:I{r}", f"=Calc!C{70 + i}", "0%", -4152)
        merge(d, f"J{r}:M{r}", f"=Calc!C{70 + i}", '0%', -4131)
        merge(d, f"N{r}:O{r}", f"=Calc!D{70 + i}", "#,##0", -4152)
        merge(d, f"P{r}:Q{r}", f"=Calc!E{70 + i}", None, -4131)
        underline(d.Range(f"C{r}:Q{r}"), LINE)
    font(d.Range("C38:C43"), 9, True, GOLD)
    d.Range("C38:Q43").VerticalAlignment = -4108
    font(d.Range("D38:Q43"), 9.5, False, IVORY)
    cf_font(d.Range("H38:I43"), "=$H38>=1", GOOD, True)
    cf_font(d.Range("H38:I43"), "=$H38>=AccelAt", GOLD_LT, True)
    cf_font(d.Range("P38:Q43"), '=$P38="Paid"', GOOD, True)
    cf_font(d.Range("P38:Q43"), '=$P38="Approved"', AMBER, True)
    gauge = d.Range("J38:M43")
    gauge.Font.Color = rgb(CARD)                     # the bar is the gauge; the value is in the Target column
    for r in range(38, 44):
        db = d.Range(f"J{r}:M{r}").FormatConditions.AddDatabar()
        db.MinPoint.Modify(0, 0); db.MaxPoint.Modify(0, 1.5)
        db.BarColor.Color = rgb(GOLD)
        db.BarFillType = 0
        db.ShowValue = False
        db.BarBorder.Type = 0
    # 100% marker: a thin line at two thirds of the gauge (scale 0-150%)
    g0, g1 = d.Range("J38"), d.Range("M43")
    x = g0.Left + (g1.Left + g1.Width - g0.Left) * (1 / 1.5)
    ln = d.Shapes.AddLine(x, g0.Top + 1, x, g1.Top + g1.Height - 1)
    ln.Line.ForeColor.RGB = rgb(IVORY); ln.Line.Weight = 1; ln.Line.DashStyle = 4; ln.Name = "target_marker"
    put(d, "C45", '="Gauge scale 0 to 150% of target; dashed line = 100%."'); font(d.Range("C45"), 7.5, False, MUTED)
    for i in range(7):
        r = 37 + i
        put(d, f"U{r}", "◆"); font(d.Range(f"U{r}"), 7, False, GOLD); d.Range(f"U{r}").HorizontalAlignment = -4108
        merge(d, f"V{r}:AI{r}", f"=Calc!A{80 + i}", None, -4131)
        font(d.Range(f"V{r}"), 9.5, False, IVORY)
        underline(d.Range(f"U{r}:AI{r}"), LINE)
    d.Rows(44).RowHeight = 18
    d.Rows(47).RowHeight = 8
    d.Rows(48).RowHeight = 20
    put(d, "B48", ('=ChecksPassed&" of 8 data checks PASS  ·  "&(TestsTotal-TestsFailed)&" of "&TestsTotal&'
                   '" hand-worked test cases PASS  ·  formulas locked, no macros"'))
    font(d.Range("B48"), 9, True, GOOD)
    cf_font(d.Range("B48"), "=(ChecksPassed<8)+(TestsFailed>0)", BAD)
    ps = d.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$AK$49"


# ================================================================================================ Test cases
COMM_CASES = [  # case, net sales, target, method, expected commission, expected bonus  (bands 0/15k/30k/45k at 2/4/6/8%)
    ("No sales", 0, 20000, "Whole amount", 0.00, 0),
    ("One cent under the first step", 14999.99, 20000, "Whole amount", 300.00, 0),       # 14,999.99 x 2% = 299.9998
    ("Exactly on the 15,000 step", 15000, 20000, "Whole amount", 600.00, 0),             # 15,000 x 4%
    ("Same 15,000, stepped", 15000, 20000, "Stepped", 300.00, 0),                        # 15,000 x 2%
    ("30,000 stepped", 30000, 35000, "Stepped", 900.00, 0),                              # 15,000 x 2% + 15,000 x 4%
    ("Top band, stepped", 52340.50, 40000, "Stepped", 2387.24, 300),                     # 300 + 600 + 900 + 7,340.50 x 8%
    ("Top band, whole amount", 52340.50, 40000, "Whole amount", 4187.24, 300),           # 52,340.50 x 8%
    ("Target met exactly", 25000, 25000, "Whole amount", 1000.00, 300),                  # 25,000 x 4%, bonus
    ("One cent under target", 24999.99, 25000, "Whole amount", 1000.00, 0),              # 999.9996 rounds to 1,000.00, no bonus
]
ACCEL_CASES = [  # case, net, target, expected accelerator  (from 120% of target, +2%)
    ("Below 120% of target", 35000, 30000, 0.00),                  # 35,000 < 36,000
    ("Exactly 120% of target", 36000, 30000, 0.00),                # nothing above the line
    ("1,000 above 120% of target", 49000, 40000, 20.00),           # 49,000 - 48,000 = 1,000 x 2%
    ("Far above target", 60000, 40000, 240.00),                    # 60,000 - 48,000 = 12,000 x 2%
    ("No target set for the rep", 5000, 0, 0.00),                  # no target = no accelerator
]
OVERRIDE_CASES = [  # case, team rep 1 net, team rep 2 net, expected override (1%)
    ("One rep in the team", 22183.31, 0, 221.83),                  # 221.8331
    ("Two reps, half a cent rounds up", 12500.50, 8000.00, 205.01),  # 20,500.50 x 1% = 205.005
    ("Not a manager (no team)", 0, 0, 0.00),
]
CLAW_CASES = [  # case, line net, qty sold, qty returned, split %, rep comm, rep net, split comm, split net,
                # expected returned net, expected clawback rep, expected clawback split
    ("1 of 4 chairs returned, no split", 1156.00, 4, 1, 0, 1924.11, 32068.52, 0, 0, 289.00, 17.34, 0.00),
    ("Split 50/50 deal, 1 of 2 desks back", 1098.00, 2, 1, 0.5, 600.00, 15000.00, 900.00, 30000.00, 549.00, 10.98, 8.24),
    ("Odd cent on a 30% split", 100.01, 1, 1, 0.3, 300.00, 15000.00, 1200.00, 20000.00, 100.01, 1.40, 1.80),
]
TOTAL_CASES = [  # case, commission, accelerator, bonus, override, clawback, expected total pay
    ("Rep above 120% with a return", 3938.26, 18.57, 300, 0, 12.30, 4244.53),
    ("Manager on target with a return", 1924.11, 0, 300, 221.83, 17.34, 2428.60),
]
WHATIF_CASES = [  # case, net, target, method, expected scenario target, commission, accelerator, bonus, pay
    # fixed test scenario: bands 0/15k/30k/45k at 2.5/4.5/6/7.5%, targets +5%, accelerator 3% from 120%, bonus 250
    ("Whole amount, below the raised target", 20000, 25000, "Whole amount", 26250.00, 900.00, 0.00, 0, 900.00),
    ("Stepped, top band, above 120%", 52340.40, 40000, "Stepped", 42000.00, 2500.53, 58.21, 250, 2808.74),
    ("Exactly on the raised target", 31500, 30000, "Whole amount", 31500.00, 1890.00, 0.00, 250, 2140.00),
]
PAYOUT_CASES = [  # case, status, paid date, total pay, expected to pay, expected check
    ("Paid with a date", "Paid", date(2026, 10, 6), 2224.11, 0.00, "OK"),
    ("Paid but no date typed", "Paid", None, 2224.11, 0.00, "Paid date missing"),
    ("Approved, waiting to be paid", "Approved", None, 887.33, 887.33, "OK"),
    ("Pending with a date typed", "Pending", date(2026, 11, 5), 1466.01, 1466.01, "Date but not Paid"),
]
ORDER_CASES = [  # case, date, rep, product, qty, discount, split rep, split %, exp net, tax, rep credit, split credit, status
    ("Discount exactly at the product max", date(2026, 10, 5), "Grace Holloway", "Gel Pens (pack 12)", 30, 0.15, None, None, 249.90, 20.62, 249.90, 0, "OK"),
    ("Discount above the product max", date(2026, 10, 14), "Diego Santos", "A4 Copy Paper (box)", 40, 0.25, None, None, 1260.00, 103.95, 1260.00, 0, "Discount above 10% max"),
    ("Quantity zero", date(2026, 10, 21), "Priya Nair", "Standing Desk", 0, 0, None, None, 0.00, 0.00, 0.00, 0, "Qty must be 1 or more"),
    ("Product not on the price list", date(2026, 10, 27), "Lena Fischer", "Desk Lamp LED", 6, 0, None, None, None, None, None, None, "Unknown product"),
    ("Sales rep name mistyped", date(2026, 9, 18), "Tom Baker", "Ergonomic Chair", 3, 0, None, None, 867.00, 71.53, 867.00, 0, "Unknown sales rep"),
    ("Several units with a discount", date(2026, 10, 9), "Marcus Reid", "Ergonomic Chair", 7, 0.05, None, None, 1921.85, 158.55, 1921.85, 0, "OK"),
    ("Tax of exactly half a cent rounds up", date(2026, 10, 12), "Tom Becker", "Monitor Arm", 2, 0, None, None, 238.00, 19.64, 238.00, 0, "OK"),
    ("Missing order date", None, "Priya Nair", "Laser Printer", 1, 0, None, None, 379.00, 31.27, 379.00, 0, "Missing date"),
    ("Split 50/50 with an odd cent", date(2026, 10, 6), "Grace Holloway", "Gel Pens (pack 12)", 3, 0.05, "Marcus Reid", 0.5, 27.93, 2.30, 13.96, 13.97, "OK"),
    ("Split 30% to a colleague", date(2026, 10, 7), "Lena Fischer", "Standing Desk", 1, 0, "Priya Nair", 0.3, 549.00, 45.29, 384.30, 164.70, "OK"),
    ("Split rep same as the sales rep", date(2026, 10, 8), "Marcus Reid", "Monitor Arm", 4, 0, "Marcus Reid", 0.5, 476.00, 39.27, 238.00, 238.00, "Split rep same as sales rep"),
    ("Split rep with 0%", date(2026, 9, 23), "Tom Becker", "Laser Printer", 2, 0, "Priya Nair", 0, 758.00, 62.54, 758.00, 0, "Split % must be 1% to 99%"),
    ("Split rep not on the plan", date(2026, 10, 9), "Diego Santos", "Toner Cartridge", 1, 0, "Dana White", 0.5, 89.90, 7.42, 44.95, 44.95, "Unknown split rep"),
    ("Split % typed, no split rep", date(2026, 10, 9), "Tom Becker", "Gel Pens (pack 12)", 5, 0, None, 0.4, 49.00, 4.04, 49.00, 0, "Split % without a split rep"),
]


def return_cases():
    """Return-status cases point at real order lines of the sample data (picked here, expected status typed)."""
    with open(ROOT / "data" / "orders.csv", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    planted = json.loads((ROOT / "test" / "planted.json").read_text())
    planted_orders = {p["order"] for p in planted["orders"]}
    names = {p[0] for p in PLAN["products"]}
    ok = [r for r in rows if r["Order no"] not in planted_orders and int(r["Qty"]) >= 3 and r["Product"] in names
          and r["Sales rep"] in REPS and not r["Split rep"] and r["Date"] >= "2026-03-01"]
    a, b = ok[0], ok[1]
    held = next(r for r in rows if r["Order no"] == planted["orders"][0]["order"])
    d = lambda r, plus=0: date.fromordinal(date.fromisoformat(r["Date"]).toordinal() + plus)
    return [  # case, return date, order, product, qty returned, expected status
        ("Part of a line comes back", d(a, 10), a["Order no"], a["Product"], 1, "OK"),
        ("Whole line back the same day", d(b), b["Order no"], b["Product"], int(b["Qty"]), "OK"),
        ("More returned than was sold", d(a, 10), a["Order no"], a["Product"], int(a["Qty"]) + 1, "Qty above qty sold"),
        ("Return dated before the order", d(a, -1), a["Order no"], a["Product"], 1, "Return before order date"),
        ("Order no typed wrong", d(a, 10), "SO-0000", a["Product"], 1, "Order line not found"),
        ("Line that was held back (never paid)", d(held, 5), held["Order no"], held["Product"], 1, "Line was held back (never paid)"),
        ("Quantity zero", d(a, 10), a["Order no"], a["Product"], 0, "Qty must be 1 or more"),
        ("Missing return date", None, a["Order no"], a["Product"], 1, "Missing date"),
    ]


def block_title(t, row, text, note):
    put(t, f"A{row}", text); font(t.Range(f"A{row}"), 12, True, INK, GEORGIA)
    put(t, f"D{row}", note); font(t.Range(f"D{row}"), 8.5, False, (110, 110, 110))


def block_header(t, row, titles):
    t.Range(t.Cells(row, 1), t.Cells(row, len(titles))).Value = tuple(titles)
    h = t.Range(t.Cells(row, 1), t.Cells(row, len(titles)))
    font(h, 8.5, True, GOLD_LT); fill(h, (40, 48, 58)); h.WrapText = True; h.VerticalAlignment = -4108
    t.Rows(row).RowHeight = 30


def build_tests(wb):
    t = wb.Worksheets("Test cases")
    header(t, "U", "Test cases",
           "Worked out by hand (the working is in scripts/build_workbook.py), run through the same formulas as the calculator. "
           "Every Result must say PASS. If the plan changes, update the expected values with it.", dark=False, tc="A")
    nav(t, "Test cases", col="A")
    widths(t, dict(zip([chr(65 + i) for i in range(21)], [4, 32, 11, 11, 13, 11, 15, 11, 15, 11, 11, 11, 11, 11, 11, 11, 13, 22, 11, 11, 9])))
    results = []   # ranges whose cells must all say PASS
    row = 8

    block_title(t, row, "1. Commission and target bonus", "bands 0 / 15,000 / 30,000 / 45,000 at 2 / 4 / 6 / 8%, bonus 300")
    block_header(t, row + 1, ["#", "Case", "Net sales", "Target", "Method", "Expected commission", "Expected bonus",
                              "Calculator commission", "Calculator bonus", "Result"])
    for i, (case, net, target, method, ec, eb) in enumerate(COMM_CASES):
        r = row + 2 + i
        t.Range(f"A{r}:G{r}").Value = (i + 1, case, net, target, method, ec, eb)
        put(t, f"H{r}", "=" + comm(f"C{r}", f"E{r}"))
        put(t, f"I{r}", "=" + bonus(f"C{r}", f"D{r}"))
        put(t, f"J{r}", f'=IF(AND(ROUND(H{r}-F{r},2)=0,ROUND(I{r}-G{r},2)=0),"PASS","FAIL")')
    end = row + 1 + len(COMM_CASES)
    t.Range(f"C{row + 2}:D{end}").NumberFormat = "#,##0.00"; t.Range(f"F{row + 2}:I{end}").NumberFormat = "#,##0.00"
    results.append(f"J{row + 2}:J{end}")
    row = end + 2

    block_title(t, row, "2. Accelerator above 120% of target", "+2% on credited net sales above 120% of the monthly target")
    block_header(t, row + 1, ["#", "Case", "Net sales", "Target", "Expected accelerator", "Calculator", "Result"])
    for i, (case, net, target, exp) in enumerate(ACCEL_CASES):
        r = row + 2 + i
        t.Range(f"A{r}:E{r}").Value = (i + 1, case, net, target, exp)
        put(t, f"F{r}", "=" + accel(f"C{r}", f"D{r}"))
        put(t, f"G{r}", f'=IF(ROUND(F{r}-E{r},2)=0,"PASS","FAIL")')
    end = row + 1 + len(ACCEL_CASES)
    t.Range(f"C{row + 2}:F{end}").NumberFormat = "#,##0.00"
    results.append(f"G{row + 2}:G{end}")
    row = end + 2

    block_title(t, row, "3. Manager override", "1% of the credited net sales of the reps the manager leads")
    block_header(t, row + 1, ["#", "Case", "Team rep 1", "Team rep 2", "Expected override", "Calculator", "Result"])
    for i, (case, a, b, exp) in enumerate(OVERRIDE_CASES):
        r = row + 2 + i
        t.Range(f"A{r}:E{r}").Value = (i + 1, case, a, b, exp)
        put(t, f"F{r}", "=" + override(f"SUM(C{r}:D{r})"))
        put(t, f"G{r}", f'=IF(ROUND(F{r}-E{r},2)=0,"PASS","FAIL")')
    end = row + 1 + len(OVERRIDE_CASES)
    t.Range(f"C{row + 2}:F{end}").NumberFormat = "#,##0.00"
    results.append(f"G{row + 2}:G{end}")
    row = end + 2

    block_title(t, row, "4. Clawback on a return", "returned share x the rate the rep was paid that month (commission / credited net)")
    block_header(t, row + 1, ["#", "Case", "Line net", "Qty sold", "Qty returned", "Split %", "Rep commission", "Rep net",
                              "Split rep commission", "Split rep net", "Expected returned", "Expected clawback rep",
                              "Expected clawback split", "Calculator returned", "Calculator rep", "Calculator split", "Result"])
    for i, (case, ln, qs, qr, sp, rc, rn, sc, sn, er, ecr, ecs) in enumerate(CLAW_CASES):
        r = row + 2 + i
        t.Range(f"A{r}:M{r}").Value = (i + 1, case, ln, qs, qr, sp, rc, rn, sc, sn, er, ecr, ecs)
        put(t, f"N{r}", "=" + returned_net(f"C{r}", f"E{r}", f"D{r}"))
        split = f"IF(N(F{r})=0,0,{split_part(f'N{r}', f'F{r}')})"
        put(t, f"O{r}", "=" + clawback(f"(N{r}-{split})", f"G{r}", f"H{r}"))
        put(t, f"P{r}", "=" + clawback(split, f"I{r}", f"J{r}"))
        put(t, f"Q{r}", f'=IF(AND(ROUND(N{r}-K{r},2)=0,ROUND(O{r}-L{r},2)=0,ROUND(P{r}-M{r},2)=0),"PASS","FAIL")')
    end = row + 1 + len(CLAW_CASES)
    t.Range(f"C{row + 2}:C{end}").NumberFormat = "#,##0.00"; t.Range(f"F{row + 2}:F{end}").NumberFormat = "0%"
    t.Range(f"G{row + 2}:P{end}").NumberFormat = "#,##0.00"
    results.append(f"Q{row + 2}:Q{end}")
    row = end + 2

    block_title(t, row, "5. Total pay", "commission + accelerator + target bonus + manager override - clawback")
    block_header(t, row + 1, ["#", "Case", "Commission", "Accelerator", "Bonus", "Override", "Clawback", "Expected total pay",
                              "Calculator", "Result"])
    for i, (case, c_, a_, b_, o_, k_, exp) in enumerate(TOTAL_CASES):
        r = row + 2 + i
        t.Range(f"A{r}:H{r}").Value = (i + 1, case, c_, a_, b_, o_, k_, exp)
        put(t, f"I{r}", f"=C{r}+D{r}+E{r}+F{r}-G{r}")
        put(t, f"J{r}", f'=IF(ROUND(I{r}-H{r},2)=0,"PASS","FAIL")')
    end = row + 1 + len(TOTAL_CASES)
    t.Range(f"C{row + 2}:I{end}").NumberFormat = "#,##0.00"
    results.append(f"J{row + 2}:J{end}")
    row = end + 2

    block_title(t, row, "6. What-if (fixed test scenario)", "bands at 2.5 / 4.5 / 6 / 7.5%, targets +5%, accelerator 3% from 120%, bonus 250")
    tsb = make_table(t, f"S{row}", ["From", "Rate"], [(0, 0.025), (15000, 0.045), (30000, 0.06), (45000, 0.075)], "TestScBands",
                     {"From": "#,##0", "Rate": "0.0%"}, "TableStyleLight9")
    band_helpers(tsb, "TestScBands")
    t.Range(f"S{row + 6}:U{row + 9}").Value = (("Target change", 0.05, None), ("Accelerator from", 1.2, None),
                                               ("Accelerator rate", 0.03, None), ("Target bonus", 250, None))
    wb.Names.Add("TestScChange", f"='Test cases'!$T${row + 6}")
    wb.Names.Add("TestScAccelAt", f"='Test cases'!$T${row + 7}")
    wb.Names.Add("TestScAccelRate", f"='Test cases'!$T${row + 8}")
    wb.Names.Add("TestScBonus", f"='Test cases'!$T${row + 9}")
    block_header(t, row + 1, ["#", "Case", "Net sales", "Target", "Method", "Expected target", "Expected commission",
                              "Expected accelerator", "Expected bonus", "Expected pay", "Calc target", "Calc commission",
                              "Calc accelerator", "Calc bonus", "Calc pay", "Result"])
    for i, (case, net, target, method, et, ec, ea, eb, ep) in enumerate(WHATIF_CASES):
        r = row + 2 + i
        t.Range(f"A{r}:J{r}").Value = (i + 1, case, net, target, method, et, ec, ea, eb, ep)
        put(t, f"K{r}", f"=ROUND(D{r}*(1+TestScChange),2)")
        put(t, f"L{r}", "=" + comm(f"C{r}", f"E{r}", "TestScBands"))
        put(t, f"M{r}", "=" + accel(f"C{r}", f"K{r}", "TestScAccelAt", "TestScAccelRate"))
        put(t, f"N{r}", "=" + bonus(f"C{r}", f"K{r}", "TestScBonus"))
        put(t, f"O{r}", f"=L{r}+M{r}+N{r}")
        put(t, f"P{r}", f'=IF(AND(ROUND(K{r}-F{r},2)=0,ROUND(L{r}-G{r},2)=0,ROUND(M{r}-H{r},2)=0,ROUND(N{r}-I{r},2)=0,'
                        f'ROUND(O{r}-J{r},2)=0),"PASS","FAIL")')
    end = row + 1 + len(WHATIF_CASES)
    t.Range(f"C{row + 2}:D{end}").NumberFormat = "#,##0.00"; t.Range(f"F{row + 2}:O{end}").NumberFormat = "#,##0.00"
    results.append(f"P{row + 2}:P{end}")
    row = max(end, row + 10) + 2

    block_title(t, row, "7. Payout status", "To pay = everything not yet Paid; Paid needs a date and only Paid has one")
    block_header(t, row + 1, ["#", "Case", "Status", "Paid date", "Total pay", "Expected to pay", "Expected check",
                              "Calculator to pay", "Calculator check", "Result"])
    for i, (case, st, pdte, tot, etp, eck) in enumerate(PAYOUT_CASES):
        r = row + 2 + i
        t.Range(f"A{r}:C{r}").Value = (i + 1, case, st)
        if pdte:
            t.Range(f"D{r}").Value2 = serial(pdte)
        t.Range(f"E{r}:G{r}").Value = (tot, etp, eck)
        put(t, f"H{r}", "=" + to_pay(f"C{r}", f"E{r}"))
        put(t, f"I{r}", "=" + payout_check(f"C{r}", f"D{r}"))
        put(t, f"J{r}", f'=IF(AND(ROUND(H{r}-F{r},2)=0,I{r}=G{r}),"PASS","FAIL")')
    end = row + 1 + len(PAYOUT_CASES)
    t.Range(f"D{row + 2}:D{end}").NumberFormat = "dd mmm yyyy"
    t.Range(f"E{row + 2}:F{end}").NumberFormat = "#,##0.00"; t.Range(f"H{row + 2}:H{end}").NumberFormat = "#,##0.00"
    results.append(f"J{row + 2}:J{end}")
    row = end + 2

    # order lines: a table with the same column names as Orders, so the very same formulas fill it
    block_title(t, row, "8. Order lines, discounts, tax and split deals", "same formulas as the Orders sheet")
    top = row + 1
    hdr2 = ["#", "Case", "Date", "Sales rep", "Product", "Qty", "Discount %", "Split rep", "Split %"] + \
           [n for n, _, _ in ORDER_CALC] + ["Expected net", "Expected tax", "Expected rep credit", "Expected split credit",
                                            "Expected status", "Result"]
    t.Range(t.Cells(top, 1), t.Cells(top, len(hdr2))).Value = tuple(hdr2)
    lo = t.ListObjects.Add(1, t.Range(t.Cells(top, 1), t.Cells(top + len(ORDER_CASES), len(hdr2))), None, 1)
    lo.Name = "OrderTests"; lo.TableStyle = "TableStyleLight9"
    for i, (case, d, rep, prod, qty, disc, srep, spct, en, et, erc, esc, es) in enumerate(ORDER_CASES, 1):
        r = top + i
        t.Range(f"A{r}:B{r}").Value = (i, case)
        if d:
            t.Range(f"C{r}").Value2 = serial(d)
        t.Range(f"D{r}:G{r}").Value = (rep, prod, qty, disc)
        if srep:
            t.Range(f"H{r}").Value = srep
        if spct is not None:
            t.Range(f"I{r}").Value = spct
        for col, val in (("Expected net", en), ("Expected tax", et), ("Expected rep credit", erc), ("Expected split credit", esc)):
            if val is not None:
                lo.ListColumns(col).DataBodyRange.Cells(i, 1).Value = val
        lo.ListColumns("Expected status").DataBodyRange.Cells(i, 1).Value = es
    for name, f, fmt in ORDER_CALC:
        lo.ListColumns(name).DataBodyRange.Formula = f
        lo.ListColumns(name).DataBodyRange.NumberFormat = fmt
    lo.ListColumns("Result").DataBodyRange.Formula = (
        '=IF(AND(ROUND(N([@Net])-N([@[Expected net]]),2)=0,ROUND(N([@Tax])-N([@[Expected tax]]),2)=0,'
        'ROUND(N([@[Rep credit]])-N([@[Expected rep credit]]),2)=0,ROUND(N([@[Split credit]])-N([@[Expected split credit]]),2)=0,'
        '[@Status]=[@[Expected status]]),"PASS","FAIL")')
    lo.ListColumns("Date").DataBodyRange.NumberFormat = "dd mmm yyyy"
    for n in ("Discount %", "Split %"):
        lo.ListColumns(n).DataBodyRange.NumberFormat = "0%"
    for n in ("Expected net", "Expected tax", "Expected rep credit", "Expected split credit"):
        lo.ListColumns(n).DataBodyRange.NumberFormat = "#,##0.00"
    lo.HeaderRowRange.WrapText = True; t.Rows(top).RowHeight = 30
    results.append(lo.ListColumns("Result").DataBodyRange.Address.replace("$", ""))
    row = top + len(ORDER_CASES) + 2

    block_title(t, row, "9. Returns", "same formulas as the Returns sheet, on real order lines of the sample data")
    top = row + 1
    rcases = return_cases()
    hdr3 = ["#", "Case"] + RET_INPUTS[:4] + [n for n, *_ in RET_CALC] + ["Expected status", "Result"]
    t.Range(t.Cells(top, 1), t.Cells(top, len(hdr3))).Value = tuple(hdr3)
    lo = t.ListObjects.Add(1, t.Range(t.Cells(top, 1), t.Cells(top + len(rcases), len(hdr3))), None, 1)
    lo.Name = "ReturnTests"; lo.TableStyle = "TableStyleLight9"
    for i, (case, rd, order, prod, qty, es) in enumerate(rcases, 1):
        r = top + i
        t.Range(f"A{r}:B{r}").Value = (i, case)
        if rd:
            t.Range(f"C{r}").Value2 = serial(rd)
        t.Range(f"D{r}:F{r}").Value = (order, prod, qty)
        lo.ListColumns("Expected status").DataBodyRange.Cells(i, 1).Value = es
    for name, f, fmt in RET_CALC:
        lo.ListColumns(name).DataBodyRange.Formula = f
        lo.ListColumns(name).DataBodyRange.NumberFormat = fmt
    lo.ListColumns("Return date").DataBodyRange.NumberFormat = "dd mmm yyyy"
    lo.ListColumns("Result").DataBodyRange.Formula = '=IF([@Status]=[@[Expected status]],"PASS","FAIL")'
    lo.HeaderRowRange.WrapText = True; t.Rows(top).RowHeight = 30
    results.append(lo.ListColumns("Result").DataBodyRange.Address.replace("$", ""))
    row = top + len(rcases) + 2

    for rng in results:
        pass_fail(t.Range(rng))
    total = sum(t.Range(r).Cells.Count for r in results)
    fail_f = "+".join(f'COUNTIF({r},"FAIL")' for r in results)
    pass_f = "+".join(f'COUNTIF({r},"PASS")' for r in results)
    put(t, "A6", f'="Cases passed: "&({pass_f})&" of {total}"'); font(t.Range("A6"), 12, True, INK)
    put(t, f"A{row}", f"={fail_f}"); t.Range(f"A{row}").Font.Color = rgb(WHITE)
    put(t, f"B{row}", f"={total}"); t.Range(f"B{row}").Font.Color = rgb(WHITE)
    wb.Names.Add("TestsFailed", f"='Test cases'!$A${row}")
    wb.Names.Add("TestsTotal", f"='Test cases'!$B${row}")
    t.Rows(6).RowHeight = 22
    t.Range("A7").RowHeight = 4


# ================================================================================================ How to use
def build_howto(wb):
    ws = wb.Worksheets("How to use")
    header(ws, "C", "How to use", "Daily steps for the sales team, month-end steps for the manager.", dark=False)
    nav(ws, "How to use")
    ws.Columns("A").ColumnWidth = 1.5
    ws.Columns("B").ColumnWidth = 120
    text = [("Every day (sales team)", True),
            ("1. Orders: type or paste each order line in the yellow columns. Rep and product come from drop-downs.", False),
            ("2. A deal won together with a colleague: pick the colleague in Split rep and their share in Split % (1% to 99%).", False),
            ("3. Price, discount, tax and total are calculated and locked. Status must say OK; red = fix that line.", False),
            ("4. Goods returned or an order cancelled: add a line on Returns (date, order no, product, qty). The commission paid on it is clawed back.", False),
            ("", False), ("Month end (sales manager)", True),
            ("1. Dashboard: pick the month. Fix every line held back (Commission sheet), then the numbers are final.", False),
            ("2. Payouts: set the month's Status to Approved; when the money goes out, set Paid and type the paid date.", False),
            ("3. Statement: pick each rep and print or save as PDF. scripts/export_statements.py exports every rep at once (optional macro: scripts/ExportStatements.bas).", False),
            ("4. Check that the 8 data checks on Commission say PASS (the Dashboard footer shows the count).", False),
            ("", False), ("Changing the plan", True),
            ("Plan is not locked: prices, max discounts, reps, managers, targets, bands, tax, bonus, override, accelerator and the method.", False),
            ("Whole amount = all credited sales at the band reached. Stepped = each slice of sales at its own band rate.", False),
            ("Accelerator = extra rate on the sales above a share of target (120% here). Override = a share of the sales of the reps a manager leads.", False),
            ("Try a change first on What-if: it recalculates all 12 months with the scenario and shows the cost difference.", False),
            ("After a plan change, update the expected values on Test cases so they still describe the agreed plan.", False),
            ("", False), ("Adding a month or a rep", True),
            ("Add the month to Plan > Months and one row per rep on the Payouts ledger (Month, Sales rep, Status). Add reps to Plan > Sales reps.", False),
            ("", False), ("Protection", True),
            ("Formula sheets are protected without a password (Review > Unprotect Sheet). Only the yellow / gold cells can be typed in.", False),
            ("Orders has room for 3,000 lines and Returns for 60; for more, unprotect and drag the table's corner down.", False),
            ("Drop-downs and validation stop typing mistakes; pasted data skips validation, which is why Status checks every line.", False),
            ("", False), ("Built and tested in Excel for Windows with Excel 2016 formulas only. No macros in this file, no Power Query.", False)]
    for i, (s, bold) in enumerate(text, 8):
        ws.Cells(i, 2).Value = s
        font(ws.Cells(i, 2), 12 if bold else 10, bold, INK, GEORGIA if bold else None)


if __name__ == "__main__":
    main()
