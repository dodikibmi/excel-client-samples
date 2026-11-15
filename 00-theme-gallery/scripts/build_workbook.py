"""Builds Theme-Gallery.xlsx with Excel COM: one sales dashboard drawn in the 8 colour themes of themes.py.

Sheets: Themes (tiles, click to open) · one dashboard per theme · Palette (every hex code) · How to use ·
Calc (every number, plain SUMIFS) · Data (the Sales table). No macros.
Run under the shared lock:  python tools/with_lock.py E0 -- <python> scripts/build_workbook.py
"""
import csv, os, sys
from datetime import date
from pathlib import Path
import win32com.client as win32

sys.path.insert(0, os.path.dirname(__file__))
from themes import THEMES, ROLES, SWATCHES, sheet_name, luminance, is_dark

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Theme-Gallery.xlsx"
REPORT = date(2026, 10, 1)
COMPANY = "Harbourline Supply Co. (fictional)"
REGIONS = ["North", "South", "East", "West"]
CATEGORIES = ["Fasteners", "Power Tools", "Paint", "Garden", "Electrical"]
CARD_COLS = [("B", "F"), ("H", "L"), ("N", "R"), ("T", "X")]  # four card columns, G/M/S are narrow gaps
SEP = ","  # local list separator (conditional formats are read in local syntax), set in main()
SWATCH_COLS = ["N", "O", "P", "Q", "R", "T", "U", "V", "W", "X"]


def serial(d):
    return (d - date(1899, 12, 30)).days


def c(h):
    h = h.lstrip("#")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


def fill(rng, h):
    rng.Interior.Color = c(h)


def font(rng, h, size=None, bold=None, name=None, italic=None):
    f = rng.Font
    f.Color = c(h)
    if size: f.Size = size
    if bold is not None: f.Bold = bold
    if name: f.Name = name
    if italic is not None: f.Italic = italic


def box(rng, h, weight=2):
    for edge in (7, 8, 9, 10):
        rng.Borders(edge).LineStyle = 1
        rng.Borders(edge).Weight = weight
        rng.Borders(edge).Color = c(h)


def window(xl, ws, headings=False, zoom=100):
    ws.Activate()
    w = xl.ActiveWindow
    w.DisplayGridlines = False
    w.DisplayHeadings = headings
    w.Zoom = zoom
    w.ScrollRow = 1
    w.ScrollColumn = 1
    ws.Range("A1").Select()


def page_setup(ws, area):
    ps = ws.PageSetup
    ps.PrintArea = area
    ps.Orientation = 2
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    ps.CenterHorizontally = True
    for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin", "HeaderMargin", "FooterMargin"):
        setattr(ps, m, 14)


# ---------------------------------------------------------------- data + calc
def load_rows():
    with (ROOT / "data" / "sales.csv").open(encoding="utf-8") as f:
        r = list(csv.reader(f))
    out = []
    for row in r[1:]:
        y, m, _ = map(int, row[0].split("-"))
        out.append([serial(date(y, m, 1)), row[1], row[2]] + [int(v) for v in row[3:]])
    return r[0], out


def build_data(wb):
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Data"
    head, rows = load_rows()
    n = len(rows)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 7)).Value = [head]
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 7)).Value = rows
    lo = ws.ListObjects.Add(1, ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, 7)), None, 1)
    lo.Name = "Sales"
    lo.TableStyle = "TableStyleLight9"
    ws.Range("A:A").NumberFormat = "mmm yyyy"
    ws.Range("D:F").NumberFormat = "#,##0"
    ws.Columns("A:G").AutoFit()
    ws.Tab.Color = c("#8E98A3")
    return ws


def build_calc(wb):
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Calc"
    S = lambda col, *crit: f"SUMIFS(Sales[{col}]" + "".join(f",Sales[{k}],{v}" for k, v in crit) + ")"
    ws.Range("A1").Value = "Reporting month"
    ws.Range("C1").Value = serial(REPORT)
    ws.Range("C1").NumberFormat = "mmm yyyy"
    ws.Range("A3:E3").Value = [["Month", "Actual", "Budget", "Prior year", "Label"]]
    for k in range(12):
        r = 4 + k
        ws.Range(f"A{r}").Formula = f"=EDATE($C$1,{k - 11})"
        ws.Range(f"B{r}").Formula = "=" + S("Revenue", ("Month", f"$A{r}"))
        ws.Range(f"C{r}").Formula = "=" + S("Budget", ("Month", f"$A{r}"))
        ws.Range(f"D{r}").Formula = "=" + S("Revenue", ("Month", f"EDATE($A{r},-12)"))
        ws.Range(f"E{r}").Formula = f'=TEXT(A{r},"mmm")&" "&TEXT(A{r},"yy")'
    ws.Range("A4:A15").NumberFormat = "mmm yyyy"
    ws.Range("A17:D17").Value = [["Category", "Revenue", "Share", "Legend label"]]
    for k, cat in enumerate(CATEGORIES):
        r = 18 + k
        ws.Range(f"A{r}").Value = cat
        ws.Range(f"B{r}").Formula = "=" + S("Revenue", ("Month", "$C$1"), ("Category", f"$A{r}"))
        ws.Range(f"C{r}").Formula = f"=B{r}/SUM($B$18:$B$22)"
        ws.Range(f"D{r}").Formula = f'=A{r}&"  "&TEXT(C{r},"0%")'
    ws.Range("A24:G24").Value = [["Region", "Actual", "Budget", "Variance", "Var %", "Margin", "Share"]]
    for k, reg in enumerate(REGIONS + ["Total"]):
        r = 25 + k
        ws.Range(f"A{r}").Value = reg
        if reg == "Total":
            for col in "BCD":
                ws.Range(f"{col}{r}").Formula = f"=SUM({col}25:{col}28)"
            ws.Range(f"F{r}").Formula = "=1-" + S("Cost", ("Month", "$C$1")) + f"/B{r}"
        else:
            ws.Range(f"B{r}").Formula = "=" + S("Revenue", ("Month", "$C$1"), ("Region", f"$A{r}"))
            ws.Range(f"C{r}").Formula = "=" + S("Budget", ("Month", "$C$1"), ("Region", f"$A{r}"))
            ws.Range(f"D{r}").Formula = f"=B{r}-C{r}"
            ws.Range(f"F{r}").Formula = "=1-" + S("Cost", ("Month", "$C$1"), ("Region", f"$A{r}")) + f"/B{r}"
        ws.Range(f"E{r}").Formula = f"=B{r}/C{r}-1"
        ws.Range(f"G{r}").Formula = f"=B{r}/$B$29"
    kpis = [  # row 32..
        ("Revenue", "=" + S("Revenue", ("Month", "$C$1"))),
        ("Budget", "=" + S("Budget", ("Month", "$C$1"))),
        ("Revenue vs budget", "=B32/B33-1"),
        ("Revenue prior year", "=" + S("Revenue", ("Month", "EDATE($C$1,-12)"))),
        ("Revenue vs prior year", "=B32/B35-1"),
        ("Gross profit", "=B32-" + S("Cost", ("Month", "$C$1"))),
        ("Gross margin", "=B37/B32"),
        ("Gross margin prior year", "=1-" + S("Cost", ("Month", "EDATE($C$1,-12)")) + "/B35"),
        ("Margin change (pts)", "=(B38-B39)*100"),
        ("Orders", "=" + S("Orders", ("Month", "$C$1"))),
        ("Orders prior year", "=" + S("Orders", ("Month", "EDATE($C$1,-12)"))),
        ("Orders vs prior year", "=B41/B42-1"),
        ("Average order", "=B32/B41"),
        ("Average order prior year", "=B35/B42"),
        ("Average order vs prior year", "=B44/B45-1"),
    ]
    ws.Range("A31:C31").Value = [["KPI", "Value", "Shown as"]]
    for k, (label, f) in enumerate(kpis):
        ws.Range(f"A{32 + k}").Value = label
        ws.Range(f"B{32 + k}").Formula = f
    ws.Range("C32").Formula = '=TEXT(B32,"#,##0")'
    ws.Range("C33").Formula = '="revenue, "&TEXT($C$1,"mmm yyyy")'
    arrow = lambda ref, fmt, tail: (f'=IF(ROUND({ref},3)=0,"● flat{tail}",IF({ref}>=0,"▲ ","▼ ")'
                                    f'&TEXT(ABS({ref}),"{fmt}")&"{tail}")')
    ws.Range("C34").Formula = arrow("B34", "0.0%", " vs budget")
    ws.Range("C36").Formula = arrow("B36", "0.0%", " vs prior year")
    ws.Range("C40").Formula = arrow("B40", "0.0", " pts vs prior year")
    ws.Range("C43").Formula = arrow("B43", "0.0%", " vs prior year")
    ws.Range("C46").Formula = arrow("B46", "0.0%", " vs prior year")
    ws.Range("C37").Formula = '="gross profit "&TEXT(B37,"#,##0")'
    ws.Range("A48").Value = "Subtitle"
    ws.Range("B48").Formula = f'="{COMPANY}  |  "&TEXT($C$1,"mmmm yyyy")&"  |  all regions  |  USD"'
    for rng, fmt in (("B4:D15", "#,##0"), ("B18:B22", "#,##0"), ("C18:C22", "0.0%"), ("B25:D29", "#,##0"),
                     ("E25:G29", "0.0%"), ("B32:B33", "#,##0"), ("B35", "#,##0"), ("B37", "#,##0"),
                     ("B34,B36,B38,B39,B43,B46", "0.0%"), ("B40", "0.00"), ("B41:B42", "#,##0"), ("B44:B45", "#,##0.0")):
        for part in rng.split(","):
            ws.Range(part).NumberFormat = fmt
    for rr in ("A3:E3", "A17:D17", "A24:G24", "A31:C31"):
        ws.Range(rr).Font.Bold = True
    ws.Columns("A:G").AutoFit()
    ws.Tab.Color = c("#8E98A3")
    return ws


# ---------------------------------------------------------------- themed dashboard
def button(ws, t, x, y, w, h, label, target, strong=False):
    shp = ws.Shapes.AddShape(5, x, y, w, h)
    shp.Adjustments.SetItem(1, 0.18)
    shp.Fill.ForeColor.RGB = c(t["accent"] if strong else t["card"])
    shp.Line.ForeColor.RGB = c(t["accent"] if strong else t["line"])
    shp.Line.Weight = 1
    tf = shp.TextFrame2
    tf.MarginLeft = tf.MarginRight = 2
    tf.MarginTop = tf.MarginBottom = 0
    tf.VerticalAnchor = 3
    tr = tf.TextRange
    tr.Text = label
    tr.Font.Size = 9
    tr.Font.Name = t["body_font"]
    tr.Font.Bold = strong
    on_accent = "#FFFFFF" if luminance(t["accent"]) < 0.35 else "#111111"
    tr.Font.Fill.ForeColor.RGB = c(on_accent if strong else t["text"])
    tr.ParagraphFormat.Alignment = 2
    shp.Placement = 3
    ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress=f"'{target}'!A1", ScreenTip=f"Go to {target}")
    return shp


def icon(ws, t, x, y, glyph, size=24):
    shp = ws.Shapes.AddShape(9, x, y, size, size)
    shp.Fill.ForeColor.RGB = c(t["card2"])
    shp.Line.ForeColor.RGB = c(t["accent"])
    shp.Line.Weight = 1
    tf = shp.TextFrame2
    tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
    tf.VerticalAnchor = 3
    tf.WordWrap = False
    tr = tf.TextRange
    tr.Text = glyph
    tr.Font.Size = 11
    tr.Font.Bold = True
    tr.Font.Name = "Segoe UI Symbol"
    tr.Font.Fill.ForeColor.RGB = c(t["accent"])
    tr.ParagraphFormat.Alignment = 2
    shp.Placement = 3


def good_bad(ws, t, addr):
    cell = ws.Range(addr)
    a = cell.Address  # absolute, e.g. $C$9
    for ch, col in (("▲", "good"), ("▼", "bad")):
        fc = cell.FormatConditions.Add(Type=2, Formula1=f'=LEFT({a}{SEP}1)="{ch}"')
        fc.Font.Color = c(t[col])


def sign_colour(ws, t, addr):
    cell = ws.Range(addr)
    a = cell.Address
    fc = cell.FormatConditions.Add(Type=2, Formula1=f"={a}<0")
    fc.Font.Color = c(t["bad"])
    fc = cell.FormatConditions.Add(Type=2, Formula1=f"={a}>=0")
    fc.Font.Color = c(t["good"])


def style_chart(ch, t, title, sub=None):
    ch.ChartArea.Format.Fill.Visible = True
    ch.ChartArea.Format.Fill.Solid()
    ch.ChartArea.Format.Fill.ForeColor.RGB = c(t["card"])
    ch.ChartArea.Format.Line.Visible = True
    ch.ChartArea.Format.Line.ForeColor.RGB = c(t["line"])
    ch.ChartArea.Format.Line.Weight = 0.75
    ch.ChartArea.RoundedCorners = False
    ch.PlotArea.Format.Fill.Visible = False
    tr = ch.ChartArea.Format.TextFrame2.TextRange
    tr.Font.Fill.ForeColor.RGB = c(t["muted"])
    tr.Font.Name = t["body_font"]
    tr.Font.Size = 9
    ch.HasTitle = True
    ch.ChartTitle.Text = title + ("\n" + sub if sub else "")
    tr = ch.ChartTitle.Format.TextFrame2.TextRange
    tr.Font.Name = t["title_font"]
    tr.Font.Size = 13
    tr.Font.Bold = False
    tr.Font.Fill.ForeColor.RGB = c(t["text"])
    if sub:
        n = len(title) + 2
        part = tr.GetCharacters(n, len(sub))
        part.Font.Size = 9
        part.Font.Name = t["body_font"]
        part.Font.Fill.ForeColor.RGB = c(t["muted"])
    ch.ChartTitle.Format.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
    ch.ChartTitle.Left = 10
    ch.ChartTitle.Top = 6


def legend(ch, t, pos):
    ch.HasLegend = True
    ch.Legend.Position = pos
    ch.Legend.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = c(t["text"])
    ch.Legend.Format.TextFrame2.TextRange.Font.Size = 9


def build_theme(wb, calc, i):
    t = THEMES[i]
    n = len(THEMES)
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheet_name(i)
    ws.Tab.Color = c(t["accent"])
    ws.Cells.Font.Name = t["body_font"]
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 2.5
    ws.Columns("B:X").ColumnWidth = 7.6
    for g in ("G", "M", "S"):
        ws.Columns(g).ColumnWidth = 1.6
    ws.Columns("Y").ColumnWidth = 2.5
    ws.Range("A1:Y44").Interior.Color = c(t["page"])
    ws.Range("A1:Y44").Font.Color = c(t["text"])
    for r, h in {1: 8, 2: 15, 3: 32, 4: 18, 5: 2.25, 6: 12, 7: 20, 8: 36, 9: 16, 10: 16, 11: 8, 12: 12}.items():
        ws.Rows(r).RowHeight = h
    for r in range(13, 31):
        ws.Rows(r).RowHeight = 15
    for r, h in {31: 12, 32: 26, 33: 15, 34: 20, 35: 19, 36: 19, 37: 19, 38: 19, 39: 20, 40: 10, 41: 12, 42: 16,
                 43: 10}.items():
        ws.Rows(r).RowHeight = h

    # header band
    fill(ws.Range("A1:Y4"), t["header"])
    fill(ws.Range("A5:Y5"), t["accent"])
    ws.Range("B2").Value = "   ".join(" ".join(w.upper()) for w in f"Theme {i + 1}  of  {n}  ·  {t['name']}".split())
    font(ws.Range("B2"), t["accent"], 8, True)
    ws.Range("B3").Value = "Sales overview"
    font(ws.Range("B3"), t["htext"], 22, False, t["title_font"])
    ws.Range("B3").VerticalAlignment = -4108
    ws.Range("B4").Formula = "=Calc!B48"
    font(ws.Range("B4"), t["hmuted"], 9)
    right = ws.Range("X1").Left + ws.Range("X1").Width
    top = ws.Range("B3").Top + 5
    prev_i, next_i = (i - 1) % n, (i + 1) % n
    bw, gap = 84, 6
    x = right - 3 * bw - 2 * gap
    button(ws, t, x, top, bw, 22, "◀  Previous", sheet_name(prev_i))
    button(ws, t, x + bw + gap, top, bw, 22, "All themes", "Themes", strong=True)
    button(ws, t, x + 2 * (bw + gap), top, bw, 22, "Next  ▶", sheet_name(next_i))

    # KPI cards
    cards = [("REVENUE", "=Calc!B32", "#,##0", "=Calc!C34", "=Calc!C36", "$"),
             ("GROSS MARGIN", "=Calc!B38", "0.0%", "=Calc!C40", "=Calc!C37", "%"),
             ("ORDERS", "=Calc!B41", "#,##0", "=Calc!C43", '="across 4 regions"', "#"),
             ("AVERAGE ORDER", "=Calc!B44", "#,##0", "=Calc!C46", '="revenue / orders"', "≈")]
    for (c1, c2), (label, big, fmt, note1, note2, glyph) in zip(CARD_COLS, cards):
        blk = ws.Range(f"{c1}7:{c2}11")
        fill(blk, t["card"])
        box(blk, t["line"])
        b = blk.Borders(7)
        b.Color = c(t["accent"])
        b.Weight = 4
        ws.Range(f"{c1}7").Value = label
        font(ws.Range(f"{c1}7"), t["muted"], 8, True)
        ws.Range(f"{c1}7").IndentLevel = 1
        ws.Range(f"{c1}8:{c2}8").Merge()
        ws.Range(f"{c1}8").Formula = big
        ws.Range(f"{c1}8").NumberFormat = fmt
        ws.Range(f"{c1}8").HorizontalAlignment = -4131
        font(ws.Range(f"{c1}8"), t["text"], 24, False, t["title_font"])
        ws.Range(f"{c1}8").IndentLevel = 1
        for r, f in ((9, note1), (10, note2)):
            ws.Range(f"{c1}{r}").Formula = f
            font(ws.Range(f"{c1}{r}"), t["muted"], 9)
            ws.Range(f"{c1}{r}").IndentLevel = 1
            good_bad(ws, t, f"{c1}{r}")
        last = ws.Range(f"{c2}7")
        icon(ws, t, last.Left + last.Width - 30, ws.Range(f"{c1}7").Top + 6, glyph)

    # chart 1: revenue vs budget and prior year
    left, top = ws.Range("B13").Left, ws.Range("B13").Top
    width = ws.Range("L13").Left + ws.Range("L13").Width - left
    height = ws.Range("B31").Top - top
    ch = ws.Shapes.AddChart2(-1, 51, left, top, width, height).Chart
    while ch.SeriesCollection().Count:
        ch.SeriesCollection(1).Delete()
    s = ch.SeriesCollection().NewSeries()
    s.Name = "Actual"
    s.Values = calc.Range("B4:B15")
    s.XValues = calc.Range("E4:E15")
    s.Format.Fill.ForeColor.RGB = c(t["accent"])
    s.Format.Line.Visible = False
    s.Points(12).Format.Fill.ForeColor.RGB = c(t["accent2"])
    ch.ChartGroups(1).GapWidth = 55
    for name, rng, col, dash in (("Budget", "C4:C15", t["muted"], 1), ("Prior year", "D4:D15", t["s2"], 4)):
        s = ch.SeriesCollection().NewSeries()
        s.Name = name
        s.Values = calc.Range(rng)
        s.XValues = calc.Range("E4:E15")
        s.ChartType = 4
        s.Format.Line.ForeColor.RGB = c(col)
        s.Format.Line.Weight = 2
        s.Format.Line.DashStyle = dash
        s.MarkerStyle = -4142
    style_chart(ch, t, "Revenue vs budget and prior year", "12 months to " + REPORT.strftime("%b %Y") +
                "  |  bars = actual, light bar = reporting month")
    legend(ch, t, -4160)
    ax = ch.Axes(2)
    ax.TickLabels.NumberFormat = '#,##0,"k"'
    ax.HasMajorGridlines = True
    ax.MajorGridlines.Format.Line.ForeColor.RGB = c(t["line"])
    ax.MajorGridlines.Format.Line.Weight = 0.5
    ax.Format.Line.Visible = False
    ax.TickLabels.Font.Color = c(t["muted"])
    ax.MinimumScale = 0
    cx = ch.Axes(1)
    cx.Format.Line.ForeColor.RGB = c(t["line"])
    cx.TickLabels.Font.Color = c(t["muted"])

    # chart 2: category doughnut
    left = ws.Range("N13").Left
    width = ws.Range("X13").Left + ws.Range("X13").Width - left
    shp = ws.Shapes.AddChart2(-1, -4120, left, top, width, height)
    ch = shp.Chart
    while ch.SeriesCollection().Count:
        ch.SeriesCollection(1).Delete()
    s = ch.SeriesCollection().NewSeries()
    s.Name = "Revenue"
    s.Values = calc.Range("B18:B22")
    s.XValues = calc.Range("D18:D22")
    for k, role in enumerate(("accent", "s2", "s3", "s4", "accent2"), 1):
        p = s.Points(k)
        p.Format.Fill.ForeColor.RGB = c(t[role])
        p.Format.Line.ForeColor.RGB = c(t["card"])
        p.Format.Line.Weight = 2
    ch.ChartGroups(1).DoughnutHoleSize = 62
    style_chart(ch, t, "Revenue by category", "share of " + REPORT.strftime("%B %Y") + " revenue")
    legend(ch, t, -4152)
    ch.Legend.Format.TextFrame2.TextRange.Font.Size = 10
    pa = ch.PlotArea
    side = min(pa.InsideWidth, pa.InsideHeight)
    cxp = shp.Left + pa.InsideLeft + pa.InsideWidth / 2
    cyp = shp.Top + pa.InsideTop + pa.InsideHeight / 2
    for txt, size, col, dy, fnt in (("=Calc!$C$32", 18, t["text"], -16, t["title_font"]),
                                     ("=Calc!$C$33", 8, t["muted"], 8, t["body_font"])):
        tb = ws.Shapes.AddTextbox(1, cxp - side * 0.3, cyp + dy, side * 0.6, 22)
        tb.Fill.Visible = False
        tb.Line.Visible = False
        tb.DrawingObject.Formula = txt
        tf = tb.TextFrame2
        tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
        tf.WordWrap = False
        tr = tf.TextRange
        tr.Font.Size = size
        tr.Font.Name = fnt
        tr.Font.Fill.ForeColor.RGB = c(col)
        tr.ParagraphFormat.Alignment = 2
        tb.Placement = 3

    # regions table
    blk = ws.Range("B32:L40")
    fill(blk, t["card"])
    box(blk, t["line"])
    ws.Range("B32").Value = "Regions compared"
    font(ws.Range("B32"), t["text"], 13, False, t["title_font"])
    ws.Range("B33").Formula = '=TEXT(Calc!$C$1,"mmmm yyyy")&" vs budget  |  red = below budget, green = above"'
    font(ws.Range("B33"), t["muted"], 9)
    heads = {"B": "Region", "D": "Revenue", "E": "Budget", "F": "Variance", "H": "Var %", "I": "Margin", "J": "Share"}
    for col, h in heads.items():
        cell = ws.Range(f"{col}34")
        cell.Value = h
        cell.HorizontalAlignment = -4131 if col == "B" else -4152
    font(ws.Range("B34:L34"), t["accent"], 9, True)
    ws.Range("B34:L34").Borders(9).LineStyle = 1
    ws.Range("B34:L34").Borders(9).Color = c(t["accent"])
    for k in range(5):
        r, cr = 35 + k, 25 + k
        ws.Range(f"B{r}").Formula = f"=Calc!A{cr}"
        for col, src in (("D", "B"), ("E", "C"), ("F", "D"), ("H", "E"), ("I", "F"), ("J", "G")):
            ws.Range(f"{col}{r}").Formula = f"=Calc!{src}{cr}"
        ws.Range(f"D{r}:E{r}").NumberFormat = "#,##0"
        ws.Range(f"F{r}").NumberFormat = "+#,##0;-#,##0;0"
        ws.Range(f"H{r}").NumberFormat = "+0.0%;-0.0%;0.0%"
        ws.Range(f"I{r}").NumberFormat = "0.0%"
        ws.Range(f"J{r}").NumberFormat = "0.0%"
        ws.Range(f"K{r}:L{r}").Merge()
        if k < 4:
            ws.Range(f"K{r}").Formula = f"=J{r}"
            ws.Range(f"K{r}").NumberFormat = ";;;"
        font(ws.Range(f"B{r}:L{r}"), t["text"], 10, k == 4)
        if k % 2 == 1:
            fill(ws.Range(f"B{r}:L{r}"), t["card2"])
        sign_colour(ws, t, f"F{r}")
        sign_colour(ws, t, f"H{r}")
    ws.Range("B39:L39").Borders(8).LineStyle = 1
    ws.Range("B39:L39").Borders(8).Color = c(t["line"])
    db = ws.Range("K35:K38").FormatConditions.AddDatabar()
    db.BarFillType = 0
    db.BarColor.Color = c(t["accent"])
    db.MinPoint.Modify(0, 0)
    db.MaxPoint.Modify(0, 0.4)
    ws.Range("D35:J39").HorizontalAlignment = -4152

    # palette panel
    blk = ws.Range("N32:X40")
    fill(blk, t["card"])
    box(blk, t["line"])
    ws.Range("N32").Value = "This theme's colours"
    font(ws.Range("N32"), t["text"], 13, False, t["title_font"])
    ws.Range("N33").Value = "Hex codes, ready for Excel, Power BI, Google Sheets or Looker Studio"
    font(ws.Range("N33"), t["muted"], 9)
    labels = {"page": "Page", "header": "Header", "card": "Card", "accent": "Accent", "accent2": "Light",
              "text": "Text", "muted": "Muted", "s2": "Series 2", "s3": "Series 3", "s4": "Series 4"}
    for col, role in zip(SWATCH_COLS, SWATCHES):
        sw = ws.Range(f"{col}35:{col}36")
        fill(sw, t[role])
        box(sw, t["line"], 2)
        ws.Range(f"{col}37").Value = labels[role]
        font(ws.Range(f"{col}37"), t["muted"], 8)
        ws.Range(f"{col}38").Value = t[role].upper()
        font(ws.Range(f"{col}38"), t["text"], 8)
        ws.Range(f"{col}37:{col}38").HorizontalAlignment = -4108
    fname = t["title_font"]
    ws.Range("N39").Value = f"Fonts: {fname} for titles, {t['body_font']} for text  ·  {t['mood']}  ·  best for {t['best_for'].lower()}"
    font(ws.Range("N39"), t["muted"], 8, italic=True)

    # footer
    ws.Range("B42").Value = f"Theme {i + 1} of {n}: {t['name']}  ·  same data and formulas on every theme  ·  fictional data  ·  no macros"
    font(ws.Range("B42"), t["muted"], 9)
    ws.Range("X42").Value = "Pick a number, it is applied to your workbook"
    ws.Range("X42").HorizontalAlignment = -4152
    font(ws.Range("X42"), t["accent"], 9, True)
    page_setup(ws, "A1:Y43")
    return ws


# ---------------------------------------------------------------- index, palette, how to use
NEUTRAL = dict(page="#ECEAE6", header="#15181C", ink="#1E2228", muted="#6B6F76", gold="#C9A35B", ivory="#F2EEE6")


def build_index(wb):
    ws = wb.Worksheets(1)
    ws.Name = "Themes"
    ws.Tab.Color = c(NEUTRAL["gold"])
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 2.5
    ws.Columns("B:X").ColumnWidth = 7.6
    for g in ("G", "M", "S"):
        ws.Columns(g).ColumnWidth = 1.6
    ws.Columns("Y").ColumnWidth = 2.5
    fill(ws.Range("A1:Y40"), NEUTRAL["page"])
    fill(ws.Range("A1:Y4"), NEUTRAL["header"])
    fill(ws.Range("A5:Y5"), NEUTRAL["gold"])
    for r, h in {1: 8, 2: 15, 3: 34, 4: 18, 5: 2.25, 6: 14}.items():
        ws.Rows(r).RowHeight = h
    ws.Range("B2").Value = "C O L O U R   T H E M E S   ·   S A L E S   D A S H B O A R D"
    font(ws.Range("B2"), NEUTRAL["gold"], 8, True)
    ws.Range("B3").Value = "Choose a colour theme"
    font(ws.Range("B3"), NEUTRAL["ivory"], 24, False, "Georgia")
    ws.Range("B3").VerticalAlignment = -4108
    ws.Range("B4").Value = ("The same dashboard, data and formulas in 8 themes. Click a tile to open it, then tell me "
                            "the number, or send your own brand colours.")
    font(ws.Range("B4"), "#A9AEB5", 9)
    rows = [7, 14]
    heights = [30, 32, 16, 20, 26, 22, 16]
    for i, t in enumerate(THEMES):
        c1, c2 = CARD_COLS[i % 4]
        r0 = rows[i // 4]
        for k, h in enumerate(heights):
            ws.Rows(r0 + k).RowHeight = h
        cols = [chr(x) for x in range(ord(c1), ord(c2) + 1)]
        band = ws.Range(f"{c1}{r0}:{c2}{r0}")
        fill(band, t["header"])
        ws.Range(f"{c1}{r0}").Value = f"{i + 1}  {t['name']}"
        font(ws.Range(f"{c1}{r0}"), t["htext"], 13, False, t["title_font"])
        ws.Range(f"{c1}{r0}").IndentLevel = 1
        ws.Range(f"{c1}{r0}").VerticalAlignment = -4108
        fill(ws.Range(f"{c1}{r0 + 1}:{c2}{r0 + 2}"), t["card"])
        cols_mid = cols[len(cols) - 3]
        ws.Range(f"{c1}{r0 + 1}:{cols_mid}{r0 + 1}").Merge()
        ws.Range(f"{cols[-2]}{r0 + 1}:{c2}{r0 + 1}").Merge()
        ws.Range(f"{c1}{r0 + 1}").Formula = "=Calc!B32"
        ws.Range(f"{c1}{r0 + 1}").NumberFormat = "#,##0"
        ws.Range(f"{c1}{r0 + 1}").HorizontalAlignment = -4131
        ws.Range(f"{c1}{r0 + 1}").IndentLevel = 1
        font(ws.Range(f"{c1}{r0 + 1}"), t["text"], 18, False, t["title_font"])
        nb = f"{cols[-2]}{r0 + 1}"
        ws.Range(nb).Formula = "=Calc!C34"
        ws.Range(nb).HorizontalAlignment = -4152
        ws.Range(nb).IndentLevel = 1
        font(ws.Range(nb), t["muted"], 8)
        good_bad(ws, t, nb)
        ws.Range(f"{c1}{r0 + 2}").Value = f"{t['mood']}  ·  revenue, Oct 2026"
        font(ws.Range(f"{c1}{r0 + 2}"), t["muted"], 8)
        ws.Range(f"{c1}{r0 + 2}").IndentLevel = 1
        for col, role in zip(cols, ("accent", "s2", "s3", "s4", "accent2")):
            fill(ws.Range(f"{col}{r0 + 3}"), t[role])
        ws.Range(f"{c1}{r0 + 3}:{c2}{r0 + 3}").Borders(9).LineStyle = 1
        ws.Range(f"{c1}{r0 + 3}:{c2}{r0 + 3}").Borders(9).Color = c(t["card"])
        ws.Range(f"{c1}{r0 + 3}:{c2}{r0 + 3}").Borders(9).Weight = 4
        fill(ws.Range(f"{c1}{r0 + 4}:{c2}{r0 + 5}"), t["page"])
        ws.Range(f"{c1}{r0 + 4}:{c2}{r0 + 4}").Merge()
        ws.Range(f"{c1}{r0 + 4}").Value = "Best for: " + t["best_for"].lower()
        ws.Range(f"{c1}{r0 + 4}").WrapText = True
        ws.Range(f"{c1}{r0 + 4}").VerticalAlignment = -4108
        ws.Range(f"{c1}{r0 + 4}").IndentLevel = 1
        font(ws.Range(f"{c1}{r0 + 4}"), t["muted"], 8)
        link = ws.Range(f"{c1}{r0 + 5}")
        ws.Hyperlinks.Add(Anchor=link, Address="", SubAddress=f"'{sheet_name(i)}'!A1",
                          ScreenTip=f"Open theme {i + 1}", TextToDisplay=f"Open theme {i + 1}  ▸")
        font(link, t["accent"], 9, True)
        link.Font.Underline = -4142
        link.IndentLevel = 1
        box(ws.Range(f"{c1}{r0}:{c2}{r0 + 5}"), t["line"] if is_dark(t) else "#C9CED6")
    r = 22
    notes = ["Every theme uses the same data, formulas, charts and layout, so you compare colours only.",
             "All 8 pass WCAG AA contrast: text on cards at least 4.5:1, secondary text, accents and red/green at least 3:1.",
             "Hex codes for every colour are on each theme page and on the Palette sheet (Excel, Power BI, Google Sheets).",
             "Your own brand colours: send the hex codes or your logo and they become theme 9 on your workbook."]
    for k, note in enumerate(notes):
        cell = ws.Range(f"B{r + k}")
        cell.Value = "✓  " + note
        font(cell, NEUTRAL["ink"], 10)
        ws.Rows(r + k).RowHeight = 18
    ws.Range(f"B{r + 5}").Value = f"{COMPANY}  ·  fictional data  ·  no macros"
    font(ws.Range(f"B{r + 5}"), NEUTRAL["muted"], 8)
    page_setup(ws, f"A1:Y{r + 6}")
    return ws


def build_palette(wb):
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Palette"
    ws.Tab.Color = c(NEUTRAL["gold"])
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    fill(ws.Range("A1:J30"), NEUTRAL["page"])
    ws.Columns("A").ColumnWidth = 2.5
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C:J").ColumnWidth = 17
    fill(ws.Range("A1:J3"), NEUTRAL["header"])
    fill(ws.Range("A4:J4"), NEUTRAL["gold"])
    ws.Rows(1).RowHeight = 8
    ws.Rows(2).RowHeight = 34
    ws.Rows(3).RowHeight = 18
    ws.Rows(4).RowHeight = 2.25
    ws.Rows(5).RowHeight = 12
    ws.Range("B2").Value = "Palette: every colour of every theme"
    font(ws.Range("B2"), NEUTRAL["ivory"], 20, False, "Georgia")
    ws.Range("B2").VerticalAlignment = -4108
    ws.Range("B3").Value = "Copy a hex code into Excel (More Colors > Hex), Power BI (theme JSON) or Google Sheets (custom colour)."
    font(ws.Range("B3"), "#A9AEB5", 9)
    ws.Range("B6").Value = "Role"
    font(ws.Range("B6"), NEUTRAL["ink"], 10, True)
    for k, t in enumerate(THEMES):
        cell = ws.Cells(6, 3 + k)
        cell.Value = f"{k + 1} {t['name']}"
        fill(cell, t["header"])
        font(cell, t["htext"], 9, True)
        cell.WrapText = True
        cell.HorizontalAlignment = -4108
        cell.VerticalAlignment = -4108
        ws.Hyperlinks.Add(Anchor=cell, Address="", SubAddress=f"'{sheet_name(k)}'!A1", ScreenTip="Open this theme")
        font(cell, t["htext"], 9, True)
        cell.Font.Underline = -4142
    ws.Rows(6).RowHeight = 32
    for j, (role, label) in enumerate(ROLES):
        r = 7 + j
        ws.Rows(r).RowHeight = 20
        ws.Cells(r, 2).Value = label
        font(ws.Cells(r, 2), NEUTRAL["ink"], 10)
        ws.Cells(r, 2).IndentLevel = 1
        for k, t in enumerate(THEMES):
            cell = ws.Cells(r, 3 + k)
            cell.Value = t[role].upper()
            fill(cell, t[role])
            font(cell, "#FFFFFF" if luminance(t[role]) < 0.3 else "#111111", 9)
            cell.HorizontalAlignment = -4108
    r = 7 + len(ROLES)
    for j, (key, label) in enumerate((("title_font", "Title font"), ("body_font", "Text font"), ("mood", "Mood"),
                                      ("best_for", "Best for"))):
        ws.Cells(r + j, 2).Value = label
        font(ws.Cells(r + j, 2), NEUTRAL["ink"], 10)
        ws.Rows(r + j).RowHeight = 30 if key == "best_for" else 20
        for k, t in enumerate(THEMES):
            cell = ws.Cells(r + j, 3 + k)
            cell.Value = t[key]
            font(cell, NEUTRAL["ink"], 9, name=t[key] if key == "title_font" else None)
            cell.WrapText = True
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108
    last = r + 3
    rng = ws.Range(ws.Cells(6, 2), ws.Cells(last, 10))
    for edge in (7, 8, 9, 10, 11, 12):
        rng.Borders(edge).LineStyle = 1
        rng.Borders(edge).Color = c("#D2CEC6")
    ws.Range(ws.Cells(6, 2), ws.Cells(last, 10)).VerticalAlignment = -4108
    ws.Range(ws.Cells(6, 2), ws.Cells(last, 2)).IndentLevel = 1
    page_setup(ws, f"A1:J{last + 1}")
    return ws


def build_howto(wb):
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "How to use"
    ws.Tab.Color = c(NEUTRAL["gold"])
    ws.Cells.Font.Name = "Segoe UI"
    ws.Columns("A").ColumnWidth = 2.5
    ws.Columns("B").ColumnWidth = 120
    lines = [
        ("How to use this gallery", "title"),
        ("1.  Open the Themes sheet and click a tile, or use Previous / Next at the top of each theme page.", ""),
        ("2.  Every theme shows the same October 2026 sales dashboard: 4 KPI cards, a 12-month chart, a category "
         "doughnut, a region table and the theme's own colour codes.", ""),
        ("3.  Pick the one you like and reply with its number (for example \"theme 3\"). Mixing is fine too: "
         "\"theme 6 with the accent of theme 1\".", ""),
        ("4.  Prefer your own brand? Send the hex codes or your logo; they become theme 9 with the same contrast checks.", ""),
        ("", ""),
        ("Behind the pages", "title"),
        ("Calc holds every number as a plain SUMIFS on the Sales table (sheet Data), so all 8 themes always agree.", ""),
        ("Red / green notes and variances are conditional formatting: they turn the right colour when the numbers change.", ""),
        ("Palette lists all 14 colour roles of every theme plus the font pairing.", ""),
        ("No macros, no add-ins: the file opens in any Excel 2016 or newer, and in Excel for the web.", ""),
        ("", ""),
        (f"{COMPANY}: all names and amounts are made up.", "muted"),
    ]
    for k, (txt, kind) in enumerate(lines, 2):
        cell = ws.Range(f"B{k}")
        cell.Value = txt
        cell.WrapText = True
        if kind == "title":
            font(cell, NEUTRAL["ink"], 15, False, "Georgia")
            ws.Rows(k).RowHeight = 26
        else:
            font(cell, NEUTRAL["muted"] if kind == "muted" else NEUTRAL["ink"], 10)
    page_setup(ws, f"A1:B{len(lines) + 2}")
    return ws


def main():
    if OUT.exists():
        OUT.unlink()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    global SEP
    SEP = xl.International[4]
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count > 1:
            wb.Worksheets(wb.Worksheets.Count).Delete()
        # data and calc first exist at the end of the tab order; theme sheets are inserted before them
        data = build_data(wb)
        calc = build_calc(wb)
        xl.Calculate()
        build_index(wb)
        for i in range(len(THEMES)):
            ws = build_theme(wb, calc, i)
            ws.Move(Before=wb.Worksheets("Data"))
        build_palette(wb).Move(Before=wb.Worksheets("Data"))
        build_howto(wb).Move(Before=wb.Worksheets("Data"))
        calc.Move(Before=wb.Worksheets("Data"))
        xl.ScreenUpdating = True
        xl.Calculate()
        for ws in wb.Worksheets:
            window(xl, ws, headings=ws.Name in ("Calc", "Data"))
        wb.Worksheets("Themes").Activate()
        wb.SaveAs(str(OUT), 51)
        wb.Close(False)
        print(OUT)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()
