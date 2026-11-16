"""Builds Sales-Dashboard-Redesign.xlsx with real Excel (COM): the client's dashboard redone in a light brand theme,
on cleaned data, with every number error of the old file listed and a bridge from the old total to the right one.
No macros, Excel 2016 functions.

Sheets
  Dashboard     KPI cards vs last year, insights written by formulas, months 2026 vs 2025, regions, categories, top products
  Issues found  the 9 mistakes in the old file: where, what, old value (read from the old file), correct, difference, fix
  Bridge        old total -> + rows outside the SUM range + text numbers - duplicates = correct total (waterfall)
  Data          both exports as received, plus helper columns: number, trimmed name, duplicate, in the old range, note
  Checks        every view adds up to the same total, the bridge closes, the planted problems are all found
  Before        a picture of the old dashboard, for comparison
  How to use
  Calc          (hidden) tables behind the charts

    python scripts/build_after.py         (needs Before-Client-Dashboard.xlsx and docs/before.png; run under tools/with_lock.py)
"""
import csv
from datetime import date
from pathlib import Path
import win32com.client as win32
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Sales-Dashboard-Redesign.xlsx"
CATS = ["Seeds & bulbs", "Plants", "Hand tools", "Power tools", "Watering", "Soil & feed"]
PRODUCTS = [r[1] for r in list(csv.reader(open(ROOT / "data" / "products.csv", encoding="utf-8")))[1:]]
REGIONS = ["North", "South", "East", "West"]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct"]


def hexrgb(h):
    h = h.lstrip("#"); r, g, b = int(h[:2], 16), int(h[2:4], 16), int(h[4:], 16)
    return r + g * 256 + b * 65536


# ---- light brand theme: cream page, forest green brand, terracotta second series
PAGE, CARD, CARD2, LINE = "#F7F4EC", "#FFFFFF", "#FBF9F3", "#E3DDCF"
ACCENT, ACCENT2, S2, S3 = "#2F6B4F", "#8DBF9F", "#C8793F", "#D9B44A"
TEXT, MUTED, GOOD, BAD, LY = "#23302A", "#6B756E", "#2F8A4E", "#B83A2E", "#C9C1B0"
GEORGIA, BODY = "Georgia", "Segoe UI"
BRAND = "   ".join(" ".join(w) for w in "GREENWAY GARDEN SUPPLY".split())

# ---- the old file, as Excel saved it
old = load_workbook(ROOT / "Before-Client-Dashboard.xlsx", data_only=True)["Dashboard"]
OLD = {k: old[k].value for k in ("B3", "B4", "B5", "B6", "B7", "E7", "B20", "E18")}
OLD["hose"] = next(old[f"H{r}"].value for r in range(4, 22) if old[f"G{r}"].value == "Hose 25m")
OLD["hose_cell"] = next(f"H{r}" for r in range(4, 22) if old[f"G{r}"].value == "Hose 25m")


import json as _json
TEXT_IDS = set(_json.loads((ROOT / "data" / "planted.json").read_text())["text_revenue_orders"])


def load(name):
    out = []
    for r in list(csv.reader(open(ROOT / "data" / name, encoding="utf-8")))[1:]:
        rev = r[7] if (r[0] in TEXT_IDS and "2026" in name) else float(r[7])
        out.append([r[0], (date.fromisoformat(r[1]) - date(1899, 12, 30)).days, r[2], r[3], r[4], int(r[5]), float(r[6]), rev, float(r[8])])
    return out


ROWS = load("sales-2025.csv") + load("sales-2026.csv")
F0, F1 = 6, 6 + len(ROWS) - 1                      # data rows on the Data sheet
rng = lambda c: f"Data!${c}${F0}:${c}${F1}"
END26, END25 = "DATE(2026,10,31)", "DATE(2025,10,31)"


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


def page(ws, title, sub):
    ws.Cells.Font.Name = BODY; ws.Cells.Font.Size = 10; ws.Cells.Font.Color = hexrgb(TEXT)
    ws.Cells.Interior.Color = hexrgb(PAGE)
    style(ws.Range("B2"), fg=ACCENT, size=8, bold=True).Value = BRAND
    style(ws.Range("B3"), size=22, name=GEORGIA).Value = title
    style(ws.Range("B4"), fg=MUTED, size=9).Value = sub
    ws.Rows(3).RowHeight = 32
    ws.Columns("A").ColumnWidth = 2


def chart_style(ch):
    ch.ChartArea.Format.Fill.ForeColor.RGB = hexrgb(CARD); ch.ChartArea.Format.Line.ForeColor.RGB = hexrgb(LINE)
    ch.PlotArea.Format.Fill.Visible = False
    tr = ch.ChartArea.Format.TextFrame2.TextRange.Font; tr.Fill.ForeColor.RGB = hexrgb(MUTED); tr.Size = 9; tr.Name = BODY


xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False; xl.ScreenUpdating = False
if OUT.exists():
    OUT.unlink()
try:
    wb = xl.Workbooks.Add()
    SHEETS = ["Dashboard", "Issues found", "Bridge", "Data", "Checks", "Before", "How to use", "Calc"]
    while wb.Worksheets.Count < len(SHEETS):
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    for i, n in enumerate(SHEETS, 1):
        wb.Worksheets(i).Name = n
    W = {n: wb.Worksheets(n) for n in SHEETS}

    # ------------------------------------------------------------------ Data
    ws = W["Data"]
    page(ws, "Data: both exports as received", "Columns A-I exactly as exported (text numbers, trailing spaces and duplicates included). Columns J-P clean them with formulas; nothing is typed over.")
    hdr = ["Order ID", "Date", "Region", "Category", "Product", "Units", "Unit price", "Revenue (as exported)", "Cost",
           "Year", "Revenue", "Product (clean)", "Duplicate?", "Counted", "In the old SUM range", "Note"]
    for i, h in enumerate(hdr, 1):
        ws.Cells(5, i).Value = h
    style(ws.Range("A5:P5"), fg="#FFFFFF", bg=ACCENT, bold=True, size=9, wrap=True, v="c")
    style(ws.Range("J5:P5"), bg="#4F8A6C")
    ws.Rows(5).RowHeight = 30
    ws.Range(f"H{F0}:H{F1}").NumberFormat = "General"
    ws.Range(f"A{F0}:I{F1}").Value = ROWS
    for i, r in enumerate(ROWS, F0):
        if isinstance(r[7], str):
            ws.Cells(i, 8).NumberFormat = "@"; ws.Cells(i, 8).Value = r[7]
    ws.Range(f"B{F0}:B{F1}").NumberFormat = "d mmm yyyy"
    f = {"J": "=YEAR(B{r})", "K": '=IF(ISTEXT(H{r}),NUMBERVALUE(H{r},".",","),H{r})', "L": "=TRIM(E{r})",
         "M": '=IF(COUNTIF(A$6:A{r},A{r})>1,"Duplicate","")', "N": '=IF(M{r}="",1,0)',
         "O": '=IF(J{r}=2026,IF(COUNTIF(J$6:J{r},2026)<=699,"Yes","No"),"")',
         "P": '=MID(IF(ISTEXT(H{r}),"; number came as text","")&IF(E{r}<>L{r},"; extra space in the name","")&IF(M{r}<>"","; exported twice: not counted",""),3,200)'}
    for col, ff in f.items():
        ws.Range(f"{col}{F0}:{col}{F1}").Formula = ff.format(r=F0)
    ws.Range(f"K{F0}:K{F1}").NumberFormat = "#,##0.00"; ws.Range(f"G{F0}:G{F1}").NumberFormat = "#,##0.00"
    ws.Range(f"I{F0}:I{F1}").NumberFormat = "#,##0.00"
    for col, w in zip("ABCDEFGHIJKLMNOP", (10, 11, 8, 13, 17, 6, 9, 12, 9, 6, 10, 17, 10, 8, 10, 40)):
        ws.Columns(col).ColumnWidth = w
    fc = ws.Range(f"P{F0}:P{F1}").FormatConditions.Add(1, 4, '=""'); fc.Font.Color = hexrgb(BAD)
    ws.Range(f"A{F0}:P{F1}").Interior.Color = hexrgb(CARD)
    ws.Activate(); ws.Range("B6").Select(); xl.ActiveWindow.FreezePanes = True; xl.ActiveWindow.DisplayGridlines = False
    ws.Range(f"A5:P{F1}").AutoFilter(1)

    # ------------------------------------------------------------------ Calc (hidden)
    ws = W["Calc"]
    S26 = lambda extra="": f'SUMIFS({rng("K")},{rng("J")},2026,{rng("N")},1,{rng("B")},"<="&{END26}{extra})'
    S25 = lambda extra="": f'SUMIFS({rng("K")},{rng("J")},2025,{rng("N")},1,{rng("B")},"<="&{END25}{extra})'
    C26 = lambda extra="": f'SUMIFS({rng("I")},{rng("J")},2026,{rng("N")},1,{rng("B")},"<="&{END26}{extra})'
    C25 = lambda extra="": f'SUMIFS({rng("I")},{rng("J")},2025,{rng("N")},1,{rng("B")},"<="&{END25}{extra})'
    ws.Range("A1").Value = "Tables behind the charts and cards. Hidden; nothing to type."
    kp = [("Rev_26", "=" + S26()), ("Rev_25", "=" + S25()), ("Cost_26", "=" + C26()), ("Cost_25", "=" + C25()),
          ("Orders_26", f'=COUNTIFS({rng("J")},2026,{rng("N")},1)'), ("Orders_25", f'=COUNTIFS({rng("J")},2025,{rng("N")},1,{rng("B")},"<="&{END25})'),
          ("Growth", "=B3/B4-1"), ("Margin_26", "=(B3-B5)/B3"), ("Margin_25", "=(B4-B6)/B4"), ("AOV_26", "=B3/B7"), ("AOV_25", "=B4/B8")]
    for i, (n, ff) in enumerate(kp, 3):
        ws.Cells(i, 1).Value = n; ws.Cells(i, 2).Formula = ff
        wb.Names.Add(n, f"=Calc!$B${i}")
    # months
    ws.Range("D2:F2").Value = ("Month", "2026", "2025")
    for m, nm in enumerate(MONTHS, 1):
        r = 2 + m
        ws.Cells(r, 4).Value = nm
        ws.Cells(r, 5).Formula = f'=SUMIFS({rng("K")},{rng("J")},2026,{rng("N")},1,{rng("B")},">="&DATE(2026,{m},1),{rng("B")},"<"&DATE(2026,{m + 1},1))'
        ws.Cells(r, 6).Formula = f'=SUMIFS({rng("K")},{rng("J")},2025,{rng("N")},1,{rng("B")},">="&DATE(2025,{m},1),{rng("B")},"<"&DATE(2025,{m + 1},1))'
    # regions
    ws.Range("H2:K2").Value = ("Region", "2026", "2025", "Growth")
    for i, reg in enumerate(REGIONS, 3):
        ws.Cells(i, 8).Value = reg
        ws.Cells(i, 9).Formula = "=" + S26(f',{rng("C")},H{i}'); ws.Cells(i, 10).Formula = "=" + S25(f',{rng("C")},H{i}')
        ws.Cells(i, 11).Formula = f"=I{i}/J{i}-1"
    # categories, then sorted largest first (for the bar chart)
    ws.Range("M2:N2").Value = ("Category", "2026")
    for i, c in enumerate(CATS, 3):
        ws.Cells(i, 13).Value = c; ws.Cells(i, 14).Formula = "=" + S26(f',{rng("D")},M{i}')
    ws.Range("P2:Q2").Value = ("Category (sorted)", "2026")
    for k in range(1, 7):
        r = 2 + k
        ws.Cells(r, 17).Formula = f"=LARGE($N$3:$N$8,{7 - k})"           # smallest first: a bar chart draws bottom-up
        ws.Cells(r, 16).Formula = f"=INDEX($M$3:$M$8,MATCH(Q{r},$N$3:$N$8,0))"
    # products (clean name), top 8 sorted
    ws.Range("S2:T2").Value = ("Product", "2026")
    for i, p in enumerate(PRODUCTS, 3):
        ws.Cells(i, 19).Value = p; ws.Cells(i, 20).Formula = "=" + S26(f',{rng("L")},S{i}')
    ws.Range("V2:W2").Value = ("Top product", "2026")
    for k in range(1, 9):
        r = 2 + k
        ws.Cells(r, 23).Formula = f"=LARGE($T$3:$T$20,{9 - k})"
        ws.Cells(r, 22).Formula = f"=INDEX($S$3:$S$20,MATCH(W{r},$T$3:$T$20,0))"
    ws.Visible = 0

    # ------------------------------------------------------------------ Dashboard
    ws = W["Dashboard"]
    page(ws, "Sales 2026, January to October", "")
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = False; xl.ActiveWindow.DisplayHeadings = False
    for c in range(2, 21):
        ws.Columns(c).ColumnWidth = 2.2 if c in (5, 9, 13, 17) else 10
    ws.Range("B4").Formula = (f'="Compared with the same months of 2025   ·   "&TEXT(COUNTA({rng("A")}),"#,##0")&" order lines checked: "'
                              f'&COUNTIF({rng("M")},"Duplicate")&" duplicates removed, "&SUMPRODUCT(--ISTEXT({rng("H")}))&" text numbers fixed, "'
                              f'&SUMPRODUCT(--({rng("E")}<>{rng("L")}))&" product names trimmed"')
    cards = [("REVENUE", "=Rev_26", "[$$-409]#,##0", '="vs "&TEXT(Rev_25,"[$$-409]#,##0")&" in 2025"'),
             ("GROWTH VS 2025", "=Growth", "+0.0%;-0.0%", '="same months, January to October"'),
             ("GROSS MARGIN", "=Margin_26", "0.0%", '=TEXT((Margin_26-Margin_25)*100,"+0.0;-0.0")&" points vs 2025"'),
             ("ORDERS", "=Orders_26", "#,##0", '=TEXT(Orders_26/Orders_25-1,"+0%;-0%")&" vs "&TEXT(Orders_25,"#,##0")&" in 2025"'),
             ("AVERAGE ORDER", "=AOV_26", "[$$-409]#,##0.00", '="vs "&TEXT(AOV_25,"[$$-409]#,##0.00")&" in 2025"')]
    for i, (lab, ff, fmt, note) in enumerate(cards):
        c0 = 2 + i * 4
        box = ws.Range(ws.Cells(6, c0), ws.Cells(10, c0 + 2))
        box.Interior.Color = hexrgb(CARD); border(box, LINE)
        border(ws.Range(ws.Cells(6, c0), ws.Cells(6, c0 + 2)), ACCENT, (8,), -4138)
        ws.Range(ws.Cells(6, c0), ws.Cells(6, c0 + 2)).Merge()
        style(ws.Cells(6, c0), fg=MUTED, size=8, bold=True).Value = lab; ws.Cells(6, c0).IndentLevel = 1
        ws.Range(ws.Cells(7, c0), ws.Cells(8, c0 + 2)).Merge()
        v = ws.Cells(7, c0); v.NumberFormat = fmt; v.Formula = ff
        style(v, fg=ACCENT, size=22, name=GEORGIA, v="c"); v.IndentLevel = 1
        ws.Range(ws.Cells(9, c0), ws.Cells(9, c0 + 2)).Merge()
        style(ws.Cells(9, c0), fg=MUTED, size=8).Formula = note; ws.Cells(9, c0).IndentLevel = 1
    for r, hgt in ((5, 8), (6, 20), (7, 18), (8, 18), (9, 16), (10, 6), (11, 10)):
        ws.Rows(r).RowHeight = hgt
    style(ws.Range("B12"), fg=ACCENT, size=8, bold=True).Value = "W H A T   T H E   N U M B E R S   S A Y"
    ins = [
        '="•  Revenue "&IF(Growth>=0,"grew ","fell ")&TEXT(ABS(Growth),"0.0%")&" on the same months of 2025, to "&TEXT(Rev_26,"[$$-409]#,##0")&"; the spring peak was "&INDEX(Calc!$D$3:$D$12,MATCH(MAX(Calc!$E$3:$E$12),Calc!$E$3:$E$12,0))&" ("&TEXT(MAX(Calc!$E$3:$E$12),"[$$-409]#,##0")&")."',
        '="•  Fastest-growing region: "&INDEX(Calc!$H$3:$H$6,MATCH(MAX(Calc!$K$3:$K$6),Calc!$K$3:$K$6,0))&" ("&TEXT(MAX(Calc!$K$3:$K$6),"+0%")&"); slowest: "&INDEX(Calc!$H$3:$H$6,MATCH(MIN(Calc!$K$3:$K$6),Calc!$K$3:$K$6,0))&" ("&TEXT(MIN(Calc!$K$3:$K$6),"+0%;-0%")&")."',
        '="•  "&Calc!$P$8&" brings "&TEXT(Calc!$Q$8/Rev_26,"0%")&" of revenue; the best seller is "&Calc!$V$10&" ("&TEXT(Calc!$W$10,"[$$-409]#,##0")&"). Gross margin "&TEXT(Margin_26,"0.0%")&"."']
    for k, ff in enumerate(ins):
        ws.Range(f"B{13 + k}:T{13 + k}").Merge()
        style(ws.Range(f"B{13 + k}"), size=10).Formula = ff
    ws.Rows(16).RowHeight = 10

    def chart(tl, br, kind, series, title):
        """series: [(name, values range, x range)] - set one by one, so a numeric header is never read as data"""
        a, b = ws.Range(tl), ws.Range(br)
        co = ws.ChartObjects().Add(a.Left, a.Top, b.Left + b.Width - a.Left, b.Top + b.Height - a.Top)
        ch = co.Chart; ch.ChartType = kind
        while ch.SeriesCollection().Count:
            ch.SeriesCollection(1).Delete()
        for nm, vals, xs in series:
            se = ch.SeriesCollection().NewSeries(); se.Name = nm; se.Values = vals; se.XValues = xs
        chart_style(ch)
        ch.HasTitle = True; ch.ChartTitle.Text = title
        t = ch.ChartTitle.Format.TextFrame2.TextRange.Font; t.Size = 11; t.Bold = True; t.Fill.ForeColor.RGB = hexrgb(TEXT); t.Name = BODY
        ch.ChartTitle.Left = 8
        return co, ch

    C = W["Calc"]
    co, ch = chart("B17", "L33", 51, [("2026", C.Range("E3:E12"), C.Range("D3:D12")), ("2025", C.Range("F3:F12"), C.Range("D3:D12"))], "Revenue by month: 2026 vs 2025")
    s26, s25 = ch.SeriesCollection(1), ch.SeriesCollection(2)
    s26.Format.Fill.ForeColor.RGB = hexrgb(ACCENT)
    s25.ChartType = 65; s25.Format.Line.ForeColor.RGB = hexrgb(S2); s25.Format.Line.Weight = 2
    s25.MarkerStyle = 8; s25.MarkerSize = 5; s25.MarkerBackgroundColor = hexrgb(S2); s25.MarkerForegroundColor = hexrgb(S2)
    ch.ChartGroups(1).GapWidth = 60
    ch.HasLegend = True; ch.Legend.Position = -4160
    ax = ch.Axes(2); ax.MajorGridlines.Format.Line.ForeColor.RGB = hexrgb(LINE); ax.TickLabels.NumberFormat = "[$$-409]#,##0"; ax.Format.Line.Visible = False
    co.Name = "chtMonths"

    co, ch = chart("N17", "T33", 57, [("2026", C.Range("I3:I6"), C.Range("H3:H6")), ("2025", C.Range("J3:J6"), C.Range("H3:H6"))], "Revenue by region")
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = hexrgb(ACCENT)
    ch.SeriesCollection(2).Format.Fill.ForeColor.RGB = hexrgb(LY)
    ch.SeriesCollection(1).HasDataLabels = True; ch.SeriesCollection(1).DataLabels().NumberFormat = "[$$-409]#,##0"
    ch.ChartGroups(1).GapWidth = 50; ch.ChartGroups(1).Overlap = 0
    ch.Axes(1).ReversePlotOrder = True
    ch.HasLegend = True; ch.Legend.Position = -4160
    ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = hexrgb(LINE); ch.Axes(2).TickLabels.NumberFormat = "[$$-409]#,##0"
    co.Name = "chtRegions"

    co, ch = chart("B35", "J51", 57, [("2026", C.Range("Q3:Q8"), C.Range("P3:P8"))], "Revenue by category")
    ch.HasLegend = False
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = hexrgb(ACCENT)
    ch.SeriesCollection(1).HasDataLabels = True; ch.SeriesCollection(1).DataLabels().NumberFormat = "[$$-409]#,##0"
    ch.ChartGroups(1).GapWidth = 50
    ch.Axes(2).MajorGridlines.Delete(); ch.Axes(2).TickLabels.NumberFormat = "[$$-409]#,##0"
    co.Name = "chtCategories"

    co, ch = chart("L35", "T51", 57, [("2026", C.Range("W3:W10"), C.Range("V3:V10"))], "Top 8 products")
    ch.HasLegend = False
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = hexrgb(S2)
    ch.SeriesCollection(1).HasDataLabels = True; ch.SeriesCollection(1).DataLabels().NumberFormat = "[$$-409]#,##0"
    ch.ChartGroups(1).GapWidth = 50
    ch.Axes(2).MajorGridlines.Delete(); ch.Axes(2).TickLabels.NumberFormat = "[$$-409]#,##0"
    co.Name = "chtProducts"
    ws.Rows(34).RowHeight = 10
    ws.Range("B53:T53").Merge()
    style(ws.Range("B53"), fg=MUTED, size=8).Formula = (
        '="Checks: "&ChecksPassed&" / "&ChecksTotal&" PASS   ·   9 number errors in the old dashboard: see Issues found and Bridge   ·   Sample project, fictional data"')
    ps = ws.PageSetup; ps.PrintArea = "$A$1:$U$54"; ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1

    # ------------------------------------------------------------------ Bridge
    ws = W["Bridge"]
    page(ws, "Bridge: from the old total to the right one", "Revenue January to October 2026. Every step is worked out from the Data sheet; the old figure is read from the old file.")
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = False
    ws.Columns("B").ColumnWidth = 68; ws.Columns("C").ColumnWidth = 14; ws.Columns("D").ColumnWidth = 9; ws.Columns("E").ColumnWidth = 3
    style(ws.Range("B6:D6"), fg="#FFFFFF", bg=ACCENT, bold=True).Value = ("Step", "Revenue", "Orders")
    T26, YES = f'({rng("J")}=2026)', f'({rng("O")}="Yes")'
    steps = [
        ("Old dashboard: TOTAL REVENUE 2026 (Dashboard!B3, =SUM of rows 2-700)", OLD["B3"], OLD["B6"]),
        ("+ October orders pasted below row 700, outside the SUM range",
         f'=SUMIFS({rng("K")},{rng("J")},2026,{rng("O")},"No",{rng("N")},1)', f'=COUNTIFS({rng("J")},2026,{rng("O")},"No",{rng("N")},1)'),
        ("+ revenue cells that came in as text, skipped by SUM",
         f'=SUMPRODUCT({T26}*{YES}*ISTEXT({rng("H")})*({rng("N")}=1),{rng("K")})', "=0"),
        ("- orders exported twice, counted twice",
         f'=-SUMPRODUCT({T26}*{YES}*ISNUMBER({rng("H")})*({rng("N")}=0),{rng("K")})', f'=-COUNTIFS({rng("J")},2026,{rng("O")},"Yes",{rng("N")},0)'),
        ("Right total (new dashboard)", "=SUM(C7:C10)", "=SUM(D7:D10)"),
    ]
    for i, (lab, rv, od) in enumerate(steps, 7):
        ws.Cells(i, 2).Value = lab
        ws.Cells(i, 3).Value = rv if not isinstance(rv, str) else None
        if isinstance(rv, str): ws.Cells(i, 3).Formula = rv
        if isinstance(od, str): ws.Cells(i, 4).Formula = od
        else: ws.Cells(i, 4).Value = od
    ws.Range("C7:C11").NumberFormat = "[$$-409]#,##0.00;-[$$-409]#,##0.00"; ws.Range("D7:D11").NumberFormat = "+#,##0;-#,##0;0"
    ws.Range("D7").NumberFormat = "#,##0"; ws.Range("D11").NumberFormat = "#,##0"
    style(ws.Range("B11:D11"), bold=True, bg="#E6EFE8")
    border(ws.Range("B6:D11"), LINE, (7, 8, 9, 10, 12))
    ws.Range("B13").Value = "Check: right total = the new dashboard"
    ws.Range("C13").Formula = '=IF(ABS(C11-Rev_26)<0.005,"PASS","FAIL")'
    ws.Range("B14").Value = "Check: the old total, recomputed from the data, = the value in the old file"
    ws.Range("C14").Formula = f'=IF(ABS(SUMPRODUCT({T26}*{YES}*ISNUMBER({rng("H")}),{rng("H")})-C7)<0.005,"PASS","FAIL")'
    ws.Range("B15").Value = "Check: orders, old count + steps = new count"
    ws.Range("C15").Formula = '=IF(D11=Orders_26,"PASS","FAIL")'
    for c in ("C13", "C14", "C15"):
        style(ws.Range(c), bold=True, h="c")
    ws.Range("B16").Value = "Note: the old order count (699) is the same range: it also misses the October orders and counts the duplicates twice."
    ws.Range("B16").Font.Color = hexrgb(MUTED); ws.Range("B16").Font.Size = 9
    # waterfall drawn as stacked columns with an invisible base (works in every Excel version)
    ws.Range("G6:J6").Value = ("", "Base", "Up", "Down")
    labels = ["Old total", "+ rows below 700", "+ text numbers", "- duplicates", "Right total"]
    for k, lab in enumerate(labels):
        r = 7 + k
        ws.Cells(r, 7).Value = lab
    ws.Range("H7").Value = 0; ws.Range("I7").Formula = "=C7"; ws.Range("J7").Value = 0
    ws.Range("H8").Formula = "=C7"; ws.Range("I8").Formula = "=C8"; ws.Range("J8").Value = 0
    ws.Range("H9").Formula = "=C7+C8"; ws.Range("I9").Formula = "=C9"; ws.Range("J9").Value = 0
    ws.Range("H10").Formula = "=C7+C8+C9+C10"; ws.Range("I10").Value = 0; ws.Range("J10").Formula = "=-C10"
    ws.Range("H11").Value = 0; ws.Range("I11").Formula = "=C11"; ws.Range("J11").Value = 0
    ws.Range("G6:J11").Font.Color = hexrgb(MUTED); ws.Range("G6:J11").Font.Size = 8
    ws.Range("H7:J11").NumberFormat = "#,##0"
    a, b = ws.Range("B18"), ws.Range("D36")
    co = ws.ChartObjects().Add(a.Left, a.Top, b.Left + b.Width - a.Left, b.Top + b.Height - a.Top)
    ch = co.Chart; ch.ChartType = 52
    while ch.SeriesCollection().Count:
        ch.SeriesCollection(1).Delete()
    for col in "HIJ":
        se = ch.SeriesCollection().NewSeries(); se.Name = ws.Range(f"{col}6").Value
        se.Values = ws.Range(f"{col}7:{col}11"); se.XValues = ws.Range("G7:G11")
    chart_style(ch)
    ch.SeriesCollection(1).Format.Fill.Visible = False
    ch.SeriesCollection(2).Format.Fill.ForeColor.RGB = hexrgb(ACCENT)
    ch.SeriesCollection(3).Format.Fill.ForeColor.RGB = hexrgb(BAD)
    for s in (2, 3):
        se = ch.SeriesCollection(s); se.HasDataLabels = True; se.DataLabels().NumberFormat = "[$$-409]#,##0;;"
        se.DataLabels().Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = hexrgb("#FFFFFF")
        se.DataLabels().Format.TextFrame2.TextRange.Font.Bold = True
    ch.SeriesCollection(2).Points(1).Format.Fill.ForeColor.RGB = hexrgb(LY)
    ch.ChartGroups(1).GapWidth = 40; ch.HasLegend = False
    ch.Axes(2).MinimumScale = 50000; ch.Axes(2).MaximumScale = 62000
    ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = hexrgb(LINE); ch.Axes(2).TickLabels.NumberFormat = "[$$-409]#,##0"
    ch.HasTitle = True; ch.ChartTitle.Text = "Old total to right total"
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11

    # ------------------------------------------------------------------ Issues found
    ws = W["Issues found"]
    page(ws, "Issues found in the old dashboard", "9 mistakes: 6 in the formulas, 3 in the data under them. Old values are read from Before-Client-Dashboard.xlsx; correct values are live formulas.")
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = False
    cols = [("#", 4), ("Where in the old file", 18), ("What was wrong", 46), ("Number affected", 20), ("Old file shows", 13),
            ("Correct", 13), ("Difference", 13), ("How the new file avoids it", 46)]
    for i, (h, w) in enumerate(cols, 1):
        ws.Cells(6, i).Value = h; ws.Columns(i).ColumnWidth = w
    ws.Columns("A").ColumnWidth = 4
    style(ws.Range("A6:H6"), fg="#FFFFFF", bg=ACCENT, bold=True, wrap=True, v="c")
    hose_new = f'SUMIFS({rng("K")},{rng("J")},2026,{rng("N")},1,{rng("L")},"Hose 25m")'
    issues = [
        ("Dashboard!B3", "TOTAL REVENUE is =SUM(H2:H700); the 27 late-October orders pasted below row 700 are left out.",
         "Total revenue 2026", None, "=Bridge!C8", "Totals use whole-column SUMIFS on the data; nothing depends on a fixed last row.", "$"),
        ("Data 2026, column H", "14 revenue cells were exported as text (\"1,234.50\"); SUM skips text without a warning.",
         "Total revenue 2026", None, "=Bridge!C9", "Column K turns text into numbers (NUMBERVALUE); Checks counts them.", "$"),
        ("Data 2026", "9 orders appear twice in the export (same order ID); both copies were added.",
         "Total revenue 2026", None, "=Bridge!C10", "Column M flags the second copy of an order ID; it is not counted.", "$"),
        ("Dashboard!E7", "The West formula adds up column I (Cost) instead of H (Revenue).",
         "Revenue West", OLD["E7"], "=Calc!I6", "One formula for all regions, from a region table.", "$"),
        ("Dashboard!B7", "Average of 699 row margins: a $3 seed packet weighs as much as a $289 mower.",
         "Margin %", OLD["B7"], "=Margin_26", "Margin = (revenue - cost) / revenue on the totals.", "%"),
        ("Dashboard!B20", "October was typed in by hand from a preliminary report (1-15 October) and never updated.",
         "Revenue October", OLD["B20"], "=Calc!E12", "Every month is a formula on the data.", "$"),
        ("Dashboard!B5", "Growth = (2026 - 2025) / 2026: divided by this year instead of last year.",
         "Growth %", OLD["B5"], "=Growth", "Growth = this year / last year - 1, on the same months.", "%"),
        ("Dashboard!E18", "TOTAL adds the 'Tools subtotal' row as well: hand and power tools are counted twice. The 3D pie shows the subtotal as a slice too.",
         "Category total", OLD["E18"], "=Rev_26", "No subtotal rows inside a table that gets summed; the chart uses the 6 categories only.", "$"),
        (f"Dashboard!{OLD['hose_cell']}", "6 rows say \"Hose 25m \" with a trailing space; SUMIF on \"Hose 25m\" does not match them.",
         "Revenue Hose 25m", OLD["hose"], "=" + hose_new, "Column L trims every product name before it is summed.", "$"),
    ]
    for k, (where, what, num, oldv, new, fix, kind) in enumerate(issues, 1):
        r = 6 + k
        ws.Cells(r, 1).Value = k; ws.Cells(r, 2).Value = where; ws.Cells(r, 3).Value = what; ws.Cells(r, 4).Value = num
        fmt = "[$$-409]#,##0.00;-[$$-409]#,##0.00" if kind == "$" else "0.0%;-0.0%"
        if oldv is None:                                      # parts of the total: their amount is the bridge step
            ws.Cells(r, 5).Value = "see Bridge"; ws.Cells(r, 6).Value = "see Bridge"
            ws.Cells(r, 7).Formula = new
        else:
            ws.Cells(r, 5).Value = oldv; ws.Cells(r, 6).Formula = new; ws.Cells(r, 7).Formula = f"=F{r}-E{r}"
        for c in (5, 6, 7):
            ws.Cells(r, c).NumberFormat = fmt
        ws.Cells(r, 8).Value = fix
    rr = ws.Range("A7:H15")
    style(rr, wrap=True, v="t", bg=CARD)
    ws.Range("E7:G15").HorizontalAlignment = -4152
    border(ws.Range("A6:H15"), LINE, (7, 8, 9, 10, 11, 12))
    style(ws.Range("G7:G15"), bold=True, fg=BAD)
    ws.Range("B17").Value = "Old file shows = the value Excel calculated in Before-Client-Dashboard.xlsx (opened and read when this file was built)."
    style(ws.Range("B17"), fg=MUTED, size=9)

    # ------------------------------------------------------------------ Checks
    ws = W["Checks"]
    page(ws, "Checks", "Every view adds up to the same total, the bridge closes, and every planted problem is found. All must say PASS.")
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = False
    ws.Columns("B").ColumnWidth = 58; ws.Columns("C").ColumnWidth = 14; ws.Columns("D").ColumnWidth = 10
    style(ws.Range("B6:D6"), fg="#FFFFFF", bg=ACCENT, bold=True).Value = ("Check", "Found", "Result")
    chk = [("Months add up to the total", "=SUM(Calc!E3:E12)-Rev_26", "=ABS(C{r})<0.005"),
           ("Regions add up to the total", "=SUM(Calc!I3:I6)-Rev_26", "=ABS(C{r})<0.005"),
           ("Categories add up to the total", "=SUM(Calc!N3:N8)-Rev_26", "=ABS(C{r})<0.005"),
           ("Products add up to the total", "=SUM(Calc!T3:T20)-Rev_26", "=ABS(C{r})<0.005"),
           ("Bridge: old total + steps = right total", "=Bridge!C11-Rev_26", "=ABS(C{r})<0.005"),
           ("Bridge: old total recomputed from the data = old file", "=Bridge!C14", '=C{r}="PASS"'),
           ("Orders exported twice found", f'=COUNTIF({rng("M")},"Duplicate")', "=C{r}=9"),
           ("Revenue cells stored as text found and converted", f'=SUMPRODUCT(--ISTEXT({rng("H")}))', "=C{r}=14"),
           ("Product names with extra spaces trimmed", f'=SUMPRODUCT(--({rng("E")}<>{rng("L")}))', "=C{r}=6"),
           ("Every line has a revenue number", f'=COUNTA({rng("A")})-COUNT({rng("K")})', "=C{r}=0"),
           ("Every line's revenue = units x unit price", f'=SUMPRODUCT(--(ABS({rng("K")}-{rng("F")}*{rng("G")})>0.005))', "=C{r}=0")]
    for k, (txt, ff, test) in enumerate(chk, 7):
        ws.Cells(k, 2).Value = txt; ws.Cells(k, 3).Formula = ff
        ws.Cells(k, 4).Formula = f'=IF({test.format(r=k)[1:]},"PASS","FAIL")'
    last = 6 + len(chk)
    ws.Range(f"C7:C{last}").NumberFormat = "#,##0.##"
    d = ws.Range(f"D7:D{last}")
    fc = d.FormatConditions.Add(1, 3, '="PASS"'); fc.Font.Color = hexrgb(GOOD); fc.Interior.Color = hexrgb("#E2F0E5")
    fc = d.FormatConditions.Add(1, 3, '="FAIL"'); fc.Font.Color = hexrgb(BAD); fc.Interior.Color = hexrgb("#F6DEDA")
    d.Font.Bold = True; d.HorizontalAlignment = -4108
    border(ws.Range(f"B6:D{last}"), LINE, (7, 8, 9, 10, 12))
    ws.Range(f"B{last + 2}").Value = "Checks passed"; ws.Range(f"D{last + 2}").Formula = f'=COUNTIF(D7:D{last},"PASS")'
    ws.Range(f"B{last + 3}").Value = "Checks in total"; ws.Range(f"D{last + 3}").Value = len(chk)
    wb.Names.Add("ChecksPassed", f"=Checks!$D${last + 2}"); wb.Names.Add("ChecksTotal", f"=Checks!$D${last + 3}")

    # ------------------------------------------------------------------ Before (picture) and How to use
    ws = W["Before"]
    page(ws, "Before: the old dashboard", "Before-Client-Dashboard.xlsx as received, for comparison. Its 9 number errors are listed on Issues found.")
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = False
    pic = ws.Shapes.AddPicture(str(ROOT / "docs" / "before.png"), False, True, ws.Range("B6").Left, ws.Range("B6").Top, -1, -1)
    pic.LockAspectRatio = True; pic.Width = 900
    ws = W["How to use"]
    page(ws, "How to use", "No macros. Excel 2016 or later.")
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = False
    ws.Columns("B").ColumnWidth = 120
    for i, t in enumerate([
            "1. Paste a new export below the last row on Data, columns A-I exactly as exported (text numbers and spaces are fine).",
            "2. Copy the formulas in columns J-P down to the new last row. Every total uses the whole data range named in Calc.",
            "3. Dashboard: cards, insights and the four charts follow at once. Checks must still say PASS.",
            "4. A line flagged 'exported twice' is not counted; the note column says why each line was changed.",
            "5. Issues found and Bridge belong to this review of the old file: keep them as the record of what was fixed.",
            "Theme: forest green and cream (brand colours can be swapped: one colour per role, nothing hard-coded in the charts).",
    ], 6):
        ws.Cells(i, 2).Value = t

    for n in ("Dashboard", "Issues found", "Bridge", "Checks", "Before", "How to use"):
        W[n].Protect(DrawingObjects=True, Contents=True)
    W["Dashboard"].Activate(); W["Dashboard"].Range("A1").Select()
    xl.ScreenUpdating = True
    xl.CalculateFull()
    wb.SaveAs(str(OUT), FileFormat=51)
    print("saved", OUT.name, "| checks", W["Checks"].Range(f"D{last + 2}").Value, "/", len(chk),
          "| revenue", round(xl.Evaluate("Rev_26"), 2) if not hasattr(xl.Evaluate("Rev_26"), "Value") else W["Calc"].Range("B3").Value)
    wb.Close(False)
finally:
    xl.Quit()
