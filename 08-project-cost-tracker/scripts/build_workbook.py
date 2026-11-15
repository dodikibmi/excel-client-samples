"""Builds Project-Cost-Tracker.xlsx with real Excel (COM).

One input sheet per project (identical layout, filled in by the PM), placed between the 'PM start' and 'PM end' sheets.
A Rollup sheet adds them up (earned value, CPI/SPI, three EAC methods, TCPI, change orders, WIP billing and cash), and
six dark-gold report sheets read it:
  Dashboard     KPI cards with CPI/SPI gauges, project RAG table, project x cost-code heatmap, charts, checks, PM notes
  Project view  pick a project: cards, waterfall budget -> forecast by cost code, earned value detail, cost codes, notes
  S-curve       pick a project: planned / earned / actual cumulative against the baseline, forecast finish, crew chart
  COs & risks   change orders (approved ones flow into contract and forecast) and a 5x5 risk matrix
  WIP & cash    percentage-of-completion WIP schedule: over/under billing, retainage, cash received vs paid
  Weekly trend  CPI, SPI, overrun and crew week by week from the History table
Data > Refresh All saves this week's snapshot into a History table (Power Query appends to its own output, so the
history survives without macros). Navigation = shapes with hyperlinks. The workbook has no macros.

Run with a Python that has pywin32:  python scripts/build_workbook.py      (after scripts/make_data.py)
"""
import csv, json, time
from datetime import date
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Project-Cost-Tracker.xlsx"
DATA = ROOT / "data"
EXP = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
WEEK = date.fromisoformat(EXP["week_ending"])

# ---------------------------------------------------------------- dark-gold palette (same as E9)
HEAD, PAGE, CARD, CARD2, LINE = (11, 15, 19), (18, 23, 29), (26, 33, 41), (32, 40, 50), (42, 51, 61)
GOLD, GOLD_LT, IVORY, MUTED = (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163)
JADE, COPPER, SKY, GREY = (79, 163, 140), (181, 117, 74), (127, 167, 201), (74, 85, 96)
GOOD, BAD, AMBER = (91, 191, 138), (229, 115, 95), (217, 164, 65)
GOOD_BG, AMBER_BG, BAD_BG = (27, 52, 41), (58, 46, 22), (64, 32, 29)
INPUT_BG = (48, 41, 26)
SERIF, SANS, SYMBOL = "Georgia", "Segoe UI", "Segoe UI Symbol"
LIGHT_HEAD = (31, 58, 90)
SERIES = [SKY, COPPER, JADE, GOLD_LT, (160, 130, 200)]

FRONT = ["Dashboard", "Project view", "S-curve", "COs & risks", "WIP & cash", "Weekly trend", "Snapshots", "PM start"]
TAIL = ["PM end", "Template", "Rollup", "Baseline", "Cost codes", "Settings", "How to use"]
NAV = ["Dashboard", "Project view", "S-curve", "COs & risks", "WIP & cash", "Weekly trend", "PM sheets"]
MONEY = '#,##0;-#,##0;"–"'
KMONEY = '#,##0,"k";-#,##0,"k";"–"'
SLOTS = 5  # project rows on the Dashboard
HIST_COLS = ["Week ending", "Sheet", "Project", "Contract", "Budget", "Committed", "Actual", "Earned", "Planned", "Field %",
             "Shop %", "CPI", "SPI", "Forecast at completion", "Planned heads", "Actual heads", "Billed", "Cash received",
             "Cash paid"]
SEP = ","  # the local list separator, set in main(): conditional formats, validation and names are read in local syntax


def rgb(c):
    return c[0] + c[1] * 256 + c[2] * 65536


def loc(f):
    return f.replace(",", SEP)


def col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def read_csv(name):
    with open(DATA / name, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def serial(d):
    return (d - date(1899, 12, 30)).days


def put(cell, v):
    if isinstance(v, str) and v.startswith("="):
        cell.Formula = v
    else:
        cell.Value = v


# ---------------------------------------------------------------- styling helpers
def dark(ws, widths, tab=GOLD):
    ws.Cells.Interior.Color = rgb(PAGE)
    ws.Cells.Font.Color = rgb(IVORY)
    ws.Cells.Font.Name = SANS
    ws.Cells.Font.Size = 10
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    ws.Tab.Color = rgb(tab)


def header(ws, last, title, sub, active, first_pm):
    ws.Range(f"A1:{last}3").Interior.Color = rgb(HEAD)
    ws.Range(f"A4:{last}4").Interior.Color = rgb(GOLD)
    for r, h in ((1, 16), (2, 30), (3, 18), (4, 2.25), (5, 9)):
        ws.Rows(r).RowHeight = h
    c = ws.Range("B1")
    c.Value = "B R I G H T W I R E    ·    P R O J E C T   C O N T R O L S"
    c.Font.Size = 8; c.Font.Bold = True; c.Font.Color = rgb(GOLD); c.VerticalAlignment = -4107
    c = ws.Range("B2")
    put(c, title)
    c.Font.Name = SERIF; c.Font.Size = 20; c.Font.Color = rgb(IVORY); c.VerticalAlignment = -4108
    c = ws.Range("B3")
    put(c, sub)
    c.Font.Size = 9; c.Font.Color = rgb(MUTED); c.VerticalAlignment = -4160
    nav(ws, last, active, first_pm)


def nav(ws, last, active, first_pm):
    right = ws.Range(f"{last}1").Left + ws.Range(f"{last}1").Width - 4
    bw, bh, gap = 74, 20, 4
    x = right - len(NAV) * bw - (len(NAV) - 1) * gap
    top = ws.Range("A2").Top + 5
    for label in NAV:
        target = first_pm if label == "PM sheets" else label
        on = label == active
        shp = ws.Shapes.AddShape(5, x, top, bw, bh)
        shp.Name = f"Nav {label}"
        shp.Fill.ForeColor.RGB = rgb(GOLD if on else CARD)
        shp.Line.ForeColor.RGB = rgb(GOLD if on else LINE)
        shp.Line.Weight = 0.75
        tf = shp.TextFrame2
        tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
        tf.VerticalAnchor = 3
        tr = tf.TextRange
        tr.Text = label
        tr.ParagraphFormat.Alignment = 2
        tr.Font.Size = 8.5; tr.Font.Name = SANS; tr.Font.Bold = on
        tr.Font.Fill.ForeColor.RGB = rgb(HEAD if on else IVORY)
        shp.Placement = 3
        ws.Hyperlinks.Add(shp, "", f"'{target}'!A1", f"Go to {label}")
        x += bw + gap


def textbox(ws, x, y, w, h, text=None, size=10, color=IVORY, bold=False, font=SANS, link=None, align=1, wrap=False, anchor=1):
    tb = ws.Shapes.AddTextbox(1, x, y, w, h)
    tb.Fill.Visible = False
    tb.Line.Visible = False
    tf = tb.TextFrame2
    tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
    tf.WordWrap = wrap
    tf.AutoSize = 0
    tf.VerticalAnchor = anchor
    if link:
        tb.DrawingObject.Formula = link
    else:
        tf.TextRange.Text = text
    tr = tf.TextRange
    tr.Font.Size = size; tr.Font.Name = font; tr.Font.Bold = bold
    tr.Font.Fill.ForeColor.RGB = rgb(color)
    tr.ParagraphFormat.Alignment = align
    tb.Placement = 3
    return tb


class Texts:
    """Formula cells on the Rollup sheet that card textboxes link to (one cell per text)."""

    def __init__(self, ws, column="BE", first=4):
        self.ws, self.col, self.r = ws, column, first
        lab = col(self.col_n() - 1)
        ws.Range(f"{lab}{first - 1}").Value = "Card texts (linked to the report shapes)"
        ws.Range(f"{lab}{first - 1}").Font.Bold = True

    def col_n(self):
        n = 0
        for ch in self.col:
            n = n * 26 + ord(ch) - 64
        return n

    def add(self, label, formula, fmt=None, name=None):
        c = self.ws.Range(f"{self.col}{self.r}")
        self.ws.Range(f"{col(self.col_n() - 1)}{self.r}").Value = label
        c.Formula = formula
        if fmt:
            c.NumberFormat = fmt
        if name:
            self.ws.Parent.Names.Add(name, f"=Rollup!${self.col}${self.r}")
        self.r += 1
        return f"=Rollup!${self.col}${self.r - 1}"


def card(ws, x, y, w, h, label, value_link, notes=(), icon=None, value_size=21, value_w=None):
    """A KPI card: rounded box, gold accent, label, big value (linked to a cell), note lines (one per colour; the
    note cells are empty unless their condition holds, so exactly one shows)."""
    bg = ws.Shapes.AddShape(5, x, y, w, h)
    try:
        bg.Adjustments.SetItem(1, 0.07)
    except Exception:
        pass
    bg.Fill.ForeColor.RGB = rgb(CARD)
    bg.Line.ForeColor.RGB = rgb(LINE); bg.Line.Weight = 0.75
    bg.Placement = 3
    bar = ws.Shapes.AddShape(1, x + 1, y + 10, 2.25, h - 20)
    bar.Fill.ForeColor.RGB = rgb(GOLD); bar.Line.Visible = False; bar.Placement = 3
    textbox(ws, x + 12, y + 9, w - 46, 13, label.upper(), 7.5, MUTED, bold=True)
    textbox(ws, x + 12, y + 25, value_w or (w - 24), 32, size=value_size, color=IVORY, font=SERIF, link=value_link)
    for link, color in notes:
        textbox(ws, x + 12, y + h - 21, w - 18, 13, size=8, color=color, link=link)
    if icon:
        ov = ws.Shapes.AddShape(9, x + w - 31, y + 9, 21, 21)
        ov.Fill.Visible = False
        ov.Line.ForeColor.RGB = rgb(GOLD); ov.Line.Weight = 1
        tf = ov.TextFrame2
        tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
        tf.VerticalAnchor = 3
        tf.TextRange.Text = icon
        tf.TextRange.ParagraphFormat.Alignment = 2
        tf.TextRange.Font.Size = 10; tf.TextRange.Font.Name = SYMBOL
        tf.TextRange.Font.Fill.ForeColor.RGB = rgb(GOLD)
        ov.Placement = 3


def gauge(ws, x, y, size, values):
    """Half-doughnut gauge: points = green arc, amber arc, red arc, rest of the track, hidden lower half."""
    co = ws.ChartObjects().Add(x, y, size, size)
    ch = co.Chart
    s = ch.SeriesCollection().NewSeries()
    s.Values = values
    ch.ChartType = -4120
    ch.HasLegend = False
    ch.HasTitle = False
    ch.ChartArea.Format.Fill.Visible = False
    ch.ChartArea.Format.Line.Visible = False
    ch.PlotArea.Format.Fill.Visible = False
    g = ch.ChartGroups(1)
    g.FirstSliceAngle = 270
    g.DoughnutHoleSize = 62
    for i, c in enumerate((GOOD, AMBER, BAD, LINE, None), 1):
        p = s.Points(i)
        p.Format.Line.Visible = False
        if c is None:
            p.Format.Fill.Visible = False
        else:
            p.Format.Fill.Visible = True
            p.Format.Fill.ForeColor.RGB = rgb(c)
    co.Placement = 3
    return co


def cards_row(ws, first_col, last_col, top_row, n, height=86, gap=10):
    """x positions for n evenly spaced cards between two columns, starting at a row."""
    left = ws.Range(f"{first_col}1").Left
    right = ws.Range(f"{last_col}1").Left + ws.Range(f"{last_col}1").Width
    w = (right - left - gap * (n - 1)) / n
    y = ws.Range(f"A{top_row}").Top
    return [(left + i * (w + gap), y, w, height) for i in range(n)]


def note3(T, label, text, good, bad):
    """Three cells: the note text when good / in between / bad (the others empty)."""
    return [(T.add(label + " (good)", f'=IFERROR(IF({good},{text},""),"")'), GOOD),
            (T.add(label + " (amber)", f'=IFERROR(IF(AND(NOT({good}),NOT({bad})),{text},""),"")'), AMBER),
            (T.add(label + " (bad)", f'=IFERROR(IF({bad},{text},""),"")'), BAD)]


def section(ws, cell, text, size=13):
    c = ws.Range(cell)
    c.Value = text
    c.Font.Name = SERIF; c.Font.Size = size; c.Font.Color = rgb(IVORY)
    c.VerticalAlignment = -4107


def table_head(rng, wrap=True):
    rng.Interior.Color = rgb(HEAD)
    rng.Font.Color = rgb(GOLD); rng.Font.Bold = True; rng.Font.Size = 8.5
    rng.WrapText = wrap
    rng.VerticalAlignment = -4108
    b = rng.Borders(9); b.LineStyle = 1; b.Color = rgb(GOLD); b.Weight = 2


def table_body(rng):
    rng.Interior.Color = rgb(CARD)
    rng.Font.Size = 9.5
    rng.VerticalAlignment = -4108
    b = rng.Borders(12); b.LineStyle = 1; b.Color = rgb(LINE)
    b = rng.Borders(9); b.LineStyle = 1; b.Color = rgb(LINE)


def table_total(rng):
    rng.Interior.Color = rgb(CARD2)
    rng.Font.Bold = True
    b = rng.Borders(8); b.LineStyle = 1; b.Color = rgb(GOLD)


def cf(rng, formula, font=None, fill=None, bold=None):
    fc = rng.FormatConditions.Add(2, None, loc(formula))
    if font:
        fc.Font.Color = rgb(font)
    if fill:
        fc.Interior.Color = rgb(fill)
    if bold:
        fc.Font.Bold = True
    return fc


def neg_red(rng):
    fc = rng.FormatConditions.Add(1, 6, "=0")  # less than 0
    fc.Font.Color = rgb(BAD)


def pills(rng):
    for text, (bg, fg) in (("On track", (GOOD_BG, GOOD)), ("Watch", (AMBER_BG, AMBER)), ("At risk", (BAD_BG, BAD))):
        fc = rng.FormatConditions.Add(1, 3, f'="{text}"')
        fc.Interior.Color = rgb(bg); fc.Font.Color = rgb(fg); fc.Font.Bold = True


def pass_fail(rng):
    fc = rng.FormatConditions.Add(1, 3, '="PASS"')
    fc.Interior.Color = rgb(GOOD_BG); fc.Font.Color = rgb(GOOD)
    fc = rng.FormatConditions.Add(1, 4, '="PASS"')
    fc.Interior.Color = rgb(BAD_BG); fc.Font.Color = rgb(BAD)


def rag_index(rng, first_cell, green, amber):
    """Font colour for CPI/SPI style indexes, three exclusive rules."""
    a = first_cell
    cf(rng, f"=AND(ISNUMBER({a}),{a}<{amber})", font=BAD, bold=True)
    cf(rng, f"=AND(ISNUMBER({a}),{a}>={amber},{a}<{green})", font=AMBER, bold=True)
    cf(rng, f"=AND(ISNUMBER({a}),{a}>={green})", font=GOOD)


def hide_errors(rng, first_cell, bg=CARD):
    cf(rng, f"=ISERROR({first_cell})", font=bg)


def dark_chart(ch, title=None, legend=True, size=12):
    ch.ChartArea.Format.Fill.Visible = True
    ch.ChartArea.Format.Fill.ForeColor.RGB = rgb(CARD)
    ch.ChartArea.Format.Line.Visible = True
    ch.ChartArea.Format.Line.ForeColor.RGB = rgb(LINE)
    ch.ChartArea.RoundedCorners = False
    ch.PlotArea.Format.Fill.Visible = False
    ch.ChartArea.Font.Name = SANS
    ch.ChartArea.Font.Size = 8.5
    ch.ChartArea.Font.Color = rgb(MUTED)
    if title:
        ch.HasTitle = True
        ch.ChartTitle.Text = title
        ch.ChartTitle.Font.Name = SERIF
        ch.ChartTitle.Font.Size = size
        ch.ChartTitle.Font.Bold = False
        ch.ChartTitle.Font.Color = rgb(IVORY)
        ch.ChartTitle.Left = 8
        ch.ChartTitle.Top = 4
    else:
        ch.HasTitle = False
    ch.HasLegend = legend
    if legend:
        ch.Legend.Position = -4160  # top
        ch.Legend.Font.Color = rgb(IVORY)
        ch.Legend.Font.Size = 8.5
    for ax_type in (1, 2):
        try:
            ax = ch.Axes(ax_type)
            ax.Format.Line.Visible = True
            ax.Format.Line.ForeColor.RGB = rgb(LINE)
            ax.TickLabels.Font.Color = rgb(MUTED)
            ax.MajorTickMark = -4142
            if ax_type == 2:
                ax.HasMajorGridlines = True
                ax.MajorGridlines.Format.Line.ForeColor.RGB = rgb(LINE)
                ax.MajorGridlines.Format.Line.Weight = 0.5
                ax.Format.Line.Visible = False
        except Exception:
            pass


def add_chart(ws, rng_addr, ctype):
    rng = ws.Range(rng_addr)
    co = ws.ChartObjects().Add(rng.Left + 2, rng.Top + 2, rng.Width - 4, rng.Height - 4)
    co.Placement = 3
    ch = co.Chart
    ch.ChartType = ctype
    while ch.SeriesCollection().Count:
        ch.SeriesCollection(1).Delete()
    return ch


def add_series(ch, name, values, xvalues=None, color=None, kind=None, line=False, weight=2.25, dash=None, marker=False):
    s = ch.SeriesCollection().NewSeries()
    s.Name = name
    s.Values = values
    if xvalues is not None:
        s.XValues = xvalues
    if kind:
        s.ChartType = kind
    if color:
        if line:
            s.Format.Line.Visible = True
            s.Format.Line.ForeColor.RGB = rgb(color)
            s.Format.Line.Weight = weight
            if dash:
                s.Format.Line.DashStyle = dash
            if marker:
                s.MarkerStyle = 8
                s.MarkerSize = 5
                s.MarkerBackgroundColor = rgb(color)
                s.MarkerForegroundColor = rgb(color)
            else:
                s.MarkerStyle = -4142
        else:
            s.Format.Fill.ForeColor.RGB = rgb(color)
            s.Format.Line.Visible = False
    return s


# ---------------------------------------------------------------- settings, cost codes, baseline
def build_settings(wb):
    st = wb.Worksheets("Settings")
    st.Range("A1").Value = "Settings"
    st.Range("A1").Font.Size = 16; st.Range("A1").Font.Bold = True
    rows = [("Company", "Brightwire Electrical Contractors", "Company", "Shown in the report titles (fictional company)", "@"),
            ("Week ending", WEEK, "WeekEnding", "Change every week, then Data > Refresh All: the snapshot for this date is saved", "d mmm yyyy"),
            ("On track if CPI at least", 0.98, "GreenCPI", "Cost performance index = earned value / actual cost", "0.00"),
            ("Watch if CPI at least", 0.92, "AmberCPI", "Below this a project is At risk on cost", "0.00"),
            ("On track if SPI at least", 0.95, "GreenSPI", "Schedule performance index = earned value / planned value", "0.00"),
            ("Watch if SPI at least", 0.90, "AmberSPI", "Below this a project is At risk on schedule", "0.00")]
    st.Range("A3:C3").Value = ("Setting", "Value", "Note")
    st.Range("A3:C3").Font.Bold = True
    for r, (label, val, nm, note, fmt) in enumerate(rows, 4):
        st.Cells(r, 1).Value = label
        st.Cells(r, 2).NumberFormat = fmt
        if isinstance(val, date):  # a serial number: a datetime would be shifted by the PC's time zone
            st.Cells(r, 2).Value2 = serial(val)
        else:
            st.Cells(r, 2).Value = val
        st.Cells(r, 2).HorizontalAlignment = -4131
        st.Cells(r, 3).Value = note
        st.Cells(r, 2).Interior.Color = rgb((255, 255, 204))
        wb.Names.Add(nm, f"=Settings!$B${r}")
    st.Range("A12").Value = "Change order status"
    st.Range("A12").Font.Bold = True
    for i, s in enumerate(("Approved", "Pending", "Rejected"), 13):
        st.Cells(i, 1).Value = s
    wb.Names.Add("StatusList", "=Settings!$A$13:$A$15")
    st.Range("B12").Value = "Heatmap metric"
    st.Range("B12").Font.Bold = True
    st.Range("B13").Value = "CPI"; st.Range("B14").Value = "SPI"
    wb.Names.Add("MetricList", "=Settings!$B$13:$B$14")
    st.Columns("A").ColumnWidth = 26; st.Columns("B").ColumnWidth = 34; st.Columns("C").ColumnWidth = 80


def build_codes(wb):
    ws = wb.Worksheets("Cost codes")
    codes = read_csv("cost-codes.csv")
    ws.Range("A1").Value = "Cost codes (the only codes a PM sheet accepts)"
    ws.Range("A1").Font.Size = 14; ws.Range("A1").Font.Bold = True
    ws.Range("A3:C3").Value = ("Cost code", "Description", "Category")
    for i, c in enumerate(codes, 4):
        ws.Cells(i, 1).NumberFormat = "@"
        ws.Cells(i, 1).Value = c["Cost code"]
        ws.Cells(i, 2).Value = c["Description"]
        ws.Cells(i, 3).Value = c["Category"]
    lo = ws.ListObjects.Add(1, ws.Range(f"A3:C{3 + len(codes)}"), None, 1)
    lo.Name = "CostCodes"
    lo.TableStyle = "TableStyleMedium2"
    wb.Names.Add("CodeList", "=CostCodes[Cost code]")
    ws.Columns("A").ColumnWidth = 11; ws.Columns("B").ColumnWidth = 44; ws.Columns("C").ColumnWidth = 18
    ws.Range("A16").Value = "Add a code: type a new row at the bottom of the table; every PM sheet's drop-down picks it up."
    ws.Range("A16").Font.Italic = True
    return [c["Cost code"] for c in codes]


def build_baseline(wb):
    ws = wb.Worksheets("Baseline")
    rows = read_csv("baseline.csv")
    ws.Range("A1").Value = "Baseline schedule (cost-loaded): planned % to date, planned value and planned crew per project, cost code and week"
    ws.Range("A1").Font.Size = 14; ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = ("From the project schedule. The PM sheets look up their planned % for the week here (the last week "
                            "once the plan is finished); the S-curve draws it to the end of the project.")
    ws.Range("A2").Font.Italic = True
    data = [["Sheet", "Cost code", "Week ending", "Planned %", "Planned value", "Planned heads"]]
    for r in rows:
        data.append([r["sheet"], r["code"], serial(date.fromisoformat(r["week"])), float(r["planned_pct"]),
                     float(r["planned_value"]), int(r["planned_heads"])])
    ws.Range(f"A4:B{3 + len(data)}").NumberFormat = "@"
    ws.Range(ws.Cells(4, 1), ws.Cells(3 + len(data), 6)).Value = data
    lo = ws.ListObjects.Add(1, ws.Range(f"A4:F{3 + len(data)}"), None, 1)
    lo.Name = "Baseline"
    lo.TableStyle = "TableStyleMedium2"
    lo.ListColumns("Week ending").DataBodyRange.NumberFormat = "d mmm yyyy"
    lo.ListColumns("Planned %").DataBodyRange.NumberFormat = "0.0%"
    lo.ListColumns("Planned value").DataBodyRange.NumberFormat = "#,##0"
    for c, w in zip("ABCDEF", (10, 11, 14, 11, 15, 14)):
        ws.Columns(c).ColumnWidth = w
    return len({r["week"] for r in rows})


# ---------------------------------------------------------------- one PM sheet
FIRST, LAST = 18, 27  # cost code rows
TOT = 28
PMC = dict(id="C6", name="C7", pm="C8", contract="C9", billed="I6", ret="I7", recv="I8", paid="I9", heads="I10",
           pheads="I11", cpi="M6", spi="M7", field="M8", shop="M9", fac="M10", margin="M11", unknown="Q6", pctbad="Q7",
           overcommit="Q8", note="C13")


def build_pm(ws, codes, first_pm, proj=None, lines=None):
    widths = {"A": 2, "B": 15, "C": 34, "D": 15, "E": 12, "F": 12, "G": 12, "H": 9, "I": 9, "J": 12, "K": 12, "L": 11,
              "M": 11, "N": 7, "O": 7, "P": 12, "Q": 12}
    dark(ws, widths)
    header(ws, "Q", '="Project cost report  ·  "&$C$6&"  ·  "&$C$7',
           "Filled in by the project manager every week. Gold cells are inputs; planned % comes from the Baseline schedule.",
           "PM sheets", first_pm)
    ws.Rows(5).RowHeight = 18
    for r in range(6, 12):
        ws.Rows(r).RowHeight = 19
    for block in ("B6:D11", "F6:J11", "K6:M11", "O5:Q11"):
        ws.Range(block).Interior.Color = rgb(CARD)
    left = [("Project ID", proj["Sheet"] if proj else "P-0000"), ("Project", proj["Project"] if proj else "New project name"),
            ("Project manager", proj["PM"] if proj else "PM name"), ("Contract value", float(proj["Contract value"]) if proj else 0),
            ("Week ending", "=WeekEnding")]
    for i, (label, v) in enumerate(left):
        r = 6 + i
        ws.Range(f"B{r}").Value = label
        if i == 0:
            ws.Range(f"C{r}").NumberFormat = "@"
        put(ws.Range(f"C{r}"), v)
    ws.Range("C9").NumberFormat = "#,##0"
    ws.Range("C10").NumberFormat = "d mmm yyyy"
    ws.Range("C6:C10").HorizontalAlignment = -4131
    ws.Range("C6:C7").Font.Bold = True
    mid = [("Billed to date", float(proj["Billed to date"]) if proj else 0, "#,##0"),
           ("Retainage %", float(proj["Retainage %"]) if proj else 0.1, "0%"),
           ("Cash received to date", float(proj["Cash received to date"]) if proj else 0, "#,##0"),
           ("Cash paid to date", float(proj["Cash paid to date"]) if proj else 0, "#,##0"),
           ("Crew on site (heads)", float(proj["Crew on site"]) if proj else 0, "0"),
           ("Planned crew (baseline)", "=SUMIFS(Baseline[Planned heads],Baseline[Sheet],$C$6,Baseline[Week ending],MIN(WeekEnding,MAX(Baseline[Week ending])))", "0")]
    for i, (label, v, fmt) in enumerate(mid):
        r = 6 + i
        ws.Range(f"F{r}").Value = label
        ws.Range(f"I{r}:J{r}").Merge()
        put(ws.Range(f"I{r}"), v)
        ws.Range(f"I{r}").NumberFormat = fmt
        ws.Range(f"I{r}").HorizontalAlignment = -4152
    right = [("Cost performance (CPI)", f"=N{TOT}", "0.00"), ("Schedule perf. (SPI)", f"=O{TOT}", "0.00"),
             ("Field % complete", f'=IF(SUMPRODUCT((D{FIRST}:D{LAST}<>"Shop fabrication")*E{FIRST}:E{LAST})=0,"",'
                                  f'SUMPRODUCT((D{FIRST}:D{LAST}<>"Shop fabrication")*E{FIRST}:E{LAST}*I{FIRST}:I{LAST})'
                                  f'/SUMPRODUCT((D{FIRST}:D{LAST}<>"Shop fabrication")*E{FIRST}:E{LAST}))', "0.0%"),
             ("Shop fab % complete", f'=IF(SUMIF(D{FIRST}:D{LAST},"Shop fabrication",E{FIRST}:E{LAST})=0,"",'
                                     f'SUMPRODUCT((D{FIRST}:D{LAST}="Shop fabrication")*E{FIRST}:E{LAST}*I{FIRST}:I{LAST})'
                                     f'/SUMIF(D{FIRST}:D{LAST},"Shop fabrication",E{FIRST}:E{LAST}))', "0.0%"),
             ("Forecast at completion", f"=P{TOT}", "#,##0"),
             ("Projected margin", '=IF(C9=0,"",C9-M10)', MONEY)]
    for i, (label, f, fmt) in enumerate(right):
        r = 6 + i
        ws.Range(f"K{r}").Value = label
        ws.Range(f"M{r}").Formula = f
        ws.Range(f"M{r}").NumberFormat = fmt
        ws.Range(f"M{r}").Font.Bold = True
        ws.Range(f"M{r}").HorizontalAlignment = -4152
    ws.Range("O5").Value = "Sheet checks"
    ws.Range("O5").Font.Bold = True; ws.Range("O5").Font.Color = rgb(GOLD)
    chk = [("Unknown codes", f'=SUMPRODUCT((B{FIRST}:B{LAST}<>"")*(COUNTIF(CodeList,B{FIRST}:B{LAST})=0))'),
           ("% outside 0-100", f'=COUNTIF(H{FIRST}:I{LAST},"<0")+COUNTIF(H{FIRST}:I{LAST},">1")'),
           ("Actual > committed", f"=SUMPRODUCT(--(G{FIRST}:G{LAST}>F{FIRST}:F{LAST}))")]
    for i, (label, f) in enumerate(chk):
        r = 6 + i
        ws.Range(f"O{r}").Value = label
        ws.Range(f"Q{r}").Formula = f
        ws.Range(f"Q{r}").HorizontalAlignment = -4108
        fc = ws.Range(f"Q{r}").FormatConditions.Add(1, 3, "=0")
        fc.Interior.Color = rgb(GOOD_BG); fc.Font.Color = rgb(GOOD)
        fc = ws.Range(f"Q{r}").FormatConditions.Add(1, 5, "=0")
        fc.Interior.Color = rgb(BAD_BG); fc.Font.Color = rgb(BAD)
    for rng in ("B6:B10", "F6:F11", "K6:K11", "O6:O8"):
        ws.Range(rng).Font.Color = rgb(MUTED)
        ws.Range(rng).Font.Size = 9
    rag_index(ws.Range("M6"), "M6", "GreenCPI", "AmberCPI")
    rag_index(ws.Range("M7"), "M7", "GreenSPI", "AmberSPI")
    neg_red(ws.Range("M11"))

    ws.Range("B13").Value = "This week's notes"
    ws.Range("B13").Font.Color = rgb(GOLD); ws.Range("B13").Font.Bold = True
    ws.Range("B13").VerticalAlignment = -4160
    ws.Range("C13:Q15").Merge()
    ws.Range("C13").Value = proj["This week's note"] if proj else ""
    ws.Range("C13:Q15").WrapText = True
    ws.Range("C13:Q15").VerticalAlignment = -4160

    hdr = ["Cost code", "Description", "Category", "Budget", "Committed", "Actual to date", "Planned % to date", "% complete",
           "Planned value", "Earned value", "Cost variance", "Schedule variance", "CPI", "SPI", "Forecast at completion",
           "Variance at completion"]
    ws.Range("B17:Q17").Value = hdr
    table_head(ws.Range("B17:Q17"))
    ws.Rows(17).RowHeight = 32
    vals = {r["code"]: r for r in (lines or [])}
    for i, code in enumerate(codes):
        r = FIRST + i
        ws.Range(f"B{r}").NumberFormat = "@"
        ws.Range(f"B{r}").Value = code
        ws.Range(f"C{r}").Formula = f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},CostCodes,2,FALSE),"Unknown code"))'
        ws.Range(f"D{r}").Formula = f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},CostCodes,3,FALSE),"Unknown"))'
        v = vals.get(code)
        if v:
            ws.Range(f"E{r}:G{r}").Value = (float(v["budget"]), float(v["committed"]), float(v["actual"]))
            ws.Range(f"I{r}").Value = float(v["pct"])
        ws.Range(f"H{r}").Formula = (f'=IF(B{r}="","",IFERROR(SUMIFS(Baseline[Planned %],Baseline[Sheet],$C$6,Baseline[Cost code],B{r},'
                                     f'Baseline[Week ending],MIN(WeekEnding,MAX(Baseline[Week ending]))),0))')
        ws.Range(f"J{r}").Formula = f"=E{r}*H{r}"
        ws.Range(f"K{r}").Formula = f"=E{r}*I{r}"
        ws.Range(f"L{r}").Formula = f"=K{r}-G{r}"
        ws.Range(f"M{r}").Formula = f"=K{r}-J{r}"
        ws.Range(f"N{r}").Formula = f'=IF(G{r}>0,K{r}/G{r},"")'
        ws.Range(f"O{r}").Formula = f'=IF(J{r}>0,K{r}/J{r},"")'
        ws.Range(f"P{r}").Formula = f"=IF(I{r}>0,G{r}/I{r},MAX(E{r},G{r}))"
        ws.Range(f"Q{r}").Formula = f"=E{r}-P{r}"
    table_body(ws.Range(f"B{FIRST}:Q{LAST}"))
    ws.Range(f"B{TOT}").Value = "Total"
    for c in "EFGJKLMPQ":
        ws.Range(f"{c}{TOT}").Formula = f"=SUM({c}{FIRST}:{c}{LAST})"
    ws.Range(f"N{TOT}").Formula = f'=IF(G{TOT}>0,K{TOT}/G{TOT},"")'
    ws.Range(f"O{TOT}").Formula = f'=IF(J{TOT}>0,K{TOT}/J{TOT},"")'
    table_total(ws.Range(f"B{TOT}:Q{TOT}"))
    for c in "EFGJKP":
        ws.Range(f"{c}{FIRST}:{c}{TOT}").NumberFormat = "#,##0"
    for c in "LMQ":
        ws.Range(f"{c}{FIRST}:{c}{TOT}").NumberFormat = MONEY
        neg_red(ws.Range(f"{c}{FIRST}:{c}{TOT}"))
    ws.Range(f"H{FIRST}:I{LAST}").NumberFormat = "0.0%"
    ws.Range(f"N{FIRST}:O{TOT}").NumberFormat = "0.00"
    ws.Range(f"N{FIRST}:O{TOT}").HorizontalAlignment = -4108
    rag_index(ws.Range(f"N{FIRST}:N{TOT}"), f"N{FIRST}", "GreenCPI", "AmberCPI")
    rag_index(ws.Range(f"O{FIRST}:O{TOT}"), f"O{FIRST}", "GreenSPI", "AmberSPI")
    inputs = [ws.Range(f"B{FIRST}:B{LAST}"), ws.Range(f"E{FIRST}:G{LAST}"), ws.Range(f"I{FIRST}:I{LAST}"), ws.Range("C6:C9"), ws.Range("I6:J10"),
              ws.Range("C13:Q15")]
    for rng in inputs:
        rng.Interior.Color = rgb(INPUT_BG)
        rng.Font.Color = rgb(GOLD_LT)
    dv = ws.Range(f"B{FIRST}:B{LAST}").Validation
    dv.Delete()
    dv.Add(3, 1, 1, "=CodeList")
    dv.ErrorMessage = "Pick a cost code from the list (Cost codes sheet)."
    for c in "I":
        dv = ws.Range(f"{c}{FIRST}:{c}{LAST}").Validation
        dv.Delete()
        dv.Add(2, 1, 1, "0", "1")  # decimal between 0 and 1
        dv.ErrorMessage = "% is between 0% and 100%."
    ws.Range("B30").Value = ("Planned value = budget x planned % (looked up in the Baseline schedule for this week); earned value = budget x % complete. "
                             "CPI = earned / actual, SPI = earned / planned (below 1.00 = over cost / behind schedule). "
                             "Forecast at completion = actual / % complete (budget if the line has not started).")
    ws.Range("B30").Font.Italic = True; ws.Range("B30").Font.Color = rgb(MUTED); ws.Range("B30").Font.Size = 8.5
    ws.Cells.Locked = True
    for rng in inputs:
        for r in range(1, rng.Rows.Count + 1):  # row by row: a range spanning several merged areas refuses
            row = rng.Rows(r)
            (row.Cells(1, 1).MergeArea if row.Cells(1, 1).MergeCells else row).Locked = False
    ws.Activate()
    ws.Range(f"A{FIRST}").Select()
    ws.Application.ActiveWindow.FreezePanes = True
    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$Q$30"


def protect(ws):
    ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True, AllowFormattingColumns=True, AllowFormattingRows=True)


def bookend(ws, text):
    ws.Tab.Color = rgb((166, 166, 166))
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 110
    ws.Range("B2").Value = text
    ws.Range("B2").Font.Size = 14; ws.Range("B2").Font.Bold = True
    lines = ["Every sheet between 'PM start' and 'PM end' is a project and is added up on the Dashboard (3D sums).",
             "Add a project: right-click 'Template' > Move or Copy > tick Create a copy > put it before 'PM end'.",
             "Rename the copy to the project ID (e.g. P-2404), fill in the gold cells, then type the same ID in the next empty "
             "row of the Projects table on the Dashboard.",
             "If a sheet sits between the bookends but is not on the Dashboard, the Dashboard check says FAIL."]
    for i, t in enumerate(lines, 4):
        ws.Range(f"B{i}").Value = t


# ---------------------------------------------------------------- Rollup: every number the report sheets show
R_HEAD, R_FIRST = 4, 5
R_LAST = R_FIRST + SLOTS - 1
R_TOT = R_LAST + 1
R_PICK, R_CURVE = R_TOT + 2, R_TOT + 3
D_FIRST = 13  # first project row on the Dashboard


def ind(cell):
    return f'INDIRECT("\'"&{{sheet}}&"\'!{cell}")'


ROLL = [  # key, header, number format, expression ({key} = that column, same row), total row: "sum" / "same" / None
    ("sheet", "Sheet", "@", None, None),
    ("project", "Project", None, ind(PMC["name"]), None),
    ("pm", "PM", None, ind(PMC["pm"]), None),
    ("contract", "Contract", MONEY, ind(PMC["contract"]), "sum"),
    ("co_price", "Approved CO price", MONEY, 'SUMIFS(ChangeOrders[Price],ChangeOrders[Sheet],{sheet},ChangeOrders[Status],"Approved")', "sum"),
    ("co_cost", "Approved CO cost", MONEY, 'SUMIFS(ChangeOrders[Cost estimate],ChangeOrders[Sheet],{sheet},ChangeOrders[Status],"Approved")', "sum"),
    ("co_pending", "Pending CO price", MONEY, 'SUMIFS(ChangeOrders[Price],ChangeOrders[Sheet],{sheet},ChangeOrders[Status],"Pending")', "sum"),
    ("rev_contract", "Revised contract", MONEY, "{contract}+{co_price}", "sum"),
    ("budget", "Budget (BAC)", MONEY, ind(f"E{TOT}"), "sum"),
    ("rev_budget", "Revised budget", MONEY, "{budget}+{co_cost}", "sum"),
    ("committed", "Committed", MONEY, ind(f"F{TOT}"), "sum"),
    ("actual", "Actual cost (AC)", MONEY, ind(f"G{TOT}"), "sum"),
    ("earned", "Earned value (EV)", MONEY, ind(f"K{TOT}"), "sum"),
    ("planned", "Planned value (PV)", MONEY, ind(f"J{TOT}"), "sum"),
    ("cv", "Cost variance", MONEY, "{earned}-{actual}", "sum"),
    ("sv", "Schedule variance", MONEY, "{earned}-{planned}", "sum"),
    ("cpi", "CPI", "0.00", 'IF(N({actual})>0,{earned}/{actual},"")', "same"),
    ("spi", "SPI", "0.00", 'IF(N({planned})>0,{earned}/{planned},"")', "same"),
    ("field", "Field %", "0.0%", ind(PMC["field"]), None),
    ("shop", "Shop %", "0.0%", ind(PMC["shop"]), None),
    ("fac", "Forecast (cost codes)", MONEY, ind(f"P{TOT}"), "sum"),
    ("fac_co", "Forecast incl. COs", MONEY, "{fac}+{co_cost}", "sum"),
    ("eac_cpi", "EAC CPI method", MONEY, 'IF(N({cpi})>0,{budget}/{cpi},"")', "same"),
    ("eac_cpi_spi", "EAC CPI x SPI method", MONEY, 'IF(N({cpi})*N({spi})>0,{actual}+({budget}-{earned})/({cpi}*{spi}),"")', "same"),
    ("tcpi", "TCPI", "0.00", 'IF({budget}-{actual}<>0,({budget}-{earned})/({budget}-{actual}),"")', "same"),
    ("etc", "Estimate to complete", MONEY, "{fac_co}-{actual}", "sum"),
    ("vac", "Variance at completion", MONEY, "{rev_budget}-{fac_co}", "sum"),
    ("margin", "Margin at completion", "0.0%", 'IF(N({rev_contract})>0,({rev_contract}-{fac_co})/{rev_contract},"")', "same"),
    ("bid_margin", "Bid margin", "0.0%", 'IF(N({contract})>0,({contract}-{budget})/{contract},"")', "same"),
    ("cost_status", "Cost status", None, 'IF({cpi}="","",IF({cpi}>=GreenCPI,"On track",IF({cpi}>=AmberCPI,"Watch","At risk")))', "same"),
    ("sched_status", "Schedule status", None, 'IF({spi}="","",IF({spi}>=GreenSPI,"On track",IF({spi}>=AmberSPI,"Watch","At risk")))', "same"),
    ("pct_cost", "% complete (cost-to-cost)", "0.0%", 'IF(N({fac_co})>0,{actual}/{fac_co},"")', "same"),
    ("earned_rev", "Earned revenue", MONEY, 'IF({pct_cost}="","",{rev_contract}*{pct_cost})', "sum"),
    ("billed", "Billed to date", MONEY, ind(PMC["billed"]), "sum"),
    ("over_under", "Over / (under) billed", MONEY, "{billed}-{earned_rev}", "sum"),
    ("ret_pct", "Retainage %", "0%", ind(PMC["ret"]), None),
    ("retainage", "Retainage held", MONEY, "{billed}*{ret_pct}", "sum"),
    ("received", "Cash received", MONEY, ind(PMC["recv"]), "sum"),
    ("paid", "Cash paid", MONEY, ind(PMC["paid"]), "sum"),
    ("net_cash", "Net cash", MONEY, "{received}-{paid}", "sum"),
    ("ar", "Receivable (excl. retainage)", MONEY, "{billed}-{retainage}-{received}", "sum"),
    ("heads", "Crew on site", "0", ind(PMC["heads"]), "sum"),
    ("pheads", "Planned crew", "0", ind(PMC["pheads"]), "sum"),
    ("baseline_rows", "Baseline rows this week", "0", "COUNTIFS(Baseline[Sheet],{sheet},Baseline[Week ending],MIN(WeekEnding,MAX(Baseline[Week ending])))", "sum"),
    ("no_baseline", "No baseline this week", "0", "--({baseline_rows}=0)", "sum"),
    ("over_contract", "Billed above contract", "0", "--({billed}>{rev_contract})", "sum"),
    ("risk_exp", "Risk exposure (expected)", MONEY, "SUMIFS(Risks[Expected cost],Risks[Sheet],{sheet})", "sum"),
    ("risk_high", "High risks (score 15+)", "0", 'COUNTIFS(Risks[Sheet],{sheet},Risks[Score],">=15")', "sum"),
    ("risk_open", "Open risks", "0", "COUNTIFS(Risks[Sheet],{sheet})", "sum"),
    ("co_app_n", "Approved COs", "0", 'COUNTIFS(ChangeOrders[Sheet],{sheet},ChangeOrders[Status],"Approved")', "sum"),
    ("co_pend_n", "Pending COs", "0", 'COUNTIFS(ChangeOrders[Sheet],{sheet},ChangeOrders[Status],"Pending")', "sum"),
    ("note", "This week's note", None, ind(PMC["note"]), None),
]
RC = {k: col(i + 1) for i, (k, *_rest) in enumerate(ROLL)}


def R(key, row=None):
    """Absolute reference to a Rollup cell: R('cpi') = total row, R('cpi', 'pick') = the Project view pick."""
    r = {None: R_TOT, "pick": R_PICK, "curve": R_CURVE}.get(row, row)
    return f"Rollup!${RC[key]}${r}"


def build_rollup(wb):
    ws = wb.Worksheets("Rollup")
    ws.Tab.Color = rgb((166, 166, 166))
    ws.Range("A1").Value = "Rollup: every figure the report sheets show, one row per project (read from the PM sheets)"
    ws.Range("A1").Font.Size = 14; ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Calculated, do not type here. Row 'Project view' and 'S-curve' = the project picked on those sheets."
    ws.Range("A2").Font.Italic = True
    for j, (k, h, fmt, expr, agg) in enumerate(ROLL, 1):
        ws.Cells(R_HEAD, j).Value = h
        for r in range(R_FIRST, R_LAST + 1):
            cells = {kk: f"{RC[kk]}{r}" for kk in RC}
            cells["sheet"] = f"$A{r}"
            if k == "sheet":
                d = D_FIRST + r - R_FIRST
                f = f'=IF(Dashboard!$B{d}="","",Dashboard!$B{d})'
            else:
                f = '=IF($A{r}="","",{e})'.format(r=r, e=expr.format(**cells))
            ws.Cells(r, j).Formula = f
        cells = {kk: f"{RC[kk]}{R_TOT}" for kk in RC}
        if k == "sheet":
            ws.Cells(R_TOT, j).Value = "Total"
        elif agg == "sum":
            ws.Cells(R_TOT, j).Formula = f"=SUM({RC[k]}{R_FIRST}:{RC[k]}{R_LAST})"
        elif agg == "same":
            ws.Cells(R_TOT, j).Formula = "=" + expr.format(**cells)
        for r, label in ((R_PICK, "Project view"), (R_CURVE, "S-curve")):
            if k == "sheet":
                ws.Cells(r, j).Formula = "=PickProject" if label == "Project view" else "=PickCurve"
            else:
                ws.Cells(r, j).Formula = (f'=IFERROR(INDEX({RC[k]}${R_FIRST}:{RC[k]}${R_LAST},'
                                          f'MATCH($A{r},$A${R_FIRST}:$A${R_LAST},0)),"")')
        if fmt:
            ws.Range(f"{RC[k]}{R_FIRST}:{RC[k]}{R_CURVE}").NumberFormat = fmt
    hdr = ws.Range(ws.Cells(R_HEAD, 1), ws.Cells(R_HEAD, len(ROLL)))
    hdr.Font.Bold = True; hdr.Interior.Color = rgb(LIGHT_HEAD); hdr.Font.Color = rgb((255, 255, 255)); hdr.WrapText = True
    ws.Rows(R_HEAD).RowHeight = 45
    ws.Range(ws.Cells(R_TOT, 1), ws.Cells(R_TOT, len(ROLL))).Font.Bold = True
    ws.Range(ws.Cells(R_PICK, 1), ws.Cells(R_CURVE, len(ROLL))).Interior.Color = rgb((242, 246, 250))
    ws.Columns("A").ColumnWidth = 12
    ws.Range(ws.Cells(1, 2), ws.Cells(1, len(ROLL))).EntireColumn.ColumnWidth = 12
    ws.Columns(RC["note"]).ColumnWidth = 40
    return ws


# gauges: rows on the Rollup sheet, below the table
G_ROW = R_CURVE + 4
GAUGES = [("Dashboard CPI", "cpi", None, "GreenCPI", "AmberCPI"), ("Dashboard SPI", "spi", None, "GreenSPI", "AmberSPI"),
          ("Project view CPI", "cpi", "pick", "GreenCPI", "AmberCPI"), ("Project view SPI", "spi", "pick", "GreenSPI", "AmberSPI")]
G_MIN, G_MAX = 0.70, 1.30


def build_gauge_cells(ws):
    ws.Range(f"A{G_ROW - 1}").Value = f"Gauges (scale {G_MIN:.2f} - {G_MAX:.2f}): value, share of the arc, green, amber, red, rest, hidden half"
    ws.Range(f"A{G_ROW - 1}").Font.Bold = True
    refs = {}
    for i, (label, key, row, g, a) in enumerate(GAUGES):
        r = G_ROW + i
        ws.Range(f"A{r}").Value = label
        ws.Range(f"B{r}").Formula = f"={R(key, row)}"
        ws.Range(f"C{r}").Formula = f'=IF(ISNUMBER(B{r}),MEDIAN(0,(B{r}-{G_MIN})/({G_MAX}-{G_MIN}),1),0)'
        ws.Range(f"D{r}").Formula = f'=IF(AND(ISNUMBER(B{r}),B{r}>={g}),C{r},0)'
        ws.Range(f"E{r}").Formula = f'=IF(AND(ISNUMBER(B{r}),B{r}<{g},B{r}>={a}),C{r},0)'
        ws.Range(f"F{r}").Formula = f'=IF(AND(ISNUMBER(B{r}),B{r}<{a}),C{r},0)'
        ws.Range(f"G{r}").Formula = f"=1-C{r}"
        ws.Range(f"H{r}").Value = 1
        refs[label] = ws.Range(f"D{r}:H{r}")
    return refs


# ---------------------------------------------------------------- dashboard
def grid(ws, last="S", width=9.4):
    widths = {"A": 2}
    for i in range(2, ord(last) - 63):
        widths[col(i)] = width
    return widths


def build_dashboard(wb, sheets, T, gauges):
    ws = wb.Worksheets("Dashboard")
    dark(ws, grid(ws))
    header(ws, "S", "Project cost dashboard",
           '=Company&" (fictional)  ·  week ending "&TEXT(WeekEnding,"d mmm yyyy")&"  ·  "&COUNTA($B$13:$B$17)&" active projects  ·  USD"',
           "Dashboard", sheets[0])
    # KPI cards: rows 6-9
    for r, h in ((6, 22), (7, 22), (8, 22), (9, 22), (10, 12)):
        ws.Rows(r).RowHeight = h
    pos = cards_row(ws, "B", "S", 6, 6, height=86)
    t_bud = f'TEXT({R("rev_budget")},"#,##0")'
    card(ws, *pos[0], "Revised contract", T.add("Dash revised contract", f"={R('rev_contract')}", "#,##0"),
         [(T.add("Dash contract note", f'="+"&TEXT({R("co_price")}/1000,"#,##0")&"k approved COs  ·  "&TEXT({R("co_pending")}/1000,"#,##0")&"k pending"'), MUTED)],
         icon="$")
    card(ws, *pos[1], "Forecast at completion", T.add("Dash forecast", f"={R('fac_co')}", "#,##0"),
         note3(T, "Dash forecast note", f'IF({R("vac")}>=0,"▲ "&TEXT({R("vac")},"#,##0")&" under budget","▼ "&TEXT(-{R("vac")},"#,##0")&" over the budget")',
               f'{R("vac")}>=0', f'{R("vac")}<0'), icon="◎")
    card(ws, *pos[2], "Margin at completion", T.add("Dash margin", f"={R('margin')}", "0.0%"),
         note3(T, "Dash margin note", f'"bid margin "&TEXT({R("bid_margin")},"0.0%")&"  ·  "&TEXT({R("margin")}-{R("bid_margin")},"+0.0%;-0.0%")',
               f'{R("margin")}>={R("bid_margin")}', f'{R("margin")}<{R("bid_margin")}-0.01'), icon="%")
    for i, (key, label, green, amber, gname) in enumerate((("cpi", "Cost performance (CPI)", "GreenCPI", "AmberCPI", "Dashboard CPI"),
                                                            ("spi", "Schedule performance (SPI)", "GreenSPI", "AmberSPI", "Dashboard SPI"))):
        x, y, w, h = pos[3 + i]
        other = "AC" if key == "cpi" else "PV"
        okey = "actual" if key == "cpi" else "planned"
        card(ws, x, y, w, h, label, T.add(f"Dash {key}", f"={R(key)}", "0.00"),
             note3(T, f"Dash {key} note", f'{R("cost_status" if key == "cpi" else "sched_status")}&"  ·  EV "&TEXT({R("earned")}/1000,"#,##0")&"k vs {other} "&TEXT({R(okey)}/1000,"#,##0")&"k"',
                   f"{R(key)}>={green}", f"{R(key)}<{amber}"), value_w=60)
        gauge(ws, x + w - 84, y + 6, 80, gauges[gname])
    card(ws, *pos[5], "Over / (under) billed", T.add("Dash over under", f"={R('over_under')}", "#,##0;(#,##0)"),
         note3(T, "Dash billing note", f'"net cash "&TEXT({R("net_cash")}/1000,"#,##0;-#,##0")&"k  ·  retainage "&TEXT({R("retainage")}/1000,"#,##0")&"k"',
               f'{R("net_cash")}>=0', f'{R("net_cash")}<0'), icon="⇄")

    # Projects table
    section(ws, "B11", "Projects")
    ws.Range("F11").Value = "One row per PM sheet: type the sheet name in a gold cell. Cost / Schedule = CPI / SPI against the Settings thresholds."
    ws.Range("F11").Font.Color = rgb(MUTED); ws.Range("F11").Font.Size = 8.5; ws.Range("F11").VerticalAlignment = -4107
    ws.Rows(11).RowHeight = 24
    hdr = {"B": "Sheet", "C": "Project", "F": "Revised contract", "G": "Budget", "H": "Actual to date", "I": "Earned value",
           "J": "CPI", "K": "SPI", "L": "Field %", "M": "Shop fab %", "N": "Forecast incl. COs", "O": "TCPI",
           "P": "Margin at completion", "Q": "Over/under billed", "R": "Cost", "S": "Schedule"}
    for c, h in hdr.items():
        ws.Range(f"{c}12").Value = h
    ws.Range("C12:E12").Merge()
    table_head(ws.Range("B12:S12"))
    ws.Rows(12).RowHeight = 30
    src = {"F": "rev_contract", "G": "budget", "H": "actual", "I": "earned", "J": "cpi", "K": "spi", "L": "field", "M": "shop",
           "N": "fac_co", "O": "tcpi", "P": "margin", "Q": "over_under", "R": "cost_status", "S": "sched_status"}
    d_last = D_FIRST + SLOTS - 1
    t = d_last + 1
    for i in range(SLOTS):
        r = D_FIRST + i
        rr = R_FIRST + i
        ws.Range(f"B{r}").NumberFormat = "@"
        if i < len(sheets):
            ws.Range(f"B{r}").Value = sheets[i]
        ws.Range(f"C{r}:E{r}").Merge()
        ws.Range(f"C{r}").Formula = f'=Rollup!{RC["project"]}{rr}'
        for c, k in src.items():
            ws.Range(f"{c}{r}").Formula = f"=Rollup!{RC[k]}{rr}"
        ws.Rows(r).RowHeight = 20
    ws.Range(f"B{t}").Value = "Total"
    ws.Range(f"C{t}:E{t}").Merge()
    ws.Range(f"C{t}").Formula = f'=COUNTA(B{D_FIRST}:B{d_last})&" projects"'
    for c, k in src.items():
        if k not in ("field", "shop"):
            ws.Range(f"{c}{t}").Formula = f"={R(k)}"
    ws.Rows(t).RowHeight = 22
    table_body(ws.Range(f"B{D_FIRST}:S{d_last}"))
    table_total(ws.Range(f"B{t}:S{t}"))
    ws.Range(f"B{D_FIRST}:B{d_last}").Interior.Color = rgb(INPUT_BG)
    ws.Range(f"B{D_FIRST}:B{d_last}").Font.Color = rgb(GOLD_LT)
    ws.Range(f"B{D_FIRST}:B{d_last}").Font.Bold = True
    for c in "FGHIN":
        ws.Range(f"{c}{D_FIRST}:{c}{t}").NumberFormat = "#,##0"
    ws.Range(f"Q{D_FIRST}:Q{t}").NumberFormat = '#,##0;(#,##0);"–"'
    neg_red(ws.Range(f"Q{D_FIRST}:Q{t}"))
    ws.Range(f"J{D_FIRST}:K{t}").NumberFormat = "0.00"
    ws.Range(f"O{D_FIRST}:O{t}").NumberFormat = "0.00"
    ws.Range(f"L{D_FIRST}:M{t}").NumberFormat = "0%"
    ws.Range(f"P{D_FIRST}:P{t}").NumberFormat = "0.0%;-0.0%"
    neg_red(ws.Range(f"P{D_FIRST}:P{t}"))
    ws.Range(f"J{D_FIRST}:M{t}").HorizontalAlignment = -4108
    ws.Range(f"O{D_FIRST}:O{t}").HorizontalAlignment = -4108
    ws.Range(f"R{D_FIRST}:S{t}").HorizontalAlignment = -4108
    rag_index(ws.Range(f"J{D_FIRST}:J{t}"), f"J{D_FIRST}", "GreenCPI", "AmberCPI")
    rag_index(ws.Range(f"K{D_FIRST}:K{t}"), f"K{D_FIRST}", "GreenSPI", "AmberSPI")
    cf(ws.Range(f"O{D_FIRST}:O{t}"), f"=AND(ISNUMBER(O{D_FIRST}),O{D_FIRST}>1.05)", font=BAD, bold=True)
    pills(ws.Range(f"R{D_FIRST}:S{t}"))
    cf(ws.Range(f"M{D_FIRST}:M{d_last}"), f"=AND(ISNUMBER(M{D_FIRST}),L{D_FIRST}-M{D_FIRST}>=0.1)", font=AMBER, fill=AMBER_BG, bold=True)
    ws.Range(f"B{t + 1}").Value = ("CPI = earned / actual, SPI = earned / planned, TCPI = CPI needed on the remaining work to finish on budget "
                                   "(red above 1.05). Forecast = actual / % complete per cost code + approved change orders. "
                                   "Amber shop % = shop fabrication 10+ points behind the field.")
    ws.Range(f"B{t + 1}").Font.Italic = True; ws.Range(f"B{t + 1}").Font.Color = rgb(MUTED); ws.Range(f"B{t + 1}").Font.Size = 8
    ws.Rows(t + 2).RowHeight = 10

    # heatmap: project x cost code
    h0 = t + 3
    section(ws, f"B{h0}", "Cost-code heatmap")
    ws.Range(f"F{h0}").Value = "Each cell = the index of one cost code on one PM sheet; red / amber / green against the Settings thresholds."
    ws.Range(f"F{h0}").Font.Color = rgb(MUTED); ws.Range(f"F{h0}").Font.Size = 8.5; ws.Range(f"F{h0}").VerticalAlignment = -4107
    ws.Range(f"Q{h0}").Value = "Show"
    ws.Range(f"Q{h0}").HorizontalAlignment = -4152; ws.Range(f"Q{h0}").Font.Color = rgb(MUTED)
    ws.Range(f"Q{h0}").VerticalAlignment = -4108
    ws.Range(f"R{h0}:S{h0}").Merge()
    m = ws.Range(f"R{h0}")
    m.Value = "CPI"
    m.Interior.Color = rgb(INPUT_BG); m.Font.Color = rgb(GOLD_LT); m.Font.Bold = True; m.HorizontalAlignment = -4108
    m.VerticalAlignment = -4108
    dv = m.Validation
    dv.Delete()
    dv.Add(3, 1, 1, "=MetricList")
    wb.Names.Add("HeatMetric", f"=Dashboard!$R${h0}")
    ws.Rows(h0).RowHeight = 24
    T.add("Heat column", '=IF(HeatMetric="SPI","O","N")', name="HeatCol")
    T.add("Heat green", '=IF(HeatMetric="SPI",GreenSPI,GreenCPI)', name="HeatGreen")
    T.add("Heat amber", '=IF(HeatMetric="SPI",AmberSPI,AmberCPI)', name="HeatAmber")
    hc, hd = h0 + 1, h0 + 2
    ws.Range(f"B{hc}:D{hd}").Merge()
    ws.Range(f"B{hc}").Formula = '="Project  ·  "&HeatMetric&" by cost code"'
    codes_cols = [col(5 + k) for k in range(10)]  # E..N
    for k, c in enumerate(codes_cols, 1):
        ws.Range(f"{c}{hc}").Formula = f"=INDEX(CostCodes[Cost code],{k})"
        ws.Range(f"{c}{hd}").Formula = f'=IFERROR(MID(INDEX(CostCodes[Description],{k}),FIND(" - ",INDEX(CostCodes[Description],{k}))+3,40),"")'
    ws.Range(f"O{hc}:S{hd}").Merge()
    ws.Range(f"O{hc}").Value = "Weakest cost code"
    table_head(ws.Range(f"B{hc}:S{hd}"))
    ws.Range(f"B{hc}:S{hc}").Borders(9).LineStyle = -4142
    ws.Range(f"E{hd}:N{hd}").Font.Bold = False; ws.Range(f"E{hd}:N{hd}").Font.Color = rgb(MUTED); ws.Range(f"E{hd}:N{hd}").Font.Size = 7.5
    ws.Range(f"E{hc}:N{hd}").HorizontalAlignment = -4108
    ws.Rows(hc).RowHeight = 15; ws.Rows(hd).RowHeight = 30
    h1 = hd + 1
    for i in range(SLOTS):
        r = h1 + i
        d = D_FIRST + i
        ws.Range(f"B{r}:D{r}").Merge()
        ws.Range(f"B{r}").Formula = (f'=IF($B${d}="","",$B${d}&"  "&IF(LEN($C${d})>24,LEFT($C${d},23)&"…",$C${d}))')
        for c in codes_cols:
            ws.Range(f"{c}{r}").Formula = (
                f'=IF($B${d}="","",IFERROR(INDEX(INDIRECT("\'"&$B${d}&"\'!"&HeatCol&"{FIRST}:"&HeatCol&"{LAST}"),'
                f'MATCH({c}${hc},INDIRECT("\'"&$B${d}&"\'!B{FIRST}:B{LAST}"),0)),""))')
        ws.Range(f"O{r}:S{r}").Merge()
        ws.Range(f"O{r}").Formula = (f'=IF(OR($B${d}="",COUNT(E{r}:N{r})=0),"",INDEX($E${hc}:$N${hc},MATCH(MIN(E{r}:N{r}),E{r}:N{r},0))'
                                     f'&"  "&INDEX($E${hd}:$N${hd},MATCH(MIN(E{r}:N{r}),E{r}:N{r},0))&"  "&TEXT(MIN(E{r}:N{r}),"0.00"))')
        ws.Rows(r).RowHeight = 20
    h_last = h1 + SLOTS - 1
    table_body(ws.Range(f"B{h1}:S{h_last}"))
    heat = ws.Range(f"E{h1}:N{h_last}")
    heat.NumberFormat = "0.00"
    heat.HorizontalAlignment = -4108
    heat.Borders(11).LineStyle = 1; heat.Borders(11).Color = rgb(PAGE); heat.Borders(11).Weight = -4138
    heat.Borders(12).LineStyle = 1; heat.Borders(12).Color = rgb(PAGE); heat.Borders(12).Weight = -4138
    a = f"E{h1}"
    cf(heat, f"=AND(ISNUMBER({a}),{a}<HeatAmber)", font=BAD, fill=BAD_BG, bold=True)
    cf(heat, f"=AND(ISNUMBER({a}),{a}>=HeatAmber,{a}<HeatGreen)", font=AMBER, fill=AMBER_BG, bold=True)
    cf(heat, f"=AND(ISNUMBER({a}),{a}>=HeatGreen)", font=GOOD, fill=GOOD_BG)
    ws.Range(f"O{h1}:S{h_last}").Font.Color = rgb(GOLD_LT)
    ws.Range(f"O{h1}:S{h_last}").Font.Size = 9
    ws.Range(f"B{h_last + 1}").Value = "Blank = cost code not started yet (no actual cost or no planned value). Switch CPI / SPI in the gold cell above."
    ws.Range(f"B{h_last + 1}").Font.Italic = True; ws.Range(f"B{h_last + 1}").Font.Color = rgb(MUTED); ws.Range(f"B{h_last + 1}").Font.Size = 8
    ws.Rows(h_last + 2).RowHeight = 10

    # charts
    c0 = h_last + 3
    c1 = c0 + 15
    for r in range(c0, c1 + 1):
        ws.Rows(r).RowHeight = 15
    sep = SEP
    cnt = f"COUNTA(Dashboard!$B${D_FIRST}:$B${d_last})"
    for nm, cc in (("PrjIds", "B"), ("PrjBudget", "G"), ("PrjActual", "H"), ("PrjFAC", "N"), ("PrjField", "L"), ("PrjShop", "M")):
        n = wb.Names.Add(nm, f"=Dashboard!${cc}${D_FIRST}")
        n.RefersToLocal = f"=OFFSET(Dashboard!${cc}${D_FIRST},0,0,MAX(1,{cnt}),1)".replace(",", sep)
    book = wb.Name
    ch = add_chart(ws, f"B{c0}:J{c1}", 51)
    for name, ref, color in (("Budget", "PrjBudget", GREY), ("Actual to date", "PrjActual", COPPER), ("Forecast incl. COs", "PrjFAC", GOLD)):
        s = ch.SeriesCollection().NewSeries()
        s.Formula = f"=SERIES(\"{name}\",'{book}'!PrjIds,'{book}'!{ref},{ch.SeriesCollection().Count})"
        s.Format.Fill.ForeColor.RGB = rgb(color)
        s.Format.Line.Visible = False
    dark_chart(ch, "Budget, actual and forecast by project")
    ch.Axes(2).TickLabels.NumberFormat = '#,##0.0,,"M";;0'
    ch.ChartGroups(1).GapWidth = 70
    ch.ChartGroups(1).Overlap = -5
    ch = add_chart(ws, f"K{c0}:S{c1}", 57)
    for name, ref, color in (("Field % complete", "PrjField", JADE), ("Shop fabrication % complete", "PrjShop", COPPER)):
        s = ch.SeriesCollection().NewSeries()
        s.Formula = f"=SERIES(\"{name}\",'{book}'!PrjIds,'{book}'!{ref},{ch.SeriesCollection().Count})"
        s.Format.Fill.ForeColor.RGB = rgb(color)
        s.Format.Line.Visible = False
        s.HasDataLabels = True
        s.DataLabels().NumberFormat = "0%"
        s.DataLabels().Font.Color = rgb(IVORY)
        s.DataLabels().Font.Size = 8.5
    dark_chart(ch, "Field vs shop fabrication progress")
    ch.Axes(1).ReversePlotOrder = True
    ch.Axes(2).MinimumScale = 0
    ch.Axes(2).MaximumScale = 1
    ch.Axes(2).TickLabels.NumberFormat = "0%"
    ch.ChartGroups(1).GapWidth = 60

    # checks + notes
    k0 = c1 + 2
    ws.Rows(c1 + 1).RowHeight = 10
    section(ws, f"B{k0}", "Checks")
    section(ws, f"K{k0}", "This week from the PMs")
    ws.Rows(k0).RowHeight = 24
    checks = [
        ("Every PM sheet between 'PM start' and 'PM end' is on this dashboard",
         f"=IF(SHEETS('PM start:PM end'!A1)-2=COUNTA(B{D_FIRST}:B{d_last}),\"PASS\",\"FAIL\")"),
        ("Dashboard budget = sum of the budget on every PM sheet",
         f"=IF(ROUND(SUM('PM start:PM end'!E{TOT})-{R('budget')},2)=0,\"PASS\",\"FAIL\")"),
        ("Dashboard actual cost = sum of the actual cost on every PM sheet",
         f"=IF(ROUND(SUM('PM start:PM end'!G{TOT})-{R('actual')},2)=0,\"PASS\",\"FAIL\")"),
        ("No unknown cost codes on any PM sheet",
         f"=IF(SUM('PM start:PM end'!{PMC['unknown']})=0,\"PASS\",\"FAIL\")"),
        ("Planned % and % complete between 0% and 100% on every line",
         f"=IF(SUM('PM start:PM end'!{PMC['pctbad']})=0,\"PASS\",\"FAIL\")"),
        ("Every project has a baseline schedule for this week (planned value)",
         f"=IF({R('no_baseline')}=0,\"PASS\",\"FAIL\")"),
        ("Every change order belongs to a project here and has a valid status",
         f"=IF(SUMPRODUCT((COUNTIF(B{D_FIRST}:B{d_last},ChangeOrders[Sheet])=0)+(COUNTIF(StatusList,ChangeOrders[Status])=0))=0,\"PASS\",\"FAIL\")"),
        ("Risk probability and impact are whole numbers 1-5",
         "=IF(SUMPRODUCT(--ISNUMBER(MATCH(Risks[Probability],{1,2,3,4,5},0)),--ISNUMBER(MATCH(Risks[Impact],{1,2,3,4,5},0)))=ROWS(Risks[ID]),\"PASS\",\"FAIL\")"),
        ("No project billed above its revised contract",
         f"=IF({R('over_contract')}=0,\"PASS\",\"FAIL\")"),
        ("This week's snapshot saved, once per project",
         f"=IF(COUNTIFS(History[Week ending],WeekEnding)=COUNTA(B{D_FIRST}:B{d_last}),\"PASS\",\"FAIL\")"),
        ("History: no project saved twice for the same week",
         "=IF(SUMPRODUCT(1/COUNTIFS(History[Week ending],History[Week ending],History[Sheet],History[Sheet]))=ROWS(History[Sheet]),\"PASS\",\"FAIL\")"),
    ]
    late = []
    for i, (label, f) in enumerate(checks, k0 + 1):
        ws.Range(f"B{i}").Value = label
        ws.Range(f"B{i}").Font.Size = 9
        cell = ws.Range(f"J{i}")
        late.append((f"J{i}", f))
        cell.Font.Bold = True; cell.Font.Size = 8.5
        cell.HorizontalAlignment = -4108
        pass_fail(cell)
        ws.Rows(i).RowHeight = 18
    k_end = k0 + len(checks)
    table_body(ws.Range(f"B{k0 + 1}:J{k_end}"))
    wb.Names.Add("CheckResults", f"=Dashboard!$J${k0 + 1}:$J${k_end}")
    ws.Range(f"G{k0}:J{k0}").Merge()
    ws.Range(f"G{k0}").Formula = f'=COUNTIF(J{k0 + 1}:J{k_end},"PASS")&" / "&ROWS(J{k0 + 1}:J{k_end})&" checks PASS"'
    ws.Range(f"G{k0}").HorizontalAlignment = -4152; ws.Range(f"G{k0}").Font.Color = rgb(GOLD); ws.Range(f"G{k0}").Font.Bold = True
    ws.Range(f"G{k0}").VerticalAlignment = -4107
    ws.Range(f"B{k_end + 1}:I{k_end + 1}").Merge()
    ws.Range(f"B{k_end + 1}").Formula = ("=\"Lines where actual is above committed (check the POs): \"&SUM('PM start:PM end'!" + PMC["overcommit"] + ")")
    ws.Range(f"B{k_end + 1}").Font.Color = rgb(MUTED); ws.Range(f"B{k_end + 1}").Font.Size = 8.5
    for i in range(SLOTS):
        r = k0 + 1 + 2 * i
        d = D_FIRST + i
        rr = R_FIRST + i
        ws.Range(f"K{r}:K{r + 1}").Merge()
        ws.Range(f"K{r}").Formula = f'=IF($B${d}="","",$B${d})'
        ws.Range(f"K{r}").Font.Bold = True; ws.Range(f"K{r}").Font.Color = rgb(GOLD); ws.Range(f"K{r}").Font.Size = 9
        ws.Range(f"K{r}").VerticalAlignment = -4160
        ws.Range(f"L{r}:S{r + 1}").Merge()
        ws.Range(f"L{r}").Formula = f'=IF($B${d}="","",Rollup!{RC["note"]}{rr}&"  ("&Rollup!{RC["pm"]}{rr}&")")'
        ws.Range(f"L{r}:S{r + 1}").WrapText = True
        ws.Range(f"L{r}:S{r + 1}").VerticalAlignment = -4160
        ws.Range(f"L{r}").Font.Size = 9
        ws.Range(f"K{r}:S{r + 1}").Interior.Color = rgb(CARD)
        b = ws.Range(f"K{r + 1}:S{r + 1}").Borders(9); b.LineStyle = 1; b.Color = rgb(LINE)
    foot = k_end + 2
    ws.Range(f"B{foot}").Value = "Brightwire Electrical Contractors is fictional. No macros: Data > Refresh All saves the week. Excel 2016+ / Microsoft 365."
    ws.Range(f"B{foot}").Font.Color = rgb(MUTED); ws.Range(f"B{foot}").Font.Size = 8
    ps = ws.PageSetup
    ps.Orientation = 1; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = f"$A$1:$S${foot}"
    return dict(heat=(h1, h_last), heat_header=hc, charts=c0, checks=(k0 + 1, k_end), end=foot, late=late)


# ---------------------------------------------------------------- project view (drill-down)
def build_project_view(wb, sheets, T, gauges):
    ws = wb.Worksheets("Project view")
    dark(ws, grid(ws))
    header(ws, "S", '="Project view  ·  "&PickProject',
           "Pick a project in the gold cell: every card, the waterfall and the cost codes follow. Read from the PM sheet.",
           "Project view", sheets[0])
    ws.Rows(6).RowHeight = 26
    ws.Range("B6").Value = "Project"
    ws.Range("B6").Font.Color = rgb(MUTED); ws.Range("B6").VerticalAlignment = -4108
    ws.Range("C6:D6").Merge()
    p = ws.Range("C6")
    p.NumberFormat = "@"
    p.Value = "P-2402" if "P-2402" in sheets else sheets[0]
    p.Interior.Color = rgb(INPUT_BG); p.Font.Color = rgb(GOLD_LT); p.Font.Bold = True; p.Font.Size = 12
    p.HorizontalAlignment = -4108; p.VerticalAlignment = -4108
    b = ws.Range("C6:D6").Borders(9); b.LineStyle = 1; b.Color = rgb(GOLD)
    dv = p.Validation
    dv.Delete()
    dv.Add(3, 1, 1, "=PrjIds")
    ws.Range("E6:M6").Merge()
    ws.Range("E6").Formula = f'={R("project", "pick")}&"  ·  PM "&{R("pm", "pick")}'
    ws.Range("E6").Font.Name = SERIF; ws.Range("E6").Font.Size = 13; ws.Range("E6").VerticalAlignment = -4108
    ws.Range("E6").IndentLevel = 1
    for rng, f in (("N6:P6", f'="Cost: "&{R("cost_status", "pick")}'), ("Q6:S6", f'="Schedule: "&{R("sched_status", "pick")}')):
        ws.Range(rng).Merge()
        c = ws.Range(rng.split(":")[0])
        c.Formula = f
        c.HorizontalAlignment = -4108; c.VerticalAlignment = -4108; c.Font.Bold = True
        for word, (bg, fg) in (("On track", (GOOD_BG, GOOD)), ("Watch", (AMBER_BG, AMBER)), ("At risk", (BAD_BG, BAD))):
            cf(ws.Range(rng), f'=ISNUMBER(SEARCH("{word}",{rng.split(":")[0]}))', font=fg, fill=bg)
    ws.Rows(7).RowHeight = 10
    for r in range(8, 12):
        ws.Rows(r).RowHeight = 22
    pos = cards_row(ws, "B", "S", 8, 6, height=86)
    for i, (key, label, green, amber, gname) in enumerate((("cpi", "Cost performance (CPI)", "GreenCPI", "AmberCPI", "Project view CPI"),
                                                            ("spi", "Schedule performance (SPI)", "GreenSPI", "AmberSPI", "Project view SPI"))):
        x, y, w, h = pos[i]
        var = "cv" if key == "cpi" else "sv"
        card(ws, x, y, w, h, label, T.add(f"PV {key}", f"={R(key, 'pick')}", "0.00"),
             note3(T, f"PV {key} note", f'"{var.upper()} "&TEXT({R(var, "pick")},"#,##0;-#,##0")',
                   f"{R(key, 'pick')}>={green}", f"{R(key, 'pick')}<{amber}"), value_w=60)
        gauge(ws, x + w - 84, y + 6, 80, gauges[gname])
    card(ws, *pos[2], "Forecast incl. COs", T.add("PV forecast", f"={R('fac_co', 'pick')}", "#,##0"),
         [(T.add("PV forecast note", f'="CPI "&TEXT({R("eac_cpi", "pick")}/1000,"#,##0")&"k  ·  CPI×SPI "&TEXT({R("eac_cpi_spi", "pick")}/1000,"#,##0")&"k"'), MUTED)],
         icon="◎")
    card(ws, *pos[3], "TCPI to finish on budget", T.add("PV tcpi", f"={R('tcpi', 'pick')}", "0.00"),
         note3(T, "PV tcpi note", f'IF({R("tcpi", "pick")}>1,"rest must beat today\'s CPI","achievable at today\'s CPI")',
               f"{R('tcpi', 'pick')}<=1", f"{R('tcpi', 'pick')}>1.05"), icon="⚑")
    card(ws, *pos[4], "Margin at completion", T.add("PV margin", f"={R('margin', 'pick')}", "0.0%"),
         note3(T, "PV margin note", f'"bid margin "&TEXT({R("bid_margin", "pick")},"0.0%")',
               f'{R("margin", "pick")}>={R("bid_margin", "pick")}', f'{R("margin", "pick")}<{R("bid_margin", "pick")}-0.01'), icon="%")
    card(ws, *pos[5], "Over / (under) billed", T.add("PV over under", f"={R('over_under', 'pick')}", "#,##0;(#,##0)"),
         note3(T, "PV billing note", f'"net cash "&TEXT({R("net_cash", "pick")},"#,##0;-#,##0")',
               f'{R("net_cash", "pick")}>=0', f'{R("net_cash", "pick")}<0'), icon="⇄")
    ws.Rows(12).RowHeight = 12

    # waterfall data on the Rollup sheet
    ro = wb.Worksheets("Rollup")
    w0 = G_ROW + len(GAUGES) + 3
    ro.Range(f"A{w0 - 1}").Value = "Waterfall (Project view): label, value, invisible base, overrun, saving, total bars (all minus the floor in B" + str(w0 - 1) + ")"
    ro.Range(f"A{w0 - 1}").Font.Bold = True
    ro.Range(f"I{w0 - 1}").Value = "Floor"
    ro.Range(f"J{w0 - 1}").Formula = f'=IFERROR(ROUNDDOWN(MIN({R("budget", "pick")},{R("fac_co", "pick")})*0.92,-4),0)'
    floor = f"$J${w0 - 1}"
    rows = []
    r = w0
    ro.Range(f"A{r}").Formula = f'="Budget"&CHAR(10)&TEXT({R("budget", "pick")}/1000000,"0.00")&"M"'
    ro.Range(f"B{r}").Formula = f'=N({R("budget", "pick")})'
    ro.Range(f"G{r}").Formula = f"=B{r}"  # running level
    ro.Range(f"F{r}").Formula = f"=B{r}-{floor}"
    rows.append(r)
    for k in range(1, 11):
        r = w0 + k
        ro.Range(f"A{r}").Formula = f'=IFERROR(INDEX(INDIRECT("\'"&PickProject&"\'!B{FIRST}:B{LAST}"),{k}),"")'
        ro.Range(f"B{r}").Formula = (f'=IFERROR(INDEX(INDIRECT("\'"&PickProject&"\'!P{FIRST}:P{LAST}"),{k})'
                                     f'-INDEX(INDIRECT("\'"&PickProject&"\'!E{FIRST}:E{LAST}"),{k}),0)')
        ro.Range(f"G{r}").Formula = f"=G{r - 1}+B{r}"
        ro.Range(f"C{r}").Formula = f"=MIN(G{r - 1},G{r})-{floor}"
        ro.Range(f"D{r}").Formula = f"=MAX(B{r},0)"
        ro.Range(f"E{r}").Formula = f"=MAX(-B{r},0)"
    r = w0 + 11
    ro.Range(f"A{r}").Value = "COs"
    ro.Range(f"B{r}").Formula = f'=N({R("co_cost", "pick")})'
    ro.Range(f"G{r}").Formula = f"=G{r - 1}+B{r}"
    ro.Range(f"C{r}").Formula = f"=MIN(G{r - 1},G{r})-{floor}"
    ro.Range(f"D{r}").Formula = f"=MAX(B{r},0)"
    ro.Range(f"E{r}").Formula = f"=MAX(-B{r},0)"
    r = w0 + 12
    ro.Range(f"A{r}").Formula = f'="Forecast"&CHAR(10)&TEXT({R("fac_co", "pick")}/1000000,"0.00")&"M"'
    ro.Range(f"B{r}").Formula = f'=N({R("fac_co", "pick")})'
    ro.Range(f"F{r}").Formula = f"=B{r}-{floor}"
    ro.Range(f"H{w0 - 1}").Value = "Check"
    ro.Range(f"H{w0}").Formula = f"=ROUND(G{w0 + 11}-B{w0 + 12},2)"  # waterfall closes on the forecast
    w_last = w0 + 12
    ro.Range(f"A{w0 - 2}:G{w0 - 2}").Value = ("Label", "Value / change", "Base", "Overrun", "Saving", "Total", "Running level")
    ro.Range(f"B{w0}:G{w_last}").NumberFormat = "#,##0"

    section(ws, "B13", "Budget to forecast by cost code")
    section(ws, "M13", "Earned value detail")
    ws.Rows(13).RowHeight = 24
    for r in range(14, 31):
        ws.Rows(r).RowHeight = 17
    ch = add_chart(ws, "B14:L30", 52)  # stacked column
    xr = ro.Range(f"A{w0}:A{w_last}")
    s = add_series(ch, "Base", ro.Range(f"C{w0}:C{w_last}"), xr)
    s.Format.Fill.Visible = False; s.Format.Line.Visible = False
    s = add_series(ch, "Overrun", ro.Range(f"D{w0}:D{w_last}"), xr, BAD)
    s.HasDataLabels = True; s.DataLabels().NumberFormat = '+#,##0,"k";;'; s.DataLabels().Position = -4108
    s.DataLabels().Font.Color = rgb(IVORY); s.DataLabels().Font.Size = 8
    s = add_series(ch, "Saving", ro.Range(f"E{w0}:E{w_last}"), xr, GOOD)
    s.HasDataLabels = True; s.DataLabels().NumberFormat = '-#,##0,"k";;'; s.DataLabels().Position = -4108
    s.DataLabels().Font.Color = rgb(IVORY); s.DataLabels().Font.Size = 8
    s = add_series(ch, "Budget / forecast", ro.Range(f"F{w0}:F{w_last}"), xr, GOLD)
    dark_chart(ch, None, legend=True)
    ch.Legend.LegendEntries(1).Delete()
    ch.ChartGroups(1).GapWidth = 35
    ch.ChartGroups(1).Overlap = 100
    ch.Axes(2).TickLabelPosition = -4142
    ch.Axes(2).HasMajorGridlines = False
    ch.Axes(1).TickLabels.Font.Size = 8
    ch.Axes(1).TickLabelSpacing = 1
    ws.Range("B31").Value = "Bars between the totals = forecast minus budget per cost code (red over, green under). The axis starts above zero to show the changes."
    ws.Range("B31").Font.Color = rgb(MUTED); ws.Range("B31").Font.Size = 8; ws.Range("B31").Font.Italic = True

    evm = [("Budget at completion (BAC)", "budget", "#,##0"), ("Planned value (PV)", "planned", "#,##0"),
           ("Earned value (EV)", "earned", "#,##0"), ("Actual cost (AC)", "actual", "#,##0"),
           ("Cost variance  CV = EV - AC", "cv", "#,##0;-#,##0"), ("Schedule variance  SV = EV - PV", "sv", "#,##0;-#,##0"),
           ("CPI = EV / AC", "cpi", "0.00"), ("SPI = EV / PV", "spi", "0.00"),
           ("EAC cost codes (actual / % complete)", "fac", "#,##0"), ("EAC CPI method  BAC / CPI", "eac_cpi", "#,##0"),
           ("EAC CPI × SPI  AC + (BAC-EV)/(CPI×SPI)", "eac_cpi_spi", "#,##0"), ("Approved change orders (cost)", "co_cost", "#,##0"),
           ("Forecast incl. change orders", "fac_co", "#,##0"), ("Estimate to complete (ETC)", "etc", "#,##0"),
           ("Variance at completion", "vac", "#,##0;-#,##0"), ("TCPI  (BAC-EV)/(BAC-AC)", "tcpi", "0.00")]
    for i, (label, key, fmt) in enumerate(evm):
        r = 14 + i
        ws.Range(f"M{r}:P{r}").Merge()
        ws.Range(f"M{r}").Value = label
        ws.Range(f"M{r}").Font.Color = rgb(MUTED); ws.Range(f"M{r}").Font.Size = 9
        ws.Range(f"Q{r}:S{r}").Merge()
        ws.Range(f"Q{r}").Formula = f"={R(key, 'pick')}"
        ws.Range(f"Q{r}").NumberFormat = fmt
        ws.Range(f"Q{r}").HorizontalAlignment = -4152
        ws.Range(f"Q{r}").Font.Bold = key in ("fac_co", "cpi", "spi", "tcpi")
    table_body(ws.Range("M14:S29"))
    ws.Range("M26:S26").Interior.Color = rgb(CARD2)
    ws.Range("Q26").Font.Color = rgb(GOLD_LT)
    neg_red(ws.Range("Q18:Q19")); neg_red(ws.Range("Q28"))
    rag_index(ws.Range("Q20"), "Q20", "GreenCPI", "AmberCPI")
    rag_index(ws.Range("Q21"), "Q21", "GreenSPI", "AmberSPI")
    ws.Range("M30").Value = "ETC = forecast incl. COs - actual. TCPI above 1 = the rest of the work must cost less than planned."
    ws.Range("M30").Font.Color = rgb(MUTED); ws.Range("M30").Font.Size = 8; ws.Range("M30").Font.Italic = True
    ws.Rows(32).RowHeight = 10

    # cost codes of the picked project + PM notes
    section(ws, "B33", "Cost codes")
    section(ws, "N33", "From the PM this week")
    ws.Rows(33).RowHeight = 24
    hdr = {"B": "Cost code", "C": "Description", "F": "Budget", "G": "Actual", "H": "Planned %", "I": "% complete",
           "J": "CPI", "K": "SPI", "L": "Forecast", "M": "Variance"}
    for c, h in hdr.items():
        ws.Range(f"{c}34").Value = h
    ws.Range("C34:E34").Merge()
    table_head(ws.Range("B34:M34"))
    ws.Rows(34).RowHeight = 22
    srcs = {"B": "B", "C": "C", "F": "E", "G": "G", "H": "H", "I": "I", "J": "N", "K": "O", "L": "P", "M": "Q"}
    for k in range(1, 11):
        r = 34 + k
        ws.Range(f"C{r}:E{r}").Merge()
        for c, sc in srcs.items():
            ws.Range(f"{c}{r}").Formula = f'=IFERROR(INDEX(INDIRECT("\'"&PickProject&"\'!{sc}{FIRST}:{sc}{LAST}"),{k}),"")'
        ws.Rows(r).RowHeight = 18
    table_body(ws.Range("B35:M44"))
    ws.Range("B35:B44").Font.Color = rgb(GOLD_LT)
    ws.Range("F35:G44").NumberFormat = "#,##0"; ws.Range("L35:L44").NumberFormat = "#,##0"
    ws.Range("M35:M44").NumberFormat = "#,##0;-#,##0;–"
    neg_red(ws.Range("M35:M44"))
    ws.Range("H35:I44").NumberFormat = "0%"
    ws.Range("J35:K44").NumberFormat = "0.00"
    ws.Range("H35:K44").HorizontalAlignment = -4108
    ws.Range("C35:C44").Font.Size = 9
    rag_index(ws.Range("J35:J44"), "J35", "GreenCPI", "AmberCPI")
    rag_index(ws.Range("K35:K44"), "K35", "GreenSPI", "AmberSPI")
    try:
        db = ws.Range("I35:I44").FormatConditions.AddDatabar()
        db.MinPoint.Modify(0, 0)
        db.MaxPoint.Modify(0, 1)
        db.BarColor.Color = rgb(JADE)
        db.BarFillType = 0
    except Exception as e:
        print("data bar skipped:", e)
    ws.Range("N35:S39").Merge()
    ws.Range("N35").Formula = f"={R('note', 'pick')}"
    ws.Range("N35:S39").WrapText = True
    ws.Range("N35:S39").VerticalAlignment = -4160
    ws.Range("N35:S39").Interior.Color = rgb(CARD)
    ws.Range("N35").Font.Size = 9.5; ws.Range("N35").IndentLevel = 1
    facts = [f'="Change orders:  "&{R("co_app_n", "pick")}&" approved (+"&TEXT({R("co_price", "pick")},"#,##0")&")  ·  "&{R("co_pend_n", "pick")}&" pending ("&TEXT({R("co_pending", "pick")},"#,##0")&")"',
             f'="Risks:  "&{R("risk_open", "pick")}&" open, "&{R("risk_high", "pick")}&" high  ·  expected cost "&TEXT({R("risk_exp", "pick")},"#,##0")',
             f'="Crew:  "&{R("heads", "pick")}&" on site vs "&{R("pheads", "pick")}&" planned this week"',
             f'="Billing:  billed "&TEXT({R("billed", "pick")},"#,##0")&" vs earned revenue "&TEXT({R("earned_rev", "pick")},"#,##0")']
    for i, f in enumerate(facts):
        r = 41 + i
        ws.Range(f"N{r}:S{r}").Merge()
        ws.Range(f"N{r}").Formula = f
        ws.Range(f"N{r}").Font.Size = 9
        ws.Range(f"N{r}").IndentLevel = 1
    table_body(ws.Range("N41:S44"))
    ws.Range("N40:S40").Interior.Color = rgb(PAGE)
    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$S$45"
    return dict(waterfall=(w0, w_last))


# ---------------------------------------------------------------- S-curve and crew
def build_scurve(wb, sheets, T, n_weeks):
    ws = wb.Worksheets("S-curve")
    dark(ws, grid(ws))
    header(ws, "S", '="S-curve  ·  "&PickCurve',
           "Planned value from the baseline schedule; earned value, actual cost and crew from the weekly History (Refresh All).",
           "S-curve", sheets[0])
    ws.Rows(6).RowHeight = 26
    ws.Range("B6").Value = "Project"
    ws.Range("B6").Font.Color = rgb(MUTED); ws.Range("B6").VerticalAlignment = -4108
    ws.Range("C6:D6").Merge()
    p = ws.Range("C6")
    p.NumberFormat = "@"
    p.Value = "P-2402" if "P-2402" in sheets else sheets[0]
    p.Interior.Color = rgb(INPUT_BG); p.Font.Color = rgb(GOLD_LT); p.Font.Bold = True; p.Font.Size = 12
    p.HorizontalAlignment = -4108; p.VerticalAlignment = -4108
    b = ws.Range("C6:D6").Borders(9); b.LineStyle = 1; b.Color = rgb(GOLD)
    dv = p.Validation
    dv.Delete()
    dv.Add(3, 1, 1, "=PrjIds")
    ws.Range("E6:S6").Merge()
    ws.Range("E6").Formula = f'={R("project", "curve")}&"  ·  PM "&{R("pm", "curve")}'
    ws.Range("E6").Font.Name = SERIF; ws.Range("E6").Font.Size = 13; ws.Range("E6").VerticalAlignment = -4108
    ws.Range("E6").IndentLevel = 1

    # week grid (right side): rows 15.. one per baseline week
    g0 = 15
    g_last = g0 + n_weeks - 1
    section(ws, "B13", "S-curve: planned, earned and actual (cumulative)")
    section(ws, "N13", "Week by week")
    ws.Rows(13).RowHeight = 24
    hdr = {"N": "Week ending", "O": "Planned value", "P": "Earned value", "Q": "Actual cost", "R": "Crew planned", "S": "Crew on site"}
    for c, h in hdr.items():
        ws.Range(f"{c}14").Value = h
    table_head(ws.Range("N14:S14"))
    ws.Rows(14).RowHeight = 30
    first_week = "MIN(Baseline[Week ending])"
    pk = "PickCurve"
    for i in range(n_weeks):
        r = g0 + i
        ws.Range(f"N{r}").Formula = f"={first_week}+7*{i}"
        ws.Range(f"O{r}").Formula = f"=SUMIFS(Baseline[Planned value],Baseline[Sheet],{pk},Baseline[Week ending],N{r})"
        for c, fld in (("P", "Earned"), ("Q", "Actual"), ("S", "Actual heads")):
            ws.Range(f"{c}{r}").Formula = (f"=IF(N{r}>WeekEnding,NA(),IFERROR(1/(1/SUMIFS(History[{fld}],History[Sheet],{pk},"
                                           f"History[Week ending],N{r})),NA()))")
        ws.Range(f"R{r}").Formula = f"=SUMIFS(Baseline[Planned heads],Baseline[Sheet],{pk},Baseline[Week ending],N{r})"
        ws.Rows(r).RowHeight = 15.5
    table_body(ws.Range(f"N{g0}:S{g_last}"))
    ws.Range(f"N{g0}:S{g_last}").Font.Size = 8.5
    ws.Range(f"N{g0}:N{g_last}").NumberFormat = "d mmm"
    ws.Range(f"O{g0}:Q{g_last}").NumberFormat = '#,##0,"k"'
    ws.Range(f"R{g0}:S{g_last}").NumberFormat = "0"
    ws.Range(f"N{g0}:S{g_last}").HorizontalAlignment = -4108
    hide_errors(ws.Range(f"P{g0}:S{g_last}"), f"P{g0}")
    cf(ws.Range(f"N{g0}:S{g_last}"), f"=$N{g0}=WeekEnding", fill=CARD2)
    cf(ws.Range(f"N{g0}:N{g_last}"), f"=N{g0}=WeekEnding", font=GOLD)
    cf(ws.Range(f"S{g0}:S{g_last}"), f"=AND(ISNUMBER(S{g0}),S{g0}>R{g0})", font=AMBER)

    # cards
    for r in range(8, 12):
        ws.Rows(r).RowHeight = 22
    ws.Rows(7).RowHeight = 10
    ws.Rows(12).RowHeight = 12
    pos = cards_row(ws, "B", "S", 8, 6, height=86)
    card(ws, *pos[0], "Planned value to date", T.add("SC pv", f"={R('planned', 'curve')}", "#,##0"),
         [(T.add("SC pv note", f'="of budget "&TEXT({R("budget", "curve")},"#,##0")&"  ·  "&TEXT({R("planned", "curve")}/{R("budget", "curve")},"0%")'), MUTED)], icon="◷")
    card(ws, *pos[1], "Earned value to date", T.add("SC ev", f"={R('earned', 'curve')}", "#,##0"),
         note3(T, "SC ev note", f'"SV "&TEXT({R("sv", "curve")},"#,##0;-#,##0")&"  ·  SPI "&TEXT({R("spi", "curve")},"0.00")',
               f"{R('spi', 'curve')}>=GreenSPI", f"{R('spi', 'curve')}<AmberSPI"), icon="▲")
    card(ws, *pos[2], "Actual cost to date", T.add("SC ac", f"={R('actual', 'curve')}", "#,##0"),
         note3(T, "SC ac note", f'"CV "&TEXT({R("cv", "curve")},"#,##0;-#,##0")&"  ·  CPI "&TEXT({R("cpi", "curve")},"0.00")',
               f"{R('cpi', 'curve')}>=GreenCPI", f"{R('cpi', 'curve')}<AmberCPI"), icon="$")
    # finish dates
    dur = T.add("SC planned weeks", f"=COUNTIF('S-curve'!$O${g0}:$O${g_last},\"<\"&{R('budget', 'curve')})+1", "0",
                name="CurvePlannedWeeks")
    fdur = T.add("SC forecast weeks", f'=IFERROR(ROUNDUP(CurvePlannedWeeks/{R("spi", "curve")},0),"")', "0", name="CurveForecastWeeks")
    card(ws, *pos[3], "Baseline finish", T.add("SC baseline finish", f"={first_week}+7*(CurvePlannedWeeks-1)", "d mmm yyyy", name="CurveBaselineFinish"),
         [(T.add("SC baseline finish note", '=CurvePlannedWeeks&" weeks planned"'), MUTED)], icon="⚑")
    card(ws, *pos[4], "Forecast finish (at SPI)", T.add("SC forecast finish", f"=IFERROR({first_week}+7*(CurveForecastWeeks-1),\"\")", "d mmm yyyy", name="CurveForecastFinish"),
         note3(T, "SC forecast finish note", 'IF(CurveForecastWeeks-CurvePlannedWeeks<=0,"on the baseline","+"&(CurveForecastWeeks-CurvePlannedWeeks)&" weeks late")',
               "CurveForecastWeeks-CurvePlannedWeeks<=0", "CurveForecastWeeks-CurvePlannedWeeks>=2"), icon="◷")
    card(ws, *pos[5], "Crew this week", T.add("SC crew", f"={R('heads', 'curve')}", "0"),
         note3(T, "SC crew note", f'"planned "&{R("pheads", "curve")}&"  ·  "&TEXT({R("heads", "curve")}-{R("pheads", "curve")},"+0;-0;0")&" heads"',
               f"{R('heads', 'curve')}<={R('pheads', 'curve')}", f"{R('heads', 'curve')}>{R('pheads', 'curve')}*1.15"), icon="⚒")

    # S-curve chart
    ch = add_chart(ws, f"B14:M{g_last}", 4)
    xr = ws.Range(f"N{g0}:N{g_last}")
    add_series(ch, "Planned value (baseline)", ws.Range(f"O{g0}:O{g_last}"), xr, GOLD, line=True, weight=2.25, dash=4)
    add_series(ch, "Earned value", ws.Range(f"P{g0}:P{g_last}"), xr, JADE, line=True, weight=3, marker=True)
    add_series(ch, "Actual cost", ws.Range(f"Q{g0}:Q{g_last}"), xr, COPPER, line=True, weight=3, marker=True)
    dark_chart(ch, None)
    ch.Axes(1).CategoryType = 2
    ch.Axes(1).TickLabels.NumberFormat = "d mmm"
    ch.Axes(1).TickLabelSpacing = 2
    ch.Axes(2).TickLabels.NumberFormat = '#,##0,"k"'
    ch.Axes(2).MinimumScale = 0
    ws.Range(f"B{g_last + 1}").Value = ("Earned value below the dashed plan = behind schedule; actual cost above earned value = over cost. "
                                        "Forecast finish = planned weeks / SPI.")
    ws.Range(f"B{g_last + 1}").Font.Color = rgb(MUTED); ws.Range(f"B{g_last + 1}").Font.Size = 8; ws.Range(f"B{g_last + 1}").Font.Italic = True
    m0 = g_last + 3
    section(ws, f"B{m0}", "Crew on site: planned vs actual heads per week")
    ws.Rows(m0).RowHeight = 24
    for r in range(m0 + 1, m0 + 15):
        ws.Rows(r).RowHeight = 15
    ch = add_chart(ws, f"B{m0 + 1}:S{m0 + 14}", 51)
    add_series(ch, "Crew planned (baseline)", ws.Range(f"R{g0}:R{g_last}"), xr, GREY)
    s = add_series(ch, "Crew on site", ws.Range(f"S{g0}:S{g_last}"), xr, SKY)
    s.HasDataLabels = True
    s.DataLabels().Font.Color = rgb(IVORY); s.DataLabels().Font.Size = 8
    dark_chart(ch, None)
    ch.Axes(1).CategoryType = 2
    ch.Axes(1).TickLabels.NumberFormat = "d mmm"
    ch.ChartGroups(1).GapWidth = 60
    ch.ChartGroups(1).Overlap = -10
    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = f"$A$1:$S${m0 + 15}"
    return dict(grid=(g0, g_last), crew=m0)


# ---------------------------------------------------------------- change orders & risks
def build_cos_cards(wb, T):
    ws = wb.Worksheets("COs & risks")
    pos = cards_row(ws, "B", "Q", 6, 4, height=86)
    card(ws, *pos[0], "Approved change orders", T.add("CO approved", f"={R('co_price')}", "#,##0"),
         [(T.add("CO approved note", f'={R("co_app_n")}&" approved  ·  cost "&TEXT({R("co_cost")},"#,##0")&"  ·  margin "&TEXT(1-{R("co_cost")}/{R("co_price")},"0%")'), GOOD)], icon="✓")
    card(ws, *pos[1], "Pending change orders", T.add("CO pending", f"={R('co_pending')}", "#,##0"),
         [(T.add("CO pending note", f'={R("co_pend_n")}&" waiting for the owner: upside not in the forecast"'), AMBER)], icon="◷")
    card(ws, *pos[2], "Risk exposure (expected cost)", T.add("Risk exposure", f"={R('risk_exp')}", "#,##0"),
         [(T.add("Risk exposure note", f'={R("risk_open")}&" open risks  ·  cost if all happen "&TEXT(SUM(Risks[Cost if it happens]),"#,##0")'), MUTED)], icon="⚠")
    card(ws, *pos[3], "High risks (score 15+)", T.add("Risk high", f"={R('risk_high')}", "0"),
         [(T.add("Risk high note", '=IFERROR("top: "&INDEX(Risks[ID],MATCH(MAX(Risks[Expected cost]),Risks[Expected cost],0))&" "&'
                                   'INDEX(Risks[Risk],MATCH(MAX(Risks[Expected cost]),Risks[Expected cost],0)),"")'), BAD)], icon="!")



def build_cos_risks(wb, sheets):
    ws = wb.Worksheets("COs & risks")
    widths = {"A": 2, "B": 8, "C": 9, "D": 46, "E": 13, "F": 11, "G": 11, "H": 11, "I": 12, "J": 12, "K": 3, "L": 13,
              "M": 9.5, "N": 9.5, "O": 9.5, "P": 9.5, "Q": 9.5}
    dark(ws, widths)
    header(ws, "Q", "Change orders & risks",
           "Approved change orders flow into the revised contract and the forecast; pending ones are the upside still to win.",
           "COs & risks", sheets[0])
    for r in range(6, 10):
        ws.Rows(r).RowHeight = 22
    ws.Rows(10).RowHeight = 12
    # change orders table
    section(ws, "B11", "Change orders")
    ws.Rows(11).RowHeight = 24
    cos = read_csv("change-orders.csv")
    data = [["CO", "Sheet", "Description", "Submitted", "Status", "Price", "Cost estimate", "CO margin"]]
    for c in cos:
        data.append([c["CO"], c["Sheet"], c["Description"], serial(date.fromisoformat(c["Submitted"])), c["Status"],
                     float(c["Price"]), float(c["Cost estimate"]), None])
    n = len(data)
    ws.Range(f"B12:C{11 + n}").NumberFormat = "@"
    ws.Range(ws.Cells(12, 2), ws.Cells(11 + n, 9)).Value = data
    lo = ws.ListObjects.Add(1, ws.Range(f"B12:I{11 + n}"), None, 1)
    lo.Name = "ChangeOrders"
    try:
        lo.TableStyle = ""
    except Exception:
        lo.TableStyle = "TableStyleLight1"
    lo.ListColumns("CO margin").DataBodyRange.Formula = '=IF([@Price]=0,"",1-[@[Cost estimate]]/[@Price])'
    table_head(lo.HeaderRowRange)
    table_body(lo.DataBodyRange)
    lo.ListColumns("Submitted").DataBodyRange.NumberFormat = "d mmm"
    lo.ListColumns("Price").DataBodyRange.NumberFormat = "#,##0"
    lo.ListColumns("Cost estimate").DataBodyRange.NumberFormat = "#,##0;-#,##0;–"
    lo.ListColumns("CO margin").DataBodyRange.NumberFormat = "0%;-0%;–"
    st = lo.ListColumns("Status").DataBodyRange
    st.HorizontalAlignment = -4108
    for text, (bg, fg) in (("Approved", (GOOD_BG, GOOD)), ("Pending", (AMBER_BG, AMBER)), ("Rejected", (BAD_BG, BAD))):
        fc = st.FormatConditions.Add(1, 3, f'="{text}"')
        fc.Interior.Color = rgb(bg); fc.Font.Color = rgb(fg); fc.Font.Bold = True
    dv = st.Validation
    dv.Delete()
    dv.Add(3, 1, 1, "=StatusList")
    lo.ListColumns("CO").DataBodyRange.Font.Color = rgb(GOLD_LT)
    co_last = 11 + n
    for r in range(13, co_last + 1):
        ws.Rows(r).RowHeight = 18

    # CO summary by project (right)
    ws.Range("L12:Q12").Value = ("Project", "Approved", "Pending", "Rejected", "Revised contract", "Margin")
    table_head(ws.Range("L12:Q12"))
    for i in range(SLOTS):
        r = 13 + i
        rr = R_FIRST + i
        ws.Range(f"L{r}").Formula = f"=Rollup!$A{rr}"
        ws.Range(f"M{r}").Formula = f"=Rollup!{RC['co_price']}{rr}"
        ws.Range(f"N{r}").Formula = f"=Rollup!{RC['co_pending']}{rr}"
        ws.Range(f"O{r}").Formula = f'=IF(L{r}="","",SUMIFS(ChangeOrders[Price],ChangeOrders[Sheet],L{r},ChangeOrders[Status],"Rejected"))'
        ws.Range(f"P{r}").Formula = f"=Rollup!{RC['rev_contract']}{rr}"
        ws.Range(f"Q{r}").Formula = f"=Rollup!{RC['margin']}{rr}"
    table_body(ws.Range(f"L13:Q{12 + SLOTS}"))
    ws.Range(f"M13:P{12 + SLOTS}").NumberFormat = '#,##0;-#,##0;"–"'
    ws.Range(f"Q13:Q{12 + SLOTS}").NumberFormat = "0.0%;-0.0%"
    neg_red(ws.Range(f"Q13:Q{12 + SLOTS}"))
    ws.Range(f"L13:L{12 + SLOTS}").Font.Color = rgb(GOLD_LT)
    ws.Range(f"L{13 + SLOTS}:Q{13 + SLOTS}").Merge()
    ws.Range(f"L{13 + SLOTS}").Value = "Approved: price added to the contract, cost estimate added to the forecast."
    ws.Range(f"L{13 + SLOTS}").Font.Color = rgb(MUTED); ws.Range(f"L{13 + SLOTS}").Font.Size = 8; ws.Range(f"L{13 + SLOTS}").Font.Italic = True

    # risk register
    r0 = max(co_last, 13 + SLOTS) + 2
    section(ws, f"B{r0}", "Risk register")
    section(ws, f"L{r0}", "Risk matrix (IDs per cell)")
    ws.Rows(r0).RowHeight = 24
    risks = read_csv("risks.csv")
    data = [["ID", "Sheet", "Risk", "Owner", "Probability", "Impact", "Score", "Cost if it happens", "Expected cost"]]
    for x in risks:
        data.append([x["ID"], x["Sheet"], x["Risk"], x["Owner"], int(x["Probability"]), int(x["Impact"]), None,
                     float(x["Cost if it happens"]), None])
    n = len(data)
    ws.Range(f"B{r0 + 1}:C{r0 + n}").NumberFormat = "@"
    ws.Range(ws.Cells(r0 + 1, 2), ws.Cells(r0 + n, 10)).Value = data
    lo = ws.ListObjects.Add(1, ws.Range(f"B{r0 + 1}:J{r0 + n}"), None, 1)
    lo.Name = "Risks"
    try:
        lo.TableStyle = ""
    except Exception:
        lo.TableStyle = "TableStyleLight1"
    lo.ListColumns("Score").DataBodyRange.Formula = "=[@Probability]*[@Impact]"
    lo.ListColumns("Expected cost").DataBodyRange.Formula = "=[@[Cost if it happens]]*IFERROR(CHOOSE([@Probability],0.1,0.3,0.5,0.7,0.9),0)"
    table_head(lo.HeaderRowRange)
    table_body(lo.DataBodyRange)
    for c in ("Probability", "Impact", "Score"):
        lo.ListColumns(c).DataBodyRange.HorizontalAlignment = -4108
    for c in ("Cost if it happens", "Expected cost"):
        lo.ListColumns(c).DataBodyRange.NumberFormat = "#,##0"
    lo.ListColumns("ID").DataBodyRange.Font.Color = rgb(GOLD_LT)
    sc = lo.ListColumns("Score").DataBodyRange
    a = sc.Cells(1, 1).Address.replace("$", "")
    cf(sc, f"={a}>=15", font=BAD, fill=BAD_BG, bold=True)
    cf(sc, f"=AND({a}>=8,{a}<15)", font=AMBER, fill=AMBER_BG, bold=True)
    cf(sc, f"={a}<8", font=GOOD, fill=GOOD_BG)
    for c in ("Probability", "Impact"):
        dv = lo.ListColumns(c).DataBodyRange.Validation
        dv.Delete()
        dv.Add(1, 1, 1, "1", "5")  # whole number 1-5
        dv.ErrorMessage = "1 (low) to 5 (high)."
    rk_last = r0 + n
    for r in range(r0 + 1, rk_last + 1):
        ws.Rows(r).RowHeight = 18
    ws.Range(f"B{rk_last + 2}").Value = "Score = probability x impact. Expected cost = cost if it happens x 10 / 30 / 50 / 70 / 90% for probability 1-5."
    ws.Range(f"B{rk_last + 2}").Font.Color = rgb(MUTED); ws.Range(f"B{rk_last + 2}").Font.Size = 8; ws.Range(f"B{rk_last + 2}").Font.Italic = True

    # risk helper on the Rollup sheet: per risk row, the IDs collected so far for the same matrix cell
    ro = wb.Worksheets("Rollup")
    h0 = 60
    ro.Range(f"A{h0 - 1}").Value = "Risk matrix helper: per risk row, matrix cell key (probability*10+impact) and the IDs in that cell so far"
    ro.Range(f"A{h0 - 1}").Font.Bold = True
    for k in range(1, 41):
        r = h0 + k - 1
        ro.Range(f"A{r}").Formula = f'=IFERROR(INDEX(Risks[ID],{k})&"","")'
        ro.Range(f"B{r}").Formula = f'=IF(A{r}="","",INDEX(Risks[Probability],{k})*10+INDEX(Risks[Impact],{k}))'
        if k == 1:
            ro.Range(f"C{r}").Formula = f'=IF(A{r}="","",A{r})'
        else:
            ro.Range(f"C{r}").Formula = (f'=IF(A{r}="","",IF(COUNTIF(B${h0}:B{r - 1},B{r})=0,A{r},'
                                         f'LOOKUP(2,1/(B${h0}:B{r - 1}=B{r}),C${h0}:C{r - 1})&", "&A{r}))')
    keys, ids = f"Rollup!$B${h0}:$B${h0 + 39}", f"Rollup!$C${h0}:$C${h0 + 39}"
    # matrix: rows impact 5..1, columns probability 1..5
    m1 = r0 + 1
    ws.Range(f"L{m1}").Value = "Impact ↓  Prob. →"
    for pcol in range(1, 6):
        ws.Cells(m1, 12 + pcol).Value = pcol
    table_head(ws.Range(f"L{m1}:Q{m1}"))
    ws.Range(f"M{m1}:Q{m1}").HorizontalAlignment = -4108
    labels = {5: "5  severe", 4: "4  major", 3: "3  moderate", 2: "2  minor", 1: "1  low"}
    for i, imp in enumerate(range(5, 0, -1)):
        r = m1 + 1 + i
        ws.Range(f"L{r}").Value = labels[imp]
        ws.Range(f"L{r}").Font.Color = rgb(MUTED); ws.Range(f"L{r}").Font.Size = 8.5
        ws.Range(f"L{r}").Interior.Color = rgb(CARD)
        for pcol in range(1, 6):
            c = ws.Cells(r, 12 + pcol)
            c.Formula = f'=IFERROR(LOOKUP(2,1/({keys}={pcol * 10 + imp}),{ids}),"")'
            score = pcol * imp
            bg, fg = (BAD_BG, BAD) if score >= 15 else (AMBER_BG, AMBER) if score >= 8 else (GOOD_BG, GOOD)
            c.Interior.Color = rgb(bg); c.Font.Color = rgb(fg); c.Font.Bold = True; c.Font.Size = 8
            c.HorizontalAlignment = -4108; c.WrapText = True
            for e in (7, 8, 9, 10):
                c.Borders(e).LineStyle = 1; c.Borders(e).Color = rgb(PAGE); c.Borders(e).Weight = -4138
        ws.Rows(r).RowHeight = 24
    ws.Range(f"L{m1 + 6}:Q{m1 + 6}").Merge()
    ws.Range(f"L{m1 + 6}").Value = "Red = score 15+, amber 8-14, green below 8."
    ws.Range(f"L{m1 + 6}").Font.Color = rgb(MUTED); ws.Range(f"L{m1 + 6}").Font.Size = 8; ws.Range(f"L{m1 + 6}").Font.Italic = True
    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = f"$A$1:$Q${rk_last + 2}"
    return dict(co=(13, co_last), risks=(r0 + 2, rk_last), matrix=m1)


# ---------------------------------------------------------------- WIP & cash
def build_wip(wb, sheets, T):
    ws = wb.Worksheets("WIP & cash")
    dark(ws, grid(ws))
    header(ws, "S", "WIP schedule & cash",
           "Percentage of completion (cost-to-cost): earned revenue vs billed = over / under billing. Billing and cash typed by the PMs.",
           "WIP & cash", sheets[0])
    for r in range(6, 10):
        ws.Rows(r).RowHeight = 22
    ws.Rows(10).RowHeight = 12
    pos = cards_row(ws, "B", "S", 6, 4, height=86)
    card(ws, *pos[0], "Over / (under) billed", T.add("WIP over under", f"={R('over_under')}", "#,##0;(#,##0)"),
         note3(T, "WIP over under note", f'IF({R("over_under")}<0,"under billed: bill the earned work","billed ahead of the work")',
               f"{R('over_under')}>=0", f"{R('over_under')}<0"), icon="⇄")
    card(ws, *pos[1], "Retainage held by owners", T.add("WIP retainage", f"={R('retainage')}", "#,##0"),
         [(T.add("WIP retainage note", f'=TEXT({R("retainage")}/{R("billed")},"0.0%")&" of "&TEXT({R("billed")},"#,##0")&" billed"'), MUTED)], icon="⛁")
    card(ws, *pos[2], "Net cash (received - paid)", T.add("WIP net cash", f"={R('net_cash')}", "#,##0;-#,##0"),
         note3(T, "WIP net cash note", f'"received "&TEXT({R("received")},"#,##0")&"  ·  paid "&TEXT({R("paid")},"#,##0")',
               f"{R('net_cash')}>=0", f"{R('net_cash')}<0"), icon="$")
    card(ws, *pos[3], "Receivable (excl. retainage)", T.add("WIP ar", f"={R('ar')}", "#,##0"),
         [(T.add("WIP ar note", '="billed, not retained, not yet paid"'), MUTED)], icon="◷")
    section(ws, "B11", "WIP schedule")
    ws.Rows(11).RowHeight = 24
    cols = [("B", "Sheet", "sheet"), ("C", "Project", "project"), ("E", "Contract", "contract"), ("F", "Approved COs", "co_price"),
            ("G", "Revised contract", "rev_contract"), ("H", "Forecast cost", "fac_co"), ("I", "Est. gross profit", None),
            ("J", "Cost to date", "actual"), ("K", "% complete", "pct_cost"), ("L", "Earned revenue", "earned_rev"),
            ("M", "Billed to date", "billed"), ("N", "Over / (under) billed", "over_under"), ("O", "Retainage held", "retainage"),
            ("P", "Cash received", "received"), ("Q", "Cash paid", "paid"), ("R", "Net cash", "net_cash"), ("S", "Crew / plan", None)]
    for c, h, _ in cols:
        ws.Range(f"{c}12").Value = h
    ws.Range("C12:D12").Merge()
    table_head(ws.Range("B12:S12"))
    ws.Rows(12).RowHeight = 30
    w_last = 12 + SLOTS
    t = w_last + 1
    for i in range(SLOTS + 1):
        r = 13 + i
        rr = R_FIRST + i if i < SLOTS else R_TOT
        ws.Range(f"C{r}:D{r}").Merge()
        for c, h, k in cols:
            if k:
                ws.Range(f"{c}{r}").Formula = f"=Rollup!{RC[k]}{rr}"
        ws.Range(f"I{r}").Formula = f'=IF(Rollup!$A{rr}="","",G{r}-H{r})'
        if i < SLOTS:
            ws.Range(f"S{r}").Formula = f'=IF(Rollup!$A{rr}="","",Rollup!{RC["heads"]}{rr}&" / "&Rollup!{RC["pheads"]}{rr})'
        ws.Rows(r).RowHeight = 20
    ws.Range(f"C{t}").Formula = f'=SUMPRODUCT(--(B13:B{w_last}<>""))&" projects"'
    ws.Range(f"S{t}").Formula = f'=Rollup!{RC["heads"]}{R_TOT}&" / "&Rollup!{RC["pheads"]}{R_TOT}'
    table_body(ws.Range(f"B13:S{w_last}"))
    table_total(ws.Range(f"B{t}:S{t}"))
    ws.Range(f"B13:B{t}").Font.Color = rgb(GOLD_LT)
    ws.Range(f"E13:R{t}").NumberFormat = '#,##0;-#,##0;"–"'
    ws.Range(f"K13:K{t}").NumberFormat = "0.0%"
    ws.Range(f"N13:N{t}").NumberFormat = '#,##0;(#,##0);"–"'
    ws.Range(f"K13:K{t}").HorizontalAlignment = -4108
    ws.Range(f"S13:S{t}").HorizontalAlignment = -4108
    neg_red(ws.Range(f"N13:N{t}")); neg_red(ws.Range(f"R13:R{t}")); neg_red(ws.Range(f"I13:I{t}"))
    ws.Range(f"C13:C{t}").Font.Size = 9
    ws.Range(f"C13:D{t}").ShrinkToFit = True
    ws.Range(f"B{t + 1}").Value = ("% complete = cost to date / forecast cost (incl. approved COs). Earned revenue = revised contract x % complete. "
                                   "Positive = billed ahead of the work (a liability), negative = under billed (cash left on the table).")
    ws.Range(f"B{t + 1}").Font.Color = rgb(MUTED); ws.Range(f"B{t + 1}").Font.Size = 8; ws.Range(f"B{t + 1}").Font.Italic = True
    c0 = t + 3
    for r in range(c0, c0 + 17):
        ws.Rows(r).RowHeight = 15
    book = wb.Name
    ch = add_chart(ws, f"B{c0}:J{c0 + 16}", 51)
    for nm, cc in (("WipEarned", "L"), ("WipBilled", "M")):
        n = wb.Names.Add(nm, f"='WIP & cash'!${cc}$13")
        n.RefersToLocal = f"=OFFSET('WIP & cash'!${cc}$13,0,0,MAX(1,COUNTA(Dashboard!$B${D_FIRST}:$B${D_FIRST + SLOTS - 1})),1)".replace(",", SEP)
    for name, ref, color in (("Earned revenue", "WipEarned", JADE), ("Billed to date", "WipBilled", GOLD)):
        s = ch.SeriesCollection().NewSeries()
        s.Formula = f"=SERIES(\"{name}\",'{book}'!PrjIds,'{book}'!{ref},{ch.SeriesCollection().Count})"
        s.Format.Fill.ForeColor.RGB = rgb(color)
        s.Format.Line.Visible = False
    dark_chart(ch, "Earned revenue vs billed, by project")
    ch.Axes(2).TickLabels.NumberFormat = '#,##0,"k"'
    ch.ChartGroups(1).GapWidth = 70
    return dict(table=(13, t), chart=c0)


def wip_cash_chart(wb, trend):
    ws = wb.Worksheets("WIP & cash")
    c0 = trend["wip_chart_row"]
    g0, g1 = trend["rows"]
    tr = wb.Worksheets("Weekly trend")
    ch = add_chart(ws, f"K{c0}:S{c0 + 16}", 4)
    xr = tr.Range(f"B{g0}:B{g1}")
    for key, color, dash in (("Billed", GOLD, 4), ("Cash received", JADE, None), ("Cash paid", COPPER, None)):
        c = trend["cols"][key]
        add_series(ch, key, tr.Range(f"{c}{g0}:{c}{g1}"), xr, color, line=True, weight=2.5, dash=dash, marker=dash is None)
    dark_chart(ch, "Cash in vs cash out, week by week (all projects, to date)")
    ch.Axes(1).CategoryType = 2
    ch.Axes(1).TickLabels.NumberFormat = "d mmm"
    ch.Axes(2).TickLabels.NumberFormat = '#,##0.0,,"M";;0'
    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = f"$A$1:$S${c0 + 17}"


# ---------------------------------------------------------------- snapshots (Power Query history)
def m_value(v, c):
    if c == "Week ending":
        d = date.fromisoformat(v)
        return f"#date({d.year},{d.month},{d.day})"
    if c in ("Sheet", "Project"):
        return '"' + v.replace('"', '""') + '"'
    return v


TYPES = "{" + ", ".join('{"%s", %s}' % (c, "type date" if c == "Week ending" else "type text" if c in ("Sheet", "Project") else "type number")
                        for c in HIST_COLS) + "}"
SEED_KEYS = ["week", "sheet", "project", "contract", "budget", "committed", "actual", "earned", "planned", "field", "shop", "cpi",
             "spi", "fac", "planned_heads", "actual_heads", "billed", "received", "paid"]
THIS_WEEK_SRC = {"Sheet": "sheet", "Project": "project", "Contract": "contract", "Budget": "budget", "Committed": "committed",
                 "Actual": "actual", "Earned": "earned", "Planned": "planned", "Field %": "field", "Shop %": "shop", "CPI": "cpi",
                 "SPI": "spi", "Forecast at completion": "fac", "Planned heads": "pheads", "Actual heads": "heads",
                 "Billed": "billed", "Cash received": "received", "Cash paid": "paid"}

CURRENT = f"""
    // this week: one row per project from the ThisWeek table (formulas that read the Rollup sheet)
    Raw      = Excel.CurrentWorkbook(){{[Name = "ThisWeek"]}}[Content],
    Filled   = Table.SelectRows(Raw, each [Sheet] <> null and [Sheet] <> ""),
    Current  = Table.TransformColumnTypes(Table.SelectColumns(Filled, {{{", ".join('"' + c + '"' for c in HIST_COLS)}}}), {TYPES}),
    Weeks    = List.Buffer(List.Distinct(Current[Week ending])),"""

TAIL_M = """
    // re-running the same week replaces that week's rows, so a week is never saved twice
    Keep     = Table.SelectRows(Old, each not List.Contains(Weeks, [Week ending])),
    All      = Table.Combine({Keep, Current}),
    Sorted   = Table.Sort(All, {{"Week ending", Order.Ascending}, {"Sheet", Order.Ascending}})
in
    Sorted"""


def seed_query():
    seed = read_csv("history-seed.csv")
    rows = ",\n        ".join("{" + ", ".join(m_value(r[k], c) for k, c in zip(SEED_KEYS, HIST_COLS)) + "}" for r in seed)
    cols = ", ".join(f'#"{c}" = {"date" if c == "Week ending" else "text" if c in ("Sheet", "Project") else "number"}' for c in HIST_COLS)
    return (f"let{CURRENT}\n    // first build only: the seven earlier weekly snapshots (fictional)\n"
            f"    Old      = #table(type table [{cols}], {{\n        {rows}}}),{TAIL_M}")


HISTORY_M = (f"let{CURRENT}\n    // the history saved so far = this query's own output table on the Snapshots sheet\n"
             f"    Old      = Table.TransformColumnTypes(Excel.CurrentWorkbook(){{[Name = \"History\"]}}[Content], {TYPES}),{TAIL_M}")


def conn_string(name):
    return f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=""'


def build_snapshots(wb, sheets):
    ws = wb.Worksheets("Snapshots")
    last = col(len(HIST_COLS))
    ws.Range(f"A1:{last}3").Interior.Color = rgb(HEAD)
    ws.Range(f"A4:{last}4").Interior.Color = rgb(GOLD)
    ws.Rows(4).RowHeight = 2.25
    ws.Range("A1").Value = "B R I G H T W I R E    ·    P R O J E C T   C O N T R O L S"
    ws.Range("A1").Font.Size = 8; ws.Range("A1").Font.Bold = True; ws.Range("A1").Font.Color = rgb(GOLD)
    ws.Range("A2").Value = "Weekly snapshots"
    ws.Range("A2").Font.Name = SERIF; ws.Range("A2").Font.Size = 18; ws.Range("A2").Font.Color = rgb(IVORY)
    ws.Rows(2).RowHeight = 28
    ws.Range("A3").Value = "Data > Refresh All copies 'This week' into History. Same week again = replaced, never doubled. No macros."
    ws.Range("A3").Font.Color = rgb(MUTED); ws.Range("A3").Font.Size = 9
    ws.Range("A6").Value = "This week (live from the Rollup sheet)"
    ws.Range("A6").Font.Bold = True; ws.Range("A6").Font.Size = 12
    ws.Range(f"A7:{last}7").Value = HIST_COLS
    for i in range(SLOTS):
        r = 8 + i
        rr = R_FIRST + i
        for j, c in enumerate(HIST_COLS, 1):
            cell = ws.Cells(r, j)
            if c == "Week ending":
                cell.Formula = f'=IF(Rollup!$A{rr}="","",WeekEnding)'
            else:
                cell.Formula = f'=IF(Rollup!$A{rr}="","",Rollup!{RC[THIS_WEEK_SRC[c]]}{rr})'
    lo = ws.ListObjects.Add(1, ws.Range(f"A7:{last}{7 + SLOTS}"), None, 1)
    lo.Name = "ThisWeek"
    lo.TableStyle = "TableStyleLight9"
    h_title = 9 + SLOTS
    ws.Range(f"A{h_title}").Value = "History (saved by Refresh All, one row per project and week)"
    ws.Range(f"A{h_title}").Font.Bold = True; ws.Range(f"A{h_title}").Font.Size = 12
    wb.Queries.Add("History", seed_query())
    lo2 = ws.ListObjects.Add(0, conn_string("History"), None, 1, ws.Range(f"A{h_title + 1}"))
    qt = lo2.QueryTable
    qt.CommandType = 2
    qt.CommandText = "SELECT * FROM [History]"
    qt.RefreshStyle = 1
    qt.AdjustColumnWidth = False
    qt.BackgroundQuery = False
    qt.PreserveFormatting = True
    qt.Refresh(False)
    lo2.Name = "History"
    lo2.TableStyle = "TableStyleMedium2"
    wb.Queries("History").Formula = HISTORY_M  # from now on the query appends to its own table
    qt.Refresh(False)
    for l in (lo, lo2):
        l.ListColumns("Week ending").Range.NumberFormat = "d mmm yyyy"
        for c in ("Contract", "Budget", "Committed", "Actual", "Earned", "Planned", "Forecast at completion", "Billed",
                  "Cash received", "Cash paid"):
            l.ListColumns(c).DataBodyRange.NumberFormat = "#,##0"
        for c in ("Field %", "Shop %"):
            l.ListColumns(c).DataBodyRange.NumberFormat = "0.0%"
        for c in ("CPI", "SPI"):
            l.ListColumns(c).DataBodyRange.NumberFormat = "0.000"
        l.HeaderRowRange.WrapText = True
        l.HeaderRowRange.EntireRow.RowHeight = 30
    for j in range(1, len(HIST_COLS) + 1):
        ws.Columns(j).ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 8
    ws.Columns("C").ColumnWidth = 30
    ws.Tab.Color = rgb(GREY)


# ---------------------------------------------------------------- weekly trend
def build_trend(wb, sheets, wip):
    ws = wb.Worksheets("Weekly trend")
    dark(ws, grid(ws, "W"))
    header(ws, "S", '="Weekly trend  ·  8 weeks to "&TEXT(WeekEnding,"d mmm yyyy")',
           "Read from the History table on the Snapshots sheet. Each Refresh All adds the new week.", "Weekly trend", sheets[0])
    n = len(sheets)
    # data grid below the charts
    g_title = 46
    section(ws, f"B{g_title}", "Chart data (from History)", 11)
    hrow = g_title + 1
    g0, g1 = hrow + 1, hrow + 8
    heads = ["Week ending"] + [f'="CPI "&Dashboard!B{D_FIRST + j}' for j in range(n)] + [f'="SPI "&Dashboard!B{D_FIRST + j}' for j in range(n)]
    heads += ["CPI on track", "CPI at risk", "SPI on track", "SPI at risk", "Actual cost", "Earned value", "Overrun so far",
              "Forecast", "Budget", "Forecast overrun", "Crew planned", "Crew on site", "Billed", "Cash received", "Cash paid"]
    cols = {}
    for j, h in enumerate(heads):
        c = col(2 + j)
        put(ws.Range(f"{c}{hrow}"), h)
        if not h.startswith("="):
            cols[h] = c
    last = col(1 + len(heads))
    table_head(ws.Range(f"B{hrow}:{last}{hrow}"))
    ws.Rows(hrow).RowHeight = 32
    for i in range(8):
        r = g0 + i
        ws.Range(f"B{r}").Formula = f"=WeekEnding-7*{7 - i}"
        for j in range(n):
            for k, fld in enumerate(("CPI", "SPI")):
                c = col(3 + j + k * n)
                ws.Range(f"{c}{r}").Formula = (f'=IFERROR(1/(1/SUMIFS(History[{fld}],History[Week ending],$B{r},'
                                               f'History[Sheet],Dashboard!$B${D_FIRST + j})),NA())')
        ws.Range(f"{cols['CPI on track']}{r}").Formula = "=GreenCPI"
        ws.Range(f"{cols['CPI at risk']}{r}").Formula = "=AmberCPI"
        ws.Range(f"{cols['SPI on track']}{r}").Formula = "=GreenSPI"
        ws.Range(f"{cols['SPI at risk']}{r}").Formula = "=AmberSPI"
        for h, fld in (("Actual cost", "Actual"), ("Earned value", "Earned"), ("Forecast", "Forecast at completion"),
                       ("Budget", "Budget"), ("Crew planned", "Planned heads"), ("Crew on site", "Actual heads"),
                       ("Billed", "Billed"), ("Cash received", "Cash received"), ("Cash paid", "Cash paid")):
            ws.Range(f"{cols[h]}{r}").Formula = f"=IFERROR(1/(1/SUMIFS(History[{fld}],History[Week ending],$B{r})),NA())"
        ws.Range(f"{cols['Overrun so far']}{r}").Formula = f"={cols['Actual cost']}{r}-{cols['Earned value']}{r}"
        ws.Range(f"{cols['Forecast overrun']}{r}").Formula = f"={cols['Forecast']}{r}-{cols['Budget']}{r}"
    table_body(ws.Range(f"B{g0}:{last}{g1}"))
    ws.Range(f"B{g0}:{last}{g1}").Font.Size = 8.5
    ws.Range(f"B{g0}:B{g1}").NumberFormat = "d mmm"
    ws.Range(f"C{g0}:{col(2 + 2 * n + 4)}{g1}").NumberFormat = "0.00"
    ws.Range(f"{cols['Actual cost']}{g0}:{last}{g1}").NumberFormat = '#,##0,"k"'
    ws.Range(f"{cols['Crew planned']}{g0}:{cols['Crew on site']}{g1}").NumberFormat = "0"
    hide_errors(ws.Range(f"C{g0}:{last}{g1}"), f"C{g0}")

    def x():
        return ws.Range(f"B{g0}:B{g1}")

    for r in range(6, 44):
        ws.Rows(r).RowHeight = 15
    ws.Rows(24).RowHeight = 10
    for idx, (fld, green, amber, title, rng) in enumerate((("CPI", "CPI on track", "CPI at risk", "CPI by project, week by week", "B6:J23"),
                                                           ("SPI", "SPI on track", "SPI at risk", "SPI by project, week by week", "K6:S23"))):
        ch = add_chart(ws, rng, 4)
        for j in range(n):
            c = col(3 + j + idx * n)
            add_series(ch, f"='Weekly trend'!${c}${hrow}", ws.Range(f"{c}{g0}:{c}{g1}"), x(), SERIES[j % len(SERIES)],
                       line=True, weight=2.5, marker=True)
        add_series(ch, "On track", ws.Range(f"{cols[green]}{g0}:{cols[green]}{g1}"), x(), GOOD, line=True, weight=1.25, dash=4)
        add_series(ch, "At risk", ws.Range(f"{cols[amber]}{g0}:{cols[amber]}{g1}"), x(), BAD, line=True, weight=1.25, dash=4)
        dark_chart(ch, title)
        ch.Axes(1).CategoryType = 2
        ch.Axes(1).TickLabels.NumberFormat = "d mmm"
        ch.Axes(2).MinimumScale = 0.7 if fld == "SPI" else 0.8
        ch.Axes(2).MaximumScale = 1.1
        ch.Axes(2).TickLabels.NumberFormat = "0.00"
    ch = add_chart(ws, "B25:J42", 51)
    add_series(ch, "Overrun so far (actual - earned)", ws.Range(f"{cols['Overrun so far']}{g0}:{cols['Overrun so far']}{g1}"), x(), COPPER)
    add_series(ch, "Forecast overrun (forecast - budget)", ws.Range(f"{cols['Forecast overrun']}{g0}:{cols['Forecast overrun']}{g1}"),
               x(), BAD, kind=65, line=True, weight=2.5, marker=True)
    dark_chart(ch, "Portfolio overrun, week by week")
    ch.Axes(1).CategoryType = 2
    ch.Axes(1).TickLabels.NumberFormat = "d mmm"
    ch.Axes(2).TickLabels.NumberFormat = '#,##0,"k"'
    ch.ChartGroups(1).GapWidth = 70
    ch = add_chart(ws, "K25:S42", 51)
    add_series(ch, "Crew planned", ws.Range(f"{cols['Crew planned']}{g0}:{cols['Crew planned']}{g1}"), x(), GREY)
    add_series(ch, "Crew on site", ws.Range(f"{cols['Crew on site']}{g0}:{cols['Crew on site']}{g1}"), x(), SKY)
    dark_chart(ch, "Crew on site vs plan, all projects (heads)")
    ch.Axes(1).CategoryType = 2
    ch.Axes(1).TickLabels.NumberFormat = "d mmm"
    ch.ChartGroups(1).GapWidth = 60
    ws.Range("B43").Value = ("A falling CPI line = the project spends more than the work done is worth; SPI below 1 = behind the baseline. "
                             "Overrun so far = actual cost minus earned value; forecast overrun = forecast at completion minus budget.")
    ws.Range("B43").Font.Italic = True; ws.Range("B43").Font.Color = rgb(MUTED); ws.Range("B43").Font.Size = 8
    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$S$43"
    return dict(rows=(g0, g1), header=hrow, cols=cols, wip_chart_row=wip["chart"])


def build_howto(wb):
    ws = wb.Worksheets("How to use")
    ws.Columns("A").ColumnWidth = 125
    text = [
        ("How to use this workbook", True), ("", False),
        ("Every week", True),
        ("1. Each PM opens their project sheet (between 'PM start' and 'PM end') and updates the gold cells: committed, "
         "actual to date and % complete per cost code; billed, cash received and paid, crew on site; notes. Planned % is looked up in the Baseline sheet.", False),
        ("2. Settings: set Week ending to this Friday.", False),
        ("3. Data > Refresh All. The report sheets are live anyway; Refresh All saves this week's snapshot in the History table "
         "(Snapshots sheet). Refreshing again in the same week replaces that week's rows, it never adds them twice.", False),
        ("4. Dashboard: all checks must say PASS. Use the buttons at the top to move between the report sheets.", False),
        ("", False),
        ("Report sheets", True),
        ("Dashboard: KPI cards with CPI / SPI gauges, one row per project, cost-code heatmap (switch CPI / SPI), charts, checks, PM notes.", False),
        ("Project view: pick a project in the gold cell; cards, waterfall from budget to forecast by cost code, earned value detail.", False),
        ("S-curve: pick a project; planned value from the Baseline sheet against earned value and actual cost from History; crew chart.", False),
        ("COs & risks: type change orders and risks in the two tables. Approved change orders add their price to the contract and "
         "their cost estimate to the forecast; pending ones are shown as upside. The 5x5 matrix lists the risk IDs per cell.", False),
        ("WIP & cash: percentage-of-completion WIP schedule (over / under billing), retainage and cash received vs paid.", False),
        ("", False),
        ("Add a project", True),
        ("Right-click 'Template' > Move or Copy > tick Create a copy > place it before 'PM end'. Rename the sheet to the project ID, "
         "fill in the gold cells, then type the sheet name in the next empty row of the Projects table on the Dashboard "
         f"({SLOTS} rows). Add the project's weeks to the Baseline table for its S-curve.", False),
        ("Close a project: move its sheet after 'PM end' (it stays in the history) and clear its row on the Dashboard.", False),
        ("", False),
        ("What the numbers mean", True),
        ("Planned value (PV) = budget x planned %; earned value (EV) = budget x % complete; actual cost (AC) as booked.", False),
        ("CPI = EV / AC (cost), SPI = EV / PV (schedule). 1.00 = on plan; CPI 0.90 = every $1.00 of work is costing $1.11.", False),
        ("Forecast at completion = actual / % complete for each cost code (the budget for lines not started) + approved change orders. "
         "Shown next to it: EAC CPI method = BAC / CPI and EAC CPI x SPI method = AC + (BAC - EV) / (CPI x SPI).", False),
        ("TCPI = (BAC - EV) / (BAC - AC): the CPI needed on the remaining work to finish on budget.", False),
        ("Margin at completion = (revised contract - forecast) / revised contract. Forecast finish = planned weeks / SPI.", False),
        ("WIP: % complete = actual / forecast cost; earned revenue = revised contract x % complete; over / (under) billed = billed - earned revenue.", False),
        ("Status: cost On track if CPI >= 0.98, Watch if >= 0.92; schedule On track if SPI >= 0.95, Watch if >= 0.90 (Settings).", False),
        ("", False),
        ("Checks", True),
        ("The Dashboard adds up every sheet between the bookends with 3D sums and compares it with its own table, so a project "
         "sheet that was copied in but not added to the Dashboard shows FAIL. It also flags unknown cost codes, % outside 0-100%, "
         "planned value that does not match the baseline, change orders for unknown projects, risk scores outside 1-5, "
         "billing above the contract and a week saved twice.", False),
        ("Formulas are locked on the PM sheets (no password: Review > Unprotect Sheet). No macros: works in Excel 2016 or later.", False),
    ]
    for i, (t, bold) in enumerate(text, 1):
        ws.Cells(i, 1).Value = t
        ws.Cells(i, 1).Font.Bold = bold
        ws.Cells(i, 1).WrapText = True
    ws.Range("A1").Font.Size = 16


def main():
    global SEP
    if OUT.exists():
        OUT.unlink()
    projects = read_csv("projects.csv")
    lines = read_csv(f"pm-week-{WEEK.isoformat()}.csv")
    sheets = [p["Sheet"] for p in projects]
    order = FRONT + sheets + TAIL
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        SEP = xl.International[4]
        print("list separator:", SEP)
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < len(order):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, title in enumerate(order, 1):
            wb.Worksheets(i).Name = title
        wb.SaveAs(str(OUT), 51)
        xl.Calculation = -4135  # manual while building
        build_settings(wb)
        codes = build_codes(wb)
        n_weeks = build_baseline(wb)
        for p in projects:
            build_pm(wb.Worksheets(p["Sheet"]), codes, sheets[0], p, [l for l in lines if l["sheet"] == p["Sheet"]])
        build_pm(wb.Worksheets("Template"), codes, sheets[0])
        wb.Worksheets("Template").Tab.Color = rgb(COPPER)
        bookend(wb.Worksheets("PM start"), "PM start: project sheets begin after this sheet")
        bookend(wb.Worksheets("PM end"), "PM end: project sheets end before this sheet")
        wb.Names.Add("PickProject", "='Project view'!$C$6")
        wb.Names.Add("PickCurve", "='S-curve'!$C$6")
        build_cos_risks(wb, sheets)  # the ChangeOrders and Risks tables: the Rollup formulas read them
        ro = build_rollup(wb)
        gauges = build_gauge_cells(ro)
        T = Texts(ro)
        info = build_dashboard(wb, sheets, T, gauges)
        build_project_view(wb, sheets, T, gauges)
        build_cos_cards(wb, T)
        wip = build_wip(wb, sheets, T)
        xl.Calculation = -4105
        xl.Calculate()
        t0 = time.time()
        build_snapshots(wb, sheets)  # the History table: the S-curve, the trend and two checks read it
        d = wb.Worksheets("Dashboard")
        for cell, f in info.pop("late"):
            d.Range(cell).Formula = f
        print(f"history query loaded in {time.time() - t0:.1f}s")
        build_scurve(wb, sheets, T, n_weeks)
        trend = build_trend(wb, sheets, wip)
        wip_cash_chart(wb, trend)
        build_howto(wb)
        wb.Worksheets("Dashboard").Tab.Color = rgb(GOLD)
        xl.CalculateFull()
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            if ws.Name not in sheets and ws.Name != "Template":
                ws.Range("A1").Select()
            xl.ActiveWindow.ScrollRow = 1
            xl.ActiveWindow.ScrollColumn = 1
        for name in sheets + ["Template"]:
            protect(wb.Worksheets(name))
        wb.Worksheets("Dashboard").Activate()
        wb.Save()
        print("checks:", [c.Value for c in wb.Names("CheckResults").RefersToRange])
        print("status:", [d.Range(f"R{r}").Value for r in range(D_FIRST, D_FIRST + len(sheets))],
              [d.Range(f"S{r}").Value for r in range(D_FIRST, D_FIRST + len(sheets))])
        print("layout:", info)
        print("saved", OUT.name)
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        xl.Quit()


if __name__ == "__main__":
    main()
