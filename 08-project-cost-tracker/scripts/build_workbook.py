"""Builds Project-Cost-Tracker.xlsx with real Excel (COM).

One input sheet per project (identical layout, filled in by the PM), placed between the 'PM start' and 'PM end' sheets.
The Dashboard rolls them up for the VP: KPI cards, a red/amber/green table per project, earned value and CPI,
budget vs actual vs forecast, field vs shop fabrication progress, the PMs' notes, and PASS/FAIL checks.
Data > Refresh All saves this week's snapshot into a History table (Power Query appends to its own output, so the
history survives without macros); the Weekly trend sheet charts it.

Run with a Python that has pywin32:  python scripts/build_workbook.py      (after scripts/make_data.py)
The workbook has no macros.
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

NAVY, TEAL, ORANGE = (31, 58, 90), (0, 128, 128), (237, 125, 49)
GREEN = ((198, 239, 206), (0, 97, 0))
AMBER = ((255, 235, 156), (156, 87, 0))
RED = ((255, 199, 206), (156, 0, 6))
YELLOW = (255, 255, 204)
MUTED = (90, 100, 115)
SHEETS = ["Dashboard", "Weekly trend", "Snapshots", "PM start"]  # then the PM sheets, then the rest
TAIL = ["PM end", "Template", "Cost codes", "Settings", "How to use"]
MONEY = '#,##0;[Red]-#,##0;"–"'
PCT = '0.0%;-0.0%;"–"'
SLOTS = 8  # project rows on the Dashboard
HIST_COLS = ["Week ending", "Sheet", "Project", "Contract", "Budget", "Committed", "Actual", "Earned", "Field %", "Shop %", "CPI",
             "Forecast at completion"]


def rgb(r, g, b):
    return r + g * 256 + b * 65536


def read_csv(name):
    with open(DATA / name, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def heading(ws, cell, text, size=13):
    ws.Range(cell).Value = text
    ws.Range(cell).Font.Size = size
    ws.Range(cell).Font.Bold = True
    ws.Range(cell).Font.Color = rgb(*NAVY)


def header_row(rng):
    rng.Font.Bold = True
    rng.Interior.Color = rgb(*NAVY)
    rng.Font.Color = rgb(255, 255, 255)
    rng.WrapText = True
    rng.VerticalAlignment = -4108


def pass_fail(rng):
    fc = rng.FormatConditions.Add(1, 3, '="PASS"')
    fc.Interior.Color = rgb(*GREEN[0]); fc.Font.Color = rgb(*GREEN[1])
    fc = rng.FormatConditions.Add(1, 4, '="PASS"')
    fc.Interior.Color = rgb(*RED[0]); fc.Font.Color = rgb(*RED[1])


def status_colors(rng):
    for text, (bg, fg) in (("On track", GREEN), ("Watch", AMBER), ("At risk", RED)):
        fc = rng.FormatConditions.Add(1, 3, f'="{text}"')
        fc.Interior.Color = rgb(*bg); fc.Font.Color = rgb(*fg); fc.Font.Bold = True


def title_band(ws, last_col, title, sub):
    ws.Range(f"A1:{last_col}2").Interior.Color = rgb(*NAVY)
    c = ws.Range("B1")
    if title.startswith("="):
        c.Formula = title
    else:
        c.Value = title
    c.Font.Size = 18; c.Font.Bold = True; c.Font.Color = rgb(255, 255, 255)
    ws.Range("B2").Value = sub
    ws.Range("B2").Font.Color = rgb(200, 215, 230)
    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 20


# NB: conditional-format formulas are read in Excel's local syntax (this PC: ';'), so they avoid commas.
# ---------------------------------------------------------------- settings, cost codes
def build_settings(wb):
    st = wb.Worksheets("Settings")
    st.Range("A1").Value = "Settings"
    st.Range("A1").Font.Size = 16; st.Range("A1").Font.Bold = True
    rows = [("Company", "Brightwire Electrical Contractors", "Company", "Shown in the dashboard title (fictional company)", "@"),
            ("Week ending", WEEK, "WeekEnding", "Change every week, then Data > Refresh All: the snapshot for this date is saved", "d mmm yyyy"),
            ("On track if CPI at least", 0.98, "GreenCPI", "Cost performance index = earned value / actual cost", "0.00"),
            ("Watch if CPI at least", 0.92, "AmberCPI", "Below this a project is At risk", "0.00")]
    st.Range("A3:C3").Value = ("Setting", "Value", "Note")
    st.Range("A3:C3").Font.Bold = True
    for r, (label, val, nm, note, fmt) in enumerate(rows, 4):
        st.Cells(r, 1).Value = label
        st.Cells(r, 2).NumberFormat = fmt
        if isinstance(val, date):  # a serial number: a datetime would be shifted by the PC's time zone
            st.Cells(r, 2).Value2 = (val - date(1899, 12, 30)).days
        else:
            st.Cells(r, 2).Value = val
        st.Cells(r, 2).HorizontalAlignment = -4131
        st.Cells(r, 3).Value = note
        st.Cells(r, 2).Interior.Color = rgb(*YELLOW)
        wb.Names.Add(nm, f"=Settings!$B${r}")
    st.Columns("A").ColumnWidth = 26; st.Columns("B").ColumnWidth = 34; st.Columns("C").ColumnWidth = 80


def build_codes(wb):
    ws = wb.Worksheets("Cost codes")
    codes = read_csv("cost-codes.csv")
    heading(ws, "A1", "Cost codes (the only codes a PM sheet accepts)", 14)
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
    ws.Range("A16").Font.Italic = True; ws.Range("A16").Font.Color = rgb(*MUTED)
    return [c["Cost code"] for c in codes]


# ---------------------------------------------------------------- one PM sheet
FIRST, LAST = 15, 24  # cost code rows
TOT = 25


def build_pm(ws, codes, proj=None, lines=None):
    widths = {"A": 2, "B": 16, "C": 34, "D": 16, "E": 13, "F": 13, "G": 13, "H": 11, "I": 13, "J": 13, "K": 7, "L": 14, "M": 14}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    title_band(ws, "M", '="Project cost report   |   "&$C$4&"   |   "&$C$5',
               "Filled in by the project manager every week. Yellow cells are inputs; everything else is calculated.")
    labels = [("B4", "Project ID"), ("B5", "Project"), ("B6", "Project manager"), ("B7", "Contract value"), ("B8", "Week ending")]
    for cell, t in labels:
        ws.Range(cell).Value = t
        ws.Range(cell).Font.Color = rgb(*MUTED)
    ws.Range("C4").NumberFormat = "@"
    ws.Range("C4").Value = proj["Sheet"] if proj else "P-0000"
    ws.Range("C5").Value = proj["Project"] if proj else "New project name"
    ws.Range("C6").Value = proj["PM"] if proj else "PM name"
    ws.Range("C7").Value = float(proj["Contract value"]) if proj else 0
    ws.Range("C7").NumberFormat = "#,##0"
    ws.Range("C8").Formula = "=WeekEnding"
    ws.Range("C8").NumberFormat = "d mmm yyyy"
    ws.Range("C4:C8").HorizontalAlignment = -4131
    ws.Range("C4:C7").Interior.Color = rgb(*YELLOW)
    ws.Range("C4:C5").Font.Bold = True

    box = [("F4", "Field % complete", "H4",
            f'=IF(SUMPRODUCT((D{FIRST}:D{LAST}<>"Shop fabrication")*E{FIRST}:E{LAST})=0,"",'
            f'SUMPRODUCT((D{FIRST}:D{LAST}<>"Shop fabrication")*E{FIRST}:E{LAST}*H{FIRST}:H{LAST})'
            f'/SUMPRODUCT((D{FIRST}:D{LAST}<>"Shop fabrication")*E{FIRST}:E{LAST}))', PCT),
           ("F5", "Shop fabrication % complete", "H5",
            f'=IF(SUMIF(D{FIRST}:D{LAST},"Shop fabrication",E{FIRST}:E{LAST})=0,"",'
            f'SUMPRODUCT((D{FIRST}:D{LAST}="Shop fabrication")*E{FIRST}:E{LAST}*H{FIRST}:H{LAST})'
            f'/SUMIF(D{FIRST}:D{LAST},"Shop fabrication",E{FIRST}:E{LAST}))', PCT),
           ("F6", "Cost performance (CPI)", "H6", f"=K{TOT}", "0.00"),
           ("F7", "Forecast at completion", "H7", f"=L{TOT}", "#,##0"),
           ("F8", "Projected margin", "H8", f"=IF(C7=0,\"\",C7-H7)", MONEY)]
    for lc, label, vc, f, fmt in box:
        ws.Range(lc).Value = label
        ws.Range(lc).Font.Color = rgb(*MUTED)
        ws.Range(vc).Formula = f
        ws.Range(vc).NumberFormat = fmt
        ws.Range(vc).Font.Bold = True
        ws.Range(vc).HorizontalAlignment = -4152
    ws.Range("I8").Formula = '=IF(OR(C7=0,H8=""),"",H8/C7)'
    ws.Range("I8").NumberFormat = '0.0%;[Red]-0.0%'
    ws.Range("F4:I8").Interior.Color = rgb(242, 246, 250)

    ws.Range("K4").Value = "Sheet checks"
    ws.Range("K4").Font.Bold = True; ws.Range("K4").Font.Color = rgb(*NAVY)
    ws.Range("K5").Value = "Unknown cost codes"
    ws.Range("M5").Formula = f'=SUMPRODUCT((B{FIRST}:B{LAST}<>"")*(COUNTIF(CodeList,B{FIRST}:B{LAST})=0))'
    ws.Range("K6").Value = "% complete outside 0-100%"
    ws.Range("M6").Formula = f'=COUNTIF(H{FIRST}:H{LAST},"<0")+COUNTIF(H{FIRST}:H{LAST},">1")'
    ws.Range("K7").Value = "Actual above committed"
    ws.Range("M7").Formula = f"=SUMPRODUCT(--(G{FIRST}:G{LAST}>F{FIRST}:F{LAST}))"
    for c in ("M5", "M6", "M7"):
        ws.Range(c).HorizontalAlignment = -4108
        fc = ws.Range(c).FormatConditions.Add(1, 3, "=0")
        fc.Interior.Color = rgb(*GREEN[0]); fc.Font.Color = rgb(*GREEN[1])
        fc = ws.Range(c).FormatConditions.Add(1, 5, "=0")
        fc.Interior.Color = rgb(*RED[0]); fc.Font.Color = rgb(*RED[1])
    ws.Range("K5:K7").Font.Color = rgb(*MUTED)

    ws.Range("B10").Value = "This week's notes"
    ws.Range("B10").Font.Bold = True; ws.Range("B10").Font.Color = rgb(*NAVY)
    ws.Range("C10:M12").Merge()
    ws.Range("C10").Value = proj["This week's note"] if proj else ""
    ws.Range("C10:M12").WrapText = True
    ws.Range("C10:M12").VerticalAlignment = -4160
    ws.Range("C10:M12").Interior.Color = rgb(*YELLOW)

    hdr = ["Cost code", "Description", "Category", "Budget", "Committed", "Actual to date", "% complete", "Earned value",
           "Cost variance", "CPI", "Forecast at completion", "Variance at completion"]
    ws.Range("B14:M14").Value = hdr
    header_row(ws.Range("B14:M14"))
    ws.Rows(14).RowHeight = 32
    vals = {r["code"]: r for r in (lines or [])}
    for i, code in enumerate(codes):
        r = FIRST + i
        ws.Range(f"B{r}").NumberFormat = "@"
        ws.Range(f"B{r}").Value = code
        ws.Range(f"C{r}").Formula = f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},CostCodes,2,FALSE),"Unknown code"))'
        ws.Range(f"D{r}").Formula = f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},CostCodes,3,FALSE),"Unknown"))'
        v = vals.get(code)
        if v:
            ws.Range(f"E{r}:H{r}").Value = (float(v["budget"]), float(v["committed"]), float(v["actual"]), float(v["pct"]))
        ws.Range(f"I{r}").Formula = f"=E{r}*H{r}"
        ws.Range(f"J{r}").Formula = f"=I{r}-G{r}"
        ws.Range(f"K{r}").Formula = f'=IF(G{r}>0,I{r}/G{r},"")'
        ws.Range(f"L{r}").Formula = f"=IF(H{r}>0,G{r}/H{r},MAX(E{r},G{r}))"
        ws.Range(f"M{r}").Formula = f"=E{r}-L{r}"
    ws.Range(f"B{TOT}").Value = "Total"
    for c in "EFGIJLM":
        ws.Range(f"{c}{TOT}").Formula = f"=SUM({c}{FIRST}:{c}{LAST})"
    ws.Range(f"K{TOT}").Formula = f'=IF(G{TOT}>0,I{TOT}/G{TOT},"")'
    ws.Range(f"B{TOT}:M{TOT}").Font.Bold = True
    ws.Range(f"B{TOT}:M{TOT}").Borders(8).LineStyle = 1
    ws.Range(f"B{TOT}:M{TOT}").Interior.Color = rgb(242, 246, 250)
    ws.Range(f"E{FIRST}:G{TOT}").NumberFormat = "#,##0"
    ws.Range(f"I{FIRST}:I{TOT}").NumberFormat = "#,##0"
    ws.Range(f"L{FIRST}:L{TOT}").NumberFormat = "#,##0"
    ws.Range(f"J{FIRST}:J{TOT}").NumberFormat = MONEY
    ws.Range(f"M{FIRST}:M{TOT}").NumberFormat = MONEY
    ws.Range(f"H{FIRST}:H{LAST}").NumberFormat = "0.0%"
    ws.Range(f"K{FIRST}:K{TOT}").NumberFormat = "0.00"
    inputs = [ws.Range(f"B{FIRST}:B{LAST}"), ws.Range(f"E{FIRST}:H{LAST}")]
    for rng in inputs:
        rng.Interior.Color = rgb(*YELLOW)
    dv = ws.Range(f"B{FIRST}:B{LAST}").Validation
    dv.Delete()
    dv.Add(3, 1, 1, "=CodeList")
    dv.ErrorMessage = "Pick a cost code from the list (Cost codes sheet)."
    dv = ws.Range(f"H{FIRST}:H{LAST}").Validation
    dv.Delete()
    dv.Add(2, 1, 1, "0", "1")  # decimal between 0 and 1
    dv.ErrorMessage = "% complete is between 0% and 100%."
    # CPI per line: red below Watch, amber below On track
    k = ws.Range(f"K{FIRST}:K{TOT}")
    fc = k.FormatConditions.Add(2, None, f'=(K{FIRST}<>"")*(K{FIRST}<AmberCPI)')
    fc.Font.Color = rgb(*RED[1]); fc.Interior.Color = rgb(*RED[0])
    fc = k.FormatConditions.Add(2, None, f'=(K{FIRST}<>"")*(K{FIRST}<GreenCPI)')
    fc.Font.Color = rgb(*AMBER[1]); fc.Interior.Color = rgb(*AMBER[0])
    ws.Range("B27").Value = ("Cost variance = earned value - actual. Earned value = budget x % complete. CPI = earned value / actual (below 1.00 = spending more than the work done is worth). "
                             "Forecast at completion = actual / % complete (budget if the line has not started).")
    ws.Range("B27").Font.Italic = True; ws.Range("B27").Font.Color = rgb(*MUTED)
    # lock formulas, leave inputs open (no password)
    ws.Cells.Locked = True
    for rng in inputs + [ws.Range("C4:C7"), ws.Range("C10:M12")]:
        rng.Locked = False
    ws.Activate()
    ws.Range("A15").Select()
    ws.Application.ActiveWindow.FreezePanes = True
    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$M$27"


def protect(ws):
    ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True, AllowFormattingColumns=True, AllowFormattingRows=True)


def bookend(ws, text):
    ws.Tab.Color = rgb(166, 166, 166)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 110
    heading(ws, "B2", text, 14)
    lines = ["Every sheet between 'PM start' and 'PM end' is a project and is added up on the Dashboard (3D sums).",
             "Add a project: right-click 'Template' > Move or Copy > tick Create a copy > put it before 'PM end'.",
             "Rename the copy to the project ID (e.g. P-2404), fill in the yellow cells, then type the same ID in the next empty "
             "row of the Projects table on the Dashboard.",
             "If a sheet sits between the bookends but is not on the Dashboard, the Dashboard check says FAIL."]
    for i, t in enumerate(lines, 4):
        ws.Range(f"B{i}").Value = t
        ws.Range(f"B{i}").Font.Color = rgb(*MUTED)


# ---------------------------------------------------------------- dashboard
D_FIRST = 13
D_LAST = D_FIRST + SLOTS - 1
D_TOT = D_LAST + 1


def ind(r, cell):
    return f'=IF($B{r}="","",INDIRECT("\'"&$B{r}&"\'!{cell}"))'


def build_dashboard(wb, sheets):
    ws = wb.Worksheets("Dashboard")
    widths = {"A": 2, "B": 9, "C": 31, "D": 16, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 7, "K": 8, "L": 8,
              "M": 12, "N": 12, "O": 9, "P": 10}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    title_band(ws, "P", '="Project cost dashboard   |   "&Company',
               "Every PM sheet between 'PM start' and 'PM end' rolls up here. Fictional data. Weekly: PMs update their sheets, then Data > Refresh All.")
    ws.Range("N1").Value = "Week ending"
    ws.Range("N1").Font.Color = rgb(200, 215, 230)
    ws.Range("N1").HorizontalAlignment = -4152
    ws.Range("N1").VerticalAlignment = -4108
    ws.Range("O1:P1").Merge()
    ws.Range("O1").Formula = "=WeekEnding"
    ws.Range("O1").NumberFormat = "d mmm yyyy"
    ws.Range("O1").Font.Size = 14; ws.Range("O1").Font.Bold = True; ws.Range("O1").Font.Color = rgb(255, 255, 255)
    ws.Range("O1").HorizontalAlignment = -4131
    ws.Range("O1").VerticalAlignment = -4108

    # Projects table
    heading(ws, "B11", "Projects: one row per PM sheet (type the sheet name in column B)")
    hdr = ["Sheet", "Project", "PM", "Contract value", "Budget", "Committed", "Actual to date", "Earned value", "CPI",
           "Field %", "Shop fab %", "Forecast at completion", "Variance at completion", "Margin", "Status"]
    ws.Range(f"B12:P12").Value = hdr
    header_row(ws.Range("B12:P12"))
    ws.Rows(12).RowHeight = 32
    for i in range(SLOTS):
        r = D_FIRST + i
        ws.Range(f"B{r}").NumberFormat = "@"
        if i < len(sheets):
            ws.Range(f"B{r}").Value = sheets[i]
        ws.Range(f"B{r}").Interior.Color = rgb(*YELLOW)
        for col, cell in (("C", "C5"), ("D", "C6"), ("E", "C7"), ("F", f"E{TOT}"), ("G", f"F{TOT}"), ("H", f"G{TOT}"),
                          ("I", f"I{TOT}"), ("K", "H4"), ("L", "H5"), ("M", f"L{TOT}")):
            ws.Range(f"{col}{r}").Formula = ind(r, cell)
        ws.Range(f"J{r}").Formula = f'=IF(N(H{r})>0,I{r}/H{r},"")'
        ws.Range(f"N{r}").Formula = f'=IF($B{r}="","",F{r}-M{r})'
        ws.Range(f"O{r}").Formula = f'=IF(N(E{r})>0,(E{r}-M{r})/E{r},"")'
        ws.Range(f"P{r}").Formula = f'=IF(J{r}="","",IF(J{r}>=GreenCPI,"On track",IF(J{r}>=AmberCPI,"Watch","At risk")))'
    t = D_TOT
    ws.Range(f"B{t}").Value = "Total"
    for c in "EFGHIMN":
        ws.Range(f"{c}{t}").Formula = f"=SUM({c}{D_FIRST}:{c}{D_LAST})"
    ws.Range(f"J{t}").Formula = f'=IF(H{t}>0,I{t}/H{t},"")'
    ws.Range(f"O{t}").Formula = f'=IF(E{t}>0,(E{t}-M{t})/E{t},"")'
    ws.Range(f"P{t}").Formula = f'=IF(J{t}="","",IF(J{t}>=GreenCPI,"On track",IF(J{t}>=AmberCPI,"Watch","At risk")))'
    ws.Range(f"B{t}:P{t}").Font.Bold = True
    ws.Range(f"B{t}:P{t}").Borders(8).LineStyle = 1
    ws.Range(f"B{t}:P{t}").Interior.Color = rgb(242, 246, 250)
    ws.Range(f"E{D_FIRST}:I{t}").NumberFormat = "#,##0"
    ws.Range(f"M{D_FIRST}:M{t}").NumberFormat = "#,##0"
    ws.Range(f"N{D_FIRST}:N{t}").NumberFormat = MONEY
    ws.Range(f"J{D_FIRST}:J{t}").NumberFormat = "0.00"
    ws.Range(f"K{D_FIRST}:L{t}").NumberFormat = "0%"
    ws.Range(f"O{D_FIRST}:O{t}").NumberFormat = '0.0%;[Red]-0.0%'
    ws.Range(f"P{D_FIRST}:P{t}").HorizontalAlignment = -4108
    status_colors(ws.Range(f"P{D_FIRST}:P{t}"))
    # shop fabrication behind the field by 10 points or more: amber
    fc = ws.Range(f"L{D_FIRST}:L{D_LAST}").FormatConditions.Add(2, None, f'=ISNUMBER(L{D_FIRST})*((K{D_FIRST}-L{D_FIRST})>=0.1)')
    fc.Interior.Color = rgb(*AMBER[0]); fc.Font.Color = rgb(*AMBER[1]); fc.Font.Bold = True
    ws.Range(f"B{t + 1}").Value = ("Earned value = budget x % complete; CPI = earned / actual. Status: On track if CPI >= 0.98, Watch if >= 0.92, else At risk "
                                   "(Settings). Amber shop % = shop fabrication 10+ points behind the field.")
    ws.Range(f"B{t + 1}").Font.Italic = True; ws.Range(f"B{t + 1}").Font.Color = rgb(*MUTED); ws.Range(f"B{t + 1}").Font.Size = 9

    # KPI cards (rows 4-9)
    cards = [("B", "C", "CONTRACT VALUE", f"=E{t}", "#,##0", f'=COUNTA(B{D_FIRST}:B{D_LAST})&" active projects"'),
             ("D", "F", "BUDGET (COST)", f"=F{t}", "#,##0", f'="committed "&TEXT(G{t},"#,##0")'),
             ("G", "I", "ACTUAL COST TO DATE", f"=H{t}", "#,##0", f'="earned value "&TEXT(I{t},"#,##0")&"  ·  CPI "&TEXT(J{t},"0.00")'),
             ("J", "M", "FORECAST AT COMPLETION", f"=M{t}", "#,##0", f'="margin at completion "&TEXT(O{t},"0.0%")'),
             ("N", "P", "VARIANCE AT COMPLETION", f"=N{t}", '#,##0;[Red]-#,##0', '="budget minus forecast"')]
    for a, b, label, formula, fmt, note in cards:
        box = ws.Range(f"{a}4:{b}9")
        box.Interior.Color = rgb(247, 249, 252)
        for edge in (7, 8, 9, 10):
            box.Borders(edge).LineStyle = 1
            box.Borders(edge).Color = rgb(210, 218, 228)
        ws.Range(f"{a}4:{b}4").Interior.Color = rgb(*TEAL)
        for rows, val, size, bold, color in (("5:5", label, 9, True, MUTED), ("6:7", formula, 24, True, NAVY),
                                             ("8:8", note, 9, False, (110, 110, 110))):
            r1, r2 = rows.split(":")
            rng = ws.Range(f"{a}{r1}:{b}{r2}")
            rng.Merge()
            cell = ws.Range(f"{a}{r1}")
            if val.startswith("="):
                cell.Formula = val
            else:
                cell.Value = val
            cell.Font.Size = size; cell.Font.Bold = bold; cell.Font.Color = rgb(*color)
            rng.HorizontalAlignment = -4108
            rng.VerticalAlignment = -4108
        ws.Range(f"{a}6").NumberFormat = fmt
    ws.Rows(4).RowHeight = 5

    # Checks (rows 24-31)
    c0 = t + 3  # 24
    heading(ws, f"B{c0}", "Checks")
    checks = [
        ("Every PM sheet between 'PM start' and 'PM end' is on this dashboard",
         f"=IF(SHEETS('PM start:PM end'!A1)-2=COUNTA(B{D_FIRST}:B{D_LAST}),\"PASS\",\"FAIL\")"),
        ("Dashboard budget = sum of the budget on every PM sheet",
         f"=IF(ROUND(SUM('PM start:PM end'!E{TOT})-F{t},2)=0,\"PASS\",\"FAIL\")"),
        ("Dashboard actual cost = sum of the actual cost on every PM sheet",
         f"=IF(ROUND(SUM('PM start:PM end'!G{TOT})-H{t},2)=0,\"PASS\",\"FAIL\")"),
        ("No unknown cost codes on any PM sheet",
         "=IF(SUM('PM start:PM end'!M5)=0,\"PASS\",\"FAIL\")"),
        ("% complete between 0% and 100% on every line",
         "=IF(SUM('PM start:PM end'!M6)=0,\"PASS\",\"FAIL\")"),
        ("This week's snapshot saved, once per project",
         f"=IF(COUNTIFS(History[Week ending],WeekEnding)=COUNTA(B{D_FIRST}:B{D_LAST}),\"PASS\",\"FAIL\")"),
        ("History: no project saved twice for the same week",
         "=IF(SUMPRODUCT(1/COUNTIFS(History[Week ending],History[Week ending],History[Sheet],History[Sheet]))=ROWS(History[Sheet]),\"PASS\",\"FAIL\")"),
    ]
    late = []  # formulas that use the History table are written after it exists
    for i, (label, f) in enumerate(checks, c0 + 1):
        ws.Range(f"B{i}").Value = label
        cell = ws.Range(f"P{i}")
        if "History[" in f:
            late.append((f"P{i}", f))
        else:
            cell.Formula = f
        cell.Font.Bold = True
        cell.HorizontalAlignment = -4108
        pass_fail(cell)
    c_end = c0 + len(checks)  # 31
    ws.Range(f"I{c0 + 1}").Value = "Lines where actual is above committed (check the POs):"
    ws.Range(f"I{c0 + 1}").Font.Color = rgb(*MUTED)
    ws.Range(f"N{c0 + 1}").Formula = "=SUM('PM start:PM end'!M7)"
    ws.Range(f"N{c0 + 1}").HorizontalAlignment = -4108

    # Notes from the PMs
    n0 = c_end + 2  # 33
    heading(ws, f"B{n0}", "This week from the PMs")
    for i in range(SLOTS):
        r = n0 + 1 + i
        src = D_FIRST + i
        ws.Range(f"B{r}").Formula = f'=IF($B${src}="","",$B${src})'
        ws.Range(f"B{r}").Font.Bold = True
        ws.Range(f"B{r}").VerticalAlignment = -4160
        ws.Range(f"C{r}:P{r}").Merge()
        ws.Range(f"C{r}").Formula = f'=IF($B${src}="","",INDIRECT("\'"&$B${src}&"\'!C10"))'
        ws.Range(f"C{r}:P{r}").WrapText = True
        ws.Range(f"C{r}:P{r}").VerticalAlignment = -4160
        ws.Rows(r).RowHeight = 30 if i < len(sheets) else 3
    n_end = n0 + SLOTS

    # Named ranges for the charts: grow with the Projects table
    cnt = f"COUNTA(Dashboard!$B${D_FIRST}:$B${D_LAST})"
    for nm, col in (("PrjIds", "B"), ("PrjBudget", "F"), ("PrjActual", "H"), ("PrjFAC", "M"), ("PrjField", "K"), ("PrjShop", "L")):
        sep = wb.Application.International[4]  # names are parsed in the local list separator on this PC
        n = wb.Names.Add(nm, f"=Dashboard!${col}${D_FIRST}")
        n.RefersToLocal = f"=OFFSET(Dashboard!${col}${D_FIRST},0,0,MAX(1,{cnt}),1)".replace(",", sep)
    book = wb.Name
    ch_top = n_end + 2
    rng = ws.Range(f"B{ch_top}:H{ch_top + 17}")
    co = ws.ChartObjects().Add(rng.Left, rng.Top, ws.Range(f"B{ch_top}:H{ch_top}").Width, rng.Height)
    ch = co.Chart
    ch.ChartType = 51  # clustered column
    for name, ref, color in (("Budget", "PrjBudget", (166, 180, 196)), ("Actual to date", "PrjActual", TEAL),
                             ("Forecast at completion", "PrjFAC", ORANGE)):
        s = ch.SeriesCollection().NewSeries()
        s.Formula = f"=SERIES(\"{name}\",'{book}'!PrjIds,'{book}'!{ref},{ch.SeriesCollection().Count})"
        s.Format.Fill.ForeColor.RGB = rgb(*color)
    ch.HasTitle = True
    ch.ChartTitle.Text = "Budget, actual and forecast at completion"
    ch.HasLegend = True
    ch.Legend.Position = -4107  # bottom
    ch.Axes(2).TickLabels.NumberFormat = '#,##0,"k"'
    ch.ChartGroups(1).GapWidth = 80
    style_chart(ch)

    rng2 = ws.Range(f"I{ch_top}:P{ch_top + 17}")
    co = ws.ChartObjects().Add(rng2.Left + 6, rng2.Top, ws.Range(f"I{ch_top}:P{ch_top}").Width - 6, rng2.Height)
    ch = co.Chart
    ch.ChartType = 57  # clustered bar
    for name, ref, color in (("Field % complete", "PrjField", TEAL), ("Shop fabrication % complete", "PrjShop", ORANGE)):
        s = ch.SeriesCollection().NewSeries()
        s.Formula = f"=SERIES(\"{name}\",'{book}'!PrjIds,'{book}'!{ref},{ch.SeriesCollection().Count})"
        s.Format.Fill.ForeColor.RGB = rgb(*color)
        s.HasDataLabels = True
        s.DataLabels().NumberFormat = "0%"
    ch.HasTitle = True
    ch.ChartTitle.Text = "Field vs shop fabrication progress"
    ch.HasLegend = True
    ch.Legend.Position = -4107
    ch.Axes(1).ReversePlotOrder = True
    ch.Axes(2).MinimumScale = 0
    ch.Axes(2).MaximumScale = 1
    ch.Axes(2).TickLabels.NumberFormat = "0%"
    style_chart(ch)
    end_row = ch_top + 17
    ps = ws.PageSetup
    ps.Orientation = 1; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = f"$A$1:$P${end_row}"
    return dict(checks=(c0 + 1, c_end), notes=n0, charts=ch_top, end=end_row, late=late)


def style_chart(ch):
    ch.ChartTitle.Font.Size = 13
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Font.Color = rgb(*NAVY)
    ch.ChartArea.Format.Line.Visible = False
    try:
        ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = rgb(225, 229, 235)
    except Exception:
        pass


# ---------------------------------------------------------------- snapshots (Power Query history)
def m_value(v, col):
    if col == "Week ending":
        d = date.fromisoformat(v)
        return f"#date({d.year},{d.month},{d.day})"
    if col in ("Sheet", "Project"):
        return '"' + v.replace('"', '""') + '"'
    return v


TYPES = ('{{"Week ending", type date}, {"Sheet", type text}, {"Project", type text}, {"Contract", type number}, '
         '{"Budget", type number}, {"Committed", type number}, {"Actual", type number}, {"Earned", type number}, '
         '{"Field %", type number}, {"Shop %", type number}, {"CPI", type number}, {"Forecast at completion", type number}}')

CURRENT = f"""
    // this week: one row per project from the ThisWeek table (formulas that read the Dashboard)
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
    keys = ["week", "sheet", "project", "contract", "budget", "committed", "actual", "earned", "field", "shop", "cpi", "fac"]
    rows = ",\n        ".join("{" + ", ".join(m_value(r[k], c) for k, c in zip(keys, HIST_COLS)) + "}" for r in seed)
    cols = ", ".join(f'#"{c}" = {"date" if c == "Week ending" else "text" if c in ("Sheet", "Project") else "number"}' for c in HIST_COLS)
    return (f"let{CURRENT}\n    // first build only: the seven earlier weekly snapshots (fictional)\n"
            f"    Old      = #table(type table [{cols}], {{\n        {rows}}}),{TAIL_M}")


HISTORY_M = (f"let{CURRENT}\n    // the history saved so far = this query's own output table on the Snapshots sheet\n"
             f"    Old      = Table.TransformColumnTypes(Excel.CurrentWorkbook(){{[Name = \"History\"]}}[Content], {TYPES}),{TAIL_M}")


def conn_string(name):
    return f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=""'


def build_snapshots(wb):
    ws = wb.Worksheets("Snapshots")
    ws.Columns("A").ColumnWidth = 12
    title_band(ws, "L", "Weekly snapshots",
               "Data > Refresh All copies 'This week' into History. Same week again = replaced, never doubled. No macros.")
    heading(ws, "A4", "This week (live from the Dashboard)")
    ws.Range("A5:L5").Value = HIST_COLS
    src = {"Sheet": "B", "Project": "C", "Contract": "E", "Budget": "F", "Committed": "G", "Actual": "H", "Earned": "I",
           "Field %": "K", "Shop %": "L", "CPI": "J", "Forecast at completion": "M"}
    for i in range(SLOTS):
        r = 6 + i
        d = D_FIRST + i
        for j, col in enumerate(HIST_COLS, 1):
            cell = ws.Cells(r, j)
            if col == "Week ending":
                cell.Formula = f'=IF(Dashboard!$B${d}="","",WeekEnding)'
            else:
                cell.Formula = f'=IF(Dashboard!$B${d}="","",Dashboard!{src[col]}{d})'
    lo = ws.ListObjects.Add(1, ws.Range(f"A5:L{5 + SLOTS}"), None, 1)
    lo.Name = "ThisWeek"
    lo.TableStyle = "TableStyleLight9"
    heading(ws, f"A{8 + SLOTS}", "History (saved by Refresh All, one row per project and week)")
    wb.Queries.Add("History", seed_query())
    lo2 = ws.ListObjects.Add(0, conn_string("History"), None, 1, ws.Range(f"A{9 + SLOTS}"))
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
    # from now on the query appends to its own table
    wb.Queries("History").Formula = HISTORY_M
    qt.Refresh(False)
    for l in (lo, lo2):
        l.ListColumns("Week ending").Range.NumberFormat = "d mmm yyyy"
        for c in ("Contract", "Budget", "Committed", "Actual", "Earned", "Forecast at completion"):
            l.ListColumns(c).DataBodyRange.NumberFormat = "#,##0"
        for c in ("Field %", "Shop %"):
            l.ListColumns(c).DataBodyRange.NumberFormat = "0.0%"
        l.ListColumns("CPI").DataBodyRange.NumberFormat = "0.000"
        l.HeaderRowRange.WrapText = True
    for l in (lo, lo2):
        l.HeaderRowRange.EntireRow.RowHeight = 30
    for c, w in zip("ABCDEFGHIJKL", (12, 8, 32, 12, 12, 12, 12, 12, 8, 8, 7, 13)):
        ws.Columns(c).ColumnWidth = w


# ---------------------------------------------------------------- weekly trend
def build_trend(wb, sheets):
    ws = wb.Worksheets("Weekly trend")
    title_band(ws, "P", '="Weekly trend   |   last 8 weeks to "&TEXT(WeekEnding,"d mmm yyyy")',
               "Read from the History table on the Snapshots sheet. Each Refresh All adds the new week.")
    ws.Range("B1").Formula = '="Weekly trend   |   8 weeks of saved snapshots"'
    ws.Columns("A").ColumnWidth = 2
    heading(ws, "B4", "CPI by project (1.00 = on budget)")
    ws.Range("B5").Value = "Week ending"
    n = len(sheets)
    for j in range(n):
        ws.Cells(5, 3 + j).Formula = f"=Dashboard!B{D_FIRST + j}"
    ws.Cells(5, 3 + n).Value = "On track line"
    ws.Cells(5, 4 + n).Value = "At risk line"
    a_col = 6 + n  # portfolio columns
    pcols = ["Actual cost", "Earned value", "Overrun so far (actual - earned)", "Forecast at completion", "Budget",
             "Forecast overrun (forecast - budget)"]
    for k, h in enumerate(pcols):
        ws.Cells(5, a_col + k).Value = h
    header_row(ws.Range(ws.Cells(5, 2), ws.Cells(5, a_col + len(pcols) - 1)))
    ws.Cells(5, 5 + n).Interior.Pattern = -4142
    ws.Rows(5).RowHeight = 60
    heading(ws, ws.Cells(4, a_col).Address, "Portfolio: overrun so far and forecast overrun")
    for i in range(8):
        r = 6 + i
        ws.Range(f"B{r}").Formula = f"=WeekEnding-7*{7 - i}"
        ws.Range(f"B{r}").NumberFormat = "d mmm"
        for j in range(n):
            c = ws.Cells(r, 3 + j)
            col = c.Address.split("$")[1]
            c.Formula = (f'=IFERROR(1/(1/SUMIFS(History[CPI],History[Week ending],$B{r},History[Sheet],{col}$5)),NA())')
            c.NumberFormat = "0.00"
        ws.Cells(r, 3 + n).Formula = "=GreenCPI"
        ws.Cells(r, 4 + n).Formula = "=AmberCPI"
        L = lambda k: ws.Cells(r, a_col + k).Address.replace("$", "")
        for k, fld in ((0, "Actual"), (1, "Earned"), (3, "Forecast at completion"), (4, "Budget")):
            ws.Cells(r, a_col + k).Formula = f'=IFERROR(1/(1/SUMIFS(History[{fld}],History[Week ending],$B{r})),NA())'
        ws.Cells(r, a_col + 2).Formula = f"={L(0)}-{L(1)}"
        ws.Cells(r, a_col + 5).Formula = f"={L(3)}-{L(4)}"
        ws.Range(ws.Cells(r, a_col), ws.Cells(r, a_col + 5)).NumberFormat = "#,##0"
    for j in range(2, a_col + len(pcols)):
        ws.Columns(j).ColumnWidth = 11
    ws.Columns(5 + n).ColumnWidth = 3
    colors = [TEAL, (192, 0, 0), ORANGE, (112, 48, 160), (0, 112, 192), (84, 130, 53)]

    r1 = ws.Range(ws.Cells(16, 2), ws.Cells(35, 4 + n))
    co = ws.ChartObjects().Add(r1.Left, r1.Top, r1.Width, r1.Height)
    ch = co.Chart
    ch.ChartType = 65  # line with markers
    for j in range(n + 2):
        s = ch.SeriesCollection().NewSeries()
        s.Name = f"='Weekly trend'!{ws.Cells(5, 3 + j).Address}"
        s.Values = ws.Range(ws.Cells(6, 3 + j), ws.Cells(13, 3 + j))
        s.XValues = ws.Range("B6:B13")
        if j < n:
            s.Format.Line.ForeColor.RGB = rgb(*colors[j % len(colors)])
            s.MarkerBackgroundColor = rgb(*colors[j % len(colors)])
            s.MarkerForegroundColor = rgb(*colors[j % len(colors)])
            s.Format.Line.Weight = 2.5
        else:
            s.ChartType = 4  # plain line
            s.Format.Line.ForeColor.RGB = rgb(*(GREEN[1] if j == n else RED[1]))
            s.Format.Line.DashStyle = 4
            s.Format.Line.Weight = 1.25
    ch.Axes(1).CategoryType = 2  # one label per week, not a date scale
    ch.Axes(1).TickLabels.NumberFormat = "d mmm"
    ch.HasTitle = True
    ch.ChartTitle.Text = "CPI by project, week by week"
    ch.Axes(2).MinimumScale = 0.8
    ch.Axes(2).MaximumScale = 1.1
    ch.Axes(2).TickLabels.NumberFormat = "0.00"
    ch.HasLegend = True
    ch.Legend.Position = -4107
    style_chart(ch)

    r2 = ws.Range(ws.Cells(16, a_col), ws.Cells(35, a_col + len(pcols) - 1))
    co = ws.ChartObjects().Add(r2.Left, r2.Top, r2.Width, r2.Height)
    ch = co.Chart
    ch.ChartType = 51  # columns: overrun so far
    s = ch.SeriesCollection().NewSeries()
    s.Name = f"='Weekly trend'!{ws.Cells(5, a_col + 2).Address}"
    s.Values = ws.Range(ws.Cells(6, a_col + 2), ws.Cells(13, a_col + 2))
    s.XValues = ws.Range("B6:B13")
    s.Format.Fill.ForeColor.RGB = rgb(*ORANGE)
    s = ch.SeriesCollection().NewSeries()
    s.Name = f"='Weekly trend'!{ws.Cells(5, a_col + 5).Address}"
    s.Values = ws.Range(ws.Cells(6, a_col + 5), ws.Cells(13, a_col + 5))
    s.XValues = ws.Range("B6:B13")
    s.ChartType = 65
    s.Format.Line.ForeColor.RGB = rgb(192, 0, 0)
    s.MarkerBackgroundColor = rgb(192, 0, 0)
    s.MarkerForegroundColor = rgb(192, 0, 0)
    s.Format.Line.Weight = 2.5
    ch.Axes(1).CategoryType = 2
    ch.Axes(1).TickLabels.NumberFormat = "d mmm"
    ch.HasTitle = True
    ch.ChartTitle.Text = "Portfolio overrun, week by week"
    ch.Axes(2).TickLabels.NumberFormat = '#,##0,"k"'
    ch.HasLegend = True
    ch.Legend.Position = -4107
    style_chart(ch)
    ws.Range("B37").Value = ("A falling CPI line = the project spends more than the work done is worth. Overrun so far = actual cost minus "
                             "earned value; forecast overrun = forecast at completion minus budget.")
    ws.Range("B37").Font.Italic = True; ws.Range("B37").Font.Color = rgb(*MUTED)
    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    last = ws.Cells(37, a_col + len(pcols) - 1).Address
    ps.PrintArea = f"$A$1:{last}"
    ws.Range(f"A1:{ws.Cells(2, a_col + len(pcols) - 1).Address}").Interior.Color = rgb(*NAVY)


def build_howto(wb):
    ws = wb.Worksheets("How to use")
    ws.Columns("A").ColumnWidth = 125
    text = [
        ("How to use this workbook", True), ("", False),
        ("Every week", True),
        ("1. Each PM opens their project sheet (between 'PM start' and 'PM end') and updates the yellow cells: committed, "
         "actual to date and % complete per cost code, plus this week's notes.", False),
        ("2. Settings: set Week ending to this Friday.", False),
        ("3. Data > Refresh All. The Dashboard is live anyway; Refresh All saves this week's snapshot in the History table "
         "(Snapshots sheet). Refreshing again in the same week replaces that week's rows, it never adds them twice.", False),
        ("4. Dashboard: all checks must say PASS. Read the status column, the charts and the PMs' notes; Weekly trend shows the history.", False),
        ("", False),
        ("Add a project", True),
        ("Right-click 'Template' > Move or Copy > tick Create a copy > place it before 'PM end'. Rename the sheet to the project ID, "
         "fill in the yellow cells, then type the sheet name in the next empty row of the Projects table on the Dashboard "
         "(8 rows; insert rows inside the table for more).", False),
        ("Close a project: move its sheet after 'PM end' (it stays in the history) and clear its row on the Dashboard.", False),
        ("", False),
        ("What the numbers mean", True),
        ("Earned value = budget x % complete: what the work done so far should have cost.", False),
        ("CPI (cost performance index) = earned value / actual cost. 1.00 = on budget; 0.90 = every $1.00 of work is costing $1.11.", False),
        ("Forecast at completion = actual / % complete for each cost code (the budget for lines that have not started).", False),
        ("Variance at completion = budget - forecast. Margin at completion = (contract - forecast) / contract.", False),
        ("Field % complete is weighted by budget over all field cost codes; shop fabrication % over the shop fabrication codes, so the two are tracked apart.", False),
        ("Status: On track if CPI >= 0.98, Watch if >= 0.92, otherwise At risk (thresholds on Settings).", False),
        ("", False),
        ("Checks", True),
        ("The Dashboard adds up every sheet between the bookends with 3D sums and compares it with its own table, so a project "
         "sheet that was copied in but not added to the Dashboard shows FAIL. It also flags unknown cost codes, % complete "
         "outside 0-100% and a week saved twice.", False),
        ("Formulas are locked on the PM sheets (no password: Review > Unprotect Sheet). No macros: works in Excel 2016 or later.", False),
    ]
    for i, (t, bold) in enumerate(text, 1):
        ws.Cells(i, 1).Value = t
        ws.Cells(i, 1).Font.Bold = bold
        ws.Cells(i, 1).WrapText = True
    ws.Range("A1").Font.Size = 16


def main():
    if OUT.exists():
        OUT.unlink()
    projects = read_csv("projects.csv")
    lines = read_csv(f"pm-week-{WEEK.isoformat()}.csv")
    sheets = [p["Sheet"] for p in projects]
    order = SHEETS + sheets + TAIL
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < len(order):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, title in enumerate(order, 1):
            wb.Worksheets(i).Name = title
        wb.SaveAs(str(OUT), 51)
        build_settings(wb)
        codes = build_codes(wb)
        for p in projects:
            ws = wb.Worksheets(p["Sheet"])
            build_pm(ws, codes, p, [l for l in lines if l["sheet"] == p["Sheet"]])
            ws.Tab.Color = rgb(*TEAL)
        build_pm(wb.Worksheets("Template"), codes)
        wb.Worksheets("Template").Tab.Color = rgb(*ORANGE)
        bookend(wb.Worksheets("PM start"), "PM start: project sheets begin after this sheet")
        bookend(wb.Worksheets("PM end"), "PM end: project sheets end before this sheet")
        info = build_dashboard(wb, sheets)
        xl.Calculate()
        t0 = time.time()
        build_snapshots(wb)
        for cell, f in info.pop("late"):
            wb.Worksheets("Dashboard").Range(cell).Formula = f
        print(f"history query loaded in {time.time() - t0:.1f}s")
        build_trend(wb, sheets)
        build_howto(wb)
        wb.Worksheets("Dashboard").Tab.Color = rgb(*NAVY)
        xl.Calculate()
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            ws.Range("A1").Select()
            xl.ActiveWindow.ScrollRow = 1
        for name in sheets + ["Template"]:
            protect(wb.Worksheets(name))
        wb.Worksheets("Dashboard").Activate()
        wb.Save()
        d = wb.Worksheets("Dashboard")
        r0, r1 = info["checks"]
        print("checks:", [d.Range(f"P{r}").Value for r in range(r0, r1 + 1)])
        print("status:", [d.Range(f"P{r}").Value for r in range(D_FIRST, D_FIRST + len(sheets))])
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
