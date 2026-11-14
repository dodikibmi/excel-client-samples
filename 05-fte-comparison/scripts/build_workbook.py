"""Builds FTE-Comparison.xlsx with real Excel (COM).

Power Query cleans an HR roster export and a Finance position budget, compares them person by person and fund center
by fund center, and three PivotTables (HR, budget, budget minus HR) summarise FTE by fund center x functional area.
Every discrepancy is listed with full details on its own tab, with a next step.

Run with a Python that has pywin32:  python scripts/build_workbook.py
The workbook has no macros; the client only needs Data -> Refresh All.
"""
import time
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "FTE-Comparison.xlsx"

# ---------------------------------------------------------------- Power Query (M)
Q = {}
Q["Setting"] = """(name as text) as text =>
    // Reads one value from a named cell on the Settings sheet (DataFolder, HRFile, BudgetFile)
    Text.From(Excel.CurrentWorkbook(){[Name = name]}[Content]{0}[Column1])"""

Q["HR_Raw"] = """let
    Path    = Setting("DataFolder") & Setting("HRFile"),
    Source  = Csv.Document(File.Contents(Path), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Headers"""

Q["Budget_Raw"] = """let
    Path    = Setting("DataFolder") & Setting("BudgetFile"),
    Source  = Csv.Document(File.Contents(Path), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Headers"""

Q["HR_Clean"] = """let
    Source   = HR_Raw,
    // 1. the HR export sometimes repeats a row: drop exact duplicates
    Distinct = Table.Distinct(Source),
    Digits   = (t as text) as text => Text.Select(t, {"0".."9"}),
    // 2. one employee key both systems agree on: E01042 and 1042 are the same person
    Key      = Table.AddColumn(Distinct, "Emp key", each Text.From(Number.From(Digits([Employee ID]))), type text),
    // 3. FC-210, fc210, "FC 210 " -> FC210 ; "80%", " 80 %" -> 0.80
    Clean    = Table.AddColumn(Key, "Fund center", each "FC" & Digits([Fund Center]), type text),
    Fte      = Table.TransformColumns(Clean, {{"FTE", each Number.Round(Number.FromText(Digits(_)) / 100, 2), type number}}),
    Named    = Table.TransformColumns(Fte, {{"Employee Name", Text.Trim, type text}, {"Functional Area", Text.Trim, type text},
                   {"Status", Text.Trim, type text}, {"Employee ID", Text.Trim, type text}}),
    Line     = Table.AddIndexColumn(Named, "HR line", 1, 1, Int64.Type),
    Result   = Table.SelectColumns(Table.RenameColumns(Line, {{"Employee Name", "Employee"}, {"Functional Area", "Functional area"},
                   {"Fund Center", "Fund center (as typed)"}}),
                   {"HR line", "Employee ID", "Emp key", "Employee", "Fund center (as typed)", "Fund center", "Functional area", "FTE", "Status"})
in
    Result"""

Q["HR_Active"] = """let
    // terminated people stay in the HR export; they are not current headcount
    Source = Table.SelectRows(HR_Clean, each [Status] <> "Terminated")
in
    Source"""

Q["Budget_Clean"] = """let
    Source  = Budget_Raw,
    Digits  = (t as text) as text => Text.Select(t, {"0".."9"}),
    // HR spells the functional areas; the budget types them in any case, with stray spaces
    Areas   = List.Buffer(List.Distinct(HR_Clean[Functional area])),
    Area    = (t as text) as text => let s = Text.Trim(t), m = List.Select(Areas, each Text.Upper(_) = Text.Upper(s))
                                     in if List.IsEmpty(m) then s else m{0},
    Key     = Table.AddColumn(Source, "Emp key", each if Digits([Emp No]) = "" then null else Text.From(Number.From(Digits([Emp No]))), type text),
    Fc      = Table.AddColumn(Key, "Fund center", each "FC" & Digits([Cost Center]), type text),
    Ar      = Table.AddColumn(Fc, "Functional area", each Area([Function]), type text),
    Fte     = Table.AddColumn(Ar, "FTE", each Number.Round(Number.FromText(Text.Trim([Budgeted FTE]), "en-US"), 2), type number),
    Named   = Table.TransformColumns(Fte, {{"Employee", Text.Trim, type text}, {"Position", Text.Trim, type text}}),
    Line    = Table.AddIndexColumn(Named, "Budget line", 1, 1, Int64.Type),
    Result  = Table.SelectColumns(Table.RenameColumns(Line, {{"Function", "Function (as typed)"}}),
                  {"Budget line", "Position", "Emp key", "Employee", "Fund center", "Function (as typed)", "Functional area", "FTE"})
in
    Result"""

Q["Compare"] = """let
    H   = Table.RenameColumns(Table.SelectColumns(HR_Active, {"HR line", "Employee ID", "Emp key", "Employee", "Fund center", "Functional area", "FTE"}),
              {{"Employee", "HR name"}, {"Fund center", "HR fund center"}, {"Functional area", "HR functional area"}, {"FTE", "HR FTE"}}),
    B   = Table.RenameColumns(Table.SelectColumns(Budget_Clean, {"Budget line", "Position", "Emp key", "Employee", "Fund center", "Functional area", "FTE"}),
              {{"Emp key", "B key"}, {"Employee", "Budget name"}, {"Fund center", "Budget fund center"},
               {"Functional area", "Budget functional area"}, {"FTE", "Budget FTE"}}),
    // 1. same person in the same fund center
    J       = Table.Join(H, {"Emp key", "HR fund center"}, B, {"B key", "Budget fund center"}, JoinKind.FullOuter),
    Both    = Table.SelectRows(J, each [HR line] <> null and [Budget line] <> null),
    HOnly   = Table.SelectColumns(Table.SelectRows(J, each [Budget line] = null), Table.ColumnNames(H)),
    BOnly   = Table.SelectColumns(Table.SelectRows(J, each [HR line] = null), Table.ColumnNames(B)),
    // 2. same person left over on both sides = booked to a different fund center
    Moved   = Table.Join(HOnly, {"Emp key"}, Table.SelectRows(BOnly, each [B key] <> null), {"B key"}, JoinKind.Inner),
    MovedK  = List.Buffer(Moved[Emp key]),
    HRest   = Table.SelectRows(HOnly, each not List.Contains(MovedK, [Emp key])),
    BRest   = Table.SelectRows(BOnly, each [B key] = null or not List.Contains(MovedK, [B key])),
    All     = Table.Combine({Both, Moved, HRest, BRest}),
    Gone    = List.Buffer(Table.SelectRows(HR_Clean, each [Status] = "Terminated")[Emp key]),
    Status  = Table.AddColumn(All, "Status", each
                  if [HR line] = null then "Only in budget"
                  else if [Budget line] = null then "Only in HR roster"
                  else if [HR fund center] <> [Budget fund center] then "Fund center differs"
                  else if [HR functional area] <> [Budget functional area] then "Functional area differs"
                  else if Number.Round([Budget FTE] - [HR FTE], 2) <> 0 then "FTE differs"
                  else "Matched", type text),
    Diff    = Table.AddColumn(Status, "FTE difference", each
                  Number.Round((if [Budget FTE] = null then 0 else [Budget FTE]) - (if [HR FTE] = null then 0 else [HR FTE]), 2), type number),
    Person  = Table.AddColumn(Diff, "Employee", each if [HR name] <> null then [HR name] else [Budget name], type text),
    Id      = Table.AddColumn(Person, "Emp ID", each
                  if [Employee ID] <> null then [Employee ID] else if [B key] <> null then "E" & Text.PadStart([B key], 5, "0") else "", type text),
    F2      = (n as nullable number) as text => if n = null then "0.00" else Number.ToText(n, "0.00", "en-US"),
    Action  = Table.AddColumn(Id, "What to do", each
                  if [Status] = "Matched" then "OK"
                  else if [Status] = "FTE differs" then "HR shows " & F2([HR FTE]) & " FTE, budget " & F2([Budget FTE]) & ": confirm the approved FTE and correct one side"
                  else if [Status] = "Functional area differs" then "HR: " & [HR functional area] & ", budget: " & [Budget functional area] & ": agree the functional area"
                  else if [Status] = "Fund center differs" then "HR books the person to " & [HR fund center] & ", the budget to " & [Budget fund center] & ": confirm which fund center pays"
                  else if [Status] = "Only in HR roster" then "On the HR roster but not funded: add a budgeted position or check the hire"
                  else if [B key] = null then "Vacant position " & [Position] & ": confirm recruitment or release the budget"
                  else if List.Contains(Gone, [B key]) then "Terminated in HR but still funded: remove from the budget"
                  else "In the budget but not on the HR roster: check the person is still employed", type text),
    Order   = Table.AddColumn(Action, "Sort", each
                  List.PositionOf({"FTE differs", "Functional area differs", "Fund center differs", "Only in HR roster", "Only in budget", "Matched"}, [Status]), Int64.Type),
    Sorted  = Table.Sort(Order, {{"Sort", Order.Ascending}, {"Employee", Order.Ascending}}),
    Result  = Table.SelectColumns(Sorted, {"Status", "Emp ID", "Employee", "HR line", "HR fund center", "HR functional area", "HR FTE",
                  "Budget line", "Position", "Budget fund center", "Budget functional area", "Budget FTE", "FTE difference", "What to do"})
in
    Result"""

Q["Discrepancies"] = """let
    Source = Table.SelectRows(Compare, each [Status] <> "Matched")
in
    Source"""

Q["FTE_Long"] = """let
    // one long table feeds all three PivotTables, so HR, budget and the difference can never disagree
    Hr  = Table.AddColumn(Table.AddColumn(Table.SelectColumns(HR_Active, {"Fund center", "Functional area", "FTE"}),
              "Source", each "HR roster", type text), "Budget minus HR", each -[FTE], type number),
    Bu  = Table.AddColumn(Table.AddColumn(Table.SelectColumns(Budget_Clean, {"Fund center", "Functional area", "FTE"}),
              "Source", each "Finance budget", type text), "Budget minus HR", each [FTE], type number)
in
    Table.Combine({Hr, Bu})"""

Q["Cleaning_Log"] = """let
    Digits    = (t as text) as text => Text.Select(t, {"0".."9"}),
    HrRows    = Table.RowCount(HR_Raw),
    HrKept    = Table.RowCount(HR_Clean),
    FcOdd     = Table.RowCount(Table.SelectRows(HR_Clean, each [#"Fund center (as typed)"] <> [Fund center])),
    Gone      = Table.SelectRows(HR_Clean, each [Status] = "Terminated"),
    Active    = Table.RowCount(HR_Active),
    BuRows    = Table.RowCount(Budget_Raw),
    Vacant    = Table.RowCount(Table.SelectRows(Budget_Clean, each [Emp key] = null)),
    AreaFix   = Table.RowCount(Table.SelectRows(Budget_Clean, each [#"Function (as typed)"] <> [Functional area])),
    Log = #table(type table [Source = text, Check = text, Rows = Int64.Type, #"What was done" = text], {
        {"HR roster", "Rows in the export file", HrRows, "Read as text first, so nothing is lost on import"},
        {"HR roster", "Exact duplicate rows", HrRows - HrKept, "Removed (the export repeated them)"},
        {"HR roster", "Employee IDs like E01042", HrKept, "Turned into one key the budget also uses (1042)"},
        {"HR roster", "Fund centers typed differently (fc210, FC-210, FC 210)", FcOdd, "Normalised to FC210"},
        {"HR roster", "FTE written as a percentage (80%)", HrKept, "Converted to 0.80"},
        {"HR roster", "Terminated employees still in the export", Table.RowCount(Gone), "Left out of headcount; used to explain budget lines"},
        {"HR roster", "Active and on-leave rows compared", Active, "Compared with the budget"},
        {"Finance budget", "Rows in the export file", BuRows, "Read as text first, so nothing is lost on import"},
        {"Finance budget", "Cost center without prefix (210)", BuRows, "Written as FC210, like HR"},
        {"Finance budget", "Functional area in a different case or with spaces", AreaFix, "Spelled the way HR spells it"},
        {"Finance budget", "Positions without an employee (vacant)", Vacant, "Kept; reported as Only in budget"}})
in
    Log"""

LOADED = [("Summary", None), ("Pivot comparison", None), ("Discrepancies", "Discrepancies"), ("Comparison", "Compare"),
          ("Cleaning log", "Cleaning_Log"), ("Settings", None), ("How to use", None)]
STATUS_COLORS = {"Matched": (226, 239, 218), "FTE differs": (252, 228, 214), "Functional area differs": (252, 228, 214),
                 "Fund center differs": (252, 228, 214), "Only in HR roster": (255, 242, 204), "Only in budget": (221, 235, 247)}
NAVY, TEAL = (31, 58, 90), (0, 128, 128)
GREEN = ((198, 239, 206), (0, 97, 0))
RED = ((255, 199, 206), (156, 0, 6))
YELLOW = (255, 255, 153)


def rgb(r, g, b):
    return r + g * 256 + b * 65536


def conn_string(name):
    return f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=""'


def load_query(wb, ws, name, table_name, at="A1"):
    lo = ws.ListObjects.Add(0, conn_string(name), None, 1, ws.Range(at))  # xlSrcExternal, xlYes
    qt = lo.QueryTable
    qt.CommandType = 2  # xlCmdSql
    qt.CommandText = f"SELECT * FROM [{name}]"
    qt.RefreshStyle = 1  # xlInsertDeleteCells
    qt.AdjustColumnWidth = False
    qt.BackgroundQuery = False  # Refresh All waits for it, so the Summary formulas see the new data
    qt.Refresh(False)
    lo.Name = table_name
    lo.TableStyle = "TableStyleMedium2"
    return lo


def heading(ws, cell, text, size=13):
    ws.Range(cell).Value = text
    ws.Range(cell).Font.Size = size
    ws.Range(cell).Font.Bold = True
    ws.Range(cell).Font.Color = rgb(*NAVY)


def pass_fail(rng, ok_value='="PASS"'):
    fc = rng.FormatConditions.Add(1, 3, ok_value)
    fc.Interior.Color = rgb(*GREEN[0]); fc.Font.Color = rgb(*GREEN[1])
    fc = rng.FormatConditions.Add(1, 4, ok_value)
    fc.Interior.Color = rgb(*RED[0]); fc.Font.Color = rgb(*RED[1])


def title_band(ws, last_col, title_formula, sub):
    ws.Range(f"A1:{last_col}2").Interior.Color = rgb(*NAVY)
    ws.Range("B1").Formula = title_formula
    ws.Range("B1").Font.Size = 18; ws.Range("B1").Font.Bold = True; ws.Range("B1").Font.Color = rgb(255, 255, 255)
    ws.Range("B2").Value = sub
    ws.Range("B2").Font.Color = rgb(200, 215, 230)
    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 20


# ---------------------------------------------------------------- PivotTables
def build_pivots(wb):
    ws = wb.Worksheets("Pivot comparison")
    ws.Columns("A").ColumnWidth = 2
    title_band(ws, "W", '="FTE by fund center and functional area   |   "&Period',
               "Three PivotTables on one Power Query source: HR roster, Finance budget, and budget minus HR. Red = the two sources disagree.")
    conn = wb.Connections.Add2("Query - FTE_Long", "FTE from both sources, one row per person and fund center",
                               conn_string("FTE_Long"), "SELECT * FROM [FTE_Long]", 2, False, False)
    conn.OLEDBConnection.BackgroundQuery = False
    pc = wb.PivotCaches().Create(2, conn, 6)  # xlExternal, version 2016+
    blocks = [("B", "HRPivot", "HR roster", "FTE", "HR roster: active FTE"),
              ("J", "BudgetPivot", "Finance budget", "FTE", "Finance budget: budgeted FTE"),
              ("R", "DiffPivot", None, "Budget minus HR", "Budget minus HR (+ = budget higher)")]
    for col, name, source, field, label in blocks:
        heading(ws, f"{col}4", label)
        pt = pc.CreatePivotTable(ws.Range(f"{col}8"), name)
        pt.RowAxisLayout(1)  # tabular
        pt.PivotFields("Fund center").Orientation = 1  # row
        pt.PivotFields("Functional area").Orientation = 2  # column
        pf = pt.PivotFields("Source")
        pf.Orientation = 3  # page
        if source:
            pf.CurrentPage = source
        df = pt.AddDataField(pt.PivotFields(field), "Total FTE" if source else "FTE difference", -4157)  # xlSum
        df.NumberFormat = "0.00" if source else '+0.00;-0.00;"–"'
        pt.ShowDrillIndicators = False
        pt.TableStyle2 = "PivotStyleMedium2" if source else "PivotStyleMedium9"
        pt.CompactLayoutRowHeader = "Fund center"
        pt.HasAutoFormat = False  # keep the column widths below on refresh
        pt.NullString = ""
        pt.DisplayNullString = True
    for c in "BJR":
        ws.Columns(c).ColumnWidth = 13
    for c in ["C", "D", "E", "F", "G", "K", "L", "M", "N", "O", "S", "T", "U", "V", "W"]:
        ws.Columns(c).ColumnWidth = 15.5
    for c in "HIPQ":
        ws.Columns(c).ColumnWidth = 2
    # red where the sources disagree (any non-zero cell of the difference pivot, totals included)
    diff = ws.PivotTables("DiffPivot")
    fc = diff.DataBodyRange.FormatConditions.Add(1, 4, "=0")  # cell value <> 0 (empty cells count as 0)
    fc.Interior.Color = rgb(*RED[0]); fc.Font.Color = rgb(*RED[1])
    fc.ScopeType = 2  # xlDataFieldScope: the rule follows the pivot when it grows on refresh
    ws.Range("B18").Value = ("How to read it: each cell is the FTE in that fund center and functional area. "
                             "The third table is the second minus the first; every red cell is explained line by line on the Discrepancies tab.")
    ws.Range("B18").Font.Italic = True
    ws.Range("B18").Font.Color = rgb(90, 100, 115)
    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$W$18"
    return ws


# ---------------------------------------------------------------- Summary
def build_summary(wb):
    ws = wb.Worksheets("Summary")
    widths = {"A": 2, "B": 30, "C": 11, "D": 11, "E": 11, "F": 13, "G": 3, "H": 12, "I": 12, "J": 12, "K": 12, "L": 12, "M": 12}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    title_band(ws, "M", '="FTE comparison   |   Riverside Community Health   |   "&Period',
               "HR roster vs Finance position budget, compared in Power Query. Fictional data. New month: new CSVs in the data folder, Data > Refresh All.")

    # Result by status (rows 12-20)
    heading(ws, "B12", "Result by status")
    for c, t in {"B": "Status", "C": "Lines", "D": "HR FTE", "E": "Budget FTE", "F": "Budget - HR"}.items():
        h = ws.Range(f"{c}13")
        h.Value = t; h.Font.Bold = True; h.Interior.Color = rgb(*NAVY); h.Font.Color = rgb(255, 255, 255)
        if c != "B":
            h.HorizontalAlignment = -4152
    for i, st in enumerate(STATUS_COLORS, 14):
        ws.Range(f"B{i}").Value = st
        ws.Range(f"B{i}").Interior.Color = rgb(*STATUS_COLORS[st])
        ws.Range(f"C{i}").Formula = f"=COUNTIFS(Compare[Status],B{i})"
        ws.Range(f"D{i}").Formula = f"=SUMIFS(Compare[HR FTE],Compare[Status],B{i})"
        ws.Range(f"E{i}").Formula = f"=SUMIFS(Compare[Budget FTE],Compare[Status],B{i})"
        ws.Range(f"F{i}").Formula = f"=SUMIFS(Compare[FTE difference],Compare[Status],B{i})"
    ws.Range("B20").Value = "Total"
    for c in "CDEF":
        ws.Range(f"{c}20").Formula = f"=SUM({c}14:{c}19)"
    for c in "BCDEF":
        ws.Range(f"{c}20").Font.Bold = True
        ws.Range(f"{c}20").Borders(8).LineStyle = 1
    ws.Range("C14:C20").NumberFormat = "0"
    ws.Range("D14:E20").NumberFormat = "0.00"
    ws.Range("F14:F20").NumberFormat = '+0.00;[Red]-0.00;"–"'

    # Data checks (rows 23-29)
    heading(ws, "B23", "Data checks")
    log = "INDEX(CleaningLog[Rows],{})".format
    checks = [
        ("No HR row lost: rows in file = duplicates removed + terminated + compared",
         f'=IF({log(1)}={log(2)}+{log(6)}+COUNT(Compare[HR line]),"PASS","FAIL")'),
        ("Every active HR line compared exactly once",
         '=IF(SUMPRODUCT((Compare[HR line]<>"")/COUNTIF(Compare[HR line],Compare[HR line]&""))=COUNT(Compare[HR line]),"PASS","FAIL")'),
        ("Every budget line compared exactly once",
         f'=IF(AND(SUMPRODUCT((Compare[Budget line]<>"")/COUNTIF(Compare[Budget line],Compare[Budget line]&""))=COUNT(Compare[Budget line]),COUNT(Compare[Budget line])={log(8)}),"PASS","FAIL")'),
        ("HR PivotTable total = HR FTE in the comparison",
         "=IF(ROUND(GETPIVOTDATA(\"Total FTE\",'Pivot comparison'!$B$8)-D20,2)=0,\"PASS\",\"FAIL\")"),
        ("Budget PivotTable total = budget FTE in the comparison",
         "=IF(ROUND(GETPIVOTDATA(\"Total FTE\",'Pivot comparison'!$J$8)-E20,2)=0,\"PASS\",\"FAIL\")"),
        ("Difference PivotTable total = sum of the line differences",
         "=IF(ROUND(GETPIVOTDATA(\"FTE difference\",'Pivot comparison'!$R$8)-F20,2)=0,\"PASS\",\"FAIL\")"),
    ]
    for i, (label, f) in enumerate(checks, 24):
        ws.Range(f"B{i}").Value = label
        cell = ws.Range(f"M{i}")
        cell.Formula = f
        cell.Font.Bold = True
        cell.HorizontalAlignment = -4108
        pass_fail(cell)

    # KPI cards (rows 4-9)
    cards = [("B", "B", "HR ROSTER FTE", "=D20", "0.00", '=COUNT(Compare[HR line])&" active lines"'),
             ("C", "E", "FINANCE BUDGET FTE", "=E20", "0.00", '=COUNT(Compare[Budget line])&" budget lines"'),
             ("F", "I", "BUDGET MINUS HR", "=F20", '+0.00;-0.00;0.00', '="net FTE difference"'),
             ("J", "M", "LINES TO REVIEW", "=C20-C14", "0", '="each with full details and a next step"')]
    for a, b, label, formula, fmt, note in cards:
        box = ws.Range(f"{a}4:{b}9")
        box.Interior.Color = rgb(247, 249, 252)
        for edge in (7, 8, 9, 10):
            box.Borders(edge).LineStyle = 1
            box.Borders(edge).Color = rgb(210, 218, 228)
        ws.Range(f"{a}4:{b}4").Interior.Color = rgb(*TEAL)
        for rows, val, size, bold, color in (("5:5", label, 9, True, (90, 100, 115)), ("6:7", formula, 26, True, NAVY),
                                             ("8:8", note, 9, False, (110, 110, 110))):
            r1, r2 = rows.split(":")
            rng = ws.Range(f"{a}{r1}:{b}{r2}")
            if rng.Cells.Count > 1:
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

    # Chart: lines to review by type
    top, left = ws.Range("H12").Top, ws.Range("H12").Left
    co = ws.ChartObjects().Add(left, top, ws.Range("H12:M12").Width, ws.Range("H12:H20").Height)
    ch = co.Chart
    ch.ChartType = 57  # xlBarClustered
    ser = ch.SeriesCollection().NewSeries()
    ser.Name = "Lines"
    ser.Values = ws.Range("C15:C19")
    ser.XValues = ws.Range("B15:B19")
    ch.HasLegend = False
    ser.Format.Fill.ForeColor.RGB = rgb(*TEAL)
    ser.HasDataLabels = True
    ch.Axes(1).ReversePlotOrder = True
    ch.Axes(2).HasMajorGridlines = False
    ch.Axes(2).Delete()
    ch.HasTitle = True
    ch.ChartTitle.Text = "Lines to review, by type"
    ch.ChartTitle.Font.Size = 13
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Font.Color = rgb(*NAVY)
    ch.ChartTitle.Left = 4
    ch.ChartArea.Format.Line.Visible = False

    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$M$29"


def format_tables(wb):
    widths = {"Status": 21, "Emp ID": 9, "Employee": 18, "HR line": 6, "HR fund center": 9, "HR functional area": 16,
              "HR FTE": 7, "Budget line": 7, "Position": 8, "Budget fund center": 9, "Budget functional area": 16,
              "Budget FTE": 8, "FTE difference": 9, "What to do": 70}
    for title in ("Comparison", "Discrepancies"):
        ws = wb.Worksheets(title)
        lo = ws.ListObjects(1)
        for c in ("HR FTE", "Budget FTE"):
            lo.ListColumns(c).DataBodyRange.NumberFormat = "0.00"
        lo.ListColumns("FTE difference").DataBodyRange.NumberFormat = '+0.00;[Red]-0.00;"–"'
        for i in range(1, lo.ListColumns.Count + 1):
            ws.Columns(i).ColumnWidth = widths.get(lo.ListColumns(i).Name, 12)
        lo.HeaderRowRange.WrapText = True
        rng = lo.ListColumns("Status").DataBodyRange
        for text, color in STATUS_COLORS.items():
            fc = rng.FormatConditions.Add(2, None, f'=$A2="{text}"')
            fc.Interior.Color = rgb(*color)
        ws.Activate()
        ws.Range("A2").Select()
        wb.Application.ActiveWindow.FreezePanes = True
    ws = wb.Worksheets("Cleaning log")
    for c, w in zip("ABCD", (15, 56, 7, 62)):
        ws.Columns(c).ColumnWidth = w


def build_settings(wb):
    st = wb.Worksheets("Settings")
    st.Range("A1").Value = "Settings"
    st.Range("A1").Font.Size = 16; st.Range("A1").Font.Bold = True
    rows = [("Data folder", '=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"data\\"', "DataFolder",
             "Folder with the two CSV exports. Default: the data folder next to this workbook. Type a path to override."),
            ("HR roster file", "hr-roster-2026-10.csv", "HRFile", "HR system export: Employee ID, Employee Name, Fund Center, Functional Area, FTE, Status"),
            ("Finance budget file", "finance-budget-2026-10.csv", "BudgetFile", "Position budget export: Position, Emp No, Employee, Cost Center, Function, Budgeted FTE"),
            ("Period", "October 2026", "Period", "Shown in the report titles")]
    st.Range("A3:C3").Value = ("Setting", "Value", "Note")
    st.Range("A3:C3").Font.Bold = True
    for r, (label, val, nm, note) in enumerate(rows, 4):
        st.Cells(r, 1).Value = label
        if val.startswith("="):
            st.Cells(r, 2).Formula2 = val
        else:
            st.Cells(r, 2).NumberFormat = "@"
            st.Cells(r, 2).Value = val
        st.Cells(r, 3).Value = note
        st.Cells(r, 2).Interior.Color = rgb(*YELLOW)
        wb.Names.Add(nm, f"=Settings!$B${r}")
    st.Columns("A").ColumnWidth = 22; st.Columns("B").ColumnWidth = 48; st.Columns("C").ColumnWidth = 90


def build_howto(wb):
    ws = wb.Worksheets("How to use")
    ws.Columns("A").ColumnWidth = 120
    text = [
        ("How to use this workbook", True), ("", False),
        ("Every month", True),
        ("1. Export the HR roster and the Finance position budget as CSV into the data folder (see Settings).", False),
        ("2. On Settings, type the two file names.", False),
        ("3. Data > Refresh All. The comparison, the three PivotTables and the Summary all update. No macros.", False),
        ("4. Summary: all data checks must say PASS. Then work through the Discrepancies tab.", False), ("", False),
        ("What each status means", True),
        ("Matched: same person, same fund center, same functional area, same FTE in both sources.", False),
        ("FTE differs: same person and fund center, different FTE (e.g. HR 1.00, budget 0.80).", False),
        ("Functional area differs: same person and fund center, booked to a different functional area.", False),
        ("Fund center differs: the person is in both sources but under a different fund center.", False),
        ("Only in HR roster: an active employee with no budgeted position.", False),
        ("Only in budget: a vacant position, someone terminated in HR, or a name HR does not know.", False), ("", False),
        ("How the two sources are lined up", True),
        ("Employee IDs: E01042 (HR) and 1042 (budget) become the same key. Fund centers: fc210, FC-210, 210 all become FC210.", False),
        ("FTE: 80% (HR) and 0.8 (budget) are both 0.80. Functional areas are spelled the way HR spells them.", False),
        ("Terminated people stay in the HR file but are not counted as headcount; every step is counted on the Cleaning log tab.", False),
        ("The steps live in Power Query (Data > Queries & Connections): HR_Clean, Budget_Clean, Compare, FTE_Long.", False),
    ]
    for i, (t, bold) in enumerate(text, 1):
        ws.Cells(i, 1).Value = t
        ws.Cells(i, 1).Font.Bold = bold
    ws.Range("A1").Font.Size = 16


def main():
    if OUT.exists():
        OUT.unlink()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < len(LOADED):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, (title, _) in enumerate(LOADED, 1):
            wb.Worksheets(i).Name = title
        build_settings(wb)
        wb.SaveAs(str(OUT), 51)  # CELL("filename") needs a saved file
        xl.Calculate()
        for name, formula in Q.items():
            wb.Queries.Add(name, formula)
        t0 = time.time()
        for title, query in LOADED:
            if query:
                load_query(wb, wb.Worksheets(title), query, query.replace("_", ""))
        print(f"queries loaded in {time.time() - t0:.1f}s")
        format_tables(wb)
        build_pivots(wb)
        build_summary(wb)
        build_howto(wb)
        xl.Calculate()
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            ws.Range("A1").Select()
        wb.Worksheets("Summary").Activate()
        wb.Save()
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
