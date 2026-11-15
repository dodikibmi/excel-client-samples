"""Builds FTE-Comparison.xlsx with real Excel (COM).

Power Query cleans an HR roster export and a Finance position budget, compares them person by person and fund center
by fund center, and three PivotTables (HR, budget, budget minus HR) summarise FTE by fund center x functional area.
Every discrepancy is listed with full details on its own tab, with a next step.

On top of that, a one-screen Dashboard: KPI cards, a Budget -> HR bridge (waterfall) by discrepancy type, a gap heatmap
fund center x functional area, a donut of reconciliation status, a 3-month trend, FTE by functional area and insight
sentences written by formulas. Four slicers (Month, Fund center, Functional area, Discrepancy type) drive the dashboard
and the three PivotTables together: they all sit on one pivot cache.

Run with a Python that has pywin32:  python scripts/build_workbook.py
The workbook has no macros; the client only needs Data -> Refresh All.
"""
import time
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "FTE-Comparison.xlsx"

# ---------------------------------------------------------------- Power Query (M)
TYPES = ["Matched", "FTE differs", "Area differs", "Fund center moved", "Only in HR", "Vacant post", "Terminated",
         "Unknown to HR"]
M_TYPES = "{" + ", ".join(f'"{t}"' for t in TYPES) + "}"
COMPARE_COLS = ["Status", "Discrepancy type", "Emp ID", "Employee", "HR line", "HR fund center", "HR functional area",
                "HR FTE", "Budget line", "Position", "Budget fund center", "Budget functional area", "Budget FTE",
                "FTE difference", "What to do"]
M_COMPARE_COLS = "{" + ", ".join(f'"{c}"' for c in COMPARE_COLS) + "}"

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

Q["fnHRClean"] = """(Source as table) as table =>
let
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

Q["fnBudgetClean"] = """(Source as table, Areas as list) as table =>
let
    Digits  = (t as text) as text => Text.Select(t, {"0".."9"}),
    // HR spells the functional areas; the budget types them in any case, with stray spaces
    Known   = List.Buffer(Areas),
    Area    = (t as text) as text => let s = Text.Trim(t), m = List.Select(Known, each Text.Upper(_) = Text.Upper(s))
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

Q["fnCompare"] = """(HRAll as table, Budget as table) as table =>
let
    // terminated people stay in the HR export; they are not current headcount
    Active  = Table.SelectRows(HRAll, each [Status] <> "Terminated"),
    H   = Table.RenameColumns(Table.SelectColumns(Active, {"HR line", "Employee ID", "Emp key", "Employee", "Fund center", "Functional area", "FTE"}),
              {{"Employee", "HR name"}, {"Fund center", "HR fund center"}, {"Functional area", "HR functional area"}, {"FTE", "HR FTE"}}),
    B   = Table.RenameColumns(Table.SelectColumns(Budget, {"Budget line", "Position", "Emp key", "Employee", "Fund center", "Functional area", "FTE"}),
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
    Gone    = List.Buffer(Table.SelectRows(HRAll, each [Status] = "Terminated")[Emp key]),
    Status  = Table.AddColumn(All, "Status", each
                  if [HR line] = null then "Only in budget"
                  else if [Budget line] = null then "Only in HR roster"
                  else if [HR fund center] <> [Budget fund center] then "Fund center differs"
                  else if [HR functional area] <> [Budget functional area] then "Functional area differs"
                  else if Number.Round([Budget FTE] - [HR FTE], 2) <> 0 then "FTE differs"
                  else "Matched", type text),
    // 3. a finer type for the dashboard: "Only in budget" splits into vacant, terminated and unknown to HR
    Type    = Table.AddColumn(Status, "Discrepancy type", each
                  if [Status] = "Matched" then "Matched"
                  else if [Status] = "FTE differs" then "FTE differs"
                  else if [Status] = "Functional area differs" then "Area differs"
                  else if [Status] = "Fund center differs" then "Fund center moved"
                  else if [Status] = "Only in HR roster" then "Only in HR"
                  else if [B key] = null then "Vacant post"
                  else if List.Contains(Gone, [B key]) then "Terminated"
                  else "Unknown to HR", type text),
    Diff    = Table.AddColumn(Type, "FTE difference", each
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
                  else if [Discrepancy type] = "Vacant post" then "Vacant position " & [Position] & ": confirm recruitment or release the budget"
                  else if [Discrepancy type] = "Terminated" then "Terminated in HR but still funded: remove from the budget"
                  else "In the budget but not on the HR roster: check the person is still employed", type text),
    Order   = Table.AddColumn(Action, "Sort", each List.PositionOf(""" + M_TYPES + """, [Discrepancy type]), Int64.Type),
    Sorted  = Table.Sort(Order, {{"Sort", Order.Ascending}, {"Employee", Order.Ascending}}),
    Result  = Table.SelectColumns(Sorted, """ + M_COMPARE_COLS + """)
in
    Result"""

Q["HR_Clean"] = """let
    // the steps live in fnHRClean, so every month in the History query is cleaned exactly the same way
    Source = fnHRClean(HR_Raw)
in
    Source"""

Q["HR_Active"] = """let
    // terminated people stay in the HR export; they are not current headcount
    Source = Table.SelectRows(HR_Clean, each [Status] <> "Terminated")
in
    Source"""

Q["Budget_Clean"] = """let
    Source = fnBudgetClean(Budget_Raw, List.Distinct(HR_Clean[Functional area]))
in
    Source"""

Q["Compare"] = """let
    Source = fnCompare(HR_Clean, Budget_Clean)
in
    Source"""

Q["Discrepancies"] = """let
    Source = Table.SelectRows(Compare, each [Status] <> "Matched")
in
    Source"""

Q["History"] = """let
    // every monthly pair in the data folder: hr-roster-YYYY-MM.csv + finance-budget-YYYY-MM.csv, compared the same way
    Path     = Setting("DataFolder"),
    Current  = Text.End(Text.BeforeDelimiter(Setting("HRFile"), ".csv"), 7),
    Files    = Table.Buffer(Table.SelectColumns(Folder.Files(Path), {"Name", "Content"})),
    ReadCsv  = (b as binary) as table => Table.PromoteHeaders(Csv.Document(b, [Delimiter = ",", Encoding = 65001,
                   QuoteStyle = QuoteStyle.Csv]), [PromoteAllScalars = true]),
    Hr       = Table.SelectRows(Files, each Text.StartsWith([Name], "hr-roster-") and Text.EndsWith([Name], ".csv")),
    Months   = Table.AddColumn(Hr, "Month", each Text.End(Text.BeforeDelimiter([Name], ".csv"), 7), type text),
    Pair     = Table.AddColumn(Months, "Budget file", each "finance-budget-" & [Month] & ".csv", type text),
    Joined   = Table.NestedJoin(Pair, {"Budget file"}, Files, {"Name"}, "B", JoinKind.Inner),
    Compared = Table.AddColumn(Joined, "Rows", each
                   let h = fnHRClean(ReadCsv([Content])),
                       b = fnBudgetClean(ReadCsv([B]{0}[Content]), List.Distinct(h[Functional area]))
                   in fnCompare(h, b)),
    Expanded = Table.ExpandTableColumn(Table.SelectColumns(Compared, {"Month", "Rows"}), "Rows", """ + M_COMPARE_COLS + """),
    Snapshot = Table.AddColumn(Expanded, "Snapshot", each if [Month] = Current then "Current" else "Earlier", type text)
in
    Snapshot"""

Q["FTE_Long"] = """let
    // one long table (one row per side of every compared line, every month) feeds every PivotTable and the dashboard,
    // so HR, budget, difference, slicers and charts can never disagree with each other
    Src  = History,
    Hr   = Table.SelectRows(Src, each [HR line] <> null),
    Bu   = Table.SelectRows(Src, each [Budget line] <> null),
    HrR  = Table.AddColumn(Hr, "R", each [Month = [Month], Snapshot = [Snapshot], Source = "HR roster",
               #"Fund center" = [HR fund center], #"Functional area" = [HR functional area],
               #"Discrepancy type" = [Discrepancy type], FTE = [HR FTE], #"Budget minus HR" = -[HR FTE], Line = 1])[R],
    BuR  = Table.AddColumn(Bu, "R", each [Month = [Month], Snapshot = [Snapshot], Source = "Finance budget",
               #"Fund center" = [Budget fund center], #"Functional area" = [Budget functional area],
               #"Discrepancy type" = [Discrepancy type], FTE = [Budget FTE], #"Budget minus HR" = [Budget FTE], Line = 1])[R],
    Long = Table.FromRecords(List.Combine({HrR, BuR}), type table [Month = text, Snapshot = text, Source = text,
               #"Fund center" = text, #"Functional area" = text, #"Discrepancy type" = text, FTE = number,
               #"Budget minus HR" = number, Line = Int64.Type]),
    // slicer items follow this order (data source order): Matched first, then the discrepancy types
    Rank = Table.AddColumn(Long, "Rank", each List.PositionOf(""" + M_TYPES + """, [Discrepancy type]), Int64.Type),
    Sorted = Table.RemoveColumns(Table.Sort(Rank, {{"Rank", Order.Ascending}, {"Month", Order.Ascending}}), {"Rank"})
in
    Sorted"""

Q["Cleaning_Log"] = """let
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

SHEETS = ["Dashboard", "Pivot comparison", "Discrepancies", "Comparison", "Checks", "Cleaning log", "Settings",
          "How to use", "Dash data"]
LOADED = {"Discrepancies": "Discrepancies", "Comparison": "Compare", "Cleaning log": "Cleaning_Log"}
NAV = [("Dashboard", "Dashboard"), ("Pivots", "Pivot comparison"), ("Discrepancies", "Discrepancies"),
       ("Comparison", "Comparison"), ("Checks", "Checks"), ("Cleaning log", "Cleaning log"), ("How to use", "How to use")]
STATUSES = ["Matched", "FTE differs", "Functional area differs", "Fund center differs", "Only in HR roster", "Only in budget"]


# ---------------------------------------------------------------- palette (same as E9)
def hexrgb(h):
    h = h.lstrip("#")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


C = {k: hexrgb(v) for k, v in dict(
    head="0B0F13", page="12171D", card="1A2129", alt="202832", line="2A333D", gold="C9A35B", lgold="E6CF9A",
    ivory="ECE6DA", muted="8E98A3", jade="4FA38C", copper="B5754A", sky="7FA7C9", grey="4A5560", good="5BBF8A",
    bad="E5735F", amber="D9A441", badbg="3A2422", goodbg="1C3529", amberbg="3A3020", skybg="1E2C3A").items()}
SERIF, SANS = "Georgia", "Segoe UI"
BRAND = "  ".join(" ".join(w) for w in "RIVERSIDE COMMUNITY HEALTH".split())


def add_name(wb, name, refers_to):
    """Workbook name. Names.Add reads the formula with the PC's list separator (here ';'), so pass it as local."""
    return wb.Names.Add(Name=name, RefersToLocal=refers_to.replace(",", ";"))


def conn_string(name):
    return f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=""'


# ---------------------------------------------------------------- styles
def add_styles(wb):
    """Dark-gold table, PivotTable and slicer styles."""
    ts = wb.TableStyles.Add("Gold Table")
    ts.ShowAsAvailableTableStyle = True; ts.ShowAsAvailablePivotTableStyle = False
    e = ts.TableStyleElements
    e(0).Interior.Color = C["card"]; e(0).Font.Color = C["ivory"]
    e(1).Interior.Color = C["head"]; e(1).Font.Color = C["gold"]; e(1).Font.Bold = True
    e(5).Interior.Color = C["alt"]

    ps = wb.TableStyles.Add("Gold Pivot")
    ps.ShowAsAvailableTableStyle = False; ps.ShowAsAvailablePivotTableStyle = True
    e = ps.TableStyleElements
    e(0).Interior.Color = C["card"]; e(0).Font.Color = C["ivory"]
    e(1).Interior.Color = C["head"]; e(1).Font.Color = C["gold"]; e(1).Font.Bold = True
    e(2).Interior.Color = C["alt"]; e(2).Font.Color = C["lgold"]; e(2).Font.Bold = True
    e(2).Borders(8).Color = C["gold"]  # top edge of the grand total row
    e(4).Font.Bold = True; e(4).Font.Color = C["lgold"]
    e(3).Font.Color = C["ivory"]
    e(26).Font.Color = C["muted"]; e(27).Font.Color = C["gold"]; e(27).Font.Bold = True

    ss = wb.TableStyles.Add("Gold Slicer")
    ss.ShowAsAvailableTableStyle = False; ss.ShowAsAvailablePivotTableStyle = False; ss.ShowAsAvailableSlicerStyle = True
    e = ss.TableStyleElements
    e(0).Interior.Color = C["card"]; e(0).Font.Color = C["ivory"]
    for i in (7, 8, 9, 10):
        e(0).Borders(i).Color = C["line"]
    e(1).Font.Color = C["gold"]; e(1).Font.Bold = True
    e(28).Interior.Color = C["alt"]; e(28).Font.Color = C["ivory"]       # unselected, with data
    e(29).Interior.Color = C["card"]; e(29).Font.Color = C["grey"]       # unselected, no data
    e(30).Interior.Color = C["gold"]; e(30).Font.Color = C["head"]; e(30).Font.Bold = True  # selected
    e(31).Interior.Color = hexrgb("6B5A38"); e(31).Font.Color = C["muted"]  # selected, no data
    e(32).Interior.Color = C["line"]; e(32).Font.Color = C["lgold"]      # hover
    e(33).Interior.Color = C["lgold"]; e(33).Font.Color = C["head"]
    for i in (28, 29, 30, 31, 32, 33):
        for b in (7, 8, 9, 10):
            e(i).Borders(b).Color = C["line"]


def page(ws, last_col=120, last_row=200):
    ws.Cells.Interior.Color = C["page"]
    ws.Cells.Font.Color = C["ivory"]
    ws.Cells.Font.Name = SANS
    ws.Cells.Font.Size = 10


def header(wb, ws, title, sub, nav_right, col="B"):
    """Charcoal band, spaced gold brand line, serif title, thin gold rule, navigation buttons (current one gold)."""
    ws.Rows("1:3").Interior.Color = C["head"]
    ws.Rows(1).RowHeight = 20; ws.Rows(2).RowHeight = 30; ws.Rows(3).RowHeight = 20
    ws.Rows(4).RowHeight = 2; ws.Rows(4).Interior.Color = C["gold"]
    ws.Rows(5).RowHeight = 10
    b = ws.Range(f"{col}1"); b.Value = BRAND; b.Font.Size = 8; b.Font.Bold = True; b.Font.Color = C["gold"]
    b.VerticalAlignment = -4107
    t = ws.Range(f"{col}2"); t.Value = title; t.Font.Name = SERIF; t.Font.Size = 20; t.Font.Color = C["ivory"]
    s = ws.Range(f"{col}3")
    if sub.startswith("="):
        s.Formula = sub
    else:
        s.Value = sub
    s.Font.Size = 9; s.Font.Color = C["muted"]; s.VerticalAlignment = -4108
    bw, gap = 84, 5
    left = nav_right - len(NAV) * (bw + gap) + gap
    for i, (label, target) in enumerate(NAV):
        shp = ws.Shapes.AddShape(5, left + i * (bw + gap), 12, bw, 24)
        shp.Adjustments.SetItem(1, 0.12)
        shp.Name = f"nav_{label}"
        cur = target == ws.Name
        shp.Fill.ForeColor.RGB = hexrgb("3A3324") if cur else C["head"]
        shp.Line.ForeColor.RGB = C["gold"] if cur else C["line"]
        shp.Line.Weight = 1
        tr = shp.TextFrame2.TextRange
        tr.Text = label
        tr.Font.Name = SANS; tr.Font.Size = 9; tr.Font.Bold = cur
        tr.Font.Fill.ForeColor.RGB = C["lgold"] if cur else C["ivory"]
        tr.ParagraphFormat.Alignment = 2
        shp.TextFrame2.VerticalAnchor = 3
        for m in ("MarginLeft", "MarginRight", "MarginTop", "MarginBottom"):
            setattr(shp.TextFrame2, m, 1)
        ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress=f"'{target}'!A1", ScreenTip=f"Go to {target}")
        shp.Placement = 3  # don't move or size with cells


def textbox(ws, left, top, w, h, text=None, link=None, size=10, color="ivory", bold=False, font=SANS, align=1, anchor=3):
    shp = ws.Shapes.AddTextbox(1, left, top, w, h)
    shp.Fill.Visible = 0; shp.Line.Visible = 0
    tf = shp.TextFrame2
    for m in ("MarginLeft", "MarginRight", "MarginTop", "MarginBottom"):
        setattr(tf, m, 0)
    tf.WordWrap = -1; tf.AutoSize = 0; tf.VerticalAnchor = anchor
    if link:
        shp.DrawingObject.Formula = link
    else:
        tf.TextRange.Text = text
    tr = tf.TextRange
    tr.Font.Name = font; tr.Font.Size = size; tr.Font.Bold = -1 if bold else 0
    tr.Font.Fill.ForeColor.RGB = C[color]
    tr.ParagraphFormat.Alignment = align
    shp.Placement = 3
    return shp


def style_chart(ch, title=None, title_link=None, legend=None):
    ch.ChartArea.Format.Fill.ForeColor.RGB = C["card"]
    ch.ChartArea.Format.Line.ForeColor.RGB = C["line"]
    ch.ChartArea.RoundedCorners = True
    ch.PlotArea.Format.Fill.Visible = 0
    f = ch.ChartArea.Format.TextFrame2.TextRange.Font
    f.Name = SANS; f.Size = 9; f.Fill.ForeColor.RGB = C["ivory"]
    if title or title_link:
        ch.HasTitle = True
        if title_link:
            try:
                ch.ChartTitle.Formula = title_link
            except Exception:
                ch.ChartTitle.Caption = title_link
        else:
            ch.ChartTitle.Text = title
        tf = ch.ChartTitle.Format.TextFrame2.TextRange.Font
        tf.Name = SERIF; tf.Size = 13; tf.Bold = 0; tf.Fill.ForeColor.RGB = C["ivory"]
        ch.ChartTitle.Left = 8; ch.ChartTitle.Top = 4
    else:
        ch.HasTitle = False
    if legend is None:
        ch.HasLegend = False
    else:
        ch.HasLegend = True
        ch.Legend.Position = legend
        ch.Legend.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = C["ivory"]


def axis_style(ax, visible_labels=True, grid=False):
    ax.Format.Line.ForeColor.RGB = C["line"]
    ax.HasMajorGridlines = grid
    if grid:
        ax.MajorGridlines.Format.Line.ForeColor.RGB = C["line"]
    ax.TickLabels.Font.Color = C["muted"]
    if not visible_labels:
        ax.TickLabelPosition = -4142  # none


# ---------------------------------------------------------------- loading
def load_query(ws, name, table_name, at="A7"):
    lo = ws.ListObjects.Add(0, conn_string(name), None, 1, ws.Range(at))  # xlSrcExternal, xlYes
    qt = lo.QueryTable
    qt.CommandType = 2  # xlCmdSql
    qt.CommandText = f"SELECT * FROM [{name}]"
    qt.RefreshStyle = 1  # xlInsertDeleteCells
    qt.AdjustColumnWidth = False
    qt.BackgroundQuery = False  # Refresh All waits for it, so the formulas see the new data
    qt.Refresh(False)
    lo.Name = table_name
    lo.TableStyle = "Gold Table"
    # the page fill/font set on the whole sheet would hide the table style: clear them inside the table
    lo.Range.Interior.ColorIndex = -4142
    lo.HeaderRowRange.Font.Color = C["gold"]; lo.HeaderRowRange.Font.Bold = True
    return lo


def format_tables(wb):
    widths = {"Status": 21, "Discrepancy type": 16, "Emp ID": 9, "Employee": 18, "HR line": 6, "HR fund center": 9,
              "HR functional area": 16, "HR FTE": 7, "Budget line": 7, "Position": 8, "Budget fund center": 9,
              "Budget functional area": 16, "Budget FTE": 8, "FTE difference": 9, "What to do": 70}
    tone = {"Matched": ("goodbg", "good"), "FTE differs": ("badbg", "bad"), "Functional area differs": ("badbg", "bad"),
            "Fund center differs": ("badbg", "bad"), "Only in HR roster": ("amberbg", "amber"),
            "Only in budget": ("skybg", "sky")}
    for title, sub in (("Comparison", "Every HR line and every budget line, side by side, with a status and what to do"),
                       ("Discrepancies", "Every line where HR and the budget disagree: both sides and the next step")):
        ws = wb.Worksheets(title)
        lo = ws.ListObjects(1)
        for c in ("HR FTE", "Budget FTE"):
            lo.ListColumns(c).DataBodyRange.NumberFormat = "0.00"
        diff = lo.ListColumns("FTE difference").DataBodyRange
        diff.NumberFormat = '+0.00;-0.00;"–"'
        fc = diff.FormatConditions.Add(1, 6, "=0"); fc.Font.Color = C["bad"]   # less than 0
        fc = diff.FormatConditions.Add(1, 5, "=0"); fc.Font.Color = C["amber"]  # greater than 0
        for i in range(1, lo.ListColumns.Count + 1):
            ws.Columns(i).ColumnWidth = widths.get(lo.ListColumns(i).Name, 12)
        lo.HeaderRowRange.WrapText = True
        lo.HeaderRowRange.VerticalAlignment = -4108
        ws.Rows(lo.HeaderRowRange.Row).RowHeight = 44
        rng = lo.ListColumns("Status").DataBodyRange
        first = rng.Row
        for text, (bg, fg) in tone.items():
            fc = rng.FormatConditions.Add(2, None, f'=$A{first}="{text}"')
            fc.Interior.Color = C[bg]; fc.Font.Color = C[fg]
        header(wb, ws, title, '="' + sub + '   ·   "&Period', ws.Range("P1").Left, col="A")
        ws.Rows(6).RowHeight = 6
        ws.Activate()
        ws.Range(f"A{first}").Select()
        wb.Application.ActiveWindow.FreezePanes = True
    ws = wb.Worksheets("Cleaning log")
    for c, w in zip("ABCDE", (15, 56, 7, 62, 30)):
        ws.Columns(c).ColumnWidth = w
    header(wb, ws, "Cleaning log", '="Every cleaning step, counted (current month: "&Period&")"', ws.Range("F1").Left, col="A")
    ws.Rows(6).RowHeight = 6


# ---------------------------------------------------------------- pivots and slicers
def pivot(pc, ws, at, name, rows=(), cols=(), pages=(), data=(), layout_tabular=True):
    pt = pc.CreatePivotTable(ws.Range(at), name)
    if layout_tabular:
        pt.RowAxisLayout(1)
    for f in rows:
        pt.PivotFields(f).Orientation = 1
    for f in cols:
        pt.PivotFields(f).Orientation = 2
    for f, item in pages:
        pf = pt.PivotFields(f)
        pf.Orientation = 3
        if item:
            pf.CurrentPage = item
    for field, caption, fmt in data:
        df = pt.AddDataField(pt.PivotFields(field), caption, -4157)  # xlSum
        df.NumberFormat = fmt
    for f in list(rows) + list(cols):
        try:
            pt.PivotFields(f).AutoSort(1, f)  # ascending
        except Exception:
            pass
    pt.ShowDrillIndicators = False
    pt.TableStyle2 = "Gold Pivot"
    pt.HasAutoFormat = False
    pt.NullString = ""
    pt.DisplayNullString = True
    return pt


def ref(pt):
    return f"'{pt.Parent.Name}'!{pt.TableRange1.Cells(1, 1).Address}"


def build_pivots(wb):
    ws = wb.Worksheets("Pivot comparison")
    conn = wb.Connections.Add2("Query - FTE_Long", "FTE from both sources, one row per side of every compared line, every month",
                               conn_string("FTE_Long"), "SELECT * FROM [FTE_Long]", 2, False, False)
    conn.OLEDBConnection.BackgroundQuery = False
    pc = wb.PivotCaches().Create(2, conn, 6)  # xlExternal, version 2016+
    P = {}
    ws.Columns("A").ColumnWidth = 2
    blocks = [("B", "HRPivot", "HR roster", "FTE", "HR roster: active FTE"),
              ("J", "BudgetPivot", "Finance budget", "FTE", "Finance budget: budgeted FTE"),
              ("R", "DiffPivot", None, "Budget minus HR", "Budget minus HR (+ = budget higher)")]
    for col, name, source, field, label in blocks:
        h = ws.Range(f"{col}13"); h.Value = label
        h.Font.Name = SERIF; h.Font.Size = 13; h.Font.Color = C["ivory"]
        pt = pivot(pc, ws, f"{col}17", name, rows=["Fund center"], cols=["Functional area"],
                   pages=[("Source", source)],
                   data=[(field, "Total FTE" if source else "FTE difference", "0.00" if source else '+0.00;-0.00;"–"')])
        pt.CompactLayoutRowHeader = "Fund center"
        P[name] = pt
    for c in "BJR":
        ws.Columns(c).ColumnWidth = 13
    for c in ["C", "D", "E", "F", "G", "K", "L", "M", "N", "O", "S", "T", "U", "V", "W"]:
        ws.Columns(c).ColumnWidth = 15.5
    for c in "HIPQ":
        ws.Columns(c).ColumnWidth = 2
    diff = P["DiffPivot"]
    fc = diff.DataBodyRange.FormatConditions.Add(1, 6, "=0")  # < 0: more HR than budget
    fc.Interior.Color = C["badbg"]; fc.Font.Color = C["bad"]; fc.ScopeType = 2
    fc = diff.DataBodyRange.FormatConditions.Add(1, 5, "=0")  # > 0: budget not filled
    fc.Interior.Color = C["amberbg"]; fc.Font.Color = C["amber"]; fc.ScopeType = 2
    n = ws.Range("B28")
    n.Value = ("How to read it: each cell is the FTE in that fund center and functional area for the month picked in the "
               "slicer. The third table is the second minus the first: red = more people than budget, amber = budget "
               "not filled. Every coloured cell is explained line by line on the Discrepancies tab.")
    n.Font.Italic = True; n.Font.Color = C["muted"]; n.Font.Size = 9

    # helper pivots for the dashboard, all on the same cache
    dd = wb.Worksheets("Dash data")
    P["PK"] = pivot(pc, dd, "A6", "PK", rows=["Discrepancy type"], cols=["Source"],
                    data=[("FTE", "Total FTE", "0.00"), ("Line", "Lines", "0")])
    P["PT"] = pivot(pc, dd, "A28", "PT", rows=["Month"], cols=["Source"],
                    data=[("FTE", "Total FTE", "0.00"), ("Budget minus HR", "FTE difference", "+0.00;-0.00;0.00")])
    P["PA"] = pivot(pc, dd, "A48", "PA", rows=["Functional area"], cols=["Source"],
                    data=[("FTE", "Total FTE", "0.00"), ("Line", "Lines", "0")])
    P["PM"] = pivot(pc, dd, "A62", "PM", rows=["Month"], data=[("Line", "Lines", "0")])
    P["PC"] = pivot(pc, dd, "J8", "PC", cols=["Source"], pages=[("Snapshot", "Current")],
                    data=[("FTE", "Total FTE", "0.00"), ("Budget minus HR", "FTE difference", "+0.00;-0.00;0.00")])
    P["PMo"] = pivot(pc, dd, "A80", "PMo", rows=["Month"], data=[("Line", "Lines", "0")])
    P["PFC"] = pivot(pc, dd, "D80", "PFC", rows=["Fund center"], data=[("Line", "Lines", "0")])
    P["PFA"] = pivot(pc, dd, "G80", "PFA", rows=["Functional area"], data=[("Line", "Lines", "0")])
    for nm, txt in (("A4", "PK: lines and FTE by discrepancy type (all slicers)"),
                    ("A26", "PT: by month (Fund center, Functional area, Discrepancy type slicers; not Month)"),
                    ("A46", "PA: by functional area (all slicers)"),
                    ("A60", "PM: month(s) picked in the Month slicer"),
                    ("J4", "PC: control pivot, current month, no slicers (for the data checks)"),
                    ("A78", "Item lists (no slicers): months, fund centers, functional areas")):
        dd.Range(nm).Value = txt; dd.Range(nm).Font.Bold = True; dd.Range(nm).Font.Color = C["gold"]
    return pc, P


def build_slicers(wb, P):
    dash, piv = wb.Worksheets("Dashboard"), wb.Worksheets("Pivot comparison")
    main = [P[n] for n in ("HRPivot", "BudgetPivot", "DiffPivot", "PK", "PA")]
    spec = [("Month", "Month", main + [P["PM"]], 3), ("Fund center", "Fund center", main + [P["PT"]], 6),
            ("Functional area", "Functional area", main + [P["PT"]], 2),
            ("Discrepancy type", "Discrepancy type", main + [P["PT"]], 4)]
    # dashboard positions (columns) and pivot-sheet positions (points)
    dpos = {"Month": (2, 13), "Fund center": (14, 33), "Functional area": (34, 49), "Discrepancy type": (50, 72)}
    caches = {}
    dash.Activate()
    for field, caption, pts, ncol in spec:
        sc = wb.SlicerCaches.Add2(pts[0], field, "Slicer_" + field.replace(" ", "_"))
        for pt in pts[1:]:
            sc.PivotTables.AddPivotTable(pt)
        c1, c2 = dpos[field]
        l = dash.Cells(6, c1).Left; r = dash.Cells(6, c2 + 1).Left - 6
        t = dash.Range("A6").Top; b = dash.Range("A11").Top - 2
        print("slicer", field, flush=True)
        s = sc.Slicers.Add(SlicerDestination=dash, Name="Dash_" + field.replace(" ", "_"), Caption=caption,
                           Top=t, Left=l, Width=r - l, Height=b - t)
        print(" added", flush=True)
        s.NumberOfColumns = ncol; print(" cols", flush=True)
        s.RowHeight = 20; print(" rowh", flush=True)
        s.Style = "Gold Slicer"; print(" style", flush=True)
        caches[field] = sc
    # the same slicers again on the pivot sheet (same caches, so they stay in sync)
    piv.Activate()
    x = piv.Range("B6").Left
    t = piv.Range("A6").Top; h = piv.Range("A11").Top - 2 - t
    total = piv.Range("X1").Left - x
    shares = {"Month": 0.17, "Fund center": 0.27, "Functional area": 0.22, "Discrepancy type": 0.34}
    for field, caption, pts, ncol in spec:
        w = total * shares[field] - 6
        s = caches[field].Slicers.Add(SlicerDestination=piv, Name="Piv_" + field.replace(" ", "_"), Caption=caption,
                                      Top=t, Left=x, Width=w, Height=h)
        s.NumberOfColumns = ncol; s.RowHeight = 20; s.Style = "Gold Slicer"
        x += w + 6
    # start on the latest month
    sc = caches["Month"]
    items = [it.Name for it in sc.SlicerItems]
    latest = max(items)
    for it in sc.SlicerItems:
        if it.Name != latest:
            it.Selected = False
    return caches


# ---------------------------------------------------------------- dashboard numbers (Dash data sheet, formulas only)
WF_TYPES = TYPES[1:]


def build_dash_data(wb, P):
    dd = wb.Worksheets("Dash data")
    dd.Cells.Font.Name = SANS; dd.Cells.Font.Size = 9
    dd.Range("A1").Value = "Dash data: the numbers behind the Dashboard (PivotTables on the left, formulas on the right)"
    dd.Range("A1").Font.Size = 13; dd.Range("A1").Font.Bold = True
    dd.Range("A2").Value = ("Everything here recalculates from the slicers. Nothing is typed by hand; "
                            "test/verify.py recomputes every figure from the CSV files.")
    PK, PT, PA, PM, PC = (ref(P[n]) for n in ("PK", "PT", "PA", "PM", "PC"))
    DIFF = ref(P["DiffPivot"])
    names = {}

    def name(nm, cell):
        add_name(wb, nm, f"='Dash data'!{dd.Range(cell).Address}")
        names[nm] = cell

    def put(cell, formula, nm=None, fmt=None):
        c = dd.Range(cell)
        c.Formula = formula
        if fmt:
            c.NumberFormat = fmt
        if nm:
            name(nm, cell)

    # item lists from the unfiltered list pivots (generic: new months / fund centers / areas show up by themselves)
    def item_list(col, pt, slots, nm):
        dd.Range(f"{col}4").Value = nm
        dd.Range(f"{col}4").Font.Bold = True
        top = P[pt].TableRange1.Row + 1
        src_col = P[pt].TableRange1.Column
        for k in range(slots):
            src = dd.Cells(top + k, src_col).Address
            dd.Range(f"{col}{5 + k}").Formula = f'=IF(OR({src}="",{src}="Grand Total"),"",{src})'
        add_name(wb, nm, f"='Dash data'!${col}$5:${col}${4 + slots}")
    item_list("R", "PMo", 12, "MonthList")
    item_list("S", "PFC", 10, "FcList")
    item_list("T", "PFA", 6, "AreaList")
    pm_top = P["PM"].TableRange1.Row + 1
    pm_col = P["PM"].TableRange1.Column
    sel_cells = ",".join(dd.Cells(pm_top + k, pm_col).Address for k in range(12))

    # key figures (column W labels, X values)
    dd.Range("W4").Value = "Key figures"; dd.Range("W4").Font.Bold = True
    rows = [
        ("SelCount", f'=SUMPRODUCT(--(COUNTIF(MonthList,{dd.Cells(pm_top, pm_col).Address}:{dd.Cells(pm_top + 11, pm_col).Address})>0))', "0", "Months picked"),
        ("SelMonth", f'=IF(SelCount=1,{dd.Cells(pm_top, pm_col).Address},"")', None, "Month shown"),
        ("MonthCount", '=COUNTIF(MonthList,"?*")', "0", "Months in the data"),
        ("PriorMonth", '=IF(SelMonth="","",IF(MATCH(SelMonth,MonthList,0)>1,INDEX(MonthList,MATCH(SelMonth,MonthList,0)-1),""))', None, "Prior month"),
        ("FirstMonth", '=INDEX(MonthList,1)', None, "First month"),
        ("K_HR", f'=IFERROR(ROUND(GETPIVOTDATA("Total FTE",{PK},"Source","HR roster"),2),0)', "0.00", "HR FTE"),
        ("K_BU", f'=IFERROR(ROUND(GETPIVOTDATA("Total FTE",{PK},"Source","Finance budget"),2),0)', "0.00", "Budget FTE"),
        ("K_GAP", '=ROUND(K_BU-K_HR,2)', "+0.00;-0.00;0.00", "Gap (budget minus HR)"),
        ("K_BLINES", f'=IFERROR(GETPIVOTDATA("Lines",{PK},"Source","Finance budget"),0)', "0", "Budget lines"),
        ("K_BMATCH", f'=IFERROR(GETPIVOTDATA("Lines",{PK},"Source","Finance budget","Discrepancy type","Matched"),0)', "0", "Budget lines matched"),
        ("K_PCT", '=IF(K_BLINES=0,0,K_BMATCH/K_BLINES)', "0.0%", "% budget lines matched"),
        ("K_VAC", f'=IFERROR(GETPIVOTDATA("Total FTE",{PK},"Source","Finance budget","Discrepancy type","Vacant post"),0)', "0.00", "Vacant FTE"),
        ("K_VACN", f'=IFERROR(GETPIVOTDATA("Lines",{PK},"Source","Finance budget","Discrepancy type","Vacant post"),0)', "0", "Vacant posts"),
        ("K_TERMN", f'=IFERROR(GETPIVOTDATA("Lines",{PK},"Source","Finance budget","Discrepancy type","Terminated"),0)', "0", "Terminated, still funded"),
        ("K_TERMF", f'=IFERROR(GETPIVOTDATA("Total FTE",{PK},"Source","Finance budget","Discrepancy type","Terminated"),0)', "0.00", "Terminated FTE"),
        ("K_HRONLY", f'=IFERROR(GETPIVOTDATA("Lines",{PK},"Source","HR roster","Discrepancy type","Only in HR"),0)', "0", "HR-only lines"),
        ("P_HR", f'=IF(PriorMonth="",0,IFERROR(ROUND(GETPIVOTDATA("Total FTE",{PT},"Month",PriorMonth,"Source","HR roster"),2),0))', "0.00", "HR FTE, prior month"),
        ("P_BU", f'=IF(PriorMonth="",0,IFERROR(ROUND(GETPIVOTDATA("Total FTE",{PT},"Month",PriorMonth,"Source","Finance budget"),2),0))', "0.00", "Budget FTE, prior month"),
        ("P_GAP", '=ROUND(P_BU-P_HR,2)', "+0.00;-0.00;0.00", "Gap, prior month"),
        ("G_FIRST", f'=IFERROR(GETPIVOTDATA("FTE difference",{PT},"Month",FirstMonth),0)', "+0.00;-0.00;0.00", "Gap, first month"),
    ]
    for i, (nm, f, fmt, label) in enumerate(rows, 5):
        dd.Range(f"W{i}").Value = label
        put(f"X{i}", f, nm, fmt)

    # waterfall: Budget -> 7 discrepancy types -> HR (rows 30-38), columns W..AF
    r0 = 30
    for j, h in enumerate(["Bar", "HR FTE", "Budget FTE", "Step (HR - budget)", "Running", "Base", "Total", "Up", "Down",
                           "HR lines", "Budget lines", "Lines", "Abs step"]):
        dd.Cells(r0 - 1, 23 + j).Value = h
        dd.Cells(r0 - 1, 23 + j).Font.Bold = True
    dd.Range(f"W{r0 - 2}").Value = "Bridge (waterfall) and donut"; dd.Range(f"W{r0 - 2}").Font.Bold = True
    first, last = r0, r0 + len(WF_TYPES) + 1
    for k in range(len(WF_TYPES) + 2):
        r = r0 + k
        if k == 0:
            dd.Range(f"W{r}").Formula = '="Budget "&TEXT(K_BU,"0.00")'
            dd.Range(f"AA{r}").Formula = "=K_BU"
        elif k == len(WF_TYPES) + 1:
            dd.Range(f"W{r}").Formula = '="HR "&TEXT(K_HR,"0.00")'
            dd.Range(f"AA{r}").Formula = f"=AA{r - 1}"
        else:
            t = WF_TYPES[k - 1]
            dd.Range(f"W{r}").Value = t
            dd.Range(f"X{r}").Formula = f'=IFERROR(GETPIVOTDATA("Total FTE",{PK},"Source","HR roster","Discrepancy type","{t}"),0)'
            dd.Range(f"Y{r}").Formula = f'=IFERROR(GETPIVOTDATA("Total FTE",{PK},"Source","Finance budget","Discrepancy type","{t}"),0)'
            dd.Range(f"Z{r}").Formula = f"=ROUND(X{r}-Y{r},2)"
            dd.Range(f"AA{r}").Formula = f"=ROUND(AA{r - 1}+Z{r},2)"
            dd.Range(f"AB{r}").Formula = f"=MIN(AA{r - 1},AA{r})-WF_FLOOR"
            dd.Range(f"AD{r}").Formula = f"=MAX(Z{r},0)"
            dd.Range(f"AE{r}").Formula = f"=MAX(-Z{r},0)"
            dd.Range(f"AF{r}").Formula = f'=IFERROR(GETPIVOTDATA("Lines",{PK},"Source","HR roster","Discrepancy type","{t}"),0)'
            dd.Range(f"AG{r}").Formula = f'=IFERROR(GETPIVOTDATA("Lines",{PK},"Source","Finance budget","Discrepancy type","{t}"),0)'
            dd.Range(f"AH{r}").Formula = f"=MAX(AF{r},AG{r})"
            dd.Range(f"AI{r}").Formula = f"=ABS(Z{r})"
        if k in (0, len(WF_TYPES) + 1):
            dd.Range(f"AB{r}").Value = 0
            dd.Range(f"AC{r}").Formula = f"=AA{r}-WF_FLOOR"
            dd.Range(f"AD{r}").Value = 0; dd.Range(f"AE{r}").Value = 0
        else:
            dd.Range(f"AC{r}").Value = 0
    dd.Range(f"X{r0}:AE{last}").NumberFormat = "0.00"
    put(f"X{last + 2}", f"=MAX(0,ROUNDDOWN(MIN(AA{first}:AA{last})-MAX(1,0.6*(MAX(AA{first}:AA{last})-MIN(AA{first}:AA{last}))),0))", "WF_FLOOR", "0")
    dd.Range(f"W{last + 2}").Value = "Bars start at (FTE)"
    put(f"X{last + 3}", f"=AA{last}", "WF_END", "0.00")
    dd.Range(f"W{last + 3}").Value = "Bridge ends at"
    put(f"X{last + 4}", f'="Bars start at "&WF_FLOOR&" FTE; blue = HR higher, copper = budget higher for that type"', "WF_TITLE")
    dd.Range(f"W{last + 4}").Value = "Bridge title"
    put(f"X{last + 5}", f"=SUM(AH{r0 + 1}:AH{last - 1})", "K_REVIEW", "0")
    dd.Range(f"W{last + 5}").Value = "Lines to review"
    add_name(wb, "WF_Labels", f"='Dash data'!$W${first}:$W${last}")
    for nm, col in (("WF_Base", "AB"), ("WF_Total", "AC"), ("WF_Up", "AD"), ("WF_Down", "AE")):
        add_name(wb, nm, f"='Dash data'!${col}${first}:${col}${last}")

    # donut: lines by discrepancy type (Matched + the 7 types), rows 46-53
    d0 = 46
    dd.Range(f"W{d0 - 1}").Value = "Donut: lines by type"; dd.Range(f"W{d0 - 1}").Font.Bold = True
    for k, t in enumerate(TYPES):
        r = d0 + k
        if t == "Matched":
            dd.Range(f"Y{r}").Formula = f'=MAX(IFERROR(GETPIVOTDATA("Lines",{PK},"Source","HR roster","Discrepancy type","Matched"),0),K_BMATCH)'
        else:
            dd.Range(f"Y{r}").Formula = f"=AH{r0 + k}"
        dd.Range(f"X{r}").Value = t
        dd.Range(f"W{r}").Formula = f'=X{r}&"  "&Y{r}'
    add_name(wb, "DonutLabels", f"='Dash data'!$W${d0}:$W${d0 + 7}")
    add_name(wb, "DonutValues", f"='Dash data'!$Y${d0}:$Y${d0 + 7}")
    put(f"X{d0 + 9}", '=TEXT(K_PCT,"0%")', "DonutCenter")
    dd.Range(f"W{d0 + 9}").Value = "Donut centre"

    # trend: one row per month in the data (rows 60-71)
    t0 = 60
    dd.Range(f"W{t0 - 2}").Value = "Trend by month"; dd.Range(f"W{t0 - 2}").Font.Bold = True
    for j, h in enumerate(["Month", "HR FTE", "Budget FTE", "Gap"]):
        dd.Cells(t0 - 1, 23 + j).Value = h; dd.Cells(t0 - 1, 23 + j).Font.Bold = True
    for k in range(12):
        r = t0 + k
        dd.Range(f"W{r}").Formula = f"=INDEX(MonthList,{k + 1})"
        dd.Range(f"X{r}").Formula = f'=IF(W{r}="","",IFERROR(GETPIVOTDATA("Total FTE",{PT},"Month",W{r},"Source","HR roster"),0))'
        dd.Range(f"Y{r}").Formula = f'=IF(W{r}="","",IFERROR(GETPIVOTDATA("Total FTE",{PT},"Month",W{r},"Source","Finance budget"),0))'
        dd.Range(f"Z{r}").Formula = f'=IF(W{r}="","",ROUND(Y{r}-X{r},2))'
    dd.Range(f"X{t0}:Z{t0 + 11}").NumberFormat = "0.00"
    for nm, col in (("TrendMonths", "W"), ("TrendHR", "X"), ("TrendBU", "Y"), ("TrendGap", "Z")):
        add_name(wb, nm, f"=OFFSET('Dash data'!${col}${t0},0,0,MAX(1,MonthCount),1)")

    # FTE by functional area (rows 76-81)
    a0 = 76
    dd.Range(f"W{a0 - 2}").Value = "By functional area"; dd.Range(f"W{a0 - 2}").Font.Bold = True
    for j, h in enumerate(["Label", "Area", "HR FTE", "Budget FTE", "HR lines"]):
        dd.Cells(a0 - 1, 23 + j).Value = h; dd.Cells(a0 - 1, 23 + j).Font.Bold = True
    for k in range(6):
        r = a0 + k
        dd.Range(f"X{r}").Formula = f"=INDEX(AreaList,{k + 1})"
        dd.Range(f"Y{r}").Formula = f'=IF(X{r}="","",IFERROR(GETPIVOTDATA("Total FTE",{PA},"Functional area",X{r},"Source","HR roster"),0))'
        dd.Range(f"Z{r}").Formula = f'=IF(X{r}="","",IFERROR(GETPIVOTDATA("Total FTE",{PA},"Functional area",X{r},"Source","Finance budget"),0))'
        dd.Range(f"AA{r}").Formula = f'=IF(X{r}="","",IFERROR(GETPIVOTDATA("Lines",{PA},"Functional area",X{r},"Source","HR roster"),0))'
        dd.Range(f"W{r}").Formula = f'=IF(X{r}="","",X{r}&" · "&AA{r}&" HR lines")'
    dd.Range(f"Y{a0}:Z{a0 + 5}").NumberFormat = "0.00"
    put(f"X{a0 + 7}", '=COUNTIF(AreaList,"?*")', "AreaCount", "0")
    dd.Range(f"W{a0 + 7}").Value = "Areas"
    for nm, col in (("AreaLabels", "W"), ("AreaHR", "Y"), ("AreaBU", "Z")):
        add_name(wb, nm, f"=OFFSET('Dash data'!${col}${a0},0,0,MAX(1,AreaCount),1)")

    # heatmap grid: budget minus HR by fund center x functional area (rows 88-97, columns X..AC) + row totals in AD
    h0 = 88
    dd.Range(f"W{h0 - 2}").Value = "Heatmap: budget minus HR"; dd.Range(f"W{h0 - 2}").Font.Bold = True
    for j in range(6):
        dd.Cells(h0 - 1, 24 + j).Formula = f"=INDEX(AreaList,{j + 1})"
    dd.Cells(h0 - 1, 30).Value = "Row total"
    for k in range(10):
        r = h0 + k
        dd.Range(f"W{r}").Formula = f"=INDEX(FcList,{k + 1})"
        for j in range(6):
            a = dd.Cells(h0 - 1, 24 + j).Address
            dd.Cells(r, 24 + j).Formula = (f'=IF(OR($W{r}="",{a}=""),"",IFERROR(ROUND(GETPIVOTDATA("FTE difference",{DIFF},'
                                            f'"Fund center",$W{r},"Functional area",{a}),2),""))')
        dd.Range(f"AD{r}").Formula = f'=IF($W{r}="","",IFERROR(ROUND(GETPIVOTDATA("FTE difference",{DIFF},"Fund center",$W{r}),2),""))'
    dd.Range(f"X{h0}:AD{h0 + 9}").NumberFormat = "+0.00;-0.00;0.00"
    add_name(wb, "HeatGrid", f"='Dash data'!$X${h0}:$AC${h0 + 9}")
    add_name(wb, "HeatRows", f"='Dash data'!$AD${h0}:$AD${h0 + 9}")
    add_name(wb, "HeatFc", f"='Dash data'!$W${h0}:$W${h0 + 9}")
    add_name(wb, "HeatAreas", f"='Dash data'!$X${h0 - 1}:$AC${h0 - 1}")

    # insights written by formulas (rows 102-...)
    i0 = 102
    dd.Range(f"W{i0 - 2}").Value = "Insights"; dd.Range(f"W{i0 - 2}").Font.Bold = True
    helpers = [
        ("I_MIN", '=IF(COUNT(HeatRows)=0,0,MIN(HeatRows))'),
        ("I_MINFC", '=IF(I_MIN<0,INDEX(HeatFc,MATCH(I_MIN,HeatRows,0)),"")'),
        ("I_MINAREA", '=IF(I_MINFC="","",INDEX(HeatAreas,MATCH(MIN(INDEX(HeatGrid,MATCH(I_MINFC,HeatFc,0),0)),INDEX(HeatGrid,MATCH(I_MINFC,HeatFc,0),0),0)))'),
        ("I_MAX", '=IF(COUNT(HeatRows)=0,0,MAX(HeatRows))'),
        ("I_MAXFC", '=IF(I_MAX>0,INDEX(HeatFc,MATCH(I_MAX,HeatRows,0)),"")'),
        ("I_MAXAREA", '=IF(I_MAXFC="","",INDEX(HeatAreas,MATCH(MAX(INDEX(HeatGrid,MATCH(I_MAXFC,HeatFc,0),0)),INDEX(HeatGrid,MATCH(I_MAXFC,HeatFc,0),0),0)))'),
        ("I_DRV", f"=MAX(AI{r0 + 1}:AI{last - 1})"),
        ("I_DRVK", f"=IF(I_DRV=0,0,MATCH(I_DRV,AI{r0 + 1}:AI{last - 1},0))"),
        ("I_TODO", '=IF(K_TERMN>0,"remove "&K_TERMN&" terminated "&IF(K_TERMN=1,"person","people")&" from the budget; ","")'
                   '&IF(K_VACN>0,"decide on "&K_VACN&" vacant "&IF(K_VACN=1,"post","posts")&"; ","")'
                   '&IF(K_HRONLY>0,"fund or check "&K_HRONLY&" HR-only "&IF(K_HRONLY=1,"hire","hires")&"; ","")'),
    ]
    for i, (nm, f) in enumerate(helpers, i0):
        dd.Range(f"W{i}").Value = nm
        put(f"X{i}", f, nm)
    s0 = i0 + len(helpers) + 1
    lines = 'K_REVIEW&IF(K_REVIEW=1," line"," lines")'
    sentences = [
        ('=IF(SelCount<>1,"Pick one month in the Month slicer (now "&SelCount&" months are added together)",'
         'IF(K_GAP>0,"Budget is "&TEXT(K_GAP,"0.00")&" FTE above the HR roster; "&' + lines + '&" explain it",'
         'IF(K_GAP<0,"HR roster is "&TEXT(-K_GAP,"0.00")&" FTE above budget; "&' + lines + '&" explain it",'
         '"HR roster and budget agree on FTE; "&' + lines + '&" to review")))',
         '=IF(SelCount<>1,"bad","neutral")'),
        ('=IF(I_MINFC="","No fund center has more people than budget",I_MINFC&" is "&TEXT(-I_MIN,"0.00")&" FTE over budget, mostly "&I_MINAREA)',
         '=IF(I_MINFC="","good","bad")'),
        ('=IF(I_MAXFC="","No fund center has budget left unfilled",I_MAXFC&" has "&TEXT(I_MAX,"0.00")&" budgeted FTE not filled, mostly "&I_MAXAREA)',
         '=IF(I_MAXFC="","good","warn")'),
        (f'=IF(I_DRVK=0,"The discrepancies net to zero FTE in this view","Biggest driver: "&INDEX(W{r0 + 1}:W{last - 1},I_DRVK)&", "'
         f'&INDEX(AH{r0 + 1}:AH{last - 1},I_DRVK)&IF(INDEX(AH{r0 + 1}:AH{last - 1},I_DRVK)=1," line"," lines")&", HR "'
         f'&TEXT(INDEX(Z{r0 + 1}:Z{last - 1},I_DRVK),"+0.00;-0.00")&" FTE vs budget")',
         '="neutral"'),
        ('=IF(SelCount<>1,"Trend: pick one month",IF(SelMonth=FirstMonth,SelMonth&" is the first month in the data",'
         '"Gap went from "&TEXT(G_FIRST,"+0.00;-0.00;0.00")&" in "&FirstMonth&" to "&TEXT(K_GAP,"+0.00;-0.00;0.00")&" in "&SelMonth'
         '&IF(ABS(K_GAP)<ABS(G_FIRST),": narrowing",IF(ABS(K_GAP)>ABS(G_FIRST),": widening",": flat"))))',
         '=IF(OR(SelCount<>1,SelMonth=FirstMonth),"neutral",IF(ABS(K_GAP)<=ABS(G_FIRST),"good","bad"))'),
        ('=IF(I_TODO="","Nothing to fix in this view","To do: "&LEFT(I_TODO,LEN(I_TODO)-2))',
         '=IF(I_TODO="","good","warn")'),
    ]
    for k, (f, tone) in enumerate(sentences):
        r = s0 + k
        dd.Range(f"W{r}").Value = f"Insight {k + 1}"
        put(f"X{r}", f, f"INSIGHT{k + 1}")
        put(f"Y{r}", tone, f"TONE{k + 1}")

    # KPI card texts: value, note, tone (good / bad / muted) -> three overlapping text boxes show the right colour
    c0 = s0 + len(sentences) + 2
    dd.Range(f"W{c0 - 1}").Value = "KPI cards"; dd.Range(f"W{c0 - 1}").Font.Bold = True
    arrow = lambda d: f'IF({d}>0,"▲ ",IF({d}<0,"▼ ",""))'
    cards = [
        ("=K_HR", "0.00", f'=IF(PriorMonth="","first month in the data",{arrow("ROUND(K_HR-P_HR,2)")}&TEXT(ROUND(K_HR-P_HR,2),"+0.00;-0.00;0.00")&" vs "&PriorMonth)',
         '=IF(PriorMonth="","muted",IF(ABS(K_GAP)<=ABS(P_GAP),"good","bad"))'),
        ("=K_BU", "0.00", f'=IF(PriorMonth="","first month in the data",{arrow("ROUND(K_BU-P_BU,2)")}&TEXT(ROUND(K_BU-P_BU,2),"+0.00;-0.00;0.00")&" vs "&PriorMonth)',
         '=IF(PriorMonth="","muted",IF(ABS(K_GAP)<=ABS(P_GAP),"good","bad"))'),
        ("=K_GAP", "+0.00;-0.00;0.00", '=IF(PriorMonth="","budget minus HR",IF(ABS(K_GAP)<ABS(P_GAP),"▼ narrowed from ",'
         'IF(ABS(K_GAP)>ABS(P_GAP),"▲ widened from ","same as "))&TEXT(P_GAP,"+0.00;-0.00;0.00")&" in "&PriorMonth)',
         '=IF(PriorMonth="","muted",IF(ABS(K_GAP)<=ABS(P_GAP),"good","bad"))'),
        ("=K_PCT", "0.0%", '=(K_BLINES-K_BMATCH)&" of "&K_BLINES&" budget lines to fix"', '=IF(K_PCT>=0.95,"good","bad")'),
        ("=K_VAC", "0.00", '=IF(K_VACN=0,"no vacant posts",K_VACN&IF(K_VACN=1," post"," posts")&": recruit or release")',
         '=IF(K_VACN=0,"good","bad")'),
        ("=K_TERMN", "0", '=IF(K_TERMN=0,"none: the budget is clean",TEXT(K_TERMF,"0.00")&" FTE still funded: remove")',
         '=IF(K_TERMN=0,"good","bad")'),
    ]
    for k, (val, fmt, note, tone) in enumerate(cards):
        r = c0 + k
        dd.Range(f"W{r}").Value = f"Card {k + 1}"
        put(f"X{r}", val, f"CARD{k + 1}", fmt)
        put(f"Y{r}", note, f"NOTE{k + 1}")
        put(f"Z{r}", tone, f"CTONE{k + 1}")
        for j, t in enumerate(("good", "bad", "muted")):
            dd.Cells(r, 27 + j).Formula = f'=IF(Z{r}="{t}",Y{r},"")'
    dd.Columns("W").ColumnWidth = 26
    dd.Columns("X").ColumnWidth = 14
    return c0


# ---------------------------------------------------------------- dashboard sheet
def dashboard_grid(wb):
    ws = wb.Worksheets("Dashboard")
    ws.Columns("A").ColumnWidth = 1.5
    ws.Range(ws.Columns(2), ws.Columns(80)).ColumnWidth = 2.71
    for r in range(6, 60):
        ws.Rows(r).RowHeight = 18
    heights = {11: 8, 17: 8, 18: 26, 19: 16, 20: 30, 27: 18, 28: 8, 29: 8, 30: 26, 31: 6, 38: 8, 39: 30, 40: 18, 41: 8}
    for r in range(21, 27):
        heights[r] = 30
    for r in range(32, 38):
        heights[r] = 22
    for r, h in heights.items():
        ws.Rows(r).RowHeight = h
    right = ws.Cells(1, 73).Left
    header(wb, ws, "FTE dashboard",
           '="HR roster vs Finance position budget   ·   "&IF(SelCount=1,"showing "&SelMonth,"pick one month")&"   ·   fictional data"',
           right)


def build_dashboard(wb, P, card_row):
    ws = wb.Worksheets("Dashboard")
    dd = wb.Worksheets("Dash data")

    def box(c1, r1, c2, r2):
        a, b = ws.Cells(r1, c1), ws.Cells(r2 + 1, c2 + 1)
        return a.Left, a.Top, b.Left - a.Left, b.Top - a.Top

    def dref(nm):
        return f"='Dash data'!{dd.Range(nm).Address}"

    # KPI cards (rows 12-16)
    cards = [("HR ROSTER FTE", "👥"), ("FINANCE BUDGET FTE", "▦"), ("GAP: BUDGET MINUS HR", "±"),
             ("BUDGET LINES MATCHED", "✓"), ("VACANT FTE", "○"), ("TERMINATED, STILL FUNDED", "✕")]
    for k, (label, icon) in enumerate(cards):
        l, t, w, h = box(2 + 12 * k, 12, 12 + 12 * k, 16)
        bg = ws.Shapes.AddShape(5, l, t, w, h)
        bg.Adjustments.SetItem(1, 0.08)
        bg.Fill.ForeColor.RGB = C["card"]; bg.Line.ForeColor.RGB = C["line"]; bg.Line.Weight = 0.75
        bg.Placement = 3
        bar = ws.Shapes.AddShape(1, l + 6, t + 12, 2.5, h - 24)
        bar.Fill.ForeColor.RGB = C["gold"]; bar.Line.Visible = 0; bar.Placement = 3
        textbox(ws, l + 16, t + 8, w - 60, 14, text=label, size=8, color="muted", bold=True)
        ic = ws.Shapes.AddShape(9, l + w - 36, t + 8, 26, 26)
        ic.Fill.ForeColor.RGB = hexrgb("3A3324"); ic.Line.ForeColor.RGB = C["gold"]; ic.Line.Weight = 1; ic.Placement = 3
        tr = ic.TextFrame2.TextRange; tr.Text = icon
        tr.Font.Name = "Segoe UI Symbol"; tr.Font.Size = 11; tr.Font.Fill.ForeColor.RGB = C["gold"]
        tr.ParagraphFormat.Alignment = 2; ic.TextFrame2.VerticalAnchor = 3
        for m in ("MarginLeft", "MarginRight", "MarginTop", "MarginBottom"):
            setattr(ic.TextFrame2, m, 0)
        textbox(ws, l + 12, t + 26, w - 24, 34, link=f"='Dash data'!$X${card_row + k}", size=24, font=SERIF, align=2)
        for j, tone in enumerate(("good", "bad", "muted")):
            col = dd.Cells(card_row + k, 27 + j).Address
            textbox(ws, l + 8, t + h - 24, w - 16, 14, link=f"='Dash data'!{col}", size=8.5, color=tone, align=2)

    # waterfall (stacked columns, invisible base), columns 2-28, rows 18-28
    l, t, w, h = box(2, 18, 28, 28)
    ch = ws.ChartObjects().Add(l, t, w, h).Chart
    ch.ChartType = 52  # stacked column
    for nm, color in (("Base", None), ("Total", None), ("HR higher", "sky"), ("Budget higher", "copper")):
        s = ch.SeriesCollection().NewSeries()
        s.Name = nm
        key = {"Base": "WF_Base", "Total": "WF_Total", "HR higher": "WF_Up", "Budget higher": "WF_Down"}[nm]
        s.Values = f"='{wb.Name}'!{key}"
        s.XValues = f"='{wb.Name}'!WF_Labels"
        if nm == "Base":
            s.Format.Fill.Visible = 0; s.Format.Line.Visible = 0
        elif nm == "Total":
            s.Format.Fill.ForeColor.RGB = C["grey"]
            s.Points(len(WF_TYPES) + 2).Format.Fill.ForeColor.RGB = C["jade"]
            s.Points(1).Format.Fill.ForeColor.RGB = C["grey"]
        else:
            s.Format.Fill.ForeColor.RGB = C[color]
            s.HasDataLabels = True
            s.DataLabels().NumberFormat = "+0.00;;" if nm == "HR higher" else '"-"0.00;;'
            s.DataLabels().Position = 3 if nm == "HR higher" else 3  # inside end
            s.DataLabels().Font.Color = C["ivory"]; s.DataLabels().Font.Size = 9; s.DataLabels().Font.Bold = True
    ch.ChartGroups(1).GapWidth = 45
    style_chart(ch, title="Bridge: budget FTE → HR FTE, by discrepancy type", legend=-4160)  # top
    ch.HasLegend = False  # the subtitle says blue = HR higher, copper = budget higher
    axis_style(ch.Axes(2), visible_labels=False)
    axis_style(ch.Axes(1))
    ch.Axes(1).TickLabels.Font.Color = C["ivory"]; ch.Axes(1).TickLabels.Font.Size = 8
    ch.Parent.Name = "chart_bridge"
    textbox(ws, l + 10, t + 28, w - 20, 12, link=f"='Dash data'!{dd.Range('WF_TITLE').Address}", size=8, color="muted")

    # heatmap panel in cells, columns 30-52, rows 18-28
    pnl = ws.Range(ws.Cells(18, 30), ws.Cells(28, 52))
    pnl.Interior.Color = C["card"]
    for e in (7, 8, 9, 10):
        pnl.Borders(e).Color = C["line"]
    tc = ws.Cells(18, 30); tc.Value = " Gap heatmap: fund center × functional area"
    tc.Font.Name = SERIF; tc.Font.Size = 13
    sc = ws.Cells(19, 30); sc.Value = "  budget minus HR FTE · red = more people than budget · blue = budget not filled"
    sc.Font.Size = 8; sc.Font.Color = C["muted"]
    for j in range(4):
        c1 = 34 + 4 * j
        hdr = ws.Range(ws.Cells(20, c1), ws.Cells(20, c1 + 3))
        hdr.Merge(); hdr.Formula = f"='Dash data'!{dd.Cells(87, 24 + j).Address}"
        hdr.HorizontalAlignment = -4108; hdr.VerticalAlignment = -4108; hdr.WrapText = True
        hdr.Font.Size = 8.5; hdr.Font.Color = C["lgold"]; hdr.Font.Bold = True
    tot = ws.Range(ws.Cells(20, 50), ws.Cells(20, 52)); tot.Merge(); tot.Value = "Total"
    tot.HorizontalAlignment = -4108; tot.VerticalAlignment = -4108; tot.Font.Size = 8.5; tot.Font.Bold = True
    tot.Font.Color = C["lgold"]
    grid_cells = []
    for k in range(6):
        r = 21 + k
        lab = ws.Range(ws.Cells(r, 30), ws.Cells(r, 33)); lab.Merge()
        lab.Formula = f"=\"  \"&'Dash data'!{dd.Cells(88 + k, 23).Address}"
        lab.Font.Bold = True; lab.VerticalAlignment = -4108; lab.Font.Color = C["ivory"]
        for j in range(4):
            c1 = 34 + 4 * j
            cell = ws.Range(ws.Cells(r, c1), ws.Cells(r, c1 + 3)); cell.Merge()
            cell.Formula = f"='Dash data'!{dd.Cells(88 + k, 24 + j).Address}"
            cell.HorizontalAlignment = -4108; cell.VerticalAlignment = -4108
            cell.NumberFormat = "+0.00;-0.00;0.00"; cell.Font.Size = 11; cell.Font.Bold = True
            for e in (7, 8, 9, 10):
                cell.Borders(e).Color = C["card"]; cell.Borders(e).Weight = -4138  # medium: gaps between tiles
            grid_cells.append(cell)
        rt = ws.Range(ws.Cells(r, 50), ws.Cells(r, 52)); rt.Merge()
        rt.Formula = f"='Dash data'!{dd.Cells(88 + k, 30).Address}"
        rt.HorizontalAlignment = -4108; rt.VerticalAlignment = -4108; rt.NumberFormat = "+0.00;-0.00;0.00"
        rt.Font.Bold = True; rt.Font.Color = C["lgold"]
    grid = ws.Range(ws.Cells(21, 34), ws.Cells(26, 49))
    cs = grid.FormatConditions.AddColorScale(3)
    cs.ColorScaleCriteria(1).Type = 1; cs.ColorScaleCriteria(1).FormatColor.Color = C["bad"]          # lowest value
    cs.ColorScaleCriteria(2).Type = 0; cs.ColorScaleCriteria(2).Value = 0                               # number 0
    cs.ColorScaleCriteria(2).FormatColor.Color = C["alt"]
    cs.ColorScaleCriteria(3).Type = 2; cs.ColorScaleCriteria(3).FormatColor.Color = hexrgb("4F7FAF")   # highest value
    lg = ws.Cells(27, 30); lg.Value = "  0.00 = HR and budget agree in that cell (or nobody is booked there)"
    lg.Font.Size = 8; lg.Font.Color = C["muted"]

    # donut, columns 54-72, rows 18-28
    l, t, w, h = box(54, 18, 72, 28)
    ch = ws.ChartObjects().Add(l, t, w, h).Chart
    ch.ChartType = -4120
    s = ch.SeriesCollection().NewSeries()
    s.Name = "Lines"
    s.Values = f"='{wb.Name}'!DonutValues"
    s.XValues = f"='{wb.Name}'!DonutLabels"
    style_chart(ch, title="Reconciliation status (lines)", legend=-4152)  # right
    ch.ChartGroups(1).DoughnutHoleSize = 62
    colors = ["jade", "copper", "amber", "sky", "bad", "lgold", "muted", "grey"]
    for i, c in enumerate(colors, 1):
        s.Points(i).Format.Fill.ForeColor.RGB = C[c]
        s.Points(i).Format.Line.ForeColor.RGB = C["card"]
    ch.Legend.Font.Size = 8.5
    ch.Parent.Name = "chart_donut"
    pa = ch.PlotArea
    cx = l + pa.InsideLeft + pa.InsideWidth / 2; cy = t + pa.InsideTop + pa.InsideHeight / 2
    textbox(ws, cx - 45, cy - 20, 90, 24, link=f"='Dash data'!{dd.Range('DonutCenter').Address}", size=20, font=SERIF,
            color="ivory", align=2)
    textbox(ws, cx - 45, cy + 4, 90, 14, text="budget lines matched", size=7.5, color="muted", align=2)

    # trend (columns + gap line on a second axis), columns 2-23, rows 30-41
    l, t, w, h = box(2, 30, 23, 41)
    ch = ws.ChartObjects().Add(l, t, w, h).Chart
    ch.ChartType = 51
    for nm, key, color in (("HR FTE", "TrendHR", "jade"), ("Budget FTE", "TrendBU", "grey"), ("Gap", "TrendGap", "gold")):
        s = ch.SeriesCollection().NewSeries()
        s.Name = nm
        s.Values = f"='{wb.Name}'!{key}"
        s.XValues = f"='{wb.Name}'!TrendMonths"
        if nm == "Gap":
            s.ChartType = 65; s.AxisGroup = 2
            s.Format.Line.ForeColor.RGB = C["gold"]; s.Format.Line.Weight = 2.5
            s.MarkerStyle = 8; s.MarkerSize = 8
            s.MarkerBackgroundColor = C["gold"]; s.MarkerForegroundColor = C["gold"]
            s.HasDataLabels = True
            s.DataLabels().NumberFormat = "+0.00;-0.00;0.00"; s.DataLabels().Position = 0  # above
            s.DataLabels().Font.Color = C["lgold"]; s.DataLabels().Font.Bold = True; s.DataLabels().Font.Size = 10
        else:
            s.Format.Fill.ForeColor.RGB = C[color]
            s.HasDataLabels = True
            s.DataLabels().NumberFormat = "0.0"; s.DataLabels().Position = 3
            s.DataLabels().Font.Color = C["ivory"]; s.DataLabels().Font.Size = 8
    ch.ChartGroups(1).GapWidth = 60
    style_chart(ch, title="3-month trend: FTE and gap", legend=-4160)
    axis_style(ch.Axes(1)); ch.Axes(1).TickLabels.Font.Color = C["ivory"]
    axis_style(ch.Axes(2), grid=True)
    ch.Axes(2, 2).TickLabelPosition = -4142
    ch.Axes(2, 2).Format.Line.Visible = 0
    ch.Axes(2, 2).MinimumScale = 0
    ch.Parent.Name = "chart_trend"

    # FTE by functional area (clustered bars), columns 25-41, rows 30-41
    l, t, w, h = box(25, 30, 41, 41)
    ch = ws.ChartObjects().Add(l, t, w, h).Chart
    ch.ChartType = 57
    for nm, key, color in (("HR FTE", "AreaHR", "jade"), ("Budget FTE", "AreaBU", "grey")):
        s = ch.SeriesCollection().NewSeries()
        s.Name = nm
        s.Values = f"='{wb.Name}'!{key}"
        s.XValues = f"='{wb.Name}'!AreaLabels"
        s.Format.Fill.ForeColor.RGB = C[color]
        s.HasDataLabels = True
        s.DataLabels().NumberFormat = "0.0"; s.DataLabels().Font.Color = C["ivory"]; s.DataLabels().Font.Size = 8
    ch.ChartGroups(1).GapWidth = 50; ch.ChartGroups(1).Overlap = -5
    style_chart(ch, title="FTE and headcount by functional area", legend=-4160)
    ch.Axes(1).ReversePlotOrder = True
    axis_style(ch.Axes(1)); ch.Axes(1).TickLabels.Font.Color = C["ivory"]; ch.Axes(1).TickLabels.Font.Size = 8
    axis_style(ch.Axes(2), visible_labels=False)
    ch.Parent.Name = "chart_areas"

    # insights + checks panel, columns 43-72, rows 30-41
    pnl = ws.Range(ws.Cells(30, 43), ws.Cells(41, 72))
    pnl.Interior.Color = C["card"]
    for e in (7, 8, 9, 10):
        pnl.Borders(e).Color = C["line"]
    tc = ws.Cells(30, 43); tc.Value = " Key insights, written by formulas"
    tc.Font.Name = SERIF; tc.Font.Size = 13
    for k in range(6):
        r = 32 + k
        b = ws.Cells(r, 44); b.Value = "●"; b.Font.Size = 9; b.HorizontalAlignment = -4108
        txt = ws.Range(ws.Cells(r, 45), ws.Cells(r, 72)); txt.Merge()
        txt.Formula = f"=INSIGHT{k + 1}"
        txt.Font.Size = 9.5; txt.VerticalAlignment = -4108
        both = ws.Range(ws.Cells(r, 44), ws.Cells(r, 72))
        for tone, color in (("bad", "bad"), ("good", "good"), ("warn", "amber")):
            fc = both.FormatConditions.Add(2, None, f'=TONE{k + 1}="{tone}"')
            fc.Font.Color = C[color]
    sep = ws.Range(ws.Cells(38, 43), ws.Cells(38, 72))
    sep.Borders(9).Color = C["line"]
    chk = ws.Range(ws.Cells(39, 44), ws.Cells(39, 56)); chk.Merge()
    chk.Formula = '=CHK_PASS&" / "&CHK_N&"  checks PASS"'
    chk.Font.Name = SERIF; chk.Font.Size = 16; chk.Font.Color = C["good"]; chk.VerticalAlignment = -4108
    fc = chk.FormatConditions.Add(2, None, "=CHK_PASS<CHK_N"); fc.Font.Color = C["bad"]
    ws.Hyperlinks.Add(Anchor=chk, Address="", SubAddress="'Checks'!A1", ScreenTip="See every data check")
    chk.Font.Underline = -4142
    sub = ws.Range(ws.Cells(39, 57), ws.Cells(40, 72)); sub.Merge(); sub.WrapText = True
    sub.Value = ("no row lost · every HR and budget line compared once · pivots = line totals · "
                 "bridge and heatmap = cards")
    sub.Font.Size = 8; sub.Font.Color = C["muted"]; sub.VerticalAlignment = -4108
    note = ws.Range(ws.Cells(40, 44), ws.Cells(40, 56)); note.Merge()
    note.Value = "click for the Checks tab"; note.Font.Size = 8; note.Font.Color = C["muted"]
    ws.Hyperlinks.Add(Anchor=note, Address="", SubAddress="'Checks'!A1")
    note.Font.Underline = -4142; note.Font.Color = C["muted"]; note.Font.Name = SANS; note.Font.Size = 8
    chk.Font.Name = SERIF; chk.Font.Size = 18; chk.Font.Color = C["good"]; chk.Font.Bold = False

    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = f"$A$1:{ws.Cells(42, 73).Address}"


# ---------------------------------------------------------------- checks sheet (result by status + data checks)
def build_checks(wb, P):
    ws = wb.Worksheets("Checks")
    widths = {"A": 2, "B": 36, "C": 10, "D": 11, "E": 11, "F": 13, "G": 3, "H": 12, "I": 4, "J": 14, "K": 14, "L": 14,
              "M": 14, "N": 14, "O": 14}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    header(wb, ws, "Data checks", '="Result by status and the data checks for the current month: "&Period', ws.Range("P1").Left)

    def head(cell, text):
        c = ws.Range(cell); c.Value = text; c.Font.Name = SERIF; c.Font.Size = 13
    head("B7", "Result by status (Comparison tab)")
    for c, t in {"B": "Status", "C": "Lines", "D": "HR FTE", "E": "Budget FTE", "F": "Budget - HR"}.items():
        h = ws.Range(f"{c}8")
        h.Value = t; h.Font.Bold = True; h.Interior.Color = C["head"]; h.Font.Color = C["gold"]
        if c != "B":
            h.HorizontalAlignment = -4152
    for i, st in enumerate(STATUSES, 9):
        ws.Range(f"B{i}").Value = st
        ws.Range(f"C{i}").Formula = f"=COUNTIFS(Compare[Status],B{i})"
        ws.Range(f"D{i}").Formula = f"=SUMIFS(Compare[HR FTE],Compare[Status],B{i})"
        ws.Range(f"E{i}").Formula = f"=SUMIFS(Compare[Budget FTE],Compare[Status],B{i})"
        ws.Range(f"F{i}").Formula = f"=SUMIFS(Compare[FTE difference],Compare[Status],B{i})"
        ws.Range(f"B{i}:F{i}").Interior.Color = C["card"] if i % 2 else C["alt"]
    ws.Range("B15").Value = "Total"
    for c in "CDEF":
        ws.Range(f"{c}15").Formula = f"=SUM({c}9:{c}14)"
    r = ws.Range("B15:F15"); r.Font.Bold = True; r.Font.Color = C["lgold"]; r.Borders(8).Color = C["gold"]
    ws.Range("C9:C15").NumberFormat = "0"
    ws.Range("D9:E15").NumberFormat = "0.00"
    ws.Range("F9:F15").NumberFormat = '+0.00;-0.00;"–"'
    fc = ws.Range("F9:F14").FormatConditions.Add(1, 6, "=0"); fc.Font.Color = C["bad"]
    fc = ws.Range("F9:F14").FormatConditions.Add(1, 5, "=0"); fc.Font.Color = C["amber"]
    ws.Range("B9").Font.Color = C["good"]

    head("B18", "Data checks (PASS / FAIL after every refresh)")
    log = "INDEX(CleaningLog[Rows],{})".format
    PC = ref(P["PC"])
    checks = [
        ("No HR row lost: rows in file = duplicates removed + terminated + compared",
         f'=IF({log(1)}={log(2)}+{log(6)}+COUNT(Compare[HR line]),"PASS","FAIL")'),
        ("Every active HR line compared exactly once",
         '=IF(SUMPRODUCT((Compare[HR line]<>"")/COUNTIF(Compare[HR line],Compare[HR line]&""))=COUNT(Compare[HR line]),"PASS","FAIL")'),
        ("Every budget line compared exactly once",
         f'=IF(AND(SUMPRODUCT((Compare[Budget line]<>"")/COUNTIF(Compare[Budget line],Compare[Budget line]&""))=COUNT(Compare[Budget line]),COUNT(Compare[Budget line])={log(8)}),"PASS","FAIL")'),
        ("Pivot cache, current month: HR FTE = HR FTE in the comparison",
         f'=IF(ROUND(GETPIVOTDATA("Total FTE",{PC},"Source","HR roster")-D15,2)=0,"PASS","FAIL")'),
        ("Pivot cache, current month: budget FTE = budget FTE in the comparison",
         f'=IF(ROUND(GETPIVOTDATA("Total FTE",{PC},"Source","Finance budget")-E15,2)=0,"PASS","FAIL")'),
        ("Pivot cache, current month: budget minus HR = sum of the line differences",
         f'=IF(ROUND(GETPIVOTDATA("FTE difference",{PC})-F15,2)=0,"PASS","FAIL")'),
        ("Dashboard bridge: budget FTE + every step = HR FTE card",
         '=IF(ROUND(WF_END-K_HR,2)=0,"PASS","FAIL")'),
        ("Dashboard heatmap: sum of the fund center totals = gap card",
         '=IF(ROUND(SUM(HeatRows)-K_GAP,2)=0,"PASS","FAIL")'),
    ]
    for i, (label, f) in enumerate(checks, 19):
        rng = ws.Range(f"B{i}:F{i}"); rng.Merge(); rng.Value = label
        rng.Interior.Color = C["card"] if i % 2 else C["alt"]
        cell = ws.Range(f"H{i}")
        cell.Formula = f
        cell.Font.Bold = True
        cell.HorizontalAlignment = -4108
        fc = cell.FormatConditions.Add(1, 3, '="PASS"'); fc.Interior.Color = C["goodbg"]; fc.Font.Color = C["good"]
        fc = cell.FormatConditions.Add(1, 4, '="PASS"'); fc.Interior.Color = C["badbg"]; fc.Font.Color = C["bad"]
    last = 18 + len(checks)
    add_name(wb, "CHK_PASS", f'=COUNTIF(Checks!$H$19:$H${last},"PASS")')
    add_name(wb, "CHK_N", f"={len(checks)}")
    # big badge
    badge = ws.Range("J8:N13"); badge.Merge()
    badge.Formula = '=CHK_PASS&" / "&CHK_N&CHAR(10)&"checks PASS"'
    badge.WrapText = True; badge.HorizontalAlignment = -4108; badge.VerticalAlignment = -4108
    badge.Font.Name = SERIF; badge.Font.Size = 28; badge.Interior.Color = C["goodbg"]; badge.Font.Color = C["good"]
    for e in (7, 8, 9, 10):
        badge.Borders(e).Color = C["good"]
    fc = badge.FormatConditions.Add(2, None, "=CHK_PASS<CHK_N"); fc.Interior.Color = C["badbg"]; fc.Font.Color = C["bad"]
    nb = ws.Range("J14:N15"); nb.Merge(); nb.WrapText = True; nb.VerticalAlignment = -4160
    nb.Value = ("Recalculated after every Data > Refresh All. The pivot checks use a control PivotTable on the same cache "
                "that no slicer touches, so they hold whatever you pick on the Dashboard.")
    nb.Font.Size = 8.5; nb.Font.Color = C["muted"]
    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = f"$A$1:$O${last + 1}"


def build_settings(wb):
    st = wb.Worksheets("Settings")
    rows = [("Data folder", '=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"data\\"', "DataFolder",
             "Folder with the CSV exports. Default: the data folder next to this workbook. Type a path to override."),
            ("HR roster file", "hr-roster-2026-10.csv", "HRFile", "Current month, HR system export: Employee ID, Employee Name, Fund Center, Functional Area, FTE, Status"),
            ("Finance budget file", "finance-budget-2026-10.csv", "BudgetFile", "Current month, position budget export: Position, Emp No, Employee, Cost Center, Function, Budgeted FTE"),
            ("Period", "October 2026", "Period", "Shown in the report titles")]
    for r, (label, val, nm, note) in enumerate(rows, 8):
        st.Cells(r, 2).Value = label
        if val.startswith("="):
            st.Cells(r, 3).Formula2 = val
        else:
            st.Cells(r, 3).NumberFormat = "@"
            st.Cells(r, 3).Value = val
        st.Cells(r, 4).Value = note
        add_name(wb, nm, f"=Settings!$C${r}")


def style_settings(wb):
    st = wb.Worksheets("Settings")
    st.Columns("A").ColumnWidth = 2; st.Columns("B").ColumnWidth = 22; st.Columns("C").ColumnWidth = 60
    st.Columns("D").ColumnWidth = 100
    header(wb, st, "Settings", "Where the CSV exports live and which month is current. The trend reads every "
                               "hr-roster-YYYY-MM.csv / finance-budget-YYYY-MM.csv pair in the folder.", st.Range("E1").Left)
    for c, t in zip("BCD", ("Setting", "Value", "Note")):
        h = st.Range(f"{c}7"); h.Value = t; h.Font.Bold = True; h.Font.Color = C["gold"]; h.Interior.Color = C["head"]
    v = st.Range("C8:C11"); v.Interior.Color = C["alt"]; v.Font.Color = C["lgold"]
    for e in (7, 8, 9, 10, 12):
        v.Borders(e).Color = C["gold"]
    st.Range("D8:D11").Font.Color = C["muted"]


def build_howto(wb):
    ws = wb.Worksheets("How to use")
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 150
    header(wb, ws, "How to use", "Every month: drop in the new CSV files, Data > Refresh All, check the PASS badge, "
                                 "work through the Discrepancies tab.", ws.Range("C1").Left)
    text = [
        ("Every month", True),
        ("1. Export the HR roster and the Finance position budget as CSV into the data folder (see Settings), named "
         "hr-roster-YYYY-MM.csv and finance-budget-YYYY-MM.csv. Keep the earlier months there: they feed the trend.", False),
        ("2. On Settings, type the two file names of the current month.", False),
        ("3. Data > Refresh All. The comparison, the three PivotTables, the dashboard and the checks all update. No macros.", False),
        ("4. Dashboard: pick the month in the Month slicer; the 'checks PASS' badge must show every check passing.", False),
        ("5. Work through the Discrepancies tab: every line has what to do.", False), ("", False),
        ("The Dashboard", True),
        ("Slicers: Month, Fund center, Functional area, Discrepancy type. They drive the KPI cards, the bridge, the heatmap, "
         "the donut, the charts, the insight sentences AND the three PivotTables (same slicers on the Pivots tab).", False),
        ("The trend chart follows every slicer except Month (it always shows every month). Pick one month at a time.", False),
        ("Bridge: from the budget FTE to the HR FTE, one bar per discrepancy type (blue = HR higher, copper = budget higher).", False),
        ("Heatmap: budget minus HR per fund center and functional area. Red = more people than budget, blue = budget not filled.", False),
        ("Insights: plain sentences written by formulas from the same numbers; they change with the slicers.", False), ("", False),
        ("What each status means", True),
        ("Matched: same person, same fund center, same functional area, same FTE in both sources.", False),
        ("FTE differs: same person and fund center, different FTE (e.g. HR 1.00, budget 0.80).", False),
        ("Functional area differs (Area differs): same person and fund center, booked to a different functional area.", False),
        ("Fund center differs (Fund center moved): the person is in both sources but under a different fund center.", False),
        ("Only in HR roster (Only in HR): an active employee with no budgeted position.", False),
        ("Only in budget: a vacant position (Vacant post), someone terminated in HR (Terminated) or a name HR does not know (Unknown to HR).", False),
        ("", False),
        ("How the two sources are lined up", True),
        ("Employee IDs: E01042 (HR) and 1042 (budget) become the same key. Fund centers: fc210, FC-210, 210 all become FC210.", False),
        ("FTE: 80% (HR) and 0.8 (budget) are both 0.80. Functional areas are spelled the way HR spells it.", False),
        ("Terminated people stay in the HR file but are not counted as headcount; every step is counted on the Cleaning log tab.", False),
        ("The steps live in Power Query (Data > Queries & Connections): fnHRClean, fnBudgetClean, fnCompare, History, FTE_Long.", False),
    ]
    for i, (t, bold) in enumerate(text, 7):
        c = ws.Cells(i, 2)
        c.Value = t
        if bold:
            c.Font.Name = SERIF; c.Font.Size = 13; c.Font.Color = C["lgold"]
        else:
            c.Font.Size = 10


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
        for i, title in enumerate(SHEETS, 1):
            wb.Worksheets(i).Name = title
        add_styles(wb)
        for title in SHEETS:
            if title != "Dash data":
                page(wb.Worksheets(title))
        build_settings(wb)
        wb.SaveAs(str(OUT), 51)  # CELL("filename") needs a saved file
        xl.Calculate()
        for name, formula in Q.items():
            for attempt in range(3):  # the Power Query pane sometimes throws on the first try
                try:
                    wb.Queries.Add(name, formula)
                    break
                except Exception:
                    if attempt == 2:
                        raise
                    time.sleep(2)
        print("queries added", flush=True)
        t0 = time.time()
        for title, query in LOADED.items():
            load_query(wb.Worksheets(title), query, query.replace("_", ""))
        print(f"queries loaded in {time.time() - t0:.1f}s")
        format_tables(wb)
        pc, P = build_pivots(wb)
        print(f"pivots in {time.time() - t0:.1f}s", flush=True)
        card_row = build_dash_data(wb, P)
        print("dash data", flush=True)
        build_checks(wb, P)  # defines CHK_PASS / CHK_N used by the dashboard
        dashboard_grid(wb)
        piv = wb.Worksheets("Pivot comparison")
        header(wb, piv, "Pivot comparison",
               "Three PivotTables on one Power Query source: HR roster, Finance budget, budget minus HR. "
               "The slicers here are the Dashboard's slicers.", piv.Range("X1").Left)
        build_slicers(wb, P)
        print("slicers", flush=True)
        build_dashboard(wb, P, card_row)
        print("dashboard", flush=True)
        ps = piv.PageSetup
        ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
        ps.PrintArea = "$A$1:$X$29"
        piv.Range("A26:X40").Interior.Color = C["page"]  # cells a shrinking pivot leaves behind
        style_settings(wb)
        build_howto(wb)
        dd = wb.Worksheets("Dash data")
        dd.Tab.Color = C["grey"]
        wb.Worksheets("Dashboard").Tab.Color = C["gold"]
        xl.Calculate()
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            xl.ActiveWindow.Zoom = 100
            ws.Range("A1").Select()
        wb.Worksheets("Dashboard").Activate()
        xl.ActiveWindow.DisplayHeadings = False
        wb.Save()
        print("saved", OUT.name, f"{time.time() - t0:.1f}s")
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        xl.Quit()


if __name__ == "__main__":
    main()
