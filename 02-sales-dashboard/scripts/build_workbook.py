"""Builds Sales-Dashboard.xlsx with real Excel (COM).

Power Query loads sales lines and the budget from CSV. Slicers (Region, Category, Sales rep, MTD/YTD view) and a
Month timeline sit on small PivotTables on the "Slicer engine" sheet; a 1/0 "In" column in the Sales and Budget tables
reads what those PivotTables show, so every number on the six report pages stays a plain, traceable SUMIFS on the
Calc sheet. Pages: Overview, Products, Customers, Sales reps, Forecast, Checks - dark-gold design, navigation buttons.
No macros. New data: Data > Refresh All.

Run with a Python that has pywin32:  python scripts/build_workbook.py
"""
import re, shutil, time, zipfile
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Sales-Dashboard.xlsx"
TIMELINE_SCROLL = "2025-10-01"   # first month visible in the timeline
DEFAULT_MONTH = ("2026-10-01", "2026-10-31")


def hexrgb(h):
    h = h.lstrip("#")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


P = {k: hexrgb(v) for k, v in dict(
    bg0="0B0F13", bg="12171D", card="1A2129", card2="202832", line="2A333D", gold="C9A35B", gold2="E6CF9A",
    ivory="ECE6DA", muted="8E98A3", jade="4FA38C", copper="B5754A", sky="7FA7C9", bgrey="4A5560", good="5BBF8A",
    bad="E5735F", amber="D9A441", goodbg="1E3A2F", badbg="3D2524", navsel="3A3121", heatbad="A9473B", heatgood="3E8E68",
    light="F4F1EA", lightline="D9D3C7", ink="26303A").items()}
SERIES = [P["gold"], P["jade"], P["copper"], P["sky"], P["gold2"]]
SERIF, SANS = "Georgia", "Segoe UI"

CATS = ["Fasteners", "Power Tools", "Paint", "Garden", "Electrical"]
REGIONS = ["North", "South", "East", "West"]
PAGES = [("Dashboard", "Overview", "Sales overview"), ("Products", "Products", "Products & categories"),
         ("Customers", "Customers", "Customers & the 80/20 rule"), ("Reps", "Sales reps", "Sales rep scorecard"),
         ("Forecast", "Forecast", "Forecast to year end"), ("Checks", "Checks", "Data checks")]
SHEETS = [p[0] for p in PAGES] + ["Calc", "Sales data", "Budget data", "Slicer engine", "Settings", "How to use"]

# ------------------------------------------------------------------------------------------------- Power Query
Q = {}
Q["Setting"] = """(name as text) as text =>
    Text.From(Excel.CurrentWorkbook(){[Name = name]}[Content]{0}[Column1])"""
Q["Sales_Raw"] = """let
    Source = Csv.Document(File.Contents(Setting("DataFolder") & Setting("SalesFile")), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv])
in
    Table.PromoteHeaders(Source, [PromoteAllScalars = true])"""
Q["Budget_Raw"] = """let
    Source = Csv.Document(File.Contents(Setting("DataFolder") & Setting("BudgetFile")), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv])
in
    Table.PromoteHeaders(Source, [PromoteAllScalars = true])"""
Q["Sales"] = """let
    Source  = Sales_Raw,
    Typed   = Table.TransformColumns(Source, {
                  {"Date", each Date.FromText(_, [Format = "yyyy-MM-dd", Culture = "en-US"]), type date},
                  {"Qty", each Number.FromText(_, "en-US"), Int64.Type},
                  {"Unit Price", each Number.FromText(_, "en-US"), type number},
                  {"Unit Cost", each Number.FromText(_, "en-US"), type number}}),
    Month   = Table.AddColumn(Typed, "Month", each Date.StartOfMonth([Date]), type date),
    MonthNo = Table.AddColumn(Month, "Month No", each Date.Year([Date]) * 100 + Date.Month([Date]), Int64.Type),
    Revenue = Table.AddColumn(MonthNo, "Revenue", each Number.Round([Qty] * [Unit Price], 2), type number),
    Cost    = Table.AddColumn(Revenue, "Cost", each Number.Round([Qty] * [Unit Cost], 2), type number),
    Margin  = Table.AddColumn(Cost, "Margin", each Number.Round([Revenue] - [Cost], 2), type number),
    // one invoice = one order: flag its first line, and list the invoice's categories ("|Paint|Garden|")
    // so an order still counts when the category slicer keeps only some of its lines
    Grouped = Table.Group(Margin, {"Invoice"}, {
                  {"Rows", each Table.AddIndexColumn(_, "Line", 1, 1, Int64.Type), type table},
                  {"Inv Categories", each "|" & Text.Combine(List.Distinct([Category]), "|") & "|", type text}}),
    Expand  = Table.ExpandTableColumn(Grouped, "Rows", {"Date", "Customer", "Region", "Sales Rep", "Category", "Qty", "Unit Price",
                  "Unit Cost", "Month", "Month No", "Revenue", "Cost", "Margin", "Line"}),
    Order   = Table.AddColumn(Expand, "Order", each if [Line] = 1 then 1 else 0, Int64.Type),
    Sorted  = Table.Sort(Order, {{"Date", Order.Ascending}, {"Invoice", Order.Ascending}, {"Line", Order.Ascending}}),
    Cols    = Table.ReorderColumns(Table.RemoveColumns(Sorted, {"Line"}), {"Date", "Invoice", "Customer", "Region", "Sales Rep",
                  "Category", "Qty", "Unit Price", "Unit Cost", "Month", "Month No", "Revenue", "Cost", "Margin", "Order", "Inv Categories"}),
    Result  = Table.TransformColumnTypes(Cols, {{"Date", type date}, {"Month", type date}, {"Month No", Int64.Type},
                  {"Customer", type text}, {"Region", type text}, {"Sales Rep", type text}, {"Category", type text},
                  {"Qty", Int64.Type}, {"Unit Price", type number}, {"Unit Cost", type number}, {"Revenue", type number},
                  {"Cost", type number}, {"Margin", type number}, {"Inv Categories", type text}})
in
    Result"""
Q["Budget"] = """let
    Source = Budget_Raw,
    Typed  = Table.TransformColumns(Source, {
                 {"Month", each Date.FromText(_ & "-01", [Format = "yyyy-MM-dd", Culture = "en-US"]), type date},
                 {"Budget Revenue", each Number.FromText(_, "en-US"), type number},
                 {"Budget Margin %", each Number.FromText(_, "en-US"), type number}}),
    GP     = Table.AddColumn(Typed, "Budget GP", each Number.Round([Budget Revenue] * [#"Budget Margin %"], 2), type number)
in
    GP"""
Q["Customers"] = """let
    Source = Table.Distinct(Table.SelectColumns(Sales, {"Customer", "Region"}))
in
    Table.Sort(Source, {"Customer"})"""
Q["Reps"] = """let
    Source = Table.Distinct(Table.SelectColumns(Sales, {"Sales Rep", "Region"}))
in
    Table.Sort(Source, {{"Region", Order.Ascending}, {"Sales Rep", Order.Ascending}})"""
Q["Checks"] = """let
    S       = Sales,
    B       = Budget,
    Keys    = Table.Distinct(Table.SelectColumns(S, {"Month", "Sales Rep", "Category"})),
    NoBud   = Table.NestedJoin(Keys, {"Month", "Sales Rep", "Category"}, B, {"Month", "Sales Rep", "Category"}, "B", JoinKind.LeftAnti),
    BudKeys = Table.RowCount(Table.Distinct(Table.SelectColumns(B, {"Month", "Sales Rep", "Category"})))
in
    #table(type table [Check = text, Value = number], {
        {"Sales lines loaded", Table.RowCount(S)},
        {"Invoices (orders)", List.Sum(S[Order])},
        {"Month/rep/category combinations sold without a budget", Table.RowCount(NoBud)},
        {"Budget rows", Table.RowCount(B)},
        {"Budget rows that repeat a month/rep/category", Table.RowCount(B) - BudKeys},
        {"Total revenue, all months", List.Sum(S[Revenue])}})"""


def load_query(ws, name, at="A1"):
    conn = f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=""'
    lo = ws.ListObjects.Add(0, conn, None, 1, ws.Range(at))
    qt = lo.QueryTable
    qt.CommandType = 2
    qt.CommandText = f"SELECT * FROM [{name}]"
    qt.RefreshStyle = 1
    qt.AdjustColumnWidth = False
    qt.BackgroundQuery = False  # Refresh All: tables first, then the PivotTables behind the slicers
    qt.Refresh(False)
    lo.Name = name
    lo.TableStyle = "TableStyleLight1"
    return lo


def main():
    """E2_STAGE=pages (developer shortcut): reuse the workbook saved after the Calc step and rebuild only the pages."""
    import os
    stage_file = Path(os.environ.get("TEMP", ".")) / "e2-stage.xlsx"
    pages_only = os.environ.get("E2_STAGE") == "pages" and stage_file.exists()
    if OUT.exists():
        OUT.unlink()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    t0 = time.time()
    log = lambda what: print(f"{time.time() - t0:6.1f}s  {what}", flush=True)
    try:
        if pages_only:
            shutil.copy(stage_file, OUT)
            wb = xl.Workbooks.Open(str(OUT))
            xl.Calculation = -4135
            caches = {k: wb.SlicerCaches(n) for k, n in (("Region", "Slicer_Region"), ("Category", "Slicer_Category"),
                      ("Rep", "Slicer_Rep"), ("View", "Slicer_View"), ("Month", "Timeline_Month"))}
            log("stage workbook opened")
        else:
            wb = xl.Workbooks.Add()
            while wb.Worksheets.Count < len(SHEETS):
                wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            for i, n in enumerate(SHEETS, 1):
                wb.Worksheets(i).Name = n
            build_settings(wb)
            wb.SaveAs(str(OUT), 51)
            xl.Calculate()
            for name, m in Q.items():
                wb.Queries.Add(name, m)
            load_query(wb.Worksheets("Sales data"), "Sales")
            load_query(wb.Worksheets("Budget data"), "Budget")
            calc = wb.Worksheets("Calc")
            load_query(calc, "Customers", "AN3")
            load_query(calc, "Reps", "BH3")
            load_query(calc, "Checks", "CA3")
            log("queries loaded")
            xl.Calculation = -4135  # manual while building
            caches = build_engine(wb)
            log("slicer engine")
            format_data(wb)
            build_calc(wb)
            log("calc")
            make_styles(wb)
            wb.Save()
            shutil.copy(OUT, stage_file)
        probe_separator(wb)
        for sheet, nav, title in PAGES:
            frame(wb, sheet, title, caches, slicers=sheet != "Checks")
            log("frame " + sheet)
        build_overview(wb); log("overview")
        build_products(wb); log("products")
        build_customers(wb); log("customers")
        build_reps(wb); log("reps")
        build_forecast(wb); log("forecast")
        build_checks(wb)
        build_howto(wb)
        xl.Calculation = -4105
        xl.CalculateFull()
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            if ws.Name in [p[0] for p in PAGES]:
                xl.ActiveWindow.DisplayHeadings = False
                xl.ActiveWindow.Zoom = 80
            ws.Range("A1").Select()
            xl.ActiveWindow.ScrollRow = 1; xl.ActiveWindow.ScrollColumn = 1
        wb.Worksheets("Dashboard").Activate()
        xl.ScreenUpdating = True
        wb.Save()
        print("saved", OUT.name)
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        xl.Quit()
    post_xml()


def set_attrs(tag, attrs):
    """Set (or replace) XML attributes on one start tag, e.g. <slicer name=... />."""
    for k, v in attrs.items():
        tag = re.sub(rf'\s{k}="[^"]*"', "", tag)
        tag = tag.replace(" ", f' {k}="{v}" ', 1)
    return tag


def post_xml():
    """Slicer and timeline layout written straight into the file.

    In this workbook Excel stops responding whenever a slicer's layout property (columns, row height, style) is set
    through COM, so those settings go into xl/slicers/*.xml and xl/timelines/*.xml after saving - the same attributes
    Excel writes itself. The timeline's scroll position has no COM property at all.
    """
    tmp = OUT.with_suffix(".tmp")
    cols = {k: n for k, _, n, _ in SIDEBAR}
    with zipfile.ZipFile(OUT) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith(("xl/slicers/", "xl/timelines/")) and item.filename.endswith(".xml"):
                s = data.decode("utf-8")

                def fix(m):
                    tag, name = m.group(0), m.group(2)
                    key, _, sheet = name.partition("_")
                    if sheet == "engine":
                        return tag
                    if m.group(1) == "slicer" and key in cols:
                        return set_attrs(tag, {"columnCount": cols[key], "rowHeight": 17 * 12700, "style": "Gold Slicer"})
                    if m.group(1) == "timeline":
                        return set_attrs(tag, {"style": "Gold Timeline", "showHeader": 1, "showSelectionLabel": 1, "showTimeLevel": 0,
                                               "showHorizontalScrollbar": 0, "scrollPosition": TIMELINE_SCROLL + "T00:00:00"})
                    return tag
                s = re.sub(r'<(slicer|timeline)\s+name="([^"]+)"[^>]*?/?>', fix, s)
                data = s.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, OUT)


# ------------------------------------------------------------------------------------------------- Settings / data
def build_settings(wb):
    st = wb.Worksheets("Settings")
    st.Range("A1").Value = "Settings"
    st.Range("A1").Font.Size = 16; st.Range("A1").Font.Bold = True
    rows = [("Data folder", '=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"data\\"', "DataFolder",
             "Folder with the CSV exports. Default: the data folder next to this workbook. Type a path to override."),
            ("Sales file", "sales-lines.csv", "SalesFile", "One row per invoice line: date, invoice, customer, region, rep, category, qty, price, cost"),
            ("Budget file", "budget.csv", "BudgetFile", "One row per month, sales rep and category (the full year, so the forecast has a target)"),
            ("Company", "Brightline Supply Co.", "Company", "Shown in the page titles (fictional company)")]
    st.Range("A3:C3").Value = ("Setting", "Value", "Note")
    st.Range("A3:C3").Font.Bold = True
    for r, (label, val, nm, note) in enumerate(rows, 4):
        st.Cells(r, 1).Value = label
        c = st.Cells(r, 2)
        if val.startswith("="):
            c.Formula2 = val  # Formula2: no implicit-intersection @ (keeps Excel 2016 happy)
        else:
            c.NumberFormat = "@"
            c.Value = val
        c.Interior.Color = hexrgb("FFF2CC")
        st.Cells(r, 3).Value = note
        wb.Names.Add(nm, f"=Settings!$B${r}")
    st.Columns("A").ColumnWidth = 16; st.Columns("B").ColumnWidth = 50; st.Columns("C").ColumnWidth = 90


def light_header(rng):
    rng.Interior.Color = P["ink"]; rng.Font.Color = P["gold2"]; rng.Font.Bold = True


def format_data(wb):
    ws = wb.Worksheets("Sales data")
    lo = ws.ListObjects("Sales")
    for c, f in (("Date", "dd mmm yyyy"), ("Month", "mmm yyyy"), ("Month No", "0")):
        lo.ListColumns(c).DataBodyRange.NumberFormat = f
    for c in ("Unit Price", "Unit Cost", "Revenue", "Cost", "Margin"):
        lo.ListColumns(c).DataBodyRange.NumberFormat = "#,##0.00"
    widths = {"Customer": 26, "Sales Rep": 13, "Category": 12, "Invoice": 11, "Inv Categories": 34, "Date": 12, "Qty": 6,
              "Order": 7, "In": 6, "Order In": 9, "Month No": 10}
    for i in range(1, lo.ListColumns.Count + 1):
        ws.Columns(i).ColumnWidth = widths.get(lo.ListColumns(i).Name, 11)
    light_header(lo.HeaderRowRange)
    ws.Activate(); ws.Range("A2").Select(); wb.Application.ActiveWindow.FreezePanes = True
    ws = wb.Worksheets("Budget data")
    lo = ws.ListObjects("Budget")
    lo.ListColumns("Month").DataBodyRange.NumberFormat = "mmm yyyy"
    lo.ListColumns("Budget Revenue").DataBodyRange.NumberFormat = "#,##0"
    lo.ListColumns("Budget GP").DataBodyRange.NumberFormat = "#,##0"
    lo.ListColumns("Budget Margin %").DataBodyRange.NumberFormat = "0%"
    light_header(lo.HeaderRowRange)
    for i in range(1, lo.ListColumns.Count + 1):
        ws.Columns(i).ColumnWidth = 15


# ------------------------------------------------------------------------------------------------- Slicer engine
def build_engine(wb):
    """Small PivotTables that the slicers and the timeline filter. Their visible row items are what the formulas read."""
    ws = wb.Worksheets("Slicer engine")
    ws.Range("A1").Value = "Slicer engine: each slicer filters a one-field PivotTable below; the Sales/Budget column 'In' = 1 when a row's region, rep and category are all visible here."
    ws.Range("A1").Font.Bold = True
    ws.Range("O3").Value = "View"; ws.Range("O4").Value = "MTD"; ws.Range("O5").Value = "YTD"
    vt = ws.ListObjects.Add(1, ws.Range("O3:O5"), None, 1); vt.Name = "ViewOptions"; vt.TableStyle = "TableStyleLight1"
    pc = wb.PivotCaches().Create(1, "Sales")
    pts = {}
    for col, (name, field) in zip("ACEG", (("PT_Region", "Region"), ("PT_Category", "Category"), ("PT_Rep", "Sales Rep"),
                                          ("PT_Month", "Month No"))):
        pt = pc.CreatePivotTable(ws.Range(f"{col}3"), name)
        pt.PivotFields(field).Orientation = 1
        pts[name] = pt
    pv = wb.PivotCaches().Create(1, "ViewOptions").CreatePivotTable(ws.Range("I3"), "PT_View")
    pv.PivotFields("View").Orientation = 1
    pts["PT_View"] = pv
    for pt in pts.values():
        pt.ColumnGrand = False; pt.RowGrand = False; pt.RowAxisLayout(1)
        pt.TableStyle2 = "PivotStyleLight1"
    for nm, col in (("SelRegions", "A"), ("SelCats", "C"), ("SelReps", "E"), ("SelMonths", "G"), ("SelView", "I")):
        wb.Names.Add(nm, f"='Slicer engine'!${col}$4:${col}$80")
    caches = {}
    for key, field, pt in (("Region", "Region", "PT_Region"), ("Category", "Category", "PT_Category"), ("Rep", "Sales Rep", "PT_Rep")):
        sc = wb.SlicerCaches.Add2(pts[pt], field, f"Slicer_{key}")
        for other in ("PT_Region", "PT_Category", "PT_Rep"):
            if other != pt:
                sc.PivotTables.AddPivotTable(pts[other])
        caches[key] = sc
    caches["View"] = wb.SlicerCaches.Add2(pts["PT_View"], "View", "Slicer_View")
    caches["Month"] = wb.SlicerCaches.Add2(pts["PT_Month"], "Month", "Timeline_Month", 2)
    # one slicer of each on this sheet too (a slicer cache without any slicer is dropped when the file is saved)
    x = ws.Range("Q3").Left
    for k, key in enumerate(("View", "Region", "Category", "Rep")):
        caches[key].Slicers.Add(SlicerDestination=ws, Name=f"{key}_engine", Caption=key, Top=ws.Range("Q3").Top,
                                Left=x + k * 130, Width=120, Height=150)
    caches["Month"].Slicers.Add(SlicerDestination=ws, Name="Month_engine", Caption="Month", Top=ws.Range("Q16").Top,
                                Left=x, Width=500, Height=110)
    caches["View"].SlicerItems("YTD").Selected = False
    caches["Month"].TimelineState.SetFilterDateRange(*DEFAULT_MONTH)
    ws.Range("A84").Value = "View: MTD = the last month picked on the timeline; YTD = January to that month. Pick one."
    ws.Columns("A:O").ColumnWidth = 14

    # 1/0 filter columns used by every SUMIFS
    sales = wb.Worksheets("Sales data").ListObjects("Sales")
    for name, f in (("In", "=IF(AND(COUNTIF(SelRegions,[@Region])>0,COUNTIF(SelReps,[@[Sales Rep]])>0,COUNTIF(SelCats,[@Category])>0),1,0)"),
                    ("Order In", '=IF([@Order]=1,IF(AND(COUNTIF(SelRegions,[@Region])>0,COUNTIF(SelReps,[@[Sales Rep]])>0,'
                                 'SUMPRODUCT(--ISNUMBER(SEARCH("|"&SelCats&"|",[@[Inv Categories]])))>0),1,0),0)')):
        col = sales.ListColumns.Add(); col.Name = name; col.DataBodyRange.Formula = f
    bud = wb.Worksheets("Budget data").ListObjects("Budget")
    col = bud.ListColumns.Add(); col.Name = "In"
    col.DataBodyRange.Formula = "=IF(AND(COUNTIF(SelRegions,[@Region])>0,COUNTIF(SelReps,[@[Sales Rep]])>0,COUNTIF(SelCats,[@Category])>0),1,0)"
    return caches


# ------------------------------------------------------------------------------------------------- Calc
def SR(col, s, e, crit=""):
    return f'SUMIFS(Sales[{col}],Sales[Month],">="&{s},Sales[Month],"<="&{e},Sales[In],1{crit})'


def SO(s, e, crit=""):
    return f'SUMIFS(Sales[Order In],Sales[Month],">="&{s},Sales[Month],"<="&{e}{crit})'


def SB(col, s, e, crit=""):
    return f'SUMIFS(Budget[{col}],Budget[Month],">="&{s},Budget[Month],"<="&{e},Budget[In],1{crit})'


def put(ws, addr, val, fmt=None):
    r = ws.Range(addr)
    if isinstance(val, str) and val.startswith("="):
        r.Formula2 = val  # Formula2: no implicit-intersection @ added by Excel 365
    else:
        r.Value = val
    if fmt:
        r.NumberFormat = fmt


def name(wb, nm, addr):
    wb.Names.Add(nm, f"=Calc!{abs_(addr)}")


def abs_(addr):
    return re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", addr)


PREDEFINED = ["RecentAch", "RecentAvg", "FcYTD", "FcFYB", "FcFY", "FcVsB", "FcLeft", "FcStraight", "FcStrVsB", "FcPY", "FcVsPY", "FcYTDB",
              "FcYTDvsB", "FcYTDPY", "FcYTDvsPY", "FcGap", "FcRest", "CustTotal", "CustActive", "CustOnFile", "CustN80", "CustN80Pct",
              "CustTop10", "CustTopName", "CustTopShare", "CustAvg", "CustActivePY", "CustAvgPY", "CustAvgVsPY", "CustActiveVsPY",
              "CustTop10VsPY", "RepsWithTarget", "RepsOnTarget", "RepLow", "RepLowPct", "RepTeamPct"]
PCT, PCTS, NUM, NUMS = "0.0%", "+0.0%;-0.0%;0.0%", "#,##0", "+#,##0;-#,##0;0"
TP, TPS, TN, TNS = "0.0%;-0.0%;–", "+0.0%;-0.0%;–", "#,##0;-#,##0;–", "+#,##0;-#,##0;–"   # page tables: zero = dash


def build_calc(wb):
    c = wb.Worksheets("Calc")
    for nm in PREDEFINED:   # placeholders so formulas can use a name before its cell is written; redefined below
        wb.Names.Add(nm, "=0")
    c.Cells.Font.Size = 10
    put(c, "C2", '="revenue, "&IF(ViewMode="YTD","year to date","month")')
    c.Range("A1").Value = "Calc: every number on the report pages comes from this sheet (SUMIFS on the Sales and Budget tables, filtered by the slicers through the In column)"
    c.Range("A1").Font.Bold = True; c.Range("A1").Font.Size = 12

    # ---- selection  A3:B20
    sel = [("RepMonth", "Reporting month (last month on the timeline)",
            '=IF(COUNT(SelMonths)=0,MAX(Sales[Month]),DATE(INT(MAX(SelMonths)/100),MOD(MAX(SelMonths),100),1))', "mmm yyyy"),
           ("ViewMode", "View (MTD or YTD)", '=IF(AND(COUNTIF(SelView,"YTD")>0,COUNTIF(SelView,"MTD")=0),"YTD","MTD")', None),
           ("PStart", "Period start", '=IF(ViewMode="YTD",DATE(YEAR(RepMonth),1,1),RepMonth)', "mmm yyyy"),
           ("PEnd", "Period end", "=RepMonth", "mmm yyyy"),
           ("PYStart", "Prior-year start", "=EDATE(PStart,-12)", "mmm yyyy"),
           ("PYEnd", "Prior-year end", "=EDATE(PEnd,-12)", "mmm yyyy"),
           ("PrevMonth", "Prior month", "=EDATE(RepMonth,-1)", "mmm yyyy"),
           ("PeriodLabel", "Period label", '=IF(ViewMode="YTD","Jan-"&TEXT(RepMonth,"mmm yyyy"),TEXT(RepMonth,"mmmm yyyy"))', None),
           ("PYLabel", "Prior-year label", '=IF(ViewMode="YTD","Jan-"&TEXT(PYEnd,"mmm yyyy"),TEXT(PYEnd,"mmm yyyy"))', None),
           ("MonLabel", "Month short", '=TEXT(RepMonth,"mmm")', None),
           ("PrevLabel", "Prior month short", '=TEXT(PrevMonth,"mmm")', None),
           ("NReg", "Regions selected", "=COUNTA(SelRegions)", None),
           ("NCat", "Categories selected", "=COUNTA(SelCats)", None),
           ("NRep", "Sales reps selected", "=COUNTA(SelReps)", None),
           ("SelText", "Selection in words",
            '=PeriodLabel&IF(ViewMode="YTD"," (year to date)","")&"  |  "&IF(NReg=4,"all regions",IF(NReg=1,INDEX(SelRegions,1),NReg&" regions"))'
            '&"  |  "&IF(NCat=5,"all categories",IF(NCat=1,INDEX(SelCats,1),NCat&" categories"))'
            '&"  |  "&IF(NRep=ROWS(Reps[Sales Rep]),"all reps",IF(NRep=1,INDEX(SelReps,1),NRep&" reps"))', None),
           ("FYear", "Forecast year", "=YEAR(RepMonth)", "0")]
    c.Range("A3:B3").Value = ("Selection", "Value")
    for i, (nm, label, f, fmt) in enumerate(sel, 4):
        c.Range(f"A{i}").Value = label
        put(c, f"B{i}", f, fmt)
        name(wb, nm, f"B{i}")

    # ---- KPI matrix  D3:H12   columns: period, prior year, reporting month, prior month
    c.Range("D3:H3").Value = ("Measure", "Period", "Prior year", "Month", "Prior month")
    per = [("PStart", "PEnd"), ("PYStart", "PYEnd"), ("RepMonth", "RepMonth"), ("PrevMonth", "PrevMonth")]
    rows = [("Revenue", lambda s, e: SR("Revenue", s, e), NUM),
            ("Gross profit", lambda s, e: SR("Margin", s, e), NUM),
            ("Orders", lambda s, e: SO(s, e), NUM),
            ("Budget revenue", lambda s, e: SB("Budget Revenue", s, e), NUM),
            ("Budget gross profit", lambda s, e: SB("Budget GP", s, e), NUM)]
    for i, (label, f, fmt) in enumerate(rows, 4):
        c.Range(f"D{i}").Value = label
        for col, (s, e) in zip("EFGH", per):
            put(c, f"{col}{i}", "=" + f(s, e), fmt)
    for i, (label, f, fmt) in enumerate([("Gross margin %", "IFERROR({c}5/{c}4,0)", PCT), ("Budget margin %", "IFERROR({c}8/{c}7,0)", PCT),
                                         ("Average order", "IFERROR({c}4/{c}6,0)", NUM), ("Budget achieved", "IFERROR({c}4/{c}7,0)", PCT)], 9):
        c.Range(f"D{i}").Value = label
        for col in "EFGH":
            put(c, f"{col}{i}", "=" + f.format(c=col), fmt)
    # deltas  D14:G20   (blank when there is nothing to compare with)
    c.Range("D14:G14").Value = ("Change", "vs budget", "vs prior year", "month vs prior month")
    d = [("Revenue %", "IF(E7=0,\"\",E4/E7-1)", "IF(F4=0,\"\",E4/F4-1)", "IF(H4=0,\"\",G4/H4-1)", PCTS),
         ("Gross profit %", "IF(E8=0,\"\",E5/E8-1)", "IF(F5=0,\"\",E5/F5-1)", "IF(H5=0,\"\",G5/H5-1)", PCTS),
         ("Gross margin (pts)", "IF(OR(E4=0,E7=0),\"\",E9-E10)", "IF(OR(E4=0,F4=0),\"\",E9-F9)", "IF(OR(G4=0,H4=0),\"\",G9-H9)", "+0.0%;-0.0%"),
         ("Orders %", "\"\"", "IF(F6=0,\"\",E6/F6-1)", "IF(H6=0,\"\",G6/H6-1)", PCTS),
         ("Average order %", "\"\"", "IF(F11=0,\"\",E11/F11-1)", "IF(H11=0,\"\",G11/H11-1)", PCTS),
         ("Budget achieved (pts)", "\"\"", "IF(OR(F7=0,F4=0),\"\",E12-F12)", "IF(OR(H7=0,H4=0),\"\",G12-H12)", "+0.0%;-0.0%")]
    for i, (label, a, b, m, fmt) in enumerate(d, 15):
        c.Range(f"D{i}").Value = label
        for col, f in zip("EFG", (a, b, m)):
            put(c, f"{col}{i}", "=" + f, fmt)

    # ---- 12-month trend  K2:AK15  (row 2 = keys used by the formulas)
    hdr = ["Month", "Label", "Actual", "Budget", "Prior year", "Gross profit", "Margin %", "Orders", "Avg order", "Achieved"]
    c.Range("K3:T3").Value = tuple(hdr)
    for j, reg in enumerate(REGIONS):
        for base, tag in ((21, "Act "), (25, "Bud "), (29, "Var% ")):
            c.Cells(2, base + j).Value = reg; c.Cells(3, base + j).Value = tag + reg
    for j, cat in enumerate(CATS):
        c.Cells(2, 33 + j).Value = cat; c.Cells(3, 33 + j).Value = "Act " + cat
    for i in range(12):
        r = 4 + i
        put(c, f"K{r}", f"=EDATE(RepMonth,{i - 11})", "mmm yyyy")
        put(c, f"L{r}", f'=TEXT(K{r},"mmm yy")')
        put(c, f"M{r}", f"=SUMIFS(Sales[Revenue],Sales[Month],K{r},Sales[In],1)", NUM)
        put(c, f"N{r}", f"=SUMIFS(Budget[Budget Revenue],Budget[Month],K{r},Budget[In],1)", NUM)
        put(c, f"O{r}", f"=SUMIFS(Sales[Revenue],Sales[Month],EDATE(K{r},-12),Sales[In],1)", NUM)
        put(c, f"P{r}", f"=SUMIFS(Sales[Margin],Sales[Month],K{r},Sales[In],1)", NUM)
        put(c, f"Q{r}", f"=IFERROR(P{r}/M{r},0)", PCT)
        put(c, f"R{r}", f"=SUMIFS(Sales[Order In],Sales[Month],K{r})", NUM)
        put(c, f"S{r}", f"=IFERROR(M{r}/R{r},0)", NUM)
        put(c, f"T{r}", f"=IFERROR(M{r}/N{r},0)", PCT)
        for j in range(4):
            a, b, v = (col_letter(21 + j), col_letter(25 + j), col_letter(29 + j))
            put(c, f"{a}{r}", f"=SUMIFS(Sales[Revenue],Sales[Month],$K{r},Sales[In],1,Sales[Region],{a}$2)", NUM)
            put(c, f"{b}{r}", f"=SUMIFS(Budget[Budget Revenue],Budget[Month],$K{r},Budget[In],1,Budget[Region],{b}$2)", NUM)
            put(c, f"{v}{r}", f'=IF({b}{r}=0,"",{a}{r}/{b}{r}-1)', PCTS)
        for j in range(5):
            a = col_letter(33 + j)
            put(c, f"{a}{r}", f"=SUMIFS(Sales[Revenue],Sales[Month],$K{r},Sales[In],1,Sales[Category],{a}$2)", NUM)

    # ---- forecast  K18:T31 + summary K33:L49
    c.Range("K18").Value = "Forecast year: actual to the reporting month, then budget x recent achievement (last 3 months actual / budget)"
    c.Range("K19:T19").Value = ("Month", "Label", "Actual", "Budget", "Prior year", "Forecast", "Actual or forecast", "Cumulative",
                                "Cumulative budget", "vs budget")
    for n in range(12):
        r = 20 + n
        put(c, f"K{r}", f"=DATE(FYear,{n + 1},1)", "mmm yyyy")
        put(c, f"L{r}", f'=TEXT(K{r},"mmm")')
        put(c, f"M{r}", f"=IF(K{r}<=RepMonth,SUMIFS(Sales[Revenue],Sales[Month],K{r},Sales[In],1),0)", NUM)
        put(c, f"N{r}", f"=SUMIFS(Budget[Budget Revenue],Budget[Month],K{r},Budget[In],1)", NUM)
        put(c, f"O{r}", f"=SUMIFS(Sales[Revenue],Sales[Month],EDATE(K{r},-12),Sales[In],1)", NUM)
        put(c, f"P{r}", f"=IF(K{r}>RepMonth,N{r}*RecentAch,0)", NUM)
        put(c, f"Q{r}", f"=M{r}+P{r}", NUM)
        put(c, f"R{r}", f"=SUM(Q$20:Q{r})", NUM)
        put(c, f"S{r}", f"=SUM(N$20:N{r})", NUM)
        put(c, f"T{r}", f'=IF(N{r}=0,"",Q{r}/N{r}-1)', PCTS)
    fc = [("RecentAch", "Recent achievement (last 3 months actual / budget)",
           '=IFERROR(SUMIFS(Sales[Revenue],Sales[Month],">="&EDATE(RepMonth,-2),Sales[Month],"<="&RepMonth,Sales[In],1)/'
           'SUMIFS(Budget[Budget Revenue],Budget[Month],">="&EDATE(RepMonth,-2),Budget[Month],"<="&RepMonth,Budget[In],1),0)', PCT),
          ("RecentAvg", "Average month, last 3 months",
           '=SUMIFS(Sales[Revenue],Sales[Month],">="&EDATE(RepMonth,-2),Sales[Month],"<="&RepMonth,Sales[In],1)/3', NUM),
          ("FcYTD", "Year to date actual", "=SUM(M20:M31)", NUM),
          ("FcFYB", "Full-year budget", "=SUM(N20:N31)", NUM),
          ("FcFY", "Year-end forecast (budget-phased run-rate)", "=SUM(Q20:Q31)", NUM),
          ("FcVsB", "Forecast vs full-year budget", '=IF(FcFYB=0,"",FcFY/FcFYB-1)', PCTS),
          ("FcLeft", "Months left", "=12-MONTH(RepMonth)", "0"),
          ("FcStraight", "Year-end, straight run-rate (YTD + 3-month average x months left)", "=FcYTD+RecentAvg*FcLeft", NUM),
          ("FcStrVsB", "Straight run-rate vs budget", '=IF(FcFYB=0,"",FcStraight/FcFYB-1)', PCTS),
          ("FcPY", "Prior year, full year", "=SUM(O20:O31)", NUM),
          ("FcVsPY", "Forecast vs prior year", '=IF(FcPY=0,"",FcFY/FcPY-1)', PCTS),
          ("FcYTDB", "Year to date budget", "=SUMPRODUCT((K20:K31<=RepMonth)*N20:N31)", NUM),
          ("FcYTDvsB", "YTD vs budget", '=IF(FcYTDB=0,"",FcYTD/FcYTDB-1)', PCTS),
          ("FcYTDPY", "Prior-year YTD", "=SUMPRODUCT((K20:K31<=RepMonth)*O20:O31)", NUM),
          ("FcYTDvsPY", "YTD vs prior year", '=IF(FcYTDPY=0,"",FcYTD/FcYTDPY-1)', PCTS),
          ("FcGap", "Forecast minus budget", "=FcFY-FcFYB", NUMS),
          ("FcRest", "Forecast for the months left", "=SUM(P20:P31)", NUM)]
    c.Range("V19:W19").Value = ("Forecast summary", "Value")
    for i, (nm, label, f, fmt) in enumerate(fc, 20):
        c.Range(f"V{i}").Value = label
        put(c, f"W{i}", f, fmt)
        name(wb, nm, f"W{i}")

    # ---- region table  A23:I29
    def block(top, key, items, sales_col, bud_col, extra=False):
        heads = [key, "Actual", "Budget", "Variance", "Var %", "Prior year", "vs PY", "Gross profit", "Margin %"]
        if extra:
            heads += ["Budget margin %", "Share"]
        c.Range(c.Cells(top, 1), c.Cells(top, len(heads))).Value = tuple(heads)
        tot = top + len(items) + 1
        for i, it in enumerate(items + ["Total"], top + 1):
            c.Range(f"A{i}").Value = it
            if it == "Total":
                for col in "BCFH":
                    put(c, f"{col}{i}", f"=SUM({col}{top + 1}:{col}{i - 1})", NUM)
            else:
                cs, cb = f",Sales[{sales_col}],$A{i}", f",Budget[{bud_col}],$A{i}"
                put(c, f"B{i}", "=" + SR("Revenue", "PStart", "PEnd", cs), NUM)
                put(c, f"C{i}", "=" + SB("Budget Revenue", "PStart", "PEnd", cb), NUM)
                put(c, f"F{i}", "=" + SR("Revenue", "PYStart", "PYEnd", cs), NUM)
                put(c, f"H{i}", "=" + SR("Margin", "PStart", "PEnd", cs), NUM)
            put(c, f"D{i}", f"=B{i}-C{i}", NUMS)
            put(c, f"E{i}", f'=IF(C{i}=0,"",B{i}/C{i}-1)', PCTS)
            put(c, f"G{i}", f'=IF(F{i}=0,"",B{i}/F{i}-1)', PCTS)
            put(c, f"I{i}", f'=IF(B{i}=0,"",H{i}/B{i})', PCT)
            if extra:
                if it == "Total":
                    put(c, f"J{i}", '=IF(C{0}=0,"",SUMPRODUCT(C{1}:C{2},J{1}:J{2})/C{0})'.format(i, top + 1, i - 1), PCT)
                else:
                    put(c, f"J{i}", f'=IF(C{i}=0,"",{SB("Budget GP", "PStart", "PEnd", cb)}/C{i})', PCT)
                put(c, f"K{i}", f'=IF($B${tot}=0,"",B{i}/$B${tot})', PCT)
        return tot

    c.Range("A22").Value = "By region, selected period"
    block(23, "Region", REGIONS, "Region", "Region")                  # rows 24-27, total 28
    c.Range("A31").Value = "By category, selected period"
    block(32, "Category", CATS, "Category", "Category", extra=True)   # rows 33-37, total 38

    # ---- waterfalls (stacked columns: invisible base + total + up + down; bars start at a floor so gaps show)
    def waterfall(top, labels, deltas_rng, budget_cell, actual_cell, nm):
        c.Range(c.Cells(top, 1), c.Cells(top, 8)).Value = ("Step", "Value", "Cumulative", "Base", "Total", "Up", "Down", "Axis label")
        n = len(labels)
        first, last = top + 1, top + n + 2
        rows_ = ["Budget"] + labels + ["Actual"]
        for k, lab in enumerate(rows_):
            r = first + k
            c.Range(f"A{r}").Value = lab
            if k == 0:
                put(c, f"B{r}", f"={budget_cell}", NUM); put(c, f"C{r}", f"=B{r}", NUM)
            elif k == n + 1:
                put(c, f"B{r}", f"={actual_cell}", NUM); put(c, f"C{r}", f"=B{r}", NUM)
            else:
                put(c, f"B{r}", f"={deltas_rng[k - 1]}", NUMS); put(c, f"C{r}", f"=C{r - 1}+B{r}", NUM)
        fl = last + 1
        c.Range(f"A{fl}").Value = "Floor"
        put(c, f"B{fl}", f"=MAX(0,ROUNDDOWN(MIN(C{first}:C{last})*0.95,2-LEN(INT(MIN(C{first}:C{last})*0.95))))", NUM)
        name(wb, nm + "Floor", f"B{fl}")
        c.Range(f"A{fl + 1}").Value = "Note"
        put(c, f"B{fl + 1}", f'=PeriodLabel&"  |  bars start at "&TEXT(B{fl},"#,##0")&" to show the steps"')
        name(wb, nm + "Note", f"B{fl + 1}")
        for k in range(n + 2):
            r = first + k
            total = k in (0, n + 1)
            put(c, f"D{r}", "=0" if total else f"=MIN(C{r - 1},C{r})-$B${fl}", NUM)
            put(c, f"E{r}", f"=C{r}-$B${fl}" if total else "=0", NUM)
            put(c, f"F{r}", "=0" if total else f"=MAX(B{r},0)", NUM)
            put(c, f"G{r}", "=0" if total else f"=MAX(-B{r},0)", NUM)
            put(c, f"H{r}", f'=A{r}&CHAR(10)&TEXT(B{r}/1000,"#,##0.0")&"k"' if total else
                f'=A{r}&CHAR(10)&TEXT(B{r}/1000,"+#,##0.0;-#,##0.0;0.0")&"k"')
        return first, last

    c.Range("A40").Value = "Waterfall: budget to actual by region"
    waterfall(41, REGIONS, [f"D{24 + i}" for i in range(4)], "C28", "B28", "WfReg")        # rows 42-47, floor 48
    c.Range("A51").Value = "Waterfall: budget to actual by category"
    waterfall(52, CATS, [f"D{33 + i}" for i in range(5)], "C38", "B38", "WfCat")           # rows 53-59, floor 60

    # ---- region x category matrix  A63:R67  (for the "driven by" insight)
    c.Range("A62").Value = "Region x category, selected period: actual | budget | variance"
    for j, cat in enumerate(CATS):
        for base in (2, 8, 14):
            c.Cells(63, base + j).Value = cat
    for i, reg in enumerate(REGIONS, 64):
        c.Range(f"A{i}").Value = reg
        for j in range(5):
            a, b, v = col_letter(2 + j), col_letter(8 + j), col_letter(14 + j)
            put(c, f"{a}{i}", "=" + SR("Revenue", "PStart", "PEnd", f",Sales[Region],$A{i},Sales[Category],{a}$63"), NUM)
            put(c, f"{b}{i}", "=" + SB("Budget Revenue", "PStart", "PEnd", f",Budget[Region],$A{i},Budget[Category],{b}$63"), NUM)
            put(c, f"{v}{i}", f"={a}{i}-{b}{i}", NUMS)
    w = [("WorstIdx", "Weakest region (lowest var %, regions with a budget)", '=IFERROR(MATCH(AGGREGATE(15,6,E24:E27,1),E24:E27,0),0)', "0"),
         ("WorstReg", "Weakest region", '=IF(WorstIdx=0,"",INDEX(A24:A27,WorstIdx))', None),
         ("WorstPct", "Its var %", '=IF(WorstIdx=0,"",INDEX(E24:E27,WorstIdx))', PCTS),
         ("WorstAmt", "Its variance", '=IF(WorstIdx=0,"",INDEX(D24:D27,WorstIdx))', NUMS)]
    for i, (nm, label, f, fmt) in enumerate(w, 70):
        c.Range(f"A{i}").Value = label; put(c, f"B{i}", f, fmt); name(wb, nm, f"B{i}")
    c.Range("A74").Value = "Its categories: variance"; c.Range("A75").Value = "Its categories: budget"
    for j in range(5):
        col = col_letter(2 + j)
        put(c, f"{col}74", f"=IF(WorstIdx=0,0,INDEX({col_letter(14 + j)}64:{col_letter(14 + j)}67,WorstIdx))", NUMS)
        put(c, f"{col}75", f"=IF(WorstIdx=0,0,INDEX({col_letter(8 + j)}64:{col_letter(8 + j)}67,WorstIdx))", NUM)
    w2 = [("WorstCatIdx", "Category with the largest shortfall there", "=IFERROR(MATCH(MIN(B74:F74),B74:F74,0),0)", "0"),
          ("WorstCat", "Category", '=IF(WorstCatIdx=0,"",INDEX(B63:F63,WorstCatIdx))', None),
          ("WorstCatPct", "Its var %", '=IF(WorstCatIdx=0,"",IF(INDEX(B75:F75,WorstCatIdx)=0,"",INDEX(B74:F74,WorstCatIdx)/INDEX(B75:F75,WorstCatIdx)))', PCTS),
          ("WorstCatAmt", "Its variance", '=IF(WorstCatIdx=0,"",INDEX(B74:F74,WorstCatIdx))', NUMS)]
    for i, (nm, label, f, fmt) in enumerate(w2, 76):
        c.Range(f"A{i}").Value = label; put(c, f"B{i}", f, fmt); name(wb, nm, f"B{i}")

    # ---- customers (query table AN3 + formula columns) and ranked list AV3:BD63
    cust = c.ListObjects("Customers")
    for nm, f in (("Revenue", "=" + SR("Revenue", "PStart", "PEnd", ",Sales[Customer],[@Customer]")),
                  ("PY Revenue", "=" + SR("Revenue", "PYStart", "PYEnd", ",Sales[Customer],[@Customer]")),
                  ("Orders", "=" + SO("PStart", "PEnd", ",Sales[Customer],[@Customer]")),
                  ("RankKey", "=IF([@Revenue]>0,[@Revenue]+ROW()/1000000000,0)")):
        col = cust.ListColumns.Add(); col.Name = nm; col.DataBodyRange.Formula = f
    for nm in ("Revenue", "PY Revenue", "Orders"):
        cust.ListColumns(nm).DataBodyRange.NumberFormat = NUM
    c.Range("AV2").Value = "Customers ranked by revenue, selected period"
    c.Range("AV3:BE3").Value = ("Rank", "Customer", "Region", "Revenue", "Prior year", "vs PY", "Orders", "Share", "Cumulative share", "Row")
    for n in range(1, 61):
        r = 3 + n
        c.Range(f"AV{r}").Value = n
        put(c, f"BE{r}", f"=IF(AV{r}>CustActive,0,MATCH(LARGE(Customers[RankKey],AV{r}),Customers[RankKey],0))", "0")
        put(c, f"AW{r}", f'=IF(BE{r}=0,"",INDEX(Customers[Customer],BE{r}))')
        put(c, f"AX{r}", f'=IF(BE{r}=0,"",INDEX(Customers[Region],BE{r}))')
        put(c, f"AY{r}", f'=IF(BE{r}=0,"",INDEX(Customers[Revenue],BE{r}))', NUM)
        put(c, f"AZ{r}", f'=IF(BE{r}=0,"",INDEX(Customers[PY Revenue],BE{r}))', NUM)
        put(c, f"BA{r}", f'=IF(BE{r}=0,"",IF(AZ{r}=0,"",AY{r}/AZ{r}-1))', PCTS)
        put(c, f"BB{r}", f'=IF(BE{r}=0,"",INDEX(Customers[Orders],BE{r}))', NUM)
        put(c, f"BC{r}", f'=IF(BE{r}=0,"",AY{r}/CustTotal)', PCT)
        put(c, f"BD{r}", f'=IF(BE{r}=0,"",SUM(AY$4:AY{r})/CustTotal)', PCT)
    cs = [("CustTotal", "Customer revenue total", "=SUM(Customers[Revenue])", NUM),
          ("CustActive", "Active customers (revenue > 0)", '=COUNTIF(Customers[Revenue],">0")', "0"),
          ("CustOnFile", "Customers on file", "=ROWS(Customers[Customer])", "0"),
          ("CustN80", "Customers that bring 80% of revenue", '=IF(CustActive=0,0,COUNTIF(BD4:BD63,"<0.8")+1)', "0"),
          ("CustN80Pct", "... as a share of active customers", "=IFERROR(CustN80/CustActive,0)", PCT),
          ("CustTop10", "Top 10 share", "=IFERROR(SUM(AY4:AY13)/CustTotal,0)", PCT),
          ("CustTopName", "Largest customer", "=AW4", None),
          ("CustTopShare", "Largest customer share", "=IFERROR(BC4*1,0)", PCT),
          ("CustAvg", "Average revenue per active customer", "=IFERROR(CustTotal/CustActive,0)", NUM),
          ("CustActivePY", "Active customers, prior year", '=COUNTIF(Customers[PY Revenue],">0")', "0"),
          ("CustAvgPY", "Average per customer, prior year", "=IFERROR(SUM(Customers[PY Revenue])/CustActivePY,0)", NUM),
          ("CustAvgVsPY", "Average per customer vs PY", '=IF(CustAvgPY=0,"",CustAvg/CustAvgPY-1)', PCTS),
          ("CustActiveVsPY", "Active customers vs PY", "=CustActive-CustActivePY", "+0;-0;0"),
          ("CustTop10VsPY", "Top 10 share, prior year", '=IFERROR(SUMPRODUCT(LARGE(Customers[PY Revenue],{1,2,3,4,5,6,7,8,9,10}))/SUM(Customers[PY Revenue]),"")', PCT)]
    for i, (nm, label, f, fmt) in enumerate(cs, 66):
        c.Range(f"AV{i}").Value = label; put(c, f"AW{i}", f, fmt); name(wb, nm, f"AW{i}")
    # Pareto chart helper: 80% line
    for n in range(20):
        put(c, f"BF{4 + n}", "=0.8", PCT)
    c.Range("BF3").Value = "80% line"

    # ---- reps (query table BH3 + formula columns) and scorecard BH14:BV21
    reps = c.ListObjects("Reps")
    rr = ",Sales[Sales Rep],[@[Sales Rep]]"
    for nm, f in (("Revenue", "=" + SR("Revenue", "PStart", "PEnd", rr)),
                  ("Target", "=" + SB("Budget Revenue", "PStart", "PEnd", ",Budget[Sales Rep],[@[Sales Rep]]")),
                  ("% of target", '=IF([@Target]=0,"",[@Revenue]/[@Target])'),
                  ("PY Revenue", "=" + SR("Revenue", "PYStart", "PYEnd", rr)),
                  ("vs PY", '=IF([@[PY Revenue]]=0,"",[@Revenue]/[@[PY Revenue]]-1)'),
                  ("Orders", "=" + SO("PStart", "PEnd", rr)),
                  ("Avg order", "=IFERROR([@Revenue]/[@Orders],0)"),
                  ("Gross profit", "=" + SR("Margin", "PStart", "PEnd", rr)),
                  ("Margin %", '=IF([@Revenue]=0,"",[@[Gross profit]]/[@Revenue])'),
                  ("Share", '=IF(SUM([Revenue])=0,"",[@Revenue]/SUM([Revenue]))'),
                  ("RankKey", "=IF([@Target]>0,[@[% of target]]+ROW()/1000000000,-1)")):
        col = reps.ListColumns.Add(); col.Name = nm; col.DataBodyRange.Formula = f
    c.Range("BH13").Value = "Scorecard: reps with a target in the selection, best % of target first"
    sc_cols = ["Rank", "Sales rep", "Region", "Revenue", "Target", "% of target", "Prior year", "vs PY", "Orders", "Avg order",
               "Margin %", "Share", "Row", "Gauge to 100%", "Gauge over 100%", "Gauge rest", "Gauge hidden", "Gauge label", "Revenue / target"]
    c.Range(c.Cells(14, 60), c.Cells(14, 60 + len(sc_cols) - 1)).Value = tuple(sc_cols)   # BH14...
    src = {"Sales rep": "Sales Rep", "Region": "Region", "Revenue": "Revenue", "Target": "Target", "% of target": "% of target",
           "Prior year": "PY Revenue", "vs PY": "vs PY", "Orders": "Orders", "Avg order": "Avg order", "Margin %": "Margin %", "Share": "Share"}
    fmts = {"Revenue": NUM, "Target": NUM, "% of target": PCT, "Prior year": NUM, "vs PY": PCTS, "Orders": NUM, "Avg order": NUM,
            "Margin %": PCT, "Share": PCT}
    for n in range(1, 9):
        r = 14 + n
        cols = {h: col_letter(60 + k) for k, h in enumerate(sc_cols)}
        c.Range(f"{cols['Rank']}{r}").Value = n
        rowc = f"{cols['Row']}{r}"
        put(c, rowc, f"=IF({cols['Rank']}{r}>RepsWithTarget,0,MATCH(LARGE(Reps[RankKey],{cols['Rank']}{r}),Reps[RankKey],0))", "0")
        for h, s in src.items():
            put(c, f"{cols[h]}{r}", f'=IF({rowc}=0,"",INDEX(Reps[{s}],{rowc}))', fmts.get(h))
        p = f"{cols['% of target']}{r}"
        put(c, f"{cols['Gauge to 100%']}{r}", f"=IF({rowc}=0,0,MIN({p},1))", "0.000")
        put(c, f"{cols['Gauge over 100%']}{r}", f"=IF({rowc}=0,0,MIN(MAX({p}-1,0),0.5))", "0.000")
        put(c, f"{cols['Gauge rest']}{r}", f"=1.5-{cols['Gauge to 100%']}{r}-{cols['Gauge over 100%']}{r}", "0.000")
        put(c, f"{cols['Gauge hidden']}{r}", "=1.5", "0.0")
        put(c, f"{cols['Gauge label']}{r}", f'=IF({rowc}=0,"-",TEXT({p},"0%"))')
        put(c, f"{cols['Revenue / target']}{r}", f'=IF({rowc}=0,"no target in this selection",TEXT({cols["Revenue"]}{r},"#,##0")&" of "&TEXT({cols["Target"]}{r},"#,##0"))')
    rs = [("RepsWithTarget", "Reps with a target in the selection", '=COUNTIF(Reps[Target],">0")', "0"),
          ("RepsOnTarget", "... on or above target", '=COUNTIFS(Reps[Target],">0",Reps[% of target],">=1")', "0"),
          ("RepLow", "Lowest vs target", '=IF(RepsWithTarget=0,"",INDEX(BI15:BI22,RepsWithTarget))', None),
          ("RepLowPct", "... % of target", '=IF(RepsWithTarget=0,"",INDEX(BM15:BM22,RepsWithTarget))', PCT),
          ("RepTeamPct", "Team % of target", '=IFERROR(SUM(Reps[Revenue])/SUM(Reps[Target]),0)', PCT)]
    for i, (nm, label, f, fmt) in enumerate(rs, 24):
        c.Range(f"BH{i}").Value = label; put(c, f"BI{i}", f, fmt); name(wb, nm, f"BI{i}")

    # ---- card helpers: KPI notes and key insights  A82:C92
    ins = [
        ('=PeriodLabel&": revenue "&TEXT(E4,"#,##0")&IF(E7=0,", no budget",", "&TEXT(E15,"+0.0%;-0.0%")&" vs budget ("&TEXT(E4-E7,"+#,##0;-#,##0")&")")'
         '&IF(F4=0,", no prior-year data.",", "&TEXT(F15,"+0.0%;-0.0%")&" vs "&PYLabel&".")', '=IF(E7=0,0,IF(E4<E7,-1,1))'),
        ('=IF(WorstIdx=0,"No budget in this selection.","Weakest region: "&WorstReg&" "&TEXT(WorstPct,"+0.0%;-0.0%")&" vs budget, driven by "'
         '&WorstCat&IF(WorstCatPct=""," "," "&TEXT(WorstCatPct,"+0.0%;-0.0%")&" ")&"("&TEXT(WorstCatAmt,"+#,##0;-#,##0")&").")',
         '=IF(WorstIdx=0,0,IF(WorstPct<0,-1,1))'),
        ('=IF(E4=0,"No sales in this selection.","Gross margin "&TEXT(E9,"0.0%")&IF(E7=0,"",": "&TEXT(E17*100,"+0.0;-0.0")&" pts vs budget")'
         '&IF(F4=0,"",", "&TEXT(F17*100,"+0.0;-0.0")&" pts vs prior year")&".")', '=IF(OR(E4=0,E7=0),0,IF(E9<E10,-1,1))'),
        ('=IF(CustActive=0,"No sales in this selection.",CustN80&" of "&CustActive&" active customers bring 80% of revenue; the largest, "'
         '&CustTopName&", brings "&TEXT(CustTopShare,"0.0%")&".")', "=0"),
        ('=IF(RepsWithTarget=0,"No sales targets in this selection.","Lowest vs target: "&RepLow&" at "&TEXT(RepLowPct,"0%")&"; "'
         '&RepsOnTarget&" of "&RepsWithTarget&" reps on or above target.")', '=IF(RepsWithTarget=0,0,IF(RepLowPct<1,-1,1))'),
        ('=IF(FcFYB=0,"No full-year budget for "&FYear&".","At the recent run-rate "&FYear&" ends at "&TEXT(FcFY,"#,##0")&", "'
         '&TEXT(FcVsB,"+0.0%;-0.0%")&" vs the full-year budget of "&TEXT(FcFYB,"#,##0")&".")', '=IF(FcFYB=0,0,IF(FcFY<FcFYB,-1,1))')]
    c.Range("A82:C82").Value = ("Key insight", "Sentence (written by formulas)", "Tone (1 good, -1 bad)")
    for i, (txt, tone) in enumerate(ins, 83):
        c.Range(f"A{i}").Value = f"Insight {i - 82}"
        put(c, f"B{i}", txt); put(c, f"C{i}", tone, "0")
    wb.Names.Add("Insights", "=Calc!$B$83:$B$88")

    for rng in ("A3:B3", "D3:H3", "D14:G14", "K3:AK3", "K19:T19", "V19:W19", "A23:I23", "A32:K32", "A41:H41", "A52:H52",
                "A82:C82", "AV3:BF3", "BH14:CB14"):
        light_header(c.Range(rng))
    c.Columns("A").ColumnWidth = 44; c.Columns("B").ColumnWidth = 40
    for col in ["C", "D"] + [col_letter(i) for i in range(5, 40)]:
        c.Columns(col).ColumnWidth = 11
    c.Columns("D").ColumnWidth = 22; c.Columns("K").ColumnWidth = 44
    c.Columns("AN").ColumnWidth = 26; c.Columns("AV").ColumnWidth = 30; c.Columns("AW").ColumnWidth = 26
    c.Columns("BH").ColumnWidth = 30; c.Columns("BI").ColumnWidth = 14; c.Columns("CA").ColumnWidth = 52
    c.ListObjects("Checks").ListColumns("Value").DataBodyRange.NumberFormat = "#,##0.##"
    for nm in ("Revenue", "Target", "PY Revenue", "Orders", "Avg order", "Gross profit"):
        reps.ListColumns(nm).DataBodyRange.NumberFormat = NUM
    for nm in ("% of target", "Margin %", "Share"):
        reps.ListColumns(nm).DataBodyRange.NumberFormat = PCT
    reps.ListColumns("vs PY").DataBodyRange.NumberFormat = PCTS


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ------------------------------------------------------------------------------------------------- styles
def make_styles(wb):
    st = wb.TableStyles.Add("Gold Slicer")
    st.ShowAsAvailableSlicerStyle = True; st.ShowAsAvailableTableStyle = False; st.ShowAsAvailablePivotTableStyle = False
    whole = st.TableStyleElements(0); whole.Interior.Color = P["card"]; whole.Font.Color = P["ivory"]
    head = st.TableStyleElements(1); head.Font.Color = P["gold"]; head.Font.Bold = True
    for idx, (bg, fg) in {28: ("202832", "ECE6DA"), 29: ("161B21", "56606A"), 30: ("C9A35B", "0B0F13"), 31: ("6E5A34", "0B0F13"),
                          32: ("2A333D", "ECE6DA"), 33: ("E6CF9A", "0B0F13"), 34: ("2A333D", "8E98A3"), 35: ("E6CF9A", "0B0F13")}.items():
        e = st.TableStyleElements(idx); e.Interior.Color = hexrgb(bg); e.Font.Color = hexrgb(fg)
    ts = wb.TableStyles.Add("Gold Timeline")
    ts.ShowAsAvailableTimelineStyle = True; ts.ShowAsAvailableTableStyle = False; ts.ShowAsAvailablePivotTableStyle = False
    ts.TableStyleElements(0).Interior.Color = P["card"]; ts.TableStyleElements(0).Font.Color = P["ivory"]
    ts.TableStyleElements(1).Font.Color = P["gold"]; ts.TableStyleElements(1).Font.Bold = True
    for idx, (bg, fg) in {36: (None, "E6CF9A"), 37: (None, "8E98A3"), 38: (None, "8E98A3"), 39: (None, "ECE6DA"),
                          40: ("C9A35B", None), 41: ("2A333D", None), 42: ("1A2129", None)}.items():
        e = ts.TableStyleElements(idx)
        if bg:
            e.Interior.Color = hexrgb(bg)
        if fg:
            e.Font.Color = hexrgb(fg)


# ------------------------------------------------------------------------------------------------- page kit
NCOL, C0 = 65, 11          # 65 grid columns; content starts at column K (11); sidebar B:I


def cc(i):
    return C0 + i


def rng(ws, r1, c1, r2=None, c2=None):
    return ws.Range(ws.Cells(r1, c1), ws.Cells(r2 or r1, c2 or c1))


def fill(r, color):
    r.Interior.Color = color


def box(ws, r1, c1, r2, c2, color=None, line=None):
    r = rng(ws, r1, c1, r2, c2)
    fill(r, P["card"] if color is None else color)
    for e in (7, 8, 9, 10):
        r.Borders(e).LineStyle = 1; r.Borders(e).Weight = 2; r.Borders(e).Color = line or P["line"]
    return r


def text(ws, r, c1, c2, val, size=9, color=None, bold=False, align=-4131, font=SANS, fmt=None, valign=-4108, wrap=False, italic=False):
    r_ = rng(ws, r, c1, r, c2)
    if c2 > c1:
        r_.Merge()
    cell = ws.Cells(r, c1)
    if isinstance(val, str) and val.startswith("="):
        cell.Formula2 = val
    elif val is not None:
        cell.Value = val
    f = r_.Font
    f.Name = font; f.Size = size; f.Bold = bold; f.Italic = italic; f.Color = P["ivory"] if color is None else color
    r_.HorizontalAlignment = align; r_.VerticalAlignment = valign
    r_.WrapText = wrap; r_.IndentLevel = 0
    if fmt:
        cell.NumberFormat = fmt
    return cell


SEP = ","


def probe_separator(wb):
    """Conditional-format formulas are parsed with the PC's list separator (',' or ';'): find out which by trying."""
    global SEP
    cell = wb.Worksheets("Settings").Range("H1")
    try:
        cell.FormatConditions.Add(2, None, "=AND(1,1)")
        SEP = ","
    except Exception:
        SEP = ";"
    cell.FormatConditions.Delete()
    print("list separator", repr(SEP))


def cfx(f):
    """Conditional-format formulas are parsed with the PC's list separator; this build replaces the commas."""
    return f.replace(",", SEP) if SEP != "," else f


def cf_sign(r, ref, good=P["good"], bad=P["bad"], fills=False):
    """Font (and optional fill) green when ref >= 0, red when < 0; ref is a formula on the top-left cell (relative)."""
    fc = r.FormatConditions.Add(2, None, cfx(f"=AND(ISNUMBER({ref}),{ref}<0)"))
    fc.Font.Color = bad
    if fills:
        fc.Interior.Color = P["badbg"]
    fc = r.FormatConditions.Add(2, None, cfx(f"=AND(ISNUMBER({ref}),{ref}>0)"))
    fc.Font.Color = good
    if fills:
        fc.Interior.Color = P["goodbg"]


def page_rows(ws, heights):
    for r, h in heights.items():
        ws.Rows(r).RowHeight = h


def frame(wb, sheet, title, caches, slicers=True):
    ws = wb.Worksheets(sheet)
    ws.Cells.Font.Name = SANS; ws.Cells.Font.Size = 9; ws.Cells.Font.Color = P["ivory"]
    ws.Columns(1).ColumnWidth = 1.3
    for i in range(2, NCOL + 1):
        ws.Columns(i).ColumnWidth = 2.29
    ws.Columns(NCOL).ColumnWidth = 1.3
    ws.Columns(10).ColumnWidth = 1.6
    ws.StandardWidth = 2.29
    rng(ws, 1, 1, 140, NCOL).Interior.Color = P["bg"]
    rng(ws, 1, 1, 4, NCOL).Interior.Color = P["bg0"]
    page_rows(ws, {1: 6, 2: 26, 3: 16, 4: 6, 5: 10})
    b = rng(ws, 4, 1, 4, NCOL).Borders(9); b.LineStyle = 1; b.Weight = 2; b.Color = P["gold"]
    text(ws, 2, 2, 9, "B R I G H T L I N E", 8, P["gold"], True, valign=-4107)
    text(ws, 3, 2, 9, "sales performance", 8, P["muted"], valign=-4160, italic=True)
    text(ws, 2, cc(0), cc(22), title, 19, P["ivory"], font=SERIF, valign=-4107)
    text(ws, 3, cc(0), cc(30), '=Company&" (fictional)  |  "&SelText&"  |  USD"', 8, P["muted"], valign=-4160)
    # navigation buttons
    left0 = ws.Cells(1, cc(24)).Left
    right = ws.Cells(1, cc(54)).Left
    n = len(PAGES); gap = 6
    w = (right - left0 - gap * (n - 1)) / n
    top, h = ws.Cells(2, 1).Top + 3, 26
    for k, (s, nav, _) in enumerate(PAGES):
        x = left0 + k * (w + gap)
        shp = ws.Shapes.AddShape(5, x, top, w, h)
        shp.Name = f"nav_{s}"
        try:
            shp.Adjustments.SetItem(1, 0.12)
        except Exception:
            pass
        cur = s == sheet
        shp.Fill.ForeColor.RGB = P["navsel"] if cur else P["card"]
        shp.Line.ForeColor.RGB = P["gold"] if cur else P["line"]; shp.Line.Weight = 1
        tr = shp.TextFrame2.TextRange
        tr.Text = nav
        tr.Font.Size = 9.5; tr.Font.Name = SANS; tr.Font.Bold = cur
        tr.Font.Fill.ForeColor.RGB = P["gold2"] if cur else P["ivory"]
        tr.ParagraphFormat.Alignment = 2
        shp.TextFrame2.VerticalAnchor = 3
        shp.TextFrame2.MarginLeft = 2; shp.TextFrame2.MarginRight = 2
        ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress=f"'{s}'!A1", ScreenTip=f"Go to {nav}")
    if slicers:
        add_slicers(wb, ws, caches)
    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
        setattr(ps, m, 10)
    ps.CenterHorizontally = True
    return ws


SIDEBAR = [("View", "VIEW", 2, 50), ("Region", "REGION", 2, 76), ("Category", "CATEGORY", 1, 136), ("Rep", "SALES REP", 1, 154)]


def add_slicers(wb, ws, caches):
    left = ws.Cells(1, 2).Left
    width = ws.Cells(1, 10).Left - left
    top = ws.Cells(6, 1).Top
    for key, cap, ncol, h in SIDEBAR:
        s = caches[key].Slicers.Add(SlicerDestination=ws, Name=f"{key}_{ws.Name}", Caption=cap, Top=top, Left=left, Width=width, Height=h)
        top += h + 8
    tb = ws.Shapes.AddTextbox(1, left, top, width, 60)
    tb.Name = "slicer_tip"
    tb.TextFrame2.TextRange.Text = "Click to filter, Ctrl+click for several, the funnel-x clears. Every number on every page follows."
    tb.TextFrame2.TextRange.Font.Size = 8; tb.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = P["muted"]
    tb.TextFrame2.TextRange.Font.Name = SANS
    tb.Fill.Visible = False; tb.Line.Visible = False
    tb.TextFrame2.MarginLeft = 2; tb.TextFrame2.MarginRight = 2
    # timeline across the top of the content area
    tl_left = ws.Cells(1, cc(0)).Left
    tl = caches["Month"].Slicers.Add(SlicerDestination=ws, Name=f"Month_{ws.Name}", Caption="MONTH  (the last month picked is the reporting month)",
                                     Top=ws.Cells(6, 1).Top, Left=tl_left, Width=ws.Cells(1, cc(38)).Left - tl_left, Height=74)
    # selection box right of the timeline
    box(ws, 6, cc(39), 9, cc(53))
    text(ws, 6, cc(39) + 1, cc(53) - 1, "SHOWING", 8, P["gold"], True)
    text(ws, 7, cc(39) + 1, cc(53) - 1, "=PeriodLabel&IF(ViewMode=\"YTD\",\"  (year to date)\",\"  (month)\")", 11, P["ivory"], font=SERIF)
    text(ws, 8, cc(39) + 1, cc(53) - 1, '=IF(NReg=4,"All regions",NReg&" of 4 regions")&"  |  "&IF(NCat=5,"all categories",NCat&" of 5 categories")'
         '&"  |  "&IF(NRep=ROWS(Reps[Sales Rep]),"all reps",IF(NRep=1,"1 rep",NRep&" reps"))', 8, P["muted"])
    t = text(ws, 9, cc(39) + 1, cc(53) - 1, "Green = better than budget or prior, red = worse", 8, P["muted"])
    page_rows(ws, {6: 18, 7: 20, 8: 16, 9: 20})


def panel(ws, r1, c1, r2, c2, title, sub=None):
    box(ws, r1, c1, r2, c2)
    text(ws, r1, c1 + 1, c2 - 1, title, 12, P["ivory"], font=SERIF, valign=-4107)
    if sub is not None:
        text(ws, r1 + 1, c1 + 1, c2 - 1, sub, 8, P["muted"], valign=-4160)


def chart_in(ws, r1, c1, r2, c2, pad=4):
    left, top = ws.Cells(r1, c1).Left + pad, ws.Cells(r1, 1).Top + pad
    width = ws.Cells(1, c2 + 1).Left - pad - left
    height = ws.Cells(r2 + 1, 1).Top - pad - top
    co = ws.ChartObjects().Add(left, top, width, height)
    ch = co.Chart
    ch.ChartArea.Format.Fill.Visible = False
    ch.ChartArea.Format.Line.Visible = False
    ch.ChartArea.Font.Name = SANS; ch.ChartArea.Font.Size = 8.5; ch.ChartArea.Font.Color = P["muted"]
    try:
        ch.PlotArea.Format.Fill.Visible = False
    except Exception:
        pass
    return co, ch


def dark_axes(ch, value_fmt=None, grid=True, secondary=False):
    try:
        ax = ch.Axes(1)
        ax.Format.Line.ForeColor.RGB = P["line"]
        ax.MajorTickMark = -4142
        ax.TickLabels.Font.Color = P["muted"]
    except Exception:
        pass
    try:
        v = ch.Axes(2)
        v.Format.Line.Visible = False
        v.HasMajorGridlines = grid
        if grid:
            v.MajorGridlines.Format.Line.ForeColor.RGB = P["line"]; v.MajorGridlines.Format.Line.Weight = 0.75
        v.TickLabels.Font.Color = P["muted"]
        if value_fmt:
            v.TickLabels.NumberFormat = value_fmt
    except Exception:
        pass


def legend(ch, pos=-4160):
    ch.HasLegend = True
    ch.Legend.Position = pos
    ch.Legend.Font.Color = P["ivory"]; ch.Legend.Font.Size = 8.5


def no_title(ch):
    ch.HasTitle = True
    ch.ChartTitle.Text = " "
    ch.ChartTitle.Font.Size = 1
    ch.HasTitle = False


def add_series(ch, name, values, xvalues=None, ctype=None, color=None, line=False, axis=None):
    s = ch.SeriesCollection().NewSeries()
    s.Name = name
    s.Values = values
    if xvalues is not None:
        s.XValues = xvalues
    if ctype is not None:
        s.ChartType = ctype
    if axis:
        s.AxisGroup = axis
    if color is not None:
        if line:
            s.Format.Line.Visible = True
            s.Format.Line.ForeColor.RGB = color; s.Format.Line.Weight = 2.25
            try:
                s.MarkerStyle = -4142
            except Exception:
                pass
        else:
            s.Format.Fill.ForeColor.RGB = color
            s.Format.Line.Visible = False
    return s


def card(ws, r, c1, c2, label, icon, big, fmt, spark=None, lines=(), big_ref=None):
    """KPI card: rows r..r+6. lines = [(formula_text, sign_ref_or_None)]"""
    box(ws, r, c1, r + 6, c2)
    lb = rng(ws, r, c1, r + 6, c1).Borders(7); lb.LineStyle = 1; lb.Weight = 4; lb.Color = P["gold"]
    text(ws, r, c1 + 1, c2 - 3, label, 8, P["muted"], True, valign=-4107)
    big_cell = text(ws, r + 1, c1 + 1, c2 - 1, big, 20, P["ivory"], font=SERIF, fmt=fmt)
    if big_ref:
        cf_sign(rng(ws, r + 1, c1 + 1, r + 1, c2 - 1), big_ref, good=P["ivory"], bad=P["bad"])
    if spark:
        sr = rng(ws, r + 2, c1 + 1, r + 2, c2 - 1); sr.Merge()
        g = ws.Cells(r + 2, c1 + 1).SparklineGroups.Add(1, spark)
        g.SeriesColor.Color = P["gold"]; g.LineWeight = 1.25
        g.Points.Lastpoint.Visible = True; g.Points.Lastpoint.Color.Color = P["gold2"]
        g.Points.Highpoint.Visible = False; g.Points.Lowpoint.Visible = False
    for k, (f, ref) in enumerate(lines):
        cell = text(ws, r + 3 + k, c1 + 1, c2 - 1, f, 8, P["muted"])
        if ref:
            cf_sign(rng(ws, r + 3 + k, c1 + 1, r + 3 + k, c2 - 1), abs_(ref))
    # icon: small gold ring with a glyph, top-right
    x = ws.Cells(1, c2).Left - 10
    y = ws.Cells(r, 1).Top + 4
    o = ws.Shapes.AddShape(9, x - 4, y, 19, 19)
    o.Fill.Visible = False; o.Line.ForeColor.RGB = P["gold"]; o.Line.Weight = 1
    tr = o.TextFrame2.TextRange; tr.Text = icon
    tr.Font.Size = 8.5; tr.Font.Bold = True; tr.Font.Name = "Segoe UI Symbol"; tr.Font.Fill.ForeColor.RGB = P["gold"]
    tr.ParagraphFormat.Alignment = 2
    o.TextFrame2.VerticalAnchor = 3
    for m in ("MarginLeft", "MarginRight", "MarginTop", "MarginBottom"):
        setattr(o.TextFrame2, m, 0)
    return big_cell


def card_rows(ws, r):
    page_rows(ws, {r: 18, r + 1: 28, r + 2: 20, r + 3: 13, r + 4: 13, r + 5: 13, r + 6: 6})


def delta(ref, label, pts=False):
    """'▲ 3.5% vs budget' from a Calc cell holding a ratio (or a difference in points)."""
    if pts:
        num = f'TEXT(ABS({ref})*100,"0.0")&" pts"'
    else:
        num = f'TEXT(ABS({ref}),"0.0%")'
    return f'=IF({ref}="","",IF({ref}>=0,"▲ ","▼ ")&{num}&" {label}")', ref


def delta_m(ref, pts=False):
    num = f'TEXT(ABS({ref})*100,"0.0")&" pts"' if pts else f'TEXT(ABS({ref}),"0.0%")'
    return f'=IF({ref}="","",IF({ref}>=0,"▲ ","▼ ")&{num}&" "&MonLabel&" vs "&PrevLabel)', ref


def table(ws, r, c1, cols, nrows, src, total=False, sign_cols=(), shade=True):
    """cols = [(header, width_units, fmt, align)]; src(i, k) -> formula for data row i, column k."""
    c = c1
    spans = []
    for h, wu, fmt, al in cols:
        spans.append((c, c + wu - 1, fmt, al)); c += wu
    hr = rng(ws, r, c1, r, c - 1)
    fill(hr, P["bg0"])
    for (a, b, _, al), (h, *_ ) in zip(spans, cols):
        text(ws, r, a, b, h, 8, P["gold"], True, align=al)
    for i in range(nrows):
        rr = r + 1 + i
        last = total and i == nrows - 1
        row = rng(ws, rr, c1, rr, c - 1)
        if shade:
            fill(row, P["card2"] if i % 2 else P["card"])
        if last:
            fill(row, P["bg0"])
            t = row.Borders(8); t.LineStyle = 1; t.Color = P["gold"]; t.Weight = 2
        for k, (a, b, fmt, al) in enumerate(spans):
            cell = text(ws, rr, a, b, src(i, k), 9, P["ivory"], bold=last, align=al, fmt=fmt)
            if k in sign_cols:
                cf_sign(rng(ws, rr, a, rr, b), cell.Address.replace("$", ""), fills=True)
    return c - 1


# ------------------------------------------------------------------------------------------------- Overview
def build_overview(wb):
    ws = wb.Worksheets("Dashboard")
    calc = wb.Worksheets("Calc")
    # key insights banner  rows 11-15
    page_rows(ws, {11: 24, 12: 30, 13: 30, 14: 30, 15: 6, 16: 10})
    box(ws, 11, cc(0), 15, cc(53))
    lb = rng(ws, 11, cc(0), 15, cc(0)).Borders(7); lb.LineStyle = 1; lb.Weight = 4; lb.Color = P["gold"]
    text(ws, 11, cc(1), cc(30), "Key insights, written by formulas", 12, P["gold2"], font=SERIF, valign=-4107)
    text(ws, 11, cc(31), cc(52), "they rewrite themselves for every slicer choice", 8, P["muted"], align=-4152, valign=-4107, italic=True)
    for k in range(6):
        row = 12 + k % 3
        a = cc(1) if k < 3 else cc(27)
        text(ws, row, a, a, "●", 10, P["muted"], align=-4108, valign=-4160)
        fc = rng(ws, row, a, row, a).FormatConditions.Add(2, None, cfx(f"=Calc!$C${83 + k}<0")); fc.Font.Color = P["bad"]
        fc = rng(ws, row, a, row, a).FormatConditions.Add(2, None, cfx(f"=Calc!$C${83 + k}>0")); fc.Font.Color = P["good"]
        text(ws, row, a + 1, a + 25, f"=Calc!B{83 + k}", 9, P["ivory"], wrap=True, valign=-4160)

    # KPI cards  rows 17-23
    r = 17
    card_rows(ws, r)
    cw = 8
    specs = [("REVENUE", "$", "=Calc!E4", "#,##0", "Calc!M4:M15",
              [delta("Calc!E15", "vs budget"), delta("Calc!F15", "vs prior year"), delta_m("Calc!G15")]),
             ("ACHIEVED", "◎", "=Calc!E12", "0.0%", "Calc!T4:T15",
              [('="budget "&TEXT(Calc!E7,"#,##0")', None), delta("Calc!F20", "vs prior year", True), delta_m("Calc!G20", True)]),
             ("GROSS PROFIT", "Σ", "=Calc!E5", "#,##0", "Calc!P4:P15",
              [delta("Calc!E16", "vs budget"), delta("Calc!F16", "vs prior year"), delta_m("Calc!G16")]),
             ("GROSS MARGIN", "%", "=Calc!E9", "0.0%", "Calc!Q4:Q15",
              [delta("Calc!E17", "vs budget", True), delta("Calc!F17", "vs prior year", True), delta_m("Calc!G17", True)]),
             ("ORDERS", "#", "=Calc!E6", "#,##0", "Calc!R4:R15",
              [('=TEXT(CustActive,"0")&" active customers"', None), delta("Calc!F18", "vs prior year"), delta_m("Calc!G18")]),
             ("AVERAGE ORDER", "≈", "=Calc!E11", "#,##0", "Calc!S4:S15",
              [('="revenue / orders"', None), delta("Calc!F19", "vs prior year"), delta_m("Calc!G19")])]
    for k, (label, icon, big, fmt, spark, lines) in enumerate(specs):
        a = cc(k * 9)
        card(ws, r, a, a + cw - 1 if k < 5 else cc(53), label, icon, big, fmt, spark, lines)
    page_rows(ws, {24: 10})

    # charts row  25-41
    for rr in range(25, 42):
        ws.Rows(rr).RowHeight = 15
    ws.Rows(25).RowHeight = 22; ws.Rows(26).RowHeight = 14
    panel(ws, 25, cc(0), 41, cc(23), "Revenue vs budget and prior year", '="12 months to "&TEXT(RepMonth,"mmm yyyy")&"  |  bars = actual, last bar = reporting month"')
    co, ch = chart_in(ws, 27, cc(0), 41, cc(23))
    co.Name = "chart_trend"
    ch.ChartType = 51
    s1 = add_series(ch, "Actual", calc.Range("M4:M15"), calc.Range("L4:L15"), color=P["gold"])
    s1.Points(12).Format.Fill.ForeColor.RGB = P["gold2"]
    add_series(ch, "Budget", calc.Range("N4:N15"), calc.Range("L4:L15"), ctype=4, color=P["bgrey"], line=True)
    s3 = add_series(ch, "Prior year", calc.Range("O4:O15"), calc.Range("L4:L15"), ctype=4, color=P["sky"], line=True)
    s3.Format.Line.DashStyle = 4; s3.Format.Line.Weight = 1.75
    ch.ChartGroups(1).GapWidth = 55
    no_title(ch); legend(ch); dark_axes(ch, '#,##0,"k"')

    panel(ws, 25, cc(25), 41, cc(39), "Budget to actual, by region", "=WfRegNote")
    co, ch = chart_in(ws, 27, cc(25), 41, cc(39))
    co.Name = "chart_waterfall_region"
    waterfall_chart(ch, calc, 42, 47)

    panel(ws, 25, cc(41), 41, cc(53), "Revenue by category", '="share of "&TEXT(Calc!B38,"#,##0")&", "&PeriodLabel')
    donut(ws, calc, 27, cc(41), 41, cc(53))

    # heatmap + region table rows 43-51
    ws.Rows(42).RowHeight = 10
    ws.Rows(43).RowHeight = 22; ws.Rows(44).RowHeight = 14
    for rr in range(45, 52):
        ws.Rows(rr).RowHeight = 18
    ws.Rows(52).RowHeight = 8
    panel(ws, 43, cc(0), 52, cc(29), "Heatmap: region x month, % vs budget", '="12 months to "&TEXT(RepMonth,"mmm yyyy")&"  |  red = below budget, green = above"')
    hr = 45
    text(ws, hr, cc(1), cc(3), "", 8)
    for m in range(12):
        a = cc(4) + m * 2
        text(ws, hr, a, a + 1, f"=Calc!L{4 + m}", 8, P["muted"], align=-4108)
    for i, reg in enumerate(REGIONS + ["Total"]):
        rr = hr + 1 + i
        text(ws, rr, cc(1), cc(3), reg, 9, P["ivory"], bold=reg == "Total")
        for m in range(12):
            a = cc(4) + m * 2
            if reg == "Total":
                f = f"=Calc!T{4 + m}-1"
                f = f'=IF(Calc!N{4 + m}=0,"",Calc!T{4 + m}-1)'
            else:
                f = f"=Calc!{col_letter(29 + i)}{4 + m}"
            cell = text(ws, rr, a, a + 1, f, 8, P["ivory"], bold=reg == "Total", align=-4108, fmt="+0%;-0%;0%")
            ws.Range(cell, ws.Cells(rr, a + 1)).Interior.Color = P["card2"]
            for e in (7, 8, 9, 10):
                bb = ws.Range(cell, ws.Cells(rr, a + 1)).Borders(e); bb.LineStyle = 1; bb.Color = P["card"]; bb.Weight = 2
    heat = rng(ws, hr + 1, cc(4), hr + 5, cc(4) + 23)
    cs = heat.FormatConditions.AddColorScale(3)
    for k, (v, col) in enumerate(((-0.12, P["heatbad"]), (0, P["card2"]), (0.08, P["heatgood"])), 1):
        cs.ColorScaleCriteria(k).Type = 0
        cs.ColorScaleCriteria(k).Value = v
        cs.ColorScaleCriteria(k).FormatColor.Color = col

    panel(ws, 43, cc(31), 52, cc(53), "Regions compared", '=PeriodLabel&" vs budget and "&PYLabel')
    cols = [("Region", 3, None, -4131), ("Actual", 3, TN, -4152), ("Budget", 3, TN, -4152), ("Var", 3, TNS, -4152),
            ("Var %", 3, TPS, -4152), ("vs PY", 3, TPS, -4152), ("Margin", 3, TP, -4152)]
    srcmap = ["A", "B", "C", "D", "E", "G", "I"]
    table(ws, 45, cc(32), cols, 5, lambda i, k: f"=Calc!{srcmap[k]}{24 + i}", total=True, sign_cols=(3, 4, 5))
    sel = rng(ws, 46, cc(32), 49, cc(52))
    fc = sel.FormatConditions.Add(2, None, cfx(f'=AND(NReg<4,COUNTIF(SelRegions,${col_letter(cc(32))}46)>0)'))
    fc.Font.Bold = True

    # checks strip  row 53
    ws.Rows(53).RowHeight = 22
    box(ws, 53, cc(0), 53, cc(53), color=P["bg0"])
    t = text(ws, 53, cc(1), cc(30), '="✔  "&COUNTIF(CheckResults,"PASS")&" of "&COUNTA(CheckResults)&" data checks PASS  (open Checks for the detail)"',
             9, P["good"], True)
    fc = rng(ws, 53, cc(1), 53, cc(30)).FormatConditions.Add(2, None, cfx('=COUNTIF(CheckResults,"PASS")<COUNTA(CheckResults)'))
    fc.Font.Color = P["bad"]
    ws.Hyperlinks.Add(Anchor=ws.Cells(53, cc(1)), Address="", SubAddress="'Checks'!A1", ScreenTip="Open the data checks")
    ws.Cells(53, cc(1)).Font.Underline = False; ws.Cells(53, cc(1)).Font.Color = P["good"]
    text(ws, 53, cc(31), cc(52), "Fictional data  |  Power Query + SUMIFS + slicers  |  no macros", 8, P["muted"], align=-4152)
    finish_page(ws, 54)


def finish_page(ws, last_row):
    ws.Rows(last_row).RowHeight = 10
    try:
        tip = ws.Shapes("slicer_tip")
        while ws.Cells(last_row + 1, 1).Top < tip.Top + tip.Height + 8:
            last_row += 1
    except Exception:
        pass
    ws.PageSetup.PrintArea = f"$A$1:${col_letter(NCOL)}${last_row}"


def waterfall_chart(ch, calc, first, last):
    ch.ChartType = 52
    x = calc.Range(f"H{first}:H{last}")
    base = add_series(ch, "Base", calc.Range(f"D{first}:D{last}"), x)
    base.Format.Fill.Visible = False; base.Format.Line.Visible = False
    tot = add_series(ch, "Total", calc.Range(f"E{first}:E{last}"), x, color=P["gold"])
    tot.Points(1).Format.Fill.ForeColor.RGB = P["bgrey"]
    add_series(ch, "Above budget", calc.Range(f"F{first}:F{last}"), x, color=P["good"])
    add_series(ch, "Below budget", calc.Range(f"G{first}:G{last}"), x, color=P["bad"])
    ch.ChartGroups(1).GapWidth = 35
    no_title(ch)
    ch.HasLegend = False
    dark_axes(ch, grid=False)
    ch.Axes(2).Delete()
    ch.Axes(1).TickLabels.Font.Color = P["ivory"]; ch.Axes(1).TickLabels.Font.Size = 8


def donut(ws, calc, r1, c1, r2, c2):
    split = c1 + 7
    co, ch = chart_in(ws, r1, c1, r2, split)
    co.Name = "chart_donut"
    ch.ChartType = -4120
    s = add_series(ch, "Revenue", calc.Range("B33:B37"), calc.Range("A33:A37"))
    for k in range(5):
        p = s.Points(k + 1)
        p.Format.Fill.ForeColor.RGB = SERIES[k]
        p.Format.Line.ForeColor.RGB = P["card"]; p.Format.Line.Weight = 1.5
    ch.ChartGroups(1).DoughnutHoleSize = 64
    no_title(ch); ch.HasLegend = False
    # centre number: text boxes linked to cells, on top of the chart
    cx = co.Left + co.Width / 2; cy = co.Top + co.Height / 2
    for k, (ref, size, color, dy) in enumerate((("Calc!$B$38", 14, P["ivory"], -7), ("Calc!$C$2", 7.5, P["muted"], 10))):
        tb = ws.Shapes.AddTextbox(1, cx - 50, cy + dy - 9, 100, 20)
        tb.Name = f"donut_centre_{k}"
        tb.DrawingObject.Formula = "=" + ref
        tf = tb.TextFrame2
        tf.TextRange.Font.Size = size; tf.TextRange.Font.Name = SERIF if k == 0 else SANS
        tf.TextRange.Font.Fill.ForeColor.RGB = color
        tf.TextRange.ParagraphFormat.Alignment = 2; tf.VerticalAnchor = 3
        tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
        tb.Fill.Visible = False; tb.Line.Visible = False
    # legend as cells: swatch, name, share
    for k in range(5):
        rr = r1 + 2 + k * 2
        text(ws, rr, split + 1, split + 1, "■", 10, SERIES[k], align=-4108)
        text(ws, rr, split + 2, c2 - 1, f"=Calc!A{33 + k}", 8.5, P["ivory"])
        text(ws, rr + 1, split + 2, c2 - 1, f'=IF(Calc!K{33 + k}="","",TEXT(Calc!K{33 + k},"0.0%")&"  |  "&TEXT(Calc!B{33 + k}/1000,"#,##0.0")&"k")', 8, P["muted"])


# ------------------------------------------------------------------------------------------------- Products
def build_products(wb):
    ws = wb.Worksheets("Products")
    calc = wb.Worksheets("Calc")
    r = 11
    card_rows(ws, r)
    glyph = ["⚙", "⚡", "◐", "❀", "⌁"]
    for k, cat in enumerate(CATS):
        a = cc(k * 11)
        b = a + 9 if k < 4 else cc(53)
        rowc = 33 + k
        card(ws, r, a, b, cat.upper(), glyph[k], f"=Calc!B{rowc}", "#,##0", f"Calc!{col_letter(33 + k)}4:{col_letter(33 + k)}15",
             [delta(f"Calc!E{rowc}", "vs budget"), delta(f"Calc!G{rowc}", "vs prior year"),
              (f'=IF(Calc!I{rowc}="","","margin "&TEXT(Calc!I{rowc},"0.0%")&"  |  share "&TEXT(Calc!K{rowc},"0.0%"))', None)])
    ws.Rows(18).RowHeight = 10
    ws.Rows(19).RowHeight = 22; ws.Rows(20).RowHeight = 14
    for rr in range(21, 33):
        ws.Rows(rr).RowHeight = 18
    panel(ws, 19, cc(0), 32, cc(29), "Categories compared", '=PeriodLabel&" vs budget and "&PYLabel')
    cols = [("Category", 5, None, -4131), ("Actual", 3, TN, -4152), ("Budget", 3, TN, -4152), ("Var", 3, TNS, -4152),
            ("Var %", 3, TPS, -4152), ("vs PY", 3, TPS, -4152), ("Margin", 3, TP, -4152), ("Bud. margin", 3, TP, -4152),
            ("Share", 2, TP, -4152)]
    srcmap = ["A", "B", "C", "D", "E", "G", "I", "J", "K"]
    table(ws, 22, cc(1), cols, 6, lambda i, k: f"=Calc!{srcmap[k]}{33 + i}", total=True, sign_cols=(3, 4, 5))
    text(ws, 30, cc(1), cc(28), "Margin red/green: compare Margin with Bud. margin (budget margin for the same lines).", 8, P["muted"], italic=True)
    mg = rng(ws, 23, cc(1) + 20, 28, cc(1) + 22)
    fc = mg.FormatConditions.Add(2, None, cfx(f"=AND(ISNUMBER(Calc!$I33),ISNUMBER(Calc!$J33),Calc!$I33<Calc!$J33)")); fc.Font.Color = P["bad"]
    fc = mg.FormatConditions.Add(2, None, cfx(f"=AND(ISNUMBER(Calc!$I33),ISNUMBER(Calc!$J33),Calc!$I33>=Calc!$J33)")); fc.Font.Color = P["good"]

    panel(ws, 19, cc(31), 32, cc(53), "Budget to actual, by category", "=WfCatNote")
    co, ch = chart_in(ws, 21, cc(31), 32, cc(53))
    co.Name = "chart_waterfall_category"
    waterfall_chart(ch, calc, 53, 59)

    ws.Rows(33).RowHeight = 10
    ws.Rows(34).RowHeight = 22; ws.Rows(35).RowHeight = 14
    for rr in range(36, 50):
        ws.Rows(rr).RowHeight = 15
    panel(ws, 34, cc(0), 49, cc(29), "Revenue by category, 12 months", '="stacked, to "&TEXT(RepMonth,"mmm yyyy")')
    co, ch = chart_in(ws, 36, cc(0), 49, cc(29))
    co.Name = "chart_category_trend"
    ch.ChartType = 52
    for k, cat in enumerate(CATS):
        add_series(ch, cat, calc.Range(f"{col_letter(33 + k)}4:{col_letter(33 + k)}15"), calc.Range("L4:L15"), color=SERIES[k])
    ch.ChartGroups(1).GapWidth = 45
    no_title(ch); legend(ch); dark_axes(ch, '#,##0,"k"')

    panel(ws, 34, cc(31), 49, cc(53), "Gross margin by category", '="actual vs budget margin, "&PeriodLabel')
    co, ch = chart_in(ws, 36, cc(31), 49, cc(53))
    co.Name = "chart_margin"
    ch.ChartType = 57
    add_series(ch, "Actual margin", calc.Range("I33:I37"), calc.Range("A33:A37"), color=P["gold"])
    add_series(ch, "Budget margin", calc.Range("J33:J37"), calc.Range("A33:A37"), color=P["bgrey"])
    ch.ChartGroups(1).GapWidth = 50; ch.ChartGroups(1).Overlap = -10
    no_title(ch); legend(ch); dark_axes(ch, "0%")
    ch.Axes(1).ReversePlotOrder = True
    ch.Axes(1).TickLabels.Font.Color = P["ivory"]
    finish_page(ws, 50)


# ------------------------------------------------------------------------------------------------- Customers
def build_customers(wb):
    ws = wb.Worksheets("Customers")
    calc = wb.Worksheets("Calc")
    r = 11
    card_rows(ws, r)
    specs = [("ACTIVE CUSTOMERS", "☺", "=CustActive", "0", [('="of "&CustOnFile&" on file"', None),
              ('=IF(CustActivePY=0,"",TEXT(CustActiveVsPY,"+0;-0;0")&" vs prior year")', "CustActiveVsPY"), ('=""', None)]),
             ("80% OF REVENUE", "∴", "=CustN80", "0", [('="customers bring 80% of revenue"', None),
              ('=TEXT(CustN80Pct,"0%")&" of active customers"', None), ('=""', None)]),
             ("TOP 10 SHARE", "★", "=CustTop10", "0.0%", [('="of revenue from 10 customers"', None),
              ('=IF(CustTop10VsPY="","","prior year "&TEXT(CustTop10VsPY,"0.0%"))', None), ('=""', None)]),
             ("LARGEST CUSTOMER", "♛", "=CustTopShare", "0.0%", [("=CustTopName", None), ('="share of revenue"', None), ('=""', None)]),
             ("REVENUE PER CUSTOMER", "$", "=CustAvg", "#,##0", [delta("CustAvgVsPY", "vs prior year"), ('="per active customer"', None), ('=""', None)])]
    for k, (label, icon, big, fmt, lines) in enumerate(specs):
        a = cc(k * 11)
        b = a + 9 if k < 4 else cc(53)
        card(ws, r, a, b, label, icon, big, fmt, None, lines)
    ws.Rows(18).RowHeight = 10
    ws.Rows(19).RowHeight = 22; ws.Rows(20).RowHeight = 14
    for rr in range(21, 36):
        ws.Rows(rr).RowHeight = 17
    panel(ws, 19, cc(0), 35, cc(30), "Pareto: who brings the revenue", '="top 20 customers, "&PeriodLabel&"  |  bars = revenue, line = cumulative share, dashed = 80%"')
    co, ch = chart_in(ws, 21, cc(0), 35, cc(30))
    co.Name = "chart_pareto"
    ch.ChartType = 51
    add_series(ch, "Revenue", calc.Range("AY4:AY23"), calc.Range("AV4:AV23"), color=P["gold"])
    s = add_series(ch, "Cumulative share", calc.Range("BD4:BD23"), calc.Range("AV4:AV23"), ctype=65, color=P["sky"], line=True, axis=2)
    s.MarkerStyle = 8; s.MarkerSize = 4; s.MarkerBackgroundColor = P["sky"]; s.MarkerForegroundColor = P["sky"]
    s2 = add_series(ch, "80%", calc.Range("BF4:BF23"), calc.Range("AV4:AV23"), ctype=4, color=P["bad"], line=True, axis=2)
    s2.Format.Line.DashStyle = 4; s2.Format.Line.Weight = 1.25
    ch.ChartGroups(1).GapWidth = 40
    no_title(ch); legend(ch); dark_axes(ch, '#,##0,"k"')
    ax2 = ch.Axes(2, 2)
    ax2.MinimumScale = 0; ax2.MaximumScale = 1; ax2.TickLabels.NumberFormat = "0%"; ax2.TickLabels.Font.Color = P["muted"]
    ax2.Format.Line.Visible = False
    ch.Axes(1).TickLabels.Font.Color = P["muted"]

    panel(ws, 19, cc(32), 35, cc(53), "Top 10 customers", '=PeriodLabel&"  |  share and cumulative share of revenue"')
    cols = [("#", 1, "0", -4108), ("Customer", 8, None, -4131), ("Revenue", 3, TN, -4152), ("vs PY", 3, TPS, -4152),
            ("Share", 3, TP, -4152), ("Cum.", 2, TP, -4152)]
    srcmap = ["AV", "AW", "AY", "BA", "BC", "BD"]
    table(ws, 22, cc(33), cols, 10, lambda i, k: f'=IF(Calc!$AW{4 + i}="","",Calc!{srcmap[k]}{4 + i})', sign_cols=(3,))
    bars = rng(ws, 23, cc(33) + 9, 32, cc(33) + 11).FormatConditions.AddDatabar()
    bars.BarColor.Color = P["copper"]; bars.MinPoint.Modify(0, 0)
    try:
        bars.BarFillType = 0
    except Exception:
        pass
    text(ws, 34, cc(33), cc(52), '="Customers beyond the top 10: "&TEXT(1-CustTop10,"0.0%")&" of revenue from "&MAX(CustActive-10,0)&" customers."', 8, P["muted"], italic=True)
    finish_page(ws, 36)


# ------------------------------------------------------------------------------------------------- Reps
def build_reps(wb):
    ws = wb.Worksheets("Reps")
    calc = wb.Worksheets("Calc")
    # gauges rows 11-20
    page_rows(ws, {11: 20, 12: 13, 13: 18, 14: 18, 15: 18, 16: 18, 17: 16, 18: 14, 19: 6, 20: 10})
    for k in range(6):
        a = cc(k * 9)
        b = a + 8 if k < 5 else cc(53)
        rr = 15 + k
        box(ws, 11, a, 19, b)
        text(ws, 11, a + 1, b - 1, f'=IF(Calc!BI{rr}="","-",Calc!BI{rr})', 10.5, P["ivory"], font=SERIF, valign=-4107)
        text(ws, 12, a + 1, b - 1, f'=IF(Calc!BI{rr}="","",Calc!BJ{rr}&"  |  #"&Calc!BH{rr})', 7.5, P["muted"])
        co, ch = chart_in(ws, 13, a, 16, b, pad=2)
        co.Name = f"gauge_{k + 1}"
        ch.ChartType = -4120
        s = ch.SeriesCollection().NewSeries()
        s.Name = "gauge"
        s.Values = calc.Range(f"BU{rr}:BX{rr}")
        for p, col in enumerate((P["gold"], P["good"], P["line"], None), 1):
            pt = s.Points(p)
            if col is None:
                pt.Format.Fill.Visible = False; pt.Format.Line.Visible = False
            else:
                pt.Format.Fill.ForeColor.RGB = col; pt.Format.Line.Visible = False
        ch.ChartGroups(1).FirstSliceAngle = 270
        ch.ChartGroups(1).DoughnutHoleSize = 62
        no_title(ch); ch.HasLegend = False
        lab = text(ws, 16, a + 1, b - 1, f"=Calc!BY{rr}", 15, P["ivory"], font=SERIF, align=-4108, valign=-4160)
        fc = rng(ws, 16, a + 1, 16, b - 1).FormatConditions.Add(2, None, cfx(f"=AND(ISNUMBER(Calc!$BM${rr}),Calc!$BM${rr}<1)")); fc.Font.Color = P["bad"]
        fc = rng(ws, 16, a + 1, 16, b - 1).FormatConditions.Add(2, None, cfx(f"=AND(ISNUMBER(Calc!$BM${rr}),Calc!$BM${rr}>=1)")); fc.Font.Color = P["good"]
        text(ws, 17, a + 1, b - 1, f"=Calc!BZ{rr}", 7.5, P["muted"], align=-4108)
        text(ws, 18, a + 1, b - 1, f'=IF(Calc!BO{rr}="","","vs prior year "&TEXT(Calc!BO{rr},"+0.0%;-0.0%"))', 7.5, P["muted"], align=-4108)
        cf_sign(rng(ws, 18, a + 1, 18, b - 1), f"Calc!$BO${rr}")
    # scorecard + chart rows 21-31
    ws.Rows(21).RowHeight = 22; ws.Rows(22).RowHeight = 14
    for rr in range(23, 33):
        ws.Rows(rr).RowHeight = 18
    panel(ws, 21, cc(0), 32, cc(33), "Scorecard", '=PeriodLabel&"  |  target = budget for the rep\'s lines  |  best % of target first"')
    cols = [("#", 1, "0", -4108), ("Sales rep", 6, None, -4131), ("Region", 3, None, -4131), ("Revenue", 3, TN, -4152),
            ("Target", 3, TN, -4152), ("% target", 3, TP, -4152), ("vs PY", 3, TPS, -4152), ("Orders", 3, TN, -4152),
            ("Avg order", 3, TN, -4152), ("Margin", 3, TP, -4152)]
    srcmap = ["BH", "BI", "BJ", "BK", "BL", "BM", "BO", "BP", "BQ", "BR"]
    table(ws, 24, cc(1), cols, 6, lambda i, k: f'=IF(Calc!$BI{15 + i}="","",Calc!{srcmap[k]}{15 + i})', sign_cols=(6,))
    pt = rng(ws, 25, cc(1) + 16, 30, cc(1) + 18)
    fc = pt.FormatConditions.Add(2, None, cfx(f"=AND(ISNUMBER(Calc!$BM15),Calc!$BM15<1)")); fc.Font.Color = P["bad"]; fc.Interior.Color = P["badbg"]
    fc = pt.FormatConditions.Add(2, None, cfx(f"=AND(ISNUMBER(Calc!$BM15),Calc!$BM15>=1)")); fc.Font.Color = P["good"]; fc.Interior.Color = P["goodbg"]
    text(ws, 31, cc(1), cc(32), '="Team: "&TEXT(RepTeamPct,"0.0%")&" of target  |  "&RepsOnTarget&" of "&RepsWithTarget&" reps on or above target"', 8.5, P["gold2"])

    panel(ws, 21, cc(35), 32, cc(53), "Revenue vs target", "=PeriodLabel")
    co, ch = chart_in(ws, 23, cc(35), 32, cc(53))
    co.Name = "chart_reps"
    ch.ChartType = 57
    add_series(ch, "Revenue", calc.Range("BK15:BK20"), calc.Range("BI15:BI20"), color=P["gold"])
    add_series(ch, "Target", calc.Range("BL15:BL20"), calc.Range("BI15:BI20"), color=P["bgrey"])
    ch.ChartGroups(1).GapWidth = 45; ch.ChartGroups(1).Overlap = -5
    no_title(ch); legend(ch); dark_axes(ch, '#,##0,"k"')
    ch.Axes(1).ReversePlotOrder = True; ch.Axes(1).TickLabels.Font.Color = P["ivory"]
    finish_page(ws, 33)


# ------------------------------------------------------------------------------------------------- Forecast
def build_forecast(wb):
    ws = wb.Worksheets("Forecast")
    calc = wb.Worksheets("Calc")
    r = 11
    card_rows(ws, r)
    specs = [("YEAR TO DATE", "$", "=FcYTD", "#,##0", [delta("FcYTDvsB", "vs YTD budget"), delta("FcYTDvsPY", "vs prior-year YTD"),
                                                       ('="Jan-"&TEXT(RepMonth,"mmm yyyy")', None)]),
             ("FULL-YEAR BUDGET", "◎", "=FcFYB", "#,##0", [('="budget "&FYear', None), ('="prior year "&TEXT(FcPY,"#,##0")', None), ('=""', None)]),
             ("YEAR-END FORECAST", "↗", "=FcFY", "#,##0", [delta("FcVsB", "vs full-year budget"), delta("FcVsPY", "vs prior year"),
                                                               ('="recent achievement "&TEXT(RecentAch,"0.0%")', None)]),
             ("GAP TO BUDGET", "Δ", "=FcGap", "+#,##0;-#,##0;0", [('="forecast minus budget"', None), ('=FcLeft&" months left"', None), ('=""', None)]),
             ("STRAIGHT RUN-RATE", "→", "=FcStraight", "#,##0", [delta("FcStrVsB", "vs full-year budget"),
                                                                     ('="YTD + 3-month average x months left"', None), ('=""', None)])]
    for k, (label, icon, big, fmt, lines) in enumerate(specs):
        a = cc(k * 11)
        b = a + 9 if k < 4 else cc(53)
        card(ws, r, a, b, label, icon, big, fmt, None, lines, big_ref="FcGap" if label == "GAP TO BUDGET" else None)
    ws.Rows(18).RowHeight = 10
    ws.Rows(19).RowHeight = 22; ws.Rows(20).RowHeight = 14
    for rr in range(21, 35):
        ws.Rows(rr).RowHeight = 16
    panel(ws, 19, cc(0), 34, cc(31), "Month by month", '=FYear&": actual to "&TEXT(RepMonth,"mmm")&", then forecast = monthly budget x recent achievement ("&TEXT(RecentAch,"0.0%")&")"')
    co, ch = chart_in(ws, 21, cc(0), 34, cc(31))
    co.Name = "chart_forecast"
    ch.ChartType = 52
    add_series(ch, "Actual", calc.Range("M20:M31"), calc.Range("L20:L31"), color=P["gold"])
    fcs = add_series(ch, "Forecast", calc.Range("P20:P31"), calc.Range("L20:L31"), color=P["copper"])
    add_series(ch, "Budget", calc.Range("N20:N31"), calc.Range("L20:L31"), ctype=4, color=P["bgrey"], line=True)
    py = add_series(ch, "Prior year", calc.Range("O20:O31"), calc.Range("L20:L31"), ctype=4, color=P["sky"], line=True)
    py.Format.Line.DashStyle = 4; py.Format.Line.Weight = 1.75
    ch.ChartGroups(1).GapWidth = 45
    no_title(ch); legend(ch); dark_axes(ch, '#,##0,"k"')

    panel(ws, 19, cc(33), 34, cc(53), "Landing: cumulative vs budget", '="cumulative "&FYear&", actual + forecast vs budget"')
    co, ch = chart_in(ws, 21, cc(33), 34, cc(53))
    co.Name = "chart_landing"
    ch.ChartType = 4
    add_series(ch, "Actual + forecast", calc.Range("R20:R31"), calc.Range("L20:L31"), color=P["gold"], line=True)
    b = add_series(ch, "Budget", calc.Range("S20:S31"), calc.Range("L20:L31"), color=P["bgrey"], line=True)
    b.Format.Line.DashStyle = 4
    no_title(ch); legend(ch); dark_axes(ch, '#,##0,"k"')

    # month table rows 36-43
    ws.Rows(35).RowHeight = 10
    ws.Rows(36).RowHeight = 22; ws.Rows(37).RowHeight = 14
    for rr in range(38, 44):
        ws.Rows(rr).RowHeight = 18
    panel(ws, 36, cc(0), 43, cc(53), "The numbers behind the forecast", '="italic copper = forecast  |  "&FYear')
    hdr = rng(ws, 38, cc(1), 38, cc(52)); fill(hdr, P["bg0"])
    text(ws, 38, cc(1), cc(4), "", 8)
    labels = [("Actual / forecast", "Q", NUM), ("Budget", "N", NUM), ("vs budget", "T", "+0%;-0%;0%"), ("Prior year", "O", NUM)]
    for m in range(12):
        a = cc(5) + m * 4
        text(ws, 38, a, a + 3, f"=Calc!L{20 + m}", 8, P["gold"], True, align=-4152)
    for i, (lab, col, fmt) in enumerate(labels):
        rr = 39 + i
        fill(rng(ws, rr, cc(1), rr, cc(52)), P["card2"] if i % 2 else P["card"])
        text(ws, rr, cc(1), cc(4), lab, 8.5, P["ivory"], bold=i == 0)
        for m in range(12):
            a = cc(5) + m * 4
            cell = text(ws, rr, a, a + 3, f"=Calc!{col}{20 + m}", 8.5, P["ivory"], align=-4152, fmt=fmt)
            if i == 0:
                fc = rng(ws, rr, a, rr, a + 3).FormatConditions.Add(2, None, cfx(f"=Calc!$K${20 + m}>RepMonth"))
                fc.Font.Color = P["copper"]; fc.Font.Italic = True
            if i == 2:
                cf_sign(rng(ws, rr, a, rr, a + 3), f"Calc!$T${20 + m}")
    text(ws, 43, cc(1), cc(52), '="Budget-phased run-rate: each month left = its budget x "&TEXT(RecentAch,"0.0%")&" (actual / budget, last 3 months). Straight run-rate for comparison: "&TEXT(FcStraight,"#,##0")&"."',
         8, P["muted"], italic=True)
    finish_page(ws, 44)


# ------------------------------------------------------------------------------------------------- Checks
CHECKS = [
    ("Sales lines loaded = rows in the sales table", "=INDEX(Checks[Value],1)", "=ROWS(Sales[Date])"),
    ("Every month/rep/category sold has a budget (combinations without one)", "=INDEX(Checks[Value],3)", "=0"),
    ("Budget has no repeated month/rep/category (repeated rows)", "=INDEX(Checks[Value],5)", "=0"),
    ("Revenue in the Sales table = revenue counted by Power Query", "=ROUND(SUM(Sales[Revenue]),2)", "=ROUND(INDEX(Checks[Value],6),2)"),
    ("Region table total = revenue KPI", "=ROUND(Calc!B28,2)", "=ROUND(Calc!E4,2)"),
    ("Category table total = revenue KPI", "=ROUND(Calc!B38,2)", "=ROUND(Calc!E4,2)"),
    ("Customer list total = revenue KPI", "=ROUND(CustTotal,2)", "=ROUND(Calc!E4,2)"),
    ("Sales rep total = revenue KPI", "=ROUND(SUM(Reps[Revenue]),2)", "=ROUND(Calc!E4,2)"),
    ("Sales rep orders = orders KPI", "=SUM(Reps[Orders])", "=Calc!E6"),
    ("Waterfall steps land on actual (region)", "=ROUND(Calc!C46,2)", "=ROUND(Calc!B47,2)"),
    ("Trend chart, reporting month = month revenue", "=ROUND(Calc!M15,2)", "=ROUND(Calc!G4,2)"),
    ("Every slicer keeps at least one item (regions x categories x reps x months)", "=NReg*NCat*NRep*COUNT(SelMonths)>0", "=TRUE"),
]
CHECK_FMT = ["#,##0", "#,##0", "#,##0", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0", "#,##0.00", "#,##0.00", "General"]


def build_checks(wb):
    ws = wb.Worksheets("Checks")
    page_rows(ws, {6: 10, 7: 22, 8: 18})
    n = len(CHECKS)
    last = 8 + n + 2
    panel(ws, 6, cc(0), last, cc(53), "Data checks", None)
    ws.Rows(6).RowHeight = 24
    text(ws, 7, cc(1), cc(52), '="They run on every refresh and every slicer click. Now: "&SelText', 8.5, P["muted"])
    cols = [("#", 1, "0", -4108), ("Check", 26, None, -4131), ("Left side", 8, "#,##0.##", -4152), ("Right side", 8, "#,##0.##", -4152),
            ("Result", 5, None, -4108)]
    table(ws, 8, cc(1), cols, n, lambda i, k: [i + 1, CHECKS[i][0], CHECKS[i][1], CHECKS[i][2],
                                                 f'=IF({col_letter(cc(1) + 27)}{9 + i}={col_letter(cc(1) + 35)}{9 + i},"PASS","FAIL")'][k])
    res = rng(ws, 9, cc(1) + 43, 8 + n, cc(1) + 47)
    wb.Names.Add("CheckResults", f"=Checks!${col_letter(cc(1) + 43)}$9:${col_letter(cc(1) + 43)}${8 + n}")
    for i, f in enumerate(CHECK_FMT):
        for off in (27, 35):
            ws.Cells(9 + i, cc(1) + off).NumberFormat = f
    for val, bg, fg in (("PASS", P["goodbg"], P["good"]), ("FAIL", P["badbg"], P["bad"])):
        fc = res.FormatConditions.Add(1, 3, f'="{val}"'); fc.Interior.Color = bg; fc.Font.Color = fg; fc.Font.Bold = True
    for i in range(n):
        for off in (27, 35):
            ws.Cells(9 + i, cc(1) + off).HorizontalAlignment = -4152
    for rr in range(9, 9 + n):
        ws.Rows(rr).RowHeight = 18
    text(ws, 9 + n + 1, cc(1), cc(52), "Power Query also counts rows, invoices and budget keys while loading (Calc sheet, Checks table); "
         "the test script test/verify.py recomputes every page in Python from the CSV files.", 8, P["muted"], italic=True)
    finish_page(ws, last + 1)


def build_howto(wb):
    ws = wb.Worksheets("How to use")
    ws.Columns("A").ColumnWidth = 120
    text_ = [("How to use this dashboard", True), ("", False),
             ("Every day", True),
             ("1. Use the slicers on the left of any page: View (MTD or YTD), Region, Category, Sales rep. Ctrl+click picks several; the funnel-x clears.", False),
             ("2. Use the Month timeline at the top: the last month you pick is the reporting month. MTD = that month, YTD = January to that month.", False),
             ("3. All six pages follow the same slicers. Navigate with the gold buttons at the top right.", False),
             ("4. Green = better than budget or the prior period, red = worse. The key insights on Overview rewrite themselves.", False),
             ("", False), ("Every month", True),
             ("1. Export sales lines and the budget as CSV into the data folder (see Settings).", False),
             ("2. Data > Refresh All. New months appear on the timeline automatically.", False),
             ("3. Open Checks: all twelve must say PASS.", False),
             ("", False), ("How it is built", True),
             ("Power Query (Data > Queries & Connections) loads and types the CSVs: Sales, Budget, Customers, Reps, Checks.", False),
             ("Slicer engine: each slicer filters a one-field PivotTable; the Sales and Budget tables get an In column (1/0) that reads what those PivotTables show.", False),
             ("Calc holds every number the pages show as SUMIFS on the Sales and Budget tables with In = 1, so any figure can be traced.", False),
             ("Orders = invoices with at least one line in the selected categories. Target = the budget for the rep's region and categories.", False),
             ("Forecast = actual to the reporting month, then each month left = its budget x the last 3 months' actual/budget.", False),
             ("No macros. Works in Excel 2016 and Microsoft 365 for Windows (slicers, timeline and sparklines are standard Excel).", False)]
    for i, (t, bold) in enumerate(text_, 1):
        ws.Cells(i, 1).Value = t
        ws.Cells(i, 1).Font.Bold = bold
    ws.Range("A1").Font.Size = 16


if __name__ == "__main__":
    main()
