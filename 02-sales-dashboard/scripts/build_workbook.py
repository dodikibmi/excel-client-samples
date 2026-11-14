"""Builds Sales-Dashboard.xlsx with real Excel (COM): Power Query loads sales lines and budget from CSV,
a Calc sheet turns them into numbers for the selected month and region, and the Dashboard sheet shows
KPI cards, budget-vs-actual charts, a region table, top 10 customers, a sales-rep leaderboard and data checks.
No macros: pick a month and a region from the two drop-downs; new data = Data > Refresh All.

Run with a Python that has pywin32:  python scripts/build_workbook.py
"""
import time
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Sales-Dashboard.xlsx"


def rgb(r, g, b):
    return r + g * 256 + b * 65536


NAVY, TEAL, ORANGE, GREY = (24, 45, 72), (0, 128, 128), (230, 126, 34), (110, 110, 110)
GREEN_BG, GREEN_FG, RED_BG, RED_FG = (198, 239, 206), (0, 97, 0), (255, 199, 206), (156, 0, 6)

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
    Revenue = Table.AddColumn(Month, "Revenue", each Number.Round([Qty] * [Unit Price], 2), type number),
    Cost    = Table.AddColumn(Revenue, "Cost", each Number.Round([Qty] * [Unit Cost], 2), type number),
    Margin  = Table.AddColumn(Cost, "Margin", each Number.Round([Revenue] - [Cost], 2), type number),
    // one invoice = one order: flag its first line so orders can be counted with SUMIFS
    Grouped = Table.Group(Margin, {"Invoice"}, {{"Rows", each Table.AddIndexColumn(_, "Line", 1, 1, Int64.Type), type table}}),
    Expand  = Table.ExpandTableColumn(Grouped, "Rows", {"Date", "Customer", "Region", "Sales Rep", "Category", "Qty", "Unit Price",
                  "Unit Cost", "Month", "Revenue", "Cost", "Margin", "Line"}),
    Order   = Table.AddColumn(Expand, "Order", each if [Line] = 1 then 1 else 0, Int64.Type),
    Sorted  = Table.Sort(Order, {{"Date", Order.Ascending}, {"Invoice", Order.Ascending}, {"Line", Order.Ascending}}),
    Result  = Table.TransformColumnTypes(Table.RemoveColumns(Sorted, {"Line"}), {{"Date", type date}, {"Month", type date},
                  {"Customer", type text}, {"Region", type text}, {"Sales Rep", type text}, {"Category", type text},
                  {"Qty", Int64.Type}, {"Unit Price", type number}, {"Unit Cost", type number}, {"Revenue", type number},
                  {"Cost", type number}, {"Margin", type number}})
in
    Result"""
Q["Budget"] = """let
    Source = Budget_Raw,
    Typed  = Table.TransformColumns(Source, {
                 {"Month", each Date.FromText(_ & "-01", [Format = "yyyy-MM-dd", Culture = "en-US"]), type date},
                 {"Budget Revenue", each Number.FromText(_, "en-US"), type number},
                 {"Budget Margin %", each Number.FromText(_, "en-US"), type number}})
in
    Typed"""
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
    Keys    = Table.Distinct(Table.SelectColumns(S, {"Month", "Region", "Category"})),
    NoBud   = Table.NestedJoin(Keys, {"Month", "Region", "Category"}, B, {"Month", "Region", "Category"}, "B", JoinKind.LeftAnti),
    BudKeys = Table.RowCount(Table.Distinct(Table.SelectColumns(B, {"Month", "Region", "Category"})))
in
    #table(type table [Check = text, Value = number], {
        {"Sales lines loaded", Table.RowCount(S)},
        {"Invoices (orders)", List.Sum(S[Order])},
        {"Month/region/category combinations sold without a budget", Table.RowCount(NoBud)},
        {"Budget rows", Table.RowCount(B)},
        {"Budget rows that repeat a month/region/category", Table.RowCount(B) - BudKeys},
        {"Total revenue, all months", List.Sum(S[Revenue])}})"""

SHEETS = ["Dashboard", "Calc", "Sales data", "Budget data", "Settings", "How to use"]


def load_query(ws, name, at="A1"):
    conn = f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=""'
    lo = ws.ListObjects.Add(0, conn, None, 1, ws.Range(at))
    qt = lo.QueryTable
    qt.CommandType = 2
    qt.CommandText = f"SELECT * FROM [{name}]"
    qt.RefreshStyle = 1
    qt.AdjustColumnWidth = False
    qt.Refresh(False)
    lo.Name = name
    lo.TableStyle = "TableStyleLight9"
    return lo


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
        build_settings(wb)
        wb.SaveAs(str(OUT), 51)
        xl.Calculate()
        for name, m in Q.items():
            wb.Queries.Add(name, m)
        t0 = time.time()
        sales = load_query(wb.Worksheets("Sales data"), "Sales")
        load_query(wb.Worksheets("Budget data"), "Budget")
        calc = wb.Worksheets("Calc")
        load_query(calc, "Customers", "P4")
        load_query(calc, "Reps", "W4")
        load_query(calc, "Checks", "AC4")
        print(f"queries loaded in {time.time() - t0:.1f}s")
        format_data(wb)
        build_calc(wb)
        build_dashboard(wb)
        build_howto(wb)
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            ws.Range("A1").Select()
        wb.Worksheets("Dashboard").Activate()
        xl.ActiveWindow.Zoom = 90
        wb.Save()
        print("saved", OUT.name)
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        xl.Quit()


# ------------------------------------------------------------------------------------------------- Settings
def build_settings(wb):
    st = wb.Worksheets("Settings")
    st.Range("A1").Value = "Settings"
    st.Range("A1").Font.Size = 16; st.Range("A1").Font.Bold = True
    rows = [("Data folder", '=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"data\\"', "DataFolder",
             "Folder with the CSV exports. Default: the data folder next to this workbook. Type a path to override."),
            ("Sales file", "sales-lines.csv", "SalesFile", "One row per invoice line: date, invoice, customer, region, rep, category, qty, price, cost"),
            ("Budget file", "budget.csv", "BudgetFile", "One row per month, region and category"),
            ("Company", "Brightline Supply Co.", "Company", "Shown in the dashboard title (fictional company)")]
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
        c.Interior.Color = rgb(255, 255, 153)
        st.Cells(r, 3).Value = note
        wb.Names.Add(nm, f"=Settings!$B${r}")
    wb.Names.Add("SelMonth", "=Dashboard!$C$4")
    wb.Names.Add("SelRegion", "=Dashboard!$G$4")
    st.Columns("A").ColumnWidth = 16; st.Columns("B").ColumnWidth = 50; st.Columns("C").ColumnWidth = 90


def format_data(wb):
    ws = wb.Worksheets("Sales data")
    lo = ws.ListObjects("Sales")
    for c in ("Date", "Month"):
        lo.ListColumns(c).DataBodyRange.NumberFormat = "dd mmm yyyy" if c == "Date" else "mmm yyyy"
    for c in ("Unit Price", "Unit Cost", "Revenue", "Cost", "Margin"):
        lo.ListColumns(c).DataBodyRange.NumberFormat = "#,##0.00"
    for i in range(1, lo.ListColumns.Count + 1):
        ws.Columns(i).ColumnWidth = {"Customer": 26, "Sales Rep": 14, "Category": 12, "Invoice": 11}.get(lo.ListColumns(i).Name, 11)
    ws.Activate(); ws.Range("A2").Select(); wb.Application.ActiveWindow.FreezePanes = True
    ws = wb.Worksheets("Budget data")
    lo = ws.ListObjects("Budget")
    lo.ListColumns("Month").DataBodyRange.NumberFormat = "mmm yyyy"
    lo.ListColumns("Budget Revenue").DataBodyRange.NumberFormat = "#,##0"
    lo.ListColumns("Budget Margin %").DataBodyRange.NumberFormat = "0%"
    for i in range(1, lo.ListColumns.Count + 1):
        ws.Columns(i).ColumnWidth = 16


# ------------------------------------------------------------------------------------------------- Calc
CATS = ["Fasteners", "Power Tools", "Paint", "Garden", "Electrical"]
REGIONS = ["North", "South", "East", "West"]


def build_calc(wb):
    c = wb.Worksheets("Calc")
    c.Range("A1").Value = "Calc - every number on the Dashboard comes from this sheet (driven by the two drop-downs)"
    c.Range("A1").Font.Bold = True
    c.Range("H2").Value = "Region filter"
    c.Range("I2").Formula = '=IF(SelRegion="All","*",SelRegion)'
    wb.Names.Add("RegionCrit", "=Calc!$I$2")
    wb.Names.Add("FirstMonth", "=Calc!$A$4")
    S = lambda col, extra="": f"SUMIFS(Sales[{col}],Sales[Month],{{m}},Sales[Region],RegionCrit{extra})"

    # 12-month trend  A3:F15
    c.Range("A3:F3").Value = ("Month", "Label", "Actual", "Budget", "Margin %", "Variance %")
    for i in range(12):
        r = 4 + i
        c.Range(f"A{r}").Formula = "=MIN(Budget[Month])" if i == 0 else f"=EDATE(A{r - 1},1)"
        c.Range(f"B{r}").Formula = f'=TEXT(A{r},"mmm yy")'
        c.Range(f"C{r}").Formula = "=" + S("Revenue").format(m=f"A{r}")
        c.Range(f"D{r}").Formula = f"=SUMIFS(Budget[Budget Revenue],Budget[Month],A{r},Budget[Region],RegionCrit)"
        c.Range(f"E{r}").Formula = f"=IFERROR({S('Margin').format(m=f'A{r}')}/C{r},0)"
        c.Range(f"F{r}").Formula = f"=IFERROR(C{r}/D{r}-1,0)"
    c.Range("A4:A15").NumberFormat = "mmm yyyy"
    wb.Names.Add("MonthList", "=Calc!$A$4:$A$15")
    # region drop-down reads a range: a typed list "All,North,..." breaks on PCs whose list separator is ";"
    c.Range("H18").Value = "Region list"
    for i, reg in enumerate(["All"] + REGIONS, 19):
        c.Range(f"H{i}").Value = reg
    wb.Names.Add("RegionList", "=Calc!$H$19:$H$23")

    # KPIs  H3:I16
    k = [("Revenue", "=" + S("Revenue").format(m="SelMonth")),
         ("Budget", "=SUMIFS(Budget[Budget Revenue],Budget[Month],SelMonth,Budget[Region],RegionCrit)"),
         ("Vs budget", "=IFERROR(I4/I5-1,0)"),
         ("Budget achieved", "=IFERROR(I4/I5,0)"),
         ("Margin %", "=IFERROR(" + S("Margin").format(m="SelMonth") + "/I4,0)"),
         ("Margin % last month", "=IFERROR(" + S("Margin").format(m="EDATE(SelMonth,-1)") + "/" + S("Revenue").format(m="EDATE(SelMonth,-1)") + ",0)"),
         ("Margin change (pp)", "=(I8-I9)*100"),
         ("Orders", "=" + S("Order").format(m="SelMonth")),
         ("Average order", "=IFERROR(I4/I11,0)"),
         ("YTD revenue", '=SUMIFS(Sales[Revenue],Sales[Month],">="&FirstMonth,Sales[Month],"<="&SelMonth,Sales[Region],RegionCrit)'),
         ("YTD budget", '=SUMIFS(Budget[Budget Revenue],Budget[Month],">="&FirstMonth,Budget[Month],"<="&SelMonth,Budget[Region],RegionCrit)'),
         ("YTD vs budget", "=IFERROR(I13/I14-1,0)"),
         ("Revenue last month", "=" + S("Revenue").format(m="EDATE(SelMonth,-1)"))]
    c.Range("H3:I3").Value = ("KPI", "Value")
    for i, (label, f) in enumerate(k, 4):
        c.Range(f"H{i}").Value = label
        c.Range(f"I{i}").Formula = f

    # Category  A18:E23
    c.Range("A18:E18").Value = ("Category", "Actual", "Budget", "Variance", "Variance %")
    for i, cat in enumerate(CATS, 19):
        c.Range(f"A{i}").Value = cat
        c.Range(f"B{i}").Formula = f"=SUMIFS(Sales[Revenue],Sales[Month],SelMonth,Sales[Region],RegionCrit,Sales[Category],A{i})"
        c.Range(f"C{i}").Formula = f"=SUMIFS(Budget[Budget Revenue],Budget[Month],SelMonth,Budget[Region],RegionCrit,Budget[Category],A{i})"
        c.Range(f"D{i}").Formula = f"=B{i}-C{i}"
        c.Range(f"E{i}").Formula = f"=IFERROR(B{i}/C{i}-1,0)"

    # Region (always all regions, selected month)  A26:F31
    c.Range("A26:F26").Value = ("Region", "Actual", "Budget", "Variance", "Variance %", "Margin %")
    for i, reg in enumerate(REGIONS + ["Total"], 27):
        crit = '"*"' if reg == "Total" else f"A{i}"
        c.Range(f"A{i}").Value = reg
        c.Range(f"B{i}").Formula = f"=SUMIFS(Sales[Revenue],Sales[Month],SelMonth,Sales[Region],{crit})"
        c.Range(f"C{i}").Formula = f"=SUMIFS(Budget[Budget Revenue],Budget[Month],SelMonth,Budget[Region],{crit})"
        c.Range(f"D{i}").Formula = f"=B{i}-C{i}"
        c.Range(f"E{i}").Formula = f"=IFERROR(B{i}/C{i}-1,0)"
        c.Range(f"F{i}").Formula = f"=IFERROR(SUMIFS(Sales[Margin],Sales[Month],SelMonth,Sales[Region],{crit})/B{i},0)"

    # Customers table (query) + formula columns, then top 10  L3:N13
    cust = c.ListObjects("Customers")
    for name, f in (("Revenue", "=SUMIFS(Sales[Revenue],Sales[Customer],[@Customer],Sales[Month],SelMonth,Sales[Region],RegionCrit)"),
                    ("RankKey", "=[@Revenue]+ROW()/1000000000")):
        col = cust.ListColumns.Add()
        col.Name = name
        col.DataBodyRange.Formula = f
    c.Range("L3:N3").Value = ("Top", "Customer", "Revenue")
    for n in range(1, 11):
        r = 3 + n
        c.Range(f"L{r}").Value = n
        c.Range(f"M{r}").Formula = (f'=IF(LARGE(Customers[RankKey],L{r})<0.5,"",'
                                    f'INDEX(Customers[Customer],MATCH(LARGE(Customers[RankKey],L{r}),Customers[RankKey],0)))')
        c.Range(f"N{r}").Formula = f'=IF(M{r}="",0,ROUND(LARGE(Customers[RankKey],L{r}),2))'

    # Reps table (query) + formula columns
    reps = c.ListObjects("Reps")
    for name, f in (("Revenue", "=SUMIFS(Sales[Revenue],Sales[Sales Rep],[@[Sales Rep]],Sales[Month],SelMonth,Sales[Region],RegionCrit)"),
                    ("Orders", "=SUMIFS(Sales[Order],Sales[Sales Rep],[@[Sales Rep]],Sales[Month],SelMonth,Sales[Region],RegionCrit)"),
                    ("Margin %", "=IFERROR(SUMIFS(Sales[Margin],Sales[Sales Rep],[@[Sales Rep]],Sales[Month],SelMonth,Sales[Region],RegionCrit)/[@Revenue],0)"),
                    ("Share", "=IFERROR([@Revenue]/SUM([Revenue]),0)"),
                    ("RankKey", "=[@Revenue]+ROW()/1000000000")):
        col = reps.ListColumns.Add()
        col.Name = name
        col.DataBodyRange.Formula = f

    for rng, fmt in (("C4:D15", "#,##0"), ("E4:F15", "0.0%"), ("I4:I16", "#,##0.00"), ("B19:D23", "#,##0"), ("E19:E23", "0.0%"),
                     ("B27:D31", "#,##0"), ("E27:F31", "0.0%"), ("N4:N13", "#,##0")):
        c.Range(rng).NumberFormat = fmt
    for col in "ABCDEFHIMN":
        c.Columns(col).ColumnWidth = 14
    c.Columns("H").ColumnWidth = 22; c.Columns("M").ColumnWidth = 26; c.Columns("P").ColumnWidth = 26
    c.Columns("W").ColumnWidth = 14; c.Columns("AC").ColumnWidth = 58; c.Columns("AD").ColumnWidth = 14
    c.ListObjects("Checks").ListColumns("Value").DataBodyRange.NumberFormat = "#,##0.##"


# ------------------------------------------------------------------------------------------------- Dashboard
CARD_COLS = [("B", "D"), ("F", "H"), ("J", "L"), ("N", "P"), ("R", "T")]


def font(rng, size=None, bold=None, color=None):
    if size: rng.Font.Size = size
    if bold is not None: rng.Font.Bold = bold
    if color: rng.Font.Color = rgb(*color)


def heading(ws, cell, text):
    ws.Range(cell).Value = text
    font(ws.Range(cell), 13, True, NAVY)


def build_dashboard(wb):
    d = wb.Worksheets("Dashboard")
    d.Cells.Font.Name = "Calibri"
    d.Columns("A").ColumnWidth = 2
    for col in "BCDFGHJKLNOPRST":
        d.Columns(col).ColumnWidth = 10.5
    for col in "EIMQ":
        d.Columns(col).ColumnWidth = 2
    d.Columns("U").ColumnWidth = 2

    # Title + selectors
    d.Range("A1:U2").Interior.Color = rgb(*NAVY)
    d.Range("B1").Formula = '="Sales dashboard   |   "&Company&"   |   "&TEXT(SelMonth,"mmmm yyyy")&IF(SelRegion="All",""," | "&SelRegion&" region")'
    font(d.Range("B1"), 18, True, (255, 255, 255))
    d.Range("B2").Value = "Budget vs actual. Fictional data. Change the month or region below; new data: Data > Refresh All."
    font(d.Range("B2"), 10, False, (200, 215, 230))
    d.Rows(1).RowHeight = 34; d.Rows(2).RowHeight = 18; d.Rows(3).RowHeight = 8
    for label, cell, name in (("Month", "C4", "SelMonth"), ("Region", "G4", "SelRegion")):
        lab = d.Range(cell).Offset(1, 0)  # B4 / F4
        lab.Value = label; font(lab, 11, True, NAVY); lab.HorizontalAlignment = -4152
        sel = d.Range(f"{cell}:{chr(ord(cell[0]) + 1)}4")
        sel.Merge()
        sel.Interior.Color = rgb(255, 242, 204)
        sel.Borders.LineStyle = 1; sel.Borders.Color = rgb(191, 143, 0)
        font(sel, 12, True, NAVY); sel.HorizontalAlignment = -4108
        wb.Names.Add(name, f"=Dashboard!${cell[0]}$4")
    d.Range("C4").Validation.Add(3, 1, 1, "=MonthList")
    d.Range("C4").NumberFormat = "mmm yyyy"
    d.Range("C4").Formula = "=MAX(MonthList)"
    d.Range("C4").Value = d.Range("C4").Value  # a fixed date, not a formula, so the drop-down can change it
    d.Range("G4").Validation.Add(3, 1, 1, "=RegionList")
    d.Range("G4").Value = "All"
    d.Range("J4").Value = "Pick a month and a region: every number and chart below follows."
    font(d.Range("J4"), 10, False, GREY)
    d.Rows(5).RowHeight = 8

    # KPI cards rows 6-10
    cards = [("REVENUE", "=Calc!I4", "#,##0", '=TEXT(Calc!I6,"+0.0%;-0.0%")&" vs budget"', "Calc!I6"),
             ("BUDGET ACHIEVED", "=Calc!I7", "0.0%", '="budget "&TEXT(Calc!I5,"#,##0")', "Calc!I6"),
             ("GROSS MARGIN", "=Calc!I8", "0.0%", '=TEXT(Calc!I10,"+0.0;-0.0")&" pp vs last month"', "Calc!I10"),
             ("ORDERS", "=Calc!I11", "#,##0", '="average order "&TEXT(Calc!I12,"#,##0")', None),
             ("YEAR TO DATE", "=Calc!I13", "#,##0", '=TEXT(Calc!I15,"+0.0%;-0.0%")&" vs YTD budget"', "Calc!I15")]
    for (a, b), (label, value, fmt, note, sign) in zip(CARD_COLS, cards):
        box = d.Range(f"{a}6:{b}10")
        box.Interior.Color = rgb(247, 249, 252)
        for edge in (7, 8, 9, 10):
            box.Borders(edge).LineStyle = 1; box.Borders(edge).Color = rgb(210, 218, 228)
        d.Range(f"{a}6:{b}6").Interior.Color = rgb(*TEAL)
        d.Rows(6).RowHeight = 5
        for row, val, size, bold, color in ((7, label, 9, True, (90, 100, 115)), (8, value, 24, True, NAVY), (10, note, 9, True, GREY)):
            rng = d.Range(f"{a}{row}:{b}{row + (1 if row == 8 else 0)}")
            rng.Merge()
            cell = d.Range(f"{a}{row}")
            if val.startswith("="):
                cell.Formula = val
            else:
                cell.Value = val
            font(cell, size, bold, color)
            rng.HorizontalAlignment = -4108; rng.VerticalAlignment = -4108
        d.Range(f"{a}8").NumberFormat = fmt
        if sign:
            n = d.Range(f"{a}10")
            fc = n.FormatConditions.Add(2, None, f"={sign}<0"); fc.Font.Color = rgb(192, 0, 0)
            fc = n.FormatConditions.Add(2, None, f"={sign}>=0"); fc.Font.Color = rgb(0, 128, 0)
    d.Rows(8).RowHeight = 22; d.Rows(9).RowHeight = 18
    d.Rows(11).RowHeight = 10

    # Charts row 12-27
    calc = wb.Worksheets("Calc")
    heading(d, "B12", "Revenue vs budget, last 12 months")
    co = d.ChartObjects().Add(d.Range("B13").Left, d.Range("B13").Top, d.Range("B13:L13").Width, d.Range("B13:B27").Height)
    ch = co.Chart
    ch.ChartType = 51  # clustered column
    s1 = ch.SeriesCollection().NewSeries(); s1.Name = "Actual"; s1.Values = calc.Range("C4:C15"); s1.XValues = calc.Range("B4:B15")
    s2 = ch.SeriesCollection().NewSeries(); s2.Name = "Budget"; s2.Values = calc.Range("D4:D15"); s2.XValues = calc.Range("B4:B15")
    s1.Format.Fill.ForeColor.RGB = rgb(*TEAL)
    s2.ChartType = 65  # line with markers
    s2.Format.Line.ForeColor.RGB = rgb(*ORANGE); s2.Format.Line.Weight = 2.25
    s2.MarkerBackgroundColor = rgb(*ORANGE); s2.MarkerForegroundColor = rgb(*ORANGE)
    s1.Points(12).Format.Fill.ForeColor.RGB = rgb(*NAVY)
    style_chart(ch)
    ch.Axes(2).TickLabels.NumberFormat = '#,##0,"k"'

    heading(d, "N12", "By category, selected month")
    co = d.ChartObjects().Add(d.Range("N13").Left, d.Range("N13").Top, d.Range("N13:T13").Width, d.Range("N13:N27").Height)
    ch = co.Chart
    ch.ChartType = 57  # clustered bar
    s1 = ch.SeriesCollection().NewSeries(); s1.Name = "Actual"; s1.Values = calc.Range("B19:B23"); s1.XValues = calc.Range("A19:A23")
    s2 = ch.SeriesCollection().NewSeries(); s2.Name = "Budget"; s2.Values = calc.Range("C19:C23"); s2.XValues = calc.Range("A19:A23")
    s1.Format.Fill.ForeColor.RGB = rgb(*TEAL); s2.Format.Fill.ForeColor.RGB = rgb(*ORANGE)
    style_chart(ch)
    ch.Axes(1).ReversePlotOrder = True
    ch.Axes(2).TickLabels.NumberFormat = '#,##0,"k"'
    d.Rows(28).RowHeight = 10

    # Region table rows 29-35 (B..H) and Top 10 chart (J..T)
    heading(d, "B29", "Budget vs actual by region, selected month")
    cols = ["B", "C", "D", "F", "G", "H"]
    for c, t in zip(cols, ("Region", "Actual", "Budget", "Variance", "Var %", "Margin %")):
        h = d.Range(f"{c}30"); h.Value = t
        font(h, 10, True, (255, 255, 255)); h.Interior.Color = rgb(*NAVY)
        if c != "B":
            h.HorizontalAlignment = -4152
    d.Range("E30").Interior.Color = rgb(*NAVY)
    for i in range(5):
        r, src = 31 + i, 27 + i
        for c, s in zip(cols, "ABCDEF"):
            d.Range(f"{c}{r}").Formula = f"=Calc!{s}{src}"
        d.Range(f"C{r}:D{r}").NumberFormat = "#,##0"; d.Range(f"F{r}").NumberFormat = "+#,##0;-#,##0;0"
        d.Range(f"G{r}:H{r}").NumberFormat = "0.0%"
        d.Range(f"B{r}:H{r}").Borders(9).LineStyle = 1; d.Range(f"B{r}:H{r}").Borders(9).Color = rgb(220, 220, 220)
    d.Range("B35:H35").Font.Bold = True
    d.Range("B35:H35").Borders(8).LineStyle = 1
    var = d.Range("G31:G35")
    fc = var.FormatConditions.Add(1, 6, "=0"); fc.Interior.Color = rgb(*RED_BG); fc.Font.Color = rgb(*RED_FG)  # less than 0
    fc = var.FormatConditions.Add(1, 7, "=0"); fc.Interior.Color = rgb(*GREEN_BG); fc.Font.Color = rgb(*GREEN_FG)  # >= 0
    sel = d.Range("B31:H34")
    fc = sel.FormatConditions.Add(2, None, '=$B31=SelRegion'); fc.Font.Bold = True; fc.Interior.Color = rgb(255, 242, 204)

    heading(d, "J29", "Top 10 customers, selected month and region")
    co = d.ChartObjects().Add(d.Range("J30").Left, d.Range("J30").Top, d.Range("J30:T30").Width, d.Range("J30:J46").Height)
    ch = co.Chart
    ch.ChartType = 57
    s1 = ch.SeriesCollection().NewSeries(); s1.Name = "Revenue"; s1.Values = calc.Range("N4:N13"); s1.XValues = calc.Range("M4:M13")
    s1.Format.Fill.ForeColor.RGB = rgb(*TEAL); s1.HasDataLabels = True
    s1.DataLabels().NumberFormat = "#,##0"
    style_chart(ch, legend=False)
    ch.Axes(1).ReversePlotOrder = True
    ch.Axes(2).Delete()
    ch.ChartGroups(1).GapWidth = 60

    # Rep leaderboard rows 37-44
    heading(d, "B37", "Sales reps, selected month and region")
    rcols = ["B", "C", "D", "F", "G", "H"]
    for c, t in zip(rcols, ("Sales rep", "Region", "Revenue", "Orders", "Margin %", "Share")):
        h = d.Range(f"{c}38"); h.Value = t
        font(h, 10, True, (255, 255, 255)); h.Interior.Color = rgb(*NAVY)
        if c not in "BC":
            h.HorizontalAlignment = -4152
    d.Range("E38").Interior.Color = rgb(*NAVY)
    for i in range(6):
        r = 39 + i
        # ranked by revenue; reps with no sales in the selection stay blank
        key = f"LARGE(Reps[RankKey],{i + 1})"
        for c, col in zip(rcols, ("Sales Rep", "Region", "Revenue", "Orders", "Margin %", "Share")):
            d.Range(f"{c}{r}").Formula = f'=IF({key}<0.5,"",INDEX(Reps[{col}],MATCH({key},Reps[RankKey],0)))'
        d.Range(f"D{r}").NumberFormat = "#,##0"; d.Range(f"F{r}").NumberFormat = "#,##0"
        d.Range(f"G{r}:H{r}").NumberFormat = "0.0%"
    bars = d.Range("D39:D44").FormatConditions.AddDatabar()
    bars.BarColor.Color = rgb(150, 205, 200)
    bars.MinPoint.Modify(0, 0)  # bars start at zero, so length = revenue


    # Data checks rows 48-53
    heading(d, "B48", "Data checks")
    checks = [("Sales lines loaded = rows in the sales file", '=IF(INDEX(Checks[Value],1)=ROWS(Sales[Date]),"PASS","FAIL")'),
              ("Every month/region/category sold has a budget", '=IF(INDEX(Checks[Value],3)=0,"PASS","FAIL")'),
              ("Budget has no repeated month/region/category", '=IF(INDEX(Checks[Value],5)=0,"PASS","FAIL")'),
              ("12-month chart (all regions) adds up to all revenue",
               '=IF(ROUND(SUMPRODUCT(SUMIFS(Sales[Revenue],Sales[Month],MonthList))-INDEX(Checks[Value],6),2)=0,"PASS","FAIL")'),
              ("Region table total = all sales in the month", '=IF(ROUND(Calc!B31-SUMIFS(Sales[Revenue],Sales[Month],SelMonth),2)=0,"PASS","FAIL")')]
    for i, (label, f) in enumerate(checks, 49):
        d.Range(f"B{i}").Value = label
        cell = d.Range(f"H{i}")
        cell.Formula = f
        font(cell, 10, True); cell.HorizontalAlignment = -4108
        for val, bg, fg in (("PASS", GREEN_BG, GREEN_FG), ("FAIL", RED_BG, RED_FG)):
            fc = cell.FormatConditions.Add(1, 3, f'="{val}"'); fc.Interior.Color = rgb(*bg); fc.Font.Color = rgb(*fg)

    ps = d.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$U$53"
    for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
        setattr(ps, m, 18)


def style_chart(ch, legend=True):
    # headings live in the cells above. Multi-series charts get no title; a one-series chart gets a blank
    # title, otherwise Excel adds the series name as a title when the file is opened.
    ch.HasTitle = not legend
    if not legend:
        ch.ChartTitle.Text = " "
        ch.ChartTitle.Font.Size = 2
    ch.HasLegend = legend
    if legend:
        ch.Legend.Position = -4160  # top
    ch.Axes(2).HasMajorGridlines = True
    ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = rgb(230, 230, 230)
    ch.ChartArea.Format.Line.Visible = False
    ch.ChartArea.Font.Size = 9


def build_howto(wb):
    ws = wb.Worksheets("How to use")
    ws.Columns("A").ColumnWidth = 110
    text = [("How to use this dashboard", True), ("", False),
            ("Every day", True),
            ("1. On Dashboard, pick a month and a region from the two yellow drop-downs. Every number and chart follows.", False),
            ("2. Red = below budget, green = on or above budget. The 12-month chart shows the selected month in dark blue.", False),
            ("", False), ("Every month", True),
            ("1. Export sales lines and the budget as CSV into the data folder (see Settings).", False),
            ("2. Data > Refresh All. New months appear in the month drop-down automatically.", False),
            ("3. Check the Data checks block at the bottom of the Dashboard: all five must say PASS.", False),
            ("", False), ("How it is built", True),
            ("Power Query (Data > Queries & Connections) loads and types the CSVs: Sales, Budget, Customers, Reps, Checks.", False),
            ("Calc holds every formula the Dashboard shows (SUMIFS on the Sales and Budget tables), so any number can be traced.", False),
            ("No macros, no pivot cache to refresh by hand; works in Excel 2016 and Microsoft 365 for Windows.", False)]
    for i, (t, bold) in enumerate(text, 1):
        ws.Cells(i, 1).Value = t
        ws.Cells(i, 1).Font.Bold = bold
    ws.Range("A1").Font.Size = 16


if __name__ == "__main__":
    main()
