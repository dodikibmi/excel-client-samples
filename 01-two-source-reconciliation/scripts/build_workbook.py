"""Builds Two-Source-Reconciliation.xlsx with real Excel (COM): Power Query cleans both CSV exports,
matches them (full outer merge on a normalised reference), and the Summary sheet turns the result into a
bank reconciliation statement whose unexplained difference must be 0.00.

Run with a Python that has pywin32:  python scripts/build_workbook.py
The workbook has no macros; the client only needs Data -> Refresh All.
"""
import time
from pathlib import Path
import win32com.client as win32
from summary_sheet import build_summary

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Two-Source-Reconciliation.xlsx"

# ---------------------------------------------------------------- Power Query (M)
Q = {}
Q["Setting"] = """(name as text) as text =>
    // Reads one value from a named cell on the Settings sheet (DataFolder, BankFile, LedgerFile)
    Text.From(Excel.CurrentWorkbook(){[Name = name]}[Content]{0}[Column1])"""

Q["Bank_Raw"] = """let
    Path    = Setting("DataFolder") & Setting("BankFile"),
    Source  = Csv.Document(File.Contents(Path), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Headers"""

Q["Ledger_Raw"] = """let
    Path    = Setting("DataFolder") & Setting("LedgerFile"),
    Source  = Csv.Document(File.Contents(Path), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Headers"""

Q["Bank_Clean"] = """let
    Source   = Bank_Raw,
    // 1. the bank export sometimes repeats a row: drop exact duplicates
    Distinct = Table.Distinct(Source),
    // 2. US bank: MM/DD/YYYY dates, amounts as text with thousand separators
    ToNumber = (t as nullable text) as number => if t = null or Text.Trim(t) = "" then 0 else Number.FromText(Text.Trim(t), "en-US"),
    Typed    = Table.TransformColumns(Distinct, {
                   {"Date", each Date.FromText(Text.Trim(_), [Format = "MM/dd/yyyy", Culture = "en-US"]), type date},
                   {"Debit", ToNumber, type number}, {"Credit", ToNumber, type number},
                   {"Reference", Text.Trim, type text}, {"Description", Text.Trim, type text}}),
    Amount   = Table.AddColumn(Typed, "Amount", each Number.Round([Credit] - [Debit], 2), type number),
    // 3. one matching key: upper case, no spaces / dashes / slashes / dots
    Key      = Table.AddColumn(Amount, "Key", each
                   let k = Text.Upper(Text.Remove([Reference], {" ", "-", "/", "."})) in if k = "" then null else k, type text),
    Line     = Table.AddIndexColumn(Key, "Bank line", 1, 1, Int64.Type),
    // 4. the same reference twice on the statement is NOT matched twice: occurrence 2+ is reported
    Grouped  = Table.Group(Line, {"Key"}, {{"Rows", each Table.AddIndexColumn(Table.Sort(_, {"Bank line"}), "Occurrence", 1, 1, Int64.Type), type table}}),
    Expand   = Table.ExpandTableColumn(Grouped, "Rows", {"Bank line", "Date", "Reference", "Description", "Amount", "Balance", "Occurrence"}),
    NoKeyOne = Table.ReplaceValue(Expand, each [Occurrence], each if [Key] = null then 1 else [Occurrence], Replacer.ReplaceValue, {"Occurrence"}),
    Sorted   = Table.Sort(NoKeyOne, {"Bank line"}),
    Result   = Table.TransformColumnTypes(Sorted, {{"Bank line", Int64.Type}, {"Date", type date}, {"Amount", type number},
                   {"Reference", type text}, {"Description", type text}, {"Key", type text}, {"Occurrence", Int64.Type}})
in
    Result"""

Q["Bank_Controls"] = """let
    // the opening balance is added on the Summary sheet: mixing Settings and file data in one query trips the privacy firewall
    Raw     = Bank_Clean,
    Printed = Number.FromText(Text.Trim(List.Last(Raw[Balance])), "en-US"),
    Moves   = Number.Round(List.Sum(Raw[Amount]), 2)
in
    #table(type table [Control = text, Value = number], {
        {"Closing balance printed on the statement", Printed},
        {"Movements on the statement (after removing repeated export rows)", Moves}})"""

Q["Ledger_Clean"] = """let
    Source  = Ledger_Raw,
    Typed   = Table.TransformColumns(Source, {
                  {"Date", each Date.FromText(Text.Trim(_), [Format = "yyyy-MM-dd", Culture = "en-US"]), type date},
                  {"Amount", each Number.Round(Number.FromText(Text.Trim(_), "en-US"), 2), type number},
                  {"Reference", Text.Trim, type text}, {"Party", Text.Trim, type text}, {"Doc No", Text.Trim, type text}}),
    Key     = Table.AddColumn(Typed, "Key", each Text.Upper(Text.Remove([Reference], {" ", "-", "/", "."})), type text),
    Line    = Table.AddIndexColumn(Key, "Ledger line", 1, 1, Int64.Type),
    // the same reference posted more than once: occurrence 1 is matched, the rest are reported as duplicates
    Grouped = Table.Group(Line, {"Key"}, {{"Rows", each Table.AddIndexColumn(Table.Sort(_, {"Ledger line"}), "Occurrence", 1, 1, Int64.Type), type table}}),
    Expand  = Table.ExpandTableColumn(Grouped, "Rows", {"Ledger line", "Date", "Doc No", "Party", "Reference", "Amount", "Memo", "Occurrence"}),
    Sorted  = Table.Sort(Expand, {"Ledger line"}),
    Result  = Table.TransformColumnTypes(Sorted, {{"Ledger line", Int64.Type}, {"Date", type date}, {"Amount", type number}, {"Occurrence", Int64.Type}})
in
    Result"""

Q["Recon"] = """let
    Bank     = Table.SelectRows(Bank_Clean, each [Occurrence] = 1),
    BankRep  = Table.SelectRows(Bank_Clean, each [Occurrence] > 1),
    Ledger   = Table.SelectRows(Ledger_Clean, each [Occurrence] = 1),
    Repeats  = Table.SelectRows(Ledger_Clean, each [Occurrence] > 1),
    Joined   = Table.NestedJoin(Bank, {"Key"}, Ledger, {"Key"}, "L", JoinKind.FullOuter),
    Expanded = Table.ExpandTableColumn(Joined, "L", {"Ledger line", "Date", "Doc No", "Party", "Amount"},
                   {"Ledger line", "Ledger date", "Doc no", "Party", "Ledger amount"}),
    Renamed  = Table.RenameColumns(Expanded, {{"Date", "Bank date"}, {"Reference", "Bank reference"},
                   {"Description", "Bank description"}, {"Amount", "Bank amount"}}),
    DupRows  = Table.RenameColumns(Table.SelectColumns(Repeats, {"Ledger line", "Date", "Doc No", "Party", "Amount", "Key"}),
                   {{"Date", "Ledger date"}, {"Doc No", "Doc no"}, {"Amount", "Ledger amount"}}),
    DupFlag  = Table.AddColumn(DupRows, "IsDuplicate", each true, type logical),
    BankDup  = Table.AddColumn(Table.RenameColumns(Table.SelectColumns(BankRep, {"Bank line", "Date", "Reference", "Description", "Amount", "Key"}),
                   {{"Date", "Bank date"}, {"Reference", "Bank reference"}, {"Description", "Bank description"}, {"Amount", "Bank amount"}}),
                   "IsBankDuplicate", each true, type logical),
    All      = Table.Combine({Table.AddColumn(Renamed, "IsDuplicate", each false, type logical), DupFlag, BankDup}),
    Status   = Table.AddColumn(All, "Status", each
                   if [IsBankDuplicate] = true then "Duplicate in bank"
                   else if [IsDuplicate] = true then "Duplicate in ledger"
                   else if [Ledger line] = null then "Only in bank"
                   else if [Bank line] = null then "Only in ledger"
                   else if Number.Round([Bank amount] - [Ledger amount], 2) = 0 then "Matched"
                   else "Amount differs", type text),
    Diff     = Table.AddColumn(Status, "Difference", each
                   if [Status] = "Amount differs" then Number.Round([Bank amount] - [Ledger amount], 2) else null, type number),
    Gap      = Table.AddColumn(Diff, "Days to clear", each
                   if [Bank date] <> null and [Ledger date] <> null then Duration.Days([Bank date] - [Ledger date]) else null, Int64.Type),
    Action   = Table.AddColumn(Gap, "What to do", each
                   let d = Text.Upper(if [Bank description] = null then "" else [Bank description]) in
                   if [Status] = "Matched" then (if [Days to clear] > 0 then "OK - cleared after " & Text.From([Days to clear]) & " day(s)" else "OK")
                   else if [Status] = "Amount differs" then "Check with " & [Party] & ": bank is " & Number.ToText(Number.Abs([Difference]), "N2", "en-US")
                        & (if [Difference] < 0 then " lower" else " higher") & " than the cash book"
                   else if [Status] = "Duplicate in ledger" then "Posted twice - reverse " & [Doc no]
                   else if [Status] = "Duplicate in bank" then "Same reference twice on the statement - check for a double payment, then post or reclaim it"
                   else if [Status] = "Only in ledger" then (if [Ledger amount] > 0 then "Deposit in transit - should clear early next month"
                                                             else "Outstanding payment - confirm it clears next month")
                   else if Text.Contains(d, "FEE") then "Bank charge - post it to the cash book"
                   else if Text.Contains(d, "INTEREST") then "Interest - post it to the cash book"
                   else if [Key] <> null then "No cash book entry for " & [Bank reference] & " - find the bill and post it"
                   else "Unidentified receipt - ask the bank who paid it", type text),
    Order    = Table.AddColumn(Action, "Sort", each
                   List.PositionOf({"Amount differs", "Duplicate in ledger", "Duplicate in bank", "Only in bank", "Only in ledger", "Matched"}, [Status]), Int64.Type),
    Sorted   = Table.Sort(Order, {{"Sort", Order.Ascending}, {"Bank date", Order.Ascending}, {"Ledger date", Order.Ascending}}),
    Result   = Table.SelectColumns(Sorted, {"Status", "Bank line", "Bank date", "Bank reference", "Bank description", "Bank amount",
                   "Ledger line", "Ledger date", "Doc no", "Party", "Ledger amount", "Difference", "Days to clear", "What to do"})
in
    Result"""

Q["Discrepancies"] = """let
    Source = Table.SelectRows(Recon, each [Status] <> "Matched")
in
    Table.RemoveColumns(Source, {"Days to clear"})"""

Q["Cleaning_Log"] = """let
    BankRows  = Table.RowCount(Bank_Raw),
    BankKept  = Table.RowCount(Bank_Clean),
    Standard  = (r as text) => r = Text.Upper(r) and not Text.Contains(r, " ") and Text.Contains(r, "-"),
    RefsFixed = Table.RowCount(Table.SelectRows(Bank_Clean, each [Key] <> null and not Standard([Reference]))),
    BankReps  = Table.RowCount(Table.SelectRows(Bank_Clean, each [Occurrence] > 1)),
    NoRef     = Table.RowCount(Table.SelectRows(Bank_Clean, each [Key] = null)),
    LedRows   = Table.RowCount(Ledger_Raw),
    Repeats   = Table.RowCount(Table.SelectRows(Ledger_Clean, each [Occurrence] > 1)),
    Log = #table(type table [Source = text, Check = text, Rows = Int64.Type, #"What was done" = text], {
        {"Bank statement", "Rows in the export file", BankRows, "Read as text first, so nothing is lost on import"},
        {"Bank statement", "Exact duplicate rows", BankRows - BankKept, "Removed (the export repeated them)"},
        {"Bank statement", "Dates in US format MM/DD/YYYY", BankKept, "Converted to real dates"},
        {"Bank statement", "Amounts stored as text (""1,250.00"")", BankKept, "Converted to numbers; Debit/Credit merged into one signed Amount"},
        {"Bank statement", "References typed differently (case, spaces, dashes)", RefsFixed, "Normalised into one matching key"},
        {"Bank statement", "Rows without a reference", NoRef, "Kept; they can only appear as Only in bank"},
        {"Bank statement", "Same reference more than once (not identical rows)", BankReps, "First one matched; the repeat is reported as Duplicate in bank"},
        {"Cash book", "Rows in the export file", LedRows, "Read as text first, so nothing is lost on import"},
        {"Cash book", "Same reference posted more than once", Repeats, "First posting matched; the repeat is reported as Duplicate in ledger"}})
in
    Log"""

LOADED = [("Summary", None), ("Discrepancies", "Discrepancies"), ("Reconciliation", "Recon"),
          ("Cleaning log", "Cleaning_Log"), ("Settings", None), ("How to use", None)]

NAVY, GREY, YELLOW = 0x5A3A1F, 0xF2F2F2, 0x99FFFF  # Excel colours are BGR


def rgb(r, g, b):
    return r + g * 256 + b * 65536


def load_query(wb, ws, name, table_name, at="A1"):
    conn = ("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;"
            f"Location={name};Extended Properties=\"\"")
    lo = ws.ListObjects.Add(0, conn, None, 1, ws.Range(at))  # xlSrcExternal, xlYes
    lo.QueryTable.CommandType = 2  # xlCmdSql
    lo.QueryTable.CommandText = f"SELECT * FROM [{name}]"
    lo.QueryTable.RefreshStyle = 1  # xlInsertDeleteCells
    lo.QueryTable.AdjustColumnWidth = False
    lo.QueryTable.Refresh(False)
    lo.Name = table_name
    lo.TableStyle = "TableStyleMedium2"
    return lo


def main():
    if OUT.exists():
        OUT.unlink()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < len(LOADED):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, (title, _) in enumerate(LOADED, 1):
            wb.Worksheets(i).Name = title

        # Settings first: the queries read their file locations from here
        st = wb.Worksheets("Settings")
        st.Range("A1").Value = "Settings"
        st.Range("A1").Font.Size = 16; st.Range("A1").Font.Bold = True
        rows = [("Data folder", '=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"data\\"', "DataFolder",
                 "Folder with the two CSV exports. Default: the data folder next to this workbook. Type a path to override."),
                ("Bank statement file", "bank-statement-2026-10.csv", "BankFile", "File name of the bank export for the month"),
                ("Cash book file", "cash-book-2026-10.csv", "LedgerFile", "File name of the cash book export for the month"),
                ("Opening balance (bank)", 18450.00, "BankOpening", "Bank balance on the first day of the month"),
                ("Opening balance (cash book)", 18450.00, "LedgerOpening", "Cash book balance on the first day of the month"),
                ("Period", "October 2026", "Period", "Shown in the report title")]
        st.Range("A3:C3").Value = ("Setting", "Value", "Note")
        st.Range("A3:C3").Font.Bold = True
        for r, (label, val, nm, note) in enumerate(rows, 4):
            st.Cells(r, 1).Value = label
            if isinstance(val, str) and val.startswith("="):
                st.Cells(r, 2).Formula2 = val  # Formula2: no implicit-intersection @ (keeps Excel 2016 happy)
            else:
                if isinstance(val, str):
                    st.Cells(r, 2).NumberFormat = "@"  # keep "October 2026" as text, not a date
                st.Cells(r, 2).Value = val
            st.Cells(r, 3).Value = note
            st.Cells(r, 2).Interior.Color = YELLOW
            wb.Names.Add(nm, f"=Settings!$B${r}")
        st.Range("B7:B8").NumberFormat = "#,##0.00"
        st.Columns("A").ColumnWidth = 28; st.Columns("B").ColumnWidth = 48; st.Columns("C").ColumnWidth = 70

        wb.SaveAs(str(OUT), 51)  # xlOpenXMLWorkbook; CELL("filename") needs a saved file
        xl.Calculate()

        for name, formula in Q.items():
            wb.Queries.Add(name, formula)

        t0 = time.time()
        tables = {}
        for title, query in LOADED:
            if query:
                tables[query] = load_query(wb, wb.Worksheets(title), query, query.replace("_", ""))
        load_query(wb, wb.Worksheets("Cleaning log"), "Bank_Controls", "BankControls", "F1")
        print(f"queries loaded in {time.time() - t0:.1f}s")

        format_tables(wb)
        build_summary(wb)
        build_howto(wb)
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            ws.Range("A1").Select()
        wb.Worksheets("Summary").Activate()
        wb.Save()
        print("saved", OUT.name)
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        xl.Quit()


def format_tables(wb):
    for title, cols_money, cols_date in [("Reconciliation", ["Bank amount", "Ledger amount", "Difference"], ["Bank date", "Ledger date"]),
                                         ("Discrepancies", ["Bank amount", "Ledger amount", "Difference"], ["Bank date", "Ledger date"])]:
        ws = wb.Worksheets(title)
        lo = ws.ListObjects(1)
        for c in cols_money:
            lo.ListColumns(c).DataBodyRange.NumberFormat = "#,##0.00;[Red]-#,##0.00"
        for c in cols_date:
            lo.ListColumns(c).DataBodyRange.NumberFormat = "dd mmm yyyy"
        widths = {"Status": 20, "Bank line": 6, "Bank date": 12, "Bank reference": 13, "Bank description": 36,
                  "Bank amount": 12, "Ledger line": 7, "Ledger date": 12, "Doc no": 9, "Party": 24,
                  "Ledger amount": 13, "Difference": 11, "Days to clear": 7, "What to do": 58}
        for i in range(1, lo.ListColumns.Count + 1):
            ws.Columns(i).ColumnWidth = widths.get(lo.ListColumns(i).Name, 12)
        lo.HeaderRowRange.WrapText = True
        rng = lo.ListColumns("Status").DataBodyRange
        for text, color in [("Matched", rgb(226, 239, 218)), ("Amount differs", rgb(252, 228, 214)),
                            ("Duplicate in ledger", rgb(252, 228, 214)), ("Only in bank", rgb(255, 242, 204)),
                            ("Only in ledger", rgb(221, 235, 247))]:
            fc = rng.FormatConditions.Add(2, None, f'=$A2="{text}"')  # xlExpression
            fc.Interior.Color = color
        ws.Activate()
        ws.Range("A2").Select()
        wb.Application.ActiveWindow.FreezePanes = True
    ws = wb.Worksheets("Cleaning log")
    ws.Columns("A").ColumnWidth = 16; ws.Columns("B").ColumnWidth = 50; ws.Columns("C").ColumnWidth = 8; ws.Columns("D").ColumnWidth = 70
    ws.Columns("E").ColumnWidth = 3; ws.Columns("F").ColumnWidth = 62; ws.Columns("G").ColumnWidth = 14
    ws.ListObjects("BankControls").ListColumns("Value").DataBodyRange.NumberFormat = "#,##0.00"


def build_howto(wb):
    ws = wb.Worksheets("How to use")
    ws.Columns("A").ColumnWidth = 110
    text = [
        ("How to use this workbook", True),
        ("", False),
        ("Every month", True),
        ("1. Export the bank statement and the cash book as CSV into the data folder (see Settings).", False),
        ("2. On Settings, type the two file names and the opening balances.", False),
        ("3. Data > Refresh All. All tabs update; no formulas to copy, no macros.", False),
        ("4. Summary: the unexplained difference must be 0.00. Work through the Discrepancies tab.", False),
        ("", False),
        ("What each status means", True),
        ("Matched: same reference and amount on both sides (Days to clear shows timing differences).", False),
        ("Amount differs: same reference, different amount. Usually a short payment, a fee deducted, or a typo.", False),
        ("Duplicate in ledger: the same reference was posted more than once in the cash book.", False),
        ("Only in bank: bank charges, interest, direct debits or receipts nobody has booked yet.", False),
        ("Only in ledger: deposits in transit and payments that have not cleared yet.", False),
        ("", False),
        ("How the match works", True),
        ("References are cleaned first (upper case, no spaces or dashes), so 'inv-1023 ', 'INV1023' and 'INV-1023' are the same key.", False),
        ("Bank rows exported twice are removed; every cleaning step is counted on the Cleaning log tab.", False),
        ("The steps live in Power Query (Data > Queries & Connections): Bank_Clean, Ledger_Clean, Recon.", False),
    ]
    for i, (t, bold) in enumerate(text, 1):
        ws.Cells(i, 1).Value = t
        ws.Cells(i, 1).Font.Bold = bold
    ws.Range("A1").Font.Size = 16


if __name__ == "__main__":
    main()
