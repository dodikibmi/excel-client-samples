"""Builds Two-Source-Reconciliation.xlsx with real Excel (COM).

Power Query reads every bank and cash book CSV in the data folder (two accounts, six months), cleans them, and matches
them with three rules (reference; reference with an amount tolerance; amount + date for lines without a usable
reference). The Summary sheet turns the result into a bank reconciliation statement for the account and month picked,
whose unexplained difference must be 0.00. The Exceptions workbench keeps Owner / Comment / Resolved date typed by the
user across refreshes (a self-referencing query).

Run with a Python that has pywin32:  python scripts/build_workbook.py
The workbook has no macros; the client only needs Data -> Refresh All.
"""
import json, time
from datetime import datetime
from pathlib import Path
import win32com.client as win32
import lux
from lux import c, fill, font
from summary_sheet import build_summary, build_calc

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Two-Source-Reconciliation.xlsx"
NOTES = json.loads((ROOT / "scripts" / "demo_notes.json").read_text(encoding="utf-8"))

# ---------------------------------------------------------------- Power Query (M)
Q = {}
Q["Setting"] = """(name as text) as text =>
    // Reads one value from a named cell on the Settings sheet (DataFolder)
    Text.From(Excel.CurrentWorkbook(){[Name = name]}[Content]{0}[Column1])"""

Q["SettingNum"] = """(name as text) as number =>
    // Reads one number from a named cell on the Settings sheet (AmountTolerance, DaysTolerance)
    Number.From(Excel.CurrentWorkbook(){[Name = name]}[Content]{0}[Column1])"""

Q["Files"] = """let
    // every CSV in the data folder named bank_<Account>_<yyyy-mm>.csv or cashbook_<Account>_<yyyy-mm>.csv
    Source = Folder.Files(Setting("DataFolder")),
    Csv    = Table.SelectRows(Source, each Text.Lower([Extension]) = ".csv"),
    Parts  = Table.AddColumn(Csv, "Parts", each Text.Split(Text.BeforeDelimiter([Name], ".", {0, RelativePosition.FromEnd}), "_")),
    Valid  = Table.SelectRows(Parts, each List.Count([Parts]) = 3 and List.Contains({"bank", "cashbook"}, Text.Lower([Parts]{0}))),
    Kind   = Table.AddColumn(Valid, "Kind", each Text.Lower([Parts]{0}), type text),
    Acct   = Table.AddColumn(Kind, "Account", each [Parts]{1}, type text),
    Period = Table.AddColumn(Acct, "Period", each [Parts]{2}, type text),
    Result = Table.Sort(Table.SelectColumns(Period, {"Name", "Kind", "Account", "Period", "Content"}), {{"Name", Order.Ascending}})
in
    Result"""

READ = """let
    Source = Table.SelectRows(Files, each [Kind] = "{kind}"),
    // read every file as text first, so nothing is lost on import; Row = line number inside the file
    Parsed = Table.AddColumn(Source, "Rows", each Table.AddIndexColumn(
                 Table.PromoteHeaders(Csv.Document([Content], [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
                 [PromoteAllScalars = true]), "Row", 1, 1, Int64.Type)),
    Expand = Table.ExpandTableColumn(Table.SelectColumns(Parsed, {{"Account", "Period", "Rows"}}), "Rows", {cols})
in
    Expand"""
Q["Bank_Raw"] = READ.format(kind="bank", cols='{"Row", "Date", "Reference", "Description", "Debit", "Credit", "Balance"}')
Q["Ledger_Raw"] = READ.format(kind="cashbook", cols='{"Row", "Date", "Doc No", "Party", "Reference", "Amount", "Memo"}')

Q["Bank_Clean"] = """let
    Source   = Bank_Raw,
    // 1. the bank export sometimes repeats a row: drop exact duplicates (same account, every field equal)
    Distinct = Table.Distinct(Source, {"Account", "Date", "Reference", "Description", "Debit", "Credit", "Balance"}),
    // 2. US bank: MM/DD/YYYY dates, amounts as text with thousand separators
    ToNumber = (t as nullable text) as number => if t = null or Text.Trim(t) = "" then 0 else Number.FromText(Text.Trim(t), "en-US"),
    Typed    = Table.TransformColumns(Distinct, {
                   {"Date", each Date.FromText(Text.Trim(_), [Format = "MM/dd/yyyy", Culture = "en-US"]), type date},
                   {"Debit", ToNumber, type number}, {"Credit", ToNumber, type number}, {"Balance", ToNumber, type number},
                   {"Reference", Text.Trim, type text}, {"Description", Text.Trim, type text}}),
    Amount   = Table.AddColumn(Typed, "Amount", each Number.Round([Credit] - [Debit], 2), type number),
    // 3. one matching key: upper case, no spaces / dashes / slashes / dots; empty reference = no key
    Key      = Table.AddColumn(Amount, "Key", each
                   let k = Text.Upper(Text.Remove([Reference], {" ", "-", "/", "."})) in if k = "" then null else k, type text),
    Id       = Table.AddColumn(Key, "Bank ID", each Text.Upper(Text.Start([Account], 3)) & "-B-" & Text.Remove([Period], "-")
                   & "-" & Text.PadStart(Text.From([Row]), 3, "0"), type text),
    // 4. the same reference twice on the statement is NOT matched twice: occurrence 2+ is reported
    Grouped  = Table.Group(Id, {"Account", "Key"}, {{"Rows", each Table.AddIndexColumn(Table.Sort(_, {"Bank ID"}), "Occurrence", 1, 1, Int64.Type), type table}}),
    Expand   = Table.ExpandTableColumn(Grouped, "Rows", {"Bank ID", "Period", "Row", "Date", "Reference", "Description", "Amount", "Balance", "Occurrence"}),
    NoKeyOne = Table.ReplaceValue(Expand, each [Occurrence], each if [Key] = null then 1 else [Occurrence], Replacer.ReplaceValue, {"Occurrence"}),
    Sorted   = Table.Sort(NoKeyOne, {"Bank ID"}),
    Result   = Table.TransformColumnTypes(Sorted, {{"Row", Int64.Type}, {"Date", type date}, {"Amount", type number}, {"Balance", type number},
                   {"Reference", type text}, {"Description", type text}, {"Key", type text}, {"Occurrence", Int64.Type}})
in
    Result"""

Q["Ledger_Clean"] = """let
    Source  = Ledger_Raw,
    Typed   = Table.TransformColumns(Source, {
                  {"Date", each Date.FromText(Text.Trim(_), [Format = "yyyy-MM-dd", Culture = "en-US"]), type date},
                  {"Amount", each Number.Round(Number.FromText(Text.Trim(_), "en-US"), 2), type number},
                  {"Reference", Text.Trim, type text}, {"Party", Text.Trim, type text}, {"Doc No", Text.Trim, type text}}),
    Key     = Table.AddColumn(Typed, "Key", each
                  let k = Text.Upper(Text.Remove([Reference], {" ", "-", "/", "."})) in if k = "" then null else k, type text),
    Id      = Table.AddColumn(Key, "Ledger ID", each Text.Upper(Text.Start([Account], 3)) & "-L-" & Text.Remove([Period], "-")
                  & "-" & Text.PadStart(Text.From([Row]), 3, "0"), type text),
    // the same reference posted more than once: occurrence 1 is matched, the rest are reported as duplicates
    Grouped = Table.Group(Id, {"Account", "Key"}, {{"Rows", each Table.AddIndexColumn(Table.Sort(_, {"Ledger ID"}), "Occurrence", 1, 1, Int64.Type), type table}}),
    Expand  = Table.ExpandTableColumn(Grouped, "Rows", {"Ledger ID", "Period", "Row", "Date", "Doc No", "Party", "Reference", "Amount", "Memo", "Occurrence"}),
    NoKey   = Table.ReplaceValue(Expand, each [Occurrence], each if [Key] = null then 1 else [Occurrence], Replacer.ReplaceValue, {"Occurrence"}),
    Sorted  = Table.Sort(NoKey, {"Ledger ID"}),
    Result  = Table.TransformColumnTypes(Sorted, {{"Row", Int64.Type}, {"Date", type date}, {"Amount", type number}, {"Occurrence", Int64.Type}})
in
    Result"""

Q["Recon"] = """let
    Tol      = SettingNum("AmountTolerance"),
    DaysTol  = SettingNum("DaysTolerance"),
    AsAt     = Date.EndOfMonth(List.Max(Bank_Clean[Date])),
    B        = Table.RenameColumns(Table.SelectColumns(Bank_Clean, {"Account", "Bank ID", "Date", "Reference", "Description", "Amount", "Key", "Occurrence"}),
                   {{"Date", "Bank date"}, {"Reference", "Bank reference"}, {"Description", "Bank description"}, {"Amount", "Bank amount"}}),
    L        = Table.RenameColumns(Table.SelectColumns(Ledger_Clean, {"Account", "Ledger ID", "Date", "Doc No", "Party", "Amount", "Key", "Occurrence"}),
                   {{"Account", "L account"}, {"Date", "Ledger date"}, {"Doc No", "Doc no"}, {"Amount", "Ledger amount"}, {"Key", "L key"}, {"Occurrence", "L occ"}}),
    // rule 1 and 2: same account and same reference key
    Pairs    = Table.Join(Table.SelectRows(B, each [Occurrence] = 1 and [Key] <> null), {"Account", "Key"},
                          Table.SelectRows(L, each [L occ] = 1 and [L key] <> null), {"L account", "L key"}, JoinKind.Inner),
    KeyRule  = Table.AddColumn(Pairs, "SR", each let d = Number.Round([Bank amount] - [Ledger amount], 2) in
                   if d = 0 then {"Matched", "Reference"} else if Number.Abs(d) <= Tol then {"Matched", "Reference ± amount"} else {"Amount differs", ""}),
    PairedB  = List.Buffer(Pairs[Bank ID]),
    PairedL  = List.Buffer(Pairs[Ledger ID]),
    BankLeft = Table.SelectRows(B, each [Occurrence] = 1 and not List.Contains(PairedB, [Bank ID])),
    LedLeft  = Table.SelectRows(L, each [L occ] = 1 and not List.Contains(PairedL, [Ledger ID])),
    // rule 3, leftovers only: same account, amount within the tolerance, dates within the days tolerance; closest first, each line once
    Cands0   = Table.Join(BankLeft, {"Account"}, LedLeft, {"L account"}, JoinKind.Inner),
    Cands1   = Table.AddColumn(Cands0, "AbsDiff", each Number.Abs(Number.Round([Bank amount] - [Ledger amount], 2)), type number),
    Cands2   = Table.AddColumn(Cands1, "AbsDays", each Number.Abs(Duration.Days([Bank date] - [Ledger date])), type number),
    Cands    = Table.Buffer(Table.Sort(Table.SelectRows(Cands2, each [AbsDiff] <= Tol and [AbsDays] <= DaysTol),
                   {{"AbsDiff", Order.Ascending}, {"AbsDays", Order.Ascending}, {"Bank ID", Order.Ascending}, {"Ledger ID", Order.Ascending}})),
    Picked   = List.Accumulate(Table.ToRecords(Cands), [b = {}, l = {}, p = {}], (s, r) =>
                   if List.Contains(s[b], r[Bank ID]) or List.Contains(s[l], r[Ledger ID]) then s
                   else [b = s[b] & {r[Bank ID]}, l = s[l] & {r[Ledger ID]}, p = s[p] & {r[Bank ID] & "|" & r[Ledger ID]}]),
    Rule3    = Table.SelectRows(Cands, each List.Contains(Picked[p], [Bank ID] & "|" & [Ledger ID])),
    KeyT     = Table.AddColumn(Table.AddColumn(KeyRule, "Status", each [SR]{0}, type text), "Match rule", each [SR]{1}, type text),
    R3T      = Table.AddColumn(Table.AddColumn(Rule3, "Status", each "Matched", type text), "Match rule", each "Amount + date", type text),
    BOnly    = Table.AddColumn(Table.SelectRows(BankLeft, each not List.Contains(Picked[b], [Bank ID])), "Status", each "Only in bank", type text),
    LOnly    = Table.AddColumn(Table.SelectRows(LedLeft, each not List.Contains(Picked[l], [Ledger ID])), "Status", each "Only in ledger", type text),
    BDup     = Table.AddColumn(Table.SelectRows(B, each [Occurrence] > 1), "Status", each "Duplicate in bank", type text),
    LDup     = Table.AddColumn(Table.SelectRows(L, each [L occ] > 1), "Status", each "Duplicate in ledger", type text),
    All      = Table.Combine({KeyT, R3T, BOnly, LOnly, BDup, LDup}),
    Acct     = Table.AddColumn(All, "Acct", each if [Account] <> null then [Account] else [L account], type text),
    Diff     = Table.AddColumn(Acct, "Difference", each
                   if [Bank amount] <> null and [Ledger amount] <> null and Number.Round([Bank amount] - [Ledger amount], 2) <> 0
                   then Number.Round([Bank amount] - [Ledger amount], 2) else null, type number),
    Gap      = Table.AddColumn(Diff, "Days to clear", each
                   if [Bank date] <> null and [Ledger date] <> null then Duration.Days([Bank date] - [Ledger date]) else null, Int64.Type),
    Month    = Table.AddColumn(Gap, "Month", each Date.StartOfMonth(if [Bank date] <> null then [Bank date] else [Ledger date]), type date),
    Value    = Table.AddColumn(Month, "Value", each Number.Abs(if [Bank amount] <> null then [Bank amount] else [Ledger amount]), type number),
    ExcId    = Table.AddColumn(Value, "Exception ID", each if [Bank ID] <> null then [Bank ID] else [Ledger ID], type text),
    Action   = Table.AddColumn(ExcId, "What to do", each
                   let d = Text.Upper(if [Bank description] = null then "" else [Bank description]) in
                   if [Status] = "Matched" then
                       (if [Match rule] = "Amount + date" then "OK - matched on amount and date (no usable reference)"
                        else if [Match rule] = "Reference ± amount" then "OK - within tolerance; write off " & Number.ToText(Number.Abs([Difference]), "N2", "en-US")
                        else if [Days to clear] > 0 then "OK - cleared after " & Text.From([Days to clear]) & " day(s)" else "OK")
                   else if [Status] = "Amount differs" then "Check with " & [Party] & ": bank is " & Number.ToText(Number.Abs([Difference]), "N2", "en-US")
                        & (if [Difference] < 0 then " lower" else " higher") & " than the cash book"
                   else if [Status] = "Duplicate in ledger" then "Posted twice - reverse " & [Doc no]
                   else if [Status] = "Duplicate in bank" then "Same reference twice on the statement - check for a double payment, then post or reclaim it"
                   else if [Status] = "Only in ledger" then
                       (if Duration.Days(AsAt - [Ledger date]) > 60 then "Stale - not cleared after 60+ days: investigate, then write back"
                        else if [Ledger amount] > 0 then "Deposit in transit - should clear early next month"
                        else "Outstanding payment - confirm it clears next month")
                   else if Text.Contains(d, "FEE") then "Bank charge - post it to the cash book"
                   else if Text.Contains(d, "INTEREST") then "Interest - post it to the cash book"
                   else if Text.Contains(d, "DIRECT DEBIT") then "Direct debit not in the cash book - post it"
                   else if [Key] <> null then "No cash book entry for " & [Bank reference] & " - find the bill and post it"
                   else if [Bank amount] > 0 then "Unidentified receipt - ask the bank who paid it"
                   else "Unidentified payment - ask the bank", type text),
    Order    = Table.AddColumn(Action, "Sort", each
                   List.PositionOf({"Amount differs", "Duplicate in ledger", "Duplicate in bank", "Only in bank", "Only in ledger", "Matched"}, [Status]), Int64.Type),
    Sorted   = Table.Sort(Order, {{"Acct", Order.Ascending}, {"Sort", Order.Ascending}, {"Month", Order.Descending}, {"Exception ID", Order.Ascending}}),
    Clean    = Table.RenameColumns(Table.RemoveColumns(Sorted, {"Account"}), {{"Acct", "Account"}}),
    Result   = Table.SelectColumns(Clean, {"Account", "Status", "Match rule", "Month", "Bank ID", "Bank date", "Bank reference", "Bank description",
                   "Bank amount", "Ledger ID", "Ledger date", "Doc no", "Party", "Ledger amount", "Difference", "Days to clear", "Value",
                   "Exception ID", "What to do"})
in
    Result"""

Q["Exceptions_Workbench"] = """let
    // every line that is not matched, plus what people typed on the workbench (kept across refreshes by Exception ID)
    AsAt    = Date.EndOfMonth(List.Max(Bank_Clean[Date])),
    Fresh   = Table.SelectRows(Recon, each [Status] <> "Matched"),
    Dated   = Table.AddColumn(Fresh, "Date", each if [Bank date] <> null then [Bank date] else [Ledger date], type date),
    Ref     = Table.AddColumn(Dated, "Reference", each if [Bank reference] <> null and [Bank reference] <> "" then [Bank reference] else [Doc no], type text),
    Desc    = Table.AddColumn(Ref, "Description", each if [Bank description] <> null then [Bank description] else [Party], type text),
    Amount  = Table.AddColumn(Desc, "Amount", each if [Status] = "Amount differs" then [Difference]
                  else if [Bank amount] <> null then [Bank amount] else [Ledger amount], type number),
    Abs     = Table.AddColumn(Amount, "Abs amount", each Number.Abs([Amount]), type number),
    Age     = Table.AddColumn(Abs, "Age (days)", each Duration.Days(AsAt - [Date]), Int64.Type),
    Empty   = #table(type table [#"Exception ID" = text, Owner = text, Comment = text, #"Resolved date" = any], {}),
    Prev    = try Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "Workbench"]}[Content], {"Exception ID", "Owner", "Comment", "Resolved date"}) otherwise Empty,
    Kept    = Table.Buffer(Table.SelectRows(Prev, each [Exception ID] <> null and [Exception ID] <> "")),
    Joined  = Table.NestedJoin(Age, {"Exception ID"}, Kept, {"Exception ID"}, "N", JoinKind.LeftOuter),
    Notes   = Table.ExpandTableColumn(Joined, "N", {"Owner", "Comment", "Resolved date"}),
    Typed   = Table.TransformColumns(Notes, {
                  {"Owner", each if _ = null or _ = "" then null else Text.From(_), type text},
                  {"Comment", each if _ = null or _ = "" then null else Text.From(_), type text},
                  {"Resolved date", each if _ = null or _ = "" then null else try Date.From(_) otherwise null, type date}}),
    State   = Table.AddColumn(Typed, "State", each if [Resolved date] <> null then "Resolved" else "Open", type text),
    // 1 = resolved (green), 2 = open (amber), 3 = open for more than 30 days (red)
    Icons   = Table.AddColumn(State, "Flag", each if [Resolved date] <> null then 1 else if [#"Age (days)"] > 30 then 3 else 2, Int64.Type),
    Sorted  = Table.Sort(Icons, {{"State", Order.Ascending}, {"Account", Order.Ascending}, {"Age (days)", Order.Descending}, {"Exception ID", Order.Ascending}}),
    Result  = Table.SelectColumns(Sorted, {"Flag", "Exception ID", "Account", "Status", "Date", "Age (days)", "Reference", "Description",
                  "Amount", "What to do", "Owner", "Comment", "Resolved date", "State", "Abs amount"})
in
    Result"""

Q["Bank_Controls"] = """let
    // per statement file: does opening + movements = the closing balance printed, and does each month start where the last one closed?
    Per    = Table.Group(Bank_Clean, {"Account", "Period"}, {
                 {"Lines", each Table.RowCount(_), Int64.Type},
                 {"Movements", each Number.Round(List.Sum([Amount]), 2), type number},
                 {"Printed opening", each let t = Table.Sort(_, {"Bank ID"}) in Number.Round(t{0}[Balance] - t{0}[Amount], 2), type number},
                 {"Printed closing", each let t = Table.Sort(_, {"Bank ID"}) in t{Table.RowCount(t) - 1}[Balance], type number}}),
    Sorted = Table.Buffer(Table.Sort(Per, {{"Account", Order.Ascending}, {"Period", Order.Ascending}})),
    Idx    = Table.AddIndexColumn(Sorted, "i", 0, 1, Int64.Type),
    Inside = Table.AddColumn(Idx, "Inside the file", each
                 if Number.Round([Printed opening] + [Movements] - [Printed closing], 2) = 0 then "OK" else "BREAK", type text),
    Chain  = Table.AddColumn(Inside, "Chain", each
                 if [i] = 0 or Sorted{[i] - 1}[Account] <> [Account] then "First file"
                 else if Number.Round(Sorted{[i] - 1}[Printed closing] - [Printed opening], 2) = 0 then "OK" else "BREAK", type text),
    Month  = Table.AddColumn(Chain, "Month", each Date.FromText([Period] & "-01", [Format = "yyyy-MM-dd", Culture = "en-US"]), type date),
    Result = Table.SelectColumns(Month, {"Account", "Month", "Lines", "Printed opening", "Movements", "Printed closing", "Inside the file", "Chain"})
in
    Result"""

Q["Cleaning_Log"] = """let
    BankRows  = Table.RowCount(Bank_Raw),
    BankKept  = Table.RowCount(Bank_Clean),
    Standard  = (r as text) => r = Text.Upper(r) and not Text.Contains(r, " ") and Text.Contains(r, "-"),
    RefsFixed = Table.RowCount(Table.SelectRows(Bank_Clean, each [Key] <> null and not Standard([Reference]))),
    BankReps  = Table.RowCount(Table.SelectRows(Bank_Clean, each [Occurrence] > 1)),
    NoRef     = Table.RowCount(Table.SelectRows(Bank_Clean, each [Key] = null)),
    LedRows   = Table.RowCount(Ledger_Raw),
    LedNoRef  = Table.RowCount(Table.SelectRows(Ledger_Clean, each [Key] = null)),
    Repeats   = Table.RowCount(Table.SelectRows(Ledger_Clean, each [Occurrence] > 1)),
    BankFiles = Table.RowCount(Table.SelectRows(Files, each [Kind] = "bank")),
    LedFiles  = Table.RowCount(Table.SelectRows(Files, each [Kind] = "cashbook")),
    Log = #table(type table [Source = text, Check = text, Rows = Int64.Type, #"What was done" = text], {
        {"Bank statement", "Files read (one per account and month)", BankFiles, "bank_<Account>_<yyyy-mm>.csv, all read the same way"},
        {"Bank statement", "Rows in the export files", BankRows, "Read as text first, so nothing is lost on import"},
        {"Bank statement", "Exact duplicate rows", BankRows - BankKept, "Removed (the export repeated them)"},
        {"Bank statement", "Dates in US format MM/DD/YYYY", BankKept, "Converted to real dates"},
        {"Bank statement", "Amounts stored as text (""1,250.00"")", BankKept, "Converted to numbers; Debit/Credit merged into one signed Amount"},
        {"Bank statement", "References typed differently (case, spaces, dashes)", RefsFixed, "Normalised into one matching key"},
        {"Bank statement", "Rows without a reference", NoRef, "Kept; matched by amount + date where possible"},
        {"Bank statement", "Same reference more than once (not identical rows)", BankReps, "First one matched; the repeat is reported as Duplicate in bank"},
        {"Cash book", "Files read (one per account and month)", LedFiles, "cashbook_<Account>_<yyyy-mm>.csv"},
        {"Cash book", "Rows in the export files", LedRows, "Read as text first, so nothing is lost on import"},
        {"Cash book", "Rows without a reference", LedNoRef, "Kept; matched by amount + date where possible"},
        {"Cash book", "Same reference posted more than once", Repeats, "First posting matched; the repeat is reported as Duplicate in ledger"}})
in
    Log"""

Q["Run_Info"] = """let
    // audit trail: when the data was last refreshed and what was read
    Info = #table(type table [Item = text, Value = any], {
        {"Last refresh", DateTime.ToText(DateTime.LocalNow(), "dd MMM yyyy HH:mm", "en-US")},
        {"Files read", Table.RowCount(Files)},
        {"Bank rows read", Table.RowCount(Bank_Raw)},
        {"Cash book rows", Table.RowCount(Ledger_Raw)},
        {"Repeats removed", Table.RowCount(Bank_Raw) - Table.RowCount(Bank_Clean)},
        {"Accounts", Text.Combine(List.Sort(List.Distinct(Files[Account])), ", ")},
        {"Latest month", Date.ToText(Date.From(List.Max(Bank_Clean[Date])), "MMM yyyy", "en-US")}})
in
    Info"""

SHEETS = ["Summary", "Exceptions workbench", "Reconciliation", "Cleaning log", "Settings", "How to use", "Calc"]


def load_query(wb, ws, name, table_name, at):
    conn = ("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;"
            f"Location={name};Extended Properties=\"\"")
    lo = ws.ListObjects.Add(0, conn, None, 1, ws.Range(at))  # xlSrcExternal, xlYes
    lo.QueryTable.CommandType = 2  # xlCmdSql
    lo.QueryTable.CommandText = f"SELECT * FROM [{name}]"
    lo.QueryTable.RefreshStyle = 1  # xlInsertDeleteCells
    lo.QueryTable.AdjustColumnWidth = False
    lo.QueryTable.PreserveFormatting = True
    lo.QueryTable.Refresh(False)
    lo.Name = table_name
    lo.TableStyle = "Lakeside Dark"
    lo.Range.ClearFormats()  # the page colour must not hide the table style
    return lo


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
        wb.Styles("Normal").Font.Name = lux.SANS
        wb.Styles("Normal").Font.Size = 10
        lux.table_style(wb)
        wb.Queries.FastCombine = True  # the queries read Settings cells and the data folder together (no privacy prompt)

        build_settings(wb)
        wb.SaveAs(str(OUT), 51)  # xlOpenXMLWorkbook; CELL("filename") needs a saved file
        xl.Calculate()

        for name, formula in Q.items():
            wb.Queries.Add(name, formula)

        t0 = time.time()
        rc = wb.Worksheets("Reconciliation")
        lux.page(rc, "W", 1000)
        lo_recon = load_query(wb, rc, "Recon", "Recon", "B8")
        cl = wb.Worksheets("Cleaning log")
        lux.page(cl, "P", 60)
        load_query(wb, cl, "Cleaning_Log", "CleaningLog", "B8")
        load_query(wb, cl, "Run_Info", "RunInfo", "B24")
        load_query(wb, cl, "Bank_Controls", "BankControls", "G8")
        wbk = wb.Worksheets("Exceptions workbench")
        lux.page(wbk, "R", 60)
        lo_wb = load_query(wb, wbk, "Exceptions_Workbench", "Workbench", "B14")
        print(f"queries loaded in {time.time() - t0:.1f}s")

        # demo notes typed by the team, then Refresh: the query reads them back and keeps them with their line
        ids = [r[0] for r in lo_wb.ListColumns("Exception ID").DataBodyRange.Value]
        col = {lo_wb.ListColumns(i).Name: i for i in range(1, lo_wb.ListColumns.Count + 1)}
        for eid, n in NOTES.items():
            r = ids.index(eid) + 1
            lo_wb.DataBodyRange.Cells(r, col["Owner"]).Value = n["owner"]
            lo_wb.DataBodyRange.Cells(r, col["Comment"]).Value = n["comment"]
            if n["resolved"]:
                # Excel serial: COM would shift a naive datetime through UTC (one day back at UTC+7)
                lo_wb.DataBodyRange.Cells(r, col["Resolved date"]).Value = (datetime.fromisoformat(n["resolved"]) - datetime(1899, 12, 30)).days
        lo_wb.QueryTable.Refresh(False)

        format_recon(wb, lo_recon)
        format_cleaning(wb)
        build_calc(wb)
        format_workbench(wb, lo_wb)
        build_summary(wb)
        build_howto(wb)
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            xl.ActiveWindow.DisplayHeadings = ws.Name == "Calc"
            xl.ActiveWindow.Zoom = 100
            ws.Range("A1").Select()
            ws.Tab.Color = c("gold") if ws.Name == "Summary" else (c("card2") if ws.Name != "Calc" else c("grey"))
        wb.Worksheets("Summary").Activate()
        wb.Save()
        print("saved", OUT.name)
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        xl.Quit()


def build_settings(wb):
    st = wb.Worksheets("Settings")
    lux.page(st, "F", 40)
    for col, w in {"A": 2, "B": 30, "C": 46, "D": 74, "E": 2}.items():
        st.Columns(col).ColumnWidth = w
    lux.header(st, "Settings", "Where the files are, the matching tolerances and the opening balances. Yellow-gold cells are yours to change.",
               last_col="E", nav_current="Settings")
    st.Range("B8:D8").Value = ("Setting", "Value", "Note")
    lux.font(st.Range("B8:D8"), "gold", 10, True)
    st.Range("B8:D8").Borders(9).LineStyle = 1; st.Range("B8:D8").Borders(9).Color = c("gold")
    rows = [("Data folder", '=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"data\\"', "DataFolder",
             "Folder with the CSV exports. Default: the data folder next to this workbook. Type a path to override."),
            ("Amount tolerance (±)", 1.00, "AmountTolerance", "Rule 2 and 3: amounts may differ by up to this much and still match"),
            ("Days tolerance (±)", 5, "DaysTolerance", "Rule 3 (no usable reference): bank and cash book dates at most this many days apart"),
            ("Company", "Lakeside Hardware Ltd", "Company", "Shown in the report titles (fictional company)")]
    for r, (label, val, nm, note) in enumerate(rows, 9):
        st.Cells(r, 2).Value = label
        if isinstance(val, str) and val.startswith("="):
            st.Cells(r, 3).Formula2 = val  # Formula2: no implicit-intersection @ (keeps Excel 2016 happy)
        else:
            if isinstance(val, str):
                st.Cells(r, 3).NumberFormat = "@"
            st.Cells(r, 3).Value = val
        st.Cells(r, 4).Value = note
        wb.Names.Add(nm, f"=Settings!$C${r}")
        st.Rows(r).RowHeight = 20
    st.Range("C10").NumberFormat = "0.00"
    st.Range("C11").NumberFormat = "0"
    inp = st.Range("C9:C12")
    lux.fill(inp, "card2"); lux.font(inp, "gold2", 10, True)
    for e in (7, 8, 9, 10, 12):
        inp.Borders(e).LineStyle = 1; inp.Borders(e).Color = c("gold")
    lux.font(st.Range("D9:D12"), "muted", 9)
    st.Range("B9:D12").VerticalAlignment = -4108

    st.Range("B15").Value = "Bank accounts"
    lux.font(st.Range("B15"), "ivory", 14, False, lux.SERIF)
    st.Range("B16").Value = "One row per account. The name must match the file names: bank_<Account>_<yyyy-mm>.csv and cashbook_<Account>_<yyyy-mm>.csv"
    lux.font(st.Range("B16"), "muted", 9)
    st.Range("B17:D17").Value = ("Account", "Opening balance, bank (1st day of the first file)", "Opening balance, cash book (same day)")
    st.Range("B18:D18").Value = ("Operating", 18450.00, 18450.00)
    st.Range("B19:D19").Value = ("Payroll", 6200.00, 6200.00)
    lo = st.ListObjects.Add(1, st.Range("B17:D19"), None, 1)  # xlSrcRange, has headers
    lo.Name = "Accounts"
    lo.ListColumns(2).Name = "Opening bank"
    lo.ListColumns(3).Name = "Opening cash book"
    lo.TableStyle = "Lakeside Dark"
    lo.Range.ClearFormats()
    lo.DataBodyRange.Columns("B:C").NumberFormat = "#,##0.00"
    lux.font(lo.DataBodyRange, "gold2", 10, True)
    lo.HeaderRowRange.HorizontalAlignment = -4131
    st.Range("C18:D19").HorizontalAlignment = -4131
    st.Range("B21").Value = ("The Summary sheet has the account and month pickers. Matching rules: 1 reference + same amount;  "
                             "2 reference + amount within the tolerance;  3 no usable reference: amount within the tolerance and dates within the days tolerance.")
    lux.font(st.Range("B21"), "muted", 9)
    st.PageSetup.PrintArea = "$A$1:$E$22"


def format_recon(wb, lo):
    ws = wb.Worksheets("Reconciliation")
    for c_ in ["Bank amount", "Ledger amount", "Difference", "Value"]:
        lo.ListColumns(c_).DataBodyRange.NumberFormat = "#,##0.00;-#,##0.00"
    for c_ in ["Bank date", "Ledger date"]:
        lo.ListColumns(c_).DataBodyRange.NumberFormat = "dd mmm yyyy"
    lo.ListColumns("Month").DataBodyRange.NumberFormat = "mmm yyyy"
    widths = {"Account": 10, "Status": 17, "Match rule": 16, "Month": 9, "Bank ID": 16, "Bank date": 11.5, "Bank reference": 13,
              "Bank description": 33, "Bank amount": 11.5, "Ledger ID": 16, "Ledger date": 11.5, "Doc no": 12, "Party": 22,
              "Ledger amount": 12, "Difference": 10, "Days to clear": 7, "Value": 11, "Exception ID": 16, "What to do": 56}
    ws.Columns("A").ColumnWidth = 2
    for i in range(1, lo.ListColumns.Count + 1):
        lo.ListColumns(i).Range.ColumnWidth = widths.get(lo.ListColumns(i).Name, 12)
    ws.Columns(lo.Range.Column + lo.ListColumns.Count).ColumnWidth = 2
    lo.HeaderRowRange.WrapText = True
    lo.HeaderRowRange.VerticalAlignment = -4108
    ws.Rows(8).RowHeight = 30
    status_colors(lo.ListColumns("Status").DataBodyRange, "$C9")
    lux.header(ws, "Reconciliation: every line of both files",
               "All accounts and months. Status, the rule that matched it, days to clear and what to do. Filter the table to drill in.",
               last_col="U", nav_current="Reconciliation")
    ws.Range("B7").Value = "Statuses:  Matched (rule 1-3) · Amount differs · Duplicate in ledger · Duplicate in bank · Only in bank · Only in ledger"
    lux.font(ws.Range("B7"), "muted", 9)
    ws.Activate()
    ws.Range("A9").Select()
    wb.Application.ActiveWindow.FreezePanes = False
    ws.Range("D9").Select()
    wb.Application.ActiveWindow.FreezePanes = True


def status_colors(rng, first_cell):
    for text, col in [("Matched", "jade"), ("Amount differs", "bad"), ("Duplicate in ledger", "bad"), ("Duplicate in bank", "bad"),
                      ("Only in bank", "amber"), ("Only in ledger", "sky")]:
        fc = rng.FormatConditions.Add(2, None, f'={first_cell}="{text}"')  # xlExpression
        fc.Font.Color = c(col)
        fc.Font.Bold = True


def format_cleaning(wb):
    ws = wb.Worksheets("Cleaning log")
    for col, w in {"A": 2, "B": 16, "C": 46, "D": 8, "E": 58, "F": 3, "G": 11, "H": 10, "I": 7, "J": 13, "K": 13,
                   "L": 13, "M": 11, "N": 10, "O": 2}.items():
        ws.Columns(col).ColumnWidth = w
    lux.header(ws, "Cleaning log and statement controls",
               "Every cleaning step counted, nothing dropped silently. Each statement file is checked against the balances printed on it.",
               last_col="O", nav_current="Cleaning log")
    ws.Range("B7").Value = "What Power Query did to the raw exports"
    ws.Range("B23").Value = "Audit trail of the last refresh"
    ws.Range("G7").Value = "Statement files: opening + movements = printed closing"
    for a in ("B7", "B23", "G7"):
        lux.font(ws.Range(a), "ivory", 12, False, lux.SERIF)
    bc = ws.ListObjects("BankControls")
    for c_ in ["Printed opening", "Movements", "Printed closing"]:
        bc.ListColumns(c_).DataBodyRange.NumberFormat = "#,##0.00;-#,##0.00"
    bc.ListColumns("Month").DataBodyRange.NumberFormat = "mmm yyyy"
    bc.Range.Columns.HorizontalAlignment = -4131
    for c_ in ["Printed opening", "Movements", "Printed closing", "Lines"]:
        bc.ListColumns(c_).Range.HorizontalAlignment = -4152
    for name in ("Inside the file", "Chain"):
        rng = bc.ListColumns(name).DataBodyRange
        addr = rng.Cells(1, 1).GetAddress(False, False)
        for text, col in [("OK", "good"), ("BREAK", "bad"), ("First file", "muted")]:
            fc = rng.FormatConditions.Add(2, None, f'={addr}="{text}"'); fc.Font.Color = c(col); fc.Font.Bold = text != "First file"
    ws.ListObjects("CleaningLog").ListColumns("Rows").Range.HorizontalAlignment = -4152
    ws.ListObjects("RunInfo").ListColumns("Value").Range.HorizontalAlignment = -4131
    bc.HeaderRowRange.WrapText = True
    ws.ListObjects("CleaningLog").HeaderRowRange.WrapText = True
    ws.Rows(8).RowHeight = 30
    ws.ListObjects("CleaningLog").DataBodyRange.Columns(4).WrapText = False


def format_workbench(wb, lo):
    ws = wb.Worksheets("Exceptions workbench")
    widths = {"Flag": 5, "Exception ID": 16, "Account": 10, "Status": 17, "Date": 11, "Age (days)": 6, "Reference": 13,
              "Description": 30, "Amount": 11, "What to do": 60, "Owner": 8, "Comment": 52, "Resolved date": 11.5,
              "State": 9, "Abs amount": 10}
    ws.Columns("A").ColumnWidth = 2
    for i in range(1, lo.ListColumns.Count + 1):
        lo.ListColumns(i).Range.ColumnWidth = widths.get(lo.ListColumns(i).Name, 12)
    lo.ListColumns("Abs amount").Range.EntireColumn.Hidden = True
    lo.ListColumns("Amount").DataBodyRange.NumberFormat = "#,##0.00;-#,##0.00"
    lo.ListColumns("Date").DataBodyRange.NumberFormat = "dd mmm yyyy"
    lo.ListColumns("Resolved date").DataBodyRange.NumberFormat = "dd mmm yyyy"
    lo.HeaderRowRange.WrapText = True
    lo.HeaderRowRange.VerticalAlignment = -4108
    ws.Rows(14).RowHeight = 30
    # status icons: 1 resolved (green tick), 2 open (amber !), 3 open over 30 days (red cross)
    flag = lo.ListColumns("Flag").DataBodyRange
    flag.HorizontalAlignment = -4108
    ic = flag.FormatConditions.AddIconSetCondition()
    ic.IconSet = wb.IconSets(8)  # xl3Symbols2 (tick, !, cross)
    ic.ReverseOrder = True
    ic.ShowIconOnly = True
    ic.IconCriteria(2).Type = 0; ic.IconCriteria(2).Value = 2; ic.IconCriteria(2).Operator = 7  # >= 2
    ic.IconCriteria(3).Type = 0; ic.IconCriteria(3).Value = 3; ic.IconCriteria(3).Operator = 7  # >= 3
    status_colors(lo.ListColumns("Status").DataBodyRange, "$E15")
    st = lo.ListColumns("State").DataBodyRange
    for text, col in [("Resolved", "good"), ("Open", "amber")]:
        fc = st.FormatConditions.Add(2, None, f'=$O15="{text}"'); fc.Font.Color = c(col); fc.Font.Bold = True
    # the three columns the team types in: gold header, slightly lighter cells
    for name in ("Owner", "Comment", "Resolved date"):
        lc = lo.ListColumns(name)
        lux.fill(lc.DataBodyRange, "#232B35")
        lux.font(lc.DataBodyRange, "gold2", 10)
    lux.header(ws, "Exceptions workbench",
               "Every line that is not matched. Type Owner, Comment and Resolved date; Refresh All keeps them with their line and the Summary follows.",
               last_col="O", nav_current="Exceptions workbench")
    lo.ListColumns("Abs amount").Range.EntireColumn.Hidden = True  # helper for the Summary formulas
    # filter hints (slicers on a Power Query table crashed Excel during the COM build; the table's own filter buttons do the job)
    hints = [("B7", "FILTER AND REVIEW", "gold", 8, True),
             ("B8", "Use the filter buttons on Account, Status, Owner and State; sort by Age (days) to see the oldest first.", "ivory", 10, False),
             ("B9", "Icons:  green tick = resolved   ·   amber ! = open   ·   red cross = open for more than 30 days", "muted", 9, False),
             ("B10", "Gold columns (Owner, Comment, Resolved date) are typed by your team; they stay with the line after Data > Refresh All.", "gold2", 9, False)]
    for cell, t, col, size, bold in hints:
        ws.Range(cell).Value = t
        lux.font(ws.Range(cell), col, size, bold)
    ws.Range("B12").Formula = '="Open: "&COUNTIFS(Workbench[State],"Open")&"   ·   Resolved: "&COUNTIFS(Workbench[State],"Resolved")&"   ·   Open amount: "&TEXT(SUMIFS(Workbench[Abs amount],Workbench[State],"Open"),"#,##0.00")'
    lux.font(ws.Range("B12"), "ivory", 11, True)
    ws.Activate()
    ws.Range("D15").Select()
    wb.Application.ActiveWindow.FreezePanes = True


def build_howto(wb):
    ws = wb.Worksheets("How to use")
    lux.page(ws, "E", 45)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 150
    ws.Columns("C").ColumnWidth = 2
    lux.header(ws, "How to use this workbook", "Monthly routine, what each status means, how the matching works.", last_col="C",
               nav_current="How to use")
    text = [
        ("Every month", True),
        ("1. Export each bank statement and cash book as CSV into the data folder: bank_<Account>_<yyyy-mm>.csv and cashbook_<Account>_<yyyy-mm>.csv.", False),
        ("2. New account? Add a row on Settings > Bank accounts with its opening balances.", False),
        ("3. Data > Refresh All. Every tab updates; no formulas to copy, no macros.", False),
        ("4. Summary: pick the account and month. The unexplained difference must be 0.00.", False),
        ("5. Exceptions workbench: give every open line an owner, a comment and, when done, a resolved date. Refresh All keeps them.", False),
        ("", False),
        ("What each status means", True),
        ("Matched: rule 1 same reference and amount; rule 2 same reference, amount within the tolerance; rule 3 no usable reference, amount and date within the tolerances.", False),
        ("Amount differs: same reference, different amount. Usually a short payment, a fee deducted, or a typo.", False),
        ("Duplicate in ledger / in bank: the same reference posted or paid more than once.", False),
        ("Only in bank: bank charges, interest, direct debits or receipts nobody has booked yet.", False),
        ("Only in ledger: deposits in transit and payments that have not cleared yet (stale after 60 days).", False),
        ("", False),
        ("How the match works", True),
        ("References are cleaned first (upper case, no spaces or dashes), so 'inv-1023 ', 'INV1023' and 'INV-1023' are the same key.", False),
        ("Matching runs per account over all months, so a cheque written on 30 September that clears on 2 October is matched, with its days to clear.", False),
        ("Rule 3 takes the closest pair first (amount, then days) and uses every line once. Tolerances are on Settings.", False),
        ("The steps live in Power Query (Data > Queries & Connections): Files, Bank_Clean, Ledger_Clean, Recon, Exceptions_Workbench.", False),
    ]
    for i, (t, bold) in enumerate(text, 8):
        cell = ws.Cells(i, 2)
        cell.Value = t
        if bold:
            lux.font(cell, "gold", 13, False, lux.SERIF)
            ws.Rows(i).RowHeight = 22
        else:
            lux.font(cell, "ivory", 10)
            ws.Rows(i).RowHeight = 17


if __name__ == "__main__":
    main()
