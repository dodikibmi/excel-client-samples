"""Builds the Power BI project (PBIP) for E9 as plain text files that Power BI Desktop opens directly:

  Copperleaf-Finance.pbip                    the file to double-click
  Copperleaf-Finance.SemanticModel/          the model in TMDL: Power Query (M) that cleans the bank and card exports,
                                             tables, star-schema relationships and every DAX measure
  Copperleaf-Finance.Report/                 the report in PBIR: four pages, every visual as one JSON file

Everything is generated from this script, so the model and the report can be reviewed in Git line by line.
Run:  python scripts/build_pbip.py [--data-folder PATH]    (default: C:/Copperleaf/data, a neutral path
      that the published files carry; copy the data folder there before Refresh)
Then open the .pbip in Power BI Desktop and press Refresh once (a PBIP keeps no data of its own).
"""
import argparse, json, shutil, uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
NAME = "Copperleaf-Finance"
SM, RP = ROOT / f"{NAME}.SemanticModel", ROOT / f"{NAME}.Report"
NS = uuid.UUID("5b0f8a4e-2f1c-4e55-9a37-0c9e6a1d2e09")


def tag(*parts):
    return str(uuid.uuid5(NS, "/".join(parts)))


def indent(text, tabs):
    return "\n".join("\t" * tabs + line if line.strip() else "" for line in text.strip("\n").splitlines())


def q(name):  # TMDL object name, quoted when needed
    return name if name.replace("_", "").isalnum() else "'" + name.replace("'", "''") + "'"


# ====================================================================================== Power Query (M)
NORM = r'''
(t as nullable text) as text =>
let
    // upper case, punctuation to spaces, single spaces: "TapPay Inc. Deposit" -> "TAPPAY INC DEPOSIT"
    u = Text.Upper(if t = null then "" else t),
    spaced = Text.Combine(List.Transform(Text.ToList(u), each if List.Contains({"*", ".", "-", "#", "/", ",", "&", "'"}, _) then " " else _)),
    words = List.Select(Text.Split(spaced, " "), each _ <> "")
in
    Text.Combine(words, " ")
'''
MONEY = r'''
(t as nullable text) as number =>
let
    // "$1,250.00" / "1,250.00" / "-64.18" / "" -> number
    s = if t = null then "" else Text.Remove(t, {"$", ",", " "})
in
    if s = "" then 0 else Number.FromText(s, "en-US")
'''


def csv_src(file, cols):
    return (f'Csv.Document(File.Contents(DataFolder & "{file}"), '
            f'[Delimiter = ",", Columns = {cols}, Encoding = 65001, QuoteStyle = QuoteStyle.Csv])')


CHECKING = f'''
let
    // bank export: 3 title lines above the header, MM/DD/YYYY, Debit and Credit as text "$1,250.00"
    Source = {csv_src("checking-2026-08-10.csv", 5)},
    FromHeader = Table.Skip(Source, each [Column1] <> "Date"),
    Promoted = Table.PromoteHeaders(FromHeader, [PromoteAllScalars = true]),
    Rows = Table.SelectRows(Promoted, each Text.Trim([Date]) <> ""),
    Indexed = Table.AddIndexColumn(Rows, "Row", 1, 1, Int64.Type),
    Out = Table.FromRecords(Table.TransformRows(Indexed, (r) => [
        Line ID = "CHK-" & Text.PadStart(Text.From(r[Row]), 3, "0"),
        Source = "Checking",
        Register No = 1000,
        Txn Date = Date.FromText(Text.Trim(r[Date]), [Format = "MM/dd/yyyy", Culture = "en-US"]),
        Post Date = Date.FromText(Text.Trim(r[Date]), [Format = "MM/dd/yyyy", Culture = "en-US"]),
        Description = Text.Trim(r[Description]),
        Amount = Money(r[Credit]) - Money(r[Debit])
    ]))
in
    Out
'''.replace("Line ID =", '#"Line ID" =').replace("Register No =", '#"Register No" =') \
   .replace("Txn Date =", '#"Txn Date" =').replace("Post Date =", '#"Post Date" =')

SAVINGS = f'''
let
    // savings export: ISO dates, one signed Amount column with thousand separators
    Source = {csv_src("savings-2026-08-10.csv", 3)},
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Rows = Table.SelectRows(Promoted, each Text.Trim([Posted]) <> ""),
    Indexed = Table.AddIndexColumn(Rows, "Row", 1, 1, Int64.Type),
    Out = Table.FromRecords(Table.TransformRows(Indexed, (r) => [
        #"Line ID" = "SAV-" & Text.PadStart(Text.From(r[Row]), 3, "0"),
        Source = "Savings",
        #"Register No" = 1010,
        #"Txn Date" = Date.FromText(Text.Trim(r[Posted]), [Format = "yyyy-MM-dd", Culture = "en-US"]),
        #"Post Date" = Date.FromText(Text.Trim(r[Posted]), [Format = "yyyy-MM-dd", Culture = "en-US"]),
        Description = Text.Trim(r[Details]),
        Amount = Money(r[Amount])
    ]))
in
    Out
'''

CARD = f'''
let
    // card export: transaction and post date (M/D/YYYY), charges negative; booked by POST date like the statement
    Source = {csv_src("visa-2026-08-10.csv", 6)},
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Rows = Table.SelectRows(Promoted, each Text.Trim([Post Date]) <> ""),
    Indexed = Table.AddIndexColumn(Rows, "Row", 1, 1, Int64.Type),
    Out = Table.FromRecords(Table.TransformRows(Indexed, (r) => [
        #"Line ID" = "VISA-" & Text.PadStart(Text.From(r[Row]), 3, "0"),
        Source = "Card",
        #"Register No" = 2000,
        #"Txn Date" = Date.FromText(Text.Trim(r[Transaction Date]), [Format = "M/d/yyyy", Culture = "en-US"]),
        #"Post Date" = Date.FromText(Text.Trim(r[Post Date]), [Format = "M/d/yyyy", Culture = "en-US"]),
        Description = Text.Trim(r[Description]),
        Amount = Money(r[Amount])
    ]))
in
    Out
'''

RULES = f'''
let
    // keyword -> account, first match wins (rules.csv is in priority order)
    Source = {csv_src("rules.csv", 5)},
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {{{{"Priority", Int64.Type}}, {{"Account No", Int64.Type}}}}, "en-US"),
    Sorted = Table.Sort(Typed, {{{{"Priority", Order.Ascending}}}}),
    Clean = Table.AddColumn(Sorted, "Clean Keyword", each Norm([Keyword]), type text)
in
    Clean
'''

BANKLINES = r'''
let
    // the three registers as one list; every line gets an account from the first rule whose keyword it contains
    Lines = Table.Combine({Checking_Clean, Savings_Clean, Card_Clean}),
    RuleList = Table.ToRecords(Table.SelectColumns(Rules, {"Clean Keyword", "Account No", "Payee"})),
    WithClean = Table.AddColumn(Lines, "Clean Description", each Norm([Description]), type text),
    WithRule = Table.AddColumn(WithClean, "Rule", (r) => List.First(List.Select(RuleList, (k) => Text.Contains(r[Clean Description], k[Clean Keyword])), null)),
    WithAccount = Table.AddColumn(WithRule, "Offset Account No", each if [Rule] = null then 6990 else [Rule][Account No], Int64.Type),
    WithKeyword = Table.AddColumn(WithAccount, "Rule Keyword", each if [Rule] = null then "(no rule: to review)" else [Rule][Clean Keyword], type text),
    WithPayee = Table.AddColumn(WithKeyword, "Payee", each if [Rule] = null then [Description] else [Rule][Payee], type text)
in
    Table.RemoveColumns(WithPayee, {"Rule"})
'''

ACCOUNTS_RAW = f'''
let
    Source = {csv_src("accounts.csv", 6)},
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {{{{"Account No", Int64.Type}}, {{"Opening Balance", Currency.Type}}}}, "en-US")
in
    Typed
'''

ADJ = f'''
let
    // month-end journal entries typed by the bookkeeper (depreciation)
    Source = {csv_src("adjustments.csv", 6)},
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {{{{"Date", type date}}, {{"Account No", Int64.Type}}, {{"Debit", type number}}, {{"Credit", type number}}}}, "en-US")
in
    Typed
'''

LEDGER = r'''
let
    // double entry: every bank line posts twice (its register and the account the rule chose), plus the journal entries
    Bank = Table.ToRecords(BankLines),
    Posted = List.Combine(List.Transform(Bank, (r) => {
        [#"Line ID" = r[Line ID] & "-R", Side = "Register", Date = r[Post Date], #"Txn Date" = r[Txn Date], Source = r[Source],
         #"Register No" = r[Register No], Description = r[Description], #"Clean Description" = r[Clean Description],
         #"Rule Keyword" = r[Rule Keyword], Payee = r[Payee], #"Account No" = r[Register No], #"Offset Account No" = r[Offset Account No], Amount = r[Amount]],
        [#"Line ID" = r[Line ID] & "-A", Side = "Account", Date = r[Post Date], #"Txn Date" = r[Txn Date], Source = r[Source],
         #"Register No" = r[Register No], Description = r[Description], #"Clean Description" = r[Clean Description],
         #"Rule Keyword" = r[Rule Keyword], Payee = r[Payee], #"Account No" = r[Offset Account No], #"Offset Account No" = r[Register No], Amount = -r[Amount]]
    })),
    Journal = List.Transform(List.Positions(Table.ToRecords(Adjustments)), (i) =>
        let a = Table.ToRecords(Adjustments){i} in
        [#"Line ID" = a[Entry] & "-" & Text.From(i + 1), Side = "Journal", Date = a[Date], #"Txn Date" = a[Date], Source = "Journal",
         #"Register No" = null, Description = a[Description], #"Clean Description" = Norm(a[Description]),
         #"Rule Keyword" = null, Payee = "Journal entry", #"Account No" = a[Account No], #"Offset Account No" = null, Amount = a[Debit] - a[Credit]]),
    All = Table.FromRecords(Posted & Journal),
    Rounded = Table.TransformColumns(All, {{"Amount", each Number.Round(_, 2), type number}}),
    WithDebit = Table.AddColumn(Rounded, "Debit", each if [Amount] > 0 then [Amount] else 0, Currency.Type),
    WithCredit = Table.AddColumn(WithDebit, "Credit", each if [Amount] < 0 then -[Amount] else 0, Currency.Type),
    // cash flow line for the register side of checking and savings (what kind of money moved)
    WithFlow = Table.AddColumn(WithCredit, "Cash Flow Line", each
        if [Side] <> "Register" then null
        else if [Offset Account No] >= 4000 and [Offset Account No] < 5000 then "Receipts from customers"
        else if [Offset Account No] >= 5000 and [Offset Account No] < 7000 then "Payments to suppliers and staff"
        else if [Offset Account No] = 3100 then "Owner's draws"
        else if Text.Contains([Clean Description], "VISA") or Text.Contains([Clean Description], "PAYMENT THANK YOU") then "Card payments"
        else "Transfers between own accounts", type text),
    WithOrder = Table.AddColumn(WithFlow, "Cash Flow Order", each
        List.PositionOf({"Receipts from customers", "Payments to suppliers and staff", "Owner's draws", "Card payments", "Transfers between own accounts"}, [Cash Flow Line]) + 1, Int64.Type),
    Typed = Table.TransformColumnTypes(WithOrder, {
        {"Line ID", type text}, {"Side", type text}, {"Date", type date}, {"Txn Date", type date}, {"Source", type text},
        {"Register No", Int64.Type}, {"Description", type text}, {"Clean Description", type text}, {"Rule Keyword", type text}, {"Payee", type text},
        {"Account No", Int64.Type}, {"Offset Account No", Int64.Type}, {"Amount", Currency.Type}})
in
    Typed
'''

ACCOUNT = r'''
let
    Source = AccountsRaw,
    Label = Table.AddColumn(Source, "Account Label", each Text.From([Account No]) & " " & [Account], type text),
    Cash = Table.AddColumn(Label, "Is Cash Account", each [Is Cash] = "Y", type logical),
    Key = Table.AddColumn(Cash, "Category Key", each [Type], type text),
    // the slices of the cost mix chart
    Group = Table.AddColumn(Key, "Cost Group", each
        if [Type] = "COGS" then "Ingredients and packaging"
        else if [Account No] = 6000 then "Payroll"
        else if [Account No] = 6100 then "Rent"
        else if [Account No] = 6200 then "Utilities"
        else if [Type] = "Expense" then "Other overheads" else null, type text),
    Out = Table.RenameColumns(Table.RemoveColumns(Group, {"Is Cash"}), {{"Is Cash Account", "Is Cash"}})
in
    Out
'''

CATEGORY = r'''
let
    Source = #table(
        type table [#"Category Key" = text, Category = text, Group = text, Statement = text, Sort = Int64.Type],
        {
            {"Revenue", "Revenue", "Income", "P&L", 1},
            {"COGS", "Cost of goods sold", "Expense", "P&L", 2},
            {"Expense", "Operating expenses", "Expense", "P&L", 3},
            {"Asset", "Assets", "Balance sheet", "Balance sheet", 4},
            {"Liability", "Liabilities", "Balance sheet", "Balance sheet", 5},
            {"Equity", "Equity", "Balance sheet", "Balance sheet", 6}
        })
in
    Source
'''

BUDGET = f'''
let
    // typed by the client as positive amounts; stored with the P&L sign (income +, costs -) so Actual - Budget = better (+) or worse (-)
    Source = {csv_src("budget.csv", 3)},
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {{{{"Account No", Int64.Type}}, {{"Month", type date}}, {{"Budget", Currency.Type}}}}, "en-US"),
    Types = Table.SelectColumns(AccountsRaw, {{"Account No", "Type"}}),
    Joined = Table.NestedJoin(Typed, {{"Account No"}}, Types, {{"Account No"}}, "A", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "A", {{"Type"}}),
    Signed = Table.AddColumn(Expanded, "Budget Amount", each if [Type] = "Revenue" then [Budget] else -[Budget], Currency.Type),
    Out = Table.RenameColumns(Table.SelectColumns(Signed, {{"Account No", "Month", "Budget", "Budget Amount"}}), {{{{"Budget", "Budget As Typed"}}}})
in
    Out
'''

STATEMENTS = f'''
let
    // month-end balance printed on each statement; the card prints what is owed as a positive number
    Source = {csv_src("statements-2026-08-10.csv", 3)},
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Accts = Table.SelectColumns(Table.SelectRows(AccountsRaw, each [Statement Name] <> ""), {{"Statement Name", "Account No", "Type"}}),
    Joined = Table.NestedJoin(Promoted, {{"Account"}}, Accts, {{"Statement Name"}}, "A", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "A", {{"Account No", "Type"}}),
    Out = Table.FromRecords(Table.TransformRows(Expanded, (r) => [
        #"Account No" = r[Account No],
        #"Statement Date" = Date.FromText(Text.Trim(r[Statement date]), [Format = "MM/dd/yyyy", Culture = "en-US"]),
        #"Printed Balance" = Money(r[Ending balance]),
        #"Ending Balance" = (if r[Type] = "Liability" then -1 else 1) * Money(r[Ending balance])
    ])),
    Typed = Table.TransformColumnTypes(Out, {{{{"Account No", Int64.Type}}, {{"Statement Date", type date}}, {{"Printed Balance", Currency.Type}}, {{"Ending Balance", Currency.Type}}}})
in
    Typed
'''

DATE = r'''
let
    // one row per day, whole months, from the first to the last month in the exports
    Start = Date.StartOfMonth(List.Min(BankLines[Post Date])),
    End = Date.EndOfMonth(List.Max(BankLines[Post Date])),
    Days = List.Dates(Start, Duration.Days(End - Start) + 1, #duration(1, 0, 0, 0)),
    Base = Table.FromList(Days, Splitter.SplitByNothing(), type table [Date = date]),
    Year = Table.AddColumn(Base, "Year", each Date.Year([Date]), Int64.Type),
    MonthNo = Table.AddColumn(Year, "Month No", each Date.Year([Date]) * 100 + Date.Month([Date]), Int64.Type),
    Month = Table.AddColumn(MonthNo, "Month", each Date.ToText([Date], [Format = "MMM yyyy", Culture = "en-US"]), type text),
    MStart = Table.AddColumn(Month, "Month Start", each Date.StartOfMonth([Date]), type date),
    MEnd = Table.AddColumn(MStart, "Month End", each Date.EndOfMonth([Date]), type date),
    Day = Table.AddColumn(MEnd, "Day", each Date.Day([Date]), Int64.Type),
    Weekday = Table.AddColumn(Day, "Weekday", each Date.ToText([Date], [Format = "ddd", Culture = "en-US"]), type text),
    WeekdayNo = Table.AddColumn(Weekday, "Weekday No", each Date.DayOfWeek([Date], Day.Monday) + 1, Int64.Type)
in
    WeekdayNo
'''

SOURCE_FILES = f'''
let
    // control totals read straight from each export (a second, independent parse) for the Checks page
    Num = (t as nullable text) as number => if t = null or Text.Trim(t) = "" then 0 else Number.FromText(Text.Select(t, {{"0".."9", ".", "-"}}), "en-US"),
    ChkAll = {csv_src("checking-2026-08-10.csv", 5)},
    ChkHeader = List.PositionOf(ChkAll[Column1], "Date"),
    Chk = Table.SelectRows(Table.Skip(ChkAll, ChkHeader + 1), each Text.Trim([Column1]) <> ""),
    Sav = Table.SelectRows(Table.Skip({csv_src("savings-2026-08-10.csv", 3)}, 1), each Text.Trim([Column1]) <> ""),
    Card = Table.SelectRows(Table.Skip({csv_src("visa-2026-08-10.csv", 6)}, 1), each Text.Trim([Column1]) <> ""),
    Rows = {{
        {{1000, "checking-2026-08-10.csv", Table.RowCount(Chk), List.Sum(List.Transform(Chk[Column4], Num)) - List.Sum(List.Transform(Chk[Column3], Num)), Num(List.Last(Chk[Column5]))}},
        {{1010, "savings-2026-08-10.csv", Table.RowCount(Sav), List.Sum(List.Transform(Sav[Column3], Num)), null}},
        {{2000, "visa-2026-08-10.csv", Table.RowCount(Card), List.Sum(List.Transform(Card[Column6], Num)), null}}
    }},
    T = #table(type table [#"Register No" = Int64.Type, File = text, #"Rows In File" = Int64.Type, #"Net Amount In File" = number, #"Bank Running Balance" = number], Rows),
    Typed = Table.TransformColumnTypes(T, {{{{"Net Amount In File", Currency.Type}}, {{"Bank Running Balance", Currency.Type}}}})
in
    Typed
'''

CHECKS = r'''
let
    Source = #table(
        type table [No = Int64.Type, Check = text, #"Why it matters" = text],
        {
            {1, "Opening balances balance", "Debits = credits in the chart of accounts before the first line is posted"},
            {2, "Ledger balances: debits = credits", "Every bank line is posted twice, so the ledger must net to zero"},
            {3, "Trial balance at every month end", "Balances of all accounts add to zero at each month end"},
            {4, "Every export line is in the model once", "Rows per file = register lines in the model (nothing lost, nothing doubled)"},
            {5, "Total per export = total in the model", "Net amount per file, read a second way, = what was posted"},
            {6, "Books = bank and card statements", "Closing balance per account and month = the printed statement"},
            {7, "Checking = the bank's own running balance", "Last Balance column in the checking export = books"},
            {8, "Every line has an account in the chart", "No ledger line points to an account that does not exist"},
            {9, "Lines waiting for a rule", "Uncategorized lines: add a keyword to rules.csv and refresh"}
        })
in
    Source
'''

INSIGHTS = r'''
let
    Source = #table(
        type table [No = Int64.Type, Topic = text],
        {{1, "Profit"}, {2, "Income and costs"}, {3, "Biggest overrun"}, {4, "Biggest saving"}, {5, "Cash"}, {6, "Next 13 weeks"}})
in
    Source
'''

CASHPLAN = r"""
let
    // the owner's plan for the next 13 weeks: one row per regular receipt or payment (cash-plan.csv);
    // the amount comes from the history unless the plan types it (fixed)
    Source = SRC_PLAN,
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Order", Int64.Type}, {"Register No", Int64.Type}, {"Lookback", Int64.Type},
        {"First Date", type date}, {"Amount", type number}}, "en-US"),
    LastDate = List.Max(BankLines[Post Date]),
    Start = Date.AddDays(LastDate, 1),
    End = Date.AddDays(Start, 13 * 7 - 1),
    Hist = Table.SelectColumns(BankLines, {"Register No", "Offset Account No", "Post Date", "Amount"}),
    Lines = (r) => let accts = List.Transform(Text.Split(r[Accounts], ";"), Number.From) in
        Table.SelectRows(Hist, each [Register No] = r[Register No] and List.Contains(accts, [Offset Account No])),
    PerTime = (r) =>
        if r[Basis] = "fixed" then r[Amount]
        else if r[Basis] = "avg_week" then
            List.Sum(Table.SelectRows(Lines(r), each [Post Date] > Date.AddDays(LastDate, -7 * r[Lookback]))[Amount]) / r[Lookback]
        else if r[Basis] = "avg_month" then
            List.Sum(Table.SelectRows(Lines(r), each [Post Date] > Date.EndOfMonth(Date.AddMonths(LastDate, -r[Lookback])))[Amount]) / r[Lookback]
        else
            // avg_event: the last N dates with a posting (pay runs), averaged
            let g = Table.Group(Lines(r), {"Post Date"}, {{"S", each List.Sum([Amount]), type number}}),
                top = Table.FirstN(Table.Sort(g, {{"Post Date", Order.Descending}}), r[Lookback])
            in List.Average(top[S]),
    Dates = (r) =>
        if r[Frequency] = "weekly" then List.Dates(Start, 13, #duration(7, 0, 0, 0))
        else if r[Frequency] = "biweekly" then List.Select(List.Dates(r[First Date], 8, #duration(14, 0, 0, 0)), each _ >= Start and _ <= End)
        else List.Select(List.Transform({0..3}, (k) => Date.AddMonths(r[First Date], k)), each _ >= Start and _ <= End),
    WithPer = Table.AddColumn(Typed, "Per Time", each Number.Round(PerTime(_), 2), type number),
    WithDates = Table.AddColumn(WithPer, "Dates", each Dates(_)),
    WithBasis = Table.AddColumn(WithDates, "How It Is Estimated", each
        if [Basis] = "fixed" then "Fixed amount"
        else if [Basis] = "avg_week" then "Average week, last " & Text.From([Lookback]) & " weeks"
        else if [Basis] = "avg_month" then "Average month, last " & Text.From([Lookback]) & " months"
        else "Average of the last " & Text.From([Lookback]) & " runs", type text),
    WithWhen = Table.AddColumn(WithBasis, "When", each
        if [Frequency] = "weekly" then "Every week"
        else if [Frequency] = "biweekly" then "Every 2nd week from " & Date.ToText([First Date], [Format = "d MMM", Culture = "en-US"])
        else "Monthly on day " & Text.From(Date.Day([First Date])), type text)
in
    WithWhen
""".replace("SRC_PLAN", csv_src("cash-plan.csv", 11))

CASH_PLAN_TABLE = r"""
let
    Source = CashPlanCalc,
    Times = Table.AddColumn(Source, "Times", each List.Count([Dates]), Int64.Type),
    Total = Table.AddColumn(Times, "Total 13 Weeks", each [Per Time] * [Times], Currency.Type),
    Out = Table.SelectColumns(Total, {"Order", "Item", "Cash Flow Line", "How It Is Estimated", "When", "Per Time", "Times", "Total 13 Weeks", "Note"}),
    Typed = Table.TransformColumnTypes(Out, {{"Per Time", Currency.Type}})
in
    Typed
"""

FORECAST = r"""
let
    // one row per expected receipt or payment in the next 13 weeks
    Start = Date.AddDays(List.Max(BankLines[Post Date]), 1),
    Rows = Table.SelectColumns(CashPlanCalc, {"Order", "Item", "Cash Flow Line", "Per Time", "Dates"}),
    Expanded = Table.ExpandListColumn(Rows, "Dates"),
    Renamed = Table.RenameColumns(Expanded, {{"Dates", "Date"}, {"Per Time", "Amount"}}),
    Week = Table.AddColumn(Renamed, "Week No", each Number.IntegerDivide(Duration.Days([Date] - Start), 7) + 1, Int64.Type),
    Typed = Table.TransformColumnTypes(Week, {{"Date", type date}, {"Amount", Currency.Type}})
in
    Typed
"""

FORECAST_WEEK = r"""
let
    Start = Date.AddDays(List.Max(BankLines[Post Date]), 1),
    Starts = List.Dates(Start, 13, #duration(7, 0, 0, 0)),
    T = Table.FromRecords(List.Transform(List.Positions(Starts), (i) => [
        #"Week No" = i + 1, #"Week Start" = Starts{i}, #"Week End" = Date.AddDays(Starts{i}, 6),
        Week = "W" & Text.From(i + 1) & " · " & Date.ToText(Date.AddDays(Starts{i}, 6), [Format = "d MMM", Culture = "en-US"])])),
    Typed = Table.TransformColumnTypes(T, {{"Week No", Int64.Type}, {"Week Start", type date}, {"Week End", type date}, {"Week", type text}})
in
    Typed
"""

CASH_BUFFER = r"""
let
    // what-if: the minimum cash the owner wants to keep (picked in a slicer, 10,000 when nothing is picked)
    Source = Table.FromList(List.Numbers(0, 13, 2500), Splitter.SplitByNothing(), type table [#"Cash Buffer" = Int64.Type])
in
    Source
"""

MEASURES_SRC = r'''
let
    Source = #table(type table [#"Measure Group" = text], {})
in
    Source
'''

# ====================================================================================== DAX measures
MEASURES = [  # name, folder, format, expression, hidden
    ("Total Income", "1 Profit and loss", "#,0", 'CALCULATE ( -SUM ( Ledger[Amount] ), KEEPFILTERS ( Category[Group] = "Income" ) )', False),
    ("Cost of Sales", "1 Profit and loss", "#,0", 'CALCULATE ( SUM ( Ledger[Amount] ), KEEPFILTERS ( Category[Category Key] = "COGS" ) )', False),
    ("Operating Expenses", "1 Profit and loss", "#,0", 'CALCULATE ( SUM ( Ledger[Amount] ), KEEPFILTERS ( Category[Category Key] = "Expense" ) )', False),
    ("Total Expense", "1 Profit and loss", "#,0", 'CALCULATE ( SUM ( Ledger[Amount] ), KEEPFILTERS ( Category[Group] = "Expense" ) )', False),
    ("Gross Profit", "1 Profit and loss", "#,0;(#,0);0", "[Total Income] - [Cost of Sales]", False),
    ("Net Profit", "1 Profit and loss", "#,0;(#,0);0", "[Total Income] - [Total Expense]", False),
    ("Net Profit Color", "1 Profit and loss", None, 'IF ( [Net Profit] < 0, "#E5735F", "#5BBF8A" )', True),
    ("Net Margin %", "1 Profit and loss", "0.0%;-0.0%;0.0%", "DIVIDE ( [Net Profit], [Total Income] )", False),
    ("Actual", "2 Budget", "#,0;(#,0);0",
     '-- P&L sign: income +, costs -. The total of the P&L is the net profit\nCALCULATE ( -SUM ( Ledger[Amount] ), KEEPFILTERS ( Category[Statement] = "P&L" ) )', False),
    ("Budget", "2 Budget", "#,0;(#,0);0", "SUM ( Budget[Budget Amount] )", False),
    ("Variance", "2 Budget", "+#,0;-#,0;0", "-- + = better than budget (more income or less cost)\n[Actual] - [Budget]", False),
    ("Variance %", "2 Budget", "+0.0%;-0.0%;0.0%", "DIVIDE ( [Variance], ABS ( [Budget] ) )", False),
    ("Variance Color", "2 Budget", None,
     'SWITCH ( TRUE (), [Variance] < -0.5, "#E5735F", [Variance] > 0.5, "#5BBF8A", "#8E98A3" )', True),
    ("Income Previous Month", "3 Time intelligence", "#,0",
     "CALCULATE ( [Total Income], PREVIOUSMONTH ( LASTDATE ( 'Date'[Date] ) ) )", False),
    ("Income Last Month", "3 Time intelligence", "#,0",
     "-- the last month in the selection, whole month\nCALCULATE ( [Total Income], DATESMTD ( LASTDATE ( 'Date'[Date] ) ) )", False),
    ("Income MoM %", "3 Time intelligence", "+0.0%;-0.0%;0.0%",
     "VAR cur = [Income Last Month]\nVAR prev = [Income Previous Month]\nRETURN\n    IF ( NOT ISBLANK ( prev ), DIVIDE ( cur - prev, prev ) )", False),
    ("Net Profit Previous Month", "3 Time intelligence", "#,0;(#,0);0",
     "CALCULATE ( [Net Profit], PREVIOUSMONTH ( LASTDATE ( 'Date'[Date] ) ) )", False),
    ("Net Profit YTD", "3 Time intelligence", "#,0;(#,0);0", "TOTALYTD ( [Net Profit], 'Date'[Date] )", False),
    ("Closing Balance", "4 Cash and balances", "#,0.00;(#,0.00);0.00",
     "-- opening balance + every posting up to the last date in the selection (debit +, credit -)\n"
     "VAR d = MAX ( 'Date'[Date] )\nRETURN\n    SUM ( Account[Opening Balance] )\n"
     "        + CALCULATE ( SUM ( Ledger[Amount] ), REMOVEFILTERS ( 'Date' ), 'Date'[Date] <= d )", False),
    ("Cash Balance", "4 Cash and balances", "#,0", "CALCULATE ( [Closing Balance], KEEPFILTERS ( Account[Is Cash] = TRUE () ) )", False),
    ("Cash Movement", "4 Cash and balances", "#,0;(#,0);0",
     "CALCULATE ( SUM ( Ledger[Amount] ), KEEPFILTERS ( Account[Is Cash] = TRUE () ) )", False),
    ("In Transit", "4 Cash and balances", "#,0.00", "CALCULATE ( [Closing Balance], KEEPFILTERS ( Account[Account No] = 1090 ) )", False),
    ("Statement Balance", "4 Cash and balances", "#,0.00;(#,0.00);0.00",
     "CALCULATE ( SUM ( Statements[Ending Balance] ), LASTDATE ( 'Date'[Date] ) )", False),
    ("Books Balance", "4 Cash and balances", "#,0.00;(#,0.00);0.00",
     "IF ( NOT ISBLANK ( [Statement Balance] ), [Closing Balance] )", False),
    ("Difference vs Statement", "4 Cash and balances", "#,0.00;(#,0.00);0.00",
     "IF ( NOT ISBLANK ( [Statement Balance] ), [Closing Balance] - [Statement Balance] )", False),
    ("Review Amount", "5 Checks", "#,0.00", "CALCULATE ( SUM ( Ledger[Amount] ), KEEPFILTERS ( Ledger[Account No] = 6990 ) )", False),
    # ---- checks: each helper ignores every slicer
    ("Chk Opening Diff", "5 Checks", "#,0.00", "CALCULATE ( SUM ( Account[Opening Balance] ), REMOVEFILTERS () )", True),
    ("Chk Debits", "5 Checks", "#,0.00", "CALCULATE ( SUM ( Ledger[Debit] ), REMOVEFILTERS () )", True),
    ("Chk Credits", "5 Checks", "#,0.00", "CALCULATE ( SUM ( Ledger[Credit] ), REMOVEFILTERS () )", True),
    ("Chk TB Months Off", "5 Checks", "0",
     "VAR ends = CALCULATETABLE ( VALUES ( 'Date'[Month End] ), REMOVEFILTERS () )\n"
     "VAR opening = CALCULATE ( SUM ( Account[Opening Balance] ), REMOVEFILTERS () )\n"
     "RETURN\n    COUNTROWS ( FILTER ( ends, VAR d = 'Date'[Month End] RETURN\n"
     "        ABS ( opening + CALCULATE ( SUM ( Ledger[Amount] ), REMOVEFILTERS (), Ledger[Date] <= d ) ) > 0.005 ) ) + 0", True),
    ("Chk Month Ends", "5 Checks", "0", "CALCULATE ( DISTINCTCOUNT ( 'Date'[Month End] ), REMOVEFILTERS () )", True),
    ("Chk Rows Off", "5 Checks", "0",
     "COUNTROWS ( FILTER ( CALCULATETABLE ( 'Source Files', REMOVEFILTERS () ),\n"
     "    VAR r = 'Source Files'[Register No]\n    VAR n = 'Source Files'[Rows In File]\n    RETURN\n"
     "        CALCULATE ( COUNTROWS ( Ledger ), REMOVEFILTERS (), Ledger[Side] = \"Register\", Ledger[Register No] = r ) <> n ) ) + 0", True),
    ("Chk Rows In Files", "5 Checks", "0", "CALCULATE ( SUM ( 'Source Files'[Rows In File] ), REMOVEFILTERS () )", True),
    ("Chk Totals Off", "5 Checks", "0",
     "COUNTROWS ( FILTER ( CALCULATETABLE ( 'Source Files', REMOVEFILTERS () ),\n"
     "    VAR r = 'Source Files'[Register No]\n    VAR t = 'Source Files'[Net Amount In File]\n    RETURN\n"
     "        ABS ( CALCULATE ( SUM ( Ledger[Amount] ), REMOVEFILTERS (), Ledger[Side] = \"Register\", Ledger[Register No] = r ) - t ) > 0.005 ) ) + 0", True),
    ("Chk Statements Off", "5 Checks", "0",
     "COUNTROWS ( FILTER ( CALCULATETABLE ( Statements, REMOVEFILTERS () ),\n"
     "    VAR a = Statements[Account No]\n    VAR d = Statements[Statement Date]\n"
     "    VAR books = CALCULATE ( SUM ( Account[Opening Balance] ), REMOVEFILTERS (), Account[Account No] = a )\n"
     "        + CALCULATE ( SUM ( Ledger[Amount] ), REMOVEFILTERS (), Ledger[Account No] = a, Ledger[Date] <= d )\n"
     "    RETURN ABS ( books - Statements[Ending Balance] ) > 0.005 ) ) + 0", True),
    ("Chk Statements", "5 Checks", "0", "CALCULATE ( COUNTROWS ( Statements ), REMOVEFILTERS () )", True),
    ("Chk Checking Books", "5 Checks", "#,0.00",
     "CALCULATE ( SUM ( Account[Opening Balance] ), REMOVEFILTERS (), Account[Account No] = 1000 )\n"
     "    + CALCULATE ( SUM ( Ledger[Amount] ), REMOVEFILTERS (), Ledger[Account No] = 1000 )", True),
    ("Chk Checking Bank", "5 Checks", "#,0.00",
     "CALCULATE ( SUM ( 'Source Files'[Bank Running Balance] ), REMOVEFILTERS (), 'Source Files'[Register No] = 1000 )", True),
    ("Chk Orphans", "5 Checks", "0",
     "COUNTROWS ( FILTER ( CALCULATETABLE ( Ledger, REMOVEFILTERS () ), ISBLANK ( RELATED ( Account[Account] ) ) ) ) + 0", True),
    ("Chk Review Lines", "5 Checks", "0",
     "CALCULATE ( COUNTROWS ( Ledger ), REMOVEFILTERS (), Ledger[Account No] = 6990 ) + 0", True),
    ("Check Status", "5 Checks", None,
     "VAR ok = SWITCH ( SELECTEDVALUE ( Checks[No] ),\n"
     "    1, ABS ( [Chk Opening Diff] ) < 0.005,\n"
     "    2, ABS ( [Chk Debits] - [Chk Credits] ) < 0.005,\n"
     "    3, [Chk TB Months Off] = 0,\n"
     "    4, [Chk Rows Off] = 0,\n"
     "    5, [Chk Totals Off] = 0,\n"
     "    6, [Chk Statements Off] = 0 && [Chk Statements] > 0,\n"
     "    7, ABS ( [Chk Checking Books] - [Chk Checking Bank] ) < 0.005,\n"
     "    8, [Chk Orphans] = 0,\n"
     "    9, [Chk Review Lines] = 0\n)\n"
     "RETURN\n    IF ( HASONEVALUE ( Checks[No] ), IF ( ok, \"PASS\", IF ( SELECTEDVALUE ( Checks[No] ) = 9, \"REVIEW\", \"FAIL\" ) ) )", False),
    ("Check Detail", "5 Checks", None,
     "SWITCH ( SELECTEDVALUE ( Checks[No] ),\n"
     "    1, \"Difference \" & FORMAT ( [Chk Opening Diff], \"#,0.00\" ),\n"
     "    2, \"Debits \" & FORMAT ( [Chk Debits], \"#,0.00\" ) & \" = credits \" & FORMAT ( [Chk Credits], \"#,0.00\" ),\n"
     "    3, FORMAT ( [Chk Month Ends] - [Chk TB Months Off], \"0\" ) & \" of \" & FORMAT ( [Chk Month Ends], \"0\" ) & \" month ends balance to 0.00\",\n"
     "    4, FORMAT ( [Chk Rows In Files], \"0\" ) & \" lines in 3 files, \" & FORMAT ( [Chk Rows Off], \"0\" ) & \" files differ\",\n"
     "    5, FORMAT ( 3 - [Chk Totals Off], \"0\" ) & \" of 3 files: net amount = posted\",\n"
     "    6, FORMAT ( [Chk Statements] - [Chk Statements Off], \"0\" ) & \" of \" & FORMAT ( [Chk Statements], \"0\" ) & \" account-months agree to the cent\",\n"
     "    7, \"Books \" & FORMAT ( [Chk Checking Books], \"#,0.00\" ) & \", bank \" & FORMAT ( [Chk Checking Bank], \"#,0.00\" ),\n"
     "    8, FORMAT ( [Chk Orphans], \"0\" ) & \" lines without an account\",\n"
     "    9, FORMAT ( [Chk Review Lines], \"0\" ) & \" lines, \" & FORMAT ( CALCULATE ( [Review Amount], REMOVEFILTERS () ), \"#,0.00\" ) & \" on 6990 Uncategorized\"\n)", False),
    ("Check Color", "5 Checks", None,
     'SWITCH ( [Check Status], "PASS", "#5BBF8A", "FAIL", "#E5735F", "REVIEW", "#D9A441" )', True),
    ("Checks Passed", "5 Checks", None,
     'VAR t = ADDCOLUMNS ( ALL ( Checks[No] ), "@s", [Check Status] )\n'
     'RETURN FORMAT ( COUNTROWS ( FILTER ( t, [@s] = "PASS" ) ), "0" ) & " of " & FORMAT ( COUNTROWS ( FILTER ( t, [@s] <> "REVIEW" ) ), "0" ) & " PASS"', False),
    ("Checks Failed", "5 Checks", "0",
     'VAR t = ADDCOLUMNS ( ALL ( Checks[No] ), "@s", [Check Status] )\nRETURN COUNTROWS ( FILTER ( t, [@s] = "FAIL" ) ) + 0', False),
    # ---- the small line under each KPI and the Key insights panel
    ("Income Budget", "6 Report text", "#,0", 'CALCULATE ( [Budget], KEEPFILTERS ( Category[Group] = "Income" ) )', False),
    ("Expense Budget", "6 Report text", "#,0", '-CALCULATE ( [Budget], KEEPFILTERS ( Category[Group] = "Expense" ) )', False),
    ("Income vs Budget %", "6 Report text", "+0.0%;-0.0%;0.0%", "DIVIDE ( [Total Income] - [Income Budget], [Income Budget] )", False),
    ("Expense vs Budget %", "6 Report text", "+0.0%;-0.0%;0.0%", "DIVIDE ( [Total Expense] - [Expense Budget], [Expense Budget] )", False),
    ("Budget Margin %", "6 Report text", "0.0%", "DIVIDE ( [Budget], [Income Budget] )", False),
    ("Cash Start", "6 Report text", "#,0",
     "-- cash (checking + savings) the day before the first date in the selection\n"
     "VAR d0 = MIN ( 'Date'[Date] )\nRETURN\n    CALCULATE (\n        SUM ( Account[Opening Balance] )\n"
     "            + CALCULATE ( SUM ( Ledger[Amount] ), 'Date'[Date] < d0 ),\n"
     "        KEEPFILTERS ( Account[Is Cash] = TRUE () ),\n        REMOVEFILTERS ( 'Date' )\n    )", False),
    ("Cash Change", "6 Report text", "+#,0;-#,0;0", "[Cash Balance] - [Cash Start]", False),
    ("Income Note", "6 Report text", None,
     'VAR p = [Income vs Budget %]\nRETURN IF ( ISBLANK ( p ), "no budget", IF ( p >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( p ), "0.0%" ) & " vs budget" )', True),
    ("Income Note Color", "6 Report text", None, 'IF ( [Income vs Budget %] < 0, "#E5735F", "#5BBF8A" )', True),
    ("Expense Note", "6 Report text", None,
     'VAR p = [Expense vs Budget %]\nRETURN IF ( ISBLANK ( p ), "no budget", IF ( p > 0, "▲ " & FORMAT ( p, "0.0%" ) & " over budget", "▼ " & FORMAT ( -p, "0.0%" ) & " under budget" ) )', True),
    ("Expense Note Color", "6 Report text", None, 'IF ( [Expense vs Budget %] > 0, "#E5735F", "#5BBF8A" )', True),
    ("Profit Note", "6 Report text", None,
     '"budget " & FORMAT ( [Budget], "#,0;(#,0);0" ) & "  ·  " & FORMAT ( [Variance], "+#,0;-#,0;0" )', True),
    ("Margin Note", "6 Report text", None, '"budget " & FORMAT ( [Budget Margin %], "0.0%;-0.0%;0.0%" )', True),
    ("Margin Note Color", "6 Report text", None, 'IF ( [Net Margin %] < [Budget Margin %], "#E5735F", "#5BBF8A" )', True),
    ("MoM Note", "6 Report text", None,
     "VAR m = MAX ( 'Date'[Month Start] )\nRETURN\n    IF ( ISBLANK ( [Income Previous Month] ), \"first month: nothing to compare\",\n"
     "        FORMAT ( m, \"mmm\" ) & \" \" & FORMAT ( [Income Last Month], \"#,0\" ) & \" vs \" & FORMAT ( EDATE ( m, -1 ), \"mmm\" ) & \" \" & FORMAT ( [Income Previous Month], \"#,0\" ) )", True),
    ("Cash Note", "6 Report text", None,
     'VAR c = [Cash Change]\nRETURN IF ( c >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( c ), "#,0" ) & " in the period"', True),
    ("Cash Note Color", "6 Report text", None, 'IF ( [Cash Change] < 0, "#E5735F", "#5BBF8A" )', True),
    ("Insight", "6 Report text", None,
     "-- one sentence per row of the Insights table, recalculated for the month and account picked in the slicers\n"
     "VAR byAccount = FILTER ( ADDCOLUMNS ( VALUES ( Account[Account Label] ), \"@v\", [Variance] ), NOT ISBLANK ( [@v] ) )\n"
     "VAR worst = TOPN ( 1, byAccount, [@v], ASC )\nVAR best = TOPN ( 1, byAccount, [@v], DESC )\n"
     "RETURN\n    SWITCH ( SELECTEDVALUE ( Insights[No] ),\n"
     "        1, \"Net profit \" & FORMAT ( [Net Profit], \"#,0;(#,0);0\" ) & \" against a budget of \" & FORMAT ( [Budget], \"#,0;(#,0);0\" )\n"
     "            & \" (\" & FORMAT ( [Variance], \"+#,0;-#,0;0\" ) & \")\",\n"
     "        2, \"Income \" & FORMAT ( [Income vs Budget %], \"+0.0%;-0.0%;0.0%\" ) & \" vs budget, costs \" & FORMAT ( [Expense vs Budget %], \"+0.0%;-0.0%;0.0%\" ),\n"
     "        3, MAXX ( worst, Account[Account Label] & \": \" & FORMAT ( [@v], \"+#,0;-#,0;0\" ) & \" vs budget\" ),\n"
     "        4, MAXX ( best, Account[Account Label] & \": \" & FORMAT ( [@v], \"+#,0;-#,0;0\" ) & \" vs budget\" ),\n"
     "        5, \"Cash \" & IF ( [Cash Change] >= 0, \"up \", \"down \" ) & FORMAT ( ABS ( [Cash Change] ), \"#,0\" ) & \" to \" & FORMAT ( [Cash Balance], \"#,0\" ),\n"
     "        6, \"Lowest cash \" & FORMAT ( [Lowest Forecast Cash], \"#,0\" ) & \" in \" & [Lowest Forecast Week] & \", \"\n"
     "            & FORMAT ( ABS ( [Headroom] ), \"#,0\" ) & IF ( [Headroom] >= 0, \" above\", \" below\" ) & \" a buffer of \" & FORMAT ( [Cash Buffer Value], \"#,0\" )\n    )", False),
    ("Insight Color", "6 Report text", None,
     'SWITCH ( SELECTEDVALUE ( Insights[No] ), 1, [Variance Color], 3, "#E5735F", 4, "#5BBF8A", 5, [Cash Note Color], 6, [Headroom Color], "#ECE6DA" )', True),
    # ---- revenue and costs
    ("Gross Margin %", "7 Revenue and costs", "0.0%;-0.0%;0.0%", "DIVIDE ( [Gross Profit], [Total Income] )", False),
    ("Budget Gross Margin %", "7 Revenue and costs", "0.0%",
     'DIVIDE ( [Income Budget] + CALCULATE ( [Budget], KEEPFILTERS ( Category[Category Key] = "COGS" ) ), [Income Budget] )', False),
    ("Gross Margin Note", "7 Revenue and costs", None, '"budget " & FORMAT ( [Budget Gross Margin %], "0.0%" )', True),
    ("Gross Margin Note Color", "7 Revenue and costs", None, 'IF ( [Gross Margin %] < [Budget Gross Margin %], "#E5735F", "#5BBF8A" )', True),
    ("Counter Sales", "7 Revenue and costs", "#,0", "CALCULATE ( [Total Income], KEEPFILTERS ( Account[Account No] = 4000 ) )", False),
    ("Wholesale Sales", "7 Revenue and costs", "#,0", "CALCULATE ( [Total Income], KEEPFILTERS ( Account[Account No] = 4100 ) )", False),
    ("Wholesale Share %", "7 Revenue and costs", "0.0%", "DIVIDE ( [Wholesale Sales], [Total Income] )", False),
    ("Wholesale Note", "7 Revenue and costs", None, '"wholesale " & FORMAT ( [Wholesale Sales], "#,0" ) & " of " & FORMAT ( [Total Income], "#,0" )', True),
    ("Counter Sales per Day", "7 Revenue and costs", "#,0",
     "-- calendar days, days without a deposit count as 0\nAVERAGEX ( VALUES ( 'Date'[Date] ), [Counter Sales] + 0 )", False),
    ("Counter Day Note", "7 Revenue and costs", None,
     '"best day " & FORMAT ( MAXX ( VALUES ( \'Date\'[Date] ), [Counter Sales] ), "#,0" )', True),
    ("Payroll % of Income", "7 Revenue and costs", "0.0%",
     "DIVIDE ( CALCULATE ( [Total Expense], KEEPFILTERS ( Account[Account No] = 6000 ) ), [Total Income] )", False),
    ("Payroll Budget %", "7 Revenue and costs", "0.0%",
     "DIVIDE ( -CALCULATE ( [Budget], KEEPFILTERS ( Account[Account No] = 6000 ) ), [Income Budget] )", False),
    ("Payroll Note", "7 Revenue and costs", None, '"budget " & FORMAT ( [Payroll Budget %], "0.0%" )', True),
    ("Payroll Note Color", "7 Revenue and costs", None, 'IF ( [Payroll % of Income] > [Payroll Budget %], "#E5735F", "#5BBF8A" )', True),
    ("Spend", "7 Revenue and costs", "#,0",
     '-- paid to suppliers and staff (bank and card lines on cost accounts; depreciation is not paid)\n'
     'CALCULATE ( SUM ( Ledger[Amount] ), KEEPFILTERS ( Ledger[Side] = "Account" ), KEEPFILTERS ( Category[Group] = "Expense" ) )', False),
    ("Payments", "7 Revenue and costs", "0",
     'CALCULATE ( COUNTROWS ( Ledger ), KEEPFILTERS ( Ledger[Side] = "Account" ), KEEPFILTERS ( Category[Group] = "Expense" ) )', False),
    ("Average Payment", "7 Revenue and costs", "#,0", "DIVIDE ( [Spend], [Payments] )", False),
    ("Share of Spend %", "7 Revenue and costs", "0.0%", "DIVIDE ( [Spend], CALCULATE ( [Spend], REMOVEFILTERS ( Ledger[Payee] ) ) )", False),
    # ---- 13-week cash forecast
    ("Cash Today", "8 Cash forecast", "#,0", "-- cash after the last bank line, whatever the slicers say\nCALCULATE ( [Cash Balance], REMOVEFILTERS ( 'Date' ) )", False),
    ("Forecast Amount", "8 Cash forecast", "#,0;(#,0);0", "SUM ( Forecast[Amount] )", False),
    ("Forecast In", "8 Cash forecast", "#,0", "CALCULATE ( SUM ( Forecast[Amount] ), KEEPFILTERS ( Forecast[Amount] > 0 ) )", False),
    ("Forecast Out", "8 Cash forecast", "#,0;(#,0);0", "CALCULATE ( SUM ( Forecast[Amount] ), KEEPFILTERS ( Forecast[Amount] < 0 ) )", False),
    ("Forecast Balance", "8 Cash forecast", "#,0",
     "-- cash today + every expected receipt and payment up to the end of the week\nVAR w = MAX ( 'Forecast Week'[Week No] )\nRETURN\n"
     "    [Cash Today] + CALCULATE ( SUM ( Forecast[Amount] ), REMOVEFILTERS ( 'Forecast Week' ), 'Forecast Week'[Week No] <= w )", False),
    ("Forecast Closing", "8 Cash forecast", "#,0", "CALCULATE ( [Forecast Balance], REMOVEFILTERS ( 'Forecast Week' ) )", False),
    ("Lowest Forecast Cash", "8 Cash forecast", "#,0", "MINX ( ALL ( 'Forecast Week'[Week No] ), [Forecast Balance] )", False),
    ("Lowest Forecast Week", "8 Cash forecast", None,
     "VAR t = ADDCOLUMNS ( ALL ( 'Forecast Week'[Week No] ), \"@b\", [Forecast Balance] )\nVAR m = MINX ( t, [@b] )\n"
     "VAR w = MAXX ( FILTER ( t, [@b] = m ), 'Forecast Week'[Week No] )\n"
     "RETURN LOOKUPVALUE ( 'Forecast Week'[Week], 'Forecast Week'[Week No], w )", False),
    ("Lowest Note", "8 Cash forecast", None, '"in " & [Lowest Forecast Week]', True),
    ("Cash Buffer Value", "8 Cash forecast", "#,0", "SELECTEDVALUE ( 'Cash Buffer'[Cash Buffer], 10000 )", False),
    ("Buffer Line", "8 Cash forecast", "#,0", "IF ( HASONEVALUE ( 'Forecast Week'[Week No] ), [Cash Buffer Value] )", False),
    ("Headroom", "8 Cash forecast", "+#,0;-#,0;0", "[Lowest Forecast Cash] - [Cash Buffer Value]", False),
    ("Headroom Color", "8 Cash forecast", None, 'IF ( [Headroom] < 0, "#E5735F", "#5BBF8A" )', True),
    ("Headroom Note", "8 Cash forecast", None, '"over a buffer of " & FORMAT ( [Cash Buffer Value], "#,0" )', True),
    ("Weeks Below Buffer", "8 Cash forecast", "0",
     "COUNTROWS ( FILTER ( ALL ( 'Forecast Week'[Week No] ), [Forecast Balance] < [Cash Buffer Value] ) ) + 0", False),
    ("Weeks Below Color", "8 Cash forecast", None, 'IF ( [Weeks Below Buffer] > 0, "#E5735F", "#5BBF8A" )', True),
    ("Weeks Below Note", "8 Cash forecast", None, '"of 13 weeks end under the buffer"', True),
    ("Today Note", "8 Cash forecast", None, '"checking + savings at " & FORMAT ( CALCULATE ( MAX ( Ledger[Date] ), REMOVEFILTERS () ), "d mmm yyyy" )', True),
    # ---- ledger page
    ("Balance Start", "9 Ledger", "#,0.00;(#,0.00);0.00",
     "VAR d0 = MIN ( 'Date'[Date] )\nRETURN\n    CALCULATE ( SUM ( Account[Opening Balance] ) + CALCULATE ( SUM ( Ledger[Amount] ), 'Date'[Date] < d0 ), REMOVEFILTERS ( 'Date' ) )", False),
    ("Balance Movement", "9 Ledger", "+#,0.00;-#,0.00;0.00", "[Closing Balance] - [Balance Start]", False),
    ("Ledger Debit", "9 Ledger", "#,0.00;(#,0.00);\"\"", "SUM ( Ledger[Debit] )", False),
    ("Ledger Credit", "9 Ledger", "#,0.00;(#,0.00);\"\"", "SUM ( Ledger[Credit] )", False),
    ("Running Balance", "9 Ledger", "#,0.00;(#,0.00);0.00",
     "-- balance of the picked account after this line (debit +): opening + every line up to this date and line\n"
     "VAR lineDate = MAX ( Ledger[Date] )\nVAR lineId = SELECTEDVALUE ( Ledger[Line ID] )\nRETURN\n"
     "    IF ( HASONEVALUE ( Account[Account No] ) && NOT ISBLANK ( lineId ),\n"
     "        SUM ( Account[Opening Balance] )\n"
     "            -- ALLEXCEPT keeps the account (picked by number or by label) and drops every other filter (month, payee, line)\n"
     "            + CALCULATE ( SUM ( Ledger[Amount] ), ALLEXCEPT ( Ledger, Account[Account No], Account[Account Label] ),\n"
     "                FILTER ( ALL ( Ledger[Date], Ledger[Line ID] ),\n"
     "                    Ledger[Date] < lineDate || ( Ledger[Date] = lineDate && Ledger[Line ID] <= lineId ) ) ) )", False),
]

# ====================================================================================== tables (TMDL)
# column: name, dataType, formatString, summarizeBy, extra lines
TABLES = {
    "_Measures": dict(src=MEASURES_SRC, cols=[("Measure Group", "string", None, "none", ["isHidden"])], measures=True),
    "Ledger": dict(src=LEDGER, desc="One row per ledger line: each bank or card line posts to its register and to the account its rule chose; plus journal entries.", cols=[
        ("Line ID", "string", None, "none", []), ("Side", "string", None, "none", []),
        ("Date", "dateTime", "d mmm yyyy", "none", []), ("Txn Date", "dateTime", "d mmm yyyy", "none", []),
        ("Source", "string", None, "none", []), ("Register No", "int64", "0", "none", []),
        ("Description", "string", None, "none", []), ("Clean Description", "string", None, "none", []),
        ("Rule Keyword", "string", None, "none", []), ("Payee", "string", None, "none", []), ("Account No", "int64", "0", "none", []),
        ("Offset Account No", "int64", "0", "none", []), ("Amount", "decimal", "#,0.00", "sum", []),
        ("Debit", "decimal", "#,0.00", "sum", []), ("Credit", "decimal", "#,0.00", "sum", []),
        ("Cash Flow Line", "string", None, "none", ["sortByColumn: 'Cash Flow Order'"]),
        ("Cash Flow Order", "int64", "0", "none", ["isHidden"])]),
    "Account": dict(src=ACCOUNT, desc="Chart of accounts with opening balances at 31 Jul 2026 (debit +).", cols=[
        ("Account No", "int64", "0", "none", ["isKey"]), ("Account", "string", None, "none", []),
        ("Type", "string", None, "none", []), ("Opening Balance", "decimal", "#,0.00", "sum", []),
        ("Statement Name", "string", None, "none", ["isHidden"]),
        ("Account Label", "string", None, "none", ["sortByColumn: 'Account No'"]),
        ("Is Cash", "boolean", None, "none", []), ("Category Key", "string", None, "none", []), ("Cost Group", "string", None, "none", [])]),
    "Category": dict(src=CATEGORY, desc="Groups accounts into the lines of the P&L and balance sheet.", cols=[
        ("Category Key", "string", None, "none", ["isKey"]), ("Category", "string", None, "none", ["sortByColumn: Sort"]),
        ("Group", "string", None, "none", []), ("Statement", "string", None, "none", []), ("Sort", "int64", "0", "none", ["isHidden"])]),
    "Date": dict(src=DATE, desc="Calendar: one row per day of the months in the exports.", dataCategory="Time", cols=[
        ("Date", "dateTime", "d mmm yyyy", "none", ["isKey"]), ("Year", "int64", "0", "none", []),
        ("Month No", "int64", "0", "none", ["isHidden"]), ("Month", "string", None, "none", ["sortByColumn: 'Month No'"]),
        ("Month Start", "dateTime", "d mmm yyyy", "none", []), ("Month End", "dateTime", "d mmm yyyy", "none", []),
        ("Day", "int64", "0", "none", []), ("Weekday", "string", None, "none", ["sortByColumn: 'Weekday No'"]),
        ("Weekday No", "int64", "0", "none", ["isHidden"])]),
    "Budget": dict(src=BUDGET, desc="Budget per account per month; Budget Amount carries the P&L sign (income +, costs -).", cols=[
        ("Account No", "int64", "0", "none", []), ("Month", "dateTime", "mmm yyyy", "none", []),
        ("Budget As Typed", "decimal", "#,0.00", "sum", []), ("Budget Amount", "decimal", "#,0.00", "sum", [])]),
    "Statements": dict(src=STATEMENTS, desc="Month-end balance printed on each bank and card statement.", cols=[
        ("Account No", "int64", "0", "none", []), ("Statement Date", "dateTime", "d mmm yyyy", "none", []),
        ("Printed Balance", "decimal", "#,0.00", "sum", []), ("Ending Balance", "decimal", "#,0.00", "sum", [])]),
    "Source Files": dict(src=SOURCE_FILES, desc="Control totals read straight from each export file.", cols=[
        ("Register No", "int64", "0", "none", []), ("File", "string", None, "none", []),
        ("Rows In File", "int64", "0", "sum", []), ("Net Amount In File", "decimal", "#,0.00", "sum", []),
        ("Bank Running Balance", "decimal", "#,0.00", "sum", [])]),
    "Checks": dict(src=CHECKS, desc="The list of checks shown on the Checks page.", cols=[
        ("No", "int64", "0", "none", []), ("Check", "string", None, "none", ["sortByColumn: No"]),
        ("Why it matters", "string", None, "none", [])]),
    "Cash Plan": dict(src=CASH_PLAN_TABLE, desc="The owner's assumptions for the 13-week cash forecast (cash-plan.csv) and the amount each one gives.", cols=[
        ("Order", "int64", "0", "none", []), ("Item", "string", None, "none", ["sortByColumn: Order"]),
        ("Cash Flow Line", "string", None, "none", []), ("How It Is Estimated", "string", None, "none", []),
        ("When", "string", None, "none", []), ("Per Time", "decimal", "#,0;(#,0);0", "none", []),
        ("Times", "int64", "0", "sum", []), ("Total 13 Weeks", "decimal", "#,0;(#,0);0", "sum", []), ("Note", "string", None, "none", [])]),
    "Forecast": dict(src=FORECAST, desc="Every expected receipt and payment in the next 13 weeks (money in +, out -).", cols=[
        ("Order", "int64", "0", "none", ["isHidden"]), ("Item", "string", None, "none", ["sortByColumn: Order"]),
        ("Cash Flow Line", "string", None, "none", []), ("Date", "dateTime", "d mmm yyyy", "none", []),
        ("Amount", "decimal", "#,0;(#,0);0", "sum", []), ("Week No", "int64", "0", "none", [])]),
    "Forecast Week": dict(src=FORECAST_WEEK, desc="The 13 weeks after the last bank line.", cols=[
        ("Week No", "int64", "0", "none", ["isKey"]), ("Week Start", "dateTime", "d mmm yyyy", "none", []),
        ("Week End", "dateTime", "d mmm yyyy", "none", []), ("Week", "string", None, "none", ["sortByColumn: 'Week No'"])]),
    "Cash Buffer": dict(src=CASH_BUFFER, desc="What-if values for the minimum cash buffer.", cols=[
        ("Cash Buffer", "int64", "#,0", "none", [])]),
    "Insights": dict(src=INSIGHTS, desc="The rows of the Key insights panel; the sentence itself is the Insight measure.", cols=[
        ("No", "int64", "0", "none", []), ("Topic", "string", None, "none", ["sortByColumn: No"])]),
}
RELATIONSHIPS = [  # many side -> one side
    ("Ledger", "Date", "Date", "Date"), ("Ledger", "Account No", "Account", "Account No"),
    ("Account", "Category Key", "Category", "Category Key"),
    ("Budget", "Month", "Date", "Date"), ("Budget", "Account No", "Account", "Account No"),
    ("Statements", "Statement Date", "Date", "Date"), ("Statements", "Account No", "Account", "Account No"),
    ("Forecast", "Week No", "Forecast Week", "Week No"),
]
EXPRESSIONS = [  # shared queries, not loaded: name, M, result type
    ("Norm", NORM, "Function"), ("Money", MONEY, "Function"),
    ("Checking_Clean", CHECKING, "Table"), ("Savings_Clean", SAVINGS, "Table"), ("Card_Clean", CARD, "Table"),
    ("Rules", RULES, "Table"), ("AccountsRaw", ACCOUNTS_RAW, "Table"), ("Adjustments", ADJ, "Table"),
    ("BankLines", BANKLINES, "Table"), ("CashPlanCalc", CASHPLAN, "Table"),
]


def table_tmdl(name, t):
    out = []
    if t.get("desc"):
        out.append(f"/// {t['desc']}")
    out += [f"table {q(name)}", f"\tlineageTag: {tag('t', name)}"]
    if t.get("dataCategory"):
        out.append(f"\tdataCategory: {t['dataCategory']}")
    out.append("")
    if t.get("measures"):
        for m, folder, fmt, expr, hidden in MEASURES:
            if "\n" in expr:
                out.append(f"\tmeasure {q(m)} =")
                out.append(indent(expr, 3))
            else:
                out.append(f"\tmeasure {q(m)} = {expr}")
            if fmt:
                out.append(f"\t\tformatString: {fmt}")
            out.append(f"\t\tdisplayFolder: {folder}")
            if hidden:
                out.append("\t\tisHidden")
            out.append(f"\t\tlineageTag: {tag('m', m)}")
            out.append("")
    for c, dt, fmt, summ, extra in t["cols"]:
        out.append(f"\tcolumn {q(c)}")
        out.append(f"\t\tdataType: {dt}")
        if fmt:
            out.append(f"\t\tformatString: {fmt}")
        for e in extra:
            out.append(f"\t\t{e}")
        out.append(f"\t\tlineageTag: {tag('c', name, c)}")
        out.append(f"\t\tsummarizeBy: {summ}")
        out.append(f"\t\tsourceColumn: {c}")
        out.append("")
        out.append("\t\tannotation SummarizationSetBy = Automatic")
        out.append("")
    out.append(f"\tpartition {q(name)} = m")
    out.append("\t\tmode: import")
    out.append("\t\tsource =")
    out.append(indent(t["src"], 4))
    out.append("")
    out.append("\tannotation PBI_ResultType = Table")
    out.append("")
    return "\n".join(out)


def build_model(data_folder):
    if SM.exists():
        shutil.rmtree(SM)
    d = SM / "definition"
    (d / "tables").mkdir(parents=True)
    (SM / "definition.pbism").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2", "settings": {}}, indent=2), encoding="utf-8")
    (d / "database.tmdl").write_text("database\n\tcompatibilityLevel: 1601\n\n", encoding="utf-8")
    order = ["DataFolder"] + [e[0] for e in EXPRESSIONS] + list(TABLES)
    model = ["model Model", "\tculture: en-US", "\tdefaultPowerBIDataSourceVersion: powerBI_V3", "\tdiscourageImplicitMeasures",
             "\tsourceQueryCulture: en-US", "\tdataAccessOptions", "\t\tlegacyRedirects", "\t\treturnErrorValuesAsNull", "",
             f"annotation PBI_QueryOrder = {json.dumps(order)}", "", "annotation __PBI_TimeIntelligenceEnabled = 0", "",
             "annotation PBI_ProTooling = [\"DevMode\"]", ""]
    model += [f"ref table {q(t)}" for t in TABLES]
    (d / "model.tmdl").write_text("\n".join(model) + "\n\n", encoding="utf-8")
    ex = [f'expression DataFolder = "{data_folder}" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]',
          f"\tlineageTag: {tag('e', 'DataFolder')}", "", "\tannotation PBI_ResultType = Text", ""]
    for name, m, rtype in EXPRESSIONS:
        ex += [f"expression {name} =", indent(m, 2), f"\tlineageTag: {tag('e', name)}", "",
               f"\tannotation PBI_ResultType = {rtype}", ""]
    (d / "expressions.tmdl").write_text("\n".join(ex) + "\n", encoding="utf-8")
    rel = []
    for ft, fc, tt, tc in RELATIONSHIPS:
        rel += [f"relationship {tag('r', ft, fc, tt)}", f"\tfromColumn: {q(ft)}.{q(fc)}", f"\ttoColumn: {q(tt)}.{q(tc)}", ""]
    (d / "relationships.tmdl").write_text("\n".join(rel) + "\n", encoding="utf-8")
    # model diagram: the fact table in the middle, dimensions around it (star schema), helper tables below
    spots = {"Date": (40, 300), "Ledger": (400, 180), "Account": (780, 300), "Category": (1120, 330),
             "Budget": (780, -40), "Statements": (400, 700), "_Measures": (40, 700), "Checks": (1120, 640),
             "Source Files": (780, 700), "Insights": (1120, 900), "Forecast": (40, 1060), "Forecast Week": (400, 1060),
             "Cash Plan": (780, 1060), "Cash Buffer": (1120, 1060)}
    nodes = [{"location": {"x": x, "y": y}, "nodeIndex": n, "nodeLineageTag": tag("t", n),
              "size": {"height": min(64 + 34 * len(TABLES[n]["cols"]), 420), "width": 260}, "zIndex": 0}
             for n, (x, y) in spots.items()]
    (SM / "diagramLayout.json").write_text(json.dumps({
        "version": "1.1.0", "diagrams": [{"ordinal": 0, "scrollPosition": {"x": 0, "y": 0}, "nodes": nodes, "name": "All tables",
                                          "zoomValue": 70, "pinKeyFieldsToTop": False, "showExtraHeaderInfo": False,
                                          "hideKeyFieldsWhenCollapsed": False, "tablesLocked": False}],
        "selectedDiagram": "All tables", "defaultDiagram": "All tables"}, indent=2), encoding="utf-8")
    for name, t in TABLES.items():
        (d / "tables" / f"{name}.tmdl").write_text(table_tmdl(name, t), encoding="utf-8")


# ====================================================================================== report (PBIR)
S = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
# "Copperleaf" dark edition: midnight background, gold hairlines, serif headings and numbers
HDR, PAGE_BG, CARD, CARD_ALT, LINE = "#0B0F13", "#12171D", "#1A2129", "#202832", "#2A333D"
GOLD, GOLD_LT, GOLD_DIM = "#C9A35B", "#E6CF9A", "#3A3020"
IVORY, MUTED, GRID = "#ECE6DA", "#8E98A3", "#262F39"
JADE, COPPER, SKY, BUDGET_GREY = "#4FA38C", "#B5754A", "#7FA7C9", "#4A5560"
GOOD, BAD, AMBER = "#5BBF8A", "#E5735F", "#D9A441"
INK = IVORY
SERIF, SANS, SANS_SB = "Georgia", "Segoe UI", "Segoe UI Semibold"
NAV_X, NAV_W = 470, 790


def lit(v):
    if isinstance(v, bool):
        s = "true" if v else "false"
    elif isinstance(v, (int, float)):
        s = f"{v}D"
    else:
        s = "'" + str(v).replace("'", "''") + "'"
    return {"expr": {"Literal": {"Value": s}}}


def color(c):
    return {"solid": {"color": lit(c)}}


def mcolor(measure):  # colour returned by a measure
    return {"solid": {"color": {"expr": mea(measure)}}}


def col(entity, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def mea(prop):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": "_Measures"}}, "Property": prop}}


def proj(field, name=None):
    kind = "Column" if "Column" in field else "Measure"
    ent, prop = field[kind]["Expression"]["SourceRef"]["Entity"], field[kind]["Property"]
    p = {"field": field, "queryRef": f"{ent}.{prop}", "nativeQueryRef": prop}
    if name:
        p["displayName"] = name
    return p


def props(**kw):
    return [{"properties": kw}]


def container(title=None, bg=CARD, border=True, transparent=False, radius=12, border_color=LINE):
    o = {"background": props(show=lit(not transparent), color=color(bg), transparency=lit(100 if transparent else 0)),
         "border": props(show=lit(border), color=color(border_color), radius=lit(radius), width=lit(1)),
         "dropShadow": props(show=lit(False))}
    if title:
        o["title"] = props(show=lit(True), text=lit(title), fontColor=color(IVORY), fontSize=lit(12), bold=lit(False),
                           fontFamily=lit(SERIF))
    else:
        o["title"] = props(show=lit(False))
    return o


class Page:
    def __init__(self, name, display):
        self.name, self.display, self.visuals = name, display, []

    def add(self, vid, x, y, w, h, visual):
        self.visuals.append({"$schema": f"{S}/visualContainer/2.14.0/schema.json", "name": vid,
                             "position": {"x": x, "y": y, "z": len(self.visuals) * 1000, "width": w, "height": h,
                                          "tabOrder": len(self.visuals) * 1000}, "visual": visual})


def textbox(page, vid, x, y, w, h, paras, bg=CARD, border=False, transparent=False, radius=12, pad=None, border_color=LINE):
    vco = container(bg=bg, border=border, transparent=transparent, radius=radius, border_color=border_color)
    if pad is not None:
        vco["padding"] = props(top=lit(pad[0]), left=lit(pad[1]), bottom=lit(0), right=lit(0))
    page.add(vid, x, y, w, h, {"visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": paras}}]},
                               "visualContainerObjects": vco})


def run(value, size, colour, font=SANS, bold=False):
    st = {"fontFamily": font, "fontSize": f"{size}pt", "color": colour}
    if bold:
        st["fontWeight"] = "bold"
    return {"value": value, "textStyle": st}


def header(page, title, sub):
    """Brand and page title on the left, page buttons on the right, a gold hairline underneath."""
    textbox(page, f"{page.name}_hdr", 0, 0, 1280, 60, [
        {"textRuns": [run("C O P P E R L E A F", 8, GOLD, SANS_SB), run("     ", 8, MUTED), run(title, 18, IVORY, SERIF)]},
        {"textRuns": [run(sub, 8, MUTED)]}], bg=HDR, radius=0, pad=(4, 22))
    textbox(page, f"{page.name}_rule", 0, 60, 1280, 2, [{"textRuns": [run(" ", 1, GOLD)]}], bg=GOLD, radius=0, pad=(0, 0))
    page.add(f"{page.name}_nav", NAV_X, 15, NAV_W, 30, {
        "visualType": "pageNavigator",
        "objects": {
            "fill": [{"properties": {"show": lit(True), "fillColor": color(HDR), "transparency": lit(0)}, "selector": {"id": "default"}},
                     {"properties": {"show": lit(True), "fillColor": color("#1B222A"), "transparency": lit(0)}, "selector": {"id": "hover"}},
                     {"properties": {"show": lit(True), "fillColor": color(GOLD_DIM), "transparency": lit(0)}, "selector": {"id": "selected"}}],
            "text": [{"properties": {"show": lit(True), "fontColor": color("#AEB6BF"), "fontSize": lit(9), "fontFamily": lit(SANS)}, "selector": {"id": "default"}},
                     {"properties": {"fontColor": color(GOLD_LT)}, "selector": {"id": "selected"}}],
            "outline": [{"properties": {"show": lit(True), "lineColor": color("#2A333D"), "weight": lit(1)}, "selector": {"id": "default"}},
                        {"properties": {"show": lit(True), "lineColor": color(GOLD), "weight": lit(1)}, "selector": {"id": "selected"}}],
            "shape": props(tileShape=lit("rectangleRounded"), roundEdge=lit(15))},
        "visualContainerObjects": container(bg=HDR, border=False, transparent=True)})


def slicer(page, vid, x, y, w, field, title, sync=None, h=64, default=None):
    v = {"visualType": "slicer",
         "query": {"queryState": {"Values": {"projections": [proj(field)]}}},
         "objects": {"data": props(mode=lit("Dropdown")),
                     "header": props(show=lit(True), text=lit(title), fontColor=color(GOLD), textSize=lit(8), fontFamily=lit(SANS_SB)),
                     "items": props(textSize=lit(10), fontColor=color(IVORY), background=color(CARD_ALT), outlineColor=color(LINE))},
         "visualContainerObjects": container()}
    if default is not None:  # a pre-selected value (single select)
        kind = "Column"
        ent, prop = field[kind]["Expression"]["SourceRef"]["Entity"], field[kind]["Property"]
        v["objects"]["selection"] = props(singleSelect=lit(True))
        v["objects"]["general"] = [{"properties": {"filter": {"filter": {
            "Version": 2, "From": [{"Name": "s", "Entity": ent, "Type": 0}],
            "Where": [{"Condition": {"In": {"Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": prop}}],
                                            "Values": [[{"Literal": {"Value": f"{default}L" if isinstance(default, int) else "'" + default + "'"}}]]}}}]}}}}]
    if sync:
        v["syncGroup"] = {"groupName": sync, "fieldChanges": True, "filterChanges": True}
    page.add(vid, x, y, w, h, v)


def value_card(page, vid, x, y, w, h, measure, fill, size, label=None, font=SERIF):
    page.add(vid, x, y, w, h, {
        "visualType": "card",
        "query": {"queryState": {"Values": {"projections": [proj(mea(measure), label)]}}},
        "objects": {"labels": props(fontSize=lit(size), color=fill, labelDisplayUnits=lit(1), fontFamily=lit(font)),
                    "categoryLabels": props(show=lit(False))},
        "visualContainerObjects": container(transparent=True, border=False)})


def kpi(page, vid, x, y, w, h, measure, label, icon, note=None, note_color=None, value_color=IVORY, size=22):
    """One KPI tile: dark card, gold bar on the left, gold ring icon, small-caps label, serif value, coloured verdict line."""
    textbox(page, f"{vid}_bg", x, y, w, h, [{"textRuns": [run(" ", 6, IVORY)]}], border=True)
    textbox(page, f"{vid}_bar", x + 1, y + 16, 3, h - 32, [{"textRuns": [run(" ", 4, GOLD)]}], bg=GOLD, radius=2, pad=(0, 0))
    textbox(page, f"{vid}_ic", x + w - 40, y + 10, 28, 28,
            [{"textRuns": [run(icon, 10, GOLD, SANS_SB)], "horizontalTextAlignment": "center"}],
            bg=GOLD_DIM, radius=14, pad=(3, 0), border=True, border_color=GOLD)
    textbox(page, f"{vid}_lb", x + 14, y + 10, w - 58, 24, [{"textRuns": [run(label.upper(), 7, MUTED, SANS_SB)]}],
            transparent=True, pad=(2, 0))
    fill = color(value_color) if value_color.startswith("#") else mcolor(value_color)
    compact = h < 100
    if note:
        vy, vh = (y + 24, 34) if compact else (y + 30, 40)
    else:
        vy, vh = y + 30, h - 40
    value_card(page, f"{vid}_v", x + 8, vy, w - 16, vh, measure, fill, size - (2 if compact and note else 0), label)
    if note:
        nfill = mcolor(note_color) if note_color and not note_color.startswith("#") else color(note_color or MUTED)
        ny = y + h - (25 if compact else 30)
        value_card(page, f"{vid}_n", x + 8, ny, w - 16, 22 if compact else 26, note, nfill, 8, font=SANS_SB)


def sort(field, direction="Ascending"):
    return {"sort": [{"field": field, "direction": direction}], "isDefaultSort": False}


def chart(page, vid, x, y, w, h, vtype, title, roles, sort_def=None, objects=None):
    v = {"visualType": vtype, "query": {"queryState": {r: {"projections": [proj(*f) if isinstance(f, tuple) else proj(f) for f in fs]}
                                                       for r, fs in roles.items()}},
         "visualContainerObjects": container(title)}
    if sort_def:
        v["query"]["sortDefinition"] = sort_def
    v["objects"] = objects or {}
    if vtype in AXIS_CHARTS:  # the title says what the axes are; faint gridlines, muted axis text
        v["objects"]["categoryAxis"] = [{"properties": {"showAxisTitle": lit(False), "labelColor": color(MUTED), "fontSize": lit(8)}
                                         | v["objects"].get("categoryAxis", [{}])[0].get("properties", {})}]
        v["objects"]["valueAxis"] = [{"properties": {"showAxisTitle": lit(False), "secShowAxisTitle": lit(False), "labelColor": color(MUTED),
                                                     "fontSize": lit(8), "gridlineColor": color(GRID), "gridlineStyle": lit("solid")}
                                      | v["objects"].get("valueAxis", [{}])[0].get("properties", {})}]
    if "legend" in v["objects"]:
        v["objects"]["legend"] = [{"properties": {"labelColor": color(IVORY)} | v["objects"]["legend"][0]["properties"]}]
    if "labels" in v["objects"]:
        first = v["objects"]["labels"][0]
        v["objects"]["labels"] = [first | {"properties": {"color": color(IVORY)} | first["properties"]}] + v["objects"]["labels"][1:]
    if not v["objects"]:
        del v["objects"]
    page.add(vid, x, y, w, h, v)
    return v


def labels(size=8, units=1, precision=None, **kw):
    p = dict(show=lit(True), labelDisplayUnits=lit(units), fontSize=lit(size))
    if precision is not None:
        p["labelPrecision"] = lit(precision)
    return props(**p, **kw)


def legend(pos="TopRight"):
    return props(show=lit(True), position=lit(pos), fontSize=lit(8))


def fills(*pairs):  # (measure name, colour) per series
    return [{"properties": {"fill": color(c)}, "selector": {"metadata": f"_Measures.{m}"}} for m, c in pairs]


def cond_color(measure_queryref, color_measure):
    return {"properties": {"fontColor": {"solid": {"color": {"expr": mea(color_measure)}}}},
            "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": measure_queryref}}


def grid_objects(font=10, extra_values=None, pad=3):
    return {"values": [{"properties": {"fontSize": lit(font), "fontColorPrimary": color(IVORY), "fontColorSecondary": color(IVORY),
                                       "backColorPrimary": color(CARD), "backColorSecondary": color(CARD_ALT)}}] + (extra_values or []),
            "columnHeaders": props(fontSize=lit(font - 1), bold=lit(False), fontColor=color(GOLD), backColor=color(HDR), fontFamily=lit(SANS_SB)),
            "rowHeaders": props(fontSize=lit(font), fontColor=color(IVORY), backColor=color(CARD)),
            "grid": props(rowPadding=lit(pad), gridHorizontal=lit(True), gridHorizontalColor=color(GRID), gridVertical=lit(False),
                          outlineColor=color(LINE)),
            "total": props(fontColor=color(GOLD_LT), backColor=color(HDR)),
            "subTotals": props(fontColor=color(GOLD_LT), backColor=color("#1F1B14"))}


AXIS_CHARTS = {"lineClusteredColumnComboChart", "clusteredBarChart", "clusteredColumnChart", "lineChart", "areaChart", "waterfallChart",
               "columnChart", "stackedColumnChart"}
NO_TOTALS = {"total": props(totals=lit(False))}
MONTH, ACCT, CATG = col("Date", "Month"), col("Account", "Account Label"), col("Category", "Category")
SUB = "Copperleaf Bakery & Cafe (fictional)  ·  Aug–Oct 2026  ·  USD"


def kpi_row(page, prefix, y, h, items, x0=274, x1=1258, gap=12, size=20):
    w = (x1 - x0 - gap * (len(items) - 1)) / len(items)
    for i, it in enumerate(items):
        m, lbl, ic = it[:3]
        note, ncol, vcol = (list(it[3:]) + [None, None, IVORY])[:3]
        kpi(page, f"{prefix}{i}", round(x0 + i * (w + gap)), y, round(w), h, m, lbl, ic, note, ncol, vcol or IVORY, size)


def build_pages():
    # ---------------------------------------------------------------- 1 Overview
    p1 = Page("overview", "Overview")
    header(p1, "Finance overview", SUB)
    slicer(p1, "ov_month", 22, 74, 240, MONTH, "MONTH", "Month", h=60)
    slicer(p1, "ov_acct", 274, 74, 320, ACCT, "ACCOUNT", "Account", h=60)
    textbox(p1, "ov_hint", 606, 74, 652, 60, [
        {"textRuns": [run("Every figure on this page follows the month and account you pick.  ", 9, MUTED),
                      run("Green", 9, GOOD, bold=True), run(" = better than budget,  ", 9, MUTED),
                      run("red", 9, BAD, bold=True), run(" = worse.", 9, MUTED)]}], transparent=True, pad=(20, 6))
    kpis = [("Total Income", "Income", "$", "Income Note", "Income Note Color", IVORY),
            ("Total Expense", "Expenses", "−", "Expense Note", "Expense Note Color", IVORY),
            ("Net Profit", "Net profit", "=", "Profit Note", "Variance Color", "Net Profit Color"),
            ("Net Margin %", "Net margin", "%", "Margin Note", "Margin Note Color", IVORY),
            ("Income MoM %", "Income vs prior month", "↗", "MoM Note", MUTED, IVORY),
            ("Cash Balance", "Cash at end", "◈", "Cash Note", "Cash Note Color", IVORY)]
    kpi_row(p1, "ov_k", 146, 100, kpis, x0=22, size=22)
    chart(p1, "ov_trend", 22, 258, 410, 236, "clusteredColumnChart", "Income vs expenses by month",
          {"Category": [MONTH], "Y": [(mea("Total Income"), "Income"), (mea("Total Expense"), "Expenses")]}, sort(MONTH),
          {"labels": labels(8, 1000, 1), "legend": legend(),
           "dataPoint": fills(("Total Income", JADE), ("Total Expense", COPPER))})
    chart(p1, "ov_bridge", 444, 258, 400, 236, "waterfallChart", "Profit bridge: income to net profit",
          {"Category": [CATG], "Y": [(mea("Actual"), "Amount")]}, sort(CATG),
          {"labels": labels(8, 1000, 1),
           "sentimentColors": props(increaseFill=color(JADE), decreaseFill=color(COPPER), totalFill=color(GOLD)),
           "legend": props(show=lit(False))})
    chart(p1, "ov_exp", 856, 258, 402, 236, "clusteredBarChart", "Where the money goes",
          {"Category": [ACCT], "Y": [(mea("Total Expense"), "Expenses")]}, sort(mea("Total Expense"), "Descending"),
          {"labels": labels(8), "dataPoint": props(fill=color(GOLD)), "categoryAxis": props(maxMarginFactor=lit(45))})
    chart(p1, "ov_np", 22, 506, 560, 194, "clusteredColumnChart", "Net profit by month: actual vs budget",
          {"Category": [MONTH], "Y": [(mea("Actual"), "Actual"), (mea("Budget"), "Budget")]}, sort(MONTH),
          {"labels": labels(8), "legend": legend(), "dataPoint": fills(("Actual", GOLD), ("Budget", BUDGET_GREY))})
    chart(p1, "ov_ins", 594, 506, 664, 194, "tableEx", "Key insights, written by DAX",
          {"Values": [(col("Insights", "Topic"), "Topic"), (mea("Insight"), "What the numbers say")]},
          sort(col("Insights", "Topic")),
          objects=grid_objects(9, [cond_color("_Measures.Insight", "Insight Color")], pad=1) | NO_TOTALS)
    # ---------------------------------------------------------------- 2 P&L vs budget
    p2 = Page("pnl", "P&L vs Budget")
    header(p2, "Profit and loss vs budget", SUB + "  ·  + = better than budget")
    slicer(p2, "pl_month", 22, 74, 240, MONTH, "MONTH", "Month", h=84)
    kpi_row(p2, "pl_k", 74, 84, [("Actual", "Net profit", "=", None, None, "Net Profit Color"),
                                 ("Budget", "Budgeted net profit", "◎"),
                                 ("Variance", "Variance vs budget", "±", None, None, "Variance Color"),
                                 ("Income vs Budget %", "Income vs budget", "$", None, None, "Income Note Color")])
    mx = chart(p2, "pl_matrix", 22, 170, 652, 530, "pivotTable", "P&L by account: actual, budget, variance",
               {"Rows": [CATG, ACCT], "Values": [mea("Actual"), mea("Budget"), mea("Variance"), mea("Variance %")]},
               objects=grid_objects(9, [cond_color("_Measures.Variance", "Variance Color"), cond_color("_Measures.Variance %", "Variance Color")], pad=1))
    mx["expansionStates"] = [{"roles": ["Rows"], "levels": [{"queryRefs": ["Category.Category"], "identityKeys": [CATG], "isPinned": True},
                                                            {"queryRefs": ["Account.Account Label"], "isPinned": True}],
                              "root": {"isToggled": True}}]
    chart(p2, "pl_trend", 686, 170, 572, 256, "clusteredColumnChart", "Net profit by month: actual vs budget",
          {"Category": [MONTH], "Y": [(mea("Actual"), "Actual"), (mea("Budget"), "Budget")]}, sort(MONTH),
          {"labels": labels(8), "legend": legend(), "dataPoint": fills(("Actual", GOLD), ("Budget", BUDGET_GREY))})
    chart(p2, "pl_var", 686, 438, 572, 262, "clusteredBarChart", "Variance by account, worst first",
          {"Category": [ACCT], "Y": [mea("Variance")]}, sort(mea("Variance"), "Ascending"),
          {"labels": labels(8), "categoryAxis": props(maxMarginFactor=lit(40)),
           "dataPoint": [{"properties": {"fill": mcolor("Variance Color")}, "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}}]})
    # ---------------------------------------------------------------- 3 Revenue and costs
    p3 = Page("revenue", "Revenue & Costs")
    header(p3, "Revenue and costs", SUB)
    slicer(p3, "rc_month", 22, 74, 240, MONTH, "MONTH", "Month", h=84)
    kpi_row(p3, "rc_k", 74, 84, [("Gross Margin %", "Gross margin", "%", "Gross Margin Note", "Gross Margin Note Color"),
                                 ("Wholesale Share %", "Wholesale share of income", "◧", "Wholesale Note", MUTED),
                                 ("Counter Sales per Day", "Counter sales per day", "☕", "Counter Day Note", MUTED),
                                 ("Payroll % of Income", "Payroll as % of income", "♟", "Payroll Note", "Payroll Note Color")])
    chart(p3, "rc_stream", 22, 170, 410, 262, "columnChart", "Income by stream and month",
          {"Category": [MONTH], "Series": [ACCT], "Y": [(mea("Total Income"), "Income")]}, sort(MONTH),
          {"labels": labels(8, 1000, 1), "legend": props(show=lit(True), position=lit("Bottom"), fontSize=lit(8), showTitle=lit(False))})
    chart(p3, "rc_mix", 444, 170, 400, 262, "donutChart", "Cost mix: where each dollar of cost goes",
          {"Category": [col("Account", "Cost Group")], "Y": [(mea("Total Expense"), "Costs")]}, sort(mea("Total Expense"), "Descending"),
          {"labels": props(show=lit(True), labelStyle=lit("Category, percent of total"), color=color(IVORY), fontSize=lit(8)),
           "legend": props(show=lit(False)), "slices": props(innerRadiusRatio=lit(62))})
    chart(p3, "rc_margin", 856, 170, 402, 262, "lineChart", "Margins by month",
          {"Category": [MONTH], "Y": [(mea("Gross Margin %"), "Gross margin"), (mea("Net Margin %"), "Net margin")]}, sort(MONTH),
          {"labels": labels(8), "legend": legend(), "dataPoint": fills(("Gross Margin %", GOLD), ("Net Margin %", JADE)),
           "lineStyles": props(strokeWidth=lit(3), showMarker=lit(True), markerSize=lit(6))})
    chart(p3, "rc_wday", 22, 444, 410, 256, "clusteredColumnChart", "Card takings by deposit weekday (average day)",
          {"Category": [col("Date", "Weekday")], "Y": [(mea("Counter Sales per Day"), "Average day")]}, sort(col("Date", "Weekday")),
          {"labels": labels(8, 1), "dataPoint": props(fill=color(JADE))})
    chart(p3, "rc_payees", 444, 444, 814, 256, "tableEx", "Who gets paid: suppliers and staff, largest first",
          {"Values": [(col("Ledger", "Payee"), "Payee"), (mea("Spend"), "Paid"), (mea("Share of Spend %"), "Share"),
                      (mea("Payments"), "Payments"), (mea("Average Payment"), "Average payment")]},
          sort(mea("Spend"), "Descending"), objects=grid_objects(9, pad=2))
    # ---------------------------------------------------------------- 4 Cash and statements
    p4 = Page("cash", "Cash & Banks")
    header(p4, "Cash and bank statements", SUB)
    slicer(p4, "ca_month", 22, 74, 240, MONTH, "MONTH", "Month", h=84)
    kpi_row(p4, "ca_k", 74, 84, [("Cash Balance", "Cash at end: checking + savings", "◈"),
                                 ("Cash Change", "Change in the period", "↕", None, None, "Cash Note Color"),
                                 ("In Transit", "Transfers still in transit", "⇄")])
    chart(p4, "ca_line", 22, 170, 690, 270, "areaChart", "Cash balance day by day (checking + savings)",
          {"Category": [col("Date", "Date")], "Y": [mea("Cash Balance")]}, sort(col("Date", "Date")),
          {"dataPoint": props(fill=color(GOLD)), "categoryAxis": props(axisType=lit("Scalar")),
           "lineStyles": props(strokeWidth=lit(2))})
    chart(p4, "ca_flow", 22, 452, 690, 248, "pivotTable", "Cash flow by month (checking + savings)",
          {"Rows": [col("Ledger", "Cash Flow Line")], "Columns": [MONTH], "Values": [(mea("Cash Movement"), "Net cash")]},
          objects=grid_objects(9, pad=2))
    chart(p4, "ca_rec", 724, 170, 534, 330, "tableEx", "Books vs statements (card owed shown negative)",
          {"Values": [(ACCT, "Account"), (MONTH, "Month"), (mea("Books Balance"), "Books"), (mea("Statement Balance"), "Statement"),
                      (mea("Difference vs Statement"), "Difference")]},
          objects=grid_objects(9, pad=2) | NO_TOTALS)
    chart(p4, "ca_bridge", 724, 512, 534, 188, "waterfallChart", "Cash bridge: what moved cash",
          {"Category": [col("Ledger", "Cash Flow Line")], "Y": [(mea("Cash Movement"), "Net cash")]}, sort(col("Ledger", "Cash Flow Line")),
          {"labels": labels(8, 1000, 1),
           "sentimentColors": props(increaseFill=color(JADE), decreaseFill=color(COPPER), totalFill=color(GOLD)),
           "legend": props(show=lit(False))})
    # ---------------------------------------------------------------- 5 Cash forecast
    p5 = Page("forecast", "Cash Forecast")
    header(p5, "13-week cash forecast", "Cash today + the owner's plan (cash-plan.csv), amounts taken from the bank history  ·  USD")
    slicer(p5, "fc_buffer", 22, 74, 240, col("Cash Buffer", "Cash Buffer"), "MINIMUM CASH BUFFER", h=84, default=10000)
    kpi_row(p5, "fc_k", 74, 84, [("Cash Today", "Cash today", "◈", "Today Note", MUTED),
                                 ("Lowest Forecast Cash", "Lowest in 13 weeks", "▼", "Lowest Note", MUTED),
                                 ("Headroom", "Headroom over buffer", "⛨", "Headroom Note", MUTED, "Headroom Color"),
                                 ("Weeks Below Buffer", "Weeks below buffer", "!", "Weeks Below Note", MUTED, "Weeks Below Color")])
    chart(p5, "fc_line", 22, 170, 760, 284, "lineChart", "Forecast cash at the end of each week vs your buffer",
          {"Category": [col("Forecast Week", "Week")], "Y": [(mea("Forecast Balance"), "Forecast cash"), (mea("Buffer Line"), "Buffer")]},
          sort(col("Forecast Week", "Week")),
          {"labels": [{"properties": {"show": lit(True), "labelDisplayUnits": lit(1000), "labelPrecision": lit(1), "fontSize": lit(8)}},
                      {"properties": {"show": lit(False)}, "selector": {"metadata": "_Measures.Buffer Line"}}],
           "legend": legend(), "dataPoint": fills(("Forecast Balance", GOLD), ("Buffer Line", BAD)),
           "lineStyles": props(strokeWidth=lit(3), showMarker=lit(True), markerSize=lit(5))})
    chart(p5, "fc_bridge", 794, 170, 464, 284, "waterfallChart", "Next 13 weeks: where the cash goes",
          {"Category": [col("Forecast", "Cash Flow Line")], "Y": [(mea("Forecast Amount"), "Net cash")]}, sort(col("Forecast", "Cash Flow Line")),
          {"labels": labels(8, 1000, 1),
           "sentimentColors": props(increaseFill=color(JADE), decreaseFill=color(COPPER), totalFill=color(GOLD)),
           "legend": props(show=lit(False))})
    chart(p5, "fc_weeks", 22, 466, 560, 234, "tableEx", "Week by week",
          {"Values": [(col("Forecast Week", "Week"), "Week"), (mea("Forecast In"), "Money in"), (mea("Forecast Out"), "Money out"),
                      (mea("Forecast Amount"), "Net"), (mea("Forecast Balance"), "Cash at week end")]},
          sort(col("Forecast Week", "Week")), objects=grid_objects(9, pad=1) | NO_TOTALS)
    chart(p5, "fc_plan", 594, 466, 664, 234, "tableEx", "The plan behind it (edit cash-plan.csv, then Refresh)",
          {"Values": [(col("Cash Plan", "Item"), "Item"), (col("Cash Plan", "How It Is Estimated"), "Estimated from"),
                      (col("Cash Plan", "When"), "When"), (col("Cash Plan", "Per Time"), "Each time")]},
          sort(col("Cash Plan", "Item")), objects=grid_objects(9, pad=1) | NO_TOTALS)
    # ---------------------------------------------------------------- 6 Ledger
    p6 = Page("ledger", "Ledger")
    header(p6, "Ledger", SUB + "  ·  debit +, credit −")
    slicer(p6, "lg_acct", 22, 74, 330, ACCT, "ACCOUNT", "Account", h=84, default="1000 Business checking")
    slicer(p6, "lg_month", 364, 74, 220, MONTH, "MONTH", "Month", h=84)
    kpi_row(p6, "lg_k", 74, 84, [("Balance Start", "Balance at start", "◁"), ("Balance Movement", "Movement", "↕"),
                                 ("Closing Balance", "Balance at end", "▷")], x0=596, size=18)
    chart(p6, "lg_table", 22, 170, 1236, 530, "tableEx", "Ledger lines for the account (running balance after each line)",
          {"Values": [(col("Ledger", "Date"), "Date"), (col("Ledger", "Payee"), "Payee"), (col("Ledger", "Description"), "As on the export"),
                      (col("Ledger", "Source"), "Source"), (col("Ledger", "Line ID"), "Line"), (mea("Ledger Debit"), "Debit"),
                      (mea("Ledger Credit"), "Credit"), (mea("Running Balance"), "Balance")]},
          {"sort": [{"field": col("Ledger", "Date"), "direction": "Ascending"}, {"field": col("Ledger", "Line ID"), "direction": "Ascending"}],
           "isDefaultSort": False}, objects=grid_objects(9, pad=1) | NO_TOTALS)
    # ---------------------------------------------------------------- 7 Checks
    p7 = Page("checks", "Checks")
    header(p7, "Can the numbers be trusted?", "Recalculated on every refresh, independent of the slicers  ·  "
           "a second read of every export file")
    kpi_row(p7, "ck_k", 74, 88, [("Checks Passed", "Checks", "✓", None, None, GOOD), ("Checks Failed", "Failed", "✕"),
                                 ("Chk Review Lines", "Lines waiting for a rule", "?", None, None, AMBER)], x0=22, x1=846, size=22)
    chart(p7, "ck_table", 22, 174, 1236, 338, "tableEx", "Data checks",
          {"Values": [(col("Checks", "No"), "#"), (col("Checks", "Check"), "Check"), (col("Checks", "Why it matters"), "Why it matters"),
                      (mea("Check Detail"), "Result"), (mea("Check Status"), "Status")]},
          sort(col("Checks", "No")),
          objects=grid_objects(10, [cond_color("_Measures.Check Status", "Check Color")]) | NO_TOTALS)
    chart(p7, "ck_review", 22, 524, 1236, 176, "tableEx", "Lines waiting for a rule (add a keyword to rules.csv, then refresh)",
          {"Values": [(col("Ledger", "Date"), "Date"), (col("Ledger", "Source"), "Source"), (col("Ledger", "Description"), "Description"),
                      (col("Ledger", "Line ID"), "Line"), (mea("Review Amount"), "Amount")]},
          sort(col("Ledger", "Date")), objects=grid_objects(9))
    return [p1, p2, p3, p4, p5, p6, p7]


def build_report():
    if RP.exists():
        shutil.rmtree(RP)
    d = RP / "definition"
    (d / "pages").mkdir(parents=True)
    (RP / "definition.pbir").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0", "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}}, indent=2), encoding="utf-8")
    (d / "version.json").write_text(json.dumps({"$schema": f"{S}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"}, indent=2),
                                    encoding="utf-8")
    base = "CY25SU11"
    theme_src = ROOT / "scripts" / "theme" / f"{base}.json"
    (RP / "StaticResources" / "SharedResources" / "BaseThemes").mkdir(parents=True)
    shutil.copyfile(theme_src, RP / "StaticResources" / "SharedResources" / "BaseThemes" / f"{base}.json")
    (RP / "StaticResources" / "RegisteredResources").mkdir(parents=True)
    custom = {"name": "Copperleaf Dark", "dataColors": [JADE, GOLD, COPPER, SKY, "#9C8AC0", "#5F7A8A", "#D9B97A", "#6FB3A0"],
              "foreground": IVORY, "foregroundNeutralSecondary": MUTED, "foregroundNeutralTertiary": "#5F6B77",
              "background": CARD, "backgroundLight": CARD_ALT, "backgroundNeutral": LINE, "tableAccent": GOLD,
              "good": GOOD, "bad": BAD, "neutral": AMBER, "maximum": GOLD, "center": CARD_ALT, "minimum": CARD,
              "textClasses": {"callout": {"fontFace": SERIF, "color": IVORY}, "title": {"fontFace": SERIF, "color": IVORY},
                              "header": {"fontFace": SANS_SB, "color": GOLD}, "label": {"fontFace": SANS, "color": IVORY}},
              "visualStyles": {"page": {"*": {"background": [{"color": {"solid": {"color": PAGE_BG}}, "transparency": 0}]}},
                               "*": {"*": {"visualTooltip": [{"titleFontColor": GOLD, "valueFontColor": IVORY, "background": {"solid": {"color": HDR}}}]}}}}
    (RP / "StaticResources" / "RegisteredResources" / "Copperleaf.json").write_text(json.dumps(custom, indent=2), encoding="utf-8")
    ver = {"visual": "2.14.0", "report": "3.3.0", "page": "2.1.0"}
    report = {"$schema": f"{S}/report/3.3.0/schema.json",
              "themeCollection": {"baseTheme": {"name": base, "reportVersionAtImport": ver, "type": "SharedResources"},
                                  "customTheme": {"name": "Copperleaf.json", "reportVersionAtImport": ver, "type": "RegisteredResources"}},
              "resourcePackages": [{"name": "SharedResources", "type": "SharedResources",
                                    "items": [{"name": base, "path": f"BaseThemes/{base}.json", "type": "BaseTheme"}]},
                                   {"name": "RegisteredResources", "type": "RegisteredResources",
                                    "items": [{"name": "Copperleaf.json", "path": "Copperleaf.json", "type": "CustomTheme"}]}],
              "settings": {"useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarized",
                           "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True, "useEnhancedTooltips": True,
                           "useDefaultAggregateDisplayName": True},
              "objects": {"outspacePane": props(expanded=lit(False))}}  # the filter pane starts collapsed
    (d / "report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    pages = build_pages()
    (d / "pages" / "pages.json").write_text(json.dumps({"$schema": f"{S}/pagesMetadata/1.1.0/schema.json",
                                                         "pageOrder": [p.name for p in pages], "activePageName": pages[0].name},
                                                        indent=2), encoding="utf-8")
    for p in pages:
        pd = d / "pages" / p.name
        (pd / "visuals").mkdir(parents=True)
        (pd / "page.json").write_text(json.dumps({"$schema": f"{S}/page/2.1.0/schema.json", "name": p.name, "displayName": p.display,
                                                  "displayOption": "FitToPage", "height": 720, "width": 1280,
                                                  "objects": {"background": props(color=color(PAGE_BG), transparency=lit(0))}},
                                                 indent=2), encoding="utf-8")
        for v in p.visuals:
            (pd / "visuals" / v["name"]).mkdir()
            (pd / "visuals" / v["name"] / "visual.json").write_text(json.dumps(v, indent=2, ensure_ascii=False), encoding="utf-8")
    return pages


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-folder", default=r"C:\Copperleaf\data")
    a = ap.parse_args()
    folder = str(Path(a.data_folder).resolve()).rstrip("\\/") + "\\"
    build_model(folder)
    pages = build_report()
    (ROOT / f"{NAME}.pbip").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}], "settings": {"enableAutoRecovery": True}},
        indent=2), encoding="utf-8")
    print(f"{NAME}.pbip: {len(TABLES)} tables, {len(MEASURES)} measures, {len(RELATIONSHIPS)} relationships, "
          f"{len(pages)} pages, {sum(len(p.visuals) for p in pages)} visuals; data folder {folder}")


if __name__ == "__main__":
    main()
