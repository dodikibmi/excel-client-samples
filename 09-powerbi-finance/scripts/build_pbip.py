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
    Source = {csv_src("rules.csv", 4)},
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
    RuleList = Table.ToRecords(Table.SelectColumns(Rules, {"Clean Keyword", "Account No"})),
    WithClean = Table.AddColumn(Lines, "Clean Description", each Norm([Description]), type text),
    WithRule = Table.AddColumn(WithClean, "Rule", (r) => List.First(List.Select(RuleList, (k) => Text.Contains(r[Clean Description], k[Clean Keyword])), null)),
    WithAccount = Table.AddColumn(WithRule, "Offset Account No", each if [Rule] = null then 6990 else [Rule][Account No], Int64.Type),
    WithKeyword = Table.AddColumn(WithAccount, "Rule Keyword", each if [Rule] = null then "(no rule: to review)" else [Rule][Clean Keyword], type text)
in
    Table.RemoveColumns(WithKeyword, {"Rule"})
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
         #"Rule Keyword" = r[Rule Keyword], #"Account No" = r[Register No], #"Offset Account No" = r[Offset Account No], Amount = r[Amount]],
        [#"Line ID" = r[Line ID] & "-A", Side = "Account", Date = r[Post Date], #"Txn Date" = r[Txn Date], Source = r[Source],
         #"Register No" = r[Register No], Description = r[Description], #"Clean Description" = r[Clean Description],
         #"Rule Keyword" = r[Rule Keyword], #"Account No" = r[Offset Account No], #"Offset Account No" = r[Register No], Amount = -r[Amount]]
    })),
    Journal = List.Transform(List.Positions(Table.ToRecords(Adjustments)), (i) =>
        let a = Table.ToRecords(Adjustments){i} in
        [#"Line ID" = a[Entry] & "-" & Text.From(i + 1), Side = "Journal", Date = a[Date], #"Txn Date" = a[Date], Source = "Journal",
         #"Register No" = null, Description = a[Description], #"Clean Description" = Norm(a[Description]),
         #"Rule Keyword" = null, #"Account No" = a[Account No], #"Offset Account No" = null, Amount = a[Debit] - a[Credit]]),
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
        {"Register No", Int64.Type}, {"Description", type text}, {"Clean Description", type text}, {"Rule Keyword", type text},
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
    Out = Table.RenameColumns(Table.RemoveColumns(Key, {"Is Cash"}), {{"Is Cash Account", "Is Cash"}})
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
    Weekday = Table.AddColumn(Day, "Weekday", each Date.ToText([Date], [Format = "ddd", Culture = "en-US"]), type text)
in
    Weekday
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
    ("Net Profit Color", "1 Profit and loss", None, 'IF ( [Net Profit] < 0, "#C0392B", "#00968A" )', True),
    ("Net Margin %", "1 Profit and loss", "0.0%;-0.0%;0.0%", "DIVIDE ( [Net Profit], [Total Income] )", False),
    ("Actual", "2 Budget", "#,0;(#,0);0",
     '-- P&L sign: income +, costs -. The total of the P&L is the net profit\nCALCULATE ( -SUM ( Ledger[Amount] ), KEEPFILTERS ( Category[Statement] = "P&L" ) )', False),
    ("Budget", "2 Budget", "#,0;(#,0);0", "SUM ( Budget[Budget Amount] )", False),
    ("Variance", "2 Budget", "+#,0;-#,0;0", "-- + = better than budget (more income or less cost)\n[Actual] - [Budget]", False),
    ("Variance %", "2 Budget", "+0.0%;-0.0%;0.0%", "DIVIDE ( [Variance], ABS ( [Budget] ) )", False),
    ("Variance Color", "2 Budget", None,
     'SWITCH ( TRUE (), [Variance] < -0.5, "#C0392B", [Variance] > 0.5, "#1E8449", "#5F6B7A" )', True),
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
     'SWITCH ( [Check Status], "PASS", "#1E8449", "FAIL", "#C0392B", "REVIEW", "#B9770E" )', True),
    ("Checks Passed", "5 Checks", None,
     'VAR t = ADDCOLUMNS ( ALL ( Checks[No] ), "@s", [Check Status] )\n'
     'RETURN FORMAT ( COUNTROWS ( FILTER ( t, [@s] = "PASS" ) ), "0" ) & " of " & FORMAT ( COUNTROWS ( FILTER ( t, [@s] <> "REVIEW" ) ), "0" ) & " PASS"', False),
    ("Checks Failed", "5 Checks", "0",
     'VAR t = ADDCOLUMNS ( ALL ( Checks[No] ), "@s", [Check Status] )\nRETURN COUNTROWS ( FILTER ( t, [@s] = "FAIL" ) ) + 0', False),
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
        ("Rule Keyword", "string", None, "none", []), ("Account No", "int64", "0", "none", []),
        ("Offset Account No", "int64", "0", "none", []), ("Amount", "decimal", "#,0.00", "sum", []),
        ("Debit", "decimal", "#,0.00", "sum", []), ("Credit", "decimal", "#,0.00", "sum", []),
        ("Cash Flow Line", "string", None, "none", ["sortByColumn: 'Cash Flow Order'"]),
        ("Cash Flow Order", "int64", "0", "none", ["isHidden"])]),
    "Account": dict(src=ACCOUNT, desc="Chart of accounts with opening balances at 31 Jul 2026 (debit +).", cols=[
        ("Account No", "int64", "0", "none", ["isKey"]), ("Account", "string", None, "none", []),
        ("Type", "string", None, "none", []), ("Opening Balance", "decimal", "#,0.00", "sum", []),
        ("Statement Name", "string", None, "none", ["isHidden"]),
        ("Account Label", "string", None, "none", ["sortByColumn: 'Account No'"]),
        ("Is Cash", "boolean", None, "none", []), ("Category Key", "string", None, "none", [])]),
    "Category": dict(src=CATEGORY, desc="Groups accounts into the lines of the P&L and balance sheet.", cols=[
        ("Category Key", "string", None, "none", ["isKey"]), ("Category", "string", None, "none", ["sortByColumn: Sort"]),
        ("Group", "string", None, "none", []), ("Statement", "string", None, "none", []), ("Sort", "int64", "0", "none", ["isHidden"])]),
    "Date": dict(src=DATE, desc="Calendar: one row per day of the months in the exports.", dataCategory="Time", cols=[
        ("Date", "dateTime", "d mmm yyyy", "none", ["isKey"]), ("Year", "int64", "0", "none", []),
        ("Month No", "int64", "0", "none", ["isHidden"]), ("Month", "string", None, "none", ["sortByColumn: 'Month No'"]),
        ("Month Start", "dateTime", "d mmm yyyy", "none", []), ("Month End", "dateTime", "d mmm yyyy", "none", []),
        ("Day", "int64", "0", "none", []), ("Weekday", "string", None, "none", [])]),
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
}
RELATIONSHIPS = [  # many side -> one side
    ("Ledger", "Date", "Date", "Date"), ("Ledger", "Account No", "Account", "Account No"),
    ("Account", "Category Key", "Category", "Category Key"),
    ("Budget", "Month", "Date", "Date"), ("Budget", "Account No", "Account", "Account No"),
    ("Statements", "Statement Date", "Date", "Date"), ("Statements", "Account No", "Account", "Account No"),
]
EXPRESSIONS = [  # shared queries, not loaded: name, M, result type
    ("Norm", NORM, "Function"), ("Money", MONEY, "Function"),
    ("Checking_Clean", CHECKING, "Table"), ("Savings_Clean", SAVINGS, "Table"), ("Card_Clean", CARD, "Table"),
    ("Rules", RULES, "Table"), ("AccountsRaw", ACCOUNTS_RAW, "Table"), ("Adjustments", ADJ, "Table"),
    ("BankLines", BANKLINES, "Table"),
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
             "Source Files": (780, 700)}
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
NAVY, TEAL, INK, MUTED, PAGE_BG = "#18324D", "#00968A", "#1B2A3A", "#5F6B7A", "#F3F5F8"


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


def container(title=None, bg="#FFFFFF", border=True):
    o = {"background": props(show=lit(True), color=color(bg), transparency=lit(0)),
         "border": props(show=lit(border), color=color("#DCE1E7"), radius=lit(8)),
         "dropShadow": props(show=lit(False))}
    if title:
        o["title"] = props(show=lit(True), text=lit(title), fontColor=color(INK), fontSize=lit(12), bold=lit(True))
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


def header(page, title, sub):
    paras = [{"textRuns": [{"value": title, "textStyle": {"fontFamily": "Segoe UI Semibold", "fontSize": "20pt", "color": "#FFFFFF"}}]},
             {"textRuns": [{"value": sub, "textStyle": {"fontFamily": "Segoe UI", "fontSize": "11pt", "color": "#B8C8DA"}}]}]
    page.add(f"{page.name}_hdr", 0, 0, 1280, 72, {
        "visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": paras}}]},
        "visualContainerObjects": container(bg=NAVY, border=False) | {"padding": props(top=lit(8), left=lit(22))}})


def slicer(page, vid, x, y, w, field, title, sync=None, h=76):
    v = {"visualType": "slicer",
         "query": {"queryState": {"Values": {"projections": [proj(field)]}}},
         "objects": {"data": props(mode=lit("Dropdown")),
                     "header": props(show=lit(True), text=lit(title), fontColor=color(MUTED), textSize=lit(10)),
                     "items": props(textSize=lit(11))},
         "visualContainerObjects": container()}
    if sync:
        v["syncGroup"] = {"groupName": sync, "fieldChanges": True, "filterChanges": True}
    page.add(vid, x, y, w, h, v)


def card(page, vid, x, y, w, h, measure, label, value_color=INK, size=24):
    # value_color: a hex colour, or the name of a measure that returns one (red when negative)
    fill = color(value_color) if value_color.startswith("#") else {"solid": {"color": {"expr": mea(value_color)}}}
    page.add(vid, x, y, w, h, {
        "visualType": "card",
        "query": {"queryState": {"Values": {"projections": [proj(mea(measure), label)]}}},
        "objects": {"labels": props(fontSize=lit(size), color=fill, labelDisplayUnits=lit(1)),
                    "categoryLabels": props(show=lit(True), fontSize=lit(10), color=color(MUTED))},
        "visualContainerObjects": container()})


def sort(field, direction="Ascending"):
    return {"sort": [{"field": field, "direction": direction}], "isDefaultSort": False}


def chart(page, vid, x, y, w, h, vtype, title, roles, sort_def=None, objects=None):
    v = {"visualType": vtype, "query": {"queryState": {r: {"projections": [proj(*f) if isinstance(f, tuple) else proj(f) for f in fs]}
                                                       for r, fs in roles.items()}},
         "visualContainerObjects": container(title)}
    if sort_def:
        v["query"]["sortDefinition"] = sort_def
    if objects:
        v["objects"] = objects
    if vtype in AXIS_CHARTS:  # the title says what the axes are
        v.setdefault("objects", {})
        v["objects"]["categoryAxis"] = [{"properties": {"showAxisTitle": lit(False)} | v["objects"].get("categoryAxis", [{}])[0].get("properties", {})}]
        v["objects"]["valueAxis"] = props(showAxisTitle=lit(False), secShowAxisTitle=lit(False))
    page.add(vid, x, y, w, h, v)
    return v


def cond_color(measure_queryref, color_measure):
    return {"properties": {"fontColor": {"solid": {"color": {"expr": mea(color_measure)}}}},
            "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": measure_queryref}}


def grid_objects(font=11, extra_values=None, pad=4):
    o = {"values": [{"properties": {"fontSize": lit(font)}}] + (extra_values or []),
         "columnHeaders": props(fontSize=lit(font), bold=lit(True), fontColor=color(INK), backColor=color("#E8EDF2")),
         "grid": props(rowPadding=lit(pad))}
    return o


AXIS_CHARTS = {"lineClusteredColumnComboChart", "clusteredBarChart", "clusteredColumnChart", "lineChart"}
NO_TOTALS = {"total": props(totals=lit(False))}
MONTH, ACCT, CATG = col("Date", "Month"), col("Account", "Account Label"), col("Category", "Category")


def build_pages():
    sub = "Copperleaf Bakery & Cafe (fictional)  ·  Aug–Oct 2026  ·  built from the bank and card exports"
    # ---------------------------------------------------------------- 1 Overview
    p1 = Page("overview", "Overview")
    header(p1, "Finance overview", sub)
    slicer(p1, "ov_month", 22, 84, 300, MONTH, "Month", "Month")
    slicer(p1, "ov_acct", 334, 84, 380, ACCT, "Account", "Account")
    kpis = [("Total Income", "Income"), ("Total Expense", "Expenses"), ("Net Profit", "Net profit"),
            ("Net Margin %", "Net margin"), ("Income MoM %", "Income vs prior month"), ("Cash Balance", "Cash, checking + savings")]
    for i, (m, lbl) in enumerate(kpis):
        card(p1, f"ov_k{i}", 22 + i * 208, 172, 196, 100, m, lbl, "Net Profit Color" if m == "Net Profit" else INK)
    chart(p1, "ov_trend", 22, 284, 764, 416, "lineClusteredColumnComboChart", "Income, expenses and net profit by month",
          {"Category": [MONTH], "Y": [(mea("Total Income"), "Income"), (mea("Total Expense"), "Expenses")], "Y2": [(mea("Net Profit"), "Net profit")]}, sort(MONTH),
          {"labels": props(show=lit(True), labelDisplayUnits=lit(1000), labelPrecision=lit(1), fontSize=lit(9)),
           "legend": props(show=lit(True), position=lit("Top")),
           "dataPoint": [{"properties": {"fill": color("#1F4E79")}, "selector": {"metadata": "_Measures.Total Income"}},
                         {"properties": {"fill": color("#E07B39")}, "selector": {"metadata": "_Measures.Total Expense"}},
                         {"properties": {"fill": color(TEAL)}, "selector": {"metadata": "_Measures.Net Profit"}}]})
    chart(p1, "ov_exp", 798, 284, 460, 416, "clusteredBarChart", "Where the money goes (expenses by account)",
          {"Category": [ACCT], "Y": [mea("Total Expense")]}, sort(mea("Total Expense"), "Descending"),
          {"labels": props(show=lit(True), labelDisplayUnits=lit(1), fontSize=lit(9)),
           "dataPoint": props(fill=color("#E07B39"))})
    # ---------------------------------------------------------------- 2 P&L vs budget
    p2 = Page("pnl", "P&L vs Budget")
    header(p2, "Profit and loss vs budget", sub + "  ·  + = better than budget")
    slicer(p2, "pl_month", 22, 84, 300, MONTH, "Month", "Month")
    card(p2, "pl_k1", 334, 84, 200, 76, "Actual", "Net profit", "Net Profit Color", size=20)
    card(p2, "pl_k2", 546, 84, 200, 76, "Budget", "Budgeted net profit", size=20)
    card(p2, "pl_k3", 758, 84, 200, 76, "Variance", "Variance vs budget", size=20)
    mx = chart(p2, "pl_matrix", 22, 172, 652, 528, "pivotTable", "P&L by account: actual, budget, variance",
               {"Rows": [CATG, ACCT], "Values": [mea("Actual"), mea("Budget"), mea("Variance"), mea("Variance %")]},
               objects=grid_objects(10, [cond_color("_Measures.Variance", "Variance Color"), cond_color("_Measures.Variance %", "Variance Color")], pad=1)
               | {"rowHeaders": props(fontSize=lit(10))})
    mx["expansionStates"] = [{"roles": ["Rows"], "levels": [{"queryRefs": ["Category.Category"], "identityKeys": [CATG], "isPinned": True},
                                                            {"queryRefs": ["Account.Account Label"], "isPinned": True}],
                              "root": {"isToggled": True}}]
    chart(p2, "pl_trend", 686, 172, 572, 256, "clusteredColumnChart", "Net profit by month: actual vs budget",
          {"Category": [MONTH], "Y": [(mea("Actual"), "Actual"), (mea("Budget"), "Budget")]}, sort(MONTH),
          {"labels": props(show=lit(True), labelDisplayUnits=lit(1), fontSize=lit(9)), "legend": props(show=lit(True), position=lit("Top")),
           "dataPoint": [{"properties": {"fill": color(TEAL)}, "selector": {"metadata": "_Measures.Actual"}},
                         {"properties": {"fill": color("#B8C2CC")}, "selector": {"metadata": "_Measures.Budget"}}]})
    chart(p2, "pl_var", 686, 440, 572, 260, "clusteredBarChart", "Variance by account (worst first)",
          {"Category": [ACCT], "Y": [mea("Variance")]}, sort(mea("Variance"), "Ascending"),
          {"labels": props(show=lit(True), labelDisplayUnits=lit(1), fontSize=lit(9)),
           "dataPoint": [{"properties": {"fill": {"solid": {"color": {"expr": mea("Variance Color")}}}},
                          "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}}]})
    # ---------------------------------------------------------------- 3 Cash
    p3 = Page("cash", "Cash & Statements")
    header(p3, "Cash and bank statements", sub)
    slicer(p3, "ca_month", 22, 84, 300, MONTH, "Month", "Month")
    card(p3, "ca_k1", 334, 84, 240, 76, "Cash Balance", "Cash at end (checking + savings)", size=20)
    card(p3, "ca_k2", 586, 84, 240, 76, "In Transit", "Transfers still in transit", size=20)
    chart(p3, "ca_line", 22, 172, 690, 272, "lineChart", "Cash balance day by day (checking + savings)",
          {"Category": [col("Date", "Date")], "Y": [mea("Cash Balance")]}, sort(col("Date", "Date")),
          {"dataPoint": props(fill=color("#1F4E79")), "categoryAxis": props(axisType=lit("Scalar"))})
    chart(p3, "ca_flow", 22, 456, 690, 244, "pivotTable", "Cash flow by month (checking + savings)",
          {"Rows": [col("Ledger", "Cash Flow Line")], "Columns": [MONTH], "Values": [(mea("Cash Movement"), "Net cash")]},
          objects=grid_objects(10) | {"rowHeaders": props(fontSize=lit(10))})
    chart(p3, "ca_rec", 724, 172, 534, 346, "tableEx", "Books vs statements (card owed shown negative)",
          {"Values": [(ACCT, "Account"), (MONTH, "Month"), (mea("Books Balance"), "Books"), (mea("Statement Balance"), "Statement"),
                      (mea("Difference vs Statement"), "Difference")]},
          objects=grid_objects(10) | NO_TOTALS)
    p3.add("ca_note", 724, 530, 534, 170, {
        "visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": [
            {"textRuns": [{"value": "How to read it", "textStyle": {"fontWeight": "bold", "fontSize": "11pt", "color": INK}}]},
            {"textRuns": [{"value": "Books = opening balance + every posting up to the statement date. A difference of 0.00 for every "
                                    "account and month proves nothing was lost, doubled or booked in the wrong month. "
                                    "Money moved between own accounts (checking to savings, card payments) waits on "
                                    "1090 Transfers in transit until the other side arrives.",
                           "textStyle": {"fontSize": "10pt", "color": MUTED}}]}]}}]},
        "visualContainerObjects": container()})
    # ---------------------------------------------------------------- 4 Checks
    p4 = Page("checks", "Checks")
    header(p4, "Checks: can these numbers be trusted?", "Recalculated on every refresh, independent of the slicers  ·  "
           "the same checks as the Excel version, plus a second read of every export file")
    card(p4, "ck_k1", 22, 84, 300, 80, "Checks Passed", "Checks")
    card(p4, "ck_k2", 334, 84, 200, 80, "Checks Failed", "Failed")
    card(p4, "ck_k3", 546, 84, 240, 80, "Chk Review Lines", "Lines waiting for a rule")
    chart(p4, "ck_table", 22, 176, 1236, 336, "tableEx", "Data checks",
          {"Values": [(col("Checks", "No"), "#"), (col("Checks", "Check"), "Check"), (col("Checks", "Why it matters"), "Why it matters"),
                      (mea("Check Detail"), "Result"), (mea("Check Status"), "Status")]},
          sort(col("Checks", "No")),
          objects=grid_objects(11, [cond_color("_Measures.Check Status", "Check Color")]) | NO_TOTALS)
    chart(p4, "ck_review", 22, 524, 1236, 176, "tableEx", "Lines waiting for a rule (add a keyword to rules.csv, then refresh)",
          {"Values": [(col("Ledger", "Date"), "Date"), (col("Ledger", "Source"), "Source"), (col("Ledger", "Description"), "Description"),
                      (col("Ledger", "Line ID"), "Line"), (mea("Review Amount"), "Amount")]},
          sort(col("Ledger", "Date")), objects=grid_objects(10))
    return [p1, p2, p3, p4]


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
    custom = {"name": "Copperleaf", "dataColors": ["#1F4E79", TEAL, "#E07B39", "#7A5195", MUTED, "#C0392B", "#1E8449", "#B9770E"],
              "foreground": INK, "background": "#FFFFFF", "tableAccent": TEAL,
              "visualStyles": {"page": {"*": {"background": [{"color": {"solid": {"color": PAGE_BG}}, "transparency": 0}]}}}}
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
