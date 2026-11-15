"""Builds Bookkeeping-Trial-Balance.xlsx with real Excel (COM).

Power Query reads three bank/card exports, cleans them, gives every line an account and a payee from the Rules tab,
pairs the two halves of every transfer and card payment, and writes a general ledger (bank lines + adjusting entries).
Formulas turn the ledger into a trial balance per month, a P&L vs budget, a balance sheet, a cash flow statement
(indirect method), a 13-week cash forecast, a bank reconciliation and a month-end close checklist. The Summary is a
dark-gold dashboard: month and comparison slicers, KPI cards, a profit waterfall, costs with sparklines, payees and
checks, with a navigation menu on every report sheet.

Run with a Python that has pywin32:  python scripts/build_workbook.py
The workbook has no macros; the client only needs Data -> Refresh All.
"""
import sys, time
from pathlib import Path
import win32com.client as win32

sys.path.insert(0, str(Path(__file__).parent))
from coa import ACCOUNTS, RULES, BUDGET, ADJUSTMENTS, MONTHS, CLEARING, UNCATEGORIZED, CASH_PLAN, CASH_BUFFER

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Bookkeeping-Trial-Balance.xlsx"

# ---------------------------------------------------------------- Power Query (M)
Q = {}
Q["Setting"] = """(name as text) as text =>
    // Reads one value from a named cell on the Settings sheet
    Text.From(Excel.CurrentWorkbook(){[Name = name]}[Content]{0}[Column1])"""

Q["Money"] = """(t as nullable text) as number =>
    // "$1,250.00", "2,000.00", "-89.99", "" -> numbers (blank = 0)
    let s = if t = null then "" else Text.Remove(Text.Trim(t), {"$", ",", " "})
    in  if s = "" then 0 else Number.FromText(s, "en-US")"""

Q["Norm"] = """(t as nullable text) as text =>
    // "TapPay Inc. Deposit" -> "TAPPAY INC DEPOSIT"; "CITY PWR&LIGHT" -> "CITY PWR LIGHT": upper case, punctuation to spaces
    let u = Text.Upper(if t = null then "" else t),
        c = Text.Combine(List.Transform(Text.ToList(u), each if List.Contains({"*", ".", "-", "#", "/", ",", "&", "'"}, _) then " " else _)),
        w = List.Select(Text.Split(c, " "), each _ <> "")
    in  Text.Combine(w, " ")"""

Q["Checking_Raw"] = """let
    Path    = Setting("DataFolder") & Setting("CheckingFile"),
    Source  = Csv.Document(File.Contents(Path), [Delimiter = ",", Columns = 5, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    // the bank puts its name, the period and a blank line above the real header
    Body    = Table.Skip(Source, 3),
    Headers = Table.PromoteHeaders(Body, [PromoteAllScalars = true])
in
    Headers"""

Q["Savings_Raw"] = """let
    Path    = Setting("DataFolder") & Setting("SavingsFile"),
    Source  = Csv.Document(File.Contents(Path), [Delimiter = ",", Columns = 3, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Headers"""

Q["Card_Raw"] = """let
    Path    = Setting("DataFolder") & Setting("CardFile"),
    Source  = Csv.Document(File.Contents(Path), [Delimiter = ",", Columns = 6, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Headers"""

Q["Statements_Raw"] = """let
    Path    = Setting("DataFolder") & Setting("StatementsFile"),
    Source  = Csv.Document(File.Contents(Path), [Delimiter = ",", Columns = 3, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Headers"""

REG = """    Row     = Table.AddIndexColumn(Typed, "Row", 1, 1, Int64.Type),
    Line    = Table.AddColumn(Row, "Line", each "{code}-" & Text.PadStart(Text.From([Row]), 3, "0"), type text),
    Reg     = Table.AddColumn(Line, "Register", each "{name}", type text),
    No      = Table.AddColumn(Reg, "Register no", each {no}, Int64.Type),
    Result  = Table.SelectColumns(No, {{"Line", "Register", "Register no", "Posted", "Bought", "Description", "Amount"}})
in
    Result"""

Q["Checking_Clean"] = """let
    Source  = Checking_Raw,
    // Debit and Credit arrive as text like "$1,250.00"; one signed amount: money in = +, money out = -
    Amt     = Table.AddColumn(Source, "Amount", each Money([Credit]) - Money([Debit]), type number),
    Dt      = Table.AddColumn(Amt, "Posted", each Date.FromText(Text.Trim([Date]), "en-US"), type date),
    Bt      = Table.AddColumn(Dt, "Bought", each [Posted], type date),
    Typed   = Table.TransformColumns(Bt, {{"Description", Text.Trim, type text}}),
""" + REG.format(code="CHK", name="Business checking ...1180", no=1000)

Q["Savings_Clean"] = """let
    Source  = Table.RenameColumns(Savings_Raw, {{"Posted", "Posted text"}, {"Details", "Description"}, {"Amount", "Amount text"}}),
    // signed amounts with thousand separators ("2,000.00"), ISO dates
    Amt     = Table.AddColumn(Source, "Amount", each Money([Amount text]), type number),
    Dt      = Table.AddColumn(Amt, "Posted", each Date.FromText(Text.Trim([Posted text])), type date),
    Bt      = Table.AddColumn(Dt, "Bought", each [Posted], type date),
    Typed   = Table.TransformColumns(Bt, {{"Description", Text.Trim, type text}}),
""" + REG.format(code="SAV", name="Business savings ...4821", no=1010)

Q["Card_Clean"] = """let
    Source  = Table.RenameColumns(Card_Raw, {{"Amount", "Amount text"}}),
    // card exports are already signed from the card's side: charges -, payments +.
    // The statement month follows the POST date; the purchase date is kept for reference.
    Amt     = Table.AddColumn(Source, "Amount", each Money([Amount text]), type number),
    Dt      = Table.AddColumn(Amt, "Posted", each Date.FromText(Text.Trim([Post Date]), "en-US"), type date),
    Bt      = Table.AddColumn(Dt, "Bought", each Date.FromText(Text.Trim([Transaction Date]), "en-US"), type date),
    Typed   = Table.TransformColumns(Bt, {{"Description", Text.Trim, type text}}),
""" + REG.format(code="VISA", name="Visa business card ...7702", no=2000)

Q["Accounts_T"] = """let
    Source = Excel.CurrentWorkbook(){[Name = "Accounts"]}[Content],
    Kept   = Table.SelectRows(Source, each [Account] <> null),
    Typed  = Table.TransformColumnTypes(Table.SelectColumns(Kept, {"Account", "Name", "Type"}),
                 {{"Account", Int64.Type}, {"Name", type text}, {"Type", type text}})
in
    Typed"""

Q["Rule_List"] = """let
    // the Rules tab, top to bottom: the first keyword found in the description decides the account
    Source = Excel.CurrentWorkbook(){[Name = "Rules"]}[Content],
    Kept   = Table.SelectRows(Source, each [Keyword] <> null and Text.Trim(Text.From([Keyword])) <> "" and [Account] <> null),
    Recs   = List.Buffer(List.Transform(Table.ToRecords(Kept), each [Keyword = Norm(Text.From([Keyword])), Account = Number.From([Account]),
                 Payee = if [Payee] = null then "" else Text.From([Payee])]))
in
    Recs"""

Q["Transactions"] = f"""let
    All     = Table.Combine({{Checking_Clean, Savings_Clean, Card_Clean}}),
    Rules   = Rule_List,
    Clean   = Table.AddColumn(All, "Clean description", each Norm([Description]), type text),
    Hit     = Table.AddColumn(Clean, "Hit", each List.First(List.Select(Rules, (r) => Text.Contains([Clean description], r[Keyword])), null)),
    Acc     = Table.AddColumn(Hit, "Account", each if [Hit] = null then {UNCATEGORIZED} else [Hit][Account], Int64.Type),
    Rule    = Table.AddColumn(Acc, "Matched rule", each if [Hit] = null then "(no rule)" else [Hit][Keyword], type text),
    // who gets (or pays) the money, from the same rule; lines no rule knows stay "Unknown" until explained
    Payee   = Table.AddColumn(Rule, "Payee", each if [Hit] = null then "Unknown (on Review)" else [Hit][Payee], type text),
    Month   = Table.AddColumn(Payee, "MonthNo", each Date.Year([Posted]) * 100 + Date.Month([Posted]), Int64.Type),
    Named   = Table.NestedJoin(Month, {{"Account"}}, Accounts_T, {{"Account"}}, "A", JoinKind.LeftOuter),
    Exp     = Table.ExpandTableColumn(Named, "A", {{"Name", "Type"}}, {{"Account name", "Type"}}),
    Sorted  = Table.Sort(Exp, {{{{"Posted", Order.Ascending}}, {{"Line", Order.Ascending}}}}),
    Result  = Table.SelectColumns(Sorted, {{"Line", "Register", "Register no", "Posted", "MonthNo", "Bought", "Description",
                  "Clean description", "Amount", "Matched rule", "Account", "Account name", "Payee", "Type"}})
in
    Result"""

Q["Ledger"] = """let
    T       = Transactions,
    // every bank or card line is two ledger lines: the register itself and the account the rule chose
    RegLeg  = Table.RenameColumns(Table.SelectColumns(T, {"Posted", "MonthNo", "Line", "Description", "Register no", "Amount"}),
                  {{"Register no", "Account"}, {"Amount", "Net"}}),
    OffLeg0 = Table.SelectColumns(T, {"Posted", "MonthNo", "Line", "Description", "Account", "Amount"}),
    OffLeg  = Table.RemoveColumns(Table.AddColumn(OffLeg0, "Net", each -[Amount], type number), {"Amount"}),
    Bank    = Table.AddColumn(Table.Combine({RegLeg, OffLeg}), "Source", each "Bank line", type text),
    // adjusting entries typed on the Adjustments tab (depreciation, accruals ...)
    AdjSrc  = Table.SelectRows(Excel.CurrentWorkbook(){[Name = "Adjustments"]}[Content], each [Account] <> null),
    Adj     = Table.FromRecords(List.Transform(Table.ToRecords(AdjSrc), each [
                  Posted = Date.From([Date]), MonthNo = Date.Year(Date.From([Date])) * 100 + Date.Month(Date.From([Date])),
                  Line = Text.From([Entry]), Description = Text.From([Description]), Account = Number.From([Account]),
                  Net = (if [Debit] = null then 0 else [Debit]) - (if [Credit] = null then 0 else [Credit]), Source = "Adjustment"])),
    All     = Table.Combine({Bank, Adj}),
    Rounded = Table.TransformColumns(All, {{"Net", each Number.Round(_, 2), type number}, {"Account", each Int64.From(_), Int64.Type},
                  {"MonthNo", each Int64.From(_), Int64.Type}, {"Posted", Date.From, type date}}),
    Named   = Table.NestedJoin(Rounded, {"Account"}, Accounts_T, {"Account"}, "A", JoinKind.LeftOuter),
    Exp     = Table.ExpandTableColumn(Named, "A", {"Name", "Type"}, {"Account name", "Type"}),
    Dr      = Table.AddColumn(Exp, "Debit", each if [Net] > 0 then [Net] else null, type number),
    Cr      = Table.AddColumn(Dr, "Credit", each if [Net] < 0 then -[Net] else null, type number),
    Sorted  = Table.Sort(Cr, {{"Posted", Order.Ascending}, {"Line", Order.Ascending}, {"Net", Order.Descending}}),
    Result  = Table.SelectColumns(Sorted, {"Posted", "MonthNo", "Source", "Line", "Description", "Account", "Account name", "Type",
                  "Debit", "Credit", "Net"})
in
    Result"""

Q["Transfers"] = f"""let
    // both halves of a transfer (checking -> savings, checking -> card) were given the clearing account {CLEARING};
    // pair them: same amount, opposite direction, posted within 5 days
    Legs    = Table.SelectRows(Transactions, each [Account] = {CLEARING}),
    Key     = Table.AddColumn(Legs, "Key", each Number.Round(Number.Abs([Amount]), 2), type number),
    Rank    = (t as table) as table => Table.Combine(Table.Group(t, {{"Key"}},
                  {{{{"G", each Table.AddIndexColumn(Table.Sort(_, {{{{"Posted", Order.Ascending}}, {{"Line", Order.Ascending}}}}), "Rank", 1, 1)}}}})[G]),
    Out     = Table.RenameColumns(Table.SelectColumns(Rank(Table.SelectRows(Key, each [Amount] < 0)), {{"Key", "Rank", "Register", "Line", "Posted"}}),
                  {{{{"Register", "From"}}, {{"Line", "From line"}}, {{"Posted", "From date"}}}}),
    In      = Table.RenameColumns(Table.SelectColumns(Rank(Table.SelectRows(Key, each [Amount] > 0)), {{"Key", "Rank", "Register", "Line", "Posted"}}),
                  {{{{"Key", "In key"}}, {{"Rank", "In rank"}}, {{"Register", "To"}}, {{"Line", "To line"}}, {{"Posted", "To date"}}}}),
    J       = Table.Join(Out, {{"Key", "Rank"}}, In, {{"In key", "In rank"}}, JoinKind.FullOuter),
    Amt     = Table.AddColumn(J, "Amount", each if [Key] <> null then [Key] else [In key], type number),
    Days    = Table.AddColumn(Amt, "Days apart", each if [From date] <> null and [To date] <> null then Duration.Days([To date] - [From date]) else null, Int64.Type),
    Status  = Table.AddColumn(Days, "Status", each
                  if [From line] = null or [To line] = null then "Unpaired"
                  else if Number.Abs([Days apart]) > 5 then "Unpaired (more than 5 days apart)" else "Paired", type text),
    Sorted  = Table.Sort(Status, {{{{"From date", Order.Ascending}}, {{"To date", Order.Ascending}}}}),
    Result  = Table.SelectColumns(Sorted, {{"Amount", "From", "From line", "From date", "To", "To line", "To date", "Days apart", "Status"}})
in
    Result"""

Q["Review"] = f"""let
    T       = Transactions,
    NoRule  = Table.AddColumn(Table.AddColumn(Table.SelectRows(T, each [Account] = {UNCATEGORIZED}),
                  "Issue", each "No rule matched", type text),
                  "What to do", each "Ask the owner what it was, add a keyword on the Rules tab, Refresh All", type text),
    Open    = Table.SelectRows(Transfers, each [Status] <> "Paired"),
    OpenL   = List.Buffer(List.Select(List.Combine({{Open[From line], Open[To line]}}), each _ <> null)),
    Legs    = Table.SelectRows(T, each List.Contains(OpenL, [Line])),
    Leg     = Table.AddColumn(Table.AddColumn(Legs, "Issue", each "Transfer without its other half", type text),
                  "What to do", each if [Amount] < 0
                      then "Left on " & Date.ToText([Posted], "dd MMM", "en-US") & ", other side not in the files yet: in transit if it posts early next month, else find it"
                      else "Money arrived in " & [Register] & " with no matching payment out: check where it came from", type text),
    All     = Table.Combine({{NoRule, Leg}}),
    Result  = Table.SelectColumns(All, {{"Issue", "Line", "Register", "Posted", "Description", "Amount", "What to do"}})
in
    Result"""

Q["Statements"] = """let
    Source  = Statements_Raw,
    No      = Table.AddColumn(Source, "Account no", each let u = Text.Upper([Account]) in
                  if Text.Contains(u, "CHECKING") then 1000 else if Text.Contains(u, "SAVINGS") then 1010
                  else if Text.Contains(u, "VISA") then 2000 else null, Int64.Type),
    Dt      = Table.AddColumn(No, "Date", each Date.FromText(Text.Trim([Statement date]), "en-US"), type date),
    Mo      = Table.AddColumn(Dt, "MonthNo", each Date.Year([Date]) * 100 + Date.Month([Date]), Int64.Type),
    Bal     = Table.AddColumn(Mo, "Balance", each Money([Ending balance]), type number),
    Result  = Table.RenameColumns(Table.SelectColumns(Bal, {"Account", "Account no", "Date", "MonthNo", "Balance"}),
                  {{"Account", "Account (as printed)"}, {"Date", "Statement date"}, {"Balance", "Ending balance"}})
in
    Result"""

Q["Cleaning_Log"] = f"""let
    T        = Transactions,
    Chk      = Table.RowCount(Checking_Raw),
    Sav      = Table.RowCount(Savings_Raw),
    Card     = Table.RowCount(Card_Raw),
    Later    = Table.RowCount(Table.SelectRows(Card_Clean, each Date.Month([Bought]) <> Date.Month([Posted]))),
    Typed    = Table.RowCount(Table.SelectRows(T, each Text.Trim([Description]) <> [Clean description])),
    Ruled    = Table.RowCount(Table.SelectRows(T, each [Account] <> {UNCATEGORIZED})),
    Legs     = Table.RowCount(Table.SelectRows(T, each [Account] = {CLEARING})),
    Paired   = Table.RowCount(Table.SelectRows(Transfers, each [Status] = "Paired")),
    Adj      = Table.RowCount(Table.SelectRows(Ledger, each [Source] = "Adjustment")),
    Log = #table(type table [Source = text, Check = text, Rows = Int64.Type, #"What was done" = text], {{
        {{"Business checking", "Rows in the export file", Chk, "Read as text first, so nothing is lost on import"}},
        {{"Business checking", "Title lines above the header (bank name, period, blank)", 3, "Skipped"}},
        {{"Business checking", "Debit and Credit as text with $ and commas", Chk, "Turned into one signed amount (in +, out -)"}},
        {{"Business savings", "Rows in the export file", Sav, "Read as text first"}},
        {{"Business savings", "Signed amounts with thousand separators, ISO dates", Sav, "Converted to numbers and dates"}},
        {{"Visa card", "Rows in the export file", Card, "Read as text first"}},
        {{"Visa card", "Bought in one month, posted in the next", Later, "Booked in the POST month, like the statement"}},
        {{"Visa card", "Card issuer's own categories", Card, "Ignored: the Rules tab decides the account"}},
        {{"All registers", "Descriptions in mixed case or with * . - # &", Typed, "Cleaned before the rules look at them"}},
        {{"All registers", "Lines given an account by a rule", Ruled, "Posted to the ledger"}},
        {{"All registers", "Lines no rule knows", Table.RowCount(T) - Ruled, "Posted to Uncategorized and listed on Review"}},
        {{"All registers", "Transfer and card-payment halves", Legs, "Paired so they are not counted twice"}},
        {{"All registers", "Pairs found", Paired, "Both halves cancel out in Transfers in transit"}},
        {{"Adjustments", "Adjusting entry lines", Adj, "Posted to the ledger as typed"}}}})
in
    Log"""


Q["Payees"] = """let
    // who gets paid: every cost line (COGS or expense) grouped by payee, one column per month, biggest first
    T       = Table.SelectRows(Transactions, each [Type] = "COGS" or [Type] = "Expense"),
    Labels  = List.Transform(Excel.CurrentWorkbook(){[Name = "MonthList"]}[Content][Column1], Text.From),
    Nos     = List.Transform(Excel.CurrentWorkbook(){[Name = "MonthNos"]}[Content][Column1], Number.From),
    G       = Table.Group(T, {"Payee"}, {
                  {"Account", each let a = List.Min([Account]) in Text.From(a) & " " & Table.SelectRows(_, (r) => r[Account] = a){0}[Account name], type text},
                  {"Payments", each Table.RowCount(_), Int64.Type},
                  {"Rows", each _}}),
    Months  = List.Accumulate(List.Positions(Nos), G, (t, i) => Table.AddColumn(t, Labels{i},
                  each Number.Round(-List.Sum(List.Combine({{0}, Table.SelectRows([Rows], (r) => r[MonthNo] = Nos{i})[Amount]})), 2), type number)),
    Total   = Table.AddColumn(Months, "Total", each Number.Round(-List.Sum([Rows][Amount]), 2), type number),
    All     = List.Sum(Total[Total]),
    Share   = Table.AddColumn(Total, "Share", each [Total] / All, type number),
    Sorted  = Table.Sort(Share, {{"Total", Order.Descending}, {"Payee", Order.Ascending}}),
    Result  = Table.RemoveColumns(Sorted, {"Rows"})
in
    Result"""

Q["Cash_Plan_Calc"] = """let
    // the owner's plan for the next 13 weeks (Cash plan tab): one row per regular receipt or payment.
    // The amount comes from the bank history unless the plan types it (fixed).
    Source  = Table.SelectRows(Excel.CurrentWorkbook(){[Name = "CashPlan"]}[Content], each [Item] <> null),
    Typed   = Table.TransformColumnTypes(Source, {{"Order", Int64.Type}, {"Register", Int64.Type}, {"Lookback", Int64.Type},
                  {"First date", type date}, {"Amount", type number}, {"Accounts", type text}}),
    LastDate = List.Max(Transactions[Posted]),
    Start   = Date.AddDays(LastDate, 1),
    End     = Date.AddDays(Start, 13 * 7 - 1),
    Hist    = Table.Buffer(Table.SelectColumns(Transactions, {"Register no", "Account", "Posted", "Amount"})),
    Lines   = (r) => let accts = List.Transform(Text.Split(r[Accounts], ";"), Number.From) in
                  Table.SelectRows(Hist, each [Register no] = r[Register] and List.Contains(accts, [Account])),
    PerTime = (r) =>
        if r[Basis] = "fixed" then r[Amount]
        else if r[Basis] = "avg_week" then
            List.Sum(List.Combine({{0}, Table.SelectRows(Lines(r), each [Posted] > Date.AddDays(LastDate, -7 * r[Lookback]))[Amount]})) / r[Lookback]
        else if r[Basis] = "avg_month" then
            List.Sum(List.Combine({{0}, Table.SelectRows(Lines(r), each [Posted] > Date.EndOfMonth(Date.AddMonths(LastDate, -r[Lookback])))[Amount]})) / r[Lookback]
        else
            // avg_event: the last N dates with a posting (pay runs), averaged
            let g = Table.Group(Lines(r), {"Posted"}, {{"S", each List.Sum([Amount]), type number}}),
                top = Table.FirstN(Table.Sort(g, {{"Posted", Order.Descending}}), r[Lookback])
            in List.Average(top[S]),
    Dates   = (r) =>
        if r[Frequency] = "weekly" then List.Dates(Start, 13, #duration(7, 0, 0, 0))
        else if r[Frequency] = "biweekly" then List.Select(List.Dates(r[First date], 8, #duration(14, 0, 0, 0)), each _ >= Start and _ <= End)
        else List.Select(List.Transform({0..3}, (k) => Date.AddMonths(r[First date], k)), each _ >= Start and _ <= End),
    WithPer = Table.AddColumn(Typed, "Each time", each Number.Round(PerTime(_), 2), type number),
    WithDt  = Table.AddColumn(WithPer, "Dates", each Dates(_)),
    Times   = Table.AddColumn(WithDt, "Times", each List.Count([Dates]), Int64.Type),
    Total   = Table.AddColumn(Times, "Total 13 weeks", each [Each time] * [Times], type number),
    Basis   = Table.AddColumn(Total, "Estimated from", each
                  if [Basis] = "fixed" then "Fixed amount"
                  else if [Basis] = "avg_week" then "Avg week, last " & Text.From([Lookback]) & " wks"
                  else if [Basis] = "avg_month" then "Avg month, last " & Text.From([Lookback])
                  else "Avg of last " & Text.From([Lookback]) & " runs", type text),
    When    = Table.AddColumn(Basis, "When", each
                  if [Frequency] = "weekly" then "Every week"
                  else if [Frequency] = "biweekly" then "Every 2nd week from " & Date.ToText([First date], "d MMM", "en-US")
                  else "Monthly on day " & Text.From(Date.Day([First date])), type text),
    Result  = Table.SelectColumns(When, {"Order", "Item", "Cash flow line", "Estimated from", "When", "Each time", "Times", "Total 13 weeks", "Dates"})
in
    Result"""

Q["Forecast"] = """let
    // one row per expected receipt (+) or payment (-) in the 13 weeks after the last bank line
    Start   = Date.AddDays(List.Max(Transactions[Posted]), 1),
    Rows    = Table.SelectColumns(Cash_Plan_Calc, {"Order", "Item", "Cash flow line", "Each time", "Dates"}),
    Exp     = Table.ExpandListColumn(Rows, "Dates"),
    Kept    = Table.SelectRows(Exp, each [Dates] <> null),
    Ren     = Table.RenameColumns(Kept, {{"Dates", "Date"}, {"Each time", "Amount"}}),
    Week    = Table.AddColumn(Ren, "Week no", each Number.IntegerDivide(Duration.Days([Date] - Start), 7) + 1, Int64.Type),
    Typed   = Table.TransformColumnTypes(Week, {{"Date", type date}, {"Amount", type number}}),
    Sorted  = Table.Sort(Typed, {{"Date", Order.Ascending}, {"Order", Order.Ascending}})
in
    Sorted"""

Q["Cash_Plan_View"] = """let
    Result = Table.RemoveColumns(Cash_Plan_Calc, {"Dates"})
in
    Result"""

# ---------------------------------------------------------------- sheets
# (sheet title, query loaded on it, cell where the query table starts)
LOADED = [("Summary", None, None), ("P&L", None, None), ("Trial balance", None, None), ("Balance sheet", None, None),
          ("Cash flow", None, None), ("Bank rec", None, None), ("Review", "Review", "B8"), ("Month-end close", None, None),
          ("Payees", "Payees", "B8"), ("Ledger", "Ledger", "A1"), ("Transactions", "Transactions", "A1"),
          ("Transfers", "Transfers", "A1"), ("Statements", "Statements", "A1"), ("Forecast", "Forecast", "A1"),
          ("Accounts", None, None), ("Rules", None, None), ("Budget", None, None), ("Adjustments", None, None),
          ("Cash plan", "Cash_Plan_View", "A18"), ("Cleaning log", "Cleaning_Log", "A1"), ("Calc", None, None),
          ("Settings", None, None), ("How to use", None, None), ("Pickers", None, None)]
# navigation menu on every report sheet: (sheet, glyph, label)
NAV = [("Summary", "⌂", "Summary"), ("P&L", "∑", "P&L"), ("Trial balance", "⚖", "Trial balance"),
       ("Balance sheet", "▤", "Balance sheet"), ("Cash flow", "⇅", "Cash flow"), ("Bank rec", "⇌", "Bank rec"),
       ("Review", "⚑", "Review"), ("Month-end close", "✔", "Close"), ("Payees", "☰", "Payees")]

# dark luxury palette (same as E9)
BG0, PAGE, CARD, CARD2, LINE = (11, 15, 19), (18, 23, 29), (26, 33, 41), (32, 40, 50), (42, 51, 61)
GOLD, GOLD_LT, IVORY, MUTED = (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163)
JADE, COPPER, SKY, BUDGETC = (79, 163, 140), (181, 117, 74), (127, 167, 201), (74, 85, 96)
GOOD, BAD, AMBER = (91, 191, 138), (229, 115, 95), (217, 164, 65)
NETROW = (44, 38, 27)          # dark gold tint for the net profit row
# light working sheets
NAVY = (31, 58, 90)
YELLOW = (255, 255, 153)
MONEY = '#,##0.00;[Red]-#,##0.00'
MONEY0 = '#,##0;[Red]-#,##0;"–"'
D0 = '#,##0;-#,##0;"–"'                 # dark sheets: colours come from conditional formats, not [Red]
D2 = '#,##0.00;-#,##0.00;"–"'
D2Z = '#,##0.00;-#,##0.00;"0.00"'
TYPES = {a[0]: a[2] for a in ACCOUNTS}
NAMES = {a[0]: a[1] for a in ACCOUNTS}
COST_ACCTS = [a for a, *_ in ACCOUNTS if TYPES[a] in ("COGS", "Expense")]
PL_ACCTS = [a for a, *_ in ACCOUNTS if TYPES[a] in ("Revenue", "COGS", "Expense")]
HEAD_ROWS = 5      # rows 1-5: brand, title, subtitle, menu, gold rule
TOP = 7            # first content row on a report sheet


def rgb(r, g, b):
    return r + g * 256 + b * 65536


def col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def colno(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def conn_string(name):
    return f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=""'


def load_query(ws, name, table_name, at="A1"):
    lo = ws.ListObjects.Add(0, conn_string(name), None, 1, ws.Range(at))
    qt = lo.QueryTable
    qt.CommandType = 2
    qt.CommandText = f"SELECT * FROM [{name}]"
    qt.RefreshStyle = 1
    qt.AdjustColumnWidth = False
    qt.BackgroundQuery = False
    qt.Refresh(False)
    lo.Name = table_name
    lo.TableStyle = "TableStyleMedium2"
    return lo


def fnt(rng, color=IVORY, size=None, bold=None, name=None, italic=None):
    f = rng.Font
    f.Color = rgb(*color)
    if size:
        f.Size = size
    if bold is not None:
        f.Bold = bold
    if name:
        f.Name = name
    if italic is not None:
        f.Italic = italic


def fill(rng, color):
    rng.Interior.Color = rgb(*color)


def line(rng, edge, color=LINE, weight=2):
    b = rng.Borders(edge)
    b.LineStyle = 1
    b.Color = rgb(*color)
    b.Weight = weight


def status_fmt(rng, open_state=False):
    """PASS / FAIL (/ OPEN) cells: a tick or cross in front, green / red / amber."""
    rng.NumberFormat = '"✔  "@'
    rng.Font.Bold = True
    rng.HorizontalAlignment = -4108
    fc = rng.FormatConditions.Add(1, 3, '="PASS"')
    fc.Font.Color = rgb(*GOOD)
    fc = rng.FormatConditions.Add(1, 3, '="FAIL"')
    fc.Font.Color = rgb(*BAD)
    fc.NumberFormat = '"✖  "@'
    if open_state:
        fc = rng.FormatConditions.Add(1, 3, '="OPEN"')
        fc.Font.Color = rgb(*AMBER)
        fc.NumberFormat = '"●  "@'


def good_bad(rng, good_when=">"):
    """Variance cells: + green, - red (dark sheets)."""
    fc = rng.FormatConditions.Add(1, 5, "=0")
    fc.Font.Color = rgb(*GOOD)
    fc = rng.FormatConditions.Add(1, 6, "=0")
    fc.Font.Color = rgb(*BAD)


def neg_red(rng):
    fc = rng.FormatConditions.Add(1, 6, "=0")
    fc.Font.Color = rgb(*BAD)


def ignore_errors(ws, area):
    for cell in ws.Range(area):
        try:
            cell.Errors(4).Ignore = True   # inconsistent formula: subtotal rows differ from their neighbours on purpose
        except Exception:
            pass


def page(ws, area, landscape=True):
    ps = ws.PageSetup
    ps.Orientation = 2 if landscape else 1
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    ps.PrintArea = area
    ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = 18
    ps.CenterHorizontally = False


def width_pt(ws, c1, c2):
    return sum(ws.Columns(i).Width for i in range(colno(c1), colno(c2) + 1))


def premium(ws, key, title, sub, last_col, last_row):
    """Dark page, brand line, serif title, navigation menu (current sheet in gold) and a thin gold rule.
    last_col is the right margin column: the page is A..last_col; the menu fits between B and the column before it."""
    inner_last = col(colno(last_col) - 1)
    n, gap, min_w = len(NAV), 5, 84
    avail = width_pt(ws, "B", inner_last)
    need = n * min_w + (n - 1) * gap
    if avail < need:   # widen the right margin so the menu always fits
        ws.Columns(last_col).ColumnWidth = ws.Columns(last_col).ColumnWidth * (1 + (need - avail) / ws.Columns(last_col).Width)
        avail = need
        inner_last = last_col
    fill(ws.Range(f"A1:{last_col}4"), BG0)
    fill(ws.Range(f"A5:{last_col}5"), GOLD)
    for r, h in ((1, 16), (2, 30), (3, 17), (4, 30), (5, 2.25), (6, 9)):
        ws.Rows(r).RowHeight = h
    ws.Range("B1").Value = "C O P P E R L E A F"
    fnt(ws.Range("B1"), GOLD, 8.5, True)
    ws.Range("B1").VerticalAlignment = -4107
    ws.Range("B2").Value = title
    fnt(ws.Range("B2"), IVORY, 20, False, "Georgia")
    ws.Range("B2").VerticalAlignment = -4108
    ws.Range("B3").Value = sub
    fnt(ws.Range("B3"), MUTED, 9)
    info = ws.Range(f"{inner_last}2")
    info.Formula = '=Company&" (fictional)   ·   "&Period&"   ·   USD"'
    info.HorizontalAlignment = -4152
    fnt(info, MUTED, 9)
    # menu
    bw = min(118, (avail - (n - 1) * gap) / n)
    total = n * bw + (n - 1) * gap
    left = ws.Range("B4").Left + (avail - total)          # right-aligned under the company line
    top = ws.Range("A4").Top + 4
    for i, (sheet, glyph, label) in enumerate(NAV):
        shp = ws.Shapes.AddShape(5, left + i * (bw + gap), top, bw, 22)
        shp.Name = f"nav_{i + 1}"
        cur = sheet == key
        shp.Fill.ForeColor.RGB = rgb(*(GOLD if cur else CARD))
        shp.Line.ForeColor.RGB = rgb(*(GOLD if cur else LINE))
        shp.Line.Weight = 0.75
        try:
            shp.Adjustments.SetItem(1, 0.18)
        except Exception:
            pass
        tf = shp.TextFrame2
        tf.MarginLeft = tf.MarginRight = 2
        tf.MarginTop = tf.MarginBottom = 0
        tf.VerticalAnchor = 3
        tf.WordWrap = False
        tr = tf.TextRange
        tr.Text = f"{glyph}  {label}"
        tr.Font.Name = "Segoe UI Symbol"
        tr.Font.Size = 9
        tr.Font.Bold = cur
        tr.Font.Fill.ForeColor.RGB = rgb(*(BG0 if cur else IVORY))
        tr.ParagraphFormat.Alignment = 2
        try:
            tr.Characters(1, 1).Font.Fill.ForeColor.RGB = rgb(*(BG0 if cur else GOLD))
        except Exception:
            try:
                tr.GetCharacters(1, 1).Font.Fill.ForeColor.RGB = rgb(*(BG0 if cur else GOLD))
            except Exception:
                pass
        ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress=f"'{sheet}'!A1", ScreenTip=f"Go to {sheet}")
        shp.Placement = 3
    ws.Activate()
    ws.Application.ActiveWindow.FreezePanes = False
    ws.Application.ActiveWindow.ScrollRow = 1
    ws.Range("A6").Select()
    ws.Application.ActiveWindow.FreezePanes = True


def page_bg(ws, last_col, rows=110):
    """Dark page behind everything; called first so the panels and tables drawn later keep their own fills."""
    whole = ws.Range(f"A1:{last_col}{rows}")
    fill(whole, PAGE)
    fnt(whole, IVORY)


def card(ws, c1, c2, r1, label, formula, fmt, note, glyph=None, big=20):
    """KPI card: label (+ icon), big number, note line. Rows r1..r1+4."""
    box = ws.Range(f"{c1}{r1}:{c2}{r1 + 4}")
    fill(box, CARD)
    line(box, 7, PAGE, 4)                # page-coloured edge = a gap when cards sit side by side
    acc = ws.Shapes.AddShape(1, box.Left + 3, box.Top + 6, 2.5, box.Height - 12)   # gold accent bar
    acc.Fill.ForeColor.RGB = rgb(*GOLD)
    acc.Line.Visible = False
    acc.Placement = 2
    ws.Range(f"{c1}{r1}").Value = label
    fnt(ws.Range(f"{c1}{r1}"), MUTED, 8.5, True)
    ws.Range(f"{c1}{r1}").IndentLevel = 1
    ws.Range(f"{c1}{r1}").VerticalAlignment = -4107
    num = ws.Range(f"{c1}{r1 + 1}:{c2}{r1 + 2}")
    num.Merge()
    num.Formula = formula
    num.NumberFormat = fmt
    num.HorizontalAlignment = -4108
    num.VerticalAlignment = -4108
    fnt(num, IVORY, big, False, "Georgia")
    nt = ws.Range(f"{c1}{r1 + 3}:{c2}{r1 + 3}")
    nt.Merge()
    nt.Formula = note
    nt.HorizontalAlignment = -4108
    fnt(nt, MUTED, 8.5)
    if glyph:
        cell = ws.Range(f"{c2}{r1}")
        d = 17
        shp = ws.Shapes.AddShape(9, cell.Left + cell.Width - d - 6, cell.Top + 4, d, d)
        shp.Fill.ForeColor.RGB = rgb(*CARD)
        shp.Line.ForeColor.RGB = rgb(*GOLD)
        shp.Line.Weight = 1
        tf = shp.TextFrame2
        tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
        tf.VerticalAnchor = 3
        tf.WordWrap = False
        tf.TextRange.Text = glyph
        tf.TextRange.Font.Name = "Segoe UI Symbol"
        tf.TextRange.Font.Size = 8.5
        tf.TextRange.Font.Bold = True
        tf.TextRange.Font.Fill.ForeColor.RGB = rgb(*GOLD)
        tf.TextRange.ParagraphFormat.Alignment = 2
        shp.Placement = 2
    return num, nt


def note_colour(rng, state_ref):
    """Colour a note line from a state cell: 1 good (green), -1 bad (red), 2 attention (amber)."""
    for v, c in ((1, GOOD), (-1, BAD), (2, AMBER)):
        fc = rng.FormatConditions.Add(2, None, f"={state_ref}={v}")
        fc.Font.Color = rgb(*c)


def dark_chart(ch, title=None):
    ch.ChartArea.Format.Fill.Visible = True
    ch.ChartArea.Format.Fill.ForeColor.RGB = rgb(*CARD)
    ch.ChartArea.Format.Line.Visible = False
    ch.PlotArea.Format.Fill.Visible = False
    ch.ChartArea.Font.Color = rgb(*MUTED)
    ch.ChartArea.Font.Size = 9
    ch.ChartArea.Font.Name = "Segoe UI"
    if title:
        ch.HasTitle = True
        ch.ChartTitle.Text = title
        ch.ChartTitle.Font.Color = rgb(*IVORY)
        ch.ChartTitle.Font.Size = 12
        ch.ChartTitle.Font.Name = "Georgia"
        ch.ChartTitle.Font.Bold = False
    else:
        ch.HasTitle = False
    try:
        ax = ch.Axes(2)
        ax.HasMajorGridlines = True
        ax.MajorGridlines.Format.Line.ForeColor.RGB = rgb(*LINE)
        ax.Format.Line.Visible = False
        ax.TickLabels.Font.Color = rgb(*MUTED)
        cx = ch.Axes(1)
        cx.Format.Line.ForeColor.RGB = rgb(*LINE)
        cx.TickLabels.Font.Color = rgb(*MUTED)
        cx.MajorTickMark = -4142
    except Exception:
        pass
    ch.PlotVisibleOnly = False


def panel(ws, area, title=None, title_cell=None):
    fill(ws.Range(area), CARD)
    if title_cell:
        ws.Range(title_cell).Formula = title
        fnt(ws.Range(title_cell), IVORY, 12, False, "Georgia")
        ws.Range(title_cell).IndentLevel = 1
        ws.Range(title_cell).VerticalAlignment = -4108


def th(rng, align_right=True):
    """Column headers inside a dark panel: gold, small, gold rule below."""
    fnt(rng, GOLD, 8.5, True)
    line(rng, 9, GOLD, 2)
    if align_right:
        rng.HorizontalAlignment = -4152


def mno(k):
    return f"INDEX(MonthNos,{k})"


def mend(k):
    return f"INDEX(MonthEnds,{k})"


def bal(acct_ref, k):
    """Balance of an account at the end of month k (debit +): opening + every ledger line up to that month."""
    return (f"(SUMIFS(Accounts[Opening debit],Accounts[Account],{acct_ref})-SUMIFS(Accounts[Opening credit],Accounts[Account],{acct_ref})"
            f"+SUMIFS(Ledger[Net],Ledger[Account],{acct_ref},Ledger[MonthNo],\"<=\"&{mno(k)}))")


def opening(acct_ref):
    return f"(SUMIFS(Accounts[Opening debit],Accounts[Account],{acct_ref})-SUMIFS(Accounts[Opening credit],Accounts[Account],{acct_ref}))"


def put_table(ws, at_row, headers, rows, name, text_cols=(), first_col="A"):
    n = len(headers)
    c0 = colno(first_col)
    last = col(c0 + n - 1)
    for c in text_cols:
        ws.Range(f"{c}{at_row}:{c}{at_row + len(rows)}").NumberFormat = "@"
    ws.Range(f"{first_col}{at_row}:{last}{at_row}").NumberFormat = "@"
    ws.Range(f"{first_col}{at_row}:{last}{at_row}").Value = tuple(headers)
    if rows:
        ws.Range(f"{first_col}{at_row + 1}:{last}{at_row + len(rows)}").Value = tuple(tuple(r) for r in rows)
    lo = ws.ListObjects.Add(1, ws.Range(f"{first_col}{at_row}:{last}{at_row + max(1, len(rows))}"), None, 1)
    lo.Name = name
    lo.TableStyle = "TableStyleMedium2"
    return lo


def intro(ws, row, text, c="A"):
    ws.Range(f"{c}{row}").Value = text
    ws.Range(f"{c}{row}").Font.Italic = True
    ws.Range(f"{c}{row}").Font.Color = rgb(90, 100, 115)


def dark_intro(ws, cell, text):
    ws.Range(cell).Value = text
    fnt(ws.Range(cell), MUTED, 9, italic=True)


# ---------------------------------------------------------------- styles
def build_styles(wb):
    st = wb.TableStyles.Add("Copperleaf Dark")
    st.ShowAsAvailableTableStyle = True
    e = st.TableStyleElements(0)            # whole table
    e.Interior.Color = rgb(*CARD)
    e.Font.Color = rgb(*IVORY)
    b = e.Borders(12)                         # inside horizontal
    b.Color = rgb(*LINE)
    b.LineStyle = 1
    e = st.TableStyleElements(1)            # header row
    e.Interior.Color = rgb(*BG0)
    e.Font.Color = rgb(*GOLD)
    e.Font.Bold = True
    b = e.Borders(9)
    b.Color = rgb(*GOLD)
    b.LineStyle = 1
    e = st.TableStyleElements(5)            # row stripe 1
    e.Interior.Color = rgb(*CARD2)
    sl = wb.TableStyles.Add("Copperleaf Slicer")
    sl.ShowAsAvailableTableStyle = False
    sl.ShowAsAvailableSlicerStyle = True
    sl.TableStyleElements(0).Interior.Color = rgb(*PAGE)
    for idx, (bg, fg, bold, border) in {28: (CARD, IVORY, False, LINE), 29: (CARD, MUTED, False, LINE), 30: (GOLD, BG0, True, GOLD),
                                        31: (GOLD, BG0, True, GOLD), 32: (CARD2, GOLD_LT, False, GOLD), 33: (GOLD_LT, BG0, True, GOLD),
                                        34: (CARD2, GOLD_LT, False, GOLD), 35: (GOLD_LT, BG0, True, GOLD)}.items():
        e = sl.TableStyleElements(idx)
        e.Interior.Color = rgb(*bg)
        e.Font.Color = rgb(*fg)
        e.Font.Bold = bold
        for k in (7, 8, 9, 10):
            try:
                e.Borders(k).Color = rgb(*border)
                e.Borders(k).LineStyle = 1
            except Exception:
                break


# ---------------------------------------------------------------- input tabs
def build_inputs(wb):
    ws = wb.Worksheets("Accounts")
    ws.Range("A1").Value = "Chart of accounts and opening balances (31 Jul 2026)"; ws.Range("A1").Font.Size = 14; ws.Range("A1").Font.Bold = True
    intro(ws, 2, "Yellow = yours to change. Type: Asset, Liability, Equity, Revenue, COGS or Expense. Opening debits must equal opening credits.")
    lo = put_table(ws, 4, ["Account", "Name", "Type", "Opening debit", "Opening credit"],
                   [(a, n, t, d or None, c or None) for a, n, t, d, c in ACCOUNTS], "Accounts")
    lo.ListColumns("Opening debit").DataBodyRange.NumberFormat = MONEY
    lo.ListColumns("Opening credit").DataBodyRange.NumberFormat = MONEY
    lo.DataBodyRange.Interior.Color = rgb(*YELLOW)
    for c, w in zip("ABCDE", (10, 34, 11, 15, 15)):
        ws.Columns(c).ColumnWidth = w

    ws = wb.Worksheets("Rules")
    ws.Range("A1").Value = "Categorisation rules"; ws.Range("A1").Font.Size = 14; ws.Range("A1").Font.Bold = True
    intro(ws, 2, "If the keyword appears in a cleaned description (upper case, punctuation removed), the line goes to that account and payee. "
                 "Top to bottom: the first match wins, so put specific keywords above general ones. Add a row, then Refresh All.")
    lo = put_table(ws, 4, ["Keyword", "Account", "Account name", "Payee", "Why"], [(k, a, None, p, w) for k, a, w, p in RULES], "Rules")
    lo.ListColumns("Account name").DataBodyRange.Formula = '=IFERROR(INDEX(Accounts[Name],MATCH([@Account],Accounts[Account],0)),"unknown account")'
    for c in ("Keyword", "Account", "Payee"):
        lo.ListColumns(c).DataBodyRange.Interior.Color = rgb(*YELLOW)
    for c, w in zip("ABCDE", (24, 10, 30, 24, 60)):
        ws.Columns(c).ColumnWidth = w

    ws = wb.Worksheets("Budget")
    ws.Range("A1").Value = "Monthly budget (profit and loss accounts)"; ws.Range("A1").Font.Size = 14; ws.Range("A1").Font.Bold = True
    intro(ws, 2, "Revenue and costs both as positive amounts.")
    lo = put_table(ws, 4, ["Account", "Name"] + [lbl for _, lbl, _ in MONTHS],
                   [(a, None) + tuple(b) for a, b in BUDGET.items()], "Budget")
    lo.ListColumns("Name").DataBodyRange.Formula = '=IFERROR(INDEX(Accounts[Name],MATCH([@Account],Accounts[Account],0)),"unknown account")'
    for _, lbl, _ in MONTHS:
        lo.ListColumns(lbl).DataBodyRange.NumberFormat = MONEY0
        lo.ListColumns(lbl).DataBodyRange.Interior.Color = rgb(*YELLOW)
    for c, w in zip("ABCDE", (10, 34, 12, 12, 12)):
        ws.Columns(c).ColumnWidth = w

    ws = wb.Worksheets("Adjustments")
    ws.Range("A1").Value = "Adjusting entries"; ws.Range("A1").Font.Size = 14; ws.Range("A1").Font.Bold = True
    intro(ws, 2, "Entries that never pass through a bank account (depreciation, accruals, prepayments). Each entry must balance.")
    lo = put_table(ws, 4, ["Date", "Entry", "Description", "Account", "Debit", "Credit"],
                   [(d.strftime("%Y-%m-%d"), e, t, a, dr or None, cr or None) for d, e, t, a, dr, cr in ADJUSTMENTS], "Adjustments")
    lo.ListColumns("Date").DataBodyRange.NumberFormat = "dd mmm yyyy"
    for i, (d, *_rest) in enumerate(ADJUSTMENTS, 1):  # real dates, not text
        c = lo.ListColumns("Date").DataBodyRange.Cells(i, 1)
        c.Formula = f"=DATE({d.year},{d.month},{d.day})"
        c.Value = c.Value
    lo.ListColumns("Debit").DataBodyRange.NumberFormat = MONEY
    lo.ListColumns("Credit").DataBodyRange.NumberFormat = MONEY
    lo.DataBodyRange.Interior.Color = rgb(*YELLOW)
    for c, w in zip("ABCDEF", (12, 9, 58, 10, 12, 12)):
        ws.Columns(c).ColumnWidth = w

    ws = wb.Worksheets("Cash plan")
    ws.Range("A1").Value = "Cash plan: regular receipts and payments for the 13-week forecast"; ws.Range("A1").Font.Size = 14; ws.Range("A1").Font.Bold = True
    intro(ws, 2, "Basis: fixed (type the Amount), avg_week / avg_month (average of the last N weeks / months in the bank history), "
                 "avg_event (average of the last N pay runs). Accounts: which accounts' history to average, separated by ;  Money in +, out -.")
    rows = [(o, it, cf, reg, acc, b, lb, fr, None, amt, nt) for o, it, cf, reg, acc, b, lb, fr, _fd, amt, nt in CASH_PLAN]
    lo = put_table(ws, 4, ["Order", "Item", "Cash flow line", "Register", "Accounts", "Basis", "Lookback", "Frequency", "First date",
                           "Amount", "Note"], rows, "CashPlan", text_cols=("E",))
    for i, row in enumerate(CASH_PLAN, 1):
        c = lo.ListColumns("First date").DataBodyRange.Cells(i, 1)
        c.Formula = f"=DATE({row[8].year},{row[8].month},{row[8].day})"
        c.Value = c.Value
    lo.ListColumns("First date").DataBodyRange.NumberFormat = "dd mmm yyyy"
    lo.ListColumns("Amount").DataBodyRange.NumberFormat = MONEY
    lo.DataBodyRange.Interior.Color = rgb(*YELLOW)
    for c, w in zip("ABCDEFGHIJK", (7, 38, 30, 9, 22, 10, 9, 10, 12, 11, 70)):
        ws.Columns(c).ColumnWidth = w
    ws.Range("A16").Value = "What the plan gives (Power Query, refreshed with the bank history)"
    ws.Range("A16").Font.Bold = True; ws.Range("A16").Font.Size = 12


def build_settings(wb):
    st = wb.Worksheets("Settings")
    st.Range("A1").Value = "Settings"
    st.Range("A1").Font.Size = 16; st.Range("A1").Font.Bold = True
    rows = [("Data folder", '=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"data\\"', "DataFolder",
             "Folder with the CSV exports. Default: the data folder next to this workbook. Type a path to override."),
            ("Checking export", "checking-2026-08-10.csv", "CheckingFile", "Bank export: 3 title lines, then Date, Description, Debit, Credit, Balance"),
            ("Savings export", "savings-2026-08-10.csv", "SavingsFile", "Posted, Details, Amount (signed)"),
            ("Card export", "visa-2026-08-10.csv", "CardFile", "Transaction Date, Post Date, Description, Category, Type, Amount (charges negative)"),
            ("Statement balances", "statements-2026-08-10.csv", "StatementsFile", "Account, Statement date, Ending balance: copied from each month's statement"),
            ("Business name", "Copperleaf Bakery & Cafe", "Company", "Shown in the report titles"),
            ("Period", "Aug–Oct 2026", "Period", "Shown in the report titles")]
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
    st.Range("A13:D13").Value = ("Month", "MonthNo", "Month end", "Short")
    st.Range("A13:D13").Font.Bold = True
    for i, (no, lbl, end) in enumerate(MONTHS, 14):
        st.Cells(i, 1).NumberFormat = "@"
        st.Cells(i, 1).Value = lbl
        st.Cells(i, 2).Value = no
        st.Cells(i, 3).Formula = f"=DATE({end.year},{end.month},{end.day})"
        st.Cells(i, 3).NumberFormat = "dd mmm yyyy"
        st.Cells(i, 4).Formula = f"=LEFT(A{i},3)"
    n = len(MONTHS)
    wb.Names.Add("MonthList", f"=Settings!$A$14:$A${13 + n}")
    wb.Names.Add("MonthNos", f"=Settings!$B$14:$B${13 + n}")
    wb.Names.Add("MonthEnds", f"=Settings!$C$14:$C${13 + n}")
    wb.Names.Add("MonthShort", f"=Settings!$D$14:$D${13 + n}")
    st.Range("A12").Value = "Months in this workbook (MonthNo = year x 100 + month, used by every report)"
    st.Range("A12").Font.Italic = True
    st.Columns("A").ColumnWidth = 22; st.Columns("B").ColumnWidth = 48; st.Columns("C").ColumnWidth = 90; st.Columns("D").ColumnWidth = 8


def build_pickers(wb):
    """Two small tables the Summary slicers filter: which month, and what to compare it with."""
    ws = wb.Worksheets("Pickers")
    ws.Range("A1").Value = "Slicer tables (the Summary slicers filter these; SUBTOTAL reads the visible row)"
    ws.Range("A3:C3").Value = ("Month", "Idx", "Label")
    for i, (no, lbl, end) in enumerate(MONTHS, 4):
        ws.Range(f"A{i}").NumberFormat = "@"
        ws.Range(f"A{i}").Value = lbl[:3]
        ws.Range(f"B{i}").Value = i - 3
        ws.Range(f"C{i}").NumberFormat = "@"
        ws.Range(f"C{i}").Value = lbl
    lo = ws.ListObjects.Add(1, ws.Range(f"A3:C{3 + len(MONTHS)}"), None, 1)
    lo.Name = "MonthPick"
    r0 = 5 + len(MONTHS) + 2
    ws.Range(f"A{r0}:B{r0}").Value = ("Compare", "Code")
    ws.Range(f"A{r0 + 1}:B{r0 + 2}").Value = (("vs Budget", 2), ("vs Last month", 1))
    lo2 = ws.ListObjects.Add(1, ws.Range(f"A{r0}:B{r0 + 2}"), None, 1)
    lo2.Name = "CompareMode"
    ws.Columns("A").ColumnWidth = 16
    return lo, lo2


# ---------------------------------------------------------------- P&L
def build_pl(wb):
    ws = wb.Worksheets("P&L")
    page_bg(ws, "P")
    for c, w in {"A": 2, "B": 7, "C": 31}.items():
        ws.Columns(c).ColumnWidth = w
    for c in "DEFGHIJKLMNO":
        ws.Columns(c).ColumnWidth = 10.5
    ws.Columns("P").ColumnWidth = 2
    R = TOP + 1   # month header row
    for k, (no, lbl, _) in enumerate(MONTHS + [(0, "Quarter", None)]):
        c1, c3 = chr(ord("D") + 3 * k), chr(ord("D") + 3 * k + 2)
        rng = ws.Range(f"{c1}{R}:{c3}{R}")
        rng.Merge(); rng.NumberFormat = "@"; rng.Value = lbl if k < 3 else "Aug–Oct"; rng.HorizontalAlignment = -4108
        fill(rng, CARD2); fnt(rng, GOLD_LT, 10, True)
        ws.Range(f"{c1}{R + 1}:{c3}{R + 1}").Value = ("Actual", "Budget", "Variance")
    th(ws.Range(f"D{R + 1}:O{R + 1}"))
    ws.Range(f"B{R + 1}").Value = "Acct"; ws.Range(f"C{R + 1}").Value = "Account"
    th(ws.Range(f"B{R + 1}:C{R + 1}"), align_right=False)
    rows = {}
    r = R + 2
    sections = [("Revenue", ["Revenue"], -1, "Total revenue"), ("Cost of goods sold", ["COGS"], 1, "Total cost of goods sold"),
                ("Operating expenses", ["Expense"], 1, "Total operating expenses")]

    def block_cols(k):
        return [chr(ord("D") + 3 * k + i) for i in range(3)]

    first_body = r
    for title, types, sign, total in sections:
        ws.Range(f"B{r}").Value = title
        fnt(ws.Range(f"B{r}"), GOLD, 10, True)
        r += 1
        first = r
        for acct in [a[0] for a in ACCOUNTS if a[2] in types]:
            ws.Range(f"B{r}").Value = acct
            ws.Range(f"C{r}").Formula = f'=IFERROR(INDEX(Accounts[Name],MATCH(B{r},Accounts[Account],0)),"?")'
            for k, (no, lbl, _) in enumerate(MONTHS):
                a, b, v = block_cols(k)
                s = "-" if sign < 0 else ""
                ws.Range(f"{a}{r}").Formula = f"={s}SUMIFS(Ledger[Net],Ledger[Account],$B{r},Ledger[MonthNo],{mno(k + 1)})"
                ws.Range(f"{b}{r}").Formula = f"=SUMIFS(Budget[{lbl}],Budget[Account],$B{r})"
                ws.Range(f"{v}{r}").Formula = f"={a}{r}-{b}{r}" if sign < 0 else f"={b}{r}-{a}{r}"
            ws.Range(f"M{r}").Formula = f"=D{r}+G{r}+J{r}"
            ws.Range(f"N{r}").Formula = f"=E{r}+H{r}+K{r}"
            ws.Range(f"O{r}").Formula = f"=F{r}+I{r}+L{r}"
            fnt(ws.Range(f"B{r}"), MUTED, 9)
            line(ws.Range(f"B{r}:O{r}"), 9, LINE, 1)
            rows[acct] = r
            r += 1
        last = r - 1
        ws.Range(f"C{r}").Value = total
        for c in "DEFGHIJKLMNO":
            ws.Range(f"{c}{r}").Formula = f"=SUM({c}{first}:{c}{last})"
        ws.Range(f"B{r}:O{r}").Font.Bold = True
        line(ws.Range(f"B{r}:O{r}"), 8, MUTED, 2)
        rows[title] = r
        r += 1
        if title == "Cost of goods sold":
            rows["Gross profit"] = r
            ws.Range(f"C{r}").Value = "Gross profit"
            for k in range(4):
                a, b, v = block_cols(k)
                rr, cc = rows["Revenue"], rows["Cost of goods sold"]
                ws.Range(f"{a}{r}").Formula = f"={a}{rr}-{a}{cc}"
                ws.Range(f"{b}{r}").Formula = f"={b}{rr}-{b}{cc}"
                ws.Range(f"{v}{r}").Formula = f"={a}{r}-{b}{r}"
            ws.Range(f"B{r}:O{r}").Font.Bold = True
            fill(ws.Range(f"B{r}:O{r}"), CARD2)
            r += 1
            rows["Gross margin"] = r
            ws.Range(f"C{r}").Value = "Gross margin"
            for k in range(4):
                a, b, v = block_cols(k)
                ws.Range(f"{a}{r}").Formula = f"=IFERROR({a}{r - 1}/{a}{rows['Revenue']},0)"
                ws.Range(f"{b}{r}").Formula = f"=IFERROR({b}{r - 1}/{b}{rows['Revenue']},0)"
                ws.Range(f"{v}{r}").Formula = f"=({a}{r}-{b}{r})*100"
                ws.Range(f"{v}{r}").NumberFormat = '+0.0" pts";-0.0" pts";"–"'
                ws.Range(f"{a}{r}:{b}{r}").NumberFormat = "0.0%"
            fnt(ws.Range(f"C{r}:O{r}"), MUTED, italic=True)
            r += 1
        r += 1
    rows["Net profit"] = r
    ws.Range(f"C{r}").Value = "Net profit"
    for k in range(4):
        a, b, v = block_cols(k)
        g, e = rows["Gross profit"], rows["Operating expenses"]
        ws.Range(f"{a}{r}").Formula = f"={a}{g}-{a}{e}"
        ws.Range(f"{b}{r}").Formula = f"={b}{g}-{b}{e}"
        ws.Range(f"{v}{r}").Formula = f"={a}{r}-{b}{r}"
    nr = ws.Range(f"B{r}:O{r}")
    fill(nr, NETROW)
    fnt(nr, GOLD_LT, 11, True)
    line(nr, 8, GOLD, 2); line(nr, 9, GOLD, 2)
    ws.Rows(r).RowHeight = 21
    r += 1
    rows["Net margin"] = r
    ws.Range(f"C{r}").Value = "Net margin"
    for k in range(4):
        a, b, v = block_cols(k)
        ws.Range(f"{a}{r}").Formula = f"=IFERROR({a}{r - 1}/{a}{rows['Revenue']},0)"
        ws.Range(f"{b}{r}").Formula = f"=IFERROR({b}{r - 1}/{b}{rows['Revenue']},0)"
        ws.Range(f"{v}{r}").Formula = f"=({a}{r}-{b}{r})*100"
        ws.Range(f"{v}{r}").NumberFormat = '+0.0" pts";-0.0" pts";"–"'
        ws.Range(f"{a}{r}:{b}{r}").NumberFormat = "0.0%"
    fnt(ws.Range(f"C{r}:O{r}"), MUTED, italic=True)
    last_row = r
    for rr in range(first_body, last_row + 1):
        if rr in (rows["Gross margin"], rows["Net margin"]):
            continue
        ws.Range(f"D{rr}:O{rr}").NumberFormat = D0
    fill(ws.Range(f"B{R}:C{R}"), PAGE)
    body = ws.Range(f"B{R + 1}:O{last_row}")
    # card behind the table; the gross profit / net rows keep their own fills
    for rr in range(R + 1, last_row + 1):
        if rr not in (rows["Gross profit"], rows["Net profit"]):
            fill(ws.Range(f"B{rr}:O{rr}"), CARD)
    for k in range(4):
        v = block_cols(k)[2]
        good_bad(ws.Range(f"{v}{R + 2}:{v}{last_row}"))
        line(ws.Range(f"{v}{R}:{v}{last_row}"), 10, LINE, 2)
    neg_red(ws.Range(f"D{rows['Net profit']}:O{rows['Net profit']}"))
    for rr in range(R + 2, last_row + 1):
        if rr not in (rows["Gross profit"], rows["Net profit"]):
            fill(ws.Range(f"M{rr}:O{rr}"), CARD2)
    note = last_row + 2
    dark_intro(ws, f"B{note}", "Amounts from the Ledger tab (bank and card lines given an account by the Rules tab, plus the adjusting entries). "
                               "Budget from the Budget tab. Variance: + = better than budget (more revenue or lower cost), - = worse.")
    ignore_errors(ws, f"D{R + 2}:O{last_row}")
    premium(ws, "P&L", "Profit and loss vs budget", "From the general ledger, by month and for the quarter. Green = better than budget, red = worse.",
            "P", note + 1)
    page(ws, f"$A$1:$P${note + 1}")
    return rows


# ---------------------------------------------------------------- trial balance
def build_tb(wb):
    ws = wb.Worksheets("Trial balance")
    page_bg(ws, "M")
    for c, w in {"A": 2, "B": 8, "C": 30, "D": 10}.items():
        ws.Columns(c).ColumnWidth = w
    for c in "EFGHIJKL":
        ws.Columns(c).ColumnWidth = 12.5
    ws.Columns("M").ColumnWidth = 2
    R = TOP + 1
    blocks = [("E", "F", "Opening 31 Jul 2026", 0)] + [(chr(ord("G") + 2 * i), chr(ord("H") + 2 * i), lbl, i + 1)
                                                       for i, (_, lbl, _) in enumerate(MONTHS)]
    ws.Range(f"B{R + 1}:D{R + 1}").Value = ("Acct", "Account", "Type")
    th(ws.Range(f"B{R + 1}:D{R + 1}"), align_right=False)
    for a, b, lbl, k in blocks:
        rng = ws.Range(f"{a}{R}:{b}{R}"); rng.Merge(); rng.Value = lbl; rng.HorizontalAlignment = -4108
        fill(rng, CARD2); fnt(rng, GOLD_LT, 10, True)
        if k:
            ws.Range(f"{a}{R}").Formula = f'="{lbl}  (to "&TEXT(INDEX(MonthEnds,{k}),"d mmm")&")"'
        ws.Range(f"{a}{R + 1}:{b}{R + 1}").Value = ("Debit", "Credit")
        th(ws.Range(f"{a}{R + 1}:{b}{R + 1}"))
    r = R + 2
    for acct, *_ in ACCOUNTS:
        ws.Range(f"B{r}").Value = acct
        ws.Range(f"C{r}").Formula = f'=IFERROR(INDEX(Accounts[Name],MATCH(B{r},Accounts[Account],0)),"?")'
        ws.Range(f"D{r}").Formula = f'=IFERROR(INDEX(Accounts[Type],MATCH(B{r},Accounts[Account],0)),"?")'
        op = opening(f"$B{r}")
        ws.Range(f"E{r}").Formula = f"=MAX({op},0)"
        ws.Range(f"F{r}").Formula = f"=MAX(-{op},0)"
        for a, b, lbl, k in blocks[1:]:
            bb = bal(f"$B{r}", k)
            ws.Range(f"{a}{r}").Formula = f"=MAX({bb},0)"
            ws.Range(f"{b}{r}").Formula = f"=MAX(-{bb},0)"
        fill(ws.Range(f"B{r}:L{r}"), CARD if (r - R) % 2 else CARD2)
        fnt(ws.Range(f"B{r}"), MUTED, 9)
        fnt(ws.Range(f"D{r}"), MUTED, 9)
        r += 1
    last = r - 1
    tot, diff, chk = r, r + 1, r + 2
    ws.Range(f"C{tot}").Value = "Total"
    ws.Range(f"C{diff}").Value = "Debits minus credits"
    ws.Range(f"C{chk}").Value = "Balanced?"
    for a, b, lbl, k in blocks:
        ws.Range(f"{a}{tot}").Formula = f"=SUM({a}{R + 2}:{a}{last})"
        ws.Range(f"{b}{tot}").Formula = f"=SUM({b}{R + 2}:{b}{last})"
        ws.Range(f"{a}{diff}").Formula = f"=ROUND({a}{tot}-{b}{tot},2)"
        ws.Range(f"{a}{chk}").Formula = f'=IF({a}{diff}=0,"PASS","FAIL")'
        status_fmt(ws.Range(f"{a}{chk}"))
        ws.Range(f"{a}{chk}").HorizontalAlignment = -4152
    t = ws.Range(f"B{tot}:L{tot}")
    fill(t, NETROW); fnt(t, GOLD_LT, 10.5, True); line(t, 8, GOLD, 2); line(t, 9, GOLD, 2)
    fill(ws.Range(f"B{diff}:L{chk}"), CARD)
    ws.Range(f"C{diff}:L{chk}").Font.Bold = True
    ws.Range(f"E{R + 2}:L{tot}").NumberFormat = '#,##0.00;-#,##0.00;"–"'
    ws.Range(f"E{diff}:L{diff}").NumberFormat = "#,##0.00"
    dark_intro(ws, f"B{chk + 2}", "Profit and loss accounts start at zero on 1 August (the year so far); last year's profit is inside Owner's equity. "
                                  "Transfers in transit = money that left one account and reached the other after the month end.")
    premium(ws, "Trial balance", "Trial balance at each month end",
            "Opening balances + every ledger line up to the month end. Debits must equal credits.", "M", chk + 3)
    page(ws, f"$A$1:$M${chk + 3}")
    return {"diff": diff, "check": chk, "total": tot}


# ---------------------------------------------------------------- balance sheet
def build_bs(wb):
    ws = wb.Worksheets("Balance sheet")
    page_bg(ws, "O")
    for c, w in {"A": 2, "B": 8, "C": 36, "D": 13.5, "E": 13.5, "F": 13.5, "G": 3}.items():
        ws.Columns(c).ColumnWidth = w
    for c in "HIJKLMN":
        ws.Columns(c).ColumnWidth = 10
    ws.Columns("O").ColumnWidth = 2
    R = TOP + 1
    for k, (_, lbl, _) in enumerate(MONTHS, 1):
        c = chr(ord("C") + k)
        ws.Range(f"{c}{R}").Formula = f'=TEXT(INDEX(MonthEnds,{k}),"d mmm yyyy")'
    ws.Range(f"B{R}").Value = "Acct"
    th(ws.Range(f"D{R}:F{R}"))
    th(ws.Range(f"B{R}:C{R}"), align_right=False)
    rows = {}
    r = R + 1

    def ln(label, acct=None, formula=None, sign=1, bold=False, key=None):
        nonlocal r
        if acct:
            ws.Range(f"B{r}").Value = acct
            fnt(ws.Range(f"B{r}"), MUTED, 9)
        ws.Range(f"C{r}").Value = label
        for k in range(1, 4):
            c = chr(ord("C") + k)
            if formula:
                ws.Range(f"{c}{r}").Formula = "=" + formula(c, k)
            elif acct:
                ws.Range(f"{c}{r}").Formula = ("=-" if sign < 0 else "=") + bal(f"$B{r}", k)
        if bold:
            ws.Range(f"B{r}:F{r}").Font.Bold = True
            line(ws.Range(f"B{r}:F{r}"), 8, MUTED, 2)
        if key:
            rows[key] = r
        r += 1

    def section(t):
        nonlocal r
        ws.Range(f"B{r}").Value = t
        fnt(ws.Range(f"B{r}"), GOLD, 10, True)
        r += 1

    section("Assets")
    a0 = r
    for acct in (1000, 1010, CLEARING, 1500, 1510):
        ln(NAMES[acct], acct, key=acct)
    ln("Total assets", formula=lambda c, k: f"SUM({c}{a0}:{c}{r - 1})", bold=True, key="assets")
    r += 1
    section("Liabilities")
    ln(NAMES[2000], 2000, sign=-1, key=2000)
    ln("Total liabilities", formula=lambda c, k: f"{c}{rows[2000]}", bold=True, key="liab")
    r += 1
    section("Equity")
    e0 = r
    ln(NAMES[3000], 3000, sign=-1, key=3000)
    ln(NAMES[3100], 3100, sign=-1, key=3100)
    ln("Profit for the year so far (from 1 Aug)", formula=lambda c, k: "-(" + "+".join(
        f'SUMIFS(Ledger[Net],Ledger[Type],"{t}",Ledger[MonthNo],"<="&{mno(k)})' for t in ("Revenue", "COGS", "Expense")) + ")", key="profit")
    ln("Total equity", formula=lambda c, k: f"SUM({c}{e0}:{c}{r - 1})", bold=True, key="equity")
    r += 1
    ln("Total liabilities and equity", formula=lambda c, k: f"{c}{rows['liab']}+{c}{rows['equity']}", bold=True, key="le")
    ln("Assets minus (liabilities + equity)", formula=lambda c, k: f"ROUND({c}{rows['assets']}-{c}{rows['le']},2)", key="diff")
    ln("Balanced?", formula=lambda c, k: f'IF({c}{rows["diff"]}=0,"PASS","FAIL")', key="check")
    fill(ws.Range(f"B{R + 1}:F{rows['check']}"), CARD)
    ws.Range(f"D{R + 1}:F{rows['diff']}").NumberFormat = '#,##0.00;-#,##0.00;"–"'
    ws.Range(f"D{rows['diff']}:F{rows['diff']}").NumberFormat = "0.00"
    ws.Range(f"C{rows['diff']}:F{rows['check']}").Font.Bold = True
    status_fmt(ws.Range(f"D{rows['check']}:F{rows['check']}"))
    ws.Range(f"D{rows['check']}:F{rows['check']}").HorizontalAlignment = -4152
    for key in ("assets", "le"):
        t = ws.Range(f"B{rows[key]}:F{rows[key]}")
        fill(t, NETROW); fnt(t, GOLD_LT, 10.5, True)
    # side panel: what the business owns and owes, chart over the panel
    last = rows["check"]
    panel(ws, f"H{R}:N{last}", '="What the business owns and owes"', f"H{R}")
    ws.Range(f"H{R + 1}").Value = "Cash = checking + savings; card = what is owed on the Visa card; equity = what is left for the owner."
    fnt(ws.Range(f"H{R + 1}"), MUTED, 8.5)
    ws.Range(f"H{R + 1}").IndentLevel = 1
    top = ws.Range(f"H{R + 2}")
    co = ws.ChartObjects().Add(top.Left + 4, top.Top, width_pt(ws, "H", "N") - 8, ws.Range(f"H{R + 2}:H{last}").Height - 4)
    co.Name = "bs_chart"
    ch = co.Chart
    ch.ChartType = 51
    helper = last + 4   # hidden helper row: cash = checking + savings
    ws.Range(f"C{helper}").Value = "Cash (chart helper)"
    for c in "DEF":
        ws.Range(f"{c}{helper}").Formula = f"={c}{rows[1000]}+{c}{rows[1010]}"
    ws.Rows(helper).Hidden = True
    for name, rr, color in (("Cash", helper, JADE), ("Card owed", rows[2000], COPPER), ("Total equity", rows["equity"], GOLD)):
        s = ch.SeriesCollection().NewSeries()
        s.Name = name
        s.Values = ws.Range(f"D{rr}:F{rr}")
        s.XValues = ws.Range(f"D{R}:F{R}")
        s.Format.Fill.ForeColor.RGB = rgb(*color)
        s.HasDataLabels = True
        s.DataLabels().NumberFormat = '#,##0.0,"K"'
        s.DataLabels().Font.Size = 8
        s.DataLabels().Font.Color = rgb(*IVORY)
    try:
        ch.SeriesCollection("Cash").PlotOrder = 1
    except Exception:
        pass
    dark_chart(ch)
    ch.HasLegend = True
    ch.Legend.Position = -4160
    ch.Legend.Font.Color = rgb(*IVORY)
    ch.ChartGroups(1).GapWidth = 70
    ch.Axes(2).TickLabels.NumberFormat = '#,##0,"K"'
    premium(ws, "Balance sheet", "Balance sheet at each month end", "Simple balance sheet from the same ledger: it must balance every month.",
            "O", last + 2)
    page(ws, f"$A$1:$O${last + 2}")
    return rows


# ---------------------------------------------------------------- bank rec
def build_rec(wb):
    ws = wb.Worksheets("Bank rec")
    page_bg(ws, "M")
    for c, w in {"A": 2, "B": 32, "C": 13, "D": 9, "E": 17, "F": 17, "G": 13, "H": 12, "I": 3, "J": 13, "K": 13, "L": 13, "M": 2}.items():
        ws.Columns(c).ColumnWidth = w
    R = TOP + 1
    ws.Range(f"B{R}:H{R}").Value = ("Account", "Month end", "Lines", "Balance in books", "Statement balance", "Difference", "Status")
    th(ws.Range(f"C{R}:H{R}"))
    th(ws.Range(f"B{R}"), align_right=False)
    ws.Range(f"H{R}").HorizontalAlignment = -4108
    r = R + 1
    for acct in (1000, 1010, 2000):
        for k, (no, lbl, end) in enumerate(MONTHS, 1):
            ws.Range(f"B{r}").Value = NAMES[acct] if k == 1 else ""
            ws.Range(f"N{r}").Value = acct
            ws.Range(f"C{r}").Formula = f"=INDEX(MonthEnds,{k})"
            ws.Range(f"C{r}").NumberFormat = "dd mmm yyyy"
            ws.Range(f"D{r}").Formula = f"=COUNTIFS(Transactions[Register no],$N{r},Transactions[MonthNo],{mno(k)})"
            ws.Range(f"E{r}").Formula = ("=-" if acct == 2000 else "=") + bal(f"$N{r}", k)
            ws.Range(f"F{r}").Formula = f"=SUMIFS(Statements[Ending balance],Statements[Account no],$N{r},Statements[MonthNo],{mno(k)})"
            ws.Range(f"G{r}").Formula = f"=ROUND(E{r}-F{r},2)"
            ws.Range(f"H{r}").Formula = f'=IF(AND(G{r}=0,COUNTIFS(Statements[Account no],$N{r},Statements[MonthNo],{mno(k)})=1),"PASS","FAIL")'
            fill(ws.Range(f"B{r}:H{r}"), CARD)
            line(ws.Range(f"B{r}:H{r}"), 9, LINE if k < 3 else MUTED, 2)
            r += 1
    fnt(ws.Range(f"B{R + 1}:B{r - 1}"), GOLD_LT, 10, True)
    ws.Columns("N").Hidden = True
    ws.Range(f"E{R + 1}:G{r - 1}").NumberFormat = D2Z
    status_fmt(ws.Range(f"H{R + 1}:H{r - 1}"))
    last = r - 1
    # cards on the right
    cards = [("ACCOUNT-MONTHS AGREED", f'=COUNTIF(H{R + 1}:H{last},"PASS")&" / "&ROWS(H{R + 1}:H{last})', "@",
              '="books = statement, to the cent"'),
             ("TOTAL DIFFERENCE", f"=SUMPRODUCT(ABS(G{R + 1}:G{last}))", "#,##0.00", '="sum of every difference"'),
             ("LINES RECONCILED", f"=SUM(D{R + 1}:D{last})", "#,##0", '="checking + savings + card, by post date"')]
    for i, (lbl, f, fmt, nt) in enumerate(cards):
        r1 = R + i * 6
        num, ntc = card(ws, "J", "L", r1, lbl, f, fmt, nt, glyph=("✔", "=", "#")[i])
        fnt(ntc, GOOD, 8.5)
    for i, t in enumerate(("Lines are counted by POST date, the way the bank and the card company build their statements.",
                           "A transfer still on its way shows up under Transfers in transit (balance sheet), not here.",
                           "Card: the amount owed, as the statement prints it."), 1):
        dark_intro(ws, f"B{last + i + 1}", t)
    end = max(last + 5, R + 18)
    premium(ws, "Bank rec", "Books vs bank statements",
            "The balance in the books must equal the statement: every account, every month.", "M", end)
    page(ws, f"$A$1:$M${end}")
    return (R + 1, last)


# ---------------------------------------------------------------- cash flow + 13-week forecast
def build_cashflow(wb, bs, pl):
    ws = wb.Worksheets("Cash flow")
    page_bg(ws, "N")
    widths = {"A": 2, "B": 40, "C": 12, "D": 12, "E": 12, "F": 14, "G": 3, "H": 8, "I": 11, "J": 11.5, "K": 11.5, "L": 11.5, "M": 13, "N": 2}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    n = len(MONTHS)
    R = TOP
    # buffer input (card 1)
    buf = ws.Range(f"B{R}:B{R + 4}")
    fill(buf, CARD)
    line(buf, 7, PAGE, 4)
    acc = ws.Shapes.AddShape(1, buf.Left + 3, buf.Top + 6, 2.5, buf.Height - 12)
    acc.Fill.ForeColor.RGB = rgb(*GOLD)
    acc.Line.Visible = False
    ws.Range(f"B{R}").Value = "MINIMUM CASH BUFFER  (type yours)"
    fnt(ws.Range(f"B{R}"), GOLD, 8.5, True)
    ws.Range(f"B{R}").IndentLevel = 1
    ws.Range(f"B{R}").VerticalAlignment = -4107
    inp = ws.Range(f"B{R + 1}:B{R + 2}")
    inp.Merge()
    inp.Value = CASH_BUFFER
    inp.NumberFormat = "#,##0"
    inp.HorizontalAlignment = -4108
    inp.VerticalAlignment = -4108
    fill(inp, BG0)
    fnt(inp, GOLD_LT, 18, False, "Georgia")
    for e in (8, 9, 10):
        line(inp, e, GOLD, 2)
    inp.Borders(7).LineStyle = 1
    inp.Borders(7).Color = rgb(*GOLD)
    inp.Locked = False
    inp.IndentLevel = 0
    wb.Names.Add("CashBuffer", f"='Cash flow'!$B${R + 1}")
    ws.Range(f"B{R + 3}").Value = "the forecast warns when cash would fall below it"
    fnt(ws.Range(f"B{R + 3}"), MUTED, 8.5)
    ws.Range(f"B{R + 3}").HorizontalAlignment = -4108
    W0 = 15   # first week row
    WL = W0 + 12
    cash_rng = f"$M${W0}:$M${WL}"
    out_rng = f"$K${W0}:$K${WL}"
    # ---- cash flow statement (left), rows 13+
    T = R + 6           # 13
    panel(ws, f"B{T}:F{T + 21}", '="Cash flow statement (indirect method, from the ledger)"', f"B{T}")
    H = T + 1
    for k, (_, lbl, _) in enumerate(MONTHS):
        ws.Range(f"{chr(ord('C') + k)}{H}").Value = lbl
    ws.Range(f"F{H}").Value = "Aug–Oct"
    ws.Range(f"C{H}:F{H}").NumberFormat = "@"
    for k, (_, lbl, _) in enumerate(MONTHS):
        ws.Range(f"{chr(ord('C') + k)}{H}").Value = lbl
    th(ws.Range(f"C{H}:F{H}"))
    rows = {}
    r = H + 1

    def mv(acct, k):   # ledger movement of an account in month k (debit +)
        return f"SUMIFS(Ledger[Net],Ledger[Account],{acct},Ledger[MonthNo],{mno(k)})"

    def ln(label, key, f_month, bold=False, section=False, indent=1):
        nonlocal r
        ws.Range(f"B{r}").Value = label
        ws.Range(f"B{r}").IndentLevel = indent
        if section:
            fnt(ws.Range(f"B{r}"), GOLD, 10, True)
            ws.Range(f"B{r}").IndentLevel = 1
        elif f_month:
            for k in range(1, 4):
                ws.Range(f"{chr(ord('C') + k - 1)}{r}").Formula = "=" + f_month(k)
            ws.Range(f"F{r}").Formula = f"=SUM(C{r}:E{r})"
        if bold:
            ws.Range(f"B{r}:F{r}").Font.Bold = True
            line(ws.Range(f"B{r}:F{r}"), 8, MUTED, 2)
        if key:
            rows[key] = r
        r += 1

    plr = pl
    ln("Operating activities", None, None, section=True)
    ln("Net profit", "np", lambda k: f"INDEX('P&L'!$D${plr['Net profit']}:$L${plr['Net profit']},1,{3 * k - 2})")
    ln("Add back: depreciation (no cash leaves)", "dep", lambda k: f"-{mv(1510, k)}")
    ln("(Increase) / decrease in transfers in transit", "transit", lambda k: f"-{mv(CLEARING, k)}")
    ln("Increase / (decrease) in card balance owed", "card", lambda k: f"-{mv(2000, k)}")
    ln("Net cash from operating activities", "op", lambda k: "+".join(f"{chr(ord('C') + k - 1)}{rows[x]}" for x in ("np", "dep", "transit", "card")), bold=True)
    ln("Investing activities", None, None, section=True)
    ln("Kitchen equipment bought", "equip", lambda k: f"-{mv(1500, k)}")
    ln("Net cash from investing activities", "inv", lambda k: f"{chr(ord('C') + k - 1)}{rows['equip']}", bold=True)
    ln("Financing activities", None, None, section=True)
    ln("Owner's draws", "draws", lambda k: f"-{mv(3100, k)}")
    ln("Owner's capital put in", "capital", lambda k: f"-{mv(3000, k)}")
    ln("Net cash from financing activities", "fin", lambda k: f"{chr(ord('C') + k - 1)}{rows['draws']}+{chr(ord('C') + k - 1)}{rows['capital']}", bold=True)
    ln("Net change in cash", "change", lambda k: "+".join(f"{chr(ord('C') + k - 1)}{rows[x]}" for x in ("op", "inv", "fin")), bold=True)
    start_f = lambda k: (f"{opening(1000)}+{opening(1010)}" if k == 1 else f"{chr(ord('C') + k - 2)}{r + 1}")
    rows["start"] = r
    ws.Range(f"B{r}").Value = "Cash at the start (checking + savings)"
    ws.Range(f"B{r}").IndentLevel = 1
    r += 1
    rows["end"] = r
    ws.Range(f"B{r}").Value = "Cash at the end"
    ws.Range(f"B{r}").IndentLevel = 1
    for k in range(1, 4):
        c = chr(ord("C") + k - 1)
        ws.Range(f"{c}{rows['start']}").Formula = "=" + (f"{opening(1000)}+{opening(1010)}" if k == 1 else f"{chr(ord('C') + k - 2)}{rows['end']}")
        ws.Range(f"{c}{rows['end']}").Formula = f"={c}{rows['start']}+{c}{rows['change']}"
    ws.Range(f"F{rows['start']}").Formula = f"=C{rows['start']}"
    ws.Range(f"F{rows['end']}").Formula = f"=F{rows['start']}+F{rows['change']}"
    r += 1
    rows["bs"] = r
    ws.Range(f"B{r}").Value = "Cash on the balance sheet"
    ws.Range(f"B{r}").IndentLevel = 1
    for k in range(1, 4):
        c = chr(ord("C") + k - 1)
        bc = chr(ord("D") + k - 1)
        ws.Range(f"{c}{r}").Formula = f"='Balance sheet'!{bc}{bs[1000]}+'Balance sheet'!{bc}{bs[1010]}"
    ws.Range(f"F{r}").Formula = f"=E{r}"
    r += 1
    rows["check"] = r
    ws.Range(f"B{r}").Value = "Cash flow ties to the balance sheet?"
    ws.Range(f"B{r}").IndentLevel = 1
    for c in "CDEF":
        ws.Range(f"{c}{r}").Formula = f'=IF(ROUND({c}{rows["end"]}-{c}{rows["bs"]},2)=0,"PASS","FAIL")'
    status_fmt(ws.Range(f"C{r}:F{r}"))
    ws.Range(f"C{r}:F{r}").HorizontalAlignment = -4152
    ws.Range(f"C{H + 1}:F{rows['bs']}").NumberFormat = D0
    for key in ("np", "change", "op", "fin", "inv"):
        neg_red(ws.Range(f"C{rows[key]}:F{rows[key]}"))
    t = ws.Range(f"B{rows['change']}:F{rows['change']}")
    fill(t, NETROW); fnt(t, GOLD_LT, 10.5, True); line(t, 8, GOLD, 2); line(t, 9, GOLD, 2)
    t = ws.Range(f"B{rows['end']}:F{rows['end']}")
    t.Font.Bold = True
    fnt(ws.Range(f"B{rows['bs']}:F{rows['bs']}"), MUTED, italic=True)
    wb.Names.Add("CashToday", f"='Cash flow'!$E${rows['end']}")
    wb.Names.Add("CashBurn", f"=('Cash flow'!$F${rows['start']}-'Cash flow'!$F${rows['end']})/{n}")
    ignore_errors(ws, f"C{H + 1}:F{rows['check']}")
    cf_end = rows["check"]
    # ---- 13-week table (right)
    panel(ws, f"H{T}:M{T + 21}", '="13-week cash forecast"', f"H{T}")
    ws.Range(f"H{H}:M{H}").Value = ("Week", "Week end", "Money in", "Money out", "Net", "Cash at end")
    th(ws.Range(f"H{H}:M{H}"))
    ws.Range(f"H{H}").HorizontalAlignment = -4131
    assert H + 1 == W0, (H, W0)
    start = "(MAX(Transactions[Posted])+1)"
    for w in range(1, 14):
        rr = W0 + w - 1
        ws.Range(f"H{rr}").Value = f"W{w}"
        ws.Range(f"I{rr}").Formula = f"={start}+{7 * w - 1}"
        ws.Range(f"I{rr}").NumberFormat = "d mmm yyyy"
        ws.Range(f"J{rr}").Formula = f'=SUMIFS(Forecast[Amount],Forecast[Week no],{w},Forecast[Amount],">0")'
        ws.Range(f"K{rr}").Formula = f'=SUMIFS(Forecast[Amount],Forecast[Week no],{w},Forecast[Amount],"<0")'
        ws.Range(f"L{rr}").Formula = f"=J{rr}+K{rr}"
        ws.Range(f"M{rr}").Formula = f"=CashToday+L{rr}" if w == 1 else f"=M{rr - 1}+L{rr}"
        line(ws.Range(f"H{rr}:M{rr}"), 9, LINE, 1)
    ws.Range(f"J{W0}:M{WL}").NumberFormat = D0
    neg_red(ws.Range(f"K{W0}:L{WL}"))
    fnt(ws.Range(f"H{W0}:H{WL}"), GOLD_LT, 9.5, True)
    fnt(ws.Range(f"I{W0}:I{WL}"), MUTED, 9)
    fc = ws.Range(f"M{W0}:M{WL}").FormatConditions.Add(1, 6, "=CashBuffer")
    fc.Font.Color = rgb(*BAD)
    fc = ws.Range(f"H{W0}:M{WL}").FormatConditions.Add(2, None, f"=$M{W0}=MIN({cash_rng})")
    fc.Interior.Color = rgb(*NETROW)
    tr = WL + 1
    ws.Range(f"H{tr}").Value = "13 weeks"
    for c in "JKL":
        ws.Range(f"{c}{tr}").Formula = f"=SUM({c}{W0}:{c}{WL})"
    ws.Range(f"M{tr}").Formula = f"=M{WL}"
    t = ws.Range(f"H{tr}:M{tr}")
    t.NumberFormat = D0
    fnt(t, GOLD_LT, 10, True)
    line(t, 8, GOLD, 2)
    ws.Range(f"H{tr + 2}").Value = "From the Cash plan tab. Highlighted: the lowest week."
    fnt(ws.Range(f"H{tr + 2}"), MUTED, 8.5, italic=True)
    ws.Range(f"H{tr + 2}").IndentLevel = 1
    # ---- chart (right, below) and the plan (left, below)
    C0 = T + 23
    panel(ws, f"H{C0}:M{C0 + 17}", '="Forecast cash at the end of each week vs your buffer"', f"H{C0}")
    top = ws.Range(f"H{C0 + 1}")
    co = ws.ChartObjects().Add(top.Left + 4, top.Top, width_pt(ws, "H", "M") - 8, ws.Range(f"H{C0 + 1}:H{C0 + 17}").Height - 4)
    co.Name = "forecast_chart"
    ch = co.Chart
    ch.ChartType = 65
    s = ch.SeriesCollection().NewSeries()
    s.Name = "Forecast cash"
    s.Values = ws.Range(f"M{W0}:M{WL}")
    s.XValues = ws.Range(f"H{W0}:H{WL}")
    s.Format.Line.ForeColor.RGB = rgb(*GOLD)
    s.Format.Line.Weight = 2.25
    s.MarkerStyle = 8
    s.MarkerSize = 6
    s.MarkerBackgroundColor = rgb(*GOLD)
    s.MarkerForegroundColor = rgb(*GOLD)
    s.HasDataLabels = True
    s.DataLabels().NumberFormat = '#,##0.0,"K"'
    s.DataLabels().Position = 0
    s.DataLabels().Font.Color = rgb(*IVORY)
    s.DataLabels().Font.Size = 8
    s2 = ch.SeriesCollection().NewSeries()
    s2.Name = "Buffer"
    for w in range(W0, WL + 1):
        ws.Range(f"O{w}").Formula = "=CashBuffer"
    ws.Columns("O").Hidden = True
    s2.Values = ws.Range(f"O{W0}:O{WL}")
    s2.XValues = ws.Range(f"H{W0}:H{WL}")
    s2.Format.Line.ForeColor.RGB = rgb(*BAD)
    s2.Format.Line.Weight = 1.75
    s2.Format.Line.DashStyle = 4
    s2.MarkerStyle = -4142
    dark_chart(ch)
    ch.HasLegend = True
    ch.Legend.Position = -4160
    ch.Legend.Font.Color = rgb(*IVORY)
    ch.Axes(2).TickLabels.NumberFormat = '#,##0,"K"'
    ch.Axes(2).MinimumScale = 0
    panel(ws, f"B{C0}:F{C0 + 17}", '="The plan behind it (edit the Cash plan tab, then Refresh All)"', f"B{C0}")
    P = C0 + 1
    ws.Range(f"B{P}:F{P}").Value = ("Item", "Each time", "Times", "13 weeks", "Estimated from")
    th(ws.Range(f"C{P}:E{P}"))
    th(ws.Range(f"B{P}"), align_right=False)
    th(ws.Range(f"F{P}"), align_right=False)
    for i in range(1, len(CASH_PLAN) + 1):
        rr = P + i
        ws.Range(f"B{rr}").Formula = f'=IFERROR(INDEX(CashPlanView[Item],{i}),"")'
        ws.Range(f"C{rr}").Formula = f'=IFERROR(INDEX(CashPlanView[Each time],{i}),"")'
        ws.Range(f"D{rr}").Formula = f'=IFERROR(INDEX(CashPlanView[Times],{i}),"")'
        ws.Range(f"E{rr}").Formula = f'=IFERROR(INDEX(CashPlanView[Total 13 weeks],{i}),"")'
        ws.Range(f"F{rr}").Formula = f'=IFERROR(INDEX(CashPlanView[Estimated from],{i}),"")'
        ws.Range(f"B{rr}").IndentLevel = 1
        line(ws.Range(f"B{rr}:F{rr}"), 9, LINE, 1)
    pl_last = P + len(CASH_PLAN)
    ws.Range(f"C{P + 1}:E{pl_last}").NumberFormat = D0
    neg_red(ws.Range(f"C{P + 1}:E{pl_last}"))
    fnt(ws.Range(f"F{P + 1}:F{pl_last}"), MUTED, 8.5)
    tr2 = pl_last + 1
    ws.Range(f"B{tr2}").Value = "Net change over 13 weeks"
    ws.Range(f"B{tr2}").IndentLevel = 1
    ws.Range(f"E{tr2}").Formula = f"=SUM(E{P + 1}:E{pl_last})"
    ws.Range(f"E{tr2}").NumberFormat = D0
    t = ws.Range(f"B{tr2}:F{tr2}")
    fnt(t, GOLD_LT, 10, True)
    line(t, 8, GOLD, 2)
    neg_red(ws.Range(f"E{tr2}"))
    # ---- KPI cards (rows 7-11), written last so every name they use exists
    low_pos = f"MATCH(MIN({cash_rng}),{cash_rng},0)"
    cards = [("C", "D", "CASH TODAY", "=CashToday", "#,##0",
              '="checking + savings, "&TEXT(INDEX(MonthEnds,ROWS(MonthEnds)),"d mmm")', "◈", "FcToday"),
             ("E", "F", "LOWEST IN 13 WEEKS", f"=MIN({cash_rng})", "#,##0",
              f'="in W"&{low_pos}&" · week to "&TEXT(INDEX($I${W0}:$I${WL},{low_pos}),"d mmm")', "▼", "FcLowest"),
             ("H", "I", "HEADROOM", f"=MIN({cash_rng})-CashBuffer", "+#,##0;-#,##0;0",
              f'=IF(COUNTIF({cash_rng},"<"&CashBuffer)=0,"no week under the buffer",COUNTIF({cash_rng},"<"&CashBuffer)&" weeks under the buffer")',
              "↑", "FcHeadroom"),
             ("J", "K", "WEEKS OF CASH", f"=CashToday/(-SUM({out_rng})/13)", '0.0" wks"',
              f'="vs "&TEXT(-SUM({out_rng})/13,"#,##0")&" paid out a week"', "⌛", "FcWeeks"),
             ("L", "M", "RUNWAY", '=IF(CashBurn>0,CashToday/CashBurn,0)', '0.0" months"',
              '=IF(CashBurn>0,"cash fell "&TEXT(CashBurn,"#,##0")&" a month, Aug–Oct","cash grew Aug–Oct")', "⧖", "FcRunway")]
    for c1, c2, lbl, f, fmt, nt, g, nm in cards:
        num, ntc = card(ws, c1, c2, R, lbl, f, fmt, nt, glyph=g, big=18)
        wb.Names.Add(nm, f"='Cash flow'!${c1}${R + 1}")
        if nm == "FcHeadroom":
            neg_red(num)
            fc = num.FormatConditions.Add(1, 5, "=0")
            fc.Font.Color = rgb(*GOOD)
    # helper cells (hidden column O): names must not contain a list separator (it differs per locale)
    ws.Range(f"O{R}").Formula = f"={low_pos}"
    ws.Range(f"O{R + 1}").Formula = f"=INDEX($I${W0}:$I${WL},{low_pos})"
    wb.Names.Add("FcLowWeek", f"='Cash flow'!$O${R}")
    wb.Names.Add("FcLowDate", f"='Cash flow'!$O${R + 1}")
    end = C0 + 18
    premium(ws, "Cash flow", "Cash flow and 13-week forecast",
            "Where the cash went (indirect method, from the ledger) and where it is going: the owner's plan, priced from the bank history.",
            "N", end)
    page(ws, f"$A$1:$N${end}")
    rows.update(week0=W0, weekL=WL, cards=R)
    return rows


# ---------------------------------------------------------------- month-end close
def build_close(wb, tb, bs, rec, cf):
    ws = wb.Worksheets("Month-end close")
    page_bg(ws, "I")
    for c, w in {"A": 2, "B": 5, "C": 50, "D": 14, "E": 14, "F": 14, "G": 3, "H": 70, "I": 2}.items():
        ws.Columns(c).ColumnWidth = w
    R = TOP + 1
    dark_intro(ws, f"B{TOP}", "Ticked by formulas after every Refresh All. PASS = done and proven, OPEN = needs a person (see Review), FAIL = something does not agree.")
    ws.Range(f"B{R}:H{R}").Value = ("#", "Close step", "", "", "", "", "How it is checked")
    for k, (_, lbl, _) in enumerate(MONTHS):
        ws.Range(f"{chr(ord('D') + k)}{R}").Value = lbl
    ws.Range(f"D{R}:F{R}").NumberFormat = "@"
    for k, (_, lbl, _) in enumerate(MONTHS):
        ws.Range(f"{chr(ord('D') + k)}{R}").Value = lbl
    th(ws.Range(f"B{R}:H{R}"), align_right=False)
    ws.Range(f"D{R}:F{R}").HorizontalAlignment = -4108
    r1, r2 = rec

    def rec_row(acct, k):
        return r1 + (0 if acct == 1000 else 3 if acct == 1010 else 6) + (k - 1)

    def start(k):
        return f"(EOMONTH({mend(k)},-1)+1)"

    def transit(k):
        e = mend(k)
        return (f"ROUND(SUMIFS(Transfers[Amount],Transfers[From date],\"<=\"&{e},Transfers[To date],\">\"&{e})"
                f"+SUMIFS(Transfers[Amount],Transfers[From date],\"<=\"&{e},Transfers[To line],\"\")"
                f"-SUMIFS(Transfers[Amount],Transfers[To date],\"<=\"&{e},Transfers[From date],\">\"&{e})"
                f"-SUMIFS(Transfers[Amount],Transfers[To date],\"<=\"&{e},Transfers[From line],\"\")"
                f"-{bal(CLEARING, k)},2)=0")

    items = [
        ("Bank, savings and card exports loaded for the month",
         lambda k: "AND(" + ",".join(f"COUNTIFS(Transactions[Register no],{a},Transactions[MonthNo],{mno(k)})>0" for a in (1000, 1010, 2000)) + ")",
         "Each of the three registers has lines in the month"),
        ("Business checking agrees with its statement", lambda k: f"'Bank rec'!$H${rec_row(1000, k)}=\"PASS\"", "Bank rec: books = statement to the cent"),
        ("Business savings agrees with its statement", lambda k: f"'Bank rec'!$H${rec_row(1010, k)}=\"PASS\"", "Bank rec: books = statement to the cent"),
        ("Visa card agrees with its statement", lambda k: f"'Bank rec'!$H${rec_row(2000, k)}=\"PASS\"", "Bank rec: amount owed = statement"),
        ("Transfers paired, money in transit explained", transit,
         "Transfers in transit at month end = halves whose other side posts later (Transfers tab)"),
        ("Accruals and depreciation posted, entries balance",
         lambda k: (f"AND(COUNTIFS(Adjustments[Date],\">=\"&{start(k)},Adjustments[Date],\"<=\"&{mend(k)})>0,"
                    f"ROUND(SUMIFS(Adjustments[Debit],Adjustments[Date],\">=\"&{start(k)},Adjustments[Date],\"<=\"&{mend(k)})"
                    f"-SUMIFS(Adjustments[Credit],Adjustments[Date],\">=\"&{start(k)},Adjustments[Date],\"<=\"&{mend(k)}),2)=0)"),
         "Adjustments tab has entries dated in the month; debits = credits"),
        ("Trial balance difference 0.00", lambda k: f"'Trial balance'!${chr(ord('G') + 2 * (k - 1))}${tb['check']}=\"PASS\"",
         "Trial balance: debits = credits at the month end"),
        ("Balance sheet balances", lambda k: f"'Balance sheet'!${chr(ord('C') + k)}${bs['check']}=\"PASS\"", "Assets = liabilities + equity"),
        ("Cash flow statement ties to the balance sheet", lambda k: f"'Cash flow'!${chr(ord('B') + k)}${cf['check']}=\"PASS\"",
         "Closing cash on the cash flow = checking + savings"),
        ("Review cleared: every line has a real account",
         None, "Lines no rule knows (6990 Uncategorized) left in the month: ask the owner, add a rule"),
    ]
    r = R + 1
    first = r
    for i, (label, f, how) in enumerate(items, 1):
        ws.Range(f"B{r}").Value = i
        ws.Range(f"C{r}").Value = label
        ws.Range(f"H{r}").Value = how
        for k in range(1, 4):
            c = chr(ord("D") + k - 1)
            if f:
                ws.Range(f"{c}{r}").Formula = f'=IF({f(k)},"PASS","FAIL")'
            else:
                ws.Range(f"{c}{r}").Formula = f'=IF(COUNTIFS(Transactions[Account],{UNCATEGORIZED},Transactions[MonthNo],{mno(k)})=0,"PASS","OPEN")'
        fill(ws.Range(f"B{r}:H{r}"), CARD if i % 2 else CARD2)
        fnt(ws.Range(f"B{r}"), GOLD, 9, True)
        ws.Range(f"B{r}").HorizontalAlignment = -4108
        fnt(ws.Range(f"H{r}"), MUTED, 9)
        ws.Rows(r).RowHeight = 20
        r += 1
    last = r - 1
    status_fmt(ws.Range(f"D{first}:F{last}"), open_state=True)
    det = r
    ws.Range(f"C{det}").Value = "Still on Review (uncategorized lines)"
    for k in range(1, 4):
        c = chr(ord("D") + k - 1)
        ws.Range(f"{c}{det}").Formula = (f'=COUNTIFS(Transactions[Account],{UNCATEGORIZED},Transactions[MonthNo],{mno(k)})&" · "&'
                                         f'TEXT(-SUMIFS(Transactions[Amount],Transactions[Account],{UNCATEGORIZED},Transactions[MonthNo],{mno(k)}),"#,##0.00")')
    ws.Range(f"D{det}:F{det}").HorizontalAlignment = -4108
    fnt(ws.Range(f"C{det}:F{det}"), AMBER, 9)
    fill(ws.Range(f"B{det}:H{det}"), CARD)
    done = det + 1
    ws.Range(f"C{done}").Value = "Steps done"
    stat = done + 1
    ws.Range(f"C{stat}").Value = "Close status"
    for k in range(1, 4):
        c = chr(ord("D") + k - 1)
        ws.Range(f"{c}{done}").Formula = f'=COUNTIF({c}{first}:{c}{last},"PASS")&" of "&ROWS({c}{first}:{c}{last})'
        ws.Range(f"{c}{stat}").Formula = (f'=IF(COUNTIF({c}{first}:{c}{last},"FAIL")>0,"Fix FAIL first",'
                                          f'IF(COUNTIF({c}{first}:{c}{last},"OPEN")>0,"Open items","Closed"))')
    t = ws.Range(f"B{done}:H{stat}")
    fill(t, NETROW)
    fnt(t, GOLD_LT, 10.5, True)
    line(ws.Range(f"B{done}:H{done}"), 8, GOLD, 2)
    ws.Range(f"D{done}:F{stat}").HorizontalAlignment = -4108
    fc = ws.Range(f"D{stat}:F{stat}").FormatConditions.Add(1, 3, '="Closed"'); fc.Font.Color = rgb(*GOOD)
    fc = ws.Range(f"D{stat}:F{stat}").FormatConditions.Add(1, 3, '="Open items"'); fc.Font.Color = rgb(*AMBER)
    fc = ws.Range(f"D{stat}:F{stat}").FormatConditions.Add(1, 3, '="Fix FAIL first"'); fc.Font.Color = rgb(*BAD)
    for e in (7, 10):
        line(ws.Range(f"D{R}:F{stat}"), e, LINE, 2)
    dark_intro(ws, f"B{stat + 2}", "Close order: load the exports, reconcile each account, pair transfers, post the adjusting entries, "
                                   "prove the trial balance and balance sheet, then clear Review. Only then send the P&L.")
    premium(ws, "Month-end close", "Month-end close checklist", "Ten close steps per month, each proven by a formula: nothing is ticked by hand.",
            "I", stat + 3)
    page(ws, f"$A$1:$I${stat + 3}")
    return dict(first=first, last=last, detail=det, done=done, status=stat, labels=[x[0] for x in items])


# ---------------------------------------------------------------- Calc: everything the Summary shows, in plain cells
def build_calc(wb, pl, bs, rec, tb, cf, close):
    ws = wb.Worksheets("Calc")
    ws.Range("A1").Value = "Calc: the numbers behind the Summary (follows the two slicers)"
    ws.Range("A1").Font.Size = 14; ws.Range("A1").Font.Bold = True
    for c, w in zip("ABCDEFGHIJ", (44, 30, 14, 14, 62, 9, 12, 12, 12, 16)):
        ws.Columns(c).ColumnWidth = w

    def head(r, labels):
        ws.Range(f"A{r}:{col(len(labels))}{r}").Value = tuple(labels)
        ws.Range(f"A{r}:{col(len(labels))}{r}").Font.Bold = True
        ws.Range(f"A{r}:{col(len(labels))}{r}").Interior.Color = rgb(221, 228, 236)

    sel = [("Month picked (index)", "=IF(SUBTOTAL(104,MonthPick[Idx])=0,ROWS(MonthNos),SUBTOTAL(104,MonthPick[Idx]))", "SelIdx"),
           ("Month", "=INDEX(MonthList,SelIdx)", "SelMonth"),
           ("Month short", "=INDEX(MonthShort,SelIdx)", "SelShort"),
           ("MonthNo", "=INDEX(MonthNos,SelIdx)", "SelNo"),
           ("Month end", "=INDEX(MonthEnds,SelIdx)", "SelEnd"),
           ("Compare with (2 = budget, 1 = last month)", "=IF(SUBTOTAL(104,CompareMode[Code])=0,2,SUBTOTAL(104,CompareMode[Code]))", "ModeNo"),
           ("Is there something to compare with?", "=OR(ModeNo=2,SelIdx>1)", "HasComp"),
           ("Compared with", '=IF(ModeNo=2,"budget",IF(SelIdx>1,INDEX(MonthShort,SelIdx-1),"-"))', "CompLbl")]
    for i, (lbl, f, nm) in enumerate(sel, 3):
        ws.Range(f"A{i}").Value = lbl
        ws.Range(f"B{i}").Formula = f
        wb.Names.Add(nm, f"=Calc!$B${i}")
    ws.Range("B7").NumberFormat = "dd mmm yyyy"

    def act(row, idx="SelIdx"):
        return f"INDEX('P&L'!$D${row}:$L${row},1,3*{idx}-2)"

    def bud(row):
        return f"INDEX('P&L'!$D${row}:$L${row},1,3*SelIdx-1)"

    def comp(row):
        return f'IF(ModeNo=2,{bud(row)},IF(SelIdx>1,{act(row, "(SelIdx-1)")},""))'

    def bsv(row, idx="SelIdx"):
        return f"INDEX('Balance sheet'!$D${row}:$F${row},1,{idx})"

    # ---- KPI cards
    K = 12
    head(K - 1, ["KPI", "Actual", "Compared with", "Better by", "Note on the card", "State"])
    R_, N_, G_, E_ = pl["Revenue"], pl["Net profit"], pl["Gross profit"], pl["Operating expenses"]
    arrow = lambda r: f'IF(B{r}>=C{r},"▲ ","▼ ")'
    kp = [
        ("Revenue", f"={act(R_)}", f"={comp(R_)}", f'=IF(C{K}="","",B{K}-C{K})',
         f'=IF(C{K}="","no earlier month in this file",{arrow(K)}&TEXT(ABS(B{K}-C{K}),"#,##0")&" vs "&CompLbl&" "&TEXT(C{K},"#,##0"))'),
        ("Net profit", f"={act(N_)}", f"={comp(N_)}", f'=IF(C{K + 1}="","",B{K + 1}-C{K + 1})',
         f'=IF(C{K + 1}="","no earlier month in this file",{arrow(K + 1)}&TEXT(ABS(B{K + 1}-C{K + 1}),"#,##0")&" vs "&CompLbl&" "&TEXT(C{K + 1},"#,##0;-#,##0"))'),
        ("Gross margin", f"=IFERROR({act(G_)}/{act(R_)},0)", f'=IF({comp(R_)}="","",IFERROR({comp(G_)}/{comp(R_)},0))',
         f'=IF(C{K + 2}="","",B{K + 2}-C{K + 2})',
         f'=IF(C{K + 2}="","no earlier month in this file",{arrow(K + 2)}&TEXT(ABS(B{K + 2}-C{K + 2})*100,"0.0")&" pts vs "&CompLbl&" "&TEXT(C{K + 2},"0.0%"))'),
        ("Operating costs % of revenue", f"=IFERROR({act(E_)}/{act(R_)},0)", f'=IF({comp(R_)}="","",IFERROR({comp(E_)}/{comp(R_)},0))',
         f'=IF(C{K + 3}="","",C{K + 3}-B{K + 3})',
         f'=IF(C{K + 3}="","no earlier month in this file",{arrow(K + 3)}&TEXT(ABS(B{K + 3}-C{K + 3})*100,"0.0")&" pts vs "&CompLbl&" "&TEXT(C{K + 3},"0.0%"))'),
        ("Cash at month end", f"={bsv(bs[1000])}+{bsv(bs[1010])}",
         f"=IF(SelIdx=1,{opening(1000)}+{opening(1010)},{bsv(bs[1000], '(SelIdx-1)')}+{bsv(bs[1010], '(SelIdx-1)')})",
         f"=B{K + 4}-C{K + 4}", f'={arrow(K + 4)}&TEXT(ABS(B{K + 4}-C{K + 4}),"#,##0")&" in the month"'),
        ("Cash in transit", f"={bsv(bs[CLEARING])}", "", "",
         f'=IF(ROUND(B{K + 5},2)=0,"nothing between accounts","on its way between accounts")'),
        ("Lines to review", f'=COUNTIFS(Review[Posted],">="&(EOMONTH(SelEnd,-1)+1),Review[Posted],"<="&SelEnd)',
         f'=-SUMIFS(Review[Amount],Review[Posted],">="&(EOMONTH(SelEnd,-1)+1),Review[Posted],"<="&SelEnd)', "",
         f'=IF(B{K + 6}=0,"nothing to review",B{K + 6}&IF(B{K + 6}=1," line"," lines")&" · "&TEXT(C{K + 6},"#,##0.00")&" on Review")'),
        ("Weeks of cash (today)", "=FcWeeks", "=FcLowest", "=FcHeadroom",
         f'="lowest ahead: "&TEXT(C{K + 7},"#,##0")&" (W"&FcLowWeek&")"'),
    ]
    for i, (lbl, a, c, d, e) in enumerate(kp):
        r = K + i
        ws.Range(f"A{r}").Value = lbl
        for cc, f in zip("BCDE", (a, c, d, e)):
            if f:
                ws.Range(f"{cc}{r}").Formula = f
        ws.Range(f"F{r}").Formula = f'=IF(D{r}="",0,IF(D{r}>=0,1,-1))'
    ws.Range(f"F{K + 5}").Formula = f"=IF(ROUND(B{K + 5},2)=0,1,2)"
    ws.Range(f"F{K + 6}").Formula = f"=IF(B{K + 6}=0,1,2)"
    ws.Range(f"B{K}:D{K + 7}").NumberFormat = "#,##0.00"
    ws.Range(f"B{K + 2}:D{K + 3}").NumberFormat = "0.0%"
    wb.Names.Add("CalcKpi", f"=Calc!$A${K}:$F${K + 7}")

    # ---- profit waterfall
    W = 23
    head(W - 1, ["Step", "Start", "End", "Base (hidden bar)", "Up", "Down", "Label position", "Label"])
    steps = [("Revenue", "0", f"={act(R_)}"), ("Cost of goods", f"=C{W}", f"=C{W}-{act(pl['Cost of goods sold'])}"),
             ("Operating costs", f"=C{W + 1}", f"=C{W + 1}-{act(E_)}"), ("Net profit", "0", f"={act(N_)}")]
    for i, (lbl, s, e) in enumerate(steps):
        r = W + i
        ws.Range(f"A{r}").Value = lbl
        ws.Range(f"B{r}").Formula = s if s.startswith("=") else "=" + s
        ws.Range(f"C{r}").Formula = e
        lo, hi = f"MIN(B{r}:C{r})", f"MAX(B{r}:C{r})"
        ws.Range(f"D{r}").Formula = f"=IF({lo}>=0,{lo},IF({hi}<=0,{hi},0))"
        ws.Range(f"E{r}").Formula = f"=IF({lo}>=0,{hi}-{lo},IF({hi}<=0,0,{hi}))"
        ws.Range(f"F{r}").Formula = f"=IF({lo}>=0,0,IF({hi}<=0,{lo}-{hi},{lo}))"
        ws.Range(f"G{r}").Formula = f"={hi}"
        ws.Range(f"H{r}").Formula = f'=TEXT(C{r}-B{r},"#,##0;-#,##0")' if i < 3 else f'=TEXT(C{r},"#,##0;-#,##0")'
    ws.Range(f"B{W}:G{W + 3}").NumberFormat = "#,##0.00"
    wb.Names.Add("CalcWaterfall", f"=Calc!$A${W}:$H${W + 3}")

    # ---- every P&L account for the month (costs sorted below)
    A0 = 30
    head(A0 - 1, ["Account", "Name", "Type", "This month", "Compared with", "Better by", "Month 1", "Month 2", "Month 3", "Sort key (costs)"])
    acc_rows = {}
    for i, a in enumerate(PL_ACCTS):
        r = A0 + i
        acc_rows[a] = r
        pr = pl[a]
        ws.Range(f"A{r}").Value = a
        ws.Range(f"B{r}").Formula = f'=IFERROR(INDEX(Accounts[Name],MATCH(A{r},Accounts[Account],0)),"?")'
        ws.Range(f"C{r}").Value = TYPES[a]
        ws.Range(f"D{r}").Formula = f"={act(pr)}"
        ws.Range(f"E{r}").Formula = f"={comp(pr)}"
        ws.Range(f"F{r}").Formula = (f'=IF(E{r}="","",D{r}-E{r})' if TYPES[a] == "Revenue" else f'=IF(E{r}="","",E{r}-D{r})')
        for k in range(3):
            ws.Range(f"{'GHI'[k]}{r}").Formula = f"='P&L'!{chr(ord('D') + 3 * k)}{pr}"
        ws.Range(f"J{r}").Formula = f'=IF(C{r}="Revenue","",ROUND(D{r},2)-A{r}/10^7)'
    AL = A0 + len(PL_ACCTS) - 1
    ws.Range(f"D{A0}:I{AL}").NumberFormat = "#,##0.00"
    ws.Range(f"J{A0}:J{AL}").NumberFormat = "0.0000000"
    S0 = AL + 4
    head(S0 - 1, ["Rank", "Account", "This month", "Better by", "Month 1", "Month 2", "Month 3"])
    keys = f"$J${A0}:$J${AL}"
    for k in range(1, len(COST_ACCTS) + 1):
        r = S0 + k - 1
        pos = f"MATCH(LARGE({keys},{k}),{keys},0)"
        ws.Range(f"A{r}").Value = k
        ws.Range(f"B{r}").Formula = f'=INDEX($A${A0}:$A${AL},{pos})&" "&INDEX($B${A0}:$B${AL},{pos})'
        ws.Range(f"C{r}").Formula = f"=INDEX($D${A0}:$D${AL},{pos})"
        ws.Range(f"D{r}").Formula = f"=INDEX($F${A0}:$F${AL},{pos})"
        for j in range(3):
            ws.Range(f"{'EFG'[j]}{r}").Formula = f"=INDEX(${'GHI'[j]}${A0}:${'GHI'[j]}${AL},{pos})"
    SL = S0 + len(COST_ACCTS) - 1
    ws.Range(f"C{S0}:G{SL}").NumberFormat = "#,##0.00"
    wb.Names.Add("CalcCosts", f"=Calc!$A${S0}:$G${SL}")

    # ---- who got paid this month (top 6, other, total)
    P0 = SL + 4
    head(P0 - 1, ["Rank", "Payee", "Spend", "Payments", "Share of the month"])
    vals = "INDEX(Payees[#Data],0,3+SelIdx)"
    keyarr = f"INDEX({vals}-ROW(Payees[Payee])/10^7,0)"
    for k in range(1, 7):
        r = P0 + k - 1
        pos = f"MATCH(LARGE({keyarr},{k}),{keyarr},0)"
        ws.Range(f"A{r}").Value = k
        ws.Range(f"B{r}").Formula = f'=IFERROR(INDEX(Payees[Payee],{pos}),"")'
        ws.Range(f"C{r}").Formula = f'=IFERROR(INDEX({vals},{pos}),0)'
        ws.Range(f"D{r}").Formula = (f'=IF(B{r}="",0,COUNTIFS(Transactions[Payee],B{r},Transactions[MonthNo],SelNo,Transactions[Type],"COGS")'
                                     f'+COUNTIFS(Transactions[Payee],B{r},Transactions[MonthNo],SelNo,Transactions[Type],"Expense"))')
        ws.Range(f"E{r}").Formula = f"=IFERROR(C{r}/C{P0 + 7},0)"
    ws.Range(f"B{P0 + 6}").Value = "Other payees"
    ws.Range(f"C{P0 + 6}").Formula = f"=C{P0 + 7}-SUM(C{P0}:C{P0 + 5})"
    ws.Range(f"D{P0 + 6}").Formula = (f'=COUNTIFS(Transactions[MonthNo],SelNo,Transactions[Type],"COGS")+COUNTIFS(Transactions[MonthNo],SelNo,Transactions[Type],"Expense")'
                                      f"-SUM(D{P0}:D{P0 + 5})")
    ws.Range(f"E{P0 + 6}").Formula = f"=IFERROR(C{P0 + 6}/C{P0 + 7},0)"
    ws.Range(f"B{P0 + 7}").Value = "Total spend"
    ws.Range(f"C{P0 + 7}").Formula = f"=SUM({vals})"
    ws.Range(f"D{P0 + 7}").Formula = f"=SUM(D{P0}:D{P0 + 6})"
    ws.Range(f"E{P0 + 7}").Formula = f"=IFERROR(C{P0 + 7}/C{P0 + 7},0)"
    ws.Range(f"C{P0}:C{P0 + 7}").NumberFormat = "#,##0.00"
    ws.Range(f"E{P0}:E{P0 + 7}").NumberFormat = "0.0%"
    wb.Names.Add("CalcPayees", f"=Calc!$A${P0}:$E${P0 + 7}")

    # ---- by month (chart)
    M0 = P0 + 11
    head(M0 - 1, ["Month", "Revenue", "Total costs", "Net profit", "Budget net profit"])
    for k in range(3):
        r = M0 + k
        a, b = chr(ord("D") + 3 * k), chr(ord("E") + 3 * k)
        ws.Range(f"A{r}").Formula = f"=INDEX(MonthList,{k + 1})"
        ws.Range(f"B{r}").Formula = f"='P&L'!{a}{R_}"
        ws.Range(f"C{r}").Formula = f"='P&L'!{a}{pl['Cost of goods sold']}+'P&L'!{a}{E_}"
        ws.Range(f"D{r}").Formula = f"='P&L'!{a}{N_}"
        ws.Range(f"E{r}").Formula = f"='P&L'!{b}{N_}"
    ws.Range(f"B{M0}:E{M0 + 2}").NumberFormat = "#,##0.00"
    wb.Names.Add("CalcMonths", f"=Calc!$A${M0}:$E${M0 + 2}")

    # ---- key insights
    I0 = M0 + 6
    head(I0 - 1, ["Topic", "What the numbers say", "State"])
    better = f"$F${A0}:$F${AL}"
    names_ = f"$A${A0}:$A${AL}"
    full = lambda pos: f'INDEX({names_},{pos})&" "&INDEX($B${A0}:$B${AL},{pos})'
    ins = [("Profit", f'=SelMonth&": "&IF({act(N_)}>=0,"profit ","loss ")&TEXT(ABS({act(N_)}),"#,##0")&IF(HasComp," vs "&CompLbl&" "&TEXT($C${K + 1},"#,##0;-#,##0")&" ("&TEXT($D${K + 1},"+#,##0;-#,##0")&")","")',
            f"=$F${K + 1}"),
           ("Biggest overrun", f'=IF(HasComp,{full(f"MATCH(MIN({better}),{better},0)")}&": "&TEXT(MIN({better}),"+#,##0;-#,##0")&" vs "&CompLbl,"no earlier month to compare with")',
            f"=IF(HasComp,IF(MIN({better})<0,-1,1),0)"),
           ("Biggest saving", f'=IF(HasComp,{full(f"MATCH(MAX({better}),{better},0)")}&": "&TEXT(MAX({better}),"+#,##0;-#,##0")&" vs "&CompLbl,"")',
            f"=IF(HasComp,IF(MAX({better})>0,1,-1),0)"),
           ("Cash", f'="Cash "&IF($D${K + 4}>=0,"up ","down ")&TEXT(ABS($D${K + 4}),"#,##0")&" in "&SelShort&" to "&TEXT($B${K + 4},"#,##0")'
                    f'&IF(ROUND($B${K + 5},2)>0,", "&TEXT($B${K + 5},"#,##0")&" in transit","")', f"=$F${K + 4}"),
           ("Next 13 weeks", '="Lowest cash "&TEXT(FcLowest,"#,##0")&" in W"&FcLowWeek&" ("&TEXT(FcLowDate,"d mmm")&"), "&TEXT(ABS(FcHeadroom),"#,##0")'
                             '&IF(FcHeadroom>=0," above"," below")&" the "&TEXT(CashBuffer,"#,##0")&" buffer"', "=IF(FcHeadroom>=0,1,-1)")]
    for i, (t, f, s) in enumerate(ins):
        r = I0 + i
        ws.Range(f"A{r}").Value = t
        ws.Range(f"B{r}").Formula = f
        ws.Range(f"C{r}").Formula = s
    wb.Names.Add("CalcInsights", f"=Calc!$A${I0}:$C${I0 + len(ins) - 1}")

    # ---- the eight data checks
    C0 = I0 + len(ins) + 3
    head(C0 - 1, ["Data check", "Status", "Short label (Summary)"])
    log = "INDEX(CleaningLog[Rows],{})".format
    r1, r2 = rec
    checks = [
        ("Opening balances balance (debits = credits at 31 Jul)", "Opening balances balance",
         '=IF(ROUND(SUM(Accounts[Opening debit])-SUM(Accounts[Opening credit]),2)=0,"PASS","FAIL")'),
        ("Trial balance balances at every month end", "Trial balance 0.00 every month end",
         f"=IF(COUNTIF('Trial balance'!$G${tb['check']}:$L${tb['check']},\"PASS\")=3,\"PASS\",\"FAIL\")"),
        ("Every bank and card line posted once (lines in the files = ledger)", "Every bank line posted exactly once",
         f'=IF(AND(ROWS(Transactions[Line])={log(1)}+{log(4)}+{log(6)},COUNTIFS(Ledger[Source],"Bank line")=2*ROWS(Transactions[Line])),"PASS","FAIL")'),
        ("Every line has a known account; lines no rule knows are all on Review", "Every line has an account or is on Review",
         f'=IF(AND(SUMPRODUCT(--ISNA(MATCH(Transactions[Account],Accounts[Account],0)))=0,'
         f'COUNTIFS(Transactions[Account],{UNCATEGORIZED})=COUNTIFS(Review[Issue],"No rule matched")),"PASS","FAIL")'),
        ("Transfers: what is still in transit = the unpaired halves on Review", "In transit = unpaired halves on Review",
         f"=IF(ROUND('Balance sheet'!$F${bs[CLEARING]}+SUMIFS(Review[Amount],Review[Issue],\"Transfer without its other half\"),2)=0,\"PASS\",\"FAIL\")"),
        ("Books = bank and card statements (3 accounts x 3 months)", "Books = all 9 bank and card statements",
         f"=IF(COUNTIF('Bank rec'!$H${r1}:$H${r2},\"PASS\")=9,\"PASS\",\"FAIL\")"),
        ("Balance sheet balances at every month end", "Balance sheet balances every month end",
         f"=IF(COUNTIF('Balance sheet'!$D${bs['check']}:$F${bs['check']},\"PASS\")=3,\"PASS\",\"FAIL\")"),
        ("Adjusting entries balance (debits = credits)", "Adjusting entries balance",
         '=IF(ROUND(SUM(Adjustments[Debit])-SUM(Adjustments[Credit]),2)=0,"PASS","FAIL")'),
    ]
    for i, (label, short, f) in enumerate(checks):
        r = C0 + i
        ws.Range(f"A{r}").Value = label
        ws.Range(f"B{r}").Formula = f
        ws.Range(f"C{r}").Value = short
    CL = C0 + len(checks) - 1
    ws.Range(f"A{CL + 1}").Value = "Checks passing / total"
    ws.Range(f"B{CL + 1}").Formula = f'=COUNTIF(B{C0}:B{CL},"PASS")'
    ws.Range(f"C{CL + 1}").Formula = f"=ROWS(B{C0}:B{CL})"
    wb.Names.Add("CheckLabels", f"=Calc!$A${C0}:$A${CL}")
    wb.Names.Add("CheckStatus", f"=Calc!$B${C0}:$B${CL}")
    wb.Names.Add("CheckShort", f"=Calc!$C${C0}:$C${CL}")

    # ---- the close checklist for the picked month
    L0 = CL + 5
    head(L0 - 1, ["Close step", "Status (picked month)"])
    cs = "'Month-end close'"
    for i in range(close["last"] - close["first"] + 1):
        r = L0 + i
        rr = close["first"] + i
        ws.Range(f"A{r}").Formula = f"={cs}!$C${rr}"
        ws.Range(f"B{r}").Formula = f"=INDEX({cs}!$D${rr}:$F${rr},1,SelIdx)"
    LL = L0 + close["last"] - close["first"]
    ws.Range(f"A{LL + 1}").Value = "Steps done"
    ws.Range(f"B{LL + 1}").Formula = f"=INDEX({cs}!$D${close['done']}:$F${close['done']},1,SelIdx)"
    ws.Range(f"A{LL + 2}").Value = "Close status"
    ws.Range(f"B{LL + 2}").Formula = f"=INDEX({cs}!$D${close['status']}:$F${close['status']},1,SelIdx)"
    wb.Names.Add("CloseSel", f"=Calc!$A${L0}:$B${LL + 2}")
    return dict(K=K, W=W, S0=S0, SL=SL, P0=P0, M0=M0, I0=I0, C0=C0, CL=CL, L0=L0, LL=LL)


# ---------------------------------------------------------------- Summary dashboard
def U(k, j=0):
    """column letter of unit j (0-2) of card slot k (0-7) on the Summary grid"""
    return col(2 + 4 * k + j)


def build_summary(wb, calc, pickers):
    ws = wb.Worksheets("Summary")
    page_bg(ws, "AG")
    ws.Columns("A").ColumnWidth = 2
    for k in range(8):
        for j in range(3):
            ws.Columns(U(k, j)).ColumnWidth = 8.3
        ws.Columns(col(5 + 4 * k)).ColumnWidth = 1.6
    ws.Columns(col(33)).ColumnWidth = 2
    LAST = col(33)                            # AG
    END = 49
    for r, h in {7: 15, 8: 17, 9: 17, 10: 9, 11: 18, 12: 19, 13: 19, 14: 16, 15: 6, 16: 10, 17: 25, 18: 18, 32: 17, 33: 10,
                 34: 25, 48: 8, 49: 14}.items():
        ws.Rows(r).RowHeight = h
    for r in list(range(19, 32)) + list(range(35, 48)):
        ws.Rows(r).RowHeight = 17.5
    premium(ws, "Summary", "Bookkeeping summary",
            "Bank and card exports categorised by rules, posted to a ledger, checked against every statement.", LAST, END)
    c = calc
    # ---- pickers
    for cell, txt in (("B7", "MONTH"), (f"{U(3)}7", "COMPARE WITH")):
        ws.Range(cell).Value = txt
        fnt(ws.Range(cell), GOLD, 8.5, True)
    ws.Range(f"{U(5)}8").Value = "Pick a month and what to compare it with: every card, chart and table below follows."
    fnt(ws.Range(f"{U(5)}8"), MUTED, 9)
    ws.Range(f"{U(5)}9").Value = "Green = better, red = worse, amber = needs a person."
    fnt(ws.Range(f"{U(5)}9"), MUTED, 9)
    lo_m, lo_c = pickers
    top = ws.Range("B8").Top
    h = ws.Range("B8:B9").Height
    for lo, field, nm, c1, c2, ncol in ((lo_m, "Month", "Month picker", U(0), U(2, 2), 3), (lo_c, "Compare", "Compare with", U(3), U(4, 2), 2)):
        sc = wb.SlicerCaches.Add2(lo, field)
        left = ws.Range(f"{c1}8").Left
        wdt = width_pt(ws, c1, c2)
        sl = sc.Slicers.Add(SlicerDestination=ws, Name=nm, Caption=nm, Top=top, Left=left, Width=wdt, Height=h)
        for prop, val in (("DisplayHeader", False), ("NumberOfColumns", ncol), ("Style", "Copperleaf Slicer"),
                          ("RowHeight", h - 8), ("ColumnWidth", (wdt - 8 - 4 * (ncol - 1)) / ncol)):
            try:
                setattr(sl, prop, val)
            except Exception as e:
                print("slicer", prop, e)
    # ---- KPI cards
    K = c["K"]
    cards = [("REVENUE", "#,##0", "$"), ("NET PROFIT", "#,##0;-#,##0", "="), ("GROSS MARGIN", "0.0%", "%"),
             ("OPEX % OF REVENUE", "0.0%", "◔"), ("CASH AT MONTH END", "#,##0", "◈"), ("CASH IN TRANSIT", "#,##0;-#,##0;0", "⇄"),
             ("LINES TO REVIEW", "0", "!"), ("WEEKS OF CASH", '0.0" wks"', "⌛")]
    for k, (lbl, fmt, g) in enumerate(cards):
        num, nt = card(ws, U(k, 0), U(k, 2), 11, lbl, f"=Calc!$B${K + k}", fmt, f"=Calc!$E${K + k}", glyph=g)
        note_colour(nt, f"Calc!$F${K + k}")
        if k == 1:
            neg_red(num)
        if k == 6:
            fc = num.FormatConditions.Add(1, 5, "=0"); fc.Font.Color = rgb(*AMBER)
            fc = num.FormatConditions.Add(1, 3, "=0"); fc.Font.Color = rgb(*GOOD)
        if k == 5:
            fc = num.FormatConditions.Add(1, 5, "=0"); fc.Font.Color = rgb(*AMBER)
    # ---- panel 1: profit waterfall
    P1, P2, P3 = (U(0), U(2, 2)), (U(3), U(5, 2)), (U(6), U(7, 2))
    panel(ws, f"{P1[0]}17:{P1[1]}32", '="Profit waterfall  ·  "&SelMonth', "B17")
    W = c["W"]
    t = ws.Range("B18")
    co = ws.ChartObjects().Add(t.Left + 4, t.Top, width_pt(ws, P1[0], P1[1]) - 8, ws.Range("B18:B32").Height - 4)
    co.Name = "waterfall"
    ch = co.Chart
    ch.ChartType = 52
    calc_ws = wb.Worksheets("Calc")
    for j, nm in enumerate(("Base", "Up", "Down")):
        s = ch.SeriesCollection().NewSeries()
        s.Name = nm
        s.Values = calc_ws.Range(f"{'DEF'[j]}{W}:{'DEF'[j]}{W + 3}")
        s.XValues = calc_ws.Range(f"A{W}:A{W + 3}")
    s4 = ch.SeriesCollection().NewSeries()
    s4.Name = "Label"
    s4.Values = calc_ws.Range(f"G{W}:G{W + 3}")
    s4.XValues = calc_ws.Range(f"A{W}:A{W + 3}")
    s4.ChartType = 65
    s4.Format.Line.Visible = False
    s4.MarkerStyle = -4142
    ch.SeriesCollection(1).Format.Fill.Visible = False
    ch.SeriesCollection(1).Format.Line.Visible = False
    colors = [JADE, COPPER, COPPER, GOLD]
    for sidx in (2, 3):
        for p in range(1, 5):
            pt = ch.SeriesCollection(sidx).Points(p)
            pt.Format.Fill.ForeColor.RGB = rgb(*(BAD if (sidx == 3 and p == 4) else colors[p - 1]))
            pt.Format.Line.Visible = False
    s4.HasDataLabels = True
    dl = s4.DataLabels()
    try:
        dl.Format.TextFrame2.TextRange.InsertChartField(7, f"=Calc!$H${W}:$H${W + 3}")
        dl.ShowRange = True
        dl.ShowValue = False
    except Exception as e:
        print("waterfall labels", e)
    dl.Position = 0
    dl.Font.Color = rgb(*IVORY)
    dl.Font.Size = 10
    dl.Font.Bold = True
    ch.ChartGroups(1).GapWidth = 55
    dark_chart(ch)
    ch.HasLegend = False
    ch.Axes(2).TickLabels.NumberFormat = '#,##0,"K";-#,##0,"K";0'
    ch.Axes(1).TickLabelPosition = -4134
    ch.Axes(1).TickLabels.Font.Color = rgb(*IVORY)
    # ---- panel 2: costs by account, sorted, with data bars and sparklines
    panel(ws, f"{P2[0]}17:{P2[1]}32", '="Costs by account  ·  "&SelMonth', f"{P2[0]}17")
    hdr = 18
    ws.Range(f"{U(3)}{hdr}").Value = "Account"
    ws.Range(f"{U(4)}{hdr}").Formula = "=SelMonth"
    ws.Range(f"{U(4, 2)}{hdr}").Formula = '="vs "&CompLbl'
    ws.Range(f"{U(5)}{hdr}").Value = "Aug–Oct trend"
    th(ws.Range(f"{U(3)}{hdr}:{U(5, 2)}{hdr}"), align_right=False)
    ws.Range(f"{U(4)}{hdr}:{U(4, 2)}{hdr}").HorizontalAlignment = -4152
    ws.Range(f"{U(5)}{hdr}").HorizontalAlignment = -4108
    S0 = c["S0"]
    n = len(COST_ACCTS)
    for i in range(n):
        r = 19 + i
        ws.Range(f"{U(3)}{r}").Formula = f"=Calc!$B${S0 + i}"
        ws.Range(f"{U(3)}{r}").IndentLevel = 1
        m = ws.Range(f"{U(4)}{r}:{U(4, 1)}{r}")
        m.Merge()
        m.Formula = f"=Calc!$C${S0 + i}"
        m.NumberFormat = D0
        ws.Range(f"{U(4, 2)}{r}").Formula = f"=Calc!$D${S0 + i}"
        ws.Range(f"{U(4, 2)}{r}").NumberFormat = '+#,##0;-#,##0;0'
        ws.Range(f"{U(5)}{r}:{U(5, 2)}{r}").Merge()
        line(ws.Range(f"{U(3)}{r}:{U(5, 2)}{r}"), 9, LINE, 1)
    fnt(ws.Range(f"{U(3)}19:{U(3)}{18 + n}"), IVORY, 9)
    good_bad(ws.Range(f"{U(4, 2)}19:{U(4, 2)}{18 + n}"))
    db = ws.Range(f"{U(4)}19:{U(4, 1)}{18 + n}").FormatConditions.AddDatabar()
    db.BarColor.Color = rgb(*COPPER)
    db.BarFillType = 0
    db.MinPoint.Modify(0, 0)
    try:
        db.BarBorder.Type = 0
    except Exception:
        pass
    sg = ws.Range(f"{U(5)}19:{U(5)}{18 + n}").SparklineGroups.Add(1, f"Calc!$E${S0}:$G${S0 + n - 1}")
    sg.SeriesColor.Color = rgb(*GOLD)
    sg.LineWeight = 1.5
    try:
        sg.Points.Markers.Visible = True
        sg.Points.Markers.Color.Color = rgb(*GOLD_LT)
    except Exception as e:
        print("sparkline", e)
    ws.Range(f"{U(3)}32").Formula = '="Sorted by "&SelMonth&"; bar = amount, line = the three months. Variance + = better."'
    fnt(ws.Range(f"{U(3)}32"), MUTED, 8, italic=True)
    ws.Range(f"{U(3)}32").IndentLevel = 1
    # ---- panel 3: who got paid + key insights
    panel(ws, f"{P3[0]}17:{P3[1]}32", '="Who got paid  ·  "&SelMonth', f"{P3[0]}17")
    ws.Range(f"{U(6)}{hdr}").Value = "Payee"
    ws.Range(f"{U(6, 2)}{hdr}:{U(7)}{hdr}").Merge()
    ws.Range(f"{U(6, 2)}{hdr}").Value = "Spend"
    ws.Range(f"{U(7, 1)}{hdr}").Value = "Paid"
    ws.Range(f"{U(7, 2)}{hdr}").Value = "Share"
    th(ws.Range(f"{U(6)}{hdr}:{U(7, 2)}{hdr}"))
    ws.Range(f"{U(6)}{hdr}").HorizontalAlignment = -4131
    P0 = c["P0"]
    for i in range(8):
        r = 19 + i
        ws.Range(f"{U(6)}{r}").Formula = f"=Calc!$B${P0 + i}"
        ws.Range(f"{U(6)}{r}").IndentLevel = 0
        m = ws.Range(f"{U(6, 2)}{r}:{U(7)}{r}")
        m.Merge()
        m.Formula = f"=Calc!$C${P0 + i}"
        m.NumberFormat = D0
        ws.Range(f"{U(7, 1)}{r}").Formula = f"=Calc!$D${P0 + i}"
        ws.Range(f"{U(7, 1)}{r}").NumberFormat = "0×"
        ws.Range(f"{U(7, 2)}{r}").Formula = f"=Calc!$E${P0 + i}"
        ws.Range(f"{U(7, 2)}{r}").NumberFormat = "0%"
        line(ws.Range(f"{U(6)}{r}:{U(7, 2)}{r}"), 9, LINE, 1)
    fnt(ws.Range(f"{U(6)}19:{U(7, 2)}26"), IVORY, 9)
    fnt(ws.Range(f"{U(6)}25:{U(7, 2)}25"), MUTED, 9)
    t = ws.Range(f"{U(6)}26:{U(7, 2)}26")
    fnt(t, GOLD_LT, 9.5, True)
    line(t, 8, GOLD, 2)
    db = ws.Range(f"{U(6, 2)}19:{U(7)}24").FormatConditions.AddDatabar()
    db.BarColor.Color = rgb(*BUDGETC)
    db.BarFillType = 0
    db.MinPoint.Modify(0, 0)
    ws.Range(f"{U(6)}27").Value = "Key insights"
    fnt(ws.Range(f"{U(6)}27"), IVORY, 11, False, "Georgia")
    ws.Range(f"{U(6)}27").IndentLevel = 1
    ws.Range(f"{U(6)}27").VerticalAlignment = -4107
    I0 = c["I0"]
    for i in range(5):
        r = 28 + i
        cell = ws.Range(f"{U(6)}{r}")
        cell.Formula = f'="▸ "&Calc!$B${I0 + i}'
        cell.IndentLevel = 1
        fnt(cell, IVORY, 8.5)
        for v, cc in ((1, GOOD), (-1, BAD)):
            fc = cell.FormatConditions.Add(2, None, f"=Calc!$C${I0 + i}={v}")
            fc.Font.Color = rgb(*cc)
    # ---- panel 4: by month chart
    P4, P5 = (U(0), U(3, 2)), (U(4), U(7, 2))
    panel(ws, f"{P4[0]}34:{P4[1]}48", '="Revenue, costs and net profit by month"', "B34")
    M0 = c["M0"]
    t = ws.Range("B35")
    co = ws.ChartObjects().Add(t.Left + 4, t.Top, width_pt(ws, P4[0], P4[1]) - 8, ws.Range("B35:B48").Height - 4)
    co.Name = "by_month"
    ch = co.Chart
    ch.ChartType = 51
    for colx, name, color in (("B", "Revenue", JADE), ("C", "Total costs", COPPER), ("D", "Net profit", GOLD), ("E", "Budget net profit", BUDGETC)):
        s = ch.SeriesCollection().NewSeries()
        s.Name = name
        s.Values = calc_ws.Range(f"{colx}{M0}:{colx}{M0 + 2}")
        s.XValues = calc_ws.Range(f"A{M0}:A{M0 + 2}")
        s.Format.Fill.ForeColor.RGB = rgb(*color)
        s.Format.Line.Visible = False
    for i in (1, 2, 3, 4):
        s = ch.SeriesCollection(i)
        s.HasDataLabels = True
        s.DataLabels().NumberFormat = '#,##0.0,"K";-#,##0.0,"K"'
        s.DataLabels().Font.Size = 8
        s.DataLabels().Font.Color = rgb(*IVORY)
        s.DataLabels().Position = 2   # outside end
    ch.ChartGroups(1).GapWidth = 80
    ch.ChartGroups(1).Overlap = -5
    dark_chart(ch)
    ch.HasLegend = True
    ch.Legend.Position = -4160
    ch.Legend.Font.Color = rgb(*IVORY)
    ch.Axes(2).TickLabels.NumberFormat = '#,##0,"K";-#,##0,"K";0'
    ch.Axes(1).TickLabelPosition = -4134
    ch.Axes(1).TickLabels.Font.Color = rgb(*IVORY)
    # ---- panel 5: data checks + month-end close
    panel(ws, f"{P5[0]}34:{P5[1]}48", '="Data checks"', f"{P5[0]}34")
    ws.Range(f"{U(6)}34").Formula = '="Month-end close  ·  "&SelMonth'
    fnt(ws.Range(f"{U(6)}34"), IVORY, 12, False, "Georgia")
    ws.Range(f"{U(6)}34").VerticalAlignment = -4108
    C0 = c["C0"]
    for i in range(8):
        r = 35 + i
        ws.Range(f"{U(4)}{r}").Formula = f"=Calc!$C${C0 + i}"
        ws.Range(f"{U(4)}{r}").IndentLevel = 1
        ws.Range(f"{U(5, 2)}{r}").Formula = f"=Calc!$B${C0 + i}"
        line(ws.Range(f"{U(4)}{r}:{U(5, 2)}{r}"), 9, LINE, 1)
    fnt(ws.Range(f"{U(4)}35:{U(4)}42"), IVORY, 9)
    status_fmt(ws.Range(f"{U(5, 2)}35:{U(5, 2)}42"))
    ws.Range(f"{U(5, 2)}35:{U(5, 2)}42").HorizontalAlignment = -4152
    big = ws.Range(f"{U(4)}44:{U(5, 2)}46")
    big.Merge()
    big.Formula = f'="✔  "&Calc!$B${c["CL"] + 1}&" of "&Calc!$C${c["CL"] + 1}&" checks PASS"'
    big.HorizontalAlignment = -4108
    big.VerticalAlignment = -4108
    fnt(big, GOOD, 16, False, "Georgia")
    fc = big.FormatConditions.Add(2, None, f'=Calc!$B${c["CL"] + 1}<Calc!$C${c["CL"] + 1}')
    fc.Font.Color = rgb(*BAD)
    ws.Range(f"{U(4)}47").Value = "after every Refresh All: trial balance, statements, transfers, nothing posted twice"
    fnt(ws.Range(f"{U(4)}47"), MUTED, 8, italic=True)
    ws.Range(f"{U(4)}47").IndentLevel = 1
    L0 = c["L0"]
    nclose = c["LL"] - L0 + 1
    for i in range(nclose):
        r = 35 + i
        ws.Range(f"{U(6)}{r}").Formula = f"=Calc!$A${L0 + i}"
        ws.Range(f"{U(6)}{r}").IndentLevel = 1
        ws.Range(f"{U(7, 2)}{r}").Formula = f"=Calc!$B${L0 + i}"
        line(ws.Range(f"{U(6)}{r}:{U(7, 2)}{r}"), 9, LINE, 1)
    fnt(ws.Range(f"{U(6)}35:{U(6)}{34 + nclose}"), IVORY, 8.5)
    status_fmt(ws.Range(f"{U(7, 2)}35:{U(7, 2)}{34 + nclose}"), open_state=True)
    ws.Range(f"{U(7, 2)}35:{U(7, 2)}{34 + nclose}").HorizontalAlignment = -4152
    sr = 35 + nclose
    ws.Range(f"{U(6)}{sr}").Formula = f'="Steps done: "&Calc!$B${c["LL"] + 1}'
    ws.Range(f"{U(7, 2)}{sr}").Formula = f"=Calc!$B${c['LL'] + 2}"
    t = ws.Range(f"{U(6)}{sr}:{U(7, 2)}{sr}")
    fnt(t, GOLD_LT, 9.5, True)
    line(t, 8, GOLD, 2)
    ws.Range(f"{U(6)}{sr}").IndentLevel = 1
    ws.Range(f"{U(7, 2)}{sr}").HorizontalAlignment = -4152
    for v, cc in (("Closed", GOOD), ("Open items", AMBER), ("Fix FAIL first", BAD)):
        fc = ws.Range(f"{U(7, 2)}{sr}").FormatConditions.Add(1, 3, f'="{v}"')
        fc.Font.Color = rgb(*cc)
    ws.Range(f"B{END}").Value = "Fictional data. No macros: the slicers filter two small tables, formulas do the rest. Next month: new exports in, Data > Refresh All."
    fnt(ws.Range(f"B{END}"), MUTED, 8, italic=True)
    page(ws, f"$A$1:${LAST}${END}")
    ws.PageSetup.Orientation = 2


# ---------------------------------------------------------------- loaded tables shown in the report (dark)
def build_payees_sheet(wb):
    ws = wb.Worksheets("Payees")
    page_bg(ws, "K")
    lo = ws.ListObjects("Payees")
    lo.TableStyle = "Copperleaf Dark"
    lo.Range.Interior.ColorIndex = -4142
    fnt(lo.HeaderRowRange, GOLD, None, True)
    fnt(lo.DataBodyRange, IVORY)
    widths = {"A": 2, "B": 28, "C": 34, "D": 11, "E": 12, "F": 12, "G": 12, "H": 13, "I": 9, "J": 16, "K": 2}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    n = lo.ListRows.Count
    first, last = 9, 8 + n
    for c in ("E", "F", "G", "H"):
        ws.Range(f"{c}{first}:{c}{last}").NumberFormat = D2
    ws.Range(f"I{first}:I{last}").NumberFormat = "0.0%"
    ws.Range(f"D{first}:D{last}").HorizontalAlignment = -4108
    ws.Range("D8:I8").HorizontalAlignment = -4152
    ws.Range("J8").Value = "Aug–Oct"
    th(ws.Range("J8"))
    fill(ws.Range("J8"), BG0)
    ws.Range("J8").HorizontalAlignment = -4108
    fill(ws.Range(f"J{first}:J{last}"), CARD)
    sg = ws.Range(f"J{first}:J{last}").SparklineGroups.Add(1, f"Payees!$E${first}:$G${last}")
    sg.SeriesColor.Color = rgb(*GOLD)
    sg.LineWeight = 1.5
    try:
        sg.Points.Markers.Visible = True
        sg.Points.Markers.Color.Color = rgb(*GOLD_LT)
    except Exception as e:
        print("sparkline", e)
    db = ws.Range(f"H{first}:H{last}").FormatConditions.AddDatabar()
    db.BarColor.Color = rgb(*COPPER)
    db.BarFillType = 0
    db.MinPoint.Modify(0, 0)
    fnt(ws.Range(f"C{first}:C{last}"), MUTED, 9)
    tr = last + 1
    ws.Range(f"B{tr}").Value = "Total spend"
    for c, nm in zip("DEFGHI", ["Payments"] + [lbl for _, lbl, _ in MONTHS] + ["Total", "Share"]):
        ws.Range(f"{c}{tr}").Formula = f"=SUM(Payees[{nm}])"
    ws.Range(f"E{tr}:H{tr}").NumberFormat = D2
    ws.Range(f"I{tr}").NumberFormat = "0%"
    ws.Range(f"D{tr}").HorizontalAlignment = -4108
    t = ws.Range(f"B{tr}:J{tr}")
    fill(t, NETROW); fnt(t, GOLD_LT, 10, True); line(t, 8, GOLD, 2)
    dark_intro(ws, f"B{TOP}", "Every cost line (bank and card) grouped by the payee its rule names. Biggest first. "
                              "Total spend = cost of goods + operating costs, less depreciation (no payee).")
    dark_intro(ws, f"B{tr + 2}", "\"Unknown (on Review)\" = lines no rule knows yet: once the owner explains them, add a rule and Refresh All.")
    premium(ws, "Payees", "Who gets paid", "Vendor and payee spend by month, from the same rules that choose the account.", "K", tr + 3)
    page(ws, f"$A$1:$K${tr + 3}")
    return tr


def build_review_sheet(wb):
    ws = wb.Worksheets("Review")
    page_bg(ws, "I")
    lo = ws.ListObjects("Review")
    lo.TableStyle = "Copperleaf Dark"
    lo.Range.Interior.ColorIndex = -4142
    fnt(lo.HeaderRowRange, GOLD, None, True)
    fnt(lo.DataBodyRange, IVORY)
    widths = {"A": 2, "B": 30, "C": 10, "D": 25, "E": 12, "F": 28, "G": 12, "H": 74, "I": 2}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    n = max(lo.ListRows.Count, 1)
    first, last = 9, 8 + n
    ws.Range(f"G{first}:G{last}").NumberFormat = D2
    ws.Range(f"E{first}:E{last}").NumberFormat = "dd mmm yyyy"
    fnt(ws.Range(f"B{first}:B{last}"), AMBER, 10, True)
    fnt(ws.Range(f"H{first}:H{last}"), MUTED, 9)
    ws.Range("G8").HorizontalAlignment = -4152
    dark_intro(ws, f"B{TOP}", "Only the lines that need a person: payments no rule knows, and transfer halves whose other side is not in the files yet.")
    end = last + 4
    dark_intro(ws, f"B{last + 2}", "Fix a line: ask the owner what it was, add a keyword (and payee) on the Rules tab, Data > Refresh All. "
                                   "The line moves to its real account and leaves this list.")
    premium(ws, "Review", "Review: lines that need a person", "Everything else was categorised by a rule and reconciled.", "I", end)
    page(ws, f"$A$1:$I${end}")


def format_tables(wb):
    def fmt(sheet, widths, money=(), dates=(), wrap=False, name=None):
        ws = wb.Worksheets(sheet)
        lo = ws.ListObjects(name) if name else ws.ListObjects(1)
        c0 = lo.Range.Column
        for i in range(1, lo.ListColumns.Count + 1):
            ws.Columns(c0 + i - 1).ColumnWidth = widths.get(lo.ListColumns(i).Name, 12)
        for c in money:
            lo.ListColumns(c).DataBodyRange.NumberFormat = MONEY
        for c in dates:
            lo.ListColumns(c).DataBodyRange.NumberFormat = "dd mmm yyyy"
        lo.HeaderRowRange.WrapText = wrap
        if lo.Range.Row == 1:
            ws.Activate(); ws.Range("A2").Select(); wb.Application.ActiveWindow.FreezePanes = True
        return lo

    fmt("Ledger", {"Posted": 12, "MonthNo": 8, "Source": 11, "Line": 9, "Description": 36, "Account": 8, "Account name": 28,
                   "Type": 9, "Debit": 12, "Credit": 12, "Net": 12}, money=("Debit", "Credit", "Net"), dates=("Posted",))
    fmt("Transactions", {"Line": 9, "Register": 24, "Register no": 8, "Posted": 12, "MonthNo": 8, "Bought": 12, "Description": 34,
                         "Clean description": 34, "Amount": 12, "Matched rule": 20, "Account": 8, "Account name": 28, "Payee": 24, "Type": 9},
        money=("Amount",), dates=("Posted", "Bought"), wrap=True)
    lo = fmt("Transfers", {"Amount": 12, "From": 26, "From line": 9, "From date": 12, "To": 26, "To line": 9, "To date": 12,
                           "Days apart": 8, "Status": 30}, money=("Amount",), dates=("From date", "To date"))
    rng = lo.ListColumns("Status").DataBodyRange
    fc = rng.FormatConditions.Add(2, None, '=$I2<>"Paired"'); fc.Interior.Color = rgb(255, 199, 206)
    fc = rng.FormatConditions.Add(2, None, '=$I2="Paired"'); fc.Interior.Color = rgb(226, 239, 218)
    fmt("Statements", {"Account (as printed)": 28, "Account no": 10, "Statement date": 14, "MonthNo": 9, "Ending balance": 15},
        money=("Ending balance",), dates=("Statement date",))
    fmt("Forecast", {"Order": 7, "Item": 40, "Cash flow line": 32, "Amount": 12, "Date": 12, "Week no": 8},
        money=("Amount",), dates=("Date",))
    ws = wb.Worksheets("Cash plan")
    lo = ws.ListObjects("CashPlanView")
    for c in ("Each time", "Total 13 weeks"):
        lo.ListColumns(c).DataBodyRange.NumberFormat = MONEY
    ws = wb.Worksheets("Cleaning log")
    for c, w in zip("ABCD", (18, 56, 7, 52)):
        ws.Columns(c).ColumnWidth = w


def build_howto(wb):
    ws = wb.Worksheets("How to use")
    ws.Columns("A").ColumnWidth = 125
    text = [
        ("How to use this workbook", True), ("", False),
        ("Every month", True),
        ("1. Download the checking, savings and card transactions as CSV into the data folder (see Settings) and type the file names there.", False),
        ("2. Copy each statement's closing balance into the statements file (or point Settings at a new one).", False),
        ("3. Data > Refresh All. Transactions, ledger, trial balance, P&L, balance sheet, cash flow, forecast, bank rec and payees all update. No macros.", False),
        ("4. Summary: pick the month and the comparison with the two slicers. Every data check must say PASS.", False),
        ("5. Month-end close tab: work down the ten steps. Clear the Review tab (add a keyword and payee on Rules for each new merchant, Refresh All again).", False),
        ("6. Type month-end entries that never touch the bank (depreciation, accruals) on Adjustments.", False), ("", False),
        ("Moving around", True),
        ("Every report sheet has the same menu at the top: Summary, P&L, Trial balance, Balance sheet, Cash flow, Bank rec, Review, Close, Payees.", False),
        ("The month slicer and the compare slicer (vs Budget / vs Last month) filter two small tables on the hidden Pickers sheet; formulas read the visible row.", False),
        ("", False),
        ("How a bank line becomes a ledger entry", True),
        ("The description is cleaned (upper case, punctuation removed) and the Rules tab is read top to bottom: the first keyword found decides the account and the payee.", False),
        ("Each line is posted twice: once to its own register (checking, savings or card) and once to the account the rule chose, so debits always equal credits.", False),
        ("Lines no rule knows go to 6990 Uncategorized and are listed on Review; they stay visible in the P&L until explained.", False),
        ("Transfers between your own accounts and card payments go to 1090 Transfers in transit from both sides. The Transfers tab pairs the two halves; "
         "what is left over is money still on its way at the month end.", False),
        ("Card lines are booked by POST date, so each month agrees with the card statement.", False), ("", False),
        ("Cash flow and forecast", True),
        ("Cash flow statement (indirect method): net profit, plus depreciation, plus or minus the change in money in transit and in the card balance, "
         "minus owner's draws = the change in checking + savings. It must tie to the balance sheet.", False),
        ("13-week forecast: the Cash plan tab lists the regular receipts and payments; each amount is a fixed number or an average from the bank history. "
         "Type your minimum cash buffer on the Cash flow sheet.", False), ("", False),
        ("What this workbook is not", True),
        ("A bookkeeping aid for a small business owner or their bookkeeper. It is not tax or accounting advice; your accountant still prepares the tax return and year-end accounts.", False),
        ("The Power Query steps are readable in Data > Queries & Connections: *_Clean, Transactions, Transfers, Ledger, Review, Payees, Cash_Plan_Calc, Forecast.", False),
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
    expand = xl.AutoCorrect.AutoExpandListRange
    xl.AutoCorrect.AutoExpandListRange = False   # cells typed next to a table must not join it
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < len(LOADED):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, (title, _, _) in enumerate(LOADED, 1):
            wb.Worksheets(i).Name = title
        st = wb.Styles("Normal")
        st.Font.Name = "Segoe UI"
        st.Font.Size = 10
        build_settings(wb)
        build_inputs(wb)
        pickers = build_pickers(wb)
        pickers[0].Range.AutoFilter(1, MONTHS[-1][1][:3])     # open on the last month ...
        pickers[1].Range.AutoFilter(1, "vs Budget")            # ... compared with the budget
        wb.SaveAs(str(OUT), 51)
        xl.Calculate()
        try:
            wb.Queries.FastCombine = True  # the queries combine the workbook's own tables with the CSV files
        except Exception as e:
            print("FastCombine not set:", e)
        for name, formula in Q.items():
            wb.Queries.Add(name, formula)
        t0 = time.time()
        for title, query, at in LOADED:
            if query:
                load_query(wb.Worksheets(title), query, query.replace("_", ""), at)
                print(" loaded", query, f"{time.time() - t0:.1f}s")
        build_styles(wb)
        format_tables(wb)
        pl = build_pl(wb)
        tb = build_tb(wb)
        bs = build_bs(wb)
        rec = build_rec(wb)
        cf = build_cashflow(wb, bs, pl)
        close = build_close(wb, tb, bs, rec, cf)
        calc = build_calc(wb, pl, bs, rec, tb, cf, close)
        build_summary(wb, calc, pickers)
        build_payees_sheet(wb)
        build_review_sheet(wb)
        build_howto(wb)
        xl.Calculate()
        dark = {"Summary", "P&L", "Trial balance", "Balance sheet", "Cash flow", "Bank rec", "Review", "Month-end close", "Payees"}
        inputs = {"Accounts", "Rules", "Budget", "Adjustments", "Cash plan"}
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            xl.ActiveWindow.DisplayHeadings = ws.Name not in dark
            if ws.Name in dark:
                ws.Tab.Color = rgb(*GOLD) if ws.Name == "Summary" else rgb(*BG0)
            elif ws.Name in inputs:
                ws.Tab.Color = rgb(255, 230, 120)
            ws.Range("A1").Select()
        wb.Worksheets("Pickers").Visible = 0      # xlSheetHidden: the slicers still work
        wb.Worksheets("Summary").Activate()
        wb.Save()
        print("saved", OUT.name)
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        xl.AutoCorrect.AutoExpandListRange = expand
        xl.Quit()


if __name__ == "__main__":
    main()
