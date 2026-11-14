"""Builds Bookkeeping-Trial-Balance.xlsx with real Excel (COM).

Power Query reads three bank/card exports, cleans them, gives every line an account from the Rules tab, pairs the two
halves of every transfer and card payment, and writes a general ledger (bank lines + adjusting entries). Formulas turn
the ledger into a trial balance per month, a P&L vs budget, a simple balance sheet and a bank reconciliation, all
checked on the Summary.

Run with a Python that has pywin32:  python scripts/build_workbook.py
The workbook has no macros; the client only needs Data -> Refresh All.
"""
import sys, time
from pathlib import Path
import win32com.client as win32

sys.path.insert(0, str(Path(__file__).parent))
from coa import ACCOUNTS, RULES, BUDGET, ADJUSTMENTS, MONTHS, CLEARING, UNCATEGORIZED

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
    Recs   = List.Buffer(List.Transform(Table.ToRecords(Kept), each [Keyword = Norm(Text.From([Keyword])), Account = Number.From([Account])]))
in
    Recs"""

Q["Transactions"] = f"""let
    All     = Table.Combine({{Checking_Clean, Savings_Clean, Card_Clean}}),
    Rules   = Rule_List,
    Clean   = Table.AddColumn(All, "Clean description", each Norm([Description]), type text),
    Hit     = Table.AddColumn(Clean, "Hit", each List.First(List.Select(Rules, (r) => Text.Contains([Clean description], r[Keyword])), null)),
    Acc     = Table.AddColumn(Hit, "Account", each if [Hit] = null then {UNCATEGORIZED} else [Hit][Account], Int64.Type),
    Rule    = Table.AddColumn(Acc, "Matched rule", each if [Hit] = null then "(no rule)" else [Hit][Keyword], type text),
    Month   = Table.AddColumn(Rule, "MonthNo", each Date.Year([Posted]) * 100 + Date.Month([Posted]), Int64.Type),
    Named   = Table.NestedJoin(Month, {{"Account"}}, Accounts_T, {{"Account"}}, "A", JoinKind.LeftOuter),
    Exp     = Table.ExpandTableColumn(Named, "A", {{"Name"}}, {{"Account name"}}),
    Sorted  = Table.Sort(Exp, {{{{"Posted", Order.Ascending}}, {{"Line", Order.Ascending}}}}),
    Result  = Table.SelectColumns(Sorted, {{"Line", "Register", "Register no", "Posted", "MonthNo", "Bought", "Description",
                  "Clean description", "Amount", "Matched rule", "Account", "Account name"}})
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

LOADED = [("Summary", None), ("P&L", None), ("Trial balance", None), ("Balance sheet", None), ("Bank rec", None),
          ("Review", "Review"), ("Ledger", "Ledger"), ("Transactions", "Transactions"), ("Transfers", "Transfers"),
          ("Statements", "Statements"), ("Accounts", None), ("Rules", None), ("Budget", None), ("Adjustments", None),
          ("Cleaning log", "Cleaning_Log"), ("Settings", None), ("How to use", None)]
NAVY, TEAL = (31, 58, 90), (0, 128, 128)
GREEN = ((198, 239, 206), (0, 97, 0))
RED = ((255, 199, 206), (156, 0, 6))
YELLOW = (255, 255, 153)
LIGHT = (242, 245, 249)
MONEY = '#,##0.00;[Red]-#,##0.00'
MONEY0 = '#,##0;[Red]-#,##0;"–"'
TYPES = {a[0]: a[2] for a in ACCOUNTS}
NAMES = {a[0]: a[1] for a in ACCOUNTS}


def rgb(r, g, b):
    return r + g * 256 + b * 65536


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


def heading(ws, cell, text, size=13):
    ws.Range(cell).Value = text
    ws.Range(cell).Font.Size = size
    ws.Range(cell).Font.Bold = True
    ws.Range(cell).Font.Color = rgb(*NAVY)


def header_row(rng):
    rng.Font.Bold = True
    rng.Interior.Color = rgb(*NAVY)
    rng.Font.Color = rgb(255, 255, 255)


def pass_fail(rng, ok='="PASS"'):
    fc = rng.FormatConditions.Add(1, 3, ok)
    fc.Interior.Color = rgb(*GREEN[0]); fc.Font.Color = rgb(*GREEN[1])
    fc = rng.FormatConditions.Add(1, 4, ok)
    fc.Interior.Color = rgb(*RED[0]); fc.Font.Color = rgb(*RED[1])


def better_worse(rng):
    fc = rng.FormatConditions.Add(1, 5, "=0")   # > : better than budget (condition formulas use the local list/decimal separator: keep them comma-free)
    fc.Font.Color = rgb(0, 128, 0)
    fc = rng.FormatConditions.Add(1, 6, "=0")  # < : worse
    fc.Font.Color = rgb(192, 0, 0)


def title_band(ws, last_col, title_formula, sub):
    ws.Range(f"A1:{last_col}2").Interior.Color = rgb(*NAVY)
    ws.Range("B1").Formula = title_formula
    ws.Range("B1").Font.Size = 18; ws.Range("B1").Font.Bold = True; ws.Range("B1").Font.Color = rgb(255, 255, 255)
    ws.Range("B2").Value = sub
    ws.Range("B2").Font.Color = rgb(200, 215, 230)
    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 20


def page(ws, area, landscape=True):
    ps = ws.PageSetup
    ps.Orientation = 2 if landscape else 1
    ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = area


def mno(k):
    return f"INDEX(MonthNos,{k})"


def bal(acct_ref, k):
    """Balance of an account at the end of month k (debit +): opening + every ledger line up to that month."""
    return (f"(SUMIFS(Accounts[Opening debit],Accounts[Account],{acct_ref})-SUMIFS(Accounts[Opening credit],Accounts[Account],{acct_ref})"
            f"+SUMIFS(Ledger[Net],Ledger[Account],{acct_ref},Ledger[MonthNo],\"<=\"&{mno(k)}))")


def put_table(ws, at_row, headers, rows, name, text_cols=()):
    n = len(headers)
    last = chr(ord("A") + n - 1)
    for c in text_cols:
        ws.Range(f"{c}{at_row}:{c}{at_row + len(rows)}").NumberFormat = "@"
    ws.Range(f"A{at_row}:{last}{at_row}").NumberFormat = "@"
    ws.Range(f"A{at_row}:{last}{at_row}").Value = tuple(headers)
    if rows:
        ws.Range(f"A{at_row + 1}:{last}{at_row + len(rows)}").Value = tuple(tuple(r) for r in rows)
    lo = ws.ListObjects.Add(1, ws.Range(f"A{at_row}:{last}{at_row + max(1, len(rows))}"), None, 1)
    lo.Name = name
    lo.TableStyle = "TableStyleMedium2"
    return lo


def intro(ws, row, text):
    ws.Range(f"A{row}").Value = text
    ws.Range(f"A{row}").Font.Italic = True
    ws.Range(f"A{row}").Font.Color = rgb(90, 100, 115)


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
    intro(ws, 2, "If the keyword appears in a cleaned description (upper case, punctuation removed), the line goes to that account. "
                 "Top to bottom: the first match wins, so put specific keywords above general ones. Add a row, then Refresh All.")
    lo = put_table(ws, 4, ["Keyword", "Account", "Account name", "Why"], [(k, a, None, w) for k, a, w in RULES], "Rules")
    lo.ListColumns("Account name").DataBodyRange.Formula = '=IFERROR(INDEX(Accounts[Name],MATCH([@Account],Accounts[Account],0)),"unknown account")'
    lo.ListColumns("Keyword").DataBodyRange.Interior.Color = rgb(*YELLOW)
    lo.ListColumns("Account").DataBodyRange.Interior.Color = rgb(*YELLOW)
    for c, w in zip("ABCD", (24, 10, 30, 62)):
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
        lo.ListColumns("Date").DataBodyRange.Cells(i, 1).Formula = f"=DATE({d.year},{d.month},{d.day})"
        lo.ListColumns("Date").DataBodyRange.Cells(i, 1).Value = lo.ListColumns("Date").DataBodyRange.Cells(i, 1).Value
    lo.ListColumns("Debit").DataBodyRange.NumberFormat = MONEY
    lo.ListColumns("Credit").DataBodyRange.NumberFormat = MONEY
    lo.DataBodyRange.Interior.Color = rgb(*YELLOW)
    for c, w in zip("ABCDEF", (12, 9, 58, 10, 12, 12)):
        ws.Columns(c).ColumnWidth = w


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
            ("Period", "August to October 2026", "Period", "Shown in the report titles")]
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
    st.Range("A13:C13").Value = ("Month", "MonthNo", "Month end")
    st.Range("A13:C13").Font.Bold = True
    for i, (no, lbl, end) in enumerate(MONTHS, 14):
        st.Cells(i, 1).NumberFormat = "@"
        st.Cells(i, 1).Value = lbl
        st.Cells(i, 2).Value = no
        st.Cells(i, 3).Formula = f"=DATE({end.year},{end.month},{end.day})"
        st.Cells(i, 3).NumberFormat = "dd mmm yyyy"
    wb.Names.Add("MonthList", "=Settings!$A$14:$A$16")
    wb.Names.Add("MonthNos", "=Settings!$B$14:$B$16")
    wb.Names.Add("MonthEnds", "=Settings!$C$14:$C$16")
    st.Range("A12").Value = "Months in this workbook (MonthNo = year x 100 + month, used by every report)"
    st.Range("A12").Font.Italic = True
    st.Columns("A").ColumnWidth = 22; st.Columns("B").ColumnWidth = 48; st.Columns("C").ColumnWidth = 90


# ---------------------------------------------------------------- P&L
def build_pl(wb):
    ws = wb.Worksheets("P&L")
    title_band(ws, "O", '="Profit and loss vs budget   |   "&Company&"   |   "&Period',
               "From the general ledger. Variance: + = better than budget (more revenue or lower cost), - = worse.")
    widths = {"A": 2, "B": 7, "C": 30}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    for c in "DEFGHIJKLMNO":
        ws.Columns(c).ColumnWidth = 10.5
    for k, (no, lbl, _) in enumerate(MONTHS + [(0, "Quarter", None)]):
        c1 = chr(ord("D") + 3 * k)
        c3 = chr(ord("D") + 3 * k + 2)
        rng = ws.Range(f"{c1}4:{c3}4")
        rng.Merge(); rng.NumberFormat = "@"; rng.Value = lbl; rng.HorizontalAlignment = -4108
        header_row(rng)
        ws.Range(f"{c1}5:{c3}5").Value = ("Actual", "Budget", "Variance")
        ws.Range(f"{c1}5:{c3}5").Font.Bold = True
        ws.Range(f"{c1}5:{c3}5").HorizontalAlignment = -4152
        ws.Range(f"{c1}5:{c3}5").Borders(9).LineStyle = 1
    ws.Range("B5").Value = "Acct"; ws.Range("C5").Value = "Account"
    ws.Range("B5:C5").Font.Bold = True; ws.Range("B5:C5").Borders(9).LineStyle = 1

    rows = {}
    r = 6
    sections = [("Revenue", ["Revenue"], -1, "Total revenue"), ("Cost of goods sold", ["COGS"], 1, "Total cost of goods sold"),
                ("Operating expenses", ["Expense"], 1, "Total operating expenses")]

    def block_cols(k):
        return [chr(ord("D") + 3 * k + i) for i in range(3)]

    for title, types, sign, total in sections:
        ws.Range(f"B{r}").Value = title
        ws.Range(f"B{r}").Font.Bold = True; ws.Range(f"B{r}").Font.Color = rgb(*NAVY)
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
            r += 1
        last = r - 1
        ws.Range(f"C{r}").Value = total
        for c in "DEFGHIJKLMNO":
            ws.Range(f"{c}{r}").Formula = f"=SUM({c}{first}:{c}{last})"
        ws.Range(f"B{r}:O{r}").Font.Bold = True
        ws.Range(f"B{r}:O{r}").Borders(8).LineStyle = 1
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
            ws.Range(f"B{r}:O{r}").Interior.Color = rgb(*LIGHT)
            r += 1
            rows["Gross margin"] = r
            ws.Range(f"C{r}").Value = "Gross margin"
            for k in range(4):
                a, b, v = block_cols(k)
                ws.Range(f"{a}{r}").Formula = f"=IFERROR({a}{r - 1}/{a}{rows['Revenue']},0)"
                ws.Range(f"{b}{r}").Formula = f"=IFERROR({b}{r - 1}/{b}{rows['Revenue']},0)"
                ws.Range(f"{a}{r}:{b}{r}").NumberFormat = "0.0%"
            ws.Range(f"C{r}:O{r}").Font.Italic = True
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
    ws.Range(f"B{r}:O{r}").Font.Bold = True; ws.Range(f"B{r}:O{r}").Font.Size = 12
    ws.Range(f"B{r}:O{r}").Interior.Color = rgb(221, 235, 247)
    ws.Range(f"B{r}:O{r}").Borders(8).LineStyle = 1; ws.Range(f"B{r}:O{r}").Borders(9).LineStyle = -4119
    r += 1
    rows["Net margin"] = r
    ws.Range(f"C{r}").Value = "Net margin"
    for k in range(4):
        a, b, v = block_cols(k)
        ws.Range(f"{a}{r}").Formula = f"=IFERROR({a}{r - 1}/{a}{rows['Revenue']},0)"
        ws.Range(f"{b}{r}").Formula = f"=IFERROR({b}{r - 1}/{b}{rows['Revenue']},0)"
        ws.Range(f"{a}{r}:{b}{r}").NumberFormat = "0.0%"
    ws.Range(f"C{r}:O{r}").Font.Italic = True
    last_row = r
    # number formats, variance colours
    for rr in range(7, last_row + 1):
        if rr in (rows["Gross margin"], rows["Net margin"]):
            continue
        ws.Range(f"D{rr}:O{rr}").NumberFormat = MONEY0
    for k in range(4):
        v = block_cols(k)[2]
        better_worse(ws.Range(f"{v}7:{v}{rows['Net profit']}"))
        ws.Range(f"{v}6:{v}{last_row}").Borders(10).LineStyle = 1
        ws.Range(f"{v}6:{v}{last_row}").Borders(10).Color = rgb(190, 200, 212)
    ws.Range(f"M4:O{last_row}").Interior.Color = rgb(*LIGHT)
    ws.Range("M4:O4").Interior.Color = rgb(*TEAL)
    ws.Range(f"B{last_row + 2}").Value = ("Amounts from the Ledger tab (bank and card lines given an account by the Rules tab, plus the adjusting entries). "
                                          "Budget from the Budget tab. Uncategorized lines stay visible here until they are explained.")
    ws.Range(f"B{last_row + 2}").Font.Italic = True; ws.Range(f"B{last_row + 2}").Font.Color = rgb(90, 100, 115)
    for cell in ws.Range(f"D6:O{last_row}"):  # subtotal rows differ from their neighbours on purpose: no green triangles
        try:
            cell.Errors(4).Ignore = True  # xlInconsistentFormula
        except Exception:
            pass
    ws.Activate(); ws.Range("D6").Select(); wb.Application.ActiveWindow.FreezePanes = True
    page(ws, f"$A$1:$O${last_row + 2}")
    return rows


# ---------------------------------------------------------------- trial balance
def build_tb(wb):
    ws = wb.Worksheets("Trial balance")
    title_band(ws, "L", '="Trial balance at each month end   |   "&Company',
               "Opening balances + every ledger line up to the month end. Debits must equal credits.")
    for c, w in {"A": 2, "B": 8, "C": 30, "D": 10}.items():
        ws.Columns(c).ColumnWidth = w
    for c in "EFGHIJKL":
        ws.Columns(c).ColumnWidth = 12
    blocks = [("E", "F", "Opening 31 Jul 2026", 0)] + [(chr(ord("G") + 2 * i), chr(ord("H") + 2 * i), lbl, i + 1)
                                                       for i, (_, lbl, _) in enumerate(MONTHS)]
    ws.Range("B5:D5").Value = ("Acct", "Account", "Type")
    ws.Range("B5:D5").Font.Bold = True; ws.Range("B5:D5").Borders(9).LineStyle = 1
    for a, b, lbl, k in blocks:
        rng = ws.Range(f"{a}4:{b}4"); rng.Merge(); rng.Value = lbl; rng.HorizontalAlignment = -4108; header_row(rng)
        if k:
            ws.Range(f"{a}4").Formula = f'="{lbl}  (to "&TEXT(INDEX(MonthEnds,{k}),"d mmm")&")"'
        ws.Range(f"{a}5:{b}5").Value = ("Debit", "Credit")
        ws.Range(f"{a}5:{b}5").Font.Bold = True; ws.Range(f"{a}5:{b}5").HorizontalAlignment = -4152
        ws.Range(f"{a}5:{b}5").Borders(9).LineStyle = 1
    r = 6
    for acct, *_ in ACCOUNTS:
        ws.Range(f"B{r}").Value = acct
        ws.Range(f"C{r}").Formula = f'=IFERROR(INDEX(Accounts[Name],MATCH(B{r},Accounts[Account],0)),"?")'
        ws.Range(f"D{r}").Formula = f'=IFERROR(INDEX(Accounts[Type],MATCH(B{r},Accounts[Account],0)),"?")'
        op = f"(SUMIFS(Accounts[Opening debit],Accounts[Account],$B{r})-SUMIFS(Accounts[Opening credit],Accounts[Account],$B{r}))"
        ws.Range(f"E{r}").Formula = f"=MAX({op},0)"
        ws.Range(f"F{r}").Formula = f"=MAX(-{op},0)"
        for a, b, lbl, k in blocks[1:]:
            bb = bal(f"$B{r}", k)
            ws.Range(f"{a}{r}").Formula = f"=MAX({bb},0)"
            ws.Range(f"{b}{r}").Formula = f"=MAX(-{bb},0)"
        r += 1
    last = r - 1
    tot, diff, chk = r, r + 1, r + 2
    ws.Range(f"C{tot}").Value = "Total"
    ws.Range(f"C{diff}").Value = "Debits minus credits"
    ws.Range(f"C{chk}").Value = "Balanced?"
    for a, b, lbl, k in blocks:
        ws.Range(f"{a}{tot}").Formula = f"=SUM({a}6:{a}{last})"
        ws.Range(f"{b}{tot}").Formula = f"=SUM({b}6:{b}{last})"
        ws.Range(f"{a}{diff}").Formula = f"=ROUND({a}{tot}-{b}{tot},2)"
        ws.Range(f"{a}{chk}").Formula = f'=IF({a}{diff}=0,"PASS","FAIL")'
        ws.Range(f"{a}{chk}").HorizontalAlignment = -4152
        pass_fail(ws.Range(f"{a}{chk}"))
    ws.Range(f"B{tot}:L{tot}").Font.Bold = True
    ws.Range(f"B{tot}:L{tot}").Borders(8).LineStyle = 1; ws.Range(f"B{tot}:L{tot}").Borders(9).LineStyle = -4119
    ws.Range(f"C{diff}:L{chk}").Font.Bold = True
    ws.Range(f"E6:L{diff}").NumberFormat = '#,##0.00;[Red]-#,##0.00;"–"'
    ws.Range(f"{'E'}{diff}:L{diff}").NumberFormat = "#,##0.00"
    for i in range(6, last + 1, 2):
        ws.Range(f"B{i}:L{i}").Interior.Color = rgb(*LIGHT)
    ws.Range(f"B{chk + 2}").Value = ("Profit and loss accounts start at zero on 1 August (the year so far); last year's profit is inside Owner's equity. "
                                    "Transfers in transit = money that left one account and reached the other after the month end.")
    ws.Range(f"B{chk + 2}").Font.Italic = True; ws.Range(f"B{chk + 2}").Font.Color = rgb(90, 100, 115)
    ws.Activate(); ws.Range("E6").Select(); wb.Application.ActiveWindow.FreezePanes = True
    page(ws, f"$A$1:$L${chk + 2}")
    return {"diff": diff, "check": chk, "total": tot}


# ---------------------------------------------------------------- balance sheet
def build_bs(wb):
    ws = wb.Worksheets("Balance sheet")
    title_band(ws, "G", '="Balance sheet at each month end   |   "&Company', "Simple balance sheet from the same ledger.")
    for c, w in {"A": 2, "B": 8, "C": 34, "D": 13, "E": 13, "F": 13, "G": 2}.items():
        ws.Columns(c).ColumnWidth = w
    for k, (_, lbl, _) in enumerate(MONTHS, 1):
        c = chr(ord("C") + k)
        ws.Range(f"{c}4").Formula = f'=TEXT(INDEX(MonthEnds,{k}),"d mmm yyyy")'
        header_row(ws.Range(f"{c}4")); ws.Range(f"{c}4").HorizontalAlignment = -4152
    header_row(ws.Range("B4:C4")); ws.Range("B4").Value = "Acct"; ws.Range("C4").Value = ""
    rows = {}
    r = 5

    def line(label, acct=None, formula=None, sign=1, bold=False, key=None):
        nonlocal r
        if acct:
            ws.Range(f"B{r}").Value = acct
        ws.Range(f"C{r}").Value = label
        for k in range(1, 4):
            c = chr(ord("C") + k)
            if formula:
                ws.Range(f"{c}{r}").Formula = "=" + formula(c, k)
            elif acct:
                ws.Range(f"{c}{r}").Formula = ("=-" if sign < 0 else "=") + bal(f"$B{r}", k)
        if bold:
            ws.Range(f"B{r}:F{r}").Font.Bold = True
            ws.Range(f"B{r}:F{r}").Borders(8).LineStyle = 1
        if key:
            rows[key] = r
        r += 1

    def section(t):
        nonlocal r
        ws.Range(f"B{r}").Value = t; ws.Range(f"B{r}").Font.Bold = True; ws.Range(f"B{r}").Font.Color = rgb(*NAVY)
        r += 1

    section("Assets")
    a0 = r
    for acct in (1000, 1010, CLEARING, 1500, 1510):
        line(NAMES[acct], acct, key=acct)
    line("Total assets", formula=lambda c, k: f"SUM({c}{a0}:{c}{r - 1})", bold=True, key="assets")
    r += 1
    section("Liabilities")
    line(NAMES[2000], 2000, sign=-1, key=2000)
    line("Total liabilities", formula=lambda c, k: f"{c}{rows[2000]}", bold=True, key="liab")
    r += 1
    section("Equity")
    e0 = r
    line(NAMES[3000], 3000, sign=-1, key=3000)
    line(NAMES[3100], 3100, sign=-1, key=3100)
    line("Profit for the year so far (from 1 Aug)", formula=lambda c, k: "-(" + "+".join(
        f'SUMIFS(Ledger[Net],Ledger[Type],"{t}",Ledger[MonthNo],"<="&{mno(k)})' for t in ("Revenue", "COGS", "Expense")) + ")", key="profit")
    line("Total equity", formula=lambda c, k: f"SUM({c}{e0}:{c}{r - 1})", bold=True, key="equity")
    r += 1
    line("Total liabilities and equity", formula=lambda c, k: f"{c}{rows['liab']}+{c}{rows['equity']}", bold=True, key="le")
    line("Assets minus (liabilities + equity)", formula=lambda c, k: f"ROUND({c}{rows['assets']}-{c}{rows['le']},2)", key="diff")
    line("Balanced?", formula=lambda c, k: f'IF({c}{rows["diff"]}=0,"PASS","FAIL")', key="check")
    ws.Range(f"D5:F{rows['diff']}").NumberFormat = '#,##0.00;[Red]-#,##0.00;"–"'
    ws.Range(f"D{rows['check']}:F{rows['check']}").HorizontalAlignment = -4152
    ws.Range(f"D{rows['diff']}:F{rows['diff']}").NumberFormat = "0.00"
    ws.Range(f"C{rows['diff']}:F{rows['check']}").Font.Bold = True
    pass_fail(ws.Range(f"D{rows['check']}:F{rows['check']}"))
    ws.Range(f"B{rows['le']}:F{rows['le']}").Interior.Color = rgb(*LIGHT)
    ws.Range(f"B{rows['assets']}:F{rows['assets']}").Interior.Color = rgb(*LIGHT)
    page(ws, f"$A$1:$G${rows['check']}", landscape=False)
    return rows


# ---------------------------------------------------------------- bank rec
def build_rec(wb):
    ws = wb.Worksheets("Bank rec")
    title_band(ws, "I", '="Books vs bank statements   |   "&Company',
               "The balance in the books must equal the statement, every account, every month. Card: amount owed, as the statement prints it.")
    for c, w in {"A": 2, "B": 30, "C": 12, "D": 10, "E": 16, "F": 16, "G": 13, "H": 10, "I": 2}.items():
        ws.Columns(c).ColumnWidth = w
    ws.Range("B4:H4").Value = ("Account", "Month end", "Lines", "Balance in books", "Statement balance", "Difference", "Status")
    header_row(ws.Range("B4:H4"))
    ws.Range("C4:H4").HorizontalAlignment = -4152
    r = 5
    for acct in (1000, 1010, 2000):
        for k, (no, lbl, end) in enumerate(MONTHS, 1):
            ws.Range(f"B{r}").Value = NAMES[acct] if k == 1 else ""
            ws.Range(f"J{r}").Value = acct
            ws.Range(f"C{r}").Formula = f"=INDEX(MonthEnds,{k})"
            ws.Range(f"C{r}").NumberFormat = "dd mmm yyyy"
            ws.Range(f"D{r}").Formula = f"=COUNTIFS(Transactions[Register no],$J{r},Transactions[MonthNo],{mno(k)})"
            ws.Range(f"E{r}").Formula = ("=-" if acct == 2000 else "=") + bal(f"$J{r}", k)
            ws.Range(f"F{r}").Formula = f"=SUMIFS(Statements[Ending balance],Statements[Account no],$J{r},Statements[MonthNo],{mno(k)})"
            ws.Range(f"G{r}").Formula = f"=ROUND(E{r}-F{r},2)"
            ws.Range(f"H{r}").Formula = f'=IF(AND(G{r}=0,COUNTIFS(Statements[Account no],$J{r},Statements[MonthNo],{mno(k)})=1),"PASS","FAIL")'
            if k == 3:
                ws.Range(f"B{r}:H{r}").Borders(9).LineStyle = 1
                ws.Range(f"B{r}:H{r}").Borders(9).Color = rgb(190, 200, 212)
            r += 1
    ws.Columns("J").Hidden = True
    ws.Range(f"E5:G{r - 1}").NumberFormat = '#,##0.00;[Red]-#,##0.00;"0.00"'
    ws.Range(f"H5:H{r - 1}").HorizontalAlignment = -4108
    ws.Range(f"H5:H{r - 1}").Font.Bold = True
    pass_fail(ws.Range(f"H5:H{r - 1}"))
    for i, t in enumerate(("Lines are counted by POST date, the way the bank and the card company build their statements.",
                           "A transfer still on its way shows up under Transfers in transit (balance sheet), not here."), 1):
        ws.Range(f"B{r + i}").Value = t
        ws.Range(f"B{r + i}").Font.Italic = True; ws.Range(f"B{r + i}").Font.Color = rgb(90, 100, 115)
    page(ws, f"$A$1:$I${r + 2}")
    return (5, r - 1)


# ---------------------------------------------------------------- summary
def build_summary(wb, pl, tb, bs, rec):
    ws = wb.Worksheets("Summary")
    widths = {"A": 2, "B": 30, "C": 12, "D": 12, "E": 12, "F": 12, "G": 3, "H": 12, "I": 12, "J": 12, "K": 12, "L": 12, "M": 12}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    title_band(ws, "M", '="Bookkeeping summary   |   "&Company&"   |   "&Period',
               "Bank and card exports categorised by rules, posted to a ledger, checked against every statement. Fictional data.")
    ws.Range("B3").Value = "Month (pick one):"
    ws.Range("B3").HorizontalAlignment = -4152; ws.Range("B3").Font.Bold = True
    sel = ws.Range("C3")
    sel.NumberFormat = "@"
    sel.Value = MONTHS[-1][1]
    sel.Interior.Color = rgb(*YELLOW); sel.Font.Bold = True; sel.HorizontalAlignment = -4108
    sel.Validation.Add(3, 1, 1, "=MonthList")
    ws.Rows(3).RowHeight = 22
    wb.Names.Add("SelMonth", "=Summary!$C$3")
    st = wb.Worksheets("Settings")  # a cell, not a name formula: Names.Add parses with the local separator
    st.Range("A18").Value = "Selected month (from Summary)"
    st.Range("B18").Formula = "=MATCH(Summary!$C$3,MonthList,0)"
    wb.Names.Add("SelIdx", "=Settings!$B$18")

    def plv(row, col):  # col 1 actual, 2 budget, 3 variance for the selected month
        return f"INDEX('P&L'!$D${row}:$L${row},1,3*SelIdx-{3 - col})"

    def bsv(row):
        return f"INDEX('Balance sheet'!$D${row}:$F${row},1,SelIdx)"

    # KPI cards rows 5-10
    cards = [("B", "C", "REVENUE", "=" + plv(pl["Revenue"], 1), "#,##0",
              f'="budget "&TEXT({plv(pl["Revenue"], 2)},"#,##0")&"   ("&TEXT({plv(pl["Revenue"], 3)},"+#,##0;-#,##0")&")"'),
             ("D", "F", "NET PROFIT", "=" + plv(pl["Net profit"], 1), "#,##0;[Red]-#,##0",
              f'="budget "&TEXT({plv(pl["Net profit"], 2)},"#,##0")&"   ("&TEXT({plv(pl["Net profit"], 3)},"+#,##0;-#,##0")&")"'),
             ("H", "J", "NET MARGIN", "=" + plv(pl["Net margin"], 1), "0.0%;[Red]-0.0%",
              f'="budget "&TEXT({plv(pl["Net margin"], 2)},"0.0%")'),
             ("K", "M", "CASH AT MONTH END", f"={bsv(bs[1000])}+{bsv(bs[1010])}", "#,##0",
              f'="checking + savings; "&TEXT({bsv(bs[CLEARING])},"#,##0")&" in transit"')]
    for a, b, label, formula, fmt, note in cards:
        box = ws.Range(f"{a}5:{b}10")
        box.Interior.Color = rgb(247, 249, 252)
        for edge in (7, 8, 9, 10):
            box.Borders(edge).LineStyle = 1
            box.Borders(edge).Color = rgb(210, 218, 228)
        ws.Range(f"{a}5:{b}5").Interior.Color = rgb(*TEAL)
        for rows_, val, size, bold, color in (("6:6", label, 9, True, (90, 100, 115)), ("7:8", formula, 24, True, NAVY),
                                              ("9:9", note, 9, False, (110, 110, 110))):
            r1, r2 = rows_.split(":")
            rng = ws.Range(f"{a}{r1}:{b}{r2}")
            if rng.Cells.Count > 1:
                rng.Merge()
            cell = ws.Range(f"{a}{r1}")
            if val.startswith("="):
                cell.Formula = val
            else:
                cell.Value = val
            cell.Font.Size = size; cell.Font.Bold = bold; cell.Font.Color = rgb(*color)
            rng.HorizontalAlignment = -4108; rng.VerticalAlignment = -4108
        ws.Range(f"{a}7").NumberFormat = fmt
    ws.Rows(5).RowHeight = 5

    # P&L for the selected month (rows 12-19)
    ws.Range("B12").Formula = '="Profit and loss, "&SelMonth'
    ws.Range("B12").Font.Size = 13; ws.Range("B12").Font.Bold = True; ws.Range("B12").Font.Color = rgb(*NAVY)
    ws.Range("B13:E13").Value = ("", "Actual", "Budget", "Variance")
    header_row(ws.Range("B13:E13")); ws.Range("C13:E13").HorizontalAlignment = -4152
    mini = [("Revenue", pl["Revenue"]), ("Cost of goods sold", pl["Cost of goods sold"]), ("Gross profit", pl["Gross profit"]),
            ("Operating expenses", pl["Operating expenses"]), ("Net profit", pl["Net profit"])]
    for i, (label, row) in enumerate(mini, 14):
        ws.Range(f"B{i}").Value = label
        for j, c in enumerate("CDE", 1):
            ws.Range(f"{c}{i}").Formula = "=" + plv(row, j)
        if label in ("Gross profit", "Net profit"):
            ws.Range(f"B{i}:E{i}").Font.Bold = True
            ws.Range(f"B{i}:E{i}").Borders(8).LineStyle = 1
    ws.Range("C14:E18").NumberFormat = "#,##0;[Red]-#,##0"
    better_worse(ws.Range("E14:E18"))
    ws.Range("B19").Formula = ('="Variance: + = better than budget.   Lines to review: "&COUNTA(Review[Line])&" (Review tab)"')
    ws.Range("B19").Font.Italic = True; ws.Range("B19").Font.Color = rgb(90, 100, 115)

    # three months at a glance (rows 21-25) = chart source
    ws.Range("B21").Value = "Three months at a glance"
    ws.Range("B21").Font.Size = 13; ws.Range("B21").Font.Bold = True; ws.Range("B21").Font.Color = rgb(*NAVY)
    ws.Range("B22:F22").Value = ("Month", "Revenue", "Total costs", "Net profit", "Cash")
    header_row(ws.Range("B22:F22")); ws.Range("C22:F22").HorizontalAlignment = -4152
    for k in range(3):
        r = 23 + k
        a = chr(ord("D") + 3 * k)
        ws.Range(f"B{r}").Formula = f"=INDEX(MonthList,{k + 1})"
        ws.Range(f"C{r}").Formula = f"='P&L'!{a}{pl['Revenue']}"
        ws.Range(f"D{r}").Formula = f"='P&L'!{a}{pl['Cost of goods sold']}+'P&L'!{a}{pl['Operating expenses']}"
        ws.Range(f"E{r}").Formula = f"='P&L'!{a}{pl['Net profit']}"
        bc = chr(ord("D") + k)
        ws.Range(f"F{r}").Formula = f"='Balance sheet'!{bc}{bs[1000]}+'Balance sheet'!{bc}{bs[1010]}"
    ws.Range("C23:F25").NumberFormat = "#,##0;[Red]-#,##0"
    fc = ws.Range("B23:F25").FormatConditions.Add(2, None, "=$B23=SelMonth")
    fc.Interior.Color = rgb(255, 242, 204)

    # chart H12:M25
    co = ws.ChartObjects().Add(ws.Range("H12").Left, ws.Range("H12").Top, ws.Range("H12:M12").Width, ws.Range("H12:H25").Height)
    ch = co.Chart
    ch.ChartType = 51
    for col, name, color in (("C", "Revenue", TEAL), ("D", "Total costs", (160, 174, 192)), ("E", "Net profit", NAVY)):
        s = ch.SeriesCollection().NewSeries()
        s.Name = name
        s.Values = ws.Range(f"{col}23:{col}25")
        s.XValues = ws.Range("B23:B25")
        s.Format.Fill.ForeColor.RGB = rgb(*color)
    ch.SeriesCollection(3).ChartType = 65
    ch.SeriesCollection(3).Format.Line.ForeColor.RGB = rgb(*NAVY)
    ch.SeriesCollection(3).MarkerBackgroundColor = rgb(*NAVY)
    ch.SeriesCollection(3).MarkerForegroundColor = rgb(*NAVY)
    ch.SeriesCollection(3).HasDataLabels = True
    ch.SeriesCollection(3).DataLabels().NumberFormat = "#,##0"
    ch.SeriesCollection(3).DataLabels().Position = 0  # above
    ch.HasLegend = True; ch.Legend.Position = -4107  # bottom
    ch.Axes(2).HasMajorGridlines = True
    ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = rgb(225, 230, 236)
    ch.Axes(2).TickLabels.NumberFormat = "#,##0"
    ch.HasTitle = True
    ch.ChartTitle.Text = "Revenue, costs and net profit by month"
    ch.ChartTitle.Font.Size = 13; ch.ChartTitle.Font.Bold = True; ch.ChartTitle.Font.Color = rgb(*NAVY)
    ch.ChartArea.Format.Line.Visible = False

    # data checks rows 27-35
    heading(ws, "B27", "Data checks")
    log = "INDEX(CleaningLog[Rows],{})".format
    r1, r2 = rec
    checks = [
        ("Opening balances balance (debits = credits at 31 Jul)",
         '=IF(ROUND(SUM(Accounts[Opening debit])-SUM(Accounts[Opening credit]),2)=0,"PASS","FAIL")'),
        ("Trial balance balances at every month end",
         f"=IF(COUNTIF('Trial balance'!$G${tb['check']}:$L${tb['check']},\"PASS\")=3,\"PASS\",\"FAIL\")"),
        ("Every bank and card line posted once (lines in the files = ledger)",
         f'=IF(AND(ROWS(Transactions[Line])={log(1)}+{log(4)}+{log(6)},COUNTIFS(Ledger[Source],"Bank line")=2*ROWS(Transactions[Line])),"PASS","FAIL")'),
        ("Every line has a known account; lines no rule knows are all on Review",
         f'=IF(AND(SUMPRODUCT(--ISNA(MATCH(Transactions[Account],Accounts[Account],0)))=0,'
         f'COUNTIFS(Transactions[Account],{UNCATEGORIZED})=COUNTIFS(Review[Issue],"No rule matched")),"PASS","FAIL")'),
        ("Transfers: what is still in transit = the unpaired halves on Review",
         f"=IF(ROUND('Balance sheet'!$F${bs[CLEARING]}+SUMIFS(Review[Amount],Review[Issue],\"Transfer without its other half\"),2)=0,\"PASS\",\"FAIL\")"),
        ("Books = bank and card statements (3 accounts x 3 months)",
         f"=IF(COUNTIF('Bank rec'!$H${r1}:$H${r2},\"PASS\")=9,\"PASS\",\"FAIL\")"),
        ("Balance sheet balances at every month end",
         f"=IF(COUNTIF('Balance sheet'!$D${bs['check']}:$F${bs['check']},\"PASS\")=3,\"PASS\",\"FAIL\")"),
        ("Adjusting entries balance (debits = credits)",
         '=IF(ROUND(SUM(Adjustments[Debit])-SUM(Adjustments[Credit]),2)=0,"PASS","FAIL")'),
    ]
    for i, (label, f) in enumerate(checks, 28):
        ws.Range(f"B{i}").Value = label
        cell = ws.Range(f"M{i}")
        cell.Formula = f
        cell.Font.Bold = True
        cell.HorizontalAlignment = -4108
        pass_fail(cell)
        if i % 2 == 0:
            ws.Range(f"B{i}:M{i}").Interior.Color = rgb(*LIGHT)
    page(ws, "$A$1:$M$35")
    return (28, 28 + len(checks) - 1)


def format_tables(wb):
    def fmt(sheet, widths, money=(), dates=(), wrap=False):
        ws = wb.Worksheets(sheet)
        lo = ws.ListObjects(1)
        for i in range(1, lo.ListColumns.Count + 1):
            ws.Columns(i).ColumnWidth = widths.get(lo.ListColumns(i).Name, 12)
        for c in money:
            lo.ListColumns(c).DataBodyRange.NumberFormat = MONEY
        for c in dates:
            lo.ListColumns(c).DataBodyRange.NumberFormat = "dd mmm yyyy"
        lo.HeaderRowRange.WrapText = wrap
        ws.Activate(); ws.Range("A2").Select(); wb.Application.ActiveWindow.FreezePanes = True
        return lo

    fmt("Ledger", {"Posted": 12, "MonthNo": 8, "Source": 11, "Line": 9, "Description": 36, "Account": 8, "Account name": 28,
                   "Type": 9, "Debit": 12, "Credit": 12, "Net": 12}, money=("Debit", "Credit", "Net"), dates=("Posted",))
    fmt("Transactions", {"Line": 9, "Register": 24, "Register no": 8, "Posted": 12, "MonthNo": 8, "Bought": 12, "Description": 34,
                         "Clean description": 34, "Amount": 12, "Matched rule": 20, "Account": 8, "Account name": 28},
        money=("Amount",), dates=("Posted", "Bought"), wrap=True)
    lo = fmt("Transfers", {"Amount": 12, "From": 26, "From line": 9, "From date": 12, "To": 26, "To line": 9, "To date": 12,
                           "Days apart": 8, "Status": 30}, money=("Amount",), dates=("From date", "To date"))
    rng = lo.ListColumns("Status").DataBodyRange
    fc = rng.FormatConditions.Add(2, None, '=$I2<>"Paired"'); fc.Interior.Color = rgb(*RED[0])
    fc = rng.FormatConditions.Add(2, None, '=$I2="Paired"'); fc.Interior.Color = rgb(226, 239, 218)
    lo = fmt("Review", {"Issue": 28, "Line": 9, "Register": 24, "Posted": 12, "Description": 30, "Amount": 12, "What to do": 72},
             money=("Amount",), dates=("Posted",))
    lo.ListColumns("Issue").DataBodyRange.Interior.Color = rgb(255, 242, 204)
    fmt("Statements", {"Account (as printed)": 28, "Account no": 10, "Statement date": 14, "MonthNo": 9, "Ending balance": 15},
        money=("Ending balance",), dates=("Statement date",))
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
        ("3. Data > Refresh All. Transactions, ledger, trial balance, P&L, balance sheet and the bank rec all update. No macros.", False),
        ("4. Summary: every data check must say PASS. Then clear the Review tab: add a keyword on Rules for each new merchant, Refresh All again.", False),
        ("5. Type month-end entries that never touch the bank (depreciation, accruals) on Adjustments.", False), ("", False),
        ("How a bank line becomes a ledger entry", True),
        ("The description is cleaned (upper case, punctuation removed) and the Rules tab is read top to bottom: the first keyword found decides the account.", False),
        ("Each line is posted twice: once to its own register (checking, savings or card) and once to the account the rule chose, so debits always equal credits.", False),
        ("Lines no rule knows go to 6990 Uncategorized and are listed on Review; they stay visible in the P&L until explained.", False),
        ("Transfers between your own accounts and card payments go to 1090 Transfers in transit from both sides. The Transfers tab pairs the two halves; "
         "what is left over is money still on its way at the month end.", False),
        ("Card lines are booked by POST date, so each month agrees with the card statement.", False), ("", False),
        ("What this workbook is not", True),
        ("A bookkeeping aid for a small business owner or their bookkeeper. It is not tax or accounting advice; your accountant still prepares the tax return and year-end accounts.", False),
        ("The Power Query steps are readable in Data > Queries & Connections: *_Clean, Transactions, Transfers, Ledger, Review.", False),
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
        build_inputs(wb)
        wb.SaveAs(str(OUT), 51)
        xl.Calculate()
        try:
            wb.Queries.FastCombine = True  # the queries combine the workbook's own tables with the CSV files
        except Exception as e:
            print("FastCombine not set:", e)
        for name, formula in Q.items():
            wb.Queries.Add(name, formula)
        t0 = time.time()
        for title, query in LOADED:
            if query:
                load_query(wb.Worksheets(title), query, query.replace("_", ""))
                print(" loaded", query, f"{time.time() - t0:.1f}s")
        format_tables(wb)
        pl = build_pl(wb)
        tb = build_tb(wb)
        bs = build_bs(wb)
        rec = build_rec(wb)
        build_summary(wb, pl, tb, bs, rec)
        build_howto(wb)
        xl.Calculate()
        tabs = {"Summary": TEAL, "P&L": NAVY, "Trial balance": NAVY, "Balance sheet": NAVY, "Bank rec": NAVY, "Review": (255, 192, 0),
                "Accounts": (255, 230, 120), "Rules": (255, 230, 120), "Budget": (255, 230, 120), "Adjustments": (255, 230, 120)}
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            if ws.Name in tabs:
                ws.Tab.Color = rgb(*tabs[ws.Name])
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
