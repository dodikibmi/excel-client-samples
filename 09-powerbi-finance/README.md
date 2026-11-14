# E9 · Power BI finance report: bank and card exports to P&L vs budget, cash and checks

**Checking, savings and credit card exports → Power Query cleans them and a rules table gives every line an account →
a star-schema model (ledger fact table, Date, Account, Category, Budget, Statements) → DAX measures → four report pages:
Overview, P&L vs Budget, Cash & Statements and a Checks page that proves the numbers.**
Next month: drop the new exports in the data folder, press **Refresh**.

![Overview: month and account slicers, six KPI cards, income/expenses/net profit by month, expenses by account](docs/overview.png)

▶ **[Demo video (59 s, MP4)](docs/demo.mp4)**: the raw bank export → Power BI Desktop → Refresh → pick October → P&L vs
budget, cash, checks, star schema. Real screen recording of Power BI Desktop, captions added.

## The problem this solves

A small business (here the fictional *Copperleaf Bakery & Cafe*, the same books as sample [E7](../07-bookkeeping-trial-balance/))
wants its monthly numbers in Power BI instead of a spreadsheet: profit and loss against budget, where the cash went, and
a quick way to see that nothing was missed. The raw material is what the bank and card portals export:

![The raw checking export: title lines above the header, amounts stored as text with $ and commas](docs/before-raw-csv.png)

- three exports, three layouts: title lines above the header, `$1,250.00` as text, Debit/Credit columns vs one signed amount, US and ISO dates
- the same supplier typed several ways (`TapPay Inc. Deposit`, `TAP PAY *DEPOSIT 72350`, `TAPPAY DEPOSIT 0804`)
- transfers and card payments appear in **two** registers, a card purchase on 31 Aug posts on 1 Sep
- a dashboard is only useful if the client can trust it: does it agree with the bank statements?

## What is built

| Step | Where | What happens |
|---|---|---|
| 1 | Power Query `Checking_Clean`, `Savings_Clean`, `Card_Clean` | Read each export, skip the title lines, one signed amount (money in +, out −), parse the dates; card lines are booked by **post** date like the statement |
| 2 | `BankLines` + `rules.csv` | Descriptions cleaned (upper case, punctuation removed); the first keyword that appears decides the account. No match → *6990 Uncategorized*, listed on the Checks page |
| 3 | `Ledger` (fact table) | Double entry: every bank line posts to its register **and** to the chosen account, plus the month-end journal entries (depreciation). 378 lines, debits = credits |
| 4 | Dimensions | `Date` (calendar, marked as date table), `Account` (chart of accounts + opening balances), `Category` (Revenue / Cost of goods sold / Operating expenses / balance sheet), plus `Budget` and `Statements` |
| 5 | 45 DAX measures | Income, expense, net profit, margin, budget, variance, variance %, month-on-month, previous month, YTD, closing balance per account, running cash balance, in transit, books vs statement, the checks |
| 6 | Report | Overview, P&L vs Budget (matrix with red/green variance), Cash & Statements, Checks |

![Model view: the Ledger fact table in the middle, Date, Account and Category around it, Budget and Statements](docs/model-view.png)

A few of the measures (all of them are readable in [`_Measures.tmdl`](Copperleaf-Finance.SemanticModel/definition/tables/_Measures.tmdl)):

```dax
Total Income  = CALCULATE ( -SUM ( Ledger[Amount] ), KEEPFILTERS ( Category[Group] = "Income" ) )
Actual        = CALCULATE ( -SUM ( Ledger[Amount] ), KEEPFILTERS ( Category[Statement] = "P&L" ) )   -- income +, costs -
Variance      = [Actual] - [Budget]                                                                   -- + = better than budget
Income MoM %  = VAR cur = [Income Last Month] VAR prev = [Income Previous Month]
                RETURN IF ( NOT ISBLANK ( prev ), DIVIDE ( cur - prev, prev ) )
Closing Balance =                     -- opening balance + every posting up to the last date in the selection
    VAR d = MAX ( 'Date'[Date] )
    RETURN SUM ( Account[Opening Balance] ) + CALCULATE ( SUM ( Ledger[Amount] ), REMOVEFILTERS ( 'Date' ), 'Date'[Date] <= d )
Cash Balance  = CALCULATE ( [Closing Balance], KEEPFILTERS ( Account[Is Cash] = TRUE () ) )
```

## Result for August to October 2026

| | Aug | Sep | Oct | Quarter |
|---|---:|---:|---:|---:|
| Income | 33,894 | 34,909 | 37,854 | 106,657 |
| Expenses (cost of sales + operating) | 31,830 | 33,895 | 39,813 | 105,538 |
| Net profit | 2,064 | 1,013 | **−1,959** | 1,119 |
| Budgeted net profit | 1,655 | 3,055 | 4,405 | 9,115 |
| Cash at month end (checking + savings) | 42,948 | 40,136 | 35,413 | |
| Books − statements (3 accounts) | **0.00** | **0.00** | **0.00** | |

The same figures as the Excel version (E7), to the cent. The story the report tells: income grew 8.4% from September
to October, but October has **three pay dates** (payroll −6,699 against budget) and turns a budgeted profit of 4,405
into a loss of 1,959; September lost on ingredients and an oven repair.

![P&L vs budget: every account under its category, variance in red or green, net profit by month against budget](docs/pnl-vs-budget.png)

![Cash: balance day by day, cash flow by type of payment, books against the bank and card statements](docs/cash-and-statements.png)

## Proof the numbers are right

The **Checks** page recalculates nine checks on every refresh, ignoring the slicers: opening balances balance, debits =
credits, the trial balance balances at every month end, every export line is in the model exactly once, the net
amount per file (read a second, independent way) = what was posted, books = all 9 bank and card statements, checking
= the bank's own running balance, every line has an account in the chart. The ninth lists the 3 lines that still need
a rule.

![Checks: 8 of 8 PASS, 0 failed, 3 lines waiting for a rule](docs/checks.png)

[`test/verify.py`](test/verify.py) does the bookkeeping again in plain Python straight from the CSV files and compares
it with **(a)** the answer planted when the data was generated (every line's true account, the P&L, cash, statements),
**(b)** the project files: the model has the tables, the 7 star-schema relationships, the measures and the time
intelligence, and every field used by the 33 visuals exists in the model, **(c)** what the model inside Power BI
Desktop answered ([`test/query_model.ps1`](test/query_model.ps1) asks the open report with DAX queries): every month,
all 48 account-months of the P&L, the 9 statements, the cash flow, the check results and the account of all 186 bank
lines. Result: **34/34 checks pass**.

## Use it with your own data

1. Open **`Copperleaf-Finance.pbix`** in Power BI Desktop (free, no account needed): the data is inside.
2. The report reads its CSVs from `C:\Copperleaf\data\`: copy the `data` folder there before pressing Refresh.
3. To load your own exports: put them in a folder, then **Transform data → Edit parameters → DataFolder** = that folder.
   Replace `accounts.csv` (chart of accounts, opening balances), `rules.csv` (keyword → account) and `budget.csv` with yours.
4. **Refresh**. Every line on the Checks page's *Lines waiting for a rule* table needs a keyword in `rules.csv`; add it and refresh again.

A different bank layout only changes its `*_Clean` query; the model, measures and pages stay the same.
The `.pbip` folder is the same report as text files (model in TMDL, report in PBIR), so every change can be reviewed in
Git; it keeps no data, so press *Refresh now* after opening it. Built and tested with Power BI Desktop 2.159 (November 2026).

**This is a reporting aid, not accounting or tax advice.**

## Files

| Path | What |
|---|---|
| `Copperleaf-Finance.pbix` | The report with its data (open this one) |
| `Copperleaf-Finance.pbip`, `*.SemanticModel/`, `*.Report/` | The same report as a Power BI project: Power Query, tables, relationships and DAX in TMDL; pages and visuals in PBIR JSON |
| `data/` | The three fictional exports and statement balances (as in E7), the chart of accounts, rules, budget and journal entries |
| `scripts/build_pbip.py` | Writes the whole project (M, DAX, visuals) from one readable script |
| `scripts/make_tables.py` | Writes the input tables and the planted answer from E7's fictional data |
| `test/verify.py`, `test/expected.json`, `test/query_model.ps1`, `test/model-snapshot.json` | Independent check |
| `scripts/shots.py`, `scripts/make_thumb.py` | Screenshots taken from Power BI Desktop, thumbnail, `.pbix` copy |

All names, businesses and amounts are fictional.
