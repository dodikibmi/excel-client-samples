# E2 · Sales dashboard in Excel: budget vs actual, one page, two drop-downs

**12 months of invoice lines + a budget file → Power Query loads them → one dashboard page that answers
"are we on budget, where not, and why", for any month and any region.**
No macros, no pivot to refresh by hand. Next month: drop in the new export, press **Refresh All**.

![Dashboard, October 2026, all regions](docs/dashboard-all.png)

▶ **[Demo video (49 s, MP4)](docs/demo.mp4)**: pick a region and a month, watch every number follow. Real screen recording, captions added.

## The problem this solves

A wholesaler (here the fictional *Brightline Supply Co.*) gets one big export from its accounting system every
month and a budget from finance. Someone spends a day each month pasting both into a spreadsheet, building
pivots and charts, and still cannot answer the manager's follow-up question ("and only West?") without starting over.

![The raw export: one row per invoice line](docs/before-raw-csv.png)

## What the workbook does

| Sheet | What it holds |
|---|---|
| **Dashboard** | Month and region drop-downs · 5 KPI cards (revenue vs budget, budget achieved, gross margin vs last month, orders and average order, year to date vs budget) · revenue vs budget for 12 months (selected month highlighted) · category actual vs budget · region table with red/green variance · top 10 customers · sales-rep leaderboard · 5 PASS/FAIL data checks · print-ready on one landscape page |
| **Calc** | Every number the dashboard shows, as plain `SUMIFS` formulas on two Excel Tables, so any figure can be traced |
| **Sales data**, **Budget data** | The two Power Query tables (typed dates and numbers, revenue, cost, margin, one order flag per invoice) |
| **Settings** | Data folder (defaults to the `data` folder next to the workbook), file names, company name |
| **How to use** | Daily and monthly steps for the client |

Power Query also builds the customer and sales-rep lists and a `Checks` table (rows loaded, sales without a budget,
repeated budget rows, total revenue), which feed the data checks on the dashboard.

## What it shows for October 2026

| | All regions | West only |
|---|---|---|
| Revenue | 392,296 | 85,324 |
| Budget achieved | 96.5% | **88.0%** |
| Gross margin | 32.3% (−0.7 pp vs September) | 32.5% |
| Orders · average order | 434 · 904 | 98 · 871 |
| Year to date vs budget | −2.3% | −3.0% |

Pick **West** and the reason is on the screen: **Garden is 61% below budget** while every other category is within 10%.
The other three regions are within 4% of budget.

![Dashboard, October 2026, West region: the Garden miss](docs/dashboard-west.png)

![Calc sheet: every dashboard number as a traceable formula](docs/calc.png)

## Proof the numbers are right

[`test/verify.py`](test/verify.py) recomputes everything in plain Python straight from the two CSV files, then opens
the workbook in Excel, switches the drop-downs to four month/region combinations and compares the KPIs, the 12-month
trend, categories, regions, top 10 customers and the rep leaderboard. It also checks the planted story (West misses
because of Garden), that all five data checks say PASS and that no formula needs Microsoft 365 only features.
Result: **67/67 checks pass**.

## Use it with your own data

1. Export invoice lines (date, invoice, customer, region, sales rep, category, qty, unit price, unit cost) and the
   budget (month, region, category, budget revenue) as CSV into the `data` folder.
2. **Data → Refresh All**. New months appear in the month drop-down automatically.
3. Check that the five data checks at the bottom of the dashboard say PASS.

Different column names: only the `Sales` and `Budget` queries change; Calc and the dashboard stay the same.
Works in Excel 2016+ / Microsoft 365 for Windows.

## Files

| Path | What |
|---|---|
| `Sales-Dashboard.xlsx` | The workbook (no macros) |
| `data/` | Fictional sales lines (9,899 rows, 5,546 invoices) and budget (240 rows) |
| `scripts/make_data.py` | Generates the fictional data, with seasonality, growth and the West/Garden miss planted on purpose |
| `scripts/build_workbook.py` | Builds the workbook with real Excel (Power Query M code and every formula are readable in the script) |
| `test/verify.py`, `test/story.json` | Independent check |
| `scripts/shots.ps1`, `scripts/make_thumb.py` | Screenshots in `docs/` |

All names, companies and amounts are fictional.
