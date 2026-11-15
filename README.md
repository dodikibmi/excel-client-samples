# Excel client samples

Working Excel files built the way small businesses and finance teams ask for them on freelance boards:
messy exports in, clean answers out, and **proof that the numbers are right**.
Every sample ships with fictional data, screenshots, a 50–60 second demo video (`docs/demo.mp4`), a How-to-use sheet and an automated check.

| # | Sample | What it shows | Status |
|---|---|---|---|
| E1 | [Bank reconciliation: two sources, one balanced answer](01-two-source-reconciliation/) | Power Query cleaning + full outer match, exceptions with next steps, reconciliation statement to 0.00, 31 automated checks | ✅ |
| E2 | [Sales dashboard: budget vs actual](02-sales-dashboard/) | Month + region drop-downs, 5 KPI cards, 12-month budget vs actual, category and region variance, top 10 customers, rep leaderboard, print-ready, 67 automated checks | ✅ |
| E3 | [Sales & commission calculator](03-commission-calculator/) | Order entry with drop-downs and validation, tiered bands (whole or stepped) + target bonus, bad lines held back, printable rep statements, locked formulas, 17 hand-worked test cases, 36 automated checks | ✅ |
| E4 | [Invoice tracker with payment reminders](04-sheets-invoice-tracker/) | Google Sheets + Apps Script: ARRAYFORMULA/QUERY/SPARKLINE dashboard, staged reminder emails (never twice, preview first), reminder log, weekly summary, daily trigger with lock, self-test 31/31 in Google + 30 offline checks in Node | ✅ |
| E5 | [FTE comparison: two sources, pivots + discrepancies](05-fte-comparison/) | HR roster vs Finance budget lined up in Power Query (IDs, fund centers, FTE % normalised), three PivotTables by fund center x functional area (HR, budget, difference), a Discrepancies tab with both sides and a next step, six in-workbook checks, 51 automated checks | ✅ |
| E6 | [Customer list cleanup: three sources, duplicates merged, CRM-ready](06-customer-dedupe/) | Webshop + POS + old sales sheet cleaned in Power Query, duplicates linked by email, phone or name + ZIP (nicknames too), newest record wins, Merge log, look-alikes left for review, HubSpot/Pipedrive import tab, 33 automated checks | ✅ |
| E7 | [Small business bookkeeping: bank and card recaps to trial balance and P&L](07-bookkeeping-trial-balance/) | Checking, savings and card exports cleaned, Rules tab categorises 183 of 186 lines, transfers paired, trial balance 0.00 every month end, P&L vs budget, balance sheet, books = statements, 65 automated checks | ✅ |
| E8 | [Project cost tracker: PM sheets roll up to a VP dashboard](08-project-cost-tracker/) | One sheet per project manager, VP dashboard with CPI, forecast at completion and red/amber/green status, field vs shop progress, weekly History saved by Refresh All (no macros, never doubled), 44 automated checks | ✅ |
| E9 | [Power BI finance report: bank and card exports to P&L vs budget, a 13-week cash forecast and checks](09-powerbi-finance/) | Power Query cleans 3 exports + rules table, star schema (ledger, Date, Account, Category, Budget, Statements, Forecast), 105 DAX measures incl. time intelligence, running balances, what-if cash buffer and DAX-written key insights; 7 pages in a dark gold theme (Overview, P&L vs Budget, Revenue & Costs, Cash & Banks, 13-week Cash Forecast, Ledger, Checks: books = 9 statements), .pbix + PBIP (TMDL/PBIR), 41 automated checks against the live model | ✅ |

Built by a banking back-office developer: reconciliation, controls and audit trails are the day job,
so every workbook here balances, logs what it changed, and can be re-run by the client without the author.

All names, companies and amounts are fictional. MIT licence.
