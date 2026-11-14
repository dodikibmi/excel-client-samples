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

Built by a banking back-office developer: reconciliation, controls and audit trails are the day job,
so every workbook here balances, logs what it changed, and can be re-run by the client without the author.

All names, companies and amounts are fictional. MIT licence.
