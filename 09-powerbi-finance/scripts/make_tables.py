"""Writes the input files the Power BI model reads, all fictional ("Copperleaf Bakery & Cafe", Aug-Oct 2026).

The four bank/card exports are the same messy files as sample E7 (copied from ../07-bookkeeping-trial-balance/data),
so the Power BI model and the Excel workbook can be compared line for line. The client's set-up comes from E7's
scripts/coa.py and is written here as plain CSV tables, the way a client would keep them:
  accounts.csv     chart of accounts: number, name, type, opening balance (debit +), cash flag, statement name
  rules.csv        keyword -> account, in priority order (first match wins)
  budget.csv       budget per account per month, typed as positive amounts
  adjustments.csv  month-end journal entries (depreciation)
"""
import csv, json, shutil, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
E7 = ROOT.parent / "07-bookkeeping-trial-balance"
sys.path.insert(0, str(E7 / "scripts"))
from coa import ACCOUNTS, RULES, BUDGET, ADJUSTMENTS, MONTHS  # noqa: E402

DATA = ROOT / "data"
DATA.mkdir(exist_ok=True)
for f in ("checking-2026-08-10.csv", "savings-2026-08-10.csv", "visa-2026-08-10.csv", "statements-2026-08-10.csv"):
    shutil.copyfile(E7 / "data" / f, DATA / f)

STATEMENT_NAME = {1000: "Business Checking ...1180", 1010: "Business Savings ...4821", 2000: "Visa Business ...7702"}
CASH = {1000, 1010}


def write(name, header, rows):
    with open(DATA / name, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"data/{name}: {len(rows)} rows")


write("accounts.csv", ["Account No", "Account", "Type", "Opening Balance", "Is Cash", "Statement Name"],
      [[no, name, typ, f"{dr - cr:.2f}", "Y" if no in CASH else "N", STATEMENT_NAME.get(no, "")]
       for no, name, typ, dr, cr in ACCOUNTS])
write("rules.csv", ["Priority", "Keyword", "Account No", "Note"],
      [[i, k, a, note] for i, (k, a, note) in enumerate(RULES, 1)])
write("budget.csv", ["Account No", "Month", "Budget"],
      [[a, f"{end:%Y-%m}-01", f"{amounts[i]:.2f}"] for a, amounts in BUDGET.items() for i, (_, _, end) in enumerate(MONTHS)])
write("adjustments.csv", ["Date", "Entry", "Description", "Account No", "Debit", "Credit"],
      [[f"{d:%Y-%m-%d}", e, desc, a, f"{dr:.2f}", f"{cr:.2f}"] for d, e, desc, a, dr, cr in ADJUSTMENTS])

# ---------------------------------------------------------------- the planted answer for test/verify.py
# E7's generator knows every line's true account and the true P&L, cash and statement balances (its test/expected.json);
# the budget is the client's own input. Nothing here is computed by the Power BI model.
planted = json.loads((E7 / "test" / "expected.json").read_text(encoding="utf-8"))
types = {a[0]: a[2] for a in ACCOUNTS}
budget_net = {no: round(sum((amounts[i] if types[a] == "Revenue" else -amounts[i]) for a, amounts in BUDGET.items()), 2)
              for i, (no, _, _) in enumerate(MONTHS)}
expected = dict(
    source="planted by 07-bookkeeping-trial-balance/scripts/make_data.py (true accounts, P&L, cash, statements) + budget.csv",
    rows=planted["rows"],
    pl={no: dict(v, budget_net_profit=budget_net[int(no)], variance=round(v["net_profit"] - budget_net[int(no)], 2))
        for no, v in planted["pl"].items()},
    cash=planted["cash"], in_transit=planted["in_transit_by_month"], statements=planted["statements"],
    uncategorized=planted["uncategorized"],
    true_account=[[t["reg"], t["post"], t["desc"], t["amount"], t["acct"]] for t in planted["true_account"]],
    checks={str(i): "PASS" for i in range(1, 9)} | {"9": "REVIEW"},
    model=dict(
        tables=["_Measures", "Ledger", "Account", "Category", "Date", "Budget", "Statements", "Source Files", "Checks"],
        relationships=[["Ledger.Date", "Date.Date"], ["Ledger.Account No", "Account.Account No"],
                       ["Account.Category Key", "Category.Category Key"], ["Budget.Month", "Date.Date"],
                       ["Budget.Account No", "Account.Account No"], ["Statements.Statement Date", "Date.Date"],
                       ["Statements.Account No", "Account.Account No"]],
        measures=["Total Income", "Total Expense", "Net Profit", "Net Margin %", "Actual", "Budget", "Variance", "Variance %",
                  "Income MoM %", "Net Profit YTD", "Closing Balance", "Cash Balance", "Cash Movement", "In Transit",
                  "Statement Balance", "Difference vs Statement", "Check Status", "Check Detail", "Checks Passed"],
        dax_functions=["CALCULATE", "KEEPFILTERS", "REMOVEFILTERS", "PREVIOUSMONTH", "DATESMTD", "TOTALYTD", "LASTDATE",
                       "DIVIDE", "SWITCH", "SELECTEDVALUE"],
        pages=["Overview", "P&L vs Budget", "Cash & Statements", "Checks"]),
)
(ROOT / "test").mkdir(exist_ok=True)
(ROOT / "test" / "expected.json").write_text(json.dumps(expected, indent=1), encoding="utf-8")
print("test/expected.json")
