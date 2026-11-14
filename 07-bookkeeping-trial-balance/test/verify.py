"""Independent check of Bookkeeping-Trial-Balance.xlsx.

1. Re-does the bookkeeping in plain Python straight from the four CSV files (no Excel involved): parse, clean,
   categorise with the rules, pair transfers, post a ledger, trial balance, P&L, statement balances.
2. Compares it with the answer planted when the data was generated (test/expected.json): every line's true account,
   every account's movement per month, the P&L, cash, in-transit money and the statement balances.
3. Opens the workbook with Excel and compares every line's account, the P&L, trial balance, balance sheet, bank rec,
   Review tab and the eight in-workbook checks with Python.
4. Copies workbook + data to another folder, refreshes there, and checks again.
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs pywin32 for steps 3-4)
"""
import csv, json, sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from coa import ACCOUNTS, RULES, ADJUSTMENTS, MONTHS, CLEARING, UNCATEGORIZED, BUDGET

results = []


def check(name, ok, detail=""):
    results.append(bool(ok))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


def money(t):
    t = (t or "").strip().replace("$", "").replace(",", "").replace(" ", "")
    return float(t) if t else 0.0


def norm(t):
    t = "".join(" " if ch in "*.-#/,&'" else ch for ch in t.upper())
    return " ".join(t.split())


def us(d):
    m, dd, y = d.strip().split("/")
    return datetime(int(y), int(m), int(dd)).date()


# ---------------------------------------------------------------- 1. plain Python
data = ROOT / "data"
lines = []
with open(data / "checking-2026-08-10.csv", encoding="utf-8") as f:
    body = f.read().splitlines()[3:]
for i, r in enumerate(csv.DictReader(body), 1):
    lines.append(dict(line=f"CHK-{i:03d}", reg=1000, post=us(r["Date"]), desc=r["Description"].strip(), amount=round(money(r["Credit"]) - money(r["Debit"]), 2)))
with open(data / "savings-2026-08-10.csv", encoding="utf-8") as f:
    for i, r in enumerate(csv.DictReader(f), 1):
        lines.append(dict(line=f"SAV-{i:03d}", reg=1010, post=datetime.strptime(r["Posted"].strip(), "%Y-%m-%d").date(),
                          desc=r["Details"].strip(), amount=money(r["Amount"])))
with open(data / "visa-2026-08-10.csv", encoding="utf-8") as f:
    for i, r in enumerate(csv.DictReader(f), 1):
        lines.append(dict(line=f"VISA-{i:03d}", reg=2000, post=us(r["Post Date"]), desc=r["Description"].strip(), amount=money(r["Amount"])))
rules = [(norm(k), a) for k, a, _ in RULES]
for l in lines:
    l["acct"] = next((a for k, a in rules if k in norm(l["desc"])), UNCATEGORIZED)
    l["mo"] = l["post"].year * 100 + l["post"].month

exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
check("Rows read = rows planted (checking, savings, card)",
      [sum(l["reg"] == r for l in lines) for r in (1000, 1010, 2000)] == [exp["rows"]["checking"], exp["rows"]["savings"], exp["rows"]["card"]])
truth = defaultdict(list)
for t in exp["true_account"]:
    truth[(t["reg"], t["post"], norm(t["desc"]), t["amount"])].append(t["acct"])
wrong = []
for l in lines:
    k = (l["reg"], l["post"].isoformat(), norm(l["desc"]), l["amount"])
    want = truth[k].pop(0) if truth[k] else None
    if want != l["acct"]:
        wrong.append(f"{l['line']} {l['desc']}: {l['acct']} vs {want}")
check(f"Rules give all {len(lines)} lines their planted account", not wrong, "; ".join(wrong[:3]))

# pair transfer halves: same amount, opposite direction, within 5 days, earliest first
outs = sorted([l for l in lines if l["acct"] == CLEARING and l["amount"] < 0], key=lambda l: (l["post"], l["line"]))
ins = sorted([l for l in lines if l["acct"] == CLEARING and l["amount"] > 0], key=lambda l: (l["post"], l["line"]))
pairs, unpaired = [], []
for o in outs:
    m = next((i for i in ins if abs(i["amount"] + o["amount"]) < 0.005 and abs((i["post"] - o["post"]).days) <= 5), None)
    if m:
        ins.remove(m); pairs.append((o, m))
    else:
        unpaired.append(o)
unpaired += ins
check("Transfers: 6 pairs, 1 half still in transit (the 30 Oct card payment)",
      len(pairs) == 6 and [round(u["amount"], 2) for u in unpaired] == [-1800.0], f"{len(pairs)} pairs, {[u['amount'] for u in unpaired]}")

net = defaultdict(float)
for l in lines:
    net[(l["reg"], l["mo"])] += l["amount"]
    net[(l["acct"], l["mo"])] -= l["amount"]
for d, _, _, a, dr, cr in ADJUSTMENTS:
    net[(a, d.year * 100 + d.month)] += dr - cr
got = {f"{a}-{mo}": round(v, 2) for (a, mo), v in net.items() if round(v, 2)}
check("Every account's movement per month = planted answer", got == exp["accounts_by_month"],
      str(set(got.items()) ^ set(exp["accounts_by_month"].items()))[:200])

types = {a[0]: a[2] for a in ACCOUNTS}
opening = {a[0]: round(a[3] - a[4], 2) for a in ACCOUNTS}
months = [no for no, _, _ in MONTHS]


def balance(a, no):
    return round(opening[a] + sum(v for (aa, mo), v in net.items() if aa == a and mo <= no), 2)


py_pl, py_tb, py_bs = {}, {}, {}
for no in months:
    s = lambda t: round(sum(-v for (a, mo), v in net.items() if mo == no and types[a] == t), 2)
    rev, cogs, ex = s("Revenue"), -s("COGS"), -s("Expense")
    py_pl[no] = dict(revenue=rev, cogs=cogs, gross_profit=round(rev - cogs, 2), expenses=ex, net_profit=round(rev - cogs - ex, 2))
    bals = [balance(a, no) for a in opening]
    py_tb[no] = (round(sum(b for b in bals if b > 0), 2), round(-sum(b for b in bals if b < 0), 2))
    check(f"{no}: P&L = planted answer", py_pl[no] == {k: round(v, 2) for k, v in exp["pl"][str(no)].items()}, f"{py_pl[no]}")
    check(f"{no}: trial balance balances in Python (debits {py_tb[no][0]:,.2f})", py_tb[no][0] == py_tb[no][1])
    check(f"{no}: cash (checking + savings) = planted", round(balance(1000, no) + balance(1010, no), 2) == exp["cash"][str(no)])
    check(f"{no}: transfers in transit = planted", balance(CLEARING, no) == exp["in_transit_by_month"][str(no)])
with open(data / "statements-2026-08-10.csv", encoding="utf-8") as f:
    stmts = {}
    for r in csv.DictReader(f):
        a = 1000 if "CHECKING" in r["Account"].upper() else 1010 if "SAVINGS" in r["Account"].upper() else 2000
        d = us(r["Statement date"])
        stmts[(a, d.year * 100 + d.month)] = money(r["Ending balance"])
bad = [(a, no) for (a, no), v in stmts.items() if round((-balance(a, no) if a == 2000 else balance(a, no)) - v, 2)]
check("Books = all 9 statement balances (Python)", not bad and len(stmts) == 9, str(bad))
check("Uncategorized lines = the 3 planted", sorted((l["desc"], l["amount"]) for l in lines if l["acct"] == UNCATEGORIZED)
      == sorted((u["desc"], u["amount"]) for u in exp["uncategorized"]))

# ---------------------------------------------------------------- 3. against the workbook
try:
    import win32com.client as win32
except ImportError:
    print("SKIP  workbook checks (pywin32 not installed)")
    win32 = None


def find_row(ws, col, text, last=80):
    for r in range(1, last):
        if ws.Range(f"{col}{r}").Value == text:
            return r
    raise KeyError(text)


def check_workbook(wb, tag=""):
    lo = wb.Worksheets("Transactions").ListObjects("Transactions")
    cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
    xrows = {r[cols.index("Line")]: r for r in lo.DataBodyRange.Value}
    diff = [l["line"] for l in lines if l["line"] not in xrows or int(xrows[l["line"]][cols.index("Account")]) != l["acct"]
            or round(xrows[l["line"]][cols.index("Amount")] - l["amount"], 2)]
    check(f"{tag}Every line in the workbook has the Python account and amount ({len(lines)} lines)", not diff and len(xrows) == len(lines), str(diff[:3]))
    led = wb.Worksheets("Ledger").ListObjects("Ledger").ListRows.Count
    check(f"{tag}Ledger = 2 lines per bank line + {len(ADJUSTMENTS)} adjustment lines", led == 2 * len(lines) + len(ADJUSTMENTS), str(led))
    p = wb.Worksheets("P&L")
    rows = {k: find_row(p, "C", t) for k, t in (("revenue", "Total revenue"), ("cogs", "Total cost of goods sold"),
                                                 ("gross_profit", "Gross profit"), ("expenses", "Total operating expenses"), ("net_profit", "Net profit"))}
    for k, no in enumerate(months):
        col = chr(ord("D") + 3 * k)
        got = {key: round(p.Range(f"{col}{r}").Value, 2) for key, r in rows.items()}
        check(f"{tag}{no}: workbook P&L = Python", got == py_pl[no], f"{got}")
        bcol = chr(ord("E") + 3 * k)
        bud = round(p.Range(f"{bcol}{rows['net_profit']}").Value, 2)
        want = sum(v[k] * (1 if types[a] == "Revenue" else -1) for a, v in BUDGET.items())
        check(f"{tag}{no}: budget net profit = Budget tab", bud == round(want, 2), f"{bud} vs {want}")
    t = wb.Worksheets("Trial balance")
    tot = find_row(t, "C", "Total")
    for k, no in enumerate(months):
        a, b = chr(ord("G") + 2 * k), chr(ord("H") + 2 * k)
        got = (round(t.Range(f"{a}{tot}").Value, 2), round(t.Range(f"{b}{tot}").Value, 2))
        check(f"{tag}{no}: trial balance totals = Python and balanced", got == py_tb[no], f"{got} vs {py_tb[no]}")
    r = wb.Worksheets("Bank rec")
    recs = [(r.Range(f"J{i}").Value, r.Range(f"E{i}").Value, r.Range(f"F{i}").Value, r.Range(f"H{i}").Value) for i in range(5, 14)]
    check(f"{tag}Bank rec: 9 rows, books = statement = Python, all PASS",
          all(st == "PASS" and round(bk - stmts[(int(a), months[i % 3])], 2) == 0 and round(sb - stmts[(int(a), months[i % 3])], 2) == 0
              for i, (a, bk, sb, st) in enumerate(recs)), str(recs[:2]))
    rv = wb.Worksheets("Review").ListObjects("Review")
    rvals = sorted((x[0], x[4], round(x[5], 2)) for x in rv.DataBodyRange.Value)
    want = sorted([("No rule matched", u["desc"], u["amount"]) for u in exp["uncategorized"]] +
                  [("Transfer without its other half", u["desc"], round(u["amount"], 2)) for u in unpaired])
    check(f"{tag}Review tab = the 3 unknown lines + the unpaired transfer half", rvals == want, str(rvals))
    s = wb.Worksheets("Summary")
    s.Range("C3").Value = "Oct 2026"; wb.Application.Calculate()
    check(f"{tag}Summary cards for Oct 2026 = Python (revenue, net profit, cash)",
          round(s.Range("B7").Value, 2) == py_pl[202610]["revenue"] and round(s.Range("D7").Value, 2) == py_pl[202610]["net_profit"]
          and round(s.Range("K7").Value, 2) == exp["cash"]["202610"])
    s.Range("C3").Value = "Aug 2026"; wb.Application.Calculate()
    check(f"{tag}Summary follows the month picker (Aug 2026)", round(s.Range("D7").Value, 2) == py_pl[202608]["net_profit"])
    s.Range("C3").Value = "Oct 2026"; wb.Application.Calculate()
    states = [(s.Range(f"B{i}").Value, s.Range(f"M{i}").Value) for i in range(28, 36)]
    for label, st in states:
        check(f"{tag}Workbook data check: {label}", st == "PASS")


if win32:
    xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(ROOT / "Bookkeeping-Trial-Balance.xlsx"), 0, True)
        check_workbook(wb)
        wb.Close(False)

        import shutil, tempfile
        with tempfile.TemporaryDirectory() as tmp:
            shutil.copy(ROOT / "Bookkeeping-Trial-Balance.xlsx", tmp)
            shutil.copytree(ROOT / "data", Path(tmp) / "data")
            wb = xl.Workbooks.Open(str(Path(tmp) / "Bookkeeping-Trial-Balance.xlsx"))
            wb.RefreshAll(); xl.CalculateUntilAsyncQueriesDone(); xl.Calculate()
            folder = wb.Worksheets("Settings").Range("B4").Value
            check("Refresh in another folder reads that folder's data", str(folder).lower().startswith(str(tmp).lower()[:12]), folder)
            check_workbook(wb, "After refresh: ")
            wb.Close(False)
    finally:
        xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
