"""Independent check of Two-Source-Reconciliation.xlsx.

1. Re-does the whole match in plain Python straight from the two CSV files (no Excel involved).
2. Compares it with the answer that was planted when the data was generated (test/expected.json).
3. Opens the workbook with Excel and compares every status, count and total with the Python result.
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs pywin32 for step 3)
"""
import csv, json, sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
results = []


def check(name, ok, detail=""):
    results.append(ok)
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


def key(ref):
    k = "".join(ch for ch in ref.upper() if ch not in " -/.")
    return k or None


# ---------------------------------------------------------------- 1. plain-Python reconciliation
with open(ROOT / "data" / "bank-statement-2026-10.csv", encoding="utf-8") as f:
    raw = list(csv.DictReader(f))
seen, bank = set(), []
for r in raw:
    t = tuple(r.values())
    if t in seen:
        continue
    seen.add(t)
    num = lambda v: float(v.replace(",", "")) if v.strip() else 0.0
    bank.append(dict(key=key(r["Reference"]), amount=round(num(r["Credit"]) - num(r["Debit"]), 2),
                     desc=r["Description"], balance=num(r["Balance"])))
with open(ROOT / "data" / "cash-book-2026-10.csv", encoding="utf-8") as f:
    ledger = [dict(key=key(r["Reference"]), amount=round(float(r["Amount"]), 2), doc=r["Doc No"]) for r in csv.DictReader(f)]

exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
status = []  # (status, bank_amount, ledger_amount, key)
bank_seen = Counter()
first_bank = {}
for b in bank:
    if b["key"]:
        bank_seen[b["key"]] += 1
        if bank_seen[b["key"]] > 1:
            status.append(("Duplicate in bank", b["amount"], None, b["key"])); continue
        first_bank[b["key"]] = b
    else:
        status.append(("Only in bank", b["amount"], None, None))
led_seen, first_led = Counter(), {}
for l in ledger:
    led_seen[l["key"]] += 1
    if led_seen[l["key"]] > 1:
        status.append(("Duplicate in ledger", None, l["amount"], l["key"])); continue
    first_led[l["key"]] = l
for k, b in first_bank.items():
    l = first_led.get(k)
    if not l:
        status.append(("Only in bank", b["amount"], None, k))
    elif round(b["amount"] - l["amount"], 2) == 0:
        status.append(("Matched", b["amount"], l["amount"], k))
    else:
        status.append(("Amount differs", b["amount"], l["amount"], k))
for k, l in first_led.items():
    if k not in first_bank:
        status.append(("Only in ledger", None, l["amount"], k))

py_counts = Counter(s[0] for s in status)
py_bank = defaultdict(float); py_led = defaultdict(float)
for s, ba, la, _ in status:
    py_bank[s] += ba or 0; py_led[s] += la or 0

bank_close = exp["opening"] + sum(b["amount"] for b in bank)
check("Bank file: closing balance printed = opening + movements", round(bank[-1]["balance"] - bank_close, 2) == 0)
check("Bank file: the 2 repeated export rows are dropped", len(raw) - len(bank) == exp["bank_duplicates_in_file"])

# ---------------------------------------------------------------- 2. against the planted answer
planted = Counter(c["status"] for c in exp["cases"])
for st in ["Matched", "Amount differs", "Only in bank", "Only in ledger", "Duplicate in ledger"]:
    check(f"Python result = planted answer: {st}", py_counts[st] == planted[st], f"{py_counts[st]} vs {planted[st]}")
py_diff = {k: round(ba - la, 2) for s, ba, la, k in status if s == "Amount differs"}
for c in exp["cases"]:
    if c["status"] == "Amount differs":
        check(f"Planted difference {c['key']} = {c['diff']:+.2f}", py_diff.get(c["key"]) == c["diff"])

adj_bank = bank_close + py_led["Only in ledger"]
adj_led = (exp["opening"] + sum(l["amount"] for l in ledger) + py_bank["Only in bank"] + py_bank["Duplicate in bank"]
           + sum(py_diff.values()) - py_led["Duplicate in ledger"])
check("Python bank reconciliation balances to 0.00", round(adj_bank - adj_led, 2) == 0, f"{adj_bank - adj_led:.2f}")

# ---------------------------------------------------------------- 3. against the workbook
try:
    import win32com.client as win32
except ImportError:
    print("SKIP  workbook checks (pywin32 not installed)")
else:
    xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(ROOT / "Two-Source-Reconciliation.xlsx"), 0, True)
        lo = wb.Worksheets("Reconciliation").ListObjects("Recon")
        cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
        rows = [dict(zip(cols, r)) for r in lo.DataBodyRange.Value]
        xl_counts = Counter(r["Status"] for r in rows)
        for st in sorted(set(py_counts) | set(xl_counts)):
            check(f"Workbook = Python: {st} lines", xl_counts[st] == py_counts[st], f"{xl_counts[st]} vs {py_counts[st]}")
            xb = round(sum(r["Bank amount"] or 0 for r in rows if r["Status"] == st), 2)
            xlg = round(sum(r["Ledger amount"] or 0 for r in rows if r["Status"] == st), 2)
            check(f"Workbook = Python: {st} amounts", xb == round(py_bank[st], 2) and xlg == round(py_led[st], 2),
                  f"bank {xb} vs {py_bank[st]:.2f}, book {xlg} vs {py_led[st]:.2f}")
        s = wb.Worksheets("Summary")
        check("Workbook Summary: unexplained difference = 0.00", s.Range("D39").Value == 0, str(s.Range("D39").Value))
        check("Workbook Summary: adjusted bank = Python", round(s.Range("D29").Value - adj_bank, 2) == 0)
        for r in range(24, 28):
            check(f"Workbook data check: {s.Range(f'F{r}').Value}", s.Range(f"N{r}").Value == "PASS")
        no_action = [r for r in rows if r["Status"] != "Matched" and not r["What to do"]]
        check("Every line to review has a 'What to do'", not no_action)
        late = sum(1 for r in rows if r["Status"] == "Matched" and (r["Days to clear"] or 0) > 0)
        planted_late = sum(1 for c in exp["cases"] if c["status"] == "Matched" and c.get("gap", 0) > 0)
        check("Timing differences found = planted", late == planted_late, f"{late} vs {planted_late}")
        wb.Close(False)

        # The client's PC: copy workbook + data to another folder, open, Refresh All
        import shutil, tempfile
        with tempfile.TemporaryDirectory() as tmp:
            shutil.copy(ROOT / "Two-Source-Reconciliation.xlsx", tmp)
            shutil.copytree(ROOT / "data", Path(tmp) / "data")
            wb = xl.Workbooks.Open(str(Path(tmp) / "Two-Source-Reconciliation.xlsx"))
            wb.RefreshAll(); xl.CalculateUntilAsyncQueriesDone(); xl.Calculate()
            folder = wb.Worksheets("Settings").Range("B4").Value
            check("Refresh in another folder reads that folder's data", str(folder).lower().startswith(str(tmp).lower()[:12]), folder)
            check("Refresh in another folder still balances to 0.00", wb.Worksheets("Summary").Range("D39").Value == 0)
            wb.Close(False)
    finally:
        xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
