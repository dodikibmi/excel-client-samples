"""Independent check of Two-Source-Reconciliation.xlsx.

1. Re-does the whole reconciliation in plain Python straight from the CSV files (test/recon_py.py, no Excel involved).
2. Compares it with the answer planted when the data was generated (test/expected.json).
3. Opens the workbook with Excel and compares every status, rule, pair, statement line, KPI card, chart series,
   trend point, aging cell, rule count, check, cleaning count and workbench note with the Python result,
   for several account/month picks.
4. Copies workbook + data to another folder and refreshes there: notes survive, a newly resolved item flows into the
   Summary, and changed tolerances give exactly the Python result for those tolerances.
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs pywin32 for steps 3-4)
"""
import json, shutil, sys, tempfile
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import recon_py as R

ROOT = Path(__file__).resolve().parent.parent
results = []
STATUSES = ["Amount differs", "Only in bank", "Only in ledger", "Duplicate in bank", "Duplicate in ledger"]
RULES = ["Reference", "Reference ± amount", "Amount + date"]


def check(name, ok, detail=""):
    results.append(bool(ok))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


def eq(a, b, tol=0.005):
    if a is None or b is None:
        return a == b
    return abs(float(a) - float(b)) <= tol


exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
notes = exp["notes"]
OPEN = exp["opening"]
MONTHS = [date.fromisoformat(m + "-01") for m in exp["months"]]

# ---------------------------------------------------------------- 1. plain-Python reconciliation from the CSVs
data = R.load()
rows = R.reconcile(data, exp["amount_tolerance"], exp["days_tolerance"])
for acct in OPEN:
    bal, ok = OPEN[acct], True
    for b in sorted([b for b in data["bank"] if b["account"] == acct], key=lambda b: b["id"]):
        bal = round(bal + b["amount"], 2)
        ok &= round(bal - b["balance"], 2) == 0
    check(f"Bank files {acct}: every printed balance = opening + movements so far", ok)
check("Bank files: the repeated export rows are dropped",
      data["raw_bank"] - len(data["bank"]) == sum(f.get("repeated", 0) for f in exp["files"]))
check("All 24 files read (bank + cash book, 2 accounts x 6 months)", len(data["files"]) == len(exp["files"]) == 24)

# ---------------------------------------------------------------- 2. against the planted answer
py_c = Counter((r["account"], r["status"], r["rule"]) for r in rows)
pl_c = Counter((c["account"], c["status"], c["rule"]) for c in exp["cases"])
for k in sorted(set(py_c) | set(pl_c)):
    check(f"Python = planted: {k[0]} / {k[1]}{' / rule ' + k[2] if k[2] else ''}", py_c[k] == pl_c[k], f"{py_c[k]} vs {pl_c[k]}")
py_pairs = {(r["b"]["id"] if r["b"] else None, r["l"]["id"] if r["l"] else None, r["status"], r["rule"]) for r in rows}
pl_pairs = {(c.get("bank_id"), c.get("ledger_id"), c["status"], c["rule"]) for c in exp["cases"]}
check("Python = planted: every bank line paired with the planted cash book line (all rules)", py_pairs == pl_pairs,
      f"{len(py_pairs ^ pl_pairs)} differ")
py_diff = {r["b"]["id"]: r["difference"] for r in rows if r["status"] == "Amount differs" or r["rule"] == "Reference ± amount"}
for c in exp["cases"]:
    if "diff" in c:
        check(f"Planted difference {c['bank_id']} = {c['diff']:+.2f}", eq(py_diff.get(c["bank_id"]), c["diff"]))
st_py = {}
for acct in OPEN:
    for s in MONTHS:
        st = R.statement(rows, acct, s, OPEN[acct], OPEN[acct])
        st_py[(acct, s)] = st
        check(f"Python statement {acct} {s:%b %Y} balances to 0.00", st["unexplained"] == 0, str(st["unexplained"]))


# ---------------------------------------------------------------- 3. against the workbook
def table(wb, sheet, name):
    lo = wb.Worksheets(sheet).ListObjects(name)
    cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
    if lo.DataBodyRange is None:
        return []
    return [dict(zip(cols, r)) for r in lo.DataBodyRange.Value]


def d(v):
    return v.date() if isinstance(v, datetime) or hasattr(v, "year") and hasattr(v, "hour") else v


def nm(wb, name):
    return wb.Names(name).RefersToRange.Value


def serial(dt):
    """Excel date serial: COM would shift a naive datetime through UTC (one day back at UTC+7)."""
    return (date(dt.year, dt.month, dt.day) - date(1899, 12, 30)).days


def pick(wb, xl, account, month):
    s = wb.Worksheets("Summary")
    s.Range("C7").Value = account
    s.Range("I7").Value = serial(month)
    xl.Calculate()


def check_pick(wb, xl, account, month, rows, notes, tag=""):
    """Every number on the Summary for one account/month pick, recomputed in Python."""
    pick(wb, xl, account, month)
    s = wb.Worksheets("Summary")
    calc = wb.Worksheets("Calc")
    t = f"{tag}{account} {month:%b %Y}"
    st = R.statement(rows, account, month, OPEN[account], OPEN[account])
    lines = {33: "bank_open", 34: "bank_moves", 35: "bank_close", 36: "in_transit", 37: "outstanding", 38: "adj_bank",
             40: "book_open", 41: "book_moves", 42: "book_close", 43: "bank_items", 44: "differences", 45: "twice",
             46: "adj_book", 48: "unexplained"}
    bad = [f"{k}: {s.Range(f'K{r}').Value} vs {st[k]}" for r, k in lines.items() if not eq(s.Range(f"K{r}").Value, st[k])]
    check(f"Statement {t}: all 14 lines = Python", not bad, "; ".join(bad))
    check(f"Statement {t}: unexplained difference 0.00", s.Range("K48").Value == 0)
    k = R.month_kpis(rows, account, month, notes)
    check(f"KPI {t}: match rate {k['match_rate']:.1%}", eq(nm(wb, "MatchRate"), k["match_rate"], 1e-9))
    check(f"KPI {t}: matched amount {k['matched_amount']:,.2f}", eq(nm(wb, "MatchedAmt"), k["matched_amount"]))
    check(f"KPI {t}: open exceptions {k['open_exceptions']}, resolved {k['resolved']}",
          nm(wb, "OpenExc") == k["open_exceptions"] and nm(wb, "ResolvedExc") == k["resolved"])
    old = sum(v for (st_, b), v in k["aging"].items() if b == "30+")
    check(f"KPI {t}: open over 30 days {old}", nm(wb, "OldExc") == old)
    donut = [k["by_rule"]["Reference"] + k["by_rule"]["Reference ± amount"], k["by_rule"]["Amount + date"], k["exceptions_in_month"]]
    xl_donut = [calc.Range(f"E{r}").Value for r in (3, 4, 5)]
    check(f"Donut {t}: {donut}", xl_donut == donut, str(xl_donut))
    bars = [round(k["by_status_value"][x], 2) for x in STATUSES]
    xl_bars = [calc.Range(f"E{r}").Value for r in range(8, 13)]
    check(f"Status bars {t}", all(eq(a, b) for a, b in zip(xl_bars, bars)), f"{xl_bars} vs {bars}")
    tr = R.trend(rows, account, MONTHS, notes)
    xl_tr = [(d(calc.Range(f"C{r}").Value), calc.Range(f"E{r}").Value, calc.Range(f"F{r}").Value) for r in range(15, 21)]
    check(f"Trend {t}: 6 months of auto-match % and open difference",
          all(a[0] == b[0] and eq(a[1], b[1], 1e-9) and eq(a[2], b[2]) for a, b in zip(xl_tr, tr)), f"{xl_tr} vs {tr}")
    grid_ok, amt_ok = True, True
    for i, stt in enumerate(STATUSES, 34):
        for col, b in (("Q", "0-7"), ("T", "8-30"), ("V", "30+")):
            grid_ok &= s.Range(f"{col}{i}").Value == k["aging"][(stt, b)]
    for col, b in (("Q", "0-7"), ("T", "8-30"), ("V", "30+")):
        amt_ok &= eq(s.Range(f"{col}39").Value, k["aging_amount"][b])
    check(f"Aging heat grid {t}: 15 counts", grid_ok)
    check(f"Aging heat grid {t}: open amount per bucket", amt_ok)
    six = Counter(r["rule"] for r in rows if r["account"] == account and r["status"] == "Matched")
    ok = all(s.Range(f"T{44 + i}").Value == k["by_rule"][rule] and s.Range(f"V{44 + i}").Value == six[rule] for i, rule in enumerate(RULES))
    check(f"Matches per rule {t}: month and 6 months", ok)
    checks = [s.Range(f"K{r}").Value for r in range(53, 61)]
    check(f"Data checks {t}: 8 of 8 PASS", checks == ["PASS"] * 8, str(checks))
    return k


try:
    import win32com.client as win32
except ImportError:
    print("SKIP  workbook checks (pywin32 not installed)")
else:
    xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(ROOT / "Two-Source-Reconciliation.xlsx"), 0, True)
        rec = table(wb, "Reconciliation", "Recon")
        xl_c = Counter((r["Account"], r["Status"], r["Match rule"] or "") for r in rec)
        for k in sorted(set(py_c) | set(xl_c)):
            check(f"Workbook = Python: {k[0]} / {k[1]}{' / rule ' + k[2] if k[2] else ''} lines", xl_c[k] == py_c[k], f"{xl_c[k]} vs {py_c[k]}")
        xb, xlg, pb, plg = defaultdict(float), defaultdict(float), defaultdict(float), defaultdict(float)
        for r in rec:
            xb[(r["Account"], r["Status"])] += r["Bank amount"] or 0; xlg[(r["Account"], r["Status"])] += r["Ledger amount"] or 0
        for r in rows:
            pb[(r["account"], r["status"])] += r["bank_amount"] or 0; plg[(r["account"], r["status"])] += r["ledger_amount"] or 0
        check("Workbook = Python: bank and cash book totals per account and status",
              all(eq(xb[k], pb[k]) and eq(xlg[k], plg[k]) for k in set(pb) | set(xb)))
        xl_pairs = {(r["Bank ID"], r["Ledger ID"], r["Status"], r["Match rule"] or "") for r in rec}
        check("Workbook = Python: the same bank/cash book pairs, status and rule for all lines", xl_pairs == py_pairs,
              f"{len(xl_pairs ^ py_pairs)} differ")
        check("Every line to review has a 'What to do'", all(r["What to do"] for r in rec))
        late = sum(1 for r in rec if r["Status"] == "Matched" and r["Match rule"] == "Reference" and (r["Days to clear"] or 0) > 0)
        planted_late = sum(1 for c in exp["cases"] if c["status"] == "Matched" and c["rule"] == "Reference" and c.get("gap", 0) > 0)
        check(f"Timing differences found = planted ({planted_late})", late == planted_late, f"{late} vs {planted_late}")
        cross = sum(1 for r in rec if r["Bank date"] and r["Ledger date"] and d(r["Bank date"]).month != d(r["Ledger date"]).month)
        check("Lines that clear in the next month are matched across the month end", cross >= 15, str(cross))

        # workbench: one line per exception, demo notes kept by Exception ID after the build's refresh
        wbk = table(wb, "Exceptions workbench", "Workbench")
        py_exc = {r["exception_id"]: r for r in rows if r["status"] != "Matched"}
        check(f"Workbench: exactly the {len(py_exc)} exceptions", sorted(r["Exception ID"] for r in wbk) == sorted(py_exc))
        ok = all(eq(r["Amount"], py_exc[r["Exception ID"]]["wb_amount"]) and d(r["Date"]) == py_exc[r["Exception ID"]]["date"] for r in wbk)
        check("Workbench: amount and date of every exception = Python", ok)
        got = {r["Exception ID"]: r for r in wbk}
        ok = all(got[i]["Owner"] == n["owner"] and got[i]["Comment"] == n["comment"]
                 and (d(got[i]["Resolved date"]).isoformat() if got[i]["Resolved date"] else None) == n["resolved"] for i, n in notes.items())
        check(f"Workbench: the {len(notes)} typed notes survived the refresh, on the right lines", ok,
              "; ".join(f"{i}: {got[i]['Resolved date']}" for i in notes))
        ok = all((r["State"] == "Resolved") == bool((notes.get(r["Exception ID"]) or {}).get("resolved")) for r in wbk)
        asat = max(b["date"] for b in data["bank"]).replace(day=31)
        ok &= all(r["Flag"] == (1 if r["State"] == "Resolved" else (3 if (asat - d(r["Date"])).days > 30 else 2)) for r in wbk)
        check("Workbench: State and status icon (flag) of every line", ok)

        # cleaning log, statement controls, audit trail
        log = {(r["Source"], r["Check"]): r["Rows"] for r in table(wb, "Cleaning log", "CleaningLog")}
        std = lambda x: x == x.upper() and " " not in x and "-" in x
        check("Cleaning log: repeated export rows", log[("Bank statement", "Exact duplicate rows")] == data["raw_bank"] - len(data["bank"]))
        check("Cleaning log: references normalised",
              log[("Bank statement", "References typed differently (case, spaces, dashes)")] == sum(1 for b in data["bank"] if b["key"] and not std(b["ref"])))
        check("Cleaning log: bank rows without a reference", log[("Bank statement", "Rows without a reference")] ==
              sum(1 for b in data["bank"] if not b["key"]), str(log[("Bank statement", "Rows without a reference")]))
        check("Cleaning log: cash book rows without a reference", log[("Cash book", "Rows without a reference")] ==
              sum(1 for l in data["ledger"] if not l["key"]))
        bc = table(wb, "Cleaning log", "BankControls")
        closes = {(f["account"], f["period"]): f["closing"] for f in exp["files"] if "closing" in f}
        check("Statement controls: printed closing of all 12 files = planted",
              all(eq(r["Printed closing"], closes[(r["Account"], f"{d(r['Month']):%Y-%m}")]) for r in bc) and len(bc) == 12)
        check("Statement controls: every file OK inside and chained to the previous month",
              all(r["Inside the file"] == "OK" and r["Chain"] in ("OK", "First file") for r in bc))
        ri = {r["Item"]: r["Value"] for r in table(wb, "Cleaning log", "RunInfo")}
        check("Audit trail: files and rows read", ri["Files read"] == 24 and ri["Bank rows read"] == data["raw_bank"]
              and ri["Cash book rows"] == data["raw_ledger"])

        # every account x month balances (Calc grid) = Python statement
        calc = wb.Worksheets("Calc")
        for r in range(3, 27):
            acct = calc.Range(f"K{r}").Value
            if not acct:
                continue
            m = d(calc.Range(f"L{r}").Value)
            st = st_py[(acct, m)]
            check(f"Calc grid {acct} {m:%b %Y}: adjusted bank {st['adj_bank']:,.2f} = adjusted book, difference 0.00",
                  eq(calc.Range(f"N{r}").Value, st["adj_bank"]) and eq(calc.Range(f"O{r}").Value, st["adj_book"])
                  and calc.Range(f"P{r}").Value == 0)

        # the Summary for several picks (the workbook is read-only; nothing is saved)
        s = wb.Worksheets("Summary")
        check("Summary opens on Operating, Oct 2026", s.Range("C7").Value == "Operating" and d(s.Range("I7").Value) == date(2026, 10, 1))
        for acct, m in [("Operating", date(2026, 10, 1)), ("Payroll", date(2026, 10, 1)), ("Operating", date(2026, 7, 1)),
                        ("Payroll", date(2026, 8, 1))]:
            check_pick(wb, xl, acct, m, rows, notes)
        pick(wb, xl, "Operating", date(2026, 10, 1))
        steps = " ".join(str(s.Range(f"O{r}").Value) for r in (53, 55, 57, 59, 61))
        n1 = sum(1 for r in rows if r["rule"] == "Reference")
        check("Proof steps quote the Python numbers (files, rows, rule 1 count, adjusted bank)",
              f"Read 24 CSV files: {data['raw_bank']:,} bank rows, {data['raw_ledger']:,} cash book rows" in steps
              and f"Matched {n1:,} on reference" in steps and f"{st_py[('Operating', date(2026, 10, 1))]['adj_bank']:,.2f}" in steps
              and "8/8 PASS" in steps, steps[:200])
        wb.Close(False)

        # ------------------------------------------------------------ 4. the client's PC: another folder, Refresh All
        with tempfile.TemporaryDirectory() as tmp:
            shutil.copy(ROOT / "Two-Source-Reconciliation.xlsx", tmp)
            shutil.copytree(ROOT / "data", Path(tmp) / "data")
            wb = xl.Workbooks.Open(str(Path(tmp) / "Two-Source-Reconciliation.xlsx"))

            def refresh():
                wb.RefreshAll(); xl.CalculateUntilAsyncQueriesDone(); xl.Calculate()

            refresh()
            folder = wb.Worksheets("Settings").Range("C9").Value
            check("Refresh in another folder reads that folder's data", str(folder).lower().startswith(str(tmp).lower()[:12]), folder)
            check("Refresh in another folder still balances to 0.00", wb.Worksheets("Summary").Range("K48").Value == 0)
            got = {r["Exception ID"]: r for r in table(wb, "Exceptions workbench", "Workbench")}
            check("Refresh in another folder keeps every typed note", all(got[i]["Owner"] == n["owner"] for i, n in notes.items()))

            # the team resolves the July cash deposit: type a date, Refresh All, the Summary follows
            target = next(i for i, n in notes.items() if "deposit slip" in n["comment"])
            lo = wb.Worksheets("Exceptions workbench").ListObjects("Workbench")
            ids = [r[0] for r in lo.ListColumns("Exception ID").DataBodyRange.Value]
            col = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)].index("Resolved date") + 1
            lo.DataBodyRange.Cells(ids.index(target) + 1, col).Value = serial(date(2026, 11, 6))
            refresh()
            notes2 = dict(notes); notes2[target] = dict(notes[target], resolved="2026-11-06")
            k = check_pick(wb, xl, "Operating", date(2026, 10, 1), rows, notes2, "after resolving 1 item: ")
            check("Resolved item flows into the Summary: open exceptions down by 1, July open difference 0",
                  k["open_exceptions"] == R.month_kpis(rows, "Operating", date(2026, 10, 1), notes)["open_exceptions"] - 1
                  and wb.Worksheets("Calc").Range("F17").Value == 0)

            # tighter tolerances on Settings: the workbook must give exactly the Python result for them
            st_ = wb.Worksheets("Settings")
            st_.Range("C10").Value = 0.5
            st_.Range("C11").Value = 1
            refresh()
            rows2 = R.reconcile(data, 0.5, 1)
            rec2 = table(wb, "Reconciliation", "Recon")
            p2 = {(r["b"]["id"] if r["b"] else None, r["l"]["id"] if r["l"] else None, r["status"], r["rule"]) for r in rows2}
            x2 = {(r["Bank ID"], r["Ledger ID"], r["Status"], r["Match rule"] or "") for r in rec2}
            c2 = Counter(r["rule"] for r in rows2 if r["status"] == "Matched")
            check(f"Tolerance ±0.50 / ±1 day: same pairs as Python (rule counts {dict(c2)})", p2 == x2, f"{len(p2 ^ x2)} differ")
            check("Tolerance ±0.50 / ±1 day: every account-month still balances to 0.00",
                  wb.Names("PeriodsFailed").RefersToRange.Value == 0 and wb.Names("PeriodsChecked").RefersToRange.Value == 12)
            wb.Close(False)
    finally:
        xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
(ROOT / "test" / "verify_result.json").write_text(json.dumps(dict(passed=sum(results), total=len(results))), encoding="utf-8")
sys.exit(0 if all(results) else 1)
