"""Independent check of FTE-Comparison.xlsx.

1. Re-does the whole comparison in plain Python straight from the two CSV files (no Excel involved).
2. Compares it with the answer that was planted when the data was generated (test/expected.json).
3. Opens the workbook with Excel and compares every status, count, FTE total and every PivotTable cell with Python.
4. Copies workbook + data to another folder, refreshes there, and checks the answer is the same.
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs pywin32 for steps 3-4)
"""
import csv, json, re, sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
results = []


def check(name, ok, detail=""):
    results.append(bool(ok))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


digits = lambda t: re.sub(r"\D", "", t)

# ---------------------------------------------------------------- 1. plain-Python comparison
with open(ROOT / "data" / "hr-roster-2026-10.csv", encoding="utf-8") as f:
    raw_hr = list(csv.DictReader(f))
seen, hr_all = set(), []
for r in raw_hr:
    t = tuple(r.values())
    if t in seen:
        continue
    seen.add(t)
    hr_all.append(dict(key=str(int(digits(r["Employee ID"]))), fc="FC" + digits(r["Fund Center"]),
                       area=r["Functional Area"].strip(), fte=round(int(digits(r["FTE"])) / 100, 2), status=r["Status"].strip()))
hr = [r for r in hr_all if r["status"] != "Terminated"]
terminated = {r["key"] for r in hr_all if r["status"] == "Terminated"}
areas = {a.upper(): a for a in (r["area"] for r in hr_all)}
with open(ROOT / "data" / "finance-budget-2026-10.csv", encoding="utf-8") as f:
    budget = [dict(key=str(int(digits(r["Emp No"]))) if digits(r["Emp No"]) else None, fc="FC" + digits(r["Cost Center"]),
                   area=areas.get(r["Function"].strip().upper(), r["Function"].strip()), fte=round(float(r["Budgeted FTE"]), 2))
              for r in csv.DictReader(f)]

H = {(r["key"], r["fc"]): r for r in hr}
B = {(r["key"], r["fc"]): r for r in budget if r["key"]}
check("HR: one row per person and fund center after removing duplicates", len(H) == len(hr))
rows = []  # (status, key, hr_fte, budget_fte)
for k, h in H.items():
    b = B.get(k)
    if b:
        st = ("Functional area differs" if h["area"] != b["area"] else
              "FTE differs" if round(b["fte"] - h["fte"], 2) else "Matched")
        rows.append((st, k[0], h["fte"], b["fte"]))
h_left = {k: h for k, h in H.items() if k not in B}
b_left = {k: b for k, b in B.items() if k not in H}
b_keys = Counter(k[0] for k in b_left)
for k, h in h_left.items():
    if b_keys[k[0]]:
        b = next(v for kk, v in b_left.items() if kk[0] == k[0])
        rows.append(("Fund center differs", k[0], h["fte"], b["fte"]))
    else:
        rows.append(("Only in HR roster", k[0], h["fte"], None))
h_keys = {k[0] for k in h_left}
for k, b in b_left.items():
    if k[0] not in h_keys:
        rows.append(("Only in budget", k[0], None, b["fte"]))
for b in budget:
    if not b["key"]:
        rows.append(("Only in budget", None, None, b["fte"]))

py_counts = Counter(r[0] for r in rows)
py_hr, py_bu = defaultdict(float), defaultdict(float)
for st, _, hf, bf in rows:
    py_hr[st] += hf or 0
    py_bu[st] += bf or 0

exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
check("HR file: the 2 repeated export rows are dropped", len(raw_hr) - len(hr_all) == exp["hr_duplicates_in_file"])
check("HR active FTE = planted total", round(sum(r["fte"] for r in hr), 2) == exp["hr_fte"])
check("Budget FTE = planted total", round(sum(r["fte"] for r in budget), 2) == exp["budget_fte"])

# ---------------------------------------------------------------- 2. against the planted answer
planted = Counter(c["status"] for c in exp["cases"])
for st in ["Matched", "FTE differs", "Functional area differs", "Fund center differs", "Only in HR roster", "Only in budget"]:
    check(f"Python result = planted answer: {st}", py_counts[st] == planted[st], f"{py_counts[st]} vs {planted[st]}")
py_diff = {k: round(bf - hf, 2) for st, k, hf, bf in rows if st == "FTE differs"}
for c in exp["cases"]:
    if c["status"] == "FTE differs":
        check(f"Planted FTE difference E{c['emp']:05d} = {c['diff']:+.2f}", py_diff.get(str(c["emp"])) == c["diff"])
gone = [k for st, k, hf, bf in rows if st == "Only in budget" and k in terminated]
check("Terminated-but-funded people found = planted", len(gone) == sum(1 for c in exp["cases"] if c.get("why") == "terminated"))

grid = defaultdict(lambda: [0.0, 0.0])  # (fc, area) -> [hr, budget]
for r in hr:
    grid[(r["fc"], r["area"])][0] += r["fte"]
for r in budget:
    grid[(r["fc"], r["area"])][1] += r["fte"]

# ---------------------------------------------------------------- 3. against the workbook
try:
    import win32com.client as win32
except ImportError:
    print("SKIP  workbook checks (pywin32 not installed)")
    win32 = None


def pivot_value(pt, field, fc, area):
    try:
        v = pt.GetPivotData(field, "Fund center", fc, "Functional area", area).Value
        return v or 0.0
    except Exception:
        return 0.0  # no cell = nothing booked there


def check_workbook(wb, tag=""):
    lo = wb.Worksheets("Comparison").ListObjects("Compare")
    cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
    xrows = [dict(zip(cols, r)) for r in lo.DataBodyRange.Value]
    xc = Counter(r["Status"] for r in xrows)
    for st in sorted(set(py_counts) | set(xc)):
        check(f"{tag}Workbook = Python: {st} lines", xc[st] == py_counts[st], f"{xc[st]} vs {py_counts[st]}")
        xh = round(sum(r["HR FTE"] or 0 for r in xrows if r["Status"] == st), 2)
        xb = round(sum(r["Budget FTE"] or 0 for r in xrows if r["Status"] == st), 2)
        check(f"{tag}Workbook = Python: {st} FTE", xh == round(py_hr[st], 2) and xb == round(py_bu[st], 2),
              f"HR {xh} vs {py_hr[st]:.2f}, budget {xb} vs {py_bu[st]:.2f}")
    return xrows


if win32:
    xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(ROOT / "FTE-Comparison.xlsx"), 0, True)
        xrows = check_workbook(wb)
        check("Every line to review has a 'What to do'", all(r["What to do"] for r in xrows if r["Status"] != "Matched"))
        d = wb.Worksheets("Discrepancies").ListObjects("Discrepancies").ListRows.Count
        check("Discrepancies tab = every line that is not Matched", d == len(rows) - py_counts["Matched"], f"{d}")
        pv = wb.Worksheets("Pivot comparison")
        hp, bp, dp = (pv.PivotTables(n) for n in ("HRPivot", "BudgetPivot", "DiffPivot"))
        bad = []
        for (fc, area), (h, b) in sorted(grid.items()):
            got = (pivot_value(hp, "Total FTE", fc, area), pivot_value(bp, "Total FTE", fc, area),
                   pivot_value(dp, "FTE difference", fc, area))
            if any(round(g - e, 2) for g, e in zip(got, (h, b, b - h))):
                bad.append(f"{fc}/{area}: {got} vs {(h, b, b - h)}")
        check(f"All {len(grid)} PivotTable cells (HR, budget, difference) = Python", not bad, "; ".join(bad[:3]))
        s = wb.Worksheets("Summary")
        for r in range(24, 30):
            check(f"Workbook data check: {s.Range(f'B{r}').Value}", s.Range(f"M{r}").Value == "PASS")
        wb.Close(False)

        # The client's PC: copy workbook + data to another folder, open, Refresh All
        import shutil, tempfile
        with tempfile.TemporaryDirectory() as tmp:
            shutil.copy(ROOT / "FTE-Comparison.xlsx", tmp)
            shutil.copytree(ROOT / "data", Path(tmp) / "data")
            wb = xl.Workbooks.Open(str(Path(tmp) / "FTE-Comparison.xlsx"))
            wb.RefreshAll(); xl.CalculateUntilAsyncQueriesDone(); xl.Calculate()
            folder = wb.Worksheets("Settings").Range("B4").Value
            check("Refresh in another folder reads that folder's data", str(folder).lower().startswith(str(tmp).lower()[:12]), folder)
            check_workbook(wb, "After refresh: ")
            s = wb.Worksheets("Summary")
            check("After refresh: all six data checks PASS", all(s.Range(f"M{r}").Value == "PASS" for r in range(24, 30)))
            wb.Close(False)
    finally:
        xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
