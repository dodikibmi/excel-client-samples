"""Independent check of Project-Cost-Tracker.xlsx.

1. Recomputes every project in plain Python from what the PMs typed (data/*.csv), no Excel involved, and compares it
   with the answer planted when the data was generated (test/expected.json).
2. Opens the workbook with Excel and compares every PM sheet line, every Dashboard figure, the History table and the
   Weekly trend grid with Python; all seven Dashboard checks must say PASS.
3. On a copy in another folder, plays a few weeks of real use:
   - Refresh All twice in the same week: History does not grow (a week is never saved twice)
   - next week: a PM updates a line, Week ending + 7, Refresh All: 3 new rows, the earlier weeks unchanged
   - a project sheet copied between the bookends but not added to the Dashboard: the check says FAIL, then PASS once added
   - an unknown cost code typed on a PM sheet: the check says FAIL
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs pywin32 for steps 2-3)
"""
import csv, json, shutil, sys, tempfile
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BOOK = "Project-Cost-Tracker.xlsx"
results = []


def check(name, ok, detail=""):
    results.append(bool(ok))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


def close(a, b, tol=0.005):
    return a is not None and b is not None and abs(float(a) - float(b)) <= tol


def read(name):
    with open(ROOT / "data" / name, encoding="utf-8") as f:
        return list(csv.DictReader(f))


# ---------------------------------------------------------------- 1. plain Python
exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
week = date.fromisoformat(exp["week_ending"])
projects = read("projects.csv")
codes = {r["Cost code"]: r["Category"] for r in read("cost-codes.csv")}
lines = defaultdict(list)
for r in read(f"pm-week-{week.isoformat()}.csv"):
    lines[r["sheet"]].append(dict(code=r["code"], budget=float(r["budget"]), committed=float(r["committed"]),
                                  actual=float(r["actual"]), pct=float(r["pct"])))


def line_calc(l):
    ev = l["budget"] * l["pct"]
    fac = l["actual"] / l["pct"] if l["pct"] > 0 else max(l["budget"], l["actual"])
    return ev, fac


py = {}
for p in projects:
    ls = lines[p["Sheet"]]
    ev = sum(line_calc(l)[0] for l in ls)
    fac = sum(line_calc(l)[1] for l in ls)
    actual = sum(l["actual"] for l in ls)
    field = [l for l in ls if codes[l["code"]] != "Shop fabrication"]
    shop = [l for l in ls if codes[l["code"]] == "Shop fabrication"]
    cpi = ev / actual
    py[p["Sheet"]] = dict(contract=float(p["Contract value"]), budget=sum(l["budget"] for l in ls),
                          committed=sum(l["committed"] for l in ls), actual=actual, earned=ev, fac=fac, cpi=cpi,
                          field=sum(l["budget"] * l["pct"] for l in field) / sum(l["budget"] for l in field),
                          shop=sum(l["budget"] * l["pct"] for l in shop) / sum(l["budget"] for l in shop),
                          status="On track" if cpi >= 0.98 else "Watch" if cpi >= 0.92 else "At risk")
check("Every cost code typed by the PMs is in the cost code list", all(l["code"] in codes for ls in lines.values() for l in ls))
for s, v in py.items():
    e = exp["projects"][s]
    check(f"{s}: Python = planted answer (budget, actual, earned, forecast, CPI)",
          all(close(v[k], e[k], 0.01) for k in ("budget", "actual", "earned", "fac")) and close(v["cpi"], e["cpi"], 1e-6))
    check(f"{s}: status = planted story ({exp['planted_status'][s]})", v["status"] == exp["planted_status"][s])
tot = {k: sum(v[k] for v in py.values()) for k in ("contract", "budget", "committed", "actual", "earned", "fac")}
check("Portfolio totals = planted answer", all(close(tot[k], exp["portfolio"][k], 0.01) for k in tot))
seed = read("history-seed.csv")
check("History seed: 7 earlier weeks x 3 projects", len(seed) == 21 and len({r["week"] for r in seed}) == 7)

# ---------------------------------------------------------------- 2. the workbook
try:
    import win32com.client as win32
except ImportError:
    print("SKIP  workbook checks (pywin32 not installed)")
    win32 = None

D_FIRST, CHECKS = 13, range(25, 32)


def serial(d):
    return (d - date(1899, 12, 30)).days


def history(wb):
    lo = wb.Worksheets("Snapshots").ListObjects("History")
    cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
    out = []
    for row in lo.DataBodyRange.Value2:
        out.append(dict(zip(cols, row)))
    return out


def checks(wb):
    d = wb.Worksheets("Dashboard")
    return [d.Range(f"P{r}").Value for r in CHECKS]


def check_workbook(wb, tag=""):
    d = wb.Worksheets("Dashboard")
    for i, (s, v) in enumerate(py.items()):
        r = D_FIRST + i
        got = dict(contract=d.Range(f"E{r}").Value, budget=d.Range(f"F{r}").Value, committed=d.Range(f"G{r}").Value,
                   actual=d.Range(f"H{r}").Value, earned=d.Range(f"I{r}").Value, cpi=d.Range(f"J{r}").Value,
                   field=d.Range(f"K{r}").Value, shop=d.Range(f"L{r}").Value, fac=d.Range(f"M{r}").Value)
        bad = [k for k in got if not close(got[k], v[k], 1e-6 if k in ("cpi", "field", "shop") else 0.01)]
        check(f"{tag}Dashboard {s}: all 9 figures = Python", d.Range(f"B{r}").Value == s and not bad, f"{bad}")
        check(f"{tag}Dashboard {s}: status {v['status']}", d.Range(f"P{r}").Value == v["status"])
    t = D_FIRST + 8
    check(f"{tag}Dashboard totals = Python", all(close(d.Range(f"{c}{t}").Value, tot[k], 0.01) for c, k in
                                              (("E", "contract"), ("F", "budget"), ("G", "committed"), ("H", "actual"), ("I", "earned"), ("M", "fac"))))
    check(f"{tag}Variance at completion card = budget - forecast", close(d.Range("N6").Value, tot["budget"] - tot["fac"], 0.01))
    check(f"{tag}All seven Dashboard checks PASS", checks(wb) == ["PASS"] * 7, f"{checks(wb)}")


if win32:
    xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(ROOT / BOOK), 0, True)
        # every PM sheet line
        bad = []
        for s in py:
            ws = wb.Worksheets(s)
            for i, l in enumerate(lines[s]):
                r = 15 + i
                ev, fac = line_calc(l)
                if ws.Range(f"B{r}").Value != l["code"] or not close(ws.Range(f"I{r}").Value, ev, 0.01) or not close(ws.Range(f"L{r}").Value, fac, 0.01):
                    bad.append(f"{s} {l['code']}")
        check("All 30 PM sheet lines: earned value and forecast at completion = Python", not bad, ", ".join(bad[:3]))
        check_workbook(wb)
        h = history(wb)
        check("History: 8 weeks x 3 projects = 24 rows", len(h) == 24 and len({r["Week ending"] for r in h}) == 8)
        sh = {(int(r["Week ending"]), r["Sheet"]): r for r in h}
        okseed = all(close(sh[(serial(date.fromisoformat(r["week"])), r["sheet"])]["Actual"], float(r["actual"]), 0.01) and
                     close(sh[(serial(date.fromisoformat(r["week"])), r["sheet"])]["CPI"], float(r["cpi"]), 1e-6) for r in seed)
        check("History: the seven earlier weeks are stored as generated", okseed)
        okw = all(close(sh[(serial(week), s)][k], py[s][kk], 1e-6 if k == "CPI" else 0.01)
                  for s in py for k, kk in (("Actual", "actual"), ("Earned", "earned"), ("CPI", "cpi"), ("Forecast at completion", "fac")))
        check("History: this week's snapshot = Python", okw)
        tr = wb.Worksheets("Weekly trend")
        check("Weekly trend: last week's CPI per project = Python",
              all(close(tr.Cells(13, 3 + i).Value, py[s]["cpi"], 1e-6) for i, s in enumerate(py)))
        check("Weekly trend: first week = 25 Sep 2026", int(tr.Range("B6").Value2) == serial(week - timedelta(days=49)))
        wb.Close(False)

        # ---------------------------------------------------------------- 3. weeks of real use, on a copy
        with tempfile.TemporaryDirectory() as tmp:
            shutil.copy(ROOT / BOOK, tmp)
            wb = xl.Workbooks.Open(str(Path(tmp) / BOOK))

            print(f"      (Excel calculation mode: {xl.Calculation}; -4105 = automatic)")

            def refresh():
                xl.Calculate()
                wb.RefreshAll(); xl.CalculateUntilAsyncQueriesDone(); xl.Calculate()

            refresh(); refresh()
            check("Same week, Refresh All twice: History still 24 rows", len(history(wb)) == 24)
            check_workbook(wb, "After refresh: ")

            before = {(int(r["Week ending"]), r["Sheet"]): r["Actual"] for r in history(wb)}
            st = wb.Worksheets("Settings")
            st.Range("B5").Value2 = serial(week + timedelta(days=7))
            pm = wb.Worksheets("P-2401")
            old_g = pm.Range("G16").Value
            pm.Range("G16").Value = old_g + 6000   # trim and finish: this week's costs
            pm.Range("H16").Value = 0.40
            refresh()
            h = history(wb)
            new = [r for r in h if int(r["Week ending"]) == serial(week + timedelta(days=7))]
            check("Next week + Refresh All: 3 new rows (27)", len(h) == 27 and len(new) == 3, f"{len(h)}")
            p1 = next(r for r in new if r["Sheet"] == "P-2401")
            check("Next week: P-2401 snapshot has the PM's update", close(p1["Actual"], py["P-2401"]["actual"] + 6000, 0.01), f"{p1['Actual']} vs {py['P-2401']['actual'] + 6000}")
            after = {(int(r["Week ending"]), r["Sheet"]): r["Actual"] for r in h}
            check("Next week: the earlier 24 rows are unchanged", all(close(after[k], v, 0.01) for k, v in before.items()))
            check("Next week: Weekly trend moves on (last row = new week)",
                  int(wb.Worksheets("Weekly trend").Range("B13").Value2) == serial(week + timedelta(days=7)))
            check("Next week: all checks PASS", checks(wb) == ["PASS"] * 7, f"{checks(wb)}")

            wb.Worksheets("Template").Copy(Before=wb.Worksheets("PM end"))
            xl.Calculate()
            copy = wb.Worksheets(wb.Worksheets("PM end").Index - 1)
            copy.Name = "P-2404"
            xl.Calculate()
            check("Sheet copied in but not on the Dashboard: check 1 says FAIL", checks(wb)[0] == "FAIL")
            d = wb.Worksheets("Dashboard")
            d.Range("B16").Value = "P-2404"
            xl.Calculate()
            check("Added to the Dashboard: check 1 PASS again", checks(wb)[0] == "PASS")
            check("New project row reads the copied sheet", d.Range("C16").Value == "New project name")
            copy.Range("E15").Value = 50000
            xl.Calculate()
            check("Budget typed on the new sheet reaches the Dashboard total, checks still PASS",
                  close(d.Range("F21").Value, tot["budget"] + 50000, 0.01) and checks(wb)[1] == "PASS")
            wb.Worksheets("P-2403").Range("B15").Value = "26-999"
            xl.Calculate()
            check("Unknown cost code typed: check 4 says FAIL", checks(wb)[3] == "FAIL")
            wb.Close(False)
    finally:
        xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
