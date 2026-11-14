"""Independent check of Sales-Dashboard.xlsx.

1. Recomputes every dashboard number in plain Python straight from the two CSV files (no Excel involved).
2. Opens the workbook with Excel, sets the month/region drop-downs to several combinations and compares
   the KPIs, the 12-month trend, categories, regions, top 10 customers and sales reps with the Python result.
3. Checks the planted story (West misses budget in Oct 2026 because of Garden) and that all data checks say PASS.
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs pywin32)
"""
import csv, json, sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
results = []


def check(name, ok, detail=""):
    results.append(ok)
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


def close(a, b, tol=0.011):
    return abs((a or 0) - (b or 0)) <= tol


# ---------------------------------------------------------------- 1. plain Python
sales = []
with open(ROOT / "data" / "sales-lines.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        d = date.fromisoformat(r["Date"])
        q, p, c = int(r["Qty"]), float(r["Unit Price"]), float(r["Unit Cost"])
        rev, cost = round(q * p, 2), round(q * c, 2)
        sales.append(dict(month=d.replace(day=1), inv=r["Invoice"], cust=r["Customer"], region=r["Region"],
                          rep=r["Sales Rep"], cat=r["Category"], rev=rev, margin=round(rev - cost, 2)))
budget = []
with open(ROOT / "data" / "budget.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        y, m = map(int, r["Month"].split("-"))
        budget.append(dict(month=date(y, m, 1), region=r["Region"], cat=r["Category"], budget=float(r["Budget Revenue"])))
story = json.loads((ROOT / "test" / "story.json").read_text())
MONTHS = sorted({b["month"] for b in budget})


def py(month, region):
    S = [s for s in sales if region == "All" or s["region"] == region]
    B = [b for b in budget if region == "All" or b["region"] == region]
    cur = [s for s in S if s["month"] == month]
    prev_m = date(month.year - (month.month == 1), (month.month - 2) % 12 + 1, 1)
    prev = [s for s in S if s["month"] == prev_m]
    rev = sum(s["rev"] for s in cur); bud = sum(b["budget"] for b in B if b["month"] == month)
    orders = len({s["inv"] for s in cur})
    out = dict(rev=rev, bud=bud, margin=sum(s["margin"] for s in cur) / rev, orders=orders, aov=rev / orders,
               prev_margin=sum(s["margin"] for s in prev) / sum(s["rev"] for s in prev) if prev else 0,
               ytd=sum(s["rev"] for s in S if s["month"] <= month),
               ytd_bud=sum(b["budget"] for b in B if b["month"] <= month))
    out["trend"] = [(sum(s["rev"] for s in S if s["month"] == m), sum(b["budget"] for b in B if b["month"] == m)) for m in MONTHS]
    out["cats"] = {c: (sum(s["rev"] for s in cur if s["cat"] == c), sum(b["budget"] for b in B if b["month"] == month and b["cat"] == c))
                   for c in ["Fasteners", "Power Tools", "Paint", "Garden", "Electrical"]}
    cust = defaultdict(float)
    for s in cur:
        cust[s["cust"]] += s["rev"]
    out["top"] = sorted(cust.items(), key=lambda kv: -kv[1])[:10]
    reps = defaultdict(float)
    for s in cur:
        reps[s["rep"]] += s["rev"]
    out["reps"] = dict(reps)
    return out


def py_regions(month):
    res = {}
    for reg in ["North", "South", "East", "West"]:
        res[reg] = (sum(s["rev"] for s in sales if s["month"] == month and s["region"] == reg),
                    sum(b["budget"] for b in budget if b["month"] == month and b["region"] == reg))
    return res


# ---------------------------------------------------------------- 2. Excel
xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(ROOT / "Sales-Dashboard.xlsx"), 0, True)
    d, c = wb.Worksheets("Dashboard"), wb.Worksheets("Calc")
    check("Sales table has every CSV line", c.Parent.Worksheets("Sales data").ListObjects("Sales").ListRows.Count == len(sales))
    implicit = []
    for ws in wb.Worksheets:
        try:
            cells = ws.UsedRange.SpecialCells(-4123)  # formulas
        except Exception:
            continue
        implicit += [f"{ws.Name}!{x.Address}" for a in cells.Areas for x in a.Cells if "@" in x.Formula2.replace("[@", "")]
    check("No implicit-intersection @ formulas (opens cleanly in Excel 2016)", not implicit, str(implicit[:3]))
    regs_dd = [x[0] for x in wb.Names("RegionList").RefersToRange.Value] if d.Range("G4").Validation.Formula1 == "=RegionList" else []
    check("Region drop-down offers All + 4 regions", regs_dd == ["All", "North", "South", "East", "West"], str(regs_dd))
    check("Month drop-down lists 12 months", [c.Range(f"A{r}").Value.date() for r in range(4, 16)] == MONTHS)
    for month, region in [(MONTHS[-1], "All"), (MONTHS[-1], "West"), (date(2026, 5, 1), "North"), (MONTHS[0], "All")]:
        d.Range("C4").Value2 = (month - date(1899, 12, 30)).days
        d.Range("G4").Value = region
        xl.CalculateFull()
        p = py(month, region)
        tag = f"{month:%b %Y}/{region}"
        v = lambda a: c.Range(a).Value
        check(f"{tag}: revenue", close(v("I4"), p["rev"]), f"{v('I4')} vs {p['rev']}")
        check(f"{tag}: budget", close(v("I5"), p["bud"]))
        check(f"{tag}: gross margin %", close(v("I8"), p["margin"], 1e-6))
        check(f"{tag}: margin last month", close(v("I9"), p["prev_margin"], 1e-6))
        check(f"{tag}: orders", v("I11") == p["orders"], f"{v('I11')} vs {p['orders']}")
        check(f"{tag}: average order", close(v("I12"), p["aov"]))
        check(f"{tag}: YTD revenue and budget", close(v("I13"), p["ytd"]) and close(v("I14"), p["ytd_bud"]))
        check(f"{tag}: 12-month trend", all(close(v(f"C{4 + i}"), a) and close(v(f"D{4 + i}"), b) for i, (a, b) in enumerate(p["trend"])))
        check(f"{tag}: categories", all(close(v(f"B{19 + i}"), p["cats"][v(f"A{19 + i}")][0]) and close(v(f"C{19 + i}"), p["cats"][v(f"A{19 + i}")][1])
                                         for i in range(5)))
        xl_top = [(v(f"M{r}"), v(f"N{r}")) for r in range(4, 14) if v(f"M{r}")]
        check(f"{tag}: top 10 customers", [n for n, _ in xl_top] == [n for n, _ in p["top"]] and
              all(close(a[1], b[1], 0.5) for a, b in zip(xl_top, p["top"])), f"{xl_top[:2]} vs {p['top'][:2]}")
        reps = c.ListObjects("Reps")
        xr = {reps.DataBodyRange.Cells(i, 1).Value: reps.ListColumns("Revenue").DataBodyRange.Cells(i, 1).Value for i in range(1, reps.ListRows.Count + 1)}
        check(f"{tag}: sales reps", all(close(xr.get(k, 0), val) for k, val in p["reps"].items()) and
              all(close(val, p["reps"].get(k, 0)) for k, val in xr.items()))
        shown = [x for x in (d.Range(f"D{r}").Value for r in range(39, 45)) if x not in ("", None)]
        want = sorted(p["reps"].values(), reverse=True)[:6]
        check(f"{tag}: rep leaderboard ranked, only reps with sales", len(shown) == len(want) and all(close(a, b) for a, b in zip(shown, want)),
              f"{shown} vs {want}")
        regs = py_regions(month)
        check(f"{tag}: region table", all(close(v(f"B{27 + i}"), regs[r][0]) and close(v(f"C{27 + i}"), regs[r][1])
                                           for i, r in enumerate(["North", "South", "East", "West"])))
        checks = [d.Range(f"H{r}").Value for r in range(49, 54)]
        check(f"{tag}: dashboard data checks all PASS", checks == ["PASS"] * 5, str(checks))
        check(f"{tag}: no error values on Dashboard", not any(isinstance(x, int) and x < 0 for row in d.Range("A1:U53").Value for x in row))

    # ---------------------------------------------------------------- 3. planted story
    d.Range("C4").Value2 = (date(2026, 10, 1) - date(1899, 12, 30)).days; d.Range("G4").Value = "West"; xl.CalculateFull()
    west = {c.Range(f"A{r}").Value: c.Range(f"E{r}").Value for r in range(19, 24)}
    worst = min(west, key=west.get)
    check("Story: West is below budget in Oct 2026", c.Range("I6").Value < -0.1, f"{c.Range('I6').Value:.1%}")
    check("Story: Garden is the category that explains it", worst == story["miss"]["category"] and west[worst] < -0.5, f"{worst} {west[worst]:.1%}")
    d.Range("G4").Value = "All"; xl.CalculateFull()
    others = [c.Range(f"E{r}").Value for r in range(27, 30)]
    check("Story: the other three regions are within 10% of budget", all(abs(x) < 0.1 for x in others), str(others))
    wb.Close(False)
finally:
    xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
