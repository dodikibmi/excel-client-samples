"""Independent check of FTE-Comparison.xlsx.

1. Re-does the whole comparison in plain Python straight from the CSV files (no Excel involved), for every month.
2. Compares it with the answer that was planted when the data was generated (test/expected.json).
3. Opens the workbook with Excel and compares every status, count, FTE total and every PivotTable cell with Python.
4. Dashboard: every KPI card and note, every bridge (waterfall) step, every heatmap cell, the donut, the trend, the
   functional-area chart and every insight sentence is recomputed in Python - for the default view and for several
   slicer picks (one fund center, an earlier month, two discrepancy types, one functional area).
5. Copies workbook + data to another folder, refreshes there, and checks the answer is the same.
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs pywin32 for steps 3-5)
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
TYPES = ["Matched", "FTE differs", "Area differs", "Fund center moved", "Only in HR", "Vacant post", "Terminated",
         "Unknown to HR"]
WF_TYPES = TYPES[1:]
STATUS_TYPE = {"Matched": "Matched", "FTE differs": "FTE differs", "Functional area differs": "Area differs",
               "Fund center differs": "Fund center moved", "Only in HR roster": "Only in HR"}


# ---------------------------------------------------------------- 1. plain-Python comparison
def compare(month):
    with open(ROOT / "data" / f"hr-roster-{month}.csv", encoding="utf-8") as f:
        raw_hr = list(csv.DictReader(f))
    seen, hr_all = set(), []
    for r in raw_hr:
        t = tuple(r.values())
        if t in seen:
            continue
        seen.add(t)
        hr_all.append(dict(key=str(int(digits(r["Employee ID"]))), fc="FC" + digits(r["Fund Center"]),
                           area=r["Functional Area"].strip(), fte=round(int(digits(r["FTE"])) / 100, 2),
                           status=r["Status"].strip()))
    hr = [r for r in hr_all if r["status"] != "Terminated"]
    terminated = {r["key"] for r in hr_all if r["status"] == "Terminated"}
    areas = {a.upper(): a for a in (r["area"] for r in hr_all)}
    with open(ROOT / "data" / f"finance-budget-{month}.csv", encoding="utf-8") as f:
        budget = [dict(key=str(int(digits(r["Emp No"]))) if digits(r["Emp No"]) else None, fc="FC" + digits(r["Cost Center"]),
                       area=areas.get(r["Function"].strip().upper(), r["Function"].strip()), fte=round(float(r["Budgeted FTE"]), 2))
                  for r in csv.DictReader(f)]
    H = {(r["key"], r["fc"]): r for r in hr}
    B = {(r["key"], r["fc"]): r for r in budget if r["key"]}
    rows = []  # dict(status, type, key, h, b)
    for k, h in H.items():
        b = B.get(k)
        if b:
            st = ("Functional area differs" if h["area"] != b["area"] else
                  "FTE differs" if round(b["fte"] - h["fte"], 2) else "Matched")
            rows.append(dict(status=st, key=k[0], h=h, b=b))
    h_left = {k: h for k, h in H.items() if k not in B}
    b_left = {k: b for k, b in B.items() if k not in H}
    b_keys = Counter(k[0] for k in b_left)
    for k, h in h_left.items():
        if b_keys[k[0]]:
            b = next(v for kk, v in b_left.items() if kk[0] == k[0])
            rows.append(dict(status="Fund center differs", key=k[0], h=h, b=b))
        else:
            rows.append(dict(status="Only in HR roster", key=k[0], h=h, b=None))
    h_keys = {k[0] for k in h_left}
    for k, b in b_left.items():
        if k[0] not in h_keys:
            rows.append(dict(status="Only in budget", key=k[0], h=None, b=b))
    for b in budget:
        if not b["key"]:
            rows.append(dict(status="Only in budget", key=None, h=None, b=b))
    for r in rows:
        if r["status"] != "Only in budget":
            r["type"] = STATUS_TYPE[r["status"]]
        else:
            r["type"] = "Vacant post" if r["key"] is None else "Terminated" if r["key"] in terminated else "Unknown to HR"
    return dict(raw_hr=raw_hr, hr_all=hr_all, hr=hr, budget=budget, terminated=terminated, rows=rows)


MONTHS = sorted(p.name[len("hr-roster-"):-4] for p in (ROOT / "data").glob("hr-roster-*.csv"))
CUR = "2026-10"
D = {m: compare(m) for m in MONTHS}
cur = D[CUR]
rows, hr, budget = cur["rows"], cur["hr"], cur["budget"]
py_counts = Counter(r["status"] for r in rows)
py_hr, py_bu = defaultdict(float), defaultdict(float)
for r in rows:
    py_hr[r["status"]] += r["h"]["fte"] if r["h"] else 0
    py_bu[r["status"]] += r["b"]["fte"] if r["b"] else 0

exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
check("HR: one row per person and fund center after removing duplicates",
      len({(r["key"], r["fc"]) for r in hr}) == len(hr))
check("HR file: the 2 repeated export rows are dropped", len(cur["raw_hr"]) - len(cur["hr_all"]) == exp["hr_duplicates_in_file"])
check("HR active FTE = planted total", round(sum(r["fte"] for r in hr), 2) == exp["hr_fte"])
check("Budget FTE = planted total", round(sum(r["fte"] for r in budget), 2) == exp["budget_fte"])

# ---------------------------------------------------------------- 2. against the planted answer
planted = Counter(c["status"] for c in exp["cases"])
for st in ["Matched", "FTE differs", "Functional area differs", "Fund center differs", "Only in HR roster", "Only in budget"]:
    check(f"Python result = planted answer: {st}", py_counts[st] == planted[st], f"{py_counts[st]} vs {planted[st]}")
py_diff = {r["key"]: round(r["b"]["fte"] - r["h"]["fte"], 2) for r in rows if r["status"] == "FTE differs"}
for c in exp["cases"]:
    if c["status"] == "FTE differs":
        check(f"Planted FTE difference E{c['emp']:05d} = {c['diff']:+.2f}", py_diff.get(str(c["emp"])) == c["diff"])
why = Counter(c.get("why") for c in exp["cases"] if c["status"] == "Only in budget")
pt = Counter(r["type"] for r in rows)
check("Only in budget splits into planted vacant / terminated / unknown",
      (pt["Vacant post"], pt["Terminated"], pt["Unknown to HR"]) == (why["vacant"], why["terminated"], why["unknown"]))
check(f"Months in the data folder = planted history {list(exp['history'])}", MONTHS == list(exp["history"]))
for m in MONTHS:
    a = [r for r in D[m]["hr"]]
    check(f"{m}: HR {exp['history'][m]['hr_fte']:.2f} / budget {exp['history'][m]['budget_fte']:.2f} FTE = planted",
          round(sum(r["fte"] for r in a), 2) == exp["history"][m]["hr_fte"]
          and round(sum(r["fte"] for r in D[m]["budget"]), 2) == exp["history"][m]["budget_fte"])

grid = defaultdict(lambda: [0.0, 0.0])  # (fc, area) -> [hr, budget], current month
for r in hr:
    grid[(r["fc"], r["area"])][0] += r["fte"]
for r in budget:
    grid[(r["fc"], r["area"])][1] += r["fte"]


# ---------------------------------------------------------------- dashboard model (what every slicer pick must show)
LONG = []  # one row per side of every compared line, every month
for m in MONTHS:
    for r in D[m]["rows"]:
        if r["h"]:
            LONG.append(dict(month=m, src="HR roster", fc=r["h"]["fc"], area=r["h"]["area"], type=r["type"], fte=r["h"]["fte"]))
        if r["b"]:
            LONG.append(dict(month=m, src="Finance budget", fc=r["b"]["fc"], area=r["b"]["area"], type=r["type"], fte=r["b"]["fte"]))
FCS = sorted({r["fc"] for r in LONG})
AREAS = sorted({r["area"] for r in LONG})


def t2(x):            # TEXT(x,"0.00")
    return f"{round(x + 0.0, 2):.2f}"


def ts(x):            # TEXT(x,"+0.00;-0.00;0.00")
    x = round(x, 2)
    return f"+{x:.2f}" if x > 0 else f"-{-x:.2f}" if x < 0 else "0.00"


def s(n):
    return str(int(n))


def model(months, fcs=None, areas=None, types=None):
    """Everything the dashboard shows for a slicer pick (months = picked months)."""
    keep = lambda r: (fcs is None or r["fc"] in fcs) and (areas is None or r["area"] in areas) and (types is None or r["type"] in types)
    rs = [r for r in LONG if r["month"] in months and keep(r)]
    hrf = lambda **k: round(sum(r["fte"] for r in rs if r["src"] == "HR roster" and all(r[a] == b for a, b in k.items())), 2)
    buf = lambda **k: round(sum(r["fte"] for r in rs if r["src"] == "Finance budget" and all(r[a] == b for a, b in k.items())), 2)
    hrn = lambda **k: sum(1 for r in rs if r["src"] == "HR roster" and all(r[a] == b for a, b in k.items()))
    bun = lambda **k: sum(1 for r in rs if r["src"] == "Finance budget" and all(r[a] == b for a, b in k.items()))
    o = {}
    sel = sorted(months)
    o["SelCount"] = len(sel)
    o["SelMonth"] = sel[0] if len(sel) == 1 else ""
    i = MONTHS.index(o["SelMonth"]) if o["SelMonth"] else -1
    o["PriorMonth"] = MONTHS[i - 1] if i > 0 else ""
    o["K_HR"], o["K_BU"] = hrf(), buf()
    o["K_GAP"] = round(o["K_BU"] - o["K_HR"], 2)
    o["K_BLINES"], o["K_BMATCH"] = bun(), bun(type="Matched")
    o["K_PCT"] = o["K_BMATCH"] / o["K_BLINES"] if o["K_BLINES"] else 0
    o["K_VAC"], o["K_VACN"] = buf(type="Vacant post"), bun(type="Vacant post")
    o["K_TERMN"], o["K_TERMF"] = bun(type="Terminated"), buf(type="Terminated")
    o["K_HRONLY"] = hrn(type="Only in HR")
    # trend: every month, same fund center / area / type picks, not the month pick
    tr = [r for r in LONG if keep(r)]
    o["trend"] = [(m, round(sum(r["fte"] for r in tr if r["month"] == m and r["src"] == "HR roster"), 2),
                   round(sum(r["fte"] for r in tr if r["month"] == m and r["src"] == "Finance budget"), 2)) for m in MONTHS]
    tg = {m: round(b - h, 2) for m, h, b in o["trend"]}
    o["P_GAP"] = tg[o["PriorMonth"]] if o["PriorMonth"] else 0
    p_hr = dict((m, h) for m, h, b in o["trend"]).get(o["PriorMonth"], 0)
    p_bu = dict((m, b) for m, h, b in o["trend"]).get(o["PriorMonth"], 0)
    o["G_FIRST"] = tg[MONTHS[0]]
    # bridge
    steps = []
    run = [o["K_BU"]]
    for t in WF_TYPES:
        st = round(hrf(type=t) - buf(type=t), 2)
        steps.append(dict(type=t, step=st, rows=max(hrn(type=t), bun(type=t))))
        run.append(round(run[-1] + st, 2))
    lo, hi = min(run), max(run)
    import math
    floor = max(0, math.floor(lo - max(1, 0.6 * (hi - lo))))
    o["steps"], o["run"], o["WF_FLOOR"], o["WF_END"] = steps, run, floor, run[-1]
    o["K_REVIEW"] = sum(x["rows"] for x in steps)
    o["base"] = [0] + [min(run[k], run[k + 1]) - floor for k in range(len(steps))] + [0]
    o["total"] = [run[0] - floor] + [0] * len(steps) + [run[-1] - floor]
    o["up"] = [0] + [max(x["step"], 0) for x in steps] + [0]
    o["down"] = [0] + [max(-x["step"], 0) for x in steps] + [0]
    o["labels"] = [f"Budget {t2(o['K_BU'])}"] + WF_TYPES + [f"HR {t2(o['K_HR'])}"]
    o["donut"] = [max(hrn(type="Matched"), o["K_BMATCH"])] + [x["rows"] for x in steps]
    # heatmap: budget minus HR; a cell exists only if some line is booked there
    heat = {}
    for fc in FCS:
        for a in AREAS:
            cell = [r for r in rs if r["fc"] == fc and r["area"] == a]
            heat[(fc, a)] = round(sum(r["fte"] if r["src"] == "Finance budget" else -r["fte"] for r in cell), 2) if cell else None
    rowt = {fc: (round(sum(v for (f, a), v in heat.items() if f == fc and v is not None), 2)
                 if any(r["fc"] == fc for r in rs) else None) for fc in FCS}
    o["heat"], o["rowt"] = heat, rowt
    o["areas"] = [(a, hrf(area=a), buf(area=a), hrn(area=a)) for a in AREAS]
    # insights (same wording as the formulas on Dash data)
    nums = [v for v in rowt.values() if v is not None]
    mn, mx = (min(nums), max(nums)) if nums else (0, 0)
    lines = f"{o['K_REVIEW']} {'line' if o['K_REVIEW'] == 1 else 'lines'}"
    if o["SelCount"] != 1:
        i1 = f"Pick one month in the Month slicer (now {o['SelCount']} months are added together)"
    elif o["K_GAP"] > 0:
        i1 = f"Budget is {t2(o['K_GAP'])} FTE above the HR roster; {lines} explain it"
    elif o["K_GAP"] < 0:
        i1 = f"HR roster is {t2(-o['K_GAP'])} FTE above budget; {lines} explain it"
    else:
        i1 = f"HR roster and budget agree on FTE; {lines} to review"

    def mostly(fc, pick):
        vals = [(heat[(fc, a)], a) for a in AREAS if heat[(fc, a)] is not None]
        target = pick(v for v, a in vals)
        return next(a for v, a in vals if v == target)
    if mn < 0:
        fc = next(f for f in FCS if rowt[f] == mn)
        i2 = f"{fc} is {t2(-mn)} FTE over budget, mostly {mostly(fc, min)}"
    else:
        i2 = "No fund center has more people than budget"
    if mx > 0:
        fc = next(f for f in FCS if rowt[f] == mx)
        i3 = f"{fc} has {t2(mx)} budgeted FTE not filled, mostly {mostly(fc, max)}"
    else:
        i3 = "No fund center has budget left unfilled"
    big = max(abs(x["step"]) for x in steps)
    if big == 0:
        i4 = "The discrepancies net to zero FTE in this view"
    else:
        d = next(x for x in steps if abs(x["step"]) == big)
        i4 = (f"Biggest driver: {d['type']}, {d['rows']} {'line' if d['rows'] == 1 else 'lines'}, "
              f"HR {ts(d['step'])} FTE vs budget")
    if o["SelCount"] != 1:
        i5 = "Trend: pick one month"
    elif o["SelMonth"] == MONTHS[0]:
        i5 = f"{o['SelMonth']} is the first month in the data"
    else:
        g0 = o["G_FIRST"]
        trend = ": narrowing" if abs(o["K_GAP"]) < abs(g0) else ": widening" if abs(o["K_GAP"]) > abs(g0) else ": flat"
        i5 = f"Gap went from {ts(g0)} in {MONTHS[0]} to {ts(o['K_GAP'])} in {o['SelMonth']}{trend}"
    todo = ""
    if o["K_TERMN"]:
        todo += f"remove {o['K_TERMN']} terminated {'person' if o['K_TERMN'] == 1 else 'people'} from the budget; "
    if o["K_VACN"]:
        todo += f"decide on {o['K_VACN']} vacant {'post' if o['K_VACN'] == 1 else 'posts'}; "
    if o["K_HRONLY"]:
        todo += f"fund or check {o['K_HRONLY']} HR-only {'hire' if o['K_HRONLY'] == 1 else 'hires'}; "
    i6 = "To do: " + todo[:-2] if todo else "Nothing to fix in this view"
    o["insights"] = [i1, i2, i3, i4, i5, i6]
    # card notes
    arrow = lambda d: "▲ " if round(d, 2) > 0 else "▼ " if round(d, 2) < 0 else ""
    pm = o["PriorMonth"]
    good = abs(o["K_GAP"]) <= abs(o["P_GAP"])
    move = "good" if good else "bad"
    o["notes"] = [
        ("first month in the data", "muted") if not pm else (f"{arrow(o['K_HR'] - p_hr)}{ts(o['K_HR'] - p_hr)} vs {pm}", move),
        ("first month in the data", "muted") if not pm else (f"{arrow(o['K_BU'] - p_bu)}{ts(o['K_BU'] - p_bu)} vs {pm}", move),
        ("budget minus HR", "muted") if not pm else (
            ("▼ narrowed from " if abs(o["K_GAP"]) < abs(o["P_GAP"]) else "▲ widened from " if abs(o["K_GAP"]) > abs(o["P_GAP"])
             else "same as ") + f"{ts(o['P_GAP'])} in {pm}", move),
        (f"{o['K_BLINES'] - o['K_BMATCH']} of {o['K_BLINES']} budget lines to fix", "good" if o["K_PCT"] >= 0.95 else "bad"),
        ("no vacant posts", "good") if not o["K_VACN"] else (f"{o['K_VACN']} {'post' if o['K_VACN'] == 1 else 'posts'}: recruit or release", "bad"),
        ("none: the budget is clean", "good") if not o["K_TERMN"] else (f"{t2(o['K_TERMF'])} FTE still funded: remove", "bad"),
    ]
    o["cards"] = [o["K_HR"], o["K_BU"], o["K_GAP"], o["K_PCT"], o["K_VAC"], o["K_TERMN"]]
    return o


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
    xt = Counter(r["Discrepancy type"] for r in xrows)
    check(f"{tag}Workbook = Python: lines per discrepancy type (8 types)", xt == Counter(r["type"] for r in rows),
          f"{dict(xt)}")
    return xrows


def eq(a, b):
    if isinstance(a, (int, float)) or isinstance(b, (int, float)):
        try:
            return round(float(a or 0) - float(b or 0), 6) == 0
        except (TypeError, ValueError):
            return False
    return (a or "") == (b or "")


def check_dashboard(wb, o, tag):
    dd = wb.Worksheets("Dash data")
    nv = lambda nm: wb.Names(nm).RefersToRange.Value
    col = lambda nm: [c[0] if isinstance(c, tuple) else c for c in wb.Names(nm).RefersToRange.Value]
    # cards: value + note + colour
    bad = [f"card {k + 1}: {nv(f'CARD{k + 1}')} vs {o['cards'][k]}" for k in range(6) if not eq(nv(f"CARD{k + 1}"), o["cards"][k])]
    check(f"{tag}6 KPI cards = Python ({t2(o['K_HR'])} / {t2(o['K_BU'])} / {ts(o['K_GAP'])} / {o['K_PCT']:.1%} / "
          f"{t2(o['K_VAC'])} / {o['K_TERMN']})", not bad, "; ".join(bad))
    bad = [f"card {k + 1}: {nv(f'NOTE{k + 1}')!r}/{nv(f'CTONE{k + 1}')} vs {o['notes'][k]}" for k in range(6)
           if (nv(f"NOTE{k + 1}"), nv(f"CTONE{k + 1}")) != o["notes"][k]]
    check(f"{tag}6 card notes and their green/red colour = Python", not bad, "; ".join(bad))
    # bridge
    got = [col("WF_Labels"), col("WF_Base"), col("WF_Total"), col("WF_Up"), col("WF_Down")]
    want = [o["labels"], o["base"], o["total"], o["up"], o["down"]]
    bad = [f"{n}[{i}] {g} vs {w}" for n, gs, ws in zip(("label", "base", "total", "up", "down"), got, want)
           for i, (g, w) in enumerate(zip(gs, ws)) if not eq(g, w)]
    check(f"{tag}Bridge: {len(o['labels'])} bars (labels, base, totals, steps) = Python; starts at {o['WF_FLOOR']}",
          not bad and eq(nv("WF_FLOOR"), o["WF_FLOOR"]) and eq(nv("WF_END"), o["WF_END"]), "; ".join(bad[:4]))
    # heatmap on the Dashboard itself (6 fund centers x 4 areas) + row totals
    ws = wb.Worksheets("Dashboard")
    bad = []
    for k in range(6):
        fc = str(ws.Cells(21 + k, 30).Value).strip()
        if eq(ws.Cells(21 + k, 50).Value, o["rowt"].get(fc)) is False:
            bad.append(f"{fc} total {ws.Cells(21 + k, 50).Value} vs {o['rowt'].get(fc)}")
        for j in range(4):
            a = ws.Cells(20, 34 + 4 * j).Value
            g, w = ws.Cells(21 + k, 34 + 4 * j).Value, o["heat"].get((fc, a))
            if not eq(g, w):
                bad.append(f"{fc}/{a}: {g!r} vs {w}")
    labels_ok = [str(ws.Cells(21 + k, 30).Value).strip() for k in range(6)] == FCS[:6] and \
                [ws.Cells(20, 34 + 4 * j).Value for j in range(4)] == AREAS[:4]
    check(f"{tag}Heatmap: 24 cells + 6 fund center totals = Python", not bad and labels_ok, "; ".join(bad[:4]))
    # donut
    got = col("DonutValues")
    check(f"{tag}Donut: lines by type {o['donut']} = Python", all(eq(g, w) for g, w in zip(got, o["donut"])), f"{got}")
    check(f"{tag}Donut centre = {o['K_PCT']:.0%}", nv("DonutCenter") == f"{round(o['K_PCT'] * 100 + 1e-9):.0f}%", nv("DonutCenter"))
    # trend
    got = list(zip(col("TrendMonths"), col("TrendHR"), col("TrendBU"), col("TrendGap")))
    want = [(m, h, b, round(b - h, 2)) for m, h, b in o["trend"]]
    check(f"{tag}Trend: {len(want)} months HR / budget / gap = Python " + ", ".join(f"{m} {ts(g)}" for m, _, _, g in want),
          len(got) == len(want) and all(eq(x, y) for g, w in zip(got, want) for x, y in zip(g, w)), f"{got}")
    # functional areas
    got = list(zip(col("AreaLabels"), col("AreaHR"), col("AreaBU")))
    want = [(f"{a} · {n} HR lines", h, b) for a, h, b, n in o["areas"]]
    check(f"{tag}FTE by functional area chart: {len(want)} areas = Python",
          len(got) == len(want) and all(eq(x, y) for g, w in zip(got, want) for x, y in zip(g, w)), f"{got}")
    # insights
    for k in range(6):
        g = nv(f"INSIGHT{k + 1}")
        check(f"{tag}Insight {k + 1}: \"{o['insights'][k]}\"", g == o["insights"][k], f"workbook: {g!r}")


def pick(wb, field, items):
    """Slicer pick: exactly these items selected (None = all)."""
    sc = wb.SlicerCaches("Slicer_" + field.replace(" ", "_"))
    sc.ClearManualFilter()
    if items is not None:
        for it in sc.SlicerItems:
            if it.Name not in items:
                it.Selected = False
    wb.Application.Calculate()


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
        ck = wb.Worksheets("Checks")
        n_checks = int(wb.Names("CHK_N").RefersTo.lstrip("="))
        for r in range(19, 19 + n_checks):
            check(f"Workbook data check: {ck.Range(f'B{r}').Value}", ck.Range(f"H{r}").Value == "PASS")

        # slicers: every one drives the three pivots and the dashboard pivots; the trend ignores Month
        main = {"HRPivot", "BudgetPivot", "DiffPivot", "PK", "PA"}
        want = {"Month": main | {"PM"}, "Fund center": main | {"PT"}, "Functional area": main | {"PT"},
                "Discrepancy type": main | {"PT"}}
        for field, pts in want.items():
            sc = wb.SlicerCaches("Slicer_" + field.replace(" ", "_"))
            got = {p.Name for p in sc.PivotTables}
            check(f"Slicer '{field}' drives {', '.join(sorted(pts))} (and sits on the Dashboard and the Pivots tab)",
                  got == pts and sc.Slicers.Count == 2, f"{sorted(got)}, {sc.Slicers.Count} slicers")
        check("Month slicer starts on the latest month", [i.Name for i in wb.SlicerCaches("Slicer_Month").SlicerItems if i.Selected] == [MONTHS[-1]])

        # the dashboard, default view and several slicer picks
        check_dashboard(wb, model([CUR]), "Default view: ")
        o = model([CUR])
        fc_over = o["insights"][1].split(" ")[0]
        pick(wb, "Fund center", [fc_over])
        check_dashboard(wb, model([CUR], fcs={fc_over}), f"Fund center = {fc_over}: ")
        pick(wb, "Fund center", None)
        pick(wb, "Month", [MONTHS[-2]])
        check_dashboard(wb, model([MONTHS[-2]]), f"Month = {MONTHS[-2]}: ")
        pick(wb, "Month", [CUR])
        pick(wb, "Discrepancy type", ["Vacant post", "Terminated"])
        check_dashboard(wb, model([CUR], types={"Vacant post", "Terminated"}), "Type = Vacant post + Terminated: ")
        pick(wb, "Discrepancy type", None)
        pick(wb, "Functional area", ["Clinical Care"])
        check_dashboard(wb, model([CUR], areas={"Clinical Care"}), "Area = Clinical Care: ")
        pick(wb, "Functional area", None)
        pick(wb, "Month", MONTHS[-2:])
        o2 = model(MONTHS[-2:])
        check("Two months picked: insight asks to pick one month", wb.Names("INSIGHT1").RefersToRange.Value == o2["insights"][0],
              wb.Names("INSIGHT1").RefersToRange.Value)
        pick(wb, "Month", [CUR])
        check("Data checks still PASS after all the slicer picks",
              all(ck.Range(f"H{r}").Value == "PASS" for r in range(19, 19 + n_checks)))
        wb.Close(False)

        # The client's PC: copy workbook + data to another folder, open, Refresh All
        import shutil, tempfile
        with tempfile.TemporaryDirectory() as tmp:
            shutil.copy(ROOT / "FTE-Comparison.xlsx", tmp)
            shutil.copytree(ROOT / "data", Path(tmp) / "data")
            wb = xl.Workbooks.Open(str(Path(tmp) / "FTE-Comparison.xlsx"))
            wb.RefreshAll(); xl.CalculateUntilAsyncQueriesDone(); xl.Calculate()
            folder = wb.Worksheets("Settings").Range("C8").Value
            check("Refresh in another folder reads that folder's data", str(folder).lower().startswith(str(tmp).lower()[:12]), folder)
            check_workbook(wb, "After refresh: ")
            check_dashboard(wb, model([CUR]), "After refresh: ")
            ck = wb.Worksheets("Checks")
            check(f"After refresh: all {n_checks} data checks PASS",
                  all(ck.Range(f"H{r}").Value == "PASS" for r in range(19, 19 + n_checks)))
            wb.Close(False)
    finally:
        xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
