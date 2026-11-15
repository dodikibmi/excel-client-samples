"""Independent check of Sales-Dashboard.xlsx.

1. Recomputes every number on the six report pages in plain Python straight from the two CSV files (no Excel).
2. Opens the workbook with Excel, drives the real slicers and the Month timeline through six selections
   (month / YTD view, one region, two categories, two reps, an early month with no prior year) and compares:
   KPI matrix and changes, 12-month trend + heatmap, region and category tables, both waterfalls, the
   region x category matrix, customers ranked + Pareto (80/20), rep scorecard + gauges, forecast to year end,
   the six key-insight sentences word for word, the KPI card notes, and that every data check says PASS.
3. Checks the planted story (West misses budget in Oct 2026 because of Garden), the page furniture (slicers,
   timeline, navigation buttons, sparklines), and that no formula needs a Microsoft 365-only function.
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs pywin32)
"""
import csv, json, sys
from collections import defaultdict
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
results = []


def check(name, ok, detail=""):
    results.append(bool(ok))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


def close(a, b, tol=0.011):
    if a in ("", None) or b in ("", None):
        return a in ("", None) and b in ("", None)
    return abs(a - b) <= tol


def closer(a, b):
    return close(a, b, 1e-7)


# ---------------------------------------------------------------- Excel-style number formatting (TEXT)
def rnd(x, nd=0):
    q = Decimal(1).scaleb(-nd)
    return Decimal(repr(round(x, 9))).quantize(q, rounding=ROUND_HALF_UP)


def t_num(x, nd=0):                     # "#,##0" / "#,##0.0"
    d = rnd(abs(x), nd)
    s = f"{d:,.{nd}f}"
    return ("-" if x < 0 and d != 0 else "") + s


def t_signed(x, nd=0):                  # "+#,##0;-#,##0"
    return ("+" if x >= 0 else "-") + f"{rnd(abs(x), nd):,.{nd}f}"


def t_pct(x, nd=1, sign=False):         # "0.0%" or "+0.0%;-0.0%"
    s = f"{rnd(abs(x) * 100, nd):,.{nd}f}%"
    if sign:
        return ("+" if x >= 0 else "-") + s
    return ("-" if x < 0 and rnd(abs(x) * 100, nd) != 0 else "") + s


def t_pts(x):                           # TEXT(x*100,"+0.0;-0.0")
    return ("+" if x >= 0 else "-") + f"{rnd(abs(x) * 100, 1):.1f}"


# ---------------------------------------------------------------- 1. plain Python model
CATS = ["Fasteners", "Power Tools", "Paint", "Garden", "Electrical"]
REGIONS = ["North", "South", "East", "West"]
sales, inv_cats, inv_first = [], defaultdict(set), {}
with open(ROOT / "data" / "sales-lines.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        d = date.fromisoformat(r["Date"])
        q, p, c = int(r["Qty"]), float(r["Unit Price"]), float(r["Unit Cost"])
        rev, cost = round(q * p, 2), round(q * c, 2)
        s = dict(month=d.replace(day=1), inv=r["Invoice"], cust=r["Customer"], region=r["Region"], rep=r["Sales Rep"],
                 cat=r["Category"], rev=rev, margin=round(rev - cost, 2))
        sales.append(s)
        inv_cats[s["inv"]].add(s["cat"])
        inv_first.setdefault(s["inv"], s)
invoices = list(inv_first.values())       # one row per invoice (month, customer, region, rep are invoice-level)
budget = []
with open(ROOT / "data" / "budget.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        y, m = map(int, r["Month"].split("-"))
        b = float(r["Budget Revenue"])
        budget.append(dict(month=date(y, m, 1), region=r["Region"], rep=r["Sales Rep"], cat=r["Category"], budget=b,
                           gp=round(b * float(r["Budget Margin %"]), 2)))
story = json.loads((ROOT / "test" / "story.json").read_text())
CUSTOMERS = sorted({s["cust"] for s in sales})
REPS = sorted({(s["region"], s["rep"]) for s in sales})            # Power Query sorts by region, then rep
REPS = [rep for _, rep in REPS]
REP_REGION = {s["rep"]: s["region"] for s in sales}
CUST_REGION = {s["cust"]: s["region"] for s in sales}
MONTHS = sorted({s["month"] for s in sales})


def edate(d, k):
    m = d.year * 12 + d.month - 1 + k
    return date(m // 12, m % 12 + 1, 1)


def model(sel):
    R, C, P_ = set(sel["regions"]), set(sel["cats"]), set(sel["reps"])
    rep_m, view = sel["month"], sel["view"]
    S = [s for s in sales if s["region"] in R and s["rep"] in P_ and s["cat"] in C]
    O = [i for i in invoices if i["region"] in R and i["rep"] in P_ and inv_cats[i["inv"]] & C]
    B = [b for b in budget if b["region"] in R and b["rep"] in P_ and b["cat"] in C]
    ps = date(rep_m.year, 1, 1) if view == "YTD" else rep_m
    pe = rep_m
    pys, pye, prev = edate(ps, -12), edate(pe, -12), edate(rep_m, -1)

    def rev(a, b, key="rev", **flt):
        return sum(s[key] for s in S if a <= s["month"] <= b and all(s[k] == v for k, v in flt.items()))

    def orders(a, b, **flt):
        return sum(1 for i in O if a <= i["month"] <= b and all(i[k] == v for k, v in flt.items()))

    def bud(a, b, key="budget", **flt):
        return sum(x[key] for x in B if a <= x["month"] <= b and all(x[k] == v for k, v in flt.items()))

    out = {}
    cols = [(ps, pe), (pys, pye), (rep_m, rep_m), (prev, prev)]
    k = {}
    for j, (a, b) in enumerate(cols):
        k[j] = dict(rev=rev(a, b), gp=rev(a, b, "margin"), ord=orders(a, b), bud=bud(a, b), bgp=bud(a, b, "gp"))
        kk = k[j]
        kk["mg"] = kk["gp"] / kk["rev"] if kk["rev"] else 0
        kk["bmg"] = kk["bgp"] / kk["bud"] if kk["bud"] else 0
        kk["aov"] = kk["rev"] / kk["ord"] if kk["ord"] else 0
        kk["ach"] = kk["rev"] / kk["bud"] if kk["bud"] else 0
    out["kpi"] = [[k[j][m] for j in range(4)] for m in ("rev", "gp", "ord", "bud", "bgp", "mg", "bmg", "aov", "ach")]
    P, PY, M, PM = k[0], k[1], k[2], k[3]
    ratio = lambda a, b: "" if b == 0 else a / b - 1
    out["delta"] = [
        [ratio(P["rev"], P["bud"]), ratio(P["rev"], PY["rev"]), ratio(M["rev"], PM["rev"])],
        [ratio(P["gp"], P["bgp"]), ratio(P["gp"], PY["gp"]), ratio(M["gp"], PM["gp"])],
        ["" if P["rev"] == 0 or P["bud"] == 0 else P["mg"] - P["bmg"], "" if P["rev"] == 0 or PY["rev"] == 0 else P["mg"] - PY["mg"],
         "" if M["rev"] == 0 or PM["rev"] == 0 else M["mg"] - PM["mg"]],
        ["", ratio(P["ord"], PY["ord"]), ratio(M["ord"], PM["ord"])],
        ["", ratio(P["aov"], PY["aov"]), ratio(M["aov"], PM["aov"])],
        ["", "" if PY["bud"] == 0 or PY["rev"] == 0 else P["ach"] - PY["ach"],
         "" if PM["bud"] == 0 or PM["rev"] == 0 else M["ach"] - PM["ach"]]]
    # 12-month trend
    tr = []
    for i in range(12):
        m = edate(rep_m, i - 11)
        a, b, py = rev(m, m), bud(m, m), rev(edate(m, -12), edate(m, -12))
        gp, o = rev(m, m, "margin"), orders(m, m)
        row = [a, b, py, gp, gp / a if a else 0, o, a / o if o else 0, a / b if b else 0]
        ra = [rev(m, m, region=r) for r in REGIONS]
        rb = [bud(m, m, region=r) for r in REGIONS]
        row += ra + rb + ["" if y == 0 else x / y - 1 for x, y in zip(ra, rb)] + [rev(m, m, cat=c) for c in CATS]
        tr.append(row)
    out["trend"] = tr
    out["trend_months"] = [edate(rep_m, i - 11) for i in range(12)]

    def block(key, items):
        rows = []
        for it in items:
            a, b = rev(ps, pe, **{key: it}), bud(ps, pe, **{key: it})
            py, gp = rev(pys, pye, **{key: it}), rev(ps, pe, "margin", **{key: it})
            bg = bud(ps, pe, "gp", **{key: it})
            rows.append([a, b, a - b, "" if b == 0 else a / b - 1, py, "" if py == 0 else a / py - 1, gp, "" if a == 0 else gp / a,
                         "" if b == 0 else bg / b, bg])
        t = [sum(r[i] for r in rows) for i in (0, 1, 4, 6, 9)]
        a, b, py, gp, bg = t
        rows.append([a, b, a - b, "" if b == 0 else a / b - 1, py, "" if py == 0 else a / py - 1, gp, "" if a == 0 else gp / a,
                     "" if b == 0 else bg / b, bg])
        for r in rows:
            r.append("" if rows[-1][0] == 0 else r[0] / rows[-1][0])
        return rows
    out["region"] = block("region", REGIONS)
    out["category"] = block("cat", CATS)

    def waterfall(rows):
        steps = [rows[-1][1]] + [r[2] for r in rows[:-1]] + [rows[-1][0]]
        cum = [steps[0]]
        for d in steps[1:-1]:
            cum.append(cum[-1] + d)
        cum.append(steps[-1])
        mn = min(cum) * 0.95
        floor = 0.0 if mn <= 0 else float(Decimal(repr(mn)).quantize(Decimal(1).scaleb(len(str(int(mn))) - 2), rounding="ROUND_DOWN"))
        floor = max(0.0, floor)
        base = [0] + [min(cum[i - 1], cum[i]) - floor for i in range(1, len(cum) - 1)] + [0]
        tot = [cum[0] - floor] + [0] * (len(cum) - 2) + [cum[-1] - floor]
        up = [0] + [max(d, 0) for d in steps[1:-1]] + [0]
        dn = [0] + [max(-d, 0) for d in steps[1:-1]] + [0]
        return dict(cum=cum, floor=floor, base=base, tot=tot, up=up, dn=dn)
    out["wf_region"] = waterfall(out["region"])
    out["wf_cat"] = waterfall(out["category"])
    mat_a = [[rev(ps, pe, region=r, cat=c) for c in CATS] for r in REGIONS]
    mat_b = [[bud(ps, pe, region=r, cat=c) for c in CATS] for r in REGIONS]
    out["matrix"] = (mat_a, mat_b)
    var = [(r[3], i) for i, r in enumerate(out["region"][:4]) if r[3] != ""]
    if var:
        wi = min(var)[1] if len({v for v, _ in var}) == len(var) else [i for v, i in var if v == min(v for v, _ in var)][0]
        amts = [mat_a[wi][j] - mat_b[wi][j] for j in range(5)]
        cj = amts.index(min(amts))
        out["worst"] = dict(reg=REGIONS[wi], pct=out["region"][wi][3], amt=out["region"][wi][2], cat=CATS[cj],
                            cat_pct="" if mat_b[wi][cj] == 0 else amts[cj] / mat_b[wi][cj], cat_amt=amts[cj])
    else:
        out["worst"] = None
    # customers
    cust = []
    for i, cname in enumerate(CUSTOMERS):
        cust.append(dict(name=cname, region=CUST_REGION[cname], rev=rev(ps, pe, cust=cname), py=rev(pys, pye, cust=cname),
                         ord=orders(ps, pe, cust=cname), row=i))
    total = sum(c["rev"] for c in cust)
    ranked = sorted([c for c in cust if c["rev"] > 0], key=lambda c: (c["rev"], c["row"]), reverse=True)
    cum = 0
    for c in ranked:
        cum += c["rev"]
        c["share"], c["cum"] = c["rev"] / total, cum / total
    active = len(ranked)
    n80 = 0 if active == 0 else sum(1 for c in ranked if c["cum"] < 0.8) + 1
    pyt = sum(c["py"] for c in cust)
    act_py = sum(1 for c in cust if c["py"] > 0)
    avg, avg_py = (total / active if active else 0), (pyt / act_py if act_py else 0)
    top10py = sorted((c["py"] for c in cust), reverse=True)[:10]
    out["cust"] = dict(ranked=ranked, total=total, active=active, n80=n80, n80pct=n80 / active if active else 0,
                       top10=sum(c["rev"] for c in ranked[:10]) / total if total else 0, top=ranked[0] if ranked else None,
                       avg=avg, active_py=act_py, avg_py=avg_py, avg_vs_py="" if avg_py == 0 else avg / avg_py - 1,
                       top10_py=sum(top10py) / pyt if pyt else "", on_file=len(CUSTOMERS))
    # reps
    reps = []
    for i, rp in enumerate(REPS):
        a, t = rev(ps, pe, rep=rp), bud(ps, pe, rep=rp)
        py, o, gp = rev(pys, pye, rep=rp), orders(ps, pe, rep=rp), rev(ps, pe, "margin", rep=rp)
        reps.append(dict(rep=rp, region=REP_REGION[rp], rev=a, target=t, pct="" if t == 0 else a / t, py=py,
                         vs_py="" if py == 0 else a / py - 1, ord=o, aov=a / o if o else 0, gp=gp, mg="" if a == 0 else gp / a, row=i))
    rt = sum(r["rev"] for r in reps)
    for r in reps:
        r["share"] = "" if rt == 0 else r["rev"] / rt
    score = sorted([r for r in reps if r["target"] > 0], key=lambda r: (r["pct"], r["row"]), reverse=True)
    out["reps"] = dict(all=reps, score=score, with_target=len(score), on_target=sum(1 for r in score if r["pct"] >= 1),
                       team=(rt / sum(r["target"] for r in reps)) if sum(r["target"] for r in reps) else 0)
    # forecast
    fy = rep_m.year
    a3 = rev(edate(rep_m, -2), rep_m)
    b3 = bud(edate(rep_m, -2), rep_m)
    ach3 = a3 / b3 if b3 else 0
    rows = []
    for n in range(12):
        m = date(fy, n + 1, 1)
        act = rev(m, m) if m <= rep_m else 0
        bb = bud(m, m)
        fc = bb * ach3 if m > rep_m else 0
        rows.append(dict(m=m, act=act, bud=bb, py=rev(edate(m, -12), edate(m, -12)), fc=fc, q=act + fc))
    c1 = c2 = 0
    for r in rows:
        c1 += r["q"]; c2 += r["bud"]
        r["cum"], r["cumb"], r["vs"] = c1, c2, ("" if r["bud"] == 0 else r["q"] / r["bud"] - 1)
    ytd, fyb, fcy = sum(r["act"] for r in rows), sum(r["bud"] for r in rows), sum(r["q"] for r in rows)
    left = 12 - rep_m.month
    straight = ytd + a3 / 3 * left
    fpy = sum(r["py"] for r in rows)
    ytdb = sum(r["bud"] for r in rows if r["m"] <= rep_m)
    ytdpy = sum(r["py"] for r in rows if r["m"] <= rep_m)
    out["fc"] = dict(rows=rows, summary=[ach3, a3 / 3, ytd, fyb, fcy, "" if fyb == 0 else fcy / fyb - 1, left, straight,
                                         "" if fyb == 0 else straight / fyb - 1, fpy, "" if fpy == 0 else fcy / fpy - 1, ytdb,
                                         "" if ytdb == 0 else ytd / ytdb - 1, ytdpy, "" if ytdpy == 0 else ytd / ytdpy - 1,
                                         fcy - fyb, sum(r["fc"] for r in rows)], fy=fy)
    # labels and insight sentences
    period = (f"Jan-{rep_m:%b %Y}" if view == "YTD" else f"{rep_m:%B %Y}")
    pylab = (f"Jan-{pye:%b %Y}" if view == "YTD" else f"{pye:%b %Y}")
    out["period"] = period
    ins, tone = [], []
    s1 = f"{period}: revenue {t_num(P['rev'])}"
    s1 += ", no budget" if P["bud"] == 0 else f", {t_pct(out['delta'][0][0], sign=True)} vs budget ({t_signed(P['rev'] - P['bud'])})"
    s1 += ", no prior-year data." if PY["rev"] == 0 else f", {t_pct(out['delta'][0][1], sign=True)} vs {pylab}."
    ins.append(s1); tone.append(0 if P["bud"] == 0 else (-1 if P["rev"] < P["bud"] else 1))
    w = out["worst"]
    if w is None:
        ins.append("No budget in this selection."); tone.append(0)
    else:
        mid = " " if w["cat_pct"] == "" else f" {t_pct(w['cat_pct'], sign=True)} "
        ins.append(f"Weakest region: {w['reg']} {t_pct(w['pct'], sign=True)} vs budget, driven by {w['cat']}{mid}({t_signed(w['cat_amt'])}).")
        tone.append(-1 if w["pct"] < 0 else 1)
    if P["rev"] == 0:
        ins.append("No sales in this selection.")
    else:
        s3 = f"Gross margin {t_pct(P['mg'])}"
        if P["bud"] != 0:
            s3 += f": {t_pts(out['delta'][2][0])} pts vs budget"
        if PY["rev"] != 0:
            s3 += f", {t_pts(out['delta'][2][1])} pts vs prior year"
        ins.append(s3 + ".")
    tone.append(0 if P["rev"] == 0 or P["bud"] == 0 else (-1 if P["mg"] < P["bmg"] else 1))
    cu = out["cust"]
    ins.append("No sales in this selection." if active == 0 else
               f"{n80} of {active} active customers bring 80% of revenue; the largest, {cu['top']['name']}, brings {t_pct(cu['top']['share'])}.")
    tone.append(0)
    rp = out["reps"]
    if rp["with_target"] == 0:
        ins.append("No sales targets in this selection."); tone.append(0)
    else:
        low = score[-1]
        ins.append(f"Lowest vs target: {low['rep']} at {t_pct(low['pct'], 0)}; {rp['on_target']} of {rp['with_target']} reps on or above target.")
        tone.append(-1 if low["pct"] < 1 else 1)
    if fyb == 0:
        ins.append(f"No full-year budget for {fy}."); tone.append(0)
    else:
        ins.append(f"At the recent run-rate {fy} ends at {t_num(fcy)}, {t_pct(fcy / fyb - 1, sign=True)} vs the full-year budget of {t_num(fyb)}.")
        tone.append(-1 if fcy < fyb else 1)
    out["insights"], out["tones"] = ins, tone
    out["cards"] = dict(
        rev_budget="" if out["delta"][0][0] == "" else ("▲ " if out["delta"][0][0] >= 0 else "▼ ") + t_pct(abs(out["delta"][0][0])) + " vs budget",
        rev_month="" if out["delta"][0][2] == "" else ("▲ " if out["delta"][0][2] >= 0 else "▼ ") + t_pct(abs(out["delta"][0][2]))
        + f" {rep_m:%b} vs {prev:%b}")
    return out


# ---------------------------------------------------------------- 2. Excel
def xl_date(d):
    return d.isoformat()


def set_items(wb, cache, wanted):
    sc = wb.SlicerCaches(cache)
    names = [it.Name for it in sc.SlicerItems]
    if set(wanted) >= set(names):
        sc.ClearManualFilter()
        return
    for it in sc.SlicerItems:
        if it.Name in wanted:
            it.Selected = True
    for it in sc.SlicerItems:
        if it.Name not in wanted:
            it.Selected = False


def apply(xl, wb, sel):
    set_items(wb, "Slicer_View", [sel["view"]])
    set_items(wb, "Slicer_Region", sel["regions"])
    set_items(wb, "Slicer_Category", sel["cats"])
    set_items(wb, "Slicer_Rep", sel["reps"])
    m = sel["month"]
    end = edate(m, 1).toordinal() - 1
    wb.SlicerCaches("Timeline_Month").TimelineState.SetFilterDateRange(xl_date(m), xl_date(date.fromordinal(end)))
    xl.Calculate()


def rows_close(xs, ps, tol=0.011):
    return all(close(a, b, tol) if not isinstance(b, str) or b == "" else a == b for a, b in zip(xs, ps))


ALL = dict(regions=REGIONS, cats=CATS, reps=REPS)
SCENARIOS = [
    dict(name="Oct 2026 month, all", month=date(2026, 10, 1), view="MTD", **ALL),
    dict(name="Oct 2026 month, West", month=date(2026, 10, 1), view="MTD", regions=["West"], cats=CATS, reps=REPS),
    dict(name="Oct 2026 YTD, all", month=date(2026, 10, 1), view="YTD", **ALL),
    dict(name="May 2026 month, North, Paint+Garden", month=date(2026, 5, 1), view="MTD", regions=["North"], cats=["Paint", "Garden"], reps=REPS),
    dict(name="Mar 2026 YTD, reps Finn Walsh + Dev Patel", month=date(2026, 3, 1), view="YTD", regions=REGIONS, cats=CATS,
         reps=["Finn Walsh", "Dev Patel"]),
    dict(name="Feb 2025 month, all (no prior year)", month=date(2025, 2, 1), view="MTD", **ALL),
]

xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(ROOT / "Sales-Dashboard.xlsx"), 0, True)
    c = wb.Worksheets("Calc")
    V = lambda a: c.Range(a).Value
    NV = lambda n: wb.Names(n).RefersToRange.Value
    check("Sales table has every CSV line", wb.Worksheets("Sales data").ListObjects("Sales").ListRows.Count == len(sales))
    check("Budget table has every CSV row", wb.Worksheets("Budget data").ListObjects("Budget").ListRows.Count == len(budget))
    implicit, modern = [], []
    import re
    MODERN = re.compile(r"(?<![A-Z0-9_.])(MINIFS|MAXIFS|TEXTJOIN|CONCAT|IFS|SWITCH|XLOOKUP|XMATCH|FILTER|LET|UNIQUE|SORT|SORTBY|SEQUENCE|LAMBDA)\(")
    for ws in wb.Worksheets:
        try:
            cells = ws.UsedRange.SpecialCells(-4123)  # formulas
        except Exception:
            continue
        for a in cells.Areas:
            for x in a.Cells:
                f2 = x.Formula2
                if "@" in f2.replace("[@", ""):
                    implicit.append(f"{ws.Name}!{x.Address}")
                if MODERN.search(f2.upper().replace("_XLFN.", "")):
                    modern.append(f"{ws.Name}!{x.Address}")
    check("No implicit-intersection @ formulas (opens cleanly in Excel 2016)", not implicit, str(implicit[:3]))
    check("No Microsoft 365 / 2019-only functions", not modern, str(modern[:3]))

    # page furniture
    pages = ["Dashboard", "Products", "Customers", "Reps", "Forecast", "Checks"]
    for p in pages:
        ws = wb.Worksheets(p)
        navs = {s.Name: s for s in ws.Shapes if s.Name.startswith("nav_")}
        links = {}
        for h in ws.Hyperlinks:
            try:
                links[h.Shape.Name] = h.SubAddress
            except Exception:
                pass
        ok = len(navs) == 6 and all(links.get(f"nav_{q}") == f"'{q}'!A1" for q in pages)
        check(f"{p}: six navigation buttons, each linked to its page", ok, str(links))
    slicers_by_sheet = defaultdict(list)
    for sc in wb.SlicerCaches:
        for s in sc.Slicers:
            slicers_by_sheet[s.Shape.Parent.Name].append(sc.Name)
    for p in pages[:5]:
        got = sorted(slicers_by_sheet[p])
        check(f"{p}: View, Region, Category, Sales rep slicers + Month timeline", got == sorted(
            ["Slicer_View", "Slicer_Region", "Slicer_Category", "Slicer_Rep", "Timeline_Month"]), str(got))
    check("Overview: six KPI sparklines", wb.Worksheets("Dashboard").Cells.SparklineGroups.Count == 6,
          str(wb.Worksheets("Dashboard").Cells.SparklineGroups.Count))
    wb.SlicerCaches("Timeline_Month").ClearDateFilter()
    xl.Calculate()
    shown = sorted(int(v[0]) for v in wb.Names("SelMonths").RefersToRange.Value if v[0] not in (None, ""))
    check("Timeline cleared = every month in the data, reporting month = the last one",
          shown == [m.year * 100 + m.month for m in MONTHS] and V("B4").date() == MONTHS[-1], f"{shown[:3]}...{shown[-1:]}")

    story_out = None
    for sel in SCENARIOS:
        apply(xl, wb, sel)
        p = model(sel)
        tag = sel["name"]
        check(f"{tag}: reporting month and view", V("B4") is not None and V("B4").date() == sel["month"] and NV("ViewMode") == sel["view"],
              f"{V('B4')} {NV('ViewMode')}")
        k = V("E4:H12")
        ok = all(close(k[i][j], p["kpi"][i][j], 0.011 if i < 5 or i == 7 else 1e-7) for i in range(9) for j in range(4))
        check(f"{tag}: KPI matrix (period, prior year, month, prior month)", ok, f"{k[0]} vs {p['kpi'][0]}")
        dl = V("E15:G20")
        check(f"{tag}: changes vs budget / prior year / prior month", all(closer(dl[i][j], p["delta"][i][j]) for i in range(6) for j in range(3)),
              f"{dl[0]} vs {p['delta'][0]}")
        t = V("M4:AK15")
        idx = [0, 1, 2, 3, 4, 5, 6, 7] + list(range(8, 25))
        ok_t = all(close(t[i][0], p["trend"][i][0]) and close(t[i][1], p["trend"][i][1]) and close(t[i][2], p["trend"][i][2]) for i in range(12))
        check(f"{tag}: 12-month actual / budget / prior year", ok_t)
        ok_t2 = all(closer(t[i][4], p["trend"][i][4]) and t[i][5] == p["trend"][i][5] and close(t[i][6], p["trend"][i][6]) and closer(t[i][7], p["trend"][i][7])
                    for i in range(12))
        check(f"{tag}: 12-month margin, orders, average order, achieved (sparklines)", ok_t2)
        ok_h = all(close(t[i][8 + j], p["trend"][i][8 + j]) and close(t[i][12 + j], p["trend"][i][12 + j]) and closer(t[i][16 + j], p["trend"][i][16 + j])
                   for i in range(12) for j in range(4))
        check(f"{tag}: heatmap region x month (% vs budget)", ok_h)
        check(f"{tag}: 12-month revenue by category", all(close(t[i][20 + j], p["trend"][i][20 + j]) for i in range(12) for j in range(5)))
        reg = V("B24:I28")
        check(f"{tag}: region table", all(rows_close(reg[i], [p["region"][i][x] for x in range(8)]) for i in range(5)), f"{reg[3]} vs {p['region'][3][:8]}")
        cat = V("B33:K38")
        check(f"{tag}: category table", all(rows_close(cat[i], [p["category"][i][x] for x in (0, 1, 2, 3, 4, 5, 6, 7, 8, 10)]) for i in range(6)),
              f"{cat[3]} vs {p['category'][3]}")
        for nm, top, w in (("region", 42, p["wf_region"]), ("category", 53, p["wf_cat"])):
            n = len(w["cum"])
            got = V(f"C{top}:G{top + n - 1}")
            ok = all(close(got[i][0], w["cum"][i]) and close(got[i][1], w["base"][i]) and close(got[i][2], w["tot"][i]) and
                     close(got[i][3], w["up"][i]) and close(got[i][4], w["dn"][i]) for i in range(n))
            check(f"{tag}: waterfall by {nm} (floor {w['floor']:,.0f})", ok and close(V(f"B{top + n}"), w["floor"]), f"{got[0]} {V(f'B{top + n}')}")
        ma, mb = V("B64:F67"), V("H64:L67")
        check(f"{tag}: region x category matrix", all(close(ma[i][j], p["matrix"][0][i][j]) and close(mb[i][j], p["matrix"][1][i][j])
                                                     for i in range(4) for j in range(5)))
        cu = p["cust"]
        rk = V("AW4:BD63")
        exp = [[x["name"], x["region"], x["rev"], x["py"], "" if x["py"] == 0 else x["rev"] / x["py"] - 1, x["ord"], x["share"], x["cum"]]
               for x in cu["ranked"]] + [[""] * 8] * (60 - len(cu["ranked"]))
        ok = all(rk[i][0] == exp[i][0] and rk[i][1] == exp[i][1] and rows_close(rk[i][2:], exp[i][2:]) for i in range(60))
        check(f"{tag}: customers ranked (all {cu['active']})", ok, f"{rk[0]} vs {exp[0]}")
        summ = [NV(n) for n in ("CustTotal", "CustActive", "CustOnFile", "CustN80", "CustN80Pct", "CustTop10", "CustTopShare", "CustAvg",
                                "CustActivePY", "CustAvgPY", "CustAvgVsPY", "CustTop10VsPY")]
        want = [cu["total"], cu["active"], cu["on_file"], cu["n80"], cu["n80pct"], cu["top10"], cu["top"]["share"] if cu["top"] else 0,
                cu["avg"], cu["active_py"], cu["avg_py"], cu["avg_vs_py"], cu["top10_py"]]
        check(f"{tag}: Pareto 80/20 and customer cards ({cu['n80']} of {cu['active']} = 80%)", rows_close(summ, want, 1e-6) or
              all(close(a, b, 0.011) for a, b in zip(summ, want)), f"{summ} vs {want}")
        rp = p["reps"]
        reps_x = c.ListObjects("Reps")
        xr = {reps_x.DataBodyRange.Cells(i, 1).Value: [reps_x.DataBodyRange.Cells(i, j).Value for j in range(3, 14)] for i in range(1, reps_x.ListRows.Count + 1)}
        ok = all(rows_close(xr[r["rep"]][:10], [r["rev"], r["target"], r["pct"], r["py"], r["vs_py"], r["ord"], r["aov"], r["gp"], r["mg"], r["share"]])
                 for r in rp["all"])
        check(f"{tag}: every rep: revenue, target, % of target, PY, orders, margin, share", ok and len(xr) == len(rp["all"]))
        sc = V("BI15:BY22")
        exp = [[r["rep"], r["region"], r["rev"], r["target"], r["pct"]] for r in rp["score"]]
        ok = all(sc[i][0] == exp[i][0] and sc[i][1] == exp[i][1] and rows_close(sc[i][2:5], exp[i][2:]) for i in range(len(exp))) and \
            all(sc[i][0] == "" for i in range(len(exp), 8))
        g_ok = all(closer(sc[i][12], min(exp[i][4], 1)) and closer(sc[i][13], min(max(exp[i][4] - 1, 0), 0.5)) and
                   sc[i][16] == t_pct(exp[i][4], 0) for i in range(len(exp)))
        check(f"{tag}: scorecard ranked by % of target + gauges", ok and g_ok, f"{sc[0][:6]} vs {exp[:1]}")
        check(f"{tag}: reps on target / lowest", NV("RepsWithTarget") == rp["with_target"] and NV("RepsOnTarget") == rp["on_target"] and
              (rp["with_target"] == 0 or NV("RepLow") == rp["score"][-1]["rep"]) and closer(NV("RepTeamPct"), rp["team"]))
        fc = V("M20:T31")
        ok = all(rows_close(fc[i], [r["act"], r["bud"], r["py"], r["fc"], r["q"], r["cum"], r["cumb"], r["vs"]]) for i, r in enumerate(p["fc"]["rows"]))
        check(f"{tag}: forecast month by month ({p['fc']['fy']})", ok, f"{fc[-1]}")
        fs = [x[0] for x in V("W20:W36")]
        check(f"{tag}: forecast summary (year-end {p['fc']['summary'][4]:,.0f})",
              all(closer(a, b) if i in (0, 5, 8, 10, 12, 14) else close(a, b) for i, (a, b) in enumerate(zip(fs, p["fc"]["summary"]))),
              f"{fs} vs {p['fc']['summary']}")
        ins = [x[0] for x in V("B83:B88")]
        tones = [x[0] for x in V("C83:C88")]
        for i in range(6):
            check(f"{tag}: insight {i + 1} word for word", ins[i] == p["insights"][i], f"\n      xl: {ins[i]}\n      py: {p['insights'][i]}")
        check(f"{tag}: insight colours (good/bad)", tones == p["tones"], f"{tones} vs {p['tones']}")
        d = wb.Worksheets("Dashboard")
        check(f"{tag}: revenue card notes", d.Range("L20").Value == p["cards"]["rev_budget"] and d.Range("L22").Value == p["cards"]["rev_month"],
              f"{d.Range('L20').Value!r} / {d.Range('L22').Value!r} vs {p['cards']}")
        res = [x[0] for x in wb.Names("CheckResults").RefersToRange.Value]
        check(f"{tag}: all {len(res)} data checks PASS", res == ["PASS"] * len(res), str(res))
        errs = []
        for pg in pages:
            ws = wb.Worksheets(pg)
            vals = ws.Range("A1:BM60").Value
            errs += [f"{pg}!r{r + 1}c{j + 1}" for r, row in enumerate(vals) for j, x in enumerate(row) if isinstance(x, int) and x < -2_000_000_000]
        check(f"{tag}: no error values on the six pages", not errs, str(errs[:5]))
        if sel["name"].startswith("Oct 2026 month, West"):
            story_out = (p, ins)

    # ---------------------------------------------------------------- 3. planted story (Oct 2026, West)
    p, ins = story_out
    check("Story: West is more than 10% below budget in Oct 2026", p["region"][3][3] < -0.1, f"{p['region'][3][3]:.1%}")
    check("Story: Garden is the category that explains it (more than 50% below budget)",
          p["worst"]["cat"] == story["miss"]["category"] and p["worst"]["cat_pct"] < -0.5, str(p["worst"]))
    check("Story: the insight banner names West and Garden", "West" in ins[1] and "driven by Garden" in ins[1], ins[1])
    allp = model(SCENARIOS[0])
    check("Story: the other three regions are within 10% of budget (all regions view)",
          all(abs(r[3]) < 0.1 for r in allp["region"][:3]), str([r[3] for r in allp["region"][:3]]))
    apply(xl, wb, SCENARIOS[0])
    wb.Close(False)
finally:
    xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
