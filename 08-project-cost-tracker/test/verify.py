"""Independent check of Project-Cost-Tracker.xlsx.

1. Recomputes every project in plain Python from what the PMs typed and the baseline schedule (data/*.csv), no Excel
   involved: earned value, planned value, CPI, SPI, CV, SV, three forecast methods, TCPI, change orders, margin, WIP
   (earned revenue, over/under billing, retainage, cash), crew, risk exposure and the 5x5 matrix, forecast finish.
   Compares it with the answer planted when the data was generated (test/expected.json).
2. Opens the workbook with Excel and compares every PM sheet line, every Rollup figure (one row per project + totals),
   the Dashboard table, heatmap and cards, the Project view (EVM detail, cost codes, waterfall), the S-curve grid and
   finish dates, change orders and the risk matrix, the WIP schedule, the History table and the Weekly trend grid with
   Python; all eleven Dashboard checks must say PASS; navigation buttons point at real sheets; no macros in the file.
3. On a copy in another folder, plays a few weeks of real use:
   - Refresh All twice in the same week: History does not grow (a week is never saved twice)
   - next week: a PM updates a line, Week ending + 7, Refresh All: 3 new rows, planned value follows the baseline
   - a pending change order is approved: contract, pending and margin move; the Project view and heatmap pickers work
   - bad input makes the right check FAIL: risk score outside 1-5, change order for an unknown project, billing above
     the contract, a project sheet copied in but not on the Dashboard / without a baseline, an unknown cost code
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs pywin32 for steps 2-3)
"""
import csv, json, math, shutil, sys, tempfile, zipfile
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
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def read(name):
    with open(ROOT / "data" / name, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def status(v, green, amber):
    return "On track" if v >= green else "Watch" if v >= amber else "At risk"


# ---------------------------------------------------------------- 1. plain Python
exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
week = date.fromisoformat(exp["week_ending"])
projects = read("projects.csv")
codes = {r["Cost code"]: r["Category"] for r in read("cost-codes.csv")}
code_list = list(codes)
baseline = read("baseline.csv")
base_pct = {(b["sheet"], b["code"], b["week"]): float(b["planned_pct"]) for b in baseline}
base_weeks = sorted({b["week"] for b in baseline})
cos = read("change-orders.csv")
risks = read("risks.csv")
PROB = {1: 0.1, 2: 0.3, 3: 0.5, 4: 0.7, 5: 0.9}


def planned_pct(sheet, code, wk):
    wk = min(wk.isoformat(), base_weeks[-1])  # after the last baseline week the plan stays at its end
    return base_pct.get((sheet, code, wk), 0.0)


lines = defaultdict(list)
for r in read(f"pm-week-{week.isoformat()}.csv"):
    lines[r["sheet"]].append(dict(code=r["code"], budget=float(r["budget"]), committed=float(r["committed"]),
                                  actual=float(r["actual"]), pct=float(r["pct"]), typed_planned=float(r["planned"])))
for s, ls in lines.items():
    for l in ls:
        l["planned"] = planned_pct(s, l["code"], week)


def line_calc(l):
    ev, pv = l["budget"] * l["pct"], l["budget"] * l["planned"]
    fac = l["actual"] / l["pct"] if l["pct"] > 0 else max(l["budget"], l["actual"])
    return dict(ev=ev, pv=pv, fac=fac, cpi=ev / l["actual"] if l["actual"] > 0 else None, spi=ev / pv if pv > 0 else None)


def project_calc(p, ls, cos, risks):
    s = p["Sheet"]
    lc = [line_calc(l) for l in ls]
    v = dict(sheet=s, project=p["Project"], pm=p["PM"], contract=float(p["Contract value"]))
    mine = [c for c in cos if c["Sheet"] == s]
    v["co_price"] = sum(float(c["Price"]) for c in mine if c["Status"] == "Approved")
    v["co_cost"] = sum(float(c["Cost estimate"]) for c in mine if c["Status"] == "Approved")
    v["co_pending"] = sum(float(c["Price"]) for c in mine if c["Status"] == "Pending")
    v["co_app_n"] = sum(c["Status"] == "Approved" for c in mine)
    v["co_pend_n"] = sum(c["Status"] == "Pending" for c in mine)
    v["rev_contract"] = v["contract"] + v["co_price"]
    v["budget"] = sum(l["budget"] for l in ls)
    v["rev_budget"] = v["budget"] + v["co_cost"]
    v["committed"] = sum(l["committed"] for l in ls)
    v["actual"] = ac = sum(l["actual"] for l in ls)
    v["earned"] = ev = sum(x["ev"] for x in lc)
    v["planned"] = pv = sum(x["pv"] for x in lc)
    v["cv"], v["sv"] = ev - ac, ev - pv
    v["cpi"], v["spi"] = ev / ac, ev / pv
    field = [l for l in ls if codes[l["code"]] != "Shop fabrication"]
    shop = [l for l in ls if codes[l["code"]] == "Shop fabrication"]
    v["field"] = sum(l["budget"] * l["pct"] for l in field) / sum(l["budget"] for l in field)
    v["shop"] = sum(l["budget"] * l["pct"] for l in shop) / sum(l["budget"] for l in shop)
    v["fac"] = sum(x["fac"] for x in lc)
    v["fac_co"] = v["fac"] + v["co_cost"]
    bac = v["budget"]
    v["eac_cpi"] = bac / v["cpi"]
    v["eac_cpi_spi"] = ac + (bac - ev) / (v["cpi"] * v["spi"])
    v["tcpi"] = (bac - ev) / (bac - ac)
    v["etc"] = v["fac_co"] - ac
    v["vac"] = v["rev_budget"] - v["fac_co"]
    v["margin"] = (v["rev_contract"] - v["fac_co"]) / v["rev_contract"]
    v["bid_margin"] = (v["contract"] - bac) / v["contract"]
    v["cost_status"] = status(v["cpi"], 0.98, 0.92)
    v["sched_status"] = status(v["spi"], 0.95, 0.90)
    v["pct_cost"] = ac / v["fac_co"]
    v["earned_rev"] = v["rev_contract"] * v["pct_cost"]
    v["billed"] = float(p["Billed to date"])
    v["over_under"] = v["billed"] - v["earned_rev"]
    v["ret_pct"] = float(p["Retainage %"])
    v["retainage"] = v["billed"] * v["ret_pct"]
    v["received"], v["paid"] = float(p["Cash received to date"]), float(p["Cash paid to date"])
    v["net_cash"] = v["received"] - v["paid"]
    v["ar"] = v["billed"] - v["retainage"] - v["received"]
    v["heads"] = float(p["Crew on site"])
    wk = min(week.isoformat(), base_weeks[-1])
    v["pheads"] = sum(int(b["planned_heads"]) for b in baseline if b["sheet"] == s and b["week"] == wk)
    v["baseline_rows"] = sum(1 for b in baseline if b["sheet"] == s and b["week"] == wk)
    v["no_baseline"] = int(v["baseline_rows"] == 0)
    v["over_contract"] = int(v["billed"] > v["rev_contract"])
    mr = [r for r in risks if r["Sheet"] == s]
    v["risk_exp"] = sum(float(r["Cost if it happens"]) * PROB[int(r["Probability"])] for r in mr)
    v["risk_high"] = sum(int(r["Probability"]) * int(r["Impact"]) >= 15 for r in mr)
    v["risk_open"] = len(mr)
    v["note"] = p["This week's note"]
    v["lines"] = lc
    return v


py = {p["Sheet"]: project_calc(p, lines[p["Sheet"]], cos, risks) for p in projects}
SUM_KEYS = ["contract", "co_price", "co_cost", "co_pending", "rev_contract", "budget", "rev_budget", "committed", "actual",
            "earned", "planned", "cv", "sv", "fac", "fac_co", "etc", "vac", "earned_rev", "billed", "over_under", "retainage",
            "received", "paid", "net_cash", "ar", "heads", "pheads", "baseline_rows", "no_baseline", "over_contract",
            "risk_exp", "risk_high", "risk_open", "co_app_n", "co_pend_n"]


def totals(py):
    t = {k: sum(v[k] for v in py.values()) for k in SUM_KEYS}
    bac, ev, ac, pv = t["budget"], t["earned"], t["actual"], t["planned"]
    t.update(cpi=ev / ac, spi=ev / pv, tcpi=(bac - ev) / (bac - ac))
    t["eac_cpi"] = bac / t["cpi"]
    t["eac_cpi_spi"] = ac + (bac - ev) / (t["cpi"] * t["spi"])
    t["margin"] = (t["rev_contract"] - t["fac_co"]) / t["rev_contract"]
    t["bid_margin"] = (t["contract"] - bac) / t["contract"]
    t["pct_cost"] = ac / t["fac_co"]
    t["cost_status"] = status(t["cpi"], 0.98, 0.92)
    t["sched_status"] = status(t["spi"], 0.95, 0.90)
    return t


tot = totals(py)

check("Every cost code typed by the PMs is in the cost code list", all(l["code"] in codes for ls in lines.values() for l in ls))
check("Planned % in the PM export = the baseline schedule for this week (30 lines)",
      all(close(l["typed_planned"], l["planned"], 1e-9) for ls in lines.values() for l in ls))
PLANT = {"budget": "budget", "actual": "actual", "earned": "earned", "planned": "planned", "fac": "fac",
         "forecast_incl_co": "fac_co", "revised_contract": "rev_contract", "eac_cpi": "eac_cpi", "eac_cpi_spi": "eac_cpi_spi",
         "earned_revenue": "earned_rev", "over_under": "over_under", "retainage": "retainage", "net_cash": "net_cash",
         "risk_expected": "risk_exp", "actual_heads": "heads", "planned_heads": "pheads", "billed": "billed",
         "co_approved_price": "co_price", "co_approved_cost": "co_cost", "co_pending_price": "co_pending"}
for s, v in py.items():
    e = exp["projects"][s]
    bad = [k for k, kk in PLANT.items() if not close(v[kk], e[k], 0.02)]
    bad += [k for k in ("cpi", "spi", "tcpi", "margin") if not close(v[k], e[k], 1e-6)]
    check(f"{s}: Python = planted answer ({len(PLANT) + 4} figures: EV, PV, CPI, SPI, 3 EACs, TCPI, COs, WIP, crew, risk)", not bad, f"{bad}")
    check(f"{s}: cost status = planted story ({exp['planted_status'][s]})", v["cost_status"] == exp["planted_status"][s])
    check(f"{s}: schedule status = planted story ({exp['planted_schedule'][s]})", v["sched_status"] == exp["planted_schedule"][s])
pf = exp["portfolio"]
check("Portfolio totals = planted answer",
      all(close(tot[kk], pf[k], 0.02) for k, kk in PLANT.items() if k in pf) and close(tot["cpi"], pf["cpi"], 1e-6)
      and close(tot["spi"], pf["spi"], 1e-6))
w = py["P-2402"]
check("Story: the warehouse is under billed, cash negative, more crew than planned, dock-door CO still pending",
      w["over_under"] < 0 and w["net_cash"] < 0 and w["heads"] > w["pheads"] and
      any(c["CO"] == "CO-07" and c["Status"] == "Pending" for c in cos))
seed = read("history-seed.csv")
check("History seed: 7 earlier weeks x 3 projects", len(seed) == 21 and len({r["week"] for r in seed}) == 7)
check("History seed: CPI = earned / actual and SPI = earned / planned on every row",
      all(close(float(r["cpi"]), float(r["earned"]) / float(r["actual"]), 1e-6) and
          close(float(r["spi"]), float(r["earned"]) / float(r["planned"]), 1e-6) for r in seed))
bpv = defaultdict(float)
for b in baseline:
    bpv[(b["sheet"], b["week"])] += float(b["planned_value"])
check("History seed: planned value = the baseline for that week",
      all(close(float(r["planned"]), bpv[(r["sheet"], r["week"])], 0.01) for r in seed))
check("Baseline: every project ends at 100% of its budget",
      all(close(bpv[(s, base_weeks[-1])], py[s]["budget"], 0.01) for s in py))

# 5x5 risk matrix, forecast finish
matrix = defaultdict(list)
for r in risks:
    matrix[(int(r["Probability"]), int(r["Impact"]))].append(r["ID"])
matrix = {k: ", ".join(v) for k, v in matrix.items()}


def finish(s, spi):
    pvw = [bpv[(s, wk)] for wk in base_weeks]
    dur = sum(1 for x in pvw if x < py[s]["budget"] - 1e-6) + 1
    fdur = math.ceil(dur / spi)
    first = date.fromisoformat(base_weeks[0])
    return dur, fdur, first + timedelta(days=7 * (dur - 1)), first + timedelta(days=7 * (fdur - 1))


# ---------------------------------------------------------------- 2. the workbook
try:
    import win32com.client as win32
except ImportError:
    print("SKIP  workbook checks (pywin32 not installed)")
    win32 = None

FIRST = 18                  # first cost-code row on a PM sheet
R_FIRST, R_TOT, R_PICK, R_CURVE = 5, 10, 12, 13
D_FIRST, HEAT_FIRST = 13, 24  # Dashboard: projects table, heatmap rows (columns E..N = the ten cost codes)
NA = -2146826246            # #N/A as read through COM


def serial(d):
    return (d - date(1899, 12, 30)).days


def rollup_cols(wb):
    ro = wb.Worksheets("Rollup")
    heads = ro.Range("A4:BZ4").Value[0]
    return {h: i + 1 for i, h in enumerate(heads) if h}


ROLL_HEAD = {"project": "Project", "contract": "Contract", "co_price": "Approved CO price", "co_cost": "Approved CO cost",
             "co_pending": "Pending CO price", "rev_contract": "Revised contract", "budget": "Budget (BAC)",
             "rev_budget": "Revised budget", "committed": "Committed", "actual": "Actual cost (AC)", "earned": "Earned value (EV)",
             "planned": "Planned value (PV)", "cv": "Cost variance", "sv": "Schedule variance", "cpi": "CPI", "spi": "SPI",
             "field": "Field %", "shop": "Shop %", "fac": "Forecast (cost codes)", "fac_co": "Forecast incl. COs",
             "eac_cpi": "EAC CPI method", "eac_cpi_spi": "EAC CPI x SPI method", "tcpi": "TCPI", "etc": "Estimate to complete",
             "vac": "Variance at completion", "margin": "Margin at completion", "bid_margin": "Bid margin",
             "cost_status": "Cost status", "sched_status": "Schedule status", "pct_cost": "% complete (cost-to-cost)",
             "earned_rev": "Earned revenue", "billed": "Billed to date", "over_under": "Over / (under) billed",
             "retainage": "Retainage held", "received": "Cash received", "paid": "Cash paid", "net_cash": "Net cash",
             "ar": "Receivable (excl. retainage)", "heads": "Crew on site", "pheads": "Planned crew",
             "baseline_rows": "Baseline rows this week", "no_baseline": "No baseline this week",
             "over_contract": "Billed above contract", "risk_exp": "Risk exposure (expected)", "risk_high": "High risks (score 15+)",
             "risk_open": "Open risks", "co_app_n": "Approved COs", "co_pend_n": "Pending COs", "note": "This week's note"}
RATIO = {"cpi", "spi", "tcpi", "margin", "bid_margin", "field", "shop", "pct_cost"}


def same(got, want, key):
    if isinstance(want, str):
        return got == want
    return close(got, want, 1e-6 if key in RATIO else 0.01)


def rollup_row(wb, row):
    ro = wb.Worksheets("Rollup")
    cols = rollup_cols(wb)
    return {k: ro.Cells(row, cols[h]).Value for k, h in ROLL_HEAD.items()}


def history(wb):
    lo = wb.Worksheets("Snapshots").ListObjects("History")
    cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
    return [dict(zip(cols, row)) for row in lo.DataBodyRange.Value2]


def checks(wb):
    return [c.Value for c in wb.Names("CheckResults").RefersToRange]


def text_cell(wb, label):
    ro = wb.Worksheets("Rollup")
    for r in range(4, 200):
        if ro.Range(f"BD{r}").Value == label:
            return ro.Range(f"BE{r}").Value2
    return None


def check_rollup(wb, py, tag=""):
    tot = totals(py)
    for i, (s, v) in enumerate(py.items()):
        got = rollup_row(wb, R_FIRST + i)
        bad = [k for k in ROLL_HEAD if not same(got[k], v[k], k)]
        check(f"{tag}Rollup {s}: all {len(ROLL_HEAD)} figures = Python", wb.Worksheets("Rollup").Cells(R_FIRST + i, 1).Value == s
              and not bad, f"{[(k, got[k], v[k]) for k in bad[:4]]}")
    got = rollup_row(wb, R_TOT)
    keys = [k for k in ROLL_HEAD if k in tot]
    bad = [k for k in keys if not same(got[k], tot[k], k)]
    check(f"{tag}Rollup totals: {len(keys)} figures = Python", not bad, f"{[(k, got[k], tot[k]) for k in bad[:4]]}")


def check_workbook(wb, tag=""):
    check_rollup(wb, py, tag)
    d = wb.Worksheets("Dashboard")
    cols = dict(F="rev_contract", G="budget", H="actual", I="earned", J="cpi", K="spi", L="field", M="shop", N="fac_co",
                O="tcpi", P="margin", Q="over_under", R="cost_status", S="sched_status")
    bad = [f"{s}.{k}" for i, s in enumerate(py) for c, k in cols.items() if not same(d.Range(f"{c}{D_FIRST + i}").Value, py[s][k], k)]
    check(f"{tag}Dashboard projects table: 14 columns x 3 projects = Python", not bad, f"{bad[:4]}")
    t = D_FIRST + 5
    bad = [k for c, k in cols.items() if k not in ("field", "shop") and not same(d.Range(f"{c}{t}").Value, tot[k], k)]
    check(f"{tag}Dashboard total row = Python", not bad, f"{bad}")
    check(f"{tag}All 11 Dashboard checks PASS", checks(wb) == ["PASS"] * 11, f"{checks(wb)}")


if win32:
    with zipfile.ZipFile(ROOT / BOOK) as z:
        check("No macros: the .xlsx has no VBA project", not any("vbaProject" in n for n in z.namelist()))
    xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(ROOT / BOOK), 0, True)
        bad = []
        for s in py:
            ws = wb.Worksheets(s)
            for i, l in enumerate(lines[s]):
                r = FIRST + i
                c = line_calc(l)
                okl = (ws.Range(f"B{r}").Value == l["code"] and close(ws.Range(f"H{r}").Value, l["planned"], 1e-9) and
                       close(ws.Range(f"J{r}").Value, c["pv"], 0.01) and close(ws.Range(f"K{r}").Value, c["ev"], 0.01) and
                       close(ws.Range(f"P{r}").Value, c["fac"], 0.01) and
                       (ws.Range(f"N{r}").Value == "" if c["cpi"] is None else close(ws.Range(f"N{r}").Value, c["cpi"], 1e-9)) and
                       (ws.Range(f"O{r}").Value == "" if c["spi"] is None else close(ws.Range(f"O{r}").Value, c["spi"], 1e-9)))
                if not okl:
                    bad.append(f"{s} {l['code']}")
        check("All 30 PM sheet lines: planned % from the baseline, PV, EV, CPI, SPI, forecast = Python", not bad, ", ".join(bad[:3]))
        check_workbook(wb)

        d = wb.Worksheets("Dashboard")
        bad = []
        for i, s in enumerate(py):
            for k, code in enumerate(code_list):
                got = d.Cells(HEAT_FIRST + i, 5 + k).Value
                want = py[s]["lines"][k]["cpi"]
                if not (got == "" if want is None else close(got, want, 1e-9)):
                    bad.append(f"{s} {code}")
        check("Heatmap: 3 projects x 10 cost codes = line CPI (blank if not started)", not bad, f"{bad[:3]}")
        check("Heatmap: weakest cost code of the warehouse = 26-100 rough-in",
              str(d.Range(f"O{HEAT_FIRST + 1}").Value).startswith("26-100"), d.Range(f"O{HEAT_FIRST + 1}").Value)
        cards = {"Dash revised contract": tot["rev_contract"], "Dash forecast": tot["fac_co"], "Dash margin": tot["margin"],
                 "Dash cpi": tot["cpi"], "Dash spi": tot["spi"], "Dash over under": tot["over_under"],
                 "WIP retainage": tot["retainage"], "WIP net cash": tot["net_cash"], "WIP ar": tot["ar"],
                 "CO approved": tot["co_price"], "CO pending": tot["co_pending"], "Risk exposure": tot["risk_exp"]}
        bad = [k for k, v in cards.items() if not close(text_cell(wb, k), v, 1e-6 if abs(v) < 5 else 0.01)]
        check(f"KPI cards: {len(cards)} card values = Python", not bad, f"{bad}")
        linked = sum(1 for sh in d.Shapes if sh.Type == 17 and sh.DrawingObject.Formula)
        check("Dashboard cards are live (text boxes linked to cells)", linked >= 12, f"{linked}")
        ro = wb.Worksheets("Rollup")
        g = ro.Range("D17:H17").Value[0]
        frac = min(1, max(0, (tot["cpi"] - 0.7) / 0.6))
        check("CPI gauge: arc = (CPI - 0.70) / 0.60, drawn in the amber band (Watch)",
              close(g[1], frac, 1e-9) and g[0] == 0 and g[2] == 0 and close(g[3], 1 - frac, 1e-9))
        nav_ok, sheets = True, {ws.Name for ws in wb.Worksheets}
        for name in ("Dashboard", "Project view", "S-curve", "COs & risks", "WIP & cash", "Weekly trend", "P-2402"):
            ws = wb.Worksheets(name)
            subs = [h.SubAddress for h in ws.Hyperlinks]
            targets = {sa.split("!")[0].strip("'") for sa in subs}
            if len(subs) != 7 or not targets <= sheets:
                nav_ok = False
        check("Navigation: 7 buttons on every report and PM sheet, all pointing at real sheets", nav_ok)

        # Project view (picked: P-2402)
        pv = wb.Worksheets("Project view")
        v = py["P-2402"]
        evm = ["budget", "planned", "earned", "actual", "cv", "sv", "cpi", "spi", "fac", "eac_cpi", "eac_cpi_spi", "co_cost",
               "fac_co", "etc", "vac", "tcpi"]
        bad = [k for i, k in enumerate(evm) if not same(pv.Range(f"Q{14 + i}").Value, v[k], k)]
        check("Project view P-2402: 16 earned value figures = Python", pv.Range("C6").Value == "P-2402" and not bad, f"{bad}")
        bad = []
        for k, l in enumerate(lines["P-2402"]):
            r = 35 + k
            c = v["lines"][k]
            if not (pv.Range(f"B{r}").Value == l["code"] and close(pv.Range(f"L{r}").Value, c["fac"], 0.01)
                    and close(pv.Range(f"M{r}").Value, l["budget"] - c["fac"], 0.01)):
                bad.append(l["code"])
        check("Project view P-2402: cost code table (forecast, variance) = Python", not bad, f"{bad}")
        wf = ro.Range("B24:B36").Value
        deltas = [c["fac"] - l["budget"] for c, l in zip(v["lines"], lines["P-2402"])]
        check("Waterfall: budget, 10 cost-code changes, approved COs, forecast = Python, and it closes",
              close(wf[0][0], v["budget"], 0.01) and all(close(wf[1 + k][0], dlt, 0.01) for k, dlt in enumerate(deltas))
              and close(wf[11][0], v["co_cost"], 0.01) and close(wf[12][0], v["fac_co"], 0.01) and close(ro.Range("H24").Value, 0, 0.01))

        # S-curve (picked: P-2402)
        sc = wb.Worksheets("S-curve")
        hist_seed = {(r["sheet"], r["week"]): r for r in seed}
        bad = []
        for i, wk in enumerate(base_weeks):
            r = 15 + i
            got = sc.Range(f"O{r}:S{r}").Value[0]
            wd = date.fromisoformat(wk)
            if int(sc.Range(f"N{r}").Value2) != serial(wd) or not close(got[0], bpv[("P-2402", wk)], 0.01):
                bad.append(wk)
            elif wd < week:
                h = hist_seed[("P-2402", wk)]
                if not (close(got[1], float(h["earned"]), 0.01) and close(got[2], float(h["actual"]), 0.01) and close(got[4], float(h["actual_heads"]))):
                    bad.append(wk)
            elif wd == week:
                if not (close(got[1], v["earned"], 0.01) and close(got[2], v["actual"], 0.01) and close(got[4], v["heads"])):
                    bad.append(wk)
            elif not (got[1] == NA and got[2] == NA):
                bad.append(wk)
        check(f"S-curve P-2402: {len(base_weeks)} weeks of planned value = baseline, EV/AC/crew = History, blank after this week",
              not bad, f"{bad[:3]}")
        dur, fdur, bfin, ffin = finish("P-2402", v["spi"])
        check(f"S-curve P-2402: baseline finish {bfin:%d %b %Y}, forecast finish at SPI {ffin:%d %b %Y} ({fdur - dur} weeks late)",
              int(text_cell(wb, "SC baseline finish")) == serial(bfin) and int(text_cell(wb, "SC forecast finish")) == serial(ffin))

        # change orders & risk matrix
        cr = wb.Worksheets("COs & risks")
        lo = cr.ListObjects("Risks")
        exp_cost = [r[0] for r in lo.ListColumns("Expected cost").DataBodyRange.Value]
        check("Risk register: expected cost per risk = Python",
              all(close(g, float(r["Cost if it happens"]) * PROB[int(r["Probability"])], 0.01) for g, r in zip(exp_cost, risks)))
        m1 = lo.HeaderRowRange.Row
        bad = []
        for i, imp in enumerate(range(5, 0, -1)):
            for p in range(1, 6):
                got = cr.Cells(m1 + 1 + i, 12 + p).Value or ""
                if got != matrix.get((p, imp), ""):
                    bad.append(f"P{p}xI{imp}: {got!r}")
        check("Risk matrix 5x5: the IDs in every cell = Python", not bad, f"{bad[:3]}")

        # WIP schedule
        wp = wb.Worksheets("WIP & cash")
        wcols = dict(G="rev_contract", H="fac_co", J="actual", K="pct_cost", L="earned_rev", M="billed", N="over_under",
                     O="retainage", P="received", Q="paid", R="net_cash")
        bad = [f"{s}.{k}" for i, s in enumerate(py) for c, k in wcols.items() if not same(wp.Range(f"{c}{13 + i}").Value, py[s][k], k)]
        check("WIP schedule: 11 columns x 3 projects = Python", not bad, f"{bad[:4]}")

        h = history(wb)
        check("History: 8 weeks x 3 projects = 24 rows", len(h) == 24 and len({r["Week ending"] for r in h}) == 8)
        sh = {(int(r["Week ending"]), r["Sheet"]): r for r in h}
        HK = {"Actual": "actual", "CPI": "cpi", "Planned": "planned", "SPI": "spi", "Planned heads": "planned_heads",
              "Actual heads": "actual_heads", "Billed": "billed", "Cash received": "received", "Cash paid": "paid"}
        okseed = all(close(sh[(serial(date.fromisoformat(r["week"])), r["sheet"])][k], float(r[kk]), 1e-6)
                     for r in seed for k, kk in HK.items())
        check("History: the seven earlier weeks are stored as generated (9 columns checked)", okseed)
        TK = {"Actual": "actual", "Earned": "earned", "Planned": "planned", "CPI": "cpi", "SPI": "spi", "Forecast at completion": "fac",
              "Planned heads": "pheads", "Actual heads": "heads", "Billed": "billed", "Cash received": "received", "Cash paid": "paid"}
        okw = all(same(sh[(serial(week), s)][k], py[s][kk], kk) for s in py for k, kk in TK.items())
        check("History: this week's snapshot = Python (11 columns)", okw)
        tr = wb.Worksheets("Weekly trend")
        head = {tr.Cells(47, c).Value: c for c in range(2, 30) if tr.Cells(47, c).Value}
        check("Weekly trend: last week's CPI and SPI per project = Python",
              all(close(tr.Cells(55, head[f"CPI {s}"]).Value, py[s]["cpi"], 1e-6) and
                  close(tr.Cells(55, head[f"SPI {s}"]).Value, py[s]["spi"], 1e-6) for s in py))
        check("Weekly trend: crew and cash this week = Python",
              close(tr.Cells(55, head["Crew on site"]).Value, tot["heads"]) and close(tr.Cells(55, head["Crew planned"]).Value, tot["pheads"])
              and close(tr.Cells(55, head["Cash received"]).Value, tot["received"], 0.01) and close(tr.Cells(55, head["Cash paid"]).Value, tot["paid"], 0.01))
        check("Weekly trend: first week = 25 Sep 2026", int(tr.Range("B48").Value2) == serial(week - timedelta(days=49)))
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
            nxt = week + timedelta(days=7)
            wb.Worksheets("Settings").Range("B5").Value2 = serial(nxt)
            pm = wb.Worksheets("P-2401")
            pm.Range(f"G{FIRST + 1}").Value = pm.Range(f"G{FIRST + 1}").Value + 6000   # trim and finish: this week's costs
            pm.Range(f"I{FIRST + 1}").Value = 0.40
            refresh()
            h = history(wb)
            new = [r for r in h if int(r["Week ending"]) == serial(nxt)]
            check("Next week + Refresh All: 3 new rows (27)", len(h) == 27 and len(new) == 3, f"{len(h)}")
            p1 = next(r for r in new if r["Sheet"] == "P-2401")
            check("Next week: P-2401 snapshot has the PM's update", close(p1["Actual"], py["P-2401"]["actual"] + 6000, 0.01))
            pv_next = {s: sum(l["budget"] * planned_pct(s, l["code"], nxt) for l in lines[s]) for s in py}
            check("Next week: planned value moved on with the baseline (no typing by the PMs)",
                  all(close(r["Planned"], pv_next[r["Sheet"]], 0.01) for r in new) and pv_next["P-2402"] > py["P-2402"]["planned"])
            after = {(int(r["Week ending"]), r["Sheet"]): r["Actual"] for r in h}
            check("Next week: the earlier 24 rows are unchanged", all(close(after[k], v, 0.01) for k, v in before.items()))
            check("Next week: Weekly trend moves on (last row = new week)",
                  int(wb.Worksheets("Weekly trend").Range("B55").Value2) == serial(nxt))
            check("Next week: all checks PASS", checks(wb) == ["PASS"] * 11, f"{checks(wb)}")
            wb.Worksheets("Settings").Range("B5").Value2 = serial(week)
            pm.Range(f"G{FIRST + 1}").Value = lines["P-2401"][1]["actual"]
            pm.Range(f"I{FIRST + 1}").Value = lines["P-2401"][1]["pct"]
            xl.Calculate()

            # a pending change order is approved
            col = wb.Worksheets("COs & risks").ListObjects("ChangeOrders")
            ids = [r[0] for r in col.ListColumns("CO").DataBodyRange.Value]
            col.ListColumns("Status").DataBodyRange.Cells(ids.index("CO-07") + 1, 1).Value = "Approved"
            xl.Calculate()
            cos2 = [dict(c, Status="Approved") if c["CO"] == "CO-07" else c for c in cos]
            py2 = {p["Sheet"]: project_calc(p, lines[p["Sheet"]], cos2, risks) for p in projects}
            got = rollup_row(wb, R_FIRST + 1)
            check("CO-07 approved: warehouse revised contract +38,500, pending -38,500, margin recomputed",
                  close(got["rev_contract"], py["P-2402"]["rev_contract"] + 38500, 0.01) and close(got["co_pending"], 0, 0.01)
                  and close(got["margin"], py2["P-2402"]["margin"], 1e-9) and close(text_cell(wb, "Dash revised contract"), totals(py2)["rev_contract"], 0.01))
            wb.Worksheets("Project view").Range("C6").Value = "P-2401"
            xl.Calculate()
            check("Project view picker: P-2401 picked, cards and EVM detail follow",
                  close(text_cell(wb, "PV cpi"), py["P-2401"]["cpi"], 1e-9) and close(wb.Worksheets("Project view").Range("Q27").Value, py["P-2401"]["etc"], 0.01))
            wb.Worksheets("Dashboard").Range("R21").Value = "SPI"
            xl.Calculate()
            check("Heatmap switched to SPI: cells = line SPI", all(
                (wb.Worksheets("Dashboard").Cells(HEAT_FIRST + i, 5 + k).Value == "" if py[s]["lines"][k]["spi"] is None
                 else close(wb.Worksheets("Dashboard").Cells(HEAT_FIRST + i, 5 + k).Value, py[s]["lines"][k]["spi"], 1e-9))
                for i, s in enumerate(py) for k in range(10)))
            rk = wb.Worksheets("COs & risks").ListObjects("Risks").ListColumns("Probability").DataBodyRange.Cells(1, 1)
            rk.Value = 6
            xl.Calculate()
            check("Risk probability 6 typed: check 8 says FAIL", checks(wb)[7] == "FAIL")
            rk.Value = int(risks[0]["Probability"])
            cs = col.ListColumns("Sheet").DataBodyRange.Cells(1, 1)
            cs.Value = "P-9999"
            xl.Calculate()
            check("Change order for an unknown project: check 7 says FAIL", checks(wb)[6] == "FAIL")
            cs.Value = cos[0]["Sheet"]
            wb.Worksheets("P-2403").Range("I6").Value = 5_000_000
            xl.Calculate()
            check("Billed above the contract: check 9 says FAIL", checks(wb)[8] == "FAIL")
            wb.Worksheets("P-2403").Range("I6").Value = py["P-2403"]["billed"]
            xl.Calculate()
            check("Inputs restored: all checks PASS again", checks(wb) == ["PASS"] * 11, f"{checks(wb)}")

            wb.Worksheets("Template").Copy(Before=wb.Worksheets("PM end"))
            xl.Calculate()
            copy = wb.Worksheets(wb.Worksheets("PM end").Index - 1)
            copy.Name = "P-2404"
            xl.Calculate()
            check("Sheet copied in but not on the Dashboard: check 1 says FAIL", checks(wb)[0] == "FAIL")
            d = wb.Worksheets("Dashboard")
            d.Range("B16").Value = "P-2404"
            copy.Unprotect()
            copy.Range("C6").Value = "P-2404"
            xl.Calculate()
            check("Added to the Dashboard: check 1 PASS, and check 6 says FAIL until its baseline is added",
                  checks(wb)[0] == "PASS" and checks(wb)[5] == "FAIL")
            check("New project row reads the copied sheet", d.Range("C16").Value == "New project name")
            copy.Range("E18").Value = 50000
            xl.Calculate()
            check("Budget typed on the new sheet reaches the Dashboard total, checks 2-3 still PASS",
                  close(d.Range("G18").Value, tot["budget"] + 50000, 0.01) and checks(wb)[1:3] == ["PASS", "PASS"])
            wb.Worksheets("P-2403").Range(f"B{FIRST}").Value = "26-999"
            xl.Calculate()
            check("Unknown cost code typed: check 4 says FAIL", checks(wb)[3] == "FAIL")
            wb.Close(False)
    finally:
        xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
