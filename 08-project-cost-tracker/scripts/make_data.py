"""Generates the fictional data for "Brightwire Electrical Contractors" (made-up company): three projects, eight weeks.

Writes into ../data/ what the project managers typed, and into ../test/expected.json the designed answer:
  projects.csv           - one row per project: sheet, project, PM, contract, this week's note, billing/cash to date, crew
  pm-week-2026-11-13.csv - this week's PM sheets: per project and cost code: budget, committed, actual to date,
                           planned % to date (looked up from the baseline) and % complete
  baseline.csv           - the cost-loaded baseline schedule: planned % to date, planned value (cumulative) and planned
                           crew per project, cost code and week, from the first week to the end of the longest project
  history-seed.csv       - the seven earlier weekly snapshots (25/09 - 06/11/2026), one row per project and week
  change-orders.csv      - change orders: Approved / Pending / Rejected, price to the owner and our cost estimate
  risks.csv              - risk register: probability and impact 1-5, cost if it happens
The story is planted on purpose:
  P-2401 clinic fit-out   on budget (CPI about 1.01), on schedule, slightly overbilled (healthy cash)
  P-2402 warehouse        labor overrun on rough-in + shop prefab behind -> At risk on cost AND schedule,
                          more crew than planned, underbilled (cash squeeze), dock-door change order still pending
  P-2403 school           copper wire price rise on material -> Watch on cost, on schedule
"""
import csv, json
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
NOW = 7                                   # index of this week (8 weeks: 0..7)
WEEK0 = date(2026, 11, 13) - timedelta(days=7 * NOW)
wk = lambda w: WEEK0 + timedelta(days=7 * w)
HEAD_COST = 2000                          # labor cost of one person for one week (40 h x $50), for crew counts

CODES = [  # code, description, category
    ("26-100", "Labor - rough-in", "Labor"),
    ("26-110", "Labor - trim and finish", "Labor"),
    ("26-200", "Material - wire, conduit and fittings", "Material"),
    ("26-210", "Material - switchgear and panels", "Material"),
    ("26-220", "Material - lighting fixtures", "Material"),
    ("26-300", "Equipment - lifts and rentals", "Equipment"),
    ("26-400", "Subcontract - fire alarm and low voltage", "Subcontract"),
    ("26-500", "Shop fabrication - prefab assemblies", "Shop fabrication"),
    ("26-510", "Shop fabrication - panel and rack build", "Shop fabrication"),
    ("26-900", "General conditions and supervision", "General"),
]
SHOP = {c for c, _, cat in CODES if cat == "Shop fabrication"}
LABOR = {c for c, _, cat in CODES if cat == "Labor"}

# per project: contract, budget per code, % complete this week, final CPI per code, first week the code starts (1-based),
# schedule performance per code this week (planned % = % complete / spi), billing factor, retainage
PROJECTS = [
    dict(sheet="P-2401", name="Riverside Medical Clinic fit-out", pm="Dana Whitfield", contract=1_250_000,
         budget=[228000, 96000, 142000, 118000, 86000, 31000, 74000, 92000, 58000, 125000],
         pct=[0.92, 0.35, 0.88, 0.95, 0.70, 0.60, 0.55, 1.00, 0.95, 0.62],
         cpi=[1.03, 1.00, 1.01, 1.00, 0.99, 1.02, 1.00, 1.04, 1.02, 1.00],
         spi=[1.00, 0.97, 1.02, 1.00, 0.98, 1.00, 1.00, 1.00, 1.00, 1.00],
         start=[1, 5, 1, 1, 3, 1, 2, 1, 1, 1], bill=1.06, ret=0.10,
         note="Level 2 rough-in complete, inspection passed 11/12. Lighting fixtures delivered. Trim crew starts Monday."),
    dict(sheet="P-2402", name="Harbor Point logistics warehouse", pm="Marcus Oyelaran", contract=2_400_000,
         budget=[512000, 168000, 336000, 284000, 212000, 74000, 152000, 168000, 96000, 224000],
         pct=[0.68, 0.0, 0.62, 0.55, 0.20, 0.50, 0.30, 0.32, 0.40, 0.47],
         cpi=[0.80, 1.00, 0.96, 0.98, 0.97, 0.90, 1.00, 0.88, 0.92, 0.95],
         spi=[0.92, 1.00, 0.90, 0.92, 0.80, 0.95, 0.85, 0.65, 0.75, 0.95],
         start=[1, 9, 1, 2, 6, 1, 3, 1, 2, 1], bill=0.90, ret=0.10,
         note="Shop prefab racks 3 weeks behind (welder out, overtime approved). Rough-in labor over: rework at dock doors, change order CO-07 for $38,500 submitted."),
    dict(sheet="P-2403", name="Oakmont Middle School modernization", pm="Priya Nandakumar", contract=1_780_000,
         budget=[318000, 142000, 236000, 204000, 128000, 46000, 118000, 96000, 64000, 178000],
         pct=[0.82, 0.30, 0.90, 0.85, 0.55, 0.70, 0.60, 0.85, 0.75, 0.71],
         cpi=[0.99, 1.00, 0.84, 0.97, 0.96, 1.00, 1.00, 1.00, 0.99, 1.00],
         spi=[0.98, 0.95, 1.00, 0.97, 0.92, 1.00, 0.96, 1.00, 0.97, 1.00],
         start=[1, 6, 1, 1, 4, 1, 2, 1, 1, 1], bill=1.00, ret=0.05,
         note="Copper wire up 14% since bid; the rest of the wire is now bought. Gym panels energized. On schedule for the winter break shutdown."),
]

CHANGE_ORDERS = [  # CO, sheet, description, submitted, status, price to owner, our cost estimate
    ("CO-01", "P-2401", "Added exam room lighting, 6 rooms", "2026-10-02", "Approved", 18400, 14200),
    ("CO-02", "P-2401", "Nurse call panel relocation", "2026-11-06", "Pending", 6900, 5300),
    ("CO-03", "P-2403", "Gym scoreboard and shot clock power", "2026-10-09", "Approved", 9800, 7400),
    ("CO-04", "P-2403", "Copper price escalation claim", "2026-10-23", "Rejected", 31000, 0),
    ("CO-05", "P-2402", "EV charger rough-in, 6 bays", "2026-10-16", "Approved", 46200, 35800),
    ("CO-06", "P-2402", "Upsized feeder to the mezzanine", "2026-10-30", "Approved", 22750, 17900),
    ("CO-07", "P-2402", "Dock door rework (owner change; cost is in rough-in)", "2026-11-12", "Pending", 38500, 0),
    ("CO-08", "P-2403", "Feeds for two portable classrooms", "2026-11-13", "Pending", 14600, 11200),
]

RISKS = [  # id, sheet, risk, owner, probability 1-5, impact 1-5, cost if it happens
    ("R-01", "P-2402", "Prefab racks slip further (welder availability)", "Shop lead", 4, 4, 60000),
    ("R-02", "P-2402", "Owner rejects CO-07 dock door rework", "M. Oyelaran", 3, 4, 38500),
    ("R-03", "P-2402", "Switchgear lead time beyond 14 weeks", "Purchasing", 3, 5, 85000),
    ("R-04", "P-2402", "Rough-in productivity stays at CPI 0.80", "M. Oyelaran", 4, 5, 120000),
    ("R-05", "P-2403", "Copper rises again before the last wire buy", "Purchasing", 2, 3, 15000),
    ("R-06", "P-2403", "Winter break shutdown window missed", "P. Nandakumar", 2, 5, 70000),
    ("R-07", "P-2403", "Asbestos found above the classroom ceilings", "P. Nandakumar", 2, 4, 40000),
    ("R-08", "P-2401", "Final inspection needs a re-test", "D. Whitfield", 2, 2, 6000),
    ("R-09", "P-2401", "Fixture substitution approval delay", "D. Whitfield", 1, 2, 4000),
    ("R-10", "P-2401", "Trim crew availability next month", "Field ops", 3, 2, 8000),
]
PROB_PCT = {1: 0.10, 2: 0.30, 3: 0.50, 4: 0.70, 5: 0.90}


def plan_pct(p, i, w):
    """Baseline planned % complete of code i in week w: a straight ramp from the start week, capped at 100%."""
    s = p["start"][i] - 1
    if w < s:
        return 0.0
    if p["pct"][i] == 0:            # not started yet this week: planned over 6 weeks from its start
        return round(min(1.0, (w - s + 1) / 6), 3)
    p_now = min(1.0, p["pct"][i] / p["spi"][i])
    return round(min(1.0, p_now * (w - s + 1) / (NOW + 1 - s)), 3)


def week_values(p, w):
    """Cost-code values of project p in week index w (0..7)."""
    rows = []
    for i, (code, desc, cat) in enumerate(CODES):
        s = p["start"][i] - 1
        if w < s or p["pct"][i] == 0:
            pct = 0.0
        else:
            frac = (w - s + 1) / (NOW + 1 - s)  # straight ramp from the start week to this week
            pct = round(p["pct"][i] * frac, 3)
        cpi = 1 + (p["cpi"][i] - 1) * (0.4 + 0.6 * (w + 1) / 8)  # the overrun grows over the weeks
        actual = round(p["budget"][i] * pct / cpi) if pct else 0
        committed = max(actual, round(p["budget"][i] * min(1.0, pct + 0.25))) if pct else round(p["budget"][i] * 0.1 * (cat == "Material"))
        rows.append(dict(code=code, budget=p["budget"][i], committed=committed, actual=actual,
                         planned=plan_pct(p, i, w), pct=pct))
    return rows


def summarise(rows):
    budget = sum(r["budget"] for r in rows)
    ev = sum(r["budget"] * r["pct"] for r in rows)
    pv = sum(r["budget"] * r["planned"] for r in rows)
    actual = sum(r["actual"] for r in rows)
    fac = sum(r["actual"] / r["pct"] if r["pct"] > 0 else max(r["budget"], r["actual"]) for r in rows)
    field = [r for r in rows if r["code"] not in SHOP]
    shop = [r for r in rows if r["code"] in SHOP]
    return dict(budget=budget, committed=sum(r["committed"] for r in rows), actual=actual, earned=round(ev, 2),
                planned=round(pv, 2),
                field=round(sum(r["budget"] * r["pct"] for r in field) / sum(r["budget"] for r in field), 6),
                shop=round(sum(r["budget"] * r["pct"] for r in shop) / sum(r["budget"] for r in shop), 6),
                cpi=round(ev / actual, 6) if actual else None, spi=round(ev / pv, 6) if pv else None, fac=round(fac, 2))


def labor(rows, key):
    return sum(r["budget"] * r[key] if key != "actual" else r["actual"] for r in rows if r["code"] in LABOR)


def main():
    (ROOT / "data").mkdir(exist_ok=True)
    (ROOT / "test").mkdir(exist_ok=True)
    # baseline horizon: until every code of every project is planned at 100%
    horizon = NOW + 1
    while any(plan_pct(p, i, horizon - 1) < 1 for p in PROJECTS for i in range(len(CODES))):
        horizon += 1
    this_week, history, baseline = [], [], []
    expected = {"week_ending": wk(NOW).isoformat(), "projects": {}, "baseline_weeks": horizon}
    extra = {}
    for p in PROJECTS:
        prev_ac_lab = 0
        cash = []
        for w in range(NOW + 1):
            rows = week_values(p, w)
            s = summarise(rows)
            ac_lab = labor(rows, "actual")
            heads = round((ac_lab - prev_ac_lab) / HEAD_COST)
            prev_ac_lab = ac_lab
            billed = round(p["contract"] * s["actual"] / s["fac"] * p["bill"] / 1000) * 1000
            cash.append(dict(billed=billed, actual=s["actual"]))
            received = round(cash[w - 2]["billed"] * (1 - p["ret"])) if w >= 2 else 0
            paid = round(cash[w - 1]["actual"] * 0.92) if w >= 1 else 0
            snap = dict(week=wk(w).isoformat(), sheet=p["sheet"], project=p["name"], contract=p["contract"], **s,
                        actual_heads=heads, billed=billed, received=received, paid=paid)
            if w < NOW:
                history.append(snap)
            else:
                this_week += [dict(sheet=p["sheet"], **r) for r in rows]
                expected["projects"][p["sheet"]] = snap
                extra[p["sheet"]] = dict(billed=billed, received=received, paid=paid, heads=heads)
        for i, (c, _, _) in enumerate(CODES):
            prev = 0.0
            for w in range(horizon):
                pp = plan_pct(p, i, w)
                pv = p["budget"][i] * pp
                heads = round((pv - prev) / HEAD_COST) if c in LABOR else 0
                prev = pv
                baseline.append(dict(sheet=p["sheet"], code=c, week=wk(w).isoformat(), planned_pct=pp,
                                     planned_value=round(pv, 2), planned_heads=heads))
    bl = {}
    for b in baseline:
        k = (b["sheet"], b["week"])
        bl[k] = bl.get(k, 0) + b["planned_heads"]
    for r in history:
        r["planned_heads"] = bl[(r["sheet"], r["week"])]
    for s, v in expected["projects"].items():
        v["planned_heads"] = bl[(s, v["week"])]

    with open(ROOT / "data" / "projects.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        wr.writerow(["Sheet", "Project", "PM", "Contract value", "This week's note", "Retainage %", "Billed to date",
                     "Cash received to date", "Cash paid to date", "Crew on site"])
        for p in PROJECTS:
            e = extra[p["sheet"]]
            wr.writerow([p["sheet"], p["name"], p["pm"], p["contract"], p["note"], p["ret"], e["billed"], e["received"],
                         e["paid"], e["heads"]])
    with open(ROOT / "data" / "cost-codes.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        wr.writerow(["Cost code", "Description", "Category"])
        wr.writerows(CODES)
    with open(ROOT / "data" / f"pm-week-{wk(NOW).isoformat()}.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.DictWriter(f, ["sheet", "code", "budget", "committed", "actual", "planned", "pct"])
        wr.writeheader()
        wr.writerows(this_week)
    with open(ROOT / "data" / "baseline.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.DictWriter(f, ["sheet", "code", "week", "planned_pct", "planned_value", "planned_heads"])
        wr.writeheader()
        wr.writerows(baseline)
    with open(ROOT / "data" / "history-seed.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.DictWriter(f, list(history[0].keys()))
        wr.writeheader()
        wr.writerows(history)
    with open(ROOT / "data" / "change-orders.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        wr.writerow(["CO", "Sheet", "Description", "Submitted", "Status", "Price", "Cost estimate"])
        wr.writerows(CHANGE_ORDERS)
    with open(ROOT / "data" / "risks.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        wr.writerow(["ID", "Sheet", "Risk", "Owner", "Probability", "Impact", "Cost if it happens"])
        wr.writerows(RISKS)

    e = expected["projects"]
    status = lambda c, g, a: "On track" if c >= g else ("Watch" if c >= a else "At risk")
    ret = {p["sheet"]: p["ret"] for p in PROJECTS}
    for k, v in e.items():
        v["status"] = status(v["cpi"], 0.98, 0.92)
        v["schedule_status"] = status(v["spi"], 0.95, 0.90)
        cos = [c for c in CHANGE_ORDERS if c[1] == k]
        v["co_approved_price"] = sum(c[5] for c in cos if c[4] == "Approved")
        v["co_approved_cost"] = sum(c[6] for c in cos if c[4] == "Approved")
        v["co_pending_price"] = sum(c[5] for c in cos if c[4] == "Pending")
        v["revised_contract"] = v["contract"] + v["co_approved_price"]
        v["forecast_incl_co"] = round(v["fac"] + v["co_approved_cost"], 2)
        v["margin"] = round((v["revised_contract"] - v["forecast_incl_co"]) / v["revised_contract"], 6)
        bac, ev, ac, cpi, spi = v["budget"], v["earned"], v["actual"], v["earned"] / v["actual"], v["earned"] / v["planned"]
        v["eac_cpi"] = round(bac / cpi, 2)
        v["eac_cpi_spi"] = round(ac + (bac - ev) / (cpi * spi), 2)
        v["tcpi"] = round((bac - ev) / (bac - ac), 6)
        v["earned_revenue"] = round(v["revised_contract"] * ac / v["forecast_incl_co"], 2)
        v["over_under"] = round(v["billed"] - v["earned_revenue"], 2)
        v["retainage"] = round(v["billed"] * ret[k], 2)
        v["net_cash"] = v["received"] - v["paid"]
        v["risk_expected"] = round(sum(r[6] * PROB_PCT[r[4]] for r in RISKS if r[1] == k), 2)
    expected["planted_status"] = {"P-2401": "On track", "P-2402": "At risk", "P-2403": "Watch"}
    expected["planted_schedule"] = {"P-2401": "On track", "P-2402": "At risk", "P-2403": "On track"}
    keys = ("contract", "budget", "committed", "actual", "earned", "planned", "fac", "revised_contract", "forecast_incl_co",
            "billed", "earned_revenue", "over_under", "retainage", "received", "paid", "risk_expected", "actual_heads", "planned_heads")
    expected["portfolio"] = {k: round(sum(v[k] for v in e.values()), 2) for k in keys}
    pf = expected["portfolio"]
    pf["cpi"] = round(pf["earned"] / pf["actual"], 6)
    pf["spi"] = round(pf["earned"] / pf["planned"], 6)
    expected["history_rows_seed"] = len(history)
    (ROOT / "test" / "expected.json").write_text(json.dumps(expected, indent=1), encoding="utf-8")
    for k, v in e.items():
        print(k, v["status"], v["schedule_status"], f"CPI {v['cpi']:.3f} SPI {v['spi']:.3f}", f"field {v['field']:.1%}",
              f"shop {v['shop']:.1%}", f"FAC {v['fac']:,.0f} incl CO {v['forecast_incl_co']:,.0f}",
              f"margin {v['margin']:.1%} over/under {v['over_under']:,.0f} heads {v['actual_heads']}/{v['planned_heads']}")
    print("portfolio", pf, "baseline weeks", horizon)


if __name__ == "__main__":
    main()
