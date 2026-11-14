"""Generates the fictional data for "Brightwire Electrical Contractors" (made-up company): three projects, eight weeks.

Writes into ../data/ what the project managers typed, and into ../test/expected.json the designed answer:
  projects.csv         - one row per project: sheet name, project, PM, contract value, this week's note
  pm-week-2026-11-13.csv - this week's PM sheets: per project and cost code: budget, committed, actual to date, % complete
  history-seed.csv     - the seven earlier weekly snapshots (25/09 - 06/11/2026), one row per project and week
The story is planted on purpose:
  P-2401 clinic fit-out   on budget (CPI about 1.01)
  P-2402 warehouse        labor overrun on rough-in + shop prefab behind the field  -> at risk
  P-2403 school           copper wire price rise on material                         -> watch
"""
import csv, json, random
from datetime import date, timedelta
from pathlib import Path

random.seed(20261113)
ROOT = Path(__file__).resolve().parent.parent
WEEKS = [date(2026, 11, 13) - timedelta(days=7 * i) for i in range(7, -1, -1)]  # 8 Fridays, the last = this week

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

# per project: contract, budget per code, final % complete this week, final CPI per code, first week the code starts
PROJECTS = [
    dict(sheet="P-2401", name="Riverside Medical Clinic fit-out", pm="Dana Whitfield", contract=1_250_000,
         budget=[228000, 96000, 142000, 118000, 86000, 31000, 74000, 92000, 58000, 125000],
         pct=[0.92, 0.35, 0.88, 0.95, 0.70, 0.60, 0.55, 1.00, 0.95, 0.62],
         cpi=[1.03, 1.00, 1.01, 1.00, 0.99, 1.02, 1.00, 1.04, 1.02, 1.00],
         start=[1, 5, 1, 1, 3, 1, 2, 1, 1, 1],
         note="Level 2 rough-in complete, inspection passed 11/12. Lighting fixtures delivered. Trim crew starts Monday."),
    dict(sheet="P-2402", name="Harbor Point logistics warehouse", pm="Marcus Oyelaran", contract=2_400_000,
         budget=[512000, 168000, 336000, 284000, 212000, 74000, 152000, 168000, 96000, 224000],
         pct=[0.68, 0.0, 0.62, 0.55, 0.20, 0.50, 0.30, 0.32, 0.40, 0.47],
         cpi=[0.80, 1.00, 0.96, 0.98, 0.97, 0.90, 1.00, 0.88, 0.92, 0.95],
         start=[1, 9, 1, 2, 6, 1, 3, 1, 2, 1],
         note="Shop prefab racks 3 weeks behind (welder out, overtime approved). Rough-in labor over: rework at dock doors, change order CO-07 for $38,500 submitted."),
    dict(sheet="P-2403", name="Oakmont Middle School modernization", pm="Priya Nandakumar", contract=1_780_000,
         budget=[318000, 142000, 236000, 204000, 128000, 46000, 118000, 96000, 64000, 178000],
         pct=[0.82, 0.30, 0.90, 0.85, 0.55, 0.70, 0.60, 0.85, 0.75, 0.71],
         cpi=[0.99, 1.00, 0.84, 0.97, 0.96, 1.00, 1.00, 1.00, 0.99, 1.00],
         start=[1, 6, 1, 1, 4, 1, 2, 1, 1, 1],
         note="Copper wire up 14% since bid; the rest of the wire is now bought. Gym panels energized. On schedule for the winter break shutdown."),
]


def week_values(p, w):
    """Cost-code values of project p in week index w (0..7)."""
    rows = []
    for i, (code, desc, cat) in enumerate(CODES):
        s = p["start"][i] - 1
        if w < s or p["pct"][i] == 0:
            pct = 0.0
        else:
            frac = (w - s + 1) / (8 - s)  # straight ramp from the start week to this week
            pct = round(p["pct"][i] * frac, 3)
        cpi = 1 + (p["cpi"][i] - 1) * (0.4 + 0.6 * (w + 1) / 8)  # the overrun grows over the weeks
        actual = round(p["budget"][i] * pct / cpi) if pct else 0
        committed = max(actual, round(p["budget"][i] * min(1.0, pct + 0.25))) if pct else round(p["budget"][i] * 0.1 * (cat == "Material"))
        rows.append(dict(code=code, budget=p["budget"][i], committed=committed, actual=actual, pct=pct))
    return rows


def summarise(rows):
    budget = sum(r["budget"] for r in rows)
    ev = sum(r["budget"] * r["pct"] for r in rows)
    actual = sum(r["actual"] for r in rows)
    fac = sum(r["actual"] / r["pct"] if r["pct"] > 0 else max(r["budget"], r["actual"]) for r in rows)
    field = [r for r in rows if r["code"] not in SHOP]
    shop = [r for r in rows if r["code"] in SHOP]
    return dict(budget=budget, committed=sum(r["committed"] for r in rows), actual=actual, earned=round(ev, 2),
                field=round(sum(r["budget"] * r["pct"] for r in field) / sum(r["budget"] for r in field), 6),
                shop=round(sum(r["budget"] * r["pct"] for r in shop) / sum(r["budget"] for r in shop), 6),
                cpi=round(ev / actual, 6) if actual else None, fac=round(fac, 2))


def main():
    (ROOT / "data").mkdir(exist_ok=True)
    (ROOT / "test").mkdir(exist_ok=True)
    with open(ROOT / "data" / "projects.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        wr.writerow(["Sheet", "Project", "PM", "Contract value", "This week's note"])
        for p in PROJECTS:
            wr.writerow([p["sheet"], p["name"], p["pm"], p["contract"], p["note"]])
    with open(ROOT / "data" / "cost-codes.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        wr.writerow(["Cost code", "Description", "Category"])
        wr.writerows(CODES)
    this_week, history, expected = [], [], {"week_ending": WEEKS[-1].isoformat(), "projects": {}}
    for p in PROJECTS:
        for w, wk in enumerate(WEEKS):
            rows = week_values(p, w)
            s = summarise(rows)
            snap = dict(week=wk.isoformat(), sheet=p["sheet"], project=p["name"], contract=p["contract"], **s)
            if w < 7:
                history.append(snap)
            else:
                this_week += [dict(sheet=p["sheet"], **r) for r in rows]
                expected["projects"][p["sheet"]] = snap
    with open(ROOT / "data" / f"pm-week-{WEEKS[-1].isoformat()}.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.DictWriter(f, ["sheet", "code", "budget", "committed", "actual", "pct"])
        wr.writeheader()
        wr.writerows(this_week)
    with open(ROOT / "data" / "history-seed.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.DictWriter(f, list(history[0].keys()))
        wr.writeheader()
        wr.writerows(history)
    e = expected["projects"]
    status = lambda c: "On track" if c >= 0.98 else ("Watch" if c >= 0.92 else "At risk")
    for k, v in e.items():
        v["status"] = status(v["cpi"])
    expected["planted_status"] = {"P-2401": "On track", "P-2402": "At risk", "P-2403": "Watch"}
    expected["portfolio"] = {k: round(sum(v[k] for v in e.values()), 2) for k in ("contract", "budget", "committed", "actual", "earned", "fac")}
    expected["portfolio"]["cpi"] = round(expected["portfolio"]["earned"] / expected["portfolio"]["actual"], 6)
    expected["history_rows_seed"] = len(history)
    (ROOT / "test" / "expected.json").write_text(json.dumps(expected, indent=1), encoding="utf-8")
    for k, v in e.items():
        print(k, v["status"], f"CPI {v['cpi']:.3f}", f"field {v['field']:.1%}", f"shop {v['shop']:.1%}",
              f"budget {v['budget']:,} actual {v['actual']:,} FAC {v['fac']:,.0f} contract {v['contract']:,}")
    print("portfolio", expected["portfolio"])


if __name__ == "__main__":
    main()
