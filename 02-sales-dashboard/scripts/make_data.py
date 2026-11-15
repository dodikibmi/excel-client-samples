"""Fictional sales and budget data for "Brightline Supply Co." (made-up wholesaler).

Sales lines: Jan 2025 - Oct 2026 (22 months, so every 2026 month has a prior-year month).
Budget: Jan 2025 - Dec 2026, per month x sales rep x category (the full-year 2026 budget feeds the forecast page).
Writes ../data/sales-lines.csv, ../data/budget.csv and ../test/story.json (facts the dashboard should reveal).
Planted story: steady growth, Garden peaks in spring, and West misses budget in Oct 2026 because Garden sales collapsed.
"""
import csv, json, random
from datetime import date, timedelta
from pathlib import Path

random.seed(2026_10)
ROOT = Path(__file__).resolve().parent.parent
MONTHS = [date(2025, k, 1) for k in range(1, 13)] + [date(2026, k, 1) for k in range(1, 11)]
BUDGET_MONTHS = MONTHS + [date(2026, 11, 1), date(2026, 12, 1)]
REGIONS = {"North": 1.00, "South": 0.85, "East": 1.15, "West": 0.95}
CATEGORIES = {  # base monthly revenue per region, margin, unit price range, monthly seasonality (Jan..Dec)
    "Fasteners":   (21000, 0.38, (2, 40),   [1.0, 1.0, 1.05, 1.05, 1.05, 1.0, 0.95, 0.95, 1.0, 1.05, 1.0, 0.9]),
    "Power Tools": (26000, 0.24, (60, 480), [0.8, 0.85, 0.95, 1.0, 1.0, 0.95, 0.9, 0.95, 1.0, 1.05, 1.2, 1.4]),
    "Paint":       (15000, 0.42, (12, 95),  [0.7, 0.75, 0.9, 1.1, 1.25, 1.3, 1.3, 1.2, 1.05, 0.9, 0.8, 0.7]),
    "Garden":      (12000, 0.35, (8, 160),  [0.4, 0.5, 0.9, 1.5, 1.8, 1.6, 1.3, 1.1, 0.9, 0.7, 0.5, 0.4]),
    "Electrical":  (17000, 0.31, (5, 220),  [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.05, 1.05, 1.0]),
}
FIRST = ["Apex", "Birch", "Cobalt", "Dune", "Ember", "Fjord", "Granite", "Harbor", "Iris", "Juniper", "Kestrel",
         "Lumen", "Maple", "Nimbus", "Orchard", "Pioneer", "Quarry", "Ridge", "Summit", "Tidewater"]
LAST = ["Builders", "Homes", "Contracting", "Property Services", "Renovations", "Hardware", "Landscapes", "Electric"]
# sales rep: (share of the region's sales, share of the region's budget) -> different "% of target" per rep
SALES_REPS = {"North": {"Amara Cole": (0.58, 0.52), "Ben Ortiz": (0.42, 0.48)}, "South": {"Chloe Park": (1, 1)},
              "East": {"Dev Patel": (0.53, 0.50), "Elena Rossi": (0.47, 0.50)}, "West": {"Finn Walsh": (1, 1)}}

customers = []
names = random.sample([f"{a} {b}" for a in FIRST for b in LAST], 44)
for i, n in enumerate(names):
    region = list(REGIONS)[i % 4]
    customers.append(dict(name=n, region=region, size=random.choice([0.4, 0.6, 0.8, 1, 1, 1.4, 2.4, 3.5])))

lines, budget, inv = [], [], 30001
for mi, m in enumerate(BUDGET_MONTHS):
    growth = 1 + 0.009 * mi
    days = ((m.replace(day=28) + timedelta(days=4)).replace(day=1) - m).days
    for region, rw in REGIONS.items():
        custs = [c for c in customers if c["region"] == region]
        reps = SALES_REPS[region]
        for cat, (base, margin, (pmin, pmax), season) in CATEGORIES.items():
            target = base * rw * growth * season[m.month - 1]
            region_budget = target * random.uniform(0.97, 1.06)
            for rep, (_, bshare) in reps.items():
                budget.append(dict(month=m, region=region, rep=rep, category=cat,
                                   budget=round(region_budget * bshare, -2), budget_margin=margin))
            if m not in MONTHS:
                continue  # budget only: the months the forecast page projects
            actual = target * random.uniform(0.9, 1.08)
            if (m, region, cat) == (date(2026, 10, 1), "West", "Garden"):
                actual = target * 0.38  # the planted miss
            if (m, region) == (date(2026, 10, 1), "West") and cat != "Garden":
                actual = target * random.uniform(0.93, 0.98)
            done = 0.0
            while done < actual:
                c = random.choices(custs, weights=[c["size"] for c in custs])[0]
                price = round(random.uniform(pmin, pmax), 2)
                qty = random.choice([1, 2, 3, 4, 5, 6, 8, 10, 12, 20, 24])
                if done + price * qty > actual * 1.02:
                    qty = max(1, int((actual - done) / price))
                cost = round(price * (1 - margin * random.uniform(0.85, 1.15)), 2)
                rep = random.choices(list(reps), weights=[s for s, _ in reps.values()])[0]
                lines.append(dict(date=m + timedelta(days=random.randrange(days)), invoice=f"INV-{inv}",
                                  customer=c["name"], region=region, rep=rep, category=cat, qty=qty, price=price, cost=cost))
                done += price * qty
                if random.random() < 0.55:
                    inv += 1  # otherwise the next line joins the same invoice
            inv += 1

# Invoices can mix categories but always one customer/region/rep: give follow-on lines the invoice's customer
first = {}
for l in lines:
    f = first.setdefault(l["invoice"], l)
    l["customer"], l["region"], l["rep"], l["date"] = f["customer"], f["region"], f["rep"], f["date"]
lines.sort(key=lambda l: (l["date"], l["invoice"]))

(ROOT / "data").mkdir(exist_ok=True); (ROOT / "test").mkdir(exist_ok=True)
with open(ROOT / "data" / "sales-lines.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Date", "Invoice", "Customer", "Region", "Sales Rep", "Category", "Qty", "Unit Price", "Unit Cost"])
    for l in lines:
        w.writerow([l["date"].isoformat(), l["invoice"], l["customer"], l["region"], l["rep"], l["category"],
                    l["qty"], f"{l['price']:.2f}", f"{l['cost']:.2f}"])
with open(ROOT / "data" / "budget.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Month", "Region", "Sales Rep", "Category", "Budget Revenue", "Budget Margin %"])
    for b in budget:
        w.writerow([b["month"].strftime("%Y-%m"), b["region"], b["rep"], b["category"], f"{b['budget']:.0f}",
                    f"{b['budget_margin']:.2f}"])
json.dump(dict(first_month="2025-01", last_month="2026-10", budget_last_month="2026-12", lines=len(lines),
               invoices=len({l["invoice"] for l in lines}), miss=dict(month="2026-10", region="West", category="Garden")),
          open(ROOT / "test" / "story.json", "w"), indent=1)
print(f"{len(lines)} sales lines, {len({l['invoice'] for l in lines})} invoices, {len(budget)} budget rows")
