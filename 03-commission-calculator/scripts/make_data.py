"""Fictional setup and order lines for "Harborview Office Supplies" (made-up company), Sep - Oct 2026.

Writes ../data/orders.csv (what the sales team typed: one row per order line) and ../data/plan.json
(products, sales reps, commission bands, tax, bonus) used by scripts/build_workbook.py, plus
../test/planted.json with the order lines that were made wrong on purpose and must be held back.
"""
import csv, json, random
from datetime import date, timedelta
from pathlib import Path

random.seed(2026_11)
ROOT = Path(__file__).resolve().parent.parent

PLAN = dict(
    company="Harborview Office Supplies",
    tax_rate=0.0825, target_bonus=300, discount_cap=0.30,
    products=[  # product, category, unit price, max discount
        ("A4 Copy Paper (box)", "Paper", 42.00, 0.10), ("Recycled Paper (box)", "Paper", 46.50, 0.10),
        ("Gel Pens (pack 12)", "Writing", 9.80, 0.15), ("Whiteboard Markers (pack 8)", "Writing", 12.40, 0.15),
        ("Ergonomic Chair", "Furniture", 289.00, 0.20), ("Standing Desk", "Furniture", 549.00, 0.20),
        ("Filing Cabinet 3-Drawer", "Furniture", 199.00, 0.20), ("Laser Printer", "Tech", 379.00, 0.12),
        ("Toner Cartridge", "Tech", 89.90, 0.12), ("Wireless Keyboard & Mouse", "Tech", 64.50, 0.12),
        ("Desk Organizer Set", "Accessories", 24.90, 0.25), ("Monitor Arm", "Accessories", 119.00, 0.25)],
    reps=[  # sales rep, region, monthly target
        ("Grace Holloway", "North", 30000), ("Marcus Reid", "North", 25000), ("Priya Nair", "South", 35000),
        ("Tom Becker", "South", 20000), ("Lena Fischer", "East", 40000), ("Diego Santos", "West", 28000)],
    bands=[(0, 0.02), (15000, 0.04), (30000, 0.06), (45000, 0.08)],  # monthly net sales from -> rate
)
CUSTOMERS = ["Bayside Dental Group", "Copperline Logistics", "Evergreen Law LLP", "Foxglove Architects",
             "Granite Peak Insurance", "Hollis & Grant Accounting", "Ironwood Schools Trust", "Juniper Health Clinic",
             "Keystone Realty", "Lighthouse Media", "Meadowbrook Council", "Northgate Engineering",
             "Oakridge Veterinary", "Pinecrest Hotels", "Quayside Marine"]
# share of monthly target each rep reaches (net, before planted errors): spreads reps over all commission bands
REACH = {date(2026, 9, 1): [0.92, 1.08, 0.80, 0.70, 1.15, 0.95], date(2026, 10, 1): [1.05, 0.88, 0.89, 0.66, 1.20, 1.02]}

lines, order_no = [], 7001
for month, reach in REACH.items():
    days = 30 if month.month == 9 else 31
    for (rep, region, target), share in zip(PLAN["reps"], reach):
        goal, done = target * share, 0.0
        while done < goal:
            d = month + timedelta(days=random.randrange(days))
            cust = random.choice(CUSTOMERS)
            for _ in range(random.choice([1, 1, 2, 2, 3])):
                prod, cat, price, maxd = random.choice(PLAN["products"])
                qty = random.choice([1, 2, 2, 3, 4, 5, 6, 8, 10]) * (5 if cat in ("Paper", "Writing") else 1)
                disc = random.choice([0, 0, 0, 0.05, 0.05, round(maxd, 2)])
                net = round(round(qty * price, 2) - round(qty * price * disc, 2), 2)
                if done + net > goal * 1.03 and done > goal * 0.9:
                    continue
                lines.append(dict(date=d, order=f"SO-{order_no}", rep=rep, cust=cust, prod=prod, qty=qty, disc=disc))
                done += net
            order_no += 1

# Mistakes a real sales team makes, planted on purpose: these lines must be held back, not paid on.
planted = [
    dict(date=date(2026, 10, 14), order=f"SO-{order_no}", rep="Diego Santos", cust="Pinecrest Hotels",
         prod="A4 Copy Paper (box)", qty=40, disc=0.25, why="Discount above 10% max"),
    dict(date=date(2026, 10, 21), order=f"SO-{order_no + 1}", rep="Priya Nair", cust="Keystone Realty",
         prod="Standing Desk", qty=0, disc=0, why="Qty must be 1 or more"),
    dict(date=date(2026, 10, 27), order=f"SO-{order_no + 2}", rep="Lena Fischer", cust="Lighthouse Media",
         prod="Desk Lamp LED", qty=6, disc=0, why="Unknown product"),
    dict(date=date(2026, 9, 18), order=f"SO-{order_no + 3}", rep="Tom Baker", cust="Bayside Dental Group",
         prod="Ergonomic Chair", qty=3, disc=0, why="Unknown sales rep"),
]
lines += [{k: v for k, v in p.items() if k != "why"} for p in planted]
lines.sort(key=lambda l: (l["date"], l["order"]))

(ROOT / "data").mkdir(exist_ok=True); (ROOT / "test").mkdir(exist_ok=True)
with open(ROOT / "data" / "orders.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Date", "Order no", "Sales rep", "Customer", "Product", "Qty", "Discount %"])
    for l in lines:
        w.writerow([l["date"].isoformat(), l["order"], l["rep"], l["cust"], l["prod"], l["qty"], f"{l['disc']:.2f}"])
json.dump(PLAN, open(ROOT / "data" / "plan.json", "w"), indent=1)
json.dump([dict(order=p["order"], why=p["why"]) for p in planted], open(ROOT / "test" / "planted.json", "w"), indent=1)
print(f"{len(lines)} order lines, {len({l['order'] for l in lines})} orders, {len(planted)} planted mistakes")
