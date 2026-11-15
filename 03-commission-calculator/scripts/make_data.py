"""Fictional setup and order lines for "Harborview Office Supplies" (made-up company), Nov 2025 - Oct 2026.

Writes
  ../data/orders.csv    what the sales team typed: one row per order line (some deals split between two reps)
  ../data/returns.csv   goods returned / orders cancelled after they were paid on (clawback)
  ../data/payouts.csv   payout status per rep and month (Pending / Approved / Paid + paid date)
  ../data/plan.json     products, reps + managers, commission bands, tax, bonus, override, accelerator, what-if scenario
  ../test/planted.json  order lines and returns that were made wrong on purpose and must be held back
used by scripts/build_workbook.py and test/verify.py.
"""
import csv, json, random
from datetime import date, timedelta
from pathlib import Path

random.seed(2026_11)
ROOT = Path(__file__).resolve().parent.parent

PLAN = dict(
    company="Harborview Office Supplies",
    tax_rate=0.0825, target_bonus=300, discount_cap=0.30,
    override_rate=0.01,              # a manager earns 1% of the credited net sales of the reps they manage
    accel_at=1.20, accel_rate=0.02,  # +2% extra on the net sales above 120% of target
    first_month="2025-11-01", months=12,
    products=[  # product, category, unit price, max discount
        ("A4 Copy Paper (box)", "Paper", 42.00, 0.10), ("Recycled Paper (box)", "Paper", 46.50, 0.10),
        ("Gel Pens (pack 12)", "Writing", 9.80, 0.15), ("Whiteboard Markers (pack 8)", "Writing", 12.40, 0.15),
        ("Ergonomic Chair", "Furniture", 289.00, 0.20), ("Standing Desk", "Furniture", 549.00, 0.20),
        ("Filing Cabinet 3-Drawer", "Furniture", 199.00, 0.20), ("Laser Printer", "Tech", 379.00, 0.12),
        ("Toner Cartridge", "Tech", 89.90, 0.12), ("Wireless Keyboard & Mouse", "Tech", 64.50, 0.12),
        ("Desk Organizer Set", "Accessories", 24.90, 0.25), ("Monitor Arm", "Accessories", 119.00, 0.25)],
    reps=[  # sales rep, region, monthly target, manager ("" = none)
        ("Grace Holloway", "North", 30000, ""), ("Marcus Reid", "North", 25000, "Grace Holloway"),
        ("Priya Nair", "South", 35000, ""), ("Tom Becker", "South", 20000, "Priya Nair"),
        ("Lena Fischer", "East", 40000, ""), ("Diego Santos", "West", 28000, "")],
    bands=[(0, 0.02), (15000, 0.04), (30000, 0.06), (45000, 0.08)],  # monthly credited net sales from -> rate
    # What-if sheet starts with this scenario: richer low bands, a lower top band, targets +5%, accelerator 3%
    scenario=dict(bands=[(0, 0.025), (15000, 0.045), (30000, 0.06), (45000, 0.07)], target_change=0.05,
                  accel_at=1.20, accel_rate=0.03, target_bonus=250),
)
CUSTOMERS = ["Bayside Dental Group", "Copperline Logistics", "Evergreen Law LLP", "Foxglove Architects",
             "Granite Peak Insurance", "Hollis & Grant Accounting", "Ironwood Schools Trust", "Juniper Health Clinic",
             "Keystone Realty", "Lighthouse Media", "Meadowbrook Council", "Northgate Engineering",
             "Oakridge Veterinary", "Pinecrest Hotels", "Quayside Marine"]
# share of monthly target each rep reaches (net, before splits and planted errors)
#          Grace Marcus Priya  Tom  Lena  Diego
REACH = [[0.85, 0.95, 0.78, 0.72, 1.05, 0.90],   # Nov 2025
         [1.10, 0.80, 0.92, 0.65, 1.28, 0.97],   # Dec
         [0.70, 0.75, 0.68, 0.60, 0.88, 0.72],   # Jan 2026
         [0.82, 0.90, 0.85, 0.78, 0.98, 0.86],   # Feb
         [1.02, 1.12, 0.95, 0.82, 1.10, 1.04],   # Mar
         [0.95, 1.00, 1.01, 0.74, 1.06, 0.93],   # Apr
         [1.26, 0.92, 0.99, 0.88, 1.12, 1.08],   # May
         [0.99, 1.05, 1.13, 0.69, 1.24, 0.96],   # Jun
         [0.88, 0.97, 0.90, 0.80, 1.01, 1.27],   # Jul
         [0.93, 1.03, 0.94, 0.71, 1.09, 1.00],   # Aug
         [0.92, 1.08, 0.80, 0.70, 1.15, 0.95],   # Sep
         [1.05, 0.88, 0.89, 0.66, 1.27, 1.02]]   # Oct 2026
MONTHS = [date(2025 + (10 + i) // 12, (10 + i) % 12 + 1, 1) for i in range(12)]
REPS = [r[0] for r in PLAN["reps"]]


def month_days(m):
    return ((m.replace(day=28) + timedelta(days=4)).replace(day=1) - m).days


lines, order_no = [], 6001
for month, reach in zip(MONTHS, REACH):
    for (rep, region, target, mgr), share in zip(PLAN["reps"], reach):
        goal, done = target * share, 0.0
        while done < goal:
            d = month + timedelta(days=random.randrange(month_days(month)))
            cust = random.choice(CUSTOMERS)
            split_rep, split_pct = "", 0.0
            if random.random() < 0.06:  # a deal won together with a colleague
                split_rep = random.choice([r for r in REPS if r != rep])
                split_pct = random.choice([0.5, 0.5, 0.3, 0.4])
            for _ in range(random.choice([1, 1, 2, 2, 3])):
                prod, cat, price, maxd = random.choice(PLAN["products"])
                qty = random.choice([1, 2, 2, 3, 4, 5, 6, 8, 10]) * (5 if cat in ("Paper", "Writing") else 1)
                disc = random.choice([0, 0, 0, 0.05, 0.05, round(maxd, 2)])
                net = round(round(qty * price, 2) - round(qty * price * disc, 2), 2)
                if done + net > goal * 1.03 and done > goal * 0.9:
                    continue
                if any(l["order"] == f"SO-{order_no}" and l["prod"] == prod for l in lines[-3:]):
                    continue  # one line per product per order (returns point at order + product)
                lines.append(dict(date=d, order=f"SO-{order_no}", rep=rep, cust=cust, prod=prod, qty=qty, disc=disc,
                                  split_rep=split_rep, split_pct=split_pct))
                done += net * (1 - split_pct)
            order_no += 1

# Mistakes a real sales team makes, planted on purpose: these lines must be held back, not paid on.
P = lambda d, rep, cust, prod, qty, disc, why, sr="", sp=0.0: dict(date=d, rep=rep, cust=cust, prod=prod, qty=qty, disc=disc,
                                                                   split_rep=sr, split_pct=sp, why=why)
planted = [
    P(date(2026, 10, 14), "Diego Santos", "Pinecrest Hotels", "A4 Copy Paper (box)", 40, 0.25, "Discount above 10% max"),
    P(date(2026, 10, 21), "Priya Nair", "Keystone Realty", "Standing Desk", 0, 0, "Qty must be 1 or more"),
    P(date(2026, 10, 27), "Lena Fischer", "Lighthouse Media", "Desk Lamp LED", 6, 0, "Unknown product"),
    P(date(2026, 10, 8), "Marcus Reid", "Oakridge Veterinary", "Monitor Arm", 4, 0, "Split rep same as sales rep", "Marcus Reid", 0.5),
    P(date(2026, 9, 18), "Tom Baker", "Bayside Dental Group", "Ergonomic Chair", 3, 0, "Unknown sales rep"),
    P(date(2026, 9, 23), "Tom Becker", "Quayside Marine", "Laser Printer", 2, 0, "Split % must be 1% to 99%", "Priya Nair", 0.0),
]
for i, p in enumerate(planted):
    p["order"] = f"SO-{order_no + i}"
lines += [{k: v for k, v in p.items() if k != "why"} for p in planted]
lines.sort(key=lambda l: (l["date"], l["order"]))

# Returns: goods sent back after the rep was paid. The commission paid on them is clawed back in the month of the return.
rng = random.Random(77)
ok_lines = [l for l in lines if l["rep"] in REPS and l not in planted and l["qty"] >= 2 and l["date"] < date(2026, 10, 1)
            and l["prod"] in {p[0] for p in PLAN["products"]} and not any(l["order"] == p["order"] for p in planted)]
picked = rng.sample([l for l in ok_lines if not l["split_rep"]], 10) + rng.sample([l for l in ok_lines if l["split_rep"]], 2)
REASONS = ["Damaged in delivery", "Customer cancelled", "Wrong item ordered", "Returned unused"]
returns = []
for l in picked:
    rd = l["date"] + timedelta(days=rng.randrange(8, 40))
    rd = min(rd, date(2026, 10, 30))
    returns.append(dict(date=rd, order=l["order"], prod=l["prod"], qty=rng.randrange(1, l["qty"] + 1) if l["qty"] < 6 else rng.randrange(1, 6),
                        reason=rng.choice(REASONS)))
some = next(l for l in ok_lines if l["date"].month == 9 and l["qty"] < 5 and l not in picked)
bad_returns = [  # planted: must not be clawed back
    dict(date=date(2026, 10, 19), order=some["order"], prod=some["prod"], qty=some["qty"] + 3, reason="Customer cancelled",
         why="Qty above qty sold"),
    dict(date=date(2026, 10, 24), order="SO-5999", prod="Laser Printer", qty=1, reason="Damaged in delivery",
         why="Order line not found"),
]
returns += [{k: v for k, v in b.items() if k != "why"} for b in bad_returns]
returns.sort(key=lambda r: (r["date"], r["order"]))

# Payouts: everything up to August is paid (on the 5th-7th of the next month); September is part paid, part approved;
# October was just calculated, so it is pending approval.
payouts = []
for m in MONTHS:
    nxt = (m + timedelta(days=32)).replace(day=1)
    for i, rep in enumerate(REPS):
        if m < date(2026, 9, 1):
            payouts.append((m, rep, "Paid", nxt + timedelta(days=4 + (i + m.month) % 3)))
        elif m == date(2026, 9, 1):
            payouts.append((m, rep, "Paid", date(2026, 10, 6)) if i < 4 else (m, rep, "Approved", None))
        else:
            payouts.append((m, rep, "Pending", None))

(ROOT / "data").mkdir(exist_ok=True); (ROOT / "test").mkdir(exist_ok=True)
with open(ROOT / "data" / "orders.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Date", "Order no", "Sales rep", "Customer", "Product", "Qty", "Discount %", "Split rep", "Split %"])
    for l in lines:
        w.writerow([l["date"].isoformat(), l["order"], l["rep"], l["cust"], l["prod"], l["qty"], f"{l['disc']:.2f}",
                    l["split_rep"], f"{l['split_pct']:.2f}" if l["split_rep"] else ""])
with open(ROOT / "data" / "returns.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Return date", "Order no", "Product", "Qty returned", "Reason"])
    for r in returns:
        w.writerow([r["date"].isoformat(), r["order"], r["prod"], r["qty"], r["reason"]])
with open(ROOT / "data" / "payouts.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Month", "Sales rep", "Status", "Paid date"])
    for m, rep, st, pd in payouts:
        w.writerow([m.isoformat(), rep, st, pd.isoformat() if pd else ""])
json.dump(PLAN, open(ROOT / "data" / "plan.json", "w"), indent=1)
json.dump(dict(orders=[dict(order=p["order"], why=p["why"]) for p in planted],
               returns=[dict(order=b["order"], why=b["why"]) for b in bad_returns]),
          open(ROOT / "test" / "planted.json", "w"), indent=1)
print(f"{len(lines)} order lines, {len({l['order'] for l in lines})} orders, "
      f"{sum(1 for l in lines if l['split_rep'])} split lines, {len(returns)} returns, {len(planted)} + {len(bad_returns)} planted mistakes")
