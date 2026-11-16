"""Fictional sales export of a garden supply shop (2025 full year, 2026 January-October) with the data problems a real
export has, planted on purpose and listed in data/planted.json:
  - 14 revenue cells exported as text ("1,234.50") in 2026      -> SUM ignores them
  - 9 orders exported twice in 2026                               -> counted twice
  - 6 rows of "Hose 25m " with a trailing space in 2026           -> SUMIF by product misses them
Output: data/sales-2025.csv, data/sales-2026.csv (as exported), data/products.csv, data/planted.json
"""
import csv, json, random
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
D = ROOT / "data"; D.mkdir(exist_ok=True)
rnd = random.Random(1111)

PRODUCTS = [  # category, product, unit price, cost ratio
    ("Seeds & bulbs", "Tomato seeds", 3.50, 0.35), ("Seeds & bulbs", "Tulip bulbs (20)", 12.00, 0.45), ("Seeds & bulbs", "Lawn seed 5kg", 34.00, 0.55),
    ("Plants", "Rose bush", 18.00, 0.50), ("Plants", "Lavender pot", 9.50, 0.45), ("Plants", "Fruit tree", 45.00, 0.55),
    ("Hand tools", "Pruning shears", 22.00, 0.48), ("Hand tools", "Garden spade", 29.00, 0.50), ("Hand tools", "Hand trowel", 8.00, 0.40),
    ("Power tools", "Lawn mower", 289.00, 0.70), ("Power tools", "Hedge trimmer", 119.00, 0.66), ("Power tools", "Leaf blower", 95.00, 0.64),
    ("Watering", "Hose 25m", 39.00, 0.52), ("Watering", "Sprinkler", 16.00, 0.45), ("Watering", "Watering can", 11.00, 0.40),
    ("Soil & feed", "Compost 50L", 7.50, 0.55), ("Soil & feed", "Plant food 1L", 9.00, 0.42), ("Soil & feed", "Bark mulch 70L", 8.50, 0.58),
]
REGIONS = {"North": 1.0, "South": 1.25, "East": 0.9, "West": 1.1}
SEASON = [0.45, 0.55, 0.95, 1.45, 1.6, 1.35, 1.15, 1.0, 0.95, 0.85, 0.6, 0.7]   # spring peak
QTY = lambda price: rnd.choice([1, 1, 2] if price > 80 else [1, 2, 3, 4, 6, 10] if price < 15 else [1, 2, 3, 4])


def year_rows(year, months, growth, start_id):
    rows, oid = [], start_id
    for m in range(1, months + 1):
        days = (date(year + (m == 12), m % 12 + 1, 1) - date(year, m, 1)).days
        n = int(62 * SEASON[m - 1] * growth + rnd.randint(-4, 4))
        for _ in range(n):
            d = date(year, m, rnd.randint(1, days))
            region = rnd.choices(list(REGIONS), weights=list(REGIONS.values()))[0]
            cat, prod, price, cr = rnd.choice(PRODUCTS)
            q = QTY(price)
            rev = round(q * price, 2)
            cost = round(rev * cr * rnd.uniform(0.95, 1.05), 2)
            rows.append([d, region, cat, prod, q, price, rev, cost])
    rows.sort(key=lambda r: r[0])
    out = []
    for r in rows:
        out.append([f"SO-{oid}"] + r); oid += 1
    return out, oid


r25, nxt = year_rows(2025, 12, 1.0, 50001)
r26, _ = year_rows(2026, 10, 1.12, nxt)

# ---- planted problems (2026 only), all inside the first 699 data rows except where said
idx = list(range(40, 690))
rnd.shuffle(idx)
dups = sorted(idx[:9])
texts = sorted(i for i in idx[9:40] if r26[i][4] != "Hose 25m")[:14]
hose = [i for i, r in enumerate(r26) if r[4] == "Hose 25m" and i < 690 and i not in dups and i not in texts][:6]
assert len(hose) == 6 and len(texts) == 14
exported = []
for i, r in enumerate(r26):
    r = list(r)
    if i in hose:
        r[4] += " "
    if i in texts:
        r[7] = f"{r[7]:,.2f}"            # text in the export
    exported.append(r)
    if i in dups:
        exported.append(list(r))          # the same order exported twice, right after itself
HDR = ["Order ID", "Date", "Region", "Category", "Product", "Units", "Unit price", "Revenue", "Cost"]


def write(name, rows):
    with open(D / name, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(HDR)
        for r in rows:
            w.writerow([r[0], r[1].isoformat()] + r[2:])


write("sales-2025.csv", r25)
write("sales-2026.csv", exported)
with open(D / "products.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f); w.writerow(["Category", "Product", "Unit price", "Cost ratio"]); w.writerows(PRODUCTS)
dup_ids = [r26[i][0] for i in dups]
(D / "planted.json").write_text(json.dumps({
    "duplicate_orders": dup_ids, "text_revenue_orders": [r26[i][0] for i in texts],
    "trailing_space_orders": [r26[i][0] for i in hose], "rows_2026_exported": len(exported),
    "rows_2026_unique": len(r26), "rows_2025": len(r25)}, indent=1))
print("2025:", len(r25), "rows · 2026:", len(exported), "rows exported,", len(r26), "unique · sum range H2:H700 misses",
      len(exported) + 1 - 700, "rows")
