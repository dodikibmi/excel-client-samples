"""Fictional monthly sales for "Harbourline Supply Co." (made-up wholesaler): Nov 2024 - Oct 2026,
4 regions x 5 categories, revenue, cost, orders and budget. Writes ../data/sales.csv.
The same numbers sit behind all 8 themed dashboards, so the client compares colours, not data."""
import csv, random
from pathlib import Path

random.seed(10_2026)
ROOT = Path(__file__).resolve().parent.parent
MONTHS = [(2024 + (10 + k) // 12, (10 + k) % 12 + 1) for k in range(24)]  # Nov 2024 .. Oct 2026
REGIONS = {"North": 1.00, "South": 0.86, "East": 1.14, "West": 0.94}
CATEGORIES = {  # base revenue per region-month, margin, seasonality Jan..Dec
    "Fasteners": (21000, 0.38, [1.0, 1.0, 1.05, 1.05, 1.05, 1.0, 0.95, 0.95, 1.0, 1.05, 1.0, 0.9]),
    "Power Tools": (26000, 0.24, [0.8, 0.85, 0.95, 1.0, 1.0, 0.95, 0.9, 0.95, 1.0, 1.05, 1.2, 1.4]),
    "Paint": (15000, 0.42, [0.7, 0.75, 0.9, 1.1, 1.25, 1.3, 1.3, 1.2, 1.05, 0.9, 0.8, 0.7]),
    "Garden": (12000, 0.35, [0.4, 0.5, 0.9, 1.5, 1.8, 1.6, 1.3, 1.1, 0.9, 0.7, 0.5, 0.4]),
    "Electrical": (17000, 0.31, [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.05, 1.05, 1.0]),
}

rows = []
for i, (y, m) in enumerate(MONTHS):
    growth = 1 + 0.006 * i
    for reg, rf in REGIONS.items():
        for cat, (base, margin, season) in CATEGORIES.items():
            rev = base * rf * season[m - 1] * growth * random.uniform(0.93, 1.07)
            bud = rev * random.uniform(0.96, 1.07)
            if (y, m) == (2026, 10) and reg == "West":  # the planted story: West misses budget in October
                rev *= 0.86
            cost = rev * (1 - margin * random.uniform(0.95, 1.05))
            orders = max(1, round(rev / random.uniform(640, 780)))
            rows.append([f"{y}-{m:02d}-01", reg, cat, round(rev), round(bud), round(cost), orders])

out = ROOT / "data" / "sales.csv"
with out.open("w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Month", "Region", "Category", "Revenue", "Budget", "Cost", "Orders"])
    w.writerows(rows)
print(out, len(rows), "rows")
