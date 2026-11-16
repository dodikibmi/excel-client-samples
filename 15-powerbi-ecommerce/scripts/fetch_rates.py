"""Saves the daily ECB exchange rates (1 USD = x EUR / GBP / CAD / AUD) from the free Frankfurter API to
data/fx-rates-saved.csv: the copy the report falls back on when the API cannot be reached during a refresh, and the
rates test/verify.py and scripts/make_data.py use.  Run:  python scripts/fetch_rates.py
"""
import csv, json, urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
URL = "https://api.frankfurter.dev/v1/2024-12-01..?base=USD&symbols=EUR,GBP,CAD,AUD"

with urllib.request.urlopen(urllib.request.Request(URL, headers={"User-Agent": "sample-report"}), timeout=60) as r:
    data = json.load(r)
rows = [(d, c, v) for d, rates in sorted(data["rates"].items()) for c, v in sorted(rates.items())]
with open(ROOT / "data" / "fx-rates-saved.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Date", "Currency", "Units per USD"])
    w.writerows(rows)
print(f"fx-rates-saved.csv: {len(rows)} rates, {data['start_date']} to {data['end_date']}")
