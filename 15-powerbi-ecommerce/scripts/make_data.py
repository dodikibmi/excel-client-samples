"""Writes the fictional input files for E15 (Larkspur Home, an online home goods store) and the planted answer.

data/shop-orders-export.csv       webshop order export, Shopify layout: one row per order line, the order fields
                                  (status, currency, totals, discount, refund) only on the first line of each order
data/marketplace-orders-eu.csv    EU marketplace export: semicolons, dd.mm.yyyy, decimal commas, monthly exports
                                  pasted together (the header repeats, 4 rows appear twice)
data/products.csv                 SKU, product, category, unit cost in USD
data/ads-spend.csv                monthly ad spend per channel (USD)
test/expected.json                the true numbers, worked out here from the generated orders (not from the files)

Exchange rates: data/fx-rates-saved.csv (run scripts/fetch_rates.py first): 1 USD = x units, ECB via Frankfurter.
Run:  python scripts/make_data.py
"""
import csv, json, math, random
from bisect import bisect_right
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA, TEST = ROOT / "data", ROOT / "test"
rnd = random.Random(15)
START, END = date(2025, 1, 1), date(2026, 10, 31)

PRODUCTS = [  # sku, name, category, USD price, unit cost USD, popularity
    ("LH-CN-001", "Amber & Oak Candle 8 oz", "Candles", 28, 7.40, 10),
    ("LH-CN-002", "Fig Leaf Candle 8 oz", "Candles", 28, 7.40, 8),
    ("LH-CN-003", "Sea Salt Candle 14 oz", "Candles", 42, 11.20, 6),
    ("LH-CN-004", "Travel Candle Trio", "Candles", 36, 9.80, 7),
    ("LH-CN-005", "Reed Diffuser Cedar", "Candles", 38, 10.60, 5),
    ("LH-CR-001", "Stoneware Mug Sand", "Ceramics", 24, 6.10, 9),
    ("LH-CR-002", "Stoneware Mug Moss", "Ceramics", 24, 6.10, 7),
    ("LH-CR-003", "Serving Bowl Large", "Ceramics", 58, 17.50, 4),
    ("LH-CR-004", "Bud Vase Set of 3", "Ceramics", 46, 12.90, 5),
    ("LH-CR-005", "Dinner Plate Set of 4", "Ceramics", 84, 27.00, 3),
    ("LH-LN-001", "Linen Napkins Set of 4", "Linens", 34, 9.40, 6),
    ("LH-LN-002", "Linen Tablecloth Oat", "Linens", 89, 26.50, 3),
    ("LH-LN-003", "Waffle Bath Towel", "Linens", 44, 12.00, 5),
    ("LH-LN-004", "Linen Throw Pillow Cover", "Linens", 39, 10.20, 5),
    ("LH-LN-005", "Tea Towel Pair", "Linens", 22, 5.30, 6),
    ("LH-KT-001", "Olive Wood Board", "Kitchen", 52, 15.80, 4),
    ("LH-KT-002", "Brass Measuring Spoons", "Kitchen", 26, 7.10, 4),
    ("LH-KT-003", "Glass Storage Jar Set", "Kitchen", 48, 14.20, 4),
    ("LH-KT-004", "Pour-Over Coffee Set", "Kitchen", 64, 21.50, 3),
    ("LH-GS-001", "Cozy Evening Gift Set", "Gift Sets", 72, 22.40, 3),
    ("LH-GS-002", "Host Gift Box", "Gift Sets", 96, 31.00, 2),
    ("LH-GS-003", "Mini Candle Advent Box", "Gift Sets", 118, 38.50, 1),
]
SKU = {p[0]: p for p in PRODUCTS}
FACTOR = {"USD": 1.0, "EUR": 0.95, "GBP": 0.82, "CAD": 1.38, "AUD": 1.55}  # fixed local price lists


def local_price(usd, cur):
    return float(usd) if cur == "USD" else math.floor(usd * FACTOR[cur]) + 0.90


MARKETS = {  # country: currency, share of webshop orders, sales tax, growth 2026 vs 2025
    "United States": ("USD", 0.44, 0.07, 1.06), "Canada": ("CAD", 0.10, 0.13, 1.04),
    "United Kingdom": ("GBP", 0.15, 0.20, 1.30), "Germany": ("EUR", 0.08, 0.19, 1.38), "France": ("EUR", 0.05, 0.20, 1.34),
    "Netherlands": ("EUR", 0.04, 0.21, 1.36), "Ireland": ("EUR", 0.03, 0.23, 1.30), "Australia": ("AUD", 0.11, 0.10, 1.18),
}
SHIP = {"USD": 6.95, "CAD": 9.95, "GBP": 4.95, "EUR": 5.95, "AUD": 11.95}
FREE_OVER = {"USD": 75, "CAD": 100, "GBP": 60, "EUR": 70, "AUD": 110}
SEASON = {1: .78, 2: .84, 3: .95, 4: .97, 5: 1.08, 6: .9, 7: .86, 8: .9, 9: 1.0, 10: 1.12, 11: 1.62, 12: 1.95}
WEEKDAY = [0.95, 0.97, 1.0, 1.02, 1.05, 1.1, 0.91]
MP_COUNTRIES = [("Germany", .35), ("France", .25), ("Italy", .20), ("Spain", .12), ("Austria", .08)]

# ---------------------------------------------------------------------------------------- exchange rates
rates = defaultdict(list)  # currency -> [(date, units per usd)]
with open(DATA / "fx-rates-saved.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        rates[r["Currency"]].append((date.fromisoformat(r["Date"]), float(r["Units per USD"])))
for c in rates:
    rates[c].sort()


def rate_on(cur, d):
    if cur == "USD":
        return 1.0, d
    lst = rates[cur]
    i = bisect_right([x[0] for x in lst], d) - 1
    return lst[i][1], lst[i][0]


def month_avg(cur, y, m):
    if cur == "USD":
        return 1.0
    v = [r for d, r in rates[cur] if d.year == y and d.month == m]
    return sum(v) / len(v)


def poisson(lam):
    l, k, p = math.exp(-lam), 0, 1.0
    while True:
        p *= rnd.random()
        if p <= l:
            return k
        k += 1


def pick(weighted):
    t = rnd.random() * sum(w for _, w in weighted)
    for v, w in weighted:
        t -= w
        if t <= 0:
            return v
    return weighted[-1][0]


# ---------------------------------------------------------------------------------------- webshop orders
orders = []      # dicts, chronological
cust_by_country = defaultdict(list)
cust_country = {}
next_cust = 1
order_no = 1001
d = START
while d <= END:
    season = SEASON[d.month] * WEEKDAY[d.weekday()]
    bfcm = date(2025, 11, 27) <= d <= date(2025, 12, 1)
    for country, (cur, share, tax, growth) in MARKETS.items():
        lam = 11.0 * share * season * (growth ** ((d - START).days / 365)) * (2.4 if bfcm else 1.0)
        for _ in range(poisson(lam)):
            existing = cust_by_country[country]
            if existing and rnd.random() < 0.30:
                cust = rnd.choice(existing)
            else:
                cust = f"c{next_cust:05d}@example.com"
                next_cust += 1
                existing.append(cust)
                cust_country[cust] = country
            orders.append(dict(day=d, country=country, cur=cur, tax=tax, cust=cust))
    d += timedelta(days=1)
for o in orders:
    o["time"] = datetime(o["day"].year, o["day"].month, o["day"].day, rnd.randint(6, 23), rnd.randint(0, 59), rnd.randint(0, 59))
orders.sort(key=lambda o: o["time"])
for o in orders:   # order numbers follow the time the order was placed
    o["no"] = f"#{order_no}"
    order_no += 1

seen_customers = set()
for o in orders:
    cur, d = o["cur"], o["day"]
    k = pick([(1, 55), (2, 28), (3, 12), (4, 5)])
    gift = 3.0 if d.month in (11, 12) else 1.0
    pool = [(p[0], p[5] * (gift if p[2] == "Gift Sets" else 1.0)) for p in PRODUCTS]
    skus = []
    while len(skus) < k:
        s = pick(pool)
        if s not in skus:
            skus.append(s)
    lines = []
    for s in skus:
        q = pick([(1, 85), (2, 12), (3, 3)])
        lines.append(dict(sku=s, qty=q, price=local_price(SKU[s][3], cur)))
    o["lines"] = lines
    gross = round(sum(l["qty"] * l["price"] for l in lines), 2)
    new = o["cust"] not in seen_customers
    code, pct = "", 0.0
    if date(2025, 11, 27) <= d <= date(2025, 12, 1) and rnd.random() < 0.65:
        code, pct = "BFCM25", 0.25
    elif date(2026, 4, 10) <= d <= date(2026, 4, 20) and rnd.random() < 0.30:
        code, pct = "SPRING15", 0.15
    elif new and rnd.random() < 0.22:
        code, pct = "WELCOME10", 0.10
    disc = round(gross * pct, 2)
    sub = round(gross - disc, 2)
    ship = 0.0 if sub >= FREE_OVER[cur] else SHIP[cur]
    taxes = round(sub * o["tax"], 2)
    total = round(sub + ship + taxes, 2)
    r = rnd.random()
    status = "voided" if r < 0.010 else "refunded" if r < 0.035 else "partially_refunded" if r < 0.065 else "paid"
    refunded = 0.0
    if status == "refunded":
        refunded = total
    elif status == "partially_refunded":
        l = rnd.choice(lines)
        refunded = round(l["qty"] * l["price"] * (sub / gross) * (1 + o["tax"]), 2)
    o.update(gross=gross, code=code, disc=disc, sub=sub, ship=ship, taxes=taxes, total=total, status=status, refunded=refunded)
    if status != "voided":
        seen_customers.add(o["cust"])

# three free fabric swatches added to linen orders: a SKU that is not in the product list
swatch_orders = [o for o in orders if o["status"] == "paid" and any(SKU[l["sku"]][2] == "Linens" for l in o["lines"])]
for o in rnd.sample(swatch_orders, 3):
    o["lines"].append(dict(sku="LH-SWATCH", qty=1, price=0.0, swatch=True))

PRODUCT_NAME = {p[0]: p[1] for p in PRODUCTS} | {"LH-SWATCH": "Fabric swatch (free)"}


def shop_time(t):
    dst = date(t.year, 3, 9) <= t.date() < date(t.year, 11, 2)
    return t.strftime("%Y-%m-%d %H:%M:%S ") + ("-0400" if dst else "-0500")


messy = 0
with open(DATA / "shop-orders-export.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Name", "Email", "Financial Status", "Paid at", "Fulfillment Status", "Currency", "Subtotal", "Shipping", "Taxes",
                "Total", "Discount Code", "Discount Amount", "Created at", "Lineitem quantity", "Lineitem name", "Lineitem price",
                "Lineitem sku", "Shipping Country", "Refunded Amount"])
    for o in orders:
        for i, l in enumerate(o["lines"]):
            sku = l["sku"]
            r = rnd.random()
            if r < 0.02:
                sku, messy = sku.lower(), messy + 1   # typed by hand in the admin
            elif r < 0.04:
                sku, messy = sku + " ", messy + 1
            line = [l["qty"], PRODUCT_NAME[l["sku"]], f"{l['price']:.2f}", sku]
            if i == 0:
                paid = "" if o["status"] == "voided" else shop_time(o["time"] + timedelta(minutes=rnd.randint(0, 3)))
                w.writerow([o["no"], o["cust"], o["status"], paid, "unfulfilled" if o["status"] == "voided" else "fulfilled", o["cur"],
                            f"{o['sub']:.2f}", f"{o['ship']:.2f}", f"{o['taxes']:.2f}", f"{o['total']:.2f}", o["code"],
                            f"{o['disc']:.2f}", shop_time(o["time"])] + line + [o["country"], f"{o['refunded']:.2f}"])
            else:
                w.writerow([o["no"], o["cust"]] + [""] * 11 + line + ["", ""])

# ---------------------------------------------------------------------------------------- marketplace orders (EUR)
mp = []
mp_customers = []
seq = 1
d = date(2025, 3, 1)
while d <= END:
    lam = 2.6 * SEASON[d.month] * (1.0 + 0.6 * (d - date(2025, 3, 1)).days / 600)
    for _ in range(poisson(lam)):
        if mp_customers and rnd.random() < 0.10:
            buyer = rnd.choice(mp_customers)
        else:
            buyer = f"MP-{rnd.randint(100000, 999999)}"
            mp_customers.append(buyer)
        country = pick(MP_COUNTRIES)
        oid = f"EU-{d.strftime('%y%m')}{seq:05d}"
        seq += 1
        n = pick([(1, 75), (2, 25)])
        skus = rnd.sample([p[0] for p in PRODUCTS if p[2] != "Gift Sets"], n)
        lines = [dict(sku=s, qty=pick([(1, 90), (2, 10)]), price=local_price(SKU[s][3], "EUR"),
                      status="Returned" if rnd.random() < 0.05 else "Shipped") for s in skus]
        mp.append(dict(no=oid, day=d, cust=buyer, country=country, cur="EUR", lines=lines))
    d += timedelta(days=1)


def eu(x):  # 1234.5 -> "1.234,50"
    s = f"{x:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


header = ["Order ID", "Order date", "Buyer ID", "SKU", "Qty", "Item price", "Line total", "Country", "Status"]
mp_rows, dup_rows, header_repeats = [], 0, 0
by_month = defaultdict(list)
for o in mp:
    for l in o["lines"]:
        by_month[(o["day"].year, o["day"].month)].append(
            [o["no"], o["day"].strftime("%d.%m.%Y"), o["cust"], l["sku"], str(l["qty"]), eu(l["price"]), eu(l["qty"] * l["price"]),
             o["country"], l["status"]])
months_mp = sorted(by_month)
dup_in = set(rnd.sample(months_mp[1:], 4))   # 4 monthly exports start with the last row of the month before
with open(DATA / "marketplace-orders-eu.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f, delimiter=";")
    prev_last = None
    for i, m in enumerate(months_mp):
        w.writerow(header)
        if i:
            header_repeats += 1
        if m in dup_in:
            w.writerow(prev_last)
            dup_rows += 1
        for row in by_month[m]:
            w.writerow(row)
        prev_last = by_month[m][-1]

# ---------------------------------------------------------------------------------------- products, ads
with open(DATA / "products.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["SKU", "Product", "Category", "List Price USD", "Unit Cost USD"])
    for p in PRODUCTS:
        w.writerow([p[0], p[1], p[2], f"{p[3]:.2f}", f"{p[4]:.2f}"])

# ---------------------------------------------------------------------------------------- the planted answer
def mo(d):
    return d.year * 100 + d.month


facts = []   # one per sold line (voided orders left out)
order_facts = []
for o in orders:
    if o["status"] == "voided":
        continue
    rate, rate_date = rate_on(o["cur"], o["day"])
    returns = o["sub"] if o["status"] == "refunded" else o["refunded"] * o["sub"] / (o["sub"] + o["taxes"]) if o["refunded"] else 0.0
    order_facts.append(dict(no=o["no"], channel="Webshop", day=o["day"], cust=o["cust"], cur=o["cur"], country=o["country"], rate=rate,
                            rate_age=(o["day"] - rate_date).days, ship=o["ship"], taxes=o["taxes"], total=o["total"], refunded=o["refunded"]))
    for l in o["lines"]:
        g = l["qty"] * l["price"]
        ld = o["disc"] * g / o["gross"]
        lr = returns * (g - ld) / o["sub"]
        cost = 0.0 if o["status"] == "refunded" or l.get("swatch") else l["qty"] * SKU[l["sku"]][4]
        facts.append(dict(no=o["no"], channel="Webshop", day=o["day"], cur=o["cur"], country=o["country"], sku=l["sku"], qty=l["qty"],
                          gross=g, disc=ld, ret=lr, cogs=cost, rate=rate))
for o in mp:
    rate, rate_date = rate_on("EUR", o["day"])
    order_facts.append(dict(no=o["no"], channel="Marketplace EU", day=o["day"], cust=o["cust"], cur="EUR", country=o["country"], rate=rate,
                            rate_age=(o["day"] - rate_date).days, ship=0.0, taxes=0.0, total=sum(l["qty"] * l["price"] for l in o["lines"]),
                            refunded=0.0))
    for l in o["lines"]:
        g = l["qty"] * l["price"]
        back = l["status"] == "Returned"
        facts.append(dict(no=o["no"], channel="Marketplace EU", day=o["day"], cur="EUR", country=o["country"], sku=l["sku"], qty=l["qty"],
                          gross=g, disc=0.0, ret=g if back else 0.0, cogs=0.0 if back else l["qty"] * SKU[l["sku"]][4], rate=rate))

for x in facts:
    x["net_local"] = x["gross"] - x["disc"] - x["ret"]
    x["net"] = x["net_local"] / x["rate"]
    x["m"] = mo(x["day"])

# customers: first order (date, then order number) decides new vs returning
first = {}
for o in sorted(order_facts, key=lambda o: (o["day"], o["channel"], int(o["no"].lstrip("#").replace("EU-", "")))):
    first.setdefault(o["cust"], o)
for o in order_facts:
    f0 = first[o["cust"]]
    o["new"] = f0 is o
    o["since"] = (o["day"].year - f0["day"].year) * 12 + o["day"].month - f0["day"].month

ads = []
months = sorted({x["m"] for x in facts})
net_shop = defaultdict(float)
for x in facts:
    if x["channel"] == "Webshop":
        net_shop[x["m"]] += x["net"]
for m in months:
    total = net_shop[m] * rnd.uniform(0.10, 0.14) * (1.25 if m % 100 in (11, 12) else 1.0)
    meta = round(total * 0.6, 2)
    ads += [(f"{m // 100}-{m % 100:02d}", "Meta Ads", meta), (f"{m // 100}-{m % 100:02d}", "Google Ads", round(total - meta, 2))]
with open(DATA / "ads-spend.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Month", "Channel", "Spend USD"])
    for a in ads:
        w.writerow([a[0], a[1], f"{a[2]:.2f}"])
ad_by_month = defaultdict(float)
for a in ads:
    ad_by_month[int(a[0].replace("-", ""))] += a[2]

exp = {"months": {}, "currencies": {}, "products": {}, "markets": {}, "channels": {}, "cohorts": {}, "fx": {}, "checks": {}}
for m in months:
    fs = [x for x in facts if x["m"] == m]
    os_ = [o for o in order_facts if mo(o["day"]) == m]
    net = sum(x["net"] for x in fs)
    cogs = sum(x["cogs"] for x in fs)
    e = dict(orders=len(os_), units=sum(x["qty"] for x in fs), gross=sum(x["gross"] / x["rate"] for x in fs),
             discounts=sum(x["disc"] / x["rate"] for x in fs), returns=sum(x["ret"] / x["rate"] for x in fs), net=net, cogs=cogs,
             gross_profit=net - cogs, new_customers=sum(o["new"] for o in os_), returning_orders=sum(not o["new"] for o in os_),
             shipping=sum(o["ship"] / o["rate"] for o in os_), taxes=sum(o["taxes"] / o["rate"] for o in os_), ad_spend=ad_by_month[m])
    e["aov"] = (e["gross"] - e["discounts"]) / e["orders"]
    y, mm = divmod(m, 100)
    if y >= 2026:
        e["net_cc"] = sum(x["net_local"] / month_avg(x["cur"], y - 1, mm) for x in fs)
        e["fx_effect"] = net - e["net_cc"]
    exp["months"][str(m)] = e
for cur in FACTOR:
    exp["currencies"][cur] = {str(m): dict(local=sum(x["net_local"] for x in facts if x["cur"] == cur and x["m"] == m),
                                           usd=sum(x["net"] for x in facts if x["cur"] == cur and x["m"] == m)) for m in months}
for s in list(SKU) + ["LH-SWATCH"]:
    fs = [x for x in facts if x["sku"] == s]
    exp["products"][s] = dict(units=sum(x["qty"] for x in fs), net=sum(x["net"] for x in fs), cogs=sum(x["cogs"] for x in fs),
                              net_2025=sum(x["net"] for x in fs if x["m"] < 202600 and x["m"] % 100 <= 10),
                              net_2026=sum(x["net"] for x in fs if x["m"] > 202600))
for c in sorted({x["country"] for x in facts}):
    exp["markets"][c] = dict(net_2025_ytd=sum(x["net"] for x in facts if x["country"] == c and 202500 < x["m"] <= 202510),
                             net_2026_ytd=sum(x["net"] for x in facts if x["country"] == c and x["m"] > 202600))
for ch in ("Webshop", "Marketplace EU"):
    exp["channels"][ch] = dict(net=sum(x["net"] for x in facts if x["channel"] == ch),
                               orders=sum(1 for o in order_facts if o["channel"] == ch))
cohort_size = defaultdict(set)
active = defaultdict(set)
for o in order_facts:
    c = mo(first[o["cust"]]["day"])
    cohort_size[c].add(o["cust"])
    active[(c, o["since"])].add(o["cust"])
for c in sorted(cohort_size):
    if c < 202600:
        exp["cohorts"][str(c)] = {str(k): len(active[(c, k)]) / len(cohort_size[c]) for k in range(13) if active[(c, k)]}
for cur in ("EUR", "GBP", "CAD", "AUD"):
    exp["fx"][cur] = dict(avg_2025=sum(r for d, r in rates[cur] if d.year == 2025) / sum(1 for d, r in rates[cur] if d.year == 2025),
                          latest=rates[cur][-1][1], latest_date=rates[cur][-1][0].isoformat())
shop_lines = sum(len(o["lines"]) for o in orders)
exp["checks"] = dict(shop_orders=len(orders), shop_lines=shop_lines, voided=sum(o["status"] == "voided" for o in orders),
                     mp_lines=sum(len(o["lines"]) for o in mp), mp_orders=len(mp), mp_duplicates=dup_rows, mp_header_repeats=header_repeats,
                     unknown_sku_lines=3, messy_skus=messy, max_rate_age=max(o["rate_age"] for o in order_facts),
                     customers=len(first), refunds_over_total=0)
ytd = lambda y: sum(e["net"] for m, e in exp["months"].items() if int(m) // 100 == y and int(m) % 100 <= 10)
exp["summary"] = dict(net_2025_ytd=ytd(2025), net_2026_ytd=ytd(2026), net_all=sum(x["net"] for x in facts),
                      orders_all=len(order_facts), fx_effect_2026=sum(e.get("fx_effect", 0) for e in exp["months"].values()),
                      net_cc_2026=sum(e.get("net_cc", 0) for e in exp["months"].values()))
(TEST / "expected.json").write_text(json.dumps(exp, indent=1), encoding="utf-8")
s = exp["summary"]
print(f"webshop {len(orders)} orders / {shop_lines} lines, marketplace {len(mp)} orders, customers {len(first)}")
print(f"net sales Jan-Oct 2025 {s['net_2025_ytd']:,.0f}  Jan-Oct 2026 {s['net_2026_ytd']:,.0f}  "
      f"({s['net_2026_ytd'] / s['net_2025_ytd'] - 1:+.1%}); constant currency {s['net_cc_2026']:,.0f} "
      f"({s['net_cc_2026'] / s['net_2025_ytd'] - 1:+.1%}); FX effect {s['fx_effect_2026']:+,.0f}")
