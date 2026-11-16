"""Independent check of the E15 Power BI project.

1. Python works out the sales again straight from the files in data/ (no Power BI involved): the webshop export with
   its order fields on the first line only, the EU marketplace export (repeated headers, rows exported twice, decimal
   commas), USD at the ECB rate of the order date (or the last rate before it), discounts and refunds spread over the
   lines, cost of goods, new vs returning customers, cohorts, constant-currency sales and the FX effect.
2. Compares that with the answer planted when the data was generated (test/expected.json).
3. Reads the project files: the TMDL model must have the tables, relationships, measures and DAX functions, the
   Power Query must call the Frankfurter API with a fallback, every field used by a visual must exist in the model.
4. Compares Python with what the model in Power BI Desktop answered (test/model-snapshot.json, written by
   test/query_model.ps1 while the report was open): every month, currency, product, market, cohort, KPI, the key
   insight sentences, the checks and the rate of every order.
5. Larkspur-Sales.pbix is a Power BI file with the data model inside.
Exit code 1 if anything differs.   Run:  python test/verify.py
"""
import csv, json, re, sys, zipfile
from bisect import bisect_right
from collections import defaultdict
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
SM = ROOT / "Larkspur-Sales.SemanticModel" / "definition"
RP = ROOT / "Larkspur-Sales.Report" / "definition"
results = []


def check(name, ok, detail=""):
    results.append(bool(ok))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


def near(a, b, tol=0.01):
    return a is not None and b is not None and abs(float(a) - float(b)) < tol


def eu(t):
    return float(t.strip().replace(".", "").replace(",", ".")) if t.strip() else 0.0


def mo(d):
    return d.year * 100 + d.month


# ====================================================================================== 1. plain Python from the files
rates = defaultdict(list)
with open(DATA / "fx-rates-saved.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        rates[r["Currency"]].append((date.fromisoformat(r["Date"]), float(r["Units per USD"])))
rate_days = {c: [d for d, _ in sorted(v)] for c, v in rates.items()}
rate_vals = {c: [x for _, x in sorted(v)] for c, v in rates.items()}


def rate_on(cur, d):
    if cur == "USD":
        return 1.0, d
    i = bisect_right(rate_days[cur], d) - 1
    return rate_vals[cur][i], rate_days[cur][i]


def month_avg(cur, y, m):
    if cur == "USD":
        return 1.0
    v = [x for dd, x in rates[cur] if dd.year == y and dd.month == m]
    return sum(v) / len(v) if v else None


products = {}
with open(DATA / "products.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        products[r["SKU"].strip().upper()] = dict(name=r["Product"], cat=r["Category"], cost=float(r["Unit Cost USD"]))

# --- webshop: order fields on the first line only
with open(DATA / "shop-orders-export.csv", encoding="utf-8") as f:
    shop_rows = [r for r in csv.DictReader(f) if r["Name"].strip()]
shop_head, shop_lines = {}, defaultdict(list)
for i, r in enumerate(shop_rows, 1):
    if r["Financial Status"].strip():
        shop_head[r["Name"]] = r
    shop_lines[r["Name"]].append(dict(line=f"S-{i:05d}", sku=r["Lineitem sku"].strip().upper(), raw=r["Lineitem sku"],
                                      qty=int(r["Lineitem quantity"]), price=float(r["Lineitem price"])))
orders, lines = [], []
voided_orders = voided_lines = 0
for no, h in shop_head.items():
    ls = shop_lines[no]
    if h["Financial Status"] == "voided":
        voided_orders += 1
        voided_lines += len(ls)
        continue
    num = lambda k: float(h[k]) if h[k].strip() else 0.0
    gross = sum(l["qty"] * l["price"] for l in ls)
    sub, taxes, refunded = num("Subtotal"), num("Taxes"), num("Refunded Amount")
    returns = sub if h["Financial Status"] == "refunded" else (refunded * sub / (sub + taxes) if refunded else 0.0)
    o = dict(no=no, seq=int(no.lstrip("#")), channel="Webshop", day=date.fromisoformat(h["Created at"][:10]),
             cust=h["Email"].strip().lower(), cur=h["Currency"].strip().upper(), country=h["Shipping Country"].strip(),
             status=h["Financial Status"], gross=gross, disc=num("Discount Amount"), sub=sub, ship=num("Shipping"), taxes=taxes,
             total=num("Total"), refunded=refunded, returns=returns)
    orders.append(o)
    for l in ls:
        g = l["qty"] * l["price"]
        ld = o["disc"] * g / gross if gross else 0.0
        lr = returns * (g - ld) / sub if sub else 0.0
        p = products.get(l["sku"])
        cost = 0.0 if o["status"] == "refunded" or p is None else l["qty"] * p["cost"]
        lines.append(dict(o=o, sku=l["sku"], raw=l["raw"], qty=l["qty"], gross=g, disc=ld, ret=lr, cogs=cost, line=l["line"]))

# --- marketplace: headers repeat, some rows exported twice, decimal commas
with open(DATA / "marketplace-orders-eu.csv", encoding="utf-8") as f:
    mp_all = [r for r in csv.reader(f, delimiter=";") if r and r[0].strip()]
header = mp_all[0]
mp_body = mp_all[1:]
mp_headers = sum(1 for r in mp_body if r[0] == "Order ID")
seen, mp_rows, mp_dups = set(), [], 0
for r in mp_body:
    if r[0] == "Order ID":
        continue
    if tuple(r) in seen:
        mp_dups += 1
        continue
    seen.add(tuple(r))
    mp_rows.append(dict(zip(header, r)))
mp_orders = {}
for i, r in enumerate(mp_rows, 1):
    no = r["Order ID"]
    d0 = date(int(r["Order date"][6:]), int(r["Order date"][3:5]), int(r["Order date"][:2]))
    o = mp_orders.setdefault(no, dict(no=no, seq=int(no.split("-")[1]), channel="Marketplace EU", day=d0, cust=r["Buyer ID"], cur="EUR",
                                      country=r["Country"], gross=0.0, disc=0.0, ship=0.0, taxes=0.0, returns=0.0, lines=[]))
    g = eu(r["Line total"])
    back = r["Status"] == "Returned"
    o["gross"] += g
    o["returns"] += g if back else 0.0
    p = products.get(r["SKU"].strip().upper())
    o["lines"].append(dict(sku=r["SKU"].strip().upper(), raw=r["SKU"], qty=int(r["Qty"]), gross=g, disc=0.0, ret=g if back else 0.0,
                           cogs=0.0 if back or p is None else int(r["Qty"]) * p["cost"], line=f"M-{i:05d}", back=back))
for o in mp_orders.values():
    o.update(sub=o["gross"], total=o["gross"], refunded=o["returns"])
    nb = sum(l["back"] for l in o["lines"])
    o["status"] = "paid" if nb == 0 else "refunded" if nb == len(o["lines"]) else "partially_refunded"
    orders.append(o)
    for l in o.pop("lines"):
        lines.append(dict(o=o, **{k: v for k, v in l.items() if k != "back"}))

for o in orders:
    o["rate"], rd = rate_on(o["cur"], o["day"])
    o["rate_age"] = (o["day"] - rd).days
    o["m"] = mo(o["day"])
for l in lines:
    l["net_local"] = l["gross"] - l["disc"] - l["ret"]
    l["net"] = l["net_local"] / l["o"]["rate"]
first = {}
for o in sorted(orders, key=lambda o: (o["day"], o["seq"])):
    first.setdefault(o["cust"], o)
for o in orders:
    f0 = first[o["cust"]]
    o["new"] = f0 is o
    o["since"] = (o["day"].year - f0["day"].year) * 12 + o["day"].month - f0["day"].month

ads = defaultdict(float)
with open(DATA / "ads-spend.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        ads[int(r["Month"].replace("-", ""))] += float(r["Spend USD"])

months = sorted({o["m"] for o in orders})


def agg(sel_lines, sel_orders):
    gross = sum(l["gross"] / l["o"]["rate"] for l in sel_lines)
    disc = sum(l["disc"] / l["o"]["rate"] for l in sel_lines)
    ret = sum(l["ret"] / l["o"]["rate"] for l in sel_lines)
    net = sum(l["net"] for l in sel_lines)
    cogs = sum(l["cogs"] for l in sel_lines)
    n = len(sel_orders)
    return dict(gross=gross, discounts=disc, returns=ret, net=net, cogs=cogs, gross_profit=net - cogs, orders=n,
                units=sum(l["qty"] for l in sel_lines), aov=(gross - disc) / n if n else None,
                new_customers=sum(o["new"] for o in sel_orders), returning_orders=sum(not o["new"] for o in sel_orders),
                shipping=sum(o["ship"] / o["rate"] for o in sel_orders), taxes=sum(o["taxes"] / o["rate"] for o in sel_orders))


def cc(sel_lines):
    """constant currency: local net at last year's average rate of the same month, and the FX effect"""
    by = defaultdict(lambda: [0.0, 0.0])
    for l in sel_lines:
        k = (l["o"]["cur"], l["o"]["m"])
        by[k][0] += l["net_local"]
        by[k][1] += l["net"]
    tot_cc = tot_fx = 0.0
    found = False
    for (cur, m), (loc, usd) in by.items():
        if m // 100 - 1 < months[0] // 100:      # the first year has no year before it
            continue
        r = month_avg(cur, m // 100 - 1, m % 100)
        if r is None:
            continue
        found = True
        tot_cc += loc / r
        tot_fx += usd - loc / r
    return (tot_cc, tot_fx) if found else (None, None)


lines_by_m, orders_by_m = defaultdict(list), defaultdict(list)
for l in lines:
    lines_by_m[l["o"]["m"]].append(l)
for o in orders:
    orders_by_m[o["m"]].append(o)
py_months = {}
for m in months:
    e = agg(lines_by_m[m], orders_by_m[m])
    e["ad_spend"] = ads[m]
    e["net_cc"], e["fx_effect"] = cc(lines_by_m[m])
    py_months[m] = e

exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))

# ====================================================================================== 2. Python = planted answer
print("— 1. Python from the files = the planted answer")
bad = [(m, k) for m in months for k, v in exp["months"][str(m)].items()
       if not (near(py_months[m][k], v, 0.02) if isinstance(v, float) else py_months[m][k] == v)]
check(f"{len(months)} months x {len(exp['months'][str(months[0])])} figures (net sales, orders, AOV, refunds, COGS, new customers, FX...)",
      not bad, str(bad[:5]))
bad = [(c, m) for c, per in exp["currencies"].items() for m, v in per.items()
       if not (near(sum(l["net_local"] for l in lines_by_m[int(m)] if l["o"]["cur"] == c), v["local"])
               and near(sum(l["net"] for l in lines_by_m[int(m)] if l["o"]["cur"] == c), v["usd"]))]
check("net sales per currency and month, local and USD", not bad, str(bad[:5]))
bad = [s for s, v in exp["products"].items()
       if not (sum(l["qty"] for l in lines if l["sku"] == s) == v["units"] and near(sum(l["net"] for l in lines if l["sku"] == s), v["net"])
               and near(sum(l["cogs"] for l in lines if l["sku"] == s), v["cogs"]))]
check(f"units, net sales and cost of all {len(exp['products'])} SKUs", not bad, str(bad))
bad = [c for c, v in exp["markets"].items()
       if not (near(sum(l["net"] for l in lines if l["o"]["country"] == c and 202500 < l["o"]["m"] <= 202510), v["net_2025_ytd"])
               and near(sum(l["net"] for l in lines if l["o"]["country"] == c and l["o"]["m"] > 202600), v["net_2026_ytd"]))]
check(f"net sales of {len(exp['markets'])} countries, Jan-Oct 2025 and 2026", not bad, str(bad))
cohort_size, active = defaultdict(set), defaultdict(set)
for o in orders:
    c = mo(first[o["cust"]]["day"])
    cohort_size[c].add(o["cust"])
    active[(c, o["since"])].add(o["cust"])
py_cohorts = {str(c): {str(k): len(active[(c, k)]) / len(cohort_size[c]) for k in range(13) if active[(c, k)]}
              for c in sorted(cohort_size) if c < 202600}
check("2025 cohorts: share who ordered again 0-12 months later", py_cohorts.keys() == exp["cohorts"].keys() and all(
    py_cohorts[c].keys() == exp["cohorts"][c].keys() and all(near(py_cohorts[c][k], v, 1e-9) for k, v in exp["cohorts"][c].items())
    for c in py_cohorts))
x = exp["checks"]
check("files: webshop orders, lines and voided orders", len(shop_head) == x["shop_orders"] and len(shop_rows) == x["shop_lines"]
      and voided_orders == x["voided"], f"{len(shop_head)} {len(shop_rows)} {voided_orders}")
check("marketplace: repeated headers and rows exported twice found and removed",
      mp_headers == x["mp_header_repeats"] and mp_dups == x["mp_duplicates"] and len(mp_rows) == x["mp_lines"],
      f"{mp_headers} {mp_dups} {len(mp_rows)}")
check("3 swatch lines have a SKU that is not in products.csv", sum(1 for l in lines if l["sku"] not in products) == x["unknown_sku_lines"])
check(f"every order has a rate at most 5 days old (oldest {max(o['rate_age'] for o in orders)} days)",
      max(o["rate_age"] for o in orders) == x["max_rate_age"] <= 5)
s = exp["summary"]
ytd = lambda y: sum(l["net"] for l in lines if l["o"]["m"] // 100 == y and l["o"]["m"] % 100 <= 10)
check(f"summary: Jan-Oct 2026 {ytd(2026):,.0f} vs 2025 {ytd(2025):,.0f}, FX effect {s['fx_effect_2026']:+,.0f}",
      near(ytd(2026), s["net_2026_ytd"]) and near(ytd(2025), s["net_2025_ytd"]) and len(orders) == s["orders_all"])
# the files themselves add up
check("every webshop order: lines - discount = subtotal and subtotal + shipping + taxes = total",
      all(abs(o["gross"] - o["disc"] - o["sub"]) < 0.005 and abs(o["sub"] + o["ship"] + o["taxes"] - o["total"]) < 0.005
          for o in orders if o["channel"] == "Webshop"))

# ====================================================================================== 3. project files
print("— 2. The Power BI project files")
tmdl = {p.stem: p.read_text(encoding="utf-8") for p in (SM / "tables").glob("*.tmdl")}
need = ["Sales", "Orders", "Product", "Customer", "Market", "Currency", "Date", "FX API", "FX Saved", "FX Rate", "Ad Spend",
        "Source Files", "File Totals", "Checks", "Insights", "_Measures", "Months Later"]
check(f"model has the {len(need)} tables", all(t in tmdl for t in need), str(set(need) - set(tmdl)))
rel = (SM / "relationships.tmdl").read_text(encoding="utf-8")
pairs = [("Sales.'Order ID'", "Orders.'Order ID'"), ("Sales.SKU", "Product.SKU"), ("Orders.Date", "Date.Date"),
         ("Orders.'Customer ID'", "Customer.'Customer ID'"), ("Orders.Country", "Market.Country"), ("Orders.Currency", "Currency.Currency"),
         ("'FX Rate'.Currency", "Currency.Currency"), ("'Ad Spend'.Month", "Date.Date")]
check("8 relationships (two fact tables, dimensions, rates by currency)",
      all(f"fromColumn: {a}\n\ttoColumn: {b}" in rel for a, b in pairs) and rel.count("relationship ") == 8)
measures = dict(re.findall(r"\tmeasure ('[^']+'|\S+) =(.*?)(?=\n\t\t(?:formatString|displayFolder))", tmdl["_Measures"], re.S))
measures = {k.strip("'"): v for k, v in measures.items()}
check(f"{len(measures)} DAX measures incl. time intelligence, constant currency, cohorts, checks and insights",
      len(measures) >= 90 and all(m in measures for m in ["Net Sales", "Net Sales LY", "Net Sales CC", "FX Effect", "Cohort Retention %",
                                                          "Insight", "Check Status", "MER", "CAC", "Net Sales YTD"]))
alltext = "\n".join(tmdl.values())
check("DAX uses SAMEPERIODLASTYEAR, TOTALYTD, SUMMARIZE, TOPN, LOOKUPVALUE, a calculated table and calculated columns",
      all(f in alltext for f in ["SAMEPERIODLASTYEAR", "TOTALYTD", "SUMMARIZE (", "TOPN", "LOOKUPVALUE", "= calculated", "column Rate ="]))
ex = (SM / "expressions.tmdl").read_text(encoding="utf-8")
fx = tmdl["FX API"]
check("Power Query calls the Frankfurter API with a static base URL, RelativePath and Query, and falls back when it fails",
      'Web.Contents("https://api.frankfurter.dev"' in fx and "RelativePath" in fx and "Query = [base" in fx and "try " in fx and "otherwise" in fx)
check("Power Query reads both exports: order fields from the first line, repeated headers and duplicates removed, decimal commas",
      'Text.Trim([Financial Status]) <> ""' in ex and '[Order ID] <> "Order ID"' in ex and "Table.Distinct" in ex and 'Text.Replace(Text.Remove' in ex)
check("Date table marked as a date table", "dataCategory: Time" in tmdl["Date"])
fields = defaultdict(set)
for t, txt in tmdl.items():
    for c in re.findall(r"\n\tcolumn ('[^']+'|[^\s=]+)", txt):
        fields[t].add(c.strip("'"))
fields["_Measures"] |= set(measures)
pages = json.loads((RP / "pages" / "pages.json").read_text(encoding="utf-8"))["pageOrder"]
used, missing, nvis = set(), [], 0
for p in pages:
    for vf in (RP / "pages" / p / "visuals").glob("*/visual.json"):
        nvis += 1
        txt = vf.read_text(encoding="utf-8")
        for ent, prop in re.findall(r'"Entity": "([^"]+)"\s*}\s*}\s*,\s*"Property": "([^"]+)"', txt):
            used.add((ent, prop))
            if prop not in fields.get(ent, ()):
                missing.append(f"{p}:{ent}.{prop}")
check(f"{len(pages)} pages, {nvis} visuals: every one of the {len(used)} fields they use exists in the model",
      pages == ["overview", "products", "customers", "fx", "checks"] and not missing and len(used) > 40, str(missing[:5]))

# ====================================================================================== 4. the model in Power BI Desktop
print("— 3. What the model in Power BI Desktop answered (test/model-snapshot.json)")
snap_file = ROOT / "test" / "model-snapshot.json"
if not snap_file.exists():
    check("model snapshot present (run test/query_model.ps1 with the report open)", False, "no snapshot yet")
else:
    snap = json.loads(snap_file.read_text(encoding="utf-8-sig"))
    rows = lambda k: snap[k] if isinstance(snap[k], list) else [snap[k]]
    keymap = dict(Net="net", Gross="gross", Discounts="discounts", Returns="returns", Orders="orders", Units="units", COGS="cogs",
                  New="new_customers", Returning="returning_orders", Shipping="shipping", Taxes="taxes", Ads="ad_spend",
                  CC="net_cc", FX="fx_effect", AOV="aov")
    bad = []
    for r in rows("months"):
        e = py_months[int(r["Month No"])]
        for k, pk in keymap.items():
            if not ((r.get(k) is None and e[pk] is None) or near(r.get(k), e[pk], 0.02)):
                bad.append((r["Month No"], k, r.get(k), e[pk]))
    check(f"{len(rows('months'))} months x {len(keymap)} measures = Python", len(rows("months")) == len(months) and not bad, str(bad[:4]))
    bad = [r for r in rows("currencies") if not (near(r["Local"], sum(l["net_local"] for l in lines_by_m[int(r["Month No"])] if l["o"]["cur"] == r["Currency"]))
                                                 and near(r["Net"], sum(l["net"] for l in lines_by_m[int(r["Month No"])] if l["o"]["cur"] == r["Currency"])))]
    check(f"{len(rows('currencies'))} currency-months, local and USD", not bad and len(rows("currencies")) > 90, str(bad[:3]))
    bad = [r for r in rows("products") if not (r["Units"] == sum(l["qty"] for l in lines if l["sku"] == r["SKU"])
                                               and near(r["Net"], sum(l["net"] for l in lines if l["sku"] == r["SKU"]))
                                               and near(r["COGS"], sum(l["cogs"] for l in lines if l["sku"] == r["SKU"])))]
    check(f"{len(rows('products'))} SKUs: units, net sales, cost", not bad and len(rows("products")) == len(exp["products"]), str(bad[:3]))
    bad = [r for r in rows("markets") if not (near(r["Net"], exp["markets"][r["Country"]]["net_2026_ytd"])
                                              and near(r["LY"], exp["markets"][r["Country"]]["net_2025_ytd"]))]
    check(f"{len(rows('markets'))} countries in 2026 and the same months of 2025", not bad, str(bad[:3]))
    got = defaultdict(dict)
    for r in rows("cohorts"):
        if r.get("R") is not None:
            got[str(r["Cohort No"])][str(r["Months Later"])] = r["R"]
    check("cohort grid = Python (every 2025 cohort, 0-12 months)", got.keys() == py_cohorts.keys() and all(
        got[c].keys() == py_cohorts[c].keys() and all(near(got[c][k], v, 1e-9) for k, v in py_cohorts[c].items()) for c in py_cohorts))
    # KPIs and key insights for the default page (2026, every channel and market)
    l26 = [l for l in lines if l["o"]["m"] // 100 == 2026]
    o26 = [o for o in orders if o["m"] // 100 == 2026]
    l25 = [l for l in lines if 202500 < l["o"]["m"] <= 202510]
    o25 = [o for o in orders if 202500 < o["m"] <= 202510]
    a26, a25 = agg(l26, o26), agg(l25, o25)
    cc26, fx26 = cc(l26)
    rr = lambda a: a["returns"] / (a["gross"] - a["discounts"])
    shop_new26 = sum(o["new"] for o in o26 if o["channel"] == "Webshop")
    ad26 = sum(v for m, v in ads.items() if m // 100 == 2026)
    k = rows("kpi")[0]
    want = dict(Net=a26["net"], LY=a25["net"], YoY=a26["net"] / a25["net"] - 1, Orders=a26["orders"], AOV=a26["aov"], AOVLY=a25["aov"],
                GM=a26["gross_profit"] / a26["net"], GMLY=a25["gross_profit"] / a25["net"], Ret=a26["returning_orders"] / a26["orders"],
                Refund=rr(a26), CC=cc26, CCG=cc26 / a25["net"] - 1, FX=fx26,
                NonUSD=sum(l["net"] for l in l26 if l["o"]["cur"] != "USD") / a26["net"],
                MER=sum(l["net"] for l in l26 if l["o"]["channel"] == "Webshop") / ad26, CAC=ad26 / shop_new26)
    bad = [(n, k.get(n), v) for n, v in want.items() if not near(k.get(n), v, 1e-6 if abs(v) < 10 else 0.02)]
    check(f"{len(want)} KPIs for 2026 vs 2025 (growth, AOV, margin, returning, refunds, constant currency, FX effect, MER, CAC)", not bad, str(bad))
    mk = defaultdict(lambda: [0.0, 0.0])
    market_of = {"United States": "United States", "Canada": "Canada", "United Kingdom": "United Kingdom", "Australia": "Australia"}
    for l in l26:
        mk[market_of.get(l["o"]["country"], "Europe (EU)")][0] += l["net"]
    for l in l25:
        mk[market_of.get(l["o"]["country"], "Europe (EU)")][1] += l["net"]
    fast = max(((m, v[0] / v[1] - 1, v[1], v[0]) for m, v in mk.items() if v[1] >= 1000), key=lambda t: t[1])
    prod = defaultdict(float)
    for l in l26:
        prod[products[l["sku"]]["name"] if l["sku"] in products else ""] += l["net"]
    best = max(prod.items(), key=lambda t: t[1])
    cat = defaultdict(lambda: [0.0, 0.0])
    for l in l26:
        if l["sku"] in products:
            c = cat[products[l["sku"]]["cat"]]
            c[0] += l["ret"] / l["o"]["rate"]
            c[1] += (l["gross"] - l["disc"]) / l["o"]["rate"]
    worst = max(((c, v[0] / v[1]) for c, v in cat.items()), key=lambda t: t[1])
    pct = lambda v: f"{v:.1%}"
    sentences = {
        1: f"Net sales {a26['net']:,.0f} USD, {want['YoY']:+.1%} on the same months last year",
        2: f"At last year's exchange rates {want['CCG']:+.1%}: currency moves {'added' if fx26 >= 0 else 'took off'} {abs(fx26):,.0f} USD",
        3: f"{fast[0]} {fast[1]:+.1%} ({fast[2]:,.0f} to {fast[3]:,.0f})",
        4: f"{best[0]}: {best[1]:,.0f} ({pct(best[1] / a26['net'])} of net sales)",
        5: f"Returning customers placed {pct(want['Ret'])} of orders; ads cost {want['CAC']:,.2f} per new webshop customer",
        6: f"Refunds {pct(want['Refund'])} of sales; highest in {worst[0]} ({pct(worst[1])})"}
    got = {int(r["No"]): r["Text"] for r in rows("insights")}
    bad = [(n, got.get(n), t) for n, t in sentences.items() if got.get(n) != t]
    check("the 6 key insight sentences for 2026, word for word", not bad, str(bad))
    ck = {int(r["No"]): (r["Status"], r["Detail"]) for r in rows("checks")}
    statuses = [ck.get(n, ("?",))[0] for n in range(1, 11)]
    check("Checks page: 9 of 9 PASS and 1 REVIEW (the 3 swatch lines)", statuses == ["PASS"] * 9 + ["REVIEW"], str(ck))
    d1 = f"{len(shop_rows):,} lines in the file = {sum(1 for l in lines if l['o']['channel'] == 'Webshop'):,} sold + {voided_lines:,} on {voided_orders:,} voided orders"
    d2 = f"{len(mp_body) - 0:,} rows: {mp_headers} repeated headers and {mp_dups} duplicates removed, {len(mp_rows):,} lines"
    check("check details: lines in each file, voided orders, repeated headers, duplicates", ck.get(1, ("", ""))[1] == d1 and ck.get(2, ("", ""))[1] == d2,
          f"{ck.get(1)} / {d1} | {ck.get(2)} / {d2}")
    fxr = {r["Currency"]: r["Latest"] for r in rows("fx") if r.get("Currency")}
    saved_latest = {c: rate_vals[c][-1] for c in rate_vals}
    api = rows("fxinfo")[0]
    check(f"rates came live from the API ({api['ApiRates']:,} rates up to {api['Latest']}) and the latest = saved copy where both exist",
          api["ApiRates"] >= len(rates["EUR"]) * 4 and all(c in fxr for c in saved_latest), str(api))
    od = {r["Order"]: r for r in rows("orders")}
    bad = [o["no"] for o in orders if o["no"] not in od or not near(od[o["no"]]["Rate"], o["rate"], 1e-9)
           or od[o["no"]]["Type"] != ("New" if o["new"] else "Returning") or od[o["no"]]["Since"] != o["since"]]
    check(f"all {len(orders):,} orders: exchange rate, new/returning, months since first order", not bad and len(od) == len(orders), str(bad[:5]))

# ====================================================================================== 5. the .pbix
pbix = ROOT / "Larkspur-Sales.pbix"
if pbix.exists():
    with zipfile.ZipFile(pbix) as z:
        names = z.namelist()
    check("Larkspur-Sales.pbix is a Power BI file with the data model inside", "DataModel" in names and any(n.startswith("Report/") for n in names))
else:
    check("Larkspur-Sales.pbix present", False, "not saved yet")

print(f"\n{sum(results)}/{len(results)} checks pass")
sys.exit(0 if all(results) else 1)
