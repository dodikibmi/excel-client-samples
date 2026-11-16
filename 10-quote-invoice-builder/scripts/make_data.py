"""Fictional data for E10 (quote & invoice builder): customers, products, settings and the history of documents
(Aug-Nov 2026) that build_workbook.py replays through the workbook's own macros, in date order.

    data/customers.csv, data/products.csv, data/settings.json, data/history.json

History events: create (draft quote), issue, accept / decline, convert (accepted quote -> draft invoice), paid.
Two drafts are left blocked on purpose: a customer without an email address and a price under the minimum.
"""
import csv, json, random
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
D = ROOT / "data"
D.mkdir(exist_ok=True)
rnd = random.Random(1010)
AS_OF = date(2026, 11, 16)

CUSTOMERS = [  # id, name, contact, email, city, terms (days), tax exempt
    ("C001", "Alder & Finch Dental", "Maya Alder", "office@alderfinch.example", "Port Ellis", 30, "No"),
    ("C002", "Brookside Fitness", "Owen Hart", "owen@brooksidefit.example", "Port Ellis", 14, "No"),
    ("C003", "Calder Estate Agents", "Nadia Calder", "accounts@calderhomes.example", "Westmere", 30, "No"),
    ("C004", "Driftwood Cafe", "Sam Ortega", "hello@driftwoodcafe.example", "Port Ellis", 14, "No"),
    ("C005", "Evergreen Landscaping", "Tom Reyes", "tom@evergreenland.example", "Ashby", 30, "No"),
    ("C006", "Fairview Veterinary Clinic", "Dr. Lena Brandt", "admin@fairviewvet.example", "Westmere", 45, "No"),
    ("C007", "Granite Peak Builders", "Marcus Webb", "ap@granitepeak.example", "Ashby", 45, "No"),
    ("C008", "Harbour Lights Festival", "Priya Shah", "events@harbourlights.example", "Port Ellis", 30, "No"),
    ("C009", "Copperpot Bakery", "Ana Lima", "", "Port Ellis", 14, "No"),   # planted: no email
    ("C010", "Ironbridge Motors", "Dev Patel", "service@ironbridgemotors.example", "Ashby", 30, "No"),
    ("C011", "Hillcrest Primary School", "Ruth Okafor", "office@hillcrest.example", "Westmere", 30, "Yes"),
    ("C012", "Juniper Yoga Studio", "Elena Varga", "elena@juniperyoga.example", "Port Ellis", 14, "No"),
]

PRODUCTS = [  # code, category, description, unit, list price, min price, max line discount %
    ("PRT-01", "Printing", "Business cards, 500, matt 400gsm", "pack", 45.00, 38.00, 10),
    ("PRT-02", "Printing", "Flyers A5, 1,000, gloss 150gsm", "pack", 120.00, 95.00, 10),
    ("PRT-03", "Printing", "Tri-fold brochures, 500", "pack", 260.00, 210.00, 10),
    ("PRT-04", "Printing", "Letterheads A4, 500", "pack", 85.00, 70.00, 10),
    ("PRT-05", "Printing", "Booklets A4 16 pages, 100", "pack", 340.00, 280.00, 10),
    ("PRT-06", "Printing", "Envelopes DL printed, 500", "pack", 75.00, 60.00, 10),
    ("SGN-01", "Signage", "Aluminium sign 600 x 400 mm", "each", 95.00, 80.00, 15),
    ("SGN-02", "Signage", "Acrylic office sign with stand-offs", "each", 140.00, 115.00, 15),
    ("SGN-03", "Signage", "Window vinyl lettering", "m2", 55.00, 45.00, 15),
    ("SGN-04", "Signage", "Pull-up banner 850 x 2000 mm", "each", 165.00, 135.00, 15),
    ("SGN-05", "Signage", "A-frame pavement sign", "each", 210.00, 175.00, 10),
    ("SGN-06", "Signage", "Vehicle door decals", "pair", 180.00, 150.00, 10),
    ("LFP-01", "Large format", "PVC banner with eyelets", "m2", 28.00, 22.00, 20),
    ("LFP-02", "Large format", "Poster A1, satin", "each", 18.00, 14.00, 20),
    ("LFP-03", "Large format", "Mesh banner", "m2", 34.00, 27.00, 20),
    ("LFP-04", "Large format", "Foam board A1", "each", 32.00, 26.00, 20),
    ("LFP-05", "Large format", "Wall graphic, laminated", "m2", 65.00, 52.00, 15),
    ("LFP-06", "Large format", "Floor sticker, anti-slip", "each", 24.00, 19.00, 20),
    ("SRV-01", "Design & services", "Design time", "hour", 60.00, 50.00, 0),
    ("SRV-02", "Design & services", "Logo refresh", "fixed", 450.00, 380.00, 10),
    ("SRV-03", "Design & services", "Proof and colour check", "each", 25.00, 25.00, 0),
    ("SRV-04", "Design & services", "Installation", "hour", 55.00, 45.00, 0),
    ("SRV-05", "Design & services", "Local delivery", "trip", 35.00, 35.00, 0),
    ("SRV-06", "Design & services", "Rush fee (48 h)", "job", 80.00, 80.00, 0),
]
P = {p[0]: p for p in PRODUCTS}
QTY = {"pack": [1, 1, 2, 3], "each": [1, 2, 4, 6, 10, 20], "m2": [3, 4.5, 6, 8, 12], "pair": [1, 2], "hour": [1, 2, 3, 4, 6],
       "fixed": [1], "trip": [1, 2], "job": [1]}

SETTINGS = {
    "company": "Harbourline Print & Signs", "address": "14 Quay Street, Port Ellis", "phone": "+00 555 0140",
    "email": "sales@harbourline.example", "tax_label": "Sales tax", "tax_rate": 0.10, "quote_valid_days": 30,
    "max_doc_discount": 0.15, "quote_prefix": "Q-", "invoice_prefix": "INV-", "pdf_folder": "PDF",
    "email_mode": ".eml file", "outbox_folder": "Outbox", "user": "Sales desk", "currency": "$",
    "as_of": AS_OF.isoformat(),
    "bank": "Harbour Mutual Bank · account 00-1234-5678901-00",
    "terms_note": "Payment by bank transfer. Please quote the invoice number.",
}


def bday(d):
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d


def lines_for(cust):
    picks = rnd.sample(PRODUCTS, rnd.choice([1, 2, 2, 3, 3, 4, 5]))
    out = []
    for code, _, _, unit, lst, mn, mx in picks:
        price = lst
        if rnd.random() < 0.15 and mn < lst:          # negotiated, never under the minimum
            price = round(rnd.uniform(mn, lst) / 0.5) * 0.5
        disc = 0 if mx == 0 or rnd.random() < 0.8 else rnd.choice([d for d in (5, 10, 15, 20) if d <= mx])
        out.append([code, rnd.choice(QTY[unit]), price, disc / 100])
    return out


events, docs = [], []
start = date(2026, 8, 3)
days = sorted(bday(start + timedelta(days=rnd.randint(0, (date(2026, 11, 13) - start).days))) for _ in range(46))
for i, d in enumerate(days):
    cust = rnd.choice([c for c in CUSTOMERS if c[0] != "C009"])
    q = {"ref": f"q{i + 1:02d}", "customer": cust[0], "date": d.isoformat(), "doc_discount": rnd.choice([0, 0, 0, 0, 0.05, 0.10]),
         "notes": "", "lines": lines_for(cust)}
    docs.append(q)
    seq = 0
    events.append({"date": d.isoformat(), "action": "create", "ref": q["ref"]})
    recent = (AS_OF - d).days < 6
    if recent and rnd.random() < 0.5:
        continue                                        # still a draft
    events.append({"date": d.isoformat(), "action": "issue", "ref": q["ref"]})
    age = (AS_OF - d).days
    r = rnd.random()
    if age < 4 or r >= 0.82:
        continue                                        # waiting for an answer (expired once past 30 days)
    if r < 0.62:
        dd = bday(d + timedelta(days=rnd.randint(2, 9)))
        if dd >= AS_OF:
            continue
        events.append({"date": dd.isoformat(), "action": "accept", "ref": q["ref"]})
        inv = dd + timedelta(days=rnd.choice([0, 0, 1, 2]))
        inv = bday(inv)
        if inv >= AS_OF:
            continue
        events.append({"date": inv.isoformat(), "action": "convert", "ref": q["ref"]})
        events.append({"date": inv.isoformat(), "action": "issue_invoice", "ref": q["ref"]})
        due = inv + timedelta(days=cust[5])
        if due < AS_OF:
            if rnd.random() < 0.8:
                paid = min(AS_OF - timedelta(days=1), due + timedelta(days=rnd.randint(-6, 9)))
                events.append({"date": max(paid, inv).isoformat(), "action": "paid", "ref": q["ref"]})
        elif rnd.random() < 0.4:
            paid = inv + timedelta(days=rnd.randint(3, 10))
            if paid < AS_OF:
                events.append({"date": paid.isoformat(), "action": "paid", "ref": q["ref"]})
    else:
        dd = bday(d + timedelta(days=rnd.randint(3, 12)))
        if dd < AS_OF:
            events.append({"date": dd.isoformat(), "action": "decline", "ref": q["ref"]})

# two drafts blocked on purpose (shown on Home as "needs attention")
docs.append({"ref": "d1", "customer": "C009", "date": "2026-11-13", "doc_discount": 0, "notes": "Shop front refresh",
             "lines": [["SGN-03", 6, 55.0, 0], ["LFP-02", 4, 18.0, 0]]})
docs.append({"ref": "d2", "customer": "C002", "date": "2026-11-13", "doc_discount": 0.05, "notes": "Spring promo",
             "lines": [["SGN-04", 3, 120.0, 0], ["PRT-02", 1, 120.0, 0]]})       # 120 < minimum 135
events += [{"date": "2026-11-13", "action": "create", "ref": "d1"}, {"date": "2026-11-13", "action": "create", "ref": "d2"}]

ORDER = {"create": 0, "issue": 1, "accept": 2, "decline": 2, "convert": 3, "issue_invoice": 4, "paid": 5}
events.sort(key=lambda e: (e["date"], ORDER[e["action"]], e["ref"]))

with open(D / "customers.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f); w.writerow(["Customer ID", "Name", "Contact", "Email", "City", "Terms (days)", "Tax exempt"]); w.writerows(CUSTOMERS)
with open(D / "products.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f); w.writerow(["Code", "Category", "Description", "Unit", "List price", "Min price", "Max discount %"])
    w.writerows(PRODUCTS)
(D / "settings.json").write_text(json.dumps(SETTINGS, indent=1), encoding="utf-8")
(D / "history.json").write_text(json.dumps({"documents": docs, "events": events}, indent=1), encoding="utf-8")
cnt = {}
for e in events:
    cnt[e["action"]] = cnt.get(e["action"], 0) + 1
print(len(docs), "documents", cnt)
