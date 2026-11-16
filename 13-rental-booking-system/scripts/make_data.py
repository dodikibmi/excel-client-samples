"""Fictional sample data for E13 (event decor rentals): catalog, customers, users, settings and a 10-week history of
bookings that the build script replays through the workbook's own macros.

The history is checked here against the stock (no item ever over-booked), except a few planted attempts that the
tool must refuse: a double booking, a staff member trying a manager action, a change inside a locked period, a wrong PIN.

    python scripts/make_data.py        -> data/catalog.csv, customers.csv, users.csv, settings.json, history.json
"""
import csv, json, random
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
D = ROOT / "data"; D.mkdir(exist_ok=True)
rnd = random.Random(1327)

# code, category, name, owned, day rate, late fee per day, replacement cost
CATALOG = [
    ("CH-GLD", "Seating", "Chiavari chair, gold", 120, 4.50, 2.00, 65),
    ("CH-CLR", "Seating", "Ghost chair, clear", 80, 5.00, 2.00, 90),
    ("SF-VLV", "Seating", "Velvet sofa, emerald", 3, 85.00, 40.00, 1400),
    ("TB-RND", "Tables", "Round table 60 in", 20, 12.00, 6.00, 180),
    ("TB-FRM", "Tables", "Farm table 8 ft", 10, 28.00, 12.00, 420),
    ("TB-CKT", "Tables", "Cocktail table", 16, 10.00, 5.00, 150),
    ("LN-WHT", "Linens", "Tablecloth 120 in, white", 60, 9.00, 4.00, 45),
    ("LN-SGE", "Linens", "Table runner, sage", 60, 3.50, 1.50, 18),
    ("LT-FRY", "Lighting", "Fairy light curtain", 12, 22.00, 10.00, 140),
    ("LT-UPL", "Lighting", "LED uplight", 24, 15.00, 7.00, 160),
    ("LT-CND", "Lighting", "Candle lantern set (6)", 30, 8.00, 3.00, 60),
    ("AR-WOD", "Arches", "Wooden hexagon arch", 2, 95.00, 45.00, 900),
    ("AR-FLR", "Arches", "Floral arch, blush", 3, 120.00, 55.00, 1500),
    ("BD-SEQ", "Backdrops", "Sequin backdrop, champagne", 4, 65.00, 30.00, 520),
    ("BD-GRN", "Backdrops", "Greenery wall 8 x 8 ft", 3, 110.00, 50.00, 1100),
    ("CP-GLD", "Tableware", "Charger plate, gold", 200, 1.20, 0.50, 9),
    ("CP-GLS", "Tableware", "Wine glass", 240, 0.90, 0.40, 6),
    ("CT-VAS", "Centerpieces", "Glass vase centerpiece", 40, 6.00, 2.50, 35),
    ("SG-EAS", "Signage", "Easel with welcome sign", 6, 18.00, 8.00, 120),
    ("DF-PNL", "Dance floor", "Dance floor panel, white", 36, 14.00, 6.00, 260),
]
CUSTOMERS = [
    ("C001", "Bloom & Bell Weddings", "Ana Ruiz", "ana@bloomandbell.example", "Planner"),
    ("C002", "Northfield Golf Club", "Tom Hale", "events@northfieldgolf.example", "Venue"),
    ("C003", "Sarah & Daniel Moss", "Sarah Moss", "sarah.moss@mail.example", "Private"),
    ("C004", "Brightline Tech", "Priya Shah", "office@brightline.example", "Corporate"),
    ("C005", "Oak Hall Events", "Lena Ward", "lena@oakhall.example", "Venue"),
    ("C006", "Petal Lane Planning", "Mia Ford", "mia@petallane.example", "Planner"),
    ("C007", "Harbor City Rotary", "Ken Cole", "secretary@harborrotary.example", "Non-profit"),
    ("C008", "Jade & Omar Karim", "Jade Karim", "jade.karim@mail.example", "Private"),
    ("C009", "Summit Realty Group", "Ben Ortiz", "marketing@summitrealty.example", "Corporate"),
    ("C010", "Willow Creek Barn", "Rosa Diaz", "bookings@willowcreek.example", "Venue"),
    ("C011", "Grace Lee", "Grace Lee", "grace.lee@mail.example", "Private"),
    ("C012", "Lumen Creative Agency", "Owen Price", "hello@lumencreative.example", "Corporate"),
    ("C013", "St. Mark's School PTA", "Ella Brooks", "pta@stmarks.example", "Non-profit"),
    ("C014", "Velvet Rope Events", "Nina Patel", "nina@velvetrope.example", "Planner"),
]
# user, display name, role, PIN (the workbook keeps only a hash of it)
USERS = [("maya", "Maya Lin", "Staff", "2468"), ("jon", "Jon Reyes", "Staff", "1357"), ("rita", "Rita Okafor", "Manager", "9024")]
SETTINGS = {"company": "Lumen & Linen Event Rentals", "address": "41 Mill Street, Harbor City", "email": "hello@lumenlinen.example",
            "currency": "$", "tax_rate": 0.08, "deposit_pct": 0.25, "grace_days": 1, "max_days": 14,
            "pdf_folder": "Invoices", "as_of": "2027-03-19"}

ITEM = {c[0]: c for c in CATALOG}
START, END = date(2027, 1, 4), date(2027, 3, 19)
booked = {}                                   # (item, day) -> qty
EVENT_KITS = [  # typical orders: (item, (min, max))
    [("CH-GLD", (40, 110)), ("TB-RND", (5, 14)), ("LN-WHT", (5, 14)), ("CT-VAS", (5, 14)), ("CP-GLD", (40, 110))],
    [("CH-CLR", (20, 60)), ("TB-CKT", (4, 10)), ("LT-UPL", (6, 16)), ("BD-SEQ", (1, 2)), ("SG-EAS", (1, 2))],
    [("TB-FRM", (3, 8)), ("LN-SGE", (6, 16)), ("LT-FRY", (2, 6)), ("AR-WOD", (1, 1)), ("CP-GLS", (40, 120))],
    [("AR-FLR", (1, 1)), ("BD-GRN", (1, 1)), ("SF-VLV", (1, 2)), ("LT-CND", (5, 15)), ("CH-GLD", (20, 60))],
    [("DF-PNL", (16, 36)), ("LT-UPL", (8, 20)), ("TB-CKT", (6, 12)), ("SG-EAS", (1, 1))],
]


def days(a, b):
    return [a + timedelta(i) for i in range((b - a).days + 1)]


def fits(lines, out, due):
    return all(booked.get((it, d), 0) + q <= ITEM[it][3] for it, q in lines for d in days(out, due))


def take(lines, out, due):
    for it, q in lines:
        for d in days(out, due):
            booked[(it, d)] = booked.get((it, d), 0) + q


bookings, n = [], 0
day = START
while day <= date(2027, 4, 10):
    for _ in range(rnd.choice((0, 1, 1, 2))):
        event = day + timedelta(rnd.choice((2, 3, 4, 5)))
        out = event - timedelta(1)
        due = event + timedelta(rnd.choice((1, 1, 2)))
        kit = rnd.choice(EVENT_KITS)
        lines = [(it, rnd.randint(*rng)) for it, rng in kit if rnd.random() < 0.85]
        if not lines or not fits(lines, out, due):
            continue
        created = day - timedelta(rnd.choice((7, 10, 14, 21))) if day > START + timedelta(21) else day - timedelta(2)
        n += 1
        bookings.append(dict(ref=f"b{n}", customer=rnd.choice(CUSTOMERS)[0], created=created, event=event, out=out, due=due,
                             lines=lines, by=rnd.choice(("maya", "jon", "jon", "maya", "rita"))))
        take(lines, out, due)
    day += timedelta(1)

AS_OF = date.fromisoformat(SETTINGS["as_of"])
BIG = {c[0] for c in CATALOG if c[3] >= 30}          # damage / loss only on items with plenty in stock
for b in bookings:
    b["delivery"] = rnd.choice((0, 0, 75, 120, 180, 250))
live = [b for b in bookings if b["created"] <= AS_OF]
cancel = set(b["ref"] for b in rnd.sample(live[5:30], 2))
enquiry = set(b["ref"] for b in rnd.sample(live[3:45], 4))
enquiry_open = set(b["ref"] for b in [b for b in live if b["out"] > AS_OF + timedelta(3)][-2:])   # not confirmed yet

# one booking that is still out after its back-by date at the as-of date (its items stay held until it comes back):
# small items with room to spare, so the late return does not clash with later bookings
ov_lines, ov_out, ov_due = [("LT-CND", 6), ("SG-EAS", 1)], AS_OF - timedelta(6), AS_OF - timedelta(4)
assert fits(ov_lines, ov_out, AS_OF + timedelta(3))
take(ov_lines, ov_out, AS_OF + timedelta(3))
n += 1
overdue = dict(ref=f"b{n}", customer="C011", created=ov_out - timedelta(12), event=ov_out + timedelta(1), out=ov_out, due=ov_due,
               lines=ov_lines, by="maya", delivery=0)
bookings.append(overdue)
live.append(overdue)
events = []
for b in live:
    c = b["created"]
    if b["ref"] in enquiry_open:
        events.append(dict(date=c, user=b["by"], action="create", ref=b["ref"], confirm=False))
        continue
    if b["ref"] in enquiry:
        events.append(dict(date=c, user=b["by"], action="create", ref=b["ref"], confirm=False))
        conf = c + timedelta(rnd.choice((2, 3, 4)))
        if conf >= b["out"] or conf > AS_OF:
            continue                                   # still an enquiry: holds no stock
        events.append(dict(date=conf, user=rnd.choice(("maya", "jon")), action="confirm", ref=b["ref"]))
        c = conf
    else:
        events.append(dict(date=c, user=b["by"], action="create", ref=b["ref"], confirm=True))
    if b["ref"] in cancel:
        events.append(dict(date=c + timedelta(2), user="rita", action="cancel", ref=b["ref"], reason="Event postponed by the customer"))
        continue
    dep = c + timedelta(rnd.choice((0, 1, 2)))
    if dep <= AS_OF and dep < b["out"]:
        events.append(dict(date=dep, user=rnd.choice(("maya", "jon")), action="deposit", ref=b["ref"]))
    if b["out"] > AS_OF:
        continue
    events.append(dict(date=b["out"], user=rnd.choice(("maya", "jon")), action="checkout", ref=b["ref"]))
    if b is overdue:
        continue
    late = rnd.choice((0, 0, 0, 0, 1, 2, 3))
    back = b["due"] + timedelta(late)
    if back > AS_OF:
        continue
    dmg = []
    big = [(it, q) for it, q in b["lines"] if it in BIG]
    if big and rnd.random() < 0.25:
        it, q = rnd.choice(big)
        dmg = [(it, min(q, rnd.choice((1, 1, 2, 3))))]
    events.append(dict(date=back, user=rnd.choice(("maya", "jon")), action="checkin", ref=b["ref"], damaged=dmg))
    events.append(dict(date=back, user=rnd.choice(("maya", "jon", "rita")), action="invoice", ref=b["ref"]))
    pay = back + timedelta(rnd.choice((0, 2, 5, 9, 14, 30)))
    if pay <= AS_OF:
        events.append(dict(date=pay, user=rnd.choice(("maya", "jon")), action="pay", ref=b["ref"]))
out_now = [overdue]
# planted attempts the tool must refuse
open_b = [b for b in live if b["out"] > AS_OF and b["ref"] not in cancel | enquiry]
paid_feb = next(e["ref"] for e in events if e["action"] == "invoice" and "2027-02-10" <= e["date"].isoformat() <= "2027-02-25")
# stock really held: confirmed (or confirmed later) and not cancelled; written off from the return date on
held_refs = {e["ref"] for e in events if (e["action"] == "create" and e["confirm"]) or e["action"] == "confirm"} - cancel
held, lost = {}, {}
BY = {b["ref"]: b for b in bookings}
for ref in held_refs:
    b = BY[ref]
    for it, q in b["lines"]:
        for d in days(b["out"], b["due"]):
            held[(it, d)] = held.get((it, d), 0) + q
for e in events:
    if e["action"] == "checkin":
        for it, q in e["damaged"]:
            lost[it] = lost.get(it, []) + [(e["date"], q)]
free = lambda it, d: ITEM[it][3] - sum(q for dd, q in lost.get(it, []) if dd <= d) - held.get((it, d), 0)
bad = [(it, d) for it in ITEM for d in days(START, date(2027, 4, 20)) if free(it, d) < 0]
assert not bad, bad[:5]
target = open_b[0]
fmin = lambda it: min(free(it, d) for d in days(target["out"], target["due"]))
it0 = max((it for it, _ in target["lines"]), key=fmin)          # an item with some stock left, asked for a bit more
over = [(it0, fmin(it0) + 6)]
events += [
    dict(date=date(2027, 3, 3), user="maya", action="lock", through="2027-02-28", expect="BLOCKED"),
    dict(date=date(2027, 3, 3), user="rita", action="lock", through="2027-02-28"),
    dict(date=date(2027, 3, 4), user="jon", action="cancel", ref=open_b[1]["ref"], reason="Customer asked", expect="BLOCKED"),
    dict(date=date(2027, 3, 8), user="jon", action="pay_backdated", ref=paid_feb, on="2027-02-20", amount=100, expect="BLOCKED"),
    dict(date=AS_OF, user="maya", action="create_over", customer="C006", event=target["event"].isoformat(),
         out=target["out"].isoformat(), due=target["due"].isoformat(), lines=over, expect="BLOCKED"),
    dict(date=AS_OF, user="jon", action="login_bad", expect="BLOCKED"),
]
ORDER = ["login_bad", "lock", "create", "create_over", "confirm", "deposit", "cancel", "pay_backdated", "checkout", "checkin", "invoice", "pay"]
events.sort(key=lambda e: (e["date"], ORDER.index(e["action"])))
ser = lambda o: o.isoformat() if isinstance(o, date) else o
for b in bookings:
    for k in ("created", "event", "out", "due"):
        b[k] = b[k].isoformat()
hist = {"bookings": bookings, "events": [{k: ser(v) for k, v in e.items()} for e in events]}
(D / "history.json").write_text(json.dumps(hist, indent=1))
with open(D / "catalog.csv", "w", newline="", encoding="utf-8") as f:
    csv.writer(f).writerows([("Code", "Category", "Item", "Owned", "Day rate", "Late fee / day", "Replacement cost")] + CATALOG)
with open(D / "customers.csv", "w", newline="", encoding="utf-8") as f:
    csv.writer(f).writerows([("ID", "Customer", "Contact", "Email", "Type")] + CUSTOMERS)
with open(D / "users.csv", "w", newline="", encoding="utf-8") as f:
    csv.writer(f).writerows([("User", "Name", "Role", "PIN")] + USERS)
(D / "settings.json").write_text(json.dumps(SETTINGS, indent=1))
acts = {}
for e in events: acts[e["action"]] = acts.get(e["action"], 0) + 1
print(len(bookings), "bookings in the plan,", len(events), "events:", acts)
print("still out at as-of:", out_now[0]["ref"] if out_now else None, "| over-booking attempt:", over, "on", target["out"])
