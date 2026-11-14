"""Generates the fictional customer exports for "Summit Ridge Outdoor" (made-up outdoor gear shop, October 2026).

Writes three deliberately messy sources into ../data/ and the designed answer into ../test/expected.json:
  webshop-customers-2026-10.csv  - online shop export: "Full Name" in any case with double spaces, emails with capitals
                                   and stray spaces, phones typed 4 ways, state written out ("Colorado"), ZIP+4,
                                   amounts as text ("$1,234.50"), people who registered twice, test accounts,
                                   rows exported twice
  pos-customers-2026-10.csv      - in-store till export: names as "Smith, Anna", nicknames (Bill, Kate), email often
                                   missing, ZIPs that lost their leading zero (2139), M/D/YYYY dates, walk-in rows
  sales-team-accounts.xlsx       - the sales team's old spreadsheet of club and trade accounts: title rows above the
                                   header, the company name in MERGED cells over its contacts, first and last name in
                                   two columns, work emails, dates typed as text or as real dates
Every person has a hidden id, so the workbook can be checked against a known answer.
"""
import csv, datetime as dt, json, random
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

random.seed(20261116)
ROOT = Path(__file__).resolve().parent.parent

NICK = {"JON": "JOHN", "JOHNNY": "JOHN", "BILL": "WILLIAM", "WILL": "WILLIAM", "BOB": "ROBERT", "ROB": "ROBERT",
        "LIZ": "ELIZABETH", "BETH": "ELIZABETH", "KATE": "KATHERINE", "KATIE": "KATHERINE", "MIKE": "MICHAEL",
        "JIM": "JAMES", "JIMMY": "JAMES", "TOM": "THOMAS", "DAN": "DANIEL", "DANNY": "DANIEL", "CHRIS": "CHRISTOPHER",
        "MATT": "MATTHEW", "NICK": "NICHOLAS", "SAM": "SAMUEL", "ALEX": "ALEXANDER", "TONY": "ANTHONY",
        "JEN": "JENNIFER", "JENNY": "JENNIFER", "MEG": "MARGARET", "PEGGY": "MARGARET", "STEVE": "STEVEN",
        "DAVE": "DAVID", "ANDY": "ANDREW", "BEN": "BENJAMIN", "JOE": "JOSEPH", "ABBY": "ABIGAIL", "GABE": "GABRIEL",
        "TED": "EDWARD", "ED": "EDWARD", "RICK": "RICHARD", "SUE": "SUSAN", "BECKY": "REBECCA", "NATE": "NATHAN"}
NICKS_OF = {}
for n, c in NICK.items():
    NICKS_OF.setdefault(c, []).append(n.title())

FIRST = sorted({c.title() for c in NICK.values()}) + [
    "Anna", "Maria", "Olivia", "Emma", "Sofia", "Lucas", "Ethan", "Noah", "Liam", "Ava", "Mia", "Grace", "Chloe", "Ella",
    "Hannah", "Leah", "Zoe", "Nora", "Ruby", "Isaac", "Owen", "Caleb", "Hugo", "Felix", "Oscar", "Leo", "Ivan", "Omar",
    "Ravi", "Priya", "Mei", "Kenji", "Yuki", "Aisha", "Fatima", "Diego", "Elena", "Lucia", "Marco", "Paolo", "Ingrid",
    "Lars", "Astrid", "Sven", "Carmen", "Rosa", "Pablo", "Tara", "Wendy", "Victor", "Hector", "Gloria", "Ivy", "Jonah",
    "Laura", "Paula", "Erin", "Megan", "Holly", "Dylan", "Logan", "Wyatt", "Claire", "Vera", "Iris", "Theo", "Milo"]
PAIR_ONLY = [("Christopher", "Christine"), ("Daniel", "Danielle"), ("Samuel", "Samantha"), ("Alexander", "Alexandra"),
             ("Joseph", "Josephine"), ("Nicholas", "Nicole")]
LAST = ["Abbott", "Barnes", "Castillo", "Dawson", "Ellis", "Fischer", "Gomez", "Hayes", "Iqbal", "Jensen", "Khan",
        "Lopez", "Morgan", "Nguyen", "Okafor", "Patel", "Quinn", "Reyes", "Silva", "Turner", "Usman", "Vance", "Walsh",
        "Young", "Zimmer", "Brooks", "Carter", "Diaz", "Evans", "Foster", "Garcia", "Hughes", "Ito", "Johansson", "Kim",
        "Larsen", "Mendez", "Novak", "Olsen", "Park", "Ramirez", "Sato", "Torres", "Ueda", "Vargas", "Weber", "Yamada",
        "Alvarez", "Bauer", "Chen", "Duarte", "Eriksen", "Ford", "Grant", "Hale", "Ingram", "Jordan", "Keller", "Lund",
        "Moreno", "Nash", "Ortiz", "Pruitt", "Russo", "Stone", "Tate", "Udall", "Voss", "Wong", "Yates", "Zhou", "Bishop",
        "Cole", "Dunn", "Ferris", "Gill", "Holt", "Irwin", "Kerr", "Lowe", "Marsh", "Noble", "Pike", "Rhodes", "Shaw",
        "Tran", "Wade", "Booth", "Crane", "Drake", "Fox", "Hart", "Lane", "Moss", "Price", "Reed", "Sharp", "Webb"]

CITIES = [  # city, state code, state name, ZIPs, area codes; Colorado is home turf
    ("Boulder", "CO", "Colorado", ["80301", "80302", "80303", "80304", "80305"], ["303", "720"]),
    ("Denver", "CO", "Colorado", ["80202", "80203", "80205", "80206", "80210"], ["303", "720"]),
    ("Fort Collins", "CO", "Colorado", ["80521", "80524", "80525"], ["970"]),
    ("Colorado Springs", "CO", "Colorado", ["80903", "80905", "80909"], ["719"]),
    ("Salt Lake City", "UT", "Utah", ["84101", "84102", "84105"], ["801"]),
    ("Seattle", "WA", "Washington", ["98101", "98103", "98105"], ["206"]),
    ("Portland", "OR", "Oregon", ["97201", "97205", "97209"], ["503"]),
    ("Oakland", "CA", "California", ["94607", "94610"], ["510"]),
    ("San Diego", "CA", "California", ["92101", "92103"], ["619"]),
    ("Phoenix", "AZ", "Arizona", ["85004", "85006"], ["602"]),
    ("Santa Fe", "NM", "New Mexico", ["87501", "87505"], ["505"]),
    ("Austin", "TX", "Texas", ["78701", "78704"], ["512"]),
    ("Chicago", "IL", "Illinois", ["60607", "60614"], ["312"]),
    ("Minneapolis", "MN", "Minnesota", ["55401", "55405"], ["612"]),
    ("Cambridge", "MA", "Massachusetts", ["02138", "02139", "02141"], ["617"]),
    ("Boston", "MA", "Massachusetts", ["02108", "02116"], ["617"]),
    ("New Haven", "CT", "Connecticut", ["06510", "06511"], ["203"]),
    ("Hoboken", "NJ", "New Jersey", ["07030"], ["201"]),
    ("Burlington", "VT", "Vermont", ["05401", "05403"], ["802"]),
    ("Portland", "ME", "Maine", ["04101", "04102"], ["207"]),
    ("Brooklyn", "NY", "New York", ["11201", "11211"], ["718"]),
]
WEIGHTS = [14, 12, 6, 5, 4, 4, 4, 3, 3, 3, 2, 3, 3, 3, 3, 2, 2, 2, 2, 2, 3]
STREETS = ["Pine", "Maple", "Cedar", "Aspen", "Spruce", "Willow", "Oak", "Birch", "Elm", "Juniper", "Canyon", "Mesa",
           "Ridge", "Summit", "Lake", "River", "Meadow", "Valley", "Highland", "Park", "Granite", "Alpine", "Sage"]
TYPES = [("Street", "St"), ("Avenue", "Ave"), ("Road", "Rd"), ("Drive", "Dr"), ("Lane", "Ln"), ("Court", "Ct"),
         ("Boulevard", "Blvd")]
DOMAINS = ["example.com", "example.net", "example.org"]

used_cf_last, used_init_last_zip, used_phone, used_email = set(), set(), set(), set()
persons = []


def canon(first):
    return NICK.get(first.upper(), first.upper())


def new_phone(area):
    while True:
        p = f"{area}555{random.randint(0, 9999):04d}"
        if p not in used_phone:
            used_phone.add(p)
            return p


def new_email(local, domain=None):
    base, n = local.lower(), 0
    while True:
        e = f"{base}{n if n else ''}@{domain or random.choice(DOMAINS)}"
        if e not in used_email:
            used_email.add(e)
            return e
        n += random.randint(1, 9)


def street():
    t = random.choice(TYPES)
    return dict(no=random.randint(10, 9899), name=random.choice(STREETS), type=t)


def person(first=None, last=None, city=None, zip_=None, allow=False):
    """allow=True: a planted look-alike, skips the uniqueness rules on purpose."""
    for _ in range(10000):
        c = city or random.choices(CITIES, WEIGHTS)[0]
        f = first or random.choice(FIRST)
        l = last or random.choice(LAST)
        z = zip_ or random.choice(c[3])
        k1, k2 = (canon(f), l.upper()), (canon(f)[0], l.upper(), z)
        if allow or (k1 not in used_cf_last and k2 not in used_init_last_zip):
            used_cf_last.add(k1); used_init_last_zip.add(k2)
            p = dict(id=len(persons) + 1, first=f, last=l, city=c[0], st=c[1], st_name=c[2], zip=z, addr=street(),
                     phone=new_phone(random.choice(c[4])),
                     email=new_email(random.choice([f"{f}.{l}", f"{f[0]}{l}", f"{f}{l[0]}", f"{f}_{l}"])),
                     optin=random.choice(["Yes"] * 6 + ["No"] * 4))
            persons.append(p)
            return p
    raise RuntimeError("could not place a person")


# ---------------------------------------------------------------- how people type things
def typed_first(p, nick_chance):
    opts = NICKS_OF.get(canon(p["first"]))
    return random.choice(opts) if opts and random.random() < nick_chance else p["first"]


def messy_case(s):
    return random.choice([s, s, s, s, s.upper(), s.lower()])


def addr_text(a):
    full, short = a["type"]
    return random.choice([f"{a['no']} {a['name']} {full}", f"{a['no']} {a['name']} {short}", f"{a['no']} {a['name']} {short}.",
                          f"{a['no']}  {a['name']} {full}", f"{a['no']} {a['name'].lower()} {short.lower()}.",
                          f"{a['no']} {a['name'].upper()} {full.upper()}"])


def phone_text(d):
    a, b, c = d[:3], d[3:6], d[6:]
    return random.choice([f"({a}) {b}-{c}", f"{a}-{b}-{c}", f"{a}.{b}.{c}", f"+1 {a}-{b}-{c}", f"{a}{b}{c}",
                          f"1-{a}-{b}-{c}", f"({a}){b}-{c}"])


def email_text(e):
    return random.choice([e] * 6 + [e.upper(), e.capitalize(), f" {e}", f"{e} ", e.replace("@", " @")])


def zip_text(z, lose_zero=False):
    if lose_zero and z.startswith("0"):
        return z[1:]  # Excel ate the leading zero before someone saved it as CSV
    return random.choice([z] * 7 + [f"{z}-{random.randint(1000, 9999)}"])


def day(a, b):
    d0 = dt.date(*a)
    return d0 + dt.timedelta(days=random.randint(0, (dt.date(*b) - d0).days))


# ---------------------------------------------------------------- the population
web_people = [person() for _ in range(975)]
pos_only = [person() for _ in range(200)]
look_alikes = []  # (person A, person B, reason)
for fa, fb in PAIR_ONLY:  # same last name and ZIP, similar first name: probably two people
    a = person(first=fa)
    c = next(x for x in CITIES if x[0] == a["city"] and x[1] == a["st"])
    b = person(first=fb, last=a["last"], city=c, zip_=a["zip"], allow=True)
    look_alikes.append((a, b, "Similar first name, same last name and ZIP"))
for _ in range(5):  # same name, different ZIP: two people, or someone who moved
    a = person()
    while True:
        c = random.choices(CITIES, WEIGHTS)[0]
        if c[3][0] != a["zip"] and (canon(a["first"])[0], a["last"].upper(), c[3][0]) not in used_init_last_zip:
            break
    b = person(first=a["first"], last=a["last"], city=c, zip_=c[3][0], allow=True)
    look_alikes.append((a, b, "Same name, different ZIP"))
web_people += [p for pair in look_alikes for p in pair[:2]]

records = []  # every row of every source, with the hidden person id


def web_row(p, email=None, created=None, nick=0.15):
    created = created or day((2022, 1, 1), (2026, 6, 30))
    last = created + dt.timedelta(days=random.randint(0, (dt.date(2026, 10, 31) - created).days))
    orders = random.randint(1, 14)
    spent = round(orders * random.uniform(35, 260), 2)
    name = f"{typed_first(p, nick)} {p['last']}"
    name = messy_case(name)
    if random.random() < 0.12:
        name = name.replace(" ", "  ")
    if random.random() < 0.1:
        name += " "
    return dict(src="W", pid=p["id"], spend=spent, last=last, row={
        "Customer No": None, "Full Name": name, "Email": email_text(email or p["email"]),
        "Phone": phone_text(p["phone"]) if random.random() < 0.93 else "",
        "Company": "", "Street": addr_text(p["addr"]), "City": messy_case(p["city"]),
        "State": random.choice([p["st"]] * 3 + [p["st_name"], p["st"].lower()]), "ZIP": zip_text(p["zip"]),
        "Created": created.isoformat(), "Last Order": last.isoformat(), "Orders": str(orders),
        "Total Spent": random.choice([f"${spent:,.2f}"] * 4 + [f"{spent:.2f}"]), "Marketing Opt-in": p["optin"]})


for p in web_people:
    records.append(web_row(p))
# 25 people registered a second time with another email (typo or new address); same phone
twice = random.sample(web_people[:975], 25)
for p in twice:
    r = web_row(p, email=new_email(f"{p['first']}{p['last']}{random.randint(1, 99)}"), nick=0.4)
    r["row"]["Phone"] = phone_text(p["phone"])
    records.append(r)
for r in records:  # the original account needs a phone too, or nothing links them
    if r["pid"] in {p["id"] for p in twice}:
        r["row"]["Phone"] = phone_text(persons[r["pid"] - 1]["phone"])
web_junk = [("Test Customer", "test@summitridge.example"), ("TEST ORDER do not ship", "orders@summitridge.example"),
            ("test test", "test@example.com"), ("Test  Account", "")]
for nm, em in web_junk:
    records.append(dict(src="W", pid="junk", spend=0.0, last=dt.date(2026, 10, 2), row={
        "Customer No": None, "Full Name": nm, "Email": em, "Phone": "", "Company": "", "Street": "1 Test Street",
        "City": "Boulder", "State": "CO", "ZIP": "80302", "Created": "2026-10-02", "Last Order": "2026-10-02",
        "Orders": "1", "Total Spent": "$0.00", "Marketing Opt-in": "No"}))
random.shuffle(records)
for i, r in enumerate(records, 1):
    r["row"]["Customer No"] = f"WS-{100000 + i * 7 + random.randint(0, 6)}"
    if random.random() < 0.04:
        r["row"]["Company"] = random.choice(["", "Self-employed", "n/a"])
for i in sorted(random.sample(range(len(records)), 3), reverse=True):  # the export repeats a row now and then
    records.insert(i + 1, dict(records[i], pid="dup"))
web = records

# ---------------------------------------------------------------- POS (till) export
pos = []


def pos_row(p, mode):
    """mode: email (email + maybe phone), phone (phone only), name (no email, no phone: name + ZIP only)."""
    f = typed_first(p, 0.3)
    name = random.choice([f"{p['last']}, {f}", f"{p['last'].upper()}, {f.upper()}", f"{f} {p['last']}", f"{p['last']},{f}"])
    email = email_text(p["email"]) if mode == "email" else ""
    phone = phone_text(p["phone"]) if mode == "phone" or (mode == "email" and random.random() < 0.8) else ""
    visit = day((2024, 1, 1), (2026, 10, 31))
    spend = round(random.uniform(15, 900), 2)
    return dict(src="P", pid=p["id"], spend=spend, last=visit, row={
        "Customer": name, "Phone": phone, "Email": email, "Address": addr_text(p["addr"]), "City": p["city"].upper(),
        "St": p["st"], "Zip": zip_text(p["zip"], lose_zero=random.random() < 0.8),
        "Last Visit": f"{visit.month}/{visit.day}/{visit.year}", "Visits": str(random.randint(1, 30)), "Spend": f"{spend:.2f}"})


both = random.sample([p for p in web_people[:975] if p not in twice], 420)
for p in both:
    pos.append(pos_row(p, random.choices(["email", "phone", "name"], [45, 40, 15])[0]))
for p in pos_only:
    pos.append(pos_row(p, random.choices(["email", "phone", "name"], [40, 55, 5])[0]))
again = random.sample(pos_only[:150], 15)  # came back and the till created a second card; phone links them
for p in again:
    pos.append(pos_row(p, "phone"))
for r in pos:
    if r["pid"] in {p["id"] for p in again}:
        r["row"]["Phone"] = phone_text(persons[r["pid"] - 1]["phone"])
for nm in ["WALK-IN", "Walk-in customer", "walk in", "", ""]:
    pos.append(dict(src="P", pid="junk", spend=round(random.uniform(5, 60), 2), last=dt.date(2026, 10, 15), row={
        "Customer": nm, "Phone": "", "Email": "", "Address": "", "City": "", "St": "", "Zip": "", "Last Visit": "10/15/2026",
        "Visits": "1", "Spend": None}))
    pos[-1]["row"]["Spend"] = f"{pos[-1]['spend']:.2f}"
for r in pos:  # walk-in and blank rows carry till money that belongs to no customer
    if r["pid"] == "junk":
        r["spend_junk"] = r["spend"]
random.shuffle(pos)
for i in sorted(random.sample(range(len(pos)), 2), reverse=True):
    pos.insert(i + 1, dict(pos[i], pid="dup"))

# ---------------------------------------------------------------- the sales team's spreadsheet
KIND = ["Outfitters", "Climbing Club", "Ski School", "Adventure Co.", "Mountain Guides", "Scout Troop 214", "Trail Runners",
        "Hiking Club", "Paddle Club", "Outdoor School"]
PRE = ["Peak", "Alpine", "Trailhead", "Canyon", "Riverbend", "High Desert", "Timberline", "Bluebird", "Red Rock",
       "Granite", "North Fork", "Switchback", "Basecamp", "Glacier", "Cottonwood", "Lakeside", "Frontier", "Wildflower"]
used_co = set()


def company_name():
    while True:
        n = f"{random.choice(PRE)} {random.choice(KIND)}"
        if n not in used_co:
            used_co.add(n)
            return n


companies = []  # (name, city tuple, zip, address, [(person, overlap?)])
overlap = random.sample([p for p in web_people[:975] if p not in twice and p not in both], 110)
by_zip = {}
for p in overlap:
    by_zip.setdefault((p["city"], p["st"], p["zip"]), []).append(p)
for (city, st, z), ps in by_zip.items():
    c = next(x for x in CITIES if x[0] == city and x[1] == st)
    while ps:
        chunk, ps = ps[:random.randint(1, 2)], ps[random.randint(1, 2):] if len(ps) > 2 else []
        companies.append([company_name(), c, z, street(), [(q, True) for q in chunk]])
sheet_only = 0
while sheet_only < 190:
    if random.random() < 0.6 and companies:
        co = random.choice(companies)
    else:
        c = random.choices(CITIES, WEIGHTS)[0]
        co = [company_name(), c, random.choice(c[3]), street(), []]
        companies.append(co)
    q = person(city=co[1], zip_=co[2])
    q["addr"] = co[3]
    co[4].append((q, False))
    sheet_only += 1
random.shuffle(companies)

sheet = []
for name, c, z, a, people in companies:
    slug = "".join(ch for ch in name.lower() if ch.isalnum())[:18]
    for p, ov in people:
        f = typed_first(p, 0.35 if ov else 0.2)
        if ov:
            mode = random.choices(["email", "phone", "name"], [30, 35, 35])[0]
        else:
            mode = "work"
        work = new_email(f"{p['first']}.{p['last']}", f"{slug}.example")
        email = p["email"] if mode == "email" else work
        phone = p["phone"] if mode == "phone" or (mode in ("email", "work") and random.random() < 0.5) else ""
        when = day((2023, 1, 1), (2026, 9, 30))
        sheet.append(dict(src="S", pid=p["id"], spend=0.0, last=when, company=name, row={
            "Company": name, "First": messy_case(f), "Last": messy_case(p["last"]), "Email": email_text(email),
            "Phone": phone_text(phone) if phone else "", "Address": addr_text(a), "City": c[0],
            "State": random.choice([c[1], c[1], c[2]]), "ZIP": zip_text(z), "Last contact": when,
            "Newsletter": random.choice(["Y", "N", ""]) if not ov else random.choice(["Y", "N", "", "N"])}))

# ---------------------------------------------------------------- write the files
(ROOT / "data").mkdir(exist_ok=True)
WCOLS = ["Customer No", "Full Name", "Email", "Phone", "Company", "Street", "City", "State", "ZIP", "Created",
         "Last Order", "Orders", "Total Spent", "Marketing Opt-in"]
with open(ROOT / "data" / "webshop-customers-2026-10.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f); w.writerow(WCOLS)
    for r in web:
        w.writerow([r["row"][c] for c in WCOLS])
PCOLS = ["Customer", "Phone", "Email", "Address", "City", "St", "Zip", "Last Visit", "Visits", "Spend"]
with open(ROOT / "data" / "pos-customers-2026-10.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f); w.writerow(PCOLS)
    for r in pos:
        w.writerow([r["row"][c] for c in PCOLS])

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Trade accounts"
ws["A1"] = "Summit Ridge Outdoor - club & trade accounts (sales team list)"
ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = "Keep this up to date!  Last cleanup: spring 2025"
ws["A2"].font = Font(italic=True, color="888888")
SCOLS = ["Company", "First", "Last", "Email", "Phone", "Address", "City", "State", "ZIP", "Last contact", "Newsletter"]
for j, c in enumerate(SCOLS, 1):
    cell = ws.cell(row=3, column=j, value=c)
    cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="2F5496")
r = 4
prev = None
for s in sheet:
    if s["company"] != prev:
        if prev is not None and random.random() < 0.5:
            r += 1  # a blank line between some companies
        start = r
        prev = s["company"]
    for j, c in enumerate(SCOLS, 1):
        v = s["row"][c]
        if c == "Company":
            v = v if r == start else None
        if c == "Last contact":
            v = v if random.random() < 0.6 else random.choice([f"{v.day} {v:%b %Y}", f"{v.day:02d} {v:%b %Y}"])
        if c == "ZIP" and isinstance(v, str):
            v = v
        cell = ws.cell(row=r, column=j, value=v)
        if c == "Last contact" and isinstance(v, dt.date):
            cell.number_format = "d-mmm-yy"
        if c == "ZIP":
            cell.number_format = "@"
    s["sheet_row"] = r
    r += 1
    nxt = sheet.index(s) + 1
    if nxt == len(sheet) or sheet[nxt]["company"] != s["company"]:
        if r - 1 > start:
            ws.merge_cells(start_row=start, start_column=1, end_row=r - 1, end_column=1)
        ws.cell(row=start, column=1).alignment = Alignment(vertical="center")
for j, wdt in enumerate([26, 12, 12, 34, 16, 26, 16, 12, 11, 13, 11], 1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = wdt
wb.save(ROOT / "data" / "sales-team-accounts.xlsx")

# ---------------------------------------------------------------- the planted answer
truth = {}
for i, r in enumerate(web, 1):
    truth[f"W-{i:04d}"] = r["pid"]
for i, r in enumerate(pos, 1):
    truth[f"P-{i:04d}"] = r["pid"]
for i, s in enumerate(sheet, 1):
    truth[f"S-{i:04d}"] = s["pid"]
real = [k for k, v in truth.items() if isinstance(v, int)]
people = sorted({truth[k] for k in real})
spend = round(sum(r["spend"] for r in web + pos if isinstance(r["pid"], int)), 2)
exp = dict(rows=dict(W=len(web), P=len(pos), S=len(sheet)), exact_duplicates=dict(W=3, P=2, S=0),
           junk=dict(W=len(web_junk), P=5, S=0), customers=len(people), merged_away=len(real) - len(people),
           lifetime_spend=spend, review_pairs=[[a["id"], b["id"], why] for a, b, why in look_alikes], truth=truth)
(ROOT / "test").mkdir(exist_ok=True)
(ROOT / "test" / "expected.json").write_text(json.dumps(exp, indent=0), encoding="utf-8")
print(f"rows: webshop {len(web)}, POS {len(pos)}, sales sheet {len(sheet)} = {len(web) + len(pos) + len(sheet)}; "
      f"customers {len(people)}, merged away {exp['merged_away']}, review pairs {len(look_alikes)}, spend {spend:,.2f}")
