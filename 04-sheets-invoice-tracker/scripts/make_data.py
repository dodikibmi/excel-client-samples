"""Fictional demo data for the Google Sheets invoice tracker ("Northbeam Studio", a made-up design agency).

Dates are stored as day offsets from the day the demo is built (day 0 = "today"), so the demo always shows
current-looking overdue invoices whenever someone makes a copy and runs "Rebuild demo data".

Writes ../src/DemoData.gs (loaded by the Apps Script project) and ../test/expected.json: the status, days overdue,
aging bucket and reminder plan worked out here in Python, independently of the JavaScript, for test/run-tests.js.
"""
import json, random
from pathlib import Path

random.seed(2026_1114)
ROOT = Path(__file__).resolve().parent.parent

CUSTOMERS = [  # name, contact, email, payer type
    ("Alder & Finch Bakery", "Maya Alder", "accounts@alderfinch.example.com", "prompt"),
    ("Brightwater Yoga", "Sam Ortega", "sam@brightwater.example.com", "prompt"),
    ("Cedar Lane Dental", "Dr. Ruth Kim", "billing@cedarlane.example.com", "slow"),
    ("Driftwood Coffee Co.", "Leo Marsh", "leo@driftwood.example.com", "slow"),
    ("Elmstead Property Group", "Hannah Price", "ap@elmstead.example.com", "late"),
    ("Fernhill Organics", "Omar Haddad", "omar@fernhill.example.com", "prompt"),
    ("Glasshouse Events", "Ivy Chen", "", "late"),  # no email on file: reminders must be skipped, not lost
    ("Harbor & Pine Law", "Daniel Brooks", "finance@harborpine.example.com", "slow"),
    ("Inkwell Bookshop", "Rosa Delgado", "rosa@inkwell.example.com", "late"),
    ("Juniper Fitness", "Kofi Mensah", "accounts@juniperfit.example.com", "prompt"),
    ("Kestrel Outdoor Gear", "Nora Lind", "ap@kestrelgear.example.com", "slow"),
    ("Lumen Architecture", "Victor Ames", "billing@lumenarch.example.com", "late"),
]
STAGES = [("Upcoming", -3), ("Friendly", 1), ("Second", 14), ("Final", 30)]
MAX_PER_RUN = 25


def round2(x):
    return round(x + 1e-9, 2)


invoices, n = [], 1001
for k in range(64):
    name, _, _, kind = CUSTOMERS[k % len(CUSTOMERS)]
    issue = -random.randint(2, 150)
    terms = random.choice([14, 30, 30])
    due = issue + terms
    amount = round2(random.choice([250, 480, 750, 900, 1200, 1850, 2400, 3200, 4750, 6400, 9800]) * random.uniform(0.85, 1.15))
    paid, paid_on = 0.0, None
    late_by = {"prompt": random.randint(-5, 3), "slow": random.randint(10, 40), "late": random.randint(45, 120)}[kind]
    if due + late_by <= 0:  # the customer has paid by now
        paid, paid_on = amount, min(0, due + late_by)
        if kind != "prompt" and random.random() < 0.25:  # paid only part
            paid = round2(amount * random.choice([0.3, 0.5]))
    invoices.append(dict(no=f"NB-{n}", customer=name, issue=issue, due=due, amount=amount, paid=paid, paid_on=paid_on, notes=""))
    n += 1
# a few cases that must always be there
invoices += [
    dict(no=f"NB-{n}", customer="Fernhill Organics", issue=-25, due=-25 + 30, amount=1500.00, paid=0, paid_on=None, notes="due in 5 days"),
    dict(no=f"NB-{n + 1}", customer="Brightwater Yoga", issue=-12, due=-12 + 14, amount=640.00, paid=0, paid_on=None, notes="due in 2 days: heads-up"),
    dict(no=f"NB-{n + 2}", customer="Juniper Fitness", issue=-40, due=-40 + 30, amount=820.00, paid=850.00, paid_on=-8, notes="paid 30 too much"),
    dict(no=f"NB-{n + 3}", customer="Cedar Lane Dental", issue=-44, due=-44 + 30, amount=2100.00, paid=0, paid_on=None, notes="exactly 14 days overdue"),
    dict(no=f"NB-{n + 4}", customer="Elmstead Property Group", issue=-60, due=-60 + 30, amount=3900.00, paid=0, paid_on=None, notes="exactly 30 days overdue"),
]
invoices.sort(key=lambda i: (i["issue"], i["no"]))

# Reminders already sent before today (the log): the script must not send the same stage twice
log = []
for inv in invoices:
    bal = round2(inv["amount"] - inv["paid"])
    rel = -inv["due"]  # days overdue as of today
    if bal > 0 and rel >= 1 and random.random() < 0.6:
        reached = [(s, d) for s, d in STAGES if d >= 1 and rel - 3 >= d]  # sent a few days ago
        for s, d in reached[-1:]:
            log.append(dict(day=inv["due"] + d, no=inv["no"], customer=inv["customer"], stage=s))
cust_email = {c[0]: c[2] for c in CUSTOMERS}
for e in log:
    e["email"] = cust_email[e["customer"]]
log.sort(key=lambda e: e["day"])


# -------------------------------------------------------------------- expected results, worked out in Python
def state(inv):
    bal = round2(inv["amount"] - inv["paid"])
    rel = -inv["due"]
    if bal < 0:
        return dict(balance=bal, status="Overpaid", days=0, aging="Paid")
    if bal == 0:
        return dict(balance=0.0, status="Paid", days=0, aging="Paid")
    days = max(0, rel)
    status = "Overdue" if rel > 0 else ("Part paid" if inv["paid"] > 0 else "Not due")
    aging = "Not due" if days == 0 else "1-30" if days <= 30 else "31-60" if days <= 60 else "61-90" if days <= 90 else "90+"
    return dict(balance=bal, status=status, days=days, aging=aging)


sent = {}
for e in log:
    sent.setdefault(e["no"], set()).add(e["stage"])
order = {s: d for s, d in STAGES}
plan = []
for inv in sorted(invoices, key=lambda i: (i["due"], i["no"])):  # most overdue first
    st = state(inv)
    if st["balance"] <= 0:
        continue
    rel = -inv["due"]
    reached = [s for s, d in STAGES if rel >= d]
    if not reached:
        continue
    target = reached[-1]
    if any(order[s] >= order[target] for s in sent.get(inv["no"], ())):
        continue
    plan.append(dict(no=inv["no"], stage=target, action="skip" if not cust_email[inv["customer"]] else "send",
                     reason="No email for this customer" if not cust_email[inv["customer"]] else ""))
sends = 0
for p in plan:
    if p["action"] == "send":
        sends += 1
        if sends > MAX_PER_RUN:
            p["action"], p["reason"] = "hold", "Run limit reached"

states = {i["no"]: state(i) for i in invoices}
summary = dict(outstanding=round2(sum(s["balance"] for s in states.values() if s["balance"] > 0)),
               overdue=round2(sum(s["balance"] for s in states.values() if s["status"] == "Overdue")),
               overdue_count=sum(1 for s in states.values() if s["status"] == "Overdue"),
               aging={b: round2(sum(s["balance"] for s in states.values() if s["aging"] == b and s["balance"] > 0))
                      for b in ["Not due", "1-30", "31-60", "61-90", "90+"]})

demo = dict(customers=[dict(name=c[0], contact=c[1], email=c[2]) for c in CUSTOMERS],
            invoices=invoices, log=log, stages=[dict(name=s, day=d) for s, d in STAGES], maxPerRun=MAX_PER_RUN)
(ROOT / "src").mkdir(exist_ok=True); (ROOT / "test").mkdir(exist_ok=True)
(ROOT / "src" / "DemoData.gs").write_text(
    "// Generated by scripts/make_data.py - fictional demo data. Dates are day offsets from the day the demo is built.\n"
    "const DEMO = " + json.dumps(demo, indent=1) + ";\n", encoding="utf-8")
json.dump(dict(states=states, plan=plan, summary=summary), open(ROOT / "test" / "expected.json", "w"), indent=1)
print(f"{len(invoices)} invoices, {len(log)} past reminders, plan: "
      f"{sum(p['action'] == 'send' for p in plan)} send / {sum(p['action'] == 'skip' for p in plan)} skip / "
      f"{sum(p['action'] == 'hold' for p in plan)} hold; outstanding {summary['outstanding']:,.2f}, "
      f"overdue {summary['overdue']:,.2f} ({summary['overdue_count']})")
