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
# Older, fully paid history (Mar-Jun) so the 6-month trends, DSO and each customer's payment habit have data.
# Own random generator: the invoices above stay exactly as they were.
hist = random.Random(2026_0601)
for k in range(20):
    name, _, _, kind = CUSTOMERS[(k * 5) % len(CUSTOMERS)]
    issue = -hist.randint(150, 250)
    due = issue + hist.choice([14, 30, 30])
    amount = round2(hist.choice([480, 750, 1200, 1850, 2400, 3200, 4750]) * hist.uniform(0.9, 1.1))
    late_by = {"prompt": hist.randint(-4, 3), "slow": hist.randint(8, 35), "late": hist.randint(40, 80)}[kind]
    invoices.append(dict(no=f"NB-{981 + k:04d}", customer=name, issue=issue, due=due, amount=amount, paid=amount,
                         paid_on=min(-1, due + late_by), notes=""))
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
    aging = "Not due" if days == 0 else "1-30" if days <= 30 else "31-60" if days <= 60 else "61-90" if days <= 90 else "91+"
    return dict(balance=bal, status=status, days=days, aging=aging)


STATUSES = ["Active", "On hold", "Collections"]   # On hold / Collections: no automatic reminders for that customer
PAUSE_DAYS = 7                                      # a reply pauses reminders for this many days
cust_status = {c[0]: "Active" for c in CUSTOMERS}

sent = {}
for e in log:
    sent.setdefault(e["no"], set()).add(e["stage"])
order = {s: d for s, d in STAGES}


def make_plan(pauses):
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
        email = cust_email[inv["customer"]]
        p = dict(no=inv["no"], stage=target, action="send", reason="")
        pz = pauses.get(inv["no"], {})
        if cust_status[inv["customer"]] != "Active":
            p.update(action="pause", reason="Customer status: " + cust_status[inv["customer"]])
        elif pz.get("paused"):
            p.update(action="pause", reason=pz["reason"])
        elif not email:
            p.update(action="skip", reason="No email for this customer")
        plan.append(p)
    sends = 0
    for p in plan:
        if p["action"] == "send":
            sends += 1
            if sends > MAX_PER_RUN:
                p["action"], p["reason"] = "hold", "Run limit reached"
    return plan


base_plan = make_plan({})
by_no = {i["no"]: i for i in invoices}

# Replies and promises to pay, logged by the owner (Replies sheet). Picked from today's plan so the pause rule shows.
first = lambda stage: next(p["no"] for p in base_plan if p["action"] == "send" and p["stage"] == stage)
promise_no, reply_no = first("Second"), first("Friendly")
old_reply_no = [p["no"] for p in base_plan if p["action"] == "send" and p["stage"] == "Final"][1]
in_plan = {p["no"] for p in base_plan}
broken_no = max((i for i in invoices if state(i)["status"] == "Overdue" and i["no"] not in in_plan and state(i)["days"] <= 60
                 and cust_email[i["customer"]]), key=lambda i: state(i)["balance"])["no"]
replies = [
    dict(day=-20, no=old_reply_no, type="Reply", promised=None, note="Asked for a copy of the invoice (sent the same day)"),
    dict(day=-12, no=broken_no, type="Promise to pay", promised=-4, note="Promised to pay in full by bank transfer"),
    dict(day=-2, no=reply_no, type="Reply", promised=None, note="Checking with their accounts team, will confirm a date"),
    dict(day=-1, no=promise_no, type="Promise to pay", promised=5, note="Will pay on their next payment run"),
]
for r in replies:
    r["customer"] = by_no[r["no"]]["customer"]


def fmt_day(off):
    import datetime as _dt
    d = ANCHOR + _dt.timedelta(days=off)
    return f"{d.day} {d.strftime('%b')} {d.year}"


def pause_state(asof_off=0):
    """Latest reply/promise per invoice -> paused / broken (same rules as pauseState in Code.gs)."""
    latest = {}
    for i, r in enumerate(replies):
        k = r["no"]
        if k not in latest or (r["day"], i) >= (latest[k][0]["day"], latest[k][1]):
            latest[k] = (r, i)
    out = {}
    for k, (r, _) in latest.items():
        if r["type"] == "Promise to pay" and r["promised"] is not None:
            if r["promised"] >= asof_off:
                out[k] = dict(paused=True, broken=False, type=r["type"], promised=r["promised"],
                              reason="Promised to pay by " + fmt_day(r["promised"]))
            else:
                out[k] = dict(paused=False, broken=True, type=r["type"], promised=r["promised"],
                              reason="Promise to pay by " + fmt_day(r["promised"]) + " not kept")
        else:
            paused = asof_off - r["day"] <= PAUSE_DAYS
            out[k] = dict(paused=paused, broken=False, type="Reply", promised=None,
                          reason=("Customer replied on " + fmt_day(r["day"])) if paused else "")
    return out


import datetime as dt
ANCHOR = dt.date(2026, 11, 14)  # test/run-tests.js uses the same day as "today"
pauses = pause_state()
plan = make_plan(pauses)

states = {i["no"]: state(i) for i in invoices}
summary = dict(outstanding=round2(sum(s["balance"] for s in states.values() if s["balance"] > 0)),
               overdue=round2(sum(s["balance"] for s in states.values() if s["status"] == "Overdue")),
               overdue_count=sum(1 for s in states.values() if s["status"] == "Overdue"),
               aging={b: round2(sum(s["balance"] for s in states.values() if s["aging"] == b and s["balance"] > 0))
                      for b in ["Not due", "1-30", "31-60", "61-90", "91+"]})



# ---- trends, DSO, collection rate (dates as day offsets; month ends worked out from the anchor date)
def off(d):
    return (d - ANCHOR).days


def eom(k):  # last day of the month k months before the anchor month (k = 0: this month)
    y, m = ANCHOR.year, ANCHOR.month - k
    while m < 1:
        y, m = y - 1, m + 12
    nxt = dt.date(y + (m == 12), m % 12 + 1, 1)
    return nxt - dt.timedelta(days=1)


def pay_day(i):
    return i["paid_on"] if i["paid_on"] is not None else 0


def metrics_at(p):
    """p = day offset. Outstanding, overdue, DSO (90 days) and collection rate (invoices due in the 90 days to p)."""
    out = ovd = sales = due_amt = due_coll = 0.0
    for i in invoices:
        paid_by = i["paid"] if i["paid"] and pay_day(i) <= p else 0.0
        bal = i["amount"] - paid_by
        if i["issue"] <= p and bal > 0.0000001:
            out += bal
            if i["due"] < p:
                ovd += bal
        if p - 90 < i["issue"] <= p:
            sales += i["amount"]
        if p - 90 < i["due"] <= p:
            due_amt += i["amount"]
            due_coll += min(paid_by, i["amount"])
    return dict(outstanding=round2(out), overdue=round2(ovd), dso=round(out / sales * 90, 4) if sales else 0.0,
                rate=round(due_coll / due_amt, 6) if due_amt else 0.0)


def collected(a, b):
    return round2(sum(i["paid"] for i in invoices if i["paid"] and a <= pay_day(i) <= b))


points = []
for k in range(5, -1, -1):
    end = off(eom(k)) if k else 0
    start = off(eom(k).replace(day=1))
    billed = round2(sum(i["amount"] for i in invoices if start <= i["issue"] <= end))
    points.append(dict(label=eom(k).strftime("%b"), end=end, collected=collected(start, end), billed=billed, **metrics_at(end)))
prev_day = min(ANCHOR.day, eom(1).day)
prev_mtd = collected(off(eom(1).replace(day=1)), off(eom(1).replace(day=prev_day)))
kpis = dict(collected_mtd=points[-1]["collected"], collected_prev_mtd=prev_mtd, dso=points[-1]["dso"], rate=points[-1]["rate"],
            dso_prev=points[-2]["dso"], rate_prev=points[-2]["rate"], outstanding_prev=points[-2]["outstanding"],
            overdue_prev=points[-2]["overdue"], cash_collected=round2(sum(i["paid"] for i in invoices if i["paid"])))

# ---- customer risk score (0-100): payment habit, how late now, how much is overdue, no email
risk = {}
for c in CUSTOMERS:
    name = c[0]
    mine = [i for i in invoices if i["customer"] == name]
    late = [max(0, i["paid_on"] - i["due"]) for i in mine if state(i)["balance"] <= 0 and i["paid_on"] is not None]
    avg = sum(late) / len(late) if late else 0.0
    maxd = max([state(i)["days"] for i in mine] or [0])
    ovd = round2(sum(state(i)["balance"] for i in mine if state(i)["status"] == "Overdue"))
    score = int(min(40, avg * 0.5) + min(30, maxd * 0.3) + min(20, ovd / 500) + (10 if not c[2] else 0) + 0.5)
    risk[name] = dict(avg_late=round(avg, 4), max_days=maxd, overdue=ovd, score=score,
                      level="High" if score >= 60 else "Medium" if score >= 30 else "Low")

# ---- owner action queue: one action per open invoice, ranked
RANK = {"Urgent": 0, "High": 1, "Medium": 2, "Low": 3}
STAGE_PRI = {"Final": "High", "Second": "High", "Friendly": "Medium", "Upcoming": "Low"}
plan_by = {p["no"]: p for p in plan}
queue = []
for i in invoices:
    st = state(i)
    if st["balance"] <= 0:
        continue
    pz, pl = pauses.get(i["no"]), plan_by.get(i["no"])
    a = None
    if pz and pz["broken"]:
        a = ("Urgent", "Call: promise not kept")
    elif cust_status[i["customer"]] != "Active":
        a = None
    elif pz and pz["paused"] and pz["type"] == "Reply":
        a = ("Medium", "Record a promise to pay")
    elif pz and pz["paused"]:
        a = ("Low", "Check the promised payment") if pz["promised"] <= 2 else None
    elif st["days"] > 60:
        a = ("Urgent", "Call: over 60 days late")
    elif pl and pl["action"] == "skip":
        a = ("High", "Call: no email on file")
    elif pl and pl["action"] == "send":
        a = (STAGE_PRI[pl["stage"]], "Send " + pl["stage"] + " reminder")
    elif pl and pl["action"] == "hold":
        a = ("Low", "Send " + pl["stage"] + " reminder (tomorrow)")
    if a:
        queue.append(dict(priority=a[0], action=a[1], no=i["no"], customer=i["customer"], balance=st["balance"], days=st["days"]))
queue.sort(key=lambda q: (RANK[q["priority"]], -q["balance"], q["no"]))

demo = dict(customers=[dict(name=c[0], contact=c[1], email=c[2], status=cust_status[c[0]]) for c in CUSTOMERS],
            invoices=invoices, log=log, replies=replies, stages=[dict(name=s, day=d) for s, d in STAGES],
            maxPerRun=MAX_PER_RUN, pauseDays=PAUSE_DAYS)
(ROOT / "src").mkdir(exist_ok=True); (ROOT / "test").mkdir(exist_ok=True)
(ROOT / "src" / "DemoData.gs").write_text(
    "// Generated by scripts/make_data.py - fictional demo data. Dates are day offsets from the day the demo is built.\n"
    "const DEMO = " + json.dumps(demo, indent=1) + ";\n", encoding="utf-8")
json.dump(dict(states=states, plan=plan, summary=summary, pauses=pauses, points=points, kpis=kpis, risk=risk, queue=queue), open(ROOT / "test" / "expected.json", "w"), indent=1)
print(f"queue {len(queue)}: " + ", ".join(f"{q['priority']} {q['no']}" for q in queue[:8]))
print("risk", {k: v["score"] for k, v in risk.items()})
print("points", [(x["label"], x["outstanding"], x["collected"], round(x["dso"], 1), round(x["rate"], 3)) for x in points])
print("kpis", kpis)
print(f"{len(invoices)} invoices, {len(log)} past reminders, plan: "
      f"{sum(p['action'] == 'send' for p in plan)} send / {sum(p['action'] == 'skip' for p in plan)} skip / "
      f"{sum(p['action'] == 'hold' for p in plan)} hold / {sum(p['action'] == 'pause' for p in plan)} pause; outstanding {summary['outstanding']:,.2f}, "
      f"overdue {summary['overdue']:,.2f} ({summary['overdue_count']})")
