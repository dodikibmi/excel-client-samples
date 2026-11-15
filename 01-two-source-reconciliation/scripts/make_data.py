"""Generates six months (May-Oct 2026) of fictional data for "Lakeside Hardware Ltd" (made-up company), two bank accounts.

Writes deliberately messy CSV exports into ../data/ (one bank file and one cash book file per account per month):
  bank_<Account>_<yyyy-mm>.csv      - US bank export: MM/DD/YYYY dates, amounts as text with thousand separators,
                                      references typed inconsistently, a few rows exported twice
  cashbook_<Account>_<yyyy-mm>.csv  - accounting export: ISO dates, signed amounts
and the designed answer into ../test/expected.json, plus demo workbench notes into scripts/demo_notes.json.

What is planted on purpose (so the workbook can be checked against a known answer):
  - normal lines that match on the reference, some clearing 1-4 days later (also across a month end)
  - customer payments without a reference and bank charges / interest / direct debits that the cash book posts a few
    days later without a reference: matched by rule 3 (amount within tolerance, date within tolerance)
  - small FX/rounding differences within the amount tolerance: rule 2
  - real exceptions: amount differences, items only in the bank, only in the cash book, posted twice, paid twice
  - three old items still open (July, August, September) so the aging shows something over 30 days
"""
import calendar, csv, json, random
from datetime import date, timedelta
from pathlib import Path

random.seed(20261031)
ROOT = Path(__file__).resolve().parent.parent
MONTHS = [(2026, m) for m in range(5, 11)]          # May .. October 2026
OPENING = {"Operating": 18450.00, "Payroll": 6200.00}
AMOUNT_TOL, DAYS_TOL = 1.00, 5                       # also the defaults on the Settings sheet

CUSTOMERS = ["Brightwater Builders", "Oakridge Property Mgmt", "Henley Joinery", "Marsh & Co Plumbing",
             "Northgate Schools Trust", "Riverside Homes", "Coastal Roofing", "Pine Street Cafe",
             "Granite Kitchens", "Willow Landscapes", "Summit Electrical", "Harbour Rentals"]
SUPPLIERS = ["Atlas Fasteners", "BlueLine Timber", "Crown Paints", "Delta Tools Wholesale",
             "Evergreen Garden Supply", "Fortis Steel", "Metro Utilities", "City Council Rates"]
STAFF = [("Ava Reid", 3420), ("Ben Okafor", 2980), ("Chloe Martin", 4150), ("Dan Kowalski", 3610), ("Ella Moss", 2875),
         ("Finn Hayes", 3330), ("Grace Lin", 4480), ("Hugo Patel", 3065), ("Isla Grant", 3790), ("Jack Byrne", 2940),
         ("Kira Novak", 3555), ("Liam Doyle", 3205)]


def messy(ref):
    """How a bank clerk / customer types a reference: case, spaces and dashes vary."""
    head, num = ref.split("-", 1) if "-" in ref else (ref[:3], ref[3:])
    return random.choice([f"{head}-{num}", f"{head}{num.replace('-', '')}", f"{head.lower()}-{num}", f" {head} {num}",
                          f"{head}-{num} ", f"{head.lower()}{num.replace('-', '')}"])


def key(ref):
    return "".join(ch for ch in ref.upper() if ch not in " -/.") or None


bank = {a: [] for a in OPENING}      # dict(date, ref, desc, amount)
ledger = {a: [] for a in OPENING}    # dict(date, ref, party, amount, memo)
cases = []                           # the planted answer, one per reconciliation line
notes = {}                           # demo workbench notes, filled after the files are written (needs row numbers)
inv, po, chq = 1001, 5001, 4501


def case(account, status, rule="", **kw):
    c = dict(account=account, status=status, rule=rule, **kw)
    cases.append(c)
    return c


def add_pair(acct, d_led, gap, ref, party, led_amt, bank_amt, bank_desc, memo, status="Matched", rule="Reference",
             bank_ref=None, **kw):
    l = dict(date=d_led, ref=ref, party=party, amount=led_amt, memo=memo)
    b = dict(date=d_led + timedelta(days=gap), ref=messy(ref) if bank_ref is None else bank_ref, desc=bank_desc,
             amount=bank_amt)
    ledger[acct].append(l); bank[acct].append(b)
    c = case(acct, status, rule, key=key(ref), gap=gap, bank_amount=bank_amt, ledger_amount=led_amt, **kw)
    l["case"] = b["case"] = c
    return l, b


def bank_only(acct, d, desc, amt, ref="", **kw):
    b = dict(date=d, ref=ref, desc=desc, amount=amt)
    bank[acct].append(b)
    b["case"] = case(acct, "Only in bank", "", key=key(ref) if ref else None, bank_amount=amt, desc=desc, **kw)
    return b


def ledger_only(acct, d, ref, party, amt, memo, **kw):
    l = dict(date=d, ref=ref, party=party, amount=amt, memo=memo)
    ledger[acct].append(l)
    l["case"] = case(acct, "Only in ledger", "", key=key(ref), ledger_amount=amt, **kw)
    return l


def nobref_pair(acct, d_bank, lag, desc, amt, party, memo, led_ref=""):
    """Bank line without a reference, posted later in the cash book (no reference or a different one): rule 3."""
    b = dict(date=d_bank, ref="", desc=desc, amount=amt)
    l = dict(date=d_bank + timedelta(days=lag), ref=led_ref, party=party, amount=amt, memo=memo)
    bank[acct].append(b); ledger[acct].append(l)
    c = case(acct, "Matched", "Amount + date", key=None, gap=-lag, bank_amount=amt, ledger_amount=amt, desc=desc)
    b["case"] = l["case"] = c


for mi, (y, m) in enumerate(MONTHS):
    dim = calendar.monthrange(y, m)[1]
    d0 = date(y, m, 1)
    last = date(y, m, dim)
    nxt = last + timedelta(days=1)
    is_oct = (m == 10)
    mon = d0.strftime("%b")
    ym = f"{y}{m:02d}"

    # ------------------------------------------------------------------ Operating account
    A = "Operating"
    for _ in range(150 if is_oct else random.randint(96, 128)):
        d = d0 + timedelta(days=random.randint(0, dim - 5))
        gap = random.choice([0] * 8 + [1, 2, 3])
        if random.random() < 0.55:
            ref = f"INV-{inv}"; inv += 1
            party = random.choice(CUSTOMERS); amt = round(random.uniform(120, 4800), 2)
            add_pair(A, d, gap, ref, party, amt, amt, f"TRANSFER FROM {party.upper()}", "Customer receipt")
        else:
            if random.random() < 0.3:
                ref = f"CHQ-00{chq}"; chq += 1; memo = "Cheque payment"; desc = None
            else:
                ref = f"PO-{po}"; po += 1; memo = "Supplier payment"; desc = None
            party = random.choice(SUPPLIERS); amt = -round(random.uniform(80, 3900), 2)
            add_pair(A, d, gap, ref, party, amt, amt, f"CHEQUE {ref[4:]}" if ref.startswith("CHQ") else f"PAYMENT TO {party.upper()}", memo)

    # transfer to the payroll account (amount fixed below, once the payroll is known)
    trf_led, trf_bank = add_pair(A, date(y, m, 23), 0, f"TRF-{ym}", "Payroll account", 0.0, 0.0,
                                 "TRANSFER TO PAYROLL ACCOUNT", "Transfer to payroll")

    # two customers paid without quoting the invoice: the bank line has no reference -> rule 3
    for _ in range(2):
        party = random.choice(CUSTOMERS); amt = round(random.uniform(300, 2600), 2)
        d = d0 + timedelta(days=random.randint(2, dim - 8))
        ref = f"INV-{inv}"; inv += 1
        nobref_pair(A, d, random.choice([0, 1, 2]), f"TRANSFER FROM {party.upper()}", amt, party, "Customer receipt", ref)

    # a small FX / rounding difference on a foreign customer, within the amount tolerance -> rule 2
    if mi in (1, 3, 5):
        ref = f"INV-{inv}"; inv += 1
        led = round(random.uniform(900, 2400), 2); diff = {1: -0.40, 3: -0.75, 5: 0.30}[mi]
        add_pair(A, d0 + timedelta(days=random.randint(5, 20)), 1, ref, "Harbour Rentals", led, round(led + diff, 2),
                 "TRANSFER FROM HARBOUR RENTALS", "Customer receipt", rule="Reference ± amount", diff=diff)

    # direct debit for insurance: no reference at the bank; booked in the cash book 1-2 days later (not in October)
    if not is_oct:
        nobref_pair(A, date(y, m, 14), random.choice([1, 2]), "DIRECT DEBIT SHIELD INSURANCE", -189.00, "Shield Insurance",
                    "Insurance direct debit")

    if not is_oct:
        # month-end items that clear in the next month (still matched on the reference)
        for _ in range(3):
            d = last - timedelta(days=random.choice([0, 1]))
            gap = random.choice([2, 3, 4])
            if random.random() < 0.5:
                ref = f"INV-{inv}"; inv += 1; party = random.choice(CUSTOMERS); amt = round(random.uniform(200, 3000), 2)
                add_pair(A, d, gap, ref, party, amt, amt, f"TRANSFER FROM {party.upper()}", "Customer receipt")
            else:
                ref = f"CHQ-00{chq}"; chq += 1; party = random.choice(SUPPLIERS); amt = -round(random.uniform(150, 2500), 2)
                add_pair(A, d, gap, ref, party, amt, amt, f"CHEQUE {ref[4:]}", "Cheque payment")
        # bank charges and interest on the last day, posted in the cash book on the 1st-2nd of next month -> rule 3
        for desc, amt in [("ACCOUNT MAINTENANCE FEE", -25.00), (f"MERCHANT CARD FEE {mon.upper()}", -round(random.uniform(30, 45), 2)),
                          ("CREDIT INTEREST", round(random.uniform(2, 6), 2))]:
            nobref_pair(A, last, random.choice([1, 2]), desc, amt, "Bank",
                        f"Bank interest {mon}" if amt > 0 else f"Bank charges {mon}")

    # old items that are still open (aging over 30 days)
    if m == 7:
        bank_only(A, date(y, m, 18), "CASH DEPOSIT BRANCH 07", 275.00, note="unidentified")
    if m == 9:
        ref = f"INV-{inv}"; inv += 1
        add_pair(A, date(y, m, 11), 1, ref, "Granite Kitchens", 1845.00, 1800.00, "TRANSFER FROM GRANITE KITCHENS",
                 "Customer receipt", status="Amount differs", rule="", diff=-45.00, why="customer short-paid 45.00")

    if is_oct:
        # the October exceptions (same story as the one-month version of this sample)
        for ref, party, led, bnk, why in [
                (f"INV-{inv}", "Henley Joinery", 1000.00, 980.00, "customer deducted a 20.00 transfer fee"),
                (f"PO-{po}", "Fortis Steel", -1452.00, -1425.00, "ledger typo: digits swapped"),
                (f"INV-{inv + 1}", "Pine Street Cafe", 1557.50, 1575.50, "customer overpaid 18.00")]:
            add_pair(A, d0 + timedelta(days=random.randint(3, 25)), 1, ref, party, led, bnk,
                     f"TRANSFER FROM {party.upper()}" if bnk > 0 else f"PAYMENT TO {party.upper()}",
                     "Customer receipt" if led > 0 else "Supplier payment", status="Amount differs", rule="",
                     diff=round(bnk - led, 2), why=why)
        inv += 2; po += 1
        for d, desc, amt in [(30, "ACCOUNT MAINTENANCE FEE", -25.00), (30, "MERCHANT CARD FEE OCT", -38.40),
                             (30, "CREDIT INTEREST", 4.12), (14, "DIRECT DEBIT SHIELD INSURANCE", -189.00),
                             (19, "CASH DEPOSIT BRANCH 12", 350.00)]:
            bank_only(A, d0 + timedelta(days=d), desc, amt)
        for d, ref, party, amt, memo in [(30, f"INV-{inv}", "Northgate Schools Trust", 2140.00, "Customer receipt"),
                                         (28, f"CHQ-00{chq}", "BlueLine Timber", -860.00, "Cheque payment"),
                                         (29, f"CHQ-00{chq + 1}", "Crown Paints", -1215.50, "Cheque payment"),
                                         (30, f"PO-{po}", "Metro Utilities", -400.00, "Supplier payment")]:
            ledger_only(A, d0 + timedelta(days=d), ref, party, amt, memo)
        inv += 1; chq += 2; po += 1

    # ------------------------------------------------------------------ Payroll account
    P = "Payroll"
    tax = -round(random.uniform(5800, 6400), 2)
    add_pair(P, date(y, m, 9), 1, f"TAX-{ym}", "Revenue Service", tax, tax, "PAYMENT TO REVENUE SERVICE", "Payroll tax")
    pays = []
    for i, (name, base) in enumerate(STAFF, 1):
        amt = -round(base * random.uniform(0.97, 1.05), 2)
        pays.append(add_pair(P, date(y, m, 25), random.choice([0, 0, 1]), f"PAY-{ym}-{i:02d}", name, amt, amt,
                             f"SALARY {name.upper()}", "Salary"))
    total = round(-tax - sum(p[0]["amount"] for p in pays), 2)
    for rec in (trf_led, trf_bank):
        rec["amount"] = -total
    trf_led["case"].update(bank_amount=-total, ledger_amount=-total)
    add_pair(P, date(y, m, 23), 0, f"TRF-{ym}", "Operating account", total, total, "TRANSFER FROM OPERATING ACCOUNT",
             "Transfer from operating")
    if not is_oct:
        nobref_pair(P, last, 1, "PAYROLL ACCOUNT FEE", -12.00, "Bank", f"Bank charges {mon}")
    else:
        bank_only(P, last, "PAYROLL ACCOUNT FEE", -12.00)
    if m == 8:  # a casual worker's cheque that was never presented
        ledger_only(P, date(y, m, 12), "CHQ-00712", "Tom Walsh (casual)", -320.00, "Cheque payment", note="stale")
    if is_oct:  # the bank paid one salary twice (second time two days later)
        first = pays[6][1]
        dup = dict(date=first["date"] + timedelta(days=2), ref=messy(f"PAY-{ym}-07"), desc=first["desc"], amount=first["amount"])
        bank[P].append(dup)
        dup["case"] = case(P, "Duplicate in bank", "", key=key(f"PAY-{ym}-07"), bank_amount=first["amount"])

# Operating October: a receipt posted twice in the cash book (bank has it once)
octo = [l for l in ledger["Operating"] if l["date"].month == 10 and l["ref"].startswith("INV") and l["amount"] > 0
        and l["case"]["status"] == "Matched" and l["case"]["rule"] == "Reference"][:40]
dup = random.choice(octo)
dup_l = dict(date=dup["date"], ref=dup["ref"], party=dup["party"], amount=dup["amount"], memo="Customer receipt")
ledger["Operating"].append(dup_l)
dup_l["case"] = case("Operating", "Duplicate in ledger", "", key=key(dup["ref"]), ledger_amount=dup["amount"])

# ---------------------------------------------------------------------- write the files, one per account and month
for f in (ROOT / "data").glob("*.csv"):
    f.unlink()
(ROOT / "data").mkdir(exist_ok=True); (ROOT / "test").mkdir(exist_ok=True)
money = lambda v: f"{v:,.2f}"
files, ids = [], {}
repeat_in = {("Operating", 6): 1, ("Operating", 10): 2}   # export glitch: rows repeated verbatim
for acct in OPENING:
    bal = OPENING[acct]
    bank[acct].sort(key=lambda r: r["date"])
    ledger[acct].sort(key=lambda r: r["date"])
    for y, m in MONTHS:
        rows = [r for r in bank[acct] if (r["date"].year, r["date"].month) == (y, m)]
        for r in rows:
            bal = round(bal + r["amount"], 2)
            r["balance"] = bal
        out = list(rows)
        for r in random.sample([b for b in rows[10:100] if b["ref"]], repeat_in.get((acct, m), 0)):
            out.insert(out.index(r) + 1, dict(r, repeat=True))
        name = f"bank_{acct}_{y}-{m:02d}.csv"
        with open(ROOT / "data" / name, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Date", "Reference", "Description", "Debit", "Credit", "Balance"])
            for i, r in enumerate(out, 1):
                w.writerow([r["date"].strftime("%m/%d/%Y"), r["ref"], r["desc"],
                            money(-r["amount"]) if r["amount"] < 0 else "", money(r["amount"]) if r["amount"] > 0 else "",
                            money(r["balance"])])
                if not r.get("repeat"):
                    r["id"] = f"{acct[:3].upper()}-B-{y}{m:02d}-{i:03d}"
        files.append(dict(file=name, account=acct, period=f"{y}-{m:02d}", rows=len(out), repeated=len(out) - len(rows),
                          closing=rows[-1]["balance"]))
        lrows = [r for r in ledger[acct] if (r["date"].year, r["date"].month) == (y, m)]
        name = f"cashbook_{acct}_{y}-{m:02d}.csv"
        with open(ROOT / "data" / name, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Date", "Doc No", "Party", "Reference", "Amount", "Memo"])
            for i, r in enumerate(lrows, 1):
                r["doc"] = f"{'CB' if acct == 'Operating' else 'PB'}-{y % 100}{m:02d}-{i:03d}"
                r["id"] = f"{acct[:3].upper()}-L-{y}{m:02d}-{i:03d}"
                w.writerow([r["date"].isoformat(), r["doc"], r["party"], r["ref"], f"{r['amount']:.2f}", r["memo"]])
        files.append(dict(file=name, account=acct, period=f"{y}-{m:02d}", rows=len(lrows)))

# line ids back into the planted cases
for acct in OPENING:
    for r in bank[acct]:
        r["case"]["bank_id"] = r["id"]
    for r in ledger[acct]:
        r["case"]["ledger_id"] = r["id"]
for c in cases:
    c["date"] = None
for acct in OPENING:
    for r in bank[acct] + ledger[acct]:
        c = r["case"]
        c["date"] = min(c["date"] or "9999", r["date"].isoformat())
for c in cases:
    c["exception_id"] = c.get("bank_id") or c.get("ledger_id")

# ---------------------------------------------------------------------- demo notes typed on the Exceptions workbench
def find(**kw):
    hits = [c for c in cases if all(c.get(k) == v for k, v in kw.items())]
    assert len(hits) == 1, (kw, len(hits))
    return hits[0]["exception_id"]


notes = {
    find(account="Operating", status="Only in bank", note="unidentified"):
        dict(owner="Maya", comment="Asked branch 07 for the deposit slip (2nd reminder sent)", resolved=None),
    find(account="Operating", status="Amount differs", diff=-45.00):
        dict(owner="Sam", comment="Customer claims a 2.5% early-payment discount; waiting for approval", resolved=None),
    find(account="Payroll", status="Only in ledger", note="stale"):
        dict(owner="Maya", comment="Cheque not presented after 60+ days; contact the worker, then write back", resolved=None),
    find(account="Operating", status="Amount differs", diff=27.00):
        dict(owner="Sam", comment="Typo fixed: cash book corrected to 1,425.00", resolved="2026-11-03"),
    find(account="Operating", status="Duplicate in ledger"):
        dict(owner="Maya", comment="Second posting reversed", resolved="2026-11-02"),
    find(account="Operating", status="Amount differs", diff=-20.00):
        dict(owner="Sam", comment="Post 20.00 to bank charges", resolved=None),
    find(account="Operating", status="Only in bank", desc="ACCOUNT MAINTENANCE FEE"):
        dict(owner="Sam", comment="Posted with the October bank charges", resolved="2026-11-04"),
    find(account="Payroll", status="Duplicate in bank"):
        dict(owner="Maya", comment="Salary paid twice: recall requested from the bank", resolved=None),
}
json.dump(notes, open(ROOT / "scripts" / "demo_notes.json", "w", encoding="utf-8"), indent=1)

json.dump(dict(opening=OPENING, amount_tolerance=AMOUNT_TOL, days_tolerance=DAYS_TOL, months=[f"{y}-{m:02d}" for y, m in MONTHS],
               files=files, notes=notes, cases=cases),
          open(ROOT / "test" / "expected.json", "w", encoding="utf-8"), indent=1, default=str)
for acct in OPENING:
    print(acct, "bank lines", len(bank[acct]), "ledger lines", len(ledger[acct]))
from collections import Counter
print(Counter((c["account"], c["status"], c["rule"]) for c in cases))
