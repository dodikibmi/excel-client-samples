"""Generates the fictional Aug-Oct 2026 bank and card exports for "Copperleaf Bakery & Cafe" (made-up business).

Writes four deliberately messy CSV files into ../data/ and the designed answer into ../test/expected.json:
  checking-2026-08-10.csv   - bank export: 3 title lines above the header, MM/DD/YYYY, Debit/Credit as text "$1,250.00"
  savings-2026-08-10.csv    - ISO dates, one signed Amount column with thousand separators
  visa-2026-08-10.csv       - card export: Transaction and Post Date (M/D/YYYY), signed Amount (charges negative),
                              the card issuer's own "Category" (not the bakery's chart of accounts)
  statements-2026-08-10.csv - the month-end balance printed on each statement
Merchant names are typed several ways. Every line carries its TRUE account here (never written to the CSV), so the
workbook's rules are checked against a known answer. Planted: 3 lines no rule can know, a Sep 30 transfer that
reaches savings on Oct 1, an Oct 30 card payment that reaches the card after the period, a charge posted a day later
in the next month, and three pay dates in October.
"""
import csv, json, random, sys
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from coa import ACCOUNTS, MONTHS, ADJUSTMENTS, CLEARING, UNCATEGORIZED

random.seed(20261115)
ROOT = Path(__file__).resolve().parent.parent
START, END = date(2026, 8, 1), date(2026, 10, 31)
lines = []  # dict(reg, date, post, desc, amount (debit +), acct, category)


def add(reg, d, desc, amount, acct, post=None, category=""):
    lines.append(dict(reg=reg, date=d, post=post or d, desc=desc, amount=round(amount, 2), acct=acct, category=category))


def v(*options):
    return random.choice(options)


def days():
    d = START
    while d <= END:
        yield d
        d += timedelta(days=1)


# ---------------------------------------------------------------- checking (1000)
for d in days():
    wd = d.weekday()  # Mon 0 .. Sun 6; closed Sunday, TapPay settles the next business day, nothing on Monday
    if wd in (0, 6):
        continue
    amt = random.uniform(1450, 1950) if wd in (5, 1) else random.uniform(1000, 1500)  # Tuesday carries the weekend
    add(1000, d, v(f"TAPPAY DEPOSIT {d:%m%d}", "TapPay Inc. Deposit", f"TAP PAY *DEPOSIT {random.randint(10000, 99999)}"), amt, 4000)
wholesale = [("Mobile Deposit - Cafe Bloom", "CAFE BLOOM ACH CREDIT"), ("ACH CREDIT BLUE HERON MARKET", "Blue Heron Mkt payment"),
             ("DEPOSIT CHECK GREENLEAF GROCERS", "Greenleaf Grocers catering")]
for m in (8, 9, 10):
    for k in range(5 if m == 10 else 4):
        add(1000, date(2026, m, 3 + k * 6), v(*random.choice(wholesale)), random.uniform(600, 1300), 4100)
    for k in range(4):
        add(1000, date(2026, m, 2 + k * 7), v("NORTHGATE FOODSVC ACH", "Northgate Food Service", "NORTHGATE FOODSERVICE INV"),
            -random.uniform(1800, 2400), 5000)
    for k in (9, 23):
        add(1000, date(2026, m, k), v("VALLEY FLOUR MILL ACH", "Valley Flour Mill Co"), -random.uniform(650, 950), 5000)
    add(1000, date(2026, m, 1), v("MAPLE ST PROPERTIES RENT", "Maple St Properties LLC"), -4200.00, 6100)
    add(1000, date(2026, m, 14), v("CITY POWER & LIGHT AUTOPAY", "CITY PWR&LIGHT AUTOPAY"), -random.uniform(600, 720), 6200)
    add(1000, date(2026, m, 18), "METRO WATER DEPT", -random.uniform(140, 170), 6200)
    add(1000, date(2026, m, 5), "SHIELD MUTUAL INS PREM", -385.00, 6500)
    add(1000, date(2026, m, 28), v("MONTHLY SERVICE FEE", "SERVICE CHARGE"), -15.00, 6800)
    add(1000, date(2026, m, 3), f"TAPPAY MONTHLY FEES {date(2026, m - 1, 1):%b}".upper(), -random.uniform(960, 1060), 6800)
    add(1000, date(2026, m, 20), "ONLINE TRANSFER TO PERSONAL ...5590", -3000.00, 3100)
# payroll every other Friday from 7 Aug: three pay dates fall in October
p = date(2026, 8, 7)
while p <= END:
    add(1000, p, "PAYPILOT PAYROLL NET PAY", -random.uniform(5100, 5450), 6000)
    add(1000, p, v("PAYPILOT TAX IMPOUND", "PayPilot Tax"), -random.uniform(1150, 1290), 6000)
    p += timedelta(days=14)
# checking -> savings: Aug 28 same day, Sep 30 reaches savings Oct 1, Oct 29 same day
for d, arrive in ((date(2026, 8, 28), date(2026, 8, 28)), (date(2026, 9, 30), date(2026, 10, 1)), (date(2026, 10, 29), date(2026, 10, 29))):
    add(1000, d, "ONLINE TRANSFER TO SAV ...4821", -2000.00, CLEARING)
    add(1010, arrive, "TRANSFER FROM CHK ...1180", 2000.00, CLEARING)
# lines no rule can categorise
add(1000, date(2026, 8, 19), "VENMO *JKELLY", -120.00, UNCATEGORIZED)
add(1000, date(2026, 9, 24), "CHECK #1043", -350.00, UNCATEGORIZED)

# ---------------------------------------------------------------- savings (1010)
for _, _, end in MONTHS:
    add(1010, end, "INTEREST PAYMENT", round(random.uniform(17, 23), 2), 4900)

# ---------------------------------------------------------------- card (2000): charges negative, payments positive
card_merchants = [
    (5000, ("COSTWISE WHOLESALE #212", "COSTWISE WHSE 212"), (250, 600), 3, "Groceries"),
    (5100, ("PACKRIGHT SUPPLY", "PACK RIGHT SUPPLY CO"), (180, 420), 2, "Shopping"),
    (6400, ("PIXELREACH ADS", "PIXELREACH*ADS"), (120, 160), 2, "Professional Services"),
    (6950, ("PAPERCLIP OFFICE #88", "PAPERCLIP OFFICE"), (40, 90), 1, "Shopping"),
]
for m in (8, 9, 10):
    for acct, names, (lo, hi), n, cat in card_merchants:
        for _ in range(n):
            d = date(2026, m, random.randint(1, 27))
            add(2000, d, v(*names), -random.uniform(lo, hi), acct, d + timedelta(days=random.randint(0, 2)), cat)
    d = date(2026, m, 9)
    add(2000, d, "BRIGHTWAVE*INTERNET", -89.99, 6200, d + timedelta(days=1), "Bills & Utilities")
    d = date(2026, m, 15)
    add(2000, d, "CLOUDPOS* SUBSCRIPTION", -69.00, 6700, d, "Bills & Utilities")
    add(2000, d, "LEDGERLY SOFTWARE", -30.00, 6700, d + timedelta(days=1), "Bills & Utilities")
d = date(2026, 9, 11)
add(2000, d, "OVEN DOCTOR REPAIRS", -486.50, 6600, d + timedelta(days=1), "Home")  # one-off repair: September over budget
add(2000, date(2026, 8, 31), "COSTWISE WHOLESALE #212", -312.44, 5000, date(2026, 9, 1), "Groceries")  # bought 31 Aug, posted 1 Sep
d = date(2026, 10, 21)
add(2000, d, "POS PURCHASE 4471 MAIN ST", -64.18, UNCATEGORIZED, d + timedelta(days=1), "Shopping")

# card payments: on the 12th, pay last month's statement balance; the card side posts the same day
opening = {a[0]: round(a[3] - a[4], 2) for a in ACCOUNTS}
card_bal = opening[2000]  # credit balance, negative
for no, label, end in MONTHS:
    owed = -round(card_bal, 2)
    pay = date(end.year, end.month, 12)
    add(1000, pay, "ONLINE PAYMENT TO VISA ...7702", -owed, CLEARING)
    add(2000, pay, v("PAYMENT THANK YOU", "AUTOPAY PAYMENT - THANK YOU"), owed, CLEARING, pay, "Payment")
    card_bal = opening[2000] + sum(l["amount"] for l in lines if l["reg"] == 2000 and l["post"] <= end)
# Oct 30: extra payment from checking; the card company posts it on 2 Nov (outside these files) -> in transit
add(1000, date(2026, 10, 30), "ONLINE PAYMENT TO VISA ...7702", -1800.00, CLEARING)


# ---------------------------------------------------------------- write the messy exports
def money(x):
    return f"${x:,.2f}"


lines.sort(key=lambda l: (l["post"], l["reg"], -l["amount"]))
data = ROOT / "data"
data.mkdir(exist_ok=True)
bal = dict(opening)
with open(data / "checking-2026-08-10.csv", "w", newline="", encoding="utf-8") as f:
    f.write("Harbor Community Bank - Business Checking ...1180\r\nStatement export 08/01/2026 - 10/31/2026\r\n\r\n")
    w = csv.writer(f)
    w.writerow(["Date", "Description", "Debit", "Credit", "Balance"])
    for l in (l for l in lines if l["reg"] == 1000):
        bal[1000] += l["amount"]
        w.writerow([f"{l['date']:%m/%d/%Y}", l["desc"] + v("", "", " "), money(-l["amount"]) if l["amount"] < 0 else "",
                    money(l["amount"]) if l["amount"] > 0 else "", money(bal[1000])])
with open(data / "savings-2026-08-10.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Posted", "Details", "Amount"])
    for l in (l for l in lines if l["reg"] == 1010):
        w.writerow([f"{l['date']:%Y-%m-%d}", l["desc"], f"{l['amount']:,.2f}"])
with open(data / "visa-2026-08-10.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Transaction Date", "Post Date", "Description", "Category", "Type", "Amount"])
    for l in (l for l in lines if l["reg"] == 2000):
        w.writerow([f"{l['date'].month}/{l['date'].day}/{l['date'].year}", f"{l['post'].month}/{l['post'].day}/{l['post'].year}",
                    l["desc"], l["category"], "Payment" if l["amount"] > 0 else "Sale", f"{l['amount']:.2f}"])

# statement balances (the card statement prints what is owed, as a positive number)
stmt_names = {1000: "Business Checking ...1180", 1010: "Business Savings ...4821", 2000: "Visa Business ...7702"}
statements = {}
with open(data / "statements-2026-08-10.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Account", "Statement date", "Ending balance"])
    for no, label, end in MONTHS:
        for reg in (1000, 1010, 2000):
            b = round(opening[reg] + sum(l["amount"] for l in lines if l["reg"] == reg and l["post"] <= end), 2)
            shown = -b if reg == 2000 else b
            statements[f"{reg}-{no}"] = shown
            w.writerow([stmt_names[reg], f"{end:%m/%d/%Y}", f"{shown:,.2f}"])

# ---------------------------------------------------------------- the planted answer
types = {a[0]: a[2] for a in ACCOUNTS}
net = defaultdict(float)  # (account, month) -> debit-positive movement
for l in lines:
    mo = l["post"].year * 100 + l["post"].month
    net[(l["reg"], mo)] += l["amount"]
    net[(l["acct"], mo)] -= l["amount"]
for d, _, _, acct, dr, cr in ADJUSTMENTS:
    net[(acct, d.year * 100 + d.month)] += dr - cr
pl = {}
for no, label, _ in MONTHS:
    s = lambda t: round(sum(-x for (a, mo), x in net.items() if mo == no and types[a] == t), 2)
    rev, cogs, exp = s("Revenue"), -s("COGS"), -s("Expense")
    pl[no] = dict(revenue=rev, cogs=cogs, gross_profit=round(rev - cogs, 2), expenses=exp, net_profit=round(rev - cogs - exp, 2))
cash = {no: round(opening[1000] + opening[1010] + sum(x for (a, mo), x in net.items() if a in (1000, 1010) and mo <= no), 2)
        for no, _, _ in MONTHS}
expected = dict(
    rows=dict(checking=sum(l["reg"] == 1000 for l in lines), savings=sum(l["reg"] == 1010 for l in lines),
              card=sum(l["reg"] == 2000 for l in lines)),
    accounts_by_month={f"{a}-{mo}": round(x, 2) for (a, mo), x in sorted(net.items()) if round(x, 2)},
    pl=pl, cash=cash, statements=statements,
    uncategorized=[dict(desc=l["desc"], amount=l["amount"]) for l in lines if l["acct"] == UNCATEGORIZED],
    in_transit_by_month={no: round(sum(x for (a, mo), x in net.items() if a == CLEARING and mo <= no), 2) for no, _, _ in MONTHS},
    true_account=[dict(reg=l["reg"], post=l["post"].isoformat(), desc=l["desc"], amount=l["amount"], acct=l["acct"]) for l in lines],
)
(ROOT / "test").mkdir(exist_ok=True)
(ROOT / "test" / "expected.json").write_text(json.dumps(expected, indent=1), encoding="utf-8")
print(expected["rows"], "lines")
for no, _, _ in MONTHS:
    print(no, pl[no], "cash", cash[no])
print("in transit", expected["in_transit_by_month"])
