"""Generates the fictional October 2026 data for "Lakeside Hardware Ltd" (made-up company).

Writes two deliberately messy CSV exports into ../data/ and the designed answer into ../test/expected.json:
  bank-statement-2026-10.csv  - US bank export: MM/DD/YYYY dates, amounts as text with thousand separators,
                                references typed inconsistently, 2 rows exported twice
  cash-book-2026-10.csv       - accounting export: ISO dates, signed amounts, one receipt posted twice
Every exception case is planted on purpose so the workbook can be checked against a known answer.
"""
import csv, json, random
from datetime import date, timedelta
from pathlib import Path

random.seed(20261031)
ROOT = Path(__file__).resolve().parent.parent
OPENING = 18450.00
MONTH_START = date(2026, 10, 1)

CUSTOMERS = ["Brightwater Builders", "Oakridge Property Mgmt", "Henley Joinery", "Marsh & Co Plumbing",
             "Northgate Schools Trust", "Riverside Homes", "Coastal Roofing", "Pine Street Cafe",
             "Granite Kitchens", "Willow Landscapes", "Summit Electrical", "Harbour Rentals"]
SUPPLIERS = ["Atlas Fasteners", "BlueLine Timber", "Crown Paints", "Delta Tools Wholesale",
             "Evergreen Garden Supply", "Fortis Steel", "Metro Utilities", "City Council Rates"]


def messy(ref):
    """How a bank clerk / customer types a reference: case, spaces and dashes vary."""
    head, num = ref.split("-") if "-" in ref else (ref[:3], ref[3:])
    return random.choice([f"{head}-{num}", f"{head}{num}", f"{head.lower()}-{num}", f" {head} {num}",
                          f"{head}-{num} ", f"{head.lower()}{num}"])


def day(n):
    return MONTH_START + timedelta(days=n)


bank, ledger, expected = [], [], []
inv, po, chq = 1001, 5001, 4501

# 1) Normal transactions that match (some clear 1-3 days after the ledger date)
for _ in range(150):
    d = random.randint(0, 27)
    gap = random.choice([0] * 8 + [1, 2, 3])
    if random.random() < 0.55:
        ref = f"INV-{inv}"; inv += 1
        party = random.choice(CUSTOMERS)
        amt = round(random.uniform(120, 4800), 2)
        ledger.append(dict(date=day(d), ref=ref, party=party, amount=amt, memo="Customer receipt"))
        bank.append(dict(date=day(d + gap), ref=messy(ref), desc=f"TRANSFER FROM {party.upper()}", amount=amt))
    else:
        if random.random() < 0.3:
            ref = f"CHQ-00{chq}"; chq += 1; memo = "Cheque payment"; desc = f"CHEQUE {ref[4:]}"
        else:
            ref = f"PO-{po}"; po += 1; memo = "Supplier payment"; desc = None
        party = random.choice(SUPPLIERS)
        amt = -round(random.uniform(80, 3900), 2)
        ledger.append(dict(date=day(d), ref=ref, party=party, amount=amt, memo=memo))
        bank.append(dict(date=day(d + gap), ref=messy(ref), desc=desc or f"PAYMENT TO {party.upper()}", amount=amt))
    expected.append(dict(key=ref.replace("-", ""), status="Matched", gap=gap))

# 2) Same reference, different amount
for ref, party, led, bnk, why in [
        (f"INV-{inv}", "Henley Joinery", 1000.00, 980.00, "customer deducted a 20.00 transfer fee"),
        (f"PO-{po}", "Fortis Steel", -1452.00, -1425.00, "ledger typo: digits swapped"),
        (f"INV-{inv + 1}", "Pine Street Cafe", 1557.50, 1575.50, "customer overpaid 18.00")]:
    d = random.randint(3, 25)
    ledger.append(dict(date=day(d), ref=ref, party=party, amount=led, memo="Customer receipt" if led > 0 else "Supplier payment"))
    bank.append(dict(date=day(d + 1), ref=messy(ref),
                     desc=(f"TRANSFER FROM {party.upper()}" if bnk > 0 else f"PAYMENT TO {party.upper()}"), amount=bnk))
    expected.append(dict(key=ref.replace("-", ""), status="Amount differs", diff=round(bnk - led, 2), why=why))
inv += 2; po += 1

# 3) Only in the bank: charges, interest, a direct debit nobody booked, an unidentified cash deposit
for d, ref, desc, amt in [(30, "", "ACCOUNT MAINTENANCE FEE", -25.00), (30, "", "MERCHANT CARD FEE OCT", -38.40),
                          (30, "", "CREDIT INTEREST", 4.12), (14, "DD-7781", "DIRECT DEBIT SHIELD INSURANCE", -189.00),
                          (19, "", "CASH DEPOSIT BRANCH 12", 350.00)]:
    bank.append(dict(date=day(d), ref=ref, desc=desc, amount=amt))
    expected.append(dict(key=ref.replace("-", "") or None, status="Only in bank", amount=amt, desc=desc))

# 4) Only in the cash book: deposit in transit and payments that have not cleared yet
for d, ref, party, amt, memo in [(30, f"INV-{inv}", "Northgate Schools Trust", 2140.00, "Customer receipt"),
                                 (28, f"CHQ-00{chq}", "BlueLine Timber", -860.00, "Cheque payment"),
                                 (29, f"CHQ-00{chq + 1}", "Crown Paints", -1215.50, "Cheque payment"),
                                 (30, f"PO-{po}", "Metro Utilities", -400.00, "Supplier payment")]:
    ledger.append(dict(date=day(d), ref=ref, party=party, amount=amt, memo=memo))
    expected.append(dict(key=ref.replace("-", ""), status="Only in ledger", amount=amt))

# 5) A receipt posted twice in the cash book (bank has it once)
dup = random.choice([l for l in ledger if l["ref"].startswith("INV") and l["amount"] > 0][:40])
ledger.append(dict(dup, memo="Customer receipt"))
expected.append(dict(key=dup["ref"].replace("-", ""), status="Duplicate in ledger", amount=dup["amount"]))

# 6) Running balance, then the export glitch: two rows repeated verbatim (same balance too)
bank.sort(key=lambda r: r["date"])
bal = OPENING
for r in bank:
    bal = round(bal + r["amount"], 2)
    r["balance"] = bal
for r in random.sample([b for b in bank[10:120] if b["ref"]], 2):
    bank.insert(bank.index(r) + 1, dict(r))

ledger.sort(key=lambda r: r["date"])

# Write files
(ROOT / "data").mkdir(exist_ok=True); (ROOT / "test").mkdir(exist_ok=True)
money = lambda v: f"{v:,.2f}"
with open(ROOT / "data" / "bank-statement-2026-10.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
    w.writerow(["Date", "Reference", "Description", "Debit", "Credit", "Balance"])
    for r in bank:
        w.writerow([r["date"].strftime("%m/%d/%Y"), r["ref"], r["desc"],
                    money(-r["amount"]) if r["amount"] < 0 else "", money(r["amount"]) if r["amount"] > 0 else "",
                    money(r["balance"])])
with open(ROOT / "data" / "cash-book-2026-10.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Date", "Doc No", "Party", "Reference", "Amount", "Memo"])
    for i, r in enumerate(ledger, 1):
        w.writerow([r["date"].isoformat(), f"CB-{i:04d}", r["party"], r["ref"], f"{r['amount']:.2f}", r["memo"]])

json.dump(dict(opening=OPENING, bank_closing=bank[-1]["balance"], bank_rows_in_file=len(bank), bank_duplicates_in_file=2,
               ledger_rows=len(ledger), cases=expected),
          open(ROOT / "test" / "expected.json", "w", encoding="utf-8"), indent=1, default=str)
print(f"bank rows {len(bank)} (2 repeated), ledger rows {len(ledger)}, cases {len(expected)}")
