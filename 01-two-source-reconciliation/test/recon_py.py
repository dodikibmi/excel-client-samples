"""The whole reconciliation again in plain Python, straight from the CSV files in ../data (no Excel involved).
Used by verify.py. Same rules as the Power Query steps, written independently:
  rule 1 "Reference"          same account + same reference key, same amount
  rule 2 "Reference ± amount" same reference key, amount within the amount tolerance
  rule 3 "Amount + date"      leftovers only: same account, amount within tolerance, dates within the days tolerance,
                              closest first (amount, then days, then line id), each line used once
"""
import calendar, csv
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def key(ref):
    return "".join(ch for ch in ref.upper() if ch not in " -/.") or None


def num(v):
    return float(v.replace(",", "")) if v.strip() else 0.0


def load(folder=ROOT / "data"):
    bank, ledger, raw_bank, raw_ledger, files = [], [], 0, 0, []
    seen = set()
    for f in sorted(folder.glob("*.csv"), key=lambda p: p.name.lower()):
        kind, acct, period = f.stem.split("_")
        rows = list(csv.DictReader(open(f, encoding="utf-8")))
        files.append((kind, acct, period, len(rows)))
        code = acct[:3].upper() + ("-B-" if kind == "bank" else "-L-") + period.replace("-", "")
        for i, r in enumerate(rows, 1):
            if kind == "bank":
                raw_bank += 1
                t = (acct,) + tuple(r.values())
                if t in seen:
                    continue
                seen.add(t)
                bank.append(dict(account=acct, period=period, id=f"{code}-{i:03d}", ref=r["Reference"].strip(),
                                 date=datetime.strptime(r["Date"].strip(), "%m/%d/%Y").date(), desc=r["Description"].strip(),
                                 amount=round(num(r["Credit"]) - num(r["Debit"]), 2), balance=num(r["Balance"]),
                                 key=key(r["Reference"])))
            else:
                raw_ledger += 1
                ledger.append(dict(account=acct, period=period, id=f"{code}-{i:03d}", ref=r["Reference"].strip(),
                                   date=date.fromisoformat(r["Date"].strip()), doc=r["Doc No"], party=r["Party"],
                                   amount=round(float(r["Amount"]), 2), key=key(r["Reference"])))
    return dict(bank=bank, ledger=ledger, raw_bank=raw_bank, raw_ledger=raw_ledger, files=files)


def reconcile(data, amount_tol=1.00, days_tol=5):
    bank, ledger = data["bank"], data["ledger"]
    rows = []  # dict(account, status, rule, b, l)

    def split(lines):
        first, rep, cnt = {}, [], Counter()
        for x in sorted(lines, key=lambda x: x["id"]):
            if x["key"] is None:
                continue
            cnt[(x["account"], x["key"])] += 1
            if cnt[(x["account"], x["key"])] > 1:
                rep.append(x)
            else:
                first[(x["account"], x["key"])] = x
        nokey = [x for x in lines if x["key"] is None]
        return first, rep, nokey

    bfirst, brep, bnokey = split(bank)
    lfirst, lrep, lnokey = split(ledger)
    for x in brep:
        rows.append(dict(account=x["account"], status="Duplicate in bank", rule="", b=x, l=None))
    for x in lrep:
        rows.append(dict(account=x["account"], status="Duplicate in ledger", rule="", b=None, l=x))
    bleft, lleft = list(bnokey), list(lnokey)
    for k, b in bfirst.items():
        l = lfirst.get(k)
        if l is None:
            bleft.append(b); continue
        d = round(b["amount"] - l["amount"], 2)
        if d == 0:
            rows.append(dict(account=b["account"], status="Matched", rule="Reference", b=b, l=l))
        elif abs(d) <= amount_tol + 1e-9:
            rows.append(dict(account=b["account"], status="Matched", rule="Reference ± amount", b=b, l=l))
        else:
            rows.append(dict(account=b["account"], status="Amount differs", rule="", b=b, l=l))
    lleft += [l for k, l in lfirst.items() if k not in bfirst]
    cands = []
    for b in bleft:
        for l in lleft:
            if b["account"] != l["account"]:
                continue
            ad, dd = abs(round(b["amount"] - l["amount"], 2)), abs((b["date"] - l["date"]).days)
            if ad <= amount_tol + 1e-9 and dd <= days_tol:
                cands.append((ad, dd, b["id"], l["id"], b, l))
    cands.sort(key=lambda c: c[:4])
    ub, ul = set(), set()
    for ad, dd, bid, lid, b, l in cands:
        if bid in ub or lid in ul:
            continue
        ub.add(bid); ul.add(lid)
        rows.append(dict(account=b["account"], status="Matched", rule="Amount + date", b=b, l=l))
    for b in bleft:
        if b["id"] not in ub:
            rows.append(dict(account=b["account"], status="Only in bank", rule="", b=b, l=None))
    for l in lleft:
        if l["id"] not in ul:
            rows.append(dict(account=l["account"], status="Only in ledger", rule="", b=None, l=l))
    for r in rows:
        b, l = r["b"], r["l"]
        r["exception_id"] = (b or l)["id"]
        r["date"] = (b or l)["date"]
        r["month"] = r["date"].replace(day=1)
        r["bank_amount"] = b["amount"] if b else None
        r["ledger_amount"] = l["amount"] if l else None
        r["value"] = abs(b["amount"] if b else l["amount"])
        r["difference"] = round(b["amount"] - l["amount"], 2) if b and l and round(b["amount"] - l["amount"], 2) != 0 else None
        r["wb_amount"] = r["difference"] if r["status"] == "Amount differs" else (b or l)["amount"]
    return rows


def month_end(s):
    return s.replace(day=calendar.monthrange(s.year, s.month)[1])


def statement(rows, account, s, opening_bank, opening_book):
    """Bank reconciliation statement for one account, month starting s (cumulative from the first file)."""
    e = month_end(s)
    R = [r for r in rows if r["account"] == account]
    bd = lambda r: r["b"]["date"] if r["b"] else None
    ld = lambda r: r["l"]["date"] if r["l"] else None
    o = {}
    o["bank_open"] = opening_bank + sum(r["bank_amount"] for r in R if bd(r) and bd(r) < s)
    o["bank_moves"] = sum(r["bank_amount"] for r in R if bd(r) and s <= bd(r) <= e)
    o["bank_close"] = o["bank_open"] + o["bank_moves"]
    real = [r for r in R if ld(r) and ld(r) <= e and r["status"] != "Duplicate in ledger" and not (bd(r) and bd(r) <= e)]
    o["in_transit"] = sum(r["ledger_amount"] for r in real if r["ledger_amount"] > 0)
    o["outstanding"] = sum(r["ledger_amount"] for r in real if r["ledger_amount"] < 0)
    o["adj_bank"] = o["bank_close"] + o["in_transit"] + o["outstanding"]
    o["book_open"] = opening_book + sum(r["ledger_amount"] for r in R if ld(r) and ld(r) < s)
    o["book_moves"] = sum(r["ledger_amount"] for r in R if ld(r) and s <= ld(r) <= e)
    o["book_close"] = o["book_open"] + o["book_moves"]
    o["bank_items"] = sum(r["bank_amount"] for r in R if bd(r) and bd(r) <= e and not (ld(r) and ld(r) <= e))
    o["differences"] = sum(r["bank_amount"] - r["ledger_amount"] for r in R if bd(r) and ld(r) and bd(r) <= e and ld(r) <= e)
    o["twice"] = -sum(r["ledger_amount"] for r in R if r["status"] == "Duplicate in ledger" and ld(r) <= e)
    o["adj_book"] = o["book_close"] + o["bank_items"] + o["differences"] + o["twice"]
    o["unexplained"] = round(o["adj_bank"] - o["adj_book"], 2)
    return {k: round(v, 2) for k, v in o.items()}


def month_kpis(rows, account, s, notes):
    """KPIs, donut, status bars, rules for one account and month; notes = {exception_id: dict(resolved=...)}."""
    e = month_end(s)
    R = [r for r in rows if r["account"] == account]
    bank_m = [r for r in R if r["b"] and s <= r["b"]["date"] <= e]
    in_m = [r for r in R if r["month"] == s]
    k = dict(bank_lines=len(bank_m), matched_lines=sum(r["status"] == "Matched" for r in bank_m))
    k["match_rate"] = k["matched_lines"] / k["bank_lines"] if k["bank_lines"] else 0
    k["matched_amount"] = round(sum(r["value"] for r in bank_m if r["status"] == "Matched"), 2)
    exc = [r for r in R if r["status"] != "Matched"]
    open_ = [r for r in exc if r["date"] <= e and not (notes.get(r["exception_id"]) or {}).get("resolved")]
    k["open_exceptions"] = len(open_)
    k["resolved"] = sum(1 for r in exc if r["date"] <= e and (notes.get(r["exception_id"]) or {}).get("resolved"))
    k["exceptions_in_month"] = sum(1 for r in in_m if r["status"] != "Matched")
    k["by_rule"] = Counter(r["rule"] for r in in_m if r["status"] == "Matched")
    k["by_status_value"] = defaultdict(float)
    k["by_status_lines"] = Counter()
    for r in in_m:
        k["by_status_value"][r["status"]] += r["value"]
        k["by_status_lines"][r["status"]] += 1
    aging = Counter(); aging_amt = defaultdict(float)
    for r in open_:
        age = (e - r["date"]).days
        b = "0-7" if age <= 7 else ("8-30" if age <= 30 else "30+")
        aging[(r["status"], b)] += 1
        aging_amt[b] += abs(r["wb_amount"])
    k["aging"], k["aging_amount"] = aging, aging_amt
    return k


def trend(rows, account, months, notes):
    """Per month: auto-match % of bank lines, and the open (unresolved) exception amount dated in that month."""
    out = []
    for s in months:
        e = month_end(s)
        R = [r for r in rows if r["account"] == account]
        bank_m = [r for r in R if r["b"] and s <= r["b"]["date"] <= e]
        rate = sum(r["status"] == "Matched" for r in bank_m) / len(bank_m) if bank_m else 0
        rem = sum(abs(r["wb_amount"]) for r in R if r["status"] != "Matched" and r["month"] == s
                  and not (notes.get(r["exception_id"]) or {}).get("resolved"))
        out.append((s, rate, round(rem, 2)))
    return out
