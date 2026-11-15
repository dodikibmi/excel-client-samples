"""The bookkeeping done again in plain Python, straight from the four CSV files in data/ (no Excel involved).

verify.py compares this with the planted answer (expected.json) and with every number in the workbook.
Run on its own to print the headline numbers:  python test/model.py
"""
import calendar, csv
from collections import defaultdict
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_EVEN
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from coa import ACCOUNTS, RULES, ADJUSTMENTS, MONTHS, CLEARING, UNCATEGORIZED, BUDGET, CASH_PLAN, CASH_BUFFER  # noqa: E402

DATA = ROOT / "data"
TYPES = {a[0]: a[2] for a in ACCOUNTS}
NAMES = {a[0]: a[1] for a in ACCOUNTS}
OPENING = {a[0]: round(a[3] - a[4], 2) for a in ACCOUNTS}
MONTH_NOS = [no for no, _, _ in MONTHS]
LABELS = [lbl for _, lbl, _ in MONTHS]
ENDS = [end for _, _, end in MONTHS]
CASH = (1000, 1010)
PL_TYPES = ("Revenue", "COGS", "Expense")
COST_ACCTS = [a for a, *_ in ACCOUNTS if TYPES[a] in ("COGS", "Expense")]
UNKNOWN_PAYEE = "Unknown (on Review)"


def money(t):
    t = (t or "").strip().replace("$", "").replace(",", "").replace(" ", "")
    return float(t) if t else 0.0


def norm(t):
    t = "".join(" " if ch in "*.-#/,&'" else ch for ch in t.upper())
    return " ".join(t.split())


def us(d):
    m, dd, y = d.strip().split("/")
    return datetime(int(y), int(m), int(dd)).date()


def r2(x):
    return round(x + 0.0, 2)


def pq_round(x):
    """Power Query's Number.Round rounds half to even: -6530.185 -> -6530.18."""
    return float(Decimal(f"{x:.6f}").quantize(Decimal("0.01"), ROUND_HALF_EVEN))


# ---------------------------------------------------------------- bank lines
lines = []
with open(DATA / "checking-2026-08-10.csv", encoding="utf-8") as f:
    body = f.read().splitlines()[3:]
for i, r in enumerate(csv.DictReader(body), 1):
    lines.append(dict(line=f"CHK-{i:03d}", reg=1000, post=us(r["Date"]), desc=r["Description"].strip(),
                      amount=r2(money(r["Credit"]) - money(r["Debit"]))))
with open(DATA / "savings-2026-08-10.csv", encoding="utf-8") as f:
    for i, r in enumerate(csv.DictReader(f), 1):
        lines.append(dict(line=f"SAV-{i:03d}", reg=1010, post=datetime.strptime(r["Posted"].strip(), "%Y-%m-%d").date(),
                          desc=r["Details"].strip(), amount=money(r["Amount"])))
with open(DATA / "visa-2026-08-10.csv", encoding="utf-8") as f:
    for i, r in enumerate(csv.DictReader(f), 1):
        lines.append(dict(line=f"VISA-{i:03d}", reg=2000, post=us(r["Post Date"]), desc=r["Description"].strip(), amount=money(r["Amount"])))
RULES_N = [(norm(k), a, p) for k, a, _, p in RULES]
for l in lines:
    hit = next(((a, p) for k, a, p in RULES_N if k in norm(l["desc"])), None)
    l["acct"], l["payee"] = hit if hit else (UNCATEGORIZED, UNKNOWN_PAYEE)
    l["mo"] = l["post"].year * 100 + l["post"].month

# ---------------------------------------------------------------- transfers
outs = sorted([l for l in lines if l["acct"] == CLEARING and l["amount"] < 0], key=lambda l: (l["post"], l["line"]))
ins = sorted([l for l in lines if l["acct"] == CLEARING and l["amount"] > 0], key=lambda l: (l["post"], l["line"]))
pairs, unpaired = [], []
for o in outs:
    m = next((i for i in ins if abs(i["amount"] + o["amount"]) < 0.005 and abs((i["post"] - o["post"]).days) <= 5), None)
    if m:
        ins.remove(m)
        pairs.append((o, m))
    else:
        unpaired.append(o)
unpaired += ins

# ---------------------------------------------------------------- ledger movements
net = defaultdict(float)
for l in lines:
    net[(l["reg"], l["mo"])] += l["amount"]
    net[(l["acct"], l["mo"])] -= l["amount"]
for d, _, _, a, dr, cr in ADJUSTMENTS:
    net[(a, d.year * 100 + d.month)] += dr - cr


def move(a, no):
    return r2(net.get((a, no), 0.0))


def balance(a, no):
    return r2(OPENING[a] + sum(v for (aa, mo), v in net.items() if aa == a and mo <= no))


def actual(a, no):
    """P&L amount as the report shows it: revenue and costs both positive."""
    return r2(-move(a, no) if TYPES[a] == "Revenue" else move(a, no))


pl, tb = {}, {}
for no in MONTH_NOS:
    s = lambda t: r2(sum(-v for (a, mo), v in net.items() if mo == no and TYPES[a] == t))
    rev, cogs, ex = s("Revenue"), -s("COGS"), -s("Expense")
    pl[no] = dict(revenue=rev, cogs=cogs, gross_profit=r2(rev - cogs), expenses=ex, net_profit=r2(rev - cogs - ex))
    bals = [balance(a, no) for a in OPENING]
    tb[no] = (r2(sum(b for b in bals if b > 0)), r2(-sum(b for b in bals if b < 0)))
budget_pl = {}
for k, no in enumerate(MONTH_NOS):
    b = lambda t: r2(sum(v[k] for a, v in BUDGET.items() if TYPES[a] == t))
    rev, cogs, ex = b("Revenue"), b("COGS"), b("Expense")
    budget_pl[no] = dict(revenue=rev, cogs=cogs, gross_profit=r2(rev - cogs), expenses=ex, net_profit=r2(rev - cogs - ex))
cash = {no: r2(balance(1000, no) + balance(1010, no)) for no in MONTH_NOS}
cash_open = r2(OPENING[1000] + OPENING[1010])
in_transit = {no: balance(CLEARING, no) for no in MONTH_NOS}

stmts = {}
with open(DATA / "statements-2026-08-10.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        a = 1000 if "CHECKING" in r["Account"].upper() else 1010 if "SAVINGS" in r["Account"].upper() else 2000
        d = us(r["Statement date"])
        stmts[(a, d.year * 100 + d.month)] = money(r["Ending balance"])

# ---------------------------------------------------------------- cash flow statement (indirect method)


def cash_flow(nos):
    m = lambda a: r2(sum(move(a, no) for no in nos))
    np_ = r2(sum(pl[no]["net_profit"] for no in nos))
    cf = dict(net_profit=np_, depreciation=-m(1510), transit=-m(CLEARING), card=-m(2000))
    cf["operating"] = r2(cf["net_profit"] + cf["depreciation"] + cf["transit"] + cf["card"])
    cf["equipment"] = -m(1500)
    cf["investing"] = cf["equipment"]
    cf["draws"], cf["capital"] = -m(3100), -m(3000)
    cf["financing"] = r2(cf["draws"] + cf["capital"])
    cf["net_change"] = r2(cf["operating"] + cf["investing"] + cf["financing"])
    first = MONTH_NOS.index(nos[0])
    cf["start"] = cash_open if first == 0 else cash[MONTH_NOS[first - 1]]
    cf["end"] = r2(cf["start"] + cf["net_change"])
    cf["cash_moves"] = r2(sum(m(a) for a in CASH))
    return cf


cash_flows = {no: cash_flow([no]) for no in MONTH_NOS}
cash_flows["quarter"] = cash_flow(MONTH_NOS)

# ---------------------------------------------------------------- 13-week cash forecast (same rules as E9)


def add_months(d, k):
    y, m = divmod(d.month - 1 + k, 12)
    y, m = d.year + y, m + 1
    return d.replace(year=y, month=m, day=min(d.day, calendar.monthrange(y, m)[1]))


def month_end(d):
    return d.replace(day=calendar.monthrange(d.year, d.month)[1])


last_date = max(l["post"] for l in lines)
fc_start, fc_end = last_date + timedelta(days=1), last_date + timedelta(days=91)
fc_in, fc_out = defaultdict(float), defaultdict(float)
plan = []
for order, item, cfl, reg, accts, basis, lb, freq, first, amount, _note in CASH_PLAN:
    accts = {int(a) for a in accts.split(";")}
    hist = [l for l in lines if l["reg"] == reg and l["acct"] in accts]
    if basis == "fixed":
        per = float(amount)
    elif basis == "avg_week":
        per = sum(l["amount"] for l in hist if l["post"] > last_date - timedelta(days=7 * lb)) / lb
    elif basis == "avg_month":
        cut = month_end(add_months(last_date, -lb))
        per = sum(l["amount"] for l in hist if l["post"] > cut) / lb
    else:
        by_day = defaultdict(float)
        for l in hist:
            by_day[l["post"]] += l["amount"]
        top = [by_day[k] for k in sorted(by_day, reverse=True)[:lb]]
        per = sum(top) / len(top)
    per = pq_round(per)
    if freq == "weekly":
        dates = [fc_start + timedelta(days=7 * k) for k in range(13)]
    elif freq == "biweekly":
        dates = [first + timedelta(days=14 * k) for k in range(8)]
    else:
        dates = [add_months(first, k) for k in range(4)]
    dates = [d for d in dates if fc_start <= d <= fc_end]
    plan.append(dict(item=item, line=cfl, per=per, times=len(dates), total=r2(per * len(dates))))
    for d in dates:
        w = (d - fc_start).days // 7 + 1
        (fc_in if per > 0 else fc_out)[w] += per
weeks = []
run = cash[MONTH_NOS[-1]]
for w in range(1, 14):
    n = r2(fc_in[w] + fc_out[w])
    run = r2(run + n)
    weeks.append(dict(week=w, end=fc_start + timedelta(days=7 * w - 1), money_in=r2(fc_in[w]), money_out=r2(fc_out[w]), net=n, cash=run))
cash_today = cash[MONTH_NOS[-1]]
lowest = min(weeks, key=lambda x: (x["cash"], x["week"]))
weekly_out = r2(-sum(w["money_out"] for w in weeks) / 13)
forecast = dict(today=cash_today, lowest=lowest["cash"], lowest_week=lowest["week"], lowest_end=lowest["end"],
                headroom=r2(lowest["cash"] - CASH_BUFFER), below=sum(w["cash"] < CASH_BUFFER for w in weeks),
                weekly_out=weekly_out, weeks_of_cash=round(cash_today / weekly_out, 1),
                end=weeks[-1]["cash"], change=r2(weeks[-1]["cash"] - cash_today))
burn = r2((cash_open - cash_today) / len(MONTH_NOS))  # average monthly fall in cash, Aug-Oct
forecast["burn_month"] = burn
forecast["runway_months"] = round(cash_today / burn, 1) if burn > 0 else None

# ---------------------------------------------------------------- payees (who gets paid)
payees = defaultdict(lambda: dict(n=0, by=defaultdict(float), accts=set()))
for l in lines:
    if TYPES[l["acct"]] in ("COGS", "Expense"):
        p = payees[l["payee"]]
        p["n"] += 1
        p["by"][l["mo"]] += -l["amount"]
        p["accts"].add(l["acct"])
spend_total = r2(sum(sum(p["by"].values()) for p in payees.values()))
payee_rows = []
for name, p in payees.items():
    tot = r2(sum(p["by"].values()))
    a = min(p["accts"])
    payee_rows.append(dict(payee=name, account=f"{a} {NAMES[a]}", n=p["n"], months=[r2(p["by"][no]) for no in MONTH_NOS],
                           total=tot, share=tot / spend_total))
payee_rows.sort(key=lambda r: (-r["total"], r["payee"]))

# ---------------------------------------------------------------- dashboard (per month, per comparison)


def comp_value(a, k, mode):
    """What account a is compared with in month index k: its budget, or last month's actual (None for the first month)."""
    if mode == "budget":
        return r2(BUDGET.get(a, (0, 0, 0))[k])
    return actual(a, MONTH_NOS[k - 1]) if k > 0 else None


def comp_pl(k, mode):
    if mode == "budget":
        return budget_pl[MONTH_NOS[k]]
    return pl[MONTH_NOS[k - 1]] if k > 0 else None


def dashboard(k, mode):
    no = MONTH_NOS[k]
    p, c = pl[no], comp_pl(k, mode)
    d = dict(month=LABELS[k], revenue=p["revenue"], net_profit=p["net_profit"],
             gross_margin=p["gross_profit"] / p["revenue"], opex_pct=p["expenses"] / p["revenue"],
             cash=cash[no], cash_change=r2(cash[no] - (cash_open if k == 0 else cash[MONTH_NOS[k - 1]])),
             in_transit=in_transit[no],
             review_lines=sum(1 for l in lines if l["mo"] == no and l["acct"] == UNCATEGORIZED)
             + sum(1 for u in unpaired if u["mo"] == no),
             weeks_of_cash=forecast["weeks_of_cash"])
    if c:
        d.update(c_revenue=c["revenue"], c_net_profit=c["net_profit"], c_gross_margin=c["gross_profit"] / c["revenue"],
                 c_opex_pct=c["expenses"] / c["revenue"])
    else:
        d.update(c_revenue=None, c_net_profit=None, c_gross_margin=None, c_opex_pct=None)
    d["waterfall"] = [p["revenue"], -p["cogs"], -p["expenses"], p["net_profit"]]
    rows = []
    for a in COST_ACCTS:
        cv = comp_value(a, k, mode)
        rows.append(dict(acct=a, value=actual(a, no), comp=cv, var=None if cv is None else r2(cv - actual(a, no)),
                         trend=[actual(a, m) for m in MONTH_NOS]))
    d["costs"] = sorted(rows, key=lambda r: (-r["value"], r["acct"]))
    d["payees"] = sorted(([r["payee"], r["months"][k]] for r in payee_rows), key=lambda x: (-x[1], payee_rows.index(
        next(r for r in payee_rows if r["payee"] == x[0]))))[:6]
    # biggest overrun / saving across every P&L account (+ = better)
    var = {}
    for a, *_ in ACCOUNTS:
        if TYPES[a] in PL_TYPES:
            cv = comp_value(a, k, mode)
            if cv is not None:
                var[a] = r2(actual(a, no) - cv) if TYPES[a] == "Revenue" else r2(cv - actual(a, no))
    if var:
        d["worst"] = min(var, key=var.get)
        d["best"] = max(var, key=var.get)
        d["worst_var"], d["best_var"] = var[d["worst"]], var[d["best"]]
    return d


# ---------------------------------------------------------------- month-end close checklist


def close_checklist(k):
    no, end = MONTH_NOS[k], ENDS[k]
    st = {}
    st["exports"] = all(any(l["reg"] == r and l["mo"] == no for l in lines) for r in (1000, 1010, 2000))
    for a in (1000, 1010, 2000):
        b = -balance(a, no) if a == 2000 else balance(a, no)
        st[f"rec{a}"] = (a, no) in stmts and r2(b - stmts[(a, no)]) == 0
    open_halves = r2(sum(o["amount"] * -1 for o, i in pairs if o["post"] <= end < i["post"])
                     - sum(i["amount"] for o, i in pairs if i["post"] <= end < o["post"])
                     + sum(-u["amount"] for u in unpaired if u["post"] <= end))
    st["transfers"] = r2(open_halves - balance(CLEARING, no)) == 0
    adj = [x for x in ADJUSTMENTS if x[0].year * 100 + x[0].month == no]
    st["adjustments"] = bool(adj) and r2(sum(x[4] for x in adj) - sum(x[5] for x in adj)) == 0
    st["tb"] = tb[no][0] == tb[no][1]
    st["bs"] = True  # recomputed in verify.py from the workbook's own balance sheet totals
    st["cashflow"] = cash_flows[no]["end"] == cash[no]
    unk = [l for l in lines if l["mo"] == no and l["acct"] == UNCATEGORIZED]
    st["review"] = (len(unk), r2(sum(-l["amount"] for l in unk)))
    return st


if __name__ == "__main__":
    for no in MONTH_NOS:
        print(no, pl[no], "cash", cash[no], "transit", in_transit[no], "cf", cash_flows[no])
    print("quarter cf", cash_flows["quarter"])
    print("forecast", forecast)
    for w in weeks:
        print("  W", w)
    for p in plan:
        print("  plan", p)
    print("payees", len(payee_rows), "spend", spend_total)
    for r in payee_rows:
        print("  ", r["payee"], r["n"], r["months"], r["total"], f"{r['share']:.1%}")
    for k in range(3):
        for mode in ("budget", "last"):
            d = dashboard(k, mode)
            print(LABELS[k], mode, {x: d.get(x) for x in ("revenue", "net_profit", "gross_margin", "opex_pct", "cash", "cash_change",
                                                            "in_transit", "review_lines", "c_net_profit", "worst", "worst_var", "best", "best_var")})
            print("   costs", [(r["acct"], r["value"], r["var"]) for r in d["costs"]][:5], "payees", d["payees"][:3])
        print("close", close_checklist(k))
