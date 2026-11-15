"""Independent check of the E9 Power BI project.

1. Python does the bookkeeping again straight from the CSV files in data/ (no Power BI involved): parse the three
   exports, clean the descriptions, apply rules.csv, post double entry + adjustments.csv, P&L per month, budget,
   variance, month-on-month, cash, money in transit, books vs statements, cash flow lines.
2. Compares that with the answer planted when the data was generated (test/expected.json).
3. Reads the project files: the TMDL model must have the tables, star-schema relationships, measures and DAX
   functions listed in expected.json; every field used by a visual in the PBIR report must exist in the model.
4. Compares Python with what the model in Power BI Desktop answered (test/model-snapshot.json, written by
   test/query_model.ps1 while the report was open): every month, account, statement, cash flow line, check and
   the account of all 186 bank lines.
5. Copperleaf-Finance.pbix is a Power BI file with the data model inside.
Exit code 1 if anything differs.   Run:  python test/verify.py
"""
import csv, json, re, sys, zipfile
from decimal import Decimal, ROUND_HALF_EVEN
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
SM = ROOT / "Copperleaf-Finance.SemanticModel" / "definition"
RP = ROOT / "Copperleaf-Finance.Report" / "definition"
results = []


def check(name, ok, detail=""):
    results.append(bool(ok))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


def near(a, b, tol=0.005):
    return a is not None and b is not None and abs(float(a) - float(b)) < tol


def money(t):
    t = (t or "").strip().replace("$", "").replace(",", "").replace(" ", "")
    return float(t) if t else 0.0


def norm(t):
    t = "".join(" " if ch in "*.-#/,&'" else ch for ch in (t or "").upper())
    return " ".join(t.split())


def us(d):
    m, dd, y = d.strip().split("/")
    return date(int(y), int(m), int(dd))


def rows(name):
    with open(DATA / name, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def mo(d):
    return d.year * 100 + d.month


# ====================================================================== 1. plain Python from the CSV files
lines = []
with open(DATA / "checking-2026-08-10.csv", encoding="utf-8") as f:
    body = f.read().splitlines()
start = next(i for i, l in enumerate(body) if l.startswith("Date,"))
for i, r in enumerate(csv.DictReader(body[start:]), 1):
    if r["Date"].strip():
        lines.append(dict(id=f"CHK-{i:03d}", reg=1000, post=us(r["Date"]), desc=r["Description"].strip(),
                          amount=round(money(r["Credit"]) - money(r["Debit"]), 2), balance=money(r["Balance"])))
for i, r in enumerate(rows("savings-2026-08-10.csv"), 1):
    lines.append(dict(id=f"SAV-{i:03d}", reg=1010, post=datetime.strptime(r["Posted"].strip(), "%Y-%m-%d").date(),
                      desc=r["Details"].strip(), amount=money(r["Amount"])))
for i, r in enumerate(rows("visa-2026-08-10.csv"), 1):
    lines.append(dict(id=f"VISA-{i:03d}", reg=2000, post=us(r["Post Date"]), desc=r["Description"].strip(), amount=money(r["Amount"])))

rules = [(norm(r["Keyword"]), int(r["Account No"])) for r in sorted(rows("rules.csv"), key=lambda r: int(r["Priority"]))]
for l in lines:
    l["acct"] = next((a for k, a in rules if k in norm(l["desc"])), 6990)

accounts = {int(r["Account No"]): r for r in rows("accounts.csv")}
types = {a: r["Type"] for a, r in accounts.items()}
opening = {a: float(r["Opening Balance"]) for a, r in accounts.items()}
cash_accts = {a for a, r in accounts.items() if r["Is Cash"] == "Y"}
net = defaultdict(float)  # (account, yyyymm) -> debit-positive movement
for l in lines:
    net[(l["reg"], mo(l["post"]))] += l["amount"]
    net[(l["acct"], mo(l["post"]))] -= l["amount"]
for r in rows("adjustments.csv"):
    net[(int(r["Account No"]), mo(datetime.strptime(r["Date"], "%Y-%m-%d").date()))] += float(r["Debit"]) - float(r["Credit"])
months = sorted({m for _, m in net})


def movement(pred, m):
    return round(sum(v for (a, mm), v in net.items() if mm == m and pred(a)), 2)


def balance(a, m):
    return round(opening[a] + sum(v for (aa, mm), v in net.items() if aa == a and mm <= m), 2) + 0.0  # no -0.00


budget = defaultdict(float)  # (account, yyyymm) -> P&L sign (income +, costs -)
for r in rows("budget.csv"):
    a, m = int(r["Account No"]), int(r["Month"][:4] + r["Month"][5:7])
    budget[(a, m)] += float(r["Budget"]) * (1 if types[a] == "Revenue" else -1)

py = {}
for m in months:
    rev = -movement(lambda a: types[a] == "Revenue", m)
    cogs = movement(lambda a: types[a] == "COGS", m)
    opex = movement(lambda a: types[a] == "Expense", m)
    bud = round(sum(v for (a, mm), v in budget.items() if mm == m), 2)
    py[m] = dict(revenue=rev, cogs=cogs, expenses=opex, gross_profit=round(rev - cogs, 2), net_profit=round(rev - cogs - opex, 2),
                 budget=bud, variance=round(rev - cogs - opex - bud, 2), cash=round(sum(balance(a, m) for a in cash_accts), 2),
                 in_transit=balance(1090, m))
for i, m in enumerate(months):
    py[m]["mom"] = (py[m]["revenue"] - py[months[i - 1]]["revenue"]) / py[months[i - 1]]["revenue"] if i else None

statements = {}
names = {r["Statement Name"]: a for a, r in accounts.items() if r["Statement Name"]}
for r in rows("statements-2026-08-10.csv"):
    a, d = names[r["Account"]], us(r["Statement date"])
    statements[(a, mo(d))] = (-1 if types[a] == "Liability" else 1) * money(r["Ending balance"])


def flow_line(l):
    o = l["acct"]
    if 4000 <= o < 5000:
        return "Receipts from customers"
    if 5000 <= o < 7000:
        return "Payments to suppliers and staff"
    if o == 3100:
        return "Owner's draws"
    d = norm(l["desc"])
    return "Card payments" if "VISA" in d or "PAYMENT THANK YOU" in d else "Transfers between own accounts"


cashflow = defaultdict(float)
for l in lines:
    if l["reg"] in cash_accts:
        cashflow[(flow_line(l), mo(l["post"]))] += l["amount"]

# ====================================================================== 2. against the planted answer
exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
check("Rows read per export = rows planted (141 checking, 6 savings, 39 card)",
      [sum(l["reg"] == r for l in lines) for r in (1000, 1010, 2000)] == [exp["rows"][k] for k in ("checking", "savings", "card")])
truth = defaultdict(list)
for reg, post, desc, amount, acct in exp["true_account"]:
    truth[(reg, post, norm(desc), round(amount, 2))].append(acct)
wrong = [l["id"] for l in lines if (truth[(l["reg"], l["post"].isoformat(), norm(l["desc"]), l["amount"])] or [None]).pop(0) != l["acct"]]
check(f"rules.csv gives all {len(lines)} lines their planted account", not wrong, ", ".join(wrong[:5]))
check("3 lines without a rule = the 3 planted (120.00, 350.00, 64.18)",
      sorted(-l["amount"] for l in lines if l["acct"] == 6990) == sorted(-u["amount"] for u in exp["uncategorized"]))
for m in months:
    e = exp["pl"][str(m)]
    ok = all(near(py[m][k], e[k]) for k in ("revenue", "cogs", "expenses", "gross_profit", "net_profit"))
    check(f"{m}: P&L = planted (income {py[m]['revenue']:,.2f}, net profit {py[m]['net_profit']:,.2f})", ok, str(py[m]))
    check(f"{m}: budget net profit {py[m]['budget']:,.0f}, variance {py[m]['variance']:,.2f} = planted",
          near(py[m]["budget"], e["budget_net_profit"]) and near(py[m]["variance"], e["variance"]))
    check(f"{m}: cash {py[m]['cash']:,.2f} and in transit {py[m]['in_transit']:,.2f} = planted",
          near(py[m]["cash"], exp["cash"][str(m)]) and near(py[m]["in_transit"], exp["in_transit"][str(m)]))
st_ok = all(near(balance(a, m), v) for (a, m), v in statements.items()) and len(statements) == 9
st_planted = all(near(abs(v), abs(exp["statements"][f"{a}-{m}"])) for (a, m), v in statements.items())
check("Books = all 9 statement balances, statements = planted", st_ok and st_planted)
check("Ledger balances: debits = credits", near(sum(net.values()), 0) and near(sum(opening.values()), 0))
check("Checking: books at 31 Oct = the bank's last running balance",
      near(balance(1000, months[-1]), [l for l in lines if l["reg"] == 1000][-1]["balance"]))

# ====================================================================== 3. the project files
tmdl = {p.stem: p.read_text(encoding="utf-8") for p in (SM / "tables").glob("*.tmdl")}
model_txt = "\n".join(tmdl.values())
check(f"Model has the {len(exp['model']['tables'])} tables (fact, dimensions, budget, statements, checks, insights, forecast)", sorted(tmdl) == sorted(exp["model"]["tables"]),
      str(sorted(tmdl)))
rel_txt = (SM / "relationships.tmdl").read_text(encoding="utf-8")
found = [(re.sub(r"'", "", f), re.sub(r"'", "", t)) for f, t in re.findall(r"fromColumn: (.+)\n\ttoColumn: (.+)", rel_txt)]
check(f"{len(exp['model']['relationships'])} relationships: star schema into Date, Account, Category, plus forecast weeks",
      sorted(found) == sorted(tuple(r) for r in exp["model"]["relationships"]), str(found))
measures = set(re.sub(r"^'|'$", "", m) for m in re.findall(r"^\tmeasure (.+?) =", tmdl["_Measures"], re.M))
missing = [m for m in exp["model"]["measures"] if m not in measures]
check(f"{len(measures)} DAX measures incl. income, expense, net profit, budget, variance, MoM, running cash, closing balance",
      not missing, str(missing))
fn_missing = [f for f in exp["model"]["dax_functions"] if f + " (" not in model_txt]
check("DAX uses CALCULATE, KEEPFILTERS and time intelligence (PREVIOUSMONTH, DATESMTD, TOTALYTD)", not fn_missing, str(fn_missing))
check("Date table is marked as a date table (dataCategory: Time, Date is the key)",
      "dataCategory: Time" in tmdl["Date"] and re.search(r"column Date\n\t\tdataType: dateTime\n.*\n\t\tisKey", tmdl["Date"]))
expr = (SM / "expressions.tmdl").read_text(encoding="utf-8")
check("Power Query reads the raw exports through one DataFolder parameter",
      "IsParameterQuery = true" in expr and expr.count('DataFolder & "') + model_txt.count('DataFolder & "') >= 9)

cols = defaultdict(set)
for t, txt in tmdl.items():
    for c in re.findall(r"^\tcolumn (.+)$", txt, re.M):
        cols[t].add(c.strip("'"))
pages = [json.loads((RP / "pages" / p / "page.json").read_text(encoding="utf-8"))["displayName"]
         for p in json.loads((RP / "pages" / "pages.json").read_text(encoding="utf-8"))["pageOrder"]]
check(f"Report has {len(pages)} pages: " + ", ".join(pages), pages == exp["model"]["pages"])
bad, n_fields = [], 0
for vf in RP.glob("pages/*/visuals/*/visual.json"):
    def walk(o):
        global n_fields
        if isinstance(o, dict):
            for kind in ("Column", "Measure"):
                if kind in o and isinstance(o[kind], dict) and "Property" in o[kind]:
                    ent = o[kind]["Expression"]["SourceRef"].get("Entity")
                    prop = o[kind]["Property"]
                    n_fields += 1
                    # slicer pre-selections refer to the table through an alias (Source), not by name
                    known = cols.get(ent, set()) if ent else set().union(*cols.values())
                    ok = prop in measures if kind == "Measure" else prop in known
                    if not ok:
                        bad.append(f"{vf.parent.name}: {ent}.{prop}")
            for v in o.values():
                walk(v)
        elif isinstance(o, list):
            for v in o:
                walk(v)
    walk(json.loads(vf.read_text(encoding="utf-8")))
check(f"Every field used by the {len(list(RP.glob('pages/*/visuals/*/visual.json')))} visuals exists in the model ({n_fields} references)",
      not bad, ", ".join(bad[:5]))

# ====================================================================== 4. what Power BI Desktop answered
snap_file = ROOT / "test" / "model-snapshot.json"
if not snap_file.exists():
    check("test/model-snapshot.json exists (run test/query_model.ps1 with the report open)", False)
else:
    snap = json.loads(snap_file.read_text(encoding="utf-8"))
    sm = {r["Month No"]: r for r in snap["months"]}
    for m in months:
        s, p = sm.get(m, {}), py[m]
        ok = all(near(s.get(k1), p[k2]) for k1, k2 in (("Income", "revenue"), ("CostOfSales", "cogs"), ("OperatingExpenses", "expenses"),
                                                      ("GrossProfit", "gross_profit"), ("NetProfit", "net_profit"), ("Budget", "budget"),
                                                      ("Variance", "variance"), ("Cash", "cash")))
        ok = ok and near(s.get("InTransit") or 0, p["in_transit"]) and (p["mom"] is None and s.get("IncomeMoM") is None or near(s.get("IncomeMoM"), p["mom"], 1e-9))
        check(f"Power BI {m}: income, costs, net profit, budget, variance, cash, in transit, MoM % = Python", ok, json.dumps(s))
    t = snap["total"][0]
    check(f"Power BI totals: income {t['Income']:,.2f}, net profit {t['NetProfit']:,.2f}, budget {t['Budget']:,.0f} = Python",
          near(t["Income"], sum(py[m]["revenue"] for m in months)) and near(t["NetProfit"], sum(py[m]["net_profit"] for m in months))
          and near(t["Budget"], sum(py[m]["budget"] for m in months)) and near(t["Cash"], py[months[-1]]["cash"]))
    check(f"Power BI ledger: {t['LedgerRows']} lines (2 per bank line + 6 journal), debits = credits {t['Debits']:,.2f}",
          t["LedgerRows"] == 2 * len(lines) + 6 and near(t["Debits"], t["Credits"]))
    acc_bad = []
    for r in snap["accounts"]:
        a, m = r["Account No"], r["Month No"]
        want_actual = -net.get((a, m), 0) if types[a] in ("Revenue", "COGS", "Expense") else None
        if want_actual is not None and not near(r["Actual"] or 0, want_actual):
            acc_bad.append(f"{a}-{m} actual")
        if not near(r["Budget"] or 0, budget.get((a, m), 0)):
            acc_bad.append(f"{a}-{m} budget")
    n_pl = len({(a, m) for (a, m) in net if types[a] in ("Revenue", "COGS", "Expense")} | set(budget))
    check(f"Power BI P&L matrix: actual and budget for all {n_pl} account-months = Python", not acc_bad and len(snap["accounts"]) >= n_pl,
          ", ".join(acc_bad[:5]))
    st = [r for r in snap["statements"] if r["Statement"] is not None]
    check(f"Power BI books vs statements: {len(st)} account-months, every difference 0.00",
          len(st) == 9 and all(near(r["Difference"], 0) and near(r["Books"], statements[(r["Account No"], r["Month No"])]) for r in st))
    cf_ok = all(near(cashflow.get((r["Cash Flow Line"], r["Month No"]), 0), r["Net"]) for r in snap["cashflow"])
    check(f"Power BI cash flow by month: {len(snap['cashflow'])} line-months = Python", cf_ok and len(snap["cashflow"]) == len([k for k in cashflow]))
    # the KPI notes and the Key insights panel (whole period, no slicer)
    inc, exp_ = sum(py[m]["revenue"] for m in months), sum(py[m]["cogs"] + py[m]["expenses"] for m in months)
    inc_b = sum(v for (a, m), v in budget.items() if types[a] == "Revenue")
    exp_b = -sum(v for (a, m), v in budget.items() if types[a] in ("COGS", "Expense"))
    cash_change = py[months[-1]]["cash"] - sum(opening[a] for a in cash_accts)
    var_acct = defaultdict(float)
    for (a, m), v in net.items():
        if types[a] in ("Revenue", "COGS", "Expense"):
            var_acct[a] += -v
    for (a, m), v in budget.items():
        var_acct[a] -= v
    worst, best = min(var_acct, key=var_acct.get), max(var_acct, key=var_acct.get)
    label = lambda a: f"{a} {accounts[a]['Account']}"
    k = snap["kpi"][0]
    check(f"Power BI KPI notes: income {(inc - inc_b) / inc_b:+.1%} and costs {(exp_ - exp_b) / exp_b:+.1%} vs budget, "
          f"cash change {cash_change:,.2f} = Python",
          near(k["IncomeVsBudget"], (inc - inc_b) / inc_b, 1e-9) and near(k["ExpenseVsBudget"], (exp_ - exp_b) / exp_b, 1e-9)
          and near(k["CashChange"], cash_change))
    ins = {r["No"]: r["Text"] for r in snap["insights"]}
    check(f"Power BI Key insights: biggest overrun {label(worst)} ({var_acct[worst]:+,.0f}), biggest saving {label(best)} ({var_acct[best]:+,.0f})",
          ins.get(3) == f"{label(worst)}: {var_acct[worst]:+,.0f} vs budget"
          and ins.get(4) == f"{label(best)}: {var_acct[best]:+,.0f} vs budget", json.dumps(ins))
    # ---- 13-week cash forecast, recomputed from cash-plan.csv and the bank lines
    import calendar
    from datetime import timedelta

    def add_months(d, k):
        y, m = divmod(d.month - 1 + k, 12)
        y, m = d.year + y, m + 1
        return d.replace(year=y, month=m, day=min(d.day, calendar.monthrange(y, m)[1]))

    def month_end(d):
        return d.replace(day=calendar.monthrange(d.year, d.month)[1])

    last = max(l["post"] for l in lines)
    start, end = last + timedelta(days=1), last + timedelta(days=91)
    fc = defaultdict(float)
    for r in rows("cash-plan.csv"):
        accts = {int(a) for a in r["Accounts"].split(";")}
        hist = [l for l in lines if l["reg"] == int(r["Register No"]) and l["acct"] in accts]
        lb = int(r["Lookback"]) if r["Lookback"] else 0
        if r["Basis"] == "fixed":
            per = float(r["Amount"])
        elif r["Basis"] == "avg_week":
            per = sum(l["amount"] for l in hist if l["post"] > last - timedelta(days=7 * lb)) / lb
        elif r["Basis"] == "avg_month":
            cut = month_end(add_months(last, -lb))
            per = sum(l["amount"] for l in hist if l["post"] > cut) / lb
        else:  # avg_event: the last N posting dates (pay runs), averaged
            by_day = defaultdict(float)
            for l in hist:
                by_day[l["post"]] += l["amount"]
            top = [by_day[k] for k in sorted(by_day, reverse=True)[:lb]]
            per = sum(top) / len(top)
        # Power Query's Number.Round rounds half to even (banker's rounding): -6530.185 -> -6530.18
        per = float(Decimal(f"{per:.6f}").quantize(Decimal("0.01"), ROUND_HALF_EVEN))
        first = datetime.strptime(r["First Date"], "%Y-%m-%d").date()
        if r["Frequency"] == "weekly":
            dates = [start + timedelta(days=7 * k) for k in range(13)]
        elif r["Frequency"] == "biweekly":
            dates = [first + timedelta(days=14 * k) for k in range(8)]
        else:
            dates = [add_months(first, k) for k in range(4)]
        for dd in dates:
            if start <= dd <= end:
                fc[(dd - start).days // 7 + 1] += per
    today = py[months[-1]]["cash"]
    bal, run = {}, today
    for w in range(1, 14):
        run += fc[w]
        bal[w] = run
    sw = {r["Week No"]: r for r in snap["forecast"]}
    check(f"Power BI 13-week forecast: every week's net and closing cash = Python (week 13: {bal[13]:,.2f})",
          len(sw) == 13 and all(near(sw[w]["Net"] or 0, fc[w], 0.02) and near(sw[w]["Balance"], bal[w], 0.05) for w in range(1, 14)),
          json.dumps({w: [round(fc[w], 2), round(bal[w], 2)] for w in bal}))
    low = min(bal.values())
    below = sum(b < 10000 for b in bal.values())
    fk = snap["forecastKpi"][0]
    check(f"Power BI forecast KPIs: cash today {today:,.2f}, lowest {low:,.2f}, headroom {low - 10000:+,.2f} over 10,000, "
          f"{below} weeks below",
          near(fk["Today"], today) and near(fk["Lowest"], low, 0.05) and near(fk["Headroom"], low - 10000, 0.05)
          and fk["Below"] == below, json.dumps(fk))
    # ---- revenue and costs page
    rv = snap["revenue"][0]
    inc_all = sum(py[m]["revenue"] for m in months)
    cogs_all = sum(py[m]["cogs"] for m in months)
    whole = -sum(v for (a, m), v in net.items() if a == 4100)
    payroll = sum(v for (a, m), v in net.items() if a == 6000)
    check(f"Power BI revenue page: gross margin {(inc_all - cogs_all) / inc_all:.1%}, wholesale {whole / inc_all:.1%}, "
          f"payroll {payroll / inc_all:.1%} of income = Python",
          near(rv["GM"], (inc_all - cogs_all) / inc_all, 1e-9) and near(rv["WS"], whole / inc_all, 1e-9)
          and near(rv["Payroll"], payroll / inc_all, 1e-9), json.dumps(rv))
    payee_of = {}
    for r in sorted(rows("rules.csv"), key=lambda r: int(r["Priority"])):
        payee_of.setdefault(norm(r["Keyword"]), r["Payee"])
    spend = defaultdict(float)
    for l in lines:
        if types[l["acct"]] in ("COGS", "Expense"):
            k = next((k for k, a in rules if k in norm(l["desc"])), None)
            spend[payee_of[k] if k else l["desc"]] += -l["amount"]
    sp = {r["Payee"]: r["Spend"] for r in snap["payees"] if r["Spend"] is not None}
    top = max(spend, key=spend.get)
    check(f"Power BI payees: {len(spend)} payees, top {top} {spend[top]:,.2f}, total {sum(spend.values()):,.2f} = Python",
          set(sp) == set(spend) and all(near(sp[p], spend[p]) for p in spend), str(sorted(set(sp) ^ set(spend)))[:200])
    # ---- ledger page: running balance of checking after every line
    chk = sorted([l for l in lines if l["reg"] == 1000], key=lambda l: (l["post"], l["id"]))
    run, want = opening[1000], {}
    for l in chk:
        run = round(run + l["amount"], 2)
        want[l["id"] + "-R"] = run
    got = {r["Line ID"]: r["Balance"] for r in snap["running"] if r["Line ID"].endswith("-R")}
    check(f"Power BI ledger page: running balance after all {len(want)} checking lines = Python (last {run:,.2f})",
          set(got) == set(want) and all(near(got[k], want[k]) for k in want),
          str([k for k in want if not near(got.get(k), want[k])][:5]))
    statuses = {str(r["No"]): r["Status"] for r in snap["checks"]}
    check("Power BI Checks page: 8 PASS + 1 REVIEW (3 lines waiting for a rule), 0 FAIL", statuses == exp["checks"], str(statuses))
    by_id = {l["id"]: l for l in lines}
    line_bad = [r["Line"] for r in snap["lines"] if not (r["Line"][:-2] in by_id and by_id[r["Line"][:-2]]["acct"] == r["Account"]
                                                         and near(by_id[r["Line"][:-2]]["amount"], r["Amount"])
                                                         and by_id[r["Line"][:-2]]["post"].isoformat() == r["Date"])]
    check(f"Power BI gives all {len(snap['lines'])} bank lines the same account, amount and date as Python",
          len(snap["lines"]) == len(lines) and not line_bad, ", ".join(line_bad[:5]))

# ====================================================================== 5. the .pbix
pbix = ROOT / "Copperleaf-Finance.pbix"
try:
    names = zipfile.ZipFile(pbix).namelist()
    check("Copperleaf-Finance.pbix is a Power BI file with the data model inside", "DataModel" in names and any(n.startswith("Report") for n in names))
except Exception as e:
    check("Copperleaf-Finance.pbix is a Power BI file with the data model inside", False, str(e))

print(f"\n{sum(results)}/{len(results)} checks pass")
sys.exit(0 if all(results) else 1)
