"""Independent check of Bookkeeping-Trial-Balance.xlsx.

1. test/model.py re-does the bookkeeping in plain Python straight from the four CSV files (no Excel involved): parse,
   clean, categorise with the rules, pair transfers, post a ledger, trial balance, P&L, statements, cash flow statement,
   13-week forecast, payees, the dashboard numbers for every month and comparison, the month-end close checklist.
2. Compares it with the answer planted when the data was generated (test/expected.json).
3. Opens the workbook with Excel and compares every number on every report sheet with Python, for all three months and
   both comparisons (the two slicers are driven through the tables they filter).
4. Copies workbook + data to another folder, refreshes there, and checks everything again.
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs pywin32 for steps 3-4)
"""
import json, sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import model as M  # noqa: E402
from model import (ROOT, ACCOUNTS, ADJUSTMENTS, MONTHS, CLEARING, UNCATEGORIZED, BUDGET, CASH_BUFFER, TYPES, NAMES,  # noqa: E402
                   MONTH_NOS, LABELS, lines, pairs, unpaired, net, pl, tb, cash, stmts, balance, norm, r2)

results = []


def check(name, ok, detail=""):
    results.append(bool(ok))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


def near(a, b, tol=0.005):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


# ---------------------------------------------------------------- 2. Python vs the planted answer
exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
check("Rows read = rows planted (checking, savings, card)",
      [sum(l["reg"] == r for l in lines) for r in (1000, 1010, 2000)] == [exp["rows"]["checking"], exp["rows"]["savings"], exp["rows"]["card"]])
truth = defaultdict(list)
for t in exp["true_account"]:
    truth[(t["reg"], t["post"], norm(t["desc"]), t["amount"])].append(t["acct"])
wrong = []
for l in lines:
    k = (l["reg"], l["post"].isoformat(), norm(l["desc"]), l["amount"])
    want = truth[k].pop(0) if truth[k] else None
    if want != l["acct"]:
        wrong.append(f"{l['line']} {l['desc']}: {l['acct']} vs {want}")
check(f"Rules give all {len(lines)} lines their planted account", not wrong, "; ".join(wrong[:3]))
check("Transfers: 6 pairs, 1 half still in transit (the 30 Oct card payment)",
      len(pairs) == 6 and [round(u["amount"], 2) for u in unpaired] == [-1800.0], f"{len(pairs)} pairs, {[u['amount'] for u in unpaired]}")
got = {f"{a}-{mo}": round(v, 2) for (a, mo), v in net.items() if round(v, 2)}
check("Every account's movement per month = planted answer", got == exp["accounts_by_month"],
      str(set(got.items()) ^ set(exp["accounts_by_month"].items()))[:200])
for no in MONTH_NOS:
    check(f"{no}: P&L = planted answer", pl[no] == {k: round(v, 2) for k, v in exp["pl"][str(no)].items()}, f"{pl[no]}")
    check(f"{no}: trial balance balances in Python (debits {tb[no][0]:,.2f})", tb[no][0] == tb[no][1])
    check(f"{no}: cash (checking + savings) = planted", cash[no] == exp["cash"][str(no)])
    check(f"{no}: transfers in transit = planted", balance(CLEARING, no) == exp["in_transit_by_month"][str(no)])
    cf = M.cash_flows[no]
    check(f"{no}: cash flow statement (indirect) = change in checking + savings ({cf['net_change']:,.2f}) and ends at the planted cash",
          cf["net_change"] == cf["cash_moves"] and cf["end"] == exp["cash"][str(no)], str(cf))
bad = [(a, no) for (a, no), v in stmts.items() if round((-balance(a, no) if a == 2000 else balance(a, no)) - v, 2)]
check("Books = all 9 statement balances (Python)", not bad and len(stmts) == 9, str(bad))
check("Uncategorized lines = the 3 planted", sorted((l["desc"], l["amount"]) for l in lines if l["acct"] == UNCATEGORIZED)
      == sorted((u["desc"], u["amount"]) for u in exp["uncategorized"]))
dep = sum(dr for _, _, _, a, dr, _ in ADJUSTMENTS if a == 6900)
check(f"Payees: spend of all {len(M.payee_rows)} payees ({M.spend_total:,.2f}) = cost of goods + operating costs - depreciation",
      near(M.spend_total, sum(pl[no]["cogs"] + pl[no]["expenses"] for no in MONTH_NOS) - dep, 0.01))
check(f"Forecast uses the same books as E9: cash today {M.forecast['today']:,.2f}, lowest {M.forecast['lowest']:,.2f} in W{M.forecast['lowest_week']}",
      M.forecast["today"] == exp["cash"][str(MONTH_NOS[-1])] and M.forecast["lowest_week"] == 12 and round(M.forecast["lowest"]) == 29496)

# ---------------------------------------------------------------- 3. against the workbook
try:
    import win32com.client as win32
except ImportError:
    print("SKIP  workbook checks (pywin32 not installed)")
    win32 = None

DARK = ["Summary", "P&L", "Trial balance", "Balance sheet", "Cash flow", "Bank rec", "Review", "Month-end close", "Payees"]
GOLD = 201 + 163 * 256 + 91 * 65536


def find_row(ws, col, text, last=90):
    for r in range(1, last):
        if ws.Range(f"{col}{r}").Value == text:
            return r
    raise KeyError(text)


def name_vals(wb, nm):
    v = wb.Names(nm).RefersToRange.Value
    return [list(r) for r in v] if isinstance(v, tuple) else [[v]]


def pick(wb, month, mode):
    """What a click on the two slicers does: filter the two small tables on the Pickers sheet."""
    p = wb.Worksheets("Pickers")
    p.ListObjects("MonthPick").Range.AutoFilter(1, month[:3])
    p.ListObjects("CompareMode").Range.AutoFilter(1, "vs Budget" if mode == "budget" else "vs Last month")
    wb.Application.Calculate()


def fmt0(x):
    return f"{x:,.0f}"


def check_workbook(wb, tag=""):
    lo = wb.Worksheets("Transactions").ListObjects("Transactions")
    cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
    xrows = {r[cols.index("Line")]: r for r in lo.DataBodyRange.Value}
    diff = [l["line"] for l in lines if l["line"] not in xrows or int(xrows[l["line"]][cols.index("Account")]) != l["acct"]
            or round(xrows[l["line"]][cols.index("Amount")] - l["amount"], 2) or xrows[l["line"]][cols.index("Payee")] != l["payee"]]
    check(f"{tag}Every line in the workbook has the Python account, payee and amount ({len(lines)} lines)", not diff and len(xrows) == len(lines), str(diff[:3]))
    led = wb.Worksheets("Ledger").ListObjects("Ledger").ListRows.Count
    check(f"{tag}Ledger = 2 lines per bank line + {len(ADJUSTMENTS)} adjustment lines", led == 2 * len(lines) + len(ADJUSTMENTS), str(led))
    # ---- P&L
    p = wb.Worksheets("P&L")
    rows = {k: find_row(p, "C", t) for k, t in (("revenue", "Total revenue"), ("cogs", "Total cost of goods sold"),
                                                 ("gross_profit", "Gross profit"), ("expenses", "Total operating expenses"), ("net_profit", "Net profit"))}
    for k, no in enumerate(MONTH_NOS):
        c = chr(ord("D") + 3 * k)
        got = {key: round(p.Range(f"{c}{r}").Value, 2) for key, r in rows.items()}
        check(f"{tag}{no}: workbook P&L = Python", got == pl[no], f"{got}")
        bud = round(p.Range(f"{chr(ord('E') + 3 * k)}{rows['net_profit']}").Value, 2)
        check(f"{tag}{no}: budget net profit = Budget tab", bud == M.budget_pl[no]["net_profit"], f"{bud}")
    # ---- trial balance, balance sheet
    t = wb.Worksheets("Trial balance")
    tot = find_row(t, "C", "Total")
    for k, no in enumerate(MONTH_NOS):
        a, b = chr(ord("G") + 2 * k), chr(ord("H") + 2 * k)
        got = (round(t.Range(f"{a}{tot}").Value, 2), round(t.Range(f"{b}{tot}").Value, 2))
        check(f"{tag}{no}: trial balance totals = Python and balanced", got == tb[no], f"{got} vs {tb[no]}")
    b = wb.Worksheets("Balance sheet")
    bchk = find_row(b, "C", "Balanced?")
    check(f"{tag}Balance sheet balances every month end (PASS x3)", [b.Range(f"{c}{bchk}").Value for c in "DEF"] == ["PASS"] * 3)
    # ---- bank rec
    r = wb.Worksheets("Bank rec")
    first = next(i for i in range(1, 40) if r.Range(f"N{i}").Value)
    recs = [(r.Range(f"N{i}").Value, r.Range(f"E{i}").Value, r.Range(f"F{i}").Value, r.Range(f"H{i}").Value) for i in range(first, first + 9)]
    check(f"{tag}Bank rec: 9 rows, books = statement = Python, all PASS",
          all(st == "PASS" and round(bk - stmts[(int(a), MONTH_NOS[i % 3])], 2) == 0 and round(sb - stmts[(int(a), MONTH_NOS[i % 3])], 2) == 0
              for i, (a, bk, sb, st) in enumerate(recs)), str(recs[:2]))
    agreed = r.Range(f"J{first}").Value
    check(f"{tag}Bank rec cards: 9 / 9 agreed, total difference 0.00, {len(lines)} lines",
          agreed == "9 / 9" and near(r.Range(f"J{first + 6}").Value, 0) and r.Range(f"J{first + 12}").Value == len(lines),
          f"{agreed} {r.Range(f'J{first + 6}').Value} {r.Range(f'J{first + 12}').Value}")
    # ---- review
    rv = wb.Worksheets("Review").ListObjects("Review")
    rvals = sorted((x[0], x[4], round(x[5], 2)) for x in rv.DataBodyRange.Value)
    want = sorted([("No rule matched", u["desc"], u["amount"]) for u in exp["uncategorized"]] +
                  [("Transfer without its other half", u["desc"], round(u["amount"], 2)) for u in unpaired])
    check(f"{tag}Review tab = the 3 unknown lines + the unpaired transfer half", rvals == want, str(rvals))
    # ---- payees
    py = wb.Worksheets("Payees").ListObjects("Payees")
    pcols = [py.ListColumns(i).Name for i in range(1, py.ListColumns.Count + 1)]
    xp = [list(x) for x in py.DataBodyRange.Value]
    want_p = [[x["payee"], x["account"], x["n"]] + x["months"] + [x["total"], x["share"]] for x in M.payee_rows]
    okp = len(xp) == len(want_p) and all(a[:3] == w[:3] and all(near(a[i], w[i], 0.005 if i < 7 else 1e-9) for i in range(3, 8))
                                         for a, w in zip(xp, want_p))
    check(f"{tag}Payees tab: all {len(want_p)} payees, payments, spend per month, total and share = Python (biggest first)",
          okp and pcols == ["Payee", "Account", "Payments"] + LABELS + ["Total", "Share"], str(xp[:2]))
    # ---- cash flow statement
    cfw = wb.Worksheets("Cash flow")
    keys = [("Net profit", "net_profit"), ("Add back: depreciation (no cash leaves)", "depreciation"),
            ("(Increase) / decrease in transfers in transit", "transit"), ("Increase / (decrease) in card balance owed", "card"),
            ("Net cash from operating activities", "operating"), ("Kitchen equipment bought", "equipment"),
            ("Net cash from investing activities", "investing"), ("Owner's draws", "draws"), ("Owner's capital put in", "capital"),
            ("Net cash from financing activities", "financing"), ("Net change in cash", "net_change"),
            ("Cash at the start (checking + savings)", "start"), ("Cash at the end", "end")]
    bad = []
    for label, key in keys:
        rr = find_row(cfw, "B", label)
        for c, per in zip("CDEF", MONTH_NOS + ["quarter"]):
            if not near(cfw.Range(f"{c}{rr}").Value, M.cash_flows[per][key]):
                bad.append(f"{label} {per}: {cfw.Range(f'{c}{rr}').Value} vs {M.cash_flows[per][key]}")
    check(f"{tag}Cash flow statement: 13 lines x (3 months + quarter) = Python", not bad, "; ".join(bad[:3]))
    tie = find_row(cfw, "B", "Cash flow ties to the balance sheet?")
    check(f"{tag}Cash flow ties to the balance sheet (PASS x4)", [cfw.Range(f"{c}{tie}").Value for c in "CDEF"] == ["PASS"] * 4)
    # ---- 13-week forecast
    w0 = find_row(cfw, "H", "W1")
    badw = []
    for i, w in enumerate(M.weeks):
        rr = w0 + i
        got = [cfw.Range(f"{c}{rr}").Value for c in "JKLM"]
        if not all(near(g, v, 0.011) for g, v in zip(got, (w["money_in"], w["money_out"], w["net"], w["cash"]))):
            badw.append(f"W{w['week']}: {got}")
        if cfw.Range(f"I{rr}").Value is None or str(cfw.Range(f"I{rr}").Value)[:10] != w["end"].isoformat():
            badw.append(f"W{w['week']} end {cfw.Range(f'I{rr}').Value}")
    check(f"{tag}13-week forecast: money in, out, net and cash for every week = Python (W13 {M.weeks[-1]['cash']:,.2f})", not badw, "; ".join(badw[:3]))
    fc = M.forecast
    nm = {n: wb.Names(n).RefersToRange.Value for n in ("FcToday", "FcLowest", "FcHeadroom", "FcWeeks", "FcRunway", "FcLowWeek")}
    check(f"{tag}Forecast cards: today {fc['today']:,.2f}, lowest {fc['lowest']:,.2f} (W{fc['lowest_week']}), headroom {fc['headroom']:+,.2f}, "
          f"{fc['weeks_of_cash']} weeks of cash, runway {fc['runway_months']} months",
          near(nm["FcToday"], fc["today"]) and near(nm["FcLowest"], fc["lowest"], 0.02) and near(nm["FcHeadroom"], fc["headroom"], 0.02)
          and round(nm["FcWeeks"], 1) == fc["weeks_of_cash"] and round(nm["FcRunway"], 1) == fc["runway_months"] and nm["FcLowWeek"] == fc["lowest_week"],
          json.dumps(nm, default=str))
    plan = wb.Worksheets("Cash plan").ListObjects("CashPlanView")
    pc = [plan.ListColumns(i).Name for i in range(1, plan.ListColumns.Count + 1)]
    xpl = [dict(zip(pc, x)) for x in plan.DataBodyRange.Value]
    check(f"{tag}Cash plan: every item's amount each time, times and 13-week total = Python ({len(M.plan)} items)",
          len(xpl) == len(M.plan) and all(x["Item"] == p_["item"] and near(x["Each time"], p_["per"]) and x["Times"] == p_["times"]
                                          and near(x["Total 13 weeks"], p_["total"], 0.01) for x, p_ in zip(xpl, M.plan)), str(xpl[:2]))
    # ---- month-end close
    cl = wb.Worksheets("Month-end close")
    r1 = find_row(cl, "C", "Bank, savings and card exports loaded for the month")
    order = ["exports", "rec1000", "rec1010", "rec2000", "transfers", "adjustments", "tb", "bs", "cashflow", "review"]
    badc = []
    for k in range(3):
        stp = M.close_checklist(k)
        for i, key in enumerate(order):
            v = stp[key]
            want_s = ("PASS" if v[0] == 0 else "OPEN") if key == "review" else ("PASS" if v else "FAIL")
            got_s = cl.Range(f"{'DEF'[k]}{r1 + i}").Value
            if got_s != want_s:
                badc.append(f"{LABELS[k]} {key}: {got_s} vs {want_s}")
        det = cl.Range(f"{'DEF'[k]}{r1 + 10}").Value
        if det != f"{stp['review'][0]} · {stp['review'][1]:,.2f}":
            badc.append(f"{LABELS[k]} detail {det}")
    check(f"{tag}Month-end close: 10 steps x 3 months = Python (9 PASS + Review OPEN each month)", not badc, "; ".join(badc[:3]))
    # ---- slicers, navigation, no error cells
    s = wb.Worksheets("Summary")
    check(f"{tag}Summary has the two slicers (month, compare)", s.Shapes.Count >= 2 and wb.SlicerCaches.Count == 2
          and sorted(wb.SlicerCaches(i).Slicers(1).Name for i in (1, 2)) == ["Compare with", "Month picker"])
    navbad = []
    for sh in DARK:
        ws = wb.Worksheets(sh)
        for i in range(1, 10):
            shp = ws.Shapes(f"nav_{i}")
            target = shp.Hyperlink.SubAddress.split("!")[0].strip("'")
            if target not in [x.Name for x in wb.Worksheets]:
                navbad.append(f"{sh}: {target}")
            if (target == sh) != (shp.Fill.ForeColor.RGB == GOLD):
                navbad.append(f"{sh}: highlight on {target}")
    check(f"{tag}Navigation menu on all {len(DARK)} report sheets: 9 links each, current sheet highlighted", not navbad, "; ".join(navbad[:3]))
    errs = []
    for sh in DARK + ["Calc"]:
        v = wb.Worksheets(sh).UsedRange.Value
        for row in (v if isinstance(v, tuple) else ((v,),)):
            for x in row:
                if isinstance(x, int) and x <= -2146826000:
                    errs.append(sh)
    check(f"{tag}No #N/A, #VALUE!, #NAME? or #REF! on the report sheets", not errs, str(sorted(set(errs))))

    # ---- the Summary for every month and both comparisons
    sc_m = next(wb.SlicerCaches(i) for i in (1, 2) if wb.SlicerCaches(i).Slicers(1).Name == "Month picker")
    for k in range(3):
        for mode in ("budget", "last"):
            pick(wb, LABELS[k], mode)
            d = M.dashboard(k, mode)
            tg = f"{tag}{LABELS[k]} vs {'budget' if mode == 'budget' else 'last month'}: "
            sel = [sc_m.SlicerItems(i).Name for i in range(1, sc_m.SlicerItems.Count + 1) if sc_m.SlicerItems(i).Selected]
            kpi = name_vals(wb, "CalcKpi")
            want_a = [d["revenue"], d["net_profit"], d["gross_margin"], d["opex_pct"], d["cash"], d["in_transit"], d["review_lines"], d["weeks_of_cash"]]
            want_c = [d["c_revenue"], d["c_net_profit"], d["c_gross_margin"], d["c_opex_pct"], r2(d["cash"] - d["cash_change"])]
            ok = sel == [LABELS[k][:3]] and all(near(kpi[i][1], w, 1e-6 if i in (2, 3) else 0.005) for i, w in enumerate(want_a[:7]))
            ok = ok and round(kpi[7][1], 1) == want_a[7]
            ok = ok and all((kpi[i][2] in (None, "") and w is None) or near(kpi[i][2], w, 1e-6 if i in (2, 3) else 0.005) for i, w in enumerate(want_c))
            rv_amt = r2(sum(-l["amount"] for l in lines if l["mo"] == MONTH_NOS[k] and l["acct"] == UNCATEGORIZED)
                        + sum(-u["amount"] for u in unpaired if u["mo"] == MONTH_NOS[k]))
            ok = ok and near(kpi[6][2], rv_amt)
            cards = [s.Range(f"{chr(ord('B') + 4 * i) if i < 6 else ('Z', 'AD')[i - 6]}12").Value for i in range(8)]
            ok = ok and all(near(cv, kv[1], 1e-9) for cv, kv in zip(cards, kpi))
            check(tg + "slicer shows the month; 8 KPI cards (revenue, net profit, GM %, opex %, cash, in transit, lines to review, "
                       "weeks of cash) and what they are compared with = Python", ok, f"{sel} {kpi}")
            notes = [x[4] for x in kpi]
            c = d["c_revenue"]
            want_note = ("no earlier month in this file" if c is None else
                         f"{'▲' if d['revenue'] >= c else '▼'} {fmt0(abs(d['revenue'] - c))} vs "
                         f"{'budget' if mode == 'budget' else LABELS[k - 1][:3]} {fmt0(c)}")
            check(tg + f"KPI note lines written from the numbers ('{want_note}')", notes[0] == want_note
                  and notes[4] == f"{'▲' if d['cash_change'] >= 0 else '▼'} {fmt0(abs(d['cash_change']))} in the month", str(notes[:5]))
            wf = name_vals(wb, "CalcWaterfall")
            p_ = pl[MONTH_NOS[k]]
            steps = [(0, p_["revenue"]), (p_["revenue"], p_["gross_profit"]), (p_["gross_profit"], p_["net_profit"]), (0, p_["net_profit"])]
            okw = True
            for (st, en), row in zip(steps, wf):
                lo_, hi_ = min(st, en), max(st, en)
                base = lo_ if lo_ >= 0 else hi_ if hi_ <= 0 else 0
                okw = okw and near(row[1], st) and near(row[2], en) and near(row[3], base) and near(row[4] - row[5], hi_ - lo_)
            okw = okw and [row[7] for row in wf] == [f"{x:,.0f}" for x in (p_["revenue"], -p_["cogs"], -p_["expenses"], p_["net_profit"])]
            check(tg + f"profit waterfall: revenue {fmt0(p_['revenue'])} - COGS - opex = net profit {fmt0(p_['net_profit'])} (bars and labels)", okw, str(wf))
            costs = name_vals(wb, "CalcCosts")
            okc = len(costs) == len(d["costs"]) and all(
                x[1] == f"{w['acct']} {NAMES[w['acct']]}" and near(x[2], w["value"])
                and ((w["var"] is None and x[3] in ("", None)) or near(x[3], w["var"])) and all(near(a, b_) for a, b_ in zip(x[4:7], w["trend"]))
                for x, w in zip(costs, d["costs"]))
            check(tg + f"costs by account, sorted biggest first ({costs[0][1]}), variance and 3-month sparkline data = Python", okc,
                  str([(x[1], x[2]) for x in costs[:3]]))
            pay = name_vals(wb, "CalcPayees")
            mo = MONTH_NOS[k]
            month_total = r2(sum(x["months"][k] for x in M.payee_rows))
            npay = lambda nm_: sum(1 for l in lines if l["payee"] == nm_ and l["mo"] == mo and TYPES[l["acct"]] in ("COGS", "Expense"))
            okp = all(x[1] == w[0] and near(x[2], w[1]) and x[3] == npay(w[0]) and near(x[4], w[1] / month_total, 1e-9)
                      for x, w in zip(pay[:6], d["payees"]))
            okp = okp and near(pay[6][2], month_total - sum(w[1] for w in d["payees"]), 0.01) and near(pay[7][2], month_total, 0.01)
            okp = okp and pay[7][3] == sum(1 for l in lines if l["mo"] == mo and TYPES[l["acct"]] in ("COGS", "Expense"))
            check(tg + f"who got paid: top 6 payees ({pay[0][1]} {fmt0(pay[0][2])}), other, total {fmt0(month_total)}, payments, share = Python", okp,
                  str(pay[:3]))
            ins = [x[1] for x in name_vals(wb, "CalcInsights")]
            comp_lbl = "budget" if mode == "budget" else (LABELS[k - 1][:3] if k else "-")
            npv = d["net_profit"]
            w1 = f"{LABELS[k]}: {'profit' if npv >= 0 else 'loss'} {fmt0(abs(npv))}"
            if d["c_net_profit"] is not None:
                w1 += f" vs {comp_lbl} {fmt0(d['c_net_profit'])} ({npv - d['c_net_profit']:+,.0f})"
            if "worst" in d:
                w2 = f"{d['worst']} {NAMES[d['worst']]}: {d['worst_var']:+,.0f} vs {comp_lbl}"
                w3 = f"{d['best']} {NAMES[d['best']]}: {d['best_var']:+,.0f} vs {comp_lbl}"
            else:
                w2, w3 = "no earlier month to compare with", ""
            w4 = (f"Cash {'up' if d['cash_change'] >= 0 else 'down'} {fmt0(abs(d['cash_change']))} in {LABELS[k][:3]} to {fmt0(d['cash'])}"
                  + (f", {fmt0(d['in_transit'])} in transit" if round(d["in_transit"], 2) > 0 else ""))
            fc_ = M.forecast
            w5 = (f"Lowest cash {fmt0(fc_['lowest'])} in W{fc_['lowest_week']} ({fc_['lowest_end'].day} {fc_['lowest_end']:%b}), "
                  f"{fmt0(abs(fc_['headroom']))} {'above' if fc_['headroom'] >= 0 else 'below'} the {fmt0(CASH_BUFFER)} buffer")
            check(tg + f"key insights written by formulas: '{w2}'", ins == [w1, w2, w3, w4, w5], str(ins))
            closesel = name_vals(wb, "CloseSel")
            stp = M.close_checklist(k)
            check(tg + "month-end close panel follows the month (9 of 10, Review open)",
                  [x[1] for x in closesel[:10]].count("PASS") == 9 and closesel[9][1] == ("OPEN" if stp["review"][0] else "PASS")
                  and closesel[10][1] == "9 of 10" and closesel[11][1] == "Open items", str(closesel[-3:]))
    pick(wb, LABELS[-1], "budget")
    states = list(zip([x[0] for x in name_vals(wb, "CheckLabels")], [x[0] for x in name_vals(wb, "CheckStatus")]))
    for label, st in states:
        check(f"{tag}Workbook data check: {label}", st == "PASS")


if win32:
    xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(ROOT / "Bookkeeping-Trial-Balance.xlsx"), 0, True)
        check_workbook(wb)
        wb.Close(False)

        import shutil, tempfile
        with tempfile.TemporaryDirectory() as tmp:
            shutil.copy(ROOT / "Bookkeeping-Trial-Balance.xlsx", tmp)
            shutil.copytree(ROOT / "data", Path(tmp) / "data")
            wb = xl.Workbooks.Open(str(Path(tmp) / "Bookkeeping-Trial-Balance.xlsx"))
            wb.RefreshAll(); xl.CalculateUntilAsyncQueriesDone(); xl.Calculate()
            folder = wb.Worksheets("Settings").Range("B4").Value
            check("Refresh in another folder reads that folder's data", str(folder).lower().startswith(str(tmp).lower()[:12]), folder)
            check_workbook(wb, "After refresh: ")
            wb.Close(False)
    finally:
        xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
