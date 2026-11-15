"""Independent check of Sales-Commission-Calculator.xlsx.

1. Recomputes in plain Python, from data/orders.csv, returns.csv, payouts.csv and plan.json (no Excel involved):
   every order line (price, discount, net, tax, split credit, status), every return (returned value, status,
   clawback), and the pay ledger for 12 months x 6 reps for both commission methods: commission, accelerator above
   120% of target, target bonus, manager override, clawback, total pay, payout "to pay", and the What-if scenario.
2. Opens the workbook in Excel (read-only, never saved), switches month, method, rep and scenario, and compares the
   ledger, Commission, Dashboard, Payouts, What-if and every rep's Statement with the Python numbers.
3. Checks the planted mistakes are exactly what is held back, all test cases and data checks say PASS, inputs are
   unlocked and validated, formula sheets are protected, there are no error values, and no formula needs Microsoft 365.
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs pywin32)
"""
import csv, json, sys
from collections import defaultdict
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
PLAN = json.loads((ROOT / "data" / "plan.json").read_text())
PLANTED = json.loads((ROOT / "test" / "planted.json").read_text())
results = []


def check(name, ok, detail=""):
    results.append(bool(ok))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""), flush=True)


def r2(x):
    """Round to cents the way Excel does (15 significant digits first, then half away from zero)."""
    return float(Decimal(f"{x:.15g}").quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


close = lambda a, b, tol=0.005: abs((a or 0) - (b or 0)) <= tol
num = lambda v: 0.0 if v in ("", None) else float(v)
price = {p[0]: p[2] for p in PLAN["products"]}
maxd = {p[0]: p[3] for p in PLAN["products"]}
REPS = [r[0] for r in PLAN["reps"]]
target = {r[0]: r[2] for r in PLAN["reps"]}
manager = {r[0]: r[3] for r in PLAN["reps"]}
FIRST = date.fromisoformat(PLAN["first_month"])
MONTHS = [date(FIRST.year + (FIRST.month - 1 + i) // 12, (FIRST.month - 1 + i) % 12 + 1, 1) for i in range(PLAN["months"])]
SC = PLAN["scenario"]
serial = lambda d: (d - date(1899, 12, 30)).days
from_serial = lambda s: date.fromordinal(date(1899, 12, 30).toordinal() + int(s))


def commission(net, method, bands):
    if method == "Stepped":
        total, prev = 0.0, 0.0
        for frm, rate in bands:
            if net > frm:
                total += (net - frm) * (rate - prev)
            prev = rate
        return r2(total)
    rate = [rt for frm, rt in bands if net >= frm][-1]
    return r2(net * rate)


def accelerator(net, tgt, at, rate):
    return r2(max(0.0, net - at * tgt) * rate) if tgt > 0 else 0.0


# ================================================================ 1. plain Python
lines = []
with open(ROOT / "data" / "orders.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        d = date.fromisoformat(r["Date"]) if r["Date"] else None
        qty, disc, prod, rep = int(r["Qty"]), float(r["Discount %"]), r["Product"], r["Sales rep"]
        srep, spct = r["Split rep"], float(r["Split %"]) if r["Split %"] else 0.0
        net = tax = split_cr = rep_cr = None
        if prod in price:
            gross = r2(qty * price[prod]); net = r2(gross - r2(gross * disc)); tax = r2(net * PLAN["tax_rate"])
            split_cr = r2(net * spct) if srep else 0.0
            rep_cr = net - split_cr
        if not d: status = "Missing date"
        elif rep not in target: status = "Unknown sales rep"
        elif prod not in price: status = "Unknown product"
        elif qty < 1: status = "Qty must be 1 or more"
        elif disc > maxd[prod] + 1e-6: status = f"Discount above {maxd[prod]:.0%} max"
        elif srep and srep not in target: status = "Unknown split rep"
        elif srep and srep == rep: status = "Split rep same as sales rep"
        elif srep and not (0.01 <= spct <= 0.99): status = "Split % must be 1% to 99%"
        elif not srep and spct != 0: status = "Split % without a split rep"
        else: status = "OK"
        lines.append(dict(order=r["Order no"], date=d, month=d.replace(day=1) if d else None, rep=rep, prod=prod, qty=qty,
                          net=net, tax=tax, split_rep=srep, split_pct=spct, split_cr=split_cr, rep_cr=rep_cr, status=status,
                          cust=r["Customer"]))
held = sorted((l["order"], l["status"]) for l in lines if l["status"] != "OK")
check(f"Order lines held back = the {len(PLANTED['orders'])} planted mistakes, with the right reason",
      held == sorted((p["order"], p["why"]) for p in PLANTED["orders"]), str(held))
split_lines = [l for l in lines if l["split_rep"] and l["status"] == "OK"]
ok_lines = [l for l in lines if l["status"] == "OK"]

key_index = {}
for i, l in enumerate(lines):
    key_index.setdefault(f"{l['order']}|{l['prod']}", i)
rets = []
with open(ROOT / "data" / "returns.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        rd = date.fromisoformat(r["Return date"]) if r["Return date"] else None
        qr = int(r["Qty returned"])
        i = key_index.get(f"{r['Order no']}|{r['Product']}")
        l = lines[i] if i is not None else None
        if not rd: st = "Missing date"
        elif l is None: st = "Order line not found"
        elif qr < 1: st = "Qty must be 1 or more"
        elif qr > l["qty"]: st = "Qty above qty sold"
        elif rd < l["date"]: st = "Return before order date"
        elif l["status"] != "OK": st = "Line was held back (never paid)"
        else: st = "OK"
        ret_net = r2(l["net"] * qr / l["qty"]) if l and l["net"] is not None and l["qty"] else None
        rets.append(dict(order=r["Order no"], date=rd, booked=rd.replace(day=1) if rd else None, line=l, qty=qr,
                         ret_net=ret_net, status=st))
bad_rets = sorted((x["order"], x["status"]) for x in rets if x["status"] != "OK")
check(f"Returns held back = the {len(PLANTED['returns'])} planted mistakes", bad_rets == sorted((p["order"], p["why"]) for p in PLANTED["returns"]),
      str(bad_rets))

payout = {}
with open(ROOT / "data" / "payouts.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        payout[(date.fromisoformat(r["Month"]), r["Sales rep"])] = (r["Status"], r["Paid date"])


def py_ledger(method, bands=PLAN["bands"], sc=SC):
    led = {}
    for m in MONTHS:
        for rep in REPS:
            mine = [l for l in ok_lines if l["month"] == m and l["rep"] == rep]
            shared = [l for l in ok_lines if l["month"] == m and l["split_rep"] == rep]
            net = r2(sum(l["rep_cr"] for l in mine) + sum(l["split_cr"] for l in shared))
            tgt = target[rep]
            led[(m, rep)] = dict(lines=len(mine) + len(shared), net=net, comm=commission(net, method, bands),
                                 accel=accelerator(net, tgt, PLAN["accel_at"], PLAN["accel_rate"]),
                                 bonus=PLAN["target_bonus"] if tgt > 0 and net >= tgt else 0, claw=0.0)
    for (m, rep), e in led.items():
        e["override"] = r2(PLAN["override_rate"] * sum(led[(m, x)]["net"] for x in REPS if manager[x] == rep))
    for x in rets:
        x["claw_rep"] = x["claw_split"] = 0.0
        if x["status"] != "OK":
            continue
        l = x["line"]
        split_share = r2(x["ret_net"] * l["split_pct"]) if l["split_rep"] else 0.0
        rep_share = x["ret_net"] - split_share
        om = l["month"]
        o = led[(om, l["rep"])]
        x["claw_rep"] = r2(rep_share * (o["comm"] / o["net"] if o["net"] else 0))
        if l["split_rep"]:
            o2 = led[(om, l["split_rep"])]
            x["claw_split"] = r2(split_share * (o2["comm"] / o2["net"] if o2["net"] else 0))
        if (x["booked"], l["rep"]) in led:
            led[(x["booked"], l["rep"])]["claw"] += x["claw_rep"]
        if l["split_rep"] and (x["booked"], l["split_rep"]) in led:
            led[(x["booked"], l["split_rep"])]["claw"] += x["claw_split"]
    for (m, rep), e in led.items():
        e["total"] = e["comm"] + e["accel"] + e["bonus"] + e["override"] - e["claw"]
        st, pd = payout[(m, rep)]
        e["status"], e["paid"] = st, pd
        e["to_pay"] = 0.0 if st == "Paid" else e["total"]
        tgt2 = r2(target[rep] * (1 + sc["target_change"]))
        e["wi_comm"] = commission(e["net"], method, sc["bands"])
        e["wi_accel"] = accelerator(e["net"], tgt2, sc["accel_at"], sc["accel_rate"])
        e["wi_bonus"] = sc["target_bonus"] if tgt2 > 0 and e["net"] >= tgt2 else 0
        e["wi_pay"] = e["wi_comm"] + e["wi_accel"] + e["wi_bonus"] + e["override"] - e["claw"]
    return led


LED = {m: py_ledger(m) for m in ("Whole amount", "Stepped")}
W = LED["Whole amount"]
check("Sample data exercises every rule (split deals, returns on split lines, accelerator, overrides, all payout statuses)",
      len(split_lines) > 50 and any(x["status"] == "OK" and x["line"]["split_rep"] for x in rets)
      and sum(1 for e in W.values() if e["accel"] > 0) >= 3 and sum(1 for e in W.values() if e["override"] > 0) >= 12
      and {e["status"] for e in W.values()} == {"Paid", "Approved", "Pending"})

# ================================================================ 2. Excel
xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(ROOT / "Sales-Commission-Calculator.xlsx"), 0, True)
    ws = {n: wb.Worksheets(n) for n in ("Dashboard", "Commission", "Statement", "What-if", "Payouts", "Orders", "Returns",
                                        "Plan", "Test cases", "Calc")}
    name_cell = lambda n: wb.Names(n).RefersToRange
    olo, rlo, llo = ws["Orders"].ListObjects("Orders"), ws["Returns"].ListObjects("Returns"), ws["Payouts"].ListObjects("Ledger")
    col = lambda lo, n: [x[0] for x in lo.ListColumns(n).DataBodyRange.Value]

    # ---- order lines
    xn, xt, xs, xr, xsc = (col(olo, n) for n in ("Net", "Tax", "Status", "Rep credit", "Split credit"))
    check(f"Every order line's net and tax match ({len(lines)} lines)",
          all(close(num(xn[i]) if xn[i] != "" else None, l["net"]) and close(num(xt[i]) if xt[i] != "" else None, l["tax"])
              for i, l in enumerate(lines)))
    check(f"Every split deal credits the right share to each rep ({sum(1 for l in lines if l['split_rep'])} split lines)",
          all(close(num(xr[i]), l["rep_cr"] or 0) and close(num(xsc[i]), l["split_cr"] or 0) for i, l in enumerate(lines)))
    check("Every order line's status matches", [xs[i] for i in range(len(lines))] == [l["status"] for l in lines])
    check("Unused order rows stay blank", all(x == "" for x in xs[len(lines):]))

    # ---- returns
    rs, rn = col(rlo, "Status"), col(rlo, "Returned net")
    check(f"Every return's status and returned value match ({len(rets)} returns)",
          [rs[i] for i in range(len(rets))] == [x["status"] for x in rets]
          and all(close(num(rn[i]) if rn[i] != "" else None, x["ret_net"]) for i, x in enumerate(rets)))

    method_cell = name_cell("Method")
    for method in ("Whole amount", "Stepped"):
        method_cell.Value = method
        xl.Calculate()
        led = py_ledger(method)   # recomputed here so the per-return clawbacks are this method's
        tag = method
        # ---- ledger: 72 rep-months
        lrep, lmon = col(llo, "Sales rep"), [from_serial(x[0]) for x in llo.ListColumns("Month").DataBodyRange.Value2]
        rows = [(lmon[i], lrep[i]) for i in range(len(lrep))]
        check(f"{tag}: ledger has one row per rep and month (72)", sorted(rows) == sorted(led.keys()))
        for field, xname in (("lines", "Lines"), ("net", "Credited net"), ("comm", "Commission"), ("accel", "Accelerator"),
                             ("bonus", "Bonus"), ("override", "Override"), ("claw", "Clawback"), ("total", "Total pay"),
                             ("to_pay", "To pay"), ("wi_pay", "What-if pay")):
            got = col(llo, xname)
            bad = [(rows[i], got[i], led[rows[i]][field]) for i in range(len(rows)) if not close(num(got[i]), led[rows[i]][field])]
            check(f"{tag}: ledger {xname.lower()} for all 72 rep-months", not bad, str(bad[:2]))
        rc1, rc2 = col(rlo, "Clawback rep"), col(rlo, "Clawback split")
        check(f"{tag}: clawback of every return (rep and split rep)",
              all(close(num(rc1[i]), x["claw_rep"]) and close(num(rc2[i]), x["claw_split"]) for i, x in enumerate(rets)))

        for month in (MONTHS[-2], MONTHS[-1]):
            name_cell("SelMonth").Value2 = serial(month)
            xl.Calculate()
            mt = f"{month:%b %Y}/{tag}"
            c = ws["Commission"]
            got = {c.Range(f"B{r}").Value: r for r in range(11, 21) if c.Range(f"B{r}").Value}
            ok = all(got.get(rep) and c.Range(f"D{got[rep]}").Value == led[(month, rep)]["lines"]
                     and all(close(c.Range(f"{cl}{got[rep]}").Value, led[(month, rep)][k])
                             for cl, k in (("E", "net"), ("I", "comm"), ("J", "accel"), ("K", "bonus"), ("L", "override"),
                                           ("M", "claw"), ("N", "total")))
                     and c.Range(f"O{got[rep]}").Value == led[(month, rep)]["status"] for rep in REPS)
            check(f"{mt}: Commission sheet, every rep (credited net, commission, accelerator, bonus, override, clawback, pay, status)", ok)
            check(f"{mt}: Commission totals", close(c.Range("N21").Value, sum(led[(month, r)]["total"] for r in REPS), 0.01))
            want = [l["order"] for l in lines if l["status"] != "OK" and l["month"] in (month, None)]
            got_h = [c.Range(f"B{r}").Value for r in range(45, 55) if c.Range(f"B{r}").Value]
            check(f"{mt}: lines held back", got_h == want, f"{got_h} vs {want}")
            want_r = [x["order"] for x in rets if x["booked"] == month]
            got_r = [c.Range(f"B{r}").Value for r in range(60, 68) if c.Range(f"B{r}").Value]
            check(f"{mt}: returns booked in the month", got_r == want_r, f"{got_r} vs {want_r}")
            checks = [c.Range(f"N{r}").Value for r in range(71, 79)]
            check(f"{mt}: the 8 data checks say PASS", checks == ["PASS"] * 8, str(checks))

            # ---- Dashboard KPIs, leaderboard, donut, insights
            d = ws["Dashboard"]
            prev = MONTHS[MONTHS.index(month) - 1]
            sales = sum(led[(month, r)]["net"] for r in REPS)
            cost = sum(led[(month, r)]["total"] for r in REPS)
            on_t = sum(1 for r in REPS if led[(month, r)]["net"] >= target[r])
            n_held = len(want)
            to_pay = sum(e["to_pay"] for e in led.values())
            ok = (close(d.Range("B11").Value, sales, 0.01) and close(d.Range("H11").Value, cost, 0.01)
                  and close(d.Range("N11").Value, cost / sales, 1e-9) and d.Range("T11").Value == f"{on_t} of {len(REPS)}"
                  and d.Range("Z11").Value == n_held and close(d.Range("AF11").Value, to_pay, 0.01))
            check(f"{mt}: Dashboard KPI cards (sales, cost, cost %, reps on target, held back, to pay)", ok,
                  str([d.Range(a).Value for a in ("B11", "H11", "N11", "T11", "Z11", "AF11")]))
            psales = sum(led[(prev, r)]["net"] for r in REPS)
            check(f"{mt}: Dashboard sales note vs prior month",
                  d.Range("B12").Text.endswith(f"vs {prev:%b}") and (("▲" in d.Range("B12").Text) == (sales > psales)))
            att = sorted(REPS, key=lambda r: (-led[(month, r)]["net"] / target[r], REPS.index(r)))
            board = [d.Range(f"D{r}").Value for r in range(38, 44)]
            check(f"{mt}: leaderboard ranked by % of target, with pay and payout status",
                  board == att and all(close(d.Range(f"N{38 + i}").Value, led[(month, rep)]["total"]) and
                                       d.Range(f"P{38 + i}").Value == led[(month, rep)]["status"] for i, rep in enumerate(att)),
                  f"{board} vs {att}")
            donut = [ws["Calc"].Range(f"B{91 + i}").Value for i in range(len(REPS))]
            check(f"{mt}: donut = each rep's share of the month's pay",
                  all(close(donut[i], max(0, led[(month, r)]["total"])) for i, r in enumerate(REPS)))
            top = max(REPS, key=lambda r: led[(month, r)]["total"])
            check(f"{mt}: insights name the top earner", d.Range("V37").Value.startswith(f"Top earner: {top}, "),
                  d.Range("V37").Value)

            # ---- statements for every rep in this month
            s = ws["Statement"]
            stmt_ok = []
            for rep in REPS:
                name_cell("StmtRep").Value = rep
                name_cell("StmtMonth").Value2 = serial(month)
                xl.Calculate()
                e = led[(month, rep)]
                listed = [x[0] for x in s.Range("I47:I146").Value if x[0] not in ("", None)]
                bands_sum = sum(num(x[0]) for x in s.Range("I27:I32").Value)
                ytd = [led[(m, rep)] for m in MONTHS if m.year == month.year and m <= month]
                ok = (close(s.Range("D20").Value, e["net"]) and close(s.Range("I33").Value, e["comm"])
                      and close(s.Range("I34").Value, e["accel"]) and close(s.Range("I35").Value, e["bonus"])
                      and close(s.Range("I36").Value, e["override"]) and close(s.Range("I37").Value, -e["claw"])
                      and close(s.Range("I38").Value, e["total"]) and s.Range("D23").Value == e["status"]
                      and close(bands_sum, e["comm"], 0.011) and len(listed) == e["lines"] and close(sum(listed), e["net"], 0.01)
                      and s.Range("I148").Value == "PASS" and s.Range("B39").Value.strip().startswith("Checked")
                      and close(s.Range("B43").Value, sum(x["net"] for x in ytd), 0.01)
                      and close(s.Range("D43").Value, sum(x["total"] for x in ytd), 0.01)
                      and close(s.Range("E43").Value, sum(x["total"] for x in ytd if x["status"] == "Paid"), 0.01)
                      and close(s.Range("F43").Value, sum(x["to_pay"] for x in ytd), 0.01))
                stmt_ok.append(ok)
            check(f"{mt}: statement of each of the {len(REPS)} reps (receipt lines, total, YTD, every credited line)",
                  all(stmt_ok), str(stmt_ok))

        # ---- Dashboard trend (12 months x 6 reps) and What-if, for this method
        k = ws["Calc"]
        trend = k.Range("B3:G14").Value
        check(f"{tag}: 12-month pay trend per rep", all(close(trend[i][j], led[(m, rep)]["total"], 0.01)
                                                         for i, m in enumerate(MONTHS) for j, rep in enumerate(REPS)))
        w = ws["What-if"]
        now = sum(e["total"] for e in led.values()); wi = sum(e["wi_pay"] for e in led.values())
        per_rep = all(close(w.Range(f"I{17 + i}").Value, sum(led[(m, rep)]["total"] for m in MONTHS), 0.01)
                      and close(w.Range(f"J{17 + i}").Value, sum(led[(m, rep)]["wi_pay"] for m in MONTHS), 0.01) for i, rep in enumerate(REPS))
        check(f"{tag}: What-if cost now vs scenario (cards and per rep)",
              per_rep and close(w.Range("B10").Value, now, 0.01) and close(w.Range("D10").Value, wi, 0.01)
              and close(w.Range("H10").Value, wi - now, 0.01))

    # ---- What-if: a scenario equal to the current plan must cost exactly the same
    method_cell.Value = "Whole amount"
    sb = ws["What-if"].ListObjects("ScBands")
    saved = {n: sb.ListColumns(n).DataBodyRange.Value for n in ("From", "Rate")}
    sb.ListColumns("From").DataBodyRange.Value = [(b[0],) for b in PLAN["bands"]]
    sb.ListColumns("Rate").DataBodyRange.Value = [(b[1],) for b in PLAN["bands"]]
    sc_cells = {n: name_cell(n) for n in ("ScTargetChange", "ScAccelAt", "ScAccelRate", "ScBonus")}
    old = {n: c.Value for n, c in sc_cells.items()}
    for n, v in (("ScTargetChange", 0), ("ScAccelAt", PLAN["accel_at"]), ("ScAccelRate", PLAN["accel_rate"]), ("ScBonus", PLAN["target_bonus"])):
        sc_cells[n].Value = v
    xl.Calculate()
    check("What-if: a scenario equal to the current plan shows a difference of 0.00",
          close(ws["What-if"].Range("H10").Value, 0, 0.001) and close(ws["What-if"].Range("K23").Value, 0, 0.001))
    # and a hand-made scenario (all rates +1 point) recomputed in Python
    plus = [(f, r + 0.01) for f, r in PLAN["bands"]]
    sb.ListColumns("Rate").DataBodyRange.Value = [(b[1],) for b in plus]
    xl.Calculate()
    led_plus = py_ledger("Whole amount", sc=dict(SC, bands=plus, target_change=0, accel_at=PLAN["accel_at"],
                                                 accel_rate=PLAN["accel_rate"], target_bonus=PLAN["target_bonus"]))
    check("What-if: every band rate +1 point gives the cost Python expects",
          close(ws["What-if"].Range("D10").Value, sum(e["wi_pay"] for e in led_plus.values()), 0.01))
    for n, v in saved.items():
        sb.ListColumns(n).DataBodyRange.Value = v
    for n, c in sc_cells.items():
        c.Value = old[n]
    xl.Calculate()

    # ---- Payouts
    p = ws["Payouts"]
    approved = sum(e["to_pay"] for e in W.values() if e["status"] == "Approved")
    pending = sum(e["to_pay"] for e in W.values() if e["status"] == "Pending")
    paid = sum(e["total"] for e in W.values() if e["status"] == "Paid")
    check("Payouts: to pay (approved), waiting for approval, paid",
          close(p.Range("B9").Value, approved, 0.01) and close(p.Range("D9").Value, pending, 0.01) and close(p.Range("G9").Value, paid, 0.01),
          str([p.Range(a).Value for a in ("B9", "D9", "G9")]))
    check("Payouts: status matrix shows each rep's pay for every month",
          all(close(p.Cells(15 + i, 3 + j).Value, W[(m, rep)]["total"], 0.01) for i, rep in enumerate(REPS) for j, m in enumerate(MONTHS)))
    ck = col(llo, "Check")
    check("Payouts: every Paid month has a paid date (Check column all OK)", ck == ["OK"] * len(ck))

    # ---- 3. tests, protection, validation, compatibility
    t = ws["Test cases"]
    res = []
    for lo_name in ("OrderTests", "ReturnTests"):
        res += [x[0] for x in t.ListObjects(lo_name).ListColumns("Result").DataBodyRange.Value]
    ur = t.UsedRange.Value
    total_cases = int(name_cell("TestsTotal").Value)
    all_res = [v for row in ur for v in row if v in ("PASS", "FAIL")]
    check(f"Test cases sheet: all {total_cases} hand-worked cases PASS", name_cell("TestsFailed").Value == 0
          and all_res.count("PASS") == total_cases and "FAIL" not in all_res and total_cases >= 50, f"{all_res.count('PASS')}/{total_cases}")
    check("Formula sheets are protected", all(ws[n].ProtectContents for n in ("Dashboard", "Commission", "Statement", "What-if",
                                                                                "Payouts", "Orders", "Returns", "Test cases")))
    o = ws["Orders"]
    first, last = olo.DataBodyRange.Row, olo.DataBodyRange.Row + olo.DataBodyRange.Rows.Count - 1
    check("Order inputs unlocked (incl. split rep and %), calculated columns locked",
          not o.Range(f"A{first}").Locked and not o.Range(f"I{last}").Locked and o.Range(f"J{first}").Locked and o.Range(f"S{last}").Locked)
    vt = {n: olo.ListColumns(n).DataBodyRange.Cells(1, 1).Validation.Type for n in ("Date", "Sales rep", "Product", "Qty", "Discount %", "Split rep", "Split %")}
    check("Drop-downs and validation on order inputs", vt == {"Date": 4, "Sales rep": 3, "Product": 3, "Qty": 1, "Discount %": 2,
                                                               "Split rep": 3, "Split %": 2}, str(vt))
    check("Payout Status / Paid date and What-if inputs are unlocked, payout status is a drop-down",
          not llo.ListColumns("Status").DataBodyRange.Cells(1, 1).Locked and not llo.ListColumns("Paid date").DataBodyRange.Cells(1, 1).Locked
          and llo.ListColumns("Status").DataBodyRange.Cells(1, 1).Validation.Type == 3
          and not name_cell("ScTargetChange").Locked and not ws["What-if"].ListObjects("ScBands").DataBodyRange.Cells(1, 2).Locked)
    dd = [wb.Names(n).RefersToRange.Rows.Count for n in ("RepList", "ProductList", "MonthList", "MethodList", "StatusList")]
    check("Drop-down lists (reps 6, products 12, months 12, methods 2, payout statuses 3)", dd == [6, 12, 12, 2, 3], str(dd))
    check("Month and rep selectors are unlocked", not name_cell("SelMonth").Locked and not name_cell("StmtRep").Locked and not name_cell("StmtMonth").Locked)
    nav_ok = all(any(sh.Name.startswith("nav_") for sh in ws[n].Shapes) for n in ("Dashboard", "Commission", "Statement", "What-if", "Payouts",
                                                                                  "Orders", "Returns", "Plan", "Test cases"))
    check("Navigation buttons on every sheet", nav_ok)
    implicit, newfn = [], []
    modern = ("FILTER(", "XLOOKUP(", "LET(", "SORT(", "SORTBY(", "UNIQUE(", "SEQUENCE(", "IFS(", "MAXIFS(", "MINIFS(", "TEXTJOIN(", "LAMBDA(")
    for sh in wb.Worksheets:
        try:
            cells = sh.UsedRange.SpecialCells(-4123)
        except Exception:
            continue
        for a in cells.Areas:
            f2 = a.Formula2
            for row in (f2 if isinstance(f2, tuple) else ((f2,),)):
                for x in row:
                    if not isinstance(x, str):
                        continue
                    if "@" in x.replace("[@", ""):
                        implicit.append(f"{sh.Name}: {x[:50]}")
                    if any(m in x.upper().replace("_XLFN.", "") for m in modern):
                        newfn.append(f"{sh.Name}: {x[:50]}")
    check("No implicit-intersection @ formulas (opens cleanly in Excel 2016)", not implicit, str(implicit[:3]))
    check("No Microsoft 365-only functions (FILTER, XLOOKUP, LET, SORT, IFS, ...)", not newfn, str(newfn[:3]))
    errs = []
    for sh in wb.Worksheets:
        try:
            errs.append(f"{sh.Name}!{sh.UsedRange.SpecialCells(-4123, 16).Address}")
        except Exception:
            pass
    check("No error values anywhere", not errs, str(errs))
    wb.Close(False)
finally:
    xl.Quit()

# summary numbers for the README (current plan, last month)
m = MONTHS[-1]
for method in ("Whole amount", "Stepped"):
    L = LED[method]
    print(f"\n{m:%b %Y} {method}: sales {sum(L[(m, r)]['net'] for r in REPS):,.2f}  commission {sum(L[(m, r)]['comm'] for r in REPS):,.2f}  "
          f"accel {sum(L[(m, r)]['accel'] for r in REPS):,.2f}  bonus {sum(L[(m, r)]['bonus'] for r in REPS):,.2f}  "
          f"override {sum(L[(m, r)]['override'] for r in REPS):,.2f}  clawback {sum(L[(m, r)]['claw'] for r in REPS):,.2f}  "
          f"total {sum(L[(m, r)]['total'] for r in REPS):,.2f}  | 12 months {sum(e['total'] for e in L.values()):,.2f}  "
          f"what-if {sum(e['wi_pay'] for e in L.values()):,.2f}  to pay {sum(e['to_pay'] for e in L.values()):,.2f}")
print(f"{len(lines)} order lines, {sum(1 for l in lines if l['split_rep'])} split, {len(rets)} returns")
print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
