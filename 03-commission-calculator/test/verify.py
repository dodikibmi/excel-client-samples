"""Independent check of Sales-Commission-Calculator.xlsx.

1. Recomputes every order line (price, discount, net, tax, status) and every rep's commission, for both months and
   both commission methods, in plain Python from data/orders.csv and data/plan.json (no Excel involved).
2. Opens the workbook in Excel (read-only), switches month and method, and compares the Commission sheet,
   the held-back list, and the Statement for every rep.
3. Checks the planted mistakes are exactly the lines held back, all test cases and data checks say PASS,
   input cells are unlocked and validated, formula sheets are protected, and no formula needs Microsoft 365.
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
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


def r2(x):
    """Round to cents the way Excel does (15 significant digits first, then half away from zero)."""
    return float(Decimal(f"{x:.15g}").quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


close = lambda a, b, tol=0.005: abs((a or 0) - (b or 0)) <= tol
price = {p[0]: p[2] for p in PLAN["products"]}
maxd = {p[0]: p[3] for p in PLAN["products"]}
target = {r[0]: r[2] for r in PLAN["reps"]}
bands = PLAN["bands"]


def commission(net, method):
    if method == "Stepped":
        total, prev = 0.0, 0.0
        for frm, rate in bands:
            if net > frm:
                total += (net - frm) * (rate - prev)
            prev = rate
        return r2(total)
    rate = [rt for frm, rt in bands if net >= frm][-1]
    return r2(net * rate)


# ---------------------------------------------------------------- 1. plain Python
lines = []
with open(ROOT / "data" / "orders.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        d = date.fromisoformat(r["Date"]) if r["Date"] else None
        qty, disc, prod, rep = int(r["Qty"]), float(r["Discount %"]), r["Product"], r["Sales rep"]
        net = None
        if prod in price:
            gross = r2(qty * price[prod]); net = r2(gross - r2(gross * disc))
        status = ("Missing date" if not d else "Unknown sales rep" if rep not in target else "Unknown product" if prod not in price
                  else "Qty must be 1 or more" if qty < 1 else f"Discount above {maxd[prod]:.0%} max" if disc > maxd[prod] + 1e-6 else "OK")
        lines.append(dict(order=r["Order no"], month=d.replace(day=1) if d else None, rep=rep, prod=prod, qty=qty,
                          net=net, tax=r2(net * PLAN["tax_rate"]) if net is not None else None, status=status))

held = sorted((l["order"], l["status"]) for l in lines if l["status"] != "OK")
check("Lines held back = the planted mistakes, with the right reason", held == sorted((p["order"], p["why"]) for p in PLANTED), str(held))


def py_month(month, method):
    reps = {}
    for rep, region, tgt in PLAN["reps"]:
        ok = [l for l in lines if l["rep"] == rep and l["month"] == month and l["status"] == "OK"]
        net = r2(sum(l["net"] for l in ok))
        reps[rep] = dict(lines=len(ok), net=net, comm=commission(net, method), bonus=PLAN["target_bonus"] if net >= tgt else 0)
    return reps


# ---------------------------------------------------------------- 2. Excel
xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(ROOT / "Sales-Commission-Calculator.xlsx"), 0, True)
    c, s, o, p, t = (wb.Worksheets(n) for n in ("Commission", "Statement", "Orders", "Plan", "Test cases"))
    lo = o.ListObjects("Orders")
    col = lambda n: [x[0] for x in lo.ListColumns(n).DataBodyRange.Value]
    xl_net, xl_tax, xl_status = col("Net"), col("Tax"), col("Status")
    check("Every order line's net and tax match", all(close(xl_net[i] if xl_net[i] != "" else None, l["net"]) and
                                                     close(xl_tax[i] if xl_tax[i] != "" else None, l["tax"]) for i, l in enumerate(lines)))
    check("Every order line's status matches", [xl_status[i] for i in range(len(lines))] == [l["status"] for l in lines])
    check("Unused table rows stay blank", all(x == "" for x in xl_status[len(lines):]))

    for month in (date(2026, 9, 1), date(2026, 10, 1)):
        for method in ("Whole amount", "Stepped"):
            c.Range("C4").Value2 = (month - date(1899, 12, 30)).days
            p.Range("B9").Value = method
            xl.CalculateFull()
            exp = py_month(month, method)
            tag = f"{month:%b %Y}/{method}"
            rows = {c.Range(f"B{r}").Value: r for r in range(14, 24) if c.Range(f"B{r}").Value}
            ok = all(rows.get(rep) and c.Range(f"D{rows[rep]}").Value == e["lines"] and close(c.Range(f"E{rows[rep]}").Value, e["net"])
                     and close(c.Range(f"I{rows[rep]}").Value, e["comm"]) and close(c.Range(f"J{rows[rep]}").Value, e["bonus"])
                     and close(c.Range(f"K{rows[rep]}").Value, e["comm"] + e["bonus"]) for rep, e in exp.items())
            check(f"{tag}: lines, net sales, commission, bonus and pay per rep", ok,
                  str({rep: (c.Range(f"I{r}").Value, exp[rep]['comm']) for rep, r in rows.items()}))
            check(f"{tag}: totals", close(c.Range("K24").Value, sum(e["comm"] + e["bonus"] for e in exp.values())))
            want_held = [l["order"] for l in lines if l["status"] != "OK" and (l["month"] in (month, None))]
            got_held = [c.Range(f"B{r}").Value for r in range(45, 55) if c.Range(f"B{r}").Value]
            check(f"{tag}: held-back list", got_held == want_held, f"{got_held} vs {want_held}")
            check(f"{tag}: KPI lines held back", c.Range("J8").Value == len(want_held))
            check(f"{tag}: data checks all PASS", [c.Range(f"K{r}").Value for r in range(58, 62)] == ["PASS"] * 4,
                  str([c.Range(f"K{r}").Value for r in range(58, 62)]))
            # statements for every rep in this month/method
            stmt_ok = []
            for rep in exp:
                s.Range("C4").Value = rep
                s.Range("G4").Value2 = (month - date(1899, 12, 30)).days
                xl.Calculate()
                e = exp[rep]
                listed = [x[0] for x in s.Range("I19:I98").Value if x[0] not in ("", None)]
                stmt_ok.append(close(s.Range("D8").Value, e["net"]) and close(s.Range("D11").Value, e["comm"]) and
                               close(s.Range("D12").Value, e["bonus"]) and len(listed) == e["lines"] and
                               close(sum(listed), e["net"], 0.01) and s.Range("I100").Value == "PASS")
            check(f"{tag}: statement for each of the {len(exp)} reps", all(stmt_ok), str(stmt_ok))
    p.Range("B9").Value = "Whole amount"; xl.CalculateFull()

    # ---------------------------------------------------------------- 3. tests, protection, validation, compatibility
    res = [x[0] for x in t.Range("J6:J14").Value] + [x[0] for x in t.ListObjects("OrderTests").ListColumns("Result").DataBodyRange.Value]
    check(f"Test cases sheet: all {len(res)} cases PASS", res == ["PASS"] * len(res), str(res))
    check("Formula sheets are protected", all(wb.Worksheets(n).ProtectContents for n in ("Commission", "Statement", "Orders", "Test cases")))
    check("Order inputs unlocked, calculated columns locked",
          not o.Range("A5").Locked and not o.Range("G604").Locked and o.Range("H5").Locked and o.Range("O604").Locked)
    vt = {n: lo.ListColumns(n).DataBodyRange.Cells(1, 1).Validation.Type for n in ("Date", "Sales rep", "Product", "Qty", "Discount %")}
    check("Drop-downs and validation on order inputs", vt == {"Date": 4, "Sales rep": 3, "Product": 3, "Qty": 1, "Discount %": 2}, str(vt))
    dd = [wb.Names(n).RefersToRange.Rows.Count for n in ("RepList", "ProductList", "MonthList", "MethodList")]
    check("Drop-down lists read ranges (reps 6, products 12, months 24, methods 2)", dd == [6, 12, 24, 2], str(dd))
    check("Month and rep selectors are unlocked", not c.Range("C4").Locked and not s.Range("C4").Locked and not s.Range("G4").Locked)
    implicit = []
    for ws in wb.Worksheets:
        try:
            cells = ws.UsedRange.SpecialCells(-4123)
        except Exception:
            continue
        implicit += [f"{ws.Name}!{x.Address}" for a in cells.Areas for x in a.Cells if "@" in x.Formula2.replace("[@", "")]
    check("No implicit-intersection @ formulas (opens cleanly in Excel 2016)", not implicit, str(implicit[:3]))
    errs = []
    for ws in wb.Worksheets:
        try:
            errs.append(f"{ws.Name}!{ws.UsedRange.SpecialCells(-4123, 16).Address}")
        except Exception:
            pass
    check("No error values anywhere", not errs, str(errs))
    wb.Close(False)
finally:
    xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
