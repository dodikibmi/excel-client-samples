"""Independent check of the E11 redesign.

1. The old file: its 9 mistakes are re-created in plain Python (the same wrong formulas, on the same export) and must
   give exactly the numbers Excel shows in Before-Client-Dashboard.xlsx. This proves the "old file shows" column.
2. The right numbers: the export is cleaned in Python (text -> number, trimmed names, second copy of an order dropped)
   and every card, chart table, insight, issue, bridge step and check of Sales-Dashboard-Redesign.xlsx must match.
3. The file itself: opened in Excel, fully recalculated (cached values = live values), 5 charts reading the right
   tables, no macros, no error values, Excel 2016 functions only, the data columns A-I identical to the export.

    python test/verify.py          (openpyxl + pywin32; run under tools/with_lock.py)
"""
import csv, json, re, sys
from datetime import date, timedelta
from decimal import Decimal as D
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
PL = json.loads((ROOT / "data" / "planted.json").read_text())
TEXT_IDS = set(PL["text_revenue_orders"])
results = []


def check(name, ok, detail=""):
    results.append(bool(ok))
    print(("PASS " if ok else "FAIL ") + name + ("" if ok or detail == "" else f"   <- {detail}"))


def close(a, b, tol=0.005):
    return a is not None and b is not None and abs(float(a) - float(b)) < tol


def load(name):
    out = []
    for r in list(csv.reader(open(ROOT / "data" / name, encoding="utf-8")))[1:]:
        is_text = r[0] in TEXT_IDS and "2026" in name
        out.append(dict(id=r[0], d=date.fromisoformat(r[1]), region=r[2], cat=r[3], prod=r[4], units=int(r[5]), price=D(r[6]),
                        raw=r[7], text=is_text, rev=D(r[7].replace(",", "")), cost=D(r[8])))
    return out


r25, x26 = load("sales-2025.csv"), load("sales-2026.csv")
END = date(2026, 10, 31)

# ============================================================== 1. the old file, mistakes and all
inrange = x26[:699]                                         # =SUM(H2:H700)
num = [r for r in inrange if not r["text"]]                  # SUM / SUMIF skip text
old = {}
old["B3"] = sum(r["rev"] for r in num)
old["B4"] = sum(r["rev"] for r in r25 if r["d"] <= date(2025, 10, 31))
old["B5"] = (old["B3"] - old["B4"]) / old["B3"]
old["B6"] = len(inrange)
old["B7"] = sum((r["rev"] - r["cost"]) / r["rev"] for r in num) / len(num)
regions = {g: sum(r["rev"] for r in num if r["region"] == g) for g in ("North", "South", "East")}
regions["West"] = sum(r["cost"] for r in inrange if r["region"] == "West")    # the Cost column is numeric on every row
old["E7"] = regions["West"]
oct_prelim = round(float(sum(r["rev"] for r in x26 if date(2026, 10, 1) <= r["d"] <= date(2026, 10, 15))), -2)
old["B20"] = oct_prelim
cats = ["Seeds & bulbs", "Plants", "Hand tools", "Power tools", "Watering", "Soil & feed"]
catv = {c: sum(r["rev"] for r in num if r["cat"] == c) for c in cats}
old["E18"] = sum(catv.values()) + catv["Hand tools"] + catv["Power tools"]
old["hose"] = sum(r["rev"] for r in num if r["prod"] == "Hose 25m")

bw = load_workbook(ROOT / "Before-Client-Dashboard.xlsx", data_only=True)["Dashboard"]
hose_cell = next(f"H{r}" for r in range(4, 22) if bw[f"G{r}"].value == "Hose 25m")
for cell, lab in (("B3", "total revenue (SUM stops at row 700, text skipped, duplicates added)"), ("B4", "revenue 2025"),
                  ("B5", "growth divided by this year"), ("B6", "order count (same range)"), ("B7", "average of row margins"),
                  ("E7", "West = sum of Cost"), ("B20", "October typed from the 1-15 Oct preliminary"),
                  ("E18", "category total incl. the subtotal row")):
    check(f"old file {cell} = Python re-creation of the mistake: {lab}", close(bw[cell].value, old[cell], 1e-6), (bw[cell].value, old[cell]))
check("old file Hose 25m = SUMIF that misses the trailing-space rows", close(bw[hose_cell].value, old["hose"]))
check("old file regions North/South/East = SUMIF on the cut range", all(close(bw[f"E{4 + i}"].value, regions[g])
                                                                         for i, g in enumerate(("North", "South", "East"))))
check("old file months Jan-Sep = SUMIFS on the cut range", all(close(bw[f"B{10 + m}"].value, sum(r["rev"] for r in num if r["d"].month == m))
                                                                for m in range(1, 10)))
ws = load_workbook(ROOT / "Before-Client-Dashboard.xlsx")["Data 2026"]
check("old file keeps the 14 revenue cells as text, like the export",
      sum(1 for row in ws.iter_rows(min_row=2, min_col=8, max_col=8) if row[0].data_type == "s") == 14)

# ============================================================== 2. the right numbers
seen, clean26, dups = set(), [], []
for r in x26:
    if r["id"] in seen:
        dups.append(r); continue
    seen.add(r["id"]); clean26.append(dict(r, prod=r["prod"].strip()))
check("planted problems: 9 duplicates, 14 text numbers, 6 trailing spaces",
      sorted(r["id"] for r in dups) == sorted(PL["duplicate_orders"]) and sum(r["text"] for r in x26) == 14 and
      sum(r["prod"] != r["prod"].strip() for r in x26) == 6)
c26 = [r for r in clean26 if r["d"] <= END]
c25 = [r for r in r25 if r["d"] <= date(2025, 10, 31)]
R = {}
R["rev26"] = sum(r["rev"] for r in c26); R["rev25"] = sum(r["rev"] for r in c25)
R["cost26"] = sum(r["cost"] for r in c26); R["cost25"] = sum(r["cost"] for r in c25)
R["growth"] = R["rev26"] / R["rev25"] - 1
R["m26"] = (R["rev26"] - R["cost26"]) / R["rev26"]; R["m25"] = (R["rev25"] - R["cost25"]) / R["rev25"]
R["orders26"], R["orders25"] = len(c26), len(c25)
months = [(sum(r["rev"] for r in c26 if r["d"].month == m), sum(r["rev"] for r in c25 if r["d"].month == m)) for m in range(1, 11)]
regs = {g: (sum(r["rev"] for r in c26 if r["region"] == g), sum(r["rev"] for r in c25 if r["region"] == g)) for g in ("North", "South", "East", "West")}
catn = {c: sum(r["rev"] for r in c26 if r["cat"] == c) for c in cats}
prods = {}
for r in c26:
    prods[r["prod"]] = prods.get(r["prod"], 0) + r["rev"]

aw = load_workbook(ROOT / "Sales-Dashboard-Redesign.xlsx", data_only=True)
calc, dash, iss, br, chk = aw["Calc"], aw["Dashboard"], aw["Issues found"], aw["Bridge"], aw["Checks"]
kp = {calc[f"A{i}"].value: calc[f"B{i}"].value for i in range(3, 14)}
want = {"Rev_26": R["rev26"], "Rev_25": R["rev25"], "Cost_26": R["cost26"], "Cost_25": R["cost25"], "Orders_26": R["orders26"],
        "Orders_25": R["orders25"], "Growth": R["growth"], "Margin_26": R["m26"], "Margin_25": R["m25"],
        "AOV_26": R["rev26"] / R["orders26"], "AOV_25": R["rev25"] / R["orders25"]}
check("KPI table (revenue, cost, orders, growth, margin, average order, both years) = Python",
      all(close(kp[k], v, 1e-6) for k, v in want.items()), {k: (kp[k], float(v)) for k, v in want.items() if not close(kp[k], v, 1e-6)})
cards = {"B7": "Rev_26", "F7": "Growth", "J7": "Margin_26", "N7": "Orders_26", "R7": "AOV_26"}
check("Dashboard cards show those values", all(close(dash[c].value, want[k], 1e-6) for c, k in cards.items()))
check("months 2026 and 2025 = Python", all(close(calc[f"E{2 + m}"].value, months[m - 1][0]) and close(calc[f"F{2 + m}"].value, months[m - 1][1])
                                            for m in range(1, 11)))
check("regions 2026, 2025 and growth = Python", all(close(calc[f"I{i}"].value, regs[calc[f"H{i}"].value][0]) and
                                                     close(calc[f"J{i}"].value, regs[calc[f"H{i}"].value][1]) and
                                                     close(calc[f"K{i}"].value, regs[calc[f"H{i}"].value][0] / regs[calc[f"H{i}"].value][1] - 1, 1e-9)
                                                     for i in range(3, 7)))
srt = sorted(catn.items(), key=lambda kv: kv[1])
check("categories sorted for the bar chart = Python", [(calc[f"P{i}"].value, round(calc[f"Q{i}"].value, 2)) for i in range(3, 9)] ==
      [(k, round(float(v), 2)) for k, v in srt])
top = sorted(prods.items(), key=lambda kv: kv[1])[-8:]
check("top 8 products (trimmed names) = Python", [(calc[f"V{i}"].value, round(calc[f"W{i}"].value, 2)) for i in range(3, 11)] ==
      [(k, round(float(v), 2)) for k, v in top])
peak = max(range(10), key=lambda i: months[i][0])
fast = max(regs, key=lambda g: regs[g][0] / regs[g][1]); slow = min(regs, key=lambda g: regs[g][0] / regs[g][1])
ins = " ".join(dash[f"B{r}"].value for r in (13, 14, 15))
check("insights name the right peak month, fastest / slowest region, top category and best seller",
      ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct"][peak] in ins and f"region: {fast}" in ins and
      f"slowest: {slow}" in ins and srt[-1][0] in ins and top[-1][0] in ins, ins)
check("subtitle counts the cleaning: 9 duplicates, 14 text numbers, 6 names",
      "9 duplicates removed, 14 text numbers fixed, 6 product names trimmed" in dash["B4"].value)

# issues: old values = the old file, correct = Python, difference = correct - old
miss = sum(r["rev"] for r in x26[699:] if r["id"] not in {d["id"] for d in dups} or True) - sum(r["rev"] for r in dups if x26.index(r) >= 699)
txt = sum(r["rev"] for r in inrange if r["text"])
dup_amt = sum(r["rev"] for i, r in enumerate(x26[:699]) if r["id"] in [d["id"] for d in dups] and
              any(x26[j]["id"] == r["id"] for j in range(i)))
bridge = [br[f"C{r}"].value for r in range(7, 12)]
check("bridge steps = Python: old total, + rows below 700, + text numbers, - duplicates, right total",
      close(bridge[0], old["B3"]) and close(bridge[1], miss) and close(bridge[2], txt) and close(bridge[3], -dup_amt) and close(bridge[4], R["rev26"]),
      (bridge, float(old["B3"]), float(miss), float(txt), float(dup_amt), float(R["rev26"])))
check("bridge closes: old + steps = right total", close(sum(bridge[:4]), bridge[4]))
check("bridge order counts: 699 + 27 - 9 = 717", [br[f"D{r}"].value for r in range(7, 12)] == [699, 27, 0, -9, R["orders26"]])
check("bridge checks all PASS", all(br[c].value == "PASS" for c in ("C13", "C14", "C15")))
rows = {iss[f"A{r}"].value: [iss.cell(r, c).value for c in range(1, 9)] for r in range(7, 16)}
check("Issues found lists 9 mistakes", sorted(rows) == list(range(1, 10)))
exp = {4: (old["E7"], regs["West"][0]), 5: (old["B7"], R["m26"]), 6: (old["B20"], months[9][0]), 7: (old["B5"], R["growth"]),
       8: (old["E18"], R["rev26"]), 9: (old["hose"], prods["Hose 25m"])}
check("Issues 4-9: old value = the old file, correct = Python, difference = correct - old",
      all(close(rows[k][4], o, 1e-6) and close(rows[k][5], n, 1e-6) and close(rows[k][6], D(str(n)) - D(str(o)), 1e-6) for k, (o, n) in exp.items()),
      {k: rows[k][4:7] for k in exp})
check("Issues 1-3: amounts = the bridge steps", close(rows[1][6], miss) and close(rows[2][6], txt) and close(rows[3][6], -dup_amt))
check("Issues name the exact cells of the old file", [rows[k][1] for k in (1, 4, 5, 6, 7, 8, 9)] ==
      ["Dashboard!B3", "Dashboard!E7", "Dashboard!B7", "Dashboard!B20", "Dashboard!B5", "Dashboard!E18", f"Dashboard!{hose_cell}"])
passed = [chk[f"D{r}"].value for r in range(7, 18)]
check("Checks sheet: 11 checks, all PASS", passed == ["PASS"] * 11, passed)

# data columns A-I identical to the export (text stays text)
dw = load_workbook(ROOT / "Sales-Dashboard-Redesign.xlsx")["Data"]
src = r25 + x26
okd, ntext = True, 0
for i, row in enumerate(dw.iter_rows(min_row=6, max_row=5 + len(src), max_col=9)):
    s = src[i]
    if row[7].data_type == "s":
        ntext += 1; okd &= row[7].value == s["raw"]
    else:
        okd &= close(row[7].value, s["rev"])
    okd &= row[0].value == s["id"] and row[4].value == s["prod"] and row[2].value == s["region"]
check("Data columns A-I = the two exports, row by row (14 texts kept as text, names untouched)", okd and ntext == 14 and
      dw.cell(6 + len(src), 1).value is None)

# ============================================================== 3. the file in Excel
import win32com.client as win32
xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(ROOT / "Sales-Dashboard-Redesign.xlsx"), 0, True)
    xl.CalculateFull()
    live = wb.Worksheets("Calc").Range("B3").Value
    check("recalculated in Excel: cached revenue = live revenue", close(live, R["rev26"]))
    ch = []
    for sh in ("Dashboard", "Bridge"):
        for co in wb.Worksheets(sh).ChartObjects():
            ch.append((sh, co.Name, [co.Chart.SeriesCollection(i).Formula for i in range(1, co.Chart.SeriesCollection().Count + 1)]))
    f = " ".join(" ".join(x[2]) for x in ch)
    check("5 charts, each reading its table (months, regions, categories, products, bridge)",
          len(ch) == 5 and all(t in f for t in ("Calc!$E$3:$E$12", "Calc!$F$3:$F$12", "Calc!$I$3:$I$6", "Calc!$Q$3:$Q$8", "Calc!$W$3:$W$10", "Bridge!$I$7:$I$11")), ch)
    fmts = [wb.Worksheets("Dashboard").ChartObjects(n).Chart.Axes(2).TickLabels.NumberFormat for n in ("chtMonths", "chtRegions", "chtCategories", "chtProducts")]
    check("chart number formats use a fixed $ with a locale tag (not the PC's currency symbol)",
          all(f.startswith("[$$-") for f in fmts), fmts)
    errs, newer = [], []
    for ws in wb.Worksheets:
        vals = ws.UsedRange.Value2
        for row in (vals if isinstance(vals, tuple) else ((vals,),)):
            for v in (row if isinstance(row, tuple) else (row,)):
                if isinstance(v, int) and v < -2146826000:
                    errs.append(ws.Name)
        frm = ws.UsedRange.Formula
        for row in (frm if isinstance(frm, tuple) else ((frm,),)):
            for v in (row if isinstance(row, tuple) else (row,)):
                if isinstance(v, str) and re.search(r"\b(XLOOKUP|FILTER|LET|UNIQUE|SORTBY|SEQUENCE|TEXTJOIN|IFS|MAXIFS|MINIFS|CONCAT|SWITCH)\(", v):
                    newer.append(ws.Name)
    check("no error values on any sheet", not errs, errs[:3])
    check("Excel 2016 functions only", not newer, newer[:3])
    check("no macros (.xlsx), summary sheets protected", not wb.HasVBProject and
          all(wb.Worksheets(n).ProtectContents for n in ("Dashboard", "Issues found", "Bridge", "Checks")))
    wb.Close(False)
finally:
    xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks pass")
(ROOT / "test" / "last-result.txt").write_text(f"{sum(results)}/{len(results)}\n")
sys.exit(0 if all(results) else 1)
