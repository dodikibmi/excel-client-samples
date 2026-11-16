"""Independent check of Three-Statement-Model.xlsx.

1. Reads the assumptions back from the workbook and runs the whole model again in plain Python (test/model.py).
2. Opens a copy in Excel, switches to each scenario and compares every line of the 3 statements and the schedules,
   36 months each; the annual sheet, the key figures, the scenario table and both sensitivity tables (50 cells).
3. Breaks the copy on purpose (a typed number, a broken balance sheet link) to prove the Checks sheet catches it,
   and adds new equity to prove an input flows all the way through.
4. Static: every cell of the monthly sheets is a formula, no error values, Excel 2016 functions only, no macros,
   charts read the right rows, fixed $ format.

    python test/verify.py         (run under tools/with_lock.py; writes test/last-result.txt)
"""
import re, shutil, sys, tempfile
from pathlib import Path
from openpyxl import load_workbook
import win32com.client as win32

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from model import run, N

ROOT = HERE.parent
BOOK = ROOT / "Three-Statement-Model.xlsx"
SCEN = ["Base", "Optimistic", "Pessimistic"]
results = []


def ok(name, cond, detail=""):
    results.append((bool(cond), name, detail))
    print(("PASS " if cond else "FAIL ") + name + (f"  [{detail}]" if detail and not cond else ""))


def close(a, b, tol=0.01):
    return a is not None and b is not None and abs(a - b) <= tol


# ------------------------------------------------------------------ 1. assumptions as the workbook holds them
wbf = load_workbook(BOOK)                      # formulas
wbv = load_workbook(BOOK, data_only=True)      # values Excel saved
asn = wbv["Assumptions"]


def name_cell(n):
    dest = list(wbf.defined_names[n].destinations)[0]
    return dest[0], dest[1].replace("$", "")


def scen_values(s):
    """live assumptions for scenario s (0..2): the driver names point at column G; take D/E/F of that row"""
    A = {}
    for n in wbf.defined_names:
        sh, ref = name_cell(n)
        if sh == "Assumptions" and ref.startswith("G"):
            A[n] = asn[f"{'DEF'[s]}{ref[1:]}"].value
    return A


sh, rng = name_cell("Season")
a, b = rng.split(":")
SEASON = [asn[f"D{r}"].value for r in range(int(a[1:]), int(b[1:]) + 1)]
OPEN = {n: asn[name_cell(n)[1]].value for n in wbf.defined_names if n.startswith("OPEN_")}
ok("assumptions read back: 39 drivers x 3 scenarios, 12 seasonality factors, opening balance sheet",
   len(scen_values(0)) == 39 and len(SEASON) == 12 and len(OPEN) == 10, f"{len(scen_values(0))} {len(SEASON)} {len(OPEN)}")
PY = [run(scen_values(s), SEASON, OPEN) for s in range(3)]

# ------------------------------------------------------------------ 4. static checks on the saved file
MONTHLY = ["Revenue & costs", "P&L", "Balance Sheet", "Cash Flow", "Schedules"]
typed = []
for s in MONTHLY:
    ws = wbf[s]
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row, min_col=5, max_col=40):
        for c in row:
            if c.value is not None and not (isinstance(c.value, str) and c.value.startswith("=")):
                typed.append(f"{s}!{c.coordinate}")
ok("every cell of the 5 monthly sheets (columns E-AN) is a formula", not typed, ", ".join(typed[:5]))
errs = [f"{ws.title}!{c.coordinate}" for ws in wbv.worksheets for row in ws.iter_rows() for c in row
        if isinstance(c.value, str) and c.value.startswith("#")]
ok("no error values anywhere", not errs, ", ".join(errs[:5]))
NEW = re.compile(r"\b(XLOOKUP|XMATCH|LET|LAMBDA|FILTER|SORT|SORTBY|UNIQUE|SEQUENCE|MINIFS|MAXIFS|IFS|SWITCH|TEXTJOIN|CONCAT)\(", re.I)
newf = [f"{ws.title}!{c.coordinate}" for ws in wbf.worksheets for row in ws.iter_rows() for c in row
        if isinstance(c.value, str) and c.value.startswith("=") and NEW.search(c.value)]
ok("Excel 2016 functions only (no XLOOKUP, LET, FILTER, MINIFS, IFS...)", not newf, ", ".join(newf[:5]))
ok("no macros (.xlsx)", BOOK.suffix == ".xlsx" and not wbf.vba_archive if hasattr(wbf, "vba_archive") else BOOK.suffix == ".xlsx")

# ------------------------------------------------------------------ 2. Excel, scenario by scenario
LINES = {  # sheet: [(label in column B, python key, sign)]
    "Revenue & costs": [("Coffee sold (kg)", "ws_kg", 1), ("Price per kg ($)", "ws_price", 1), ("Wholesale revenue", "ws_rev", 1),
                        ("Orders", "on_orders", 1), ("Average order ($)", "on_aov", 1), ("Online revenue", "on_rev", 1),
                        ("Total revenue", "rev", 1), ("Total cost of goods", "cogs", 1), ("Headcount", "heads", 1),
                        ("Payroll", "payroll", 1), ("Marketing", "mkt", 1), ("Card fees and free shipping", "fees", 1),
                        ("Other overheads", "other", 1), ("Total operating costs", "opex", 1)],
    "P&L": [("Revenue", "rev", 1), ("Cost of goods", "cogs", -1), ("Gross profit", "gross", 1), ("Total operating costs", "opex", -1),
            ("EBITDA", "ebitda", 1), ("Depreciation", "dep", -1), ("Operating profit (EBIT)", "ebit", 1), ("Interest", "interest", -1),
            ("Profit before tax", "ebt", 1), ("Income tax", "tax", -1), ("Net income", "ni", 1)],
    "Schedules": [("Capital spending", "capex", 1), ("Depreciation", "dep", 1), ("Equipment at cost", "fa_cost", 1),
                  ("Accumulated depreciation", "accdep", 1), ("Net book value", "nfa", 1), ("Drawn", "draw", 1),
                  ("Interest", "interest", 1), ("Payment (interest + principal)", "payment", 1), ("Principal repaid", "principal", 1),
                  ("Balance at the end of the month", "loan", 1), ("Customers owe us (wholesale, DSO)", "ar", 1),
                  ("Stock (days of cost of goods)", "inv", 1), ("Purchases", "purch", 1), ("We owe suppliers (DPO)", "ap", 1),
                  ("Tax for the month", "tax", 1), ("Tax paid (January, April, July, October)", "tax_paid", 1),
                  ("Tax payable", "tax_payable", 1), ("New equity paid in", "equity_in", 1)],
    "Balance Sheet": [("Cash", "cash", 1), ("Total assets", "assets", 1), ("Total liabilities", "liab", 1), ("Share capital", "share", 1),
                      ("Retained earnings", "re", 1), ("Total equity", "equity", 1)],
    "Cash Flow": [("Cash from operations", "cfo", 1), ("Cash from investing", "cfi", 1), ("Cash from financing", "cff", 1),
                  ("Net change in cash", "net_cash", 1), ("Cash at the end of the month", "cash", 1)],
}
TMP = Path(tempfile.gettempdir()) / "e12-verify.xlsx"
shutil.copy(BOOK, TMP)
xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(TMP))
    for w in wb.Worksheets:
        w.Unprotect()
    xl.Calculation = -4105
    W = lambda n: wb.Worksheets(n)
    NM = lambda n: wb.Names(n).RefersToRange.Value

    def rows_of(sheet):
        ws = W(sheet); out = {}
        for r in range(8, 80):
            v = ws.Cells(r, 2).Value
            if v and v not in out and (ws.Cells(r, 5).Formula or ws.Cells(r, 4).Formula):   # skip section headings
                out[v] = r
        return out

    def checks_all_pass():
        ws = W("Checks")
        res = [ws.Cells(r, 4).Value for r in range(7, 40) if ws.Cells(r, 4).Value in ("PASS", "FAIL")]
        return res

    ROWS = {s: rows_of(s) for s in LINES}
    ann = rows_of("Annual")
    scn = W("Scenarios")
    for s, nm in enumerate(SCEN):
        scn.Range("D6").Value = nm
        xl.CalculateFull()
        py = PY[s]
        bad, n = [], 0
        for sheet, lines in LINES.items():
            ws = W(sheet)
            for label, key, sign in lines:
                r = ROWS[sheet][label]
                vals = ws.Range(f"E{r}:AN{r}").Value[0]
                for m in range(N):
                    n += 1
                    if not close(vals[m], sign * py[key][m]):
                        bad.append(f"{sheet}/{label}/m{m + 1}: {vals[m]} vs {sign * py[key][m]}")
        ok(f"{nm}: {n} monthly numbers (54 lines x 36 months) = Python model", not bad, "; ".join(bad[:3]))
        diffs = W("Balance Sheet").Range(f"E{ROWS['Balance Sheet']['Check: assets - liabilities - equity (must be 0)']}:AN{ROWS['Balance Sheet']['Check: assets - liabilities - equity (must be 0)']}").Value[0]
        ok(f"{nm}: balance sheet balances in all 36 months (largest gap {max(abs(x) for x in diffs):.2e})", max(abs(x) for x in diffs) < 0.005)
        bad = []
        for y, col in ((1, "D"), (2, "E"), (3, "F")):
            Y = py["year"][y]
            for label, val in (("Revenue", Y["rev"]), ("EBITDA", Y["ebitda"]), ("Net income", Y["ni"]), ("Cash at 31 December", Y["cash_end"]),
                               ("Cash from operations", Y["cfo"]), ("Tax paid", Y["tax_paid"])):
                v = W("Annual").Range(f"{col}{ann[label]}").Value
                if not close(v, val): bad.append(f"{label} {y}: {v} vs {val}")
            d = W("Annual").Range(f"{col}{ann['Debt service cover: (EBITDA - tax paid) / loan payments']}").Value
            if not close(d, Y["dscr"], 1e-6): bad.append(f"DSCR {y}: {d} vs {Y['dscr']}")
        ok(f"{nm}: annual sheet (revenue, EBITDA, net income, cash, CFO, tax paid, DSCR x 3 years) = Python", not bad, "; ".join(bad[:3]))
        be = py["breakeven"] or 37
        kf = [("LowCash", py["low_cash"]), ("LowMonth", py["low_month"]), ("BreakEvenM", be), ("Funding", py["funding"]),
              ("MonthsBelow", py["below_min"]), ("Cash_Y3", py["year"][3]["cash_end"])]
        bad = [f"{k}: {NM(k)} vs {v}" for k, v in kf if not close(NM(k), v)]
        ok(f"{nm}: key figures (lowest cash and month, break-even, funding needed, months below minimum) = Python", not bad, "; ".join(bad))
        res = checks_all_pass()
        ok(f"{nm}: the workbook's own {len(res)} checks all PASS", res and all(x == "PASS" for x in res), str(res))
    # scenario comparison table, with Base picked again
    scn.Range("D6").Value = "Base"; xl.CalculateFull()
    labels = {scn.Cells(r, 2).Value: r for r in range(15, 27)}
    bad = []
    for s, col in enumerate("DEF"):
        py = PY[s]
        exp = {"Revenue 2027": py["year"][1]["rev"], "Revenue 2029": py["year"][3]["rev"], "EBITDA 2029": py["year"][3]["ebitda"],
               "Net income 2029": py["year"][3]["ni"], "Cash at 31 December 2029": py["year"][3]["cash_end"],
               "Lowest month-end cash": py["low_cash"], "Month of the lowest cash": py["low_month"],
               "Break-even month (EBITDA stays above 0)": py["breakeven"] or 37, "New funding needed": py["funding"]}
        for lab, v in exp.items():
            x = scn.Range(f"{col}{labels[lab]}").Value
            if not close(x, v): bad.append(f"{SCEN[s]} {lab}: {x} vs {v}")
    ok("scenario table: all 3 columns (Excel data table) = Python model run for each scenario", not bad, "; ".join(bad[:3]))
    # sensitivity tables
    A0 = scen_values(0)
    for title, key in (("Sensitivity: EBITDA 2029", "ebitda"), ("Sensitivity: lowest month-end cash", "low")):
        top = next(r for r in range(25, 60) if scn.Cells(r, 2).Value == title)
        hdr = top + 2
        bad = []
        for i in range(5):
            vol = scn.Cells(hdr + 1 + i, 3).Value
            for j in range(5):
                price = scn.Cells(hdr, 4 + j).Value
                r = run(A0, SEASON, OPEN, price_adj=price, vol_adj=vol)
                exp = r["year"][3]["ebitda"] if key == "ebitda" else r["low_cash"]
                x = scn.Cells(hdr + 1 + i, 4 + j).Value
                if not close(x, exp): bad.append(f"{vol}/{price}: {x} vs {exp}")
        ok(f"{title.lower()}: 25 cells (price -10%..+10% x volume -20%..+20%) = Python", not bad, "; ".join(bad[:3]))
    # 3. an input flows through: Pessimistic + $90,000 of new equity in month 20
    asw = W("Assumptions")
    eq_r = next(r for r in range(9, 80) if asw.Cells(r, 2).Value == "New equity from owners")
    eqm_r = next(r for r in range(9, 80) if asw.Cells(r, 2).Value == "New equity in month")
    scn.Range("D6").Value = "Pessimistic"; asw.Range(f"F{eq_r}").Value = 90000; asw.Range(f"F{eqm_r}").Value = 20
    xl.CalculateFull()
    A2 = scen_values(2); A2["EQUITY_AMT"], A2["EQUITY_M"] = 90000, 20
    py = run(A2, SEASON, OPEN)
    cash_r = ROWS["Cash Flow"]["Cash at the end of the month"]
    vals = W("Cash Flow").Range(f"E{cash_r}:AN{cash_r}").Value[0]
    ok(f"new input flows through: Pessimistic + $90,000 equity in month 20 -> cash in all 36 months = Python, funding needed {NM('Funding'):,.0f} (was {PY[2]['funding']:,.0f})",
       all(close(vals[m], py["cash"][m]) for m in range(N)) and close(NM("Funding"), py["funding"]) and all(x == "PASS" for x in checks_all_pass()))
    asw.Range(f"F{eq_r}").Value = 0; asw.Range(f"F{eqm_r}").Value = 12; scn.Range("D6").Value = "Base"
    # the checks catch mistakes
    chk = W("Checks")
    crow = {chk.Cells(r, 2).Value: r for r in range(7, 40) if chk.Cells(r, 2).Value}
    pl_r = ROWS["P&L"]["Payroll"]
    W("P&L").Range(f"J{pl_r}").Value = -12000; xl.CalculateFull()
    t = chk.Cells(crow["No typed numbers in the monthly sheets (cells with a number that is not a formula)"], 4).Value
    ok("a number typed over a formula (P&L payroll, June 2027) is caught: 'No typed numbers' says FAIL", t == "FAIL", str(t))
    W("P&L").Range(f"J{pl_r}").FormulaR1C1 = W("P&L").Range(f"I{pl_r}").FormulaR1C1
    bs_r = ROWS["Balance Sheet"]["Cash"]
    old = W("Balance Sheet").Range(f"M{bs_r}").Formula
    W("Balance Sheet").Range(f"M{bs_r}").Formula = f"=L{bs_r}"; xl.CalculateFull()
    t = chk.Cells(crow["Balance sheet balances: months where assets <> liabilities + equity"], 4).Value
    ok("a broken link (balance sheet cash, Sep 2027 = August) is caught: the balance check says FAIL", t == "FAIL", str(t))
    W("Balance Sheet").Range(f"M{bs_r}").Formula = old; xl.CalculateFull()
    ok("after undoing both, all checks PASS again", all(x == "PASS" for x in checks_all_pass()))
    # charts
    dash = W("Dashboard")
    f = {co.Name: [co.Chart.SeriesCollection(i).Formula for i in range(1, co.Chart.SeriesCollection().Count + 1)] for co in dash.ChartObjects()}
    rr = lambda sheet, label: f"$E${ROWS[sheet][label]}:$AN${ROWS[sheet][label]}"
    good = (len(f) == 3 and rr("P&L", "EBITDA") in f["chtRevenue"][2] and rr("Cash Flow", "Cash at the end of the month") in f["chtCash"][0]
            and "Scenarios" in f["chtScenarios"][0])
    ok("dashboard charts read the right rows (revenue/EBITDA, cash vs minimum, scenarios)", good, str(f)[:200])
    fmts = [dash.ChartObjects(n).Chart.Axes(2).TickLabels.NumberFormat for n in ("chtRevenue", "chtCash", "chtScenarios")]
    ok("chart axes use a fixed $ format (not the PC's currency)", all("[$$-" in x for x in fmts), str(fmts))
    wb.Close(False)
finally:
    xl.Quit()
    TMP.unlink(missing_ok=True)

passed = sum(1 for r in results if r[0])
print(f"\n{passed}/{len(results)} checks pass")
(HERE / "last-result.txt").write_text(f"{passed}/{len(results)}")
sys.exit(0 if passed == len(results) else 1)
