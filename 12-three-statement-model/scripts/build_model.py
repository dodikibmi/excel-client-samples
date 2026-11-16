"""Builds Three-Statement-Model.xlsx with real Excel (COM): one assumptions sheet drives a 36-month income statement,
balance sheet and cash flow, with capex/depreciation, loan and working-capital schedules, three scenarios, two
sensitivity tables and an investor dashboard. No macros, Excel 2016 functions.

Sheets
  Dashboard        6 cards, 3 insights written by formulas, revenue/EBITDA and cash charts, 3-year table, scenarios
  Assumptions      every input, Base / Optimistic / Pessimistic, the column in use, seasonality, opening balance sheet
  Scenarios        scenario picker, price/volume nudges, scenario comparison and 2 sensitivity tables (Excel data tables)
  Revenue & costs  monthly revenue build (wholesale kg, online orders), cost of goods, people and overheads
  P&L, Balance Sheet, Cash Flow   36 months, every cell a formula
  Schedules        capital spending and depreciation, bank loan (annuity), working capital, income tax, new equity
  Annual           the 3 statements by year, DSCR, and the key figures the dashboard uses
  Checks           the balance sheet balances every month, retained earnings roll, loan, tax, no typed numbers...
  How to use

    python scripts/build_model.py        (run under tools/with_lock.py)
"""
import json, re
from datetime import date
from pathlib import Path
import win32com.client as win32
from openpyxl.utils import get_column_letter as CL

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Three-Statement-Model.xlsx"
J = json.loads((ROOT / "data" / "assumptions.json").read_text())
N, E0 = 36, 5                                   # 36 months in columns E..AN; D = opening (31 Dec 2026)
LAST = CL(E0 + N - 1)                           # AN
MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def hexrgb(h):
    h = h.lstrip("#"); return int(h[:2], 16) + int(h[2:4], 16) * 256 + int(h[4:], 16) * 65536


# ---- light theme: white-blue page, navy brand, teal and amber series
PAGE, CARD, LINE = "#F4F6F9", "#FFFFFF", "#DCE2EA"
ACCENT, ACCENT2, S2, S3 = "#1F3A5F", "#5B7DA8", "#2A9D8F", "#E9A23B"
TEXT, MUTED, GOOD, BAD, INPUT, INPUTBG = "#1E2833", "#66717E", "#2F8A4E", "#B83A2E", "#1D4ED8", "#FFF9E6"
SERIF, BODY = "Georgia", "Segoe UI"
BRAND = "   ".join(" ".join(w) for w in "HARBOR & PINE COFFEE ROASTERS".split())
K = '#,##0;-#,##0;"-"'
P = '0.0%;-0.0%;"-"'
USD = "[$$-409]#,##0;-[$$-409]#,##0"


def mtxt(m):
    """month number 1..36 (an expression) -> 'Mar 2028', without TEXT() date codes (they follow the PC's language)"""
    return (f'CHOOSE(MOD(({m})-1,12)+1,' + ",".join(f'"{x}"' for x in MON) +
            f')&" "&(YEAR(StartDate)+INT((({m})-1)/12))')


def style(r, fg=None, bg=None, size=None, bold=None, name=None, italic=None, h=None, v=None, fmt=None, wrap=None):
    if fg: r.Font.Color = hexrgb(fg)
    if bg: r.Interior.Color = hexrgb(bg)
    if size: r.Font.Size = size
    if bold is not None: r.Font.Bold = bold
    if name: r.Font.Name = name
    if italic is not None: r.Font.Italic = italic
    if h: r.HorizontalAlignment = {"l": -4131, "c": -4108, "r": -4152}[h]
    if v: r.VerticalAlignment = {"t": -4160, "c": -4108, "b": -4107}[v]
    if fmt: r.NumberFormat = fmt
    if wrap is not None: r.WrapText = wrap
    return r


def border(r, color, edges=(7, 8, 9, 10), weight=2):
    for e in edges:
        b = r.Borders(e); b.LineStyle = 1; b.Weight = weight; b.Color = hexrgb(color)


def page(ws, title, sub):
    ws.Cells.Font.Name = BODY; ws.Cells.Font.Size = 10; ws.Cells.Font.Color = hexrgb(TEXT)
    ws.Cells.Interior.Color = hexrgb(PAGE)
    style(ws.Range("B2"), fg=ACCENT2, size=8, bold=True).Value = BRAND
    style(ws.Range("B3"), size=20, name=SERIF, fg=ACCENT).Value = title
    style(ws.Range("B4"), fg=MUTED, size=9).Value = sub
    ws.Rows(3).RowHeight = 30
    ws.Columns("A").ColumnWidth = 2


def chart_style(ch):
    ch.ChartArea.Format.Fill.ForeColor.RGB = hexrgb(CARD); ch.ChartArea.Format.Line.ForeColor.RGB = hexrgb(LINE)
    ch.PlotArea.Format.Fill.Visible = False
    tr = ch.ChartArea.Format.TextFrame2.TextRange.Font; tr.Fill.ForeColor.RGB = hexrgb(MUTED); tr.Size = 9; tr.Name = BODY


# ---------------------------------------------------------------------------------------------- monthly sheet specs
# (key, label, number format, formula for column E (filled right), formula for column D (opening) or None, look)
# {key} = same column of that row, {key.p} = previous column, {key.cum} = SUM from E to this column,
# {c} / {p} = this / previous column letter, {row} = this row. look: "" | "sub" | "total" | "grand" | "pct" | "memo"
OPS = [
    ("h", "Prices and seasonality"),
    ("idx", "Price index (1.000 = 2027 prices)", "0.000", "=(1+PRICE_RISE)^({c}$6-1)", None, ""),
    ("season", "Seasonality (1.00 = an average month)", "0.00", "=INDEX(Season,MONTH({c}$5))", None, ""),
    ("h", "Wholesale: cafés and offices"),
    ("ws_kg", "Coffee sold (kg)", "#,##0", "=WS_KG0*(1+WS_GROWTH)^({c}$7-1)*{season}*(1+VolAdj)", None, ""),
    ("ws_price", "Price per kg ($)", "#,##0.00", "=WS_PRICE*{idx}*(1+PriceAdj)", None, ""),
    ("ws_rev", "Wholesale revenue", K, "={ws_kg}*{ws_price}", None, "sub"),
    ("h", "Online shop"),
    ("on_orders", "Orders", "#,##0", "=ON_ORD0*(1+ON_GROWTH)^({c}$7-1)*{season}*(1+VolAdj)", None, ""),
    ("on_aov", "Average order ($)", "#,##0.00", "=ON_AOV*{idx}*(1+PriceAdj)", None, ""),
    ("on_rev", "Online revenue", K, "={on_orders}*{on_aov}", None, "sub"),
    ("rev", "Total revenue", K, "={ws_rev}+{on_rev}", None, "grand"),
    ("h", "Cost of goods"),
    ("ws_cogs", "Wholesale", K, "={ws_rev}*WS_COGS", None, ""),
    ("on_cogs", "Online", K, "={on_rev}*ON_COGS", None, ""),
    ("cogs", "Total cost of goods", K, "={ws_cogs}+{on_cogs}", None, "total"),
    ("h", "People and overheads"),
    ("heads", "Headcount", "0", "=CHOOSE({c}$6,HEAD_Y1,HEAD_Y2,HEAD_Y3)", None, ""),
    ("payroll", "Payroll", K, "={heads}*PAY_HEAD*(1+PAY_RISE)^({c}$6-1)", None, ""),
    ("rent", "Rent", K, "=RENT", None, ""),
    ("mkt", "Marketing", K, "={on_rev}*MKT", None, ""),
    ("fees", "Card fees and free shipping", K, "={on_rev}*ON_FEES", None, ""),
    ("other", "Other overheads", K, "=OTHER_OPEX*(1+COST_INFL)^({c}$6-1)", None, ""),
    ("opex", "Total operating costs", K, "={payroll}+{rent}+{mkt}+{fees}+{other}", None, "total"),
]
PL = [
    ("h", "Revenue"),
    ("pl_ws", "Wholesale", K, "={ws_rev}", None, ""),
    ("pl_on", "Online", K, "={on_rev}", None, ""),
    ("pl_rev", "Revenue", K, "={pl_ws}+{pl_on}", None, "total"),
    ("pl_cogs", "Cost of goods", K, "=-{cogs}", None, ""),
    ("gross", "Gross profit", K, "={pl_rev}+{pl_cogs}", None, "total"),
    ("gm", "Gross margin", P, "=IF({pl_rev}=0,0,{gross}/{pl_rev})", None, "pct"),
    ("h", "Operating costs"),
    ("pl_pay", "Payroll", K, "=-{payroll}", None, ""),
    ("pl_rent", "Rent", K, "=-{rent}", None, ""),
    ("pl_mkt", "Marketing", K, "=-{mkt}", None, ""),
    ("pl_fees", "Card fees and free shipping", K, "=-{fees}", None, ""),
    ("pl_other", "Other overheads", K, "=-{other}", None, ""),
    ("pl_opex", "Total operating costs", K, "={pl_pay}+{pl_rent}+{pl_mkt}+{pl_fees}+{pl_other}", None, "total"),
    ("ebitda", "EBITDA", K, "={gross}+{pl_opex}", None, "grand"),
    ("ebitda_m", "EBITDA margin", P, "=IF({pl_rev}=0,0,{ebitda}/{pl_rev})", None, "pct"),
    ("pl_dep", "Depreciation", K, "=-{dep}", None, ""),
    ("ebit", "Operating profit (EBIT)", K, "={ebitda}+{pl_dep}", None, "total"),
    ("pl_int", "Interest", K, "=-{interest}", None, ""),
    ("ebt", "Profit before tax", K, "={ebit}+{pl_int}", None, "total"),
    ("pl_tax", "Income tax", K, "=-{tax}", None, ""),
    ("ni", "Net income", K, "={ebt}+{pl_tax}", None, "grand"),
]
CAPEX_NAMES = {1: "15 kg roaster", 2: "Packaging line", 3: "Delivery van"}
SCH = [("h", "Capital spending and depreciation (straight line, from the month bought)")]
for k, nm in CAPEX_NAMES.items():
    SCH += [(f"cx{k}", f"{nm}: bought", K, f"=IF({{c}}$7=CAPEX{k}_M,CAPEX{k}_AMT,0)", None, ""),
            (f"dp{k}", f"{nm}: depreciation", K,
             f"=IF(AND({{c}}$7>=CAPEX{k}_M,{{c}}$7<CAPEX{k}_M+CAPEX{k}_LIFE*12),CAPEX{k}_AMT/(CAPEX{k}_LIFE*12),0)", None, "")]
SCH += [
    ("dp0", "Equipment owned at the start: depreciation", K, "=MIN(OPEN_DEP,OPEN_FA_COST-OPEN_ACCDEP-SUM($D{row}:{p}{row}))", None, ""),
    ("capex", "Capital spending", K, "={cx1}+{cx2}+{cx3}", None, "total"),
    ("dep", "Depreciation", K, "={dp0}+{dp1}+{dp2}+{dp3}", None, "total"),
    ("fa_cost", "Equipment at cost", K, "={fa_cost.p}+{capex}", "=OPEN_FA_COST", ""),
    ("accdep", "Accumulated depreciation", K, "={accdep.p}+{dep}", "=OPEN_ACCDEP", ""),
    ("nfa", "Net book value", K, "={fa_cost}-{accdep}", "={fa_cost}-{accdep}", "sub"),
    ("h", "Bank loan (annuity: the same payment every month)"),
    ("loan_open", "Balance at the start of the month", K, "={loan.p}", None, ""),
    ("draw", "Drawn", K, "=IF({c}$7=LOAN_M,LOAN_AMT,0)", None, ""),
    ("interest", "Interest", K, "=({loan_open}+{draw})*LOAN_RATE/12", None, ""),
    ("payment", "Payment (interest + principal)", K,
     "=IF(AND({c}$7>=LOAN_M,{c}$7<LOAN_M+LOAN_TERM),PMT(LOAN_RATE/12,LOAN_TERM,-LOAN_AMT),0)", None, ""),
    ("principal", "Principal repaid", K, "=IF({payment}=0,0,{payment}-{interest})", None, ""),
    ("loan", "Balance at the end of the month", K, "={loan_open}+{draw}-{principal}", None, "sub"),
    ("h", "Working capital"),
    ("ar", "Customers owe us (wholesale, DSO)", K, "={ws_rev}*DSO/30", "=OPEN_AR", ""),
    ("inv", "Stock (days of cost of goods)", K, "={cogs}*DIO/30", "=OPEN_INV", ""),
    ("purch", "Purchases", K, "={cogs}+{inv}-{inv.p}", None, ""),
    ("ap", "We owe suppliers (DPO)", K, "={purch}*DPO/30", "=OPEN_AP", ""),
    ("h", "Income tax (losses carried forward, paid after each quarter)"),
    ("ebt_cum", "Profit before tax since January 2027", K, "={ebt.cum}", None, ""),
    ("tax", "Tax for the month", K, "=MAX(0,TAX_RATE*MAX(0,{ebt_cum})-SUM($D{row}:{p}{row}))", None, ""),
    ("tax_paid", "Tax paid (January, April, July, October)", K, "=IF(MOD(MONTH({c}$5),3)=1,{tax_payable.p},0)", None, ""),
    ("tax_payable", "Tax payable", K, "={tax_payable.p}+{tax}-{tax_paid}", "=OPEN_TAX", "sub"),
    ("h", "Owners"),
    ("equity_in", "New equity paid in", K, "=IF({c}$7=EQUITY_M,EQUITY_AMT,0)", None, ""),
]
BS = [
    ("h", "Assets"),
    ("bs_cash", "Cash", K, "={cash}", "={cash}", ""),
    ("bs_ar", "Customers owe us", K, "={ar}", "={ar}", ""),
    ("bs_inv", "Stock", K, "={inv}", "={inv}", ""),
    ("bs_ca", "Current assets", K, "={bs_cash}+{bs_ar}+{bs_inv}", "={bs_cash}+{bs_ar}+{bs_inv}", "total"),
    ("bs_fa", "Equipment at cost", K, "={fa_cost}", "={fa_cost}", ""),
    ("bs_acc", "Accumulated depreciation", K, "=-{accdep}", "=-{accdep}", ""),
    ("bs_nfa", "Equipment, net", K, "={bs_fa}+{bs_acc}", "={bs_fa}+{bs_acc}", "total"),
    ("bs_ta", "Total assets", K, "={bs_ca}+{bs_nfa}", "={bs_ca}+{bs_nfa}", "grand"),
    ("h", "Liabilities"),
    ("bs_ap", "We owe suppliers", K, "={ap}", "={ap}", ""),
    ("bs_tax", "Tax payable", K, "={tax_payable}", "={tax_payable}", ""),
    ("bs_loan", "Bank loan", K, "={loan}", "={loan}", ""),
    ("bs_tl", "Total liabilities", K, "={bs_ap}+{bs_tax}+{bs_loan}", "={bs_ap}+{bs_tax}+{bs_loan}", "total"),
    ("h", "Equity"),
    ("bs_share", "Share capital", K, "={bs_share.p}+{equity_in}", "=OPEN_SHARE", ""),
    ("bs_re", "Retained earnings", K, "={bs_re.p}+{ni}", "=OPEN_RE", ""),
    ("bs_te", "Total equity", K, "={bs_share}+{bs_re}", "={bs_share}+{bs_re}", "total"),
    ("bs_tle", "Total liabilities and equity", K, "={bs_tl}+{bs_te}", "={bs_tl}+{bs_te}", "grand"),
    ("bs_diff", "Check: assets - liabilities - equity (must be 0)", '0.00;-0.00;"0.00"', "={bs_ta}-{bs_tle}", "={bs_ta}-{bs_tle}", "memo"),
    ("min_cash", "Memo: minimum cash the bank wants", K, "=MIN_CASH", None, "memo"),
]
CF = [
    ("h", "Operating activities"),
    ("cf_ni", "Net income", K, "={ni}", None, ""),
    ("cf_dep", "Add back depreciation (no cash leaves)", K, "={dep}", None, ""),
    ("cf_ar", "Customers owe us: less (+) / more (-)", K, "={ar.p}-{ar}", None, ""),
    ("cf_inv", "Stock: less (+) / more (-)", K, "={inv.p}-{inv}", None, ""),
    ("cf_ap", "We owe suppliers: more (+) / less (-)", K, "={ap}-{ap.p}", None, ""),
    ("cf_tax", "Tax payable: more (+) / less (-)", K, "={tax_payable}-{tax_payable.p}", None, ""),
    ("cfo", "Cash from operations", K, "={cf_ni}+{cf_dep}+{cf_ar}+{cf_inv}+{cf_ap}+{cf_tax}", None, "total"),
    ("h", "Investing activities"),
    ("cf_capex", "Equipment bought", K, "=-{capex}", None, ""),
    ("cfi", "Cash from investing", K, "={cf_capex}", None, "total"),
    ("h", "Financing activities"),
    ("cf_draw", "Loan drawn", K, "={draw}", None, ""),
    ("cf_prin", "Loan principal repaid", K, "=-{principal}", None, ""),
    ("cf_eq", "New equity", K, "={equity_in}", None, ""),
    ("cff", "Cash from financing", K, "={cf_draw}+{cf_prin}+{cf_eq}", None, "total"),
    ("net", "Net change in cash", K, "={cfo}+{cfi}+{cff}", None, "grand"),
    ("cash_open", "Cash at the start of the month", K, "={cash.p}", None, ""),
    ("cash", "Cash at the end of the month", K, "={cash_open}+{net}", "=OPEN_CASH", "grand"),
]
MONTHLY = [("Revenue & costs", OPS, "Revenue and costs, month by month",
            "Wholesale kg x price, online orders x average order, seasonality; cost of goods, people and overheads. Every cell is a formula on Assumptions."),
           ("P&L", PL, "Income statement (P&L), 36 months", "Costs are shown as negative numbers. Every cell is a formula."),
           ("Balance Sheet", BS, "Balance sheet, end of each month", "Column D is the opening balance sheet (31 Dec 2026). The check row must be 0.00 in every month."),
           ("Cash Flow", CF, "Cash flow statement (indirect method)", "Starts from net income, adds back depreciation and the changes in working capital. Ending cash feeds the balance sheet."),
           ("Schedules", SCH, "Supporting schedules", "Capital spending and depreciation, bank loan, working capital, income tax, new equity.")]

ROW = {}                                          # key -> (sheet, row)
LAYOUT = {}
for sheet, spec, *_ in MONTHLY:
    r, rows = 8, []
    for i, it in enumerate(spec):
        if it[0] == "h":
            r += 1 if i else 0
            rows.append((r, it)); r += 1
        else:
            ROW[it[0]] = (sheet, r); rows.append((r, it)); r += 1
    LAYOUT[sheet] = rows


def F(tpl, c, here, row):
    """resolve a row template for column letter c on sheet `here`"""
    p = CL(max(1, ws_col(c) - 1))

    def sub(m):
        key, mod = m.group(1), m.group(3)
        if key == "c": return c
        if key == "p": return p
        if key == "row": return str(row)
        sh, r = ROW[key]
        pre = "" if sh == here else f"'{sh}'!"
        if mod == "p": return f"{pre}{p}{r}"
        if mod == "cum": return f"SUM({pre}$E${r}:{c}${r})"
        return f"{pre}{c}{r}"
    return re.sub(r"\{(\w+)(\.(p|cum))?\}", sub, tpl)


def ws_col(c):
    n = 0
    for ch in c: n = n * 26 + ord(ch) - 64
    return n


xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False; xl.ScreenUpdating = False
if OUT.exists():
    OUT.unlink()
try:
    wb = xl.Workbooks.Add()
    SHEETS = ["Dashboard", "Assumptions", "Scenarios", "Revenue & costs", "P&L", "Balance Sheet", "Cash Flow",
              "Schedules", "Annual", "Checks", "How to use"]
    while wb.Worksheets.Count < len(SHEETS):
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    for i, n in enumerate(SHEETS, 1):
        wb.Worksheets(i).Name = n
    W = {n: wb.Worksheets(n) for n in SHEETS}
    xl.Calculation = -4135                          # manual while building; full recalculation at the end

    # ------------------------------------------------------------------ Scenarios: the picker first (names used everywhere)
    sc = W["Scenarios"]
    page(sc, "Scenarios and sensitivity", "Pick a scenario here: the whole model follows. The tables below run the model for every scenario and for price and volume changes (Excel data tables).")
    sc.Range("B6").Value = "Scenario used in the model"; sc.Range("D6").Value = "Base"
    sc.Range("B7").Value = "Scenario number"; sc.Range("D7").Formula = "=MATCH(D6,Assumptions!$D$8:$F$8,0)"
    sc.Range("B9").Value = "Price change on top of the scenario"; sc.Range("D9").Value = 0
    sc.Range("B10").Value = "Volume change on top of the scenario"; sc.Range("D10").Value = 0
    for a, n in (("D6", "ScenarioName"), ("D7", "ScenarioNo"), ("D9", "PriceAdj"), ("D10", "VolAdj")):
        wb.Names.Add(n, f"=Scenarios!${a[0]}${a[1:]}")

    # ------------------------------------------------------------------ Assumptions
    ws = W["Assumptions"]
    page(ws, "Assumptions", "Blue cells on yellow are inputs. Base / Optimistic / Pessimistic side by side; 'Used now' is the column the model reads.")
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = False
    for col, w in zip("BCDEFGH", (40, 20, 12, 12, 12, 12, 44)):
        ws.Columns(col).ColumnWidth = w
    ws.Range("B6").Value = "Scenario used now"
    style(ws.Range("D6"), fg=ACCENT, bold=True, size=12).Formula = "=ScenarioName"
    style(ws.Range("E6"), fg=MUTED, size=9).Value = "change it on the Scenarios sheet"
    style(ws.Range("B8:H8"), fg="#FFFFFF", bg=ACCENT, bold=True).Value = ("Driver", "Unit", "Base", "Optimistic", "Pessimistic", "Used now", "Note")
    r = 9
    for d in J["drivers"]:
        if d[1] is None:
            style(ws.Range(f"B{r}"), fg=ACCENT, bold=True, size=10).Value = d[0]
            ws.Range(f"B{r}:H{r}").Interior.Color = hexrgb("#E8EDF4")
            r += 1; continue
        name, label, unit, fmt, vals, note = d
        ws.Cells(r, 2).Value = label; ws.Cells(r, 3).Value = unit
        ws.Range(f"D{r}:F{r}").Value = vals
        style(ws.Range(f"D{r}:F{r}"), fg=INPUT, bg=INPUTBG, fmt=fmt)
        style(ws.Range(f"G{r}"), bold=True, fmt=fmt).Formula = f"=INDEX(D{r}:F{r},ScenarioNo)"
        ws.Cells(r, 8).Value = note
        wb.Names.Add(name, f"=Assumptions!$G${r}")
        r += 1
    DRV_LAST = r - 1
    style(ws.Range(f"C9:C{DRV_LAST}"), fg=MUTED, size=9); style(ws.Range(f"H9:H{DRV_LAST}"), fg=MUTED, size=9)
    border(ws.Range(f"B8:H{DRV_LAST}"), LINE, (7, 8, 9, 10, 12))
    r += 1
    style(ws.Range(f"B{r}"), fg=ACCENT, bold=True).Value = "Seasonality: each month vs an average month (all scenarios)"
    ws.Range(f"B{r}:H{r}").Interior.Color = hexrgb("#E8EDF4"); r += 1
    s0 = r
    for i, v in enumerate(J["season"]):
        ws.Cells(r, 2).Value = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
                                "October", "November", "December"][i]
        ws.Cells(r, 3).Value = "x average"; ws.Cells(r, 4).Value = v
        style(ws.Cells(r, 4), fg=INPUT, bg=INPUTBG, fmt="0.00"); r += 1
    wb.Names.Add("Season", f"=Assumptions!$D${s0}:$D${r - 1}")
    style(ws.Range(f"C{s0}:C{r - 1}"), fg=MUTED, size=9)
    r += 1
    style(ws.Range(f"B{r}"), fg=ACCENT, bold=True).Value = "Opening balance sheet, 31 December 2026 (all scenarios)"
    ws.Range(f"B{r}:H{r}").Interior.Color = hexrgb("#E8EDF4"); r += 1
    ws.Cells(r, 2).Value = "First month of the plan"; ws.Cells(r, 4).Value = (date(2027, 1, 1) - date(1899, 12, 30)).days
    style(ws.Cells(r, 4), fg=INPUT, bg=INPUTBG, fmt="mmm yyyy"); wb.Names.Add("StartDate", f"=Assumptions!$D${r}"); r += 1
    o0 = r
    for name, label, v in J["opening"]:
        ws.Cells(r, 2).Value = label; ws.Cells(r, 4).Value = v
        style(ws.Cells(r, 4), fg=INPUT, bg=INPUTBG, fmt="#,##0")
        ws.Cells(r, 3).Value = "$ / month" if name == "OPEN_DEP" else "$"
        wb.Names.Add(name, f"=Assumptions!$D${r}"); r += 1
    style(ws.Range(f"C{o0}:C{r - 1}"), fg=MUTED, size=9)
    ws.Cells(r, 2).Value = "Check: assets - liabilities - equity (must be 0)"
    ws.Cells(r, 4).Formula = "=OPEN_CASH+OPEN_AR+OPEN_INV+OPEN_FA_COST-OPEN_ACCDEP-OPEN_AP-OPEN_TAX-OPEN_SHARE-OPEN_RE"
    style(ws.Range(f"B{r}:D{r}"), bold=True); ws.Cells(r, 4).NumberFormat = '0.00;-0.00;"0.00"'
    OPEN_CHECK = f"Assumptions!$D${r}"
    ws.Range("D9:F200").Locked = False; ws.Range(f"D{s0}:D{r - 1}").Locked = False
    ws.Range(f"G9:G{DRV_LAST}").Locked = True; ws.Range(f"D{r}").Locked = True
    ws.Range("A9").Select(); xl.ActiveWindow.FreezePanes = True

    # ------------------------------------------------------------------ the five monthly sheets
    for sheet, spec, title, sub in MONTHLY:
        ws = W[sheet]
        page(ws, title, sub)
        ws.Activate(); xl.ActiveWindow.DisplayGridlines = False
        ws.Columns("B").ColumnWidth = 42; ws.Columns("C").ColumnWidth = 2; ws.Columns("D").ColumnWidth = 11
        ws.Range(f"E1:{LAST}1").EntireColumn.ColumnWidth = 9.5
        ws.Range(f"E5:{LAST}5").Formula = "=EOMONTH(StartDate,E$7-1)"
        ws.Range(f"E6:{LAST}6").Formula = "=YEAR(E$5)-YEAR(StartDate)+1"
        ws.Range(f"E7:{LAST}7").Formula = "=COLUMN()-4"
        ws.Range("B5").Value = "Month end"; ws.Range("B6").Value = "Year of the plan"; ws.Range("B7").Value = "Month of the plan"
        style(ws.Range(f"B5:{LAST}5"), fg="#FFFFFF", bg=ACCENT, bold=True)
        ws.Range(f"E5:{LAST}5").NumberFormat = "mmm yy"
        style(ws.Range(f"B6:{LAST}7"), fg=MUTED, size=8)
        if any(it[0] != "h" and it[4] for it in spec):
            ws.Range("D5").Formula = "=StartDate-1"; ws.Range("D5").NumberFormat = '"Opening"'
            style(ws.Range("D5"), fg="#FFFFFF", bg=ACCENT, bold=True, h="r")
        ws.Range(f"E5:{LAST}5").HorizontalAlignment = -4152
        for r, it in LAYOUT[sheet]:
            if it[0] == "h":
                style(ws.Range(f"B{r}"), fg=ACCENT, bold=True).Value = it[1]
                continue
            key, label, fmt, te, td, look = it
            ws.Cells(r, 2).Value = label
            rng = ws.Range(f"E{r}:{LAST}{r}")
            rng.Formula = F(te, "E", sheet, r)
            rng.NumberFormat = fmt
            if td:
                ws.Range(f"D{r}").Formula = F(td, "D", sheet, r); ws.Range(f"D{r}").NumberFormat = fmt
            full = ws.Range(f"B{r}:{LAST}{r}")
            if look == "sub":
                full.Font.Bold = True
            elif look == "total":
                full.Font.Bold = True; border(full, LINE, (8,))
            elif look == "grand":
                full.Font.Bold = True; full.Interior.Color = hexrgb("#E8EDF4"); border(full, ACCENT2, (8,))
            elif look == "pct":
                full.Font.Italic = True; full.Font.Color = hexrgb(MUTED)
            elif look == "memo":
                full.Font.Color = hexrgb(MUTED); full.Font.Size = 9
        ws.Range("E8").Select(); xl.ActiveWindow.FreezePanes = True
        ws.Range(f"A1:{LAST}100").Locked = True

    # cash below the minimum: red (formula without commas, so it works whatever the PC's list separator)
    sh, r = ROW["bs_cash"]
    fc = W[sh].Range(f"E{r}:{LAST}{r}").FormatConditions.Add(1, 6, "=MIN_CASH"); fc.Font.Color = hexrgb(BAD)
    sh, r = ROW["bs_diff"]
    W[sh].Range(f"D{r}:{LAST}{r}").Font.Color = hexrgb(GOOD)

    # ------------------------------------------------------------------ Annual
    ws = W["Annual"]
    page(ws, "Annual summary", "The three statements by year, summed from the monthly sheets (balance sheet items: at 31 December).")
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = False
    ws.Columns("B").ColumnWidth = 44; ws.Columns("C").ColumnWidth = 2
    for col in "DEF": ws.Columns(col).ColumnWidth = 14
    ws.Range("D5:F5").Formula = "=YEAR(StartDate)+D$6-1"; ws.Range("D6:F6").Formula = "=COLUMN()-3"
    ws.Range("B5").Value = "Year"; style(ws.Range("B5:F5"), fg="#FFFFFF", bg=ACCENT, bold=True); ws.Range("D5:F5").HorizontalAlignment = -4152
    ws.Range("D5:F5").NumberFormat = "0"
    style(ws.Range("B6:F6"), fg=MUTED, size=8); ws.Range("B6").Value = "Year of the plan"

    def S(key):
        sh, r = ROW[key]; return f"=SUMIF('{sh}'!$E$6:${LAST}$6,D$6,'{sh}'!$E${r}:${LAST}${r})"

    def END(key):
        sh, r = ROW[key]; return f"=INDEX('{sh}'!$E${r}:${LAST}${r},D$6*12)"
    ANN = [("h", "Income statement"),
           ("a_rev", "Revenue", K, S("pl_rev"), "total"),
           ("a_growth", "Revenue growth on the year before", P, '=IF(D$6=1,"",D{a_rev}/C{a_rev}-1)', "pct"),
           ("a_ws", "   Wholesale", K, S("pl_ws"), ""), ("a_on", "   Online", K, S("pl_on"), ""),
           ("a_gross", "Gross profit", K, S("gross"), "total"), ("a_gm", "Gross margin", P, "=D{a_gross}/D{a_rev}", "pct"),
           ("a_opex", "Operating costs", K, S("pl_opex"), ""),
           ("a_ebitda", "EBITDA", K, S("ebitda"), "grand"), ("a_em", "EBITDA margin", P, "=D{a_ebitda}/D{a_rev}", "pct"),
           ("a_dep", "Depreciation", K, S("pl_dep"), ""), ("a_int", "Interest", K, S("pl_int"), ""),
           ("a_tax", "Income tax", K, S("pl_tax"), ""),
           ("a_ni", "Net income", K, S("ni"), "grand"), ("a_nm", "Net margin", P, "=D{a_ni}/D{a_rev}", "pct"),
           ("h", "Cash"),
           ("a_cfo", "Cash from operations", K, S("cfo"), ""), ("a_cfi", "Cash from investing", K, S("cfi"), ""),
           ("a_cff", "Cash from financing", K, S("cff"), ""),
           ("a_cash", "Cash at 31 December", K, END("cash"), "total"),
           ("a_low", "Lowest month-end cash in the year", K,
            f"=AGGREGATE(15,6,'Cash Flow'!$E${ROW['cash'][1]}:${LAST}${ROW['cash'][1]}/('Cash Flow'!$E$6:${LAST}$6=D$6),1)", ""),
           ("h", "Balance sheet at 31 December"),
           ("a_ta", "Total assets", K, END("bs_ta"), ""), ("a_loan", "Bank loan", K, END("bs_loan"), ""),
           ("a_te", "Total equity", K, END("bs_te"), ""),
           ("h", "Bank and people"),
           ("a_pay", "Loan payments (interest + principal)", K, S("payment"), ""),
           ("a_taxpaid", "Tax paid", K, S("tax_paid"), ""),
           ("a_dscr", "Debt service cover: (EBITDA - tax paid) / loan payments", '0.00"x";-0.00"x"', '=IF(D{a_pay}=0,"",(D{a_ebitda}-D{a_taxpaid})/D{a_pay})', "total"),
           ("a_heads", "Headcount at year end", "0", END("heads"), "")]
    AROW, r = {}, 8
    for i, it in enumerate(ANN):
        if it[0] == "h":
            r += 1 if i else 0; style(ws.Range(f"B{r}"), fg=ACCENT, bold=True).Value = it[1]; r += 1; continue
        AROW[it[0]] = r; r += 1
    for it in ANN:
        if it[0] == "h": continue
        key, label, fmt, ff, look = it
        r = AROW[key]
        ws.Cells(r, 2).Value = label
        ff = re.sub(r"\{(\w+)\}", lambda m: str(AROW[m.group(1)]), ff)
        ws.Range(f"D{r}:F{r}").Formula = ff; ws.Range(f"D{r}:F{r}").NumberFormat = fmt
        full = ws.Range(f"B{r}:F{r}")
        if look == "total": full.Font.Bold = True; border(full, LINE, (8,))
        if look == "grand": full.Font.Bold = True; full.Interior.Color = hexrgb("#E8EDF4")
        if look == "pct": full.Font.Italic = True; full.Font.Color = hexrgb(MUTED)
    # key figures (named; the dashboard and the scenario table read these)
    r = max(AROW.values()) + 2
    style(ws.Range(f"B{r}"), fg=ACCENT, bold=True).Value = "Key figures (used by the Dashboard and Scenarios)"; r += 1
    cash_r, eb_r = ROW["cash"][1], ROW["ebitda"][1]
    CASH = f"'Cash Flow'!$E${cash_r}:${LAST}${cash_r}"
    KF = [("Rev_Y1", "Revenue 2027", f"=D{AROW['a_rev']}", USD), ("Rev_Y2", "Revenue 2028", f"=E{AROW['a_rev']}", USD),
          ("Rev_Y3", "Revenue 2029", f"=F{AROW['a_rev']}", USD), ("EBITDA_Y3", "EBITDA 2029", f"=F{AROW['a_ebitda']}", USD),
          ("EBITDA_M_Y3", "EBITDA margin 2029", f"=F{AROW['a_em']}", "0.0%"), ("NI_Y3", "Net income 2029", f"=F{AROW['a_ni']}", USD),
          ("Cash_Y3", "Cash at 31 December 2029", f"=F{AROW['a_cash']}", USD),
          ("DSCR_Y3", "Debt service cover 2029", f"=F{AROW['a_dscr']}", '0.00"x"'),
          ("LowCash", "Lowest month-end cash", f"=MIN({CASH})", USD),
          ("LowMonth", "Month of the lowest cash (1-36)", f"=MATCH(LowCash,{CASH},0)", "0"),
          ("LowDateTxt", "  as a month", "=" + mtxt("LowMonth"), "@"),
          ("BreakEvenM", "Break-even: EBITDA above 0 from this month on (37 = not reached)",
           f"=IFERROR(LOOKUP(2,1/('P&L'!$E${eb_r}:${LAST}${eb_r}<=0),'P&L'!$E$7:${LAST}$7)+1,1)", "0"),
          ("BreakEvenTxt", "  as a month", f'=IF(BreakEvenM>36,"not reached",{mtxt("BreakEvenM")})', "@"),
          ("Funding", "New funding needed to keep the minimum cash", "=MAX(0,MIN_CASH-LowCash)", USD),
          ("MonthsBelow", "Months below the minimum cash", f'=COUNTIF({CASH},"<"&MIN_CASH)', "0")]
    for name, label, ff, fmt in KF:
        ws.Cells(r, 2).Value = label
        c = ws.Cells(r, 4); c.NumberFormat = "General" if fmt == "@" else fmt; c.Formula = ff
        c.Font.Bold = True; c.HorizontalAlignment = -4152
        wb.Names.Add(name, f"=Annual!$D${r}")
        r += 1

    # ------------------------------------------------------------------ Scenarios: comparison and sensitivity (data tables)
    sc.Activate(); xl.ActiveWindow.DisplayGridlines = False
    sc.Columns("B").ColumnWidth = 34
    for col in "CDEFGH": sc.Columns(col).ColumnWidth = 14
    style(sc.Range("D6"), fg=INPUT, bg=INPUTBG, bold=True, size=12)
    style(sc.Range("D9:D10"), fg=INPUT, bg=INPUTBG, fmt="+0%;-0%;0%")
    sc.Range("D7").Font.Color = hexrgb(MUTED)
    dv = sc.Range("D6").Validation; dv.Delete(); dv.Add(3, 1, 1, "=Assumptions!$D$8:$F$8")
    sc.Range("E6").Value = "<- pick Base, Optimistic or Pessimistic"; style(sc.Range("E6"), fg=MUTED, size=9)
    sc.Range("E9").Value = "e.g. +5% = every price 5% higher (0% = as in the scenario)"; style(sc.Range("E9:E10"), fg=MUTED, size=9)
    sc.Range("E10").Value = "e.g. -10% = 10% fewer kg and orders every month"
    style(sc.Range("B12"), fg=ACCENT, bold=True, size=11).Value = "Scenario comparison (price and volume changes above apply to all three)"
    sc.Range("B14").Value = "Result"; sc.Range("C14").Value = "Used now"
    sc.Range("D14:F14").Value = ("Base", "Optimistic", "Pessimistic")
    style(sc.Range("B14:F14"), fg="#FFFFFF", bg=ACCENT, bold=True); sc.Range("C14:F14").HorizontalAlignment = -4152
    SCN = [("Revenue 2027", "=Rev_Y1", USD), ("Revenue 2028", "=Rev_Y2", USD), ("Revenue 2029", "=Rev_Y3", USD),
           ("EBITDA 2029", "=EBITDA_Y3", USD), ("EBITDA margin 2029", "=EBITDA_M_Y3", "0.0%"),
           ("Net income 2029", "=NI_Y3", USD), ("Cash at 31 December 2029", "=Cash_Y3", USD),
           ("Lowest month-end cash", "=LowCash", USD), ("Month of the lowest cash", "=LowMonth", '"month "0'),
           ("Break-even month (EBITDA stays above 0)", "=BreakEvenM", '[>36]"not reached";"month "0'),
           ("New funding needed", "=Funding", USD), ("Debt service cover 2029", "=DSCR_Y3", '0.00"x"')]
    for i, (lab, ff, fmt) in enumerate(SCN, 15):
        sc.Cells(i, 2).Value = lab; sc.Cells(i, 3).Formula = ff
        sc.Range(f"C{i}:F{i}").NumberFormat = fmt
    sc.Range("C14").Value = "Used now"
    # one-variable data table: the scenario names in D14:F14 are put into D6 one by one
    sc.Range("D14:F14").Value = ("Base", "Optimistic", "Pessimistic")
    sc.Range("C14:F26").Table(RowInput=sc.Range("D6"))
    sc.Range("C14").Value = "Used now"
    border(sc.Range("B14:F26"), LINE, (7, 8, 9, 10, 12)); sc.Range("C15:C26").Font.Color = hexrgb(MUTED)
    sc.Range("D15:F26").Font.Bold = True
    SC_ROW = {lab: i for i, (lab, _, _) in enumerate(SCN, 15)}

    def sens(top, title, out_ff, fmt):
        style(sc.Range(f"B{top}"), fg=ACCENT, bold=True, size=11).Value = title
        sc.Range(f"B{top + 1}").Value = "Volume change (down) / price change (across)"
        style(sc.Range(f"B{top + 1}"), fg=MUTED, size=9)
        hdr = top + 2
        c = sc.Range(f"C{hdr}"); c.Formula = out_ff; c.NumberFormat = ';;;'          # the corner formula, hidden
        sc.Range(f"D{hdr}:H{hdr}").Value = (-0.10, -0.05, 0, 0.05, 0.10)
        for k, v in enumerate((-0.20, -0.10, 0, 0.10, 0.20)):
            sc.Cells(hdr + 1 + k, 3).Value = v
        sc.Range(f"C{hdr}:H{hdr + 5}").Table(RowInput=sc.Range("D9"), ColumnInput=sc.Range("D10"))
        style(sc.Range(f"D{hdr}:H{hdr}"), fg="#FFFFFF", bg=ACCENT, bold=True, fmt="+0%;-0%;0%", h="c")
        style(sc.Range(f"C{hdr + 1}:C{hdr + 5}"), fg="#FFFFFF", bg=ACCENT, bold=True, fmt="+0%;-0%;0%", h="c")
        body = sc.Range(f"D{hdr + 1}:H{hdr + 5}"); body.NumberFormat = fmt; body.HorizontalAlignment = -4108
        cs = body.FormatConditions.AddColorScale(3)
        cs.ColorScaleCriteria(1).FormatColor.Color = hexrgb("#F2C4BE")
        cs.ColorScaleCriteria(2).FormatColor.Color = hexrgb("#FFFFFF")
        cs.ColorScaleCriteria(3).FormatColor.Color = hexrgb("#BFE3D2")
        border(sc.Range(f"C{hdr}:H{hdr + 5}"), LINE, (7, 8, 9, 10, 11, 12))
        sc.Range(f"F{hdr + 3}").Font.Bold = True; border(sc.Range(f"F{hdr + 3}"), ACCENT, (7, 8, 9, 10), -4138)
        return hdr
    S1 = sens(29, "Sensitivity: EBITDA 2029", "=EBITDA_Y3", USD)
    S2r = sens(38, "Sensitivity: lowest month-end cash", "=LowCash", USD)
    sc.Range("B46").Value = "Tables are Excel data tables: they recalculate the whole model for every cell. They use the scenario picked above."
    style(sc.Range("B46"), fg=MUTED, size=9)
    sc.Range("A1:H60").Locked = True; sc.Range("D6").Locked = False; sc.Range("D9:D10").Locked = False

    # ------------------------------------------------------------------ Dashboard
    ws = W["Dashboard"]
    page(ws, "3-year plan, 2027-2029", "")
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = False; xl.ActiveWindow.DisplayHeadings = False
    for c in range(2, 26):
        ws.Columns(c).ColumnWidth = 2.2 if c in (5, 9, 13, 17, 21) else 9.5
    ws.Range("B4").Formula = ('="Scenario: "&ScenarioName&IF(OR(PriceAdj<>0,VolAdj<>0),"  (prices "&TEXT(PriceAdj,"+0%;-0%;0%")&", volume "&TEXT(VolAdj,"+0%;-0%;0%")&")","")'
                              '&"   ·   change it on Scenarios   ·   36 months, every number a formula on Assumptions"')
    ws.Range("B4").Font.Color = hexrgb(MUTED)
    cards = [("REVENUE 2029", "=Rev_Y3", USD, '="from "&TEXT(Rev_Y1,"[$$-409]#,##0")&" in 2027"'),
             ("EBITDA MARGIN 2029", "=EBITDA_M_Y3", "0.0%", '=TEXT(EBITDA_Y3,"[$$-409]#,##0;-[$$-409]#,##0")&" EBITDA in 2029"'),
             ("BREAK-EVEN", "=BreakEvenTxt", "General", '="EBITDA above 0 every month after"'),
             ("LOWEST CASH", "=LowCash", USD, '=LowDateTxt&"  ·  minimum "&TEXT(MIN_CASH,"[$$-409]#,##0")'),
             ("NEW FUNDING NEEDED", "=Funding", USD, '=IF(Funding=0,"cash stays above the minimum",MonthsBelow&" months below the minimum")'),
             ("DEBT SERVICE COVER 2029", "=DSCR_Y3", '0.00"x";-0.00"x"', '="banks often ask for 1.25x or more"')]
    for i, (lab, ff, fmt, note) in enumerate(cards):
        c0 = 2 + i * 4
        box = ws.Range(ws.Cells(6, c0), ws.Cells(10, c0 + 2))
        box.Interior.Color = hexrgb(CARD); border(box, LINE)
        border(ws.Range(ws.Cells(6, c0), ws.Cells(6, c0 + 2)), ACCENT, (8,), -4138)
        ws.Range(ws.Cells(6, c0), ws.Cells(6, c0 + 2)).Merge()
        style(ws.Cells(6, c0), fg=MUTED, size=8, bold=True).Value = lab; ws.Cells(6, c0).IndentLevel = 1
        ws.Range(ws.Cells(7, c0), ws.Cells(8, c0 + 2)).Merge()
        v = ws.Cells(7, c0); v.NumberFormat = fmt; v.Formula = ff
        style(v, fg=ACCENT, size=20, name=SERIF, v="c", h="l"); v.IndentLevel = 1
        ws.Range(ws.Cells(9, c0), ws.Cells(9, c0 + 2)).Merge()
        style(ws.Cells(9, c0), fg=MUTED, size=8).Formula = note; ws.Cells(9, c0).IndentLevel = 1
        if lab in ("LOWEST CASH", "NEW FUNDING NEEDED"):
            cond = "=LowCash<MIN_CASH"
            fc = ws.Cells(7, c0).FormatConditions.Add(Type=2, Formula1=cond); fc.Font.Color = hexrgb(BAD)
    for r, hgt in ((5, 8), (6, 20), (7, 18), (8, 18), (9, 16), (10, 6), (11, 10)):
        ws.Rows(r).RowHeight = hgt
    style(ws.Range("B12"), fg=ACCENT2, size=8, bold=True).Value = "W H A T   T H E   M O D E L   S A Y S"
    ins = [
        '="•  "&ScenarioName&" case: revenue grows from "&TEXT(Rev_Y1,"[$$-409]#,##0")&" in 2027 to "&TEXT(Rev_Y3,"[$$-409]#,##0")&" in 2029; "&IF(BreakEvenM>36,"EBITDA does not stay above 0 within the plan.","EBITDA stays above 0 from "&BreakEvenTxt&" on.")',
        '="•  Cash is lowest in "&LowDateTxt&" ("&TEXT(LowCash,"[$$-409]#,##0;-[$$-409]#,##0")&"): "&IF(Funding=0,TEXT(LowCash-MIN_CASH,"[$$-409]#,##0")&" above the "&TEXT(MIN_CASH,"[$$-409]#,##0")&" the bank wants.","below the "&TEXT(MIN_CASH,"[$$-409]#,##0")&" minimum: "&TEXT(Funding,"[$$-409]#,##0")&" of new funding needed.")',
        f'="•  Pessimistic case: "&IF(Scenarios!F{SC_ROW["New funding needed"]}=0,"no new funding needed","needs "&TEXT(Scenarios!F{SC_ROW["New funding needed"]},"[$$-409]#,##0")&" of new funding (cash lowest in "&{mtxt("Scenarios!F" + str(SC_ROW["Month of the lowest cash"]))}&")")&".  Optimistic case: EBITDA 2029 "&TEXT(Scenarios!E{SC_ROW["EBITDA 2029"]},"[$$-409]#,##0")&"."']
    for k, ff in enumerate(ins):
        ws.Range(f"B{13 + k}:X{13 + k}").Merge()
        style(ws.Range(f"B{13 + k}"), size=10).Formula = ff
    ws.Rows(16).RowHeight = 10

    def chart(tl, br, kind, series, title):
        a, b = ws.Range(tl), ws.Range(br)
        co = ws.ChartObjects().Add(a.Left, a.Top, b.Left + b.Width - a.Left, b.Top + b.Height - a.Top)
        ch = co.Chart; ch.ChartType = kind
        while ch.SeriesCollection().Count:
            ch.SeriesCollection(1).Delete()
        for nm, vals, xs in series:
            se = ch.SeriesCollection().NewSeries(); se.Name = nm; se.Values = vals; se.XValues = xs
        chart_style(ch)
        ch.HasTitle = True; ch.ChartTitle.Text = title
        t = ch.ChartTitle.Format.TextFrame2.TextRange.Font; t.Size = 11; t.Bold = True; t.Fill.ForeColor.RGB = hexrgb(TEXT); t.Name = BODY
        ch.ChartTitle.Left = 8
        return co, ch

    def mrow(key):
        sh, r = ROW[key]; return W[sh].Range(f"E{r}:{LAST}{r}")
    DATES = W["P&L"].Range(f"E5:{LAST}5")
    co, ch = chart("B17", "L33", 52, [("Wholesale", mrow("pl_ws"), DATES), ("Online", mrow("pl_on"), DATES),
                                      ("EBITDA", mrow("ebitda"), DATES)], "Revenue and EBITDA by month")
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = hexrgb(ACCENT)
    ch.SeriesCollection(2).Format.Fill.ForeColor.RGB = hexrgb(S2)
    e = ch.SeriesCollection(3); e.ChartType = 4; e.Format.Line.ForeColor.RGB = hexrgb(S3); e.Format.Line.Weight = 2.25
    ch.ChartGroups(1).GapWidth = 40
    ax = ch.Axes(1); ax.CategoryType = 2; ax.TickLabels.NumberFormat = "mmm yy"; ax.TickLabelSpacing = 6; ax.TickMarkSpacing = 6
    ax.TickLabelPosition = -4134                              # labels at the bottom, below negative EBITDA
    a2 = ch.Axes(2); a2.MajorGridlines.Format.Line.ForeColor.RGB = hexrgb(LINE); a2.TickLabels.NumberFormat = "[$$-409]#,##0"
    ch.HasLegend = True; ch.Legend.Position = -4160
    co.Name = "chtRevenue"
    sh, rc = ROW["cash"]; sh2, rm = ROW["min_cash"]
    co, ch = chart("N17", "X33", 4, [("Cash at month end", W[sh].Range(f"E{rc}:{LAST}{rc}"), DATES),
                                     ("Minimum cash", W[sh2].Range(f"E{rm}:{LAST}{rm}"), DATES)], "Cash at month end vs the minimum")
    c1 = ch.SeriesCollection(1); c1.Format.Line.ForeColor.RGB = hexrgb(ACCENT); c1.Format.Line.Weight = 2.5
    c2 = ch.SeriesCollection(2); c2.Format.Line.ForeColor.RGB = hexrgb(BAD); c2.Format.Line.Weight = 1.5; c2.Format.Line.DashStyle = 4
    ax = ch.Axes(1); ax.CategoryType = 2; ax.TickLabels.NumberFormat = "mmm yy"; ax.TickLabelSpacing = 6; ax.TickMarkSpacing = 6
    ax.TickLabelPosition = -4134
    a2 = ch.Axes(2); a2.MajorGridlines.Format.Line.ForeColor.RGB = hexrgb(LINE); a2.TickLabels.NumberFormat = "[$$-409]#,##0;-[$$-409]#,##0"
    ch.HasLegend = True; ch.Legend.Position = -4160
    co.Name = "chtCash"
    # 3-year table
    ws.Rows(34).RowHeight = 10
    style(ws.Range("B35"), fg=ACCENT2, size=8, bold=True).Value = "T H R E E   Y E A R S   A T   A   G L A N C E"
    ws.Range("B36:D36").Merge(); ws.Range("B36").Value = ""
    for j, col in enumerate("FGH"):
        ws.Range(f"{col}36").Formula = f"=Annual!{'DEF'[j]}5"
    style(ws.Range("B36:H36"), fg="#FFFFFF", bg=ACCENT, bold=True); ws.Range("F36:H36").HorizontalAlignment = -4152
    TBL = [("Revenue", "a_rev", USD), ("Revenue growth", "a_growth", "+0.0%;-0.0%"), ("Gross margin", "a_gm", "0.0%"),
           ("EBITDA", "a_ebitda", USD), ("EBITDA margin", "a_em", "0.0%;-0.0%"), ("Net income", "a_ni", USD),
           ("Cash at 31 December", "a_cash", USD), ("Lowest cash in the year", "a_low", USD),
           ("Debt service cover", "a_dscr", '0.00"x";-0.00"x"'), ("Headcount", "a_heads", "0")]
    for i, (lab, key, fmt) in enumerate(TBL, 37):
        ws.Range(f"B{i}:D{i}").Merge(); ws.Range(f"B{i}").Value = lab
        for j, col in enumerate("FGH"):
            c = ws.Range(f"{col}{i}"); c.NumberFormat = fmt; c.Formula = f"=Annual!{'DEF'[j]}{AROW[key]}"
        ws.Range(f"B{i}:H{i}").Interior.Color = hexrgb(CARD if i % 2 else "#F8FAFC")
    border(ws.Range("B36:H46"), LINE, (7, 8, 9, 10, 12)); ws.Range("F37:H46").HorizontalAlignment = -4152
    ws.Range("J36:L46").Interior.Color = hexrgb(PAGE)
    style(ws.Range("B48:L48"), fg=MUTED, size=8)
    ws.Range("B48").Value = "Debt service cover = (EBITDA - tax paid) / loan payments. Cash and balance sheet items at 31 December."
    co, ch = chart("N35", "X51", 51, [("Lowest month-end cash", sc.Range(f"D{SC_ROW['Lowest month-end cash']}:F{SC_ROW['Lowest month-end cash']}"), sc.Range("D14:F14")),
                                      ("Cash at 31 Dec 2029", sc.Range(f"D{SC_ROW['Cash at 31 December 2029']}:F{SC_ROW['Cash at 31 December 2029']}"), sc.Range("D14:F14"))],
                   "Cash in the three scenarios")
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = hexrgb(S3)
    ch.SeriesCollection(2).Format.Fill.ForeColor.RGB = hexrgb(ACCENT)
    for s in (1, 2):
        se = ch.SeriesCollection(s); se.HasDataLabels = True; se.DataLabels().NumberFormat = "[$$-409]#,##0;-[$$-409]#,##0"
        se.DataLabels().Format.TextFrame2.TextRange.Font.Size = 8
    ch.ChartGroups(1).GapWidth = 60; ch.ChartGroups(1).Overlap = -10
    a2 = ch.Axes(2); a2.MajorGridlines.Format.Line.ForeColor.RGB = hexrgb(LINE); a2.TickLabels.NumberFormat = "[$$-409]#,##0;-[$$-409]#,##0"
    ch.Axes(1).TickLabelPosition = -4134
    ch.HasLegend = True; ch.Legend.Position = -4160
    co.Name = "chtScenarios"
    ws.Rows(52).RowHeight = 8
    ws.Range("B53:X53").Merge()
    style(ws.Range("B53"), fg=MUTED, size=8).Formula = (
        '="Checks: "&ChecksPassed&" / "&ChecksTotal&" PASS   ·   balance sheet balances in all 36 months   ·   Sample project, fictional company and numbers"')
    ps = ws.PageSetup; ps.PrintArea = "$A$1:$Y$54"; ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1

    # ------------------------------------------------------------------ Checks
    ws = W["Checks"]
    page(ws, "Checks", "The model checks itself. All must say PASS after any change to the assumptions.")
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = False
    ws.Columns("B").ColumnWidth = 70; ws.Columns("C").ColumnWidth = 14; ws.Columns("D").ColumnWidth = 10
    style(ws.Range("B6:D6"), fg="#FFFFFF", bg=ACCENT, bold=True).Value = ("Check", "Found", "Result")

    def RR(key, c0="E", c1=LAST, absolute=True):
        sh, r = ROW[key]; return f"'{sh}'!${c0}${r}:${c1}${r}"
    d_r = ROW["bs_diff"][1]
    re_r, ni_r = ROW["bs_re"][1], ROW["ni"][1]
    tax_r, cum_r = ROW["tax"][1], ROW["ebt_cum"][1]
    sh_s = "Schedules"
    calc_ranges = [f"'{s}'!$E$5:${LAST}${max(r for r, _ in LAYOUT[s])}" for s, *_ in MONTHLY]
    typed = "+".join(f"SUMPRODUCT(ISNUMBER({x})*NOT(ISFORMULA({x})))" for x in calc_ranges)
    n_drv = sum(1 for d in J["drivers"] if d[1])
    chk = [("Opening balance sheet balances (Assumptions)", f"={OPEN_CHECK}", "=ABS(C{r})<0.005"),
           ("Balance sheet balances: months where assets <> liabilities + equity", f"=SUMPRODUCT(--(ABS('Balance Sheet'!$E${d_r}:${LAST}${d_r})>0.005))", "=C{r}=0"),
           ("Retained earnings move by the month's net income (months that don't)",
            f"=SUMPRODUCT(--(ABS('Balance Sheet'!$E${re_r}:${LAST}${re_r}-'Balance Sheet'!$D${re_r}:$AM${re_r}-'P&L'!$E${ni_r}:${LAST}${ni_r})>0.005))", "=C{r}=0"),
           ("Cash: opening + 36 months of net change = cash at Dec 2029", f"=OPEN_CASH+SUM({RR('net')})-'Cash Flow'!${LAST}${ROW['cash'][1]}", "=ABS(C{r})<0.005"),
           ("Annual = sum of the months (revenue and net income, 3 years)",
            f"=ABS(SUM(Annual!D{AROW['a_rev']}:F{AROW['a_rev']})-SUM({RR('pl_rev')}))+ABS(SUM(Annual!D{AROW['a_ni']}:F{AROW['a_ni']})-SUM({RR('ni')}))", "=C{r}<0.005"),
           ("Equipment: net value never below 0, cost - depreciation = net",
            f"=MIN({RR('nfa')})", f"=AND(C{{r}}>=0,ABS('{sh_s}'!${LAST}${ROW['nfa'][1]}-'{sh_s}'!${LAST}${ROW['fa_cost'][1]}+'{sh_s}'!${LAST}${ROW['accdep'][1]})<0.005)"),
           ("Loan: interest + principal = payment in every month of the term",
            f"=SUMPRODUCT(--(ABS({RR('interest')}+{RR('principal')}-{RR('payment')})>0.005),--({RR('payment')}>0))", "=C{r}=0"),
           ("Loan: balance at Dec 2029 = the remaining payments, discounted (PV)",
            f"='{sh_s}'!${LAST}${ROW['loan'][1]}-IF(LOAN_M+LOAN_TERM-1-36>0,PV(LOAN_RATE/12,LOAN_M+LOAN_TERM-1-36,-PMT(LOAN_RATE/12,LOAN_TERM,-LOAN_AMT)),0)", "=ABS(C{r})<0.01"),
           ("Tax: total tax = tax rate x the highest cumulative profit (losses carried forward)",
            f"=SUM({RR('tax')})-TAX_RATE*MAX(0,MAX({RR('ebt_cum')}))", "=ABS(C{r})<0.005"),
           ("Cash flow: changes in working capital match the balance sheet",
            f"=ABS(SUM({RR('cf_ar')})-('Balance Sheet'!$D${ROW['bs_ar'][1]}-'Balance Sheet'!${LAST}${ROW['bs_ar'][1]}))+ABS(SUM({RR('cf_ap')})-('Balance Sheet'!${LAST}${ROW['bs_ap'][1]}-'Balance Sheet'!$D${ROW['bs_ap'][1]}))", "=C{r}<0.005"),
           ("No typed numbers in the monthly sheets (cells with a number that is not a formula)", "=" + typed, "=C{r}=0"),
           ("Every driver has a value in all 3 scenarios", f"=COUNT(Assumptions!$D$9:$F${DRV_LAST})", f"=C{{r}}={n_drv * 3}"),
           ("Scenario table: 'Used now' = the column of the scenario picked",
            f"=SUMPRODUCT(ABS(INDEX(Scenarios!$D$15:$F$26,0,ScenarioNo)-Scenarios!$C$15:$C$26))", "=C{r}<0.01"),
           ("Sensitivity: the 0% / 0% cell = the model (when no change is set)",
            f"=IF(AND(PriceAdj=0,VolAdj=0),Scenarios!F{S1 + 3}-EBITDA_Y3+Scenarios!F{S2r + 3}-LowCash,0)", "=ABS(C{r})<0.01")]
    for k, (txt, ff, test) in enumerate(chk, 7):
        ws.Cells(k, 2).Value = txt; ws.Cells(k, 3).Formula = ff
        ws.Cells(k, 4).Formula = f'=IF({test.format(r=k)[1:]},"PASS","FAIL")'
    last = 6 + len(chk)
    ws.Range(f"C7:C{last}").NumberFormat = "#,##0.00;-#,##0.00;0"
    d = ws.Range(f"D7:D{last}")
    fc = d.FormatConditions.Add(1, 3, '="PASS"'); fc.Font.Color = hexrgb(GOOD); fc.Interior.Color = hexrgb("#E2F0E5")
    fc = d.FormatConditions.Add(1, 3, '="FAIL"'); fc.Font.Color = hexrgb(BAD); fc.Interior.Color = hexrgb("#F6DEDA")
    d.Font.Bold = True; d.HorizontalAlignment = -4108
    border(ws.Range(f"B6:D{last}"), LINE, (7, 8, 9, 10, 12))
    ws.Range(f"B{last + 2}").Value = "Checks passed"; ws.Range(f"D{last + 2}").Formula = f'=COUNTIF(D7:D{last},"PASS")'
    ws.Range(f"B{last + 3}").Value = "Checks in total"; ws.Range(f"D{last + 3}").Value = len(chk)
    wb.Names.Add("ChecksPassed", f"=Checks!$D${last + 2}"); wb.Names.Add("ChecksTotal", f"=Checks!$D${last + 3}")

    # ------------------------------------------------------------------ How to use
    ws = W["How to use"]
    page(ws, "How to use", "No macros. Excel 2016 or later. Calculation must be Automatic (Formulas > Calculation Options).")
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = False
    ws.Columns("B").ColumnWidth = 125
    for i, t in enumerate([
            "1. Assumptions: change any blue cell. Base, Optimistic and Pessimistic sit side by side; 'Used now' is what the model reads.",
            "2. Scenarios: pick the scenario in the yellow cell; the price and volume boxes nudge the picked scenario up or down.",
            "3. Dashboard: cards, insights, both charts and the 3-year table follow at once. The scenario chart always shows all three.",
            "4. Revenue & costs, P&L, Balance Sheet, Cash Flow and Schedules: 36 months, every cell a formula. Nothing to type there.",
            "5. Checks must all say PASS. The balance sheet check row is 0.00 in every month; if a check fails, the last change broke a link.",
            "6. Not enough cash in a scenario? Type an amount and a month for 'New equity from owners' on Assumptions and watch the cash chart.",
            "Read across: a month column = one month end. Column D on Balance Sheet and Schedules = the opening balance sheet, 31 Dec 2026.",
            "Signs: on the P&L costs are negative; on the cash flow a minus means cash going out.",
    ], 6):
        ws.Cells(i, 2).Value = t

    for n in SHEETS:
        if n not in ("How to use",):
            W[n].Protect(DrawingObjects=True, Contents=True)
    W["Dashboard"].Activate(); W["Dashboard"].Range("A1").Select()
    xl.ScreenUpdating = True
    xl.Calculation = -4105
    xl.CalculateFull()
    wb.SaveAs(str(OUT), FileFormat=51)
    print("saved", OUT.name, "| checks", W["Checks"].Range(f"D{last + 2}").Value, "/", len(chk),
          "| rev 2029", round(W["Annual"].Range(f"F{AROW['a_rev']}").Value), "| low cash", round(wb.Names("LowCash").RefersToRange.Value), "| scen", [sc.Cells(r, c).Value for r in (22, 25) for c in (4, 5, 6)])
    for k in range(7, last + 1):
        print("  ", W["Checks"].Cells(k, 4).Value, W["Checks"].Cells(k, 3).Value, W["Checks"].Cells(k, 2).Value)
    wb.Close(False)
finally:
    xl.Quit()
