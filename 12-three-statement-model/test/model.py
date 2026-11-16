"""The whole 3-statement model again, in plain Python, written from the rules in the README (not from the workbook's
formulas). verify.py feeds it the assumptions read from the workbook and compares every month.

    run(A, season, O, price_adj=0, vol_adj=0) -> dict of 36-month lists (plus a few single values)
    A: live assumptions {name: value}; season: 12 factors; O: opening balance sheet {name: value}
"""
import calendar
from datetime import date

N = 36
START = date(2027, 1, 1)


def month_end(m):                       # m = 1..36
    y, mo = 2027 + (m - 1) // 12, (m - 1) % 12 + 1
    return date(y, mo, calendar.monthrange(y, mo)[1])


def pmt(r, n, pv):
    return pv * r / (1 - (1 + r) ** -n) if r else pv / n


def run(A, season, O, price_adj=0.0, vol_adj=0.0):
    out = {k: [0.0] * N for k in (
        "ws_kg", "ws_price", "ws_rev", "on_orders", "on_aov", "on_rev", "rev", "ws_cogs", "on_cogs", "cogs", "gross",
        "heads", "payroll", "rent", "mkt", "fees", "other", "opex", "ebitda", "capex", "dep", "ebit", "loan_open", "draw",
        "interest", "payment", "principal", "loan", "ebt", "tax", "tax_paid", "tax_payable", "ni", "ar", "inv", "purch",
        "ap", "fa_cost", "accdep", "nfa", "share", "re", "cfo", "cfi", "cff", "net_cash", "cash_open", "cash",
        "assets", "liab", "equity", "diff", "equity_in")}
    loan_bal, cum_ebt, cum_tax, payable = 0.0, 0.0, 0.0, O["OPEN_TAX"]
    ar0, inv0, ap0 = O["OPEN_AR"], O["OPEN_INV"], O["OPEN_AP"]
    fa, acc, share, re, cash = O["OPEN_FA_COST"], O["OPEN_ACCDEP"], O["OPEN_SHARE"], O["OPEN_RE"], O["OPEN_CASH"]
    open_nbv = O["OPEN_FA_COST"] - O["OPEN_ACCDEP"]
    pay = pmt(A["LOAN_RATE"] / 12, A["LOAN_TERM"], A["LOAN_AMT"])
    for i in range(N):
        m = i + 1
        d = month_end(m)
        year = (m - 1) // 12 + 1
        s = season[d.month - 1]
        idx = (1 + A["PRICE_RISE"]) ** (year - 1)
        o = lambda k, v: out[k].__setitem__(i, v)
        o("ws_kg", A["WS_KG0"] * (1 + A["WS_GROWTH"]) ** (m - 1) * s * (1 + vol_adj))
        o("ws_price", A["WS_PRICE"] * idx * (1 + price_adj))
        o("ws_rev", out["ws_kg"][i] * out["ws_price"][i])
        o("on_orders", A["ON_ORD0"] * (1 + A["ON_GROWTH"]) ** (m - 1) * s * (1 + vol_adj))
        o("on_aov", A["ON_AOV"] * idx * (1 + price_adj))
        o("on_rev", out["on_orders"][i] * out["on_aov"][i])
        o("rev", out["ws_rev"][i] + out["on_rev"][i])
        o("ws_cogs", out["ws_rev"][i] * A["WS_COGS"]); o("on_cogs", out["on_rev"][i] * A["ON_COGS"])
        o("cogs", out["ws_cogs"][i] + out["on_cogs"][i])
        o("gross", out["rev"][i] - out["cogs"][i])
        o("heads", A[f"HEAD_Y{year}"])
        o("payroll", out["heads"][i] * A["PAY_HEAD"] * (1 + A["PAY_RISE"]) ** (year - 1))
        o("rent", A["RENT"]); o("mkt", out["on_rev"][i] * A["MKT"]); o("fees", out["on_rev"][i] * A["ON_FEES"])
        o("other", A["OTHER_OPEX"] * (1 + A["COST_INFL"]) ** (year - 1))
        o("opex", sum(out[k][i] for k in ("payroll", "rent", "mkt", "fees", "other")))
        o("ebitda", out["gross"][i] - out["opex"][i])
        # capital spending and depreciation (from the month bought, straight line)
        capex = dep = 0.0
        for k in (1, 2, 3):
            mb, amt, life = A[f"CAPEX{k}_M"], A[f"CAPEX{k}_AMT"], A[f"CAPEX{k}_LIFE"]
            if m == mb: capex += amt
            if mb <= m < mb + life * 12: dep += amt / (life * 12)
        od = min(O["OPEN_DEP"], open_nbv); open_nbv -= od; dep += od
        o("capex", capex); o("dep", dep)
        o("ebit", out["ebitda"][i] - dep)
        # loan: drawn at the start of a month, annuity paid at the end of each month of the term
        o("loan_open", loan_bal)
        draw = A["LOAN_AMT"] if m == A["LOAN_M"] else 0.0
        interest = (loan_bal + draw) * A["LOAN_RATE"] / 12
        payment = pay if A["LOAN_M"] <= m < A["LOAN_M"] + A["LOAN_TERM"] else 0.0
        principal = payment - interest if payment else 0.0
        loan_bal = loan_bal + draw - principal
        o("draw", draw); o("interest", interest); o("payment", payment); o("principal", principal); o("loan", loan_bal)
        # tax: on the profit made so far, losses carried forward; paid in Jan / Apr / Jul / Oct for the quarter before
        ebt = out["ebit"][i] - interest
        cum_ebt += ebt
        tax = max(0.0, A["TAX_RATE"] * max(0.0, cum_ebt) - cum_tax)
        cum_tax += tax
        paid = payable if d.month in (1, 4, 7, 10) else 0.0
        payable = payable + tax - paid
        o("ebt", ebt); o("tax", tax); o("tax_paid", paid); o("tax_payable", payable)
        o("ni", ebt - tax)
        # working capital
        ar = out["ws_rev"][i] * A["DSO"] / 30
        inv = out["cogs"][i] * A["DIO"] / 30
        purch = out["cogs"][i] + inv - inv0
        ap = purch * A["DPO"] / 30
        o("ar", ar); o("inv", inv); o("purch", purch); o("ap", ap)
        eq_in = A["EQUITY_AMT"] if m == A["EQUITY_M"] else 0.0
        o("equity_in", eq_in)
        # cash flow, indirect method
        cfo = out["ni"][i] + dep - (ar - ar0) - (inv - inv0) + (ap - ap0) + (tax - paid)
        cfi = -capex
        cff = draw - principal + eq_in
        o("cfo", cfo); o("cfi", cfi); o("cff", cff); o("net_cash", cfo + cfi + cff)
        o("cash_open", cash); cash += cfo + cfi + cff; o("cash", cash)
        fa += capex; acc += dep; share += eq_in; re += out["ni"][i]
        o("fa_cost", fa); o("accdep", acc); o("nfa", fa - acc); o("share", share); o("re", re)
        o("assets", cash + ar + inv + fa - acc); o("liab", ap + payable + loan_bal); o("equity", share + re)
        o("diff", out["assets"][i] - out["liab"][i] - out["equity"][i])
        ar0, inv0, ap0 = ar, inv, ap
    yr = lambda k, y: sum(out[k][(y - 1) * 12: y * 12])
    out["year"] = {y: {k: yr(k, y) for k in ("rev", "ws_rev", "on_rev", "cogs", "gross", "opex", "ebitda", "dep", "ebit",
                                              "interest", "ebt", "tax", "ni", "cfo", "cfi", "cff", "tax_paid", "payment",
                                              "principal", "capex")} for y in (1, 2, 3)}
    for y in (1, 2, 3):
        Y = out["year"][y]
        Y["cash_end"] = out["cash"][y * 12 - 1]
        Y["dscr"] = (Y["ebitda"] - Y["tax_paid"]) / Y["payment"] if Y["payment"] else None
    lows = min(out["cash"])
    out["low_cash"] = lows
    out["low_month"] = out["cash"].index(lows) + 1
    neg = [m + 1 for m in range(N) if out["ebitda"][m] <= 0]          # from this month on EBITDA stays above 0
    out["breakeven"] = (neg[-1] + 1 if neg else 1) if not neg or neg[-1] < N else None
    out["funding"] = max(0.0, A["MIN_CASH"] - lows)
    out["below_min"] = sum(1 for c in out["cash"] if c < A["MIN_CASH"])
    return out


if __name__ == "__main__":                      # quick look at the three scenarios from data/assumptions.json
    import json, pathlib
    J = json.loads((pathlib.Path(__file__).resolve().parent.parent / "data" / "assumptions.json").read_text())
    O = {k: v for k, _, v in J["opening"]}
    for s, nm in enumerate(("Base", "Optimistic", "Pessimistic")):
        A = {r[0]: r[4][s] for r in J["drivers"] if r[1]}
        r = run(A, J["season"], O)
        print(nm, {y: (round(r["year"][y]["rev"]), round(r["year"][y]["ebitda"]), round(r["year"][y]["ni"]),
                       round(r["year"][y]["cash_end"]), r["year"][y]["dscr"] and round(r["year"][y]["dscr"], 2)) for y in (1, 2, 3)},
              "low", round(r["low_cash"]), "m", r["low_month"], "be", r["breakeven"], "fund", round(r["funding"]),
              "maxdiff", max(abs(x) for x in r["diff"]), "loan", round(r["loan"][-1]))
