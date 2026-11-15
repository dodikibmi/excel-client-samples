"""Summary dashboard and Calc sheet of Two-Source-Reconciliation.xlsx (dark luxury style). Called by build_workbook.py.

Summary: account + month pickers, 4 KPI cards with icons, donut (matched vs exceptions), exception amount by status,
6-month trend (auto-match % and open difference), aging heat grid, matches per rule, bank reconciliation statement
(must end at 0.00), 8 data checks and a proof-steps / audit-trail box. Every number is an Excel formula over the
Power Query tables Recon, Workbench, BankControls, CleaningLog, RunInfo and the Settings tables, so it follows a Refresh.
"""
import lux
from lux import c, fill, font

STATUSES = ["Amount differs", "Only in bank", "Only in ledger", "Duplicate in bank", "Duplicate in ledger"]
RULES = [("Reference", "1  Reference + same amount"), ("Reference ± amount", None), ("Amount + date", None)]
MONTHS = 'CHOOSE(MONTH({0}),"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")'


def statement_parts(a, s, e, ob, obk):
    """Formulas of the bank reconciliation statement for account a, month s..e (cumulative from the first file)."""
    A = f"Recon[Account],{a}"
    both = f'{A},Recon[Bank date],"<="&{e},Recon[Ledger date],"<="&{e}'

    def transit(sign):
        crit = f'{A},Recon[Ledger date],"<="&{e},Recon[Ledger amount],"{sign}0",Recon[Status],"<>Duplicate in ledger"'
        return f'SUMIFS(Recon[Ledger amount],{crit})-SUMIFS(Recon[Ledger amount],{crit},Recon[Bank date],"<="&{e})'
    return dict(
        bank_open=f'{ob}+SUMIFS(Recon[Bank amount],{A},Recon[Bank date],"<"&{s})',
        bank_moves=f'SUMIFS(Recon[Bank amount],{A},Recon[Bank date],">="&{s},Recon[Bank date],"<="&{e})',
        in_transit=transit(">"),
        outstanding=transit("<"),
        book_open=f'{obk}+SUMIFS(Recon[Ledger amount],{A},Recon[Ledger date],"<"&{s})',
        book_moves=f'SUMIFS(Recon[Ledger amount],{A},Recon[Ledger date],">="&{s},Recon[Ledger date],"<="&{e})',
        bank_items=f'SUMIFS(Recon[Bank amount],{A},Recon[Bank date],"<="&{e})-SUMIFS(Recon[Bank amount],{both})',
        differences=f'SUMIFS(Recon[Bank amount],{both})-SUMIFS(Recon[Ledger amount],{both})',
        twice=f'-SUMIFS(Recon[Ledger amount],{A},Recon[Status],"Duplicate in ledger",Recon[Ledger date],"<="&{e})')


def name(wb, nm, ref):
    wb.Names.Add(nm, ref)


def build_calc(wb):
    ws = wb.Worksheets("Calc")
    ws.Cells.Font.Name = lux.SANS
    ws.Range("A1").Value = "Calculation area: feeds the Summary charts and checks. Nothing to type here."
    ws.Range("A1").Font.Bold = True
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 14
    # the pickers live on the Summary sheet
    sm = wb.Worksheets("Summary")
    sm.Range("C7").Value = "Operating"
    name(wb, "SelAccount", "=Summary!$C$7")
    name(wb, "SelMonth", "=Summary!$I$7")
    sm.Range("I7").NumberFormat = "mmm yyyy"
    rows = [
        ("MStart", "=DATE(YEAR(SelMonth),MONTH(SelMonth),1)", "d mmm yyyy"),
        ("MEnd", "=EOMONTH(MStart,0)", "d mmm yyyy"),
        ("PStart", "=EDATE(MStart,-1)", "d mmm yyyy"),
        ("PEnd", "=MStart-1", "d mmm yyyy"),
        ("OpenBank", "=SUMIFS(Accounts[Opening bank],Accounts[Account],SelAccount)", "#,##0.00"),
        ("OpenBook", "=SUMIFS(Accounts[Opening cash book],Accounts[Account],SelAccount)", "#,##0.00"),
        ("LatestMonth", "=MAX(Recon[Month])", "mmm yyyy"),
        ("BankLinesM", '=COUNTIFS(Recon[Account],SelAccount,Recon[Bank date],">="&MStart,Recon[Bank date],"<="&MEnd)', "0"),
        ("MatchedM", '=COUNTIFS(Recon[Account],SelAccount,Recon[Bank date],">="&MStart,Recon[Bank date],"<="&MEnd,Recon[Status],"Matched")', "0"),
        ("MatchRate", "=IF(BankLinesM=0,0,MatchedM/BankLinesM)", "0.0%"),
        ("BankLinesP", '=COUNTIFS(Recon[Account],SelAccount,Recon[Bank date],">="&PStart,Recon[Bank date],"<="&PEnd)', "0"),
        ("MatchedP", '=COUNTIFS(Recon[Account],SelAccount,Recon[Bank date],">="&PStart,Recon[Bank date],"<="&PEnd,Recon[Status],"Matched")', "0"),
        ("MatchRateP", "=IF(BankLinesP=0,0,MatchedP/BankLinesP)", "0.0%"),
        ("MatchedAmt", '=SUMIFS(Recon[Value],Recon[Account],SelAccount,Recon[Bank date],">="&MStart,Recon[Bank date],"<="&MEnd,Recon[Status],"Matched")', "#,##0.00"),
        ("MatchedAmtP", '=SUMIFS(Recon[Value],Recon[Account],SelAccount,Recon[Bank date],">="&PStart,Recon[Bank date],"<="&PEnd,Recon[Status],"Matched")', "#,##0.00"),
        ("OpenExc", '=COUNTIFS(Workbench[Account],SelAccount,Workbench[Date],"<="&MEnd,Workbench[Resolved date],"")', "0"),
        ("ResolvedExc", '=COUNTIFS(Workbench[Account],SelAccount,Workbench[Date],"<="&MEnd,Workbench[Resolved date],"<>")', "0"),
        ("OldExc", '=COUNTIFS(Workbench[Account],SelAccount,Workbench[Date],"<"&MEnd-30,Workbench[Resolved date],"")', "0"),
        ("OpenAmt", '=SUMIFS(Workbench[Abs amount],Workbench[Account],SelAccount,Workbench[Date],"<="&MEnd,Workbench[Resolved date],"")', "#,##0.00"),
        ("ExcInMonth", '=COUNTIFS(Recon[Account],SelAccount,Recon[Month],MStart,Recon[Status],"<>Matched")', "0"),
        ("MonthLabel", "=" + MONTHS.format("MStart") + '&" "&YEAR(MStart)', "@"),
        ("PrevLabel", "=" + MONTHS.format("PStart"), "@"),
        ("MatchedRows", '=COUNTIFS(Recon[Account],SelAccount,Recon[Month],MStart,Recon[Status],"Matched")', "0"),
    ]
    for i, (nm, f, fmt) in enumerate(rows, 3):
        ws.Cells(i, 1).Value = nm
        ws.Cells(i, 2).Formula = f
        if fmt != "@":
            ws.Cells(i, 2).NumberFormat = fmt
        name(wb, nm, f"=Calc!$B${i}")

    # donut: lines of the month (D3:E5)
    ws.Range("D2").Value = "Donut: lines this month"; ws.Range("D2").Font.Bold = True
    donut = [("Matched on reference", '=COUNTIFS(Recon[Account],SelAccount,Recon[Month],MStart,Recon[Match rule],"Reference")'
              '+COUNTIFS(Recon[Account],SelAccount,Recon[Month],MStart,Recon[Match rule],"Reference ± amount")'),
             ("Matched on amount + date", '=COUNTIFS(Recon[Account],SelAccount,Recon[Month],MStart,Recon[Match rule],"Amount + date")'),
             ("Exceptions", "=ExcInMonth")]
    for i, (lab, f) in enumerate(donut, 3):
        ws.Cells(i, 4).Value = lab
        ws.Cells(i, 5).Formula = f
    # exception amount by status (D8:E12)
    ws.Range("D7").Value = "Bar: exception amount by status, this month"; ws.Range("D7").Font.Bold = True
    for i, st in enumerate(STATUSES, 8):
        ws.Cells(i, 4).Value = st
        ws.Cells(i, 5).Formula = f'=SUMIFS(Recon[Value],Recon[Account],SelAccount,Recon[Month],MStart,Recon[Status],D{i})'
    ws.Range("E8:E12").NumberFormat = "#,##0.00"
    # 6-month trend (C15:G20)
    ws.Range("C14:G14").Value = ("Month", "Label", "Auto-match %", "Open difference", "Bank lines")
    ws.Range("C14:G14").Font.Bold = True
    for j in range(6):
        r = 15 + j
        ws.Cells(r, 3).Formula = f"=EDATE(LatestMonth,{j - 5})"
        ws.Cells(r, 4).Formula = "=" + MONTHS.format(f"C{r}")
        bank = f'Recon[Account],SelAccount,Recon[Bank date],">="&C{r},Recon[Bank date],"<="&EOMONTH(C{r},0)'
        ws.Cells(r, 7).Formula = f"=COUNTIFS({bank})"
        ws.Cells(r, 5).Formula = f'=IF(G{r}=0,0,COUNTIFS({bank},Recon[Status],"Matched")/G{r})'
        ws.Cells(r, 6).Formula = (f'=SUMIFS(Workbench[Abs amount],Workbench[Account],SelAccount,Workbench[Date],">="&C{r},'
                                  f'Workbench[Date],"<="&EOMONTH(C{r},0),Workbench[Resolved date],"")')
    ws.Range("C15:C20").NumberFormat = "mmm yyyy"
    ws.Range("E15:E20").NumberFormat = "0.0%"
    ws.Range("F15:F20").NumberFormat = "#,##0.00"
    # month list for the picker (I3:I8)
    ws.Range("I2").Value = "Month picker list"; ws.Range("I2").Font.Bold = True
    for j in range(6):
        ws.Cells(3 + j, 9).Formula = f"=EDATE(LatestMonth,{j - 5})"
    ws.Range("I3:I8").NumberFormat = "mmm yyyy"
    name(wb, "MonthList", "=Calc!$I$3:$I$8")
    # every account x every month must balance to 0.00 (K3:Q26; up to 4 accounts)
    ws.Range("K2:Q2").Value = ("Account", "Month", "Month end", "Adjusted bank", "Adjusted cash book", "Difference", "Fails")
    ws.Range("K2:Q2").Font.Bold = True
    r = 3
    for i in range(1, 5):
        for j in range(6):
            ws.Cells(r, 11).Formula = f'=IFERROR(INDEX(Accounts[Account],{i}),"")&""'
            ws.Cells(r, 12).Formula = f"=I{3 + j}"
            ws.Cells(r, 13).Formula = f"=EOMONTH(L{r},0)"
            ob = f"SUMIFS(Accounts[Opening bank],Accounts[Account],K{r})"
            obk = f"SUMIFS(Accounts[Opening cash book],Accounts[Account],K{r})"
            p = statement_parts(f"K{r}", f"L{r}", f"M{r}", ob, obk)
            ws.Cells(r, 14).Formula = f'=IF(K{r}="","",{p["bank_open"]}+{p["bank_moves"]}+{p["in_transit"]}+{p["outstanding"]})'
            ws.Cells(r, 15).Formula = f'=IF(K{r}="","",{p["book_open"]}+{p["book_moves"]}+{p["bank_items"]}+{p["differences"]}+{p["twice"]})'
            ws.Cells(r, 16).Formula = f'=IF(K{r}="","",ROUND(N{r}-O{r},2))'
            ws.Cells(r, 17).Formula = f'=IF(K{r}="",0,IF(P{r}=0,0,1))'
            r += 1
    ws.Range("L3:M26").NumberFormat = "mmm yyyy"
    ws.Range("N3:P26").NumberFormat = "#,##0.00"
    name(wb, "PeriodsChecked", '=Calc!$S$3')
    name(wb, "PeriodsFailed", '=Calc!$S$4')
    ws.Range("R3").Value = "Account-months"; ws.Range("S3").Formula = '=COUNTIF(K3:K26,"?*")'
    ws.Range("R4").Value = "Not balanced"; ws.Range("S4").Formula = "=SUM(Q3:Q26)"
    for col, w in {"C": 10, "D": 24, "E": 13, "F": 15, "G": 10, "I": 11, "K": 11, "L": 10, "M": 11, "N": 14, "O": 16, "P": 11, "R": 15}.items():
        ws.Columns(col).ColumnWidth = w
    wb.Application.Calculate()
    sm.Range("I7").Value = ws.Range("B9").Value  # default: the latest month


def text(ws, rng, value, color="ivory", size=10, bold=False, name=None, align=None, merge=False, fmt=None, indent=None):
    r = ws.Range(rng)
    if merge and r.Cells.Count > 1:
        r.Merge()
    cell = r.Cells(1, 1)
    if isinstance(value, str) and value.startswith("="):
        cell.Formula = value
    elif value is not None:
        cell.Value = value
    font(r, color, size, bold, name)
    if align:
        r.HorizontalAlignment = {"l": -4131, "c": -4108, "r": -4152}[align]
    r.VerticalAlignment = -4108
    if fmt:
        r.NumberFormat = fmt
    if indent:
        cell.IndentLevel = indent
    return r


def build_summary(wb):
    ws = wb.Worksheets("Summary")
    calc = wb.Worksheets("Calc")
    xl = wb.Application
    lux.page(ws, "AA", 80)
    for nm, cell in (("StUnexplained", "K48"), ("StAdjBank", "K38"), ("StAdjBook", "K46"), ("StBankClose", "K35")):
        name(wb, nm, f"=Summary!${cell[0]}${cell[1:]}")
    widths = {"A": 2}
    for blk in ("BCDEF", "HIJKL", "NOPQR", "TUVWX"):
        for col in blk:
            widths[col] = 7.6
    widths.update({"G": 1.6, "M": 1.6, "S": 1.6, "Y": 2})
    for col, w in widths.items():
        ws.Columns(col).ColumnWidth = w

    lux.header(ws, "Bank reconciliation",
               "Bank statements vs cash book, matched in Power Query with 3 rules. Fictional data. New month: drop the CSV exports in the data folder, Data > Refresh All.",
               last_col="Y", nav_current="Summary")

    # pickers (row 7) ---------------------------------------------------------------------------------------------
    ws.Rows(7).RowHeight = 24
    text(ws, "B7", "ACCOUNT", "gold", 8, True)
    text(ws, "H7", "MONTH", "gold", 8, True, align="r")
    for rng, cell in (("C7:E7", "C7"), ("I7:K7", "I7")):
        r = ws.Range(rng)
        r.Merge()
        fill(r, "card2")
        font(r, "ivory", 11, True)
        r.HorizontalAlignment = -4131
        r.Cells(1, 1).IndentLevel = 1
        r.VerticalAlignment = -4108
        for e in (7, 8, 9, 10):
            r.Borders(e).LineStyle = 1; r.Borders(e).Color = c("gold")
    for cell in ("F7", "L7"):
        text(ws, cell, "▾", "gold", 12, True, align="l")
    ws.Range("C7").Validation.Add(3, 1, 1, '=INDIRECT("Accounts[Account]")')
    ws.Range("I7").Validation.Add(3, 1, 1, "=MonthList")
    text(ws, "N7", "Every card, chart and table below follows the account and month you pick. Status as of the last refresh.", "muted", 9)

    # KPI cards (rows 9-13) ---------------------------------------------------------------------------------------
    ws.Rows(8).RowHeight = 10
    for r_, h in ((9, 6), (10, 16), (11, 34), (12, 16), (13, 6), (14, 12)):
        ws.Rows(r_).RowHeight = h
    cards = [("BCDEF", "MATCH RATE", "✓", "=MatchRate", "0.0%",
              '=MatchedM&" of "&BankLinesM&" bank lines matched by the rules"', "good"),
             ("HIJKL", "MATCHED AMOUNT", "$", "=MatchedAmt", "#,##0",
              '=IF(MatchedAmtP=0,"gross value of matched lines",IF(MatchedAmt>=MatchedAmtP,"▲ ","▼ ")&TEXT(ABS(MatchedAmt/MatchedAmtP-1),"0.0%")&" vs "&PrevLabel)', None),
             ("NOPQR", "OPEN EXCEPTIONS", "!", "=OpenExc", "0",
              '=ResolvedExc&" resolved  ·  "&OldExc&" open over 30 days"', None),
             ("TUVWX", "UNEXPLAINED DIFFERENCE", "=", "=StUnexplained", "#,##0.00;-#,##0.00;0.00",
              '=IF(StUnexplained=0,"adjusted bank = adjusted cash book","check the statement below")', None)]
    for cols, label, glyph, big, fmt, note, note_col in cards:
        a, b = cols[0], cols[-1]
        box = ws.Range(f"{a}9:{b}13")
        fill(box, "card")
        for e in (7, 8, 9, 10):
            box.Borders(e).LineStyle = 1; box.Borders(e).Color = c("line")
        text(ws, f"{a}10:{b}10", label, "muted", 8, True, indent=2)
        text(ws, f"{a}11:{b}11", big, "ivory", 24, False, lux.SERIF, "c", True, fmt)
        text(ws, f"{a}12:{b}12", note, note_col or "muted", 8.5, False, None, "c", True)
        top, left = ws.Range(f"{a}9").Top, ws.Range(f"{a}9").Left
        lux.bar(ws, left + 1, top + 4, ws.Range(f"{a}9:{a}13").Height - 8)
        right = ws.Range(f"{b}9").Left + ws.Range(f"{b}9").Width
        lux.icon(ws, right - 30, top + 6, glyph)
    # note colours: green better / red worse
    def note_cf(cell, good_formula, bad_formula, good="good", bad="bad"):
        fc = ws.Range(cell).FormatConditions.Add(2, None, good_formula); fc.Font.Color = c(good)
        fc = ws.Range(cell).FormatConditions.Add(2, None, bad_formula); fc.Font.Color = c(bad)
    note_cf("H12", "=MatchedAmt>=MatchedAmtP", "=MatchedAmt<MatchedAmtP")
    note_cf("N12", "=OldExc=0", "=OldExc>0", "good", "amber")
    note_cf("T12", "=StUnexplained=0", "=StUnexplained<>0")
    note_cf("T11", "=StUnexplained=0", "=StUnexplained<>0")

    # charts (rows 15-30) ----------------------------------------------------------------------------------------
    for r_ in range(15, 31):
        ws.Rows(r_).RowHeight = 15
    ws.Rows(31).RowHeight = 12
    top = ws.Range("B15").Top
    height = ws.Range("B15:B30").Height
    x0 = ws.Range("B15").Left
    x1 = ws.Range("X15").Left + ws.Range("X15").Width
    gap = ws.Range("G15").Width
    w = (x1 - x0 - 2 * gap) / 3
    donut_chart(ws, calc, x0, top, w, height)
    status_chart(ws, calc, x0 + w + gap, top, w, height)
    trend_chart(ws, calc, x0 + 2 * (w + gap), top, w, height)

    # bank reconciliation statement (B32:L49) ----------------------------------------------------------------------
    for r_ in range(32, 63):
        ws.Rows(r_).RowHeight = 17
    ws.Rows(32).RowHeight = 28
    ws.Rows(50).RowHeight = 12
    ws.Rows(51).RowHeight = 12
    lux.panel(ws, "B32:L49", "Bank reconciliation statement")
    text(ws, "I32:L32", '=SelAccount&"  ·  "&MonthLabel', "gold2", 9, False, None, "r", True, indent=1)
    p = statement_parts("SelAccount", "MStart", "MEnd", "OpenBank", "OpenBook")
    lines = [
        ('="Balance per bank statement, "&TEXT(MStart,"d mmm yyyy")', "=" + p["bank_open"], False),
        ("Movements on the bank statement", "=" + p["bank_moves"], False),
        ('="Balance per bank statement, "&TEXT(MEnd,"d mmm yyyy")', "=K33+K34", True),
        ("Add: deposits in transit", "=" + p["in_transit"], False),
        ("Less: outstanding payments", "=" + p["outstanding"], False),
        ("Adjusted bank balance", "=K35+K36+K37", True),
        (None, None, False),
        ('="Balance per cash book, "&TEXT(MStart,"d mmm yyyy")', "=" + p["book_open"], False),
        ("Movements in the cash book", "=" + p["book_moves"], False),
        ('="Balance per cash book, "&TEXT(MEnd,"d mmm yyyy")', "=K40+K41", True),
        ("Add/less: bank items not yet in the cash book", "=" + p["bank_items"], False),
        ("Add/less: amount differences (incl. within tolerance)", "=" + p["differences"], False),
        ("Less: postings made twice", "=" + p["twice"], False),
        ("Adjusted cash book balance", "=K42+K43+K44+K45", True),
        (None, None, False),
        ("Unexplained difference (must be 0.00)", "=ROUND(K38-K46,2)", True),
    ]
    for i, (label, f, bold) in enumerate(lines, 33):
        if not label:
            continue
        text(ws, f"B{i}", label, "ivory" if bold else "#C9C3B7", 10, bold, indent=1)
        text(ws, f"K{i}:L{i}", f, "ivory", 10, bold, None, "r", True, "#,##0.00;-#,##0.00;0.00", indent=1)
    for r_ in (38, 46):
        rng = ws.Range(f"B{r_}:L{r_}")
        rng.Borders(8).LineStyle = 1; rng.Borders(8).Color = c("gold"); rng.Borders(8).Weight = 2
    tot = ws.Range("B48:L48")
    fill(tot, "#2B2416")
    font(tot, "gold2", 11, True)
    note_cf("K48", "=K48=0", "=K48<>0")

    # aging heat grid (N32:X40) ------------------------------------------------------------------------------------
    lux.panel(ws, "N32:X40", "Aging of open exceptions")
    text(ws, "U32:X32", '="as at "&TEXT(MEnd,"d mmm yyyy")', "gold2", 9, False, None, "r", True, indent=1)
    heads = [("N33", "Status", "l"), ("Q33:R33", "0–7 days", "c"), ("T33:U33", "8–30 days", "c"), ("V33:W33", "over 30", "c"), ("X33", "Total", "c")]
    for rng, t, al in heads:
        text(ws, rng, t, "gold", 9, True, None, al, True, indent=1 if al == "l" else None)
    ws.Range("N33:X33").Borders(9).LineStyle = 1; ws.Range("N33:X33").Borders(9).Color = c("line")
    buckets = [("Q", '">="&MEnd-7', '"<="&MEnd'), ("T", '">="&MEnd-30', '"<"&MEnd-7'), ("V", '">="&0', '"<"&MEnd-30')]
    for i, st in enumerate(STATUSES, 34):
        text(ws, f"N{i}", st, "#C9C3B7", 9.5, indent=1)
        for col, lo_, hi in buckets:
            nxt = chr(ord(col) + 1)
            f = (f'=COUNTIFS(Workbench[Account],SelAccount,Workbench[Status],$N{i},Workbench[Resolved date],"",'
                 f'Workbench[Date],{lo_},Workbench[Date],{hi})')
            text(ws, f"{col}{i}:{nxt}{i}", f, "ivory", 10, True, None, "c", True, '0;-0;"–"')
        text(ws, f"X{i}", f"=Q{i}+T{i}+V{i}", "gold2", 10, True, None, "c", False, '0;-0;"–"')
    text(ws, "N39", "Open amount", "muted", 9, True, indent=1)
    for col, lo_, hi in buckets:
        nxt = chr(ord(col) + 1)
        f = (f'=SUMIFS(Workbench[Abs amount],Workbench[Account],SelAccount,Workbench[Resolved date],"",'
             f'Workbench[Date],{lo_},Workbench[Date],{hi})')
        text(ws, f"{col}39:{nxt}39", f, "muted", 9, False, None, "c", True, '#,##0;-#,##0;"–"')
    text(ws, "X39", "=Q39+T39+V39", "muted", 9, True, None, "c", False, '#,##0;-#,##0;"–"')
    heat = ws.Range("Q34:W38")
    cs = heat.FormatConditions.AddColorScale(3)
    for k, (typ, val, col) in enumerate([(0, 0, "card"), (0, 1, "#7A5A2A"), (0, 4, "#B5523F")], 1):
        crit = cs.ColorScaleCriteria(k)
        crit.Type = typ
        crit.Value = val
        crit.FormatColor.Color = c(col)

    # matches per rule (N42:X49) -----------------------------------------------------------------------------------
    lux.panel(ws, "N42:X49", "Matches per rule")
    text(ws, "T42:X42", '="tolerance ±"&TEXT(AmountTolerance,"0.00")&"  ·  ±"&DaysTolerance&" days"', "gold2", 9, False, None, "r", True, indent=1)
    text(ws, "N43", "Rule", "gold", 9, True, indent=1)
    text(ws, "T43:U43", "=MonthLabel", "gold", 9, True, None, "c", True)
    text(ws, "V43:W43", "6 months", "gold", 9, True, None, "c", True)
    text(ws, "X43", "", "gold", 9, True)
    ws.Range("N43:X43").Borders(9).LineStyle = 1; ws.Range("N43:X43").Borders(9).Color = c("line")
    labels = ["1  Reference + same amount", '="2  Reference, amount within ±"&TEXT(AmountTolerance,"0.00")',
              '="3  Amount + date (±"&DaysTolerance&" days), no reference"']
    six = '">="&EDATE(LatestMonth,-5)'
    for i, ((rule, _), lab) in enumerate(zip(RULES, labels), 44):
        text(ws, f"N{i}", lab, "#C9C3B7", 9.5, indent=1)
        text(ws, f"T{i}:U{i}", f'=COUNTIFS(Recon[Account],SelAccount,Recon[Month],MStart,Recon[Match rule],"{rule}")',
             "ivory", 10, True, None, "c", True, "#,##0")
        text(ws, f"V{i}:W{i}", f'=COUNTIFS(Recon[Account],SelAccount,Recon[Month],{six},Recon[Match rule],"{rule}")',
             "ivory", 10, True, None, "c", True, "#,##0")
    text(ws, "N47", "Matched, all rules", "ivory", 10, True, indent=1)
    text(ws, "T47:U47", "=SUM(T44:T46)", "gold2", 10, True, None, "c", True, "#,##0")
    text(ws, "V47:W47", "=SUM(V44:V46)", "gold2", 10, True, None, "c", True, "#,##0")
    text(ws, "N48", "Not matched (exceptions)", "#C9C3B7", 9.5, indent=1)
    text(ws, "T48:U48", "=ExcInMonth", "amber", 10, True, None, "c", True, "#,##0")
    text(ws, "V48:W48", f'=COUNTIFS(Recon[Account],SelAccount,Recon[Month],{six},Recon[Status],"<>Matched")',
         "amber", 10, True, None, "c", True, "#,##0")
    ws.Range("N47:X47").Borders(8).LineStyle = 1; ws.Range("N47:X47").Borders(8).Color = c("line")
    for rng in ("T44:U46", "V44:W46"):
        db = ws.Range(rng).FormatConditions.AddDatabar()
        db.BarColor.Color = c("#6B5730")
        db.BarFillType = 0  # solid
        db.ShowValue = True

    # data checks (B52:L62) ----------------------------------------------------------------------------------------
    lux.panel(ws, "B52:L62", "Data checks")
    RI = lambda item: f'INDEX(RunInfo[Value],MATCH("{item}",RunInfo[Item],0))'
    checks = [
        ("Every statement file: opening + movements = printed closing", '=IF(COUNTIF(BankControls[Inside the file],"BREAK")=0,"PASS","FAIL")'),
        ("Every month starts where the previous one closed", '=IF(COUNTIF(BankControls[Chain],"BREAK")=0,"PASS","FAIL")'),
        ("Printed closing balance (picked month) = statement closing",
         '=IF(ROUND(SUMIFS(BankControls[Printed closing],BankControls[Account],SelAccount,BankControls[Month],MStart)-StBankClose,2)=0,"PASS","FAIL")'),
        ("Every bank line used exactly once",
         '=IF(AND(SUMPRODUCT((Recon[Bank ID]<>"")/COUNTIF(Recon[Bank ID],Recon[Bank ID]&""))=COUNTIF(Recon[Bank ID],"?*"),'
         f'COUNTIF(Recon[Bank ID],"?*")={RI("Bank rows read")}-{RI("Repeats removed")}),"PASS","FAIL")'),
        ("Every cash book line used exactly once",
         '=IF(AND(SUMPRODUCT((Recon[Ledger ID]<>"")/COUNTIF(Recon[Ledger ID],Recon[Ledger ID]&""))=COUNTIF(Recon[Ledger ID],"?*"),'
         f'COUNTIF(Recon[Ledger ID],"?*")={RI("Cash book rows")}),"PASS","FAIL")'),
        ("Picked account and month balance to 0.00", '=IF(StUnexplained=0,"PASS","FAIL")'),
        ('=PeriodsChecked&" account-months (every account, every month) balance to 0.00"', '=IF(AND(PeriodsChecked>0,PeriodsFailed=0),"PASS","FAIL")'),
        ("Every exception is on the workbench with a next step",
         '=IF(AND(COUNTIFS(Recon[Status],"<>Matched")=ROWS(Workbench[Exception ID]),COUNTIFS(Workbench[What to do],"")=0),"PASS","FAIL")'),
    ]
    for i, (label, f) in enumerate(checks, 53):
        text(ws, f"B{i}", label, "#C9C3B7", 9.5, indent=1)
        r = text(ws, f"K{i}:L{i}", f, "good", 9, True, None, "c", True)
        fc = r.FormatConditions.Add(1, 3, '="PASS"'); fc.Font.Color = c("good"); fc.Interior.Color = c("#16352A")
        fc = r.FormatConditions.Add(1, 4, '="PASS"'); fc.Font.Color = c("bad"); fc.Interior.Color = c("#3A1E1A")
    name(wb, "ChecksPassed", "=Calc!$B$27")
    calc.Range("A27").Value = "ChecksPassed"
    calc.Range("B27").Formula = '=COUNTIF(Summary!$K$53:$K$60,"PASS")'
    text(ws, "B61", '="Checks passed: "&ChecksPassed&" of 8"', "gold2", 10, True, indent=1)
    ws.Range("B61:L61").Borders(8).LineStyle = 1; ws.Range("B61:L61").Borders(8).Color = c("line")

    # proof steps / audit trail (N52:X62) ---------------------------------------------------------------------------
    lux.panel(ws, "N52:X62", "Proof steps / audit trail")
    CL = 'INDEX(CleaningLog[Rows],MATCH("References typed differently (case, spaces, dashes)",CleaningLog[Check],0))'
    steps = [
        f'="Read "&{RI("Files read")}&" CSV files: "&TEXT({RI("Bank rows read")},"#,##0")&" bank rows, "&TEXT({RI("Cash book rows")},"#,##0")&" cash book rows ("&{RI("Accounts")}&")"',
        f'="Cleaned: "&{RI("Repeats removed")}&" repeated export rows removed, "&{CL}&" references normalised, US dates and text amounts converted"',
        '="Matched "&TEXT(COUNTIFS(Recon[Match rule],"Reference"),"#,##0")&" on reference, "&COUNTIFS(Recon[Match rule],"Reference ± amount")'
        '&" within ±"&TEXT(AmountTolerance,"0.00")&", "&COUNTIFS(Recon[Match rule],"Amount + date")&" on amount + date; "&COUNTIFS(Recon[Status],"<>Matched")&" exceptions"',
        '=SelAccount&" "&MonthLabel&": adjusted bank "&TEXT(StAdjBank,"#,##0.00")&" = adjusted cash book "&TEXT(StAdjBook,"#,##0.00")&", difference "&TEXT(StUnexplained,"0.00")',
        f'="Checks "&ChecksPassed&"/8 PASS  ·  "&OpenExc&" open, "&ResolvedExc&" resolved on the workbench  ·  refreshed "&{RI("Last refresh")}',
    ]
    for k, f in enumerate(steps):
        r = 53 + 2 * k
        num = ws.Range(f"N{r}:N{r + 1}")
        num.Merge()
        text(ws, f"N{r}", str(k + 1), "gold", 14, False, lux.SERIF, "c")
        num.HorizontalAlignment = -4108
        body = ws.Range(f"O{r}:X{r + 1}")
        body.Merge()
        body.Cells(1, 1).Formula = f
        font(body, "ivory", 9.5)
        body.WrapText = True
        body.VerticalAlignment = -4108
        body.HorizontalAlignment = -4131
    for k in range(1, 5):
        rng = ws.Range(f"N{53 + 2 * k}:X{53 + 2 * k}")
        rng.Borders(8).LineStyle = 1; rng.Borders(8).Color = c("line")

    ws.Rows(63).RowHeight = 10
    ps = ws.PageSetup
    ps.Orientation = 2
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$Y$63"
    xl.Calculate()


def donut_chart(ws, calc, left, top, w, h):
    co = ws.ChartObjects().Add(left, top, w, h)
    co.Name = "chDonut"
    ch = co.Chart
    ch.ChartType = -4120  # xlDoughnut
    s = ch.SeriesCollection().NewSeries()
    s.Name = "Lines"
    s.Values = calc.Range("E3:E5")
    s.XValues = calc.Range("D3:D5")
    lux.chart_dark(ch, "Lines this month: matched vs exceptions")
    ch.ChartGroups(1).DoughnutHoleSize = 64
    for i, col in enumerate(["jade", "sky", "copper"], 1):
        s.Points(i).Format.Fill.ForeColor.RGB = c(col)
    s.Format.Line.ForeColor.RGB = c("card")
    s.Format.Line.Weight = 1.5
    lux.labels(s, "ink", 9, "0;;;")
    s.DataLabels().Font.Bold = True
    ch.HasLegend = True
    ch.Legend.Position = -4107  # bottom
    ch.Legend.Font.Color = c("muted")
    ch.Legend.Font.Size = 9
    ch.PlotArea.InsideTop  # force layout
    pa = ch.PlotArea
    cx = left + pa.InsideLeft + pa.InsideWidth / 2
    cy = top + pa.InsideTop + pa.InsideHeight / 2
    for txt, dy, size, col, fname, link in (("=Calc!$B$25", -16, 22, "gold2", lux.SERIF, True),
                                             ("lines matched", 12, 8.5, "muted", lux.SANS, False)):
        tb = ws.Shapes.AddTextbox(1, cx - 60, cy + dy - 12, 120, 24)
        tb.Fill.Visible = False
        tb.Line.Visible = False
        tf = tb.TextFrame2
        tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
        tf.VerticalAnchor = 3
        if link:
            tb.DrawingObject.Formula = txt
        else:
            tf.TextRange.Text = txt
        tr = tf.TextRange
        tr.Font.Size = size
        tr.Font.Name = fname
        tr.Font.Fill.ForeColor.RGB = c(col)
        tr.ParagraphFormat.Alignment = 2
        tb.Placement = 3
        tb.Name = "donutCenter" if link else "donutCaption"


def status_chart(ws, calc, left, top, w, h):
    co = ws.ChartObjects().Add(left, top, w, h)
    co.Name = "chStatus"
    ch = co.Chart
    ch.ChartType = 57  # xlBarClustered
    s = ch.SeriesCollection().NewSeries()
    s.Name = "Amount"
    s.Values = calc.Range("E8:E12")
    s.XValues = calc.Range("D8:D12")
    lux.chart_dark(ch, "Exceptions this month: amount by status")
    ch.HasLegend = False
    s.Format.Fill.ForeColor.RGB = c("copper")
    ch.ChartGroups(1).GapWidth = 70
    lux.labels(s, "ivory", 9, "#,##0.00;-#,##0.00;;")
    cat = ch.Axes(1)
    cat.ReversePlotOrder = True
    lux.axis_dark(cat)
    cat.TickLabels.Font.Color = c("#C9C3B7")
    ch.Axes(2).HasMajorGridlines = False
    ch.Axes(2).Delete()
    # leave room on the right so the labels of the longest bars stay outside the bars
    pa = ch.PlotArea
    pa.Width = pa.Width * 0.80


def trend_chart(ws, calc, left, top, w, h):
    co = ws.ChartObjects().Add(left, top, w, h)
    co.Name = "chTrend"
    ch = co.Chart
    ch.ChartType = 51  # xlColumnClustered
    s1 = ch.SeriesCollection().NewSeries()
    s1.Name = "Open difference"
    s1.Values = calc.Range("F15:F20")
    s1.XValues = calc.Range("D15:D20")
    s2 = ch.SeriesCollection().NewSeries()
    s2.Name = "Auto-match %"
    s2.Values = calc.Range("E15:E20")
    s2.XValues = calc.Range("D15:D20")
    s2.ChartType = 65  # xlLineMarkers
    s2.AxisGroup = 2
    lux.chart_dark(ch, "6-month trend: auto-match % and open difference")
    s1.Format.Fill.ForeColor.RGB = c("copper")
    ch.ChartGroups(1).GapWidth = 80
    s2.Format.Line.ForeColor.RGB = c("gold")
    s2.Format.Line.Weight = 2.25
    s2.MarkerStyle = 8  # circle
    s2.MarkerSize = 6
    s2.MarkerBackgroundColor = c("gold")
    s2.MarkerForegroundColor = c("gold")
    lux.labels(s2, "gold2", 8.5, "0%", 0)  # above
    lux.labels(s1, "#E0B79A", 8.5, "#,##0;-#,##0;;", 2)  # outside end
    ax2 = ch.Axes(2, 2)
    ax2.MinimumScale = 0.5   # keeps the % line in the middle, clear of the bar labels at the top
    ax2.MaximumScale = 1.25
    ax2.MajorUnit = 0.25
    lux.axis_dark(ax2, False, "0%")
    ax1 = ch.Axes(2, 1)
    ax1.MinimumScale = 0
    lux.axis_dark(ax1, True, "#,##0")
    lux.axis_dark(ch.Axes(1))
    ch.HasLegend = True
    ch.Legend.Position = -4107
    ch.Legend.Font.Color = c("muted")
    ch.Legend.Font.Size = 9
