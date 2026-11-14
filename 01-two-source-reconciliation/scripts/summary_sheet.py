"""Summary sheet of Two-Source-Reconciliation.xlsx: KPI cards, result by status, chart,
bank reconciliation statement and data checks. Called by build_workbook.py (Excel COM)."""


def rgb(r, g, b):
    return r + g * 256 + b * 65536


NAVY, TEAL = (31, 58, 90), (0, 128, 128)
STATUS_COLORS = {"Matched": (226, 239, 218), "Amount differs": (252, 228, 214), "Duplicate in ledger": (252, 228, 214),
                 "Duplicate in bank": (252, 228, 214), "Only in bank": (255, 242, 204), "Only in ledger": (221, 235, 247)}
GREEN = ((198, 239, 206), (0, 97, 0))
RED = ((255, 199, 206), (156, 0, 6))


def heading(ws, cell, text):
    ws.Range(cell).Value = text
    ws.Range(cell).Font.Size = 13
    ws.Range(cell).Font.Bold = True
    ws.Range(cell).Font.Color = rgb(*NAVY)


def pass_fail(rng, ok_value="=0", ok_is_equal=True):
    """Green when the cell passes, red otherwise (cell-value rules)."""
    fc = rng.FormatConditions.Add(1, 3 if ok_is_equal else 4, ok_value)
    fc.Interior.Color = rgb(*GREEN[0]); fc.Font.Color = rgb(*GREEN[1])
    fc = rng.FormatConditions.Add(1, 4 if ok_is_equal else 3, ok_value)
    fc.Interior.Color = rgb(*RED[0]); fc.Font.Color = rgb(*RED[1])


def build_summary(wb):
    ws = wb.Worksheets("Summary")
    widths = {"A": 2, "B": 46, "C": 2, "D": 16, "E": 2, "F": 16, "G": 2, "H": 16, "I": 3,
              "J": 12, "K": 12, "L": 12, "M": 12, "N": 12}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w

    # Title band ------------------------------------------------------------------------------------
    ws.Range("A1:N2").Interior.Color = rgb(*NAVY)
    ws.Range("B1").Formula = '="Bank reconciliation   |   Lakeside Hardware Ltd   |   "&Period'
    ws.Range("B1").Font.Size = 18; ws.Range("B1").Font.Bold = True; ws.Range("B1").Font.Color = rgb(255, 255, 255)
    ws.Range("B2").Value = ("Bank statement vs cash book, matched in Power Query. Fictional data. "
                            "New month: put the two CSV exports in the data folder, then Data > Refresh All.")
    ws.Range("B2").Font.Color = rgb(200, 215, 230)
    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 20

    # Result by status (rows 12-20) -----------------------------------------------------------------
    heading(ws, "B12", "Result by status")
    for c, t in {"B": "Status", "D": "Lines", "F": "Bank amount", "H": "Cash book amount"}.items():
        h = ws.Range(f"{c}13")
        h.Value = t; h.Font.Bold = True; h.Interior.Color = rgb(*NAVY); h.Font.Color = rgb(255, 255, 255)
        if c != "B":
            h.HorizontalAlignment = -4152
    statuses = list(STATUS_COLORS)
    for i, st in enumerate(statuses, 14):
        ws.Range(f"B{i}").Value = st
        ws.Range(f"B{i}").Interior.Color = rgb(*STATUS_COLORS[st])
        ws.Range(f"D{i}").Formula = f"=COUNTIFS(Recon[Status],B{i})"
        ws.Range(f"F{i}").Formula = f"=SUMIFS(Recon[Bank amount],Recon[Status],B{i})"
        ws.Range(f"H{i}").Formula = f"=SUMIFS(Recon[Ledger amount],Recon[Status],B{i})"
    ws.Range("B20").Value = "Total"
    for c in "DFH":
        ws.Range(f"{c}20").Formula = f"=SUM({c}14:{c}19)"
    for c in "BDFH":
        ws.Range(f"{c}20").Font.Bold = True
        ws.Range(f"{c}20").Borders(8).LineStyle = 1
    ws.Range("D14:D20").NumberFormat = "0"
    ws.Range("F14:H20").NumberFormat = "#,##0.00;[Red]-#,##0.00"

    # Bank reconciliation statement (rows 23-39) ----------------------------------------------------
    heading(ws, "B23", "Bank reconciliation statement")
    lines = [
        ("Balance per bank statement, opening", "=BankOpening"),
        ("Movements on the bank statement", "=SUM(Recon[Bank amount])"),
        ("Balance per bank statement, closing", "=D24+D25"),
        ("Add: deposits in transit", '=SUMIFS(Recon[Ledger amount],Recon[Status],"Only in ledger",Recon[Ledger amount],">0")'),
        ("Less: outstanding payments", '=SUMIFS(Recon[Ledger amount],Recon[Status],"Only in ledger",Recon[Ledger amount],"<0")'),
        ("Adjusted bank balance", "=D26+D27+D28"),
        (None, None),
        ("Balance per cash book, opening", "=LedgerOpening"),
        ("Movements in the cash book", "=SUM(Recon[Ledger amount])"),
        ("Balance per cash book, closing", "=D31+D32"),
        ("Add/less: bank items not yet in the cash book",
         '=SUMIFS(Recon[Bank amount],Recon[Status],"Only in bank")+SUMIFS(Recon[Bank amount],Recon[Status],"Duplicate in bank")'),
        ("Add/less: amount differences to correct", '=SUMIFS(Recon[Difference],Recon[Status],"Amount differs")'),
        ("Less: postings made twice", '=-SUMIFS(Recon[Ledger amount],Recon[Status],"Duplicate in ledger")'),
        ("Adjusted cash book balance", "=D33+D34+D35+D36"),
        (None, None),
        ("Unexplained difference (must be 0.00)", "=ROUND(D29-D37,2)"),
    ]
    for i, (label, f) in enumerate(lines, 24):
        if label:
            ws.Range(f"B{i}").Value = label
            ws.Range(f"D{i}").Formula = f
    ws.Range("D24:D39").NumberFormat = "#,##0.00;[Red]-#,##0.00"
    for r in (26, 29, 33, 37, 39):
        ws.Range(f"B{r}:D{r}").Font.Bold = True
    for r in (29, 37):
        ws.Range(f"B{r}:D{r}").Borders(8).LineStyle = 1
    ws.Range("B39:C39").Interior.Color = rgb(242, 242, 242)
    pass_fail(ws.Range("D39"))

    # Data checks (rows 23-27, right) ---------------------------------------------------------------
    heading(ws, "F23", "Data checks")
    checks = [
        ("Statement closing balance = opening + movements", '=IF(ROUND(INDEX(BankControls[Value],1)-(BankOpening+INDEX(BankControls[Value],2)),2)=0,"PASS","FAIL")'),
        ("Every bank line used exactly once",
         '=IF(SUMPRODUCT((Recon[Bank line]<>"")/COUNTIF(Recon[Bank line],Recon[Bank line]&""))=COUNT(Recon[Bank line]),"PASS","FAIL")'),
        ("Every cash book line used exactly once",
         '=IF(SUMPRODUCT((Recon[Ledger line]<>"")/COUNTIF(Recon[Ledger line],Recon[Ledger line]&""))=COUNT(Recon[Ledger line]),"PASS","FAIL")'),
        ("Reconciliation balances to 0.00", '=IF(D39=0,"PASS","FAIL")'),
    ]
    for i, (label, f) in enumerate(checks, 24):
        ws.Range(f"F{i}").Value = label
        cell = ws.Range(f"N{i}")
        cell.Formula = f
        cell.Font.Bold = True
        cell.HorizontalAlignment = -4108
        pass_fail(cell, '="PASS"')

    # KPI cards (rows 4-9) --------------------------------------------------------------------------
    cards = [("B", "B", "MATCHED AUTOMATICALLY", "=D14/COUNT(Recon[Bank line])", "0.0%",
              '=D14&" of "&COUNT(Recon[Bank line])&" bank lines, no manual work"'),
             ("D", "F", "NEED A PERSON", "=D20-D14", "0", '="lines, each with a next step"'),
             ("H", "L", "TO POST IN THE CASH BOOK", "=D34+D35+D36", "#,##0.00;-#,##0.00",
              '="fees, interest, differences, duplicates"'),
             ("M", "N", "UNEXPLAINED", "=D39", "#,##0.00", '=IF(D39=0,"fully reconciled","check")')]
    for a, b, label, formula, fmt, note in cards:
        box = ws.Range(f"{a}4:{b}9")
        box.Interior.Color = rgb(247, 249, 252)
        for edge in (7, 8, 9, 10):
            box.Borders(edge).LineStyle = 1
            box.Borders(edge).Color = rgb(210, 218, 228)
        ws.Range(f"{a}4:{b}4").Interior.Color = rgb(*TEAL)
        for rows, val, size, bold, color in (("5:5", label, 9, True, (90, 100, 115)), ("6:7", formula, 26, True, NAVY),
                                             ("8:8", note, 9, False, (110, 110, 110))):
            r1, r2 = rows.split(":")
            rng = ws.Range(f"{a}{r1}:{b}{r2}")
            if rng.Cells.Count > 1:
                rng.Merge()
            cell = ws.Range(f"{a}{r1}")
            if val.startswith("="):
                cell.Formula = val
            else:
                cell.Value = val
            cell.Font.Size = size; cell.Font.Bold = bold; cell.Font.Color = rgb(*color)
            rng.HorizontalAlignment = -4108
            rng.VerticalAlignment = -4108
        ws.Range(f"{a}6").NumberFormat = fmt
    ws.Rows(4).RowHeight = 5
    fc = ws.Range("M6").FormatConditions.Add(1, 3, "=0"); fc.Font.Color = rgb(0, 128, 0)
    fc = ws.Range("M6").FormatConditions.Add(1, 4, "=0"); fc.Font.Color = rgb(192, 0, 0)

    # Chart: lines that need a person, by type ------------------------------------------------------
    top, left = ws.Range("J12").Top, ws.Range("J12").Left
    co = ws.ChartObjects().Add(left, top, ws.Range("J12:N12").Width, ws.Range("J12:J20").Height)
    ch = co.Chart
    ch.ChartType = 57  # xlBarClustered
    ser = ch.SeriesCollection().NewSeries()
    ser.Name = "Lines"
    ser.Values = ws.Range("D15:D19")
    ser.XValues = ws.Range("B15:B19")
    ch.HasLegend = False
    ch.HasTitle = False
    ser.Format.Fill.ForeColor.RGB = rgb(*TEAL)
    ser.HasDataLabels = True
    ch.Axes(1).ReversePlotOrder = True
    ch.Axes(2).HasMajorGridlines = False
    ch.Axes(2).Delete()
    # Excel re-creates a one-series chart's automatic title on open, so the title is set on purpose instead
    ch.HasTitle = True
    ch.ChartTitle.Text = "Lines to review, by type"
    ch.ChartTitle.Font.Size = 13
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Font.Color = rgb(*NAVY)
    ch.ChartTitle.Left = 4
    ch.ChartArea.Format.Line.Visible = False

    ps = ws.PageSetup
    ps.Orientation = 2
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$N$39"
