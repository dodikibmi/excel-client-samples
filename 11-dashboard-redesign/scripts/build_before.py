"""Builds Before-Client-Dashboard.xlsx: the kind of dashboard a client sends with "can you make this look professional?".
It looks dated (3D pie, 3D columns, rainbow fills, long decimals) and, more importantly, 6 of its formulas are wrong
and the data under it has 3 export problems. Every mistake is planted on purpose and listed in data/planted.json and
on the Issues sheet of the redesign. Real Excel (COM), so every number on it is Excel's own result.

    python scripts/build_before.py        (run under tools/with_lock.py)
"""
import csv, json
from datetime import date
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Before-Client-Dashboard.xlsx"
CATS = ["Seeds & bulbs", "Plants", "Hand tools", "Power tools", "Watering", "Soil & feed"]
PRODUCTS = [r for r in csv.reader(open(ROOT / "data" / "products.csv", encoding="utf-8"))][1:]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct"]
rgb = lambda r, g, b: r + g * 256 + b * 65536


import json as _json
TEXT_IDS = set(_json.loads((ROOT / "data" / "planted.json").read_text())["text_revenue_orders"])


def load(name):
    rows = list(csv.reader(open(ROOT / "data" / name, encoding="utf-8")))[1:]
    out = []
    for r in rows:
        d = date.fromisoformat(r[1])
        rev = r[7] if (r[0] in TEXT_IDS and "2026" in name) else float(r[7])   # the 14 planted cells stay text
        out.append([r[0], (d - date(1899, 12, 30)).days, r[2], r[3], r[4], int(r[5]), float(r[6]), rev, float(r[8])])
    return out


r25, r26 = load("sales-2025.csv"), load("sales-2026.csv")
# the preliminary October figure someone typed in on 16 Oct (revenue of 1-15 October, rounded)
prelim = round(sum(r[7] if isinstance(r[7], float) else float(r[7].replace(",", "")) for r in r26
                   if date(2026, 10, 1) <= date(1899, 12, 30) + __import__("datetime").timedelta(days=r[1]) <= date(2026, 10, 15)), -2)

if OUT.exists():
    OUT.unlink()
xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Add()
    while wb.Worksheets.Count < 3:
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    dash, d26, d25 = wb.Worksheets(1), wb.Worksheets(2), wb.Worksheets(3)
    dash.Name, d26.Name, d25.Name = "Dashboard", "Data 2026", "Data 2025"
    hdr = ["Order ID", "Date", "Region", "Category", "Product", "Units", "Unit price", "Revenue", "Cost", "Margin"]
    for ws, rows in ((d26, r26), (d25, r25)):
        ws.Range("A1:J1").Value = hdr
        ws.Range(f"H2:H{len(rows) + 1}").NumberFormat = "General"
        ws.Range(ws.Cells(2, 1), ws.Cells(len(rows) + 1, 9)).Value = rows
        ws.Range(f"B2:B{len(rows) + 1}").NumberFormat = "m/d/yyyy"
        ws.Range(f"J2:J{len(rows) + 1}").Formula = "=IF(ISNUMBER(H2),(H2-I2)/H2,\"\")"
        ws.Range("A1:J1").Interior.Color = rgb(255, 255, 0); ws.Range("A1:J1").Font.Bold = True
        ws.Columns("A:J").AutoFit()
    # text cells must stay text (Excel would convert "1,234.50" when it is written as a value)
    for i, r in enumerate(r26, 2):
        if isinstance(r[7], str):
            d26.Cells(i, 8).NumberFormat = "@"; d26.Cells(i, 8).Value = r[7]

    D = "'Data 2026'!"
    ws = dash
    ws.Cells.Font.Name = "Calibri"; ws.Cells.Font.Size = 11
    ws.Range("A1:N1").Merge()
    t = ws.Range("A1"); t.Value = "SALES DASHBOARD 2026 !!!"; t.Font.Size = 24; t.Font.Bold = True
    t.Font.Color = rgb(255, 0, 0); t.Font.Underline = 2; t.Interior.Color = rgb(255, 255, 0); t.HorizontalAlignment = -4108
    kpis = [("TOTAL REVENUE 2026", f"=SUM({D}H2:H700)"),                                   # 1: range not extended
            ("Revenue 2025 (Jan-Oct)", "=SUMIFS('Data 2025'!H:H,'Data 2025'!B:B,\"<=\"&DATE(2025,10,31))"),
            ("Growth %", "=(B3-B4)/B3"),                                               # 7: divided by this year
            ("Number of orders", f"=COUNTA({D}A2:A700)"),
            ("Avg margin %", f"=AVERAGE({D}J2:J700)")]                                   # 5: average of row margins
    for i, (lab, f) in enumerate(kpis, 3):
        ws.Cells(i, 1).Value = lab; ws.Cells(i, 2).Formula = f
        ws.Cells(i, 1).Interior.Color = rgb(0, 176, 240); ws.Cells(i, 2).Interior.Color = rgb(146, 208, 80)
    ws.Range("A3:A7").Font.Bold = True
    # regions (4: West sums the Cost column)
    ws.Range("D3").Value = "REGION"; ws.Range("E3").Value = "Revenue"
    for i, reg in enumerate(["North", "South", "East", "West"], 4):
        col = "I" if reg == "West" else "H"
        ws.Cells(i, 4).Value = reg
        ws.Cells(i, 5).Formula = f'=SUMIF({D}C2:C700,"{reg}",{D}{col}2:{col}700)'
    ws.Range("D3:E3").Interior.Color = rgb(255, 192, 0)
    ws.Range("D4:E7").Interior.Color = rgb(255, 230, 153)
    # months (6: October typed in by hand from a preliminary report)
    ws.Range("A10").Value = "MONTH"; ws.Range("B10").Value = "Revenue 2026"
    for m, nm in enumerate(MONTHS, 1):
        ws.Cells(10 + m, 1).Value = nm
        if m < 10:
            ws.Cells(10 + m, 2).Formula = (f'=SUMIFS({D}H2:H700,{D}B2:B700,">="&DATE(2026,{m},1),'
                                           f'{D}B2:B700,"<"&DATE(2026,{m + 1},1))')
        else:
            ws.Cells(10 + m, 2).Value = prelim
    ws.Range("A10:B10").Interior.Color = rgb(112, 48, 160); ws.Range("A10:B10").Font.Color = rgb(255, 255, 255)
    # categories (8: the total adds a subtotal row too)
    ws.Range("D10").Value = "CATEGORY"; ws.Range("E10").Value = "Revenue"
    for i, c in enumerate(CATS, 11):
        ws.Cells(i, 4).Value = c
        ws.Cells(i, 5).Formula = f'=SUMIF({D}D2:D700,"{c}",{D}H2:H700)'
    ws.Range("D17").Value = "Tools subtotal"; ws.Range("E17").Formula = "=E13+E14"
    ws.Range("D18").Value = "TOTAL"; ws.Range("E18").Formula = "=SUM(E11:E17)"
    ws.Range("D10:E10").Interior.Color = rgb(255, 0, 0); ws.Range("D10:E10").Font.Color = rgb(255, 255, 0)
    ws.Range("D17:E18").Font.Bold = True
    # products (9: "Hose 25m " with a trailing space is not found by SUMIF)
    ws.Range("G3").Value = "PRODUCT"; ws.Range("H3").Value = "Revenue"
    for i, p in enumerate(PRODUCTS, 4):
        ws.Cells(i, 7).Value = p[1]
        ws.Cells(i, 8).Formula = f'=SUMIF({D}E2:E700,G{i},{D}H2:H700)'
    ws.Range("G3:H3").Interior.Color = rgb(0, 176, 80)
    for r in range(4, 4 + len(PRODUCTS)):
        ws.Range(f"G{r}:H{r}").Interior.Color = rgb(*[(255, 204, 255), (204, 255, 255), (255, 255, 153)][r % 3])
    for a in ("A3:B7", "D3:E7", "A10:B20", "D10:E18", f"G3:H{3 + len(PRODUCTS)}"):
        for e in (7, 8, 9, 10, 11, 12):
            try:
                b = ws.Range(a).Borders(e); b.LineStyle = 1; b.Weight = -4138
            except Exception:
                pass
    ws.Columns("A:H").ColumnWidth = 18
    # charts: 3D pie of the category table (subtotal included), 3D columns of the months
    co = ws.ChartObjects().Add(ws.Range("J3").Left, ws.Range("J3").Top, 420, 260)
    co.Chart.ChartType = -4102
    co.Chart.SetSourceData(ws.Range("D10:E17"))
    co.Chart.HasTitle = True; co.Chart.ChartTitle.Text = "CATEGORY SALES"
    co.Chart.SeriesCollection(1).HasDataLabels = True
    co.Chart.SeriesCollection(1).DataLabels().ShowPercentage = True; co.Chart.SeriesCollection(1).DataLabels().ShowValue = True
    co2 = ws.ChartObjects().Add(ws.Range("J22").Left, ws.Range("J22").Top, 420, 260)
    co2.Chart.ChartType = 54
    co2.Chart.SetSourceData(ws.Range("A10:B20"))
    co2.Chart.HasTitle = True; co2.Chart.ChartTitle.Text = "Monthly Sales!!"
    co2.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgb(255, 0, 0)
    ws.Activate(); ws.Range("A1").Select()
    wb.SaveAs(str(OUT), FileFormat=51)
    print("saved", OUT.name, "| total shown:", ws.Range("B3").Value, "| October typed:", prelim)
    wb.Close(False)
finally:
    xl.Quit()
