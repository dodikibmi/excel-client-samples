"""Generates the fictional supplier invoices for "Cedar Lane Property Management" (made-up, September-October 2026).

300 PDF files land in ../invoices/ the way an accounts payable inbox collects them:
  - 8 suppliers, 4 different invoice layouts (labels, date formats, number formats and column names all differ)
  - long invoices that run over 2-3 pages, with the table header repeated on each page
  - credit notes written as negative amounts in brackets, "(1,234.50)"
  - the same invoice sent twice under another file name (duplicates)
  - a supplier whose lines do not add up to the subtotal it printed (arithmetic errors)
  - scanned invoices (a picture of the page, no text to read), a password-protected file and two broken files
The designed answer goes to ../test/expected.json, so the workbook can be checked against a known result.
"""
import datetime as dt
import io
import json
import random
import shutil
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageFilter
from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib.pdfencrypt import StandardEncryption

random.seed(20261116)
ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "invoices"
TEST = ROOT / "test"
C2 = Decimal("0.01")


def money(x):
    return Decimal(x).quantize(C2, rounding=ROUND_HALF_UP)


PROPERTIES = ["Maple Court Apartments", "Harbor View Lofts", "Elm Street Plaza", "Willow Park Residences",
              "Oakridge Business Center", "Birch Hill Townhomes"]
PROP_CODE = {"Maple Court Apartments": "MCA", "Harbor View Lofts": "HVL", "Elm Street Plaza": "ESP",
             "Willow Park Residences": "WPR", "Oakridge Business Center": "OBC", "Birch Hill Townhomes": "BHT"}

# supplier, layout, invoice number pattern, tax rate, terms (days), weight (how many invoices), file name style
SUPPLIERS = [
    dict(name="Brightway Janitorial Co.", layout="A", prefix="BJ-", start=10412, tax="0.0000", terms=30, n=46,
         file="Brightway {no}.pdf",
         items=[("Common area cleaning - monthly", 1, 1450, 2600), ("Window cleaning, exterior", 1, 380, 720),
                ("Carpet shampoo, hallways", 1, 260, 540), ("Trash room sanitising", 1, 120, 240),
                ("Move-out unit clean", 1, 180, 260), ("Floor strip and wax, lobby", 1, 340, 610)]),
    dict(name="Summit Landscaping Group", layout="A", prefix="SLG-", start=7731, tax="0.0000", terms=15, n=38,
         file="SLG invoice {no}.pdf",
         items=[("Grounds maintenance - monthly", 1, 900, 1800), ("Leaf removal", 1, 280, 650),
                ("Irrigation repair", 1, 140, 480), ("Seasonal planting", 1, 420, 980),
                ("Tree trimming", 1, 380, 1200), ("Mulch, delivered and spread", 1, 260, 540)]),
    dict(name="Northgate Electrical Services", layout="B", prefix="NES/26/", start=301, tax="0.0725", terms=30,
         n=36, file="invoice_{no}.pdf",
         hours=[("Electrician - standard hours", 85), ("Electrician - after hours", 127.5),
                ("Apprentice", 48)],
         parts=[("LED panel 2x4", 64.9), ("GFCI outlet", 18.75), ("Breaker 20A", 22.4), ("Exit sign, LED", 49.0),
                ("Photocell sensor", 31.5), ("Wire 12/2, 250 ft", 118.0)]),
    dict(name="Riverbend Plumbing & Heating", layout="B", prefix="RPH", start=55120, tax="0.0725", terms=14, n=40,
         file="Riverbend_RPH{no}.pdf",
         hours=[("Plumber - call-out and first hour", 145), ("Plumber - additional hours", 95),
                ("Boiler technician", 120)],
         parts=[("Shut-off valve 3/4in", 26.8), ("Water heater element", 42.5), ("Toilet flapper kit", 12.9),
                ("PEX pipe 1/2in, 100 ft", 54.0), ("Pump seal kit", 88.0), ("Drain snake rental", 35.0)]),
    dict(name="Clearwater Maintenance Supply", layout="C", prefix="CMS-", start=880100, tax="0.0725", terms=30,
         n=44, file="CMS_{no}.pdf",
         catalog=[("Air filter MERV 8 20x20x1", "AF-2020-8", 6.45), ("Air filter MERV 11 16x25x1", "AF-1625-11", 9.8),
                  ("LED bulb A19 9W (4 pk)", "LB-A19-4", 11.2), ("Paint, eggshell white, 1 gal", "PT-EW-1", 32.5),
                  ("Caulk, silicone, clear", "CK-SIL", 6.95), ("Toilet seat, elongated", "TS-ELG", 24.0),
                  ("Faucet cartridge", "FC-1222", 18.4), ("Door closer, grade 1", "DC-G1", 96.0),
                  ("Smoke detector, 10-year", "SD-10Y", 21.9), ("Trash bags 55 gal (100)", "TB-55", 38.0),
                  ("Ice melt, 50 lb", "IM-50", 17.5), ("Mop head, looped", "MH-L", 7.4),
                  ("Light switch, decora", "LS-DEC", 3.85), ("Outlet cover plate", "OCP-1", 1.15),
                  ("Weatherstrip, door kit", "WS-DK", 14.6), ("HVAC coil cleaner", "HC-CC", 12.75),
                  ("Drain cleaner, enzyme", "DC-ENZ", 15.2), ("Batteries AA (24)", "BAT-AA", 16.9),
                  ("Padlock, keyed alike", "PL-KA", 13.3), ("Fire extinguisher 5 lb", "FE-5", 54.0)]),
    dict(name="Keystone Office & Janitorial", layout="C", prefix="KOJ", start=40210, tax="0.0725", terms=30, n=30,
         file="Keystone invoice KOJ{no}.pdf",
         catalog=[("Copy paper, case", "CP-10", 48.9), ("Paper towels, case", "PTW-12", 41.5),
                  ("Toilet tissue, case", "TT-96", 64.0), ("Hand soap refill", "HS-R", 9.9),
                  ("Disinfectant spray", "DS-1", 6.3), ("Gloves nitrile (100)", "GN-100", 11.8),
                  ("Glass cleaner, gal", "GC-G", 12.4), ("Floor cleaner, gal", "FC-G", 18.6),
                  ("Toner, black", "TN-BK", 89.0), ("Envelopes #10 (500)", "EN-10", 27.5),
                  ("Sticky notes (12)", "SN-12", 9.6), ("Ballpoint pens (12)", "BP-12", 5.2)]),
    dict(name="Atlas Elevator Maintenance", layout="D", prefix="AEM-", start=2291, tax="0.0000", terms=45, n=34,
         file="Atlas Elevator {mon} {no}.pdf",
         items=[("Monthly maintenance contract", 1, 610, 980), ("Annual safety test", 1, 1250, 1850),
                ("Emergency call-out", 1, 395, 720), ("Door operator repair", 1, 860, 2400),
                ("Cab lighting replacement", 1, 210, 460)]),
    dict(name="Pinecrest Security Systems", layout="D", prefix="PSS", start=19044, tax="0.0725", terms=30, n=32,
         file="PSS{no}.pdf",
         items=[("Monitoring - monthly", 1, 140, 260), ("Access control service", 1, 220, 640),
                ("Camera replacement", 1, 380, 950), ("Key fobs (10)", 1, 85, 85), ("Gate motor repair", 1, 540, 1300)]),
]
# 46+38+36+40+44+30+34+32 = 300 files

START = dt.date(2026, 9, 1)
END = dt.date(2026, 10, 31)


def rand_date():
    return START + dt.timedelta(days=random.randint(0, (END - START).days))


def fmt_amount(x, style):
    """How each supplier prints numbers."""
    x = money(x)
    if style == "A":  # 1,234.50
        return f"{x:,.2f}"
    if style == "B":  # 1234.50 (no thousands separator)
        return f"{x:.2f}"
    if style == "C":  # $1,234.50
        return f"${x:,.2f}"
    if style == "D":  # $1,234.50 and credits in brackets ($1,234.50)
        return f"(${-x:,.2f})" if x < 0 else f"${x:,.2f}"
    raise ValueError(style)


def fmt_date(d, layout):
    if layout == "A":
        return d.strftime("%d/%m/%Y")   # day first (the supplier's own system)
    if layout == "B":
        return f"{d.day} {d.strftime('%b %Y')}"  # 7 Oct 2026
    if layout == "C":
        return d.strftime("%Y-%m-%d")
    return f"{d.month}/{d.day}/{d.year}"  # US, no leading zeros


def make_invoice(sup, seq):
    d = rand_date()
    prop = random.choice(PROPERTIES)
    inv = dict(supplier=sup["name"], layout=sup["layout"], number=f"{sup['prefix']}{sup['start'] + seq}",
               date=d.isoformat(), due=(d + dt.timedelta(days=sup["terms"])).isoformat(), property=prop,
               po=f"PO-{PROP_CODE[prop]}-{random.randint(1100, 1999)}" if random.random() < 0.75 else "",
               kind="Invoice", lines=[])
    L = sup["layout"]
    if L in ("A", "D"):
        for desc, qty, lo, hi in random.sample(sup["items"], random.randint(1, 4)):
            q = qty if random.random() < 0.8 else random.randint(2, 4)
            price = money(random.uniform(lo, hi))
            inv["lines"].append(dict(desc=desc, qty=q, unit=str(price), amount=str(money(price * q))))
    elif L == "B":
        for desc, rate in random.sample(sup["hours"], random.randint(1, 2)):
            h = Decimal(random.choice([1, 1.5, 2, 2.5, 3, 4, 5, 6, 8]))
            inv["lines"].append(dict(desc=desc, qty=float(h), unit=str(money(rate)), amount=str(money(h * money(rate)))))
        for desc, p in random.sample(sup["parts"], random.randint(0, 3)):
            q = random.randint(1, 6)
            inv["lines"].append(dict(desc=desc, qty=q, unit=str(money(p)), amount=str(money(money(p) * q))))
    else:  # C: supply orders, some long enough for 2-3 pages
        n = random.choice([random.randint(3, 12)] * 3 + [random.randint(28, 45), random.randint(46, 70)])
        for _ in range(n):
            desc, sku, p = random.choice(sup["catalog"])
            q = random.randint(1, 24)
            inv["lines"].append(dict(desc=desc, sku=sku, qty=q, unit=str(money(p)), amount=str(money(money(p) * q))))
    sub = sum(Decimal(l["amount"]) for l in inv["lines"])
    tax = money(sub * Decimal(sup["tax"]))
    inv.update(subtotal=str(money(sub)), tax=str(tax), total=str(money(sub + tax)), tax_rate=sup["tax"])
    return inv


def to_credit(inv):
    inv["kind"] = "Credit note"
    for l in inv["lines"]:
        l["amount"] = str(-Decimal(l["amount"]))
        l["unit"] = str(-Decimal(l["unit"]))
    for k in ("subtotal", "tax", "total"):
        inv[k] = str(-Decimal(inv[k]))
    inv["lines"][0]["desc"] = "Credit: " + inv["lines"][0]["desc"]


# ---------------------------------------------------------------- PDF layouts
BASE = ParagraphStyle("base", fontName="Helvetica", fontSize=9, leading=11)
HEAD = ParagraphStyle("head", fontName="Helvetica-Bold", fontSize=15, leading=20)
SMALL = ParagraphStyle("small", parent=BASE, fontSize=8, leading=10, textColor=colors.HexColor("#555555"))


def P(t, st=BASE):
    return Paragraph(t, st)


SUP_ADDR = {
    "Brightway Janitorial Co.": "418 Commerce Way, Suite 2, Riverton, OR 97301",
    "Summit Landscaping Group": "77 Orchard Road, Riverton, OR 97302",
    "Northgate Electrical Services": "1290 Industrial Pkwy, Unit 5, Riverton, OR 97304",
    "Riverbend Plumbing & Heating": "36 Mill Street, Riverton, OR 97301",
    "Clearwater Maintenance Supply": "5500 Distribution Dr, Salem, OR 97305",
    "Keystone Office & Janitorial": "210 Market Ave, Salem, OR 97302",
    "Atlas Elevator Maintenance": "901 Lift Lane, Portland, OR 97210",
    "Pinecrest Security Systems": "64 Watchtower Ct, Portland, OR 97211",
}
BILL_TO = "Cedar Lane Property Management<br/>Accounts Payable<br/>1500 Cedar Lane, Riverton, OR 97301"


def build_pdf(inv, path, encrypt=None, broken_math=None):
    L = inv["layout"]
    amt_style = {"A": "A", "B": "B", "C": "C", "D": "D"}[L]
    lines = inv["lines"]
    sub, tax, total = Decimal(inv["subtotal"]), Decimal(inv["tax"]), Decimal(inv["total"])
    if broken_math:  # supplier printed a subtotal that does not match its own lines
        sub = sub + broken_math
        total = sub + tax
    story = []
    title = "CREDIT NOTE" if inv["kind"] == "Credit note" else ("TAX INVOICE" if L == "B" else "INVOICE")
    d = dt.date.fromisoformat(inv["date"])
    due = dt.date.fromisoformat(inv["due"])
    fa = lambda x: fmt_amount(x, amt_style)
    if L == "A":
        story += [P(inv["supplier"], HEAD), P(SUP_ADDR[inv['supplier']], SMALL),
                  Spacer(1, 10), P(title, HEAD), Spacer(1, 6)]
        head = [["Invoice No", inv["number"]], ["Invoice Date", fmt_date(d, L)], ["Due Date", fmt_date(due, L)],
                ["PO Number", inv["po"] or "-"], ["Site", inv["property"]]]
        ht = Table(head, colWidths=[1.3 * inch, 2.6 * inch])
        ht.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.grey), ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 9),
                                ("FONT", (1, 0), (1, -1), "Helvetica", 9)]))
        story += [Table([[ht, P("<b>Bill To</b><br/>" + BILL_TO)]], colWidths=[4.1 * inch, 2.8 * inch],
                        style=[("VALIGN", (0, 0), (-1, -1), "TOP")]), Spacer(1, 14)]
        rows = [["Description", "Qty", "Unit Price", "Amount"]]
        rows += [[l["desc"], str(l["qty"]), fa(l["unit"]), fa(l["amount"])] for l in lines]
        widths = [3.9 * inch, 0.6 * inch, 1.1 * inch, 1.2 * inch]
        tot = [["Subtotal", fa(sub)], [f"Sales Tax ({Decimal(inv['tax_rate']) * 100:.2f}%)", fa(tax)], ["Total Due", fa(total)]]
    elif L == "B":
        story += [Table([[P(f"<b><font size=16 color='#1f4e79'>{inv['supplier']}</font></b><br/>" + SUP_ADDR[inv['supplier']], SMALL),
                          P(f"<para align=right><b><font size=14>{title}</font></b></para>")]],
                        colWidths=[4.6 * inch, 2.3 * inch]), Spacer(1, 12)]
        head = [["Invoice #", inv["number"], "Customer", "Cedar Lane Property Mgmt"],
                ["Issued", fmt_date(d, L), "Job site", inv["property"]],
                ["Payment due", fmt_date(due, L), "Your ref", inv["po"] or "n/a"]]
        ht = Table(head, colWidths=[1.0 * inch, 1.6 * inch, 1.0 * inch, 3.3 * inch])
        ht.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "Helvetica", 9), ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 9),
                                ("FONT", (2, 0), (2, -1), "Helvetica-Bold", 9),
                                ("LINEBELOW", (0, -1), (-1, -1), 0.8, colors.HexColor("#1f4e79"))]))
        story += [ht, Spacer(1, 14)]
        rows = [["Item", "Hrs/Qty", "Rate", "Line Total"]]
        rows += [[l["desc"], f"{l['qty']:g}" if isinstance(l["qty"], float) else str(l["qty"]), fa(l["unit"]), fa(l["amount"])]
                 for l in lines]
        widths = [3.8 * inch, 0.8 * inch, 1.0 * inch, 1.3 * inch]
        tot = [["Net", fa(sub)], [f"Tax {Decimal(inv['tax_rate']) * 100:.2f}%", fa(tax)], ["Amount due (USD)", fa(total)]]
    elif L == "C":
        story += [P(inv["supplier"].upper(), HEAD), P(SUP_ADDR[inv['supplier']], SMALL),
                  Spacer(1, 8), P(title, HEAD), Spacer(1, 4)]
        head = [["Inv. Number", "Inv. Date", "Terms", "Ship To", "Customer PO"],
                [inv["number"], fmt_date(d, L), f"Net {(due - d).days}", inv["property"], inv["po"] or ""]]
        ht = Table(head, colWidths=[1.1 * inch, 1.0 * inch, 0.8 * inch, 2.4 * inch, 1.6 * inch])
        ht.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dddddd")),
                                ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8), ("FONT", (0, 1), (-1, 1), "Helvetica", 8)]))
        story += [ht, Spacer(1, 10)]
        rows = [["Line", "SKU", "Description", "Qty", "Unit $", "Ext $"]]
        rows += [[str(i), l["sku"], l["desc"], str(l["qty"]), fa(l["unit"]), fa(l["amount"])] for i, l in enumerate(lines, 1)]
        widths = [0.45 * inch, 1.0 * inch, 3.0 * inch, 0.5 * inch, 0.85 * inch, 1.0 * inch]
        tot = [["Merchandise", fa(sub)], ["Sales tax", fa(tax)], ["Invoice total", fa(total)]]
    else:  # D
        story += [P(inv["supplier"], HEAD), P(SUP_ADDR[inv['supplier']], SMALL),
                  Spacer(1, 10), P(title, HEAD), Spacer(1, 6),
                  P(f"<b>Document no.:</b> {inv['number']}"), P(f"<b>Date:</b> {fmt_date(d, L)}"),
                  P(f"<b>Pay by:</b> {fmt_date(due, L)}"), P(f"<b>Purchase order:</b> {inv['po'] or 'none'}"),
                  P(f"<b>Location:</b> {inv['property']}"), Spacer(1, 12)]
        rows = [["Service", "Qty", "Price", "Total"]]
        rows += [[l["desc"], str(l["qty"]), fa(l["unit"]), fa(l["amount"])] for l in lines]
        widths = [3.9 * inch, 0.6 * inch, 1.1 * inch, 1.2 * inch]
        tot = [["Subtotal", fa(sub)], ["Tax", fa(tax)], ["Balance due" if inv["kind"] == "Invoice" else "Credit total", fa(total)]]

    it = Table(rows, colWidths=widths, repeatRows=1)
    style = [("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 9), ("FONT", (0, 1), (-1, -1), "Helvetica", 9),
             ("ALIGN", (-3 if L != "C" else -3, 0), (-1, -1), "RIGHT")]
    if L in ("A", "C"):
        style.append(("GRID", (0, 0), (-1, -1), 0.5, colors.grey))
    elif L == "B":
        style += [("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.HexColor("#1f4e79")),
                  ("LINEBELOW", (0, -1), (-1, -1), 0.8, colors.HexColor("#1f4e79"))]
    else:
        style += [("BOX", (0, 0), (-1, -1), 0.8, colors.black), ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.black),
                  ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee"))]
    it.setStyle(TableStyle(style))
    tt = Table(tot, colWidths=[1.8 * inch, 1.2 * inch], hAlign="RIGHT")
    tt.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "Helvetica", 9), ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 10),
                            ("ALIGN", (1, 0), (1, -1), "RIGHT"), ("LINEABOVE", (0, -1), (-1, -1), 0.8, colors.black)]))
    story += [it, Spacer(1, 8), tt, Spacer(1, 16),
              P(f"Please quote {inv['number']} with your payment. Thank you for your business.", SMALL)]

    def footer(c, doc):
        c.setFont("Helvetica", 7)
        c.setFillColor(colors.HexColor("#777777"))
        c.drawString(0.75 * inch, 0.5 * inch, f"{inv['supplier']} - {inv['number']}")
        c.drawRightString(LETTER[0] - 0.75 * inch, 0.5 * inch, f"Page {doc.page}")

    doc = SimpleDocTemplate(str(path), pagesize=LETTER, leftMargin=0.75 * inch, rightMargin=0.75 * inch,
                            topMargin=0.7 * inch, bottomMargin=0.8 * inch, title=f"{inv['supplier']} {inv['number']}",
                            author=inv["supplier"], encrypt=encrypt)
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return doc.page


def scan_pdf(inv, path):
    """A scanned paper invoice: the page is a picture, there is no text in the file."""
    # draw a rough picture of the first page instead of rendering the real PDF (no renderer needed)
    W, H = 1275, 1650
    img = Image.new("L", (W, H), 246)
    dr = ImageDraw.Draw(img)
    try:
        big = ImageFont.truetype("arial.ttf", 34)
        f = ImageFont.truetype("arial.ttf", 22)
    except OSError:
        big = f = ImageFont.load_default()
    y = 110
    dr.text((110, y), inv["supplier"], font=big, fill=30); y += 70
    dr.text((110, y), "INVOICE", font=big, fill=30); y += 70
    for k, v in (("Invoice No", inv["number"]), ("Date", inv["date"]), ("Site", inv["property"]), ("PO", inv["po"] or "-")):
        dr.text((110, y), f"{k}: {v}", font=f, fill=40); y += 36
    y += 30
    for l in inv["lines"]:
        dr.text((110, y), l["desc"], font=f, fill=40)
        dr.text((900, y), f"{Decimal(l['amount']):,.2f}", font=f, fill=40); y += 34
    y += 20
    dr.text((700, y), f"Total due  {Decimal(inv['total']):,.2f}", font=big, fill=20)
    img = img.rotate(random.uniform(-1.2, 1.2), fillcolor=246).filter(ImageFilter.GaussianBlur(0.6))
    c = rl_canvas.Canvas(str(path), pagesize=LETTER)
    c.drawImage(ImageReader(img), 0, 0, LETTER[0], LETTER[1])
    c.showPage()
    c.save()


def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    TEST.mkdir(exist_ok=True)
    invoices = []
    for sup in SUPPLIERS:
        for seq in range(sup["n"]):
            inv = make_invoice(sup, seq)
            inv["sup"] = sup
            invoices.append(inv)
    # credit notes: Atlas and Pinecrest (layout D) and two from Clearwater (layout C)
    for inv in random.sample([i for i in invoices if i["layout"] == "D"], 5) + \
            random.sample([i for i in invoices if i["layout"] == "C" and len(i["lines"]) < 10], 2):
        to_credit(inv)
    random.shuffle(invoices)

    pool = [i for i in invoices if i["kind"] == "Invoice"]
    picks = random.sample(pool, 4 + 3 + 6 + 1 + 2)
    dup_src, bad_math, scans, locked, broken = picks[:4], picks[4:7], picks[7:13], picks[13], picks[14:16]
    for inv in bad_math:
        inv["broken_math"] = str(money(random.choice([10, 25, 100, 9.9, -50])))
    # duplicates replace other invoices so the folder still holds 300 files
    dup_targets = random.sample([i for i in pool if i not in picks], 4)
    files = []
    for inv in invoices:
        if inv in dup_targets:
            continue
        files.append((inv, "original"))
    for src in dup_src:
        files.append((src, "duplicate"))
    assert len(files) == 300, len(files)
    dropped = {id(i) for i in dup_targets}
    invoices = [i for i in invoices if id(i) not in dropped]

    used = set()
    manifest = []
    for inv, role in files:
        sup = inv["sup"]
        d = dt.date.fromisoformat(inv["date"])
        num = inv["number"][len(sup["prefix"]):]
        name = sup["file"].format(no=num, mon=d.strftime("%b"))
        if role == "duplicate":
            name = random.choice(["Fwd {n}", "{n0} (1).pdf", "scan_{r}.pdf"]).format(
                n=name, n0=name[:-4], r=random.randint(1000, 9999))
            if not name.endswith(".pdf"):
                name += ""
        if inv["kind"] == "Credit note":
            name = name.replace(".pdf", " CREDIT.pdf")
        while name in used:
            name = name.replace(".pdf", "_.pdf")
        used.add(name)
        path = OUT / name
        status = "Read"
        pages = 1
        if inv in scans:
            scan_pdf(inv, path)
            status = "Scanned image"
        elif inv is locked and role == "original":
            build_pdf(inv, path, encrypt=StandardEncryption("open-sesame", canPrint=1))
            status = "Password protected"
        elif inv in broken and role == "original":
            build_pdf(inv, path)
            data = path.read_bytes()
            path.write_bytes(data[: len(data) // 3])  # download cut off half way
            status = "Damaged file"
        else:
            bm = Decimal(inv["broken_math"]) if inv.get("broken_math") else None
            pages = build_pdf(inv, path, broken_math=bm)
        manifest.append(dict(file=name, role=role, status=status, pages=pages, supplier=inv["supplier"],
                             number=inv["number"]))

    expected = dict(files=manifest, invoices=[{k: v for k, v in i.items() if k != "sup"} for i in invoices])
    (TEST / "expected.json").write_text(json.dumps(expected, indent=1), encoding="utf-8")
    by = {}
    for m in manifest:
        by[m["status"]] = by.get(m["status"], 0) + 1
    print(len(manifest), "files", by, "multi-page:", sum(1 for m in manifest if m["pages"] > 1),
          "credit notes:", sum(1 for i in invoices if i["kind"] == "Credit note"))


if __name__ == "__main__":
    main()
