"""Independent check of Quote-Invoice-Builder.xlsm.

Part A (the saved sample, opened read-only): every document is recomputed in plain Python from data/history.json
(lines, rounding half-up to the cent, discount, tax, exempt customers), then compared with the Register, the Lines,
the printed Document page of every document, the PDF files (text read back with pypdf), the Log, the Home cards and
chart table, and the Checks sheet. Numbering, statuses, dates, links and Excel 2016 compatibility are checked too.

Part B (a throw-away copy): the macros are driven like a user would: new quote, hard input errors, every reason that
blocks issuing, fix and issue (PDF + .eml whose attachment is the same PDF), frozen documents, illegal status changes,
convert to invoice (twice = refused, prices kept after a price-list change), paid, tampering caught by the Checks
sheet, the UserForm itself (cascading lists, validation, live totals and checks), and Reset.

    python test/verify.py          (pywin32 + pypdf; run under tools/with_lock.py)
"""
import base64, csv, email, json, re, shutil, sys, tempfile
from datetime import date, datetime, timedelta
from decimal import Decimal as D, ROUND_HALF_UP
from pathlib import Path
import win32com.client as win32
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from vbom import vba_access
from watchdog import watch

BOOK = ROOT / "Quote-Invoice-Builder.xlsm"
SET = json.loads((ROOT / "data" / "settings.json").read_text())
HIST = json.loads((ROOT / "data" / "history.json").read_text())
NUM = json.loads((ROOT / "data" / "numbers.json").read_text())
CUST = {r["Customer ID"]: r for r in csv.DictReader(open(ROOT / "data" / "customers.csv", encoding="utf-8"))}
PROD = {r["Code"]: r for r in csv.DictReader(open(ROOT / "data" / "products.csv", encoding="utf-8"))}
AS_OF = date.fromisoformat(SET["as_of"])
TAX = D(str(SET["tax_rate"]))
results = []


def check(name, ok, detail=""):
    results.append((name, bool(ok)))
    print(("PASS " if ok else "FAIL ") + name + ("" if ok or not detail else f"   <- {detail}"))


def r2(x):
    return D(x).quantize(D("0.01"), ROUND_HALF_UP)


def ser(v):
    return None if v in (None, "") else date(1899, 12, 30) + timedelta(days=int(v))


def money(x):
    return f"{x:,.2f}"


# ===================================================================================== expected, from the history
docs = {d["ref"]: d for d in HIST["documents"]}
exp = {}          # number -> dict


def totals(cust, disc, lines):
    nets = [r2(D(str(q)) * D(str(p)) * (1 - D(str(dd)))) for _, q, p, dd in lines]
    sub = sum(nets, D(0))
    dsc = r2(sub * D(str(disc)))
    rate = D(0) if CUST[cust]["Tax exempt"] == "Yes" else TAX
    tax = r2((sub - dsc) * rate)
    return nets, sub, dsc, rate, tax, sub - dsc + tax


for e in HIST["events"]:
    ref, a, d = e["ref"], e["action"], date.fromisoformat(e["date"])
    doc = docs[ref]
    if a == "create":
        no = NUM["quotes"][ref]
        nets, sub, dsc, rate, tax, tot = totals(doc["customer"], doc["doc_discount"], doc["lines"])
        exp[no] = dict(type="Quote", date=d, cust=doc["customer"], lines=doc["lines"], nets=nets, sub=sub, disc=D(str(doc["doc_discount"])),
                       dsc=dsc, rate=rate, tax=tax, tot=tot, status="Draft", due=d + timedelta(days=SET["quote_valid_days"]),
                       sent=None, done=None, link="", order=len(exp))
    elif a == "issue":
        exp[NUM["quotes"][ref]].update(status="Sent", sent=d)
    elif a in ("accept", "decline"):
        exp[NUM["quotes"][ref]].update(status="Accepted" if a == "accept" else "Declined", done=d)
    elif a == "convert":
        q = exp[NUM["quotes"][ref]]; no = NUM["invoices"][ref]
        exp[no] = dict(q, type="Invoice", date=d, status="Draft", due=d + timedelta(days=int(CUST[q["cust"]]["Terms (days)"])),
                       sent=None, done=None, link=NUM["quotes"][ref], order=len(exp))
        q["link"] = no
    elif a == "issue_invoice":
        exp[NUM["invoices"][ref]].update(status="Sent", sent=d)
    elif a == "paid":
        exp[NUM["invoices"][ref]].update(status="Paid", done=d)
for no, x in exp.items():
    x["now"] = ("Expired" if x["type"] == "Quote" else "Overdue") if x["status"] == "Sent" and x["due"] < AS_OF else x["status"]

with vba_access():
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False
    watch(xl)
    tmp = Path(tempfile.mkdtemp(prefix="e10_"))
    try:
        # ============================================================================= Part A: the saved sample
        print("== Part A: the sample as saved")
        copyA = tmp / "A" / BOOK.name; copyA.parent.mkdir()
        shutil.copy(BOOK, copyA)
        wb = xl.Workbooks.Open(str(copyA))
        W = lambda n: wb.Worksheets(n)
        last = lambda ws: ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        reg = W("Register").Range(f"A6:Z{last(W('Register'))}").Value2
        lin = W("Lines").Range(f"A6:N{last(W('Lines'))}").Value2
        log = W("Log").Range(f"A6:E{last(W('Log'))}").Value2
        R = {r[0]: r for r in reg}

        check("register holds every document of the history, nothing else", set(R) == set(exp), set(R) ^ set(exp))
        check("register rows are in creation order", [r[0] for r in reg] == sorted(exp, key=lambda n: exp[n]["order"]))
        for pre in ("Q-2026-", "INV-2026-"):
            nums = sorted(int(n[len(pre):]) for n in R if n.startswith(pre))
            check(f"{pre}#### numbers run 1..{len(nums)} without a gap or a double", nums == list(range(1, len(nums) + 1)))
        bad = []
        for no, x in exp.items():
            r = R[no]
            got = (r[1], ser(r[2]), r[3], r[6], r2(r[7]), D(str(r[8])), r2(r[9]), D(str(r[10])), r2(r[11]), r2(r[12]), r[13], ser(r[14]),
                   ser(r[15]), ser(r[16]), r[17] or "")
            want = (x["type"], x["date"], x["cust"], len(x["lines"]), x["sub"], x["disc"], x["dsc"], x["rate"], x["tax"], x["tot"],
                    x["status"], x["due"], x["sent"], x["done"], x["link"])
            if got != want:
                bad.append((no, [(i, g, w) for i, (g, w) in enumerate(zip(got, want)) if g != w]))
        check("every register row = Python recompute (type, date, customer, lines, subtotal, discount, tax rate, tax, total, "
              "status, valid/due, sent, accepted/paid, link)", not bad, bad[:3])
        check("customer name and email copied from Customers", all(R[n][4] == CUST[x["cust"]]["Name"] and (R[n][5] or "") == CUST[x["cust"]]["Email"]
                                                                  for n, x in exp.items()))
        check("Status now (Expired / Overdue against the as-of date)", all(R[n][21] == x["now"] for n, x in exp.items()),
              [(n, R[n][21], x["now"]) for n, x in exp.items() if R[n][21] != x["now"]][:3])
        check("days late = as-of - due date for overdue invoices", all(
            (R[n][22] == (AS_OF - x["due"]).days) if x["now"] == "Overdue" else R[n][22] in ("", None) for n, x in exp.items()))
        check("tax exempt customer pays 0% tax", all(R[n][10] == 0 for n, x in exp.items() if x["cust"] == "C011") and
              any(x["cust"] == "C011" for x in exp.values()))
        bad = []
        L = {}
        for r in lin:
            L.setdefault(r[0], []).append(r)
        for no, x in exp.items():
            rows = sorted(L.get(no, []), key=lambda r: r[1])
            want = [(i + 1, c, PROD[c]["Description"], PROD[c]["Unit"], D(str(q)), D(str(p)), D(str(dd)), n)
                    for i, ((c, q, p, dd), n) in enumerate(zip(x["lines"], x["nets"]))]
            got = [(int(r[1]), r[2], r[3], r[4], D(str(r[5])), D(str(r[6])), D(str(r[7])), r2(r[8])) for r in rows]
            if got != want:
                bad.append((no, got[:1], want[:1]))
        check("every line (code, description, unit, qty, price, discount, amount) = Python recompute", not bad, bad[:2])
        check("invoices carry the quote's own lines and prices", all(
            [r[2:9] for r in sorted(L[n], key=lambda r: r[1])] == [r[2:9] for r in sorted(L[x["link"]], key=lambda r: r[1])]
            for n, x in exp.items() if x["type"] == "Invoice"))
        check("Lines helper columns: no amount 'Wrong', no line 'Missing', no issued line 'Under min'",
              not any(r[10] == "Wrong" or r[12] == "Missing" or r[13] == "Under min" for r in lin))

        # drafts blocked on purpose
        d1, d2 = NUM["quotes"]["d1"], NUM["quotes"]["d2"]
        check("draft for a customer without email: issue check says so", "has no email address" in (R[d1][19] or ""), R[d1][19])
        check("draft priced under the minimum: issue check names the line and the minimum",
              "Line 1 SGN-04: price 120.00 is under the minimum 135.00" in (R[d2][19] or ""), R[d2][19])
        check("issued documents keep an empty issue check", all(not R[n][19] for n, x in exp.items() if x["status"] != "Draft"))

        # log
        acts = {}
        for r in log:
            acts.setdefault(r[3], []).append(r[2])
        check("Log: every document Created once", all(acts.get(n, []).count("Created") == 1 for n in exp))
        check("Log: every issued document has one Issued entry", all(acts.get(n, []).count("Issued") == (1 if x["sent"] else 0) for n, x in exp.items()))
        check("Log: Accepted / Declined / Paid entries match the statuses", all(
            acts.get(n, []).count(x["status"]) == (1 if x["status"] in ("Accepted", "Declined", "Paid") else 0) for n, x in exp.items()))
        check("Log: every conversion logged on the quote", all(acts.get(x["link"], []).count("Converted") == 1 for x in exp.values() if x["type"] == "Invoice"))
        check("Log: the two planted drafts show 'Issue blocked' with the reason",
              acts.get(d1, []).count("Issue blocked") == 1 and acts.get(d2, []).count("Issue blocked") == 1 and
              any("no email" in r[4] for r in log if r[3] == d1) and any("under the minimum" in r[4] for r in log if r[3] == d2))
        check("Log is in time order and every entry has a user", [r[0] for r in log] == sorted(r[0] for r in log) and
              all(r[1] == SET["user"] for r in log))

        # PDFs
        pdf_dir = ROOT / SET["pdf_folder"]
        issued = [n for n, x in exp.items() if x["sent"]]
        check("every issued document has its PDF on disk, and no other PDF exists",
              sorted(p.name for p in pdf_dir.glob("*.pdf")) == sorted(R[n][18] for n in issued), )
        bad = []
        for n in issued:
            x = exp[n]
            t = " ".join(PdfReader(pdf_dir / R[n][18]).pages[0].extract_text().split())
            need = [n, CUST[x["cust"]]["Name"].replace("&", "&"), money(x["tot"]), "INVOICE" if x["type"] == "Invoice" else "QUOTE",
                    money(x["sub"])] + [PROD[c]["Description"] for c, *_ in x["lines"]]
            miss = [s for s in need if s not in t]
            if miss or "DRAFT" in t or len(PdfReader(pdf_dir / R[n][18]).pages) != 1:
                bad.append((n, miss))
        check(f"PDF text read back ({len(issued)} files): number, customer, every line, subtotal, total, one page, no DRAFT mark",
              not bad, bad[:3])

        # every document on the printed page
        doc = W("Document")
        bad = []
        for n, x in exp.items():
            doc.Range("M2").Value = n
            xl.Calculate()
            got = (doc.Range("H2").Value, r2(doc.Range("I29").Value), r2(-doc.Range("I30").Value), r2(doc.Range("I31").Value),
                   r2(doc.Range("I32").Value), sum(1 for i in range(13, 28) if doc.Range(f"C{i}").Value not in ("", None)),
                   doc.Range("M33").Value, doc.Range("B11").Value or "")
            want = (x["type"].upper(), x["sub"], x["dsc"], x["tax"], x["tot"], len(x["lines"]), "matches register",
                    "DRAFT  ·  not issued yet" if x["status"] == "Draft" else "")
            if got != want:
                bad.append((n, got, want))
        check(f"Document page for all {len(exp)} documents: title, subtotal, discount, tax, total, line count, draft mark",
              not bad, bad[:2])

        # checks sheet and Home
        check("Checks sheet: every check PASS", wb.Names("ChecksPassed").RefersToRange.Value == wb.Names("ChecksTotal").RefersToRange.Value == 11,
              (wb.Names("ChecksPassed").RefersToRange.Value, wb.Names("ChecksTotal").RefersToRange.Value))
        h = W("Home")
        q = [x for x in exp.values() if x["type"] == "Quote"]; inv = [x for x in exp.values() if x["type"] == "Invoice"]
        won = sum(x["status"] == "Accepted" for x in q); lost = sum(x["status"] == "Declined" for x in q)
        expd = sum(x["now"] == "Expired" for x in q)
        first = AS_OF.replace(day=1)
        cards = {
            "B14": float(sum(x["tot"] for x in q if x["now"] == "Sent")),
            "E14": won / (won + lost + expd),
            "H14": float(sum(x["tot"] for x in inv if x["status"] != "Draft" and first <= x["date"] <= AS_OF)),
            "K14": float(sum(x["tot"] for x in inv if x["status"] == "Sent")),
            "N14": float(sum(x["tot"] for x in inv if x["now"] == "Overdue")),
        }
        check("Home cards: open quotes, win rate, invoiced this month, unpaid, overdue = Python",
              all(abs(h.Range(c).Value - v) < 0.005 for c, v in cards.items()), {c: (h.Range(c).Value, v) for c, v in cards.items()})
        att = sum(1 for x in exp.values() if x["now"] == "Overdue" or x["status"] == "Draft" or
                  (x["type"] == "Quote" and x["now"] == "Sent" and (x["due"] - AS_OF).days <= 7))
        check("Home 'needs attention' count and list = overdue + expiring quotes + drafts",
              h.Range("Q14").Value == att and sum(1 for r in range(21, 33) if h.Range(f"J{r}").Value) == min(att, 12),
              (h.Range("Q14").Value, att))
        check("Home card notes: won / declined / expired counts",
              h.Range("E16").Value == f"{won} won · {lost} declined · {expd} expired", h.Range("E16").Value)
        calc = W("Calc").Range("A4:D7").Value2
        okm = True
        for row in calc:
            m0 = ser(row[0]); m1 = (m0.replace(day=28) + timedelta(days=4)).replace(day=1)
            wq = sum(x["tot"] for x in q if x["status"] != "Draft" and m0 <= x["date"] < m1)
            wi = sum(x["tot"] for x in inv if x["status"] != "Draft" and m0 <= x["date"] < m1)
            wp = sum(x["tot"] for x in inv if x["status"] == "Paid" and m0 <= x["done"] < m1)
            okm &= abs(row[1] - float(wq)) < .005 and abs(row[2] - float(wi)) < .005 and abs(row[3] - float(wp)) < .005
        check("chart table: quoted / invoiced / collected per month = Python", okm)
        lastn = sorted(exp, key=lambda n: exp[n]["order"])[::-1][:10]
        check("Home: last 10 documents, newest first", [h.Range(f"B{r}").Value for r in range(38, 48)] == lastn)

        # structure
        prot = {n: W(n).ProtectContents for n in ("Home", "Document", "Register", "Lines", "Log", "Checks")}
        check("Register, Lines, Log, Home, Document, Checks are protected", all(prot.values()), prot)
        check("only the picker cells are unlocked on Home / Document",
              not h.Range("B9").Locked and not doc.Range("M2").Locked and h.Range("B14").Locked and doc.Range("I32").Locked)
        comps = {c.Name: c for c in wb.VBProject.VBComponents}
        check("VBA: modCore, modButtons and the frmDocument UserForm are in the file", {"modCore", "modButtons", "frmDocument"} <= set(comps))
        ctl = {c.Name for c in comps["frmDocument"].Designer.Controls}
        check("UserForm has customer, category -> product, qty, price, discount, lines list, totals, checks and the 3 buttons",
              {"cboCustomer", "cboCategory", "cboProduct", "txtQty", "txtPrice", "txtDisc", "lstLines", "lblTotal", "lblChecks",
               "cmdSave", "cmdIssue", "cmdClose"} <= ctl)
        errs, newer = [], []
        for ws in wb.Worksheets:
            ur = ws.UsedRange
            for row in (ur.Value2 or ()):
                for v in (row if isinstance(row, tuple) else (row,)):
                    if isinstance(v, int) and v in (-2146826281, -2146826246, -2146826259, -2146826288, -2146826252, -2146826265, -2146826273):
                        errs.append(ws.Name)
            try:
                fr = ws.UsedRange.Formula
            except Exception:
                fr = ()
            for row in fr:
                for f in (row if isinstance(row, tuple) else (row,)):
                    if isinstance(f, str) and re.search(r"\b(XLOOKUP|FILTER|LET|UNIQUE|SORTBY|SEQUENCE|TEXTJOIN|IFS|MAXIFS|MINIFS|CONCAT|SWITCH)\(", f):
                        newer.append((ws.Name, f[:60]))
        check("no error values on any sheet", not errs, errs[:3])
        check("Excel 2016 functions only (no XLOOKUP, FILTER, LET, IFS, MAXIFS ...)", not newer, newer[:3])
        wb.Close(False)

        # ============================================================================= Part B: the macros on a copy
        print("== Part B: driving the macros on a throw-away copy")
        copyB = tmp / "B" / BOOK.name; copyB.parent.mkdir()
        shutil.copy(BOOK, copyB)
        wb = xl.Workbooks.Open(str(copyB))
        W = lambda n: wb.Worksheets(n)
        run = lambda *a: xl.Run(*a)
        for nm, v in (("PdfFolder", str(tmp / "B" / "PDF")), ("OutboxFolder", str(tmp / "B" / "Outbox")), ("EmailMode", ".eml file")):
            wb.Names(nm).RefersToRange.Value = v
        nrows = lambda: last(W("Register")) - 5
        today = 46342  # 16 Nov 2026 as an Excel serial (a datetime would be shifted by the time zone)
        n0 = nrows()
        nextq = f"Q-2026-{len(NUM['quotes']) + 1:04d}"
        spec = "SGN-04~2~165~0|PRT-02~1~120~0.1|SRV-04~3.5~55~0"
        no = run("NewDocument", "Quote", "C005", today, 0.05, "Test quote", spec)
        lines = [("SGN-04", 2, 165, 0), ("PRT-02", 1, 120, 0.1), ("SRV-04", 3.5, 55, 0)]
        _, sub, dsc, _, tax, tot = totals("C005", 0.05, lines)
        rr = W("Register").Range(f"A{5 + nrows()}:N{5 + nrows()}").Value2[0]
        check("new quote gets the next number", no == nextq, no)
        check("new quote: subtotal, discount, tax, total = Python, status Draft",
              (r2(rr[7]), r2(rr[9]), r2(rr[11]), r2(rr[12]), rr[13]) == (sub, dsc, tax, tot, "Draft"), (rr, sub, dsc, tax, tot))
        for label, args in (("unknown product", ("C005", "XXX-01~1~10~0")), ("quantity 0", ("C005", "PRT-01~0~45~0")),
                            ("no lines", ("C005", "")), ("unknown customer", ("C999", "PRT-01~1~45~0")),
                            ("discount 100%", ("C005", "PRT-01~1~45~1")), ("16 lines", ("C005", "|".join(["PRT-01~1~45~0"] * 16)))):
            res = run("NewDocument", "Quote", args[0], today, 0, "", args[1])
            check(f"hard input error refused, nothing saved: {label}", res.startswith("ERR:") and nrows() == n0 + 1, res)
        blocked = [("price under the minimum", "C005", 0, "SGN-04~2~120~0", "under the minimum 135.00"),
                   ("line discount above the product max", "C005", 0, "LFP-01~10~28~0.25", "above the 20% allowed"),
                   ("document discount above the limit", "C005", 0.2, "LFP-01~10~28~0", "above the 15% limit"),
                   ("customer without email", "C009", 0, "LFP-01~10~28~0", "has no email address")]
        pdfs = lambda: sorted((tmp / "B" / "PDF").glob("*.pdf")) if (tmp / "B" / "PDF").exists() else []
        for label, c, dd, sp, msg in blocked:
            n = run("NewDocument", "Quote", c, today, dd, "", sp)
            before = len(pdfs())
            res = run("IssueDocument", n, today)
            st = W("Register").Cells(run("DocRow", n), 14).Value
            lg = W("Log").Range(f"C{last(W('Log'))}:E{last(W('Log'))}").Value[0]
            check(f"issue blocked: {label} (reason shown, still Draft, no PDF, logged)",
                  res.startswith("BLOCKED:") and msg in res and st == "Draft" and len(pdfs()) == before and lg[0] == "Issue blocked" and msg in lg[2],
                  (res, st, lg))
        # duplicate number planted by hand -> blocked + Checks FAIL
        ws = W("Register"); ws.Unprotect()
        rlast = last(ws)
        ws.Rows(rlast).Copy(ws.Rows(rlast + 1))
        dup = ws.Cells(rlast, 1).Value
        res = run("CheckDocument", dup); xl.CalculateFull()
        chk = W("Checks").Range("D9").Value
        check("a number typed twice is caught: issue check + Checks sheet FAIL", "used twice" in res and chk == "FAIL", (res, chk))
        ws.Rows(rlast + 1).Delete(); xl.CalculateFull()
        # fix the under-minimum draft and issue: PDF + eml
        n = nextq
        res = run("UpdateDraft", n, "C005", today, 0.05, "Test quote", spec)
        res = run("IssueDocument", n, today)
        r = run("DocRow", n)
        pdf = tmp / "B" / "PDF" / W("Register").Cells(r, 19).Value
        eml = tmp / "B" / "Outbox" / f"{n}.eml"
        check("issue: status Sent, sent date, PDF written", res == "" and W("Register").Cells(r, 14).Value == "Sent" and pdf.exists(), res)
        msg = email.message_from_bytes(eml.read_bytes()) if eml.exists() else None
        att = [p for p in msg.walk() if p.get_content_type() == "application/pdf"] if msg else []
        check(".eml: To = customer email, unsent draft, subject has the number, body has the total",
              msg and msg["To"] == CUST["C005"]["Email"] and msg["X-Unsent"] == "1" and n in msg["Subject"] and
              money(tot) in msg.get_payload()[0].get_payload())
        check(".eml attachment is byte-for-byte the PDF", len(att) == 1 and att[0].get_payload(decode=True) == pdf.read_bytes())
        res = run("UpdateDraft", n, "C005", today, 0, "", spec)
        check("an issued document is frozen (edit refused)", res.startswith("ERR:") and "frozen" in res, res)
        res = run("SetStatus", n, "Paid", today)
        check("a quote cannot be marked paid", res.startswith("ERR:"), res)
        res = run("ConvertToInvoice", n, today)
        check("a quote that is only Sent cannot become an invoice", res.startswith("ERR:"), res)
        res = run("SetStatus", n, "Accepted", 46386)
        check("an expired quote cannot be accepted", res.startswith("ERR:") and "expired" in res, res)
        res = run("SetStatus", n, "Accepted", today)
        check("accept a sent quote", res == "" and W("Register").Cells(r, 14).Value == "Accepted", res)
        # price list changes after the quote: the invoice keeps the agreed price
        pr = W("Products").Range("A6:A40").Value
        prow = 6 + [x[0] for x in pr].index("SGN-04")
        W("Products").Cells(prow, 5).Value = 199
        inv = run("ConvertToInvoice", n, today)
        ri = run("DocRow", inv)
        check("convert: next invoice number, linked both ways",
              inv == f"INV-2026-{len(NUM['invoices']) + 1:04d}" and W("Register").Cells(ri, 18).Value == n and W("Register").Cells(r, 18).Value == inv, inv)
        check("invoice due date = date + customer terms; totals = the quote",
              ser(W("Register").Cells(ri, 15).Value2) == date(2026, 11, 16) + timedelta(days=int(CUST["C005"]["Terms (days)"])) and
              r2(W("Register").Cells(ri, 13).Value) == tot)
        check("invoice keeps the quoted price after a price-list change", run("SpecOf", inv) == run("SpecOf", n) and "SGN-04~2~165~0" in run("SpecOf", inv))
        res = run("ConvertToInvoice", n, today)
        check("a quote cannot be invoiced twice", res.startswith("ERR:") and "already invoiced" in res, res)
        res = run("IssueDocument", inv, today)
        check("issue the invoice (PDF + email)", res == "" and (tmp / "B" / "Outbox" / f"{inv}.eml").exists(), res)
        res = run("SetStatus", inv, "Paid", 46327)
        check("payment dated before the invoice is refused", res.startswith("ERR:"), res)
        res = run("SetStatus", inv, "Paid", 46346)
        check("mark the invoice paid", res == "" and W("Register").Cells(ri, 14).Value == "Paid", res)
        xl.CalculateFull()
        check("Checks sheet still all PASS after the new documents", wb.Names("ChecksPassed").RefersToRange.Value == wb.Names("ChecksTotal").RefersToRange.Value)
        # tampering
        wsl = W("Lines"); wsl.Unprotect()
        lr = last(wsl); old = wsl.Cells(lr, 9).Value
        wsl.Cells(lr, 9).Value = old + 1; xl.CalculateFull()
        check("an amount typed over on Lines is caught (2 checks FAIL)", W("Checks").Range("D8").Value == "FAIL" and W("Checks").Range("D6").Value == "FAIL")
        wsl.Cells(lr, 9).Value = old; xl.CalculateFull()
        ws.Unprotect(); ot = ws.Cells(r, 13).Value; ws.Cells(r, 13).Value = ot + 10; xl.CalculateFull()
        check("a register total typed over is caught", W("Checks").Range("D7").Value == "FAIL")
        ws.Cells(r, 13).Value = ot; xl.CalculateFull()
        check("restored: all PASS again", wb.Names("ChecksPassed").RefersToRange.Value == wb.Names("ChecksTotal").RefersToRange.Value)

        # the UserForm itself (driven through a small test module added to this copy only)
        tm = wb.VBProject.VBComponents.Add(1)
        tm.CodeModule.AddFromString('''
Public Sub T_New(): frmDocument.StartNew "Quote": frmDocument.Show vbModeless: End Sub
Public Sub T_Set(c As String, v As String): frmDocument.Controls(c).Value = v: End Sub
Public Sub T_Pick(c As String, i As Long): frmDocument.Controls(c).ListIndex = i: End Sub
Public Function T_Get(c As String) As String
    On Error Resume Next
    T_Get = frmDocument.Controls(c).Caption
    If Err.Number <> 0 Then Err.Clear: T_Get = frmDocument.Controls(c).Value
End Function
Public Function T_Count(c As String) As Long: T_Count = frmDocument.Controls(c).ListCount: End Function
Public Function T_Item(c As String, i As Long) As String: T_Item = frmDocument.Controls(c).List(i): End Function
Public Sub T_Add(): frmDocument.AddLine: End Sub
Public Sub T_Issue(): frmDocument.IssueNow: End Sub
Public Sub T_Close(): Unload frmDocument: End Sub
''')
        run("T_New")
        names = [run("T_Item", "cboCustomer", i) for i in range(run("T_Count", "cboCustomer"))]
        run("T_Pick", "cboCustomer", [i for i, s in enumerate(names) if s.startswith("C009")][0])
        check("form: picking a customer without email shows it in red text", "NO EMAIL ADDRESS" in run("T_Get", "lblCust"))
        cats = [run("T_Item", "cboCategory", i) for i in range(run("T_Count", "cboCategory"))]
        check("form: categories listed once each", cats == ["Printing", "Signage", "Large format", "Design & services"], cats)
        run("T_Pick", "cboCategory", cats.index("Signage"))
        prods = [run("T_Item", "cboProduct", i) for i in range(run("T_Count", "cboProduct"))]
        check("form: product list follows the category (cascading)", len(prods) == 6 and all(p.startswith("SGN-") for p in prods), prods)
        run("T_Pick", "cboProduct", [i for i, p in enumerate(prods) if p.startswith("SGN-04")][0])
        check("form: picking a product fills the current list price (changed to 199 above) and shows min / max discount",
              run("T_Get", "txtPrice") == "199.00" and "minimum 135.00" in run("T_Get", "lblProd"),
              (run("T_Get", "txtPrice"), run("T_Get", "lblProd")))
        run("T_Set", "txtQty", "abc"); run("T_Add")
        check("form: a quantity that is not a number is refused", "Quantity must be a number" in run("T_Get", "lblMsg") and run("T_Count", "lstLines") == 0)
        run("T_Set", "txtQty", "2"); run("T_Set", "txtPrice", "120"); run("T_Add")
        check("form: line added, totals live (2 x 120 + 10% tax = $264.00)",
              run("T_Count", "lstLines") == 1 and run("T_Get", "lblTotal") == "$264.00", run("T_Get", "lblTotal"))
        chk = run("T_Get", "lblChecks")
        check("form: live checks list both problems before saving", "no email address" in chk and "under the minimum 135.00" in chk, chk)
        n_before = nrows()
        run("T_Issue")
        newno = W("Register").Cells(5 + nrows(), 1).Value
        check("form: Issue saves a draft but refuses to issue it, and says so",
              nrows() == n_before + 1 and W("Register").Cells(5 + nrows(), 14).Value == "Draft" and "NOT issued" in run("T_Get", "lblMsg"),
              run("T_Get", "lblMsg"))
        run("T_Close")
        # reset
        run("ResetDocuments"); xl.CalculateFull()
        check("Reset: register, lines empty; log keeps one 'Reset' entry; numbering restarts at 0001",
              nrows() == 0 and last(W("Lines")) == 5 and last(W("Log")) == 6 and W("Log").Range("C6").Value == "Reset" and
              run("NextNumber", "Quote", today) == "Q-2026-0001")
        check("Reset: customers, products and settings kept", last(W("Customers")) - 5 == len(CUST) and last(W("Products")) - 5 == len(PROD))
        wb.Close(False)
    finally:
        xl.Quit()
        shutil.rmtree(tmp, ignore_errors=True)

passed = sum(ok for _, ok in results)
print(f"\n{passed}/{len(results)} checks pass")
(ROOT / "test" / "last-result.txt").write_text(f"{passed}/{len(results)}\n")
sys.exit(0 if passed == len(results) else 1)
