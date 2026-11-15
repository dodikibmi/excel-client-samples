"""Independent check of Customer-Cleanup.xlsx.

1. Re-does the whole cleanup in plain Python straight from the three source files (no Excel involved):
   clean every field, drop exact duplicates and test rows, link records that share an email, a phone or
   first name + last name + ZIP (nicknames: Bill = William), keep the most recent record as master.
2. Compares it with the answer that was planted when the data was generated (test/expected.json): every group of
   records must be exactly one real person, and the look-alike pairs must be flagged for review, not merged.
3. Opens the workbook with Excel and compares every customer, merge, review pair and data check with Python, and every
   number of the cockpit: KPI cards, data-quality gauges (before/after), funnel, rule chart, field-completeness heatmap,
   the before/after card of one customer (incl. which cells are highlighted) and the Match explorer for several picks.
4. Copies workbook + data to another folder, refreshes there, and checks the answer is the same; then marks one
   look-alike pair Merge on Needs review, refreshes, and checks the two customers became one exactly as Python predicts.
Exit code 1 if anything differs.   Run:  python test/verify.py   (needs openpyxl; pywin32 for steps 3-4)
"""
import csv, datetime as dt, json, re, sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
results = []


def check(name, ok, detail=""):
    results.append(bool(ok))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


# ---------------------------------------------------------------- the cleaning rules (same as the Power Query steps)
NICK = {"JON": "JOHN", "JOHNNY": "JOHN", "BILL": "WILLIAM", "WILL": "WILLIAM", "BOB": "ROBERT", "ROB": "ROBERT",
        "LIZ": "ELIZABETH", "BETH": "ELIZABETH", "KATE": "KATHERINE", "KATIE": "KATHERINE", "MIKE": "MICHAEL",
        "JIM": "JAMES", "JIMMY": "JAMES", "TOM": "THOMAS", "DAN": "DANIEL", "DANNY": "DANIEL", "CHRIS": "CHRISTOPHER",
        "MATT": "MATTHEW", "NICK": "NICHOLAS", "SAM": "SAMUEL", "ALEX": "ALEXANDER", "TONY": "ANTHONY",
        "JEN": "JENNIFER", "JENNY": "JENNIFER", "MEG": "MARGARET", "PEGGY": "MARGARET", "STEVE": "STEVEN",
        "DAVE": "DAVID", "ANDY": "ANDREW", "BEN": "BENJAMIN", "JOE": "JOSEPH", "ABBY": "ABIGAIL", "GABE": "GABRIEL",
        "TED": "EDWARD", "ED": "EDWARD", "RICK": "RICHARD", "SUE": "SUSAN", "BECKY": "REBECCA", "NATE": "NATHAN"}
STATES = {"COLORADO": "CO", "UTAH": "UT", "WASHINGTON": "WA", "OREGON": "OR", "CALIFORNIA": "CA", "ARIZONA": "AZ",
          "NEW MEXICO": "NM", "TEXAS": "TX", "ILLINOIS": "IL", "MINNESOTA": "MN", "MASSACHUSETTS": "MA",
          "CONNECTICUT": "CT", "NEW JERSEY": "NJ", "VERMONT": "VT", "MAINE": "ME", "NEW YORK": "NY"}
ABBR = {"STREET": "St", "ST": "St", "AVENUE": "Ave", "AVE": "Ave", "ROAD": "Rd", "RD": "Rd", "DRIVE": "Dr", "DR": "Dr",
        "LANE": "Ln", "LN": "Ln", "COURT": "Ct", "CT": "Ct", "BOULEVARD": "Blvd", "BLVD": "Blvd"}
NO_COMPANY = {"N/A", "NA", "NONE", "-", "SELF-EMPLOYED"}
PRIORITY = {"Webshop": 1, "POS": 2, "Sales sheet": 3}

T = lambda x: "" if x is None else str(x).strip()
squash = lambda s: " ".join(p for p in s.strip().split(" ") if p)
digits = lambda s: re.sub(r"\D", "", s)
proper = lambda s: s.lower().title()


def clean(rec):
    """rec: Source, Source row, Name typed, First typed, Last typed, raw fields, Last activity, Spend."""
    if rec["first_t"] or rec["last_t"]:
        first, last = squash(T(rec["first_t"])), squash(T(rec["last_t"]))
    else:
        s = squash(T(rec["name"]))
        if "," in s:
            p = s.split(",")
            last, first = p[0].strip(), ",".join(p[1:]).strip()
        else:
            p = s.split(" ")
            first, last = p[0], " ".join(p[1:])
    first, last = proper(first), proper(last)
    e = T(rec["email"]).replace(" ", "").lower() or None
    d = digits(T(rec["phone"]))
    d = d[1:] if len(d) == 11 and d.startswith("1") else d
    phone = f"({d[:3]}) {d[3:6]}-{d[6:]}" if len(d) == 10 else None
    z = digits(T(rec["zip"]).split("-")[0])
    zip_ = z if len(z) == 5 else "0" + z if len(z) == 4 else None
    st = T(rec["state"]).upper()
    st = (st if len(st) == 2 else STATES.get(st, st)) or None
    a = squash(T(rec["addr"]).replace(".", ""))
    addr = " ".join(ABBR.get(t.upper()) or proper(t) for t in a.split(" ")) if a else None
    city = proper(squash(T(rec["city"]))) or None
    co = squash(T(rec["company"]))
    co = None if not co or co.upper() in NO_COMPANY else co
    o = T(rec["optin"]).upper()
    optin = "Yes" if o in ("YES", "Y") else "No" if o in ("NO", "N") else None
    canon = NICK.get(first.upper(), first.upper())
    typed = T(rec["name"]) if rec["name"] is not None else f"{T(rec['first_t'])} {T(rec['last_t'])}"
    return dict(rec, first=first, last=last, email=e, phone=phone, zip=zip_, st=st, address=addr, cityc=city,
                co=co, opt=optin, canon=canon, typed=typed,
                kname=f"{canon}|{last.upper()}|{zip_}" if first and last and zip_ else None,
                filled=sum(v is not None for v in (e, phone, co, addr)))


# ---------------------------------------------------------------- 1. read the three sources
recs = []
with open(ROOT / "data" / "webshop-customers-2026-10.csv", encoding="utf-8") as f:
    for i, r in enumerate(csv.DictReader(f), 1):
        recs.append(dict(id=f"W-{i:04d}", src="Webshop", raw=tuple(r.values()), name=r["Full Name"], first_t="", last_t="",
                         email=r["Email"], phone=r["Phone"], company=r["Company"], addr=r["Street"], city=r["City"],
                         state=r["State"], zip=r["ZIP"], optin=r["Marketing Opt-in"],
                         date=dt.date.fromisoformat(r["Last Order"].strip()),
                         spend=float(r["Total Spent"].replace("$", "").replace(",", ""))))
with open(ROOT / "data" / "pos-customers-2026-10.csv", encoding="utf-8") as f:
    for i, r in enumerate(csv.DictReader(f), 1):
        m, d_, y = map(int, r["Last Visit"].split("/"))
        recs.append(dict(id=f"P-{i:04d}", src="POS", raw=tuple(r.values()), name=r["Customer"], first_t="", last_t="",
                         email=r["Email"], phone=r["Phone"], company="", addr=r["Address"], city=r["City"],
                         state=r["St"], zip=r["Zip"], optin="", date=dt.date(y, m, d_), spend=float(r["Spend"] or 0)))
ws = openpyxl.load_workbook(ROOT / "data" / "sales-team-accounts.xlsx").active
cols = [c.value for c in ws[3]]
company, n = None, 0
for row in ws.iter_rows(min_row=4, values_only=True):
    if all(v is None or str(v).strip() == "" for v in row):
        continue
    r = dict(zip(cols, row))
    company = r["Company"] if r["Company"] not in (None, "") else company  # merged cells: fill down
    n += 1
    v = r["Last contact"]
    date = v.date() if isinstance(v, dt.datetime) else dt.datetime.strptime(str(v).strip(), "%d %b %Y").date()
    recs.append(dict(id=f"S-{n:04d}", src="Sales sheet", raw=tuple(row), name=None, first_t=T(r["First"]),
                     last_t=T(r["Last"]), email=r["Email"], phone=r["Phone"], company=company, addr=r["Address"],
                     city=r["City"], state=r["State"], zip=r["ZIP"], optin=r["Newsletter"], date=date, spend=0.0))
for i, r in enumerate(recs):
    r["order"] = i
for r in recs:  # the values exactly as they arrived
    r.update(raw_email="" if r["email"] is None else str(r["email"]), raw_phone="" if r["phone"] is None else str(r["phone"]),
             raw_addr="" if r["addr"] is None else str(r["addr"]), raw_city="" if r["city"] is None else str(r["city"]),
             raw_state="" if r["state"] is None else str(r["state"]), raw_zip="" if r["zip"] is None else str(r["zip"]))
recs = [clean(r) for r in recs]

seen = set()
for r in recs:
    k = (r["src"], r["raw"])
    up = r["typed"].upper()
    if k in seen:
        r["status"] = "Discarded: exact duplicate row"
    elif "TEST" in up or (r["email"] or "").startswith("test@") or "WALK IN" in up.replace("-", " "):
        r["status"] = "Discarded: test or walk-in record"
    elif not r["last"] and not r["email"] and not r["phone"]:
        r["status"] = "Discarded: no name or contact"
    else:
        r["status"] = "Kept"
    seen.add(k)
kept = [r for r in recs if r["status"] == "Kept"]


# ---------------------------------------------------------------- link records (union-find), then build customers
def union_find(ids, links):
    parent = {i: i for i in ids}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for a, b in links:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra
    return find


def rule_links():
    for key in ("email", "phone", "kname"):
        first_of = {}
        for r in kept:
            if r[key]:
                if r[key] in first_of:
                    yield first_of[r[key]], r["id"]
                else:
                    first_of[r[key]] = r["id"]


SORT = lambda r: (-r["date"].toordinal(), -r["filled"], PRIORITY[r["src"]], r["order"])
find0 = union_find([r["id"] for r in kept], list(rule_links()))       # the groups the three rules make (G0)
auto = defaultdict(list)
for r in kept:
    auto[find0(r["id"])].append(r)
auto_master = {root: sorted(g, key=SORT)[0]["id"] for root, g in auto.items()}   # G0 -> its master record
rec_auto = {r["id"]: find0(r["id"]) for r in kept}


def model(merge_pairs=()):
    """Customers when the given look-alike pairs (master A, master B of two rule groups) are marked Merge."""
    links = [(auto_master[rec_auto[a]], auto_master[rec_auto[b]]) for a, b in merge_pairs]
    roots = list(auto)
    find = union_find(roots, [(rec_auto[a], rec_auto[b]) for a, b in links])
    groups = defaultdict(list)
    for r in kept:
        groups[find(rec_auto[r["id"]])].append(r)
    customers, log = {}, {}
    for g in groups.values():
        g.sort(key=SORT)
        m = g[0]
        first_val = lambda k: next((r[k] for r in g if r[k]), None)
        nm = next((r for r in g if r["last"]), m)
        ad = next((r for r in g if r["address"]), m)
        email = first_val("email")
        opts = {r["opt"] for r in g}
        srcs = sorted({r["src"] for r in g}, key=PRIORITY.get)
        customers[m["id"]] = dict(first=nm["first"], last=nm["last"], email=email, phone=first_val("phone"),
                                  other=sorted({r["email"] for r in g if r["email"] and r["email"] != email}),
                                  company=first_val("co"), address=ad["address"], city=ad["cityc"], st=ad["st"], zip=ad["zip"],
                                  opt="No" if "No" in opts else "Yes" if "Yes" in opts else "Unknown",
                                  spend=round(sum(r["spend"] for r in g), 2), n=len(g), sources=" + ".join(srcs),
                                  lastact=m["date"], members=g)
        for r in g[1:]:
            cnt = lambda k: sum(1 for x in g if x[k] and x[k] == r[k])
            why = ("Reviewer decision" if rec_auto[r["id"]] != rec_auto[m["id"]] else
                   "Same email" if r["email"] and cnt("email") > 1 else "Same phone" if r["phone"] and cnt("phone") > 1
                   else "Same name and ZIP")
            log[r["id"]] = (m["id"], why)
    return dict(customers=customers, groups=list(groups.values()), log=log)


BASE = model()
customers, groups = BASE["customers"], BASE["groups"]
pairs = {}   # look-alike pairs on the rule groups: (master A, master B) -> reason
for keyf, why in ((lambda r: f"{r['canon'][:1]}|{r['last'].upper()}|{r['zip']}" if r["first"] and r["last"] and r["zip"] else None,
                   "Similar first name, same last name and ZIP"),
                  (lambda r: f"{r['canon']}|{r['last'].upper()}" if r["first"] and r["last"] else None, "Same name, different ZIP")):
    bucket = defaultdict(set)
    for r in kept:
        k = keyf(r)
        if k:
            bucket[k].add(auto_master[rec_auto[r["id"]]])
    for ms in bucket.values():
        ms = sorted(ms)
        for i, a in enumerate(ms):
            for b in ms[i + 1:]:
                pairs.setdefault((a, b), why)

received = Counter(r["src"] for r in recs)
status = Counter(r["status"] for r in recs)
merged = len(kept) - len(customers)
print(f"Python: received {len(recs)}, discarded {len(recs) - len(kept)}, customers {len(customers)}, merged {merged}, review {len(pairs)}")


# ---------------------------------------------------------------- data quality before / after, field completeness
def is_email(s):
    return (s != "" and s == s.lower() and " " not in s and s.count("@") == 1 and s.split("@")[0] != ""
            and "." in s.split("@")[1])


def is_phone(s):
    d = re.sub(r"\D", "", s)
    return len(d) == 10 and s == f"({d[:3]}) {d[3:6]}-{d[6:]}"


def is_zip(s):
    return len(s) == 5 and s.isdigit()


S = lambda v: "" if v is None else str(v)
RAWF = ("raw_email", "raw_phone", "raw_addr", "raw_city", "raw_state", "raw_zip")


def quality(custs):
    n, c = len(recs), list(custs.values())
    return {"Valid email": (sum(is_email(r["raw_email"]) for r in recs), n, sum(is_email(S(x["email"])) for x in c), len(c)),
            "Formatted phone": (sum(is_phone(r["raw_phone"]) for r in recs), n, sum(is_phone(S(x["phone"])) for x in c), len(c)),
            "ZIP complete": (sum(is_zip(r["raw_zip"]) for r in recs), n, sum(is_zip(S(x["zip"])) for x in c), len(c)),
            "Fields filled": (sum(sum(r[f].strip() != "" for f in RAWF) for r in recs), 6 * n,
                              sum(sum(x[k] is not None for k in ("email", "phone", "address", "city", "st", "zip")) for x in c), 6 * len(c))}


QUAL = quality(customers)
HEAT_FIELDS = [("Email", "email", "email"), ("Phone", "phone", "phone"), ("Company", "co", "company"),
               ("Address", "address", "address"), ("City", "cityc", "city"), ("State", "st", "st"), ("ZIP", "zip", "zip"),
               ("Marketing opt-in", "opt", "opt")]


def heat(custs):
    out = []
    for label, rk, ck in HEAT_FIELDS:
        row = []
        for src in ("Webshop", "POS", "Sales sheet"):
            rs = [r for r in kept if r["src"] == src]
            row.append(sum(r[rk] is not None for r in rs) / len(rs))
        row.append(sum((x[ck] != "Unknown") if ck == "opt" else (x[ck] is not None) for x in custs.values()) / len(custs))
        out.append(row)
    return out


squash_all = lambda s: " ".join(p for p in s.split(" ") if p)  # Excel TRIM

# ---------------------------------------------------------------- 2. against the planted answer
exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
truth = exp["truth"]
check("Rows read per source = rows written", all(received[s] == exp["rows"][k] for s, k in
                                                  (("Webshop", "W"), ("POS", "P"), ("Sales sheet", "S"))), received)
check("Exact duplicate rows found = planted (3 webshop + 2 POS)", status["Discarded: exact duplicate row"] == 5)
check("Test and walk-in rows found = planted", sum(v == "junk" for v in truth.values()) ==
      status["Discarded: test or walk-in record"] + status["Discarded: no name or contact"])
check("Every discarded row is a planted duplicate or test row, every kept row a real person",
      all((r["status"] == "Kept") == isinstance(truth[r["id"]], int) for r in recs))
pure = all(len({truth[r["id"]] for r in g}) == 1 for g in groups)
check("Every merged group is one real person (no false merge)", pure)
check("Every real person ends up as exactly one customer (no missed duplicate)",
      len(customers) == exp["customers"] == len({truth[r["id"]] for r in kept}))
check("Records merged away = planted", merged == exp["merged_away"], f"{merged} vs {exp['merged_away']}")
check("Lifetime spend of all customers = planted spend", round(sum(c["spend"] for c in customers.values()), 2) == exp["lifetime_spend"])
pid_master = {truth[m]: m for m in customers}
planted = {tuple(sorted((pid_master[a], pid_master[b]))): why for a, b, why in exp["review_pairs"]}
check(f"Look-alike pairs flagged for review = the {len(planted)} planted pairs", pairs == planted,
      f"extra {set(pairs) - set(planted)}, missing {set(planted) - set(pairs)}")
check("Look-alikes are NOT merged (they stay two customers)", all(a != b for a, b in planted))
check("No email belongs to two customers", len({c["email"] for c in customers.values() if c["email"]}) ==
      sum(1 for c in customers.values() if c["email"]))
check("Opt-out wins: a customer is opted in only if no record says No",
      all(c["opt"] != "Yes" or all(r["opt"] != "No" for r in c["members"]) for c in customers.values()))
check("Data quality is better after the cleanup on all four measures (Python)",
      all(a / b > x / y for x, y, a, b in QUAL.values()), QUAL)

# ---------------------------------------------------------------- 3. against the workbook
try:
    import win32com.client as win32
except ImportError:
    print("SKIP  workbook checks (pywin32 not installed)")
    win32 = None


def table(wb, sheet, name):
    lo = wb.Worksheets(sheet).ListObjects(name)
    cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
    if lo.DataBodyRange is None:
        return []
    return [dict(zip(cols, r)) for r in lo.DataBodyRange.Value]


def xdate(v):
    return dt.date(v.year, v.month, v.day) if v else None


def hx(h):
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


COPPER = hx("B5754A")
near = lambda a, b: abs(float(a) - float(b)) < 1e-9


def check_workbook(wb, M=BASE, tag=""):
    custs = M["customers"]
    cust = table(wb, "Customers", "Customers")
    check(f"{tag}Workbook customers = Python ({len(custs)})", len(cust) == len(custs), f"{len(cust)}")
    bad = []
    for c in cust:
        p = custs.get(c["Master record"])
        if p is None:
            bad.append(f"{c['Master record']} not a Python master"); continue
        got = (c["First name"], c["Last name"], c["Email"], c["Phone"], c["Company"], c["Address"], c["City"], c["State"],
               c["ZIP"], c["Marketing opt-in"], round(c["Lifetime spend"], 2), int(c["Records merged"]), c["Sources"],
               xdate(c["Last activity"]), "; ".join(p["other"]) or None)
        want = (p["first"], p["last"], p["email"], p["phone"], p["company"], p["address"], p["city"], p["st"], p["zip"],
                p["opt"], p["spend"], p["n"], p["sources"], p["lastact"], "; ".join(p["other"]) or None)
        got = tuple(None if g == "" else g for g in got)
        if got != want:
            bad.append(f"{c['Master record']}: " + ", ".join(f"{g!r}<>{w!r}" for g, w in zip(got, want) if g != w))
    check(f"{tag}Every customer field = Python (survivorship, other emails, opt-in, spend, sources)", not bad, "; ".join(bad[:3]))
    ids = {c["Master record"]: c["Customer ID"] for c in cust}
    log = table(wb, "Merge log", "MergeLog")
    py_log = {rid: (ids.get(m), why) for rid, (m, why) in M["log"].items()}
    x_log = {r["Record"]: (r["Merged into"], r["Matched by"]) for r in log}
    check(f"{tag}Merge log = Python: every merged record, its customer and the rule ({len(py_log)})", x_log == py_log,
          f"{len(x_log)} rows; first diff {next(((k, x_log.get(k), v) for k, v in py_log.items() if x_log.get(k) != v), None)}")
    rev = table(wb, "Needs review", "Review")
    x_pairs = {tuple(sorted((r["Master A"], r["Master B"]))): r["Reason"] for r in rev}
    check(f"{tag}Needs review = Python ({len(pairs)} pairs, with the reason)", x_pairs == pairs, f"{len(x_pairs)}")
    allr = table(wb, "All records", "AllRecords")
    xs = Counter(("Discarded" if r["Status"].startswith("Discarded") else r["Status"]) for r in allr)
    check(f"{tag}All records: every source row once, with the right status",
          len(allr) == len(recs) and {r["Record"]: r["Status"] for r in allr if r["Status"].startswith("Discarded")} ==
          {r["id"]: r["status"] for r in recs if r["status"] != "Kept"}, f"{xs}")
    by_id = {r["id"]: r for r in recs}
    bad_raw = [r["Record"] for r in allr if tuple(S(r[f"{k} as received"]) for k in ("Email", "Phone", "Address", "City", "State", "ZIP"))
               != tuple(by_id[r["Record"]][f] for f in RAWF)]
    check(f"{tag}All records: the 'as received' columns = the source files, character for character", not bad_raw, bad_raw[:3])
    crm = table(wb, "CRM import", "CRMImport")
    check(f"{tag}CRM import: one row per customer, same emails", len(crm) == len(cust) and
          sorted(str(r["Email"] or "") for r in crm) == sorted(str(c["Email"] or "") for c in cust))
    # data quality score
    q = {r["Metric"]: r for r in table(wb, "Quality", "Quality")}
    qp = quality(custs)
    check(f"{tag}Data quality score before/after = Python (4 measures, counts and %)",
          set(q) == set(qp) and all((q[m]["Before ok"], q[m]["Before of"], q[m]["After ok"], q[m]["After of"]) == qp[m]
                                    and near(q[m]["Before"], qp[m][0] / qp[m][1]) and near(q[m]["After"], qp[m][2] / qp[m][3])
                                    for m in qp), {m: (tuple(q[m][k] for k in ("Before ok", "Before of", "After ok", "After of")) if m in q else None, qp[m]) for m in qp})
    s = wb.Worksheets("Summary")
    # cockpit: KPI cards
    n_open = sum(1 for r in rev if not r["Decision"])
    want = (len(recs), len(recs) - len(kept), len(kept) - len(custs), len(custs), len(pairs))
    got = tuple(s.Range(c).Value for c in ("D8", "I8", "N8", "S8", "X8"))
    check(f"{tag}Cockpit KPI cards = Python {want}", got == want, got)
    check(f"{tag}Cockpit review card: decided / open = the Decision column ({len(pairs) - n_open} / {n_open})",
          s.Range("X9").Value == f"{len(pairs) - n_open} decided  ·  {n_open} open", s.Range("X9").Value)
    # gauges: each doughnut plots before (inner) and after (outer) of the Python score; centre label linked
    ok, why = True, []
    for k, m in enumerate(("Valid email", "Formatted phone", "ZIP complete", "Fields filled"), 1):
        ch = s.ChartObjects(f"gauge_{k}").Chart
        b, a = ch.SeriesCollection(1).Values[0], ch.SeriesCollection(2).Values[0]
        x = qp[m]
        if not (near(b, x[0] / x[1]) and near(a, x[2] / x[3])):
            ok = False; why.append((m, b, a))
    check(f"{tag}Cockpit gauges: before/after rings = Python for all four measures", ok, why)
    # funnel
    ch = s.ChartObjects("funnel").Chart
    fv = tuple(ch.SeriesCollection(2).Values)
    check(f"{tag}Cockpit funnel: rows in -> kept -> unique customers = Python ({len(recs)}, {len(kept)}, {len(custs)})",
          fv == (len(recs), len(kept), len(custs)), fv)
    # rule chart
    rv = tuple(s.ChartObjects("rules").Chart.SeriesCollection(1).Values)
    pr = Counter(why for _, why in M["log"].values())
    check(f"{tag}Cockpit merged-by-rule chart = Python", rv == tuple(pr.get(k, 0) for k in
                                                                    ("Same email", "Same phone", "Same name and ZIP", "Reviewer decision")), rv)
    # heatmap
    hp = heat(custs)
    hw = s.Range("I31:L38").Value
    check(f"{tag}Cockpit field-completeness heatmap = Python (8 fields x 3 sources + merged list)",
          all(near(hw[i][j], hp[i][j]) for i in range(8) for j in range(4)),
          [(HEAT_FIELDS[i][0], j, hw[i][j], hp[i][j]) for i in range(8) for j in range(4) if not near(hw[i][j], hp[i][j])][:3])
    # data checks
    rng = wb.Names("DataChecks").RefersToRange
    vals = [c.Value for c in rng.Cells]
    check(f"{tag}All {len(vals)} data checks in the workbook PASS", len(vals) == 9 and all(v == "PASS" for v in vals),
          [(s.Range(f"S{c.Row}").Value, c.Value) for c in rng.Cells if c.Value != "PASS"])
    return cust


def check_example(wb, M=BASE, tag=""):
    """Before/after card on the cockpit: the example customer's source rows exactly as received vs its golden record."""
    s = wb.Worksheets("Summary")
    cid = s.Range("Z43").Value
    cust = {c["Customer ID"]: c["Master record"] for c in table(wb, "Customers", "Customers")}
    p = M["customers"][cust[cid]]
    golden = [f"{p['first']} {p['last']}", S(p["email"]), S(p["phone"]), S(p["address"]),
              squash_all(f"{S(p['city'])} {S(p['st'])} {S(p['zip'])}")]
    cols = ("H", "L", "P", "S", "V")  # Name, Email, Phone, Address, City-State-ZIP
    got_rows = {}
    copper_ok, changed = True, 0
    for k in range(4):
        r = 46 + k
        rid = s.Range(f"D{r}").Value
        if not rid:
            continue
        vals = [s.Range(f"{c}{r}").Value for c in cols]
        got_rows[rid] = vals
        for c, v, g in zip(cols, vals, golden):
            diff = v not in ("", "(blank)") and v != g
            changed += diff
            if (s.Range(f"{c}{r}").DisplayFormat.Font.Color == COPPER) != diff:
                copper_ok = False
    want_rows = {}
    for m in p["members"]:
        raw = [m["typed"], m["raw_email"], m["raw_phone"], m["raw_addr"],
               squash_all(f"{m['raw_city']} {m['raw_state']} {m['raw_zip']}")]
        want_rows[m["id"]] = [v if v != "" else "(blank)" for v in raw]
    check(f"{tag}Before/after card ({cid}): every source row, exactly as received = Python", got_rows == want_rows,
          (got_rows, want_rows))
    gold = [s.Range(f"{c}50").Value for c in cols]
    check(f"{tag}Before/after card: golden record = Python customer", gold == golden, (gold, golden))
    want_changed = sum(1 for vals in want_rows.values() for v, g in zip(vals, golden) if v != "(blank)" and v != g)
    check(f"{tag}Before/after card: {want_changed} changed values highlighted in copper, no others", copper_ok and changed == want_changed,
          (changed, want_changed))


def check_explorer(wb, xl, ids, M=BASE, tag=""):
    """Match explorer: pick customers from the drop-down, compare every record and its linking rule with Python."""
    ws = wb.Worksheets("Match explorer")
    keep = ws.Range("D6").Value
    cust = {c["Customer ID"]: c["Master record"] for c in table(wb, "Customers", "Customers")}
    bad = []
    for cid in ids:
        ws.Range("D6").Value = cid
        xl.Calculate()
        p = M["customers"][cust[cid]]
        want = {}
        for m in p["members"]:
            why = "Master (newest record)" if m["id"] == cust[cid] else M["log"][m["id"]][1]
            want[m["id"]] = (m["src"], why, m["raw_email"], m["raw_phone"], m["email"] or "", round(m["spend"], 2))
        got = {}
        r = 16
        while ws.Range(f"B{r}").Value:
            v = ws.Range(f"B{r}:M{r}").Value[0]
            got[v[0]] = (v[1], v[3], v[5], v[6], v[8], round(v[11], 2))
            r += 1
        name_ok = ws.Range("D9").Value == f"{p['first']} {p['last']}" and ws.Range("F9").Value == S(p["email"])
        if got != want or not name_ok or ws.Range("D16").Value != "Master":
            bad.append((cid, got, want))
    ws.Range("D6").Value = keep
    xl.Calculate()
    check(f"{tag}Match explorer: {len(ids)} customers picked, every record + linking rule + golden card = Python", not bad, bad[:1])


def check_nav(wb):
    sheets = {ws.Name for ws in wb.Worksheets}
    n, bad = 0, []
    for ws in wb.Worksheets:
        for shp in ws.Shapes:
            if shp.Name.startswith("nav_"):
                n += 1
                try:
                    target = shp.Hyperlink.SubAddress.split("!")[0].strip("'")
                except Exception:
                    target = None
                if target not in sheets:
                    bad.append((ws.Name, shp.Name, target))
    check(f"Navigation: all {n} menu buttons link to an existing sheet", n >= 8 * 8 and not bad, bad[:3])


def decisions(wb):
    return {(r["Master A"], r["Master B"]): (r["Decision"] or "", r["Result"]) for r in table(wb, "Needs review", "Review")}


if win32:
    xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(ROOT / "Customer-Cleanup.xlsx"), 0, True)
        check_workbook(wb)
        check_example(wb)
        cust = table(wb, "Customers", "Customers")
        multi = [c["Customer ID"] for c in cust if c["Records merged"] >= 2]
        three = [c["Customer ID"] for c in cust if c["Records merged"] == 3]
        picks = [str(wb.Worksheets("Match explorer").Range("D6").Value).split()[0]] + three[:2] + multi[::250][:3] + [cust[0]["Customer ID"]]
        check_explorer(wb, xl, picks)
        check_nav(wb)
        dec0 = decisions(wb)
        want0 = {k: ("Keep separate", "Kept as two customers") if why.startswith("Similar") else ("", "Open: waiting for a decision")
                 for k, why in pairs.items()}
        check("Needs review: shipped decisions = 6 Keep separate, 5 open", dec0 == want0, dec0)
        wb.Close(False)

        import shutil, tempfile
        with tempfile.TemporaryDirectory() as tmp:
            shutil.copy(ROOT / "Customer-Cleanup.xlsx", tmp)
            shutil.copytree(ROOT / "data", Path(tmp) / "data")
            wb = xl.Workbooks.Open(str(Path(tmp) / "Customer-Cleanup.xlsx"))
            wb.RefreshAll(); xl.CalculateUntilAsyncQueriesDone(); xl.Calculate()
            folder = wb.Worksheets("Settings").Range("B4").Value
            check("Refresh in another folder reads that folder's data", str(folder).lower().startswith(str(tmp).lower()[:12]), folder)
            check_workbook(wb, tag="After refresh: ")
            check_example(wb, tag="After refresh: ")
            check("After refresh: the typed decisions survive Refresh All", decisions(wb) == want0)

            # a person marks one open look-alike pair Merge -> on the next refresh the two customers become one
            lo = wb.Worksheets("Needs review").ListObjects("Review")
            target = None
            for i in range(1, lo.ListRows.Count + 1):
                row = lo.ListRows(i).Range
                if row.Cells(1, 3).Value.startswith("Same name") and target is None:
                    row.Cells(1, 1).Value = "Merge"
                    target = (row.Cells(1, 17).Value, row.Cells(1, 18).Value)
            wb.RefreshAll(); xl.CalculateUntilAsyncQueriesDone(); xl.Calculate()
            M1 = model([target])
            check(f"Decision Merge on {target[0]} + {target[1]}: one customer fewer ({len(M1['customers'])})",
                  len(M1["customers"]) == len(customers) - 1)
            check_workbook(wb, M1, "After Merge decision: ")
            d1 = decisions(wb)
            cust1 = {c["Master record"]: c["Customer ID"] for c in table(wb, "Customers", "Customers")}
            merged_master = next(m for m, c in M1["customers"].items() if any(x["id"] in target for x in c["members"]))
            check("After Merge decision: the pair stays listed, Decision kept, Result = 'Merged into' its customer",
                  d1[target] == ("Merge", f"Merged into {cust1[merged_master]}") and
                  all(d1[k] == v for k, v in want0.items() if k != target), d1[target])
            check("After Merge decision: the absorbed records are logged as 'Reviewer decision'",
                  sum(1 for _, w in M1["log"].values() if w == "Reviewer decision") >= 1 and
                  sum(1 for r in table(wb, "Merge log", "MergeLog") if r["Matched by"] == "Reviewer decision") ==
                  sum(1 for _, w in M1["log"].values() if w == "Reviewer decision"))
            mid = cust1[merged_master]
            check_explorer(wb, xl, [mid], M1, "After Merge decision: ")
            wb.RefreshAll(); xl.CalculateUntilAsyncQueriesDone(); xl.Calculate()
            check("Refresh again: the Merge is stable (same customers, decision still there)",
                  len(table(wb, "Customers", "Customers")) == len(M1["customers"]) and decisions(wb)[target][0] == "Merge")
            wb.Close(False)
    finally:
        xl.Quit()

print(f"\n{sum(results)}/{len(results)} checks passed")
sys.exit(0 if all(results) else 1)
