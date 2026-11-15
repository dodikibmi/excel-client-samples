"""Independent checks for Theme-Gallery.xlsx: every number is recomputed in Python from data/sales.csv, every colour
is read back from the saved file (cell fills, fonts, conditional formats, chart XML) and compared with themes.py,
and the contrast of the colours actually used is measured (WCAG). Excel is not needed.
    python test/verify.py"""
import csv, re, sys, zipfile
from collections import defaultdict
from pathlib import Path
import warnings
from openpyxl import load_workbook

warnings.filterwarnings("ignore", module="openpyxl")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from themes import THEMES, ROLES, SWATCHES, sheet_name, contrast

XLSX = ROOT / "Theme-Gallery.xlsx"
results = []


def check(name, ok, detail=""):
    results.append((bool(ok), name, detail))


def near(a, b, tol=0.6):
    return a is not None and b is not None and abs(float(a) - float(b)) <= tol


def hexof(color):
    if color is None or color.rgb is None or not isinstance(color.rgb, str):
        return None
    return "#" + color.rgb[-6:].upper()


# ---------------------------------------------------------------- python side
rows = list(csv.DictReader((ROOT / "data" / "sales.csv").open(encoding="utf-8")))
for r in rows:
    for k in ("Revenue", "Budget", "Cost", "Orders"):
        r[k] = int(r[k])
REP, PY = "2026-10-01", "2025-10-01"
MONTHS = sorted({r["Month"] for r in rows})[-12:]


def tot(field, **crit):
    return sum(r[field] for r in rows if all(r[k] == v for k, v in crit.items()))


rev, bud, cost, orders = (tot(f, Month=REP) for f in ("Revenue", "Budget", "Cost", "Orders"))
rev_py, cost_py, orders_py = (tot(f, Month=PY) for f in ("Revenue", "Cost", "Orders"))
gm, gm_py = 1 - cost / rev, 1 - cost_py / rev_py
avg, avg_py = rev / orders, rev_py / orders_py
REGIONS = ["North", "South", "East", "West"]
reg = {g: (tot("Revenue", Month=REP, Region=g), tot("Budget", Month=REP, Region=g), tot("Cost", Month=REP, Region=g))
       for g in REGIONS}


def note(x, fmt, tail):
    if round(x, 3) == 0:
        return f"● flat{tail}"
    return ("▲ " if x >= 0 else "▼ ") + format(abs(x), fmt) + tail


expected_notes = {"B9": note(rev / bud - 1, ".1%", " vs budget"), "B10": note(rev / rev_py - 1, ".1%", " vs prior year"),
                  "H9": note((gm - gm_py) * 100, ".1f", " pts vs prior year"),
                  "N9": note(orders / orders_py - 1, ".1%", " vs prior year"),
                  "T9": note(avg / avg_py - 1, ".1%", " vs prior year")}

# ---------------------------------------------------------------- workbook
wb = load_workbook(XLSX, data_only=True)
wbf = load_workbook(XLSX)  # formulas + formats
names = [sheet_name(i) for i in range(len(THEMES))]
check("sheets: Themes, 8 theme pages, Palette, How to use, Calc, Data in that order",
      wb.sheetnames == ["Themes"] + names + ["Palette", "How to use", "Calc", "Data"], wb.sheetnames)
check("no macros: .xlsx without vbaProject", "xl/vbaProject.bin" not in zipfile.ZipFile(XLSX).namelist())

data = wb["Data"]
vals = [tuple(row) for row in data.iter_rows(min_row=2, values_only=True)]
check(f"Data table holds all {len(rows)} CSV rows with the same amounts",
      len(vals) == len(rows) and all(v[3:] == (r["Revenue"], r["Budget"], r["Cost"], r["Orders"]) and v[1] == r["Region"]
                                       and v[2] == r["Category"] for v, r in zip(vals, rows)))

errors = [(ws.title, c.coordinate, c.value) for ws in wb for row in ws.iter_rows() for c in row
          if isinstance(c.value, str) and c.value.startswith("#") and c.value[1:4] in ("REF", "N/A", "VAL", "DIV", "NAM", "NUM")]
check("no #REF!, #N/A, #VALUE!, #DIV/0! anywhere", not errors, errors[:5])

calc = wb["Calc"]
series = [(calc[f"B{4 + k}"].value, calc[f"C{4 + k}"].value, calc[f"D{4 + k}"].value) for k in range(12)]
py_series = [(tot("Revenue", Month=m), tot("Budget", Month=m),
              tot("Revenue", Month=f"{int(m[:4]) - 1}{m[4:]}")) for m in MONTHS]
check("12-month chart data (actual, budget, prior year) = Python", series == py_series)
cats = ["Fasteners", "Power Tools", "Paint", "Garden", "Electrical"]
check("category doughnut data = Python",
      all(calc[f"B{18 + k}"].value == tot("Revenue", Month=REP, Category=c) for k, c in enumerate(cats)))
worst = min(REGIONS, key=lambda g: reg[g][0] / reg[g][1])
check("story: West is the region furthest below budget in Oct 2026", worst == "West" and reg["West"][0] / reg["West"][1] - 1 < -0.10,
      f"{worst} {reg[worst][0] / reg[worst][1] - 1:.1%}")

for i, t in enumerate(THEMES):
    ws, wf = wb[names[i]], wbf[names[i]]
    tag = f"theme {i + 1} {t['name']}"
    kp = (ws["B8"].value, ws["H8"].value, ws["N8"].value, ws["T8"].value)
    check(f"{tag}: KPI cards = Python (revenue, margin, orders, average order)",
          near(kp[0], rev) and near(kp[1], gm, 1e-9) and near(kp[2], orders) and near(kp[3], avg, 1e-6), kp)
    got = {a: ws[a].value for a in expected_notes}
    check(f"{tag}: ▲/▼ notes = Python", got == expected_notes, got)
    ok = True
    for k, g in enumerate(REGIONS + ["Total"]):
        r = 35 + k
        a, b, cst = (sum(v[j] for v in reg.values()) for j in range(3)) if g == "Total" else reg[g]
        want = (g, a, b, a - b if g != "Total" else sum(reg[x][0] - reg[x][1] for x in REGIONS), a / b - 1, 1 - cst / a,
                a / sum(v[0] for v in reg.values()))
        row = tuple(ws[f"{col}{r}"].value for col in "BDEFHIJ")
        ok &= row[0] == want[0] and all(near(x, y, 1e-6) for x, y in zip(row[1:], want[1:]))
    check(f"{tag}: region table = Python (revenue, budget, variance, var %, margin, share)", ok)

    fills = {"header": hexof(wf["B2"].fill.fgColor), "page": hexof(wf["A6"].fill.fgColor),
             "accent": hexof(wf["C5"].fill.fgColor), "card": hexof(wf["C9"].fill.fgColor),
             "card2": hexof(wf["C36"].fill.fgColor)}
    check(f"{tag}: header, page, gold rule, cards and stripes painted in the theme colours",
          all(fills[k] == t[k] for k in fills), fills)
    fonts = {"text": hexof(wf["B8"].font.color), "muted": hexof(wf["B7"].font.color),
             "htext": hexof(wf["B3"].font.color), "hmuted": hexof(wf["B4"].font.color), "accent": hexof(wf["B34"].font.color)}
    check(f"{tag}: title, KPI, label and table-header fonts in the theme colours",
          all(fonts[k] == t[k] for k in fonts) and wf["B3"].font.name == t["title_font"], fonts)
    sw = {role: (hexof(wf[f"{col}35"].fill.fgColor), ws[f"{col}38"].value) for col, role in
          zip(["N", "O", "P", "Q", "R", "T", "U", "V", "W", "X"], SWATCHES)}
    check(f"{tag}: colour strip shows the right 10 swatches with matching hex codes",
          all(v == (t[r], t[r]) for r, v in sw.items()), sw)
    cf = {}
    for rng, rules in wf.conditional_formatting._cf_rules.items():
        for rule in rules:
            if rule.dxf is not None and rule.dxf.font is not None and rule.dxf.font.color is not None:
                cf.setdefault(str(rng.sqref), set()).add(hexof(rule.dxf.font.color))
    want_cf = {t["good"], t["bad"]}
    cells = ["B9", "B10", "H9", "N9", "T9", "F35", "H38"]
    check(f"{tag}: red/green conditional formats use the theme's good/bad colours",
          all(cf.get(c) == want_cf for c in cells), {c: cf.get(c) for c in cells})
    measured = {
        "text/card": contrast(fonts["text"], fills["card"]), "muted/card": contrast(fonts["muted"], fills["card"]),
        "htext/header": contrast(fonts["htext"], fills["header"]),
        "hmuted/header": contrast(fonts["hmuted"], fills["header"]),
        "accent/card": contrast(t["accent"], fills["card"]), "good/card": contrast(t["good"], fills["card"]),
        "bad/card": contrast(t["bad"], fills["card"]), "text/stripe": contrast(fonts["text"], fills["card2"])}
    floors = {"text/card": 4.5, "muted/card": 4.5, "htext/header": 4.5, "hmuted/header": 4.5, "text/stripe": 4.5,
              "accent/card": 3, "good/card": 3, "bad/card": 3}
    check(f"{tag}: WCAG AA contrast measured on the painted cells",
          all(measured[k] >= floors[k] for k in floors), {k: round(v, 1) for k, v in measured.items()})

# ---------------------------------------------------------------- charts and buttons (raw XML)
z = zipfile.ZipFile(XLSX)
wbx = z.read("xl/workbook.xml").decode()
rels = z.read("xl/_rels/workbook.xml.rels").decode()
rid_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels)) | \
    {a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rels)}
sheet_file = {}
for name, rid in re.findall(r'<sheet name="([^"]+)" sheetId="\d+" r:id="(rId\d+)"', wbx):
    sheet_file[name.replace("&amp;", "&")] = "xl/" + rid_target[rid].lstrip("/").replace("xl/", "")


def rel_targets(part):
    p = part.rsplit("/", 1)
    rp = f"{p[0]}/_rels/{p[1]}.rels"
    if rp not in z.namelist():
        return {}
    x = z.read(rp).decode()
    out = {}
    for m in re.finditer(r"<Relationship [^>]*>", x):
        tag = m.group(0)
        out[re.search(r'Id="([^"]+)"', tag).group(1)] = re.search(r'Target="([^"]+)"', tag).group(1)
    return out


def resolve(base, target):
    parts = base.rsplit("/", 1)[0].split("/")
    for seg in target.split("/"):
        if seg == "..":
            parts.pop()
        elif seg != ".":
            parts.append(seg)
    return "/".join(parts)


for i, t in enumerate(THEMES):
    tag = f"theme {i + 1} {t['name']}"
    sf = sheet_file[names[i]]
    drawing = [resolve(sf, v) for v in rel_targets(sf).values() if "drawing" in v and "vml" not in v.lower()][0]
    dr = rel_targets(drawing)
    charts = sorted(resolve(drawing, v) for v in dr.values() if "charts/chart" in v)
    links = sorted(v.replace("&amp;", "&") for v in dr.values() if v.startswith("#"))
    want_links = sorted(f"#'{n}'!A1" for n in (names[(i - 1) % 8], "Themes", names[(i + 1) % 8]))
    check(f"{tag}: Previous / All themes / Next buttons link to the right sheets", links == want_links, links)
    kinds, colours = [], []
    for cpart in charts:
        x = z.read(cpart).decode()
        kinds.append("doughnut" if "doughnutChart" in x else "bar+line" if "barChart" in x and "lineChart" in x else "?")
        colours.append(set(re.findall(r'srgbClr val="([0-9A-F]{6})"', x)))
    ok = sorted(kinds) == ["bar+line", "doughnut"]
    if ok:
        bar = colours[kinds.index("bar+line")]
        dn = colours[kinds.index("doughnut")]
        ok = {t["accent"][1:], t["accent2"][1:], t["muted"][1:], t["s2"][1:]} <= bar and \
            {t[r][1:] for r in ("accent", "s2", "s3", "s4", "accent2")} <= dn and t["card"][1:] in bar
    check(f"{tag}: 2 charts (12-month bars + lines, category doughnut) drawn in the theme colours", ok, kinds)

idx = wbf["Themes"]
links = {c.coordinate: c.hyperlink.location for row in idx.iter_rows() for c in row if c.hyperlink}
want = {f"{c1}{r0 + 5}": f"'{names[i]}'!A1" for i, ((c1, c2), r0) in
        enumerate(zip([("B", "F"), ("H", "L"), ("N", "R"), ("T", "X")] * 2, [7] * 4 + [14] * 4))}
check("Themes: 8 tiles, each 'Open theme N' link opens its own sheet", links == want, links)
tiles = all(hexof(idx[f"{c1}{r0}"].fill.fgColor) == THEMES[i]["header"] for i, ((c1, _), r0) in
            enumerate(zip([("B", "F"), ("H", "L"), ("N", "R"), ("T", "X")] * 2, [7] * 4 + [14] * 4)))
check("Themes: every tile header painted in its theme's header colour", tiles)

pal = wbf["Palette"]
ok = True
for j, (role, _) in enumerate(ROLES):
    for k, t in enumerate(THEMES):
        cell = pal.cell(7 + j, 3 + k)
        ok &= cell.value == t[role].upper() and hexof(cell.fill.fgColor) == t[role]
check(f"Palette: {len(ROLES)} roles x {len(THEMES)} themes, every cell painted in the colour its hex code names", ok)
check("Palette: font pairing listed for every theme",
      all(pal.cell(7 + len(ROLES), 3 + k).value == t["title_font"] and pal.cell(8 + len(ROLES), 3 + k).value == t["body_font"]
          for k, t in enumerate(THEMES)))

passed = sum(ok for ok, _, _ in results)
for ok, name, detail in results:
    print(("PASS " if ok else "FAIL ") + name + ("" if ok else f"   -> {detail}"))
print(f"\n{passed}/{len(results)} checks passed")
sys.exit(0 if passed == len(results) else 1)
