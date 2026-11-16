"""Builds PDF-Invoices-to-Excel.xlsx with real Excel (COM).

Power Query reads every PDF in the invoices folder with the PDF connector (Pdf.Tables), finds the supplier from the
first page, reads the header fields through an editable label table (Invoice No / Invoice # / Inv. Number /
Document no. all mean the same field), reads the line table through an editable column table, turns the dates the
way each supplier writes them into real dates, and checks every invoice: lines = subtotal, subtotal + tax = total,
the same invoice sent twice. Files it cannot read (scanned pictures, a password, a broken download) are listed with
the reason and a next step, and can be typed once on the Manual entry sheet; they are then counted on every refresh.

Run with a Python that has pywin32:  python scripts/build_workbook.py
The workbook has no macros; the client only needs Data -> Refresh All.
"""
import csv
import datetime as dt
import json
import sys
import time
from pathlib import Path

import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "PDF-Invoices-to-Excel.xlsx"
REPORT_DATE = dt.date(2026, 10, 31)
PERIOD = (dt.date(2026, 9, 1), dt.date(2026, 10, 31))

SUPPLIERS = [  # supplier, text that identifies it on page 1, date format, account code, account name, terms
    ("Brightway Janitorial Co.", "Brightway Janitorial", "dd/MM/yyyy", "6110", "Cleaning & janitorial", 30),
    ("Summit Landscaping Group", "Summit Landscaping", "dd/MM/yyyy", "6120", "Landscaping & grounds", 15),
    ("Northgate Electrical Services", "Northgate Electrical", "d MMM yyyy", "6210", "Repairs - electrical", 30),
    ("Riverbend Plumbing & Heating", "Riverbend Plumbing", "d MMM yyyy", "6220", "Repairs - plumbing & HVAC", 14),
    ("Clearwater Maintenance Supply", "Clearwater Maintenance", "yyyy-MM-dd", "6300", "Maintenance supplies", 30),
    ("Keystone Office & Janitorial", "Keystone Office", "yyyy-MM-dd", "6310", "Janitorial & office supplies", 30),
    ("Atlas Elevator Maintenance", "Atlas Elevator", "M/d/yyyy", "6230", "Elevator service", 45),
    ("Pinecrest Security Systems", "Pinecrest Security", "M/d/yyyy", "6240", "Security & access", 30),
]
FIELD_LABELS = [
    ("Invoice No", "Invoice No"), ("Invoice #", "Invoice No"), ("Inv. Number", "Invoice No"), ("Document no.", "Invoice No"),
    ("Invoice Date", "Invoice Date"), ("Issued", "Invoice Date"), ("Inv. Date", "Invoice Date"), ("Date", "Invoice Date"),
    ("Due Date", "Due Date"), ("Payment due", "Due Date"), ("Pay by", "Due Date"),
    ("Terms", "Terms"),
    ("PO Number", "PO"), ("Your ref", "PO"), ("Customer PO", "PO"), ("Purchase order", "PO"),
    ("Site", "Property"), ("Job site", "Property"), ("Ship To", "Property"), ("Location", "Property"),
    ("Subtotal", "Subtotal"), ("Net", "Subtotal"), ("Merchandise", "Subtotal"),
    ("Sales Tax", "Tax"), ("Tax", "Tax"),
    ("Total Due", "Total"), ("Amount due", "Total"), ("Invoice total", "Total"), ("Balance due", "Total"),
    ("Credit total", "Total"),
]
COLUMN_LABELS = [
    ("Line", "Line"), ("SKU", "SKU"),
    ("Description", "Description"), ("Item", "Description"), ("Service", "Description"),
    ("Qty", "Qty"), ("Hrs/Qty", "Qty"), ("Quantity", "Qty"),
    ("Unit Price", "Unit price"), ("Rate", "Unit price"), ("Unit $", "Unit price"), ("Price", "Unit price"),
    ("Amount", "Amount"), ("Line Total", "Amount"), ("Ext $", "Amount"), ("Total", "Amount"),
]
MANUAL_COLS = ["File", "Supplier", "Type", "Invoice No", "Invoice Date", "Due Date", "PO", "Property",
               "Subtotal", "Tax", "Total", "Entered by", "Note"]
N_MANUAL_PREFILL = 2  # the shipped workbook has 2 of the 6 scanned invoices typed in, to show how it works

# ---------------------------------------------------------------- Power Query (M)
Q = {}
Q["Setting"] = """(name as text) as any =>
    // one value from a named cell on the Settings sheet (InvoiceFolder, ReportDate, PeriodStart, PeriodEnd)
    Excel.CurrentWorkbook(){[Name = name]}[Content]{0}[Column1]"""

Q["NormLabel"] = """(t as any) as text =>
    // "Sales Tax (7.25%)", "Tax 7.25%", "Document no.:" -> "sales tax", "tax", "document no"
    let
        s = Text.Lower(Text.Trim(if t = null then "" else Text.From(t))),
        a = if Text.Contains(s, "(") then Text.BeforeDelimiter(s, "(") else s,
        b = Text.Remove(a, {"0".."9", "%", ".", ":"}),
        c = Text.Combine(List.Select(Text.Split(b, " "), each _ <> ""), " ")
    in
        c"""

Q["Num"] = """(t as any) as nullable number =>
    // "1,234.50", "1234.50", "$1,234.50", "($1,234.50)", "$-6.30" -> number; anything else -> null
    let
        s  = Text.Remove(if t = null then "" else Text.From(t), {"$", ",", " "}),
        ng = Text.StartsWith(s, "(") and Text.EndsWith(s, ")"),
        n  = try Number.FromText(Text.Remove(s, {"(", ")"}), "en-US") otherwise null
    in
        if s = "" or n = null then null else if ng then -n else n"""

Q["FieldMap"] = """let
    // Label on invoice -> field. Add a row when a new supplier calls a field something else.
    Src  = Excel.CurrentWorkbook(){[Name = "FieldLabels"]}[Content],
    Real = Table.SelectRows(Src, each [Label on invoice] <> null and [Field] <> null),
    Key  = Table.AddColumn(Real, "Key", each NormLabel([Label on invoice]), type text),
    Raw  = Table.AddColumn(Key, "Raw", each Text.Lower(Text.Trim(Text.From([Label on invoice]))), type text)
in
    Table.Buffer(Table.Distinct(Raw, {"Key"}))"""

Q["ColumnMap"] = """let
    // Column header on invoice -> column of the Lines table
    Src  = Excel.CurrentWorkbook(){[Name = "ColumnLabels"]}[Content],
    Real = Table.SelectRows(Src, each [Column header on invoice] <> null and [Column] <> null),
    Key  = Table.AddColumn(Real, "Key", each NormLabel([Column header on invoice]), type text)
in
    Table.Buffer(Table.Distinct(Key, {"Key"}))"""

Q["SupplierList"] = """let
    Src  = Excel.CurrentWorkbook(){[Name = "Suppliers"]}[Content],
    Real = Table.SelectRows(Src, each [Supplier] <> null and [Text on invoice] <> null)
in
    Table.Buffer(Real)"""

Q["ReadPdf"] = """(content as binary) as record =>
    // everything one PDF says: status, header fields, totals and the line table
    let
        Opened  = try Table.Buffer(Pdf.Tables(content, [Implementation = "1.3"])),
        Fail    = if Opened[HasError] then
                      let m = Opened[Error][Message] in
                      if Text.Contains(m, "Password") then "Password protected"
                      else if Text.Contains(m, "Corrupt") then "Damaged file"
                      else "Could not open"
                  else null,
        Tbls    = if Fail = null then Opened[Value] else #table({"Id", "Name", "Kind", "Data"}, {}),
        ToRows  = (t as table) as list => List.Transform(Table.ToRows(t),
                      each List.Transform(_, each if _ = null then "" else Text.Trim(Text.From(_)))),
        WithRows = Table.Buffer(Table.AddColumn(Tbls, "Rows", each ToRows([Data]))),
        Pages   = Table.SelectRows(WithRows, each [Kind] = "Page"),
        Tables  = Table.SelectRows(WithRows, each [Kind] = "Table"),
        PageCount = Table.RowCount(Pages),
        RowText = (rows as list) as text => Text.Combine(List.Transform(rows, each Text.Combine(List.Select(_, each _ <> ""), " ")), " "),
        FirstText = Text.Upper(if PageCount = 0 then "" else RowText(Pages{0}[Rows])),
        AllText = RowText(List.Combine(Pages[Rows])),
        NoText  = Fail = null and Text.Length(Text.Trim(AllText)) = 0,

        // label/value candidates: header row over a value row, cells side by side, "Label: value"
        TwoRow  = List.Combine(List.Transform(List.Select(Tables[Rows], each List.Count(_) = 2), each List.Zip({_{0}, _{1}}))),
        Side    = List.Combine(List.Transform(List.Combine(Tables[Rows] & Pages[Rows]), (r) =>
                      let c = List.Select(r, each _ <> "") in List.Transform({0 .. List.Count(c) - 2}, (i) => {c{i}, c{i + 1}}))),
        Colon   = List.Transform(List.Select(List.Combine(List.Combine(Pages[Rows])), each Text.Contains(_, ": ")),
                      each {Text.BeforeDelimiter(_, ": "), Text.AfterDelimiter(_, ": ")}),
        Keys    = FieldMap[Key],
        Raws    = FieldMap[Raw],
        Hits    = List.Select(List.Transform(TwoRow & Side & Colon, each [Key = NormLabel(_{0}), Value = Text.Trim(_{1} ?? "")]),
                      each List.Contains(Keys, [Key]) and [Value] <> "" and not List.Contains(Raws, Text.Lower([Value]))),
        // first value for the field that makes sense for it ("Net 30" is a term, not the Net amount)
        GetIf   = (field as text, ok as function) as nullable text =>
                      let h = List.Select(Hits, each FieldMap{[Key = [Key]]}[Field] = field and ok([Value]))
                      in if h = {} then null else h{0}[Value],
        Get     = (field as text) as nullable text => GetIf(field, each true),
        GetNum  = (field as text) as nullable number => Num(GetIf(field, each Num(_) <> null)),

        Sup     = let m = Table.SelectRows(SupplierList, each Text.Contains(FirstText, Text.Upper([Text on invoice])))
                  in if Table.RowCount(m) = 0 then null else m{0},
        AsDate  = (t as nullable text) as nullable date =>
                      if t = null or Sup = null then null
                      else try Date.FromText(t, [Format = Sup[Date format], Culture = "en-US"]) otherwise null,
        Blank   = (t as nullable text) as nullable text =>
                      if t = null or List.Contains({"-", "n/a", "none", ""}, Text.Lower(t)) then null else t,
        InvDate = AsDate(GetIf("Invoice Date", each AsDate(_) <> null)),
        TermsDays = let t = Get("Terms") in
                      if t <> null and Text.StartsWith(Text.Lower(t), "net ") then Number.FromText(Text.AfterDelimiter(t, " ")) else null,
        Due0    = AsDate(GetIf("Due Date", each AsDate(_) <> null)),
        Due     = if Due0 <> null then Due0
                  else if InvDate = null then null
                  else Date.AddDays(InvDate, if TermsDays <> null then TermsDays else if Sup <> null then Sup[Terms days] else 30),

        // line tables: a table whose first row has a Description column and an Amount column
        ColOf   = (h as text) as nullable text => let k = NormLabel(h), m = Table.SelectRows(ColumnMap, each [Key] = k)
                      in if Table.RowCount(m) = 0 then null else m{0}[Column],
        // a line table starts at the row whose headers include a Description and an Amount column; the rows above
        // it (when the PDF joined the header block and the lines into one table) are skipped
        FromHeader = (rows as list) as list =>
                      let p = List.PositionOf(List.Transform(rows, (r) => List.ContainsAll(List.Transform(r, ColOf), {"Description", "Amount"})), true)
                      in if p < 0 then {} else List.Skip(rows, p),
        InTables = List.Select(List.Transform(Tables[Rows], FromHeader), each List.Count(_) > 1),
        // no table found (some PDFs draw the lines without borders): read the same rows from the page text instead
        ItemTbls = if InTables <> {} then InTables
                   else List.Select(List.Transform(Pages[Rows], FromHeader), each List.Count(_) > 1),
        LinesOf = (rows as list) as list =>
                      let
                          names = List.Transform(rows{0}, ColOf),
                          pos   = (c as text) => List.PositionOf(names, c),
                          cell  = (r as list, c as text) => let p = pos(c) in if p < 0 then null else r{p},
                          body  = List.Select(List.Skip(rows, 1), (r) => r <> rows{0} and cell(r, "Description") <> ""
                                      and Num(cell(r, "Amount")) <> null)
                      in
                          List.Transform(body, (r) => [SKU = Blank(cell(r, "SKU")), Description = cell(r, "Description"),
                              Qty = Num(cell(r, "Qty")), #"Unit price" = Num(cell(r, "Unit price")), Amount = Num(cell(r, "Amount"))]),
        Lines   = List.Combine(List.Transform(ItemTbls, LinesOf))
    in
        [Fail = Fail, NoText = NoText, Pages = PageCount,
         Supplier = if Sup = null then null else Sup[Supplier],
         Account = if Sup = null then null else Sup[Account code],
         Type = if Text.Contains(FirstText, "CREDIT NOTE") then "Credit note" else "Invoice",
         InvoiceNo = Get("Invoice No"), InvoiceDate = InvDate, DueDate = Due,
         PO = Blank(Get("PO")), Property = Get("Property"),
         Subtotal = GetNum("Subtotal"), Tax = GetNum("Tax"), Total = GetNum("Total"),
         Lines = Lines]"""

Q["Parsed"] = """let
    // every PDF in the folder, read once per refresh, with its status
    Folder  = Text.From(Setting("InvoiceFolder")),
    Src     = Folder.Files(Folder),
    Pdfs    = Table.SelectRows(Src, each Text.Lower([Extension]) = ".pdf" and not Text.StartsWith([Name], "~")),
    Read    = Table.AddColumn(Pdfs, "R", each ReadPdf([Content])),
    Slim    = Table.SelectColumns(Read, {"Name", "Date modified", "R"}),
    Exp     = Table.ExpandRecordColumn(Slim, "R",
                {"Fail", "NoText", "Pages", "Supplier", "Account", "Type", "InvoiceNo", "InvoiceDate", "DueDate", "PO",
                 "Property", "Subtotal", "Tax", "Total", "Lines"}),
    Buf     = Table.Buffer(Exp),
    Base    = Table.AddColumn(Buf, "Base", each
                if [Fail] <> null then [Fail]
                else if [NoText] then "Scanned image (no text)"
                else if [Supplier] = null then "Unknown supplier"
                else if [InvoiceNo] = null or [InvoiceDate] = null or [Total] = null then "Missing fields"
                else "Read", type text),
    // the same invoice sent twice: the copy received first is kept
    Readable = Table.SelectRows(Base, each [Base] = "Read"),
    Keyed   = Table.AddColumn(Readable, "Key", each [Supplier] & "|" & [InvoiceNo]),
    Groups  = Table.Group(Keyed, {"Key"}, {{"Files", each Table.Sort(_, {{"Date modified", Order.Ascending}, {"Name", Order.Ascending}})[Name]}}),
    FirstOf = Table.Buffer(Table.ExpandListColumn(Table.AddColumn(Groups, "First", each [Files]{0}), "Files")),
    WithDup = Table.AddColumn(Base, "Duplicate of", each
                let m = Table.SelectRows(FirstOf, (g) => g[Files] = [Name]) in
                if Table.RowCount(m) = 0 or m{0}[First] = [Name] then null else m{0}[First], type nullable text),
    Sums    = Table.AddColumn(WithDup, "Lines total", each
                if [Lines] = null or List.Count([Lines]) = 0 then null
                else Number.Round(List.Sum(List.Transform([Lines], each [Amount])), 2), type nullable number),
    Status  = Table.AddColumn(Sums, "Status", each
                if [Base] <> "Read" then [Base]
                else if [Duplicate of] <> null then "Duplicate"
                else if [Lines total] = null then "No line table found"
                else if [Subtotal] <> null and Number.Abs([Lines total] - [Subtotal]) > 0.005 then "Lines don't add up"
                else if Number.Abs(([Subtotal] ?? [Lines total]) + ([Tax] ?? 0) - [Total]) > 0.005 then "Total doesn't add up"
                else "OK", type text)
in
    Status"""

Q["Register"] = """let
    // one row per PDF: what was read, whether it can be paid, and what to do when it cannot
    Src     = Parsed,
    Manual0 = try Excel.CurrentWorkbook(){[Name = "ManualEntry"]}[Content] otherwise null,
    Manual  = if Manual0 = null then #table({"File"}, {}) else Table.SelectRows(Manual0, each [File] <> null and [Total] <> null),
    MKeys   = List.Transform(Manual[File], each Text.Lower(Text.Trim(Text.From(_)))),
    M       = (f as text) => let p = List.PositionOf(MKeys, Text.Lower(f)) in if p < 0 then null else Manual{p},
    Rows    = Table.AddColumn(Src, "Row", each
                let
                    m      = if [Status] = "OK" or [Status] = "Duplicate" then null else M([Name]),
                    d      = (v) => if v = null then null else if v is date then v else if v is datetime then Date.From(v)
                                    else try Date.From(v) otherwise null,
                    n      = (v) => if v = null then null else Number.From(v),
                    status = if m <> null then "Entered by hand" else [Status],
                    group  = if status = "OK" then "Read automatically"
                             else if status = "Entered by hand" then "Entered by hand"
                             else if status = "Duplicate" then "Duplicate (not counted)"
                             else if status = "Lines don't add up" or status = "Total doesn't add up" then "Held: totals don't add up"
                             else if status = "Scanned image (no text)" then "Scanned image"
                             else if status = "Password protected" or status = "Damaged file" or status = "Could not open" then "Password / damaged"
                             else "Unknown supplier / layout",
                    payable = if group = "Read automatically" or group = "Entered by hand" then "Yes"
                              else if group = "Held: totals don't add up" then "Hold" else "No",
                    nextstep =
                        if status = "Duplicate" then "Already counted from " & [Duplicate of] & ". Nothing to pay twice."
                        else if status = "Scanned image (no text)" then "Type it once on Manual entry (or ask the supplier for the original PDF)."
                        else if status = "Password protected" then "Ask the supplier for a copy without a password, or type it on Manual entry."
                        else if status = "Damaged file" or status = "Could not open" then "The file is cut off. Ask the supplier to send it again."
                        else if status = "Lines don't add up" then "The supplier's lines come to " & Number.ToText([Lines total], "N2", "en-US")
                             & ", the invoice says " & Number.ToText([Subtotal], "N2", "en-US") & ". Ask for a corrected invoice before paying."
                        else if status = "Total doesn't add up" then "Subtotal + tax is not the total printed. Ask the supplier."
                        else if status = "Unknown supplier" then "New supplier: add it on the Suppliers sheet (name and date format), then Refresh All."
                        else if status = "Missing fields" then "A field has a label we do not know: add it on Labels, then Refresh All."
                        else if status = "No line table found" then "The line table has headers we do not know: add them on Labels."
                        else ""
                in
                    if m <> null then
                        [Group = group, Status = status, Payable = payable, #"Next step" = "",
                         Supplier = m[Supplier], Type = m[Type] ?? "Invoice", #"Invoice No" = Text.From(m[#"Invoice No"]),
                         #"Invoice Date" = d(m[#"Invoice Date"]), #"Due Date" = d(m[#"Due Date"]),
                         PO = m[PO], Property = m[Property], Subtotal = n(m[Subtotal]), Tax = n(m[Tax]), Total = n(m[Total]),
                         #"Lines total" = null, Lines = null,
                         #"Account code" = try SupplierList{[Supplier = m[Supplier]]}[Account code] otherwise null]
                    else
                        [Group = group, Status = status, Payable = payable, #"Next step" = nextstep,
                         Supplier = [Supplier], Type = if [Supplier] = null then null else [Type], #"Invoice No" = [InvoiceNo],
                         #"Invoice Date" = [InvoiceDate], #"Due Date" = [DueDate], PO = [PO], Property = [Property],
                         Subtotal = [Subtotal], Tax = [Tax], Total = [Total], #"Lines total" = [Lines total],
                         Lines = if [Lines] = null or List.Count([Lines]) = 0 then null else List.Count([Lines]),
                         #"Account code" = [Account]]),
    Keep    = Table.SelectColumns(Rows, {"Name", "Date modified", "Pages", "Duplicate of", "Row"}),
    Exp     = Table.ExpandRecordColumn(Keep, "Row", {"Group", "Status", "Payable", "Next step", "Supplier", "Type",
                "Invoice No", "Invoice Date", "Due Date", "PO", "Property", "Subtotal", "Tax", "Total", "Lines total",
                "Lines", "Account code"}),
    Ren     = Table.RenameColumns(Exp, {{"Name", "File"}, {"Date modified", "Received"}}),
    Order   = Table.ReorderColumns(Ren, {"File", "Received", "Pages", "Group", "Status", "Payable", "Next step",
                "Supplier", "Type", "Invoice No", "Invoice Date", "Due Date", "PO", "Property", "Subtotal", "Tax",
                "Total", "Lines total", "Lines", "Account code", "Duplicate of"}),
    Typed   = Table.TransformColumnTypes(Order, {{"File", type text}, {"Received", type datetime}, {"Pages", Int64.Type},
                {"Group", type text}, {"Status", type text}, {"Payable", type text}, {"Next step", type text},
                {"Supplier", type text}, {"Type", type text}, {"Invoice No", type text}, {"Invoice Date", type date},
                {"Due Date", type date}, {"PO", type text}, {"Property", type text}, {"Subtotal", Currency.Type},
                {"Tax", Currency.Type}, {"Total", Currency.Type}, {"Lines total", Currency.Type}, {"Lines", Int64.Type},
                {"Account code", type text}, {"Duplicate of", type text}}),
    Sorted  = Table.Sort(Typed, {{"Supplier", Order.Ascending}, {"Invoice Date", Order.Ascending}, {"File", Order.Ascending}})
in
    Sorted"""

Q["Lines"] = """let
    // one row per invoice line of every invoice that is paid or held (duplicates and unreadable files are left out)
    Src     = Table.SelectRows(Parsed, each [Status] = "OK" or [Status] = "Lines don't add up" or [Status] = "Total doesn't add up"),
    Cols    = Table.SelectColumns(Src, {"Name", "Status", "Supplier", "Account", "Type", "InvoiceNo", "InvoiceDate",
                "DueDate", "Property", "Lines"}),
    Numbered = Table.AddColumn(Cols, "L", each List.Transform(List.Positions([Lines]), (i) => [Line = i + 1] & [Lines]{i})),
    NoList  = Table.RemoveColumns(Numbered, {"Lines"}),
    One     = Table.ExpandListColumn(NoList, "L"),
    Exp     = Table.ExpandRecordColumn(One, "L", {"Line", "SKU", "Description", "Qty", "Unit price", "Amount"}),
    Pay     = Table.AddColumn(Exp, "Payable", each if [Status] = "OK" then "Yes" else "Hold", type text),
    Ren     = Table.RenameColumns(Pay, {{"Name", "File"}, {"Account", "Account code"}, {"InvoiceNo", "Invoice No"},
                {"InvoiceDate", "Invoice Date"}, {"DueDate", "Due Date"}}),
    Order   = Table.SelectColumns(Ren, {"Supplier", "Invoice No", "Type", "Invoice Date", "Due Date", "Property", "Line",
                "SKU", "Description", "Qty", "Unit price", "Amount", "Account code", "Payable", "File"}),
    Typed   = Table.TransformColumnTypes(Order, {{"Supplier", type text}, {"Invoice No", type text}, {"Type", type text},
                {"Invoice Date", type date}, {"Due Date", type date}, {"Property", type text}, {"Line", Int64.Type},
                {"SKU", type text}, {"Description", type text}, {"Qty", type number}, {"Unit price", Currency.Type},
                {"Amount", Currency.Type}, {"Account code", type text}, {"Payable", type text}, {"File", type text}})
in
    Table.Sort(Typed, {{"Supplier", Order.Ascending}, {"Invoice Date", Order.Ascending}, {"Invoice No", Order.Ascending}, {"Line", Order.Ascending}})"""

LOADED = [("Register", "Register"), ("Lines", "Lines")]
LAYOUT_SHEETS = ["Needs attention", "Accounting import", "Checks", "Calc", "Dashboard", "How to use"]


# ---------------------------------------------------------------- look
def hx(h):
    h = h.lstrip("#")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


BG, PAGE, CARD, CARD2, LINE = map(hx, ("0B0F13", "12171D", "1A2129", "202832", "2A333D"))
GOLD, GOLD_LT, IVORY, MUTED = map(hx, ("C9A35B", "E6CF9A", "ECE6DA", "8E98A3"))
JADE, COPPER, SKY, GREY = map(hx, ("4FA38C", "B5754A", "7FA7C9", "4A5560"))
GOOD, BAD, AMBER = map(hx, ("5BBF8A", "E5735F", "D9A441"))
T_GOOD, T_BAD, T_AMBER = map(hx, ("1E3A2E", "45231E", "3D3220"))
INPUT = hx("FFF5CC")
HEAD_FILL, HEAD_INK, ROW_LINE = hx("1A2129"), hx("E6CF9A"), hx("D9D9D9")
SERIF, SANS = "Georgia", "Segoe UI"

SHEETS = ["Dashboard", "Needs attention", "Register", "Lines", "Accounting import", "Manual entry", "Checks",
          "Suppliers", "Labels", "Settings", "How to use", "Calc"]


def conn_string(name):
    return f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=""'


def load_query(wb, ws, name, table_name, at="A4"):
    lo = ws.ListObjects.Add(0, conn_string(name), None, 1, ws.Range(at))
    qt = lo.QueryTable
    qt.CommandType = 2
    qt.CommandText = f"SELECT * FROM [{name}]"
    qt.RefreshStyle = 1
    qt.AdjustColumnWidth = False
    qt.BackgroundQuery = False
    qt.Refresh(False)
    lo.Name = table_name
    lo.TableStyle = "TableStyleLight1"
    return lo


def fmt(rng, fill=None, color=None, size=None, bold=None, name=None, italic=None, h=None, v=None, wrap=None, numfmt=None):
    if fill is not None:
        rng.Interior.Color = fill
    f = rng.Font
    if color is not None:
        f.Color = color
    if size is not None:
        f.Size = size
    if bold is not None:
        f.Bold = bold
    if name is not None:
        f.Name = name
    if italic is not None:
        f.Italic = italic
    if h is not None:
        rng.HorizontalAlignment = {"l": -4131, "c": -4108, "r": -4152}[h]
    if v is not None:
        rng.VerticalAlignment = {"t": -4160, "c": -4108, "b": -4107}[v]
    if wrap is not None:
        rng.WrapText = wrap
    if numfmt is not None:
        rng.NumberFormat = numfmt


def sheet_title(ws, title, sub, width_cols=12):
    ws.Range("A1").Value = title
    fmt(ws.Range("A1"), size=16, bold=True, name=SERIF, color=hx("1A2129"))
    ws.Range("A2").Value = sub
    fmt(ws.Range("A2"), size=10, italic=True, color=hx("5A6570"))


def style_table(lo, widths=None, money=(), dates=(), wrap=()):
    hdr = lo.HeaderRowRange
    fmt(hdr, fill=HEAD_FILL, color=HEAD_INK, bold=True, v="c")
    hdr.RowHeight = 22
    ws = lo.Range.Worksheet
    for i in range(1, lo.ListColumns.Count + 1):
        name = lo.ListColumns(i).Name
        col = lo.ListColumns(i).Range
        if widths and name in widths:
            col.ColumnWidth = widths[name]
        body = lo.ListColumns(i).DataBodyRange
        if body is None:
            continue
        if name in money:
            body.NumberFormat = '#,##0.00;[Red](#,##0.00);"-"'
        if name in dates:
            body.NumberFormat = "dd mmm yyyy"
        if name in wrap:
            body.WrapText = True
    b = lo.Range.Borders(12)  # inside horizontal
    b.LineStyle = 1
    b.Color = ROW_LINE


# ---------------------------------------------------------------- input sheets
def put_table(ws, top_left_row, headers, rows, name, yellow_cols=None, widths=None):
    for j, h in enumerate(headers, 1):
        ws.Cells(top_left_row, j).Value = h
    for i, r in enumerate(rows, top_left_row + 1):
        for j, v in enumerate(r, 1):
            c = ws.Cells(i, j)
            if isinstance(v, str) and not v.startswith("="):
                c.NumberFormat = "@"
            c.Value = v
    last = top_left_row + max(len(rows), 1)
    rng = ws.Range(ws.Cells(top_left_row, 1), ws.Cells(last, len(headers)))
    lo = ws.ListObjects.Add(1, rng, None, 1)
    lo.Name = name
    lo.TableStyle = "TableStyleLight1"
    fmt(lo.HeaderRowRange, fill=HEAD_FILL, color=HEAD_INK, bold=True)
    if lo.DataBodyRange is not None:
        cols = yellow_cols or range(1, len(headers) + 1)
        for j in cols:
            lo.ListColumns(j).DataBodyRange.Interior.Color = INPUT
    if widths:
        for j, w in enumerate(widths, 1):
            ws.Columns(j).ColumnWidth = w
    return lo


def build_settings(wb):
    ws = wb.Worksheets("Settings")
    sheet_title(ws, "Settings", "Yellow cells can be changed. The folder defaults to the 'invoices' folder next to this workbook.")
    rows = [("Invoice folder", '=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"invoices\\"', "InvoiceFolder",
             "Every PDF in this folder (and its sub-folders) is read. Type a path to use another folder."),
            ("Report date", REPORT_DATE, "ReportDate", "Overdue and 'due in 14 days' are counted from this date"),
            ("Period start", PERIOD[0], "PeriodStart", "Invoices dated outside the period are flagged on Checks"),
            ("Period end", PERIOD[1], "PeriodEnd", ""),
            ("Company", "Cedar Lane Property Management", "Company", "Shown on the dashboard")]
    ws.Range("A4:C4").Value = ("Setting", "Value", "Note")
    fmt(ws.Range("A4:C4"), fill=HEAD_FILL, color=HEAD_INK, bold=True)
    for r, (label, val, nm, note) in enumerate(rows, 5):
        ws.Cells(r, 1).Value = label
        c = ws.Cells(r, 2)
        if isinstance(val, str) and val.startswith("="):
            c.Formula = val
        elif isinstance(val, dt.date):
            c.Formula = f"=DATE({val.year},{val.month},{val.day})"
            c.Value = c.Value
            c.NumberFormat = "dd mmm yyyy"
            c.HorizontalAlignment = -4131
        else:
            c.Value = val
        c.Interior.Color = INPUT
        ws.Cells(r, 3).Value = note
        wb.Names.Add(nm, f"=Settings!$B${r}")
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 70
    ws.Columns("C").ColumnWidth = 80


def build_suppliers(wb):
    ws = wb.Worksheets("Suppliers")
    sheet_title(ws, "Suppliers", "How each supplier is recognised and how it writes dates. "
                "A new supplier = one new row, then Refresh All.")
    put_table(ws, 4, ["Supplier", "Text on invoice", "Date format", "Account code", "Account name", "Terms days"],
              [list(s) for s in SUPPLIERS], "Suppliers", widths=[32, 26, 14, 14, 30, 12])
    ws.Range("A14").Value = "Date format: dd/MM/yyyy = 09/10/2026 is 9 October · M/d/yyyy = 10/9/2026 is 9 October · " \
                            "d MMM yyyy = 9 Oct 2026 · yyyy-MM-dd = 2026-10-09."
    ws.Range("A15").Value = "Text on invoice: words printed on page 1 of that supplier's invoices (any case). " \
                            "Terms days: used when the invoice has no due date and no 'Net 30' term."
    fmt(ws.Range("A14:A15"), italic=True, color=hx("5A6570"))


def build_labels(wb):
    ws = wb.Worksheets("Labels")
    sheet_title(ws, "Labels", "Suppliers name the same field differently. Each row teaches the workbook one more name. "
                "Case, numbers, % and brackets are ignored: 'Sales Tax (7.25%)' = 'Sales Tax'.")
    put_table(ws, 4, ["Label on invoice", "Field"], FIELD_LABELS, "FieldLabels", widths=[24, 16])
    for j, w in ((4, 28), (5, 16)):
        ws.Columns(j).ColumnWidth = w
    # second table next to the first (columns D:E)
    for j, h in enumerate(["Column header on invoice", "Column"], 4):
        ws.Cells(4, j).Value = h
    for i, (a, b) in enumerate(COLUMN_LABELS, 5):
        ws.Cells(i, 4).Value = a
        ws.Cells(i, 5).Value = b
    lo = ws.ListObjects.Add(1, ws.Range(ws.Cells(4, 4), ws.Cells(4 + len(COLUMN_LABELS), 5)), None, 1)
    lo.Name = "ColumnLabels"
    lo.TableStyle = "TableStyleLight1"
    fmt(lo.HeaderRowRange, fill=HEAD_FILL, color=HEAD_INK, bold=True)
    lo.DataBodyRange.Interior.Color = INPUT
    ws.Columns(3).ColumnWidth = 4


def manual_prefill():
    exp = json.loads((ROOT / "test" / "expected.json").read_text(encoding="utf-8"))
    inv = {i["number"]: i for i in exp["invoices"]}
    scans = sorted([f for f in exp["files"] if f["status"] == "Scanned image"], key=lambda f: f["file"])
    rows = []
    for f in scans[:N_MANUAL_PREFILL]:
        i = inv[f["number"]]
        d, due = dt.date.fromisoformat(i["date"]), dt.date.fromisoformat(i["due"])
        rows.append([f["file"], i["supplier"], i["kind"], i["number"], d, due, i["po"] or None, i["property"],
                     float(i["subtotal"]), float(i["tax"]), float(i["total"]), "A. Clerk", "Typed from the scan, 03/11"])
    return rows


def build_manual(wb):
    ws = wb.Worksheets("Manual entry")
    sheet_title(ws, "Manual entry", "For PDFs the workbook cannot read (scanned, password, damaged). Type the invoice once: "
                "it is counted from the next Refresh All, and it stays here.")
    rows = manual_prefill()
    for j, h in enumerate(MANUAL_COLS, 1):
        ws.Cells(4, j).Value = h
    for i, r in enumerate(rows, 5):
        for j, v in enumerate(r, 1):
            c = ws.Cells(i, j)
            if isinstance(v, dt.date):
                c.Formula = f"=DATE({v.year},{v.month},{v.day})"
                c.Value = c.Value
            else:
                if isinstance(v, str):
                    c.NumberFormat = "@"
                c.Value = v
    last = 4 + max(len(rows), 1) + 6  # a few empty rows ready to type in
    lo = ws.ListObjects.Add(1, ws.Range(ws.Cells(4, 1), ws.Cells(last, len(MANUAL_COLS))), None, 1)
    lo.Name = "ManualEntry"
    lo.TableStyle = "TableStyleLight1"
    fmt(lo.HeaderRowRange, fill=HEAD_FILL, color=HEAD_INK, bold=True)
    lo.DataBodyRange.Interior.Color = INPUT
    for nm in ("Invoice Date", "Due Date"):
        lo.ListColumns(nm).DataBodyRange.NumberFormat = "dd mmm yyyy"
    for nm in ("Subtotal", "Tax", "Total"):
        lo.ListColumns(nm).DataBodyRange.NumberFormat = "#,##0.00"
    for nm in ("File", "Invoice No", "PO"):
        lo.ListColumns(nm).DataBodyRange.NumberFormat = "@"
    widths = [30, 30, 12, 13, 13, 13, 14, 26, 11, 9, 11, 12, 30]
    for j, w in enumerate(widths, 1):
        ws.Columns(j).ColumnWidth = w
    # drop-downs: file = a PDF that still needs a person, supplier from the Suppliers table, type
    lo.ListColumns("Supplier").DataBodyRange.Validation.Add(3, 1, 1, "=SupplierNames")
    lo.ListColumns("Type").DataBodyRange.Validation.Add(3, 1, 1, "Invoice,Credit note")
    ws.Range("A16").Value = ("File = the PDF's name exactly as on Needs attention (copy it from there). "
                             "Supplier from the list. Subtotal + Tax must equal Total.")
    fmt(ws.Range("A16"), italic=True, color=hx("5A6570"))


# ---------------------------------------------------------------- output sheets
def format_outputs(wb):
    reg = wb.Worksheets("Register").ListObjects("Register")
    style_table(reg, widths={"File": 34, "Received": 15, "Pages": 6, "Group": 22, "Status": 21, "Payable": 8,
                             "Next step": 50, "Supplier": 28, "Type": 11, "Invoice No": 13, "Invoice Date": 12,
                             "Due Date": 12, "PO": 13, "Property": 24, "Subtotal": 11, "Tax": 9, "Total": 11,
                             "Lines total": 11, "Lines": 6, "Account code": 8, "Duplicate of": 30},
                money=("Subtotal", "Tax", "Total", "Lines total"), dates=("Invoice Date", "Due Date"))
    reg.ListColumns("Received").DataBodyRange.NumberFormat = "dd mmm yyyy hh:mm"
    lin = wb.Worksheets("Lines").ListObjects("Lines")
    style_table(lin, widths={"Supplier": 28, "Invoice No": 13, "Type": 11, "Invoice Date": 12, "Due Date": 12,
                             "Property": 24, "Line": 6, "SKU": 11, "Description": 34, "Qty": 6, "Unit price": 10,
                             "Amount": 11, "Account code": 8, "Payable": 8, "File": 34},
                money=("Unit price", "Amount"), dates=("Invoice Date", "Due Date"))
    for ws, t, s in ((wb.Worksheets("Register"), "Register: one row per PDF",
                      "Read by Power Query from every PDF in the invoice folder. Payable = Yes is counted; Hold = the "
                      "supplier's own sums are wrong; No = duplicate or not readable (see Next step)."),
                     (wb.Worksheets("Lines"), "Lines: every invoice line",
                      "Invoices that are paid (Yes) or held (Hold). Duplicates and unreadable files are left out.")):
        sheet_title(ws, t, s)
        ws.Activate()
        ws.Range("A5").Select()
        wb.Application.ActiveWindow.FreezePanes = False
        wb.Application.ActiveWindow.SplitRow = 0
        wb.Application.ActiveWindow.FreezePanes = True


def build_needs(wb):
    """Needs attention: every PDF that is not paid automatically, with the reason, the next step and a link to the file."""
    ws = wb.Worksheets("Needs attention")
    sheet_title(ws, "Needs attention", "Every PDF that a person has to look at: why, what to do, and a link that opens it. "
                "Fixed items drop off at the next Refresh All.")
    hdr = ["#", "Open", "File", "Problem", "Supplier", "Invoice No", "Total", "Next step"]
    ws.Range("A4:H4").Value = hdr
    fmt(ws.Range("A4:H4"), fill=HEAD_FILL, color=HEAD_INK, bold=True, v="c")
    ws.Range("A4:H4").RowHeight = 22
    # the list itself: FILTER over the register, sorted by how urgent the problem is
    ws.Range("J4").Value = "list"
    ws.Range("J5").Formula2 = (
        '=LET(r,Register,keep,(Register[Payable]<>"Yes")*(Register[Group]<>"Duplicate (not counted)"),'
        'rank,MATCH(Register[Group],{"Held: totals don\'t add up","Password / damaged","Scanned image",'
        '"Unknown supplier / layout"},0),'
        'f,FILTER(CHOOSE({1,2,3,4,5,6},Register[File],Register[Status],Register[Supplier]&"",Register[Invoice No]&"",'
        'Register[Total],Register[Next step]),keep,""),'
        'k,FILTER(IFERROR(rank,9),keep,9),IF(ROWS(f)=0,"",SORTBY(f,k,1,INDEX(f,0,1),1)))')
    ws.Columns("J:O").Hidden = True
    n = 40
    for i in range(n):
        r = 5 + i
        ws.Cells(r, 1).Formula = f'=IF(C{r}="","",ROW()-4)'
        ws.Cells(r, 2).Formula = f'=IF(C{r}="","",HYPERLINK(InvoiceFolder&C{r},"Open PDF"))'
        for j, col in enumerate("JKLMNO"):
            ws.Cells(r, 3 + j).Formula = f'=IFERROR(IF(INDEX($J$5#,{i + 1},{j + 1})="","",INDEX($J$5#,{i + 1},{j + 1})),"")'
    ws.Range(f"G5:G{4 + n}").NumberFormat = '#,##0.00;[Red](#,##0.00);""'
    ws.Range(f"B5:B{4 + n}").Font.Color = hx("1F6FB2")
    ws.Range(f"B5:B{4 + n}").Font.Underline = 2
    for col, w in zip("ABCDEFGH", (4, 10, 34, 24, 28, 13, 11, 100)):
        ws.Columns(col).ColumnWidth = w
    ws.Range(f"H5:H{4 + n}").WrapText = True
    b = ws.Range(f"A4:H{4 + n}").Borders(12)
    b.LineStyle = 1
    b.Color = ROW_LINE
    fc = ws.Range(f"D5:D{4 + n}").FormatConditions
    for txt, fill, ink in (("add up", T_AMBER, AMBER), ("Scanned", T_BAD, BAD), ("Password", T_BAD, BAD),
                           ("Damaged", T_BAD, BAD), ("Unknown", T_AMBER, AMBER), ("Missing", T_AMBER, AMBER)):
        c = fc.Add(Type=9, String=txt, TextOperator=0)
        c.Interior.Color = hx("FBE3DE") if ink == BAD else hx("FBF0D9")
        c.Font.Color = hx("9C2F1F") if ink == BAD else hx("8A5A00")
    ws.Range("A47").Formula2 = '=IF(J5<>"","","Nothing needs a person: every PDF is paid, entered by hand or a caught duplicate.")'


def build_import(wb):
    """Accounting import: one row per bill line (tax as its own line), ready to save as CSV for Xero / QuickBooks."""
    ws = wb.Worksheets("Accounting import")
    sheet_title(ws, "Accounting import: bills ready to load",
                "Payable invoices only (Payable = Yes): one row per line, sales tax as its own line, invoices typed by hand as one "
                "line. Save this sheet as CSV and import it as bills (column names can be renamed to your system's).")
    hdr = ["Supplier", "Invoice No", "Reference (PO)", "Invoice Date", "Due Date", "Description", "Quantity",
           "Unit Price", "Amount", "Account Code", "Tracking: Property"]
    for j, h in enumerate(hdr, 1):
        ws.Cells(4, j).Value = h
    fmt(ws.Range(ws.Cells(4, 1), ws.Cells(4, len(hdr))), fill=HEAD_FILL, color=HEAD_INK, bold=True, v="c")
    ws.Range("A4:K4").RowHeight = 22
    ws.Range("A5").Formula2 = (
        '=LET('
        'L,FILTER(Lines,Lines[Payable]="Yes"),'
        'li,CHOOSE({1,2,3,4,5,6,7,8,9,10,11},INDEX(L,0,1),INDEX(L,0,2),'
        'XLOOKUP(INDEX(L,0,1)&"|"&INDEX(L,0,2),Register[Supplier]&"|"&Register[Invoice No],Register[PO]&"","",0),'
        'INDEX(L,0,4),INDEX(L,0,5),INDEX(L,0,9),INDEX(L,0,10),INDEX(L,0,11),INDEX(L,0,12),INDEX(L,0,13),INDEX(L,0,6)),'
        'tk,(Register[Payable]="Yes")*(Register[Tax]<>0)*(Register[Group]="Read automatically"),'
        'T,FILTER(Register,tk),'
        'tx,CHOOSE({1,2,3,4,5,6,7,8,9,10,11},INDEX(T,0,8),INDEX(T,0,10),INDEX(T,0,13),INDEX(T,0,11),INDEX(T,0,12),'
        '"Sales tax",1,INDEX(T,0,16),INDEX(T,0,16),"2210",INDEX(T,0,14)),'
        'mk,Register[Group]="Entered by hand",'
        'M,FILTER(Register,mk,""),'
        'ma,IF(ROWS(M)*(INDEX(M,1,1)<>"")=0,"",CHOOSE({1,2,3,4,5,6,7,8,9,10,11},INDEX(M,0,8),INDEX(M,0,10),INDEX(M,0,13),'
        'INDEX(M,0,11),INDEX(M,0,12),"As per invoice (typed from "&INDEX(M,0,1)&")",1,INDEX(M,0,17),INDEX(M,0,17),'
        'INDEX(M,0,20),INDEX(M,0,14))),'
        'all,IF(INDEX(ma,1,1)="",VSTACK(li,tx),VSTACK(li,tx,ma)),'
        'SORTBY(all,INDEX(all,0,1),1,INDEX(all,0,2),1,INDEX(all,0,6)="Sales tax",1))')
    ws.Range("D5:E3000").NumberFormat = "dd mmm yyyy"
    ws.Range("H5:I3000").NumberFormat = "#,##0.00;[Red]-#,##0.00"
    ws.Range("C5:C3000").NumberFormat = '@'
    for col, w in zip("ABCDEFGHIJK", (28, 13, 14, 12, 12, 40, 9, 11, 11, 9, 26)):
        ws.Columns(col).ColumnWidth = w
    ws.Range("M4").Value = "Total of this sheet"
    ws.Range("M5").Formula2 = "=SUM(INDEX(A5#,0,9))"
    ws.Range("M5").NumberFormat = "#,##0.00"
    fmt(ws.Range("M4"), bold=True)
    ws.Columns("M").ColumnWidth = 18
    ws.Range("M7").Value = "Account 2210 = sales tax paid"
    fmt(ws.Range("M7"), italic=True, color=hx("5A6570"))


CHECKS = [
    ("Every PDF in the folder has one row on Register",
     '=COUNTA(Register[File])', '=ROWS(UNIQUE(Register[File]))', "count"),
    ("Every file is accounted for: paid + held + duplicates + needs a person = files",
     '=COUNTIFS(Register[Payable],"Yes")+COUNTIFS(Register[Payable],"Hold")+COUNTIFS(Register[Payable],"No")',
     '=COUNTA(Register[File])', "count"),
    ("Paid invoices read from PDF: lines add up to the subtotal (to the cent)",
     '=SUMIFS(Register[Lines total],Register[Group],"Read automatically")',
     '=SUMIFS(Register[Subtotal],Register[Group],"Read automatically")', "money"),
    ("Paid invoices: subtotal + tax = total, every invoice",
     '=SUMPRODUCT((Register[Payable]="Yes")*ABS(Register[Subtotal]+Register[Tax]-Register[Total])>0.005)', '0', "count"),
    ("Lines table = lines total on Register (paid + held)",
     '=SUM(Lines[Amount])', '=SUMIFS(Register[Lines total],Register[Payable],"Yes",Register[Group],"Read automatically")'
     '+SUMIFS(Register[Lines total],Register[Payable],"Hold")', "money"),
    ("No invoice counted twice (same supplier + number among paid)",
     '=SUMPRODUCT((Register[Payable]="Yes")*(COUNTIFS(Register[Supplier],Register[Supplier],Register[Invoice No],'
     'Register[Invoice No],Register[Payable],"Yes")>1))', '0', "count"),
    ("Accounting import total = paid total on Register",
     "=SUM(INDEX('Accounting import'!A5#,0,9))", '=SUMIFS(Register[Total],Register[Payable],"Yes")', "money"),
    ("Every paid invoice has supplier, number, date, due date and property",
     '=SUMPRODUCT((Register[Payable]="Yes")*((Register[Supplier]="")+(Register[Invoice No]="")+(Register[Invoice Date]="")'
     '+(Register[Due Date]="")+(Register[Property]="")>0))', '0', "count"),
    ("Invoice dates inside the period (Settings)",
     '=SUMPRODUCT((Register[Payable]<>"No")*(Register[Invoice Date]<>"")*((Register[Invoice Date]<PeriodStart)+(Register[Invoice Date]>PeriodEnd)>0))',
     '0', "count"),
    ("Due date is on or after the invoice date",
     '=SUMPRODUCT((Register[Payable]<>"No")*(Register[Due Date]<>"")*(Register[Due Date]<Register[Invoice Date]))', '0', "count"),
    ("Every row on Manual entry matches a PDF in the folder",
     '=SUMPRODUCT((ManualEntry[File]<>"")*(COUNTIFS(Register[File],ManualEntry[File])=0))', '0', "count"),
    ("Manual entries add up: subtotal + tax = total",
     '=SUMPRODUCT((ManualEntry[File]<>"")*(ABS(N(+ManualEntry[Subtotal])+N(+ManualEntry[Tax])-N(+ManualEntry[Total]))>0.005))',
     '0', "count"),
]


def build_checks(wb):
    ws = wb.Worksheets("Checks")
    sheet_title(ws, "Checks", "Proof that the numbers out of the PDFs are complete and add up. All must say PASS. "
                "A FAIL names where to look.")
    ws.Range("A4:E4").Value = ("#", "Check", "Found", "Expected", "Result")
    fmt(ws.Range("A4:E4"), fill=HEAD_FILL, color=HEAD_INK, bold=True, v="c")
    for i, (label, a, b, kind) in enumerate(CHECKS, 1):
        r = 4 + i
        ws.Cells(r, 1).Value = i
        ws.Cells(r, 2).Value = label
        ws.Cells(r, 3).Formula2 = a
        ws.Cells(r, 4).Formula2 = "=" + b if not b.startswith("=") else b
        ws.Cells(r, 5).Formula = f'=IF(ABS(C{r}-D{r})<0.005,"PASS","FAIL")'
        ws.Range(f"C{r}:D{r}").NumberFormat = "#,##0.00" if kind == "money" else "#,##0"
    last = 4 + len(CHECKS)
    ws.Range(f"B{last + 2}").Value = "Checks passed"
    ws.Range(f"C{last + 2}").Formula = f'=COUNTIF(E5:E{last},"PASS")'
    ws.Range(f"D{last + 2}").Value = len(CHECKS)
    fmt(ws.Range(f"B{last + 2}:D{last + 2}"), bold=True)
    wb.Names.Add("ChecksPassed", f"=Checks!$C${last + 2}")
    wb.Names.Add("ChecksTotal", f"=Checks!$D${last + 2}")
    fc = ws.Range(f"E5:E{last}").FormatConditions
    c = fc.Add(Type=1, Operator=3, Formula1='="PASS"')
    c.Interior.Color = hx("DDF2E6"); c.Font.Color = hx("1E6B43"); c.Font.Bold = True
    c = fc.Add(Type=1, Operator=3, Formula1='="FAIL"')
    c.Interior.Color = hx("FBE3DE"); c.Font.Color = hx("9C2F1F"); c.Font.Bold = True
    for col, w in zip("ABCDE", (4, 78, 14, 14, 9)):
        ws.Columns(col).ColumnWidth = w
    ws.Range(f"E5:E{last}").HorizontalAlignment = -4108


GROUPS = ["Read automatically", "Entered by hand", "Duplicate (not counted)", "Held: totals don't add up",
          "Scanned image", "Password / damaged", "Unknown supplier / layout"]


def build_calc(wb):
    ws = wb.Worksheets("Calc")
    ws.Range("A1").Value = "Helper ranges for the dashboard (hidden)"
    ws.Range("A3:B3").Value = ("Group", "Files")
    for i, g in enumerate(GROUPS, 4):
        ws.Cells(i, 1).Value = g
        ws.Cells(i, 2).Formula = f'=COUNTIFS(Register[Group],A{i})'
    # payable by supplier (sorted, largest first)
    ws.Range("D3:E3").Value = ("Supplier", "Payable")
    ws.Range("D4").Formula2 = ('=LET(s,Suppliers[Supplier],v,SUMIFS(Register[Total],Register[Supplier],s,Register[Payable],"Yes"),'
                               'SORTBY(HSTACK(s,v),v,-1))')
    ws.Range("G3:H3").Value = ("Property", "Payable")
    ws.Range("G4").Formula2 = ('=LET(p,UNIQUE(FILTER(Register[Property],Register[Payable]="Yes")),'
                               'v,SUMIFS(Register[Total],Register[Property],p,Register[Payable],"Yes"),SORTBY(HSTACK(p,v),v,-1))')
    # due profile from the report date
    ws.Range("J3:K3").Value = ("When due", "Payable")
    buckets = [("Overdue", '"<"&ReportDate', None), ("Next 7 days", '">="&ReportDate', '"<"&ReportDate+7'),
               ("8-14 days", '">="&ReportDate+7', '"<"&ReportDate+14'), ("15-30 days", '">="&ReportDate+14', '"<"&ReportDate+30'),
               ("Later", '">="&ReportDate+30', None)]
    for i, (lab, a, b) in enumerate(buckets, 4):
        ws.Cells(i, 10).Value = lab
        cond = f'Register[Due Date],{a}' + (f',Register[Due Date],{b}' if b else "")
        ws.Cells(i, 11).Formula = f'=SUMIFS(Register[Total],Register[Payable],"Yes",{cond})'
    # by month
    ws.Range("M3:N3").Value = ("Month", "Payable")
    for i, m in enumerate((9, 10), 4):
        ws.Cells(i, 13).Value = dt.date(2026, m, 1).strftime("%b %Y")
        ws.Cells(i, 14).Formula = (f'=SUMIFS(Register[Total],Register[Payable],"Yes",Register[Invoice Date],">="&DATE(2026,{m},1),'
                                   f'Register[Invoice Date],"<"&DATE(2026,{m + 1},1))')
    ws.Range("B4:N40").NumberFormat = "#,##0"
    for c in ("E4:E40", "H4:H40", "K4:K40", "N4:N40"):
        ws.Range(c).NumberFormat = "[$$-409]#,##0"


def card(ws, rng, label, formula, sub_formula, numfmt, accent=GOLD):
    r = ws.Range(rng)
    top = r.Rows(1)
    fmt(r, fill=CARD)
    for i in range(1, r.Rows.Count + 1):
        r.Rows(i).Merge()
    r.Cells(1, 1).Value = label
    fmt(r.Rows(1), color=MUTED, size=9, bold=True, h="l", v="b", name=SANS)
    r.Cells(2, 1).Formula2 = formula
    fmt(r.Rows(2), color=accent, size=22, name=SERIF, h="l", v="c", numfmt=numfmt)
    r.Cells(3, 1).Formula2 = sub_formula
    fmt(r.Rows(3), color=IVORY, size=9, h="l", v="t", name=SANS)
    r.IndentLevel = 1
    b = r.Borders(8)  # top edge accent
    b.LineStyle = 1
    b.Weight = 4
    b.Color = accent


def dark_chart(ch, title):
    ch.HasTitle = True
    ch.SetElement(2)  # title above the chart
    ch.ChartTitle.Text = title
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IVORY
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Name = SERIF
    ch.ChartArea.Format.Fill.ForeColor.RGB = CARD
    ch.ChartArea.Format.Line.Visible = False
    ch.PlotArea.Format.Fill.Visible = False
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = MUTED
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = SANS
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Size = 9


def build_dashboard(wb):
    ws = wb.Worksheets("Dashboard")
    ws.Cells.Interior.Color = BG
    ws.Cells.Font.Color = IVORY
    ws.Cells.Font.Name = SANS
    widths = [2] + [11.5] * 16 + [2]
    for j, w in enumerate(widths, 1):
        ws.Columns(j).ColumnWidth = w
    ws.Range("B2").Value = "Supplier invoices · 300 PDFs to one Excel register"
    fmt(ws.Range("B2"), size=22, name=SERIF, color=GOLD_LT)
    ws.Range("B3").Formula = ('=Company&" · accounts payable inbox · "&TEXT(PeriodStart,"d mmm")&" – "&TEXT(PeriodEnd,"d mmm yyyy")'
                              '&" · as at "&TEXT(ReportDate,"d mmm yyyy")&" · checks "&ChecksPassed&"/"&ChecksTotal&" PASS"')
    fmt(ws.Range("B3"), size=10, color=MUTED)
    ws.Rows(2).RowHeight = 34
    ws.Rows(4).RowHeight = 8
    for r, h in ((5, 20), (6, 36), (7, 22)):
        ws.Rows(r).RowHeight = h
    cards = [
        ("B5:C7", "PDF FILES IN THE FOLDER", '=COUNTA(Register[File])',
         '=TEXT(SUM(Register[Pages]),"#,##0")&" pages read"', "#,##0", GOLD),
        ("D5:F7", "READ WITH NO TYPING", '=COUNTIFS(Register[Group],"Read automatically")',
         '=TEXT(COUNTIFS(Register[Group],"Read automatically")/COUNTA(Register[File]),"0%")&" of the files · "&TEXT(SUM(Register[Lines]),"#,##0")&" lines"',
         "#,##0", JADE),
        ("G5:H7", "DUPLICATES CAUGHT", '=COUNTIFS(Register[Group],"Duplicate (not counted)")',
         '="not paid twice: "&TEXT(SUMIFS(Register[Total],Register[Group],"Duplicate (not counted)"),"$#,##0")', "#,##0", SKY),
        ("I5:K7", "NEED A PERSON", "=ROWS('Needs attention'!J5#)-('Needs attention'!J5=\"\")",
         '=COUNTIFS(Register[Group],"Held*")&" held · "&COUNTIFS(Register[Group],"Scanned image")&" scanned · "'
         '&COUNTIFS(Register[Group],"Password / damaged")&" password/damaged"', "#,##0", BAD),
        ("L5:N7", "PAYABLE (CHECKED)", '=SUMIFS(Register[Total],Register[Payable],"Yes")',
         '=TEXT(COUNTIFS(Register[Payable],"Yes"),"#,##0")&" invoices incl. "&COUNTIFS(Register[Type],"Credit note",Register[Payable],"Yes")&" credit notes"',
         '"$"#,##0', GOLD),
        ("O5:Q7", "OVERDUE AT REPORT DATE", '=Calc!K4',
         '="due next 14 days: "&TEXT(Calc!K5+Calc!K6,"$#,##0")', '"$"#,##0', AMBER),
    ]
    for rng, lab, f, sub, nf, acc in cards:
        card(ws, rng, lab, f, sub, nf, acc)
    # insight sentences
    ws.Range("B9").Formula = ('="▸ "&COUNTIFS(Register[Group],"Read automatically")&" of "&COUNTA(Register[File])&'
                              '" PDFs went straight to the register: 4 layouts, 8 suppliers, dates read the way each supplier writes them."')
    ws.Range("B10").Formula = ('="▸ "&COUNTIFS(Register[Group],"Held*")&" invoices are held: the supplier\'s own lines do not add up to its subtotal. '
                               'Ask for a corrected invoice before paying (Needs attention)."')
    ws.Range("B11").Formula = ('="▸ Largest supplier: "&INDEX(Calc!D4#,1,1)&", "&TEXT(INDEX(Calc!D4#,1,2),"$#,##0")&" ("&'
                               'TEXT(INDEX(Calc!D4#,1,2)/SUMIFS(Register[Total],Register[Payable],"Yes"),"0%")&" of the payable total)."')
    fmt(ws.Range("B9:B11"), color=IVORY, size=10)
    # charts
    ch_top = ws.Range("B13").Top
    left = ws.Range("B13").Left
    w_all = ws.Range("B13:Q13").Width
    gap = 8
    w3 = (w_all - 2 * gap) / 3
    calc = wb.Worksheets("Calc")

    def add_chart(kind, x, y, w, h, src, title, color):
        shp = ws.Shapes.AddChart2(-1, kind, x, y, w, h)
        ch = shp.Chart
        ch.SetSourceData(src)
        dark_chart(ch, title)
        ch.HasLegend = False
        s = ch.SeriesCollection(1)
        s.Format.Fill.ForeColor.RGB = color
        s.HasDataLabels = True
        s.DataLabels().Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IVORY
        s.DataLabels().Format.TextFrame2.TextRange.Font.Size = 8
        return ch

    ch1 = add_chart(57, left, ch_top, w3, 250, calc.Range("A4:B10"), "What happened to the 300 files", GOLD)
    ch1.Axes(1).ReversePlotOrder = True
    ch1.Axes(2).Delete()
    s = ch1.SeriesCollection(1)
    pal = [JADE, GOLD, SKY, AMBER, BAD, BAD, GREY]
    for i, c in enumerate(pal, 1):
        s.Points(i).Format.Fill.ForeColor.RGB = c
    ch1.ChartGroups(1).GapWidth = 60
    ch2 = add_chart(57, left + w3 + gap, ch_top, w3, 250, calc.Range("D4:E11"), "Payable by supplier", GOLD)
    ch2.Axes(1).ReversePlotOrder = True
    ch2.Axes(2).Delete()
    ch2.ChartGroups(1).GapWidth = 60
    ch2.SeriesCollection(1).DataLabels().NumberFormat = '[$$-409]#,##0'
    ch3 = add_chart(51, left + 2 * (w3 + gap), ch_top, w3, 250, calc.Range("J4:K8"), "Payable by due date (from report date)", COPPER)
    ch3.Axes(2).Delete()
    ch3.ChartGroups(1).GapWidth = 60
    ch3.SeriesCollection(1).DataLabels().NumberFormat = '[$$-409]#,##0'
    ch3.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = BAD
    for ch in (ch1, ch2, ch3):
        try:
            ch.Axes(2).MajorGridlines.Delete()
        except Exception:
            pass
        ch.Axes(1).Format.Line.ForeColor.RGB = LINE
        ch.Axes(1).TickLabels.Font.Color = IVORY
    # needs a person, top of the list
    top = ws.Range("B31").Row
    ws.Range(f"B{top}").Value = "Needs a person (top 8, full list on Needs attention)"
    fmt(ws.Range(f"B{top}"), color=GOLD_LT, size=12, name=SERIF)
    heads = [("B", "C", "File"), ("D", "E", "Problem"), ("F", "H", "Supplier"), ("I", "I", "Total"), ("J", "Q", "Next step")]
    for a, b, t in heads:
        ws.Range(f"{a}{top + 1}:{b}{top + 1}").Merge()
        ws.Range(f"{a}{top + 1}").Value = t
    fmt(ws.Range(f"B{top + 1}:Q{top + 1}"), fill=CARD2, color=MUTED, bold=True, size=9)
    src_col = {"File": "C", "Problem": "D", "Supplier": "E", "Total": "G", "Next step": "H"}
    for i in range(8):
        r = top + 2 + i
        for a, b, t in heads:
            ws.Range(f"{a}{r}:{b}{r}").Merge()
            ws.Range(f"{a}{r}").Formula = f"='Needs attention'!{src_col[t]}{5 + i}&\"\""
        ws.Range(f"I{r}").Formula = f"=IF('Needs attention'!G{5 + i}=\"\",\"\",'Needs attention'!G{5 + i})"
        ws.Range(f"I{r}").NumberFormat = '"$"#,##0.00;[Red]("$"#,##0.00);'
        fmt(ws.Range(f"B{r}:Q{r}"), fill=CARD if i % 2 == 0 else PAGE, color=IVORY, size=9)
        ws.Range(f"D{r}").Font.Color = AMBER
    ws.Range(f"B{top + 11}").Value = ("Read by Power Query (no macros) · Data > Refresh All after new PDFs land in the folder · "
                                      "all names, companies and amounts are fictional")
    fmt(ws.Range(f"B{top + 11}"), color=MUTED, size=8, italic=True)


def build_howto(wb):
    ws = wb.Worksheets("How to use")
    ws.Columns("A").ColumnWidth = 125
    text = [
        ("How to use this workbook", True), ("", False),
        ("Every week (or every day)", True),
        ("1. Save the supplier PDFs from the accounts payable inbox into the invoices folder (see Settings). Sub-folders are fine.", False),
        ("2. Data > Refresh All. Every PDF is read again; nothing is typed.", False),
        ("3. Checks: every line must say PASS.", False),
        ("4. Needs attention: one row per PDF a person has to look at, with the reason, the next step and an 'Open PDF' link.", False),
        ("   - Scanned or password-protected: type the invoice once on Manual entry (File = the PDF's name). It is counted from the next refresh.", False),
        ("   - Held (lines do not add up): ask the supplier for a corrected invoice; drop the new PDF into the folder.", False),
        ("   - Unknown supplier: add one row on Suppliers (name, words on its page 1, date format, account code).", False),
        ("5. Accounting import: File > Save a Copy > CSV with this sheet active, then import it as bills (Xero, QuickBooks, Sage).", False),
        ("", False),
        ("What the workbook reads from each PDF", True),
        ("Supplier: found from words on page 1 (Suppliers sheet). Dates are read with that supplier's own format: 09/10/2026 is 9 October", False),
        ("for a supplier that writes day first, and 9/10/2026 is 10 September for one that writes month first.", False),
        ("Header fields: invoice number, invoice date, due date (or 'Net 30' terms), PO, property, subtotal, tax, total. Each supplier labels", False),
        ("them differently (Invoice # / Inv. Number / Document no.); the Labels sheet lists every name. Add a row when a new one appears.", False),
        ("Lines: the table with a description and an amount column, over as many pages as the invoice has.", False),
        ("Amounts: 1,234.50 · 1234.50 · $1,234.50 · ($1,234.50) for credit notes, all read as numbers.", False),
        ("", False),
        ("What is checked on every refresh", True),
        ("Lines add up to the subtotal, subtotal + tax = total, the same supplier + invoice number is not counted twice (the copy received", False),
        ("first is kept), dates are in the period, the import sheet totals the paid invoices. See Checks.", False),
        ("", False),
        ("Where the steps are", True),
        ("Data > Queries & Connections: ReadPdf (one PDF), Parsed (all PDFs + status), Register, Lines. No macros, no add-ins.", False),
        ("Needs Excel for Microsoft 365 or Excel 2021 (PDF connector, dynamic arrays).", False),
    ]
    for i, (t, bold) in enumerate(text, 1):
        ws.Cells(i, 1).Value = t
        ws.Cells(i, 1).Font.Bold = bold
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Name = SERIF


def export_csv(wb):
    ws = wb.Worksheets("Accounting import")
    rng = ws.Range("A5").SpillingToRange
    hdr = [ws.Cells(4, j).Value for j in range(1, 12)]
    (ROOT / "output").mkdir(exist_ok=True)
    out = ROOT / "output" / "bills-import-2026-09-10.csv"
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(hdr)
        for r in rng.Value:
            row = []
            for j, v in enumerate(r):
                if hasattr(v, "year"):
                    row.append(v.strftime("%Y-%m-%d"))
                elif isinstance(v, float) and j in (7, 8):
                    row.append(f"{v:.2f}")
                elif isinstance(v, float) and v.is_integer():
                    row.append(str(int(v)))
                else:
                    row.append("" if v is None else v)
            w.writerow(row)
    print("wrote", out.relative_to(ROOT), rng.Rows.Count, "rows")


def finish(xl, wb):
    build_needs(wb)
    build_import(wb)
    build_checks(wb)
    build_calc(wb)
    xl.Calculate()
    build_dashboard(wb)
    build_howto(wb)
    format_outputs(wb)
    xl.Calculate()
    wb.Worksheets("Needs attention").Rows("5:44").AutoFit()
    wb.Worksheets("Calc").Visible = 0
    for ws in wb.Worksheets:
        if ws.Visible:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            xl.ActiveWindow.ScrollRow = 1
            xl.ActiveWindow.ScrollColumn = 1
    tab = {"Dashboard": GOLD, "Needs attention": BAD, "Checks": JADE, "Manual entry": hx("F2C94C")}
    for nm, c in tab.items():
        wb.Worksheets(nm).Tab.Color = c
    wb.Worksheets("Dashboard").Activate()
    wb.Worksheets("Dashboard").Range("A1").Select()
    export_csv(wb)
    wb.Save()
    print("saved", OUT.name)


def main():
    xl = win32.DispatchEx("Excel.Application")  # its own Excel: never touches another open Excel
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        if "--layout" in sys.argv:  # rebuild the formula sheets only, on the saved workbook (queries untouched)
            wb = xl.Workbooks.Open(str(OUT))
            for name in LAYOUT_SHEETS:
                ws = wb.Worksheets(name)
                ws.Visible = -1
                for i in range(ws.Shapes.Count, 0, -1):
                    ws.Shapes(i).Delete()
                ws.Cells.Clear()
                ws.Cells.FormatConditions.Delete()
                ws.Cells.UnMerge()
                ws.Columns.Hidden = False
                ws.Rows.RowHeight = 15
            for nm in ("ChecksPassed", "ChecksTotal"):
                try:
                    wb.Names(nm).Delete()
                except Exception:
                    pass
            if "--refresh" in sys.argv:
                t1 = time.time()
                wb.RefreshAll()
                xl.CalculateUntilAsyncQueriesDone()
                print(f"refreshed in {time.time() - t1:.1f}s", flush=True)
            finish(xl, wb)
            return
        if OUT.exists():
            OUT.unlink()
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < len(SHEETS):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, name in enumerate(SHEETS, 1):
            wb.Worksheets(i).Name = name
        st = wb.Styles("Normal")
        st.Font.Name = SANS
        st.Font.Size = 10
        build_settings(wb)
        build_suppliers(wb)
        wb.Names.Add("SupplierNames", "=Suppliers[Supplier]")
        build_labels(wb)
        build_manual(wb)
        wb.SaveAs(str(OUT), 51)
        try:
            wb.Queries.FastCombine = True
        except Exception as e:
            print("FastCombine not set:", e)
        for name, formula in Q.items():
            wb.Queries.Add(name, formula)
        t0 = time.time()
        for sheet, query in LOADED:
            t1 = time.time()
            load_query(wb, wb.Worksheets(sheet), query, query)
            print(f"  {query} loaded in {time.time() - t1:.1f}s", flush=True)
        print(f"queries loaded in {time.time() - t0:.1f}s", flush=True)
        wb.Save()
        finish(xl, wb)
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        xl.Quit()


if __name__ == "__main__":
    main()
