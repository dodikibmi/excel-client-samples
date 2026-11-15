"""Builds Customer-Cleanup.xlsx with real Excel (COM).

Power Query reads three customer sources (webshop CSV, POS CSV, the sales team's old .xlsx with merged cells),
cleans every field, drops exact duplicates and test rows, links records that share an email, a phone or
first name + last name + ZIP (nicknames from an editable table: Bill = William), keeps the most recent record as
master and fills its blanks from the duplicates. Output: one row per customer, a Merge log, a Needs review tab for
look-alikes, a CRM-ready import tab, and a Summary with data checks.

Run with a Python that has pywin32:  python scripts/build_workbook.py
The workbook has no macros; the client only needs Data -> Refresh All.
"""
import csv, time
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Customer-Cleanup.xlsx"

NICK = [("Jon", "John"), ("Johnny", "John"), ("Bill", "William"), ("Will", "William"), ("Bob", "Robert"), ("Rob", "Robert"),
        ("Liz", "Elizabeth"), ("Beth", "Elizabeth"), ("Kate", "Katherine"), ("Katie", "Katherine"), ("Mike", "Michael"),
        ("Jim", "James"), ("Jimmy", "James"), ("Tom", "Thomas"), ("Dan", "Daniel"), ("Danny", "Daniel"),
        ("Chris", "Christopher"), ("Matt", "Matthew"), ("Nick", "Nicholas"), ("Sam", "Samuel"), ("Alex", "Alexander"),
        ("Tony", "Anthony"), ("Jen", "Jennifer"), ("Jenny", "Jennifer"), ("Meg", "Margaret"), ("Peggy", "Margaret"),
        ("Steve", "Steven"), ("Dave", "David"), ("Andy", "Andrew"), ("Ben", "Benjamin"), ("Joe", "Joseph"),
        ("Abby", "Abigail"), ("Gabe", "Gabriel"), ("Ted", "Edward"), ("Ed", "Edward"), ("Rick", "Richard"),
        ("Sue", "Susan"), ("Becky", "Rebecca"), ("Nate", "Nathan")]

# ---------------------------------------------------------------- Power Query (M)
Q = {}
Q["Setting"] = """(name as text) as text =>
    // Reads one value from a named cell on the Settings sheet (DataFolder, WebshopFile, PosFile, SalesFile)
    Text.From(Excel.CurrentWorkbook(){[Name = name]}[Content]{0}[Column1])"""

# data-quality tests, used for the before/after score (same definitions in test/verify.py)
Q["IsEmail"] = """(s as text) as logical =>
    // usable email exactly as written: lower case, no spaces, one @, something before it and a dot after it
    s <> "" and s = Text.Lower(s) and not Text.Contains(s, " ") and List.Count(Text.Split(s, "@")) = 2
    and Text.BeforeDelimiter(s, "@") <> "" and Text.Contains(Text.AfterDelimiter(s, "@"), ".")"""

Q["IsPhone"] = """(s as text) as logical =>
    // phone exactly in the house format (303) 555-0142
    let d = Text.Select(s, {"0".."9"}) in
    Text.Length(d) = 10 and s = "(" & Text.Start(d, 3) & ") " & Text.Middle(d, 3, 3) & "-" & Text.End(d, 4)"""

Q["IsZip"] = """(s as text) as logical => Text.Length(s) = 5 and Text.Length(Text.Select(s, {"0".."9"})) = 5"""

Q["Decisions"] = """let
    // the Decision column a person fills in on Needs review (Merge / Keep separate). It is read back from the
    // sheet on every refresh, so a decision survives Refresh All and a Merge is applied in Keyed.
    Empty  = #table(type table [#"Master A" = text, #"Master B" = text, Decision = text], {}),
    Src    = try Excel.CurrentWorkbook(){[Name = "Review"]}[Content] otherwise null,
    Cols   = if Src = null or not List.ContainsAll(Table.ColumnNames(Src), {"Master A", "Master B", "Decision"}) then Empty
             else Table.SelectColumns(Src, {"Master A", "Master B", "Decision"}),
    Real   = Table.SelectRows(Cols, each [Master A] <> null and [Master B] <> null),
    Norm   = Table.TransformColumns(Real, {{"Decision", each let u = Text.Upper(Text.Trim(if _ = null then "" else Text.From(_))) in
                 if Text.StartsWith(u, "MERGE") then "Merge" else if Text.StartsWith(u, "KEEP") then "Keep separate" else ""}}),
    Keys   = Table.AddColumn(Norm, "Pair", each Text.From([Master A]) & "|" & Text.From([Master B]), type text)
in
    Table.Buffer(Table.Distinct(Keys, {"Pair"}))"""

Q["Webshop_Raw"] = """let
    Source  = Csv.Document(File.Contents(Setting("DataFolder") & Setting("WebshopFile")),
                  [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Headers"""

Q["POS_Raw"] = """let
    Source  = Csv.Document(File.Contents(Setting("DataFolder") & Setting("PosFile")),
                  [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Headers"""

Q["Sales_Raw"] = """let
    // the sales team's spreadsheet: two title rows above the header, company name in merged cells
    Book    = Excel.Workbook(File.Contents(Setting("DataFolder") & Setting("SalesFile")), null, true),
    Sheet   = Book{[Item = "Trade accounts", Kind = "Sheet"]}[Data],
    Skipped = Table.Skip(Sheet, 2),
    Headers = Table.PromoteHeaders(Skipped, [PromoteAllScalars = true]),
    NoBlank = Table.SelectRows(Headers, each not List.IsEmpty(List.Select(Record.FieldValues(_), each _ <> null and Text.Trim(Text.From(_)) <> "")))
in
    NoBlank"""

RAWKEY = 'Text.Combine(List.Transform(Record.FieldValues(Record.RemoveFields(r, {"Source row"})), each if _ = null then "" else Text.From(_)), "|")'

Q["Webshop_In"] = f"""let
    S = Table.AddIndexColumn(Webshop_Raw, "Source row", 1, 1, Int64.Type),
    R = Table.FromRecords(Table.TransformRows(S, (r) => [
          Record = "W-" & Text.PadStart(Text.From(r[Source row]), 4, "0"), Source = "Webshop", #"Source row" = r[Source row],
          RawKey = {RAWKEY},
          #"Name typed" = r[Full Name], #"First typed" = "", #"Last typed" = "", #"Email raw" = r[Email], #"Phone raw" = r[Phone],
          #"Company raw" = r[Company], #"Address raw" = r[Street], #"City raw" = r[City], #"State raw" = r[State],
          #"ZIP raw" = r[ZIP], #"Opt-in raw" = r[#"Marketing Opt-in"], #"Company filled" = false,
          #"Last activity" = Date.FromText(Text.Trim(r[Last Order]), "en-US"),
          Spend = Number.FromText(Text.Remove(Text.Trim(r[Total Spent]), {{"$", ","}}), "en-US")]))
in
    R"""

Q["POS_In"] = f"""let
    S = Table.AddIndexColumn(POS_Raw, "Source row", 1, 1, Int64.Type),
    R = Table.FromRecords(Table.TransformRows(S, (r) => [
          Record = "P-" & Text.PadStart(Text.From(r[Source row]), 4, "0"), Source = "POS", #"Source row" = r[Source row],
          RawKey = {RAWKEY},
          #"Name typed" = r[Customer], #"First typed" = "", #"Last typed" = "", #"Email raw" = r[Email], #"Phone raw" = r[Phone],
          #"Company raw" = "", #"Address raw" = r[Address], #"City raw" = r[City], #"State raw" = r[St],
          #"ZIP raw" = r[Zip], #"Opt-in raw" = "", #"Company filled" = false,
          #"Last activity" = Date.FromText(Text.Trim(r[Last Visit]), "en-US"),     // 7/5/2025 = July 5
          Spend = if Text.Trim(r[Spend]) = "" then 0 else Number.FromText(Text.Trim(r[Spend]), "en-US")]))
in
    R"""

Q["Sales_In"] = f"""let
    S0 = Table.AddIndexColumn(Sales_Raw, "Source row", 1, 1, Int64.Type),
    // merged cells arrive as one value on the first row and blanks below it: remember which, then fill down
    Keys   = Table.AddColumn(S0, "RawKey", (r) => {RAWKEY}, type text),
    Flag   = Table.AddColumn(Keys, "Company filled", each [Company] = null or Text.Trim(Text.From([Company])) = "", type logical),
    Filled = Table.FillDown(Table.ReplaceValue(Flag, "", null, Replacer.ReplaceValue, {{"Company"}}), {{"Company"}}),
    T      = (x) => if x = null then "" else Text.Trim(Text.From(x)),
    AsDate = (v) => if Value.Is(v, type datetime) then Date.From(v) else if Value.Is(v, type date) then v
                    else if Value.Is(v, type number) then Date.From(v)     // a real date cell arrives as its serial number
                    else Date.FromText(T(v), "en-GB"),   // "3 Oct 2025" typed as text
    R = Table.FromRecords(Table.TransformRows(Filled, (r) => [
          Record = "S-" & Text.PadStart(Text.From(r[Source row]), 4, "0"), Source = "Sales sheet", #"Source row" = r[Source row],
          RawKey = r[RawKey],
          #"Name typed" = T(r[First]) & " " & T(r[Last]), #"First typed" = T(r[First]), #"Last typed" = T(r[Last]),
          #"Email raw" = r[Email], #"Phone raw" = r[Phone], #"Company raw" = r[Company], #"Address raw" = r[Address],
          #"City raw" = r[City], #"State raw" = r[State], #"ZIP raw" = r[ZIP], #"Opt-in raw" = r[Newsletter],
          #"Company filled" = r[Company filled], #"Last activity" = AsDate(r[Last contact]), Spend = 0]))
in
    R"""

Q["Records"] = """let
    Nick     = Excel.CurrentWorkbook(){[Name = "Nicknames"]}[Content],
    NickRec  = Record.FromList(List.Transform(Nick[Full name], each Text.Upper(Text.Trim(_))),
                               List.Transform(Nick[Nickname], each Text.Upper(Text.Trim(_)))),
    StateRec = [COLORADO = "CO", UTAH = "UT", WASHINGTON = "WA", OREGON = "OR", CALIFORNIA = "CA", ARIZONA = "AZ",
                #"NEW MEXICO" = "NM", TEXAS = "TX", ILLINOIS = "IL", MINNESOTA = "MN", MASSACHUSETTS = "MA",
                CONNECTICUT = "CT", #"NEW JERSEY" = "NJ", VERMONT = "VT", MAINE = "ME", #"NEW YORK" = "NY"],
    AbbrRec  = [STREET = "St", ST = "St", AVENUE = "Ave", AVE = "Ave", ROAD = "Rd", RD = "Rd", DRIVE = "Dr", DR = "Dr",
                LANE = "Ln", LN = "Ln", COURT = "Ct", CT = "Ct", BOULEVARD = "Blvd", BLVD = "Blvd"],
    NoCo     = {"N/A", "NA", "NONE", "-", "SELF-EMPLOYED"},
    T        = (x) => if x = null then "" else Text.Trim(Text.From(x)),
    Squash   = (s as text) as text => Text.Combine(List.Select(Text.Split(Text.Trim(s), " "), each _ <> ""), " "),
    Digits   = (s as text) as text => Text.Select(s, {"0".."9"}),
    Proper   = (s as text) as text => Text.Proper(Text.Lower(s)),
    Clean    = (r as record) as record => let
        Split   = T(r[First typed]) <> "" or T(r[Last typed]) <> "",
        S       = Squash(T(r[Name typed])),
        Flip    = not Split and Text.Contains(S, ","),
        // "Smith, Anna" -> Anna Smith ; "anna  SMITH" -> Anna Smith
        Parts   = if Split then {Squash(T(r[First typed])), Squash(T(r[Last typed]))}
                  else if Flip then let p = Text.Split(S, ",") in {Text.Trim(Text.Combine(List.Skip(p, 1), ",")), Text.Trim(p{0})}
                  else let p = Text.Split(S, " ") in {p{0}, Text.Combine(List.Skip(p, 1), " ")},
        First   = Proper(Parts{0}),
        Last    = Proper(Parts{1}),
        Canon   = Record.FieldOrDefault(NickRec, Text.Upper(First), Text.Upper(First)),
        E       = Text.Lower(Text.Remove(T(r[Email raw]), {" "})),
        Email   = if E = "" then null else E,
        D0      = Digits(T(r[Phone raw])),
        D       = if Text.Length(D0) = 11 and Text.StartsWith(D0, "1") then Text.Middle(D0, 1) else D0,
        Phone   = if Text.Length(D) = 10 then "(" & Text.Start(D, 3) & ") " & Text.Middle(D, 3, 3) & "-" & Text.End(D, 4) else null,
        Z       = Digits(Text.Split(T(r[ZIP raw]), "-"){0}),
        Zip     = if Text.Length(Z) = 5 then Z else if Text.Length(Z) = 4 then "0" & Z else null,   // Excel ate the leading zero
        St0     = Text.Upper(T(r[State raw])),
        State   = if St0 = "" then null else if Text.Length(St0) = 2 then St0 else Record.FieldOrDefault(StateRec, St0, St0),
        A       = Squash(Text.Remove(T(r[Address raw]), {"."})),
        Address = if A = "" then null else Text.Combine(List.Transform(Text.Split(A, " "),
                      each Record.FieldOrDefault(AbbrRec, Text.Upper(_), Proper(_))), " "),
        C       = Proper(Squash(T(r[City raw]))),
        Co      = Squash(T(r[Company raw])),
        O       = Text.Upper(T(r[#"Opt-in raw"])),
        Typed   = T(r[Name typed]),
        Up      = Text.Upper(Typed),
        // the values exactly as they arrived (for the before/after score and the before/after card)
        Raw     = (x) => if x = null then "" else Text.From(x),
        RawList = {Raw(r[Email raw]), Raw(r[Phone raw]), Raw(r[Address raw]), Raw(r[City raw]), Raw(r[State raw]), Raw(r[ZIP raw])},
        OrNull  = (s as text) => if s = "" then null else s
      in [Record = r[Record], Source = r[Source], #"Source row" = r[Source row], RawKey = r[RawKey], #"Name typed" = Typed,
          First = First, Last = Last, Canon = Canon, Email = Email, Phone = Phone,
          Company = if Co = "" or List.Contains(NoCo, Text.Upper(Co)) then null else Co,
          Address = Address, City = if C = "" then null else C, State = State, ZIP = Zip,
          #"Opt-in" = if O = "YES" or O = "Y" then "Yes" else if O = "NO" or O = "N" then "No" else null,
          #"Last activity" = r[Last activity], Spend = r[Spend],
          Priority = List.PositionOf({"Webshop", "POS", "Sales sheet"}, r[Source]) + 1,
          Filled   = List.Count(List.RemoveNulls({Email, Phone, if Co = "" or List.Contains(NoCo, Text.Upper(Co)) then null else Co, Address})),
          KName    = if First <> "" and Last <> "" and Zip <> null then Canon & "|" & Text.Upper(Last) & "|" & Zip else null,
          RA       = if First <> "" and Last <> "" and Zip <> null then Text.Start(Canon, 1) & "|" & Text.Upper(Last) & "|" & Zip else null,
          RB       = if First <> "" and Last <> "" then Canon & "|" & Text.Upper(Last) else null,
          Junk     = Text.Contains(Up, "TEST") or (Email <> null and Text.StartsWith(Email, "test@"))
                     or Text.Contains(Text.Replace(Up, "-", " "), "WALK IN"),
          NoContact = Last = "" and Email = null and Phone = null,
          // flags for the Cleaning log
          fFlip = Flip, fName = not Flip and Typed <> "" and Typed <> First & (if Last = "" then "" else " " & Last),
          fNick = First <> "" and Canon <> Text.Upper(First),
          fEmail = Email <> null and T(r[Email raw]) <> Email, fPhone = Phone <> null and T(r[Phone raw]) <> Phone,
          fPhoneBad = T(r[Phone raw]) <> "" and Phone = null, fZip0 = Text.Length(Z) = 4,
          fZip4 = Text.Contains(T(r[ZIP raw]), "-"), fState = Text.Length(St0) > 2,
          fAddr = Address <> null and T(r[Address raw]) <> Address, fCo = r[Company filled],
          #"Email as received" = OrNull(RawList{0}), #"Phone as received" = OrNull(RawList{1}),
          #"Address as received" = OrNull(RawList{2}), #"City as received" = OrNull(RawList{3}),
          #"State as received" = OrNull(RawList{4}), #"ZIP as received" = OrNull(RawList{5}),
          qEmail = IsEmail(RawList{0}), qPhone = IsPhone(RawList{1}), qZip = IsZip(RawList{5}),
          qFilled = List.Count(List.Select(RawList, each Text.Trim(_) <> ""))],
    All      = Table.Combine({Webshop_In, POS_In, Sales_In}),
    Ordered  = Table.AddIndexColumn(All, "Order", 0, 1, Int64.Type),
    Cleaned  = Table.FromRecords(Table.TransformRows(Ordered, (r) => Record.AddField(Clean(r), "Order", r[Order]))),
    // a row the export repeated exactly: the first one stays
    FirstOf  = Table.Group(Cleaned, {"Source", "RawKey"}, {{"m", each List.Min([Order])}}),
    FirstSet = List.Buffer(FirstOf[m]),
    Status   = Table.AddColumn(Cleaned, "Status", each
                  if not List.Contains(FirstSet, [Order]) then "Discarded: exact duplicate row"
                  else if [Junk] then "Discarded: test or walk-in record"
                  else if [NoContact] then "Discarded: no name or contact"
                  else "Kept", type text),
    Typed    = Table.TransformColumnTypes(Status, {{"Last activity", type date}, {"Spend", type number}, {"Source row", Int64.Type}})
in
    Table.Buffer(Typed)"""

Q["Keyed"] = """let
    // records that share an email, a phone, or first name + last name + ZIP belong to one customer.
    // Every record starts in its own group; each pass gives it the lowest group number among the records it shares
    // a key with, until nothing changes (so A=B by email and B=C by phone puts A, B and C together).
    Kept   = Table.SelectRows(Records, each [Status] = "Kept"),
    Start  = Table.AddIndexColumn(Kept, "G", 1, 1, Int64.Type),
    MinMap = (t as table, k as text) as record => let
                 rows = Table.SelectRows(t, each Record.Field(_, k) <> null),
                 g    = Table.Group(rows, {k}, {{"m", each List.Min([G])}})
             in Record.FromList(g[m], Table.Column(g, k)),
    Look   = (m as record, v) => if v = null then null else Record.Field(m, v),
    Step   = (t as table) as table => let
                 E = MinMap(t, "Email"), P = MinMap(t, "Phone"), N = MinMap(t, "KName"),
                 New = Table.AddColumn(t, "G2", each List.Min(List.RemoveNulls({[G], Look(E, [Email]), Look(P, [Phone]), Look(N, [KName])})))
             in Table.Buffer(Table.RenameColumns(Table.RemoveColumns(New, {"G"}), {{"G2", "G"}})),
    Run    = (t as table, n as number) as table => let s = Step(t) in
                 if n >= 50 or List.Sum(s[G]) = List.Sum(t[G]) then s else @Run(s, n + 1),
    Auto   = Run(Table.Buffer(Start), 1),
    // G0 = the group the three rules found. Pairs marked Merge on Needs review join two G0 groups into one G.
    Base   = Table.RenameColumns(Auto, {{"G", "G0"}}),
    RecG   = Record.FromList(Auto[G], Auto[Record]),
    Edges  = List.Select(List.Transform(Table.ToRecords(Table.SelectRows(Decisions, each [Decision] = "Merge")),
                 each {Record.FieldOrDefault(RecG, [Master A], null), Record.FieldOrDefault(RecG, [Master B], null)}),
                 each _{0} <> null and _{1} <> null and _{0} <> _{1}),
    Pass   = (lbl as record) as record => List.Accumulate(Edges, lbl, (s, e) => let
                 a = Text.From(e{0}), b = Text.From(e{1}),
                 m = List.Min({Record.FieldOrDefault(s, a, e{0}), Record.FieldOrDefault(s, b, e{1})})
             in Record.Combine({s, Record.FromList({m, m}, {a, b})})),
    Labels = List.Accumulate({0..List.Count(Edges)}, [], (s, i) => Pass(s)),
    Final  = Table.AddColumn(Base, "G", each Record.FieldOrDefault(Labels, Text.From([G0]), [G0]), Int64.Type)
in
    Table.Buffer(Final)"""

Q["CustomersAll"] = """let
    // inside a group: most recent record first (then the most complete, then webshop > POS > sales sheet)
    Groups = Table.Group(Keyed, {"G"}, {{"Rows", each Table.Sort(_, {{"Last activity", Order.Descending},
                 {"Filled", Order.Descending}, {"Priority", Order.Ascending}, {"Order", Order.Ascending}})}}),
    Build  = (T as table) as record => let
        M      = T{0},
        Named  = Table.SelectRows(T, each [Last] <> ""),
        Nm     = if Table.IsEmpty(Named) then M else Named{0},
        HasAd  = Table.SelectRows(T, each [Address] <> null),
        Ad     = if Table.IsEmpty(HasAd) then M else HasAd{0},
        // every field comes from the most recent record that has it
        FirstNN = (c as text) => List.First(List.RemoveNulls(Table.Column(T, c)), null),
        Email  = FirstNN("Email"),
        Others = List.Sort(List.Distinct(List.Select(List.RemoveNulls(T[Email]), each _ <> Email)), Comparer.Ordinal),
        Opts   = T[#"Opt-in"]
      in [#"First name" = Nm[First], #"Last name" = Nm[Last], Email = Email,
          #"Other emails" = if List.IsEmpty(Others) then null else Text.Combine(Others, "; "),
          Phone = FirstNN("Phone"), Company = FirstNN("Company"), Address = Ad[Address], City = Ad[City],
          State = Ad[State], ZIP = Ad[ZIP],
          // an opt-out anywhere wins: never email someone who said no in one of the systems
          #"Marketing opt-in" = if List.Contains(Opts, "No") then "No" else if List.Contains(Opts, "Yes") then "Yes" else "Unknown",
          #"Lifetime spend" = Number.Round(List.Sum(T[Spend]), 2), #"Last activity" = M[Last activity],
          Sources = Text.Combine(List.Select({"Webshop", "POS", "Sales sheet"}, each List.Contains(T[Source], _)), " + "),
          #"Records merged" = Table.RowCount(T), #"Master record" = M[Record]],
    Cust   = Table.AddColumn(Groups, "C", each Build([Rows])),
    Wide   = Table.ExpandRecordColumn(Cust, "C", {"First name", "Last name", "Email", "Other emails", "Phone", "Company",
                 "Address", "City", "State", "ZIP", "Marketing opt-in", "Lifetime spend", "Last activity", "Sources",
                 "Records merged", "Master record"}),
    Sorted = Table.Sort(Wide, {{"Last name", Order.Ascending}, {"First name", Order.Ascending}, {"Master record", Order.Ascending}}),
    Ids    = Table.AddIndexColumn(Sorted, "N", 1, 1, Int64.Type),
    WithId = Table.AddColumn(Ids, "Customer ID", each "C" & Text.PadStart(Text.From([N]), 5, "0"), type text),
    Typed  = Table.TransformColumnTypes(WithId, {{"Lifetime spend", type number}, {"Last activity", type date}, {"Records merged", Int64.Type}})
in
    Table.Buffer(Typed)"""

Q["Customers"] = """let
    Source = Table.SelectColumns(CustomersAll, {"Customer ID", "First name", "Last name", "Email", "Phone", "Company",
                 "Address", "City", "State", "ZIP", "Marketing opt-in", "Lifetime spend", "Last activity", "Sources",
                 "Records merged", "Master record", "Other emails"})
in
    Source"""

Q["MergeLog"] = """let
    Multi = Table.SelectRows(CustomersAll, each [Records merged] > 1),
    Logs  = List.Transform(Table.ToRecords(Multi), (c) => let
              T    = c[Rows],
              Cnt  = (col as text, v) => List.Count(List.Select(Table.Column(T, col), each _ = v)),
              Rest = Table.Skip(T, 1),
              G0M  = T{0}[G0],
              Why  = Table.AddColumn(Rest, "Matched by", each
                         if [G0] <> G0M then "Reviewer decision"
                         else if [Email] <> null and Cnt("Email", [Email]) > 1 then "Same email"
                         else if [Phone] <> null and Cnt("Phone", [Phone]) > 1 then "Same phone"
                         else "Same name and ZIP", type text),
              Into = Table.AddColumn(Table.AddColumn(Table.AddColumn(Why, "Merged into", each c[Customer ID], type text),
                         "Customer", each c[First name] & " " & c[Last name], type text), "Master record", each c[Master record], type text),
              Note = Table.AddColumn(Into, "Note", each
                         if [Matched by] = "Reviewer decision" then "Marked Merge on Needs review"
                         else if [Email] <> null and [Email] <> c[Email] then "Different email: kept under Other emails"
                         else if [First] <> "" and [Canon] <> Text.Upper([First]) then "Nickname: " & [First] & " = " & Text.Proper([Canon])
                         else "", type text)
            in Note),
    All   = Table.Combine(Logs),
    Out   = Table.SelectColumns(All, {"Record", "Source", "Name typed", "Email", "Phone", "ZIP", "Last activity",
                 "Merged into", "Customer", "Master record", "Matched by", "Note"}),
    Named = Table.RenameColumns(Out, {{"Name typed", "Name as typed"}}),
    Typed = Table.TransformColumnTypes(Named, {{"Last activity", type date}})
in
    Table.Sort(Typed, {{"Merged into", Order.Ascending}, {"Record", Order.Ascending}})"""

Q["Review"] = """let
    // look-alikes in different groups: similar first name with the same last name and ZIP, or the same name in another ZIP.
    // Pairs are found on the groups the rules made (G0), so a pair stays listed after a person marks it Merge.
    Pairs  = (k as text, reason as text) as table => let
                 rows  = Table.SelectRows(Keyed, each Record.Field(_, k) <> null),
                 g     = Table.Group(rows, {k}, {{"Gs", each List.Sort(List.Distinct([G0]))}}),
                 multi = Table.SelectRows(g, each List.Count([Gs]) > 1),
                 recs  = List.Combine(List.Transform(multi[Gs], (gs) => List.Combine(List.Transform(List.Positions(gs),
                             (i) => List.Transform(List.Skip(gs, i + 1), (b) => [A = gs{i}, B = b, Reason = reason])))))
             in Table.FromRecords(recs, type table [A = Int64.Type, B = Int64.Type, Reason = text]),
    Both   = Table.Distinct(Table.Combine({Pairs("RA", "Similar first name, same last name and ZIP"),
                                           Pairs("RB", "Same name, different ZIP")}), {"A", "B"}),
    // one line per rule-made group: its master record and the values its customer would have on its own
    Side   = Table.Group(Keyed, {"G0"}, {{"S", (t) => let
                 T  = Table.Sort(t, {{"Last activity", Order.Descending}, {"Filled", Order.Descending},
                          {"Priority", Order.Ascending}, {"Order", Order.Ascending}}),
                 M  = T{0},
                 Nm = let n = Table.SelectRows(T, each [Last] <> "") in if Table.IsEmpty(n) then M else n{0},
                 Ad = let n = Table.SelectRows(T, each [Address] <> null) in if Table.IsEmpty(n) then M else n{0},
                 NN = (c as text) => List.First(List.RemoveNulls(Table.Column(T, c)), null)
               in [Master = M[Record], G = M[G], F = Nm[First], L = Nm[Last], Email = NN("Email"), Phone = NN("Phone"),
                   City = Ad[City], ZIP = Ad[ZIP]]}}),
    SideT  = Table.ExpandRecordColumn(Side, "S", {"Master", "G", "F", "L", "Email", "Phone", "City", "ZIP"}),
    Ids    = Table.SelectColumns(CustomersAll, {"G", "Customer ID"}),
    SideC  = Table.ExpandTableColumn(Table.NestedJoin(SideT, {"G"}, Ids, {"G"}, "c", JoinKind.LeftOuter), "c", {"Customer ID"}),
    Cols   = {"Customer ID", "F", "L", "Email", "Phone", "City", "ZIP", "Master"},
    JA     = Table.ExpandTableColumn(Table.NestedJoin(Both, {"A"}, SideC, {"G0"}, "a", JoinKind.LeftOuter), "a", Cols,
                 {"Customer A", "FA", "LA", "Email A", "Phone A", "City A", "ZIP A", "Master A"}),
    JB     = Table.ExpandTableColumn(Table.NestedJoin(JA, {"B"}, SideC, {"G0"}, "b", JoinKind.LeftOuter), "b", Cols,
                 {"Customer B", "FB", "LB", "Email B", "Phone B", "City B", "ZIP B", "Master B"}),
    Names  = Table.AddColumn(Table.AddColumn(JB, "Name A", each [FA] & " " & [LA], type text), "Name B", each [FB] & " " & [LB], type text),
    Check  = Table.AddColumn(Names, "What to check", each
                 if Text.StartsWith([Reason], "Similar") then [FA] & " / " & [FB] & ": different email and phone, probably two people. Merge only if confirmed."
                 else [City A] & " / " & [City B] & ": two people, or one who moved? Check order history first.", type text),
    // the decision typed on this sheet before the refresh (see Decisions); blank = still open
    DecRec = Record.FromList(Decisions[Decision], Decisions[Pair]),
    Dec    = Table.AddColumn(Check, "Decision", each Record.FieldOrDefault(DecRec, [Master A] & "|" & [Master B], ""), type text),
    Res    = Table.AddColumn(Dec, "Result", each
                 if [Decision] = "Merge" then (if [Customer A] = [Customer B] then "Merged into " & [Customer A] else "Merge: refresh again")
                 else if [Decision] = "Keep separate" then "Kept as two customers" else "Open: waiting for a decision", type text),
    Out    = Table.SelectColumns(Res, {"Decision", "Result", "Reason", "Customer A", "Name A", "Email A", "Phone A", "City A", "ZIP A",
                 "Customer B", "Name B", "Email B", "Phone B", "City B", "ZIP B", "What to check", "Master A", "Master B"})
in
    Table.Sort(Out, {{"Reason", Order.Ascending}, {"Master A", Order.Ascending}})"""

Q["AllRecords"] = """let
    G      = Table.SelectColumns(Keyed, {"Record", "G"}),
    C      = Table.SelectColumns(CustomersAll, {"G", "Customer ID", "Master record"}),
    L      = Table.SelectColumns(MergeLog, {"Record", "Matched by"}),
    J1     = Table.ExpandTableColumn(Table.NestedJoin(Records, {"Record"}, G, {"Record"}, "g", JoinKind.LeftOuter), "g", {"G"}),
    J2     = Table.ExpandTableColumn(Table.NestedJoin(J1, {"G"}, C, {"G"}, "c", JoinKind.LeftOuter), "c", {"Customer ID", "Master record"}),
    J3     = Table.ExpandTableColumn(Table.NestedJoin(J2, {"Record"}, L, {"Record"}, "l", JoinKind.LeftOuter), "l", {"Matched by"}),
    St     = Table.AddColumn(J3, "Result", each
                 if [Status] <> "Kept" then [Status] else if [Record] = [Master record] then "Master" else "Merged", type text),
    Sorted = Table.Sort(St, {{"Order", Order.Ascending}}),
    Out    = Table.SelectColumns(Sorted, {"Record", "Source", "Source row", "Result", "Customer ID", "Matched by", "Name typed",
                 "First", "Last", "Email", "Phone", "Company", "Address", "City", "State", "ZIP", "Opt-in", "Last activity", "Spend",
                 "Email as received", "Phone as received", "Address as received", "City as received", "State as received",
                 "ZIP as received"})
in
    Table.RenameColumns(Out, {{"Result", "Status"}, {"Name typed", "Name as typed"}})"""

Q["Quality"] = """let
    // data quality before (every row as it arrived) and after (one row per customer). Same tests both sides.
    R    = Records,
    C    = CustomersAll,
    nR   = Table.RowCount(R),
    nC   = Table.RowCount(C),
    Cnt  = (t as table, f as function) as number => Table.RowCount(Table.SelectRows(t, f)),
    S    = (x) => if x = null then "" else Text.From(x),
    CF   = List.Sum(List.Transform(Table.ToRecords(C), (c) => List.Count(List.Select({c[Email], c[Phone], c[Address], c[City],
               c[State], c[ZIP]}, each _ <> null and Text.Trim(Text.From(_)) <> "")))),
    Rows = {
        {"Valid email", "Email present and written correctly: lower case, no spaces, one @, a dot in the domain",
            Cnt(R, each [qEmail]), nR, Cnt(C, each IsEmail(S([Email]))), nC},
        {"Formatted phone", "Phone present and written as (303) 555-0142",
            Cnt(R, each [qPhone]), nR, Cnt(C, each IsPhone(S([Phone]))), nC},
        {"ZIP complete", "ZIP present as exactly 5 digits",
            Cnt(R, each [qZip]), nR, Cnt(C, each IsZip(S([ZIP]))), nC},
        {"Fields filled", "Email, phone, address, city, state and ZIP: share of the boxes that are filled",
            List.Sum(R[qFilled]), nR * 6, CF, nC * 6}},
    T    = #table(type table [Metric = text, Meaning = text, #"Before ok" = Int64.Type, #"Before of" = Int64.Type,
               #"After ok" = Int64.Type, #"After of" = Int64.Type], Rows),
    P    = Table.AddColumn(Table.AddColumn(T, "Before", each [Before ok] / [Before of], type number),
               "After", each [After ok] / [After of], type number)
in
    P"""

Q["CRM_Import"] = """let
    // column names a CRM import wizard recognises (HubSpot / Pipedrive map them automatically)
    Source = Table.SelectColumns(CustomersAll, {"First name", "Last name", "Email", "Phone", "Company", "Address", "City",
                 "State", "ZIP", "Marketing opt-in", "Lifetime spend", "Customer ID", "Other emails"}),
    Named  = Table.RenameColumns(Source, {{"First name", "First Name"}, {"Last name", "Last Name"}, {"Phone", "Phone Number"},
                 {"Company", "Company Name"}, {"Address", "Street Address"}, {"State", "State/Region"}, {"ZIP", "Postal Code"},
                 {"Marketing opt-in", "Marketing Opt-in"}, {"Lifetime spend", "Lifetime Spend"}, {"Customer ID", "External ID"},
                 {"Other emails", "Additional Emails"}})
in
    Named"""

Q["Cleaning_Log"] = """let
    R     = Records,
    Src   = {"Webshop", "POS", "Sales sheet"},
    N     = (s as text, f as function) as number => Table.RowCount(Table.SelectRows(R, each [Source] = s and f(_))),
    Line  = (check as text, f as function, done as text) as list =>
                List.Transform(Src, (s) => {s, check, N(s, f), done}),
    Rows  = List.Combine({
        Line("Rows received", each true, "Every row read as text first, so nothing is lost on import"),
        Line("Exact duplicate rows", each [Status] = "Discarded: exact duplicate row", "Removed (the export repeated them)"),
        Line("Test accounts and walk-in rows", each [Status] = "Discarded: test or walk-in record", "Removed; listed on All records"),
        Line("Rows with no name, email or phone", each [Status] = "Discarded: no name or contact", "Removed; listed on All records"),
        Line("Names typed Last, First", each [fFlip], "Turned around: Smith, Anna -> Anna Smith"),
        Line("Names in capitals, lower case or with extra spaces", each [fName], "Proper case, single spaces"),
        Line("Nicknames recognised (Bill, Kate, Mike ...)", each [fNick], "Matched on the full name from the Name rules tab; the name itself is kept"),
        Line("Emails with capitals or spaces", each [fEmail], "Lower case, spaces removed"),
        Line("Phones in another format", each [fPhone], "Written as (303) 555-0142; +1 removed"),
        Line("Phones that are not 10 digits", each [fPhoneBad], "Left blank (not guessed)"),
        Line("ZIPs that lost their leading zero (2139)", each [fZip0], "Restored to 02139"),
        Line("ZIP+4 codes (80302-1234)", each [fZip4], "Shortened to the 5-digit ZIP"),
        Line("States written out (Colorado)", each [fState], "Two-letter code (CO)"),
        Line("Addresses with mixed case, dots or long street types", each [fAddr], "Proper case; Street -> St, Avenue -> Ave"),
        Line("Company only in the first row of merged cells", each [fCo], "Filled down to every contact of that company")}),
    Log   = Table.FromRows(List.Select(Rows, each _{0} <> "x"), type table [Source = text, Check = text, Rows = Int64.Type, #"What was done" = text])
in
    Table.SelectRows(Log, each [Rows] > 0 or [Check] = "Rows received")"""

LOADED = [("Summary", None), ("Match explorer", None), ("Customers", "Customers"), ("Needs review", "Review"),
          ("Merge log", "MergeLog"), ("CRM import", "CRM_Import"), ("All records", "AllRecords"),
          ("Cleaning log", "Cleaning_Log"), ("Quality", "Quality"), ("Name rules", None), ("Settings", None), ("How to use", None)]
RECLIST = ('=LET(id,{id},m,AllRecords[Customer ID]=id,n,SUM(--m),IF(OR(id="",n=0),"",IF(n=1,FILTER(AllRecords[Record],m),'
           'VSTACK(FILTER(AllRecords[Record],m*(AllRecords[Status]="Master")),FILTER(AllRecords[Record],m*(AllRecords[Status]<>"Master"))))))')
TABLE_ROW = 5  # loaded tables start here, under the dark header band with the navigation buttons


def hx(h):
    h = h.lstrip("#")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


# dark luxury palette (same as the E9 Power BI report)
BG, PAGE, CARD, CARD2, LINE = map(hx, ("0B0F13", "12171D", "1A2129", "202832", "2A333D"))
GOLD, GOLD_LT, IVORY, MUTED = map(hx, ("C9A35B", "E6CF9A", "ECE6DA", "8E98A3"))
JADE, COPPER, SKY, GREY = map(hx, ("4FA38C", "B5754A", "7FA7C9", "4A5560"))
GOOD, BAD, AMBER = map(hx, ("5BBF8A", "E5735F", "D9A441"))
T_GOOD, T_BAD, T_COPPER, T_SKY = map(hx, ("1E3A2E", "45231E", "3A2A20", "1C2C3B"))  # dark tints for fills
ICON, SERIF, SANS = "Segoe MDL2 Assets", "Georgia", "Segoe UI"
YELLOW = (255, 255, 153)

# sheet, button label, Segoe MDL2 glyph
PAGES = [("Summary", "Cockpit", ""), ("Match explorer", "Match explorer", ""), ("Customers", "Customers", ""),
         ("Merge log", "Merge log", ""), ("Needs review", "Needs review", ""), ("CRM import", "CRM import", ""),
         ("All records", "All records", ""), ("Cleaning log", "Cleaning log", "")]


def rgb(r, g, b):
    return r + g * 256 + b * 65536


def conn_string(name):
    return f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=""'


def load_query(wb, ws, name, table_name, at=f"A{TABLE_ROW}"):
    lo = ws.ListObjects.Add(0, conn_string(name), None, 1, ws.Range(at))
    qt = lo.QueryTable
    qt.CommandType = 2
    qt.CommandText = f"SELECT * FROM [{name}]"
    qt.RefreshStyle = 1
    qt.AdjustColumnWidth = False
    qt.BackgroundQuery = False
    qt.Refresh(False)
    lo.Name = table_name
    lo.TableStyle = ""
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
    return rng


def put(ws, addr, value, merge=True, **kw):
    """Value or formula into a (merged) range, then format it."""
    rng = ws.Range(addr)
    if merge and rng.Cells.Count > 1:
        rng.Merge()
    cell = rng.Cells(1, 1)
    if isinstance(value, str) and value.startswith("="):
        cell.Formula2 = value
    elif value is not None:
        cell.Value = value
    fmt(rng, **kw)
    return rng


def cf_formula(rng, formula, font=None, fill=None, bold=None, first=False):
    """Conditional format by formula. Formulas here never use a comma: rule formulas use the local list separator."""
    fc = rng.FormatConditions.Add(2, None, formula)
    if first:
        fc.SetFirstPriority()
    if font is not None:
        fc.Font.Color = font
    if fill is not None:
        fc.Interior.Color = fill
    if bold is not None:
        fc.Font.Bold = bold
    return fc


def pass_fail(rng):
    for op, fill, font in ((3, T_GOOD, GOOD), (4, T_BAD, BAD)):
        fc = rng.FormatConditions.Add(1, op, '="PASS"')
        fc.Interior.Color = fill
        fc.Font.Color = font


def button(ws, left, top, w, h, glyph, label, target, current, align_left=True, size=10):
    shp = ws.Shapes.AddShape(5, left, top, w, h)  # rounded rectangle
    try:
        shp.Adjustments.SetItem(1, 0.18)
    except Exception:
        pass
    shp.Name = f"nav_{label}"
    shp.Fill.ForeColor.RGB = GOLD if current else CARD
    shp.Line.ForeColor.RGB = GOLD if current else LINE
    shp.Line.Weight = 0.75
    shp.Shadow.Visible = False
    tf = shp.TextFrame2
    tf.MarginLeft, tf.MarginRight, tf.MarginTop, tf.MarginBottom = 9, 4, 0, 0
    tf.VerticalAnchor = 3
    tf.WordWrap = False
    tr = tf.TextRange
    tr.Text = f"{glyph}   {label}"
    tr.ParagraphFormat.Alignment = 1 if align_left else 2
    tr.Font.Name = SANS
    tr.Font.Size = size
    tr.Font.Bold = current
    tr.Font.Fill.ForeColor.RGB = BG if current else IVORY
    g = shp.TextFrame.Characters(1, 1)
    g.Font.Name = ICON
    g.Font.Size = size + 1
    g.Font.Color = BG if current else GOLD
    ws.Hyperlinks.Add(shp, "", f"'{target}'!A1", f"Go to {label}")
    shp.Placement = 3
    return shp


def top_nav(ws, current, top):
    x = 6
    for sheet, label, glyph in PAGES:
        w = 104 if len(label) < 12 else 116
        button(ws, x, top, w, 24, glyph, label, sheet, sheet == current, align_left=False, size=9)
        x += w + 6


def page_header(ws, title, sub, current, width_cols=40):
    """Dark header band: serif title, spaced gold brand, navigation buttons, thin gold rule."""
    ws.Cells.Interior.Color = PAGE
    ws.Cells.Font.Color = IVORY
    ws.Range(ws.Cells(1, 1), ws.Cells(3, width_cols)).Interior.Color = BG
    ws.Range(ws.Cells(4, 1), ws.Cells(4, width_cols)).Interior.Color = GOLD
    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 16
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 2.5
    c = ws.Range("A1")
    c.Value = title
    fmt(c, color=IVORY, size=18, name=SERIF, v="c")
    c = ws.Range("A2")
    c.Value = "S U M M I T   R I D G E   O U T D O O R     ·     " + sub
    fmt(c, color=GOLD, size=8.5, v="c")
    top_nav(ws, current, ws.Range("A3").Top + 4)


def header_row(lo):
    h = lo.HeaderRowRange
    fmt(h, fill=CARD2, color=GOLD_LT, bold=True, size=9.5, wrap=True, v="c")
    h.Borders(9).LineStyle = 1
    h.Borders(9).Color = GOLD
    h.RowHeight = 32


def band(ws, lo, last_col, first_col=1):
    """Zebra rows for a query table (conditional fill, so it grows with the table on refresh)."""
    first = lo.HeaderRowRange.Row + 1
    rng = ws.Range(ws.Cells(first, first_col), ws.Cells(first + 4000, last_col))
    cf_formula(rng, f"=ISEVEN(ROW())*($A{first}<>\"\")", fill=CARD)
    cf_formula(rng, f"=ISODD(ROW())*($A{first}<>\"\")", fill=PAGE)


# ---------------------------------------------------------------- data sheets
SUBS = {"Customers": "one row per customer: newest values, blanks filled from the duplicates, spend added up",
        "Needs review": "look-alikes are not merged automatically: type Merge or Keep separate in Decision, then Refresh All",
        "Merge log": "every record folded into a customer, the customer it went into and the rule that linked it",
        "CRM import": "the same customers with the column names HubSpot / Pipedrive / Mailchimp recognise",
        "All records": "every source row once, with its fate: master, merged into a customer, or set aside",
        "Cleaning log": "every cleaning step, counted per source",
        "Quality": "data quality before (rows as received) and after (one row per customer); feeds the cockpit gauges"}
TITLES = {"Customers": "Customers", "Needs review": "Needs review", "Merge log": "Merge log", "CRM import": "CRM import",
          "All records": "All records", "Cleaning log": "Cleaning log", "Quality": "Data quality score"}


def format_tables(wb):
    nfmt = {"Lifetime spend": "#,##0.00", "Lifetime Spend": "#,##0.00", "Spend": "#,##0.00", "Last activity": "yyyy-mm-dd",
            "Before": "0.0%", "After": "0.0%"}
    widths = {"Customers": {"Customer ID": 9, "First name": 11, "Last name": 11, "Email": 30, "Phone": 15, "Company": 24,
                            "Address": 22, "City": 14, "State": 6, "ZIP": 7, "Marketing opt-in": 9, "Lifetime spend": 11,
                            "Last activity": 11, "Sources": 24, "Records merged": 8, "Master record": 9, "Other emails": 34},
              "Needs review": {"Decision": 14, "Result": 25, "Reason": 37, "Customer A": 9, "Name A": 18, "Email A": 28,
                               "Phone A": 15, "City A": 13, "ZIP A": 7, "Customer B": 9, "Name B": 18, "Email B": 28,
                               "Phone B": 15, "City B": 13, "ZIP B": 7, "What to check": 79, "Master A": 9, "Master B": 9},
              "Merge log": {"Record": 8, "Source": 11, "Name as typed": 22, "Email": 30, "Phone": 15, "ZIP": 7,
                            "Last activity": 11, "Merged into": 9, "Customer": 20, "Master record": 9, "Matched by": 17, "Note": 40},
              "CRM import": {"Email": 30, "Street Address": 22, "Company Name": 24, "Additional Emails": 34, "City": 14,
                             "Phone Number": 15, "First Name": 11, "Last Name": 11},
              "All records": {"Status": 28, "Name as typed": 24, "Email": 30, "Address": 22, "Company": 22, "Phone": 15,
                              "Last activity": 11, "Matched by": 16, "Email as received": 30, "Phone as received": 16,
                              "Address as received": 24, "City as received": 14},
              "Cleaning log": {"Source": 12, "Check": 52, "Rows": 7, "What was done": 96},
              "Quality": {"Metric": 16, "Meaning": 80, "Before ok": 10, "Before of": 10, "After ok": 10, "After of": 10}}
    for title, w in widths.items():
        ws = wb.Worksheets(title)
        lo = ws.ListObjects(1)
        page_header(ws, TITLES[title], SUBS[title], title)
        n = lo.ListColumns.Count
        for i in range(1, n + 1):
            name = lo.ListColumns(i).Name
            ws.Columns(i).ColumnWidth = w.get(name, 10)
            if name in nfmt and lo.DataBodyRange is not None:
                lo.ListColumns(i).DataBodyRange.NumberFormat = nfmt[name]
        header_row(lo)
        band(ws, lo, n, 2 if title == "Needs review" else 1)
        ws.Activate()
        wb.Application.ActiveWindow.FreezePanes = False
        ws.Range(f"A{TABLE_ROW + 1}").Select()
        wb.Application.ActiveWindow.FreezePanes = True
    # status colours on All records, rule colours on the merge log
    lo = wb.Worksheets("All records").ListObjects(1)
    rng = lo.ListColumns("Status").DataBodyRange
    r0 = TABLE_ROW + 1
    for text, fill, font in (("Master", T_GOOD, GOOD), ("Merged", T_SKY, SKY),
                             ("Discarded: exact duplicate row", T_COPPER, COPPER),
                             ("Discarded: test or walk-in record", T_COPPER, COPPER),
                             ("Discarded: no name or contact", T_COPPER, COPPER)):
        cf_formula(rng, f'=$D{r0}="{text}"', font=font, fill=fill, first=True)
    lo = wb.Worksheets("Merge log").ListObjects(1)
    rng = lo.ListColumns("Matched by").DataBodyRange
    col = lo.ListColumns("Matched by").Range.Column
    L = chr(64 + col)
    for text, font in (("Same email", GOLD_LT), ("Same phone", SKY), ("Same name and ZIP", JADE), ("Reviewer decision", AMBER)):
        cf_formula(rng, f'=${L}{r0}="{text}"', font=font, bold=True, first=True)
    build_review_input(wb)


def build_review_input(wb):
    """Needs review: the Decision column is the one place a person types; drop-down Merge / Keep separate."""
    ws = wb.Worksheets("Needs review")
    lo = ws.ListObjects("Review")
    dec = lo.ListColumns("Decision").DataBodyRange
    fmt(dec, fill=CARD2, color=GOLD_LT, bold=True, h="c")
    for edge in (7, 8, 9, 10, 11, 12):
        dec.Borders(edge).LineStyle = 1
        dec.Borders(edge).Color = GOLD
    dec.Validation.Delete()
    dec.Validation.Add(3, 1, 1, "=DecisionList")
    dec.Validation.IgnoreBlank = True
    dec.Validation.InputTitle = "Decision"
    dec.Validation.InputMessage = "Merge = same person, join the two customers. Keep separate = two people. Then Data > Refresh All."
    res = lo.ListColumns("Result").DataBodyRange
    for text, font in (("Merged", GOOD), ("Kept", MUTED), ("Open", AMBER), ("Merge:", AMBER)):
        fc = res.FormatConditions.Add(9, None, None, None, text, 2)  # text begins with
        fc.Font.Color = font
        fc.Font.Bold = True


# ---------------------------------------------------------------- Summary: data quality cockpit
# content columns: D-G, I-L, N-Q, S-V, X-AA (H, M, R, W are narrow gutters); helper cells from AD
GROUPS = [(4, 7), (9, 12), (14, 17), (19, 22), (24, 27)]
HC = 30  # AD: helper cells for the charts and cards (outside the print area)


def A(col, row):
    s = ""
    while col:
        col, r = divmod(col - 1, 26)
        s = chr(65 + r) + s
    return f"{s}{row}"


def R(ws, r1, c1, r2=None, c2=None):
    return ws.Range(A(c1, r1) + ":" + A(c2 or c1, r2 or r1))


def panel(ws, r1, r2, c1, c2, fill=CARD):
    rng = R(ws, r1, c1, r2, c2)
    rng.Interior.Color = fill
    for e in (7, 8, 9, 10):
        rng.Borders(e).LineStyle = 1
        rng.Borders(e).Color = LINE
    return rng


def linked_box(ws, left, top, w, h, ref, size, color, bold=False, name=SANS, align=2):
    tb = ws.Shapes.AddTextbox(1, left, top, w, h)
    tb.Fill.Visible = False
    tb.Line.Visible = False
    try:
        tb.DrawingObject.Formula = ref
    except Exception:
        tb.OLEFormat.Object.Formula = ref
    tf = tb.TextFrame2
    tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
    tf.VerticalAnchor = 3
    tf.WordWrap = True
    tr = tf.TextRange
    tr.ParagraphFormat.Alignment = align
    tr.Font.Name = name
    tr.Font.Size = size
    tr.Font.Bold = bold
    tr.Font.Fill.ForeColor.RGB = color
    tb.Placement = 3
    return tb


def dark_chart(ch):
    ch.ChartArea.Format.Fill.Visible = False
    ch.ChartArea.Format.Line.Visible = False
    ch.PlotArea.Format.Fill.Visible = False
    ch.ChartArea.Font.Color = IVORY
    ch.ChartArea.Font.Name = SANS
    ch.HasLegend = False


def build_summary(wb):
    ws = wb.Worksheets("Summary")
    ws.Cells.Interior.Color = PAGE
    ws.Cells.Font.Color = IVORY
    widths = {1: 1.5, 2: 22, 3: 2}
    for c in range(4, 28):
        widths[c] = 1.3 if c in (8, 13, 18, 23) else 7.6
    widths[28] = 2
    widths[29] = 4
    for c in range(HC, HC + 10):
        widths[c] = 11
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    heights = {1: 10, 2: 36, 3: 18, 4: 2.5, 5: 12, 6: 5, 7: 18, 8: 36, 9: 16, 10: 7, 11: 12,
               12: 6, 13: 24, 14: 16, 23: 18, 24: 16, 25: 8, 26: 12, 27: 6, 28: 24, 29: 16, 30: 19, 40: 8, 41: 12,
               42: 6, 43: 24, 44: 16, 45: 19, 50: 22, 51: 16, 52: 8, 53: 12}
    for r in range(15, 23):
        heights[r] = 17
    for r in range(31, 40):
        heights[r] = 19
    for r in range(46, 50):
        heights[r] = 20
    for r, h in heights.items():
        ws.Rows(r).RowHeight = h
    R(ws, 1, 1, 3, 28).Interior.Color = BG
    R(ws, 4, 1, 4, 28).Interior.Color = GOLD
    put(ws, "B2", "S U M M I T   R I D G E", color=GOLD, size=9, bold=True, v="c")
    put(ws, "B3", "OUTDOOR  ·  CRM CLEANUP", color=MUTED, size=8, v="t")
    put(ws, "D2", "Data quality cockpit", color=IVORY, size=24, name=SERIF, v="c")
    put(ws, "D3", '="Customer list cleanup   ·   "&Period&"   ·   webshop + POS + sales sheet  →  one CRM-ready list   ·   fictional data"',
        color=MUTED, size=9, v="t")
    put(ws, "X2:AA2", "Data > Refresh All updates every number", color=MUTED, size=8, italic=True, h="r", v="c")

    # ---- helper cells (live formulas; the cards, gauges and charts read these)
    h = HC
    ws.Range("Z43").Value = "C00001"  # the example customer (replaced by choose_examples)
    wb.Names.Add("ExampleID", "=Summary!$Z$43")
    put(ws, A(h, 5), "Helper cells for the cockpit (live formulas)", color=MUTED, size=8, italic=True, merge=False)
    for j, t in enumerate(("Source", "Received", "Set aside", "Customers", "Merged")):
        put(ws, A(h + j, 6), t, color=MUTED, size=8, bold=True, merge=False)
    for i, src in enumerate(["Webshop", "POS", "Sales sheet"], 7):
        ws.Cells(i, h).Value = src
        ws.Cells(i, h + 1).Formula = f"=COUNTIFS(AllRecords[Source],{A(h, i)})"
        ws.Cells(i, h + 2).Formula = f'=COUNTIFS(AllRecords[Source],{A(h, i)},AllRecords[Status],"Discarded*")'
        ws.Cells(i, h + 3).Formula = f'=COUNTIFS(AllRecords[Source],{A(h, i)},AllRecords[Status],"Master")'
        ws.Cells(i, h + 4).Formula = f'=COUNTIFS(AllRecords[Source],{A(h, i)},AllRecords[Status],"Merged")'
    ws.Cells(10, h).Value = "Total"
    for j in range(1, 5):
        ws.Cells(10, h + j).Formula = f"=SUM({A(h + j, 7)}:{A(h + j, 9)})"
    rules = ["Same email", "Same phone", "Same name and ZIP", "Reviewer decision"]
    for i, rule in enumerate(rules, 12):
        ws.Cells(i, h).Value = rule
        ws.Cells(i, h + 1).Formula = f"=COUNTIFS(MergeLog[Matched by],{A(h, i)})"
    tot, setaside, merged_ = A(h + 1, 10), A(h + 2, 10), A(h + 4, 10)
    funnel = [(f'="Rows in"', f"={tot}"),
              (f'="Kept after removing "&TEXT({setaside},"#,##0")', f"={tot}-{setaside}"),
              (f'="Unique customers after merging "&TEXT({merged_},"#,##0")', "=ROWS(Customers[Customer ID])")]
    for i, (lab, val) in enumerate(funnel, 17):
        ws.Cells(i, h).Formula = lab
        ws.Cells(i, h + 1).Formula = val
        ws.Cells(i, h + 2).Formula = f"=(MAX({A(h + 1, 17)}:{A(h + 1, 19)})-{A(h + 1, i)})/2"
    for k in range(4):
        r = 21 + k
        ws.Cells(r, h).Formula = f"=INDEX(Quality[Metric],{k + 1})"
        ws.Cells(r, h + 1).Formula = f"=INDEX(Quality[Before],{k + 1})"
        ws.Cells(r, h + 2).Formula = f"=1-{A(h + 1, r)}"
        ws.Cells(r, h + 3).Formula = f"=INDEX(Quality[After],{k + 1})"
        ws.Cells(r, h + 4).Formula = f"=1-{A(h + 3, r)}"
        ws.Cells(r, h + 5).Formula = f'=TEXT({A(h + 3, r)},"0%")'
        ws.Cells(r, h + 6).Formula = f'="before "&TEXT({A(h + 1, r)},"0%")&"  →  after "&TEXT({A(h + 3, r)},"0.0%")'
        ws.Cells(r, h + 7).Formula = f'="+"&TEXT(({A(h + 3, r)}-{A(h + 1, r)})*100,"0.0")&" pts"'
    for i, (lab, st) in enumerate((("Repeated rows", "Discarded: exact duplicate row"), ("Test / walk-in", "Discarded: test or walk-in record"),
                                   ("No contact", "Discarded: no name or contact")), 26):
        ws.Cells(i, h).Value = lab
        ws.Cells(i, h + 1).Formula = f'=COUNTIFS(AllRecords[Status],"{st}")'
    ws.Cells(30, h).Value = "Decided"
    ws.Cells(30, h + 1).Formula = '=COUNTIFS(Review[Decision],"Merge")+COUNTIFS(Review[Decision],"Keep separate")'
    ws.Cells(31, h).Value = "Open"
    ws.Cells(31, h + 1).Formula = f"=ROWS(Review[Reason])-{A(h + 1, 30)}"
    fmt(R(ws, 6, h, 45, h + 9), color=MUTED, size=8)
    R(ws, 21, h + 1, 24, h + 4).NumberFormat = "0.0%"

    print("KPI cards", flush=True)
    # ---- KPI cards
    cards = [("ROWS IN", "", f"={tot}",
              f'="webshop "&TEXT({A(h + 1, 7)},"#,##0")&"  ·  POS "&TEXT({A(h + 1, 8)},"#,##0")&"  ·  sales "&TEXT({A(h + 1, 9)},"#,##0")', MUTED),
             ("SET ASIDE", "", f"={setaside}",
              f'={A(h + 1, 26)}&" repeated · "&{A(h + 1, 27)}&" test/walk-in · "&{A(h + 1, 28)}&" no contact"', MUTED),
             ("DUPLICATES MERGED", "", f"={merged_}",
              f'={A(h + 1, 12)}&" email · "&{A(h + 1, 13)}&" phone · "&{A(h + 1, 14)}&" name+ZIP"&IF({A(h + 1, 15)}>0," · "&{A(h + 1, 15)}&" reviewer","")', MUTED),
             ("UNIQUE CUSTOMERS", "", "=ROWS(Customers[Customer ID])",
              f'="▼ "&TEXT(1-ROWS(Customers[Customer ID])/{tot},"0.0%")&" fewer rows than received"', GOOD),
             ("PAIRS TO REVIEW", "", "=ROWS(Review[Reason])",
              f'={A(h + 1, 30)}&" decided  ·  "&{A(h + 1, 31)}&" open"', AMBER)]
    for (c1, c2), (label, glyph, val, note, note_color) in zip(GROUPS, cards):
        panel(ws, 6, 10, c1, c2)
        box = R(ws, 6, c1, 10, c1)
        box.Borders(7).Weight = 4
        box.Borders(7).Color = GOLD
        put(ws, A(c1, 7) + ":" + A(c2 - 1, 7), label, color=MUTED, size=8, bold=True, v="c")
        R(ws, 7, c1).IndentLevel = 1
        put(ws, A(c2, 7), glyph, color=GOLD, size=12, name=ICON, h="r", v="c")
        put(ws, A(c1, 8) + ":" + A(c2, 8), val, color=IVORY, size=24, name=SERIF, h="l", v="c", numfmt="#,##0")
        R(ws, 8, c1).IndentLevel = 1
        put(ws, A(c1, 9) + ":" + A(c2, 9), note, color=note_color, size=8, v="c")
        R(ws, 9, c1).IndentLevel = 1
        R(ws, 9, c1).ShrinkToFit = True
    cf_formula(R(ws, 9, 24), f"={A(h + 1, 31)}=0", font=GOOD)

    print("panel: data quality score", flush=True)
    # ---- panel: data quality score, before vs after (four gauges)
    panel(ws, 12, 25, 4, 17)
    put(ws, "D13:Q13", "Data quality score: before → after", color=IVORY, size=13, name=SERIF, v="c")
    put(ws, "D14:Q14", "Inner ring = the rows as they arrived (1 per source row) · outer ring = the clean list (1 per customer)",
        color=MUTED, size=8.5, v="t")
    left0, width = ws.Range("D15").Left, ws.Range("D15:Q15").Width
    slot = width / 4
    d = 122
    top = ws.Range("D15").Top + 2
    for k in range(4):
        r = 21 + k
        x = left0 + slot * k + (slot - d) / 2
        co = ws.ChartObjects().Add(x, top, d, d)
        co.Name = f"gauge_{k + 1}"
        ch = co.Chart
        ch.ChartType = -4120  # doughnut
        for vals, on in ((R(ws, r, h + 1, r, h + 2), COPPER), (R(ws, r, h + 3, r, h + 4), JADE if k != 3 else GOLD)):
            s = ch.SeriesCollection().NewSeries()
            s.Values = vals
            s.Format.Line.Visible = False
            s.Points(1).Format.Fill.ForeColor.RGB = on
            s.Points(2).Format.Fill.ForeColor.RGB = LINE
        dark_chart(ch)
        ch.ChartGroups(1).DoughnutHoleSize = 58
        ch.ChartGroups(1).FirstSliceAngle = 0
        ch.HasTitle = False
        cx = x + d / 2
        linked_box(ws, cx - 34, top + d / 2 - 16, 68, 30, f"=Summary!{A(h + 5, r)}", 18, IVORY, name=SERIF)
        linked_box(ws, left0 + slot * k + 4, ws.Range("D23").Top, slot - 8, 18, f"=Summary!{A(h, r)}", 10.5, GOLD_LT, bold=True)
        linked_box(ws, left0 + slot * k + 4, ws.Range("D24").Top, slot - 8, 16, f"=Summary!{A(h + 6, r)}", 8.5, MUTED)

    print("panel: funnel", flush=True)
    # ---- panel: funnel rows in -> unique customers
    panel(ws, 12, 25, 19, 27)
    put(ws, "S13:AA13", "From rows to customers", color=IVORY, size=13, name=SERIF, v="c")
    put(ws, "S14:AA14", '="removed "&TEXT(' + setaside + ',"#,##0")&" rows, merged "&TEXT(' + merged_ + ',"#,##0")&" duplicates"',
        color=MUTED, size=8.5, v="t")
    rr = ws.Range("S15:AA24")
    co = ws.ChartObjects().Add(rr.Left + 2, rr.Top, rr.Width - 4, rr.Height)
    co.Name = "funnel"
    ch = co.Chart
    ch.ChartType = 58  # stacked bar
    s0 = ch.SeriesCollection().NewSeries()
    s0.Values = R(ws, 17, h + 2, 19, h + 2)
    s0.XValues = R(ws, 17, h, 19, h)
    s0.Format.Fill.Visible = False
    s0.Format.Line.Visible = False
    s1 = ch.SeriesCollection().NewSeries()
    s1.Values = R(ws, 17, h + 1, 19, h + 1)
    s1.XValues = R(ws, 17, h, 19, h)
    for i, col in enumerate((GOLD, SKY, JADE), 1):
        s1.Points(i).Format.Fill.ForeColor.RGB = col
    s1.Format.Line.Visible = False
    s1.HasDataLabels = True
    s1.DataLabels().Position = 3  # centre
    s1.DataLabels().NumberFormat = "#,##0"
    s1.DataLabels().Font.Color = BG
    s1.DataLabels().Font.Bold = True
    s1.DataLabels().Font.Size = 11
    ch.ChartGroups(1).GapWidth = 35
    dark_chart(ch)
    ch.Axes(1).ReversePlotOrder = True
    ch.Axes(1).Format.Line.Visible = False
    ch.Axes(1).TickLabels.Font.Color = IVORY
    ch.Axes(1).TickLabels.Font.Size = 8.5
    ch.Axes(2).HasMajorGridlines = False
    ch.Axes(2).Delete()

    print("panel: field completeness heatmap", flush=True)
    # ---- panel: field completeness heatmap
    panel(ws, 27, 40, 4, 12)
    put(ws, "D28:L28", "Field completeness by source", color=IVORY, size=13, name=SERIF, v="c")
    put(ws, "D29:L29", "share of kept rows with the field filled · last column = the merged list", color=MUTED, size=8.5, v="t")
    put(ws, "D30:G30", "Field", color=GOLD_LT, size=8.5, bold=True, v="c")
    for c, t in zip((9, 10, 11, 12), ("Webshop", "POS", "Sales", "Customers")):
        put(ws, A(c, 30), t, color=GOLD_LT, size=8.5, bold=True, h="c", v="c")
        R(ws, 30, c).ShrinkToFit = True
    fields = [("Email", "Email", "Email"), ("Phone", "Phone", "Phone"), ("Company", "Company", "Company"),
              ("Address", "Address", "Address"), ("City", "City", "City"), ("State", "State", "State"), ("ZIP", "ZIP", "ZIP"),
              ("Marketing opt-in", "Opt-in", "Marketing opt-in")]
    for i, (label, rec_col, cust_col) in enumerate(fields, 31):
        put(ws, f"D{i}:G{i}", label, color=IVORY, size=9, v="c")
        for c, src in zip((9, 10, 11), ("Webshop", "POS", "Sales sheet")):
            ws.Cells(i, c).Formula = (f'=COUNTIFS(AllRecords[Source],"{src}",AllRecords[Status],"<>Discarded*",AllRecords[{rec_col}],"<>")'
                                      f'/COUNTIFS(AllRecords[Source],"{src}",AllRecords[Status],"<>Discarded*")')
        crit = '"<>Unknown"' if cust_col == "Marketing opt-in" else '"<>"'
        ws.Cells(i, 12).Formula = f"=COUNTIFS(Customers[{cust_col}],{crit})/ROWS(Customers[Customer ID])"
    heat = ws.Range("I31:L38")
    fmt(heat, color=IVORY, size=9, bold=True, h="c", v="c", numfmt="0%")
    cs = heat.FormatConditions.AddColorScale(3)
    for i, (val, col) in enumerate(((0, hx("6E3B28")), (0.5, hx("3A4450")), (1, hx("2F7F6A"))), 1):
        cs.ColorScaleCriteria(i).Type = 0
        cs.ColorScaleCriteria(i).Value = val
        cs.ColorScaleCriteria(i).FormatColor.Color = col
    for e in (7, 8, 9, 10, 11, 12):
        heat.Borders(e).LineStyle = 1
        heat.Borders(e).Color = CARD
        heat.Borders(e).Weight = 4 if e in (7, 8, 9, 10) else 2
    ws.Range("L30:L38").Borders(7).Color = GOLD
    put(ws, "D39:L39", "POS never asks for a company or opt-in; the merge fills them from the other two sources.",
        color=MUTED, size=8, italic=True, v="c")

    print("panel: merged by rule", flush=True)
    # ---- panel: merged by rule
    panel(ws, 27, 40, 14, 17)
    put(ws, "N28:Q28", "Merged by rule", color=IVORY, size=13, name=SERIF, v="c")
    put(ws, "N29:Q29", "what linked each duplicate", color=MUTED, size=8.5, v="t")
    rr = ws.Range("N30:Q39")
    co = ws.ChartObjects().Add(rr.Left, rr.Top, rr.Width, rr.Height)
    co.Name = "rules"
    ch = co.Chart
    ch.ChartType = 57
    s = ch.SeriesCollection().NewSeries()
    s.Values = R(ws, 12, h + 1, 15, h + 1)
    s.XValues = R(ws, 12, h, 15, h)
    s.Format.Fill.ForeColor.RGB = GOLD
    s.HasDataLabels = True
    s.DataLabels().Font.Color = GOLD_LT
    s.DataLabels().Font.Size = 9
    ch.ChartGroups(1).GapWidth = 45
    dark_chart(ch)
    ch.Axes(1).ReversePlotOrder = True
    ch.Axes(1).Format.Line.ForeColor.RGB = LINE
    ch.Axes(1).TickLabels.Font.Color = IVORY
    ch.Axes(1).TickLabels.Font.Size = 8
    ch.Axes(2).HasMajorGridlines = False
    ch.Axes(2).Delete()

    print("panel: data checks", flush=True)
    # ---- panel: data checks
    panel(ws, 27, 40, 19, 27)
    put(ws, "S28:AA28", "Data checks", color=IVORY, size=13, name=SERIF, v="c")
    put(ws, "S29:AA29", "recalculated after every refresh", color=MUTED, size=8.5, v="t")
    checks = [
        ("No row lost: received = set aside + customers + merged",
         '=IF(COUNTA(AllRecords[Record])=COUNTIF(AllRecords[Status],"Discarded*")+ROWS(Customers[Customer ID])+COUNTIF(AllRecords[Status],"Merged"),"PASS","FAIL")'),
        ("Every source row listed once on All records",
         '=IF(AND(SUMPRODUCT(1/COUNTIF(AllRecords[Record],AllRecords[Record]))=COUNTA(AllRecords[Record]),COUNTA(AllRecords[Record])=SUMIFS(CleaningLog[Rows],CleaningLog[Check],"Rows received")),"PASS","FAIL")'),
        ("No email belongs to two customers",
         '=IF(SUMPRODUCT((Customers[Email]<>"")/COUNTIF(Customers[Email],Customers[Email]&""))=COUNTIF(Customers[Email],"?*"),"PASS","FAIL")'),
        ("No phone belongs to two customers",
         '=IF(SUMPRODUCT((Customers[Phone]<>"")/COUNTIF(Customers[Phone],Customers[Phone]&""))=COUNTIF(Customers[Phone],"?*"),"PASS","FAIL")'),
        ("Every merged record points to a customer; counts add up",
         '=IF(AND(SUMPRODUCT(--(COUNTIF(Customers[Customer ID],MergeLog[Merged into])=0))=0,SUM(Customers[Records merged])=ROWS(Customers[Customer ID])+ROWS(MergeLog[Record])),"PASS","FAIL")'),
        ("Lifetime spend = spend on the kept records",
         '=IF(ROUND(SUM(Customers[Lifetime spend])-SUMIFS(AllRecords[Spend],AllRecords[Status],"Master")-SUMIFS(AllRecords[Spend],AllRecords[Status],"Merged"),2)=0,"PASS","FAIL")'),
        ("CRM import has exactly one row per customer",
         '=IF(ROWS(CRMImport[Email])=ROWS(Customers[Customer ID]),"PASS","FAIL")'),
        ("Every pair marked Merge is now one customer",
         '=IF(SUMPRODUCT((Review[Decision]="Merge")*(Review[Customer A]<>Review[Customer B]))=0,"PASS","FAIL")'),
        ("Quality after ≥ before on all four measures",
         '=IF(SUMPRODUCT(--(Quality[After]>=Quality[Before]))=ROWS(Quality[Metric]),"PASS","FAIL")'),
    ]
    for i, (label, f) in enumerate(checks, 30):
        put(ws, f"S{i}:Z{i}", label, color=IVORY, size=9, v="c")
        R(ws, i, 19).IndentLevel = 0
        cell = ws.Range(f"AA{i}")
        cell.Formula = f
        fmt(cell, size=9, bold=True, h="c", v="c")
        pass_fail(cell)
        for e in (7, 8, 9, 10):
            cell.Borders(e).LineStyle = 1
            cell.Borders(e).Color = CARD
            cell.Borders(e).Weight = 4
    wb.Names.Add("DataChecks", f"=Summary!$AA$30:$AA${29 + len(checks)}")
    put(ws, "S39:AA39", f'=COUNTIF(AA30:AA{29 + len(checks)},"PASS")&" of {len(checks)} checks pass"', color=GOOD, size=8.5,
        bold=True, h="r", v="c")

    print("panel: one customer", flush=True)
    # ---- panel: one customer, before and after
    panel(ws, 42, 52, 4, 27)
    put(ws, "D43:U43", "One customer, before and after", color=IVORY, size=13, name=SERIF, v="c")
    put(ws, "D44:Y44", "Every source row exactly as it arrived, then the golden record the merge produced. "
                       "Copper = a value the cleanup changed or a newer record replaced.", color=MUTED, size=8.5, v="t")
    put(ws, "V43:Y43", "Customer  ▸", color=MUTED, size=8.5, h="r", v="c")
    pick = put(ws, "Z43:AA43", None, color=GOLD_LT, size=10, bold=True, h="c", v="c", fill=CARD2)
    for e in (7, 8, 9, 10):
        pick.Borders(e).LineStyle = 1
        pick.Borders(e).Color = GOLD
    pick.Cells(1, 1).Value = "C00001"  # replaced by choose_examples
    pick.Validation.Delete()
    pick.Validation.Add(3, 1, 1, "=CustomerIDs")
    spans = [("Record", 4, 5), ("Source", 6, 7), ("Name as typed", 8, 11), ("Email", 12, 15), ("Phone", 16, 18),
             ("Address", 19, 21), ("City · State · ZIP", 22, 25), ("Last activity", 26, 27)]
    for t, c1, c2 in spans:
        put(ws, A(c1, 45) + ":" + A(c2, 45), t, color=GOLD_LT, size=8.5, bold=True, v="c")
    R(ws, 45, 4, 45, 27).Borders(9).LineStyle = 1
    R(ws, 45, 4, 45, 27).Borders(9).Color = LINE
    spill = "Calc!$B$3#"
    XR = lambda r, col: f'XLOOKUP($D{r},AllRecords[Record],AllRecords[{col}])'
    for k in range(1, 5):
        r = 45 + k
        cols = [f'=IFERROR(INDEX({spill},{k})&"","")', f'={XR(r, "Source")}&""', f'={XR(r, "Name as typed")}&""',
                f'={XR(r, "Email as received")}&""', f'={XR(r, "Phone as received")}&""', f'={XR(r, "Address as received")}&""',
                f'=TRIM({XR(r, "City as received")}&" "&{XR(r, "State as received")}&" "&{XR(r, "ZIP as received")})',
                f'={XR(r, "Last activity")}']
        for j, ((t, c1, c2), f) in enumerate(zip(spans, cols), 1):
            if j in (2, 8):
                f = f'=IF($D{r}="","",{f[1:]})'
            elif j > 2:
                f = f'=IF($D{r}="","",LET(v,{f[1:]},IF(v="","(blank)",v)))'
            put(ws, A(c1, r) + ":" + A(c2, r), f, color=IVORY, size=9, v="c", numfmt="yyyy-mm-dd" if j == 8 else None)
            R(ws, r, c1).ShrinkToFit = j in (2, 4, 5, 6, 7)
    r = 50
    gold_row = R(ws, r, 4, r, 27)
    gold_row.Interior.Color = CARD2
    gold_row.Borders(8).LineStyle = 1
    gold_row.Borders(8).Color = GOLD
    X = lambda col: f'XLOOKUP(ExampleID,Customers[Customer ID],Customers[{col}],"")'
    golden = ['="Golden"', f'={X("Sources")}&""', f'={X("First name")}&" "&{X("Last name")}', f'={X("Email")}&""',
              f'={X("Phone")}&""', f'={X("Address")}&""', f'=TRIM({X("City")}&" "&{X("State")}&" "&{X("ZIP")})',
              f'={X("Last activity")}']
    for (t, c1, c2), f in zip(spans, golden):
        put(ws, A(c1, r) + ":" + A(c2, r), f, color=GOLD, size=9.5, bold=True, v="c",
            numfmt="yyyy-mm-dd" if t == "Last activity" else None)
        R(ws, r, c1).ShrinkToFit = True
    # flags: 1 = this source value differs from the golden record (exact compare, so case changes count)
    put(ws, A(h, 39), "Changed? (1 = differs from golden)", color=MUTED, size=8, merge=False)
    for k in range(4):
        rr_ = 46 + k
        for j, (t, c1, c2) in enumerate(spans[2:7]):
            cell, gold_cell = A(c1, rr_), A(c1, 50)
            ws.Cells(40 + k, h + j).Formula = f'=IF(OR({cell}="",{cell}="(blank)"),0,IF(EXACT({cell},{gold_cell}),0,1))'
            fc = cf_formula(R(ws, rr_, c1), f"={A(h + j, 40 + k).replace(A(h + j, 40 + k).rstrip('0123456789'), '$' + A(h + j, 40 + k).rstrip('0123456789') + '$')}=1",
                            font=COPPER, fill=T_COPPER, bold=True)
            cf_formula(R(ws, rr_, c1), f'={cell}="(blank)"', font=GREY)
    put(ws, "D51:AA51", '="Changed values: "&SUM(' + A(h, 40) + ":" + A(h + 4, 43) + ')&"   ·   blanks filled from another record: '
                        'shown as (blank) in the source row, filled in the golden record"', color=MUTED, size=8, italic=True, v="c")

    print("side menu", flush=True)
    # ---- side menu
    menu = R(ws, 6, 2, 52, 2)
    menu.Interior.Color = CARD
    for e in (7, 8, 9, 10):
        menu.Borders(e).LineStyle = 1
        menu.Borders(e).Color = LINE
    put(ws, "B7", "  NAVIGATE", color=GOLD, size=8, bold=True, v="c")
    left, w = ws.Range("B8").Left + 7, ws.Range("B8").Width - 14
    y = ws.Range("B8").Top + 4
    for sheet, label, glyph in PAGES:
        button(ws, left, y, w, 27, glyph, label, sheet, sheet == "Summary")
        y += 34
    put(ws, "B42", "  HOW IT WORKS", color=GOLD, size=8, bold=True, v="c")
    put(ws, "B43:B51", "1  Save the new exports in the data folder\n\n2  Data > Refresh All\n\n3  Decide the open pairs on "
                       "Needs review (Merge / Keep separate), Refresh again\n\n4  Upload the CRM import tab",
        color=MUTED, size=8.5, wrap=True, v="t")
    R(ws, 43, 2).IndentLevel = 1

    ps = ws.PageSetup
    ps.Orientation = 2
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$AB$53"


# ---------------------------------------------------------------- Match explorer
def build_match_explorer(wb):
    ws = wb.Worksheets("Match explorer")
    page_header(ws, "Match explorer", "pick a customer: every source record behind it and the rule that linked it", "Match explorer")
    widths = {1: 1.5, 2: 9, 3: 11, 4: 10, 5: 21, 6: 20, 7: 28, 8: 17, 9: 10, 10: 27, 11: 15, 12: 11, 13: 10, 14: 2}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    ws.Range("A1").Value = None
    put(ws, "B1", "Match explorer", color=IVORY, size=18, name=SERIF, v="c", merge=False)
    put(ws, "B2", ws.Range("A2").Value, color=GOLD, size=8.5, v="c", merge=False)
    ws.Range("A2").Value = None
    for r, hgt in {5: 10, 6: 26, 7: 10, 8: 15, 9: 22, 10: 15, 11: 22, 12: 22, 13: 10, 14: 24, 15: 30}.items():
        ws.Rows(r).RowHeight = hgt
    # the drop-down list: "C00012  ·  Anna Smith" for every customer (spill in column Z, out of sight)
    put(ws, "B6:C6", "PICK A CUSTOMER  ▸", color=GOLD, size=9, bold=True, v="c")
    pick = put(ws, "D6:G6", None, color=IVORY, size=12, bold=True, v="c", fill=CARD2)
    for e in (7, 8, 9, 10):
        pick.Borders(e).LineStyle = 1
        pick.Borders(e).Color = GOLD
        pick.Borders(e).Weight = 2
    pick.IndentLevel = 1
    pick.Cells(1, 1).Value = "C00001"  # replaced by choose_examples
    pick.Validation.Delete()
    pick.Validation.Add(3, 1, 1, "=PickList")
    pick.Validation.ShowError = False  # an ID typed on its own (C00012) works too
    put(ws, "H6:M6", '="Open the list ("&TEXT(ROWS(Customers[Customer ID]),"#,##0")&" customers) or type an ID such as C00012. Everything below is live: FILTER + XLOOKUP."',
        color=MUTED, size=8.5, italic=True, v="c")
    X = lambda col: f'XLOOKUP(MatchID,Customers[Customer ID],Customers[{col}],"")'
    card = [(8, [("CUSTOMER", "B:C", f'={X("Customer ID")}&""'), ("NAME", "D:E", f'={X("First name")}&" "&{X("Last name")}'),
                 ("EMAIL", "F:G", f'={X("Email")}&""'), ("PHONE", "H:H", f'={X("Phone")}&""'),
                 ("COMPANY", "I:J", f'={X("Company")}&""'), ("OPT-IN", "K:K", f'={X("Marketing opt-in")}&""'),
                 ("LIFETIME SPEND", "L:M", f'={X("Lifetime spend")}')]),
            (10, [("ADDRESS", "B:D", f'={X("Address")}&""'), ("CITY · STATE · ZIP", "E:F", f'=TRIM({X("City")}&" "&{X("State")}&" "&{X("ZIP")})'),
                  ("SOURCES", "G:G", f'={X("Sources")}&""'), ("RECORDS", "H:H", f'={X("Records merged")}'),
                  ("LAST ACTIVITY", "I:J", f'={X("Last activity")}'), ("MASTER RECORD", "K:K", f'={X("Master record")}&""'),
                  ("OTHER EMAILS", "L:M", f'={X("Other emails")}&""')])]
    panel(ws, 8, 11, 2, 13, fill=CARD2)
    for r, items in card:
        for label, cols, f in items:
            c1, c2 = cols.split(":")
            put(ws, f"{c1}{r}:{c2}{r}", label, color=MUTED, size=7.5, bold=True, v="b")
            put(ws, f"{c1}{r + 1}:{c2}{r + 1}", f, color=GOLD_LT if label in ("CUSTOMER", "NAME") else IVORY, size=10.5,
                bold=label in ("CUSTOMER", "NAME"), v="c")
            ws.Range(f"{c1}{r + 1}").ShrinkToFit = True
            ws.Range(f"{c1}{r}").IndentLevel = 1
            ws.Range(f"{c1}{r + 1}").IndentLevel = 1
    ws.Range("L9").NumberFormat = "#,##0.00"
    ws.Range("I11").NumberFormat = "yyyy-mm-dd"
    put(ws, "B12:M12", '=LET(id,MatchID,a,XLOOKUP(id,Review[Customer A],Review[Customer B]&"  ("&Review[Reason]&";  "&Review[Result]&")",""),'
                       'b,XLOOKUP(id,Review[Customer B],Review[Customer A]&"  ("&Review[Reason]&";  "&Review[Result]&")",""),'
                       'IF(a&b="","No look-alike of this customer is waiting on Needs review.","Look-alike on Needs review: "&IF(a<>"",a,b)))',
        color=AMBER, size=9, italic=True, v="c")
    ws.Range("B12").IndentLevel = 1
    put(ws, "B14:M14", "Source records behind this customer, and the rule that linked each one", color=IVORY, size=13,
        name=SERIF, v="c")
    heads = ["Record", "Source", "Status", "Linked by", "Name as typed", "Email as received", "Phone as received",
             "ZIP as received", "Email (clean)", "Phone (clean)", "Last activity", "Spend"]
    for j, t in enumerate(heads, 2):
        c = ws.Cells(15, j)
        c.Value = t
    fmt(ws.Range("B15:M15"), fill=CARD2, color=GOLD_LT, bold=True, size=9, wrap=True, v="c")
    ws.Range("B15:M15").Borders(9).LineStyle = 1
    ws.Range("B15:M15").Borders(9).Color = GOLD
    for k in range(1, 9):
        ws.Range(f"B{15 + k}").Formula2 = (f'=IF(AND(MatchID<>"",Calc!$D$3=""),IF({k}=1,"No customer with this ID",""),'
                                            f'IFERROR(INDEX(Calc!$D$3#,{k})&"",""))')
    XR = lambda r, col: f'XLOOKUP($B{r},AllRecords[Record],AllRecords[{col}])'
    for r in range(16, 24):
        cols = [f'={XR(r, "Source")}&""', f'={XR(r, "Status")}&""',
                f'=IF({XR(r, "Status")}="Master","Master (newest record)",{XR(r, "Matched by")}&"")',
                f'={XR(r, "Name as typed")}&""', f'={XR(r, "Email as received")}&""', f'={XR(r, "Phone as received")}&""',
                f'={XR(r, "ZIP as received")}&""', f'={XR(r, "Email")}&""', f'={XR(r, "Phone")}&""',
                f'={XR(r, "Last activity")}', f'={XR(r, "Spend")}']
        for c, f in enumerate(cols, 3):
            ws.Cells(r, c).Formula2 = f'=IF(OR($B{r}="",$B{r}="No customer with this ID"),"",{f[1:]})'
    body = ws.Range("B16:M23")
    fmt(body, size=9.5, v="c")
    for c in ("G", "J"):
        ws.Range(f"{c}16:{c}23").ShrinkToFit = True
    ws.Range("L16:L23").NumberFormat = "yyyy-mm-dd"
    ws.Range("M16:M23").NumberFormat = "#,##0.00"
    for r in range(16, 24):
        ws.Rows(r).RowHeight = 20
    cf_formula(body, '=$D16="Master"', font=GOLD_LT, fill=T_GOOD, bold=True)
    cf_formula(body, '=$B16<>""', fill=CARD)
    for text, font in (("Same email", GOLD_LT), ("Same phone", SKY), ("Same name and ZIP", JADE), ("Reviewer decision", AMBER)):
        cf_formula(ws.Range("E16:E23"), f'=$E16="{text}"', font=font, bold=True)
    put(ws, "B25:M25", "How records are linked", color=GOLD, size=9, bold=True, v="c")
    legend = ["Same email: the cleaned emails are identical (lower case, spaces removed).",
              "Same phone: the same 10 digits, whatever way they were typed ((303) 555-0142, 303.555.0142, +1 303-555-0142 ...).",
              "Same name and ZIP: first name + last name + ZIP agree; nicknames count (Bill = William, see the Name rules tab).",
              "Reviewer decision: a person marked the pair Merge on Needs review; it is applied on the next Refresh All.",
              "Links chain: A = B by email and B = C by phone puts all three in one customer. The newest record is the master."]
    for i, t in enumerate(legend, 26):
        put(ws, f"B{i}:M{i}", t, color=MUTED, size=8.5, v="c")
        ws.Range(f"B{i}").IndentLevel = 1
    ws.Activate()
    wb.Application.ActiveWindow.FreezePanes = False


def build_calc(wb):
    """Hidden helper sheet with the three spilled lists (a multi-row spill on the cockpit sheet crashes Excel)."""
    ws = wb.Worksheets("Calc")
    wb.Worksheets("Match explorer").Range("D6").Value = "C00001"  # replaced by choose_examples
    wb.Names.Add("MatchID", "=Calc!$H$3")
    ws.Range("A1").Value = "Helper lists for the cockpit and Match explorer (live formulas; the sheet is hidden)"
    ws.Range("B2:H2").Value = ("Example customer: records", None, "Match explorer: records", None, "Pick list", None, "Picked ID")
    ws.Range("H3").Formula2 = "=TEXTBEFORE(TRIM('Match explorer'!$D$6)&\" \",\" \")"
    ws.Range("B2:H2").Font.Bold = True
    ws.Range("B3").Formula2 = RECLIST.format(id="ExampleID")
    ws.Range("D3").Formula2 = RECLIST.format(id="MatchID")
    ws.Range("F3").Formula2 = '=Customers[Customer ID]&"  ·  "&Customers[First name]&" "&Customers[Last name]'
    wb.Names.Add("PickList", "=Calc!$F$3#")
    ws.Columns("B:F").ColumnWidth = 24
    ws.Visible = 0  # hidden


def choose_examples(wb):
    """Pick the customers shown on the cockpit card and on Match explorer from the refreshed workbook."""
    lo = wb.Worksheets("All records").ListObjects("AllRecords")
    cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
    rows = [dict(zip(cols, r)) for r in lo.DataBodyRange.Value]
    by = {}
    for r in rows:
        if r["Customer ID"]:
            by.setdefault(r["Customer ID"], []).append(r)

    def changed(r):
        return sum(1 for raw, clean in ((r["Email as received"], r["Email"]), (r["Phone as received"], r["Phone"]),
                                        (r["ZIP as received"], r["ZIP"])) if raw and raw != clean)

    def score(cid):
        g = by[cid]
        rules = {r["Matched by"] for r in g if r["Matched by"]}
        return (len(g), len({r["Source"] for r in g}), len(rules), sum(changed(r) for r in g))

    ranked = sorted(by, key=lambda c: (score(c), c), reverse=True)
    example = ranked[0]
    explorer = next((c for c in ranked if c != example and len({r["Matched by"] for r in by[c] if r["Matched by"]}) >= 2), ranked[1])
    wb.Worksheets("Summary").Range("Z43").Value = example
    lo = wb.Worksheets("Customers").ListObjects("Customers")
    names = {r[0]: f"{r[1]} {r[2]}" for r in lo.DataBodyRange.Value}
    name = names.get(explorer, "")
    wb.Worksheets("Match explorer").Range("D6").Value = f"{explorer}  ·  {name}"
    print("example", example, "explorer", explorer)


def build_settings(wb):
    st = wb.Worksheets("Settings")
    st.Range("A1").Value = "Settings"
    st.Range("A1").Font.Size = 16; st.Range("A1").Font.Bold = True
    rows = [("Data folder", '=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"data\\"', "DataFolder",
             "Folder with the three exports. Default: the data folder next to this workbook. Type a path to override."),
            ("Webshop file", "webshop-customers-2026-10.csv", "WebshopFile", "Online shop customer export (CSV)"),
            ("POS file", "pos-customers-2026-10.csv", "PosFile", "In-store till customer export (CSV)"),
            ("Sales sheet file", "sales-team-accounts.xlsx", "SalesFile", "The sales team's spreadsheet, sheet 'Trade accounts' (.xlsx)"),
            ("Period", "October 2026", "Period", "Shown in the report title")]
    st.Range("A3:C3").Value = ("Setting", "Value", "Note")
    st.Range("A3:C3").Font.Bold = True
    for r, (label, val, nm, note) in enumerate(rows, 4):
        st.Cells(r, 1).Value = label
        if val.startswith("="):
            st.Cells(r, 2).Formula2 = val
        else:
            st.Cells(r, 2).NumberFormat = "@"
            st.Cells(r, 2).Value = val
        st.Cells(r, 3).Value = note
        st.Cells(r, 2).Interior.Color = rgb(*YELLOW)
        wb.Names.Add(nm, f"=Settings!$B${r}")
    st.Range("A11").Value = "Decision choices"
    st.Range("A11").Font.Bold = True
    st.Range("B11").Value = "Merge"
    st.Range("B12").Value = "Keep separate"
    st.Range("C11").Value = "The drop-down in the Decision column on Needs review"
    wb.Names.Add("DecisionList", "=Settings!$B$11:$B$12")
    st.Columns("A").ColumnWidth = 18; st.Columns("B").ColumnWidth = 48; st.Columns("C").ColumnWidth = 90


def build_name_rules(wb):
    ws = wb.Worksheets("Name rules")
    ws.Range("A1").Value = "Name rules: nicknames that mean the same first name"
    ws.Range("A1").Font.Size = 14; ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = ("Used only for matching (Bill Ortiz + same ZIP = William Ortiz). The name on the customer stays as typed. "
                            "Add a row for any nickname your customers use, then Refresh All.")
    ws.Range("A2").Font.Italic = True
    ws.Range("A4:B4").Value = ("Nickname", "Full name")
    for i, (n, f) in enumerate(NICK, 5):
        ws.Cells(i, 1).Value = n; ws.Cells(i, 2).Value = f
    lo = ws.ListObjects.Add(1, ws.Range(f"A4:B{4 + len(NICK)}"), None, 1)
    lo.Name = "Nicknames"
    lo.TableStyle = "TableStyleLight9"
    ws.Range(f"A5:B{4 + len(NICK)}").Interior.Color = rgb(*YELLOW)
    ws.Columns("A").ColumnWidth = 14; ws.Columns("B").ColumnWidth = 16


def build_howto(wb):
    ws = wb.Worksheets("How to use")
    ws.Columns("A").ColumnWidth = 125
    text = [
        ("How to use this workbook", True), ("", False),
        ("Every time", True),
        ("1. Save the webshop and POS customer exports (CSV) and the sales team's spreadsheet into the data folder (see Settings).", False),
        ("2. On Settings, type the three file names.", False),
        ("3. Data > Refresh All. Every tab and the cockpit update. No macros.", False),
        ("4. Cockpit (Summary): all nine data checks must say PASS. The gauges show data quality before and after the cleanup.", False),
        ("5. Needs review: type Merge or Keep separate in the Decision column (drop-down), then Refresh All again. A Merge joins the two customers;", False),
        ("   the decision stays on the sheet and is applied on every later refresh. Blank = still open.", False),
        ("6. Match explorer: pick any customer to see every source record behind it and the rule that linked it.", False),
        ("7. CRM import tab: File > Save a Copy > CSV UTF-8 (with this tab active) and upload it to HubSpot, Pipedrive or Mailchimp.", False), ("", False),
        ("How duplicates are found", True),
        ("Two records are the same customer when they share an email, a phone number, or first name + last name + ZIP.", False),
        ("Matching works on cleaned values: emails in lower case, phones as 10 digits, ZIPs as 5 digits, nicknames as the full name (Name rules tab).", False),
        ("Links chain: A = B by email and B = C by phone puts A, B and C in one customer. Every merge is listed on the Merge log with its rule.", False), ("", False),
        ("Which values are kept", True),
        ("The most recent record is the master. Every field comes from the most recent record that has it, so blanks are filled from the duplicates.", False),
        ("Other email addresses are kept in 'Other emails'. Lifetime spend adds up all merged records.", False),
        ("Marketing opt-in: if any system says No, the customer is No. An opt-out is never overwritten by an older Yes.", False), ("", False),
        ("What is not merged automatically", True),
        ("Look-alikes (Christopher and Christine at the same ZIP; the same name in two cities) go to Needs review with what to check.", False),
        ("Test accounts, walk-in rows and rows the export repeated are removed and listed on All records, so nothing disappears silently.", False),
        ("The steps live in Power Query (Data > Queries & Connections): Records, Keyed, CustomersAll, MergeLog, Review, Quality.", False),
    ]
    for i, (t, bold) in enumerate(text, 1):
        ws.Cells(i, 1).Value = t
        ws.Cells(i, 1).Font.Bold = bold
    ws.Range("A1").Font.Size = 16


def export_crm_csv(wb):
    lo = wb.Worksheets("CRM import").ListObjects(1)
    cols = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
    (ROOT / "output").mkdir(exist_ok=True)
    out = ROOT / "output" / "crm-import-2026-10.csv"
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(cols)
        for r in lo.DataBodyRange.Value:
            w.writerow(["" if v is None else f"{v:.2f}" if isinstance(v, float) else v for v in r])
    print("wrote", out.relative_to(ROOT))


def finish(xl, wb):
    """Cockpit, Match explorer, How to use, examples, CRM file; then save."""
    xl.Calculation = -4135  # manual while the layout is built
    build_summary(wb)
    print("summary built", flush=True)
    build_calc(wb)
    build_match_explorer(wb)
    print("match explorer built", flush=True)
    build_howto(wb)
    xl.Calculation = -4105
    xl.Calculate()
    choose_examples(wb)
    xl.Calculate()
    export_crm_csv(wb)
    for ws in wb.Worksheets:
        ws.Activate()
        xl.ActiveWindow.DisplayGridlines = False
        xl.ActiveWindow.ScrollRow = 1
        xl.ActiveWindow.ScrollColumn = 1
        if ws.Name in ("Summary", "Match explorer"):
            ws.Range("A1").Select()
    wb.Worksheets("Summary").Activate()
    wb.Save()
    print("saved", OUT.name)


def relayout(xl):
    """--layout: rebuild only the cockpit and Match explorer on the saved workbook (queries untouched)."""
    wb = xl.Workbooks.Open(str(OUT))
    if "Calc" not in [w.Name for w in wb.Worksheets]:
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)).Name = "Calc"
    for name in ("Summary", "Match explorer", "Calc"):
        ws = wb.Worksheets(name)
        for i in range(ws.Shapes.Count, 0, -1):
            ws.Shapes(i).Delete()
        ws.Cells.Clear()
        ws.Cells.FormatConditions.Delete()
        ws.Cells.Validation.Delete()
        ws.Columns.Hidden = False
        ws.Rows.RowHeight = 15
    for nm in ("ExampleID", "MatchID", "DataChecks", "PickList"):
        try:
            wb.Names(nm).Delete()
        except Exception:
            pass
    return wb


def main():
    import sys
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        if "--layout" in sys.argv:
            wb = relayout(xl)
            finish(xl, wb)
            return
        if OUT.exists():
            OUT.unlink()
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < len(LOADED):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, (title, _) in enumerate(LOADED, 1):
            wb.Worksheets(i).Name = title
        st = wb.Styles("Normal")
        st.Font.Name = SANS
        st.Font.Size = 10
        build_settings(wb)
        build_name_rules(wb)
        wb.SaveAs(str(OUT), 51)
        xl.Calculate()
        try:
            wb.Queries.FastCombine = True  # queries read files and tables of this workbook: skip the privacy-level prompt
        except Exception as e:
            print("FastCombine not set:", e)
        for name, formula in Q.items():
            wb.Queries.Add(name, formula)
        t0 = time.time()
        for title, query in LOADED:
            if query:
                t1 = time.time()
                load_query(wb, wb.Worksheets(title), query, query.replace("_", ""))
                print(f"  {query} {time.time() - t1:.1f}s")
        print(f"queries loaded in {time.time() - t0:.1f}s")
        wb.Names.Add("CustomerIDs", "=Customers[Customer ID]")
        format_tables(wb)
        # the shipped workbook has the six similar-first-name pairs decided (different email and phone: two people);
        # the five same-name-different-ZIP pairs stay open for the client
        lo = wb.Worksheets("Needs review").ListObjects("Review")
        for i in range(1, lo.ListRows.Count + 1):
            row = lo.ListRows(i).Range
            if str(row.Cells(1, 3).Value).startswith("Similar"):
                row.Cells(1, 1).Value = "Keep separate"
        t1 = time.time()
        wb.RefreshAll()
        xl.CalculateUntilAsyncQueriesDone()
        print(f"refresh with decisions {time.time() - t1:.1f}s")
        wb.Save()
        # lay out the cockpit on the reopened file: a spilled formula on a sheet that existed while the queries were
        # first loaded crashes this Excel build, so the hidden Calc sheet with the spills is added after reopening
        wb.Close(False)
        wb = relayout(xl)
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
