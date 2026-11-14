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
        Up      = Text.Upper(Typed)
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
          fAddr = Address <> null and T(r[Address raw]) <> Address, fCo = r[Company filled]],
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
                 if n >= 50 or List.Sum(s[G]) = List.Sum(t[G]) then s else @Run(s, n + 1)
in
    Run(Table.Buffer(Start), 1)"""

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
              Why  = Table.AddColumn(Rest, "Matched by", each
                         if [Email] <> null and Cnt("Email", [Email]) > 1 then "Same email"
                         else if [Phone] <> null and Cnt("Phone", [Phone]) > 1 then "Same phone"
                         else "Same name and ZIP", type text),
              Into = Table.AddColumn(Table.AddColumn(Table.AddColumn(Why, "Merged into", each c[Customer ID], type text),
                         "Customer", each c[First name] & " " & c[Last name], type text), "Master record", each c[Master record], type text),
              Note = Table.AddColumn(Into, "Note", each
                         if [Email] <> null and [Email] <> c[Email] then "Different email: kept under Other emails"
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
    // look-alikes in different groups: similar first name with the same last name and ZIP, or the same name in another ZIP
    Pairs  = (k as text, reason as text) as table => let
                 rows  = Table.SelectRows(Keyed, each Record.Field(_, k) <> null),
                 g     = Table.Group(rows, {k}, {{"Gs", each List.Sort(List.Distinct([G]))}}),
                 multi = Table.SelectRows(g, each List.Count([Gs]) > 1),
                 recs  = List.Combine(List.Transform(multi[Gs], (gs) => List.Combine(List.Transform(List.Positions(gs),
                             (i) => List.Transform(List.Skip(gs, i + 1), (b) => [A = gs{i}, B = b, Reason = reason])))))
             in Table.FromRecords(recs, type table [A = Int64.Type, B = Int64.Type, Reason = text]),
    Both   = Table.Distinct(Table.Combine({Pairs("RA", "Similar first name, same last name and ZIP"),
                                           Pairs("RB", "Same name, different ZIP")}), {"A", "B"}),
    C      = Table.SelectColumns(CustomersAll, {"G", "Customer ID", "First name", "Last name", "Email", "Phone", "City", "ZIP", "Master record"}),
    JA     = Table.ExpandTableColumn(Table.NestedJoin(Both, {"A"}, C, {"G"}, "a", JoinKind.LeftOuter), "a",
                 {"Customer ID", "First name", "Last name", "Email", "Phone", "City", "ZIP", "Master record"},
                 {"Customer A", "FA", "LA", "Email A", "Phone A", "City A", "ZIP A", "Master A"}),
    JB     = Table.ExpandTableColumn(Table.NestedJoin(JA, {"B"}, C, {"G"}, "b", JoinKind.LeftOuter), "b",
                 {"Customer ID", "First name", "Last name", "Email", "Phone", "City", "ZIP", "Master record"},
                 {"Customer B", "FB", "LB", "Email B", "Phone B", "City B", "ZIP B", "Master B"}),
    Names  = Table.AddColumn(Table.AddColumn(JB, "Name A", each [FA] & " " & [LA], type text), "Name B", each [FB] & " " & [LB], type text),
    Check  = Table.AddColumn(Names, "What to check", each
                 if Text.StartsWith([Reason], "Similar") then [FA] & " / " & [FB] & ": different email and phone, probably two people. Merge only if confirmed."
                 else [City A] & " / " & [City B] & ": two people, or one who moved? Check order history first.", type text),
    Out    = Table.SelectColumns(Check, {"Reason", "Customer A", "Name A", "Email A", "Phone A", "City A", "ZIP A",
                 "Customer B", "Name B", "Email B", "Phone B", "City B", "ZIP B", "What to check", "Master A", "Master B"})
in
    Table.Sort(Out, {{"Reason", Order.Ascending}, {"Customer A", Order.Ascending}})"""

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
                 "First", "Last", "Email", "Phone", "Company", "Address", "City", "State", "ZIP", "Opt-in", "Last activity", "Spend"})
in
    Table.RenameColumns(Out, {{"Result", "Status"}, {"Name typed", "Name as typed"}})"""

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

LOADED = [("Summary", None), ("Customers", "Customers"), ("Needs review", "Review"), ("Merge log", "MergeLog"),
          ("CRM import", "CRM_Import"), ("All records", "AllRecords"), ("Cleaning log", "Cleaning_Log"),
          ("Name rules", None), ("Settings", None), ("How to use", None)]
NAVY, TEAL = (31, 58, 90), (0, 128, 128)
GREEN = ((198, 239, 206), (0, 97, 0))
RED = ((255, 199, 206), (156, 0, 6))
YELLOW = (255, 255, 153)


def rgb(r, g, b):
    return r + g * 256 + b * 65536


def conn_string(name):
    return f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=""'


def load_query(wb, ws, name, table_name, at="A1"):
    lo = ws.ListObjects.Add(0, conn_string(name), None, 1, ws.Range(at))
    qt = lo.QueryTable
    qt.CommandType = 2
    qt.CommandText = f"SELECT * FROM [{name}]"
    qt.RefreshStyle = 1
    qt.AdjustColumnWidth = False
    qt.BackgroundQuery = False
    qt.Refresh(False)
    lo.Name = table_name
    lo.TableStyle = "TableStyleMedium2"
    return lo


def heading(ws, cell, text, size=13):
    ws.Range(cell).Value = text
    ws.Range(cell).Font.Size = size
    ws.Range(cell).Font.Bold = True
    ws.Range(cell).Font.Color = rgb(*NAVY)


def pass_fail(rng, ok_value='="PASS"'):
    fc = rng.FormatConditions.Add(1, 3, ok_value)
    fc.Interior.Color = rgb(*GREEN[0]); fc.Font.Color = rgb(*GREEN[1])
    fc = rng.FormatConditions.Add(1, 4, ok_value)
    fc.Interior.Color = rgb(*RED[0]); fc.Font.Color = rgb(*RED[1])


def title_band(ws, last_col, title_formula, sub):
    ws.Range(f"A1:{last_col}2").Interior.Color = rgb(*NAVY)
    ws.Range("B1").Formula = title_formula
    ws.Range("B1").Font.Size = 18; ws.Range("B1").Font.Bold = True; ws.Range("B1").Font.Color = rgb(255, 255, 255)
    ws.Range("B2").Value = sub
    ws.Range("B2").Font.Color = rgb(200, 215, 230)
    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 20


def hdr(cell, text, right=False):
    cell.Value = text; cell.Font.Bold = True; cell.Interior.Color = rgb(*NAVY); cell.Font.Color = rgb(255, 255, 255)
    if right:
        cell.HorizontalAlignment = -4152


# ---------------------------------------------------------------- Summary
def build_summary(wb):
    ws = wb.Worksheets("Summary")
    widths = {"A": 2, "B": 30, "C": 11, "D": 11, "E": 11, "F": 11, "G": 3, "H": 12, "I": 12, "J": 12, "K": 12, "L": 12, "M": 12}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    title_band(ws, "M", '="Customer list cleanup   |   Summit Ridge Outdoor   |   "&Period',
               "Webshop + POS + sales team spreadsheet, cleaned and de-duplicated in Power Query. Fictional data. New exports: Data > Refresh All.")

    heading(ws, "B11", "By source")
    for c, t in {"B": "Source", "C": "Received", "D": "Discarded", "E": "Customers", "F": "Merged"}.items():
        hdr(ws.Range(f"{c}12"), t, c != "B")
    for i, src in enumerate(["Webshop", "POS", "Sales sheet"], 13):
        ws.Range(f"B{i}").Value = src
        ws.Range(f"C{i}").Formula = f"=COUNTIFS(AllRecords[Source],B{i})"
        ws.Range(f"D{i}").Formula = f'=COUNTIFS(AllRecords[Source],B{i},AllRecords[Status],"Discarded*")'
        ws.Range(f"E{i}").Formula = f'=COUNTIFS(AllRecords[Source],B{i},AllRecords[Status],"Master")'
        ws.Range(f"F{i}").Formula = f'=COUNTIFS(AllRecords[Source],B{i},AllRecords[Status],"Merged")'
    ws.Range("B16").Value = "Total"
    for c in "CDEF":
        ws.Range(f"{c}16").Formula = f"=SUM({c}13:{c}15)"
    for c in "BCDEF":
        ws.Range(f"{c}16").Font.Bold = True
        ws.Range(f"{c}16").Borders(8).LineStyle = 1
    ws.Range("C13:F16").NumberFormat = "#,##0"
    ws.Range("B17").Value = "Customers = the source of the record kept as master; merged = folded into a customer."
    ws.Range("B17").Font.Italic = True; ws.Range("B17").Font.Size = 9; ws.Range("B17").Font.Color = rgb(110, 110, 110)

    # merged records by rule (feeds the chart), H11:I14 area kept small on the right of the chart data
    rules = ["Same email", "Same phone", "Same name and ZIP"]
    for i, rule in enumerate(rules, 13):
        ws.Range(f"L{i + 20}").Value = rule  # helper cells far below the print area
        ws.Range(f"M{i + 20}").Formula = f'=COUNTIFS(MergeLog[Matched by],L{i + 20})'
    ws.Range("L33:M35").Font.Color = rgb(255, 255, 255)

    # Data checks (rows 22-29)
    heading(ws, "B22", "Data checks")
    checks = [
        ("No row lost: received = discarded + customers + merged",
         '=IF(COUNTA(AllRecords[Record])=COUNTIF(AllRecords[Status],"Discarded*")+ROWS(Customers[Customer ID])+COUNTIF(AllRecords[Status],"Merged"),"PASS","FAIL")'),
        ("Every source row listed once on All records",
         '=IF(AND(SUMPRODUCT(1/COUNTIF(AllRecords[Record],AllRecords[Record]))=COUNTA(AllRecords[Record]),COUNTA(AllRecords[Record])=SUMIFS(CleaningLog[Rows],CleaningLog[Check],"Rows received")),"PASS","FAIL")'),
        ("No email belongs to two customers",
         '=IF(SUMPRODUCT((Customers[Email]<>"")/COUNTIF(Customers[Email],Customers[Email]&""))=COUNTIF(Customers[Email],"?*"),"PASS","FAIL")'),
        ("No phone belongs to two customers",
         '=IF(SUMPRODUCT((Customers[Phone]<>"")/COUNTIF(Customers[Phone],Customers[Phone]&""))=COUNTIF(Customers[Phone],"?*"),"PASS","FAIL")'),
        ("Every merged record points to a customer, and the counts add up",
         '=IF(AND(SUMPRODUCT(--(COUNTIF(Customers[Customer ID],MergeLog[Merged into])=0))=0,SUM(Customers[Records merged])=ROWS(Customers[Customer ID])+ROWS(MergeLog[Record])),"PASS","FAIL")'),
        ("Lifetime spend of the customers = spend on the kept records",
         '=IF(ROUND(SUM(Customers[Lifetime spend])-SUMIFS(AllRecords[Spend],AllRecords[Status],"Master")-SUMIFS(AllRecords[Spend],AllRecords[Status],"Merged"),2)=0,"PASS","FAIL")'),
        ("CRM import has exactly one row per customer",
         '=IF(ROWS(CRMImport[Email])=ROWS(Customers[Customer ID]),"PASS","FAIL")'),
    ]
    for i, (label, f) in enumerate(checks, 23):
        ws.Range(f"B{i}").Value = label
        cell = ws.Range(f"M{i}")
        cell.Formula = f
        cell.Font.Bold = True
        cell.HorizontalAlignment = -4108
        pass_fail(cell)

    cards = [("B", "B", "ROWS RECEIVED", "=C16", "#,##0", '="from 3 sources; "&D16&" discarded"'),
             ("C", "E", "UNIQUE CUSTOMERS", "=E16", "#,##0", '="one row each, ready for the CRM"'),
             ("F", "I", "DUPLICATES MERGED", "=F16", "#,##0", '="records folded into a customer"'),
             ("J", "M", "PAIRS TO REVIEW", '=ROWS(Review[Reason])', "0", '="look-alikes a person should decide"')]
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
            rng.WrapText = r1 == "8"
        ws.Range(f"{a}6").NumberFormat = fmt
    ws.Rows(4).RowHeight = 5

    top, left = ws.Range("H11").Top, ws.Range("H11").Left
    co = ws.ChartObjects().Add(left, top, ws.Range("H11:M11").Width, ws.Range("H11:H20").Height)
    ch = co.Chart
    ch.ChartType = 57
    ser = ch.SeriesCollection().NewSeries()
    ser.Name = "Records"
    ser.Values = ws.Range("M33:M35")
    ser.XValues = ws.Range("L33:L35")
    ch.HasLegend = False
    ser.Format.Fill.ForeColor.RGB = rgb(*TEAL)
    ser.HasDataLabels = True
    ch.Axes(1).ReversePlotOrder = True
    ch.Axes(2).HasMajorGridlines = False
    ch.Axes(2).Delete()
    ch.HasTitle = True
    ch.ChartTitle.Text = "Duplicates merged, by match rule"
    ch.ChartTitle.Font.Size = 13
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Font.Color = rgb(*NAVY)
    ch.ChartTitle.Left = 4
    ch.ChartArea.Format.Line.Visible = False

    ps = ws.PageSetup
    ps.Orientation = 2; ps.Zoom = False; ps.FitToPagesWide = 1; ps.FitToPagesTall = 1
    ps.PrintArea = "$A$1:$M$29"


def format_tables(wb):
    fmt = {"Lifetime spend": "#,##0.00", "Lifetime Spend": "#,##0.00", "Spend": "#,##0.00",
           "Last activity": "yyyy-mm-dd"}
    widths = {"Customers": {"Customer ID": 9, "First name": 11, "Last name": 11, "Email": 30, "Phone": 15, "Company": 24,
                            "Address": 22, "City": 14, "State": 6, "ZIP": 7, "Marketing opt-in": 9, "Lifetime spend": 11,
                            "Last activity": 11, "Sources": 24, "Records merged": 8, "Master record": 9, "Other emails": 34},
              "Needs review": {"Reason": 37, "Customer A": 9, "Name A": 18, "Email A": 28, "Phone A": 15, "City A": 13,
                               "ZIP A": 7, "Customer B": 9, "Name B": 18, "Email B": 28, "Phone B": 15, "City B": 13,
                               "ZIP B": 7, "What to check": 79, "Master A": 9, "Master B": 9},
              "Merge log": {"Record": 8, "Source": 11, "Name as typed": 22, "Email": 30, "Phone": 15, "ZIP": 7,
                            "Last activity": 11, "Merged into": 9, "Customer": 20, "Master record": 9, "Matched by": 17, "Note": 40},
              "CRM import": {"Email": 30, "Street Address": 22, "Company Name": 24, "Additional Emails": 34, "City": 14,
                             "Phone Number": 15},
              "All records": {"Status": 28, "Name as typed": 24, "Email": 30, "Address": 22, "Company": 22, "Phone": 15,
                              "Last activity": 11, "Matched by": 16}}
    for title, w in widths.items():
        ws = wb.Worksheets(title)
        lo = ws.ListObjects(1)
        for i in range(1, lo.ListColumns.Count + 1):
            name = lo.ListColumns(i).Name
            ws.Columns(i).ColumnWidth = w.get(name, 10)
            if name in fmt and lo.DataBodyRange is not None:
                lo.ListColumns(i).DataBodyRange.NumberFormat = fmt[name]
        lo.HeaderRowRange.WrapText = True
        ws.Activate()
        ws.Range("A2").Select()
        wb.Application.ActiveWindow.FreezePanes = True
    ws = wb.Worksheets("Cleaning log")
    for c, w in zip("ABCD", (12, 52, 7, 70)):
        ws.Columns(c).ColumnWidth = w
    # colour the rule on the merge log and the status on all records
    lo = wb.Worksheets("All records").ListObjects(1)
    rng = lo.ListColumns("Status").DataBodyRange
    for text, color in (("Master", (226, 239, 218)), ("Merged", (221, 235, 247)),
                        ("Discarded: exact duplicate row", (252, 228, 214)), ("Discarded: test or walk-in record", (252, 228, 214)),
                        ("Discarded: no name or contact", (252, 228, 214))):
        fc = rng.FormatConditions.Add(2, None, f'=$D2="{text}"')  # no commas: rule formulas use the local list separator
        fc.Interior.Color = rgb(*color)


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
    st.Columns("A").ColumnWidth = 18; st.Columns("B").ColumnWidth = 48; st.Columns("C").ColumnWidth = 90


def build_name_rules(wb):
    ws = wb.Worksheets("Name rules")
    ws.Range("A1").Value = "Name rules: nicknames that mean the same first name"
    ws.Range("A1").Font.Size = 14; ws.Range("A1").Font.Bold = True; ws.Range("A1").Font.Color = rgb(*NAVY)
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
        ("3. Data > Refresh All. Customers, Merge log, Needs review, CRM import and the Summary all update. No macros.", False),
        ("4. Summary: all seven data checks must say PASS. Then decide the pairs on Needs review.", False),
        ("5. CRM import tab: File > Save a Copy > CSV UTF-8 (with this tab active) and upload it to HubSpot, Pipedrive or Mailchimp.", False), ("", False),
        ("How duplicates are found", True),
        ("Two records are the same customer when they share an email, a phone number, or first name + last name + ZIP.", False),
        ("Matching works on cleaned values: emails in lower case, phones as 10 digits, ZIPs as 5 digits, nicknames as the full name (Name rules tab).", False),
        ("Links chain: A = B by email and B = C by phone puts A, B and C in one customer. Every merge is listed on the Merge log with its rule.", False), ("", False),
        ("Which values are kept", True),
        ("The most recent record is the master. Every field comes from the most recent record that has it, so blanks are filled from the duplicates.", False),
        ("Other email addresses are kept in 'Other emails'. Lifetime spend adds up all merged records.", False),
        ("Marketing opt-in: if any system says No, the customer is No. An opt-out is never overwritten by an older Yes.", False), ("", False),
        ("What is not merged automatically", True),
        ("Look-alikes (Christopher and Christine Hale at the same ZIP; the same name in two cities) go to Needs review with what to check.", False),
        ("Test accounts, walk-in rows and rows the export repeated are removed and listed on All records, so nothing disappears silently.", False),
        ("The steps live in Power Query (Data > Queries & Connections): Records, Keyed, CustomersAll, MergeLog, Review.", False),
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


def main():
    if OUT.exists():
        OUT.unlink()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < len(LOADED):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, (title, _) in enumerate(LOADED, 1):
            wb.Worksheets(i).Name = title
        build_settings(wb)
        build_name_rules(wb)
        wb.SaveAs(str(OUT), 51)
        xl.Calculate()
        try:
            wb.Queries.FastCombine = True  # one query reads files and the Name rules table: skip the privacy-level prompt
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
        format_tables(wb)
        build_summary(wb)
        build_howto(wb)
        xl.Calculate()
        export_crm_csv(wb)
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            ws.Range("A1").Select()
        wb.Worksheets("Summary").Activate()
        wb.Save()
        print("saved", OUT.name)
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        xl.Quit()


if __name__ == "__main__":
    main()
