"""Builds the Power BI project (PBIP) for E15 as plain text files that Power BI Desktop opens directly:

  Larkspur-Sales.pbip                  the file to double-click
  Larkspur-Sales.SemanticModel/        the model in TMDL: Power Query (M) that reads the webshop and marketplace exports
                                       and the live exchange rates (Frankfurter API), tables, relationships, DAX
  Larkspur-Sales.Report/               the report in PBIR: five pages, every visual as one JSON file

Everything is generated from this script, so the model and the report can be reviewed in Git line by line.
Run:  python scripts/build_pbip.py [--data-folder PATH]    (default C:/Larkspur/data: copy the data folder there)
Then open the .pbip in Power BI Desktop and press Refresh once (a PBIP keeps no data of its own).
"""
import argparse, json, shutil, uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
NAME = "Larkspur-Sales"
SM, RP = ROOT / f"{NAME}.SemanticModel", ROOT / f"{NAME}.Report"
NS = uuid.UUID("9e3c1d7a-4b2f-4c61-8f0e-15e15e15e150")


def tag(*parts):
    return str(uuid.uuid5(NS, "/".join(parts)))


def indent(text, tabs):
    return "\n".join("\t" * tabs + line if line.strip() else "" for line in text.strip("\n").splitlines())


def q(name):  # TMDL object name, quoted when needed
    return name if name.replace("_", "").isalnum() else "'" + name.replace("'", "''") + "'"


def csv_src(file, cols, delim=","):
    return (f'Csv.Document(File.Contents(DataFolder & "{file}"), '
            f'[Delimiter = "{delim}", Columns = {cols}, Encoding = 65001, QuoteStyle = QuoteStyle.Csv])')


# ====================================================================================== Power Query (M)
NUM = r'''
(t as nullable text) as number =>
    // "1250.00" / "" -> number (webshop export, US format)
    if t = null or Text.Trim(t) = "" then 0 else Number.FromText(Text.Trim(t), "en-US")
'''
NUM_EU = r'''
(t as nullable text) as number =>
    // "1.024,50" / "24,90" -> number (EU marketplace export: dot for thousands, comma for decimals)
    if t = null or Text.Trim(t) = "" then 0 else Number.FromText(Text.Replace(Text.Remove(Text.Trim(t), "."), ",", "."), "en-US")
'''

SHOP_RAW = f'''
let
    // webshop order export (Shopify layout): one row per order line
    Source = {csv_src("shop-orders-export.csv", 19)},
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Rows = Table.SelectRows(Promoted, each Text.Trim([Name]) <> ""),
    Indexed = Table.AddIndexColumn(Rows, "Row", 1, 1, Int64.Type)
in
    Table.Buffer(Indexed)
'''

SHOP_LINES = r'''
let
    // every line of every order; SKUs typed by hand ("lh-cn-001", "LH-CN-001 ") are trimmed and upper-cased
    Out = Table.FromRecords(Table.TransformRows(ShopRaw, (r) => [
        #"Line ID" = "S-" & Text.PadStart(Text.From(r[Row]), 5, "0"),
        #"Order ID" = r[Name],
        SKU = Text.Upper(Text.Trim(r[Lineitem sku])),
        #"SKU As Exported" = r[Lineitem sku],
        Qty = Number.From(r[Lineitem quantity]),
        #"Price Local" = Num(r[Lineitem price])
    ])),
    Gross = Table.AddColumn(Out, "Gross Local", each [Qty] * [Price Local], type number)
in
    Table.Buffer(Gross)
'''

SHOP_ORDERS = r'''
let
    // the order fields (status, currency, totals, discount, refund) are only on the FIRST line of each order
    Firsts = Table.SelectRows(ShopRaw, each Text.Trim([Financial Status]) <> ""),
    Out = Table.FromRecords(Table.TransformRows(Firsts, (r) => [
        #"Order ID" = r[Name],
        #"Order Seq" = Number.From(Text.AfterDelimiter(r[Name], "#")),
        Channel = "Webshop",
        // "2026-10-31 18:22:05 -0400": the store's own date
        Date = Date.FromText(Text.Start(r[Created at], 10), [Format = "yyyy-MM-dd", Culture = "en-US"]),
        #"Customer ID" = Text.Lower(Text.Trim(r[Email])),
        Currency = Text.Upper(Text.Trim(r[Currency])),
        Country = Text.Trim(r[Shipping Country]),
        Status = r[Financial Status],
        #"Discount Code" = r[Discount Code],
        #"Discount Local" = Num(r[Discount Amount]),
        #"Subtotal Local" = Num(r[Subtotal]),
        #"Shipping Local" = Num(r[Shipping]),
        #"Taxes Local" = Num(r[Taxes]),
        #"Total Local" = Num(r[Total]),
        #"Refunded Local" = Num(r[Refunded Amount])
    ])),
    PerOrder = Table.Group(ShopLines, {"Order ID"}, {{"Gross Local", each List.Sum([Gross Local]), type number}, {"Lines", each Table.RowCount(_), Int64.Type}}),
    Joined = Table.Join(Out, "Order ID", Table.RenameColumns(PerOrder, {{"Order ID", "Key"}}), "Key"),
    // refunds in net sales: a full refund returns the subtotal; a partial refund (tax included) is taken back to the amount before tax
    Returns = Table.AddColumn(Table.RemoveColumns(Joined, {"Key"}), "Returns Local", each
        if [Status] = "refunded" then [Subtotal Local]
        else if [Refunded Local] = 0 then 0
        else [Refunded Local] * [Subtotal Local] / ([Subtotal Local] + [Taxes Local]), type number)
in
    Table.Buffer(Returns)
'''

MARKET_LINES = f'''
let
    // EU marketplace export: semicolons, dd.mm.yyyy, decimal commas; monthly exports pasted together, so the header
    // line repeats and an export can start with the last row of the export before it
    Source = {csv_src("marketplace-orders-eu.csv", 9, ";")},
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    NoHeaders = Table.SelectRows(Promoted, each [Order ID] <> "Order ID" and Text.Trim([Order ID]) <> ""),
    Distinct = Table.Distinct(NoHeaders),
    Indexed = Table.AddIndexColumn(Distinct, "Row", 1, 1, Int64.Type),
    Out = Table.FromRecords(Table.TransformRows(Indexed, (r) => [
        #"Line ID" = "M-" & Text.PadStart(Text.From(r[Row]), 5, "0"),
        #"Order ID" = r[Order ID],
        Date = Date.FromText(r[Order date], [Format = "dd.MM.yyyy", Culture = "en-US"]),
        #"Customer ID" = r[Buyer ID],
        SKU = Text.Upper(Text.Trim(r[SKU])),
        #"SKU As Exported" = r[SKU],
        Qty = Number.From(r[Qty]),
        #"Price Local" = NumEU(r[Item price]),
        #"Gross Local" = NumEU(r[Line total]),
        Country = r[Country],
        #"Line Status" = r[Status]
    ]))
in
    Table.Buffer(Out)
'''

MARKET_ORDERS = r'''
let
    G = Table.Group(MarketLines, {"Order ID"}, {
        {"Date", each List.Min([Date]), type date}, {"Customer ID", each [Customer ID]{0}, type text},
        {"Country", each [Country]{0}, type text}, {"Gross Local", each List.Sum([Gross Local]), type number},
        {"Returns Local", each List.Sum(Table.SelectRows(_, each [Line Status] = "Returned")[Gross Local]) ?? 0, type number},
        {"Returned Lines", each Table.RowCount(Table.SelectRows(_, each [Line Status] = "Returned")), Int64.Type},
        {"Lines", each Table.RowCount(_), Int64.Type}}),
    Out = Table.FromRecords(Table.TransformRows(G, (r) => [
        #"Order ID" = r[Order ID], #"Order Seq" = Number.From(Text.AfterDelimiter(r[Order ID], "-")), Channel = "Marketplace EU",
        Date = r[Date], #"Customer ID" = r[Customer ID], Currency = "EUR", Country = r[Country],
        Status = if r[Returned Lines] = 0 then "paid" else if r[Returned Lines] = r[Lines] then "refunded" else "partially_refunded",
        #"Discount Code" = "", #"Discount Local" = 0, #"Subtotal Local" = r[Gross Local], #"Shipping Local" = 0, #"Taxes Local" = 0,
        #"Total Local" = r[Gross Local], #"Refunded Local" = r[Returns Local], #"Gross Local" = r[Gross Local], Lines = r[Lines],
        #"Returns Local" = r[Returns Local]]))
in
    Table.Buffer(Out)
'''

PRODUCTS_RAW = f'''
let
    Source = {csv_src("products.csv", 5)},
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {{{{"List Price USD", Currency.Type}}, {{"Unit Cost USD", Currency.Type}}}}, "en-US"),
    Clean = Table.TransformColumns(Typed, {{{{"SKU", each Text.Upper(Text.Trim(_)), type text}}}})
in
    Clean
'''

ORDERS = r'''
let
    // one row per order from both channels; voided webshop orders (payment never taken) are left out
    Cols = {"Order ID", "Order Seq", "Channel", "Date", "Customer ID", "Currency", "Country", "Status", "Discount Code", "Lines",
            "Gross Local", "Discount Local", "Subtotal Local", "Shipping Local", "Taxes Local", "Total Local", "Refunded Local", "Returns Local"},
    Shop = Table.SelectColumns(Table.SelectRows(ShopOrders, each [Status] <> "voided"), Cols),
    All = Table.Combine({Shop, Table.SelectColumns(MarketOrders, Cols)}),
    // the first order of each customer (date, then order number) makes a new customer; later orders are returning
    Keyed = Table.Buffer(Table.AddColumn(All, "Sort Key", each Number.From([Date]) * 10000000000 + [Order Seq], type number)),
    First = Table.Group(Keyed, {"Customer ID"}, {{"First Key", each List.Min([Sort Key]), type number}, {"First Date", each List.Min([Date]), type date}}),
    Joined = Table.Join(Keyed, "Customer ID", Table.RenameColumns(First, {{"Customer ID", "CK"}}), "CK"),
    IsNew = Table.AddColumn(Joined, "Is New Customer", each [Sort Key] = [First Key], type logical),
    CType = Table.AddColumn(IsNew, "Customer Type", each if [Is New Customer] then "New" else "Returning", type text),
    Since = Table.AddColumn(CType, "Months Since First", each
        (Date.Year([Date]) - Date.Year([First Date])) * 12 + Date.Month([Date]) - Date.Month([First Date]), Int64.Type),
    Out = Table.RemoveColumns(Since, {"CK", "Sort Key", "First Key", "First Date"}),
    Typed = Table.TransformColumnTypes(Out, {{"Order ID", type text}, {"Order Seq", Int64.Type}, {"Channel", type text}, {"Date", type date},
        {"Customer ID", type text}, {"Currency", type text}, {"Country", type text}, {"Status", type text}, {"Discount Code", type text},
        {"Lines", Int64.Type}, {"Gross Local", type number}, {"Discount Local", type number}, {"Subtotal Local", type number},
        {"Shipping Local", type number}, {"Taxes Local", type number}, {"Total Local", type number}, {"Refunded Local", type number},
        {"Returns Local", type number}})
in
    Typed
'''

SALES = r'''
let
    // one row per sold line; order discounts and refunds are spread over the lines of the order by their value
    O = Table.Buffer(Table.SelectColumns(Orders, {"Order ID", "Channel", "Status", "Gross Local", "Discount Local", "Subtotal Local", "Returns Local"})),
    Ok = Table.RenameColumns(O, {{"Order ID", "OK"}, {"Status", "Order Status"}, {"Gross Local", "Order Gross"},
        {"Discount Local", "Order Discount"}, {"Subtotal Local", "Order Subtotal"}, {"Returns Local", "Order Returns"}}),
    ShopJ = Table.Join(ShopLines, "Order ID", Table.SelectRows(Ok, each [Channel] = "Webshop"), "OK"),
    ShopX = Table.AddColumn(ShopJ, "Discount Local", each if [Order Gross] = 0 then 0 else [Order Discount] * [Gross Local] / [Order Gross], type number),
    ShopR = Table.AddColumn(ShopX, "Returns Local", each if [Order Subtotal] = 0 then 0 else [Order Returns] * ([Gross Local] - [Discount Local]) / [Order Subtotal], type number),
    ShopB = Table.AddColumn(ShopR, "Goods Back", each [Order Status] = "refunded", type logical),
    MktJ = Table.Join(Table.SelectColumns(MarketLines, {"Line ID", "Order ID", "SKU", "SKU As Exported", "Qty", "Price Local", "Gross Local", "Line Status"}),
        "Order ID", Table.SelectRows(Ok, each [Channel] = "Marketplace EU"), "OK"),
    MktX = Table.AddColumn(MktJ, "Discount Local", each 0, type number),
    MktR = Table.AddColumn(MktX, "Returns Local", each if [Line Status] = "Returned" then [Gross Local] else 0, type number),
    MktB = Table.AddColumn(MktR, "Goods Back", each [Line Status] = "Returned", type logical),
    Cols = {"Line ID", "Order ID", "Channel", "SKU", "SKU As Exported", "Qty", "Price Local", "Gross Local", "Discount Local", "Returns Local", "Goods Back"},
    All = Table.Combine({Table.SelectColumns(ShopB, Cols), Table.SelectColumns(MktB, Cols)}),
    Net = Table.AddColumn(All, "Net Local", each [Gross Local] - [Discount Local] - [Returns Local], type number),
    // cost of goods in USD from the product list; returned goods go back to stock, an unknown SKU has no cost
    Costs = Table.Buffer(Table.SelectColumns(ProductsRaw, {"SKU", "Unit Cost USD"})),
    WithCost = Table.NestedJoin(Net, {"SKU"}, Costs, {"SKU"}, "P", JoinKind.LeftOuter),
    Cost = Table.ExpandTableColumn(WithCost, "P", {"Unit Cost USD"}),
    Cogs = Table.AddColumn(Cost, "COGS USD", each if [Goods Back] or [Unit Cost USD] = null then 0 else [Qty] * [Unit Cost USD], type number),
    Out = Table.RemoveColumns(Cogs, {"Unit Cost USD", "Goods Back"}),
    Typed = Table.TransformColumnTypes(Out, {{"Line ID", type text}, {"Order ID", type text}, {"Channel", type text}, {"SKU", type text},
        {"SKU As Exported", type text}, {"Qty", Int64.Type}, {"Price Local", type number}, {"Gross Local", type number},
        {"Discount Local", type number}, {"Returns Local", type number}, {"Net Local", type number}, {"COGS USD", type number}})
in
    Typed
'''

PRODUCT = r'''
let
    // the product list, plus one row per SKU that appears in the exports but not in products.csv (shown, not hidden)
    Known = Table.AddColumn(ProductsRaw, "In Product List", each true, type logical),
    Used = List.Distinct(ShopLines[SKU] & MarketLines[SKU]),
    Missing = List.Difference(Used, Known[SKU]),
    Extra = Table.FromRecords(List.Transform(Missing, (s) => [SKU = s, Product = s & " (not in products.csv)",
        Category = "Not in product list", #"List Price USD" = null, #"Unit Cost USD" = null, #"In Product List" = false])),
    All = Table.Combine({Known, Extra}),
    Typed = Table.TransformColumnTypes(All, {{"SKU", type text}, {"Product", type text}, {"Category", type text},
        {"List Price USD", Currency.Type}, {"Unit Cost USD", Currency.Type}, {"In Product List", type logical}})
in
    Typed
'''

CUSTOMER = r'''
let
    // one row per customer: country and channel of the first order, the month they became a customer (cohort)
    Sorted = Table.Buffer(Table.Sort(Orders, {{"Date", Order.Ascending}, {"Order Seq", Order.Ascending}})),
    G = Table.Group(Sorted, {"Customer ID"}, {
        {"First Order Date", each List.Min([Date]), type date}, {"Channel", each [Channel]{0}, type text},
        {"Country", each [Country]{0}, type text}, {"Orders", each Table.RowCount(_), Int64.Type}}),
    No = Table.AddColumn(G, "Cohort No", each Date.Year([First Order Date]) * 100 + Date.Month([First Order Date]), Int64.Type),
    Cohort = Table.AddColumn(No, "Cohort", each Date.ToText([First Order Date], [Format = "MMM yyyy", Culture = "en-US"]), type text),
    Repeat = Table.AddColumn(Cohort, "Repeat Customer", each [Orders] > 1, type logical)
in
    Repeat
'''

MARKET = r'''
let
    Source = #table(
        type table [Country = text, Market = text, #"Market Sort" = Int64.Type],
        {
            {"United States", "United States", 1}, {"Canada", "Canada", 2}, {"United Kingdom", "United Kingdom", 3},
            {"Germany", "Europe (EU)", 4}, {"France", "Europe (EU)", 4}, {"Netherlands", "Europe (EU)", 4}, {"Ireland", "Europe (EU)", 4},
            {"Italy", "Europe (EU)", 4}, {"Spain", "Europe (EU)", 4}, {"Austria", "Europe (EU)", 4}, {"Australia", "Australia", 5}
        })
in
    Source
'''

CURRENCY = r'''
let
    Source = #table(
        type table [Currency = text, #"Currency Name" = text, Sort = Int64.Type],
        {{"USD", "US dollar", 1}, {"EUR", "Euro", 2}, {"GBP", "Pound sterling", 3}, {"CAD", "Canadian dollar", 4}, {"AUD", "Australian dollar", 5}})
in
    Source
'''

DATE = r'''
let
    // one row per day, whole months, from the first to the last order
    Start = Date.StartOfMonth(List.Min(Orders[Date])),
    End = Date.EndOfMonth(List.Max(Orders[Date])),
    Days = List.Dates(Start, Duration.Days(End - Start) + 1, #duration(1, 0, 0, 0)),
    Base = Table.FromList(Days, Splitter.SplitByNothing(), type table [Date = date]),
    Year = Table.AddColumn(Base, "Year", each Date.Year([Date]), Int64.Type),
    MonthNo = Table.AddColumn(Year, "Month No", each Date.Year([Date]) * 100 + Date.Month([Date]), Int64.Type),
    Month = Table.AddColumn(MonthNo, "Month", each Date.ToText([Date], [Format = "MMM yyyy", Culture = "en-US"]), type text),
    MoY = Table.AddColumn(Month, "Month Of Year", each Date.Month([Date]), Int64.Type),
    MName = Table.AddColumn(MoY, "Month Name", each Date.ToText([Date], [Format = "MMM", Culture = "en-US"]), type text),
    MStart = Table.AddColumn(MName, "Month Start", each Date.StartOfMonth([Date]), type date),
    Quarter = Table.AddColumn(MStart, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text)
in
    Quarter
'''

FX_API = r'''
let
    // daily ECB reference rates from the free Frankfurter API (no key): 1 USD = x units of each currency.
    // Static base URL + RelativePath/Query, so the query also refreshes in the Power BI service.
    // If the API cannot be reached, this table stays empty and the report falls back on data/fx-rates-saved.csv.
    Response = try Json.Document(Web.Contents("https://api.frankfurter.dev",
        [RelativePath = "v1/" & FxStart & "..", Query = [base = "USD", symbols = "EUR,GBP,CAD,AUD"], Timeout = #duration(0, 0, 0, 30)])) otherwise null,
    Rates = if Response = null then [] else Response[rates],
    Rows = List.Combine(List.Transform(Record.FieldNames(Rates), (d) =>
        let day = Record.Field(Rates, d) in
            List.Transform(Record.FieldNames(day), (c) => {Date.FromText(d, [Format = "yyyy-MM-dd", Culture = "en-US"]), c, Record.Field(day, c)}))),
    T = #table(type table [Date = date, Currency = text, #"Units per USD" = number], Rows)
in
    T
'''

FX_SAVED = f'''
let
    // the copy saved by scripts/fetch_rates.py: used for any day the API did not return
    Source = {csv_src("fx-rates-saved.csv", 3)},
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {{{{"Date", type date}}, {{"Currency", type text}}, {{"Units per USD", type number}}}}, "en-US")
in
    Typed
'''

FX_RATE_DAX = r'''
-- live rates first; the saved copy fills every day/currency the API did not return (all of them when it is down)
VAR api = SELECTCOLUMNS ( 'FX API', "Date", 'FX API'[Date], "Currency", 'FX API'[Currency], "Units per USD", 'FX API'[Units per USD], "Source", "API" )
VAR apiKeys = SELECTCOLUMNS ( 'FX API', "K", FORMAT ( 'FX API'[Date], "yyyymmdd" ) & 'FX API'[Currency] )
VAR saved =
    FILTER (
        SELECTCOLUMNS ( 'FX Saved', "Date", 'FX Saved'[Date], "Currency", 'FX Saved'[Currency], "Units per USD", 'FX Saved'[Units per USD], "Source", "Saved copy" ),
        NOT ( ( FORMAT ( [Date], "yyyymmdd" ) & [Currency] ) IN apiKeys )
    )
RETURN
    ADDCOLUMNS (
        UNION ( api, saved ),
        "Year", YEAR ( [Date] ),
        "Month Of Year", MONTH ( [Date] ),
        "USD per Unit", DIVIDE ( 1, [Units per USD] )
    )
'''

ADS = f'''
let
    Source = {csv_src("ads-spend.csv", 3)},
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Month = Table.TransformColumns(Promoted, {{{{"Month", each Date.FromText(_ & "-01", [Format = "yyyy-MM-dd", Culture = "en-US"]), type date}}}}),
    Typed = Table.TransformColumnTypes(Month, {{{{"Channel", type text}}, {{"Spend USD", type number}}}}, "en-US"),
    Out = Table.RenameColumns(Typed, {{{{"Channel", "Ad Channel"}}}})
in
    Out
'''

SOURCE_FILES = f'''
let
    // control figures read straight from each export a second, independent way (column positions, no clean-up steps)
    Shop = Table.Buffer(Table.Skip({csv_src("shop-orders-export.csv", 19)}, 1)),
    ShopOrd = Table.SelectRows(Shop, each [Column3] <> ""),
    VoidedIds = List.Buffer(Table.SelectRows(ShopOrd, each [Column3] = "voided")[Column1]),   // read once, not once per row
    VoidedLines = Table.RowCount(Table.SelectRows(Shop, each List.Contains(VoidedIds, [Column1]))),
    Mkt = Table.Skip({csv_src("marketplace-orders-eu.csv", 9, ";")}, 1),
    MktRows = Table.SelectRows(Mkt, each Text.Trim([Column1]) <> ""),
    Headers = Table.RowCount(Table.SelectRows(MktRows, each [Column1] = "Order ID")),
    MktData = Table.SelectRows(MktRows, each [Column1] <> "Order ID"),
    Dups = Table.RowCount(MktData) - Table.RowCount(Table.Distinct(MktData)),
    T = #table(type table [Channel = text, File = text, #"Rows In File" = Int64.Type, #"Orders In File" = Int64.Type,
                           #"Voided Orders" = Int64.Type, #"Voided Lines" = Int64.Type, #"Header Repeats" = Int64.Type, Duplicates = Int64.Type], {{
        {{"Webshop", "shop-orders-export.csv", Table.RowCount(Shop), Table.RowCount(ShopOrd), List.Count(VoidedIds), VoidedLines, 0, 0}},
        {{"Marketplace EU", "marketplace-orders-eu.csv", Table.RowCount(MktRows), List.Count(List.Distinct(MktData[Column1])), 0, 0, Headers, Dups}}
    }})
in
    T
'''

FILE_TOTALS = f'''
let
    // order totals per currency as typed in the files: webshop Total of every order that was not voided,
    // marketplace line totals (each row once)
    Shop = Table.SelectRows(Table.Skip({csv_src("shop-orders-export.csv", 19)}, 1), each [Column3] <> "" and [Column3] <> "voided"),
    ShopT = Table.Group(Shop, {{"Column6"}}, {{{{"Total", each List.Sum(List.Transform([Column10], each Number.FromText(_, "en-US"))), type number}}}}),
    Mkt = Table.Distinct(Table.SelectRows(Table.Skip({csv_src("marketplace-orders-eu.csv", 9, ";")}, 1), each [Column1] <> "Order ID" and Text.Trim([Column1]) <> "")),
    MktTotal = List.Sum(List.Transform(Mkt[Column7], each Number.FromText(Text.Replace(Text.Remove(_, "."), ",", "."), "en-US"))),
    T = Table.FromRecords(List.Transform(Table.ToRecords(ShopT), (r) => [Currency = r[Column6],
        #"Total In File" = r[Total] + (if r[Column6] = "EUR" then MktTotal else 0)])),
    Typed = Table.TransformColumnTypes(T, {{{{"Currency", type text}}, {{"Total In File", type number}}}})
in
    Typed
'''

CHECKS = r'''
let
    Source = #table(
        type table [No = Int64.Type, Check = text, #"Why it matters" = text],
        {
            {1, "Every webshop line is in the model once", "Lines in the file = sold lines in the model + lines of voided orders"},
            {2, "Every marketplace line is in the model once", "Rows in the file - repeated headers - rows exported twice = lines in the model"},
            {3, "Every order appears once", "No order number twice after the two channels are put together"},
            {4, "Order lines add up to the subtotal", "Sum of quantity x price - discount = subtotal, for every order"},
            {5, "Order totals add up", "Subtotal + shipping + taxes = total, for every order"},
            {6, "Totals per currency = the files", "Order totals per currency, read a second way from the files, = the model"},
            {7, "Every order has an exchange rate", "A rate on the order date or the last working day before it (at most 5 days old)"},
            {8, "Live rates = the saved copy", "The API answer agrees with data/fx-rates-saved.csv on every day both have"},
            {9, "No refund is larger than its order", "Refunded amount <= order total"},
            {10, "Lines waiting for a product", "SKUs that are not in products.csv: add them and refresh"}
        })
in
    Source
'''

INSIGHTS = r'''
let
    Source = #table(type table [No = Int64.Type, Topic = text],
        {{1, "Sales"}, {2, "Exchange rates"}, {3, "Fastest growing market"}, {4, "Best seller"}, {5, "Customers"}, {6, "Refunds"}})
in
    Source
'''

MONTHS_LATER = r'''
let
    Source = Table.FromList({0..12}, Splitter.SplitByNothing(), type table [#"Months Later" = Int64.Type])
in
    Source
'''

MEASURES_SRC = r'''
let
    Source = #table(type table [#"Measure Group" = text], {})
in
    Source
'''

# ====================================================================================== DAX
PCT, PCT_SIGNED, USD, USD_SIGNED = "0.0%;-0.0%;0.0%", "+0.0%;-0.0%;0.0%", "#,0", "+#,0;-#,0;0"
MEASURES = [  # name, folder, format, expression, hidden
    # ---- sales (USD at the rate of the order date)
    ("Gross Sales", "1 Sales", USD, "SUM ( Sales[Gross USD] )", False),
    ("Discounts", "1 Sales", USD, "SUM ( Sales[Discount USD] )", False),
    ("Returns", "1 Sales", USD, "-- refunds, in the month of the order\nSUM ( Sales[Returns USD] )", False),
    ("Net Sales", "1 Sales", USD, "-- gross - discounts - refunds; without shipping and taxes\nSUM ( Sales[Net USD] )", False),
    ("Orders", "1 Sales", "#,0", "COUNTROWS ( Orders )", False),
    ("Units", "1 Sales", "#,0", "SUM ( Sales[Qty] )", False),
    ("Average Order Value", "1 Sales", "#,0.00", "-- order value after discounts, before refunds\nDIVIDE ( [Gross Sales] - [Discounts], [Orders] )", False),
    ("COGS", "1 Sales", USD, "SUM ( Sales[COGS USD] )", False),
    ("Gross Profit", "1 Sales", USD, "[Net Sales] - [COGS]", False),
    ("Gross Margin %", "1 Sales", PCT, "DIVIDE ( [Gross Profit], [Net Sales] )", False),
    ("Refund Rate %", "1 Sales", PCT, "DIVIDE ( [Returns], [Gross Sales] - [Discounts] )", False),
    ("Shipping Charged", "1 Sales", USD, "SUM ( Orders[Shipping USD] )", False),
    ("Taxes Collected", "1 Sales", USD, "SUM ( Orders[Taxes USD] )", False),
    ("Share of Net Sales %", "1 Sales", PCT, "DIVIDE ( [Net Sales], CALCULATE ( [Net Sales], REMOVEFILTERS ( Product ) ) )", False),
    # ---- time intelligence
    ("Net Sales LY", "2 Time intelligence", USD, "CALCULATE ( [Net Sales], SAMEPERIODLASTYEAR ( 'Date'[Date] ) )", False),
    ("Net Sales YoY %", "2 Time intelligence", PCT_SIGNED, "VAR ly = [Net Sales LY]\nRETURN IF ( NOT ISBLANK ( ly ), DIVIDE ( [Net Sales] - ly, ly ) )", False),
    ("Orders LY", "2 Time intelligence", "#,0", "CALCULATE ( [Orders], SAMEPERIODLASTYEAR ( 'Date'[Date] ) )", False),
    ("Orders YoY %", "2 Time intelligence", PCT_SIGNED, "VAR ly = [Orders LY]\nRETURN IF ( NOT ISBLANK ( ly ), DIVIDE ( [Orders] - ly, ly ) )", False),
    ("AOV LY", "2 Time intelligence", "#,0.00", "CALCULATE ( [Average Order Value], SAMEPERIODLASTYEAR ( 'Date'[Date] ) )", False),
    ("Gross Margin LY %", "2 Time intelligence", PCT, "CALCULATE ( [Gross Margin %], SAMEPERIODLASTYEAR ( 'Date'[Date] ) )", False),
    ("Refund Rate LY %", "2 Time intelligence", PCT, "CALCULATE ( [Refund Rate %], SAMEPERIODLASTYEAR ( 'Date'[Date] ) )", False),
    ("Net Sales YTD", "2 Time intelligence", USD, "TOTALYTD ( [Net Sales], 'Date'[Date] )", False),
    ("Net Sales LY YTD", "2 Time intelligence", USD, "CALCULATE ( [Net Sales YTD], SAMEPERIODLASTYEAR ( 'Date'[Date] ) )", False),
    # ---- currencies
    ("Net Sales Local", "3 Currencies", "#,0", "-- in the currency of the orders: only meaningful for one currency\nSUM ( Sales[Net Local] )", False),
    ("Effective Rate", "3 Currencies", "0.0000", "-- units per USD actually used, weighted by sales\nDIVIDE ( [Net Sales Local], [Net Sales] )", False),
    ("Non-USD Share %", "3 Currencies", PCT,
     'DIVIDE ( CALCULATE ( [Net Sales], KEEPFILTERS ( Orders[Currency] <> "USD" ) ), [Net Sales] )', False),
    ("Net Sales CC", "3 Currencies", USD,
     "-- constant currency: this year's sales in local currency at last year's average rate for the same month\n"
     "VAR firstYear = CALCULATE ( MIN ( 'Date'[Year] ), REMOVEFILTERS ( 'Date' ) )   -- the first year has no year before it\nRETURN\n"
     "SUMX (\n    SUMMARIZE ( Sales, Orders[Currency], 'Date'[Year], 'Date'[Month Of Year] ),\n"
     "    VAR c = Orders[Currency]\n    VAR y = 'Date'[Year]\n    VAR m = 'Date'[Month Of Year]\n"
     "    VAR r = IF ( y - 1 < firstYear, BLANK (), IF ( c = \"USD\", 1,\n        CALCULATE ( AVERAGE ( 'FX Rate'[Units per USD] ), REMOVEFILTERS ( 'FX Rate' ),\n"
     "            'FX Rate'[Currency] = c, 'FX Rate'[Year] = y - 1, 'FX Rate'[Month Of Year] = m ) ) )\n"
     "    RETURN IF ( NOT ISBLANK ( r ), DIVIDE ( CALCULATE ( SUM ( Sales[Net Local] ) ), r ) )\n)", False),
    ("FX Effect", "3 Currencies", USD_SIGNED,
     "-- what exchange rates added to (or took from) USD sales compared with last year's rates\n"
     "VAR firstYear = CALCULATE ( MIN ( 'Date'[Year] ), REMOVEFILTERS ( 'Date' ) )   -- the first year has no year before it\nRETURN\n"
     "SUMX (\n    SUMMARIZE ( Sales, Orders[Currency], 'Date'[Year], 'Date'[Month Of Year] ),\n"
     "    VAR c = Orders[Currency]\n    VAR y = 'Date'[Year]\n    VAR m = 'Date'[Month Of Year]\n"
     "    VAR r = IF ( y - 1 < firstYear, BLANK (), IF ( c = \"USD\", 1,\n        CALCULATE ( AVERAGE ( 'FX Rate'[Units per USD] ), REMOVEFILTERS ( 'FX Rate' ),\n"
     "            'FX Rate'[Currency] = c, 'FX Rate'[Year] = y - 1, 'FX Rate'[Month Of Year] = m ) ) )\n"
     "    RETURN IF ( NOT ISBLANK ( r ), CALCULATE ( SUM ( Sales[Net USD] ) ) - DIVIDE ( CALCULATE ( SUM ( Sales[Net Local] ) ), r ) )\n)", False),
    ("CC Growth %", "3 Currencies", PCT_SIGNED,
     "VAR ly = [Net Sales LY]\nVAR cc = [Net Sales CC]\nRETURN IF ( NOT ISBLANK ( ly ) && NOT ISBLANK ( cc ), DIVIDE ( cc - ly, ly ) )", False),
    ("Latest Rate", "3 Currencies", "0.0000",
     "VAR d = CALCULATE ( MAX ( 'FX Rate'[Date] ), REMOVEFILTERS ( 'FX Rate'[Date] ) )\n"
     "RETURN IF ( HASONEVALUE ( 'Currency'[Currency] ), CALCULATE ( AVERAGE ( 'FX Rate'[Units per USD] ), 'FX Rate'[Date] = d ) )", False),
    ("USD per Unit", "3 Currencies", "0.0000", "AVERAGE ( 'FX Rate'[USD per Unit] )", False),
    ("FX Index", "3 Currencies", "0.0",
     "-- value of one unit in USD, first day of the rates = 100\n"
     "VAR d0 = CALCULATE ( MIN ( 'FX Rate'[Date] ), REMOVEFILTERS ( 'FX Rate'[Date] ) )\n"
     "VAR v0 = CALCULATE ( AVERAGE ( 'FX Rate'[USD per Unit] ), REMOVEFILTERS ( 'FX Rate'[Date] ), 'FX Rate'[Date] = d0 )\n"
     "RETURN DIVIDE ( [USD per Unit], v0 ) * 100", False),
    ("Rates Latest Date", "3 Currencies", "d mmm yyyy", "CALCULATE ( MAX ( 'FX Rate'[Date] ), REMOVEFILTERS () )", False),
    ("API Rates", "3 Currencies", "#,0", "CALCULATE ( COUNTROWS ( 'FX Rate' ), REMOVEFILTERS (), 'FX Rate'[Source] = \"API\" ) + 0", False),
    ("Rate Source", "3 Currencies", None,
     'IF ( [API Rates] > 0, "Live API", "Saved copy" )', False),
    ("Rate Source Note", "3 Currencies", None,
     "IF ( [API Rates] > 0, \"Frankfurter (ECB): \", \"API not reached: \" ) & FORMAT ( CALCULATE ( COUNTROWS ( 'FX Rate' ), REMOVEFILTERS () ), \"#,0\" ) & \" rates to \" & FORMAT ( [Rates Latest Date], \"d mmm yyyy\" )", True),
    ("Rate Source Color", "3 Currencies", None, 'IF ( [API Rates] > 0, "#3F8A4E", "#D9A441" )', True),
    # ---- customers and marketing
    ("Customers", "4 Customers", "#,0", "DISTINCTCOUNT ( Orders[Customer ID] )", False),
    ("New Customers", "4 Customers", "#,0", "CALCULATE ( COUNTROWS ( Orders ), KEEPFILTERS ( Orders[Is New Customer] = TRUE () ) )", False),
    ("Returning Orders", "4 Customers", "#,0", "CALCULATE ( COUNTROWS ( Orders ), KEEPFILTERS ( Orders[Is New Customer] = FALSE () ) )", False),
    ("New Customer Orders", "4 Customers", "#,0", "[New Customers]", False),
    ("Returning Share %", "4 Customers", PCT, "DIVIDE ( [Returning Orders], [Orders] )", False),
    ("Returning Share LY %", "4 Customers", PCT, "CALCULATE ( [Returning Share %], SAMEPERIODLASTYEAR ( 'Date'[Date] ) )", False),
    ("Ad Spend", "4 Customers", USD, "SUM ( 'Ad Spend'[Spend USD] )", False),
    ("MER", "4 Customers", '0.0"x"',
     "-- marketing efficiency ratio: webshop net sales per dollar of ads (the marketplace does its own marketing)\n"
     "DIVIDE ( CALCULATE ( [Net Sales], KEEPFILTERS ( Orders[Channel] = \"Webshop\" ) ), [Ad Spend] )", False),
    ("CAC", "4 Customers", "#,0.00",
     "-- ad spend per new webshop customer\nDIVIDE ( [Ad Spend], CALCULATE ( [New Customers], KEEPFILTERS ( Orders[Channel] = \"Webshop\" ) ) )", False),
    ("Cohort Customers", "4 Customers", "#,0",
     "CALCULATE ( [New Customers], REMOVEFILTERS ( 'Date' ), REMOVEFILTERS ( Orders[Months Since First] ) )", False),
    ("Cohort Retention %", "4 Customers", "0%",
     "-- share of a month's new customers who ordered again N months later (2025 cohorts, first 12 months)\n"
     "VAR k = SELECTEDVALUE ( 'Months Later'[Months Later] )\nVAR c = SELECTEDVALUE ( Customer[Cohort No] )\n"
     "VAR active = CALCULATE ( DISTINCTCOUNT ( Orders[Customer ID] ), REMOVEFILTERS ( 'Date' ), Orders[Months Since First] = k )\n"
     "RETURN IF ( NOT ISBLANK ( k ) && NOT ISBLANK ( c ) && c < 202601, DIVIDE ( active, [Cohort Customers] ) )", False),
    ("Cohort Color", "4 Customers", None,
     "VAR p = [Cohort Retention %]\nRETURN SWITCH ( TRUE (), ISBLANK ( p ), \"#FFFFFF\", p >= 0.99, \"#E8DFD2\", p >= 0.06, \"#C2603D\", p >= 0.04, \"#D98B6A\",\n"
     "    p >= 0.025, \"#E3A587\", p >= 0.01, \"#F1CDB9\", \"#FBF0E8\" )", True),
    ("Cohort Font", "4 Customers", None, 'IF ( [Cohort Retention %] >= 0.06 && [Cohort Retention %] < 0.99, "#FFFFFF", "#3B2F2A" )', True),
    # ---- checks: each helper ignores every slicer
    ("Chk Shop Rows", "5 Checks", "#,0", "CALCULATE ( SUM ( 'Source Files'[Rows In File] ), REMOVEFILTERS (), 'Source Files'[Channel] = \"Webshop\" )", True),
    ("Chk Shop Voided Lines", "5 Checks", "#,0", "CALCULATE ( SUM ( 'Source Files'[Voided Lines] ), REMOVEFILTERS () )", True),
    ("Chk Shop Voided Orders", "5 Checks", "#,0", "CALCULATE ( SUM ( 'Source Files'[Voided Orders] ), REMOVEFILTERS () )", True),
    ("Chk Shop Lines", "5 Checks", "#,0", "CALCULATE ( COUNTROWS ( Sales ), REMOVEFILTERS (), Sales[Channel] = \"Webshop\" ) + 0", True),
    ("Chk Mkt Rows", "5 Checks", "#,0", "CALCULATE ( SUM ( 'Source Files'[Rows In File] ), REMOVEFILTERS (), 'Source Files'[Channel] = \"Marketplace EU\" )", True),
    ("Chk Mkt Headers", "5 Checks", "#,0", "CALCULATE ( SUM ( 'Source Files'[Header Repeats] ), REMOVEFILTERS () )", True),
    ("Chk Mkt Duplicates", "5 Checks", "#,0", "CALCULATE ( SUM ( 'Source Files'[Duplicates] ), REMOVEFILTERS () )", True),
    ("Chk Mkt Lines", "5 Checks", "#,0", "CALCULATE ( COUNTROWS ( Sales ), REMOVEFILTERS (), Sales[Channel] = \"Marketplace EU\" ) + 0", True),
    ("Chk Orders", "5 Checks", "#,0", "CALCULATE ( COUNTROWS ( Orders ), REMOVEFILTERS () )", True),
    ("Chk Order IDs", "5 Checks", "#,0", "CALCULATE ( DISTINCTCOUNT ( Orders[Order ID] ), REMOVEFILTERS () )", True),
    ("Chk Lines Off", "5 Checks", "0",
     "COUNTROWS ( FILTER ( CALCULATETABLE ( Orders, REMOVEFILTERS () ),\n"
     "    ABS ( Orders[Gross Local] - Orders[Discount Local] - Orders[Subtotal Local] ) > 0.005 ) ) + 0", True),
    ("Chk Totals Off", "5 Checks", "0",
     "COUNTROWS ( FILTER ( CALCULATETABLE ( Orders, REMOVEFILTERS () ),\n"
     "    ABS ( Orders[Subtotal Local] + Orders[Shipping Local] + Orders[Taxes Local] - Orders[Total Local] ) > 0.005 ) ) + 0", True),
    ("Chk Currencies Off", "5 Checks", "0",
     "COUNTROWS ( FILTER ( CALCULATETABLE ( 'File Totals', REMOVEFILTERS () ),\n"
     "    VAR c = 'File Totals'[Currency]\n"
     "    RETURN ABS ( CALCULATE ( SUM ( Orders[Total Local] ), REMOVEFILTERS (), Orders[Currency] = c ) - 'File Totals'[Total In File] ) > 0.005 ) ) + 0", True),
    ("Chk Currencies", "5 Checks", "0", "CALCULATE ( COUNTROWS ( 'File Totals' ), REMOVEFILTERS () ) + 0", True),
    ("Chk Rate Missing", "5 Checks", "0",
     "COUNTROWS ( FILTER ( CALCULATETABLE ( Orders, REMOVEFILTERS () ), ISBLANK ( Orders[Rate] ) || Orders[Rate Age Days] > 5 ) ) + 0", True),
    ("Chk Rate Max Age", "5 Checks", "0", "CALCULATE ( MAX ( Orders[Rate Age Days] ), REMOVEFILTERS () )", True),
    ("Chk FX Compared", "5 Checks", "#,0",
     "COUNTROWS ( FILTER ( CALCULATETABLE ( 'FX API', REMOVEFILTERS () ),\n"
     "    NOT ISBLANK ( LOOKUPVALUE ( 'FX Saved'[Units per USD], 'FX Saved'[Date], 'FX API'[Date], 'FX Saved'[Currency], 'FX API'[Currency] ) ) ) ) + 0", True),
    ("Chk FX Diff", "5 Checks", "0",
     "COUNTROWS ( FILTER ( CALCULATETABLE ( 'FX API', REMOVEFILTERS () ),\n"
     "    VAR s = LOOKUPVALUE ( 'FX Saved'[Units per USD], 'FX Saved'[Date], 'FX API'[Date], 'FX Saved'[Currency], 'FX API'[Currency] )\n"
     "    RETURN NOT ISBLANK ( s ) && ABS ( s - 'FX API'[Units per USD] ) > 0.0000001 ) ) + 0", True),
    ("Chk FX New", "5 Checks", "0", "[API Rates] - [Chk FX Compared]", True),
    ("Chk Refund Over", "5 Checks", "0",
     "COUNTROWS ( FILTER ( CALCULATETABLE ( Orders, REMOVEFILTERS () ), Orders[Refunded Local] > Orders[Total Local] + 0.005 ) ) + 0", True),
    ("Chk Unknown SKU", "5 Checks", "0",
     "COUNTROWS ( FILTER ( CALCULATETABLE ( Sales, REMOVEFILTERS () ), RELATED ( Product[In Product List] ) = FALSE () ) ) + 0", True),
    ("Unknown SKU Lines", "5 Checks", "0", "COUNTROWS ( FILTER ( Sales, RELATED ( Product[In Product List] ) = FALSE () ) )", False),
    ("Check Status", "5 Checks", None,
     "VAR n = SELECTEDVALUE ( Checks[No] )\nVAR ok = SWITCH ( n,\n"
     "    1, [Chk Shop Rows] = [Chk Shop Lines] + [Chk Shop Voided Lines],\n"
     "    2, [Chk Mkt Rows] - [Chk Mkt Headers] - [Chk Mkt Duplicates] = [Chk Mkt Lines],\n"
     "    3, [Chk Orders] = [Chk Order IDs],\n"
     "    4, [Chk Lines Off] = 0,\n"
     "    5, [Chk Totals Off] = 0,\n"
     "    6, [Chk Currencies Off] = 0 && [Chk Currencies] > 0,\n"
     "    7, [Chk Rate Missing] = 0,\n"
     "    8, [Chk FX Diff] = 0 && [API Rates] > 0,\n"
     "    9, [Chk Refund Over] = 0,\n"
     "    10, [Chk Unknown SKU] = 0\n)\n"
     "RETURN\n    IF ( HASONEVALUE ( Checks[No] ),\n"
     "        IF ( ok, \"PASS\", IF ( n = 10 || ( n = 8 && [API Rates] = 0 ), \"REVIEW\", \"FAIL\" ) ) )", False),
    ("Check Detail", "5 Checks", None,
     "SWITCH ( SELECTEDVALUE ( Checks[No] ),\n"
     "    1, FORMAT ( [Chk Shop Rows], \"#,0\" ) & \" lines in the file = \" & FORMAT ( [Chk Shop Lines], \"#,0\" ) & \" sold + \"\n"
     "        & FORMAT ( [Chk Shop Voided Lines], \"#,0\" ) & \" on \" & FORMAT ( [Chk Shop Voided Orders], \"#,0\" ) & \" voided orders\",\n"
     "    2, FORMAT ( [Chk Mkt Rows], \"#,0\" ) & \" rows: \" & FORMAT ( [Chk Mkt Headers], \"0\" ) & \" repeated headers and \"\n"
     "        & FORMAT ( [Chk Mkt Duplicates], \"0\" ) & \" duplicates removed, \" & FORMAT ( [Chk Mkt Lines], \"#,0\" ) & \" lines\",\n"
     "    3, FORMAT ( [Chk Orders], \"#,0\" ) & \" orders, \" & FORMAT ( [Chk Order IDs], \"#,0\" ) & \" different order numbers\",\n"
     "    4, FORMAT ( [Chk Lines Off], \"0\" ) & \" orders differ\",\n"
     "    5, FORMAT ( [Chk Totals Off], \"0\" ) & \" orders differ\",\n"
     "    6, FORMAT ( [Chk Currencies] - [Chk Currencies Off], \"0\" ) & \" of \" & FORMAT ( [Chk Currencies], \"0\" ) & \" currencies agree to the cent\",\n"
     "    7, FORMAT ( [Chk Rate Missing], \"0\" ) & \" orders without a rate; oldest rate used: \" & FORMAT ( [Chk Rate Max Age], \"0\" ) & \" days (weekends, holidays)\",\n"
     "    8, IF ( [API Rates] = 0, \"API not reached: every rate comes from the saved copy\",\n"
     "        FORMAT ( [Chk FX Compared] - [Chk FX Diff], \"#,0\" ) & \" of \" & FORMAT ( [Chk FX Compared], \"#,0\" ) & \" rates agree, \"\n"
     "        & FORMAT ( [Chk FX New], \"0\" ) & \" newer than the copy\" ),\n"
     "    9, FORMAT ( [Chk Refund Over], \"0\" ) & \" refunds larger than the order\",\n"
     "    10, FORMAT ( [Chk Unknown SKU], \"0\" ) & \" lines with a SKU not in products.csv\"\n)", False),
    ("Check Color", "5 Checks", None, 'SWITCH ( [Check Status], "PASS", "#3F8A4E", "FAIL", "#C0392B", "REVIEW", "#B7862B" )', True),
    ("Checks Passed", "5 Checks", None,
     'VAR t = ADDCOLUMNS ( ALL ( Checks[No] ), "@s", [Check Status] )\n'
     'RETURN FORMAT ( COUNTROWS ( FILTER ( t, [@s] = "PASS" ) ), "0" ) & " of " & FORMAT ( COUNTROWS ( FILTER ( t, [@s] <> "REVIEW" ) ), "0" ) & " PASS"', False),
    ("Checks Failed", "5 Checks", "0",
     'VAR t = ADDCOLUMNS ( ALL ( Checks[No] ), "@s", [Check Status] )\nRETURN COUNTROWS ( FILTER ( t, [@s] = "FAIL" ) ) + 0', False),
    # ---- the line under each KPI and the Key insights panel
    ("Sales Note", "6 Report text", None,
     'VAR p = [Net Sales YoY %]\nRETURN IF ( ISBLANK ( p ), "no earlier year to compare", IF ( p >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( p ), "0.0%" ) & " vs last year" )', True),
    ("Sales Note Color", "6 Report text", None, 'IF ( [Net Sales YoY %] < 0, "#C0392B", "#3F8A4E" )', True),
    ("Orders Note", "6 Report text", None,
     'VAR p = [Orders YoY %]\nRETURN IF ( ISBLANK ( p ), "no earlier year to compare", IF ( p >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( p ), "0.0%" ) & " vs last year" )', True),
    ("Orders Note Color", "6 Report text", None, 'IF ( [Orders YoY %] < 0, "#C0392B", "#3F8A4E" )', True),
    ("AOV Note", "6 Report text", None,
     'IF ( ISBLANK ( [AOV LY] ), "no earlier year to compare", "last year " & FORMAT ( [AOV LY], "#,0.00" ) )', True),
    ("AOV Note Color", "6 Report text", None, 'IF ( [Average Order Value] < [AOV LY], "#C0392B", "#3F8A4E" )', True),
    ("Margin Note", "6 Report text", None,
     'IF ( ISBLANK ( [Gross Margin LY %] ), "no earlier year to compare", "last year " & FORMAT ( [Gross Margin LY %], "0.0%" ) )', True),
    ("Margin Note Color", "6 Report text", None, 'IF ( [Gross Margin %] < [Gross Margin LY %], "#C0392B", "#3F8A4E" )', True),
    ("Returning Note", "6 Report text", None,
     'IF ( ISBLANK ( [Returning Share LY %] ), "of orders", "last year " & FORMAT ( [Returning Share LY %], "0.0%" ) )', True),
    ("Returning Note Color", "6 Report text", None, 'IF ( [Returning Share %] < [Returning Share LY %], "#C0392B", "#3F8A4E" )', True),
    ("Refund Note", "6 Report text", None,
     'IF ( ISBLANK ( [Refund Rate LY %] ), "of sales", "last year " & FORMAT ( [Refund Rate LY %], "0.0%" ) )', True),
    ("Refund Note Color", "6 Report text", None, 'IF ( [Refund Rate %] > [Refund Rate LY %], "#C0392B", "#3F8A4E" )', True),
    ("CC Note", "6 Report text", None,
     'IF ( ISBLANK ( [CC Growth %] ), "no earlier year to compare", "reported " & FORMAT ( [Net Sales YoY %], "+0.0%;-0.0%" ) )', True),
    ("FX Note", "6 Report text", None, '"vs last year\'s rates"', True),
    ("FX Color", "6 Report text", None, 'IF ( [FX Effect] < 0, "#C0392B", "#3F8A4E" )', True),
    ("Non-USD Note", "6 Report text", None, '"of net sales in EUR, GBP, CAD, AUD"', True),
    ("MER Note", "6 Report text", None, '"webshop sales per $1 of ads"', True),
    ("CAC Note", "6 Report text", None, '"ads per new webshop customer"', True),
    ("Insight", "6 Report text", None,
     "-- one sentence per row of the Insights table, recalculated for the year, channel and market picked in the slicers\n"
     "VAR markets = FILTER ( ADDCOLUMNS ( VALUES ( Market[Market] ), \"@ly\", [Net Sales LY], \"@g\", [Net Sales YoY %] ), [@ly] >= 1000 )\n"
     "VAR fast = TOPN ( 1, markets, [@g], DESC )\n"
     "VAR products = ADDCOLUMNS ( VALUES ( Product[Product] ), \"@v\", [Net Sales] )\n"
     "VAR best = TOPN ( 1, products, [@v], DESC )\n"
     "VAR cats = FILTER ( ADDCOLUMNS ( VALUES ( Product[Category] ), \"@r\", [Refund Rate %] ), NOT ISBLANK ( [@r] ) )\n"
     "VAR worst = TOPN ( 1, cats, [@r], DESC )\n"
     "VAR total = [Net Sales]   -- outside MAXX: inside it the row context would turn it into one product's sales\n"
     "RETURN\n    SWITCH ( SELECTEDVALUE ( Insights[No] ),\n"
     "        1, \"Net sales \" & FORMAT ( [Net Sales], \"#,0\" ) & \" USD\" & IF ( ISBLANK ( [Net Sales YoY %] ), \"\",\n"
     "            \", \" & FORMAT ( [Net Sales YoY %], \"+0.0%;-0.0%\" ) & \" on the same months last year\" ),\n"
     "        2, IF ( ISBLANK ( [CC Growth %] ), \"No earlier year to compare rates with\",\n"
     "            \"At last year's exchange rates \" & FORMAT ( [CC Growth %], \"+0.0%;-0.0%\" ) & \": currency moves \"\n"
     "            & IF ( [FX Effect] >= 0, \"added \", \"took off \" ) & FORMAT ( ABS ( [FX Effect] ), \"#,0\" ) & \" USD\" ),\n"
     "        3, IF ( ISEMPTY ( fast ), \"No earlier year to compare\",\n"
     "            MAXX ( fast, Market[Market] & \" \" & FORMAT ( [@g], \"+0.0%;-0.0%\" ) & \" (\" & FORMAT ( [@ly], \"#,0\" ) & \" to \"\n"
     "            & FORMAT ( [@ly] * ( 1 + [@g] ), \"#,0\" ) & \")\" ) ),\n"
     "        4, MAXX ( best, Product[Product] & \": \" & FORMAT ( [@v], \"#,0\" ) & \" (\" & FORMAT ( DIVIDE ( [@v], total ), \"0.0%\" ) & \" of net sales)\" ),\n"
     "        5, \"Returning customers placed \" & FORMAT ( [Returning Share %], \"0.0%\" ) & \" of orders; ads cost \"\n"
     "            & FORMAT ( [CAC], \"#,0.00\" ) & \" per new webshop customer\",\n"
     "        6, \"Refunds \" & FORMAT ( [Refund Rate %], \"0.0%\" ) & \" of sales; highest in \"\n"
     "            & MAXX ( worst, Product[Category] & \" (\" & FORMAT ( [@r], \"0.0%\" ) & \")\" )\n    )", False),
    ("Insight Color", "6 Report text", None,
     'SWITCH ( SELECTEDVALUE ( Insights[No] ), 1, [Sales Note Color], 2, [FX Color], 3, "#3F8A4E", 6, "#C0392B", "#3B2F2A" )', True),
]

# calculated columns: table -> [(name, dataType, format, DAX)]
CALC_COLUMNS = {
    "Orders": [
        ("Rate", "double", "0.0000",
         "-- units per USD on the order date, or the last rate before it (no ECB rate at weekends and holidays)\n"
         "VAR c = Orders[Currency]\nVAR d = Orders[Date]\nRETURN\n    IF ( c = \"USD\", 1,\n"
         "        VAR last = MAXX ( FILTER ( 'FX Rate', 'FX Rate'[Currency] = c && 'FX Rate'[Date] <= d ), 'FX Rate'[Date] )\n"
         "        RETURN MAXX ( FILTER ( 'FX Rate', 'FX Rate'[Currency] = c && 'FX Rate'[Date] = last ), 'FX Rate'[Units per USD] ) )"),
        ("Rate Date", "dateTime", "d mmm yyyy",
         "VAR c = Orders[Currency]\nVAR d = Orders[Date]\nRETURN\n    IF ( c = \"USD\", d, MAXX ( FILTER ( 'FX Rate', 'FX Rate'[Currency] = c && 'FX Rate'[Date] <= d ), 'FX Rate'[Date] ) )"),
        ("Rate Age Days", "int64", "0", "INT ( Orders[Date] - Orders[Rate Date] )"),
        ("Shipping USD", "double", "#,0.00", "DIVIDE ( Orders[Shipping Local], Orders[Rate] )"),
        ("Taxes USD", "double", "#,0.00", "DIVIDE ( Orders[Taxes Local], Orders[Rate] )"),
    ],
    "Sales": [
        ("Gross USD", "double", "#,0.00", "DIVIDE ( Sales[Gross Local], RELATED ( Orders[Rate] ) )"),
        ("Discount USD", "double", "#,0.00", "DIVIDE ( Sales[Discount Local], RELATED ( Orders[Rate] ) )"),
        ("Returns USD", "double", "#,0.00", "DIVIDE ( Sales[Returns Local], RELATED ( Orders[Rate] ) )"),
        ("Net USD", "double", "#,0.00", "-- in USD at the rate of the order date\nDIVIDE ( Sales[Net Local], RELATED ( Orders[Rate] ) )"),
    ],
}

# ====================================================================================== tables (TMDL)
S_, I_, D_, DT_, B_ = "string", "int64", "double", "dateTime", "boolean"
TABLES = {
    "_Measures": dict(src=MEASURES_SRC, cols=[("Measure Group", S_, None, "none", ["isHidden"])], measures=True),
    "Sales": dict(src=SALES, desc="One row per sold order line from both channels; discounts and refunds spread over the lines of the order.", cols=[
        ("Line ID", S_, None, "none", []), ("Order ID", S_, None, "none", []), ("Channel", S_, None, "none", []),
        ("SKU", S_, None, "none", []), ("SKU As Exported", S_, None, "none", []), ("Qty", I_, "#,0", "sum", []),
        ("Price Local", D_, "#,0.00", "none", []), ("Gross Local", D_, "#,0.00", "sum", []), ("Discount Local", D_, "#,0.00", "sum", []),
        ("Returns Local", D_, "#,0.00", "sum", []), ("Net Local", D_, "#,0.00", "sum", []), ("COGS USD", D_, "#,0.00", "sum", [])]),
    "Orders": dict(src=ORDERS, desc="One row per order (webshop and marketplace), amounts in the order currency; voided orders left out.", cols=[
        ("Order ID", S_, None, "none", ["isKey"]), ("Order Seq", I_, "0", "none", ["isHidden"]), ("Channel", S_, None, "none", []),
        ("Date", DT_, "d mmm yyyy", "none", []), ("Customer ID", S_, None, "none", []), ("Currency", S_, None, "none", []),
        ("Country", S_, None, "none", []), ("Status", S_, None, "none", []), ("Discount Code", S_, None, "none", []),
        ("Lines", I_, "0", "sum", []), ("Gross Local", D_, "#,0.00", "sum", []), ("Discount Local", D_, "#,0.00", "sum", []),
        ("Subtotal Local", D_, "#,0.00", "sum", []), ("Shipping Local", D_, "#,0.00", "sum", []), ("Taxes Local", D_, "#,0.00", "sum", []),
        ("Total Local", D_, "#,0.00", "sum", []), ("Refunded Local", D_, "#,0.00", "sum", []), ("Returns Local", D_, "#,0.00", "sum", []),
        ("Is New Customer", B_, None, "none", []), ("Customer Type", S_, None, "none", []), ("Months Since First", I_, "0", "none", [])]),
    "Product": dict(src=PRODUCT, desc="Product list with unit cost in USD.", cols=[
        ("SKU", S_, None, "none", ["isKey"]), ("Product", S_, None, "none", []), ("Category", S_, None, "none", []),
        ("List Price USD", D_, "#,0.00", "none", []), ("Unit Cost USD", D_, "#,0.00", "none", []),
        ("In Product List", B_, None, "none", [])]),
    "Customer": dict(src=CUSTOMER, desc="One row per customer; the cohort is the month of the first order.", cols=[
        ("Customer ID", S_, None, "none", ["isKey"]), ("First Order Date", DT_, "d mmm yyyy", "none", []), ("Channel", S_, None, "none", []),
        ("Country", S_, None, "none", []), ("Orders", I_, "0", "sum", []), ("Cohort No", I_, "0", "none", ["isHidden"]),
        ("Cohort", S_, None, "none", ["sortByColumn: 'Cohort No'"]), ("Repeat Customer", B_, None, "none", [])]),
    "Market": dict(src=MARKET, desc="Shipping country to sales market.", cols=[
        ("Country", S_, None, "none", ["isKey"]), ("Market", S_, None, "none", ["sortByColumn: 'Market Sort'"]),
        ("Market Sort", I_, "0", "none", ["isHidden"])]),
    "Currency": dict(src=CURRENCY, desc="The five currencies the store sells in.", cols=[
        ("Currency", S_, None, "none", ["isKey", "sortByColumn: Sort"]), ("Currency Name", S_, None, "none", []), ("Sort", I_, "0", "none", ["isHidden"])]),
    "Date": dict(src=DATE, desc="Calendar: one row per day of the months with orders.", dataCategory="Time", cols=[
        ("Date", DT_, "d mmm yyyy", "none", ["isKey"]), ("Year", I_, "0", "none", []),
        ("Month No", I_, "0", "none", ["isHidden"]), ("Month", S_, None, "none", ["sortByColumn: 'Month No'"]),
        ("Month Of Year", I_, "0", "none", ["isHidden"]), ("Month Name", S_, None, "none", ["sortByColumn: 'Month Of Year'"]),
        ("Month Start", DT_, "d mmm yyyy", "none", []), ("Quarter", S_, None, "none", [])]),
    "FX API": dict(src=FX_API, desc="Daily ECB rates from the Frankfurter API (1 USD = x units). Empty when the API cannot be reached.", hidden=True, cols=[
        ("Date", DT_, "d mmm yyyy", "none", []), ("Currency", S_, None, "none", []), ("Units per USD", D_, "0.00000", "none", [])]),
    "FX Saved": dict(src=FX_SAVED, desc="Saved copy of the rates (data/fx-rates-saved.csv).", hidden=True, cols=[
        ("Date", DT_, "d mmm yyyy", "none", []), ("Currency", S_, None, "none", []), ("Units per USD", D_, "0.00000", "none", [])]),
    "FX Rate": dict(calc=FX_RATE_DAX, desc="The rates the report uses: live from the API, the saved copy for any day the API did not return.", cols=[
        ("Date", DT_, "d mmm yyyy", "none", []), ("Currency", S_, None, "none", []), ("Units per USD", D_, "0.00000", "none", []),
        ("Source", S_, None, "none", []), ("Year", I_, "0", "none", []), ("Month Of Year", I_, "0", "none", []),
        ("USD per Unit", D_, "0.0000", "none", [])]),
    "Ad Spend": dict(src=ADS, desc="Monthly ad spend per ad channel (USD).", cols=[
        ("Month", DT_, "mmm yyyy", "none", []), ("Ad Channel", S_, None, "none", []), ("Spend USD", D_, "#,0.00", "sum", [])]),
    "Source Files": dict(src=SOURCE_FILES, desc="Control figures read straight from each export file.", cols=[
        ("Channel", S_, None, "none", []), ("File", S_, None, "none", []), ("Rows In File", I_, "#,0", "sum", []),
        ("Orders In File", I_, "#,0", "sum", []), ("Voided Orders", I_, "0", "sum", []), ("Voided Lines", I_, "0", "sum", []),
        ("Header Repeats", I_, "0", "sum", []), ("Duplicates", I_, "0", "sum", [])]),
    "File Totals": dict(src=FILE_TOTALS, desc="Order totals per currency read straight from the files.", cols=[
        ("Currency", S_, None, "none", []), ("Total In File", D_, "#,0.00", "sum", [])]),
    "Checks": dict(src=CHECKS, desc="The list of checks shown on the Checks page.", cols=[
        ("No", I_, "0", "none", []), ("Check", S_, None, "none", ["sortByColumn: No"]), ("Why it matters", S_, None, "none", [])]),
    "Insights": dict(src=INSIGHTS, desc="The rows of the Key insights panel; the sentence itself is the Insight measure.", cols=[
        ("No", I_, "0", "none", []), ("Topic", S_, None, "none", ["sortByColumn: No"])]),
    "Months Later": dict(src=MONTHS_LATER, desc="Columns of the cohort grid (0-12 months after the first order); not related, so the year slicer does not cut it.", cols=[
        ("Months Later", I_, "0", "none", [])]),
}
RELATIONSHIPS = [  # many side -> one side
    ("Sales", "Order ID", "Orders", "Order ID"), ("Sales", "SKU", "Product", "SKU"),
    ("Orders", "Date", "Date", "Date"), ("Orders", "Customer ID", "Customer", "Customer ID"),
    ("Orders", "Country", "Market", "Country"), ("Orders", "Currency", "Currency", "Currency"),
    ("FX Rate", "Currency", "Currency", "Currency"), ("Ad Spend", "Month", "Date", "Date"),
]
EXPRESSIONS = [  # shared queries, not loaded: name, M, result type
    ("Num", NUM, "Function"), ("NumEU", NUM_EU, "Function"), ("ShopRaw", SHOP_RAW, "Table"), ("ShopLines", SHOP_LINES, "Table"),
    ("ShopOrders", SHOP_ORDERS, "Table"), ("MarketLines", MARKET_LINES, "Table"), ("MarketOrders", MARKET_ORDERS, "Table"),
    ("ProductsRaw", PRODUCTS_RAW, "Table"),
]


def table_tmdl(name, t):
    out = []
    if t.get("desc"):
        out.append(f"/// {t['desc']}")
    out += [f"table {q(name)}"]
    if t.get("hidden"):
        out.append("\tisHidden")
    out.append(f"\tlineageTag: {tag('t', name)}")
    if t.get("dataCategory"):
        out.append(f"\tdataCategory: {t['dataCategory']}")
    out.append("")
    if t.get("measures"):
        for m, folder, fmt, expr, hidden in MEASURES:
            if "\n" in expr:
                out.append(f"\tmeasure {q(m)} =")
                out.append(indent(expr, 3))
            else:
                out.append(f"\tmeasure {q(m)} = {expr}")
            if fmt:
                out.append(f"\t\tformatString: {fmt}")
            out.append(f"\t\tdisplayFolder: {folder}")
            if hidden:
                out.append("\t\tisHidden")
            out.append(f"\t\tlineageTag: {tag('m', m)}")
            out.append("")
    calc = "calc" in t
    for c, dt, fmt, summ, extra in t["cols"]:
        out.append(f"\tcolumn {q(c)}")
        out.append(f"\t\tdataType: {dt}")
        if calc:
            out.append("\t\tisNameInferred")
        if fmt:
            out.append(f"\t\tformatString: {fmt}")
        for e in extra:
            out.append(f"\t\t{e}")
        if t.get("hidden"):
            out.append("\t\tisHidden")
        out.append(f"\t\tlineageTag: {tag('c', name, c)}")
        out.append(f"\t\tsummarizeBy: {summ}")
        out.append(f"\t\tsourceColumn: [{c}]" if calc else f"\t\tsourceColumn: {c}")
        out.append("")
        out.append("\t\tannotation SummarizationSetBy = Automatic")
        out.append("")
    for c, dt, fmt, expr in CALC_COLUMNS.get(name, []):
        if "\n" in expr:
            out.append(f"\tcolumn {q(c)} =")
            out.append(indent(expr, 3))
        else:
            out.append(f"\tcolumn {q(c)} = {expr}")
        out.append(f"\t\tdataType: {dt}")
        if fmt:
            out.append(f"\t\tformatString: {fmt}")
        out.append(f"\t\tlineageTag: {tag('c', name, c)}")
        out.append("\t\tsummarizeBy: none")
        out.append("")
        out.append("\t\tannotation SummarizationSetBy = Automatic")
        out.append("")
    if calc:
        out.append(f"\tpartition {q(name)} = calculated")
        out.append("\t\tmode: import")
        out.append("\t\tsource =")
        out.append(indent(t["calc"], 4))
        out.append("")
    else:
        out.append(f"\tpartition {q(name)} = m")
        out.append("\t\tmode: import")
        out.append("\t\tsource =")
        out.append(indent(t["src"], 4))
        out.append("")
        out.append("\tannotation PBI_ResultType = Table")
    out.append("")
    return "\n".join(out)


def build_model(data_folder, fx_start):
    if SM.exists():
        shutil.rmtree(SM)
    d = SM / "definition"
    (d / "tables").mkdir(parents=True)
    (SM / "definition.pbism").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2", "settings": {}}, indent=2), encoding="utf-8")
    (d / "database.tmdl").write_text("database\n\tcompatibilityLevel: 1601\n\n", encoding="utf-8")
    order = ["DataFolder", "FxStart"] + [e[0] for e in EXPRESSIONS] + [t for t in TABLES if "calc" not in TABLES[t]]
    model = ["model Model", "\tculture: en-US", "\tdefaultPowerBIDataSourceVersion: powerBI_V3", "\tdiscourageImplicitMeasures",
             "\tsourceQueryCulture: en-US", "\tdataAccessOptions", "\t\tlegacyRedirects", "\t\treturnErrorValuesAsNull", "",
             f"annotation PBI_QueryOrder = {json.dumps(order)}", "", "annotation __PBI_TimeIntelligenceEnabled = 0", "",
             "annotation PBI_ProTooling = [\"DevMode\"]", ""]
    model += [f"ref table {q(t)}" for t in TABLES]
    (d / "model.tmdl").write_text("\n".join(model) + "\n\n", encoding="utf-8")
    ex = [f'expression DataFolder = "{data_folder}" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]',
          f"\tlineageTag: {tag('e', 'DataFolder')}", "", "\tannotation PBI_ResultType = Text", "",
          f'expression FxStart = "{fx_start}" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]',
          f"\tlineageTag: {tag('e', 'FxStart')}", "", "\tannotation PBI_ResultType = Text", ""]
    for name, m, rtype in EXPRESSIONS:
        ex += [f"expression {name} =", indent(m, 2), f"\tlineageTag: {tag('e', name)}", "",
               f"\tannotation PBI_ResultType = {rtype}", ""]
    (d / "expressions.tmdl").write_text("\n".join(ex) + "\n", encoding="utf-8")
    rel = []
    for ft, fc, tt, tc in RELATIONSHIPS:
        rel += [f"relationship {tag('r', ft, fc, tt)}", f"\tfromColumn: {q(ft)}.{q(fc)}", f"\ttoColumn: {q(tt)}.{q(tc)}", ""]
    (d / "relationships.tmdl").write_text("\n".join(rel) + "\n", encoding="utf-8")
    # model diagram: the two fact tables in the middle, dimensions around them, helper tables below
    # rows 460 apart and columns 380 apart, so no box overlaps another (the tallest are capped at 420)
    spots = {"Ad Spend": (40, 40), "Date": (420, 40), "Customer": (800, 40),
             "Market": (40, 500), "Orders": (420, 500), "Sales": (800, 500), "Product": (1180, 500), "Months Later": (1560, 500),
             "Currency": (420, 960), "FX Rate": (800, 960), "FX API": (1180, 960), "FX Saved": (1560, 960),
             "_Measures": (40, 1420), "Source Files": (420, 1420), "File Totals": (800, 1420), "Checks": (1180, 1420),
             "Insights": (1560, 1420)}
    nodes = [{"location": {"x": x, "y": y}, "nodeIndex": n, "nodeLineageTag": tag("t", n),
              "size": {"height": min(64 + 34 * (len(TABLES[n]["cols"]) + len(CALC_COLUMNS.get(n, []))), 420), "width": 260}, "zIndex": 0}
             for n, (x, y) in spots.items()]
    (SM / "diagramLayout.json").write_text(json.dumps({
        "version": "1.1.0", "diagrams": [{"ordinal": 0, "scrollPosition": {"x": 0, "y": 0}, "nodes": nodes, "name": "All tables",
                                          "zoomValue": 70, "pinKeyFieldsToTop": False, "showExtraHeaderInfo": False,
                                          "hideKeyFieldsWhenCollapsed": False, "tablesLocked": False}],
        "selectedDiagram": "All tables", "defaultDiagram": "All tables"}, indent=2), encoding="utf-8")
    for name, t in TABLES.items():
        (d / "tables" / f"{name}.tmdl").write_text(table_tmdl(name, t), encoding="utf-8")


# ====================================================================================== report (PBIR)
S = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
# "Sand & Terracotta" light theme from E0: warm paper, white cards, terracotta accent, sage, amber and slate blue
PAGE_BG, HDR, CARD, CARD_ALT, LINE = "#FAF6F0", "#FFFDF9", "#FFFFFF", "#FBF7F1", "#E8DFD2"
ACCENT, ACCENT_LT, ACCENT_DIM = "#C2603D", "#E3A587", "#F6E3D9"
INK, MUTED, GRID = "#3B2F2A", "#7D6E63", "#EFE8DD"
SAGE, AMBER, SLATE, LY_GREY = "#6B8F71", "#D9A441", "#5E7FA3", "#D8CCBD"
GOOD, BAD, REVIEW = "#3F8A4E", "#C0392B", "#B7862B"
SERIF, SANS, SANS_SB = "Georgia", "Segoe UI", "Segoe UI Semibold"
NAV_X, NAV_W = 560, 700


def lit(v):
    if isinstance(v, bool):
        s = "true" if v else "false"
    elif isinstance(v, (int, float)):
        s = f"{v}D"
    else:
        s = "'" + str(v).replace("'", "''") + "'"
    return {"expr": {"Literal": {"Value": s}}}


def color(c):
    return {"solid": {"color": lit(c)}}


def mcolor(measure):  # colour returned by a measure
    return {"solid": {"color": {"expr": mea(measure)}}}


def col(entity, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def mea(prop):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": "_Measures"}}, "Property": prop}}


def proj(field, name=None):
    kind = "Column" if "Column" in field else "Measure"
    ent, prop = field[kind]["Expression"]["SourceRef"]["Entity"], field[kind]["Property"]
    p = {"field": field, "queryRef": f"{ent}.{prop}", "nativeQueryRef": prop}
    if name:
        p["displayName"] = name
    return p


def props(**kw):
    return [{"properties": kw}]


def container(title=None, bg=CARD, border=True, transparent=False, radius=10, border_color=LINE):
    o = {"background": props(show=lit(not transparent), color=color(bg), transparency=lit(100 if transparent else 0)),
         "border": props(show=lit(border), color=color(border_color), radius=lit(radius), width=lit(1)),
         "dropShadow": props(show=lit(False))}
    if title:
        o["title"] = props(show=lit(True), text=lit(title), fontColor=color(INK), fontSize=lit(12), bold=lit(False), fontFamily=lit(SERIF))
    else:
        o["title"] = props(show=lit(False))
    return o


class Page:
    def __init__(self, name, display):
        self.name, self.display, self.visuals = name, display, []

    def add(self, vid, x, y, w, h, visual):
        self.visuals.append({"$schema": f"{S}/visualContainer/2.14.0/schema.json", "name": vid,
                             "position": {"x": x, "y": y, "z": len(self.visuals) * 1000, "width": w, "height": h,
                                          "tabOrder": len(self.visuals) * 1000}, "visual": visual})


def textbox(page, vid, x, y, w, h, paras, bg=CARD, border=False, transparent=False, radius=10, pad=None, border_color=LINE):
    vco = container(bg=bg, border=border, transparent=transparent, radius=radius, border_color=border_color)
    if pad is not None:
        vco["padding"] = props(top=lit(pad[0]), left=lit(pad[1]), bottom=lit(0), right=lit(0))
    page.add(vid, x, y, w, h, {"visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": paras}}]},
                               "visualContainerObjects": vco})


def run(value, size, colour, font=SANS, bold=False):
    st = {"fontFamily": font, "fontSize": f"{size}pt", "color": colour}
    if bold:
        st["fontWeight"] = "bold"
    return {"value": value, "textStyle": st}


def header(page, title, sub):
    """Brand and page title on the left, page buttons on the right, a terracotta hairline underneath."""
    textbox(page, f"{page.name}_hdr", 0, 0, 1280, 60, [
        {"textRuns": [run("L A R K S P U R", 8, ACCENT, SANS_SB), run("     ", 8, MUTED), run(title, 18, INK, SERIF)]},
        {"textRuns": [run(sub, 8, MUTED)]}], bg=HDR, radius=0, pad=(4, 22))
    textbox(page, f"{page.name}_rule", 0, 60, 1280, 2, [{"textRuns": [run(" ", 1, ACCENT)]}], bg=ACCENT, radius=0, pad=(0, 0))
    page.add(f"{page.name}_nav", NAV_X, 15, NAV_W, 30, {
        "visualType": "pageNavigator",
        "objects": {
            "fill": [{"properties": {"show": lit(True), "fillColor": color(HDR), "transparency": lit(0)}, "selector": {"id": "default"}},
                     {"properties": {"show": lit(True), "fillColor": color(CARD_ALT), "transparency": lit(0)}, "selector": {"id": "hover"}},
                     {"properties": {"show": lit(True), "fillColor": color(ACCENT_DIM), "transparency": lit(0)}, "selector": {"id": "selected"}}],
            "text": [{"properties": {"show": lit(True), "fontColor": color(MUTED), "fontSize": lit(9), "fontFamily": lit(SANS)}, "selector": {"id": "default"}},
                     {"properties": {"fontColor": color(ACCENT)}, "selector": {"id": "selected"}}],
            "outline": [{"properties": {"show": lit(True), "lineColor": color(LINE), "weight": lit(1)}, "selector": {"id": "default"}},
                        {"properties": {"show": lit(True), "lineColor": color(ACCENT), "weight": lit(1)}, "selector": {"id": "selected"}}],
            "shape": props(tileShape=lit("rectangleRounded"), roundEdge=lit(15))},
        "visualContainerObjects": container(bg=HDR, border=False, transparent=True)})


def slicer(page, vid, x, y, w, field, title, sync=None, h=64, default=None):
    v = {"visualType": "slicer",
         "query": {"queryState": {"Values": {"projections": [proj(field)]}}},
         "objects": {"data": props(mode=lit("Dropdown")),
                     "header": props(show=lit(True), text=lit(title), fontColor=color(ACCENT), textSize=lit(8), fontFamily=lit(SANS_SB)),
                     "items": props(textSize=lit(10), fontColor=color(INK), background=color(CARD_ALT), outlineColor=color(LINE))},
         "visualContainerObjects": container()}
    if default is not None:  # a pre-selected value (single select)
        ent, prop = field["Column"]["Expression"]["SourceRef"]["Entity"], field["Column"]["Property"]
        v["objects"]["selection"] = props(singleSelect=lit(True))
        v["objects"]["general"] = [{"properties": {"filter": {"filter": {
            "Version": 2, "From": [{"Name": "s", "Entity": ent, "Type": 0}],
            "Where": [{"Condition": {"In": {"Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": prop}}],
                                            "Values": [[{"Literal": {"Value": f"{default}L" if isinstance(default, int) else "'" + default + "'"}}]]}}}]}}}}]
    if sync:
        v["syncGroup"] = {"groupName": sync, "fieldChanges": True, "filterChanges": True}
    page.add(vid, x, y, w, h, v)


def value_card(page, vid, x, y, w, h, measure, fill, size, label=None, font=SERIF):
    page.add(vid, x, y, w, h, {
        "visualType": "card",
        "query": {"queryState": {"Values": {"projections": [proj(mea(measure), label)]}}},
        "objects": {"labels": props(fontSize=lit(size), color=fill, labelDisplayUnits=lit(1), fontFamily=lit(font)),
                    "categoryLabels": props(show=lit(False))},
        "visualContainerObjects": container(transparent=True, border=False)})


def kpi(page, vid, x, y, w, h, measure, label, icon, note=None, note_color=None, value_color=INK, size=22):
    """One KPI tile: white card, terracotta bar on the left, ringed icon, small-caps label, serif value, coloured verdict line."""
    textbox(page, f"{vid}_bg", x, y, w, h, [{"textRuns": [run(" ", 6, INK)]}], border=True)
    textbox(page, f"{vid}_bar", x + 1, y + 16, 3, h - 32, [{"textRuns": [run(" ", 4, ACCENT)]}], bg=ACCENT, radius=2, pad=(0, 0))
    textbox(page, f"{vid}_ic", x + w - 40, y + 10, 28, 28,
            [{"textRuns": [run(icon, 10, ACCENT, SANS_SB)], "horizontalTextAlignment": "center"}],
            bg=ACCENT_DIM, radius=14, pad=(3, 0), border=True, border_color=ACCENT_LT)
    textbox(page, f"{vid}_lb", x + 14, y + 10, w - 58, 24, [{"textRuns": [run(label.upper(), 7, MUTED, SANS_SB)]}],
            transparent=True, pad=(2, 0))
    fill = color(value_color) if value_color.startswith("#") else mcolor(value_color)
    compact = h < 100
    if note:
        vy, vh = (y + 24, 34) if compact else (y + 30, 40)
    else:
        vy, vh = y + 30, h - 40
    value_card(page, f"{vid}_v", x + 8, vy, w - 16, vh, measure, fill, size - (2 if compact and note else 0), label)
    if note:
        nfill = mcolor(note_color) if note_color and not note_color.startswith("#") else color(note_color or MUTED)
        ny = y + h - (25 if compact else 30)
        value_card(page, f"{vid}_n", x + 8, ny, w - 16, 22 if compact else 26, note, nfill, 8, font=SANS_SB)


def sort(field, direction="Ascending"):
    return {"sort": [{"field": field, "direction": direction}], "isDefaultSort": False}


AXIS_CHARTS = {"lineClusteredColumnComboChart", "clusteredBarChart", "clusteredColumnChart", "lineChart", "areaChart", "waterfallChart",
               "columnChart", "stackedColumnChart", "barChart"}


def chart(page, vid, x, y, w, h, vtype, title, roles, sort_def=None, objects=None):
    v = {"visualType": vtype, "query": {"queryState": {r: {"projections": [proj(*f) if isinstance(f, tuple) else proj(f) for f in fs]}
                                                       for r, fs in roles.items()}},
         "visualContainerObjects": container(title)}
    if sort_def:
        v["query"]["sortDefinition"] = sort_def
    v["objects"] = objects or {}
    if vtype in AXIS_CHARTS:  # the title says what the axes are; faint gridlines, muted axis text
        v["objects"]["categoryAxis"] = [{"properties": {"showAxisTitle": lit(False), "labelColor": color(MUTED), "fontSize": lit(8)}
                                         | v["objects"].get("categoryAxis", [{}])[0].get("properties", {})}]
        v["objects"]["valueAxis"] = [{"properties": {"showAxisTitle": lit(False), "secShowAxisTitle": lit(False), "labelColor": color(MUTED),
                                                     "fontSize": lit(8), "gridlineColor": color(GRID), "gridlineStyle": lit("solid")}
                                      | v["objects"].get("valueAxis", [{}])[0].get("properties", {})}]
    if "legend" in v["objects"]:
        v["objects"]["legend"] = [{"properties": {"labelColor": color(INK)} | v["objects"]["legend"][0]["properties"]}]
    if "labels" in v["objects"]:
        first = v["objects"]["labels"][0]
        v["objects"]["labels"] = [first | {"properties": {"color": color(INK)} | first["properties"]}] + v["objects"]["labels"][1:]
    if not v["objects"]:
        del v["objects"]
    page.add(vid, x, y, w, h, v)
    return v


def labels(size=8, units=1, precision=None, **kw):
    p = dict(show=lit(True), labelDisplayUnits=lit(units), fontSize=lit(size))
    if precision is not None:
        p["labelPrecision"] = lit(precision)
    return props(**p, **kw)


def legend(pos="TopRight"):
    return props(show=lit(True), position=lit(pos), fontSize=lit(8))


def fills(*pairs):  # (measure name, colour) per series
    return [{"properties": {"fill": color(c)}, "selector": {"metadata": f"_Measures.{m}"}} for m, c in pairs]


def series_fills(entity, prop, pairs):  # (category value, colour) per series of a column
    return [{"properties": {"fill": color(c)},
             "selector": {"data": [{"scopeId": {"Comparison": {"ComparisonKind": 0, "Left": col(entity, prop),
                                                               "Right": {"Literal": {"Value": f"'{v}'"}}}}}]}} for v, c in pairs]


def cond(measure_queryref, prop, color_measure):
    return {"properties": {prop: {"solid": {"color": {"expr": mea(color_measure)}}}},
            "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": measure_queryref}}


def grid_objects(font=10, extra_values=None, pad=3):
    return {"values": [{"properties": {"fontSize": lit(font), "fontColorPrimary": color(INK), "fontColorSecondary": color(INK),
                                       "backColorPrimary": color(CARD), "backColorSecondary": color(CARD_ALT)}}] + (extra_values or []),
            "columnHeaders": props(fontSize=lit(font - 1), bold=lit(False), fontColor=color(ACCENT), backColor=color(HDR), fontFamily=lit(SANS_SB)),
            "rowHeaders": props(fontSize=lit(font), fontColor=color(INK), backColor=color(CARD)),
            "grid": props(rowPadding=lit(pad), gridHorizontal=lit(True), gridHorizontalColor=color(GRID), gridVertical=lit(False),
                          outlineColor=color(LINE)),
            "total": props(fontColor=color(INK), backColor=color(ACCENT_DIM)),
            "subTotals": props(fontColor=color(INK), backColor=color(CARD_ALT))}


NO_TOTALS = {"total": props(totals=lit(False))}
YEAR, MONTH_NAME, MONTH = col("Date", "Year"), col("Date", "Month Name"), col("Date", "Month")
CHANNEL, MARKET, CATEGORY, PRODUCT_ = col("Orders", "Channel"), col("Market", "Market"), col("Product", "Category"), col("Product", "Product")
SUB = "Larkspur Home (fictional)  ·  webshop + EU marketplace  ·  5 currencies, in USD"


def kpi_row(page, prefix, y, h, items, x0=274, x1=1258, gap=12, size=20):
    w = (x1 - x0 - gap * (len(items) - 1)) / len(items)
    for i, it in enumerate(items):
        m, lbl, ic = it[:3]
        note, ncol, vcol = (list(it[3:]) + [None, None, INK])[:3]
        kpi(page, f"{prefix}{i}", round(x0 + i * (w + gap)), y, round(w), h, m, lbl, ic, note, ncol, vcol or INK, size)


def year_slicer(page, prefix, x=22, w=150, h=60):
    slicer(page, f"{prefix}_year", x, 74, w, YEAR, "YEAR", "Year", h=h, default=2026)


def build_pages():
    # ---------------------------------------------------------------- 1 Overview
    p1 = Page("overview", "Overview")
    header(p1, "Sales overview", SUB)
    year_slicer(p1, "ov")
    slicer(p1, "ov_chan", 184, 74, 200, CHANNEL, "CHANNEL", "Channel", h=60)
    slicer(p1, "ov_mkt", 396, 74, 220, MARKET, "MARKET", "Market", h=60)
    textbox(p1, "ov_hint", 628, 74, 630, 60, [
        {"textRuns": [run("Every figure follows the year, channel and market you pick, compared with the same months a year before.  ", 9, MUTED),
                      run("Green", 9, GOOD, bold=True), run(" = better,  ", 9, MUTED), run("red", 9, BAD, bold=True), run(" = worse.", 9, MUTED)]}],
            transparent=True, pad=(12, 6))
    kpi_row(p1, "ov_k", 146, 100, [
        ("Net Sales", "Net sales (USD)", "$", "Sales Note", "Sales Note Color"),
        ("Orders", "Orders", "#", "Orders Note", "Orders Note Color"),
        ("Average Order Value", "Average order value", "Ø", "AOV Note", "AOV Note Color"),
        ("Gross Margin %", "Gross margin", "%", "Margin Note", "Margin Note Color"),
        ("Returning Share %", "Returning orders", "↺", "Returning Note", "Returning Note Color"),
        ("Refund Rate %", "Refunds", "↩", "Refund Note", "Refund Note Color")], x0=22, size=22)
    chart(p1, "ov_trend", 22, 258, 600, 236, "clusteredColumnChart", "Net sales by month: this year vs last year",
          {"Category": [MONTH_NAME], "Y": [(mea("Net Sales"), "This year"), (mea("Net Sales LY"), "Last year")]}, sort(MONTH_NAME),
          {"labels": labels(7, 1000, 0), "legend": legend(), "dataPoint": fills(("Net Sales", ACCENT), ("Net Sales LY", LY_GREY))})
    chart(p1, "ov_mkts", 634, 258, 300, 236, "clusteredBarChart", "Net sales by market",
          {"Category": [MARKET], "Y": [(mea("Net Sales"), "This year"), (mea("Net Sales LY"), "Last year")]}, sort(mea("Net Sales"), "Descending"),
          {"labels": labels(7, 1000, 0), "legend": props(show=lit(False)), "dataPoint": fills(("Net Sales", ACCENT), ("Net Sales LY", LY_GREY))})
    chart(p1, "ov_cat", 946, 258, 312, 236, "donutChart", "Net sales by category",
          {"Category": [CATEGORY], "Y": [(mea("Net Sales"), "Net sales")]}, sort(mea("Net Sales"), "Descending"),
          {"labels": props(show=lit(True), labelStyle=lit("Category, percent of total"), color=color(INK), fontSize=lit(8)),
           "legend": props(show=lit(False)), "slices": props(innerRadiusRatio=lit(60))})
    chart(p1, "ov_ytd", 22, 506, 480, 194, "lineChart", "Running total for the year vs last year",
          {"Category": [MONTH_NAME], "Y": [(mea("Net Sales YTD"), "This year"), (mea("Net Sales LY YTD"), "Last year")]}, sort(MONTH_NAME),
          {"legend": legend(), "dataPoint": fills(("Net Sales YTD", ACCENT), ("Net Sales LY YTD", LY_GREY)),
           "lineStyles": props(strokeWidth=lit(3), showMarker=lit(True), markerSize=lit(5))})
    chart(p1, "ov_ins", 514, 506, 744, 194, "tableEx", "Key insights, written by DAX",
          {"Values": [(col("Insights", "Topic"), "Topic"), (mea("Insight"), "What the numbers say")]}, sort(col("Insights", "Topic")),
          objects=grid_objects(9, [cond("_Measures.Insight", "fontColor", "Insight Color")], pad=1) | NO_TOTALS)
    # ---------------------------------------------------------------- 2 Products
    p2 = Page("products", "Products")
    header(p2, "Products", SUB)
    year_slicer(p2, "pr", h=84)
    slicer(p2, "pr_cat", 184, 74, 200, CATEGORY, "CATEGORY", "Category", h=84)
    kpi_row(p2, "pr_k", 74, 84, [("Units", "Units sold", "▦"), ("Net Sales", "Net sales (USD)", "$", "Sales Note", "Sales Note Color"),
                                 ("Gross Profit", "Gross profit", "="), ("Gross Margin %", "Gross margin", "%", "Margin Note", "Margin Note Color")],
            x0=396, size=20)
    mx = chart(p2, "pr_matrix", 22, 170, 760, 530, "pivotTable", "Category and product: units, sales, growth, margin, refunds",
               {"Rows": [CATEGORY, PRODUCT_], "Values": [(mea("Units"), "Units"), (mea("Net Sales"), "Net sales"),
                                                       (mea("Net Sales YoY %"), "vs last year"), (mea("Gross Margin %"), "Margin"),
                                                       (mea("Refund Rate %"), "Refunds"), (mea("Share of Net Sales %"), "Share")]},
               objects=grid_objects(9, pad=1))
    mx["expansionStates"] = [{"roles": ["Rows"], "levels": [{"queryRefs": ["Product.Category"], "identityKeys": [CATEGORY], "isPinned": True},
                                                            {"queryRefs": ["Product.Product"], "isPinned": True}],
                              "root": {"isToggled": True}}]
    chart(p2, "pr_top", 794, 170, 464, 300, "clusteredBarChart", "Top products by net sales",
          {"Category": [PRODUCT_], "Y": [(mea("Net Sales"), "Net sales")]}, sort(mea("Net Sales"), "Descending"),
          {"labels": labels(7, 1000, 1), "dataPoint": props(fill=color(ACCENT)), "categoryAxis": props(maxMarginFactor=lit(45))})
    chart(p2, "pr_cat_m", 794, 482, 464, 218, "columnChart", "Net sales by category and month",
          {"Category": [MONTH_NAME], "Series": [CATEGORY], "Y": [(mea("Net Sales"), "Net sales")]}, sort(MONTH_NAME),
          {"legend": props(show=lit(True), position=lit("Bottom"), fontSize=lit(8), showTitle=lit(False))})
    # ---------------------------------------------------------------- 3 Customers & marketing
    p3 = Page("customers", "Customers")
    header(p3, "Customers and marketing", SUB)
    year_slicer(p3, "cu", h=84)
    slicer(p3, "cu_chan", 184, 74, 200, CHANNEL, "CHANNEL", "Channel", h=84)
    kpi_row(p3, "cu_k", 74, 84, [("New Customers", "New customers", "+"),
                                 ("Returning Share %", "Returning orders", "↺", "Returning Note", "Returning Note Color"),
                                 ("Ad Spend", "Ad spend", "◎"), ("MER", "Marketing efficiency", "×", "MER Note", MUTED),
                                 ("CAC", "Cost / new customer", "¤", "CAC Note", MUTED)], x0=396, size=18)
    chart(p3, "cu_nr", 22, 170, 600, 250, "columnChart", "Orders by month: new vs returning customers",
          {"Category": [MONTH_NAME], "Series": [col("Orders", "Customer Type")], "Y": [(mea("Orders"), "Orders")]}, sort(MONTH_NAME),
          {"legend": legend(), "dataPoint": series_fills("Orders", "Customer Type", [("New", ACCENT_LT), ("Returning", ACCENT)])})
    chart(p3, "cu_ads", 634, 170, 624, 250, "lineClusteredColumnComboChart", "Ad spend and what it brought in (webshop)",
          {"Category": [MONTH_NAME], "Y": [(mea("Ad Spend"), "Ad spend")], "Y2": [(mea("MER"), "Sales per $1 of ads")]}, sort(MONTH_NAME),
          {"legend": legend(), "dataPoint": [{"properties": {"fill": color(SLATE)}}] + fills(("Ad Spend", SLATE), ("MER", ACCENT)),
           "lineStyles": props(strokeWidth=lit(3), showMarker=lit(True), markerSize=lit(5))})
    chart(p3, "cu_cohort", 22, 432, 1236, 268, "pivotTable",
          "2025 customer cohorts: share of each month's new customers who ordered again N months later (whatever year is picked)",
          {"Rows": [(col("Customer", "Cohort"), "First order")], "Columns": [(col("Months Later", "Months Later"), "Months later")],
           "Values": [(mea("Cohort Retention %"), "Ordered")]},
          objects=grid_objects(8, [cond("_Measures.Cohort Retention %", "backColor", "Cohort Color"),
                                   cond("_Measures.Cohort Retention %", "fontColor", "Cohort Font")], pad=1)
          | {"total": props(rowTotals=lit(False), columnTotals=lit(False)), "subTotals": props(rowSubtotals=lit(False), columnSubtotals=lit(False))})
    # ---------------------------------------------------------------- 4 Markets & FX
    p4 = Page("fx", "Markets & FX")
    header(p4, "Markets and exchange rates", "Daily ECB rates from the Frankfurter API  ·  constant currency = last year's rates")
    year_slicer(p4, "fx", h=84)
    kpi_row(p4, "fx_k", 74, 84, [("Non-USD Share %", "Sales outside the US dollar", "€", "Non-USD Note", MUTED),
                                 ("CC Growth %", "Growth at constant currency", "↗", "CC Note", MUTED),
                                 ("FX Effect", "What exchange rates added", "±", "FX Note", MUTED, "FX Color"),
                                 ("Rate Source", "Exchange rates", "⟳", "Rate Source Note", MUTED, "Rate Source Color")], x0=184, size=17)
    chart(p4, "fx_table", 22, 170, 760, 230, "tableEx", "By currency: local sales, the rate used, USD, and at last year's rates",
          {"Values": [(col("Currency", "Currency"), "Currency"), (mea("Net Sales Local"), "Net sales (local)"),
                      (mea("Effective Rate"), "Rate used (per USD)"), (mea("Net Sales"), "Net sales USD"),
                      (mea("Net Sales CC"), "At last year's rates"), (mea("FX Effect"), "FX effect"), (mea("Latest Rate"), "Latest rate")]},
          sort(col("Currency", "Currency")), objects=grid_objects(9, [cond("_Measures.FX Effect", "fontColor", "FX Color")], pad=2) | NO_TOTALS)
    chart(p4, "fx_mkt", 794, 170, 464, 230, "clusteredBarChart", "Net sales by market: this year vs last year",
          {"Category": [MARKET], "Y": [(mea("Net Sales"), "This year"), (mea("Net Sales LY"), "Last year")]}, sort(mea("Net Sales"), "Descending"),
          {"labels": labels(7, 1000, 0), "legend": legend(), "dataPoint": fills(("Net Sales", ACCENT), ("Net Sales LY", LY_GREY))})
    chart(p4, "fx_line", 22, 412, 760, 288, "lineChart", "What one unit of each currency is worth in USD (first day = 100), daily ECB rates",
          {"Category": [col("FX Rate", "Date")], "Series": [col("FX Rate", "Currency")], "Y": [(mea("FX Index"), "Index")]},
          sort(col("FX Rate", "Date")),
          {"legend": legend(), "categoryAxis": props(axisType=lit("Scalar")), "lineStyles": props(strokeWidth=lit(2)),
           "dataPoint": series_fills("FX Rate", "Currency", [("EUR", ACCENT), ("GBP", SLATE), ("CAD", SAGE), ("AUD", AMBER)]),
           "valueAxis": props(start=lit(85))})
    chart(p4, "fx_month", 794, 412, 464, 288, "clusteredColumnChart", "FX effect by month (USD)",
          {"Category": [MONTH_NAME], "Y": [(mea("FX Effect"), "FX effect")]}, sort(MONTH_NAME),
          {"labels": labels(7, 1, 0), "dataPoint": [{"properties": {"fill": mcolor("FX Color")},
                                                     "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}}]})
    # ---------------------------------------------------------------- 5 Checks
    p5 = Page("checks", "Checks")
    header(p5, "Can the numbers be trusted?", "Recalculated on every refresh, independent of the slicers  ·  a second read of every export file")
    kpi_row(p5, "ck_k", 74, 88, [("Checks Passed", "Checks", "✓", None, None, GOOD), ("Checks Failed", "Failed", "✕"),
                                 ("Chk Unknown SKU", "Lines waiting for a product", "?", None, None, REVIEW),
                                 ("Rate Source", "Exchange rates", "⟳", "Rate Source Note", MUTED, "Rate Source Color")], x0=22, x1=1258, size=20)
    chart(p5, "ck_table", 22, 174, 1236, 380, "tableEx", "Data checks",
          {"Values": [(col("Checks", "No"), "#"), (col("Checks", "Check"), "Check"), (col("Checks", "Why it matters"), "Why it matters"),
                      (mea("Check Detail"), "Result"), (mea("Check Status"), "Status")]},
          sort(col("Checks", "No")), objects=grid_objects(9, [cond("_Measures.Check Status", "fontColor", "Check Color")]) | NO_TOTALS)
    chart(p5, "ck_review", 22, 566, 1236, 134, "tableEx", "Lines waiting for a product (add the SKU to products.csv, then refresh)",
          {"Values": [(col("Sales", "Order ID"), "Order"), (col("Sales", "Channel"), "Channel"), (col("Sales", "SKU As Exported"), "SKU as exported"),
                      (col("Sales", "Line ID"), "Line"), (mea("Unknown SKU Lines"), "Lines")]},
          sort(col("Sales", "Order ID")), objects=grid_objects(9) | NO_TOTALS)
    return [p1, p2, p3, p4, p5]


def build_report():
    if RP.exists():
        shutil.rmtree(RP)
    d = RP / "definition"
    (d / "pages").mkdir(parents=True)
    (RP / "definition.pbir").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0", "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}}, indent=2), encoding="utf-8")
    (d / "version.json").write_text(json.dumps({"$schema": f"{S}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"}, indent=2),
                                    encoding="utf-8")
    base = "CY25SU11"
    theme_src = ROOT / "scripts" / "theme" / f"{base}.json"
    (RP / "StaticResources" / "SharedResources" / "BaseThemes").mkdir(parents=True)
    shutil.copyfile(theme_src, RP / "StaticResources" / "SharedResources" / "BaseThemes" / f"{base}.json")
    (RP / "StaticResources" / "RegisteredResources").mkdir(parents=True)
    custom = {"name": "Larkspur Sand & Terracotta", "dataColors": [ACCENT, SLATE, SAGE, AMBER, "#9C6B98", ACCENT_LT, "#8FA8C8", "#A7C4A0"],
              "foreground": INK, "foregroundNeutralSecondary": MUTED, "foregroundNeutralTertiary": "#A89A8E",
              "background": CARD, "backgroundLight": CARD_ALT, "backgroundNeutral": LINE, "tableAccent": ACCENT,
              "good": GOOD, "bad": BAD, "neutral": AMBER, "maximum": ACCENT, "center": ACCENT_LT, "minimum": CARD,
              "textClasses": {"callout": {"fontFace": SERIF, "color": INK}, "title": {"fontFace": SERIF, "color": INK},
                              "header": {"fontFace": SANS_SB, "color": ACCENT}, "label": {"fontFace": SANS, "color": INK}},
              "visualStyles": {"page": {"*": {"background": [{"color": {"solid": {"color": PAGE_BG}}, "transparency": 0}]}},
                               "*": {"*": {"visualTooltip": [{"titleFontColor": ACCENT, "valueFontColor": INK, "background": {"solid": {"color": CARD}}}]}}}}
    (RP / "StaticResources" / "RegisteredResources" / "Larkspur.json").write_text(json.dumps(custom, indent=2), encoding="utf-8")
    ver = {"visual": "2.14.0", "report": "3.3.0", "page": "2.1.0"}
    report = {"$schema": f"{S}/report/3.3.0/schema.json",
              "themeCollection": {"baseTheme": {"name": base, "reportVersionAtImport": ver, "type": "SharedResources"},
                                  "customTheme": {"name": "Larkspur.json", "reportVersionAtImport": ver, "type": "RegisteredResources"}},
              "resourcePackages": [{"name": "SharedResources", "type": "SharedResources",
                                    "items": [{"name": base, "path": f"BaseThemes/{base}.json", "type": "BaseTheme"}]},
                                   {"name": "RegisteredResources", "type": "RegisteredResources",
                                    "items": [{"name": "Larkspur.json", "path": "Larkspur.json", "type": "CustomTheme"}]}],
              "settings": {"useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarized",
                           "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True, "useEnhancedTooltips": True,
                           "useDefaultAggregateDisplayName": True},
              "objects": {"outspacePane": props(expanded=lit(False))}}  # the filter pane starts collapsed
    (d / "report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    pages = build_pages()
    (d / "pages" / "pages.json").write_text(json.dumps({"$schema": f"{S}/pagesMetadata/1.1.0/schema.json",
                                                         "pageOrder": [p.name for p in pages], "activePageName": pages[0].name},
                                                        indent=2), encoding="utf-8")
    for p in pages:
        pd = d / "pages" / p.name
        (pd / "visuals").mkdir(parents=True)
        (pd / "page.json").write_text(json.dumps({"$schema": f"{S}/page/2.1.0/schema.json", "name": p.name, "displayName": p.display,
                                                  "displayOption": "FitToPage", "height": 720, "width": 1280,
                                                  "objects": {"background": props(color=color(PAGE_BG), transparency=lit(0))}},
                                                 indent=2), encoding="utf-8")
        for v in p.visuals:
            (pd / "visuals" / v["name"]).mkdir()
            (pd / "visuals" / v["name"] / "visual.json").write_text(json.dumps(v, indent=2, ensure_ascii=False), encoding="utf-8")
    return pages


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-folder", default=r"C:\Larkspur\data")
    ap.add_argument("--fx-start", default="2024-12-01")
    a = ap.parse_args()
    folder = str(Path(a.data_folder).resolve()).rstrip("\\/") + "\\"
    build_model(folder, a.fx_start)
    pages = build_report()
    (ROOT / f"{NAME}.pbip").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}], "settings": {"enableAutoRecovery": True}},
        indent=2), encoding="utf-8")
    ncalc = sum(len(v) for v in CALC_COLUMNS.values())
    print(f"{NAME}.pbip: {len(TABLES)} tables, {len(MEASURES)} measures, {ncalc} calculated columns, {len(RELATIONSHIPS)} relationships, "
          f"{len(pages)} pages, {sum(len(p.visuals) for p in pages)} visuals; data folder {folder}")


if __name__ == "__main__":
    main()
