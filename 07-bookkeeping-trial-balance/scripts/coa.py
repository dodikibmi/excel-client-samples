"""The client's set-up for "Copperleaf Bakery & Cafe" (fictional): chart of accounts with opening balances,
categorisation rules, monthly budget and adjusting entries. build_workbook.py writes these into the workbook's
input tabs; verify.py uses the same rules to recompute everything in plain Python. make_data.py does NOT use the
rules: it knows the true account of every line, so the rules are tested against a planted answer."""
from datetime import date

MONTHS = [(202608, "Aug 2026", date(2026, 8, 31)), (202609, "Sep 2026", date(2026, 9, 30)), (202610, "Oct 2026", date(2026, 10, 31))]

# number, name, type, opening debit, opening credit (balances at 31 Jul 2026)
ACCOUNTS = [
    (1000, "Business checking", "Asset", 18420.55, 0),
    (1010, "Business savings", "Asset", 25000.00, 0),
    (1090, "Transfers in transit", "Asset", 0, 0),
    (1500, "Kitchen equipment", "Asset", 48000.00, 0),
    (1510, "Accumulated depreciation", "Asset", 0, 6000.00),
    (2000, "Visa business card", "Liability", 0, 2315.40),
    (3000, "Owner's equity", "Equity", 0, 83105.15),
    (3100, "Owner's draws", "Equity", 0, 0),
    (4000, "Counter sales (card and cash)", "Revenue", 0, 0),
    (4100, "Wholesale and catering sales", "Revenue", 0, 0),
    (4900, "Interest income", "Revenue", 0, 0),
    (5000, "Ingredients", "COGS", 0, 0),
    (5100, "Packaging", "COGS", 0, 0),
    (6000, "Payroll and payroll taxes", "Expense", 0, 0),
    (6100, "Rent", "Expense", 0, 0),
    (6200, "Utilities and internet", "Expense", 0, 0),
    (6400, "Marketing", "Expense", 0, 0),
    (6500, "Insurance", "Expense", 0, 0),
    (6600, "Repairs and maintenance", "Expense", 0, 0),
    (6700, "Software subscriptions", "Expense", 0, 0),
    (6800, "Bank and card fees", "Expense", 0, 0),
    (6900, "Depreciation", "Expense", 0, 0),
    (6950, "Office and sundry", "Expense", 0, 0),
    (6990, "Uncategorized (to review)", "Expense", 0, 0),
]
REGISTERS = {1000: "Business checking ...1180", 1010: "Business savings ...4821", 2000: "Visa business card ...7702"}
CLEARING, UNCATEGORIZED = 1090, 6990

# keyword (matched inside the cleaned, upper-case description), account. First match wins: specific rules first.
RULES = [
    ("TAPPAY MONTHLY FEE", 6800, "Processor fees: must sit above the TAPPAY deposit rule"),
    ("TAPPAY", 4000, "Daily card takings, net of nothing (fees are charged monthly)"),
    ("TAP PAY", 4000, "Same processor, typed with a space"),
    ("CAFE BLOOM", 4100, "Wholesale customer"),
    ("BLUE HERON", 4100, "Wholesale customer"),
    ("GREENLEAF", 4100, "Catering customer"),
    ("INTEREST", 4900, "Savings interest"),
    ("NORTHGATE", 5000, "Main food supplier"),
    ("VALLEY FLOUR", 5000, "Flour mill"),
    ("COSTWISE", 5000, "Cash-and-carry top-ups"),
    ("PACKRIGHT", 5100, "Boxes and cups"),
    ("PACK RIGHT", 5100, "Same supplier, typed with a space"),
    ("PAYPILOT", 6000, "Payroll provider: net pay and tax impounds"),
    ("MAPLE ST", 6100, "Landlord"),
    ("CITY POWER", 6200, "Electricity"),
    ("CITY PWR", 6200, "Electricity (autopay text)"),
    ("METRO WATER", 6200, "Water"),
    ("BRIGHTWAVE", 6200, "Internet"),
    ("PIXELREACH", 6400, "Social media ads"),
    ("SHIELD MUTUAL", 6500, "Business insurance"),
    ("OVEN DOCTOR", 6600, "Equipment repairs"),
    ("CLOUDPOS", 6700, "Point-of-sale software"),
    ("LEDGERLY", 6700, "Bookkeeping software"),
    ("SERVICE FEE", 6800, "Bank fee"),
    ("SERVICE CHARGE", 6800, "Bank fee (other wording)"),
    ("PAPERCLIP", 6950, "Office supplies"),
    ("TO PERSONAL", 3100, "Owner pays themself: a draw, not an expense"),
    ("TRANSFER TO SAV", CLEARING, "Checking to savings: paired with the savings side"),
    ("TRANSFER FROM CHK", CLEARING, "Savings side of the same transfer"),
    ("PAYMENT TO VISA", CLEARING, "Card payment from checking: paired with the card side"),
    ("PAYMENT THANK YOU", CLEARING, "Card side of the same payment"),
]

BUDGET = {  # per month: Aug, Sep, Oct
    4000: (30500, 31500, 33000), 4100: (3500, 3500, 4000), 4900: (15, 15, 15),
    5000: (10800, 10500, 11000), 5100: (900, 900, 950),
    6000: (13000, 13000, 13000), 6100: (4200, 4200, 4200), 6200: (950, 900, 900), 6400: (250, 250, 300),
    6500: (385, 385, 385), 6600: (150, 150, 150), 6700: (100, 100, 100), 6800: (1050, 1000, 1050),
    6900: (500, 500, 500), 6950: (75, 75, 75), 6990: (0, 0, 0),
}

# date, entry, description, account, debit, credit
ADJUSTMENTS = []
for _, label, end in MONTHS:
    ADJUSTMENTS += [(end, f"ADJ-{end:%m}", f"Depreciation {label}: kitchen equipment, 8 years straight line", 6900, 500.00, 0),
                    (end, f"ADJ-{end:%m}", f"Depreciation {label}: kitchen equipment, 8 years straight line", 1510, 0, 500.00)]
