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

# keyword (matched inside the cleaned, upper-case description), account, why, payee (who gets / pays the money).
# First match wins: specific rules first.
RULES = [
    ("TAPPAY MONTHLY FEE", 6800, "Processor fees: must sit above the TAPPAY deposit rule", "TapPay (processor fees)"),
    ("TAPPAY", 4000, "Daily card takings, net of nothing (fees are charged monthly)", "TapPay (card takings)"),
    ("TAP PAY", 4000, "Same processor, typed with a space", "TapPay (card takings)"),
    ("CAFE BLOOM", 4100, "Wholesale customer", "Cafe Bloom"),
    ("BLUE HERON", 4100, "Wholesale customer", "Blue Heron"),
    ("GREENLEAF", 4100, "Catering customer", "Greenleaf Catering"),
    ("INTEREST", 4900, "Savings interest", "Savings interest"),
    ("NORTHGATE", 5000, "Main food supplier", "Northgate Foodservice"),
    ("VALLEY FLOUR", 5000, "Flour mill", "Valley Flour Mill"),
    ("COSTWISE", 5000, "Cash-and-carry top-ups", "Costwise"),
    ("PACKRIGHT", 5100, "Boxes and cups", "PackRight"),
    ("PACK RIGHT", 5100, "Same supplier, typed with a space", "PackRight"),
    ("PAYPILOT", 6000, "Payroll provider: net pay and tax impounds", "PayPilot (payroll)"),
    ("MAPLE ST", 6100, "Landlord", "Maple St Properties"),
    ("CITY POWER", 6200, "Electricity", "City Power & Light"),
    ("CITY PWR", 6200, "Electricity (autopay text)", "City Power & Light"),
    ("METRO WATER", 6200, "Water", "Metro Water"),
    ("BRIGHTWAVE", 6200, "Internet", "Brightwave Internet"),
    ("PIXELREACH", 6400, "Social media ads", "PixelReach Ads"),
    ("SHIELD MUTUAL", 6500, "Business insurance", "Shield Mutual"),
    ("OVEN DOCTOR", 6600, "Equipment repairs", "Oven Doctor"),
    ("CLOUDPOS", 6700, "Point-of-sale software", "CloudPOS"),
    ("LEDGERLY", 6700, "Bookkeeping software", "Ledgerly"),
    ("SERVICE FEE", 6800, "Bank fee", "Harbor Community Bank"),
    ("SERVICE CHARGE", 6800, "Bank fee (other wording)", "Harbor Community Bank"),
    ("PAPERCLIP", 6950, "Office supplies", "Paperclip Office"),
    ("TO PERSONAL", 3100, "Owner pays themself: a draw, not an expense", "Owner"),
    ("TRANSFER TO SAV", CLEARING, "Checking to savings: paired with the savings side", "Own accounts"),
    ("TRANSFER FROM CHK", CLEARING, "Savings side of the same transfer", "Own accounts"),
    ("PAYMENT TO VISA", CLEARING, "Card payment from checking: paired with the card side", "Own accounts"),
    ("PAYMENT THANK YOU", CLEARING, "Card side of the same payment", "Own accounts"),
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

# The owner's plan for the 13-week cash forecast (same as E9's cash-plan.csv). The amount comes from the bank history
# unless the plan types it (fixed): avg_week = average week of the last N weeks, avg_month = average month of the
# last N months, avg_event = average of the last N dates with a posting (pay runs). Money in +, money out -.
# order, item, cash flow line, register, accounts, basis, lookback, frequency, first date, fixed amount, note
CASH_PLAN = [
    (1, "Customer receipts (counter and wholesale)", "Receipts from customers", 1000, "4000;4100", "avg_week", 8, "weekly", date(2026, 11, 1), None,
     "Card takings and wholesale customers: the average week of the last 8 weeks"),
    (2, "Savings interest", "Receipts from customers", 1010, "4900", "avg_month", 3, "monthly", date(2026, 11, 30), None,
     "Paid on the last day of the month"),
    (3, "Payroll and payroll taxes", "Payments to suppliers and staff", 1000, "6000", "avg_event", 6, "biweekly", date(2026, 11, 13), None,
     "PayPilot every second Friday: net pay + tax impound; average of the last 6 pay runs"),
    (4, "Food suppliers (bank)", "Payments to suppliers and staff", 1000, "5000", "avg_week", 8, "weekly", date(2026, 11, 1), None,
     "Northgate invoices paid from checking"),
    (5, "Rent", "Payments to suppliers and staff", 1000, "6100", "fixed", None, "monthly", date(2026, 11, 1), -4200, "Lease: 4200 on the 1st"),
    (6, "Insurance", "Payments to suppliers and staff", 1000, "6500", "fixed", None, "monthly", date(2026, 11, 5), -385, "Premium on the 5th"),
    (7, "Utilities (power and water)", "Payments to suppliers and staff", 1000, "6200", "avg_month", 3, "monthly", date(2026, 11, 14), None,
     "Autopay mid-month"),
    (8, "Bank and processor fees", "Payments to suppliers and staff", 1000, "6800", "avg_month", 3, "monthly", date(2026, 11, 28), None,
     "TapPay monthly fee + bank service fee"),
    (9, "Card payment", "Card payments", 2000, "5000;5100;6200;6400;6600;6700;6950;6990", "avg_month", 3, "monthly", date(2026, 11, 12), None,
     "Pays the card spend: the average month of card purchases"),
    (10, "Owner's draws", "Owner's draws", 1000, "3100", "fixed", None, "monthly", date(2026, 11, 20), -3000, "Owner pays themself 3000 on the 20th"),
]
CASH_BUFFER = 10000  # minimum cash the owner wants to keep (editable on the Cash flow sheet)
