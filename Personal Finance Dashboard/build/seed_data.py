"""Seed data for the Personal Finance Planner workbook: real Accounts/Categories/Vendors
supplied by the user, with no sample transactions — a clean slate to start entering into.
"""
import datetime as dt

TODAY = dt.date(2026, 7, 29)

def months_back(n):
    y, m = TODAY.year, TODAY.month
    m -= n
    while m <= 0:
        m += 12
        y -= 1
    return y, m

def d(y, m, day):
    day = min(day, 28)
    return dt.date(y, m, day)

# ---------------- Dim_Category (flat — no subcategories) ----------------
# CategoryID, CategoryName, CategoryType, EssentialCategory
CATEGORIES = [
    (1, "Mortgage", "Expense", "Yes"),
    (2, "Utilities", "Expense", "Yes"),
    (3, "Insurance Home", "Expense", "Yes"),
    (4, "Car Lease/Finance", "Expense", "Yes"),
    (5, "Car Expenses", "Expense", "Yes"),
    (6, "Insurance Car", "Expense", "Yes"),
    (7, "Gas", "Expense", "Yes"),
    (8, "Groceries", "Expense", "Yes"),
    (9, "Dining & Restaurants", "Expense", "No"),
    (10, "Healthcare", "Expense", "Yes"),
    (11, "Entertainment", "Expense", "No"),
    (12, "Personal Care", "Expense", "No"),
    (13, "Men's Clothing", "Expense", "No"),
    (14, "Women's Clothing", "Expense", "No"),
    (15, "Kid's Clothing", "Expense", "No"),
    (16, "Kid's Accessories", "Expense", "No"),
    (17, "Gifts", "Expense", "No"),
    (18, "Toys", "Expense", "No"),
    (19, "Cleaning Supplies", "Expense", "No"),
    (20, "Amazon", "Expense", "No"),
    (21, "Walmart", "Expense", "No"),
    (22, "Target", "Expense", "No"),
    (23, "Books", "Expense", "No"),
    (24, "Legal Fees", "Expense", "Yes"),
    (25, "Taxes", "Expense", "Yes"),
    (26, "Payroll Direct DB", "Income", "No"),
    (27, "Payroll Cash RB", "Income", "No"),
    (28, "Payroll Direct RB", "Income", "No"),
]
_CAT_ID = {name: cid for (cid, name, _typ, _ess) in CATEGORIES}

# ---------------- Dim_Vendor ----------------
# VendorID, VendorName, DefaultCategoryID
# DefaultCategoryID is None where nothing in the category list above is a confident fit
# (mainly: schools/yeshivas, shuls/congregations, tzedakah & charity organizations, bank
# fees/interest line items, home repair, and Zelle transfers — none of these have a matching
# category yet). Those vendors still work everywhere; they just won't auto-suggest a category
# in Bank Import until you either pick one manually or add a category that fits.
_VENDOR_DEFS = [
    ("99 CENTS OUTLET", None),
    ("AIRCO MECHANICAL", None),
    ("AMAZING SAVINGS", None),
    ("AMAZON", "Amazon"),
    ("AMAZON PRIME", "Amazon"),
    ("ASI / PROGRESSIVE INSURANCE", "Insurance Car"),
    ("BABY DREAMS", "Kid's Accessories"),
    ("BAIS HASFORIM", "Books"),
    ("BETH ROCHEL SCHOOL", None),
    ("BINGO WHOLESALE", "Groceries"),
    ("BNEI YAKOV YOSEF OF MONSEY", None),
    ("CAFE CHOCOLAT", "Dining & Restaurants"),
    ("CAFE CORNER", "Dining & Restaurants"),
    ("CELL 2 GET", "Utilities"),
    ("CHAI LIFELINE", None),
    ("CHURRASKO GRILL", "Dining & Restaurants"),
    ("CLAIRE'S", "Kid's Accessories"),
    ("CM WINDOWS & DOORS", None),
    ("CONG TALMUDEI ISRAEL", None),
    ("Cong Zwehil Of Monsey", None),
    ("CONG. BIRKAS SHESH", None),
    ("CONG. BUILDING EMANUEL", None),
    ("CONG. KIPAS CHASANIM", None),
    ("COPE INSTITUTE", None),
    ("COSTUME CENTRAL", "Toys"),
    ("CROSSROADS WINE & SPIRITS", "Entertainment"),
    ("CVS PHARMACY", "Healthcare"),
    ("DAVE & BUSTERS", "Entertainment"),
    ("Der Shtiebel", None),
    ("DOLLAR TREE", None),
    ("DSW", "Women's Clothing"),
    ("Eli Meisels Inc", None),
    ("ERIE INSURANCE", None),
    ("EVERGREEN KOSHER MARKET", "Groceries"),
    ("EXXON GAS", "Gas"),
    ("E-Z PASS", "Car Expenses"),
    ("EZER YESHIVA SEDER RAFFLE", None),
    ("FAME CN CENTRAL", None),
    ("FIVE BELOW", "Toys"),
    ("FOREIGN TRANSACTION FEE", None),
    ("FRANKEL'S DESIGNER SHOES", "Women's Clothing"),
    ("GEICO", "Insurance Car"),
    ("GOOGLE", None),
    ("GRAPE WINE & SPIRIT", "Entertainment"),
    ("H&M", "Women's Clothing"),
    ("HAVA JAVA", "Dining & Restaurants"),
    ("HEAVEN SCENT", "Gifts"),
    ("HIVE DISCOUNT", None),
    ("IKEA", None),
    ("INTEREST PAYMENT", None),
    ("INVITE WITH CLASS", "Gifts"),
    ("IRS", "Taxes"),
    ("JOSEPH DANITTI - CUFF & CO", "Gifts"),
    ("KAYX", None),
    ("KEREN CHASANIM", None),
    ("KIDICHIC", "Kid's Clothing"),
    ("KIDS FIRST PEO", "Healthcare"),
    ("KRISPY BY GG", "Dining & Restaurants"),
    ("LATE FEE", None),
    ("LE BRICK", None),
    ("LILY & TODD", "Kid's Clothing"),
    ("LILY AND TODD", "Kid's Clothing"),
    ("LOWES", None),
    ("LUIBELLE", "Women's Clothing"),
    ("LUKOIL", "Gas"),
    ("LULU KIDS CLOTHING", "Kid's Clothing"),
    ("MACY'S", "Women's Clothing"),
    ("MALBISH", None),
    ("MELT", "Dining & Restaurants"),
    ("MERKAZ SEFORIM", "Books"),
    ("METRO BY T-MOBILE", "Utilities"),
    ("MIDAS - MONSEY", "Car Expenses"),
    ("Mikes Burger", "Dining & Restaurants"),
    ("MONSEY GLATT", "Groceries"),
    ("MONSEY URGENT CARE", "Healthcare"),
    ("MONSEY WINE & LIQUOR", "Entertainment"),
    ("MONTHLY SERVICE FEE", None),
    ("MONTVALE WINE LIQUOR", "Entertainment"),
    ("MORC INC 8453710211", None),
    ("MUNCH HEARTY", "Dining & Restaurants"),
    ("NAME CHEAP", None),
    ("NEW YORK STATE DMV", "Car Expenses"),
    ("NEWDAY", None),
    ("Nissan Auto Lease", "Car Lease/Finance"),
    ("NOD", None),
    ("NORDSTROM RACK", "Women's Clothing"),
    ("NORWICH COMMERCIAL", None),
    ("NU TREND CLEANERS", "Personal Care"),
    ("NUMBER BARN", "Utilities"),
    ("NY STATE TAX", "Taxes"),
    ("NYS TAX", "Taxes"),
    ("OHR CHAIM - MIKVA", None),
    ("OHR CHAIM CHARITY CAMPAIG", None),
    ("OLYMPIA PITA", "Dining & Restaurants"),
    ("OORAH", None),
    ("ORANGE & ROCKLAND", "Utilities"),
    ("OVERDRAFT FEE", None),
    ("PHARMACY PLUS", "Healthcare"),
    ("PIES N FRIES", "Dining & Restaurants"),
    ("PITA LAND", "Dining & Restaurants"),
    ("PRIM", None),
    ("Prime Video", "Entertainment"),
    ("PURCHASE INTEREST CHARGE", None),
    ("PYRAMID PLUMBING", None),
    ("RAILWAY.COM", None),
    ("ROCKET MORTGAGE", "Mortgage"),
    ("ROCKLAND KOSHER", "Groceries"),
    ("SANDERS BAKERY", "Dining & Restaurants"),
    ("SEASONS EXPRESS", "Groceries"),
    ("SEPHORA", "Personal Care"),
    ("SHAWARMA DELIGHT", "Dining & Restaurants"),
    ("SHELIS IN TOWN SQ", "Dining & Restaurants"),
    ("SHELL OIL", "Gas"),
    ("SHOPRITE", "Groceries"),
    ("SLONIM", None),
    ("SOCK SHOPPE", None),
    ("SPARK CAR WASH", "Car Expenses"),
    ("SPRINKLES", "Dining & Restaurants"),
    ("SW DESIGNER SHOES", "Women's Clothing"),
    ("SWADDLES BABY", "Kid's Accessories"),
    ("SWEET EXPRESSIONS", "Gifts"),
    ("TALMED TORAH IMREI BINA", None),
    ("TALMUDEI YISROEL STANISLUV", None),
    ("TARGET", "Target"),
    ("TEN YAD KALLAH GEMACH", None),
    ("THE CHILDRENS PLACE", "Kid's Clothing"),
    ("THE ROBE GALLERY", None),
    ("TOIREM.ORG", None),
    ("TOTTINI", "Kid's Accessories"),
    ("TOYOTA", "Car Lease/Finance"),
    ("TOYS 4 U", "Toys"),
    ("TURTLE BACK ZOO", "Entertainment"),
    ("TWILIO", None),
    ("UBER EATS", "Dining & Restaurants"),
    ("ULTA", "Personal Care"),
    ("US MOBILE", "Utilities"),
    ("VEOLIA", "Utilities"),
    ("WAL-MART", "Walmart"),
    ("Walmart+ Membership", "Walmart"),
    ("WEGMANS MONTVALE", "Groceries"),
    ("WESLEY KOSHER", "Groceries"),
    ("WESTCHESTER MEDICAL CENTER", "Healthcare"),
    ("WINE ON 59", "Entertainment"),
    ("YOFFEE COFFEE", "Dining & Restaurants"),
    ("Zadarma", "Utilities"),
    ("Zara", "Women's Clothing"),
    ("Zelle payment from", None),
    ("Zelle payment to", None),
    # Added so Bank Import can auto-categorize payroll deposits the same way it does expenses —
    # not part of your list, remove the row on the Vendors tab if you don't want it.
    ("Payroll Direct DB", "Payroll Direct DB"),
    ("Payroll Cash RB", "Payroll Cash RB"),
    ("Payroll Direct RB", "Payroll Direct RB"),
]
VENDORS = [
    (i + 1, name, (_CAT_ID[cat] if cat else None))
    for i, (name, cat) in enumerate(_VENDOR_DEFS)
]

# ---------------- Dim_Account ----------------
# AccountID, AccountName, AccountType, Balance
ACCOUNTS = [
    (1, "Checking 2833", "Checking", 0),
    (2, "Savings 5950", "Savings", 0),
    (3, "Credit Card Sapphire 7750", "Credit Card", 0),
    (4, "Business Checking 4841", "Checking", 0),
    (5, "Credit Card Freedom 3361", "Credit Card", 0),
]

# ---------------- Dim_Goal ----------------
# GoalID, GoalName, TargetAmount, TargetDate, Priority
GOALS = [
    (1, "Emergency Fund", 25000, dt.date(2026, 12, 31), "High"),
]

# ---------------- Fact_Income — no sample transactions ----------------
INCOME_ROWS = []

# ---------------- Fact_Expenses — no sample transactions ----------------
EXPENSE_ROWS = []

# ---------------- Fact_SavingsTransfers — no sample transactions ----------------
SAVINGS_TRANSFER_ROWS = []
