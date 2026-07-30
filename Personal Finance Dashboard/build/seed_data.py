"""Seed / sample data for the simplified Personal Finance Planner workbook.
All rows here are illustrative demo data the user is expected to replace or delete
(see README). They exist so charts, KPIs and reports render real output out of the box.
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
    (3, "Car", "Expense", "Yes"),
    (4, "Groceries", "Expense", "Yes"),
    (5, "Insurance", "Expense", "Yes"),
    (6, "Healthcare", "Expense", "Yes"),
    (7, "Dining & Entertainment", "Expense", "No"),
    (8, "Personal Care", "Expense", "No"),
    (9, "Shopping", "Expense", "No"),
    (10, "Miscellaneous", "Expense", "No"),
    (11, "Payroll", "Income", "No"),
    (12, "Other Income", "Income", "No"),
]

# ---------------- Dim_Vendor ----------------
# VendorID, VendorName, DefaultCategoryID
VENDORS = [
    (1, "Mortgage Lender", 1),
    (2, "City Electric Co", 2),
    (3, "Comcast Xfinity", 2),
    (4, "Verizon Wireless", 2),
    (5, "Shell Gas Station", 3),
    (6, "AutoZone", 3),
    (7, "Whole Foods", 4),
    (8, "Costco", 4),
    (9, "State Farm Insurance", 5),
    (10, "Kaiser Permanente", 6),
    (11, "Chipotle", 7),
    (12, "Netflix", 7),
    (13, "AMC Theatres", 7),
    (14, "Planet Fitness", 8),
    (15, "Amazon", 9),
    (16, "Target", 9),
    (17, "Employer Payroll Inc", 11),
    (18, "Other Income Source", 12),
]

# ---------------- Dim_Account ----------------
# AccountID, AccountName, AccountType, Balance
ACCOUNTS = [
    (1, "Checking", "Checking", 4820.55),
    (2, "Savings", "Savings", 18250.00),
    (3, "Credit Card", "Credit Card", -1240.33),
]

# ---------------- Dim_Goal ----------------
# GoalID, GoalName, TargetAmount, TargetDate, Priority
GOALS = [
    (1, "Emergency Fund", 25000, dt.date(2026, 12, 31), "High"),
    (2, "Vacation", 6000, dt.date(2027, 6, 1), "Medium"),
    (3, "New Car", 12000, dt.date(2027, 12, 31), "Medium"),
]

# ---------------- Fact_Income (sample, 13 months, Payroll only) ----------------
# Date, VendorName, CategoryName, AccountName, GrossAmount, TaxesWithheld, RecurringFlag, Notes
INCOME_ROWS = []
for i in range(12, -1, -1):
    y, m = months_back(i)
    INCOME_ROWS.append((d(y, m, 1), "Employer Payroll Inc", "Payroll", "Checking", 5100.00, 1224.00, "Yes", "Semi-monthly payroll"))
    INCOME_ROWS.append((d(y, m, 15), "Employer Payroll Inc", "Payroll", "Checking", 5100.00, 1224.00, "Yes", "Semi-monthly payroll"))

# ---------------- Fact_Expenses (sample) ----------------
# Vendor, Category, Account, PaymentMethod, Amount, NeedWantFlag, RecurringFlag, Notes
EXPENSE_TEMPLATE = [
    ("Mortgage Lender", "Mortgage", "Checking", "ACH", 2150.00, "Need", "Yes", "Monthly mortgage payment"),
    ("Whole Foods", "Groceries", "Credit Card", "Credit Card", 186.42, "Need", "No", ""),
    ("Costco", "Groceries", "Credit Card", "Credit Card", 245.10, "Need", "No", "Bulk groceries"),
    ("Shell Gas Station", "Car", "Credit Card", "Credit Card", 52.30, "Need", "No", "Gas"),
    ("AutoZone", "Car", "Credit Card", "Credit Card", 38.75, "Need", "No", "Wiper blades / oil"),
    ("City Electric Co", "Utilities", "Checking", "ACH", 128.60, "Need", "Yes", ""),
    ("Comcast Xfinity", "Utilities", "Checking", "ACH", 79.99, "Need", "Yes", "Internet"),
    ("Verizon Wireless", "Utilities", "Checking", "ACH", 95.00, "Need", "Yes", "Mobile phone"),
    ("State Farm Insurance", "Insurance", "Checking", "ACH", 142.00, "Need", "Yes", "Auto + home insurance"),
    ("Kaiser Permanente", "Healthcare", "Credit Card", "Credit Card", 45.00, "Need", "No", "Copay"),
    ("Chipotle", "Dining & Entertainment", "Credit Card", "Credit Card", 14.25, "Want", "No", ""),
    ("Netflix", "Dining & Entertainment", "Credit Card", "Credit Card", 15.49, "Want", "Yes", "Streaming subscription"),
    ("AMC Theatres", "Dining & Entertainment", "Credit Card", "Credit Card", 32.00, "Want", "No", ""),
    ("Planet Fitness", "Personal Care", "Checking", "ACH", 24.99, "Want", "Yes", "Gym membership"),
    ("Amazon", "Shopping", "Credit Card", "Credit Card", 68.47, "Want", "No", ""),
    ("Target", "Shopping", "Credit Card", "Credit Card", 54.32, "Want", "No", ""),
]

EXPENSE_ROWS = []
import random
random.seed(42)
for i in range(11, -1, -1):
    y, m = months_back(i)
    for idx, tmpl in enumerate(EXPENSE_TEMPLATE):
        day = ((idx * 3 + 2) % 27) + 1
        amount = tmpl[4]
        jitter = round(amount * random.uniform(-0.08, 0.08), 2) if amount > 20 else 0
        row = (d(y, m, day), tmpl[0], tmpl[1], tmpl[2], tmpl[3], round(amount + jitter, 2)) + tmpl[5:]
        EXPENSE_ROWS.append(row)

# ---------------- Fact_SavingsTransfers (sample) ----------------
# Date, GoalName, AccountName, Amount
SAVINGS_TRANSFER_ROWS = []
for i in range(11, -1, -1):
    y, m = months_back(i)
    SAVINGS_TRANSFER_ROWS.append((d(y, m, 3), "Emergency Fund", "Savings", 500.00))
    if i % 2 == 0:
        SAVINGS_TRANSFER_ROWS.append((d(y, m, 17), "Vacation", "Savings", 150.00))
    if i % 3 == 0:
        SAVINGS_TRANSFER_ROWS.append((d(y, m, 20), "New Car", "Savings", 200.00))
