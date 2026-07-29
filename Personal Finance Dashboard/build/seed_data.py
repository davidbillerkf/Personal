"""Seed / sample data for the Finance Dashboard workbook.
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

# ---------------- Dim_Category ----------------
# CategoryID, CategoryName, ParentCategory, CategoryType, EssentialCategory
CATEGORIES = [
    (1, "Housing", "", "Expense", "Yes"),
    (2, "Food", "", "Expense", "Yes"),
    (3, "Transportation", "", "Expense", "Yes"),
    (4, "Utilities", "", "Expense", "Yes"),
    (5, "Insurance", "", "Expense", "Yes"),
    (6, "Healthcare", "", "Expense", "Yes"),
    (7, "Debt Payments", "", "Expense", "Yes"),
    (8, "Entertainment", "", "Expense", "No"),
    (9, "Shopping", "", "Expense", "No"),
    (10, "Subscriptions", "", "Expense", "No"),
    (11, "Personal Care", "", "Expense", "No"),
    (12, "Education", "", "Expense", "No"),
    (13, "Savings", "", "Expense", "No"),
    (14, "Investments", "", "Expense", "No"),
    (15, "Gifts & Donations", "", "Expense", "No"),
    (16, "Travel", "", "Expense", "No"),
    (17, "Salary", "", "Income", "No"),
    (18, "Bonus", "", "Income", "No"),
    (19, "Freelance", "", "Income", "No"),
    (20, "Investment Income", "", "Income", "No"),
    (21, "Other Income", "", "Income", "No"),
]

# ---------------- Dim_Subcategory ----------------
# SubcategoryID, CategoryID, SubcategoryName
SUBCATEGORIES = [
    (1, 1, "Rent / Mortgage"), (2, 1, "Home Insurance"), (3, 1, "Home Maintenance"),
    (4, 2, "Groceries"), (5, 2, "Dining Out"), (6, 2, "Coffee Shops"),
    (7, 3, "Gas"), (8, 3, "Public Transit"), (9, 3, "Car Maintenance"), (10, 3, "Rideshare"),
    (11, 4, "Electric"), (12, 4, "Water"), (13, 4, "Internet"), (14, 4, "Mobile Phone"),
    (15, 5, "Health Insurance"), (16, 5, "Auto Insurance"), (17, 5, "Life Insurance"),
    (18, 6, "Doctor Visits"), (19, 6, "Pharmacy"), (20, 6, "Dental"),
    (21, 7, "Credit Card Payment"), (22, 7, "Loan Payment"),
    (23, 8, "Streaming"), (24, 8, "Movies"), (25, 8, "Games"), (26, 8, "Concerts/Events"),
    (27, 9, "Clothing"), (28, 9, "Electronics"), (29, 9, "Amazon Purchases"), (30, 9, "Home Goods"),
    (31, 10, "Streaming Services"), (32, 10, "Software"), (33, 10, "Memberships"),
    (34, 11, "Haircuts"), (35, 11, "Gym"), (36, 11, "Cosmetics"),
    (37, 12, "Tuition"), (38, 12, "Books & Supplies"), (39, 12, "Online Courses"),
    (40, 16, "Flights"), (41, 16, "Hotels"), (42, 16, "Activities"),
]

# ---------------- Dim_Vendor ----------------
# VendorID, VendorName, VendorType, DefaultCategoryID
VENDORS = [
    (1, "Amazon", "Retail", 9), (2, "Costco", "Retail", 4), (3, "Walmart", "Retail", 4),
    (4, "Starbucks", "Food & Beverage", 6), (5, "Target", "Retail", 9),
    (6, "Whole Foods", "Grocery", 4), (7, "Shell", "Fuel", 7), (8, "Netflix", "Subscription", 23),
    (9, "Spotify", "Subscription", 23), (10, "Landlord Property Mgmt", "Housing", 1),
    (11, "City Electric Co", "Utility", 11), (12, "Comcast Xfinity", "Utility", 13),
    (13, "Verizon Wireless", "Utility", 14), (14, "Chase Bank", "Financial", 21),
    (15, "Kaiser Permanente", "Healthcare", 15), (16, "State Farm Insurance", "Insurance", 16),
    (17, "Uber", "Transportation", 10), (18, "Chipotle", "Restaurant", 5),
    (19, "AMC Theatres", "Entertainment", 24), (20, "Planet Fitness", "Personal Care", 35),
    (21, "Employer Payroll Inc", "Employer", 17), (22, "ACME Freelance Client", "Client", 19),
    (23, "Vanguard Brokerage", "Financial", 20), (24, "Delta Airlines", "Travel", 40),
    (25, "Local Coffee Roasters", "Food & Beverage", 6),
]

# ---------------- Dim_Account ----------------
# AccountID, AccountName, AccountType, Institution, Balance
ACCOUNTS = [
    (1, "Primary Checking", "Checking", "Chase Bank", 4820.55),
    (2, "High-Yield Savings", "Savings", "Ally Bank", 18250.00),
    (3, "Rewards Credit Card", "Credit Card", "Chase Bank", -1240.33),
    (4, "Brokerage Account", "Brokerage", "Vanguard", 32500.00),
    (5, "401(k)", "Retirement", "Fidelity", 68400.00),
    (6, "Roth IRA", "Retirement", "Vanguard", 21750.00),
]

# ---------------- Dim_PaymentMethod ----------------
PAYMENT_METHODS = [
    (1, "Credit Card"), (2, "Debit Card"), (3, "Cash"), (4, "ACH"),
    (5, "Check"), (6, "Bank Transfer"), (7, "Mobile Payment"),
]

# ---------------- Dim_Subscription ----------------
# SubscriptionID, ServiceName, CategoryID, MonthlyCost, AnnualCost, KeepFlag
SUBSCRIPTIONS = [
    (1, "Netflix", 10, 15.49, 185.88, "Yes"),
    (2, "Spotify Premium", 10, 11.99, 143.88, "Yes"),
    (3, "Amazon Prime", 10, 14.99, 179.88, "Yes"),
    (4, "Disney+", 10, 13.99, 167.88, "No"),
    (5, "Adobe Creative Cloud", 10, 54.99, 659.88, "Yes"),
    (6, "Planet Fitness", 11, 24.99, 299.88, "Yes"),
    (7, "NYTimes Digital", 10, 17.00, 204.00, "No"),
    (8, "iCloud Storage", 10, 2.99, 35.88, "Yes"),
    (9, "HBO Max", 10, 15.99, 191.88, "No"),
    (10, "Audible", 10, 14.95, 179.40, "Yes"),
]

# ---------------- Dim_Debt ----------------
# DebtID, DebtName, DebtType, InterestRate, OriginalBalance, CurrentBalance, MinimumPayment
DEBTS = [
    (1, "Chase Sapphire Credit Card", "Credit Card", 0.2199, 6500, 4210.50, 125),
    (2, "Toyota Camry Auto Loan", "Car Loan", 0.0549, 28000, 15230.00, 420),
    (3, "Federal Student Loan", "Student Loan", 0.0450, 32000, 21875.00, 310),
    (4, "Home Mortgage", "Mortgage", 0.0625, 380000, 351200.00, 2150),
    (5, "Personal Loan - LendingClub", "Personal Loan", 0.1099, 8000, 3120.00, 245),
]

# ---------------- Dim_Investment ----------------
# InvestmentID, InvestmentName, Ticker, AssetClass
INVESTMENTS = [
    (1, "Vanguard Total Stock Market ETF", "VTI", "US Equity"),
    (2, "Vanguard Total Intl Stock ETF", "VXUS", "Intl Equity"),
    (3, "Vanguard Total Bond Market ETF", "BND", "Bonds"),
    (4, "Apple Inc.", "AAPL", "US Equity"),
    (5, "Bitcoin", "BTC", "Crypto"),
    (6, "401(k) Target Date Fund 2055", "TDF2055", "Retirement Fund"),
    (7, "Cash Reserve", "CASH", "Cash"),
]

# ---------------- Dim_Goal ----------------
# GoalID, GoalName, TargetAmount, TargetDate, Priority
GOALS = [
    (1, "Emergency Fund", 25000, dt.date(2026, 12, 31), "High"),
    (2, "Hawaii Vacation", 6000, dt.date(2027, 6, 1), "Medium"),
    (3, "New Vehicle", 12000, dt.date(2027, 12, 31), "Medium"),
    (4, "House Down Payment", 60000, dt.date(2029, 1, 1), "High"),
    (5, "Kids Education Fund", 40000, dt.date(2035, 1, 1), "Low"),
]

# ---------------- Fact_Income (sample, 12 months) ----------------
# Date, SourceName, CategoryName, AccountName, GrossAmount, TaxesWithheld, RecurringFlag, Notes
INCOME_ROWS = []
for i in range(12, -1, -1):
    y, m = months_back(i)
    INCOME_ROWS.append((d(y, m, 1), "Employer Payroll Inc", "Salary", "Primary Checking", 5100.00, 1224.00, "Yes", "Semi-monthly payroll"))
    INCOME_ROWS.append((d(y, m, 15), "Employer Payroll Inc", "Salary", "Primary Checking", 5100.00, 1224.00, "Yes", "Semi-monthly payroll"))
INCOME_ROWS.append((d(*months_back(3), 20), "ACME Freelance Client", "Freelance", "Primary Checking", 1200.00, 0.00, "No", "Website project"))
INCOME_ROWS.append((d(*months_back(6), 5), "Vanguard Brokerage", "Investment Income", "Brokerage Account", 340.00, 0.00, "No", "Dividend payout"))
INCOME_ROWS.append((d(*months_back(1), 5), "Vanguard Brokerage", "Investment Income", "Brokerage Account", 355.00, 0.00, "No", "Dividend payout"))
INCOME_ROWS.append((d(*months_back(11), 22), "Employer Payroll Inc", "Bonus", "Primary Checking", 2500.00, 625.00, "No", "Year-end bonus"))

# ---------------- Fact_Expenses (sample) ----------------
# Date, VendorName, CategoryName, SubcategoryName, AccountName, PaymentMethod, Amount,
# TaxIncluded, EssentialFlag, NeedWantFlag, ImpulseFlag, CanCutFlag, RecurringFlag, ReimbursableFlag, Notes
EXPENSE_TEMPLATE = [
    ("Landlord Property Mgmt", "Housing", "Rent / Mortgage", "Primary Checking", "ACH", 2150.00, "No", "Yes", "Need", "No", "No", "Yes", "No", "Monthly rent"),
    ("Whole Foods", "Food", "Groceries", "Rewards Credit Card", "Credit Card", 186.42, "Yes", "Yes", "Need", "No", "No", "No", "No", ""),
    ("Costco", "Food", "Groceries", "Rewards Credit Card", "Credit Card", 245.10, "Yes", "Yes", "Need", "No", "No", "No", "No", "Bulk groceries"),
    ("Starbucks", "Food", "Coffee Shops", "Rewards Credit Card", "Credit Card", 6.75, "Yes", "No", "Want", "Yes", "Yes", "No", "No", ""),
    ("Local Coffee Roasters", "Food", "Coffee Shops", "Rewards Credit Card", "Credit Card", 5.50, "Yes", "No", "Want", "Yes", "Yes", "No", "No", ""),
    ("Chipotle", "Food", "Dining Out", "Rewards Credit Card", "Credit Card", 14.25, "Yes", "No", "Want", "No", "Yes", "No", "No", ""),
    ("Shell", "Transportation", "Gas", "Rewards Credit Card", "Credit Card", 52.30, "Yes", "Yes", "Need", "No", "No", "No", "No", ""),
    ("Uber", "Transportation", "Rideshare", "Rewards Credit Card", "Credit Card", 24.80, "No", "No", "Want", "Yes", "Yes", "No", "No", ""),
    ("City Electric Co", "Utilities", "Electric", "Primary Checking", "ACH", 128.60, "No", "Yes", "Need", "No", "No", "Yes", "No", ""),
    ("Comcast Xfinity", "Utilities", "Internet", "Primary Checking", "ACH", 79.99, "No", "Yes", "Need", "No", "No", "Yes", "No", ""),
    ("Verizon Wireless", "Utilities", "Mobile Phone", "Primary Checking", "ACH", 95.00, "No", "Yes", "Need", "No", "No", "Yes", "No", ""),
    ("State Farm Insurance", "Insurance", "Auto Insurance", "Primary Checking", "ACH", 142.00, "No", "Yes", "Need", "No", "No", "Yes", "No", ""),
    ("Kaiser Permanente", "Healthcare", "Doctor Visits", "Rewards Credit Card", "Credit Card", 45.00, "No", "Yes", "Need", "No", "No", "No", "No", "Copay"),
    ("Netflix", "Subscriptions", "Streaming Services", "Rewards Credit Card", "Credit Card", 15.49, "No", "No", "Want", "No", "Yes", "Yes", "No", ""),
    ("Spotify", "Subscriptions", "Streaming Services", "Rewards Credit Card", "Credit Card", 11.99, "No", "No", "Want", "No", "Yes", "Yes", "No", ""),
    ("Amazon", "Shopping", "Amazon Purchases", "Rewards Credit Card", "Credit Card", 68.47, "Yes", "No", "Want", "Yes", "Yes", "No", "No", ""),
    ("Amazon", "Shopping", "Electronics", "Rewards Credit Card", "Credit Card", 129.99, "Yes", "No", "Want", "Yes", "Yes", "No", "No", "Impulse buy"),
    ("Target", "Shopping", "Home Goods", "Rewards Credit Card", "Credit Card", 54.32, "Yes", "No", "Want", "No", "No", "No", "No", ""),
    ("AMC Theatres", "Entertainment", "Movies", "Rewards Credit Card", "Credit Card", 32.00, "Yes", "No", "Want", "No", "Yes", "No", "No", ""),
    ("Planet Fitness", "Personal Care", "Gym", "Primary Checking", "ACH", 24.99, "No", "No", "Want", "No", "No", "Yes", "No", ""),
]

EXPENSE_ROWS = []
import random
random.seed(42)
for i in range(11, -1, -1):
    y, m = months_back(i)
    for idx, tmpl in enumerate(EXPENSE_TEMPLATE):
        day = ((idx * 3 + 2) % 27) + 1
        amount = tmpl[5]
        jitter = round(amount * random.uniform(-0.08, 0.08), 2) if amount > 20 else 0
        row = (d(y, m, day),) + (tmpl[0], tmpl[1], tmpl[2], tmpl[3], tmpl[4], round(amount + jitter, 2)) + tmpl[6:]
        EXPENSE_ROWS.append(row)

# ---------------- Fact_Bills (sample, next ~45 days + recent history) ----------------
# BillName, DueDate, CategoryName, AccountName, Amount, Frequency, PaidStatus, AutoPayFlag, Reminder, Notes
BILLS_ROWS = [
    ("Rent", TODAY + dt.timedelta(days=2), "Housing", "Primary Checking", 2150.00, "Monthly", "Unpaid", "No", "3 days before", ""),
    ("Electric Bill", TODAY + dt.timedelta(days=5), "Utilities", "Primary Checking", 128.60, "Monthly", "Unpaid", "Yes", "3 days before", ""),
    ("Internet", TODAY + dt.timedelta(days=6), "Utilities", "Primary Checking", 79.99, "Monthly", "Unpaid", "Yes", "3 days before", ""),
    ("Mobile Phone", TODAY + dt.timedelta(days=9), "Utilities", "Primary Checking", 95.00, "Monthly", "Unpaid", "Yes", "3 days before", ""),
    ("Auto Insurance", TODAY + dt.timedelta(days=12), "Insurance", "Primary Checking", 142.00, "Monthly", "Unpaid", "Yes", "5 days before", ""),
    ("Credit Card Payment", TODAY + dt.timedelta(days=14), "Debt Payments", "Primary Checking", 250.00, "Monthly", "Unpaid", "No", "5 days before", "Min payment + extra"),
    ("Car Loan Payment", TODAY + dt.timedelta(days=18), "Debt Payments", "Primary Checking", 420.00, "Monthly", "Unpaid", "Yes", "5 days before", ""),
    ("Student Loan Payment", TODAY + dt.timedelta(days=20), "Debt Payments", "Primary Checking", 310.00, "Monthly", "Unpaid", "Yes", "5 days before", ""),
    ("Netflix", TODAY + dt.timedelta(days=1), "Subscriptions", "Rewards Credit Card", 15.49, "Monthly", "Unpaid", "Yes", "1 day before", ""),
    ("Gym Membership", TODAY - dt.timedelta(days=3), "Personal Care", "Primary Checking", 24.99, "Monthly", "Paid", "Yes", "1 day before", ""),
    ("Mortgage/Rent (Prior)", TODAY - dt.timedelta(days=28), "Housing", "Primary Checking", 2150.00, "Monthly", "Paid", "No", "3 days before", ""),
    ("Water Bill", TODAY - dt.timedelta(days=10), "Utilities", "Primary Checking", 42.15, "Monthly", "Paid", "No", "3 days before", ""),
]

# ---------------- Fact_DebtPayments (sample, 6 months) ----------------
# Date, DebtName, AccountName, PaymentAmount, PrincipalAmount
DEBT_PAYMENTS_ROWS = []
for i in range(5, -1, -1):
    y, m = months_back(i)
    DEBT_PAYMENTS_ROWS.append((d(y, m, 5), "Chase Sapphire Credit Card", "Primary Checking", 200.00, 122.00))
    DEBT_PAYMENTS_ROWS.append((d(y, m, 10), "Toyota Camry Auto Loan", "Primary Checking", 420.00, 350.00))
    DEBT_PAYMENTS_ROWS.append((d(y, m, 12), "Federal Student Loan", "Primary Checking", 310.00, 228.00))
    DEBT_PAYMENTS_ROWS.append((d(y, m, 1), "Home Mortgage", "Primary Checking", 2150.00, 470.00))
    DEBT_PAYMENTS_ROWS.append((d(y, m, 15), "Personal Loan - LendingClub", "Primary Checking", 245.00, 198.00))

# ---------------- Fact_Investments (sample) ----------------
# Date, InvestmentName, AccountName, Shares, Price, CostBasis
INVESTMENT_ROWS = [
    (d(*months_back(11), 3), "Vanguard Total Stock Market ETF", "Brokerage Account", 80, 245.10, 18200.00),
    (d(*months_back(6), 3), "Vanguard Total Intl Stock ETF", "Brokerage Account", 90, 58.20, 5000.00),
    (d(*months_back(9), 3), "Vanguard Total Bond Market ETF", "Brokerage Account", 60, 72.10, 4100.00),
    (d(*months_back(4), 3), "Apple Inc.", "Brokerage Account", 15, 189.50, 2650.00),
    (d(*months_back(2), 3), "Bitcoin", "Brokerage Account", 0.05, 62000.00, 2600.00),
    (d(*months_back(11), 1), "401(k) Target Date Fund 2055", "401(k)", 1200, 57.00, 62000.00),
    (d(*months_back(1), 1), "Cash Reserve", "Brokerage Account", 950, 1.00, 950.00),
]

# ---------------- Fact_SavingsTransfers (sample) ----------------
# Date, GoalName, AccountName, Amount
SAVINGS_TRANSFER_ROWS = []
for i in range(11, -1, -1):
    y, m = months_back(i)
    SAVINGS_TRANSFER_ROWS.append((d(y, m, 3), "Emergency Fund", "High-Yield Savings", 500.00))
    if i % 2 == 0:
        SAVINGS_TRANSFER_ROWS.append((d(y, m, 17), "Hawaii Vacation", "High-Yield Savings", 150.00))
    if i % 3 == 0:
        SAVINGS_TRANSFER_ROWS.append((d(y, m, 20), "New Vehicle", "High-Yield Savings", 200.00))
SAVINGS_TRANSFER_ROWS.append((d(*months_back(8), 5), "House Down Payment", "High-Yield Savings", 1000.00))
SAVINGS_TRANSFER_ROWS.append((d(*months_back(2), 5), "House Down Payment", "High-Yield Savings", 1000.00))

# ---------------- Net worth history (sample, 13 months incl. current) ----------------
# Month label handled in generator; values: Assets total override optional
