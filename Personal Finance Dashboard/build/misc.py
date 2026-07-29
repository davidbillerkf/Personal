"""Data Entry, Data Model, README sheets."""
from common import *


def build_data_entry(wb):
    ws = wb.create_sheet("Data Entry")
    style_sheet_title(ws, "Data Entry", span_cols=8)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 60

    r = 4
    ws.cell(row=r, column=2, value="Your daily workflow").font = f(13, bold=True)
    r += 1
    steps = [
        "1. Enter each transaction on the Income, Expenses, Bills, Debt, Savings Goals, or Investments tab as it happens.",
        "2. Use the dropdowns for Vendor / Category / Subcategory / Account / Payment Method — never type free text — so every report stays consistent.",
        "3. If this workbook is connected to bank exports via Power Query (see the PowerQuery/ folder), click Data > Queries & Connections > Refresh All.",
        "4. Press Ctrl+Alt+F9 (Excel: Formulas > Calculate Now) so every KPI, chart and report recalculates from your new data.",
        "5. Open the Dashboard tab to review updated KPIs, charts, the Cut-Back Analyzer and your Financial Health Score.",
    ]
    for s in steps:
        ws.cell(row=r, column=2, value=s).font = f(10)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.cell(row=r, column=2).alignment = LEFT_WRAP
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    ws.cell(row=r, column=2, value="Jump to a data-entry tab").font = f(13, bold=True)
    r += 1
    links = [
        ("Income", "Log paychecks, freelance income, dividends"),
        ("Expenses", "Log every purchase — the most-used tab"),
        ("Bills", "Track recurring & one-time bills, due dates, autopay"),
        ("Debt", "Log loan/credit-card payments; manage payoff strategy"),
        ("Savings Goals", "Log transfers toward Emergency Fund, Vacation, etc."),
        ("Investments", "Log brokerage/401(k)/crypto lots and prices"),
        ("Budget", "Set your monthly budget per category"),
    ]
    for name, desc in links:
        c = ws.cell(row=r, column=2, value=name)
        c.font = f(10, bold=True, color=BLUE_ACCENT)
        c.hyperlink = f"#'{name}'!A1"
        ws.cell(row=r, column=3, value=desc).font = f(9, italic=True, color="6B7280")
        r += 1

    r += 2
    ws.cell(row=r, column=2, value="Adding new categories, vendors, or accounts").font = f(13, bold=True)
    r += 1
    notes = [
        "New category: add a row to Dim_Category on the Categories tab. It appears in every dropdown and report automatically.",
        "New subcategory: add a row to Dim_Subcategory on the Categories tab, with the CategoryID it belongs under.",
        "New vendor: add a row to Dim_Vendor on the Vendors tab.",
        "New account: add a row to Dim_Account on the Accounts tab, with its starting Balance.",
        "No formulas anywhere else need to change — every report reads from these dimension tables live.",
    ]
    for note in notes:
        ws.cell(row=r, column=2, value=note).font = f(10)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.cell(row=r, column=2).alignment = LEFT_WRAP
        ws.row_dimensions[r].height = 24
        r += 1

    ws.sheet_properties.tabColor = BLUE_ACCENT
    return {}


def build_data_model(wb):
    ws = wb.create_sheet("Data Model")
    style_sheet_title(ws, "Data Model", span_cols=8)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 60

    r = 4
    ws.cell(row=r, column=1, value="Star Schema Architecture").font = f(13, bold=True)
    r += 1
    intro = (
        "This workbook is modeled as a star schema: FACT tables hold transactions "
        "(Expenses, Income, Bills, Debt Payments, Investments, Savings Transfers); DIMENSION "
        "tables hold controlled vocabularies (Category, Subcategory, Vendor, Account, Payment "
        "Method, Debt, Investment, Goal, Subscription). Every fact row stores its dimension "
        "value as text chosen from a dropdown (e.g. VendorName), plus a hidden ID column that "
        "looks the surrogate key up via INDEX/MATCH against the dimension table — this "
        "preserves referential integrity and a true ID-based star schema while keeping SUMIFS-based "
        "reports simple and fast at 100,000+ rows."
    )
    ws.cell(row=r, column=1, value=intro).font = f(10)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(row=r, column=1).alignment = LEFT_WRAP
    ws.row_dimensions[r].height = 75
    r += 2

    note2 = (
        "Dates are handled the same way a real Dim_Date would, without needing surrogate-key "
        "joins: DateID = INT(Date), i.e. Excel's native date serial number, so a hidden "
        "Dim_Date sheet (2025-2027, one row per calendar day with Day/Week/Month/Quarter/Year/"
        "FiscalYear) already lines up with every fact table by construction. Reports filter "
        "dates directly with SUMIFS date-bound criteria rather than walking through a join, "
        "which is both simpler and faster than formula-based lookups at scale."
    )
    ws.cell(row=r, column=1, value=note2).font = f(10)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(row=r, column=1).alignment = LEFT_WRAP
    ws.row_dimensions[r].height = 60
    r += 2

    note3 = (
        "On engineering choices: this workbook intentionally avoids native Excel PivotTables "
        "and a Power Pivot/xVelocity Data Model — those are authored inside the Excel "
        "application itself and cannot be generated by an external build script. In their "
        "place, every 'pivot-style' report (Monthly Reports, Annual Reports, Budget, "
        "Dashboard) is built from SUMIFS/COUNTIFS/INDEX/MATCH formulas against the star "
        "schema tables below, which recalculates live and needs no manual refresh step. "
        "If you want true Power Pivot: Data > Get Data > From Table/Range on each table "
        "below, check 'Add this data to the Data Model', then build the relationships shown "
        "here in Power Pivot > Manage > Diagram View."
    )
    ws.cell(row=r, column=1, value=note3).font = f(10)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(row=r, column=1).alignment = LEFT_WRAP
    ws.row_dimensions[r].height = 90
    r += 2

    r = style_section_header(ws, r, 1, "Tables & Relationships", span=8)
    headers = ["Table", "Type", "Sheet", "Key Fields", "Relates To"]
    rows = [
        ("Fact_Expenses", "Fact", "Expenses", "ExpenseID, Date, Amount", "Dim_Vendor, Dim_Category, Dim_Subcategory, Dim_Account, Dim_PaymentMethod, Dim_Date"),
        ("Fact_Income", "Fact", "Income", "IncomeID, Date, NetAmount", "Dim_Source, Dim_Category, Dim_Account, Dim_Date"),
        ("Fact_Bills", "Fact", "Bills", "BillID, DueDate, Amount", "Dim_Category, Dim_Account, Dim_Date"),
        ("Fact_DebtPayments", "Fact", "Debt", "PaymentID, Date, PaymentAmount", "Dim_Debt, Dim_Account, Dim_Date"),
        ("Fact_Investments", "Fact", "Investments", "TransactionID, Date, MarketValue", "Dim_Investment, Dim_Account, Dim_Date"),
        ("Fact_SavingsTransfers", "Fact", "Savings Goals", "TransferID, Date, Amount", "Dim_Goal, Dim_Account, Dim_Date"),
        ("Dim_Category", "Dimension", "Categories", "CategoryID, CategoryName", "—"),
        ("Dim_Subcategory", "Dimension", "Categories", "SubcategoryID, CategoryID", "Dim_Category"),
        ("Dim_Vendor", "Dimension", "Vendors", "VendorID, VendorName", "Dim_Category (DefaultCategoryID)"),
        ("Dim_Account", "Dimension", "Accounts", "AccountID, AccountName, Balance", "—"),
        ("Dim_PaymentMethod", "Dimension", "Settings", "PaymentMethodID, MethodName", "—"),
        ("Dim_Subscription", "Dimension", "Subscriptions", "SubscriptionID, ServiceName", "Dim_Category"),
        ("Dim_Debt", "Dimension", "Debt", "DebtID, DebtName, CurrentBalance", "—"),
        ("Dim_Investment", "Dimension", "Investments", "InvestmentID, Ticker, AssetClass", "—"),
        ("Dim_Goal", "Dimension", "Savings Goals", "GoalID, GoalName, TargetAmount", "—"),
        ("Dim_Date", "Dimension (hidden)", "Dim_Date", "DateID = INT(Date)", "—"),
    ]
    for j, h in enumerate(headers):
        c = ws.cell(row=r, column=1 + j, value=h)
        c.font = f(10, bold=True, color=WHITE); c.fill = fill(NAVY)
    for i, row_ in enumerate(rows):
        rr = r + 1 + i
        for j, val in enumerate(row_):
            cell = ws.cell(row=rr, column=1 + j, value=val)
            cell.font = f(9)
            cell.alignment = LEFT_WRAP
    ws.column_dimensions["D"].width = 26
    ws.column_dimensions["E"].width = 46

    ws.sheet_properties.tabColor = "6B7280"
    return {}


def build_readme_sheet(wb):
    ws = wb.create_sheet("README")
    style_sheet_title(ws, "README", span_cols=8)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 100

    r = 4
    lines = [
        ("Personal Finance Intelligence Dashboard", 16, True),
        ("", 10, False),
        ("A star-schema-based personal finance system built entirely with native Excel "
         "formulas, tables, and charts — no VBA, no macros required.", 10, False),
        ("", 10, False),
        ("QUICK START", 12, True),
        ("1. Start on the Data Entry tab for a guided workflow.", 10, False),
        ("2. Log transactions on Income / Expenses / Bills / Debt / Savings Goals / Investments.", 10, False),
        ("3. Set category budgets on the Budget tab.", 10, False),
        ("4. Review the Dashboard for KPIs, charts, and recommendations.", 10, False),
        ("", 10, False),
        ("KEY CONCEPTS", 12, True),
        ("- Dimension tables (Categories, Vendors, Accounts, Settings) define controlled "
         "vocabularies used everywhere else via dropdowns.", 10, False),
        ("- Fact tables (Income, Expenses, Bills, Debt, Investments, Savings Goals) hold your "
         "transactions and reference dimension values by name, with a hidden ID lookup column "
         "for true star-schema integrity.", 10, False),
        ("- Every dashboard, report, and chart is formula-driven (SUMIFS/COUNTIFS/INDEX/MATCH) "
         "so it updates the instant you enter data and recalculate — no manual refresh of "
         "reports needed.", 10, False),
        ("", 10, False),
        ("KNOWN ENVIRONMENT LIMITATION (please read)", 12, True),
        ("This workbook was authored programmatically outside of Excel. Formulas were written "
         "to be fully compatible with Excel 365 and use only widely-supported functions "
         "(SUMIFS, COUNTIFS, INDEX, MATCH, IFERROR, SUMPRODUCT, RANK, LARGE/SMALL, EOMONTH, "
         "EDATE, TEXTJOIN, IFS — deliberately avoiding volatile functions like NOW/RAND/OFFSET/"
         "INDIRECT and avoiding dynamic-array-only functions like XLOOKUP/FILTER/UNIQUE/SORT "
         "for maximum compatibility). Because of a LibreOffice/document-type-detection defect "
         "in the build sandbox, automated formula recalculation could not be run in this "
         "environment before delivery. The very first time you open this file in Excel, press "
         "Ctrl+Alt+F9 (Formulas > Calculate Now) to force a full recalculation and populate "
         "every cached value. Please review the Dashboard and Financial Health tabs after that "
         "first recalculation and report anything that looks off.", 10, False),
        ("", 10, False),
        ("See the full README.md and CHANGELOG.md shipped alongside this file for more detail.", 10, False),
    ]
    for text, size, bold in lines:
        c = ws.cell(row=r, column=1, value=text)
        c.font = f(size, bold=bold)
        c.alignment = LEFT_WRAP
        ws.row_dimensions[r].height = max(18, size + 6) if not text or len(text) < 90 else 30
        r += 1

    ws.sheet_properties.tabColor = NAVY
    return {}
