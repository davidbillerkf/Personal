"""Data Entry sheet (simplified planner scope) — merges quick-start + README content."""
from common import *


def build_data_entry(wb):
    ws = wb.create_sheet("Data Entry")
    style_sheet_title(ws, "Data Entry", span_cols=8)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 60

    r = 4
    ws.cell(row=r, column=2, value="Personal Finance Planner — Quick Start").font = f(13, bold=True)
    r += 1
    steps = [
        "1. Got a bank/credit-card CSV export? Go to the Bank Import tab first — paste it there, pick a Vendor per"
        " row, and it auto-suggests the Category and splits Expense vs Income for you.",
        "2. Otherwise, enter transactions directly: Income for paychecks, Expenses for purchases, Savings for"
        " transfers toward a goal.",
        "3. Use the dropdowns for Vendor / Category / Account / Payment Method — never type free text — so every"
        " report stays consistent.",
        "4. Set your monthly budget per category on the Budget tab.",
        "5. Press Ctrl+Alt+F9 (Formulas > Calculate Now) so every KPI, chart, and report recalculates.",
        "6. Open the Dashboard tab to review updated KPIs, charts, and the Cut-Back Analyzer.",
    ]
    for s in steps:
        ws.cell(row=r, column=2, value=s).font = f(10)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.cell(row=r, column=2).alignment = LEFT_WRAP
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    ws.cell(row=r, column=2, value="Jump to a tab").font = f(13, bold=True)
    r += 1
    links = [
        ("Bank Import", "Paste a bank/credit-card CSV export here first"),
        ("Income", "Log paychecks"),
        ("Expenses", "Log every purchase — the most-used tab"),
        ("Savings", "Set goals and log transfers toward them"),
        ("Budget", "Set your monthly budget per category"),
        ("Categories", "Add a new expense/income category"),
        ("Vendors", "Add a new vendor"),
        ("Accounts", "Add a new bank/credit account"),
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
        "New category: add a row to Dim_Category on the Categories tab. It appears in every dropdown and report"
        " automatically — no formulas anywhere else need to change.",
        "New vendor: add a row to Dim_Vendor on the Vendors tab, with its default category.",
        "New account: add a row to Dim_Account on the Accounts tab, with its starting Balance.",
        "All dropdowns use structured references (e.g. =Dim_Category[CategoryName]), which automatically grow to"
        " include new rows — not a fixed cell range — so nothing else needs to be touched.",
    ]
    for note in notes:
        ws.cell(row=r, column=2, value=note).font = f(10)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.cell(row=r, column=2).alignment = LEFT_WRAP
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    ws.cell(row=r, column=2, value="Where your data lives").font = f(13, bold=True)
    r += 1
    storage_note = (
        "Everything is stored inside this workbook — the Excel Tables on each tab (Fact_Income, Fact_Expenses, "
        "Fact_SavingsTransfers, Dim_Category, Dim_Vendor, Dim_Account, Dim_Goal, Budget) ARE the database. There is "
        "no external server, cloud account, or database connection involved. That means: (1) this file is the only "
        "copy of your data — back it up (save a copy, or keep it in a synced folder / version control) the same way "
        "you would any important document; (2) if you keep this file in a git repo for versioning, remember that "
        "committing it puts your real transaction data into that repo's history — use a private repo if the numbers "
        "are real."
    )
    ws.cell(row=r, column=2, value=storage_note).font = f(10)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws.cell(row=r, column=2).alignment = LEFT_WRAP
    ws.row_dimensions[r].height = 75

    ws.sheet_properties.tabColor = BLUE_ACCENT
    return {}
