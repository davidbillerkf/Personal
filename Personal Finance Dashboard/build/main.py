import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from openpyxl import Workbook
from common import *
import dims
import facts1
import facts2
import reports
import dashboard
import misc
import bank_import
import trends

OUT = "/home/user/Personal/Personal Finance Dashboard/Personal Finance Intelligence Dashboard.xlsx"

SHEET_ORDER = [
    "Dashboard", "Data Entry", "Bank Import", "Income", "Expenses", "Savings", "Budget", "Trends",
    "Categories", "Vendors", "Accounts", "Settings", "Dim_Date", "ChartData",
]


def main():
    wb = Workbook()
    wb.remove(wb.active)

    dims.build_settings(wb)
    dims.build_dim_date(wb)
    dims.build_categories(wb)
    dims.build_vendors(wb)
    dims.build_accounts(wb)

    facts1.build_income(wb)
    facts1.build_expenses(wb)

    facts2.build_savings(wb)

    bud_ctx = reports.build_budget(wb)

    bank_import.build_bank_import(wb)

    cd_ctx = dashboard.build_chartdata(wb)
    dash_ctx = {
        "chartdata": cd_ctx,
        "bud_info": bud_ctx["bud_info"],
    }
    dashboard.build_dashboard(wb, dash_ctx)

    trends.build_trends(wb)

    misc.build_data_entry(wb)

    # Reorder sheets into the spec-defined order
    by_title = {ws.title: ws for ws in wb.worksheets}
    wb._sheets = [by_title[name] for name in SHEET_ORDER if name in by_title]
    wb.active = 0

    wb.save(OUT)
    print("Saved:", OUT)


if __name__ == "__main__":
    main()
