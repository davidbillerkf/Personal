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

OUT = "/home/user/Personal/Personal Finance Dashboard/Personal Finance Intelligence Dashboard.xlsx"

SHEET_ORDER = [
    "Dashboard", "Data Entry", "Income", "Expenses", "Categories", "Vendors", "Accounts",
    "Budget", "Bills", "Subscriptions", "Debt", "Savings Goals", "Investments", "Net Worth",
    "Monthly Reports", "Annual Reports", "Cash Flow Forecast", "Financial Health",
    "Data Model", "Settings", "README", "Dim_Date", "ChartData",
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
    bill_ctx = facts1.build_bills(wb)
    facts1.build_subscriptions(wb)

    facts2.build_debt(wb)
    facts2.build_savings_goals(wb)
    inv_ctx = facts2.build_investments(wb)
    nw_ctx = facts2.build_net_worth(wb)

    bud_ctx = reports.build_budget(wb)
    mr_ctx = reports.build_monthly_reports(wb)
    reports.build_annual_reports(wb)
    reports.build_cash_flow_forecast(wb)
    reports.build_financial_health(wb)

    cd_ctx = dashboard.build_chartdata(wb)
    dash_ctx = {
        "chartdata": cd_ctx,
        "bud_info": bud_ctx["bud_info"],
        "nw_info": nw_ctx,
        "inv_alloc": {"start": inv_ctx["alloc_start"], "end": inv_ctx["alloc_end"]},
        "top10_start": mr_ctx["top10_start"],
    }
    dashboard.build_dashboard(wb, dash_ctx)

    misc.build_data_entry(wb)
    misc.build_data_model(wb)
    misc.build_readme_sheet(wb)

    # Reorder sheets into the spec-defined order
    by_title = {ws.title: ws for ws in wb.worksheets}
    wb._sheets = [by_title[name] for name in SHEET_ORDER if name in by_title]
    wb.active = 0

    wb.save(OUT)
    print("Saved:", OUT)


if __name__ == "__main__":
    main()
