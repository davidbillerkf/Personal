"""Budget, Monthly Reports, Annual Reports, Cash Flow Forecast, Financial Health sheets."""
from openpyxl.workbook.defined_name import DefinedName
from common import *
import seed_data as sd

EXPENSE_CATEGORIES = [c[1] for c in sd.CATEGORIES if c[3] == "Expense"]
BUDGET_AMOUNTS = {
    "Housing": 2200, "Food": 650, "Transportation": 250, "Utilities": 320, "Insurance": 300,
    "Healthcare": 150, "Debt Payments": 1000, "Entertainment": 120, "Shopping": 250,
    "Subscriptions": 120, "Personal Care": 80, "Education": 100, "Savings": 500,
    "Investments": 400, "Gifts & Donations": 60, "Travel": 150,
}


def build_budget(wb):
    ws = wb.create_sheet("Budget")
    style_sheet_title(ws, "Budget", span_cols=9)
    ws.sheet_view.showGridLines = False
    for col in range(1, 10):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions["A"].width = 20

    kpi_specs = [
        ("Total Monthly Budget", "=SUM(Budget[MonthlyBudget])", CUR_FMT0),
        ("Total Actual (This Month)", "=SUM(Budget[Actual])", CUR_FMT0),
        ("Remaining", "=SUM(Budget[MonthlyBudget])-SUM(Budget[Actual])", CUR_FMT0),
        ("Overall Utilization", "=SUM(Budget[Actual])/SUM(Budget[MonthlyBudget])", PCT_FMT),
    ]
    for i, (label, formula, fmt) in enumerate(kpi_specs):
        kpi_card(ws, 4, 1 + i * 2, label, formula, number_format=fmt, width_cols=2)

    r = 9
    r = style_section_header(ws, r, 1,
        "Category Budgets  —  set MonthlyBudget per category; Actual pulls live from Expenses this month", span=8)
    headers = ["CategoryName", "MonthlyBudget", "Actual", "Remaining", "Variance %", "Status", "Progress"]
    data_rows = [[cat, BUDGET_AMOUNTS.get(cat, 100)] + [None] * 5 for cat in EXPENSE_CATEGORIES]
    bud_info = add_table(ws, r, 1, headers, "Budget", data_rows,
                          col_widths=[20, 14, 13, 13, 12, 14, 12])
    fr, lr = bud_info["first_data_row"], bud_info["end_row"]
    for rr in range(fr, lr + 1):
        ws.cell(row=rr, column=2).number_format = CUR_FMT
        ws.cell(row=rr, column=2).font = f(10, color=BLUE_ACCENT)
        ws.cell(row=rr, column=3,
                value=f'=SUMIFS(Fact_Expenses[Amount],Fact_Expenses[CategoryName],$A{rr},Fact_Expenses[Date],">="&CurMonthStart,Fact_Expenses[Date],"<="&CurMonthEnd)')
        ws.cell(row=rr, column=3).number_format = CUR_FMT
        ws.cell(row=rr, column=4, value=f"=$B{rr}-$C{rr}")
        ws.cell(row=rr, column=4).number_format = CUR_FMT
        ws.cell(row=rr, column=5, value=f'=IFERROR($C{rr}/$B{rr}-1,0)')
        ws.cell(row=rr, column=5).number_format = PCT_FMT
        ws.cell(row=rr, column=6,
                value=(f'=_xlfn.IFS($C{rr}/$B{rr}>1,"Over Budget",$C{rr}/$B{rr}>=0.8,"Near Limit",TRUE,"Under Budget")'))
        ws.cell(row=rr, column=7, value=f"=MIN($C{rr}/$B{rr},1.5)")
        ws.cell(row=rr, column=7).number_format = PCT_FMT
    ws.conditional_formatting.add(f"F{fr}:F{lr}", CellIsRule(operator="equal", formula=['"Over Budget"'], fill=fill(RED_LIGHT), font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(f"F{fr}:F{lr}", CellIsRule(operator="equal", formula=['"Near Limit"'], fill=fill(ORANGE_LIGHT), font=Font(color=ORANGE, bold=True)))
    ws.conditional_formatting.add(f"F{fr}:F{lr}", CellIsRule(operator="equal", formula=['"Under Budget"'], fill=fill(GREEN_LIGHT), font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"G{fr}:G{lr}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.5, color=BLUE_ACCENT))

    ws.sheet_properties.tabColor = BLUE_ACCENT
    return {"bud_info": bud_info}


def build_monthly_reports(wb):
    ws = wb.create_sheet("Monthly Reports")
    style_sheet_title(ws, "Monthly Reports", span_cols=14)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 20
    for col in range(2, 15):
        ws.column_dimensions[get_column_letter(col)].width = 11

    r = 4
    r = style_section_header(ws, r, 1, "Income vs Expenses — Trailing 13 Months", span=13)
    ws.cell(row=r, column=1, value="Month").font = f(10, bold=True)
    month_hdr_row = r
    n = 13
    for i in range(n):
        months_ago = n - 1 - i
        c = ws.cell(row=r, column=2 + i, value=f"=EDATE(TodayDate,{-months_ago})")
        c.number_format = "mmm-yy"
        c.font = f(9, bold=True)
    metrics = ["Income", "Expenses", "Net Savings", "Savings Rate %"]
    metric_rows = {}
    for mi, metric in enumerate(metrics):
        rr = r + 1 + mi
        ws.cell(row=rr, column=1, value=metric).font = f(10, bold=(metric != "Savings Rate %"))
        metric_rows[metric] = rr
    for i in range(n):
        col = 2 + i
        cl = get_column_letter(col)
        start_expr = f"EOMONTH({cl}{month_hdr_row},-1)+1"
        end_expr = f"EOMONTH({cl}{month_hdr_row},0)"
        ic = ws.cell(row=metric_rows["Income"], column=col,
                      value=f'=SUMIFS(Fact_Income[NetAmount],Fact_Income[Date],">="&{start_expr},Fact_Income[Date],"<="&{end_expr})')
        ic.number_format = CUR_FMT0
        ec = ws.cell(row=metric_rows["Expenses"], column=col,
                      value=f'=SUMIFS(Fact_Expenses[Amount],Fact_Expenses[Date],">="&{start_expr},Fact_Expenses[Date],"<="&{end_expr})')
        ec.number_format = CUR_FMT0
        nc = ws.cell(row=metric_rows["Net Savings"], column=col, value=f"={cl}{metric_rows['Income']}-{cl}{metric_rows['Expenses']}")
        nc.number_format = CUR_FMT0
        sc = ws.cell(row=metric_rows["Savings Rate %"], column=col,
                      value=f"=IFERROR({cl}{metric_rows['Net Savings']}/{cl}{metric_rows['Income']},0)")
        sc.number_format = PCT_FMT
    ws.conditional_formatting.add(
        f"B{metric_rows['Net Savings']}:N{metric_rows['Net Savings']}",
        CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))

    r2 = metric_rows["Savings Rate %"] + 3
    r2 = style_section_header(ws, r2, 1, "Spending by Category — Trailing 13 Months", span=13)
    cat_hdr_row = r2
    for i in range(n):
        months_ago = n - 1 - i
        c = ws.cell(row=r2, column=2 + i, value=f"=EDATE(TodayDate,{-months_ago})")
        c.number_format = "mmm-yy"
        c.font = f(9, bold=True)
    cat_rows = {}
    for ci, cat in enumerate(EXPENSE_CATEGORIES):
        rr = r2 + 1 + ci
        ws.cell(row=rr, column=1, value=cat).font = f(9)
        cat_rows[cat] = rr
        for i in range(n):
            col = 2 + i
            cl = get_column_letter(col)
            start_expr = f"EOMONTH({cl}{cat_hdr_row},-1)+1"
            end_expr = f"EOMONTH({cl}{cat_hdr_row},0)"
            cc = ws.cell(row=rr, column=col,
                          value=f'=SUMIFS(Fact_Expenses[Amount],Fact_Expenses[CategoryName],$A{rr},Fact_Expenses[Date],">="&{start_expr},Fact_Expenses[Date],"<="&{end_expr})')
            cc.number_format = CUR_FMT0
    total_row = r2 + 1 + len(EXPENSE_CATEGORIES)
    ws.cell(row=total_row, column=1, value="TOTAL").font = f(9, bold=True)
    for i in range(n):
        col = 2 + i
        cl = get_column_letter(col)
        ws.cell(row=total_row, column=col, value=f"=SUM({cl}{r2+1}:{cl}{r2+len(EXPENSE_CATEGORIES)})").font = f(9, bold=True)
        ws.cell(row=total_row, column=col).number_format = CUR_FMT0
        ws.cell(row=total_row, column=col).border = Border(top=Side(style="thin"))

    r3 = total_row + 3
    r3 = style_section_header(ws, r3, 1, "This Month — Top 10 Expenses", span=4)
    ws.cell(row=r3, column=1, value="Rank").font = f(10, bold=True)
    ws.cell(row=r3, column=2, value="Vendor").font = f(10, bold=True)
    ws.cell(row=r3, column=3, value="Category").font = f(10, bold=True)
    ws.cell(row=r3, column=4, value="Amount").font = f(10, bold=True)
    # Uses Fact_Expenses[CurMonthAmt], a materialized helper column (Amount if in the
    # current month else 0) so LARGE/MATCH work as ordinary (non-array) formulas.
    for i in range(10):
        rr = r3 + 1 + i
        ws.cell(row=rr, column=1, value=i + 1)
        ws.cell(row=rr, column=4, value=f'=IFERROR(LARGE(Fact_Expenses[CurMonthAmt],{i+1}),0)')
        ws.cell(row=rr, column=4).number_format = CUR_FMT
        ws.cell(row=rr, column=2,
                value=f'=IFERROR(INDEX(Fact_Expenses[VendorName],MATCH($D{rr},Fact_Expenses[CurMonthAmt],0)),"")')
        ws.cell(row=rr, column=3,
                value=f'=IFERROR(INDEX(Fact_Expenses[CategoryName],MATCH($D{rr},Fact_Expenses[CurMonthAmt],0)),"")')
        for c in (2, 3, 4):
            ws.cell(row=rr, column=c).font = f(9)

    ws.sheet_properties.tabColor = "6B7280"
    return {"metric_rows": metric_rows, "month_hdr_row": month_hdr_row, "cat_hdr_row": cat_hdr_row,
            "cat_rows": cat_rows, "n_months": n, "top10_start": r3 + 1}


def build_annual_reports(wb):
    ws = wb.create_sheet("Annual Reports")
    style_sheet_title(ws, "Annual Reports", span_cols=6)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 24
    for col in range(2, 6):
        ws.column_dimensions[get_column_letter(col)].width
        ws.column_dimensions[get_column_letter(col)].width = 16

    r = 4
    r = style_section_header(ws, r, 1, "Year-Over-Year Summary", span=4)
    ws.cell(row=r, column=1, value="Metric").font = f(10, bold=True)
    ws.cell(row=r, column=2, value="Last Year").font = f(10, bold=True)
    ws.cell(row=r, column=3, value="This Year").font = f(10, bold=True)
    ws.cell(row=r, column=4, value="YoY Change %").font = f(10, bold=True)
    rows = [
        ("Total Income",
         '=SUMIFS(Fact_Income[NetAmount],Fact_Income[Date],">="&LastYearStart,Fact_Income[Date],"<="&LastYearEnd)',
         '=SUMIFS(Fact_Income[NetAmount],Fact_Income[Date],">="&CurYearStart,Fact_Income[Date],"<="&CurYearEnd)'),
        ("Total Expenses",
         '=SUMIFS(Fact_Expenses[Amount],Fact_Expenses[Date],">="&LastYearStart,Fact_Expenses[Date],"<="&LastYearEnd)',
         '=SUMIFS(Fact_Expenses[Amount],Fact_Expenses[Date],">="&CurYearStart,Fact_Expenses[Date],"<="&CurYearEnd)'),
    ]
    for i, (label, lyf, cyf) in enumerate(rows):
        rr = r + 1 + i
        ws.cell(row=rr, column=1, value=label).font = f(10)
        ws.cell(row=rr, column=2, value=lyf).number_format = CUR_FMT0
        ws.cell(row=rr, column=3, value=cyf).number_format = CUR_FMT0
        ws.cell(row=rr, column=4, value=f"=IFERROR(C{rr}/B{rr}-1,0)").number_format = PCT_FMT
    savings_row = r + 3
    ws.cell(row=savings_row, column=1, value="Net Savings").font = f(10, bold=True)
    ws.cell(row=savings_row, column=2, value=f"=B{r+1}-B{r+2}").number_format = CUR_FMT0
    ws.cell(row=savings_row, column=3, value=f"=C{r+1}-C{r+2}").number_format = CUR_FMT0
    ws.cell(row=savings_row, column=4, value=f"=IFERROR(C{savings_row}/B{savings_row}-1,0)").number_format = PCT_FMT
    sr_row = savings_row + 1
    ws.cell(row=sr_row, column=1, value="Savings Rate").font = f(10)
    ws.cell(row=sr_row, column=2, value=f"=IFERROR(B{savings_row}/B{r+1},0)").number_format = PCT_FMT
    ws.cell(row=sr_row, column=3, value=f"=IFERROR(C{savings_row}/C{r+1},0)").number_format = PCT_FMT
    nw_row = sr_row + 2
    ws.cell(row=nw_row, column=1, value="Net Worth (current)").font = f(10)
    ws.cell(row=nw_row, column=3, value="=NetWorthCell").number_format = CUR_FMT0
    ws.conditional_formatting.add(f"D{r+1}:D{savings_row}", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED)))
    ws.conditional_formatting.add(f"D{r+1}:D{savings_row}", CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=Font(color=GREEN)))

    ws.sheet_properties.tabColor = "6B7280"
    return {}


def build_cash_flow_forecast(wb):
    ws = wb.create_sheet("Cash Flow Forecast")
    style_sheet_title(ws, "Cash Flow Forecast", span_cols=9)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 16
    for col in range(2, 9):
        ws.column_dimensions[get_column_letter(col)].width = 14

    r = 4
    r = style_section_header(ws, r, 1, "Forecast Assumptions (based on trailing 3 complete months)", span=6)
    labels = [
        ("Avg Monthly Income", '=SUMIFS(Fact_Income[NetAmount],Fact_Income[Date],">="&EOMONTH(TodayDate,-4)+1,Fact_Income[Date],"<="&EOMONTH(TodayDate,-1))/3'),
        ("Avg Monthly Expenses", '=SUMIFS(Fact_Expenses[Amount],Fact_Expenses[Date],">="&EOMONTH(TodayDate,-4)+1,Fact_Expenses[Date],"<="&EOMONTH(TodayDate,-1))/3'),
        ("Monthly Debt Payments (min.)", "=SUM(Dim_Debt[MinimumPayment])"),
        ("Avg Monthly Savings Transfers", '=SUMIFS(Fact_SavingsTransfers[Amount],Fact_SavingsTransfers[Date],">="&EOMONTH(TodayDate,-4)+1,Fact_SavingsTransfers[Date],"<="&EOMONTH(TodayDate,-1))/3'),
        ("Starting Balance (Checking + Savings)", '=SUMIFS(Dim_Account[Balance],Dim_Account[AccountType],"Checking")+SUMIFS(Dim_Account[Balance],Dim_Account[AccountType],"Savings")'),
    ]
    assump = {}
    for i, (label, formula) in enumerate(labels):
        rr = r + i
        ws.cell(row=rr, column=1, value=label).font = f(10)
        c = ws.cell(row=rr, column=2, value=formula)
        c.number_format = CUR_FMT
        c.font = f(10, bold=True)
        assump[label] = f"$B${rr}"

    r2 = r + len(labels) + 2
    r2 = style_section_header(ws, r2, 1, "12-Month Rolling Forecast", span=8)
    headers = ["Month", "Starting Balance", "Projected Income", "Projected Expenses",
               "Debt Payments", "Savings Transfers", "Ending Balance", "Status"]
    for j, h in enumerate(headers):
        c = ws.cell(row=r2, column=1 + j, value=h)
        c.font = f(10, bold=True, color=WHITE); c.fill = fill(NAVY); c.alignment = CENTER_WRAP
    fr = r2 + 1
    for i in range(12):
        rr = fr + i
        ws.cell(row=rr, column=1, value=f"=EDATE(TodayDate,{i+1})")
        ws.cell(row=rr, column=1).number_format = "mmm-yy"
        if i == 0:
            ws.cell(row=rr, column=2, value=f"={assump['Starting Balance (Checking + Savings)']}")
        else:
            ws.cell(row=rr, column=2, value=f"=G{rr-1}")
        ws.cell(row=rr, column=3, value=f"={assump['Avg Monthly Income']}")
        ws.cell(row=rr, column=4, value=f"={assump['Avg Monthly Expenses']}")
        ws.cell(row=rr, column=5, value=f"={assump['Monthly Debt Payments (min.)']}")
        ws.cell(row=rr, column=6, value=f"={assump['Avg Monthly Savings Transfers']}")
        ws.cell(row=rr, column=7, value=f"=B{rr}+C{rr}-D{rr}-E{rr}-F{rr}")
        ws.cell(row=rr, column=8, value=f'=IF(G{rr}<0,"Negative Cash Flow","OK")')
        for c in (2, 3, 4, 5, 6, 7):
            ws.cell(row=rr, column=c).number_format = CUR_FMT0
    lr = fr + 11
    ws.conditional_formatting.add(f"G{fr}:G{lr}", CellIsRule(operator="lessThan", formula=["0"], fill=fill(RED_LIGHT), font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(f"H{fr}:H{lr}", CellIsRule(operator="equal", formula=['"Negative Cash Flow"'], fill=fill(RED_LIGHT), font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(f"H{fr}:H{lr}", CellIsRule(operator="equal", formula=['"OK"'], font=Font(color=GREEN)))

    ws.sheet_properties.tabColor = ORANGE
    return {"fr": fr, "lr": lr}


def build_financial_health(wb):
    ws = wb.create_sheet("Financial Health")
    style_sheet_title(ws, "Financial Health", span_cols=8)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 40

    r = 4
    r = style_section_header(ws, r, 1, "Score Factors (0-100 each, weighted to a composite 0-100 score)", span=5)
    headers = ["Factor", "Raw Value", "Score (0-100)", "Weight", "How it's scored"]
    for j, h in enumerate(headers):
        c = ws.cell(row=r, column=1 + j, value=h)
        c.font = f(10, bold=True, color=WHITE); c.fill = fill(NAVY)
    fr = r + 1
    factors = [
        ("Savings Rate", '=IFERROR((SUMIFS(Fact_Income[NetAmount],Fact_Income[Date],">="&CurMonthStart,Fact_Income[Date],"<="&CurMonthEnd)-SUMIFS(Fact_Expenses[Amount],Fact_Expenses[Date],">="&CurMonthStart,Fact_Expenses[Date],"<="&CurMonthEnd))/SUMIFS(Fact_Income[NetAmount],Fact_Income[Date],">="&CurMonthStart,Fact_Income[Date],"<="&CurMonthEnd),0)',
         PCT_FMT, 0.20, "100 pts at 20%+ savings rate, scaled linearly"),
        ("Debt-to-Income Ratio", '=IFERROR(SUM(Dim_Debt[MinimumPayment])/SUMIFS(Fact_Income[NetAmount],Fact_Income[Date],">="&CurMonthStart,Fact_Income[Date],"<="&CurMonthEnd),0)',
         PCT_FMT, 0.15, "100 pts at 0% DTI, 0 pts at 36%+ DTI"),
        ("Emergency Fund (months)", '=IFERROR(SUMIFS(Fact_SavingsTransfers[Amount],Fact_SavingsTransfers[GoalName],"Emergency Fund")/(SUMIFS(Fact_Expenses[Amount],Fact_Expenses[Date],">="&EOMONTH(TodayDate,-4)+1,Fact_Expenses[Date],"<="&EOMONTH(TodayDate,-1))/3),0)',
         '0.0"  mo"', 0.15, "100 pts at 6+ months of expenses saved"),
        ("Budget Compliance", "=IFERROR(1-MAX(SUM(Budget[Actual])/SUM(Budget[MonthlyBudget])-1,0),1)",
         PCT_FMT, 0.10, "100 pts if at/under budget, declines as overspend grows"),
        ("Net Worth Growth (mo.)", "=NetWorthGrowthPct", PCT_FMT, 0.10, "100 pts at 2%+ monthly growth, 0 pts at -2% or worse"),
        ("Investment Return", "=IFERROR(SUM(Fact_Investments[GainLoss])/SUM(Fact_Investments[CostBasis]),0)",
         PCT_FMT, 0.10, "100 pts at 10%+ return, 0 pts at 0% or below"),
        ("Bill Payment History", '=IFERROR(COUNTIFS(Fact_Bills[PaidStatus],"Paid")/COUNTA(Fact_Bills[BillID]),1)',
         PCT_FMT, 0.05, "% of logged bills marked Paid (not overdue)"),
        ("Need vs Want Ratio", '=IFERROR(SUMIFS(Fact_Expenses[Amount],Fact_Expenses[NeedWantFlag],"Need")/SUM(Fact_Expenses[Amount]),0)',
         PCT_FMT, 0.05, "100 pts at 70%+ spend on Needs"),
        ("Subscription Spending", "=IFERROR(SUM(Dim_Subscription[MonthlyCost])/SUMIFS(Fact_Income[NetAmount],Fact_Income[Date],\">=\"&CurMonthStart,Fact_Income[Date],\"<=\"&CurMonthEnd),0)",
         PCT_FMT, 0.05, "100 pts at 0% of income, 0 pts at 5%+ of income"),
        ("Impulse Spending", '=IFERROR(SUMIFS(Fact_Expenses[Amount],Fact_Expenses[ImpulseFlag],"Yes")/SUM(Fact_Expenses[Amount]),0)',
         PCT_FMT, 0.05, "100 pts at 0% impulse, 0 pts at 10%+ impulse"),
    ]
    score_formulas = [
        '=MIN($B{0}/0.2,1)*100',
        '=MAX(1-$B{0}/0.36,0)*100',
        '=MIN($B{0}/6,1)*100',
        '=$B{0}*100',
        '=MAX(MIN(($B{0}+0.02)/0.04,1),0)*100',
        '=MAX(MIN($B{0}/0.1,1),0)*100',
        '=$B{0}*100',
        '=MIN($B{0}/0.7,1)*100',
        '=MAX(1-$B{0}/0.05,0)*100',
        '=MAX(1-$B{0}/0.1,0)*100',
    ]
    for i, (label, raw_formula, raw_fmt, weight, note) in enumerate(factors):
        rr = fr + i
        ws.cell(row=rr, column=1, value=label).font = f(10)
        rc = ws.cell(row=rr, column=2, value=raw_formula)
        rc.number_format = raw_fmt
        sc = ws.cell(row=rr, column=3, value=score_formulas[i].format(rr))
        sc.number_format = '0.0'
        wc = ws.cell(row=rr, column=4, value=weight)
        wc.number_format = PCT_FMT
        nc = ws.cell(row=rr, column=5, value=note)
        nc.font = f(8, italic=True, color="6B7280")
        nc.alignment = LEFT_WRAP
    lr = fr + len(factors) - 1
    # fix Net Worth Growth row placeholder formula (needs net worth history rows, wired below)
    nwg_row = fr + 4

    score_row = lr + 2
    ws.cell(row=score_row, column=1, value="COMPOSITE FINANCIAL HEALTH SCORE").font = f(13, bold=True, color=WHITE)
    ws.cell(row=score_row, column=1).fill = fill(NAVY)
    score_cell = ws.cell(row=score_row, column=3, value=f"=SUMPRODUCT(C{fr}:C{lr},D{fr}:D{lr})")
    score_cell.font = f(16, bold=True, color=WHITE)
    score_cell.fill = fill(NAVY)
    score_cell.number_format = '0.0'
    ws.cell(row=score_row, column=2).fill = fill(NAVY)
    ws.cell(row=score_row, column=4).fill = fill(NAVY)
    ws.cell(row=score_row, column=5).fill = fill(NAVY)
    wb.defined_names["FinHealthScore"] = DefinedName("FinHealthScore", attr_text=f"'Financial Health'!$C${score_row}")

    rating_row = score_row + 1
    ws.cell(row=rating_row, column=1, value="Rating").font = f(10, bold=True)
    ws.cell(row=rating_row, column=3,
            value=f'=_xlfn.IFS(C{score_row}>=85,"Excellent",C{score_row}>=70,"Good",C{score_row}>=50,"Fair",TRUE,"Needs Attention")')
    ws.conditional_formatting.add(f"C{rating_row}", CellIsRule(operator="equal", formula=['"Excellent"'], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"C{rating_row}", CellIsRule(operator="equal", formula=['"Good"'], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"C{rating_row}", CellIsRule(operator="equal", formula=['"Fair"'], font=Font(color=ORANGE, bold=True)))
    ws.conditional_formatting.add(f"C{rating_row}", CellIsRule(operator="equal", formula=['"Needs Attention"'], font=Font(color=RED, bold=True)))

    # Gauge-style visual: horizontal stacked bar 0-100 built from cells + data bar CF
    gauge_row = rating_row + 2
    ws.cell(row=gauge_row, column=1, value="Score Gauge").font = f(10, bold=True)
    ws.merge_cells(start_row=gauge_row, start_column=2, end_row=gauge_row, end_column=6)
    gc = ws.cell(row=gauge_row, column=2, value=f"=C{score_row}/100")
    gc.number_format = '0.0%'
    ws.conditional_formatting.add(ws.cell(row=gauge_row, column=2).coordinate,
                                    DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=GREEN))

    r3 = gauge_row + 3
    r3 = style_section_header(ws, r3, 1, "Recommendations", span=5)
    recs = [
        (1, f'=IF(C{fr}<60,"Increase your savings rate — aim for at least 20% of net income.","Savings rate is healthy — keep it up.")'),
        (2, f'=IF(C{fr+1}<60,"Debt payments are consuming a large share of income — consider the debt avalanche method.","Debt load relative to income is under control.")'),
        (3, f'=IF(C{fr+2}<60,"Build your Emergency Fund toward 6 months of expenses.","Emergency fund is on track.")'),
        (4, f'=IF(C{fr+3}<60,"You are over budget in one or more categories — review the Budget tab.","Spending is within budget.")'),
        (5, f'=IF(C{fr+7}<60,"Too much spending is going to Wants — rebalance toward Needs and savings.","Need vs Want balance looks healthy.")'),
        (6, f'=IF(C{fr+8}<60,"Subscription costs are a meaningful share of income — review the Subscriptions tab for cuts.","Subscription spending is reasonable.")'),
        (7, f'=IF(C{fr+9}<60,"Impulse purchases are elevated — use the Cut-Back Analyzer on the Dashboard.","Impulse spending is well controlled.")'),
    ]
    for i, (n, formula) in enumerate(recs):
        rr = r3 + i
        ws.cell(row=rr, column=1, value=f"{n}.").font = f(9)
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=8)
        c = ws.cell(row=rr, column=2, value=formula)
        c.font = f(9)
        c.alignment = LEFT_WRAP

    ws.sheet_properties.tabColor = NAVY
    return {"fr": fr, "lr": lr, "score_row": score_row, "nwg_row": nwg_row}
