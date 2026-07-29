"""Income, Expenses, Bills, Subscriptions sheets."""
from openpyxl.workbook.defined_name import DefinedName
from common import *
import seed_data as sd


def build_income(wb):
    ws = wb.create_sheet("Income")
    style_sheet_title(ws, "Income", span_cols=10)
    ws.sheet_view.showGridLines = False

    r = style_section_header(ws, 4, 1, "Dim_Source", span=2)
    sources = [(1, "Employer Payroll Inc"), (2, "ACME Freelance Client"), (3, "Vanguard Brokerage"), (4, "Other")]
    src_info = add_table(ws, r, 1, ["SourceID", "SourceName"], "Dim_Source", sources, col_widths=[10, 24])
    wb.defined_names["SourceNameList"] = DefinedName(
        "SourceNameList", attr_text="Dim_Source[SourceName]")

    r2 = src_info["end_row"] + 3
    r2 = style_section_header(ws, r2, 1,
        "Fact_Income  —  enter each paycheck/deposit here. Dropdowns keep entries consistent; ID columns (hidden) auto-lookup for reporting.",
        span=13)
    headers = ["IncomeID", "Date", "SourceName", "SourceID", "CategoryName", "CategoryID",
               "AccountName", "AccountID", "GrossAmount", "TaxesWithheld", "NetAmount", "RecurringFlag", "Notes"]
    data_rows = []
    for row in sd.INCOME_ROWS:
        date_, source, cat, acct, gross, tax, rec, notes = row
        data_rows.append([None, date_, source, None, cat, None, acct, None, gross, tax, None, rec, notes])
    inc_info = add_table(ws, r2, 1, headers, "Fact_Income", data_rows,
                          col_widths=[10, 12, 22, 10, 16, 10, 20, 10, 13, 13, 13, 13, 24])

    fr, lr = inc_info["first_data_row"], inc_info["end_row"]
    for rr in range(fr, lr + 1):
        ws.cell(row=rr, column=1, value=f"=ROW()-{fr-1}")
        ws.cell(row=rr, column=2).number_format = DATE_FMT
        ws.cell(row=rr, column=4, value=f'=IFERROR(INDEX(Dim_Source[SourceID],MATCH([@SourceName],Dim_Source[SourceName],0)),"")'.replace("[@SourceName]", f"$C{rr}"))
        ws.cell(row=rr, column=6, value=f'=IFERROR(INDEX(Dim_Category[CategoryID],MATCH($E{rr},Dim_Category[CategoryName],0)),"")')
        ws.cell(row=rr, column=8, value=f'=IFERROR(INDEX(Dim_Account[AccountID],MATCH($G{rr},Dim_Account[AccountName],0)),"")')
        for col in (9, 10, 11):
            ws.cell(row=rr, column=col).number_format = CUR_FMT
        ws.cell(row=rr, column=11, value=f"=$I{rr}-$J{rr}")
    # hide ID helper columns
    for col_letter in ("D", "F", "H"):
        ws.column_dimensions[col_letter].hidden = True

    # Data validation dropdowns
    add_list_validation(ws, f"C{fr}:C{lr+200}", "=SourceNameList")
    add_list_validation(ws, f"E{fr}:E{lr+200}", "=ExpenseCategoryList")
    add_list_validation(ws, f"G{fr}:G{lr+200}", "=AccountNameList")
    add_list_validation(ws, f"L{fr}:L{lr+200}", "=List_YesNo")

    ws.freeze_panes = f"A{fr}"
    ws.sheet_properties.tabColor = GREEN
    return {"inc_info": inc_info}


def build_expenses(wb):
    ws = wb.create_sheet("Expenses")
    style_sheet_title(ws, "Expenses", span_cols=16)
    ws.sheet_view.showGridLines = False

    r = style_section_header(ws, 4, 1,
        "Fact_Expenses  —  every purchase goes here. Pick Vendor/Category/Subcategory/Account/Payment Method from dropdowns; flags drive the Cut-Back Analyzer and Financial Health Score.",
        span=20)
    headers = ["ExpenseID", "Date", "VendorName", "VendorID", "CategoryName", "CategoryID",
               "SubcategoryName", "SubcategoryID", "AccountName", "AccountID", "PaymentMethod", "PaymentMethodID",
               "Amount", "TaxIncluded", "EssentialFlag", "NeedWantFlag", "ImpulseFlag", "CanCutFlag",
               "RecurringFlag", "ReimbursableFlag", "Notes", "CurMonthAmt"]
    data_rows = []
    for row in sd.EXPENSE_ROWS:
        (date_, vendor, cat, subcat, acct, pm, amount, tax, ess, nw, imp, cut, rec, reim, notes) = row
        data_rows.append([None, date_, vendor, None, cat, None, subcat, None, acct, None, pm, None,
                           amount, tax, ess, nw, imp, cut, rec, reim, notes, None])
    exp_info = add_table(ws, r, 1, headers, "Fact_Expenses", data_rows,
                          col_widths=[10, 12, 20, 10, 16, 10, 18, 12, 18, 10, 14, 14, 12, 11, 11, 11, 10, 10, 11, 13, 22, 13])

    fr, lr = exp_info["first_data_row"], exp_info["end_row"]
    for rr in range(fr, lr + 1):
        ws.cell(row=rr, column=1, value=f"=ROW()-{fr-1}")
        ws.cell(row=rr, column=2).number_format = DATE_FMT
        ws.cell(row=rr, column=4, value=f'=IFERROR(INDEX(Dim_Vendor[VendorID],MATCH($C{rr},Dim_Vendor[VendorName],0)),"")')
        ws.cell(row=rr, column=6, value=f'=IFERROR(INDEX(Dim_Category[CategoryID],MATCH($E{rr},Dim_Category[CategoryName],0)),"")')
        ws.cell(row=rr, column=8, value=f'=IFERROR(INDEX(Dim_Subcategory[SubcategoryID],MATCH($G{rr},Dim_Subcategory[SubcategoryName],0)),"")')
        ws.cell(row=rr, column=10, value=f'=IFERROR(INDEX(Dim_Account[AccountID],MATCH($I{rr},Dim_Account[AccountName],0)),"")')
        ws.cell(row=rr, column=12, value=f'=IFERROR(INDEX(Dim_PaymentMethod[PaymentMethodID],MATCH($K{rr},Dim_PaymentMethod[MethodName],0)),"")')
        ws.cell(row=rr, column=13).number_format = CUR_FMT
        ws.cell(row=rr, column=22,
                value=f'=IF(AND($B{rr}>=CurMonthStart,$B{rr}<=CurMonthEnd),$M{rr},0)')
        ws.cell(row=rr, column=22).number_format = CUR_FMT
    for col_letter in ("D", "F", "H", "J", "L", "V"):
        ws.column_dimensions[col_letter].hidden = True

    buf = 300
    add_list_validation(ws, f"C{fr}:C{lr+buf}", "=VendorNameList")
    add_list_validation(ws, f"E{fr}:E{lr+buf}", "=ExpenseCategoryList")
    add_list_validation(ws, f"G{fr}:G{lr+buf}", "=SubcategoryNameList")
    add_list_validation(ws, f"I{fr}:I{lr+buf}", "=AccountNameList")
    add_list_validation(ws, f"K{fr}:K{lr+buf}", "=PaymentMethodList")
    add_list_validation(ws, f"N{fr}:N{lr+buf}", "=List_YesNo")
    add_list_validation(ws, f"O{fr}:O{lr+buf}", "=List_YesNo")
    add_list_validation(ws, f"P{fr}:P{lr+buf}", "=List_NeedWant")
    add_list_validation(ws, f"Q{fr}:Q{lr+buf}", "=List_YesNo")
    add_list_validation(ws, f"R{fr}:R{lr+buf}", "=List_YesNo")
    add_list_validation(ws, f"S{fr}:S{lr+buf}", "=List_YesNo")
    add_list_validation(ws, f"T{fr}:T{lr+buf}", "=List_YesNo")

    # Conditional formatting: large purchases (>$200) highlighted orange; impulse flag red text
    ws.conditional_formatting.add(
        f"M{fr}:M{lr}",
        CellIsRule(operator="greaterThan", formula=["200"], fill=fill(ORANGE_LIGHT)))
    ws.conditional_formatting.add(
        f"Q{fr}:Q{lr}",
        FormulaRule(formula=[f'$Q{fr}="Yes"'], font=Font(color=RED, bold=True)))

    ws.freeze_panes = f"A{fr}"
    ws.sheet_properties.tabColor = RED
    return {"exp_info": exp_info}


def build_bills(wb):
    ws = wb.create_sheet("Bills")
    style_sheet_title(ws, "Bills", span_cols=13)
    ws.sheet_view.showGridLines = False

    r = style_section_header(ws, 4, 1, "Bill Summary", span=6)
    kpi_specs = [
        ("Upcoming (Next 7 Days)", '=SUMPRODUCT((Fact_Bills[DueDate]>=TodayDate)*(Fact_Bills[DueDate]<=SevenDaysOut)*(Fact_Bills[PaidStatus]<>"Paid")*Fact_Bills[Amount])'),
        ("Overdue", '=SUMPRODUCT((Fact_Bills[DueDate]<TodayDate)*(Fact_Bills[PaidStatus]<>"Paid")*Fact_Bills[Amount])'),
        ("Monthly Obligations", '=SUMPRODUCT((Fact_Bills[Frequency]="Monthly")*Fact_Bills[Amount])'),
        ("Annual Cost (All Bills, Monthly x12)", '=SUMPRODUCT((Fact_Bills[Frequency]="Monthly")*Fact_Bills[Amount])*12'),
    ]
    for i, (label, formula) in enumerate(kpi_specs):
        kpi_card(ws, 5, 1 + i * 3, label, formula, number_format=CUR_FMT0, width_cols=3)
    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 11

    r2 = 10
    r2 = style_section_header(ws, r2, 1,
        "Fact_Bills  —  recurring & one-time bills. PaidStatus/AutoPayFlag drive the Upcoming/Overdue KPIs above.", span=13)
    headers = ["BillID", "BillName", "DueDate", "CategoryName", "CategoryID", "AccountName", "AccountID",
               "Amount", "Frequency", "PaidStatus", "AutoPayFlag", "Reminder", "Notes", "UnpaidDueSort"]
    data_rows = []
    for row in sd.BILLS_ROWS:
        name_, due, cat, acct, amt, freq, status, auto, remind, notes = row
        data_rows.append([None, name_, due, cat, None, acct, None, amt, freq, status, auto, remind, notes, None])
    bill_info = add_table(ws, r2, 1, headers, "Fact_Bills", data_rows,
                           col_widths=[10, 22, 12, 16, 10, 20, 10, 12, 12, 12, 12, 16, 20, 13])

    fr, lr = bill_info["first_data_row"], bill_info["end_row"]
    for rr in range(fr, lr + 1):
        ws.cell(row=rr, column=1, value=f"=ROW()-{fr-1}")
        ws.cell(row=rr, column=3).number_format = DATE_FMT
        ws.cell(row=rr, column=5, value=f'=IFERROR(INDEX(Dim_Category[CategoryID],MATCH($D{rr},Dim_Category[CategoryName],0)),"")')
        ws.cell(row=rr, column=7, value=f'=IFERROR(INDEX(Dim_Account[AccountID],MATCH($F{rr},Dim_Account[AccountName],0)),"")')
        ws.cell(row=rr, column=8).number_format = CUR_FMT
        ws.cell(row=rr, column=14, value=f'=IF($J{rr}<>"Paid",$C{rr},"")')
        ws.cell(row=rr, column=14).number_format = DATE_FMT
    for col_letter in ("E", "G", "N"):
        ws.column_dimensions[col_letter].hidden = True

    buf = 200
    add_list_validation(ws, f"D{fr}:D{lr+buf}", "=ExpenseCategoryList")
    add_list_validation(ws, f"F{fr}:F{lr+buf}", "=AccountNameList")
    add_list_validation(ws, f"I{fr}:I{lr+buf}", "=List_Frequency")
    add_list_validation(ws, f"J{fr}:J{lr+buf}", "=List_BillStatus")
    add_list_validation(ws, f"K{fr}:K{lr+buf}", "=List_YesNo")

    # Conditional formatting on DueDate + PaidStatus
    ws.conditional_formatting.add(
        f"J{fr}:J{lr}", CellIsRule(operator="equal", formula=['"Paid"'], fill=fill(GREEN_LIGHT), font=Font(color=GREEN)))
    ws.conditional_formatting.add(
        f"J{fr}:J{lr}", CellIsRule(operator="equal", formula=['"Overdue"'], fill=fill(RED_LIGHT), font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(
        f"C{fr}:C{lr}",
        FormulaRule(formula=[f'AND($C{fr}<TodayDate,$J{fr}<>"Paid")'], fill=fill(RED_LIGHT)))
    ws.conditional_formatting.add(
        f"C{fr}:C{lr}",
        FormulaRule(formula=[f'AND($C{fr}>=TodayDate,$C{fr}<=SevenDaysOut,$J{fr}<>"Paid")'], fill=fill(ORANGE_LIGHT)))

    ws.freeze_panes = f"A{fr}"
    ws.sheet_properties.tabColor = ORANGE
    return {"bill_info": bill_info}


def build_subscriptions(wb):
    ws = wb.create_sheet("Subscriptions")
    style_sheet_title(ws, "Subscriptions", span_cols=8)
    ws.sheet_view.showGridLines = False

    r = style_section_header(ws, 4, 1, "Subscription Summary", span=6)
    kpi_specs = [
        ("Total Monthly Cost", "=SUM(Dim_Subscription[MonthlyCost])"),
        ("Total Annual Cost", "=SUM(Dim_Subscription[AnnualCost])"),
        ("Potential Savings (KeepFlag = No)", '=SUMIFS(Dim_Subscription[AnnualCost],Dim_Subscription[KeepFlag],"No")'),
    ]
    for i, (label, formula) in enumerate(kpi_specs):
        kpi_card(ws, 5, 1 + i * 3, label, formula, number_format=CUR_FMT0, width_cols=3)
    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 14

    r2 = 10
    r2 = style_section_header(ws, r2, 1, "Dim_Subscription  —  every recurring subscription. Set KeepFlag to No to flag for cancellation.", span=8)
    headers = ["SubscriptionID", "ServiceName", "CategoryName", "CategoryID", "MonthlyCost", "AnnualCost", "KeepFlag"]
    data_rows = []
    for (sid, name_, catid, mc, ac, keep) in sd.SUBSCRIPTIONS:
        catname = next((c[1] for c in sd.CATEGORIES if c[0] == catid), "")
        data_rows.append([sid, name_, catname, catid, mc, ac, keep])
    sub_info = add_table(ws, r2, 1, headers, "Dim_Subscription", data_rows,
                          col_widths=[14, 24, 16, 10, 13, 13, 12])
    fr, lr = sub_info["first_data_row"], sub_info["end_row"]
    for rr in range(fr, lr + 1):
        ws.cell(row=rr, column=5).number_format = CUR_FMT
        ws.cell(row=rr, column=6).number_format = CUR_FMT
        ws.cell(row=rr, column=6, value=f"=$E{rr}*12")
    ws.column_dimensions["D"].hidden = True
    buf = 100
    add_list_validation(ws, f"C{fr}:C{lr+buf}", "=ExpenseCategoryList")
    add_list_validation(ws, f"G{fr}:G{lr+buf}", "=List_KeepFlag")
    ws.conditional_formatting.add(
        f"G{fr}:G{lr}", CellIsRule(operator="equal", formula=['"No"'], fill=fill(RED_LIGHT), font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(
        f"G{fr}:G{lr}", CellIsRule(operator="equal", formula=['"Evaluate"'], fill=fill(ORANGE_LIGHT)))

    ws.freeze_panes = f"A{fr}"
    ws.sheet_properties.tabColor = PURPLE
    return {"sub_info": sub_info}
