"""Budget sheet (simplified planner scope)."""
from common import *
import seed_data as sd

EXPENSE_CATEGORIES = [c[1] for c in sd.CATEGORIES if c[2] == "Expense"]
BUDGET_AMOUNTS = {
    "Mortgage": 2200, "Utilities": 300, "Insurance Home": 150, "Car Lease/Finance": 450,
    "Car Expenses": 150, "Insurance Car": 200, "Gas": 200, "Groceries": 900,
    "Dining & Restaurants": 250, "Healthcare": 200, "Entertainment": 150, "Personal Care": 100,
    "Men's Clothing": 75, "Women's Clothing": 100, "Kid's Clothing": 100,
    "Kid's Accessories": 75, "Gifts": 100, "Toys": 75, "Cleaning Supplies": 50,
    "Amazon": 150, "Walmart": 100, "Target": 100, "Books": 50, "Legal Fees": 100, "Taxes": 200,
}


def build_budget(wb):
    ws = wb.create_sheet("Budget")
    style_sheet_title(ws, "Budget", span_cols=9)
    ws.sheet_view.showGridLines = False
    for col in range(1, 10):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions["A"].width = 22

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
        "Category Budgets  —  set MonthlyBudget per category; Actual pulls live from Expenses + Bank Import this month", span=8)
    headers = ["CategoryName", "MonthlyBudget", "Actual", "Remaining", "Variance %", "Status", "Progress"]
    data_rows = [[cat, BUDGET_AMOUNTS.get(cat, 100)] + [None] * 5 for cat in EXPENSE_CATEGORIES]
    bud_info = add_table(ws, r, 1, headers, "Budget", data_rows,
                          col_widths=[22, 14, 13, 13, 12, 14, 12])
    fr, lr = bud_info["first_data_row"], bud_info["end_row"]
    for rr in range(fr, lr + 1):
        ws.cell(row=rr, column=2).number_format = CUR_FMT
        ws.cell(row=rr, column=2).font = f(10, color=BLUE_ACCENT)
        ws.cell(row=rr, column=3,
                value=(f'=SUMIFS(Fact_Expenses[Amount],Fact_Expenses[CategoryName],$A{rr},Fact_Expenses[Date],">="&CurMonthStart,Fact_Expenses[Date],"<="&CurMonthEnd)'
                       f'+SUMIFS(BankImportRaw[AbsAmount],BankImportRaw[CategoryName],$A{rr},BankImportRaw[TransactionType],"Expense",BankImportRaw[Date],">="&CurMonthStart,BankImportRaw[Date],"<="&CurMonthEnd)'))
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
