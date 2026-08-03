"""Trends sheet — compare every month, category, and vendor side by side.

Combines Fact_Expenses/Fact_Income (manually entered) with BankImportRaw (bank-derived,
netting Return rows against Expense rows) exactly like the Dashboard/Budget/Vendors totals do.
"""
import datetime as dt
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import LineChart, Reference
from common import *

MONTH_START = dt.date(2024, 1, 1)
N_MONTHS = 48  # Jan 2024 - Dec 2027: generous history + runway, avoids re-editing this sheet often
N_CAT_ROWS = 60      # padding above the current category count, for future growth
N_VENDOR_ROWS = 220  # padding above the current vendor count, for future growth


def _expense_sum(cat_or_vendor_col, criteria_cell, date_start_expr, date_end_expr, table_col):
    """SUMIFS(Fact_Expenses) + SUMIFS(BankImportRaw, Expense) - SUMIFS(BankImportRaw, Return),
    filtered by CategoryName or VendorName = criteria_cell and a date window."""
    return (
        f'SUMIFS(Fact_Expenses[Amount],Fact_Expenses[{table_col}],{criteria_cell},'
        f'Fact_Expenses[Date],">="&{date_start_expr},Fact_Expenses[Date],"<="&{date_end_expr})'
        f'+SUMIFS(BankImportRaw[AbsAmount],BankImportRaw[{table_col}],{criteria_cell},'
        f'BankImportRaw[TransactionType],"Expense",BankImportRaw[Date],">="&{date_start_expr},BankImportRaw[Date],"<="&{date_end_expr})'
        f'-SUMIFS(BankImportRaw[AbsAmount],BankImportRaw[{table_col}],{criteria_cell},'
        f'BankImportRaw[TransactionType],"Return",BankImportRaw[Date],">="&{date_start_expr},BankImportRaw[Date],"<="&{date_end_expr})'
    )


def build_trends(wb):
    ws = wb.create_sheet("Trends")
    style_sheet_title(ws, "Trends", subtitle="Compare every month, category, and vendor side by side", span_cols=16)
    ws.sheet_view.showGridLines = False

    def _month_at(i):
        y = MONTH_START.year + (MONTH_START.month - 1 + i) // 12
        m = (MONTH_START.month - 1 + i) % 12 + 1
        return dt.date(y, m, 1)

    # ================= Section 1: Monthly Summary =================
    r = style_section_header(ws, 4, 1,
        f"Monthly Summary  —  {_month_at(0).strftime('%b %Y')} through {_month_at(N_MONTHS - 1).strftime('%b %Y')}",
        span=7)
    headers1 = ["Month", "MonthEnd", "Income", "Expenses", "Net Savings", "Cumulative", "MoM Expense Change %"]
    for j, h in enumerate(headers1):
        c = ws.cell(row=r, column=1 + j, value=h)
        c.font = f(10, bold=True, color=WHITE)
        c.fill = fill(NAVY)
    m1_fr = r + 1
    for i in range(N_MONTHS):
        rr = m1_fr + i
        ws.cell(row=rr, column=1, value=_month_at(i))
        ws.cell(row=rr, column=1).number_format = "mmm-yy"
        ws.cell(row=rr, column=2, value=f"=EOMONTH(A{rr},0)")
        ws.cell(row=rr, column=2).number_format = DATE_FMT
        ws.cell(row=rr, column=3,
                value=(f'=SUMIFS(Fact_Income[NetAmount],Fact_Income[Date],">="&A{rr},Fact_Income[Date],"<="&B{rr})'
                       f'+SUMIFS(BankImportRaw[AbsAmount],BankImportRaw[TransactionType],"Income",BankImportRaw[Date],">="&A{rr},BankImportRaw[Date],"<="&B{rr})'))
        ws.cell(row=rr, column=4,
                value=(f'=SUMIFS(Fact_Expenses[Amount],Fact_Expenses[Date],">="&A{rr},Fact_Expenses[Date],"<="&B{rr})'
                       f'+SUMIFS(BankImportRaw[AbsAmount],BankImportRaw[TransactionType],"Expense",BankImportRaw[Date],">="&A{rr},BankImportRaw[Date],"<="&B{rr})'
                       f'-SUMIFS(BankImportRaw[AbsAmount],BankImportRaw[TransactionType],"Return",BankImportRaw[Date],">="&A{rr},BankImportRaw[Date],"<="&B{rr})'))
        ws.cell(row=rr, column=5, value=f"=C{rr}-D{rr}")
        ws.cell(row=rr, column=6, value=f"=SUM($E${m1_fr}:E{rr})")
        if i == 0:
            ws.cell(row=rr, column=7, value="")
        else:
            ws.cell(row=rr, column=7, value=f'=IFERROR(D{rr}/D{rr-1}-1,"")')
        for col in (3, 4, 5, 6):
            ws.cell(row=rr, column=col).number_format = CUR_FMT0
        ws.cell(row=rr, column=7).number_format = PCT_FMT
    m1_lr = m1_fr + N_MONTHS - 1
    for col in range(1, 8):
        ws.column_dimensions[get_column_letter(col)].width = 14

    # Chart: Income vs Expenses, full history
    ch = LineChart()
    ch.title = "Income vs Expenses — Full History"
    ch.height, ch.width = 9, 24
    cats = Reference(ws, min_col=1, min_row=m1_fr, max_col=1, max_row=m1_lr)
    ch.add_data(Reference(ws, min_col=3, min_row=r, max_col=4, max_row=m1_lr), titles_from_data=True)
    ch.set_categories(cats)
    ch.series[0].graphicalProperties.line.solidFill = GREEN
    ch.series[1].graphicalProperties.line.solidFill = RED
    ws.add_chart(ch, f"I{r}")

    # ================= Section 2: Category x Month matrix =================
    r2 = m1_lr + 20  # leave room for the chart above, which spans ~18 rows
    r2 = style_section_header(ws, r2, 1,
        "Category × Month  —  every category's spend by month (color scale highlights high/low months)",
        span=3 + N_MONTHS)
    ws.cell(row=r2, column=1, value="Category").font = f(10, bold=True, color=WHITE)
    ws.cell(row=r2, column=1).fill = fill(NAVY)
    month_col_first = 2
    for i in range(N_MONTHS):
        col = month_col_first + i
        c = ws.cell(row=r2, column=col, value=f"=B{m1_fr + i}")
        c.number_format = "mmm-yy"
        c.font = f(9, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        ws.column_dimensions[get_column_letter(col)].width = 11
    total_col = month_col_first + N_MONTHS
    ws.cell(row=r2, column=total_col, value="Total").font = f(10, bold=True, color=WHITE)
    ws.cell(row=r2, column=total_col).fill = fill(NAVY)
    ws.column_dimensions[get_column_letter(total_col)].width = 14
    ws.column_dimensions["A"].width = 20

    cat_fr = r2 + 1
    for j in range(N_CAT_ROWS):
        rr = cat_fr + j
        ws.cell(row=rr, column=1, value=f'=IFERROR(INDEX(Dim_Category[CategoryName],{j+1}),"")')
        for i in range(N_MONTHS):
            col = month_col_first + i
            hdr = get_column_letter(col) + str(r2)
            start_expr = f"EOMONTH({hdr},-1)+1"
            cell = ws.cell(row=rr, column=col,
                    value=f"=IF($A{rr}=\"\",\"\"," + _expense_sum("CategoryName", f"$A{rr}", start_expr, hdr, "CategoryName") + ")")
            cell.number_format = CUR_FMT0
        tot = ws.cell(row=rr, column=total_col,
                value=f'=IF($A{rr}="","",' + _expense_sum("CategoryName", f"$A{rr}", "DATE(1900,1,1)", "DATE(2100,12,31)", "CategoryName") + ")")
        tot.number_format = CUR_FMT0
        tot.font = f(9, bold=True)
    cat_lr = cat_fr + N_CAT_ROWS - 1
    ws.conditional_formatting.add(
        f"{get_column_letter(month_col_first)}{cat_fr}:{get_column_letter(month_col_first + N_MONTHS - 1)}{cat_lr}",
        ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="F87171"))

    # ================= Section 3: Category Drilldown =================
    r3 = cat_lr + 2
    ws.cell(row=r3, column=1, value="Category trend — pick a category:").font = f(10, bold=True)
    ws.merge_cells(start_row=r3, start_column=1, end_row=r3, end_column=3)
    cat_pick_cell = ws.cell(row=r3, column=4, value="Groceries")
    cat_pick_cell.font = f(10, bold=True, color=BLUE_ACCENT)
    cat_pick_cell.fill = fill(LIGHT_GRAY)
    add_list_validation(ws, cat_pick_cell.coordinate, "=ExpenseCategoryList")
    wb.defined_names["TrendsCategoryPick"] = DefinedName("TrendsCategoryPick", attr_text=f"Trends!$D${r3}")

    r3b = r3 + 2
    ws.cell(row=r3b, column=1, value="Month").font = f(9, bold=True, color=WHITE)
    ws.cell(row=r3b, column=1).fill = fill(NAVY)
    ws.cell(row=r3b, column=2, value="Spend").font = f(9, bold=True, color=WHITE)
    ws.cell(row=r3b, column=2).fill = fill(NAVY)
    cd_fr = r3b + 1
    for i in range(N_MONTHS):
        rr = cd_fr + i
        ws.cell(row=rr, column=1, value=f"=A{m1_fr + i}")
        ws.cell(row=rr, column=1).number_format = "mmm-yy"
        hdr = f"B{m1_fr + i}"
        start_expr = f"EOMONTH({hdr},-1)+1"
        c = ws.cell(row=rr, column=2,
                value="=" + _expense_sum("CategoryName", "TrendsCategoryPick", start_expr, hdr, "CategoryName"))
        c.number_format = CUR_FMT0
    cd_lr = cd_fr + N_MONTHS - 1

    ch2 = LineChart()
    ch2.title = "Selected Category — Monthly Spend"
    ch2.height, ch2.width = 9, 20
    ch2.add_data(Reference(ws, min_col=2, min_row=r3b, max_col=2, max_row=cd_lr), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=1, min_row=cd_fr, max_col=1, max_row=cd_lr))
    ch2.series[0].graphicalProperties.line.solidFill = BLUE_ACCENT
    ws.add_chart(ch2, f"D{r3b}")

    # ================= Section 4: Vendor Summary =================
    r4 = cd_lr + 20
    r4 = style_section_header(ws, r4, 1,
        "Vendor Summary  —  every vendor, all-time. Click a column header to sort/filter.", span=7)
    headers4 = ["VendorName", "DefaultCategory", "Total (All-Time)", "YTD", "Avg Monthly (Trailing 12mo)",
                "Transactions", "Last Transaction"]
    for j, h in enumerate(headers4):
        c = ws.cell(row=r4, column=1 + j, value=h)
        c.font = f(9, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = CENTER_WRAP
    ven_fr = r4 + 1
    for j in range(N_VENDOR_ROWS):
        rr = ven_fr + j
        ws.cell(row=rr, column=1, value=f'=IFERROR(INDEX(Dim_Vendor[VendorName],{j+1}),"")')
        ws.cell(row=rr, column=2, value=f'=IFERROR(INDEX(Dim_Vendor[DefaultCategory],{j+1}),"")')
        total = _expense_sum("VendorName", f"$A{rr}", "DATE(1900,1,1)", "DATE(2100,12,31)", "VendorName")
        ws.cell(row=rr, column=3, value=f'=IF($A{rr}="","",' + total + ")")
        ws.cell(row=rr, column=3).number_format = CUR_FMT
        ytd = _expense_sum("VendorName", f"$A{rr}", "CurYearStart", "CurYearEnd", "VendorName")
        ws.cell(row=rr, column=4, value=f'=IF($A{rr}="","",' + ytd + ")")
        ws.cell(row=rr, column=4).number_format = CUR_FMT
        trail12 = _expense_sum("VendorName", f"$A{rr}", "EDATE(TodayDate,-12)", "TodayDate", "VendorName")
        ws.cell(row=rr, column=5, value=f'=IF($A{rr}="","",(' + trail12 + ")/12)")
        ws.cell(row=rr, column=5).number_format = CUR_FMT
        ws.cell(row=rr, column=6,
                value=(f'=IF($A{rr}="","",COUNTIFS(Fact_Expenses[VendorName],$A{rr})'
                       f'+COUNTIFS(BankImportRaw[VendorName],$A{rr},BankImportRaw[TransactionType],"Expense"))'))
        ws.cell(row=rr, column=7,
                value=(f'=IF(OR($A{rr}="",$F{rr}=0),"",MAX('
                       f'_xlfn.MAXIFS(Fact_Expenses[Date],Fact_Expenses[VendorName],$A{rr}),'
                       f'_xlfn.MAXIFS(BankImportRaw[Date],BankImportRaw[VendorName],$A{rr},BankImportRaw[TransactionType],"Expense")))'))
        ws.cell(row=rr, column=7).number_format = DATE_FMT
    ven_lr = ven_fr + N_VENDOR_ROWS - 1
    for col, w in zip(range(1, 8), [24, 20, 15, 13, 18, 12, 14]):
        ws.column_dimensions[get_column_letter(col)].width = w

    # ================= Section 5: Vendor Drilldown =================
    r5 = ven_lr + 2
    ws.cell(row=r5, column=1, value="Vendor trend — pick a vendor:").font = f(10, bold=True)
    ws.merge_cells(start_row=r5, start_column=1, end_row=r5, end_column=3)
    ven_pick_cell = ws.cell(row=r5, column=4, value="AMAZON")
    ven_pick_cell.font = f(10, bold=True, color=BLUE_ACCENT)
    ven_pick_cell.fill = fill(LIGHT_GRAY)
    add_list_validation(ws, ven_pick_cell.coordinate, "=VendorNameList")
    wb.defined_names["TrendsVendorPick"] = DefinedName("TrendsVendorPick", attr_text=f"Trends!$D${r5}")

    r5b = r5 + 2
    ws.cell(row=r5b, column=1, value="Month").font = f(9, bold=True, color=WHITE)
    ws.cell(row=r5b, column=1).fill = fill(NAVY)
    ws.cell(row=r5b, column=2, value="Spend").font = f(9, bold=True, color=WHITE)
    ws.cell(row=r5b, column=2).fill = fill(NAVY)
    vd_fr = r5b + 1
    for i in range(N_MONTHS):
        rr = vd_fr + i
        ws.cell(row=rr, column=1, value=f"=A{m1_fr + i}")
        ws.cell(row=rr, column=1).number_format = "mmm-yy"
        hdr = f"B{m1_fr + i}"
        start_expr = f"EOMONTH({hdr},-1)+1"
        c = ws.cell(row=rr, column=2,
                value="=" + _expense_sum("VendorName", "TrendsVendorPick", start_expr, hdr, "VendorName"))
        c.number_format = CUR_FMT0
    vd_lr = vd_fr + N_MONTHS - 1

    ch3 = LineChart()
    ch3.title = "Selected Vendor — Monthly Spend"
    ch3.height, ch3.width = 9, 20
    ch3.add_data(Reference(ws, min_col=2, min_row=r5b, max_col=2, max_row=vd_lr), titles_from_data=True)
    ch3.set_categories(Reference(ws, min_col=1, min_row=vd_fr, max_col=1, max_row=vd_lr))
    ch3.series[0].graphicalProperties.line.solidFill = ORANGE
    ws.add_chart(ch3, f"D{r5b}")

    ws.freeze_panes = "B2"
    ws.sheet_properties.tabColor = PURPLE
    return {"m1_fr": m1_fr, "m1_lr": m1_lr, "cat_fr": cat_fr, "cat_lr": cat_lr, "ven_fr": ven_fr, "ven_lr": ven_lr}
