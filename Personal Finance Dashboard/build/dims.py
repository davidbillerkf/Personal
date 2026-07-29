import datetime as dt
from openpyxl.workbook.defined_name import DefinedName
from common import *
import seed_data as sd


def build_settings(wb):
    ws = wb.create_sheet("Settings")
    style_sheet_title(ws, "Settings", span_cols=8)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 16

    r = style_section_header(ws, 4, 1, "System Date Anchors (used by all SUMIFS date filters)", span=5)
    labels = [
        ("Today's Date", "=TODAY()", "TodayDate"),
        ("Current Month Start", "=EOMONTH(TodayDate,-1)+1", "CurMonthStart"),
        ("Current Month End", "=EOMONTH(TodayDate,0)", "CurMonthEnd"),
        ("Last Month Start", "=EOMONTH(TodayDate,-2)+1", "LastMonthStart"),
        ("Last Month End", "=EOMONTH(TodayDate,-1)", "LastMonthEnd"),
        ("Current Year Start", "=DATE(YEAR(TodayDate),1,1)", "CurYearStart"),
        ("Current Year End", "=DATE(YEAR(TodayDate),12,31)", "CurYearEnd"),
        ("Last Year Start", "=DATE(YEAR(TodayDate)-1,1,1)", "LastYearStart"),
        ("Last Year End", "=DATE(YEAR(TodayDate)-1,12,31)", "LastYearEnd"),
        ("7 Days From Today", "=TodayDate+7", "SevenDaysOut"),
        ("30 Days From Today", "=TodayDate+30", "ThirtyDaysOut"),
    ]
    named_cells = {}
    for i, (label, formula, name) in enumerate(labels):
        row = r + i
        lc = ws.cell(row=row, column=1, value=label)
        lc.font = f(10)
        lc.alignment = LEFT
        vc = ws.cell(row=row, column=2, value=formula)
        vc.font = f(10, bold=True, color=BLUE_ACCENT)
        vc.number_format = DATE_FMT
        named_cells[name] = f"Settings!$B${row}"
    r2 = r + len(labels) + 1

    r2 = style_section_header(ws, r2 + 1, 1, "Dim_PaymentMethod", span=5)
    pm_headers = ["PaymentMethodID", "MethodName"]
    pm_info = add_table(ws, r2, 1, pm_headers, "Dim_PaymentMethod", sd.PAYMENT_METHODS, col_widths=[16, 20])
    r3 = pm_info["end_row"] + 2

    r3 = style_section_header(ws, r3, 1, "Controlled Vocabulary Lists (used for data-validation dropdowns)", span=5)
    lists = {
        "List_YesNo": ["Yes", "No"],
        "List_NeedWant": ["Need", "Want"],
        "List_Frequency": ["Weekly", "Bi-Weekly", "Monthly", "Quarterly", "Semi-Annual", "Annual", "One-Time"],
        "List_BillStatus": ["Paid", "Unpaid", "Overdue", "Scheduled"],
        "List_Priority": ["High", "Medium", "Low"],
        "List_KeepFlag": ["Yes", "No", "Evaluate"],
    }
    col = 1
    start_row = r3
    list_ranges = {}
    for name, values in lists.items():
        ws.cell(row=start_row, column=col, value=name).font = f(10, bold=True, color=WHITE)
        ws.cell(row=start_row, column=col).fill = fill(NAVY)
        for i, v in enumerate(values):
            ws.cell(row=start_row + 1 + i, column=col, value=v).font = f(10)
        first = start_row + 1
        last = start_row + len(values)
        list_ranges[name] = f"Settings!${get_column_letter(col)}${first}:${get_column_letter(col)}${last}"
        col += 1
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 16
    ws.column_dimensions["I"].width = 16
    ws.column_dimensions["J"].width = 14
    ws.column_dimensions["K"].width = 14

    # Theme note
    theme_row = start_row + 10
    style_section_header(ws, theme_row, 1, "Design Theme", span=5)
    theme_notes = [
        ("Header Navy", "#0B1F3A"), ("Positive Green", "#15803D"), ("Negative Red", "#B91C1C"),
        ("Warning Orange", "#C2650B"), ("Accent Blue", "#2563EB"), ("Card Background", "#FFFFFF"),
    ]
    for i, (k, v) in enumerate(theme_notes):
        ws.cell(row=theme_row + 1 + i, column=1, value=k).font = f(10)
        ws.cell(row=theme_row + 1 + i, column=2, value=v).font = f(10)

    for name, ref in named_cells.items():
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
    for name, ref in list_ranges.items():
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
    wb.defined_names["PaymentMethodList"] = DefinedName(
        "PaymentMethodList", attr_text="Dim_PaymentMethod[MethodName]")

    ws.sheet_properties.tabColor = "6B7280"
    return {"named": named_cells, "lists": list_ranges, "pm_info": pm_info}


def build_dim_date(wb):
    """Hidden helper sheet: Dim_Date, DateID = INT(date serial). 2023-01-01 .. 2030-12-31."""
    ws = wb.create_sheet("Dim_Date")
    headers = ["DateID", "Date", "Day", "Week", "Month", "MonthNumber", "Quarter", "Year", "FiscalYear"]
    for j, h in enumerate(headers):
        c = ws.cell(row=1, column=1 + j, value=h)
        c.font = f(10, bold=True, color=WHITE)
        c.fill = fill(NAVY)
    start = dt.date(2025, 1, 1)
    end = dt.date(2027, 12, 31)
    n_days = (end - start).days + 1
    row = 2
    cur = start
    # Write in bulk for performance
    for i in range(n_days):
        this_date = cur + dt.timedelta(days=i)
        r = row + i
        ws.cell(row=r, column=1, value=f"=INT(B{r})")
        dc = ws.cell(row=r, column=2, value=this_date)
        dc.number_format = DATE_FMT
        ws.cell(row=r, column=3, value=f'=TEXT(B{r},"ddd")')
        ws.cell(row=r, column=4, value=f"=ISOWEEKNUM(B{r})")
        ws.cell(row=r, column=5, value=f'=TEXT(B{r},"mmm")')
        ws.cell(row=r, column=6, value=f"=MONTH(B{r})")
        ws.cell(row=r, column=7, value=f"=ROUNDUP(MONTH(B{r})/3,0)")
        ws.cell(row=r, column=8, value=f"=YEAR(B{r})")
        ws.cell(row=r, column=9, value=f"=YEAR(B{r})")
    end_row = row + n_days - 1
    ref = f"A1:I{end_row}"
    tab = Table(displayName="Dim_Date", ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)
    ws.sheet_state = "hidden"
    return {"end_row": end_row}


def build_categories(wb):
    ws = wb.create_sheet("Categories")
    style_sheet_title(ws, "Categories", span_cols=6)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 16

    r = style_section_header(ws, 4, 1, "Dim_Category  —  add a row here to make a new category available everywhere", span=5)
    cat_info = add_table(ws, r, 1, ["CategoryID", "CategoryName", "ParentCategory", "CategoryType", "EssentialCategory"],
                          "Dim_Category", sd.CATEGORIES, col_widths=[12, 22, 16, 14, 16])
    r2 = cat_info["end_row"] + 3
    r2 = style_section_header(ws, r2, 1, "Dim_Subcategory  —  CategoryID links each subcategory to its parent category above", span=5)
    sub_info = add_table(ws, r2, 1, ["SubcategoryID", "CategoryID", "SubcategoryName"],
                          "Dim_Subcategory", sd.SUBCATEGORIES, col_widths=[14, 12, 24])

    # Named ranges for dropdown lists — structured references so newly added rows
    # (typed into the blank row directly below the table) appear automatically.
    wb.defined_names["CategoryNameList"] = DefinedName(
        "CategoryNameList", attr_text="Dim_Category[CategoryName]")
    wb.defined_names["ExpenseCategoryList"] = DefinedName(
        "ExpenseCategoryList", attr_text="Dim_Category[CategoryName]")
    wb.defined_names["SubcategoryNameList"] = DefinedName(
        "SubcategoryNameList", attr_text="Dim_Subcategory[SubcategoryName]")
    ws.sheet_properties.tabColor = "6B7280"
    return {"cat_info": cat_info, "sub_info": sub_info}


def build_vendors(wb):
    ws = wb.create_sheet("Vendors")
    style_sheet_title(ws, "Vendors", span_cols=6)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 16

    r = style_section_header(ws, 4, 1, "Dim_Vendor  —  add a row here to make a new vendor available in Expenses", span=6)
    ven_info = add_table(ws, r, 1, ["VendorID", "VendorName", "VendorType", "DefaultCategoryID",
                                     "CurMonthSpend", "YTDSpend"],
                          "Dim_Vendor", [list(v) + [None, None] for v in sd.VENDORS],
                          col_widths=[12, 26, 18, 18, 15, 13])
    fr, lr = ven_info["first_data_row"], ven_info["end_row"]
    for rr in range(fr, lr + 1):
        ws.cell(row=rr, column=5,
                value=f'=SUMIFS(Fact_Expenses[Amount],Fact_Expenses[VendorName],$B{rr},Fact_Expenses[Date],">="&CurMonthStart,Fact_Expenses[Date],"<="&CurMonthEnd)')
        ws.cell(row=rr, column=5).number_format = CUR_FMT
        ws.cell(row=rr, column=6,
                value=f'=SUMIFS(Fact_Expenses[Amount],Fact_Expenses[VendorName],$B{rr},Fact_Expenses[Date],">="&CurYearStart,Fact_Expenses[Date],"<="&CurYearEnd)')
        ws.cell(row=rr, column=6).number_format = CUR_FMT
    wb.defined_names["VendorNameList"] = DefinedName(
        "VendorNameList", attr_text="Dim_Vendor[VendorName]")
    ws.sheet_properties.tabColor = "6B7280"
    return {"ven_info": ven_info}


def build_accounts(wb):
    ws = wb.create_sheet("Accounts")
    style_sheet_title(ws, "Accounts", span_cols=6)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 16

    r = style_section_header(ws, 4, 1, "Dim_Account  —  add a row here to make a new account available everywhere", span=5)
    acc_info = add_table(ws, r, 1, ["AccountID", "AccountName", "AccountType", "Institution", "Balance"],
                          "Dim_Account", sd.ACCOUNTS, col_widths=[12, 24, 16, 20, 16],
                          number_formats={4: CUR_FMT})
    for rr in range(acc_info["first_data_row"], acc_info["end_row"] + 1):
        ws.cell(row=rr, column=5).number_format = CUR_FMT
        ws.cell(row=rr, column=5).font = f(10, color=BLUE_ACCENT)

    r2 = acc_info["end_row"] + 3
    r2 = style_section_header(ws, r2, 1, "Account Totals", span=5)
    labels = [("Total Cash (Checking + Savings)", "Checking,Savings"),
              ("Total Credit Card Debt", "Credit Card"),
              ("Total Investment Accounts", "Brokerage,Retirement")]
    ws.cell(row=r2, column=1, value="Total All Accounts").font = f(10, bold=True)
    ws.cell(row=r2, column=2, value=f"=SUM(Dim_Account[Balance])").font = f(10, bold=True)
    ws.cell(row=r2, column=2).number_format = CUR_FMT

    wb.defined_names["AccountNameList"] = DefinedName(
        "AccountNameList", attr_text="Dim_Account[AccountName]")
    ws.sheet_properties.tabColor = "6B7280"
    return {"acc_info": acc_info}
