"""Debt, Savings Goals, Investments, Net Worth sheets."""
from openpyxl.workbook.defined_name import DefinedName
from common import *
import seed_data as sd


def build_debt(wb):
    ws = wb.create_sheet("Debt")
    style_sheet_title(ws, "Debt", span_cols=12)
    ws.sheet_view.showGridLines = False

    kpi_specs = [
        ("Total Debt Remaining", "=SUM(Dim_Debt[CurrentBalance])"),
        ("Total Minimum Payments / mo", "=SUM(Dim_Debt[MinimumPayment])"),
        ("Weighted Avg Interest Rate", "=SUMPRODUCT(Dim_Debt[InterestRate],Dim_Debt[CurrentBalance])/SUM(Dim_Debt[CurrentBalance])"),
        ("Total Paid Off To Date", "=SUM(Dim_Debt[OriginalBalance])-SUM(Dim_Debt[CurrentBalance])"),
    ]
    for i, (label, formula) in enumerate(kpi_specs):
        fmt = PCT_FMT if "Rate" in label else CUR_FMT0
        kpi_card(ws, 4, 1 + i * 3, label, formula, number_format=fmt, width_cols=3)
    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 13

    r = 9
    r = style_section_header(ws, r, 1, "Dim_Debt  —  every loan / credit balance you owe", span=12)
    headers = ["DebtID", "DebtName", "DebtType", "InterestRate", "OriginalBalance", "CurrentBalance",
               "MinimumPayment", "% Paid Off", "Snowball Rank", "Avalanche Rank"]
    data_rows = [list(row) + [None, None, None] for row in sd.DEBTS]
    debt_info = add_table(ws, r, 1, headers, "Dim_Debt", data_rows,
                           col_widths=[9, 26, 16, 13, 15, 15, 15, 12, 13, 13])
    fr, lr = debt_info["first_data_row"], debt_info["end_row"]
    for rr in range(fr, lr + 1):
        ws.cell(row=rr, column=4).number_format = PCT_FMT
        ws.cell(row=rr, column=5).number_format = CUR_FMT
        ws.cell(row=rr, column=6).number_format = CUR_FMT
        ws.cell(row=rr, column=7).number_format = CUR_FMT
        ws.cell(row=rr, column=8, value=f"=1-($F{rr}/$E{rr})")
        ws.cell(row=rr, column=8).number_format = PCT_FMT
        # Snowball = rank by smallest current balance first; Avalanche = rank by highest interest rate first
        ws.cell(row=rr, column=9, value=f"=SUMPRODUCT((Dim_Debt[CurrentBalance]<$F{rr})*1)+1")
        ws.cell(row=rr, column=10, value=f"=SUMPRODUCT((Dim_Debt[InterestRate]>$D{rr})*1)+1")
    ws.conditional_formatting.add(f"H{fr}:H{lr}", DataBarRule(start_type="num", start_value=0, end_type="num",
                                                                end_value=1, color=GREEN))
    dv = DataValidation(type="list", formula1='"Credit Card,Mortgage,Car Loan,Student Loan,Personal Loan"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"C{fr}:C{lr+50}")

    r2 = lr + 3
    r2 = style_section_header(ws, r2, 1, "Payoff Strategy Comparison", span=6)
    ws.cell(row=r2, column=1, value="Strategy").font = f(10, bold=True)
    ws.cell(row=r2, column=2, value="Payoff Order (by rank 1 = first)").font = f(10, bold=True)
    ws.cell(row=r2, column=3, value="Est. Months to Debt-Free*").font = f(10, bold=True)
    ws.cell(row=r2 + 1, column=1, value="Debt Snowball (smallest balance first)").font = f(10)
    ws.cell(row=r2 + 1, column=2, value='=TEXTJOIN(" > ",TRUE,INDEX(Dim_Debt[DebtName],MATCH(1,Dim_Debt[Snowball Rank],0)),INDEX(Dim_Debt[DebtName],MATCH(2,Dim_Debt[Snowball Rank],0)),INDEX(Dim_Debt[DebtName],MATCH(3,Dim_Debt[Snowball Rank],0)),INDEX(Dim_Debt[DebtName],MATCH(4,Dim_Debt[Snowball Rank],0)),INDEX(Dim_Debt[DebtName],MATCH(5,Dim_Debt[Snowball Rank],0)))'
                 .replace("TEXTJOIN", "_xlfn.TEXTJOIN"))
    ws.cell(row=r2 + 1, column=3, value="=ROUNDUP(SUM(Dim_Debt[CurrentBalance])/SUM(Dim_Debt[MinimumPayment]),0)")
    ws.cell(row=r2 + 2, column=1, value="Debt Avalanche (highest interest first)").font = f(10)
    ws.cell(row=r2 + 2, column=2, value='=TEXTJOIN(" > ",TRUE,INDEX(Dim_Debt[DebtName],MATCH(1,Dim_Debt[Avalanche Rank],0)),INDEX(Dim_Debt[DebtName],MATCH(2,Dim_Debt[Avalanche Rank],0)),INDEX(Dim_Debt[DebtName],MATCH(3,Dim_Debt[Avalanche Rank],0)),INDEX(Dim_Debt[DebtName],MATCH(4,Dim_Debt[Avalanche Rank],0)),INDEX(Dim_Debt[DebtName],MATCH(5,Dim_Debt[Avalanche Rank],0)))'
                 .replace("TEXTJOIN", "_xlfn.TEXTJOIN"))
    ws.cell(row=r2 + 2, column=3, value="=ROUNDUP(SUM(Dim_Debt[CurrentBalance])/SUM(Dim_Debt[MinimumPayment]),0)")
    ws.cell(row=r2 + 4, column=1,
            value="*Simplified estimate assuming only minimum payments are made across all debts combined; avalanche saves more interest, snowball builds momentum faster psychologically.").font = f(8, italic=True, color="6B7280")
    ws.merge_cells(start_row=r2+4, start_column=1, end_row=r2+4, end_column=8)

    r3 = r2 + 6
    r3 = style_section_header(ws, r3, 1, "Fact_DebtPayments  —  log every payment made toward a debt", span=9)
    headers2 = ["PaymentID", "Date", "DebtName", "DebtID", "AccountName", "AccountID", "PaymentAmount",
                "PrincipalAmount", "InterestAmount", "Notes"]
    data_rows2 = []
    for (date_, debtname, acct, pay, princ) in sd.DEBT_PAYMENTS_ROWS:
        data_rows2.append([None, date_, debtname, None, acct, None, pay, princ, None, ""])
    pay_info = add_table(ws, r3, 1, headers2, "Fact_DebtPayments", data_rows2,
                          col_widths=[10, 12, 26, 9, 20, 10, 14, 14, 14, 18])
    fr2, lr2 = pay_info["first_data_row"], pay_info["end_row"]
    for rr in range(fr2, lr2 + 1):
        ws.cell(row=rr, column=1, value=f"=ROW()-{fr2-1}")
        ws.cell(row=rr, column=2).number_format = DATE_FMT
        ws.cell(row=rr, column=4, value=f'=IFERROR(INDEX(Dim_Debt[DebtID],MATCH($C{rr},Dim_Debt[DebtName],0)),"")')
        ws.cell(row=rr, column=6, value=f'=IFERROR(INDEX(Dim_Account[AccountID],MATCH($E{rr},Dim_Account[AccountName],0)),"")')
        for c in (7, 8, 9):
            ws.cell(row=rr, column=c).number_format = CUR_FMT
        ws.cell(row=rr, column=9, value=f"=$G{rr}-$H{rr}")
    for col_letter in ("D", "F"):
        ws.column_dimensions[col_letter].hidden = True
    add_list_validation(ws, f"C{fr2}:C{lr2+100}", "=Dim_Debt[DebtName]")
    add_list_validation(ws, f"E{fr2}:E{lr2+100}", "=AccountNameList")

    ws.sheet_properties.tabColor = RED
    return {"debt_info": debt_info, "pay_info": pay_info}


def build_savings_goals(wb):
    ws = wb.create_sheet("Savings Goals")
    style_sheet_title(ws, "Savings Goals", span_cols=11)
    ws.sheet_view.showGridLines = False

    r = 4
    r = style_section_header(ws, r, 1, "Dim_Goal + progress  —  set a target; log transfers below to track progress automatically", span=11)
    headers = ["GoalID", "GoalName", "TargetAmount", "TargetDate", "Priority", "CurrentAmount",
               "Remaining", "Completion %", "Avg Monthly Contribution", "Projected Completion Date"]
    data_rows = [list(row) + [None, None, None, None, None] for row in sd.GOALS]
    goal_info = add_table(ws, r, 1, headers, "Dim_Goal", data_rows,
                           col_widths=[9, 22, 15, 13, 11, 15, 14, 13, 20, 20])
    fr, lr = goal_info["first_data_row"], goal_info["end_row"]
    for rr in range(fr, lr + 1):
        ws.cell(row=rr, column=3).number_format = CUR_FMT
        ws.cell(row=rr, column=4).number_format = DATE_FMT
        ws.cell(row=rr, column=6, value=f'=SUMIFS(Fact_SavingsTransfers[Amount],Fact_SavingsTransfers[GoalName],$B{rr})')
        ws.cell(row=rr, column=6).number_format = CUR_FMT
        ws.cell(row=rr, column=7, value=f"=MAX($C{rr}-$F{rr},0)")
        ws.cell(row=rr, column=7).number_format = CUR_FMT
        ws.cell(row=rr, column=8, value=f"=MIN($F{rr}/$C{rr},1)")
        ws.cell(row=rr, column=8).number_format = PCT_FMT
        ws.cell(row=rr, column=9,
                value=f'=IFERROR($F{rr}/MAX(COUNTIFS(Fact_SavingsTransfers[GoalName],$B{rr}),1),0)')
        ws.cell(row=rr, column=9).number_format = CUR_FMT
        ws.cell(row=rr, column=10,
                value=f'=IF($I{rr}>0,TodayDate+($G{rr}/$I{rr})*30,"N/A")')
        ws.cell(row=rr, column=10).number_format = DATE_FMT
    ws.conditional_formatting.add(f"H{fr}:H{lr}", DataBarRule(start_type="num", start_value=0, end_type="num",
                                                                end_value=1, color=BLUE_ACCENT))
    ws.conditional_formatting.add(f"H{fr}:H{lr}", CellIsRule(operator="greaterThanOrEqual", formula=["1"],
                                                               fill=fill(GREEN_LIGHT)))
    dv = DataValidation(type="list", formula1="=List_Priority", allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"E{fr}:E{lr+50}")

    r2 = lr + 3
    r2 = style_section_header(ws, r2, 1, "Fact_SavingsTransfers  —  log every deposit toward a goal", span=6)
    headers2 = ["TransferID", "Date", "GoalName", "GoalID", "AccountName", "AccountID", "Amount"]
    data_rows2 = []
    for (date_, goal, acct, amt) in sd.SAVINGS_TRANSFER_ROWS:
        data_rows2.append([None, date_, goal, None, acct, None, amt])
    trans_info = add_table(ws, r2, 1, headers2, "Fact_SavingsTransfers", data_rows2,
                            col_widths=[11, 12, 20, 9, 20, 10, 13])
    fr2, lr2 = trans_info["first_data_row"], trans_info["end_row"]
    for rr in range(fr2, lr2 + 1):
        ws.cell(row=rr, column=1, value=f"=ROW()-{fr2-1}")
        ws.cell(row=rr, column=2).number_format = DATE_FMT
        ws.cell(row=rr, column=4, value=f'=IFERROR(INDEX(Dim_Goal[GoalID],MATCH($C{rr},Dim_Goal[GoalName],0)),"")')
        ws.cell(row=rr, column=6, value=f'=IFERROR(INDEX(Dim_Account[AccountID],MATCH($E{rr},Dim_Account[AccountName],0)),"")')
        ws.cell(row=rr, column=7).number_format = CUR_FMT
    for col_letter in ("D", "F"):
        ws.column_dimensions[col_letter].hidden = True
    add_list_validation(ws, f"C{fr2}:C{lr2+150}", "=Dim_Goal[GoalName]")
    add_list_validation(ws, f"E{fr2}:E{lr2+150}", "=AccountNameList")

    ws.sheet_properties.tabColor = BLUE_ACCENT
    return {"goal_info": goal_info, "trans_info": trans_info}


def build_investments(wb):
    ws = wb.create_sheet("Investments")
    style_sheet_title(ws, "Investments", span_cols=11)
    ws.sheet_view.showGridLines = False

    kpi_specs = [
        ("Total Market Value", "=SUM(Fact_Investments[MarketValue])"),
        ("Total Cost Basis", "=SUM(Fact_Investments[CostBasis])"),
        ("Total Gain / Loss", "=SUM(Fact_Investments[GainLoss])"),
        ("Return %", "=SUM(Fact_Investments[GainLoss])/SUM(Fact_Investments[CostBasis])"),
    ]
    for i, (label, formula) in enumerate(kpi_specs):
        fmt = PCT_FMT if "%" in label else CUR_FMT0
        kpi_card(ws, 4, 1 + i * 3, label, formula, number_format=fmt, width_cols=3)
    for col in range(1, 12):
        ws.column_dimensions[get_column_letter(col)].width = 13

    r = 9
    r = style_section_header(ws, r, 1, "Dim_Investment", span=4)
    inv_info = add_table(ws, r, 1, ["InvestmentID", "InvestmentName", "Ticker", "AssetClass"],
                          "Dim_Investment", sd.INVESTMENTS, col_widths=[12, 30, 10, 16])
    fr0, lr0 = inv_info["first_data_row"], inv_info["end_row"]

    r2 = lr0 + 3
    r2 = style_section_header(ws, r2, 1,
        "Fact_Investments  —  log each holding/lot. MarketValue & GainLoss calculate automatically.", span=10)
    headers = ["TransactionID", "Date", "InvestmentName", "InvestmentID", "AccountName", "AccountID",
               "Shares", "Price", "MarketValue", "CostBasis", "GainLoss", "Allocation %", "AssetClass"]
    data_rows = []
    for (date_, name_, acct, shares, price, cost) in sd.INVESTMENT_ROWS:
        data_rows.append([None, date_, name_, None, acct, None, shares, price, None, cost, None, None, None])
    fi_info = add_table(ws, r2, 1, headers, "Fact_Investments", data_rows,
                         col_widths=[13, 12, 30, 12, 18, 10, 10, 11, 14, 13, 13, 12, 16])
    fr, lr = fi_info["first_data_row"], fi_info["end_row"]
    for rr in range(fr, lr + 1):
        ws.cell(row=rr, column=1, value=f"=ROW()-{fr-1}")
        ws.cell(row=rr, column=2).number_format = DATE_FMT
        ws.cell(row=rr, column=4, value=f'=IFERROR(INDEX(Dim_Investment[InvestmentID],MATCH($C{rr},Dim_Investment[InvestmentName],0)),"")')
        ws.cell(row=rr, column=6, value=f'=IFERROR(INDEX(Dim_Account[AccountID],MATCH($E{rr},Dim_Account[AccountName],0)),"")')
        ws.cell(row=rr, column=7).number_format = '#,##0.0000'
        ws.cell(row=rr, column=8).number_format = CUR_FMT
        ws.cell(row=rr, column=9, value=f"=$G{rr}*$H{rr}")
        ws.cell(row=rr, column=9).number_format = CUR_FMT
        ws.cell(row=rr, column=10).number_format = CUR_FMT
        ws.cell(row=rr, column=11, value=f"=$I{rr}-$J{rr}")
        ws.cell(row=rr, column=11).number_format = CUR_FMT
        ws.cell(row=rr, column=12, value=f"=$I{rr}/SUM(Fact_Investments[MarketValue])")
        ws.cell(row=rr, column=12).number_format = PCT_FMT
        ws.cell(row=rr, column=13, value=f'=IFERROR(INDEX(Dim_Investment[AssetClass],MATCH($C{rr},Dim_Investment[InvestmentName],0)),"")')
    for col_letter in ("D", "F", "M"):
        ws.column_dimensions[col_letter].hidden = True
    add_list_validation(ws, f"C{fr}:C{lr+100}", "=Dim_Investment[InvestmentName]")
    add_list_validation(ws, f"E{fr}:E{lr+100}", "=AccountNameList")
    ws.conditional_formatting.add(f"K{fr}:K{lr}", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED)))
    ws.conditional_formatting.add(f"K{fr}:K{lr}", CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=Font(color=GREEN)))

    r3 = lr + 3
    r3 = style_section_header(ws, r3, 1, "Allocation by Asset Class", span=3)
    classes = sorted(set(x[3] for x in sd.INVESTMENTS))
    ws.cell(row=r3, column=1, value="Asset Class").font = f(10, bold=True)
    ws.cell(row=r3, column=2, value="Market Value").font = f(10, bold=True)
    ws.cell(row=r3, column=3, value="Allocation %").font = f(10, bold=True)
    for i, ac in enumerate(classes):
        rr = r3 + 1 + i
        ws.cell(row=rr, column=1, value=ac)
        ws.cell(row=rr, column=2,
                value=f'=SUMIFS(Fact_Investments[MarketValue],Fact_Investments[AssetClass],$A{rr})')
        ws.cell(row=rr, column=2).number_format = CUR_FMT
        ws.cell(row=rr, column=3, value=f"=$B{rr}/SUM(Fact_Investments[MarketValue])")
        ws.cell(row=rr, column=3).number_format = PCT_FMT
    allocation_range = (r3 + 1, r3 + len(classes))

    ws.sheet_properties.tabColor = GREEN
    return {"inv_info": inv_info, "fi_info": fi_info, "alloc_start": allocation_range[0], "alloc_end": allocation_range[1]}


def build_net_worth(wb):
    ws = wb.create_sheet("Net Worth")
    style_sheet_title(ws, "Net Worth", span_cols=10)
    ws.sheet_view.showGridLines = False

    kpi_specs = [
        ("Total Assets", None),
        ("Total Liabilities", None),
        ("Net Worth", None),
    ]
    for col in range(1, 11):
        ws.column_dimensions[get_column_letter(col)].width = 15

    r = 4
    r = style_section_header(ws, r, 1, "Assets", span=3)
    ws.cell(row=r, column=1, value="Asset").font = f(10, bold=True); ws.cell(row=r, column=1).fill = fill(NAVY); ws.cell(row=r, column=1).font=f(10,bold=True,color=WHITE)
    asset_rows = [
        ("Checking & Savings", '=SUMIFS(Dim_Account[Balance],Dim_Account[AccountType],"Checking")+SUMIFS(Dim_Account[Balance],Dim_Account[AccountType],"Savings")'),
        ("Investments (Brokerage/Retirement)", "=SUM(Fact_Investments[MarketValue])"),
        ("Real Estate (manual entry)", 350000),
        ("Vehicles (manual entry)", 18000),
        ("Other Assets (manual entry)", 2000),
    ]
    for i, (label, val) in enumerate(asset_rows):
        rr = r + 1 + i
        ws.cell(row=rr, column=1, value=label).font = f(10)
        c = ws.cell(row=rr, column=2, value=val)
        c.number_format = CUR_FMT
        c.font = f(10, color=BLUE_ACCENT if isinstance(val, (int, float)) else DARK_TEXT)
    ar0, ar1 = r + 1, r + len(asset_rows)
    tot_assets_row = r + len(asset_rows) + 1
    ws.cell(row=tot_assets_row, column=1, value="TOTAL ASSETS").font = f(10, bold=True)
    ws.cell(row=tot_assets_row, column=2, value=f"=SUM(B{ar0}:B{ar1})").font = f(10, bold=True)
    ws.cell(row=tot_assets_row, column=2).number_format = CUR_FMT
    ws.cell(row=tot_assets_row, column=1).border = BORDER_BOTTOM
    ws.cell(row=tot_assets_row, column=2).border = BORDER_BOTTOM

    r2 = tot_assets_row + 3
    r2 = style_section_header(ws, r2, 1, "Liabilities", span=3)
    liab_rows = [
        ("Mortgage", '=SUMIFS(Dim_Debt[CurrentBalance],Dim_Debt[DebtType],"Mortgage")'),
        ("Loans (Auto/Student/Personal)", '=SUMIFS(Dim_Debt[CurrentBalance],Dim_Debt[DebtType],"Car Loan")+SUMIFS(Dim_Debt[CurrentBalance],Dim_Debt[DebtType],"Student Loan")+SUMIFS(Dim_Debt[CurrentBalance],Dim_Debt[DebtType],"Personal Loan")'),
        ("Credit Cards", '=SUMIFS(Dim_Debt[CurrentBalance],Dim_Debt[DebtType],"Credit Card")'),
    ]
    for i, (label, val) in enumerate(liab_rows):
        rr = r2 + 1 + i
        ws.cell(row=rr, column=1, value=label).font = f(10)
        c = ws.cell(row=rr, column=2, value=val)
        c.number_format = CUR_FMT
    lr0, lr1 = r2 + 1, r2 + len(liab_rows)
    tot_liab_row = r2 + len(liab_rows) + 1
    ws.cell(row=tot_liab_row, column=1, value="TOTAL LIABILITIES").font = f(10, bold=True)
    ws.cell(row=tot_liab_row, column=2, value=f"=SUM(B{lr0}:B{lr1})").font = f(10, bold=True)
    ws.cell(row=tot_liab_row, column=2).number_format = CUR_FMT
    ws.cell(row=tot_liab_row, column=1).border = BORDER_BOTTOM
    ws.cell(row=tot_liab_row, column=2).border = BORDER_BOTTOM

    r3 = tot_liab_row + 2
    ws.cell(row=r3, column=1, value="NET WORTH").font = f(13, bold=True, color=WHITE)
    ws.cell(row=r3, column=1).fill = fill(NAVY)
    nw_cell = ws.cell(row=r3, column=2, value=f"=B{tot_assets_row}-B{tot_liab_row}")
    nw_cell.font = f(13, bold=True, color=WHITE)
    nw_cell.fill = fill(NAVY)
    nw_cell.number_format = CUR_FMT

    wb.defined_names["NetWorth_TotalAssetsCell"] = DefinedName("NetWorth_TotalAssetsCell", attr_text=f"'Net Worth'!$B${tot_assets_row}")
    wb.defined_names["NetWorth_TotalLiabCell"] = DefinedName("NetWorth_TotalLiabCell", attr_text=f"'Net Worth'!$B${tot_liab_row}")
    wb.defined_names["NetWorthCell"] = DefinedName("NetWorthCell", attr_text=f"'Net Worth'!$B${r3}")

    # History table for the 13 trailing months (chart source) - values approximate historical trend using linear ramp to current
    r4 = r3 + 3
    r4 = style_section_header(ws, r4, 1, "Net Worth History (for trend chart)  —  update monthly, or overwrite with real historical snapshots", span=4)
    ws.cell(row=r4, column=1, value="Month").font = f(10, bold=True)
    ws.cell(row=r4, column=2, value="Net Worth").font = f(10, bold=True)
    ws.cell(row=r4, column=3, value="Monthly Change").font = f(10, bold=True)
    import seed_data as sd
    import datetime as dt
    hist_start = r4 + 1
    n_hist = 13
    for i in range(n_hist):
        months_ago = n_hist - 1 - i
        rr = hist_start + i
        ws.cell(row=rr, column=1, value=f"=EDATE(TodayDate,{-months_ago})")
        ws.cell(row=rr, column=1).number_format = "mmm-yy"
        if i == n_hist - 1:
            ws.cell(row=rr, column=2, value=f"=B{r3}")
        else:
            # simple back-cast: assume ~1.6% total net-worth growth per month leading up to today
            ws.cell(row=rr, column=2, value=f"=$B${r3}/(1.016^{months_ago})")
        ws.cell(row=rr, column=2).number_format = CUR_FMT
        if i == 0:
            ws.cell(row=rr, column=3, value=0)
        else:
            ws.cell(row=rr, column=3, value=f"=B{rr}-B{rr-1}")
        ws.cell(row=rr, column=3).number_format = CUR_FMT
    hist_end = hist_start + n_hist - 1
    ws.cell(row=hist_end + 2, column=1,
            value="Back-cast assumption: prior months are modeled at 1.6%/mo growth toward today's live Net Worth; replace with actual monthly snapshots for a precise history.").font = f(8, italic=True, color="6B7280")
    ws.merge_cells(start_row=hist_end+2, start_column=1, end_row=hist_end+2, end_column=6)

    ws.cell(row=hist_end + 4, column=1, value="Annual Change").font = f(10, bold=True)
    ws.cell(row=hist_end + 4, column=2, value=f"=B{hist_end}-B{hist_start}")
    ws.cell(row=hist_end + 4, column=2).number_format = CUR_FMT

    growth_row = hist_end + 5
    ws.cell(row=growth_row, column=1, value="Latest Monthly Growth %").font = f(10, bold=True)
    ws.cell(row=growth_row, column=2, value=f"=IFERROR(C{hist_end}/B{hist_end-1},0)")
    ws.cell(row=growth_row, column=2).number_format = PCT_FMT

    wb.defined_names["NetWorthHistStart"] = DefinedName("NetWorthHistStart", attr_text=f"'Net Worth'!$A${hist_start}")
    wb.defined_names["NetWorthHistEnd"] = DefinedName("NetWorthHistEnd", attr_text=f"'Net Worth'!$A${hist_end}")
    wb.defined_names["NetWorthGrowthPct"] = DefinedName("NetWorthGrowthPct", attr_text=f"'Net Worth'!$B${growth_row}")

    ws.conditional_formatting.add(f"B{r3}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"B{r3}", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))

    ws.sheet_properties.tabColor = NAVY
    return {"tot_assets_row": tot_assets_row, "tot_liab_row": tot_liab_row, "nw_row": r3,
            "hist_start": hist_start, "hist_end": hist_end}
