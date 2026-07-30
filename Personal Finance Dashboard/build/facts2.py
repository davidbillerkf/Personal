"""Savings sheet (simplified planner scope)."""
from openpyxl.workbook.defined_name import DefinedName
from common import *
import seed_data as sd


def build_savings(wb):
    ws = wb.create_sheet("Savings")
    style_sheet_title(ws, "Savings", span_cols=11)
    ws.sheet_view.showGridLines = False

    kpi_specs = [
        ("Total Saved (All Goals)", "=SUM(Dim_Goal[CurrentAmount])", CUR_FMT0),
        ("Monthly Savings (This Month)", '=SUMIFS(Fact_SavingsTransfers[Amount],Fact_SavingsTransfers[Date],">="&CurMonthStart,Fact_SavingsTransfers[Date],"<="&CurMonthEnd)', CUR_FMT0),
        ("Savings Rate", '=IFERROR(SUMIFS(Fact_SavingsTransfers[Amount],Fact_SavingsTransfers[Date],">="&CurMonthStart,Fact_SavingsTransfers[Date],"<="&CurMonthEnd)/SUMIFS(Fact_Income[NetAmount],Fact_Income[Date],">="&CurMonthStart,Fact_Income[Date],"<="&CurMonthEnd),0)', PCT_FMT),
    ]
    for i, (label, formula, fmt) in enumerate(kpi_specs):
        kpi_card(ws, 4, 1 + i * 3, label, formula, number_format=fmt, width_cols=3)
    for col in range(1, 12):
        ws.column_dimensions[get_column_letter(col)].width = 15

    r = 9
    r = style_section_header(ws, r, 1, "Savings Goals  —  set a target; log transfers below to track progress automatically", span=10)
    headers = ["GoalID", "GoalName", "TargetAmount", "TargetDate", "Priority", "CurrentAmount",
               "Remaining", "Completion %", "Avg Monthly Contribution", "Projected Completion Date"]
    data_rows = [list(row) + [None, None, None, None, None] for row in sd.GOALS]
    goal_info = add_table(ws, r, 1, headers, "Dim_Goal", data_rows,
                           col_widths=[9, 20, 15, 13, 11, 15, 14, 13, 20, 20])
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
    add_list_validation(ws, f"E{fr}:E{lr+50}", '"High,Medium,Low"')

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
