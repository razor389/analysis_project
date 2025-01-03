import os
import sys
import json
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from datetime import datetime
import textwrap
import re

# Define Custom Fills
label_fill = PatternFill(start_color="00FFFF", end_color="00FFFF", fill_type="solid")  # Light blue
data_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")   # Cornsilk

# Define custom fonts
title_font = Font(name="Times New Roman", size=14, bold=True, italic=True)
label_font = Font(name="Times New Roman", size=10, bold=True, italic=True)
data_tnr_font = Font(name = "Times New Roman", size=10)
data_tnr_italic_font = Font(name='Times New Roman', size=10, italic=True)
data_tnr_bold_font = Font(name = "Times New Roman", bold=True)
data_arial_font = Font(name = "Arial", size=10)
data_arial_bold_font = Font(name ="Arial", size=10, bold=True)
data_arial_italic_font = Font(name = "Arial", size=10, italic=True)

center_alignment = Alignment(horizontal="center", vertical="center")
right_alignment = Alignment(horizontal="right", vertical="center")
# Define a thin black border
thin_border = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='thin', color='000000')
)

# Define a thick black border for the outer edges
thick_border = Border(
    left=Side(style='thick', color='000000'),
    right=Side(style='thick', color='000000'),
    top=Side(style='thick', color='000000'),
    bottom=Side(style='thick', color='000000')
)

def apply_table_border(ws, row, start_col, end_col):
    """
    Applies a thin border around a group of cells in a specified row from start_col to end_col.

    :param ws: The worksheet object.
    :param row: The row number where the group is located.
    :param start_col: The starting column number of the group.
    :param end_col: The ending column number of the group.
    """
    # Define a thin black border
    thin_side = Side(style='thin', color='000000')
    
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        existing_border = cell.border
        
        # Initialize new border components with existing borders
        new_border = Border(
            top=existing_border.top,
            bottom=existing_border.bottom,
            left=existing_border.left,
            right=existing_border.right,
            diagonal=existing_border.diagonal,
            diagonal_direction=existing_border.diagonal_direction,
            outline=existing_border.outline,
            vertical=existing_border.vertical,
            horizontal=existing_border.horizontal
        )
        
        # Apply top and bottom borders to all cells in the group
        new_border.top = thin_side
        new_border.bottom = thin_side
        
        # Apply left border to the first cell and right border to the last cell
        if col == start_col:
            new_border.left = thin_side
        if col == end_col:
            new_border.right = thin_side
        
        # Assign the updated border back to the cell
        cell.border = new_border

def title_fill_range(ws, row_number, left_col, right_col):
        for cc in range(left_col, right_col + 1):
            cell = ws.cell(row_number, cc)
            if cell.fill.patternType is None:
                cell.fill = label_fill

def format_workbook(writer):
    """
    removes gridlines from all worksheets.
    """
    for sheetname in writer.book.sheetnames:
        ws = writer.book[sheetname]
        # Remove gridlines
        ws.sheet_view.showGridLines = False
        
def load_final_output(ticker):
    file_path = os.path.join("output", f"{ticker}_yoy_consolidated.json")
    if not os.path.exists(file_path):
        print(f"Error: {file_path} does not exist. Please generate {ticker}_yoy_consolidated.json first.")
        sys.exit(1)
    with open(file_path, "r") as f:
        data = json.load(f)
    return data

def create_xls(xls_filename):
    if os.path.exists(xls_filename):
        print(f"Overwriting existing file: {xls_filename}")
    return pd.ExcelWriter(xls_filename, engine='openpyxl', mode='w')

def write_summary_sheet(writer, final_output):
    wb = writer.book

    if 'Summary' not in wb.sheetnames:
        wb.create_sheet('Summary')
    ws = wb['Summary']

    summary_data = final_output["summary"]
    company_name = summary_data["company_name"]
    exchange = summary_data["exchange"]
    symbol = summary_data["symbol"]
    description = summary_data["description"]

    # Write and format the combined title in E1
    combined_title = f"{company_name.upper()} ({exchange}) - {symbol}"
    ws['E1'] = combined_title
    ws['E1'].font = title_font

    # Write "Company Description" in A4 with label formatting
    ws['A4'] = "Company Description"
    ws['A4'].font = label_font

    # Wrap and write the description
    wrapped_lines = textwrap.wrap(description, width=150)
    start_row = 5
    col = 2  # Column B
    for i, line in enumerate(wrapped_lines):
        cell = ws.cell(row=start_row + i, column=col, value=line)
        # Optionally, you can set font here if needed

def write_company_description(writer, final_output):
    reported_currency = final_output["summary"]["reported_currency"]
    cd_info = final_output["company_description"]
    cd_data = cd_info["data"]
    wb = writer.book

    if "Co. Desc" not in wb.sheetnames:
        wb.create_sheet("Co. Desc")
    ws = wb["Co. Desc"]
    ws.freeze_panes = "B1"

    # Write and format labels
    ws.cell(row=1, column=3, value="Currency").fill = label_fill
    ws.cell(row=1, column=3).font = label_font
    ws.cell(row=1, column=3).border = thin_border

    ws.cell(row=1, column=5, value="FY End").fill = label_fill
    ws.cell(row=1, column=5).font = label_font
    ws.cell(row=1, column=5).border = thin_border

    ws.cell(row=1, column=7, value="Stock Price").fill = label_fill
    ws.cell(row=1, column=7).font = label_font
    ws.cell(row=1, column=7).border = thin_border

    ws.cell(row=1, column=9, value="Market Cap").fill = label_fill
    ws.cell(row=1, column=9).font = label_font
    ws.cell(row=1, column=9).border = thin_border

    # Write and format data cells
    rc_cell = ws.cell(row=2, column=3, value=reported_currency)
    rc_cell.fill = data_fill
    rc_cell.font = label_font
    rc_cell.border = thin_border

    fiscal_year_end = cd_info.get("fiscal_year_end")
    fy_cell = ws.cell(row=2, column=5, value=fiscal_year_end)
    fy_cell.fill = data_fill
    fy_cell.font = label_font
    fy_cell.border = thin_border

    stock_price = to_float(cd_info.get("stock_price"))
    sp_cell = ws.cell(row=2, column=7, value=stock_price)
    sp_cell.fill = data_fill
    sp_cell.font = label_font
    sp_cell.number_format = '$#,##0.00'
    sp_cell.border = thin_border

    market_cap = to_float(cd_info.get("marketCapitalization")) / 1_000_000
    mc_cell = ws.cell(row=2, column=9, value=market_cap)
    mc_cell.fill = data_fill
    mc_cell.font = label_font
    mc_cell.number_format = '$#,##0'
    mc_cell.border = thin_border

    # Sort years for consistent column ordering
    sorted_years = sorted(cd_data.keys(), key=lambda x: int(x))

    # Determine the next two years
    if sorted_years:
        max_year = max(int(year) for year in sorted_years)
        new_years = [str(max_year + 1), str(max_year + 2)]
    else:
        new_years = ["2024", "2025"]

    # Append the new years to the sorted_years list
    all_years = sorted_years + new_years

    # Write the years at row=3 starting at column B
    start_col = 2  # Column B
    for i, year in enumerate(all_years):
        year_cell = ws.cell(row=3, column=start_col + i, value=year)
        year_cell.fill = label_fill
        year_cell.font = label_font
        year_cell.border = thin_border

    # Define metrics that should be displayed in millions
    million_scale_metrics = {
        "net_profit",
        "dividends_paid", 
        "shares_outstanding",
        "buyback",
        "share_equity",
        "long_term_debt"
    }

    # Define the metrics and their corresponding row positions
    metric_positions = {
        "net_profit": 4,            
        "diluted_eps": 5,           
        "operating_eps": 6,         
        "pe_ratio": 8,              
        "price_low": 9,             
        "price_high": 10,           
        "dividends_paid": 12,       
        "dividends_per_share": 13,  
        "avg_dividend_yield": 14,   
        "shares_outstanding": 16,   
        "buyback": 17,              
        "share_equity": 19,         
        "book_value_per_share": 20, 
        "long_term_debt": 22,       
        "roe": 24,                  
        "roc": 25                   
    }

    # Define human-readable labels for each metric
    metric_labels = {
        "net_profit": "Net Profit",
        "diluted_eps": "Diluted EPS",
        "operating_eps": "Operating EPS",
        "pe_ratio": "P/E Ratio",
        "price_low": "Yrly Price Low",
        "price_high": "Yrly Price High",
        "dividends_paid": "Dividends Paid",
        "dividends_per_share": "Dividends/Share",
        "avg_dividend_yield": "Avg Div Yield",
        "shares_outstanding": "Shares Outstanding",
        "buyback": "Buyback",
        "share_equity": "Share Equity",
        "book_value_per_share": "Book Value/Share",
        "long_term_debt": "Long-Term Debt",
        "roe": "ROE",
        "roc": "ROC"
    }

    # Updated number formats according to specifications
    number_formats = {
        "net_profit": '#,##0',          # Millions with commas, no $ no M
        "diluted_eps": '$#,##0.00',     # Dollars and cents
        "operating_eps": '$#,##0.00',    # Dollars and cents
        "pe_ratio": '#,##0.0',          # One decimal place
        "price_low": '#,##0.0',         # One decimal place
        "price_high": '#,##0.0',        # One decimal place
        "dividends_paid": '#,##0',      # Millions with commas, no $ no M
        "dividends_per_share": '$#,##0.00',  # Dollars and cents
        "avg_dividend_yield": '0.00%',   # Percentage with two decimal places
        "shares_outstanding": '#,##0',   # Millions with commas, no $ no M
        "buyback": '#,##0',             # Millions with commas, no $ no M
        "share_equity": '#,##0',        # Millions with commas, no $ no M
        "book_value_per_share": '$#,##0.00',  # Dollars and cents
        "long_term_debt": '#,##0',      # Millions with commas, no $ no M
        "roe": '0.0%',                  # Percentage with one decimal place
        "roc": '0.0%'                   # Percentage with one decimal place
    }

    # Write metric labels in column A with label formatting
    for metric, metric_row in metric_positions.items():
        label_cell = ws.cell(row=metric_row, column=1, value=metric_labels[metric])
        label_cell.fill = label_fill
        label_cell.font = label_font
        label_cell.border = thin_border

    # For each metric, write the data for each year in the specified rows
    for metric, metric_row in metric_positions.items():
        for i, year in enumerate(all_years):
            col = start_col + i
            col_letter = get_column_letter(col)
            data_cell = None

            if year in new_years:
                # For new years, we'll add formulas for certain metrics
                if metric == "diluted_eps":
                    formula = f"={col_letter}{metric_positions['net_profit']}/{col_letter}{metric_positions['shares_outstanding']}"
                    data_cell = ws.cell(row=metric_row, column=col, value=formula)
                
                elif metric == "operating_eps":
                    formula = f"={col_letter}{metric_positions['diluted_eps']}"
                    data_cell = ws.cell(row=metric_row, column=col, value=formula)
                
                elif metric == "pe_ratio":
                    formula = f"=(({col_letter}{metric_positions['price_low']}+{col_letter}{metric_positions['price_high']})/2)/{col_letter}{metric_positions['diluted_eps']}"
                    data_cell = ws.cell(row=metric_row, column=col, value=formula)
                
                elif metric == "dividends_per_share":
                    formula = f"={col_letter}{metric_positions['dividends_paid']}/{col_letter}{metric_positions['shares_outstanding']}"
                    data_cell = ws.cell(row=metric_row, column=col, value=formula)
                
                elif metric == "buyback":
                    prev_col = get_column_letter(col-1)
                    formula = f"=({prev_col}{metric_positions['shares_outstanding']}-{col_letter}{metric_positions['shares_outstanding']})*({col_letter}{metric_positions['price_low']}+{col_letter}{metric_positions['price_high']})/2"
                    data_cell = ws.cell(row=metric_row, column=col, value=formula)
                
                elif metric == "avg_dividend_yield":
                    formula = f"={col_letter}{metric_positions['dividends_per_share']}/((({col_letter}{metric_positions['price_low']}+{col_letter}{metric_positions['price_high']})/2))"
                    data_cell = ws.cell(row=metric_row, column=col, value=formula)
                
                elif metric == "book_value_per_share":
                    formula = f"={col_letter}{metric_positions['share_equity']}/{col_letter}{metric_positions['shares_outstanding']}"
                    data_cell = ws.cell(row=metric_row, column=col, value=formula)
                
                elif metric == "roe":
                    formula = f"={col_letter}{metric_positions['net_profit']}/{col_letter}{metric_positions['share_equity']}"
                    data_cell = ws.cell(row=metric_row, column=col, value=formula)
                
                elif metric == "roc":
                    formula = f"={col_letter}{metric_positions['net_profit']}/({col_letter}{metric_positions['share_equity']}+{col_letter}{metric_positions['long_term_debt']})"
                    data_cell = ws.cell(row=metric_row, column=col, value=formula)
                
                else:
                    value = cd_data.get(year, {}).get(metric)
                    if value is not None and metric in million_scale_metrics:
                        value = value / 1_000_000
                    data_cell = ws.cell(row=metric_row, column=col, value=value)

                data_cell.font = data_tnr_italic_font
                data_cell.alignment = right_alignment
            
            else:
                value = cd_data.get(year, {}).get(metric)
                if value is not None and metric in million_scale_metrics:
                    value = value / 1_000_000
                data_cell = ws.cell(row=metric_row, column=col, value=value)
                data_cell.font = data_tnr_font

            # Apply common formatting to all cells, whether formula or value
            data_cell.fill = data_fill
            data_cell.border = thin_border
            data_cell.number_format = number_formats[metric]

def write_analyses_sheet(writer, final_output):
    reported_currency = final_output["summary"]["reported_currency"]
    analyses = final_output["analyses"]
    inv_char = analyses["investment_characteristics"]
    data = analyses["data"]
    wb = writer.book

    if "Analyses" not in wb.sheetnames:
        wb.create_sheet("Analyses")
    ws = wb["Analyses"]
    ws.freeze_panes = "B1"

    # Write and format the "Investment Characteristics" title
    ic_cell = ws.cell(row=1, column=6, value="Investment Characteristics (in mlns "+reported_currency + ")")
    ic_cell.fill = label_fill
    ic_cell.font = title_font
    apply_table_border(ws, 1, 5, 9)
    title_fill_range(ws, 1, 5, 9)

    # Extract required fields and write labels with formatting
    labels_with_positions = [
        (3, 3, "Earnings Analysis:"),
        (5, 4, "Growth Rate %:"),
        (7, 4, "Quality %:"),
        (3, 8, "Use Of Earnings Analysis:"),
        (5, 9, "Avg Div Payout Rate:"),
        (7, 9, "Avg Stk Buyback Rate:"),
        (13, 3, "Sales Analysis:"),
        (15, 4, "Growth Rate %:"),
        (17, 4, "Growth Rate PS %:"),
        (13, 8, "Sales Analysis (last 5 yrs.):"),
        (15, 9, "Growth Rate %:"),
        (17, 9, "Growth Rate PS %:")
    ]

    for row, col, text in labels_with_positions:
        cell = ws.cell(row=row, column=col, value=text)
        cell.fill = label_fill
        cell.font = label_font
        apply_table_border(ws, row, col, col+1)
        title_fill_range(ws, row, col, col+1)

    # Fetch data for the investment characteristics
    growth_rate_operating_eps = inv_char["earnings_analysis"].get("growth_rate_percent_operating_eps")
    quality_percent = inv_char["earnings_analysis"].get("quality_percent")

    avg_div_payout = inv_char["use_of_earnings_analysis"].get("avg_dividend_payout_percent")
    avg_stk_buyback = inv_char["use_of_earnings_analysis"].get("avg_stock_buyback_percent")

    growth_rate_rev = inv_char["sales_analysis"].get("growth_rate_percent_revenues")
    growth_rate_sps = inv_char["sales_analysis"].get("growth_rate_percent_sales_per_share")

    growth_rate_rev_5y = inv_char["sales_analysis_last_5_years"].get("growth_rate_percent_revenues")
    growth_rate_sps_5y = inv_char["sales_analysis_last_5_years"].get("growth_rate_percent_sales_per_share")

    # Write the investment characteristics data (data cells)
    data_cells = {
        (5, 6): growth_rate_operating_eps,    # F5
        (7, 6): quality_percent,              # F7
        (5, 11): avg_div_payout,              # K5
        (7, 11): avg_stk_buyback,             # K7
        (15, 6): growth_rate_rev,             # F15
        (17, 6): growth_rate_sps,             # F17
        (15, 11): growth_rate_rev_5y,         # K15
        (17, 11): growth_rate_sps_5y          # K17
    }

    for (row, col), value in data_cells.items():
        cell = ws.cell(row=row, column=col, value=value)
        cell.fill = data_fill
        cell.font = data_tnr_bold_font
        cell.border = thin_border
        # Apply percentage format for these values
        cell.number_format = '0.0%'

    # Handle the data by years
    sorted_years = sorted(data.keys(), key=lambda x: int(x))
    
    # Determine the next two years
    if sorted_years:
        max_year = max(int(year) for year in sorted_years)
        new_years = [str(max_year + 1), str(max_year + 2)]
    else:
        new_years = ["2024", "2025"]

    # Append the new years to the sorted_years list
    all_years = sorted_years + new_years
    
    start_col = 2  # Column B

    # Write years at row 9 and row 19 (year labels)
    for i, year in enumerate(all_years):
        y9_cell = ws.cell(row=9, column=start_col + i, value=year)
        y9_cell.fill = label_fill
        y9_cell.font = label_font
        y9_cell.border = thin_border

        y19_cell = ws.cell(row=19, column=start_col + i, value=year)
        y19_cell.fill = label_fill
        y19_cell.font = label_font
        y19_cell.border = thin_border

    # Define the metrics and their row positions
    metric_rows_1 = {
        "revenues": 10,
        "sales_per_share": 11
    }

    metric_rows_2 = {
        "op_margin_percent": 20,
        "tax_rate": 21,
        "depreciation": 22,
        "depreciation_percent": 23
    }

    # Add labels in column A for these metrics with label formatting
    additional_labels = {
        10: "Revenues",
        11: "Sales/Share",
        20: "Operating Margin %",
        21: "Tax Rate %",
        22: "Depreciation",
        23: "Depreciation %"
    }

    # Define the formatting rules for each metric
    number_formats = {
        "revenues": '#,##0',  # Millions with commas
        "sales_per_share": '$#,##0.00',  # Dollars and cents
        "op_margin_percent": '0.0%',  # Percentage with one decimal place
        "tax_rate": '0.0%',  # Percentage with one decimal place
        "depreciation": '#,##0',  # Millions with commas
        "depreciation_percent": '0.0%'  # Percentage with one decimal place
    }

    # Define metrics that should be displayed in millions
    million_scale_metrics = {
        "revenues",
        "depreciation"
    }

    for row, label in additional_labels.items():
        cell = ws.cell(row=row, column=1, value=label)
        cell.fill = label_fill
        cell.font = label_font
        cell.border = thin_border

    # Write first set of metrics (data cells)
    for metric, row_num in metric_rows_1.items():
        for i, year in enumerate(all_years):
            col = start_col + i
            col_letter = get_column_letter(col)
            
            if year in new_years:
                if metric == "sales_per_share":
                    # revenue this year / shares outstanding from Co. Desc sheet
                    formula = f"={col_letter}{metric_rows_1['revenues']}/('Co. Desc'!{col_letter}16)"
                    cell = ws.cell(row=row_num, column=col, value=formula)
                else:
                    val = data.get(year, {}).get(metric)
                    if val is not None and metric in million_scale_metrics:
                        val = val / 1_000_000
                    cell = ws.cell(row=row_num, column=col, value=val)
            else:
                val = data.get(year, {}).get(metric)
                if val is not None and metric in million_scale_metrics:
                    val = val / 1_000_000
                cell = ws.cell(row=row_num, column=col, value=val)
            
            cell.fill = data_fill
            cell.font = data_tnr_italic_font if year in new_years else data_tnr_font
            cell.border = thin_border
            cell.number_format = number_formats[metric]

    # Write second set of metrics (data cells)
    for metric, row_num in metric_rows_2.items():
        for i, year in enumerate(all_years):
            col = start_col + i
            col_letter = get_column_letter(col)
            
            if year in new_years:
                if metric == "depreciation_percent":
                    # depreciation this year / net profit from Co. Desc sheet
                    formula = f"={col_letter}{metric_rows_2['depreciation']}/('Co. Desc'!{col_letter}4)"
                    cell = ws.cell(row=row_num, column=col, value=formula)
                else:
                    val = data.get(year, {}).get(metric)
                    if val is not None and metric in million_scale_metrics:
                        val = val / 1_000_000
                    cell = ws.cell(row=row_num, column=col, value=val)
            else:
                val = data.get(year, {}).get(metric)
                if val is not None and metric in million_scale_metrics:
                    val = val / 1_000_000
                cell = ws.cell(row=row_num, column=col, value=val)
            
            cell.fill = data_fill
            cell.font = data_tnr_italic_font if year in new_years else data_tnr_font
            cell.border = thin_border
            cell.number_format = number_formats[metric]

def write_profit_desc_sheet(writer, final_output):
    reported_currency = final_output["summary"]["reported_currency"]
    pd_info = final_output["profit_description"]
    pchar = pd_info["profit_description_characteristics"]
    pdata = pd_info["data"]
    wb = writer.book

    # Create or access the "Profit.Desc." sheet
    if "Profit.Desc." not in wb.sheetnames:
        wb.create_sheet("Profit.Desc.")
    ws = wb["Profit.Desc."]
    ws.freeze_panes = "D1"

    # Write and format the title
    title_cell = ws.cell(row=1, column=4, value=f"Description & Analysis of Profitability (in mlns {reported_currency})")
    title_cell.fill = label_fill
    title_cell.font = title_font
    title_fill_range(ws, 1, 3, 10)
    apply_table_border(ws, 1, 3, 10)

    # Define the order of metrics to display
    metrics_order = [
        "revenues",
        "expenses",
        "ebitda",
        "amortization_depreciation",
        "free_cash_flow",
        "capex",
        "operating_earnings",
        "external_costs",
        "earnings",
        "earnings_percent_revenue",
        "dividend_paid",
        "dividend_paid_pct_fcf",
        "share_buybacks_from_stmt_cf",
        "net_biz_acquisition"
    ]

    # Define metric labels mapping
    metric_labels = {
        "revenues": "Net Revenues:",
        "expenses": "Internal Costs:",
        "ebitda": "EBITDA:",
        "free_cash_flow": "Free Cash Flow:",
        "operating_earnings": "Operating Margin:",
        "external_costs": "External Costs:",
        "dividend_paid": "Dividend Paid:",
        "dividend_paid_pct_fcf": "Dividend Paid % of FCF:",
        "share_buybacks_from_stmt_cf": "Share Buybacks (Stmt of CFs):",
        "net_biz_acquisition": "Net Biz Acquisitions:",
        "amortization_depreciation": "Amortization & Depreciation:",
        "capex": "Capital Expenditures:",
        "earnings": "Earnings:",
        "earnings_percent_revenue": "Earnings % of Revenue:"
    }

    percent_metrics = ["dividend_paid_pct_fcf", "earnings_percent_revenue"]

    current_row = 5  # Starting row for metrics
    metric_rows = {}  # To track the row number for each metric
    breakdown_rows = {}  # To track the row numbers for each breakdown item

    # Step 1: Collect all unique revenue breakdown items across all years
    all_revenue_breakdowns = set()
    all_external_costs_breakdowns = set()
    for year_data in pdata.values():
        revenues = year_data.get("revenues", {})
        external_costs = year_data.get("external_costs", {})
        revenue_breakdown = revenues.get("breakdown", {})
        external_costs_breakdown = external_costs.get("breakdown", {})
        if revenue_breakdown:
            all_revenue_breakdowns.update(revenue_breakdown.keys())
        if external_costs_breakdown:
            all_external_costs_breakdowns.update(external_costs_breakdown.keys())

    # Step 2: Iterate through metrics and define rows
    for metric in metrics_order:
        # Get the label from the mapping
        label = metric_labels.get(metric, metric.capitalize() + ":")

        # Write the metric label cell with formatting
        label_cell = ws.cell(row=current_row, column=1, value=label)
        label_cell.fill = label_fill
        label_cell.font = label_font
        # Fill columns A to C for the label
        for col in range(1, 4):
            cell = ws.cell(row=current_row, column=col)
            cell.fill = label_fill
            cell.font = label_font
        metric_rows[metric] = current_row
        apply_table_border(ws, current_row, 1, 3)
        current_row += 1

        # Handle breakdowns
        if metric == "revenues":
            for bkey in sorted(all_revenue_breakdowns):
                breakdown_cell = ws.cell(row=current_row, column=2, value=bkey)
                breakdown_cell.font = Font(italic=True)
                breakdown_rows[(metric, bkey)] = current_row
                current_row += 1
        elif metric == "expenses" or metric == "external_costs":
            metric_data = pdata[next(iter(pdata))].get(metric, {})
            if isinstance(metric_data, dict) and "breakdown" in metric_data:
                breakdown_items = all_external_costs_breakdowns if metric == "external_costs" else metric_data["breakdown"].keys()
                for bkey in sorted(breakdown_items):
                    breakdown_cell = ws.cell(row=current_row, column=2, value=bkey)
                    breakdown_cell.font = Font(italic=True)
                    breakdown_rows[(metric, bkey)] = current_row
                    current_row += 1

    # Step 3: Write Year Headers
    sorted_years = sorted(pdata.keys(), key=lambda x: int(x))
    start_col_for_years = 4  # Starting at column D

    for i, year in enumerate(sorted_years):
        year_col = start_col_for_years + i * 2  # Each year takes 2 columns
        y_cell = ws.cell(row=3, column=year_col, value=year)
        y_cell.fill = label_fill
        y_cell.font = label_font
        y_cell.border = thin_border

    # Step 4: Write Metric Values and Breakdown Values
    for i, year in enumerate(sorted_years):
        year_col = start_col_for_years + i * 2  # Column for the value
        year_data = pdata[year]
        
        for metric in metrics_order:
            metric_val = year_data.get(metric)
            metric_row = metric_rows.get(metric)

            if metric_val is None:
                continue

            if isinstance(metric_val, dict) and "breakdown" in metric_val:
                # Handle metrics with breakdowns
                total_key = None
                if metric == "revenues":
                    total_key = "total_revenues"
                elif metric == "expenses":
                    total_key = "total_expenses"
                elif metric == "external_costs":
                    total_key = "total_external_costs"

                if total_key and total_key in metric_val:
                    val = to_float(metric_val[total_key])
                    if val is not None:
                        val = val / 1_000_000
                    cell = ws.cell(row=metric_row, column=year_col, value=val)
                    cell.fill = data_fill
                    cell.font = Font(name="Arial", italic=True)
                    cell.number_format = '#,##0'
                    cell.border = thin_border

                # Write breakdown items
                for (m, bkey), brow in breakdown_rows.items():
                    if m == metric:
                        breakdown_val = metric_val["breakdown"].get(bkey)
                        if breakdown_val is not None:
                            breakdown_val = to_float(breakdown_val)
                            if breakdown_val is not None:
                                breakdown_val = breakdown_val / 1_000_000
                            bdata_cell = ws.cell(row=brow, column=year_col, value=breakdown_val)
                            # Apply italic font
                            bdata_cell.font = data_arial_italic_font
                            # Apply number format
                            bdata_cell.number_format = '#,##0'
                            
                            # Add CAGR for revenue breakdowns
                            if metric == "revenues":
                                cagr_key = f"cagr_revenues_{bkey}_percent"
                                cagr_value = pchar.get("cagr_revenues_breakdown_percent", {}).get(cagr_key)
                                if cagr_value is not None:
                                    cagr_cell = ws.cell(row=brow, column=3, value=cagr_value)
                                    cagr_cell.font = Font(name="Arial", italic=True, size=8)
                                    cagr_cell.number_format = '0.0%'
                            
                            # Add CAGR for external costs breakdowns
                            elif metric == "external_costs":
                                cagr_key = f"cagr_external_costs_{bkey}_percent"
                                cagr_value = pchar.get("cagr_external_costs_breakdown_percent", {}).get(cagr_key)
                                if cagr_value is not None:
                                    cagr_cell = ws.cell(row=brow, column=3, value=cagr_value)
                                    cagr_cell.font = Font(name="Arial", italic=True, size=8)
                                    cagr_cell.number_format = '0.0%'
                        else:
                            bdata_cell = ws.cell(row=brow, column=year_col, value=None)
                            bdata_cell.number_format = '#,##0'
            else:
                # Handle metrics without breakdowns
                val = to_float(metric_val)
                if val is not None and metric not in percent_metrics:
                    val = val / 1_000_000
                cell = ws.cell(row=metric_row, column=year_col, value=val)
                cell.fill = data_fill
                cell.font = Font(name="Arial", italic=True)
                cell.border = thin_border
                if metric not in percent_metrics:
                    cell.number_format = '#,##0'
                else:
                    cell.number_format = '0.0%'

    # Step 5: Write CAGR Values
    cagr_map = {
        "revenues": "cagr_revenues_percent",
        "expenses": "cagr_total_expenses_percent",
        "ebitda": "cagr_ebitda_percent",
        "free_cash_flow": "cagr_free_cash_flow_percent",
        "operating_earnings": "cagr_operating_earnings_percent",
        "external_costs": "cagr_total_external_costs_percent",
        "earnings": "cagr_earnings_percent"
    }

    for metric, cagr_key in cagr_map.items():
        cagr_value = pchar.get(cagr_key)
        if cagr_value is not None and metric in metric_rows:
            row = metric_rows[metric]
            cagr_cell = ws.cell(row=row, column=3, value=cagr_value)
            cagr_cell.font = Font(name="Arial", italic=True, size=8)
            cagr_cell.number_format = '0.00%'

    breakdown_cagr_map = {
        "cost_of_revenue": "cagr_cost_of_revenue_percent",
        "research_and_development": "cagr_research_and_development_percent",
        "selling_marketing_general_admin": "cagr_selling_marketing_general_admin_percent"
    }

    for (metric, bkey), brow in breakdown_rows.items():
        cagr_field = breakdown_cagr_map.get(bkey)
        if cagr_field:
            cagr_value = pchar.get(cagr_field)
            if cagr_value is not None:
                cagr_cell = ws.cell(row=brow, column=3, value=cagr_value)
                cagr_cell.font = Font(name="Arial", italic=True, size=8)
                cagr_cell.number_format = '0.00%'

    # Step 6: Compute and Write Percentages
    revenues_row = metric_rows.get("revenues")
    expense_breakdowns = ["cost_of_revenue", "research_and_development", "selling_marketing_general_admin"]
    top_metrics = ["ebitda", "operating_earnings", "earnings"]

    for i, year in enumerate(sorted_years):
        year_col = start_col_for_years + i * 2
        rev_val = ws.cell(row=revenues_row, column=year_col).value
        if rev_val is None or rev_val == 0:
            continue

        # Calculate percentages for all revenue breakdowns
        for bkey in all_revenue_breakdowns:
            br_key = ("revenues", bkey)
            if br_key in breakdown_rows:
                brow = breakdown_rows[br_key]
                metric_val = ws.cell(row=brow, column=year_col).value
                if isinstance(metric_val, (int, float)) and metric_val is not None and rev_val != 0:
                    percent = (metric_val / rev_val) * 100
                    percent_cell = ws.cell(row=brow, column=year_col + 1, value=f"{percent:.1f}%")
                    percent_cell.font = Font(name="Arial", italic=True, size=8)

        # Calculate percentages for expense breakdowns
        for bkey in expense_breakdowns:
            br_key = ("expenses", bkey)
            if br_key in breakdown_rows:
                brow = breakdown_rows[br_key]
                metric_val = ws.cell(row=brow, column=year_col).value
                if isinstance(metric_val, (int, float)) and metric_val is not None and rev_val != 0:
                    percent = (metric_val / rev_val) * 100
                    percent_cell = ws.cell(row=brow, column=year_col + 1, value=f"{percent:.1f}%")
                    percent_cell.font = Font(name="Arial", italic=True, size=8)

        # Calculate percentages for external costs breakdowns
        for bkey in all_external_costs_breakdowns:
            br_key = ("external_costs", bkey)
            if br_key in breakdown_rows:
                brow = breakdown_rows[br_key]
                metric_val = ws.cell(row=brow, column=year_col).value
                if isinstance(metric_val, (int, float)) and metric_val is not None and rev_val != 0:
                    percent = (metric_val / rev_val) * 100
                    percent_cell = ws.cell(row=brow, column=year_col + 1, value=f"{percent:.1f}%")
                    percent_cell.font = Font(name="Arial", italic=True, size=8)

        # Calculate percentages for top metrics
        for tm in top_metrics:
            if tm in metric_rows:
                tm_row = metric_rows[tm]
                metric_val = ws.cell(row=tm_row, column=year_col).value
                if isinstance(metric_val, (int, float)) and metric_val is not None and rev_val != 0:
                    percent = (metric_val / rev_val) * 100
                    percent_cell = ws.cell(row=tm_row, column=year_col + 1, value=f"{percent:.1f}%")
                    percent_cell.font = Font(name="Arial", italic=True, size=8)

def write_balance_sheet_sheet(writer, final_output):
    reported_currency = final_output["summary"]["reported_currency"]
    bs_info = final_output["balance_sheet"]
    bs_char = bs_info["balance_sheet_characteristics"]
    bs_data = bs_info["data"]
    wb = writer.book

    if "Balance Sht." not in wb.sheetnames:
        wb.create_sheet("Balance Sht.")
    ws = wb["Balance Sht."]
    ws.freeze_panes = "F1"
    # Write and format the title
    title_cell = ws.cell(row=1, column=4, value=f"Balance Sheet (in mlns {reported_currency}):")
    title_cell.fill = label_fill
    title_cell.font = title_font
    title_fill_range(ws, 1, 4,7)
    apply_table_border(ws, 1, 4, 7)

    # Extract characteristics
    cagr_assets = bs_char.get("cagr_total_assets_percent")
    cagr_liabilities = bs_char.get("cagr_total_liabilities_percent")
    cagr_equity = bs_char.get("cagr_total_shareholders_equity_percent")

    top_sections = [
        ("assets", "Assets:"),
        ("liabilities", "Liabilities:"),
        ("shareholders_equity", "Shareholder's Equity:")
    ]

    sorted_years = sorted(bs_data.keys(), key=lambda x: int(x))
    start_col_for_years = 6  # Column F

    # Write and format years in row 3
    for i, year in enumerate(sorted_years):
        year_col = start_col_for_years + i
        y_cell = ws.cell(row=3, column=year_col, value=year)
        y_cell.fill = label_fill
        y_cell.font = label_font
        y_cell.border = thin_border

    current_row = 5
    section_rows = {}

    for section_key, section_label in top_sections:
        # Main heading label in column A with label formatting
        sec_label_cell = ws.cell(row=current_row, column=1, value=section_label)
        sec_label_cell.fill = label_fill
        sec_label_cell.font = label_font
        title_fill_range(ws, current_row, 1,5)
        apply_table_border(ws, current_row, 1, 5)
        section_rows[section_key] = current_row

        first_year = sorted_years[0]
        section_data = bs_data[first_year].get(section_key, {})
        total_key = "total_" + section_key

        # Write total values for each year (data fill)
        for i, year in enumerate(sorted_years):
            year_col = start_col_for_years + i
            val = to_float(bs_data[year][section_key].get(total_key))
            if val is not None:
                val = val / 1_000_000
            d_cell = ws.cell(row=current_row, column=year_col, value=val)
            # This is the main heading's data row, so data fill it
            d_cell.fill = data_fill
            d_cell.font = data_arial_bold_font
            d_cell.border = thin_border
            # Apply number format
            if isinstance(val, (int, float)):
                d_cell.number_format = '#,##0'

        # Apply CAGR to column E
        if section_key == "assets":
            if cagr_assets is not None:
                ws.cell(row=section_rows["assets"], column=5, value=cagr_assets).number_format = '0.0%'
        elif section_key == "liabilities":
            if cagr_liabilities is not None:
                ws.cell(row=section_rows["liabilities"], column=5, value=cagr_liabilities).number_format = '0.0%'
        elif section_key == "shareholders_equity":
            if cagr_equity is not None:
                ws.cell(row=section_rows["shareholders_equity"], column=5, value=cagr_equity).number_format = '0.0%'

        current_row += 1

        # Breakdown items (no fill, no font)
        breakdown_data = section_data.get("breakdown", {})
        breakdown_rows = {}
        for bkey in breakdown_data.keys():
            # Breakdown label no fill
            ws.cell(row=current_row, column=2, value=bkey)
            breakdown_rows[bkey] = current_row
            current_row += 1

        # Write breakdown data (no fill)
        for bkey, brow in breakdown_rows.items():
            for i, year in enumerate(sorted_years):
                year_col = start_col_for_years + i
                val = to_float(bs_data[year][section_key]["breakdown"].get(bkey))
                if val is not None:
                    val = val / 1_000_000
                bdata_cell = ws.cell(row=brow, column=year_col, value=val)
                bdata_cell.font = data_arial_italic_font
                # Apply number format
                if isinstance(val, (int, float)):
                    bdata_cell.number_format = '#,##0'

        # Add a blank line before the next section
        current_row += 1

def write_studies_sheet(writer, final_output):
    reported_currency = final_output["summary"]["reported_currency"]
    studies = final_output.get("studies", {})
    analysis = studies.get("analysis_of_debt_levels", {})

    tdc = analysis.get("total_debt_capital", {})
    ltd = analysis.get("long_term_debt", {})
    nip = analysis.get("net_income_payback", {})
    anip = analysis.get("addback_net_inc_payback", {})

    wb = writer.book
    if "Studies" not in wb.sheetnames:
        wb.create_sheet("Studies")
    ws = wb["Studies"]

    # Write and format the title
    title_cell = ws.cell(row=1, column=4, value=f"Description & Analysis of Debt Levels (in mlns {reported_currency}):")
    title_cell.font = title_font
    apply_table_border(ws, 1 ,3, 10)
    title_fill_range(ws, 1, 3,10)

    # Summary section labels (apply label_fill and label_font)
    summary_labels = ["Summary:", "Total Debt-Capital:", "Long Term Debt-Cap.:", "Net Income Payback:", "Addback Net Inc Payback:"]
    summary_rows = [3, 7, 15, 23, 35]
    for label, row in zip(summary_labels, summary_rows):
        cell = ws.cell(row=row, column=1, value=label)
        cell.fill = label_fill
        cell.font = label_font
        cell.border = thin_border

    # Summary Section Texts (no font changes)
    ws.cell(row=3, column=2, value="Debt is a four-letter word.  Debt causes the years of repayment of capital to equity shareholders to stretch").font = data_arial_font
    ws.cell(row=4, column=2, value="out into the more distant future.  Even worse, debt can cause the best business model to become the").font = data_arial_font
    ws.cell(row=5, column=2, value="property of bondholders in a rough economic environment.").font = data_arial_font

    # Total Debt-Capital Section Texts
    ws.cell(row=7, column=2, value="The measure of total debt to total capital is useful when book value is a good measure of a firm's worth.  This").font = data_arial_font
    ws.cell(row=8, column=2, value="is particularly true of traditional businesses where property, plant and equipment are important.  Further, it").font = data_arial_font
    ws.cell(row=9, column=2, value="helps to have this ratio in capital intensive businesses with cyclical earnings.").font = data_arial_font

    # Total Debt-Capital Data
    ws.cell(row=11, column=3, value="Total Debt:").font = label_font
    val = tdc.get("total_debt")
    if val is not None:
        val = val / 1_000_000
    data_cell = ws.cell(row=11, column=4, value=val)
    data_cell.number_format = '#,##0'
    data_cell.font = data_arial_font
    ws.cell(row=11, column=5, value="Here, deferred income taxes have been excluded.").font = data_arial_font

    ws.cell(row=12, column=3, value="Total Capital:").font = label_font
    val = tdc.get("total_capital")
    if val is not None:
        val = val / 1_000_000
    data_cell = ws.cell(row=12, column=4, value=val)
    data_cell.number_format = '#,##0'
    data_cell.font = data_arial_font
    ws.cell(row=12, column=5, value="Here, deferred income taxes have been excluded.").font = data_arial_font

    ws.cell(row=13, column=3, value="Ratio:").font = label_font
    data_cell = ws.cell(row=13, column=4, value=tdc.get("total_debt_ratio"))
    data_cell.number_format = '0.0%'
    data_cell.font = data_arial_font

    # Long Term Debt-Cap. Section Texts
    ws.cell(row=15, column=2, value="The measure of long term debt to total capital is useful when total debt is distorted by the high presence").font = data_arial_font
    ws.cell(row=16, column=2, value="of current assets being financed by current liabilities.  Again, the measure works best within a traditional").font = data_arial_font
    ws.cell(row=17, column=2, value="industry setting.  The ratio helps position the equity shareholders.").font = data_arial_font

    # Long Term Debt-Cap. Data
    ws.cell(row=19, column=3, value="L. T. Debt:").font = label_font
    val = ltd.get("lt_debt")
    if val is not None:
        val = val / 1_000_000
    data_cell = ws.cell(row=19, column=4, value=val)
    data_cell.number_format = '#,##0'
    data_cell.font = data_arial_font
    ws.cell(row=19, column=5, value="Here, the current liabilities have been excluded.").font = data_arial_font

    ws.cell(row=20, column=3, value="L. T. Capital:").font = label_font
    val = ltd.get("lt_capital")
    if val is not None:
        val = val / 1_000_000
    data_cell = ws.cell(row=20, column=4, value=val)
    data_cell.number_format = '#,##0'
    data_cell.font = data_arial_font
    ws.cell(row=20, column=5, value="Here, the current liabilities have been excluded.").font = data_arial_font

    ws.cell(row=21, column=3, value="Ratio:").font = label_font
    data_cell = ws.cell(row=21, column=4, value=ltd.get("lt_debt_ratio"))
    data_cell.number_format = '0.0%'
    data_cell.font = data_arial_font

    # Net Income Payback Section Texts
    ws.cell(row=23, column=2, value="The measure of how quickly total debt is repaid by net income is a conservative measure, as it includes").font = data_arial_font
    ws.cell(row=24, column=2, value="debt such as current liabilities, that are financed by current assets and excludes some sources of cash, such").font = data_arial_font
    ws.cell(row=25, column=2, value="as noncash amortization numbers.").font = data_arial_font

    # Net Income Payback Data
    ws.cell(row=27, column=3, value="Total Debt:").font = label_font
    val = nip.get("total_debt")
    if val is not None:
        val = val / 1_000_000
    data_cell = ws.cell(row=27, column=4, value=val)
    data_cell.number_format = '#,##0'
    data_cell.font = data_arial_font

    ws.cell(row=28, column=3, value="Net Income:").font = label_font
    val = nip.get("net_income")
    if val is not None:
        val = val / 1_000_000
    data_cell = ws.cell(row=28, column=4, value=val)
    data_cell.number_format = '#,##0'
    data_cell.font = data_arial_font

    ws.cell(row=29, column=3, value="Years Payback:").font = label_font
    data_cell = ws.cell(row=29, column=4, value=nip.get("years_payback_total_debt"))
    data_cell.number_format = '#,##0.0'
    data_cell.font = data_arial_font

    ws.cell(row=31, column=3, value="L.T. Debt:").font = label_font
    val = nip.get("lt_debt")
    if val is not None:
        val = val / 1_000_000
    data_cell = ws.cell(row=31, column=4, value=val)
    data_cell.number_format = '#,##0'
    data_cell.font = data_arial_font

    ws.cell(row=32, column=3, value="Net Income:").font = label_font
    val = nip.get("net_income")
    if val is not None:
        val = val / 1_000_000
    data_cell = ws.cell(row=32, column=4, value=val)
    data_cell.number_format = '#,##0'
    data_cell.font = data_arial_font

    ws.cell(row=33, column=3, value="Years Payback:").font = label_font
    data_cell = ws.cell(row=33, column=4, value=nip.get("years_payback_lt_debt"))
    data_cell.number_format = '#,##0.0'
    data_cell.font = data_arial_font

    # Addback Net Inc Payback Section Texts
    ws.cell(row=35, column=2, value="The measure of how quickly debt is repaid by addback net income is a good measure, as it starts with GAAP").font = data_arial_font
    ws.cell(row=36, column=2, value="net income and adds back expenses on an after-tax basis that are clearly discretionary, such as business").font = data_arial_font
    ws.cell(row=37, column=2, value="acquisitions to better analyze the strength of the repayment stream.").font = data_arial_font

    # Addback Net Inc Payback Data
    ws.cell(row=39, column=3, value="L.T. Debt:").font = label_font
    val = anip.get("lt_debt")
    if val is not None:
        val = val / 1_000_000
    data_cell = ws.cell(row=39, column=4, value=val)
    data_cell.number_format = '#,##0'
    data_cell.font = data_arial_font

    ws.cell(row=40, column=3, value="Net Income:").font = label_font
    val = anip.get("net_income")
    if val is not None:
        val = val / 1_000_000
    data_cell = ws.cell(row=40, column=4, value=val)
    data_cell.number_format = '#,##0'
    data_cell.font = data_arial_font

    ws.cell(row=41, column=3, value="Addback:").font = label_font
    val = anip.get("addback")
    if val is not None:
        val = val / 1_000_000
    data_cell = ws.cell(row=41, column=4, value=val)
    data_cell.number_format = '#,##0'
    data_cell.font = data_arial_font
    ws.cell(row=41, column=5, value="Merger charges, writedowns above the line, dep. Amort below the line less capex").font = data_arial_font

    ws.cell(row=42, column=3, value="Years Payback:").font = label_font
    data_cell = ws.cell(row=42, column=4, value=anip.get("years_payback"))
    data_cell.number_format = '#,##0'
    data_cell.font = data_arial_font
    
    # Apply data fills to specified ranges
    def data_fill_range(ws, top_row, left_col, bottom_row, right_col):
        for rr in range(top_row, bottom_row + 1):
            for cc in range(left_col, right_col + 1):
                cell = ws.cell(rr, cc)
                if cell.fill.patternType is None:
                    cell.fill = data_fill

    data_fill_range(ws, 3, 2, 5, 12)   # B3:L5
    data_fill_range(ws, 7, 2, 9, 12)   # B7:L9
    data_fill_range(ws, 15, 2, 17, 12) # B15:L17
    data_fill_range(ws, 23, 2, 25, 12) # B23:L25
    data_fill_range(ws, 35, 2, 37, 12) # B35:L37

def to_float(val):
    if val is None or val == "":
        return None
    val_str = str(val).replace("%", "").replace(",", "").strip()
    if val_str == "":
        return None
    try:
        return float(val_str)
    except ValueError:
        return None
    
def write_qualities_sheet(writer, final_output):
    """
    Create or update a sheet called 'Qualities' that displays the text from
    final_output['qualities'] with bold formatting for specific segments.
    Each '\n' in the text forces a new line in the sheet.
    Within each line, we wrap at ~100 characters. Times New Roman, size 10 font.
    """
    # Early return if qualities is None/null
    if not final_output.get("qualities"):
        return
    
    wb = writer.book

    # If the sheet doesn't exist yet, create it
    if "Qualities" not in wb.sheetnames:
        wb.create_sheet("Qualities")
    ws = wb["Qualities"]

    # Retrieve the forum summary text
    text = final_output.get("qualities", "No forum summary available.")

    # Optional: set a descriptive title in the first cell
    ws["A1"] = "Core Analysis"
    ws["A1"].font = Font(name="Times New Roman", size=14, bold=True)
    ws["A1"].fill = label_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].border = thin_border

    # Split the text by newlines, so each separate line starts on a new row
    lines = text.split("\n")

    start_row = 3
    col = 1
    current_row = start_row

    bold_pattern = re.compile(r"\*\*(.*?)\*\*")

    for line in lines:
        # 1) Trim leading/trailing spaces
        line = line.strip()

        # 2) Remove a leading dash (e.g. "- " or just "-")
        #    This regex says: if the line starts with a dash and optional spaces, remove them.
        line = re.sub(r"^-\s*", "", line)

        # If the line is now empty after removing dash and space, still move down a row
        if not line:
            ws.cell(row=current_row, column=col, value="")
            current_row += 1
            continue

        # Word-wrap the line at ~100 characters while preserving bold formatting
        segments = []
        last_index = 0

        # Find all bold segments (those surrounded by **)
        for match in bold_pattern.finditer(line):
            start, end = match.span()
            if last_index < start:
                # Add text before the bold part
                segments.append((line[last_index:start], False))
            # Add the bold part
            segments.append((match.group(1), True))
            last_index = end

        # Add any text after the last bold segment
        if last_index < len(line):
            segments.append((line[last_index:], False))

        # Now write the segments into the Excel sheet
        for segment, is_bold in segments:
            wrapped_lines = textwrap.wrap(segment, width=100)
            for wrapped_segment in wrapped_lines:
                cell = ws.cell(row=current_row, column=col, value=wrapped_segment)
                cell.font = Font(name="Arial", size=10, bold=is_bold)
                current_row += 1

    # Optionally set column width so text fits nicely
    ws.column_dimensions[get_column_letter(col)].width = 110

    # Remove gridlines on this sheet if you prefer
    ws.sheet_view.showGridLines = False

def write_industry_sheet(writer, final_output):
    """
    Write the Industry sheet with operating and market statistics
    
    Parameters:
    writer: ExcelWriter object
    final_output: Dictionary containing the full output data including industry statistics
    """
    # Early return if qualities is None/null
    if not final_output.get("industry"):
        return
    
    wb = writer.book
    
    # If the sheet doesn't exist yet, create it
    if "Industry" not in wb.sheetnames:
        wb.create_sheet("Industry")
    ws = wb["Industry"]

    industry_data = final_output["industry"]
    industry_name = final_output["summary"]["industry"]
    # Write and format the "Industry Overview" title
    title_cell = ws.cell(row=1, column=6, value=f"Industry Comparison: {industry_name}")
    title_cell.fill = label_fill
    title_cell.font = title_font
    title_cell.alignment = center_alignment
    apply_table_border(ws, 1, 4, 8)
    title_fill_range(ws, 1, 4, 8)

    # Write Operating Statistics section
    op_stats_cell = ws.cell(row=3, column=2, value="Operating Statistics:")
    op_stats_cell.fill = label_fill
    op_stats_cell.font = label_font
    op_stats_cell.border = thin_border
    
    # Get companies (tickers)
    companies = industry_data["operatingStatistics"].keys()
    
    # Define the operating statistics columns and their formats
    op_stats_columns = {
        "Company": (2, None),  # Column B, no special format
        "Debt(yrs.)": (4, '#,##0.0'),  # Column D
        "Sales": (6, '#,##0'),  # Column F
        "ROC": (8, '0.0%'),  # Column H
        "Operating Margin": (10, '0.0%')  # Column J
    }

    # Write operating statistics headers and data
    row = 5  # Start at row 5
    
    # Write headers
    for label, (col, _) in op_stats_columns.items():
        cell = ws.cell(row=row, column=col, value=label)
        cell.fill = label_fill
        cell.font = label_font
        cell.border = thin_border
        cell.alignment = center_alignment

    # Write company data
    for company in companies:
        row += 1
        # Write company name
        company_cell = ws.cell(row=row, column=2, value=company)
        company_cell.font = data_arial_font
        company_cell.alignment = center_alignment
        company_data = industry_data["operatingStatistics"][company]
        
        # Write operating statistics
        col_mappings = {
            "Debt(yrs.)": (4, "Debt(yrs.)"),
            "Sales": (6, "Sales"),
            "ROC": (8, "ROC"),
            "Operating Margin": (10, "Operating Margin")
        }

        for label, (col, key) in col_mappings.items():
            value = company_data[key]
            cell = ws.cell(row=row, column=col, value=value)
            cell.font = data_arial_font
            cell.alignment = center_alignment
            cell.number_format = op_stats_columns[label][1]

            # Convert Sales to millions
            if key == "Sales":
                cell.value = value / 1_000_000

    last_op_stats_row = row

    # Market Statistics section
    market_stats_start_row = last_op_stats_row + 3
    market_stats_cell = ws.cell(row=market_stats_start_row, column=2, value="Market Statistics:")
    market_stats_cell.fill = label_fill
    market_stats_cell.font = label_font
    market_stats_cell.border = thin_border

    # Define the market statistics columns and their formats
    market_stats_columns = {
        "Company": (2, None),  # Column B, no special format
        "P/B": (4, '#,##0.00'),  # Column D
        "P/E": (6, '#,##0.0'),  # Column F
        "Div. Yld.": (8, '0.00%'),  # Column H
        "EV/Sales": (10, '#,##0.00')  # Column J
    }

    # Write market statistics headers
    row = market_stats_start_row + 2
    for label, (col, _) in market_stats_columns.items():
        cell = ws.cell(row=row, column=col, value=label)
        cell.fill = label_fill
        cell.font = label_font
        cell.border = thin_border
        cell.alignment = center_alignment

    # Write market statistics data
    for company in companies:
        row += 1
        # Write company name
        company_cell = ws.cell(row=row, column=2, value=company)
        company_cell.font = data_arial_font
        company_cell.alignment = center_alignment
        company_data = industry_data["marketStatistics"][company]
        
        # Write market statistics
        col_mappings = {
            "P/B": (4, "P/B"),
            "P/E": (6, "P/E"),
            "Div. Yld.": (8, "Div. Yld."),
            "EV/Sales": (10, "EV/Sales")
        }

        for label, (col, key) in col_mappings.items():
            value = company_data[key]
            cell = ws.cell(row=row, column=col, value=value)
            cell.font = data_arial_font
            cell.alignment = center_alignment
            cell.number_format = market_stats_columns[label][1]

    # Adjust column widths
    for col in range(1, 11):
        ws.column_dimensions[get_column_letter(col)].width = 15

def write_hist_pricing_sheet(writer, final_output):
    """
    Write the Historical Pricing sheet with average ratios and price implications in a 2x2 grid layout
    
    Parameters:
    writer: ExcelWriter object
    final_output: Dictionary containing the full output data including historical pricing and current metrics
    """
    wb = writer.book
    
    if "Hist. Pricing" not in wb.sheetnames:
        wb.create_sheet("Hist. Pricing")
    ws = wb["Hist. Pricing"]

    hist_pricing = final_output.get("historical_pricing", {})
    current_metrics = final_output.get("current_metrics", {})
    
    # Write and format the title
    title_cell = ws.cell(row=1, column=4, value="Historical Pricing Analysis")
    title_cell.fill = label_fill
    title_cell.font = title_font
    title_cell.alignment = center_alignment
    apply_table_border(ws, 1, 3, 6)
    title_fill_range(ws, 1, 3, 6)

    # Get the first new year column letter from Co. Desc sheet for formulas
    sorted_years = sorted(final_output["company_description"]["data"].keys(), key=lambda x: int(x))
    if sorted_years:
        max_year = max(int(year) for year in sorted_years)
        new_year = str(max_year + 1)
        # In Co. Desc sheet, years start at column B (2)
        first_new_year_col = get_column_letter(2 + len(sorted_years))
    else:
        first_new_year_col = "B"  # fallback

    # Define the grid positions for each metric
    metrics = {
        # Top Left - P/E Ratio
        "P/E Ratio": {
            "low_key": "avg_pe_low",
            "high_key": "avg_pe_high",
            "current_formula": f"='Co. Desc'!{first_new_year_col}5",  # References diluted EPS
            "start_row": 3,
            "start_col": 2,
            "format": '#,##0.0',
            "value_type": "earnings"
        },
        # Top Right - P/S Ratio
        "P/S Ratio": {
            "low_key": "avg_ps_low",
            "high_key": "avg_ps_high",
            "current_formula": f"='Analyses'!{first_new_year_col}11",  # References Sales/Share
            "start_row": 3,
            "start_col": 8,
            "format": '#,##0.00',
            "value_type": "sales"
        },
        # Bottom Left - P/B Ratio
        "P/B Ratio": {
            "low_key": "avg_pb_low",
            "high_key": "avg_pb_high",
            "current_formula": f"='Co. Desc'!{first_new_year_col}20",  # References Book Value/Share
            "start_row": 10,
            "start_col": 2,
            "format": '#,##0.00',
            "value_type": "book_value"
        },
        # Bottom Right - P/CF Ratio
        "P/CF Ratio": {
            "low_key": "avg_pcf_low",
            "high_key": "avg_pcf_high",
            "current_formula": f"=('Co. Desc'!{first_new_year_col}4+'Analyses'!{first_new_year_col}22)/'Co. Desc'!{first_new_year_col}16",  # (Net Profit + Depreciation) / Shares Outstanding
            "start_row": 10,
            "start_col": 8,
            "format": '#,##0.00',
            "value_type": "cash_flow"
        }
    }

    # Write each metric box
    for metric, props in metrics.items():
        start_row = props["start_row"]
        start_col = props["start_col"]
        
        # Write box title
        title_cell = ws.cell(row=start_row, column=start_col, value=metric)
        title_cell.fill = label_fill
        title_cell.font = label_font
        title_cell.alignment = center_alignment
        ws.merge_cells(start_row=start_row, start_column=start_col, 
                      end_row=start_row, end_column=start_col + 1)
        
        # Write metric rows vertically
        metrics_data = [
            ("Used", props["current_formula"]),  # Now using formula instead of value
            ("Avg Low", hist_pricing.get(props["low_key"])),
            ("Avg High", hist_pricing.get(props["high_key"]))
        ]
        
        for idx, (label, value) in enumerate(metrics_data):
            # Write label
            label_cell = ws.cell(row=start_row + 1 + idx, column=start_col, value=label)
            label_cell.fill = label_fill
            label_cell.font = label_font
            label_cell.border = thin_border
            label_cell.alignment = center_alignment
            
            # Write value or formula
            if idx == 0:  # "Used" row
                value_cell = ws.cell(row=start_row + 1 + idx, column=start_col + 1)
                value_cell.value = value  # This is now a formula
            else:
                value_cell = ws.cell(row=start_row + 1 + idx, column=start_col + 1, value=value)
            
            value_cell.fill = data_fill
            value_cell.font = data_arial_font
            value_cell.border = thin_border
            value_cell.number_format = props["format"]
            value_cell.alignment = right_alignment
        
        # Write Buy and Sell rows with formulas
        used_cell = f"{get_column_letter(start_col + 1)}{start_row + 1}"
        avg_low_cell = f"{get_column_letter(start_col + 1)}{start_row + 2}"
        avg_high_cell = f"{get_column_letter(start_col + 1)}{start_row + 3}"
        
        # Write Buy row (Used * Avg Low)
        buy_label = ws.cell(row=start_row + 4, column=start_col, value="Buy")
        buy_label.fill = label_fill
        buy_label.font = label_font
        buy_label.border = thin_border
        buy_label.alignment = center_alignment
        
        buy_cell = ws.cell(row=start_row + 4, column=start_col + 1, 
                          value=f"={used_cell}*{avg_low_cell}")
        buy_cell.fill = data_fill
        buy_cell.font = data_arial_font
        buy_cell.border = thin_border
        buy_cell.number_format = '#,##0.00'
        buy_cell.alignment = right_alignment
        
        # Write Sell row (Used * Avg High)
        sell_label = ws.cell(row=start_row + 5, column=start_col, value="Sell")
        sell_label.fill = label_fill
        sell_label.font = label_font
        sell_label.border = thin_border
        sell_label.alignment = center_alignment
        
        sell_cell = ws.cell(row=start_row + 5, column=start_col + 1, 
                           value=f"={used_cell}*{avg_high_cell}")
        sell_cell.fill = data_fill
        sell_cell.font = data_arial_font
        sell_cell.border = thin_border
        sell_cell.number_format = '#,##0.00'
        sell_cell.alignment = right_alignment
        
        # Add border around the entire box
        for row in range(start_row, start_row + 6):
            for col in range(start_col, start_col + 2):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border

    # Adjust column widths
    for base_col in [2, 8]:  # Starting columns for left and right sections
        ws.column_dimensions[get_column_letter(base_col)].width = 12      # Labels
        ws.column_dimensions[get_column_letter(base_col + 1)].width = 20  # Values

def generate_excel_for_ticker_year(ticker: str, year: int):
    """
    Generate the Excel file for the given ticker and year, writing to:
       ./output/{ticker}.{last 2 digits of year}.2.xlsx
    
    :param ticker: The company symbol/ticker.
    :param year:   The 4-digit year (e.g., 2024). We'll use only the last two digits in the filename.
    """
    # Convert to uppercase and build the filename, e.g.: ./output/ABC.24.2.xlsx
    ticker = ticker.upper()
    year_2_digits = str(year)[-2:]  # last two chars of the year
    xls_filename = os.path.join("output", f"{ticker}.{year_2_digits}.2.xlsx")

    # 1. Load the JSON data
    final_output = load_final_output(ticker)

    # 2. Create or append to the Excel file
    writer = create_xls(xls_filename)

    # 3. Write data to each sheet
    write_summary_sheet(writer, final_output)
    write_company_description(writer, final_output)
    write_analyses_sheet(writer, final_output)
    write_profit_desc_sheet(writer, final_output)
    write_balance_sheet_sheet(writer, final_output)
    write_studies_sheet(writer, final_output)
    write_qualities_sheet(writer, final_output)
    write_industry_sheet(writer, final_output)
    write_hist_pricing_sheet(writer, final_output)

    # 4. Apply workbook formatting (remove gridlines, etc.)
    format_workbook(writer)

    # 5. Save
    writer.close()
    print(f"Data for {ticker} written to {xls_filename} successfully.")


if __name__ == "__main__":
    # Usage: python write_excel.py TICKER target_file.xls
    if len(sys.argv) < 3:
        print("Usage: python gen_excel.py TICKER XLS_FILENAME")
        sys.exit(1)

    ticker = sys.argv[1].upper()
    xls_filename = sys.argv[2]

    final_output = load_final_output(ticker)

    # Create or append to xls file
    writer = create_xls(xls_filename)

    # Now write data to each sheet
    write_summary_sheet(writer, final_output)
    write_company_description(writer, final_output)
    write_analyses_sheet(writer, final_output)
    write_profit_desc_sheet(writer, final_output)
    write_balance_sheet_sheet(writer, final_output)
    write_studies_sheet(writer, final_output)
    write_qualities_sheet(writer, final_output)
    write_industry_sheet(writer, final_output)
    write_hist_pricing_sheet(writer, final_output)

    # Apply formatting: set font to Arial size 10 for non-formatted cells and remove gridlines
    format_workbook(writer)

    # Save changes
    writer.close()
    print(f"Data for {ticker} written to {xls_filename} successfully.")
