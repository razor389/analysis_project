# gen_excel.py

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

from utils import get_current_quote_yahoo

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
    sp_cell.number_format = '#,##0.00'
    sp_cell.border = thin_border

    market_cap = to_float(cd_info.get("marketCapitalization")) / 1_000_000
    mc_cell = ws.cell(row=2, column=9, value=market_cap)
    mc_cell.fill = data_fill
    mc_cell.font = label_font
    mc_cell.number_format = '#,##0'
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
        "assets"
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
        "assets": 22,               
        "return_on_equity": 24,     
        "return_on_assets": 25,     
        "leverage_ratio": 26        
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
        "assets": "Total Assets",
        "return_on_equity": "Return on Equity",
        "return_on_assets": "Return on Assets",
        "leverage_ratio": "Leverage Ratio"
    }

    # Updated number formats according to specifications
    number_formats = {
        "net_profit": '#,##0',          # Millions with commas, no $ no M
        "diluted_eps": '#,##0.00',      # Dollars and cents
        "operating_eps": '#,##0.00',    # Dollars and cents
        "pe_ratio": '#,##0.0',          # One decimal place
        "price_low": '#,##0.00',        # Dollars and cents
        "price_high": '#,##0.00',       # Dollars and cents
        "dividends_paid": '#,##0',      # Millions with commas, no $ no M
        "dividends_per_share": '#,##0.00',  # Dollars and cents
        "avg_dividend_yield": '0.00%',   # Percentage with two decimal places
        "shares_outstanding": '#,##0',   # Millions with commas, no $ no M
        "buyback": '#,##0',             # Millions with commas, no $ no M
        "share_equity": '#,##0',        # Millions with commas, no $ no M
        "book_value_per_share": '#,##0.00',  # Dollars and cents
        "assets": '#,##0',              # Millions with commas, no $ no M
        "return_on_equity": '0.00%',    # Percentage with two decimal places
        "return_on_assets": '0.00%',    # Percentage with two decimal places
        "leverage_ratio": '#,##0.0'     # One decimal place
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
                
                elif metric == "return_on_equity":
                    formula = f"={col_letter}{metric_positions['net_profit']}/{col_letter}{metric_positions['share_equity']}"
                    data_cell = ws.cell(row=metric_row, column=col, value=formula)
                
                elif metric == "return_on_assets":
                    formula = f"={col_letter}{metric_positions['net_profit']}/{col_letter}{metric_positions['assets']}"
                    data_cell = ws.cell(row=metric_row, column=col, value=formula)
                
                elif metric == "leverage_ratio":
                    formula = f"={col_letter}{metric_positions['assets']}/{col_letter}{metric_positions['share_equity']}"
                    data_cell = ws.cell(row=metric_row, column=col, value=formula)
                
                else:
                    value = cd_data.get(year, {}).get(metric)
                    if value is not None and metric in million_scale_metrics:
                        value = value / 1_000_000
                    data_cell = ws.cell(row=metric_row, column=col, value=value)

                data_cell.font = data_tnr_italic_font
                data_cell.alignment = right_alignment
            
            else:
                # For historical years, map the new field names to their values
                if metric == "return_on_equity":
                    value = cd_data.get(year, {}).get("return_on_equity")
                elif metric == "return_on_assets":
                    value = cd_data.get(year, {}).get("return_on_assets")
                elif metric == "leverage_ratio":
                    value = cd_data.get(year, {}).get("leverage_ratio")
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
        (16, 3, "Premium Analysis:"),
        (18, 4, "Growth Rate %:"),
        (20, 4, "Growth Rate PS %:"),
        (16, 8, "Premium Analysis (last 5 yrs.):"),
        (18, 9, "Growth Rate %:"),
        (20, 9, "Growth Rate PS %:")
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

    # Write the investment characteristics data
    data_cells = {
        (5, 5): growth_rate_operating_eps,
        (7, 5): quality_percent,
        (5, 10): avg_div_payout,
        (7, 10): avg_stk_buyback,
        (18, 5): growth_rate_rev,
        (20, 5): growth_rate_sps,
        (18, 10): growth_rate_rev_5y,
        (20, 10): growth_rate_sps_5y
    }

    for (row, col), value in data_cells.items():
        cell = ws.cell(row=row, column=col, value=value)
        cell.fill = data_fill
        cell.font = data_tnr_bold_font
        cell.border = thin_border
        cell.number_format = '0.0%'

    # Handle the data by years
    sorted_years = sorted(data.keys(), key=lambda x: int(x))
    
    # Determine the next two years
    if sorted_years:
        max_year = max(int(year) for year in sorted_years)
        new_years = [str(max_year + 1), str(max_year + 2)]
    else:
        new_years = ["2024", "2025"]

    all_years = sorted_years + new_years
    start_col = 2  # Column B

    # Write years at row 9 and row 22
    for i, year in enumerate(all_years):
        for row in [9, 22]:
            cell = ws.cell(row=row, column=start_col + i, value=year)
            cell.fill = label_fill
            cell.font = label_font
            cell.border = thin_border

    # Define the insurance metrics and their row positions
    metric_rows_1 = {
        "premium_earned": 10,
        "benefit_claims": 11,
        "gross_underwriting_profit": 12,
        "underwriting_yield_on_asset": 13,
        "investment_income": 14,
        "investment_yield_on_asset": 15
    }

    metric_rows_2 = {
        "non_claim_expenses": 23,
        "expense_yield_on_asset": 24,
        "tax_rate": 25,
        "premium_equity_ratio": 26
    }

    # Add labels for metrics
    additional_labels = {
        10: "Premium Earned",
        11: "Benefit Claims",
        12: "Gross Underwriting Profit",
        13: "Underwriting Yield on Asset",
        14: "Investment Income",
        15: "Investment Yield on Asset",
        23: "Non-Claim Expenses",
        24: "Expense Yield on Asset",
        25: "Tax Rate",
        26: "Premium/Equity Ratio"
    }

    # Define formatting rules
    number_formats = {
        "premium_earned": '#,##0',
        "benefit_claims": '#,##0',
        "gross_underwriting_profit": '#,##0',
        "underwriting_yield_on_asset": '0.00%',
        "investment_income": '#,##0',
        "investment_yield_on_asset": '0.00%',
        "non_claim_expenses": '#,##0',
        "expense_yield_on_asset": '0.00%',
        "tax_rate": '0.00%',
        "premium_equity_ratio": '0.00'
    }

    # Define metrics that should be displayed in millions
    million_scale_metrics = {
        "premium_earned",
        "benefit_claims",
        "gross_underwriting_profit",
        "investment_income",
        "non_claim_expenses"
    }

    for row, label in additional_labels.items():
        cell = ws.cell(row=row, column=1, value=label)
        cell.fill = label_fill
        cell.font = label_font
        cell.border = thin_border

    # Write both sets of metrics
    for metric_rows in [metric_rows_1, metric_rows_2]:
        for metric, row_num in metric_rows.items():
            for i, year in enumerate(all_years):
                col = start_col + i
                
                if year in new_years:
                    val = None  # Leave future years blank for insurance metrics
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
    """
    Write the profit description sheet with updated handling for operating earnings breakdowns
    
    Parameters:
    writer: ExcelWriter object
    final_output: Dictionary containing the full output data
    """
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

    # Step 1: Collect all unique breakdowns across all years
    all_revenue_breakdowns = set()
    all_expense_breakdowns = set()
    all_operating_earnings_breakdowns = set()
    all_external_costs_breakdowns = set()
    
    for year_data in pdata.values():
        # Handle revenue breakdowns
        revenues = year_data.get("revenues", {})
        if revenue_breakdown := revenues.get("breakdown", {}):
            all_revenue_breakdowns.update(revenue_breakdown.keys())
            
        # Handle expense breakdowns
        expenses = year_data.get("expenses", {})
        if expense_breakdown := expenses.get("breakdown", {}):
            all_expense_breakdowns.update(expense_breakdown.keys())
            
        # Handle operating earnings breakdowns
        operating_earnings = year_data.get("operating_earnings", {})
        if operating_breakdown := operating_earnings.get("breakdown", {}):
            all_operating_earnings_breakdowns.update(operating_breakdown.keys())
            
        # Handle external costs breakdowns
        external_costs = year_data.get("external_costs", {})
        if external_costs_breakdown := external_costs.get("breakdown", {}):
            all_external_costs_breakdowns.update(external_costs_breakdown.keys())

    # Step 2: Write the metrics and their labels
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

        # Handle breakdowns based on metric type
        if metric == "revenues" and all_revenue_breakdowns:
            for bkey in sorted(all_revenue_breakdowns):
                breakdown_cell = ws.cell(row=current_row, column=2, value=bkey)
                breakdown_cell.font = Font(italic=True)
                breakdown_rows[(metric, bkey)] = current_row
                current_row += 1
        elif metric == "expenses" and all_expense_breakdowns:
            for bkey in sorted(all_expense_breakdowns):
                breakdown_cell = ws.cell(row=current_row, column=2, value=bkey)
                breakdown_cell.font = Font(italic=True)
                breakdown_rows[(metric, bkey)] = current_row
                current_row += 1
        elif metric == "operating_earnings" and all_operating_earnings_breakdowns:
            for bkey in sorted(all_operating_earnings_breakdowns):
                breakdown_cell = ws.cell(row=current_row, column=2, value=bkey)
                breakdown_cell.font = Font(italic=True)
                breakdown_rows[(metric, bkey)] = current_row
                current_row += 1
        elif metric == "external_costs" and all_external_costs_breakdowns:
            for bkey in sorted(all_external_costs_breakdowns):
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
        y_cell.alignment = center_alignment

        # Add hyperlink if filing URL exists for this year
        filing_url = pdata.get(year, {}).get("filing_url")
        if filing_url:
            y_cell.hyperlink = filing_url
            y_cell.font = Font(name="Times New Roman", size=10, bold=True, italic=True, underline="single", color="0000FF")
            
    # Step 4: Write Metric Values and Breakdown Values
    for i, year in enumerate(sorted_years):
        year_col = start_col_for_years + i * 2  # Column for the value
        year_data = pdata[year]
        
        for metric in metrics_order:
            metric_val = year_data.get(metric)
            metric_row = metric_rows.get(metric)

            if metric_val is None:
                continue

            if isinstance(metric_val, dict) and (
                ("total_" + metric) in metric_val  # for operating_earnings
                or ("total_" + metric.replace("_earnings", "")) in metric_val  # for revenues/external_costs
            ):
                # Handle metrics with breakdowns (including operating earnings)
                # Handle both operating_earnings and other metrics
                total_key = "total_operating_earnings" if metric == "operating_earnings" else "total_" + metric.replace("_earnings", "")
                
                if total_key in metric_val:
                    val = to_float(metric_val[total_key])
                    if val is not None:
                        val = val / 1_000_000
                    cell = ws.cell(row=metric_row, column=year_col, value=val)
                    cell.fill = data_fill
                    cell.font = Font(name="Arial", italic=True)
                    cell.number_format = '#,##0'
                    cell.border = thin_border

                # Write breakdown items
                breakdown_items = metric_val.get("breakdown", {})
                for (m, bkey), brow in breakdown_rows.items():
                    if m == metric and bkey in breakdown_items:
                        breakdown_val = breakdown_items[bkey]
                        if breakdown_val is not None:
                            breakdown_val = to_float(breakdown_val)
                            if breakdown_val is not None:
                                breakdown_val = breakdown_val / 1_000_000
                            bdata_cell = ws.cell(row=brow, column=year_col, value=breakdown_val)
                            bdata_cell.font = data_arial_italic_font
                            bdata_cell.number_format = '#,##0'
                            
                            # Add appropriate CAGR values
                            if metric == "revenues":
                                cagr_key = f"cagr_revenues_{bkey}_percent"
                                cagr_value = pchar.get("cagr_revenues_breakdown_percent", {}).get(cagr_key)
                                if cagr_value is not None:
                                    cagr_cell = ws.cell(row=brow, column=3, value=cagr_value)
                                    cagr_cell.font = Font(name="Arial", italic=True, size=8)
                                    cagr_cell.number_format = '0.0%'
                            elif metric == "operating_earnings":
                                cagr_key = f"cagr_operating_earnings_{bkey}_percent"
                                cagr_value = pchar.get("cagr_operating_earnings_breakdown_percent", {}).get(cagr_key)
                                if cagr_value is not None:
                                    cagr_cell = ws.cell(row=brow, column=3, value=cagr_value)
                                    cagr_cell.font = Font(name="Arial", italic=True, size=8)
                                    cagr_cell.number_format = '0.0%'

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

    # Handle expense breakdown CAGRs
    expense_cagr_map = {
        "cost_of_revenue": "cagr_cost_of_revenue_percent",
        "research_and_development": "cagr_research_and_development_percent",
        "selling_marketing_general_admin": "cagr_selling_marketing_general_admin_percent"
    }

    # Write expense breakdown CAGRs
    for (metric, bkey), brow in breakdown_rows.items():
        if metric == "expenses" and bkey in expense_cagr_map:
            cagr_key = expense_cagr_map[bkey]
            cagr_value = pchar.get(cagr_key)
            if cagr_value is not None:
                cagr_cell = ws.cell(row=brow, column=3, value=cagr_value)
                cagr_cell.font = Font(name="Arial", italic=True, size=8)
                cagr_cell.number_format = '0.00%'
        elif metric == "external_costs":
            # Handle external costs breakdown CAGRs
            cagr_key = f"cagr_external_costs_{bkey}_percent"
            cagr_value = pchar.get("cagr_external_costs_breakdown_percent", {}).get(cagr_key)
            if cagr_value is not None:
                cagr_cell = ws.cell(row=brow, column=3, value=cagr_value)
                cagr_cell.font = Font(name="Arial", italic=True, size=8)
                cagr_cell.number_format = '0.00%'

    for metric, cagr_key in cagr_map.items():
        cagr_value = pchar.get(cagr_key)
        if cagr_value is not None and metric in metric_rows:
            row = metric_rows[metric]
            cagr_cell = ws.cell(row=row, column=3, value=cagr_value)
            cagr_cell.font = Font(name="Arial", italic=True, size=8)
            cagr_cell.number_format = '0.00%'

    # Step 6: Compute and Write Percentages
    revenues_row = metric_rows.get("revenues")
    expense_breakdowns = ["cost_of_revenue", "research_and_development", "selling_marketing_general_admin"]
    top_metrics = ["ebitda", "operating_earnings", "earnings"]  # Metrics that get percentage calculations

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

        # Calculate percentages for operating earnings breakdowns
        for bkey in all_operating_earnings_breakdowns:
            br_key = ("operating_earnings", bkey)
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
    """
    Writes a Balance Sheet worksheet to the Excel workbook using data from final_output.
    This version is flexible with the input data structure and preserves the order of breakdown items:
      - If a section (assets, liabilities, shareholders_equity) has a key equal to its name,
        that value is treated as the total for that section.
      - Any other keys in the section (or, if provided, in a nested "breakdown" dict)
        are treated as breakdown items.
    Values are converted from raw units to millions.
    """
    # Assume these style objects and helper functions are defined elsewhere:
    #   label_fill, label_font, title_font, data_fill, data_arial_bold_font,
    #   data_arial_italic_font, thin_border, title_fill_range, apply_table_border, to_float

    reported_currency = final_output["summary"]["reported_currency"]
    bs_info = final_output["balance_sheet"]
    bs_char = bs_info["balance_sheet_characteristics"]
    bs_data = bs_info["data"]
    wb = writer.book

    # Create or get the worksheet
    sheet_name = "Balance Sht."
    if sheet_name not in wb.sheetnames:
        wb.create_sheet(sheet_name)
    ws = wb[sheet_name]
    ws.freeze_panes = "F1"

    # Write and format the title
    title_cell = ws.cell(row=1, column=4, value=f"Balance Sheet (in mlns {reported_currency}):")
    title_cell.fill = label_fill
    title_cell.font = title_font
    title_fill_range(ws, 1, 4, 7)
    apply_table_border(ws, 1, 4, 7)

    # Extract CAGR values (if available)
    cagr_assets = bs_char.get("cagr_total_assets_percent")
    cagr_liabilities = bs_char.get("cagr_total_liabilities_percent")
    cagr_equity = bs_char.get("cagr_total_shareholders_equity_percent")

    # Define the top-level sections and their labels
    top_sections = [
        ("assets", "Assets:"),
        ("liabilities", "Liabilities:"),
        ("shareholders_equity", "Shareholder's Equity:")
    ]

    # Get the years in sorted order (assumes year keys are numeric strings)
    sorted_years = sorted(bs_data.keys(), key=lambda x: int(x))
    start_col_for_years = 6  # Column F

    # Write and format the header row with years (row 3)
    for i, year in enumerate(sorted_years):
        year_col = start_col_for_years + i
        y_cell = ws.cell(row=3, column=year_col, value=year)
        y_cell.fill = label_fill
        y_cell.font = label_font
        y_cell.border = thin_border

    # Start writing section rows from row 5
    current_row = 5

    for section_key, section_label in top_sections:
        # Write the section label in column A with header formatting.
        sec_label_cell = ws.cell(row=current_row, column=1, value=section_label)
        sec_label_cell.fill = label_fill
        sec_label_cell.font = label_font
        title_fill_range(ws, current_row, 1, 5)
        apply_table_border(ws, current_row, 1, 5)
        total_row = current_row

        # -- Write the total row for the section (if present) --
        # In the new format, we expect that if a total is provided it is stored under the key
        # that is the same as the section_key (e.g. "assets" for assets). Otherwise,
        # there is no total row.
        total_found = False
        for year in sorted_years:
            sec_data = bs_data.get(year, {}).get(section_key, {})
            if section_key in sec_data:
                total_found = True
                break

        if total_found:
            for i, year in enumerate(sorted_years):
                sec_data = bs_data.get(year, {}).get(section_key, {})
                # Use the value from the key equal to the section name (if available)
                raw_val = sec_data.get(section_key)
                val = to_float(raw_val) if raw_val is not None else None
                if val is not None:
                    val = val / 1_000_000  # Convert to millions
                d_cell = ws.cell(row=total_row, column=start_col_for_years + i, value=val)
                d_cell.fill = data_fill
                d_cell.font = data_arial_bold_font
                d_cell.border = thin_border
                if isinstance(val, (int, float)):
                    d_cell.number_format = '#,##0'
            # Apply the corresponding CAGR value in column E (if available)
            if section_key == "assets" and cagr_assets is not None:
                ws.cell(row=total_row, column=5, value=cagr_assets).number_format = '0.0%'
            elif section_key == "liabilities" and cagr_liabilities is not None:
                ws.cell(row=total_row, column=5, value=cagr_liabilities).number_format = '0.0%'
            elif section_key == "shareholders_equity" and cagr_equity is not None:
                ws.cell(row=total_row, column=5, value=cagr_equity).number_format = '0.0%'
            current_row += 1  # Advance to the next row after the total row

        # -- Determine breakdown keys while preserving original order --
        # Instead of sorting keys alphabetically, we iterate over the years and add keys
        # in the order they first appear.
        breakdown_keys = []
        for year in sorted_years:
            sec_data = bs_data.get(year, {}).get(section_key, {})
            if "breakdown" in sec_data and isinstance(sec_data["breakdown"], dict):
                for key in sec_data["breakdown"]:
                    if key not in breakdown_keys:
                        breakdown_keys.append(key)
            else:
                for key in sec_data:
                    if key != section_key and key not in breakdown_keys:
                        breakdown_keys.append(key)

        # -- Write breakdown rows --
        for bkey in breakdown_keys:
            # Write the breakdown label in column B (no special fill)
            ws.cell(row=current_row, column=2, value=bkey)
            # Write data for each year for this breakdown item
            for i, year in enumerate(sorted_years):
                sec_data = bs_data.get(year, {}).get(section_key, {})
                # If a nested breakdown exists, use it; otherwise, read directly.
                if "breakdown" in sec_data and isinstance(sec_data["breakdown"], dict):
                    raw_val = sec_data["breakdown"].get(bkey)
                else:
                    raw_val = sec_data.get(bkey)
                val = to_float(raw_val) if raw_val is not None else None
                if val is not None:
                    val = val / 1_000_000  # convert to millions
                bdata_cell = ws.cell(row=current_row, column=start_col_for_years + i, value=val)
                bdata_cell.font = data_arial_italic_font
                if isinstance(val, (int, float)):
                    bdata_cell.number_format = '#,##0'
            current_row += 1

        # Add a blank line before the next section
        current_row += 1

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
    final_output['qualities'] with:
    - Numbered items (1-10)
    - Bold headers inline with numbers
    - Wrapped descriptive text starting on the next line
    Each quality is separated by a blank line for readability.
    """
    if not final_output.get("qualities"):
        return
    
    wb = writer.book

    # Create sheet if it doesn't exist
    if "Qualities" not in wb.sheetnames:
        wb.create_sheet("Qualities")
    ws = wb["Qualities"]

    # Get the qualities text
    text = final_output.get("qualities", "No summary available.")

    # Set title
    ws["A1"] = "Core Analysis"
    ws["A1"].font = Font(name="Times New Roman", size=14, bold=True)
    ws["A1"].fill = label_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].border = thin_border

    # Split into individual quality entries (split on numbered items)
    qualities = re.split(r'\n\n(?=\d+\.)', text.strip())
    
    current_row = 3
    col = 1

    for quality in qualities:
        if not quality.strip():
            continue
            
        # Extract the number, header, and description
        # Pattern matches: number, bold header, and description
        match = re.match(r'(\d+)\.\s*\*\*(.*?)\*\*:(.+)', quality.strip(), re.DOTALL)
        
        if match:
            number, header, description = match.groups()
            
            # Write the numbered header line
            header_cell = ws.cell(row=current_row, column=col)
            header_cell.value = f"{number}. {header}:"
            header_cell.font = Font(name="Arial", size=10, bold=True)
            current_row += 1
            
            # Write the wrapped description on subsequent lines
            description = description.strip()
            wrapped_lines = textwrap.wrap(description, width=100)
            
            for line in wrapped_lines:
                desc_cell = ws.cell(row=current_row, column=col)
                desc_cell.value = line
                desc_cell.font = Font(name="Arial", size=10)
                current_row += 1
            
            # Add a blank line between qualities
            current_row += 1

    # Set column width
    ws.column_dimensions[get_column_letter(col)].width = 110

    # Remove gridlines
    ws.sheet_view.showGridLines = False

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

def write_segmentation_sheet(writer, final_output):
    """
    Write the Segmentation sheet showing revenue breakdown by business segment over time
    
    Parameters:
    writer: ExcelWriter object
    final_output: Dictionary containing the full output data including segmentation data
    """
    # Early return if segmentation data is not present
    if not final_output.get("segmentation"):
        return
    
    reported_currency = final_output["summary"]["reported_currency"]
    segmentation_data = final_output["segmentation"]
    wb = writer.book
    
    # Create the sheet if it doesn't exist
    if "Segmentation" not in wb.sheetnames:
        wb.create_sheet("Segmentation")
    ws = wb["Segmentation"]
    ws.freeze_panes = "B1"

    # Write and format the title
    title_cell = ws.cell(row=1, column=5, value=f"Revenue Segmentation (in mlns {reported_currency})")
    title_cell.fill = label_fill
    title_cell.font = title_font
    title_cell.alignment = center_alignment
    apply_table_border(ws, 1, 3, 7)
    title_fill_range(ws, 1, 3, 7)

    # Get sorted years and all unique segments
    sorted_years = sorted(segmentation_data.keys(), key=lambda x: int(x))
    all_segments = set()
    for year_data in segmentation_data.values():
        all_segments.update(year_data.keys())
    sorted_segments = sorted(all_segments)

    # Write year headers starting at row 3
    for i, year in enumerate(sorted_years):
        year_col = i + 2  # Start at column B
        y_cell = ws.cell(row=3, column=year_col, value=year)
        y_cell.fill = label_fill
        y_cell.font = label_font
        y_cell.border = thin_border
        y_cell.alignment = center_alignment

    # Write segment data
    current_row = 4
    for segment in sorted_segments:
        # Write segment name in column A
        segment_cell = ws.cell(row=current_row, column=1, value=segment)
        segment_cell.font = label_font
        segment_cell.fill = label_fill
        segment_cell.border = thin_border

        # Write values for each year
        for i, year in enumerate(sorted_years):
            year_col = i + 2  # Start at column B
            value = segmentation_data[year].get(segment)
            
            # Create cell and apply styling regardless of value
            value_cell = ws.cell(row=current_row, column=year_col)
            value_cell.fill = data_fill
            value_cell.font = data_arial_font
            value_cell.border = thin_border
            value_cell.alignment = right_alignment
            
            if value is not None:
                # Convert to millions and set value
                value = value / 1_000_000
                value_cell.value = value
                value_cell.number_format = '#,##0'

        current_row += 1

    # Calculate and write growth rates in the rightmost column
    growth_col = len(sorted_years) + 3  # One column after percentages
    growth_header = ws.cell(row=3, column=growth_col, value="CAGR")
    growth_header.fill = label_fill
    growth_header.font = label_font
    growth_header.border = thin_border
    growth_header.alignment = center_alignment

    for segment_idx, segment in enumerate(sorted_segments):
        row = segment_idx + 4
        
        # Get first and last valid values
        first_val = next((segmentation_data[year].get(segment) for year in sorted_years if segmentation_data[year].get(segment) is not None), None)
        last_val = next((segmentation_data[year].get(segment) for year in reversed(sorted_years) if segmentation_data[year].get(segment) is not None), None)
        
        if first_val and last_val and first_val != 0:
            years_between = len(sorted_years) - 1
            if years_between > 0:
                cagr = (last_val / first_val) ** (1/years_between) - 1
                growth_cell = ws.cell(row=row, column=growth_col, value=cagr)
                growth_cell.number_format = '0.0%'
                growth_cell.font = Font(name="Arial", size=8, italic=True)
                growth_cell.fill = data_fill  # Added fill for consistency
                growth_cell.border = thin_border  # Added border for consistency

        else:
            # Create empty growth cell with consistent styling
            growth_cell = ws.cell(row=row, column=growth_col)
            growth_cell.fill = data_fill
            growth_cell.border = thin_border

    # Adjust column widths
    ws.column_dimensions['A'].width = 30  # Segment names
    for col in range(2, growth_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 12

def generate_config_note(ticker, wb):
    """
    Add a note in cell B1 of the profit_desc sheet if there are any
    configuration overrides for the company from financial_data_config.json.
    
    Args:
        ticker (str): Company ticker symbol
        wb (openpyxl.Workbook): Excel workbook object
    """
    import json
    import logging
    from pathlib import Path

    logger = logging.getLogger(__name__)
    
    # Get the profit_desc sheet
    try:
        sheet = wb["Profit.Desc."]
    except KeyError:
        logger.warning("profit_desc sheet not found in workbook")
        return
    
    # Load the configuration file
    try:
        with open('financial_data_config.json', 'r') as f:
            config = json.load(f)
    except FileNotFoundError:
        logger.warning("financial_data_config.json not found")
        return
    except json.JSONDecodeError:
        logger.warning("Error parsing financial_data_config.json")
        return
    
    # Check if we have any configurations for this ticker
    ticker_config = config.get(ticker.upper(), {})
    if not ticker_config:
        return
    
    # Build the configuration note
    notes = []
    for statement, fields in ticker_config.items():
        field_notes = []
        for field, value in fields.items():
            field_notes.append(f"set {field} to {value}")
        if field_notes:
            notes.append(f"In {statement}, {', '.join(field_notes)}")
    
    if notes:
        note = '. '.join(notes)
        # Add note to cell B1
        sheet['A2'] = f"Configuration overrides: {note}."
        
        # Style the cell
        from openpyxl.styles import Font, PatternFill
        cell = sheet['B1']
        cell.font = Font(italic=True, size=9, color="666666")
        cell.fill = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")

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
    write_qualities_sheet(writer, final_output)
    write_industry_sheet(writer, final_output)
    write_segmentation_sheet(writer, final_output)
    generate_config_note(ticker, writer.book)

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
    write_balance_sheet_sheet(writer, final_output)
    #write_profit_desc_sheet(writer, final_output)
    
    #write_qualities_sheet(writer, final_output)
    #write_industry_sheet(writer, final_output)
    #write_segmentation_sheet(writer, final_output)
    #generate_config_note(ticker, writer.book)

    # Apply formatting: set font to Arial size 10 for non-formatted cells and remove gridlines
    format_workbook(writer)

    # Save changes
    writer.close()
    print(f"Data for {ticker} written to {xls_filename} successfully.")
