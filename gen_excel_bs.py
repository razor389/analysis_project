# gen_excel_bs.py

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
    file_path = os.path.join("output", f"{ticker}_yoy_consolidated_bs.json")
    if not os.path.exists(file_path):
        print(f"Error: {file_path} does not exist. Please generate {ticker}_yoy_consolidated_bs.json first.")
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
    Write the profit description sheet using the new profit_description data format.
    This version dynamically gathers breakdown keys for those metrics that are provided
    as dictionaries (with a "total" and a "breakdown" section). It is written to be flexible
    so that different companies (with different breakdown items) are handled automatically.
    
    Note: The operating margin breakdown items "pretax_combined_ratio", 
          "pretax_insurance_yield_on_equity", and "pretax_return_on_equity" are assumed to be 
          percentages (e.g. 0.86 means 86.0%) and are not divided by 1,000,000.
    
    Parameters:
        writer: ExcelWriter object
        final_output: Dictionary containing the full output data.
                     Expected to have final_output["profit_description"] with keys:
                         - "profit_description_characteristics": a dict with overall CAGR values
                           and breakdown CAGR dictionaries (e.g. "cagr_gross_revenues_breakdown_percent")
                         - "data": a dict with keys (years) whose values are profit metrics
                           e.g. "gross_revenues", "investment_income", "internal_costs",
                           "operating_margin", "external_costs", "earnings", etc.
    """
    # Get reported currency, profit description characteristics and data by year
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
    title_text = f"Description & Analysis of Profitability (in mlns {reported_currency})"
    title_cell = ws.cell(row=1, column=4, value=title_text)
    title_cell.fill = label_fill
    title_cell.font = title_font
    title_fill_range(ws, 1, 3, 10)
    apply_table_border(ws, 1, 3, 10)

    # Define the metric order and labels based on the new format.
    metrics_order = [
        "gross_revenues",
        "investment_income",
        "internal_costs",
        "operating_margin",
        "external_costs",
        "earnings",
        "equity_employed",
        "shares_repurchased"
    ]
    metric_labels = {
        "gross_revenues": "Gross Revenues:",
        "investment_income": "Investment Income:",
        "internal_costs": "Internal Costs:",
        "operating_margin": "Operating Margin:",
        "external_costs": "External Costs:",
        "earnings": "Earnings:",
        "equity_employed": "Equity Employed:",
        "shares_repurchased": "Shares Repurchased:"
    }

    # List of metrics that are percentages (if any)
    percent_metrics = []  # In the new format, most numbers are monetary

    # STEP 1: Collect all unique breakdown keys for those metrics that come as dicts with a breakdown.
    breakdown_source_metrics = ["gross_revenues", "internal_costs", "operating_margin", "external_costs"]
    breakdown_keys = {metric: set() for metric in breakdown_source_metrics}
    for year_data in pdata.values():
        for metric in breakdown_source_metrics:
            mdata = year_data.get(metric)
            if isinstance(mdata, dict):
                bd = mdata.get("breakdown", {})
                if bd:
                    breakdown_keys[metric].update(bd.keys())

    # STEP 2: Write the metric labels and, when available, the breakdown labels.
    current_row = 5  # Starting row for metric rows
    metric_rows = {}      # Track the row for each metric label row
    breakdown_rows = {}   # Track the row for each breakdown item; key = (metric, breakdown_key)
    for metric in metrics_order:
        label = metric_labels.get(metric, metric.capitalize() + ":")
        # Write the main metric label in column 1 (across columns 1-3 for formatting)
        for col in range(1, 4):
            cell = ws.cell(row=current_row, column=col, value=label if col == 1 else None)
            cell.fill = label_fill
            cell.font = label_font
        metric_rows[metric] = current_row
        apply_table_border(ws, current_row, 1, 3)
        current_row += 1

        # If this metric has a breakdown (as discovered above), write each breakdown item as a sub-row.
        if metric in breakdown_keys and breakdown_keys[metric]:
            for bkey in sorted(breakdown_keys[metric]):
                bd_cell = ws.cell(row=current_row, column=2, value=bkey)
                bd_cell.font = Font(italic=True)
                breakdown_rows[(metric, bkey)] = current_row
                current_row += 1

    # STEP 3: Write Year Headers (starting at column D)
    sorted_years = sorted(pdata.keys(), key=lambda x: int(x))
    start_col_for_years = 4  # Column D
    for i, year in enumerate(sorted_years):
        year_col = start_col_for_years + i * 2  # Each year uses two columns: one for value, one for percentage
        y_cell = ws.cell(row=3, column=year_col, value=year)
        y_cell.fill = label_fill
        y_cell.font = label_font
        y_cell.border = thin_border
        y_cell.alignment = center_alignment
        # Add hyperlink to filing URL if available
        filing_url = pdata.get(year, {}).get("filing_url")
        if filing_url:
            y_cell.hyperlink = filing_url
            y_cell.font = Font(name="Times New Roman", size=10, bold=True, italic=True, underline="single", color="0000FF")

    # Define the set of operating_margin breakdown keys that should be treated as percentages.
    operating_margin_pct_keys = {"pretax_combined_ratio", "pretax_insurance_yield_on_equity", "pretax_return_on_equity"}

    # STEP 4: Write Metric Values and Breakdown Values for each year.
    for i, year in enumerate(sorted_years):
        year_col = start_col_for_years + i * 2  # Column for the main value
        year_data = pdata[year]
        for metric in metrics_order:
            mdata = year_data.get(metric)
            if mdata is None:
                continue

            # If the value is given as a dict with a "total" field (and possibly a breakdown)
            if isinstance(mdata, dict) and "total" in mdata:
                val = to_float(mdata["total"])
                if val is not None:
                    # For monetary metrics, divide by 1,000,000
                    val = val / 1_000_000
                cell = ws.cell(row=metric_rows[metric], column=year_col, value=val)
                cell.fill = data_fill
                cell.font = Font(name="Arial", italic=True)
                cell.number_format = '#,##0'
                cell.border = thin_border

                # Write breakdown values if present
                bd_items = mdata.get("breakdown", {})
                for (m, bkey), brow in breakdown_rows.items():
                    if m == metric and bkey in bd_items:
                        bd_val = to_float(bd_items[bkey])
                        if bd_val is not None:
                            if metric == "operating_margin" and bkey in operating_margin_pct_keys:
                                # These values are percentages; do not divide.
                                display_val = bd_val
                                number_format = '0.0%'
                            else:
                                display_val = bd_val / 1_000_000
                                number_format = '#,##0'
                        else:
                            display_val = None
                            number_format = '#,##0'
                        bd_cell = ws.cell(row=brow, column=year_col, value=display_val)
                        bd_cell.font = data_arial_italic_font
                        bd_cell.number_format = number_format

                        # Write CAGR values for this breakdown if available.
                        if metric in breakdown_source_metrics:
                            cagr_map = {
                                "gross_revenues": ("cagr_gross_revenues_breakdown_percent", "cagr_gross_revenues_{}_percent"),
                                "internal_costs": ("cagr_internal_costs_breakdown_percent", "cagr_internal_costs_{}_percent"),
                                "operating_margin": ("cagr_operating_margin_breakdown_percent", "cagr_operating_margin_{}_percent"),
                                "external_costs": ("cagr_external_costs_breakdown_percent", "cagr_external_costs_{}_percent")
                            }
                            base, pattern = cagr_map[metric]
                            cagr_key = pattern.format(bkey)
                            cagr_value = pchar.get(base, {}).get(cagr_key)
                            if cagr_value is not None:
                                cagr_cell = ws.cell(row=brow, column=3, value=cagr_value)
                                cagr_cell.font = Font(name="Arial", italic=True, size=8)
                                cagr_cell.number_format = '0.0%'
            else:
                # For metrics provided as plain numbers (e.g. investment_income, earnings, etc.)
                val = to_float(mdata)
                if val is not None and metric not in percent_metrics:
                    val = val / 1_000_000
                cell = ws.cell(row=metric_rows[metric], column=year_col, value=val)
                cell.fill = data_fill
                cell.font = Font(name="Arial", italic=True)
                cell.border = thin_border
                cell.number_format = '#,##0' if metric not in percent_metrics else '0.0%'

    # STEP 5: Write overall CAGR values (for the main metric rows).
    cagr_map = {
        "gross_revenues": "cagr_gross_revenues_percent",
        "investment_income": "cagr_investment_income_percent",
        "internal_costs": "cagr_internal_costs_percent",
        "operating_margin": "cagr_operating_margin_percent",
        "external_costs": "cagr_external_costs_percent",
        "earnings": "cagr_earnings_percent"
        # Additional metrics may be added here if applicable.
    }
    for metric, cagr_key in cagr_map.items():
        cagr_value = pchar.get(cagr_key)
        if cagr_value is not None and metric in metric_rows:
            row = metric_rows[metric]
            cagr_cell = ws.cell(row=row, column=3, value=cagr_value)
            cagr_cell.font = Font(name="Arial", italic=True, size=8)
            cagr_cell.number_format = '0.00%'

    # STEP 6: Compute and Write Percentages.
    # We assume that percentages for all breakdowns and top-level metrics are relative to gross revenues.
    base_row = metric_rows.get("gross_revenues")
    if base_row is None:
        # If there is no gross_revenues row, skip percentage calculations.
        return
    for i, year in enumerate(sorted_years):
        year_col = start_col_for_years + i * 2
        rev_val = ws.cell(row=base_row, column=year_col).value
        if rev_val is None or rev_val == 0:
            continue

        # For every breakdown row (across any metric that has a breakdown), compute its percentage of gross revenues.
        for (m, bkey), brow in breakdown_rows.items():
            metric_val = ws.cell(row=brow, column=year_col).value
            # For breakdowns that are already percentages (e.g. operating margin pct values), skip this calculation.
            if m == "operating_margin" and bkey in operating_margin_pct_keys:
                continue
            if isinstance(metric_val, (int, float)) and rev_val != 0:
                percent = (metric_val / rev_val) * 100
                pct_cell = ws.cell(row=brow, column=year_col + 1, value=f"{percent:.1f}%")
                pct_cell.font = Font(name="Arial", italic=True, size=8)

        # Calculate percentages for top-level metrics (compared to gross revenues)
        top_metrics = ["investment_income", "internal_costs", "operating_margin", "external_costs", "earnings"]
        for tm in top_metrics:
            if tm in metric_rows:
                tm_row = metric_rows[tm]
                metric_val = ws.cell(row=tm_row, column=year_col).value
                if isinstance(metric_val, (int, float)) and rev_val != 0:
                    percent = (metric_val / rev_val) * 100
                    pct_cell = ws.cell(row=tm_row, column=year_col + 1, value=f"{percent:.1f}%")
                    pct_cell.font = Font(name="Arial", italic=True, size=8)

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
    if not final_output.get("industry_comparison"):
        return
    
    wb = writer.book
    
    # If the sheet doesn't exist yet, create it
    if "Industry" not in wb.sheetnames:
        wb.create_sheet("Industry")
    ws = wb["Industry"]

    industry_data = final_output["industry_comparison"]
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
    Write the Segmentation sheet showing revenue breakdown by business segment over time.
    
    Parameters:
      writer: ExcelWriter object
      final_output: Dictionary containing the full output data including segmentation data.
    """
    # Early return if segmentation data is not present
    if not final_output.get("segmentation"):
        return

    reported_currency = final_output["summary"]["reported_currency"]
    segmentation_data = final_output["segmentation"]["data"]
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

    # Get sorted years and all unique segments from the "breakdown" dictionaries
    sorted_years = sorted(segmentation_data.keys(), key=lambda x: int(x))
    all_segments = set()
    for year in sorted_years:
        # For each year, get the breakdown sub-dictionary
        breakdown = segmentation_data[year].get("breakdown", {})
        all_segments.update(breakdown.keys())
    sorted_segments = sorted(all_segments)

    # Write year headers starting at row 3 (starting at column B)
    for i, year in enumerate(sorted_years):
        year_col = i + 2  # Column B is index 2
        y_cell = ws.cell(row=3, column=year_col, value=year)
        y_cell.fill = label_fill
        y_cell.font = label_font
        y_cell.border = thin_border
        y_cell.alignment = center_alignment

    # Write segment data for each segment (one segment per row)
    current_row = 4
    for segment in sorted_segments:
        # Write the segment name in column A
        segment_cell = ws.cell(row=current_row, column=1, value=segment)
        segment_cell.font = label_font
        segment_cell.fill = label_fill
        segment_cell.border = thin_border

        # Write values for each year
        for i, year in enumerate(sorted_years):
            year_col = i + 2  # Data starts at column B
            # Retrieve the breakdown for the given year and get the segment value
            breakdown = segmentation_data[year].get("breakdown", {})
            value = breakdown.get(segment)

            # Create the cell and apply the common styling
            value_cell = ws.cell(row=current_row, column=year_col)
            value_cell.fill = data_fill
            value_cell.font = data_arial_font
            value_cell.border = thin_border
            value_cell.alignment = right_alignment

            if value is not None:
                # Convert the value to millions and set the cell value/formatting
                value_cell.value = value / 1_000_000
                value_cell.number_format = '#,##0'

        current_row += 1

    # Calculate and write growth rates (CAGR) in the rightmost column.
    # The growth column is placed one column after the last year's column plus an extra spacer.
    growth_col = len(sorted_years) + 3
    growth_header = ws.cell(row=3, column=growth_col, value="CAGR")
    growth_header.fill = label_fill
    growth_header.font = label_font
    growth_header.border = thin_border
    growth_header.alignment = center_alignment

    # For each segment, calculate the compound annual growth rate (CAGR)
    for segment_idx, segment in enumerate(sorted_segments):
        row = segment_idx + 4

        # Get the first valid value for the segment (from the earliest year)
        first_val = None
        for year in sorted_years:
            breakdown = segmentation_data[year].get("breakdown", {})
            val = breakdown.get(segment)
            if val is not None:
                first_val = val
                break

        # Get the last valid value for the segment (from the most recent year)
        last_val = None
        for year in reversed(sorted_years):
            breakdown = segmentation_data[year].get("breakdown", {})
            val = breakdown.get(segment)
            if val is not None:
                last_val = val
                break

        if first_val is not None and last_val is not None and first_val != 0:
            years_between = len(sorted_years) - 1
            if years_between > 0:
                cagr = (last_val / first_val) ** (1 / years_between) - 1
                growth_cell = ws.cell(row=row, column=growth_col, value=cagr)
                growth_cell.number_format = '0.0%'
                growth_cell.font = Font(name="Arial", size=8, italic=True)
                growth_cell.fill = data_fill
                growth_cell.border = thin_border
        else:
            # Create an empty growth cell with consistent styling
            growth_cell = ws.cell(row=row, column=growth_col)
            growth_cell.fill = data_fill
            growth_cell.border = thin_border

    # Adjust column widths: set column A wider for segment names, and other columns to a fixed width.
    ws.column_dimensions['A'].width = 30  # Segment names
    for col in range(2, growth_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 12

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
    reported_currency = final_output["summary"]["reported_currency"]

    # Write and format the title
    title_cell = ws.cell(row=1, column=4, value=f"Historical Pricing Analysis ({reported_currency})")
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
        buy_cell.font = data_arial_bold_font
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
        sell_cell.font = data_arial_bold_font
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

def write_valuation_sheet(writer, final_output, ticker):
    """Write the valuation analysis sheet with 2x2 grid layout"""
    reported_currency = final_output["summary"]["reported_currency"]
    wb = writer.book
    
    if "Valuation" not in wb.sheetnames:
        wb.create_sheet("Valuation")
    ws = wb["Valuation"]

    # Title
    title_cell = ws.cell(row=1, column=4, value="Valuation (USD)")
    title_cell.fill = label_fill
    title_cell.font = title_font
    title_fill_range(ws, 1, 3, 5)
    apply_table_border(ws, 1, 3, 5)

    current_price = get_current_quote_yahoo(ticker)

    # Get forecast year column reference
    sorted_years = sorted(
        final_output["company_description"]["data"].keys(),
        key=lambda x: int(x)
    )
    first_forecast_col = (
        get_column_letter(2 + len(sorted_years)) if sorted_years else "B"
    )
    last_col_bal_sht = (
        get_column_letter(5 + len(sorted_years)) if sorted_years else "F"
    )
    # =========================================================================
    # REARRANGED SETTINGS
    # =========================================================================
    #   B3 => ADR Multiple
    #   B4 => Currency Ratio
    #   D3 => EPS Growth
    #   D4 => Dividend Growth
    #   F3 => Purchase Discount
    #   F4 => Sell Discount
    #   F5 => PE Multiple
    # =========================================================================
    settings = {
        (3, 2): ("ORD:ADR:", 1),                           # B3
        (4, 2): (f"USD:{reported_currency} rate:", 1),          # B4
        (3, 4): ("EPS growth rate:", 0.10),                     # D3
        (4, 4): ("Dividend growth rate:", 0.10),                # D4
        (3, 6): ("Purchase Discount:", 0.14),                   # F3
        (4, 6): ("Sell Discount:", 0.05),                       # F4
        (5, 6): ("PE Multiple:", 25),                           # F5
    }

    for (row, col), (label, value) in settings.items():
        label_cell = ws.cell(row=row, column=col - 1, value=label)
        label_cell.fill = label_fill
        label_cell.font = label_font
        label_cell.border = thin_border
        
        value_cell = ws.cell(row=row, column=col, value=value)
        value_cell.fill = data_fill
        value_cell.font = data_arial_font
        value_cell.border = thin_border
        value_cell.alignment = center_alignment

        # Format anything < 1 as percentage
        if isinstance(value, float) and value < 1:
            value_cell.number_format = '0.00%'
    
    new_settings = {
        (3, 7): ("Buy %:", 0.60),    # G3 (label), H3 (value = 60%)
        (4, 7): ("Sell %:", 1.20),   # G4 (label), H4 (value = 120%)
    }

    for (row, col), (label, value) in new_settings.items():
        label_cell = ws.cell(row=row, column=col, value=label)
        label_cell.fill = label_fill
        label_cell.font = label_font
        label_cell.border = thin_border
        
        value_cell = ws.cell(row=row, column=col + 1, value=value)
        value_cell.fill = data_fill
        value_cell.font = data_arial_font
        value_cell.border = thin_border
        value_cell.alignment = center_alignment

        # Format them as percentages
        value_cell.number_format = '0.00%'

    # If you need the numeric value of PE multiple in code:
    # pe_multiple_val = ws.cell(row=5, column=6).value  # F5

    # =========================================================================
    # 2x2 GRID LAYOUT
    # =========================================================================
    # We keep top-left and bottom-left in columns B/C.
    # We move top-right and bottom-right to columns E/F.
    # Then fix all formula references accordingly.
    # =========================================================================
    grid_segments = {
        # ---------------------------------------------------------------------
        # TOP LEFT (columns B/C)
        # ---------------------------------------------------------------------
        "Initial Rate of Investment:": {
            "start_row": 8,
            "start_col": 2,  # B
            "metrics": {
                "Current Price:": f"={current_price}",
                # Currency ratio is B4, ADR multiple is B3:
                "Current EPS:": f"='Co. Desc'!{first_forecast_col}5 * B4 * B3",
                # = B10 / B9 once written to the sheet
                "Initial ROI:": "=B10/B9",
            },
        },
        # ---------------------------------------------------------------------
        # TOP RIGHT (columns E/F)
        # ---------------------------------------------------------------------
        "Relative Value to Investment In T-Bonds:": {
            "start_row": 8,
            "start_col": 5,  # E
            "metrics": {
                # Same logic as top-left for 'Current EPS'
                "Current EPS:": f"='Co. Desc'!{first_forecast_col}5 * B4 * B3",
                "T-Bond Rate:": 0.04,
                # Was =H9/H10 in original; now =E9/E10
                "Relative Value:": "=E9/E10",
            },
        },
        # ---------------------------------------------------------------------
        # BOTTOM LEFT (columns B/C)
        # ---------------------------------------------------------------------
        "Valuation as an Equity Bond:": {
            "start_row": 15,
            "start_col": 2,  # B
            "metrics": {
                # Use B3 (ADR multiple) and B4 (currency ratio)
                "Current BV:": f"='Co. Desc'!{first_forecast_col}20 * B3 * B4",
                "Current ROE:": f"='Co. Desc'!{first_forecast_col}24",
                "Retained % adjustment:": 0.10,
                # unchanged, presumably references other sheet cells
                "Retained %:": "=1 - 'Analyses'!J5 - 'Analyses'!J7 - B18",
                "Net BV growth:": "=B17*B19",
                "BV in year 10:": "=FV(B20, 10, , -B16)",
                "EPS Adjustment Factor:": 1.5,
                "EPS in Year 10:": "=B17 * B21 * B22",
                # PE multiple is at F5
                "Value at PE Multiple:": "=F5 * B23",
                # Dividend growth is at D4 ⇒ use FV(D4,10,…)
                "Total Dividends:": (
                    f"=(('Co. Desc'!{first_forecast_col}13 + "
                    f"FV(D4, 10, , -'Co. Desc'!{first_forecast_col}13))/2)*10*B3*B4"
                ),
                "Total Future Value:": "=B24+B25",
                # Purchase discount is at F3
                "Purchase at Discount:": "=PV(F3, 10, , B26)*-1",
            },
        },
        # ---------------------------------------------------------------------
        # BOTTOM RIGHT (columns E/F)
        # ---------------------------------------------------------------------
        "Valuation on Earnings Growth:": {
            "start_row": 15,
            "start_col": 5,  # E
            "metrics": {
                # Same B3/B4 for ADR/currency
                "Current EPS:": f"='Co. Desc'!{first_forecast_col}5 * B4 * B3",
                # EPS Growth is at D3 => FV(D3,10,…)
                "EPS in year 10:": "=FV(D3, 10, , -E16)",
                "Avg PE Ratio:": f"=AVERAGE('Co. Desc'!B8:{first_forecast_col}8)",
                # Was =B5*H17 + B25; now =F5*E17 + B25
                "Value at PE Multiple:": "=F5*E17 + B25",
                # Was =RATE(10, , B9, -H19 + B25); now =RATE(10, , B9, -E19 + B25)
                "Price Return:": "=RATE(10, , B9, -E19 + B25)",
                "Dividend Return:": f"='Co. Desc'!{first_forecast_col}14",
                # Was =H20+H21; now =E20+E21
                "Total Return:": "=E20 + E21",
                # Purchase discount is F3
                "Purchase at Discount:": "=PV(F3, 10, , -E19)",
                # Sell discount is F4
                "Sell at Discount:": "=PV(F4, 10, , -E19)",
            },
        },
        "Float Valuation Approach:": {
            "start_row": 8,
            "start_col": 8,  # G
            "metrics": {
                "Float:": f"=('Balance Sht.'!{last_col_bal_sht}24 + 'Balance Sht.'!{last_col_bal_sht}23 -'Balance Sht.'!{last_col_bal_sht}13 -'Balance Sht.'!{last_col_bal_sht}11)*B4",
                
                "Float Growth Rate:": 0.05,
                
                "Cost of Float:": -0.04,
                
                "10yr FV:": "=FV(H10,10,,-H9)",
                
                "Investment Return:": 0.048,
                
                "Return on Float:": "=(H13-H11)",
                
                "Tax Burden on Float:": f"='Analyses'!{first_forecast_col}25 * H14",

                "After Tax Return on Float:": "=(H14-H15)",

                "Income on Float:": "=(H12 * H16)",

                "Discount Rate:": 0.08,

                "Capitalization Factor:": "=(H18-H10)",

                "Value of Float EOY10:": "=(H17/H19)",

                "RFR:": 0.06,

                "PV at RFR:": "=PV(H21,10,,-H20)",

                "Value of Insurance Equity:": f"='Co. Desc'!{first_forecast_col}19 * B4",

                "Total Value:": "=(H22+H23)",

                "Shares Outstanding:": f"='Co. Desc'!{first_forecast_col}16 * (1/B3)",
                
                # Share Value = Total Value / Shares Outstanding
                "Per Share Value:": "=(H24 / H25)",
                
                # Buy at => Share Value * Buy % (which is in H1)
                "Buy At:": "=H26 * $H$3",
                
                # Sell at => Share Value * Sell % (which is in H2)
                "Sell At:": "=H26 * $H$4",
            },
        }
    }

    #
    # Write each 2×2 grid segment
    #
    for title, config in grid_segments.items():
        start_row = config["start_row"]
        start_col = config["start_col"]
        
        # Segment title
        title_cell = ws.cell(row=start_row, column=start_col - 1, value=title)
        title_cell.fill = label_fill
        title_cell.font = Font(name="Times New Roman", size=12, bold=True, italic=True)
        title_cell.border = thin_border
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        # Merge the two columns for the label
        ws.merge_cells(
            start_row=start_row, start_column=start_col - 1, 
            end_row=start_row, end_column=start_col
        )
        
        # Write metrics
        current_row = start_row + 1
        for label, formula in config["metrics"].items():
            # Label
            label_cell = ws.cell(row=current_row, column=start_col - 1, value=label)
            label_cell.fill = label_fill
            label_cell.font = label_font
            label_cell.border = thin_border
            
            # Value/Formula
            value_cell = ws.cell(row=current_row, column=start_col, value=formula)
            value_cell.fill = data_fill
            value_cell.font = data_arial_font
            value_cell.border = thin_border
            value_cell.alignment = center_alignment

            # Make certain items bold
            if label.lower() in ["relative value:", "purchase at discount:", "sell at discount:", "buy at:", "sell at:"]:
                value_cell.font = data_arial_bold_font
            
            # Format numeric cells
            label_lower = label.lower()
            if label_lower == "t-bond rate:":
                value_cell.number_format = '0.00%'
            elif label_lower == "eps adjustment factor:":
                value_cell.number_format = '0.00'
            elif label_lower == "avg pe ratio:":
                value_cell.number_format = '0.00'
            elif label_lower == "rfr:":
                value_cell.number_format = '0.00%'
            elif label_lower == "float:":
                value_cell.number_format = '"$"#,##0.00'
            elif "net bv growth" in label_lower:
                value_cell.number_format = '0.00%'
            elif "shares outstanding" in label_lower:
                value_cell.number_format = '#,##0'
            elif any(x in label_lower for x in ["rate", "tax", "factor", "roi", "return", "roe", "%", "cost"]):
                value_cell.number_format = '0.00%'
            elif any(x in label_lower for x in ["price", "value", "eps", "bv", "dividends", "purchase", "sell",
                                                "ebit", "value", "income", "debt", "share", "pv", "fv", "buy at", "sell at"]):
                value_cell.number_format = '"$"#,##0.00'
            
            current_row += 1

    # =========================================================================
    # Column Widths
    # =========================================================================
    # Example adjustments for label columns vs. value columns
    ws.column_dimensions[get_column_letter(1)].width = 25  # A
    ws.column_dimensions[get_column_letter(2)].width = 15  # B
    ws.column_dimensions[get_column_letter(3)].width = 25  # C
    ws.column_dimensions[get_column_letter(4)].width = 20  # D
    ws.column_dimensions[get_column_letter(5)].width = 25  # E
    ws.column_dimensions[get_column_letter(6)].width = 15  # F
    ws.column_dimensions[get_column_letter(7)].width = 25  # G
    ws.column_dimensions[get_column_letter(8)].width = 15  # H


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
    write_hist_pricing_sheet(writer, final_output)
    write_valuation_sheet(writer, final_output, ticker)
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
    write_profit_desc_sheet(writer, final_output)
    write_qualities_sheet(writer, final_output)
    write_industry_sheet(writer, final_output)
    write_segmentation_sheet(writer, final_output)
    write_hist_pricing_sheet(writer, final_output)
    write_valuation_sheet(writer, final_output, ticker)
    generate_config_note(ticker, writer.book)

    # Apply formatting: set font to Arial size 10 for non-formatted cells and remove gridlines
    format_workbook(writer)

    # Save changes
    writer.close()
    print(f"Data for {ticker} written to {xls_filename} successfully.")
