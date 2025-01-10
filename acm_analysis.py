# acm_analysis.py

import argparse
import json
import os
import sys
import requests
import datetime
from forum_post_summary import generate_post_summary
from forum_posts import fetch_all_for_ticker
from dotenv import load_dotenv
from gen_excel import generate_excel_for_ticker_year
from outlook_ticker_search import filter_emails_by_ticker
from industry_comp import get_industry_peers_with_stats
from utils import get_company_profile, get_current_market_cap_yahoo, get_current_quote_yahoo, get_reported_currency, get_yahoo_ticker, get_yearly_high_low_yahoo

# Load the .env file
load_dotenv()

# Retrieve FMP_API_KEY from environment
FMP_API_KEY = os.getenv("FMP_API_KEY")
if not FMP_API_KEY:
    print("ERROR: FMP_API_KEY not found in environment. Please set FMP_API_KEY in your .env file.")
    sys.exit(1)

def save_to_file(data: dict, filename: str):
    """Save the fetched data to a JSON file."""
    try:
        # Ensure the output directory exists
        os.makedirs("output", exist_ok=True)
        filepath = os.path.join("output", filename)
        with open(filepath, "w") as f:
            json.dump(data, f, indent=4)
        print(f"Data saved to {filepath}")
    except Exception as e:
        print(f"Error saving data to file: {e}")

def get_fiscal_year_end(symbol: str) -> str:
    """
    Fetch the fiscalYearEnd from the company-core-information endpoint.
    """
    url = f"https://financialmodelingprep.com/api/v4/company-core-information?symbol={symbol}&apikey={FMP_API_KEY}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        if isinstance(data, list) and len(data) > 0:
            return data[0].get("fiscalYearEnd")
        return None
    except Exception as e:
        print(f"Error fetching fiscalYearEnd for {symbol}: {e}")
        return None
    
def derive_most_recent_fiscal_year(fiscal_year_end: str) -> int:
    """
    Determine the most recent completed fiscal year given the company's fiscal year end (MM-DD)
    and today's date. For example, if the fiscal year end is 09-30 and today's date is 2024-01-15,
    the most recent completed year might be 2023 if the date is before 09-30, or 2024 if the date
    is after 09-30 in the same year.

    Returns: An integer representing the last completed fiscal year, or None if we can't parse.
    """
    if not fiscal_year_end:
        return None  # Can't parse

    try:
        # Extract month and day from e.g. "09-30"
        month_str, day_str = fiscal_year_end.split("-")
        fye_month = int(month_str)
        fye_day = int(day_str)
    except Exception:
        return None

    today = datetime.date.today()
    current_year = today.year
    # Construct a date object for the current year's FYE
    fye_this_year = datetime.date(current_year, fye_month, fye_day)

    # If today's date is *after or equal* to the FYE date, then the most recent complete
    # fiscal year is the current year. Otherwise, it is the previous year.
    if today >= fye_this_year:
        return current_year
    else:
        return current_year - 1

def get_financials(symbol: str, statement_type: str, frequency: str):
    """
    Fetch financial data from Financial Modeling Prep and save it to a file.
    statement_type can be one of: "bs" (Balance Sheet), "ic" (Income Statement), "cf" (Cash Flow)
    frequency is "annual" or "quarter"
    """
    endpoints = {
        "ic": "income-statement",
        "bs": "balance-sheet-statement",
        "cf": "cash-flow-statement",
        "bs-ar": "balance-sheet-statement-as-reported"
    }
    endpoint = endpoints.get(statement_type)
    if not endpoint:
        print(f"Invalid statement type: {statement_type}")
        return

    url = f"https://financialmodelingprep.com/api/v3/{endpoint}/{symbol}?period={frequency}&apikey={FMP_API_KEY}"

    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()

        if isinstance(data, list) and len(data) > 0:
            # Wrap the list into {"financials": data} to keep compatibility with the extraction logic
            wrapped_data = {"financials": data}
            filename = f"{symbol}_{statement_type}_{frequency}.json"
            save_to_file(wrapped_data, filename)
        else:
            print(f"No data returned for {symbol}, {statement_type}, {frequency}")
    except Exception as e:
        print(f"Error fetching data: {e}")

def get_basic_financials(symbol: str):
    """
    Fetch key metrics from FMP as a stand-in for basic financials.
    This will return an array of yearly metrics. We'll store them under "keyMetrics".
    """
    url = f"https://financialmodelingprep.com/api/v3/key-metrics/{symbol}?apikey={FMP_API_KEY}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()

        # data is a list of metrics by year descending
        # We'll store them as {"keyMetrics": data}
        filename = f"{symbol}_basic_financials.json"
        to_store = {"keyMetrics": data}
        save_to_file(to_store, filename)
    except Exception as e:
        print(f"Error fetching basic financials for {symbol}: {e}")

def get_revenue_segmentation(symbol: str):
    """
    Fetch revenue product segmentation data from FMP and organize it by year.

    Returns:
        A dictionary where each key is a year (int) and the value is another dictionary
        mapping revenue segments to their respective amounts.
    """
    url = f"https://financialmodelingprep.com/api/v4/revenue-product-segmentation?symbol={symbol}&structure=flat&period=annual&apikey={FMP_API_KEY}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()

        # Example provided structure:
        # [
        #   {
        #     "2024-09-28": {
        #       "Mac": 29984000000,
        #       "Service": 96169000000,
        #       "Wearables, Home and Accessories": 37005000000,
        #       "iPad": 26694000000,
        #       "iPhone": 201183000000
        #     }
        #   },
        #   // ... more entries
        # ]

        segmentation = {}
        for entry in data:
            for date_str, segments in entry.items():
                # Extract the year from the date string
                year_str = date_str.split('-')[0]
                try:
                    year = int(year_str)
                except ValueError:
                    print(f"Invalid date format: {date_str}")
                    continue

                # Ensure segments is a dictionary
                if isinstance(segments, dict):
                    segmentation[year] = segments
                else:
                    print(f"Invalid segments format for year {year}: {segments}")

        return segmentation
    except Exception as e:
        print(f"Error fetching revenue segmentation for {symbol}: {e}")
        return {}


def load_json(filename: str) -> dict:
    """Load JSON file into a dictionary."""
    filepath = os.path.join("output", filename)
    if not os.path.exists(filepath):
        return {}
    with open(filepath, "r") as f:
        return json.load(f)

def extract_series_values_by_year(basic_data: dict, key: str) -> dict:
    """
    Adapted for FMP's key-metrics data.
    We'll look for the key in each year's data under `keyMetrics`.
    Each entry in keyMetrics has a "date" field and various metrics.
    We'll try to find a metric matching 'pe ratio' for example.
    
    For P/E ratio, the key is often "peRatio" or "PE ratio" in the key-metrics data.
    Check the returned data fields from key-metrics for the exact naming.
    
    Example key metrics structure (truncated):
    [
      {
        "date": "2023-06-30",
        "symbol": "AAPL",
        "peRatio": 37.3678,
        ...
      },
      {
        "date": "2022-06-30",
        "peRatio": 28.5012,
        ...
      },
      ...
    ]

    We'll return {2023: 37.3678, 2022: 28.5012, ...}
    """
    # According to FMP docs, PE ratio field is often "peRatio"
    # If you need a different metric, adjust the field name here.
    field_map = {
        "pe": "peRatio"
    }
    fmp_field = field_map.get(key, key)

    result = {}
    key_metrics = basic_data.get("keyMetrics", [])
    for entry in key_metrics:
        date_str = entry.get("date")
        value = entry.get(fmp_field)
        if date_str and value is not None:
            year_str = date_str.split("-")[0]
            try:
                year_int = int(year_str)
                result[year_int] = value
            except ValueError:
                pass

    return result

def extract_yoy_data(symbol: str, years: list, segmentation_data: dict, profile: dict):
    # Load the three financial statements
    bs_data = load_json(f"{symbol}_bs_annual.json")
    ic_data = load_json(f"{symbol}_ic_annual.json")
    cf_data = load_json(f"{symbol}_cf_annual.json")
    basic_data = load_json(f"{symbol}_basic_financials.json")

    # Extract P/E ratios by year from key metrics
    pe_by_year = extract_series_values_by_year(basic_data, 'pe')
    
    # The data is wrapped as {"financials": [...]}, each entry has a 'date'
    def by_year_dict(data):
        res = {}
        for item in data.get('financials', []):
            date_str = item.get('date')
            if date_str:
                y = date_str.split('-')[0]
                try:
                    year_int = int(y)
                    res[year_int] = item
                except:
                    pass
        return res

    bs_by_year = by_year_dict(bs_data)
    ic_by_year = by_year_dict(ic_data)
    cf_by_year = by_year_dict(cf_data)

    results = {}
    prev_shares_outstanding = None  # For buyback calculation
    yahoo_symbol = get_yahoo_ticker(profile)

    for year in years:
        bs = bs_by_year.get(year, {})
        ic = ic_by_year.get(year, {})
        cf = cf_by_year.get(year, {})

        # Extract relevant financial data
        shares_outstanding = ic.get('weightedAverageShsOut')
        net_profit = ic.get('netIncome')
        revenues = ic.get('revenue')
        diluted_eps = ic.get('epsdiluted')
        ebit = ic.get('operatingIncome')
        dividends_paid = -1 * cf.get('dividendsPaid', 0)
        shareholder_equity = bs.get('totalStockholdersEquity')
        long_term_debt =  (bs.get("longTermDebt") or 0) + (bs.get("shortTermDebt") or 0) - (bs.get("capitalLeaseObligations") or 0)
        provision_for_taxes = ic.get('incomeTaxExpense')
        pretax_income = ic.get('incomeBeforeTax')
        depreciation = cf.get('depreciationAndAmortization')
        cost_of_revenue = ic.get('costOfRevenue') or 0
        cost_of_res_and_dev = ic.get('researchAndDevelopmentExpenses') or 0
        cost_of_selling_and_marketing_gen_and_admin = ic.get('sellingGeneralAndAdministrativeExpenses') or 0
        expenses = cost_of_selling_and_marketing_gen_and_admin + cost_of_res_and_dev + cost_of_revenue
        amort_dep = ic.get('depreciationAndAmortization')
        ebitda = revenues - expenses + amort_dep if revenues and expenses and amort_dep else None
        capex = cf.get('capitalExpenditure')
        fcf = ebitda + capex if ebitda and capex else None

        # Operating margin = ebit / revenue
        operating_margin = (ebit / revenues) if (ebit and revenues) else None

        # Operating EPS = Net Profit / Shares Outstanding
        operating_eps = (net_profit / shares_outstanding) if (net_profit and shares_outstanding) else None

        # Operating earnings = revenues - expenses
        operating_earnings = (revenues - expenses) if (expenses is not None and revenues is not None) else None

        # Operating earnings percentage = operating earnings / revenues
        operating_earnings_pct = (operating_earnings / revenues) if (operating_earnings is not None and revenues) else None

        income_tax_expense = ic.get('incomeTaxExpense')
        interest_and_other_income_expense = -1 * ic.get('totalOtherIncomeExpensesNet')
        extern_costs = income_tax_expense + interest_and_other_income_expense if (income_tax_expense and interest_and_other_income_expense) else None

        earnings = operating_earnings - extern_costs if (operating_earnings and extern_costs) else None
        stmt_cf_share_repurchase = cf.get('commonStockRepurchased')
        net_acquisitions = cf.get('acquisitionsNet')

        earnings_pct_revenue = (earnings / revenues) if (earnings and revenues) else None

        dividends_paid_pct_fcf = (dividends_paid / fcf) if (dividends_paid and fcf) else None

        # Tax rate = provision_for_taxes / pretax_income
        tax_rate = (provision_for_taxes / pretax_income) if (provision_for_taxes and pretax_income) else None

        # Yearly high/low prices
        
        yearly_high, yearly_low = get_yearly_high_low_yahoo(yahoo_symbol, year)
        average_price = None
        if yearly_high is not None and yearly_low is not None:
            average_price = (yearly_high + yearly_low) / 2

        # P/E Ratio from basic financials key metrics
        pe_ratio = pe_by_year.get(year)

        # Book Value per share = shareholder_equity / shares_outstanding
        book_value_per_share = (shareholder_equity / shares_outstanding) if (shareholder_equity and shares_outstanding and shares_outstanding != 0) else None

        # Dividends per share = dividends_paid / shares_outstanding
        dividends_per_share = (dividends_paid / shares_outstanding) if (dividends_paid and shares_outstanding and shares_outstanding != 0) else None

        # Avg dividend yield = dividends_per_share / average_price
        avg_dividend_yield = (dividends_per_share / average_price) if (dividends_per_share and average_price and average_price != 0) else None

        # Buyback = (change in shares_outstanding) * average_price
        buyback = None
        if prev_shares_outstanding is not None and shares_outstanding is not None and average_price:
            share_change = prev_shares_outstanding - shares_outstanding
            # Negative share_change implies buyback
            buyback = share_change * average_price

        # ROE = net_profit / shareholder_equity
        roe = (net_profit / shareholder_equity) if (net_profit and shareholder_equity and shareholder_equity != 0) else None

        # ROC = net_profit / (shareholder_equity + long_term_debt)
        roc = None
        capital = (shareholder_equity + long_term_debt) if (shareholder_equity and long_term_debt) else None
        if net_profit is not None and capital and capital != 0:
            roc = net_profit / capital

        # Sales per share = revenues / shares_outstanding
        sales_per_share = (revenues / shares_outstanding) if (revenues and shares_outstanding and shares_outstanding != 0) else None

        # Depreciation % = depreciation / net_profit
        depreciation_percent = (depreciation / net_profit) if (depreciation and net_profit and net_profit != 0) else None

        # p/e low high
        pe_low = yearly_low / diluted_eps if diluted_eps !=0 else None
        pe_high = yearly_high / diluted_eps if diluted_eps != 0  else None

        # p/b low high
        pb_low = yearly_low / book_value_per_share if book_value_per_share != 0 else None
        pb_high = yearly_high / book_value_per_share if book_value_per_share != 0 else None

        # p/s low high
        ps_low = yearly_low / sales_per_share if sales_per_share !=0 else None
        ps_high = yearly_high / sales_per_share if sales_per_share != 0 else None

        # pcf low high
        addback_dep_earnings_ps = (net_profit + depreciation) / shares_outstanding if shares_outstanding != 0 else None

        pcfs_low = yearly_low / addback_dep_earnings_ps if addback_dep_earnings_ps != 0 else None
        pcfs_high = yearly_high / addback_dep_earnings_ps if addback_dep_earnings_ps != 0 else None

        year_segments = segmentation_data.get(str(year), {})

        # Prepare the three sections
        company_description = {
            "net_profit": net_profit,
            "diluted_eps": diluted_eps,
            "operating_eps": operating_eps,
            "pe_ratio": pe_ratio,
            "price_low": yearly_low,
            "price_high": yearly_high,
            "dividends_paid": dividends_paid,
            "dividends_per_share": dividends_per_share,
            "avg_dividend_yield": avg_dividend_yield,
            "shares_outstanding": shares_outstanding,
            "buyback": buyback,
            "share_equity": shareholder_equity,
            "book_value_per_share": book_value_per_share,
            "long_term_debt": long_term_debt,
            "roe": roe,
            "roc": roc
        }

        analysis = {
            "revenues": revenues,
            "sales_per_share": sales_per_share,
            "op_margin_percent": operating_margin,
            "tax_rate": tax_rate,
            "depreciation": depreciation,
            "depreciation_percent": depreciation_percent
        }

        # Structure the revenues with breakdown
        revenues_structured = {
            "total_revenues": revenues,
            "breakdown": year_segments.get("revenue", {})
        }

        profit_description = {
            "revenues": revenues_structured,
            "expenses": {
                "total_expenses": expenses,
                "breakdown": {
                    "cost_of_revenue": cost_of_revenue,
                    "research_and_development": cost_of_res_and_dev,
                    "selling_marketing_general_admin": cost_of_selling_and_marketing_gen_and_admin
                }
            },
            "ebitda": ebitda,
            "amortization_depreciation": amort_dep,  
            "free_cash_flow": fcf,
            "capex": capex,
            "operating_earnings": {
                "total_operating_earnings": operating_earnings,
                "breakdown": year_segments.get("operating_income", {})
            },
            "operating_earnings_percent_revenue": operating_earnings_pct,
            "external_costs": {
                "total_external_costs": extern_costs,
                "breakdown": {
                    "income_taxes": income_tax_expense,
                    "interest_and_other_income": interest_and_other_income_expense
                }
            },
            "earnings": earnings,
            "earnings_percent_revenue": earnings_pct_revenue,
            "dividend_paid": dividends_paid,
            "dividend_paid_pct_fcf": dividends_paid_pct_fcf,
            "share_buybacks_from_stmt_cf": stmt_cf_share_repurchase,
            "net_biz_acquisition": net_acquisitions
        }

        # Directly extract and structure the balance sheet data from 'bs'
        balance_sheet = {
            "assets": {"total_assets": bs.get("totalAssets"), 
                       "breakdown": {
                            "cash_and_cash_equivalents": bs.get("cashAndCashEquivalents"),
                            "short_term_investment": bs.get("shortTermInvestments"),
                            "accounts_receivable_net": bs.get("netReceivables"),
                            "other_current_assets": bs.get("otherCurrentAssets"),
                            "land_property_equipment_net": bs.get("propertyPlantEquipmentNet"),
                            "goodwill_and_intangible_assets": bs.get("goodwillAndIntangibleAssets"),
                            "other_non_current": bs.get("otherNonCurrentAssets"),
                            "long_term_equity_investment": bs.get("longTermInvestments")  # Adjust if a better field exists
                        }
            },
            "liabilities": {"total_liabilities": bs.get("totalLiabilities"),
                            "breakdown": {
                                # current liabilities
                                "accounts_payable": bs.get("accountPayables"),
                                "tax_payables": bs.get("taxPayables"),
                                "other_current_liabilities": bs.get("otherCurrentLiabilities"),
                                "deferred_revenue": bs.get("deferredRevenue"),
                                "short_term_debt": bs.get("shortTermDebt"),
                                # non current liabilites
                                "long_term_debt_minus_capital_lease_obligation": (bs.get("longTermDebt") or 0) - (bs.get("capitalLeaseObligations") or 0),
                                "other_non_current_liabilities": bs.get("otherNonCurrentLiabilities"),
                                "capital_lease_obligations": bs.get("capitalLeaseObligations") 
                            }
            },
            "shareholders_equity": {"total_shareholders_equity": bs.get("totalEquity"),
                                    "breakdown": {
                                        "common_stock": bs.get("commonStock"),
                                        "additional_paid_in_capital": bs.get("othertotalStockholdersEquity"),  # Adjust if separate field exists
                                        "retained_earnings": bs.get("retainedEarnings"),
                                        "accumulated_other_comprehensive_income_loss": bs.get("accumulatedOtherComprehensiveIncomeLoss")
                                    }
            }
        }

        hist_pricing = {
            "pe_low": pe_low,
            "pe_high": pe_high,
            "pb_low": pb_low,
            "pb_high": pb_high,
            "ps_low": ps_low,
            "ps_high": ps_high,
            "pcf_low": pcfs_low,
            "pcf_high": pcfs_high
        }

        # Populate the results dictionary with the four sections
        results[year] = {
            "year": year,
            "company_description": company_description,
            "analyses": analysis,
            "profit_description": profit_description,
            "balance_sheet": balance_sheet,
            "hist_pricing": hist_pricing,
            "segmentation": year_segments.get("segmentation", {})
        }

        prev_shares_outstanding = shares_outstanding

    return results

def calculate_cagr(values_by_year: list):
    """
    Calculate the Compound Annual Growth Rate (CAGR) given a list of (year, value) tuples.
    The list should be sorted by year in ascending order.
    We will use the first and last years as endpoints, but if the initial value is <= 0, 
    we will move forward year-over-year until we find a positive start value.
    Print to screen if the start was adjusted.
    
    Returns:
    - CAGR as a percentage (float), or None if not calculable.
    """
    if len(values_by_year) < 2:
        return None

    # Extract years and values
    years = [y for (y, v) in values_by_year]
    vals = [v for (y, v) in values_by_year]

    # initial assumption: begin_value is vals[0], end_value is vals[-1]
    begin_value = vals[0]
    end_value = vals[-1]
    periods = years[-1] - years[0]

    if periods <= 0:
        return None

    # If begin_value <= 0, move forward until we find a positive start
    idx = 0
    adjusted = False
    while idx < len(vals) and (vals[idx] is None or vals[idx] <= 0):
        idx += 1

    if idx >= len(vals):
        # No positive start found
        print("No positive start value found for CAGR calculation.")
        return None

    if idx > 0:
        # We found a later start
        begin_value = vals[idx]
        # End stays the last one
        end_value = vals[-1]
        periods = years[-1] - years[idx]
        adjusted = True

    if adjusted:
        print(f"Start was adjusted for CAGR calculation from year {years[0]} to {years[idx]} due to non-positive start value.")

    if begin_value <= 0 or end_value <= 0 or periods <= 0:
        return None

    try:
        cagr = ((end_value / begin_value) ** (1 / periods) - 1)
        return round(cagr, 4)
    except:
        return None

def compute_investment_characteristics(yoy_data: dict):
    """
    Compute investment characteristics based on YOY data using the new calculate_cagr function.
    Instead of directly passing begin and end values, we build a list of (year, value) tuples
    and pass them to calculate_cagr for each CAGR calculation. If the initial value is non-positive,
    calculate_cagr will attempt to adjust the start.
    """
    investment_characteristics = {
        "earnings_analysis": {
            "growth_rate_percent_operating_eps": None,
            "quality_percent": None
        },
        "use_of_earnings_analysis": {
            "avg_dividend_payout_percent": None,
            "avg_stock_buyback_percent": None
        },
        "sales_analysis": {
            "growth_rate_percent_revenues": None,
            "growth_rate_percent_sales_per_share": None
        },
        "sales_analysis_last_5_years": {
            "growth_rate_percent_revenues": None,
            "growth_rate_percent_sales_per_share": None
        }
    }

    sorted_years = sorted(yoy_data.keys())

    def build_values_by_year(metric_key_chain):
        # metric_key_chain is something like ("company_description", "operating_eps")
        # We'll return a list of (year, value) for all years that have a non-None value.
        vals = []
        for year in sorted_years:
            d = yoy_data[year]
            for k in metric_key_chain:
                d = d.get(k, {})
            # After traversing keys, d should be the value or {}
            if not isinstance(d, dict):
                val = d
            else:
                val = None
            if val is not None:
                vals.append((year, val))
        return vals

    # Earnings Analysis: Growth Rate % (of Operating EPS)
    # Previously used begin_net_profit/end_net_profit, now we build a values list:
    ops_eps_values = build_values_by_year(("company_description", "operating_eps"))
    if len(ops_eps_values) >= 2:
        cagr_ops = calculate_cagr(ops_eps_values)
        investment_characteristics["earnings_analysis"]["growth_rate_percent_operating_eps"] = cagr_ops

    # Earnings Analysis: Quality % (unchanged logic)
    total_eps = 0
    total_operating_eps = 0
    count_eps = 0
    count_operating_eps = 0
    for year in sorted_years:
        eps = yoy_data[year]["company_description"].get("diluted_eps")
        operating_eps = yoy_data[year]["company_description"].get("operating_eps")
        if eps is not None:
            total_eps += eps
            count_eps += 1
        if operating_eps is not None:
            total_operating_eps += operating_eps
            count_operating_eps += 1
    avg_eps = total_eps / count_eps if count_eps > 0 else None
    avg_operating_eps = total_operating_eps / count_operating_eps if count_operating_eps > 0 else None
    if avg_eps and avg_operating_eps and avg_operating_eps != 0:
        quality_percent = (avg_eps / avg_operating_eps)
        investment_characteristics["earnings_analysis"]["quality_percent"] = round(quality_percent, 2)

    # Avg Dividend Payout % (unchanged logic)
    sum_dividends_per_share = 0
    sum_operating_eps_for_div = 0
    count_for_div_payout = 0
    for year in sorted_years:
        dividends_per_share = yoy_data[year]["company_description"].get("dividends_per_share")
        operating_eps = yoy_data[year]["company_description"].get("operating_eps")
        if dividends_per_share is not None and operating_eps is not None and operating_eps != 0:
            sum_dividends_per_share += dividends_per_share
            sum_operating_eps_for_div += operating_eps
            count_for_div_payout += 1
    avg_dividend_payout_percent = None
    if count_for_div_payout > 0:
        average_dividends_per_share = sum_dividends_per_share / count_for_div_payout
        average_operating_eps = sum_operating_eps_for_div / count_for_div_payout
        if average_operating_eps != 0:
            avg_dividend_payout_percent = (average_dividends_per_share / average_operating_eps)
            avg_dividend_payout_percent = round(avg_dividend_payout_percent, 2)
    investment_characteristics["use_of_earnings_analysis"]["avg_dividend_payout_percent"] = avg_dividend_payout_percent

    # Avg Stock Buyback % (unchanged logic)
    sum_of_buybacks = 0
    sum_of_net_profits = 0
    for year in sorted_years:
        buyback = yoy_data[year]["company_description"].get("buyback")
        net_profit = yoy_data[year]["company_description"].get("net_profit")
        if buyback is not None and net_profit and net_profit != 0:
            sum_of_buybacks += buyback
            sum_of_net_profits += net_profit
    avg_stock_buyback_percent = None
    if sum_of_net_profits != 0:
        avg_stock_buyback_percent = (sum_of_buybacks / sum_of_net_profits)
        avg_stock_buyback_percent = round(avg_stock_buyback_percent, 2)
    investment_characteristics["use_of_earnings_analysis"]["avg_stock_buyback_percent"] = avg_stock_buyback_percent

    # Sales Analysis: Growth Rates
    # Revenues
    rev_values = []
    for year in sorted_years:
        val = yoy_data[year]["analyses"].get("revenues")
        if val is not None:
            rev_values.append((year, val))
    if len(rev_values) >= 2:
        investment_characteristics["sales_analysis"]["growth_rate_percent_revenues"] = calculate_cagr(rev_values)

    # Sales per Share
    sps_values = []
    for year in sorted_years:
        val = yoy_data[year]["analyses"].get("sales_per_share")
        if val is not None:
            sps_values.append((year, val))
    if len(sps_values) >= 2:
        investment_characteristics["sales_analysis"]["growth_rate_percent_sales_per_share"] = calculate_cagr(sps_values)

    # Last 5 Years: Revenues
    last_5_years = sorted_years[-5:]
    rev_values_5y = []
    for y in last_5_years:
        val = yoy_data[y]["analyses"].get("revenues")
        if val is not None:
            rev_values_5y.append((y, val))
    if len(rev_values_5y) >= 2:
        investment_characteristics["sales_analysis_last_5_years"]["growth_rate_percent_revenues"] = calculate_cagr(rev_values_5y)

    # Last 5 Years: Sales per Share
    sps_values_5y = []
    for y in last_5_years:
        val = yoy_data[y]["analyses"].get("sales_per_share")
        if val is not None:
            sps_values_5y.append((y, val))
    if len(sps_values_5y) >= 2:
        investment_characteristics["sales_analysis_last_5_years"]["growth_rate_percent_sales_per_share"] = calculate_cagr(sps_values_5y)

    return investment_characteristics

def compute_balance_sheet_characteristics(yoy_data: dict):
    """
    Compute CAGR for Total Assets, Total Liabilities, and Total Shareholders' Equity using the new calculate_cagr.
    Build values_by_year lists for each metric.
    """
    balance_sheet_characteristics = {
        "cagr_total_assets_percent": None,
        "cagr_total_liabilities_percent": None,
        "cagr_total_shareholders_equity_percent": None
    }

    sorted_years = sorted(yoy_data.keys())
    if len(sorted_years) < 2:
        print("Not enough data to compute CAGR for balance sheet characteristics.")
        return balance_sheet_characteristics

    def build_bs_values(metric_key_chain):
        vals = []
        for year in sorted_years:
            d = yoy_data[year]["balance_sheet"]
            for k in metric_key_chain:
                d = d.get(k, {})
            # d should be value or {}
            if not isinstance(d, dict):
                val = d
            else:
                val = None
            if val is not None:
                vals.append((year, val))
        return vals

    # Total Assets
    assets_values = build_bs_values(("assets", "total_assets"))
    if len(assets_values) >= 2:
        balance_sheet_characteristics["cagr_total_assets_percent"] = calculate_cagr(assets_values)

    # Total Liabilities
    liab_values = build_bs_values(("liabilities", "total_liabilities"))
    if len(liab_values) >= 2:
        balance_sheet_characteristics["cagr_total_liabilities_percent"] = calculate_cagr(liab_values)

    # Total Shareholders' Equity
    eq_values = build_bs_values(("shareholders_equity", "total_shareholders_equity"))
    if len(eq_values) >= 2:
        balance_sheet_characteristics["cagr_total_shareholders_equity_percent"] = calculate_cagr(eq_values)

    return balance_sheet_characteristics

def compute_profit_description_characteristics(yoy_data: dict):
    """
    Compute the CAGR of various profit-related metrics using the new calculate_cagr function.
    Includes CAGR calculations for each item in external_costs and revenues breakdowns.
    """
    sorted_years = sorted(yoy_data.keys())
    if len(sorted_years) < 2:
        return {
            "cagr_revenues_percent": None,
            "cagr_total_expenses_percent": None,
            "cagr_ebitda_percent": None,
            "cagr_free_cash_flow_percent": None,
            "cagr_operating_earnings_percent": None,
            "cagr_total_external_costs_percent": None,
            "cagr_earnings_percent": None,
            "cagr_cost_of_revenue_percent": None,
            "cagr_research_and_development_percent": None,
            "cagr_selling_marketing_general_admin_percent": None,
            # Initialize empty dicts for dynamic breakdown items
            "cagr_external_costs_breakdown_percent": {},
            "cagr_operating_earnings_breakdown_percent": {},
            "cagr_revenues_breakdown_percent": {}
        }

    results = {}

    # Initialize nested dictionaries before assignment
    results["cagr_external_costs_breakdown_percent"] = {}
    results["cagr_revenues_breakdown_percent"] = {}
    results["cagr_operating_earnings_breakdown_percent"] = {}

    def get_val(year, key_chain):
        d = yoy_data[year]
        for k in key_chain:
            d = d.get(k, {})
        if not isinstance(d, dict):
            return d
        return None

    def build_values(key_chain):
        vals = []
        for y in sorted_years:
            val = get_val(y, key_chain)
            if val is not None:
                # Ensure the value is numeric for CAGR calculation
                try:
                    numeric_val = float(val)
                    vals.append((y, numeric_val))
                except ValueError:
                    pass
        return vals

    metrics = {
        "cagr_revenues_percent": ("profit_description", "revenues", "total_revenues"),
        "cagr_total_expenses_percent": ("profit_description", "expenses", "total_expenses"),
        "cagr_ebitda_percent": ("profit_description", "ebitda"),
        "cagr_free_cash_flow_percent": ("profit_description", "free_cash_flow"),
        "cagr_operating_earnings_percent": ("profit_description", "operating_earnings", "total_operating_earnings"),
        "cagr_total_external_costs_percent": ("profit_description", "external_costs", "total_external_costs"),
        "cagr_earnings_percent": ("profit_description", "earnings"),
        "cagr_cost_of_revenue_percent": ("profit_description", "expenses", "breakdown", "cost_of_revenue"),
        "cagr_research_and_development_percent": ("profit_description", "expenses", "breakdown", "research_and_development"),
        "cagr_selling_marketing_general_admin_percent": ("profit_description", "expenses", "breakdown", "selling_marketing_general_admin"),
    }

    # Calculate CAGR for predefined metrics
    for metric_name, chain in metrics.items():
        vals = build_values(chain)
        if len(vals) >= 2:
            results[metric_name] = calculate_cagr(vals)
        else:
            results[metric_name] = None

    # Compute CAGR for each external_costs breakdown item
    external_costs_keys = set()
    for year in sorted_years:
        breakdown = yoy_data[year].get("profit_description", {}).get("external_costs", {}).get("breakdown", {})
        external_costs_keys.update(breakdown.keys())

    for cost_item in external_costs_keys:
        key_chain = ("profit_description", "external_costs", "breakdown", cost_item)
        vals = build_values(key_chain)
        if len(vals) >= 2:
            cagr_key = f"cagr_external_costs_{cost_item}_percent"
            results["cagr_external_costs_breakdown_percent"][cagr_key] = calculate_cagr(vals)
        else:
            cagr_key = f"cagr_external_costs_{cost_item}_percent"
            results["cagr_external_costs_breakdown_percent"][cagr_key] = None

    # Compute CAGR for each revenues breakdown item
    revenues_keys = set()
    for year in sorted_years:
        breakdown = yoy_data[year].get("profit_description", {}).get("revenues", {}).get("breakdown", {})
        revenues_keys.update(breakdown.keys())

    for rev_item in revenues_keys:
        key_chain = ("profit_description", "revenues", "breakdown", rev_item)
        vals = build_values(key_chain)
        if len(vals) >= 2:
            cagr_key = f"cagr_revenues_{rev_item}_percent"
            results["cagr_revenues_breakdown_percent"][cagr_key] = calculate_cagr(vals)
        else:
            cagr_key = f"cagr_revenues_{rev_item}_percent"
            results["cagr_revenues_breakdown_percent"][cagr_key] = None

    operating_earnings_segments = set()
    for year in sorted_years:
        breakdown = yoy_data[year].get("profit_description", {}).get("operating_earnings", {}).get("breakdown", {})
        operating_earnings_segments.update(breakdown.keys())

    for segment in operating_earnings_segments:
        vals = []
        for year in sorted_years:
            breakdown = yoy_data[year].get("profit_description", {}).get("operating_earnings", {}).get("breakdown", {})
            val = breakdown.get(segment)
            if val is not None:
                vals.append((int(year), float(val)))
        
        if len(vals) >= 2:
            cagr_key = f"cagr_operating_earnings_{segment}_percent"
            results["cagr_operating_earnings_breakdown_percent"][cagr_key] = calculate_cagr(vals)
        else:
            cagr_key = f"cagr_operating_earnings_{segment}_percent"
            results["cagr_operating_earnings_breakdown_percent"][cagr_key] = None

    return results

def compute_historical_pricing_averages(yoy_data):
    """
    Compute average historical pricing metrics across all years.
    """
    metrics = ['pe_low', 'pe_high', 'pb_low', 'pb_high', 'ps_low', 'ps_high', 
              'pcf_low', 'pcf_high']
    
    # Initialize sums and counts for each metric
    sums = {metric: 0 for metric in metrics}
    counts = {metric: 0 for metric in metrics}
    
    # Accumulate values
    for year_data in yoy_data.values():
        hist_vals = year_data.get('hist_pricing', {})
        for metric in metrics:
            value = hist_vals.get(metric)
            if value is not None and isinstance(value, (int, float)):
                sums[metric] += value
                counts[metric] += 1
    
    # Calculate averages
    averages = {}
    for metric in metrics:
        if counts[metric] > 0:
            averages[f'avg_{metric}'] = sums[metric] / counts[metric]
        else:
            averages[f'avg_{metric}'] = None
            
    return averages

def transform_final_output(final_output: dict, stock_price: float = None):
    yoy_data = final_output.get("data", {})
    sorted_years = sorted(yoy_data.keys())

    # We no longer use 'period', we now want 'fiscal_year_end'
    fiscal_year_end = get_fiscal_year_end(final_output["symbol"])

    company_description_data = {}
    for year in sorted_years:
        company_description_data[str(year)] = yoy_data[year]["company_description"]

    analyses_data = {}
    for year in sorted_years:
        analyses_data[str(year)] = yoy_data[year]["analyses"]

    profit_description_data = {}
    for year in sorted_years:
        profit_description_data[str(year)] = yoy_data[year]["profit_description"]

    balance_sheet_data = {}
    for year in sorted_years:
        balance_sheet_data[str(year)] = yoy_data[year]["balance_sheet"]

    # Compute the "studies" section using the latest year
    studies = {}
    if sorted_years:
        latest_year = sorted_years[-1]
        latest_data = yoy_data[latest_year]
        bs = latest_data["balance_sheet"]
        pd_data = latest_data["profit_description"]

        # Extract required fields
        total_liabilities = bs["liabilities"]["total_liabilities"] or 0
        total_shareholders_equity = bs["shareholders_equity"]["total_shareholders_equity"] or 0

        def val(x): return x if x is not None else 0
        liab_breakdown = bs["liabilities"]["breakdown"]

        # Now we have capital lease obligations from the breakdown
        capital_lease_obligations = val(liab_breakdown.get("capital_lease_obligations"))

        total_debt = total_liabilities - capital_lease_obligations
        total_capital = total_debt + total_shareholders_equity

        # Current liabilities
        accounts_payable = val(liab_breakdown.get("accounts_payable"))
        tax_payables = val(liab_breakdown.get("tax_payables"))
        other_current_liabilities = val(liab_breakdown.get("other_current_liabilities"))
        deferred_revenue = val(liab_breakdown.get("deferred_revenue"))
        short_term_debt = val(liab_breakdown.get("short_term_debt"))

        total_current_liabilities = (accounts_payable + tax_payables +
                                     other_current_liabilities + deferred_revenue +
                                     short_term_debt)

        lt_debt = val(liab_breakdown.get("long_term_debt_minus_capital_lease_obligation"))+short_term_debt
        lt_capital = total_capital - total_current_liabilities

        # net income from profit_description (earnings)
        net_income = pd_data.get("earnings", 0) or 0

        # Addback = cash_and_cash_equivalents + short_term_investment
        assets_breakdown = bs["assets"]["breakdown"]
        cash_equiv = val(assets_breakdown.get("cash_and_cash_equivalents"))
        sti = val(assets_breakdown.get("short_term_investment"))
        addback = cash_equiv + sti

        total_debt_ratio = (total_debt / total_capital) if total_capital != 0 else None
        lt_debt_ratio = (lt_debt / lt_capital) if lt_capital != 0 else None

        years_payback_total_debt = (total_debt / net_income) if net_income != 0 else None
        years_payback_lt_debt = (lt_debt / net_income) if net_income != 0 else None
        years_payback_addback = ((lt_debt - addback) / net_income) if (net_income != 0) else None

        studies = {
            "analysis_of_debt_levels": {
                "total_debt_capital": {
                    "total_debt": total_debt,
                    "total_capital": total_capital,
                    "total_debt_ratio": total_debt_ratio
                },
                "long_term_debt": {
                    "lt_debt": lt_debt,
                    "lt_capital": lt_capital,
                    "lt_debt_ratio": lt_debt_ratio
                },
                "net_income_payback": {
                    "total_debt": total_debt,
                    "net_income": net_income,
                    "years_payback_total_debt": years_payback_total_debt,
                    "lt_debt": lt_debt,
                    "years_payback_lt_debt": years_payback_lt_debt
                },
                "addback_net_inc_payback": {
                    "lt_debt": lt_debt,
                    "net_income": net_income,
                    "addback": addback,
                    "years_payback": years_payback_addback
                }
            }
        }
    
    historical_pricing_averages = compute_historical_pricing_averages(yoy_data)

    rearranged = {
        "summary": {
            "symbol": final_output.get("symbol"),
            "company_name": final_output.get("company_name"),
            "exchange": final_output.get("exchange"),
            "description": final_output.get("description"),
            "sector": final_output.get("sector"),
            "industry": final_output.get("industry"),
            "reported_currency": final_output.get("reported_currency"),
            "isAdr": final_output.get("isAdr")
        },
        "company_description": {
            "fiscal_year_end": fiscal_year_end,
            "stock_price": stock_price,
            "marketCapitalization": final_output.get("marketCapitalization"),
            "data": company_description_data
        },
        "analyses": {
            "investment_characteristics": final_output.get("investment_characteristics", {}),
            "data": analyses_data
        },
        "profit_description": {
            "profit_description_characteristics": final_output.get("profit_description_characteristics", {}),
            "data": profit_description_data
        },
        "balance_sheet": {
            "balance_sheet_characteristics": final_output.get("balance_sheet_characteristics", {}),
            "data": balance_sheet_data
        },
        "studies": studies,
        "historical_pricing": historical_pricing_averages,
        "industry": final_output.get("industry_comparison", {})
    }

    # Add segmentation data directly from yoy_data
    segmentation = {}
    for year in sorted_years:
        year_data = yoy_data[year]
        if "segmentation" in year_data and year_data["segmentation"]:
            segmentation[str(year)] = year_data["segmentation"]
    
    if segmentation:  # Only add if we have segmentation data
        rearranged["segmentation"] = segmentation

    rearranged["qualities"] = final_output.get("qualities")

    return rearranged

def is_percentage_field(key: str) -> bool:
    # Determine if a field should be treated as a percentage
    # Adjust this logic as needed
    key_lower = key.lower()
    percentage_keywords = ["percent", "rate", "roe", "roc", "yield", "margin", "payout",
                           "earnings_percent_revenue", "operating_earnings_percent_revenue", "op_margin_percent", 
                           "tax_rate", "total_debt_ratio", "lt_debt_ratio"]
    return any(kw in key_lower for kw in percentage_keywords)

def format_number(value, key=None):
    """
    Format a number according to the rules:
    - If it's a percentage field, multiply by 100 and append '%', limit to two decimals.
    - If value > 100000, convert to millions with 'M' and two decimals.
    - Otherwise, two decimals max.
    """
    if value is None:
        return value
    
    # If not a number, just return
    if not isinstance(value, (int, float)):
        return value

    # Handle percentage fields
    # if key and is_percentage_field(key):
    #     # Treat value as a ratio: multiply by 100 to get percentage
    #     # Example: 0.2742 => 27.42%
    #     percentage_val = value * 100
    #     return f"{percentage_val:.2f}%"

    # Handle large numbers > 100000
    if abs(value) > 100000:
        # Convert to millions
        value_in_millions = (int)(value / 1_000_000)
        return f"{value_in_millions}"

    # Otherwise, just two decimal places
    # Note: If it's an integer (like dividends_paid=0), formatting will give "0.00"
    # If you want to keep integers as integers, add a check.
    return f"{value:.2f}"

def normalize_data(data, key=None):
    """
    Recursively traverse the dictionary and normalize all numeric values.
    Pass the key along so that we can determine if something is a percentage field.
    """
    if isinstance(data, dict):
        new_dict = {}
        for k, v in data.items():
            new_dict[k] = normalize_data(v, key=k)  # pass the current key down
        return new_dict
    elif isinstance(data, list):
        new_list = []
        for item in data:
            # For lists, we don't have a key, so we just pass None or the parent's key context if available
            new_list.append(normalize_data(item, key=key))
        return new_list
    else:
        # For scalar values, format them using the current key context
        return format_number(data, key=key)

def process_qualities(symbol, ignore_qualities=False, debug=False):
    if not ignore_qualities:
        # Fetch and filter data
        fetch_all_for_ticker(symbol)
        filter_emails_by_ticker(symbol)
        
        # Define filenames
        posts_filename = os.path.join("output", f"{symbol}_posts.json")
        emails_filename = os.path.join("output", f"{symbol}_sent_emails.json")
        combined_filename = os.path.join("output", f"{symbol}_combined_debug.json")
        
        try:
            # Initialize combined data
            combined_data = []

            # Load posts if the file exists
            if os.path.exists(posts_filename):
                with open(posts_filename, "r", encoding="utf-8") as f:
                    posts = json.load(f)
                    combined_data.extend(posts)

            # Load emails if the file exists
            if os.path.exists(emails_filename):
                with open(emails_filename, "r", encoding="utf-8") as f:
                    emails = json.load(f)
                    combined_data.extend(emails)

            # Optionally write combined data to a debug file
            if debug:
                with open(combined_filename, "w", encoding="utf-8") as f:
                    json.dump(combined_data, f, indent=4)
                print(f"Combined data written to {combined_filename} for debugging.")

            if combined_data:
                # Call your existing generate_post_summary() function
                print(f"Attempting to generate post summaries for {symbol}")
                return generate_post_summary(combined_data, symbol)
            else:
                print(f"No posts or emails found for {symbol}. Skipping summary.")
                return "No forum summary available."

        except Exception as e:
            print(f"Error processing data for {symbol}: {e}")
            return "Error generating summary."

def finalize_output(rearranged_output):
    """
    Apply normalization and formatting to the rearranged output.
    """
    # First pass: convert numeric fields
    processed = normalize_data(rearranged_output)
    return processed

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process financial data analysis')
    parser.add_argument('symbol', type=str, help='Stock ticker symbol')
    parser.add_argument('start_year', type=int, help='Start year for analysis')
    parser.add_argument('--ignore_qualities', action='store_true', help='Skip qualities analysis')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode')
    parser.add_argument('--basic_segmentation', action='store_true', 
                       help='Use basic FMP segmentation instead of unified')
    args = parser.parse_args()

    symbol = args.symbol.upper()
    start_year = int(args.start_year)
    ignore_qualities = args.ignore_qualities
    debug = args.debug

    if not FMP_API_KEY:
        print("ERROR: FMP_API_KEY not found in environment. Please set FMP_API_KEY in .env")
        sys.exit(1)
    
     # Get the company's fiscal year end
    fye_string = get_fiscal_year_end(symbol)
    derived_end_year = derive_most_recent_fiscal_year(fye_string)

    # Fallback if we can't parse the FYE
    if not derived_end_year:
        derived_end_year = datetime.date.today().year - 1

    statement_types = ["bs", "ic", "cf", "bs-ar"]
    frequency = "annual"

    # Fetch and save the financial statements
    for statement_type in statement_types:
        get_financials(symbol, statement_type, frequency)

    # Fetch and save the basic financials
    get_basic_financials(symbol)

    # Determine the latest available year from financial statements
    def get_latest_available_year(symbol: str, statement_types: list) -> int:
        available_years = set()
        for stmt in statement_types:
            filename = f"{symbol}_{stmt}_annual.json"
            data = load_json(filename)
            financials = data.get("financials", [])
            for entry in financials:
                date_str = entry.get("date")
                if date_str:
                    year = int(date_str.split('-')[0])
                    available_years.add(year)
        if available_years:
            return max(available_years)
        return derived_end_year  # Fallback to derived_end_year if no data found

    latest_data_year = get_latest_available_year(symbol, statement_types)
    
    # Set end_year to the lesser of derived_end_year and latest_data_year
    end_year = min(derived_end_year, latest_data_year)

    # Validate years
    if start_year > end_year:
        print("START_YEAR cannot be greater than END_YEAR.")
        sys.exit(1)

     # Get unified segmentation if available and not using basic segmentation
    segmentation_data = {}
    if not args.basic_segmentation:
        try:
            from unified_segmentation import process_years
            segmentation_data = process_years(symbol, end_year)
            if not segmentation_data:
                print("No unified segmentation data available, using basic segmentation")
        except Exception as e:
            print(f"Error getting unified segmentation: {e}, using basic segmentation")
    
    # Get basic revenue segmentation if needed
    if not segmentation_data:
        revenue_segmentation = get_revenue_segmentation(symbol)
        if revenue_segmentation:
            # Convert to unified format
            segmentation_data = {}
            for year, segments in revenue_segmentation.items():
                segmentation_data[year] = {
                    "revenue": segments,
                    "segmentation": segments
                }

    years_to_extract = list(range(start_year, end_year + 1))

    # Fetch company profile
    profile = get_company_profile(symbol)
    
    reported_currency = get_reported_currency(symbol)

    # Extract YOY data
    yoy_data = extract_yoy_data(symbol, years_to_extract, 
                                segmentation_data, profile)

    # Compute Investment Characteristics
    investment_characteristics = compute_investment_characteristics(yoy_data)

    # Compute Balance Sheet Characteristics (CAGRs)
    balance_sheet_characteristics = compute_balance_sheet_characteristics(yoy_data)

    # Compute Profit Description Characteristics
    profit_description_characteristics = compute_profit_description_characteristics(yoy_data)

    # Fetch industry comparison data
    try:
        industry_data = get_industry_peers_with_stats(symbol)
    except Exception as e:
        print(f"Error fetching industry data: {e}")
        industry_data = {
            "operatingStatistics": {},
            "marketStatistics": {}
        }
    
    # Get current stock price from quote-short API
    yahoo_symbol = get_yahoo_ticker(profile)
    current_stock_price = get_current_quote_yahoo(yahoo_symbol)
    market_cap = get_current_market_cap_yahoo(yahoo_symbol)

    # Create final output structure with a "header"
    final_output = {
        "symbol": profile.get("symbol", symbol),
        "company_name": profile.get("companyName"),
        "exchange": profile.get("exchange"),
        "description": profile.get("description"),
        "marketCapitalization": market_cap,
        "reported_currency": reported_currency,
        "sector": profile.get("sector"),
        "industry": profile.get("industry"),
        "investment_characteristics": investment_characteristics,
        "balance_sheet_characteristics": balance_sheet_characteristics,
        "profit_description_characteristics": profit_description_characteristics,
        "data": yoy_data,
        "qualities": "",
        "industry_comparison": industry_data,
        "isAdr": profile.get("isAdr", False),
    }

    # Process qualities
    qualities = process_qualities(symbol, ignore_qualities=ignore_qualities, debug=debug)
    final_output["qualities"] = qualities

    rearranged_output = transform_final_output(final_output, stock_price=current_stock_price)
    # rearranged_output = finalize_output(rearranged_output)  # apply normalization

    # Save the consolidated YOY data with header
    output_path = os.path.join("output", f"{symbol}_yoy_consolidated.json")
    with open(output_path, "w") as f:
        json.dump(rearranged_output, f, indent=4)
    print(f"Consolidated YOY data (including company header) saved to {output_path}")

    # 2) Generate the Excel file using our function:
    try:
        generate_excel_for_ticker_year(symbol, end_year)
    except Exception as e:
        print(f"Error generating Excel for {symbol} - {end_year}: {e}")
