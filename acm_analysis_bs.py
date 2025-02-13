# analysis_project/acm_analysis_bs.py
#!/usr/bin/env python3
import os
import json
import logging
import re
import requests
from bs4 import BeautifulSoup
import time
import argparse
from dotenv import load_dotenv
import xml.etree.ElementTree as ET
from urllib.parse import urljoin
from urllib3.util.retry import Retry
import datetime

# Import helper functions from your modules.
from acm_analysis import calculate_cagr, process_qualities
from financial_data_preprocessor import process_financial_statements
from gen_excel_bs import generate_excel_for_ticker_year
from unified_segmentation import get_filing_contents
from utils import get_company_profile, get_current_market_cap_yahoo, get_current_quote_yahoo, get_yahoo_ticker, get_yearly_high_low_yahoo
from industry_comp import get_industry_peers_with_stats

# Load environment variables
load_dotenv()
FMP_API_KEY = os.getenv("FMP_API_KEY")
if not FMP_API_KEY:
    raise ValueError("FMP_API_KEY not found in environment variables")

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

def find_context(soup, context_ref):
    """
    Fallback helper: returns a tag whose id equals context_ref and whose tag name
    (lowercased) contains 'context'.
    """
    return soup.find(lambda tag: tag.get("id") == context_ref and "context" in tag.name.lower())

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

def get_financials(symbol: str, statement_type: str, frequency: str):
    """
    Fetch financial data from Financial Modeling Prep.
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
        return response.json()
    except Exception as e:
        print(f"Error fetching data: {e}")

def get_basic_financials(symbol: str):
    """
    Fetch key metrics from FMP as a stand-in for basic financials.
    """
    url = f"https://financialmodelingprep.com/api/v3/key-metrics/{symbol}?apikey={FMP_API_KEY}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        return {"keyMetrics": data}
    except Exception as e:
        print(f"Error fetching basic financials for {symbol}: {e}")

def extract_series_values_by_year(basic_data: dict, key: str) -> dict:
    """
    Extracts series values (e.g. P/E ratio) by year from FMP key-metrics data.
    """
    field_map = {"pe": "peRatio"}
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
    """
    Extracts FMP data across years (company description, earnings, etc.).
    NOTE: We use SEC data for the balance sheet.
    Now, we also compute the FMP tax rate as:
         incomeTaxExpense / incomeBeforeTax
    """
    bs_data = get_financials(symbol, 'bs', "annual")
    ic_data = get_financials(symbol, 'ic', "annual")
    cf_data = get_financials(symbol, 'cf', "annual")
    basic_data = get_basic_financials(symbol)
    
    try:
        ic_data, bs_data, cf_data = process_financial_statements(
            ticker=symbol,
            ic_data=ic_data,
            bs_data=bs_data,
            cf_data=cf_data
        )
    except Exception as e:
        logger.error(f"Error processing financial statements: {e}")
        return {}
    
    pe_by_year = extract_series_values_by_year(basic_data, 'pe')
    
    def by_year_dict(data):
        res = {}
        if isinstance(data, list):
            for item in data:
                date_str = item.get("date")
                if date_str:
                    y = date_str.split('-')[0]
                    try:
                        year_int = int(y)
                        res[year_int] = item
                    except Exception:
                        pass
            return res
        elif isinstance(data, dict):
            for item in data.get('financials', []):
                date_str = item.get('date')
                if date_str:
                    y = date_str.split('-')[0]
                    try:
                        year_int = int(y)
                        res[year_int] = item
                    except Exception:
                        pass
            return res
        return res

    bs_by_year = by_year_dict(bs_data)
    ic_by_year = by_year_dict(ic_data)
    cf_by_year = by_year_dict(cf_data)
    
    results = {}
    prev_shares_outstanding = None
    yahoo_symbol = get_yahoo_ticker(profile)
    
    for year in years:
        bs = bs_by_year.get(year, {})
        ic = ic_by_year.get(year, {})
        cf = cf_by_year.get(year, {})
        
        filing_url = bs.get('link')
        net_profit = ic.get('netIncome')
        revenues = ic.get('revenue')
        diluted_eps = ic.get('epsdiluted')
        operating_income = ic.get('operatingIncome')
        shares_outstanding = ic.get('weightedAverageShsOutDil')
        operating_eps = operating_income / shares_outstanding if (operating_income is not None and shares_outstanding) else None
        dividends_paid = -1 * cf.get('dividendsPaid', 0)
        shareholder_equity = bs.get('totalStockholdersEquity')
        buyback = None
        if prev_shares_outstanding is not None and shares_outstanding and yahoo_symbol:
            share_change = prev_shares_outstanding - shares_outstanding
            high, low = get_yearly_high_low_yahoo(yahoo_symbol, year)
            avg_price = (high + low) / 2 if (high and low) else None
            if avg_price:
                buyback = share_change * avg_price
        book_value_per_share = (shareholder_equity / shares_outstanding) if (shareholder_equity and shares_outstanding and shares_outstanding != 0) else None
        
        # Compute FMP tax rate using incomeTaxExpense and incomeBeforeTax.
        provision_for_taxes = ic.get('incomeTaxExpense')
        pretax_income = ic.get('incomeBeforeTax')
        fmp_tax_rate = (provision_for_taxes / pretax_income) if (provision_for_taxes is not None and pretax_income and pretax_income != 0) else None
        
        # Get yearly high and low prices once:
        yearly_prices = get_yearly_high_low_yahoo(yahoo_symbol, year)
        yearly_high = yearly_prices[0] if yearly_prices else None
        yearly_low = yearly_prices[1] if yearly_prices else None
        
        # Compute historical pricing metrics if possible:
        pe_low = (yearly_low / diluted_eps) if (yearly_low and diluted_eps and diluted_eps != 0) else None
        pe_high = (yearly_high / diluted_eps) if (yearly_high and diluted_eps and diluted_eps != 0) else None
        pb_low = (yearly_low / book_value_per_share) if (yearly_low and book_value_per_share and book_value_per_share != 0) else None
        pb_high = (yearly_high / book_value_per_share) if (yearly_high and book_value_per_share and book_value_per_share != 0) else None
        sales_per_share = (revenues / shares_outstanding) if (revenues and shares_outstanding and shares_outstanding != 0) else None
        ps_low = (yearly_low / sales_per_share) if (yearly_low and sales_per_share and sales_per_share != 0) else None
        ps_high = (yearly_high / sales_per_share) if (yearly_high and sales_per_share and sales_per_share != 0) else None
        depreciation = cf.get('depreciationAndAmortization')
        if depreciation is not None and net_profit is not None and shares_outstanding and shares_outstanding != 0:
            addback_dep_earnings_ps = (net_profit + depreciation) / shares_outstanding
        else:
            addback_dep_earnings_ps = None
        pcf_low = (yearly_low / addback_dep_earnings_ps) if (yearly_low and addback_dep_earnings_ps and addback_dep_earnings_ps != 0) else None
        pcf_high = (yearly_high / addback_dep_earnings_ps) if (yearly_high and addback_dep_earnings_ps and addback_dep_earnings_ps != 0) else None

        company_description = {
            "net_profit": net_profit,
            "diluted_eps": diluted_eps,
            "operating_eps": operating_eps,
            "pe_ratio": pe_by_year.get(year),
            "price_low": yearly_low,
            "price_high": yearly_high,
            "dividends_paid": dividends_paid,
            "dividends_per_share": (dividends_paid / shares_outstanding
                                    if (dividends_paid and shares_outstanding and shares_outstanding != 0) else None),
            "avg_dividend_yield": None,  # computed later in unified mapping
            "shares_outstanding": shares_outstanding,
            "buyback": buyback,
            "share_equity": shareholder_equity,
            "book_value_per_share": book_value_per_share
        }
        
        hist_pricing = {
            "pe_low": pe_low,
            "pe_high": pe_high,
            "pb_low": pb_low,
            "pb_high": pb_high,
            "ps_low": ps_low,
            "ps_high": ps_high,
            "pcf_low": pcf_low,
            "pcf_high": pcf_high
        }

        analyses = {
            "revenue": revenues,
            "tax_rate": fmp_tax_rate  # FMP tax rate computed above
        }
        profit_description = {
            "gross_revenues": revenues,
            "filing_url": filing_url
        }
        
        results[year] = {
            "company_description": company_description,
            "analyses": analyses,
            "profit_description": profit_description,
            "hist_pricing": hist_pricing
            # FMP balance sheet data is ignored; we use SEC balance sheet.
        }
        prev_shares_outstanding = shares_outstanding
        
    return results

def compute_investment_characteristics(yoy_data: dict):
    """
    Compute investment characteristics based on year‐over‐year (YOY) data.
    For the “sales_analysis” subsections, we use the growth in total assets (and assets per share)
    rather than revenues and sales per share.
    Assumes that each year’s unified data (yoy_data[year]) contains a "company_description" with keys:
       - "operating_eps", "diluted_eps", "dividends_per_share", "buyback", "net_profit", "shares_outstanding"
       - "assets" (from SEC balance sheet, as computed in create_unified_year_output)
    Also assumes that a calculate_cagr(values_by_year: list) function is available.
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
            "growth_rate_percent_revenues": None,              # now: growth rate in assets
            "growth_rate_percent_sales_per_share": None          # now: growth rate in assets per share
        },
        "sales_analysis_last_5_years": {
            "growth_rate_percent_revenues": None,              # now: growth rate in assets
            "growth_rate_percent_sales_per_share": None          # now: growth rate in assets per share
        }
    }

    sorted_years = sorted(yoy_data.keys())

    def build_values_by_year(metric_key_chain):
        # Follow a chain of keys (e.g. ("company_description", "operating_eps"))
        # to return a list of (year, value) pairs.
        vals = []
        for year in sorted_years:
            d = yoy_data[year]
            for k in metric_key_chain:
                d = d.get(k, {})
            # If d is not a dict, then it is our value.
            if not isinstance(d, dict):
                val = d
            else:
                val = None
            if val is not None:
                vals.append((year, val))
        return vals

    # --- Earnings Analysis (unchanged) ---

    # (a) Growth Rate in Operating EPS
    ops_eps_values = build_values_by_year(("company_description", "operating_eps"))
    if len(ops_eps_values) >= 2:
        cagr_ops = calculate_cagr(ops_eps_values)
        investment_characteristics["earnings_analysis"]["growth_rate_percent_operating_eps"] = cagr_ops

    # (b) Earnings Quality: ratio of diluted_eps to operating_eps averaged over years
    total_eps = total_operating_eps = count_eps = count_operating_eps = 0
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
        quality_percent = avg_eps / avg_operating_eps
        investment_characteristics["earnings_analysis"]["quality_percent"] = round(quality_percent, 2)

    # --- Use of Earnings (unchanged) ---

    # (a) Average Dividend Payout %
    sum_dividends_per_share = sum_operating_eps_for_div = count_for_div = 0
    for year in sorted_years:
        dividends = yoy_data[year]["company_description"].get("dividends_per_share")
        operating_eps = yoy_data[year]["company_description"].get("operating_eps")
        if dividends is not None and operating_eps and operating_eps != 0:
            sum_dividends_per_share += dividends
            sum_operating_eps_for_div += operating_eps
            count_for_div += 1
    if count_for_div > 0:
        avg_dividend = sum_dividends_per_share / count_for_div
        avg_operating = sum_operating_eps_for_div / count_for_div
        investment_characteristics["use_of_earnings_analysis"]["avg_dividend_payout_percent"] = round(avg_dividend / avg_operating, 2)

    # (b) Average Stock Buyback %
    sum_buyback = sum_net_profit = 0
    for year in sorted_years:
        buyback = yoy_data[year]["company_description"].get("buyback")
        net_profit = yoy_data[year]["company_description"].get("net_profit")
        if buyback is not None and net_profit and net_profit != 0:
            sum_buyback += buyback
            sum_net_profit += net_profit
    if sum_net_profit != 0:
        investment_characteristics["use_of_earnings_analysis"]["avg_stock_buyback_percent"] = round(sum_buyback / sum_net_profit, 2)

    # --- Sales Analysis (asset-based now) ---

    # (a) Total Assets Growth (instead of revenues)
    asset_values = []
    for year in sorted_years:
        assets = yoy_data[year]["company_description"].get("assets")
        if assets is not None:
            asset_values.append((year, assets))
    if len(asset_values) >= 2:
        # Note: The label "growth_rate_percent_revenues" remains for backward compatibility;
        # you might want to rename the key in your final output.
        investment_characteristics["sales_analysis"]["growth_rate_percent_revenues"] = calculate_cagr(asset_values)

    # (b) Assets per Share Growth
    asset_per_share_values = []
    for year in sorted_years:
        comp = yoy_data[year]["company_description"]
        assets = comp.get("assets")
        shares = comp.get("shares_outstanding")
        if assets is not None and shares and shares != 0:
            asset_per_share = assets / shares
            asset_per_share_values.append((year, asset_per_share))
    if len(asset_per_share_values) >= 2:
        investment_characteristics["sales_analysis"]["growth_rate_percent_sales_per_share"] = calculate_cagr(asset_per_share_values)

    # --- Sales Analysis Last 5 Years (asset-based) ---

    last_5_years = sorted_years[-5:]
    asset_values_5y = []
    for y in last_5_years:
        assets = yoy_data[y]["company_description"].get("assets")
        if assets is not None:
            asset_values_5y.append((y, assets))
    if len(asset_values_5y) >= 2:
        investment_characteristics["sales_analysis_last_5_years"]["growth_rate_percent_revenues"] = calculate_cagr(asset_values_5y)

    asset_per_share_values_5y = []
    for y in last_5_years:
        comp = yoy_data[y]["company_description"]
        assets = comp.get("assets")
        shares = comp.get("shares_outstanding")
        if assets is not None and shares and shares != 0:
            asset_per_share = assets / shares
            asset_per_share_values_5y.append((y, asset_per_share))
    if len(asset_per_share_values_5y) >= 2:
        investment_characteristics["sales_analysis_last_5_years"]["growth_rate_percent_sales_per_share"] = calculate_cagr(asset_per_share_values_5y)

    return investment_characteristics

def compute_balance_sheet_characteristics(bs_data: dict) -> dict:
    """
    Given the inverted balance sheet data (a dict keyed by year),
    compute the CAGR (as a percent string) for total assets, total liabilities,
    and total shareholders' equity.
    
    Expects each year’s balance sheet data to have:
       - assets: a dict with key "assets" (or "total_assets") containing a numeric value
       - liabilities: a dict with key "liabilities" containing a numeric value
       - shareholders_equity: a dict with key "shareholders_equity" containing a numeric value
    """
    balance_sheet_characteristics = {
        "cagr_total_assets_percent": None,
        "cagr_total_liabilities_percent": None,
        "cagr_total_shareholders_equity_percent": None
    }
    
    sorted_years = sorted(bs_data.keys())
    if len(sorted_years) < 2:
        return balance_sheet_characteristics

    def build_bs_values(key_chain: tuple) -> list:
        # key_chain is a tuple of keys to traverse; for example: ("assets", "assets")
        vals = []
        for year in sorted_years:
            d = bs_data[year]
            for key in key_chain:
                d = d.get(key, {})
            if isinstance(d, (int, float)):
                vals.append((year, d))
        return vals

    # Build values for total assets from path ("assets", "assets")
    assets_values = build_bs_values(("assets", "assets"))
    if len(assets_values) >= 2:
        balance_sheet_characteristics["cagr_total_assets_percent"] = calculate_cagr(assets_values)

    # Build values for total liabilities from path ("liabilities", "liabilities")
    liabilities_values = build_bs_values(("liabilities", "liabilities"))
    if len(liabilities_values) >= 2:
        balance_sheet_characteristics["cagr_total_liabilities_percent"] = calculate_cagr(liabilities_values)

    # Build values for total shareholders' equity from path ("shareholders_equity", "shareholders_equity")
    equity_values = build_bs_values(("shareholders_equity", "shareholders_equity"))
    if len(equity_values) >= 2:
        balance_sheet_characteristics["cagr_total_shareholders_equity_percent"] = calculate_cagr(equity_values)

    return balance_sheet_characteristics

def compute_profit_description_characteristics(unified_results: dict) -> dict:
    """
    Compute CAGR metrics for profit description based on unified per‐year data.
    For each metric, the function builds a list of (year, value) pairs and calls calculate_cagr()
    if at least two data points are available.
    
    It computes CAGR for:
      - Gross revenues total (from profit_description → gross_revenues → total)
      - Gross revenues breakdown items (from profit_description → gross_revenues → breakdown)
      - Investment income (from profit_description → investment_income)
      - Internal costs total (from profit_description → internal_costs → total)
      - Internal costs breakdown items (from profit_description → internal_costs → breakdown)
      - Operating margin total (from profit_description → operating_margin → total)
      - Operating margin breakdown items (from profit_description → operating_margin → breakdown)
      - External costs total (from profit_description → external_costs → total)
      - External costs breakdown items (from profit_description → external_costs → breakdown)
      - Earnings (from profit_description → earnings)
    """
    years = sorted(unified_results.keys())
    result = {}

    def get_total(metric_key: str) -> list:
        values = []
        for year in years:
            pd = unified_results[year].get("profit_description", {})
            val = pd.get(metric_key)
            if isinstance(val, (int, float)):
                values.append((year, val))
        return values

    def get_nested_total(metric_path: tuple) -> list:
        values = []
        for year in years:
            pd = unified_results[year].get("profit_description", {})
            d = pd
            for key in metric_path:
                d = d.get(key, {})
            if isinstance(d, (int, float)):
                values.append((year, d))
        return values

    def get_breakdown(metric_path: tuple) -> dict:
        breakdown_values = {}
        for year in years:
            pd = unified_results[year].get("profit_description", {})
            d = pd
            for key in metric_path:
                d = d.get(key, {})
            if isinstance(d, dict):
                for bk, val in d.items():
                    if isinstance(val, (int, float)):
                        breakdown_values.setdefault(bk, []).append((year, val))
        return breakdown_values

    # Gross Revenues Total
    gross_total = get_nested_total(("gross_revenues", "total"))
    result["cagr_gross_revenues_percent"] = calculate_cagr(gross_total) if len(gross_total) >= 2 else None

    # Investment Income
    inv_income = get_total("investment_income")
    result["cagr_investment_income_percent"] = calculate_cagr(inv_income) if len(inv_income) >= 2 else None

    # Internal Costs Total
    internal_total = get_nested_total(("internal_costs", "total"))
    result["cagr_internal_costs_percent"] = calculate_cagr(internal_total) if len(internal_total) >= 2 else None

    # Operating Margin Total
    op_margin_total = get_nested_total(("operating_margin", "total"))
    result["cagr_operating_margin_percent"] = calculate_cagr(op_margin_total) if len(op_margin_total) >= 2 else None

    # External Costs Total
    ext_total = get_nested_total(("external_costs", "total"))
    result["cagr_external_costs_percent"] = calculate_cagr(ext_total) if len(ext_total) >= 2 else None

    # Earnings
    earnings = get_total("earnings")
    result["cagr_earnings_percent"] = calculate_cagr(earnings) if len(earnings) >= 2 else None

    # Breakdown for Gross Revenues
    gr_breakdown = get_breakdown(("gross_revenues", "breakdown"))
    result["cagr_gross_revenues_breakdown_percent"] = {}
    for bk, values in gr_breakdown.items():
        result["cagr_gross_revenues_breakdown_percent"][f"cagr_gross_revenues_{bk}_percent"] = calculate_cagr(values) if len(values) >= 2 else None

    # Breakdown for Internal Costs
    ic_breakdown = get_breakdown(("internal_costs", "breakdown"))
    result["cagr_internal_costs_breakdown_percent"] = {}
    for bk, values in ic_breakdown.items():
        result["cagr_internal_costs_breakdown_percent"][f"cagr_internal_costs_{bk}_percent"] = calculate_cagr(values) if len(values) >= 2 else None

    # Breakdown for Operating Margin
    om_breakdown = get_breakdown(("operating_margin", "breakdown"))
    result["cagr_operating_margin_breakdown_percent"] = {}
    for bk, values in om_breakdown.items():
        result["cagr_operating_margin_breakdown_percent"][f"cagr_operating_margin_{bk}_percent"] = calculate_cagr(values) if len(values) >= 2 else None

    # Breakdown for External Costs
    ec_breakdown = get_breakdown(("external_costs", "breakdown"))
    result["cagr_external_costs_breakdown_percent"] = {}
    for bk, values in ec_breakdown.items():
        result["cagr_external_costs_breakdown_percent"][f"cagr_external_costs_{bk}_percent"] = calculate_cagr(values) if len(values) >= 2 else None

    return result

class EDGARExhibit13Finder:
    """
    Retrieves company filings via EDGAR and locates the XML filing document.
    """
    BASE_URL = "https://www.sec.gov"
    
    def __init__(self, user_agent: str):
        self.headers = {
            "User-Agent": user_agent,
            "Accept-Encoding": "gzip, deflate",
        }
        self.session = requests.Session()
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["HEAD", "GET", "OPTIONS"],
        )
        adapter = requests.adapters.HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("https://", adapter)
        
    def get_cik_from_ticker(self, ticker: str) -> str:
        url = f"https://financialmodelingprep.com/api/v3/balance-sheet-statement/{ticker}"
        params = {
            "period": "annual",
            "apikey": FMP_API_KEY,
            "limit": 1,
        }
        try:
            response = self.session.get(url, params=params, timeout=(10, 30))
            response.raise_for_status()
            data = response.json()
            if not data:
                raise ValueError(f"No data found for ticker {ticker}")
            cik = data[0].get("cik", "").lstrip("0")
            if not cik:
                raise ValueError(f"No CIK found for ticker {ticker}")
            return cik
        except requests.exceptions.Timeout:
            logger.error(f"Timeout while fetching CIK for ticker {ticker}")
            raise
        except requests.exceptions.RequestException as e:
            logger.error(f"Request error getting CIK for ticker {ticker}: {e}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error getting CIK for ticker {ticker}: {e}")
            raise
        
    def get_company_filings(self, cik: str) -> dict:
        url = urljoin(
            self.BASE_URL,
            f"/cgi-bin/browse-edgar?action=getcompany&CIK={cik}&type=10-K&dateb=&owner=exclude&start=0&count=40&output=atom",
        )
        response = self.session.get(url, headers=self.headers, timeout=(10, 30))
        response.raise_for_status()
        try:
            root = ET.fromstring(response.content)
        except ET.ParseError as e:
            logger.error("Error parsing XML from SEC response.")
            raise e
        
        entries = []
        ns = {"atom": "http://www.w3.org/2005/Atom"}
        for entry in root.findall("atom:entry", ns):
            accession_number = ""
            filing_href = ""
            filing_date = ""
            id_elem = entry.find("atom:id", ns)
            if id_elem is not None and id_elem.text:
                accession_match = re.search(r"accession-number=(\d{10}-\d{2}-\d{6})", id_elem.text)
                if accession_match:
                    accession_number = accession_match.group(1)
            date_elem = entry.find("atom:updated", ns)
            if date_elem is not None and date_elem.text:
                filing_date = date_elem.text.split("T")[0]
            link_elem = entry.find("atom:link", ns)
            if link_elem is not None:
                href = link_elem.get("href")
                if href:
                    filing_href = urljoin(self.BASE_URL, href)
            if accession_number and filing_href:
                entries.append({
                    "accession_number": accession_number,
                    "filing_href": filing_href,
                    "filing_date": filing_date,
                })
        return {"filings": entries}
    
    def get_filing_detail(self, filing_url: str) -> dict:
        """
        Retrieves the filing detail page and locates the XML filing document.
        """
        response = self.session.get(filing_url, headers=self.headers, timeout=(10, 30))
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        documents = {"xml": None}
        table = soup.find("table", class_="tableFile")
        if table:
            for row in table.find_all("tr"):
                cells = row.find_all("td")
                if len(cells) >= 3:
                    description = cells[1].get_text(strip=True)
                    if ("extracted" in description.lower() and 
                        "instance document" in description.lower() and 
                        "xbrl" in description.lower()):
                        document_link = cells[2].find("a")
                        if document_link:
                            href = document_link.get("href")
                            if href and href.lower().endswith(".xml"):
                                documents["xml"] = urljoin(self.BASE_URL, href)
                                break
        if not documents["xml"]:
            xml_link = soup.select_one('a[href$="_htm.xml"]')
            if xml_link:
                documents["xml"] = urljoin(self.BASE_URL, xml_link.get("href"))
        if not documents["xml"]:
            logger.warning(f"No XML filing document found in filing page: {filing_url}")
        return documents

class MetricsExtractor:
    """
    Extracts metrics from an SEC XML filing containing inline XBRL facts.
    The output for each year includes:
      - profit_desc
      - balance_sheet
      - segmentation
    """
    def __init__(self, user_agent: str, config: dict):
        if not config:
            raise ValueError("A valid metrics configuration must be provided. Terminating.")
        
        # Extract required mappings from the config.
        self.profit_desc_metrics = config.get("profit_desc_metrics")
        self.balance_sheet_metrics = config.get("balance_sheet_metrics")
        self.segmentation_mapping = config.get("segmentation_mapping")
        self.balance_sheet_categories = config.get("balance_sheet_categories")
        
        # Ensure all required config sections are present.
        if not (self.profit_desc_metrics and self.balance_sheet_metrics and 
                self.segmentation_mapping and self.balance_sheet_categories):
            raise ValueError("Incomplete metrics configuration. Please provide profit, balance sheet, segmentation mappings, and balance sheet categories.")

        self.config = config  # Save entire config if needed later.
        
        self.headers = {
            "User-Agent": user_agent,
            "Accept-Encoding": "gzip, deflate",
        }
        self.session = requests.Session()
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["HEAD", "GET", "OPTIONS"],
        )
        adapter = requests.adapters.HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("https://", adapter)
    
    def parse_context(self, soup, context_ref):
        logger.debug(f"Parsing context: {context_ref}")
        context = soup.find("context", {"id": context_ref})
        if context:
            period_elem = context.find("period")
            if period_elem:
                start = period_elem.find("startDate")
                end = period_elem.find("endDate")
                instant = period_elem.find("instant")
                if start and end:
                    return {"period": f"{start.text.strip()} to {end.text.strip()}"}
                elif instant:
                    return {"period": f"As of {instant.text.strip()}"}
        context = find_context(soup, context_ref)
        if context:
            period_elem = context.find("period")
            if period_elem:
                start = period_elem.find("startDate")
                end = period_elem.find("endDate")
                instant = period_elem.find("instant")
                if start and end:
                    return {"period": f"{start.text.strip()} to {end.text.strip()}"}
                elif instant:
                    return {"period": f"As of {instant.text.strip()}"}
        year_pattern = r"(\d{4})"
        year_match = re.search(year_pattern, context_ref)
        if year_match:
            year = year_match.group(1)
            return {"period": f"{year}-01-01 to {year}-12-31"}
        return {}
    
    def process_mapping(self, soup, mapping):
        local = {}
        for metric_name, tag in mapping.items():
            elems = soup.find_all(tag)
            for elem in elems:
                try:
                    numeric_value = float(elem.get_text(strip=True))
                    scale = elem.get("scale", "0")
                    if scale and scale != "0":
                        numeric_value *= 10 ** int(scale)
                    context_ref = elem.get("contextRef") or elem.get("contextref")
                    if not context_ref:
                        continue
                    context_data = self.parse_context(soup, context_ref)
                    period_text = context_data.get("period", "")
                    if not period_text:
                        continue
                    year_match = re.search(r"(\d{4})", period_text)
                    if not year_match:
                        continue
                    year = year_match.group(1)
                    if year not in local:
                        local[year] = {}
                    local[year][metric_name] = numeric_value
                except (ValueError, TypeError) as e:
                    logger.error(f"Error processing {metric_name}: {e}")
                    continue
        return local
    
    def process_segmentation(self, soup):
        seg_results = {}
        for seg_key, seg_info in self.segmentation_mapping.items():
            tag = seg_info["tag"]
            required = seg_info["explicitMembers"]
            elems = soup.find_all(tag)
            for elem in elems:
                try:
                    context_ref = elem.get("contextRef") or elem.get("contextref")
                    if not context_ref:
                        continue
                    context = soup.find("context", {"id": context_ref})
                    if not context:
                        continue
                    entity = context.find("entity")
                    if not entity:
                        continue
                    segment = entity.find("segment")
                    if not segment:
                        continue
                    explicit_members = segment.find_all(lambda t: "explicitmember" in t.name.lower())
                    criteria_met = True
                    for dim, expected in required.items():
                        match_found = any(
                            (exp.get("dimension") == dim and exp.get_text(strip=True) == expected)
                            for exp in explicit_members
                        )
                        if not match_found:
                            criteria_met = False
                            break
                    if not criteria_met:
                        continue
                    numeric_value = float(elem.get_text(strip=True))
                    scale = elem.get("scale", "0")
                    if scale and scale != "0":
                        numeric_value *= 10 ** int(scale)
                    context_data = self.parse_context(soup, context_ref)
                    period_text = context_data.get("period", "")
                    if not period_text:
                        continue
                    year_match = re.search(r"(\d{4})", period_text)
                    if not year_match:
                        continue
                    year = year_match.group(1)
                    if year not in seg_results:
                        seg_results[year] = {}
                    seg_results[year][seg_key] = seg_results[year].get(seg_key, 0) + numeric_value
                except Exception as e:
                    logger.error(f"Error processing segmentation {seg_key}: {e}")
                    continue
        return seg_results
    
    def extract_metrics(self, xml_url: str) -> dict:
        max_retries = 3
        attempt = 0
        content = None
        while attempt < max_retries:
            try:
                logger.info("Fetching filing data...")
                content, meta_info = get_filing_contents(xml_url)
                if content:
                    break
            except requests.exceptions.Timeout as e:
                attempt += 1
                logger.error(f"Timeout error fetching filing data (attempt {attempt}/{max_retries}): {e}")
                time.sleep(5)
        if not content:
            logger.error("No filing content retrieved after maximum retries.")
            return {}
        
        logger.info("Parsing XML filing...")
        soup = BeautifulSoup(content, "lxml-xml")
        results = {}
        
        profit_data = self.process_mapping(soup, self.profit_desc_metrics)
        balance_data = self.process_mapping(soup, self.balance_sheet_metrics)
        segmentation_data = self.process_segmentation(soup)
        
        # Retrieve balance_sheet_categories configuration from the config.
        balance_sheet_categories = self.balance_sheet_categories
        if not balance_sheet_categories:
            raise ValueError("Missing 'balance_sheet_categories' in configuration.")

        years = set(profit_data.keys()) | set(balance_data.keys()) | set(segmentation_data.keys())
        for year in years:
            results[year] = {
                "profit_desc": profit_data.get(year, {}),
                "balance_sheet": {
                    "assets": {},
                    "liabilities": {},
                    "shareholders_equity": {}
                },
                "segmentation": segmentation_data.get(year, {})
            }
            if year in balance_data:
                for metric_name, value in balance_data[year].items():
                    assigned = False
                    # Loop through each category (assets, liabilities, shareholders_equity) as defined in the config.
                    for category, metrics_list in balance_sheet_categories.items():
                        if metric_name in metrics_list:
                            results[year]["balance_sheet"][category][metric_name] = value
                            assigned = True
                            break
                    if not assigned:
                        # If the metric isn't found in any category, assign it to assets by default.
                        results[year]["balance_sheet"]["assets"][metric_name] = value
        logger.info(f"Extracted SEC metrics: {results}")
        return results

def create_unified_year_output(year, fmp_data, sec_data):
    """
    Remaps FMP and SEC data into the unified structure:
    
    "company_description": {
       "net_profit",
       "diluted_eps",
       "operating_eps",
       "pe_ratio",
       "price_low",
       "price_high",
       "dividends_paid",
       "dividends_per_share",
       "avg_dividend_yield",
       "shares_outstanding",
       "buyback",
       "share_equity",
       "book_value_per_share",
       "assets" (from SEC balance sheet),
       "return_on_equity" = net_profit / share_equity,
       "return_on_assets" = net_profit / assets,
       "leverage_ratio" = assets / share_equity
    },
    "analyses": {
       "premium_earned" = gross_revenues from SEC profit_desc,
       "benefit_claims" = losses_and_expenses from SEC profit_desc,
       "gross_underwriting_profit" = premium_earned - benefit_claims,
       "underwriting_yield_on_asset" = gross_underwriting_profit / assets,
       "investment_income" = investment_income from SEC profit_desc,
       "investment_yield_on_asset" = investment_income / assets,
       "non_claim_expenses" = acquisition_costs + underwriting_expenses,
       "expense_yield_on_asset" = non_claim_expenses / assets,
       "tax_rate" = FMP tax rate,
       "premium_equity_ratio" = premium_earned / share_equity
    },
    "balance_sheet": { SEC balance sheet data (total vs. breakdown) },
    "profit_description": {
       "gross_revenues": { "total": gross_revenues, "breakdown": segmentation breakdown },
       "investment_income",
       "internal_costs": {
         "total" = losses_and_expenses + acquisition_costs + underwriting_expenses + service_expenses,
         "breakdown": { "losses_and_expenses", "acquisition_costs", "underwriting_expenses", "service_expenses" }
       },
       "operating_margin": {
         "total" = gross_revenues + investment_income - total internal costs,
         "breakdown": {
            "underwriting": (sum(gross_revenue segments) - (losses_and_expenses + acquisition_costs + underwriting_expenses)),
            "pretax_combined_ratio": (sum(gross_revenue segments) - underwriting_margin) / (sum(gross_revenue segments)),
            "pretax_insurance_yield_on_equity": underwriting_margin / share_equity,
            "pretax_return_on_equity": operating_margin_total / share_equity
         }
       },
       "external_costs": {
         "total": taxes + interest_expenses,
         "breakdown": { "taxes", "interest_expenses" }
       },
       "earnings": operating_margin_total - external_costs_total,
       "equity_employed": share_equity,
       "shares_repurchased": buyback,
       "filing_url": filing_url
    },
    "segmentation": {
       "total": (sum of segmentation values),
       "breakdown": segmentation dict
    }
    """
    # FMP company data:
    fmp_comp = fmp_data.get("company_description", {})
    net_profit           = fmp_comp.get("net_profit")
    diluted_eps          = fmp_comp.get("diluted_eps")
    operating_eps        = fmp_comp.get("operating_eps")
    pe_ratio             = fmp_comp.get("pe_ratio")
    price_low            = fmp_comp.get("price_low")
    price_high           = fmp_comp.get("price_high")
    dividends_paid       = fmp_comp.get("dividends_paid")
    dividends_per_share  = fmp_comp.get("dividends_per_share")
    shares_outstanding   = fmp_comp.get("shares_outstanding")
    buyback              = fmp_comp.get("buyback")
    share_equity         = fmp_comp.get("share_equity")
    book_value_per_share = fmp_comp.get("book_value_per_share")
    filing_url = fmp_data.get("profit_description", {}).get("filing_url")

    # SEC balance sheet (use total assets):
    sec_balance = sec_data.get("balance_sheet", {})
    sec_assets = sec_balance.get("assets", {})
    assets = sec_assets.get("assets") or sec_assets.get("total_assets")
    
    # Derived ratios:
    return_on_equity = (net_profit / share_equity) if (net_profit is not None and share_equity) else None
    return_on_assets = (net_profit / assets) if (net_profit is not None and assets) else None
    leverage_ratio   = (assets / share_equity) if (assets is not None and share_equity) else None
    avg_dividend_yield = (dividends_per_share / ((price_low + price_high) / 2)
                          if (dividends_per_share and price_low and price_high and ((price_low+price_high)/2) != 0)
                          else None)
    
    company_description = {
        "net_profit": net_profit,
        "diluted_eps": diluted_eps,
        "operating_eps": operating_eps,
        "pe_ratio": pe_ratio,
        "price_low": price_low,
        "price_high": price_high,
        "dividends_paid": dividends_paid,
        "dividends_per_share": dividends_per_share,
        "avg_dividend_yield": avg_dividend_yield,
        "shares_outstanding": shares_outstanding,
        "buyback": buyback,
        "share_equity": share_equity,
        "book_value_per_share": book_value_per_share,
        "assets": assets,
        "return_on_equity": return_on_equity,
        "return_on_assets": return_on_assets,
        "leverage_ratio": leverage_ratio
    }
    
    # SEC profit data:
    sec_profit = sec_data.get("profit_desc", {})
    premium_earned    = sec_profit.get("gross_revenues")
    benefit_claims    = sec_profit.get("losses_and_expenses")
    acquisition_costs = sec_profit.get("acquisition_costs")
    underwriting_expenses = sec_profit.get("underwriting_expenses")
    service_expenses  = sec_profit.get("service_expenses")
    taxes             = sec_profit.get("taxes")
    interest_expenses = sec_profit.get("interest_expenses")
    investment_income = sec_profit.get("investment_income")
    
    gross_underwriting_profit = (premium_earned - benefit_claims
                                 if premium_earned is not None and benefit_claims is not None
                                 else None)
    underwriting_yield_on_asset = (gross_underwriting_profit / assets
                                   if gross_underwriting_profit is not None and assets
                                   else None)
    non_claim_expenses = (acquisition_costs + underwriting_expenses
                          if acquisition_costs is not None and underwriting_expenses is not None
                          else None)
    expense_yield_on_asset = (non_claim_expenses / assets
                              if non_claim_expenses is not None and assets
                              else None)
    premium_equity_ratio = (premium_earned / share_equity
                            if premium_earned is not None and share_equity
                            else None)
    
    # Note: tax rate now comes from FMP. Retrieve it from FMP analysis:
    fmp_analyses = fmp_data.get("analyses", {})
    fmp_tax_rate = fmp_analyses.get("tax_rate")
    
    analyses = {
        "premium_earned": premium_earned,
        "benefit_claims": benefit_claims,
        "gross_underwriting_profit": gross_underwriting_profit,
        "underwriting_yield_on_asset": underwriting_yield_on_asset,
        "investment_income": investment_income,
        "investment_yield_on_asset": (investment_income / assets if investment_income is not None and assets else None),
        "non_claim_expenses": non_claim_expenses,
        "expense_yield_on_asset": expense_yield_on_asset,
        "tax_rate": fmp_tax_rate,
        "premium_equity_ratio": premium_equity_ratio
    }
    
    internal_costs_total = ((benefit_claims or 0) +
                            (acquisition_costs or 0) +
                            (underwriting_expenses or 0) +
                            (service_expenses or 0))
    operating_margin_total = ((premium_earned or 0) + (investment_income or 0) - internal_costs_total)
    external_costs_total = ((taxes or 0) + (interest_expenses or 0))
    earnings = operating_margin_total - external_costs_total
    
    seg_breakdown = sec_data.get("segmentation", {})
    gross_revenue_segments = sum(seg_breakdown.values()) if seg_breakdown else None
    underwriting_margin = (gross_revenue_segments - ((benefit_claims or 0) +
                                                     (acquisition_costs or 0) +
                                                     (underwriting_expenses or 0))
                           if gross_revenue_segments is not None else None)
    pretax_combined_ratio = ((gross_revenue_segments - underwriting_margin) / gross_revenue_segments
                             if gross_revenue_segments and gross_revenue_segments != 0 else None)
    pretax_insurance_yield_on_equity = (underwriting_margin / share_equity
                                        if underwriting_margin is not None and share_equity else None)
    pretax_return_on_equity = (operating_margin_total / share_equity
                               if share_equity and share_equity != 0 else None)
    
    profit_description = {
        "gross_revenues": {
            "total": premium_earned,
            "breakdown": seg_breakdown
        },
        "investment_income": investment_income,
        "internal_costs": {
            "total": internal_costs_total,
            "breakdown": {
                "losses_and_expenses": benefit_claims,
                "acquisition_costs": acquisition_costs,
                "underwriting_expenses": underwriting_expenses,
                "service_expenses": service_expenses
            }
        },
        "operating_margin": {
            "total": operating_margin_total,
            "breakdown": {
                "underwriting": underwriting_margin,
                "pretax_combined_ratio": pretax_combined_ratio,
                "pretax_insurance_yield_on_equity": pretax_insurance_yield_on_equity,
                "pretax_return_on_equity": pretax_return_on_equity
            }
        },
        "external_costs": {
            "total": external_costs_total,
            "breakdown": {
                "taxes": taxes,
                "interest_expenses": interest_expenses
            }
        },
        "earnings": earnings,
        "equity_employed": share_equity,
        "shares_repurchased": buyback,
        "filing_url": filing_url
    }
    
    seg_total = sum(seg_breakdown.values()) if seg_breakdown else None
    segmentation = {
        "total": seg_total,
        "breakdown": seg_breakdown
    }
    
    unified = {
        "company_description": company_description,
        "analyses": analyses,
        "balance_sheet": sec_data.get("balance_sheet", {}),
        "profit_description": profit_description,
        "segmentation": segmentation
    }
    historical_pricing = fmp_data.get("hist_pricing", {})
    unified["historical_pricing"] = historical_pricing
    return unified

def compute_historical_pricing_averages(yoy_data: dict):
    metrics = ['pe_low', 'pe_high', 'pb_low', 'pb_high', 'ps_low', 'ps_high', 'pcf_low', 'pcf_high']
    sums = {metric: 0 for metric in metrics}
    counts = {metric: 0 for metric in metrics}
    
    for year_data in yoy_data.values():
        # Changed key from 'hist_pricing' to 'historical_pricing'
        hist = year_data.get('historical_pricing', {})
        for metric in metrics:
            value = hist.get(metric)
            if value is not None and isinstance(value, (int, float)):
                sums[metric] += value
                counts[metric] += 1
    
    averages = {}
    for metric in metrics:
        averages[f'avg_{metric}'] = (sums[metric] / counts[metric]) if counts[metric] > 0 else None
    return averages

def main():
    parser = argparse.ArgumentParser(description="Extract unified insurance metrics from FMP and SEC filings")
    parser.add_argument("ticker", type=str, help="Company ticker symbol")
    parser.add_argument("start_year", type=str, help="Start year (YYYY)")
    parser.add_argument("--email", type=str, default="rgranowski@gmail.com", help="Your email for SEC user agent")
    parser.add_argument("--output", type=str, help="Output JSON file")
    parser.add_argument("--config", type=str, default="metrics_config.json",
                        help="Path to the metrics config file (JSON format)")
    parser.add_argument("--ignore_qualities", action="store_true",
                    help="Skip processing qualities analysis")
    args = parser.parse_args()
    
    ticker = args.ticker.upper()
    # Load the metrics configuration file.
    try:
        with open(args.config, "r") as f:
            metrics_config = json.load(f)
    except Exception as e:
        logger.error(f"Error loading config file {args.config}: {e}")
        return  # Terminate if the config file cannot be loaded.
    
    # Look for a ticker-specific configuration.
    ticker_config = metrics_config.get(ticker)
    if not ticker_config:
        logger.error(f"No metrics configuration found for ticker {ticker} in {args.config}. Terminating.")
        return

    # Retrieve company profile using the utils function.
    profile = get_company_profile(ticker)
    if not profile:
        logger.error(f"No company profile found for {ticker}")
        return
    
    finder = EDGARExhibit13Finder(f"Insurance Research - Contact: {args.email}")
    try:
        extractor = MetricsExtractor(f"Insurance Research - Contact: {args.email}", config=ticker_config)
    except ValueError as ve:
        logger.error(f"Configuration error: {ve}")
        return
    
    try:
        cik = finder.get_cik_from_ticker(ticker)
        logger.info(f"Retrieved CIK for {ticker}: {cik}")
        filings_data = finder.get_company_filings(cik)
        if not filings_data.get("filings"):
            logger.error("No filings found.")
            return
        
        current_year = datetime.datetime.now().year
        start_year_int = int(args.start_year)
        years = list(range(int(args.start_year)-1, current_year + 1))
        # Pass the valid company profile instead of an empty dict.
        fmp_results = extract_yoy_data(ticker, years, segmentation_data={}, profile=profile)
        
        sec_results = {}
        for filing in filings_data["filings"]:
            if filing["filing_date"] < f"{args.start_year}-01-01":
                continue
            filing_docs = finder.get_filing_detail(filing["filing_href"])
            xml_url = filing_docs.get("xml")
            if xml_url:
                logger.info(f"Processing SEC XML filing dated {filing['filing_date']}")
                sec_year_data = extractor.extract_metrics(xml_url)
                for year_str, data in sec_year_data.items():
                    year_int = int(year_str)
                    if year_int not in sec_results:
                        sec_results[year_int] = data
                    else:
                        sec_results[year_int]["profit_desc"].update(data.get("profit_desc", {}))
                        for subsec in ["assets", "liabilities", "shareholders_equity"]:
                            sec_results[year_int]["balance_sheet"].setdefault(subsec, {}).update(
                                data.get("balance_sheet", {}).get(subsec, {}))
                        sec_results[year_int]["segmentation"].update(data.get("segmentation", {}))
            else:
                logger.warning(f"No XML filing document found for filing dated {filing['filing_date']}")
            time.sleep(0.1)
        
        # Only include years for which SEC data exists.
        unified_results = {}
        for year in sec_results.keys():
            fmp_year_data = fmp_results.get(year, {})
            sec_year_data = sec_results[year]
            unified_results[year] = create_unified_year_output(year, fmp_year_data, sec_year_data)
        
        # Now drop any years in unified_results before our desired start year.
        unified_results = {year: data for year, data in unified_results.items() if year >= start_year_int}
        # Determine the most recent year for which we have unified data.
        if unified_results:
            end_year = max(unified_results.keys())
        else:
            end_year = current_year  # Fallback if no data is available

        ordered_results = dict(sorted(unified_results.items(), reverse=True))
        
        # Invert the output structure so that the top level keys are the metric categories.
        inverted_output = {
            "company_description": {"data": {}},
            "analyses": {"data": {}},
            "balance_sheet": {"data": {}},
            "profit_description": {"data": {}},
            "segmentation": {"data": {}}
        }
        for year, metrics in ordered_results.items():
            inverted_output["company_description"]["data"][year] = metrics.get("company_description", {})
            inverted_output["analyses"]["data"][year] = metrics.get("analyses", {})
            inverted_output["balance_sheet"]["data"][year] = metrics.get("balance_sheet", {})
            inverted_output["profit_description"]["data"][year] = metrics.get("profit_description", {})
            inverted_output["segmentation"]["data"][year] = metrics.get("segmentation", {})
        
        # Compute the investment characteristics based on the unified year-by-year data.
        investment_characteristics = compute_investment_characteristics(unified_results)
        historical_pricing_averages = compute_historical_pricing_averages(unified_results)
        # Now insert the new section into the "analyses" top-level key.
        # (You can choose whether to place it at the same level as "data" or merge it with per‑year analyses.)
        inverted_output["analyses"]["investment_characteristics"] = investment_characteristics

        
        summary = {
            "symbol": profile.get("symbol"),
            "company_name": profile.get("companyName"),
            "exchange": profile.get("exchange"),
            "description": profile.get("description"),
            "sector": profile.get("sector"),
            "industry": profile.get("industry"),
            "reported_currency": profile.get("reported_currency") or profile.get("currency"),
            "isAdr": profile.get("isAdr")
        }
        # -------------------------------------------------------------------------
        # Retrieve the fiscal year end using the helper function (assumed to be imported)
        fiscal_year_end = get_fiscal_year_end(ticker)  # e.g., returns "12-31"
        # Get current stock price and market capitalization (assumed available via utils)
        stock_price = get_current_quote_yahoo(get_yahoo_ticker(profile))
        market_cap = get_current_market_cap_yahoo(get_yahoo_ticker(profile))
        
        # Compute the balance sheet characteristics using the inverted balance sheet data.
        balance_sheet_characteristics = compute_balance_sheet_characteristics(inverted_output["balance_sheet"]["data"])
        profit_description_characteristics = compute_profit_description_characteristics(unified_results)
        inverted_output["profit_description"]["profit_description_characteristics"] = profit_description_characteristics

                # Fetch industry comparison data using the helper function
        try:
            industry_data = get_industry_peers_with_stats(ticker)
        except Exception as e:
            logger.error(f"Error fetching industry data for {ticker}: {e}")
            industry_data = {}

        # Process qualities using the same function as in acm_analysis unless flag is set
        if not args.ignore_qualities:
            qualities = process_qualities(ticker, ignore_qualities=args.ignore_qualities, debug=False)
        else:
            qualities = ""

        # Construct final output including qualities
        final_output = {
            "summary": summary,
            "company_description": {
                "fiscal_year_end": fiscal_year_end,
                "stock_price": stock_price,
                "marketCapitalization": market_cap,
                "data": inverted_output["company_description"]["data"]
            },
            "analyses": inverted_output["analyses"],
            "balance_sheet": {
                "balance_sheet_characteristics": balance_sheet_characteristics,
                "data": inverted_output["balance_sheet"]["data"]
            },
            "profit_description": {
                "profit_description_characteristics": profit_description_characteristics,
                "data": inverted_output["profit_description"]["data"]
            },
            "segmentation": inverted_output["segmentation"],
            "historical_pricing": historical_pricing_averages,  
            "industry_comparison": industry_data,
            "qualities": qualities
        }

        os.makedirs("output", exist_ok=True)
        output_file = args.output or os.path.join("output", f"{ticker.lower()}_yoy_consolidated_bs.json")
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(final_output, f, indent=2)
        logger.info(f"Unified BS results saved to {output_file}")

        try:
            generate_excel_for_ticker_year(ticker, end_year)
        except Exception as e:
            print(f"Error generating Excel for {ticker} - {end_year}: {e}")

    except Exception as e:
        logger.error(f"Error: {e}")
        raise

if __name__ == "__main__":
    main()
