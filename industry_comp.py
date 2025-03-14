# analysis_project/industry_comp.py
import datetime
import os
import json
import time
import requests
import logging
from dotenv import load_dotenv
from requests.exceptions import Timeout, RequestException

from utils import get_current_quote_yahoo, get_yahoo_ticker, get_yearly_high_low_yahoo

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def load_api_key():
    """Load the API key from the .env file."""
    logger.info("Loading API key from .env file")
    load_dotenv()
    api_key = os.getenv("FMP_API_KEY")
    if api_key:
        logger.info("API key loaded successfully")
    else:
        logger.error("Failed to load API key")
    return api_key

def fetch_with_retry(url, params=None, timeout=10, retries=3):
    """
    Helper function to fetch data from a URL with a specified number of retries.
    If the call fails due to a timeout or other RequestException, it will be retried.
    """
    for attempt in range(retries):
        try:
            logger.debug(f"Attempt {attempt + 1} for URL: {url}")
            response = requests.get(url, params=params, timeout=timeout)
            response.raise_for_status()
            return response
        except (Timeout, RequestException) as e:
            logger.error(f"Attempt {attempt + 1} of {retries} failed for URL {url}: {e}")
            if attempt < retries - 1:
                time.sleep(2)  # wait 2 seconds before retrying
    # If all attempts fail, raise an exception
    raise RequestException(f"All {retries} attempts failed for URL {url}")

def get_financial_data(ticker, api_key):
    """Fetch all required financial statements for a ticker with retries."""
    logger.info(f"Fetching financial data for ticker: {ticker}")
    base_url = "https://financialmodelingprep.com/api/v3"
    
    endpoints = {
        "ic": f"{base_url}/income-statement/{ticker}",
        "bs": f"{base_url}/balance-sheet-statement/{ticker}",
        "cf": f"{base_url}/cash-flow-statement/{ticker}",
        "profile": f"{base_url}/profile/{ticker}"
    }
    
    data = {}
    timeout_value = 10
    for key, url in endpoints.items():
        logger.debug(f"Requesting {key} data from endpoint: {url}")
        try:
            response = fetch_with_retry(url, params={"apikey": api_key}, timeout=timeout_value, retries=3)
            json_data = response.json()
            # For this example, we assume each endpoint returns a list.
            data[key] = json_data[0] if json_data else None
            logger.debug(f"Successfully retrieved {key} data")
        except RequestException as e:
            logger.error(f"Failed to fetch {key} data for ticker {ticker} after retries: {e}")
            data[key] = None
    return data

def calculate_statistics(ticker, financial_data):
    """Calculate operating and market statistics for a ticker."""
    logger.info(f"Calculating statistics for ticker: {ticker}")
    try:
        ic = financial_data.get('ic', {})
        bs = financial_data.get('bs', {})
        cf = financial_data.get('cf', {})
        profile = financial_data.get('profile', {})
        
        yahoo_ticker = get_yahoo_ticker(profile)
        current_stock_price = get_current_quote_yahoo(yahoo_ticker)
        
        # Basic financial metrics
        shares_outstanding = ic.get('weightedAverageShsOut', 0)
        revenues = ic.get('revenue', 0)
        net_profit = ic.get('netIncome', 0)
        cost_of_revenue = ic.get('costOfRevenue') or 0
        cost_of_res_and_dev = ic.get('researchAndDevelopmentExpenses') or 0
        cost_of_selling_and_marketing_gen_and_admin = ic.get('sellingGeneralAndAdministrativeExpenses') or 0
        expenses = cost_of_selling_and_marketing_gen_and_admin + cost_of_res_and_dev + cost_of_revenue
        shareholder_equity = bs.get('totalStockholdersEquity', 0)
        logger.debug(f"Base metrics - Price: {current_stock_price}, Shares: {shares_outstanding}, Revenue: {revenues}")
        
        # Calculate liabilities and debt
        total_liabilities = bs.get('totalLiabilities', 0)
        capital_lease_obligations = bs.get('capitalLeaseObligations', 0)
        total_debt = total_liabilities - capital_lease_obligations
        lt_debt = (bs.get("longTermDebt") or 0) + (bs.get("shortTermDebt") or 0) - (bs.get("capitalLeaseObligations") or 0)
        logger.debug(f"Debt calculations - Total debt: {total_debt}, LT debt: {lt_debt}")
        
        # Operating Statistics Calculations
        operating_margin = (revenues - expenses) / revenues if revenues else None

        total_capital = shareholder_equity + lt_debt
        roc = (net_profit / total_capital) if total_capital else 0
        
        cash_equiv = bs.get('cashAndCashEquivalents', 0)
        sti = bs.get('shortTermInvestments', 0)
        addback = cash_equiv + sti
        years_payback = ((lt_debt - addback) / net_profit) if net_profit else None
        logger.debug(f"Operating metrics - Margin: {operating_margin}, ROC: {roc}, Years payback: {years_payback}")
        
        # Market Statistics Calculations
        book_value_per_share = shareholder_equity / shares_outstanding if shares_outstanding else 0
        pb_ratio = current_stock_price / book_value_per_share if book_value_per_share else 0
        operating_eps = net_profit / shares_outstanding if shares_outstanding else 0
        pe_ratio = current_stock_price / operating_eps if operating_eps else 0
        
        statement_date = ic.get('date')
        statement_year = int(statement_date.split('-')[0]) if statement_date else datetime.datetime.now().year
        
        yearly_high, yearly_low = get_yearly_high_low_yahoo(yahoo_ticker, statement_year)
        average_price = (yearly_high + yearly_low) / 2 if (yearly_high is not None and yearly_low is not None) else current_stock_price
        
        dividends_paid = -1 * cf.get('dividendsPaid', 0)
        dividends_per_share = dividends_paid / shares_outstanding if shares_outstanding else 0
        div_yield = (dividends_per_share / average_price) if average_price else 0
        
        market_cap = current_stock_price * shares_outstanding
        ev_sales = ((market_cap + lt_debt) / revenues) if revenues else 0
        
        logger.debug(f"Market metrics - P/B: {pb_ratio}, P/E: {pe_ratio}, Div Yield: {div_yield}")
        
        return {
            "operatingStatistics": {
                ticker: {
                    "Debt(yrs.)": round(years_payback, 1) if years_payback else None,
                    "Sales": revenues,
                    "ROC": roc,
                    "Operating Margin": operating_margin
                }
            },
            "marketStatistics": {
                ticker: {
                    "P/B": round(pb_ratio, 2),
                    "P/E": round(pe_ratio, 1),
                    "Div. Yld.": div_yield,
                    "EV/Sales": round(ev_sales, 2)
                }
            }
        }
    except Exception as e:
        logger.error(f"Error calculating statistics for {ticker}: {str(e)}", exc_info=True)
        return None

def check_adr_mapping(ticker, peers, adr_mapping):
    """
    Check if any peer tickers map to the same ordinary shares as the input ticker.
    Remove duplicates, keeping the ADR ticker when found.
    """
    # Create reverse mapping from ordinary to ADR
    ord_to_adr = {v: k for k, v in adr_mapping.items()}
    
    filtered_peers = []
    for peer in peers:
        if peer in ord_to_adr and ord_to_adr[peer] == ticker:
            continue
        filtered_peers.append(peer)
    return filtered_peers

def get_industry_peers_with_stats(ticker, num_comps=5, save_to_file=False):
    """Get industry peers and calculate statistics for all companies."""
    logger.info(f"Getting industry peers and stats for ticker: {ticker}")
    api_key = load_api_key()
    if not api_key:
        logger.error("API key not found")
        raise ValueError("API key not found")
    
    # Load ADR mapping
    with open('adr_to_ord_mapping.json', 'r') as f:
        adr_mapping = json.load(f)
    
    # Get company profile
    profile_url = f"https://financialmodelingprep.com/api/v3/profile/{ticker}"
    try:
        response = fetch_with_retry(profile_url, params={"apikey": api_key}, timeout=10, retries=3)
    except RequestException as e:
        logger.error(f"Error fetching profile for {ticker}: {e}")
        raise

    profile = response.json()[0]
    industry = profile.get("industry")
    logger.info(f"Industry identified: {industry}")
    
    # Get peers via the screener endpoint
    screener_url = "https://financialmodelingprep.com/api/v3/stock-screener"
    params = {
        "industry": industry,
        "isEtf": False,
        "isActivelyTrading": True,
        "apikey": api_key
    }
    try:
        response = fetch_with_retry(screener_url, params=params, timeout=10, retries=3)
    except RequestException as e:
        logger.error(f"Error fetching peers for {ticker}: {e}")
        raise

    peers = response.json()
    sorted_peers = sorted(peers, key=lambda x: x.get('marketCap', 0), reverse=True)
    
    searched_company = next((p for p in sorted_peers if p['symbol'] == ticker), None)
    company_name = searched_company.get('companyName') if searched_company else None
    logger.info(f"Found company: {company_name}")
    
    unique_results = []
    seen_names = set()
    for stock in sorted_peers:
        stock_name = stock.get("companyName")
        stock_symbol = stock.get("symbol")
        if (stock_name and 
            stock_symbol != ticker and 
            stock_name != company_name and 
            stock_name not in seen_names):
            unique_results.append(stock)
            seen_names.add(stock_name)
    
    top_peers = [stock['symbol'] for stock in unique_results[:num_comps]]
    top_peers = check_adr_mapping(ticker, top_peers, adr_mapping)
    logger.info(f"Selected peers after ADR check: {', '.join(top_peers)}")
    
    all_tickers = [ticker] + top_peers
    result = {
        "operatingStatistics": {},
        "marketStatistics": {}
    }
    
    for t in all_tickers:
        logger.info(f"Processing ticker: {t}")
        financial_data = get_financial_data(t, api_key)
        stats = calculate_statistics(t, financial_data)
        if stats:
            result["operatingStatistics"].update(stats["operatingStatistics"])
            result["marketStatistics"].update(stats["marketStatistics"])
        else:
            logger.warning(f"No statistics calculated for ticker: {t}")
    
    if save_to_file:
        output_file = os.path.join('output', f'{ticker}_peer_analysis.json')
        logger.info(f"Saving results to file: {output_file}")
        with open(output_file, 'w') as f:
            json.dump(result, f, indent=4)
    
    return result

def main():
    ticker = input("Enter a ticker symbol: ").strip().upper()
    if not ticker:
        logger.error("No ticker symbol provided")
        print("Ticker symbol is required.")
        return
    
    logger.info(f"Starting industry comparison analysis for ticker: {ticker}")
    try:
        result = get_industry_peers_with_stats(ticker)
        print("Analysis complete.\nSummary:")
        print(json.dumps(result, indent=2))
        logger.info("Industry comparison analysis completed successfully")
    except Exception as e:
        logger.error(f"Industry comparison analysis failed: {str(e)}", exc_info=True)
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
