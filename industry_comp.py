import datetime
import os
import json
import requests
import logging
from dotenv import load_dotenv

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

def get_financial_data(ticker, api_key):
    """Fetch all required financial statements for a ticker."""
    logger.info(f"Fetching financial data for ticker: {ticker}")
    base_url = "https://financialmodelingprep.com/api/v3"
    
    endpoints = {
        "ic": f"{base_url}/income-statement/{ticker}",
        "bs": f"{base_url}/balance-sheet-statement/{ticker}",
        "cf": f"{base_url}/cash-flow-statement/{ticker}",
        "profile": f"{base_url}/profile/{ticker}"
    }
    
    data = {}
    for key, url in endpoints.items():
        logger.debug(f"Requesting data from endpoint: {key}")
        response = requests.get(url, params={"apikey": api_key})
        if response.status_code == 200:
            if key == "quote":
                data[key] = response.json()[0] if response.json() else None
            else:
                data[key] = response.json()[0] if response.json() else None
            logger.debug(f"Successfully retrieved {key} data")
        else:
            logger.error(f"Failed to fetch {key} data. Status code: {response.status_code}")
    
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
        ebit = ic.get('operatingIncome', 0)
        shareholder_equity = bs.get('totalStockholdersEquity', 0)
        
        logger.debug(f"Base metrics - Price: {current_stock_price}, Shares: {shares_outstanding}, Revenue: {revenues}")
        
        # Calculate liabilities and debt
        total_liabilities = bs.get('totalLiabilities', 0)
        capital_lease_obligations = bs.get('capitalLeaseObligations', 0)
        total_debt = total_liabilities - capital_lease_obligations
        
        lt_debt =  (bs.get("longTermDebt") or 0) + (bs.get("shortTermDebt") or 0) - (bs.get("capitalLeaseObligations") or 0)
        
        logger.debug(f"Debt calculations - Total debt: {total_debt}, LT debt: {lt_debt}")
        
        # Operating Statistics Calculations
        operating_margin = (ebit / revenues) if revenues else 0
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
        statement_year = int(statement_date.split('-')[0]) if statement_date else datetime.now().year
        
        yearly_high, yearly_low = get_yearly_high_low_yahoo(yahoo_ticker, statement_year)
        average_price = (yearly_high + yearly_low) / 2 if yearly_high is not None and yearly_low is not None else current_stock_price
        
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

def get_industry_peers_with_stats(ticker, num_comps=5, save_to_file=False):
    """Get industry peers and calculate statistics for all companies."""
    logger.info(f"Getting industry peers and stats for ticker: {ticker}")
    api_key = load_api_key()
    if not api_key:
        logger.error("API key not found")
        raise ValueError("API key not found")
    
    # First get company industry
    profile_url = f"https://financialmodelingprep.com/api/v3/profile/{ticker}"
    response = requests.get(profile_url, params={"apikey": api_key})
    if response.status_code != 200:
        logger.error(f"Error fetching profile. Status code: {response.status_code}")
        raise Exception(f"Error fetching profile: {response.status_code}")
    
    profile = response.json()[0]
    industry = profile.get("industry")
    logger.info(f"Industry identified: {industry}")
    
    # Get peers
    screener_url = "https://financialmodelingprep.com/api/v3/stock-screener"
    params = {
        "industry": industry,
        "isEtf": False,
        "isActivelyTrading": True,
        "apikey": api_key
    }
    
    response = requests.get(screener_url, params=params)
    if response.status_code != 200:
        logger.error(f"Error fetching peers. Status code: {response.status_code}")
        raise Exception(f"Error fetching peers: {response.status_code}")
    
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
    logger.info(f"Selected peers: {', '.join(top_peers)}")
    
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
        print(f"Analysis complete.")
        print("\nSummary:")
        print(json.dumps(result, indent=2))
        logger.info("Industry comparison analysis completed successfully")
    except Exception as e:
        logger.error(f"Industry comparison analysis failed: {str(e)}", exc_info=True)
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()