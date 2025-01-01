import datetime
import os
import json
import requests
from dotenv import load_dotenv
from acm_analysis import get_yearly_high_low_yahoo

def load_api_key():
    """Load the API key from the .env file."""
    load_dotenv()
    return os.getenv("FMP_API_KEY")

def get_financial_data(ticker, api_key):
    """Fetch all required financial statements for a ticker."""
    base_url = "https://financialmodelingprep.com/api/v3"
    
    endpoints = {
        "ic": f"{base_url}/income-statement/{ticker}",
        "bs": f"{base_url}/balance-sheet-statement/{ticker}",
        "cf": f"{base_url}/cash-flow-statement/{ticker}",
        "quote": f"{base_url}/quote-short/{ticker}"
    }
    
    data = {}
    for key, url in endpoints.items():
        response = requests.get(url, params={"apikey": api_key})
        if response.status_code == 200:
            if key == "quote":
                data[key] = response.json()[0] if response.json() else None
            else:
                data[key] = response.json()[0] if response.json() else None
    
    return data

def calculate_statistics(ticker, financial_data):
    """Calculate operating and market statistics for a ticker."""
    try:
        ic = financial_data.get('ic', {})
        bs = financial_data.get('bs', {})
        cf = financial_data.get('cf', {})
        quote = financial_data.get('quote', {})
        
        # Basic financial metrics
        current_stock_price = quote.get('price', 0)
        shares_outstanding = ic.get('weightedAverageShsOut', 0)
        revenues = ic.get('revenue', 0)
        net_profit = ic.get('netIncome', 0)
        ebit = ic.get('operatingIncome', 0)
        shareholder_equity = bs.get('totalStockholdersEquity', 0)
        
        # Calculate liabilities and debt
        total_liabilities = bs.get('totalLiabilities', 0)
        capital_lease_obligations = bs.get('capitalLeaseObligations', 0)
        total_debt = total_liabilities - capital_lease_obligations
        
        # Current liabilities
        current_liabilities = bs.get('totalCurrentLiabilities', 0)
        lt_debt = total_debt - current_liabilities
        bs_long_term_debt = bs.get('longTermDebt', 0)
        # Operating Statistics Calculations
        
        # Operating Margin = EBIT / Revenue
        operating_margin = (ebit / revenues) if revenues else 0
        
        # ROC = Net Profit / (Shareholder Equity + Long Term Debt)
        total_capital = shareholder_equity + bs_long_term_debt
        roc = (net_profit / total_capital) if total_capital else 0
        
        # Years Payback calculation
        cash_equiv = bs.get('cashAndCashEquivalents', 0)
        sti = bs.get('shortTermInvestments', 0)
        addback = cash_equiv + sti
        years_payback = ((lt_debt - addback) / net_profit) if net_profit else None
        
        # Market Statistics Calculations
        
        # Book value per share
        book_value_per_share = shareholder_equity / shares_outstanding if shares_outstanding else 0
        
        # P/B ratio
        pb_ratio = current_stock_price / book_value_per_share if book_value_per_share else 0
        
        # Operating EPS and P/E
        operating_eps = net_profit / shares_outstanding if shares_outstanding else 0
        pe_ratio = current_stock_price / operating_eps if operating_eps else 0
        
        # Get year from financial statements
        statement_date = ic.get('date')
        statement_year = int(statement_date.split('-')[0]) if statement_date else datetime.now().year
        
        # Get yearly high/low for average price calculation
        yearly_high, yearly_low = get_yearly_high_low_yahoo(ticker, statement_year)
        average_price = (yearly_high + yearly_low) / 2 if yearly_high is not None and yearly_low is not None else current_stock_price
        
        # Dividend Yield
        dividends_paid = -1 * cf.get('dividendsPaid', 0)
        dividends_per_share = dividends_paid / shares_outstanding if shares_outstanding else 0
        div_yield = (dividends_per_share / average_price) if average_price else 0
        
        # EV/Sales
        market_cap = current_stock_price * shares_outstanding
        ev_sales = ((market_cap + bs_long_term_debt) / revenues) if revenues else 0
        
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
        print(f"Error calculating statistics for {ticker}: {str(e)}")
        return None

def get_industry_peers_with_stats(ticker, num_comps=5, save_to_file = False):
    """Get industry peers and calculate statistics for all companies."""
    api_key = load_api_key()
    if not api_key:
        raise ValueError("API key not found")
    
    # First get company industry
    profile_url = f"https://financialmodelingprep.com/api/v3/profile/{ticker}"
    response = requests.get(profile_url, params={"apikey": api_key})
    if response.status_code != 200:
        raise Exception(f"Error fetching profile: {response.status_code}")
    
    profile = response.json()[0]
    industry = profile.get("industry")
    
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
        raise Exception(f"Error fetching peers: {response.status_code}")
    
    peers = response.json()
    sorted_peers = sorted(peers, key=lambda x: x.get('marketCap', 0), reverse=True)
    
    # Get the searched company's name
    searched_company = next((p for p in sorted_peers if p['symbol'] == ticker), None)
    company_name = searched_company.get('companyName') if searched_company else None
    
    # Deduplicate results by excluding the searched company by both symbol and name
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
    
    # Get top peers from deduplicated results
    top_peers = [stock['symbol'] for stock in unique_results[:num_comps]]
    all_tickers = [ticker] + top_peers
    
    # Calculate statistics for all companies
    result = {
        "operatingStatistics": {},
        "marketStatistics": {}
    }
    
    for t in all_tickers:
        financial_data = get_financial_data(t, api_key)
        stats = calculate_statistics(t, financial_data)
        if stats:
            result["operatingStatistics"].update(stats["operatingStatistics"])
            result["marketStatistics"].update(stats["marketStatistics"])
    
    if save_to_file:
        # Save to JSON file
        output_file = f"{ticker}_peer_analysis.json"
        with open(output_file, 'w') as f:
            json.dump(result, f, indent=4)
    
    return result

def main():
    ticker = input("Enter a ticker symbol: ").strip().upper()
    if not ticker:
        print("Ticker symbol is required.")
        return
    
    try:
        result = get_industry_peers_with_stats(ticker)
        print(f"Analysis complete. Results saved to {ticker}_peer_analysis.json")
        print("\nSummary:")
        print(json.dumps(result, indent=2))
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()