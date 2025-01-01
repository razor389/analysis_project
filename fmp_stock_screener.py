import os
import requests
from dotenv import load_dotenv

# Global parameter for the number of results to display
NUM_RESULTS = 5

def load_api_key():
    """Load the API key from the .env file."""
    load_dotenv()
    return os.getenv("FMP_API_KEY")

def search_ticker_by_name(company_name, exchange="NASDAQ", limit=10):
    """Search for a ticker symbol by company name."""
    base_url = "https://financialmodelingprep.com/api/v3/search-name"
    api_key = load_api_key()
    if not api_key:
        raise ValueError("API key not found. Please add it to your .env file under the key 'FMP_API_KEY'.")

    params = {
        "query": company_name,
        "limit": limit,
        "exchange": exchange,
        "apikey": api_key
    }

    response = requests.get(base_url, params=params)
    if response.status_code != 200:
        raise Exception(f"Error: {response.status_code} - {response.text}")

    data = response.json()
    if data:
        return data  # Return all matching results
    return []

def get_stock_screener_results(params):
    """Call the FMP API stock screener endpoint with the given parameters."""
    base_url = "https://financialmodelingprep.com/api/v3/stock-screener"
    api_key = load_api_key()
    if not api_key:
        raise ValueError("API key not found. Please add it to your .env file under the key 'FMP_API_KEY'.")

    # Append API key to parameters
    params['apikey'] = api_key

    response = requests.get(base_url, params=params)
    if response.status_code != 200:
        raise Exception(f"Error: {response.status_code} - {response.text}")

    return response.json()


def get_company_sector_and_industry(ticker):
    """Fetch sector and industry information for a given ticker."""
    base_url = f"https://financialmodelingprep.com/api/v3/profile/{ticker}"
    api_key = load_api_key()
    if not api_key:
        raise ValueError("API key not found. Please add it to your .env file under the key 'FMP_API_KEY'.")

    response = requests.get(base_url, params={"apikey": api_key})
    if response.status_code != 200:
        raise Exception(f"Error: {response.status_code} - {response.text}")

    data = response.json()
    if data:
        profile = data[0]
        return profile.get("companyName"), profile.get("sector"), profile.get("industry")
    return None, None, None


def get_industry_peers(ticker, searched_name):
    """Get industry peers for the given ticker, sorted by market cap."""
    try:
        # Get company details and industry
        company_name, sector, industry = get_company_sector_and_industry(ticker)
        if not industry:
            print(f"Industry information not found for ticker {ticker}.")
            return

        print(f"Fetching peers for industry: {industry}...")
        params = {
            "isEtf": False,
            "isFund": False,
            "isActivelyTrading": True,
            "industry": industry
        }

        # Fetch peers, including the searched company
        results = get_stock_screener_results(params)
        sorted_results = sorted(results, key=lambda x: x.get('marketCap', 0), reverse=True)

        # Find the searched company's market cap from the peers list
        searched_market_cap = None
        for stock in sorted_results:
            if stock.get("symbol") == ticker:
                searched_market_cap = stock.get("marketCap")
                break

        # Display the searched company details
        print(f"Searched Company: Symbol: {ticker}, Name: {company_name}, Market Cap: {searched_market_cap or 'Unavailable'}")

        # Deduplicate results by excluding the searched company by both symbol and name
        unique_results = []
        seen_names = set()
        for stock in sorted_results:
            stock_name = stock.get("companyName")
            stock_symbol = stock.get("symbol")
            if stock_name and stock_symbol != ticker and stock_name != company_name and stock_name not in seen_names:
                unique_results.append(stock)
                seen_names.add(stock_name)

        # Display top peers based on NUM_RESULTS
        print(f"\nTop {NUM_RESULTS} unique peers for {ticker} ({industry}):")
        for stock in unique_results[:NUM_RESULTS]:  # Use NUM_RESULTS for slicing
            print(f"Symbol: {stock.get('symbol')}, Name: {stock.get('companyName')}, Market Cap: {stock.get('marketCap')}")
    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    """Main function to execute the script."""
    ticker = input("Enter a ticker symbol: ").strip().upper()
    if not ticker:
        print("Ticker symbol is required.")
        return

    # Pass the ticker directly to the function
    get_industry_peers(ticker, searched_name=None)


if __name__ == "__main__":
    main()
