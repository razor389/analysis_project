import json
import requests
import yfinance as yf
from dotenv import load_dotenv
import datetime
import os
import sys

# Load the .env file
load_dotenv()

# Retrieve FMP_API_KEY from environment
FMP_API_KEY = os.getenv("FMP_API_KEY")
if not FMP_API_KEY:
    print("ERROR: FMP_API_KEY not found in environment. Please set FMP_API_KEY in your .env file.")
    sys.exit(1)

def get_company_profile(symbol: str):
    """Fetch the company's profile from FMP."""
    url = f"https://financialmodelingprep.com/api/v3/profile/{symbol}?apikey={FMP_API_KEY}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        if isinstance(data, list) and len(data) > 0:
            return data[0]  # Take the first profile item
        return {}
    except Exception as e:
        print(f"Error fetching company profile for {symbol}: {e}")
        return {}

def get_reported_currency(symbol: str):
    """
    Fetch the reported currency from the most recent balance sheet statement for a given symbol.
    
    Args:
        symbol (str): The stock symbol (e.g., 'ASML')
        
    Returns:
        str: The reported currency (e.g., 'EUR') or empty string if not found
    """
    try:
        url = f"https://financialmodelingprep.com/api/v3/balance-sheet-statement/{symbol}?period=annual&apikey={FMP_API_KEY}"
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        
        if isinstance(data, list) and len(data) > 0:
            # Get the most recent statement (first item in the list)
            return data[0].get('reportedCurrency', '')
        return ''
        
    except Exception as e:
        print(f"Error fetching reported currency for {symbol}: {e}")
        return ''

def get_yahoo_ticker(profile: dict) -> str:
    """
    Determine the appropriate Yahoo Finance ticker based on company profile.
    For ADRs, attempts to convert to the ordinary shares ticker if mapping exists.
    """
    if not profile:
        raise ValueError("Invalid company profile")
        
    symbol = profile.get("symbol")
    if not symbol:
        raise ValueError("No symbol found in company profile")
    
    # Check if it's an ADR
    if profile.get("isAdr", False):
        try:
            # Look for mapping file in current directory
            with open("adr_to_ord_mapping.json", "r") as f:
                mapping = json.load(f)
                
            if symbol in mapping:
                return mapping[symbol]
            else:
                print(f"Warning: No ordinary shares mapping found for ADR {symbol}")
                return symbol
        except Exception as e:
            print(f"Error reading ADR mapping file: {e}")
            return symbol
    
    return symbol

def get_current_market_cap_yahoo(symbol: str) -> float:
    """
    Calculate current market cap using Yahoo Finance data (price * shares outstanding).
    
    Args:
        symbol (str): The stock symbol
        
    Returns:
        float: Current market cap in the trading currency, or None if there's an error
    """
    try:
        ticker = yf.Ticker(symbol)
        info = ticker.info
        
        # Get current price and shares outstanding
        current_price = info.get('currentPrice')
        shares_outstanding = info.get('sharesOutstanding')
        
        if current_price is None or shares_outstanding is None:
            print(f"Error: Missing price or shares data for {symbol}")
            return None
            
        market_cap = current_price * shares_outstanding
        return market_cap
        
    except Exception as e:
        print(f"Error calculating market cap for {symbol}: {e}")
        return None

def get_yearly_high_low_yahoo(symbol: str, year: int):
    """
    Fetch daily stock data from Yahoo Finance for the given year and return the yearly high and low.
    """
    start_date = f"{year}-01-01"
    end_date = f"{year}-12-31"
    df = yf.download(symbol, start=start_date, end=end_date, progress=False)
    if df.empty:
        return None, None

    yearly_high = df['High'].max()
    yearly_low = df['Low'].min()

    # Convert from numpy floats to Python floats
    if yearly_high is not None:
        yearly_high = yearly_high.item()
    if yearly_low is not None:
        yearly_low = yearly_low.item()

    return yearly_high, yearly_low

def get_current_quote_yahoo(symbol: str) -> float:
    """
    Fetch current stock price from Yahoo Finance.
    Returns None if there's an error.
    """
    try:
        ticker = yf.Ticker(symbol)
        info = ticker.info
        return info.get('currentPrice')
    except Exception as e:
        print(f"Error fetching current quote for {symbol}: {e}")
        return None
    
def get_long_term_rate():
    """
    Fetch the long-term rates from the backend API and return the 20-year bond yield.
    
    Expects the environment variable BACKEND_URL to be set.
    
    Returns:
        float: The bond_yield_20y value, or None if an error occurs.
    """
    backend_url = os.getenv("BACKEND_URL")
    if not backend_url:
        print("Error: BACKEND_URL environment variable not set.")
        return None

    url = f"{backend_url}/api/v1/long_term_rates"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        # Assuming the API returns a structure like: {"rates": {"bond_yield_20y": 3.5, ...}}
        rates = data.get("rates", {})
        bond_yield_20y = rates.get("bond_yield_20y")
        return (bond_yield_20y/100)
    except Exception as e:
        print("Error fetching long-term rates:", e)
        return None
    
def get_eoy_fx_rate(reporting_currency: str,
                    base_currency: str,
                    year: int) -> float:
    """
    Fetch the FX close on the last trading day of December `year`
    for the pair REPORTING/BASE (e.g. EUR/USD) by using yfinance.
    Returns the 'Close' price for the latest available date in Dec.
    """
    # Yahoo ticker for FX pair is e.g. "EURUSD=X"
    pair = f"{reporting_currency}{base_currency}=X"
    try:
        ticker = yf.Ticker(pair)
        # try just December data
        start = f"{year}-12-01"
        end   = f"{year}-12-31"
        df = ticker.history(start=start, end=end, auto_adjust=False)
        if not df.empty:
            # last available close in December
            return df['Close'].iloc[-1]
        # fallback: grab 1-year history and take last
        df = ticker.history(period="1y", auto_adjust=False)
        if not df.empty:
            return df['Close'].iloc[-1]
    except Exception as e:
        print(f"Error fetching EOY FX rate {reporting_currency}/{base_currency} {year}: {e}")
    return None