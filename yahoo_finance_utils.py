import yfinance as yf

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