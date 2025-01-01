# ACM Analysis Project

This project provides a comprehensive suite of financial analysis tools for stock market research, combining data from multiple sources including industry comparisons, forum posts, email analysis, and more.

## Features

- Industry comparison analysis with peer companies
- Historical stock data retrieval from Yahoo Finance
- Forum post collection and analysis
- Email analysis for specific stock tickers
- OpenAI-powered summarization of collected data
- Automated Excel report generation

## Prerequisites

- Python 3.7 or newer
- Microsoft Outlook (for email analysis features)
- Windows OS (for setup.bat and Outlook integration)
- Active API keys for:
  - Financial Modeling Prep (FMP)
  - Website Toolbox
  - OpenAI
  - Other service-specific requirements

## Installation

1. Clone the repository:

```bash
git clone [repository-url]
cd analysis_project
```

2. Create a `.env` file in the root directory with the following content:

```
WEBSITETOOLBOX_API_KEY='your_websitetoolbox_api_key'
WEBSITETOOLBOX_USERNAME='your_websitetoolbox_username'
FMP_API_KEY='your_fmp_api_key'
OPENAI_API_KEY='your_openai_api_key'
SENDER_EMAIL='your_email@domain.com'
```

3. Run the setup script:

```bash
setup.bat
```

This will:

- Create a virtual environment
- Install required dependencies
- Set up the initial configuration

## Project Structure

```
analysis_project/
├── acm_analysis.py         # Main entry point
├── industry_comp.py        # Industry comparison functionality
├── yahoo_finance_utils.py  # Yahoo Finance data retrieval
├── forum_posts.py          # Forum post collection
├── forum_post_summary.py   # Forum post analysis and summarization
├── outlook_ticker_search.py # Email analysis for tickers
├── requirements.txt        # Project dependencies
├── setup.bat              # Windows setup script
└── README.md              # This file
```

## Usage

The main entry point is `acm_analysis.py`. Run it with a ticker symbol:

```bash
python acm_analysis.py TICKER
```

### Individual Components

Each component can also be run independently:

- Industry comparison:

```bash
python industry_comp.py TICKER
```

- Forum posts analysis:

```bash
python forum_posts.py TICKER
```

- Email analysis:

```bash
python outlook_ticker_search.py TICKER
```

## API Keys and Configuration

### Required API Keys

1. **Financial Modeling Prep (FMP)**
   - Sign up at [financialmodelingprep.com](https://financialmodelingprep.com)
   - Used for financial data retrieval

2. **Website Toolbox**
   - Obtain from your Website Toolbox admin panel
   - Used for forum data collection

3. **OpenAI**
   - Sign up at [platform.openai.com](https://platform.openai.com)
   - Used for content summarization

### Environment Variables

Create a `.env` file with the following structure:

```
WEBSITETOOLBOX_API_KEY='x8fHbXqWYl8J'
WEBSITETOOLBOX_USERNAME='your_username'
FMP_API_KEY="your_fmp_key"
OPENAI_API_KEY='your_openai_key'
SENDER_EMAIL='your_email'
```

## Output

The analysis results are stored in the `output` directory:

- Industry comparison data: `{ticker}_peer_analysis.json`
- Forum posts: `{ticker}_posts.json`
- Email analysis: `{ticker}_sent_emails.json`
- Summaries: `{ticker}_post_summary.txt`

## Dependencies

Key dependencies include:

- yfinance: Yahoo Finance data retrieval
- pandas: Data manipulation
- openai: GPT integration
- win32com: Outlook integration
- beautifulsoup4: HTML parsing
- requests: API calls

For a complete list, see `requirements.txt`.

## Error Handling

The project includes comprehensive logging throughout all components. Check the console output for any errors during execution. Common issues include:

- Invalid API keys
- Network connectivity problems
- Missing Outlook configuration
- Invalid ticker symbols

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request

## License

[Your chosen license]

## Acknowledgments

- Financial Modeling Prep for financial data
- Website Toolbox for forum integration
- OpenAI for GPT integration
- Yahoo Finance for historical data
