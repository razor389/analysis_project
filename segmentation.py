import sys
import json
from typing import Dict, List, Optional, Union
from datetime import datetime
import os
from dotenv import load_dotenv
import math
import requests
from bs4 import BeautifulSoup
import re
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()
FMP_API_KEY = os.getenv('FMP_API_KEY')

HEADERS = {
    "User-Agent": "Custom Research Agent - Contact: rgranowski@gmail.com" 
}

def get_financial_statement(symbol: str, year: int) -> dict:
    """
    Fetch the balance sheet statement for the given symbol and year.
    Returns the filing data including the SEC document link.
    """
    url = f"https://financialmodelingprep.com/api/v3/balance-sheet-statement/{symbol}"
    params = {
        "period": "annual",
        "apikey": FMP_API_KEY
    }
    
    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        statements = response.json()
        
        # Find the statement for the specified year
        target_statement = next(
            (stmt for stmt in statements 
             if int(stmt["calendarYear"]) == year),
            None
        )
        
        if not target_statement:
            raise ValueError(f"No financial statement found for {symbol} in {year}")
            
        return target_statement
        
    except Exception as e:
        logger.error(f"Error fetching financial statement: {e}")
        return None

def get_filing_url(symbol: str, year: int) -> str:
    """
    Get the iXBRL filing URL for the given symbol and year.
    """
    statement = get_financial_statement(symbol, year)
    if not statement:
        return None
        
    # Get the final link and transform it to ix?doc format
    final_link = statement.get("finalLink")
    if not final_link:
        logger.error("No filing link found in statement data")
        return None
        
    transformed_url = final_link.replace("/Archives/", "/ix?doc=/Archives/")
    return transformed_url

def get_filing_metadata(url):
    """Extract the CIK, accession number and filename from the URL."""
    match = re.search(r'/Archives/edgar/data/(\d+)/(\d+)/([^/]+)$', url)
    if not match:
        raise ValueError("Invalid SEC URL format")
    cik = match.group(1)
    accession_number = match.group(2)
    filename = match.group(3)
    return cik, accession_number.replace("-", ""), filename

def get_filing_contents(url):
    """Get the full filing contents using SEC's data endpoints."""
    logger.info("Fetching filing data...")
    
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        
        cik, accession_number, filename = get_filing_metadata(url)
        
        base_url = "https://www.sec.gov/Archives/edgar/data"
        filing_url = f"{base_url}/{cik}/{accession_number}/{filename}"
        meta_url = f"{base_url}/{cik}/{accession_number}/MetaLinks.json"
        
        filing_response = requests.get(filing_url, headers=HEADERS)
        filing_response.raise_for_status()
        
        meta_response = requests.get(meta_url, headers=HEADERS)
        meta_response.raise_for_status()
        
        return filing_response.text, meta_response.json()
        
    except requests.exceptions.RequestException as e:
        logger.error(f"Error fetching filing: {e}")
        return None, None

def format_axis_name(axis_name):
    """Format axis name to standard format."""
    # Remove any namespace prefix
    if ':' in axis_name:
        axis_name = axis_name.split(':')[-1]
    
    # Handle common axis names
    if axis_name == "StatementBusinessSegmentsAxis":
        return "US-GAAP Statement Business Segments Axis"
    
    # For other axis names, add spaces before capital letters
    formatted = re.sub(r'(?<!^)(?=[A-Z])', ' ', axis_name)
    if "Axis" not in formatted:
        formatted += " Axis"
    return formatted

def format_member_name(member_value, company_prefix=None):
    """Format member name to standard format."""
    if ':' not in member_value:
        return member_value
    
    prefix, member = member_value.split(':')
    
    # Add spaces before capital letters for readability
    formatted = re.sub(r'(?<!^)(?=[A-Z])', ' ', member)
    
    # If we have a company prefix, add it to the member name
    if company_prefix:
        formatted = f"{company_prefix}{formatted}"
    
    if "Member" not in formatted:
        formatted += " Member"
    
    return formatted

def get_scale_factor(decimals, scale):
    """Determine the scale factor based on decimals and scale attributes."""
    if scale:
        if int(scale) == 6:
            return "Millions"
        elif int(scale) == 3:
            return "Thousands"
        else:
            return "Units" 
    elif decimals:
        decimals = int(decimals)
        if decimals == -6:
            return "Millions"
        elif decimals == -3:
            return "Thousands"
    return "Units"

def parse_context(soup, context_ref):
    """Extract period and dimensional information from context."""
    context = soup.find(attrs={"id": context_ref})
    if not context:
        return {}
    
    context_data = {}
    
    # Extract period information
    period = context.find('xbrli:period') or context.find('period')
    if period:
        instant = period.find('xbrli:instant') or period.find('instant')
        if instant:
            context_data['period'] = f"As of {instant.text}"
        else:
            start_date = period.find('xbrli:startdate') or period.find('startdate')
            end_date = period.find('xbrli:enddate') or period.find('enddate')
            if start_date and end_date:
                start = datetime.strptime(start_date.text, '%Y-%m-%d')
                end = datetime.strptime(end_date.text, '%Y-%m-%d')
                days = (end - start).days
                months = math.ceil(days / 30.44)  # Round up to nearest month
                end_formatted = end.strftime('%m/%d/%Y')  # Format as MM/DD/YYYY
                context_data['period'] = f"{months} months ending {end_formatted}"
    
    # Extract entity and segment information
    entity = context.find('xbrli:entity') or context.find('entity')
    if entity:
        identifier = entity.find('xbrli:identifier') or entity.find('identifier')
        if identifier:
            context_data['entity_id'] = identifier.text
            context_data['entity_scheme'] = identifier.get('scheme', '')

        segment = entity.find('xbrli:segment') or entity.find('segment')
        if segment:
            explicit_members = segment.find_all(['xbrldi:explicitmember', 'explicitmember', 'xbrldi:typedmember', 'typedmember'], recursive=True)
            
            # Initialize lists to collect multiple axes and members
            axes = []
            members = []
            explicit_members_list = []
            
            for dimension in explicit_members:
                dimension_name = dimension.get('dimension', '')
                if dimension_name:
                    axes.append(format_axis_name(dimension_name))
                
                member_value = dimension.text.strip()
                if ':' in member_value:
                    company_prefix = member_value.split(':')[0].upper()
                    members.append(format_member_name(member_value, company_prefix))
                    explicit_members_list.append(member_value)
            
            # Join multiple axes and members with newlines if there are multiple
            if axes:
                context_data['axis'] = '\n'.join(axes)
            if members:
                context_data['member'] = '\n'.join(members)
                context_data['explicit_member'] = '\n'.join(explicit_members_list)

    return context_data

def extract_inline_xbrl_data(url, target_tag):
    """Extract Inline XBRL data from SEC filings for a specific tag."""
    content, meta_info = get_filing_contents(url)
    if not content:
        return []
    
    logger.info("Parsing document...")
    soup = BeautifulSoup(content, 'html.parser')
    
    xbrl_elements = []
    for element in soup.find_all(True):
        if element.get('name') == target_tag or element.get('data-xlinkLabel') == target_tag:
            xbrl_elements.append(element)
    
    results = []
    for elem in xbrl_elements:
        #print(f"elem: {elem}")
        raw_value = elem.get_text(strip=True)
        #print(f"value: {raw_value}")
        decimals = elem.get("decimals")
        scale = elem.get("scale")
        format_attr = elem.get("format", "")
        sign_attr = elem.get("sign")  # Get the explicit sign attribute
        
        if format_attr.startswith("ixt:"):
            format_attr = format_attr[4:]  # Remove ixt: prefix
        
        # Process the numeric value
        try:
            # Remove commas and convert to float
            numeric_value = float(raw_value.replace(",", ""))
            
            # Apply sign based on both the sign attribute and any minus prefix
            if sign_attr == "-" or raw_value.startswith("-"):
                numeric_value = -abs(numeric_value)  # Ensure negative
            
            # Apply scaling
            scale_factor = 1000000 if get_scale_factor(decimals, scale) == "Millions" else \
                          1000 if get_scale_factor(decimals, scale) == "Thousands" else 1
            scaled_value = numeric_value * scale_factor
            
            # Format with commas, handling negative values
            formatted_value = f"{scaled_value:,.0f}"
            
        except (ValueError, TypeError):
            formatted_value = raw_value
        
        # Build the fact dictionary
        fact = {
            "tag": target_tag,
            "fact": formatted_value,
            "scale": get_scale_factor(decimals, scale),
            "decimals": get_scale_factor(decimals, scale),
            "format": format_attr,
            "sign": "Negative" if (sign_attr == "-" or raw_value.startswith("-")) else "Positive",
            "type": "Monetary Item Type" if elem.get("unitref") or elem.get("unitRef") else "String Type"
        }
        
        # Rest of the code remains the same...
        unit_ref = elem.get("unitref") or elem.get("unitRef")
        if unit_ref:
            fact["unit_ref"] = unit_ref
            unit = soup.find(id=unit_ref)
            if unit:
                measure = unit.find(['measure', 'xbrli:measure'])
                if measure:
                    fact["measure"] = measure.text.split(':')[-1]
        
        context_ref = elem.get("contextref") or elem.get("contextRef")
        if context_ref:
            context_data = parse_context(soup, context_ref)
            fact.update(context_data)
        
        if 'IncomeLoss' in target_tag:
            fact["balance"] = "Credit"
        
        fact = {k: v for k, v in fact.items() if v is not None and k not in ['entity_id', 'entity_scheme']}
        
        results.append(fact)
    
    return results

def load_config(ticker: str) -> Dict:
    """Load the configuration for a specific ticker."""
    try:
        with open('segmentation_config.json', 'r') as f:
            config = json.load(f)
        
        if ticker not in config:
            raise ValueError(f"No configuration found for ticker {ticker}")
        
        logger.info(f"Loaded configuration for {ticker}:")
        logger.info(json.dumps(config[ticker], indent=2))
        return config[ticker]
    except FileNotFoundError:
        raise FileNotFoundError("segmentation_config.json not found")
    except json.JSONDecodeError:
        raise ValueError("Invalid JSON in configuration file")

def filter_facts(facts: List[Dict], axes: Optional[Union[List[str], str]], year: int) -> List[Dict]:
    """
    Filter facts by axes and year, handling different fiscal year end dates.
    Matches facts where the period ends in the specified calendar year.
    """
    if not facts:
        return []

    # First filter by period - match any period ending in the specified year
    period_filtered = [
        fact for fact in facts 
        if fact.get('period', '') and str(year) in fact.get('period', '').split('/')[-1]
    ]
    
    logger.info(f"Filtering by period ending in year {year}:")
    logger.info(f"Before period filtering: {len(facts)} facts")
    logger.info(f"After period filtering: {len(period_filtered)} facts")
    
    # Then filter by axes if specified
    if not axes:
        logger.info(f"No axes specified, returning {len(period_filtered)} facts")
        return period_filtered
    
    # Convert single axis to list for consistent handling
    if isinstance(axes, str):
        axes = [axes]
    
    # Filter facts that have all required axes
    axis_filtered = []
    for fact in period_filtered:
        fact_axes = fact.get('axis', '').split('\n')  # Split multiple axes
        # Check if all required axes are present (case-insensitive)
        if all(
            any(req_axis.lower() in ax.lower() for ax in fact_axes)
            for req_axis in axes
        ):
            axis_filtered.append(fact)
    
    logger.info(f"Filtering by axes {axes}:")
    logger.info(f"Before axis filtering: {len(period_filtered)} facts")
    logger.info(f"After axis filtering: {len(axis_filtered)} facts")
    
    if len(axis_filtered) == 0:
        logger.warning("No facts found after filtering. Available axes:")
        all_axes = set()
        for fact in period_filtered:
            fact_axes = fact.get('axis', '').split('\n')
            all_axes.update(fact_axes)
        for ax in all_axes:
            if ax:  # Only print non-empty axes
                logger.info(f"- {ax}")
            
        logger.info("Available periods:")
        periods = set(fact.get('period') for fact in facts if fact.get('period'))
        for period in sorted(periods):  # Sort periods for clearer output
            logger.info(f"- {period}")
    
    return axis_filtered

def extract_segment_data(ticker: str, year: int, metric_config: Dict) -> List[Dict]:
    """Extract segment data for a specific metric."""
    tag = metric_config.get('tag')
    if not tag:
        return []
    
    logger.info(f"Processing tag: {tag}")
    filing_url = get_filing_url(ticker, year)
    if not filing_url:
        raise ValueError(f"Could not get filing URL for {ticker} {year}")
    
    logger.info(f"Filing URL: {filing_url}")
    facts = extract_inline_xbrl_data(filing_url, tag)
    
    logger.info(f"Extracted {len(facts)} facts for tag {tag}")
    if facts:
        logger.debug("Sample fact structure:")
        logger.debug(json.dumps(facts[0], indent=2))
    
    axes = metric_config.get('axes')  # Now getting 'axes' instead of 'axis'
    return filter_facts(facts, axes, year)

def deduplicate_metrics(facts):
    """
    Deduplicate facts based on fact value, member, and axis.
    Preserves original order of entries.
    """
    seen = set()
    deduplicated = []
    
    for fact in facts:
        # Create a unique key for comparison
        key = (str(fact.get('fact')), 
               str(fact.get('member')), 
               str(fact.get('axis')))
        
        if key not in seen:
            seen.add(key)
            deduplicated.append(fact)
    
    return deduplicated

def process_ticker(ticker: str, year: int) -> Dict:
    """Process a ticker and return the structured data."""
    config = load_config(ticker)
    result = {
        "ticker": ticker,
        "year": year,
        "metrics": {}
    }
    
    for metric_name, metric_config in config.items():
        if not metric_config.get('tag'):
            logger.warning(f"No tag configuration found for {metric_name}")
            continue
            
        try:
            facts = extract_segment_data(ticker, year, metric_config)
            if facts:
                # Deduplicate the facts for this metric
                deduped_facts = deduplicate_metrics(facts)
                result["metrics"][metric_name] = deduped_facts
                logger.info(f"Successfully extracted {len(facts)} facts for {metric_name}")
                logger.info(f"After deduplication: {len(deduped_facts)} unique facts")
            else:
                logger.warning(f"No data found for {metric_name}")
        except Exception as e:
            logger.error(f"Error processing {metric_name} for {ticker}: {e}")
            result["metrics"][metric_name] = []
    
    return result

def process_years(ticker: str, end_year: int) -> Dict:
    """
    Process a ticker for multiple years, going backwards from end_year until no data is found.
    Returns a dictionary containing data for all available years.
    """
    result = {
        "ticker": ticker,
        "end_year": end_year,
        "years": {}
    }
    
    current_year = end_year
    consecutive_failures = 0
    MAX_CONSECUTIVE_FAILURES = 1  # Stop after 1 year with no data
    
    while consecutive_failures < MAX_CONSECUTIVE_FAILURES:
        try:
            # Try to get filing URL for the current year
            filing_url = get_filing_url(ticker, current_year)
            if not filing_url:
                logger.warning(f"No filing found for {ticker} in {current_year}")
                consecutive_failures += 1
                current_year -= 1
                continue
                
            # Process the year's data
            year_data = process_ticker(ticker, current_year)
            
            # Check if we got any actual metrics data
            if any(year_data["metrics"].values()):
                result["years"][str(current_year)] = year_data["metrics"]
                consecutive_failures = 0  # Reset counter on success
                logger.info(f"Successfully processed {ticker} for {current_year}")
            else:
                consecutive_failures += 1
                logger.warning(f"No metrics data found for {ticker} in {current_year}")
            
        except Exception as e:
            logger.error(f"Error processing {ticker} for {current_year}: {e}")
            consecutive_failures += 1
            
        current_year -= 1
        
    if not result["years"]:
        logger.warning(f"No data found for {ticker} in any year")
    else:
        logger.info(f"Successfully processed {len(result['years'])} years for {ticker}")
        
    return result

def save_combined_results(result: Dict, ticker: str, end_year: int):
    """Save combined results to a JSON file."""
    filename = f"{ticker}_{end_year}_historical_breakdown.json"
    with open(filename, 'w') as f:
        json.dump(result, f, indent=2)
    logger.info(f"Combined results saved to {filename}")

def main():
    if len(sys.argv) != 3:
        logger.error("Usage: python sec_breakdown.py <TICKER> <END_YEAR>")
        sys.exit(1)
    
    # Load environment variables
    load_dotenv()
    if not FMP_API_KEY:
        logger.error("Error: FMP_API_KEY not found in .env file")
        sys.exit(1)
    
    ticker = sys.argv[1].upper()
    try:
        end_year = int(sys.argv[2])
    except ValueError:
        logger.error("Error: End year must be a valid integer")
        sys.exit(1)
    
    try:
        # Process multiple years
        result = process_years(ticker, end_year)
        save_combined_results(result, ticker, end_year)
    except Exception as e:
        logger.error(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()