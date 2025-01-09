import argparse
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
    level=logging.INFO,  # Changed from INFO to WARNING
    format='%(levelname)s - %(message)s',  # Simplified format
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
    
def process_fact_entry(fact: Dict, debug: bool = False) -> Dict:
    """
    Process a single fact entry, converting numeric values to integers
    and selecting relevant fields based on debug mode.
    """
    # Convert fact value to integer
    try:
        raw_value = fact['fact'].replace(',', '')
        numeric_value = int(float(raw_value))
    except (ValueError, TypeError):
        # If conversion fails, return None to filter out non-numeric entries
        logger.warning(f"Failed to convert fact value to integer: {fact.get('fact')}, setting to 0")
        numeric_value = 0

    # Basic fields to always include
    processed_fact = {
        'tag': fact.get('tag'),
        'fact': numeric_value,
        'axis': fact.get('axis'),
        #'member': fact.get('member'),
        'explicit_member': fact.get('explicit_member')
    }

    # If in debug mode, include all original fields
    if debug:
        processed_fact.update({
            k: v for k, v in fact.items()
            if k not in processed_fact  # Don't overwrite already processed fields
        })

    return processed_fact

def filter_facts(facts: List[Dict], axes: Optional[Union[List[str], str]], year: int, debug: bool = False) -> List[Dict]:
    """
    Filter facts by axes and year, handling different fiscal year end dates.
    Matches facts where the period ends in the specified calendar year.
    """
    if not facts:
        return []

    # First filter by period
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
        processed_facts = [process_fact_entry(fact, debug) for fact in period_filtered]
        return [f for f in processed_facts if f is not None]  # Filter out None values
    
    # Convert single axis to list
    if isinstance(axes, str):
        axes = [axes]
    
    # Filter by axes and process facts
    axis_filtered = []
    for fact in period_filtered:
        fact_axes = fact.get('axis', '').split('\n')
        if all(
            any(req_axis.lower() in ax.lower() for ax in fact_axes)
            for req_axis in axes
        ):
            processed_fact = process_fact_entry(fact, debug)
            if processed_fact:  # Only add if processing succeeded
                axis_filtered.append(processed_fact)
    
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
            if ax:
                logger.info(f"- {ax}")
            
        logger.info("Available periods:")
        periods = set(fact.get('period') for fact in facts if fact.get('period'))
        for period in sorted(periods):
            logger.info(f"- {period}")
    
    return axis_filtered

def extract_segment_data(ticker: str, year: int, metric_config: Dict, debug: bool = False) -> List[Dict]:
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
    if facts and debug:
        logger.debug("Sample fact structure:")
        logger.debug(json.dumps(facts[0], indent=2))
    
    axes = metric_config.get('axes')
    return filter_facts(facts, axes, year, debug)

def deduplicate_metrics(facts):
    """
    Deduplicate facts using a normalized comparison approach that properly handles
    multi-line fields and all relevant attributes.
    """
    def normalize_string(s):
        """Normalize string values by removing whitespace and sorting multi-line values."""
        if not s:
            return ''
        # Split by newlines, strip each part, sort, and filter empty strings
        parts = [p.strip() for p in str(s).split('\n')]
        return '\n'.join(sorted(filter(None, parts)))

    def create_comparison_key(fact):
        def strip_prefix(value: str) -> str:
            """
            If a field has 'something:' in front, remove it.
            E.g. 'fb:FamilyOfAppsMember' -> 'FamilyOfAppsMember'
            """
            # If multiline, handle each line separately
            lines = [x.strip() for x in value.split('\n')]
            stripped_lines = []
            for line in lines:
                # Remove the leading 'xxxx:' if present
                if ':' in line:
                    line = line.split(':', 1)[1]
                stripped_lines.append(line)
            # Sort lines so that order differences (Axes reversed, etc.) don’t cause duplicates
            stripped_lines = sorted(stripped_lines)
            return '\n'.join(stripped_lines).strip()
        
        # Normalize the actual “fact” field
        fact_val = normalize_string(str(fact.get('fact', '')))
        
        # For axis, member, explicit_member, strip and sort lines
        axis_val = strip_prefix(fact.get('axis', ''))
        member_val = strip_prefix(fact.get('member', ''))
        explicit_val = strip_prefix(fact.get('explicit_member', ''))
        
        # Similarly for period if you only care about the year, unify that as well:
        # e.g. "12 months ending 12/31/2023" -> "2023"
        # (If you do want month precision, ignore this step.)
        period_text = fact.get('period', '')
        
        key_parts = [
            fact_val,
            axis_val,
            member_val,
            explicit_val,
            normalize_string(fact.get('tag', '')),
            # Maybe you only want to unify unit and measure if they're different, up to you
            normalize_string(fact.get('unit_ref', '')),
            normalize_string(fact.get('measure', '')),
            # Potentially skip format or unify blank with 'num-dot-decimal'
            # skip scale/decimals entirely
            normalize_string(fact.get('sign', '')),  
            normalize_string(fact.get('type', '')),  
            period_text,
        ]
        
        return '|||'.join(key_parts)


    seen_keys = {}
    deduplicated = []
    
    for fact in facts:
        key = create_comparison_key(fact)
        if key not in seen_keys:
            seen_keys[key] = fact
            deduplicated.append(fact)
    
    return deduplicated

def process_ticker(ticker: str, year: int, config: Dict, debug: bool = False) -> Dict:
    """Process a ticker and return the structured data."""
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
            facts = extract_segment_data(ticker, year, metric_config, debug)
            if facts:
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

def process_years(ticker: str, end_year: int, debug: bool = False) -> Dict:
    """
    Process a ticker for multiple years. When we can't find matching data in a year's filing,
    go back to the most recent successful filing to search for that and subsequent years' data.
    """
    config = load_config(ticker)

    result = {
        "ticker": ticker,
        "end_year": end_year,
        "years": {}
    }
    
    current_year = end_year
    last_successful_data_year = None
    last_successful_filing_url = None
    
    while True:
        try:
            # Get filing URL for the current year
            filing_url = get_filing_url(ticker, current_year)
            if not filing_url:
                logger.warning(f"No filing found for {ticker} in {current_year}")
                break
                
            # Try to get data from this year's filing
            logger.info(f"Processing {current_year} filing...")
            year_data = process_ticker(ticker, current_year, config, debug)
            
            if any(year_data["metrics"].values()):
                # Found matching data in this year's filing
                result["years"][str(current_year)] = year_data["metrics"]
                last_successful_data_year = current_year
                last_successful_filing_url = filing_url
                logger.info(f"Found matching data in {current_year} filing")
            else:
                # No matching data in this year's filing - try last successful filing
                logger.warning(f"No matching data found in {current_year} filing")
                
                if last_successful_filing_url and last_successful_data_year:
                    logger.info(f"Searching {last_successful_data_year} filing for {current_year} and earlier years...")
                    
                    # Use the config for the ticker to get metric tags
                    
                    search_year = current_year
                    
                    while True:
                        has_data = False
                        year_metrics = {}
                        
                        # For each metric in the config, try to find data for the search year
                        for metric_name, metric_config in config.items():
                            tag = metric_config.get('tag')
                            if not tag:
                                continue
                                
                            # Get all facts for this metric from the last successful filing
                            facts = extract_inline_xbrl_data(last_successful_filing_url, tag)
                            
                            # Filter facts for the search year
                            filtered_facts = filter_facts(facts, metric_config.get('axes'), search_year, debug)
                            
                            if filtered_facts:
                                # Deduplicate the filtered facts
                                deduped_facts = deduplicate_metrics(filtered_facts)
                                if deduped_facts:
                                    has_data = True
                                    year_metrics[metric_name] = deduped_facts
                        
                        if has_data:
                            result["years"][str(search_year)] = year_metrics
                            logger.info(f"Found data for {search_year} in {last_successful_data_year} filing")
                            search_year -= 1
                        else:
                            logger.warning(f"No data found for {search_year} in {last_successful_data_year} filing")
                            break
                
                break  # Exit main loop after checking historical data
                
        except Exception as e:
            logger.error(f"Error processing {ticker} for {current_year}: {e}")
            break
            
        current_year -= 1
        
    if not result["years"]:
        logger.warning(f"No matching data found for {ticker} in any year")
    else:
        logger.info(f"Successfully processed {len(result['years'])} years for {ticker}")
        years_found = sorted(map(int, result["years"].keys()))
        logger.info(f"Data found for years: {years_found}")

    return result  

def save_combined_results(result: Dict, ticker: str, end_year: int):
    """Save combined results to a JSON file."""
    filename = f"{ticker}_{end_year}_historical_breakdown.json"
    with open(filename, 'w') as f:
        json.dump(result, f, indent=2)
    logger.info(f"Combined results saved to {filename}")

def main():
    parser = argparse.ArgumentParser(description='Process SEC filings for segment breakdown analysis')
    parser.add_argument('ticker', type=str, help='Stock ticker symbol')
    parser.add_argument('end_year', type=int, help='End year for analysis')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode to retain all fields')
    args = parser.parse_args()
    
    # Load environment variables
    load_dotenv()
    if not FMP_API_KEY:
        logger.error("Error: FMP_API_KEY not found in .env file")
        sys.exit(1)
    
    try:
        result = process_years(args.ticker.upper(), args.end_year, args.debug)
        save_combined_results(result, args.ticker.upper(), args.end_year)
    except Exception as e:
        logger.error(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()