# unified_segmentation.py

import os
import re
import json
import requests
import logging
from bs4 import BeautifulSoup
from datetime import datetime
import math
import argparse
import sys
from typing import Dict, List, Any, Optional, Union
from dotenv import load_dotenv

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()
FMP_API_KEY = os.getenv('FMP_API_KEY')

HEADERS = {
    "User-Agent": "Custom Research Agent - Contact: rgranowski@gmail.com"
}

def load_config(ticker: str) -> Dict:
    """Load the configuration for a specific ticker."""
    try:
        with open('unified_segmentation_config.json', 'r') as f:
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

def get_financial_statement(symbol: str, year: int) -> dict:
    """Fetch the balance sheet statement for the given symbol and year."""
    url = f"https://financialmodelingprep.com/api/v3/balance-sheet-statement/{symbol}"
    params = {
        "period": "annual",
        "apikey": FMP_API_KEY
    }
    
    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        statements = response.json()
        
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
    """Get the iXBRL filing URL for the given symbol and year."""
    statement = get_financial_statement(symbol, year)
    if not statement:
        return None
        
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
    if ':' in axis_name:
        axis_name = axis_name.split(':')[-1]
    
    if axis_name == "StatementBusinessSegmentsAxis":
        return "US-GAAP Statement Business Segments Axis"
    
    formatted = re.sub(r'(?<!^)(?=[A-Z])', ' ', axis_name)
    if "Axis" not in formatted:
        formatted += " Axis"
    return formatted

def format_member_name(member_value, company_prefix=None):
    """Format member name to standard format."""
    if ':' not in member_value:
        return member_value
    
    prefix, member = member_value.split(':')
    formatted = re.sub(r'(?<!^)(?=[A-Z])', ' ', member)
    
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
                months = math.ceil(days / 30.44)
                end_formatted = end.strftime('%m/%d/%Y')
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
        raw_value = elem.get_text(strip=True)
        decimals = elem.get("decimals")
        scale = elem.get("scale")
        format_attr = elem.get("format", "")
        sign_attr = elem.get("sign")
        
        if format_attr.startswith("ixt:"):
            format_attr = format_attr[4:]
        
        try:
            numeric_value = float(raw_value.replace(",", ""))
            
            if sign_attr == "-" or raw_value.startswith("-"):
                numeric_value = -abs(numeric_value)
            
            scale_factor = 1000000 if get_scale_factor(decimals, scale) == "Millions" else \
                          1000 if get_scale_factor(decimals, scale) == "Thousands" else 1
            scaled_value = numeric_value * scale_factor
            
            formatted_value = f"{scaled_value:,.0f}"
            
        except (ValueError, TypeError):
            formatted_value = raw_value
        
        fact = {
            "tag": target_tag,
            "fact": formatted_value,
            "scale": get_scale_factor(decimals, scale),
            "decimals": get_scale_factor(decimals, scale),
            "format": format_attr,
            "sign": "Negative" if (sign_attr == "-" or raw_value.startswith("-")) else "Positive",
            "type": "Monetary Item Type" if elem.get("unitref") or elem.get("unitRef") else "String Type"
        }
        
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

def filter_facts(facts: List[Dict], axes: Optional[Union[List[str], str]], year: int, debug: bool = False) -> List[Dict]:
    """Filter facts by axes and year."""
    if not facts:
        return []

    period_filtered = [
        fact for fact in facts 
        if fact.get('period', '') and str(year) in fact.get('period', '').split('/')[-1]
    ]
    
    logger.info(f"Filtering by period ending in year {year}:")
    logger.info(f"Before period filtering: {len(facts)} facts")
    logger.info(f"After period filtering: {len(period_filtered)} facts")
    
    if not axes:
        logger.info(f"No axes specified, returning {len(period_filtered)} facts")
        return period_filtered
    
    if isinstance(axes, str):
        axes = [axes]
    
    axis_filtered = []
    for fact in period_filtered:
        fact_axes = fact.get('axis', '').split('\n')
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
            if ax:
                logger.info(f"- {ax}")
            
        logger.info("Available periods:")
        periods = set(fact.get('period') for fact in facts if fact.get('period'))
        for period in sorted(periods):
            logger.info(f"- {period}")
    
    return axis_filtered

def deduplicate_metrics(facts):
    """Deduplicate facts using a normalized comparison approach."""
    def normalize_string(s):
        if not s:
            return ''
        parts = [p.strip() for p in str(s).split('\n')]
        return '\n'.join(sorted(filter(None, parts)))

    def strip_prefix(value: str) -> str:
        lines = [x.strip() for x in value.split('\n')]
        stripped_lines = []
        for line in lines:
            if ':' in line:
                line = line.split(':', 1)[1]
            stripped_lines.append(line)
        stripped_lines = sorted(stripped_lines)
        return '\n'.join(stripped_lines).strip()

    def create_comparison_key(fact):
        fact_val = normalize_string(str(fact.get('fact', '')))
        axis_val = strip_prefix(fact.get('axis', ''))
        member_val = strip_prefix(fact.get('member', ''))
        explicit_val = strip_prefix(fact.get('explicit_member', ''))
        period_text = fact.get('period', '')
        
        key_parts = [
            fact_val,
            axis_val,
            member_val,
            explicit_val,
            normalize_string(fact.get('tag', '')),
            normalize_string(fact.get('unit_ref', '')),
            normalize_string(fact.get('measure', '')),
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

def process_raw_segmentation(ticker: str, year: int, metric_config: Dict, debug: bool = False) -> List[Dict]:
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

def transform_facts(raw_facts: List[Dict], name_mapping: Dict[str, str]) -> Dict[str, int]:
    """Transform raw facts into structured data using name mappings."""
    transformed = {}
    
    for fact in raw_facts:
        entry_member_parts = set(fact["explicit_member"].split("\n"))
        for config_member, mapped_name in name_mapping.items():
            config_member_parts = set(config_member.split("\n"))
            if entry_member_parts == config_member_parts:
                # Convert fact value to float, removing commas
                value = int(fact["fact"].replace(",", ""))
                transformed[mapped_name] = value
                break
    
    return transformed

def process_years(ticker: str, end_year: int, raw_output: bool = False, debug: bool = False) -> Dict:
    """Process multiple years of data with support for raw and transformed output."""
    config = load_config(ticker)
    
    result = {}
    
    current_year = end_year
    last_successful_data_year = None
    last_successful_filing_url = None
    
    while True:
        try:
            filing_url = get_filing_url(ticker, current_year)
            if not filing_url:
                break
                
            year_metrics = {}
            metrics_found = False
            
            for metric_name, metric_config in config.items():
                raw_facts = process_raw_segmentation(ticker, current_year, metric_config, debug)
                
                if raw_facts:
                    metrics_found = True
                    if raw_output:
                        year_metrics[metric_name] = raw_facts
                    else:
                        transformed_data = transform_facts(raw_facts, metric_config.get('name_mapping', {}))
                        year_metrics[metric_name] = transformed_data
            
            if metrics_found:
                # Copy revenue data to segmentation if segmentation is empty
                if 'segmentation' in year_metrics and not year_metrics['segmentation'] and 'revenue' in year_metrics:
                    year_metrics['segmentation'] = year_metrics['revenue'].copy()
                
                # Add "Total Services" to segmentation if we have Google Services in revenue
                if 'segmentation' in year_metrics and 'revenue' in year_metrics:
                    if 'Google Services' in year_metrics['revenue']:
                        year_metrics['segmentation']['Total Services'] = year_metrics['revenue']['Google Services']
                    if 'Google Cloud' in year_metrics['revenue']:
                        year_metrics['segmentation']['Google Cloud'] = year_metrics['revenue']['Google Cloud']
                    if 'Other Segments' in year_metrics['revenue']:
                        year_metrics['segmentation']['Other Segments'] = year_metrics['revenue']['Other Segments']
                
                result[str(current_year)] = year_metrics
                last_successful_data_year = current_year
                last_successful_filing_url = filing_url
            else:
                if last_successful_filing_url and last_successful_data_year:
                    logger.info(f"Searching {last_successful_data_year} filing for {current_year} data...")
                    
                    search_year = current_year
                    while True:
                        year_metrics = {}
                        metrics_found = False
                        
                        for metric_name, metric_config in config.items():
                            facts = extract_inline_xbrl_data(last_successful_filing_url, metric_config['tag'])
                            filtered_facts = filter_facts(facts, metric_config.get('axes'), search_year, debug)
                            
                            if filtered_facts:
                                metrics_found = True
                                if raw_output:
                                    year_metrics[metric_name] = filtered_facts
                                else:
                                    transformed_data = transform_facts(filtered_facts, metric_config.get('name_mapping', {}))
                                    year_metrics[metric_name] = transformed_data
                        
                        if metrics_found:
                            # Copy revenue data to segmentation if segmentation is empty
                            if 'segmentation' in year_metrics and not year_metrics['segmentation'] and 'revenue' in year_metrics:
                                year_metrics['segmentation'] = year_metrics['revenue'].copy()
                            
                            # Add "Total Services" to segmentation if we have Google Services in revenue
                            if 'segmentation' in year_metrics and 'revenue' in year_metrics:
                                if 'Google Services' in year_metrics['revenue']:
                                    year_metrics['segmentation']['Total Services'] = year_metrics['revenue']['Google Services']
                                if 'Google Cloud' in year_metrics['revenue']:
                                    year_metrics['segmentation']['Google Cloud'] = year_metrics['revenue']['Google Cloud']
                                if 'Other Segments' in year_metrics['revenue']:
                                    year_metrics['segmentation']['Other Segments'] = year_metrics['revenue']['Other Segments']
                            
                            result[str(search_year)] = year_metrics
                            search_year -= 1
                        else:
                            break
                
                break
                
        except Exception as e:
            logger.error(f"Error processing {ticker} for {current_year}: {e}")
            break
            
        current_year -= 1
    
    return result

def save_results(result: Dict, ticker: str, end_year: int, raw_output: bool = False):
    """Save results to a JSON file."""
    suffix = "raw_segmentation" if raw_output else "segmentation"
    filename = f"{ticker}_{end_year}_{suffix}.json"
    
    # Sort years in descending order
    ordered_result = {}
    for year in sorted(result.keys(), reverse=True):
        ordered_result[year] = result[year]
    with open(filename, 'w') as f:
        json.dump(ordered_result, f, indent=2)
    logger.info(f"Results saved to {filename}")

def main():
    parser = argparse.ArgumentParser(description='Process SEC filings for segment breakdown analysis')
    parser.add_argument('ticker', type=str, help='Stock ticker symbol')
    parser.add_argument('end_year', type=int, help='End year for analysis')
    parser.add_argument('--raw_segmentation', action='store_true', help='Output raw segmentation data')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode')
    args = parser.parse_args()
    
    if not FMP_API_KEY:
        logger.error("Error: FMP_API_KEY not found in .env file")
        sys.exit(1)
    
    try:
        result = process_years(args.ticker.upper(), args.end_year, args.raw_segmentation, args.debug)
        save_results(result, args.ticker.upper(), args.end_year, args.raw_segmentation)
    except Exception as e:
        logger.error(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()