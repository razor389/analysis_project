#!/usr/bin/env python3
import os
import json
import logging
import re
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import time
from typing import Dict, List
import argparse
from dotenv import load_dotenv
import xml.etree.ElementTree as ET
from urllib.parse import urljoin
from urllib3.util.retry import Retry

from unified_segmentation import get_filing_contents, get_scale_factor

# Load environment variables
load_dotenv()
FMP_API_KEY = os.getenv('FMP_API_KEY')
if not FMP_API_KEY:
    raise ValueError("FMP_API_KEY not found in environment variables")

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

class EDGARExhibit13Finder:
    BASE_URL = "https://www.sec.gov"

    def __init__(self, user_agent: str):
        self.headers = {
            "User-Agent": user_agent,
            "Accept-Encoding": "gzip, deflate"
        }
        self.session = requests.Session()
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["HEAD", "GET", "OPTIONS"]
        )
        adapter = requests.adapters.HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("https://", adapter)

    def get_cik_from_ticker(self, ticker: str) -> str:
        """
        Get CIK from ticker using FMP API.
        Uses the session (with retries) and proper timeout.
        """
        url = f"https://financialmodelingprep.com/api/v3/balance-sheet-statement/{ticker}"
        params = {
            "period": "annual",
            "apikey": FMP_API_KEY,
            "limit": 1
        }

        try:
            response = self.session.get(url, params=params, timeout=(10, 30))
            response.raise_for_status()
            data = response.json()

            if not data:
                raise ValueError(f"No data found for ticker {ticker}")

            # Some endpoints return the CIK with leading zeros.
            cik = data[0].get("cik", "").lstrip("0")
            if not cik:
                raise ValueError(f"No CIK found for ticker {ticker}")

            return cik

        except requests.exceptions.Timeout:
            logger.error(f"Timeout while fetching CIK for ticker {ticker}")
            raise
        except requests.exceptions.RequestException as e:
            logger.error(f"Request error getting CIK for ticker {ticker}: {e}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error getting CIK for ticker {ticker}: {e}")
            raise

    def get_company_filings(self, cik: str) -> Dict:
        """
        Get recent 10-K filings using the EDGAR browse API.
        Returns a dictionary with a list of filings.
        """
        url = urljoin(
            self.BASE_URL,
            f"/cgi-bin/browse-edgar?action=getcompany&CIK={cik}&type=10-K&dateb=&owner=exclude&start=0&count=40&output=atom"
        )

        response = self.session.get(url, headers=self.headers, timeout=(10, 30))
        response.raise_for_status()

        try:
            root = ET.fromstring(response.content)
        except ET.ParseError as e:
            logger.error("Error parsing XML from SEC response.")
            raise e

        entries = []
        ns = {'atom': 'http://www.w3.org/2005/Atom'}

        for entry in root.findall('atom:entry', ns):
            accession_number = ''
            filing_href = ''
            filing_date = ''

            id_elem = entry.find('atom:id', ns)
            if id_elem is not None and id_elem.text:
                accession_match = re.search(r'accession-number=(\d{10}-\d{2}-\d{6})', id_elem.text)
                if accession_match:
                    accession_number = accession_match.group(1)

            date_elem = entry.find('atom:updated', ns)
            if date_elem is not None and date_elem.text:
                # Filing date in YYYY-MM-DD format
                filing_date = date_elem.text.split('T')[0]

            link_elem = entry.find('atom:link', ns)
            if link_elem is not None:
                href = link_elem.get('href')
                if href:
                    filing_href = urljoin(self.BASE_URL, href)

            if accession_number and filing_href:
                entries.append({
                    'accession_number': accession_number,
                    'filing_href': filing_href,
                    'filing_date': filing_date
                })

        return {'filings': entries}

    def get_filing_detail(self, filing_url: str) -> Dict:
        """
        Get the filing detail page from the SEC website and locate the Exhibit 13 document.
        """
        response = self.session.get(filing_url, headers=self.headers, timeout=(10, 30))
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')
        documents = {'exhibit13': None}

        # Try finding the table by its summary attribute
        table = soup.find('table', summary=lambda value: value and 'Document Format Files' in value)
        if not table:
            table = soup.find('table', class_='tableFile')

        if table:
            # Skip the header row
            for row in table.find_all('tr')[1:]:
                cells = row.find_all('td')
                if len(cells) >= 4:
                    description = cells[1].get_text(strip=True).lower()
                    document_link = cells[2].find('a')
                    if document_link and ('ex-13' in description or 'exhibit 13' in description):
                        documents['exhibit13'] = urljoin(self.BASE_URL, document_link.get('href'))
                        break  # Stop after finding the first matching Exhibit 13

        if not documents['exhibit13']:
            logger.warning(f"No Exhibit 13 found in filing page: {filing_url}")

        return documents


class InsuranceMetricsExtractor:
    def __init__(self, user_agent: str):
        self.headers = {
            "User-Agent": user_agent,
            "Accept-Encoding": "gzip, deflate"
        }
        self.session = requests.Session()
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["HEAD", "GET", "OPTIONS"]
        )
        adapter = requests.adapters.HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("https://", adapter)
        # Mapping of our target metrics to the corresponding inline XBRL tags.
        self.metrics = {
            "gross_revenues": "us-gaap:PremiumsEarnedNetPropertyAndCasualty",
            "investment_income": "us-gaap:InterestAndDividendIncomeOperating",
            "losses_and_expenses": "us-gaap:IncurredClaimsPropertyCasualtyAndLiability",
            "acquisition_costs": "us-gaap:DeferredPolicyAcquisitionCostAmortizationExpense",
            "underwriting_expenses": "us-gaap:OtherUnderwritingExpense",
            "service_expenses": "pgr:NonInsuranceServiceExpenses"
        }

    def parse_context(self, soup, context_ref):
        """
        Enhanced context parser that also extracts dates from context reference IDs.
        """
        logger.debug(f"\n{'='*50}\nParsing context: {context_ref}\n{'='*50}")
        
        # First try to extract dates directly from the context_ref ID
        # Look for patterns like D20210101-20211231 in the ID
        date_pattern = r'D(\d{8})-(\d{8})'
        id_match = re.search(date_pattern, context_ref)
        if id_match:
            start_date = id_match.group(1)
            end_date = id_match.group(2)
            # Format dates nicely
            start_formatted = f"{start_date[:4]}-{start_date[4:6]}-{start_date[6:]}"
            end_formatted = f"{end_date[:4]}-{end_date[4:6]}-{end_date[6:]}"
            return {'period': f"{start_formatted} to {end_formatted}"}
        
        # If no date in ID, try finding the context element
        context = soup.find(attrs={"id": context_ref})
        if not context:
            # Try different case variations
            context = (soup.find(attrs={"id": context_ref.lower()}) or 
                    soup.find(attrs={"id": context_ref.upper()}) or
                    soup.find(attrs={"ID": context_ref}))
        
        if context:
            logger.debug(f"Found context element: {str(context)[:200]}")
            
            # Try to find period information in the element
            period = context.find(['xbrli:period', 'period'])
            if period:
                start = period.find(['xbrli:startdate', 'startdate'])
                end = period.find(['xbrli:enddate', 'enddate'])
                instant = period.find(['xbrli:instant', 'instant'])
                
                if start and end:
                    return {'period': f"{start.text.strip()} to {end.text.strip()}"}
                elif instant:
                    return {'period': f"As of {instant.text.strip()}"}
        
        # If we haven't found a date yet, try looking for date patterns in the context_ref
        year_pattern = r'(\d{4})'
        year_match = re.search(year_pattern, context_ref)
        if year_match:
            year = year_match.group(1)
            return {'period': f"{year}-01-01 to {year}-12-31"}
        
        logger.debug("No period information found")
        return {}

    def extract_metrics(self, filing_url: str) -> Dict:
        """
        Enhanced metric extractor with better error handling and debugging.
        """
        content, meta_info = get_filing_contents(filing_url)
        if not content:
            logger.error("No filing content retrieved.")
            return {}

        logger.info("Parsing document...")
        soup = BeautifulSoup(content, 'html.parser')
        results = {}

        # Log some document structure information
        logger.debug(f"Document contains {len(soup.find_all())} total elements")
        logger.debug(f"Found {len(soup.find_all('ix:nonfraction'))} ix:nonfraction elements")
        
        for metric_name, tag in self.metrics.items():
            logger.info(f"\nProcessing metric: {metric_name} (tag: {tag})")
            
            # Find elements using multiple approaches
            xbrl_elements = []
            xbrl_elements.extend(soup.find_all(attrs={"name": tag}))
            xbrl_elements.extend(soup.find_all(attrs={"data-xlinkLabel": tag}))
            xbrl_elements.extend(soup.find_all("ix:nonfraction", attrs={"name": tag}))
            
            # Remove duplicates while preserving order
            xbrl_elements = list(dict.fromkeys(xbrl_elements))
            
            logger.info(f"Found {len(xbrl_elements)} elements for metric '{metric_name}'")
            
            for idx, elem in enumerate(xbrl_elements, 1):
                logger.debug(f"\nProcessing element {idx}/{len(xbrl_elements)} for {metric_name}:")
                logger.debug(f"Element: {str(elem)[:200]}")
                
                # Extract and clean the value
                raw_value = elem.get_text(strip=True)
                logger.debug(f"Raw value: '{raw_value}'")
                
                try:
                    # Handle parentheses, commas, and negative values
                    cleaned_value = raw_value.replace(",", "")
                    if cleaned_value.startswith("(") and cleaned_value.endswith(")"):
                        cleaned_value = "-" + cleaned_value[1:-1]
                    numeric_value = float(cleaned_value)
                    
                    # Handle scaling
                    scale = elem.get("scale", "0")
                    decimals = elem.get("decimals", "0")
                    logger.debug(f"Scale: {scale}, Decimals: {decimals}")
                    
                    if scale and scale != "0":
                        numeric_value *= 10 ** int(scale)
                    
                    logger.debug(f"Final numeric value: {numeric_value}")
                    
                    # Get context information
                    context_ref = elem.get("contextref") or elem.get("contextRef")
                    if not context_ref:
                        logger.debug("No context reference found")
                        continue
                    
                    context_data = self.parse_context(soup, context_ref)
                    period_text = context_data.get("period", "")
                    
                    if not period_text:
                        logger.debug("No period text found")
                        continue
                    
                    # Extract year
                    year_match = re.search(r'(\d{4})', period_text)
                    if not year_match:
                        logger.debug(f"No year found in period text: {period_text}")
                        continue
                    
                    year = year_match.group(1)
                    
                    # Store the result
                    if year not in results:
                        results[year] = {}
                    results[year][metric_name] = numeric_value
                    logger.debug(f"Stored {metric_name} = {numeric_value} for year {year}")
                    
                except (ValueError, TypeError) as e:
                    logger.error(f"Error processing value for {metric_name}: {e}")
                    continue
        
        logger.info(f"Final extracted results: {results}")
        return results

def main():
    parser = argparse.ArgumentParser(description='Extract insurance metrics from SEC filings')
    parser.add_argument('ticker', type=str, help='Company ticker symbol')
    parser.add_argument('start_year', type=str, help='Start year (YYYY)')
    parser.add_argument('--email', type=str, required=True, help='Your email for SEC user agent')
    parser.add_argument('--output', type=str, help='Output JSON file')
    args = parser.parse_args()

    # Initialize finder and extractor using the supplied email in the user-agent string.
    finder = EDGARExhibit13Finder(f"Insurance Research - Contact: {args.email}")
    extractor = InsuranceMetricsExtractor(f"Insurance Research - Contact: {args.email}")

    try:
        # Get the company's CIK using the ticker
        cik = finder.get_cik_from_ticker(args.ticker.upper())
        logger.info(f"Retrieved CIK for {args.ticker.upper()}: {cik}")

        # Get recent 10-K filings for the company
        filings_data = finder.get_company_filings(cik)
        if not filings_data.get('filings'):
            logger.error("No filings found.")
            return

        all_results = {}
        for filing in filings_data['filings']:
            # Only process filings on or after the start year
            if filing['filing_date'] < f"{args.start_year}-01-01":
                continue

            filing_docs = finder.get_filing_detail(filing['filing_href'])
            exhibit_url = filing_docs.get('exhibit13')
            if exhibit_url:
                logger.info(f"Processing Exhibit 13 from filing dated {filing['filing_date']}")
                year_results = extractor.extract_metrics(exhibit_url)
                for year, metrics in year_results.items():
                    if year not in all_results:
                        all_results[year] = {}
                    all_results[year].update(metrics)
            else:
                logger.warning(f"Exhibit 13 not found for filing dated {filing['filing_date']}")

            # Respect SEC rate limits
            time.sleep(0.1)

        # Order the results with years in descending order
        ordered_results = dict(sorted(all_results.items(), reverse=True))

        # Save the results to a JSON file
        output_file = args.output or f"{args.ticker.lower()}_insurance_metrics.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(ordered_results, f, indent=2)
        logger.info(f"Results saved to {output_file}")

    except Exception as e:
        logger.error(f"Error: {e}")
        raise


if __name__ == "__main__":
    main()
