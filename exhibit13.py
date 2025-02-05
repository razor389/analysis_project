import requests
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional
import time
from datetime import datetime
import logging
import re
from urllib.parse import urljoin
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
FMP_API_KEY = os.getenv('FMP_API_KEY')

# Configure logging
logging.basicConfig(
    level=logging.INFO,  # Change to DEBUG to see more details
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class EDGARExhibit13Finder:
    BASE_URL = "https://www.sec.gov"
    
    def __init__(self, user_agent: str):
        """
        Initialize the finder with your contact information (required by SEC)
        Args:
            user_agent: String with your name and email, e.g. "Name (email@domain.com)"
        """
        self.headers = {
            "User-Agent": user_agent,
            "Accept-Encoding": "gzip, deflate"
        }
        self.session = requests.Session()
        
        # Configure session with timeouts and retries using urllib3 Retry
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
        Get CIK from ticker using FMP API
        Args:
            ticker: Company ticker symbol
        Returns:
            CIK number as string with leading zeros removed
        """
        url = f"https://financialmodelingprep.com/api/v3/balance-sheet-statement/{ticker}"
        params = {
            "period": "annual",
            "apikey": FMP_API_KEY,
            "limit": 1
        }
        
        try:
            response = requests.get(url, params=params)
            response.raise_for_status()
            data = response.json()
            
            if not data:
                raise ValueError(f"No data found for ticker {ticker}")
                
            # Extract CIK and remove leading zeros
            cik = data[0].get("cik", "").lstrip("0")
            if not cik:
                raise ValueError(f"No CIK found for ticker {ticker}")
                
            return cik
        
        except Exception as e:
            logger.error(f"Error getting CIK for ticker {ticker}: {e}")
            raise

    def get_company_filings(self, cik: str) -> Dict:
        """
        Get company filings using EDGAR browse API
        Args:
            cik: Company CIK number
        """
        url = urljoin(
            self.BASE_URL,
            f"/cgi-bin/browse-edgar?action=getcompany&CIK={cik}&type=10-K&dateb=&owner=exclude&start=0&count=40&output=atom"
        )
        logger.info(f"Fetching filings from URL: {url}")
        try:
            response = self.session.get(
                url, 
                headers=self.headers,
                timeout=(10, 30)
            )
            response.raise_for_status()
            
            # Debug: Uncomment to log the raw XML
            # logger.debug(response.text)
            
            # Parse XML response
            root = ET.fromstring(response.content)
            
            entries = []
            # Namespace for Atom XML
            ns = {'atom': 'http://www.w3.org/2005/Atom'}
            for entry in root.findall('atom:entry', ns):
                accession_number = ''
                filing_href = ''
                filing_date = ''
                
                # Get accession number from id
                id_elem = entry.find('atom:id', ns)
                if id_elem is not None and id_elem.text:
                    accession_match = re.search(r'accession-number=(\d{10}-\d{2}-\d{6})', id_elem.text)
                    if accession_match:
                        accession_number = accession_match.group(1)
                    else:
                        logger.debug(f"No accession number match found in id: {id_elem.text}")
                
                # Get filing date
                date_elem = entry.find('atom:updated', ns)
                if date_elem is not None and date_elem.text:
                    # The date is in ISO format; we extract the date portion.
                    filing_date = date_elem.text.split('T')[0]
                
                # Get filing detail URL
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
            
            logger.info(f"Found {len(entries)} filings for CIK {cik}.")
            return {'filings': entries}
            
        except requests.exceptions.Timeout:
            logger.error("Timeout while fetching company filings")
            raise
        except requests.exceptions.RequestException as e:
            logger.error(f"Error fetching company filings: {e}")
            raise
        except ET.ParseError as e:
            logger.error(f"Error parsing company filings XML: {e}")
            raise

    def get_filing_detail(self, filing_url: str) -> Dict:
        """
        Get the filing detail page to find links to documents.
        This version uses BeautifulSoup to parse the HTML and then look for rows
        corresponding to Exhibit 13 (based on the text in the Description or Type cells).
        
        Args:
            filing_url: URL of the filing detail page
        Returns:
            A dictionary with keys '10k' and 'exhibit13' (if found)
        """
        logger.info(f"Fetching filing detail page: {filing_url}")
        try:
            response = self.session.get(
                filing_url,
                headers=self.headers,
                timeout=(10, 30)
            )
            response.raise_for_status()
            html = response.text

            soup = BeautifulSoup(html, 'html.parser')
            
            # Initialize our dictionary for documents
            documents = {
                '10k': None,
                'exhibit13': None
            }
            
            # Try to locate the table that lists the Document Format Files.
            # Many filing pages use a table with summary="Document Format Files" or a class like "tableFile"
            table = soup.find('table', summary=lambda value: value and 'Document Format Files' in value)
            if not table:
                table = soup.find('table', class_='tableFile')
            
            if table:
                # Iterate over all table rows (skip the header row)
                rows = table.find_all('tr')
                for row in rows:
                    cells = row.find_all('td')
                    # Expecting at least 4 columns: Seq, Description, Document, Type, (and Size)
                    if len(cells) >= 4:
                        # Extract text and normalize to lowercase for matching.
                        description_text = cells[1].get_text(strip=True).lower()
                        type_text = cells[3].get_text(strip=True).lower()
                        
                        # Extract the link from the Document cell (usually third column)
                        link_tag = cells[2].find('a')
                        if link_tag and link_tag.has_attr('href'):
                            full_url = urljoin(self.BASE_URL, link_tag['href'])
                        else:
                            full_url = None

                        # Identify the 10-K by looking for "10-k" in description or type
                        if (('10-k' in description_text or '10-k' in type_text) and 
                            full_url and '/ix?doc=' in full_url):
                            documents['10k'] = full_url
                            
                        # Identify Exhibit 13 by checking for "ex-13" or "exhibit 13"
                        # in either the description or type cell.
                        elif (('ex-13' in description_text or 'exhibit 13' in description_text or
                            'ex-13' in type_text or 'exhibit 13' in type_text or
                            'exhibit13' in description_text or 'exhibit13' in type_text) and
                            full_url):
                            documents['exhibit13'] = full_url
                            logger.debug(f"Found Exhibit 13 link: {full_url}")
            else:
                logger.warning("Could not locate the Document Format Files table using BeautifulSoup.")

            return documents

        except requests.exceptions.RequestException as e:
            logger.error(f"Error getting filing detail: {e}")
            return {}

    def find_exhibit13_ixbrl(self, ticker: str, start_date: Optional[str] = None) -> List[Dict]:
        """
        Find Exhibit 13 iXBRL documents for a company.
        Args:
            cik: Company CIK number.
            start_date: Optional start date in YYYY-MM-DD format. Filings before this date are skipped.
        Returns:
            List of dictionaries containing filing info and URLs.
        """
        results = []
        
        try:
            cik = self.get_cik_from_ticker(ticker)
            logger.info(f"Found CIK {cik} for ticker {ticker}")
            filings_data = self.get_company_filings(cik)
        except Exception as e:
            logger.error(f"Could not get company filings: {e}")
            return results

        for filing in filings_data['filings']:
            try:
                # Skip filings older than start_date if specified
                if start_date and filing['filing_date'] < start_date:
                    logger.debug(f"Skipping filing {filing['accession_number']} due to date {filing['filing_date']}")
                    continue
                    
                # Get the filing detail page
                filing_docs = self.get_filing_detail(filing['filing_href'])
                
                # Proceed if both a 10-K and an Exhibit 13 document were found.
                if filing_docs.get('10k') and filing_docs.get('exhibit13'):
                    # Check if exhibit13 is in iXBRL format
                    is_ixbrl = self._check_if_ixbrl(filing_docs['exhibit13'])
                    
                    results.append({
                        'filing_date': filing['filing_date'],
                        'accession_number': filing['accession_number'],
                        '10k_url': filing_docs['10k'],
                        'exhibit13_url': filing_docs['exhibit13'],
                        'is_ixbrl': is_ixbrl
                    })
                    logger.info(f"Added filing {filing['accession_number']} dated {filing['filing_date']}")
                
            except Exception as e:
                logger.error(f"Error processing filing {filing['accession_number']}: {e}")
            
            # Respect SEC rate limits (adjust the sleep duration if necessary)
            time.sleep(0.1)
            
        return results

    def _check_if_ixbrl(self, url: str) -> bool:
        """
        Check if a document is in iXBRL format by looking for the iXBRL namespace.
        Args:
            url: URL of the document to check.
        Returns:
            True if the document appears to be iXBRL, False otherwise.
        """
        logger.info(f"Checking if document at {url} is iXBRL")
        try:
            response = self.session.get(
                url,
                headers=self.headers,
                timeout=(10, 30)
            )
            response.raise_for_status()
            # Look for the iXBRL namespace in the document
            return 'xmlns:ix="http://www.xbrl.org/2013/inlineXBRL"' in response.text
        except requests.exceptions.RequestException as e:
            logger.error(f"Error checking iXBRL document: {e}")
            return False

# Example usage
if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Find Exhibit 13 documents for a company')
    parser.add_argument('ticker', type=str, help='Company ticker symbol')
    parser.add_argument('start_year', type=str, help='Start year (YYYY)')
    parser.add_argument('--email', type=str, required=True, help='Your email for SEC user agent')
    args = parser.parse_args()
    
    if not FMP_API_KEY:
        raise ValueError("FMP_API_KEY not found in environment variables")
    
    start_date = f"{args.start_year}-01-01"
    finder = EDGARExhibit13Finder(f"Custom Research Agent - Contact: {args.email}")
    
    try:
        results = finder.find_exhibit13_ixbrl(args.ticker.upper(), start_date=start_date)
        
        print("\nFound Exhibit 13 documents:")
        for result in results:
            print(f"\nFiling Date: {result['filing_date']}")
            print(f"Accession Number: {result['accession_number']}")
            print(f"10-K URL: {result['10k_url']}")
            print(f"Exhibit 13 URL: {result['exhibit13_url']}")
            print(f"Is iXBRL: {result['is_ixbrl']}")
    except Exception as e:
        logger.error(f"Error in main execution: {e}")