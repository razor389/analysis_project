#!/usr/bin/env python3
import os
import json
import logging
import re
import requests
from bs4 import BeautifulSoup
import time
import argparse
from dotenv import load_dotenv
import xml.etree.ElementTree as ET
from urllib.parse import urljoin
from urllib3.util.retry import Retry

# Assume these helper functions are provided by your unified_segmentation module.
# They should return the filing content (as text) and any related metadata.
from unified_segmentation import get_filing_contents, get_scale_factor

# Load environment variables
load_dotenv()
FMP_API_KEY = os.getenv("FMP_API_KEY")
if not FMP_API_KEY:
    raise ValueError("FMP_API_KEY not found in environment variables")

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

def find_context(soup, context_ref):
    """
    Fallback helper: returns a tag whose id equals context_ref and whose name
    (lowercased) contains 'context'.
    """
    return soup.find(lambda tag: tag.get("id") == context_ref and "context" in tag.name.lower())

class EDGARExhibit13Finder:
    """
    Retrieves company filings via EDGAR and locates the XML filing document.
    (Although the class name remains for compatibility, we now use the XML file.)
    """
    BASE_URL = "https://www.sec.gov"

    def __init__(self, user_agent: str):
        self.headers = {
            "User-Agent": user_agent,
            "Accept-Encoding": "gzip, deflate",
        }
        self.session = requests.Session()
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["HEAD", "GET", "OPTIONS"],
        )
        adapter = requests.adapters.HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("https://", adapter)

    def get_cik_from_ticker(self, ticker: str) -> str:
        url = f"https://financialmodelingprep.com/api/v3/balance-sheet-statement/{ticker}"
        params = {
            "period": "annual",
            "apikey": FMP_API_KEY,
            "limit": 1,
        }
        try:
            response = self.session.get(url, params=params, timeout=(10, 30))
            response.raise_for_status()
            data = response.json()
            if not data:
                raise ValueError(f"No data found for ticker {ticker}")
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

    def get_company_filings(self, cik: str) -> dict:
        url = urljoin(
            self.BASE_URL,
            f"/cgi-bin/browse-edgar?action=getcompany&CIK={cik}&type=10-K&dateb=&owner=exclude&start=0&count=40&output=atom",
        )
        response = self.session.get(url, headers=self.headers, timeout=(10, 30))
        response.raise_for_status()
        try:
            root = ET.fromstring(response.content)
        except ET.ParseError as e:
            logger.error("Error parsing XML from SEC response.")
            raise e

        entries = []
        ns = {"atom": "http://www.w3.org/2005/Atom"}
        for entry in root.findall("atom:entry", ns):
            accession_number = ""
            filing_href = ""
            filing_date = ""
            id_elem = entry.find("atom:id", ns)
            if id_elem is not None and id_elem.text:
                accession_match = re.search(
                    r"accession-number=(\d{10}-\d{2}-\d{6})", id_elem.text
                )
                if accession_match:
                    accession_number = accession_match.group(1)
            date_elem = entry.find("atom:updated", ns)
            if date_elem is not None and date_elem.text:
                filing_date = date_elem.text.split("T")[0]
            link_elem = entry.find("atom:link", ns)
            if link_elem is not None:
                href = link_elem.get("href")
                if href:
                    filing_href = urljoin(self.BASE_URL, href)
            if accession_number and filing_href:
                entries.append(
                    {
                        "accession_number": accession_number,
                        "filing_href": filing_href,
                        "filing_date": filing_date,
                    }
                )
        return {"filings": entries}

    def get_filing_detail(self, filing_url: str) -> dict:
        """
        Retrieves the filing detail (index) page and locates the XML filing document.
        It looks through the table rows for one whose description contains keywords
        like "extracted", "instance document", and "xbrl" (case-insensitive). If not found,
        it falls back to a CSS selector.
        """
        response = self.session.get(filing_url, headers=self.headers, timeout=(10, 30))
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        documents = {"xml": None}
        table = soup.find("table", class_="tableFile")
        if table:
            for row in table.find_all("tr"):
                cells = row.find_all("td")
                if len(cells) >= 3:
                    description = cells[1].get_text(strip=True)
                    logger.debug(f"Row description: {description}")
                    desc_lower = description.lower()
                    if (
                        "extracted" in desc_lower
                        and "instance document" in desc_lower
                        and "xbrl" in desc_lower
                    ):
                        document_link = cells[2].find("a")
                        if document_link:
                            href = document_link.get("href")
                            if href and href.lower().endswith(".xml"):
                                documents["xml"] = urljoin(self.BASE_URL, href)
                                break
        if not documents["xml"]:
            xml_link = soup.select_one('a[href$="_htm.xml"]')
            if xml_link:
                documents["xml"] = urljoin(self.BASE_URL, xml_link.get("href"))
        if not documents["xml"]:
            logger.warning(
                f"No XML filing document found in filing page: {filing_url}"
            )
        return documents

class InsuranceMetricsExtractor:
    """
    Extracts metrics from an XML filing that contains inline XBRL facts and context definitions.
    This version uses the XML parser.
    """
    def __init__(self, user_agent: str):
        self.headers = {
            "User-Agent": user_agent,
            "Accept-Encoding": "gzip, deflate",
        }
        self.session = requests.Session()
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["HEAD", "GET", "OPTIONS"],
        )
        adapter = requests.adapters.HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("https://", adapter)
        self.metrics = {
            "gross_revenues": "us-gaap:PremiumsEarnedNetPropertyAndCasualty",
            "investment_income": "us-gaap:InterestAndDividendIncomeOperating",
            "losses_and_expenses": "us-gaap:IncurredClaimsPropertyCasualtyAndLiability",
            "acquisition_costs": "us-gaap:DeferredPolicyAcquisitionCostAmortizationExpense",
            "underwriting_expenses": "us-gaap:OtherUnderwritingExpense",
            "assets": "us-gaap:LiabilitiesAndStockholdersEquity",
            "cash": "us-gaap:CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
        }

    def parse_context(self, soup, context_ref):
        """
        Parses a <context> element from the XML filing to extract period information.
        Looks for <startDate> and <endDate> (or <instant>) within the <period> element.
        """
        logger.debug(f"Parsing context: {context_ref}")
        context = soup.find("context", {"id": context_ref})
        if context:
            period_elem = context.find("period")
            if period_elem:
                start = period_elem.find("startDate")
                end = period_elem.find("endDate")
                instant = period_elem.find("instant")
                if start and end:
                    logger.debug(
                        f"Found period: {start.text.strip()} to {end.text.strip()}"
                    )
                    return {"period": f"{start.text.strip()} to {end.text.strip()}"}
                elif instant:
                    logger.debug(f"Found instant: {instant.text.strip()}")
                    return {"period": f"As of {instant.text.strip()}"}
        # Fallback using the helper
        context = find_context(soup, context_ref)
        if context:
            period_elem = context.find("period")
            if period_elem:
                start = period_elem.find("startDate")
                end = period_elem.find("endDate")
                instant = period_elem.find("instant")
                if start and end:
                    logger.debug(
                        f"Found period: {start.text.strip()} to {end.text.strip()}"
                    )
                    return {"period": f"{start.text.strip()} to {end.text.strip()}"}
                elif instant:
                    logger.debug(f"Found instant: {instant.text.strip()}")
                    return {"period": f"As of {instant.text.strip()}"}
        # Fallback: extract a year from context_ref itself.
        year_pattern = r"(\d{4})"
        year_match = re.search(year_pattern, context_ref)
        if year_match:
            year = year_match.group(1)
            logger.debug(f"Falling back to year {year} from context ref")
            return {"period": f"{year}-01-01 to {year}-12-31"}
        logger.debug("No period information found")
        return {}

    def extract_metrics(self, xml_url: str) -> dict:
        """
        Retrieves the XML filing from xml_url, saves a debug copy,
        parses it using the lxml-xml parser, and extracts the desired metrics.
        """
        content, meta_info = get_filing_contents(xml_url)
        if not content:
            logger.error("No filing content retrieved.")
            return {}
        
        # with open("debug_filing.xml", "w", encoding="utf-8") as f:
        #     f.write(content)
        # logger.info("Saved raw filing XML to debug_filing.xml")
        
        logger.info("Parsing document using lxml-xml parser...")
        soup = BeautifulSoup(content, "lxml-xml")
        results = {}
        
        for metric_name, metric_tag in self.metrics.items():
            logger.info(f"Processing metric: {metric_name} (tag: {metric_tag})")
            # Here we search for elements by tag name.
            xbrl_elements = soup.find_all(metric_tag)
            logger.info(f"Found {len(xbrl_elements)} elements for metric '{metric_name}'")
            
            for elem in xbrl_elements:
                raw_value = elem.get_text(strip=True)
                logger.debug(f"Raw value for {metric_name}: '{raw_value}'")
                try:
                    # Since the value is now plain (no commas, parentheses, etc.), we simply convert.
                    numeric_value = float(raw_value)
                    
                    scale = elem.get("scale", "0")
                    if scale and scale != "0":
                        numeric_value *= 10 ** int(scale)
                    logger.debug(f"Final numeric value: {numeric_value}")
                    
                    context_ref = elem.get("contextRef") or elem.get("contextref")
                    if not context_ref:
                        logger.debug("No context reference found")
                        continue
                    logger.debug(f"Found contextRef: {context_ref}")
                    
                    context_data = self.parse_context(soup, context_ref)
                    logger.debug(f"Context data for {context_ref}: {context_data}")
                    period_text = context_data.get("period", "")
                    if not period_text:
                        logger.debug("No period text found")
                        continue
                    
                    # Extract the year from the startDate of the period.
                    # We assume the startDate is in the format YYYY-MM-DD.
                    year_match = re.search(r"(\d{4})", period_text)
                    if not year_match:
                        logger.debug(f"No year found in period text: {period_text}")
                        continue
                    year = year_match.group(1)
                    
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
    parser = argparse.ArgumentParser(description="Extract insurance metrics from SEC filings")
    parser.add_argument("ticker", type=str, help="Company ticker symbol")
    parser.add_argument("start_year", type=str, help="Start year (YYYY)")
    parser.add_argument("--email", type=str, required=True, help="Your email for SEC user agent")
    parser.add_argument("--output", type=str, help="Output JSON file")
    args = parser.parse_args()
    
    finder = EDGARExhibit13Finder(f"Insurance Research - Contact: {args.email}")
    extractor = InsuranceMetricsExtractor(f"Insurance Research - Contact: {args.email}")
    
    try:
        cik = finder.get_cik_from_ticker(args.ticker.upper())
        logger.info(f"Retrieved CIK for {args.ticker.upper()}: {cik}")
        filings_data = finder.get_company_filings(cik)
        if not filings_data.get("filings"):
            logger.error("No filings found.")
            return
        
        all_results = {}
        for filing in filings_data["filings"]:
            if filing["filing_date"] < f"{args.start_year}-01-01":
                continue
            
            filing_docs = finder.get_filing_detail(filing["filing_href"])
            xml_url = filing_docs.get("xml")
            if xml_url:
                logger.info(f"Processing XML filing from filing dated {filing['filing_date']}")
                year_results = extractor.extract_metrics(xml_url)
                for year, metrics in year_results.items():
                    if year not in all_results:
                        all_results[year] = {}
                    all_results[year].update(metrics)
            else:
                logger.warning(
                    f"XML filing document not found for filing dated {filing['filing_date']}"
                )
            time.sleep(0.1)
        
        ordered_results = dict(sorted(all_results.items(), reverse=True))
        output_file = args.output or f"{args.ticker.lower()}_insurance_metrics.json"
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(ordered_results, f, indent=2)
        logger.info(f"Results saved to {output_file}")
    except Exception as e:
        logger.error(f"Error: {e}")
        raise

if __name__ == "__main__":
    main()
