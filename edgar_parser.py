# analysis_project/edgar_parser.py
#!/usr/bin/env python3
import os
import logging
import re
import requests
import time
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from urllib.parse import urljoin
from urllib3.util.retry import Retry

# Import helper functions from your modules.
from unified_segmentation import get_filing_contents

# Load environment variables
load_dotenv()
FMP_API_KEY = os.getenv("FMP_API_KEY")
if not FMP_API_KEY:
    raise ValueError("FMP_API_KEY not found in environment variables")

# Configure logging
logger = logging.getLogger(__name__)

def find_context(soup, context_ref):
    """
    Fallback helper: returns a tag whose id equals context_ref and whose tag name
    (lowercased) contains 'context'.
    """
    return soup.find(lambda tag: tag.get("id") == context_ref and "context" in tag.name.lower())

class EDGARExhibit13Finder:
    """
    Retrieves company filings via EDGAR and locates the XML filing document.
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
                accession_match = re.search(r"accession-number=(\d{10}-\d{2}-\d{6})", id_elem.text)
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
                entries.append({
                    "accession_number": accession_number,
                    "filing_href": filing_href,
                    "filing_date": filing_date,
                })
        return {"filings": entries}
    
    def get_filing_detail(self, filing_url: str) -> dict:
        """
        Retrieves the filing detail page and locates the XML filing document.
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
                    if ("extracted" in description.lower() and 
                        "instance document" in description.lower() and 
                        "xbrl" in description.lower()):
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
            logger.warning(f"No XML filing document found in filing page: {filing_url}")
        return documents

class MetricsExtractor:
    """
    Extracts metrics from an SEC XML filing containing inline XBRL facts.
    The output for each year includes:
      - profit_desc
      - balance_sheet
      - segmentation
    """
    def __init__(self, user_agent: str, config: dict):
        if not config:
            raise ValueError("A valid metrics configuration must be provided. Terminating.")
        
        # Extract required mappings from the config.
        self.profit_desc_metrics = config.get("profit_desc_metrics")
        self.balance_sheet_metrics = config.get("balance_sheet_metrics")
        self.segmentation_mapping = config.get("segmentation_mapping")
        self.balance_sheet_categories = config.get("balance_sheet_categories")
        
        # Ensure all required config sections are present.
        if not (self.profit_desc_metrics and self.balance_sheet_metrics and 
                self.segmentation_mapping and self.balance_sheet_categories):
            raise ValueError("Incomplete metrics configuration. Please provide profit, balance sheet, segmentation mappings, and balance sheet categories.")

        self.config = config  # Save entire config if needed later.
        
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
    
    def parse_context(self, soup, context_ref):
        logger.debug(f"Parsing context: {context_ref}")
        context = soup.find("context", {"id": context_ref})
        if context:
            period_elem = context.find("period")
            if period_elem:
                start = period_elem.find("startDate")
                end = period_elem.find("endDate")
                instant = period_elem.find("instant")
                if start and end:
                    return {"period": f"{start.text.strip()} to {end.text.strip()}"}
                elif instant:
                    return {"period": f"As of {instant.text.strip()}"}
        context = find_context(soup, context_ref)
        if context:
            period_elem = context.find("period")
            if period_elem:
                start = period_elem.find("startDate")
                end = period_elem.find("endDate")
                instant = period_elem.find("instant")
                if start and end:
                    return {"period": f"{start.text.strip()} to {end.text.strip()}"}
                elif instant:
                    return {"period": f"As of {instant.text.strip()}"}
        year_pattern = r"(\d{4})"
        year_match = re.search(year_pattern, context_ref)
        if year_match:
            year = year_match.group(1)
            return {"period": f"{year}-01-01 to {year}-12-31"}
        return {}
    
    def process_mapping(self, soup, mapping):
        local = {}
        for metric_name, tag in mapping.items():
            elems = soup.find_all(tag)
            for elem in elems:
                try:
                    context_ref = elem.get("contextRef") or elem.get("contextref")
                    if not context_ref:
                        continue

                    # ---- START OF MODIFICATION ----
                    # Find the context associated with the fact
                    context = soup.find("context", {"id": context_ref})
                    if not context:
                        # Use the fallback helper if the primary find fails
                        context = find_context(soup, context_ref)
                    if not context:
                        logger.warning(f"Could not find context for ref: {context_ref}")
                        continue
                    
                    # **CRITICAL FIX**: Check if the context contains a <segment> element. 
                    # If it does, this fact is for a business segment, not the consolidated
                    # entity, and we should skip it for this general mapping.
                    if context.find("segment"):
                        continue
                    # ---- END OF MODIFICATION ----

                    numeric_value = float(elem.get_text(strip=True))
                    scale = elem.get("scale", "0")
                    if scale and scale != "0":
                        numeric_value *= 10 ** int(scale)
                    
                    context_data = self.parse_context(soup, context_ref)
                    period_text = context_data.get("period", "")
                    if not period_text:
                        continue
                    year_match = re.search(r"(\d{4})", period_text)
                    if not year_match:
                        continue
                    year = year_match.group(1)
                    if year not in local:
                        local[year] = {}
                    local[year][metric_name] = numeric_value
                except (ValueError, TypeError) as e:
                    logger.error(f"Error processing {metric_name}: {e}")
                    continue
        return local
    
    def process_segmentation(self, soup):
        seg_results = {}
        for seg_key, seg_info in self.segmentation_mapping.items():
            tag = seg_info["tag"]
            required = seg_info["explicitMembers"]
            elems = soup.find_all(tag)
            for elem in elems:
                try:
                    context_ref = elem.get("contextRef") or elem.get("contextref")
                    if not context_ref:
                        continue
                    context = soup.find("context", {"id": context_ref})
                    if not context:
                        continue
                    entity = context.find("entity")
                    if not entity:
                        continue
                    segment = entity.find("segment")
                    if not segment:
                        continue
                    explicit_members = segment.find_all(lambda t: "explicitmember" in t.name.lower())
                    criteria_met = True
                    for dim, expected in required.items():
                        match_found = any(
                            (exp.get("dimension") == dim and exp.get_text(strip=True) == expected)
                            for exp in explicit_members
                        )
                        if not match_found:
                            criteria_met = False
                            break
                    if not criteria_met:
                        continue
                    numeric_value = float(elem.get_text(strip=True))
                    scale = elem.get("scale", "0")
                    if scale and scale != "0":
                        numeric_value *= 10 ** int(scale)
                    context_data = self.parse_context(soup, context_ref)
                    period_text = context_data.get("period", "")
                    if not period_text:
                        continue
                    year_match = re.search(r"(\d{4})", period_text)
                    if not year_match:
                        continue
                    year = year_match.group(1)
                    if year not in seg_results:
                        seg_results[year] = {}
                    seg_results[year][seg_key] = seg_results[year].get(seg_key, 0) + numeric_value
                except Exception as e:
                    logger.error(f"Error processing segmentation {seg_key}: {e}")
                    continue
        return seg_results
    
    def extract_metrics(self, xml_url: str) -> dict:
        max_retries = 3
        attempt = 0
        content = None
        while attempt < max_retries:
            try:
                logger.info("Fetching filing data...")
                content, meta_info = get_filing_contents(xml_url)
                if content:
                    break
            except requests.exceptions.Timeout as e:
                attempt += 1
                logger.error(f"Timeout error fetching filing data (attempt {attempt}/{max_retries}): {e}")
                time.sleep(5)
        if not content:
            logger.error("No filing content retrieved after maximum retries.")
            return {}
        
        logger.info("Parsing XML filing...")
        soup = BeautifulSoup(content, "lxml-xml")
        results = {}
        
        profit_data = self.process_mapping(soup, self.profit_desc_metrics)
        balance_data = self.process_mapping(soup, self.balance_sheet_metrics)
        segmentation_data = self.process_segmentation(soup)
        
        # Retrieve balance_sheet_categories configuration from the config.
        balance_sheet_categories = self.balance_sheet_categories
        if not balance_sheet_categories:
            raise ValueError("Missing 'balance_sheet_categories' in configuration.")

        years = set(profit_data.keys()) | set(balance_data.keys()) | set(segmentation_data.keys())
        for year in years:
            results[year] = {
                "profit_desc": profit_data.get(year, {}),
                "balance_sheet": {
                    "assets": {},
                    "liabilities": {},
                    "shareholders_equity": {}
                },
                "segmentation": segmentation_data.get(year, {})
            }
            if year in balance_data:
                for metric_name, value in balance_data[year].items():
                    assigned = False
                    # Loop through each category (assets, liabilities, shareholders_equity) as defined in the config.
                    for category, metrics_list in balance_sheet_categories.items():
                        if metric_name in metrics_list:
                            results[year]["balance_sheet"][category][metric_name] = value
                            assigned = True
                            break
                    if not assigned:
                        # If the metric isn't found in any category, assign it to assets by default.
                        results[year]["balance_sheet"]["assets"][metric_name] = value
        logger.info(f"Extracted SEC metrics: {results}")
        return results