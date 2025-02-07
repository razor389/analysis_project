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
    Fallback helper: returns a tag whose id equals context_ref and whose tag name
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
        Retrieves the filing detail (index) page and locates the XML filing document.
        It scans table rows for one whose description contains keywords such as
        "extracted", "instance document", and "xbrl" (case-insensitive). Falls back
        to a CSS selector.
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
                    if ("extracted" in desc_lower and
                        "instance document" in desc_lower and
                        "xbrl" in desc_lower):
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

class InsuranceMetricsExtractor:
    """
    Extracts metrics from an XML filing that contains inline XBRL facts and context definitions.
    The results are organized per year into three sections: "profit_desc", "balance_sheet", and
    "segmentation". The balance_sheet section is subdivided into "assets", "liabilities", and
    "shareholders_equity".
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
        # Profit & loss metrics (profit_desc)
        self.profit_desc_metrics = {
            "gross_revenues": "us-gaap:PremiumsEarnedNetPropertyAndCasualty",
            "investment_income": "us-gaap:InterestAndDividendIncomeOperating",
            "losses_and_expenses": "us-gaap:IncurredClaimsPropertyCasualtyAndLiability",
            "acquisition_costs": "us-gaap:DeferredPolicyAcquisitionCostAmortizationExpense",
            "underwriting_expenses": "us-gaap:OtherUnderwritingExpense",
            "taxes": "us-gaap:IncomeTaxExpenseBenefit",
            "interest_expenses": "us-gaap:InterestExpenseDebt",
            "service_expenses": "pgr:NonInsuranceServiceExpenses"
        }
        # Balance sheet metrics
        self.balance_sheet_metrics = {
            "assets": "us-gaap:Assets",
            "cash": "us-gaap:CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
            "fixed_income": "pgr:DebtSecuritiesAvailableforsaleFixedMaturities",
            "preferred_stocks": "pgr:EquitySecuritiesFVNINonredeemablePreferredStock",
            "common_equities": "pgr:EquitySecuritiesFVNICommonEquities",
            "short_term": "us-gaap:ShortTermInvestments",
            "accrued_investment_income": "us-gaap:AccruedInvestmentIncomeReceivable",
            "premiums_receivable": "us-gaap:PremiumsReceivableAtCarryingValue",
            "reinsurance_recoverables": "us-gaap:ReinsuranceRecoverables",
            "prepaid_reinsurance_premiums": "us-gaap:PrepaidReinsurancePremiums",
            "deferred_acquisition_costs": "us-gaap:DeferredPolicyAcquisitionCosts",
            "income_taxes": "us-gaap:DeferredTaxAssetsLiabilitiesNet",
            "property_and_equipment": "us-gaap:PropertyPlantAndEquipmentNet",
            "goodwill": "us-gaap:Goodwill",
            "intangibles": "us-gaap:IntangibleAssetsNetExcludingGoodwill",
            "other_assets": "us-gaap:OtherAssets",
            "liabilities": "us-gaap:Liabilities",
            "unearned_premiums": "us-gaap:UnearnedPremiums",
            "loss_reserves": "us-gaap:LiabilityForClaimsAndClaimsAdjustmentExpense",
            "accounts_payable": "us-gaap:AccountsPayableAndAccruedLiabilitiesCurrentAndNoncurrent",
            "income_taxes_liabilities": "us-gaap:DeferredTaxLiabilities",
            "debt": "us-gaap:DebtLongtermAndShorttermCombinedAmount",
            "shareholders_equity": "us-gaap:StockholdersEquity",
            "common_stock": "us-gaap:CommonStockValueOutstanding",
            "additional_capital": "us-gaap:AdditionalPaidInCapitalCommonStock",
            "unamortized_restricted_stock": "us-gaap:PreferredStockValueOutstanding",
            "unrealized_net_capital_gains_losses": "us-gaap:AccumulatedOtherComprehensiveIncomeLossAvailableForSaleSecuritiesAdjustmentNetOfTax",
            "hedges": "us-gaap:AccumulatedOtherComprehensiveIncomeLossCumulativeChangesInNetGainLossFromCashFlowHedgesEffectNetOfTax",
            "foreign_currency_translation": "pgr:AccumulatedOtherComprehensiveIncomeLossAttributableToNoncontrollingInterestNetOfTax",
            "retained_earnings": "us-gaap:RetainedEarningsAccumulatedDeficit"
        }
        # Map each balance sheet metric to a subsection.
        self.balance_sheet_mapping = {
            # Assets subsection
            "assets": "assets",
            "cash": "assets",
            "fixed_income": "assets",
            "preferred_stocks": "assets",
            "common_equities": "assets",
            "short_term": "assets",
            "accrued_investment_income": "assets",
            "premiums_receivable": "assets",
            "reinsurance_recoverables": "assets",
            "prepaid_reinsurance_premiums": "assets",
            "deferred_acquisition_costs": "assets",
            "income_taxes": "assets",
            "property_and_equipment": "assets",
            "goodwill": "assets",
            "intangibles": "assets",
            "other_assets": "assets",
            # Liabilities subsection
            "liabilities": "liabilities",
            "unearned_premiums": "liabilities",
            "loss_reserves": "liabilities",
            "accounts_payable": "liabilities",
            "income_taxes_liabilities": "liabilities",
            "debt": "liabilities",
            # Shareholders' Equity subsection
            "shareholders_equity": "shareholders_equity",
            "common_stock": "shareholders_equity",
            "additional_capital": "shareholders_equity",
            "unamortized_restricted_stock": "shareholders_equity",
            "unrealized_net_capital_gains_losses": "shareholders_equity",
            "hedges": "shareholders_equity",
            "foreign_currency_translation": "shareholders_equity",
            "retained_earnings": "shareholders_equity",
        }
        # Segmentation mapping.
        # For each segmentation key, we specify:
        #   - "tag": the fact tag to search (here "us-gaap:Revenues")
        #   - "explicitMembers": a dictionary of required dimension/value pairs.
        self.segmentation_mapping = {
            "personal_lines_agency": {
                "tag": "us-gaap:Revenues",
                "explicitMembers": {
                    "srt:ProductOrServiceAxis": "pgr:UnderwritingOperationsMember",
                    "us-gaap:StatementBusinessSegmentsAxis": "pgr:PersonalLinesSegmentMember",
                    "us-gaap:SubsegmentsAxis": "pgr:AgencyChannelMember"
                }
            },
            "personal_lines_direct": {
                "tag": "us-gaap:Revenues",
                "explicitMembers": {
                    "srt:ProductOrServiceAxis": "pgr:UnderwritingOperationsMember",
                    "us-gaap:StatementBusinessSegmentsAxis": "pgr:PersonalLinesSegmentMember",
                    "us-gaap:SubsegmentsAxis": "pgr:DirectChannelMember"
                }
            },
            "commercial_lines": {
                "tag": "us-gaap:Revenues",
                "explicitMembers": {
                    "srt:ConsolidationItemsAxis": "us-gaap:OperatingSegmentsMember",
                    "srt:ProductOrServiceAxis": "pgr:UnderwritingOperationsMember",
                    "us-gaap:StatementBusinessSegmentsAxis": "pgr:CommercialLinesSegmentMember"
                }
            },
            "property_lines": {
                "tag": "us-gaap:Revenues",
                "explicitMembers": {
                    "srt:ConsolidationItemsAxis": "us-gaap:OperatingSegmentsMember",
                    "srt:ProductOrServiceAxis": "pgr:UnderwritingOperationsMember",
                    "us-gaap:StatementBusinessSegmentsAxis": "pgr:PropertySegmentMember"
                }
            }
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
                    logger.debug(f"Found period: {start.text.strip()} to {end.text.strip()}")
                    return {"period": f"{start.text.strip()} to {end.text.strip()}"}
                elif instant:
                    logger.debug(f"Found instant: {instant.text.strip()}")
                    return {"period": f"As of {instant.text.strip()}"}
        context = find_context(soup, context_ref)
        if context:
            period_elem = context.find("period")
            if period_elem:
                start = period_elem.find("startDate")
                end = period_elem.find("endDate")
                instant = period_elem.find("instant")
                if start and end:
                    logger.debug(f"Found period: {start.text.strip()} to {end.text.strip()}")
                    return {"period": f"{start.text.strip()} to {end.text.strip()}"}
                elif instant:
                    logger.debug(f"Found instant: {instant.text.strip()}")
                    return {"period": f"As of {instant.text.strip()}"}
        year_pattern = r"(\d{4})"
        year_match = re.search(year_pattern, context_ref)
        if year_match:
            year = year_match.group(1)
            logger.debug(f"Falling back to year {year} from context ref")
            return {"period": f"{year}-01-01 to {year}-12-31"}
        logger.debug("No period information found")
        return {}

    def process_mapping(self, soup, mapping):
        """
        Helper to process a mapping of fact metrics.
        Returns a dictionary keyed by year with the metric values.
        """
        local = {}
        for metric_name, tag in mapping.items():
            elems = soup.find_all(tag)
            for elem in elems:
                try:
                    numeric_value = float(elem.get_text(strip=True))
                    scale = elem.get("scale", "0")
                    if scale and scale != "0":
                        numeric_value *= 10 ** int(scale)
                    context_ref = elem.get("contextRef") or elem.get("contextref")
                    if not context_ref:
                        continue
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
        """
        Processes segmentation data. For each segmentation key in the segmentation_mapping,
        this function searches for facts with the given tag (typically "us-gaap:Revenues"),
        then looks up their context. For each fact, it finds the corresponding <context> element,
        and if that context contains an <entity><segment> element, it collects all <xbrldi:explicitMember>
        children. If the set of explicitMember dimensions and values exactly match the required criteria,
        the fact's value is recorded for that segmentation key (for the year determined from the period).
        If multiple facts match for a given segmentation key and year, their values are summed.
        """
        seg_results = {}
        for seg_key, seg_info in self.segmentation_mapping.items():
            tag = seg_info["tag"]
            required = seg_info["explicitMembers"]
            # Find all facts with the given tag.
            elems = soup.find_all(tag)
            for elem in elems:
                try:
                    context_ref = elem.get("contextRef") or elem.get("contextref")
                    if not context_ref:
                        continue
                    # Look up the corresponding context element.
                    context = soup.find("context", {"id": context_ref})
                    if not context:
                        continue
                    # Find the segment element within the entity.
                    entity = context.find("entity")
                    if not entity:
                        continue
                    segment = entity.find("segment")
                    if not segment:
                        continue
                    # Gather all explicitMember elements (namespace prefixes might vary).
                    explicit_members = segment.find_all(lambda t: "explicitmember" in t.name.lower())
                    # For each required dimension, check if an explicitMember with that dimension and value is present.
                    criteria_met = True
                    for dim, expected in required.items():
                        # We assume that the explicitMember element has an attribute "dimension"
                        # and its text equals the expected value.
                        match_found = any(
                            (exp.get("dimension") == dim and exp.get_text(strip=True) == expected)
                            for exp in explicit_members
                        )
                        if not match_found:
                            criteria_met = False
                            break
                    if not criteria_met:
                        continue
                    # If criteria are met, get the fact value.
                    numeric_value = float(elem.get_text(strip=True))
                    scale = elem.get("scale", "0")
                    if scale and scale != "0":
                        numeric_value *= 10 ** int(scale)
                    # Get the year from the context period.
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
                    # Sum values if multiple facts match for the same segmentation key.
                    seg_results[year][seg_key] = seg_results[year].get(seg_key, 0) + numeric_value
                except Exception as e:
                    logger.error(f"Error processing segmentation {seg_key}: {e}")
                    continue
        return seg_results

    def extract_metrics(self, xml_url: str) -> dict:
        """
        Retrieves the XML filing from xml_url, retrying a few times if a timeout occurs,
        then parses it using the lxml-xml parser and extracts the desired metrics into three groups:
        profit_desc, balance_sheet (subdivided), and segmentation.
        """
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
        
        logger.info("Parsing document using lxml-xml parser...")
        soup = BeautifulSoup(content, "lxml-xml")
        results = {}
        
        profit_data = self.process_mapping(soup, self.profit_desc_metrics)
        balance_data = self.process_mapping(soup, self.balance_sheet_metrics)
        segmentation_data = self.process_segmentation(soup)
        
        # Combine profit and balance sheet data into final results.
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
                    subsection = self.balance_sheet_mapping.get(metric_name)
                    if subsection:
                        results[year]["balance_sheet"][subsection][metric_name] = value
                    else:
                        results[year]["balance_sheet"][metric_name] = value
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
                for year, data in year_results.items():
                    if year not in all_results:
                        all_results[year] = data
                    else:
                        all_results[year]["profit_desc"].update(data.get("profit_desc", {}))
                        for subsec, subdata in data.get("balance_sheet", {}).items():
                            if subsec not in all_results[year]["balance_sheet"]:
                                all_results[year]["balance_sheet"][subsec] = subdata
                            else:
                                all_results[year]["balance_sheet"][subsec].update(subdata)
                        all_results[year]["segmentation"].update(data.get("segmentation", {}))
            else:
                logger.warning(f"XML filing document not found for filing dated {filing['filing_date']}")
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
