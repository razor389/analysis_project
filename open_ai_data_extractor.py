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
import datetime

# Import helper functions from your modules.
from financial_data_preprocessor import process_financial_statements
from unified_segmentation import get_filing_contents
from utils import get_company_profile, get_yahoo_ticker, get_yearly_high_low_yahoo

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

def get_financials(symbol: str, statement_type: str, frequency: str):
    """
    Fetch financial data from Financial Modeling Prep.
    statement_type can be one of: "bs" (Balance Sheet), "ic" (Income Statement), "cf" (Cash Flow)
    frequency is "annual" or "quarter"
    """
    endpoints = {
        "ic": "income-statement",
        "bs": "balance-sheet-statement",
        "cf": "cash-flow-statement",
        "bs-ar": "balance-sheet-statement-as-reported"
    }
    endpoint = endpoints.get(statement_type)
    if not endpoint:
        print(f"Invalid statement type: {statement_type}")
        return

    url = f"https://financialmodelingprep.com/api/v3/{endpoint}/{symbol}?period={frequency}&apikey={FMP_API_KEY}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"Error fetching data: {e}")

def get_basic_financials(symbol: str):
    """
    Fetch key metrics from FMP as a stand-in for basic financials.
    """
    url = f"https://financialmodelingprep.com/api/v3/key-metrics/{symbol}?apikey={FMP_API_KEY}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        return {"keyMetrics": data}
    except Exception as e:
        print(f"Error fetching basic financials for {symbol}: {e}")

def extract_series_values_by_year(basic_data: dict, key: str) -> dict:
    """
    Extracts series values (e.g. P/E ratio) by year from FMP key-metrics data.
    """
    field_map = {"pe": "peRatio"}
    fmp_field = field_map.get(key, key)
    result = {}
    key_metrics = basic_data.get("keyMetrics", [])
    for entry in key_metrics:
        date_str = entry.get("date")
        value = entry.get(fmp_field)
        if date_str and value is not None:
            year_str = date_str.split("-")[0]
            try:
                year_int = int(year_str)
                result[year_int] = value
            except ValueError:
                pass
    return result

def extract_yoy_data(symbol: str, years: list, segmentation_data: dict, profile: dict):
    """
    Extracts FMP data across years (company description, earnings, etc.).
    NOTE: We use SEC data for the balance sheet.
    Now, we also compute the FMP tax rate as:
         incomeTaxExpense / incomeBeforeTax
    """
    bs_data = get_financials(symbol, 'bs', "annual")
    ic_data = get_financials(symbol, 'ic', "annual")
    cf_data = get_financials(symbol, 'cf', "annual")
    basic_data = get_basic_financials(symbol)
    
    try:
        ic_data, bs_data, cf_data = process_financial_statements(
            ticker=symbol,
            ic_data=ic_data,
            bs_data=bs_data,
            cf_data=cf_data
        )
    except Exception as e:
        logger.error(f"Error processing financial statements: {e}")
        return {}
    
    pe_by_year = extract_series_values_by_year(basic_data, 'pe')
    
    def by_year_dict(data):
        res = {}
        if isinstance(data, list):
            for item in data:
                date_str = item.get("date")
                if date_str:
                    y = date_str.split('-')[0]
                    try:
                        year_int = int(y)
                        res[year_int] = item
                    except Exception:
                        pass
            return res
        elif isinstance(data, dict):
            for item in data.get('financials', []):
                date_str = item.get('date')
                if date_str:
                    y = date_str.split('-')[0]
                    try:
                        year_int = int(y)
                        res[year_int] = item
                    except Exception:
                        pass
            return res
        return res

    bs_by_year = by_year_dict(bs_data)
    ic_by_year = by_year_dict(ic_data)
    cf_by_year = by_year_dict(cf_data)
    
    results = {}
    prev_shares_outstanding = None
    yahoo_symbol = get_yahoo_ticker(profile)
    
    for year in years:
        bs = bs_by_year.get(year, {})
        ic = ic_by_year.get(year, {})
        cf = cf_by_year.get(year, {})
        
        filing_url = bs.get('finalLink')
        shares_outstanding = ic.get('weightedAverageShsOutDil')
        net_profit = ic.get('netIncome')
        revenues = ic.get('revenue')
        diluted_eps = ic.get('epsdiluted')
        operating_income = ic.get('operatingIncome')
        shares_outstanding = ic.get('weightedAverageShsOutDil')
        operating_eps = operating_income / shares_outstanding if (operating_income is not None and shares_outstanding) else None
        dividends_paid = -1 * cf.get('dividendsPaid', 0)
        shareholder_equity = bs.get('totalStockholdersEquity')
        buyback = None
        if prev_shares_outstanding is not None and shares_outstanding and yahoo_symbol:
            share_change = prev_shares_outstanding - shares_outstanding
            high, low = get_yearly_high_low_yahoo(yahoo_symbol, year)
            avg_price = (high + low) / 2 if (high and low) else None
            if avg_price:
                buyback = share_change * avg_price
        book_value_per_share = (shareholder_equity / shares_outstanding) if (shareholder_equity and shares_outstanding and shares_outstanding != 0) else None
        
        # Compute FMP tax rate using incomeTaxExpense and incomeBeforeTax.
        provision_for_taxes = ic.get('incomeTaxExpense')
        pretax_income = ic.get('incomeBeforeTax')
        fmp_tax_rate = (provision_for_taxes / pretax_income) if (provision_for_taxes is not None and pretax_income and pretax_income != 0) else None
        
        company_description = {
            "net_profit": net_profit,
            "diluted_eps": diluted_eps,
            "operating_eps": operating_eps,
            "pe_ratio": pe_by_year.get(year),
            "price_low": (get_yearly_high_low_yahoo(yahoo_symbol, year)[1]
                          if get_yearly_high_low_yahoo(yahoo_symbol, year) else None),
            "price_high": (get_yearly_high_low_yahoo(yahoo_symbol, year)[0]
                           if get_yearly_high_low_yahoo(yahoo_symbol, year) else None),
            "dividends_paid": dividends_paid,
            "dividends_per_share": (dividends_paid / shares_outstanding
                                    if (dividends_paid and shares_outstanding and shares_outstanding != 0) else None),
            "avg_dividend_yield": None,  # computed later in unified mapping
            "shares_outstanding": shares_outstanding,
            "buyback": buyback,
            "share_equity": shareholder_equity,
            "book_value_per_share": book_value_per_share
        }
        
        analysis = {
            "revenue": revenues,
            "tax_rate": fmp_tax_rate  # FMP tax rate computed above
        }
        profit_description = {
            "gross_revenues": revenues,
            "filing_url": filing_url
        }
        
        results[year] = {
            "company_description": company_description,
            "analysis": analysis,
            "profit_description": profit_description
            # FMP balance sheet data is ignored; we use SEC balance sheet.
        }
        prev_shares_outstanding = shares_outstanding
        
    return results

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
        self.profit_desc_metrics = {
            "gross_revenues": "us-gaap:PremiumsEarnedNetPropertyAndCasualty",
            "investment_income": "us-gaap:InterestAndDividendIncomeOperating",
            "losses_and_expenses": "us-gaap:IncurredClaimsPropertyCasualtyAndLiability",
            "acquisition_costs": "us-gaap:DeferredPolicyAcquisitionCostAmortizationExpense",
            "underwriting_expenses": "us-gaap:OtherUnderwritingExpense",
            "service_expenses": "pgr:NonInsuranceServiceExpenses",
            "taxes": "us-gaap:IncomeTaxExpenseBenefit",
            "interest_expenses": "us-gaap:InterestExpenseDebt"
        }
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
            "debt": "us-gaap:DebtLongtermAndShorttermCombinedAmount",
            "shareholders_equity": "us-gaap:StockholdersEquity"
        }
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
                    if metric_name in ["assets", "cash", "fixed_income", "preferred_stocks", "common_equities",
                                       "short_term", "accrued_investment_income", "premiums_receivable",
                                       "reinsurance_recoverables", "prepaid_reinsurance_premiums",
                                       "deferred_acquisition_costs", "income_taxes", "property_and_equipment",
                                       "goodwill", "intangibles", "other_assets"]:
                        results[year]["balance_sheet"]["assets"][metric_name] = value
                    elif metric_name in ["liabilities", "unearned_premiums", "loss_reserves", "accounts_payable", "debt"]:
                        results[year]["balance_sheet"]["liabilities"][metric_name] = value
                    elif metric_name in ["shareholders_equity"]:
                        results[year]["balance_sheet"]["shareholders_equity"][metric_name] = value
                    else:
                        results[year]["balance_sheet"]["assets"][metric_name] = value
        logger.info(f"Extracted SEC metrics: {results}")
        return results

def create_unified_year_output(year, fmp_data, sec_data):
    """
    Remaps FMP and SEC data into the unified structure:
    
    "company_description": {
       "net_profit",
       "diluted_eps",
       "operating_eps",
       "pe_ratio",
       "price_low",
       "price_high",
       "dividends_paid",
       "dividends_per_share",
       "avg_dividend_yield",
       "shares_outstanding",
       "buyback",
       "share_equity",
       "book_value_per_share",
       "assets" (from SEC balance sheet),
       "return_on_equity" = net_profit / share_equity,
       "return_on_assets" = net_profit / assets,
       "leverage_ratio" = assets / share_equity
    },
    "analysis": {
       "premium_earned" = gross_revenues from SEC profit_desc,
       "benefit_claims" = losses_and_expenses from SEC profit_desc,
       "gross_underwriting_profit" = premium_earned - benefit_claims,
       "underwriting_yield_on_asset" = gross_underwriting_profit / assets,
       "investment_income" = investment_income from SEC profit_desc,
       "investment_yield_on_asset" = investment_income / assets,
       "non_claim_expenses" = acquisition_costs + underwriting_expenses,
       "expense_yield_on_asset" = non_claim_expenses / assets,
       "tax_rate" = FMP tax rate,
       "premium_equity_ratio" = premium_earned / share_equity
    },
    "balance_sheet": { SEC balance sheet data (total vs. breakdown) },
    "profit_description": {
       "gross_revenues": { "total": gross_revenues, "breakdown": segmentation breakdown },
       "investment_income",
       "internal_costs": {
         "total" = losses_and_expenses + acquisition_costs + underwriting_expenses + service_expenses,
         "breakdown": { "losses_and_expenses", "acquisition_costs", "underwriting_expenses", "service_expenses" }
       },
       "operating_margin": {
         "total" = gross_revenues + investment_income - total internal costs,
         "breakdown": {
            "underwriting": (sum(gross_revenue segments) - (losses_and_expenses + acquisition_costs + underwriting_expenses)),
            "pretax_combined_ratio": (sum(gross_revenue segments) - underwriting_margin) / (sum(gross_revenue segments)),
            "pretax_insurance_yield_on_equity": underwriting_margin / share_equity,
            "pretax_return_on_equity": operating_margin_total / share_equity
         }
       },
       "external_costs": {
         "total": taxes + interest_expenses,
         "breakdown": { "taxes", "interest_expenses" }
       },
       "earnings": operating_margin_total - external_costs_total,
       "equity_employed": share_equity,
       "shares_repurchased": buyback,
       "filing_url": filing_url
    },
    "segmentation": {
       "total": (sum of segmentation values),
       "breakdown": segmentation dict
    }
    """
    # FMP company data:
    fmp_comp = fmp_data.get("company_description", {})
    net_profit           = fmp_comp.get("net_profit")
    diluted_eps          = fmp_comp.get("diluted_eps")
    operating_eps        = fmp_comp.get("operating_eps")
    pe_ratio             = fmp_comp.get("pe_ratio")
    price_low            = fmp_comp.get("price_low")
    price_high           = fmp_comp.get("price_high")
    dividends_paid       = fmp_comp.get("dividends_paid")
    dividends_per_share  = fmp_comp.get("dividends_per_share")
    shares_outstanding   = fmp_comp.get("shares_outstanding")
    buyback              = fmp_comp.get("buyback")
    share_equity         = fmp_comp.get("share_equity")
    book_value_per_share = fmp_comp.get("book_value_per_share")
    filing_url = fmp_data.get("profit_description", {}).get("filing_url")

    # SEC balance sheet (use total assets):
    sec_balance = sec_data.get("balance_sheet", {})
    sec_assets = sec_balance.get("assets", {})
    assets = sec_assets.get("assets") or sec_assets.get("total_assets")
    
    # Derived ratios:
    return_on_equity = (net_profit / share_equity) if (net_profit is not None and share_equity) else None
    return_on_assets = (net_profit / assets) if (net_profit is not None and assets) else None
    leverage_ratio   = (assets / share_equity) if (assets is not None and share_equity) else None
    avg_dividend_yield = (dividends_per_share / ((price_low + price_high) / 2)
                          if (dividends_per_share and price_low and price_high and ((price_low+price_high)/2) != 0)
                          else None)
    
    company_description = {
        "net_profit": net_profit,
        "diluted_eps": diluted_eps,
        "operating_eps": operating_eps,
        "pe_ratio": pe_ratio,
        "price_low": price_low,
        "price_high": price_high,
        "dividends_paid": dividends_paid,
        "dividends_per_share": dividends_per_share,
        "avg_dividend_yield": avg_dividend_yield,
        "shares_outstanding": shares_outstanding,
        "buyback": buyback,
        "share_equity": share_equity,
        "book_value_per_share": book_value_per_share,
        "assets": assets,
        "return_on_equity": return_on_equity,
        "return_on_assets": return_on_assets,
        "leverage_ratio": leverage_ratio
    }
    
    # SEC profit data:
    sec_profit = sec_data.get("profit_desc", {})
    premium_earned    = sec_profit.get("gross_revenues")
    benefit_claims    = sec_profit.get("losses_and_expenses")
    acquisition_costs = sec_profit.get("acquisition_costs")
    underwriting_expenses = sec_profit.get("underwriting_expenses")
    service_expenses  = sec_profit.get("service_expenses")
    taxes             = sec_profit.get("taxes")
    interest_expenses = sec_profit.get("interest_expenses")
    investment_income = sec_profit.get("investment_income")
    
    gross_underwriting_profit = (premium_earned - benefit_claims
                                 if premium_earned is not None and benefit_claims is not None
                                 else None)
    underwriting_yield_on_asset = (gross_underwriting_profit / assets
                                   if gross_underwriting_profit is not None and assets
                                   else None)
    non_claim_expenses = (acquisition_costs + underwriting_expenses
                          if acquisition_costs is not None and underwriting_expenses is not None
                          else None)
    expense_yield_on_asset = (non_claim_expenses / assets
                              if non_claim_expenses is not None and assets
                              else None)
    premium_equity_ratio = (premium_earned / share_equity
                            if premium_earned is not None and share_equity
                            else None)
    
    # Note: tax rate now comes from FMP. Retrieve it from FMP analysis:
    fmp_analysis = fmp_data.get("analysis", {})
    fmp_tax_rate = fmp_analysis.get("tax_rate")
    
    analysis = {
        "premium_earned": premium_earned,
        "benefit_claims": benefit_claims,
        "gross_underwriting_profit": gross_underwriting_profit,
        "underwriting_yield_on_asset": underwriting_yield_on_asset,
        "investment_income": investment_income,
        "investment_yield_on_asset": (investment_income / assets if investment_income is not None and assets else None),
        "non_claim_expenses": non_claim_expenses,
        "expense_yield_on_asset": expense_yield_on_asset,
        "tax_rate": fmp_tax_rate,
        "premium_equity_ratio": premium_equity_ratio
    }
    
    internal_costs_total = ((benefit_claims or 0) +
                            (acquisition_costs or 0) +
                            (underwriting_expenses or 0) +
                            (service_expenses or 0))
    operating_margin_total = ((premium_earned or 0) + (investment_income or 0) - internal_costs_total)
    external_costs_total = ((taxes or 0) + (interest_expenses or 0))
    earnings = operating_margin_total - external_costs_total
    
    seg_breakdown = sec_data.get("segmentation", {})
    gross_revenue_segments = sum(seg_breakdown.values()) if seg_breakdown else None
    underwriting_margin = (gross_revenue_segments - ((benefit_claims or 0) +
                                                     (acquisition_costs or 0) +
                                                     (underwriting_expenses or 0))
                           if gross_revenue_segments is not None else None)
    pretax_combined_ratio = ((gross_revenue_segments - underwriting_margin) / gross_revenue_segments
                             if gross_revenue_segments and gross_revenue_segments != 0 else None)
    pretax_insurance_yield_on_equity = (underwriting_margin / share_equity
                                        if underwriting_margin is not None and share_equity else None)
    pretax_return_on_equity = (operating_margin_total / share_equity
                               if share_equity and share_equity != 0 else None)
    
    profit_description = {
        "gross_revenues": {
            "total": premium_earned,
            "breakdown": seg_breakdown
        },
        "investment_income": investment_income,
        "internal_costs": {
            "total": internal_costs_total,
            "breakdown": {
                "losses_and_expenses": benefit_claims,
                "acquisition_costs": acquisition_costs,
                "underwriting_expenses": underwriting_expenses,
                "service_expenses": service_expenses
            }
        },
        "operating_margin": {
            "total": operating_margin_total,
            "breakdown": {
                "underwriting": underwriting_margin,
                "pretax_combined_ratio": pretax_combined_ratio,
                "pretax_insurance_yield_on_equity": pretax_insurance_yield_on_equity,
                "pretax_return_on_equity": pretax_return_on_equity
            }
        },
        "external_costs": {
            "total": external_costs_total,
            "breakdown": {
                "taxes": taxes,
                "interest_expenses": interest_expenses
            }
        },
        "earnings": earnings,
        "equity_employed": share_equity,
        "shares_repurchased": buyback,
        "filing_url": filing_url
    }
    
    seg_total = sum(seg_breakdown.values()) if seg_breakdown else None
    segmentation = {
        "total": seg_total,
        "breakdown": seg_breakdown
    }
    
    unified = {
        "company_description": company_description,
        "analysis": analysis,
        "balance_sheet": sec_data.get("balance_sheet", {}),
        "profit_description": profit_description,
        "segmentation": segmentation
    }
    return unified

def main():
    parser = argparse.ArgumentParser(description="Extract unified insurance metrics from FMP and SEC filings")
    parser.add_argument("ticker", type=str, help="Company ticker symbol")
    parser.add_argument("start_year", type=str, help="Start year (YYYY)")
    parser.add_argument("--email", type=str, required=True, help="Your email for SEC user agent")
    parser.add_argument("--output", type=str, help="Output JSON file")
    args = parser.parse_args()
    
    ticker = args.ticker.upper()
    # Retrieve company profile using the utils function.
    profile = get_company_profile(ticker)
    if not profile:
        logger.error(f"No company profile found for {ticker}")
        return
    
    finder = EDGARExhibit13Finder(f"Insurance Research - Contact: {args.email}")
    extractor = MetricsExtractor(f"Insurance Research - Contact: {args.email}")
    
    try:
        cik = finder.get_cik_from_ticker(ticker)
        logger.info(f"Retrieved CIK for {ticker}: {cik}")
        filings_data = finder.get_company_filings(cik)
        if not filings_data.get("filings"):
            logger.error("No filings found.")
            return
        
        current_year = datetime.datetime.now().year
        start_year_int = int(args.start_year)
        years = list(range(int(args.start_year)-1, current_year + 1))
        # Pass the valid company profile instead of an empty dict.
        fmp_results = extract_yoy_data(ticker, years, segmentation_data={}, profile=profile)
        
        sec_results = {}
        for filing in filings_data["filings"]:
            if filing["filing_date"] < f"{args.start_year}-01-01":
                continue
            filing_docs = finder.get_filing_detail(filing["filing_href"])
            xml_url = filing_docs.get("xml")
            if xml_url:
                logger.info(f"Processing SEC XML filing dated {filing['filing_date']}")
                sec_year_data = extractor.extract_metrics(xml_url)
                for year_str, data in sec_year_data.items():
                    year_int = int(year_str)
                    if year_int not in sec_results:
                        sec_results[year_int] = data
                    else:
                        sec_results[year_int]["profit_desc"].update(data.get("profit_desc", {}))
                        for subsec in ["assets", "liabilities", "shareholders_equity"]:
                            sec_results[year_int]["balance_sheet"].setdefault(subsec, {}).update(
                                data.get("balance_sheet", {}).get(subsec, {}))
                        sec_results[year_int]["segmentation"].update(data.get("segmentation", {}))
            else:
                logger.warning(f"No XML filing document found for filing dated {filing['filing_date']}")
            time.sleep(0.1)
        
        # Only include years for which SEC data exists.
        unified_results = {}
        for year in sec_results.keys():
            fmp_year_data = fmp_results.get(year, {})
            sec_year_data = sec_results[year]
            unified_results[year] = create_unified_year_output(year, fmp_year_data, sec_year_data)
        
         # Now drop any years in unified_results before our desired start year.
        unified_results = {year: data for year, data in unified_results.items() if year >= start_year_int}
        ordered_results = dict(sorted(unified_results.items(), reverse=True))
        output_file = args.output or f"{ticker.lower()}_unified_insurance_metrics.json"
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(ordered_results, f, indent=2)
        logger.info(f"Unified results saved to {output_file}")
    except Exception as e:
        logger.error(f"Error: {e}")
        raise

if __name__ == "__main__":
    main()
