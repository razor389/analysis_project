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
        self.profit_rollups = (config.get("profit_rollups") or [])
        self.suppress_profit_keys = set(config.get("suppress_profit_keys") or [])
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

    def _year_ok(self, ystr: str, rule: dict) -> bool:
        try:
            y = int(ystr)
        except Exception:
            return True  # if we can't parse year, don't block
        if "year_gte" in rule and y < int(rule["year_gte"]): return False
        if "year_lte" in rule and y > int(rule["year_lte"]): return False
        if "years" in rule and isinstance(rule["years"], (list, tuple)) and y not in rule["years"]: return False
        if "exclude_years" in rule and isinstance(rule["exclude_years"], (list, tuple)) and y in rule["exclude_years"]: return False
        return True
    
    NUMERIC_NEG_PAT = re.compile(r"^\(([\d,\.]+)\)$")

    def _parse_numeric(self, elem):
        raw = elem.get_text(strip=True)
        if not raw:
            return None
        m = self.NUMERIC_NEG_PAT.match(raw)
        if m:
            raw = "-" + m.group(1)
        raw = raw.replace(",", "")
        try:
            return float(raw)
        except Exception:
            return None
        
    def _debug_rollup(self, year, metrics, rule):
        target = rule.get("target")
        adds = rule.get("add", []) or []
        subs = rule.get("subtract", []) or []
        def gv(k): 
            v = metrics.get(k)
            return v if isinstance(v, (int, float)) else None
        add_parts = [(k, gv(k)) for k in adds]
        sub_parts = [(k, gv(k)) for k in subs]
        logger.info(f"[rollup dbg] y={year} target={target} "
                    f"base={gv(target)} "
                    f"adds={add_parts} subs={sub_parts}")

        
    def _apply_profit_rollups(self, profit_data: dict) -> None:
        """
        Apply per-year rollups like:
        - add/subtract intermediate keys into a target metric
        - optional year filters: year_gte, year_lte, years (list), exclude_years (list)
        Mutates profit_data in place.
        """
        if not self.profit_rollups:
            return

        for year, metrics in profit_data.items():
            for rule in self.profit_rollups:
                if not self._year_ok(year, rule):
                    continue
                # DEBUG (remove after verifying):
                # self._debug_rollup(year, metrics, rule)
                target = rule.get("target")
                adds = rule.get("add", []) or []
                subs = rule.get("subtract", []) or []

                delta = 0.0
                for k in adds:
                    v = metrics.get(k)
                    if isinstance(v, (int, float)):
                        delta += v
                for k in subs:
                    v = metrics.get(k)
                    if isinstance(v, (int, float)):
                        delta -= v

                if target:
                    metrics[target] = (metrics.get(target, 0.0) or 0.0) + delta

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
    
    def is_consolidated_context(self, context) -> bool:
        seg = context.find("segment")
        if not seg:
            return True

        explicit = seg.find_all(lambda t: "explicitmember" in t.name.lower())
        if not explicit:
            return True

        pairs = [(e.get("dimension", "").strip(), e.get_text(strip=True)) for e in explicit]

        # --- BRK consolidated buckets on ProductOrServiceAxis ---
        ALLOWED_CONSOLIDATED_PAIRS = {
            ("srt:ProductOrServiceAxis", "brka:InsuranceAndOtherMember"),
            ("us-gaap:ProductOrServiceAxis", "brka:InsuranceAndOtherMember"),
            ("srt:ProductOrServiceAxis", "brka:RailroadUtilitiesAndEnergyMember"),
            ("us-gaap:ProductOrServiceAxis", "brka:RailroadUtilitiesAndEnergyMember"),
        }
        if any((dim, mem) in ALLOWED_CONSOLIDATED_PAIRS for dim, mem in pairs):
            return True
        # --------------------------------------------------------

        BUSINESS_AXES = {
            "us-gaap:StatementBusinessSegmentsAxis",
            "us-gaap:SubsegmentsAxis",
            "srt:ProductOrServiceAxis",
            "srt:GeographicalAreasAxis",
            "srt:MajorCustomersAxis",
        }
        if any(dim in BUSINESS_AXES for dim, _ in pairs):
            return False

        CONSOL_AXES = {"us-gaap:ConsolidationItemsAxis", "srt:ConsolidationItemsAxis"}
        CONSOL_MEMBERS = {
            "us-gaap:ConsolidatedEntitiesMember",
            "srt:ConsolidatedEntitiesMember",
            "us-gaap:ConsolidatedGroupMember",
        }
        has_consol_axis = any(dim in CONSOL_AXES for dim, _ in pairs)
        if has_consol_axis and any(mem in CONSOL_MEMBERS for _, mem in pairs):
            return True

        if all(dim.endswith("LegalEntityAxis") for dim, _ in pairs):
            return True

        return False

    def process_mapping(self, soup, mapping):
        local = {}

        def year_from_context_ref(context_ref):
            context_data = self.parse_context(soup, context_ref)
            period_text = context_data.get("period", "")
            m = re.search(r"(\d{4})", period_text)
            return m.group(1) if m else None

        CONSOL_AXES = {"us-gaap:ConsolidationItemsAxis", "srt:ConsolidationItemsAxis"}
        CONSOL_MEMBERS = {
            "us-gaap:ConsolidatedEntitiesMember",
            "srt:ConsolidatedEntitiesMember",
            "us-gaap:ConsolidatedGroupMember",
        }
        BUSINESS_AXES = {
            "us-gaap:StatementBusinessSegmentsAxis",
            "us-gaap:SubsegmentsAxis",
            "srt:ProductOrServiceAxis",
            "srt:GeographicalAreasAxis",
            "srt:MajorCustomersAxis",
        }

        def _segment_node(context):
            seg = context.find("segment")
            if seg is not None:
                return seg
            entity = context.find("entity")
            return (entity.find("segment") if entity is not None else None)

        def _pairs_for_context(context):
            seg = _segment_node(context)
            if not seg:
                return []
            explicit = seg.find_all(lambda t: "explicitmember" in t.name.lower())
            return [(e.get("dimension", "").strip(), e.get_text(strip=True)) for e in explicit]

        def _score_preference(pairs, required_pairs_set):
            """
            Higher is better. We prefer:
            1) Explicit consolidated contexts
            2) Fewer extra business axes beyond what's required
            3) No StatementBusinessSegmentsAxis
            """
            score = 0
            if any(dim in CONSOL_AXES and mem in CONSOL_MEMBERS for dim, mem in pairs):
                score += 10
            # count extra axes beyond required
            extra = [(d, m) for (d, m) in pairs if (d, m) not in required_pairs_set]
            # prefer fewer extras
            score += max(0, 5 - len(extra))
            # penalize SBS axis
            if any(dim == "us-gaap:StatementBusinessSegmentsAxis" for dim, _ in pairs):
                score -= 2
            return score

        for metric_name, spec in mapping.items():
            components = spec if isinstance(spec, list) else [spec]

            # If the metric wrapper itself specifies aggregate, capture it;
            # per-component setting overrides metric-level.
            metric_aggregate = None
            if isinstance(spec, dict):
                metric_aggregate = spec.get("aggregate")

            # For aggregation:
            # - when we "sum", we keep totals[year]
            # - when we "pick_one", we keep candidates[year] = list[(val, pairs, required_pairs_set)]
            totals = {}
            candidates = {}

            for comp in components:
                if isinstance(comp, str):
                    tag = comp
                    required = None
                    aggregate_mode = metric_aggregate or "sum"
                else:
                    tag = comp.get("tag")
                    required = comp.get("explicitMembers", None)
                    aggregate_mode = comp.get("aggregate", metric_aggregate or "sum")
                    if not tag:
                        continue

                elems = soup.find_all(tag)
                for elem in elems:
                    try:
                        context_ref = elem.get("contextRef") or elem.get("contextref")
                        if not context_ref:
                            continue
                        context = soup.find("context", {"id": context_ref}) or find_context(soup, context_ref)
                        if not context:
                            continue

                        # Keep only consolidated contexts using your existing logic
                        if not self.is_consolidated_context(context):
                            continue

                        pairs = _pairs_for_context(context)
                        pairs_set = set(pairs)

                        # Enforce required explicit members (if any). We must look under context.segment or entity.segment.
                        if required:
                            req_set = {(dim, expected) for dim, expected in required.items()}
                            if not req_set.issubset(pairs_set):
                                continue
                            # (Optional robustness) Reject LegalEntityAxis to avoid per-subsidiary duplicates
                            if any(dim.endswith("LegalEntityAxis") for dim, _ in pairs):
                                continue
                        else:
                            req_set = set()

                        # Determine year
                        year = year_from_context_ref(context_ref)
                        if not year:
                            continue

                        # Respect any per-component year filters
                        if isinstance(comp, dict) and any(k in comp for k in ("year_gte","year_lte","years","exclude_years")):
                            if not self._year_ok(year, comp):
                                continue

                        # Parse number (no scale multiplication; Inline XBRL already encodes plain numbers textually)
                        raw = elem.get_text(strip=True)
                        if not raw:
                            continue
                        if raw.startswith("(") and raw.endswith(")"):
                            raw = "-" + raw[1:-1]
                        raw = raw.replace(",", "")
                        try:
                            val = float(raw)
                        except Exception:
                            continue

                        if aggregate_mode == "pick_one":
                            # store as candidate; we'll choose best per year
                            candidates.setdefault(year, []).append((val, pairs, req_set))
                        else:
                            # default behavior: sum
                            totals[year] = totals.get(year, 0.0) + val

                    except Exception as e:
                        logger.error(f"Error processing {metric_name}: {e}")
                        continue

            # finalize: write chosen values for pick_one
            for y, lst in candidates.items():
                if not lst:
                    continue
                # score and pick the best candidate
                chosen_val = max(lst, key=lambda t: (_score_preference(t[1], t[2]), abs(t[0])))[0]
                local.setdefault(y, {})[metric_name] = chosen_val

            # finalize: write sums
            for y, v in totals.items():
                local.setdefault(y, {})[metric_name] = v

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
        self._apply_profit_rollups(profit_data)
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
            if year in profit_data:
                # Optionally suppress internal keys
                clean = {
                    k: v for k, v in profit_data[year].items()
                    if k not in self.suppress_profit_keys
                }
                results[year]["profit_desc"] = clean
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