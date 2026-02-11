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
        self.balance_sheet_rollups = config.get("balance_sheet_rollups") or []
        
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

    def _apply_balance_rollups(self, balance_data: dict) -> None:
        if not self.balance_sheet_rollups:
            return
        for year, metrics in balance_data.items():
            for rule in self.balance_sheet_rollups:
                if not self._year_ok(year, rule):
                    continue
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

        # NEW: treat these as neutral (don’t disqualify consolidation)
        RELATED_PARTY_AXIS = "us-gaap:RelatedPartyTransactionsByRelatedPartyAxis"
        NONRELATED_MEMBER  = "us-gaap:NonrelatedPartyMember"

        # Drop neutral axes: LegalEntity and RelatedParty=Nonrelated
        effective_pairs = []
        for dim, mem in pairs:
            if dim.endswith("LegalEntityAxis"):
                continue
            if dim == RELATED_PARTY_AXIS and mem == NONRELATED_MEMBER:
                continue
            effective_pairs.append((dim, mem))

        # Use the remaining pairs for the original checks
        pairs = effective_pairs

        ALLOWED_CONSOLIDATED_PAIRS = {
            ("srt:ProductOrServiceAxis", "brka:InsuranceAndOtherMember"),
            ("us-gaap:ProductOrServiceAxis", "brka:InsuranceAndOtherMember"),
            ("srt:ProductOrServiceAxis", "brka:RailroadUtilitiesAndEnergyMember"),
            ("us-gaap:ProductOrServiceAxis", "brka:RailroadUtilitiesAndEnergyMember"),
        }
        if any((dim, mem) in ALLOWED_CONSOLIDATED_PAIRS for dim, mem in pairs):
            return True

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
        if any(dim in CONSOL_AXES for dim, _ in pairs) and any(mem in CONSOL_MEMBERS for _, mem in pairs):
            return True

        if effective_pairs == []:
            return True  # only neutral axes were present

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
            Higher is better. Prefer:
            1) Explicit consolidated contexts
            2) Fewer extra business axes beyond what's required
            3) No StatementBusinessSegmentsAxis
            """
            score = 0
            if any(dim in CONSOL_AXES and mem in CONSOL_MEMBERS for dim, mem in pairs):
                score += 10
            extra = [(d, m) for (d, m) in pairs if (d, m) not in required_pairs_set]
            score += max(0, 5 - len(extra))
            if any(dim == "us-gaap:StatementBusinessSegmentsAxis" for dim, _ in pairs):
                score -= 2
            return score

        def _parse_numeric(elem):
            # Robust numeric parse with () negatives and optional @scale / @sign
            raw = (elem.get_text(strip=True) or "").replace(",", "")
            if not raw:
                return None
            if raw.startswith("(") and raw.endswith(")"):
                raw = "-" + raw[1:-1]
            try:
                val = float(raw)
            except Exception:
                return None
            scale = elem.get("scale") or elem.get("Scale") or "0"
            if scale and scale != "0":
                try:
                    val *= 10 ** int(scale)
                except Exception:
                    pass
            sign_attr = (elem.get("sign") or "").lower()
            if sign_attr in {"-", "neg", "negative"}:
                val = -abs(val)
            return val

        for metric_name, spec in mapping.items():
            components = spec if isinstance(spec, list) else [spec]

            # We will ALWAYS accumulate into totals across components.
            # Within each component we "pick one" context by default.
            totals = {}  # year -> float

            # Optional metric-level override (component-level still wins)
            metric_aggregate = None
            if isinstance(spec, dict):
                metric_aggregate = spec.get("aggregate")

            for comp in components:
                if isinstance(comp, str):
                    tag = comp
                    required = None
                    aggregate_mode = metric_aggregate or "pick_one"   # <-- default changed
                    year_filter = None
                else:
                    tag = comp.get("tag")
                    required = comp.get("explicitMembers", None)
                    aggregate_mode = comp.get("aggregate", metric_aggregate or "pick_one")  # <-- default changed
                    year_filter = comp  # we’ll read year_gte/year_lte/years/exclude_years from here
                    if not tag:
                        continue

                # Gather candidates for THIS component, then pick one per year (or sum if explicitly requested)
                comp_candidates = {}  # year -> [(val, pairs, req_set)]
                comp_sums = {}        # year -> float (only used if aggregate_mode == "sum")

                elems = soup.find_all(tag)
                for elem in elems:
                    try:
                        context_ref = elem.get("contextRef") or elem.get("contextref")
                        if not context_ref:
                            continue
                        context = soup.find("context", {"id": context_ref}) or find_context(soup, context_ref)
                        if not context:
                            continue

                        # Keep only consolidated contexts (your existing logic)
                        if not self.is_consolidated_context(context):
                            continue

                        pairs = _pairs_for_context(context)
                        pairs_set = set(pairs)

                        # Enforce required explicit members
                        if required:
                            req_set = {(dim, expected) for dim, expected in required.items()}
                            if not req_set.issubset(pairs_set):
                                continue
                            # Optional: avoid LegalEntityAxis duplicates when required members present
                            if any(dim.endswith("LegalEntityAxis") for dim, _ in pairs):
                                continue
                        else:
                            req_set = set()

                        # Year + year filters
                        year = year_from_context_ref(context_ref)
                        if not year:
                            continue
                        if isinstance(year_filter, dict) and any(k in year_filter for k in ("year_gte","year_lte","years","exclude_years")):
                            if not self._year_ok(year, year_filter):
                                continue

                        val = _parse_numeric(elem)
                        if val is None:
                            continue

                        if aggregate_mode == "sum":
                            comp_sums[year] = comp_sums.get(year, 0.0) + val
                        else:
                            comp_candidates.setdefault(year, []).append((val, pairs, req_set))

                    except Exception as e:
                        logger.error(f"Error processing {metric_name}: {e}")
                        continue

                # Fold THIS component into metric totals
                if aggregate_mode == "sum":
                    for y, v in comp_sums.items():
                        totals[y] = totals.get(y, 0.0) + v
                else:
                    for y, lst in comp_candidates.items():
                        # pick ONE best context for this component/year
                        chosen_val = max(lst, key=lambda t: (_score_preference(t[1], t[2]), abs(t[0])))[0]
                        totals[y] = totals.get(y, 0.0) + chosen_val

            # Emit final metric values
            for y, v in totals.items():
                local.setdefault(y, {})[metric_name] = v

        return local

    def process_segmentation(self, soup):
        seg_results = {}

        BUSINESS_AXES = {
            "us-gaap:StatementBusinessSegmentsAxis",
            "us-gaap:SubsegmentsAxis",
            "srt:ProductOrServiceAxis",
            "us-gaap:ProductOrServiceAxis",  # <-- add this
            "srt:GeographicalAreasAxis",
            "srt:MajorCustomersAxis",
        }
        CONSOL_AXES = {"us-gaap:ConsolidationItemsAxis", "srt:ConsolidationItemsAxis"}
        RELATED_PARTY_AXES = {"us-gaap:RelatedPartyTransactionsByRelatedPartyAxis"}

        # Axes where srt/us-gaap prefixes can vary year-to-year
        AXIS_ALIASES = {
            "srt:ConsolidationItemsAxis": {"srt:ConsolidationItemsAxis", "us-gaap:ConsolidationItemsAxis"},
            "us-gaap:ConsolidationItemsAxis": {"srt:ConsolidationItemsAxis", "us-gaap:ConsolidationItemsAxis"},
            "srt:ProductOrServiceAxis": {"srt:ProductOrServiceAxis", "us-gaap:ProductOrServiceAxis"},
            "us-gaap:ProductOrServiceAxis": {"srt:ProductOrServiceAxis", "us-gaap:ProductOrServiceAxis"},
        }

        def _segment_node(context):
            seg = context.find("segment")
            if seg is not None:
                return seg
            ent = context.find("entity")
            return ent.find("segment") if ent is not None else None

        def _pairs_for_context(context):
            seg = _segment_node(context)
            if not seg:
                return []
            explicit = seg.find_all(lambda t: "explicitmember" in t.name.lower())
            return [(e.get("dimension", "").strip(), e.get_text(strip=True)) for e in explicit]

        def _normalize_pairs(pairs):
            # Remove neutral axes that cause duplicates
            RELATED_PARTY_AXIS = "us-gaap:RelatedPartyTransactionsByRelatedPartyAxis"
            NONRELATED_MEMBER  = "us-gaap:NonrelatedPartyMember"
            out = []
            for dim, mem in pairs:
                if dim.endswith("LegalEntityAxis"):
                    continue
                if dim == RELATED_PARTY_AXIS and mem == NONRELATED_MEMBER:
                    continue
                out.append((dim, mem))
            return out

        def _duration_days(context):
            period = context.find("period")
            if not period:
                return None
            start = period.find("startDate")
            end = period.find("endDate")
            if not (start and end):
                return None
            try:
                from datetime import date
                return (date.fromisoformat(end.text.strip()) -
                        date.fromisoformat(start.text.strip())).days
            except Exception:
                return None

        def _year_from_context_ref(context_ref):
            ctx = self.parse_context(soup, context_ref)
            m = re.search(r"(\d{4})", ctx.get("period", ""))
            return m.group(1) if m else None

        def _required_ok(pairs_set, required):
            # “either/or” for aliased axes
            for dim, mem in (required or {}).items():
                if dim in AXIS_ALIASES:
                    if not any((d, mem) in pairs_set for d in AXIS_ALIASES[dim]):
                        return False
                else:
                    if (dim, mem) not in pairs_set:
                        return False
            return True

        def _strict_accept(pairs, required):
            pairs_set = set(pairs)
            if not _required_ok(pairs_set, required):
                return False

            # Reject intersegment eliminations outright
            if any(mem == "us-gaap:IntersegmentEliminationMember" for _, mem in pairs):
                return False

            # Reject extra axes in these families unless explicitly required (with alias expansion)
            req_dims = set()
            for d in (required or {}).keys():
                req_dims |= AXIS_ALIASES.get(d, {d})

            for dim, _ in pairs:
                if (dim in BUSINESS_AXES or dim in CONSOL_AXES or dim in RELATED_PARTY_AXES) and dim not in req_dims:
                    return False

            return True

        def _score(pairs, days):
            score = 0
            if days is not None and 360 <= days <= 370:
                score += 3
            score += max(0, 5 - len(pairs))
            return score

        for seg_key, seg_info in self.segmentation_mapping.items():
            specs = seg_info if isinstance(seg_info, list) else [seg_info]

            # best_by_year across specs (in case you ever overlap years)
            best_by_year = {}

            for spec in specs:
                tag = spec["tag"]
                required = spec.get("explicitMembers", {}) or {}
                year_filter = spec

                seen = set()  # (contextRef, tag, raw_text)

                for elem in soup.find_all(tag):
                    try:
                        context_ref = elem.get("contextRef") or elem.get("contextref")
                        if not context_ref:
                            continue
                        raw = elem.get_text(strip=True)
                        if not raw:
                            continue
                        sig = (context_ref, tag, raw)
                        if sig in seen:
                            continue
                        seen.add(sig)

                        context = soup.find("context", {"id": context_ref}) or find_context(soup, context_ref)
                        if not context:
                            continue

                        pairs = _normalize_pairs(_pairs_for_context(context))
                        if not _strict_accept(pairs, required):
                            continue

                        year = _year_from_context_ref(context_ref)
                        if not year:
                            continue
                        if isinstance(year_filter, dict) and any(k in year_filter for k in ("year_gte","year_lte","years","exclude_years")):
                            if not self._year_ok(year, year_filter):
                                continue

                        # keep your numeric parse (no scale handling here)
                        val = self._parse_numeric(elem)
                        if val is None:
                            continue

                        days = _duration_days(context)
                        score = _score(pairs, days)

                        prev = best_by_year.get(year)
                        cand = (score, abs(val), val)
                        if prev is None or (cand[0] > prev["rank"][0]) or (cand[0] == prev["rank"][0] and cand[1] > prev["rank"][1]):
                            best_by_year[year] = {"val": val, "rank": cand}

                    except Exception as e:
                        logger.error(f"Error processing segmentation {seg_key}: {e}")
                        continue

            for y, rec in best_by_year.items():
                seg_results.setdefault(y, {})[seg_key] = rec["val"]

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
        self._apply_balance_rollups(balance_data)
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