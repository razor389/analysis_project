import win32com.client
import json
import logging
import sys
from typing import List, Dict, Any, Set, Optional, Iterable, Tuple
import re
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv

# Create output directory if it doesn't exist
os.makedirs("output", exist_ok=True)

# Configure logging for console output only
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s:%(message)s",
)

# Load environment variables from .env file
load_dotenv()

# Get sender email from environment variables
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
EXCLUDED_EMAIL = "derekr@academycapitalmgmt.com".lower()

if not SENDER_EMAIL:
    logging.error("SENDER_EMAIL not found in .env file")
    sys.exit(1)

SENDER_EMAIL = SENDER_EMAIL.strip().lower()

# MAPI property tags for SMTP addresses (works better than SenderEmailAddress for Exchange)
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
PR_RECEIVED_BY_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D07001E"
PR_SENT_REPRESENTING_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D02001E"

def to_naive(dt: datetime) -> datetime:
    """Return a tz-naive datetime (drop tzinfo if present)."""
    if isinstance(dt, datetime) and dt.tzinfo is not None:
        return dt.replace(tzinfo=None)
    return dt

def load_ticker_config(config_path: str = "ticker_email_config.json") -> Dict[str, List[str]]:
    """
    Load ticker configuration from JSON file.
    Returns empty dict if file not found or invalid.
    """
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        if not isinstance(config, dict):
            logging.warning(f"Config file is not a dict: {config_path}")
            return {}
        # Normalize config values to list[str]
        normalized: Dict[str, List[str]] = {}
        for k, v in config.items():
            if isinstance(v, list):
                normalized[str(k).upper()] = [str(x) for x in v if str(x).strip()]
            else:
                normalized[str(k).upper()] = [str(v)]
        return normalized
    except FileNotFoundError:
        logging.warning(f"Config file not found: {config_path}")
        return {}
    except json.JSONDecodeError:
        logging.warning(f"Config file invalid JSON: {config_path}")
        return {}


def email_to_unix(email_dt: datetime) -> int:
    """
    Convert a datetime to a Unix timestamp (seconds since epoch).
    Note: Outlook COM usually returns naive datetimes in local time.
    We convert as-is using .timestamp(), which assumes local time for naive dt.
    """
    return int(email_dt.timestamp())


def initialize_outlook():
    """Initialize and return the Outlook namespace."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        # Use existing profile/session; avoid prompting if possible.
        # If you want to force prompt, set Prompt=True.
        namespace.Logon("", "", False, False)
        return namespace
    except Exception as e:
        logging.error(f"Failed to initialize Outlook: {e}")
        sys.exit(1)


def safe_get_smtp_from_accessor(message, prop_tag: str) -> Optional[str]:
    """Try to get an SMTP address using PropertyAccessor; return None if unavailable."""
    try:
        accessor = getattr(message, "PropertyAccessor", None)
        if accessor is None:
            return None
        val = accessor.GetProperty(prop_tag)
        if val:
            return str(val).strip().lower()
    except Exception:
        return None
    return None


def safe_get_sender_smtp(message) -> Optional[str]:
    """
    Get sender SMTP address robustly (Exchange often has non-SMTP SenderEmailAddress).
    """
    # Try common SMTP props
    for tag in (PR_SENDER_SMTP_ADDRESS, PR_SENT_REPRESENTING_SMTP_ADDRESS):
        smtp = safe_get_smtp_from_accessor(message, tag)
        if smtp:
            return smtp

    # Fallback
    try:
        v = getattr(message, "SenderEmailAddress", None)
        if v:
            return str(v).strip().lower()
    except Exception:
        pass
    return None


def safe_iter_recipients_addresses(message) -> Iterable[str]:
    """
    Yield recipient addresses. On Exchange, Recipients[i].Address may not be SMTP,
    but it's still useful as a fallback. This function is best-effort and never throws.
    """
    try:
        recipients = getattr(message, "Recipients", None)
        if not recipients:
            return
        # Recipients is COM collection
        for r in recipients:
            try:
                addr = getattr(r, "Address", None)
                if addr:
                    yield str(addr).strip().lower()
            except Exception:
                continue
    except Exception:
        return


def email_contains_excluded_address(message, excluded_email: str) -> bool:
    """
    Returns True if the excluded email appears as sender or recipient
    (To, CC, or BCC) in the message.

    IMPORTANT: This is best-effort. If anything goes wrong, we FAIL OPEN (return False),
    so we don't accidentally exclude everything.
    """
    try:
        fields: List[str] = []

        sender_smtp = safe_get_sender_smtp(message)
        if sender_smtp:
            fields.append(sender_smtp)

        for addr in safe_iter_recipients_addresses(message):
            fields.append(addr)

        combined = " ".join(fields).lower()
        return excluded_email in combined
    except Exception:
        # Fail open
        return False


def is_valid_search_term(term: str) -> bool:
    """
    Validate search term format.
    Allow either ticker format (1-5 uppercase letters) or company names (word characters and spaces)
    """
    return bool(re.match(r"^[A-Z]{1,5}$", term) or re.match(r"^[\w\s-]+$", term))


def clean_message(raw_message: str) -> str:
    """
    Cleans the raw email message by removing excessive line breaks,
    email signatures, and other boilerplate text.
    """
    signature_patterns = [
        r"Scott Granowski CFA®, CFP®\s+Academy Capital Management.*",
        r"Sent via .*",
        r"-------- Original message --------.*",
        r"From: .*",
        r"[\r\n]{2,}",
    ]

    cleaned = raw_message or ""

    for pattern in signature_patterns:
        cleaned = re.sub(pattern, "", cleaned, flags=re.DOTALL | re.IGNORECASE)

    # Replace multiple line breaks with single space
    cleaned = re.sub(r"[\r\n]+", " ", cleaned)

    # Remove any remaining excessive whitespace
    cleaned = re.sub(r"\s{2,}", " ", cleaned)

    return cleaned.strip()


def build_items_sources(namespace) -> List[Tuple[str, Any]]:
    """
    Return list of (source_name, ItemsCollection) for Sent Items.
    Recommended change: search Sent Items across ALL stores, so multi-account/shared mailbox works.
    """
    sources: List[Tuple[str, Any]] = []

    # Try all stores
    try:
        top_folders = namespace.Folders
        for store in top_folders:
            try:
                store_name = str(getattr(store, "Name", "UnknownStore"))
                sent_folder = None

                # Most common English name
                try:
                    sent_folder = store.Folders["Sent Items"]
                except Exception:
                    # Fallback to default folder for this store isn't directly exposed here;
                    # skip if not found.
                    sent_folder = None

                if sent_folder is None:
                    continue

                items = sent_folder.Items
                # Sort descending by SentOn
                try:
                    items.Sort("[SentOn]", True)
                except Exception:
                    pass

                count = getattr(items, "Count", None)
                if count is None:
                    sources.append((f"{store_name} / Sent Items", items))
                else:
                    sources.append((f"{store_name} / Sent Items (Count={count})", items))
            except Exception:
                continue
    except Exception as e:
        logging.warning(f"Failed to enumerate stores. Falling back to default Sent Items. Error: {e}")

    # Fallback: default Sent Items
    if not sources:
        try:
            sent_folder = namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
            items = sent_folder.Items
            try:
                items.Sort("[SentOn]", True)
            except Exception:
                pass
            sources.append(("Default Store / Sent Items", items))
        except Exception as e:
            logging.error(f"Error fetching default Sent Items: {e}")

    return sources


def filter_emails(
    items_sources: List[Tuple[str, Any]],
    primary_ticker: str,
    search_terms: Set[str],
    lookback_years: int = 15,
    require_sender_match: bool = True,
) -> List[Dict[str, Any]]:
    """
    Filter emails that contain any of the search terms in the subject line.

    Recommended changes applied:
    - Fix timezone-aware vs naive datetime bug by using naive cutoff_date.
    - Fail-open exclusion filter.
    - Robust SMTP extraction for sender (Exchange safe).
    - Optionally require sender match with SENDER_EMAIL.
    - Search across all stores' Sent Items.
    """
    filtered_emails: List[Dict[str, Any]] = []
    processed_count = 0
    seen_emails = set()

    # IMPORTANT FIX: use naive datetime to match Outlook COM naive datetimes
    cutoff_date = to_naive(datetime.now() - timedelta(days=lookback_years * 365))

    patterns = {
        term: re.compile(r"\b" + re.escape(term) + r"\b", re.IGNORECASE)
        for term in search_terms
    }

    for source_name, items in items_sources:
        try:
            total = getattr(items, "Count", None)
            logging.info(f"Scanning {source_name} ...")
            if total is not None:
                logging.info(f"Total messages in {source_name}: {total}")
        except Exception:
            logging.info(f"Scanning {source_name} ... (Count unavailable)")

        for message in items:
            try:
                # 43 = olMailItem
                if getattr(message, "Class", None) != 43:
                    continue

                # Exclusion check (fail-open)
                if email_contains_excluded_address(message, EXCLUDED_EMAIL):
                    continue

                sent_time_dt = getattr(message, "SentOn", None)
                if not isinstance(sent_time_dt, datetime):
                    continue

                sent_time_dt = to_naive(sent_time_dt)

                # Cutoff
                if sent_time_dt < cutoff_date:
                    # Items are sorted descending; we can break early for this source
                    break

                # Require sender match (optional, but recommended)
                if require_sender_match:
                    sender_smtp = safe_get_sender_smtp(message)
                    # If we can't determine sender SMTP, skip (conservative)
                    if not sender_smtp or sender_smtp != SENDER_EMAIL:
                        continue

                subject = str(getattr(message, "Subject", "") or "").strip()
                if not subject:
                    continue

                if any(pattern.search(subject) for pattern in patterns.values()):
                    unix_timestamp = email_to_unix(sent_time_dt)
                    email_id = f"{unix_timestamp}_{subject}"

                    if email_id in seen_emails:
                        continue
                    seen_emails.add(email_id)

                    raw_body = str(getattr(message, "Body", "") or "").strip()
                    cleaned_body = clean_message(raw_body)

                    found_terms = [
                        term for term, pattern in patterns.items() if pattern.search(subject)
                    ]
                    logging.info(f"[{source_name}] Found terms {found_terms} in subject: {subject}")

                    filtered_emails.append(
                        {
                            "timestamp": unix_timestamp,
                            "message": cleaned_body,
                            "authorEmail": SENDER_EMAIL,
                            "sourceFolder": source_name,
                            "subject": subject,
                        }
                    )

                processed_count += 1
                if processed_count % 1000 == 0:
                    logging.info(f"Processed {processed_count} messages...")

            except Exception as e:
                # Recommended change: log the exception with basic context
                subj = None
                try:
                    subj = getattr(message, "Subject", None)
                except Exception:
                    subj = None
                logging.warning(f"Failed to process message (Subject={subj!r}): {e}")
                continue

    return filtered_emails


def filter_emails_by_config(
    ticker: str,
    config_path: str = "ticker_email_config.json",
    lookback_years: int = 15,
) -> str:
    """
    Main function to filter sent emails by ticker and its related terms from config.
    If ticker not in config, searches for just the ticker symbol.
    """
    ticker = ticker.upper()

    # Validate ticker format
    if not is_valid_search_term(ticker):
        raise ValueError(f"Invalid ticker format: {ticker}")

    # Load config and get search terms
    config = load_ticker_config(config_path)

    # Create set of search terms - if ticker not in config, just use the ticker
    search_terms = set([ticker] + config.get(ticker, []))

    logging.info(f"Searching for terms: {search_terms} in Sent Items")

    namespace = initialize_outlook()
    items_sources = build_items_sources(namespace)

    if not items_sources:
        logging.info("No Sent Items sources found to scan.")
        return ""

    filtered_emails = filter_emails(
        items_sources=items_sources,
        primary_ticker=ticker,
        search_terms=search_terms,
        lookback_years=lookback_years,
        require_sender_match=True,  # recommended
    )

    if not filtered_emails:
        logging.info(f"No emails found containing any search terms for {ticker}")
        return ""

    output_file = os.path.join("output", f"{ticker}_sent_emails.json")

    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(filtered_emails, f, indent=4, ensure_ascii=False)

    logging.info(f"Email filtering complete. Results saved to {output_file}")

    email_count = len(filtered_emails)
    print(f"\nFound {email_count} sent emails from {SENDER_EMAIL} containing search terms for '{ticker}'")
    print(f"Results saved to: {output_file}")

    return output_file


def main():
    """Command-line interface for filtering emails by ticker."""
    try:
        if len(sys.argv) != 2:
            print("Usage: python outlook_ticker_search.py TICKER")
            sys.exit(1)

        ticker = sys.argv[1].upper()
        filter_emails_by_config(ticker)

    except ValueError as ve:
        logging.error(f"Validation Error: {ve}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)
