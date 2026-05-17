import win32com.client
import argparse
import json
import logging
import sys
from typing import List, Dict, Any, Set, Optional, Iterable, Tuple
import re
from datetime import datetime
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
OL_FOLDER_SENT_MAIL = 5
SENT_FOLDER_NAMES = {
    "sent",
    "sent items",
    "sent mail",
    "sent messages",
    "sent e-mail",
}


def safe_getattr(obj, name, default=None):
    """Safely get an attribute from a COM object without raising."""
    try:
        return getattr(obj, name, default)
    except Exception:
        return default


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
        namespace.Logon("", "", False, False)
        return namespace
    except Exception as e:
        logging.error(f"Failed to initialize Outlook: {e}")
        sys.exit(1)


def safe_get_smtp_from_accessor(message, prop_tag: str) -> Optional[str]:
    """Try to get an SMTP address using PropertyAccessor; return None if unavailable."""
    try:
        accessor = safe_getattr(message, "PropertyAccessor", None)
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
        v = safe_getattr(message, "SenderEmailAddress", None)
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
        recipients = safe_getattr(message, "Recipients", None)
        if not recipients:
            return
        for r in recipients:
            try:
                addr = safe_getattr(r, "Address", None)
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


def normalize_folder_name(name: Any) -> str:
    """Normalize an Outlook folder name for matching."""
    return str(name or "").strip().lower()


def is_sent_folder_name(name: Any) -> bool:
    """Return True for common Outlook/Gmail/IMAP sent folder names."""
    normalized = normalize_folder_name(name)
    if normalized in SENT_FOLDER_NAMES:
        return True

    return normalized.startswith(
        (
            "sent items ",
            "sent items (",
            "sent mail ",
            "sent mail (",
            "sent messages ",
            "sent messages (",
            "sent e-mail ",
            "sent e-mail (",
        )
    )


def get_folder_path(folder) -> str:
    """Build a best-effort display path for a COM folder."""
    names: List[str] = []
    current = folder
    seen_ids: Set[int] = set()

    for _ in range(25):
        if current is None:
            break

        object_id = id(current)
        if object_id in seen_ids:
            break
        seen_ids.add(object_id)

        name = safe_getattr(current, "Name", None)
        if name:
            names.append(str(name))

        parent = safe_getattr(current, "Parent", None)
        if parent is None:
            break
        current = parent

    return " / ".join(reversed(names)) if names else "Unknown Sent Folder"


def get_folder_key(folder) -> str:
    """Return a stable key for de-duplicating Outlook folders."""
    entry_id = safe_getattr(folder, "EntryID", None)
    store_id = safe_getattr(folder, "StoreID", None)
    if entry_id:
        return f"{store_id or ''}:{entry_id}"
    return get_folder_path(folder)


def add_sent_folder_source(sources: List[Tuple[str, Any]], seen_folders: Set[str], folder) -> None:
    """Add a sent folder's Items collection to sources once."""
    if folder is None:
        return

    folder_key = get_folder_key(folder)
    if folder_key in seen_folders:
        return
    seen_folders.add(folder_key)

    try:
        items = folder.Items
    except Exception as e:
        logging.warning(f"Unable to read Items for sent folder {get_folder_path(folder)}: {e}")
        return

    try:
        items.Sort("[SentOn]", True)
    except Exception:
        pass

    source_name = get_folder_path(folder)
    count = safe_getattr(items, "Count", None)
    if count is not None:
        source_name = f"{source_name} (Count={count})"

    sources.append((source_name, items))


def iter_child_folders(folder) -> Iterable[Any]:
    """Yield direct child folders from an Outlook folder, best-effort."""
    try:
        folders = safe_getattr(folder, "Folders", None)
        if folders is None:
            return
        for child in folders:
            yield child
    except Exception:
        return


def discover_sent_folders(root_folder) -> Iterable[Any]:
    """Recursively yield folders whose names look like Sent folders."""
    stack = [root_folder]
    seen_paths: Set[str] = set()

    while stack:
        folder = stack.pop()
        path = get_folder_path(folder)
        if path in seen_paths:
            continue
        seen_paths.add(path)

        if is_sent_folder_name(safe_getattr(folder, "Name", "")):
            yield folder

        children = list(iter_child_folders(folder) or [])
        stack.extend(reversed(children))


def build_items_sources(namespace) -> List[Tuple[str, Any]]:
    """
    Return list of (source_name, ItemsCollection) for Sent folders.

    Outlook accounts can expose the real sent folder as a store default, a
    top-level "Sent Items" folder, or a nested IMAP/Gmail folder named
    "Sent Mail". Check all of those and de-duplicate by folder EntryID.
    """
    sources: List[Tuple[str, Any]] = []
    seen_folders: Set[str] = set()

    # Include the profile default first.
    try:
        add_sent_folder_source(
            sources,
            seen_folders,
            namespace.GetDefaultFolder(OL_FOLDER_SENT_MAIL),
        )
    except Exception as e:
        logging.warning(f"Unable to read profile default Sent folder: {e}")

    # Include per-store default Sent folders when Outlook exposes Stores.
    try:
        stores = safe_getattr(namespace, "Stores", None)
        if stores:
            for store in stores:
                try:
                    add_sent_folder_source(
                        sources,
                        seen_folders,
                        store.GetDefaultFolder(OL_FOLDER_SENT_MAIL),
                    )
                except Exception:
                    continue
    except Exception as e:
        logging.warning(f"Failed to enumerate Outlook stores: {e}")

    # Search all mailbox roots for common Sent folder names, including nested
    # IMAP/Gmail layouts such as "[Gmail] / Sent Mail".
    try:
        top_folders = namespace.Folders
        for root_folder in top_folders:
            try:
                for sent_folder in discover_sent_folders(root_folder):
                    add_sent_folder_source(sources, seen_folders, sent_folder)
            except Exception:
                continue
    except Exception as e:
        logging.warning(f"Failed to recursively enumerate Sent folders: {e}")

    if sources:
        logging.info(f"Discovered {len(sources)} Sent folder source(s).")
    else:
        logging.error("No Sent folders found.")

    return sources

def filter_emails(
    items_sources: List[Tuple[str, Any]],
    primary_ticker: str,
    search_terms: Set[str],
    min_year: int = 2018,
    max_emails: Optional[int] = None,
    require_sender_match: bool = True,
) -> List[Dict[str, Any]]:
    """
    Filter emails that contain any of the search terms in the subject line.
    """
    filtered_emails: List[Dict[str, Any]] = []
    processed_count = 0
    seen_emails = set()

    if min_year < 1:
        raise ValueError(f"min_year must be >= 1, got {min_year}")
    if max_emails is not None and max_emails <= 0:
        raise ValueError(f"max_emails must be > 0 when provided, got {max_emails}")

    # Define cutoff as the start of the selected year.
    cutoff_date = datetime(min_year, 1, 1)

    patterns = {
        term: re.compile(r"\b" + re.escape(term) + r"\b", re.IGNORECASE)
        for term in search_terms
    }

    for source_name, items in items_sources:
        logging.info(f"Scanning Source: {source_name} ...")

        for message in items:
            # 1. Class Check (43 = olMailItem)
            if safe_getattr(message, "Class", None) != 43:
                continue

            # 2. Date Check
            sent_time_dt_raw = safe_getattr(message, "SentOn", None)
            if not sent_time_dt_raw:
                continue

            # Convert to naive datetime
            try:
                sent_time_dt = datetime(
                    sent_time_dt_raw.year, sent_time_dt_raw.month, sent_time_dt_raw.day,
                    sent_time_dt_raw.hour, sent_time_dt_raw.minute, sent_time_dt_raw.second
                )
            except Exception:
                continue

            # 3. Cutoff Check
            # Since items are sorted Descending, we can break once we hit old emails
            if sent_time_dt < cutoff_date:
                break

            # 4. Excluded Email Check
            if email_contains_excluded_address(message, EXCLUDED_EMAIL):
                continue

            # 5. Sender Match Check REMOVED
            # We trust "Sent Items" contains only emails sent by the user.
            # This avoids the Exchange X.500 address mismatch issue.

            subject = str(safe_getattr(message, "Subject", "") or "").strip()
            if not subject:
                continue

            # 6. Term Match Check
            found_terms = [
                term for term, pattern in patterns.items() if pattern.search(subject)
            ]
            
            if found_terms:
                logging.info(f"MATCH FOUND: '{subject}' with terms {found_terms}")
                
                unix_timestamp = int(sent_time_dt.timestamp())
                email_id = f"{unix_timestamp}_{subject}"

                if email_id in seen_emails:
                    continue
                seen_emails.add(email_id)

                filtered_emails.append({
                    "timestamp": unix_timestamp,
                    "message": clean_message(str(safe_getattr(message, "Body", "") or "")),
                    "authorEmail": SENDER_EMAIL,
                    "sourceFolder": source_name,
                    "subject": subject,
                })

            processed_count += 1
            if processed_count % 1000 == 0:
                logging.info(f"Processed {processed_count} messages...")

    filtered_emails.sort(key=lambda email: email.get("timestamp", 0), reverse=True)

    if max_emails is not None:
        filtered_emails = filtered_emails[:max_emails]

    return filtered_emails

def filter_emails_by_config(
    ticker: str,
    config_path: str = "ticker_email_config.json",
    min_year: int = 2018,
    max_emails: Optional[int] = None,
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
        min_year=min_year,
        max_emails=max_emails,
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
    print(f"Included emails from {min_year} or later.")
    if max_emails is not None:
        print(f"Applied max email cap of {max_emails}, keeping the newest matches.")
    print(f"Results saved to: {output_file}")

    return output_file


def main():
    """Command-line interface for filtering emails by ticker."""
    try:
        parser = argparse.ArgumentParser(description="Filter Outlook sent emails by ticker")
        parser.add_argument("ticker", type=str, help="Ticker symbol")
        parser.add_argument(
            "--min_year",
            type=int,
            default=2018,
            help="Only include emails sent in this year or later (default: 2018)",
        )
        parser.add_argument(
            "--max_emails",
            type=int,
            default=None,
            help="Optional cap on matched emails; keeps the newest emails and drops older overflow",
        )
        args = parser.parse_args()

        ticker = args.ticker.upper()
        filter_emails_by_config(ticker, min_year=args.min_year, max_emails=args.max_emails)

    except ValueError as ve:
        logging.error(f"Validation Error: {ve}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
