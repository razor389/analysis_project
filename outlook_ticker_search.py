import win32com.client
import json
import logging
import sys
from typing import List, Dict, Any, Set
import re
from datetime import datetime
import os
from dotenv import load_dotenv

# Create output directory if it doesn't exist
os.makedirs('output', exist_ok=True)

# Configure logging for console output only
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s:%(message)s'
)

# Load environment variables from .env file
load_dotenv()

# Get sender email from environment variables
SENDER_EMAIL = os.getenv('SENDER_EMAIL')
if not SENDER_EMAIL:
    logging.error("SENDER_EMAIL not found in .env file")
    sys.exit(1)

def load_ticker_config(config_path: str = 'ticker_email_config.json') -> Dict[str, List[str]]:
    """
    Load ticker configuration from JSON file.
    Returns empty dict if file not found or invalid.
    """
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
        return config
    except (FileNotFoundError, json.JSONDecodeError):
        logging.warning(f"Config file not found or invalid: {config_path}")
        return {}

def email_to_unix(email_timestamp):
    """
    Convert an email timestamp string to a Unix timestamp.
    Args:
        email_timestamp (str): Timestamp in '%Y-%m-%d %H:%M:%S' format.
    Returns:
        int: Unix timestamp (seconds since epoch).
    """
    dt = datetime.strptime(email_timestamp, '%Y-%m-%d %H:%M:%S')
    return int(dt.timestamp())

def initialize_outlook():
    """Initialize and return the Outlook namespace."""
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        namespace.Logon()  # Ensure that Outlook is logged on
        return namespace
    except Exception as e:
        logging.error(f"Failed to initialize Outlook: {e}")
        sys.exit(1)

def fetch_sent_emails(namespace):
    """Fetch and return sent emails from Outlook."""
    try:
        sent_folder = namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
        messages = sent_folder.Items
        messages.Sort("[SentOn]", Descending=True)
        if messages.Count == 0:
            logging.warning("No messages found in Sent Items folder.")
            return []
        logging.info(f"Total messages found in Sent Items: {messages.Count}")
        return messages
    except Exception as e:
        logging.error(f"Error fetching sent emails: {e}")
        return []

def is_valid_search_term(term: str) -> bool:
    """
    Validate search term format.
    Allow either ticker format (1-5 uppercase letters) or company names (word characters and spaces)
    """
    return bool(re.match(r'^[A-Z]{1,5}$', term) or re.match(r'^[\w\s-]+$', term))

def clean_message(raw_message: str) -> str:
    """
    Cleans the raw email message by removing excessive line breaks,
    email signatures, and other boilerplate text.
    """
    signature_patterns = [
        r'Scott Granowski CFA®, CFP®\s+Academy Capital Management.*',
        r'Sent via .*',
        r'-------- Original message --------.*',
        r'From: .*',
        r'[\r\n]{2,}',
    ]
    
    cleaned = raw_message

    for pattern in signature_patterns:
        cleaned = re.sub(pattern, '', cleaned, flags=re.DOTALL | re.IGNORECASE)

    # Replace multiple line breaks with single space
    cleaned = re.sub(r'[\r\n]+', ' ', cleaned)

    # Remove any remaining excessive whitespace
    cleaned = re.sub(r'\s{2,}', ' ', cleaned)

    return cleaned.strip()

def filter_emails(messages, primary_ticker: str, search_terms: Set[str]) -> List[Dict[str, Any]]:
    """
    Filter emails that contain any of the search terms in the subject line.
    """
    filtered_emails = []
    processed_count = 0
    seen_emails = set()

    patterns = {term: re.compile(r'\b' + re.escape(term) + r'\b', re.IGNORECASE) 
               for term in search_terms}

    for message in messages:
        try:
            if message.Class != 43:  # Skip non-mail items
                continue

            subject = str(message.Subject).strip()
            
            if any(pattern.search(subject) for pattern in patterns.values()):
                sent_time = message.SentOn.strftime('%Y-%m-%d %H:%M:%S')
                unix_timestamp = email_to_unix(sent_time)
                email_id = f"{unix_timestamp}_{subject}"
                
                if email_id not in seen_emails:
                    seen_emails.add(email_id)
                    raw_body = str(message.Body).strip()
                    cleaned_body = clean_message(raw_body)

                    found_terms = [term for term, pattern in patterns.items() 
                                 if pattern.search(subject)]
                    logging.info(f"Found terms {found_terms} in email subject: {subject}")
                    
                    filtered_emails.append({
                        "timestamp": unix_timestamp,
                        "message": cleaned_body,
                        "authorEmail": SENDER_EMAIL
                    })

            processed_count += 1
            if processed_count % 1000 == 0:
                logging.info(f"Processed {processed_count} messages...")

        except Exception as e:
            logging.warning(f"Failed to process a message: {e}")
            continue

    return filtered_emails

def filter_emails_by_config(ticker: str, config_path: str = 'ticker_email_config.json') -> str:
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
    messages = fetch_sent_emails(namespace)

    if not messages:
        logging.info("No messages to process.")
        return ""

    filtered_emails = filter_emails(messages, ticker, search_terms)

    if not filtered_emails:
        logging.info(f"No emails found containing any search terms for {ticker}")
        return ""

    output_file = os.path.join('output', f'{ticker}_sent_emails.json')

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(filtered_emails, f, indent=4, ensure_ascii=False)

    logging.info(f"Email filtering complete. Results saved to {output_file}")

    email_count = len(filtered_emails)
    print(f"\nFound {email_count} emails sent by {SENDER_EMAIL} containing search terms for '{ticker}'")
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

if __name__ == "__main__":
    main()