import win32com.client
import json
import logging
import sys
from typing import List, Dict, Any
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
        # Use the correct constant for Sent Mail (5 = olFolderSentMail)
        sent_folder = namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
        messages = sent_folder.Items
        messages.Sort("[SentOn]", Descending=True)  # Optional: Sort messages by sent date
        if messages.Count == 0:
            logging.warning("No messages found in Sent Items folder.")
            return []
        logging.info(f"Total messages found in Sent Items: {messages.Count}")
        return messages
    except Exception as e:
        logging.error(f"Error fetching sent emails: {e}")
        return []

def is_valid_ticker(ticker: str) -> bool:
    """Validate ticker format (1-5 uppercase letters)."""
    return bool(re.match(r'^[A-Z]{1,5}$', ticker))

def clean_message(raw_message: str) -> str:
    """
    Cleans the raw email message by removing excessive line breaks,
    email signatures, and other boilerplate text.
    """
    # Remove email signatures (common patterns)
    signature_patterns = [
        r'Scott Granowski CFA®, CFP®\s+Academy Capital Management.*',  # Adjust as needed
        r'Sent via .*',  # Remove lines like "Sent via the Samsung Galaxy..."
        r'-------- Original message --------.*',  # Remove original message blocks
        r'From: .*',  # Remove lines starting with From:
        r'[\r\n]{2,}',  # Replace multiple line breaks with two
    ]
    
    cleaned = raw_message

    for pattern in signature_patterns:
        cleaned = re.sub(pattern, '', cleaned, flags=re.DOTALL | re.IGNORECASE)

    # Replace multiple line breaks with single space
    cleaned = re.sub(r'[\r\n]+', ' ', cleaned)

    # Remove any remaining excessive whitespace
    cleaned = re.sub(r'\s{2,}', ' ', cleaned)

    # Trim leading and trailing whitespace
    cleaned = cleaned.strip()

    return cleaned

def filter_emails(messages, ticker: str) -> List[Dict[str, Any]]:
    """
    Filter emails that contain the ticker in the subject line.

    Args:
        messages: Outlook messages to filter.
        ticker: Ticker symbol to search for.

    Returns:
        A list of dictionaries containing filtered email data.
    """
    filtered_emails = []
    processed_count = 0

    ticker_upper = ticker.upper()
    pattern = r'\b' + re.escape(ticker_upper) + r'\b'

    for message in messages:
        try:
            if message.Class != 43:  # Skip non-mail items
                continue

            subject = str(message.Subject).strip()

            if re.search(pattern, subject.upper()):
                sent_time = message.SentOn.strftime('%Y-%m-%d %H:%M:%S')
                unix_timestamp = email_to_unix(sent_time)
                raw_body = str(message.Body).strip()
                cleaned_body = clean_message(raw_body) 

                logging.info(f"Found {ticker_upper} in email subject: {subject}")
                filtered_emails.append({
                    "timestamp": unix_timestamp,
                    "message": cleaned_body,
                    "authorEmail": "smgacm@gmail.com"  # Use the sender email from environment
                })

            processed_count += 1
            if processed_count % 1000 == 0:
                logging.info(f"Processed {processed_count} messages...")

        except Exception as e:
            logging.warning(f"Failed to process a message: {e}")
            continue

    return filtered_emails

def filter_emails_by_ticker(ticker: str) -> str:
    """
    Main function to filter sent emails by ticker.

    Args:
        ticker: The ticker symbol to search for in email subjects.

    Returns:
        The path to the JSON file containing filtered emails.
    """
    ticker = ticker.upper()

    if not is_valid_ticker(ticker):
        raise ValueError("Ticker must be 1-5 uppercase letters")

    logging.info(f"Searching for ticker: {ticker} in Sent Items")

    namespace = initialize_outlook()
    messages = fetch_sent_emails(namespace)

    if not messages:
        logging.info("No messages to process.")
        return ""

    filtered_emails = filter_emails(messages, ticker)

    if not filtered_emails:
        logging.info(f"No emails found containing '{ticker}' in the subject line.")
        return ""

    output_file = os.path.join('output', f'{ticker}_sent_emails.json')

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(filtered_emails, f, indent=4, ensure_ascii=False)

    logging.info(f"Email filtering complete. Results saved to {output_file}")

    email_count = len(filtered_emails)
    print(f"\nFound {email_count} emails sent by {SENDER_EMAIL} containing '{ticker}' in the subject line.")
    print(f"Results saved to: {output_file}")

    return output_file

def main():
    """Command-line interface for filtering emails by ticker."""
    try:
        if len(sys.argv) != 2:
            print("Usage: python outlook_ticker_search.py TICKER")
            sys.exit(1)

        ticker = sys.argv[1].upper()

        filter_emails_by_ticker(ticker)

    except ValueError as ve:
        logging.error(f"Validation Error: {ve}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
