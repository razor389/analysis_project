# analysis_project/gemini_summarizer.py

import os
import json
import asyncio
import re
from datetime import datetime
from typing import List, Dict, Any
import google.generativeai as genai
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configure the Gemini API client
gemini_api_key = os.getenv("GEMINI_API_KEY")
if not gemini_api_key:
    raise ValueError("GEMINI_API_KEY not found in environment. Please set it in your .env file.")
genai.configure(api_key=gemini_api_key)

async def _generate_post_summary_async(posts: List[Dict[str, Any]], ticker: str) -> str:
    """Async helper function to generate a summary using the Gemini API."""
    sorted_posts = sorted(posts, key=lambda x: x.get('timestamp', 0), reverse=True)
    
    content_lines = [
        f"Date: {datetime.fromtimestamp(post['timestamp'])}\nMessage:\n{post['message']}\n\n"
        for post in sorted_posts
    ]
    
    system_prompt = """You are an expert financial analyst. Your task is to summarize the key insights about the given ticker in exactly 10 concise bullet points.

Requirements:
- Each bullet point must start with a bold topic phrase followed by a colon (e.g., **Topic:** Detail).
- Focus on concrete facts, financial metrics, and in-depth analysis.
- Give more weight to the most recent information and any content from 'smgacm@gmail.com'.
- Do not mention the sources, dates, or email addresses in your summary.
- Present all information as direct market insights.
- Number each bullet point from 1 to 10."""

    user_prompt = f"Here are the posts about {ticker} for you to analyze (newest first):\n\n" + "".join(content_lines)

    model = genai.GenerativeModel(
        model_name="gemini-2.5-pro",
        system_instruction=system_prompt
    )
    
    print("[GEMINI] ==> Calling the API...")
    response = await model.generate_content_async([user_prompt])
    print("[GEMINI] ==> Response received.")
    
    summary_text = str(response.text.strip())
    match = re.search(r'1\.\s+\*\*', summary_text)

    if match:
        return summary_text[match.start():]
    else:
        print("[GEMINI] ==> WARNING: '1. **' pattern not found. Returning raw response.")
        return summary_text

def generate_post_summary(posts: list, ticker: str) -> str:
    """
    The main interface function that calls the async helper and returns the result.
    """
    print("[GEMINI] ==> Summarizer initiated.")
    if not posts:
        return "No posts were provided to summarize."
        
    try:
        # Run the async function and explicitly store its result
        summary = asyncio.run(_generate_post_summary_async(posts, ticker))
        
        # Final check to ensure the summary is not empty before returning
        if summary:
            print(f"[GEMINI] ==> Summary generated successfully ({len(summary)} characters). Returning to main script.")
            return summary
        else:
            print("[GEMINI] ==> ERROR: The summarizer returned an empty string.")
            return "Error: Summary was empty."
            
    except Exception as e:
        import traceback
        print(f"[GEMINI] ==> An exception occurred: {e}")
        traceback.print_exc()
        return "Error: Could not generate summary due to an exception."

# --- Standalone Execution Logic (for testing) ---
def process_ticker_posts(ticker: str, debug: bool = False):
    posts_filename = os.path.join("output", f"{ticker}_posts.json")
    emails_filename = os.path.join("output", f"{ticker}_sent_emails.json")
    
    try:
        combined_data = []
        if os.path.exists(posts_filename):
            with open(posts_filename, "r", encoding="utf-8") as f:
                combined_data.extend(json.load(f))
        if os.path.exists(emails_filename):
            with open(emails_filename, "r", encoding="utf-8") as f:
                combined_data.extend(json.load(f))

        if combined_data:
            summary = generate_post_summary(combined_data, ticker)
            output_filename = os.path.join("output", f"{ticker}_post_summary.txt")
            with open(output_filename, "w", encoding="utf-8") as f:
                f.write(summary)
            print(f"Summary saved to {output_filename}")
        else:
            print(f"No data for {ticker}.")
            
    except Exception as e:
        print(f"Error processing data for {ticker}: {e}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='Generate summaries of ticker posts and emails using Gemini.')
    parser.add_argument('ticker', type=str, help='Stock ticker symbol (e.g., AAPL)')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode')
    args = parser.parse_args()
    process_ticker_posts(args.ticker.upper(), args.debug)