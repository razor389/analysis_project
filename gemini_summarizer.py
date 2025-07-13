# analysis_project/gemini_summarizer.py

import os
import json
import asyncio
import re
from datetime import datetime
from typing import List, Dict, Any
import google.generativeai as genai
from dotenv import load_dotenv

# --- Start of Diagnostic Version ---

def run_gemini_summarizer(posts: List[Dict[str, Any]], ticker: str) -> str:
    """
    Diagnostic wrapper for the async Gemini summarizer.
    """
    print("[GEMINI_DIAGNOSTIC] ==> Entering summarizer.")
    
    # 1. Check for API Key
    gemini_api_key = os.getenv("GEMINI_API_KEY")
    if not gemini_api_key:
        print("[GEMINI_DIAGNOSTIC] ==> ERROR: GEMINI_API_KEY is not set in the .env file.")
        return "Error: GEMINI_API_KEY not found."
    
    # Mask key for printing
    masked_key = f"{gemini_api_key[:4]}...{gemini_api_key[-4:]}"
    print(f"[GEMINI_DIAGNOSTIC] ==> API key loaded successfully: {masked_key}")

    try:
        # 2. Configure the library
        print("[GEMINI_DIAGNOSTIC] ==> Configuring the google.generativeai library.")
        genai.configure(api_key=gemini_api_key)
        print("[GEMINI_DIAGNOSTIC] ==> Library configured.")
        
        # 3. Prepare data for the prompt
        if not posts:
            print("[GEMINI_DIAGNOSTIC] ==> No posts provided to summarize.")
            return "No posts were provided to summarize."
            
        sorted_posts = sorted(posts, key=lambda x: x.get('timestamp', 0), reverse=True)
        content_lines = [
            f"Date: {datetime.fromtimestamp(post['timestamp'])}\nMessage:\n{post['message']}\n\n"
            for post in sorted_posts
        ]
        user_prompt = f"Here are the posts about {ticker} for you to analyze (newest first):\n\n" + "".join(content_lines)
        
        print(f"[GEMINI_DIAGNOSTIC] ==> Prepared {len(sorted_posts)} posts for summarization ({len(user_prompt)} characters).")
        
        # 4. Run the async API call
        print("[GEMINI_DIAGNOSTIC] ==> Starting asyncio event loop to call the API.")
        summary = asyncio.run(_generate_post_summary_async(user_prompt, ticker))
        print("[GEMINI_DIAGNOSTIC] ==> Successfully received summary from API.")
        return summary
        
    except Exception as e:
        # 5. Catch and print any exception with full details
        import traceback
        print("[GEMINI_DIAGNOSTIC] ==> An unhandled exception occurred!")
        print(f"[GEMINI_DIAGNOSTIC] ==> Exception Type: {type(e).__name__}")
        print(f"[GEMINI_DIAGNOSTIC] ==> Exception Details: {e}")
        traceback.print_exc()
        return f"Error generating summary: {e}"

async def _generate_post_summary_async(user_prompt: str, ticker: str) -> str:
    """
    Async helper function to generate a summary using the Gemini API.
    """
    system_prompt = """You are an expert financial analyst. Your task is to summarize the key insights about the given ticker in exactly 10 concise bullet points.

Requirements:
- Each bullet point must start with a bold topic phrase followed by a colon (e.g., **Topic:** Detail).
- Focus on concrete facts, financial metrics, and in-depth analysis.
- Give more weight to the most recent information and any content from 'smgacm@gmail.com'.
- Do not mention the sources, dates, or email addresses in your summary.
- Present all information as direct market insights.
- Number each bullet point from 1 to 10."""

    generation_config = {
        "temperature": 0.3,
        "max_output_tokens": 2000,
    }

    model = genai.GenerativeModel(
        model_name="gemini-1.5-pro",
        generation_config=generation_config,
        system_instruction=system_prompt
    )
    
    print("[GEMINI_DIAGNOSTIC] ==> Inside async function, attempting to call model.generate_content_async()...")
    response = await model.generate_content_async([user_prompt])
    summary_text = response.text.strip()
    
    match = re.search(r'1\.\s+\*\*', summary_text)
    if match:
        summary_text = summary_text[match.start():]
        
    return summary_text

# --- Main Interface Function ---
# This is the function that acm_analysis.py calls.
def generate_post_summary(posts: list, ticker: str) -> str:
    # We call our new diagnostic wrapper instead of the original function.
    return run_gemini_summarizer(posts, ticker)

# --- Standalone Execution Logic (Unchanged) ---
def process_ticker_posts(ticker: str, debug: bool = False):
    posts_filename = os.path.join("output", f"{ticker}_posts.json")
    emails_filename = os.path.join("output", f"{ticker}_sent_emails.json")
    combined_filename = os.path.join("output", f"{ticker}_combined_debug.json")
    
    try:
        combined_data = []
        if os.path.exists(posts_filename):
            with open(posts_filename, "r", encoding="utf-8") as f:
                combined_data.extend(json.load(f))
        if os.path.exists(emails_filename):
            with open(emails_filename, "r", encoding="utf-8") as f:
                combined_data.extend(json.load(f))

        if debug:
            sorted_posts = sorted(combined_data, key=lambda x: x['timestamp'], reverse=True)
            with open(combined_filename, "w", encoding="utf-8") as f:
                json.dump(sorted_posts, f, indent=4)
            print(f"Combined data written to {combined_filename} for debugging.")

        if combined_data:
            print(f"Attempting to generate post summaries for {ticker}")
            summary = generate_post_summary(combined_data, ticker)
            
            output_filename = os.path.join("output", f"{ticker}_post_summary.txt")
            with open(output_filename, "w", encoding="utf-8") as f:
                f.write(summary)
            print(f"Summary has been saved to {output_filename}")
            return summary
        else:
            print(f"No posts or emails found for {ticker}. Skipping summary.")
            return "No forum summary available."
    except Exception as e:
        print(f"Error processing data for {ticker}: {e}")
        return "Error generating summary."

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='Generate weighted summaries of ticker posts and emails using Gemini.')
    parser.add_argument('ticker', type=str, help='Stock ticker symbol (e.g., AAPL)')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode')
    args = parser.parse_args()
    process_ticker_posts(args.ticker.upper(), args.debug)