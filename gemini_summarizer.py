# analysis_project/gemini_summarizer.py

import os
import json
import asyncio
import re
from datetime import datetime
from typing import List, Dict, Any
import google.generativeai as genai
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Configure the Gemini API client
gemini_api_key = os.getenv("GEMINI_API_KEY")
if not gemini_api_key:
    raise ValueError("GEMINI_API_KEY not found in environment. Please set it in your .env file.")
genai.configure(api_key=gemini_api_key)

async def _generate_post_summary_async(posts: List[Dict[str, Any]], ticker: str) -> str:
    """
    Async helper function to generate a summary using the Gemini API.
    """
    # Sort posts by timestamp (most recent first)
    sorted_posts = sorted(posts, key=lambda x: x.get('timestamp', 0), reverse=True)
    
    # Format the posts into a single string for the prompt
    content_lines = [
        f"Date: {datetime.fromtimestamp(post['timestamp'])}\nMessage:\n{post['message']}\n\n"
        for post in sorted_posts
    ]
    
    # Define the system prompt with clear instructions for the model
    system_prompt = """You are an expert financial analyst. Your task is to summarize the key insights about the given ticker in exactly 10 concise bullet points.

Requirements:
- Each bullet point must start with a bold topic phrase followed by a colon (e.g., **Topic:** Detail).
- Focus on concrete facts, financial metrics, and in-depth analysis.
- Give more weight to the most recent information and any content from 'smgacm@gmail.com'.
- Do not mention the sources, dates, or email addresses in your summary.
- Present all information as direct market insights.
- Number each bullet point from 1 to 10."""

    # Combine the instructions and the content into a single user prompt
    user_prompt = f"Here are the posts about {ticker} for you to analyze (newest first):\n\n" + "".join(content_lines)

    # Set up the generation configuration
    generation_config = {
        "temperature": 0.3,
        "max_output_tokens": 2000,
    }

    # Initialize the Gemini model with the more powerful Gemini 1.5 Pro
    model = genai.GenerativeModel(
        model_name="gemini-1.5-pro",
        generation_config=generation_config,
        system_instruction=system_prompt
    )

    try:
        # Generate the content asynchronously
        response = await model.generate_content_async([user_prompt])
        summary_text = response.text.strip()
        
        # In case the model adds an introductory phrase, find the start of the list
        match = re.search(r'1\.\s+\*\*', summary_text)
        if match:
            summary_text = summary_text[match.start():]
            
        return summary_text
    except Exception as e:
        print(f"Error during Gemini API call: {str(e)}")
        raise

def generate_post_summary(posts: list, ticker: str) -> str:
    """
    Generates a weighted summary of forum posts and emails using the Gemini API.
    This function maintains the same interface as the original summarizer.
    """
    if not posts:
        return "No posts were provided to summarize."
        
    try:
        return asyncio.run(_generate_post_summary_async(posts, ticker))
    except Exception as e:
        print(f"An error occurred while generating the summary: {str(e)}")
        return "Error: Summary could not be generated."

def process_ticker_posts(ticker: str, debug: bool = False):
    """
    Process posts and emails for a given ticker and generate a summary.
    """
    # Define filenames
    posts_filename = os.path.join("output", f"{ticker}_posts.json")
    emails_filename = os.path.join("output", f"{ticker}_sent_emails.json")
    combined_filename = os.path.join("output", f"{ticker}_combined_debug.json")
    
    try:
        combined_data = []

        # Load posts if the file exists
        if os.path.exists(posts_filename):
            with open(posts_filename, "r", encoding="utf-8") as f:
                combined_data.extend(json.load(f))

        # Load emails if the file exists
        if os.path.exists(emails_filename):
            with open(emails_filename, "r", encoding="utf-8") as f:
                combined_data.extend(json.load(f))

        # Optionally write combined data to a debug file
        if debug:
            sorted_posts = sorted(combined_data, key=lambda x: x['timestamp'], reverse=True)
            with open(combined_filename, "w", encoding="utf-8") as f:
                json.dump(sorted_posts, f, indent=4)
            print(f"Combined data written to {combined_filename} for debugging.")

        if combined_data:
            print(f"Attempting to generate post summaries for {ticker}")
            sorted_posts = sorted(combined_data, key=lambda x: x['timestamp'], reverse=True)
            summary = generate_post_summary(sorted_posts, ticker)
            
            # Write summary to output file
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
    parser.add_argument('--debug', action='store_true', help='Enable debug mode to write combined data to file')
    
    args = parser.parse_args()
    
    process_ticker_posts(args.ticker.upper(), args.debug)