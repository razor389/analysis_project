# analysis_project/gemini_summarizer.py

import os
import json
import asyncio
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

    # Initialize the Gemini model
    model = genai.GenerativeModel(
        model_name="gemini-1.5-flash",
        generation_config=generation_config,
        system_instruction=system_prompt
    )

    try:
        # Generate the content asynchronously
        response = await model.generate_content_async([user_prompt])
        
        # Extract the text and clean it up
        summary_text = response.text.strip()
        
        # In case the model adds an introductory phrase, find the start of the list
        import re
        match = re.search(r'1\.\s+\*\*', summary_text)
        if match:
            summary_text = summary_text[match.start():]
            
        return summary_text
    except Exception as e:
        print(f"Error during Gemini API call: {str(e)}")
        # Raise the exception to be caught by the calling function
        raise

def generate_post_summary(posts: list, ticker: str) -> str:
    """
    Generates a weighted summary of forum posts and emails using the Gemini API.
    This function maintains the same interface as the original summarizer.
    """
    if not posts:
        return "No posts were provided to summarize."
        
    try:
        # Run the asynchronous summary generation function
        return asyncio.run(_generate_post_summary_async(posts, ticker))
    except Exception as e:
        print(f"An error occurred while generating the summary: {str(e)}")
        return "Error: Summary could not be generated."

if __name__ == '__main__':
    # This part is for standalone testing of the summarizer
    import argparse
    
    parser = argparse.ArgumentParser(description='Generate summaries of ticker-related posts and emails using Gemini.')
    parser.add_argument('ticker', type=str, help='The stock ticker symbol (e.g., AAPL).')
    
    args = parser.parse_args()
    ticker = args.ticker.upper()
    
    # Define filenames for input
    posts_filename = os.path.join("output", f"{ticker}_posts.json")
    emails_filename = os.path.join("output", f"{ticker}_sent_emails.json")
    
    combined_data = []
    # Load posts if available
    if os.path.exists(posts_filename):
        with open(posts_filename, "r", encoding="utf-8") as f:
            combined_data.extend(json.load(f))

    # Load emails if available
    if os.path.exists(emails_filename):
        with open(emails_filename, "r", encoding="utf-8") as f:
            combined_data.extend(json.load(f))
            
    if combined_data:
        print(f"Generating summary for {ticker}...")
        summary = generate_post_summary(combined_data, ticker)
        
        # Save the summary to an output file
        output_filename = os.path.join("output", f"{ticker}_gemini_summary.txt")
        with open(output_filename, "w", encoding="utf-8") as f:
            f.write(summary)
            
        print(f"Summary saved successfully to {output_filename}")
        print("\n--- Summary ---")
        print(summary)
    else:
        print(f"No data found for ticker {ticker}. Could not generate summary.")