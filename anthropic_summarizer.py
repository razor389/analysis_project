import os
import json
import asyncio
from datetime import datetime
from typing import List, Dict, Any
from anthropic import AsyncAnthropic
from dotenv import load_dotenv

load_dotenv()

async def _generate_post_summary_async(posts: List[Dict[str, Any]], ticker: str) -> str:
    """
    Async helper function to generate summary using Claude.
    """
    client = AsyncAnthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    
    # Sort posts by timestamp (most recent first)
    sorted_posts = sorted(posts, key=lambda x: x['timestamp'], reverse=True)
    
    content_lines = [
        f"Date: {datetime.fromtimestamp(post['timestamp'])}\nMessage:\n{post['message']}\n\n"
        for post in sorted_posts
    ]
    
    system_prompt = """You are an expert financial analyst. Summarize the key insights about the given ticker in exactly 10 concise bullet points.
    
Requirements:
- Each bullet point must start with a bold topic phrase followed by a colon
- Focus on concrete facts, metrics, and analyses
- Give more weight to recent information and content from smgacm@gmail.com
- Do not mention sources, dates, or emails in the summary
- Present information directly as market insights
- Number each bullet point 1-10"""

    user_prompt = f"Here are posts about {ticker} to analyze (newest first):\n\n" + "".join(content_lines)
    
    try:
        message = await client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=2000,
            temperature=0.3,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}]
        )
        # Get the response and remove any introductory text before the numbered list
        response = message.content[0].text.strip()
        # Find the first numbered bullet point
        import re
        match = re.search(r'1\.\s+\*\*', response)
        if match:
            response = response[match.start():].strip()
        return response
    except Exception as e:
        print(f"API call failed: {str(e)}")
        raise

def generate_post_summary(posts: list, ticker: str) -> str:
    """
    Generate a weighted summary of posts using the Anthropic API.
    Maintains the same interface as the original OpenAI version.
    """
    try:
        return asyncio.run(_generate_post_summary_async(posts, ticker))
    except Exception as e:
        print(f"Error generating summary: {str(e)}")
        return "Error generating summary."

def process_ticker_posts(ticker: str, debug: bool = False):
    """
    Process posts and emails for a given ticker and generate a summary.
    """
    # Define filenames
    posts_filename = os.path.join("output", f"{ticker}_posts.json")
    emails_filename = os.path.join("output", f"{ticker}_sent_emails.json")
    combined_filename = os.path.join("output", f"{ticker}_combined_debug.json")
    
    try:
        # Initialize combined data
        combined_data = []

        # Load posts if the file exists
        if os.path.exists(posts_filename):
            with open(posts_filename, "r", encoding="utf-8") as f:
                posts = json.load(f)
                combined_data.extend(posts)

        # Load emails if the file exists
        if os.path.exists(emails_filename):
            with open(emails_filename, "r", encoding="utf-8") as f:
                emails = json.load(f)
                combined_data.extend(emails)

        # Optionally write combined data to a debug file
        if debug:
            sorted_posts = sorted(combined_data, key=lambda x: x['timestamp'], reverse=True)
            with open(combined_filename, "w", encoding="utf-8") as f:
                json.dump(sorted_posts, f, indent=4)
            print(f"Combined data written to {combined_filename} for debugging.")

        if combined_data:
            print(f"Attempting to generate post summaries for {ticker}")
            # Sort posts by timestamp (most recent first)
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
    
    parser = argparse.ArgumentParser(description='Generate weighted summaries of ticker posts and emails')
    parser.add_argument('ticker', type=str, help='Stock ticker symbol (e.g., AAPL)')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode to write combined data to file')
    
    args = parser.parse_args()
    
    process_ticker_posts(args.ticker.upper(), args.debug)