import os
import json
import asyncio
import random
import time
from datetime import datetime
from typing import List, Dict, Any
from anthropic import AsyncAnthropic
from dotenv import load_dotenv

load_dotenv()

async def _generate_post_summary_async_with_retry(posts: List[Dict[str, Any]], ticker: str, max_retries: int = 5) -> str:
    """
    Async helper function to generate summary using Claude with enhanced retry logic for 529 errors.
    """
    client = AsyncAnthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    
    # Sort posts by timestamp (most recent first)
    sorted_posts = sorted(posts, key=lambda x: x['timestamp'], reverse=True)
    
    # Limit the number of posts to prevent overwhelming the API
    max_posts = 70  # Reduced from potentially unlimited posts
    if len(sorted_posts) > max_posts:
        print(f"Limiting analysis to {max_posts} most recent posts (from {len(sorted_posts)} total)")
        sorted_posts = sorted_posts[:max_posts]
    
    content_lines = []
    total_chars = 0
    max_chars = 8000  # Conservative limit to stay well under API limits
    
    for post in sorted_posts:
        post_content = f"Date: {datetime.fromtimestamp(post['timestamp'])}\nMessage:\n{post['message']}\n\n"
        if total_chars + len(post_content) > max_chars:
            print(f"Truncating content at {total_chars} characters to stay within API limits")
            break
        content_lines.append(post_content)
        total_chars += len(post_content)
    
    system_prompt = """You are an expert financial analyst. Summarize the key insights about the given ticker in exactly 10 concise bullet points.
    
Requirements:
- Each bullet point must start with a bold topic phrase followed by a colon
- Focus on concrete facts, metrics, and analyses
- Give more weight to recent information and content from smgacm@gmail.com
- Do not mention sources, dates, or emails in the summary
- Present information directly as market insights
- Number each bullet point 1-10"""

    user_prompt = f"Here are posts about {ticker} to analyze (newest first):\n\n" + "".join(content_lines)
    
    for attempt in range(max_retries):
        try:
            # Add a small delay before each attempt to be respectful to the API
            if attempt > 0:
                base_delay = 2.0
                # Exponential backoff with jitter
                delay = base_delay * (2 ** (attempt - 1)) + random.uniform(0, 2)
                print(f"Retrying API call in {delay:.1f} seconds... (attempt {attempt + 1}/{max_retries})")
                await asyncio.sleep(delay)
            
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
            
            print(f"✅ Successfully generated summary for {ticker}")
            return response
            
        except Exception as e:
            error_str = str(e)
            print(f"API call attempt {attempt + 1} failed: {error_str}")
            
            # Check if it's a 529 overload error
            if "529" in error_str or "overloaded" in error_str.lower():
                if attempt == max_retries - 1:
                    print(f"❌ Max retries ({max_retries}) exceeded for overload errors")
                    raise Exception(f"API overloaded after {max_retries} attempts. Please try again later.")
                else:
                    print(f"🔄 API overloaded, will retry...")
                    continue
            else:
                # For non-529 errors, don't retry
                print(f"❌ Non-retryable error: {error_str}")
                raise e
    
    raise Exception(f"Failed after {max_retries} attempts")

async def _generate_post_summary_async(posts: List[Dict[str, Any]], ticker: str) -> str:
    """
    Original async helper function - now calls the retry version.
    """
    return await _generate_post_summary_async_with_retry(posts, ticker)

def generate_post_summary(posts: list, ticker: str) -> str:
    """
    Generate a weighted summary of posts using the Anthropic API.
    Enhanced with better error handling and retry logic.
    """
    try:
        # Validate input
        if not posts:
            return "No posts available for analysis."
        
        if not os.getenv("ANTHROPIC_API_KEY"):
            return "Error: ANTHROPIC_API_KEY not found in environment variables."
        
        print(f"🔄 Processing {len(posts)} posts for {ticker}...")
        return asyncio.run(_generate_post_summary_async(posts, ticker))
        
    except Exception as e:
        error_msg = str(e)
        print(f"❌ Error generating summary: {error_msg}")
        
        # Provide different messages based on error type
        if "overloaded" in error_msg.lower() or "529" in error_msg:
            return f"Summary temporarily unavailable due to API overload. Please try again in a few minutes. ({len(posts)} posts were ready for analysis)"
        elif "api_key" in error_msg.lower():
            return "Error: Invalid or missing API key."
        else:
            return f"Summary generation failed: {error_msg}"

def process_ticker_posts(ticker: str, debug: bool = False, max_posts: int = 20):
    """
    Process posts and emails for a given ticker and generate a summary.
    Enhanced with post limiting and better error handling.
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
                print(f"📄 Loaded {len(posts)} posts from {posts_filename}")

        # Load emails if the file exists
        if os.path.exists(emails_filename):
            with open(emails_filename, "r", encoding="utf-8") as f:
                emails = json.load(f)
                combined_data.extend(emails)
                print(f"📧 Loaded {len(emails)} emails from {emails_filename}")

        if not combined_data:
            print(f"⚠️  No posts or emails found for {ticker}. Skipping summary.")
            return "No forum summary available."

        # Sort posts by timestamp (most recent first)
        sorted_posts = sorted(combined_data, key=lambda x: x['timestamp'], reverse=True)
        
        # Limit posts to prevent API overload
        if len(sorted_posts) > max_posts:
            print(f"📊 Limiting to {max_posts} most recent posts (from {len(sorted_posts)} total)")
            sorted_posts = sorted_posts[:max_posts]

        # Optionally write combined data to a debug file
        if debug:
            with open(combined_filename, "w", encoding="utf-8") as f:
                json.dump(sorted_posts, f, indent=4)
            print(f"🐛 Debug: Combined data written to {combined_filename}")

        print(f"🤖 Attempting to generate summary for {ticker} using {len(sorted_posts)} items...")
        summary = generate_post_summary(sorted_posts, ticker)
        
        # Write summary to output file
        output_filename = os.path.join("output", f"{ticker}_post_summary.txt")
        with open(output_filename, "w", encoding="utf-8") as f:
            f.write(summary)
        
        print(f"💾 Summary saved to {output_filename}")
        return summary
            
    except Exception as e:
        error_msg = f"Error processing data for {ticker}: {e}"
        print(f"❌ {error_msg}")
        return f"Error processing posts: {str(e)}"

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Generate weighted summaries of ticker posts and emails')
    parser.add_argument('ticker', type=str, help='Stock ticker symbol (e.g., AAPL)')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode to write combined data to file')
    parser.add_argument('--max_posts', type=int, default=20, help='Maximum number of posts to analyze (default: 20)')
    
    args = parser.parse_args()
    
    process_ticker_posts(args.ticker.upper(), args.debug, args.max_posts)