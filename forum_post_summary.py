# forum_post_summary.py

import os
import json
from openai import OpenAI
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()  # Load .env variables

def generate_post_summary(posts: list, ticker: str) -> str:
    """
    Generate a weighted summary of posts using the OpenAI API.
    Emphasis is on more recent data and smgacm@gmail.com content,
    but we do not explicitly mention 'posts' or 'authors' in the final summary.
    """

    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
    
    # Sort posts by timestamp (most recent first)
    sorted_posts = sorted(posts, key=lambda x: x['timestamp'], reverse=True)
    
    # Build a descriptive message
    prompt_intro = f"""
    Summarize key insights about {ticker} in concise bullet points (no more than 10).
    Give more weight to newer information and content from the email "smgacm@gmail.com," 
    but do not mention authors, emails, or 'posts' in your summary. 
    Present the summary directly, as if stating the facts or perspectives 
    without referencing that they come from older or newer messages.

    Important formatting requirements:
    - No title.
    - Produce numbered bullet points (1 to 10 max).
    - Each bullet point must begin with a bold phrase summarizing the point, 
      like this example:
        **1) De-globalization and Apple’s Challenges in China:** Explanation here...

    ---
    Here are the raw messages (newest first):
    """

    content = []
    for post in sorted_posts:
        timestamp = datetime.fromtimestamp(post['timestamp'])
        # Convert the post to a direct snippet (no mention of 'author' or 'post')
        # You can optionally keep or remove the date, depending on your preference
        snippet = f"Date: {timestamp}\nMessage:\n{post['message']}\n\n"
        content.append(snippet)
    
    # Join everything into a single message to the model
    system_prompt = (
        "You are a helpful assistant that produces direct, factual summaries. "
        "Follow the exact formatting requirements provided. "
        "Do not include a title, only numbered bullet points. "
        "Each bullet point's header is bold, followed by a colon."
    )
    user_prompt = prompt_intro + "\n".join(content)

    # Now create the ChatCompletion
    completion = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        max_tokens=1000,
        temperature=0.3,
    )

    # Return the assistant's content
    return completion.choices[0].message.content.strip()

def process_ticker_posts(ticker: str):
    """
    Process posts for a given ticker and generate a summary.
    """
    filename = f"{ticker}_posts.json"
    with open(filename, 'r', encoding='utf-8') as f:
        posts = json.load(f)
    
    summary = generate_post_summary(posts, ticker)
    
    output_filename = f"{ticker}_post_summary.txt"
    with open(output_filename, 'w', encoding='utf-8') as f:
        f.write(summary)
    
    print(f"Summary has been saved to {output_filename}")

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Generate weighted summaries of ticker posts')
    parser.add_argument('ticker', type=str, help='Stock ticker symbol (e.g., AAPL)')
    
    args = parser.parse_args()
    process_ticker_posts(args.ticker.upper())
