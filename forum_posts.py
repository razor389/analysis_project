# forum_posts.py

import os
import re
import json
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

WEBSITETOOLBOX_API_KEY = os.getenv("WEBSITETOOLBOX_API_KEY")
WEBSITETOOLBOX_USERNAME = os.getenv("WEBSITETOOLBOX_USERNAME")

# If your forum ID is "my_forum_id", you might need:
# BASE_URL = "https://api.websitetoolbox.com/v1/my_forum_id"
BASE_URL = 'https://api.websitetoolbox.com/v1/api'

# Common request headers
HEADERS = {
    'Accept': 'application/json',
    'x-api-key': WEBSITETOOLBOX_API_KEY
}

def get_categories():
    """
    Fetch all categories (only first page).
    If you have many categories, handle pagination with has_more, total_size, etc.
    """
    url = f"{BASE_URL}/categories"
    response = requests.get(url, params={}, headers=HEADERS)
    response.raise_for_status()
    return response.json()

def find_category_by_title(title):
    """
    Searches through the category list for a title match.
    Returns the category dict if found, else None.
    """
    categories_data = get_categories()
    for category in categories_data.get("data", []):
        if category.get("title") == title:
            return category
    return None

def get_subcategories(all_categories, parent_id):
    """
    Recursively find all subcategories whose parentId == parent_id.
    Return a list of subcategory objects (including deeper sub-subcategories).
    """
    subcats = []
    for cat in all_categories.get("data", []):
        if cat.get("parentId") == parent_id:
            subcats.append(cat)
            subcats += get_subcategories(all_categories, cat["categoryId"])
    return subcats

def get_topics_for_category(category_id):
    """
    Fetches topics for the given categoryId (first page).
    If there are many topics, handle pagination if needed.
    """
    url = f"{BASE_URL}/topics"
    params = {"categoryId": category_id}
    response = requests.get(url, headers=HEADERS, params=params)
    response.raise_for_status()
    return response.json()

def get_posts_for_topic(topic_id):
    """
    Fetches posts for the given topicId (first page).
    If there are many posts, handle pagination if needed.
    """
    url = f"{BASE_URL}/posts"
    params = {"topicId": topic_id}
    response = requests.get(url, headers=HEADERS, params=params)
    response.raise_for_status()
    return response.json()

def fetch_all_for_ticker(ticker):
    """
    1) Find the parent category for 'ticker'
    2) Gather all subcategories
    3) For each (sub)category, fetch topics, then posts
    4) Collect all posts, remove duplicates
    5) Clean HTML from each post, sort by timestamp, save to {ticker}_posts.json
    """
    # Get all categories up front
    all_categories = get_categories()
    cat_list = all_categories.get("data", [])
    if not cat_list:
        print("No categories returned from the API.")
        return

    # Find the main category by title
    parent_cat = find_category_by_title(ticker)
    if not parent_cat:
        print(f"No category found with title '{ticker}'.")
        return

    parent_id = parent_cat["categoryId"]
    print(f"Found category '{ticker}' (ID={parent_id}).")

    # Get all subcategories recursively
    subcategories = get_subcategories(all_categories, parent_id)
    print(f"Found {len(subcategories)} subcategory(ies) under '{ticker}'.")

    # Combine parent + subcategories
    relevant_categories = [parent_cat] + subcategories

    # We'll collect unique posts in a dict keyed by postId
    unique_posts = {}

    # For each category, fetch topics -> posts
    for cat in relevant_categories:
        cat_id = cat["categoryId"]
        cat_title = cat["title"]
        topics_data = get_topics_for_category(cat_id)
        topics_list = topics_data.get("data", [])

        print(f"Category '{cat_title}' (ID={cat_id}) -> {len(topics_list)} topic(s).")

        for topic in topics_list:
            topic_id = topic.get("topicId")
            topic_title = topic.get("title")

            posts_data = get_posts_for_topic(topic_id)
            posts_list = posts_data.get("data", [])

            print(f"  Topic '{topic_title}' (ID={topic_id}) -> {len(posts_list)} post(s).")

            # Collect each post
            for post in posts_list:
                post_id = post.get("postId")
                if post_id not in unique_posts:
                    unique_posts[post_id] = post

    # Now let's create a simplified, cleaned list of posts
    simplified_posts = []
    for post_id, post in unique_posts.items():
        # Extract the raw HTML message
        raw_html = post.get("message", "")
        # Clean out HTML tags using BeautifulSoup
        soup = BeautifulSoup(raw_html, "html.parser")
        clean_message = soup.get_text(separator=" ").strip()

        # Extract the author's email
        author_email = post.get("author", {}).get("email", "")

        # Build a simplified post record (adding the author email)
        simplified_posts.append({
            "timestamp": post.get("postTimestamp", 0),
            "message": clean_message,
            "authorEmail": author_email
        })

    # Sort all simplified posts by timestamp (ascending)
    simplified_posts.sort(key=lambda p: p["timestamp"])

    # Save to output/{ticker}_posts.json
    os.makedirs("output", exist_ok=True)  # Ensure 'output' directory exists
    json_filename = os.path.join("output", f"{ticker}_posts.json")
    with open(json_filename, "w", encoding="utf-8") as f:
        json.dump(simplified_posts, f, indent=2, ensure_ascii=False)

    print(f"Saved {len(simplified_posts)} posts to '{json_filename}'.")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) != 2:
        print("Usage: python forum_posts.py TICKER")
        print("Example: python forum_posts.py AAPL")
        sys.exit(1)
        
    ticker = sys.argv[1]
    fetch_all_for_ticker(ticker)