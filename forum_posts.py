# forum_posts.py

import os
import json
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

WEBSITETOOLBOX_API_KEY = os.getenv("WEBSITETOOLBOX_API_KEY")
WEBSITETOOLBOX_USERNAME = os.getenv("WEBSITETOOLBOX_USERNAME")

BASE_URL = "https://api.websitetoolbox.com/v1/api"

HEADERS = {
    "Accept": "application/json",
    "x-api-key": WEBSITETOOLBOX_API_KEY
}


def get_categories():
    url = f"{BASE_URL}/categories"
    response = requests.get(url, headers=HEADERS)
    response.raise_for_status()
    return response.json()


def find_category_by_title(title):
    categories_data = get_categories()
    for category in categories_data.get("data", []):
        if category.get("title") == title:
            return category
    return None


def get_subcategories(all_categories, parent_id):
    subcats = []
    for cat in all_categories.get("data", []):
        if cat.get("parentId") == parent_id:
            subcats.append(cat)
            subcats.extend(get_subcategories(all_categories, cat["categoryId"]))
    return subcats


def get_topics_for_category(category_id):
    url = f"{BASE_URL}/topics"
    params = {"categoryId": category_id}
    response = requests.get(url, headers=HEADERS, params=params)
    response.raise_for_status()
    return response.json()


def get_posts_for_topic(topic_id):
    """
    Fetch posts for a topic.
    Handles WebsiteToolbox quirks:
    - pagination required
    - restricted/archived topics return 400
    - empty body responses
    """
    url = f"{BASE_URL}/posts"
    params = {
        "topicId": topic_id,
        "page": 1,
        "pageSize": 100
    }

    response = requests.get(url, headers=HEADERS, params=params)

    # WebsiteToolbox uses 400 for restricted/invalid topics
    if response.status_code == 400:
        print(f"Skipping topic {topic_id}: API returned 400")
        return {"data": []}

    response.raise_for_status()

    # Guard against empty body
    if not response.content:
        return {"data": []}

    return response.json()


def load_ticker_config():
    try:
        with open("forum_search_config.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
    except json.JSONDecodeError:
        print("Warning: forum_search_config.json is invalid JSON")
        return {}


def get_search_ticker(input_ticker, config):
    search_ticker = config.get(input_ticker)
    if search_ticker:
        print(f"Found mapping: using '{search_ticker}' instead of '{input_ticker}'")
        return search_ticker

    print(f"No mapping found: using provided ticker '{input_ticker}'")
    return input_ticker


def fetch_all_for_ticker(input_ticker):
    config = load_ticker_config()
    ticker = get_search_ticker(input_ticker, config)

    all_categories = get_categories()
    cat_list = all_categories.get("data", [])

    if not cat_list:
        print("No categories returned from the API.")
        return

    parent_cat = find_category_by_title(ticker)
    if not parent_cat:
        print(f"No category found with title '{ticker}'.")
        return

    parent_id = parent_cat["categoryId"]
    print(f"Found category '{ticker}' (ID={parent_id}).")

    subcategories = get_subcategories(all_categories, parent_id)
    print(f"Found {len(subcategories)} subcategory(ies) under '{ticker}'.")

    relevant_categories = [parent_cat] + subcategories
    unique_posts = {}

    for cat in relevant_categories:
        cat_id = cat["categoryId"]
        cat_title = cat["title"]

        try:
            topics_data = get_topics_for_category(cat_id)
        except requests.HTTPError as e:
            print(f"Skipping category '{cat_title}' (ID={cat_id}): {e}")
            continue

        topics_list = topics_data.get("data", [])
        print(f"Category '{cat_title}' (ID={cat_id}) -> {len(topics_list)} topic(s).")

        for topic in topics_list:
            topic_id = topic.get("topicId")
            topic_title = topic.get("title")

            try:
                posts_data = get_posts_for_topic(topic_id)
            except requests.HTTPError as e:
                print(f"Skipping topic '{topic_title}' (ID={topic_id}): {e}")
                continue

            posts_list = posts_data.get("data", [])
            print(f"  Topic '{topic_title}' (ID={topic_id}) -> {len(posts_list)} post(s).")

            for post in posts_list:
                post_id = post.get("postId")
                if post_id not in unique_posts:
                    unique_posts[post_id] = post

    simplified_posts = []

    for post in unique_posts.values():
        raw_html = post.get("message", "")
        soup = BeautifulSoup(raw_html, "html.parser")
        clean_message = soup.get_text(separator=" ").strip()

        author_email = post.get("author", {}).get("email", "")

        simplified_posts.append({
            "timestamp": post.get("postTimestamp", 0),
            "message": clean_message,
            "authorEmail": author_email
        })

    simplified_posts.sort(key=lambda p: p["timestamp"])

    os.makedirs("output", exist_ok=True)
    output_path = os.path.join("output", f"{ticker}_posts.json")

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(simplified_posts, f, indent=2, ensure_ascii=False)

    print(f"Saved {len(simplified_posts)} posts to '{output_path}'.")


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 2:
        print("Usage: python forum_posts.py TICKER")
        print("Example: python forum_posts.py AAPL")
        sys.exit(1)

    fetch_all_for_ticker(sys.argv[1])
