# forum_posts.py

import os
import json
import time
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

WEBSITETOOLBOX_API_KEY = os.getenv("WEBSITETOOLBOX_API_KEY")

BASE_URL = "https://api.websitetoolbox.com/v1/api"

HEADERS = {
    "Accept": "application/json",
    "x-api-key": WEBSITETOOLBOX_API_KEY
}

MAX_RETRIES = 3
PAGE_SIZE = 100


# ---------------------------
# Low-level helpers
# ---------------------------

def _request_with_retry(url, params):
    for attempt in range(MAX_RETRIES):
        try:
            response = requests.get(url, headers=HEADERS, params=params)

            # Restricted / archived resources
            if response.status_code == 400:
                return None

            response.raise_for_status()

            if not response.content:
                return None

            return response.json()

        except requests.exceptions.RequestException as e:
            if attempt < MAX_RETRIES - 1:
                time.sleep(2 ** attempt)
            else:
                print(f"❌ Giving up on {url} params={params}: {e}")
                return None


def _paginate(endpoint, base_params):
    """
    Generator yielding items across all pages.
    """
    page = 1
    while True:
        params = dict(base_params)
        params["page"] = page
        params["pageSize"] = PAGE_SIZE

        data = _request_with_retry(f"{BASE_URL}/{endpoint}", params)
        if not data:
            break

        items = data.get("data", [])
        if not items:
            break

        for item in items:
            yield item

        # Stop if last page
        total = data.get("totalSize")
        if total is not None and page * PAGE_SIZE >= total:
            break

        page += 1


# ---------------------------
# API wrappers
# ---------------------------

def get_categories():
    data = _request_with_retry(f"{BASE_URL}/categories", {})
    return data or {"data": []}


def find_category_by_title(title):
    categories = get_categories().get("data", [])
    for category in categories:
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
    return list(_paginate("topics", {"categoryId": category_id}))


def get_posts_for_topic(topic_id):
    return list(_paginate("posts", {"topicId": topic_id}))


# ---------------------------
# Ticker config
# ---------------------------

def load_ticker_config():
    try:
        with open("forum_search_config.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def get_search_ticker(input_ticker, config):
    return config.get(input_ticker, input_ticker)


# ---------------------------
# Main pipeline
# ---------------------------

def fetch_all_for_ticker(input_ticker):
    config = load_ticker_config()
    ticker = get_search_ticker(input_ticker, config)

    all_categories = get_categories()
    parent_cat = find_category_by_title(ticker)

    if not parent_cat:
        print(f"No category found with title '{ticker}'.")
        return

    parent_id = parent_cat["categoryId"]
    print(f"Found category '{ticker}' (ID={parent_id}).")

    subcategories = get_subcategories(all_categories, parent_id)
    relevant_categories = [parent_cat] + subcategories

    # Deduplicate categories
    seen_cat_ids = set()
    unique_categories = []
    for cat in relevant_categories:
        cid = cat["categoryId"]
        if cid not in seen_cat_ids:
            unique_categories.append(cat)
            seen_cat_ids.add(cid)

    unique_posts = {}
    seen_topics = set()

    for cat in unique_categories:
        cat_id = cat["categoryId"]
        cat_title = cat["title"]

        topics = get_topics_for_category(cat_id)
        print(f"Category '{cat_title}' (ID={cat_id}) -> {len(topics)} topic(s).")

        for topic in topics:
            topic_id = topic.get("topicId")
            topic_title = topic.get("title")

            if topic_id in seen_topics:
                continue
            seen_topics.add(topic_id)

            posts = get_posts_for_topic(topic_id)
            print(f"  Topic '{topic_title}' (ID={topic_id}) -> {len(posts)} post(s).")

            for post in posts:
                post_id = post.get("postId")
                if post_id not in unique_posts:
                    unique_posts[post_id] = post

    # ---------------------------
    # Clean + save output
    # ---------------------------

    simplified_posts = []

    for post in unique_posts.values():
        soup = BeautifulSoup(post.get("message", ""), "html.parser")
        clean_message = soup.get_text(separator=" ").strip()

        simplified_posts.append({
            "timestamp": post.get("postTimestamp", 0),
            "message": clean_message,
            "authorEmail": post.get("author", {}).get("email", "")
        })

    simplified_posts.sort(key=lambda p: p["timestamp"])

    os.makedirs("output", exist_ok=True)
    output_path = os.path.join("output", f"{ticker}_posts.json")

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(simplified_posts, f, indent=2, ensure_ascii=False)

    print(f"Saved {len(simplified_posts)} posts to '{output_path}'.")


# ---------------------------
# CLI
# ---------------------------

if __name__ == "__main__":
    import sys

    if len(sys.argv) != 2:
        print("Usage: python forum_posts.py TICKER")
        sys.exit(1)

    fetch_all_for_ticker(sys.argv[1])
