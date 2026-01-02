# forum_posts.py

import os
import json
import time
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from collections import Counter

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

SCOTT_EMAIL = "smgacm@gmail.com"


# ---------------------------
# Low-level helpers
# ---------------------------

def _extract_author_email(post: dict) -> str:
    """
    WebsiteToolbox payloads can vary. Try a few common shapes.
    """
    # 1) current assumed shape
    author = post.get("author") or {}
    email = author.get("email")
    if email:
        return str(email)

    # 2) sometimes direct
    email = post.get("authorEmail")
    if email:
        return str(email)

    # 3) other possible shapes (defensive)
    user = post.get("user") or {}
    email = user.get("email")
    if email:
        return str(email)

    created_by = post.get("createdBy") or {}
    email = created_by.get("email")
    if email:
        return str(email)

    return ""


def _category_tree_ids(all_categories: dict, root_id: int) -> list[int]:
    """
    root + all descendants (ids)
    """
    ids = [root_id]
    for c in get_subcategories(all_categories, root_id):
        cid = c.get("categoryId")
        if cid is not None:
            ids.append(cid)
    return ids

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


def _clean_html_to_text(html: str) -> str:
    soup = BeautifulSoup(html or "", "html.parser")
    return soup.get_text(separator=" ").strip()


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
    """
    Recursively collect all descendants under parent_id.
    """
    subcats = []
    for cat in all_categories.get("data", []):
        if cat.get("parentId") == parent_id:
            subcats.append(cat)
            subcats.extend(get_subcategories(all_categories, cat["categoryId"]))
    return subcats


def get_direct_subcategories(all_categories, parent_id):
    """
    Only the immediate children of parent_id (non-recursive).
    """
    return [
        cat for cat in all_categories.get("data", [])
        if cat.get("parentId") == parent_id
    ]


def find_descendant_category_by_title(all_categories, parent_id, title):
    """
    Find a descendant (or direct child) of parent_id whose title matches exactly.
    """
    for cat in get_subcategories(all_categories, parent_id):
        if cat.get("title") == title:
            return cat
    return None


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
# New: pre-LLM moat-threat data gathering
# ---------------------------

def fetch_moat_threat_source_for_ticker(
    input_ticker,
    author_email=SCOTT_EMAIL,
    *,
    include_descendants=True,     # search topics in each moat subcat's subtree
    require_author=True,          # if False, do not filter by author
    debug=False                   # print under-the-hood stats
):
    config = load_ticker_config()
    ticker = get_search_ticker(input_ticker, config)

    all_categories = get_categories()

    parent_cat = None
    for cat in all_categories.get("data", []):
        if cat.get("title") == ticker:
            parent_cat = cat
            break

    if not parent_cat:
        print(f"Ticker not found: no category found with title '{ticker}'.")
        return None

    parent_id = parent_cat["categoryId"]

    thesis_title = f"{ticker} Investment Thesis"
    thesis_cat = find_descendant_category_by_title(all_categories, parent_id, thesis_title)

    if not thesis_cat:
        print(f"Investment thesis subcategory not found for '{ticker}' (expected '{thesis_title}').")
        return None

    thesis_id = thesis_cat["categoryId"]
    moat_subcats = get_direct_subcategories(all_categories, thesis_id)

    if not moat_subcats:
        print(f"'{thesis_title}' has no moat threat subcategories.")
        return None

    if debug:
        print(f"[DEBUG] Ticker categoryId={parent_id}")
        print(f"[DEBUG] Investment Thesis categoryId={thesis_id}")
        print(f"[DEBUG] Found {len(moat_subcats)} moat threat subcategories (direct children).")
        print(f"[DEBUG] Author filter: {'ON' if require_author else 'OFF'} ({author_email})")
        print(f"[DEBUG] Include descendants under each moat subcat: {include_descendants}")

    moat_data = {}
    moat_debug = {}  # optional diagnostics

    for moat_cat in moat_subcats:
        moat_cat_id = moat_cat["categoryId"]
        moat_cat_title = moat_cat.get("title", f"category_{moat_cat_id}")

        # Where we will look for topics
        cat_ids = [moat_cat_id]
        if include_descendants:
            cat_ids = _category_tree_ids(all_categories, moat_cat_id)

        # Pull topics across all these categories (dedupe by topicId)
        topics_by_id = {}
        for cid in cat_ids:
            for t in get_topics_for_category(cid):
                tid = t.get("topicId")
                if tid is not None and tid not in topics_by_id:
                    topics_by_id[tid] = t

        topics = list(topics_by_id.values())

        posts_for_subcat = []
        emails_seen = Counter()
        total_posts_seen = 0

        for topic in topics:
            topic_id = topic.get("topicId")
            topic_title = topic.get("title", "")

            if not topic_id:
                continue

            posts = get_posts_for_topic(topic_id)
            total_posts_seen += len(posts)

            for post in posts:
                email = _extract_author_email(post).strip()
                if email:
                    emails_seen[email.lower()] += 1

                if require_author:
                    if email.lower() != (author_email or "").lower():
                        continue

                posts_for_subcat.append({
                    "timestamp": post.get("postTimestamp", 0),
                    "topicId": topic_id,
                    "topicTitle": topic_title,
                    "postId": post.get("postId"),
                    "authorEmail": email,
                    "message": _clean_html_to_text(post.get("message", "")),
                })

        posts_for_subcat.sort(key=lambda p: p.get("timestamp", 0))
        
        moat_data[moat_cat_title] = posts_for_subcat

        if debug:
            print(f"\n[DEBUG] Moat subcat: '{moat_cat_title}' (categoryId={moat_cat_id})")
            print(f"[DEBUG]   Category IDs searched: {len(cat_ids)}")
            print(f"[DEBUG]   Topics found: {len(topics)}")
            print(f"[DEBUG]   Posts fetched (all authors): {total_posts_seen}")
            print(f"[DEBUG]   Posts kept (after filter): {len(posts_for_subcat)}")
            if emails_seen:
                top = emails_seen.most_common(8)
                print(f"[DEBUG]   Top author emails seen (lowercased): {top}")
            else:
                print(f"[DEBUG]   No author emails found in payloads for these posts.")

        moat_debug[moat_cat_title] = {
            "moatCategoryId": moat_cat_id,
            "searchedCategoryIdsCount": len(cat_ids),
            "topicsFound": len(topics),
            "postsFetchedAllAuthors": total_posts_seen,
            "postsKeptAfterFilter": len(posts_for_subcat),
            "topEmailsSeen": emails_seen.most_common(10),
        }

    if not moat_data:
        print(f"No posts found for moat-threat subcategories of '{ticker}'.")
        return None
    
    assembled = {
        "ticker": ticker,
        "category": {"title": parent_cat.get("title"), "categoryId": parent_id},
        "investmentThesisCategory": {"title": thesis_cat.get("title"), "categoryId": thesis_id},
        "moatThreatSubcategories": moat_data,
        "filters": {"authorEmail": author_email if require_author else None, "sortedBy": "timestamp_asc"},
    }

    # Save main output
    os.makedirs("output", exist_ok=True)
    output_path = os.path.join("output", f"{ticker}_moat_threat_source.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(assembled, f, indent=2, ensure_ascii=False)

    # Save debug diagnostics (only if debug)
    if debug:
        dbg_path = os.path.join("output", f"{ticker}_moat_threat_debug.json")
        with open(dbg_path, "w", encoding="utf-8") as f:
            json.dump(moat_debug, f, indent=2, ensure_ascii=False)
        print(f"\n[DEBUG] Saved diagnostics to '{dbg_path}'")

    print(
        f"Saved moat-threat source for '{ticker}' to '{output_path}'. "
        f"Subcategories: {len(moat_subcats)}."
    )
    return assembled


# ---------------------------
# Main pipeline (existing behavior)
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

    # Clean + save output
    simplified_posts = []

    for post in unique_posts.values():
        clean_message = _clean_html_to_text(post.get("message", ""))

        simplified_posts.append({
            "timestamp": post.get("postTimestamp", 0),
            "message": clean_message,
            "authorEmail": (post.get("author") or {}).get("email", "")
        })

    simplified_posts.sort(key=lambda p: p["timestamp"])

    os.makedirs("output", exist_ok=True)
    output_path = os.path.join("output", f"{ticker}_posts.json")

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(simplified_posts, f, indent=2, ensure_ascii=False)

    print(f"Saved {len(simplified_posts)} posts to '{output_path}'.")


# ---------------------------
# CLI (updated)
# ---------------------------

def _print_cli_usage():
    print(
        "Usage:\n"
        "  python forum_posts.py TICKER\n"
        "  python forum_posts.py TICKER --moat\n"
        "  python forum_posts.py TICKER --all\n"
        "\nMoat options:\n"
        "  --debug                 Print detailed counts and save output/{TICKER}_moat_threat_debug.json\n"
        "  --author EMAIL          Override author email filter (default scott@academycapitalmgmt.com)\n"
        "  --no-author-filter      Do not filter by author (show all authors)\n"
        "  --no-descendants        Only look for topics directly under each moat subcategory\n"
    )


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        _print_cli_usage()
        sys.exit(1)

    input_ticker = sys.argv[1]

    run_all = "--all" in sys.argv
    run_moat = "--moat" in sys.argv or run_all
    run_old = (not run_moat) or run_all

    debug = "--debug" in sys.argv
    require_author = "--no-author-filter" not in sys.argv
    include_descendants = "--no-descendants" not in sys.argv

    author_email = SCOTT_EMAIL
    if "--author" in sys.argv:
        try:
            author_email = sys.argv[sys.argv.index("--author") + 1]
        except (ValueError, IndexError):
            print("❌ Missing value after --author")
            _print_cli_usage()
            sys.exit(1)

    if run_old:
        fetch_all_for_ticker(input_ticker)

    if run_moat:
        fetch_moat_threat_source_for_ticker(
            input_ticker,
            author_email=author_email,
            include_descendants=include_descendants,
            require_author=require_author,
            debug=debug,
        )
