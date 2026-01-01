import os
import json
import asyncio
import re
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
from dotenv import load_dotenv

# Import the provider factory
from .providers import get_provider

# Load environment variables
load_dotenv()

DEFAULT_EMPHASIS_EMAIL = os.getenv("SUMMARY_EMPHASIS_EMAIL", "smgacm@gmail.com")


# ---------------------------
# Internal helpers (Preserved Authoritative Logic)
# ---------------------------

def _strip_urls(text: str) -> str:
    # Remove URLs which often trigger safety and add noise
    return re.sub(r"https?://\S+", "[link]", text or "")

def _sanitize_message(text: str) -> str:
    text = text or ""
    text = _strip_urls(text)
    # Collapse whitespace
    text = re.sub(r"\s+", " ", text).strip()
    return text

def _is_blocked_response(text: str) -> bool:
    if not text:
        return True
    t = text.strip().lower()
    return (
        "blocked" in t
        or "response was empty" in t
        or t.startswith("error:")
        or "summary generation failed" in t
    )

def _safe_dt(ts: int) -> str:
    try:
        return str(datetime.fromtimestamp(int(ts)))
    except Exception:
        return ""


def _build_content_lines(posts: List[Dict[str, Any]], newest_first: bool = True) -> List[str]:
    sorted_posts = sorted(posts, key=lambda x: x.get("timestamp", 0), reverse=newest_first)
    lines = []
    for post in sorted_posts:
        ts = post.get("timestamp", 0) or 0
        msg = _sanitize_message(post.get("message", ""))
        dt = _safe_dt(ts)
        if dt:
            lines.append(f"Date: {dt}\nMessage:\n{msg}\n\n")
        else:
            lines.append(f"Message:\n{msg}\n\n")
    return lines

def _trim_to_char_budget(lines: List[str], max_chars: int) -> Tuple[str, int]:
    """
    Join lines until max_chars reached. Returns (text, used_chars).
    """
    out = []
    used = 0
    for line in lines:
        if used + len(line) > max_chars:
            break
        out.append(line)
        used += len(line)
    return "".join(out), used

def _extract_numbered_bullets(summary_text: str) -> str:
    """
    For the 10-bullet ticker summary, strip any preamble before '1. **'
    """
    match = re.search(r'1\.\s+\*\*', summary_text)
    if match:
        return summary_text[match.start():].strip()
    
    # Fallback for models that might miss the bolding
    match_lax = re.search(r'1\.\s+', summary_text)
    if match_lax:
         return summary_text[match_lax.start():].strip()

    print("[SUMMARIZER] ==> WARNING: '1. **' pattern not found. Returning raw response.")
    return summary_text.strip()


# ---------------------------
# Post Summarization (Async & Sync)
# ---------------------------

async def _generate_post_summary_async(posts: List[Dict[str, Any]], ticker: str) -> str:
    """Async helper function to generate a 10-bullet summary using the configured Provider."""
    content_lines = _build_content_lines(posts, newest_first=True)
    
    # Instantiate the provider dynamically based on env
    llm = get_provider()

    system_prompt = f"""You are an expert financial analyst. Your task is to summarize the key insights about the given ticker in exactly 10 concise bullet points.

Requirements:
- Each bullet point must start with a bold topic phrase followed by a colon (e.g., **Topic:** Detail).
- Focus on concrete facts, financial metrics, and in-depth analysis.
- Give more weight to the most recent information and any content from '{DEFAULT_EMPHASIS_EMAIL}'.
- Do not mention the sources, dates, or email addresses in your summary.
- Present all information as direct market insights.
- Number each bullet point from 1 to 10."""

    user_prompt = f"Here are the posts about {ticker} for you to analyze (newest first):\n\n" + "".join(content_lines)

    summary_text = await llm.generate_content_async(
        system_prompt=system_prompt,
        user_prompt=user_prompt,
        temperature=0.5,
    )

    return _extract_numbered_bullets(summary_text)


def generate_post_summary(posts: list, ticker: str) -> str:
    """
    The main interface function that calls the async helper and returns the result.
    """
    print(f"[SUMMARIZER] ==> Summarizer initiated for {ticker}.")
    if not posts:
        return "No posts were provided to summarize."

    try:
        summary = asyncio.run(_generate_post_summary_async(posts, ticker))
        if summary:
            print(f"[SUMMARIZER] ==> Summary generated successfully ({len(summary)} chars).")
            return summary
        print("[SUMMARIZER] ==> ERROR: The summarizer returned an empty string.")
        return "Error: Summary was empty."
    except Exception as e:
        import traceback
        print(f"[SUMMARIZER] ==> An exception occurred: {e}")
        traceback.print_exc()
        return "Error: Could not generate summary due to an exception."


# ---------------------------
# Moat Threat Summarization
# ---------------------------

async def _generate_moat_subcategory_summary_async(
    posts: List[Dict[str, Any]],
    ticker: str,
    moat_subcategory: str,
    *,
    emphasis_email: str = DEFAULT_EMPHASIS_EMAIL,
    temperature: float = 0.3,
    max_input_chars: int = 80_000,
    max_output_tokens: int = 8192,
) -> str:
    if not posts:
        return "No posts available for this moat threat subcategory."

    llm = get_provider()

    # Attempt 1: normal prompt, newest-first
    content_lines = _build_content_lines(posts, newest_first=True)
    trimmed_blob, used_chars = _trim_to_char_budget(content_lines, max_input_chars)

    system_prompt_1 = f"""You are an expert financial analyst.

Write exactly ONE concise paragraph summarizing the moat threat theme for {ticker} under "{moat_subcategory}".

Requirements:
- Exactly 1 paragraph (no bullets, no headings).
- 4–6 sentences, specific and analytical.
- Weight the most recent messages more heavily.
- Do not mention sources, dates, email addresses, or that these were posts.
- Do not include URLs or direct quotes."""
    
    user_prompt_1 = (
        f"Ticker: {ticker}\n"
        f"Theme: {moat_subcategory}\n\n"
        f"Messages (newest first):\n\n{trimmed_blob}"
    )

    text = await llm.generate_content_async(
        system_prompt=system_prompt_1,
        user_prompt=user_prompt_1,
        temperature=temperature,
        max_output_tokens=max_output_tokens,
    )

    if not _is_blocked_response(text):
        return re.sub(r"\s+", " ", text.strip())

    # Attempt 2: safer / more neutral instructions + smaller input
    print(f"[SUMMARIZER] ==> Blocked/Empty for '{moat_subcategory}'. Retrying with safer prompt + smaller input.")
    smaller_blob, _ = _trim_to_char_budget(content_lines, max_input_chars // 2)

    system_prompt_2 = f"""Summarize the business and competitive risks described for {ticker} related to "{moat_subcategory}".

Output rules:
- Exactly 1 paragraph.
- 4–6 sentences.
- Neutral, factual tone.
- No politics commentary, no sensitive labels, no quotes, no URLs, no personal opinions.
- Weight the most recent messages more heavily."""
    
    user_prompt_2 = f"{smaller_blob}"

    text2 = await llm.generate_content_async(
        system_prompt=system_prompt_2,
        user_prompt=user_prompt_2,
        temperature=0.2, # Lower temp for retry
        max_output_tokens=max_output_tokens,
    )

    if not _is_blocked_response(text2):
        return re.sub(r"\s+", " ", text2.strip())

    return "__BLOCKED_BY_PROVIDER__"


async def generate_moat_threat_summary_async(
    ticker: str,
    *,
    source_path: Optional[str] = None,
    output_path: Optional[str] = None,
    emphasis_email: str = DEFAULT_EMPHASIS_EMAIL,
    temperature: float = 0.4,
    concurrency: int = 4,
    max_input_chars: int = 120_000,
    max_output_tokens: int = 8192,
) -> Dict[str, Any]:
    """
    Reads output/{ticker}_moat_threat_source.json and writes output/{ticker}_moat_threat_summary.json
    """
    source_path = source_path or os.path.join("output", f"{ticker}_moat_threat_source.json")
    output_path = output_path or os.path.join("output", f"{ticker}_moat_threat_summary.json")

    if not os.path.exists(source_path):
        raise FileNotFoundError(f"Moat threat source file not found: {source_path}")

    with open(source_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    moat_map = data.get("moatThreatSubcategories") or {}
    if not isinstance(moat_map, dict) or not moat_map:
        result = {"ticker": ticker, "moatThreatSummaries": {}}
        os.makedirs("output", exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as out:
            json.dump(result, out, indent=2, ensure_ascii=False)
        return result

    # Limit concurrency for API safety
    sem = asyncio.Semaphore(max(1, int(concurrency)))

    async def _summarize_one(name: str, posts: List[Dict[str, Any]]) -> Tuple[str, str]:
        async with sem:
            summary = await _generate_moat_subcategory_summary_async(
                posts=posts,
                ticker=ticker,
                moat_subcategory=name,
                emphasis_email=emphasis_email,
                temperature=temperature,
                max_input_chars=max_input_chars,
                max_output_tokens=max_output_tokens,
            )
            return name, summary

    tasks = [_summarize_one(name, posts) for name, posts in moat_map.items()]
    pairs = await asyncio.gather(*tasks)

    summaries = {name: summary for name, summary in pairs}
    result = {"ticker": ticker, "moatThreatSummaries": summaries}

    os.makedirs("output", exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as out:
        json.dump(result, out, indent=2, ensure_ascii=False)

    print(f"[SUMMARIZER] ==> Moat threat summary saved to {output_path}")
    return result


def generate_moat_threat_summary(
    ticker: str,
    *,
    source_path: Optional[str] = None,
    output_path: Optional[str] = None,
    emphasis_email: str = DEFAULT_EMPHASIS_EMAIL,
    temperature: float = 0.4,
    concurrency: int = 4,
    max_input_chars: int = 120_000,
    max_output_tokens: int = 8192,
) -> Dict[str, Any]:
    """
    Sync wrapper to generate moat threat summaries.
    """
    try:
        return asyncio.run(
            generate_moat_threat_summary_async(
                ticker=ticker,
                source_path=source_path,
                output_path=output_path,
                emphasis_email=emphasis_email,
                temperature=temperature,
                concurrency=concurrency,
                max_input_chars=max_input_chars,
                max_output_tokens=max_output_tokens,
            )
        )
    except Exception as e:
        import traceback
        print(f"[SUMMARIZER] ==> Error generating moat threat summary for {ticker}: {e}")
        traceback.print_exc()
        return {"ticker": ticker, "moatThreatSummaries": {}, "error": str(e)}


# --- Standalone Execution Logic (for testing/CLI) ---

def process_ticker_posts(ticker: str, debug: bool = False):
    posts_filename = os.path.join("output", f"{ticker}_posts.json")
    emails_filename = os.path.join("output", f"{ticker}_sent_emails.json")

    try:
        combined_data = []
        if os.path.exists(posts_filename):
            with open(posts_filename, "r", encoding="utf-8") as f:
                combined_data.extend(json.load(f))
        if os.path.exists(emails_filename):
            with open(emails_filename, "r", encoding="utf-8") as f:
                combined_data.extend(json.load(f))

        if combined_data:
            summary = generate_post_summary(combined_data, ticker)
            output_filename = os.path.join("output", f"{ticker}_post_summary.txt")
            with open(output_filename, "w", encoding="utf-8") as f:
                f.write(summary)
            print(f"Summary saved to {output_filename}")
        else:
            print(f"No data for {ticker}.")

    except Exception as e:
        print(f"Error processing data for {ticker}: {e}")

def process_ticker_moat_threat(ticker: str):
    """
    Standalone runner for moat-threat summaries.
    Assumes output/{ticker}_moat_threat_source.json already exists.
    """
    generate_moat_threat_summary(ticker)

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Generate summaries using Configured Provider.")
    parser.add_argument("ticker", type=str, help="Stock ticker symbol (e.g., AAPL)")
    parser.add_argument("--debug", action="store_true", help="Enable debug mode")
    parser.add_argument("--moat", action="store_true", help="Generate moat threat summaries")
    args = parser.parse_args()

    t = args.ticker.upper()
    if args.moat:
        process_ticker_moat_threat(t)
    else:
        process_ticker_posts(t, args.debug)