#!/usr/bin/env python3
import os
import sys
import json
import datetime
from typing import Dict, Any, List

# Third-party libraries
import requests
import finnhub
from dotenv import load_dotenv

import logging
logger = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s"
)

# Assume this function returns "MM-DD" for a given ticker or CIK.
# If you don't have it, stub it with a default "12-31".
from acm_analysis import get_fiscal_year_end

# -----------------------------------------------------------------------------
# PART 1: Fetch revenue breakdown from Finnhub
# -----------------------------------------------------------------------------

def get_revenue_breakdown(symbol_or_cik: str) -> Dict[str, Any]:
    """
    Fetch revenue breakdown data for a given symbol or CIK using Finnhub's API.
    """
    load_dotenv()
    api_key = os.getenv("FINNHUB_API_KEY")
    if not api_key:
        raise ValueError("Finnhub API key not found. Set FINNHUB_API_KEY in your .env file.")

    finnhub_client = finnhub.Client(api_key=api_key)
    data = finnhub_client.stock_revenue_breakdown(symbol_or_cik)
    return data

# -----------------------------------------------------------------------------
# PART 2: Date parsing and FY logic
# -----------------------------------------------------------------------------

def parse_date(date_str: str) -> datetime.date:
    """Safely parse a date string (YYYY-MM-DD) into a datetime.date object."""
    return datetime.datetime.strptime(date_str, "%Y-%m-%d").date()

def determine_fiscal_year(end_date: datetime.date, fy_end_mm_dd: str) -> int:
    """
    Given the end_date of a reporting period and the company's fiscal year-end (MM-DD),
    figure out which *fiscal year* this end_date belongs to.
    """
    mm, dd = map(int, fy_end_mm_dd.split("-"))
    this_year_fy_end = datetime.date(end_date.year, mm, dd)
    if end_date <= this_year_fy_end:
        return end_date.year
    else:
        return end_date.year + 1

def is_approx_full_year(
    start_date: datetime.date,
    end_date: datetime.date,
    min_days: int = 360,
    max_days: int = 370
) -> bool:
    """
    Return True if the period is between min_days and max_days (inclusive),
    to allow for typical 52–53 week variations, leap years, etc.
    """
    diff = (end_date - start_date).days + 1
    return min_days <= diff <= max_days

# -----------------------------------------------------------------------------
# PART 3: Filtering & Consolidation
# -----------------------------------------------------------------------------

def filter_breakdown_fields(breakdown: Dict[str, Any]) -> None:
    """
    From 'breakdown', remove the top-level 'unit' and 'concept' fields if present.
    """
    breakdown.pop("unit", None)
    breakdown.pop("concept", None)

def remove_unwanted_fields_in_data(revenue_breakdown_list: List[Dict[str, Any]]) -> None:
    """
    Within the 'data' of each axis entry, remove 'unit' and 'member'.
    """
    for axis_entry in revenue_breakdown_list:
        # axis_entry["data"] is a list of objects
        data_list = axis_entry.get("data", [])
        for d in data_list:
            d.pop("unit", None)
            d.pop("member", None)
            d.pop("percentage", None)

def deduplicate_revenue_breakdown_entries(entries: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    De-duplicate identical axis+data sets. We'll:
      - skip duplicates if (axis_name, sorted(data_tuples)) is the same.
    Return a new list with unique entries in order of first occurrence.
    """
    seen = set()
    unique_entries = []

    for entry in entries:
        axis_name = entry.get("axis", "")
        data_list = entry.get("data", [])
        # Sort data by (label, value, percentage) to get a canonical representation
        # ignoring potential small floating differences. If you want fuzzy matching,
        # you'd have to approach differently.
        sorted_data = sorted(
            data_list, 
            key=lambda d: (d.get("label", ""), d.get("value", 0.0))
        )
        data_tuple = tuple(
            (d.get("label", ""), d.get("value", 0.0))
            for d in sorted_data
        )
        signature = (axis_name, data_tuple)
        if signature not in seen:
            seen.add(signature)
            unique_entries.append(entry)

    return unique_entries

def is_strict_subset(smaller_data: List[Dict[str, Any]], bigger_data: List[Dict[str, Any]]) -> bool:
    """
    Return True if smaller_data is a strict subset of bigger_data.
    Specifically, for each (label, value) in smaller_data, 
    it must appear exactly in bigger_data. The bigger_data may have more items.
    
    We do NOT compare 'percentage' for subset logic because 
    the 'percentage' can be slightly off. If needed, adjust logic.
    """
    # Convert bigger_data to a set of (label, value) pairs:
    bigger_set = {(d["label"], d["value"]) for d in bigger_data}
    smaller_set = {(d["label"], d["value"]) for d in smaller_data}

    if smaller_set == bigger_set:
        return False  # They are the same, not a *strict* subset
    return smaller_set.issubset(bigger_set)

def remove_subset_entries(entries: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    If we have two entries with the same 'axis', and the data of one is a 
    strict subset of the data of the other, remove the smaller.
    We'll do an O(n^2) pass. For each pair that have the same axis:
        if A is a strict subset of B => remove A
    Priority is preserving order of first occurrence. 
    We'll remove subsets from the final result.
    """
    to_remove_indices = set()
    n = len(entries)

    for i in range(n):
        if i in to_remove_indices:
            continue
        axis_i = entries[i].get("axis", "")
        data_i = entries[i].get("data", [])

        for j in range(i+1, n):
            if j in to_remove_indices:
                continue
            axis_j = entries[j].get("axis", "")
            data_j = entries[j].get("data", [])

            # Only compare if axis is the same
            if axis_i == axis_j:
                # Check if i is subset of j
                if is_strict_subset(data_i, data_j):
                    to_remove_indices.add(i)
                    break  # i is removed, move on
                # Or if j is subset of i
                elif is_strict_subset(data_j, data_i):
                    to_remove_indices.add(j)
                else:
                    # Not subsets
                    pass

    return [
        e for idx, e in enumerate(entries)
        if idx not in to_remove_indices
    ]

def remove_single_data_point(year_item: Dict[str, Any]) -> None:
    """
    For a given 'year_item' structure, remove any revenueBreakdown entry that
    has exactly one data point, where that single 'value' equals the top-level
    'breakdown["value"]'.

    Modifies 'year_item' in place by filtering out those entries.
    """
    breakdown = year_item.get("breakdown", {})

    original_list = breakdown.get("revenueBreakdown", [])
    filtered_list = []
    for rbe in original_list:
        data_list = rbe.get("data", [])
        # If there's exactly one data point...
        if len(data_list) == 1:
            # skip it
            continue
        filtered_list.append(rbe)

    breakdown["revenueBreakdown"] = filtered_list

def filter_revenue_breakdown(breakdown: Dict[str, Any]) -> bool:
    """
    1) Remove axis == "srt_StatementGeographicalAxis".
    2) Remove empty data entries (where "data" is empty or missing).
    3) Remove 'unit'/'concept' from top-level breakdown.
    4) Remove 'unit'/'member' from each 'data' item.

    Return True if there's still meaningful data left; else False.
    """
    # Step 3: remove 'unit'/'concept' at top level
    filter_breakdown_fields(breakdown)

    revenue_breakdown = breakdown.get("revenueBreakdown")
    if not isinstance(revenue_breakdown, list):
        return False

    filtered_entries = []
    for axis_entry in revenue_breakdown:
        axis_name = axis_entry.get("axis", "")
        data_list = axis_entry.get("data", [])

        # Step A: detect any duplicate labels in data_list
        labels_seen = set()
        has_duplicate_labels = False
        for item in data_list:
            lbl = item.get("label")
            if lbl in labels_seen:
                has_duplicate_labels = True
                break
            labels_seen.add(lbl)

        # If we have duplicate labels, skip this entire axis entry
        if has_duplicate_labels:
            continue

        # Skip if axis == "srt_StatementGeographicalAxis"
        if "StatementGeographicalAxis" in axis_name:
            continue

        if "DerivativeInstrumentRiskAxis" in axis_name:
            continue
        
        if "AdjustmentsForNewAccountingPronouncementsAxis" in axis_name:
            continue

        # Skip if data is empty
        if not data_list:
            continue

        filtered_entries.append(axis_entry)

    # Now remove 'unit'/'member' fields from the sub-entries we kept
    remove_unwanted_fields_in_data(filtered_entries)

    # 5) De-duplicate
    filtered_entries = deduplicate_revenue_breakdown_entries(filtered_entries)

     # Step 6: remove subsets
    filtered_entries = remove_subset_entries(filtered_entries)

    breakdown["revenueBreakdown"] = filtered_entries
    # NEW: Now remove any single-data-point entry matching top-level:
    temp_item = {"breakdown": breakdown}  # wrap in a year-like dict
    remove_single_data_point(temp_item)
    # The breakdown is now updated in-place.

    # We can re-check how many remain:
    final_list = breakdown.get("revenueBreakdown", [])
    return len(final_list) > 0

def consolidate_by_fiscal_year(raw_data: Dict[str, Any], fy_end_mm_dd: str) -> Dict[str, List[Any]]:
    """
    Consolidate 'raw_data' by fiscal year, skipping partial-year items and
    skipping any item that has no meaningful data after filtering.
    """
    consolidated = {}
    # 'raw_data' from Finnhub has a top-level "data" list
    for item in raw_data.get("data", []):
        # Remove 'accessNumber'
        item.pop("accessNumber", None)

        breakdown = item.get("breakdown", {})
        start_date_str = breakdown.get("startDate")
        end_date_str   = breakdown.get("endDate")
        if not (start_date_str and end_date_str):
            continue

        # Filter the breakdown in place
        if not filter_revenue_breakdown(breakdown):
            # Means there's no meaningful data left
            continue

        # Now parse the dates
        try:
            start_date = parse_date(start_date_str)
            end_date   = parse_date(end_date_str)
        except ValueError:
            # Invalid date format => skip
            continue

        # 1) Skip partial-year items
        if not is_approx_full_year(start_date, end_date):
            continue

        # 2) Determine the fiscal year
        fy = determine_fiscal_year(end_date, fy_end_mm_dd)
        fy_str = str(fy)
        if fy_str not in consolidated:
            consolidated[fy_str] = []
        consolidated[fy_str].append(item)

    return consolidated

# -----------------------------------------------------------------------------
# PART 4: Main routine
# -----------------------------------------------------------------------------

def main():
    """
    Usage:
        python this_script.py <ticker_or_cik>
    Example:
        python this_script.py MSFT
        python this_script.py 320193   # Apple CIK
    """
    if len(sys.argv) < 2:
        logger.error("Usage: python this_script.py <ticker_or_cik>")
        sys.exit(1)

    symbol_or_cik = sys.argv[1].upper()
    store_raw = "--store_raw" in sys.argv
    logger.info(f"Fetching revenue breakdown for: {symbol_or_cik}")

    # 1) Fetch raw data
    raw_data = get_revenue_breakdown(symbol_or_cik)
    if not raw_data:
        logger.error(f"Finnhub returned no data for symbol/CIK: {symbol_or_cik}")
        sys.exit(1)

    # If --store_raw is present, save the raw data to <ticker>_revenue_breakdown.json
    if store_raw:
        raw_output_path = f"{symbol_or_cik}_revenue_breakdown.json"
        with open(raw_output_path, "w", encoding="utf-8") as raw_file:
            json.dump(raw_data, raw_file, indent=2)
        logger.info(f"Raw revenue breakdown data saved to {raw_output_path}")

    # 2) Determine the company's fiscal year-end
    fy_end = get_fiscal_year_end(symbol_or_cik)
    if not fy_end:
        logger.warning(f"No fiscalYearEnd found for {symbol_or_cik}. Defaulting to 12-31.")
        fy_end = "12-31"

    # 3) Consolidate by FY
    consolidated = consolidate_by_fiscal_year(raw_data, fy_end)

    # 4) Output
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    outpath = os.path.join(output_dir, f"{symbol_or_cik}_seg_consolidated.json")

    with open(outpath, "w", encoding="utf-8") as f_out:
        json.dump(consolidated, f_out, indent=2)

    logger.info(
        f"Consolidated segment data for '{symbol_or_cik}' (full-year items only) => {outpath}"
    )

if __name__ == "__main__":
    main()
