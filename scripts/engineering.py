#!/usr/bin/env python3
"""
engineering.py — Pure data analysis for Genesys call flow exports.

Input:  CSV from Genesys Cloud (outcomes, durations, queue paths).
Output: Structured dicts with metrics + call paths.

No I/O side-effects, no Excel. Just analysis.
"""

import csv
from datetime import timedelta
from collections import Counter


# Batch definitions — where the CSV files live and which belong together.
# Source file names as stored in Genesys cache (doc_<uuid>_<original>.csv).
BATCHES = {
    "Switchboard": [
        ("April",  "doc_b696083cdb96_Switchbaord-April.csv"),
        ("May",    "doc_91b7d70ec942_Switchbaord-May.csv"),
        ("June",   "doc_73661aaf8493_Switchboard June.csv"),
    ],
    "Non-ATTY": [
        ("April", "doc_d98d6ef6429d_Non-ATTY April.csv"),
        ("May",   "doc_65dfd2d8d7f7_Non-ATTY May.csv"),
        ("June",  "doc_14b5e25c2c25_Non-ATTY June.csv"),
    ],
    "ATTY": [
        ("April", "doc_d0cec11475c8_ATTY April.csv"),
        ("May",   "doc_38dfa13d8eb9_ATTY May.csv"),
        ("June",  "doc_74e69e181d95_ATTY June.csv"),
    ],
    "Bar Examiners": [
        ("April", "doc_c848269b044c_Bar Examiners April.csv"),
        ("May",   "doc_7ec1019bfcbe_Bar Examiners May.csv"),
        ("June",  "doc_a7d02a33e3f6_Bar Examiners June.csv"),
    ],
    "LRS Spanish": [
        ("April", "doc_87985000e90b_LRS Spanish April.csv"),
        ("May",   "doc_66820355f921_LRS Spanish May.csv"),
        ("June",  "doc_edceaaca07cc_LRS Spanish June.csv"),
    ],
}

# LRS Spanish uses a different outcome classification (both dimensions
# tracked independently: a call can be "successful" AND "abandoned").
LRS_SPANISH_BATES = {"LRS Spanish"}


def parse_duration(raw):
    """Parse duration string like ' 00:04:12.710' → total seconds (float).

    Returns None for malformed values or outliers > 24 hours (data error).
    """
    if not raw or not isinstance(raw, str):
        return None
    raw = raw.strip()
    if not raw:
        return None
    parts = raw.split(":")
    if len(parts) != 3:
        return None
    try:
        h, m, s = parts[0].strip(), parts[1].strip(), parts[2].strip()
        total = int(h) * 3600 + int(m) * 60 + float(s)
        if total > 86400:  # outlier filter
            return None
        return total
    except (ValueError, IndexError):
        return None


def classify_outcome(success_str, abandoned_str):
    """Unified outcome classification for ALL flow types.

    The same formula applies to Switchboard, Non-ATTY, ATTY,
    Bar Examiners, and LRS Spanish:

      Successful = Outcome Success == 1 AND Abandoned != 'YES'
      Failed     = Outcome Success != 1 AND Abandoned != 'YES'
      Abandoned  = Abandoned == 'YES' (regardless of Outcome Success)

    PITFALL: Failed ≠ Total - Successful - Abandoned.
    Some rows appear in both Successful AND Abandoned (users who
    completed their call but hung up before wrap-up).
    """
    outcome_success = success_str.strip() == "1"
    abandoned_flag = abandoned_str.strip() == "YES"

    if outcome_success and not abandoned_flag:
        return "successful"
    elif abandoned_flag:
        return "abandoned"
    else:
        return "failed"


def parse_queue_path(raw):
    """Parse semicolon-separated Queue column into a chain string.

    e.g., 'SWITCHBOARD; LRS' → 'SWITCHBOARD → LRS'
    Returns None if empty (voicemail, callback, etc.).
    """
    if not raw or not str(raw).strip():
        return None
    parts = [p.strip() for p in str(raw).split(";") if p.strip()]
    return " → ".join(parts) if parts else None


def analyze_csv(csv_path, batches=None):
    """Analyze a single Genesys CSV file.

    Args:
        csv_path: Path to the CSV file.
        batches: Optional dict mapping batch_name → list of (month, filename).
                 If provided, returns per-batch aggregations across months.

    Returns:
        If batches is None:
            dict with keys:
                - total_calls: int
                - successful: int
                - failed: int
                - abandoned: int
                - success_pct: float (0–100)
                - abandon_pct: float (0–100)
                - avg_duration: float (seconds)
                - avg_duration_str: str (timedelta string)
                - call_paths: list of (path, count) sorted by count desc
                - path_counter: Counter

        If batches is provided (dict of batch → [(month, fname)]):
            dict of batch_name → analysis dict (above).
            Only batches that have valid files are returned.
    """
    with open(csv_path, "r", encoding="utf-8", errors="replace") as f:
        reader = csv.DictReader(f)
        total_calls = 0
        successful = 0
        failed = 0
        abandoned = 0
        durations = []
        path_counter = Counter()

        for row in reader:
            total_calls += 1

            # Outcome classification (unified for all flow types)
            success_str = row.get("Outcome Success", "")
            abandoned_str = row.get("Abandoned", "")
            category = classify_outcome(success_str, abandoned_str)

            if category == "successful":
                successful += 1
            elif category == "abandoned":
                abandoned += 1
            else:
                failed += 1

            # Duration (trimmed, outlier-filtered)
            dur_raw = row.get("Duration", "")
            dur_secs = parse_duration(dur_raw)
            if dur_secs is not None:
                durations.append(dur_secs)

            # Queue paths (semicolon-separated → arrow-joined)
            queue_raw = row.get("Queue", "")
            path = parse_queue_path(queue_raw)
            if path:
                path_counter[path] += 1

    avg_duration = sum(durations) / len(durations) if durations else 0
    call_paths = sorted(path_counter.items(), key=lambda x: x[1], reverse=True)

    return {
        "total_calls": total_calls,
        "successful": successful,
        "failed": failed,
        "abandoned": abandoned,
        "success_pct": (successful / total_calls * 100) if total_calls else 0,
        "abandon_pct": (abandoned / total_calls * 100) if total_calls else 0,
        "avg_duration": avg_duration,
        "avg_duration_str": str(timedelta(seconds=round(avg_duration)))
            if avg_duration > 0 else "0:00:00",
        "call_paths": call_paths,
        "path_counter": path_counter,
    }


def analyze_batch(batch_name, monthly_files, csv_dir):
    """Analyze all months for one batch.

    Args:
        batch_name: e.g. 'Switchboard'
        monthly_files: list of (month, filename) tuples.
        csv_dir: directory where CSV files live.

    Returns:
        dict {month_name: analysis_dict}, only for months with valid files.
    """
    results = {}
    for month_name, fname in monthly_files:
        csv_full = f"{csv_dir}/{fname}"
        try:
            data = analyze_csv(csv_full)
            data["filename"] = fname  # keep traceability
            results[month_name] = data
        except (FileNotFoundError, KeyError, ValueError) as e:
            print(f"  ⚠ {batch_name}/{month_name}: {e}")
    return results


def weight_avg(values, weights):
    """Compute weighted average. Returns 0 if no weights."""
    total_weight = sum(weights)
    if total_weight == 0:
        return 0
    return sum(v * w for v, w in zip(values, weights)) / total_weight
