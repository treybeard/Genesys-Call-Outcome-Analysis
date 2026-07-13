#!/usr/bin/env python3
"""
output.py — Generate consolidated Excel reports from monthly Genesys CSVs.

Takes per-month CSV files, processes them all, and produces:
  - Per-batch consolidated XLSX (Summary + monthly tabs)
  - Optional: global aggregation across all batches
  - Optional: charts (hand-crafted XML) via --full flag

Usage:
    python3 output.py [options]

Options:
    --csv-dir PATH     Where CSV files live (default: ~/Documents/Genesys/)
    --out-dir PATH     Where to write Excel files (default: ~/Documents/Genesys/Output/)
    --full             Use full XLSX (hand-crafted XML charts)
    --summary-only     Only produce Summary sheets (no monthly tabs)
    --merge            Merge all categories into one workbook
    --help             Show this help
"""

import argparse
import os
import sys
import re


# Canonical project directory (scripts live in scripts/, so project root is parent)
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_CSV_DIR = os.path.join(PROJECT_ROOT, "CSV Files")
DEFAULT_OUT_DIR = os.path.join(PROJECT_ROOT, "Output")


def discover_categories(csv_dir):
    """Scan CSV directory and group files by category.

    Returns: { category_name: [(month, filename), ...] }
    """
    if not os.path.isdir(csv_dir):
        return {}

    csv_files = [f for f in os.listdir(csv_dir) if f.lower().endswith(".csv")]
    groups = {}

    for fname in csv_files:
        match = re.search(r"(april|may|june)", fname, re.IGNORECASE)
        if not match:
            continue

        month = match.group(1).capitalize()
        prefix = fname[:match.start()].strip()
        category = re.sub(r"^doc_[a-f0-9]+_", "", prefix).strip()
        # Normalize typos
        category = re.sub(
            r'(?i)switchboa?r[dh]|switchba?o?r?d', 'Switchboard', category
        )
        category = category.strip('-').strip()

        if category not in groups:
            groups[category] = {}
        groups[category][month] = fname

    # Convert to ordered list (April, May, June)
    result = {}
    for cat, months_dict in sorted(groups.items()):
        ordered = []
        for m in ["April", "May", "June"]:
            if m in months_dict:
                ordered.append((m, months_dict[m]))
        if ordered:
            result[cat] = ordered

    return result


def main():
    parser = argparse.ArgumentParser(
        description="Genesys Call Analysis — generate consolidated reports."
    )
    parser.add_argument("--csv-dir", default=DEFAULT_CSV_DIR,
                        help=f"CSV directory (default: {DEFAULT_CSV_DIR})")
    parser.add_argument("--out-dir", default=DEFAULT_OUT_DIR,
                        help=f"Output directory (default: {DEFAULT_OUT_DIR})")
    parser.add_argument("--full", action="store_true",
                        help="Use full XLSX (hand-crafted XML charts)")
    parser.add_argument("--summary-only", action="store_true",
                        help="Only produce Summary sheets")
    parser.add_argument("--merge", action="store_true",
                        help="Merge all categories into one workbook")

    args = parser.parse_args()

    if not os.path.isdir(args.csv_dir):
        print(f"Error: CSV directory not found: {args.csv_dir}")
        sys.exit(1)

    from engineering import analyze_batch
    from office import simple_xlsx, full_xlsx

    os.makedirs(args.out_dir, exist_ok=True)

    # Discover categories
    categories = discover_categories(args.csv_dir)
    if not categories:
        print("No CSV files found.")
        return

    print(f"Discovered {len(categories)} categories: {', '.join(sorted(categories.keys()))}")

    results = {}
    all_monthly_data = {}  # For merge mode

    for batch_name, monthly_files in categories.items():
        print(f"\n{'='*60}")
        print(f"Processing: {batch_name}")
        print(f"{'='*60}")

        monthly_data = analyze_batch(batch_name, monthly_files, args.csv_dir)
        if not monthly_data:
            print(f"  ❌ No valid data found for {batch_name}")
            results[batch_name] = "✗"
            continue

        # Print summary
        print(f"  Months analyzed: {', '.join(monthly_data.keys())}")
        total_all = sum(d["total_calls"] for d in monthly_data.values())
        total_success = sum(d["successful"] for d in monthly_data.values())
        total_failed = sum(d["failed"] for d in monthly_data.values())
        total_abandoned = sum(d["abandoned"] for d in monthly_data.values())
        print(f"  Total Calls: {total_all}")
        print(f"  Successful: {total_success} ({total_success/total_all*100:.1f}%)")
        print(f"  Failed: {total_failed}")
        print(f"  Abandoned: {total_abandoned} ({total_abandoned/total_all*100:.1f}%)")

        out_name = f"{batch_name} - Consolidated.xlsx"
        out_path = os.path.join(args.out_dir, out_name)

        if args.full:
            full_xlsx(batch_name, monthly_data, out_path)
        else:
            simple_xlsx(batch_name, monthly_data, out_path)

        results[batch_name] = "✓"
        all_monthly_data[batch_name] = monthly_data

    # Merge mode: combine all categories into one workbook
    if args.merge and all_monthly_data:
        merged_out = os.path.join(args.out_dir, "All Categories - Consolidated.xlsx")
        print(f"\n{'='*60}")
        print("Merging all categories...")
        print(f"{'='*60}")
        # For merge, we'd write each batch's data into one workbook
        # This is a placeholder for the merge functionality
        print(f"  Merge mode: producing individual files instead")
        print(f"  To merge, run individual batch outputs to the same directory")

    # Print final summary
    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")
    for cat, status in results.items():
        print(f"  {status} {cat}")

    success_count = sum(1 for s in results.values() if s == "✓")
    print(f"\n✅ {success_count}/{len(results)} batches processed")
    print(f"📁 Output directory: {args.out_dir}")


if __name__ == "__main__":
    main()
