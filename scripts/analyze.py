#!/usr/bin/env python3
"""
analyze.py — CLI entry point for Genesys call analysis.

Processes all categories from raw CSV files and produces per-month
consolidated Excel workbooks (Summary + monthly tabs, no charts).

Usage:
    python3 analyze.py [options]

Options:
    --csv-dir PATH     Where CSV files live (default: ~/Documents/Genesys/)
    --out-dir PATH     Where to write Excel files (default: ~/Documents/Genesys/Output/)
    --batch NAME       Only process one batch (e.g. "Switchboard")
    --simple           Use simple XLSX (openpyxl only, no charts)
    --full             Use full XLSX (hand-crafted XML charts)
    --list             List available batches and CSV files
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


def list_csvs(csv_dir):
    """List all CSV files in the CSV directory, grouped by category."""
    if not os.path.isdir(csv_dir):
        print(f"⚠ CSV directory not found: {csv_dir}")
        return {}

    csv_files = [f for f in os.listdir(csv_dir) if f.lower().endswith(".csv")]
    categories = {}

    for fname in sorted(csv_files):
        match = re.search(r"(april|may|june)", fname, re.IGNORECASE)
        if not match:
            print(f"  ⚠ No month found in '{fname}', skipping")
            continue

        month = match.group(1).capitalize()
        prefix = fname[:match.start()].strip()
        category = re.sub(r"^doc_[a-f0-9]+_", "", prefix).strip()
        category = re.sub(r'(?i)switchboa?r[dh]|switchba?o?r?d', 'Switchboard', category)
        category = category.strip('-').strip()

        if category not in categories:
            categories[category] = {}
        categories[category][month] = fname

    return categories


def main():
    parser = argparse.ArgumentParser(
        description="Genesys Call Analysis — process CSVs into Excel reports."
    )
    parser.add_argument("--csv-dir", default=DEFAULT_CSV_DIR,
                        help=f"CSV directory (default: {DEFAULT_CSV_DIR})")
    parser.add_argument("--out-dir", default=DEFAULT_OUT_DIR,
                        help=f"Output directory (default: {DEFAULT_OUT_DIR})")
    parser.add_argument("--batch", default=None,
                        help="Only process one batch (e.g., 'Switchboard')")
    parser.add_argument("--simple", action="store_true",
                        help="Use simple XLSX (openpyxl only, no charts)")
    parser.add_argument("--full", action="store_true",
                        help="Use full XLSX (hand-crafted XML charts)")
    parser.add_argument("--list", action="store_true",
                        help="List available batches and CSV files")

    args = parser.parse_args()

    if args.list:
        categories = list_csvs(args.csv_dir)
        if not categories:
            print("No CSV files found.")
            return
        print(f"Discovered {len(categories)} categories: {', '.join(sorted(categories.keys()))}")
        for cat, months in sorted(categories.items()):
            print(f"\n  {cat}:")
            for month, fname in sorted(months.items()):
                fpath = os.path.join(args.csv_dir, fname)
                size = os.path.getsize(fpath) if os.path.exists(fpath) else 0
                print(f"    {month}: {fname} ({size:,} bytes)")
        return

    if not os.path.isdir(args.csv_dir):
        print(f"Error: CSV directory not found: {args.csv_dir}")
        print(f"Create it and copy your Genesys CSV exports there.")
        sys.exit(1)

    # Import from our modules
    from engineering import BATCHES, analyze_batch
    from office import simple_xlsx, full_xlsx

    os.makedirs(args.out_dir, exist_ok=True)

    # Select batches
    if args.batch:
        batches = {args.batch: BATCHES[args.batch]}
    else:
        batches = BATCHES

    results = {}
    for batch_name, monthly_files in batches.items():
        print(f"\n{'='*60}")
        print(f"Processing: {batch_name}")
        print(f"{'='*60}")

        # Analyze monthly data
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

        # Write Excel output
        out_name = f"{batch_name} - Consolidated.xlsx"
        out_path = os.path.join(args.out_dir, out_name)

        if args.full:
            full_xlsx(batch_name, monthly_data, out_path)
        else:
            simple_xlsx(batch_name, monthly_data, out_path)

        results[batch_name] = "✓"

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
