#!/usr/bin/env python3
"""analyse_calls.py

Read the Genesys Cloud CSV export, compute metrics and write an Excel summary.
"""

import sys
import os
import pandas as pd
import re

def _parse_duration(series: pd.Series) -> pd.Series:
    """
    Convert a Series that contains strings like "4m 13s" into total seconds.
    Returns a new Series of integers (seconds).
    """
    # minutes part (optional spaces)
    mins = series.str.extract(r'(\d+)\s*m', expand=False).fillna(0).astype(int)
    # seconds part (optional spaces)
    secs = series.str.extract(r'(\d+)\s*s', expand=False).fillna(0).astype(int)
    return mins * 60 + secs

def main(csv_path: str):
    # --------------------------------------------------------------
    # 1️⃣ Load the CSV – keep everything as strings so we can search freely
    # --------------------------------------------------------------
    try:
        df = pd.read_csv(csv_path, dtype=str)
    except FileNotFoundError:
        raise FileNotFoundError(f"CSV file not found: {csv_path}")
    except Exception as e:
        raise RuntimeError(f"Failed to read CSV '{csv_path}': {e}")

    # --------------------------------------------------------------
    # 2️⃣ Clean out completely empty columns
    # --------------------------------------------------------------
    df = df.dropna(axis=1, how='all')

    total_calls = len(df)

    # --------------------------------------------------------------
    # 3️⃣ Validate required columns exist
    # --------------------------------------------------------------
    direction_col = next((col for col in df.columns if col.lower() == 'direction'), None)
    duration_col = next((col for col in df.columns if col.lower() == 'duration'), None)
    queue_col = next((col for col in df.columns if col.lower() == 'queue'), None)

    if direction_col is None:
        raise KeyError("Column named 'Direction' not found – check the CSV header.")
    if duration_col is None:
        raise KeyError("Column named 'Duration' not found – check the CSV header.")
    if queue_col is None:
        raise KeyError("Column named 'Queue' not found – check the CSV header.")

    # --------------------------------------------------------------
    # 4️⃣ Inbound vs Outbound split (using Direction)
    # --------------------------------------------------------------
    direction_series = df[direction_col].astype(str).replace('', '')
    inbound_calls = direction_series.str.contains('inbound', case=False, na=False).sum()
    outbound_calls = total_calls - inbound_calls

    # --------------------------------------------------------------
    # 5️⃣ Average call duration (minutes)
    # --------------------------------------------------------------
    duration_series = df[duration_col].astype(str).replace('', pd.NA)
    dur_seconds = _parse_duration(duration_series)
    avg_duration_minutes = dur_seconds.dropna().mean() if not dur_seconds.dropna().empty else 0.0

    # --------------------------------------------------------------
    # 6️⃣ Count calls per unique queue (call‑path)
    # --------------------------------------------------------------
    queue_counts = df[queue_col].value_counts(dropna=True)

    # --------------------------------------------------------------
    # 7️⃣ Overall success rate (YES / NO)
    # --------------------------------------------------------------
    # Look for a column whose header contains "success" (case‑insensitive)
    success_col = None
    for col in df.columns:
        if 'success' in col.lower():
            success_col = col
            break
    # Fallback: any column that mentions yes/no
    if success_col is None:
        for col in df.columns:
            if any(word in col.lower() for word in ['yes', 'no']):
                success_col = col
                break

    if success_col is None:
        raise ValueError("Could not locate a Success/YES‑No column to compute success rate.")

    # Extract YES/NO values, ignore missing entries
    success_series = df[success_col].astype(str).str.upper().replace('', pd.NA)
    yes_count = success_series.str.contains('YES', case=False, na=False).sum()
    total_non_na = success_series.notna().sum()
    if total_non_na == 0:
        success_rate_pct = 0.0
    else:
        success_rate_pct = (yes_count / total_non_na) * 100

    # ---------------------------------------------------------------
    # 8️⃣ Output the computed metrics
    # ---------------------------------------------------------------
    print(f"Total calls: {total_calls}")
    print(f"Inbound calls: {inbound_calls} ({inbound_calls/total_calls*100:.2f}%)")
    print(f"Outbound calls: {outbound_calls} ({outbound_calls/total_calls*100:.2f}%)")
    print(f"Average duration (minutes): {avg_duration_minutes:.2f}")
    print(f"Success rate (YES/NO): {success_rate_pct:.2f}%")

    # ---------------------------------------------------------------
    # 9️⃣ Output per‑queue counts
    # ---------------------------------------------------------------
    print("\\nCalls per queue (call‑path):")
    for queue, count in queue_counts.items():
        queue_display = str(queue) if pd.notna(queue) else "<missing>"
        print(f"{queue_display}: {count}")

    # ---------------------------------------------------------------
    # 10️⃣ Write results to an Excel workbook
    # ---------------------------------------------------------------
    summary_df = pd.DataFrame({
        "Metric": [
            "Total Calls",
            "Inbound Calls",
            "Outbound Calls",
            "Average Duration (minutes)",
            "Success Rate (%)"
        ],
        "Value": [
            total_calls,
            inbound_calls,
            outbound_calls,
            avg_duration_minutes,
            success_rate_pct
        ]
    })

    # Write per‑queue counts to another sheet
    queue_df = queue_counts.reset_index()
    queue_df.columns = ["Queue", "Calls"]
    excel_path = "genesys_call_summary.xlsx"

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        queue_df.to_excel(writer, sheet_name="Calls per Queue", index=False)

    print(f"\n✅ Summary written to {excel_path}")

# ----------------------------------------------------------------------
if __name__ == "__main__":
    # Allow the user to pass a CSV path on the command line; otherwise use a
    # sensible default that points to a typical Genesys export location.
    if len(sys.argv) > 1:
        csv_file = sys.argv[1]
    else:
        # Default location – adjust if your export lives elsewhere.
        csv_file = "/Users/trey/.hermes/cache/documents/doc_20fba0e8569e_2026-06-24 Interactions.csv"

    main(csv_file)