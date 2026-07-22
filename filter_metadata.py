#!/usr/bin/env python3
"""
filter_metadata.py – Pre‑filter metadata extract for date and publisher tasks.

Filters rows where eprints_type is in:
  - conference_item, exhibition, performance  (for dates)
  - article, conference_item                   (for publishers)

Combined set: article, conference_item, exhibition, performance.

Usage:
    python filter_metadata.py --input inputs/metadata_extract_20260127.xlsx --output Outputs/metadata_filtered.xlsx
"""

import argparse
from pathlib import Path

import pandas as pd


def filter_metadata(input_file, output_file):
    """Read raw metadata, filter to relevant rows, and save."""
    df = pd.read_excel(input_file, engine="openpyxl")

    # Allowed types for both tasks
    allowed_types = ["article", "conference_item", "exhibition", "performance"]
    df_filtered = df[df["eprints_type"].isin(allowed_types)].copy()

    # Check for required columns and warn if missing
    required_cols = [
        "eprints_type",
        "publisher",
        "event_dates_start",
        "event_dates_end",
    ]
    missing = [c for c in required_cols if c not in df_filtered.columns]
    if missing:
        print(f"Warning: Missing columns in filtered output: {missing}")
        # Continue anyway – the downstream scripts will handle missing columns

    df_filtered.to_excel(output_file, index=False, engine="openpyxl")

    print(f"Filtered file saved to: {output_file}")
    print(f"  - Rows kept: {len(df_filtered)}")
    print(f"  - Types present: {df_filtered['eprints_type'].unique().tolist()}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Filter metadata for date and publisher tasks."
    )
    parser.add_argument(
        "--input",
        default="inputs/metadata_extract_20260127.xlsx",
        help="Input raw metadata file",
    )
    parser.add_argument(
        "--output",
        default="Outputs/metadata_filtered.xlsx",
        help="Output filtered file",
    )
    args = parser.parse_args()

    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    filter_metadata(args.input, args.output)
