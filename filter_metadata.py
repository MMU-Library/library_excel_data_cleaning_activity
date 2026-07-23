#!/usr/bin/env python3
"""
filter_metadata.py – Pre‑filter metadata extract for date and publisher tasks.

Filters:
  - Dates: conference_item, exhibition, performance → Outputs/metadata_dates_filtered.xlsx
  - Publishers: article, conference_item → Outputs/metadata_publishers_filtered.xlsx

Usage:
    python filter_metadata.py --input inputs/metadata_extract_20260127.xlsx
"""

import argparse
from pathlib import Path

import pandas as pd


def find_column(df, col_name):
    """Find column name case‑insensitively; return the actual column name or None."""
    for col in df.columns:
        if col.lower() == col_name.lower():
            return col
    return None


def filter_metadata(input_file):
    """Read raw metadata, filter to separate date and publisher files."""
    df = pd.read_excel(input_file, engine="openpyxl")

    # Find the eprints_type column (case‑insensitive)
    eprint_col = find_column(df, "eprints_type")
    if eprint_col is None:
        print(f"Error: Column 'eprints_type' not found in the input file.")
        print(f"Available columns: {list(df.columns)}")
        return

    # Date filter: conference_item, exhibition, performance
    date_types = ["conference_item", "exhibition", "performance"]
    df_dates = df[df[eprint_col].isin(date_types)].copy()
    df_dates = df_dates.reset_index(drop=False).rename(
        columns={"index": "original_index"}
    )

    # Publisher filter: article, conference_item
    pub_types = ["article", "conference_item"]
    df_publishers = df[df[eprint_col].isin(pub_types)].copy()
    df_publishers = df_publishers.reset_index(drop=False).rename(
        columns={"index": "original_index"}
    )

    # Create Outputs directory
    output_dir = Path("Outputs")
    output_dir.mkdir(exist_ok=True)

    # Save date filtered file
    date_output = output_dir / "metadata_dates_filtered.xlsx"
    df_dates.to_excel(date_output, index=False, engine="openpyxl")
    print(f"Dates filtered file saved: {date_output}")
    print(f"  - {len(df_dates)} rows (conference_item, exhibition, performance)")

    # Save publisher filtered file
    pub_output = output_dir / "metadata_publishers_filtered.xlsx"
    df_publishers.to_excel(pub_output, index=False, engine="openpyxl")
    print(f"Publishers filtered file saved: {pub_output}")
    print(f"  - {len(df_publishers)} rows (article, conference_item)")

    # Save combined filtered file (for backwards compatibility)
    all_types = date_types + pub_types
    df_all = df[df[eprint_col].isin(all_types)].copy()
    df_all = df_all.reset_index(drop=False).rename(columns={"index": "original_index"})
    all_output = output_dir / "metadata_filtered.xlsx"
    df_all.to_excel(all_output, index=False, engine="openpyxl")
    print(f"Combined filtered file saved: {all_output}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Filter metadata for date and publisher tasks."
    )
    parser.add_argument(
        "--input",
        default="inputs/metadata_extract_20260127.xlsx",
        help="Input raw metadata file",
    )
    args = parser.parse_args()

    filter_metadata(args.input)
