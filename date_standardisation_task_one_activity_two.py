#!/usr/bin/env python3
"""
Date standardisation script – accepts command‑line arguments for input and output.

Usage:
    python date_standardisation_task_one_activity_two.py --input input.xlsx --output output.xlsx
"""

import argparse
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from dateutil.parser import parse

# ------------------------------------------------------------
# Parse command‑line arguments
# ------------------------------------------------------------
parser = argparse.ArgumentParser(
    description="Standardise event dates in metadata exports."
)
parser.add_argument("--input", required=True, help="Path to input Excel file")
parser.add_argument("--output", required=True, help="Path to output Excel file")
args = parser.parse_args()

# ------------------------------------------------------------
# Read the input file
# ------------------------------------------------------------
input_path = Path(args.input)
if not input_path.exists():
    print(f"Error: Input file not found: {input_path}")
    exit(1)

df_filtered = pd.read_excel(input_path, engine="openpyxl")
print(f"Loaded {len(df_filtered)} rows from {input_path}")


# ------------------------------------------------------------
# Define the date parser with logging of errors
# ------------------------------------------------------------
def parse_date_to_ddmmyyyy(value, row_id_col_name, error_log):
    """Convert a date string to dd-mm-yyyy format. If conversion fails, return None and log the error."""
    if pd.isna(value) or value == "":
        return None
    val = str(value).strip()

    # Primary parser
    try:
        dt = pd.to_datetime(val, dayfirst=True, errors="raise")
        if dt.year < 1900 or dt.year > 2025:
            error_log.append(f"year out of range: row={row_id_col_name}, value='{val}'")
            return None
        return dt.strftime("%d-%m-%Y")
    except Exception:
        pass

    # Secondary parser – remove suffixes (1st, 2nd, etc.)
    try:
        clean = re.sub(r"(\d{1,2})(st|nd|rd|th)", r"\1", val)
        dt = parse(clean, dayfirst=True)
        if dt.year < 1900 or dt.year > 2025:
            error_log.append(f"year out of range: row={row_id_col_name}, value='{val}'")
            return None
        return dt.strftime("%d-%m-%Y")
    except Exception:
        error_log.append(f"unparseable date: row={row_id_col_name}, value='{val}'")
        return None


# ------------------------------------------------------------
# Apply the function to event_dates_start and event_dates_end
# ------------------------------------------------------------
error_log = []

for col in ["event_dates_start", "event_dates_end"]:
    if col in df_filtered.columns:
        df_filtered[col] = df_filtered.apply(
            lambda row: parse_date_to_ddmmyyyy(
                row[col], row_id_col_name="id", error_log=error_log
            ),
            axis=1,
        )
    else:
        print(f"Warning: Column '{col}' not found in input file.")

# ------------------------------------------------------------
# Save the error log
# ------------------------------------------------------------
log_file = "dates_standardisation_error_log.txt"
with open(log_file, "w") as log:
    if error_log:
        log.write("Date parsing errors (see below):\n")
        for entry in error_log:
            log.write(entry + "\n")
    else:
        log.write("No errors found.\n")
print(f"Date parsing complete. Errors logged to {log_file}")

# ------------------------------------------------------------
# Save the output file
# ------------------------------------------------------------
output_path = Path(args.output)
output_path.parent.mkdir(parents=True, exist_ok=True)
df_filtered.to_excel(output_path, index=False, engine="openpyxl")
print(f"Standardised dates saved to {output_path}")
