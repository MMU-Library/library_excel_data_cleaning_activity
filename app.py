#!/usr/bin/env python3
"""
app.py – Streamlit UI for the Open Research Excel Data Cleansing Pipeline.

This UI orchestrates three data‑cleaning tasks:
  1. Date standardisation – converts event dates to DD‑MM‑YYYY.
  2. Publisher name standardisation – merges variant publisher names.
  3. Author name standardisation – standardises author names.

Usage:
    streamlit run app.py
"""

import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import pandas as pd
import streamlit as st

# ------------------------------------------------------------
# Determine where the scripts are (relative to this file)
# ------------------------------------------------------------
BASE_DIR = Path(__file__).parent.absolute()

# Map friendly names to actual script filenames
SCRIPT_MAP = {
    "dates": BASE_DIR / "date_standardisation_task_one_activity_two.py",
    "publishers": BASE_DIR / "publisher_name_standardisation_task_two_activity_one.py",
    "authors": BASE_DIR / "author_name_standardisation_task_three_activity_one.py",
}

# Check which scripts exist
for name, path in SCRIPT_MAP.items():
    if not path.exists():
        st.sidebar.warning(f"⚠️ {name} script not found: {path.name}")

# ------------------------------------------------------------
# Session state initialisation
# ------------------------------------------------------------
if "processed_metadata" not in st.session_state:
    st.session_state.processed_metadata = False
if "processed_authors" not in st.session_state:
    st.session_state.processed_authors = False
if "final_data" not in st.session_state:
    st.session_state.final_data = None
if "review_data" not in st.session_state:
    st.session_state.review_data = None
if "cluster_data" not in st.session_state:
    st.session_state.cluster_data = None
if "authors_data" not in st.session_state:
    st.session_state.authors_data = None


# ------------------------------------------------------------
# Helper function to run subprocesses with live output
# ------------------------------------------------------------
def run_with_live_output(cmd, placeholder, log_prefix=""):
    """Run a subprocess and stream its stdout/stderr to a Streamlit placeholder."""
    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        universal_newlines=True,
    )
    output_lines = []
    for line in iter(process.stdout.readline, ""):
        if line:
            output_lines.append(line)
            if len(output_lines) > 50:
                output_lines = output_lines[-50:]
            placeholder.text(f"{log_prefix}\n" + "".join(output_lines))
    process.wait()
    return process.returncode, "".join(output_lines)


# ------------------------------------------------------------
# Helper to merge date and publisher results back into original
# ------------------------------------------------------------
def merge_results(original_path, dates_path, publishers_path, output_path):
    """Merge date and publisher standardisation results back into the original dataframe."""
    # Load dataframes
    df_original = pd.read_excel(original_path, engine="openpyxl")
    df_dates = pd.read_excel(dates_path, engine="openpyxl")
    df_publishers = pd.read_excel(publishers_path, engine="openpyxl")

    # Ensure original_index columns exist
    if "original_index" not in df_dates.columns:
        df_dates["original_index"] = df_dates.index
    if "original_index" not in df_publishers.columns:
        df_publishers["original_index"] = df_publishers.index

    # Create a copy of original to update
    df_merged = df_original.copy()

    # Update dates for rows that are in the dates file
    for idx in df_dates["original_index"]:
        if idx in df_merged.index:
            row_dates = df_dates[df_dates["original_index"] == idx]
            if not row_dates.empty:
                if "event_dates_start" in row_dates.columns:
                    df_merged.at[idx, "event_dates_start"] = row_dates.iloc[0][
                        "event_dates_start"
                    ]
                if "event_dates_end" in row_dates.columns:
                    df_merged.at[idx, "event_dates_end"] = row_dates.iloc[0][
                        "event_dates_end"
                    ]

    # Update publishers for rows that are in the publishers file
    for idx in df_publishers["original_index"]:
        if idx in df_merged.index:
            row_pubs = df_publishers[df_publishers["original_index"] == idx]
            if not row_pubs.empty:
                # If the publisher script added a standardised column, use that
                if "publisher_standardised" in row_pubs.columns:
                    df_merged.at[idx, "publisher"] = row_pubs.iloc[0][
                        "publisher_standardised"
                    ]
                elif "publisher" in row_pubs.columns:
                    df_merged.at[idx, "publisher"] = row_pubs.iloc[0]["publisher"]

    # Save merged file
    df_merged.to_excel(output_path, index=False, engine="openpyxl")
    return output_path


# ------------------------------------------------------------
# Main UI
# ------------------------------------------------------------
st.set_page_config(
    page_title="Open Research Data Cleansing",
    page_icon="📊",
    layout="centered",
)

st.title("📊 Open Research Excel Data Cleansing")
st.markdown(
    """
    This tool cleans metadata exports for the Research Repository.
    Select a tab below to get started.
    """
)

# ------------------------------------------------------------
# Create two tabs – one for each workflow
# ------------------------------------------------------------
tab_metadata, tab_authors = st.tabs(
    ["📁 Metadata Cleanup (Dates + Publishers)", "✍️ Author Name Cleanup"]
)

# ================================================================
# TAB 1: Metadata Cleanup (Dates + Publishers)
# ================================================================
with tab_metadata:
    st.markdown(
        """
        Upload your **metadata extract file** and this will:

        1. **Standardise event dates** – for conference items, exhibitions, and performances.
        2. **Standardise publisher names** – for articles and conference papers.
        """
    )

    uploaded_metadata = st.file_uploader(
        "Choose your metadata Excel file",
        type=["xlsx"],
        key="metadata_file",
        help="Upload the main metadata extract file (metadata_extract_*.xlsx).",
    )

    override_file = st.file_uploader(
        "Optional: Publisher override file (CSV/Excel)",
        type=["csv", "xlsx"],
        key="override_file",
        help="If you have manual overrides for publisher names, upload them here.",
    )

    run_metadata = st.button(
        "🚀 Run Metadata Cleanup",
        type="primary",
        use_container_width=True,
        disabled=uploaded_metadata is None,
    )

    if run_metadata and uploaded_metadata is not None:
        st.session_state.processed_metadata = False
        st.session_state.final_data = None
        st.session_state.review_data = None
        st.session_state.cluster_data = None

        progress_bar = st.progress(0, text="Initialising...")
        status_text = st.empty()
        log_placeholder = st.empty()

        try:
            # Save uploaded file to temp
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
                tmp_input.write(uploaded_metadata.getbuffer())
                input_path = tmp_input.name

            override_path = None
            if override_file is not None:
                suffix = ".xlsx" if override_file.name.endswith(".xlsx") else ".csv"
                with tempfile.NamedTemporaryFile(
                    delete=False, suffix=suffix
                ) as tmp_override:
                    tmp_override.write(override_file.getbuffer())
                    override_path = tmp_override.name

            # Output paths
            interim_dates_path = os.path.join(
                tempfile.gettempdir(), "dates_cleaned.xlsx"
            )
            interim_pubs_path = os.path.join(
                tempfile.gettempdir(), "publishers_cleaned.xlsx"
            )
            final_path = os.path.join(tempfile.gettempdir(), "final_cleaned.xlsx")

            (BASE_DIR / "Outputs").mkdir(exist_ok=True)
            review_path = str(BASE_DIR / "Outputs" / "publisher_review_index.xlsx")
            cluster_path = str(BASE_DIR / "Outputs" / "publisher_cluster_summary.csv")

            # --------------------------------------------------------
            # Step 0: Filter metadata into separate date and publisher files
            # --------------------------------------------------------
            progress_bar.progress(10, text="Filtering metadata...")
            status_text.info("⏳ Filtering rows for date and publisher tasks...")

            cmd_filter = [
                sys.executable,
                str(BASE_DIR / "filter_metadata.py"),
                "--input",
                input_path,
            ]
            result_filter = subprocess.run(cmd_filter, capture_output=True, text=True)

            if result_filter.returncode != 0:
                st.error(f"❌ Filtering failed:\n{result_filter.stderr}")
                st.stop()

            st.success("✅ Filtering complete.")
            dates_filtered_path = BASE_DIR / "Outputs" / "metadata_dates_filtered.xlsx"
            pubs_filtered_path = (
                BASE_DIR / "Outputs" / "metadata_publishers_filtered.xlsx"
            )

            # --------------------------------------------------------
            # Step 1: Date standardisation (using dates‑filtered file)
            # --------------------------------------------------------
            progress_bar.progress(20, text="Step 1/2: Standardising event dates...")
            status_text.info("⏳ Processing date fields...")

            cmd1 = [
                sys.executable,
                str(SCRIPT_MAP["dates"]),
                "--input",
                str(dates_filtered_path),
                "--output",
                interim_dates_path,
            ]
            returncode1, _ = run_with_live_output(cmd1, log_placeholder, "📅 Date log:")

            if returncode1 != 0:
                st.error(f"❌ Date standardisation failed with code {returncode1}")
                st.stop()

            progress_bar.progress(50, text="Step 1 complete. Dates standardised.")
            st.success("✅ Step 1 complete: Dates standardised.")
            log_placeholder.empty()

            # --------------------------------------------------------
            # Step 2: Publisher name standardisation (using publishers‑filtered file)
            # --------------------------------------------------------
            progress_bar.progress(60, text="Step 2/2: Standardising publisher names...")
            status_text.info("⏳ Analysing publisher names...")

            cmd2 = [
                sys.executable,
                str(SCRIPT_MAP["publishers"]),
                "--input",
                str(pubs_filtered_path),
                "--output",
                interim_pubs_path,
                "--review",
                review_path,
                "--clusters",
                cluster_path,
            ]
            if override_path:
                cmd2.extend(["--override", override_path])

            returncode2, _ = run_with_live_output(
                cmd2, log_placeholder, "📚 Publisher log:"
            )

            if returncode2 != 0:
                st.error(f"❌ Publisher standardisation failed with code {returncode2}")
                st.stop()

            progress_bar.progress(90, text="Step 2 complete. Publishers standardised.")
            st.success("✅ Step 2 complete: Publishers standardised.")
            log_placeholder.empty()

            # --------------------------------------------------------
            # Step 3: Merge results back into the original file
            # --------------------------------------------------------
            progress_bar.progress(95, text="Merging results...")
            status_text.info("⏳ Combining cleaned fields into final file...")

            merge_results(input_path, interim_dates_path, interim_pubs_path, final_path)

            # --------------------------------------------------------
            # Store results in session state
            # --------------------------------------------------------
            with open(final_path, "rb") as f:
                st.session_state.final_data = f.read()

            if Path(review_path).exists():
                with open(review_path, "rb") as f:
                    st.session_state.review_data = f.read()

            if Path(cluster_path).exists():
                with open(cluster_path, "rb") as f:
                    st.session_state.cluster_data = f.read()

            st.session_state.processed_metadata = True

            # Clean up temp files
            os.unlink(input_path)
            if override_path and os.path.exists(override_path):
                os.unlink(override_path)
            # Optionally delete filtered files (they are in Outputs, keep them for reference)

            progress_bar.progress(100, text="Metadata cleanup complete!")
            status_text.empty()

            st.balloons()
            st.success("🎉 Metadata cleanup completed successfully!")

        except Exception as e:
            st.error(f"An unexpected error occurred: {e}")
            st.stop()

    # --------------------------------------------------------
    # Download buttons for metadata results
    # --------------------------------------------------------
    if st.session_state.processed_metadata:
        st.divider()
        st.subheader("📥 Downloads")

        col1, col2, col3 = st.columns(3)

        with col1:
            if st.session_state.final_data is not None:
                st.download_button(
                    label="📥 Final Cleaned Metadata",
                    data=st.session_state.final_data,
                    file_name="metadata_final_cleaned.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        with col2:
            if st.session_state.review_data is not None:
                st.download_button(
                    label="📋 Publisher Review Index",
                    data=st.session_state.review_data,
                    file_name="publisher_review_index.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        with col3:
            if st.session_state.cluster_data is not None:
                st.download_button(
                    label="📊 Cluster Summary (CSV)",
                    data=st.session_state.cluster_data,
                    file_name="publisher_cluster_summary.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

        if st.button(
            "🔄 Clear Results", use_container_width=True, key="clear_metadata"
        ):
            for key in [
                "processed_metadata",
                "final_data",
                "review_data",
                "cluster_data",
            ]:
                if key in st.session_state:
                    st.session_state[key] = (
                        False if key == "processed_metadata" else None
                    )
            st.rerun()

# ================================================================
# TAB 2: Author Name Cleanup
# ================================================================
with tab_authors:
    st.markdown(
        """
        Upload your **authors file** and this will standardise author names.

        The tool will analyse `first_name` and `last_name` columns, using the `id`
        (university ID) where available to validate duplicates.
        """
    )

    uploaded_authors = st.file_uploader(
        "Choose your authors Excel file",
        type=["xlsx"],
        key="authors_file",
        help="Upload the authors file (authors_*.xlsx).",
    )

    authors_script_exists = SCRIPT_MAP["authors"].exists()

    if not authors_script_exists:
        st.warning(
            "⚠️ The author standardisation script is not yet available. "
            "This feature will be enabled when the script is added to the project."
        )

    run_authors = st.button(
        "✍️ Run Author Name Standardisation",
        type="primary",
        use_container_width=True,
        disabled=uploaded_authors is None or not authors_script_exists,
    )

    if run_authors and uploaded_authors is not None:
        st.session_state.processed_authors = False
        st.session_state.authors_data = None

        authors_progress = st.progress(0, text="Processing authors...")
        authors_status = st.empty()
        log_placeholder = st.empty()

        try:
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=".xlsx"
            ) as tmp_authors:
                tmp_authors.write(uploaded_authors.getbuffer())
                authors_input = tmp_authors.name

            authors_output = os.path.join(tempfile.gettempdir(), "authors_cleaned.xlsx")

            (BASE_DIR / "Outputs").mkdir(exist_ok=True)
            review_author_path = str(BASE_DIR / "Outputs" / "author_review_index.xlsx")
            cluster_author_path = str(
                BASE_DIR / "Outputs" / "author_cluster_summary.csv"
            )

            authors_progress.progress(30, text="Analysing author names...")
            authors_status.info("⏳ Standardising author names...")

            cmd3 = [
                sys.executable,
                str(SCRIPT_MAP["authors"]),
                "--input",
                authors_input,
                "--output",
                authors_output,
                "--review",
                review_author_path,
                "--clusters",
                cluster_author_path,
            ]

            returncode3, _ = run_with_live_output(
                cmd3, log_placeholder, "📋 Author log:"
            )

            if returncode3 != 0:
                st.error(f"❌ Author standardisation failed with code {returncode3}")
                st.stop()

            authors_progress.progress(90, text="Authors processed.")
            st.success("✅ Author standardisation complete.")

            # Small delay to ensure file is written
            time.sleep(1)

            # Check for output file
            output_paths_to_try = [
                Path(authors_output),
                BASE_DIR / "Outputs" / "authors_cleaned.xlsx",
                Path(tempfile.gettempdir()) / "authors_cleaned.xlsx",
            ]

            found_path = None
            for p in output_paths_to_try:
                if p.exists():
                    found_path = p
                    break

            if found_path:
                with open(found_path, "rb") as f:
                    st.session_state.authors_data = f.read()
                st.session_state.processed_authors = True
                st.success(f"✅ Output file loaded from: {found_path}")
            else:
                st.error("❌ Output file not found. Check the logs above for errors.")
                st.info(f"Expected location: {authors_output}")
                temp_dir = Path(tempfile.gettempdir())
                xlsx_files = list(temp_dir.glob("*.xlsx"))
                if xlsx_files:
                    st.info(
                        f"Found these .xlsx files in temp: {[f.name for f in xlsx_files]}"
                    )

            os.unlink(authors_input)

            authors_progress.progress(100, text="Author cleanup complete!")
            authors_status.empty()
            log_placeholder.empty()

            if st.session_state.processed_authors:
                st.balloons()
                st.success("🎉 Author standardisation completed successfully!")

        except Exception as e:
            st.error(f"An unexpected error occurred: {e}")
            st.stop()

    if st.session_state.processed_authors:
        st.divider()
        st.subheader("📥 Download")

        if st.session_state.authors_data is not None:
            st.download_button(
                label="📝 Cleaned Authors File",
                data=st.session_state.authors_data,
                file_name="authors_cleaned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.warning("No data available for download.")

        if st.button("🔄 Clear Results", use_container_width=True, key="clear_authors"):
            st.session_state.processed_authors = False
            st.session_state.authors_data = None
            st.rerun()

# ------------------------------------------------------------
# Footer
# ------------------------------------------------------------
st.divider()
st.caption(
    "Built for the MMU Research Repository data cleaning project. "
    "Scripts: date_standardisation_task_one_activity_two.py, "
    "publisher_name_standardisation_task_two_activity_one.py, "
    "and author_name_standardisation_task_three_activity_one.py."
)
