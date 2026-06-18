#!/usr/bin/env python3

"""

app.py – Streamlit UI for the Open Research Excel Data Cleansing Pipeline.
 
Runs date standardisation and publisher name standardisation sequentially,

then offers the final cleaned file for download.
 
Usage:

    streamlit run app.py

"""
 
import os

import subprocess

import sys

import tempfile

from pathlib import Path
 
import streamlit as st
 
# ------------------------------------------------------------

# Determine where the scripts are (relative to this file)

# ------------------------------------------------------------

BASE_DIR = Path(__file__).parent.absolute()
 
# Student's scripts are in the root directory (not in a 'scripts' subfolder)

# Map friendly names to actual script filenames

SCRIPT_MAP = {

    "dates": BASE_DIR / "date_standardisation_task_one_activity_two.py",

    "publishers": BASE_DIR / "publisher_name_standardisation_task_two_activity_one.py",

}
 
# Check if scripts exist

for name, path in SCRIPT_MAP.items():

    if not path.exists():

        st.error(f"❌ {name} script not found at: {path}")

        st.info(

            f"Please ensure '{path.name}' is in the same folder as this app."

        )

        st.stop()
 
# ------------------------------------------------------------

# Streamlit UI

# ------------------------------------------------------------

st.set_page_config(

    page_title="Open Research Data Cleansing",

    page_icon="📊",

    layout="centered",

)
 
st.title("📊 Open Research Excel Data Cleansing")

st.markdown(

    """

    Upload your `metadata_extract.xlsx` file, and this tool will:

    1. **Standardise event dates** – for conference items, exhibitions, and performances

    2. **Standardise publisher names** – for articles and conference papers

    3. **Give you the final cleaned file** – preserving all original styling

    """

)
 
# ------------------------------------------------------------

# Session state initialisation

# ------------------------------------------------------------

if "processed" not in st.session_state:

    st.session_state.processed = False

if "final_data" not in st.session_state:

    st.session_state.final_data = None

if "review_data" not in st.session_state:

    st.session_state.review_data = None
 
# ------------------------------------------------------------

# File uploaders

# ------------------------------------------------------------

uploaded_file = st.file_uploader(

    "📁 Choose your input Excel file",

    type=["xlsx"],

    key="input_file",

    help="Upload the original metadata extract file (metadata_extract_*.xlsx).",

)
 
override_file = st.file_uploader(

    "📋 Optional: Publisher override file (CSV/Excel)",

    type=["csv", "xlsx"],

    key="override_file",

    help=(

        "If you have manual overrides for publisher names, upload them here. "

        "Columns should be: publisher_1, publisher_2, action"

    ),

)
 
# ------------------------------------------------------------

# Run button and pipeline execution

# ------------------------------------------------------------

if uploaded_file is not None:

    if st.button("🚀 Run Full Cleanup Pipeline", type="primary"):

        # Reset session state

        st.session_state.processed = False

        st.session_state.final_data = None

        st.session_state.review_data = None
 
        try:

            # ------------------------------------------------------------

            # 1. Save uploaded files to temporary locations

            # ------------------------------------------------------------

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:

                tmp_input.write(uploaded_file.getbuffer())

                input_path = tmp_input.name
 
            override_path = None

            if override_file is not None:

                suffix = ".xlsx" if override_file.name.endswith(".xlsx") else ".csv"

                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_override:

                    tmp_override.write(override_file.getbuffer())

                    override_path = tmp_override.name
 
            # Define output paths

            interim_path = os.path.join(tempfile.gettempdir(), "interim_dates_cleaned.xlsx")

            final_path = os.path.join(tempfile.gettempdir(), "final_cleaned.xlsx")
 
            # Ensure outputs directory exists (for logs and review files)

            (BASE_DIR / "Outputs").mkdir(exist_ok=True)

            review_path = str(BASE_DIR / "Outputs" / "publisher_review_index.xlsx")
 
            # ------------------------------------------------------------

            # 2. Run date standardisation

            # ------------------------------------------------------------

            st.info("⏳ Step 1/2: Standardising event dates...")

            progress_placeholder = st.empty()

            progress_placeholder.text("Running date_standardisation_task_one_activity_two.py...")
 
            cmd1 = [

                sys.executable,

                str(SCRIPT_MAP["dates"]),

                "--input", input_path,

                "--output", interim_path,

            ]

            result1 = subprocess.run(cmd1, capture_output=True, text=True)
 
            if result1.returncode != 0:

                st.error(f"❌ Date standardisation failed:\n{result1.stderr}")

                st.stop()

            else:

                st.success("✅ Step 1 complete: Dates standardised.")

                with st.expander("📋 View task output (last 500 chars)"):

                    st.text(result1.stdout[-500:] if result1.stdout else "No output")
 
            # ------------------------------------------------------------

            # 3. Run publisher name standardisation

            # ------------------------------------------------------------

            st.info("⏳ Step 2/2: Standardising publisher names...")

            progress_placeholder.text("Running publisher_name_standardisation_task_two_activity_one.py...")
 
            cmd2 = [

                sys.executable,

                str(SCRIPT_MAP["publishers"]),

                "--input", interim_path,

                "--output", final_path,

                "--review", review_path,

            ]

            if override_path:

                cmd2.extend(["--override", override_path])
 
            result2 = subprocess.run(cmd2, capture_output=True, text=True)
 
            if result2.returncode != 0:

                st.error(f"❌ Publisher standardisation failed:\n{result2.stderr}")

                st.stop()

            else:

                st.success("✅ Step 2 complete: Publishers standardised.")

                with st.expander("📋 View task output (last 500 chars)"):

                    st.text(result2.stdout[-500:] if result2.stdout else "No output")
 
            # ------------------------------------------------------------

            # 4. Store results in session state

            # ------------------------------------------------------------

            with open(final_path, "rb") as f:

                st.session_state.final_data = f.read()
 
            if Path(review_path).exists():

                with open(review_path, "rb") as f:

                    st.session_state.review_data = f.read()

            else:

                st.session_state.review_data = None
 
            st.session_state.processed = True
 
            # Clean up temp files

            os.unlink(input_path)

            if override_path and os.path.exists(override_path):

                os.unlink(override_path)
 
            st.balloons()

            st.success(

                "🎉 Pipeline completed successfully! You can now download the cleaned files below."

            )
 
        except Exception as e:

            st.error(f"An unexpected error occurred: {e}")

            st.stop()
 
    # ------------------------------------------------------------

    # 5. Show download buttons if processing is complete

    # ------------------------------------------------------------

    if st.session_state.processed:

        st.divider()

        st.subheader("📥 Downloads")
 
        col1, col2 = st.columns(2)
 
        with col1:

            if st.session_state.final_data is not None:

                st.download_button(

                    label="📥 Download Final Cleaned File",

                    data=st.session_state.final_data,

                    file_name="metadata_final_cleaned.xlsx",

                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                    use_container_width=True,

                )

            else:

                st.warning("Final file data not available.")
 
        with col2:

            if st.session_state.review_data is not None:

                st.download_button(

                    label="📋 Download Publisher Review Index",

                    data=st.session_state.review_data,

                    file_name="publisher_review_index.xlsx",

                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                    use_container_width=True,

                )

            else:

                st.info("No review index generated (no uncertain publisher pairs).")
 
        # Reset button

        if st.button("🔄 Start Over (clear results)", use_container_width=True):

            st.session_state.processed = False

            st.session_state.final_data = None

            st.session_state.review_data = None

            st.rerun()
 
else:

    st.info("👆 Please upload your metadata file to begin.")
 
# ------------------------------------------------------------

# Footer

# ------------------------------------------------------------

st.divider()

st.caption(

    "Built for the MMU Research Repository data cleaning project. "

    "Scripts: date_standardisation_task_one_activity_two.py and "

    "publisher_name_standardisation_task_two_activity_one.py."

)
 