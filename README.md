# Research Metadata Cleanup Pipeline

A suite of Python scripts to clean and standardise Excel exports from the institutional repository (EPrints) before upload.

The pipeline handles three separate tasks:

1. **Date standardisation** – converts event dates to `DD‑MM‑YYYY` for conference items, exhibitions, and performances.  
2. **Publisher name standardisation** – merges variant publisher names using fuzzy matching and hard‑coded mappings.  
3. **Author name standardisation** – analyses first and last names using staff IDs (where available) and fuzzy matching to identify duplicates.

---

## 📂 Project Structure
research-data-cleanup/
├── app.py # Streamlit UI – launch this for interactive use
├── config.py # Centralised settings (thresholds, paths)
├── pyproject.toml # Package metadata (optional)
├── requirements.txt # Python dependencies
├── .gitignore # Files/folders to ignore in Git
├── README.md # This file
│
├── date_standardisation_task_one_activity_two.py
├── publisher_name_standardisation_task_two_activity_one.py
├── author_name_standardisation_task_three_activity_one.py
│
├── inputs/ # Place your input Excel files here
│ └── metadata_extract_20260127.xlsx
│ └── authors_20260127_WorkingFile.xlsx
│
├── Outputs/ # All generated outputs are saved here
│ ├── metadata_final_cleaned.xlsx
│ ├── publisher_review_index.xlsx
│ ├── publisher_cluster_summary.csv
│ ├── author_review_index.xlsx
│ ├── author_cluster_summary.csv
│ └── *.log # Log files for each task
│
└── DeadFiles/ # Archived/obsolete files (ignored by Git)
└── (old scripts, exploratory files)

text

---

## ⚙️ Setup

### Prerequisites

- Python **3.8** or higher
- A virtual environment (recommended)

### Install dependencies

```bash
# Create and activate a virtual environment
python -m venv venv
venv\Scripts\activate      # Windows
source venv/bin/activate   # macOS/Linux

# Install required packages
pip install --upgrade pip
pip install -r requirements.txt
Configuration
Adjust settings in config.py if needed:

SIMILARITY_THRESHOLD_AUTO – high‑confidence auto‑merge threshold (default 85)

SIMILARITY_THRESHOLD_REVIEW – lower bound for review (default 65)

FREQ_RATIO_FOR_AUTO – frequency ratio for auto‑merge (default 3)

AUTHOR_SIMILARITY_THRESHOLD_AUTO – for author matching (default 85)

AUTHOR_SIMILARITY_THRESHOLD_REVIEW – for author review (default 75)

🚀 Usage
You can run the pipeline in two ways:

Option A – Interactive UI (recommended for non‑coders)
bash
streamlit streamlit run app.py
A browser tab will open. Upload your files and choose which tasks to run.

Metadata Cleanup – expects metadata_extract_*.xlsx; runs dates then publishers.

Author Name Cleanup – expects authors_*.xlsx; runs author standardisation separately.

The UI provides live logs as each script runs, showing progress and any warnings in real time. This is especially useful for long‑running tasks like author name standardisation.

Option B – Command‑line (individual scripts)
Each script accepts --input and --output arguments:

Dates
bash
python date_standardisation_task_one_activity_two.py --input inputs/metadata_extract_20260127.xlsx --output Outputs/dates_cleaned.xlsx
Publishers
bash
python publisher_name_standardisation_task_two_activity_one.py --input Outputs/dates_cleaned.xlsx --output Outputs/publishers_cleaned.xlsx --review Outputs/publisher_review_index.xlsx --clusters Outputs/publisher_cluster_summary.csv
Authors
bash
python author_name_standardisation_task_three_activity_one.py --input inputs/authors_20260127_WorkingFile.xlsx --output Outputs/authors_cleaned.xlsx --review Outputs/author_review_index.xlsx --clusters Outputs/author_cluster_summary.csv
You can also pass an optional --override file (CSV/Excel) to force‑merge or keep specific pairs.

📥 Outputs
File	Description
metadata_final_cleaned.xlsx	Original metadata with standardised dates and publisher names
publisher_review_index.xlsx	Pairs of publisher names that need human judgement
publisher_cluster_summary.csv	Shows how publisher variants were grouped and the chosen canonical name
authors_cleaned.xlsx	Original authors file with additional columns for standardised first/last names
author_review_index.xlsx	Pairs of author records needing human validation. Includes resource_id, id, similarity, freq_1/2, same_id flag, and a suggested action (Merge or Review)
author_cluster_summary.csv	Shows how author records were grouped into clusters
All output files preserve the original Excel styling (fonts, colours, column widths, formulas).

🛠️ Troubleshooting
Excel warning: “Removed Records: Formula from /xl/worksheets/sheet1.xml”
This is caused by Excel formulas that openpyxl cannot fully preserve. The data is not affected.

Fix: All scripts now use data_only=True when loading workbooks, so formulas are converted to static values. If you still see the warning, you can safely ignore it – the content is correct.

Author script seems to take a long time
The author script performs pairwise fuzzy matching on up to 117,000 rows. This is computationally intensive. On a typical machine, it completes in about 10–15 minutes for the full dataset. The UI shows live logs so you can monitor progress. You can also run a sample first to test logic before running the full file.

“ModuleNotFoundError” when running
Make sure all dependencies are installed:

bash
pip install -r requirements.txt
File paths not found
Always use absolute paths or run scripts from the root folder. The UI handles temporary files automatically.

🤝 Contributing
Use a virtual environment (venv) for development.

Follow the existing folder structure.

Add clear comments in your code.

Update this README.md when adding new functionality.

Run black or ruff to format code (optional).

📄 License
This project is provided under the MIT License

📬 Contact
For questions or support, please contact the Digital Library Services Team.