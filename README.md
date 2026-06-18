# Library Excel Data Cleaning Activity

## Description

This repository contains Python scripts developed to support the cleansing and standardisation of metadata exported from the Manchester Metropolitan University (MMU) Research Repository.

The project addresses a series of data quality tasks involving event records, publication metadata, and publisher name authority control. The scripts automate filtering, standardisation, and quality assurance processes to improve metadata consistency and prepare records for downstream reporting and analysis.

---

## Project Objectives

The repository was created to address the following data-cleaning activities.

### Task 1 – Event Date Standardisation

**Source file:** `metadata_extract_20260127.xlsx`

**Requirements**

1. Filter records where `eprints_type` contains:

   * `conference_item`
   * `exhibition`
   * `performance`

2. Analyse and standardise the following date fields:

   * `event_dates_start`
   * `event_dates_end`

3. Convert date values into a consistent format.

**Chosen standard format**

```text
DD-MM-YYYY
```

Examples:

```text
01-03-2024
15-10-2025
```

---

### Task 2 – Publisher Name Standardisation

**Source file:** `metadata_extract_20260127.xlsx`

**Requirements**

1. Filter records where `eprints_type` contains:

   * `article`
   * `conference_item`

2. Analyse publisher name variants in the `publisher` field.

3. Develop a methodology for publisher authority control and standardisation.

---

### Task 3 – Author Name Standardisation (Methodology Proposal)

**Source file:** `authors_20260127_WorkingFile.xlsx`

**Requirements**

Analyse:

* `first_name`
* `last_name`

Use the staff identifier (`id`) where available to assist in duplicate detection and identity resolution.

The objective is to identify potential duplicate author records and create a review process for manual validation where confidence is low.

**Note:** This repository currently contains implemented solutions for Tasks 1 and 2. Task 3 remains a proposed methodology and has not yet been implemented as a script.

---

# Repository Structure

```text
library_excel_data_cleaning_activity/
│
├── Inputs/
│   ├── metadata_extract_20260127.xlsx
│   └── authors_20260127_WorkingFile.xlsx
│
├── Outputs/
│   ├── metadata_extract_20260127_filtered_task_one_activity_one.xlsx
│   ├── metadata_extract_20260127_standardised_dates_activity_one_task_two.xlsx
│   ├── metadata_extract_20260127_filtered_task_two_activity_one.xlsx
│   ├── metadata_extract_20260127_publishers_cleaned.xlsx
│   ├── publisher_review_index.xlsx
│   └── publisher_cluster_summary.csv
│
├── value_filtering_task_one_activity_one.py
├── date_standardisation_task_one_activity_two.py
├── value_filtering_task_two_activity_one.py
├── publisher_name_standardisation_task_two_activity_two.py
│
├── dates_standardsation_error_log.txt
└── publisher_standardisation.log
```

---

# System Requirements

## Python Version

Developed using:

```text
Python 3.14.5
```

## Required Packages

Install dependencies using:

```bash
pip install pandas openpyxl python-dateutil rapidfuzz
```

### Package Usage

| Package         | Purpose                                             |
| --------------- | --------------------------------------------------- |
| pandas          | Reading, filtering and transforming Excel data      |
| openpyxl        | Excel file handling and workbook preservation       |
| python-dateutil | Flexible date parsing                               |
| rapidfuzz       | Fuzzy string matching for publisher standardisation |

---

# Script Documentation

## 1. value_filtering_task_one_activity_one.py

### Purpose

Filters repository metadata to retain only event-related outputs.

### Accepted Record Types

* conference_item
* exhibition
* performance

### Processing Workflow

1. Read metadata extract into a Pandas DataFrame.
2. Filter records using `isin()`.
3. Validate results by printing sample records and record count.
4. Export filtered records to a new Excel file.

### Output

```text
Outputs/metadata_extract_20260127_filtered_task_one_activity_one.xlsx
```

---

## 2. date_standardisation_task_one_activity_two.py

### Purpose

Standardises event dates into a consistent format.

### Input

```text
Outputs/metadata_extract_20260127_filtered_task_one_activity_one.xlsx
```

### Target Fields

* event_dates_start
* event_dates_end

### Methodology

#### Primary Parsing

Uses:

```python
pd.to_datetime()
```

with:

* day-first parsing
* strict validation
* year range checking

#### Secondary Parsing

Uses:

```python
dateutil.parser.parse()
```

with additional cleaning to remove ordinal suffixes such as:

```text
1st
2nd
3rd
4th
```

before attempting conversion.

#### Validation Rules

Dates are rejected when:

* Year < 1900
* Year > 2025
* Value cannot be parsed

### Error Handling

Invalid dates are recorded in:

```text
dates_standardsation_error_log.txt
```

This supports manual review of problematic records.

### Output

```text
Outputs/metadata_extract_20260127_standardised_dates_activity_one_task_two.xlsx
```

---

## 3. value_filtering_task_two_activity_one.py

### Purpose

Filters publication metadata prior to publisher standardisation.

### Accepted Record Types

* article
* conference_item

### Processing Workflow

1. Read metadata extract.
2. Filter publication records using `isin()`.
3. Export filtered dataset.

### Output

```text
Outputs/metadata_extract_20260127_filtered_task_two_activity_one.xlsx
```

---

## 4. publisher_name_standardisation_task_two_activity_two.py

### Purpose

Standardises publisher names through a combination of authority control rules and fuzzy matching techniques.

### Methodology

The standardisation process consists of five stages.

### Stage 1 – Hard-Coded Canonical Mapping

Known publisher variants are automatically normalised.

Examples include:

| Variant            | Canonical Form                                           |
| ------------------ | -------------------------------------------------------- |
| Taylor and Francis | Taylor & Francis                                         |
| IEEE               | Institute of Electrical and Electronics Engineers (IEEE) |
| BMC                | BioMed Central                                           |
| Elsevier Ltd       | Elsevier                                                 |
| Springer Nature    | Springer                                                 |
| Wiley-Blackwell    | Wiley                                                    |

---

### Stage 2 – Text Preprocessing

Publisher names are normalised before comparison.

Processing includes:

* Lowercasing
* Punctuation removal
* Tokenisation
* Stop-word removal
* Normalisation of "&" to "and"

Common publishing terms such as:

```text
press
group
ltd
limited
publications
books
journal
association
foundation
```

are excluded from similarity calculations.

---

### Stage 3 – Fuzzy Matching

The script uses RapidFuzz similarity metrics:

* `ratio()`
* `token_sort_ratio()`

to identify likely publisher variants.

Examples:

```text
Elsevier BV
Elsevier B.V.
Elsevier
```

can be recognised as equivalent publisher names.

---

### Stage 4 – Cluster Generation

Matched publisher names are grouped into clusters.

For each cluster:

* A canonical publisher name is selected.
* Frequency statistics are calculated.
* Variant mappings are recorded.

---

### Stage 5 – Manual Review Support

Publisher pairs with uncertain similarity scores are exported for human review rather than automatically merged.

This reduces the risk of incorrect standardisation decisions.

---

## Outputs

### Cleaned Metadata File

```text
Outputs/metadata_extract_20260127_publishers_cleaned.xlsx
```

Contains a standardised publisher field:

```text
publisher_standardised
```

### Review Index

```text
Outputs/publisher_review_index.xlsx
```

Contains publisher pairs requiring manual validation.

### Cluster Summary

```text
Outputs/publisher_cluster_summary.csv
```

Provides an audit trail of publisher clustering decisions.

### Processing Log

```text
publisher_standardisation.log
```

Records processing activity, matching decisions, warnings, and summary statistics.

---

# Data Quality Benefits

The workflow improves metadata quality by:

* Removing irrelevant records.
* Standardising event date formats.
* Reducing publisher name duplication.
* Supporting authority control practices.
* Providing audit trails for review and validation.
* Improving consistency for reporting and analytics.

---

# Future Enhancements

Potential future developments include:

* Author name standardisation implementation.
* Duplicate author detection.
* ORCID-based identity resolution.
* DOI validation.
* Configurable publisher authority files.
* Automated quality assurance reporting.
* Unit testing and continuous integration workflows.

---

## Author

Library Metadata Data Cleaning Activity

A Python-based workflow for research metadata cleansing, standardisation, and quality assurance.







