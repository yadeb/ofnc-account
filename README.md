# OFNC Gift Aid Processing Project

This project provides tools for processing, cleaning, and matching financial data (such as bank statements and member lists) for the Overseas Fellowship of Nigerian Christians (OFNC) Gift Aid claims. The codebase is written in Python and is designed to automate the extraction, transformation, and reconciliation of donation records with member consent lists.

## Main Components

### 1. `process_ga.py`
- **Purpose:** Main entry point for processing consolidated income data and matching payments to members.
- **Key Features:**
  - Command-line interface for specifying input folder and files.
  - `GAProcessor` class encapsulates all logic for data loading, cleaning, matching, and output.
  - Uses fuzzy matching to associate bank statement entries with member records.
  - Outputs processed Excel files for further use or reporting.

### 2. `income_data_loader.py`
- **Purpose:** Loads and parses income data from Excel files with flexible header detection.
- **Key Features:**
  - `IncomeDataLoader` class can find the correct header row in various Excel sheets.
  - Handles multiple sheets and extracts relevant dataframes for each branch.
  - Designed to be robust to changes in Excel file structure.

### 3. `data_loader.py`
- **Purpose:** General-purpose loader for bank statements and member lists.
- **Key Features:**
  - `DataLoader` class loads and preprocesses both bank statements and member lists.
  - Handles missing files and basic data cleaning.

### 4. `match_payments.py`
- **Purpose:** Standalone script for matching bank statement entries to members and summarizing payments.
- **Key Features:**
  - Uses fuzzy matching to link payment descriptions to member names.
  - Outputs summary Excel files of matched payments.

### 5. `OFNCAccount.py`
- **Purpose:** (Legacy/experimental) Functions for extracting and matching names from bank statements to CQ (central records) templates.
- **Key Features:**
  - Contains logic for parsing transaction descriptions and matching to CQ records.

### 6. `ga_script.py`
- **Purpose:** Minimal script for loading a specific Excel sheet as a DataFrame.

## Testing
- Tests are provided in `test_ga.py` (unittest) and `test/income_data_loader_test.py` (pytest).
- Tests cover data loading, header detection, and error handling.

## Usage

### Command-Line Example
```sh
python process_ga.py --folder /path/to/folder --account_file ConsolidatedAccounts2024Final1_GA.xlsx --consent_file ga_consent_list.xlsx
```
or
```sh
/Users/yadebisi/code/OFNC_GA/.venv/bin/python /Users/yadebisi/code/OFNC_GA/process_ga.py
```

### Output
- Processed Excel files with matched and cleaned donation/member data.

## Requirements
- Python 3.8+
- pandas
- rapidfuzz
- xlsxwriter

Install dependencies with:
```sh
pip install -r requirements.txt
```

## Project Structure
- `process_ga.py` — Main processing logic and CLI
- `income_data_loader.py` — Flexible Excel loader
- `data_loader.py` — General data loading utilities
- `match_payments.py` — Standalone payment matching
- `OFNCAccount.py` — CQ matching utilities
- `ga_script.py` — Minimal loader
- `test_ga.py`, `test/income_data_loader_test.py` — Tests

## License
This project is for internal OFNC use. Contact the maintainers for licensing details.
