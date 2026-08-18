# UC Library Search Batch Verification Tool (`check_uc_library.py`)

## Overview
This script automates the process of cross-referencing a large spreadsheet of books against the UC Berkeley Library Catalog (Primo VE). Instead of manually searching for hundreds of titles, the script reads an Excel file, queries the library's backend API, and strictly verifies the results. 

**This script is built on a "Zero False Positives" philosophy.** Because the UC Library catalog aggregates records from other campuses (Interlibrary Loans / Network Zones), this script forces a strict local-ownership check. If the script marks a book as "Found," it is guaranteed to be physically or electronically owned by UC Berkeley. Everything else is safely flagged as "Not Found" for human review.

## Key Features
* **Interactive Configuration Wizard:** No need to memorize complex command-line flags. Simply run the script, and a terminal wizard will walk you through selecting your files, priority filters, and format types.
* **Strict Ownership Verification (The Anti-ILL Filter):** The script digs into the raw JSON of the library's `delivery` blocks and ALMA fallback records to verify the presence of the `01UCS_BER` organizational code. It actively rejects exact matches if they belong to other UC campuses.
* **Smart Text Parsing:** Automatically strips cataloging fluff like `(ed.)`, `(eds.)`, and `(... Series)` from your spreadsheet's titles and authors so they don't break the exact-match requirements. 
* **Resume & Recheck Logic:** If the script is interrupted, it will automatically skip rows that already have a result. You can also force it to recheck previously verified "Yes" rows via the setup wizard.
* **Auto-Formatting Export:** The script dynamically generates the output Excel file, automatically widening columns and turning on text-wrapping so the final data is immediately readable without manual spreadsheet tweaking.

## Prerequisites
* **Python 3.6+**
* **Required Libraries:** `requests` (for handling API web traffic) and `openpyxl` (for reading, writing, and formatting Excel files).

Install dependencies via pip:

```bash
pip install requests openpyxl
```

## Usage
Run the script directly from your terminal. The interactive wizard will prompt you for everything it needs:

```bash
python3 check_uc_library.py
```

### Wizard Prompts
1. **Input File:** The name of your source Excel file.
2. **Output File:** The name for your results file.
3. **Priority Filter:** Target specific priorities (e.g., `1`, `5`, `1,2`, or `all`).
4. **Format Filter:** Target specific formats (e.g., `book`, `monograph`, or `all`).
5. **Recheck Rule:** Choose whether to skip or re-process rows that already have a result.

## How It Works
1. **Data Extraction:** The script reads the active sheet of the input Excel file and isolates rows that match your chosen Priority and Format filters.
2. **String Cleaning:** It removes trailing punctuation, normalizes cases, and strips out editor/series tags to create clean search tokens.
3. **API Query:** It sends a multi-clause query to the Primo VE `pnxs` endpoint to fetch potential catalog matches.
4. **Strict Match Validation:** The candidate record must be classified as a "book", the exact title phrase must match, and the primary author tokens must be a subset of the catalog's creator string.
5. **Ownership Gate:** The script verifies local ownership (`01UCS_BER`) via the brief delivery data or by directly querying the ALMA record endpoint.
6. **Data Export:** Results are saved continuously. The script appends five new columns directly to your spreadsheet: `UCLS_Lookup_Result`, `UCLS_Matched_Title`, `UCLS_Matched_Author`, `UCLS_Record_URL`, and `UCLS_Location`.
