# UC Library Search Batch Verification Tool (`check_uc_library.py`)

## Overview
This script automates the process of cross-referencing a large spreadsheet of books against the UC Berkeley Library Catalog (Primo VE). Instead of manually searching for hundreds of titles, the script reads an Excel file, queries the library's backend API, strictly verifies the results to prevent false positives, and outputs a new Excel file with the lookup results, match scores, and direct links to the library records.

## Key Features
* **Automated Batch Processing:** Reads `Title` and `Author / Editor` columns from an Excel (`.xlsx`) spreadsheet and processes them sequentially.
* **The "Python Bouncer" (Strict Validation):** Bypasses unreliable exact-match API filters. It uses a broad API query to retrieve candidate books, then uses strict local Python logic to ensure the exact title phrase and primary author match before flagging a book as "Found."
* **Smart Author Parsing:** Automatically strips out cataloging fluff (e.g., "translated by", "introduction by", "edited by") to accurately match the primary author's surname against the library's metadata.
* **Manual Review Flagging:** If an author string is highly complex (containing multiple co-authors or editor roles), the script flags the row for a quick human review.

## Prerequisites
* **Python 3.6+**
* **Required Libraries:** `requests` (for handling API web traffic) and `openpyxl` (for reading and writing Excel files).

Install dependencies via pip:

```bash
pip install requests openpyxl
```

## Usage
Run the script from your terminal or command prompt by providing an input filename and your desired output filename:

```bash
python3 check_uc_library.py input_file.xlsx output_file.xlsx
```

### Optional Arguments
* **`--sheet`:** Specify a specific sheet name (defaults to Sheet1).
* **`--delay`:** Time in seconds to wait between API calls to prevent rate-limiting (defaults to 1.0).
* **`--limit`:** Stop the script after processing a specific number of rows (useful for testing).

## How It Works
1. **Data Extraction:** The script reads the input Excel file and isolates rows that have both a Title and an Author.
2. **API Query:** It sends a contains query to the Primo VE pnxs endpoint, restricted to the `Default_UCLibrarySearch` tab.
3. **Title Check:** The exact phrase from the spreadsheet must exist word-for-word inside the library's title string (Score: 100 or 0).
4. **Author Check:** The primary author's surname must be a subset of the library's creator string, or it falls back to a similarity ratio.
5. **Threshold validation:** If the title fails the strict phrase check or the author score falls below 60%, the combined score drops to 0, and the book is marked as Not Found.
6. **Data Export:** Results are saved dynamically. The script will generate six new columns in the output Excel file: `UCLS_Lookup_Result`, `UCLS_Match_Score`, `UCLS_Matched_Title`, `UCLS_Matched_Author`, `UCLS_Record_URL`, and `UCLS_Review_Note`.
