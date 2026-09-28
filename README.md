# Week 1 Internship Project

## Adult Census Income: Data Acquisition, Cleaning, and Preprocessing

### Files

- `adult_data_cleaning.py` — complete, commented Python program
- `adultfin.csv` — input dataset
- `Week1_Adult_Data_Preprocessing_Report.docx` — final internship report
- `outputs/` — generated cleaned files and audit summaries

### Run in VS Code

1. Install Python 3.10 or newer.
2. Open this folder in VS Code.
3. Open **Terminal > New Terminal**.
4. Optional but recommended:

   **Windows**
   ```powershell
   python -m venv .venv
   .venv\Scripts\activate
   ```

   **macOS/Linux**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

5. Install the packages:
   ```bash
   pip install -r requirements.txt
   ```
6. Run:
   ```bash
   python adult_data_cleaning.py
   ```

### Generated outputs

- `adult_cleaned.csv` — readable cleaned dataset with explicit missing categories,
  outlier flags, capped features, log features, and a binary target.
- `adult_model_ready.csv` — robust-scaled numeric features, one-hot encoded
  categories, and the target.
- `quality_summary.csv` — before/after quality checks.
- `outlier_summary.csv` — IQR fences and counts.
- `cleaning_metadata.json` — reproducible thresholds and audit metadata.

### Notes

- The script does not delete statistically unusual but plausible people.
- `?` placeholders are treated as missing and changed to `Unknown`.
- Exact duplicate rows are removed.
- For a future prediction model, fit capping/scaling/encoding rules on the
  training split only to avoid data leakage.