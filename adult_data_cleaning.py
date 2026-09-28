"""
Week 1 Internship Project: Adult Census Income data cleaning and preprocessing.

Run in VS Code:
    python adult_data_cleaning.py

Expected input:
    adultfin.csv in the same folder as this script.

Outputs are written to the "outputs" folder.
"""

from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "adultfin.csv"
OUTPUT_DIR = BASE_DIR / "outputs"

CATEGORICAL_COLUMNS = [
    "workclass",
    "education",
    "marital_status",
    "occupation",
    "relationship",
    "race",
    "sex",
    "native_country",
    "income",
]

MISSING_CATEGORY_COLUMNS = ["workclass", "occupation", "native_country"]

NUMERIC_COLUMNS = [
    "age",
    "fnlwgt",
    "education_num",
    "capital_gain",
    "capital_loss",
    "hours_per_week",
]

OUTLIER_COLUMNS = [
    "age",
    "fnlwgt",
    "capital_gain",
    "capital_loss",
    "hours_per_week",
]


def snake_case_columns(columns: list[str]) -> list[str]:
    """Return consistent snake_case column names."""
    return [
        str(column).strip().lower().replace("-", "_").replace(" ", "_")
        for column in columns
    ]


def iqr_bounds(series: pd.Series) -> tuple[float, float]:
    """Calculate Tukey's 1.5 × IQR lower and upper fences."""
    q1 = float(series.quantile(0.25))
    q3 = float(series.quantile(0.75))
    iqr = q3 - q1
    return q1 - (1.5 * iqr), q3 + (1.5 * iqr)


def build_quality_table(
    raw: pd.DataFrame, cleaned: pd.DataFrame, missing_before: pd.Series
) -> pd.DataFrame:
    """Create a compact before/after data-quality audit table."""
    return pd.DataFrame(
        [
            {
                "metric": "Rows",
                "before": len(raw),
                "after": len(cleaned),
                "action": "Removed exact duplicate rows only",
            },
            {
                "metric": "Columns",
                "before": raw.shape[1],
                "after": cleaned.shape[1],
                "action": "Added engineered features and outlier flags",
            },
            {
                "metric": "Exact duplicate rows",
                "before": int(raw.duplicated().sum()),
                "after": int(cleaned.duplicated().sum()),
                "action": "Dropped duplicates",
            },
            {
                "metric": "Placeholder/missing values",
                "before": int(missing_before.sum()),
                "after": int(cleaned.isna().sum().sum()),
                "action": "Converted '?' to missing, then filled with 'Unknown'",
            },
        ]
    )


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input file not found: {INPUT_FILE}\n"
            "Place adultfin.csv in the same folder as this script."
        )

    # 1. DATA ACQUISITION AND INITIAL INSPECTION
    raw = pd.read_csv(INPUT_FILE)
    raw.columns = snake_case_columns(raw.columns.tolist())
    rows_loaded, columns_loaded = raw.shape

    expected_columns = set(NUMERIC_COLUMNS + CATEGORICAL_COLUMNS)
    missing_columns = expected_columns.difference(raw.columns)
    if missing_columns:
        raise ValueError(f"Required columns are missing: {sorted(missing_columns)}")

    # 2. STANDARDIZE TEXT AND EXPOSE HIDDEN MISSING VALUES
    cleaned = raw.copy()
    for column in CATEGORICAL_COLUMNS:
        cleaned[column] = cleaned[column].astype("string").str.strip()

    cleaned[CATEGORICAL_COLUMNS] = cleaned[CATEGORICAL_COLUMNS].replace(
        {"?": pd.NA, "": pd.NA}
    )
    missing_before = cleaned.isna().sum()

    # 3. REMOVE EXACT DUPLICATES
    duplicate_rows = int(cleaned.duplicated().sum())
    cleaned = cleaned.drop_duplicates().reset_index(drop=True)

    # 4. VALIDATE NUMERIC AND TARGET VALUES
    numeric_rules = {
        "age": cleaned["age"].between(17, 100),
        "fnlwgt": cleaned["fnlwgt"] > 0,
        "education_num": cleaned["education_num"].between(1, 16),
        "capital_gain": cleaned["capital_gain"] >= 0,
        "capital_loss": cleaned["capital_loss"] >= 0,
        "hours_per_week": cleaned["hours_per_week"].between(1, 99),
    }
    invalid_numeric_counts = {
        column: int((~valid_mask).sum())
        for column, valid_mask in numeric_rules.items()
    }

    valid_income = cleaned["income"].isin(["<=50K", ">50K"])
    invalid_income_count = int((~valid_income).sum())
    if any(invalid_numeric_counts.values()) or invalid_income_count:
        combined_valid_mask = valid_income.copy()
        for valid_mask in numeric_rules.values():
            combined_valid_mask &= valid_mask
        cleaned = cleaned.loc[combined_valid_mask].reset_index(drop=True)

    # Check that the two education fields agree one-to-one.
    education_mapping_check = (
        cleaned.groupby("education", dropna=False)["education_num"].nunique()
    )
    inconsistent_education_categories = int((education_mapping_check > 1).sum())

    # 5. HANDLE MISSING CATEGORIES
    # "Unknown" preserves all otherwise valid rows and makes missingness explicit.
    for column in MISSING_CATEGORY_COLUMNS:
        cleaned[column] = cleaned[column].fillna("Unknown")

    # Any unexpected missing category is also made explicit.
    for column in CATEGORICAL_COLUMNS:
        cleaned[column] = cleaned[column].fillna("Unknown")

    # 6. IDENTIFY OUTLIERS WITHOUT BLINDLY DELETING VALID PEOPLE
    outlier_records: list[dict[str, float | int | str]] = []
    for column in OUTLIER_COLUMNS:
        lower, upper = iqr_bounds(cleaned[column])
        flag_column = f"{column}_outlier_iqr"
        cleaned[flag_column] = (
            (cleaned[column] < lower) | (cleaned[column] > upper)
        ).astype("int8")
        outlier_records.append(
            {
                "column": column,
                "lower_fence": lower,
                "upper_fence": upper,
                "outlier_count": int(cleaned[flag_column].sum()),
                "outlier_percent": float(cleaned[flag_column].mean() * 100),
            }
        )
    outlier_summary = pd.DataFrame(outlier_records)

    # 7. ENGINEER ROBUST FEATURES
    # 1st/99th percentile capping reduces leverage while keeping every row.
    cap_columns = ["age", "fnlwgt", "hours_per_week"]
    cap_limits: dict[str, dict[str, float]] = {}
    for column in cap_columns:
        lower = float(cleaned[column].quantile(0.01))
        upper = float(cleaned[column].quantile(0.99))
        cleaned[f"{column}_capped"] = cleaned[column].clip(lower, upper)
        cap_limits[column] = {"p01": lower, "p99": upper}

    # Log transforms handle strongly right-skewed non-negative variables.
    cleaned["fnlwgt_log1p"] = np.log1p(cleaned["fnlwgt"])
    cleaned["capital_gain_log1p"] = np.log1p(cleaned["capital_gain"])
    cleaned["capital_loss_log1p"] = np.log1p(cleaned["capital_loss"])

    # Binary target for future machine-learning work.
    cleaned["income_high"] = cleaned["income"].map({"<=50K": 0, ">50K": 1})

    # 8. CREATE A MODEL-READY TABLE
    # Keep education_num and omit education to avoid duplicate encodings.
    model_numeric = pd.DataFrame(
        {
            "age": cleaned["age_capped"],
            "fnlwgt_log1p": cleaned["fnlwgt_log1p"],
            "education_num": cleaned["education_num"],
            "capital_gain_log1p": cleaned["capital_gain_log1p"],
            "capital_loss_log1p": cleaned["capital_loss_log1p"],
            "hours_per_week": cleaned["hours_per_week_capped"],
        }
    )

    # Median/IQR robust scaling limits the influence of extreme values.
    scaled_numeric = pd.DataFrame(index=model_numeric.index)
    scaling_parameters: dict[str, dict[str, float]] = {}
    for column in model_numeric.columns:
        median = float(model_numeric[column].median())
        q1 = float(model_numeric[column].quantile(0.25))
        q3 = float(model_numeric[column].quantile(0.75))
        iqr = q3 - q1
        divisor = iqr if iqr != 0 else 1.0
        scaled_numeric[column] = (model_numeric[column] - median) / divisor
        scaling_parameters[column] = {"median": median, "iqr": iqr}

    model_categories = [
        "workclass",
        "marital_status",
        "occupation",
        "relationship",
        "race",
        "sex",
        "native_country",
    ]
    encoded_categories = pd.get_dummies(
        cleaned[model_categories],
        prefix=model_categories,
        drop_first=False,
        dtype="int8",
    )
    model_ready = pd.concat(
        [
            scaled_numeric,
            encoded_categories,
            cleaned[["income_high"]].astype("int8"),
        ],
        axis=1,
    )

    # 9. SAVE OUTPUTS AND AUDIT MATERIALS
    quality_table = build_quality_table(raw, cleaned, missing_before)
    cleaned.to_csv(OUTPUT_DIR / "adult_cleaned.csv", index=False)
    model_ready.to_csv(OUTPUT_DIR / "adult_model_ready.csv", index=False)
    quality_table.to_csv(OUTPUT_DIR / "quality_summary.csv", index=False)
    outlier_summary.to_csv(OUTPUT_DIR / "outlier_summary.csv", index=False)

    metadata = {
        "input_file": INPUT_FILE.name,
        "rows_loaded": rows_loaded,
        "columns_loaded": columns_loaded,
        "duplicate_rows_removed": duplicate_rows,
        "rows_after_cleaning": len(cleaned),
        "columns_in_cleaned_file": cleaned.shape[1],
        "model_ready_shape": list(model_ready.shape),
        "placeholder_missing_counts": {
            column: int(missing_before[column])
            for column in MISSING_CATEGORY_COLUMNS
        },
        "invalid_numeric_counts": invalid_numeric_counts,
        "invalid_income_count": invalid_income_count,
        "inconsistent_education_categories": inconsistent_education_categories,
        "cap_limits": cap_limits,
        "scaling_parameters": scaling_parameters,
        "target_distribution": {
            str(key): int(value)
            for key, value in cleaned["income"].value_counts().items()
        },
    }
    (OUTPUT_DIR / "cleaning_metadata.json").write_text(
        json.dumps(metadata, indent=2), encoding="utf-8"
    )

    print("=" * 68)
    print("ADULT DATA CLEANING AND PREPROCESSING COMPLETE")
    print("=" * 68)
    print(f"Loaded shape:              {rows_loaded:,} rows × {columns_loaded} columns")
    print(f"Duplicates removed:        {duplicate_rows:,}")
    print(f"Cleaned shape:             {cleaned.shape[0]:,} rows × {cleaned.shape[1]} columns")
    print(
        "Hidden missing values:     "
        f"{int(missing_before[MISSING_CATEGORY_COLUMNS].sum()):,} cells"
    )
    print(
        "Rows with high income:     "
        f"{int(cleaned['income_high'].sum()):,} "
        f"({cleaned['income_high'].mean() * 100:.2f}%)"
    )
    print(
        f"Model-ready shape:         {model_ready.shape[0]:,} rows × "
        f"{model_ready.shape[1]} columns"
    )
    print(f"Outputs saved in:          {OUTPUT_DIR}")
    print("\nOutlier audit:")
    print(outlier_summary.to_string(index=False, formatters={
        "lower_fence": "{:.2f}".format,
        "upper_fence": "{:.2f}".format,
        "outlier_percent": "{:.2f}".format,
    }))


if __name__ == "__main__":
    main()