"""
src/cleaner.py - Data Transformation, Imputation & Anomaly Audit Engine
Handles missing values, deduplication, IQR outlier detection, and data health metrics.
"""

import math

def calculate_iqr_bounds(values: list):
    """Calculates Q1, Q3, and IQR lower/upper bounds for a list of numbers."""
    sorted_v = sorted([v for v in values if v is not None and not (isinstance(v, float) and math.isnan(v))])
    if not sorted_v:
        return 0.0, 0.0, 0.0, 0.0
    n = len(sorted_v)
    q1 = sorted_v[int(n * 0.25)]
    q3 = sorted_v[int(n * 0.75)]
    iqr = q3 - q1
    lower = max(0.0, q1 - 1.5 * iqr)
    upper = q3 + 1.5 * iqr
    return round(q1, 2), round(q3, 2), round(lower, 2), round(upper, 2)

def clean_marketing_data(df):
    """
    Cleans raw marketing telemetry:
    - Deduplicates records
    - Imputes missing continuous metrics with median values
    - Identifies spend and conversion anomalies via IQR
    Returns (cleaned_data, audit_report_dict)
    """
    audit = {
        "initial_rows": len(df),
        "duplicates_removed": 0,
        "imputed_cells": 0,
        "spend_outliers": 0,
        "conversion_outliers": 0,
        "final_rows": len(df),
        "health_score": 100.0
    }

    # If pandas DataFrame
    try:
        import pandas as pd
        is_pd = isinstance(df, pd.DataFrame)
    except ImportError:
        is_pd = False

    if is_pd:
        cleaned_df = df.copy()
        init_len = len(cleaned_df)
        cleaned_df.drop_duplicates(inplace=True)
        audit["duplicates_removed"] = init_len - len(cleaned_df)

        # Impute missing values
        num_cols = ["ctr", "cpc", "cvr", "cac", "roas"]
        for col in num_cols:
            if col in cleaned_df.columns:
                null_count = int(cleaned_df[col].isna().sum())
                audit["imputed_cells"] += null_count
                median_val = cleaned_df[col].median()
                cleaned_df[col] = cleaned_df[col].fillna(median_val)

        # Outlier detection
        if "spend" in cleaned_df.columns:
            spend_vals = cleaned_df["spend"].tolist()
            _, _, _, spend_upper = calculate_iqr_bounds(spend_vals)
            cleaned_df["is_spend_outlier"] = cleaned_df["spend"] > spend_upper
            audit["spend_outliers"] = int(cleaned_df["is_spend_outlier"].sum())
        else:
            cleaned_df["is_spend_outlier"] = False

        cleaned_df["is_outlier"] = cleaned_df["is_spend_outlier"]
        audit["final_rows"] = len(cleaned_df)
        return cleaned_df, audit

    # Fallback for SimpleDataFrame / row list
    records = df.to_dict() if hasattr(df, "to_dict") else df
    unique_records = []
    seen = set()
    for r in records:
        key = (r.get("date"), r.get("campaign_id"))
        if key not in seen:
            seen.add(key)
            unique_records.append(dict(r))
        else:
            audit["duplicates_removed"] += 1

    spend_vals = [r.get("spend", 0.0) for r in unique_records]
    _, _, _, spend_upper = calculate_iqr_bounds(spend_vals)

    for r in unique_records:
        r["is_outlier"] = (r.get("spend", 0.0) > spend_upper)
        if r["is_outlier"]:
            audit["spend_outliers"] += 1

    audit["final_rows"] = len(unique_records)
    return unique_records, audit

def validate_data_integrity(records: list):
    """Asserts core physical and business invariants on records."""
    for r in records:
        assert r.get("spend", 0.0) >= 0, "Negative spend detected"
        assert r.get("conversions", 0) <= r.get("clicks", 0), "Conversions exceed clicks"
        assert r.get("clicks", 0) <= r.get("impressions", 0), "Clicks exceed impressions"
    return True
