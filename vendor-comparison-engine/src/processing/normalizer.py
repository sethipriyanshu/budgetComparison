from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List

import pandas as pd


NUMERIC_KEYS = {
    "qty",
    "unit_price_hardware",
    "unit_price_service",
    "unit_price_sum",
    "total_price",
}


@dataclass
class NormalizedSheet:
    scope_category: str
    vendor_id: str
    dataframe: pd.DataFrame


def _to_float(value) -> float | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    # Remove thousands separators, keep decimal point
    text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def normalize_sheet(
    df: pd.DataFrame,
    column_mapping: Dict[str, str],
    scope_category: str,
    vendor_id: str,
) -> NormalizedSheet:
    """
    Apply basic Phase 1 normalization:
    - Rename vendor columns to internal keys using column_mapping.
    - Normalize numeric fields to floats.
    - Trim text fields and add lowercased helper columns for later matching.
    - Attach scope_category and vendor_id columns.
    """
    # Rename columns to internal keys where mappings exist, keep others as-is.
    renamed_cols = {
        original: column_mapping.get(original, original)
        for original in df.columns
    }
    df_norm = df.rename(columns=renamed_cols).copy()

    # Normalize numeric fields
    for key in NUMERIC_KEYS:
        if key in df_norm.columns:
            df_norm[key] = df_norm[key].map(_to_float)

    # Normalize text-ish fields: create helper "<col>_norm" lowercased copies.
    # We treat all non-numeric-key columns as text-like, even if they sometimes hold numbers.
    # Use column *position* to avoid surprises when there are duplicate column names.
    for idx, col in enumerate(list(df_norm.columns)):
        if col in NUMERIC_KEYS:
            continue

        base_name = str(col) if col not in (None, "") else f"col_{idx}"
        norm_col = f"{base_name}_norm"

        series = df_norm.iloc[:, idx]
        df_norm[norm_col] = series.map(
            lambda v: ""
            if v is None or (isinstance(v, float) and pd.isna(v))
            else str(v).strip().lower()
        )

    df_norm["scope_category"] = scope_category
    df_norm["vendor_id"] = vendor_id

    # Row type: OPTION (alternatives), SUBTOTAL (total price / sum rows — not line items), NORMAL
    desc_norm_col = "description_norm"
    if desc_norm_col in df_norm.columns:
        def classify_row(desc: str) -> str:
            if not isinstance(desc, str):
                return "NORMAL"
            text = desc.lower()
            option_keywords = ["option", "alternative", "alt", "instead of", "optional"]
            if any(k in text for k in option_keywords):
                return "OPTION"
            total_keywords = ["total price", "subtotal", "sum ", " total", "section total"]
            if any(k in text for k in total_keywords):
                return "SUBTOTAL"
            return "NORMAL"

        df_norm["row_type"] = df_norm[desc_norm_col].map(classify_row)
    else:
        df_norm["row_type"] = "NORMAL"

    return NormalizedSheet(scope_category=scope_category, vendor_id=vendor_id, dataframe=df_norm)

