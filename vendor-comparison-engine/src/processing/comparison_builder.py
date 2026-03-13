from __future__ import annotations

import re
from typing import Dict, List, Any

import pandas as pd

from src.processing.matcher import MatchRecord

# Column names that typically hold the line description (not item ref)
_DESC_KEYWORDS = ("scope", "description", "item", "service", "work package", "name")


def _looks_like_item_ref(value: Any) -> bool:
    if value is None:
        return True
    s = str(value).strip()
    if not s:
        return True
    # e.g. 1, 1.1, 2.3.4
    if re.match(r"^\d+(\.\d+)*$", s):
        return True
    if len(s) <= 4 and s.replace(".", "").isdigit():
        return True
    return False


def _get_item_name(row: pd.Series) -> str:
    """Get the best available description text from a row; avoid using item ref (e.g. 1.1) as name."""
    # Prefer explicit 'description' column
    val = row.get("description")
    if val is not None and str(val).strip() and not _looks_like_item_ref(val):
        return str(val).strip()

    # Fallback: any column whose name suggests description content
    for key in row.index:
        if key in ("item_id", "item_id_norm") or "_norm" in str(key):
            continue
        key_lower = str(key).lower()
        if not any(kw in key_lower for kw in _DESC_KEYWORDS):
            continue
        val = row.get(key)
        if val is None or (isinstance(val, float) and pd.isna(val)):
            continue
        s = str(val).strip()
        if s and not _looks_like_item_ref(s):
            return s

    return ""


def build_comparison_rows_for_scope(
    scope: str,
    matches_with_deltas: pd.DataFrame,
    vendor_a_sheets: Dict[str, pd.DataFrame],
    vendor_b_sheets: Dict[str, pd.DataFrame],
) -> List[List[object]]:
    """
    Build row data for a single scope's comparison sheet (decision-focused: item name, quantities, prices, deltas).
    """
    rows: List[List[object]] = []

    scope_matches = matches_with_deltas[
        (matches_with_deltas["scope_category"] == scope)
        & matches_with_deltas["vendor_a_idx"].notna()
        & matches_with_deltas["vendor_b_idx"].notna()
    ]

    for _, m_row in scope_matches.iterrows():
        a_sheet_name = m_row.get("vendor_a_sheet")
        b_sheet_name = m_row.get("vendor_b_sheet")
        a_idx = int(m_row["vendor_a_idx"])
        b_idx = int(m_row["vendor_b_idx"])

        a_df = vendor_a_sheets.get(a_sheet_name)
        b_df = vendor_b_sheets.get(b_sheet_name)
        if a_df is None or b_df is None:
            continue

        a = a_df.loc[a_idx]
        b = b_df.loc[b_idx]

        item_id = a.get("item_id") or b.get("item_id")
        # Resolve real description from whichever column holds it (e.g. "Scope - Mechanics")
        item_name = _get_item_name(a) or _get_item_name(b)
        if not item_name and item_id is not None:
            item_name = str(item_id)

        qty_a = a.get("qty")
        qty_b = b.get("qty")

        unit_price_a = (
            a.get("unit_price_sum")
            or a.get("unit_price_hardware")
            or a.get("unit_price_service")
        )
        unit_price_b = (
            b.get("unit_price_sum")
            or b.get("unit_price_hardware")
            or b.get("unit_price_service")
        )

        total_a = a.get("total_price")
        total_b = b.get("total_price")

        delta_abs = m_row.get("price_delta_abs")
        delta_pct = m_row.get("price_delta_pct")

        rows.append(
            [
                item_id,
                item_name,
                qty_a,
                unit_price_a,
                total_a,
                qty_b,
                unit_price_b,
                total_b,
                delta_abs,
                delta_pct,
            ]
        )

    return rows

