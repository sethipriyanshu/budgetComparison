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


# Keys that are metadata (sheet/section name, vendor) — never use as item name
_SKIP_KEYS = ("item_id", "item_id_norm", "scope_category", "vendor_id")
# Values that are section/sheet names, not line descriptions
_SECTION_NAME_VALUES = frozenset(("mechanics", "steelwork", "electrics", "overall", "options"))

# Item name phrases that indicate a total/subtotal row (exclude from comparison output)
_TOTAL_LIKE_PHRASES = ("total price", "subtotal", "sum ", " total", "section total")


def _get_item_name(row: pd.Series) -> str:
    """Get the best available description text from a row; avoid item ref and sheet/section names."""
    # Prefer explicit 'description' column (if it holds real text, not a section name)
    val = row.get("description")
    if val is not None and str(val).strip() and not _looks_like_item_ref(val):
        s = str(val).strip()
        if s.lower() not in _SECTION_NAME_VALUES:
            return s

    # Fallback: columns that look like description content (exclude metadata and _norm)
    for key in row.index:
        if key in _SKIP_KEYS or "_norm" in str(key):
            continue
        key_lower = str(key).lower()
        if not any(kw in key_lower for kw in _DESC_KEYWORDS):
            continue
        # Skip if key is exactly scope_category-like (redundant check)
        if key_lower == "scope_category" or key_lower == "vendor_id":
            continue
        val = row.get(key)
        if val is None or (isinstance(val, float) and pd.isna(val)):
            continue
        s = str(val).strip()
        if not s or _looks_like_item_ref(s) or s.lower() in _SECTION_NAME_VALUES:
            continue
        return s

    return ""


def _is_total_like_row(item_name: str) -> bool:
    """True if this looks like a total/subtotal label row (exclude from comparison output)."""
    if not item_name or not isinstance(item_name, str):
        return False
    text = item_name.strip().lower()
    return any(phrase in text for phrase in _TOTAL_LIKE_PHRASES)


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
        # Only use real description; leave blank if none (don't use item ref as "name")
        item_name = _get_item_name(a) or _get_item_name(b) or ""

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

        # Show blank instead of 0 for prices/deltas when there's no real quote (cleaner for decision-making)
        def _or_blank(val):
            if val is None:
                return None
            try:
                if float(val) == 0:
                    return None
            except (TypeError, ValueError):
                pass
            return val

        # Skip phantom rows: no item name and no totals on either side (don't show "items that do not exist")
        if not (item_name or _or_blank(total_a) is not None or _or_blank(total_b) is not None):
            continue

        # Skip total/subtotal label rows (e.g. "Total Price:") so they don't appear as extra lines
        if _is_total_like_row(item_name):
            continue

        # Comments: combine A and B comments when present
        comment_a = a.get("comments")
        comment_b = b.get("comments")
        comment_a_str = "" if comment_a is None or (isinstance(comment_a, float) and pd.isna(comment_a)) else str(comment_a).strip()
        comment_b_str = "" if comment_b is None or (isinstance(comment_b, float) and pd.isna(comment_b)) else str(comment_b).strip()
        parts = []
        if comment_a_str:
            parts.append(f"A: {comment_a_str}")
        if comment_b_str:
            parts.append(f"B: {comment_b_str}")
        comments_cell = " | ".join(parts) if parts else ""

        rows.append(
            [
                item_id,
                item_name,
                qty_a,
                _or_blank(unit_price_a),
                _or_blank(total_a),
                qty_b,
                _or_blank(unit_price_b),
                _or_blank(total_b),
                _or_blank(delta_abs),
                _or_blank(delta_pct) if delta_pct is not None else None,
                comments_cell,
            ]
        )

    return rows

