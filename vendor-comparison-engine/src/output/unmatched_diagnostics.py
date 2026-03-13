"""Build a diagnostics table for unmatched rows: best match from the other vendor and score."""
from __future__ import annotations

from typing import Dict, List

import pandas as pd
from rapidfuzz import fuzz

from src.processing.matcher import MatchRecord


def _best_score_and_row(
    desc: str,
    other_df: pd.DataFrame,
    desc_col: str = "description_norm",
) -> tuple[float, str, int]:
    best_score = -1.0
    best_desc = ""
    best_idx = -1
    for idx, row in other_df.iterrows():
        other_desc = row.get(desc_col) or ""
        if not isinstance(other_desc, str):
            other_desc = str(other_desc) if other_desc else ""
        score = max(
            fuzz.token_sort_ratio(desc, other_desc),
            fuzz.token_set_ratio(desc, other_desc),
        )
        if score > best_score:
            best_score = score
            best_desc = other_desc[:200]
            best_idx = int(idx)
    return best_score, best_desc, best_idx


def build_unmatched_diagnostics(
    matches: List[MatchRecord],
    vendor_a_sheets: Dict[str, pd.DataFrame],
    vendor_b_sheets: Dict[str, pd.DataFrame],
) -> pd.DataFrame:
    rows = []
    for m in matches:
        if m.method != "UNMATCHED":
            continue
        scope = m.scope_category
        a_sheet = m.vendor_a_sheet
        b_sheet = m.vendor_b_sheet
        a_df = vendor_a_sheets.get(a_sheet)
        b_df = vendor_b_sheets.get(b_sheet)
        if a_df is None or b_df is None:
            continue

        if m.vendor_a_idx is not None:
            vendor_side = "A"
            row_idx = m.vendor_a_idx
            r = a_df.loc[m.vendor_a_idx]
            desc = r.get("description_norm") or r.get("description") or ""
            item_id = r.get("item_id") or r.get("item_id_norm") or ""
            other_df = b_df
        else:
            vendor_side = "B"
            row_idx = m.vendor_b_idx
            r = b_df.loc[m.vendor_b_idx]
            desc = r.get("description_norm") or r.get("description") or ""
            item_id = r.get("item_id") or r.get("item_id_norm") or ""
            other_df = a_df

        if not isinstance(desc, str):
            desc = str(desc) if desc else ""
        best_score, best_desc, best_idx = _best_score_and_row(desc, other_df)

        rows.append({
            "Vendor": vendor_side,
            "Scope": scope,
            "Row_Index": row_idx,
            "Item_ID": item_id,
            "Description": (desc or "")[:200],
            "Best_Match_Score": round(best_score, 1),
            "Best_Match_Description": best_desc[:200],
            "Best_Match_Index": best_idx,
        })

    return pd.DataFrame(rows)
