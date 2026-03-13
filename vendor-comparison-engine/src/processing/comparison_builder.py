from __future__ import annotations

from typing import Dict, List

import pandas as pd

from src.processing.matcher import MatchRecord


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
        description = a.get("description") or b.get("description")
        # Single clear "item name" for decision-making
        item_name = description if description else (str(item_id) if item_id else "")

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

