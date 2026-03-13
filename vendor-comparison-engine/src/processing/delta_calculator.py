from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional

import pandas as pd

from .matcher import MatchRecord


@dataclass
class DeltaResult:
    matches_with_deltas: pd.DataFrame
    unmatched: pd.DataFrame


def _safe_price_delta(
    a: Optional[float], b: Optional[float]
) -> tuple[Optional[float], Optional[float]]:
    if a is None or b is None:
        return None, None
    try:
        delta_abs = b - a
        if a == 0:
            delta_pct = None
        else:
            delta_pct = delta_abs / a
        return delta_abs, delta_pct
    except Exception:
        return None, None


def calculate_deltas(
    matches: List[MatchRecord],
    vendor_a_sheets: Dict[str, pd.DataFrame],
    vendor_b_sheets: Dict[str, pd.DataFrame],
) -> DeltaResult:
    rows = []
    unmatched_rows = []

    for m in matches:
        a_df = vendor_a_sheets.get(m.vendor_a_sheet)
        b_df = vendor_b_sheets.get(m.vendor_b_sheet)

        a_row = (
            a_df.loc[m.vendor_a_idx] if a_df is not None and m.vendor_a_idx is not None else None
        )
        b_row = (
            b_df.loc[m.vendor_b_idx] if b_df is not None and m.vendor_b_idx is not None else None
        )

        if a_row is None or b_row is None:
            # Unmatched on one side
            unmatched_rows.append(
                {
                    "scope_category": m.scope_category,
                    "vendor_a_sheet": m.vendor_a_sheet,
                    "vendor_b_sheet": m.vendor_b_sheet,
                    "vendor_a_idx": m.vendor_a_idx,
                    "vendor_b_idx": m.vendor_b_idx,
                    "method": m.method,
                    "confidence": m.confidence,
                }
            )
            continue

        a_total = a_row.get("total_price")
        b_total = b_row.get("total_price")
        delta_abs, delta_pct = _safe_price_delta(a_total, b_total)

        qty_mismatch = False
        if a_df is not None and b_df is not None and "qty" in a_df.columns and "qty" in b_df.columns:
            qty_mismatch = a_row.get("qty") != b_row.get("qty")

        rows.append(
            {
                "scope_category": m.scope_category,
                "vendor_a_sheet": m.vendor_a_sheet,
                "vendor_b_sheet": m.vendor_b_sheet,
                "vendor_a_idx": m.vendor_a_idx,
                "vendor_b_idx": m.vendor_b_idx,
                "method": m.method,
                "confidence": m.confidence,
                "score": m.score,
                "vendor_a_total": a_total,
                "vendor_b_total": b_total,
                "price_delta_abs": delta_abs,
                "price_delta_pct": delta_pct,
                "qty_mismatch": qty_mismatch,
            }
        )

    return DeltaResult(
        matches_with_deltas=pd.DataFrame(rows),
        unmatched=pd.DataFrame(unmatched_rows),
    )

