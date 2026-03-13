from __future__ import annotations

from typing import List

import pandas as pd

from src.processing.matcher import MatchRecord


def build_audit_trail(matches: List[MatchRecord]) -> pd.DataFrame:
    """
    Convert a list of MatchRecord objects into a flat DataFrame for the
    Audit Trail sheet and/or JSON/CSV export.
    """
    rows = []
    for m in matches:
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
            }
        )
    return pd.DataFrame(rows)

