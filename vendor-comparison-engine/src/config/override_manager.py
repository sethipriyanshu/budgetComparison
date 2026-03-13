from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List

import yaml
import pandas as pd


@dataclass
class OverrideRule:
    vendor_a_description: str
    vendor_b_description: str


def load_overrides(path: str | Path) -> List[OverrideRule]:
    p = Path(path)
    if not p.exists():
        return []

    with open(p, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or []

    rules: List[OverrideRule] = []
    for entry in data:
        if not isinstance(entry, dict):
            continue
        if not entry.get("force_match"):
            continue
        a = str(entry.get("vendor_a_description", "")).strip().lower()
        b = str(entry.get("vendor_b_description", "")).strip().lower()
        if not a or not b:
            continue
        rules.append(OverrideRule(vendor_a_description=a, vendor_b_description=b))
    return rules


def apply_overrides(
    vendor_a_df: pd.DataFrame,
    vendor_b_df: pd.DataFrame,
    rules: List[OverrideRule],
) -> List[Dict[str, Any]]:
    """
    Apply manual overrides and return a list of forced matches.

    Each forced match is a dict with:
      - vendor_a_idx
      - vendor_b_idx
      - method = "MANUAL_OVERRIDE"
      - confidence = "MANUAL"
    """
    matches: List[Dict[str, Any]] = []

    if not rules:
        return matches

    # We rely on the presence of description_norm columns created in normalization.
    a_desc_col = "description_norm"
    b_desc_col = "description_norm"
    if a_desc_col not in vendor_a_df.columns or b_desc_col not in vendor_b_df.columns:
        return matches

    used_b_indices: set[int] = set()

    for rule in rules:
        a_candidates = vendor_a_df.index[
            vendor_a_df[a_desc_col] == rule.vendor_a_description
        ]
        b_candidates = vendor_b_df.index[
            vendor_b_df[b_desc_col] == rule.vendor_b_description
        ]

        if not len(a_candidates) or not len(b_candidates):
            continue

        # Take first available indices that are not already used on B side.
        for a_idx in a_candidates:
            for b_idx in b_candidates:
                if b_idx in used_b_indices:
                    continue
                matches.append(
                    {
                        "vendor_a_idx": int(a_idx),
                        "vendor_b_idx": int(b_idx),
                        "method": "MANUAL_OVERRIDE",
                        "confidence": "MANUAL",
                        "score": 100.0,
                    }
                )
                used_b_indices.add(int(b_idx))
                break

    return matches

