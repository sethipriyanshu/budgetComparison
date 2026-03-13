from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd
from rapidfuzz import fuzz


@dataclass
class MatchRecord:
    scope_category: str
    vendor_a_sheet: str
    vendor_b_sheet: str
    vendor_a_idx: Optional[int]
    vendor_b_idx: Optional[int]
    method: str
    confidence: str
    score: float


def _best_fuzzy_match(
    a_desc: str,
    b_candidates: pd.Series,
    used_b_indices: set[int],
) -> Tuple[Optional[int], float]:
    """Use best of token_sort_ratio and token_set_ratio so partial/similar wording still matches."""
    best_idx: Optional[int] = None
    best_score: float = -1.0
    for idx, b_desc in b_candidates.items():
        if idx in used_b_indices:
            continue
        b_str = b_desc if isinstance(b_desc, str) else (str(b_desc) if b_desc else "")
        if not b_str:
            continue
        score = max(
            fuzz.token_sort_ratio(a_desc, b_str),
            fuzz.token_set_ratio(a_desc, b_str),
        )
        if score > best_score:
            best_score = score
            best_idx = int(idx)
    return best_idx, best_score


def match_scope(
    scope: str,
    vendor_a_df: pd.DataFrame,
    vendor_b_df: pd.DataFrame,
    auto_threshold: float,
    review_threshold: float,
    vendor_a_sheet_name: str,
    vendor_b_sheet_name: str,
    manual_matches: Optional[List[Dict]] = None,
    use_position_fallback: bool = False,
    position_min_score: float = 50.0,
) -> List[MatchRecord]:
    """
    Match items between Vendor A and Vendor B for a single scope category.
    """
    matches: List[MatchRecord] = []
    used_b_indices: set[int] = set()

    # 1. Apply manual matches first
    if manual_matches:
        for m in manual_matches:
            used_b_indices.add(m["vendor_b_idx"])
            matches.append(
                MatchRecord(
                    scope_category=scope,
                    vendor_a_sheet=vendor_a_sheet_name,
                    vendor_b_sheet=vendor_b_sheet_name,
                    vendor_a_idx=m["vendor_a_idx"],
                    vendor_b_idx=m["vendor_b_idx"],
                    method=m["method"],
                    confidence=m["confidence"],
                    score=m["score"],
                )
            )

    # Build views for quick lookups
    a_remaining = vendor_a_df.drop(index=[m.vendor_a_idx for m in matches if m.vendor_a_idx is not None])

    # 2. Exact item_id match (use _norm when present so "1.1" and " 1.1 " match)
    a_id_col = "item_id_norm" if "item_id_norm" in vendor_a_df.columns else "item_id"
    b_id_col = "item_id_norm" if "item_id_norm" in vendor_b_df.columns else "item_id"
    if "item_id" in vendor_a_df.columns and "item_id" in vendor_b_df.columns:
        for idx, row in a_remaining.iterrows():
            a_id = row.get(a_id_col) or row.get("item_id")
            if a_id is None:
                continue
            a_id_val = str(a_id).strip().lower()
            if not a_id_val:
                continue
            candidate_b = vendor_b_df.index[
                (vendor_b_df[b_id_col].fillna("").astype(str).str.strip().str.lower() == a_id_val)
            ]
            for b_idx in candidate_b:
                if b_idx in used_b_indices:
                    continue
                matches.append(
                    MatchRecord(
                        scope_category=scope,
                        vendor_a_sheet=vendor_a_sheet_name,
                        vendor_b_sheet=vendor_b_sheet_name,
                        vendor_a_idx=int(idx),
                        vendor_b_idx=int(b_idx),
                        method="ITEM_ID_EXACT",
                        confidence="AUTO",
                        score=100.0,
                    )
                )
                used_b_indices.add(int(b_idx))
                break

        a_remaining = vendor_a_df.drop(
            index=[m.vendor_a_idx for m in matches if m.method == "ITEM_ID_EXACT"]
        )

    # 3. Exact description match (normalized)
    if "description_norm" in vendor_a_df.columns and "description_norm" in vendor_b_df.columns:
        for idx, row in a_remaining.iterrows():
            desc = row["description_norm"]
            if not desc:
                continue
            candidate_b = vendor_b_df.index[
                (vendor_b_df["description_norm"] == desc)
            ]
            for b_idx in candidate_b:
                if b_idx in used_b_indices:
                    continue
                matches.append(
                    MatchRecord(
                        scope_category=scope,
                        vendor_a_sheet=vendor_a_sheet_name,
                        vendor_b_sheet=vendor_b_sheet_name,
                        vendor_a_idx=int(idx),
                        vendor_b_idx=int(b_idx),
                        method="DESCRIPTION_EXACT",
                        confidence="AUTO",
                        score=99.0,
                    )
                )
                used_b_indices.add(int(b_idx))
                break

        a_remaining = vendor_a_df.drop(
            index=[m.vendor_a_idx for m in matches if m.method == "DESCRIPTION_EXACT"]
        )

    # 4 & 5. Fuzzy description matching
    if "description_norm" in vendor_a_df.columns and "description_norm" in vendor_b_df.columns:
        b_desc_series = vendor_b_df["description_norm"]

        for idx, row in a_remaining.iterrows():
            a_desc = row["description_norm"]
            if not a_desc:
                continue

            b_idx, score = _best_fuzzy_match(a_desc, b_desc_series, used_b_indices)
            if b_idx is None:
                continue

            if score >= auto_threshold:
                confidence = "AUTO"
                method = "FUZZY_AUTO"
            elif score >= review_threshold:
                confidence = "REVIEW"
                method = "FUZZY_REVIEW"
            else:
                continue

            matches.append(
                MatchRecord(
                    scope_category=scope,
                    vendor_a_sheet=vendor_a_sheet_name,
                    vendor_b_sheet=vendor_b_sheet_name,
                    vendor_a_idx=int(idx),
                    vendor_b_idx=int(b_idx),
                    method=method,
                    confidence=confidence,
                    score=float(score),
                )
            )
            used_b_indices.add(int(b_idx))

        a_remaining = vendor_a_df.drop(
            index=[m.vendor_a_idx for m in matches if m.method.startswith("FUZZY")]
        )

    # 5b. Optional: position-based fallback (pair 1st remaining A with 1st remaining B, etc.) when similarity >= position_min_score
    if use_position_fallback and "description_norm" in vendor_a_df.columns and "description_norm" in vendor_b_df.columns:
        a_idx_list = sorted(a_remaining.index.tolist())
        b_remaining_idx = sorted(set(vendor_b_df.index) - used_b_indices)
        b_desc_series = vendor_b_df["description_norm"]
        for k, a_idx in enumerate(a_idx_list):
            if k >= len(b_remaining_idx):
                break
            b_idx = b_remaining_idx[k]
            a_desc = vendor_a_df.loc[a_idx, "description_norm"] if isinstance(vendor_a_df.loc[a_idx, "description_norm"], str) else ""
            b_desc = b_desc_series.loc[b_idx] if isinstance(b_desc_series.loc[b_idx], str) else ""
            if not a_desc and not b_desc:
                score = 50.0
            else:
                score = max(
                    fuzz.token_sort_ratio(a_desc or "", b_desc or ""),
                    fuzz.token_set_ratio(a_desc or "", b_desc or ""),
                )
            if score < position_min_score:
                continue
            matches.append(
                MatchRecord(
                    scope_category=scope,
                    vendor_a_sheet=vendor_a_sheet_name,
                    vendor_b_sheet=vendor_b_sheet_name,
                    vendor_a_idx=int(a_idx),
                    vendor_b_idx=int(b_idx),
                    method="POSITION_FALLBACK",
                    confidence="REVIEW",
                    score=float(score),
                )
            )
            used_b_indices.add(int(b_idx))
        a_remaining = vendor_a_df.drop(
            index=[m.vendor_a_idx for m in matches if m.method == "POSITION_FALLBACK"]
        )

    # 6. Remaining unmatched Vendor A rows
    for idx in a_remaining.index:
        matches.append(
            MatchRecord(
                scope_category=scope,
                vendor_a_sheet=vendor_a_sheet_name,
                vendor_b_sheet=vendor_b_sheet_name,
                vendor_a_idx=int(idx),
                vendor_b_idx=None,
                method="UNMATCHED",
                confidence="UNMATCHED",
                score=0.0,
            )
        )

    # 7. Unmatched Vendor B rows (those whose index not in used_b_indices)
    for b_idx in vendor_b_df.index:
        if b_idx in used_b_indices:
            continue
        matches.append(
            MatchRecord(
                scope_category=scope,
                vendor_a_sheet=vendor_a_sheet_name,
                vendor_b_sheet=vendor_b_sheet_name,
                vendor_a_idx=None,
                vendor_b_idx=int(b_idx),
                method="UNMATCHED",
                confidence="UNMATCHED",
                score=0.0,
            )
        )

    return matches

