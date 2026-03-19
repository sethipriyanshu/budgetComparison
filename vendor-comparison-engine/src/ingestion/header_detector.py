from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple

import pandas as pd


ANCHOR_KEYWORDS = {
    "description",
    "item",
    "scope",
    "price",
    "unit",
    "qty",
    "quantity",
    "total",
    "amount",
}

MIN_ANCHORS_FOR_HEADER = 3


@dataclass
class HeaderDetectionResult:
    header_row_index: int
    metadata_rows: pd.DataFrame
    data_with_headers: pd.DataFrame
    metadata: Dict[str, str]


def _count_anchor_hits(cells: Sequence[object]) -> int:
    hits = 0
    for value in cells:
        if value is None:
            continue
        text = str(value).strip().lower()
        if any(keyword in text for keyword in ANCHOR_KEYWORDS):
            hits += 1
    return hits


def detect_header_row(df: pd.DataFrame) -> Optional[int]:
    """
    Return the index of the header row in a raw sheet DataFrame, or None.

    We treat the first row that contains at least MIN_ANCHORS_FOR_HEADER occurrences
    of known keywords as the header row.
    """
    for idx, row in df.iterrows():
        hits = _count_anchor_hits(row.tolist())
        if hits >= MIN_ANCHORS_FOR_HEADER:
            return int(idx)
    # Fallback: treat the first row as header if we can't find anchors.
    # (apply_header_detection also falls back, but some unit tests call this
    # helper directly.)
    return 0


def _extract_metadata_rows(df: pd.DataFrame, header_row_index: int) -> pd.DataFrame:
    if header_row_index <= 0:
        return df.iloc[0:0]
    return df.iloc[:header_row_index]


def _guess_metadata(meta_df: pd.DataFrame) -> Dict[str, str]:
    """Very light heuristic metadata extraction from rows above header."""
    metadata: Dict[str, str] = {}

    text_lines: List[str] = []
    for _, row in meta_df.iterrows():
        parts = [str(v).strip() for v in row.tolist() if v is not None and str(v).strip()]
        if parts:
            text_lines.append(" ".join(parts))

    full_text = "\n".join(text_lines).lower()

    if "vendor" in full_text and "name" in full_text:
        metadata["vendor_name"] = "UNKNOWN_VENDOR"

    if "project" in full_text:
        metadata["project"] = "UNKNOWN_PROJECT"

    for symbol in ("$", "€", "£", "aed"):
        if symbol in full_text:
            metadata["currency_hint"] = symbol
            break

    return metadata


def apply_header_detection(df: pd.DataFrame) -> HeaderDetectionResult:
    """
    Run header detection and return cleaned DataFrame with header row applied.

    If no header row can be detected, we fall back to treating the first row
    as the header to avoid dropping data; downstream stages can still inspect
    and refine behavior.
    """
    header_idx = detect_header_row(df)
    if header_idx is None:
        header_idx = 0

    meta_rows = _extract_metadata_rows(df, header_idx)
    metadata = _guess_metadata(meta_rows)

    header_row = df.iloc[header_idx].tolist()
    data_rows = df.iloc[header_idx + 1 :].reset_index(drop=True)

    data_with_headers = data_rows.copy()
    data_with_headers.columns = header_row

    return HeaderDetectionResult(
        header_row_index=header_idx,
        metadata_rows=meta_rows,
        data_with_headers=data_with_headers,
        metadata=metadata,
    )

