from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, Tuple

import yaml
from rapidfuzz import fuzz, process


DEFAULT_THRESHOLD = 80.0


@dataclass
class SchemaMappingResult:
    column_mapping: Dict[str, str]
    unmapped_columns: Dict[str, float]


def _load_keywords(path: str | Path) -> Dict[str, Iterable[str]]:
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    return {k: [str(v).lower() for v in (vals or [])] for k, vals in data.items()}


def build_schema_mapping(
    columns: Iterable[str],
    keywords_path: str | Path,
    threshold: float = DEFAULT_THRESHOLD,
) -> SchemaMappingResult:
    """
    Map vendor column headers to internal schema keys using fuzzy matching.

    Returns both the mapping and a record of unmapped columns with their best
    (but below-threshold) scores for diagnostics.
    """
    keywords = _load_keywords(keywords_path)

    # Build a flat search space of (header_variant, internal_key)
    search_strings = []
    for internal_key, variants in keywords.items():
        for variant in variants:
            search_strings.append((variant, internal_key))

    column_mapping: Dict[str, str] = {}
    unmapped: Dict[str, float] = {}

    for col in columns:
        if col is None:
            continue
        col_str = str(col).strip()
        if not col_str:
            continue

        col_lower = col_str.lower()

        # Use RapidFuzz process to find best match across all variants.
        best_score = -1.0
        best_internal: str | None = None
        for variant, internal_key in search_strings:
            score = fuzz.token_sort_ratio(col_lower, variant)
            if score > best_score:
                best_score = score
                best_internal = internal_key

        if best_internal is not None and best_score >= threshold:
            column_mapping[col_str] = best_internal
        else:
            unmapped[col_str] = best_score

    return SchemaMappingResult(column_mapping=column_mapping, unmapped_columns=unmapped)

