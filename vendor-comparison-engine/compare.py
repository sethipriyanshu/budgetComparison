from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
from rapidfuzz import fuzz

from src.ingestion.reader import read_vendor_workbook
from src.ingestion.header_detector import apply_header_detection
from src.ingestion.schema_mapper import build_schema_mapping
from src.processing.normalizer import NormalizedSheet, normalize_sheet
from src.config.override_manager import apply_overrides, load_overrides
from src.processing.matcher import match_scope, MatchRecord
from src.processing.delta_calculator import calculate_deltas
from src.processing.validator import validate_totals
from src.output.audit_logger import build_audit_trail
from src.output.excel_writer import write_excel_workbook


def _normalize_vendor(
    path: Path, vendor_id: str, keywords_path: Path
) -> Dict[str, NormalizedSheet]:
    sheets = read_vendor_workbook(path, vendor_id=vendor_id)
    normalized: Dict[str, NormalizedSheet] = {}

    for sheet_name, sheet_data in sheets.items():
        header_result = apply_header_detection(sheet_data.dataframe)
        mapping_result = build_schema_mapping(
            columns=header_result.data_with_headers.columns,
            keywords_path=keywords_path,
        )

        norm = normalize_sheet(
            df=header_result.data_with_headers,
            column_mapping=mapping_result.column_mapping,
            scope_category=sheet_name,
            vendor_id=vendor_id,
        )
        normalized[sheet_name] = norm

    return normalized


def _pair_sheets(
    vendor_a_norm: Dict[str, NormalizedSheet],
    vendor_b_norm: Dict[str, NormalizedSheet],
    sheet_threshold: float,
) -> List[Tuple[str, str, str]]:
    """
    Return list of (scope_name, vendor_a_sheet_name, vendor_b_sheet_name) pairs.

    Uses exact name match first, then fuzzy matching on remaining sheet names so
    that e.g. 'Mechanics' pairs with 'ATS Mechanics v2'.
    """
    a_names = list(vendor_a_norm.keys())
    b_names = list(vendor_b_norm.keys())

    pairs: List[Tuple[str, str, str]] = []
    used_b: set[str] = set()

    # 1) Exact matches
    for a in a_names:
        if a in vendor_b_norm:
            pairs.append((a, a, a))
            used_b.add(a)

    # 2) Fuzzy matches for remaining A sheets
    for a in a_names:
        if any(p[1] == a for p in pairs):
            continue

        best_name = None
        best_score = -1.0
        for b in b_names:
            if b in used_b:
                continue
            score = fuzz.token_sort_ratio(a.lower(), b.lower())
            if score > best_score:
                best_score = score
                best_name = b

        if best_name is not None and best_score >= sheet_threshold:
            scope = a  # use Vendor A's name as canonical scope label
            pairs.append((scope, a, best_name))
            used_b.add(best_name)

    return pairs


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Phase 2: Normalize two vendor Excel quotes, match items, "
            "compute deltas, and run basic validation."
        )
    )
    parser.add_argument("--vendor-a", required=True, help="Path to Vendor A .xlsx file")
    parser.add_argument("--vendor-b", required=True, help="Path to Vendor B .xlsx file")
    parser.add_argument(
        "--debug-output",
        help="Optional folder to write intermediate CSV files for inspection",
    )
    parser.add_argument(
        "--output",
        help="Optional path to write the final comparison Excel workbook (.xlsm)",
    )

    args = parser.parse_args()

    base_dir = Path(__file__).parent
    keywords_path = base_dir / "config" / "keywords.yaml"
    thresholds_path = base_dir / "config" / "thresholds.yaml"
    overrides_path = base_dir / "config" / "overrides.yaml"

    vendor_a_path = Path(args.vendor_a)
    vendor_b_path = Path(args.vendor_b)

    vendor_a_norm = _normalize_vendor(vendor_a_path, "A", keywords_path)
    vendor_b_norm = _normalize_vendor(vendor_b_path, "B", keywords_path)

    print("=== Phase 2 Matching & Validation ===")
    print(f"Vendor A sheets: {list(vendor_a_norm.keys())}")
    print(f"Vendor B sheets: {list(vendor_b_norm.keys())}")

    # Load config for thresholds and overrides
    import yaml

    with open(thresholds_path, "r", encoding="utf-8") as f:
        thresholds_cfg = yaml.safe_load(f) or {}
    fuzzy_cfg = thresholds_cfg.get("fuzzy", {})
    auto_threshold = float(fuzzy_cfg.get("auto", 92))
    review_threshold = float(fuzzy_cfg.get("review", 65))
    use_position_fallback = bool(fuzzy_cfg.get("use_position_fallback", False))
    position_min_score = float(fuzzy_cfg.get("position_min_score", 50))

    overrides = load_overrides(overrides_path)

    # Sheet pairing (exact + fuzzy) so 'Mechanics' matches 'ATS Mechanics v2'
    sheet_threshold = float(thresholds_cfg.get("sheet_name_threshold", 80))
    sheet_pairs = _pair_sheets(vendor_a_norm, vendor_b_norm, sheet_threshold)

    print(f"Sheet pairs (scope, A, B): {sheet_pairs}")

    all_matches: list[MatchRecord] = []

    for scope, a_sheet_name, b_sheet_name in sheet_pairs:
        norm_a = vendor_a_norm[a_sheet_name]
        norm_b = vendor_b_norm[b_sheet_name]

        a_df = norm_a.dataframe
        b_df = norm_b.dataframe

        manual_matches = apply_overrides(a_df, b_df, overrides)

        scope_matches = match_scope(
            scope=scope,
            vendor_a_df=a_df,
            vendor_b_df=b_df,
            auto_threshold=auto_threshold,
            review_threshold=review_threshold,
            vendor_a_sheet_name=a_sheet_name,
            vendor_b_sheet_name=b_sheet_name,
            manual_matches=manual_matches,
            use_position_fallback=use_position_fallback,
            position_min_score=position_min_score,
        )
        all_matches.extend(scope_matches)

    # Basic validation across all sheets
    validation_result = validate_totals(
        (n.dataframe for n in vendor_a_norm.values()),
        (n.dataframe for n in vendor_b_norm.values()),
        thresholds_path,
    )
    for msg in validation_result.messages:
        print(f"[VALIDATION] {msg}")

    # Compute deltas and audit trail
    vendor_a_sheets = {name: n.dataframe for name, n in vendor_a_norm.items()}
    vendor_b_sheets = {name: n.dataframe for name, n in vendor_b_norm.items()}

    deltas = calculate_deltas(all_matches, vendor_a_sheets, vendor_b_sheets)
    audit_df = build_audit_trail(all_matches)

    print(f"\nMatched rows with deltas: {len(deltas.matches_with_deltas)}")
    print(f"Unmatched rows: {len(deltas.unmatched)}")

    if args.debug_output:
        out_dir = Path(args.debug_output)
        out_dir.mkdir(parents=True, exist_ok=True)

        deltas.matches_with_deltas.to_csv(
            out_dir / "matches_with_deltas.csv", index=False
        )
        deltas.unmatched.to_csv(out_dir / "unmatched.csv", index=False)
        audit_df.to_csv(out_dir / "audit_trail.csv", index=False)

        print(f"\nDebug CSVs written to: {out_dir}")

    if args.output:
        output_path = Path(args.output)
        if output_path.suffix.lower() != ".xlsm":
            output_path = output_path.with_suffix(".xlsm")
            print(f"[OUTPUT] Switching to template-compatible Excel output: {output_path}")
        write_excel_workbook(
            output_path=output_path,
            matches_with_deltas=deltas.matches_with_deltas,
            unmatched=deltas.unmatched,
            audit_df=audit_df,
            vendor_a_sheets=vendor_a_sheets,
            vendor_b_sheets=vendor_b_sheets,
            validation_ok=validation_result.ok,
            sheet_pairs=sheet_pairs,
        )
        print(f"\nExcel comparison workbook written to: {output_path}")


if __name__ == "__main__":
    main()
