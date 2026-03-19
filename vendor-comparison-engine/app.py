from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Dict, List, Tuple

import streamlit as st
import yaml

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
from rapidfuzz import fuzz


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
    a_names = list(vendor_a_norm.keys())
    b_names = list(vendor_b_norm.keys())

    pairs: List[Tuple[str, str, str]] = []
    used_b: set[str] = set()

    # Exact matches first
    for a in a_names:
        if a in vendor_b_norm:
            pairs.append((a, a, a))
            used_b.add(a)

    # Fuzzy matches
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
            scope = a
            pairs.append((scope, a, best_name))
            used_b.add(best_name)

    return pairs


def run_comparison(
    vendor_a_path: Path,
    vendor_b_path: Path,
    sheet_pairs: List[Tuple[str, str, str]],
) -> Dict[str, object]:
    base_dir = Path(__file__).parent
    keywords_path = base_dir / "config" / "keywords.yaml"
    thresholds_path = base_dir / "config" / "thresholds.yaml"
    overrides_path = base_dir / "config" / "overrides.yaml"

    vendor_a_norm = _normalize_vendor(vendor_a_path, "A", keywords_path)
    vendor_b_norm = _normalize_vendor(vendor_b_path, "B", keywords_path)

    with open(thresholds_path, "r", encoding="utf-8") as f:
        thresholds_cfg = yaml.safe_load(f) or {}
    fuzzy_cfg = thresholds_cfg.get("fuzzy", {})
    auto_threshold = float(fuzzy_cfg.get("auto", 92))
    review_threshold = float(fuzzy_cfg.get("review", 65))
    use_position_fallback = bool(fuzzy_cfg.get("use_position_fallback", False))
    position_min_score = float(fuzzy_cfg.get("position_min_score", 50))

    overrides = load_overrides(overrides_path)
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

    validation_result = validate_totals(
        (n.dataframe for n in vendor_a_norm.values()),
        (n.dataframe for n in vendor_b_norm.values()),
        thresholds_path,
    )

    vendor_a_sheets = {name: n.dataframe for name, n in vendor_a_norm.items()}
    vendor_b_sheets = {name: n.dataframe for name, n in vendor_b_norm.items()}

    deltas = calculate_deltas(all_matches, vendor_a_sheets, vendor_b_sheets)
    audit_df = build_audit_trail(all_matches)

    # Build Excel workbook in a temporary file and return bytes
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir) / "comparison.xlsm"
        write_excel_workbook(
            output_path=tmp_path,
            matches_with_deltas=deltas.matches_with_deltas,
            unmatched=deltas.unmatched,
            audit_df=audit_df,
            vendor_a_sheets=vendor_a_sheets,
            vendor_b_sheets=vendor_b_sheets,
            validation_ok=validation_result.ok,
            sheet_pairs=sheet_pairs,
        )
        data = tmp_path.read_bytes()

    return {
        "excel_bytes": data,
        "validation": validation_result,
        "matches_with_deltas": deltas.matches_with_deltas,
        "unmatched": deltas.unmatched,
        "audit": audit_df,
        "vendor_a_sheets": vendor_a_sheets,
        "vendor_b_sheets": vendor_b_sheets,
    }


def main() -> None:
    st.set_page_config(page_title="Vendor Comparison Engine", layout="wide")

    # Leadec-style header: dark blue bar, white text
    st.markdown(
        """
        <style>
        .leadec-header {
            background: #002855;
            color: white;
            padding: 1.5rem 2rem;
            margin-bottom: 1.5rem;
            border-radius: 0;
        }
        .leadec-header h1 {
            color: white !important;
            margin: 0;
            font-size: 1.75rem;
            font-weight: 700;
        }
        .leadec-header p {
            color: rgba(255,255,255,0.9);
            margin: 0.35rem 0 0 0;
            font-size: 0.95rem;
        }
        /* Primary buttons: solid dark blue */
        .stButton > button {
            background-color: #002855 !important;
            color: white !important;
        }
        .stButton > button:hover {
            background-color: #003366 !important;
            color: white !important;
        }
        /* Section headings in Leadec blue */
        h2, h3 { color: #002855 !important; }
        </style>
        <div class="leadec-header">
            <h1>Vendor Comparison Engine</h1>
            <p>Upload two vendor Excel quote files to generate a comparison workbook.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)
    with col1:
        vendor_a_file = st.file_uploader("Vendor A (.xlsx)", type=["xlsx"], key="vendor_a")
    with col2:
        vendor_b_file = st.file_uploader("Vendor B (.xlsx)", type=["xlsx"], key="vendor_b")

    sheet_pairs: List[Tuple[str, str, str]] = []
    if vendor_a_file and vendor_b_file:
        # Inspect sheet names and propose pairings with confidence scores
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)
            a_tmp = tmpdir_path / "vendor_a.xlsx"
            b_tmp = tmpdir_path / "vendor_b.xlsx"
            a_tmp.write_bytes(vendor_a_file.getvalue())
            b_tmp.write_bytes(vendor_b_file.getvalue())

            wb_a = read_vendor_workbook(a_tmp, vendor_id="A")
            wb_b = read_vendor_workbook(b_tmp, vendor_id="B")
            a_names = list(wb_a.keys())
            b_names = list(wb_b.keys())

            with open(Path(__file__).parent / "config" / "thresholds.yaml", "r", encoding="utf-8") as f:
                thresholds_cfg = yaml.safe_load(f) or {}
            sheet_threshold = float(thresholds_cfg.get("sheet_name_threshold", 80))

            proposals = []
            for a_name in a_names:
                best_b = None
                best_score = -1.0
                for b_name in b_names:
                    score = fuzz.token_sort_ratio(a_name.lower(), b_name.lower())
                    if score > best_score:
                        best_score = score
                        best_b = b_name
                proposals.append((a_name, best_b, best_score))

        st.subheader("Sheet matching")
        st.caption(
            "Review and adjust how sheets from Vendor A are matched to sheets from Vendor B. "
            "Leave as 'None' if there is no corresponding sheet."
        )

        selected_pairs: List[Tuple[str, str, str]] = []
        for a_name, best_b, score in proposals:
            cols = st.columns([2, 2, 1])
            with cols[0]:
                st.markdown(f"**Vendor A sheet:** `{a_name}`")
            with cols[1]:
                options = ["None"] + b_names
                default = best_b if score >= sheet_threshold else "None"
                selection = st.selectbox(
                    f"Match for `{a_name}`",
                    options=options,
                    index=options.index(default) if default in options else 0,
                    key=f"pair_{a_name}",
                )
            with cols[2]:
                st.markdown(f"Confidence: `{int(score)}`")

            if selection != "None":
                # Use A sheet name as canonical scope label
                selected_pairs.append((a_name, a_name, selection))

        sheet_pairs = selected_pairs

    run_button = st.button(
        "Run comparison",
        type="primary",
        disabled=not (vendor_a_file and vendor_b_file and sheet_pairs),
    )

    if run_button and vendor_a_file and vendor_b_file and sheet_pairs:
        with st.spinner("Running comparison..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = Path(tmpdir)
                a_path = tmpdir_path / "vendor_a.xlsx"
                b_path = tmpdir_path / "vendor_b.xlsx"
                a_path.write_bytes(vendor_a_file.getvalue())
                b_path.write_bytes(vendor_b_file.getvalue())

                result = run_comparison(a_path, b_path, sheet_pairs)

        validation = result["validation"]

        status_color = "green" if validation.ok else "orange"
        status_label = "PASS" if validation.ok else "Totals differ (report still generated)"
        st.markdown(
            f"**Validation:** "
            f"<span style='color:{status_color}'>{status_label}</span>",
            unsafe_allow_html=True,
        )
        for msg in validation.messages:
            st.write(f"- {msg}")
        if not validation.ok:
            st.caption("The workbook was still generated so you can compare line-by-line. Adjust totals in your analysis as needed.")

        st.download_button(
            label="Download comparison workbook (.xlsx)",
            data=result["excel_bytes"],
            file_name="vendor_comparison.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
        )


if __name__ == "__main__":
    main()
