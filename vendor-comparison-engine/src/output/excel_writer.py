from __future__ import annotations

from pathlib import Path
from typing import Dict, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formatting.rule import CellIsRule

from src.processing.comparison_builder import build_comparison_rows_for_scope


def _auto_fit_columns(ws) -> None:
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # type: ignore[attr-defined]
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                continue
        ws.column_dimensions[column].width = min(max_length + 2, 60)


def write_excel_workbook(
    output_path: Path,
    matches_with_deltas: pd.DataFrame,
    unmatched: pd.DataFrame,
    audit_df: pd.DataFrame,
    vendor_a_sheets: Dict[str, pd.DataFrame],
    vendor_b_sheets: Dict[str, pd.DataFrame],
    validation_ok: bool,
    unmatched_diagnostics: Optional[pd.DataFrame] = None,
) -> None:
    wb = Workbook()

    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    # Summary sheet (simple, for decision-making)
    summary = wb.create_sheet("Summary")
    summary["A1"] = "Vendor comparison"
    summary["A1"].font = Font(bold=True, size=14)
    summary["A3"] = "Lines compared (matched)"
    summary["B3"] = int(len(matches_with_deltas))
    summary["A4"] = "Lines not matched (only one vendor quoted)"
    summary["B4"] = int(len(unmatched))
    summary["A6"] = "Note: Totals can differ — not all vendors quote every item. Use this file for line-by-line comparison."
    summary["A6"].font = Font(italic=True)

    # Per-scope comparison sheets
    scopes = sorted(set(matches_with_deltas["scope_category"].dropna().tolist()))

    for scope in scopes:
        sheet_name = f"{scope} Comparison"
        ws = wb.create_sheet(sheet_name[:31])  # Excel sheet name limit

        headers = [
            "Item ref",
            "Item name",
            "Qty (A)",
            "Unit price (A)",
            "Total (A)",
            "Qty (B)",
            "Unit price (B)",
            "Total (B)",
            "Price difference ($)",
            "Price difference (%)",
        ]
        ws.append(headers)

        rows = build_comparison_rows_for_scope(
            scope,
            matches_with_deltas,
            vendor_a_sheets,
            vendor_b_sheets,
        )
        for row in rows:
            ws.append(row)

        # Simple conditional formatting on Price Delta (%) column (J)
        max_row = ws.max_row
        pct_col = "J"
        if max_row >= 2:
            cell_range = f"{pct_col}2:{pct_col}{max_row}"
            ws.conditional_formatting.add(
                cell_range,
                CellIsRule(
                    operator="greaterThan",
                    formula=["0.05"],
                    fill=PatternFill(
                        start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
                    ),
                ),
            )
            ws.conditional_formatting.add(
                cell_range,
                CellIsRule(
                    operator="between",
                    formula=["0.01", "0.05"],
                    fill=PatternFill(
                        start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"
                    ),
                ),
            )
            ws.conditional_formatting.add(
                cell_range,
                CellIsRule(
                    operator="lessThanOrEqual",
                    formula=["0.01"],
                    fill=PatternFill(
                        start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
                    ),
                ),
            )

        _auto_fit_columns(ws)

    # Items not matched (only one vendor quoted — useful for decision-making)
    unmatched_ws = wb.create_sheet("Items not matched")
    unmatched_ws.append(["Section", "Vendor", "Item ref", "Item name", "Quoted total"])
    for _, row in unmatched.iterrows():
        scope = row.get("scope_category")
        a_sheet = row.get("vendor_a_sheet")
        b_sheet = row.get("vendor_b_sheet")
        a_idx = row.get("vendor_a_idx")
        b_idx = row.get("vendor_b_idx")
        if pd.notna(a_idx) and a_sheet and a_sheet in vendor_a_sheets:
            r = vendor_a_sheets[a_sheet].loc[int(a_idx)]
            item_ref = r.get("item_id") or ""
            item_name = r.get("description") or ""
            total = r.get("total_price")
            unmatched_ws.append([scope, "A", item_ref, item_name, total])
        elif pd.notna(b_idx) and b_sheet and b_sheet in vendor_b_sheets:
            r = vendor_b_sheets[b_sheet].loc[int(b_idx)]
            item_ref = r.get("item_id") or ""
            item_name = r.get("description") or ""
            total = r.get("total_price")
            unmatched_ws.append([scope, "B", item_ref, item_name, total])
    _auto_fit_columns(unmatched_ws)

    # Optional: why unmatched (closest match from other vendor — for tuning)
    if unmatched_diagnostics is not None and not unmatched_diagnostics.empty:
        diag_ws = wb.create_sheet("Why not matched")
        diag_ws.append(["Vendor", "Section", "Item name", "Quoted total", "Closest match (other vendor)", "Similarity %"])
        for r in unmatched_diagnostics.itertuples(index=False):
            diag_ws.append([
                getattr(r, "Vendor", ""),
                getattr(r, "Scope", ""),
                getattr(r, "Description", ""),
                "",  # total not in diagnostics; could add
                getattr(r, "Best_Match_Description", ""),
                getattr(r, "Best_Match_Score", ""),
            ])
        _auto_fit_columns(diag_ws)

    # Optional items (options / alternatives — excluded from main totals)
    options_ws = wb.create_sheet("Options")
    options_ws.append(["Section", "Vendor", "Item name", "Total"])
    for vendor_label, sheets in (("A", vendor_a_sheets), ("B", vendor_b_sheets)):
        for scope, df in sheets.items():
            if "row_type" not in df.columns:
                continue
            is_option = df["row_type"] == "OPTION"
            for idx in df[is_option].index:
                row = df.loc[idx]
                options_ws.append([
                    scope,
                    vendor_label,
                    row.get("description") or "",
                    row.get("total_price"),
                ])
    _auto_fit_columns(options_ws)

    wb.save(output_path)

