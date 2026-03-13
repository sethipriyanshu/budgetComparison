from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import CellIsRule

# Column color coding: A=shared, B=shared, C-E=Vendor A, F-H=Vendor B, I-K=comparison/comments
FILL_VENDOR_A = PatternFill(start_color="DAE3F3", end_color="DAE3F3", fill_type="solid")  # light blue
FILL_VENDOR_B = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")  # light orange
FILL_NEUTRAL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")   # light gray

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


def _apply_comparison_column_colors(ws, max_row: int) -> None:
    """Apply Vendor A / Vendor B / neutral background to columns for easy differentiation."""
    # A-B: shared (neutral), C-E: Vendor A, F-H: Vendor B, I-K: comparison/comments (neutral)
    for row in range(1, max_row + 1):
        for col in range(1, 12):  # columns 1-11 (A-K)
            cell = ws.cell(row=row, column=col)
            if col <= 2:
                cell.fill = FILL_NEUTRAL
            elif col <= 5:
                cell.fill = FILL_VENDOR_A
            elif col <= 8:
                cell.fill = FILL_VENDOR_B
            else:
                cell.fill = FILL_NEUTRAL


def write_excel_workbook(
    output_path: Path,
    matches_with_deltas: pd.DataFrame,
    unmatched: pd.DataFrame,
    audit_df: pd.DataFrame,
    vendor_a_sheets: Dict[str, pd.DataFrame],
    vendor_b_sheets: Dict[str, pd.DataFrame],
    validation_ok: bool,
    unmatched_diagnostics: Optional[pd.DataFrame] = None,
    sheet_pairs: Optional[List[Tuple[str, str, str]]] = None,
) -> None:
    wb = Workbook()

    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    # Summary sheet: comparison totals + counts
    summary = wb.create_sheet("Summary")
    summary["A1"] = "Vendor comparison"
    summary["A1"].font = Font(bold=True, size=14)

    def _sum_total(sheets: Dict[str, pd.DataFrame]) -> float:
        total = 0.0
        for df in sheets.values():
            if "total_price" in df.columns:
                total += df["total_price"].fillna(0.0).sum()
        return round(total, 2)

    total_a = _sum_total(vendor_a_sheets)
    total_b = _sum_total(vendor_b_sheets)
    diff_abs = round(total_b - total_a, 2) if total_a is not None and total_b is not None else None
    diff_pct = round((diff_abs / total_a * 100), 1) if total_a and total_a != 0 and diff_abs is not None else None

    summary["A3"] = "Total (Vendor A)"
    summary["B3"] = total_a
    summary["A4"] = "Total (Vendor B)"
    summary["B4"] = total_b
    summary["A5"] = "Difference (B − A)"
    summary["B5"] = diff_abs if diff_abs is not None else ""
    summary["A6"] = "Difference (%)"
    summary["B6"] = f"{diff_pct}%" if diff_pct is not None else ""

    summary["A8"] = "Lines compared (matched)"
    summary["B8"] = int(len(matches_with_deltas))
    summary["A9"] = "Lines not matched (only one vendor quoted)"
    summary["B9"] = int(len(unmatched))

    summary["A11"] = "Note: Totals can differ — not all vendors quote every item. Use this file for line-by-line comparison."
    summary["A11"].font = Font(italic=True)

    # By-section comparison: use sheet_pairs so "Mechanics" and "ATS Mechanics v2" show as one section
    summary["A13"] = "By section"
    summary["A13"].font = Font(bold=True)
    summary["A14"] = "Section"
    summary["B14"] = "Total (A)"
    summary["C14"] = "Total (B)"
    summary["D14"] = "Difference"
    summary["A14"].font = Font(bold=True)
    summary["B14"].font = Font(bold=True)
    summary["C14"].font = Font(bold=True)
    summary["D14"].font = Font(bold=True)
    row = 15
    if sheet_pairs:
        for scope, a_sheet, b_sheet in sheet_pairs:
            a_df = vendor_a_sheets.get(a_sheet)
            b_df = vendor_b_sheets.get(b_sheet)
            sa = round(a_df["total_price"].fillna(0).sum(), 2) if a_df is not None and "total_price" in a_df.columns else 0.0
            sb = round(b_df["total_price"].fillna(0).sum(), 2) if b_df is not None and "total_price" in b_df.columns else 0.0
            summary.cell(row=row, column=1, value=scope)
            summary.cell(row=row, column=2, value=sa)
            summary.cell(row=row, column=3, value=sb)
            summary.cell(row=row, column=4, value=round(sb - sa, 2))
            row += 1
    else:
        scopes_for_summary = sorted(set(list(vendor_a_sheets.keys()) + list(vendor_b_sheets.keys())))
        for scope in scopes_for_summary:
            a_df = vendor_a_sheets.get(scope)
            b_df = vendor_b_sheets.get(scope)
            sa = round(a_df["total_price"].fillna(0).sum(), 2) if a_df is not None and "total_price" in a_df.columns else 0.0
            sb = round(b_df["total_price"].fillna(0).sum(), 2) if b_df is not None and "total_price" in b_df.columns else 0.0
            summary.cell(row=row, column=1, value=scope)
            summary.cell(row=row, column=2, value=sa)
            summary.cell(row=row, column=3, value=sb)
            summary.cell(row=row, column=4, value=round(sb - sa, 2))
            row += 1

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
            "Comments",
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

        max_row = ws.max_row
        _apply_comparison_column_colors(ws, max_row)

        # Simple conditional formatting on Price Delta (%) column (J)
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

    wb.save(output_path)

