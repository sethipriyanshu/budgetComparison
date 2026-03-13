from __future__ import annotations

from pathlib import Path
from typing import Dict, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formatting.rule import CellIsRule

from src.processing.comparison_builder import build_comparison_rows_for_scope, _get_item_name


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

    # By-section comparison (same sheet)
    scopes_for_summary = sorted(set(list(vendor_a_sheets.keys()) + list(vendor_b_sheets.keys())))
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
            item_name = _get_item_name(r)
            total = r.get("total_price")
            unmatched_ws.append([scope, "A", item_ref, item_name, total])
        elif pd.notna(b_idx) and b_sheet and b_sheet in vendor_b_sheets:
            r = vendor_b_sheets[b_sheet].loc[int(b_idx)]
            item_ref = r.get("item_id") or ""
            item_name = _get_item_name(r)
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
                    _get_item_name(row),
                    row.get("total_price"),
                ])
    _auto_fit_columns(options_ws)

    wb.save(output_path)

