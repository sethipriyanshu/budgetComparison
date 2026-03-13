from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List

import pandas as pd
import yaml


@dataclass
class ValidationResult:
    ok: bool
    messages: List[str]


def _load_validation_config(config_path: Path) -> Dict:
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    return cfg.get("validation", {})


def _sum_total_price_from_frames(frames: Iterable[pd.DataFrame]) -> float:
    total = 0.0
    for df in frames:
        if "total_price" not in df.columns:
            continue
        total += df["total_price"].fillna(0.0).sum()
    return float(total)


def validate_totals(
    vendor_a_frames: Iterable[pd.DataFrame],
    vendor_b_frames: Iterable[pd.DataFrame],
    config_path: Path,
) -> ValidationResult:
    """
    Simple Phase 2 validation: compare sum of total_price columns between all sheets
    for each vendor. Later phases can extend this to compare against declared grand totals.
    """
    cfg = _load_validation_config(config_path)
    tolerance = float(cfg.get("total_tolerance", 0.01))

    messages: List[str] = []

    a_sum = _sum_total_price_from_frames(vendor_a_frames)
    b_sum = _sum_total_price_from_frames(vendor_b_frames)

    if a_sum == 0.0 and b_sum == 0.0:
        messages.append("No total_price values found for either vendor.")
        return ValidationResult(ok=False, messages=messages)

    if abs(a_sum - b_sum) > tolerance:
        messages.append(
            f"Total mismatch: Vendor A={a_sum:.2f}, Vendor B={b_sum:.2f}, "
            f"tolerance={tolerance:.2f}"
        )
        return ValidationResult(ok=False, messages=messages)

    messages.append(
        f"Totals match within tolerance: Vendor A={a_sum:.2f}, Vendor B={b_sum:.2f}"
    )
    return ValidationResult(ok=True, messages=messages)

