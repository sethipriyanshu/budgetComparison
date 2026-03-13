import pandas as pd

from src.ingestion.header_detector import detect_header_row


def test_detect_header_row_with_standard_headers():
    df = pd.DataFrame(
        [
            ["Some project", None, None],
            ["Vendor ABC", None, None],
            ["Item", "Description", "Qty"],
            ["1", "Conveyor", 10],
        ]
    )

    header_idx = detect_header_row(df)

    assert header_idx == 2


def test_detect_header_row_when_no_anchors_found_uses_first_row():
    df = pd.DataFrame(
        [
            ["Col1", "Col2"],
            ["A", "B"],
        ]
    )

    header_idx = detect_header_row(df)

    # Fallback behavior should treat row 0 as header
    assert header_idx == 0

