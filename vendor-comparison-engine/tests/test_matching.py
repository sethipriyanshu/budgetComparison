import pandas as pd

from src.processing.matcher import match_scope


def test_match_scope_exact_item_id():
    # Use descriptions that don't fuzzy-match so we get one exact match + two unmatched
    a_df = pd.DataFrame(
        {
            "item_id": ["1", "2"],
            "description": ["Item A", "Item B"],
            "description_norm": ["item a", "item b"],
        }
    )
    b_df = pd.DataFrame(
        {
            "item_id": ["1", "3"],
            "description": ["Item A", "Zebra wiring assembly"],
            "description_norm": ["item a", "zebra wiring assembly"],
        }
    )

    matches = match_scope(
        scope="Test",
        vendor_a_df=a_df,
        vendor_b_df=b_df,
        auto_threshold=92,
        review_threshold=75,
        vendor_a_sheet_name="A",
        vendor_b_sheet_name="B",
        manual_matches=None,
    )

    # One AUTO match on item_id, plus two unmatched (A idx=1, B idx=1)
    assert any(m.method == "ITEM_ID_EXACT" and m.vendor_a_idx == 0 and m.vendor_b_idx == 0 for m in matches)
    assert any(m.method == "UNMATCHED" and m.vendor_a_idx == 1 for m in matches)
    assert any(m.method == "UNMATCHED" and m.vendor_b_idx == 1 for m in matches)

