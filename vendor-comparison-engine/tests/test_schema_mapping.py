from pathlib import Path

from src.ingestion.schema_mapper import build_schema_mapping


def test_schema_mapping_basic(tmp_path):
    keywords_content = """
qty:
  - "quantity"
  - "qty"
description:
  - "description"
  - "item"
"""
    keywords_file = tmp_path / "keywords.yaml"
    keywords_file.write_text(keywords_content, encoding="utf-8")

    columns = ["Item", "Quantity", "Random"]

    result = build_schema_mapping(columns, keywords_file, threshold=80.0)

    assert result.column_mapping["Item"] == "description"
    assert result.column_mapping["Quantity"] == "qty"
    assert "Random" in result.unmapped_columns

