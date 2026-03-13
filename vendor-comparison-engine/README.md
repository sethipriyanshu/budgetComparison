# Vendor Comparison Engine

Compare two vendor Excel quote files with different layouts. The engine normalizes and matches line items across sheets, then produces a single comparison workbook with side-by-side prices, deltas, and comments.

## Features

- **Multi-sheet support**: Pairs sheets between vendors (e.g. "Mechanics" ↔ "ATS Mechanics v2") via exact or fuzzy sheet-name matching.
- **Flexible schema**: Auto-maps vendor columns to a common schema using configurable keywords (item id, description, qty, unit price, total, comments).
- **Line-item matching**: Matches rows by item ID, exact description, then fuzzy description; optional position fallback for remaining rows.
- **Clean output**: Excludes total/subtotal rows and phantom lines; includes a Comments column and color-coded Vendor A / Vendor B columns.
- **Excel output**: Summary sheet with totals and per-section comparison sheets with conditional formatting on price difference %.

## Requirements

- Python 3.11+
- Dependencies in `requirements.txt`: pandas, openpyxl, rapidfuzz, PyYAML, streamlit, pytest

## Installation

```bash
cd vendor-comparison-engine
pip install -r requirements.txt
```

## Usage

### Streamlit UI (recommended)

1. Run the app:
   ```bash
   streamlit run app.py
   ```
2. Upload **Vendor A** and **Vendor B** Excel files (.xlsx).
3. Review **Sheet matching**: confirm or change how each Vendor A sheet is paired with a Vendor B sheet (or set to "None").
4. Click **Run comparison**.
5. Download the generated **comparison workbook** (.xlsx).

### CLI

```bash
python compare.py --vendor-a path/to/vendor_a.xlsx --vendor-b path/to/vendor_b.xlsx --output-dir ./output
```

- Outputs the comparison workbook under `--output-dir` and optionally CSV exports of matches/unmatched (if `--debug` is used).

## Configuration

| File | Purpose |
|------|---------|
| `config/keywords.yaml` | Column name synonyms for schema mapping (item_id, description, qty, total_price, comments, etc.). |
| `config/thresholds.yaml` | Fuzzy match thresholds (`auto` / `review`), `use_position_fallback`, `position_min_score`, `sheet_name_threshold`. |
| `config/overrides.yaml` | Manual match overrides (force specific A row ↔ B row) when needed. |

## Output workbook

- **Summary**: Totals for Vendor A and B, difference, line counts; “By section” table using your sheet pairs.
- **Per-section sheets** (e.g. "Mechanics Comparison", "Electrics Comparison"): Item ref, Item name, Qty/Unit/Total for A and B, price difference ($ and %), Comments. Vendor A columns are light blue, Vendor B light orange; difference % has conditional formatting (green / yellow / red).

## Project structure

```
vendor-comparison-engine/
├── app.py              # Streamlit UI
├── compare.py          # CLI entry point
├── requirements.txt
├── config/
│   ├── keywords.yaml
│   ├── thresholds.yaml
│   └── overrides.yaml
├── src/
│   ├── ingestion/      # Reader, header detection, schema mapping
│   ├── processing/     # Normalizer, matcher, deltas, comparison builder, validator
│   └── output/         # Excel writer, audit logger
└── tests/
```

## Tests

```bash
pytest tests/ -v
```

## License

Internal use. See your organization’s policy for distribution.
