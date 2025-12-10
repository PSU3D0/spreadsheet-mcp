# Region Detection Heuristics

This document describes the heuristics used for automatic region detection and classification in spreadsheets. These are best-effort guesses based on structural patterns - always verify with actual data.

## Current Limitations

**Important**: These heuristics were developed and tested primarily against a single complex financial spreadsheet. They may not generalize well to:
- Simple flat data tables
- Pivot tables or crosstab layouts
- Heavily styled/formatted sheets where structure comes from formatting not data
- Non-English spreadsheets with different labeling conventions
- Spreadsheets with merged cells (not yet handled)

## Classification Labels

All classifications are prefixed with `likely_` to indicate uncertainty:
- `likely_data` - Tabular data with headers
- `likely_parameters` - Key-value configuration/input cells
- `likely_calculator` - Formula-heavy computation regions (>55% formulas)
- `likely_outputs` - Mixed formula regions (25-55% formulas)
- `likely_metadata` - Labels, titles, or sparse informational content
- `unknown` - Could not classify with confidence

## Key-Value Layout Detection

**Purpose**: Identify vertical parameter/config regions (label in col A, value in col B pattern)

**Assumptions**:
1. Key-value layouts have exactly 2 "dense" columns (≥40% fill rate in sampled rows)
2. Labels are short text (≤25 chars), contain letters, no digits
3. Values are numbers, dates, or text longer than 2 chars
4. At least 3 valid label-value pairs in first 15 rows
5. At least 30% of sampled rows are valid pairs

**Known Issues**:
- Sparse columns (like occasional notes in col D) can trip detection if they happen to have values in the sampled rows
- Doesn't handle horizontal key-value layouts (row 1 = labels, row 2 = values)
- Labels with numbers (e.g., "Rate 1", "Tier 2") are rejected as keys

## Header Detection

**Purpose**: Find the row containing column headers for tabular data

**Assumptions**:
1. Headers are in one of the first 3 rows of a region
2. Header rows have more text cells than numeric cells
3. Header values are relatively unique (not repeated)
4. Data-like values (proper nouns >5 chars, strings with digits >3 chars, very long strings >40 chars) are penalized
5. Date values in a potential header row reduce its score

**Known Issues**:
- Proper noun detection is naive (just checks capitalization + length)
- Doesn't handle multi-row headers well
- English-centric assumptions about what "looks like" a header

## Region Classification

**Formula Ratio Thresholds**:
- `likely_calculator`: >55% formula cells
- `likely_outputs`: 25-55% formula cells
- `likely_parameters`: <25% formulas AND (key-value layout OR narrow ≤3 cols)
- `likely_metadata`: Few non-empty cells, mostly text
- `likely_data`: Default fallback

**Confidence Scoring**:
- Based on formula consistency, cell density, header quality
- Confidence <0.5 indicates uncertain classification
- Always check `confidence` field before trusting classification

## Column Density Calculation

When detecting key-value layouts in regions wider than 2 columns:
1. Sample first 20 rows (or all rows if fewer)
2. Count non-null cells per column
3. Columns with ≥40% fill rate are "dense"
4. If exactly 2 dense columns exist, treat as potential key-value layout

## Future Improvements Needed

1. **Diverse test corpus**: Current tests use synthetic micro-sheets; need real-world variety
2. **Merged cell handling**: Currently ignored, can break region detection
3. **Style-based hints**: Bold headers, borders, background colors carry semantic meaning
4. **Horizontal key-value**: Support label-row / value-row patterns
5. **Multi-region awareness**: Adjacent regions with different structures (e.g., config block next to data table)
6. **Confidence calibration**: Current confidence scores are not well-calibrated to actual accuracy
