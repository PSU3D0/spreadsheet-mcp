# Explore Skill — Workbook Discovery & Layout Mapping

## Overview

This skill describes the required discovery workflow for understanding workbook
structure before making any edits. Thorough exploration prevents structural
blowups and ensures agents understand the data model.

## Discovery Workflow

### Step 1: Workbook Overview

```bash
# List all sheets with summary metadata
asp list-sheets <workbook.xlsx>

# Get workbook-level metadata
asp describe <workbook.xlsx>
```

### Step 2: Sheet Structure

```bash
# Detect regions, headers, and table structures
asp sheet-overview <workbook.xlsx> "<Sheet Name>"

# Profile table columns and data types
asp table-profile <workbook.xlsx> --sheet "<Sheet Name>"
```

### Step 3: Named Ranges & Definitions

```bash
# List all named ranges and table definitions
asp named-ranges <workbook.xlsx>

# Filter by sheet or prefix
asp named-ranges <workbook.xlsx> --sheet "<Sheet Name>" --name-prefix "Sales"
```

### Step 4: Layout & Value Reading

For layout mapping (understanding where data lives):

```bash
# Row-oriented format for direct mapping (preferred for agents)
asp range-values <workbook.xlsx> "<Sheet>" A1:Z50 --format rows

# Paginated sheet reading
asp sheet-page <workbook.xlsx> "<Sheet>" --format compact --page-size 100

# Visual layout with column widths, borders, formatting
asp layout-page <workbook.xlsx> "<Sheet>" --range A1:T50 --render json
```

### Step 5: Formula Analysis

```bash
# Map formula patterns by complexity
asp formula-map <workbook.xlsx> "<Sheet>" --sort-by complexity

# Trace dependencies from a key cell
asp formula-trace <workbook.xlsx> "<Sheet>" C2 precedents --depth 2
asp formula-trace <workbook.xlsx> "<Sheet>" C2 dependents --depth 2

# Find specific formula patterns
asp find-formula <workbook.xlsx> "SUM(" --sheet "<Sheet>"

# Check for volatile functions
asp scan-volatiles <workbook.xlsx> --sheet "<Sheet>"
```

### Step 6: Detailed Cell Inspection

```bash
# Inspect specific cells with full metadata (formula + value + style)
asp inspect-cells <workbook.xlsx> "<Sheet>" A1:C3,D4,F7:F10

# Use --budget for larger inspections (up to 200 cells)
asp inspect-cells <workbook.xlsx> "<Sheet>" A1:Z10 --budget 200
```

## Output Format Preferences

| Scenario | Command | Format |
|----------|---------|--------|
| Layout mapping | `range-values` | `--format rows` |
| Data extraction | `read-table` | `--table-format values` |
| Pagination | `sheet-page` | `--format compact` |
| Visual layout | `layout-page` | `--render json` |
| Table export | `range-export` | `--format csv` |

**Avoid** relying on dense encoding (`--format dense`) for layout discovery.
Use `--format rows` or `--format json` for direct row/cell mapping.

## Pre-Edit Checklist

Before any edit operation, confirm:

1. Named ranges identified (`named-ranges`)
2. Formula dependencies traced for affected area (`formula-trace`)
3. Table boundaries understood (`sheet-overview` + `table-profile`)
4. Impact of structural changes assessed (`check-ref-impact`)
5. Non-dense read format used for layout mapping (`--format rows`)
