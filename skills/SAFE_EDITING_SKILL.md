# Safe Editing Skill — Event-Sourced Session Workflow

## Overview

This skill describes the required workflow for safely editing spreadsheet workbooks
using the `asp` CLI's event-sourced session model. All structural mutations must
go through session staging to ensure undo/redo, impact visibility, and audit trails.

## Required Workflow

### 1. Start a Session

```bash
asp session start --base <workbook.xlsx> --label "Description of edit session"
# Returns: session_id
```

### 2. Explore Before Editing

Before any structural edit, **always** run these discovery commands:

```bash
# Understand named ranges and table structures
asp named-ranges <workbook.xlsx>

# Trace formula dependencies from key cells
asp formula-trace <workbook.xlsx> <Sheet> <Cell> precedents --depth 2

# Review current layout
asp layout-page <workbook.xlsx> <Sheet> --range <area>

# Read values in row-oriented format for layout mapping
asp range-values <workbook.xlsx> <Sheet> <range> --format rows
```

### 3. Preflight Structural Risk

Before insert/delete row/column operations, **always** run impact analysis:

```bash
# Standalone preflight (read-only)
asp check-ref-impact <workbook.xlsx> --ops @structure_ops.json --show-formula-delta

# Or via structure-batch dry-run
asp structure-batch <workbook.xlsx> --ops @structure_ops.json --dry-run \
    --impact-report --show-formula-delta
```

Review the output for:
- `absolute_ref_warnings`: `$`-anchored refs crossing insertion/deletion zones
- `shifted_spans`: which rows/cols shift and by how much
- `formula_delta_preview`: before/after formula comparisons

### 4. Stage Operations

Stage the operation to compute impact without advancing HEAD:

```bash
asp session op --session <session_id> --ops @edits.json
# Returns: staged_id + impact analysis
```

### 5. Apply Staged Operations

Only apply after reviewing the staged impact:

```bash
asp session apply --session <session_id> <staged_id>
```

### 6. Verify After Apply

```bash
# Diff to confirm changes
asp diff <base.xlsx> <materialized.xlsx> --details --limit 50

# Recalculate and check for errors
asp recalculate <output.xlsx> --changed-cells

# Inspect specific cells
asp inspect-cells <output.xlsx> <Sheet> <targets>
```

### 7. Recovery

If an edit causes problems, use session navigation:

```bash
# Undo the last operation
asp session undo --session <session_id>

# Or checkout a specific point
asp session checkout --session <session_id> <op_id>

# Or fork a new branch to try an alternative approach
asp session fork --session <session_id> alt-approach --from <op_id>
```

### 8. Materialize Final Output

```bash
asp session materialize --session <session_id> --output final_result.xlsx
```

## Structural Edit Recipes

### Inserting Rows in Modeled Zones

**Prefer `clone_row` over raw `insert_rows`** in modeled zones (zones with formulas):

```json
{
  "ops": [{
    "kind": "clone_row",
    "sheet_name": "Sheet1",
    "source_row": 85,
    "insert_at": 86,
    "expand_adjacent_sums": true
  }]
}
```

`clone_row` preserves:
- Row formatting and styles
- Formula patterns (relative references adjust)
- Adjacent SUM range expansion (when `expand_adjacent_sums: true`)

### Copying Ranges with Style Inheritance

```json
{
  "ops": [{
    "kind": "copy_range",
    "sheet_name": "Source",
    "dest_sheet_name": "Target",
    "src_range": "A1:F10",
    "dest_anchor": "A1",
    "include_styles": true,
    "include_formulas": true
  }]
}
```

## Anti-Patterns

- **Never** use raw `insert_rows` in formula-heavy zones without `--impact-report`
- **Never** skip `named-ranges` and `formula-trace` before structural edits
- **Never** restart from scratch when an edit fails — use `session undo` or `session checkout`
- **Never** decode dense encoding manually — use `--format rows` or `--format json`
- **Never** apply structural operations without reviewing `absolute_ref_warnings`
