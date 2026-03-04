# Safe Editing Skill — Event-Sourced Session Workflow

## Overview

This skill describes the required workflow for safely editing spreadsheet workbooks
using the `asp` CLI's event-sourced session model. All mutations must go through
session staging to ensure undo/redo, impact visibility, and audit trails.

**All operation families are supported through sessions:** cell writes (`write_matrix`),
structural edits, style changes, formula patterns, replace-in-formulas, column sizing,
sheet layout, rules (data validation / conditional formatting), and named range operations.

## Required Workflow

### 1. Start a Session

```bash
asp session start --base <workbook.xlsx> --label "Description of edit session" \
    --workspace <project_dir>
# Returns: session_id
```

### 2. Explore Before Editing

Before any structural edit, **always** run these discovery commands.
Once a session exists, use `--session` to read from the session's current state
without materializing to disk:

```bash
# Understand named ranges and table structures
asp named-ranges <workbook.xlsx> --session <session_id> --session-workspace <dir>

# Trace formula dependencies from key cells
asp formula-trace <workbook.xlsx> <Sheet> <Cell> precedents --depth 2 \
    --session <session_id> --session-workspace <dir>

# Review current layout
asp layout-page <workbook.xlsx> <Sheet> --range <area> \
    --session <session_id> --session-workspace <dir>

# Read values in row-oriented format for layout mapping
asp range-values <workbook.xlsx> <Sheet> <range> --format rows \
    --session <session_id> --session-workspace <dir>
```

> **Session-aware reads** resolve against the session's current HEAD state.
> This lets you inspect the workbook after prior operations without materializing
> to a file. When `--session` is omitted, reads target the file path directly.

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

Stage the operation to compute dry-run impact without advancing HEAD:

```bash
asp session op --session <session_id> --ops @edits.json --workspace <dir>
# Returns: staged_id + dry_run_impact
```

**Inspect the `dry_run_impact` in the staging response before applying:**

```json
{
  "staged_id": "stg_...",
  "dry_run_impact": {
    "cells_changed": 9,
    "formulas_rewritten": 45,
    "shifted_spans": [
      { "op_index": 0, "sheet_name": "Sheet1", "axis": "row", "at": 86, "count": 1, "direction": "insert" }
    ],
    "warnings": [],
    "boundary_warnings": ["$A$85 in formula crosses insertion zone"]
  }
}
```

Key fields to check:
- `cells_changed` — number of cells affected
- `formulas_rewritten` — number of formula tokens that will be rewritten
- `shifted_spans` — which rows/cols shift and by how much
- `boundary_warnings` — absolute references at risk

### 5. Apply Staged Operations

Only apply after reviewing the staged impact:

```bash
asp session apply --session <session_id> <staged_id> --workspace <dir>
```

### 6. Verify After Apply

Use session-aware reads to verify without materializing:

```bash
# Read specific cells from session state
asp range-values <workbook.xlsx> <Sheet> <range> --format rows \
    --session <session_id> --session-workspace <dir>

# Inspect critical cells
asp inspect-cells <workbook.xlsx> <Sheet> <targets> \
    --session <session_id> --session-workspace <dir>

# For diff verification, materialize then compare
asp session materialize --session <session_id> --output <temp.xlsx> --workspace <dir>
asp diff <base.xlsx> <temp.xlsx> --details --limit 50

# Recalculate and check for errors
asp recalculate <temp.xlsx> --changed-cells
```

### 7. Recovery

If an edit causes problems, use session navigation — **never restart from scratch**:

```bash
# Undo the last operation
asp session undo --session <session_id> --workspace <dir>

# Or checkout a specific point
asp session checkout --session <session_id> <op_id> --workspace <dir>

# Or fork a new branch to try an alternative approach
asp session fork --session <session_id> alt-approach --from <op_id> --workspace <dir>
```

### 8. Materialize Final Output

```bash
asp session materialize --session <session_id> --output final_result.xlsx --workspace <dir>
```

## Supported Operation Families

All of these work through `session op --ops @file.json`. The payload must include
a `"kind"` field so the session routes to the correct handler.

| Kind prefix | Payload `kind` value | Description |
|---|---|---|
| `structure.*` | `structure.insert_rows`, `structure.clone_row`, etc. | Row/column/sheet mutations |
| `transform.*` | `transform.write_matrix`, `transform.clear_range`, etc. | Cell value operations |
| `style.apply` | `style.apply` | Font, fill, border, alignment |
| `formula.apply_pattern` | `formula.apply_pattern` | Relative formula filling |
| `formula.replace_in_formulas` | `formula.replace_in_formulas` | Find/replace in formula text |
| `column.size` | `column.size` | Column width adjustment |
| `layout.apply` | `layout.apply` | Freeze panes, zoom, gridlines, page setup |
| `rules.apply` | `rules.apply` | Data validation, conditional formatting |
| `name.define` | `name.define` | Define a named range |
| `name.update` | `name.update` | Update a named range |
| `name.delete` | `name.delete` | Delete a named range |

### Payload Conventions

Batch ops use the `{"kind": "...", "ops": [...]}` envelope:

```json
{
  "kind": "structure.insert_rows",
  "ops": [{
    "kind": "insert_rows",
    "sheet_name": "Sheet1",
    "at_row": 5,
    "count": 2
  }]
}
```

Single-object ops (names, replace-in-formulas) use flat payloads:

```json
{
  "kind": "name.define",
  "name": "SalesTotal",
  "refers_to": "Sheet1!$C$100"
}
```

## Preconditions

Staged artifacts can include `preconditions` for CAS-style assertions on cell values.
If a precondition fails at apply time, the operation is rejected — preventing stale
edits when multiple agents or manual edits are interleaved.

```json
{
  "preconditions": {
    "cell_matches": [
      { "address": "Sheet1!A1", "value": "Name" },
      { "address": "Sheet1!B2", "value": 10.0 }
    ],
    "workbook_hash_before": "sha256:abc123..."
  }
}
```

- `cell_matches` — assert that specific cells hold expected values before applying
- `workbook_hash_before` — assert the entire workbook hash matches (stronger but more expensive)

Both are evaluated after the CAS HEAD check and before sealing the event.

## Structural Edit Recipes

### Inserting Rows in Modeled Zones

**Prefer `clone_row` over raw `insert_rows`** in modeled zones (zones with formulas):

```json
{
  "kind": "structure.clone_row",
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
  "kind": "structure.copy_range",
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
- **Never** apply without checking `dry_run_impact` in the staging response
- **Never** materialize just to read — use `--session` flag on read commands instead
