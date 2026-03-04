# CLI Batch Write Skill — Stateless Batch Operations

## Overview

This skill describes batch write operations available through the `asp` CLI.
All batch commands follow a consistent pattern with mutation modes and JSON
payloads.

## Mutation Modes

Every batch write command requires exactly **one** of:

| Mode | Flag | Behavior |
|------|------|----------|
| Dry Run | `--dry-run` | Validate without mutation |
| In-Place | `--in-place` | Atomically replace source file |
| Output | `--output <PATH>` | Write to new file (add `--force` to overwrite) |

## Batch Commands

### Transform Batch

Cell value/formula range operations:

```bash
asp transform-batch workbook.xlsx --ops @transform_ops.json --dry-run
```

Operations: `clear_range`, `fill_range`, `replace_in_range`, `write_matrix`

### Style Batch

Appearance modifications:

```bash
asp style-batch workbook.xlsx --ops @style_ops.json --in-place
```

### Structure Batch

Spreadsheet shape changes (rows, columns, sheets):

```bash
# Always dry-run first with impact analysis
asp structure-batch workbook.xlsx --ops @structure_ops.json --dry-run \
    --impact-report --show-formula-delta

# Then apply
asp structure-batch workbook.xlsx --ops @structure_ops.json --in-place
```

Operations: `insert_rows`, `delete_rows`, `insert_cols`, `delete_cols`,
`clone_row`, `merge_cells`, `unmerge_cells`, `rename_sheet`, `create_sheet`,
`delete_sheet`, `copy_range`, `move_range`

### Formula Pattern

Apply relative formula patterns:

```bash
asp apply-formula-pattern workbook.xlsx --ops @formula_ops.json --in-place
```

### Column Size

Adjust column widths:

```bash
asp column-size-batch workbook.xlsx --ops @column_ops.json --in-place
```

### Sheet Layout

Page setup and freeze panes:

```bash
asp sheet-layout-batch workbook.xlsx --ops @layout_ops.json --in-place
```

### Rules

Data validation and conditional formatting:

```bash
asp rules-batch workbook.xlsx --ops @rules_ops.json --in-place
```

### Replace in Formulas

Find/replace in formula text:

```bash
asp replace-in-formulas workbook.xlsx Sheet1 --find '$64' --replace '$65' --dry-run
```

## Payload Format

All `--ops` payloads use `@<path>` file references with a top-level `ops` array:

```json
{
  "ops": [
    { "kind": "...", ... }
  ]
}
```

Use `--print-schema` on any batch command to see the full JSON schema.

## Post-Write Checklist

After any write operation:

1. Run `asp recalculate` if formulas were affected
2. Run `asp diff` to verify changes match intent
3. Use `asp inspect-cells` to spot-check critical cells
4. Use `asp recalculate --changed-cells` for a summary of what changed

## Session Integration

For complex multi-step edits, use the session workflow instead of raw batch commands.
**All batch operation families** are supported through sessions and replay correctly
during materialization.

### Session Workflow

```bash
asp session start --base workbook.xlsx --workspace <dir>
asp session op --session <id> --ops @edits.json --workspace <dir>
asp session apply --session <id> <staged_id> --workspace <dir>
asp session materialize --session <id> --output result.xlsx --workspace <dir>
```

### Session Payload Convention

When using `session op`, the ops payload must include a `"kind"` field so the
session can route to the correct replay handler:

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

### Supported Session Op Kinds

| Batch command | Session `kind` value |
|---|---|
| `transform-batch` | `transform.clear_range`, `transform.fill_range`, `transform.replace_in_range` |
| `edit` / write_matrix | `transform.write_matrix` or `edit.batch` |
| `structure-batch` | `structure.insert_rows`, `structure.clone_row`, etc. |
| `style-batch` | `style.apply` |
| `apply-formula-pattern` | `formula.apply_pattern` |
| `replace-in-formulas` | `formula.replace_in_formulas` |
| `column-size-batch` | `column.size` |
| `sheet-layout-batch` | `layout.apply` |
| `rules-batch` | `rules.apply` |
| `define-name` | `name.define` |
| `update-name` | `name.update` |
| `delete-name` | `name.delete` |

### Dry-Run Impact

Staging (`session op`) computes `dry_run_impact` and returns it in the response:

```json
{
  "staged_id": "stg_...",
  "dry_run_impact": {
    "cells_changed": 9,
    "formulas_rewritten": 45,
    "shifted_spans": [...],
    "warnings": [],
    "boundary_warnings": []
  }
}
```

For structure ops, the impact analysis uses `compute_structure_impact()` and
reports absolute reference warnings and shifted spans. For write_matrix ops,
impact is the cell count from payload dimensions.

### Session-Aware Reads

After applying operations, verify results without materializing:

```bash
asp range-values workbook.xlsx Sheet1 A1:C10 --format rows \
    --session <id> --session-workspace <dir>
```

This provides undo/redo, branching, and atomic apply semantics.
See `SAFE_EDITING_SKILL.md` for the full session workflow.
