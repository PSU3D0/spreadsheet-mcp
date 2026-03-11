# CLI Batch Write Skill — Canonical Batch Workflow

Use this skill for stateless batch writes through `asp`.

## Short checklist

1. Choose the right command family
2. Discover the exact payload shape
3. Pick exactly one mutation mode
4. Dry-run first when the change is non-trivial
5. Recalculate if formulas were affected
6. Diff and inspect critical cells

## Discoverability first

When unsure of payload shape, ask the CLI directly:

```bash
asp schema transform-batch
asp example transform-batch
asp schema structure-batch
asp example structure-batch
asp schema session-op transform.write_matrix
asp example session-op transform.write_matrix
```

## Mutation modes

Every batch write command requires exactly **one** of:

| Mode | Flag | Behavior |
|------|------|----------|
| Dry run | `--dry-run` | Validate without mutation |
| In-place | `--in-place` | Atomically replace source file |
| Output | `--output <PATH>` | Write to a new file (`--force` to overwrite) |

## Canonical commands

```bash
asp transform-batch workbook.xlsx --ops @transform_ops.json --dry-run
asp style-batch workbook.xlsx --ops @style_ops.json --in-place
asp apply-formula-pattern workbook.xlsx --ops @formula_ops.json --in-place
asp structure-batch workbook.xlsx --ops @structure_ops.json --dry-run --impact-report --show-formula-delta
asp column-size-batch workbook.xlsx --ops @column_ops.json --in-place
asp sheet-layout-batch workbook.xlsx --ops @layout_ops.json --in-place
asp rules-batch workbook.xlsx --ops @rules_ops.json --in-place
asp replace-in-formulas workbook.xlsx Sheet1 --find '$64' --replace '$65' --dry-run
```

## Exact payload conventions

### Most batch commands
Use a top-level `ops` array:

```json
{
  "ops": [
    { "kind": "...", ... }
  ]
}
```

### `column-size-batch`
Preferred canonical form includes top-level `sheet_name`:

```json
{
  "sheet_name": "Sheet1",
  "ops": [
    {
      "target": { "kind": "columns", "range": "A:C" },
      "size": { "kind": "width", "width_chars": 18.0 }
    }
  ]
}
```

## Exact examples

### `transform-batch`

```json
{
  "ops": [{
    "kind": "fill_range",
    "sheet_name": "Sheet1",
    "target": { "kind": "range", "range": "B2:B4" },
    "value": "0"
  }]
}
```

### `style-batch`

```json
{
  "ops": [{
    "sheet_name": "Sheet1",
    "target": { "kind": "range", "range": "B2:B2" },
    "patch": { "font": { "bold": true } }
  }]
}
```

### `structure-batch`

```json
{
  "ops": [{
    "kind": "rename_sheet",
    "old_name": "Summary",
    "new_name": "Dashboard"
  }]
}
```

## Post-write checklist

1. Run `asp recalculate` if formulas changed
2. Run `asp verify <baseline> <current> --targets <Sheet!A1,...>` for explicit proof (target classifications + new/resolved/preexisting errors). Use `--errors-only` for a sheet-scoped QA pass or `--targets-only` for pure target proof.
3. Run `asp diff` to confirm intent. Add `--exclude-recalc-result` when you want a lower-noise review focused on direct edits.
4. Use `asp inspect-cells` on critical cells/ranges
5. Use `asp recalculate --changed-cells` for a change summary

## Session integration

Use sessions for multi-step edits or anything that needs undo/redo, branching, or staged apply:

```bash
asp session start --base workbook.xlsx --workspace <dir>
asp example session-op transform.write_matrix
asp session op --session <id> --ops @edits.json --workspace <dir>
asp session apply --session <id> <staged_id> --workspace <dir>
asp session materialize --session <id> --output result.xlsx --workspace <dir>
```

### Session kind mapping

| Batch command | Session `kind` |
|---|---|
| `transform-batch` | `transform.clear_range`, `transform.fill_range`, `transform.replace_in_range` |
| write_matrix | `transform.write_matrix` |
| `structure-batch` | `structure.insert_rows`, `structure.clone_row`, etc. |
| `style-batch` | `style.apply` |
| `apply-formula-pattern` | `formula.apply_pattern` |
| `replace-in-formulas` | `formula.replace_in_formulas` |
| `column-size-batch` | `column.size` |
| `sheet-layout-batch` | `layout.apply` |
| `rules-batch` | `rules.apply` |
| named range CRUD | `name.define`, `name.update`, `name.delete` |

## Hard rules

- Never mix mutation modes.
- Never guess a payload shape when `asp schema` / `asp example` can tell you.
- Always dry-run structure changes first.
- Recalculate after formula-affecting writes.
