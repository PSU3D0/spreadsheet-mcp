# Safe Editing Skill — Canonical Session Workflow

Use this skill when making non-trivial workbook edits.

## Short checklist

1. **Start a session**
2. **Explore before mutating**
3. **Discover the exact payload shape**
4. **Stage first**
5. **Inspect `dry_run_impact`**
6. **Apply only after review**
7. **Verify from session state**
8. **Materialize only at the end**

## Canonical workflow

```bash
# 1) Start
asp session start --base <workbook.xlsx> --label "edit session" --workspace <dir>

# 2) Explore
asp read names <workbook.xlsx> --session <session_id> --session-workspace <dir>
asp analyze formula-trace <workbook.xlsx> <Sheet> <Cell> precedents --depth 2 \
  --session <session_id> --session-workspace <dir>
asp read values <workbook.xlsx> <Sheet> <range> --format rows \
  --session <session_id> --session-workspace <dir>

# 3) Discover exact payload contract
asp schema session op transform.write_matrix
asp example session op transform.write_matrix

# 4) Stage
asp session op --session <session_id> --ops @edits.json --workspace <dir>

# 5) Apply
asp session apply --session <session_id> <staged_id> --workspace <dir>

# 6) Verify
asp read cells <workbook.xlsx> <Sheet> <targets> \
  --session <session_id> --session-workspace <dir>
asp session materialize --session <session_id> --output <temp.xlsx> --workspace <dir>
asp verify proof <base.xlsx> <temp.xlsx> --targets <Sheet!A1,...>
asp verify diff <base.xlsx> <temp.xlsx> --details --limit 50 --exclude-recalc-result

`asp verify` is the summary-first proof step: check target classifications plus new/resolved/preexisting errors, then use `asp diff` when you need deeper detail. Use `--errors-only` for a sheet-scoped QA pass or `--targets-only` when you only want explicit target proof. Add `--exclude-recalc-result` when you want a lower-noise change review focused on direct edits.

# 7) Finalize
asp session materialize --session <session_id> --output final_result.xlsx --workspace <dir>
```

## Exact session payload rules

- Every `session op` payload must include a top-level **`kind`**.
- `transform.write_matrix` is a **flat payload**.
- Batch families use a **top-level `kind` + `ops` array**.
- `name.*` and `formula.replace_in_formulas` are **flat payloads**.
- `edit.batch` is **not supported** through `session op`; use `transform.write_matrix` instead.

## Exact examples

### `transform.write_matrix`

```json
{
  "kind": "transform.write_matrix",
  "sheet_name": "Sheet1",
  "anchor": "B7",
  "rows": [[{"v": "Revenue"}, {"v": 100}]],
  "overwrite_formulas": false
}
```

### `structure.insert_rows`

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

### `name.define`

```json
{
  "kind": "name.define",
  "name": "SalesTotal",
  "refers_to": "Sheet1!$C$100",
  "scope": "workbook"
}
```

## Structural edits

Before insert/delete row/column operations, always preflight:

```bash
asp analyze ref-impact <workbook.xlsx> --ops @structure_ops.json --show-formula-delta
```

Prefer `clone_row` over raw `insert_rows` in modeled zones.

## What to inspect in `dry_run_impact`

- `cells_changed`
- `formulas_rewritten`
- `shifted_spans`
- `boundary_warnings`

If structural edits are involved, do **not** apply until these look sane.

## Recovery

```bash
asp session undo --session <session_id> --workspace <dir>
asp session checkout --session <session_id> <op_id> --workspace <dir>
asp session fork --session <session_id> alt-approach --from <op_id> --workspace <dir>
```

## Hard rules

- Never mutate first and inspect later.
- Never guess a session payload shape; use `asp schema ...` / `asp example ...`.
- Never apply structural edits without preflight impact analysis.
- Never materialize just to read; use `--session` reads first.
