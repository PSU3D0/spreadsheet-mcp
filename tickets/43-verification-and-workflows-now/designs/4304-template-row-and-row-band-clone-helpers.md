# Design: 4304 Template Row + Row-Band Clone Helpers

## Problem

Many practical spreadsheet workflows are template-driven:

- add one more modeled line item like the row above
- repeat a small row band
- preserve formulas/styles/validation where safe
- patch only the intended input cells afterward

Today the repo already has lower-level structure primitives such as `clone_row`, `insert_rows`, and `copy_range`, but agents still need too much manual reasoning to turn them into a reliable workflow.

## Goals

- Add preview-first workflow helpers for template row and row-band cloning.
- Compile down to existing structure operations where possible.
- Return explicit follow-up patch targets so the next mutation step is obvious.
- Surface uncertainty via warnings / confidence metadata instead of silently guessing.
- Treat tests as part of the contract: behavior should be documented as executable fixtures.

## Non-Goals

- Automatic template discovery in the first slice.
- Cross-sheet cloning.
- Full semantic understanding of workbook intent.
- Perfect merge preservation for every workbook pattern.

## Existing substrate to reuse

The current structure layer already supports:

- `structure.clone_row`
- `structure.insert_rows`
- formula rewrite for row insertion
- defined-name rewrite for row insertion
- optional adjacent `SUM(...)` expansion

This ticket should promote those primitives into a higher-level helper, not fork a new structural engine.

## Proposed CLI helpers

### 1. `clone-template-row`

```bash
asp clone-template-row <file> \
  --sheet <sheet> \
  --source-row <row> \
  (--before <row> | --after <row> | --insert-at <row>) \
  [--count <n>] \
  [--expand-adjacent-sums] \
  [--patch-targets likely-inputs|all-non-formula|none] \
  [--merge-policy safe|strict] \
  (--dry-run | --in-place | --output <path>) \
  [--force]
```

#### Semantics
- Clone a single template row.
- Insert one or more copies.
- Shift formulas relative to the new location.
- Preserve styles / validation where supported by the existing structure layer.
- Report likely follow-up patch targets in dry-run output.

### 2. `clone-row-band`

```bash
asp clone-row-band <file> \
  --sheet <sheet> \
  --source-rows <start:end> \
  (--before <row> | --after <row> | --insert-at <row>) \
  [--repeat <n>] \
  [--expand-adjacent-sums] \
  [--patch-targets likely-inputs|all-non-formula|none] \
  [--merge-policy safe|strict] \
  (--dry-run | --in-place | --output <path>) \
  [--force]
```

#### Semantics
- Clone a contiguous row band as a unit.
- Repeat the band one or more times.
- Preserve internal formula relationships within the band.
- Warn or fail for merge patterns that cross the source-band boundary.

## Anchor rules

Exactly one of:

- `--before ROW`
- `--after ROW`
- `--insert-at ROW`

### Meaning
- `--before 20` → insert at row 20
- `--after 20` → insert at row 21
- `--insert-at 20` → raw insertion row

## Patch-target modes

### `likely-inputs` (default)
Return cells in the inserted rows that are likely to need user/agent patching:

- non-formula cells
- non-header / non-footer-like cells
- includes intentionally blank input cells where template structure implies editable entry fields

### `all-non-formula`
Return all inserted non-formula cells.

### `none`
Return no patch-target suggestions.

## Merge policies

### `safe` (default)
- preserve merges fully contained within the source row / band
- do not reproduce merges that cross the source boundary
- emit warnings

### `strict`
- fail when any merge crosses the source boundary

## Dry-run contract

Helpers should return a stable, machine-consumable preview response.

### `clone-template-row` dry-run example

```json
{
  "mode": "dry_run",
  "file": "model.xlsx",
  "sheet_name": "Inputs",
  "helper_kind": "clone_template_row",
  "source_row": 12,
  "source_row_range": "12:12",
  "anchor_kind": "after",
  "anchor_row": 12,
  "insert_at_row": 13,
  "count": 2,
  "rows_inserted": 2,
  "inserted_row_range": "13:14",
  "expand_adjacent_sums": true,
  "patch_target_mode": "likely_inputs",
  "merge_policy": "safe",
  "template_summary": {
    "non_empty_cell_count": 8,
    "formula_cell_count": 3,
    "style_cell_count": 8,
    "validation_cell_count": 2,
    "merged_ranges_fully_contained": [],
    "merged_ranges_crossing_boundary": []
  },
  "formula_targets": ["D13", "E13", "F13", "D14", "E14", "F14"],
  "likely_patch_targets": ["A13", "B13", "C13", "A14", "B14", "C14"],
  "adjacent_sum_targets": ["F15"],
  "warnings": [],
  "confidence": "high",
  "confidence_reason": "template row cloned cleanly with no merge or layout conflicts",
  "would_change": true
}
```

### `clone-row-band` dry-run example

```json
{
  "mode": "dry_run",
  "file": "model.xlsx",
  "sheet_name": "Schedule",
  "helper_kind": "clone_row_band",
  "source_row_range": "20:22",
  "source_row_count": 3,
  "anchor_kind": "after",
  "anchor_row": 22,
  "insert_at_row": 23,
  "repeat": 2,
  "rows_inserted": 6,
  "inserted_row_range": "23:28",
  "inserted_blocks": [
    {"block_index": 0, "row_range": "23:25"},
    {"block_index": 1, "row_range": "26:28"}
  ],
  "expand_adjacent_sums": true,
  "patch_target_mode": "likely_inputs",
  "merge_policy": "safe",
  "template_summary": {
    "non_empty_cell_count": 21,
    "formula_cell_count": 8,
    "style_cell_count": 21,
    "validation_cell_count": 4,
    "merged_ranges_fully_contained": ["A20:A22"],
    "merged_ranges_crossing_boundary": []
  },
  "formula_targets": ["D23", "D24", "D25", "D26", "D27", "D28"],
  "likely_patch_targets": [
    "A23", "B23", "A24", "B24", "A25", "B25",
    "A26", "B26", "A27", "B27", "A28", "B28"
  ],
  "adjacent_sum_targets": ["F29"],
  "warnings": [],
  "confidence": "high",
  "confidence_reason": "row band is contiguous and all merges are fully contained within the cloned block",
  "would_change": true
}
```

### Warning-heavy dry-run example

```json
{
  "mode": "dry_run",
  "file": "model.xlsx",
  "sheet_name": "Report",
  "helper_kind": "clone_row_band",
  "source_row_range": "10:12",
  "insert_at_row": 20,
  "rows_inserted": 3,
  "inserted_row_range": "20:22",
  "expand_adjacent_sums": false,
  "patch_target_mode": "likely_inputs",
  "merge_policy": "safe",
  "template_summary": {
    "non_empty_cell_count": 14,
    "formula_cell_count": 4,
    "style_cell_count": 14,
    "validation_cell_count": 0,
    "merged_ranges_fully_contained": [],
    "merged_ranges_crossing_boundary": ["A9:A10", "C12:D13"]
  },
  "formula_targets": ["E20", "E21", "E22"],
  "likely_patch_targets": ["A20", "B20", "A21", "B21", "A22", "B22"],
  "adjacent_sum_targets": [],
  "warnings": [
    "merge A9:A10 crosses the source band boundary and will not be reproduced under merge_policy=safe",
    "merge C12:D13 crosses the source band boundary and will not be reproduced under merge_policy=safe"
  ],
  "confidence": "medium",
  "confidence_reason": "clone is structurally possible but merge-boundary conflicts require caution",
  "would_change": true
}
```

### Strict failure example

```json
{
  "code": "UNSAFE_CLONE_TEMPLATE",
  "message": "source rows 10:12 intersect merged ranges that cross the clone boundary",
  "details": {
    "sheet_name": "Report",
    "source_row_range": "10:12",
    "merge_policy": "strict",
    "blocking_merged_ranges": ["A9:A10", "C12:D13"]
  }
}
```

## Implementation phases

### Phase 1 — `clone-template-row`
- Add CLI command.
- Build preview/plan response.
- Compile to the existing `structure.clone_row` path.
- Add follow-up patch target reporting.
- Add merge-policy inspection (even if single-row cases are usually simple).

### Phase 2 — `clone-row-band`
- Add contiguous row-band helper.
- Implement band planning + preview metadata.
- Preserve formula relationships and repeated-band targeting.
- Add merge-boundary handling.

### Phase 3 — hardening
- Improve patch-target heuristics.
- Add richer warnings and confidence reasons.
- Consider future structural targeting once the core row/band contract is stable.

## Testing strategy

Treat tests as the public contract.

### 1. Parser / help tests
- parse single-row helper with `--after`
- parse single-row helper with `--before`
- parse band helper with `--repeat`
- help mentions `--patch-targets`
- help mentions `--merge-policy`
- reject invalid anchor combinations
- reject invalid band syntax
- reject zero `--count`
- reject zero `--repeat`

### 2. Plan-generation tests
- single-row plan computes insertion rows/ranges correctly
- repeated single-row clone computes all inserted targets correctly
- band plan computes repeated block ranges correctly
- formula targets are stable and explicit
- likely patch targets are stable and explicit
- adjacent sum targets are reported correctly
- empty template rows warn
- pure-formula template rows do not fabricate patch targets
- fully-contained merges are preserved under `safe`
- cross-boundary merges warn under `safe`
- cross-boundary merges fail under `strict`

### 3. Apply / mutation tests
- in-place single-row clone
- output-mode single-row clone
- formulas shift correctly
- styles survive where supported
- validation survives where supported
- row heights / structure survive where supported
- adjacent sums expand correctly
- repeated band clone preserves internal formula relationships

### 4. Regression tests
- last row contains calculated formula but is not a footer
- source row below insertion anchor
- source band contains blank spacer rows
- insertion directly below header row
- insertion directly above subtotal row
- malformed merge-boundary conflicts
- formulas with named ranges remain stable

### 5. Documentation-as-tests
- help examples stay in sync with behavior
- README examples execute against local fixtures
- dry-run response fields remain explicit, including empty arrays where meaningful

## Suggested test counts

A robust first pass should roughly target:

- 8–10 parser/help tests
- 18–20 plan tests
- 10–12 apply/integration tests
- 10–12 regression tests

Approximate total: 46–54 tests.

## Success criteria

- Agents can create new modeled rows/bands in a small number of high-level steps.
- Preview responses clearly identify formula targets, patch targets, and safety concerns.
- Ambiguous merge/layout cases warn or fail explicitly.
- The behavior is documented by executable tests rather than only prose.
