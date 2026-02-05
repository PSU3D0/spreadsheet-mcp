# Ticket: 1001 sheet_layout_batch (Freeze Panes + Basic View)

## Why (Human Operator Replacement)
Freezing headers/label columns is one of the most common human actions when making sheets usable. Without it, AI-generated workbooks feel unfinished and are harder to audit.

## Scope
- Add a new write tool: `sheet_layout_batch`.
- Implement v1 ops:
  - `freeze_panes`
  - `set_zoom`
  - `set_gridlines` (view gridlines)
- Support preview/apply (consistent with other write tools).

## Non-Goals
- No UNO/LibreOffice editing.
- No comprehensive sheetView flag coverage yet (row/col headings etc. deferred).

## Proposed Tool Surface
Tool: `sheet_layout_batch`

Canonical request:
```json
{
  "fork_id": "fork-123",
  "ops": [
    {"kind":"freeze_panes","sheet_name":"Dashboard","freeze_rows":1,"freeze_cols":1},
    {"kind":"set_zoom","sheet_name":"Dashboard","zoom_percent":110},
    {"kind":"set_gridlines","sheet_name":"Dashboard","show":false}
  ],
  "mode": "preview",
  "label": "Make dashboard readable"
}
```

Response:
- Standard ChangeSummary with `affected_sheets`, `counts`, `warnings`.

Warnings:
- `WARN_FREEZE_PANES_TOPLEFT_DEFAULTED` if top_left_cell inferred.

## Implementation Notes
- Implement in `src/tools/fork.rs` following existing batch patterns.
- Underlying umya-spreadsheet types:
  - `SheetView`, `Pane`, `Selection`.
- Policy:
  - if `top_left_cell` omitted: infer as next cell after frozen rows/cols (e.g., freeze_rows=1, freeze_cols=1 -> topLeftCell=B2).
- Ensure writer emits correct `pane.state` for freeze.

## Tests
- Unit tests creating a workbook, applying freeze_panes, re-reading:
  - pane xSplit/ySplit correct
  - state indicates freeze
  - topLeftCell correct
- Preview mode test: staged change applies and persists.

## Definition of Done
- Excel opens with correct frozen panes.
- Preview/apply workflow matches other write tools.

## Rollout Notes
- Update README with a single example; do not add additional layout tools.
