# Ticket: 1002 sheet_layout_batch (Print Setup)

## Why (Human Operator Replacement)
Human operators prepare printable/shareable artifacts: margins, orientation, scaling, print area, and page breaks. Without these, "board-ready" outputs are incomplete.

## Scope
Extend `sheet_layout_batch` with print/page ops:
- `set_page_margins`
- `set_page_setup`
- `set_print_area` (via defined name `_xlnm.Print_Area`)
- `set_page_breaks` (rows/cols)

## Non-Goals
- No print titles (rows to repeat) yet.
- No headers/footers beyond minimal defaults.

## Proposed Tool Surface
```json
{
  "fork_id": "fork-123",
  "ops": [
    {"kind":"set_page_setup","sheet_name":"Dashboard","orientation":"landscape","fit_to_width":1,"fit_to_height":1},
    {"kind":"set_page_margins","sheet_name":"Dashboard","left":0.25,"right":0.25,"top":0.5,"bottom":0.5},
    {"kind":"set_print_area","sheet_name":"Dashboard","range":"A1:G30"},
    {"kind":"set_page_breaks","sheet_name":"Dashboard","row_breaks":[31],"col_breaks":[8]}
  ],
  "mode": "apply"
}
```

## Implementation Notes
- Use umya PageMargins/PageSetup/RowBreaks/ColumnBreaks.
- Print area:
  - create/update defined name `_xlnm.Print_Area` scoped to the sheet.
- Validate:
  - `fit_to_width/height` positive ints
  - break indices are >= 1

## Tests
- Set print area creates defined name and persists.
- Margins/setup/breaks persist after write + read.
- Preview staging if supported.

## Definition of Done
- Print settings visible in Excel print preview.
- No tool proliferation; stays within sheet_layout_batch.
