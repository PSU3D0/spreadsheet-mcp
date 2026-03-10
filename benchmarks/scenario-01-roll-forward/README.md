# scenario-01-roll-forward

Sanitized benchmark anchor for a moderate spreadsheet workflow.

## Goal

Change a core input, recalculate the workbook, and prove the downstream output deltas in a single verification loop.

## Workflow under test

1. `asp edit ... --output draft.xlsx`
2. `asp recalculate draft.xlsx --output result.xlsx`
3. `asp verify base.xlsx result.xlsx --targets Summary!B1,Summary!B2 --sheet Summary`

## Regression gates

See `budget.json` for:
- tool-call budget
- wall-time budget
- output-size budget
- correctness postconditions
