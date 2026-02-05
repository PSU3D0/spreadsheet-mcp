# Ticket: 9002 (Deferred) Extend umya-spreadsheet for Structural Rewrite Correctness

Deferred by design.

## Why
To approach true human replacement, structural edits must adjust:
- data validations
- conditional formatting formulas
- table ranges and structured references
- workbook-scope defined names

## Scope (Future)
- Implement AdjustmentCoordinate traits for validations/tables/CF rules.
- Add structured reference parser/rewriter or defer to UNO.
