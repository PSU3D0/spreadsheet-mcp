# Ticket: 9001 (Deferred) UNO/LibreOffice Editing for Sort/Filter + Excel-Parity Ops

Deferred by design: record the work but do not start now.

## Why
Applying sort/filter like a human (reordering/hiding rows) and performing Excel-like structural refactors across tables/structured references are not well-supported by OOXML state alone.

## Scope (Future)
- Provide LO/UNO-backed operations for:
  - apply sort
  - apply filter criteria
  - structural edits that rewrite CF/DV/table/structured refs robustly

## Notes
- Keep tool surface minimal; ideally a single `office_automation_batch` tool if pursued.
