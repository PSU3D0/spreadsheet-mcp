# Ticket: 4303 Region Append, CSV Append, and Footer-Aware Insert Primitives

## Depends On
- Existing region/layout detection
- Existing structure and transform write paths
- [Design: region append + footer-aware insert helpers](./designs/4303-region-append-and-footer-aware-insert.md)

## Why
Moderate spreadsheet tasks still require too many primitive steps for common workflows like appending rows into a detected block or inserting before a subtotal/footer row.

## Owner / Effort / Risk
- Owner (proposed): Workflow Helpers / Write Surface
- Effort: L
- Risk: High

## Scope
Add generic, domain-neutral helpers for common row-oriented workflows.

### Candidate helpers
- append rows into detected region
- append CSV into detected region
- insert rows before footer/subtotal row
- rebuild/extend affected totals where policy allows

### Output goals
- explicit plan/preview mode
- explicit affected-range summary
- explicit warnings when confidence is low

## Non-Goals
- Domain-specific business logic.
- Full structural workbook model.

## Tests
- Append helpers find the right insertion point for representative region shapes.
- Footer-aware insert avoids inserting after subtotal/footer rows.
- Preview mode reports affected totals and warnings clearly.

## Definition of Done
- Common append/insert workflows can be expressed in a few high-level steps rather than many primitive edits.
