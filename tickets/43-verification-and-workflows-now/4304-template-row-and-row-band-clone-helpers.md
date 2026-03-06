# Ticket: 4304 Template Row and Row-Band Clone Helpers

## Depends On
- Existing clone-row / structure helpers
- 4303 (recommended)

## Why
Many practical workflows are template-driven: insert a new row based on a nearby modeled row or small row band, preserve formulas/styles, and then patch a few values.

## Owner / Effort / Risk
- Owner (proposed): Structure Ops / Workflow Helpers
- Effort: M
- Risk: Med

## Scope
Promote row cloning into a higher-level workflow helper.

### Candidate capabilities
- clone one template row before/after a target row
- clone a contiguous row band
- extend formulas relative to the new position
- preserve styles/validation/merges where safe
- return explicit follow-up patch targets

## Non-Goals
- Full semantic understanding of every workbook pattern.
- Multi-sheet workflow automation in this ticket.

## Tests
- Relative formulas shift correctly.
- Styles and row structure are preserved where supported.
- Unsafe clone scenarios emit warnings or fail cleanly.

## Definition of Done
- A user can create new modeled rows from nearby templates without reconstructing the whole pattern manually.
