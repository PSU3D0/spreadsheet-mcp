# Ticket: 4302 Grouped Diff + Change Summary

## Depends On
- Existing file diff / changeset infrastructure
- 4301 (recommended, not strictly required)

## Why
Raw cell-level diffs are powerful but too noisy for many practical edit flows.

Agents need grouped summaries closer to how humans reason about spreadsheet edits.

## Owner / Effort / Risk
- Owner (proposed): Diff / Verification
- Effort: M
- Risk: Med

## Scope
Add grouped diff output modes and summary views.

### Candidate groupings
- inserted/deleted row or column blocks
- contiguous formula rewrite bands
- named-range changes
- target-cell deltas
- recalc-result changes vs direct formula edits

### UX goals
- summary-first output
- ability to drill into one group
- filters for noisy diff types

## Non-Goals
- Removing the existing low-level diff modes.
- Designing a visual renderer.

## Tests
- Grouping logic is deterministic.
- Summary counts match underlying detailed diffs.
- Filters do not hide direct edit changes unintentionally.

## Definition of Done
- Moderate workbook edits produce an interpretable grouped summary rather than an overwhelming flat diff.
