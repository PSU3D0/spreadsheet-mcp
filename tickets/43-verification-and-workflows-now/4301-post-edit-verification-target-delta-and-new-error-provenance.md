# Ticket: 4301 Post-Edit Verification: Target Delta + New Error Provenance

## Depends On
- Existing session/checkpoint or fork/change comparison primitives
- Existing read/analysis surfaces
- [Design: post-edit verification + provenance](./designs/4301-post-edit-verification-and-provenance.md)

## Why
After edits and recalc, agents currently spend too much effort answering:

- what changed?
- what newly broke?
- was this already broken before?
- did my target outputs move the way I expected?

This should be first-class product behavior.

## Owner / Effort / Risk
- Owner (proposed): Verification / Recalc UX
- Effort: L
- Risk: High

## Scope
Add a verification surface that compares a baseline and current state and reports meaningful post-edit proof.

### Required outputs
- target cell deltas
- newly introduced errors
- baseline/pre-existing errors
- optional named-range deltas
- optional dependency-scoped impact summary

### Surfaces
- CLI verification command(s)
- MCP tool(s)
- SDK normalized response/helpers

## Non-Goals
- Full structural workbook model.
- Replacing generic diff entirely.

## Tests
### Positive
- Detect changed targets correctly.
- Separate pre-existing errors from newly introduced errors.
- Handle no-change cases explicitly.

### Negative
- Verification should not silently collapse missing baseline/current references.
- Unsupported comparison scopes should fail with guidance.

## Definition of Done
- A user can prove the effect of an edit without reconstructing the answer manually from raw diffs and recalc output.
