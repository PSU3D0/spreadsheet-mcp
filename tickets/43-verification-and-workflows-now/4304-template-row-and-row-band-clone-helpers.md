# Ticket: 4304 Template Row and Row-Band Clone Helpers

## Depends On
- Existing clone-row / structure helpers
- 4303 (recommended)
- [Design: template row + row-band clone helpers](./designs/4304-template-row-and-row-band-clone-helpers.md)

## Why
Many practical workflows are template-driven: insert a new row based on a nearby modeled row or small row band, preserve formulas/styles, and then patch a few values.

## Owner / Effort / Risk
- Owner (proposed): Structure Ops / Workflow Helpers
- Effort: M
- Risk: Med

## Scope
Promote row cloning into a higher-level workflow helper.

### Candidate capabilities
- `clone-template-row` helper for one template row
- `clone-row-band` helper for a contiguous source row band
- extend formulas relative to the new position
- preserve styles/validation/merges where safe
- return explicit follow-up patch targets
- support preview-first dry-run responses with warnings and confidence metadata

### Recommended phasing
1. `clone-template-row` first slice
2. `clone-row-band` second slice
3. hardening for merge-policy / patch-target heuristics / warning contracts

## Non-Goals
- Full semantic understanding of every workbook pattern.
- Multi-sheet workflow automation in this ticket.

## Tests
Treat tests as the executable contract for clone behavior.

### Parser / help
- anchor flags require exactly one of `--before`, `--after`, or `--insert-at`
- invalid row-band syntax fails cleanly
- help surface documents patch-target and merge-policy controls

### Planning / dry-run
- relative formulas shift correctly
- inserted row ranges / repeated block ranges are explicit and stable
- patch-target suggestions are explicit and deterministic
- adjacent sum targets are reported explicitly
- empty arrays remain explicit where meaningful (`formula_targets`, `likely_patch_targets`, `adjacent_sum_targets`, `warnings`)

### Mutation / regression
- styles and row structure are preserved where supported
- validation survives where supported
- unsafe merge-boundary scenarios emit warnings or fail cleanly depending on policy
- calculated data rows are not misclassified as summary/footer rows during clone planning
- documentation examples should execute as tests where practical

## Definition of Done
- A user can create new modeled rows from nearby templates without reconstructing the whole pattern manually.
