# Ticket: 0003 Recalc Needed State + Stale Formula Warnings

## Why (Human Operator Replacement)
Humans understand that changing inputs requires recalculation; agents often validate outputs too early. This is the highest-risk "silent wrongness" class: calls succeed but reported results are incorrect.

## Scope
- Track whether a fork likely needs recalculation.
- Surface this state in write responses and warn in read responses when appropriate.

## Non-Goals
- Do not implement a new calc engine.
- Do not change formula caching semantics beyond existing behavior.

## Proposed Tool Surface
- Add `recalc_needed: bool` to fork metadata and expose via:
  - write responses `summary.flags` or warnings (preferred: structured field)
  - optional field on `list_forks` output
- Add warning code for reads when operating on a fork with `recalc_needed=true`:
  - `WARN_STALE_FORMULAS` (include message and suggested action: call `recalculate`).

Heuristics for setting `recalc_needed=true`:
- any `edit_batch` that sets formulas or edits cells that might be referenced (conservative: any edit_batch)
- `apply_formula_pattern`
- `structure_batch` ops that insert/delete rows/cols or rename sheets
- `transform_batch` when include_formulas is true or target intersects known formula regions (if known)

Heuristics for reads:
- when reading from a fork workbook id and formulas are present in the requested range/region.

## Implementation Notes
- Extend fork state in `src/fork.rs`:
  - add `recalc_needed: bool` on fork record
  - set true in mutation path (`with_fork_mut` wrapper or at tool handler level)
  - set false only on successful `recalculate`
- Extend `ChangeSummary` to include `flags: Vec<String>` or structured bool.
- Read tools in `src/tools/mod.rs`:
  - detect fork ids (if workbook_id namespace indicates fork)
  - if recalc_needed, add warning in response payload.

## Tests
- `edit_batch` sets `recalc_needed=true`.
- `recalculate` sets `recalc_needed=false`.
- `range_values` or `read_table` on fork with recalc_needed emits warning.

## Definition of Done
- Agents have a clear, machine-readable signal when numbers may be stale.
- No regression for non-fork workbook reads.

## Rollout Notes
- Conservative warnings are acceptable; false positives are safer than false negatives.
