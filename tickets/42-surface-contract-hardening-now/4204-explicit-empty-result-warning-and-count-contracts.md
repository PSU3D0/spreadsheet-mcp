# Ticket: 4204 Explicit Empty Result, Warning, and Count Contracts

## Depends On
- Existing read and mutation response models across CLI/MCP/SDK

## Why
Agents should never have to infer "no results" or "probably nothing changed" from missing fields.

Empty results, warnings, and mutation counts need explicit contracts.

## Owner / Effort / Risk
- Owner (proposed): Shared Response Surface
- Effort: M
- Risk: Med

## Scope
Standardize explicit empty and summary semantics across shared surfaces.

### Response Contract Work
- Return explicit empty arrays/objects where appropriate:
  - `matches: []`
  - `warnings: []`
  - `changed_targets: []`
  - `new_errors: []`
- Standardize count fields and their semantics:
  - requested vs matched vs changed vs applied
- Distinguish informational warnings from blockers.

### Surface Alignment
- CLI JSON output
- MCP tool output
- SDK normalized output

## Non-Goals
- Full response-shape redesign for every existing tool.
- Changing human-readable prose output modes.

## Tests
- No-match scenarios produce explicit empty arrays.
- No-op mutation previews are explicit.
- Warnings/blockers are distinguishable in stable fields.
- SDK normalization preserves explicit empties.

## Definition of Done
- Agents no longer infer absence from missing fields in common read/write/verify flows.
