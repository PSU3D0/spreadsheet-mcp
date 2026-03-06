# Ticket: 4402 JS SDK Session / Workflow Objects + Verification Helpers

## Depends On
- 4301
- 4401
- [Design: JS SDK workflow object model](./designs/4402-sdk-session-workflow-object-model.md)

## Why
The SDK currently normalizes transport/runtime differences, but it does not yet provide a strong workflow-oriented programming model.

## Owner / Effort / Risk
- Owner (proposed): SDK / JS Surface
- Effort: L
- Risk: Med

## Scope
Add a higher-level SDK layer that makes stateful workflows easier to express.

### Candidate capabilities
- workbook/session handles
- workflow-oriented methods
- verification helpers
- capability-gated advanced methods

### Example direction
- open workbook/session
- mutate
- recalc
- verify targets/errors
- export/materialize

## Non-Goals
- Creating a fourth semantics fork.
- Forcing CLI concepts directly into the SDK.

## Tests
- Shared workflow helpers behave consistently across supported backends.
- Unsupported backend features fail as capability errors.
- Verification helpers normalize output cleanly.

## Definition of Done
- SDK is useful at the workflow level, not only the RPC-shape level.
