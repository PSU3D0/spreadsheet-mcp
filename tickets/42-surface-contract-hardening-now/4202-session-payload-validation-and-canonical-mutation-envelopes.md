# Ticket: 4202 Session Payload Validation + Canonical Mutation Envelopes

## Depends On
- Existing `session op` staging/apply pipeline
- Existing write op families in `core.session.*`
- [Design: canonical session op contract](./designs/4202-session-op-canonical-contract.md)

## Why
The benchmark exposed a serious footgun: mutation payloads were not obvious enough, and wrapped vs direct shapes could appear to partially work.

That is a contract failure.

## Owner / Effort / Risk
- Owner (proposed): Session / Write Surface
- Effort: L
- Risk: High

## Scope
Define and enforce canonical mutation envelopes for session operations.

### Contract Work
- For each supported session op kind, define one canonical payload shape.
- Reject ambiguous or malformed payloads with actionable validation errors.
- If shorthand forms are intentionally supported, normalize them explicitly and report the normalization.

### CLI
- `session op` validates payload shape before staging.
- Error output identifies:
  - op kind
  - expected shape
  - unexpected fields
  - example fix path

### MCP / SDK
- Shared validation behavior where semantics are shared.
- Optional schema/example exposure for clients.

### Backward Compatibility
- Decide per op whether to:
  - fail immediately
  - warn for one release, then fail
- Document the deprecation path explicitly.

## Non-Goals
- Designing brand new write primitives.
- Reworking the entire session lifecycle model.

## Tests
### Positive
- Canonical payloads stage/apply/replay correctly.
- Shared payloads behave consistently across CLI/MCP/SDK where applicable.

### Negative
- Wrapped vs direct shape mistakes fail loudly.
- Unknown fields do not silently disappear.
- Ambiguous partially valid payloads are rejected.

### Regression
- Replay behavior matches documented examples exactly.

## Definition of Done
- Session mutation inputs are canonical, strict, and test-backed.
- An agent cannot accidentally get partial success from the wrong shape.
