# Design: 4202 Canonical Session Op Contract

## Problem

`session op` currently has too much shape ambiguity across operation kinds. In practice, agents can confuse:

- direct payloads vs wrapped payloads
- shorthand examples vs replay payloads
- staging-time shapes vs apply-time shapes

That creates the worst possible failure mode: a mutation that appears to partly work.

## Goals

- Define one canonical input contract per supported session op kind.
- Reject malformed or ambiguous payloads before staging.
- Make example generation and schema discovery derive from the same source of truth.
- Keep cross-surface semantics aligned where the underlying op family is shared.

## Non-Goals

- Designing new workflow helpers.
- Replacing the session lifecycle.
- Solving all CLI ergonomics in this design.

## Proposed model

### 1) Per-kind canonical spec registry

Introduce a registry mapping `op_kind -> spec`.

Each spec should define:

- canonical payload shape
- optional accepted shorthands
- normalization rules (if any)
- validation rules
- example payload
- response preview/count semantics

This registry becomes the source of truth for:

- CLI validation
- MCP validation (where applicable)
- schema/example discovery
- docs/examples

### 2) Strict validation phases

Validation should happen in this order:

1. envelope validation
   - required fields present
   - op kind recognized
2. payload-shape validation
   - canonical shape or accepted shorthand only
3. semantic validation
   - required field combinations
   - forbidden field combinations
   - obvious type/range errors
4. normalization
   - only if a shorthand is explicitly supported
   - emit explicit metadata when normalization occurred

### 3) Error contract

Validation errors should be actionable and structured.

Suggested fields:

- `code`
- `message`
- `op_kind`
- `expected_shape`
- `unexpected_fields`
- `missing_fields`
- `example`

CLI should also provide concise prose guidance, but machine-readable fields are required.

## Contract options

### Option A: direct-shape only

For each op kind, require the exact replay payload and reject everything else.

**Pros**
- simplest contract
- lowest ambiguity
- best replay consistency

**Cons**
- less ergonomic for some hand-authored payloads

### Option B: canonical shape + explicit shorthand normalization

Accept one or two narrowly defined shorthands, but normalize into the canonical payload and report that fact.

**Pros**
- friendlier transition path
- can preserve some human ergonomics

**Cons**
- more complexity
- higher drift risk if poorly constrained

## Recommendation

Use **Option B** only when there is a strong, repeated ergonomics case. Otherwise use **Option A**.

Default bias should be strictness.

## Example rollout

1. Add registry/specs for the highest-risk op families first.
2. Make examples/schema output derive from the registry.
3. Add validation errors for known ambiguous shapes.
4. If needed, keep one-release warning mode for legacy shorthand forms.
5. Remove legacy acceptance after the warning window.

## Cross-surface implications

- CLI should expose `schema` / `example` from the same registry.
- MCP can project the same canonical payload rules into tool metadata.
- SDK can ship typed helpers generated from or aligned with the same specs.

## Success criteria

- An agent cannot accidentally stage the wrong shape and think it worked.
- Example output is guaranteed to match accepted input.
- Replay semantics are no longer discoverable only by reading source.
