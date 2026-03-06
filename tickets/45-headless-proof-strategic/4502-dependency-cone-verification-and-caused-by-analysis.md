# Ticket: 4502 Dependency-Cone Verification + Caused-By Analysis

## Depends On
- 4301
- 4501
- [Design: dependency-cone verification engine](./designs/4502-dependency-cone-verification-engine.md)

## Why
A key long-term headless advantage is not merely showing what changed, but showing what downstream effects were caused by those changes.

## Owner / Effort / Risk
- Owner (proposed): Verification / Analysis
- Effort: XL
- Risk: High

## Scope
Build dependency-aware verification that can explain downstream impact.

### Candidate outputs
- dependency-cone change summaries
- newly broken cells in downstream cones
- direct edit vs propagated impact distinction
- target-output impact proofs

## Non-Goals
- Perfect spreadsheet theorem proving.
- Immediate parity for every unsupported formula edge case.

## Tests
- Dependency-scoped impacts are deterministic on representative fixtures.
- Direct vs propagated changes remain distinguishable.
- Error provenance can point to likely upstream causes.

## Definition of Done
- The system can answer "what did my change cause?" in a structured, machine-usable way.
