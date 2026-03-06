# Ticket: 4305 Scenario Benchmark Harness + Budget Regression Gates

## Depends On
- Existing scenario fixtures/harnesses
- Existing CLI/MCP/SDK test infrastructure

## Why
The recent benchmark gave high-value feedback, but the project currently lacks a standing pass/fail mechanism for tool-call cost, token cost, and verification quality.

## Owner / Effort / Risk
- Owner (proposed): Benchmarks / CI / Agent UX
- Effort: M
- Risk: Med

## Scope
Turn representative scenarios into measurable regression harnesses.

### Initial benchmark anchor
- `scenario-01-roll-forward`

### Metrics to track
- tool calls
- wall time
- prompt/input size (where measurable)
- output size / token proxy
- number of manual verification loops
- correctness postconditions

### CI / Review goals
- record budgets
- flag regressions
- allow explicit baseline updates when intentional

## Non-Goals
- Perfect token accounting across all vendors/runtime layers.
- One benchmark to rule them all.

## Tests
- Harness is deterministic enough for regression use.
- Budget regressions are visible in CI artifacts or summaries.
- Correctness assertions remain separate from cost assertions.

## Definition of Done
- Scenario ergonomics become measurable and reviewable rather than anecdotal.
