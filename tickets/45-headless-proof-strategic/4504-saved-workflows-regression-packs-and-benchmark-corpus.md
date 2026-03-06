# Ticket: 4504 Saved Workflows, Regression Packs, and Benchmark Corpus

## Depends On
- 4305
- 4503

## Why
Once workflows and assertions exist, the project needs a durable corpus of scenarios and regression packs to improve the product deliberately.

## Owner / Effort / Risk
- Owner (proposed): Benchmarks / Product Engineering
- Effort: L
- Risk: Med

## Scope
Build a reusable benchmark/regression corpus for spreadsheet-agent workflows.

### Candidate components
- saved workflow definitions
- reusable assertion bundles
- scenario fixture metadata
- budget baselines and historical comparisons

## Non-Goals
- One corpus that covers every spreadsheet domain.
- Public inclusion of sensitive/private scenario material.

## Tests
- Workflow packs are replayable.
- Regression assertions are stable.
- Budget/correctness drift is visible over time.

## Definition of Done
- The project can track product progress against a durable, reusable workflow corpus.
