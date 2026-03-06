# Tranche 43: Verification + Workflow Helpers (Now)

This tranche focuses on the next layer of near-term competitiveness: reducing token burn for moderate spreadsheet tasks and making post-edit proof first-class.

## Why this tranche exists

The benchmark surfaced a core truth: the architecture is promising, but the agent still spends too much effort proving that a change was safe and correct.

This tranche turns verification into a product feature and adds generic workflow helpers for common edit patterns.

## Tickets

1. [4301](./4301-post-edit-verification-target-delta-and-new-error-provenance.md) — target delta + new error provenance verification surface
2. [4302](./4302-grouped-diff-and-change-summary.md) — grouped diffs and lower-noise change summaries
3. [4303](./4303-region-append-csv-append-and-footer-aware-insert-primitives.md) — append/insert helpers for detected regions
4. [4304](./4304-template-row-and-row-band-clone-helpers.md) — template row / row-band clone helpers and formula extension
5. [4305](./4305-scenario-benchmark-harness-and-budget-regression-gates.md) — scenario benchmark harness with token/time/tool-call budgets

## Design docs

- [Design: post-edit verification + provenance](./designs/4301-post-edit-verification-and-provenance.md)
- [Design: region append + footer-aware insert helpers](./designs/4303-region-append-and-footer-aware-insert.md)

## Suggested order

- **P0:** 4301, 4305
- **P1:** 4302, 4303
- **P2:** 4304

## Acceptance gate

- A moderate scenario can prove target outcomes and isolate newly introduced problems.
- Change summaries are grouped and interpretable.
- Common region/append/footer workflows require fewer primitive operations.
- Scenario regressions are measured, not anecdotal.
