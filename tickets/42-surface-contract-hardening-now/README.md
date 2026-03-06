# Tranche 42: Surface Contract Hardening (Now)

This tranche focuses on the highest-value near-term hardening work for `asp`, CLI, MCP, and SDK agent ergonomics.

## Why this tranche exists

Recent benchmark feedback showed that the biggest failures were not core spreadsheet logic failures. They were surface-contract failures:

- command/binary discovery was expensive
- mutation payload shapes were not obvious enough
- examples/docs were not canonical enough
- empty results and warnings were too implicit

This tranche is intended to make the substrate reliable and obvious before we add more power.

## Tickets

1. [4201](./4201-asp-command-obviousness-install-and-help-surface.md) — `asp` command obviousness, install, aliasing, and help alignment
2. [4202](./4202-session-payload-validation-and-canonical-mutation-envelopes.md) — strict session payload validation + canonical envelopes
3. [4203](./4203-schema-example-discoverability-and-self-describing-surfaces.md) — schema/example discovery across CLI, MCP, and SDK
4. [4204](./4204-explicit-empty-result-warning-and-count-contracts.md) — explicit empty results, warnings, and count semantics
5. [4205](./4205-thin-canonical-docs-skills-and-agent-workflows.md) — thinner docs/skills with one canonical workflow per op family

## Design docs

- [Design: canonical session op contract](./designs/4202-session-op-canonical-contract.md)

## Suggested order

- **P0:** 4201, 4202, 4204
- **P1:** 4203, 4205

## Acceptance gate

- An agent can discover and invoke the right binary/entrypoint immediately.
- Session mutations reject malformed or ambiguous payloads with actionable errors.
- Empty results are explicit, not inferred.
- Docs and skills match actual replay/runtime semantics.
