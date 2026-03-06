# Tranche 44: MCP + SDK Ergonomics (Next)

This tranche productizes the stateful surfaces after the immediate contract and verification work is in place.

## Why this tranche exists

Once the contracts are hardened, the next opportunity is to make the stateful surfaces feel task-oriented rather than plumbing-oriented.

This tranche is about:

- MCP workflow helpers
- SDK session/workflow objects
- cross-surface recipe parity
- CLI vnext grouping and workflow-oriented help

## Tickets

1. [4401](./4401-task-oriented-mcp-workflow-helpers-and-recipes.md) — task-oriented MCP helpers and safe recipe layer
2. [4402](./4402-js-sdk-session-workflow-objects-and-verification-helpers.md) — JS SDK workflow objects and verification helpers
3. [4403](./4403-cli-vnext-command-grouping-and-workflow-help-surface.md) — CLI grouping and workflow-oriented help surface
4. [4404](./4404-cross-surface-recipe-parity-and-capability-gating.md) — recipe parity and capability gating across MCP/WASM/SDK

## Design docs

- [Design: JS SDK workflow object model](./designs/4402-sdk-session-workflow-object-model.md)

## Suggested order

- **P1:** 4401, 4402
- **P2:** 4403, 4404

## Acceptance gate

- MCP offers fewer, sharper task helpers for common stateful workflows.
- SDK is useful at the workflow level, not just the transport-normalization level.
- Cross-surface helpers respect the existing boundary rules.
