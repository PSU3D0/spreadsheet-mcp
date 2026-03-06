# Design: 4402 JS SDK Session / Workflow Object Model

## Problem

The SDK currently provides useful transport normalization, but callers still operate mostly at the raw method level. That leaves too much workflow assembly burden on the application or agent.

## Goals

- Add a workflow-oriented SDK layer without creating a fourth semantics fork.
- Make common stateful flows easier to express.
- Keep backend capability differences explicit.

## Non-Goals

- Replacing backend-specific advanced flows entirely.
- Forcing CLI/path concepts into the SDK.

## Proposed object model

### Base backend layer
Keep the existing backend abstraction for shared primitive calls.

### Workflow layer on top
Introduce higher-level objects such as:

- `WorkbookHandle`
- `SessionHandle`
- optional `VerificationResult` helpers

Example direction:

```js
const wb = await sdk.openWorkbook(ctx)
const session = await wb.createSession()

await session.transformBatch(...)
await session.recalculate(...)
const proof = await session.verifyTargets(...)
await session.exportWorkbook(...)
```

## Capability model

Workflow objects must expose or inherit capability metadata.

Examples:
- MCP may support richer staged/fork flows.
- WASM may support local session/export flows.
- shared helpers should exist only where semantics are truly shared.

## Error model

Workflow helpers should preserve the SDK error contract:
- capability errors stay explicit
- backend failures remain attributable
- validation failures remain structured

## Verification helpers

The workflow layer is the natural home for helpers like:
- `verifyTargets`
- `collectNewErrors`
- `summarizeChanges`
- `compareNamedRanges`

These should wrap shared semantics, not invent new ones.

## Success criteria

- Applications can express common workbook flows more directly.
- Backend differences remain explicit and capability-gated.
- The SDK becomes a real product surface, not just a transport adapter.
