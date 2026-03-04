# spreadsheet-kit-sdk

Backend-agnostic JavaScript SDK for `spreadsheet-kit` surfaces.

It lets you write one app-level integration and switch between:

- **MCP backend** (`McpBackend`) for server/remote workflows
- **WASM backend** (`WasmBackend`) for in-process session workflows

---

## Install

```bash
npm i spreadsheet-kit-sdk
```

Node: `>=18`

---

## What this SDK normalizes

- Shared method names (`describeWorkbook`, `sheetOverview`, `rangeValues`, etc.)
- Input aliases (`camelCase` + `snake_case` tolerance)
- Output shape normalization (SDK returns `camelCase` top-level fields)
- Typed errors for capability and backend failures

---

## Quick start

```js
const { McpBackend, WasmBackend } = require("spreadsheet-kit-sdk")

// 1) MCP backend
const mcp = new McpBackend({ transport: myMcpTransport })

// 2) WASM backend
const wasm = new WasmBackend({ bindings: myWasmBindings })

async function sharedReadFlow(backend, ctx) {
  const desc = await backend.describeWorkbook(ctx)
  const sheets = await backend.listSheets(ctx)
  const page = await backend.sheetPage({
    ...ctx,
    sheetName: sheets[0],
    startRow: 1,
    pageSize: 50,
    format: "compact"
  })
  return { desc, sheets, page }
}
```

`ctx` typically includes one of:

- MCP: `{ workbookId: "..." }`
- WASM: `{ sessionId: "..." }`
- Cross-surface helper: `{ contextId: "..." }`

---

## Shared methods (currently parity-wired)

### Read/analysis

- `describeWorkbook(input)`
- `namedRanges(input)`
- `sheetOverview(input)`
- `listSheets(input)`
- `rangeValues(input)`
- `findValue(input)`
- `readTable(input)`
- `sheetPage(input)`
- `gridExport(input)`

### Write

- `transformBatch(input)`

### Lifecycle / backend-specific

- MCP-only fork lifecycle: `createFork`, `listForks`, `saveFork`, `discardFork`, staged-change methods
- WASM-only session lifecycle: `createSession`, `exportWorkbook`, `disposeSession`

Always branch using capabilities before backend-specific methods.

---

## Capability model

```js
const caps = backend.getCapabilities()

if (caps.supportsForkLifecycle) {
  // MCP path
}
if (caps.supportsSessionLifecycle) {
  // WASM path
}
```

If you call an unsupported method, SDK throws `CapabilityError` with code:

- `UNSUPPORTED_CAPABILITY`

---

## Error model

SDK exports:

- `SpreadsheetSdkError`
- `CapabilityError`
- `BackendOperationError`

Typical codes:

- `INVALID_ARGUMENT`
- `INVALID_RESPONSE`
- `UNSUPPORTED_CAPABILITY`
- backend-provided error envelopes are normalized when possible

---

## MCP transport contract

`McpBackend` accepts either:

- an `invoke(operation, params)` function
- or per-operation methods (`transport.list_sheets`, etc.)

Example:

```js
const backend = new McpBackend({
  transport: {
    async invoke(operation, params) {
      // call your MCP client
      return client.callTool(operation, params)
    }
  }
})
```

---

## WASM bindings contract

`WasmBackend` expects function bindings named like:

- `createSession`, `describeWorkbook`, `namedRanges`, `sheetOverview`
- `listSheets`, `rangeValues`, `findValue`, `readTable`, `sheetPage`, `gridExport`
- `transformBatch`, `exportWorkbook`, `disposeSession`

These map directly to methods exposed by `spreadsheet-kit-wasm`.

---

## Publishing channels / npm dist-tags

Release lanes:

- `sdk-vX.Y.Z` tags publish this package

Dist-tag policy:

- stable `X.Y.Z` -> `latest`
- prerelease `X.Y.Z-rc.N` -> `rc`
- prerelease `X.Y.Z-beta.N` -> `beta`
- prerelease `X.Y.Z-alpha.N` -> `alpha`

So prereleases never override `latest` unintentionally.

---

## Related packages

- `agent-spreadsheet` (npm CLI wrapper)
- `spreadsheet-kit` (Rust core crate)
- `spreadsheet-kit-wasm` (Rust/WASM adapter crate)

Repo: <https://github.com/PSU3D0/spreadsheet-mcp>
