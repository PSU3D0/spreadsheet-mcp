# spreadsheet-kit-sdk

Backend abstraction for shared spreadsheet-kit JavaScript workflows.

This package provides:

- `McpBackend` wrapper for MCP tool transports
- `WasmBackend` wrapper for in-process WASM bindings
- capability model (`supportsForkLifecycle`, `supportsStaging`, etc.)
- typed-ish SDK errors (`SpreadsheetSdkError`, `CapabilityError`, `BackendOperationError`)

## Install

```bash
npm i spreadsheet-kit-sdk
```

## Basic usage

```js
const { McpBackend, WasmBackend } = require("spreadsheet-kit-sdk")

const mcp = new McpBackend({ transport: myMcpTransport })
const wasm = new WasmBackend({ bindings: myWasmBindings })

async function readSummary(backend, ctx) {
  const sheets = await backend.listSheets(ctx)
  const preview = await backend.rangeValues({
    ...ctx,
    sheetName: sheets[0],
    ranges: "A1:C5"
  })
  return { sheets, preview }
}
```

Callers can switch backend instances without changing shared-method callsites.

## Capability checks

Use `backend.getCapabilities()` to branch on backend-specific flows.

```js
const caps = backend.getCapabilities()
if (caps.supportsForkLifecycle) {
  await backend.createFork({ workbookId: "wb-1" })
}
```

If you call an unsupported method, a `CapabilityError` is thrown with code `UNSUPPORTED_CAPABILITY`.

## Notes

- `McpBackend` and `WasmBackend` intentionally differ for host-specific concerns.
- Shared read/write methods are normalized where semantics match.
- This scaffold is tranche-35 MVP surface and will expand incrementally.
