# Cross-Surface Parity Harness (Tranche 35)

This harness is the tranche-35 entrypoint for fixture parity checks across core/wasm/mcp/sdk layers.

## Entrypoint

```bash
node scripts/parity-harness.js --list
node scripts/parity-harness.js
```

- `--list` prints the full fixture/check matrix and status.
- default run executes the currently automated subset.

## Fixture matrix

| Check id | Surfaces | Status | Notes |
| --- | --- | --- | --- |
| `basic-read-parity` | core, wasm, mcp, sdk | automated | Shared read fixture via SDK wrappers (`listSheets`, `rangeValues`) |
| `capability-divergence-guards` | wasm, mcp, sdk | automated | Capability misuse must return typed SDK capability errors |
| `formula-diagnostics-parity` | core, wasm, mcp | planned | Shared fixtures for formula parse policy (`off\|warn\|fail`) |
| `grid-roundtrip-parity` | core, wasm, mcp | planned | Grid export/import parity for values/formulas/styles/merges/column widths |
| `csv-edge-case-parity` | core, wasm, mcp | planned | CSV edge cases (quotes, CRLF, embedded newline, blanks) |

## Automated subset (current)

Automated checks intentionally focus on tranche-35 SDK backend abstraction behavior:

- backend switching for shared methods (`McpBackend` vs `WasmBackend`)
- typed capability misuse errors for unsupported backend-only methods

Planned checks are tracked but not yet wired into CI by this script.
