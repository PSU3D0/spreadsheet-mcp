# spreadsheet-kit

[![Crates.io](https://img.shields.io/crates/v/spreadsheet-kit.svg)](https://crates.io/crates/spreadsheet-kit)
[![License](https://img.shields.io/crates/l/spreadsheet-kit.svg)](https://github.com/PSU3D0/spreadsheet-mcp/blob/main/LICENSE)

Core shared primitives for the [spreadsheet-kit](https://github.com/PSU3D0/spreadsheet-mcp) workspace.

## What's in this crate

- **Shared engine modules** — workbook model, repository layer, tools, diff/recalc pipelines
- **`agent-spreadsheet` CLI binary** — stateless command surface for reads/edits/diff/recalc
- **`CellEdit`** — canonical cell edit type (address + value + is_formula)
- **`CoreWarning`** — structured warning type (code + message)
- **`BasicDiffChange` / `BasicDiffResponse`** — diff result types
- **`RecalculateOutcome`** — recalc result metadata (backend, duration, cells evaluated, errors)
- **Edit normalization** — `normalize_shorthand_edit()` parses `"A1=100"` / `"B2==SUM(...)"` into `CellEdit`
- **`apply_edits_to_file()`** — applies a batch of `CellEdit`s to an `.xlsx` file via `umya-spreadsheet`
- **`SessionRuntime` trait** — scaffold for stateful session backends (open / apply_edits / recalculate / save_as)

## What's intentionally not here

- **MCP server logic** — see [`spreadsheet-mcp`](../spreadsheet-mcp/)
- **MCP transport adapter** — lives in `spreadsheet-mcp`

This crate is kept minimal so both the MCP server and the CLI can depend on it without pulling in server-specific dependencies.

## Usage

```toml
[dependencies]
spreadsheet-kit = "0.1"
```

```rust
use spreadsheet_kit::write::normalize_shorthand_edit;

let edit = normalize_shorthand_edit("A1=42").unwrap();
assert_eq!(edit.address, "A1");
assert_eq!(edit.value, "42");
assert!(!edit.is_formula);

let formula = normalize_shorthand_edit("B2==SUM(A1:A2)").unwrap();
assert_eq!(formula.address, "B2");
assert_eq!(formula.value, "SUM(A1:A2)");
assert!(formula.is_formula);
```

## Consumers

| Crate | Role |
| --- | --- |
| [`spreadsheet-mcp`](../spreadsheet-mcp/) | Stateful MCP server adapter |
| [`spreadsheet-kit`](./) | Shared engine + `agent-spreadsheet` CLI |

## License

Apache-2.0 — see [LICENSE](../../LICENSE).
