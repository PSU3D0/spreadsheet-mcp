# agent-spreadsheet

[![Crates.io](https://img.shields.io/crates/v/agent-spreadsheet.svg)](https://crates.io/crates/agent-spreadsheet)
[![npm](https://img.shields.io/npm/v/agent-spreadsheet.svg)](https://www.npmjs.com/package/agent-spreadsheet)
[![License](https://img.shields.io/crates/l/agent-spreadsheet.svg)](https://github.com/PSU3D0/spreadsheet-mcp/blob/main/LICENSE)

Stateless spreadsheet CLI for AI agents. Part of the [spreadsheet-kit](https://github.com/PSU3D0/spreadsheet-mcp) workspace.

Every command takes a file path and returns JSON. No server, no state, no setup.

## Install

```bash
# npm (recommended — downloads prebuilt binary, no Rust needed)
npm i -g agent-spreadsheet

# Cargo
cargo install agent-spreadsheet
```

Or download from [GitHub Releases](https://github.com/PSU3D0/spreadsheet-mcp/releases).

## Quickstart

```bash
# Explore
agent-spreadsheet list-sheets data.xlsx
agent-spreadsheet describe data.xlsx
agent-spreadsheet sheet-overview data.xlsx "Sheet1"

# Read
agent-spreadsheet read-table data.xlsx --sheet "Sheet1"
agent-spreadsheet range-values data.xlsx "Sheet1" "A1:D10"
agent-spreadsheet find-value data.xlsx "Revenue" --mode label
agent-spreadsheet table-profile data.xlsx --sheet "Sheet1"

# Analyze formulas
agent-spreadsheet formula-map data.xlsx "Sheet1"
agent-spreadsheet formula-trace data.xlsx "Sheet1" "C10" precedents

# Edit → recalculate → diff
agent-spreadsheet copy data.xlsx /tmp/draft.xlsx
agent-spreadsheet edit /tmp/draft.xlsx Sheet1 "B2=500" "C2==B2*1.1"
agent-spreadsheet recalculate /tmp/draft.xlsx
agent-spreadsheet diff data.xlsx /tmp/draft.xlsx
```

## Command families

### Discovery

| Command | Description |
| --- | --- |
| `list-sheets <file>` | List sheets with summaries |
| `describe <file>` | Workbook metadata (sheets, named ranges, format) |
| `sheet-overview <file> <sheet>` | Region detection + orientation |

### Reading

| Command | Description |
| --- | --- |
| `read-table <file> [--sheet S] [--range R]` | Structured table read (`--table-format json\|values\|csv`) |
| `range-values <file> <sheet> [ranges...]` | Raw cell values |
| `table-profile <file> [--sheet S]` | Column types, cardinality, distributions |
| `find-value <file> <query> [--sheet S] [--mode M]` | Search values (`value`) or labels (`label`) |

### Formula analysis

| Command | Description |
| --- | --- |
| `formula-map <file> <sheet>` | Formula inventory (`--sort-by complexity\|count`) |
| `formula-trace <file> <sheet> <cell> <dir>` | Trace `precedents` or `dependents` |

### Writing

| Command | Description |
| --- | --- |
| `copy <source> <dest>` | Copy workbook for editing |
| `edit <file> <sheet> <edits...>` | Apply edits: `A1=42` (value), `B2==SUM(...)` (formula) |
| `recalculate <file>` | Recalculate formulas (Formualizer, included by default) |
| `diff <original> <modified>` | Diff two workbook versions |

## Output

All commands emit JSON to stdout. Global flags:

| Flag | Effect |
| --- | --- |
| `--format json` | JSON output (default) |
| `--compact` | Minimize JSON whitespace |
| `--quiet` | Suppress warnings |

Global `--format csv` is reserved and currently not implemented.
Use command-specific CSV options such as `read-table --table-format csv`.

Errors go to stderr as structured JSON with `code` and `message` fields.

## Agent workflow example

A typical LLM agent loop using the CLI:

```
1. agent-spreadsheet list-sheets budget.xlsx       → pick sheet
2. agent-spreadsheet sheet-overview budget.xlsx Q1  → find regions
3. agent-spreadsheet read-table budget.xlsx --sheet Q1  → read data
4. agent-spreadsheet copy budget.xlsx /tmp/edit.xlsx
5. agent-spreadsheet edit /tmp/edit.xlsx Q1 "B5=1200" "B6==B5*0.15"
6. agent-spreadsheet diff budget.xlsx /tmp/edit.xlsx    → verify changes
```

## Related

| Package | Role |
| --- | --- |
| [`spreadsheet-mcp`](../spreadsheet-mcp/) | Stateful MCP server (for multi-turn agent sessions) |
| [`spreadsheet-kit`](../spreadsheet-kit/) | Shared core primitives |
| [`agent-spreadsheet` (npm)](../../npm/agent-spreadsheet/) | npm wrapper for binary distribution |

For the full tool surface, recalc backends, deployment options, and architecture docs, see the [root README](https://github.com/PSU3D0/spreadsheet-mcp#readme).

## License

Apache-2.0 — see [LICENSE](../../LICENSE).
