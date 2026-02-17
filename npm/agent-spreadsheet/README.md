# agent-spreadsheet

[![npm](https://img.shields.io/npm/v/agent-spreadsheet.svg)](https://www.npmjs.com/package/agent-spreadsheet)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](https://github.com/PSU3D0/spreadsheet-mcp/blob/main/LICENSE)

Stateless spreadsheet CLI for AI agents. This npm package downloads a prebuilt native binary for your platform — no Rust toolchain required.

Part of the [spreadsheet-kit](https://github.com/PSU3D0/spreadsheet-mcp) workspace.

## Install

```bash
npm i -g agent-spreadsheet
agent-spreadsheet --help
```

## Quickstart

```bash
# List sheets in a workbook
agent-spreadsheet list-sheets data.xlsx

# Read a table as structured JSON
agent-spreadsheet read-table data.xlsx --sheet "Sheet1"

# Deterministic row paging
agent-spreadsheet sheet-page data.xlsx Sheet1 --format compact --page-size 200

# Profile column types and cardinality
agent-spreadsheet table-profile data.xlsx --sheet "Sheet1"

# Value search
agent-spreadsheet find-value data.xlsx "Revenue" --mode value

# Label lookup: match label cell and return adjacent value
agent-spreadsheet find-value data.xlsx "Net Income" --mode label --label-direction below

# Stateless transform dry-run
agent-spreadsheet transform-batch data.xlsx --ops @ops.json --dry-run

# Edit → diff workflow
agent-spreadsheet copy data.xlsx /tmp/draft.xlsx
agent-spreadsheet edit /tmp/draft.xlsx Sheet1 "B2=500" "C2==B2*1.1"
agent-spreadsheet diff data.xlsx /tmp/draft.xlsx
```

All commands output JSON to stdout.
Use `--shape canonical|compact` (default: `canonical`) to control response shape.

For `range-values`, shape policy is:
- **Canonical (default/omitted): return `values: [...]` when entries are present; omit `values` when all requested ranges are pruned (for example, invalid ranges).**
- **Compact (single range):** flatten that entry to top-level fields (`range`, payload, optional `next_start_row`).
- **Compact (multiple ranges):** keep `values: [...]` with per-entry `range`.

For other high-traffic commands:
- `read-table` and `sheet-page` compact mode preserves active response branches and continuation fields (`next_offset`, `next_start_row`).
- `formula-trace` compact mode omits per-layer highlights but preserves `layers` and `next_cursor`.

Use `--compact` to minimize whitespace.
Global `--output-format csv` is currently unsupported; use command-specific CSV options such as `read-table --table-format csv`.

`apply-formula-pattern` clears cached results for touched formula cells; run `recalculate` to refresh computed values.

### CLI command reference (high-traffic)

| Command | Description |
| --- | --- |
| `read-table <file> [--sheet S] [--range R] [--table-format json\|values\|csv] [--limit N] [--offset N]` | Structured table read with deterministic offset pagination |
| `sheet-page <file> <sheet> --format <full|compact|values_only> [--start-row ROW] [--page-size N]` | Deterministic sheet paging with `next_start_row` continuation |
| `range-values <file> <sheet> <range> [range...]` | Raw values for one or more ranges |
| `find-value <file> <query> [--sheet S] [--mode value\|label] [--label-direction right\|below\|any]` | Search values, or match labels and return adjacent values |
| `transform-batch <file> --ops @ops.json (--dry-run|--in-place|--output PATH)` | Generic stateless batch writes |
| `style-batch <file> --ops @ops.json (--dry-run|--in-place|--output PATH)` | Stateless style operations |
| `apply-formula-pattern <file> --ops @ops.json (--dry-run|--in-place|--output PATH)` | Stateless formula fill/pattern operations (clears touched caches) |
| `structure-batch <file> --ops @ops.json (--dry-run|--in-place|--output PATH)` | Stateless structure operations |
| `rules-batch <file> --ops @ops.json (--dry-run|--in-place|--output PATH)` | Stateless validation/conditional-format operations |

## Platform support

Prebuilt binaries are downloaded on `npm install` for:

| Platform | Architecture | Asset |
| --- | --- | --- |
| Linux | x86_64 | `agent-spreadsheet-linux-x86_64` |
| macOS | x86_64 | `agent-spreadsheet-macos-x86_64` |
| macOS | arm64 (Apple Silicon) | `agent-spreadsheet-macos-aarch64` |
| Windows | x86_64 | `agent-spreadsheet-windows-x86_64.exe` |

Binaries are downloaded from [GitHub Releases](https://github.com/PSU3D0/spreadsheet-mcp/releases) and placed in `vendor/` inside the package directory.

## Environment variables

| Variable | Description |
| --- | --- |
| `AGENT_SPREADSHEET_LOCAL_BINARY` | Path to a local binary to use instead of downloading. Useful for development or air-gapped environments. |
| `AGENT_SPREADSHEET_DOWNLOAD_BASE_URL` | Override the release download host (default: `https://github.com/PSU3D0/spreadsheet-mcp/releases/download`). |

## Troubleshooting

### Binary not found after install

If you see `BINARY_NOT_INSTALLED`, the postinstall download may have failed. Try:

```bash
# Reinstall to re-trigger the download
npm i -g agent-spreadsheet

# Check network access to GitHub releases
curl -I https://github.com/PSU3D0/spreadsheet-mcp/releases/latest
```

### Unsupported platform

The install script supports Linux x64, macOS x64/arm64, and Windows x64. For other platforms, build from source:

```bash
cargo install agent-spreadsheet
```

### Using a local binary

For development or CI where you've already built the binary:

```bash
AGENT_SPREADSHEET_LOCAL_BINARY=./target/release/agent-spreadsheet npm i -g agent-spreadsheet
```

### Corporate proxy / air-gapped

Set `AGENT_SPREADSHEET_DOWNLOAD_BASE_URL` to point to an internal mirror hosting the release assets:

```bash
AGENT_SPREADSHEET_DOWNLOAD_BASE_URL=https://internal-mirror.example.com/releases npm i -g agent-spreadsheet
```

## How it works

1. `npm install` runs `scripts/install.js` as a postinstall hook
2. The script detects your platform and architecture
3. Downloads the matching binary from GitHub Releases (or copies from `AGENT_SPREADSHEET_LOCAL_BINARY`)
4. Places it in `vendor/` and marks it executable
5. `bin/agent-spreadsheet.js` spawns the vendored binary with your arguments

## Full documentation

For the complete command reference, MCP server setup, Docker deployment, and architecture docs, see the [root README](https://github.com/PSU3D0/spreadsheet-mcp#readme).

## License

Apache-2.0
