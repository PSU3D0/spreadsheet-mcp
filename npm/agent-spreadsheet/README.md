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

# Profile column types and cardinality
agent-spreadsheet table-profile data.xlsx --sheet "Sheet1"

# Edit → diff workflow
agent-spreadsheet copy data.xlsx /tmp/draft.xlsx
agent-spreadsheet edit /tmp/draft.xlsx Sheet1 "B2=500" "C2==B2*1.1"
agent-spreadsheet diff data.xlsx /tmp/draft.xlsx
```

All commands output JSON to stdout. Use `--compact` to minimize whitespace.
For CSV, use command-specific options such as `read-table --table-format csv`.

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
