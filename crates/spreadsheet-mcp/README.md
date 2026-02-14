# spreadsheet-mcp

[![Crates.io](https://img.shields.io/crates/v/spreadsheet-mcp.svg)](https://crates.io/crates/spreadsheet-mcp)
[![License](https://img.shields.io/crates/l/spreadsheet-mcp.svg)](https://github.com/PSU3D0/spreadsheet-mcp/blob/main/LICENSE)

Stateful MCP server for spreadsheet analysis and editing. Part of the [spreadsheet-kit](https://github.com/PSU3D0/spreadsheet-mcp) workspace.

## Install

```bash
cargo install spreadsheet-mcp
```

Formualizer (native Rust recalc engine) is included by default.

Or use Docker:

```bash
# Read-only (~15 MB)
docker pull ghcr.io/psu3d0/spreadsheet-mcp:latest

# With write/recalc (~800 MB, includes LibreOffice)
docker pull ghcr.io/psu3d0/spreadsheet-mcp:full
```

## Quickstart

### Stdio (MCP clients)

```bash
spreadsheet-mcp --workspace-root /path/to/workbooks --transport stdio
```

Add to Claude Code / Claude Desktop (`~/.claude.json` or `.mcp.json`):

```json
{
  "mcpServers": {
    "spreadsheet": {
      "command": "spreadsheet-mcp",
      "args": ["--workspace-root", "/path/to/workbooks", "--transport", "stdio"]
    }
  }
}
```

### HTTP

```bash
spreadsheet-mcp --workspace-root /path/to/workbooks
# → http://127.0.0.1:8079 — POST /mcp
```

### Docker (with write/recalc)

```bash
docker run -v /path/to/workbooks:/data -p 8079:8079 \
  ghcr.io/psu3d0/spreadsheet-mcp:full
```

## What this crate provides

- **50+ MCP tools** — read-only analysis, VBA inspection, fork-based editing, recalculation, screenshots
- **LRU workbook cache** — configurable capacity, lazy sheet metrics
- **Region detection** — automatic identification of tables, parameters, and output blocks
- **Token-dense output** — sparse JSON by default, pagination only when needed
- **Fork lifecycle** — create, edit, checkpoint, recalculate, diff, save
- **Recalc backends** — Formualizer (default) or LibreOffice (feature: `recalc-libreoffice`, Docker `:full`)

## Binaries

| Binary | Purpose |
| --- | --- |
| `spreadsheet-mcp` | MCP server (primary) |

## Configuration

Key flags (all have `SPREADSHEET_MCP_` env equivalents):

```
--workspace-root <DIR>       Workspace directory (default: cwd)
--transport stdio|http       Transport mode (default: http)
--cache-capacity <N>         LRU cache size (default: 5)
--recalc-enabled             Enable write/recalc tools
--vba-enabled                Enable VBA inspection tools
--output-profile <P>         token-dense (default) or verbose
```

See the [full configuration reference](https://github.com/PSU3D0/spreadsheet-mcp#configuration-reference) in the root README.

## Related crates

| Crate | Role |
| --- | --- |
| [`spreadsheet-kit`](../spreadsheet-kit/) | Shared core primitives |

## Full documentation

See the [root README](https://github.com/PSU3D0/spreadsheet-mcp#readme) for the complete tool surface, deployment guide, Docker configuration, write tool shapes, and agent workflow recommendations.

## License

Apache-2.0 — see [LICENSE](../../LICENSE).
