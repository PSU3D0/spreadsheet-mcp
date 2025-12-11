# Spreadsheet MCP

[![Crates.io](https://img.shields.io/crates/v/spreadsheet-mcp.svg)](https://crates.io/crates/spreadsheet-mcp)
[![Documentation](https://docs.rs/spreadsheet-mcp/badge.svg)](https://docs.rs/spreadsheet-mcp)
[![License](https://img.shields.io/crates/l/spreadsheet-mcp.svg)](https://github.com/PSU3D0/spreadsheet-mcp/blob/main/LICENSE)

![Spreadsheet MCP](https://raw.githubusercontent.com/PSU3D0/spreadsheet-mcp/main/assets/banner.jpeg)

MCP server for spreadsheet analysis and editing. Slim, token-efficient tool surface designed for LLM agents.

## Why?

Dumping a 50,000-row spreadsheet into an LLM context is expensive and usually unnecessary. Most spreadsheet tasks need surgical access: find a region, profile its structure, read a filtered slice. This server exposes tools that let agents **discover → profile → extract** without burning tokens on cells they don't need.

- **Full support:** `.xlsx` (via `umya-spreadsheet`)
- **Discovery only:** `.xls`, `.xlsb` (enumerated, not parsed)

## Architecture

![Architecture Overview](https://raw.githubusercontent.com/PSU3D0/spreadsheet-mcp/main/assets/architecture_overview.jpeg)

- **LRU cache** keeps recently-accessed workbooks in memory (configurable capacity)
- **Lazy sheet metrics** computed once per sheet, reused across tools
- **Region detection** runs once and caches bounds for `sheet_overview`, `find_value`, `read_table`, `table_profile`

## Tool Surface

| Tool | Purpose |
| --- | --- |
| `list_workbooks`, `list_sheets` | Discover targets and get sheet summaries |
| `workbook_summary` | Region counts/kinds, named ranges, suggested entry points |
| `sheet_overview` | Detected regions (bounds/id/kind/confidence), narrative, key ranges |
| `find_value` | Value/label-mode lookup with region/table scoping, neighbors, row context |
| `read_table` | Structured read (range/region/table/named range), headers, filters, sampling |
| `table_profile` | Lightweight column profiling with samples and inferred types |
| `range_values` | Minimal range fetch for spot checks |
| `sheet_page` | Fallback pagination; supports `compact`/`values_only` |
| `formula_trace` | Precedents/dependents with paging |

## Write & Recalc Support

Write tools allow "what-if" analysis: fork a workbook, edit cells, recalculate formulas via LibreOffice, and diff the results.

### Enabling Write Tools

Use the `:full` Docker image (write tools enabled by default):
```bash
docker run -v /path/to/workbooks:/data -p 8079:8079 ghcr.io/psu3d0/spreadsheet-mcp:full
```

Or with a binary build:
```bash
cargo install spreadsheet-mcp --features recalc
spreadsheet-mcp --workspace-root /path/to/workbooks --recalc-enabled
```

**Note:** Recalculation requires LibreOffice installed on the host when not using the `:full` Docker image.

### Write Tools

| Tool | Purpose |
| --- | --- |
| `create_fork` | Create a temporary editable copy for "what-if" analysis |
| `edit_batch` | Apply values or formulas to cells in a fork |
| `recalculate` | Trigger LibreOffice to update formula results |
| `get_changeset` | Diff the fork against the original (cells, tables, named ranges) |
| `save_fork` | Commit changes to a file (overwrite or new path) |
| `discard_fork` | Delete the temporary fork |

See [docs/RECALC.md](docs/RECALC.md) for architecture details.

## Example

**Request:** Profile a detected region
```json
{
  "tool": "table_profile",
  "arguments": {
    "workbook_id": "budget-2024-a1b2c3",
    "sheet_name": "Q1 Actuals",
    "region_id": 1,
    "sample_size": 10,
    "sample_mode": "distributed"
  }
}
```

**Response:**
```json
{
  "sheet_name": "Q1 Actuals",
  "headers": ["Date", "Category", "Amount", "Notes"],
  "column_types": [
    {"name": "Date", "inferred_type": "date", "nulls": 0, "distinct": 87},
    {"name": "Category", "inferred_type": "text", "nulls": 2, "distinct": 12, "top_values": ["Payroll", "Marketing", "Infrastructure"]},
    {"name": "Amount", "inferred_type": "number", "nulls": 0, "min": 150.0, "max": 84500.0, "mean": 12847.32},
    {"name": "Notes", "inferred_type": "text", "nulls": 45, "distinct": 38}
  ],
  "row_count": 1247,
  "samples": [...]
}
```

The agent now knows column types, cardinality, and value distributions—without reading 1,247 rows.

## Recommended Agent Workflow

![Token Efficiency Workflow](https://raw.githubusercontent.com/PSU3D0/spreadsheet-mcp/main/assets/token_efficiency.jpeg)

1. `list_workbooks` → `list_sheets` → `workbook_summary` for orientation
2. `sheet_overview` to get `detected_regions` (ids/bounds/kind/confidence)
3. `table_profile` → `read_table` with `region_id`, small `limit`, and `sample_mode` (`distributed` preferred)
4. Use `find_value` (label mode) or `range_values` for targeted pulls
5. Reserve `sheet_page` for unknown layouts or calculator inspection; prefer `compact`/`values_only`
6. Keep payloads small; page/filter rather than full-sheet reads

## Region Detection

![Region Detection Visualization](https://raw.githubusercontent.com/PSU3D0/spreadsheet-mcp/main/assets/region_detection_viz.jpeg)

Spreadsheets often contain multiple logical tables, parameter blocks, and output areas on a single sheet. The server detects these automatically:

1. **Gutter detection** — Scans for empty rows/columns that separate content blocks
2. **Recursive splitting** — Subdivides large areas along detected gutters
3. **Border trimming** — Removes sparse edges to tighten bounds
4. **Header detection** — Identifies header rows (including multi-row merged headers)
5. **Classification** — Labels each region: `data`, `parameters`, `outputs`, `calculator`, `metadata`
6. **Confidence scoring** — Higher scores for well-structured regions with clear headers

Regions are cached per sheet. Tools like `read_table` accept a `region_id` to scope reads without manually specifying ranges.

## Quick Start

### Docker (Recommended)

Two image variants are published:

| Image | Size | Write/Recalc |
| --- | --- | --- |
| `ghcr.io/psu3d0/spreadsheet-mcp:latest` | ~15MB | No |
| `ghcr.io/psu3d0/spreadsheet-mcp:full` | ~800MB | Yes (includes LibreOffice) |

```bash
# Read-only (slim image)
docker run -v /path/to/workbooks:/data -p 8079:8079 ghcr.io/psu3d0/spreadsheet-mcp:latest

# With write/recalc support (full image)
docker run -v /path/to/workbooks:/data -p 8079:8079 ghcr.io/psu3d0/spreadsheet-mcp:full
```

### Cargo Install

```bash
# Read-only
cargo install spreadsheet-mcp
spreadsheet-mcp --workspace-root /path/to/workbooks

# With write/recalc support
cargo install spreadsheet-mcp --features recalc
spreadsheet-mcp --workspace-root /path/to/workbooks --recalc-enabled
```

**Note:** The `--recalc-enabled` flag requires LibreOffice installed on the host.

### Build from Source

```bash
cargo run --release -- --workspace-root /path/to/workbooks
```

Default transport: HTTP streaming at `127.0.0.1:8079`. Endpoint: `POST /mcp`.

Use `--transport stdio` for CLI pipelines.

## MCP Client Configuration

### Claude Code / Claude Desktop

Add to `~/.claude.json` or project `.mcp.json`:

**Read-only (slim image):**
```json
{
  "mcpServers": {
    "spreadsheet": {
      "command": "docker",
      "args": ["run", "-i", "--rm", "-v", "/path/to/workbooks:/data", "ghcr.io/psu3d0/spreadsheet-mcp:latest", "--transport", "stdio"]
    }
  }
}
```

**With write/recalc (full image):**
```json
{
  "mcpServers": {
    "spreadsheet": {
      "command": "docker",
      "args": ["run", "-i", "--rm", "-v", "/path/to/workbooks:/data", "ghcr.io/psu3d0/spreadsheet-mcp:full", "--transport", "stdio"]
    }
  }
}
```

**Binary (no Docker):**
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

### Cursor / VS Code

**Read-only (slim image):**
```json
{
  "mcp.servers": {
    "spreadsheet": {
      "command": "docker",
      "args": ["run", "-i", "--rm", "-v", "${workspaceFolder}:/data", "ghcr.io/psu3d0/spreadsheet-mcp:latest", "--transport", "stdio"]
    }
  }
}
```

**With write/recalc (full image):**
```json
{
  "mcp.servers": {
    "spreadsheet": {
      "command": "docker",
      "args": ["run", "-i", "--rm", "-v", "${workspaceFolder}:/data", "ghcr.io/psu3d0/spreadsheet-mcp:full", "--transport", "stdio"]
    }
  }
}
```

**Binary (no Docker):**
```json
{
  "mcp.servers": {
    "spreadsheet": {
      "command": "spreadsheet-mcp",
      "args": ["--workspace-root", "${workspaceFolder}", "--transport", "stdio"]
    }
  }
}
```

### HTTP Mode

```bash
docker run -v /path/to/workbooks:/data -p 8079:8079 ghcr.io/psu3d0/spreadsheet-mcp:latest
```

Connect via `POST http://localhost:8079/mcp`.

## Local Development

To test local changes without rebuilding Docker:

```bash
cargo build --release
```

Then point your MCP client to the binary:
```json
{
  "mcpServers": {
    "spreadsheet": {
      "command": "/path/to/spreadsheet-mcp/target/release/spreadsheet-mcp",
      "args": ["--workspace-root", "/path/to/workbooks", "--transport", "stdio"]
    }
  }
}
```

## Configuration

| Flag | Env | Description |
| --- | --- | --- |
| `--workspace-root <DIR>` | `SPREADSHEET_MCP_WORKSPACE` | Workspace root to scan (default: cwd) |
| `--cache-capacity <N>` | `SPREADSHEET_MCP_CACHE_CAPACITY` | Workbook cache size (default: 5) |
| `--extensions <list>` | `SPREADSHEET_MCP_EXTENSIONS` | Allowed extensions (default: `xlsx,xls,xlsb`) |
| `--workbook <FILE>` | `SPREADSHEET_MCP_WORKBOOK` | Single-workbook mode |
| `--enabled-tools <list>` | `SPREADSHEET_MCP_ENABLED_TOOLS` | Whitelist exposed tools |
| `--transport <http\|stdio>` | `SPREADSHEET_MCP_TRANSPORT` | Transport selection (default: http) |
| `--http-bind <ADDR>` | `SPREADSHEET_MCP_HTTP_BIND` | Bind address (default: `127.0.0.1:8079`) |
| `--recalc-enabled` | `SPREADSHEET_MCP_RECALC_ENABLED` | Enable write/recalc tools (default: false) |
| `--max-concurrent-recalcs <N>` | `SPREADSHEET_MCP_MAX_CONCURRENT_RECALCS` | Parallel recalc limit (default: 2) |

## Performance

- **LRU workbook cache** — Recently opened workbooks stay in memory; oldest evicted when capacity exceeded
- **Lazy metrics** — Sheet metrics computed on first access, cached for subsequent calls
- **Region caching** — Detection runs once per sheet; `region_id` lookups are O(1)
- **Sampling modes** — `distributed` sampling reads evenly across rows without loading everything
- **Compact formats** — `values_only` and `compact` output modes reduce response size

## Testing

```bash
cargo test
```

Covers: region detection, region-scoped tools, `read_table` edge cases (merged headers, filters, large sheets), workbook summary.

## Behavior & Limits

- **Read-only by default**; write/recalc features require `--recalc-enabled` or the `:full` image
- **XLSX supported for write**; `.xls`/`.xlsb` are read-only
- Bounded in-memory cache honors `cache_capacity`
- Prefer region-scoped reads and sampling for token/latency efficiency
