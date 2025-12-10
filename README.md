# Spreadsheet Read MCP

# Spreadsheet Read MCP
![Spreadsheet Read MCP](assets/banner.jpeg)

Read-only MCP server for spreadsheet analysis with a slim, token-efficient tool surface. Emphasis: find the right region fast, profile lightly, then read only what you need. XLSX-first (via `umya-spreadsheet`), enumerates `.xls`/`.xlsb` files.

## Tool Surface (≤10 exposed)
| Tool | Purpose |
| --- | --- |
| `list_workbooks`, `list_sheets` | Discover targets and get sheet summaries. |
| `workbook_summary` | Region counts/kinds, named ranges, suggested entry points. |
| `sheet_overview` | Detected regions (bounds/id/kind/confidence), narrative, key ranges. |
| `find_value` | Value/label-mode lookup with region/table scoping, neighbors, row context. |
| `read_table` | Structured read (range/region/table/named range), headers, filters, sampling. |
| `table_profile` | Lightweight column profiling with samples and inferred types. |
| `range_values` | Minimal range fetch for spot checks. |
| `sheet_page` | Fallback pagination; supports `compact`/`values_only`. |
| `formula_trace` | Precedents/dependents with paging. |

Hidden/optional: styles, statistics, volatiles, manifest stub, close_workbook.

## Recommended Agent Workflow
1) `list_workbooks` → `list_sheets` → `workbook_summary` for orientation.  
2) `sheet_overview` to get `detected_regions` (ids/bounds/kind/confidence).  
3) `table_profile` → `read_table` with `region_id`, small `limit`, and `sample_mode` (`distributed` preferred).  
4) Use `find_value` (label mode) or `range_values` for targeted pulls.  
5) Reserve `sheet_page` for unknown layouts or calculator inspection; prefer `compact`/`values_only`.  
6) Keep payloads small; page/filter rather than full-sheet reads.

## Region Detection
- Gutter-based recursive splits, trimming sparse borders.
- Header detection (multi-row) with merged-cell expansion.
- Region kinds: data, parameters, outputs, calculator, metadata, other; confidence score per region.
- Regions cached per sheet and reused by `sheet_overview`, `find_value`, `read_table`, `table_profile`.

## Quick Start
```bash
# Run HTTP streaming transport at 127.0.0.1:8079
cargo run --release -- --workspace-root /path/to/workbooks

# or install
cargo install --path .
spreadsheet-read-mcp --workspace-root /path/to/workbooks
```
Transport: HTTP streaming (default). Endpoints: `POST /mcp` for bidirectional streaming. `--transport stdio` is available for CLI pipelines.

## Configuration (flags/env)
| Flag | Env | Description |
| --- | --- | --- |
| `--workspace-root <DIR>` | `SPREADSHEET_MCP_WORKSPACE` | Workspace root to scan (default cwd). |
| `--cache-capacity <N>` | `SPREADSHEET_MCP_CACHE_CAPACITY` | Workbook cache size (default 5). |
| `--extensions <list>` | `SPREADSHEET_MCP_EXTENSIONS` | Allowed extensions (`xlsx,xls,xlsb`). |
| `--workbook <FILE>` | `SPREADSHEET_MCP_WORKBOOK` | Single-workbook mode. |
| `--enabled-tools <list>` | `SPREADSHEET_MCP_ENABLED_TOOLS` | Whitelist exposed tools. |
| `--transport <http|stdio>` | `SPREADSHEET_MCP_TRANSPORT` | Transport selection (HTTP streaming default). |
| `--http-bind <ADDR>` | `SPREADSHEET_MCP_HTTP_BIND` | Bind address (default `127.0.0.1:8079`). |

## Testing
`cargo test` runs unit + integration suites, including region detection, region-scoped tools, read_table edge cases (merged headers, filters, large sheets), and workbook summary.

## Behavior & Limits
- Strictly read-only; no mutation/recalc/VBA execution.
- XLSX parsed fully; `.xls`/`.xlsb` discovered and validated before load.
- Bounded in-memory cache honors `cache_capacity`.
- Prefer region-scoped reads and sampling for token/latency efficiency.***
