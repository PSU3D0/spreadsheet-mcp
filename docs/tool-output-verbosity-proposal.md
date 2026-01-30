# Proposal: Token-Dense Outputs and Size-Aware Pagination

## Background
The current tool surface favors rich, typed JSON payloads (e.g., `CellValue` with kind/value wrappers). That is useful for precise reasoning, but it can be very expensive for large tables, and the transcript shows `read_table` dominating token usage. The goal here is to preserve capability while making the default experience lighter and more scalable.

## Goals
- Reduce token usage for common exploration tasks.
- Keep rich/typed outputs available when needed.
- Add pagination only when output sizes exceed safe thresholds.
- Make defaults predictable and discoverable via tool docs.

## Non-goals
- Changing core semantics of tools or spreadsheet parsing logic.
- Removing any existing capabilities.
- Introducing a new tool set (prefer incremental changes).

## Cross-Tool Conventions

### Output Profiles
Introduce a server-level default profile, with per-request overrides.

- `output_profile` (config): `token_dense` (default), `verbose`.
- Request-level overrides: `format`, `include_*`, `summary_only`.

### Output Formats (for table-like tools)
Standardize across tools:

- `json` (current): typed `CellValue` objects.
- `values`: arrays of raw values (`null` for empty). No type wrappers.
- `csv`: RFC 4180 compatible string. Header row included by default.

### Size-Aware Pagination (only when needed)
Pagination is opt-in if the caller passes `limit/offset`. Otherwise, the server should only paginate when output size would exceed configured limits.

**Config defaults (proposed):**
- `SPREADSHEET_MCP_MAX_PAYLOAD_BYTES=65536`
- `SPREADSHEET_MCP_MAX_CELLS=10000`
- `SPREADSHEET_MCP_MAX_ITEMS=500`

**Response behavior:**
- If the response fits, omit pagination fields entirely.
- If truncated, include `truncated=true` and the next cursor fields (`next_offset`, `next_start_row`, or `next_cursor` depending on tool).
- Use `serde(skip_serializing_if = "Option::is_none")` for these optional fields to avoid noise.

## Default Output Matrix

| Tool | Proposed default output | Verbose option | Pagination fields only when too large |
| --- | --- | --- | --- |
| `read_table` | `format=csv` | `format=json` | `truncated`, `next_offset` |
| `range_values` | `format=values` | `format=json` | `truncated`, `next_start_row` |
| `sheet_page` | `format=compact` + no formulas | `format=full` | `truncated`, `next_start_row` |
| `table_profile` | `summary_only=true` | `include_samples=true` | `truncated`, `next_offset` |
| `sheet_statistics` | `summary_only=true` | `include_samples=true` | `truncated`, `next_offset` |
| `sheet_styles` | `summary_only=true` | `include_descriptor=true` | `truncated`, `next_offset` |
| `workbook_style_summary` | `summary_only=true` | full style payload | `truncated`, `next_offset` |
| `sheet_formula_map` | `summary_only=true` | `include_addresses=true` | `truncated`, `next_offset` |
| `find_value` | `context=none` | `context=neighbors/row` | `truncated`, `next_offset` |
| `find_formula` | no context | `include_context=true` | `truncated`, `next_offset` |
| `named_ranges` | minimal fields | full metadata | `truncated`, `next_offset` |
| `scan_volatiles` | summary counts | include addresses | `truncated`, `next_offset` |
| `list_workbooks` | minimal fields | include full path/metadata | `truncated`, `next_offset` |
| `list_sheets` | names only | include bounds/kind | `truncated`, `next_offset` |

## Tool-Specific Proposals

### read_table
**Motivation:** Highest token cost in the transcript.

**New params:**
- `format: "csv" | "values" | "json"` (default `csv`)
- `include_headers: bool` (default `true` for csv)
- `include_types: bool` (default `false`)
- `max_rows`, `max_cells`, `max_bytes` (optional overrides)

**Response (examples):**
- `format=csv`: `{ "csv": "colA,colB\n1,2\n", "total_rows": 19 }`
- `format=values`: `{ "headers": [..], "rows": [[..],[..]] }`
- `format=json`: existing structure

**Pagination:**
- Only return `truncated` and `next_offset` when output is cut due to size limits.
- If caller sets `limit/offset`, honor them and still include `truncated` only if partial.

### range_values
**Motivation:** Often used for spot checks but can explode with large ranges.

**New params:**
- `format: "values" | "csv" | "json"` (default `values`)
- `include_addresses: bool` (default `false`)
- `page_size` (optional)

**Response:**
- `format=values`: `ranges: [{ range, values: [[..]] }]`
- `format=csv`: `ranges: [{ range, csv: "..." }]`

**Pagination:**
- For wide or tall ranges, truncate by rows and return `next_start_row` only when needed.

### sheet_page
**Motivation:** Raw cell dumps are heavy; the tool already has format options.

**Proposed defaults:**
- `format=Compact`
- `include_formulas=false`
- `include_styles=false`

**Pagination:**
- Keep existing page size controls.
- Only include `has_more/next_start_row` when truncation occurs.

### table_profile
**Motivation:** Samples and `CellValue` wrappers add weight, especially for wide tables.

**New params:**
- `summary_only: bool` (default `true`)
- `include_samples: bool` (default `false`)
- `sample_format: "values" | "csv" | "json"` (default `values`)
- `top_k` for top values (default `5`)

**Pagination:**
- Only if `samples` or large `column_types` exceed limits; return `truncated` + `next_offset`.

### sheet_statistics
**Motivation:** Similar to `table_profile` but scoped to sheets; can be large.

**New params:**
- `summary_only: bool` (default `true`)
- `include_samples: bool` (default `false`)
- `top_k` (default `5`)

**Pagination:**
- Return `truncated` only when column stats exceed size limits.

### sheet_styles
**Motivation:** Style descriptors and ranges get large quickly.

**New params:**
- `summary_only: bool` (default `true`)
- `include_descriptor: bool` (default `false`)
- `include_ranges: bool` (default `false`)
- `include_example_cells: bool` (default `false`)

**Pagination:**
- `truncated` and `next_offset` only if style list exceeds limits.

### workbook_style_summary
**Motivation:** Whole-workbook style dumps are heavy.

**New params:**
- `summary_only: bool` (default `true`)
- `include_conditional_formats: bool` (default `false`)
- `include_theme: bool` (default `false`)

**Pagination:**
- Return `truncated` only when style lists exceed limits.

### sheet_formula_map
**Motivation:** Address lists are large; most users want counts first.

**New params:**
- `summary_only: bool` (default `true`)
- `include_addresses: bool` (default `false`)
- `addresses_limit` (default `15` when enabled)

**Response:**
- Summary payload includes `formula`, `count`, `complexity` (if available).

**Pagination:**
- Only include `next_offset` if formula group list is truncated.

### find_value
**Motivation:** `neighbors` and `row_context` are useful but often unnecessary.

**New params:**
- `context: "none" | "neighbors" | "row" | "both"` (default `none`)
- `context_width` (optional)
- `max_matches` (optional)

**Pagination:**
- `truncated` + `next_offset` only when results exceed limits.

### find_formula
**Motivation:** Context rows can get large; many users just want addresses.

**New params:**
- `context_rows`, `context_cols` (optional)
- `max_matches` (optional)

**Pagination:**
- Keep `next_offset` but only include it when truncated.

### named_ranges
**Motivation:** Large workbooks may have many named ranges.

**New params:**
- `summary_only: bool` (default `true`)
- `limit/offset` (optional)

**Pagination:**
- `truncated` + `next_offset` only when needed.

### scan_volatiles
**Motivation:** Addresses per volatile can be noisy.

**New params:**
- `summary_only: bool` (default `true`)
- `include_addresses: bool` (default `false`)
- `addresses_limit` (optional)

**Pagination:**
- `truncated` only if address lists exceed limits.

### list_workbooks / list_sheets / workbook_summary
**Motivation:** In large workspaces, these can be huge.

**New params:**
- `limit/offset` (optional)
- `include_paths: bool` (default `false` for list_workbooks)
- `include_bounds: bool` (default `false` for list_sheets)

**Pagination:**
- Only include `truncated` + `next_offset` when results exceed limits.

## Backwards Compatibility
- Add new params in a backwards-compatible way; existing clients can still request `format=json`.
- For defaults, gate changes behind `output_profile=token_dense` with a config flag to keep legacy behavior.
- Document the new defaults in `BASE_INSTRUCTIONS` to reduce confusion for LLMs.

## Implementation Notes
- Reuse existing `CellValue` conversion for `values` and `csv` formats.
- Ensure `csv` output uses RFC 4180 quoting and consistent line endings.
- Use `serde(skip_serializing_if = "Option::is_none")` for pagination fields.
- Add tests to validate that pagination fields are omitted when not needed.

## Suggested Next Steps
1. Implement `format` and size-aware pagination in `read_table` and `range_values`.
2. Flip `sheet_page` defaults to compact output.
3. Add `summary_only` defaults for styles and statistics tools.
4. Update tool documentation and a few unit tests to reflect new defaults.
