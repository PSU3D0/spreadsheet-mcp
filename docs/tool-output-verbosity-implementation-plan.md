# Implementation Plan: Token-Dense Defaults + Size-Aware Pagination

This plan translates the proposal in `docs/tool-output-verbosity-proposal.md` into phased implementation with per-tool work, tests, and definition of done.

## Guiding Principles
- Preserve current capabilities; make lightweight output the default.
- Use clean pagination: `limit`/`offset` params, `next_offset` response field only when more data exists.
- No `truncated` or `has_more` fields - `next_offset` presence implies more data.
- Keep changes opt-out via config for legacy behavior.

## Phase 0: Shared Infrastructure

**Scope**
- Introduce output profile and size caps.
- Build reusable size/pagination helpers.

**Changes**
- Config: `SPREADSHEET_MCP_OUTPUT_PROFILE={token_dense|verbose}` (default `token_dense`).
- Config: `SPREADSHEET_MCP_MAX_PAYLOAD_BYTES`, `SPREADSHEET_MCP_MAX_CELLS`, `SPREADSHEET_MCP_MAX_ITEMS`.
- Add helpers for:
  - payload size estimation
  - truncation decisions
  - optional pagination fields (`truncated`, `next_offset`, etc.)

**Tests**
- Unit tests for config defaults and overrides.
- Unit tests for pagination field omission when not truncated.

**Definition of Done**
- Configs documented and plumbed into `AppState`.
- Helpers available and used in at least one tool in Phase 1.
- No new fields serialized when not needed.

## Phase 1: Core Data Reads

### read_table
**Changes**
- Add `format: csv|values|json` (default `csv` in token_dense profile).
- Add `include_headers`/`include_types` (default `include_headers=true`, `include_types=false`).
- Add size-aware truncation with `truncated` + `next_offset` only when needed.

**Tests**
- `read_table_format_csv_default`: default output is CSV when profile token_dense.
- `read_table_format_values`: returns headers + value arrays without `CellValue` wrappers.
- `read_table_pagination_omitted_when_small`: no pagination fields for small tables.
- `read_table_truncates_large`: truncated + next_offset present when size limits hit.

**DoD**
- Existing JSON format preserved via `format=json` or verbose profile.
- CSV output RFC 4180 compliant.
- No regression in existing `read_table_polish` tests.

### range_values
**Changes**
- Add `format: values|csv|json` (default `values`).
- Add size-aware row truncation with `next_start_row` only when needed.

**Tests**
- `range_values_format_values_default`.
- `range_values_format_csv`.
- `range_values_truncation_only_when_large`.

**DoD**
- Range values remain accurate across formats.
- No pagination fields for small ranges.

### sheet_page
**Changes**
- Default `format=Compact` when profile token_dense.
- Default `include_formulas=false`, `include_styles=false` in token_dense profile.
- `has_more/next_start_row` only when truncated.

**Tests**
- `sheet_page_defaults_compact_token_dense`.
- `sheet_page_full_verbose` (ensures legacy behavior still possible).
- `sheet_page_no_has_more_when_small`.

**DoD**
- Compact output smaller than full for typical sheets.
- No behavior change in verbose profile.

## Phase 2: Summary/Profiling Tools

### table_profile
**Changes**
- Add `summary_only` (default `true`).
- Add `include_samples` (default `false`).
- Add `sample_format` + `top_k`.
- Size-aware truncation for `samples` and large column lists.

**Tests**
- `table_profile_summary_only_default`.
- `table_profile_samples_values_format`.
- `table_profile_truncation_fields_only_when_needed`.

### sheet_statistics
**Changes**
- Add `summary_only` (default `true`).
- Add `include_samples` and `top_k`.
- Size-aware truncation.

**Tests**
- `sheet_statistics_summary_only_default`.
- `sheet_statistics_top_k_limit`.

**DoD (Phase 2)**
- Summary-only payloads are smaller and still useful for quick scans.
- Optional samples are available and correct.

## Phase 3: Styles Tools

### sheet_styles
**Changes**
- Add `summary_only` (default `true`).
- Add `include_descriptor`, `include_ranges`, `include_example_cells` toggles.
- Size-aware truncation for large style lists.

**Tests**
- `sheet_styles_summary_only_default`.
- `sheet_styles_toggle_fields`.
- `sheet_styles_truncation_fields_only_when_large`.

### workbook_style_summary
**Changes**
- Add `summary_only` (default `true`).
- Add `include_conditional_formats`, `include_theme` toggles.
- Size-aware truncation.

**Tests**
- `workbook_style_summary_summary_only_default`.
- `workbook_style_summary_toggle_fields`.

**DoD (Phase 3)**
- Full detail still accessible.
- Summaries do not include heavy descriptors/ranges unless requested.

## Phase 4: Formula and Search Tools

### sheet_formula_map
**Changes**
- Add `summary_only` (default `true`).
- Add `include_addresses` and `addresses_limit`.
- Size-aware truncation with `next_offset` only when needed.

**Tests**
- `sheet_formula_map_summary_only_default`.
- `sheet_formula_map_addresses_limit`.

### find_value
**Changes**
- Add `context: none|neighbors|row|both` (default `none`).
- Add `context_width` (optional).
- Size-aware result truncation.

**Tests**
- `find_value_context_none_default`.
- `find_value_context_neighbors`.

### find_formula
**Changes**
- Add `context_rows`/`context_cols` limits.
- Size-aware truncation fields only when needed.

**Tests**
- `find_formula_context_bounds`.
- `find_formula_no_next_offset_when_small`.

**DoD (Phase 4)**
- Searches return minimal info by default but support rich context on demand.
- Pagination fields appear only on truncation.

## Phase 5: Pagination Cleanup (Align All Tools)

**Scope**
- Remove `truncated` and `has_more` fields from all responses.
- Standardize on `next_offset` (or `next_start_row` for row-based tools) only.
- Use `skip_serializing_if` so pagination fields disappear when not needed.

**Tools to update:**
| Tool | Current | Target |
|------|---------|--------|
| `read_table` | `has_more` + `next_offset` | `next_offset` only |
| `sheet_page` | `has_more` + `next_start_row` | `next_start_row` only |
| `range_values` | `truncated` + `next_start_row` | `next_start_row` only |
| `table_profile` | `truncated` | remove (no pagination) |
| `sheet_statistics` | `truncated` | remove (no pagination) |
| `sheet_formula_map` | `truncated` + `next_offset` | `next_offset` only |
| `find_formula` | `truncated` + `next_offset` | `next_offset` only |
| `find_value` | `truncated` | add `next_offset` |
| `scan_volatiles` | `truncated` + `next_offset` | `next_offset` only |
| `sheet_styles` | `styles_truncated` | keep (different semantic: style list truncated) |
| `workbook_style_summary` | `styles_truncated` + `conditional_formats_truncated` | keep (different semantic) |

**Calculation pattern:**
```rust
let returned_count = items.len();
let next_offset = if offset + returned_count < total_count {
    Some((offset + returned_count) as u32)
} else {
    None
};
```

**Tests**
- Verify `next_offset` absent when all data returned.
- Verify `next_offset` present and correct when more data exists.
- Verify no `truncated` or `has_more` fields in any response.

**DoD**
- All list-returning tools use consistent pagination pattern.
- Response JSON is minimal when data fits in one page.

## Phase 6: Metadata and Listing Tools

### list_workbooks / list_sheets / workbook_summary
**Changes**
- Add `limit/offset` optional pagination.
- Add `include_paths` (list_workbooks) and `include_bounds` (list_sheets) toggles.
- Follow clean pagination pattern (no `truncated`, only `next_offset`).

**Tests**
- `list_workbooks_limit_offset`.
- `list_sheets_include_bounds`.
- `list_workbooks_no_pagination_when_small`.

**DoD (Phase 6)**
- Large workspaces can be paged without heavy payloads by default.

## Phase 7: Docs, Migration, and Rollout

**Changes**
- Update `BASE_INSTRUCTIONS` to reflect new defaults and recommended usage.
- Update README/ERGONOMICS with format and pagination details.
- Add changelog note (if repository has one).

**Tests**
- Documentation lint (if applicable).
- Ensure existing integration tests pass with verbose profile enabled when needed.

**DoD**
- Clear migration guidance for clients.
- All tests green.

## Test Coverage Summary
- New unit tests for each tool’s default output format.
- Truncation behavior tests to validate pagination fields appear only on truncation.
- Regression tests to ensure legacy verbose behavior still works via profile or explicit params.

## Rollout Strategy
1. Land Phase 0 + Phase 1 behind `output_profile=token_dense` default.
2. Gate verbose output via `output_profile=verbose` or explicit params.
3. Document and communicate the change; update MCP instructions.

## Project Definition of Done
- All phases implemented and merged.
- Tests for defaults and truncation behavior added for each tool class.
- Docs updated with new defaults, formats, and pagination rules.
- Backwards compatibility maintained via `output_profile=verbose`.
