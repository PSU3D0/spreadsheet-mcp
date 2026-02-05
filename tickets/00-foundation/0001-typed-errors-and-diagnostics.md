# Ticket: 0001 Typed Errors + Actionable Diagnostics

## Why (Human Operator Replacement)
Autonomous agents need to self-correct. If validation/deserialization failures surface as generic `internal_error`, agents cannot reliably repair calls and will burn tokens retrying. Human operators get precise UI errors; we need the tool equivalent.

## Scope
- Introduce a consistent error taxonomy for tool failures and map to MCP error codes.
- Wrap serde/validation errors with actionable hints and minimal canonical examples.
- Keep the on-wire schema stable unless needed to improve diagnostics.

## Non-Goals
- No behavioral changes to spreadsheet edits themselves.
- No preview/apply default changes.

## Proposed Tool Surface
- No new tools.
- Standardize error shape at MCP level:
  - invalid request -> MCP `-32602` (invalid params)
  - missing fork/workbook -> MCP `not_found` (if supported by server), otherwise `-32004`
  - conflicts (base changed, overwrite denied) -> MCP conflict code
  - internal -> MCP internal error

Message guidance:
- Include `path` to failing field when possible (e.g., `ops[0].target.kind`).
- Include 1 minimal example payload string per tool on invalid params.
- For enums, include valid variants + closest suggestion (Levenshtein) when safe.

## Implementation Notes
- Add a small, shared error type:
  - `src/errors.rs` (recommended) or `src/model.rs`
  - variants: InvalidParams, NotFound, Conflict, ResourceExhausted, Internal
- Update error mapping in `src/server.rs`:
  - Detect `serde_json::Error` and return InvalidParams with rich message.
  - Detect common `anyhow!` contexts (fork not found, sheet not found, etc.) and map.
- Update write-tool normalization layers to return `InvalidParams` rather than `anyhow!`.

## Tests
- Add unit tests that call a tool handler with intentionally malformed params and assert:
  - MCP error code matches expectation
  - message includes valid variants and a minimal example
Examples to cover:
- `structure_batch` missing `kind`/`op`
- `style_batch` wrong `fill.kind`
- `edit_batch` shorthand missing '='

## Definition of Done
- Top 10 common schema mistakes produce `invalid params` with actionable hints.
- No more generic `internal_error` for user-caused request shape errors.

## Rollout Notes
- This is backward compatible and should reduce support load.
