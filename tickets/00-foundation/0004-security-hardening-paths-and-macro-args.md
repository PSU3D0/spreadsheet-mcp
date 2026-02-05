# Ticket: 0004 Security Hardening (Paths + LibreOffice Macro Args)

## Why (Human Operator Replacement)
Scaling spreadsheet automation increases blast radius. If path boundaries or macro argument construction are exploitable, this becomes a production blocker.

## Scope
1) Workspace path boundary enforcement
- Canonicalize paths before boundary checks.
- Reject paths that escape workspace root after normalization.
- Add symlink-aware checks.

2) LibreOffice macro argument hardening
- Sanitize/escape sheet names and range arguments passed into macro URIs.
- Reject unsafe characters where escaping is ambiguous.

## Non-Goals
- No new authentication/authorization layer.
- No multi-tenant isolation guarantees.

## Proposed Tool Surface
- No new tools.
- Improve errors for unsafe inputs (invalid params with precise field paths).

## Implementation Notes
- Files likely involved:
  - `src/config.rs` (path resolution)
  - `src/fork.rs` (save_fork target path enforcement)
  - `src/recalc/screenshot.rs` and other LO executor code
- Prefer `std::fs::canonicalize` on both workspace root and resolved target, then compare.
- For macro args:
  - allowlist sheet names for macro calls OR escape for Basic string literal context.

## Tests
- Path traversal tests:
  - `../` components
  - symlink inside workspace pointing outside
  - unicode normalization edge cases (if applicable)
- Macro arg tests:
  - sheet names with quotes/semicolons
  - ensure sanitized output is safe or rejected

## Definition of Done
- Demonstrably impossible to save outside workspace root.
- LibreOffice macro invocations cannot be injected via user-controlled sheet/range strings.

## Rollout Notes
- This may break previously-accepted unsafe paths; document in CHANGELOG.
