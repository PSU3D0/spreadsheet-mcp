# Tickets

This directory defines the roadmap as detailed, parallelizable work units.

Goals:
- Drive spreadsheet-mcp toward robust, production-grade "AI replaces a human spreadsheet operator" workflows.
- Keep the tool surface tight and model-proof.
- Prefer umya-spreadsheet-backed edits; defer LibreOffice/UNO editing paths and deep umya expansions.

## Tool Surface Strategy (Keep Tight)

Existing write tools remain the backbone (fork + preview/apply + diff + checkpoint):
- `edit_batch`, `transform_batch`, `style_batch`, `structure_batch`, `apply_formula_pattern`, `column_size_batch`

We add only two new write tools in the near term:
- `sheet_layout_batch`: view/layout actions (freeze panes, zoom, gridlines, print setup)
- `rules_batch`: data integrity + board-ready rules (data validation, conditional formatting)

## Tranches (Parallel Worktrees)

Suggested parallel worktrees by tranche:

1) `tickets/00-foundation/` (P0/P1)
- Error typing/mapping and correctness signals.
- Security hardening.

2) `tickets/10-layout/` (umya-backed)
- Freeze panes + basic view.
- Print/page layout primitives.

3) `tickets/20-rules/` (umya-backed)
- Data validation v1.
- Conditional formatting v1.
- Number-format shorthands (extends style_batch normalization).

4) `tickets/90-deferred/` (tracked, not started)
- UNO/LibreOffice editing path.
- Deep umya expansions for structural rewrite correctness.

## Ticket Format

Each ticket includes:
- Why it matters for "human spreadsheet operator replacement"
- Proposed user-facing surface (tool schema, defaults, warnings)
- Implementation plan (files/paths)
- Tests to add
- Definition of done
