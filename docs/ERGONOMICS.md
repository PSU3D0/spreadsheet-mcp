# Spreadsheet MCP ergonomics review (from traces)

This doc captures ergonomic frictions observed in recent agent traces (fork/recalc + styling work) and proposes fix pathways with expected impact.

## Goals

- Make the happy path obvious: `create_fork → edits → recalculate (only if needed) → inspect fork directly → save_fork at end`.
- Reduce token blowups / recompress loops.
- Reduce accidental destructive workflows ("save early just to inspect", cell-by-cell clearing, manual rebuilds).

## Observed frictions and root causes

### 1) `workbook_id` vs `fork_id` mental model keeps breaking

**Observed**
- The model repeatedly assumed read/inspect tools require a “real” `workbook_id` and that forks are only usable for write tools.
- This led to awkward workarounds (e.g., `save_fork` just to inspect; re-opening a saved copy).

**Reality (current implementation)**
- Read/inspect tools already accept fork ids in most cases because the internal `WorkbookId` type is a transparent string and `open_workbook` resolves fork ids via the fork registry.

**Root cause**
- Schema naming: many inspect tools accept `workbook_id`, which reads as “base workbook only”.
- Tool descriptions/ServerInfo do not consistently state “fork_id is valid here too”.

**Fix pathways**
- **Schema aliasing**: introduce `workbook_ref` (or `workbook_or_fork_id`) as the canonical argument name, and keep `workbook_id` as a deprecated alias.
- **Docs/ServerInfo**: explicitly state for every read/inspect tool: “`workbook_id` accepts base workbook id *or* fork id”.
- **Error messaging**: when lookup fails, include a hint like “If you meant a fork, pass fork_id here; forks are valid workbook ids.”

**Impact**
- Fewer wrong turns; fewer “save early” exports.
- Lower risk of exporting unintended intermediate states.
- Token reduction by removing extra steps/roundtrips.

---

### 2) Range clearing is too low-level (missing primitives)

**Observed**
- Clearing ranges (e.g., “blank A40:C88”) required massive `edit_batch` payloads (per-cell writes), causing cognitive overhead, token bloat, and accidental formula loss.

**Root cause**
- Only write primitive for values/formulas is per-cell `edit_batch`.
- No first-class range transform tools (clear/fill/replace).

**Fix pathways**
- Add a range-oriented transform tool (name suggestions):
  - `transform_batch` (recommended umbrella)
  - or narrower tools: `clear_range`, `fill_range`, `replace_in_range`.
- Example ops (orthogonal flags):
  - `clear_range`: `values|formulas|styles|comments` (pick a subset); optionally `keep_formulas=true`.
  - `replace_in_range`: find/replace in `values` and optionally in `formulas`.
  - `blank_values_keep_formulas`: explicit convenience op.

**Impact**
- Massive token reduction for common sheet refactors.
- Fewer accidental formula deletions.
- Better alignment with spreadsheet workflows (range-first editing).

---

### 3) Copy/move semantics were incomplete (cross-sheet)

**Observed**
- The trace hit “can’t copy between sheets” and had to checkpoint-restore + rebuild.

**Current state**
- This has been addressed by adding `dest_sheet_name` to `structure_batch` copy/move ops.

**Remaining ergonomic work**
- Ensure tool descriptions/docs clearly call out cross-sheet support and expected formula shifting behavior.

**Impact**
- Eliminates a brutal workflow cliff.

---

### 4) Style ops are hard to reason about without a cheap preview/inspect loop

**Observed**
- The model struggled with `style_batch` merge-vs-set semantics and couldn’t confidently know what changed visually.
- Default `merge` is safe but makes “normalize everything” hard.

**Root causes**
- Style is inherently hard to validate with text-only inspection.
- Current preview mode returns a summary but not a concrete, minimal “what changed” report.
- `sheet_styles` exists and supports range/region scoping, but that capability isn’t “marketed” as a preview/inspect loop and agents don’t discover it.

**Fix pathways**
- **Tooling**
  - Add `style_batch` preview enhancements:
    - return “top changed style ids”, “cells changed”, and a small set of before/after descriptors.
    - optionally return a “staged screenshot suggestion” (ranges to screenshot) rather than forcing screenshot spam.
  - Add a lighter-weight style inspection tool for ranges:
    - e.g., `range_style_summary(workbook_ref, sheet, range)` → dominant styles + counts + sample cells.
- **Docs/ServerInfo**
  - Emphasize a recommended loop:
    - `sheet_styles(scope=range|region)` → plan targets
    - `style_batch(mode=preview)` → verify
    - `screenshot_sheet` only when ambiguous.

**Additional bug surfaced: underline clearing semantics**
- The trace indicates attempts to clear underline frequently fail or appear unchanged.
- Likely causes (to verify):
  - patch JSON shape confusion for “double option” fields.
  - default font underline value ambiguity.
  - merge/set semantics not producing “explicitly none” for underline.

**Fix pathways (underline)**
- Tighten style patch schema guidance:
  - Document clearly: to clear a field use `null` (not the string `"none"`).
- Consider adding explicit enum semantics for underline:
  - `underline: "none" | "single" | ...` with a dedicated “set” behavior.
- Add a unit test proving “clear underline” works end-to-end.

**Impact**
- Less trial-and-error; fewer screenshots; fewer recompressions.

---

### 5) Recalc changeset noise / semantics aren’t obvious

**Observed**
- Anxiety about `#VALUE!` surfacing only after `recalculate`.
- Confusion about cached results showing up as diffs.
- The model feels like validation itself “introduces changes”.

**Root cause**
- `get_changeset` includes `recalc_result` diffs (same formula, different cached value).
- No easy “formulas-only diff” / “ignore recalc results” option.

**Fix pathways**
- Add `get_changeset` filters:
  - `include_types` / `exclude_types` where types include: `formula_edit`, `value_edit`, `style_edit`, `recalc_result`, table/name diffs.
  - convenience flags: `formulas_only=true`, `ignore_recalc_results=true`.
- Add pagination/summary:
  - `summary_only=true` → counts grouped by sheet + type.
  - `limit/offset` for changes list.

**Optional: “delta changeset” default**
- Proposal: default `get_changeset` should return only changes since last `get_changeset` call for that fork (unless `full_diff=true`).
- This requires storing per-fork diff baselines (see next section).

**Impact**
- Validation becomes safe/low-noise.
- Large token reductions when diff is huge.

---

### 6) Changeset “new changes only” requires a baseline concept

**Observed**
- The trace repeatedly calls `get_changeset` (whole fork, then per-sheet) and gets huge payloads, forcing recompress.

**Proposed semantics**
- Introduce an explicit baseline selector for `get_changeset`:
  - `baseline=base` (current behavior)
  - `baseline=checkpoint:<id>`
  - `baseline=last_changeset` (delta since last call)
  - `baseline=last_save` (delta since last `save_fork`)

**Implementation pathways**
- Store one or more baseline snapshots per fork (path-based) and diff against them.
  - Minimal: store `last_changeset_snapshot_path` for `baseline=last_changeset`.
  - More robust: store hashes + sheet map, avoid copying full xlsx.

**Impact**
- `get_changeset` becomes cheap by default.
- Better UX: “show me what changed since I last looked”.

---

### 7) Token inefficiency: payloads are too heavy by default

This trace shows recompression loops even without screenshots.

#### 7.1 `find_formula` returns row context

**Observed**
- `find_formula` returns `context: Vec<RowSnapshot>` per match.

**Is it an entire row?**
- A `RowSnapshot` is a row containing a vector of cells (addresses + values + optional formulas).
- In practice this can be “a lot of cells” and multiplies by matches.

**Fix pathways**
- Add `include_context` flag (default `false`).
- Add knobs: `context_rows=0|1|2`, `context_cols=0|1|…` (or `context_mode=none|neighbors|row_window`).
- Or split into separate tools:
  - `find_formula_addresses` (fast)
  - `get_formula_context` (slow, opt-in)

**Impact**
- Avoids runaway payloads on large sheets.

#### 7.2 `sheet_page format=full` is huge

**Observed**
- Models repeatedly call `sheet_page` with `format=full`, which returns per-cell objects.

**Fix pathways**
- Improve docs/ServerInfo: “prefer `values_only` / `compact`”.
- Consider making default `format=values_only` (breaking change risk) or add a new `sheet_values_page` tool with minimal payload.
- Add response trimming options: `max_cells`, `max_cols`, `max_bytes`.

**Impact**
- Fewer recompressions; faster iterations.

#### 7.3 `get_changeset` is unbounded

**Fix pathways**
- Add summary/pagination + baseline/delta (above).

**Impact**
- Prevents the single biggest token bomb.

---

### 8) Screenshot tool semantics and docs drift

**Observed**
- Trace still passes `return_image=true` even though it was removed.
- Tool descriptions/docs still mention “returns file URI” which implies “not inline image”.

**Fix pathways**
- Update tool schema + description to:
  - always return inline `image/png` content
  - always persist PNG under `workspace_root/screenshots/`
  - ignore/remove `return_image` everywhere; ideally reject unknown args to surface drift early.
- Update README and ServerInfo instructions accordingly.

**Impact**
- Fewer wrong turns; fewer “save just to see image”.

---

## Proposed roadmap (phased)

### Phase 0: documentation + instruction fixes (lowest risk)
- Normalize tool descriptions across all inspect tools: “`workbook_id` accepts base workbook id or fork id”.
- Update ServerInfo WRITE_INSTRUCTIONS to highlight:
  - use `sheet_styles` / `range_values` / `sheet_page values_only`
  - screenshots don’t require `save_fork`.

### Phase 1: token controls + payload knobs
- Add `include_context=false` defaults and knobs to `find_formula`/`find_value`.
- Add `get_changeset summary_only`, plus `limit/offset`.

### Phase 2: diff noise control
- Add `get_changeset exclude_types=[recalc_result]` and `formulas_only`.

### Phase 3: range transform primitives
- Add `transform_batch` with `clear_range` + `replace_in_range` at minimum.

### Phase 4: baseline/delta diffs
- Add baseline selector to `get_changeset` (or a separate `get_changeset_delta`).

### Phase 5: style preview improvements
- Enrich `style_batch preview` with meaningful diffs and/or “suggested screenshot ranges”.

---

## Acceptance criteria / measurable impact

- Agents can inspect forks with `sheet_page`, `range_values`, `find_*`, `sheet_styles` without `save_fork`.
- Typical “clear helper block” workflows use a single range op, not hundreds of cell edits.
- `get_changeset` can be made safe/low-noise for validation by excluding `recalc_result`.
- Recompression frequency drops substantially on large sheets:
  - fewer calls to `sheet_page format=full`
  - fewer repeated full-fork changeset fetches.

---

## Open questions

- For `get_changeset` default behavior, do we want:
  - default = `baseline=base` (current), or
  - default = `baseline=last_changeset` with `full_diff=true` opt-in?
- For `clear_range`, what’s the default safety stance?
  - clear values but keep formulas, or
  - clear values+formulas unless specified?
- How should “error-like strings” (e.g., `#N/A`) be handled in `edit_batch`?
  - treat as literal strings, or
  - provide an explicit `CellValue::Error` write path?
