# Architecture Plan: Universal Ops & Session DSL

## 1. Executive Summary & Vision

The `asp` CLI and underlying `spreadsheet-kit` library currently operate in a **file-mutator paradigm**. While operations are grouped conceptually (e.g., `structure-batch`, `style-batch`), they execute against a target workbook via destructive replacement or ephemeral `/tmp` files (via the `ForkRegistry`). 

This architecture introduces significant friction for autonomous agents and complex workflows:
- **No Native Undo:** If an agent inserts a row and breaks the model, the only recovery is completely restarting from the base copy.
- **State Drift:** Without a strict event log, determining *why* a cell holds a specific value requires manual sleuthing.
- **Safety Limitations:** Applying structural edits across huge grids often results in `#REF!` cascades. Currently, agents find out *after* the file is overwritten.
- **Split Brain Evaluation:** Structural edits are performed by the `umya` abstraction layer (via `spreadsheet-kit` rewriters), while formula recalculation is performed by the `formualizer` engine, risking parity drift.
- **Discoverability Drift:** Several safety/analysis capabilities already exist (e.g., `structure-batch --impact-report`, `--show-formula-delta`, `named-ranges`, `formula-trace`, scoped `diff`), but agents routinely miss them due to skill/docs gaps and inconsistent defaults.

### The Solution: Event-Sourced Sessions ("Git for Spreadsheets")
We will transition from a destructive file-mutator model to a **persistent, event-sourced Binlog model**. Every effectful action is expressed as a normalized `OpEvent`, including all writable command families (cell edits, imports, transforms, styles, formula patterns, structure, names, layout/rules, and formula text replacement). A "session" is a base workbook reference plus an append-only JSONL log (`events.jsonl`) with a movable `HEAD`. Materialized `.xlsx` files are compiled projections (snapshots) of this log.

### Key Benefits Realized
1. **Deterministic Replay & Debugging:** Any session can be replayed from the base file. Buggy agent logic is reproducible.
2. **Time Travel (Undo/Redo & Branching):** `HEAD` navigation enables branch-aware undo/redo and scenario forks.
3. **Agent Safety via Guardrails:** Ops declare `preconditions` and staged impact. Drift or policy violations reject apply.
4. **Auditable Provenance:** Every mutation and derived value can trace back to a specific `OpEvent`, timestamp, actor, and parent lineage.
5. **CI/CD Integration:** Event streams become reviewable change artifacts that can be validated before materialization.
6. **Unified Ops DSL:** A consistent, schema-versioned opcode layer simplifies automation, observability, and future engine convergence.

### Current Reality Check from Scenario Runs
- The toolset already supports critical checks that must be elevated in guidance:
  - `structure-batch --dry-run --impact-report --show-formula-delta`
  - `named-ranges`
  - `formula-trace`
  - `diff --sheet/--range`
  - `range-values --format json|values` (avoid manual `dense_v1` decoding)
- High-friction gaps still remain and should be productized:
  - safer preflight (`check-ref-impact` style command)
  - better post-flight observability (`recalculate` baseline filtering + changed-cell summaries)
  - inspection ergonomics (`inspect-cells` budget or row/rect detail mode)
  - easier row-oriented range reads (to avoid dense encoding mistakes)

---

## 2. The Universal `OpEvent` Envelope (Backend)

Currently, operations are fragmented into `StructureOp`, `StyleOp`, `TransformOp`, etc. These will be wrapped into a universal, canonical envelope that records intent, provenance, preconditions, and calculated impact. The envelope must be stable across versions via explicit schema versioning and canonical payload normalization/hashing.

### Core Schema (`OpEvent`)

```json
{
  "schema_version": "ops.v1",
  "op_id": "op_01J8F3...",            // ULID or sorted UUID
  "parent_id": "op_01J8F2...",        // Previous op in this branch
  "session_id": "sess_fpna_q3",
  "timestamp": "2024-10-24T10:00:00Z",
  "actor": {
    "id": "agent:fp_and_a_bot",
    "run_id": "run_abc",
    "source": "cli"
  },

  "kind": "structure.clone_row",      // Namespaced operation kind
  "payload": {
    "sheet_name": "Provider",
    "source_row": 85,
    "insert_at": 86,
    "expand_adjacent_sums": false
  },
  "canonical_payload_hash": "sha256:...", // Hash over canonicalized payload

  "preconditions": {
    "cell_matches": [
      { "address": "Provider!A85", "value": "susan.devers@heartplace.com" }
    ],
    "workbook_hash_before": "sha256:...",
    "head_at_stage": "op_01J8F2..."
  },

  "dry_run_impact": {
    "cells_changed": 9,
    "formulas_rewritten": 45,
    "shifted_spans": [
      { "sheet": "Provider", "axis": "row", "at": 86, "count": 1 }
    ],
    "ref_errors_generated": 0,
    "warnings": []
  },

  "apply_result": {
    "status": "applied",               // applied | rejected | superseded
    "duration_ms": 142,
    "warnings": [],
    "workbook_hash_after": "sha256:..."
  }
}
```

For audit integrity, each event record should optionally include `prev_event_hash`/`event_hash` (hash chain) so tampering is detectable in regulated workflows.

---

## 3. Storage & State Mechanics

The current `ForkRegistry` in `spreadsheet-kit/src/fork.rs` relies on ephemeral `/tmp` storage (`/tmp/mcp-forks`, `/tmp/mcp-checkpoints`). We will migrate this to a persistent, project-local `.asp/` directory.

### Directory Structure
```text
.asp/
  sessions/
    sess_fpna_q3/
      base.xlsx                 # Immutable base file (symlinked or copied)
      events.jsonl              # Append-only OpEvent log
      HEAD                      # Active op_id for current branch
      CURRENT_BRANCH            # Branch pointer (e.g., main, alt-1)
      branches.json             # Branch metadata: tip op_id, parent op_id, labels
      staged/
        stg_01J....json         # Staged op payload + computed impact
      snapshots/
        manifest.json           # Snapshot index: op_id -> file, lineage, hash
        op_01J8F0..._v.xlsx     # Materialized snapshot every N operations
      locks/
        session.lock            # Exclusive apply lock / lease metadata
```

### State Resolution
When `asp` receives a read request in session mode (e.g., `asp range-values --session sess_fpna_q3 ...`):
1. Resolve session + branch + `HEAD` `op_id`.
2. Acquire a read lock (or optimistic read with HEAD verification).
3. Find the nearest snapshot from `snapshots/manifest.json`.
4. Load snapshot and replay pending `OpEvents` between snapshot tip and `HEAD`.
5. Serve the read request from that resolved state.
6. If `HEAD` changed during replay, retry or fail fast with a stale-head error.

For non-session file reads, current behavior remains unchanged.

---

## 4. The CLI Surface (`asp session`)

The CLI will expose these session mechanics natively, deprecating the use of raw `--in-place` edits on source files in favor of session-backed commands.

### Lifecycle & Navigation
```bash
# Start a new session tracking a base file
asp session start --base model.xlsx --label "Q3 Rollforward"
# -> Outputs session ID: sess_abc123

# View timeline / branches / filtered changes
asp session log --session sess_abc123
asp session log --session sess_abc123 --since op_01J8... --kind structure.*
asp session branches --session sess_abc123

# Move HEAD pointer
asp session checkout --session sess_abc123 <op_id>
asp session undo --session sess_abc123
asp session redo --session sess_abc123   # redo is branch-local

# Branching
asp session fork --session sess_abc123 --from <op_id> --label "Alternative Scenario"
asp session switch --session sess_abc123 --branch alt-scenario

# Read at session HEAD
asp range-values --session sess_abc123 --sheet "Provider" --range A1:I10
```

### Staging & Applying Ops
The separation of staging and applying is critical for agent safety and deterministic commit semantics.

```bash
# 1. STAGE: computes dry_run_impact and records a staged artifact
# Does NOT advance HEAD.
asp session op --session sess_abc123 --stage --ops @clone_vance.json

# 2. APPLY: compare-and-swap apply against current HEAD and append OpEvent atomically
asp session apply --session sess_abc123 <staged_op_id>

# Optional one-shot (still enforces preconditions + CAS)
asp session op --session sess_abc123 --apply --ops @clone_vance.json
```

`apply` must be atomic: append event, advance HEAD, and persist commit metadata as a single commit unit (with crash-safe recovery on restart).

### Export
```bash
# Compile the current HEAD into a standalone Excel file
asp session materialize --output final_q3_model.xlsx
```

---

## 5. Migration Phasing & Implementation Plan

The migration will be executed in five phases to ensure backward compatibility and minimal disruption to existing workflows.

### Phase 0: Immediate UX/Skill Alignment (Short-Term)
**Goal:** Reduce avoidable agent failures now by closing discoverability gaps and adding targeted ergonomics before full session DSL rollout.

#### Subtasks:
1. **Skill Hardening:** Update `SAFE_EDITING_SKILL.md` and `EXPLORE_SKILL.md` to require: `--impact-report`, `--show-formula-delta`, `named-ranges`, `formula-trace`, scoped `diff`, and non-dense range reads for layout discovery.
2. **Dense Read Guardrail:** Add explicit warning/help hints when `range-values` returns dense encoding without `--format`, and introduce a row-oriented output mode (`--format rows`) for direct row/cell mapping.
3. **Inspection Ergonomics:** Raise `inspect-cells` cap or add an `inspect-rows`/rectangular detail mode for 100-200 cells while preserving payload budgets.
4. **Recalc Triage UX:** Add optional baseline filtering (`--baseline-errors <file>` and/or `--ignore-sheets`) and `changed_cells` summaries.
5. **Diff Ergonomics:** Add multi-sheet filtering (`--sheet` repeatable or `--sheets`) in addition to existing single-sheet/range filters.
6. **Preflight Ref Risk:** Add a dedicated preflight command (`check-ref-impact`) for row/column insert/delete risk on absolute references.

#### Files Affected:
- `skills/SAFE_EDITING_SKILL.md`
- `skills/EXPLORE_SKILL.md`
- `crates/spreadsheet-kit/src/cli/mod.rs`
- `crates/spreadsheet-kit/src/cli/commands/read.rs`
- `crates/spreadsheet-kit/src/cli/commands/write.rs`
- `crates/spreadsheet-kit/src/cli/commands/recalc.rs`
- `crates/spreadsheet-kit/src/tools/structure_impact.rs`

#### New Tests:
- `cli_range_values_rows_format_maps_row_numbers`
- `cli_inspect_cells_extended_budget_or_rect_mode`
- `cli_recalc_baseline_filter_and_changed_cells`
- `cli_diff_multi_sheet_filter`
- `cli_check_ref_impact_flags_absolute_ref_risk`

#### Definition of Done:
- Agents can complete scenario workflows without manual dense decoding, with preflight structural risk signals and baseline-aware recalc diagnostics available directly in CLI.

### Phase 1: Storage Layer & Universal Schema
**Goal:** Replace ephemeral `/tmp` storage with persistent project-local sessions and introduce the `OpEvent` schema.

#### Subtasks:
1. **Define Schema + Registry:** Create `OpEvent` and an opcode registry in `src/core/session.rs` (or `src/core/events.rs`) that wraps all effectful write families (`edit`, `range-import`, transform/style/formula/structure/column/layout/rules, name ops, replace-in-formulas, and session materialize metadata).
2. **Versioning + Canonicalization:** Add `schema_version`, canonical payload normalization, and `canonical_payload_hash` to make replay deterministic across releases.
3. **Persistent Registry:** Refactor `ForkRegistry` in `src/fork.rs` to `.asp/sessions/` storage; preserve compatibility with existing staged-change/checkpoint concepts while moving storage off `/tmp`.
4. **Binlog Writer/Reader:** Implement append-only `events.jsonl`, event parsing, lineage validation, and corruption detection.
5. **Snapshot Manager:** Add `snapshots/manifest.json`, periodic snapshotting (e.g., every 10 ops), and compaction/replay bounds.
6. **Head + Branch Pointers:** Implement `HEAD`, branch metadata, and branch-tip update semantics.
7. **Session Locking:** Implement lock/lease semantics for apply operations to prevent concurrent writer corruption.

#### Files Affected:
- `crates/spreadsheet-kit/src/fork.rs`
- `crates/spreadsheet-kit/src/core/session.rs`
- `crates/spreadsheet-kit/src/core/types.rs` (new/modified)
- `crates/spreadsheet-kit/src/core/events.rs` (new)

#### New Tests:
- `test_binlog_append_and_read`: `events.jsonl` serialization/deserialization for all registered op kinds.
- `test_session_recovery_from_disk`: restart recovery with `HEAD`, branch tips, and snapshot replay.
- `test_snapshot_generation_and_manifest`: snapshot indexing and nearest-snapshot lookup.
- `test_session_locking_conflict`: concurrent apply attempts fail with deterministic lock errors.
- `test_schema_version_forward_compat`: unknown fields tolerated; unsupported version rejected with explicit error.

#### Definition of Done:
- Backend can create sessions, append validated `OpEvents`, persist branch/HEAD pointers, and reconstruct workbook state from snapshot+replay deterministically with locking and version-aware parsing.

---

### Phase 2: CLI Interface & Session Management
**Goal:** Expose the session mechanics to the user/agent via the `asp session` subcommand tree.

#### Subtasks:
1. **Subcommand Routing:** Add `asp session` subcommands: `start`, `log`, `branches`, `switch`, `checkout`, `undo`, `redo`, `op`, `apply`, `materialize`.
2. **Session-Aware Reads:** Add `--session <id>` (and optional `--branch`) support to read commands so reads can resolve against session HEAD explicitly.
3. **`session op --stage` Implementation:** Run dry-run impact in the same execution semantics as apply path, compute `dry_run_impact`, and persist staged artifacts.
4. **Precondition + CAS Evaluator:** Validate `cell_matches` and enforce compare-and-swap apply against expected `HEAD`.
5. **Atomic Apply Protocol:** Ensure event append + HEAD move + commit metadata are atomic and crash-recoverable.
6. **Materializer:** Implement `asp session materialize` from current HEAD.

#### Files Affected:
- `crates/spreadsheet-kit/src/cli/mod.rs`
- `crates/spreadsheet-kit/src/cli/commands/session.rs` (new)
- `crates/spreadsheet-kit/src/tools/fork.rs` (redirecting staging/apply logic)
- `crates/spreadsheet-kit/src/tools/structure_impact.rs` (integrate into staging)

#### New Tests:
- `cli_test_session_start_and_materialize`: start session, no ops, materialize equals base.
- `cli_test_session_undo_redo_branch_local`: undo/redo behavior is branch-local and deterministic.
- `cli_test_precondition_failure`: stage with precondition, mutate out-of-band, apply fails.
- `cli_test_apply_cas_head_mismatch`: staged op apply fails when HEAD advanced by another writer.
- `cli_test_session_reads_resolve_head`: read commands with `--session` return HEAD-consistent values.

#### Definition of Done:
- Users can complete end-to-end workflows via `asp session` only, with explicit session-aware reads, staged apply with CAS, and crash-safe atomic commits.

---

### Phase 3: Agent Prompts & Policy Guardrails
**Goal:** Teach autonomous agents to utilize the new session mechanics to safely mutate spreadsheets.

#### Subtasks:
1. **Rewrite `SAFE_EDITING_SKILL.md`:** Move from manual copy workflows to `asp session` lifecycle (`start` -> `stage` -> `apply` -> `materialize`).
2. **Enforce Two-Step Commits:** Require `asp session op --stage`, inspect impact JSON, then `asp session apply` only if policies pass.
3. **Codify Existing Power Tools:** Explicitly require `named-ranges`, `formula-trace`, scoped `diff`, and `structure-batch --impact-report --show-formula-delta` in structural edit workflows.
4. **Safer Structural Recipes:** Teach explicit row-add patterns using `clone_row` / `copy_range` with style inheritance and bounded value overwrites (avoid raw blank insert in modeled zones).
5. **Table & Stray Write Linting:** During stage, warn/fail if edits spill outside declared virtual-table boundaries or expected columns.
6. **Fallback Recovery Playbook:** Agents must use `undo`/`checkout`/branch fork instead of restarting from scratch.

#### Files Affected:
- `skills/SAFE_EDITING_SKILL.md`
- `skills/CLI_BATCH_WRITE_SKILL.md`
- `skills/EXPLORE_SKILL.md`
- `crates/spreadsheet-kit/src/tools/structure_impact.rs` (add linting logic)

#### New Tests:
- E2E Benchmark Test: Run `scenario-01-roll-forward` using updated instructions to ensure the agent uses session staging and avoids structural blowups.
- Prompt-Contract Test: Agent structural workflow must invoke `--impact-report`, `--show-formula-delta`, and `named-ranges` before row insertions in modeled sheets.
- Read-Surface Test: Agent discovery path should prefer non-dense row-readable outputs for layout mapping in high-risk regions.

#### Definition of Done:
- Agents reliably use session navigation (`undo`, `checkout`, branch fork) when mistakes occur, rather than restarting. Structural edit regressions drop sharply due to staged impact + boundary linting + safer row-add recipes, and agents consistently invoke built-in safety/introspection commands before high-risk edits.

---

### Phase 4: Engine Convergence (Long Term)
**Goal:** Unify the structural operation execution path. Currently, `spreadsheet-kit` uses `umya` to mutate rows/cols, rewrites token ASTs via string manipulation, and then passes the file to Formualizer to evaluate. Formualizer already has internal `ReferenceAdjuster` and `VertexEditor` mechanics capable of structural mutation.

#### Subtasks:
1. **Use Formualizer Natively in Rust Path:** Route structural operations in `spreadsheet-kit` directly to `formualizer_eval::Engine` APIs (`insert_rows`, `delete_rows`, `insert_columns`, `delete_columns`, and move/copy equivalents) instead of the `umya` token-rewrite path.
2. **Shift Execution to Graph:** Update `apply_structure_ops_to_file` and staged apply execution to mutate Formualizer graph/store first, then serialize through `umya` only at materialization boundaries.
3. **Unify `ChangeEvent` & `OpEvent`:** Map frontend `OpEvent` DSL onto backend `ChangeEvent`/journal semantics so replay and diff are engine-native.
4. **Stage/Apply Parity Enforcement:** Ensure dry-run impact is produced by the same mutation engine used for apply.

#### Files Affected:
- `crates/formualizer-eval/src/engine/eval.rs`
- `crates/formualizer-eval/src/engine/graph/editor/vertex_editor.rs`
- `crates/spreadsheet-kit/src/tools/fork.rs` (remove `umya` structural rewriter logic)
- `crates/spreadsheet-kit/src/cli/commands/write.rs` (route structure staging/apply through unified backend)

#### New Tests:
- Parity Integration Tests: inserting/deleting rows/cols via `asp session` must match Formualizer-native dependency shifts and formula rewrites.
- Stage/Apply Equivalence Tests: `dry_run_impact` predictions match applied mutation results.
- Replay Determinism Tests: replaying `events.jsonl` from same base produces hash-equal logical workbook state.

#### Definition of Done:
- `umya` is used primarily as serialization I/O. Structural mutation, dependency rewrites, and recalculation execute in Formualizer-native structures, with stage/apply parity and measurable performance gain.