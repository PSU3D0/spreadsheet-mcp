# Surface Capability Matrix (CLI / MCP / WASM / SDK)

Status: draft (Tranche 35 foundation with drift checks)
Owner: Tranche 35 (tickets/35-js-surface-migration)

This matrix is the planning baseline for cross-surface migration.

## Legend

- **Classification**
  - `ALL` = intended shared capability (CLI + MCP + WASM via shared core)
  - `CLI_ONLY` = host/operator concern (no MCP/WASM parity required)
  - `MCP_ONLY` = agent/session orchestration concern
  - `SHARED_PARTIAL` = shared semantics, but currently only implemented on subset
- **WASM target**
  - `mvp` = planned for initial WASM surface
  - `later` = planned after MVP
  - `n/a` = intentionally not a WASM concern

Boundary contract: `docs/architecture/surface-boundary-rules.md`

---

## A) CLI command catalog

| CLI command/subcommand | MCP equivalent tool(s) | Classification | Core projection target | WASM target | Notes | Implementation module path | Parity test owner |
|---|---|---:|---|---:|---|---|---|
| `list-sheets` | `list_sheets` | ALL | `core.read.list_sheets` | mvp | Shared read primitive | `crates/spreadsheet-kit/src/cli/commands/read.rs::list_sheets` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `sheet-overview` | `sheet_overview` | ALL | `core.read.sheet_overview` | mvp | Shared read primitive | `crates/spreadsheet-kit/src/cli/commands/read.rs::sheet_overview` | `crates/spreadsheet-kit/tests/sheet_overview_truncation.rs` |
| `range-values` | `range_values` | ALL | `core.read.range_values` | mvp | Shared read primitive | `crates/spreadsheet-kit/src/cli/commands/read.rs::range_values` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `range-export --format json/csv` | `range_values` | ALL | `core.read.range_values` + formatter | mvp | CSV serialization shared; CLI handles output path/stdout | `crates/spreadsheet-kit/src/cli/commands/read.rs::range_export` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `range-export --format grid` | `grid_export` | ALL | `core.read.grid_export` | mvp | Rich payload export | `crates/spreadsheet-kit/src/cli/commands/read.rs::range_export` | `crates/spreadsheet-kit/tests/unit_grid_roundtrip.rs` |
| `range-import --from-grid` | `grid_import` | ALL | `core.write.grid_import` | mvp | Shared grid import semantics | `crates/spreadsheet-kit/src/cli/commands/write.rs::range_import` | `crates/spreadsheet-kit/tests/unit_grid_roundtrip.rs` |
| `range-import --from-csv` | _(none today)_ | SHARED_PARTIAL | `core.write.csv_import` | mvp | CLI has path; MCP may add later | `crates/spreadsheet-kit/src/cli/commands/write.rs::range_import` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `inspect-cells` | `inspect_cells` | ALL | `core.read.inspect_cells` | mvp | Strict detail-view: up to 25 cells with full metadata; returns budget object | `crates/spreadsheet-kit/src/cli/commands/read.rs::inspect_cells` | `crates/spreadsheet-kit/tests/read_guardrails.rs` |
| `sheet-page` | `sheet_page` | ALL | `core.read.sheet_page` | mvp | Shared pagination contract | `crates/spreadsheet-kit/src/cli/commands/read.rs::sheet_page` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `read-table` | `read_table` | ALL | `core.read.read_table` | mvp | Shared table read primitive | `crates/spreadsheet-kit/src/cli/commands/read.rs::read_table` | `crates/spreadsheet-kit/tests/read_table_polish.rs` |
| `find-value` | `find_value` | ALL | `core.analysis.find_value` | mvp | Shared analysis primitive | `crates/spreadsheet-kit/src/cli/commands/read.rs::find_value` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `named-ranges` | `named_ranges` | ALL | `core.read.named_ranges` | mvp | Shared read primitive | `crates/spreadsheet-kit/src/cli/commands/read.rs::named_ranges` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `define-name` | `define_name` | ALL | `core.write.define_name` | mvp | Named range CRUD (create) | `crates/spreadsheet-kit/src/cli/commands/write.rs::define_name` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `update-name` | `update_name` | ALL | `core.write.update_name` | mvp | Named range CRUD (update) | `crates/spreadsheet-kit/src/cli/commands/write.rs::update_name` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `delete-name` | `delete_name` | ALL | `core.write.delete_name` | mvp | Named range CRUD (delete) | `crates/spreadsheet-kit/src/cli/commands/write.rs::delete_name` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `find-formula` | `find_formula` | ALL | `core.analysis.find_formula` | mvp | Shared analysis primitive | `crates/spreadsheet-kit/src/cli/commands/read.rs::find_formula` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `scan-volatiles` | `scan_volatiles` | ALL | `core.analysis.scan_volatiles` | mvp | Shared analysis primitive | `crates/spreadsheet-kit/src/cli/commands/read.rs::scan_volatiles` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `sheet-statistics` | `sheet_statistics` | ALL | `core.analysis.sheet_statistics` | mvp | Shared analysis primitive | `crates/spreadsheet-kit/src/cli/commands/read.rs::sheet_statistics` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `formula-map` | `sheet_formula_map` | ALL | `core.analysis.sheet_formula_map` | mvp | Shared analysis primitive | `crates/spreadsheet-kit/src/cli/commands/read.rs::formula_map` | `crates/spreadsheet-kit/tests/heuristic_scenarios.rs` |
| `formula-trace` | `formula_trace` | ALL | `core.analysis.formula_trace` | later | Shared but heavier graph concerns | `crates/spreadsheet-kit/src/cli/commands/read.rs::formula_trace` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `describe` | `describe_workbook` | ALL | `core.read.describe_workbook` | mvp | Contract naming differs by surface | `crates/spreadsheet-kit/src/cli/commands/read.rs::describe` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `table-profile` | `table_profile` | ALL | `core.analysis.table_profile` | mvp | Shared profiling primitive | `crates/spreadsheet-kit/src/cli/commands/read.rs::table_profile` | `crates/spreadsheet-kit/tests/read_table_polish.rs` |
| `layout-page` | `layout_page` | ALL | `core.read.layout_page` | mvp | Shared layout primitive | `crates/spreadsheet-kit/src/cli/commands/read.rs::layout_page` | `crates/spreadsheet-kit/tests/unit_layout_page.rs` |
| `create-workbook` | _(none today)_ | SHARED_PARTIAL | `core.write.create_workbook_bytes` (planned) | later | CLI path-based today | `crates/spreadsheet-kit/src/cli/commands/write.rs::create_workbook` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `copy` | _(none today)_ | CLI_ONLY | `adapter-cli.copy_path` | n/a | Stateless file orchestration | `crates/spreadsheet-kit/src/cli/commands/write.rs::copy` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `edit` | `edit_batch` | ALL | `core.write.edit_batch` | mvp | CLI shorthand parsing is adapter concern | `crates/spreadsheet-kit/src/cli/commands/write.rs::edit` | `crates/spreadsheet-kit/tests/unit_edit_batch.rs` |
| `transform-batch` | `transform_batch` | ALL | `core.write.transform_batch` | mvp | Shared write primitive | `crates/spreadsheet-kit/src/cli/commands/write.rs::transform_batch` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `style-batch` | `style_batch` | ALL | `core.write.style_batch` | mvp | Shared write primitive | `crates/spreadsheet-kit/src/cli/commands/write.rs::style_batch` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `apply-formula-pattern` | `apply_formula_pattern` | ALL | `core.write.apply_formula_pattern` | later | Shared write primitive | `crates/spreadsheet-kit/src/cli/commands/write.rs::apply_formula_pattern` | `crates/spreadsheet-kit/tests/unit_formula_pattern.rs` |
| `structure-batch` | `structure_batch` | ALL | `core.write.structure_batch` | later | Shared write primitive | `crates/spreadsheet-kit/src/cli/commands/write.rs::structure_batch` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `column-size-batch` | `column_size_batch` | ALL | `core.write.column_size_batch` | later | Shared write primitive | `crates/spreadsheet-kit/src/cli/commands/write.rs::column_size_batch` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `sheet-layout-batch` | `sheet_layout_batch` | ALL | `core.write.sheet_layout_batch` | later | Shared write primitive | `crates/spreadsheet-kit/src/cli/commands/write.rs::sheet_layout_batch` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `rules-batch` | `rules_batch` | ALL | `core.write.rules_batch` | later | Shared write primitive | `crates/spreadsheet-kit/src/cli/commands/write.rs::rules_batch` | `crates/spreadsheet-kit/tests/core_runtime_parity.rs` |
| `replace-in-formulas` | `replace_in_formulas` | ALL | `core.write.replace_in_formulas` | later | Formula-only find/replace with dry-run | `crates/spreadsheet-kit/src/cli/commands/write.rs::replace_in_formulas` | `crates/spreadsheet-kit/tests/unit_replace_in_formulas.rs` |
| `sheetport manifest candidates` | `get_manifest_stub` | SHARED_PARTIAL | `core.sheetport.manifest_stub` | later | Naming differs | `crates/spreadsheet-kit/src/cli/commands/read.rs::sheetport_manifest_candidates` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `sheetport manifest schema` | _(none today)_ | CLI_ONLY | `adapter-cli.sheetport_schema` | n/a | Local schema print UX | `crates/spreadsheet-kit/src/cli/commands/read.rs::sheetport_manifest_schema` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `sheetport manifest validate` | _(none today)_ | CLI_ONLY | `adapter-cli.sheetport_validate_yaml` | n/a | Local manifest file validation | `crates/spreadsheet-kit/src/cli/commands/read.rs::sheetport_manifest_validate` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `sheetport manifest normalize` | _(none today)_ | CLI_ONLY | `adapter-cli.sheetport_normalize_yaml` | n/a | Local file transform concern | `crates/spreadsheet-kit/src/cli/commands/read.rs::sheetport_manifest_normalize` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `sheetport bind-check` | _(none direct)_ | SHARED_PARTIAL | `core.sheetport.bind_check` | later | Could be unified later | `crates/spreadsheet-kit/src/cli/commands/read.rs::sheetport_bind_check` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `sheetport run` | `execute_manifest` | ALL | `core.sheetport.execute_manifest` | later | Shared core semantics expected | `crates/spreadsheet-kit/src/cli/commands/read.rs::sheetport_run` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `recalculate` | `recalculate` | SHARED_PARTIAL | `core.recalc.recalculate` | later | Backend constraints in WASM | `crates/spreadsheet-kit/src/cli/commands/recalc.rs::recalculate` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `diff` | `get_changeset` (partial overlap) | SHARED_PARTIAL | `core.diff.diff_workbooks` | later | CLI is file-vs-file; MCP is fork-oriented | `crates/spreadsheet-kit/src/cli/commands/diff.rs::diff` | `crates/spreadsheet-kit/tests/diff_engine.rs` |
| `run-manifest` (deprecated) | `execute_manifest` | SHARED_PARTIAL | `core.sheetport.execute_manifest` | later | Kept for backward compatibility | `crates/spreadsheet-kit/src/cli/commands/read.rs::run_manifest` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `check-ref-impact` | _(none today)_ | CLI_ONLY | `core.analysis.structure_impact` | n/a | Read-only structural impact preflight; uses same engine as `structure-batch --dry-run --impact-report` | `crates/spreadsheet-kit/src/cli/commands/write.rs::check_ref_impact` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `session` | _(none today)_ | CLI_ONLY | `core.session.*` | n/a | Event-sourced session management (start, log, branches, switch, checkout, undo, redo, fork, op, apply, materialize) | `crates/spreadsheet-kit/src/cli/commands/session.rs` | `crates/spreadsheet-kit/tests/cli_integration.rs` |

---

## B) MCP tool catalog

| MCP tool | CLI equivalent | Classification | Core projection target | WASM target | Notes | Implementation module path | Parity test owner |
|---|---|---:|---|---:|---|---|---|
| `list_workbooks` | _(none)_ | MCP_ONLY | `adapter-mcp.workspace.list_workbooks` | n/a | Workspace/repository concern | `crates/spreadsheet-kit/src/tools/mod.rs::list_workbooks` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `describe_workbook` | `describe` | ALL | `core.read.describe_workbook` | mvp | Shared read primitive | `crates/spreadsheet-kit/src/tools/mod.rs::describe_workbook` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `workbook_summary` | _(none direct)_ | SHARED_PARTIAL | `core.analysis.workbook_summary` | later | Candidate future CLI command | `crates/spreadsheet-kit/src/tools/mod.rs::workbook_summary` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `list_sheets` | `list-sheets` | ALL | `core.read.list_sheets` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::list_sheets` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `sheet_overview` | `sheet-overview` | ALL | `core.read.sheet_overview` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::sheet_overview` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `sheet_page` | `sheet-page` | ALL | `core.read.sheet_page` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::sheet_page` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `find_value` | `find-value` | ALL | `core.analysis.find_value` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::find_value` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `read_table` | `read-table` | ALL | `core.read.read_table` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::read_table` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `table_profile` | `table-profile` | ALL | `core.analysis.table_profile` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::table_profile` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `range_values` | `range-values` | ALL | `core.read.range_values` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::range_values` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `inspect_cells` | `inspect-cells` | ALL | `core.read.inspect_cells` | mvp | Strict detail-view (≤25 cells); returns budget metadata | `crates/spreadsheet-kit/src/tools/mod.rs::inspect_cells` | `crates/spreadsheet-mcp/tests/read_guardrails_mcp.rs` |
| `sheet_statistics` | `sheet-statistics` | ALL | `core.analysis.sheet_statistics` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::sheet_statistics` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `sheet_formula_map` | `formula-map` | ALL | `core.analysis.sheet_formula_map` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::sheet_formula_map` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `formula_trace` | `formula-trace` | ALL | `core.analysis.formula_trace` | later | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::formula_trace` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `named_ranges` | `named-ranges` | ALL | `core.read.named_ranges` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::named_ranges` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `define_name` | `define-name` | ALL | `core.write.define_name` | mvp | Named range CRUD (create) | `crates/spreadsheet-kit/src/tools/mod.rs::define_name` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `update_name` | `update-name` | ALL | `core.write.update_name` | mvp | Named range CRUD (update) | `crates/spreadsheet-kit/src/tools/mod.rs::update_name` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `delete_name` | `delete-name` | ALL | `core.write.delete_name` | mvp | Named range CRUD (delete) | `crates/spreadsheet-kit/src/tools/mod.rs::delete_name` | `crates/spreadsheet-kit/tests/cli_integration.rs` |
| `find_formula` | `find-formula` | ALL | `core.analysis.find_formula` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::find_formula` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `scan_volatiles` | `scan-volatiles` | ALL | `core.analysis.scan_volatiles` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::scan_volatiles` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `sheet_styles` | _(none)_ | SHARED_PARTIAL | `core.read.sheet_styles` | later | Candidate future CLI/WASM surface | `crates/spreadsheet-kit/src/tools/mod.rs::sheet_styles` | `crates/spreadsheet-mcp/tests/unit_styles.rs` |
| `layout_page` | `layout-page` | ALL | `core.read.layout_page` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::layout_page` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `grid_export` | `range-export --format grid` | ALL | `core.read.grid_export` | mvp | Shared | `crates/spreadsheet-kit/src/tools/mod.rs::grid_export` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `workbook_style_summary` | _(none)_ | SHARED_PARTIAL | `core.analysis.workbook_style_summary` | later | Candidate future CLI/WASM surface | `crates/spreadsheet-kit/src/tools/mod.rs::workbook_style_summary` | `crates/spreadsheet-mcp/tests/unit_workbook_style_summary_recalc.rs` |
| `get_manifest_stub` | `sheetport manifest candidates` | SHARED_PARTIAL | `core.sheetport.manifest_stub` | later | Shared semantic target | `crates/spreadsheet-kit/src/tools/mod.rs::get_manifest_stub` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `execute_manifest` | `sheetport run`/`run-manifest` | ALL | `core.sheetport.execute_manifest` | later | Shared semantic target | `crates/spreadsheet-kit/src/tools/mod.rs::execute_manifest` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `close_workbook` | _(none)_ | MCP_ONLY | `adapter-mcp.session.close_workbook` | n/a | MCP resource lifecycle | `crates/spreadsheet-kit/src/tools/mod.rs::close_workbook` | `crates/spreadsheet-mcp/tests/server_smoke.rs` |
| `vba_project_summary` | _(none)_ | SHARED_PARTIAL | `core.vba.project_summary` | later | Parser/runtime constraints for WASM | `crates/spreadsheet-kit/src/tools/vba.rs::vba_project_summary` | `crates/spreadsheet-mcp/tests/unit_vba.rs` |
| `vba_module_source` | _(none)_ | SHARED_PARTIAL | `core.vba.module_source` | later | Same | `crates/spreadsheet-kit/src/tools/vba.rs::vba_module_source` | `crates/spreadsheet-mcp/tests/unit_vba.rs` |
| `create_fork` | _(none)_ | MCP_ONLY | `adapter-mcp.fork.create` | n/a | MCP orchestration | `crates/spreadsheet-kit/src/tools/fork.rs::create_fork` | `crates/spreadsheet-mcp/tests/fork_workflow.rs` |
| `edit_batch` | `edit` | ALL | `core.write.edit_batch` | mvp | Shared write semantics | `crates/spreadsheet-kit/src/tools/fork.rs::edit_batch` | `crates/spreadsheet-mcp/tests/fork_workflow.rs` |
| `transform_batch` | `transform-batch` | ALL | `core.write.transform_batch` | mvp | Shared | `crates/spreadsheet-kit/src/tools/fork.rs::transform_batch` | `crates/spreadsheet-mcp/tests/unit_transform_batch.rs` |
| `style_batch` | `style-batch` | ALL | `core.write.style_batch` | mvp | Shared | `crates/spreadsheet-kit/src/tools/fork.rs::style_batch` | `crates/spreadsheet-mcp/tests/unit_style_batch.rs` |
| `grid_import` | `range-import --from-grid` | ALL | `core.write.grid_import` | mvp | Shared | `crates/spreadsheet-kit/src/tools/fork.rs::grid_import` | `crates/spreadsheet-mcp/tests/fork_workflow.rs` |
| `column_size_batch` | `column-size-batch` | ALL | `core.write.column_size_batch` | later | Shared | `crates/spreadsheet-kit/src/tools/fork.rs::column_size_batch` | `crates/spreadsheet-mcp/tests/unit_column_size_batch.rs` |
| `sheet_layout_batch` | `sheet-layout-batch` | ALL | `core.write.sheet_layout_batch` | later | Shared | `crates/spreadsheet-kit/src/tools/sheet_layout.rs::sheet_layout_batch` | `crates/spreadsheet-mcp/tests/unit_sheet_layout_batch.rs` |
| `apply_formula_pattern` | `apply-formula-pattern` | ALL | `core.write.apply_formula_pattern` | later | Shared | `crates/spreadsheet-kit/src/tools/fork.rs::apply_formula_pattern` | `crates/spreadsheet-mcp/tests/unit_apply_formula_pattern.rs` |
| `structure_batch` | `structure-batch` | ALL | `core.write.structure_batch` | later | Shared | `crates/spreadsheet-kit/src/tools/fork.rs::structure_batch` | `crates/spreadsheet-mcp/tests/unit_structure_batch.rs` |
| `rules_batch` | `rules-batch` | ALL | `core.write.rules_batch` | later | Shared | `crates/spreadsheet-kit/src/tools/rules_batch.rs::rules_batch` | `crates/spreadsheet-mcp/tests/unit_rules_batch_cf.rs` |
| `replace_in_formulas` | `replace-in-formulas` | ALL | `core.write.replace_in_formulas` | later | Formula-only find/replace | `crates/spreadsheet-kit/src/tools/fork.rs::replace_in_formulas` | `crates/spreadsheet-mcp/tests/unit_replace_in_formulas.rs` |
| `get_edits` | _(none)_ | MCP_ONLY | `adapter-mcp.fork.edit_log` | n/a | Fork audit trail | `crates/spreadsheet-kit/src/tools/fork.rs::get_edits` | `crates/spreadsheet-mcp/tests/fork_workflow.rs` |
| `get_changeset` | `diff` (partial overlap) | SHARED_PARTIAL | `core.diff.get_changeset` + adapter projection | later | MCP is fork diff, CLI is file diff | `crates/spreadsheet-kit/src/tools/fork.rs::get_changeset` | `crates/spreadsheet-mcp/tests/fork_workflow.rs` |
| `recalculate` | `recalculate` | SHARED_PARTIAL | `core.recalc.recalculate` | later | Backend constraints | `crates/spreadsheet-kit/src/tools/fork.rs::recalculate` | `crates/spreadsheet-mcp/tests/unit_recalc_needed.rs` |
| `list_forks` | _(none)_ | MCP_ONLY | `adapter-mcp.fork.list` | n/a | MCP orchestration | `crates/spreadsheet-kit/src/tools/fork.rs::list_forks` | `crates/spreadsheet-mcp/tests/fork_workflow.rs` |
| `discard_fork` | _(none)_ | MCP_ONLY | `adapter-mcp.fork.discard` | n/a | MCP orchestration | `crates/spreadsheet-kit/src/tools/fork.rs::discard_fork` | `crates/spreadsheet-mcp/tests/fork_workflow.rs` |
| `save_fork` | _(none)_ | MCP_ONLY | `adapter-mcp.fork.save` | n/a | MCP orchestration | `crates/spreadsheet-kit/src/tools/fork.rs::save_fork` | `crates/spreadsheet-mcp/tests/fork_workflow.rs` |
| `checkpoint_fork` | _(none)_ | MCP_ONLY | `adapter-mcp.fork.checkpoint_create` | n/a | MCP orchestration | `crates/spreadsheet-kit/src/tools/fork.rs::checkpoint_fork` | `crates/spreadsheet-mcp/tests/fork_workflow.rs` |
| `list_checkpoints` | _(none)_ | MCP_ONLY | `adapter-mcp.fork.checkpoint_list` | n/a | MCP orchestration | `crates/spreadsheet-kit/src/tools/fork.rs::list_checkpoints` | `crates/spreadsheet-mcp/tests/fork_workflow.rs` |
| `restore_checkpoint` | _(none)_ | MCP_ONLY | `adapter-mcp.fork.checkpoint_restore` | n/a | MCP orchestration | `crates/spreadsheet-kit/src/tools/fork.rs::restore_checkpoint` | `crates/spreadsheet-mcp/tests/fork_workflow.rs` |
| `delete_checkpoint` | _(none)_ | MCP_ONLY | `adapter-mcp.fork.checkpoint_delete` | n/a | MCP orchestration | `crates/spreadsheet-kit/src/tools/fork.rs::delete_checkpoint` | `crates/spreadsheet-mcp/tests/fork_workflow.rs` |
| `list_staged_changes` | _(none)_ | MCP_ONLY | `adapter-mcp.fork.staged_list` | n/a | MCP orchestration | `crates/spreadsheet-kit/src/tools/fork.rs::list_staged_changes` | `crates/spreadsheet-mcp/tests/unit_staging.rs` |
| `apply_staged_change` | _(none)_ | MCP_ONLY | `adapter-mcp.fork.staged_apply` | n/a | MCP orchestration | `crates/spreadsheet-kit/src/tools/fork.rs::apply_staged_change` | `crates/spreadsheet-mcp/tests/unit_staging.rs` |
| `discard_staged_change` | _(none)_ | MCP_ONLY | `adapter-mcp.fork.staged_discard` | n/a | MCP orchestration | `crates/spreadsheet-kit/src/tools/fork.rs::discard_staged_change` | `crates/spreadsheet-mcp/tests/unit_staging.rs` |
| `screenshot_sheet` | _(none)_ | MCP_ONLY | `adapter-mcp.render.screenshot` | n/a | Rendering/tooling concern; browser has native rendering paths | `crates/spreadsheet-kit/src/tools/fork.rs::screenshot_sheet` | `crates/spreadsheet-mcp/tests/screenshot_docker.rs` |

---

## C) Enforcement hooks

- Boundary contract (non-negotiable): `docs/architecture/surface-boundary-rules.md`
- Matrix drift checker: `scripts/check_surface_matrix_drift.py`
- Local/CI invocation:
  - `python3 scripts/check_surface_matrix_drift.py`
  - `cargo test -p spreadsheet-kit surface_matrix_drift_check`
