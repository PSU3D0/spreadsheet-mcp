use crate::config::ServerConfig;
use crate::model::{
    CloseWorkbookResponse, FindFormulaResponse, FindValueResponse, FormulaTraceResponse,
    ManifestStubResponse, NamedRangesResponse, RangeValuesResponse, ReadTableResponse,
    SheetFormulaMapResponse, SheetListResponse, SheetOverviewResponse, SheetPageResponse,
    SheetStatisticsResponse, SheetStylesResponse, TableProfileResponse, VolatileScanResponse,
    WorkbookDescription, WorkbookListResponse, WorkbookStyleSummaryResponse,
    WorkbookSummaryResponse,
};
use crate::state::AppState;
use crate::tools;
use anyhow::Result;
use rmcp::{
    ErrorData as McpError, Json, ServerHandler, ServiceExt,
    handler::server::{router::tool::ToolRouter, wrapper::Parameters},
    model::{Implementation, ServerCapabilities, ServerInfo},
    tool, tool_handler, tool_router,
    transport::stdio,
};
use std::sync::Arc;
use thiserror::Error;

const BASE_INSTRUCTIONS: &str = "\
Spreadsheet MCP: optimized for spreadsheet analysis.

WORKFLOW:
1) list_workbooks → list_sheets → workbook_summary for orientation
2) sheet_overview for region detection (ids/bounds/kind/confidence)
3) For structured data: table_profile for quick column sense, then read_table with region_id/range, filters, sampling
4) For spot checks: range_values or find_value (label mode for key-value sheets)

TOOL SELECTION:
- table_profile: Fast column/type summary before wide reads.
- read_table: Structured table extraction. Prefer region_id or tight range; use limit + sample_mode.
- sheet_formula_map: Get formula overview. Use limit param for large sheets (e.g., limit=10). \
Use sort_by='complexity' for most complex formulas first, or 'count' for most repeated. \
Use range param to scope to specific region.
- formula_trace: Trace ONE cell's precedents/dependents. Use AFTER formula_map \
to dive deep on specific outputs (e.g., trace the total cell to understand calc flow).
- sheet_page: Raw cell dump. Use ONLY when region detection fails or for \
unstructured sheets. Prefer read_table for tabular data.
- find_value with mode='label': For key-value layouts (label in col A, value in col B). \
Use direction='right' or 'below' hints.

RANGES: Use A1 notation (e.g., A1:C10). Prefer region_id when available.

DATES: Cells with date formats return ISO-8601 strings (YYYY-MM-DD).

Keep payloads small. Page through large sheets.";

const WRITE_INSTRUCTIONS: &str = "

WRITE/RECALC TOOLS (enabled):
Fork-based editing allows 'what-if' analysis without modifying original files.

WORKFLOW:
1) create_fork: Create editable copy of a workbook. Returns fork_id.
2) Optional: checkpoint_fork before large edits.
3) edit_batch/style_batch/structure_batch/apply_formula_pattern: Apply edits to the fork.
4) recalculate: Trigger LibreOffice to recompute all formulas.
5) get_changeset: Diff fork against original. Shows cell/table/name changes.
   Optional: screenshot_sheet to capture a visual view of a range (original or fork).
6) save_fork: Write changes to file.
7) discard_fork: Delete fork without saving.

SAFETY:
- checkpoint_fork before large/structural edits; restore_checkpoint to rollback if needed.
- Preview-based tools (future) will create staged changes; use apply_staged_change or discard_staged_change.

TOOL DETAILS:
- create_fork: Only .xlsx supported. Returns fork_id for subsequent operations.
- edit_batch: {fork_id, sheet_name, edits:[{address, value, is_formula}]}. \
Formulas should NOT include leading '='.
- recalculate: Required after edit_batch to update formula results. \
May take several seconds for complex workbooks.
- get_changeset: Returns cell-level diffs with modification types: \
ValueEdit, FormulaEdit, RecalcResult, Added, Deleted. \
Use sheet_name param to filter to specific sheet.
- screenshot_sheet: {workbook_id, sheet_name, range?, return_image?}. Renders a cropped PNG for inspecting an area visually.
  workbook_id may be either a real workbook_id OR a fork_id (to screenshot an edited fork).
  Prefer return_image=true to receive the PNG in the tool response.
  If return_image=false, the PNG is written under workspace_root/screenshots/ (Docker default: /data/screenshots/).
  DO NOT call save_fork just to get a screenshot.
  If formulas changed, run recalculate on the fork first.
- save_fork: Requires target_path for new file location.
  If target_path is relative, it is resolved under workspace_root (Docker default: `/data`).
  Overwriting original requires server --allow-overwrite flag.
  Use drop_fork=false to keep fork active after saving (default: true drops fork).
  Validates base file unchanged since fork creation.
- get_edits: List all edits applied to a fork (before recalculate).
- list_forks: See all active forks.
- checkpoint_fork: Snapshot a fork to a checkpoint for high-fidelity undo.
- list_checkpoints: List checkpoints for a fork.
- restore_checkpoint: Restore a fork to a checkpoint (overwrites fork file; clears newer staged changes).
- delete_checkpoint: Delete a checkpoint.
- list_staged_changes: List staged (previewed) changes for a fork.
- apply_staged_change: Apply a staged change to the fork.
- discard_staged_change: Discard a staged change.

BEST PRACTICES:
- Always recalculate after edit_batch before get_changeset.
- Review changeset before save_fork to verify expected changes.
- Use screenshot_sheet(return_image=true) for quick visual inspection; save_fork is ONLY for exporting a workbook file.
- Discard forks when done to free resources (auto-cleanup after 1 hour).
- For large edits, batch multiple cells in single edit_batch call.";

fn build_instructions(recalc_enabled: bool) -> String {
    let mut instructions = BASE_INSTRUCTIONS.to_string();
    if recalc_enabled {
        instructions.push_str(WRITE_INSTRUCTIONS);
    } else {
        instructions.push_str("\n\nRead-only mode. Write/recalc tools disabled.");
    }
    instructions
}

#[derive(Clone)]
pub struct SpreadsheetServer {
    state: Arc<AppState>,
    tool_router: ToolRouter<SpreadsheetServer>,
}

impl SpreadsheetServer {
    pub async fn new(config: Arc<ServerConfig>) -> Result<Self> {
        config.ensure_workspace_root()?;
        let state = Arc::new(AppState::new(config));
        Ok(Self::from_state(state))
    }

    pub fn from_state(state: Arc<AppState>) -> Self {
        #[allow(unused_mut)]
        let mut router = Self::tool_router();

        #[cfg(feature = "recalc")]
        {
            router.merge(Self::fork_tool_router());
        }

        Self {
            state,
            tool_router: router,
        }
    }

    pub async fn run_stdio(self) -> Result<()> {
        let service = self
            .serve(stdio())
            .await
            .inspect_err(|error| tracing::error!("serving error: {:?}", error))?;
        service.waiting().await?;
        Ok(())
    }

    pub async fn run(self) -> Result<()> {
        self.run_stdio().await
    }

    fn ensure_tool_enabled(&self, tool: &str) -> Result<()> {
        tracing::info!(tool = tool, "tool invocation requested");
        if self.state.config().is_tool_enabled(tool) {
            Ok(())
        } else {
            Err(ToolDisabledError::new(tool).into())
        }
    }

    #[cfg(feature = "recalc")]
    fn ensure_recalc_enabled(&self, tool: &str) -> Result<()> {
        self.ensure_tool_enabled(tool)?;
        if self.state.config().recalc_enabled {
            Ok(())
        } else {
            Err(RecalcDisabledError.into())
        }
    }
}

#[tool_router]
impl SpreadsheetServer {
    #[tool(
        name = "list_workbooks",
        description = "List spreadsheet files in the workspace"
    )]
    pub async fn list_workbooks(
        &self,
        Parameters(params): Parameters<tools::ListWorkbooksParams>,
    ) -> Result<Json<WorkbookListResponse>, McpError> {
        self.ensure_tool_enabled("list_workbooks")
            .map_err(to_mcp_error)?;
        tools::list_workbooks(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "describe_workbook", description = "Describe workbook metadata")]
    pub async fn describe_workbook(
        &self,
        Parameters(params): Parameters<tools::DescribeWorkbookParams>,
    ) -> Result<Json<WorkbookDescription>, McpError> {
        self.ensure_tool_enabled("describe_workbook")
            .map_err(to_mcp_error)?;
        tools::describe_workbook(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "workbook_summary",
        description = "Summarize workbook regions and entry points"
    )]
    pub async fn workbook_summary(
        &self,
        Parameters(params): Parameters<tools::WorkbookSummaryParams>,
    ) -> Result<Json<WorkbookSummaryResponse>, McpError> {
        self.ensure_tool_enabled("workbook_summary")
            .map_err(to_mcp_error)?;
        tools::workbook_summary(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "list_sheets", description = "List sheets with summaries")]
    pub async fn list_sheets(
        &self,
        Parameters(params): Parameters<tools::ListSheetsParams>,
    ) -> Result<Json<SheetListResponse>, McpError> {
        self.ensure_tool_enabled("list_sheets")
            .map_err(to_mcp_error)?;
        tools::list_sheets(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "sheet_overview",
        description = "Get narrative overview for a sheet"
    )]
    pub async fn sheet_overview(
        &self,
        Parameters(params): Parameters<tools::SheetOverviewParams>,
    ) -> Result<Json<SheetOverviewResponse>, McpError> {
        self.ensure_tool_enabled("sheet_overview")
            .map_err(to_mcp_error)?;
        tools::sheet_overview(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "sheet_page", description = "Page through sheet cells")]
    pub async fn sheet_page(
        &self,
        Parameters(params): Parameters<tools::SheetPageParams>,
    ) -> Result<Json<SheetPageResponse>, McpError> {
        self.ensure_tool_enabled("sheet_page")
            .map_err(to_mcp_error)?;
        tools::sheet_page(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "find_value", description = "Search cell values or labels")]
    pub async fn find_value(
        &self,
        Parameters(params): Parameters<tools::FindValueParams>,
    ) -> Result<Json<FindValueResponse>, McpError> {
        self.ensure_tool_enabled("find_value")
            .map_err(to_mcp_error)?;
        tools::find_value(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "read_table",
        description = "Read structured data from a range or table"
    )]
    pub async fn read_table(
        &self,
        Parameters(params): Parameters<tools::ReadTableParams>,
    ) -> Result<Json<ReadTableResponse>, McpError> {
        self.ensure_tool_enabled("read_table")
            .map_err(to_mcp_error)?;
        tools::read_table(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "table_profile", description = "Profile a region or table")]
    pub async fn table_profile(
        &self,
        Parameters(params): Parameters<tools::TableProfileParams>,
    ) -> Result<Json<TableProfileResponse>, McpError> {
        self.ensure_tool_enabled("table_profile")
            .map_err(to_mcp_error)?;
        tools::table_profile(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "range_values",
        description = "Fetch raw values for specific ranges"
    )]
    pub async fn range_values(
        &self,
        Parameters(params): Parameters<tools::RangeValuesParams>,
    ) -> Result<Json<RangeValuesResponse>, McpError> {
        self.ensure_tool_enabled("range_values")
            .map_err(to_mcp_error)?;
        tools::range_values(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "sheet_statistics",
        description = "Get aggregated sheet statistics"
    )]
    pub async fn sheet_statistics(
        &self,
        Parameters(params): Parameters<tools::SheetStatisticsParams>,
    ) -> Result<Json<SheetStatisticsResponse>, McpError> {
        self.ensure_tool_enabled("sheet_statistics")
            .map_err(to_mcp_error)?;
        tools::sheet_statistics(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "sheet_formula_map",
        description = "Summarize formula groups across a sheet"
    )]
    pub async fn sheet_formula_map(
        &self,
        Parameters(params): Parameters<tools::SheetFormulaMapParams>,
    ) -> Result<Json<SheetFormulaMapResponse>, McpError> {
        self.ensure_tool_enabled("sheet_formula_map")
            .map_err(to_mcp_error)?;
        tools::sheet_formula_map(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "formula_trace",
        description = "Trace formula precedents or dependents"
    )]
    pub async fn formula_trace(
        &self,
        Parameters(params): Parameters<tools::FormulaTraceParams>,
    ) -> Result<Json<FormulaTraceResponse>, McpError> {
        self.ensure_tool_enabled("formula_trace")
            .map_err(to_mcp_error)?;
        tools::formula_trace(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "named_ranges", description = "List named ranges and tables")]
    pub async fn named_ranges(
        &self,
        Parameters(params): Parameters<tools::NamedRangesParams>,
    ) -> Result<Json<NamedRangesResponse>, McpError> {
        self.ensure_tool_enabled("named_ranges")
            .map_err(to_mcp_error)?;
        tools::named_ranges(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "find_formula", description = "Search formulas containing text")]
    pub async fn find_formula(
        &self,
        Parameters(params): Parameters<tools::FindFormulaParams>,
    ) -> Result<Json<FindFormulaResponse>, McpError> {
        self.ensure_tool_enabled("find_formula")
            .map_err(to_mcp_error)?;
        tools::find_formula(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "scan_volatiles", description = "Scan for volatile formulas")]
    pub async fn scan_volatiles(
        &self,
        Parameters(params): Parameters<tools::ScanVolatilesParams>,
    ) -> Result<Json<VolatileScanResponse>, McpError> {
        self.ensure_tool_enabled("scan_volatiles")
            .map_err(to_mcp_error)?;
        tools::scan_volatiles(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "sheet_styles",
        description = "Summarise style usage and properties for a sheet"
    )]
    pub async fn sheet_styles(
        &self,
        Parameters(params): Parameters<tools::SheetStylesParams>,
    ) -> Result<Json<SheetStylesResponse>, McpError> {
        self.ensure_tool_enabled("sheet_styles")
            .map_err(to_mcp_error)?;
        tools::sheet_styles(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "workbook_style_summary",
        description = "Summarise style usage, theme colors, and conditional formats across a workbook"
    )]
    pub async fn workbook_style_summary(
        &self,
        Parameters(params): Parameters<tools::WorkbookStyleSummaryParams>,
    ) -> Result<Json<WorkbookStyleSummaryResponse>, McpError> {
        self.ensure_tool_enabled("workbook_style_summary")
            .map_err(to_mcp_error)?;
        tools::workbook_style_summary(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "get_manifest_stub",
        description = "Generate manifest scaffold for workbook"
    )]
    pub async fn get_manifest_stub(
        &self,
        Parameters(params): Parameters<tools::ManifestStubParams>,
    ) -> Result<Json<ManifestStubResponse>, McpError> {
        self.ensure_tool_enabled("get_manifest_stub")
            .map_err(to_mcp_error)?;
        tools::get_manifest_stub(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "close_workbook", description = "Evict a workbook from cache")]
    pub async fn close_workbook(
        &self,
        Parameters(params): Parameters<tools::CloseWorkbookParams>,
    ) -> Result<Json<CloseWorkbookResponse>, McpError> {
        self.ensure_tool_enabled("close_workbook")
            .map_err(to_mcp_error)?;
        tools::close_workbook(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }
}

#[cfg(feature = "recalc")]
#[tool_router(router = fork_tool_router)]
impl SpreadsheetServer {
    #[tool(
        name = "create_fork",
        description = "Create a temporary editable copy of a workbook for what-if analysis"
    )]
    pub async fn create_fork(
        &self,
        Parameters(params): Parameters<tools::fork::CreateForkParams>,
    ) -> Result<Json<tools::fork::CreateForkResponse>, McpError> {
        self.ensure_recalc_enabled("create_fork")
            .map_err(to_mcp_error)?;
        tools::fork::create_fork(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "edit_batch",
        description = "Apply batch edits (values or formulas) to a fork"
    )]
    pub async fn edit_batch(
        &self,
        Parameters(params): Parameters<tools::fork::EditBatchParams>,
    ) -> Result<Json<tools::fork::EditBatchResponse>, McpError> {
        self.ensure_recalc_enabled("edit_batch")
            .map_err(to_mcp_error)?;
        tools::fork::edit_batch(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "style_batch",
        description = "Apply batch style edits to a fork. Supports targets by range, region_id, or explicit cells. \
Mode: preview or apply (default apply). Op mode: merge (default), set, or clear."
    )]
    pub async fn style_batch(
        &self,
        Parameters(params): Parameters<tools::fork::StyleBatchParams>,
    ) -> Result<Json<tools::fork::StyleBatchResponse>, McpError> {
        self.ensure_recalc_enabled("style_batch")
            .map_err(to_mcp_error)?;
        tools::fork::style_batch(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "apply_formula_pattern",
        description = "Autofill-like formula pattern application over a target range in a fork. \
Provide base_formula at anchor_cell, then fill across target_range. \
Mode: preview or apply (default apply). relative_mode: excel (default), abs_cols, abs_rows. \
fill_direction: down, right, both (default both)."
    )]
    pub async fn apply_formula_pattern(
        &self,
        Parameters(params): Parameters<tools::fork::ApplyFormulaPatternParams>,
    ) -> Result<Json<tools::fork::ApplyFormulaPatternResponse>, McpError> {
        self.ensure_recalc_enabled("apply_formula_pattern")
            .map_err(to_mcp_error)?;
        tools::fork::apply_formula_pattern(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "structure_batch",
        description = "Apply structural edits to a fork (rows/cols/sheets). \
Mode: preview or apply (default apply). Note: structural edits may not fully rewrite formulas/named ranges like Excel; \
run recalculate and review get_changeset after applying."
    )]
    pub async fn structure_batch(
        &self,
        Parameters(params): Parameters<tools::fork::StructureBatchParams>,
    ) -> Result<Json<tools::fork::StructureBatchResponse>, McpError> {
        self.ensure_recalc_enabled("structure_batch")
            .map_err(to_mcp_error)?;
        tools::fork::structure_batch(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "get_edits", description = "List all edits applied to a fork")]
    pub async fn get_edits(
        &self,
        Parameters(params): Parameters<tools::fork::GetEditsParams>,
    ) -> Result<Json<tools::fork::GetEditsResponse>, McpError> {
        self.ensure_recalc_enabled("get_edits")
            .map_err(to_mcp_error)?;
        tools::fork::get_edits(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "get_changeset",
        description = "Calculate diff between fork and base workbook"
    )]
    pub async fn get_changeset(
        &self,
        Parameters(params): Parameters<tools::fork::GetChangesetParams>,
    ) -> Result<Json<tools::fork::GetChangesetResponse>, McpError> {
        self.ensure_recalc_enabled("get_changeset")
            .map_err(to_mcp_error)?;
        tools::fork::get_changeset(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "recalculate",
        description = "Recalculate all formulas in a fork using LibreOffice"
    )]
    pub async fn recalculate(
        &self,
        Parameters(params): Parameters<tools::fork::RecalculateParams>,
    ) -> Result<Json<tools::fork::RecalculateResponse>, McpError> {
        self.ensure_recalc_enabled("recalculate")
            .map_err(to_mcp_error)?;
        tools::fork::recalculate(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "list_forks", description = "List all active forks")]
    pub async fn list_forks(
        &self,
        Parameters(params): Parameters<tools::fork::ListForksParams>,
    ) -> Result<Json<tools::fork::ListForksResponse>, McpError> {
        self.ensure_recalc_enabled("list_forks")
            .map_err(to_mcp_error)?;
        tools::fork::list_forks(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "discard_fork", description = "Discard a fork without saving")]
    pub async fn discard_fork(
        &self,
        Parameters(params): Parameters<tools::fork::DiscardForkParams>,
    ) -> Result<Json<tools::fork::DiscardForkResponse>, McpError> {
        self.ensure_recalc_enabled("discard_fork")
            .map_err(to_mcp_error)?;
        tools::fork::discard_fork(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "save_fork",
        description = "Save fork changes to target path (defaults to overwriting original)"
    )]
    pub async fn save_fork(
        &self,
        Parameters(params): Parameters<tools::fork::SaveForkParams>,
    ) -> Result<Json<tools::fork::SaveForkResponse>, McpError> {
        self.ensure_recalc_enabled("save_fork")
            .map_err(to_mcp_error)?;
        tools::fork::save_fork(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "checkpoint_fork",
        description = "Create a high-fidelity checkpoint snapshot of a fork"
    )]
    pub async fn checkpoint_fork(
        &self,
        Parameters(params): Parameters<tools::fork::CheckpointForkParams>,
    ) -> Result<Json<tools::fork::CheckpointForkResponse>, McpError> {
        self.ensure_recalc_enabled("checkpoint_fork")
            .map_err(to_mcp_error)?;
        tools::fork::checkpoint_fork(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(name = "list_checkpoints", description = "List checkpoints for a fork")]
    pub async fn list_checkpoints(
        &self,
        Parameters(params): Parameters<tools::fork::ListCheckpointsParams>,
    ) -> Result<Json<tools::fork::ListCheckpointsResponse>, McpError> {
        self.ensure_recalc_enabled("list_checkpoints")
            .map_err(to_mcp_error)?;
        tools::fork::list_checkpoints(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "restore_checkpoint",
        description = "Restore a fork to a checkpoint"
    )]
    pub async fn restore_checkpoint(
        &self,
        Parameters(params): Parameters<tools::fork::RestoreCheckpointParams>,
    ) -> Result<Json<tools::fork::RestoreCheckpointResponse>, McpError> {
        self.ensure_recalc_enabled("restore_checkpoint")
            .map_err(to_mcp_error)?;
        tools::fork::restore_checkpoint(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "delete_checkpoint",
        description = "Delete a checkpoint from a fork"
    )]
    pub async fn delete_checkpoint(
        &self,
        Parameters(params): Parameters<tools::fork::DeleteCheckpointParams>,
    ) -> Result<Json<tools::fork::DeleteCheckpointResponse>, McpError> {
        self.ensure_recalc_enabled("delete_checkpoint")
            .map_err(to_mcp_error)?;
        tools::fork::delete_checkpoint(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "list_staged_changes",
        description = "List previewed/staged changes for a fork"
    )]
    pub async fn list_staged_changes(
        &self,
        Parameters(params): Parameters<tools::fork::ListStagedChangesParams>,
    ) -> Result<Json<tools::fork::ListStagedChangesResponse>, McpError> {
        self.ensure_recalc_enabled("list_staged_changes")
            .map_err(to_mcp_error)?;
        tools::fork::list_staged_changes(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "apply_staged_change",
        description = "Apply a staged change to a fork"
    )]
    pub async fn apply_staged_change(
        &self,
        Parameters(params): Parameters<tools::fork::ApplyStagedChangeParams>,
    ) -> Result<Json<tools::fork::ApplyStagedChangeResponse>, McpError> {
        self.ensure_recalc_enabled("apply_staged_change")
            .map_err(to_mcp_error)?;
        tools::fork::apply_staged_change(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "discard_staged_change",
        description = "Discard a staged change without applying it"
    )]
    pub async fn discard_staged_change(
        &self,
        Parameters(params): Parameters<tools::fork::DiscardStagedChangeParams>,
    ) -> Result<Json<tools::fork::DiscardStagedChangeResponse>, McpError> {
        self.ensure_recalc_enabled("discard_staged_change")
            .map_err(to_mcp_error)?;
        tools::fork::discard_staged_change(self.state.clone(), params)
            .await
            .map(Json)
            .map_err(to_mcp_error)
    }

    #[tool(
        name = "screenshot_sheet",
        description = "Capture a visual screenshot of a spreadsheet region as PNG. \
	Returns file URI. Max range: 100 rows x 30 columns. Default: A1:M40."
    )]
    pub async fn screenshot_sheet(
        &self,
        Parameters(params): Parameters<tools::fork::ScreenshotSheetParams>,
    ) -> Result<rmcp::model::CallToolResult, McpError> {
        use base64::Engine;
        use rmcp::model::Content;

        self.ensure_recalc_enabled("screenshot_sheet")
            .map_err(to_mcp_error)?;

        let return_image = params.return_image;
        let response = tools::fork::screenshot_sheet(self.state.clone(), params)
            .await
            .map_err(to_mcp_error)?;

        let mut content = Vec::new();

        if return_image {
            let fs_path = response
                .output_path
                .strip_prefix("file://")
                .ok_or_else(|| {
                    to_mcp_error(anyhow::anyhow!("unexpected screenshot output_path"))
                })?;
            let bytes = tokio::fs::read(fs_path)
                .await
                .map_err(|e| to_mcp_error(anyhow::anyhow!("failed to read screenshot: {}", e)))?;

            let data = base64::engine::general_purpose::STANDARD.encode(bytes);
            content.push(Content::image(data, "image/png"));
        }

        // Always include a small text hint for clients that ignore structured_content.
        content.push(Content::text(response.output_path.clone()));

        let structured_content = serde_json::to_value(&response)
            .map_err(|e| to_mcp_error(anyhow::anyhow!("failed to serialize response: {}", e)))?;

        Ok(rmcp::model::CallToolResult {
            content,
            structured_content: Some(structured_content),
            is_error: Some(false),
            meta: None,
        })
    }
}

#[tool_handler(router = self.tool_router)]
impl ServerHandler for SpreadsheetServer {
    fn get_info(&self) -> ServerInfo {
        let recalc_enabled = {
            #[cfg(feature = "recalc")]
            {
                self.state.config().recalc_enabled
            }
            #[cfg(not(feature = "recalc"))]
            {
                false
            }
        };

        ServerInfo {
            capabilities: ServerCapabilities::builder().enable_tools().build(),
            server_info: Implementation::from_build_env(),
            instructions: Some(build_instructions(recalc_enabled)),
            ..ServerInfo::default()
        }
    }
}

fn to_mcp_error(error: anyhow::Error) -> McpError {
    if error.downcast_ref::<ToolDisabledError>().is_some() {
        McpError::invalid_request(error.to_string(), None)
    } else {
        McpError::internal_error(error.to_string(), None)
    }
}

#[derive(Debug, Error)]
#[error("tool '{tool_name}' is disabled by server configuration")]
struct ToolDisabledError {
    tool_name: String,
}

impl ToolDisabledError {
    fn new(tool_name: &str) -> Self {
        Self {
            tool_name: tool_name.to_ascii_lowercase(),
        }
    }
}

#[cfg(feature = "recalc")]
#[derive(Debug, Error)]
#[error("recalc/write tools are disabled (set SPREADSHEET_MCP_RECALC_ENABLED=true)")]
struct RecalcDisabledError;
