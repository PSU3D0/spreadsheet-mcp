use crate::fork::EditOp;
use crate::model::WorkbookId;
use crate::state::AppState;
use anyhow::{Result, anyhow};
use chrono::Utc;
use schemars::JsonSchema;
use serde::{Deserialize, Serialize};
use std::sync::Arc;

#[derive(Debug, Deserialize, JsonSchema)]
pub struct CreateForkParams {
    pub workbook_id: WorkbookId,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct CreateForkResponse {
    pub fork_id: String,
    pub base_workbook: String,
    pub ttl_seconds: u64,
}

pub async fn create_fork(
    state: Arc<AppState>,
    params: CreateForkParams,
) -> Result<CreateForkResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available (recalc disabled?)"))?;

    let workbook = state.open_workbook(&params.workbook_id).await?;
    let base_path = &workbook.path;
    let workspace_root = &state.config().workspace_root;

    let fork_id = registry.create_fork(base_path, workspace_root)?;

    Ok(CreateForkResponse {
        fork_id,
        base_workbook: base_path.display().to_string(),
        ttl_seconds: 3600,
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct EditBatchParams {
    pub fork_id: String,
    pub sheet_name: String,
    pub edits: Vec<CellEdit>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct CellEdit {
    pub address: String,
    pub value: String,
    #[serde(default)]
    pub is_formula: bool,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct EditBatchResponse {
    pub fork_id: String,
    pub edits_applied: usize,
    pub total_edits: usize,
}

pub async fn edit_batch(
    state: Arc<AppState>,
    params: EditBatchParams,
) -> Result<EditBatchResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let fork_ctx = registry.get_fork(&params.fork_id)?;
    let work_path = fork_ctx.work_path.clone();

    let edits_to_apply: Vec<_> = params
        .edits
        .iter()
        .map(|e| EditOp {
            timestamp: Utc::now(),
            sheet: params.sheet_name.clone(),
            address: e.address.clone(),
            value: e.value.clone(),
            is_formula: e.is_formula,
        })
        .collect();

    let edit_count = edits_to_apply.len();

    tokio::task::spawn_blocking({
        let sheet_name = params.sheet_name.clone();
        let edits = params.edits.clone();
        move || apply_edits_to_file(&work_path, &sheet_name, &edits)
    })
    .await??;

    let total = registry.with_fork_mut(&params.fork_id, |ctx| {
        ctx.edits.extend(edits_to_apply);
        Ok(ctx.edits.len())
    })?;

    let fork_workbook_id = WorkbookId(params.fork_id.clone());
    let _ = state.close_workbook(&fork_workbook_id);

    Ok(EditBatchResponse {
        fork_id: params.fork_id,
        edits_applied: edit_count,
        total_edits: total,
    })
}

fn apply_edits_to_file(path: &std::path::Path, sheet_name: &str, edits: &[CellEdit]) -> Result<()> {
    let mut book = umya_spreadsheet::reader::xlsx::read(path)?;

    let sheet = book
        .get_sheet_by_name_mut(sheet_name)
        .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))?;

    for edit in edits {
        let cell = sheet.get_cell_mut(edit.address.as_str());
        if edit.is_formula {
            cell.set_formula(edit.value.clone());
        } else {
            cell.set_value(edit.value.clone());
        }
    }

    umya_spreadsheet::writer::xlsx::write(&book, path)?;
    Ok(())
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct GetEditsParams {
    pub fork_id: String,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct GetEditsResponse {
    pub fork_id: String,
    pub edits: Vec<EditRecord>,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct EditRecord {
    pub timestamp: String,
    pub sheet: String,
    pub address: String,
    pub value: String,
    pub is_formula: bool,
}

pub async fn get_edits(state: Arc<AppState>, params: GetEditsParams) -> Result<GetEditsResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let fork_ctx = registry.get_fork(&params.fork_id)?;

    let edits: Vec<EditRecord> = fork_ctx
        .edits
        .iter()
        .map(|e| EditRecord {
            timestamp: e.timestamp.to_rfc3339(),
            sheet: e.sheet.clone(),
            address: e.address.clone(),
            value: e.value.clone(),
            is_formula: e.is_formula,
        })
        .collect();

    Ok(GetEditsResponse {
        fork_id: params.fork_id,
        edits,
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct GetChangesetParams {
    pub fork_id: String,
    pub sheet_name: Option<String>,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct GetChangesetResponse {
    pub fork_id: String,
    pub base_workbook: String,
    pub changes: Vec<crate::diff::Change>,
}

pub async fn get_changeset(
    state: Arc<AppState>,
    params: GetChangesetParams,
) -> Result<GetChangesetResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let fork_ctx = registry.get_fork(&params.fork_id)?;

    let changes = tokio::task::spawn_blocking({
        let base_path = fork_ctx.base_path.clone();
        let work_path = fork_ctx.work_path.clone();
        let sheet_filter = params.sheet_name.clone();
        move || crate::diff::calculate_changeset(&base_path, &work_path, sheet_filter.as_deref())
    })
    .await??;

    Ok(GetChangesetResponse {
        fork_id: params.fork_id,
        base_workbook: fork_ctx.base_path.display().to_string(),
        changes,
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct RecalculateParams {
    pub fork_id: String,
    #[serde(default = "default_timeout")]
    pub timeout_ms: u64,
}

fn default_timeout() -> u64 {
    30_000
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct RecalculateResponse {
    pub fork_id: String,
    pub duration_ms: u64,
    pub backend: String,
}

pub async fn recalculate(
    state: Arc<AppState>,
    params: RecalculateParams,
) -> Result<RecalculateResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let backend = state
        .recalc_backend()
        .ok_or_else(|| anyhow!("recalc backend not available (soffice not found?)"))?;

    let semaphore = state
        .recalc_semaphore()
        .ok_or_else(|| anyhow!("recalc semaphore not available"))?;

    let fork_ctx = registry.get_fork(&params.fork_id)?;

    let _permit = semaphore
        .0
        .acquire()
        .await
        .map_err(|e| anyhow!("failed to acquire recalc permit: {}", e))?;

    let result = backend.recalculate(&fork_ctx.work_path).await?;

    let fork_workbook_id = WorkbookId(params.fork_id.clone());
    let _ = state.close_workbook(&fork_workbook_id);

    Ok(RecalculateResponse {
        fork_id: params.fork_id,
        duration_ms: result.duration_ms,
        backend: result.executor_type.to_string(),
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ListForksParams {}

#[derive(Debug, Serialize, JsonSchema)]
pub struct ListForksResponse {
    pub forks: Vec<ForkSummary>,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct ForkSummary {
    pub fork_id: String,
    pub base_path: String,
    pub age_seconds: u64,
    pub edit_count: usize,
}

pub async fn list_forks(
    state: Arc<AppState>,
    _params: ListForksParams,
) -> Result<ListForksResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let forks: Vec<ForkSummary> = registry
        .list_forks()
        .into_iter()
        .map(|f| ForkSummary {
            fork_id: f.fork_id,
            base_path: f.base_path,
            age_seconds: f.created_at.elapsed().as_secs(),
            edit_count: f.edit_count,
        })
        .collect();

    Ok(ListForksResponse { forks })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DiscardForkParams {
    pub fork_id: String,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct DiscardForkResponse {
    pub fork_id: String,
    pub discarded: bool,
}

pub async fn discard_fork(
    state: Arc<AppState>,
    params: DiscardForkParams,
) -> Result<DiscardForkResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    registry.discard_fork(&params.fork_id)?;

    Ok(DiscardForkResponse {
        fork_id: params.fork_id,
        discarded: true,
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct SaveForkParams {
    pub fork_id: String,
    /// Target path to save to. If omitted, saves to original location (requires --allow-overwrite).
    pub target_path: Option<String>,
    /// If true, discard the fork after saving. If false, fork remains active for further edits.
    #[serde(default = "default_drop_fork")]
    pub drop_fork: bool,
}

fn default_drop_fork() -> bool {
    true
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct SaveForkResponse {
    pub fork_id: String,
    pub saved_to: String,
    pub fork_dropped: bool,
}

pub async fn save_fork(state: Arc<AppState>, params: SaveForkParams) -> Result<SaveForkResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let fork_ctx = registry.get_fork(&params.fork_id)?;
    let config = state.config();
    let workspace_root = &config.workspace_root;

    let (target, is_overwrite) = match params.target_path {
        Some(p) => {
            let resolved = config.resolve_path(&p);
            let is_overwrite = resolved == fork_ctx.base_path;
            (resolved, is_overwrite)
        }
        None => (fork_ctx.base_path.clone(), true),
    };

    if is_overwrite && !config.allow_overwrite {
        return Err(anyhow!(
            "overwriting original file is disabled. Use --allow-overwrite flag or specify a different target_path"
        ));
    }

    let base_path = fork_ctx.base_path.clone();
    registry.save_fork(&params.fork_id, &target, workspace_root, params.drop_fork)?;

    if is_overwrite {
        state.evict_by_path(&base_path);
    }

    Ok(SaveForkResponse {
        fork_id: params.fork_id,
        saved_to: target.display().to_string(),
        fork_dropped: params.drop_fork,
    })
}

const MAX_SCREENSHOT_ROWS: u32 = 100;
const MAX_SCREENSHOT_COLS: u32 = 30;
const DEFAULT_SCREENSHOT_RANGE: &str = "A1:M40";
const DEFAULT_MAX_PNG_DIM_PX: u32 = 4096;
const DEFAULT_MAX_PNG_AREA_PX: u64 = 12_000_000;

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ScreenshotSheetParams {
    pub workbook_id: WorkbookId,
    pub sheet_name: String,
    #[serde(default)]
    pub range: Option<String>,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct ScreenshotSheetResponse {
    pub workbook_id: String,
    pub sheet_name: String,
    pub range: String,
    pub output_path: String,
    pub size_bytes: u64,
    pub duration_ms: u64,
}

pub async fn screenshot_sheet(
    state: Arc<AppState>,
    params: ScreenshotSheetParams,
) -> Result<ScreenshotSheetResponse> {
    let range = params.range.as_deref().unwrap_or(DEFAULT_SCREENSHOT_RANGE);
    let bounds = validate_screenshot_range(range)?;

    let workbook = state.open_workbook(&params.workbook_id).await?;
    let workbook_path = workbook.path.clone();

    let _ = workbook.with_sheet(&params.sheet_name, |_| Ok::<_, anyhow::Error>(()))?;

    let safe_range = range.replace(':', "-");
    let filename = format!(
        "{}_{}_{}.png",
        workbook.slug,
        params.sheet_name.replace(' ', "_"),
        safe_range
    );

    let screenshot_dir = state.config().workspace_root.join("screenshots");
    tokio::fs::create_dir_all(&screenshot_dir).await?;
    let output_path = screenshot_dir.join(&filename);

    let executor = crate::recalc::ScreenshotExecutor::new(&crate::recalc::RecalcConfig::default());
    let result = executor
        .screenshot(
            &workbook_path,
            &output_path,
            &params.sheet_name,
            Some(range),
        )
        .await?;

    enforce_png_pixel_limits(&result.output_path, range, &bounds).await?;

    Ok(ScreenshotSheetResponse {
        workbook_id: params.workbook_id.0,
        sheet_name: params.sheet_name,
        range: range.to_string(),
        output_path: format!("file://{}", result.output_path.display()),
        size_bytes: result.size_bytes,
        duration_ms: result.duration_ms,
    })
}

#[derive(Debug, Clone, Copy)]
struct ScreenshotBounds {
    min_col: u32,
    max_col: u32,
    min_row: u32,
    max_row: u32,
    rows: u32,
    cols: u32,
}

fn validate_screenshot_range(range: &str) -> Result<ScreenshotBounds> {
    let bounds = parse_range_bounds(range)?;

    if bounds.rows > MAX_SCREENSHOT_ROWS || bounds.cols > MAX_SCREENSHOT_COLS {
        let row_tiles = div_ceil(bounds.rows, MAX_SCREENSHOT_ROWS);
        let col_tiles = div_ceil(bounds.cols, MAX_SCREENSHOT_COLS);
        let total_tiles = row_tiles * col_tiles;

        let display_limit = 50usize;
        let display_ranges = suggest_tiled_ranges(
            &bounds,
            MAX_SCREENSHOT_ROWS,
            MAX_SCREENSHOT_COLS,
            Some(display_limit),
        );

        let mut msg = format!(
            "Requested range {range} is too large for a single screenshot ({} rows x {} cols; max {} x {}). \
Split into {} tile(s) ({} row tiles x {} col tiles). Suggested ranges: {}",
            bounds.rows,
            bounds.cols,
            MAX_SCREENSHOT_ROWS,
            MAX_SCREENSHOT_COLS,
            total_tiles,
            row_tiles,
            col_tiles,
            display_ranges.join(", ")
        );
        if total_tiles as usize > display_limit {
            msg.push_str(&format!(
                " ... and {} more.",
                total_tiles as usize - display_limit
            ));
        }
        return Err(anyhow!(msg));
    }

    Ok(bounds)
}

fn parse_cell_ref(cell: &str) -> Result<(u32, u32)> {
    use umya_spreadsheet::helper::coordinate::index_from_coordinate;
    let (col, row, _, _) = index_from_coordinate(cell);
    match (col, row) {
        (Some(c), Some(r)) => Ok((c, r)),
        _ => Err(anyhow!("Invalid cell reference: {}", cell)),
    }
}

fn parse_range_bounds(range: &str) -> Result<ScreenshotBounds> {
    let parts: Vec<&str> = range.split(':').collect();
    if parts.len() != 2 {
        return Err(anyhow!("Invalid range format. Expected 'A1:Z99'"));
    }

    let start = parse_cell_ref(parts[0])?;
    let end = parse_cell_ref(parts[1])?;

    let min_col = start.0.min(end.0);
    let max_col = start.0.max(end.0);
    let min_row = start.1.min(end.1);
    let max_row = start.1.max(end.1);

    let rows = max_row - min_row + 1;
    let cols = max_col - min_col + 1;

    Ok(ScreenshotBounds {
        min_col,
        max_col,
        min_row,
        max_row,
        rows,
        cols,
    })
}

fn div_ceil(n: u32, d: u32) -> u32 {
    n.div_ceil(d)
}

fn suggest_tiled_ranges(
    bounds: &ScreenshotBounds,
    max_rows: u32,
    max_cols: u32,
    limit: Option<usize>,
) -> Vec<String> {
    use umya_spreadsheet::helper::coordinate::coordinate_from_index;

    let mut out = Vec::new();
    let mut row_start = bounds.min_row;
    while row_start <= bounds.max_row {
        let row_end = (row_start + max_rows - 1).min(bounds.max_row);
        let mut col_start = bounds.min_col;
        while col_start <= bounds.max_col {
            let col_end = (col_start + max_cols - 1).min(bounds.max_col);
            let start_cell = coordinate_from_index(&col_start, &row_start);
            let end_cell = coordinate_from_index(&col_end, &row_end);
            out.push(format!("{start_cell}:{end_cell}"));
            if let Some(lim) = limit
                && out.len() >= lim
            {
                return out;
            }
            col_start = col_end + 1;
        }
        row_start = row_end + 1;
        if let Some(lim) = limit
            && out.len() >= lim
        {
            return out;
        }
    }
    out
}

fn suggest_split_single_tile(bounds: &ScreenshotBounds) -> Vec<String> {
    use umya_spreadsheet::helper::coordinate::coordinate_from_index;

    if bounds.rows >= bounds.cols && bounds.rows > 1 {
        let mid_row = bounds.min_row + (bounds.rows / 2) - 1;
        let start1 = coordinate_from_index(&bounds.min_col, &bounds.min_row);
        let end1 = coordinate_from_index(&bounds.max_col, &mid_row);
        let start2 = coordinate_from_index(&bounds.min_col, &(mid_row + 1));
        let end2 = coordinate_from_index(&bounds.max_col, &bounds.max_row);
        vec![format!("{start1}:{end1}"), format!("{start2}:{end2}")]
    } else if bounds.cols > 1 {
        let mid_col = bounds.min_col + (bounds.cols / 2) - 1;
        let start1 = coordinate_from_index(&bounds.min_col, &bounds.min_row);
        let end1 = coordinate_from_index(&mid_col, &bounds.max_row);
        let start2 = coordinate_from_index(&(mid_col + 1), &bounds.min_row);
        let end2 = coordinate_from_index(&bounds.max_col, &bounds.max_row);
        vec![format!("{start1}:{end1}"), format!("{start2}:{end2}")]
    } else {
        vec![range_from_bounds(bounds)]
    }
}

fn range_from_bounds(bounds: &ScreenshotBounds) -> String {
    use umya_spreadsheet::helper::coordinate::coordinate_from_index;
    let start = coordinate_from_index(&bounds.min_col, &bounds.min_row);
    let end = coordinate_from_index(&bounds.max_col, &bounds.max_row);
    format!("{start}:{end}")
}

async fn enforce_png_pixel_limits(
    path: &std::path::Path,
    range: &str,
    bounds: &ScreenshotBounds,
) -> Result<()> {
    use image::GenericImageView;
    use image::ImageReader;

    let max_dim_px = std::env::var("SPREADSHEET_MCP_MAX_PNG_DIM_PX")
        .ok()
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(DEFAULT_MAX_PNG_DIM_PX);
    let max_area_px = std::env::var("SPREADSHEET_MCP_MAX_PNG_AREA_PX")
        .ok()
        .and_then(|v| v.parse::<u64>().ok())
        .unwrap_or(DEFAULT_MAX_PNG_AREA_PX);

    let img = ImageReader::open(path)
        .and_then(|r| r.with_guessed_format())
        .map_err(|e| anyhow!("failed to read png {}: {}", path.display(), e))?
        .decode()
        .map_err(|e| anyhow!("failed to decode png {}: {}", path.display(), e))?;
    let (w, h) = img.dimensions();
    let area = (w as u64) * (h as u64);

    if w > max_dim_px || h > max_dim_px || area > max_area_px {
        let _ = tokio::fs::remove_file(path).await;

        let mut suggestions =
            suggest_tiled_ranges(bounds, MAX_SCREENSHOT_ROWS, MAX_SCREENSHOT_COLS, Some(50));
        let row_tiles = div_ceil(bounds.rows, MAX_SCREENSHOT_ROWS);
        let col_tiles = div_ceil(bounds.cols, MAX_SCREENSHOT_COLS);
        let total_tiles = row_tiles * col_tiles;
        if total_tiles == 1 {
            suggestions = suggest_split_single_tile(bounds);
        }

        return Err(anyhow!(
            "Rendered PNG for range {range} is {w}x{h}px (area {area}px), exceeding limits (max_dim={max_dim_px}px, max_area={max_area_px}px). \
Try smaller ranges. Suggested ranges: {}",
            suggestions.join(", ")
        ));
    }

    Ok(())
}
