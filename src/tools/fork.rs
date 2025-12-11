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
