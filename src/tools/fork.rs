use crate::fork::{ChangeSummary, EditOp, StagedChange, StagedOp};
use crate::formula::pattern::{RelativeMode, parse_base_formula, shift_formula_ast};
use crate::model::{StylePatch, WorkbookId};
use crate::state::AppState;
use anyhow::{Result, anyhow, bail};
use chrono::Utc;
use formualizer_parse::tokenizer::Tokenizer;
use schemars::JsonSchema;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use uuid::Uuid;

#[derive(Debug, Deserialize, JsonSchema)]
pub struct CreateForkParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
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

    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
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

fn default_clear_values() -> bool {
    true
}

fn default_overwrite_formulas() -> bool {
    false
}

fn default_replace_case_sensitive() -> bool {
    true
}

fn default_replace_match_mode() -> String {
    "exact".to_string()
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct TransformBatchParams {
    pub fork_id: String,
    pub ops: Vec<TransformOp>,
    #[serde(default)]
    pub mode: Option<String>, // preview|apply (default apply)
    pub label: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum TransformOp {
    ClearRange {
        sheet_name: String,
        target: TransformTarget,
        #[serde(default = "default_clear_values")]
        clear_values: bool,
        #[serde(default)]
        clear_formulas: bool,
    },
    FillRange {
        sheet_name: String,
        target: TransformTarget,
        value: String,
        #[serde(default)]
        is_formula: bool,
        #[serde(default = "default_overwrite_formulas")]
        overwrite_formulas: bool,
    },
    ReplaceInRange {
        sheet_name: String,
        target: TransformTarget,
        find: String,
        replace: String,
        #[serde(default = "default_replace_match_mode")]
        match_mode: String, // exact|contains
        #[serde(default = "default_replace_case_sensitive")]
        case_sensitive: bool,
        #[serde(default)]
        include_formulas: bool,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum TransformTarget {
    Range { range: String },
    Region { region_id: u32 },
    Cells { cells: Vec<String> },
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct TransformBatchResponse {
    pub fork_id: String,
    pub mode: String,
    pub change_id: Option<String>,
    pub ops_applied: usize,
    pub summary: ChangeSummary,
}

#[derive(Debug, Serialize, Deserialize)]
struct TransformBatchStagedPayload {
    ops: Vec<TransformOp>,
}

pub async fn transform_batch(
    state: Arc<AppState>,
    params: TransformBatchParams,
) -> Result<TransformBatchResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let fork_ctx = registry.get_fork(&params.fork_id)?;
    let work_path = fork_ctx.work_path.clone();

    let fork_workbook_id = WorkbookId(params.fork_id.clone());
    let workbook = state.open_workbook(&fork_workbook_id).await?;

    let mut resolved_ops = Vec::with_capacity(params.ops.len());
    for op in &params.ops {
        let (sheet_name, target) = match op {
            TransformOp::ClearRange {
                sheet_name, target, ..
            }
            | TransformOp::FillRange {
                sheet_name, target, ..
            }
            | TransformOp::ReplaceInRange {
                sheet_name, target, ..
            } => (sheet_name, target),
        };

        let resolved_target = match target {
            TransformTarget::Region { region_id } => {
                let metrics = workbook.get_sheet_metrics(sheet_name)?;
                let region = metrics
                    .detected_regions
                    .iter()
                    .find(|r| r.id == *region_id)
                    .ok_or_else(|| {
                        anyhow!(
                            "region_id {} not found on sheet '{}'",
                            region_id,
                            sheet_name
                        )
                    })?;
                TransformTarget::Range {
                    range: region.bounds.clone(),
                }
            }
            other => other.clone(),
        };

        match op {
            TransformOp::ClearRange {
                sheet_name,
                clear_values,
                clear_formulas,
                ..
            } => {
                resolved_ops.push(TransformOp::ClearRange {
                    sheet_name: sheet_name.clone(),
                    target: resolved_target,
                    clear_values: *clear_values,
                    clear_formulas: *clear_formulas,
                });
            }
            TransformOp::FillRange {
                sheet_name,
                value,
                is_formula,
                overwrite_formulas,
                ..
            } => {
                resolved_ops.push(TransformOp::FillRange {
                    sheet_name: sheet_name.clone(),
                    target: resolved_target,
                    value: value.clone(),
                    is_formula: *is_formula,
                    overwrite_formulas: *overwrite_formulas,
                });
            }
            TransformOp::ReplaceInRange {
                sheet_name,
                find,
                replace,
                match_mode,
                case_sensitive,
                include_formulas,
                ..
            } => {
                resolved_ops.push(TransformOp::ReplaceInRange {
                    sheet_name: sheet_name.clone(),
                    target: resolved_target,
                    find: find.clone(),
                    replace: replace.clone(),
                    match_mode: match_mode.clone(),
                    case_sensitive: *case_sensitive,
                    include_formulas: *include_formulas,
                });
            }
        }
    }

    let mode = params
        .mode
        .as_deref()
        .unwrap_or("apply")
        .to_ascii_lowercase();

    if mode == "preview" {
        let change_id = Uuid::new_v4().to_string();
        let snapshot_path = stage_snapshot_path(&params.fork_id, &change_id);
        fs::create_dir_all(snapshot_path.parent().unwrap())?;
        fs::copy(&work_path, &snapshot_path)?;

        let snapshot_for_apply = snapshot_path.clone();
        let apply_result = tokio::task::spawn_blocking({
            let ops = resolved_ops.clone();
            move || apply_transform_ops_to_file(&snapshot_for_apply, &ops)
        })
        .await??;

        let mut summary = apply_result.summary;
        summary.op_kinds = vec!["transform_batch".to_string()];

        let staged_op = StagedOp {
            kind: "transform_batch".to_string(),
            payload: serde_json::to_value(TransformBatchStagedPayload {
                ops: resolved_ops.clone(),
            })?,
        };

        let staged = StagedChange {
            change_id: change_id.clone(),
            created_at: Utc::now(),
            label: params.label.clone(),
            ops: vec![staged_op],
            summary: summary.clone(),
            fork_path_snapshot: Some(snapshot_path),
        };

        registry.add_staged_change(&params.fork_id, staged)?;

        Ok(TransformBatchResponse {
            fork_id: params.fork_id,
            mode,
            change_id: Some(change_id),
            ops_applied: apply_result.ops_applied,
            summary,
        })
    } else {
        let apply_result = tokio::task::spawn_blocking({
            let ops = resolved_ops.clone();
            let work_path = work_path.clone();
            move || apply_transform_ops_to_file(&work_path, &ops)
        })
        .await??;

        let mut summary = apply_result.summary;
        summary.op_kinds = vec!["transform_batch".to_string()];

        let _ = state.close_workbook(&fork_workbook_id);

        Ok(TransformBatchResponse {
            fork_id: params.fork_id,
            mode,
            change_id: None,
            ops_applied: apply_result.ops_applied,
            summary,
        })
    }
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct StyleBatchParams {
    pub fork_id: String,
    pub ops: Vec<StyleOp>,
    #[serde(default)]
    pub mode: Option<String>, // "preview" | "apply" (default apply)
    pub label: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct StyleOp {
    pub sheet_name: String,
    pub target: StyleTarget,
    pub patch: StylePatch,
    #[serde(default)]
    pub op_mode: Option<String>, // "merge" | "set" | "clear"
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum StyleTarget {
    Range { range: String },
    Region { region_id: u32 },
    Cells { cells: Vec<String> },
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct StyleBatchResponse {
    pub fork_id: String,
    pub mode: String,
    pub change_id: Option<String>,
    pub ops_applied: usize,
    pub summary: ChangeSummary,
}

#[derive(Debug, Serialize, Deserialize)]
struct StyleBatchStagedPayload {
    ops: Vec<StyleOp>,
}

pub async fn style_batch(
    state: Arc<AppState>,
    params: StyleBatchParams,
) -> Result<StyleBatchResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let fork_ctx = registry.get_fork(&params.fork_id)?;
    let work_path = fork_ctx.work_path.clone();

    // Resolve any region targets against current fork regions.
    let fork_workbook_id = WorkbookId(params.fork_id.clone());
    let workbook = state.open_workbook(&fork_workbook_id).await?;
    let mut resolved_ops = Vec::with_capacity(params.ops.len());
    for op in &params.ops {
        let mut resolved = op.clone();
        if let StyleTarget::Region { region_id } = &op.target {
            let metrics = workbook.get_sheet_metrics(&op.sheet_name)?;
            let region = metrics
                .detected_regions
                .iter()
                .find(|r| r.id == *region_id)
                .ok_or_else(|| {
                    anyhow!(
                        "region_id {} not found on sheet '{}'",
                        region_id,
                        op.sheet_name
                    )
                })?;
            resolved.target = StyleTarget::Range {
                range: region.bounds.clone(),
            };
        }
        resolved_ops.push(resolved);
    }

    let mode = params
        .mode
        .as_deref()
        .unwrap_or("apply")
        .to_ascii_lowercase();

    if mode == "preview" {
        let change_id = Uuid::new_v4().to_string();
        let snapshot_path = stage_snapshot_path(&params.fork_id, &change_id);
        fs::create_dir_all(snapshot_path.parent().unwrap())?;
        fs::copy(&work_path, &snapshot_path)?;

        let snapshot_path_for_apply = snapshot_path.clone();
        let apply_result = tokio::task::spawn_blocking({
            let ops = resolved_ops.clone();
            move || apply_style_ops_to_file(&snapshot_path_for_apply, &ops)
        })
        .await??;

        let mut summary = apply_result.summary;
        summary.op_kinds = vec!["style_batch".to_string()];

        let staged_op = StagedOp {
            kind: "style_batch".to_string(),
            payload: serde_json::to_value(StyleBatchStagedPayload {
                ops: resolved_ops.clone(),
            })?,
        };

        let staged = StagedChange {
            change_id: change_id.clone(),
            created_at: Utc::now(),
            label: params.label.clone(),
            ops: vec![staged_op],
            summary: summary.clone(),
            fork_path_snapshot: Some(snapshot_path),
        };

        registry.add_staged_change(&params.fork_id, staged)?;

        Ok(StyleBatchResponse {
            fork_id: params.fork_id,
            mode,
            change_id: Some(change_id),
            ops_applied: resolved_ops.len(),
            summary,
        })
    } else {
        let apply_result = tokio::task::spawn_blocking({
            let ops = resolved_ops.clone();
            let work_path = work_path.clone();
            move || apply_style_ops_to_file(&work_path, &ops)
        })
        .await??;

        let mut summary = apply_result.summary;
        summary.op_kinds = vec!["style_batch".to_string()];

        let _ = state.close_workbook(&fork_workbook_id);

        Ok(StyleBatchResponse {
            fork_id: params.fork_id,
            mode,
            change_id: None,
            ops_applied: apply_result.ops_applied,
            summary,
        })
    }
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ApplyFormulaPatternParams {
    pub fork_id: String,
    pub sheet_name: String,
    pub target_range: String,
    pub anchor_cell: String,
    pub base_formula: String,
    #[serde(default)]
    pub fill_direction: Option<String>, // down|right|both (default both)
    #[serde(default)]
    pub relative_mode: Option<String>, // excel|abs_cols|abs_rows
    #[serde(default)]
    pub mode: Option<String>, // preview|apply (default apply)
    pub label: Option<String>,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct ApplyFormulaPatternResponse {
    pub fork_id: String,
    pub sheet_name: String,
    pub target_range: String,
    pub mode: String,
    pub change_id: Option<String>,
    pub cells_filled: u64,
    pub summary: ChangeSummary,
}

#[derive(Debug, Serialize, Deserialize)]
struct ApplyFormulaPatternStagedPayload {
    sheet_name: String,
    target_range: String,
    anchor_cell: String,
    base_formula: String,
    fill_direction: Option<String>,
    relative_mode: Option<String>,
}

pub async fn apply_formula_pattern(
    state: Arc<AppState>,
    params: ApplyFormulaPatternParams,
) -> Result<ApplyFormulaPatternResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let fork_ctx = registry.get_fork(&params.fork_id)?;
    let work_path = fork_ctx.work_path.clone();

    let bounds = parse_range_bounds(&params.target_range)?;
    let (anchor_col, anchor_row) = parse_cell_ref(&params.anchor_cell)?;
    validate_formula_pattern_bounds(
        &bounds,
        anchor_col,
        anchor_row,
        params.fill_direction.as_deref(),
    )?;

    let fork_workbook_id = WorkbookId(params.fork_id.clone());
    let workbook = state.open_workbook(&fork_workbook_id).await?;
    let _ = workbook.with_sheet(&params.sheet_name, |_| Ok::<_, anyhow::Error>(()))?;

    let relative_mode = RelativeMode::parse(params.relative_mode.as_deref())?;
    let mode = params
        .mode
        .as_deref()
        .unwrap_or("apply")
        .to_ascii_lowercase();

    if mode == "preview" {
        let change_id = Uuid::new_v4().to_string();
        let snapshot_path = stage_snapshot_path(&params.fork_id, &change_id);
        fs::create_dir_all(snapshot_path.parent().unwrap())?;
        fs::copy(&work_path, &snapshot_path)?;

        let sheet_name = params.sheet_name.clone();
        let target_range = params.target_range.clone();
        let anchor_cell = params.anchor_cell.clone();
        let base_formula = params.base_formula.clone();
        let fill_direction = params.fill_direction.clone();
        let relative_mode_str = params.relative_mode.clone();
        let snapshot_for_apply = snapshot_path.clone();
        let sheet_name_for_apply = sheet_name.clone();
        let target_range_for_apply = target_range.clone();
        let base_formula_for_apply = base_formula.clone();

        let apply_result = tokio::task::spawn_blocking(move || {
            apply_formula_pattern_to_file(
                &snapshot_for_apply,
                &sheet_name_for_apply,
                &target_range_for_apply,
                anchor_col,
                anchor_row,
                &base_formula_for_apply,
                relative_mode,
            )
        })
        .await??;

        let mut summary = apply_result.summary;
        summary.op_kinds = vec!["apply_formula_pattern".to_string()];

        let staged_op = StagedOp {
            kind: "apply_formula_pattern".to_string(),
            payload: serde_json::to_value(ApplyFormulaPatternStagedPayload {
                sheet_name: sheet_name.clone(),
                target_range: target_range.clone(),
                anchor_cell: anchor_cell.clone(),
                base_formula: base_formula.clone(),
                fill_direction,
                relative_mode: relative_mode_str,
            })?,
        };

        let staged = StagedChange {
            change_id: change_id.clone(),
            created_at: Utc::now(),
            label: params.label.clone(),
            ops: vec![staged_op],
            summary: summary.clone(),
            fork_path_snapshot: Some(snapshot_path),
        };

        registry.add_staged_change(&params.fork_id, staged)?;

        Ok(ApplyFormulaPatternResponse {
            fork_id: params.fork_id,
            sheet_name,
            target_range,
            mode,
            change_id: Some(change_id),
            cells_filled: apply_result.cells_filled,
            summary,
        })
    } else {
        let sheet_name = params.sheet_name.clone();
        let target_range = params.target_range.clone();
        let base_formula = params.base_formula.clone();
        let sheet_name_for_apply = sheet_name.clone();
        let target_range_for_apply = target_range.clone();
        let base_formula_for_apply = base_formula.clone();
        let apply_result = tokio::task::spawn_blocking(move || {
            apply_formula_pattern_to_file(
                &work_path,
                &sheet_name_for_apply,
                &target_range_for_apply,
                anchor_col,
                anchor_row,
                &base_formula_for_apply,
                relative_mode,
            )
        })
        .await??;

        let mut summary = apply_result.summary;
        summary.op_kinds = vec!["apply_formula_pattern".to_string()];

        let _ = state.close_workbook(&fork_workbook_id);

        Ok(ApplyFormulaPatternResponse {
            fork_id: params.fork_id,
            sheet_name,
            target_range,
            mode,
            change_id: None,
            cells_filled: apply_result.cells_filled,
            summary,
        })
    }
}

struct FormulaPatternApplyResult {
    cells_filled: u64,
    summary: ChangeSummary,
}

fn apply_formula_pattern_to_file(
    path: &Path,
    sheet_name: &str,
    target_range: &str,
    anchor_col: u32,
    anchor_row: u32,
    base_formula: &str,
    relative_mode: RelativeMode,
) -> Result<FormulaPatternApplyResult> {
    let ast = parse_base_formula(base_formula)?;
    let bounds = parse_range_bounds(target_range)?;

    let mut book = umya_spreadsheet::reader::xlsx::read(path)?;
    let sheet = book
        .get_sheet_by_name_mut(sheet_name)
        .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))?;

    let mut cells_filled: u64 = 0;
    for row in bounds.min_row..=bounds.max_row {
        for col in bounds.min_col..=bounds.max_col {
            let delta_col = col as i32 - anchor_col as i32;
            let delta_row = row as i32 - anchor_row as i32;
            let shifted = shift_formula_ast(&ast, delta_col, delta_row, relative_mode)?;
            let shifted_for_umya = shifted.strip_prefix('=').unwrap_or(&shifted);
            let addr = crate::utils::cell_address(col, row);
            sheet
                .get_cell_mut(addr.as_str())
                .set_formula(shifted_for_umya.to_string());
            cells_filled += 1;
        }
    }

    umya_spreadsheet::writer::xlsx::write(&book, path)?;

    let mut counts = BTreeMap::new();
    counts.insert("cells_filled".to_string(), cells_filled);

    let summary = ChangeSummary {
        op_kinds: vec!["apply_formula_pattern".to_string()],
        affected_sheets: vec![sheet_name.to_string()],
        affected_bounds: vec![target_range.to_string()],
        counts,
        warnings: Vec::new(),
    };

    Ok(FormulaPatternApplyResult {
        cells_filled,
        summary,
    })
}

fn validate_formula_pattern_bounds(
    bounds: &ScreenshotBounds,
    anchor_col: u32,
    anchor_row: u32,
    fill_direction: Option<&str>,
) -> Result<()> {
    if anchor_col < bounds.min_col
        || anchor_col > bounds.max_col
        || anchor_row < bounds.min_row
        || anchor_row > bounds.max_row
    {
        let bounds_range = format!(
            "{}:{}",
            crate::utils::cell_address(bounds.min_col, bounds.min_row),
            crate::utils::cell_address(bounds.max_col, bounds.max_row)
        );
        bail!(
            "anchor_cell must be inside target_range (anchor {} not within {})",
            crate::utils::cell_address(anchor_col, anchor_row),
            bounds_range
        );
    }

    if bounds.min_col != anchor_col || bounds.min_row != anchor_row {
        bail!("target_range must start at anchor_cell (anchor should be top-left of fill range)");
    }

    let dir = fill_direction.unwrap_or("both").to_ascii_lowercase();
    match dir.as_str() {
        "down" => {
            if bounds.min_col != bounds.max_col {
                bail!("fill_direction=down requires a single-column target_range");
            }
        }
        "right" => {
            if bounds.min_row != bounds.max_row {
                bail!("fill_direction=right requires a single-row target_range");
            }
        }
        "both" => {}
        other => bail!("invalid fill_direction: {}", other),
    }
    Ok(())
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct StructureBatchParams {
    pub fork_id: String,
    pub ops: Vec<StructureOp>,
    #[serde(default)]
    pub mode: Option<String>, // preview|apply (default apply)
    pub label: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum StructureOp {
    InsertRows {
        sheet_name: String,
        at_row: u32,
        count: u32,
    },
    DeleteRows {
        sheet_name: String,
        start_row: u32,
        count: u32,
    },
    InsertCols {
        sheet_name: String,
        at_col: String,
        count: u32,
    },
    DeleteCols {
        sheet_name: String,
        start_col: String,
        count: u32,
    },
    RenameSheet {
        old_name: String,
        new_name: String,
    },
    CreateSheet {
        name: String,
        #[serde(default)]
        position: Option<u32>,
    },
    DeleteSheet {
        name: String,
    },
    CopyRange {
        sheet_name: String,
        #[serde(default)]
        dest_sheet_name: Option<String>,
        src_range: String,
        dest_anchor: String,
        include_styles: bool,
        include_formulas: bool,
    },
    MoveRange {
        sheet_name: String,
        #[serde(default)]
        dest_sheet_name: Option<String>,
        src_range: String,
        dest_anchor: String,
        include_styles: bool,
        include_formulas: bool,
    },
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct StructureBatchResponse {
    pub fork_id: String,
    pub mode: String,
    pub change_id: Option<String>,
    pub ops_applied: usize,
    pub summary: ChangeSummary,
}

#[derive(Debug, Serialize, Deserialize)]
struct StructureBatchStagedPayload {
    ops: Vec<StructureOp>,
}

pub async fn structure_batch(
    state: Arc<AppState>,
    params: StructureBatchParams,
) -> Result<StructureBatchResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let fork_ctx = registry.get_fork(&params.fork_id)?;
    let work_path = fork_ctx.work_path.clone();

    let mode = params
        .mode
        .as_deref()
        .unwrap_or("apply")
        .to_ascii_lowercase();

    if mode == "preview" {
        let change_id = Uuid::new_v4().to_string();
        let snapshot_path = stage_snapshot_path(&params.fork_id, &change_id);
        fs::create_dir_all(snapshot_path.parent().unwrap())?;
        fs::copy(&work_path, &snapshot_path)?;

        let snapshot_for_apply = snapshot_path.clone();
        let ops_for_apply = params.ops.clone();

        let apply_result = tokio::task::spawn_blocking(move || {
            apply_structure_ops_to_file(&snapshot_for_apply, &ops_for_apply)
        })
        .await??;

        let mut summary = apply_result.summary;
        summary.op_kinds = vec!["structure_batch".to_string()];
        // Best-effort preview diff size: compare current fork to preview snapshot.
        // This is intentionally summarized as a count to avoid large payloads.
        if let Ok(change_count) = tokio::task::spawn_blocking({
            let base_path = work_path.clone();
            let preview_path = snapshot_path.clone();
            move || {
                crate::diff::calculate_changeset(&base_path, &preview_path, None)
                    .map(|changes| changes.len() as u64)
            }
        })
        .await?
        {
            summary
                .counts
                .insert("preview_change_items".to_string(), change_count);
        } else {
            summary.warnings.push(
                "Preview diff computation failed; run get_changeset after applying to inspect changes."
                    .to_string(),
            );
        }

        let staged_op = StagedOp {
            kind: "structure_batch".to_string(),
            payload: serde_json::to_value(StructureBatchStagedPayload {
                ops: params.ops.clone(),
            })?,
        };

        let staged = StagedChange {
            change_id: change_id.clone(),
            created_at: Utc::now(),
            label: params.label.clone(),
            ops: vec![staged_op],
            summary: summary.clone(),
            fork_path_snapshot: Some(snapshot_path),
        };

        registry.add_staged_change(&params.fork_id, staged)?;

        Ok(StructureBatchResponse {
            fork_id: params.fork_id,
            mode,
            change_id: Some(change_id),
            ops_applied: apply_result.ops_applied,
            summary,
        })
    } else {
        let ops_for_apply = params.ops.clone();
        let apply_result = tokio::task::spawn_blocking(move || {
            apply_structure_ops_to_file(&work_path, &ops_for_apply)
        })
        .await??;

        let mut summary = apply_result.summary;
        summary.op_kinds = vec!["structure_batch".to_string()];

        let fork_workbook_id = WorkbookId(params.fork_id.clone());
        let _ = state.close_workbook(&fork_workbook_id);

        Ok(StructureBatchResponse {
            fork_id: params.fork_id,
            mode,
            change_id: None,
            ops_applied: apply_result.ops_applied,
            summary,
        })
    }
}

struct StructureApplyResult {
    ops_applied: usize,
    summary: ChangeSummary,
}

fn apply_structure_ops_to_file(path: &Path, ops: &[StructureOp]) -> Result<StructureApplyResult> {
    let mut book = umya_spreadsheet::reader::xlsx::read(path)?;

    let mut affected_sheets: BTreeSet<String> = BTreeSet::new();
    let affected_bounds: Vec<String> = Vec::new();
    let mut counts: BTreeMap<String, u64> = BTreeMap::new();
    let mut warnings: Vec<String> = vec![
        "Structural edits may not fully rewrite formulas/named ranges like Excel. After apply, run recalculate and review get_changeset.".to_string(),
    ];

    for op in ops {
        match op {
            StructureOp::InsertRows {
                sheet_name,
                at_row,
                count,
            } => {
                if *at_row == 0 || *count == 0 {
                    bail!("insert_rows requires at_row>=1 and count>=1");
                }
                {
                    let sheet = book
                        .get_sheet_by_name_mut(sheet_name)
                        .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))?;
                    sheet.insert_new_row(at_row, count);
                }
                rewrite_formulas_for_sheet_row_insert(&mut book, sheet_name, *at_row, *count)?;
                rewrite_defined_name_formulas_for_sheet_row_insert(
                    &mut book, sheet_name, *at_row, *count,
                )?;
                affected_sheets.insert(sheet_name.clone());
                counts
                    .entry("rows_inserted".to_string())
                    .and_modify(|v| *v += *count as u64)
                    .or_insert(*count as u64);
            }
            StructureOp::DeleteRows {
                sheet_name,
                start_row,
                count,
            } => {
                if *start_row == 0 || *count == 0 {
                    bail!("delete_rows requires start_row>=1 and count>=1");
                }
                {
                    let sheet = book
                        .get_sheet_by_name_mut(sheet_name)
                        .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))?;
                    sheet.remove_row(start_row, count);
                }
                rewrite_formulas_for_sheet_row_delete(&mut book, sheet_name, *start_row, *count)?;
                rewrite_defined_name_formulas_for_sheet_row_delete(
                    &mut book, sheet_name, *start_row, *count,
                )?;
                affected_sheets.insert(sheet_name.clone());
                counts
                    .entry("rows_deleted".to_string())
                    .and_modify(|v| *v += *count as u64)
                    .or_insert(*count as u64);
            }
            StructureOp::InsertCols {
                sheet_name,
                at_col,
                count,
            } => {
                if at_col.trim().is_empty() || *count == 0 {
                    bail!("insert_cols requires at_col and count>=1");
                }
                let col_letters = normalize_col_letters(at_col)?;
                let root_col =
                    umya_spreadsheet::helper::coordinate::column_index_from_string(&col_letters);
                {
                    let sheet = book
                        .get_sheet_by_name_mut(sheet_name)
                        .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))?;
                    sheet.insert_new_column(&col_letters, count);
                }
                rewrite_formulas_for_sheet_col_insert(&mut book, sheet_name, root_col, *count)?;
                rewrite_defined_name_formulas_for_sheet_col_insert(
                    &mut book, sheet_name, root_col, *count,
                )?;
                affected_sheets.insert(sheet_name.clone());
                counts
                    .entry("cols_inserted".to_string())
                    .and_modify(|v| *v += *count as u64)
                    .or_insert(*count as u64);
            }
            StructureOp::DeleteCols {
                sheet_name,
                start_col,
                count,
            } => {
                if start_col.trim().is_empty() || *count == 0 {
                    bail!("delete_cols requires start_col and count>=1");
                }
                let col_letters = normalize_col_letters(start_col)?;
                let root_col =
                    umya_spreadsheet::helper::coordinate::column_index_from_string(&col_letters);
                {
                    let sheet = book
                        .get_sheet_by_name_mut(sheet_name)
                        .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))?;
                    sheet.remove_column(&col_letters, count);
                }
                rewrite_formulas_for_sheet_col_delete(&mut book, sheet_name, root_col, *count)?;
                rewrite_defined_name_formulas_for_sheet_col_delete(
                    &mut book, sheet_name, root_col, *count,
                )?;
                affected_sheets.insert(sheet_name.clone());
                counts
                    .entry("cols_deleted".to_string())
                    .and_modify(|v| *v += *count as u64)
                    .or_insert(*count as u64);
            }
            StructureOp::RenameSheet { old_name, new_name } => {
                let old_name = old_name.trim();
                let new_name = new_name.trim();
                if old_name.is_empty() || new_name.is_empty() {
                    bail!("rename_sheet requires non-empty old_name and new_name");
                }

                let sheet_index = book
                    .get_sheet_collection_no_check()
                    .iter()
                    .position(|s| s.get_name() == old_name)
                    .ok_or_else(|| anyhow!("sheet '{}' not found", old_name))?;
                book.set_sheet_name(sheet_index, new_name.to_string())
                    .map_err(|e| anyhow!("failed to rename sheet '{}': {}", old_name, e))?;

                rewrite_formulas_for_sheet_rename(&mut book, old_name, new_name)?;
                rewrite_defined_name_formulas_for_sheet_rename(&mut book, old_name, new_name)?;

                affected_sheets.insert(old_name.to_string());
                affected_sheets.insert(new_name.to_string());
                counts
                    .entry("sheets_renamed".to_string())
                    .and_modify(|v| *v += 1)
                    .or_insert(1);
            }
            StructureOp::CreateSheet { name, position } => {
                let name_trimmed = name.trim();
                if name_trimmed.is_empty() {
                    bail!("create_sheet requires non-empty name");
                }
                let requested_position = *position;
                book.new_sheet(name_trimmed.to_string())
                    .map_err(|e| anyhow!("failed to create sheet '{}': {}", name_trimmed, e))?;

                if let Some(pos) = requested_position {
                    let desired = pos as usize;
                    let len = book.get_sheet_collection_no_check().len();
                    if desired >= len {
                        warnings.push(format!(
                            "create_sheet position {} is out of range (sheet_count {}). Appended at end.",
                            desired, len
                        ));
                    } else if desired != len - 1 {
                        let sheets = book.get_sheet_collection_mut();
                        let created = sheets.remove(len - 1);
                        sheets.insert(desired, created);
                    }
                }

                affected_sheets.insert(name_trimmed.to_string());
                counts
                    .entry("sheets_created".to_string())
                    .and_modify(|v| *v += 1)
                    .or_insert(1);
            }
            StructureOp::DeleteSheet { name } => {
                let name_trimmed = name.trim();
                if name_trimmed.is_empty() {
                    bail!("delete_sheet requires non-empty name");
                }
                if book.get_sheet_collection_no_check().len() <= 1 {
                    bail!("cannot delete the last remaining sheet");
                }
                book.remove_sheet_by_name(name_trimmed)
                    .map_err(|e| anyhow!("failed to delete sheet '{}': {}", name_trimmed, e))?;
                affected_sheets.insert(name_trimmed.to_string());
                counts
                    .entry("sheets_deleted".to_string())
                    .and_modify(|v| *v += 1)
                    .or_insert(1);
            }
            StructureOp::CopyRange {
                sheet_name,
                dest_sheet_name,
                src_range,
                dest_anchor,
                include_styles,
                include_formulas,
            } => {
                let dest_sheet_name = dest_sheet_name.as_deref().unwrap_or(sheet_name);
                let result = copy_or_move_range(
                    &mut book,
                    sheet_name,
                    dest_sheet_name,
                    src_range,
                    dest_anchor,
                    *include_styles,
                    *include_formulas,
                    false,
                )?;
                affected_sheets.insert(sheet_name.clone());
                affected_sheets.insert(dest_sheet_name.to_string());
                counts
                    .entry("cells_copied".to_string())
                    .and_modify(|v| *v += result.cells_written)
                    .or_insert(result.cells_written);
                counts
                    .entry("ranges_copied".to_string())
                    .and_modify(|v| *v += 1)
                    .or_insert(1);
                warnings.extend(result.warnings);
            }
            StructureOp::MoveRange {
                sheet_name,
                dest_sheet_name,
                src_range,
                dest_anchor,
                include_styles,
                include_formulas,
            } => {
                let dest_sheet_name = dest_sheet_name.as_deref().unwrap_or(sheet_name);
                let result = copy_or_move_range(
                    &mut book,
                    sheet_name,
                    dest_sheet_name,
                    src_range,
                    dest_anchor,
                    *include_styles,
                    *include_formulas,
                    true,
                )?;
                affected_sheets.insert(sheet_name.clone());
                affected_sheets.insert(dest_sheet_name.to_string());
                counts
                    .entry("cells_moved".to_string())
                    .and_modify(|v| *v += result.cells_written)
                    .or_insert(result.cells_written);
                counts
                    .entry("ranges_moved".to_string())
                    .and_modify(|v| *v += 1)
                    .or_insert(1);
                warnings.extend(result.warnings);
            }
        }
    }

    umya_spreadsheet::writer::xlsx::write(&book, path)?;

    let summary = ChangeSummary {
        op_kinds: vec!["structure_batch".to_string()],
        affected_sheets: affected_sheets.into_iter().collect(),
        affected_bounds,
        counts,
        warnings,
    };

    Ok(StructureApplyResult {
        ops_applied: ops.len(),
        summary,
    })
}

fn normalize_col_letters(col: &str) -> Result<String> {
    let letters = col.trim().to_ascii_uppercase();
    if letters.is_empty() || !letters.chars().all(|c| c.is_ascii_alphabetic()) {
        bail!("invalid column reference: {}", col);
    }
    Ok(letters)
}

struct CopyMoveApplyResult {
    cells_written: u64,
    warnings: Vec<String>,
}

#[allow(clippy::too_many_arguments)]
fn ranges_intersect(
    a_min_col: u32,
    a_min_row: u32,
    a_max_col: u32,
    a_max_row: u32,
    b_min_col: u32,
    b_min_row: u32,
    b_max_col: u32,
    b_max_row: u32,
) -> bool {
    !(a_max_col < b_min_col
        || b_max_col < a_min_col
        || a_max_row < b_min_row
        || b_max_row < a_min_row)
}

#[allow(clippy::too_many_arguments)]
fn copy_or_move_range(
    book: &mut umya_spreadsheet::Spreadsheet,
    src_sheet_name: &str,
    dest_sheet_name: &str,
    src_range: &str,
    dest_anchor: &str,
    include_styles: bool,
    include_formulas: bool,
    clear_source: bool,
) -> Result<CopyMoveApplyResult> {
    let src_bounds = parse_range_bounds(src_range)?;
    let (dest_start_col, dest_start_row) = parse_cell_ref(dest_anchor)?;

    let width = src_bounds.cols;
    let height = src_bounds.rows;

    let dest_end_col = dest_start_col
        .checked_add(width.saturating_sub(1))
        .ok_or_else(|| anyhow!("destination range overflows column bounds"))?;
    let dest_end_row = dest_start_row
        .checked_add(height.saturating_sub(1))
        .ok_or_else(|| anyhow!("destination range overflows row bounds"))?;

    let same_sheet = src_sheet_name == dest_sheet_name;

    if same_sheet
        && ranges_intersect(
            src_bounds.min_col,
            src_bounds.min_row,
            src_bounds.max_col,
            src_bounds.max_row,
            dest_start_col,
            dest_start_row,
            dest_end_col,
            dest_end_row,
        )
    {
        let dest_range = if width == 1 && height == 1 {
            crate::utils::cell_address(dest_start_col, dest_start_row)
        } else {
            format!(
                "{}:{}",
                crate::utils::cell_address(dest_start_col, dest_start_row),
                crate::utils::cell_address(dest_end_col, dest_end_row)
            )
        };
        bail!(
            "copy/move destination overlaps source (src {}, dest {})",
            src_range,
            dest_range
        );
    }

    let delta_col = dest_start_col as i32 - src_bounds.min_col as i32;
    let delta_row = dest_start_row as i32 - src_bounds.min_row as i32;

    let mut warnings: Vec<String> = Vec::new();
    let mut formula_shift_failures: u64 = 0;
    let mut formula_value_copies: u64 = 0;

    let (src_sheet_index, dest_sheet_index) = {
        let sheets = book.get_sheet_collection_no_check();
        let src = sheets
            .iter()
            .position(|s| s.get_name() == src_sheet_name)
            .ok_or_else(|| anyhow!("sheet '{}' not found", src_sheet_name))?;
        let dest = sheets
            .iter()
            .position(|s| s.get_name() == dest_sheet_name)
            .ok_or_else(|| anyhow!("sheet '{}' not found", dest_sheet_name))?;
        (src, dest)
    };

    let sheets = book.get_sheet_collection_mut();

    if src_sheet_index == dest_sheet_index {
        let sheet = &mut sheets[src_sheet_index];

        for row in 0..height {
            for col in 0..width {
                let src_col = src_bounds.min_col + col;
                let src_row = src_bounds.min_row + row;
                let dest_col = dest_start_col + col;
                let dest_row = dest_start_row + row;

                let Some(src_cell) = sheet.get_cell((src_col, src_row)) else {
                    sheet.remove_cell((dest_col, dest_row));
                    continue;
                };

                let mut set_value = true;
                let mut dest_formula: Option<String> = None;

                if include_formulas && src_cell.is_formula() {
                    let src_formula = src_cell.get_formula().to_string();
                    match parse_base_formula(&src_formula).and_then(|ast| {
                        shift_formula_ast(&ast, delta_col, delta_row, RelativeMode::Excel)
                    }) {
                        Ok(shifted) => {
                            let shifted = shifted.strip_prefix('=').unwrap_or(&shifted).to_string();
                            dest_formula = Some(shifted);
                            set_value = false;
                        }
                        Err(_) => {
                            dest_formula = Some(src_formula);
                            set_value = false;
                            formula_shift_failures += 1;
                        }
                    }
                } else if !include_formulas && src_cell.is_formula() {
                    formula_value_copies += 1;
                }

                let src_value = src_cell.get_value().to_string();
                let src_style = src_cell.get_style().clone();

                let dest_cell = sheet.get_cell_mut((dest_col, dest_row));
                if include_styles {
                    dest_cell.set_style(src_style);
                }

                dest_cell.get_cell_value_mut().remove_formula();
                if let Some(formula) = dest_formula {
                    dest_cell.set_formula(formula);
                    dest_cell.set_formula_result_default("");
                }
                if set_value {
                    dest_cell.set_value(src_value);
                }
            }
        }

        if clear_source {
            for row in 0..height {
                for col in 0..width {
                    let src_col = src_bounds.min_col + col;
                    let src_row = src_bounds.min_row + row;
                    sheet.remove_cell((src_col, src_row));
                }
            }
        }
    } else {
        let (src_sheet, dest_sheet) = if src_sheet_index < dest_sheet_index {
            let (left, right) = sheets.split_at_mut(dest_sheet_index);
            (&mut left[src_sheet_index], &mut right[0])
        } else {
            let (left, right) = sheets.split_at_mut(src_sheet_index);
            (&mut right[0], &mut left[dest_sheet_index])
        };

        for row in 0..height {
            for col in 0..width {
                let src_col = src_bounds.min_col + col;
                let src_row = src_bounds.min_row + row;
                let dest_col = dest_start_col + col;
                let dest_row = dest_start_row + row;

                let Some(src_cell) = src_sheet.get_cell((src_col, src_row)) else {
                    dest_sheet.remove_cell((dest_col, dest_row));
                    continue;
                };

                let mut set_value = true;
                let mut dest_formula: Option<String> = None;

                if include_formulas && src_cell.is_formula() {
                    let src_formula = src_cell.get_formula().to_string();
                    match parse_base_formula(&src_formula).and_then(|ast| {
                        shift_formula_ast(&ast, delta_col, delta_row, RelativeMode::Excel)
                    }) {
                        Ok(shifted) => {
                            let shifted = shifted.strip_prefix('=').unwrap_or(&shifted).to_string();
                            dest_formula = Some(shifted);
                            set_value = false;
                        }
                        Err(_) => {
                            dest_formula = Some(src_formula);
                            set_value = false;
                            formula_shift_failures += 1;
                        }
                    }
                } else if !include_formulas && src_cell.is_formula() {
                    formula_value_copies += 1;
                }

                let src_value = src_cell.get_value().to_string();
                let src_style = src_cell.get_style().clone();

                let dest_cell = dest_sheet.get_cell_mut((dest_col, dest_row));
                if include_styles {
                    dest_cell.set_style(src_style);
                }

                dest_cell.get_cell_value_mut().remove_formula();
                if let Some(formula) = dest_formula {
                    dest_cell.set_formula(formula);
                    dest_cell.set_formula_result_default("");
                }
                if set_value {
                    dest_cell.set_value(src_value);
                }
            }
        }

        if clear_source {
            for row in 0..height {
                for col in 0..width {
                    let src_col = src_bounds.min_col + col;
                    let src_row = src_bounds.min_row + row;
                    src_sheet.remove_cell((src_col, src_row));
                }
            }
        }
    }

    if include_formulas && formula_shift_failures > 0 {
        warnings.push(format!(
            "Failed to shift {} formula(s); copied original formula text.",
            formula_shift_failures
        ));
    }
    if !include_formulas && formula_value_copies > 0 {
        warnings.push(format!(
            "Copied cached values for {} formula cell(s) (include_formulas=false); run recalculate for fresh results.",
            formula_value_copies
        ));
    }

    Ok(CopyMoveApplyResult {
        cells_written: width as u64 * height as u64,
        warnings,
    })
}

fn rewrite_formulas_for_sheet_rename(
    book: &mut umya_spreadsheet::Spreadsheet,
    old_name: &str,
    new_name: &str,
) -> Result<()> {
    let new_prefix = format_sheet_prefix_for_formula(new_name);

    for sheet in book.get_sheet_collection_mut().iter_mut() {
        for cell in sheet.get_cell_collection_mut() {
            if !cell.is_formula() {
                continue;
            }
            let formula_text = cell.get_formula();
            if formula_text.is_empty() {
                continue;
            }
            let formula_with_equals = if formula_text.starts_with('=') {
                formula_text.to_string()
            } else {
                format!("={}", formula_text)
            };

            let tokenizer = Tokenizer::new(&formula_with_equals)
                .map_err(|e| anyhow!("failed to tokenize formula: {}", e.message))?;

            let tokens = tokenizer.items;
            let mut out = String::with_capacity(formula_with_equals.len());
            let mut cursor = 0usize;

            for token in &tokens {
                if token.start > cursor {
                    out.push_str(&formula_with_equals[cursor..token.start]);
                }

                let mut value = token.value.clone();
                if token.subtype == formualizer_parse::TokenSubType::Range
                    && value.contains('!')
                    && let Some((sheet_part, tail)) = value.split_once('!')
                    && sheet_part_matches(sheet_part, old_name)
                {
                    value = format!("{}{}", new_prefix, tail);
                }

                out.push_str(&value);
                cursor = token.end;
            }

            if cursor < formula_with_equals.len() {
                out.push_str(&formula_with_equals[cursor..]);
            }

            let new_formula = out.strip_prefix('=').unwrap_or(&out);
            cell.set_formula(new_formula.to_string());
        }
    }

    Ok(())
}

fn rewrite_defined_name_formulas_for_sheet_rename(
    book: &mut umya_spreadsheet::Spreadsheet,
    old_name: &str,
    new_name: &str,
) -> Result<()> {
    let new_prefix = format_sheet_prefix_for_formula(new_name);

    for defined in book.get_defined_names_mut() {
        let refers_to = defined.get_address();
        let trimmed = refers_to.trim();
        let had_equals = trimmed.starts_with('=');
        let looks_like_formula = had_equals || trimmed.contains('(');
        if !looks_like_formula {
            continue;
        }

        let formula_in = if had_equals {
            trimmed.to_string()
        } else {
            format!("={}", trimmed)
        };

        let tokenizer = Tokenizer::new(&formula_in)
            .map_err(|e| anyhow!("failed to tokenize formula: {}", e.message))?;
        let tokens = tokenizer.items;

        let mut out = String::with_capacity(formula_in.len());
        let mut cursor = 0usize;
        let mut changed = false;

        for token in &tokens {
            if token.start > cursor {
                out.push_str(&formula_in[cursor..token.start]);
            }

            let mut value = token.value.clone();
            if token.subtype == formualizer_parse::TokenSubType::Range
                && value.contains('!')
                && let Some((sheet_part, tail)) = value.split_once('!')
                && sheet_part_matches(sheet_part, old_name)
            {
                value = format!("{}{}", new_prefix, tail);
                changed = true;
            }

            out.push_str(&value);
            cursor = token.end;
        }

        if cursor < formula_in.len() {
            out.push_str(&formula_in[cursor..]);
        }

        if changed {
            let out_final = if had_equals {
                out
            } else {
                out.strip_prefix('=').unwrap_or(&out).to_string()
            };
            defined.set_address(out_final);
        }
    }

    Ok(())
}

fn rewrite_defined_name_formulas_for_sheet_col_insert(
    book: &mut umya_spreadsheet::Spreadsheet,
    sheet_name: &str,
    at_col: u32,
    count: u32,
) -> Result<()> {
    rewrite_defined_name_formulas_for_sheet_structure_change(
        book,
        sheet_name,
        StructureAxis::Col,
        StructureEdit::Insert { at: at_col, count },
    )
}

fn rewrite_defined_name_formulas_for_sheet_col_delete(
    book: &mut umya_spreadsheet::Spreadsheet,
    sheet_name: &str,
    start_col: u32,
    count: u32,
) -> Result<()> {
    rewrite_defined_name_formulas_for_sheet_structure_change(
        book,
        sheet_name,
        StructureAxis::Col,
        StructureEdit::Delete {
            start: start_col,
            count,
        },
    )
}

fn rewrite_defined_name_formulas_for_sheet_row_insert(
    book: &mut umya_spreadsheet::Spreadsheet,
    sheet_name: &str,
    at_row: u32,
    count: u32,
) -> Result<()> {
    rewrite_defined_name_formulas_for_sheet_structure_change(
        book,
        sheet_name,
        StructureAxis::Row,
        StructureEdit::Insert { at: at_row, count },
    )
}

fn rewrite_defined_name_formulas_for_sheet_row_delete(
    book: &mut umya_spreadsheet::Spreadsheet,
    sheet_name: &str,
    start_row: u32,
    count: u32,
) -> Result<()> {
    rewrite_defined_name_formulas_for_sheet_structure_change(
        book,
        sheet_name,
        StructureAxis::Row,
        StructureEdit::Delete {
            start: start_row,
            count,
        },
    )
}

fn rewrite_defined_name_formulas_for_sheet_structure_change(
    book: &mut umya_spreadsheet::Spreadsheet,
    sheet_name: &str,
    axis: StructureAxis,
    edit: StructureEdit,
) -> Result<()> {
    for defined in book.get_defined_names_mut() {
        let refers_to = defined.get_address();
        let trimmed = refers_to.trim();
        let had_equals = trimmed.starts_with('=');
        let looks_like_formula = had_equals || trimmed.contains('(');
        if !looks_like_formula {
            continue;
        }

        let formula_in = if had_equals {
            trimmed.to_string()
        } else {
            format!("={}", trimmed)
        };

        let tokenizer = Tokenizer::new(&formula_in)
            .map_err(|e| anyhow!("failed to tokenize formula: {}", e.message))?;
        let tokens = tokenizer.items;

        let mut out = String::with_capacity(formula_in.len());
        let mut cursor = 0usize;
        let mut changed = false;

        for token in &tokens {
            if token.start > cursor {
                out.push_str(&formula_in[cursor..token.start]);
            }

            let mut value = token.value.clone();
            if token.subtype == formualizer_parse::TokenSubType::Range
                && value.contains('!')
                && let Some((sheet_part, coord_part)) = value.split_once('!')
                && sheet_part_matches(sheet_part, sheet_name)
            {
                let adjusted = adjust_ref_coord_part(coord_part, axis, edit)?;
                value = format!("{sheet_part}!{adjusted}");
                changed = true;
            }

            out.push_str(&value);
            cursor = token.end;
        }

        if cursor < formula_in.len() {
            out.push_str(&formula_in[cursor..]);
        }

        if changed {
            let out_final = if had_equals {
                out
            } else {
                out.strip_prefix('=').unwrap_or(&out).to_string()
            };
            defined.set_address(out_final);
        }
    }

    Ok(())
}

fn rewrite_formulas_for_sheet_col_insert(
    book: &mut umya_spreadsheet::Spreadsheet,
    sheet_name: &str,
    at_col: u32,
    count: u32,
) -> Result<()> {
    rewrite_formulas_for_sheet_structure_change(
        book,
        sheet_name,
        StructureAxis::Col,
        StructureEdit::Insert { at: at_col, count },
    )
}

fn rewrite_formulas_for_sheet_col_delete(
    book: &mut umya_spreadsheet::Spreadsheet,
    sheet_name: &str,
    start_col: u32,
    count: u32,
) -> Result<()> {
    rewrite_formulas_for_sheet_structure_change(
        book,
        sheet_name,
        StructureAxis::Col,
        StructureEdit::Delete {
            start: start_col,
            count,
        },
    )
}

fn rewrite_formulas_for_sheet_row_insert(
    book: &mut umya_spreadsheet::Spreadsheet,
    sheet_name: &str,
    at_row: u32,
    count: u32,
) -> Result<()> {
    rewrite_formulas_for_sheet_structure_change(
        book,
        sheet_name,
        StructureAxis::Row,
        StructureEdit::Insert { at: at_row, count },
    )
}

fn rewrite_formulas_for_sheet_row_delete(
    book: &mut umya_spreadsheet::Spreadsheet,
    sheet_name: &str,
    start_row: u32,
    count: u32,
) -> Result<()> {
    rewrite_formulas_for_sheet_structure_change(
        book,
        sheet_name,
        StructureAxis::Row,
        StructureEdit::Delete {
            start: start_row,
            count,
        },
    )
}

#[derive(Debug, Clone, Copy)]
enum StructureAxis {
    Row,
    Col,
}

#[derive(Debug, Clone, Copy)]
enum StructureEdit {
    Insert { at: u32, count: u32 },
    Delete { start: u32, count: u32 },
}

fn rewrite_formulas_for_sheet_structure_change(
    book: &mut umya_spreadsheet::Spreadsheet,
    sheet_name: &str,
    axis: StructureAxis,
    edit: StructureEdit,
) -> Result<()> {
    for sheet in book.get_sheet_collection_mut().iter_mut() {
        if sheet.get_name() == sheet_name {
            continue;
        }
        for cell in sheet.get_cell_collection_mut() {
            if !cell.is_formula() {
                continue;
            }
            let formula_text = cell.get_formula();
            if formula_text.is_empty() {
                continue;
            }
            let formula_with_equals = if formula_text.starts_with('=') {
                formula_text.to_string()
            } else {
                format!("={}", formula_text)
            };
            let tokenizer = Tokenizer::new(&formula_with_equals)
                .map_err(|e| anyhow!("failed to tokenize formula: {}", e.message))?;
            let tokens = tokenizer.items;

            let mut out = String::with_capacity(formula_with_equals.len());
            let mut cursor = 0usize;
            let mut changed = false;

            for token in &tokens {
                if token.start > cursor {
                    out.push_str(&formula_with_equals[cursor..token.start]);
                }

                let mut value = token.value.clone();
                if token.subtype == formualizer_parse::TokenSubType::Range
                    && value.contains('!')
                    && let Some((sheet_part, coord_part)) = value.split_once('!')
                    && sheet_part_matches(sheet_part, sheet_name)
                {
                    let adjusted = adjust_ref_coord_part(coord_part, axis, edit)?;
                    value = format!("{sheet_part}!{adjusted}");
                    changed = true;
                }

                out.push_str(&value);
                cursor = token.end;
            }

            if cursor < formula_with_equals.len() {
                out.push_str(&formula_with_equals[cursor..]);
            }

            if changed {
                let new_formula = out.strip_prefix('=').unwrap_or(&out);
                cell.set_formula(new_formula.to_string());
            }
        }
    }
    Ok(())
}

fn adjust_ref_coord_part(
    coord_part: &str,
    axis: StructureAxis,
    edit: StructureEdit,
) -> Result<String> {
    if coord_part == "#REF!" {
        return Ok(coord_part.to_string());
    }
    if let Some((start, end)) = coord_part.split_once(':') {
        let start_adj = adjust_ref_segment(start, axis, edit)?;
        let end_adj = adjust_ref_segment(end, axis, edit)?;
        if start_adj == "#REF!" || end_adj == "#REF!" {
            return Ok("#REF!".to_string());
        }
        Ok(format!("{start_adj}:{end_adj}"))
    } else {
        Ok(adjust_ref_segment(coord_part, axis, edit)?)
    }
}

fn adjust_ref_segment(segment: &str, axis: StructureAxis, edit: StructureEdit) -> Result<String> {
    use umya_spreadsheet::helper::coordinate::{
        coordinate_from_index_with_lock, index_from_coordinate, string_from_column_index,
    };

    let (col, row, col_lock, row_lock) = index_from_coordinate(segment);
    let mut col = col;
    let mut row = row;

    match axis {
        StructureAxis::Col => {
            if let Some(c) = col {
                col = match edit {
                    StructureEdit::Insert { at, count } => Some(adjust_insert(c, at, count)),
                    StructureEdit::Delete { start, count } => adjust_delete(c, start, count),
                };
            }
        }
        StructureAxis::Row => {
            if let Some(r) = row {
                row = match edit {
                    StructureEdit::Insert { at, count } => Some(adjust_insert(r, at, count)),
                    StructureEdit::Delete { start, count } => adjust_delete(r, start, count),
                };
            }
        }
    }

    if col.is_none() && row.is_none() {
        return Ok("#REF!".to_string());
    }

    match (col, row) {
        (Some(c), Some(r)) => Ok(coordinate_from_index_with_lock(
            &c,
            &r,
            &col_lock.unwrap_or(false),
            &row_lock.unwrap_or(false),
        )),
        (Some(c), None) => {
            let col_str = string_from_column_index(&c);
            Ok(format!(
                "{}{}",
                if col_lock.unwrap_or(false) { "$" } else { "" },
                col_str
            ))
        }
        (None, Some(r)) => Ok(format!(
            "{}{}",
            if row_lock.unwrap_or(false) { "$" } else { "" },
            r
        )),
        (None, None) => Ok("#REF!".to_string()),
    }
}

fn adjust_insert(value: u32, at: u32, count: u32) -> u32 {
    if value >= at { value + count } else { value }
}

fn adjust_delete(value: u32, start: u32, count: u32) -> Option<u32> {
    let end = start.saturating_add(count.saturating_sub(1));
    if value >= start && value <= end {
        None
    } else if value > end {
        Some(value - count)
    } else {
        Some(value)
    }
}

fn sheet_part_matches(sheet_part: &str, old_name: &str) -> bool {
    let trimmed = sheet_part.trim();
    if let Some(stripped) = trimmed.strip_prefix('\'')
        && let Some(inner) = stripped.strip_suffix('\'')
    {
        return inner.replace("''", "'") == old_name;
    }
    trimmed == old_name
}

fn format_sheet_prefix_for_formula(sheet_name: &str) -> String {
    if sheet_name_needs_quoting_for_formula(sheet_name) {
        let escaped = sheet_name.replace('\'', "''");
        format!("'{escaped}'!")
    } else {
        format!("{sheet_name}!")
    }
}

fn sheet_name_needs_quoting_for_formula(name: &str) -> bool {
    if name.is_empty() {
        return false;
    }
    let bytes = name.as_bytes();
    if bytes[0].is_ascii_digit() {
        return true;
    }
    for &byte in bytes {
        match byte {
            b' ' | b'!' | b'"' | b'#' | b'$' | b'%' | b'&' | b'\'' | b'(' | b')' | b'*' | b'+'
            | b',' | b'-' | b'.' | b'/' | b':' | b';' | b'<' | b'=' | b'>' | b'?' | b'@' | b'['
            | b'\\' | b']' | b'^' | b'`' | b'{' | b'|' | b'}' | b'~' => return true,
            _ => {}
        }
    }
    let upper = name.to_uppercase();
    matches!(
        upper.as_str(),
        "TRUE" | "FALSE" | "NULL" | "REF" | "DIV" | "NAME" | "NUM" | "VALUE" | "N/A"
    )
}

struct StyleApplyResult {
    ops_applied: usize,
    summary: ChangeSummary,
}

fn stage_snapshot_path(fork_id: &str, change_id: &str) -> PathBuf {
    PathBuf::from("/tmp/mcp-staged").join(format!("{fork_id}_{change_id}.xlsx"))
}

struct TransformApplyResult {
    ops_applied: usize,
    summary: ChangeSummary,
}

fn apply_transform_ops_to_file(path: &Path, ops: &[TransformOp]) -> Result<TransformApplyResult> {
    let mut book = umya_spreadsheet::reader::xlsx::read(path)?;

    let mut sheets: BTreeSet<String> = BTreeSet::new();
    let mut affected_bounds: Vec<String> = Vec::new();

    let mut cells_touched: u64 = 0;
    let mut cells_value_cleared: u64 = 0;
    let mut cells_formula_cleared: u64 = 0;
    let mut cells_skipped_keep_formulas: u64 = 0;

    let mut cells_value_set: u64 = 0;
    let mut cells_formula_set: u64 = 0;
    let mut cells_value_replaced: u64 = 0;
    let mut cells_formula_replaced: u64 = 0;

    for op in ops {
        match op {
            TransformOp::ClearRange {
                sheet_name,
                target,
                clear_values,
                clear_formulas,
            } => {
                let sheet = book
                    .get_sheet_by_name_mut(sheet_name)
                    .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))?;
                sheets.insert(sheet_name.clone());

                match target {
                    TransformTarget::Range { range } => {
                        let bounds = parse_range_bounds(range)?;
                        affected_bounds.push(range.clone());

                        for row in bounds.min_row..=bounds.max_row {
                            for col in bounds.min_col..=bounds.max_col {
                                let exists = sheet.get_cell((col, row)).is_some();
                                if !exists {
                                    continue;
                                }

                                let cell = sheet.get_cell_mut((col, row));
                                let was_formula = cell.is_formula();
                                cells_touched += 1;

                                if *clear_formulas && was_formula {
                                    cell.set_formula(String::new());
                                    cells_formula_cleared += 1;
                                }

                                if *clear_values {
                                    if was_formula && !*clear_formulas {
                                        cells_skipped_keep_formulas += 1;
                                    } else {
                                        if !cell.get_value().is_empty() {
                                            cells_value_cleared += 1;
                                        }
                                        cell.set_value(String::new());
                                    }
                                }
                            }
                        }
                    }
                    TransformTarget::Cells { cells } => {
                        affected_bounds.extend(cells.iter().cloned());
                        for addr in cells {
                            let exists = sheet.get_cell(addr.as_str()).is_some();
                            if !exists {
                                continue;
                            }

                            let cell = sheet.get_cell_mut(addr.as_str());
                            let was_formula = cell.is_formula();
                            cells_touched += 1;

                            if *clear_formulas && was_formula {
                                cell.set_formula(String::new());
                                cells_formula_cleared += 1;
                            }

                            if *clear_values {
                                if was_formula && !*clear_formulas {
                                    cells_skipped_keep_formulas += 1;
                                } else {
                                    if !cell.get_value().is_empty() {
                                        cells_value_cleared += 1;
                                    }
                                    cell.set_value(String::new());
                                }
                            }
                        }
                    }
                    TransformTarget::Region { .. } => {
                        return Err(anyhow!(
                            "region_id targets must be resolved before apply_transform_ops_to_file"
                        ));
                    }
                }
            }
            TransformOp::FillRange {
                sheet_name,
                target,
                value,
                is_formula,
                overwrite_formulas,
            } => {
                let sheet = book
                    .get_sheet_by_name_mut(sheet_name)
                    .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))?;
                sheets.insert(sheet_name.clone());

                match target {
                    TransformTarget::Range { range } => {
                        let bounds = parse_range_bounds(range)?;
                        affected_bounds.push(range.clone());

                        for row in bounds.min_row..=bounds.max_row {
                            for col in bounds.min_col..=bounds.max_col {
                                let cell = sheet.get_cell_mut((col, row));
                                cells_touched += 1;

                                if !*is_formula && cell.is_formula() {
                                    if !*overwrite_formulas {
                                        cells_skipped_keep_formulas += 1;
                                        continue;
                                    }
                                    cell.set_formula(String::new());
                                    cells_formula_cleared += 1;
                                }

                                if *is_formula {
                                    cell.set_formula(value.clone());
                                    cells_formula_set += 1;
                                } else {
                                    cell.set_value(value.clone());
                                    cells_value_set += 1;
                                }
                            }
                        }
                    }
                    TransformTarget::Cells { cells } => {
                        affected_bounds.extend(cells.iter().cloned());
                        for addr in cells {
                            let cell = sheet.get_cell_mut(addr.as_str());
                            cells_touched += 1;

                            if !*is_formula && cell.is_formula() {
                                if !*overwrite_formulas {
                                    cells_skipped_keep_formulas += 1;
                                    continue;
                                }
                                cell.set_formula(String::new());
                                cells_formula_cleared += 1;
                            }

                            if *is_formula {
                                cell.set_formula(value.clone());
                                cells_formula_set += 1;
                            } else {
                                cell.set_value(value.clone());
                                cells_value_set += 1;
                            }
                        }
                    }
                    TransformTarget::Region { .. } => {
                        return Err(anyhow!(
                            "region_id targets must be resolved before apply_transform_ops_to_file"
                        ));
                    }
                }
            }
            TransformOp::ReplaceInRange {
                sheet_name,
                target,
                find,
                replace,
                match_mode,
                case_sensitive,
                include_formulas,
            } => {
                let sheet = book
                    .get_sheet_by_name_mut(sheet_name)
                    .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))?;
                sheets.insert(sheet_name.clone());

                let match_mode = match_mode.to_ascii_lowercase();
                if match_mode != "exact" && match_mode != "contains" {
                    return Err(anyhow!(
                        "invalid match_mode '{}'; expected 'exact' or 'contains'",
                        match_mode
                    ));
                }
                if match_mode == "contains" && !*case_sensitive {
                    return Err(anyhow!(
                        "match_mode 'contains' requires case_sensitive=true"
                    ));
                }

                let replace_value = |input: &str| -> Option<String> {
                    if match_mode == "exact" {
                        if *case_sensitive {
                            (input == find).then(|| replace.clone())
                        } else {
                            input.eq_ignore_ascii_case(find).then(|| replace.clone())
                        }
                    } else if input.contains(find) {
                        Some(input.replace(find, replace))
                    } else {
                        None
                    }
                };

                match target {
                    TransformTarget::Range { range } => {
                        let bounds = parse_range_bounds(range)?;
                        affected_bounds.push(range.clone());

                        for row in bounds.min_row..=bounds.max_row {
                            for col in bounds.min_col..=bounds.max_col {
                                let exists = sheet.get_cell((col, row)).is_some();
                                if !exists {
                                    continue;
                                }

                                let cell = sheet.get_cell_mut((col, row));
                                cells_touched += 1;

                                if cell.is_formula() {
                                    if !*include_formulas {
                                        cells_skipped_keep_formulas += 1;
                                        continue;
                                    }

                                    let formula = cell.get_formula().to_string();
                                    if formula.is_empty() {
                                        continue;
                                    }
                                    if let Some(next) = replace_value(&formula) {
                                        cell.set_formula(next);
                                        cells_formula_replaced += 1;
                                    }
                                    continue;
                                }

                                let value = cell.get_value().to_string();
                                if value.is_empty() {
                                    continue;
                                }
                                if let Some(next) = replace_value(&value) {
                                    cell.set_value(next);
                                    cells_value_replaced += 1;
                                }
                            }
                        }
                    }
                    TransformTarget::Cells { cells } => {
                        affected_bounds.extend(cells.iter().cloned());
                        for addr in cells {
                            let exists = sheet.get_cell(addr.as_str()).is_some();
                            if !exists {
                                continue;
                            }

                            let cell = sheet.get_cell_mut(addr.as_str());
                            cells_touched += 1;

                            if cell.is_formula() {
                                if !*include_formulas {
                                    cells_skipped_keep_formulas += 1;
                                    continue;
                                }

                                let formula = cell.get_formula().to_string();
                                if formula.is_empty() {
                                    continue;
                                }
                                if let Some(next) = replace_value(&formula) {
                                    cell.set_formula(next);
                                    cells_formula_replaced += 1;
                                }
                                continue;
                            }

                            let value = cell.get_value().to_string();
                            if value.is_empty() {
                                continue;
                            }
                            if let Some(next) = replace_value(&value) {
                                cell.set_value(next);
                                cells_value_replaced += 1;
                            }
                        }
                    }
                    TransformTarget::Region { .. } => {
                        return Err(anyhow!(
                            "region_id targets must be resolved before apply_transform_ops_to_file"
                        ));
                    }
                }
            }
        }
    }

    umya_spreadsheet::writer::xlsx::write(&book, path)?;

    let mut counts = BTreeMap::new();
    counts.insert("cells_touched".to_string(), cells_touched);
    counts.insert("cells_value_cleared".to_string(), cells_value_cleared);
    counts.insert("cells_formula_cleared".to_string(), cells_formula_cleared);
    counts.insert(
        "cells_skipped_keep_formulas".to_string(),
        cells_skipped_keep_formulas,
    );

    counts.insert("cells_value_set".to_string(), cells_value_set);
    counts.insert("cells_formula_set".to_string(), cells_formula_set);
    counts.insert("cells_value_replaced".to_string(), cells_value_replaced);
    counts.insert("cells_formula_replaced".to_string(), cells_formula_replaced);

    let summary = ChangeSummary {
        op_kinds: vec!["transform_batch".to_string()],
        affected_sheets: sheets.into_iter().collect(),
        affected_bounds,
        counts,
        warnings: Vec::new(),
    };

    Ok(TransformApplyResult {
        ops_applied: ops.len(),
        summary,
    })
}

fn apply_style_ops_to_file(path: &Path, ops: &[StyleOp]) -> Result<StyleApplyResult> {
    use crate::styles::{
        StylePatchMode, apply_style_patch, descriptor_from_style, stable_style_id,
    };

    let mut book = umya_spreadsheet::reader::xlsx::read(path)?;

    let mut sheets: BTreeSet<String> = BTreeSet::new();
    let mut affected_bounds: Vec<String> = Vec::new();
    let mut cells_touched: u64 = 0;
    let mut cells_style_changed: u64 = 0;

    for op in ops {
        let sheet = book
            .get_sheet_by_name_mut(&op.sheet_name)
            .ok_or_else(|| anyhow!("sheet '{}' not found", op.sheet_name))?;
        sheets.insert(op.sheet_name.clone());

        let op_mode = match op
            .op_mode
            .as_deref()
            .unwrap_or("merge")
            .to_ascii_lowercase()
            .as_str()
        {
            "set" => StylePatchMode::Set,
            "clear" => StylePatchMode::Clear,
            _ => StylePatchMode::Merge,
        };

        match &op.target {
            StyleTarget::Range { range } => {
                let bounds = parse_range_bounds(range)?;
                affected_bounds.push(range.clone());
                for row in bounds.min_row..=bounds.max_row {
                    for col in bounds.min_col..=bounds.max_col {
                        let addr = crate::utils::cell_address(col, row);
                        let cell = sheet.get_cell_mut(addr.as_str());
                        let before = stable_style_id(&descriptor_from_style(cell.get_style()));
                        let next_style = apply_style_patch(cell.get_style(), &op.patch, op_mode);
                        cell.set_style(next_style);
                        let after = stable_style_id(&descriptor_from_style(cell.get_style()));
                        cells_touched += 1;
                        if before != after {
                            cells_style_changed += 1;
                        }
                    }
                }
            }
            StyleTarget::Cells { cells } => {
                affected_bounds.extend(cells.iter().cloned());
                for addr in cells {
                    let cell = sheet.get_cell_mut(addr.as_str());
                    let before = stable_style_id(&descriptor_from_style(cell.get_style()));
                    let next_style = apply_style_patch(cell.get_style(), &op.patch, op_mode);
                    cell.set_style(next_style);
                    let after = stable_style_id(&descriptor_from_style(cell.get_style()));
                    cells_touched += 1;
                    if before != after {
                        cells_style_changed += 1;
                    }
                }
            }
            StyleTarget::Region { .. } => {
                return Err(anyhow!(
                    "region_id targets must be resolved before apply_style_ops_to_file"
                ));
            }
        }
    }

    umya_spreadsheet::writer::xlsx::write(&book, path)?;

    let mut counts = BTreeMap::new();
    counts.insert("cells_touched".to_string(), cells_touched);
    counts.insert("cells_style_changed".to_string(), cells_style_changed);

    let summary = ChangeSummary {
        op_kinds: vec!["style_batch".to_string()],
        affected_sheets: sheets.into_iter().collect(),
        affected_bounds,
        counts,
        warnings: Vec::new(),
    };

    Ok(StyleApplyResult {
        ops_applied: ops.len(),
        summary,
    })
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

fn default_get_changeset_limit() -> u32 {
    200
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct GetChangesetParams {
    pub fork_id: String,
    pub sheet_name: Option<String>,
    #[serde(default = "default_get_changeset_limit")]
    pub limit: u32,
    #[serde(default)]
    pub offset: u32,
    #[serde(default)]
    pub summary_only: bool,
    #[serde(default)]
    pub include_types: Option<Vec<String>>,
    #[serde(default)]
    pub exclude_types: Option<Vec<String>>,
    #[serde(default)]
    pub include_subtypes: Option<Vec<String>>,
    #[serde(default)]
    pub exclude_subtypes: Option<Vec<String>>,
}

impl Default for GetChangesetParams {
    fn default() -> Self {
        Self {
            fork_id: String::new(),
            sheet_name: None,
            limit: default_get_changeset_limit(),
            offset: 0,
            summary_only: false,
            include_types: None,
            exclude_types: None,
            include_subtypes: None,
            exclude_subtypes: None,
        }
    }
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct ChangesetSummary {
    pub total_changes: u32,
    pub returned_changes: u32,
    pub truncated: bool,
    pub next_offset: Option<u32>,
    pub counts_by_kind: BTreeMap<String, u32>,
    pub counts_by_type: BTreeMap<String, u32>,
    pub counts_by_subtype: BTreeMap<String, u32>,
    pub affected_sheets: Vec<String>,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct GetChangesetResponse {
    pub fork_id: String,
    pub base_workbook: String,
    pub changes: Vec<crate::diff::Change>,
    pub summary: ChangesetSummary,
}

fn normalize_filter(values: &Option<Vec<String>>) -> Option<BTreeSet<String>> {
    values.as_ref().map(|items| {
        items
            .iter()
            .map(|s| s.to_ascii_lowercase())
            .collect::<BTreeSet<_>>()
    })
}

fn change_kind_key(change: &crate::diff::Change) -> &'static str {
    match change {
        crate::diff::Change::Cell(_) => "cell",
        crate::diff::Change::Table(_) => "table",
        crate::diff::Change::Name(_) => "name",
    }
}

fn change_type_key(change: &crate::diff::Change) -> &'static str {
    use crate::diff::merge::CellDiff;
    match change {
        crate::diff::Change::Cell(cell) => match &cell.diff {
            CellDiff::Added { .. } => "added",
            CellDiff::Deleted { .. } => "deleted",
            CellDiff::Modified { .. } => "modified",
        },
        crate::diff::Change::Table(table) => match table {
            crate::diff::tables::TableDiff::TableAdded { .. } => "table_added",
            crate::diff::tables::TableDiff::TableDeleted { .. } => "table_deleted",
            crate::diff::tables::TableDiff::TableModified { .. } => "table_modified",
        },
        crate::diff::Change::Name(name) => match name {
            crate::diff::names::NameDiff::NameAdded { .. } => "name_added",
            crate::diff::names::NameDiff::NameDeleted { .. } => "name_deleted",
            crate::diff::names::NameDiff::NameModified { .. } => "name_modified",
        },
    }
}

fn change_subtype_key(change: &crate::diff::Change) -> Option<&'static str> {
    use crate::diff::merge::{CellDiff, ModificationType};
    match change {
        crate::diff::Change::Cell(cell) => match &cell.diff {
            CellDiff::Modified { subtype, .. } => Some(match subtype {
                ModificationType::FormulaEdit => "formula_edit",
                ModificationType::RecalcResult => "recalc_result",
                ModificationType::ValueEdit => "value_edit",
                ModificationType::StyleEdit => "style_edit",
            }),
            _ => None,
        },
        _ => None,
    }
}

fn change_sheet_name(change: &crate::diff::Change) -> Option<&str> {
    match change {
        crate::diff::Change::Cell(cell) => Some(cell.sheet.as_str()),
        crate::diff::Change::Table(table) => match table {
            crate::diff::tables::TableDiff::TableAdded { sheet, .. }
            | crate::diff::tables::TableDiff::TableDeleted { sheet, .. }
            | crate::diff::tables::TableDiff::TableModified { sheet, .. } => Some(sheet.as_str()),
        },
        crate::diff::Change::Name(name) => match name {
            crate::diff::names::NameDiff::NameAdded { scope_sheet, .. }
            | crate::diff::names::NameDiff::NameDeleted { scope_sheet, .. }
            | crate::diff::names::NameDiff::NameModified { scope_sheet, .. } => {
                scope_sheet.as_deref()
            }
        },
    }
}

fn change_passes_filters(
    change: &crate::diff::Change,
    include_types: &Option<BTreeSet<String>>,
    exclude_types: &Option<BTreeSet<String>>,
    include_subtypes: &Option<BTreeSet<String>>,
    exclude_subtypes: &Option<BTreeSet<String>>,
) -> bool {
    let type_key = change_type_key(change);
    let subtype_key = change_subtype_key(change);

    if let Some(include) = include_types
        && !include.contains(type_key)
    {
        return false;
    }
    if let Some(exclude) = exclude_types
        && exclude.contains(type_key)
    {
        return false;
    }

    if let Some(include) = include_subtypes
        && subtype_key.is_none_or(|subtype| !include.contains(subtype))
    {
        return false;
    }
    if let Some(exclude) = exclude_subtypes
        && subtype_key.is_some_and(|subtype| exclude.contains(subtype))
    {
        return false;
    }

    true
}

pub async fn get_changeset(
    state: Arc<AppState>,
    params: GetChangesetParams,
) -> Result<GetChangesetResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let fork_ctx = registry.get_fork(&params.fork_id)?;

    let raw_changes = tokio::task::spawn_blocking({
        let base_path = fork_ctx.base_path.clone();
        let work_path = fork_ctx.work_path.clone();
        let sheet_filter = params.sheet_name.clone();
        move || crate::diff::calculate_changeset(&base_path, &work_path, sheet_filter.as_deref())
    })
    .await??;

    let include_types = normalize_filter(&params.include_types);
    let exclude_types = normalize_filter(&params.exclude_types);
    let include_subtypes = normalize_filter(&params.include_subtypes);
    let exclude_subtypes = normalize_filter(&params.exclude_subtypes);

    let mut affected_sheets: BTreeSet<String> = BTreeSet::new();
    let mut counts_by_kind: BTreeMap<String, u32> = BTreeMap::new();
    let mut counts_by_type: BTreeMap<String, u32> = BTreeMap::new();
    let mut counts_by_subtype: BTreeMap<String, u32> = BTreeMap::new();

    let mut filtered: Vec<crate::diff::Change> = Vec::new();
    for change in raw_changes {
        if !change_passes_filters(
            &change,
            &include_types,
            &exclude_types,
            &include_subtypes,
            &exclude_subtypes,
        ) {
            continue;
        }

        *counts_by_kind
            .entry(change_kind_key(&change).to_string())
            .or_default() += 1;
        *counts_by_type
            .entry(change_type_key(&change).to_string())
            .or_default() += 1;
        if let Some(subtype) = change_subtype_key(&change) {
            *counts_by_subtype.entry(subtype.to_string()).or_default() += 1;
        }
        if let Some(sheet) = change_sheet_name(&change) {
            affected_sheets.insert(sheet.to_string());
        }

        filtered.push(change);
    }

    let limit = params.limit.clamp(1, 2000) as usize;
    let offset = params.offset as usize;
    let total = filtered.len();

    let (returned_changes, changes, truncated, next_offset) = if params.summary_only {
        (0u32, Vec::new(), false, None)
    } else {
        let end = offset.saturating_add(limit);
        let truncated = end < total;
        let next_offset = truncated.then_some(end as u32);
        let changes: Vec<_> = filtered.into_iter().skip(offset).take(limit).collect();
        (changes.len() as u32, changes, truncated, next_offset)
    };

    let summary = ChangesetSummary {
        total_changes: total as u32,
        returned_changes,
        truncated,
        next_offset,
        counts_by_kind,
        counts_by_type,
        counts_by_subtype,
        affected_sheets: affected_sheets.into_iter().collect(),
    };

    Ok(GetChangesetResponse {
        fork_id: params.fork_id,
        base_workbook: fork_ctx.base_path.display().to_string(),
        changes,
        summary,
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

#[derive(Debug, Deserialize, JsonSchema)]
pub struct CheckpointForkParams {
    pub fork_id: String,
    pub label: Option<String>,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct CheckpointInfo {
    pub checkpoint_id: String,
    pub created_at: String,
    pub label: Option<String>,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct CheckpointForkResponse {
    pub fork_id: String,
    pub checkpoint: CheckpointInfo,
    pub total_checkpoints: usize,
}

pub async fn checkpoint_fork(
    state: Arc<AppState>,
    params: CheckpointForkParams,
) -> Result<CheckpointForkResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    registry.get_fork(&params.fork_id)?;
    let checkpoint = registry.create_checkpoint(&params.fork_id, params.label.clone())?;
    let total = registry.list_checkpoints(&params.fork_id)?.len();

    Ok(CheckpointForkResponse {
        fork_id: params.fork_id,
        checkpoint: CheckpointInfo {
            checkpoint_id: checkpoint.checkpoint_id,
            created_at: checkpoint.created_at.to_rfc3339(),
            label: checkpoint.label,
        },
        total_checkpoints: total,
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ListCheckpointsParams {
    pub fork_id: String,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct ListCheckpointsResponse {
    pub fork_id: String,
    pub checkpoints: Vec<CheckpointInfo>,
}

pub async fn list_checkpoints(
    state: Arc<AppState>,
    params: ListCheckpointsParams,
) -> Result<ListCheckpointsResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let checkpoints = registry.list_checkpoints(&params.fork_id)?;
    let checkpoints = checkpoints
        .into_iter()
        .map(|cp| CheckpointInfo {
            checkpoint_id: cp.checkpoint_id,
            created_at: cp.created_at.to_rfc3339(),
            label: cp.label,
        })
        .collect();

    Ok(ListCheckpointsResponse {
        fork_id: params.fork_id,
        checkpoints,
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct RestoreCheckpointParams {
    pub fork_id: String,
    pub checkpoint_id: String,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct RestoreCheckpointResponse {
    pub fork_id: String,
    pub restored_checkpoint: CheckpointInfo,
}

pub async fn restore_checkpoint(
    state: Arc<AppState>,
    params: RestoreCheckpointParams,
) -> Result<RestoreCheckpointResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let checkpoint = registry.restore_checkpoint(&params.fork_id, &params.checkpoint_id)?;
    let fork_workbook_id = WorkbookId(params.fork_id.clone());
    let _ = state.close_workbook(&fork_workbook_id);

    Ok(RestoreCheckpointResponse {
        fork_id: params.fork_id,
        restored_checkpoint: CheckpointInfo {
            checkpoint_id: checkpoint.checkpoint_id,
            created_at: checkpoint.created_at.to_rfc3339(),
            label: checkpoint.label,
        },
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DeleteCheckpointParams {
    pub fork_id: String,
    pub checkpoint_id: String,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct DeleteCheckpointResponse {
    pub fork_id: String,
    pub checkpoint_id: String,
    pub deleted: bool,
}

pub async fn delete_checkpoint(
    state: Arc<AppState>,
    params: DeleteCheckpointParams,
) -> Result<DeleteCheckpointResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    registry.delete_checkpoint(&params.fork_id, &params.checkpoint_id)?;

    Ok(DeleteCheckpointResponse {
        fork_id: params.fork_id,
        checkpoint_id: params.checkpoint_id,
        deleted: true,
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ListStagedChangesParams {
    pub fork_id: String,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct StagedChangeInfo {
    pub change_id: String,
    pub created_at: String,
    pub label: Option<String>,
    pub summary: ChangeSummary,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct ListStagedChangesResponse {
    pub fork_id: String,
    pub staged_changes: Vec<StagedChangeInfo>,
}

pub async fn list_staged_changes(
    state: Arc<AppState>,
    params: ListStagedChangesParams,
) -> Result<ListStagedChangesResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let staged = registry.list_staged_changes(&params.fork_id)?;
    let staged_changes = staged
        .into_iter()
        .map(|c| StagedChangeInfo {
            change_id: c.change_id,
            created_at: c.created_at.to_rfc3339(),
            label: c.label,
            summary: c.summary,
        })
        .collect();

    Ok(ListStagedChangesResponse {
        fork_id: params.fork_id,
        staged_changes,
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ApplyStagedChangeParams {
    pub fork_id: String,
    pub change_id: String,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct ApplyStagedChangeResponse {
    pub fork_id: String,
    pub change_id: String,
    pub ops_applied: usize,
    pub summary: ChangeSummary,
}

#[derive(Debug, Deserialize)]
struct EditBatchStagedPayload {
    sheet_name: String,
    edits: Vec<CellEdit>,
}

pub async fn apply_staged_change(
    state: Arc<AppState>,
    params: ApplyStagedChangeParams,
) -> Result<ApplyStagedChangeResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let staged_list = registry.list_staged_changes(&params.fork_id)?;
    let staged = staged_list
        .iter()
        .find(|c| c.change_id == params.change_id)
        .cloned()
        .ok_or_else(|| anyhow!("staged change not found: {}", params.change_id))?;

    let fork_ctx = registry.get_fork(&params.fork_id)?;
    let work_path = fork_ctx.work_path.clone();

    let mut ops_applied = 0usize;

    for op in &staged.ops {
        match op.kind.as_str() {
            "edit_batch" => {
                let payload: EditBatchStagedPayload = serde_json::from_value(op.payload.clone())
                    .map_err(|e| anyhow!("invalid edit_batch payload: {}", e))?;

                let edits_to_apply: Vec<_> = payload
                    .edits
                    .iter()
                    .map(|e| EditOp {
                        timestamp: Utc::now(),
                        sheet: payload.sheet_name.clone(),
                        address: e.address.clone(),
                        value: e.value.clone(),
                        is_formula: e.is_formula,
                    })
                    .collect();

                tokio::task::spawn_blocking({
                    let sheet_name = payload.sheet_name.clone();
                    let edits = payload.edits.clone();
                    let work_path = work_path.clone();
                    move || apply_edits_to_file(&work_path, &sheet_name, &edits)
                })
                .await??;

                registry.with_fork_mut(&params.fork_id, |ctx| {
                    ctx.edits.extend(edits_to_apply);
                    Ok(())
                })?;

                ops_applied += 1;
            }
            "style_batch" => {
                let payload: StyleBatchStagedPayload =
                    serde_json::from_value(op.payload.clone())
                        .map_err(|e| anyhow!("invalid style_batch payload: {}", e))?;

                tokio::task::spawn_blocking({
                    let ops = payload.ops.clone();
                    let work_path = work_path.clone();
                    move || apply_style_ops_to_file(&work_path, &ops)
                })
                .await??;

                ops_applied += 1;
            }
            "transform_batch" => {
                let payload: TransformBatchStagedPayload =
                    serde_json::from_value(op.payload.clone())
                        .map_err(|e| anyhow!("invalid transform_batch payload: {}", e))?;

                tokio::task::spawn_blocking({
                    let ops = payload.ops.clone();
                    let work_path = work_path.clone();
                    move || apply_transform_ops_to_file(&work_path, &ops)
                })
                .await??;

                ops_applied += 1;
            }
            "apply_formula_pattern" => {
                let payload: ApplyFormulaPatternStagedPayload =
                    serde_json::from_value(op.payload.clone())
                        .map_err(|e| anyhow!("invalid apply_formula_pattern payload: {}", e))?;

                let bounds = parse_range_bounds(&payload.target_range)?;
                let (anchor_col, anchor_row) = parse_cell_ref(&payload.anchor_cell)?;
                validate_formula_pattern_bounds(
                    &bounds,
                    anchor_col,
                    anchor_row,
                    payload.fill_direction.as_deref(),
                )?;
                let relative_mode = RelativeMode::parse(payload.relative_mode.as_deref())?;

                tokio::task::spawn_blocking({
                    let sheet_name = payload.sheet_name.clone();
                    let target_range = payload.target_range.clone();
                    let base_formula = payload.base_formula.clone();
                    let work_path = work_path.clone();
                    move || {
                        apply_formula_pattern_to_file(
                            &work_path,
                            &sheet_name,
                            &target_range,
                            anchor_col,
                            anchor_row,
                            &base_formula,
                            relative_mode,
                        )
                    }
                })
                .await??;

                ops_applied += 1;
            }
            "structure_batch" => {
                let payload: StructureBatchStagedPayload =
                    serde_json::from_value(op.payload.clone())
                        .map_err(|e| anyhow!("invalid structure_batch payload: {}", e))?;

                tokio::task::spawn_blocking({
                    let ops = payload.ops.clone();
                    let work_path = work_path.clone();
                    move || apply_structure_ops_to_file(&work_path, &ops)
                })
                .await??;

                ops_applied += 1;
            }
            other => {
                return Err(anyhow!("unsupported staged op kind: {}", other));
            }
        }
    }

    registry.discard_staged_change(&params.fork_id, &params.change_id)?;
    let fork_workbook_id = WorkbookId(params.fork_id.clone());
    let _ = state.close_workbook(&fork_workbook_id);

    Ok(ApplyStagedChangeResponse {
        fork_id: params.fork_id,
        change_id: params.change_id,
        ops_applied,
        summary: staged.summary,
    })
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DiscardStagedChangeParams {
    pub fork_id: String,
    pub change_id: String,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct DiscardStagedChangeResponse {
    pub fork_id: String,
    pub change_id: String,
    pub discarded: bool,
}

pub async fn discard_staged_change(
    state: Arc<AppState>,
    params: DiscardStagedChangeParams,
) -> Result<DiscardStagedChangeResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    registry.discard_staged_change(&params.fork_id, &params.change_id)?;

    Ok(DiscardStagedChangeResponse {
        fork_id: params.fork_id,
        change_id: params.change_id,
        discarded: true,
    })
}

const MAX_SCREENSHOT_ROWS: u32 = 100;
const MAX_SCREENSHOT_COLS: u32 = 30;
const DEFAULT_SCREENSHOT_RANGE: &str = "A1:M40";
const DEFAULT_MAX_PNG_DIM_PX: u32 = 4096;
const DEFAULT_MAX_PNG_AREA_PX: u64 = 12_000_000;

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ScreenshotSheetParams {
    #[serde(alias = "workbook_id")]
    pub workbook_or_fork_id: WorkbookId,
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

    let workbook = state.open_workbook(&params.workbook_or_fork_id).await?;
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

    let semaphore = state
        .screenshot_semaphore()
        .ok_or_else(|| anyhow!("screenshot semaphore not available"))?;

    // LibreOffice profile/macro export is not concurrency-safe. Serialize screenshot calls.
    let _permit = semaphore
        .0
        .acquire()
        .await
        .map_err(|e| anyhow!("failed to acquire screenshot permit: {}", e))?;

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
        workbook_id: params.workbook_or_fork_id.0,
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
    if parts.is_empty() || parts.len() > 2 {
        return Err(anyhow!("Invalid range format. Expected 'A1' or 'A1:Z99'"));
    }

    let start = parse_cell_ref(parts[0])?;
    let end = if parts.len() == 2 {
        parse_cell_ref(parts[1])?
    } else {
        start
    };

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
