use crate::core::types::CellEdit;
use crate::model::Warning;
use crate::runtime::stateless::StatelessRuntime;
use crate::tools::fork::{
    ApplyFormulaPatternOpInput, ColumnSizeOp, ColumnSizeOpInput, StructureBatchParamsInput,
    StructureOp, StructureOpInput, StyleBatchParamsInput, StyleOp, StyleOpInput, TransformOp,
    apply_column_size_ops_to_file, apply_formula_pattern_ops_to_file, apply_structure_ops_to_file,
    apply_style_ops_to_file, apply_transform_ops_to_file, normalize_column_size_payload,
    normalize_structure_batch, normalize_style_batch, resolve_style_ops_for_workbook,
    resolve_transform_ops_for_workbook,
};
use crate::tools::rules_batch::{RulesOp, apply_rules_ops_to_file};
use crate::tools::sheet_layout::{SheetLayoutOp, apply_sheet_layout_ops_to_file};
use anyhow::{Context, Result, anyhow, bail};
use serde::{Deserialize, Serialize, de::DeserializeOwned};
use serde_json::Value;
use std::collections::BTreeMap;
use std::fs::{self, File, OpenOptions};
use std::io::ErrorKind;
use std::path::{Path, PathBuf};
use tempfile::{Builder, TempPath};

#[derive(Debug, Serialize)]
struct CopyResponse {
    source: String,
    dest: String,
    bytes_copied: u64,
}

#[derive(Debug, Serialize)]
struct EditResponse {
    file: String,
    sheet: String,
    edits_applied: usize,
    recalc_needed: bool,
    warnings: Vec<Warning>,
}

#[derive(Debug, Deserialize)]
struct OpsPayload<T> {
    ops: Vec<T>,
}

#[derive(Debug, Deserialize)]
struct ColumnSizeOpsPayload {
    sheet_name: String,
    ops: Vec<ColumnSizeOpInput>,
}

#[derive(Debug)]
enum BatchMutationMode {
    DryRun,
    InPlace,
    Output { target: PathBuf, force: bool },
}

#[derive(Debug, Serialize)]
struct DryRunSummary {
    operation_counts: BTreeMap<String, u64>,
    result_counts: BTreeMap<String, u64>,
}

#[derive(Debug, Serialize)]
struct BatchDryRunResponse {
    op_count: usize,
    validated_count: usize,
    would_change: bool,
    warnings: Vec<Warning>,
    summary: DryRunSummary,
}

#[derive(Debug, Serialize)]
struct BatchApplyResponse {
    op_count: usize,
    applied_count: usize,
    warnings: Vec<Warning>,
    changed: bool,
    target_path: String,
    source_path: String,
}

pub async fn copy(source: PathBuf, dest: PathBuf) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&source)?;
    let dest = runtime.normalize_destination_path(&dest)?;
    let bytes_copied = runtime.copy_file(&source, &dest).with_context(|| {
        format!(
            "failed to copy workbook from '{}' to '{}'",
            source.display(),
            dest.display()
        )
    })?;

    Ok(serde_json::to_value(CopyResponse {
        source: source.display().to_string(),
        dest: dest.display().to_string(),
        bytes_copied,
    })?)
}

pub async fn edit(file: PathBuf, sheet: String, edits: Vec<String>) -> Result<Value> {
    if edits.is_empty() {
        bail!("at least one edit must be provided");
    }

    let runtime = StatelessRuntime;
    let file = runtime.normalize_existing_file(&file)?;

    let mut normalized_edits = Vec::with_capacity(edits.len());
    let mut warnings = Vec::new();
    for (idx, entry) in edits.into_iter().enumerate() {
        let (edit, entry_warnings) = crate::core::write::normalize_shorthand_edit(&entry)
            .with_context(|| format!("invalid shorthand edit at index {}", idx))?;
        normalized_edits.push(edit);
        warnings.extend(entry_warnings.into_iter().map(|warning| Warning {
            code: warning.code,
            message: warning.message,
        }));
    }

    runtime.apply_edits(&file, &sheet, &normalized_edits)?;

    Ok(serde_json::to_value(EditResponse {
        file: file.display().to_string(),
        sheet,
        edits_applied: normalized_edits.len(),
        recalc_needed: true,
        warnings,
    })?)
}

pub async fn transform_batch(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: OpsPayload<TransformOp> = parse_ops_payload(&ops, r#"{"ops":[...]}"#)?;

    let (state, workbook_id) = runtime.open_state_for_file(&source).await?;
    let workbook = state.open_workbook(&workbook_id).await?;
    let resolved_ops = resolve_transform_ops_for_workbook(&workbook, &payload.ops)
        .map_err(|error| invalid_ops_payload(error.to_string()))?;
    let _ = state.close_workbook(&workbook_id);

    let op_count = resolved_ops.len();
    let operation_counts = summarize_transform_operation_counts(&resolved_ops);

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) =
                apply_to_temp_copy(&source, source.parent(), ".transform-batch-", |path| {
                    apply_transform_ops_to_file(path, &resolved_ops).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let would_change = transform_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
            )
        }
        BatchMutationMode::InPlace => {
            let apply_result = apply_in_place_with_temp(&source, ".transform-batch-", |path| {
                apply_transform_ops_to_file(path, &resolved_ops).map_err(classify_apply_error)
            })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = transform_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let apply_result =
                apply_to_output_with_temp(&source, &target, force, ".transform-batch-", |path| {
                    apply_transform_ops_to_file(path, &resolved_ops).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = transform_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
            )
        }
    }
}

pub async fn style_batch(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: OpsPayload<StyleOpInput> = parse_ops_payload(&ops, r#"{"ops":[...]}"#)?;
    let (normalized, base_warnings) = normalize_style_batch(StyleBatchParamsInput {
        fork_id: String::new(),
        ops: payload.ops,
        mode: None,
        label: None,
    })
    .map_err(|error| invalid_ops_payload(error.to_string()))?;

    let (state, workbook_id) = runtime.open_state_for_file(&source).await?;
    let workbook = state.open_workbook(&workbook_id).await?;
    let resolved_ops = resolve_style_ops_for_workbook(&workbook, &normalized.ops)
        .map_err(|error| invalid_ops_payload(error.to_string()))?;
    let _ = state.close_workbook(&workbook_id);

    let op_count = resolved_ops.len();
    let operation_counts = summarize_style_operation_counts(&resolved_ops);

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) =
                apply_to_temp_copy(&source, source.parent(), ".style-batch-", |path| {
                    apply_style_ops_to_file(path, &resolved_ops).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings.clone(),
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let would_change = style_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
            )
        }
        BatchMutationMode::InPlace => {
            let apply_result = apply_in_place_with_temp(&source, ".style-batch-", |path| {
                apply_style_ops_to_file(path, &resolved_ops).map_err(classify_apply_error)
            })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings.clone(),
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let changed = style_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let apply_result =
                apply_to_output_with_temp(&source, &target, force, ".style-batch-", |path| {
                    apply_style_ops_to_file(path, &resolved_ops).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings,
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let changed = style_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
            )
        }
    }
}

pub async fn apply_formula_pattern(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: OpsPayload<ApplyFormulaPatternOpInput> =
        parse_ops_payload(&ops, r#"{"ops":[...]}"#)?;

    let op_count = payload.ops.len();
    let operation_counts = summarize_formula_pattern_operation_counts(&payload.ops);

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) = apply_to_temp_copy(
                &source,
                source.parent(),
                ".apply-formula-pattern-",
                |path| {
                    apply_formula_pattern_ops_to_file(path, &payload.ops)
                        .map_err(classify_apply_error)
                },
            )?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let would_change = formula_pattern_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
            )
        }
        BatchMutationMode::InPlace => {
            let apply_result =
                apply_in_place_with_temp(&source, ".apply-formula-pattern-", |path| {
                    apply_formula_pattern_ops_to_file(path, &payload.ops)
                        .map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = formula_pattern_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let apply_result = apply_to_output_with_temp(
                &source,
                &target,
                force,
                ".apply-formula-pattern-",
                |path| {
                    apply_formula_pattern_ops_to_file(path, &payload.ops)
                        .map_err(classify_apply_error)
                },
            )?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = formula_pattern_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
            )
        }
    }
}

pub async fn structure_batch(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: OpsPayload<StructureOpInput> = parse_ops_payload(&ops, r#"{"ops":[...]}"#)?;
    let (normalized, base_warnings) = normalize_structure_batch(StructureBatchParamsInput {
        fork_id: String::new(),
        ops: payload.ops,
        mode: None,
        label: None,
    })
    .map_err(|error| invalid_ops_payload(error.to_string()))?;

    let op_count = normalized.ops.len();
    let operation_counts = summarize_structure_operation_counts(&normalized.ops);

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) =
                apply_to_temp_copy(&source, source.parent(), ".structure-batch-", |path| {
                    apply_structure_ops_to_file(path, &normalized.ops).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings.clone(),
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let would_change = structure_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
            )
        }
        BatchMutationMode::InPlace => {
            let apply_result = apply_in_place_with_temp(&source, ".structure-batch-", |path| {
                apply_structure_ops_to_file(path, &normalized.ops).map_err(classify_apply_error)
            })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings.clone(),
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let changed = structure_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let apply_result =
                apply_to_output_with_temp(&source, &target, force, ".structure-batch-", |path| {
                    apply_structure_ops_to_file(path, &normalized.ops).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings,
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let changed = structure_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
            )
        }
    }
}

pub async fn column_size_batch(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: ColumnSizeOpsPayload =
        parse_ops_payload(&ops, r#"{"sheet_name":"...","ops":[...]}"#)?;
    let (normalized_ops, base_warnings) =
        normalize_column_size_payload(payload.sheet_name.clone(), payload.ops)
            .map_err(|error| invalid_ops_payload(error.to_string()))?;

    let op_count = normalized_ops.len();
    let operation_counts = summarize_column_size_operation_counts(&normalized_ops);

    match mode {
        BatchMutationMode::DryRun => {
            let sheet_name = payload.sheet_name.clone();
            let (apply_result, _temp_path) =
                apply_to_temp_copy(&source, source.parent(), ".column-size-batch-", |path| {
                    apply_column_size_ops_to_file(path, &sheet_name, &normalized_ops)
                        .map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings.clone(),
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let would_change = column_size_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
            )
        }
        BatchMutationMode::InPlace => {
            let sheet_name = payload.sheet_name.clone();
            let apply_result = apply_in_place_with_temp(&source, ".column-size-batch-", |path| {
                apply_column_size_ops_to_file(path, &sheet_name, &normalized_ops)
                    .map_err(classify_apply_error)
            })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings.clone(),
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let changed = column_size_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let sheet_name = payload.sheet_name;
            let apply_result = apply_to_output_with_temp(
                &source,
                &target,
                force,
                ".column-size-batch-",
                |path| {
                    apply_column_size_ops_to_file(path, &sheet_name, &normalized_ops)
                        .map_err(classify_apply_error)
                },
            )?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings,
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let changed = column_size_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
            )
        }
    }
}

pub async fn sheet_layout_batch(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: OpsPayload<SheetLayoutOp> = parse_ops_payload(&ops, r#"{"ops":[...]}"#)?;

    let op_count = payload.ops.len();
    let operation_counts = summarize_sheet_layout_operation_counts(&payload.ops);

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) =
                apply_to_temp_copy(&source, source.parent(), ".sheet-layout-batch-", |path| {
                    apply_sheet_layout_ops_to_file(path, &payload.ops).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let would_change = sheet_layout_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
            )
        }
        BatchMutationMode::InPlace => {
            let apply_result = apply_in_place_with_temp(&source, ".sheet-layout-batch-", |path| {
                apply_sheet_layout_ops_to_file(path, &payload.ops).map_err(classify_apply_error)
            })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = sheet_layout_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let apply_result = apply_to_output_with_temp(
                &source,
                &target,
                force,
                ".sheet-layout-batch-",
                |path| {
                    apply_sheet_layout_ops_to_file(path, &payload.ops).map_err(classify_apply_error)
                },
            )?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = sheet_layout_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
            )
        }
    }
}

pub async fn rules_batch(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: OpsPayload<RulesOp> = parse_ops_payload(&ops, r#"{"ops":[...]}"#)?;

    let op_count = payload.ops.len();
    let operation_counts = summarize_rules_operation_counts(&payload.ops);

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) =
                apply_to_temp_copy(&source, source.parent(), ".rules-batch-", |path| {
                    apply_rules_ops_to_file(path, &payload.ops).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let would_change = rules_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
            )
        }
        BatchMutationMode::InPlace => {
            let apply_result = apply_in_place_with_temp(&source, ".rules-batch-", |path| {
                apply_rules_ops_to_file(path, &payload.ops).map_err(classify_apply_error)
            })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = rules_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let apply_result =
                apply_to_output_with_temp(&source, &target, force, ".rules-batch-", |path| {
                    apply_rules_ops_to_file(path, &payload.ops).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = rules_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
            )
        }
    }
}

fn validate_batch_mode(
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<BatchMutationMode> {
    if force && output.is_none() {
        return Err(invalid_argument("--force requires --output <PATH>"));
    }

    if dry_run {
        if in_place {
            return Err(invalid_argument(
                "--dry-run cannot be combined with --in-place",
            ));
        }
        if output.is_some() {
            return Err(invalid_argument(
                "--dry-run cannot be combined with --output <PATH>",
            ));
        }
        return Ok(BatchMutationMode::DryRun);
    }

    if in_place && output.is_some() {
        return Err(invalid_argument(
            "--in-place cannot be combined with --output <PATH>",
        ));
    }

    if in_place {
        return Ok(BatchMutationMode::InPlace);
    }

    if let Some(target) = output {
        return Ok(BatchMutationMode::Output { target, force });
    }

    Err(invalid_argument(
        "choose exactly one mutation mode: --dry-run, --in-place, or --output <PATH>",
    ))
}

fn parse_ops_payload<T: DeserializeOwned>(raw: &str, schema_hint: &str) -> Result<T> {
    let path = raw
        .strip_prefix('@')
        .ok_or_else(|| invalid_ops_payload("--ops must be provided as @<path>"))?;
    if path.is_empty() {
        return Err(invalid_ops_payload(
            "--ops file reference cannot be empty; expected @<path>",
        ));
    }

    let raw_payload = fs::read_to_string(path).map_err(|error| {
        invalid_ops_payload(format!("unable to read ops payload '{}': {}", path, error))
    })?;

    let json_value: serde_json::Value = serde_json::from_str(&raw_payload).map_err(|error| {
        invalid_ops_payload(format!("ops payload is not valid JSON: {}", error))
    })?;

    if !json_value.is_object() {
        return Err(invalid_ops_payload(
            "ops payload must be a JSON object with top-level key 'ops'",
        ));
    }

    serde_json::from_value(json_value).map_err(|error| {
        invalid_ops_payload(format!(
            "ops payload does not match required schema {}: {}",
            schema_hint, error
        ))
    })
}

fn summarize_transform_operation_counts(ops: &[TransformOp]) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    for op in ops {
        let key = match op {
            TransformOp::ClearRange { .. } => "clear_range",
            TransformOp::FillRange { .. } => "fill_range",
            TransformOp::ReplaceInRange { .. } => "replace_in_range",
        };
        *counts.entry(key.to_string()).or_insert(0) += 1;
    }
    counts
}

fn summarize_style_operation_counts(ops: &[StyleOp]) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    counts.insert("style_ops".to_string(), ops.len() as u64);
    counts
}

fn summarize_formula_pattern_operation_counts(
    ops: &[ApplyFormulaPatternOpInput],
) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    counts.insert("apply_formula_pattern_ops".to_string(), ops.len() as u64);
    counts
}

fn summarize_structure_operation_counts(ops: &[StructureOp]) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    for op in ops {
        let key = match op {
            StructureOp::InsertRows { .. } => "insert_rows",
            StructureOp::DeleteRows { .. } => "delete_rows",
            StructureOp::InsertCols { .. } => "insert_cols",
            StructureOp::DeleteCols { .. } => "delete_cols",
            StructureOp::RenameSheet { .. } => "rename_sheet",
            StructureOp::CreateSheet { .. } => "create_sheet",
            StructureOp::DeleteSheet { .. } => "delete_sheet",
            StructureOp::CopyRange { .. } => "copy_range",
            StructureOp::MoveRange { .. } => "move_range",
        };
        *counts.entry(key.to_string()).or_insert(0) += 1;
    }
    counts
}

fn summarize_column_size_operation_counts(ops: &[ColumnSizeOp]) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    for op in ops {
        let key = match op.size {
            crate::tools::fork::ColumnSizeSpec::Auto { .. } => "auto",
            crate::tools::fork::ColumnSizeSpec::Width { .. } => "width",
        };
        *counts.entry(key.to_string()).or_insert(0) += 1;
    }
    counts
}

fn summarize_sheet_layout_operation_counts(ops: &[SheetLayoutOp]) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    for op in ops {
        let key = match op {
            SheetLayoutOp::FreezePanes { .. } => "freeze_panes",
            SheetLayoutOp::SetZoom { .. } => "set_zoom",
            SheetLayoutOp::SetGridlines { .. } => "set_gridlines",
            SheetLayoutOp::SetPageMargins { .. } => "set_page_margins",
            SheetLayoutOp::SetPageSetup { .. } => "set_page_setup",
            SheetLayoutOp::SetPrintArea { .. } => "set_print_area",
            SheetLayoutOp::SetPageBreaks { .. } => "set_page_breaks",
        };
        *counts.entry(key.to_string()).or_insert(0) += 1;
    }
    counts
}

fn summarize_rules_operation_counts(ops: &[RulesOp]) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    for op in ops {
        let key = match op {
            RulesOp::SetDataValidation { .. } => "set_data_validation",
            RulesOp::AddConditionalFormat { .. } => "add_conditional_format",
            RulesOp::SetConditionalFormat { .. } => "set_conditional_format",
            RulesOp::ClearConditionalFormats { .. } => "clear_conditional_formats",
        };
        *counts.entry(key.to_string()).or_insert(0) += 1;
    }
    counts
}

fn transform_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    const CHANGE_KEYS: &[&str] = &[
        "cells_value_cleared",
        "cells_formula_cleared",
        "cells_value_set",
        "cells_formula_set",
        "cells_value_replaced",
        "cells_formula_replaced",
    ];
    any_count_non_zero(counts, CHANGE_KEYS)
}

fn style_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    any_count_non_zero(counts, &["cells_style_changed"])
}

fn formula_pattern_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    any_count_non_zero(counts, &["cells_filled"])
}

fn structure_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    any_count_non_zero(
        counts,
        &[
            "rows_inserted",
            "rows_deleted",
            "cols_inserted",
            "cols_deleted",
            "sheets_renamed",
            "sheets_created",
            "sheets_deleted",
            "cells_copied",
            "cells_moved",
            "ranges_copied",
            "ranges_moved",
        ],
    )
}

fn column_size_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    any_count_non_zero(counts, &["columns_sized"])
}

fn sheet_layout_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    any_count_non_zero(
        counts,
        &[
            "ops",
            "freeze_panes_ops",
            "set_zoom_ops",
            "set_gridlines_ops",
            "set_page_margins_ops",
            "set_page_setup_ops",
            "set_print_area_ops",
            "set_page_breaks_ops",
        ],
    )
}

fn rules_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    any_count_non_zero(
        counts,
        &[
            "validations_set",
            "validations_replaced",
            "conditional_formats_added",
            "conditional_formats_set",
            "conditional_formats_replaced",
            "conditional_formats_cleared",
        ],
    )
}

fn any_count_non_zero(counts: &BTreeMap<String, u64>, keys: &[&str]) -> bool {
    keys.iter()
        .any(|key| counts.get(*key).copied().unwrap_or(0) > 0)
}

fn warning_strings_to_cli_warnings(messages: Vec<String>) -> Vec<Warning> {
    messages.into_iter().map(parse_warning_message).collect()
}

fn merge_cli_warnings(mut left: Vec<Warning>, mut right: Vec<Warning>) -> Vec<Warning> {
    left.append(&mut right);
    left
}

fn parse_warning_message(message: String) -> Warning {
    if let Some((code, detail)) = message.split_once(':') {
        let code = code.trim();
        let detail = detail.trim();
        if is_warning_code(code) && !detail.is_empty() {
            return Warning {
                code: code.to_string(),
                message: detail.to_string(),
            };
        }
    }

    Warning {
        code: "WARN_INFO".to_string(),
        message,
    }
}

fn is_warning_code(value: &str) -> bool {
    value.starts_with("WARN_")
        && value
            .chars()
            .all(|ch| ch.is_ascii_uppercase() || ch == '_' || ch.is_ascii_digit())
}

fn dry_run_response(
    op_count: usize,
    operation_counts: BTreeMap<String, u64>,
    result_counts: BTreeMap<String, u64>,
    warnings: Vec<Warning>,
    would_change: bool,
) -> Result<Value> {
    Ok(serde_json::to_value(BatchDryRunResponse {
        op_count,
        validated_count: op_count,
        would_change,
        warnings,
        summary: DryRunSummary {
            operation_counts,
            result_counts,
        },
    })?)
}

fn apply_response(
    op_count: usize,
    applied_count: usize,
    warnings: Vec<Warning>,
    changed: bool,
    target_path: String,
    source_path: String,
) -> Result<Value> {
    Ok(serde_json::to_value(BatchApplyResponse {
        op_count,
        applied_count,
        warnings,
        changed,
        target_path,
        source_path,
    })?)
}

fn apply_in_place_with_temp<T, F>(source: &Path, temp_prefix: &str, apply_fn: F) -> Result<T>
where
    F: FnOnce(&Path) -> Result<T>,
{
    let (apply_result, temp_path) =
        apply_to_temp_copy(source, source.parent(), temp_prefix, apply_fn)?;
    atomic_replace_target(temp_path, source, true)?;
    Ok(apply_result)
}

fn apply_to_output_with_temp<T, F>(
    source: &Path,
    target: &Path,
    force: bool,
    temp_prefix: &str,
    apply_fn: F,
) -> Result<T>
where
    F: FnOnce(&Path) -> Result<T>,
{
    let target_exists = path_entry_exists(target)?;
    if target_exists && !force {
        return Err(output_exists(format!(
            "output path '{}' already exists",
            target.display()
        )));
    }

    let (apply_result, temp_path) =
        apply_to_temp_copy(source, target.parent(), temp_prefix, apply_fn)?;
    atomic_replace_target(temp_path, target, force)?;
    Ok(apply_result)
}

fn apply_to_temp_copy<T, F>(
    source: &Path,
    directory: Option<&Path>,
    temp_prefix: &str,
    apply_fn: F,
) -> Result<(T, TempPath)>
where
    F: FnOnce(&Path) -> Result<T>,
{
    let parent = directory.ok_or_else(|| {
        write_failed(format!(
            "unable to create temp file: '{}' has no parent directory",
            source.display()
        ))
    })?;
    let temp_path = Builder::new()
        .prefix(temp_prefix)
        .suffix(".tmp.xlsx")
        .tempfile_in(parent)
        .map_err(|error| {
            write_failed(format!(
                "unable to allocate temp file in '{}': {}",
                parent.display(),
                error
            ))
        })?
        .into_temp_path();

    let temp_path_ref: &Path = temp_path.as_ref();

    fs::copy(source, temp_path_ref).map_err(|error| {
        write_failed(format!(
            "unable to stage temp workbook from '{}' to '{}': {}",
            source.display(),
            temp_path.display(),
            error
        ))
    })?;

    let apply_result = apply_fn(temp_path_ref)?;

    fsync_file(temp_path_ref)?;

    Ok((apply_result, temp_path))
}

fn atomic_replace_target(temp_path: TempPath, target: &Path, allow_overwrite: bool) -> Result<()> {
    if allow_overwrite {
        let target_exists = path_entry_exists(target)?;
        if target_exists && !atomic_overwrite_supported() {
            return Err(write_failed(
                "atomic overwrite is not supported on this platform",
            ));
        }

        let temp_path_ref: &Path = temp_path.as_ref();
        fs::rename(temp_path_ref, target).map_err(|error| {
            write_failed(format!(
                "unable to atomically replace '{}' from '{}': {}",
                target.display(),
                temp_path.display(),
                error
            ))
        })?;
    } else {
        temp_path.persist_noclobber(target).map_err(|error| {
            if error.error.kind() == ErrorKind::AlreadyExists {
                output_exists(format!("output path '{}' already exists", target.display()))
            } else {
                write_failed(format!(
                    "unable to move staged workbook '{}' to '{}': {}",
                    error.path.display(),
                    target.display(),
                    error.error
                ))
            }
        })?;
    }

    if let Some(parent) = target.parent() {
        fsync_directory(parent)?;
    }

    Ok(())
}

fn fsync_file(path: &Path) -> Result<()> {
    let file = OpenOptions::new()
        .read(true)
        .write(true)
        .open(path)
        .map_err(|error| {
            write_failed(format!(
                "unable to open '{}' for fsync: {}",
                path.display(),
                error
            ))
        })?;
    file.sync_all().map_err(|error| {
        write_failed(format!(
            "unable to fsync temp file '{}': {}",
            path.display(),
            error
        ))
    })
}

#[cfg(unix)]
fn fsync_directory(path: &Path) -> Result<()> {
    let dir = File::open(path).map_err(|error| {
        write_failed(format!(
            "unable to open directory '{}' for fsync: {}",
            path.display(),
            error
        ))
    })?;
    dir.sync_all().map_err(|error| {
        write_failed(format!(
            "unable to fsync directory '{}': {}",
            path.display(),
            error
        ))
    })
}

#[cfg(not(unix))]
fn fsync_directory(_path: &Path) -> Result<()> {
    Ok(())
}

fn path_entry_exists(path: &Path) -> Result<bool> {
    match fs::symlink_metadata(path) {
        Ok(_) => Ok(true),
        Err(error) if error.kind() == ErrorKind::NotFound => Ok(false),
        Err(error) => Err(write_failed(format!(
            "unable to inspect output path '{}': {}",
            path.display(),
            error
        ))),
    }
}

fn ensure_output_path_is_distinct(source: &Path, output: &Path) -> Result<()> {
    let source_identity = canonical_identity_path(source)?;
    let output_identity = canonical_identity_path(output)?;
    if source_identity == output_identity {
        return Err(invalid_argument(
            "--output path resolves to the same file as input",
        ));
    }
    Ok(())
}

fn canonical_identity_path(path: &Path) -> Result<PathBuf> {
    if path.exists() {
        return fs::canonicalize(path).with_context(|| {
            format!(
                "failed to resolve canonical identity path for '{}'",
                path.display()
            )
        });
    }

    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    let name = path
        .file_name()
        .ok_or_else(|| invalid_argument("output path must include a file name"))?;

    let parent_canonical = fs::canonicalize(parent).with_context(|| {
        format!(
            "failed to resolve output parent directory '{}': {}",
            parent.display(),
            "directory does not exist or is inaccessible"
        )
    })?;

    Ok(parent_canonical.join(name))
}

#[cfg(unix)]
fn atomic_overwrite_supported() -> bool {
    true
}

#[cfg(not(unix))]
fn atomic_overwrite_supported() -> bool {
    false
}

fn classify_apply_error(error: anyhow::Error) -> anyhow::Error {
    if error
        .chain()
        .any(|cause| cause.downcast_ref::<std::io::Error>().is_some())
    {
        write_failed(format!("failed while applying ops payload: {}", error))
    } else {
        invalid_ops_payload(format!("{}", error))
    }
}

fn invalid_argument(message: impl AsRef<str>) -> anyhow::Error {
    anyhow!("invalid argument: {}", message.as_ref())
}

fn invalid_ops_payload(message: impl AsRef<str>) -> anyhow::Error {
    anyhow!("invalid ops payload: {}", message.as_ref())
}

fn output_exists(message: impl AsRef<str>) -> anyhow::Error {
    anyhow!("output exists: {}", message.as_ref())
}

fn write_failed(message: impl AsRef<str>) -> anyhow::Error {
    anyhow!("write failed: {}", message.as_ref())
}

pub fn parse_shorthand_for_tests(entries: Vec<String>) -> Result<(Vec<CellEdit>, Vec<Warning>)> {
    let mut edits = Vec::with_capacity(entries.len());
    let mut warnings = Vec::new();
    for entry in entries {
        let (edit, entry_warnings) = crate::core::write::normalize_shorthand_edit(&entry)?;
        edits.push(edit);
        warnings.extend(entry_warnings.into_iter().map(|warning| Warning {
            code: warning.code,
            message: warning.message,
        }));
    }
    Ok((edits, warnings))
}
