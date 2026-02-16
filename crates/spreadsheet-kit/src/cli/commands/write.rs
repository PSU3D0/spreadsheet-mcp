use crate::core::types::CellEdit;
use crate::model::Warning;
use crate::runtime::stateless::StatelessRuntime;
use crate::tools::fork::{
    TransformApplyResult, TransformOp, apply_transform_ops_to_file,
    resolve_transform_ops_for_workbook,
};
use anyhow::{Context, Result, anyhow, bail};
use serde::{Deserialize, Serialize};
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
struct TransformOpsPayload {
    ops: Vec<TransformOp>,
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
struct TransformBatchDryRunResponse {
    op_count: usize,
    validated_count: usize,
    would_change: bool,
    warnings: Vec<Warning>,
    summary: DryRunSummary,
}

#[derive(Debug, Serialize)]
struct TransformBatchApplyResponse {
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

    let payload = parse_ops_payload(&ops)?;

    let (state, workbook_id) = runtime.open_state_for_file(&source).await?;
    let workbook = state.open_workbook(&workbook_id).await?;
    let resolved_ops = resolve_transform_ops_for_workbook(&workbook, &payload.ops)
        .map_err(|error| invalid_ops_payload(error.to_string()))?;
    let _ = state.close_workbook(&workbook_id);

    let op_count = resolved_ops.len();
    let operation_counts = summarize_operation_counts(&resolved_ops);

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) =
                apply_ops_to_temp_copy(&source, source.parent(), &resolved_ops)?;
            let would_change = summary_indicates_change(&apply_result.summary.counts);

            Ok(serde_json::to_value(TransformBatchDryRunResponse {
                op_count,
                validated_count: op_count,
                would_change,
                warnings: Vec::new(),
                summary: DryRunSummary {
                    operation_counts,
                    result_counts: apply_result.summary.counts,
                },
            })?)
        }
        BatchMutationMode::InPlace => {
            let apply_result = apply_ops_in_place(&source, &resolved_ops)?;
            let changed = summary_indicates_change(&apply_result.summary.counts);

            Ok(serde_json::to_value(TransformBatchApplyResponse {
                op_count,
                applied_count: apply_result.ops_applied,
                warnings: Vec::new(),
                changed,
                target_path: source.display().to_string(),
                source_path: source.display().to_string(),
            })?)
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;
            if path_entry_exists(&target)? && !force {
                return Err(output_exists(format!(
                    "output path '{}' already exists",
                    target.display()
                )));
            }

            let apply_result = apply_ops_to_output(&source, &target, force, &resolved_ops)?;
            let changed = summary_indicates_change(&apply_result.summary.counts);

            Ok(serde_json::to_value(TransformBatchApplyResponse {
                op_count,
                applied_count: apply_result.ops_applied,
                warnings: Vec::new(),
                changed,
                target_path: target.display().to_string(),
                source_path: source.display().to_string(),
            })?)
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

fn parse_ops_payload(raw: &str) -> Result<TransformOpsPayload> {
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
            "ops payload does not match required schema {{\"ops\":[...]}}: {}",
            error
        ))
    })
}

fn summarize_operation_counts(ops: &[TransformOp]) -> BTreeMap<String, u64> {
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

fn summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    const CHANGE_KEYS: &[&str] = &[
        "cells_value_cleared",
        "cells_formula_cleared",
        "cells_value_set",
        "cells_formula_set",
        "cells_value_replaced",
        "cells_formula_replaced",
    ];

    CHANGE_KEYS
        .iter()
        .any(|key| counts.get(*key).copied().unwrap_or(0) > 0)
}

fn apply_ops_in_place(source: &Path, ops: &[TransformOp]) -> Result<TransformApplyResult> {
    let (apply_result, temp_path) = apply_ops_to_temp_copy(source, source.parent(), ops)?;
    atomic_replace_target(temp_path, source, true)?;
    Ok(apply_result)
}

fn apply_ops_to_output(
    source: &Path,
    target: &Path,
    force: bool,
    ops: &[TransformOp],
) -> Result<TransformApplyResult> {
    let target_exists = path_entry_exists(target)?;
    if target_exists && !force {
        return Err(output_exists(format!(
            "output path '{}' already exists",
            target.display()
        )));
    }

    let (apply_result, temp_path) = apply_ops_to_temp_copy(source, target.parent(), ops)?;
    atomic_replace_target(temp_path, target, force)?;
    Ok(apply_result)
}

fn apply_ops_to_temp_copy(
    source: &Path,
    directory: Option<&Path>,
    ops: &[TransformOp],
) -> Result<(TransformApplyResult, TempPath)> {
    let parent = directory.ok_or_else(|| {
        write_failed(format!(
            "unable to create temp file: '{}' has no parent directory",
            source.display()
        ))
    })?;
    let temp_path = Builder::new()
        .prefix(".transform-batch-")
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

    let apply_result =
        apply_transform_ops_to_file(temp_path_ref, ops).map_err(classify_apply_transform_error)?;

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

fn classify_apply_transform_error(error: anyhow::Error) -> anyhow::Error {
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
