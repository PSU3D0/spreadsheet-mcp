use crate::runtime::stateless::StatelessRuntime;
use anyhow::{Result, anyhow, bail};
use serde::Serialize;
use serde_json::Value;
use std::fs;
use std::path::{Path, PathBuf};
use tempfile::Builder;

#[derive(Debug, Serialize)]
struct RecalculateResponse {
    file: String,
    backend: String,
    duration_ms: u64,
    cells_evaluated: Option<u64>,
    eval_errors: Option<Vec<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    source_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    target_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    changed: Option<bool>,
}

pub async fn recalculate(file: PathBuf, output: Option<PathBuf>, force: bool) -> Result<Value> {
    if force && output.is_none() {
        bail!("invalid argument: --force requires --output <PATH>");
    }

    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;

    match output {
        None => {
            // In-place mode (existing behavior)
            let outcome = runtime.recalculate_file(&source).await?;
            Ok(serde_json::to_value(RecalculateResponse {
                file: source.display().to_string(),
                backend: outcome.backend,
                duration_ms: outcome.duration_ms,
                cells_evaluated: outcome.cells_evaluated,
                eval_errors: outcome.eval_errors,
                source_path: None,
                target_path: None,
                changed: None,
            })?)
        }
        Some(output_path) => {
            // Output mode: copy source to a temp file in the target directory,
            // recalculate temp, then atomically move into place. This keeps an
            // existing target untouched if recalc fails.
            let target = runtime.normalize_destination_path(&output_path)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let target_exists = target.exists();
            if target_exists && !force {
                bail!(
                    "output exists: output path '{}' already exists",
                    target.display()
                );
            }

            let target_parent = target.parent().unwrap_or_else(|| Path::new("."));
            let temp_file = Builder::new()
                .prefix(".recalculate-")
                .suffix(".xlsx")
                .tempfile_in(target_parent)
                .map_err(|error| {
                    anyhow!(
                        "write failed: unable to create temp output in '{}': {}",
                        target_parent.display(),
                        error
                    )
                })?;
            let temp_path = temp_file.path().to_path_buf();

            runtime.copy_file(&source, &temp_path).map_err(|error| {
                anyhow!(
                    "write failed: unable to copy workbook from '{}' to '{}': {}",
                    source.display(),
                    target.display(),
                    error
                )
            })?;

            let outcome = runtime.recalculate_file(&temp_path).await?;

            if target_exists {
                fs::remove_file(&target).map_err(|error| {
                    anyhow!(
                        "write failed: unable to remove existing output '{}': {}",
                        target.display(),
                        error
                    )
                })?;
            }

            temp_file.persist(&target).map_err(|error| {
                anyhow!(
                    "write failed: unable to persist recalculated output to '{}': {}",
                    target.display(),
                    error.error
                )
            })?;

            Ok(serde_json::to_value(RecalculateResponse {
                file: target.display().to_string(),
                backend: outcome.backend,
                duration_ms: outcome.duration_ms,
                cells_evaluated: outcome.cells_evaluated,
                eval_errors: outcome.eval_errors,
                source_path: Some(source.display().to_string()),
                target_path: Some(target.display().to_string()),
                changed: Some(true),
            })?)
        }
    }
}

fn ensure_output_path_is_distinct(source: &Path, output: &Path) -> Result<()> {
    let source_identity = canonical_identity_path(source)?;
    let output_identity = canonical_identity_path(output)?;
    if source_identity == output_identity {
        bail!("invalid argument: --output path resolves to the same file as input");
    }
    Ok(())
}

fn canonical_identity_path(path: &Path) -> Result<PathBuf> {
    if path.exists() {
        return fs::canonicalize(path).map_err(|e| {
            anyhow!(
                "failed to resolve canonical identity path for '{}': {}",
                path.display(),
                e
            )
        });
    }

    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    let name = path
        .file_name()
        .ok_or_else(|| anyhow!("invalid argument: output path must include a file name"))?;

    let parent_canonical = fs::canonicalize(parent).map_err(|_| {
        anyhow!(
            "invalid argument: output parent directory '{}' does not exist or is inaccessible",
            parent.display()
        )
    })?;

    Ok(parent_canonical.join(name))
}
