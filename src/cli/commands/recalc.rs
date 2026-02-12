use crate::runtime::stateless::StatelessRuntime;
use anyhow::Result;
use serde::Serialize;
use serde_json::Value;
use std::path::PathBuf;

#[derive(Debug, Serialize)]
struct RecalculateResponse {
    file: String,
    backend: String,
    duration_ms: u64,
    cells_evaluated: Option<u64>,
    eval_errors: Option<Vec<String>>,
}

pub async fn recalculate(file: PathBuf) -> Result<Value> {
    let runtime = StatelessRuntime;
    let file = runtime.normalize_existing_file(&file)?;
    let outcome = runtime.recalculate_file(&file).await?;
    Ok(serde_json::to_value(RecalculateResponse {
        file: file.display().to_string(),
        backend: outcome.backend,
        duration_ms: outcome.duration_ms,
        cells_evaluated: outcome.cells_evaluated,
        eval_errors: outcome.eval_errors,
    })?)
}
