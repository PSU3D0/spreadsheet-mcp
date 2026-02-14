use crate::runtime::stateless::StatelessRuntime;
use anyhow::Result;
use serde_json::Value;
use std::path::PathBuf;

pub async fn diff(original: PathBuf, modified: PathBuf) -> Result<Value> {
    let runtime = StatelessRuntime;
    let original = runtime.normalize_existing_file(&original)?;
    let modified = runtime.normalize_existing_file(&modified)?;
    runtime.diff_json(&original, &modified)
}
