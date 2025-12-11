use anyhow::Result;
use async_trait::async_trait;
use std::path::Path;

#[async_trait]
pub trait RecalcExecutor: Send + Sync {
    async fn recalculate(&self, workbook_path: &Path) -> Result<RecalcResult>;
    fn is_available(&self) -> bool;
}

#[derive(Debug, Clone)]
pub struct RecalcResult {
    pub duration_ms: u64,
    pub was_warm: bool,
    pub executor_type: &'static str,
}
