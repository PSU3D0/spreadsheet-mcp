use super::executor::{RecalcExecutor, RecalcResult};
use anyhow::Result;
use async_trait::async_trait;
use std::path::Path;
use std::sync::Arc;

#[async_trait]
pub trait RecalcBackend: Send + Sync {
    async fn recalculate(&self, fork_work_path: &Path) -> Result<RecalcResult>;
    fn is_available(&self) -> bool;
    fn name(&self) -> &'static str;
}

pub struct LibreOfficeBackend {
    executor: Arc<dyn RecalcExecutor>,
}

impl LibreOfficeBackend {
    pub fn new(executor: Arc<dyn RecalcExecutor>) -> Self {
        Self { executor }
    }
}

#[async_trait]
impl RecalcBackend for LibreOfficeBackend {
    async fn recalculate(&self, fork_work_path: &Path) -> Result<RecalcResult> {
        self.executor.recalculate(fork_work_path).await
    }

    fn is_available(&self) -> bool {
        self.executor.is_available()
    }

    fn name(&self) -> &'static str {
        "libreoffice"
    }
}
