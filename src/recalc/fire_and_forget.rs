use super::RecalcConfig;
use super::executor::{RecalcExecutor, RecalcResult};
use anyhow::{Result, anyhow};
use async_trait::async_trait;
use std::path::{Path, PathBuf};
use std::process::Stdio;
use std::time::{Duration, Instant};
use tokio::process::Command;
use tokio::time;

pub struct FireAndForgetExecutor {
    soffice_path: PathBuf,
    timeout: Duration,
}

impl FireAndForgetExecutor {
    pub fn new(config: &RecalcConfig) -> Self {
        Self {
            soffice_path: config
                .soffice_path
                .clone()
                .unwrap_or_else(|| PathBuf::from("/usr/bin/soffice")),
            timeout: Duration::from_millis(config.timeout_ms.unwrap_or(30_000)),
        }
    }
}

#[async_trait]
impl RecalcExecutor for FireAndForgetExecutor {
    async fn recalculate(&self, workbook_path: &Path) -> Result<RecalcResult> {
        let start = Instant::now();

        let abs_path = workbook_path
            .canonicalize()
            .map_err(|e| anyhow!("failed to canonicalize path: {}", e))?;

        let file_url = format!("file://{}", abs_path.to_str().unwrap());
        let macro_uri = format!(
            "macro:///Standard.Module1.RecalculateAndSave(\"{}\")",
            file_url
        );

        let output_result = time::timeout(
            self.timeout,
            Command::new(&self.soffice_path)
                .args([
                    "--headless",
                    "--norestore",
                    "--nodefault",
                    "--nofirststartwizard",
                    "--nolockcheck",
                    "--calc",
                    &macro_uri,
                ])
                .stdout(Stdio::piped())
                .stderr(Stdio::piped())
                .output(),
        )
        .await
        .map_err(|_| anyhow!("soffice timed out after {:?}", self.timeout))
        .and_then(|res| res.map_err(|e| anyhow!("failed to spawn soffice: {}", e)));

        let output = output_result?;

        if !output.status.success() {
            let stderr = String::from_utf8_lossy(&output.stderr);
            let stdout = String::from_utf8_lossy(&output.stdout);
            return Err(anyhow!(
                "soffice failed (exit {}): stderr={}, stdout={}",
                output.status.code().unwrap_or(-1),
                stderr,
                stdout
            ));
        }

        Ok(RecalcResult {
            duration_ms: start.elapsed().as_millis() as u64,
            was_warm: false,
            executor_type: "fire_and_forget",
        })
    }

    fn is_available(&self) -> bool {
        self.soffice_path.exists()
    }
}
