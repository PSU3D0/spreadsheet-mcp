use super::executor::{RecalcExecutor, RecalcResult};
use super::RecalcConfig;
use anyhow::{anyhow, Result};
use async_trait::async_trait;
use std::path::{Path, PathBuf};
use std::process::Stdio;
use std::time::{Duration, Instant};
use tokio::process::Command;

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

        let profile_dir = format!("/tmp/lo-profile-{}", uuid::Uuid::new_v4());
        let abs_path = workbook_path
            .canonicalize()
            .map_err(|e| anyhow!("failed to canonicalize path: {}", e))?;

        let macro_uri =
            "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application";

        let output = tokio::time::timeout(
            self.timeout,
            Command::new(&self.soffice_path)
                .args([
                    "--headless",
                    "--norestore",
                    &format!("-env:UserInstallation=file://{}", profile_dir),
                    macro_uri,
                    abs_path.to_str().unwrap(),
                ])
                .stdout(Stdio::piped())
                .stderr(Stdio::piped())
                .output(),
        )
        .await
        .map_err(|_| anyhow!("soffice timed out after {:?}", self.timeout))?
        .map_err(|e| anyhow!("failed to spawn soffice: {}", e))?;

        let _ = tokio::fs::remove_dir_all(&profile_dir).await;

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
