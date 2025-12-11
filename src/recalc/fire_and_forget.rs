use super::RecalcConfig;
use super::executor::{RecalcExecutor, RecalcResult};
use anyhow::{Result, anyhow};
use async_trait::async_trait;
use std::path::{Path, PathBuf};
use std::process::Stdio;
use std::time::{Duration, Instant};
use tokio::fs;
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

        let profile_dir = format!("/tmp/lo-profile-{}", uuid::Uuid::new_v4());
        fs::create_dir_all(&profile_dir)
            .await
            .map_err(|e| anyhow!("failed to create profile dir: {}", e))?;

        // Seed profile with macro security + shipped macros so the macro call works from a clean profile.
        let profile_basic = format!("{}/user/basic/Standard", profile_dir);
        fs::create_dir_all(&profile_basic)
            .await
            .map_err(|e| anyhow!("failed to create profile basic dir: {}", e))?;
        fs::copy(
            "/etc/libreoffice/4/user/basic/Standard/Module1.xba",
            format!("{}/Module1.xba", profile_basic),
        )
        .await
        .map_err(|e| anyhow!("failed to copy Module1.xba into temp profile: {}", e))?;
        fs::copy(
            "/etc/libreoffice/4/user/basic/Standard/script.xlb",
            format!("{}/script.xlb", profile_basic),
        )
        .await
        .map_err(|e| anyhow!("failed to copy script.xlb into temp profile: {}", e))?;
        fs::copy(
            "/etc/libreoffice/4/user/registrymodifications.xcu",
            format!("{}/user/registrymodifications.xcu", profile_dir),
        )
        .await
        .map_err(|e| {
            anyhow!(
                "failed to copy registrymodifications.xcu into temp profile: {}",
                e
            )
        })?;

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
                    &format!("-env:UserInstallation=file://{}", profile_dir),
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

        let _ = fs::remove_dir_all(&profile_dir).await;

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
