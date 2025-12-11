use anyhow::{Context, Result, bail};
use bollard::Docker;
use bollard::image::ListImagesOptions;
use std::collections::HashMap;
use std::path::Path;
use std::process::Stdio;
use tokio::process::Command;

const IMAGE_HASH: &str = include_str!("../../docker/image.hash");
const IMAGE_NAME: &str = "gridbench-full";

pub fn image_tag() -> String {
    let hash = IMAGE_HASH.trim();
    format!("{}:{}", IMAGE_NAME, &hash[..12])
}

async fn image_exists(tag: &str) -> Result<bool> {
    let docker = Docker::connect_with_local_defaults().context("failed to connect to docker")?;

    let filters: HashMap<String, Vec<String>> =
        HashMap::from([("reference".to_string(), vec![tag.to_string()])]);

    let images = docker
        .list_images(Some(ListImagesOptions {
            filters,
            ..Default::default()
        }))
        .await
        .context("failed to list images")?;

    Ok(!images.is_empty())
}

async fn build_image(tag: &str) -> Result<()> {
    let manifest_dir = env!("CARGO_MANIFEST_DIR");
    let dockerfile_path = Path::new(manifest_dir).join("Dockerfile.full");

    let status = Command::new("docker")
        .args([
            "build",
            "-f",
            dockerfile_path.to_str().unwrap(),
            "-t",
            tag,
            manifest_dir,
        ])
        .stdout(Stdio::inherit())
        .stderr(Stdio::inherit())
        .status()
        .await
        .context("failed to run docker build")?;

    if !status.success() {
        bail!("docker build failed with status: {}", status);
    }

    Ok(())
}

pub async fn ensure_image() -> Result<String> {
    let tag = image_tag();

    if !image_exists(&tag).await? {
        eprintln!("Building docker image: {}", tag);
        build_image(&tag).await?;
    }

    Ok(tag)
}

pub struct LibreOfficeRecalc {
    workspace_path: String,
    image_tag: String,
}

impl LibreOfficeRecalc {
    pub async fn new(workspace_path: &Path) -> Result<Self> {
        let tag = ensure_image().await?;
        Ok(Self {
            workspace_path: workspace_path.to_str().unwrap().to_string(),
            image_tag: tag,
        })
    }

    pub async fn recalc(&self, file_path: &str) -> Result<()> {
        let container_path = format!("/data/{}", file_path);
        let workspace_mount = format!("{}:/data", self.workspace_path);
        let file_url = format!("file://{}", container_path);

        let mut cmd = Command::new("docker");
        cmd.args([
            "run",
            "--rm",
            "-v",
            &workspace_mount,
            "--entrypoint",
            "soffice",
            &self.image_tag,
            "--headless",
            "--norestore",
            "--nodefault",
            "--nofirststartwizard",
            "--nolockcheck",
            "--calc",
            &format!(
                "macro:///Standard.Module1.RecalculateAndSave(\"{}\")",
                file_url
            ),
        ]);

        // Guard against hangs by enforcing a timeout at the test harness level.
        let output = tokio::time::timeout(std::time::Duration::from_secs(90), cmd.output())
            .await
            .context("docker run timed out")?
            .context("failed to run soffice container")?;

        if !output.status.success() {
            let stderr = String::from_utf8_lossy(&output.stderr);
            let stdout = String::from_utf8_lossy(&output.stdout);
            bail!(
                "soffice failed (exit {}): stderr={}, stdout={}",
                output.status.code().unwrap_or(-1),
                stderr,
                stdout
            );
        }

        Ok(())
    }
}
