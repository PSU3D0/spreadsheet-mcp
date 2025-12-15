use anyhow::Result;
use rmcp::{
    ServiceExt,
    model::CallToolRequestParam,
    transport::{ConfigureCommandExt, TokioChildProcess},
};
use serde_json::Value;
use std::fs;
use std::process::Stdio;
use tokio::process::Command;

use super::{TestWorkspace, docker::ensure_image};

pub fn call_tool(name: &'static str, args: Value) -> CallToolRequestParam {
    CallToolRequestParam {
        name: name.into(),
        arguments: args.as_object().cloned(),
    }
}

pub struct McpTestClient {
    workspace: TestWorkspace,
    workspace_path: String,
    allow_overwrite: bool,
    vba_enabled: bool,
    env_overrides: Vec<(String, String)>,
}

impl McpTestClient {
    pub fn new() -> Self {
        let workspace = TestWorkspace::new();
        let workspace_path = workspace.root().to_str().unwrap().to_string();
        Self {
            workspace,
            workspace_path,
            allow_overwrite: false,
            vba_enabled: false,
            env_overrides: Vec::new(),
        }
    }

    pub fn with_allow_overwrite(mut self) -> Self {
        self.allow_overwrite = true;
        self
    }

    pub fn with_vba_enabled(mut self) -> Self {
        self.vba_enabled = true;
        self
    }

    pub fn with_env_override(mut self, key: impl Into<String>, value: impl Into<String>) -> Self {
        self.env_overrides.push((key.into(), value.into()));
        self
    }

    pub fn workspace(&self) -> &TestWorkspace {
        &self.workspace
    }

    pub async fn connect(&self) -> Result<rmcp::service::RunningService<rmcp::RoleClient, ()>> {
        let image_tag = ensure_image().await?;
        let workspace_path = self.workspace_path.clone();
        let allow_overwrite = self.allow_overwrite;
        let vba_enabled = self.vba_enabled;
        let mut env_overrides = self.env_overrides.clone();
        if let Ok(v) = std::env::var("SPREADSHEET_MCP_MAX_PNG_DIM_PX") {
            env_overrides.push(("SPREADSHEET_MCP_MAX_PNG_DIM_PX".to_string(), v));
        }
        if let Ok(v) = std::env::var("SPREADSHEET_MCP_MAX_PNG_AREA_PX") {
            env_overrides.push(("SPREADSHEET_MCP_MAX_PNG_AREA_PX".to_string(), v));
        }
        let screenshot_mount = std::env::var("SPREADSHEET_MCP_TEST_SCREENSHOT_DIR")
            .ok()
            .filter(|s| !s.trim().is_empty())
            .map(|s| {
                if s == "1" || s.eq_ignore_ascii_case("true") {
                    "./test-screenshots".to_string()
                } else {
                    s
                }
            });
        if let Some(dir) = &screenshot_mount {
            fs::create_dir_all(dir)?;
        }

        let (transport, stderr) =
            TokioChildProcess::builder(Command::new("docker").configure(move |cmd| {
                let volume_mount = format!("{}:/data", workspace_path);
                let screenshot_volume_mount = screenshot_mount
                    .as_ref()
                    .map(|dir| format!("{dir}:/data/screenshots"));
                let mut args: Vec<String> = vec![
                    "run".into(),
                    "--rm".into(),
                    "-i".into(),
                    "-v".into(),
                    volume_mount,
                ];
                for (k, v) in &env_overrides {
                    args.push("-e".into());
                    args.push(format!("{k}={v}"));
                }
                if let Some(mount) = &screenshot_volume_mount {
                    args.push("-v".into());
                    args.push(mount.clone());
                }
                args.extend([
                    image_tag.clone(),
                    "--transport".into(),
                    "stdio".into(),
                    "--recalc-enabled".into(),
                    "--workspace-root".into(),
                    "/data".into(),
                ]);
                if vba_enabled {
                    args.push("--vba-enabled".into());
                }
                if allow_overwrite {
                    args.push("--allow-overwrite".into());
                }
                cmd.args(args);
            }))
            .stderr(Stdio::piped())
            .spawn()?;

        if let Some(stderr) = stderr {
            tokio::spawn(async move {
                use tokio::io::{AsyncBufReadExt, BufReader};
                let mut lines = BufReader::new(stderr).lines();
                while let Ok(Some(line)) = lines.next_line().await {
                    eprintln!("[container] {}", line);
                }
            });
        }

        let client = ().serve(transport).await?;
        Ok(client)
    }
}

pub fn extract_json(result: &rmcp::model::CallToolResult) -> Result<Value> {
    result
        .structured_content
        .clone()
        .ok_or_else(|| anyhow::anyhow!("no structured content in response"))
}

pub fn cell_value(page: &Value, row: usize, col: usize) -> Option<String> {
    let cell = &page["rows"][row]["cells"][col];
    let value = &cell["value"];

    match value["kind"].as_str()? {
        "Number" => value["value"].as_f64().map(|n| {
            if n.fract() == 0.0 {
                format!("{}", n as i64)
            } else {
                format!("{}", n)
            }
        }),
        "String" | "Text" => value["value"].as_str().map(|s| s.to_string()),
        "Bool" => value["value"].as_bool().map(|b| b.to_string()),
        _ => None,
    }
}

pub fn cell_value_f64(page: &Value, row: usize, col: usize) -> Option<f64> {
    let cell = &page["rows"][row]["cells"][col];
    cell["value"]["value"].as_f64()
}

pub fn cell_is_error(page: &Value, row: usize, col: usize) -> bool {
    let cell = &page["rows"][row]["cells"][col];
    let kind = cell["value"]["kind"].as_str();
    if kind == Some("Error") {
        return true;
    }
    if let Some(val) = cell["value"]["value"].as_str() {
        return val.starts_with("#");
    }
    false
}

pub fn cell_error_type(page: &Value, row: usize, col: usize) -> Option<String> {
    let cell = &page["rows"][row]["cells"][col];
    if cell["value"]["kind"].as_str() == Some("Error") {
        return cell["value"]["value"].as_str().map(|s| s.to_string());
    }
    if let Some(val) = cell["value"]["value"].as_str()
        && val.starts_with("#")
    {
        return Some(val.to_string());
    }
    None
}
