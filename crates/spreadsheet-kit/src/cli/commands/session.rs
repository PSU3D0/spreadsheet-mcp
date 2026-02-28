//! CLI commands for `asp session` subcommand tree.
//!
//! These commands expose the event-sourced session mechanics to the user/agent
//! via stateless, path-driven CLI invocations.

use crate::core::events::{Actor, OpEvent, OpKind};
use crate::core::session_store::{SessionHandle, SessionStore};
use anyhow::{Result, bail};
use serde_json::{Value, json};
use std::path::{Path, PathBuf};

// ---------------------------------------------------------------------------
// Session start
// ---------------------------------------------------------------------------

pub async fn session_start(
    base: PathBuf,
    label: Option<String>,
    workspace: Option<PathBuf>,
) -> Result<Value> {
    let workspace_root = workspace.unwrap_or_else(|| {
        std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."))
    });

    if !base.exists() {
        bail!("base file not found: {}", base.display());
    }

    let store = SessionStore::open(&workspace_root)?;
    let handle = store.create_session(&base, label.as_deref())?;

    Ok(json!({
        "session_id": handle.session_id,
        "base_path": base.display().to_string(),
        "label": label,
        "workspace_root": workspace_root.display().to_string(),
    }))
}

// ---------------------------------------------------------------------------
// Session log
// ---------------------------------------------------------------------------

pub async fn session_log(
    session_id: String,
    workspace: Option<PathBuf>,
    since: Option<String>,
    kind_filter: Option<String>,
) -> Result<Value> {
    let handle = open_session(&session_id, workspace.as_deref())?;
    let events = handle.read_events()?;

    let filtered: Vec<&OpEvent> = events
        .iter()
        .filter(|e| {
            if let Some(ref since_id) = since {
                // Skip events before the since marker
                let pos = events.iter().position(|ev| ev.op_id == *since_id);
                let cur = events.iter().position(|ev| ev.op_id == e.op_id);
                match (pos, cur) {
                    (Some(p), Some(c)) => c >= p,
                    _ => true,
                }
            } else {
                true
            }
        })
        .filter(|e| {
            if let Some(ref kind) = kind_filter {
                e.kind.0.starts_with(kind.as_str())
            } else {
                true
            }
        })
        .collect();

    let entries: Vec<Value> = filtered
        .into_iter()
        .map(|e| {
            json!({
                "op_id": e.op_id,
                "parent_id": e.parent_id,
                "kind": e.kind.0,
                "timestamp": e.timestamp.to_rfc3339(),
                "actor": e.actor.id,
                "has_impact": e.dry_run_impact.is_some(),
                "status": e.apply_result.as_ref().map(|r| format!("{:?}", r.status)),
            })
        })
        .collect();

    let head = handle.read_head()?;
    let branch = handle.current_branch()?;

    Ok(json!({
        "session_id": session_id,
        "branch": branch,
        "head": head,
        "event_count": entries.len(),
        "events": entries,
    }))
}

// ---------------------------------------------------------------------------
// Session branches
// ---------------------------------------------------------------------------

pub async fn session_branches(
    session_id: String,
    workspace: Option<PathBuf>,
) -> Result<Value> {
    let handle = open_session(&session_id, workspace.as_deref())?;
    let branches = handle.list_branches()?;
    let current = handle.current_branch()?;

    let branch_list: Vec<Value> = branches
        .into_iter()
        .map(|b| {
            json!({
                "name": b.name,
                "tip_op_id": b.tip_op_id,
                "fork_point": b.fork_point,
                "label": b.label,
                "current": b.name == current,
            })
        })
        .collect();

    Ok(json!({
        "session_id": session_id,
        "current_branch": current,
        "branches": branch_list,
    }))
}

// ---------------------------------------------------------------------------
// Session switch
// ---------------------------------------------------------------------------

pub async fn session_switch(
    session_id: String,
    branch: String,
    workspace: Option<PathBuf>,
) -> Result<Value> {
    let handle = open_session(&session_id, workspace.as_deref())?;
    handle.switch_branch(&branch)?;

    let head = handle.read_head()?;
    Ok(json!({
        "session_id": session_id,
        "branch": branch,
        "head": head,
    }))
}

// ---------------------------------------------------------------------------
// Session checkout
// ---------------------------------------------------------------------------

pub async fn session_checkout(
    session_id: String,
    op_id: String,
    workspace: Option<PathBuf>,
) -> Result<Value> {
    let handle = open_session(&session_id, workspace.as_deref())?;
    handle.checkout(&op_id)?;

    Ok(json!({
        "session_id": session_id,
        "head": op_id,
    }))
}

// ---------------------------------------------------------------------------
// Session undo / redo
// ---------------------------------------------------------------------------

pub async fn session_undo(
    session_id: String,
    workspace: Option<PathBuf>,
) -> Result<Value> {
    let handle = open_session(&session_id, workspace.as_deref())?;
    let new_head = handle.undo()?;

    Ok(json!({
        "session_id": session_id,
        "head": new_head,
        "undone": true,
    }))
}

pub async fn session_redo(
    session_id: String,
    workspace: Option<PathBuf>,
) -> Result<Value> {
    let handle = open_session(&session_id, workspace.as_deref())?;
    let new_head = handle.redo()?;

    Ok(json!({
        "session_id": session_id,
        "head": new_head,
        "redone": true,
    }))
}

// ---------------------------------------------------------------------------
// Session fork (create branch)
// ---------------------------------------------------------------------------

pub async fn session_fork(
    session_id: String,
    from: Option<String>,
    label: Option<String>,
    branch_name: String,
    workspace: Option<PathBuf>,
) -> Result<Value> {
    let handle = open_session(&session_id, workspace.as_deref())?;

    // If no explicit fork point, use current HEAD
    let head = handle.read_head()?;
    let fork_point_str = from.as_deref().or(head.as_deref());

    handle.create_branch(&branch_name, fork_point_str, label.as_deref())?;

    Ok(json!({
        "session_id": session_id,
        "branch": branch_name,
        "fork_point": fork_point_str,
        "label": label,
    }))
}

// ---------------------------------------------------------------------------
// Session op (stage)
// ---------------------------------------------------------------------------

pub async fn session_op_stage(
    session_id: String,
    ops_ref: String,
    workspace: Option<PathBuf>,
) -> Result<Value> {
    let handle = open_session(&session_id, workspace.as_deref())?;

    // Load the ops payload
    let payload_json = load_ops_payload(&ops_ref)?;
    let head = handle.read_head()?;

    // Create a staged artifact
    let staged_id = format!("stg_{:013x}_{:08x}",
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis(),
        rand::random::<u32>()
    );

    let staged_artifact = json!({
        "staged_id": staged_id,
        "session_id": session_id,
        "head_at_stage": head,
        "ops_payload": payload_json,
        "created_at": chrono::Utc::now().to_rfc3339(),
    });

    let staged_path = handle.staged_dir().join(format!("{}.json", staged_id));
    std::fs::write(&staged_path, serde_json::to_string_pretty(&staged_artifact)?)?;

    Ok(json!({
        "staged_id": staged_id,
        "session_id": session_id,
        "head_at_stage": head,
        "staged_path": staged_path.display().to_string(),
    }))
}

// ---------------------------------------------------------------------------
// Session apply (from staged)
// ---------------------------------------------------------------------------

pub async fn session_apply(
    session_id: String,
    staged_id: String,
    workspace: Option<PathBuf>,
) -> Result<Value> {
    let handle = open_session(&session_id, workspace.as_deref())?;

    // Load staged artifact
    let staged_path = handle.staged_dir().join(format!("{}.json", staged_id));
    if !staged_path.exists() {
        bail!("staged operation not found: {}", staged_id);
    }

    let staged_content = std::fs::read_to_string(&staged_path)?;
    let staged: Value = serde_json::from_str(&staged_content)?;

    // Verify CAS: HEAD must match head_at_stage
    let current_head = handle.read_head()?;
    let staged_head = staged
        .get("head_at_stage")
        .and_then(|v| v.as_str())
        .map(|s| s.to_string());

    // Normalize: None and null/"" are equivalent (both mean "at base")
    let current_head_normalized = current_head.as_deref().filter(|s| !s.is_empty());
    let staged_head_normalized = staged_head.as_deref().filter(|s| !s.is_empty());

    if current_head_normalized != staged_head_normalized {
        bail!(
            "CAS conflict: HEAD has advanced since staging. HEAD={:?}, staged expected={:?}. \
             Re-stage the operation against the current HEAD.",
            current_head,
            staged_head,
        );
    }

    let payload = staged
        .get("ops_payload")
        .cloned()
        .unwrap_or(json!({}));

    // Infer the operation kind from the payload so that replay_event_on_session
    // can route to the correct handler. If the payload contains a `kind` field,
    // use it directly; otherwise, infer from structural hints (rows → write_matrix,
    // fallback to edit.batch).
    let kind = if let Some(kind_val) = payload.get("kind").and_then(|v| v.as_str()) {
        let parts: Vec<&str> = kind_val.splitn(2, '.').collect();
        if parts.len() == 2 {
            OpKind::new(parts[0], parts[1])
        } else {
            OpKind::edit_batch()
        }
    } else if payload.get("rows").is_some() {
        OpKind::transform_write_matrix()
    } else {
        OpKind::edit_batch()
    };

    // Create and append the event
    let event = OpEvent::new(
        session_id.clone(),
        current_head.clone(),
        Actor {
            id: "cli:user".to_string(),
            run_id: None,
            source: "cli".to_string(),
        },
        kind,
        payload,
    );

    let op_id = event.op_id.clone();
    handle.append_event(event)?;

    // Clean up staged file
    let _ = std::fs::remove_file(&staged_path);

    Ok(json!({
        "session_id": session_id,
        "op_id": op_id,
        "staged_id": staged_id,
        "applied": true,
        "head": op_id,
    }))
}

// ---------------------------------------------------------------------------
// Session materialize
// ---------------------------------------------------------------------------

pub async fn session_materialize(
    session_id: String,
    output: PathBuf,
    workspace: Option<PathBuf>,
    force: bool,
) -> Result<Value> {
    let handle = open_session(&session_id, workspace.as_deref())?;

    if output.exists() && !force {
        bail!(
            "output file already exists: {}. Use --force to overwrite.",
            output.display()
        );
    }

    let bytes = handle.materialize()?;
    std::fs::write(&output, &bytes)?;

    let head = handle.read_head()?;
    let all_events = handle.read_events()?;
    // Count events actually replayed (up to and including HEAD)
    let events_replayed = if let Some(ref head_id) = head {
        all_events
            .iter()
            .position(|e| e.op_id == *head_id)
            .map(|pos| pos + 1)
            .unwrap_or(all_events.len())
    } else {
        0 // at base, nothing replayed
    };

    Ok(json!({
        "session_id": session_id,
        "output_path": output.display().to_string(),
        "head": head,
        "events_replayed": events_replayed,
        "output_size_bytes": bytes.len(),
    }))
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn open_session(session_id: &str, workspace: Option<&Path>) -> Result<SessionHandle> {
    let workspace_root = workspace
        .map(|p| p.to_path_buf())
        .unwrap_or_else(|| std::env::current_dir().unwrap_or_else(|_| PathBuf::from(".")));
    let store = SessionStore::open(&workspace_root)?;
    store.open_session(session_id)
}

fn load_ops_payload(ops_ref: &str) -> Result<Value> {
    let path = if let Some(stripped) = ops_ref.strip_prefix('@') {
        stripped
    } else {
        ops_ref
    };

    let content = std::fs::read_to_string(path)
        .map_err(|e| anyhow::anyhow!("failed to read ops payload from '{}': {}", path, e))?;

    serde_json::from_str(&content)
        .map_err(|e| anyhow::anyhow!("failed to parse ops payload JSON from '{}': {}", path, e))
}
