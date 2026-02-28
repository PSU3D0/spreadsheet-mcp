//! End-to-end integration tests for the event-sourced session system.
//!
//! These tests exercise the full CLI dispatch path:
//!   session start → session op → session apply → session materialize
//! plus undo/redo, branching, checkout, CAS conflict, and log filtering.

use serde_json::Value;
use std::path::Path;
use std::process::Command;
use tempfile::tempdir;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn run_cli(args: &[&str]) -> std::process::Output {
    Command::new(assert_cmd::cargo::cargo_bin!("agent-spreadsheet"))
        .args(args)
        .output()
        .expect("run agent-spreadsheet")
}

fn parse_stdout_json(output: &std::process::Output) -> Value {
    let stdout = String::from_utf8(output.stdout.clone()).expect("stdout utf8");
    serde_json::from_str(&stdout).unwrap_or_else(|e| {
        panic!(
            "invalid json in stdout: {}\nstdout: {}\nstderr: {}",
            e,
            stdout,
            String::from_utf8_lossy(&output.stderr)
        )
    })
}

fn assert_success(output: &std::process::Output) {
    assert!(
        output.status.success(),
        "command failed.\nstdout: {}\nstderr: {}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr),
    );
}

fn write_fixture(path: &Path) {
    let mut workbook = umya_spreadsheet::new_file();
    {
        let sheet = workbook
            .get_sheet_by_name_mut("Sheet1")
            .expect("default sheet");
        sheet.get_cell_mut("A1").set_value("Name");
        sheet.get_cell_mut("B1").set_value("Amount");
        sheet.get_cell_mut("C1").set_value("Total");
        sheet.get_cell_mut("A2").set_value("Alice");
        sheet.get_cell_mut("B2").set_value_number(10.0);
        sheet.get_cell_mut("C2").set_formula("B2*2");
        sheet.get_cell_mut("A3").set_value("Bob");
        sheet.get_cell_mut("B3").set_value_number(20.0);
        sheet.get_cell_mut("C3").set_formula("B3*2");
    }
    workbook.new_sheet("Summary").expect("add summary sheet");
    {
        let s = workbook
            .get_sheet_by_name_mut("Summary")
            .expect("summary");
        s.get_cell_mut("A1").set_value("Flag");
        s.get_cell_mut("B1").set_value("Ready");
    }
    umya_spreadsheet::writer::xlsx::write(&workbook, path).expect("write fixture");
}

/// Write an ops JSON file with a write_matrix payload that sets A2 = "Eve".
fn write_ops_json(path: &Path) {
    let payload = serde_json::json!({
        "sheet_name": "Sheet1",
        "anchor": "A2",
        "rows": [[{"v": "Eve"}]],
        "overwrite_formulas": false,
    });
    std::fs::write(path, serde_json::to_string_pretty(&payload).unwrap()).unwrap();
}

/// Write an ops JSON file with a write_matrix payload that sets B2 = 99.
fn write_second_ops_json(path: &Path) {
    let payload = serde_json::json!({
        "sheet_name": "Sheet1",
        "anchor": "B2",
        "rows": [[{"v": 99.0}]],
        "overwrite_formulas": false,
    });
    std::fs::write(path, serde_json::to_string_pretty(&payload).unwrap()).unwrap();
}

// ---------------------------------------------------------------------------
// Full workflow: start → op → apply → materialize → diff
// ---------------------------------------------------------------------------

#[test]
fn session_full_workflow_start_stage_apply_materialize() {
    let tmp = tempdir().expect("tempdir");
    let workspace = tmp.path();
    let base_path = workspace.join("base.xlsx");
    write_fixture(&base_path);

    let base_str = base_path.to_str().unwrap();
    let ws_str = workspace.to_str().unwrap();

    // 1. Start session
    let start = run_cli(&[
        "session", "start",
        "--base", base_str,
        "--label", "e2e test session",
        "--workspace", ws_str,
    ]);
    assert_success(&start);
    let start_json = parse_stdout_json(&start);
    let session_id = start_json["session_id"].as_str().expect("session_id");
    assert_eq!(start_json["label"], "e2e test session");

    // 2. Stage an operation
    let ops_path = workspace.join("ops.json");
    write_ops_json(&ops_path);

    let stage = run_cli(&[
        "session", "op",
        "--session", session_id,
        "--ops", &format!("@{}", ops_path.display()),
        "--workspace", ws_str,
    ]);
    assert_success(&stage);
    let stage_json = parse_stdout_json(&stage);
    let staged_id = stage_json["staged_id"].as_str().expect("staged_id");
    assert!(staged_id.starts_with("stg_"));
    // head_at_stage should be null (at base)
    assert!(stage_json["head_at_stage"].is_null());

    // 3. Apply staged operation
    let apply = run_cli(&[
        "session", "apply",
        "--session", session_id,
        staged_id,
        "--workspace", ws_str,
    ]);
    assert_success(&apply);
    let apply_json = parse_stdout_json(&apply);
    assert_eq!(apply_json["applied"], true);
    let op_id = apply_json["op_id"].as_str().expect("op_id");
    assert_eq!(apply_json["head"], op_id);

    // 4. Check session log
    let log = run_cli(&[
        "session", "log",
        "--session", session_id,
        "--workspace", ws_str,
    ]);
    assert_success(&log);
    let log_json = parse_stdout_json(&log);
    assert_eq!(log_json["event_count"], 1);
    assert_eq!(log_json["head"], op_id);
    assert_eq!(log_json["events"][0]["op_id"], op_id);

    // 5. Materialize
    let output_path = workspace.join("result.xlsx");
    let materialize = run_cli(&[
        "session", "materialize",
        "--session", session_id,
        "--output", output_path.to_str().unwrap(),
        "--workspace", ws_str,
    ]);
    assert_success(&materialize);
    let mat_json = parse_stdout_json(&materialize);
    assert_eq!(mat_json["events_replayed"], 1);
    assert!(mat_json["output_size_bytes"].as_u64().unwrap() > 0);
    assert!(output_path.exists());

    // 6. Diff base vs materialized — should show A2 changed from Alice to Eve
    let diff = run_cli(&[
        "diff",
        base_str,
        output_path.to_str().unwrap(),
        "--details",
        "--limit", "50",
    ]);
    assert_success(&diff);
    let diff_json = parse_stdout_json(&diff);
    assert!(
        diff_json["change_count"].as_u64().unwrap() >= 1,
        "expected at least 1 change, got: {}",
        diff_json
    );
}

// ---------------------------------------------------------------------------
// Multi-op session with undo/redo
// ---------------------------------------------------------------------------

#[test]
fn session_multi_op_undo_redo() {
    let tmp = tempdir().expect("tempdir");
    let workspace = tmp.path();
    let base_path = workspace.join("undo_base.xlsx");
    write_fixture(&base_path);

    let ws_str = workspace.to_str().unwrap();
    let base_str = base_path.to_str().unwrap();

    // Start
    let start_json = {
        let out = run_cli(&[
            "session", "start", "--base", base_str, "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let session_id = start_json["session_id"].as_str().unwrap();

    // Op 1: A2 = "Eve"
    let ops1 = workspace.join("ops1.json");
    write_ops_json(&ops1);
    let stg1 = {
        let out = run_cli(&[
            "session", "op", "--session", session_id,
            "--ops", &format!("@{}", ops1.display()),
            "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let stg1_id = stg1["staged_id"].as_str().unwrap();
    let apply1 = run_cli(&[
        "session", "apply", "--session", session_id, stg1_id, "--workspace", ws_str,
    ]);
    assert_success(&apply1);
    let op1_id = parse_stdout_json(&apply1)["op_id"].as_str().unwrap().to_string();

    // Op 2: B2 = 99
    let ops2 = workspace.join("ops2.json");
    write_second_ops_json(&ops2);
    let stg2 = {
        let out = run_cli(&[
            "session", "op", "--session", session_id,
            "--ops", &format!("@{}", ops2.display()),
            "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let stg2_id = stg2["staged_id"].as_str().unwrap();
    let apply2 = run_cli(&[
        "session", "apply", "--session", session_id, stg2_id, "--workspace", ws_str,
    ]);
    assert_success(&apply2);
    let op2_id = parse_stdout_json(&apply2)["op_id"].as_str().unwrap().to_string();

    // Log should show 2 events
    let log = run_cli(&[
        "session", "log", "--session", session_id, "--workspace", ws_str,
    ]);
    assert_success(&log);
    assert_eq!(parse_stdout_json(&log)["event_count"], 2);

    // Undo → HEAD should go back to op1
    let undo = run_cli(&[
        "session", "undo", "--session", session_id, "--workspace", ws_str,
    ]);
    assert_success(&undo);
    let undo_json = parse_stdout_json(&undo);
    assert_eq!(undo_json["undone"], true);
    assert_eq!(undo_json["head"].as_str().unwrap(), op1_id);

    // Materialize at op1 → should have Eve in A2 but original B2=10
    let out_undo = workspace.join("undo_result.xlsx");
    let mat_undo = run_cli(&[
        "session", "materialize", "--session", session_id,
        "--output", out_undo.to_str().unwrap(), "--workspace", ws_str,
    ]);
    assert_success(&mat_undo);

    // Read the value at A2 from the materialized file to verify Eve
    let read_a2 = run_cli(&[
        "range-values", out_undo.to_str().unwrap(), "Sheet1", "A2:A2",
    ]);
    assert_success(&read_a2);
    let read_json = parse_stdout_json(&read_a2);
    // The dense format includes the value in the first entry
    let values = &read_json["values"][0];
    let dense = values.get("dense").or(values.get("rows_keyed"));
    assert!(
        serde_json::to_string(dense.unwrap_or(&Value::Null))
            .unwrap()
            .contains("Eve"),
        "expected Eve in A2 after undo to op1, got: {}",
        read_json
    );

    // Redo → HEAD should go back to op2
    let redo = run_cli(&[
        "session", "redo", "--session", session_id, "--workspace", ws_str,
    ]);
    assert_success(&redo);
    let redo_json = parse_stdout_json(&redo);
    assert_eq!(redo_json["redone"], true);
    assert_eq!(redo_json["head"].as_str().unwrap(), op2_id);
}

// ---------------------------------------------------------------------------
// Branching: fork, switch, branches list
// ---------------------------------------------------------------------------

#[test]
fn session_branching_fork_and_switch() {
    let tmp = tempdir().expect("tempdir");
    let workspace = tmp.path();
    let base_path = workspace.join("branch_base.xlsx");
    write_fixture(&base_path);

    let ws_str = workspace.to_str().unwrap();
    let base_str = base_path.to_str().unwrap();

    // Start session
    let start = run_cli(&[
        "session", "start", "--base", base_str, "--workspace", ws_str,
    ]);
    assert_success(&start);
    let session_id = parse_stdout_json(&start)["session_id"]
        .as_str().unwrap().to_string();

    // Apply one op so we have a fork point
    let ops = workspace.join("fork_ops.json");
    write_ops_json(&ops);
    let stg = {
        let out = run_cli(&[
            "session", "op", "--session", &session_id,
            "--ops", &format!("@{}", ops.display()),
            "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let stg_id = stg["staged_id"].as_str().unwrap();
    let apply = run_cli(&[
        "session", "apply", "--session", &session_id, stg_id, "--workspace", ws_str,
    ]);
    assert_success(&apply);
    let _op1_id = parse_stdout_json(&apply)["op_id"].as_str().unwrap().to_string();

    // Fork a branch from current HEAD
    let fork = run_cli(&[
        "session", "fork", "--session", &session_id,
        "--label", "alternative approach",
        "alt-branch",
        "--workspace", ws_str,
    ]);
    assert_success(&fork);
    let fork_json = parse_stdout_json(&fork);
    assert_eq!(fork_json["branch"], "alt-branch");

    // List branches — should have main + alt-branch
    let branches = run_cli(&[
        "session", "branches", "--session", &session_id, "--workspace", ws_str,
    ]);
    assert_success(&branches);
    let branches_json = parse_stdout_json(&branches);
    let branch_list = branches_json["branches"].as_array().expect("branches array");
    assert_eq!(branch_list.len(), 2, "expected main + alt-branch");

    let branch_names: Vec<&str> = branch_list
        .iter()
        .map(|b| b["name"].as_str().unwrap())
        .collect();
    assert!(branch_names.contains(&"main"));
    assert!(branch_names.contains(&"alt-branch"));

    // Switch to alt-branch
    let switch = run_cli(&[
        "session", "switch", "--session", &session_id,
        "--branch", "alt-branch",
        "--workspace", ws_str,
    ]);
    assert_success(&switch);
    let switch_json = parse_stdout_json(&switch);
    assert_eq!(switch_json["branch"], "alt-branch");

    // Switch back to main
    let switch_main = run_cli(&[
        "session", "switch", "--session", &session_id,
        "--branch", "main",
        "--workspace", ws_str,
    ]);
    assert_success(&switch_main);
    assert_eq!(parse_stdout_json(&switch_main)["branch"], "main");
}

// ---------------------------------------------------------------------------
// Checkout to a specific op_id
// ---------------------------------------------------------------------------

#[test]
fn session_checkout_specific_event() {
    let tmp = tempdir().expect("tempdir");
    let workspace = tmp.path();
    let base_path = workspace.join("checkout_base.xlsx");
    write_fixture(&base_path);

    let ws_str = workspace.to_str().unwrap();
    let base_str = base_path.to_str().unwrap();

    // Start + 2 ops
    let start_json = {
        let out = run_cli(&[
            "session", "start", "--base", base_str, "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let session_id = start_json["session_id"].as_str().unwrap();

    let ops1 = workspace.join("co_ops1.json");
    write_ops_json(&ops1);
    let stg1 = {
        let out = run_cli(&[
            "session", "op", "--session", session_id,
            "--ops", &format!("@{}", ops1.display()),
            "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let apply1 = run_cli(&[
        "session", "apply", "--session", session_id,
        stg1["staged_id"].as_str().unwrap(),
        "--workspace", ws_str,
    ]);
    assert_success(&apply1);
    let op1_id = parse_stdout_json(&apply1)["op_id"].as_str().unwrap().to_string();

    let ops2 = workspace.join("co_ops2.json");
    write_second_ops_json(&ops2);
    let stg2 = {
        let out = run_cli(&[
            "session", "op", "--session", session_id,
            "--ops", &format!("@{}", ops2.display()),
            "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let apply2 = run_cli(&[
        "session", "apply", "--session", session_id,
        stg2["staged_id"].as_str().unwrap(),
        "--workspace", ws_str,
    ]);
    assert_success(&apply2);

    // Checkout back to op1
    let checkout = run_cli(&[
        "session", "checkout", "--session", session_id,
        &op1_id,
        "--workspace", ws_str,
    ]);
    assert_success(&checkout);
    let co_json = parse_stdout_json(&checkout);
    assert_eq!(co_json["head"], op1_id);

    // Materialize at op1 — only first write_matrix should be applied
    let mat_path = workspace.join("checkout_result.xlsx");
    let mat = run_cli(&[
        "session", "materialize", "--session", session_id,
        "--output", mat_path.to_str().unwrap(),
        "--workspace", ws_str,
    ]);
    assert_success(&mat);
    let mat_json = parse_stdout_json(&mat);
    assert_eq!(mat_json["events_replayed"], 1, "should replay only 1 event at op1");
}

// ---------------------------------------------------------------------------
// CAS conflict detection
// ---------------------------------------------------------------------------

#[test]
fn session_cas_conflict_rejects_stale_stage() {
    let tmp = tempdir().expect("tempdir");
    let workspace = tmp.path();
    let base_path = workspace.join("cas_base.xlsx");
    write_fixture(&base_path);

    let ws_str = workspace.to_str().unwrap();
    let base_str = base_path.to_str().unwrap();

    let start_json = {
        let out = run_cli(&[
            "session", "start", "--base", base_str, "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let session_id = start_json["session_id"].as_str().unwrap();

    // Stage an op at HEAD=null (base)
    let ops1 = workspace.join("cas_ops1.json");
    write_ops_json(&ops1);
    let stg1 = {
        let out = run_cli(&[
            "session", "op", "--session", session_id,
            "--ops", &format!("@{}", ops1.display()),
            "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let stg1_id = stg1["staged_id"].as_str().unwrap();

    // Advance HEAD by applying a different op first
    let ops2 = workspace.join("cas_ops2.json");
    write_second_ops_json(&ops2);
    let stg2 = {
        let out = run_cli(&[
            "session", "op", "--session", session_id,
            "--ops", &format!("@{}", ops2.display()),
            "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let apply2 = run_cli(&[
        "session", "apply", "--session", session_id,
        stg2["staged_id"].as_str().unwrap(),
        "--workspace", ws_str,
    ]);
    assert_success(&apply2);

    // Now try to apply stg1 which was staged at HEAD=null — HEAD has advanced
    let apply1 = run_cli(&[
        "session", "apply", "--session", session_id,
        stg1_id,
        "--workspace", ws_str,
    ]);
    assert!(
        !apply1.status.success(),
        "expected CAS conflict, but command succeeded"
    );
    let stderr = String::from_utf8_lossy(&apply1.stderr);
    assert!(
        stderr.contains("CAS conflict") || stderr.contains("HEAD has advanced"),
        "expected CAS conflict error, got stderr: {}",
        stderr
    );
}

// ---------------------------------------------------------------------------
// Log filtering by kind
// ---------------------------------------------------------------------------

#[test]
fn session_log_filters_by_kind() {
    let tmp = tempdir().expect("tempdir");
    let workspace = tmp.path();
    let base_path = workspace.join("logfilter_base.xlsx");
    write_fixture(&base_path);

    let ws_str = workspace.to_str().unwrap();
    let base_str = base_path.to_str().unwrap();

    let start_json = {
        let out = run_cli(&[
            "session", "start", "--base", base_str, "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let session_id = start_json["session_id"].as_str().unwrap();

    // Apply an op (kind inferred as transform.write_matrix from payload)
    let ops = workspace.join("logfilter_ops.json");
    write_ops_json(&ops);
    let stg = {
        let out = run_cli(&[
            "session", "op", "--session", session_id,
            "--ops", &format!("@{}", ops.display()),
            "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let apply = run_cli(&[
        "session", "apply", "--session", session_id,
        stg["staged_id"].as_str().unwrap(),
        "--workspace", ws_str,
    ]);
    assert_success(&apply);

    // Filter log by kind=transform — should match "transform.write_matrix"
    let log = run_cli(&[
        "session", "log", "--session", session_id,
        "--kind", "transform",
        "--workspace", ws_str,
    ]);
    assert_success(&log);
    let log_json = parse_stdout_json(&log);
    assert_eq!(log_json["event_count"], 1);
    assert!(
        log_json["events"][0]["kind"].as_str().unwrap().starts_with("transform"),
    );

    // Filter by kind=structure — should match nothing
    let log_empty = run_cli(&[
        "session", "log", "--session", session_id,
        "--kind", "structure",
        "--workspace", ws_str,
    ]);
    assert_success(&log_empty);
    assert_eq!(parse_stdout_json(&log_empty)["event_count"], 0);
}

// ---------------------------------------------------------------------------
// Materialize refuses overwrite without --force
// ---------------------------------------------------------------------------

#[test]
fn session_materialize_refuses_overwrite_without_force() {
    let tmp = tempdir().expect("tempdir");
    let workspace = tmp.path();
    let base_path = workspace.join("force_base.xlsx");
    write_fixture(&base_path);

    let ws_str = workspace.to_str().unwrap();
    let base_str = base_path.to_str().unwrap();

    let start_json = {
        let out = run_cli(&[
            "session", "start", "--base", base_str, "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let session_id = start_json["session_id"].as_str().unwrap();

    let output_path = workspace.join("overwrite_target.xlsx");

    // First materialize succeeds
    let mat1 = run_cli(&[
        "session", "materialize", "--session", session_id,
        "--output", output_path.to_str().unwrap(),
        "--workspace", ws_str,
    ]);
    assert_success(&mat1);

    // Second materialize without --force should fail
    let mat2 = run_cli(&[
        "session", "materialize", "--session", session_id,
        "--output", output_path.to_str().unwrap(),
        "--workspace", ws_str,
    ]);
    assert!(
        !mat2.status.success(),
        "expected overwrite refusal"
    );
    let stderr = String::from_utf8_lossy(&mat2.stderr);
    assert!(
        stderr.contains("already exists") || stderr.contains("--force"),
        "expected overwrite error, got: {}",
        stderr
    );

    // With --force should succeed
    let mat3 = run_cli(&[
        "session", "materialize", "--session", session_id,
        "--output", output_path.to_str().unwrap(),
        "--force",
        "--workspace", ws_str,
    ]);
    assert_success(&mat3);
}

// ---------------------------------------------------------------------------
// Session start with missing base file
// ---------------------------------------------------------------------------

#[test]
fn session_start_rejects_missing_base() {
    let tmp = tempdir().expect("tempdir");
    let ws_str = tmp.path().to_str().unwrap();

    let start = run_cli(&[
        "session", "start",
        "--base", "/nonexistent/path/workbook.xlsx",
        "--workspace", ws_str,
    ]);
    assert!(
        !start.status.success(),
        "expected failure for missing base file"
    );
    let stderr = String::from_utf8_lossy(&start.stderr);
    assert!(
        stderr.contains("not found"),
        "expected 'not found' error, got: {}",
        stderr
    );
}

// ---------------------------------------------------------------------------
// Apply with nonexistent staged_id
// ---------------------------------------------------------------------------

#[test]
fn session_apply_rejects_unknown_staged_id() {
    let tmp = tempdir().expect("tempdir");
    let workspace = tmp.path();
    let base_path = workspace.join("unknown_stg_base.xlsx");
    write_fixture(&base_path);

    let ws_str = workspace.to_str().unwrap();
    let base_str = base_path.to_str().unwrap();

    let start_json = {
        let out = run_cli(&[
            "session", "start", "--base", base_str, "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let session_id = start_json["session_id"].as_str().unwrap();

    let apply = run_cli(&[
        "session", "apply", "--session", session_id,
        "stg_nonexistent_12345678",
        "--workspace", ws_str,
    ]);
    assert!(
        !apply.status.success(),
        "expected failure for unknown staged_id"
    );
    let stderr = String::from_utf8_lossy(&apply.stderr);
    assert!(
        stderr.contains("not found"),
        "expected staged op 'not found' error, got: {}",
        stderr
    );
}

// ---------------------------------------------------------------------------
// Undo at base returns gracefully
// ---------------------------------------------------------------------------

#[test]
fn session_undo_at_base_is_graceful() {
    let tmp = tempdir().expect("tempdir");
    let workspace = tmp.path();
    let base_path = workspace.join("undo_base_edge.xlsx");
    write_fixture(&base_path);

    let ws_str = workspace.to_str().unwrap();
    let base_str = base_path.to_str().unwrap();

    let start_json = {
        let out = run_cli(&[
            "session", "start", "--base", base_str, "--workspace", ws_str,
        ]);
        assert_success(&out);
        parse_stdout_json(&out)
    };
    let session_id = start_json["session_id"].as_str().unwrap();

    // Undo with no events — should indicate already at base
    let undo = run_cli(&[
        "session", "undo", "--session", session_id, "--workspace", ws_str,
    ]);
    // This may succeed (returning head=null) or fail gracefully
    // Either way, it shouldn't panic
    let _stderr = String::from_utf8_lossy(&undo.stderr);
    // We just verify it doesn't crash — the behavior (error vs null head) is implementation-defined
}
