use serde_json::Value;
use std::path::Path;
use std::process::Command;
use tempfile::tempdir;

fn write_fixture(path: &Path) {
    let mut workbook = umya_spreadsheet::new_file();
    let sheet = workbook
        .get_sheet_by_name_mut("Sheet1")
        .expect("default sheet exists");
    sheet.get_cell_mut("A1").set_value_number(1.0);
    sheet.get_cell_mut("B1").set_formula("SUM(A1:A1)");
    umya_spreadsheet::writer::xlsx::write(&workbook, path).expect("write workbook");
}

fn run_cli(args: &[&str]) -> std::process::Output {
    Command::new(assert_cmd::cargo::cargo_bin!("spreadsheet-cli"))
        .args(args)
        .output()
        .expect("run spreadsheet-cli")
}

fn parse_stdout_json(output: &std::process::Output) -> Value {
    let stdout = String::from_utf8(output.stdout.clone()).expect("stdout utf8");
    serde_json::from_str(&stdout).expect("valid json")
}

#[test]
fn cli_list_sheets_returns_default_sheet() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("read.xlsx");
    write_fixture(&workbook_path);

    let output = run_cli(&["list-sheets", workbook_path.to_str().expect("path utf8")]);
    assert!(output.status.success(), "stderr: {:?}", output.stderr);

    let payload = parse_stdout_json(&output);
    let sheets = payload["sheets"].as_array().expect("sheets array");
    assert!(
        sheets
            .iter()
            .any(|entry| entry["name"].as_str() == Some("Sheet1"))
    );
}

#[test]
fn cli_copy_edit_diff_reports_changes() {
    let tmp = tempdir().expect("tempdir");
    let original = tmp.path().join("original.xlsx");
    let modified = tmp.path().join("modified.xlsx");
    write_fixture(&original);

    let copy = run_cli(&[
        "copy",
        original.to_str().expect("path utf8"),
        modified.to_str().expect("path utf8"),
    ]);
    assert!(copy.status.success(), "stderr: {:?}", copy.stderr);

    let edit = run_cli(&[
        "edit",
        modified.to_str().expect("path utf8"),
        "Sheet1",
        "A1=42",
        "B1==SUM(A1:A1)",
    ]);
    assert!(edit.status.success(), "stderr: {:?}", edit.stderr);

    let diff = run_cli(&[
        "diff",
        original.to_str().expect("path utf8"),
        modified.to_str().expect("path utf8"),
    ]);
    assert!(diff.status.success(), "stderr: {:?}", diff.stderr);

    let payload = parse_stdout_json(&diff);
    assert!(
        payload["change_count"].as_u64().unwrap_or(0) >= 1,
        "expected at least one diff change"
    );
}
