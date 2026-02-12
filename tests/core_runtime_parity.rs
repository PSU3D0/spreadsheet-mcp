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

#[test]
fn core_write_and_cli_write_produce_parity_diff_counts() {
    let tmp = tempdir().expect("tempdir");
    let original = tmp.path().join("original.xlsx");
    let cli_modified = tmp.path().join("cli.xlsx");
    let core_modified = tmp.path().join("core.xlsx");
    write_fixture(&original);
    std::fs::copy(&original, &cli_modified).expect("copy cli target");
    std::fs::copy(&original, &core_modified).expect("copy core target");

    let edit = run_cli(&[
        "edit",
        cli_modified.to_str().expect("path utf8"),
        "Sheet1",
        "A1=42",
        "B1==SUM(A1:A1)",
    ]);
    assert!(edit.status.success(), "stderr: {:?}", edit.stderr);

    let shorthand = vec!["A1=42", "B1==SUM(A1:A1)"];
    let mut core_edits = Vec::new();
    for entry in shorthand {
        let (edit, _warnings) = spreadsheet_mcp::core::write::normalize_shorthand_edit(entry)
            .expect("normalize shorthand");
        core_edits.push(edit);
    }
    spreadsheet_mcp::core::write::apply_edits_to_file(&core_modified, "Sheet1", &core_edits)
        .expect("apply core edits");

    let cli_diff = spreadsheet_mcp::core::diff::diff_workbooks_json(&original, &cli_modified)
        .expect("cli diff");
    let core_diff = spreadsheet_mcp::core::diff::diff_workbooks_json(&original, &core_modified)
        .expect("core diff");

    assert_eq!(cli_diff["change_count"], core_diff["change_count"]);
    assert!(cli_diff["change_count"].as_u64().unwrap_or(0) >= 1);
}

#[cfg(feature = "recalc")]
#[test]
fn mcp_normalize_wrapper_matches_core_write_normalization() {
    use spreadsheet_mcp::tools::write_normalize::{EditBatchParamsInput, normalize_edit_batch};

    let input = serde_json::json!({
        "fork_id": "fork-1",
        "sheet_name": "Sheet1",
        "edits": [
            "A1=Hello",
            { "address": "B2", "formula": "=SUM(A1:A1)" }
        ]
    });
    let params: EditBatchParamsInput = serde_json::from_value(input).expect("valid params");
    let (wrapped, wrapped_warnings) = normalize_edit_batch(params).expect("normalize wrapper");

    let (s1, mut expected_warnings) =
        spreadsheet_mcp::core::write::normalize_shorthand_edit("A1=Hello")
            .expect("shorthand normalize");
    let (s2, more_warnings) = spreadsheet_mcp::core::write::normalize_object_edit(
        "B2",
        None,
        Some("=SUM(A1:A1)".to_string()),
        None,
    )
    .expect("object normalize");
    expected_warnings.extend(more_warnings);

    assert_eq!(wrapped.edits.len(), 2);
    assert_eq!(wrapped.edits[0].address, s1.address);
    assert_eq!(wrapped.edits[0].value, s1.value);
    assert_eq!(wrapped.edits[0].is_formula, s1.is_formula);
    assert_eq!(wrapped.edits[1].address, s2.address);
    assert_eq!(wrapped.edits[1].value, s2.value);
    assert_eq!(wrapped.edits[1].is_formula, s2.is_formula);

    let wrapped_codes: Vec<_> = wrapped_warnings.iter().map(|w| w.code.as_str()).collect();
    let expected_codes: Vec<_> = expected_warnings.iter().map(|w| w.code.as_str()).collect();
    assert_eq!(wrapped_codes, expected_codes);
}
