use serde_json::Value;
use std::path::Path;
use std::process::Command;
use tempfile::tempdir;

fn write_fixture(path: &Path) {
    let mut workbook = umya_spreadsheet::new_file();
    {
        let sheet = workbook
            .get_sheet_by_name_mut("Sheet1")
            .expect("default sheet exists");
        sheet.get_cell_mut("A1").set_value("Name");
        sheet.get_cell_mut("B1").set_value("Amount");
        sheet.get_cell_mut("C1").set_value("Total");
        sheet.get_cell_mut("A2").set_value("Alice");
        sheet.get_cell_mut("B2").set_value_number(10.0);
        sheet.get_cell_mut("C2").set_formula("B2*2");
        sheet.get_cell_mut("A3").set_value("Bob");
        sheet.get_cell_mut("B3").set_value_number(20.0);
        sheet.get_cell_mut("C3").set_formula("B3*2");
        sheet.get_cell_mut("A4").set_value("Carol");
        sheet.get_cell_mut("B4").set_value_number(30.0);
        sheet.get_cell_mut("C4").set_formula("B4*2");
    }

    workbook.new_sheet("Summary").expect("add summary sheet");
    {
        let summary = workbook
            .get_sheet_by_name_mut("Summary")
            .expect("summary sheet exists");
        summary.get_cell_mut("A1").set_value("Flag");
        summary.get_cell_mut("B1").set_value("Ready");
    }

    umya_spreadsheet::writer::xlsx::write(&workbook, path).expect("write workbook");
}

fn run_cli(args: &[&str]) -> std::process::Output {
    Command::new(assert_cmd::cargo::cargo_bin!("agent-spreadsheet"))
        .args(args)
        .output()
        .expect("run agent-spreadsheet")
}

fn parse_stdout_json(output: &std::process::Output) -> Value {
    let stdout = String::from_utf8(output.stdout.clone()).expect("stdout utf8");
    serde_json::from_str(&stdout).expect("valid json")
}

fn parse_stderr_json(output: &std::process::Output) -> Value {
    let stderr = String::from_utf8(output.stderr.clone()).expect("stderr utf8");
    serde_json::from_str(&stderr).expect("valid json error")
}

fn parse_stdout_text(output: &std::process::Output) -> String {
    String::from_utf8(output.stdout.clone()).expect("stdout utf8")
}

#[test]
fn cli_help_surfaces_include_descriptions_and_examples() {
    let root_help = run_cli(&["--help"]);
    assert!(root_help.status.success(), "stderr: {:?}", root_help.stderr);
    let root = parse_stdout_text(&root_help);
    assert!(root.contains("Stateless spreadsheet CLI for AI and automation workflows"));
    assert!(root.contains("Common workflows:"));
    assert!(root.contains("global --format csv is currently unsupported"));
    assert!(root.contains("find-value"));
    assert!(root.contains("Find cells matching a text query by value or label"));

    let find_help = run_cli(&["find-value", "--help"]);
    assert!(find_help.status.success(), "stderr: {:?}", find_help.stderr);
    let find = parse_stdout_text(&find_help);
    assert!(find.contains("Find cells matching a text query by value or label"));
    assert!(find.contains("Examples:"));
    assert!(
        find.contains("find-value data.xlsx \"Net Income\" --sheet \"Q1 Actuals\" --mode label")
    );

    let formula_help = run_cli(&["formula-map", "--help"]);
    assert!(
        formula_help.status.success(),
        "stderr: {:?}",
        formula_help.stderr
    );
    let formula = parse_stdout_text(&formula_help);
    assert!(formula.contains("Summarize formulas on a sheet by complexity or frequency"));
    assert!(formula.contains("Examples:"));
    assert!(formula.contains("formula-map data.xlsx \"Q1 Actuals\" --sort-by count --limit 25"));

    let table_help = run_cli(&["table-profile", "--help"]);
    assert!(
        table_help.status.success(),
        "stderr: {:?}",
        table_help.stderr
    );
    let table = parse_stdout_text(&table_help);
    assert!(table.contains("Profile table headers, types, and column distributions"));
    assert!(table.contains("Examples:"));
    assert!(table.contains("table-profile data.xlsx --sheet \"Q1 Actuals\""));

    let diff_help = run_cli(&["diff", "--help"]);
    assert!(diff_help.status.success(), "stderr: {:?}", diff_help.stderr);
    let diff = parse_stdout_text(&diff_help);
    assert!(diff.contains("Diff two workbook versions and report changed cells"));
    assert!(diff.contains("Examples:"));
    assert!(diff.contains("diff baseline.xlsx candidate.xlsx"));

    let range_help = run_cli(&["range-values", "--help"]);
    assert!(
        range_help.status.success(),
        "stderr: {:?}",
        range_help.stderr
    );
    let range = parse_stdout_text(&range_help);
    assert!(range.contains("Read raw values for one or more A1 ranges"));
    assert!(range.contains("Examples:"));
    assert!(range.contains("range-values data.xlsx \"Q1 Actuals\" A1:B5 D10:E20"));
}

#[test]
fn cli_read_commands_cover_ticket_surface() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("read.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let list = run_cli(&["list-sheets", file]);
    assert!(list.status.success(), "stderr: {:?}", list.stderr);
    let list_payload = parse_stdout_json(&list);
    assert_eq!(list_payload["sheets"].as_array().map(Vec::len), Some(2));

    let overview = run_cli(&["sheet-overview", file, "Sheet1"]);
    assert!(overview.status.success(), "stderr: {:?}", overview.stderr);
    let overview_payload = parse_stdout_json(&overview);
    assert_eq!(overview_payload["sheet_name"], "Sheet1");
    assert!(
        overview_payload["detected_region_count"]
            .as_u64()
            .unwrap_or(0)
            >= 1
    );

    let read_table = run_cli(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--table-format",
        "values",
    ]);
    assert!(
        read_table.status.success(),
        "stderr: {:?}",
        read_table.stderr
    );
    let read_table_payload = parse_stdout_json(&read_table);
    assert_eq!(read_table_payload["sheet_name"], "Sheet1");
    assert!(read_table_payload["values"].is_array());

    let range_values = run_cli(&["range-values", file, "Sheet1", "A1:C4"]);
    assert!(
        range_values.status.success(),
        "stderr: {:?}",
        range_values.stderr
    );
    let range_values_payload = parse_stdout_json(&range_values);
    assert!(range_values_payload.get("workbook_id").is_some());
    assert!(range_values_payload.get("workbook_short_id").is_none());
    let entries = range_values_payload["values"]
        .as_array()
        .expect("range values entries");
    assert_eq!(entries.len(), 1);

    let find_value = run_cli(&["find-value", file, "Bob", "--sheet", "Sheet1"]);
    assert!(
        find_value.status.success(),
        "stderr: {:?}",
        find_value.stderr
    );
    let find_payload = parse_stdout_json(&find_value);
    assert_eq!(find_payload["matches"][0]["address"], "A3");

    let formula_map = run_cli(&[
        "formula-map",
        file,
        "Sheet1",
        "--limit",
        "10",
        "--sort-by",
        "count",
    ]);
    assert!(
        formula_map.status.success(),
        "stderr: {:?}",
        formula_map.stderr
    );
    let formula_map_payload = parse_stdout_json(&formula_map);
    assert!(formula_map_payload["groups"].as_array().is_some());

    let formula_trace = run_cli(&["formula-trace", file, "Sheet1", "C2", "precedents"]);
    assert!(
        formula_trace.status.success(),
        "stderr: {:?}",
        formula_trace.stderr
    );
    let trace_payload = parse_stdout_json(&formula_trace);
    assert_eq!(trace_payload["origin"], "C2");
    assert!(trace_payload["layers"].as_array().is_some());

    let describe = run_cli(&["describe", file]);
    assert!(describe.status.success(), "stderr: {:?}", describe.stderr);
    let describe_payload = parse_stdout_json(&describe);
    assert_eq!(describe_payload["sheet_count"], 2);

    let table_profile = run_cli(&["table-profile", file, "--sheet", "Sheet1"]);
    assert!(
        table_profile.status.success(),
        "stderr: {:?}",
        table_profile.stderr
    );
    let profile_payload = parse_stdout_json(&table_profile);
    assert_eq!(profile_payload["sheet_name"], "Sheet1");
    assert!(
        profile_payload["headers"]
            .as_array()
            .map(Vec::len)
            .unwrap_or(0)
            >= 3
    );
}

#[test]
fn cli_range_values_shape_single_range_canonical_vs_compact() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("shape-single.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let canonical = run_cli(&["range-values", file, "Sheet1", "A1:C4"]);
    assert!(canonical.status.success(), "stderr: {:?}", canonical.stderr);
    let canonical_payload = parse_stdout_json(&canonical);
    assert!(canonical_payload.get("workbook_id").is_some());
    assert!(canonical_payload.get("workbook_short_id").is_none());
    assert_eq!(
        canonical_payload
            .get("values")
            .and_then(Value::as_array)
            .map(Vec::len),
        Some(1)
    );

    let compact = run_cli(&[
        "--shape",
        "compact",
        "range-values",
        file,
        "Sheet1",
        "A1:C4",
    ]);
    assert!(compact.status.success(), "stderr: {:?}", compact.stderr);
    let compact_payload = parse_stdout_json(&compact);
    assert!(compact_payload.get("workbook_id").is_some());
    assert!(compact_payload.get("workbook_short_id").is_none());
    assert!(compact_payload.get("values").is_none());
    assert_eq!(compact_payload["range"], "A1:C4");
    assert!(compact_payload.get("rows").is_some());
}

#[test]
fn cli_range_values_shape_multi_range_canonical_vs_compact() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("shape-multi.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let canonical = run_cli(&["range-values", file, "Sheet1", "A1:A2", "B1:B2"]);
    assert!(canonical.status.success(), "stderr: {:?}", canonical.stderr);
    let canonical_payload = parse_stdout_json(&canonical);
    assert!(canonical_payload.get("workbook_id").is_some());
    assert!(canonical_payload.get("workbook_short_id").is_none());
    let canonical_values = canonical_payload["values"]
        .as_array()
        .expect("canonical multi-range values");
    assert_eq!(canonical_values.len(), 2);

    let compact = run_cli(&[
        "--shape",
        "compact",
        "range-values",
        file,
        "Sheet1",
        "A1:A2",
        "B1:B2",
    ]);
    assert!(compact.status.success(), "stderr: {:?}", compact.stderr);
    let compact_payload = parse_stdout_json(&compact);
    assert!(compact_payload.get("workbook_id").is_some());
    assert!(compact_payload.get("workbook_short_id").is_none());
    let compact_values = compact_payload["values"]
        .as_array()
        .expect("compact multi-range values");
    assert_eq!(compact_values.len(), 2);
    assert!(
        compact_values
            .iter()
            .all(|entry| entry.get("range").and_then(Value::as_str).is_some())
    );
}

#[test]
fn cli_copy_edit_diff_are_stateless_and_persisted() {
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
    let copy_payload = parse_stdout_json(&copy);
    assert!(copy_payload["bytes_copied"].as_u64().unwrap_or(0) > 0);

    let edit = run_cli(&[
        "edit",
        modified.to_str().expect("path utf8"),
        "Sheet1",
        "B2=11",
        "C2==B2*3",
    ]);
    assert!(edit.status.success(), "stderr: {:?}", edit.stderr);
    let edit_payload = parse_stdout_json(&edit);
    assert_eq!(edit_payload["edits_applied"], 2);
    assert_eq!(edit_payload["recalc_needed"], true);

    let book = umya_spreadsheet::reader::xlsx::read(&modified).expect("read modified");
    let sheet = book
        .get_sheet_by_name("Sheet1")
        .expect("modified sheet exists");
    assert_eq!(sheet.get_cell("B2").expect("B2 exists").get_value(), "11");
    assert_eq!(
        sheet.get_cell("C2").expect("C2 exists").get_formula(),
        "B2*3"
    );

    let diff = run_cli(&[
        "diff",
        original.to_str().expect("path utf8"),
        modified.to_str().expect("path utf8"),
    ]);
    assert!(diff.status.success(), "stderr: {:?}", diff.stderr);
    let diff_payload = parse_stdout_json(&diff);
    assert!(diff_payload["change_count"].as_u64().unwrap_or(0) >= 2);
}

#[test]
fn cli_errors_use_machine_envelope() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("read.xlsx");
    write_fixture(&workbook_path);

    let output = run_cli(&[
        "formula-map",
        workbook_path.to_str().expect("path utf8"),
        "Shet1",
    ]);
    assert!(!output.status.success(), "command unexpectedly succeeded");

    let err = parse_stderr_json(&output);
    assert_eq!(err["code"], "SHEET_NOT_FOUND");
    assert_eq!(err["did_you_mean"], "Sheet1");
    assert!(
        err["message"]
            .as_str()
            .unwrap_or_default()
            .contains("was not found")
    );
    assert!(
        err["try_this"]
            .as_str()
            .unwrap_or_default()
            .contains("list-sheets")
    );
}

#[cfg(feature = "recalc-formualizer")]
#[test]
fn cli_recalculate_flow_runs_after_copy_and_edit() {
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
        "B2=25",
    ]);
    assert!(edit.status.success(), "stderr: {:?}", edit.stderr);

    let recalc = run_cli(&["recalculate", modified.to_str().expect("path utf8")]);
    assert!(recalc.status.success(), "stderr: {:?}", recalc.stderr);
    let recalc_payload = parse_stdout_json(&recalc);
    assert_eq!(recalc_payload["backend"], "formualizer");
    assert!(recalc_payload["duration_ms"].as_u64().is_some());

    let diff = run_cli(&[
        "diff",
        original.to_str().expect("path utf8"),
        modified.to_str().expect("path utf8"),
    ]);
    assert!(diff.status.success(), "stderr: {:?}", diff.stderr);
    let diff_payload = parse_stdout_json(&diff);
    assert!(diff_payload["change_count"].as_u64().unwrap_or(0) >= 1);
}
