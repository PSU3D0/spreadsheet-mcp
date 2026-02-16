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

fn write_trace_pagination_fixture(path: &Path) {
    let mut workbook = umya_spreadsheet::new_file();
    {
        let sheet = workbook
            .get_sheet_by_name_mut("Sheet1")
            .expect("default sheet exists");
        sheet.get_cell_mut("A1").set_value_number(1.0);
        for row in 1..=18 {
            let address = format!("B{row}");
            let formula = format!("A1+{row}");
            sheet.get_cell_mut(address.as_str()).set_formula(formula);
        }
    }

    umya_spreadsheet::writer::xlsx::write(&workbook, path).expect("write workbook");
}

fn write_phase1_read_surface_fixture(path: &Path) {
    let mut workbook = umya_spreadsheet::new_file();
    {
        let sheet = workbook
            .get_sheet_by_name_mut("Sheet1")
            .expect("default sheet exists");
        sheet.get_cell_mut("A1").set_value("Name");
        sheet.get_cell_mut("B1").set_value("Amount");
        sheet.get_cell_mut("C1").set_value("Calc");
        sheet.get_cell_mut("D1").set_value("Volatile");

        sheet.get_cell_mut("A2").set_value("Alice");
        sheet.get_cell_mut("B2").set_value_number(10.0);
        sheet.get_cell_mut("C2").set_formula("SUM(B2:B2)");
        sheet.get_cell_mut("D2").set_formula("NOW()");

        sheet.get_cell_mut("A3").set_value("Bob");
        sheet.get_cell_mut("B3").set_value_number(20.0);
        sheet.get_cell_mut("C3").set_formula("SUM(B3:B3)");
        sheet.get_cell_mut("D3").set_formula("RAND()");

        sheet.get_cell_mut("A4").set_value("Carol");
        sheet.get_cell_mut("B4").set_value_number(30.0);
        sheet.get_cell_mut("C4").set_formula("SUM(B4:B4)");
        sheet.get_cell_mut("D4").set_formula("TODAY()");

        let mut table = umya_spreadsheet::structs::Table::new("SalesTable", ("A1", "D4"));
        table.set_display_name("SalesTable");
        sheet.add_table(table);
    }

    workbook.new_sheet("Summary").expect("add summary sheet");
    {
        let summary = workbook
            .get_sheet_by_name_mut("Summary")
            .expect("summary sheet exists");
        summary.get_cell_mut("A1").set_value("Flag");
        summary.get_cell_mut("B1").set_value("Ready");
    }

    let sheet1 = workbook
        .get_sheet_by_name_mut("Sheet1")
        .expect("sheet1 exists");
    sheet1
        .add_defined_name("Sales_Amount", "Sheet1!$B$2:$B$4")
        .expect("defined name Sales_Amount");
    sheet1
        .add_defined_name("Sales_First", "Sheet1!$A$2")
        .expect("defined name Sales_First");
    let summary = workbook
        .get_sheet_by_name_mut("Summary")
        .expect("summary exists");
    summary
        .add_defined_name("Meta_Flag", "Summary!$A$1")
        .expect("defined name Meta_Flag");

    umya_spreadsheet::writer::xlsx::write(&workbook, path).expect("write workbook");
}

fn write_workbook_short_id_column_fixture(path: &Path) {
    let mut workbook = umya_spreadsheet::new_file();
    {
        let sheet = workbook
            .get_sheet_by_name_mut("Sheet1")
            .expect("default sheet exists");
        sheet.get_cell_mut("A1").set_value("workbook_short_id");
        sheet.get_cell_mut("B1").set_value("Name");
        sheet.get_cell_mut("A2").set_value("user-data-id");
        sheet.get_cell_mut("B2").set_value("Alice");
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

fn assert_invalid_argument(args: &[&str]) -> Value {
    let output = run_cli(args);
    assert!(
        !output.status.success(),
        "command unexpectedly succeeded for args: {args:?}"
    );
    let err = parse_stderr_json(&output);
    assert_eq!(
        err["code"], "INVALID_ARGUMENT",
        "unexpected error envelope: {err}"
    );
    err
}

#[test]
fn cli_help_surfaces_include_descriptions_and_examples() {
    let root_help = run_cli(&["--help"]);
    assert!(root_help.status.success(), "stderr: {:?}", root_help.stderr);
    let root = parse_stdout_text(&root_help);
    assert!(root.contains("Stateless spreadsheet CLI for AI and automation workflows"));
    assert!(root.contains("Common workflows:"));
    assert!(root.contains("global --output-format csv is currently unsupported"));
    assert!(root.contains("find-value"));
    assert!(root.contains("named-ranges"));
    assert!(root.contains("find-formula"));
    assert!(root.contains("scan-volatiles"));
    assert!(root.contains("sheet-statistics"));
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

    let named_ranges_help = run_cli(&["named-ranges", "--help"]);
    assert!(
        named_ranges_help.status.success(),
        "stderr: {:?}",
        named_ranges_help.stderr
    );
    let named_ranges = parse_stdout_text(&named_ranges_help);
    assert!(named_ranges.contains("List workbook named ranges and table/formula named items"));
    assert!(named_ranges.contains("Examples:"));
    assert!(named_ranges.contains("named-ranges data.xlsx"));
    assert!(
        named_ranges.contains("named-ranges data.xlsx --sheet \"Q1 Actuals\" --name-prefix Sales")
    );

    let find_formula_help = run_cli(&["find-formula", "--help"]);
    assert!(
        find_formula_help.status.success(),
        "stderr: {:?}",
        find_formula_help.stderr
    );
    let find_formula = parse_stdout_text(&find_formula_help);
    assert!(find_formula.contains("Find formulas containing a text query with pagination"));
    assert!(find_formula.contains("Examples:"));
    assert!(find_formula.contains("find-formula data.xlsx SUM("));
    assert!(
        find_formula.contains(
            "find-formula data.xlsx VLOOKUP --sheet \"Q1 Actuals\" --limit 25 --offset 50"
        )
    );

    let scan_volatiles_help = run_cli(&["scan-volatiles", "--help"]);
    assert!(
        scan_volatiles_help.status.success(),
        "stderr: {:?}",
        scan_volatiles_help.stderr
    );
    let scan_volatiles = parse_stdout_text(&scan_volatiles_help);
    assert!(scan_volatiles.contains("Scan workbook formulas for volatile functions"));
    assert!(scan_volatiles.contains("Examples:"));
    assert!(scan_volatiles.contains("scan-volatiles data.xlsx"));
    assert!(
        scan_volatiles
            .contains("scan-volatiles data.xlsx --sheet \"Q1 Actuals\" --limit 10 --offset 10")
    );

    let sheet_statistics_help = run_cli(&["sheet-statistics", "--help"]);
    assert!(
        sheet_statistics_help.status.success(),
        "stderr: {:?}",
        sheet_statistics_help.stderr
    );
    let sheet_statistics = parse_stdout_text(&sheet_statistics_help);
    assert!(sheet_statistics.contains("Compute per-sheet statistics for density and column types"));
    assert!(sheet_statistics.contains("Examples:"));
    assert!(sheet_statistics.contains("sheet-statistics data.xlsx Sheet1"));
    assert!(sheet_statistics.contains("sheet-statistics data.xlsx \"Q1 Actuals\""));

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
fn cli_phase1_named_ranges_filters_are_deterministic() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase1-named-ranges.xlsx");
    write_phase1_read_surface_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let baseline = run_cli(&["named-ranges", file]);
    assert!(baseline.status.success(), "stderr: {:?}", baseline.stderr);
    let baseline_payload = parse_stdout_json(&baseline);
    let baseline_items = baseline_payload["items"].as_array().expect("items array");
    assert!(!baseline_items.is_empty());

    let by_sheet = run_cli(&["named-ranges", file, "--sheet", "Sheet1"]);
    assert!(by_sheet.status.success(), "stderr: {:?}", by_sheet.stderr);
    let by_sheet_payload = parse_stdout_json(&by_sheet);
    let by_sheet_items = by_sheet_payload["items"].as_array().expect("items array");
    assert!(!by_sheet_items.is_empty());
    assert!(
        by_sheet_items
            .iter()
            .all(|item| item["sheet_name"] == "Sheet1")
    );

    let by_prefix_first = run_cli(&["named-ranges", file, "--name-prefix", "Sales"]);
    assert!(
        by_prefix_first.status.success(),
        "stderr: {:?}",
        by_prefix_first.stderr
    );
    let by_prefix_first_payload = parse_stdout_json(&by_prefix_first);
    let by_prefix_first_items = by_prefix_first_payload["items"]
        .as_array()
        .expect("items array");
    assert!(!by_prefix_first_items.is_empty());
    assert!(by_prefix_first_items.iter().all(|item| {
        item["name"]
            .as_str()
            .map(|name| name.starts_with("Sales"))
            .unwrap_or(false)
    }));

    let by_prefix_second = run_cli(&["named-ranges", file, "--name-prefix", "Sales"]);
    assert!(
        by_prefix_second.status.success(),
        "stderr: {:?}",
        by_prefix_second.stderr
    );
    let by_prefix_second_payload = parse_stdout_json(&by_prefix_second);
    assert_eq!(by_prefix_first_payload, by_prefix_second_payload);
}

#[test]
fn cli_phase1_find_formula_supports_limit_offset_continuation() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase1-find-formula.xlsx");
    write_phase1_read_surface_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let first = run_cli(&[
        "find-formula",
        file,
        "SUM(",
        "--sheet",
        "Sheet1",
        "--limit",
        "1",
        "--offset",
        "0",
    ]);
    assert!(first.status.success(), "stderr: {:?}", first.stderr);
    let first_payload = parse_stdout_json(&first);
    let first_matches = first_payload["matches"].as_array().expect("matches array");
    assert_eq!(first_matches.len(), 1);
    let first_next = first_payload["next_offset"]
        .as_u64()
        .expect("next_offset on first page");

    let second_offset = first_next.to_string();
    let second = run_cli(&[
        "find-formula",
        file,
        "SUM(",
        "--sheet",
        "Sheet1",
        "--limit",
        "1",
        "--offset",
        second_offset.as_str(),
    ]);
    assert!(second.status.success(), "stderr: {:?}", second.stderr);
    let second_payload = parse_stdout_json(&second);
    let second_matches = second_payload["matches"].as_array().expect("matches array");
    assert_eq!(second_matches.len(), 1);
    let second_next = second_payload["next_offset"].as_u64().unwrap_or(first_next);
    assert!(second_next >= first_next);

    let terminal = run_cli(&[
        "find-formula",
        file,
        "SUM(",
        "--sheet",
        "Sheet1",
        "--limit",
        "10",
        "--offset",
        "2",
    ]);
    assert!(terminal.status.success(), "stderr: {:?}", terminal.stderr);
    let terminal_payload = parse_stdout_json(&terminal);
    assert!(
        terminal_payload["matches"]
            .as_array()
            .map(Vec::len)
            .unwrap_or(0)
            >= 1
    );
    assert!(terminal_payload.get("next_offset").is_none());
}

#[test]
fn cli_phase1_scan_volatiles_detects_and_paginates_deterministically() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase1-scan-volatiles.xlsx");
    write_phase1_read_surface_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let full = run_cli(&["scan-volatiles", file]);
    assert!(full.status.success(), "stderr: {:?}", full.stderr);
    let full_payload = parse_stdout_json(&full);
    let full_items = full_payload["items"].as_array().expect("items array");
    assert!(!full_items.is_empty());

    let first = run_cli(&["scan-volatiles", file, "--limit", "1", "--offset", "0"]);
    assert!(first.status.success(), "stderr: {:?}", first.stderr);
    let first_payload = parse_stdout_json(&first);
    let first_items = first_payload["items"].as_array().expect("items array");
    assert_eq!(first_items.len(), 1);
    let first_entry = first_items[0].clone();
    let first_next = first_payload["next_offset"]
        .as_u64()
        .expect("next_offset for first volatile page");

    let second_offset = first_next.to_string();
    let second = run_cli(&[
        "scan-volatiles",
        file,
        "--limit",
        "1",
        "--offset",
        second_offset.as_str(),
    ]);
    assert!(second.status.success(), "stderr: {:?}", second.stderr);
    let second_payload = parse_stdout_json(&second);
    let second_items = second_payload["items"].as_array().expect("items array");
    assert_eq!(second_items.len(), 1);
    let second_entry = second_items[0].clone();
    assert_ne!(
        first_entry, second_entry,
        "continuation repeated first entry"
    );

    let second_again = run_cli(&[
        "scan-volatiles",
        file,
        "--limit",
        "1",
        "--offset",
        second_offset.as_str(),
    ]);
    assert!(
        second_again.status.success(),
        "stderr: {:?}",
        second_again.stderr
    );
    let second_again_payload = parse_stdout_json(&second_again);
    assert_eq!(second_payload, second_again_payload);
}

#[test]
fn cli_phase1_sheet_statistics_returns_expected_fields() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase1-sheet-statistics.xlsx");
    write_phase1_read_surface_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let output = run_cli(&["sheet-statistics", file, "Sheet1"]);
    assert!(output.status.success(), "stderr: {:?}", output.stderr);
    let payload = parse_stdout_json(&output);

    assert!(payload["row_count"].as_u64().unwrap_or(0) >= 4);
    assert!(payload["column_count"].as_u64().unwrap_or(0) >= 4);
    assert!(payload["density"].as_f64().unwrap_or(0.0) > 0.0);
    assert!(payload["numeric_columns"].is_array());
    assert!(payload["text_columns"].is_array());
}

#[test]
fn cli_phase1_sheet_scoped_commands_unknown_sheet_return_sheet_not_found() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase1-sheet-not-found.xlsx");
    write_phase1_read_surface_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let cases: Vec<Vec<&str>> = vec![
        vec!["named-ranges", file, "--sheet", "Shet1"],
        vec!["find-formula", file, "SUM(", "--sheet", "Shet1"],
        vec!["scan-volatiles", file, "--sheet", "Shet1"],
        vec!["sheet-statistics", file, "Shet1"],
    ];

    for args in cases {
        let output = run_cli(&args);
        assert!(
            !output.status.success(),
            "command unexpectedly succeeded: {args:?}"
        );
        let err = parse_stderr_json(&output);
        assert_eq!(err["code"], "SHEET_NOT_FOUND", "unexpected envelope: {err}");
    }
}

#[test]
fn cli_phase1_invalid_limit_flags_return_invalid_argument() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase1-invalid-limit.xlsx");
    write_phase1_read_surface_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    assert_invalid_argument(&["find-formula", file, "SUM(", "--limit", "0"]);
    assert_invalid_argument(&["scan-volatiles", file, "--limit", "0"]);
}

#[test]
fn cli_phase1_malformed_usage_prints_help_and_exits_non_zero() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase1-malformed-usage.xlsx");
    write_phase1_read_surface_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let missing_query = run_cli(&["find-formula", file]);
    assert!(
        !missing_query.status.success(),
        "find-formula without query should fail"
    );
    let missing_query_stderr = String::from_utf8(missing_query.stderr).expect("stderr utf8");
    assert!(missing_query_stderr.contains("Usage:"));
    assert!(missing_query_stderr.contains("find-formula <FILE> <QUERY>"));
}

#[test]
fn cli_read_table_pagination_round_trips_next_offset_with_sample_mode_first() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("read-table-pagination.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let mut offset = 0u32;
    let mut saw_continuation = false;
    let mut saw_terminal = false;

    for _ in 0..10 {
        let offset_arg = offset.to_string();
        let page = run_cli(&[
            "read-table",
            file,
            "--sheet",
            "Sheet1",
            "--range",
            "A1:C4",
            "--table-format",
            "json",
            "--sample-mode",
            "first",
            "--limit",
            "1",
            "--offset",
            offset_arg.as_str(),
        ]);
        assert!(page.status.success(), "stderr: {:?}", page.stderr);

        let payload = parse_stdout_json(&page);
        assert!(payload["rows"].is_array());

        if let Some(next_offset) = payload["next_offset"].as_u64() {
            saw_continuation = true;
            assert!(
                next_offset > offset as u64,
                "next_offset must strictly increase for sample-mode=first"
            );
            offset = next_offset as u32;
        } else {
            saw_terminal = true;
            break;
        }
    }

    assert!(saw_continuation, "expected at least one continuation page");
    assert!(saw_terminal, "pagination did not reach a terminal page");
}

#[test]
fn cli_formula_trace_pagination_round_trips_next_cursor_until_terminal() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("formula-trace-pagination.xlsx");
    write_trace_pagination_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let first_page = run_cli(&[
        "formula-trace",
        file,
        "Sheet1",
        "A1",
        "dependents",
        "--depth",
        "1",
        "--page-size",
        "5",
    ]);
    assert!(
        first_page.status.success(),
        "stderr: {:?}",
        first_page.stderr
    );
    let first_payload = parse_stdout_json(&first_page);
    let first_cursor = first_payload["next_cursor"]
        .as_object()
        .expect("expected next_cursor on first trace page");
    let mut cursor_depth = first_cursor["depth"].as_u64().expect("cursor depth") as u32;
    let mut cursor_offset = first_cursor["offset"].as_u64().expect("cursor offset") as usize;

    let mut saw_terminal = false;
    for _ in 0..10 {
        let depth_arg = cursor_depth.to_string();
        let offset_arg = cursor_offset.to_string();
        let page = run_cli(&[
            "formula-trace",
            file,
            "Sheet1",
            "A1",
            "dependents",
            "--depth",
            "1",
            "--page-size",
            "5",
            "--cursor-depth",
            depth_arg.as_str(),
            "--cursor-offset",
            offset_arg.as_str(),
        ]);
        assert!(page.status.success(), "stderr: {:?}", page.stderr);

        let payload = parse_stdout_json(&page);
        if let Some(next_cursor) = payload["next_cursor"].as_object() {
            let next_depth = next_cursor["depth"].as_u64().expect("next depth");
            let next_offset = next_cursor["offset"].as_u64().expect("next offset");
            assert_eq!(
                next_depth, cursor_depth as u64,
                "cursor depth should round-trip unchanged"
            );
            assert!(
                next_offset > cursor_offset as u64,
                "cursor offset should strictly increase while paginating"
            );
            cursor_depth = next_depth as u32;
            cursor_offset = next_offset as usize;
        } else {
            saw_terminal = true;
            break;
        }
    }

    assert!(
        saw_terminal,
        "formula-trace pagination did not reach a terminal page"
    );
}

#[test]
fn cli_sheet_page_first_page_emits_next_start_row() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("sheet-page-first.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let page = run_cli(&[
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        "2",
        "--page-size",
        "1",
        "--format",
        "full",
    ]);
    assert!(page.status.success(), "stderr: {:?}", page.stderr);

    let payload = parse_stdout_json(&page);
    assert_eq!(payload["format"], "full");
    assert_eq!(payload["rows"].as_array().map(Vec::len), Some(1));
    assert_eq!(payload["rows"][0]["row_index"].as_u64(), Some(2));
    assert_eq!(payload["next_start_row"].as_u64(), Some(3));
}

#[test]
fn cli_sheet_page_continuation_round_trips_deterministically() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("sheet-page-continuation.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let first = run_cli(&[
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        "2",
        "--page-size",
        "1",
        "--format",
        "full",
    ]);
    assert!(first.status.success(), "stderr: {:?}", first.stderr);
    let first_payload = parse_stdout_json(&first);
    let next_start_row = first_payload["next_start_row"]
        .as_u64()
        .expect("next_start_row present")
        .to_string();

    let continuation = run_cli(&[
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        next_start_row.as_str(),
        "--page-size",
        "1",
        "--format",
        "full",
    ]);
    assert!(
        continuation.status.success(),
        "stderr: {:?}",
        continuation.stderr
    );
    let continuation_payload = parse_stdout_json(&continuation);

    let direct = run_cli(&[
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        "3",
        "--page-size",
        "1",
        "--format",
        "full",
    ]);
    assert!(direct.status.success(), "stderr: {:?}", direct.stderr);
    let direct_payload = parse_stdout_json(&direct);

    assert_eq!(continuation_payload, direct_payload);
}

#[test]
fn cli_sheet_page_terminal_page_omits_next_start_row() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("sheet-page-terminal.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let terminal = run_cli(&[
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        "4",
        "--page-size",
        "2",
        "--format",
        "full",
    ]);
    assert!(terminal.status.success(), "stderr: {:?}", terminal.stderr);

    let payload = parse_stdout_json(&terminal);
    assert_eq!(payload["rows"][0]["row_index"].as_u64(), Some(4));
    assert!(payload.get("next_start_row").is_none());
}

#[test]
fn cli_sheet_page_column_filters_support_union_and_sheet_order() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("sheet-page-columns.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let columns_only = run_cli(&[
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        "2",
        "--page-size",
        "1",
        "--columns",
        "C:A",
        "--format",
        "compact",
    ]);
    assert!(
        columns_only.status.success(),
        "stderr: {:?}",
        columns_only.stderr
    );
    let columns_only_payload = parse_stdout_json(&columns_only);
    let columns_only_headers = columns_only_payload["compact"]["headers"]
        .as_array()
        .expect("compact headers")
        .iter()
        .map(|v| v.as_str().expect("header string"))
        .collect::<Vec<_>>();
    assert_eq!(columns_only_headers, vec!["Row", "Name", "Amount", "Total"]);

    let header_only = run_cli(&[
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        "2",
        "--page-size",
        "1",
        "--columns-by-header",
        "Total,Name",
        "--format",
        "compact",
    ]);
    assert!(
        header_only.status.success(),
        "stderr: {:?}",
        header_only.stderr
    );
    let header_only_payload = parse_stdout_json(&header_only);
    let header_only_headers = header_only_payload["compact"]["headers"]
        .as_array()
        .expect("compact headers")
        .iter()
        .map(|v| v.as_str().expect("header string"))
        .collect::<Vec<_>>();
    assert_eq!(header_only_headers, vec!["Row", "Name", "Total"]);

    let combined = run_cli(&[
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        "2",
        "--page-size",
        "1",
        "--columns",
        "B",
        "--columns-by-header",
        "Amount,Name,Total",
        "--format",
        "compact",
    ]);
    assert!(combined.status.success(), "stderr: {:?}", combined.stderr);
    let combined_payload = parse_stdout_json(&combined);
    let combined_headers = combined_payload["compact"]["headers"]
        .as_array()
        .expect("compact headers")
        .iter()
        .map(|v| v.as_str().expect("header string"))
        .collect::<Vec<_>>();
    assert_eq!(combined_headers, vec!["Row", "Name", "Amount", "Total"]);
}

#[test]
fn cli_sheet_page_accepts_all_formats_and_sets_expected_payload_branch() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("sheet-page-formats.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    for format in ["full", "compact", "values_only"] {
        let page = run_cli(&[
            "sheet-page",
            file,
            "Sheet1",
            "--start-row",
            "2",
            "--page-size",
            "1",
            "--format",
            format,
        ]);
        assert!(page.status.success(), "stderr: {:?}", page.stderr);
        let payload = parse_stdout_json(&page);

        assert_eq!(payload["format"], format);
        match format {
            "full" => {
                assert!(payload["rows"].is_array());
                assert!(payload.get("compact").is_none());
                assert!(payload.get("values_only").is_none());
            }
            "compact" => {
                assert!(payload["compact"].is_object());
                assert!(payload.get("rows").is_none());
                assert!(payload.get("values_only").is_none());
            }
            "values_only" => {
                assert!(payload["values_only"].is_object());
                assert!(payload.get("rows").is_none());
                assert!(payload.get("compact").is_none());
            }
            _ => unreachable!(),
        }
    }
}

#[test]
fn cli_sheet_page_preserves_next_start_row_in_canonical_and_compact_shapes() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("sheet-page-shape-next-start-row.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let canonical = run_cli(&[
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        "2",
        "--page-size",
        "1",
        "--format",
        "compact",
    ]);
    assert!(canonical.status.success(), "stderr: {:?}", canonical.stderr);
    let canonical_payload = parse_stdout_json(&canonical);

    let compact_shape = run_cli(&[
        "--shape",
        "compact",
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        "2",
        "--page-size",
        "1",
        "--format",
        "compact",
    ]);
    assert!(
        compact_shape.status.success(),
        "stderr: {:?}",
        compact_shape.stderr
    );
    let compact_shape_payload = parse_stdout_json(&compact_shape);

    assert_eq!(canonical_payload["next_start_row"].as_u64(), Some(3));
    assert_eq!(compact_shape_payload["next_start_row"].as_u64(), Some(3));
    assert_eq!(
        canonical_payload["next_start_row"],
        compact_shape_payload["next_start_row"]
    );
}

#[test]
fn cli_shape_3109_read_table_compact_preserves_contract_branches() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("shape-3109-read-table-branches.xlsx");
    write_phase1_read_surface_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    for (table_format, branch) in [("json", "rows"), ("values", "values"), ("csv", "csv")] {
        let canonical = run_cli(&[
            "read-table",
            file,
            "--sheet",
            "Sheet1",
            "--table-name",
            "SalesTable",
            "--table-format",
            table_format,
            "--sample-mode",
            "first",
            "--limit",
            "1",
            "--offset",
            "0",
        ]);
        assert!(canonical.status.success(), "stderr: {:?}", canonical.stderr);
        let canonical_payload = parse_stdout_json(&canonical);

        let compact = run_cli(&[
            "--shape",
            "compact",
            "read-table",
            file,
            "--sheet",
            "Sheet1",
            "--table-name",
            "SalesTable",
            "--table-format",
            table_format,
            "--sample-mode",
            "first",
            "--limit",
            "1",
            "--offset",
            "0",
        ]);
        assert!(compact.status.success(), "stderr: {:?}", compact.stderr);
        let compact_payload = parse_stdout_json(&compact);

        assert_eq!(
            compact_payload["workbook_id"],
            canonical_payload["workbook_id"]
        );
        assert_eq!(compact_payload["sheet_name"], "Sheet1");
        assert_eq!(compact_payload["table_name"], "SalesTable");
        assert_eq!(
            compact_payload["total_rows"],
            canonical_payload["total_rows"]
        );
        assert_eq!(
            compact_payload["next_offset"],
            canonical_payload["next_offset"]
        );

        match branch {
            "rows" => {
                assert!(compact_payload["rows"].is_array());
                assert!(compact_payload.get("values").is_none());
                assert!(compact_payload.get("csv").is_none());
            }
            "values" => {
                assert!(compact_payload["values"].is_array());
                assert!(compact_payload.get("rows").is_none());
                assert!(compact_payload.get("csv").is_none());
            }
            "csv" => {
                assert!(compact_payload["csv"].is_string());
                assert!(compact_payload.get("rows").is_none());
                assert!(compact_payload.get("values").is_none());
            }
            _ => unreachable!(),
        }

        assert_eq!(compact_payload, canonical_payload);
    }
}

#[test]
fn cli_shape_3109_read_table_compact_round_trips_next_offset_until_terminal() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("shape-3109-read-table-next-offset.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let canonical_first = run_cli(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--table-format",
        "json",
        "--sample-mode",
        "first",
        "--limit",
        "1",
        "--offset",
        "0",
    ]);
    assert!(
        canonical_first.status.success(),
        "stderr: {:?}",
        canonical_first.stderr
    );
    let canonical_first_payload = parse_stdout_json(&canonical_first);

    let compact_first = run_cli(&[
        "--shape",
        "compact",
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--table-format",
        "json",
        "--sample-mode",
        "first",
        "--limit",
        "1",
        "--offset",
        "0",
    ]);
    assert!(
        compact_first.status.success(),
        "stderr: {:?}",
        compact_first.stderr
    );
    let compact_first_payload = parse_stdout_json(&compact_first);

    assert_eq!(
        compact_first_payload["next_offset"],
        canonical_first_payload["next_offset"]
    );

    let mut offset = compact_first_payload["next_offset"]
        .as_u64()
        .expect("next_offset on compact first page") as u32;
    let mut saw_terminal = false;

    for _ in 0..10 {
        let offset_arg = offset.to_string();
        let page = run_cli(&[
            "--shape",
            "compact",
            "read-table",
            file,
            "--sheet",
            "Sheet1",
            "--range",
            "A1:C4",
            "--table-format",
            "json",
            "--sample-mode",
            "first",
            "--limit",
            "1",
            "--offset",
            offset_arg.as_str(),
        ]);
        assert!(page.status.success(), "stderr: {:?}", page.stderr);

        let payload = parse_stdout_json(&page);
        if let Some(next_offset) = payload["next_offset"].as_u64() {
            assert!(
                next_offset > offset as u64,
                "next_offset must strictly increase"
            );
            offset = next_offset as u32;
        } else {
            saw_terminal = true;
            break;
        }
    }

    assert!(
        saw_terminal,
        "compact read-table pagination did not reach a terminal page"
    );
}

#[test]
fn cli_shape_3109_read_table_compact_preserves_user_workbook_short_id_columns() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp
        .path()
        .join("shape-3109-read-table-workbook-short-id-column.xlsx");
    write_workbook_short_id_column_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let canonical = run_cli(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:B2",
        "--table-format",
        "json",
    ]);
    assert!(canonical.status.success(), "stderr: {:?}", canonical.stderr);
    let canonical_payload = parse_stdout_json(&canonical);

    let compact = run_cli(&[
        "--shape",
        "compact",
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:B2",
        "--table-format",
        "json",
    ]);
    assert!(compact.status.success(), "stderr: {:?}", compact.stderr);
    let compact_payload = parse_stdout_json(&compact);

    assert_eq!(compact_payload, canonical_payload);

    let row = compact_payload["rows"]
        .as_array()
        .and_then(|rows| rows.first())
        .and_then(Value::as_object)
        .expect("first compact row object");
    assert!(row.contains_key("workbook_short_id"));
}

#[test]
fn cli_shape_3109_sheet_page_compact_preserves_active_branches_without_collapse() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("shape-3109-sheet-page-branches.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    for format in ["full", "compact", "values_only"] {
        let canonical = run_cli(&[
            "sheet-page",
            file,
            "Sheet1",
            "--start-row",
            "2",
            "--page-size",
            "2",
            "--format",
            format,
        ]);
        assert!(canonical.status.success(), "stderr: {:?}", canonical.stderr);
        let canonical_payload = parse_stdout_json(&canonical);

        let compact_shape = run_cli(&[
            "--shape",
            "compact",
            "sheet-page",
            file,
            "Sheet1",
            "--start-row",
            "2",
            "--page-size",
            "2",
            "--format",
            format,
        ]);
        assert!(
            compact_shape.status.success(),
            "stderr: {:?}",
            compact_shape.stderr
        );
        let compact_payload = parse_stdout_json(&compact_shape);

        assert_eq!(
            compact_payload["workbook_id"],
            canonical_payload["workbook_id"]
        );
        assert_eq!(compact_payload["sheet_name"], "Sheet1");
        assert_eq!(compact_payload["format"], format);
        assert_eq!(
            compact_payload["next_start_row"],
            canonical_payload["next_start_row"]
        );

        match format {
            "full" => {
                let compact_rows = compact_payload["rows"].as_array().expect("full rows");
                let canonical_rows = canonical_payload["rows"]
                    .as_array()
                    .expect("canonical full rows");
                assert_eq!(compact_rows.len(), canonical_rows.len());
                assert!(compact_rows.len() > 1, "expected multi-row full payload");
                assert!(compact_payload.get("compact").is_none());
                assert!(compact_payload.get("values_only").is_none());
            }
            "compact" => {
                let compact_rows = compact_payload["compact"]["rows"]
                    .as_array()
                    .expect("compact branch rows");
                let canonical_rows = canonical_payload["compact"]["rows"]
                    .as_array()
                    .expect("canonical compact branch rows");
                assert_eq!(compact_rows.len(), canonical_rows.len());
                assert!(compact_rows.len() > 1, "expected multi-row compact payload");
                assert!(compact_payload.get("rows").is_none());
                assert!(compact_payload.get("values_only").is_none());
            }
            "values_only" => {
                let compact_rows = compact_payload["values_only"]["rows"]
                    .as_array()
                    .expect("values_only branch rows");
                let canonical_rows = canonical_payload["values_only"]["rows"]
                    .as_array()
                    .expect("canonical values_only branch rows");
                assert_eq!(compact_rows.len(), canonical_rows.len());
                assert!(
                    compact_rows.len() > 1,
                    "expected multi-row values_only payload"
                );
                assert!(compact_payload.get("rows").is_none());
                assert!(compact_payload.get("compact").is_none());
            }
            _ => unreachable!(),
        }

        assert_eq!(compact_payload, canonical_payload);
    }
}

#[test]
fn cli_shape_3109_sheet_page_compact_round_trips_next_start_row() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("shape-3109-sheet-page-next-start-row.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let canonical_first = run_cli(&[
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        "2",
        "--page-size",
        "1",
        "--format",
        "compact",
    ]);
    assert!(
        canonical_first.status.success(),
        "stderr: {:?}",
        canonical_first.stderr
    );
    let canonical_first_payload = parse_stdout_json(&canonical_first);

    let compact_first = run_cli(&[
        "--shape",
        "compact",
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        "2",
        "--page-size",
        "1",
        "--format",
        "compact",
    ]);
    assert!(
        compact_first.status.success(),
        "stderr: {:?}",
        compact_first.stderr
    );
    let compact_first_payload = parse_stdout_json(&compact_first);

    assert_eq!(
        compact_first_payload["next_start_row"],
        canonical_first_payload["next_start_row"]
    );

    let next_start_row = compact_first_payload["next_start_row"]
        .as_u64()
        .expect("next_start_row on compact first page")
        .to_string();

    let continuation = run_cli(&[
        "--shape",
        "compact",
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        next_start_row.as_str(),
        "--page-size",
        "1",
        "--format",
        "compact",
    ]);
    assert!(
        continuation.status.success(),
        "stderr: {:?}",
        continuation.stderr
    );
    let continuation_payload = parse_stdout_json(&continuation);

    let direct = run_cli(&[
        "--shape",
        "compact",
        "sheet-page",
        file,
        "Sheet1",
        "--start-row",
        "3",
        "--page-size",
        "1",
        "--format",
        "compact",
    ]);
    assert!(direct.status.success(), "stderr: {:?}", direct.stderr);
    let direct_payload = parse_stdout_json(&direct);

    assert_eq!(continuation_payload, direct_payload);
}

#[test]
fn cli_shape_3109_formula_trace_compact_omits_layer_highlights_and_preserves_cursor() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp
        .path()
        .join("shape-3109-formula-trace-compact-contract.xlsx");
    write_trace_pagination_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let canonical = run_cli(&[
        "formula-trace",
        file,
        "Sheet1",
        "A1",
        "dependents",
        "--depth",
        "1",
        "--page-size",
        "5",
    ]);
    assert!(canonical.status.success(), "stderr: {:?}", canonical.stderr);
    let canonical_payload = parse_stdout_json(&canonical);

    let compact_shape = run_cli(&[
        "--shape",
        "compact",
        "formula-trace",
        file,
        "Sheet1",
        "A1",
        "dependents",
        "--depth",
        "1",
        "--page-size",
        "5",
    ]);
    assert!(
        compact_shape.status.success(),
        "stderr: {:?}",
        compact_shape.stderr
    );
    let compact_payload = parse_stdout_json(&compact_shape);

    assert_eq!(
        compact_payload["workbook_id"],
        canonical_payload["workbook_id"]
    );
    assert_eq!(
        compact_payload["sheet_name"],
        canonical_payload["sheet_name"]
    );
    assert_eq!(compact_payload["origin"], canonical_payload["origin"]);
    assert_eq!(compact_payload["direction"], canonical_payload["direction"]);
    assert_eq!(
        compact_payload["next_cursor"],
        canonical_payload["next_cursor"]
    );
    assert_eq!(compact_payload["notes"], canonical_payload["notes"]);

    let canonical_layers = canonical_payload["layers"]
        .as_array()
        .expect("canonical layers")
        .clone();
    assert!(!canonical_layers.is_empty(), "expected canonical layers");
    assert!(
        canonical_layers
            .iter()
            .all(|layer| layer.get("highlights").is_some()),
        "canonical layers should include highlights"
    );

    let compact_layers = compact_payload["layers"]
        .as_array()
        .expect("compact layers");
    assert_eq!(compact_layers.len(), canonical_layers.len());
    assert!(
        compact_layers
            .iter()
            .all(|layer| layer.get("highlights").is_none()),
        "compact layers must omit highlights"
    );
    assert!(compact_layers.iter().all(|layer| {
        layer.get("depth").is_some()
            && layer.get("summary").is_some()
            && layer.get("edges").is_some()
            && layer.get("has_more").is_some()
    }));

    for (canonical_layer, compact_layer) in canonical_layers.iter().zip(compact_layers.iter()) {
        assert_eq!(compact_layer["depth"], canonical_layer["depth"]);
        assert_eq!(compact_layer["summary"], canonical_layer["summary"]);
        assert_eq!(compact_layer["has_more"], canonical_layer["has_more"]);

        let mut canonical_edges = canonical_layer["edges"]
            .as_array()
            .cloned()
            .unwrap_or_default();
        let mut compact_edges = compact_layer["edges"]
            .as_array()
            .cloned()
            .unwrap_or_default();

        canonical_edges.sort_by(|a, b| {
            serde_json::to_string(a)
                .expect("serialize canonical edge")
                .cmp(&serde_json::to_string(b).expect("serialize canonical edge"))
        });
        compact_edges.sort_by(|a, b| {
            serde_json::to_string(a)
                .expect("serialize compact edge")
                .cmp(&serde_json::to_string(b).expect("serialize compact edge"))
        });

        assert_eq!(compact_edges, canonical_edges);
    }
}

#[test]
fn cli_shape_3109_formula_trace_compact_round_trips_next_cursor_until_terminal() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("shape-3109-formula-trace-next-cursor.xlsx");
    write_trace_pagination_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let canonical_first = run_cli(&[
        "formula-trace",
        file,
        "Sheet1",
        "A1",
        "dependents",
        "--depth",
        "1",
        "--page-size",
        "5",
    ]);
    assert!(
        canonical_first.status.success(),
        "stderr: {:?}",
        canonical_first.stderr
    );
    let canonical_first_payload = parse_stdout_json(&canonical_first);

    let compact_first = run_cli(&[
        "--shape",
        "compact",
        "formula-trace",
        file,
        "Sheet1",
        "A1",
        "dependents",
        "--depth",
        "1",
        "--page-size",
        "5",
    ]);
    assert!(
        compact_first.status.success(),
        "stderr: {:?}",
        compact_first.stderr
    );
    let compact_first_payload = parse_stdout_json(&compact_first);

    assert_eq!(
        compact_first_payload["next_cursor"],
        canonical_first_payload["next_cursor"]
    );

    let first_cursor = compact_first_payload["next_cursor"]
        .as_object()
        .expect("next_cursor on first compact trace page");
    let mut cursor_depth = first_cursor["depth"].as_u64().expect("cursor depth") as u32;
    let mut cursor_offset = first_cursor["offset"].as_u64().expect("cursor offset") as usize;

    let mut saw_terminal = false;
    for _ in 0..10 {
        let depth_arg = cursor_depth.to_string();
        let offset_arg = cursor_offset.to_string();
        let page = run_cli(&[
            "--shape",
            "compact",
            "formula-trace",
            file,
            "Sheet1",
            "A1",
            "dependents",
            "--depth",
            "1",
            "--page-size",
            "5",
            "--cursor-depth",
            depth_arg.as_str(),
            "--cursor-offset",
            offset_arg.as_str(),
        ]);
        assert!(page.status.success(), "stderr: {:?}", page.stderr);

        let payload = parse_stdout_json(&page);
        let layers = payload["layers"].as_array().expect("layers array");
        assert!(layers.iter().all(|layer| layer.get("highlights").is_none()));

        if let Some(next_cursor) = payload["next_cursor"].as_object() {
            let next_depth = next_cursor["depth"].as_u64().expect("next depth") as u32;
            let next_offset = next_cursor["offset"].as_u64().expect("next offset") as usize;
            assert_eq!(next_depth, cursor_depth);
            assert!(next_offset > cursor_offset, "cursor offset must increase");
            cursor_depth = next_depth;
            cursor_offset = next_offset;
        } else {
            saw_terminal = true;
            break;
        }
    }

    assert!(
        saw_terminal,
        "compact formula-trace pagination did not reach a terminal page"
    );
}

#[test]
fn cli_shape_3109_compact_does_not_over_apply_to_unrelated_find_value_payloads() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("shape-3109-over-apply-find-value.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let canonical = run_cli(&["find-value", file, "Bob", "--sheet", "Sheet1"]);
    assert!(canonical.status.success(), "stderr: {:?}", canonical.stderr);
    let canonical_payload = parse_stdout_json(&canonical);

    let compact_shape = run_cli(&[
        "--shape",
        "compact",
        "find-value",
        file,
        "Bob",
        "--sheet",
        "Sheet1",
    ]);
    assert!(
        compact_shape.status.success(),
        "stderr: {:?}",
        compact_shape.stderr
    );
    let compact_payload = parse_stdout_json(&compact_shape);

    assert_eq!(compact_payload, canonical_payload);
}

#[test]
fn cli_shape_3109_default_shape_matches_explicit_canonical_for_ticket_commands() {
    let tmp = tempdir().expect("tempdir");

    let read_table_workbook = tmp
        .path()
        .join("shape-3109-default-canonical-read-table.xlsx");
    write_fixture(&read_table_workbook);
    let read_table_file = read_table_workbook.to_str().expect("path utf8");
    let read_table_default = run_cli(&[
        "read-table",
        read_table_file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--table-format",
        "json",
        "--sample-mode",
        "first",
        "--limit",
        "1",
        "--offset",
        "0",
    ]);
    assert!(
        read_table_default.status.success(),
        "stderr: {:?}",
        read_table_default.stderr
    );
    let read_table_canonical = run_cli(&[
        "--shape",
        "canonical",
        "read-table",
        read_table_file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--table-format",
        "json",
        "--sample-mode",
        "first",
        "--limit",
        "1",
        "--offset",
        "0",
    ]);
    assert!(
        read_table_canonical.status.success(),
        "stderr: {:?}",
        read_table_canonical.stderr
    );
    assert_eq!(
        parse_stdout_json(&read_table_default),
        parse_stdout_json(&read_table_canonical)
    );

    let sheet_page_workbook = tmp
        .path()
        .join("shape-3109-default-canonical-sheet-page.xlsx");
    write_fixture(&sheet_page_workbook);
    let sheet_page_file = sheet_page_workbook.to_str().expect("path utf8");
    let sheet_page_default = run_cli(&[
        "sheet-page",
        sheet_page_file,
        "Sheet1",
        "--start-row",
        "2",
        "--page-size",
        "1",
        "--format",
        "full",
    ]);
    assert!(
        sheet_page_default.status.success(),
        "stderr: {:?}",
        sheet_page_default.stderr
    );
    let sheet_page_canonical = run_cli(&[
        "--shape",
        "canonical",
        "sheet-page",
        sheet_page_file,
        "Sheet1",
        "--start-row",
        "2",
        "--page-size",
        "1",
        "--format",
        "full",
    ]);
    assert!(
        sheet_page_canonical.status.success(),
        "stderr: {:?}",
        sheet_page_canonical.stderr
    );
    assert_eq!(
        parse_stdout_json(&sheet_page_default),
        parse_stdout_json(&sheet_page_canonical)
    );

    let trace_workbook = tmp
        .path()
        .join("shape-3109-default-canonical-formula-trace.xlsx");
    write_trace_pagination_fixture(&trace_workbook);
    let trace_file = trace_workbook.to_str().expect("path utf8");
    let trace_default = run_cli(&[
        "formula-trace",
        trace_file,
        "Sheet1",
        "A1",
        "dependents",
        "--depth",
        "1",
        "--page-size",
        "5",
    ]);
    assert!(
        trace_default.status.success(),
        "stderr: {:?}",
        trace_default.stderr
    );
    let trace_canonical = run_cli(&[
        "--shape",
        "canonical",
        "formula-trace",
        trace_file,
        "Sheet1",
        "A1",
        "dependents",
        "--depth",
        "1",
        "--page-size",
        "5",
    ]);
    assert!(
        trace_canonical.status.success(),
        "stderr: {:?}",
        trace_canonical.stderr
    );

    let trace_default_payload = parse_stdout_json(&trace_default);
    let trace_canonical_payload = parse_stdout_json(&trace_canonical);
    assert_eq!(
        trace_default_payload["workbook_id"],
        trace_canonical_payload["workbook_id"]
    );
    assert_eq!(
        trace_default_payload["sheet_name"],
        trace_canonical_payload["sheet_name"]
    );
    assert_eq!(
        trace_default_payload["origin"],
        trace_canonical_payload["origin"]
    );
    assert_eq!(
        trace_default_payload["direction"],
        trace_canonical_payload["direction"]
    );
    assert_eq!(
        trace_default_payload["next_cursor"],
        trace_canonical_payload["next_cursor"]
    );
    assert_eq!(
        trace_default_payload["notes"],
        trace_canonical_payload["notes"]
    );

    let default_layers = trace_default_payload["layers"]
        .as_array()
        .expect("default layers");
    let canonical_layers = trace_canonical_payload["layers"]
        .as_array()
        .expect("canonical layers");
    assert_eq!(default_layers.len(), canonical_layers.len());

    for (default_layer, canonical_layer) in default_layers.iter().zip(canonical_layers.iter()) {
        assert_eq!(default_layer["depth"], canonical_layer["depth"]);
        assert_eq!(default_layer["summary"], canonical_layer["summary"]);
        assert_eq!(default_layer["has_more"], canonical_layer["has_more"]);
        assert_eq!(
            default_layer.get("highlights").is_some(),
            canonical_layer.get("highlights").is_some()
        );

        let mut default_edges = default_layer["edges"]
            .as_array()
            .cloned()
            .unwrap_or_default();
        let mut canonical_edges = canonical_layer["edges"]
            .as_array()
            .cloned()
            .unwrap_or_default();

        default_edges.sort_by(|a, b| {
            serde_json::to_string(a)
                .expect("serialize default edge")
                .cmp(&serde_json::to_string(b).expect("serialize default edge"))
        });
        canonical_edges.sort_by(|a, b| {
            serde_json::to_string(a)
                .expect("serialize canonical edge")
                .cmp(&serde_json::to_string(b).expect("serialize canonical edge"))
        });

        assert_eq!(default_edges, canonical_edges);
    }
}

#[test]
fn cli_sheet_page_page_size_zero_returns_invalid_argument() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("sheet-page-page-size-zero.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    assert_invalid_argument(&[
        "sheet-page",
        file,
        "Sheet1",
        "--page-size",
        "0",
        "--format",
        "full",
    ]);
}

#[test]
fn cli_sheet_page_invalid_column_spec_returns_invalid_argument() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("sheet-page-invalid-column.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    assert_invalid_argument(&[
        "sheet-page",
        file,
        "Sheet1",
        "--columns",
        "A,NOT$",
        "--format",
        "full",
    ]);
}

#[test]
fn cli_sheet_page_unknown_sheet_returns_sheet_not_found() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("sheet-page-sheet-not-found.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let output = run_cli(&["sheet-page", file, "Shet1", "--format", "full"]);
    assert!(!output.status.success(), "command unexpectedly succeeded");

    let err = parse_stderr_json(&output);
    assert_eq!(err["code"], "SHEET_NOT_FOUND");
    assert_eq!(err["did_you_mean"], "Sheet1");
}

#[test]
fn cli_sheet_page_unknown_format_value_fails_clap_parse() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("sheet-page-unknown-format.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let output = run_cli(&["sheet-page", file, "Sheet1", "--format", "bogus"]);
    assert!(!output.status.success(), "command unexpectedly succeeded");

    let stderr = String::from_utf8(output.stderr).expect("stderr utf8");
    assert!(stderr.contains("invalid value 'bogus'"), "stderr: {stderr}");
    assert!(
        stderr.contains("--format <FORMAT>"),
        "expected clap parse error for --format, got: {stderr}"
    );
    assert!(
        stderr.contains("full") && stderr.contains("compact") && stderr.contains("values_only"),
        "expected sheet-page format choices in error, got: {stderr}"
    );
}

#[test]
fn cli_read_table_filters_support_unfiltered_json_and_file_inputs() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("read-table-filters.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let unfiltered = run_cli(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--table-format",
        "json",
    ]);
    assert!(
        unfiltered.status.success(),
        "stderr: {:?}",
        unfiltered.stderr
    );
    let unfiltered_payload = parse_stdout_json(&unfiltered);
    assert_eq!(unfiltered_payload["rows"].as_array().map(Vec::len), Some(3));

    let filters_json = r#"[{"column":"Name","op":"eq","value":"Alice"}]"#;
    let filtered_json = run_cli(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--table-format",
        "json",
        "--filters-json",
        filters_json,
    ]);
    assert!(
        filtered_json.status.success(),
        "stderr: {:?}",
        filtered_json.stderr
    );
    let filtered_json_payload = parse_stdout_json(&filtered_json);
    assert_eq!(
        filtered_json_payload["rows"].as_array().map(Vec::len),
        Some(1)
    );

    let filters_file = tmp.path().join("filters.json");
    std::fs::write(&filters_file, filters_json).expect("write filters file");
    let filters_file_path = filters_file.to_str().expect("filters path utf8");
    let filtered_file = run_cli(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--table-format",
        "json",
        "--filters-file",
        filters_file_path,
    ]);
    assert!(
        filtered_file.status.success(),
        "stderr: {:?}",
        filtered_file.stderr
    );
    let filtered_file_payload = parse_stdout_json(&filtered_file);
    assert_eq!(
        filtered_file_payload["rows"].as_array().map(Vec::len),
        Some(1)
    );
}

#[test]
fn cli_read_table_allows_last_and_distributed_sampling_at_zero_offset() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("read-table-sample-modes.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let last = run_cli(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--table-format",
        "json",
        "--sample-mode",
        "last",
        "--offset",
        "0",
        "--limit",
        "2",
    ]);
    assert!(last.status.success(), "stderr: {:?}", last.stderr);
    let last_payload = parse_stdout_json(&last);
    assert!(last_payload["rows"].is_array());

    let distributed = run_cli(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--table-format",
        "json",
        "--sample-mode",
        "distributed",
        "--offset",
        "0",
        "--limit",
        "2",
    ]);
    assert!(
        distributed.status.success(),
        "stderr: {:?}",
        distributed.stderr
    );
    let distributed_payload = parse_stdout_json(&distributed);
    assert!(distributed_payload["rows"].is_array());
}

#[test]
fn cli_pagination_surface_validation_failures_use_invalid_argument() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("validation.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let filter_file = tmp.path().join("filters.json");
    let filter_json = r#"[{"column":"Name","op":"eq","value":"Alice"}]"#;
    std::fs::write(&filter_file, filter_json).expect("write filters file");
    let filter_file_path = filter_file.to_str().expect("path utf8");

    let malformed_filter_file = tmp.path().join("bad-filters.json");
    std::fs::write(&malformed_filter_file, "{not-json").expect("write malformed filter file");
    let malformed_filter_file_path = malformed_filter_file.to_str().expect("path utf8");

    assert_invalid_argument(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--filters-json",
        filter_json,
        "--filters-file",
        filter_file_path,
    ]);

    assert_invalid_argument(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--filters-json",
        "{",
    ]);

    assert_invalid_argument(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--filters-file",
        malformed_filter_file_path,
    ]);

    assert_invalid_argument(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--limit",
        "0",
    ]);

    assert_invalid_argument(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--sample-mode",
        "last",
        "--offset",
        "1",
    ]);

    assert_invalid_argument(&[
        "read-table",
        file,
        "--sheet",
        "Sheet1",
        "--range",
        "A1:C4",
        "--sample-mode",
        "distributed",
        "--offset",
        "1",
    ]);

    assert_invalid_argument(&[
        "formula-trace",
        file,
        "Sheet1",
        "C2",
        "precedents",
        "--cursor-depth",
        "1",
    ]);

    assert_invalid_argument(&[
        "formula-trace",
        file,
        "Sheet1",
        "C2",
        "precedents",
        "--cursor-offset",
        "1",
    ]);

    assert_invalid_argument(&[
        "formula-trace",
        file,
        "Sheet1",
        "C2",
        "precedents",
        "--depth",
        "0",
    ]);

    assert_invalid_argument(&[
        "formula-trace",
        file,
        "Sheet1",
        "C2",
        "precedents",
        "--depth",
        "6",
    ]);

    assert_invalid_argument(&[
        "formula-trace",
        file,
        "Sheet1",
        "C2",
        "precedents",
        "--page-size",
        "4",
    ]);

    assert_invalid_argument(&[
        "formula-trace",
        file,
        "Sheet1",
        "C2",
        "precedents",
        "--page-size",
        "201",
    ]);

    assert_invalid_argument(&[
        "formula-trace",
        file,
        "Sheet1",
        "C2",
        "precedents",
        "--cursor-depth",
        "0",
        "--cursor-offset",
        "0",
    ]);
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
    let canonical_values = canonical_payload["values"]
        .as_array()
        .expect("canonical single-range values");
    assert_eq!(canonical_values.len(), 1);
    let canonical_entry = canonical_values.first().expect("single range entry");
    assert_eq!(canonical_entry["range"], "A1:C4");
    assert!(canonical_entry.get("rows").is_some());

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
fn cli_range_values_shape_continuation_representable_canonical_and_compact() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("shape-continuation.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    // `A1:XFD1` is wider than the CLI max-cells cap (10_000), so the response keeps
    // a continuation cursor but no materialized row payload after pruning.
    let canonical = run_cli(&["range-values", file, "Sheet1", "A1:XFD1"]);
    assert!(canonical.status.success(), "stderr: {:?}", canonical.stderr);
    let canonical_payload = parse_stdout_json(&canonical);
    assert!(canonical_payload.get("workbook_id").is_some());
    assert!(canonical_payload.get("workbook_short_id").is_none());
    let canonical_values = canonical_payload["values"]
        .as_array()
        .expect("canonical continuation values");
    assert_eq!(canonical_values.len(), 1);
    let canonical_entry = canonical_values.first().expect("single continuation entry");
    assert_eq!(canonical_entry["range"], "A1:XFD1");
    assert_eq!(canonical_entry["next_start_row"].as_u64(), Some(1));

    let compact = run_cli(&[
        "--shape",
        "compact",
        "range-values",
        file,
        "Sheet1",
        "A1:XFD1",
    ]);
    assert!(compact.status.success(), "stderr: {:?}", compact.stderr);
    let compact_payload = parse_stdout_json(&compact);
    assert!(compact_payload.get("workbook_id").is_some());
    assert!(compact_payload.get("workbook_short_id").is_none());
    assert!(compact_payload.get("values").is_none());
    assert_eq!(compact_payload["range"], "A1:XFD1");
    assert_eq!(compact_payload["next_start_row"].as_u64(), Some(1));
}

#[test]
fn cli_range_values_invalid_range_omits_values_in_both_shapes() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("shape-invalid-range.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let canonical = run_cli(&["range-values", file, "Sheet1", "NOT_A_RANGE"]);
    assert!(canonical.status.success(), "stderr: {:?}", canonical.stderr);
    let canonical_payload = parse_stdout_json(&canonical);
    assert!(canonical_payload.get("workbook_id").is_some());
    assert!(canonical_payload.get("sheet_name").is_some());
    assert!(canonical_payload.get("values").is_none());

    let compact = run_cli(&[
        "--shape",
        "compact",
        "range-values",
        file,
        "Sheet1",
        "NOT_A_RANGE",
    ]);
    assert!(compact.status.success(), "stderr: {:?}", compact.stderr);
    let compact_payload = parse_stdout_json(&compact);
    assert!(compact_payload.get("workbook_id").is_some());
    assert!(compact_payload.get("sheet_name").is_some());
    assert!(compact_payload.get("values").is_none());
    assert!(compact_payload.get("range").is_none());
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
    assert!(canonical_values.iter().all(|entry| {
        entry.get("range").and_then(Value::as_str).is_some() && entry.get("rows").is_some()
    }));

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
    assert!(compact_payload.get("range").is_none());
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
fn cli_range_values_shape_default_matches_explicit_canonical() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("shape-default-canonical.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let default_shape = run_cli(&["range-values", file, "Sheet1", "A1:C4", "B1:B2"]);
    assert!(
        default_shape.status.success(),
        "stderr: {:?}",
        default_shape.stderr
    );

    let explicit_canonical = run_cli(&[
        "--shape",
        "canonical",
        "range-values",
        file,
        "Sheet1",
        "A1:C4",
        "B1:B2",
    ]);
    assert!(
        explicit_canonical.status.success(),
        "stderr: {:?}",
        explicit_canonical.stderr
    );

    let default_payload = parse_stdout_json(&default_shape);
    let canonical_payload = parse_stdout_json(&explicit_canonical);
    assert_eq!(default_payload, canonical_payload);
}

#[test]
fn cli_range_values_shape_compact_multi_range_preserves_next_start_row_without_flattening() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("shape-multi-continuation.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let compact = run_cli(&[
        "--shape",
        "compact",
        "range-values",
        file,
        "Sheet1",
        "A1:XFD1",
        "B1:B2",
    ]);
    assert!(compact.status.success(), "stderr: {:?}", compact.stderr);
    let compact_payload = parse_stdout_json(&compact);
    assert!(compact_payload.get("range").is_none());

    let compact_values = compact_payload["values"]
        .as_array()
        .expect("compact multi-range continuation values");
    assert_eq!(compact_values.len(), 2);

    let paged_entry = compact_values
        .iter()
        .find(|entry| entry.get("range").and_then(Value::as_str) == Some("A1:XFD1"))
        .expect("paged entry present");
    assert_eq!(paged_entry["next_start_row"].as_u64(), Some(1));
    assert!(
        compact_values
            .iter()
            .any(|entry| entry.get("range").and_then(Value::as_str) == Some("B1:B2"))
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

#[test]
fn cli_legacy_global_format_csv_returns_output_format_unsupported_envelope() {
    let output = run_cli(&["--format", "csv", "list-sheets", "/tmp/does-not-exist.xlsx"]);
    assert!(!output.status.success(), "command unexpectedly succeeded");

    let err = parse_stderr_json(&output);
    assert_eq!(err["code"], "OUTPUT_FORMAT_UNSUPPORTED");
    assert!(
        err["message"]
            .as_str()
            .unwrap_or_default()
            .contains("csv output is not implemented")
    );
}

#[test]
fn cli_legacy_global_format_json_is_accepted_for_existing_commands() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("legacy-format-json.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let output = run_cli(&["--format", "json", "list-sheets", file]);
    assert!(output.status.success(), "stderr: {:?}", output.stderr);

    let payload = parse_stdout_json(&output);
    assert_eq!(payload["sheets"].as_array().map(Vec::len), Some(2));
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
