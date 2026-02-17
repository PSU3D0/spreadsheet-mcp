use serde_json::Value;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use tempfile::tempdir;

#[cfg(unix)]
use std::os::unix::fs::{PermissionsExt, symlink};

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
    assert_error_code(args, "INVALID_ARGUMENT")
}

fn assert_error_code(args: &[&str], expected_code: &str) -> Value {
    let output = run_cli(args);
    assert!(
        !output.status.success(),
        "command unexpectedly succeeded for args: {args:?}"
    );
    let err = parse_stderr_json(&output);
    assert_eq!(
        err["code"], expected_code,
        "unexpected error envelope: {err}"
    );
    err
}

fn write_ops_payload(path: &Path, payload: &str) {
    fs::write(path, payload).expect("write ops payload");
}

fn repo_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
}

fn read_repo_doc(relative_path: &str) -> String {
    fs::read_to_string(repo_root().join(relative_path))
        .unwrap_or_else(|err| panic!("read {relative_path}: {err}"))
}

fn assert_batch_mode_matrix(command: &str, file: &str, ops_ref: &str) {
    assert_invalid_argument(&[command, file, "--ops", ops_ref]);
    assert_invalid_argument(&[command, file, "--ops", ops_ref, "--dry-run", "--in-place"]);
    assert_invalid_argument(&[
        command,
        file,
        "--ops",
        ops_ref,
        "--dry-run",
        "--output",
        "out.xlsx",
    ]);
    assert_invalid_argument(&[
        command,
        file,
        "--ops",
        ops_ref,
        "--in-place",
        "--output",
        "out.xlsx",
    ]);
    assert_invalid_argument(&[command, file, "--ops", ops_ref, "--force"]);
    assert_invalid_argument(&[command, file, "--ops", ops_ref, "--output", file]);
}

#[test]
fn cli_help_surfaces_include_descriptions_and_examples() {
    let root_help = run_cli(&["--help"]);
    assert!(root_help.status.success(), "stderr: {:?}", root_help.stderr);
    let root = parse_stdout_text(&root_help);
    assert!(root.contains("Stateless spreadsheet CLI for AI and automation workflows"));
    assert!(root.contains("Common workflows:"));
    assert!(
        root.contains("Inspect a workbook: list-sheets → sheet-overview → table-profile"),
        "missing inspect workflow anchor: {root}"
    );
    assert!(
        root.contains(
            "Deterministic pagination loops: sheet-page (--format + next_start_row) and read-table (--limit/--offset + next_offset)"
        ),
        "missing pagination workflow anchor: {root}"
    );
    assert!(
        root.contains(
            "Stateless batch writes: transform/style/formula/structure/column/layout/rules via --ops @ops.json + one mode (--dry-run|--in-place|--output)"
        ),
        "missing batch workflow anchor: {root}"
    );
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
        find.contains(
            "find-value data.xlsx \"Net Income\" --sheet \"Q1 Actuals\" --mode label --label-direction below"
        )
    );
    assert!(find.contains("Label mode behavior:"));
    assert!(find.contains("--label-direction any (default) checks right first, then below"));

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

    let sheet_page_help = run_cli(&["sheet-page", "--help"]);
    assert!(
        sheet_page_help.status.success(),
        "stderr: {:?}",
        sheet_page_help.stderr
    );
    let sheet_page = parse_stdout_text(&sheet_page_help);
    assert!(sheet_page.contains("Read one sheet page with deterministic continuation"));
    assert!(sheet_page.contains("Examples:"));
    assert!(sheet_page.contains("sheet-page data.xlsx Sheet1 --format compact --page-size 200"));
    assert!(
        sheet_page.contains(
            "sheet-page data.xlsx Sheet1 --format compact --page-size 200 --start-row 201"
        )
    );
    assert!(sheet_page.contains("Pagination loop:"));

    let read_table_help = run_cli(&["read-table", "--help"]);
    assert!(
        read_table_help.status.success(),
        "stderr: {:?}",
        read_table_help.stderr
    );
    let read_table = parse_stdout_text(&read_table_help);
    assert!(read_table.contains("Read a table-like region as json, values, or csv"));
    assert!(read_table.contains("Examples:"));
    assert!(
        read_table.contains(
            "read-table data.xlsx --sheet Sheet1 --table-format csv --limit 50 --offset 0"
        )
    );
    assert!(read_table.contains(
        "read-table data.xlsx --table-name SalesTable --sample-mode distributed --limit 20"
    ));
    assert!(read_table.contains("Repeat with --offset set to next_offset"));

    let formula_trace_help = run_cli(&["formula-trace", "--help"]);
    assert!(
        formula_trace_help.status.success(),
        "stderr: {:?}",
        formula_trace_help.stderr
    );
    let formula_trace = parse_stdout_text(&formula_trace_help);
    assert!(formula_trace.contains("Trace formula precedents or dependents from one origin cell"));
    assert!(formula_trace.contains("Examples:"));
    assert!(formula_trace.contains("formula-trace data.xlsx Sheet1 C2 precedents --depth 2"));
    assert!(formula_trace.contains(
        "formula-trace data.xlsx Sheet1 C2 precedents --cursor-depth 1 --cursor-offset 25"
    ));
    assert!(
        formula_trace.contains(
            "Reuse next_cursor.depth/next_cursor.offset as --cursor-depth/--cursor-offset"
        )
    );

    let transform_help = run_cli(&["transform-batch", "--help"]);
    assert!(
        transform_help.status.success(),
        "stderr: {:?}",
        transform_help.stderr
    );
    let transform = parse_stdout_text(&transform_help);
    assert!(transform.contains("Apply stateless transform operations from an @ops payload"));
    assert!(transform.contains("Examples:"));
    assert!(transform.contains("transform-batch workbook.xlsx --ops @ops.json --dry-run"));
    assert!(transform.contains(
        "transform-batch workbook.xlsx --ops @ops.json --output transformed.xlsx --force"
    ));
    assert!(transform.contains("Choose exactly one of --dry-run, --in-place, or --output <PATH>"),);
}

#[test]
fn readme_cli_docs_parity_examples_execute_with_local_fixtures() {
    let readme = read_repo_doc("README.md");
    for anchor in [
        "agent-spreadsheet sheet-page data.xlsx Sheet1 --format compact --page-size 200",
        "agent-spreadsheet read-table data.xlsx --sheet \"Sheet1\" --table-format values --limit 200 --offset 0",
        "agent-spreadsheet transform-batch data.xlsx --ops @ops.json --dry-run",
        "agent-spreadsheet style-batch data.xlsx --ops @style_ops.json --dry-run",
        "agent-spreadsheet find-value data.xlsx \"Net Income\" --mode label --label-direction below",
        "`sheet-page <file> <sheet> --format <full|compact|values_only>",
        "`range-values <file> <sheet> <range> [range...]`",
        "`find-value <file> <query> [--sheet S] [--mode value\\|label] [--label-direction right\\|below\\|any]`",
        "`transform-batch <file> --ops @ops.json (--dry-run|--in-place|--output PATH)`",
        "Compact (single range):** flatten that entry to top-level fields",
        "read-table and sheet-page: compact preserves the active branch and continuation fields (`next_offset`, `next_start_row`)",
        "Global `--output-format csv` is currently unsupported; use command-specific CSV options like `read-table --table-format csv`.",
        "`apply-formula-pattern` clears cached results for touched formula cells; run `recalculate` to refresh computed values.",
    ] {
        assert!(
            readme.contains(anchor),
            "missing README CLI anchor: {anchor}\n--- README excerpt check failed ---"
        );
    }
    assert!(
        !readme.contains("workbook_short_id"),
        "README should not advertise obsolete workbook_short_id fields"
    );

    let tmp = tempdir().expect("tempdir");
    let data_path = tmp.path().join("data.xlsx");
    let draft_path = tmp.path().join("draft.xlsx");
    let transform_ops_path = tmp.path().join("ops.json");
    let style_ops_path = tmp.path().join("style_ops.json");

    write_fixture(&data_path);
    write_ops_payload(
        &transform_ops_path,
        r#"{"ops":[{"kind":"fill_range","sheet_name":"Sheet1","target":{"kind":"cells","cells":["B2"]},"value":"77"}]}"#,
    );
    write_ops_payload(
        &style_ops_path,
        r#"{"ops":[{"sheet_name":"Sheet1","range":"B2:B2","style":{"font":{"bold":true}}}]}"#,
    );

    let file = data_path.to_str().expect("data path utf8");
    let draft = draft_path.to_str().expect("draft path utf8");
    let transform_ops_ref = format!("@{}", transform_ops_path.to_str().expect("ops utf8"));
    let style_ops_ref = format!("@{}", style_ops_path.to_str().expect("style ops utf8"));

    for args in [
        vec!["list-sheets", file],
        vec!["sheet-overview", file, "Sheet1"],
        vec![
            "read-table",
            file,
            "--sheet",
            "Sheet1",
            "--table-format",
            "values",
        ],
        vec![
            "sheet-page",
            file,
            "Sheet1",
            "--format",
            "compact",
            "--page-size",
            "2",
        ],
        vec![
            "sheet-page",
            file,
            "Sheet1",
            "--format",
            "compact",
            "--page-size",
            "2",
            "--start-row",
            "3",
        ],
        vec!["range-values", file, "Sheet1", "A1:C4"],
        vec![
            "find-value",
            file,
            "Amount",
            "--sheet",
            "Sheet1",
            "--mode",
            "label",
            "--label-direction",
            "below",
        ],
        vec![
            "transform-batch",
            file,
            "--ops",
            transform_ops_ref.as_str(),
            "--dry-run",
        ],
        vec![
            "style-batch",
            file,
            "--ops",
            style_ops_ref.as_str(),
            "--dry-run",
        ],
        vec!["copy", file, draft],
        vec!["edit", draft, "Sheet1", "B2=500", "C2==B2*1.1"],
        vec!["recalculate", draft],
        vec!["diff", file, draft],
    ] {
        let output = run_cli(args.as_slice());
        assert!(
            output.status.success(),
            "args={args:?}, stderr={:?}",
            output.stderr
        );
    }
}

#[test]
fn npm_readme_cli_docs_parity_examples_execute_with_local_fixtures() {
    let readme = read_repo_doc("npm/agent-spreadsheet/README.md");
    for anchor in [
        "agent-spreadsheet sheet-page data.xlsx Sheet1 --format compact --page-size 200",
        "agent-spreadsheet transform-batch data.xlsx --ops @ops.json --dry-run",
        "agent-spreadsheet find-value data.xlsx \"Net Income\" --mode label --label-direction below",
        "`sheet-page <file> <sheet> --format <full|compact|values_only>",
        "`find-value <file> <query> [--sheet S] [--mode value\\|label] [--label-direction right\\|below\\|any]`",
        "`transform-batch <file> --ops @ops.json (--dry-run|--in-place|--output PATH)`",
        "Canonical (default/omitted): return `values: [...]` when entries are present; omit `values` when all requested ranges are pruned (for example, invalid ranges).",
        "Global `--output-format csv` is currently unsupported; use command-specific CSV options such as `read-table --table-format csv`.",
        "`apply-formula-pattern` clears cached results for touched formula cells; run `recalculate` to refresh computed values.",
    ] {
        assert!(
            readme.contains(anchor),
            "missing npm README CLI anchor: {anchor}\n--- npm README excerpt check failed ---"
        );
    }
    assert!(
        !readme.contains("workbook_short_id"),
        "npm README should not advertise obsolete workbook_short_id fields"
    );

    let tmp = tempdir().expect("tempdir");
    let data_path = tmp.path().join("data.xlsx");
    let transform_ops_path = tmp.path().join("ops.json");
    write_fixture(&data_path);
    write_ops_payload(
        &transform_ops_path,
        r#"{"ops":[{"kind":"fill_range","sheet_name":"Sheet1","target":{"kind":"cells","cells":["B2"]},"value":"88"}]}"#,
    );

    let file = data_path.to_str().expect("data path utf8");
    let transform_ops_ref = format!("@{}", transform_ops_path.to_str().expect("ops utf8"));

    for args in [
        vec!["list-sheets", file],
        vec!["read-table", file, "--sheet", "Sheet1"],
        vec![
            "sheet-page",
            file,
            "Sheet1",
            "--format",
            "compact",
            "--page-size",
            "2",
        ],
        vec!["table-profile", file, "--sheet", "Sheet1"],
        vec![
            "find-value",
            file,
            "Amount",
            "--sheet",
            "Sheet1",
            "--mode",
            "label",
            "--label-direction",
            "below",
        ],
        vec![
            "transform-batch",
            file,
            "--ops",
            transform_ops_ref.as_str(),
            "--dry-run",
        ],
    ] {
        let output = run_cli(args.as_slice());
        assert!(
            output.status.success(),
            "args={args:?}, stderr={:?}",
            output.stderr
        );
    }
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
fn cli_find_value_label_mode_uses_query_as_label_and_direction() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("find-value-label-mode.xlsx");
    write_fixture(&workbook_path);
    let file = workbook_path.to_str().expect("path utf8");

    let below = run_cli(&[
        "find-value",
        file,
        "Amount",
        "--sheet",
        "Sheet1",
        "--mode",
        "label",
        "--label-direction",
        "below",
    ]);
    assert!(below.status.success(), "stderr: {:?}", below.stderr);
    let below_payload = parse_stdout_json(&below);
    let below_matches = below_payload["matches"].as_array().expect("matches array");
    assert_eq!(below_matches.len(), 1);
    assert_eq!(below_matches[0]["address"], "B1");
    assert_eq!(below_matches[0]["label_hit"]["label"], "Amount");
    assert_eq!(below_matches[0]["value"]["kind"], "Number");
    assert_eq!(below_matches[0]["value"]["value"], 10.0);

    let any = run_cli(&[
        "find-value",
        file,
        "Amount",
        "--sheet",
        "Sheet1",
        "--mode",
        "label",
    ]);
    assert!(any.status.success(), "stderr: {:?}", any.stderr);
    let any_payload = parse_stdout_json(&any);
    let any_matches = any_payload["matches"].as_array().expect("matches array");
    assert_eq!(any_matches.len(), 1);
    assert_eq!(any_matches[0]["address"], "B1");
    assert_eq!(any_matches[0]["value"]["kind"], "Text");
    assert_eq!(any_matches[0]["value"]["value"], "Total");
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
fn cli_phase1_scan_volatiles_skips_unparsable_formulas_instead_of_failing() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase1-scan-volatiles-parser-failure.xlsx");

    let mut workbook = umya_spreadsheet::new_file();
    {
        let sheet = workbook
            .get_sheet_by_name_mut("Sheet1")
            .expect("default sheet exists");
        sheet.get_cell_mut("A1").set_value("Input");
        sheet.get_cell_mut("B1").set_value("Result");
        // Intentionally malformed: one extra closing parenthesis.
        sheet.get_cell_mut("B2").set_formula(
            r#"IF(C70="","",IF(C70="N/A","",IF(C70="Unknown",0,IF(LEFT(C70,1)="0",0,IF(LEFT(C70,1)="1",25,IF(LEFT(C70,1)="2",50,IF(LEFT(C70,1)="3",75,IF(LEFT(C70,1)="4",100,"")))))))))"#,
        );
        sheet.get_cell_mut("B3").set_formula("NOW()");
    }
    umya_spreadsheet::writer::xlsx::write(&workbook, &workbook_path).expect("write workbook");
    let file = workbook_path.to_str().expect("path utf8");

    let output = run_cli(&["scan-volatiles", file, "--sheet", "Sheet1"]);
    assert!(output.status.success(), "stderr: {:?}", output.stderr);

    let payload = parse_stdout_json(&output);
    let items = payload["items"].as_array().expect("items array");
    assert!(
        items.iter().any(|item| {
            item["address"] == "B3"
                && item["function"] == "volatile"
                && item["sheet_name"] == "Sheet1"
        }),
        "expected volatile match from valid formula"
    );
}

#[test]
fn cli_formula_map_skips_unparsable_formulas_instead_of_failing() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("formula-map-parser-failure.xlsx");

    let mut workbook = umya_spreadsheet::new_file();
    {
        let sheet = workbook
            .get_sheet_by_name_mut("Sheet1")
            .expect("default sheet exists");
        sheet.get_cell_mut("A1").set_value("Input");
        sheet.get_cell_mut("B1").set_value("Result");
        // Intentionally malformed: one extra closing parenthesis.
        sheet.get_cell_mut("B2").set_formula(
            r#"IF(C70="","",IF(C70="N/A","",IF(C70="Unknown",0,IF(LEFT(C70,1)="0",0,IF(LEFT(C70,1)="1",25,IF(LEFT(C70,1)="2",50,IF(LEFT(C70,1)="3",75,IF(LEFT(C70,1)="4",100,"")))))))))"#,
        );
        sheet.get_cell_mut("B3").set_formula("SUM(1,2)");
    }
    umya_spreadsheet::writer::xlsx::write(&workbook, &workbook_path).expect("write workbook");
    let file = workbook_path.to_str().expect("path utf8");

    let output = run_cli(&["formula-map", file, "Sheet1", "--limit", "10"]);
    assert!(output.status.success(), "stderr: {:?}", output.stderr);

    let payload = parse_stdout_json(&output);
    let groups = payload["groups"].as_array().expect("groups array");
    assert!(
        !groups.is_empty(),
        "expected at least one parseable formula group"
    );
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
fn cli_transform_batch_dry_run_validates_contract_and_preserves_source() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("transform-batch-dry-run.xlsx");
    let ops_path = tmp.path().join("ops.json");
    write_fixture(&workbook_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"fill_range","sheet_name":"Sheet1","target":{"kind":"cells","cells":["B2"]},"value":"77"}]}"#,
    );

    let before = fs::read(&workbook_path).expect("read source before dry-run");
    let file = workbook_path.to_str().expect("path utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops path utf8"));

    let output = run_cli(&[
        "transform-batch",
        file,
        "--ops",
        ops_ref.as_str(),
        "--dry-run",
    ]);
    assert!(output.status.success(), "stderr: {:?}", output.stderr);
    let payload = parse_stdout_json(&output);

    assert_eq!(payload["op_count"].as_u64(), Some(1));
    assert_eq!(payload["validated_count"].as_u64(), Some(1));
    assert!(payload["would_change"].as_bool().unwrap_or(false));
    assert!(payload["warnings"].is_array());
    assert!(payload["summary"].is_object());
    assert!(payload["summary"]["operation_counts"].is_object());
    assert!(payload["summary"]["result_counts"].is_object());

    let after = fs::read(&workbook_path).expect("read source after dry-run");
    assert_eq!(before, after, "dry-run mutated the source workbook");
}

#[test]
fn cli_transform_batch_in_place_applies_atomically() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("transform-batch-in-place.xlsx");
    let ops_path = tmp.path().join("ops.json");
    write_fixture(&workbook_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"fill_range","sheet_name":"Sheet1","target":{"kind":"cells","cells":["B2"]},"value":"44"}]}"#,
    );

    let file = workbook_path.to_str().expect("path utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops path utf8"));

    let output = run_cli(&[
        "transform-batch",
        file,
        "--ops",
        ops_ref.as_str(),
        "--in-place",
    ]);
    assert!(output.status.success(), "stderr: {:?}", output.stderr);
    let payload = parse_stdout_json(&output);

    assert_eq!(payload["op_count"].as_u64(), Some(1));
    assert_eq!(payload["applied_count"].as_u64(), Some(1));
    assert!(payload["warnings"].is_array());
    assert!(payload["changed"].as_bool().unwrap_or(false));
    assert_eq!(payload["source_path"].as_str(), Some(file));
    assert_eq!(payload["target_path"].as_str(), Some(file));

    let book = umya_spreadsheet::reader::xlsx::read(&workbook_path).expect("read workbook");
    let sheet = book.get_sheet_by_name("Sheet1").expect("sheet exists");
    assert_eq!(sheet.get_cell("B2").expect("B2 exists").get_value(), "44");
}

#[test]
fn cli_transform_batch_output_and_force_modes_apply_with_overwrite_checks() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("transform-batch-source.xlsx");
    let output_path = tmp.path().join("transform-batch-output.xlsx");
    let ops_path_first = tmp.path().join("ops-first.json");
    let ops_path_second = tmp.path().join("ops-second.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path_first,
        r#"{"ops":[{"kind":"fill_range","sheet_name":"Sheet1","target":{"kind":"cells","cells":["B2"]},"value":"51"}]}"#,
    );
    write_ops_payload(
        &ops_path_second,
        r#"{"ops":[{"kind":"fill_range","sheet_name":"Sheet1","target":{"kind":"cells","cells":["B3"]},"value":"91"}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let ops_first_ref = format!("@{}", ops_path_first.to_str().expect("ops path utf8"));
    let ops_second_ref = format!("@{}", ops_path_second.to_str().expect("ops path utf8"));

    let first = run_cli(&[
        "transform-batch",
        source,
        "--ops",
        ops_first_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(first.status.success(), "stderr: {:?}", first.stderr);

    let source_book = umya_spreadsheet::reader::xlsx::read(&source_path).expect("read source");
    let source_sheet = source_book
        .get_sheet_by_name("Sheet1")
        .expect("sheet exists");
    assert_eq!(
        source_sheet
            .get_cell("B2")
            .expect("source B2 exists")
            .get_value(),
        "10"
    );

    let output_book = umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output");
    let output_sheet = output_book
        .get_sheet_by_name("Sheet1")
        .expect("sheet exists");
    assert_eq!(
        output_sheet
            .get_cell("B2")
            .expect("output B2 exists")
            .get_value(),
        "51"
    );

    assert_error_code(
        &[
            "transform-batch",
            source,
            "--ops",
            ops_second_ref.as_str(),
            "--output",
            output,
        ],
        "OUTPUT_EXISTS",
    );

    let forced = run_cli(&[
        "transform-batch",
        source,
        "--ops",
        ops_second_ref.as_str(),
        "--output",
        output,
        "--force",
    ]);
    assert!(forced.status.success(), "stderr: {:?}", forced.stderr);
    let forced_payload = parse_stdout_json(&forced);
    assert_eq!(forced_payload["target_path"].as_str(), Some(output));

    let overwritten = umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output");
    let overwritten_sheet = overwritten
        .get_sheet_by_name("Sheet1")
        .expect("sheet exists");
    assert_eq!(
        overwritten_sheet
            .get_cell("B3")
            .expect("output B3 exists")
            .get_value(),
        "91"
    );
}

#[cfg(unix)]
#[test]
fn cli_transform_batch_rejects_dangling_symlink_output_without_force() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp
        .path()
        .join("transform-batch-source-dangling-symlink.xlsx");
    let ops_path = tmp.path().join("ops.json");
    let output_link = tmp.path().join("dangling-output.xlsx");
    let missing_target = tmp.path().join("missing-target.xlsx");

    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"fill_range","sheet_name":"Sheet1","target":{"kind":"cells","cells":["B2"]},"value":"66"}]}"#,
    );

    symlink(&missing_target, &output_link).expect("create dangling symlink");

    let source = source_path.to_str().expect("source utf8");
    let output = output_link.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops path utf8"));

    let err = assert_error_code(
        &[
            "transform-batch",
            source,
            "--ops",
            ops_ref.as_str(),
            "--output",
            output,
        ],
        "OUTPUT_EXISTS",
    );
    assert!(
        err["message"]
            .as_str()
            .unwrap_or_default()
            .contains("already exists")
    );

    assert!(
        fs::symlink_metadata(&output_link).is_ok(),
        "dangling symlink should remain in place"
    );
}

#[test]
fn cli_transform_batch_rejects_invalid_mode_combinations() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("transform-batch-mode-matrix.xlsx");
    let ops_path = tmp.path().join("ops.json");
    write_fixture(&workbook_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"fill_range","sheet_name":"Sheet1","target":{"kind":"cells","cells":["B2"]},"value":"7"}]}"#,
    );

    let file = workbook_path.to_str().expect("path utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops path utf8"));

    assert_invalid_argument(&["transform-batch", file, "--ops", ops_ref.as_str()]);
    assert_invalid_argument(&[
        "transform-batch",
        file,
        "--ops",
        ops_ref.as_str(),
        "--dry-run",
        "--in-place",
    ]);
    assert_invalid_argument(&[
        "transform-batch",
        file,
        "--ops",
        ops_ref.as_str(),
        "--dry-run",
        "--output",
        "out.xlsx",
    ]);
    assert_invalid_argument(&[
        "transform-batch",
        file,
        "--ops",
        ops_ref.as_str(),
        "--in-place",
        "--output",
        "out.xlsx",
    ]);
    assert_invalid_argument(&[
        "transform-batch",
        file,
        "--ops",
        ops_ref.as_str(),
        "--force",
    ]);
    assert_invalid_argument(&[
        "transform-batch",
        file,
        "--ops",
        ops_ref.as_str(),
        "--output",
        file,
    ]);
}

#[test]
fn cli_transform_batch_rejects_invalid_ops_payloads() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("transform-batch-invalid-ops.xlsx");
    let malformed_path = tmp.path().join("ops-malformed.json");
    let schema_path = tmp.path().join("ops-schema.json");
    write_fixture(&workbook_path);
    write_ops_payload(&malformed_path, "{not-json}");
    write_ops_payload(&schema_path, r#"{"ops":[{"kind":"unknown_op"}]}"#);

    let file = workbook_path.to_str().expect("path utf8");

    assert_error_code(
        &["transform-batch", file, "--ops", "ops.json", "--dry-run"],
        "INVALID_OPS_PAYLOAD",
    );

    let malformed_ref = format!("@{}", malformed_path.to_str().expect("ops path utf8"));
    assert_error_code(
        &[
            "transform-batch",
            file,
            "--ops",
            malformed_ref.as_str(),
            "--dry-run",
        ],
        "INVALID_OPS_PAYLOAD",
    );

    let schema_ref = format!("@{}", schema_path.to_str().expect("ops path utf8"));
    assert_error_code(
        &[
            "transform-batch",
            file,
            "--ops",
            schema_ref.as_str(),
            "--dry-run",
        ],
        "INVALID_OPS_PAYLOAD",
    );
}

#[cfg(unix)]
#[test]
fn cli_transform_batch_maps_write_failures_and_preserves_source() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("transform-batch-write-fail-source.xlsx");
    let blocked_dir = tmp.path().join("blocked");
    let blocked_output = blocked_dir.join("output.xlsx");
    let ops_path = tmp.path().join("ops.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"fill_range","sheet_name":"Sheet1","target":{"kind":"cells","cells":["B2"]},"value":"123"}]}"#,
    );
    fs::create_dir(&blocked_dir).expect("create blocked dir");

    let mut perms = fs::metadata(&blocked_dir)
        .expect("blocked metadata")
        .permissions();
    perms.set_mode(0o555);
    fs::set_permissions(&blocked_dir, perms.clone()).expect("set blocked perms");

    let before = fs::read(&source_path).expect("read source before write failure");
    let source = source_path.to_str().expect("source utf8");
    let output = blocked_output.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops path utf8"));

    let err = assert_error_code(
        &[
            "transform-batch",
            source,
            "--ops",
            ops_ref.as_str(),
            "--output",
            output,
        ],
        "WRITE_FAILED",
    );
    assert!(
        err["message"]
            .as_str()
            .unwrap_or_default()
            .contains("unable to allocate temp file")
            || err["message"]
                .as_str()
                .unwrap_or_default()
                .contains("Permission denied")
    );

    let mut restore = perms;
    restore.set_mode(0o755);
    fs::set_permissions(&blocked_dir, restore).expect("restore blocked perms");

    let after = fs::read(&source_path).expect("read source after write failure");
    assert_eq!(before, after, "source workbook changed after write failure");
}

#[test]
fn phase_a_help_examples_for_style_and_formula_commands() {
    let style_help = run_cli(&["style-batch", "--help"]);
    assert!(
        style_help.status.success(),
        "stderr: {:?}",
        style_help.stderr
    );
    let style = parse_stdout_text(&style_help);
    assert!(style.contains("Examples:"));
    assert!(style.contains("style-batch workbook.xlsx --ops @style_ops.json --dry-run"));
    assert!(
        style.contains(
            "style-batch workbook.xlsx --ops @style_ops.json --output styled.xlsx --force"
        )
    );

    let formula_help = run_cli(&["apply-formula-pattern", "--help"]);
    assert!(
        formula_help.status.success(),
        "stderr: {:?}",
        formula_help.stderr
    );
    let formula = parse_stdout_text(&formula_help);
    assert!(formula.contains("Examples:"));
    assert!(
        formula.contains("apply-formula-pattern workbook.xlsx --ops @formula_ops.json --in-place")
    );
    assert!(
        formula.contains("apply-formula-pattern workbook.xlsx --ops @formula_ops.json --dry-run")
    );
    assert!(formula.contains(
        "Updated formula cells clear cached results. Run recalculate to refresh computed values."
    ));
}

#[test]
fn phase_a_style_batch_positive_dry_run_and_output_target_only() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-a-style-source.xlsx");
    let output_path = tmp.path().join("phase-a-style-output.xlsx");
    let ops_path = tmp.path().join("style-ops.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"sheet_name":"Sheet1","range":"B2:B2","style":{"font":{"bold":true}}}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops utf8"));

    let before = fs::read(&source_path).expect("read source before dry-run");
    let dry_run = run_cli(&[
        "style-batch",
        source,
        "--ops",
        ops_ref.as_str(),
        "--dry-run",
    ]);
    assert!(dry_run.status.success(), "stderr: {:?}", dry_run.stderr);
    let dry_payload = parse_stdout_json(&dry_run);
    assert_eq!(dry_payload["op_count"].as_u64(), Some(1));
    assert_eq!(dry_payload["validated_count"].as_u64(), Some(1));
    assert!(dry_payload["would_change"].as_bool().unwrap_or(false));

    let after_dry = fs::read(&source_path).expect("read source after dry-run");
    assert_eq!(before, after_dry, "dry-run mutated source file");

    let output_run = run_cli(&[
        "style-batch",
        source,
        "--ops",
        ops_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(
        output_run.status.success(),
        "stderr: {:?}",
        output_run.stderr
    );
    let output_payload = parse_stdout_json(&output_run);
    assert_eq!(output_payload["target_path"].as_str(), Some(output));
    assert_eq!(output_payload["source_path"].as_str(), Some(source));
    assert!(output_payload["changed"].as_bool().unwrap_or(false));

    let source_after = fs::read(&source_path).expect("read source after output mode");
    let output_after = fs::read(&output_path).expect("read output after output mode");
    assert_eq!(before, source_after, "source changed during --output mode");
    assert_ne!(source_after, output_after, "output file did not change");
}

#[test]
fn phase_a_style_batch_output_force_overwrite_semantics() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-a-style-force-source.xlsx");
    let output_path = tmp.path().join("phase-a-style-force-output.xlsx");
    let ops_first_path = tmp.path().join("style-ops-first.json");
    let ops_second_path = tmp.path().join("style-ops-second.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_first_path,
        r#"{"ops":[{"sheet_name":"Sheet1","range":"B2:B2","style":{"font":{"bold":true}}}]}"#,
    );
    write_ops_payload(
        &ops_second_path,
        r#"{"ops":[{"sheet_name":"Sheet1","range":"B2:B2","style":{"font":{"italic":true}}}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let first_ref = format!("@{}", ops_first_path.to_str().expect("ops utf8"));
    let second_ref = format!("@{}", ops_second_path.to_str().expect("ops utf8"));

    let first = run_cli(&[
        "style-batch",
        source,
        "--ops",
        first_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(first.status.success(), "stderr: {:?}", first.stderr);
    let first_bytes = fs::read(&output_path).expect("read first output bytes");

    assert_error_code(
        &[
            "style-batch",
            source,
            "--ops",
            second_ref.as_str(),
            "--output",
            output,
        ],
        "OUTPUT_EXISTS",
    );
    let after_failed_bytes = fs::read(&output_path).expect("read output after failed overwrite");
    assert_eq!(first_bytes, after_failed_bytes);

    let forced = run_cli(&[
        "style-batch",
        source,
        "--ops",
        second_ref.as_str(),
        "--output",
        output,
        "--force",
    ]);
    assert!(forced.status.success(), "stderr: {:?}", forced.stderr);

    let forced_bytes = fs::read(&output_path).expect("read forced output bytes");
    assert_ne!(
        first_bytes, forced_bytes,
        "force overwrite did not update output"
    );
}

#[test]
fn phase_a_apply_formula_pattern_positive_dry_run_and_output_target_only() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-a-formula-source.xlsx");
    let output_path = tmp.path().join("phase-a-formula-output.xlsx");
    let ops_path = tmp.path().join("formula-ops.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"sheet_name":"Sheet1","target_range":"C2:C4","anchor_cell":"C2","base_formula":"B2*3","fill_direction":"down","relative_mode":"excel"}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops utf8"));

    let before = fs::read(&source_path).expect("read source before dry-run");
    let dry_run = run_cli(&[
        "apply-formula-pattern",
        source,
        "--ops",
        ops_ref.as_str(),
        "--dry-run",
    ]);
    assert!(dry_run.status.success(), "stderr: {:?}", dry_run.stderr);
    let dry_payload = parse_stdout_json(&dry_run);
    assert_eq!(dry_payload["op_count"].as_u64(), Some(1));
    assert!(dry_payload["would_change"].as_bool().unwrap_or(false));
    let after_dry = fs::read(&source_path).expect("read source after dry-run");
    assert_eq!(before, after_dry, "dry-run mutated source file");

    let output_run = run_cli(&[
        "apply-formula-pattern",
        source,
        "--ops",
        ops_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(
        output_run.status.success(),
        "stderr: {:?}",
        output_run.stderr
    );
    let output_payload = parse_stdout_json(&output_run);
    assert!(output_payload["changed"].as_bool().unwrap_or(false));

    let source_book =
        umya_spreadsheet::reader::xlsx::read(&source_path).expect("read source workbook");
    let source_sheet = source_book
        .get_sheet_by_name("Sheet1")
        .expect("source sheet");
    assert_eq!(
        source_sheet
            .get_cell("C2")
            .expect("C2 source")
            .get_formula(),
        "B2*2"
    );

    let output_book =
        umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output workbook");
    let output_sheet = output_book
        .get_sheet_by_name("Sheet1")
        .expect("output sheet");
    assert_eq!(
        output_sheet
            .get_cell("C2")
            .expect("C2 output")
            .get_formula()
            .replace(' ', ""),
        "B2*3"
    );
    assert_eq!(
        output_sheet
            .get_cell("C3")
            .expect("C3 output")
            .get_formula()
            .replace(' ', ""),
        "B3*3"
    );
    assert_eq!(
        output_sheet
            .get_cell("C4")
            .expect("C4 output")
            .get_formula()
            .replace(' ', ""),
        "B4*3"
    );
}

#[test]
fn phase_a_apply_formula_pattern_clears_formula_cache_for_touched_cells() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase-a-formula-cache-clear.xlsx");
    let ops_path = tmp.path().join("formula-cache-ops.json");

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
        let c2 = sheet.get_cell_mut("C2");
        c2.set_formula("B2*2");
        c2.get_cell_value_mut().set_formula_result_default("20");
    }
    umya_spreadsheet::writer::xlsx::write(&workbook, &workbook_path).expect("write workbook");

    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"sheet_name":"Sheet1","target_range":"C2:C2","anchor_cell":"C2","base_formula":"B2*3","fill_direction":"down","relative_mode":"excel"}]}"#,
    );

    let file = workbook_path.to_str().expect("path utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops utf8"));

    let output = run_cli(&[
        "apply-formula-pattern",
        file,
        "--ops",
        ops_ref.as_str(),
        "--in-place",
    ]);
    assert!(output.status.success(), "stderr: {:?}", output.stderr);

    let book = umya_spreadsheet::reader::xlsx::read(&workbook_path).expect("read workbook");
    let sheet = book.get_sheet_by_name("Sheet1").expect("sheet exists");
    let c2 = sheet.get_cell("C2").expect("C2 exists");
    assert_eq!(c2.get_formula().replace(' ', ""), "B2*3");
    assert_eq!(c2.get_value(), "", "expected formula cache to be cleared");

    let read = run_cli(&["range-values", file, "Sheet1", "C2", "--shape", "compact"]);
    assert!(read.status.success(), "stderr: {:?}", read.stderr);
    let payload = parse_stdout_json(&read);
    assert!(
        payload["rows"][0][0].is_null(),
        "range-values should report null until recalculate refreshes cache"
    );
}

#[test]
fn phase_a_apply_formula_pattern_output_force_overwrite_semantics() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-a-force-source.xlsx");
    let output_path = tmp.path().join("phase-a-force-output.xlsx");
    let ops_first_path = tmp.path().join("formula-ops-first.json");
    let ops_second_path = tmp.path().join("formula-ops-second.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_first_path,
        r#"{"ops":[{"sheet_name":"Sheet1","target_range":"C2:C2","anchor_cell":"C2","base_formula":"B2*3","fill_direction":"down"}]}"#,
    );
    write_ops_payload(
        &ops_second_path,
        r#"{"ops":[{"sheet_name":"Sheet1","target_range":"C2:C2","anchor_cell":"C2","base_formula":"B2*5","fill_direction":"down"}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let first_ref = format!("@{}", ops_first_path.to_str().expect("ops utf8"));
    let second_ref = format!("@{}", ops_second_path.to_str().expect("ops utf8"));

    let first = run_cli(&[
        "apply-formula-pattern",
        source,
        "--ops",
        first_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(first.status.success(), "stderr: {:?}", first.stderr);

    assert_error_code(
        &[
            "apply-formula-pattern",
            source,
            "--ops",
            second_ref.as_str(),
            "--output",
            output,
        ],
        "OUTPUT_EXISTS",
    );

    let forced = run_cli(&[
        "apply-formula-pattern",
        source,
        "--ops",
        second_ref.as_str(),
        "--output",
        output,
        "--force",
    ]);
    assert!(forced.status.success(), "stderr: {:?}", forced.stderr);

    let output_book = umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output");
    let output_sheet = output_book
        .get_sheet_by_name("Sheet1")
        .expect("sheet exists");
    assert_eq!(
        output_sheet
            .get_cell("C2")
            .expect("C2 output")
            .get_formula()
            .replace(' ', ""),
        "B2*5"
    );
}

#[test]
fn phase_a_negative_invalid_ops_payloads() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase-a-invalid-ops.xlsx");
    let style_bad_path = tmp.path().join("style-bad.json");
    let formula_bad_path = tmp.path().join("formula-bad.json");
    write_fixture(&workbook_path);
    write_ops_payload(
        &style_bad_path,
        r#"{"ops":[{"sheet_name":"Sheet1","target":{"kind":"unknown"},"patch":{}}]}"#,
    );
    write_ops_payload(
        &formula_bad_path,
        r#"{"ops":[{"sheet_name":"Sheet1","target_range":"C2:C4","anchor_cell":"C1","base_formula":"B2*3","fill_direction":"down"}]}"#,
    );

    let file = workbook_path.to_str().expect("path utf8");
    let style_ref = format!("@{}", style_bad_path.to_str().expect("ops utf8"));
    let formula_ref = format!("@{}", formula_bad_path.to_str().expect("ops utf8"));

    assert_error_code(
        &[
            "style-batch",
            file,
            "--ops",
            style_ref.as_str(),
            "--dry-run",
        ],
        "INVALID_OPS_PAYLOAD",
    );
    assert_error_code(
        &[
            "apply-formula-pattern",
            file,
            "--ops",
            formula_ref.as_str(),
            "--dry-run",
        ],
        "INVALID_OPS_PAYLOAD",
    );
}

#[test]
fn phase_a_safety_mode_matrix_for_style_and_formula_commands() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase-a-safety.xlsx");
    let style_ops_path = tmp.path().join("style-ops.json");
    let formula_ops_path = tmp.path().join("formula-ops.json");
    write_fixture(&workbook_path);
    write_ops_payload(
        &style_ops_path,
        r#"{"ops":[{"sheet_name":"Sheet1","range":"B2:B2","style":{"font":{"bold":true}}}]}"#,
    );
    write_ops_payload(
        &formula_ops_path,
        r#"{"ops":[{"sheet_name":"Sheet1","target_range":"C2:C4","anchor_cell":"C2","base_formula":"B2*3","fill_direction":"down"}]}"#,
    );

    let file = workbook_path.to_str().expect("path utf8");
    let style_ref = format!("@{}", style_ops_path.to_str().expect("ops utf8"));
    let formula_ref = format!("@{}", formula_ops_path.to_str().expect("ops utf8"));

    assert_batch_mode_matrix("style-batch", file, style_ref.as_str());
    assert_batch_mode_matrix("apply-formula-pattern", file, formula_ref.as_str());
}

#[cfg(unix)]
#[test]
fn phase_a_style_batch_maps_write_failures_and_preserves_source() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-a-style-write-fail-source.xlsx");
    let blocked_dir = tmp.path().join("blocked");
    let blocked_output = blocked_dir.join("output.xlsx");
    let ops_path = tmp.path().join("ops.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"sheet_name":"Sheet1","range":"B2:B2","style":{"font":{"bold":true}}}]}"#,
    );
    fs::create_dir(&blocked_dir).expect("create blocked dir");

    let mut perms = fs::metadata(&blocked_dir)
        .expect("blocked metadata")
        .permissions();
    perms.set_mode(0o555);
    fs::set_permissions(&blocked_dir, perms.clone()).expect("set blocked perms");

    let before = fs::read(&source_path).expect("read source before write failure");
    let source = source_path.to_str().expect("source utf8");
    let output = blocked_output.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops path utf8"));

    assert_error_code(
        &[
            "style-batch",
            source,
            "--ops",
            ops_ref.as_str(),
            "--output",
            output,
        ],
        "WRITE_FAILED",
    );
    assert!(
        !blocked_output.exists(),
        "write failure left a partial output artifact"
    );

    let mut restore = perms;
    restore.set_mode(0o755);
    fs::set_permissions(&blocked_dir, restore).expect("restore blocked perms");

    let after = fs::read(&source_path).expect("read source after write failure");
    assert_eq!(before, after, "source workbook changed after write failure");
}

#[cfg(unix)]
#[test]
fn phase_a_apply_formula_pattern_maps_write_failures_and_preserves_source() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-a-formula-write-fail-source.xlsx");
    let blocked_dir = tmp.path().join("blocked-formula");
    let blocked_output = blocked_dir.join("output.xlsx");
    let ops_path = tmp.path().join("formula-ops.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"sheet_name":"Sheet1","target_range":"C2:C4","anchor_cell":"C2","base_formula":"B2*3","fill_direction":"down"}]}"#,
    );
    fs::create_dir(&blocked_dir).expect("create blocked dir");

    let mut perms = fs::metadata(&blocked_dir)
        .expect("blocked metadata")
        .permissions();
    perms.set_mode(0o555);
    fs::set_permissions(&blocked_dir, perms.clone()).expect("set blocked perms");

    let before = fs::read(&source_path).expect("read source before write failure");
    let source = source_path.to_str().expect("source utf8");
    let output = blocked_output.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops path utf8"));

    assert_error_code(
        &[
            "apply-formula-pattern",
            source,
            "--ops",
            ops_ref.as_str(),
            "--output",
            output,
        ],
        "WRITE_FAILED",
    );
    assert!(
        !blocked_output.exists(),
        "write failure left a partial output artifact"
    );

    let mut restore = perms;
    restore.set_mode(0o755);
    fs::set_permissions(&blocked_dir, restore).expect("restore blocked perms");

    let after = fs::read(&source_path).expect("read source after write failure");
    assert_eq!(before, after, "source workbook changed after write failure");
}

#[test]
fn phase_b_help_examples_for_structure_column_and_layout_commands() {
    let structure_help = run_cli(&["structure-batch", "--help"]);
    assert!(
        structure_help.status.success(),
        "stderr: {:?}",
        structure_help.stderr
    );
    let structure = parse_stdout_text(&structure_help);
    assert!(structure.contains("Examples:"));
    assert!(
        structure.contains("structure-batch workbook.xlsx --ops @structure_ops.json --dry-run")
    );
    assert!(structure.contains(
        "structure-batch workbook.xlsx --ops @structure_ops.json --output structured.xlsx"
    ));

    let column_help = run_cli(&["column-size-batch", "--help"]);
    assert!(
        column_help.status.success(),
        "stderr: {:?}",
        column_help.stderr
    );
    let column = parse_stdout_text(&column_help);
    assert!(column.contains("Examples:"));
    assert!(
        column.contains("column-size-batch workbook.xlsx --ops @column_size_ops.json --in-place")
    );
    assert!(column.contains(
        "column-size-batch workbook.xlsx --ops @column_size_ops.json --output columns.xlsx"
    ));

    let layout_help = run_cli(&["sheet-layout-batch", "--help"]);
    assert!(
        layout_help.status.success(),
        "stderr: {:?}",
        layout_help.stderr
    );
    let layout = parse_stdout_text(&layout_help);
    assert!(layout.contains("Examples:"));
    assert!(layout.contains("sheet-layout-batch workbook.xlsx --ops @layout_ops.json --dry-run"));
    assert!(layout.contains("sheet-layout-batch workbook.xlsx --ops @layout_ops.json --in-place"));
}

#[test]
fn phase_b_structure_batch_positive_in_place_renames_sheet() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase-b-structure-in-place.xlsx");
    let ops_path = tmp.path().join("structure-ops.json");
    write_fixture(&workbook_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"rename_sheet","old_name":"Summary","new_name":"Dashboard"}]}"#,
    );

    let file = workbook_path.to_str().expect("path utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops utf8"));

    let output = run_cli(&[
        "structure-batch",
        file,
        "--ops",
        ops_ref.as_str(),
        "--in-place",
    ]);
    assert!(output.status.success(), "stderr: {:?}", output.stderr);
    let payload = parse_stdout_json(&output);
    assert_eq!(payload["op_count"].as_u64(), Some(1));
    assert!(payload["changed"].as_bool().unwrap_or(false));

    let book = umya_spreadsheet::reader::xlsx::read(&workbook_path).expect("read workbook");
    assert!(book.get_sheet_by_name("Dashboard").is_some());
    assert!(book.get_sheet_by_name("Summary").is_none());
}

#[test]
fn phase_b_structure_batch_positive_dry_run_and_output_target_only() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-b-structure-source.xlsx");
    let output_path = tmp.path().join("phase-b-structure-output.xlsx");
    let ops_path = tmp.path().join("structure-ops-output.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"rename_sheet","old_name":"Summary","new_name":"Dashboard"}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops utf8"));

    let before = fs::read(&source_path).expect("read source before dry-run");

    let dry_run = run_cli(&[
        "structure-batch",
        source,
        "--ops",
        ops_ref.as_str(),
        "--dry-run",
    ]);
    assert!(dry_run.status.success(), "stderr: {:?}", dry_run.stderr);
    let dry_payload = parse_stdout_json(&dry_run);
    assert!(dry_payload["would_change"].as_bool().unwrap_or(false));

    let source_after_dry = fs::read(&source_path).expect("read source after dry-run");
    assert_eq!(before, source_after_dry, "dry-run mutated source workbook");

    let output_run = run_cli(&[
        "structure-batch",
        source,
        "--ops",
        ops_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(
        output_run.status.success(),
        "stderr: {:?}",
        output_run.stderr
    );
    let payload = parse_stdout_json(&output_run);
    assert!(payload["changed"].as_bool().unwrap_or(false));

    let source_book = umya_spreadsheet::reader::xlsx::read(&source_path).expect("read source");
    assert!(source_book.get_sheet_by_name("Summary").is_some());
    assert!(source_book.get_sheet_by_name("Dashboard").is_none());

    let output_book = umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output");
    assert!(output_book.get_sheet_by_name("Dashboard").is_some());
    assert!(output_book.get_sheet_by_name("Summary").is_none());
}

#[test]
fn phase_b_structure_batch_output_force_overwrite_semantics() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-b-structure-force-source.xlsx");
    let output_path = tmp.path().join("phase-b-structure-force-output.xlsx");
    let ops_first_path = tmp.path().join("structure-ops-first.json");
    let ops_second_path = tmp.path().join("structure-ops-second.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_first_path,
        r#"{"ops":[{"kind":"rename_sheet","old_name":"Summary","new_name":"Dashboard"}]}"#,
    );
    write_ops_payload(
        &ops_second_path,
        r#"{"ops":[{"kind":"rename_sheet","old_name":"Summary","new_name":"Board"}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let first_ref = format!("@{}", ops_first_path.to_str().expect("ops utf8"));
    let second_ref = format!("@{}", ops_second_path.to_str().expect("ops utf8"));

    let first = run_cli(&[
        "structure-batch",
        source,
        "--ops",
        first_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(first.status.success(), "stderr: {:?}", first.stderr);

    assert_error_code(
        &[
            "structure-batch",
            source,
            "--ops",
            second_ref.as_str(),
            "--output",
            output,
        ],
        "OUTPUT_EXISTS",
    );

    let forced = run_cli(&[
        "structure-batch",
        source,
        "--ops",
        second_ref.as_str(),
        "--output",
        output,
        "--force",
    ]);
    assert!(forced.status.success(), "stderr: {:?}", forced.stderr);

    let output_book = umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output");
    assert!(output_book.get_sheet_by_name("Board").is_some());
    assert!(output_book.get_sheet_by_name("Summary").is_none());
}

#[test]
fn phase_b_column_size_batch_positive_output_mutates_target_only() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-b-column-source.xlsx");
    let output_path = tmp.path().join("phase-b-column-output.xlsx");
    let ops_path = tmp.path().join("column-ops.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"sheet_name":"Sheet1","ops":[{"range":"A:A","size":{"kind":"width","width_chars":25.0}}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops utf8"));

    let before = fs::read(&source_path).expect("read source before dry-run");

    let dry_run = run_cli(&[
        "column-size-batch",
        source,
        "--ops",
        ops_ref.as_str(),
        "--dry-run",
    ]);
    assert!(dry_run.status.success(), "stderr: {:?}", dry_run.stderr);
    let dry_payload = parse_stdout_json(&dry_run);
    assert!(dry_payload["would_change"].as_bool().unwrap_or(false));

    let source_after_dry = fs::read(&source_path).expect("read source after dry-run");
    assert_eq!(before, source_after_dry, "dry-run mutated source workbook");

    let run = run_cli(&[
        "column-size-batch",
        source,
        "--ops",
        ops_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(run.status.success(), "stderr: {:?}", run.stderr);
    let payload = parse_stdout_json(&run);
    assert!(payload["changed"].as_bool().unwrap_or(false));

    let source_after = fs::read(&source_path).expect("read source after output mode");
    assert_eq!(before, source_after, "source changed during --output mode");

    let output_book =
        umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output workbook");
    let output_sheet = output_book
        .get_sheet_by_name("Sheet1")
        .expect("sheet exists");
    let width = *output_sheet
        .get_column_dimension("A")
        .expect("A column")
        .get_width();
    assert!((width - 25.0).abs() < 0.001);
}

#[test]
fn phase_b_column_size_batch_output_force_overwrite_semantics() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-b-column-force-source.xlsx");
    let output_path = tmp.path().join("phase-b-column-force-output.xlsx");
    let ops_first_path = tmp.path().join("column-ops-first.json");
    let ops_second_path = tmp.path().join("column-ops-second.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_first_path,
        r#"{"sheet_name":"Sheet1","ops":[{"range":"A:A","size":{"kind":"width","width_chars":25.0}}]}"#,
    );
    write_ops_payload(
        &ops_second_path,
        r#"{"sheet_name":"Sheet1","ops":[{"range":"A:A","size":{"kind":"width","width_chars":18.0}}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let first_ref = format!("@{}", ops_first_path.to_str().expect("ops utf8"));
    let second_ref = format!("@{}", ops_second_path.to_str().expect("ops utf8"));

    let first = run_cli(&[
        "column-size-batch",
        source,
        "--ops",
        first_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(first.status.success(), "stderr: {:?}", first.stderr);

    assert_error_code(
        &[
            "column-size-batch",
            source,
            "--ops",
            second_ref.as_str(),
            "--output",
            output,
        ],
        "OUTPUT_EXISTS",
    );

    let without_force_book =
        umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output without force");
    let without_force_sheet = without_force_book
        .get_sheet_by_name("Sheet1")
        .expect("sheet exists");
    let without_force_width = *without_force_sheet
        .get_column_dimension("A")
        .expect("A column")
        .get_width();
    assert!((without_force_width - 25.0).abs() < 0.001);

    let forced = run_cli(&[
        "column-size-batch",
        source,
        "--ops",
        second_ref.as_str(),
        "--output",
        output,
        "--force",
    ]);
    assert!(forced.status.success(), "stderr: {:?}", forced.stderr);

    let forced_book = umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output");
    let forced_sheet = forced_book
        .get_sheet_by_name("Sheet1")
        .expect("sheet exists");
    let forced_width = *forced_sheet
        .get_column_dimension("A")
        .expect("A column")
        .get_width();
    assert!((forced_width - 18.0).abs() < 0.001);
}

#[test]
fn phase_b_sheet_layout_batch_positive_dry_run_and_in_place() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase-b-layout.xlsx");
    let ops_path = tmp.path().join("layout-ops.json");
    write_fixture(&workbook_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"freeze_panes","sheet_name":"Sheet1","freeze_rows":1,"freeze_cols":1}]}"#,
    );

    let file = workbook_path.to_str().expect("path utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops utf8"));

    let before = fs::read(&workbook_path).expect("read before dry-run");
    let dry_run = run_cli(&[
        "sheet-layout-batch",
        file,
        "--ops",
        ops_ref.as_str(),
        "--dry-run",
    ]);
    assert!(dry_run.status.success(), "stderr: {:?}", dry_run.stderr);
    let dry_payload = parse_stdout_json(&dry_run);
    assert!(dry_payload["would_change"].as_bool().unwrap_or(false));
    let after_dry = fs::read(&workbook_path).expect("read after dry-run");
    assert_eq!(before, after_dry, "dry-run mutated workbook");

    let in_place = run_cli(&[
        "sheet-layout-batch",
        file,
        "--ops",
        ops_ref.as_str(),
        "--in-place",
    ]);
    assert!(in_place.status.success(), "stderr: {:?}", in_place.stderr);

    let book = umya_spreadsheet::reader::xlsx::read(&workbook_path).expect("read workbook");
    let sheet = book.get_sheet_by_name("Sheet1").expect("sheet exists");
    let views = sheet.get_sheets_views().get_sheet_view_list();
    let pane = views
        .first()
        .and_then(|view| view.get_pane())
        .expect("pane");
    assert_eq!(*pane.get_horizontal_split(), 1.0);
    assert_eq!(*pane.get_vertical_split(), 1.0);
    assert_eq!(pane.get_top_left_cell().to_string(), "B2");
}

#[test]
fn phase_b_sheet_layout_batch_positive_output_mutates_target_only() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-b-layout-source.xlsx");
    let output_path = tmp.path().join("phase-b-layout-output.xlsx");
    let ops_path = tmp.path().join("layout-output-ops.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"freeze_panes","sheet_name":"Sheet1","freeze_rows":1,"freeze_cols":1}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops utf8"));

    let before = fs::read(&source_path).expect("read source before output mode");

    let run = run_cli(&[
        "sheet-layout-batch",
        source,
        "--ops",
        ops_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(run.status.success(), "stderr: {:?}", run.stderr);
    let payload = parse_stdout_json(&run);
    assert!(payload["changed"].as_bool().unwrap_or(false));

    let source_after = fs::read(&source_path).expect("read source after output mode");
    assert_eq!(before, source_after, "source changed during --output mode");

    let output_book = umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output");
    let output_sheet = output_book
        .get_sheet_by_name("Sheet1")
        .expect("sheet exists");
    let pane = output_sheet
        .get_sheets_views()
        .get_sheet_view_list()
        .first()
        .and_then(|view| view.get_pane())
        .expect("pane");
    assert_eq!(pane.get_top_left_cell().to_string(), "B2");
}

#[test]
fn phase_b_sheet_layout_batch_output_force_overwrite_semantics() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-b-layout-force-source.xlsx");
    let output_path = tmp.path().join("phase-b-layout-force-output.xlsx");
    let ops_first_path = tmp.path().join("layout-ops-first.json");
    let ops_second_path = tmp.path().join("layout-ops-second.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_first_path,
        r#"{"ops":[{"kind":"freeze_panes","sheet_name":"Sheet1","freeze_rows":1,"freeze_cols":1}]}"#,
    );
    write_ops_payload(
        &ops_second_path,
        r#"{"ops":[{"kind":"freeze_panes","sheet_name":"Sheet1","freeze_rows":2,"freeze_cols":0}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let first_ref = format!("@{}", ops_first_path.to_str().expect("ops utf8"));
    let second_ref = format!("@{}", ops_second_path.to_str().expect("ops utf8"));

    let first = run_cli(&[
        "sheet-layout-batch",
        source,
        "--ops",
        first_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(first.status.success(), "stderr: {:?}", first.stderr);

    assert_error_code(
        &[
            "sheet-layout-batch",
            source,
            "--ops",
            second_ref.as_str(),
            "--output",
            output,
        ],
        "OUTPUT_EXISTS",
    );

    let without_force_book =
        umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output without force");
    let without_force_sheet = without_force_book
        .get_sheet_by_name("Sheet1")
        .expect("sheet exists");
    let without_force_pane = without_force_sheet
        .get_sheets_views()
        .get_sheet_view_list()
        .first()
        .and_then(|view| view.get_pane())
        .expect("pane without force");
    assert_eq!(without_force_pane.get_top_left_cell().to_string(), "B2");

    let forced = run_cli(&[
        "sheet-layout-batch",
        source,
        "--ops",
        second_ref.as_str(),
        "--output",
        output,
        "--force",
    ]);
    assert!(forced.status.success(), "stderr: {:?}", forced.stderr);

    let forced_book = umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output");
    let forced_sheet = forced_book
        .get_sheet_by_name("Sheet1")
        .expect("sheet exists");
    let forced_pane = forced_sheet
        .get_sheets_views()
        .get_sheet_view_list()
        .first()
        .and_then(|view| view.get_pane())
        .expect("forced pane");
    assert_eq!(forced_pane.get_top_left_cell().to_string(), "A3");
}

#[test]
fn phase_b_negative_invalid_ops_payloads() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase-b-invalid-ops.xlsx");
    let structure_bad_path = tmp.path().join("structure-bad.json");
    let column_bad_path = tmp.path().join("column-bad.json");
    let layout_bad_path = tmp.path().join("layout-bad.json");
    write_fixture(&workbook_path);
    write_ops_payload(&structure_bad_path, r#"{"ops":[{"kind":"unknown_kind"}]}"#);
    write_ops_payload(
        &column_bad_path,
        r#"{"ops":[{"range":"A:A","size":{"kind":"width","width_chars":12.0}}]}"#,
    );
    write_ops_payload(
        &layout_bad_path,
        r#"{"ops":[{"kind":"set_zoom","sheet_name":"Sheet1","zoom_percent":5}]}"#,
    );

    let file = workbook_path.to_str().expect("path utf8");
    let structure_ref = format!("@{}", structure_bad_path.to_str().expect("ops utf8"));
    let column_ref = format!("@{}", column_bad_path.to_str().expect("ops utf8"));
    let layout_ref = format!("@{}", layout_bad_path.to_str().expect("ops utf8"));

    assert_error_code(
        &[
            "structure-batch",
            file,
            "--ops",
            structure_ref.as_str(),
            "--dry-run",
        ],
        "INVALID_OPS_PAYLOAD",
    );
    assert_error_code(
        &[
            "column-size-batch",
            file,
            "--ops",
            column_ref.as_str(),
            "--dry-run",
        ],
        "INVALID_OPS_PAYLOAD",
    );
    assert_error_code(
        &[
            "sheet-layout-batch",
            file,
            "--ops",
            layout_ref.as_str(),
            "--dry-run",
        ],
        "INVALID_OPS_PAYLOAD",
    );
}

#[test]
fn phase_b_safety_mode_matrix_for_structure_column_layout_commands() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase-b-safety.xlsx");
    let structure_ops_path = tmp.path().join("structure-ops.json");
    let column_ops_path = tmp.path().join("column-ops.json");
    let layout_ops_path = tmp.path().join("layout-ops.json");
    write_fixture(&workbook_path);
    write_ops_payload(
        &structure_ops_path,
        r#"{"ops":[{"kind":"rename_sheet","old_name":"Summary","new_name":"Dashboard"}]}"#,
    );
    write_ops_payload(
        &column_ops_path,
        r#"{"sheet_name":"Sheet1","ops":[{"range":"A:A","size":{"kind":"width","width_chars":20.0}}]}"#,
    );
    write_ops_payload(
        &layout_ops_path,
        r#"{"ops":[{"kind":"freeze_panes","sheet_name":"Sheet1","freeze_rows":1,"freeze_cols":1}]}"#,
    );

    let file = workbook_path.to_str().expect("path utf8");
    let structure_ref = format!("@{}", structure_ops_path.to_str().expect("ops utf8"));
    let column_ref = format!("@{}", column_ops_path.to_str().expect("ops utf8"));
    let layout_ref = format!("@{}", layout_ops_path.to_str().expect("ops utf8"));

    assert_batch_mode_matrix("structure-batch", file, structure_ref.as_str());
    assert_batch_mode_matrix("column-size-batch", file, column_ref.as_str());
    assert_batch_mode_matrix("sheet-layout-batch", file, layout_ref.as_str());
}

#[cfg(unix)]
#[test]
fn phase_b_structure_batch_maps_write_failures_and_preserves_source() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-b-structure-write-fail-source.xlsx");
    let blocked_dir = tmp.path().join("blocked");
    let blocked_output = blocked_dir.join("output.xlsx");
    let ops_path = tmp.path().join("ops.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"rename_sheet","old_name":"Summary","new_name":"Dashboard"}]}"#,
    );
    fs::create_dir(&blocked_dir).expect("create blocked dir");

    let mut perms = fs::metadata(&blocked_dir)
        .expect("blocked metadata")
        .permissions();
    perms.set_mode(0o555);
    fs::set_permissions(&blocked_dir, perms.clone()).expect("set blocked perms");

    let before = fs::read(&source_path).expect("read source before write failure");
    let source = source_path.to_str().expect("source utf8");
    let output = blocked_output.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops path utf8"));

    assert_error_code(
        &[
            "structure-batch",
            source,
            "--ops",
            ops_ref.as_str(),
            "--output",
            output,
        ],
        "WRITE_FAILED",
    );
    assert!(
        !blocked_output.exists(),
        "write failure left a partial output artifact"
    );

    let mut restore = perms;
    restore.set_mode(0o755);
    fs::set_permissions(&blocked_dir, restore).expect("restore blocked perms");

    let after = fs::read(&source_path).expect("read source after write failure");
    assert_eq!(before, after, "source workbook changed after write failure");
}

#[cfg(unix)]
#[test]
fn phase_b_column_size_batch_maps_write_failures_and_preserves_source() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-b-column-write-fail-source.xlsx");
    let blocked_dir = tmp.path().join("blocked-column");
    let blocked_output = blocked_dir.join("output.xlsx");
    let ops_path = tmp.path().join("ops-column.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"sheet_name":"Sheet1","ops":[{"range":"A:A","size":{"kind":"width","width_chars":20.0}}]}"#,
    );
    fs::create_dir(&blocked_dir).expect("create blocked dir");

    let mut perms = fs::metadata(&blocked_dir)
        .expect("blocked metadata")
        .permissions();
    perms.set_mode(0o555);
    fs::set_permissions(&blocked_dir, perms.clone()).expect("set blocked perms");

    let before = fs::read(&source_path).expect("read source before write failure");
    let source = source_path.to_str().expect("source utf8");
    let output = blocked_output.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops path utf8"));

    assert_error_code(
        &[
            "column-size-batch",
            source,
            "--ops",
            ops_ref.as_str(),
            "--output",
            output,
        ],
        "WRITE_FAILED",
    );
    assert!(
        !blocked_output.exists(),
        "write failure left a partial output artifact"
    );

    let mut restore = perms;
    restore.set_mode(0o755);
    fs::set_permissions(&blocked_dir, restore).expect("restore blocked perms");

    let after = fs::read(&source_path).expect("read source after write failure");
    assert_eq!(before, after, "source workbook changed after write failure");
}

#[cfg(unix)]
#[test]
fn phase_b_sheet_layout_batch_maps_write_failures_and_preserves_source() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-b-layout-write-fail-source.xlsx");
    let blocked_dir = tmp.path().join("blocked-layout");
    let blocked_output = blocked_dir.join("output.xlsx");
    let ops_path = tmp.path().join("ops-layout.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"freeze_panes","sheet_name":"Sheet1","freeze_rows":1,"freeze_cols":1}]}"#,
    );
    fs::create_dir(&blocked_dir).expect("create blocked dir");

    let mut perms = fs::metadata(&blocked_dir)
        .expect("blocked metadata")
        .permissions();
    perms.set_mode(0o555);
    fs::set_permissions(&blocked_dir, perms.clone()).expect("set blocked perms");

    let before = fs::read(&source_path).expect("read source before write failure");
    let source = source_path.to_str().expect("source utf8");
    let output = blocked_output.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops path utf8"));

    assert_error_code(
        &[
            "sheet-layout-batch",
            source,
            "--ops",
            ops_ref.as_str(),
            "--output",
            output,
        ],
        "WRITE_FAILED",
    );
    assert!(
        !blocked_output.exists(),
        "write failure left a partial output artifact"
    );

    let mut restore = perms;
    restore.set_mode(0o755);
    fs::set_permissions(&blocked_dir, restore).expect("restore blocked perms");

    let after = fs::read(&source_path).expect("read source after write failure");
    assert_eq!(before, after, "source workbook changed after write failure");
}

#[test]
fn phase_c_help_examples_for_rules_command() {
    let rules_help = run_cli(&["rules-batch", "--help"]);
    assert!(
        rules_help.status.success(),
        "stderr: {:?}",
        rules_help.stderr
    );
    let rules = parse_stdout_text(&rules_help);
    assert!(rules.contains("Examples:"));
    assert!(rules.contains("rules-batch workbook.xlsx --ops @rules_ops.json --dry-run"));
    assert!(
        rules.contains(
            "rules-batch workbook.xlsx --ops @rules_ops.json --output ruled.xlsx --force"
        )
    );
}

#[test]
fn phase_c_rules_batch_positive_in_place_sets_validation() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase-c-rules-in-place.xlsx");
    let ops_path = tmp.path().join("rules-ops.json");
    write_fixture(&workbook_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"set_data_validation","sheet_name":"Sheet1","target_range":"B2:B4","validation":{"kind":"list","formula1":"\"A,B,C\""}}]}"#,
    );

    let file = workbook_path.to_str().expect("path utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops utf8"));

    let output = run_cli(&["rules-batch", file, "--ops", ops_ref.as_str(), "--in-place"]);
    assert!(output.status.success(), "stderr: {:?}", output.stderr);
    let payload = parse_stdout_json(&output);
    assert!(payload["changed"].as_bool().unwrap_or(false));

    let book = umya_spreadsheet::reader::xlsx::read(&workbook_path).expect("read workbook");
    let sheet = book.get_sheet_by_name("Sheet1").expect("sheet exists");
    let dvs = sheet.get_data_validations().expect("data validations");
    let list = dvs.get_data_validation_list();
    assert_eq!(list.len(), 1);
    assert_eq!(list[0].get_sequence_of_references().get_sqref(), "B2:B4");
}

#[test]
fn phase_c_rules_batch_positive_dry_run_and_output_target_only() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-c-rules-source.xlsx");
    let output_path = tmp.path().join("phase-c-rules-output.xlsx");
    let ops_path = tmp.path().join("rules-ops-output.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"set_data_validation","sheet_name":"Sheet1","target_range":"C2:C4","validation":{"kind":"list","formula1":"\"X,Y,Z\""}}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops utf8"));

    let before = fs::read(&source_path).expect("read source before dry-run");

    let dry_run = run_cli(&[
        "rules-batch",
        source,
        "--ops",
        ops_ref.as_str(),
        "--dry-run",
    ]);
    assert!(dry_run.status.success(), "stderr: {:?}", dry_run.stderr);
    let dry_payload = parse_stdout_json(&dry_run);
    assert!(dry_payload["would_change"].as_bool().unwrap_or(false));

    let source_after_dry = fs::read(&source_path).expect("read source after dry-run");
    assert_eq!(before, source_after_dry, "dry-run mutated source workbook");

    let output_run = run_cli(&[
        "rules-batch",
        source,
        "--ops",
        ops_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(
        output_run.status.success(),
        "stderr: {:?}",
        output_run.stderr
    );

    let source_after_output = fs::read(&source_path).expect("read source after output mode");
    assert_eq!(
        before, source_after_output,
        "source changed during --output mode"
    );

    let output_book = umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output");
    let output_sheet = output_book
        .get_sheet_by_name("Sheet1")
        .expect("sheet exists");
    let dvs = output_sheet
        .get_data_validations()
        .expect("data validations");
    let list = dvs.get_data_validation_list();
    assert_eq!(list.len(), 1);
    assert_eq!(list[0].get_sequence_of_references().get_sqref(), "C2:C4");
}

#[test]
fn phase_c_rules_batch_output_force_overwrite_semantics() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-c-rules-force-source.xlsx");
    let output_path = tmp.path().join("phase-c-rules-force-output.xlsx");
    let ops_first_path = tmp.path().join("rules-ops-first.json");
    let ops_second_path = tmp.path().join("rules-ops-second.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_first_path,
        r#"{"ops":[{"kind":"set_data_validation","sheet_name":"Sheet1","target_range":"B2:B4","validation":{"kind":"list","formula1":"\"A,B,C\""}}]}"#,
    );
    write_ops_payload(
        &ops_second_path,
        r#"{"ops":[{"kind":"set_data_validation","sheet_name":"Sheet1","target_range":"C2:C4","validation":{"kind":"list","formula1":"\"X,Y,Z\""}}]}"#,
    );

    let source = source_path.to_str().expect("source utf8");
    let output = output_path.to_str().expect("output utf8");
    let first_ref = format!("@{}", ops_first_path.to_str().expect("ops utf8"));
    let second_ref = format!("@{}", ops_second_path.to_str().expect("ops utf8"));

    let first = run_cli(&[
        "rules-batch",
        source,
        "--ops",
        first_ref.as_str(),
        "--output",
        output,
    ]);
    assert!(first.status.success(), "stderr: {:?}", first.stderr);

    assert_error_code(
        &[
            "rules-batch",
            source,
            "--ops",
            second_ref.as_str(),
            "--output",
            output,
        ],
        "OUTPUT_EXISTS",
    );

    let forced = run_cli(&[
        "rules-batch",
        source,
        "--ops",
        second_ref.as_str(),
        "--output",
        output,
        "--force",
    ]);
    assert!(forced.status.success(), "stderr: {:?}", forced.stderr);

    let output_book = umya_spreadsheet::reader::xlsx::read(&output_path).expect("read output");
    let output_sheet = output_book
        .get_sheet_by_name("Sheet1")
        .expect("sheet exists");
    let dvs = output_sheet
        .get_data_validations()
        .expect("data validations");
    let list = dvs.get_data_validation_list();
    assert_eq!(list.len(), 1);
    assert_eq!(list[0].get_sequence_of_references().get_sqref(), "C2:C4");
}

#[test]
fn phase_c_negative_invalid_ops_payload() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase-c-invalid-ops.xlsx");
    let bad_ops_path = tmp.path().join("rules-bad.json");
    write_fixture(&workbook_path);
    write_ops_payload(&bad_ops_path, r#"{"ops":[{"kind":"unknown_rule"}]}"#);

    let file = workbook_path.to_str().expect("path utf8");
    let ops_ref = format!("@{}", bad_ops_path.to_str().expect("ops utf8"));
    assert_error_code(
        &["rules-batch", file, "--ops", ops_ref.as_str(), "--dry-run"],
        "INVALID_OPS_PAYLOAD",
    );
}

#[test]
fn phase_c_safety_mode_matrix_for_rules_command() {
    let tmp = tempdir().expect("tempdir");
    let workbook_path = tmp.path().join("phase-c-safety.xlsx");
    let ops_path = tmp.path().join("rules-ops.json");
    write_fixture(&workbook_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"set_data_validation","sheet_name":"Sheet1","target_range":"B2:B4","validation":{"kind":"list","formula1":"\"A,B,C\""}}]}"#,
    );

    let file = workbook_path.to_str().expect("path utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops utf8"));
    assert_batch_mode_matrix("rules-batch", file, ops_ref.as_str());
}

#[cfg(unix)]
#[test]
fn phase_c_rules_batch_maps_write_failures_and_preserves_source() {
    let tmp = tempdir().expect("tempdir");
    let source_path = tmp.path().join("phase-c-rules-write-fail-source.xlsx");
    let blocked_dir = tmp.path().join("blocked");
    let blocked_output = blocked_dir.join("output.xlsx");
    let ops_path = tmp.path().join("ops.json");
    write_fixture(&source_path);
    write_ops_payload(
        &ops_path,
        r#"{"ops":[{"kind":"set_data_validation","sheet_name":"Sheet1","target_range":"B2:B4","validation":{"kind":"list","formula1":"\"A,B,C\""}}]}"#,
    );
    fs::create_dir(&blocked_dir).expect("create blocked dir");

    let mut perms = fs::metadata(&blocked_dir)
        .expect("blocked metadata")
        .permissions();
    perms.set_mode(0o555);
    fs::set_permissions(&blocked_dir, perms.clone()).expect("set blocked perms");

    let before = fs::read(&source_path).expect("read source before write failure");
    let source = source_path.to_str().expect("source utf8");
    let output = blocked_output.to_str().expect("output utf8");
    let ops_ref = format!("@{}", ops_path.to_str().expect("ops path utf8"));

    assert_error_code(
        &[
            "rules-batch",
            source,
            "--ops",
            ops_ref.as_str(),
            "--output",
            output,
        ],
        "WRITE_FAILED",
    );
    assert!(
        !blocked_output.exists(),
        "write failure left a partial output artifact"
    );

    let mut restore = perms;
    restore.set_mode(0o755);
    fs::set_permissions(&blocked_dir, restore).expect("restore blocked perms");

    let after = fs::read(&source_path).expect("read source after write failure");
    assert_eq!(before, after, "source workbook changed after write failure");
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
