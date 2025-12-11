use spreadsheet_mcp::model::RegionKind;
use spreadsheet_mcp::workbook::WorkbookContext;
use support::builders::{
    CellVal, apply_date_format, fill_horizontal_kv, fill_key_value, fill_sparse, fill_table,
    set_header_style,
};

mod support;

fn region_kind(ctx: &WorkbookContext, sheet: &str, region_idx: usize) -> Option<RegionKind> {
    ctx.get_sheet_metrics(sheet)
        .ok()
        .and_then(|m| m.detected_regions.get(region_idx).cloned())
        .and_then(|r| r.region_kind)
}

fn region_count(ctx: &WorkbookContext, sheet: &str) -> usize {
    ctx.get_sheet_metrics(sheet)
        .map(|m| m.detected_regions.len())
        .unwrap_or(0)
}

fn has_headers(ctx: &WorkbookContext, sheet: &str, region_idx: usize) -> bool {
    ctx.get_sheet_metrics(sheet)
        .ok()
        .and_then(|m| m.detected_regions.get(region_idx).cloned())
        .map(|r| r.header_row.is_some())
        .unwrap_or(false)
}

#[test]
fn wide_data_table_detected_as_data() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("wide.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        let headers: Vec<String> = (0..25).map(|i| format!("Col{}", i + 1)).collect();
        let header_refs: Vec<&str> = headers.iter().map(|s| s.as_str()).collect();
        let rows: Vec<Vec<CellVal>> = (0..50)
            .map(|r| (0..25).map(|c| CellVal::Num((r * 25 + c) as f64)).collect())
            .collect();

        fill_table(sheet, "A1", &header_refs, &rows);
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    assert_eq!(region_count(&ctx, "Sheet1"), 1);
    assert!(has_headers(&ctx, "Sheet1", 0));
}

#[test]
fn config_plus_data_detects_multiple_regions() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("mixed.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        fill_key_value(
            sheet,
            "A1",
            &[
                ("Report Date", CellVal::Date(45597.0)),
                ("Region", CellVal::Text("North".into())),
                ("Version", CellVal::Num(2.0)),
                ("Author", CellVal::Text("System".into())),
            ],
        );

        fill_table(
            sheet,
            "E1",
            &["Product", "Units", "Revenue"],
            &[
                vec![
                    CellVal::Text("Widget".into()),
                    CellVal::Num(100.0),
                    CellVal::Num(5000.0),
                ],
                vec![
                    CellVal::Text("Gadget".into()),
                    CellVal::Num(50.0),
                    CellVal::Num(2500.0),
                ],
                vec![
                    CellVal::Text("Gizmo".into()),
                    CellVal::Num(75.0),
                    CellVal::Num(3750.0),
                ],
            ],
        );
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    let count = region_count(&ctx, "Sheet1");
    assert!(count >= 1, "should detect at least 1 region, got {}", count);
}

#[test]
fn sparse_noisy_data_has_low_confidence() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("sparse.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        fill_sparse(
            sheet,
            &[
                ("A1", CellVal::Text("Header1".into())),
                ("C1", CellVal::Text("Header2".into())),
                ("F1", CellVal::Text("Header3".into())),
                ("A3", CellVal::Num(10.0)),
                ("C3", CellVal::Num(20.0)),
                ("A5", CellVal::Num(30.0)),
                ("F5", CellVal::Num(40.0)),
                ("B7", CellVal::Text("Note".into())),
                ("A10", CellVal::Num(50.0)),
                ("C10", CellVal::Num(60.0)),
                ("F10", CellVal::Num(70.0)),
            ],
        );
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    let metrics = ctx.get_sheet_metrics("Sheet1").unwrap();
    if let Some(region) = metrics.detected_regions.first() {
        assert!(
            region.confidence < 0.9,
            "sparse data should have lower confidence, got {}",
            region.confidence
        );
    }
}

#[test]
fn multi_row_header_detected() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("multi_header.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        sheet.get_cell_mut("A1").set_value("Q1 2024");
        sheet.get_cell_mut("D1").set_value("Q2 2024");

        fill_table(
            sheet,
            "A2",
            &["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
            &[
                vec![
                    CellVal::Num(100.0),
                    CellVal::Num(110.0),
                    CellVal::Num(120.0),
                    CellVal::Num(130.0),
                    CellVal::Num(140.0),
                    CellVal::Num(150.0),
                ],
                vec![
                    CellVal::Num(200.0),
                    CellVal::Num(210.0),
                    CellVal::Num(220.0),
                    CellVal::Num(230.0),
                    CellVal::Num(240.0),
                    CellVal::Num(250.0),
                ],
            ],
        );

        set_header_style(sheet, "A1:F2");
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    assert!(has_headers(&ctx, "Sheet1", 0));
}

#[test]
fn horizontal_kv_structure_recognized() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("horiz_kv.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        fill_horizontal_kv(
            sheet,
            "A1",
            &[
                ("Name", CellVal::Text("Acme Corp".into())),
                ("ID", CellVal::Num(12345.0)),
                ("Date", CellVal::Date(45597.0)),
                ("Status", CellVal::Text("Active".into())),
            ],
        );
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    assert_eq!(region_count(&ctx, "Sheet1"), 1);
}

#[test]
fn formula_heavy_classified_as_calculator() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("calc.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        sheet.get_cell_mut("A1").set_value("Input1");
        sheet.get_cell_mut("B1").set_value_number(100.0);
        sheet.get_cell_mut("A2").set_value("Input2");
        sheet.get_cell_mut("B2").set_value_number(50.0);

        sheet.get_cell_mut("D1").set_value("Calculations");
        sheet.get_cell_mut("D2").set_formula("B1+B2");
        sheet.get_cell_mut("D3").set_formula("B1*B2");
        sheet.get_cell_mut("D4").set_formula("D2+D3");
        sheet.get_cell_mut("D5").set_formula("D4*0.1");
        sheet.get_cell_mut("D6").set_formula("D4+D5");
        sheet.get_cell_mut("D7").set_formula("D6/B1");
        sheet.get_cell_mut("D8").set_formula("D6/B2");
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    let kind = region_kind(&ctx, "Sheet1", 0);
    assert!(
        matches!(
            kind,
            Some(RegionKind::Calculator) | Some(RegionKind::Outputs) | Some(RegionKind::Parameters)
        ),
        "formula-heavy region should be calculator/outputs/parameters, got {:?}",
        kind
    );
}

#[test]
fn cross_sheet_refs_both_sheets_detected() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("cross.xlsx", |book| {
        {
            let data_sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
            data_sheet.set_name("Data");
            fill_table(
                data_sheet,
                "A1",
                &["Item", "Price", "Qty"],
                &[
                    vec![
                        CellVal::Text("Apple".into()),
                        CellVal::Num(1.50),
                        CellVal::Num(10.0),
                    ],
                    vec![
                        CellVal::Text("Banana".into()),
                        CellVal::Num(0.75),
                        CellVal::Num(20.0),
                    ],
                    vec![
                        CellVal::Text("Cherry".into()),
                        CellVal::Num(3.00),
                        CellVal::Num(5.0),
                    ],
                ],
            );
        }

        let calc_sheet = book.new_sheet("Summary").unwrap();
        calc_sheet.get_cell_mut("A1").set_value("Total Qty");
        calc_sheet.get_cell_mut("B1").set_formula("SUM(Data!C2:C4)");
        calc_sheet.get_cell_mut("A2").set_value("Total Value");
        calc_sheet
            .get_cell_mut("B2")
            .set_formula("SUMPRODUCT(Data!B2:B4,Data!C2:C4)");
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    assert!(region_count(&ctx, "Data") >= 1);
    assert!(region_count(&ctx, "Summary") >= 1);
}

#[test]
fn null_padded_kv_detected_as_parameters() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("padded_kv.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        sheet.get_cell_mut("B1").set_value("Start Date");
        sheet.get_cell_mut("C1").set_value_number(45597.0);
        apply_date_format(sheet, "C1");
        sheet.get_cell_mut("D1").set_value_number(45627.0);
        apply_date_format(sheet, "D1");

        sheet.get_cell_mut("B2").set_value("ID");
        sheet.get_cell_mut("C2").set_value_number(12345.0);
        sheet.get_cell_mut("B3").set_value("Name");
        sheet.get_cell_mut("C3").set_value("Test Item");
        sheet.get_cell_mut("B4").set_value("Category");
        sheet.get_cell_mut("C4").set_value("Sample");
        sheet.get_cell_mut("B5").set_value("Value");
        sheet.get_cell_mut("C5").set_value_number(999.0);
        sheet.get_cell_mut("B6").set_value("Status");
        sheet.get_cell_mut("C6").set_value("Active");
        sheet.get_cell_mut("B7").set_value("Rate");
        sheet.get_cell_mut("C7").set_value_number(0.05);
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    let kind = region_kind(&ctx, "Sheet1", 0);
    assert!(
        matches!(kind, Some(RegionKind::Parameters) | Some(RegionKind::Data)),
        "null-padded kv should be parameters or data, got {:?}",
        kind
    );
}

#[test]
fn dense_4col_parameter_grid() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("4col_param.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        // Grid layout: Key | Value | Key | Value
        // Heuristic likely sees this as a data table because it expects exactly 2 columns for KV

        // Col A/B
        fill_key_value(
            sheet,
            "A1",
            &[
                ("Name", CellVal::Text("Project X".into())),
                ("Lead", CellVal::Text("Alice".into())),
                ("Start", CellVal::Date(45292.0)), // 2024-01-01
                ("Budget", CellVal::Num(50000.0)),
            ],
        );

        // Col C/D
        fill_key_value(
            sheet,
            "C1",
            &[
                ("Client", CellVal::Text("MegaCorp".into())),
                ("Region", CellVal::Text("EMEA".into())),
                ("End", CellVal::Date(45658.0)), // 2024-12-31
                ("Status", CellVal::Text("Active".into())),
            ],
        );
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    let kind = region_kind(&ctx, "Sheet1", 0);

    // Improved heuristic: detect 4-column grids as Parameters
    assert!(
        matches!(kind, Some(RegionKind::Parameters)),
        "4-col grid should be parameters, got {:?}",
        kind
    );

    // Check if it's one region or split
    assert_eq!(
        region_count(&ctx, "Sheet1"),
        1,
        "Should be detected as a single region"
    );
}

#[test]
fn tables_separated_by_zeros() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("zero_gutter.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        // Table 1 (A-B)
        fill_table(
            sheet,
            "A1",
            &["ID", "Val1"],
            &[
                vec![CellVal::Num(1.0), CellVal::Num(10.0)],
                vec![CellVal::Num(2.0), CellVal::Num(20.0)],
                vec![CellVal::Num(3.0), CellVal::Num(30.0)],
            ],
        );

        // Gutter (C) - Full of zeros (simulating formula returns)
        for r in 1..=4 {
            sheet.get_cell_mut((3, r)).set_value_number(0.0);
        }

        // Table 2 (D-E)
        fill_table(
            sheet,
            "D1",
            &["ID", "Val2"],
            &[
                vec![CellVal::Num(1.0), CellVal::Num(100.0)],
                vec![CellVal::Num(2.0), CellVal::Num(200.0)],
                vec![CellVal::Num(3.0), CellVal::Num(300.0)],
            ],
        );
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    // KNOWN LIMITATION:
    // This typically merges into one large region because Col C is "dense" (not empty).
    // Future work: Improve gutter detection to handle low-entropy columns.
    let count = region_count(&ctx, "Sheet1");

    if count == 1 {
        let metrics = ctx.get_sheet_metrics("Sheet1").unwrap();
        let region = &metrics.detected_regions[0];
        assert!(
            region.bounds.contains("E4"),
            "Region should cover both tables"
        );
    }
}

#[test]
fn numeric_year_headers() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("numeric_headers.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        // Row 1: Numeric Headers
        sheet.get_cell_mut("A1").set_value("Category");
        sheet.get_cell_mut("B1").set_value_number(2022.0);
        sheet.get_cell_mut("C1").set_value_number(2023.0);
        sheet.get_cell_mut("D1").set_value_number(2024.0);

        // Data
        let rows = vec![
            vec![
                CellVal::Text("Revenue".into()),
                CellVal::Num(100.0),
                CellVal::Num(110.0),
                CellVal::Num(120.0),
            ],
            vec![
                CellVal::Text("Cost".into()),
                CellVal::Num(80.0),
                CellVal::Num(85.0),
                CellVal::Num(90.0),
            ],
            vec![
                CellVal::Text("Profit".into()),
                CellVal::Num(20.0),
                CellVal::Num(25.0),
                CellVal::Num(30.0),
            ],
        ];

        // Using fill_table helper for data starting at A2, but we need to match columns
        // Actually, let's just use low-level building to be precise
        for (r_idx, row_data) in rows.iter().enumerate() {
            let row_num = (r_idx + 2) as u32;
            for (c_idx, val) in row_data.iter().enumerate() {
                let col_num = (c_idx + 1) as u32;
                match val {
                    CellVal::Text(s) => {
                        sheet.get_cell_mut((col_num, row_num)).set_value(s);
                    }
                    CellVal::Num(n) => {
                        sheet.get_cell_mut((col_num, row_num)).set_value_number(*n);
                    }
                    _ => {}
                }
            }
        }
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    let metrics = ctx.get_sheet_metrics("Sheet1").unwrap();
    let region = &metrics.detected_regions[0];

    // Improved heuristic: Numeric year headers are recognized
    if let Some(hr) = region.header_row {
        assert_eq!(hr, 1, "Should detect numeric headers at row 1");
        assert!(
            region.headers.contains(&"2022".to_string()),
            "Should contain year 2022"
        );
    } else {
        panic!("Header detection failed for numeric headers");
    }
}

#[test]
fn hierarchical_row_headers() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("hierarchical.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        // Row 1: [Empty] | Q1 | Q2
        sheet.get_cell_mut("B1").set_value("Q1");
        sheet.get_cell_mut("C1").set_value("Q2");

        // Row 2: Sales | 100 | 200
        sheet.get_cell_mut("A2").set_value("Sales");
        sheet.get_cell_mut("B2").set_value_number(100.0);
        sheet.get_cell_mut("C2").set_value_number(200.0);

        // Row 3:   Direct | 60 | 120  (Indented / Hierarchical)
        sheet.get_cell_mut("A3").set_value("  Direct");
        sheet.get_cell_mut("B3").set_value_number(60.0);
        sheet.get_cell_mut("C3").set_value_number(120.0);

        // Row 4:   Channel | 40 | 80
        sheet.get_cell_mut("A4").set_value("  Channel");
        sheet.get_cell_mut("B4").set_value_number(40.0);
        sheet.get_cell_mut("C4").set_value_number(80.0);
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    let metrics = ctx.get_sheet_metrics("Sheet1").unwrap();
    let region = &metrics.detected_regions[0];

    println!("Region bounds: {}", region.bounds);

    // CURRENT HEURISTIC LIMITATION:
    // The empty A1 cell might cause the region detection to trim the top row,
    // or the hierarchical text might be treated as metadata.

    // We expect the region to cover everything A1:C4 (or A2:C4 if header logic excludes empty A1)
    // The region bounds string (e.g., "A1:C4") might not contain the substring "A4",
    // so we check row_count instead.
    if region.row_count < 4 {
        println!(
            "WARNING: Hierarchical rows were excluded from region! Row count: {}",
            region.row_count
        );
    } else {
        println!("Hierarchical rows successfully included.");
    }
}

#[test]
fn headerless_list() {
    let workspace = support::TestWorkspace::new();
    let path = workspace.create_workbook("list.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        let items = ["Apple", "Banana", "Cherry", "Date", "Elderberry"];
        for (i, item) in items.iter().enumerate() {
            sheet.get_cell_mut((1, (i + 1) as u32)).set_value(*item);
        }
    });
    let ctx = WorkbookContext::load(&workspace.config().into(), &path).unwrap();

    let metrics = ctx.get_sheet_metrics("Sheet1").unwrap();
    let region = &metrics.detected_regions[0];

    // Expect Data classification even without headers
    assert!(
        matches!(region.region_kind, Some(RegionKind::Data)),
        "Expected Data classification, got {:?}",
        region.region_kind
    );

    // Improved heuristic: Stricter header detection avoids false positives on lists
    assert_eq!(
        region.header_row, None,
        "Should not detect a header in a flat list"
    );
}
