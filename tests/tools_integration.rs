use std::sync::Arc;

use anyhow::Result;
use spreadsheet_mcp::model::{TraceDirection, WorkbookId};
use spreadsheet_mcp::state::AppState;
use spreadsheet_mcp::tools::{
    DescribeWorkbookParams, FindFormulaParams, FormulaTraceParams, ListSheetsParams,
    ListWorkbooksParams, ManifestStubParams, NamedRangesParams, ScanVolatilesParams,
    SheetFormulaMapParams, SheetOverviewParams, SheetPageParams, SheetStatisticsParams,
    SheetStylesParams, describe_workbook, find_formula, formula_trace, get_manifest_stub,
    list_sheets, list_workbooks, named_ranges, scan_volatiles, sheet_formula_map, sheet_overview,
    sheet_page, sheet_statistics, sheet_styles,
};
use umya_spreadsheet::{NumberingFormat, Spreadsheet};

mod support;

#[tokio::test(flavor = "current_thread")]
async fn tool_suite_exercises_feature_rich_workbook() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("analysis.xlsx", build_featured_workbook);
    let state = workspace.app_state();

    let list_response = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    assert_eq!(list_response.workbooks.len(), 1);
    let descriptor = &list_response.workbooks[0];
    assert!(descriptor.short_id.len() < descriptor.workbook_id.as_str().len());
    let workbook_id = descriptor.workbook_id.clone();

    describe_and_overview_suite(state.clone(), workbook_id.clone()).await?;
    paging_and_stats_suite(state.clone(), workbook_id.clone()).await?;
    formula_and_dependency_suite(state.clone(), workbook_id.clone()).await?;
    naming_and_styles_suite(state.clone(), workbook_id.clone()).await?;
    manifest_suite(state, workbook_id).await?;

    Ok(())
}

async fn describe_and_overview_suite(state: Arc<AppState>, workbook_id: WorkbookId) -> Result<()> {
    let description = describe_workbook(
        state.clone(),
        DescribeWorkbookParams {
            workbook_or_fork_id: workbook_id.clone(),
        },
    )
    .await?;
    assert_eq!(description.sheet_count, 2);
    assert!(description.bytes > 0);
    assert!(description.caps.supports_formula_graph);

    let sheets = list_sheets(
        state.clone(),
        ListSheetsParams {
            workbook_or_fork_id: workbook_id.clone(),
        },
    )
    .await?;
    assert_eq!(sheets.sheets.len(), 2);
    let data_sheet = sheets
        .sheets
        .iter()
        .find(|s| s.name == "Data")
        .expect("data sheet present");
    assert!(data_sheet.formula_cells > 0);

    let overview = sheet_overview(
        state,
        SheetOverviewParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: "Data".to_string(),
        },
    )
    .await?;
    assert!(!overview.narrative.is_empty());
    assert!(!overview.regions.is_empty());
    assert!(!overview.key_ranges.is_empty());
    assert!(overview.formula_ratio > 0.0);

    Ok(())
}

async fn paging_and_stats_suite(state: Arc<AppState>, workbook_id: WorkbookId) -> Result<()> {
    let page = sheet_page(
        state.clone(),
        SheetPageParams {
            workbook_or_fork_id: workbook_id.clone(),
            sheet_name: "Data".to_string(),
            start_row: 2,
            page_size: 5,
            columns: Some(vec!["A".into(), "D".into(), "E".into(), "G".into()]),
            columns_by_header: None,
            include_formulas: true,
            include_styles: true,
            include_header: true,
            format: None,
        },
    )
    .await?;
    assert_eq!(page.rows.len(), 5);
    assert!(page.has_more);
    assert!(page.next_start_row.unwrap() > 5);
    assert!(page.header_row.is_some());
    let first_row = &page.rows[0];
    assert_eq!(first_row.cells.len(), 4);
    assert!(first_row.cells.iter().any(|cell| cell.value.is_some()));
    let page_has_formula = page
        .rows
        .iter()
        .flat_map(|row| row.cells.iter())
        .any(|cell| cell.formula.is_some());
    assert!(page_has_formula);

    let stats = sheet_statistics(
        state.clone(),
        SheetStatisticsParams {
            workbook_or_fork_id: workbook_id.clone(),
            sheet_name: "Data".to_string(),
            sample_rows: None,
        },
    )
    .await?;
    assert!(stats.density > 0.1);
    assert!(
        stats
            .numeric_columns
            .iter()
            .any(|col| col.column == "B" && col.min.unwrap() > 0.0)
    );
    assert!(
        stats
            .text_columns
            .iter()
            .any(|col| col.column == "A" && !col.samples.is_empty())
    );
    assert!(
        stats
            .null_counts
            .get("F")
            .map(|count| *count > 0)
            .unwrap_or(false)
    );

    let formula_map = sheet_formula_map(
        state.clone(),
        SheetFormulaMapParams {
            workbook_or_fork_id: workbook_id.clone(),
            sheet_name: "Data".to_string(),
            range: Some("D2:D21".to_string()),
            expand: false,
            limit: None,
            sort_by: None,
        },
    )
    .await?;
    assert!(!formula_map.groups.is_empty());
    let primary_group = &formula_map.groups[0];
    assert!(primary_group.formula.contains("*"));
    assert!(!primary_group.addresses.is_empty());

    let limited_map = sheet_formula_map(
        state.clone(),
        SheetFormulaMapParams {
            workbook_or_fork_id: workbook_id.clone(),
            sheet_name: "Data".to_string(),
            range: None,
            expand: false,
            limit: Some(2),
            sort_by: Some(spreadsheet_mcp::tools::FormulaSortBy::Complexity),
        },
    )
    .await?;
    assert!(
        limited_map.groups.len() <= 2,
        "limit should cap groups to 2"
    );
    if limited_map.groups.len() == 2 {
        assert!(
            limited_map.groups[0].formula.len() >= limited_map.groups[1].formula.len(),
            "complexity sort should order by formula length descending"
        );
    }

    Ok(())
}

async fn formula_and_dependency_suite(state: Arc<AppState>, workbook_id: WorkbookId) -> Result<()> {
    let trace = formula_trace(
        state.clone(),
        FormulaTraceParams {
            workbook_or_fork_id: workbook_id.clone(),
            sheet_name: "Data".to_string(),
            cell_address: "E21".to_string(),
            direction: TraceDirection::Precedents,
            depth: Some(2),
            limit: None,
            page_size: Some(12),
            cursor: None,
        },
    )
    .await?;
    assert_eq!(trace.origin, "E21");
    if !matches!(trace.direction, TraceDirection::Precedents) {
        panic!("expected precedents trace");
    }
    assert!(
        !trace.layers.is_empty(),
        "E21 formula SUM($D$2:D21) should have precedents but got empty layers"
    );
    let layer = trace.layers.first().unwrap();
    assert!(
        layer.summary.total_nodes >= 1,
        "expected at least 1 precedent node"
    );

    let matches = find_formula(
        state.clone(),
        FindFormulaParams {
            workbook_or_fork_id: workbook_id.clone(),
            query: "SUM(".to_string(),
            sheet_name: None,
            case_sensitive: false,
            include_context: true,
            limit: 50,
            offset: 0,
        },
    )
    .await?;
    assert!(!matches.matches.is_empty());
    assert!(
        matches
            .matches
            .iter()
            .any(|m| m.address.starts_with("D") || m.address.starts_with("E"))
    );
    assert!(!matches.matches[0].context.is_empty());

    let volatiles = scan_volatiles(
        state,
        ScanVolatilesParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: Some("Data".to_string()),
        },
    )
    .await?;
    assert!(volatiles.items.len() <= 2);
    assert!(!volatiles.truncated);

    Ok(())
}

async fn naming_and_styles_suite(state: Arc<AppState>, workbook_id: WorkbookId) -> Result<()> {
    let _names = named_ranges(
        state.clone(),
        NamedRangesParams {
            workbook_or_fork_id: workbook_id.clone(),
            sheet_name: None,
            name_prefix: Some("Sales".to_string()),
        },
    )
    .await?;

    let styles = sheet_styles(
        state,
        SheetStylesParams {
            workbook_or_fork_id: workbook_id.clone(),
            sheet_name: "Data".to_string(),
            scope: None,
            granularity: None,
            max_items: None,
        },
    )
    .await?;
    assert!(!styles.styles.is_empty());
    assert!(
        styles
            .styles
            .iter()
            .any(|style| style.tags.iter().any(|tag| tag == "header"))
    );

    Ok(())
}

async fn manifest_suite(state: Arc<AppState>, workbook_id: WorkbookId) -> Result<()> {
    let manifest = get_manifest_stub(
        state,
        ManifestStubParams {
            workbook_or_fork_id: workbook_id,
            sheet_filter: None,
        },
    )
    .await?;
    assert_eq!(manifest.sheets.len(), 2);
    assert!(
        manifest
            .sheets
            .iter()
            .any(|sheet| !sheet.candidate_expectations.is_empty())
    );

    Ok(())
}

fn build_featured_workbook(book: &mut Spreadsheet) {
    let data = book.get_sheet_by_name_mut("Sheet1").unwrap();
    data.set_name("Data");
    let headers = [
        "Item",
        "Qty",
        "Price",
        "Total",
        "RunningTotal",
        "Notes",
        "Volatile",
    ];
    for (idx, header) in headers.iter().enumerate() {
        let col = (idx as u32) + 1;
        data.get_cell_mut((col, 1)).set_value(header.to_string());
        let style = data.get_style_mut((col, 1));
        style.get_font_mut().set_bold(true);
        if matches!(*header, "Price" | "Total" | "RunningTotal") {
            style
                .get_number_format_mut()
                .set_format_code(NumberingFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
        }
    }

    let mut cumulative = 0.0f64;
    let mut qty_sum = 0.0f64;
    let mut last_total = 0.0f64;
    for row in 2..=21 {
        let qty = (row * 2) as f64;
        let price = 10.0 + row as f64;
        let total = qty * price;
        cumulative += total;
        qty_sum += qty;
        last_total = total;

        if row % 5 == 0 {
            data.get_cell_mut((1, row))
                .set_value("ItemRepeat".to_string());
        } else {
            data.get_cell_mut((1, row)).set_value(format!("Item{row}"));
        }
        data.get_cell_mut((2, row)).set_value_number(qty);
        data.get_cell_mut((3, row)).set_value_number(price);
        data.get_cell_mut((4, row))
            .set_formula(format!("B{row}*C{row}"))
            .set_formula_result_default(format!("{total:.2}"));
        data.get_cell_mut((5, row))
            .set_formula(format!("SUM($D$2:D{row})"))
            .set_formula_result_default(format!("{cumulative:.2}"));
        if row % 3 == 0 {
            data.get_cell_mut((6, row))
                .set_value(format!("Cycle {row}"));
        }
        data.get_style_mut((4, row))
            .get_number_format_mut()
            .set_format_code(NumberingFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
        data.get_style_mut((5, row))
            .get_number_format_mut()
            .set_format_code(NumberingFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
    }

    data.get_cell_mut((7, 2)).set_formula("NOW()");
    data.get_cell_mut((7, 3)).set_formula("RAND()");

    {
        let calc = book.new_sheet("Calc").expect("calc sheet");
        calc.get_cell_mut("A1")
            .set_formula("SUM(Data!D2:D21)")
            .set_formula_result_default(format!("{cumulative:.2}"));
        calc.get_cell_mut("A2")
            .set_formula("A1*2")
            .set_formula_result_default(format!("{:.2}", cumulative * 2.0));
        calc.get_cell_mut("B2")
            .set_formula("A2-Data!D21")
            .set_formula_result_default(format!("{:.2}", cumulative * 2.0 - last_total));
        calc.get_cell_mut("C3")
            .set_formula("SUM(Data!B2:B21)")
            .set_formula_result_default(format!("{qty_sum:.2}"));
        calc.get_cell_mut((4, 4)).set_formula("B2+A2");
        calc.get_cell_mut((2, 5)).set_value_number(42.0);
    }

    let data_sheet = book.get_sheet_by_name_mut("Data").unwrap();
    data_sheet
        .add_defined_name("SalesTotal", "Data!$D$2:$D$21")
        .expect("global defined name");
    data_sheet
        .add_defined_name("SalesLatest", "Data!$E$21")
        .expect("sheet defined name");
}

#[tokio::test(flavor = "current_thread")]
async fn find_formula_defaults_and_paging() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("find_formula_paging.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.set_name("Sheet1");
        for row in 1..=5 {
            sheet
                .get_cell_mut((2, row))
                .set_formula(format!("SUM(A{row}:A{row})"));
        }
    });

    let state = workspace.app_state();
    let list_response = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    let workbook_id = list_response.workbooks[0].workbook_id.clone();

    let first_page = find_formula(
        state.clone(),
        FindFormulaParams {
            workbook_or_fork_id: workbook_id.clone(),
            query: "SUM(".to_string(),
            sheet_name: Some("Sheet1".to_string()),
            case_sensitive: false,
            include_context: false,
            limit: 2,
            offset: 0,
        },
    )
    .await?;

    assert_eq!(first_page.matches.len(), 2);
    assert!(first_page.matches.iter().all(|m| m.context.is_empty()));
    assert!(first_page.truncated);
    assert_eq!(first_page.next_offset, Some(2));

    let second_page = find_formula(
        state.clone(),
        FindFormulaParams {
            workbook_or_fork_id: workbook_id.clone(),
            query: "SUM(".to_string(),
            sheet_name: Some("Sheet1".to_string()),
            case_sensitive: false,
            include_context: false,
            limit: 2,
            offset: first_page.next_offset.unwrap(),
        },
    )
    .await?;

    assert!(!second_page.matches.is_empty());
    assert_ne!(
        first_page.matches[0].address,
        second_page.matches[0].address
    );

    let with_context = find_formula(
        state,
        FindFormulaParams {
            workbook_or_fork_id: workbook_id,
            query: "SUM(".to_string(),
            sheet_name: Some("Sheet1".to_string()),
            case_sensitive: false,
            include_context: true,
            limit: 50,
            offset: 0,
        },
    )
    .await?;

    assert!(!with_context.matches.is_empty());
    assert!(!with_context.matches[0].context.is_empty());
    assert!(!with_context.truncated);
    assert!(with_context.next_offset.is_none());

    Ok(())
}
