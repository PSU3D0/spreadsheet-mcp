use anyhow::Result;
use spreadsheet_mcp::model::{SheetPageFormat, TableOutputFormat};
use spreadsheet_mcp::tools::{
    ListWorkbooksParams, RangeValuesParams, ReadTableParams, SheetPageParams, list_workbooks,
    range_values, read_table, sheet_page,
};
use umya_spreadsheet::Spreadsheet;

mod support;

#[tokio::test(flavor = "current_thread")]
async fn read_table_defaults_to_csv() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("defaults.xlsx", build_simple_workbook);
    let state = workspace.app_state();

    let workbooks = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    let workbook_id = workbooks.workbooks[0].workbook_id.clone();

    let table = read_table(
        state,
        ReadTableParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: Some("Sheet1".into()),
            header_row: Some(1),
            limit: Some(10),
            ..Default::default()
        },
    )
    .await?;

    let csv = table.csv.expect("csv output expected by default");
    assert!(csv.lines().next().unwrap_or_default().contains("Name,Value,Flag"));
    assert!(table.rows.is_empty());
    assert!(table.values.is_none());
    assert!(table.headers.is_empty());
    assert_eq!(table.has_more, None);
    assert_eq!(table.next_offset, None);
    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn range_values_defaults_to_values() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("defaults.xlsx", build_simple_workbook);
    let state = workspace.app_state();

    let workbooks = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    let workbook_id = workbooks.workbooks[0].workbook_id.clone();

    let ranges = range_values(
        state,
        RangeValuesParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: "Sheet1".into(),
            ranges: vec!["A1:C2".into()],
            include_headers: Some(false),
            format: None,
            page_size: None,
        },
    )
    .await?;

    let entry = &ranges.values[0];
    assert!(entry.values.is_some());
    assert!(entry.rows.is_none());
    assert!(entry.csv.is_none());
    assert_eq!(entry.truncated, None);
    assert_eq!(entry.next_start_row, None);
    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn sheet_page_defaults_to_compact() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("defaults.xlsx", build_simple_workbook);
    let state = workspace.app_state();

    let workbooks = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    let workbook_id = workbooks.workbooks[0].workbook_id.clone();

    let page = sheet_page(
        state,
        SheetPageParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: "Sheet1".into(),
            start_row: 1,
            page_size: 20,
            include_formulas: true,
            include_styles: true,
            format: None,
            ..Default::default()
        },
    )
    .await?;

    assert!(matches!(page.format, SheetPageFormat::Compact));
    assert!(page.compact.is_some());
    assert!(page.rows.is_empty());
    assert_eq!(page.has_more, None);
    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn read_table_truncates_with_max_cells() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("truncate.xlsx", build_simple_workbook);
    let config = workspace.config_with(|config| {
        config.max_cells = Some(6);
        config.max_payload_bytes = None;
    });
    let state = support::app_state_with_config(config);

    let workbooks = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    let workbook_id = workbooks.workbooks[0].workbook_id.clone();

    let table = read_table(
        state,
        ReadTableParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: Some("Sheet1".into()),
            header_row: Some(1),
            limit: Some(50),
            format: Some(TableOutputFormat::Json),
            ..Default::default()
        },
    )
    .await?;

    assert_eq!(table.rows.len(), 2);
    assert!(table.total_rows as usize > table.rows.len());
    assert_eq!(table.has_more, Some(true));
    assert_eq!(table.next_offset, Some(2));
    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn range_values_truncates_with_max_cells() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("truncate.xlsx", build_simple_workbook);
    let config = workspace.config_with(|config| {
        config.max_cells = Some(6);
        config.max_payload_bytes = None;
    });
    let state = support::app_state_with_config(config);

    let workbooks = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    let workbook_id = workbooks.workbooks[0].workbook_id.clone();

    let ranges = range_values(
        state,
        RangeValuesParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: "Sheet1".into(),
            ranges: vec!["A1:C6".into()],
            include_headers: Some(false),
            format: Some(TableOutputFormat::Values),
            page_size: None,
        },
    )
    .await?;

    let entry = &ranges.values[0];
    let values = entry.values.as_ref().expect("values output expected");
    assert_eq!(values.len(), 2);
    assert_eq!(entry.truncated, Some(true));
    assert_eq!(entry.next_start_row, Some(3));
    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn sheet_page_truncates_with_max_cells() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    let _path = workspace.create_workbook("truncate.xlsx", build_simple_workbook);
    let config = workspace.config_with(|config| {
        config.max_cells = Some(6);
        config.max_payload_bytes = None;
    });
    let state = support::app_state_with_config(config);

    let workbooks = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    let workbook_id = workbooks.workbooks[0].workbook_id.clone();

    let page = sheet_page(
        state,
        SheetPageParams {
            workbook_or_fork_id: workbook_id,
            sheet_name: "Sheet1".into(),
            start_row: 1,
            page_size: 50,
            include_formulas: false,
            include_styles: false,
            format: Some(SheetPageFormat::Full),
            ..Default::default()
        },
    )
    .await?;

    assert_eq!(page.rows.len(), 2);
    assert_eq!(page.has_more, Some(true));
    assert_eq!(page.next_start_row, Some(3));
    Ok(())
}

fn build_simple_workbook(book: &mut Spreadsheet) {
    let sheet = book.get_sheet_by_name_mut("Sheet1").expect("Sheet1");
    sheet.get_cell_mut("A1").set_value("Name");
    sheet.get_cell_mut("B1").set_value("Value");
    sheet.get_cell_mut("C1").set_value("Flag");

    for i in 0..5 {
        let row = i + 2;
        sheet
            .get_cell_mut(format!("A{row}").as_str())
            .set_value(format!("Item {}", i + 1));
        sheet
            .get_cell_mut(format!("B{row}").as_str())
            .set_value_number((i + 1) as f64);
        sheet
            .get_cell_mut(format!("C{row}").as_str())
            .set_value("Y");
    }
}
