//! Docker E2E tests for structure_batch (Phase 4).

use anyhow::Result;
use serde_json::json;

use crate::support::mcp::{
    McpTestClient, call_tool, cell_is_error, cell_value, cell_value_f64, extract_json,
};

#[tokio::test]
async fn test_structure_batch_insert_rows_updates_cross_sheet_formulas_in_docker() -> Result<()> {
    let test = McpTestClient::new();
    test.workspace()
        .create_workbook("structure_rows.xlsx", |book| {
            let inputs = book.get_sheet_mut(&0).unwrap();
            inputs.set_name("Inputs");
            inputs.get_cell_mut("A1").set_value_number(1);
            inputs.get_cell_mut("A2").set_value_number(2);

            book.new_sheet("Calc").unwrap();
            let calc = book.get_sheet_by_name_mut("Calc").unwrap();
            calc.get_cell_mut("A1")
                .set_formula("SUM(Inputs!A1:A2)".to_string());
        });

    let client = test.connect().await?;
    let workbooks = extract_json(
        &client
            .call_tool(call_tool("list_workbooks", json!({})))
            .await?,
    )?;
    let workbook_id = workbooks["workbooks"][0]["workbook_id"].as_str().unwrap();

    let fork = extract_json(
        &client
            .call_tool(call_tool(
                "create_fork",
                json!({ "workbook_id": workbook_id }),
            ))
            .await?,
    )?;
    let fork_id = fork["fork_id"].as_str().unwrap();

    let _ = client
        .call_tool(call_tool(
            "structure_batch",
            json!({
              "fork_id": fork_id,
              "mode": "apply",
              "ops": [{
                "kind": "insert_rows",
                "sheet_name": "Inputs",
                "at_row": 2,
                "count": 1
              }]
            }),
        ))
        .await?;

    let _ = client
        .call_tool(call_tool("recalculate", json!({ "fork_id": fork_id })))
        .await?;

    let calc_page = extract_json(
        &client
            .call_tool(call_tool(
                "sheet_page",
                json!({
                    "workbook_id": fork_id,
                    "sheet_name": "Calc",
                    "start_row": 1,
                    "page_size": 1,
                    "columns": ["A"],
                    "include_formulas": false
                }),
            ))
            .await?,
    )?;
    assert!(!cell_is_error(&calc_page, 0, 0));
    assert_eq!(cell_value_f64(&calc_page, 0, 0), Some(3.0));

    client.cancel().await?;
    Ok(())
}

#[tokio::test]
async fn test_structure_batch_rename_sheet_preserves_formulas_in_docker() -> Result<()> {
    let test = McpTestClient::new();
    test.workspace()
        .create_workbook("structure_rename.xlsx", |book| {
            let inputs = book.get_sheet_mut(&0).unwrap();
            inputs.set_name("Inputs");
            inputs.get_cell_mut("A1").set_value_number(3);

            book.new_sheet("Calc").unwrap();
            let calc = book.get_sheet_by_name_mut("Calc").unwrap();
            calc.get_cell_mut("A1").set_formula("Inputs!A1".to_string());
        });

    let client = test.connect().await?;
    let workbooks = extract_json(
        &client
            .call_tool(call_tool("list_workbooks", json!({})))
            .await?,
    )?;
    let workbook_id = workbooks["workbooks"][0]["workbook_id"].as_str().unwrap();

    let fork = extract_json(
        &client
            .call_tool(call_tool(
                "create_fork",
                json!({ "workbook_id": workbook_id }),
            ))
            .await?,
    )?;
    let fork_id = fork["fork_id"].as_str().unwrap();

    let _ = client
        .call_tool(call_tool(
            "structure_batch",
            json!({
              "fork_id": fork_id,
              "mode": "apply",
              "ops": [{
                "kind": "rename_sheet",
                "old_name": "Inputs",
                "new_name": "Data"
              }]
            }),
        ))
        .await?;

    let _ = client
        .call_tool(call_tool("recalculate", json!({ "fork_id": fork_id })))
        .await?;

    let calc_page = extract_json(
        &client
            .call_tool(call_tool(
                "sheet_page",
                json!({
                    "workbook_id": fork_id,
                    "sheet_name": "Calc",
                    "start_row": 1,
                    "page_size": 1,
                    "columns": ["A"],
                    "include_formulas": false
                }),
            ))
            .await?,
    )?;
    assert!(!cell_is_error(&calc_page, 0, 0));
    assert_eq!(cell_value_f64(&calc_page, 0, 0), Some(3.0));

    client.cancel().await?;
    Ok(())
}

#[tokio::test]
async fn test_structure_batch_insert_cols_updates_cross_sheet_formulas_in_docker() -> Result<()> {
    let test = McpTestClient::new();
    test.workspace()
        .create_workbook("structure_cols.xlsx", |book| {
            let inputs = book.get_sheet_mut(&0).unwrap();
            inputs.set_name("Inputs");
            inputs.get_cell_mut("A1").set_value_number(1);
            inputs.get_cell_mut("B1").set_value_number(2);

            book.new_sheet("Calc").unwrap();
            let calc = book.get_sheet_by_name_mut("Calc").unwrap();
            calc.get_cell_mut("A1")
                .set_formula("SUM(Inputs!A1:B1)".to_string());
        });

    let client = test.connect().await?;
    let workbooks = extract_json(
        &client
            .call_tool(call_tool("list_workbooks", json!({})))
            .await?,
    )?;
    let workbook_id = workbooks["workbooks"][0]["workbook_id"].as_str().unwrap();

    let fork = extract_json(
        &client
            .call_tool(call_tool(
                "create_fork",
                json!({ "workbook_id": workbook_id }),
            ))
            .await?,
    )?;
    let fork_id = fork["fork_id"].as_str().unwrap();

    let _ = client
        .call_tool(call_tool(
            "structure_batch",
            json!({
              "fork_id": fork_id,
              "mode": "apply",
              "ops": [{
                "kind": "insert_cols",
                "sheet_name": "Inputs",
                "at_col": "A",
                "count": 1
              }]
            }),
        ))
        .await?;

    let _ = client
        .call_tool(call_tool("recalculate", json!({ "fork_id": fork_id })))
        .await?;

    let calc_page = extract_json(
        &client
            .call_tool(call_tool(
                "sheet_page",
                json!({
                    "workbook_id": fork_id,
                    "sheet_name": "Calc",
                    "start_row": 1,
                    "page_size": 1,
                    "columns": ["A"],
                    "include_formulas": false
                }),
            ))
            .await?,
    )?;
    assert!(!cell_is_error(&calc_page, 0, 0));
    assert_eq!(cell_value_f64(&calc_page, 0, 0), Some(3.0));

    client.cancel().await?;
    Ok(())
}

#[tokio::test]
async fn test_structure_batch_delete_rows_preserves_formula_result_in_docker() -> Result<()> {
    let test = McpTestClient::new();
    test.workspace()
        .create_workbook("structure_delete_rows.xlsx", |book| {
            let inputs = book.get_sheet_mut(&0).unwrap();
            inputs.set_name("Inputs");
            inputs.get_cell_mut("A1").set_value_number(1);
            inputs.get_cell_mut("A2").set_value_number(2);
            inputs.get_cell_mut("A3").set_value_number(3);

            book.new_sheet("Calc").unwrap();
            let calc = book.get_sheet_by_name_mut("Calc").unwrap();
            calc.get_cell_mut("A1")
                .set_formula("SUM(Inputs!A1:A3)".to_string());
        });

    let client = test.connect().await?;
    let workbooks = extract_json(
        &client
            .call_tool(call_tool("list_workbooks", json!({})))
            .await?,
    )?;
    let workbook_id = workbooks["workbooks"][0]["workbook_id"].as_str().unwrap();

    let fork = extract_json(
        &client
            .call_tool(call_tool(
                "create_fork",
                json!({ "workbook_id": workbook_id }),
            ))
            .await?,
    )?;
    let fork_id = fork["fork_id"].as_str().unwrap();

    let _ = client
        .call_tool(call_tool(
            "structure_batch",
            json!({
              "fork_id": fork_id,
              "mode": "apply",
              "ops": [{
                "kind": "delete_rows",
                "sheet_name": "Inputs",
                "start_row": 2,
                "count": 1
              }]
            }),
        ))
        .await?;

    let _ = client
        .call_tool(call_tool("recalculate", json!({ "fork_id": fork_id })))
        .await?;

    let calc_page = extract_json(
        &client
            .call_tool(call_tool(
                "sheet_page",
                json!({
                    "workbook_id": fork_id,
                    "sheet_name": "Calc",
                    "start_row": 1,
                    "page_size": 1,
                    "columns": ["A"],
                    "include_formulas": false
                }),
            ))
            .await?,
    )?;
    assert!(!cell_is_error(&calc_page, 0, 0));
    // Should remain SUM of remaining values 1 and 3.
    assert_eq!(cell_value_f64(&calc_page, 0, 0), Some(4.0));

    client.cancel().await?;
    Ok(())
}

#[tokio::test]
async fn test_structure_batch_delete_cols_preserves_formula_result_in_docker() -> Result<()> {
    let test = McpTestClient::new();
    test.workspace()
        .create_workbook("structure_delete_cols.xlsx", |book| {
            let inputs = book.get_sheet_mut(&0).unwrap();
            inputs.set_name("Inputs");
            inputs.get_cell_mut("A1").set_value_number(1);
            inputs.get_cell_mut("B1").set_value_number(2);
            inputs.get_cell_mut("C1").set_value_number(3);

            book.new_sheet("Calc").unwrap();
            let calc = book.get_sheet_by_name_mut("Calc").unwrap();
            calc.get_cell_mut("A1")
                .set_formula("SUM(Inputs!A1:C1)".to_string());
        });

    let client = test.connect().await?;
    let workbooks = extract_json(
        &client
            .call_tool(call_tool("list_workbooks", json!({})))
            .await?,
    )?;
    let workbook_id = workbooks["workbooks"][0]["workbook_id"].as_str().unwrap();

    let fork = extract_json(
        &client
            .call_tool(call_tool(
                "create_fork",
                json!({ "workbook_id": workbook_id }),
            ))
            .await?,
    )?;
    let fork_id = fork["fork_id"].as_str().unwrap();

    let _ = client
        .call_tool(call_tool(
            "structure_batch",
            json!({
              "fork_id": fork_id,
              "mode": "apply",
              "ops": [{
                "kind": "delete_cols",
                "sheet_name": "Inputs",
                "start_col": "B",
                "count": 1
              }]
            }),
        ))
        .await?;

    let _ = client
        .call_tool(call_tool("recalculate", json!({ "fork_id": fork_id })))
        .await?;

    let calc_page = extract_json(
        &client
            .call_tool(call_tool(
                "sheet_page",
                json!({
                    "workbook_id": fork_id,
                    "sheet_name": "Calc",
                    "start_row": 1,
                    "page_size": 1,
                    "columns": ["A"],
                    "include_formulas": false
                }),
            ))
            .await?,
    )?;
    assert!(!cell_is_error(&calc_page, 0, 0));
    // Should remain SUM of remaining values 1 and 3.
    assert_eq!(cell_value_f64(&calc_page, 0, 0), Some(4.0));

    client.cancel().await?;
    Ok(())
}

#[tokio::test]
async fn test_structure_batch_rename_quoted_sheet_preserves_formulas_in_docker() -> Result<()> {
    let test = McpTestClient::new();
    test.workspace()
        .create_workbook("structure_rename_quoted.xlsx", |book| {
            let inputs = book.get_sheet_mut(&0).unwrap();
            inputs.set_name("My Sheet");
            inputs.get_cell_mut("A1").set_value_number(3);

            book.new_sheet("Calc").unwrap();
            let calc = book.get_sheet_by_name_mut("Calc").unwrap();
            calc.get_cell_mut("A1")
                .set_formula("'My Sheet'!A1".to_string());
        });

    let client = test.connect().await?;
    let workbooks = extract_json(
        &client
            .call_tool(call_tool("list_workbooks", json!({})))
            .await?,
    )?;
    let workbook_id = workbooks["workbooks"][0]["workbook_id"].as_str().unwrap();

    let fork = extract_json(
        &client
            .call_tool(call_tool(
                "create_fork",
                json!({ "workbook_id": workbook_id }),
            ))
            .await?,
    )?;
    let fork_id = fork["fork_id"].as_str().unwrap();

    let _ = client
        .call_tool(call_tool(
            "structure_batch",
            json!({
              "fork_id": fork_id,
              "mode": "apply",
              "ops": [{
                "kind": "rename_sheet",
                "old_name": "My Sheet",
                "new_name": "Data"
              }]
            }),
        ))
        .await?;

    let _ = client
        .call_tool(call_tool("recalculate", json!({ "fork_id": fork_id })))
        .await?;

    let calc_page = extract_json(
        &client
            .call_tool(call_tool(
                "sheet_page",
                json!({
                    "workbook_id": fork_id,
                    "sheet_name": "Calc",
                    "start_row": 1,
                    "page_size": 1,
                    "columns": ["A"],
                    "include_formulas": false
                }),
            ))
            .await?,
    )?;
    assert!(!cell_is_error(&calc_page, 0, 0));
    assert_eq!(cell_value_f64(&calc_page, 0, 0), Some(3.0));

    client.cancel().await?;
    Ok(())
}

#[tokio::test]
async fn test_structure_batch_copy_range_shifts_formulas_in_docker() -> Result<()> {
    let test = McpTestClient::new();
    test.workspace()
        .create_workbook("structure_copy_range.xlsx", |book| {
            let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
            sheet.get_cell_mut("A1").set_value_number(1);
            sheet.get_cell_mut("B1").set_value_number(10);
            sheet.get_cell_mut("A2").set_value_number(2);
            sheet.get_cell_mut("B2").set_value_number(20);
            sheet.get_cell_mut("C1").set_formula("A1+B1".to_string());
        });

    let client = test.connect().await?;
    let workbooks = extract_json(
        &client
            .call_tool(call_tool("list_workbooks", json!({})))
            .await?,
    )?;
    let workbook_id = workbooks["workbooks"][0]["workbook_id"].as_str().unwrap();

    let fork = extract_json(
        &client
            .call_tool(call_tool(
                "create_fork",
                json!({ "workbook_id": workbook_id }),
            ))
            .await?,
    )?;
    let fork_id = fork["fork_id"].as_str().unwrap();

    let _ = client
        .call_tool(call_tool(
            "structure_batch",
            json!({
              "fork_id": fork_id,
              "mode": "apply",
              "ops": [{
                "kind": "copy_range",
                "sheet_name": "Sheet1",
                "src_range": "C1:C1",
                "dest_anchor": "D1",
                "include_styles": false,
                "include_formulas": true
              }]
            }),
        ))
        .await?;

    let _ = client
        .call_tool(call_tool("recalculate", json!({ "fork_id": fork_id })))
        .await?;

    let page = extract_json(
        &client
            .call_tool(call_tool(
                "sheet_page",
                json!({
                    "workbook_id": fork_id,
                    "sheet_name": "Sheet1",
                    "start_row": 1,
                    "page_size": 1,
                    "columns": ["C", "D"],
                    "include_formulas": false
                }),
            ))
            .await?,
    )?;

    assert!(!cell_is_error(&page, 0, 0));
    assert!(!cell_is_error(&page, 0, 1));
    assert_eq!(cell_value_f64(&page, 0, 0), Some(11.0));
    assert_eq!(cell_value_f64(&page, 0, 1), Some(21.0));

    client.cancel().await?;
    Ok(())
}

#[tokio::test]
async fn test_structure_batch_move_range_moves_and_clears_source_in_docker() -> Result<()> {
    let test = McpTestClient::new();
    test.workspace()
        .create_workbook("structure_move_range.xlsx", |book| {
            let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
            sheet.get_cell_mut("A1").set_value("x");
        });

    let client = test.connect().await?;
    let workbooks = extract_json(
        &client
            .call_tool(call_tool("list_workbooks", json!({})))
            .await?,
    )?;
    let workbook_id = workbooks["workbooks"][0]["workbook_id"].as_str().unwrap();

    let fork = extract_json(
        &client
            .call_tool(call_tool(
                "create_fork",
                json!({ "workbook_id": workbook_id }),
            ))
            .await?,
    )?;
    let fork_id = fork["fork_id"].as_str().unwrap();

    let _ = client
        .call_tool(call_tool(
            "structure_batch",
            json!({
              "fork_id": fork_id,
              "mode": "apply",
              "ops": [{
                "kind": "move_range",
                "sheet_name": "Sheet1",
                "src_range": "A1:A1",
                "dest_anchor": "C3",
                "include_styles": false,
                "include_formulas": false
              }]
            }),
        ))
        .await?;

    let page = extract_json(
        &client
            .call_tool(call_tool(
                "sheet_page",
                json!({
                    "workbook_id": fork_id,
                    "sheet_name": "Sheet1",
                    "start_row": 1,
                    "page_size": 3,
                    "columns": ["A", "C"],
                    "include_formulas": false
                }),
            ))
            .await?,
    )?;

    assert_eq!(cell_value(&page, 0, 0), None);
    assert_eq!(cell_value(&page, 2, 1).as_deref(), Some("x"));

    client.cancel().await?;
    Ok(())
}
