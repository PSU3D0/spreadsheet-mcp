//! Docker E2E tests for structure_batch (Phase 4).

use anyhow::Result;
use serde_json::json;

use crate::support::mcp::{McpTestClient, call_tool, cell_is_error, cell_value_f64, extract_json};

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

