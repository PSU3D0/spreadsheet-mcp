#![cfg(feature = "recalc")]

use anyhow::Result;
use spreadsheet_mcp::model::WorkbookId;
use spreadsheet_mcp::tools::fork::{
    ApplyStagedChangeParams, CreateForkParams, StructureBatchParams, StructureOp, apply_staged_change,
    create_fork, structure_batch,
};
use spreadsheet_mcp::tools::{ListWorkbooksParams, list_workbooks};

mod support;

fn recalc_state(
    workspace: &support::TestWorkspace,
) -> std::sync::Arc<spreadsheet_mcp::state::AppState> {
    let config = workspace.config_with(|cfg| {
        cfg.recalc_enabled = true;
    });
    support::app_state_with_config(config)
}

#[tokio::test(flavor = "current_thread")]
async fn structure_batch_insert_rows_moves_cells() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("structure_rows.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_value("keep");
        sheet.get_cell_mut("A2").set_value("move");
    });

    let state = recalc_state(&workspace);
    let list = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    let workbook_id = list.workbooks[0].workbook_id.clone();
    let fork = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    structure_batch(
        state.clone(),
        StructureBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StructureOp::InsertRows {
                sheet_name: "Sheet1".to_string(),
                at_row: 2,
                count: 1,
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_wb = state.open_workbook(&WorkbookId(fork.fork_id.clone())).await?;
    let values = fork_wb.with_sheet("Sheet1", |sheet| {
        let a1 = sheet.get_cell("A1").unwrap().get_value().to_string();
        let a2 = sheet
            .get_cell("A2")
            .map(|c| c.get_value().to_string())
            .unwrap_or_default();
        let a3 = sheet
            .get_cell("A3")
            .map(|c| c.get_value().to_string())
            .unwrap_or_default();
        (a1, a2, a3)
    })?;

    assert_eq!(values.0, "keep");
    assert_eq!(values.1, "");
    assert_eq!(values.2, "move");

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn structure_batch_rename_sheet_rewrites_formula_refs() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("structure_rename.xlsx", |book| {
        let inputs = book.get_sheet_mut(&0).unwrap();
        inputs.set_name("Inputs");
        inputs.get_cell_mut("A1").set_value_number(3);

        book.new_sheet("Calc").unwrap();
        let calc = book.get_sheet_by_name_mut("Calc").unwrap();
        calc.get_cell_mut("A1").set_formula("Inputs!A1".to_string());
    });

    let state = recalc_state(&workspace);
    let list = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    let workbook_id = list.workbooks[0].workbook_id.clone();
    let fork = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    structure_batch(
        state.clone(),
        StructureBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StructureOp::RenameSheet {
                old_name: "Inputs".to_string(),
                new_name: "Data".to_string(),
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_wb = state.open_workbook(&WorkbookId(fork.fork_id.clone())).await?;
    let formula = fork_wb.with_sheet("Calc", |sheet| {
        sheet.get_cell("A1").unwrap().get_formula().to_string()
    })?;
    assert_eq!(formula, "Data!A1");

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn structure_batch_preview_stages_and_apply() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("structure_preview.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("B1").set_value("move");
    });

    let state = recalc_state(&workspace);
    let list = list_workbooks(
        state.clone(),
        ListWorkbooksParams {
            slug_prefix: None,
            folder: None,
            path_glob: None,
        },
    )
    .await?;
    let workbook_id = list.workbooks[0].workbook_id.clone();
    let fork = create_fork(state.clone(), CreateForkParams { workbook_id }).await?;

    let preview = structure_batch(
        state.clone(),
        StructureBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StructureOp::InsertCols {
                sheet_name: "Sheet1".to_string(),
                at_col: "B".to_string(),
                count: 1,
            }],
            mode: Some("preview".to_string()),
            label: Some("insert col".to_string()),
        },
    )
    .await?;
    let change_id = preview.change_id.clone().expect("change_id");

    // Preview should not mutate the fork.
    let fork_wb = state.open_workbook(&WorkbookId(fork.fork_id.clone())).await?;
    let b1 = fork_wb.with_sheet("Sheet1", |sheet| {
        sheet.get_cell("B1").unwrap().get_value().to_string()
    })?;
    assert_eq!(b1, "move");

    apply_staged_change(
        state.clone(),
        ApplyStagedChangeParams {
            fork_id: fork.fork_id.clone(),
            change_id,
        },
    )
    .await?;

    let fork_wb = state.open_workbook(&WorkbookId(fork.fork_id.clone())).await?;
    let moved = fork_wb.with_sheet("Sheet1", |sheet| {
        (
            sheet
                .get_cell("B1")
                .map(|c| c.get_value().to_string())
                .unwrap_or_default(),
            sheet.get_cell("C1").unwrap().get_value().to_string(),
        )
    })?;
    assert_eq!(moved.0, "");
    assert_eq!(moved.1, "move");

    Ok(())
}
