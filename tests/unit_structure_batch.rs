#![cfg(feature = "recalc")]

use anyhow::Result;
use spreadsheet_mcp::model::WorkbookId;
use spreadsheet_mcp::styles::descriptor_from_style;
use spreadsheet_mcp::tools::fork::{
    ApplyStagedChangeParams, CreateForkParams, StructureBatchParams, StructureOp,
    apply_staged_change, create_fork, structure_batch,
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

    let fork_wb = state
        .open_workbook(&WorkbookId(fork.fork_id.clone()))
        .await?;
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
async fn structure_batch_copy_range_shifts_formulas_and_copies_style() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("structure_copy_range.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_value_number(1);
        sheet.get_cell_mut("B1").set_value_number(10);
        sheet.get_cell_mut("A2").set_value_number(2);
        sheet.get_cell_mut("B2").set_value_number(20);

        sheet.get_cell_mut("C1").set_formula("A1+B1".to_string());
        sheet.get_style_mut("C1").get_font_mut().set_bold(true);
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
            ops: vec![StructureOp::CopyRange {
                sheet_name: "Sheet1".to_string(),
                src_range: "C1:C1".to_string(),
                dest_anchor: "D1".to_string(),
                include_styles: true,
                include_formulas: true,
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_wb = state
        .open_workbook(&WorkbookId(fork.fork_id.clone()))
        .await?;
    let (src_formula, dest_formula, dest_bold) = fork_wb.with_sheet("Sheet1", |sheet| {
        let src = sheet.get_cell("C1").expect("C1").get_formula().to_string();
        let dest = sheet.get_cell("D1").expect("D1").get_formula().to_string();
        let desc = descriptor_from_style(sheet.get_cell("D1").expect("D1").get_style());
        (src, dest, desc.font.and_then(|f| f.bold).unwrap_or(false))
    })?;

    assert_eq!(src_formula, "A1+B1");
    assert_eq!(dest_formula.replace(' ', ""), "B1+C1");
    assert!(dest_bold);

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn structure_batch_move_range_moves_and_clears_source() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("structure_move_range.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_value("x");
        sheet.get_style_mut("A1").get_font_mut().set_bold(true);
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
            ops: vec![StructureOp::MoveRange {
                sheet_name: "Sheet1".to_string(),
                src_range: "A1:A1".to_string(),
                dest_anchor: "C3".to_string(),
                include_styles: true,
                include_formulas: false,
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_wb = state
        .open_workbook(&WorkbookId(fork.fork_id.clone()))
        .await?;
    let (a1_val, c3_val, c3_bold) = fork_wb.with_sheet("Sheet1", |sheet| {
        let a1 = sheet
            .get_cell("A1")
            .map(|c| c.get_value().to_string())
            .unwrap_or_default();
        let c3 = sheet.get_cell("C3").expect("C3");
        let desc = descriptor_from_style(c3.get_style());
        (
            a1,
            c3.get_value().to_string(),
            desc.font.and_then(|f| f.bold).unwrap_or(false),
        )
    })?;

    assert_eq!(a1_val, "");
    assert_eq!(c3_val, "x");
    assert!(c3_bold);

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn structure_batch_copy_range_rejects_overlap() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("structure_copy_overlap.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_value("x");
        sheet.get_cell_mut("B2").set_value("y");
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

    let err = structure_batch(
        state.clone(),
        StructureBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StructureOp::CopyRange {
                sheet_name: "Sheet1".to_string(),
                src_range: "A1:B2".to_string(),
                dest_anchor: "B2".to_string(),
                include_styles: false,
                include_formulas: false,
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await
    .unwrap_err();

    assert!(err.to_string().contains("overlaps source"));
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
    let fork_wb = state
        .open_workbook(&WorkbookId(fork.fork_id.clone()))
        .await?;
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

    let fork_wb = state
        .open_workbook(&WorkbookId(fork.fork_id.clone()))
        .await?;
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

#[tokio::test(flavor = "current_thread")]
async fn structure_batch_preview_includes_change_count() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("structure_preview_count.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_value("x");
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
                at_col: "A".to_string(),
                count: 1,
            }],
            mode: Some("preview".to_string()),
            label: None,
        },
    )
    .await?;

    assert!(
        preview.summary.counts.contains_key("preview_change_items"),
        "preview should include preview_change_items"
    );

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn structure_batch_rename_sheet_handles_quoted_sheet_names() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("structure_rename_quoted.xlsx", |book| {
        let inputs = book.get_sheet_mut(&0).unwrap();
        inputs.set_name("My Sheet");
        inputs.get_cell_mut("A1").set_value_number(3);

        book.new_sheet("Calc").unwrap();
        let calc = book.get_sheet_by_name_mut("Calc").unwrap();
        calc.get_cell_mut("A1")
            .set_formula("'My Sheet'!A1".to_string());
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
                old_name: "My Sheet".to_string(),
                new_name: "Data".to_string(),
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_wb = state
        .open_workbook(&WorkbookId(fork.fork_id.clone()))
        .await?;
    let formula = fork_wb.with_sheet("Calc", |sheet| {
        sheet.get_cell("A1").unwrap().get_formula().to_string()
    })?;
    assert_eq!(formula, "Data!A1");

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn structure_batch_create_sheet_inserts_at_position() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("structure_create_sheet.xlsx", |_| {});

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
            ops: vec![StructureOp::CreateSheet {
                name: "First".to_string(),
                position: Some(0),
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_wb = state
        .open_workbook(&WorkbookId(fork.fork_id.clone()))
        .await?;
    let sheets = fork_wb.sheet_names();
    assert_eq!(sheets[0], "First");

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn structure_batch_delete_sheet_guard_prevents_last_sheet() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("structure_delete_last.xlsx", |_| {});

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

    let err = structure_batch(
        state.clone(),
        StructureBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StructureOp::DeleteSheet {
                name: "Sheet1".to_string(),
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await
    .unwrap_err();

    assert!(err.to_string().contains("last remaining sheet"));

    Ok(())
}
