#![cfg(feature = "recalc")]

use anyhow::Result;
use spreadsheet_mcp::model::{FillPatch, FontPatch, PatternFillPatch, StylePatch};
use spreadsheet_mcp::tools::fork::{
    ApplyStagedChangeParams, CreateForkParams, StyleBatchParams, StyleOp, StyleTarget,
    apply_staged_change, create_fork, style_batch,
};
use spreadsheet_mcp::tools::{ListWorkbooksParams, list_workbooks};
use umya_spreadsheet::{
    ConditionalFormatValues, ConditionalFormatting, ConditionalFormattingRule, Formula,
    PatternValues,
};

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
async fn style_batch_merge_set_clear_semantics() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("style.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_value("x");
        let style = sheet.get_style_mut("A1");
        style.get_font_mut().set_bold(true);
        style
            .get_fill_mut()
            .get_pattern_fill_mut()
            .set_pattern_type(umya_spreadsheet::PatternValues::Solid)
            .get_foreground_color_mut()
            .set_argb("FF0000FF");
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

    let fork = create_fork(
        state.clone(),
        CreateForkParams {
            workbook_or_fork_id: workbook_id,
        },
    )
    .await?;

    // Merge: remove bold (explicit false) and clear fill.
    let patch_merge = StylePatch {
        font: Some(Some(FontPatch {
            bold: Some(Some(false)),
            ..Default::default()
        })),
        fill: Some(None),
        ..Default::default()
    };

    style_batch(
        state.clone(),
        StyleBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StyleOp {
                sheet_name: "Sheet1".to_string(),
                target: StyleTarget::Range {
                    range: "A1:A1".to_string(),
                },
                patch: patch_merge,
                op_mode: Some("merge".to_string()),
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_wb = state
        .open_workbook(&spreadsheet_mcp::model::WorkbookId(fork.fork_id.clone()))
        .await?;
    let desc_a1 = fork_wb.with_sheet("Sheet1", |sheet| {
        spreadsheet_mcp::styles::descriptor_from_style(
            sheet.get_cell("A1").expect("A1 cell").get_style(),
        )
    })?;
    assert!(desc_a1.font.as_ref().and_then(|f| f.bold).is_none());
    assert!(desc_a1.fill.is_none());

    // Set: apply solid red fill only, wiping other direct formatting.
    let patch_set = StylePatch {
        fill: Some(Some(spreadsheet_mcp::model::FillPatch::Pattern(
            PatternFillPatch {
                pattern_type: Some(Some("solid".to_string())),
                foreground_color: Some(Some("FFFF0000".to_string())),
                ..Default::default()
            },
        ))),
        ..Default::default()
    };

    style_batch(
        state.clone(),
        StyleBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StyleOp {
                sheet_name: "Sheet1".to_string(),
                target: StyleTarget::Cells {
                    cells: vec!["A1".to_string()],
                },
                patch: patch_set,
                op_mode: Some("set".to_string()),
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_wb = state
        .open_workbook(&spreadsheet_mcp::model::WorkbookId(fork.fork_id.clone()))
        .await?;
    let desc_a1 = fork_wb.with_sheet("Sheet1", |sheet| {
        spreadsheet_mcp::styles::descriptor_from_style(
            sheet.get_cell("A1").expect("A1 cell").get_style(),
        )
    })?;
    assert!(desc_a1.font.as_ref().and_then(|f| f.bold).is_none());
    let fill = desc_a1.fill.as_ref().expect("fill");
    match fill {
        spreadsheet_mcp::model::FillDescriptor::Pattern(p) => {
            assert_eq!(p.foreground_color.as_deref(), Some("FFFF0000"));
        }
        _ => panic!("expected pattern fill"),
    }

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn style_batch_preview_stages_and_apply() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("preview.xlsx", |book| {
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

    let fork = create_fork(
        state.clone(),
        CreateForkParams {
            workbook_or_fork_id: workbook_id,
        },
    )
    .await?;

    let patch = StylePatch {
        font: Some(Some(FontPatch {
            bold: Some(Some(true)),
            ..Default::default()
        })),
        ..Default::default()
    };

    let preview = style_batch(
        state.clone(),
        StyleBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StyleOp {
                sheet_name: "Sheet1".to_string(),
                target: StyleTarget::Cells {
                    cells: vec!["A1".to_string()],
                },
                patch,
                op_mode: None,
            }],
            mode: Some("preview".to_string()),
            label: Some("bold headers".to_string()),
        },
    )
    .await?;
    let change_id = preview.change_id.clone().expect("change_id");

    // Preview should not mutate the fork.
    let fork_wb = state
        .open_workbook(&spreadsheet_mcp::model::WorkbookId(fork.fork_id.clone()))
        .await?;
    let desc_a1 = fork_wb.with_sheet("Sheet1", |sheet| {
        spreadsheet_mcp::styles::descriptor_from_style(
            sheet.get_cell("A1").expect("A1 cell").get_style(),
        )
    })?;
    assert!(desc_a1.font.is_none());

    apply_staged_change(
        state.clone(),
        ApplyStagedChangeParams {
            fork_id: fork.fork_id.clone(),
            change_id,
        },
    )
    .await?;

    let fork_wb = state
        .open_workbook(&spreadsheet_mcp::model::WorkbookId(fork.fork_id.clone()))
        .await?;
    let desc_a1 = fork_wb.with_sheet("Sheet1", |sheet| {
        spreadsheet_mcp::styles::descriptor_from_style(
            sheet.get_cell("A1").expect("A1 cell").get_style(),
        )
    })?;
    assert_eq!(desc_a1.font.as_ref().and_then(|f| f.bold), Some(true));

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn style_batch_overlap_ordering_last_wins() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("overlap.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        for addr in ["A1", "B1", "C1", "A2"] {
            sheet.get_cell_mut(addr).set_value("x");
        }
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
    let fork = create_fork(
        state.clone(),
        CreateForkParams {
            workbook_or_fork_id: workbook_id,
        },
    )
    .await?;

    let base_fill = StylePatch {
        fill: Some(Some(FillPatch::Pattern(PatternFillPatch {
            pattern_type: Some(Some("solid".to_string())),
            foreground_color: Some(Some("FFCCE5FF".to_string())),
            ..Default::default()
        }))),
        ..Default::default()
    };
    let header_bold = StylePatch {
        font: Some(Some(FontPatch {
            bold: Some(Some(true)),
            ..Default::default()
        })),
        ..Default::default()
    };

    style_batch(
        state.clone(),
        StyleBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![
                StyleOp {
                    sheet_name: "Sheet1".to_string(),
                    target: StyleTarget::Range {
                        range: "A1:C3".to_string(),
                    },
                    patch: base_fill,
                    op_mode: Some("set".to_string()),
                },
                StyleOp {
                    sheet_name: "Sheet1".to_string(),
                    target: StyleTarget::Range {
                        range: "A1:C1".to_string(),
                    },
                    patch: header_bold,
                    op_mode: Some("merge".to_string()),
                },
            ],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_wb = state
        .open_workbook(&spreadsheet_mcp::model::WorkbookId(fork.fork_id.clone()))
        .await?;
    let desc_a1 = fork_wb.with_sheet("Sheet1", |sheet| {
        spreadsheet_mcp::styles::descriptor_from_style(
            sheet.get_cell("A1").expect("A1").get_style(),
        )
    })?;
    let desc_a2 = fork_wb.with_sheet("Sheet1", |sheet| {
        spreadsheet_mcp::styles::descriptor_from_style(
            sheet.get_cell("A2").expect("A2").get_style(),
        )
    })?;

    assert_eq!(desc_a1.font.as_ref().and_then(|f| f.bold), Some(true));
    assert!(desc_a1.fill.is_some());
    assert!(desc_a2.fill.is_some());
    assert!(desc_a2.font.as_ref().and_then(|f| f.bold).is_none());

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn style_batch_nested_null_clear_only_subfield() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("null_clear.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_value("x");
        let style = sheet.get_style_mut("A1");
        style.get_font_mut().set_bold(true);
        style.get_font_mut().get_color_mut().set_argb("FFFF0000");
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
    let fork = create_fork(
        state.clone(),
        CreateForkParams {
            workbook_or_fork_id: workbook_id,
        },
    )
    .await?;

    let patch = StylePatch {
        font: Some(Some(FontPatch {
            color: Some(None),
            ..Default::default()
        })),
        ..Default::default()
    };

    style_batch(
        state.clone(),
        StyleBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StyleOp {
                sheet_name: "Sheet1".to_string(),
                target: StyleTarget::Cells {
                    cells: vec!["A1".to_string()],
                },
                patch,
                op_mode: Some("merge".to_string()),
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_wb = state
        .open_workbook(&spreadsheet_mcp::model::WorkbookId(fork.fork_id.clone()))
        .await?;
    let desc_a1 = fork_wb.with_sheet("Sheet1", |sheet| {
        spreadsheet_mcp::styles::descriptor_from_style(
            sheet.get_cell("A1").expect("A1").get_style(),
        )
    })?;
    assert_eq!(desc_a1.font.as_ref().and_then(|f| f.bold), Some(true));
    assert!(
        desc_a1
            .font
            .as_ref()
            .and_then(|f| f.color.clone())
            .is_none()
    );

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn style_batch_region_target_resolves() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("region.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_value("H1");
        sheet.get_cell_mut("B1").set_value("H2");
        sheet.get_cell_mut("C1").set_value("H3");
        for r in 2..=5 {
            sheet
                .get_cell_mut(format!("A{r}").as_str())
                .set_value_number(r);
            sheet
                .get_cell_mut(format!("B{r}").as_str())
                .set_value_number(r);
            sheet
                .get_cell_mut(format!("C{r}").as_str())
                .set_value_number(r);
        }
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
    let fork = create_fork(
        state.clone(),
        CreateForkParams {
            workbook_or_fork_id: workbook_id,
        },
    )
    .await?;

    let fork_wb = state
        .open_workbook(&spreadsheet_mcp::model::WorkbookId(fork.fork_id.clone()))
        .await?;
    let metrics = fork_wb.get_sheet_metrics("Sheet1")?;
    let regions = metrics.detected_regions();
    let region_id = regions.first().expect("region detected").id;

    let patch = StylePatch {
        font: Some(Some(FontPatch {
            bold: Some(Some(true)),
            ..Default::default()
        })),
        ..Default::default()
    };

    style_batch(
        state.clone(),
        StyleBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StyleOp {
                sheet_name: "Sheet1".to_string(),
                target: StyleTarget::Region { region_id },
                patch,
                op_mode: Some("merge".to_string()),
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_wb = state
        .open_workbook(&spreadsheet_mcp::model::WorkbookId(fork.fork_id.clone()))
        .await?;
    let desc_a1 = fork_wb.with_sheet("Sheet1", |sheet| {
        spreadsheet_mcp::styles::descriptor_from_style(
            sheet.get_cell("A1").expect("A1").get_style(),
        )
    })?;
    let desc_j1 = fork_wb.with_sheet("Sheet1", |sheet| {
        spreadsheet_mcp::styles::descriptor_from_style(sheet.get_style("J1"))
    })?;
    assert_eq!(desc_a1.font.as_ref().and_then(|f| f.bold), Some(true));
    assert!(desc_j1.font.is_none());

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn style_batch_idempotent_noop_counts_and_no_diff() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("noop.xlsx", |book| {
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
    let fork = create_fork(
        state.clone(),
        CreateForkParams {
            workbook_or_fork_id: workbook_id,
        },
    )
    .await?;

    let patch = StylePatch {
        font: Some(Some(FontPatch {
            bold: Some(Some(true)),
            ..Default::default()
        })),
        ..Default::default()
    };

    let resp = style_batch(
        state.clone(),
        StyleBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StyleOp {
                sheet_name: "Sheet1".to_string(),
                target: StyleTarget::Cells {
                    cells: vec!["A1".to_string()],
                },
                patch,
                op_mode: Some("merge".to_string()),
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    assert_eq!(
        resp.summary.counts.get("cells_style_changed").copied(),
        Some(0)
    );

    let changes = spreadsheet_mcp::tools::fork::get_changeset(
        state.clone(),
        spreadsheet_mcp::tools::fork::GetChangesetParams {
            fork_id: fork.fork_id.clone(),
            sheet_name: None,
            ..Default::default()
        },
    )
    .await?;
    use spreadsheet_mcp::diff::Change;
    use spreadsheet_mcp::diff::merge::{CellDiff, ModificationType};
    let non_style_change = changes.changes.iter().any(|c| match c {
        Change::Cell(cell) => match &cell.diff {
            CellDiff::Modified { subtype, .. } => !matches!(subtype, ModificationType::StyleEdit),
            CellDiff::Added { .. } | CellDiff::Deleted { .. } => true,
        },
        Change::Table(_) | Change::Name(_) => true,
    });
    assert!(!non_style_change);

    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn style_batch_preserves_conditional_formats() -> Result<()> {
    let workspace = support::TestWorkspace::new();
    workspace.create_workbook("cf.xlsx", |book| {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        sheet.get_cell_mut("A1").set_value_number(1);

        let mut cf = ConditionalFormatting::default();
        cf.get_sequence_of_references_mut().set_sqref("A1:A3");

        let mut rule = ConditionalFormattingRule::default();
        rule.set_type(ConditionalFormatValues::Expression);
        rule.set_priority(1);
        let mut formula = Formula::default();
        formula.set_string_value("A1>0");
        rule.set_formula(formula);

        let mut style = umya_spreadsheet::Style::default();
        style
            .get_fill_mut()
            .get_pattern_fill_mut()
            .set_pattern_type(PatternValues::Solid)
            .get_foreground_color_mut()
            .set_argb("FFFFFF00");
        rule.set_style(style);

        cf.add_conditional_collection(rule);
        sheet.add_conditional_formatting_collection(cf);
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
    let fork = create_fork(
        state.clone(),
        CreateForkParams {
            workbook_or_fork_id: workbook_id,
        },
    )
    .await?;

    let patch = StylePatch {
        font: Some(Some(FontPatch {
            italic: Some(Some(true)),
            ..Default::default()
        })),
        ..Default::default()
    };

    style_batch(
        state.clone(),
        StyleBatchParams {
            fork_id: fork.fork_id.clone(),
            ops: vec![StyleOp {
                sheet_name: "Sheet1".to_string(),
                target: StyleTarget::Range {
                    range: "A1:A3".to_string(),
                },
                patch,
                op_mode: Some("merge".to_string()),
            }],
            mode: Some("apply".to_string()),
            label: None,
        },
    )
    .await?;

    let fork_wb = state
        .open_workbook(&spreadsheet_mcp::model::WorkbookId(fork.fork_id.clone()))
        .await?;
    let cf_count = fork_wb.with_sheet("Sheet1", |sheet| {
        sheet.get_conditional_formatting_collection().len()
    })?;
    assert_eq!(cf_count, 1);

    Ok(())
}
