use crate::fork::{ChangeSummary, StagedChange, StagedOp};
use crate::model::WorkbookId;
use crate::state::AppState;
use crate::utils::make_short_random_id;
use anyhow::{Result, anyhow, bail};
use chrono::Utc;
use schemars::JsonSchema;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::Path;
use std::sync::Arc;
use umya_spreadsheet::{
    DataValidation, DataValidationOperatorValues, DataValidationValues, DataValidations,
};

fn default_mode() -> String {
    "apply".to_string()
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct RulesBatchParams {
    pub fork_id: String,
    pub ops: Vec<RulesOp>,
    #[serde(default = "default_mode")]
    pub mode: String, // preview|apply
    pub label: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum RulesOp {
    SetDataValidation {
        sheet_name: String,
        target_range: String,
        validation: DataValidationSpec,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct DataValidationSpec {
    pub kind: DataValidationKind,
    pub formula1: String,
    #[serde(default)]
    pub formula2: Option<String>,
    #[serde(default)]
    pub allow_blank: Option<bool>,
    #[serde(default)]
    pub prompt: Option<ValidationMessage>,
    #[serde(default)]
    pub error: Option<ValidationMessage>,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum DataValidationKind {
    List,
    Whole,
    Decimal,
    Date,
    Custom,
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
pub struct ValidationMessage {
    pub title: String,
    pub message: String,
}

#[derive(Debug, Serialize, JsonSchema)]
pub struct RulesBatchResponse {
    pub fork_id: String,
    pub mode: String,
    pub change_id: Option<String>,
    pub ops_applied: usize,
    pub summary: ChangeSummary,
}

#[derive(Debug, Serialize, Deserialize)]
pub(crate) struct RulesBatchStagedPayload {
    pub(crate) ops: Vec<RulesOp>,
}

pub async fn rules_batch(
    state: Arc<AppState>,
    params: RulesBatchParams,
) -> Result<RulesBatchResponse> {
    let registry = state
        .fork_registry()
        .ok_or_else(|| anyhow!("fork registry not available"))?;

    let fork_ctx = registry.get_fork(&params.fork_id)?;
    let work_path = fork_ctx.work_path.clone();

    // Validate sheet names early against current fork snapshot.
    let fork_workbook_id = WorkbookId(params.fork_id.clone());
    let workbook = state.open_workbook(&fork_workbook_id).await?;
    for op in &params.ops {
        match op {
            RulesOp::SetDataValidation { sheet_name, .. } => {
                let _ = workbook.with_sheet(sheet_name, |_| Ok::<_, anyhow::Error>(()))?;
            }
        }
    }

    let mode = params.mode.to_ascii_lowercase();
    if mode != "apply" && mode != "preview" {
        bail!(
            "invalid mode: {} (expected 'apply' or 'preview')",
            params.mode
        );
    }

    if mode == "preview" {
        let change_id = make_short_random_id("chg", 12);
        let snapshot_path = crate::tools::fork::stage_snapshot_path(&params.fork_id, &change_id);
        fs::create_dir_all(snapshot_path.parent().unwrap())?;
        fs::copy(&work_path, &snapshot_path)?;

        let snapshot_for_apply = snapshot_path.clone();
        let ops_for_apply = params.ops.clone();
        let apply_result = tokio::task::spawn_blocking(move || {
            apply_rules_ops_to_file(&snapshot_for_apply, &ops_for_apply)
        })
        .await??;

        let mut summary = apply_result.summary;
        summary
            .flags
            .insert("recalc_needed".to_string(), fork_ctx.recalc_needed);

        let staged_op = StagedOp {
            kind: "rules_batch".to_string(),
            payload: serde_json::to_value(RulesBatchStagedPayload {
                ops: params.ops.clone(),
            })?,
        };
        let staged = StagedChange {
            change_id: change_id.clone(),
            created_at: Utc::now(),
            label: params.label.clone(),
            ops: vec![staged_op],
            summary: summary.clone(),
            fork_path_snapshot: Some(snapshot_path),
        };
        registry.add_staged_change(&params.fork_id, staged)?;

        Ok(RulesBatchResponse {
            fork_id: params.fork_id,
            mode,
            change_id: Some(change_id),
            ops_applied: apply_result.ops_applied,
            summary,
        })
    } else {
        let work_path_for_apply = work_path.clone();
        let ops_for_apply = params.ops.clone();
        let apply_result = tokio::task::spawn_blocking(move || {
            apply_rules_ops_to_file(&work_path_for_apply, &ops_for_apply)
        })
        .await??;

        let mut summary = apply_result.summary;
        summary
            .flags
            .insert("recalc_needed".to_string(), fork_ctx.recalc_needed);

        let _ = state.close_workbook(&fork_workbook_id);

        Ok(RulesBatchResponse {
            fork_id: params.fork_id,
            mode,
            change_id: None,
            ops_applied: apply_result.ops_applied,
            summary,
        })
    }
}

pub(crate) struct RulesApplyResult {
    pub(crate) ops_applied: usize,
    pub(crate) summary: ChangeSummary,
}

pub(crate) fn apply_rules_ops_to_file(path: &Path, ops: &[RulesOp]) -> Result<RulesApplyResult> {
    let mut book = umya_spreadsheet::reader::xlsx::read(path)?;

    let mut affected_sheets: BTreeSet<String> = BTreeSet::new();
    let mut affected_bounds: Vec<String> = Vec::new();
    let mut counts: BTreeMap<String, u64> = BTreeMap::new();
    let mut warnings: Vec<String> = Vec::new();

    let mut validations_set: u64 = 0;
    let mut validations_replaced: u64 = 0;

    let mut warned_not_parsed = false;

    for op in ops {
        match op {
            RulesOp::SetDataValidation {
                sheet_name,
                target_range,
                validation,
            } => {
                let sheet = book
                    .get_sheet_by_name_mut(sheet_name)
                    .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))?;

                affected_sheets.insert(sheet_name.clone());
                affected_bounds.push(target_range.clone());

                if !warned_not_parsed {
                    warnings.push(
                        "WARN_VALIDATION_FORMULA_NOT_PARSED: Validation formulas are applied verbatim (not parsed or validated)."
                            .to_string(),
                    );
                    warned_not_parsed = true;
                }

                let (set_inc, replaced_inc) =
                    set_data_validation(sheet, target_range, validation, &mut warnings)?;
                validations_set += set_inc;
                validations_replaced += replaced_inc;
            }
        }
    }

    umya_spreadsheet::writer::xlsx::write(&book, path)?;

    counts.insert("validations_set".to_string(), validations_set);
    counts.insert("validations_replaced".to_string(), validations_replaced);

    Ok(RulesApplyResult {
        ops_applied: ops.len(),
        summary: ChangeSummary {
            op_kinds: vec!["rules_batch".to_string()],
            affected_sheets: affected_sheets.into_iter().collect(),
            affected_bounds,
            counts,
            warnings,
            ..Default::default()
        },
    })
}

fn normalize_sqref(input: &str) -> Result<String> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        bail!("target_range is required");
    }
    // DV sqref is space-separated list of ranges; v1 uses a single range.
    Ok(trimmed.replace(' ', "").to_ascii_uppercase())
}

fn normalize_dv_formula(field: &str, value: &str, warnings: &mut Vec<String>) -> String {
    let trimmed = value.trim();
    if let Some(stripped) = trimmed.strip_prefix('=') {
        warnings.push(format!(
            "WARN_VALIDATION_FORMULA_PREFIX: Stripped leading '=' from {field}"
        ));
        stripped.to_string()
    } else {
        trimmed.to_string()
    }
}

fn set_data_validation(
    sheet: &mut umya_spreadsheet::Worksheet,
    target_range: &str,
    spec: &DataValidationSpec,
    warnings: &mut Vec<String>,
) -> Result<(u64, u64)> {
    let sqref = normalize_sqref(target_range)?;

    if sheet.get_data_validations_mut().is_none() {
        sheet.set_data_validations(DataValidations::default());
    }
    let dvs = sheet
        .get_data_validations_mut()
        .ok_or_else(|| anyhow!("failed to initialize data validations"))?;

    // Remove any existing validations targeting the same sqref.
    let list = dvs.get_data_validation_list_mut();
    let before = list.len();
    list.retain(|dv| {
        let existing = dv.get_sequence_of_references().get_sqref();
        let existing_norm = existing.replace(' ', "").to_ascii_uppercase();
        existing_norm != sqref
    });
    let removed = before.saturating_sub(list.len());

    let mut dv = DataValidation::default();
    dv.set_type(spec.kind.to_umya());
    dv.get_sequence_of_references_mut().set_sqref(sqref.clone());

    if let Some(allow_blank) = spec.allow_blank {
        dv.set_allow_blank(allow_blank);
    }

    // Excel stores DV formulas without a leading '=' in OOXML.
    let formula1 = normalize_dv_formula("formula1", &spec.formula1, warnings);
    dv.set_formula1(formula1);
    if let Some(f2) = spec.formula2.as_ref() {
        let formula2 = normalize_dv_formula("formula2", f2, warnings);
        if !formula2.is_empty() {
            dv.set_formula2(formula2);
        }
    }

    // Operator: keep surface minimal; default to between when formula2 is provided.
    match spec.kind {
        DataValidationKind::Whole | DataValidationKind::Decimal | DataValidationKind::Date => {
            let op = if spec.formula2.as_ref().is_some_and(|s| !s.trim().is_empty()) {
                DataValidationOperatorValues::Between
            } else {
                DataValidationOperatorValues::Equal
            };
            dv.set_operator(op);
        }
        DataValidationKind::List | DataValidationKind::Custom => {}
    }

    if let Some(prompt) = spec.prompt.as_ref() {
        dv.set_show_input_message(true);
        if !prompt.title.is_empty() {
            dv.set_prompt_title(prompt.title.clone());
        }
        if !prompt.message.is_empty() {
            dv.set_prompt(prompt.message.clone());
        }
    }

    if let Some(error) = spec.error.as_ref() {
        dv.set_show_error_message(true);
        if !error.title.is_empty() {
            dv.set_error_title(error.title.clone());
        }
        if !error.message.is_empty() {
            dv.set_error_message(error.message.clone());
        }
    }

    dvs.add_data_validation_list(dv);

    Ok((1, if removed > 0 { 1 } else { 0 }))
}

impl DataValidationKind {
    fn to_umya(self) -> DataValidationValues {
        match self {
            DataValidationKind::List => DataValidationValues::List,
            DataValidationKind::Whole => DataValidationValues::Whole,
            DataValidationKind::Decimal => DataValidationValues::Decimal,
            DataValidationKind::Date => DataValidationValues::Date,
            DataValidationKind::Custom => DataValidationValues::Custom,
        }
    }
}
