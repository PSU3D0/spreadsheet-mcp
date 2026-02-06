use crate::fork::{ChangeSummary, StagedChange, StagedOp};
use crate::model::WorkbookId;
use crate::state::AppState;
use crate::utils::make_short_random_id;
use crate::{rules::conditional_format, styles::normalize_color_hex};
use anyhow::{Result, anyhow, bail};
use chrono::Utc;
use schemars::JsonSchema;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::Path;
use std::sync::Arc;
use umya_spreadsheet::{
    ConditionalFormattingOperatorValues, DataValidation, DataValidationOperatorValues,
    DataValidationValues, DataValidations,
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
    AddConditionalFormat {
        sheet_name: String,
        target_range: String,
        rule: ConditionalFormatRuleSpec,
        #[serde(default)]
        style: ConditionalFormatStyleSpec,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize, JsonSchema)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum ConditionalFormatRuleSpec {
    CellIs {
        operator: ConditionalFormatOperator,
        formula: String,
    },
    Expression {
        formula: String,
    },
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum ConditionalFormatOperator {
    #[serde(alias = "lessThan")]
    LessThan,
    #[serde(alias = "lessThanOrEqual")]
    LessThanOrEqual,
    #[serde(alias = "greaterThan")]
    GreaterThan,
    #[serde(alias = "greaterThanOrEqual")]
    GreaterThanOrEqual,
    #[serde(alias = "equal")]
    Equal,
    #[serde(alias = "notEqual")]
    NotEqual,
    #[serde(alias = "between")]
    Between,
    #[serde(alias = "notBetween")]
    NotBetween,
}

impl ConditionalFormatOperator {
    fn to_umya(self) -> ConditionalFormattingOperatorValues {
        match self {
            Self::LessThan => ConditionalFormattingOperatorValues::LessThan,
            Self::LessThanOrEqual => ConditionalFormattingOperatorValues::LessThanOrEqual,
            Self::GreaterThan => ConditionalFormattingOperatorValues::GreaterThan,
            Self::GreaterThanOrEqual => ConditionalFormattingOperatorValues::GreaterThanOrEqual,
            Self::Equal => ConditionalFormattingOperatorValues::Equal,
            Self::NotEqual => ConditionalFormattingOperatorValues::NotEqual,
            Self::Between => ConditionalFormattingOperatorValues::Between,
            Self::NotBetween => ConditionalFormattingOperatorValues::NotBetween,
        }
    }
}

#[derive(Debug, Clone, Default, Serialize, Deserialize, JsonSchema)]
pub struct ConditionalFormatStyleSpec {
    #[serde(default)]
    pub fill_color: Option<String>,
    #[serde(default)]
    pub font_color: Option<String>,
    #[serde(default)]
    pub bold: Option<bool>,
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
            RulesOp::AddConditionalFormat { sheet_name, .. } => {
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
    let mut conditional_formats_added: u64 = 0;
    let mut conditional_formats_skipped: u64 = 0;

    let mut warned_not_parsed = false;
    let mut warned_cf_structure = false;

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
            RulesOp::AddConditionalFormat {
                sheet_name,
                target_range,
                rule,
                style,
            } => {
                let sheet = book
                    .get_sheet_by_name_mut(sheet_name)
                    .ok_or_else(|| anyhow!("sheet '{}' not found", sheet_name))?;

                affected_sheets.insert(sheet_name.clone());
                affected_bounds.push(target_range.clone());

                if !warned_cf_structure {
                    warnings.push("WARN_CF_FORMULA_NOT_ADJUSTED_ON_STRUCTURE: Conditional format formulas are not automatically rewritten on structural edits; re-apply or review after row/col insertion/deletion.".to_string());
                    warned_cf_structure = true;
                }

                let (added, skipped) =
                    add_conditional_format(sheet, target_range, rule, style, &mut warnings)?;
                conditional_formats_added += added;
                conditional_formats_skipped += skipped;
            }
        }
    }

    umya_spreadsheet::writer::xlsx::write(&book, path)?;

    counts.insert("validations_set".to_string(), validations_set);
    counts.insert("validations_replaced".to_string(), validations_replaced);
    counts.insert(
        "conditional_formats_added".to_string(),
        conditional_formats_added,
    );
    counts.insert(
        "conditional_formats_skipped".to_string(),
        conditional_formats_skipped,
    );

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

fn normalize_cf_formula(field: &str, value: &str, warnings: &mut Vec<String>) -> Result<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        bail!("{field} is required");
    }
    if let Some(stripped) = trimmed.strip_prefix('=') {
        warnings.push(format!(
            "WARN_CF_FORMULA_PREFIX: Stripped leading '=' from {field}"
        ));
        return Ok(stripped.to_string());
    }
    Ok(trimmed.to_string())
}

fn normalize_argb_color(field: &str, input: &str, warnings: &mut Vec<String>) -> Result<String> {
    let trimmed = input.trim();
    let Some((argb, defaulted_alpha)) = normalize_color_hex(trimmed) else {
        bail!("invalid color for {field}: expected #RGB/#RRGGBB/#AARRGGBB");
    };
    if defaulted_alpha {
        warnings.push(format!(
            "WARN_COLOR_ALPHA_DEFAULT: Defaulted alpha to FF for {field}"
        ));
    }
    Ok(argb)
}

fn add_conditional_format(
    sheet: &mut umya_spreadsheet::Worksheet,
    target_range: &str,
    rule: &ConditionalFormatRuleSpec,
    style: &ConditionalFormatStyleSpec,
    warnings: &mut Vec<String>,
) -> Result<(u64, u64)> {
    let sqref = normalize_sqref(target_range)?;

    let desired = match rule {
        ConditionalFormatRuleSpec::Expression { formula } => (
            umya_spreadsheet::ConditionalFormatValues::Expression,
            None,
            normalize_cf_formula("rule.formula", formula, warnings)?,
        ),
        ConditionalFormatRuleSpec::CellIs { operator, formula } => (
            umya_spreadsheet::ConditionalFormatValues::CellIs,
            Some(operator.to_umya()),
            normalize_cf_formula("rule.formula", formula, warnings)?,
        ),
    };

    // Defaults aim for determinism and readability.
    let fill = style.fill_color.as_deref().unwrap_or("FFFFE0E0");
    let font = style.font_color.as_deref().unwrap_or("FF000000");
    let bold = style.bold.unwrap_or(false);

    let fill_argb = normalize_argb_color("style.fill_color", fill, warnings)?;
    let font_argb = normalize_argb_color("style.font_color", font, warnings)?;

    // Deduplicate exact matches (sqref + kind/operator + formula).
    for existing in sheet.get_conditional_formatting_collection() {
        if existing.get_sequence_of_references().get_sqref() != sqref {
            continue;
        }
        for existing_rule in existing.get_conditional_collection() {
            if existing_rule.get_type() != &desired.0 {
                continue;
            }
            if let Some(ref op) = desired.1 {
                if existing_rule.get_operator() != op {
                    continue;
                }
            }
            let existing_formula = existing_rule
                .get_formula()
                .map(|f| f.get_address_str())
                .unwrap_or_default();
            if existing_formula == desired.2 {
                return Ok((0, 1));
            }
        }
    }

    let dxf_style = conditional_format::build_simple_dxf_style(&fill_argb, &font_argb, bold);

    match desired.0 {
        umya_spreadsheet::ConditionalFormatValues::Expression => {
            conditional_format::append_cf_expression_rule(sheet, &sqref, &desired.2, dxf_style);
        }
        umya_spreadsheet::ConditionalFormatValues::CellIs => {
            conditional_format::append_cf_cellis_rule(
                sheet,
                &sqref,
                desired
                    .1
                    .clone()
                    .unwrap_or(ConditionalFormattingOperatorValues::LessThan),
                &desired.2,
                dxf_style,
            );
        }
        _ => unreachable!("only expression and cellIs are supported"),
    }

    Ok((1, 0))
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
