use crate::core::types::CellEdit;
use crate::model::{
    CommandClass, FORMULA_PARSE_FAILED_PREFIX, FormulaParseDiagnostics,
    FormulaParseDiagnosticsBuilder, FormulaParsePolicy, Warning, validate_formula,
};
use crate::runtime::stateless::StatelessRuntime;
use crate::tools::fork::{
    ApplyFormulaPatternOpInput, ColumnSizeOp, ColumnSizeOpInput, StructureBatchParamsInput,
    StructureOp, StructureOpInput, StyleBatchParamsInput, StyleOp, StyleOpInput, TransformOp,
    TransformTarget, apply_column_size_ops_to_file, apply_formula_pattern_ops_to_file,
    apply_structure_ops_to_file, apply_style_ops_to_file, apply_transform_ops_to_file,
    normalize_column_size_payload, normalize_structure_batch, normalize_style_batch,
    resolve_style_ops_for_workbook, resolve_transform_ops_for_workbook,
};
use crate::tools::rules_batch::{RulesOp, apply_rules_ops_to_file};
use crate::tools::sheet_layout::{SheetLayoutOp, apply_sheet_layout_ops_to_file};
use anyhow::{Context, Result, anyhow, bail};
use schemars::{JsonSchema, schema_for};
use serde::{Deserialize, Serialize, de::DeserializeOwned};
use serde_json::Value;
use std::collections::BTreeMap;
use std::fs::{self, File, OpenOptions};
use std::io::ErrorKind;
use std::path::{Path, PathBuf};
use tempfile::{Builder, TempPath};

#[derive(Debug, Serialize)]
struct CopyResponse {
    source: String,
    dest: String,
    bytes_copied: u64,
}

#[derive(Debug, Serialize)]
struct CreateWorkbookResponse {
    path: String,
    sheets: Vec<String>,
    overwritten: bool,
}

#[derive(Debug, Clone, Serialize)]
struct WritePathProvenance {
    written_via: String,
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    formula_targets: Vec<String>,
}

#[derive(Debug, Serialize)]
struct EditResponse {
    file: String,
    sheet: String,
    edits_applied: usize,
    recalc_needed: bool,
    warnings: Vec<Warning>,
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    affected_cells: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    source_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    target_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    changed: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    formula_parse_diagnostics: Option<FormulaParseDiagnostics>,
    #[serde(skip_serializing_if = "Option::is_none")]
    write_path_provenance: Option<WritePathProvenance>,
}

#[derive(Debug, Serialize)]
struct EditDryRunResponse {
    file: String,
    sheet: String,
    edits_provided: usize,
    edits_validated: usize,
    would_change: bool,
    recalc_needed: bool,
    warnings: Vec<Warning>,
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    affected_cells: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    formula_parse_diagnostics: Option<FormulaParseDiagnostics>,
    #[serde(skip_serializing_if = "Option::is_none")]
    write_path_provenance: Option<WritePathProvenance>,
}

#[derive(Debug, Deserialize, JsonSchema)]
struct OpsPayload<T> {
    ops: Vec<T>,
}

#[derive(Debug, Deserialize, JsonSchema)]
struct ColumnSizeOpsPayload {
    sheet_name: String,
    ops: Vec<ColumnSizeOpInput>,
}

#[derive(Debug, Deserialize, JsonSchema)]
#[serde(untagged)]
enum ColumnSizeOpWithSheetInput {
    Canonical {
        sheet_name: String,
        target: crate::tools::fork::ColumnTarget,
        size: crate::tools::fork::ColumnSizeSpec,
    },
    Shorthand {
        sheet_name: String,
        range: String,
        size: crate::tools::fork::ColumnSizeSpec,
    },
}

impl ColumnSizeOpWithSheetInput {
    fn sheet_name(&self) -> &str {
        match self {
            Self::Canonical { sheet_name, .. } | Self::Shorthand { sheet_name, .. } => sheet_name,
        }
    }

    fn into_op_input(self) -> ColumnSizeOpInput {
        match self {
            Self::Canonical { target, size, .. } => {
                ColumnSizeOpInput::Canonical(ColumnSizeOp { target, size })
            }
            Self::Shorthand { range, size, .. } => ColumnSizeOpInput::Shorthand { range, size },
        }
    }
}

const TRANSFORM_PAYLOAD_SHAPE: &str = r#"{"ops":[{"kind":"<transform_kind>",...}]}"#;
const TRANSFORM_PAYLOAD_MINIMAL_EXAMPLE: &str = r#"{"ops":[{"kind":"fill_range","sheet_name":"Sheet1","target":{"kind":"range","range":"A1:A1"},"value":"1"}]}"#;
const STYLE_PAYLOAD_SHAPE: &str =
    r#"{"ops":[{"sheet_name":"...","target":{"kind":"range","range":"A1"},"patch":{...}}]}"#;
const STYLE_PAYLOAD_MINIMAL_EXAMPLE: &str = r#"{"ops":[{"sheet_name":"Sheet1","target":{"kind":"range","range":"B2:B2"},"patch":{"font":{"bold":true}}}]}"#;
const APPLY_FORMULA_PATTERN_PAYLOAD_SHAPE: &str = r#"{"ops":[{"sheet_name":"...","target_range":"A1:A1","anchor_cell":"A1","base_formula":"..."}]}"#;
const APPLY_FORMULA_PATTERN_PAYLOAD_MINIMAL_EXAMPLE: &str = r#"{"ops":[{"sheet_name":"Sheet1","target_range":"C2:C4","anchor_cell":"C2","base_formula":"B2*2"}]}"#;
const STRUCTURE_PAYLOAD_SHAPE: &str = r#"{"ops":[{"kind":"<structure_kind>",...}]}"#;
const STRUCTURE_PAYLOAD_MINIMAL_EXAMPLE: &str =
    r#"{"ops":[{"kind":"rename_sheet","old_name":"Summary","new_name":"Dashboard"}]}"#;
const COLUMN_SIZE_PAYLOAD_SHAPE: &str =
    r#"{"sheet_name":"...","ops":[{"range":"A:A","size":{"kind":"width","width_chars":12.0}}]}"#;
const COLUMN_SIZE_PAYLOAD_ALTERNATE_SHAPE: &str =
    r#"{"ops":[{"sheet_name":"...","range":"A:A","size":{"kind":"width","width_chars":12.0}}]}"#;
const COLUMN_SIZE_PAYLOAD_MINIMAL_EXAMPLE: &str =
    r#"{"sheet_name":"Sheet1","ops":[{"range":"A:A","size":{"kind":"width","width_chars":12.0}}]}"#;
const COLUMN_SIZE_PAYLOAD_ALTERNATE_EXAMPLE: &str =
    r#"{"ops":[{"sheet_name":"Sheet1","range":"A:A","size":{"kind":"width","width_chars":12.0}}]}"#;
const SHEET_LAYOUT_PAYLOAD_SHAPE: &str = r#"{"ops":[{"kind":"<layout_kind>",...}]}"#;
const SHEET_LAYOUT_PAYLOAD_MINIMAL_EXAMPLE: &str =
    r#"{"ops":[{"kind":"freeze_panes","sheet_name":"Sheet1","freeze_rows":1,"freeze_cols":1}]}"#;
const RULES_PAYLOAD_SHAPE: &str = r#"{"ops":[{"kind":"<rules_kind>",...}]}"#;
const RULES_PAYLOAD_MINIMAL_EXAMPLE: &str = r#"{"ops":[{"kind":"set_data_validation","sheet_name":"Sheet1","target_range":"B2:B4","validation":{"kind":"list","formula1":"\"A,B,C\""}}]}"#;
const EDIT_FORMULA_HINT: &str =
    "Tip: formulas in edit shorthand use double equals, e.g. A1==SUM(B1:B5).";

#[allow(dead_code)]
#[derive(Debug, JsonSchema)]
struct ColumnSizeOpsPerOpPayload {
    ops: Vec<ColumnSizeOpWithSheetInput>,
}

#[allow(dead_code)]
#[derive(Debug, JsonSchema)]
#[serde(untagged)]
enum ColumnSizeOpsSchemaPayload {
    Canonical(ColumnSizeOpsPayload),
    PerOp(ColumnSizeOpsPerOpPayload),
}

#[derive(Debug, Clone, Copy)]
pub enum BatchSchemaCommand {
    Transform,
    Style,
    ApplyFormulaPattern,
    Structure,
    ColumnSize,
    SheetLayout,
    Rules,
}

pub fn batch_payload_schema(command: BatchSchemaCommand) -> Result<Value> {
    let schema_value = match command {
        BatchSchemaCommand::Transform => {
            serde_json::to_value(schema_for!(OpsPayload<TransformOp>))?
        }
        BatchSchemaCommand::Style => serde_json::to_value(schema_for!(OpsPayload<StyleOpInput>))?,
        BatchSchemaCommand::ApplyFormulaPattern => {
            serde_json::to_value(schema_for!(OpsPayload<ApplyFormulaPatternOpInput>))?
        }
        BatchSchemaCommand::Structure => {
            serde_json::to_value(schema_for!(OpsPayload<StructureOpInput>))?
        }
        BatchSchemaCommand::ColumnSize => {
            serde_json::to_value(schema_for!(ColumnSizeOpsSchemaPayload))?
        }
        BatchSchemaCommand::SheetLayout => {
            serde_json::to_value(schema_for!(OpsPayload<SheetLayoutOp>))?
        }
        BatchSchemaCommand::Rules => serde_json::to_value(schema_for!(OpsPayload<RulesOp>))?,
    };

    Ok(serde_json::json!({
        "schema_kind": "ops_payload",
        "schema": schema_value,
    }))
}

#[derive(Debug)]
enum EditMutationMode {
    DryRun,
    InPlace,
    Output { target: PathBuf, force: bool },
}

#[derive(Debug)]
enum BatchMutationMode {
    DryRun,
    InPlace,
    Output { target: PathBuf, force: bool },
}

#[derive(Debug, Serialize)]
struct DryRunSummary {
    operation_counts: BTreeMap<String, u64>,
    result_counts: BTreeMap<String, u64>,
}

#[derive(Debug, Serialize)]
struct BatchDryRunResponse {
    op_count: usize,
    validated_count: usize,
    would_change: bool,
    warnings: Vec<Warning>,
    summary: DryRunSummary,
    #[serde(skip_serializing_if = "Option::is_none")]
    formula_parse_diagnostics: Option<FormulaParseDiagnostics>,
    #[serde(skip_serializing_if = "Option::is_none")]
    write_path_provenance: Option<WritePathProvenance>,
}

#[derive(Debug, Serialize)]
struct BatchApplyResponse {
    op_count: usize,
    applied_count: usize,
    warnings: Vec<Warning>,
    changed: bool,
    target_path: String,
    source_path: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    formula_parse_diagnostics: Option<FormulaParseDiagnostics>,
    #[serde(skip_serializing_if = "Option::is_none")]
    write_path_provenance: Option<WritePathProvenance>,
}

pub async fn copy(source: PathBuf, dest: PathBuf) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&source)?;
    let dest = runtime.normalize_destination_path(&dest)?;
    let bytes_copied = runtime.copy_file(&source, &dest).with_context(|| {
        format!(
            "failed to copy workbook from '{}' to '{}'",
            source.display(),
            dest.display()
        )
    })?;

    Ok(serde_json::to_value(CopyResponse {
        source: source.display().to_string(),
        dest: dest.display().to_string(),
        bytes_copied,
    })?)
}

pub async fn create_workbook(
    path: PathBuf,
    sheets: Option<Vec<String>>,
    overwrite: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let path = runtime.normalize_destination_path(&path)?;

    let existed = path.exists();
    if existed {
        if !overwrite {
            bail!(
                "file '{}' already exists; pass --overwrite to replace it",
                path.display()
            );
        }
        if !path.is_file() {
            bail!("path '{}' is not a file", path.display());
        }
    }

    let mut sheet_names = sheets.unwrap_or_else(|| vec!["Sheet1".to_string()]);
    if sheet_names.is_empty() {
        sheet_names.push("Sheet1".to_string());
    }

    let mut normalized_sheet_names = Vec::new();
    for name in sheet_names {
        let trimmed = name.trim();
        if trimmed.is_empty() {
            bail!("sheet names must be non-empty");
        }
        if normalized_sheet_names
            .iter()
            .any(|existing: &String| existing.eq_ignore_ascii_case(trimmed))
        {
            bail!("duplicate sheet name '{}'", trimmed);
        }
        normalized_sheet_names.push(trimmed.to_string());
    }

    let mut workbook = umya_spreadsheet::new_file();
    let first_sheet_name = normalized_sheet_names
        .first()
        .cloned()
        .ok_or_else(|| anyhow!("at least one sheet is required"))?;
    workbook
        .get_sheet_by_name_mut("Sheet1")
        .ok_or_else(|| anyhow!("failed to initialize workbook default sheet"))?
        .set_name(first_sheet_name.as_str());

    for sheet_name in normalized_sheet_names.iter().skip(1) {
        workbook
            .new_sheet(sheet_name.as_str())
            .map_err(|err| anyhow!("failed to create sheet '{}': {}", sheet_name, err))?;
    }

    umya_spreadsheet::writer::xlsx::write(&workbook, &path)
        .with_context(|| format!("failed to write workbook '{}'", path.display()))?;

    Ok(serde_json::to_value(CreateWorkbookResponse {
        path: path.display().to_string(),
        sheets: normalized_sheet_names,
        overwritten: existed,
    })?)
}

#[allow(clippy::too_many_arguments)]
pub async fn edit(
    file: PathBuf,
    sheet: String,
    edits: Vec<String>,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
    formula_parse_policy: Option<FormulaParsePolicy>,
) -> Result<Value> {
    if edits.is_empty() {
        bail!("at least one edit must be provided");
    }

    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_edit_mode(dry_run, in_place, output, force)?;

    let mut normalized_edits = Vec::with_capacity(edits.len());
    let mut warnings = Vec::new();
    for (idx, entry) in edits.into_iter().enumerate() {
        let (edit, entry_warnings) = crate::core::write::normalize_shorthand_edit(&entry)
            .with_context(|| {
                format!(
                    "invalid shorthand edit at index {}. {}",
                    idx, EDIT_FORMULA_HINT
                )
            })?;
        normalized_edits.push(edit);
        warnings.extend(entry_warnings.into_iter().map(|warning| Warning {
            code: warning.code,
            message: warning.message,
        }));
    }
    let edits_provided = normalized_edits.len();

    let policy = formula_parse_policy.unwrap_or(FormulaParsePolicy::default_for_command_class(
        CommandClass::SingleWrite,
    ));

    let (edits_to_write, formula_parse_diagnostics) = if policy == FormulaParsePolicy::Off {
        (normalized_edits, None)
    } else {
        let mut builder = FormulaParseDiagnosticsBuilder::new(policy);
        let mut valid_edits = Vec::new();
        for edit in normalized_edits {
            if edit.is_formula {
                match validate_formula(&edit.value) {
                    Ok(()) => valid_edits.push(edit),
                    Err(err_msg) => {
                        if policy == FormulaParsePolicy::Fail {
                            bail!(
                                "{}edit at {} failed: {}",
                                FORMULA_PARSE_FAILED_PREFIX,
                                edit.address,
                                err_msg
                            );
                        }
                        builder.record_error(&sheet, &edit.address, &edit.value, &err_msg);
                    }
                }
            } else {
                valid_edits.push(edit);
            }
        }
        let diagnostics = if builder.has_errors() {
            Some(builder.build())
        } else {
            None
        };
        (valid_edits, diagnostics)
    };

    let affected_cells = edits_to_write
        .iter()
        .map(|edit| edit.address.clone())
        .collect::<Vec<_>>();
    let changed = !edits_to_write.is_empty();
    let sheet_name = sheet;
    let write_path_provenance = formula_write_provenance(
        "edit",
        edits_to_write
            .iter()
            .filter(|edit| edit.is_formula)
            .map(|edit| format!("{}!{}", sheet_name, edit.address))
            .collect(),
    );

    match mode {
        EditMutationMode::DryRun => {
            let _ = apply_to_temp_copy(&source, source.parent(), ".edit-", |path| {
                runtime.apply_edits(path, &sheet_name, &edits_to_write)
            })?;

            Ok(serde_json::to_value(EditDryRunResponse {
                file: source.display().to_string(),
                sheet: sheet_name,
                edits_provided,
                edits_validated: edits_to_write.len(),
                would_change: changed,
                recalc_needed: false,
                warnings,
                affected_cells,
                formula_parse_diagnostics,
                write_path_provenance: write_path_provenance.clone(),
            })?)
        }
        EditMutationMode::InPlace => {
            apply_in_place_with_temp(&source, ".edit-", |path| {
                runtime.apply_edits(path, &sheet_name, &edits_to_write)
            })?;

            Ok(serde_json::to_value(EditResponse {
                file: source.display().to_string(),
                sheet: sheet_name,
                edits_applied: edits_to_write.len(),
                recalc_needed: true,
                warnings,
                affected_cells,
                source_path: None,
                target_path: None,
                changed: Some(changed),
                formula_parse_diagnostics,
                write_path_provenance: write_path_provenance.clone(),
            })?)
        }
        EditMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            apply_to_output_with_temp(&source, &target, force, ".edit-", |path| {
                runtime.apply_edits(path, &sheet_name, &edits_to_write)
            })?;

            Ok(serde_json::to_value(EditResponse {
                file: target.display().to_string(),
                sheet: sheet_name,
                edits_applied: edits_to_write.len(),
                recalc_needed: true,
                warnings,
                affected_cells,
                source_path: Some(source.display().to_string()),
                target_path: Some(target.display().to_string()),
                changed: Some(changed),
                formula_parse_diagnostics,
                write_path_provenance: write_path_provenance.clone(),
            })?)
        }
    }
}

pub async fn transform_batch(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
    formula_parse_policy: Option<FormulaParsePolicy>,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: OpsPayload<TransformOp> = parse_ops_payload(
        &ops,
        TRANSFORM_PAYLOAD_SHAPE,
        TRANSFORM_PAYLOAD_MINIMAL_EXAMPLE,
    )?;

    let (state, workbook_id) = runtime.open_state_for_file(&source).await?;
    let workbook = state.open_workbook(&workbook_id).await?;
    let resolved_ops = resolve_transform_ops_for_workbook(&workbook, &payload.ops)
        .map_err(|error| invalid_ops_payload(error.to_string()))?;
    let _ = state.close_workbook(&workbook_id);

    let policy = formula_parse_policy.unwrap_or(FormulaParsePolicy::default_for_command_class(
        CommandClass::BatchWrite,
    ));

    let (ops_to_apply, formula_parse_diagnostics) = if policy == FormulaParsePolicy::Off {
        (resolved_ops, None)
    } else {
        let mut builder = FormulaParseDiagnosticsBuilder::new(policy);
        let mut valid_ops = Vec::new();
        for op in resolved_ops {
            match &op {
                TransformOp::FillRange {
                    sheet_name,
                    value,
                    is_formula,
                    ..
                } if *is_formula => match validate_formula(value) {
                    Ok(()) => valid_ops.push(op),
                    Err(err_msg) => {
                        if policy == FormulaParsePolicy::Fail {
                            bail!(
                                "{}FillRange formula failed: {}",
                                FORMULA_PARSE_FAILED_PREFIX,
                                err_msg
                            );
                        }
                        builder.record_error(sheet_name, "FillRange", value, &err_msg);
                    }
                },
                _ => valid_ops.push(op),
            }
        }
        let diagnostics = if builder.has_errors() {
            Some(builder.build())
        } else {
            None
        };
        (valid_ops, diagnostics)
    };

    let op_count = ops_to_apply.len();
    let operation_counts = summarize_transform_operation_counts(&ops_to_apply);
    let write_path_provenance =
        formula_write_provenance("transform_batch", transform_formula_targets(&ops_to_apply));

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) =
                apply_to_temp_copy(&source, source.parent(), ".transform-batch-", |path| {
                    apply_transform_ops_to_file(path, &ops_to_apply).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let would_change = transform_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
                formula_parse_diagnostics,
                write_path_provenance.clone(),
            )
        }
        BatchMutationMode::InPlace => {
            let apply_result = apply_in_place_with_temp(&source, ".transform-batch-", |path| {
                apply_transform_ops_to_file(path, &ops_to_apply).map_err(classify_apply_error)
            })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = transform_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
                formula_parse_diagnostics,
                write_path_provenance.clone(),
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let apply_result =
                apply_to_output_with_temp(&source, &target, force, ".transform-batch-", |path| {
                    apply_transform_ops_to_file(path, &ops_to_apply).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = transform_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
                formula_parse_diagnostics,
                write_path_provenance.clone(),
            )
        }
    }
}

pub async fn style_batch(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: OpsPayload<StyleOpInput> =
        parse_ops_payload(&ops, STYLE_PAYLOAD_SHAPE, STYLE_PAYLOAD_MINIMAL_EXAMPLE)?;
    let (normalized, base_warnings) = normalize_style_batch(StyleBatchParamsInput {
        fork_id: String::new(),
        ops: payload.ops,
        mode: None,
        label: None,
    })
    .map_err(|error| invalid_ops_payload(error.to_string()))?;

    let (state, workbook_id) = runtime.open_state_for_file(&source).await?;
    let workbook = state.open_workbook(&workbook_id).await?;
    let resolved_ops = resolve_style_ops_for_workbook(&workbook, &normalized.ops)
        .map_err(|error| invalid_ops_payload(error.to_string()))?;
    let _ = state.close_workbook(&workbook_id);

    let op_count = resolved_ops.len();
    let operation_counts = summarize_style_operation_counts(&resolved_ops);

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) =
                apply_to_temp_copy(&source, source.parent(), ".style-batch-", |path| {
                    apply_style_ops_to_file(path, &resolved_ops).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings.clone(),
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let would_change = style_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
                None,
                None,
            )
        }
        BatchMutationMode::InPlace => {
            let apply_result = apply_in_place_with_temp(&source, ".style-batch-", |path| {
                apply_style_ops_to_file(path, &resolved_ops).map_err(classify_apply_error)
            })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings.clone(),
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let changed = style_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
                None,
                None,
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let apply_result =
                apply_to_output_with_temp(&source, &target, force, ".style-batch-", |path| {
                    apply_style_ops_to_file(path, &resolved_ops).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings,
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let changed = style_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
                None,
                None,
            )
        }
    }
}

pub async fn apply_formula_pattern(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: OpsPayload<ApplyFormulaPatternOpInput> = parse_ops_payload(
        &ops,
        APPLY_FORMULA_PATTERN_PAYLOAD_SHAPE,
        APPLY_FORMULA_PATTERN_PAYLOAD_MINIMAL_EXAMPLE,
    )?;

    let op_count = payload.ops.len();
    let operation_counts = summarize_formula_pattern_operation_counts(&payload.ops);
    let write_path_provenance = formula_write_provenance(
        "apply_formula_pattern",
        apply_formula_pattern_targets(&payload.ops),
    );

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) = apply_to_temp_copy(
                &source,
                source.parent(),
                ".apply-formula-pattern-",
                |path| {
                    apply_formula_pattern_ops_to_file(path, &payload.ops)
                        .map_err(classify_apply_error)
                },
            )?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let would_change = formula_pattern_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
                None,
                write_path_provenance.clone(),
            )
        }
        BatchMutationMode::InPlace => {
            let apply_result =
                apply_in_place_with_temp(&source, ".apply-formula-pattern-", |path| {
                    apply_formula_pattern_ops_to_file(path, &payload.ops)
                        .map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = formula_pattern_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
                None,
                write_path_provenance.clone(),
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let apply_result = apply_to_output_with_temp(
                &source,
                &target,
                force,
                ".apply-formula-pattern-",
                |path| {
                    apply_formula_pattern_ops_to_file(path, &payload.ops)
                        .map_err(classify_apply_error)
                },
            )?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = formula_pattern_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
                None,
                write_path_provenance.clone(),
            )
        }
    }
}

pub async fn structure_batch(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
    formula_parse_policy: Option<FormulaParsePolicy>,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: OpsPayload<StructureOpInput> = parse_ops_payload(
        &ops,
        STRUCTURE_PAYLOAD_SHAPE,
        STRUCTURE_PAYLOAD_MINIMAL_EXAMPLE,
    )?;
    let (normalized, base_warnings) = normalize_structure_batch(StructureBatchParamsInput {
        fork_id: String::new(),
        ops: payload.ops,
        mode: None,
        label: None,
        formula_parse_policy,
    })
    .map_err(|error| invalid_ops_payload(error.to_string()))?;

    let policy =
        normalized
            .formula_parse_policy
            .unwrap_or(FormulaParsePolicy::default_for_command_class(
                CommandClass::BatchWrite,
            ));

    let op_count = normalized.ops.len();
    let operation_counts = summarize_structure_operation_counts(&normalized.ops);

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) =
                apply_to_temp_copy(&source, source.parent(), ".structure-batch-", |path| {
                    apply_structure_ops_to_file(path, &normalized.ops, policy)
                        .map_err(classify_apply_error)
                })?;

            let formula_parse_diagnostics = apply_result.formula_parse_diagnostics;
            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings.clone(),
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let would_change = structure_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
                formula_parse_diagnostics,
                None,
            )
        }
        BatchMutationMode::InPlace => {
            let apply_result = apply_in_place_with_temp(&source, ".structure-batch-", |path| {
                apply_structure_ops_to_file(path, &normalized.ops, policy)
                    .map_err(classify_apply_error)
            })?;

            let formula_parse_diagnostics = apply_result.formula_parse_diagnostics;
            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings.clone(),
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let changed = structure_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
                formula_parse_diagnostics,
                None,
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let apply_result =
                apply_to_output_with_temp(&source, &target, force, ".structure-batch-", |path| {
                    apply_structure_ops_to_file(path, &normalized.ops, policy)
                        .map_err(classify_apply_error)
                })?;

            let formula_parse_diagnostics = apply_result.formula_parse_diagnostics;
            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings,
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let changed = structure_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
                formula_parse_diagnostics,
                None,
            )
        }
    }
}

pub async fn column_size_batch(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: ColumnSizeOpsPayload = parse_column_size_ops_payload(&ops)?;
    let (normalized_ops, base_warnings) =
        normalize_column_size_payload(payload.sheet_name.clone(), payload.ops)
            .map_err(|error| invalid_ops_payload(error.to_string()))?;

    let op_count = normalized_ops.len();
    let operation_counts = summarize_column_size_operation_counts(&normalized_ops);

    match mode {
        BatchMutationMode::DryRun => {
            let sheet_name = payload.sheet_name.clone();
            let (apply_result, _temp_path) =
                apply_to_temp_copy(&source, source.parent(), ".column-size-batch-", |path| {
                    apply_column_size_ops_to_file(path, &sheet_name, &normalized_ops)
                        .map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings.clone(),
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let would_change = column_size_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
                None,
                None,
            )
        }
        BatchMutationMode::InPlace => {
            let sheet_name = payload.sheet_name.clone();
            let apply_result = apply_in_place_with_temp(&source, ".column-size-batch-", |path| {
                apply_column_size_ops_to_file(path, &sheet_name, &normalized_ops)
                    .map_err(classify_apply_error)
            })?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings.clone(),
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let changed = column_size_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
                None,
                None,
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let sheet_name = payload.sheet_name;
            let apply_result = apply_to_output_with_temp(
                &source,
                &target,
                force,
                ".column-size-batch-",
                |path| {
                    apply_column_size_ops_to_file(path, &sheet_name, &normalized_ops)
                        .map_err(classify_apply_error)
                },
            )?;

            let result_counts = apply_result.summary.counts;
            let warnings = merge_cli_warnings(
                base_warnings,
                warning_strings_to_cli_warnings(apply_result.summary.warnings),
            );
            let changed = column_size_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
                None,
                None,
            )
        }
    }
}

pub async fn sheet_layout_batch(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: OpsPayload<SheetLayoutOp> = parse_ops_payload(
        &ops,
        SHEET_LAYOUT_PAYLOAD_SHAPE,
        SHEET_LAYOUT_PAYLOAD_MINIMAL_EXAMPLE,
    )?;

    let op_count = payload.ops.len();
    let operation_counts = summarize_sheet_layout_operation_counts(&payload.ops);

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) =
                apply_to_temp_copy(&source, source.parent(), ".sheet-layout-batch-", |path| {
                    apply_sheet_layout_ops_to_file(path, &payload.ops).map_err(classify_apply_error)
                })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let would_change = sheet_layout_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
                None,
                None,
            )
        }
        BatchMutationMode::InPlace => {
            let apply_result = apply_in_place_with_temp(&source, ".sheet-layout-batch-", |path| {
                apply_sheet_layout_ops_to_file(path, &payload.ops).map_err(classify_apply_error)
            })?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = sheet_layout_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
                None,
                None,
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let apply_result = apply_to_output_with_temp(
                &source,
                &target,
                force,
                ".sheet-layout-batch-",
                |path| {
                    apply_sheet_layout_ops_to_file(path, &payload.ops).map_err(classify_apply_error)
                },
            )?;

            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = sheet_layout_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
                None,
                None,
            )
        }
    }
}

pub async fn rules_batch(
    file: PathBuf,
    ops: String,
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
    formula_parse_policy: Option<FormulaParsePolicy>,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let source = runtime.normalize_existing_file(&file)?;
    let mode = validate_batch_mode(dry_run, in_place, output, force)?;

    let payload: OpsPayload<RulesOp> =
        parse_ops_payload(&ops, RULES_PAYLOAD_SHAPE, RULES_PAYLOAD_MINIMAL_EXAMPLE)?;

    let policy = formula_parse_policy.unwrap_or(FormulaParsePolicy::default_for_command_class(
        CommandClass::BatchWrite,
    ));

    let op_count = payload.ops.len();
    let operation_counts = summarize_rules_operation_counts(&payload.ops);

    match mode {
        BatchMutationMode::DryRun => {
            let (apply_result, _temp_path) =
                apply_to_temp_copy(&source, source.parent(), ".rules-batch-", |path| {
                    apply_rules_ops_to_file(path, &payload.ops, policy)
                        .map_err(classify_apply_error)
                })?;

            let formula_parse_diagnostics = apply_result.formula_parse_diagnostics;
            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let would_change = rules_summary_indicates_change(&result_counts);

            dry_run_response(
                op_count,
                operation_counts,
                result_counts,
                warnings,
                would_change,
                formula_parse_diagnostics,
                None,
            )
        }
        BatchMutationMode::InPlace => {
            let apply_result = apply_in_place_with_temp(&source, ".rules-batch-", |path| {
                apply_rules_ops_to_file(path, &payload.ops, policy).map_err(classify_apply_error)
            })?;

            let formula_parse_diagnostics = apply_result.formula_parse_diagnostics;
            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = rules_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                source.display().to_string(),
                source.display().to_string(),
                formula_parse_diagnostics,
                None,
            )
        }
        BatchMutationMode::Output { target, force } => {
            let target = runtime.normalize_destination_path(&target)?;
            ensure_output_path_is_distinct(&source, &target)?;

            let apply_result =
                apply_to_output_with_temp(&source, &target, force, ".rules-batch-", |path| {
                    apply_rules_ops_to_file(path, &payload.ops, policy)
                        .map_err(classify_apply_error)
                })?;

            let formula_parse_diagnostics = apply_result.formula_parse_diagnostics;
            let result_counts = apply_result.summary.counts;
            let warnings = warning_strings_to_cli_warnings(apply_result.summary.warnings);
            let changed = rules_summary_indicates_change(&result_counts);

            apply_response(
                op_count,
                apply_result.ops_applied,
                warnings,
                changed,
                target.display().to_string(),
                source.display().to_string(),
                formula_parse_diagnostics,
                None,
            )
        }
    }
}

fn validate_edit_mode(
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<EditMutationMode> {
    if force && output.is_none() {
        return Err(invalid_argument("--force requires --output <PATH>"));
    }

    if dry_run {
        if in_place {
            return Err(invalid_argument(
                "--dry-run cannot be combined with --in-place",
            ));
        }
        if output.is_some() {
            return Err(invalid_argument(
                "--dry-run cannot be combined with --output <PATH>",
            ));
        }
        return Ok(EditMutationMode::DryRun);
    }

    if in_place && output.is_some() {
        return Err(invalid_argument(
            "--in-place cannot be combined with --output <PATH>",
        ));
    }

    if let Some(target) = output {
        return Ok(EditMutationMode::Output { target, force });
    }

    Ok(EditMutationMode::InPlace)
}

fn validate_batch_mode(
    dry_run: bool,
    in_place: bool,
    output: Option<PathBuf>,
    force: bool,
) -> Result<BatchMutationMode> {
    if force && output.is_none() {
        return Err(invalid_argument("--force requires --output <PATH>"));
    }

    if dry_run {
        if in_place {
            return Err(invalid_argument(
                "--dry-run cannot be combined with --in-place",
            ));
        }
        if output.is_some() {
            return Err(invalid_argument(
                "--dry-run cannot be combined with --output <PATH>",
            ));
        }
        return Ok(BatchMutationMode::DryRun);
    }

    if in_place && output.is_some() {
        return Err(invalid_argument(
            "--in-place cannot be combined with --output <PATH>",
        ));
    }

    if in_place {
        return Ok(BatchMutationMode::InPlace);
    }

    if let Some(target) = output {
        return Ok(BatchMutationMode::Output { target, force });
    }

    Err(invalid_argument(
        "choose exactly one mutation mode: --dry-run, --in-place, or --output <PATH>",
    ))
}

fn parse_ops_payload_object(raw: &str, guidance: &str) -> Result<serde_json::Map<String, Value>> {
    let path = raw
        .strip_prefix('@')
        .ok_or_else(|| invalid_ops_payload("--ops must be provided as @<path>"))?;
    if path.is_empty() {
        return Err(invalid_ops_payload(
            "--ops file reference cannot be empty; expected @<path>",
        ));
    }

    let raw_payload = fs::read_to_string(path).map_err(|error| {
        invalid_ops_payload(format!("unable to read ops payload '{}': {}", path, error))
    })?;

    let json_value: serde_json::Value = serde_json::from_str(&raw_payload).map_err(|error| {
        invalid_ops_payload(format!(
            "ops payload is not valid JSON: {error}; {guidance}"
        ))
    })?;

    let object = json_value.as_object().ok_or_else(|| {
        invalid_ops_payload(format!("ops payload must be a JSON object; {guidance}"))
    })?;

    Ok(object.clone())
}

fn parse_column_size_ops_payload(raw: &str) -> Result<ColumnSizeOpsPayload> {
    let guidance = format!(
        "expected top-level shape: {} OR {}; minimal valid example: {} OR {}",
        COLUMN_SIZE_PAYLOAD_SHAPE,
        COLUMN_SIZE_PAYLOAD_ALTERNATE_SHAPE,
        COLUMN_SIZE_PAYLOAD_MINIMAL_EXAMPLE,
        COLUMN_SIZE_PAYLOAD_ALTERNATE_EXAMPLE,
    );

    let object = parse_ops_payload_object(raw, &guidance)?;

    if object.contains_key("sheet_name") {
        let top_level_sheet = object
            .get("sheet_name")
            .and_then(Value::as_str)
            .map(str::to_string);

        if let (Some(top_level_sheet), Some(ops_array)) =
            (top_level_sheet, object.get("ops").and_then(Value::as_array))
        {
            for (index, raw_entry) in ops_array.iter().enumerate() {
                if let Some(per_op_sheet) = raw_entry
                    .as_object()
                    .and_then(|entry| entry.get("sheet_name"))
                    .and_then(Value::as_str)
                    && per_op_sheet != top_level_sheet
                {
                    return Err(invalid_ops_payload(format!(
                        "ops payload has mixed sheet_name values between top-level and ops[{index}] ('{}' vs '{}'); {guidance}",
                        top_level_sheet, per_op_sheet
                    )));
                }
            }
        }

        return serde_json::from_value(Value::Object(object)).map_err(|error| {
            invalid_ops_payload(format!(
                "ops payload does not match required schema: {error}; {guidance}"
            ))
        });
    }

    let ops_value = object.get("ops").ok_or_else(|| {
        invalid_ops_payload(format!("ops payload must include 'ops'; {guidance}"))
    })?;
    let ops_array = ops_value.as_array().ok_or_else(|| {
        invalid_ops_payload(format!(
            "ops payload field 'ops' must be an array; {guidance}"
        ))
    })?;

    let mut normalized_ops = Vec::with_capacity(ops_array.len());
    let mut inferred_sheet_name: Option<String> = None;

    for (index, raw_entry) in ops_array.iter().enumerate() {
        let op_with_sheet: ColumnSizeOpWithSheetInput = serde_json::from_value(raw_entry.clone())
            .map_err(|error| {
            invalid_ops_payload(format!(
                "ops payload does not match required schema at ops[{index}]: {error}; {guidance}"
            ))
        })?;

        let sheet_name = op_with_sheet.sheet_name().to_string();
        match &inferred_sheet_name {
            Some(existing) if existing != &sheet_name => {
                return Err(invalid_ops_payload(format!(
                    "ops payload has mixed sheet_name values in per-op shape; found '{}' and '{}'; {guidance}",
                    existing, sheet_name
                )));
            }
            None => inferred_sheet_name = Some(sheet_name),
            _ => {}
        }

        normalized_ops.push(op_with_sheet.into_op_input());
    }

    let sheet_name = inferred_sheet_name.ok_or_else(|| {
        invalid_ops_payload(format!(
            "ops payload must provide top-level sheet_name or per-op sheet_name values; {guidance}"
        ))
    })?;

    Ok(ColumnSizeOpsPayload {
        sheet_name,
        ops: normalized_ops,
    })
}

fn parse_ops_payload<T: DeserializeOwned>(
    raw: &str,
    expected_shape: &str,
    minimal_example: &str,
) -> Result<T> {
    let guidance = format!(
        "expected top-level shape: {expected_shape}; minimal valid example: {minimal_example}"
    );
    let object = parse_ops_payload_object(raw, &guidance)?;

    serde_json::from_value(Value::Object(object)).map_err(|error| {
        invalid_ops_payload(format!(
            "ops payload does not match required schema: {error}; {guidance}"
        ))
    })
}

fn summarize_transform_operation_counts(ops: &[TransformOp]) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    for op in ops {
        let key = match op {
            TransformOp::ClearRange { .. } => "clear_range",
            TransformOp::FillRange { .. } => "fill_range",
            TransformOp::ReplaceInRange { .. } => "replace_in_range",
        };
        *counts.entry(key.to_string()).or_insert(0) += 1;
    }
    counts
}

fn summarize_style_operation_counts(ops: &[StyleOp]) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    counts.insert("style_ops".to_string(), ops.len() as u64);
    counts
}

fn summarize_formula_pattern_operation_counts(
    ops: &[ApplyFormulaPatternOpInput],
) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    counts.insert("apply_formula_pattern_ops".to_string(), ops.len() as u64);
    counts
}

fn summarize_structure_operation_counts(ops: &[StructureOp]) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    for op in ops {
        let key = match op {
            StructureOp::InsertRows { .. } => "insert_rows",
            StructureOp::DeleteRows { .. } => "delete_rows",
            StructureOp::InsertCols { .. } => "insert_cols",
            StructureOp::DeleteCols { .. } => "delete_cols",
            StructureOp::RenameSheet { .. } => "rename_sheet",
            StructureOp::CreateSheet { .. } => "create_sheet",
            StructureOp::DeleteSheet { .. } => "delete_sheet",
            StructureOp::CopyRange { .. } => "copy_range",
            StructureOp::MoveRange { .. } => "move_range",
        };
        *counts.entry(key.to_string()).or_insert(0) += 1;
    }
    counts
}

fn summarize_column_size_operation_counts(ops: &[ColumnSizeOp]) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    for op in ops {
        let key = match op.size {
            crate::tools::fork::ColumnSizeSpec::Auto { .. } => "auto",
            crate::tools::fork::ColumnSizeSpec::Width { .. } => "width",
        };
        *counts.entry(key.to_string()).or_insert(0) += 1;
    }
    counts
}

fn summarize_sheet_layout_operation_counts(ops: &[SheetLayoutOp]) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    for op in ops {
        let key = match op {
            SheetLayoutOp::FreezePanes { .. } => "freeze_panes",
            SheetLayoutOp::SetZoom { .. } => "set_zoom",
            SheetLayoutOp::SetGridlines { .. } => "set_gridlines",
            SheetLayoutOp::SetPageMargins { .. } => "set_page_margins",
            SheetLayoutOp::SetPageSetup { .. } => "set_page_setup",
            SheetLayoutOp::SetPrintArea { .. } => "set_print_area",
            SheetLayoutOp::SetPageBreaks { .. } => "set_page_breaks",
        };
        *counts.entry(key.to_string()).or_insert(0) += 1;
    }
    counts
}

fn summarize_rules_operation_counts(ops: &[RulesOp]) -> BTreeMap<String, u64> {
    let mut counts = BTreeMap::new();
    for op in ops {
        let key = match op {
            RulesOp::SetDataValidation { .. } => "set_data_validation",
            RulesOp::AddConditionalFormat { .. } => "add_conditional_format",
            RulesOp::SetConditionalFormat { .. } => "set_conditional_format",
            RulesOp::ClearConditionalFormats { .. } => "clear_conditional_formats",
        };
        *counts.entry(key.to_string()).or_insert(0) += 1;
    }
    counts
}

fn transform_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    const CHANGE_KEYS: &[&str] = &[
        "cells_value_cleared",
        "cells_formula_cleared",
        "cells_value_set",
        "cells_formula_set",
        "cells_value_replaced",
        "cells_formula_replaced",
    ];
    any_count_non_zero(counts, CHANGE_KEYS)
}

fn style_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    any_count_non_zero(counts, &["cells_style_changed"])
}

fn formula_pattern_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    any_count_non_zero(counts, &["cells_filled"])
}

fn structure_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    any_count_non_zero(
        counts,
        &[
            "rows_inserted",
            "rows_deleted",
            "cols_inserted",
            "cols_deleted",
            "sheets_renamed",
            "sheets_created",
            "sheets_deleted",
            "cells_copied",
            "cells_moved",
            "ranges_copied",
            "ranges_moved",
        ],
    )
}

fn column_size_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    any_count_non_zero(counts, &["columns_sized"])
}

fn sheet_layout_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    any_count_non_zero(
        counts,
        &[
            "ops",
            "freeze_panes_ops",
            "set_zoom_ops",
            "set_gridlines_ops",
            "set_page_margins_ops",
            "set_page_setup_ops",
            "set_print_area_ops",
            "set_page_breaks_ops",
        ],
    )
}

fn rules_summary_indicates_change(counts: &BTreeMap<String, u64>) -> bool {
    any_count_non_zero(
        counts,
        &[
            "validations_set",
            "validations_replaced",
            "conditional_formats_added",
            "conditional_formats_set",
            "conditional_formats_replaced",
            "conditional_formats_cleared",
        ],
    )
}

fn any_count_non_zero(counts: &BTreeMap<String, u64>, keys: &[&str]) -> bool {
    keys.iter()
        .any(|key| counts.get(*key).copied().unwrap_or(0) > 0)
}

fn warning_strings_to_cli_warnings(messages: Vec<String>) -> Vec<Warning> {
    messages.into_iter().map(parse_warning_message).collect()
}

fn merge_cli_warnings(mut left: Vec<Warning>, mut right: Vec<Warning>) -> Vec<Warning> {
    left.append(&mut right);
    left
}

fn parse_warning_message(message: String) -> Warning {
    if let Some((code, detail)) = message.split_once(':') {
        let code = code.trim();
        let detail = detail.trim();
        if is_warning_code(code) && !detail.is_empty() {
            return Warning {
                code: code.to_string(),
                message: detail.to_string(),
            };
        }
    }

    Warning {
        code: "WARN_INFO".to_string(),
        message,
    }
}

fn is_warning_code(value: &str) -> bool {
    value.starts_with("WARN_")
        && value
            .chars()
            .all(|ch| ch.is_ascii_uppercase() || ch == '_' || ch.is_ascii_digit())
}

fn formula_write_provenance(
    written_via: &str,
    formula_targets: Vec<String>,
) -> Option<WritePathProvenance> {
    if formula_targets.is_empty() {
        None
    } else {
        Some(WritePathProvenance {
            written_via: written_via.to_string(),
            formula_targets,
        })
    }
}

fn transform_formula_targets(ops: &[TransformOp]) -> Vec<String> {
    ops.iter()
        .filter_map(|op| match op {
            TransformOp::FillRange {
                sheet_name,
                target,
                is_formula,
                ..
            } if *is_formula => Some(format!("{}!{}", sheet_name, transform_target_label(target))),
            TransformOp::ReplaceInRange {
                sheet_name,
                target,
                include_formulas,
                ..
            } if *include_formulas => {
                Some(format!("{}!{}", sheet_name, transform_target_label(target)))
            }
            _ => None,
        })
        .collect()
}

fn transform_target_label(target: &TransformTarget) -> String {
    match target {
        TransformTarget::Range { range } => range.clone(),
        TransformTarget::Region { region_id } => format!("region:{}", region_id),
        TransformTarget::Cells { cells } => {
            if cells.is_empty() {
                "cells".to_string()
            } else {
                format!("cells:{}", cells.join(","))
            }
        }
    }
}

fn apply_formula_pattern_targets(ops: &[ApplyFormulaPatternOpInput]) -> Vec<String> {
    ops.iter()
        .map(|op| format!("{}!{}", op.sheet_name, op.target_range))
        .collect()
}

fn dry_run_response(
    op_count: usize,
    operation_counts: BTreeMap<String, u64>,
    result_counts: BTreeMap<String, u64>,
    warnings: Vec<Warning>,
    would_change: bool,
    formula_parse_diagnostics: Option<FormulaParseDiagnostics>,
    write_path_provenance: Option<WritePathProvenance>,
) -> Result<Value> {
    Ok(serde_json::to_value(BatchDryRunResponse {
        op_count,
        validated_count: op_count,
        would_change,
        warnings,
        summary: DryRunSummary {
            operation_counts,
            result_counts,
        },
        formula_parse_diagnostics,
        write_path_provenance,
    })?)
}

#[allow(clippy::too_many_arguments)]
fn apply_response(
    op_count: usize,
    applied_count: usize,
    warnings: Vec<Warning>,
    changed: bool,
    target_path: String,
    source_path: String,
    formula_parse_diagnostics: Option<FormulaParseDiagnostics>,
    write_path_provenance: Option<WritePathProvenance>,
) -> Result<Value> {
    Ok(serde_json::to_value(BatchApplyResponse {
        op_count,
        applied_count,
        warnings,
        changed,
        target_path,
        source_path,
        formula_parse_diagnostics,
        write_path_provenance,
    })?)
}

fn apply_in_place_with_temp<T, F>(source: &Path, temp_prefix: &str, apply_fn: F) -> Result<T>
where
    F: FnOnce(&Path) -> Result<T>,
{
    let (apply_result, temp_path) =
        apply_to_temp_copy(source, source.parent(), temp_prefix, apply_fn)?;
    atomic_replace_target(temp_path, source, true)?;
    Ok(apply_result)
}

fn apply_to_output_with_temp<T, F>(
    source: &Path,
    target: &Path,
    force: bool,
    temp_prefix: &str,
    apply_fn: F,
) -> Result<T>
where
    F: FnOnce(&Path) -> Result<T>,
{
    let target_exists = path_entry_exists(target)?;
    if target_exists && !force {
        return Err(output_exists(format!(
            "output path '{}' already exists",
            target.display()
        )));
    }

    let (apply_result, temp_path) =
        apply_to_temp_copy(source, target.parent(), temp_prefix, apply_fn)?;
    atomic_replace_target(temp_path, target, force)?;
    Ok(apply_result)
}

fn apply_to_temp_copy<T, F>(
    source: &Path,
    directory: Option<&Path>,
    temp_prefix: &str,
    apply_fn: F,
) -> Result<(T, TempPath)>
where
    F: FnOnce(&Path) -> Result<T>,
{
    let parent = directory.ok_or_else(|| {
        write_failed(format!(
            "unable to create temp file: '{}' has no parent directory",
            source.display()
        ))
    })?;
    let temp_path = Builder::new()
        .prefix(temp_prefix)
        .suffix(".tmp.xlsx")
        .tempfile_in(parent)
        .map_err(|error| {
            write_failed(format!(
                "unable to allocate temp file in '{}': {}",
                parent.display(),
                error
            ))
        })?
        .into_temp_path();

    let temp_path_ref: &Path = temp_path.as_ref();

    fs::copy(source, temp_path_ref).map_err(|error| {
        write_failed(format!(
            "unable to stage temp workbook from '{}' to '{}': {}",
            source.display(),
            temp_path.display(),
            error
        ))
    })?;

    let apply_result = apply_fn(temp_path_ref)?;

    fsync_file(temp_path_ref)?;

    Ok((apply_result, temp_path))
}

fn atomic_replace_target(temp_path: TempPath, target: &Path, allow_overwrite: bool) -> Result<()> {
    if allow_overwrite {
        let target_exists = path_entry_exists(target)?;
        if target_exists && !atomic_overwrite_supported() {
            return Err(write_failed(
                "atomic overwrite is not supported on this platform",
            ));
        }

        let temp_path_ref: &Path = temp_path.as_ref();
        fs::rename(temp_path_ref, target).map_err(|error| {
            write_failed(format!(
                "unable to atomically replace '{}' from '{}': {}",
                target.display(),
                temp_path.display(),
                error
            ))
        })?;
    } else {
        temp_path.persist_noclobber(target).map_err(|error| {
            if error.error.kind() == ErrorKind::AlreadyExists {
                output_exists(format!("output path '{}' already exists", target.display()))
            } else {
                write_failed(format!(
                    "unable to move staged workbook '{}' to '{}': {}",
                    error.path.display(),
                    target.display(),
                    error.error
                ))
            }
        })?;
    }

    if let Some(parent) = target.parent() {
        fsync_directory(parent)?;
    }

    Ok(())
}

fn fsync_file(path: &Path) -> Result<()> {
    let file = OpenOptions::new()
        .read(true)
        .write(true)
        .open(path)
        .map_err(|error| {
            write_failed(format!(
                "unable to open '{}' for fsync: {}",
                path.display(),
                error
            ))
        })?;
    file.sync_all().map_err(|error| {
        write_failed(format!(
            "unable to fsync temp file '{}': {}",
            path.display(),
            error
        ))
    })
}

#[cfg(unix)]
fn fsync_directory(path: &Path) -> Result<()> {
    let dir = File::open(path).map_err(|error| {
        write_failed(format!(
            "unable to open directory '{}' for fsync: {}",
            path.display(),
            error
        ))
    })?;
    dir.sync_all().map_err(|error| {
        write_failed(format!(
            "unable to fsync directory '{}': {}",
            path.display(),
            error
        ))
    })
}

#[cfg(not(unix))]
fn fsync_directory(_path: &Path) -> Result<()> {
    Ok(())
}

fn path_entry_exists(path: &Path) -> Result<bool> {
    match fs::symlink_metadata(path) {
        Ok(_) => Ok(true),
        Err(error) if error.kind() == ErrorKind::NotFound => Ok(false),
        Err(error) => Err(write_failed(format!(
            "unable to inspect output path '{}': {}",
            path.display(),
            error
        ))),
    }
}

fn ensure_output_path_is_distinct(source: &Path, output: &Path) -> Result<()> {
    let source_identity = canonical_identity_path(source)?;
    let output_identity = canonical_identity_path(output)?;
    if source_identity == output_identity {
        return Err(invalid_argument(
            "--output path resolves to the same file as input",
        ));
    }
    Ok(())
}

fn canonical_identity_path(path: &Path) -> Result<PathBuf> {
    if path.exists() {
        return fs::canonicalize(path).with_context(|| {
            format!(
                "failed to resolve canonical identity path for '{}'",
                path.display()
            )
        });
    }

    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    let name = path
        .file_name()
        .ok_or_else(|| invalid_argument("output path must include a file name"))?;

    let parent_canonical = fs::canonicalize(parent).with_context(|| {
        format!(
            "failed to resolve output parent directory '{}': {}",
            parent.display(),
            "directory does not exist or is inaccessible"
        )
    })?;

    Ok(parent_canonical.join(name))
}

#[cfg(unix)]
fn atomic_overwrite_supported() -> bool {
    true
}

#[cfg(not(unix))]
fn atomic_overwrite_supported() -> bool {
    false
}

fn classify_apply_error(error: anyhow::Error) -> anyhow::Error {
    let message = error.to_string();
    if message.starts_with(FORMULA_PARSE_FAILED_PREFIX) {
        return error;
    }

    if error
        .chain()
        .any(|cause| cause.downcast_ref::<std::io::Error>().is_some())
    {
        write_failed(format!("failed while applying ops payload: {}", message))
    } else {
        invalid_ops_payload(message)
    }
}

fn invalid_argument(message: impl AsRef<str>) -> anyhow::Error {
    anyhow!("invalid argument: {}", message.as_ref())
}

fn invalid_ops_payload(message: impl AsRef<str>) -> anyhow::Error {
    anyhow!("invalid ops payload: {}", message.as_ref())
}

fn output_exists(message: impl AsRef<str>) -> anyhow::Error {
    anyhow!("output exists: {}", message.as_ref())
}

fn write_failed(message: impl AsRef<str>) -> anyhow::Error {
    anyhow!("write failed: {}", message.as_ref())
}

pub fn parse_shorthand_for_tests(entries: Vec<String>) -> Result<(Vec<CellEdit>, Vec<Warning>)> {
    let mut edits = Vec::with_capacity(entries.len());
    let mut warnings = Vec::new();
    for entry in entries {
        let (edit, entry_warnings) = crate::core::write::normalize_shorthand_edit(&entry)?;
        edits.push(edit);
        warnings.extend(entry_warnings.into_iter().map(|warning| Warning {
            code: warning.code,
            message: warning.message,
        }));
    }
    Ok((edits, warnings))
}
