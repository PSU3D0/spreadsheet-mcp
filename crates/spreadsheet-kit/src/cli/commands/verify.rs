use crate::model::{CellValue, NamedItemKind, NamedRangeDescriptor, NamedRangeScope};
use crate::runtime::stateless::StatelessRuntime;
use crate::tools::{self, NamedRangesParams};
use crate::workbook::{WorkbookContext, cell_to_value};
use anyhow::{Result, anyhow, bail};
use serde::Serialize;
use serde_json::Value;
use std::collections::{BTreeMap, BTreeSet};
use std::path::PathBuf;

#[derive(Debug, Serialize)]
struct VerifySummary {
    target_count: u32,
    changed_targets: u32,
    new_error_count: u32,
    preexisting_error_count: u32,
    named_range_delta_count: u32,
}

#[derive(Debug, Serialize)]
struct TargetDelta {
    address: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    before: Option<CellValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    after: Option<CellValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    before_formula: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    after_formula: Option<String>,
    changed: bool,
}

#[derive(Debug, Clone, Serialize)]
struct ErrorDelta {
    address: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    before_error: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    after_error: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    before_formula: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    after_formula: Option<String>,
}

#[derive(Debug, Serialize)]
struct NamedRangeDelta {
    name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    scope_kind: Option<NamedRangeScope>,
    #[serde(skip_serializing_if = "Option::is_none")]
    scope_sheet_name: Option<String>,
    change: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    before_refers_to: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    after_refers_to: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    before_kind: Option<NamedItemKind>,
    #[serde(skip_serializing_if = "Option::is_none")]
    after_kind: Option<NamedItemKind>,
}

#[derive(Debug, Serialize)]
struct VerifyResponse {
    baseline: String,
    current: String,
    target_deltas: Vec<TargetDelta>,
    new_errors: Vec<ErrorDelta>,
    preexisting_errors: Vec<ErrorDelta>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    named_range_deltas: Vec<NamedRangeDelta>,
    summary: VerifySummary,
}

#[derive(Debug, Clone)]
struct TargetCellSnapshot {
    value: Option<CellValue>,
    formula: Option<String>,
}

#[derive(Debug, Clone)]
struct ErrorCellSnapshot {
    error: String,
    formula: Option<String>,
}

pub async fn verify(
    baseline: PathBuf,
    current: PathBuf,
    targets: Option<Vec<String>>,
    named_ranges: bool,
) -> Result<Value> {
    let runtime = StatelessRuntime;
    let baseline = runtime.normalize_existing_file(&baseline)?;
    let current = runtime.normalize_existing_file(&current)?;

    let (baseline_state, baseline_id) = runtime.open_state_for_file(&baseline).await?;
    let (current_state, current_id) = runtime.open_state_for_file(&current).await?;

    let baseline_workbook = baseline_state.open_workbook(&baseline_id).await?;
    let current_workbook = current_state.open_workbook(&current_id).await?;

    let target_deltas = collect_target_deltas(
        &baseline_workbook,
        &current_workbook,
        targets.unwrap_or_default(),
    )?;

    let baseline_errors = collect_error_cells(&baseline_workbook)?;
    let current_errors = collect_error_cells(&current_workbook)?;
    let (new_errors, preexisting_errors) = compare_error_maps(&baseline_errors, &current_errors);

    let named_range_deltas = if named_ranges {
        let baseline_named = tools::named_ranges(
            baseline_state.clone(),
            NamedRangesParams {
                workbook_or_fork_id: baseline_id,
                sheet_name: None,
                name_prefix: None,
            },
        )
        .await?;
        let current_named = tools::named_ranges(
            current_state.clone(),
            NamedRangesParams {
                workbook_or_fork_id: current_id,
                sheet_name: None,
                name_prefix: None,
            },
        )
        .await?;
        compare_named_ranges(&baseline_named.items, &current_named.items)
    } else {
        Vec::new()
    };

    let response = VerifyResponse {
        baseline: baseline.display().to_string(),
        current: current.display().to_string(),
        summary: VerifySummary {
            target_count: target_deltas.len() as u32,
            changed_targets: target_deltas.iter().filter(|d| d.changed).count() as u32,
            new_error_count: new_errors.len() as u32,
            preexisting_error_count: preexisting_errors.len() as u32,
            named_range_delta_count: named_range_deltas.len() as u32,
        },
        target_deltas,
        new_errors,
        preexisting_errors,
        named_range_deltas,
    };

    Ok(serde_json::to_value(response)?)
}

fn collect_target_deltas(
    baseline: &WorkbookContext,
    current: &WorkbookContext,
    targets: Vec<String>,
) -> Result<Vec<TargetDelta>> {
    let mut deltas = Vec::new();
    for target in targets {
        let (sheet_name, cell_ref) = parse_sheet_cell_ref(&target)?;
        let before = read_target_cell(baseline, &sheet_name, &cell_ref)?;
        let after = read_target_cell(current, &sheet_name, &cell_ref)?;
        let changed = !cell_values_equal(before.value.as_ref(), after.value.as_ref())
            || before.formula != after.formula;
        deltas.push(TargetDelta {
            address: target,
            before: before.value,
            after: after.value,
            before_formula: before.formula,
            after_formula: after.formula,
            changed,
        });
    }
    Ok(deltas)
}

fn parse_sheet_cell_ref(raw: &str) -> Result<(String, String)> {
    let (sheet_name, cell_ref) = raw.rsplit_once('!').ok_or_else(|| {
        anyhow!(
            "invalid argument: target '{}' must use Sheet!A1 notation",
            raw
        )
    })?;
    if sheet_name.trim().is_empty() || cell_ref.trim().is_empty() {
        bail!(
            "invalid argument: target '{}' must use Sheet!A1 notation",
            raw
        );
    }
    Ok((sheet_name.to_string(), cell_ref.to_string()))
}

fn read_target_cell(
    workbook: &WorkbookContext,
    sheet_name: &str,
    cell_ref: &str,
) -> Result<TargetCellSnapshot> {
    workbook.with_sheet(sheet_name, |sheet| {
        if let Some(cell) = sheet.get_cell(cell_ref) {
            let formula = non_empty_formula(cell.get_formula());
            TargetCellSnapshot {
                value: cell_to_value(cell),
                formula,
            }
        } else {
            TargetCellSnapshot {
                value: None,
                formula: None,
            }
        }
    })
}

fn collect_error_cells(workbook: &WorkbookContext) -> Result<BTreeMap<String, ErrorCellSnapshot>> {
    let mut out = BTreeMap::new();
    for sheet_name in workbook.sheet_names() {
        let sheet_errors = workbook.with_sheet(&sheet_name, |sheet| {
            let mut items = Vec::new();
            for cell in sheet.get_cell_collection() {
                let raw = cell.get_value();
                if !is_error_text(&raw) {
                    continue;
                }
                let address = format!("{}!{}", sheet_name, cell.get_coordinate().get_coordinate());
                items.push((
                    address,
                    ErrorCellSnapshot {
                        error: raw.to_string(),
                        formula: non_empty_formula(cell.get_formula()),
                    },
                ));
            }
            items
        })?;
        for (address, snapshot) in sheet_errors {
            out.insert(address, snapshot);
        }
    }
    Ok(out)
}

fn compare_error_maps(
    baseline: &BTreeMap<String, ErrorCellSnapshot>,
    current: &BTreeMap<String, ErrorCellSnapshot>,
) -> (Vec<ErrorDelta>, Vec<ErrorDelta>) {
    let mut new_errors = Vec::new();
    let mut preexisting_errors = Vec::new();

    for (address, after) in current {
        if let Some(before) = baseline.get(address) {
            preexisting_errors.push(ErrorDelta {
                address: address.clone(),
                before_error: Some(before.error.clone()),
                after_error: Some(after.error.clone()),
                before_formula: before.formula.clone(),
                after_formula: after.formula.clone(),
            });
        } else {
            new_errors.push(ErrorDelta {
                address: address.clone(),
                before_error: None,
                after_error: Some(after.error.clone()),
                before_formula: None,
                after_formula: after.formula.clone(),
            });
        }
    }

    (new_errors, preexisting_errors)
}

fn compare_named_ranges(
    baseline: &[NamedRangeDescriptor],
    current: &[NamedRangeDescriptor],
) -> Vec<NamedRangeDelta> {
    let base_map: BTreeMap<String, &NamedRangeDescriptor> = baseline
        .iter()
        .map(|item| (named_range_key(item), item))
        .collect();
    let current_map: BTreeMap<String, &NamedRangeDescriptor> = current
        .iter()
        .map(|item| (named_range_key(item), item))
        .collect();

    let keys: BTreeSet<String> = base_map
        .keys()
        .cloned()
        .chain(current_map.keys().cloned())
        .collect();

    let mut deltas = Vec::new();
    for key in keys {
        let before = base_map.get(&key).copied();
        let after = current_map.get(&key).copied();
        match (before, after) {
            (Some(b), Some(a)) if b.refers_to != a.refers_to || b.kind != a.kind => {
                deltas.push(NamedRangeDelta {
                    name: a.name.clone(),
                    scope_kind: a.scope_kind,
                    scope_sheet_name: a.scope_sheet_name.clone(),
                    change: "changed".to_string(),
                    before_refers_to: Some(b.refers_to.clone()),
                    after_refers_to: Some(a.refers_to.clone()),
                    before_kind: Some(b.kind.clone()),
                    after_kind: Some(a.kind.clone()),
                });
            }
            (Some(b), None) => {
                deltas.push(NamedRangeDelta {
                    name: b.name.clone(),
                    scope_kind: b.scope_kind,
                    scope_sheet_name: b.scope_sheet_name.clone(),
                    change: "removed".to_string(),
                    before_refers_to: Some(b.refers_to.clone()),
                    after_refers_to: None,
                    before_kind: Some(b.kind.clone()),
                    after_kind: None,
                });
            }
            (None, Some(a)) => {
                deltas.push(NamedRangeDelta {
                    name: a.name.clone(),
                    scope_kind: a.scope_kind,
                    scope_sheet_name: a.scope_sheet_name.clone(),
                    change: "added".to_string(),
                    before_refers_to: None,
                    after_refers_to: Some(a.refers_to.clone()),
                    before_kind: None,
                    after_kind: Some(a.kind.clone()),
                });
            }
            _ => {}
        }
    }

    deltas
}

fn named_range_key(item: &NamedRangeDescriptor) -> String {
    format!(
        "{}|{:?}|{}|{:?}",
        item.name,
        item.scope_kind,
        item.scope_sheet_name.as_deref().unwrap_or(""),
        item.kind
    )
}

fn cell_values_equal(left: Option<&CellValue>, right: Option<&CellValue>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(l), Some(r)) => serde_json::to_value(l).ok() == serde_json::to_value(r).ok(),
        _ => false,
    }
}

fn non_empty_formula(raw: &str) -> Option<String> {
    let trimmed = raw.trim();
    (!trimmed.is_empty()).then(|| trimmed.to_string())
}

fn is_error_text(raw: &str) -> bool {
    let upper = raw.trim().to_ascii_uppercase();
    matches!(
        upper.as_str(),
        "#DIV/0!"
            | "#VALUE!"
            | "#NAME?"
            | "#REF!"
            | "#N/A"
            | "#NULL!"
            | "#NUM!"
            | "#SPILL!"
            | "#CALC!"
            | "#BUSY!"
            | "#FIELD!"
            | "#UNKNOWN!"
    )
}
