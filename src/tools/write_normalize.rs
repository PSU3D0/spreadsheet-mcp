use crate::model::Warning;
use crate::tools::fork::{CellEdit, EditBatchParams};
use anyhow::{Result, bail};
use schemars::JsonSchema;
use serde::Deserialize;

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct EditBatchParamsInput {
    pub fork_id: String,
    pub sheet_name: String,
    pub edits: Vec<CellEditInput>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
#[serde(untagged)]
pub enum CellEditInput {
    Shorthand(String),
    Object(CellEditV2),
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct CellEditV2 {
    pub address: String,
    #[serde(default)]
    pub value: Option<String>,
    #[serde(default)]
    pub formula: Option<String>,
    #[serde(default)]
    pub is_formula: Option<bool>,
}

pub fn normalize_edit_batch(
    params: EditBatchParamsInput,
) -> Result<(EditBatchParams, Vec<Warning>)> {
    let mut warnings = Vec::new();
    let mut edits = Vec::with_capacity(params.edits.len());

    for edit in params.edits {
        match edit {
            CellEditInput::Shorthand(entry) => {
                let Some((address_raw, rhs_raw)) = entry.split_once('=') else {
                    bail!("invalid shorthand edit: '{}'", entry);
                };

                let address = address_raw.trim();
                if address.is_empty() {
                    bail!("invalid shorthand edit: '{}'", entry);
                }

                let mut value = rhs_raw.to_string();
                let mut is_formula = false;

                warnings.push(Warning {
                    code: "WARN_SHORTHAND_EDIT".to_string(),
                    message: format!("Parsed shorthand edit '{}'", entry),
                });

                let rhs_trimmed = rhs_raw.trim_start();
                if let Some(stripped) = rhs_trimmed.strip_prefix('=') {
                    value = stripped.to_string();
                    is_formula = true;
                    warnings.push(Warning {
                        code: "WARN_FORMULA_PREFIX".to_string(),
                        message: format!("Stripped leading '=' for formula '{}'", entry),
                    });
                }

                edits.push(CellEdit {
                    address: address.to_string(),
                    value,
                    is_formula,
                });
            }
            CellEditInput::Object(obj) => {
                let address = obj.address.trim();
                if address.is_empty() {
                    bail!("edit address is required");
                }

                let (value, is_formula) = if let Some(formula) = obj.formula {
                    if let Some(stripped) = formula.strip_prefix('=') {
                        warnings.push(Warning {
                            code: "WARN_FORMULA_PREFIX".to_string(),
                            message: format!("Stripped leading '=' for formula at {}", address),
                        });
                        (stripped.to_string(), true)
                    } else {
                        (formula, true)
                    }
                } else if let Some(value) = obj.value {
                    if let Some(stripped) = value.strip_prefix('=') {
                        warnings.push(Warning {
                            code: "WARN_FORMULA_PREFIX".to_string(),
                            message: format!("Stripped leading '=' for formula at {}", address),
                        });
                        (stripped.to_string(), true)
                    } else {
                        (value, obj.is_formula.unwrap_or(false))
                    }
                } else {
                    bail!("edit value or formula is required for {}", address);
                };

                edits.push(CellEdit {
                    address: address.to_string(),
                    value,
                    is_formula,
                });
            }
        }
    }

    Ok((
        EditBatchParams {
            fork_id: params.fork_id,
            sheet_name: params.sheet_name,
            edits,
        },
        warnings,
    ))
}
